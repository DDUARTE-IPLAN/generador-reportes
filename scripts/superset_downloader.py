# scripts/superset_downloader.py
from __future__ import annotations

import re
import time
import os
from datetime import datetime
from typing import Callable, List, Optional
import pathlib

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout, expect


# =========================================================
# Utils
# =========================================================
def _log(log: Optional[Callable[[str], None]], msg: str) -> None:
    (log or print)(msg)


def _ensure_dir(p: pathlib.Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def _is_dashboard_url(url: str) -> bool:
    return "/superset/dashboard/" in url


def _is_explore_url(url: str) -> bool:
    return "/superset/explore/" in url


def _wait_idle(page, timeout_ms: int = 15000) -> None:
    page.wait_for_load_state("domcontentloaded", timeout=timeout_ms)
    try:
        page.wait_for_load_state("networkidle", timeout=timeout_ms)
    except PWTimeout:
        # OK: a veces nunca llega a networkidle por streams/WS
        pass


def _name_with_stamp(suggested: str, dest_dir: pathlib.Path) -> pathlib.Path:
    base = suggested.rsplit(".", 1)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    if len(base) == 2:
        return dest_dir / f"{base[0]}_{stamp}.{base[1]}"
    return dest_dir / f"{suggested}_{stamp}.csv"


# =========================================================
# Keycloak login (si aparece)
# =========================================================
def _login_keycloak_if_present(page, user: str, pwd: str, log) -> None:
    try:
        u = page.locator("input#username, input[name='username']").first
        p = page.locator("input#password, input[name='password']").first
        if u.count() and p.count():
            _log(log, "🔐 Keycloak detectado. Iniciando sesión…")
            u.fill(user or "")
            p.fill(pwd or "")
            btn = page.locator(
                "button#kc-login, "
                "button:has-text('Ingresar'), button:has-text('Sign in'), "
                "button[type='submit'], input[type='submit']"
            ).first
            with page.expect_navigation(wait_until="networkidle", timeout=60000):
                btn.click()
            _log(log, "✅ Login OK")
    except Exception as e:
        _log(log, f"⚠️ No pude completar login: {e}")


# =========================================================
# Selectores / helpers de menú
# =========================================================
def _open_header_menu(btn, page) -> bool:
    try:
        btn.scroll_into_view_if_needed(timeout=3000)
        btn.hover(timeout=1500)
        expect(btn).to_be_enabled(timeout=3000)
        btn.click(timeout=3000)
        expect(_menu_verticals(page).last).to_be_visible(timeout=5000)
        return True
    except Exception:
        return False


def _menu_verticals(page):
    return page.locator(
        "div.ant-dropdown:visible ul.ant-dropdown-menu.ant-dropdown-menu-root.ant-dropdown-menu-vertical:visible"
    )


def _submenu_verticals(page):
    return page.locator("div.ant-dropdown:visible ul.ant-dropdown-menu-sub:visible")


def _click_item_by_patterns(container, patterns) -> bool:
    # primero por role=menuitem con regex, después por texto simple
    for pat in patterns:
        try:
            loc = container.get_by_role("menuitem", name=re.compile(pat, re.I))
            if loc.count() > 0 and loc.first.is_visible():
                loc.first.click()
                return True
        except Exception:
            pass
    for pat in patterns:
        try:
            loc = container.locator(f"text=/{pat}/i")
            if loc.count() > 0 and loc.first.is_visible():
                loc.first.click()
                return True
        except Exception:
            pass
    return False


def _open_submenu_if_any(page, menu):
    parents = [r"Download", r"Export", r"Descargar", r"Exportar"]
    for pat in parents:
        try:
            cand = menu.get_by_role("menuitem", name=re.compile(pat, re.I))
            if cand.count():
                item = cand.first
                item.hover()
                page.wait_for_timeout(200)
                subs = _submenu_verticals(page)
                if subs.count() > 0:
                    return subs.last
        except Exception:
            pass
    return None


def _find_kebab_buttons(page):
    # prioridad: triggers cercanos al header de chart
    prioridad = page.locator(
        "[data-test='slice-header'] .ant-dropdown-trigger, "
        "[data-test='chart-header'] .ant-dropdown-trigger, "
        ".dashboard-component-chart-holder .ant-dropdown-trigger, "
        ".dashboard-component-chart-holder button[aria-haspopup='menu']"
    )
    if prioridad.count() > 0:
        return prioridad
    # fallback genérico
    sel = (
        "button[aria-haspopup='menu'], "
        "button[aria-expanded][aria-haspopup='menu'], "
        "button[aria-label='More options'], "
        ".ant-dropdown-trigger, "
        ".anticon[tabindex], .anticon-ellipsis[tabindex]"
    )
    return page.locator(sel)


def _click_export_dialog_ok_if_present(page):
    try:
        dialog = page.locator("[role='dialog']:has-text('Export'), [role='dialog']:has-text('Exportar')")
        if dialog.count() and dialog.first.is_visible():
            d = dialog.first
            for name in [r"Export", r"Descargar", r"Aceptar", r"OK"]:
                btn = d.get_by_role("button", name=re.compile(name, re.I))
                if btn.count():
                    btn.first.click()
                    return True
    except Exception:
        pass
    return False


# =========================================================
# Patrones CSV
# =========================================================
_CSV_PATTERNS = [
    r"Export\s*to\s*\.?CSV",
    r"Export\s*to\s*CSV",
    r"Export\s*CSV",
    r"Download\s*as\s*\.?CSV",
    r"Download\s*CSV",
    r"Full\s*CSV",
    r"CSV\s*\(.*\)",
    r"Export\s*full\s*CSV",
    r"Descargar\s*como\s*CSV",
    r"Descargar\s*CSV",
    r"Exportar\s*CSV",
    r"\bCSV\b",
]


# =========================================================
# Descarga desde DASHBOARD
# =========================================================
def _download_from_dashboard(
    page, url: str, dest: pathlib.Path, title_filter_regex: str,
    max_panels: int, timeout_per_panel: int, log
) -> List[pathlib.Path]:
    _log(log, "📊 Detecté enlace de Dashboard. Cargando…")
    page.goto(url, wait_until="domcontentloaded")
    _wait_idle(page, 60000)

    # scroll inicial para que carguen paneles visibles
    try:
        page.mouse.wheel(0, -4000)
    except Exception:
        pass

    buttons = _find_kebab_buttons(page)
    total = buttons.count()
    if total == 0:
        for _ in range(4):
            page.mouse.wheel(0, 1200)
            time.sleep(0.2)
            buttons = _find_kebab_buttons(page)
            total = buttons.count()
            if total > 0:
                break

    if total == 0:
        # captura para debug
        try:
            shot = dest / "debug_dashboard.png"
            page.screenshot(path=str(shot), full_page=True)
        except Exception:
            pass
        return []

    results: List[pathlib.Path] = []
    descargados = 0

    title_regex = None
    if title_filter_regex:
        try:
            title_regex = re.compile(title_filter_regex, re.I)
        except re.error:
            title_regex = None

    for idx in range(total):
        if max_panels and descargados >= max_panels:
            break

        btn = buttons.nth(idx)
        if not _open_header_menu(btn, page):
            continue

        try:
            menus = _menu_verticals(page)
            expect(menus.last).to_be_visible(timeout=8000)
            menu = menus.last
        except PWTimeout:
            page.keyboard.press("Escape")
            continue

        # filtro opcional por título del panel
        try:
            if title_regex:
                container = btn.locator("xpath=ancestor::*[self::div or self::section][1]")
                title_loc = container.locator("h1, h2, h3, [data-test='dashboard-chart-title'], [class*='chart-title']")
                title = title_loc.inner_text().strip() if title_loc.count() else ""
                if title and not title_regex.search(title):
                    page.keyboard.press("Escape")
                    continue
        except Exception:
            pass

        # click en CSV (directo o dentro de submenú)
        if _click_item_by_patterns(menu, _CSV_PATTERNS) or (
            (submenu := _open_submenu_if_any(page, menu)) and _click_item_by_patterns(submenu, _CSV_PATTERNS)
        ):
            _click_export_dialog_ok_if_present(page)
            try:
                dl = page.wait_for_event("download", timeout=timeout_per_panel * 1000)
                outfile = _name_with_stamp(dl.suggested_filename or "export.csv", dest)
                dl.save_as(str(outfile))
                results.append(outfile)
                descargados += 1
                _log(log, f"✅ CSV guardado: {outfile}")
                try:
                    page.keyboard.press("Escape")
                except Exception:
                    pass
                continue
            except PWTimeout:
                pass
            except Exception:
                pass

        try:
            page.keyboard.press("Escape")
        except Exception:
            pass

    return results


# =========================================================
# Descarga desde EXPLORE (gráfico individual)
# =========================================================
def _download_from_explore(
    page, url: str, dest: pathlib.Path, timeout_per_panel: int, log
) -> List[pathlib.Path]:
    _log(log, "📈 Detecté enlace de Explore. Cargando…")
    page.goto(url, wait_until="domcontentloaded")
    _wait_idle(page, 15000)

    # 1) Intentar abrir el menú/botón de descarga
    openers = [
        "button:has-text('Download')",
        "button[aria-label='Download']",
        "[data-test='query-download-button']",
        "[data-test='btn-download']",
        "button:has-text('Export')",
        "button:has-text('More')",
        "button[aria-label*='more' i]",
        # Español:
        "button:has-text('Descargar')",
        "button:has-text('Exportar')",
        "button:has-text('Datos')",
        "role=tab >> text=/^Data$|^Datos$/i",
    ]
    for sel in openers:
        try:
            el = page.locator(sel).first
            if el.count():
                el.click()
                break
        except Exception:
            pass

    # 2) Opción CSV en cualquier menú visible o diálogo
    csv_candidates = [
        "text=Export to .CSV",
        "text=Export CSV",
        "text=Download as CSV",
        "text=Full CSV",
        "text=/\\bCSV\\b/i",
        "text=/CSV\\s*\\(.*\\)/i",
        "text=/Export\\s*full\\s*CSV/i",
        "role=menuitem >> text=/CSV/i",
        "button:has-text('CSV')",
        "[data-test='download-csv']",
        "a[download$='.csv']",
        "text=Descargar CSV",
        "text=Exportar CSV",
        "text=Descargar como CSV",
    ]

    for sel in csv_candidates:
        try:
            with page.expect_download(timeout=timeout_per_panel * 1000) as dl:
                el = page.locator(sel).first
                if not el.count():
                    continue
                el.click()
            d = dl.value
            outfile = _name_with_stamp(d.suggested_filename or "export.csv", dest)
            d.save_as(str(outfile))
            _log(log, f"✅ CSV guardado: {outfile}")
            return [outfile]
        except PWTimeout:
            continue
        except Exception:
            continue

    # 3) Captura para depurar si algo falla
    try:
        shot = dest / "debug_explore.png"
        page.screenshot(path=str(shot), full_page=True)
    except Exception:
        pass

    return []


# =========================================================
# API pública (usada por la app)
# =========================================================
def download_superset_csvs(
    dashboard_url: str,
    download_dir: pathlib.Path | str,
    keycloak_user: str = "",
    keycloak_pass: str = "",
    title_filter_regex: str = "",
    max_panels: int = 0,
    panel_timeout: int = 25,
    headless: bool = False,
    log: Callable[[str], None] | None = None,
) -> List[pathlib.Path]:
    """
    Descarga CSVs desde:
      • Dashboard (varios charts)
      • Explore (un chart)
    Devuelve la lista de archivos guardados.
    """
    dest = pathlib.Path(download_dir).expanduser().resolve()
    _ensure_dir(dest)

    target_url = dashboard_url.strip()
    _log(log, f"Dashboard: {target_url}")
    _log(log, f"Destino: {dest}")

    with sync_playwright() as p:
        def _running_in_docker() -> bool:
            # Sin $DISPLAY o presencia de /.dockerenv
            return (not os.environ.get("DISPLAY")) or pathlib.Path("/.dockerenv").exists()

        # Forzar headless si estamos en server / docker, o si la env var lo pide
        force_headless = _running_in_docker() or os.getenv("PLAYWRIGHT_HEADLESS", "0") in ("1", "true", "True")
        run_headless = True if force_headless else bool(headless)

        launch_args = ["--no-sandbox", "--disable-dev-shm-usage"]

        try:
            browser = p.chromium.launch(headless=run_headless, args=launch_args)
        except Exception:
            # por si alguien pasa headless=False y igual falla en server
            browser = p.chromium.launch(headless=True, args=launch_args)

        context = browser.new_context(accept_downloads=True, permissions=["notifications"])
        page = context.new_page()
        page.set_viewport_size({"width": 1366, "height": 900})

        # ⚡ Bloqueo de recursos pesados para acelerar carga/acciones
        context.route(
            "**/*",
            lambda route: (
                route.abort()
                if route.request.resource_type in {"image", "font", "media"}
                or "googletagmanager" in route.request.url
                or "analytics" in route.request.url
                else route.continue_()
            ),
        )

        page.set_default_timeout(15000)
        page.set_default_navigation_timeout(30000)

        # 1) Ir al target y esperar
        page.goto(target_url, wait_until="domcontentloaded")
        _wait_idle(page, 15000)

        # 2) Si aparece Keycloak, loguear
        if keycloak_user or keycloak_pass:
            _login_keycloak_if_present(page, keycloak_user, keycloak_pass, log)

        # 3) Si nos mandaron a /welcome u otra, reabrimos el permalink
        current = page.url
        if "superset/welcome" in current or (not _is_dashboard_url(current) and not _is_explore_url(current)):
            page.goto(target_url, wait_until="domcontentloaded")
            _wait_idle(page, 15000)

        # 4) Detectar de nuevo con la URL real en pantalla y descargar
        final_url = page.url
        files: List[pathlib.Path] = []
        try:
            if _is_dashboard_url(final_url):
                _log(log, "")
                files = _download_from_dashboard(
                    page, final_url, dest, title_filter_regex, max_panels, panel_timeout, log
                )
            elif _is_explore_url(final_url):
                _log(log, "")
                files = _download_from_explore(page, final_url, dest, panel_timeout, log)
            else:
                # intento Explore y luego Dashboard, sin logs extra
                files = _download_from_explore(page, final_url, dest, panel_timeout, log)
                if not files:
                    files = _download_from_dashboard(
                        page, final_url, dest, title_filter_regex, max_panels, panel_timeout, log
                    )
        finally:
            context.close()
            browser.close()

        return files
