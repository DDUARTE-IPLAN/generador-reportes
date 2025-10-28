from st_aggrid import GridOptionsBuilder, JsCode

def build_date_comparators():
    cmp_fecha_crea = JsCode("""
        function(valueA, valueB, nodeA, nodeB, isInverted) {
            const A = nodeA.data? nodeA.data._FECHA_CREACION_ISO : null;
            const B = nodeB.data? nodeB.data._FECHA_CREACION_ISO : null;
            if (!A && !B) return 0; if (!A) return -1; if (!B) return 1;
            return A < B ? -1 : (A > B ? 1 : 0);
        }
    """)
    cmp_fecha_act = JsCode("""
        function(valueA, valueB, nodeA, nodeB, isInverted) {
            const A = nodeA.data? nodeA.data._FECHA_ACTIVACION_ISO : null;
            const B = nodeB.data? nodeB.data._FECHA_ACTIVACION_ISO : null;
            if (!A && !B) return 0; if (!A) return -1; if (!B) return 1;
            return A < B ? -1 : (A > B ? 1 : 0);
        }
    """)
    return cmp_fecha_crea, cmp_fecha_act

def configure_common_grid(gb: GridOptionsBuilder):
    gb.configure_pagination(enabled=True, paginationAutoPageSize=False, paginationPageSize=100)
    gb.configure_default_column(sortable=True, filter=True, resizable=True)
    gb.configure_grid_options(
        enableRangeSelection=True,
        enableCellTextSelection=True,
        copyHeadersToClipboard=True,
        ensureDomOrder=True,
    )
