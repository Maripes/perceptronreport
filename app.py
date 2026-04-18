import streamlit as st
import pandas as pd
import io
from io import StringIO
import plotly.express as px
from xlsxwriter.utility import xl_col_to_name

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Comparación cmm", layout="wide")
st.title("Comparación Vs CMM")

# =========================
# CARGA DE ARCHIVOS
# =========================
archivo_L = st.file_uploader(
    "Carga el archivo TXT - Lado Izquierdo",
    type=["txt"]
)

archivo_R = st.file_uploader(
    "Carga el archivo TXT - Lado Derecho",
    type=["txt"]
)

# =========================
# LECTOR ROBUSTO TXT
# =========================
def leer_txt(archivo):
    contenido = archivo.read().decode("utf-8", errors="ignore")
    lineas = contenido.splitlines()

    inicio = None
    for i, linea in enumerate(lineas):
        if (
            "Cycle Time" in linea and
            "Corr. Coef." in linea and
            "Offset" in linea and
            "T-Test" in linea and
            "F-Test" in linea
        ):
            inicio = i
            break

    if inicio is None:
        st.error("No se encontró la fila de encabezados reales")
        st.stop()

    datos = "\n".join(lineas[inicio:])

    df = pd.read_csv(
        StringIO(datos),
        sep=r"\t+",
        engine="python",
        header=0,
        on_bad_lines="skip"
    )

    return df

# =========================
# FUNCIONES DE COLOR
# =========================
def color_t_test(val):
    try:
        val = float(val)
        return (
            "background-color: #00C853; color: black; font-weight: bold"
            if val < 0.005
            else "background-color: #D50000; color: white; font-weight: bold"
        )
    except:
        return ""

def color_f_test(val):
    try:
        val = float(val)
        return (
            "background-color: #D50000; color: white; font-weight: bold"
            if val < 0.005
            else "background-color: #00C853; color: black; font-weight: bold"
        )
    except:
        return ""

def color_corr(val):
    try:
        val = float(val)
        if val >= 0.95:
            return "background-color: #00C853; color: black; font-weight: bold"
        elif 0.90 <= val <= 0.94:
            return "background-color: #FFD600; color: black; font-weight: bold"
        else:
            return "background-color: #D50000; color: white; font-weight: bold"
    except:
        return ""

def color_offset(val):
    try:
        val = float(val)
        return (
            "background-color: #AA00FF; color: white; font-weight: bold"
            if abs(val) > 0.5
            else ""
        )
    except:
        return ""


# =========================
# PROCESO PRINCIPAL
# =========================
if archivo_L and archivo_R:

    df_L = leer_txt(archivo_L)
    df_R = leer_txt(archivo_R)

    df = pd.concat([df_L, df_R], ignore_index=True)
    for col in ["T-Test", "F-Test", "Corr. Coef.", "Offset"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # =========================
    # LIMPIEZA
    # =========================
    col_cycle = df.columns[0]

    df[col_cycle] = df[col_cycle].astype(str).str.strip()
    df = df[df[col_cycle] != ""]
    df = df[~df[col_cycle].str.startswith("CT", na=False)]

    # =========================
    # EXTRAER NOMBRE / EJE
    # =========================
    df["Nombre"] = df[col_cycle].str.extract(r"(^\d+)")
    #df["Punto"] = df[col_cycle].str.extract(r"^(.+?)(?=\[)")
    df["Eje"] = df[col_cycle].str.extract(r"\[([A-Z])\]")
    df["Base"] = df[col_cycle].str.extract(r"^(.+?)(?=[LR]\[)")
    df["Lado"] = df[col_cycle].str.extract(r"([LR])(?=\[)")
    # =========================
    # FILTROS INTERACTIVOS
    # =========================
    st.sidebar.header("Filtros")

    nombres = sorted(df["Nombre"].dropna().unique())
    ejes = sorted(df["Eje"].dropna().unique())

    nombre_sel = st.sidebar.multiselect(
        "Filtrar por nombre (Cycle Time)",
        nombres,
        default=nombres
    )

    eje_sel = st.sidebar.multiselect(
        "Filtrar por eje",
        ejes,
        default=ejes
    )

    df_filtrado = df[
        df["Nombre"].isin(nombre_sel) &
        df["Eje"].isin(eje_sel)
    ]

    df_filtrado = df_filtrado.copy()

    # Limpiar nombres de columnas (quita espacios invisibles)
    df_filtrado.columns = df_filtrado.columns.str.strip()

    # Limpiar espacios dentro de las celdas (blindaje extra)
    df_filtrado = df_filtrado.applymap(
        lambda x: x.strip() if isinstance(x, str) else x
    )
    # =========================
    # ORDEN TIPO EXCEL
    # =========================
    def orden_excel(nombre):
        nombre = str(nombre)
        lado = 0 if "L[" in nombre else 1
        eje = 0 if "[Y]" in nombre else 1
        return lado, eje

    df_filtrado["__orden"] = df_filtrado[col_cycle].apply(orden_excel)
    df_filtrado = df_filtrado.sort_values("__orden").drop(columns="__orden")

    # =========================
    # SALIDA
    # =========================
    st.success("Filtro aplicado correctamente")
    st.write("Filas visibles:", df_filtrado.shape[0])

    styled_df = (
        df_filtrado.style
        .apply(lambda col: col.map(color_t_test) if col.name == "T-Test" else [""]*len(col), axis=0)
        .apply(lambda col: col.map(color_f_test) if col.name == "F-Test" else [""]*len(col), axis=0)
        .apply(lambda col: col.map(color_corr) if col.name == "Corr. Coef." else [""]*len(col), axis=0)
        .apply(lambda col: col.map(color_offset) if col.name == "Offset" else [""]*len(col), axis=0)
    )

    st.dataframe(styled_df, use_container_width=True)
    st.subheader("📊 Heatmap Histórico")

    columnas_fecha = [
        col for col in df_filtrado.columns
        if "/" in col and ":" in col
    ]

    columnas_fecha = []

    for col in df_filtrado.columns:
        try:
            pd.to_datetime(col)
            columnas_fecha.append(col)
        except:
            pass

    df_heat = df_filtrado.set_index(col_cycle)[columnas_fecha]

    fig = px.imshow(
        df_heat,
        aspect="auto",
        color_continuous_scale="RdYlGn_r"
    )
    st.plotly_chart(fig, use_container_width=True)
    
    st.subheader("📈 Tendencia por Fecha")

    # Detectar automáticamente columnas que son fechas
    columnas_fecha = [
        col for col in df_filtrado.columns
        if "/" in col and ":" in col
    ]

    # Selector de Cycle Time
    cycle_sel = st.selectbox(
        "Selecciona Cycle Time",
        df_filtrado[col_cycle].unique()
    )

    #df_graf = df_filtrado[df_filtrado[col_cycle] == cycle_sel]
    df_filtrado[col_cycle] = df_filtrado[col_cycle].str.strip()
    cycle_sel = cycle_sel.strip()

    df_graf = df_filtrado[df_filtrado[col_cycle] == cycle_sel]

    # Transponer para convertir columnas en eje X
    serie = df_graf[columnas_fecha].T
    serie.columns = ["Valor"]

    # Convertir índice a datetime
    serie.index = pd.to_datetime(serie.index, errors="coerce")

    st.line_chart(serie)
    # =========================
    # DESCARGAR A EXCEL PRO (CON COLORES + FILTROS + AUTOAJUSTE + KPIs)
    # =========================
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        df_filtrado["Punto"] = df_filtrado["Nombre"].astype(str).str.strip()
        df_filtrado.to_excel(writer, index=False, sheet_name="Comparacion")

        workbook = writer.book
        worksheet = writer.sheets["Comparacion"]

        # =========================
        # FORMATOS
        # =========================
        format_green = workbook.add_format({"bg_color": "#00C853", "font_color": "black", "bold": True})
        format_red = workbook.add_format({"bg_color": "#D50000", "font_color": "white", "bold": True})
        format_yellow = workbook.add_format({"bg_color": "#FFD600", "font_color": "black", "bold": True})
        format_purple = workbook.add_format({"bg_color": "#AA00FF", "font_color": "white", "bold": True})
        format_header = workbook.add_format({"bold": True, "border": 1})
        format_border = workbook.add_format({"border": 1})

        # =========================
        # FORMATO ENCABEZADO
        # =========================
        for col_num, value in enumerate(df_filtrado.columns):
            worksheet.write(0, col_num, value, format_header)

        # =========================
        # AUTO AJUSTAR COLUMNAS
        # =========================
        for i, col in enumerate(df_filtrado.columns):
            max_len = max(
                df_filtrado[col].astype(str).map(len).max(),
                len(col)
            ) + 2
            worksheet.set_column(i, i, max_len)

        # =========================
        # CONGELAR ENCABEZADO
        # =========================
        worksheet.freeze_panes(1, 0)

        # =========================
        # FILTRO AUTOMÁTICO
        # =========================
        worksheet.autofilter(
            0, 0,
            df_filtrado.shape[0],
            df_filtrado.shape[1] - 1
        )

        # =========================
        # COLORES CELDA POR CELDA
        # =========================
        col_t = df_filtrado.columns.get_loc("T-Test")
        col_f = df_filtrado.columns.get_loc("F-Test")
        col_corr = df_filtrado.columns.get_loc("Corr. Coef.")
        col_offset = df_filtrado.columns.get_loc("Offset")

        #for row_num, row in df_filtrado.iterrows():
        #    excel_row = row_num + 1

        for excel_row, (_, row) in enumerate(df_filtrado.iterrows(), start=1):

            for col_num in range(len(df_filtrado.columns)):
                valor = row.iloc[col_num]

                if pd.isna(valor) or valor in [float("inf"), float("-inf")]:
                    worksheet.write(excel_row, col_num, "", format_border)
                else:
                    worksheet.write(excel_row, col_num, valor, format_border)

            # T-Test
            try:
                if row["T-Test"] < 0.005:
                    worksheet.write(excel_row, col_t, row["T-Test"], format_green)
                else:
                    worksheet.write(excel_row, col_t, row["T-Test"], format_red)
            except:
                pass

            # F-Test
            try:
                if row["F-Test"] < 0.005:
                    worksheet.write(excel_row, col_f, row["F-Test"], format_red)
                else:
                    worksheet.write(excel_row, col_f, row["F-Test"], format_green)
            except:
                pass

            # Corr. Coef.
            try:
                val = row["Corr. Coef."]
                if val >= 0.95:
                    worksheet.write(excel_row, col_corr, val, format_green)
                elif 0.90 <= val < 0.95:
                    worksheet.write(excel_row, col_corr, val, format_yellow)
                else:
                    worksheet.write(excel_row, col_corr, val, format_red)
            except:
                pass

            # Offset
            try:
                val = row["Offset"]
                if abs(val) > 0.5:
                    worksheet.write(excel_row, col_offset, val, format_purple)
            except:
                pass

        # ======================================================
        # ================= DASHBOARD DINÁMICO =================
        # ======================================================

        total = len(df_filtrado)
        fallas_t = (pd.to_numeric(df_filtrado["T-Test"], errors="coerce") >= 0.005).sum()
        fallas_corr = (pd.to_numeric(df_filtrado["Corr. Coef."], errors="coerce") < 0.95).sum()
        offsets_altos = (pd.to_numeric(df_filtrado["Offset"], errors="coerce").abs() > 0.5).sum()

        dashboard = workbook.add_worksheet("Dashboard")

        dashboard.write("A1", "DASHBOARD EJECUTIVO", format_header)

        dashboard.write("A3", "Total Mediciones")
        dashboard.write("B3", total)

        dashboard.write("A4", "T-Test fuera de rango")
        dashboard.write("B4", fallas_t)

        dashboard.write("A5", "Correlación baja")
        dashboard.write("B5", fallas_corr)

        dashboard.write("A6", "Offset alto")
        dashboard.write("B6", offsets_altos)

        # ================= DROPDOWN POR PUNTO =================

        puntos_unicos = (
            df_filtrado["Punto"]
            .dropna()
            .unique()
            .tolist()
        )

        dashboard.write("A9", "Selecciona Punto:")

        col_lista = 25  # Columna Z

        for i, valor in enumerate(puntos_unicos):
            dashboard.write(i, col_lista, valor)

        rango_dropdown = f"=Dashboard!${xl_col_to_name(col_lista)}$1:${xl_col_to_name(col_lista)}${len(puntos_unicos)}"

        dashboard.data_validation(
            "B9",
            {
                "validate": "list",
                "source": rango_dropdown,
            },
        )

        dashboard.set_column(col_lista, col_lista, None, None, {"hidden": True})

        # ================= COLUMNAS FECHA =================

        columnas_fecha = []

        for col in df_filtrado.columns:
            try:
                pd.to_datetime(col)
                columnas_fecha.append(col)
            except:
                pass

        fila_inicio = 12
        dashboard.write_row(
            fila_inicio, 0,
            ["Fecha", "Lado Izq", "Lado Der", "Diferencia"]
        )

        # Columnas necesarias
        col_punto = df_filtrado.columns.get_loc("Punto")
        col_lado = df_filtrado.columns.get_loc("Lado")

        col_punto_excel = xl_col_to_name(col_punto)
        col_lado_excel = xl_col_to_name(col_lado)

        rango_punto = f"Comparacion!${col_punto_excel}$2:${col_punto_excel}${len(df_filtrado)+1}"
        rango_lado = f"Comparacion!${col_lado_excel}$2:${col_lado_excel}${len(df_filtrado)+1}"

        for i, fecha in enumerate(columnas_fecha):

            col_idx = df_filtrado.columns.get_loc(fecha)
            col_excel = xl_col_to_name(col_idx)

            rango_valores = f"Comparacion!${col_excel}$2:${col_excel}${len(df_filtrado)+1}"

            formula_L = (
                f'=IFERROR(AVERAGEIFS('
                f'{rango_valores},'
                f'{rango_punto},$B$9,'
                f'{rango_lado},"L"),0)'
            )

            formula_R = (
                f'=IFERROR(AVERAGEIFS('
                f'{rango_valores},'
                f'{rango_punto},$B$9,'
                f'{rango_lado},"R"),0)'
            )

            dashboard.write(fila_inicio + i + 1, 0, fecha)
            dashboard.write_formula(fila_inicio + i + 1, 1, formula_L)
            dashboard.write_formula(fila_inicio + i + 1, 2, formula_R)

            dashboard.write_formula(
                fila_inicio + i + 1,
                3,
                f"=B{fila_inicio+i+2}-C{fila_inicio+i+2}"
            )

        # ================= GRAFICA DINÁMICA =================

        chart_dyn = workbook.add_chart({"type": "line"})

        chart_dyn.add_series({
            "name": "Lado Izquierdo",
            "categories": ["Dashboard", fila_inicio + 1, 0, fila_inicio + len(columnas_fecha), 0],
            "values": ["Dashboard", fila_inicio + 1, 1, fila_inicio + len(columnas_fecha), 1],
        })

        chart_dyn.add_series({
            "name": "Lado Derecho",
            "categories": ["Dashboard", fila_inicio + 1, 0, fila_inicio + len(columnas_fecha), 0],
            "values": ["Dashboard", fila_inicio + 1, 2, fila_inicio + len(columnas_fecha), 2],
        })

        chart_dyn.add_series({
            "name": "Diferencia",
            "categories": ["Dashboard", fila_inicio + 1, 0, fila_inicio + len(columnas_fecha), 0],
            "values": ["Dashboard", fila_inicio + 1, 3, fila_inicio + len(columnas_fecha), 3],
        })

        chart_dyn.set_title({"name": "Comparación Dinámica L vs R"})
        chart_dyn.set_style(10)

        dashboard.insert_chart("F12", chart_dyn)

        # ================= TOP 10 OFFSETS =================

        df_offset_top = df_filtrado.copy()
        df_offset_top["AbsOffset"] = df_offset_top["Offset"].abs()
        top10 = df_offset_top.sort_values("AbsOffset", ascending=False).head(10)

        fila_top = 30

        dashboard.write(fila_top, 0, "Top 10 Offsets Críticos", format_header)

        dashboard.write_row(fila_top + 1, 0, ["Cycle", "Offset"])

        for i, row in top10.iterrows():
            dashboard.write_row(fila_top + 2, 0, [row[col_cycle], row["Offset"]])
            fila_top += 1

        chart_offset = workbook.add_chart({"type": "column"})

        chart_offset.add_series({
            "categories": ["Dashboard", fila_top - 9, 0, fila_top - 1, 0],
            "values": ["Dashboard", fila_top - 9, 1, fila_top - 1, 1],
            "name": "Offset Absoluto"
        })

        chart_offset.set_title({"name": "Top 10 Offsets"})
        dashboard.insert_chart("E30", chart_offset)


    excel_data = output.getvalue()

    st.download_button(
        label="📥 Descargar Excel PRO",
        data=excel_data,
        file_name="Comparacion_CMM_PRO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


else:
    st.info("Carga ambos archivos TXT para continuar")

#streamlit run "C:\Users\maripes3\Documents\Comparaciones\MIS DATOS\MIS DATOS\ComparacionVsCMM.py"
