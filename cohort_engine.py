import pandas as pd
from pathlib import Path
from io import BytesIO
from openpyxl.styles import Font

CASE_ID_COL = "Case ID"
EAID_COL = "Enterprise Id"
DATE_COL = "On Hold Reason"


def normalizar_nombre(texto):

    return (
        str(texto)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("_", "")
    )


def detectar_columna(columnas, candidatos):

    mapa = {
        normalizar_nombre(col): col
        for col in columnas
    }

    for candidato in candidatos:

        key = normalizar_nombre(candidato)

        if key in mapa:
            return mapa[key]

    return None


def validar_columnas(df):

    global CASE_ID_COL
    global EAID_COL
    global DATE_COL

    CASE_ID_COL = detectar_columna(
        df.columns,
        [
            "Case ID",
            "CaseId",
            "Case Number",
            "Case"
        ]
    )

    EAID_COL = detectar_columna(
        df.columns,
        [
            "Enterprise Id",
            "Enterprise ID",
            "EnterpriseId",
            "EAID"
        ]
    )

    DATE_COL = detectar_columna(
        df.columns,
        [
            "On Hold Reason",
            "On Hold Date",
            "OnHoldReason"
        ]
    )

    faltantes = []

    if CASE_ID_COL is None:
        faltantes.append("Case ID")

    if EAID_COL is None:
        faltantes.append("Enterprise Id")

    if DATE_COL is None:
        faltantes.append("On Hold Reason")

    if faltantes:

        raise Exception(
            "Columnas requeridas no encontradas:\n\n"
            + "\n".join(faltantes)
            + "\n\nColumnas encontradas:\n\n"
            + "\n".join(map(str, df.columns))
        )


def preparar_datos(df):

    df = df.copy()

    df[DATE_COL] = pd.to_datetime(
        df[DATE_COL],
        errors="coerce"
    )

    df["GroupDate"] = (
        df.groupby(EAID_COL)[DATE_COL]
        .transform("min")
    )

    df = df.sort_values(
        by=[
            "GroupDate",
            EAID_COL,
            DATE_COL
        ],
        ascending=True
    )

    return df


def construir_grupos(df):

    grupos = []

    grouped = df.groupby(
        EAID_COL,
        sort=False,
        dropna=False
    )

    for eaid, rows in grouped:

        grupos.append(
            {
                "eaid": eaid,
                "rows": rows.copy(),
                "size": len(rows),
                "oldest_date": rows["GroupDate"].iloc[0]
            }
        )

    grupos.sort(
        key=lambda g:
        (
            pd.Timestamp.max
            if pd.isna(g["oldest_date"])
            else g["oldest_date"]
        )
    )

    return grupos


def seleccionar_grupos_dia(
        grupos,
        capacidad_objetivo):

    seleccionados = []
    acumulado = 0

    for grupo in grupos:

        seleccionados.append(grupo)
        acumulado += grupo["size"]

        if acumulado >= capacidad_objetivo:
            break

    return seleccionados, acumulado


def distribuir_grupos(
        grupos,
        cantidad_personas):

    personas = {}

    for i in range(cantidad_personas):

        personas[f"Person_{i+1}"] = {
            "cases": 0,
            "groups": []
        }

    grupos = sorted(
        grupos,
        key=lambda x: x["size"],
        reverse=True
    )

    for grupo in grupos:

        destino = min(
            personas,
            key=lambda p:
            personas[p]["cases"]
        )

        personas[destino]["groups"].append(
            grupo
        )

        personas[destino]["cases"] += (
            grupo["size"]
        )

    return personas


def crear_detalle(personas):

    frames = []

    for owner, data in personas.items():

        for grupo in data["groups"]:

            temp = grupo["rows"].copy()

            temp["Assigned To"] = owner

            frames.append(temp)

    if len(frames) == 0:
        return pd.DataFrame()

    resultado = pd.concat(
        frames,
        ignore_index=True
    )

    resultado = resultado[
        resultado[CASE_ID_COL].notna()
    ]

    return resultado


def crear_summary(personas):

    filas = []

    for owner, data in personas.items():

        filas.append(
            {
                "Owner": owner,
                "Cases": data["cases"],
                "EAIDs": len(
                    data["groups"]
                )
            }
        )

    return pd.DataFrame(filas)


def crear_dashboard(
        detalle,
        summary,
        personas,
        objetivo,
        asignados,
        total_archivo):

    max_load = (
        summary["Cases"].max()
        if not summary.empty
        else 0
    )

    min_load = (
        summary["Cases"].min()
        if not summary.empty
        else 0
    )

    return pd.DataFrame(
        {
            "Metric": [
                "People",
                "Target per Person",
                "Target Capacity",
                "Total Cases In File",
                "Cases Assigned",
                "Cases Remaining",
                "Enterprise IDs",
                "Maximum Load",
                "Minimum Load",
                "Load Difference",
                "EAIDs Split"
            ],
            "Value": [
                personas,
                objetivo,
                personas * objetivo,
                total_archivo,
                len(detalle),
                total_archivo - len(detalle),
                detalle[EAID_COL].nunique(),
                max_load,
                min_load,
                max_load - min_load,
                0
            ]
        }
    )


def autoajustar(writer):

    workbook = writer.book

    for hoja in workbook.sheetnames:

        ws = workbook[hoja]

        for cell in ws[1]:
            cell.font = Font(
                bold=True
            )

        for column in ws.columns:

            max_length = 0

            column_letter = (
                column[0].column_letter
            )

            for cell in column:

                value = (
                    str(cell.value)
                    if cell.value is not None
                    else ""
                )

                max_length = max(
                    max_length,
                    len(value)
                )

            ws.column_dimensions[
                column_letter
            ].width = max_length + 4


def exportar_memoria(
        dashboard,
        summary,
        detalle):

    buffer = BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        dashboard.to_excel(
            writer,
            sheet_name="Dashboard",
            index=False
        )

        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        detalle.to_excel(
            writer,
            sheet_name="Detail",
            index=False
        )

        detalle[
            [CASE_ID_COL]
        ].to_excel(
            writer,
            sheet_name="CaseIDs",
            index=False
        )

        autoajustar(writer)

    buffer.seek(0)

    return buffer


def generar_asignacion(
        archivo,
        personas,
        objetivo):

    if personas <= 0:
        raise ValueError(
            "La cantidad de personas debe ser mayor que cero."
        )

    if objetivo <= 0:
        raise ValueError(
            "Los casos por persona deben ser mayores que cero."
        )

    extension = (
        Path(
            archivo.name
        ).suffix.lower()
    )

    if extension == ".xlsb":

        df = pd.read_excel(
            archivo,
            engine="pyxlsb"
        )

    else:

        df = pd.read_excel(
            archivo,
            engine="openpyxl"
        )

    if df.empty:

        raise Exception(
            "El archivo no contiene datos."
        )

    total_archivo = len(df)

    validar_columnas(df)

    df = preparar_datos(df)

    grupos = construir_grupos(df)

    grupos_dia, asignados = (
        seleccionar_grupos_dia(
            grupos,
            personas * objetivo
        )
    )

    personas_data = (
        distribuir_grupos(
            grupos_dia,
            personas
        )
    )

    detalle = crear_detalle(
        personas_data
    )

    summary = crear_summary(
        personas_data
    )

    dashboard = crear_dashboard(
        detalle,
        summary,
        personas,
        objetivo,
        asignados,
        total_archivo
    )

    return exportar_memoria(
        dashboard,
        summary,
        detalle
    )