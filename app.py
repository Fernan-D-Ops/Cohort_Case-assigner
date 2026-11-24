import streamlit as st

from cohort_engine import generar_asignacion

st.set_page_config(
    page_title="Case Assigner",
    layout="wide"
)

try:
    AGENTES = dict(
        st.secrets["agents"]
    )
except Exception:
    st.error(
        "No se encontraron los agentes en .streamlit/secrets.toml"
    )
    st.stop()

st.title("📊 Case Assigner")

st.markdown(
    "Asigna los casos más antiguos respetando los grupos de Enterprise ID."
)

archivo = st.file_uploader(
    "Seleccione el archivo",
    type=[
        "xlsx",
        "xlsb"
    ]
)

st.subheader(
    "Analistas que recibirán casos"
)

agentes_seleccionados = []

columnas = st.columns(3)

lista_agentes = sorted(
    AGENTES.keys()
)

for idx, agente in enumerate(
    lista_agentes
):

    columna = columnas[
        idx % 3
    ]

    with columna:

        if st.checkbox(
            agente,
            key=f"agent_{idx}"
        ):

            agentes_seleccionados.append(
                agente
            )

casos = st.number_input(
    "Casos por persona",
    min_value=1,
    value=10
)

total_objetivo = (
    len(
        agentes_seleccionados
    )
    * int(casos)
)

col1, col2 = st.columns(2)

with col1:

    st.info(
        f"Analistas seleccionados: "
        f"{len(agentes_seleccionados)}"
    )

with col2:

    st.info(
        f"Capacidad objetivo: "
        f"{total_objetivo} casos"
    )

if st.button("Generar Asignación"):

    if archivo is None:

        st.error(
            "Seleccione un archivo."
        )

    elif len(
        agentes_seleccionados
    ) == 0:

        st.error(
            "Seleccione al menos un analista."
        )

    else:

        try:

            with st.spinner(
                "Procesando archivo..."
            ):

                excel_generado = (
                    generar_asignacion(
                        archivo,
                        agentes_seleccionados,
                        int(casos)
                    )
                )

            st.success(
                f"Asignación generada correctamente. "
                f"Se procesará una capacidad objetivo de "
                f"{total_objetivo} casos."
            )

            st.download_button(
                label="📥 Descargar Assignments.xlsx",
                data=excel_generado,
                file_name="Assignments.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as ex:

            st.error(
                str(ex)
            )