import streamlit as st
from cohort_engine import generar_asignacion

st.set_page_config(
    page_title="Case Assigner",
    layout="wide"
)

st.title("📊 Case Assigner")

archivo = st.file_uploader(
    "Seleccione el archivo",
    type=[
        "xlsx",
        "xlsb"
    ]
)

personas = st.number_input(
    "Cantidad de personas",
    min_value=1,
    value=15
)

casos = st.number_input(
    "Casos por persona",
    min_value=1,
    value=10
)

if st.button("Generar Asignación"):

    if archivo is None:

        st.error(
            "Seleccione un archivo."
        )

    else:

        try:

            with st.spinner(
                "Procesando archivo..."
            ):

                excel_generado = (
                    generar_asignacion(
                        archivo,
                        int(personas),
                        int(casos)
                    )
                )

            st.success(
                "Asignación generada correctamente."
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