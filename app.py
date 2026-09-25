
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Centro de Control de Proyectos", layout="wide")

CATEGORIAS_PESO = {
    "Oferta": 10, "Compras": 15, "Project Manager": 10,
    "Ejecucion": 40, "Pruebas": 10, "Documentacion": 10, "Objetivos": 5
}

CHECKLIST_ESTANDAR = [
    ("Alcance", "PLC", 0), ("Alcance", "Offline PLC", 0), ("Alcance", "Fabricacion Armario", 0),
    ("Alcance", "Fabricacion Cajas", 0), ("Alcance", "Fabricacion Etiquetas", 0), ("Alcance", "Especialidad Eplan", 0),
    ("Oferta", "Pedido cliente recibido", 3), ("Oferta", "Alcance revisado", 3),
    ("Oferta", "Material instalacion identificado", 2), ("Oferta", "Proveedor definido", 2),
    ("Compras", "Pedido emitido", 5), ("Compras", "Confirmacion proveedor", 3),
    ("Compras", "Material recibido", 5), ("Compras", "Material revisado", 1), ("Compras", "Material completo", 1),
    ("Project Manager", "Carpeta creada", 2), ("Project Manager", "Kick Off realizado", 2),
    ("Project Manager", "Planificacion realizada", 3), ("Project Manager", "Recursos asignados", 1),
    ("Project Manager", "Seguimiento actualizado", 2),
    ("Ejecucion", "Ingenieria completada", 8), ("Ejecucion", "Fabricacion iniciada", 4),
    ("Ejecucion", "Fabricacion terminada", 8), ("Ejecucion", "PLC completado", 8),
    ("Ejecucion", "Instalacion iniciada", 4), ("Ejecucion", "Instalacion terminada", 8),
    ("Pruebas", "FAT realizada", 5), ("Pruebas", "FAT aprobada", 3),
    ("Pruebas", "SAT realizada", 1), ("Pruebas", "SAT aprobada", 1),
    ("Documentacion", "Backup PLC", 2), ("Documentacion", "Eplan final", 2),
    ("Documentacion", "Carpeta Eplan", 1), ("Documentacion", "Manuales", 2),
    ("Documentacion", "Lista materiales", 2), ("Documentacion", "As Built", 1),
    ("Objetivos", "Requisito 1 cumplido", 2), ("Objetivos", "Requisito 2 cumplido", 2),
    ("Objetivos", "Requisito 3 cumplido", 1),
]

ESTADOS = ["No iniciada", "En curso", "Bloqueada", "Completada", "No aplica"]
ESTADO_ICONO = {"Completada": "\U0001F7E2", "En curso": "\U0001F7E1", "Bloqueada": "\U0001F534",
                 "No iniciada": "\u2b1c", "No aplica": "\u26AB"}

for key, default in [("proyectos", None), ("acciones", None), ("pendientes", None)]:
    if key not in st.session_state:
        st.session_state[key] = default


def generar_checklist_para_proyecto(id_proyecto):
    filas = []
    for cat, accion, peso in CHECKLIST_ESTANDAR:
        filas.append({
            "ID_Accion": f"{id_proyecto}-{accion.replace(' ', '_')}",
            "ID_Proyecto": id_proyecto, "Categoria": cat, "Accion": accion,
            "Tipo_Check": "Estandar", "Peso": peso, "Estado": "No iniciada",
            "Responsable": "", "Fecha_Actualizacion": "", "Comentario": ""
        })
    return pd.DataFrame(filas)


def calcular_avance(df_acciones, id_proyecto):
    sub = df_acciones[(df_acciones["ID_Proyecto"] == id_proyecto) & (df_acciones["Categoria"] != "Alcance")]
    peso_total = sub["Peso"].sum()
    peso_completado = sub.loc[sub["Estado"] == "Completada", "Peso"].sum()
    avance = peso_completado / peso_total if peso_total > 0 else 0
    if avance >= 0.8:
        semaforo = "\U0001F7E2 Correcto"
    elif avance >= 0.5:
        semaforo = "\U0001F7E1 Atencion"
    else:
        semaforo = "\U0001F534 Critico"
    return round(avance, 4), semaforo


def cargar_excel_origen(file):
    xls = pd.ExcelFile(file)
    proyectos, acciones, pendientes = None, None, None
    for sheet in xls.sheet_names:
        low = sheet.lower()
        if "proyecto" in low and "bd" in low:
            proyectos = pd.read_excel(xls, sheet_name=sheet)
        elif "accion" in low or "checklist" in low:
            acciones = pd.read_excel(xls, sheet_name=sheet)
        elif "pendiente" in low:
            pendientes = pd.read_excel(xls, sheet_name=sheet)
    if proyectos is None and len(xls.sheet_names) >= 1:
        proyectos = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    return proyectos, acciones, pendientes


def exportar_a_excel(df_proyectos, df_acciones, df_pendientes):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_proyectos.to_excel(writer, sheet_name="02_BD_PROYECTOS", index=False)
        df_acciones.to_excel(writer, sheet_name="03_BD_ACCIONES", index=False)
        df_pendientes.to_excel(writer, sheet_name="04_BD_PENDIENTES", index=False)
        resumen = df_proyectos[["ID_Proyecto", "Nombre", "Avance_Operativo", "Semaforo"]].copy() \
            if "Nombre" in df_proyectos.columns else df_proyectos.copy()
        resumen.to_excel(writer, sheet_name="01_DASHBOARD", index=False)
    output.seek(0)
    return output


st.title("Centro de Control de Proyectos")
st.caption("Sube tus datos actualizados, revisa el estado, marca el checklist, y exporta de vuelta a Excel.")

tab1, tab2, tab3, tab4 = st.tabs(["Cargar datos", "Dashboard", "Ficha de Proyecto", "Exportar a Excel"])

with tab1:
    st.subheader("Cargar tu Excel o CSV actualizado")
    st.write("Sube tu `Centro_Control_Proyectos.xlsx` (o los CSV sueltos) para trabajar con tus datos reales.")

    modo = st.radio("Como quieres cargar los datos", ["Subir Excel completo", "Subir CSV por separado", "Usar datos de ejemplo"])

    if modo == "Subir Excel completo":
        file = st.file_uploader("Sube tu archivo .xlsx", type=["xlsx"])
        if file is not None:
            proyectos, acciones, pendientes = cargar_excel_origen(file)
            if proyectos is not None:
                st.session_state["proyectos"] = proyectos
            if acciones is not None:
                st.session_state["acciones"] = acciones
            if pendientes is not None:
                st.session_state["pendientes"] = pendientes
            st.success("Datos cargados correctamente. Ve a la pestana Dashboard.")

    elif modo == "Subir CSV por separado":
        c1, c2, c3 = st.columns(3)
        with c1:
            f1 = st.file_uploader("proyectos.csv", type=["csv"], key="p")
            if f1:
                st.session_state["proyectos"] = pd.read_csv(f1, sep=None, engine="python")
        with c2:
            f2 = st.file_uploader("acciones.csv (checklist)", type=["csv"], key="a")
            if f2:
                st.session_state["acciones"] = pd.read_csv(f2, sep=None, engine="python")
        with c3:
            f3 = st.file_uploader("pendientes.csv", type=["csv"], key="pe")
            if f3:
                st.session_state["pendientes"] = pd.read_csv(f3, sep=None, engine="python")
        if f1 or f2 or f3:
            st.success("CSV cargados. Ve a la pestana Dashboard.")

    else:
        proyectos_demo = pd.DataFrame([
            {"ID_Proyecto": "PRJ-1036-25", "Nombre": "WBS", "Cliente": "EISENMANN", "PM": "Xavier Arevalo",
             "Estado": "Activo", "Importe_Oferta": 250000},
            {"ID_Proyecto": "PRJ-1182-26", "Nombre": "DESMANTELAMIENTO_ELECTRICO_P2", "Cliente": "EBRO FACTORY",
             "PM": "Marcos Gimenez", "Estado": "Activo", "Importe_Oferta": 224724},
        ])
        acciones_demo = pd.concat([
            generar_checklist_para_proyecto("PRJ-1036-25"),
            generar_checklist_para_proyecto("PRJ-1182-26"),
        ], ignore_index=True)
        acciones_demo.loc[acciones_demo["ID_Proyecto"] == "PRJ-1036-25", "Estado"] = \
            np.random.choice(ESTADOS, size=(acciones_demo["ID_Proyecto"] == "PRJ-1036-25").sum())
        pendientes_demo = pd.DataFrame([
            {"ID_Pendiente": "PEN-001", "ID_Proyecto": "PRJ-1036-25", "Descripcion": "Falta confirmar proveedor de armarios",
             "Estado": "Abierto", "Prioridad": "Alta", "Responsable": "Xavier Arevalo"},
        ])
        st.session_state["proyectos"] = proyectos_demo
        st.session_state["acciones"] = acciones_demo
        st.session_state["pendientes"] = pendientes_demo
        st.success("Datos de ejemplo cargados.")

    st.divider()
    st.subheader("Dar de alta un proyecto nuevo (genera su checklist automaticamente)")
    with st.form("nuevo_proyecto"):
        nid = st.text_input("ID de proyecto (formato PRJ-NNNN-AA)")
        nnombre = st.text_input("Nombre del proyecto")
        ncliente = st.text_input("Cliente")
        npm = st.text_input("Project Manager")
        submit = st.form_submit_button("Crear proyecto y checklist estandar (39 acciones)")
        if submit and nid:
            nueva_fila = pd.DataFrame([{"ID_Proyecto": nid, "Nombre": nnombre, "Cliente": ncliente,
                                         "PM": npm, "Estado": "Activo"}])
            if st.session_state["proyectos"] is None:
                st.session_state["proyectos"] = nueva_fila
            else:
                st.session_state["proyectos"] = pd.concat([st.session_state["proyectos"], nueva_fila], ignore_index=True)
            nuevo_checklist = generar_checklist_para_proyecto(nid)
            if st.session_state["acciones"] is None:
                st.session_state["acciones"] = nuevo_checklist
            else:
                st.session_state["acciones"] = pd.concat([st.session_state["acciones"], nuevo_checklist], ignore_index=True)
            st.success(f"Proyecto {nid} creado con 39 acciones de checklist.")

with tab2:
    st.subheader("Dashboard")
    if st.session_state["proyectos"] is None:
        st.info("Carga datos primero en la pestana 'Cargar datos'.")
    else:
        proyectos = st.session_state["proyectos"].copy()
        acciones = st.session_state["acciones"] if st.session_state["acciones"] is not None else pd.DataFrame(
            columns=["ID_Proyecto", "Categoria", "Peso", "Estado"])
        pendientes = st.session_state["pendientes"] if st.session_state["pendientes"] is not None else pd.DataFrame(
            columns=["ID_Proyecto", "Estado"])

        avances, semaforos, pend_abiertos = [], [], []
        for pid in proyectos["ID_Proyecto"]:
            av, sem = calcular_avance(acciones, pid) if not acciones.empty else (0, "\u26AA Sin datos")
            avances.append(av)
            semaforos.append(sem)
            if not pendientes.empty:
                ab = pendientes[(pendientes["ID_Proyecto"] == pid) &
                                (pendientes["Estado"].isin(["Abierto", "En progreso", "Bloqueado"]))].shape[0]
            else:
                ab = 0
            pend_abiertos.append(ab)

        proyectos["Avance_Operativo"] = avances
        proyectos["Semaforo"] = semaforos
        proyectos["Pendientes_Abiertos"] = pend_abiertos
        st.session_state["proyectos"] = proyectos

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Proyectos totales", len(proyectos))
        c2.metric("Activos", (proyectos.get("Estado", pd.Series(dtype=str)) == "Activo").sum())
        c3.metric("Avance medio", f"{proyectos['Avance_Operativo'].mean()*100:.0f}%")
        c4.metric("Pendientes abiertos", int(proyectos["Pendientes_Abiertos"].sum()))

        st.dataframe(
            proyectos[["ID_Proyecto"] + [c for c in ["Nombre", "Cliente", "PM", "Estado"] if c in proyectos.columns] +
                      ["Avance_Operativo", "Semaforo", "Pendientes_Abiertos"]],
            use_container_width=True,
            column_config={"Avance_Operativo": st.column_config.ProgressColumn(
                "Avance", min_value=0, max_value=1, format="%.0f%%")}
        )

with tab3:
    st.subheader("Ficha de Proyecto")
    if st.session_state["proyectos"] is None:
        st.info("Carga datos primero.")
    else:
        proyectos = st.session_state["proyectos"]
        pid = st.selectbox("Selecciona un proyecto", proyectos["ID_Proyecto"].unique())
        fila = proyectos[proyectos["ID_Proyecto"] == pid].iloc[0]

        acciones = st.session_state["acciones"]
        pendientes = st.session_state["pendientes"]

        av, sem = calcular_avance(acciones, pid) if acciones is not None else (0, "Sin datos")
        c1, c2 = st.columns([1, 2])
        with c1:
            st.metric("Avance operativo", f"{av*100:.0f}%")
            st.write(f"**Semaforo:** {sem}")
            for campo in ["Nombre", "Cliente", "PM", "Estado", "Importe_Oferta"]:
                if campo in fila:
                    st.write(f"**{campo}:** {fila[campo]}")

        with c2:
            st.write("### Checklist")
            if acciones is not None:
                sub = acciones[acciones["ID_Proyecto"] == pid].copy()
                for idx, row in sub.iterrows():
                    cols = st.columns([3, 2, 2])
                    cols[0].write(f"{ESTADO_ICONO.get(row['Estado'], '')} **{row['Accion']}** ({row['Categoria']}, peso {row['Peso']})")
                    nuevo_estado = cols[1].selectbox("Estado", ESTADOS, index=ESTADOS.index(row["Estado"]) if row["Estado"] in ESTADOS else 0,
                                                      key=f"estado_{pid}_{idx}", label_visibility="collapsed")
                    if nuevo_estado != row["Estado"]:
                        acciones.loc[idx, "Estado"] = nuevo_estado
                        acciones.loc[idx, "Fecha_Actualizacion"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                        st.session_state["acciones"] = acciones
                        st.rerun()

            st.write("### Pendientes")
            if pendientes is not None:
                sub_p = pendientes[pendientes["ID_Proyecto"] == pid]
                st.dataframe(sub_p, use_container_width=True)

with tab4:
    st.subheader("Exportar de vuelta a Excel")
    st.write("Descarga un `.xlsx` con tus datos actualizados (proyectos, checklist y pendientes) para integrarlo en tu Centro de Control de Proyectos.")
    if st.session_state["proyectos"] is None:
        st.info("Carga y trabaja con tus datos primero.")
    else:
        proyectos = st.session_state["proyectos"]
        acciones = st.session_state["acciones"] if st.session_state["acciones"] is not None else pd.DataFrame()
        pendientes = st.session_state["pendientes"] if st.session_state["pendientes"] is not None else pd.DataFrame()
        excel_bytes = exportar_a_excel(proyectos, acciones, pendientes)
        st.download_button(
            label="Descargar Excel actualizado",
            data=excel_bytes,
            file_name=f"Centro_Control_Proyectos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.caption("Este archivo trae 4 hojas: 01_DASHBOARD, 02_BD_PROYECTOS, 03_BD_ACCIONES, 04_BD_PENDIENTES, listas para copiar/pegar o reemplazar en tu archivo maestro.")
