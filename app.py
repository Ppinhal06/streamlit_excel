import streamlit as st
import pandas as pd
import io
import json
import copy
from datetime import datetime
import google.generativeai as genai
from openpyxl import load_workbook

# 1. UI Configuration & API
st.set_page_config(page_title="BEMS Estimator Pro - DC Controls", layout="wide")

st.title("Automated BEMS Points List & Estimator")
st.markdown("---")
st.markdown("Professional generator with dynamic Excel template injection.")

# --- INICIALIZAR MEMORIA DE STREAMLIT ---
if "generado" not in st.session_state:
    st.session_state.generado = False

# --- PROFESSIONAL SIDEBAR ---
with st.sidebar:
    st.header("System Configuration")
    api_key = st.text_input("Enter API Key (Gemini):", type="password")
    
    st.markdown("---")
    st.info("✅ Standard DC Controls template is pre-loaded from the cloud server.")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash') 

# 2. Knowledge Base 
engineering_rules = {
    "Common LPHW/CHW Devices (System Level)": {
        "mandatory": ["Header Flow Immersion Temperature Sensor", "Header Return Immersion Temperature Sensor", "Outside Frost Thermostat", "Outside Temperature Sensor", "Immersion Frost Thermostat"]
    },
    "Boiler": {
        "per_unit": ["Boiler Enable", "Boiler Common Fault", "Boiler Control Signal", "Boiler Flow Immersion Temperature Sensor", "Boiler Return Immersion Temperature Sensor"]
    },
    "Pump (Primary/Secondary)": {
        "per_unit": ["Pump Enable", "Pump Status", "Variable Speed Drive", "Flow Immersion Temperature Sensor", "3 Port Control Valve / Actuator"]
    },
    "Pressurisation Unit": {
        "per_unit": ["Pressurisation Unit High Pressure", "Pressurisation Unit Low Pressure"]
    },
    "Calorifier / Hot Water Generator": {
        "per_unit": ["Enable", "Common Fault", "Control Signal", "Immersion Temperature Sensor", "High Limit Thermostat (60-95 Man. Reset)"]
    },
    "AHU (Air Handling Unit)": {
        "per_unit": ["Enable", "Status", "Control Signal", "Supply Air Temp Sensor", "Return Air Temp Sensor", "Frost Stat"]
    },
    "FCU (Fan Coil Unit)": {
        "per_unit": ["Space Temperature Sensor", "Control Valve Actuator"]
    }
}

# 3. Technical Catalog 
component_catalog = {
    "Header Flow Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Header Return Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Outside Frost Thermostat": {"Part": "DBET-23U", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Outside Temperature Sensor": {"Part": "TB/TO", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Immersion Frost Thermostat": {"Part": "DBTV-2U", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Boiler Enable": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Boiler Common Fault": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Boiler Control Signal": {"Part": "0…10V dc", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50},
    "Boiler Flow Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Boiler Return Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Pump Enable": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Pump Status": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Variable Speed Drive": {"Part": "Built in to Pump", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50},
    "Flow Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "3 Port Control Valve / Actuator": {"Part": "Valve 40mm", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50},
    "Pressurisation Unit High Pressure": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Pressurisation Unit Low Pressure": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Enable": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Common Fault": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Status": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Control Signal": {"Part": "0…10V dc", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50},
    "Immersion Temperature Sensor": {"Part": "TI/Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "High Limit Thermostat (60-95 Man. Reset)": {"Part": "RAK TW 1000B", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Space Temperature Sensor": {"Part": "RS-Temp", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Control Valve Actuator": {"Part": "MVC / DB_VZ", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50},
    "Supply Air Temp Sensor": {"Part": "Duct Temp Sensor", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Return Air Temp Sensor": {"Part": "Duct Temp Sensor", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Frost Stat": {"Part": "DBET-23U", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50}
}

# --- FUNCIÓN MAESTRA (ESTRUCTURA SEGURA Y LIGERA) ---
def crear_excel_formateado(datos, nombre_proyecto, es_io_schedule=False):
    try:
        wb = load_workbook("template.xlsx")
    except FileNotFoundError:
        st.error("⚠️ The 'template.xlsx' file is missing from the repository.")
        st.stop()
    
    sheet_name = "Points List" if "Points List" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    
    ws['B1'] = f"Project: {nombre_proyecto}"
    ws['C4'] = datetime.now().strftime("%d/%m/%Y")
    ws['C5'] = "Generated by AI Estimator - I/O Schedule Only" if es_io_schedule else "Generated by AI Estimator"
    
    start_row = 7 
    
    # Extraer estilo base una sola vez desde la primera celda descriptiva para no saturar el XML
    fuente_base = copy.copy(ws.cell(row=start_row, column=2).font)
    borde_base = copy.copy(ws.cell(row=start_row, column=2).border)
    alineacion_base = copy.copy(ws.cell(row=start_row, column=2).alignment)

    # Inyectar los datos
    for idx, row_data in enumerate(datos):
        current_row = start_row + idx
        col_map = {
            2: row_data.get("Description", ""), 3: row_data.get("AI", ""), 4: row_data.get("AO", ""),
            5: row_data.get("DI", ""), 6: row_data.get("DO", ""), 7: row_data.get("MCC", ""),
            8: row_data.get("Quantity", ""), 9: row_data.get("Part No.", ""), 10: row_data.get("Panel At 20%", ""),
            11: row_data.get("Parts At 0%", ""), 12: row_data.get("Labour At 20%", "")
        }
        
        for col_num, val in col_map.items():
            try:
                cell = ws.cell(row=current_row, column=col_num, value=val)
                # Aplicar el estilo base uniformemente
                cell.font = fuente_base
                cell.border = borde_base
                cell.alignment = alineacion_base
            except AttributeError:
                # Ignora silenciosamente si choca con una celda combinada preexistente
                pass

    # --- LIMPIEZA SEGURA (OCULTAR FILAS EN LUGAR DE BORRARLAS) ---
    if es_io_schedule:
        last_row = start_row + len(datos) - 1
        delete_start = last_row + 1
        
        # Ocultar todas las filas restantes previene la corrupción de celdas combinadas al final del documento
        for r in range(delete_start, ws.max_row + 1):
            ws.row_dimensions[r].hidden = True

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def tiene_puntos(row):
    for io_type in ["AI", "AO", "DI", "DO"]:
        val = str(row.get(io_type, "")).strip()
        if val and val.replace('.', '', 1).isdigit() and float(val) > 0:
            return True
    return False

st.subheader("Project Details")
col1, col2 = st.columns([1, 2])
with col1:
    project_name = st.text_input("Project Name (Header):", placeholder="e.g. IDA Cavan - Building 2")
with col2:
    project_description = st.text_area(
        "Project Scope Description:", 
        placeholder="Example: We need a plant with 2 Boilers, 3 Primary Pumps, 2 LPHW Pressurisation Units, and 1 Calorifier."
    )

st.markdown("---")

# 4. Generation Logic
if st.button("Generate Points List"):
    if not api_key or not project_description:
        st.warning("Please ensure you have entered your API Key and provided a project description.")
    else:
        with st.spinner("Engineering system points and cloning corporate styles..."):
            
            prompt = f"""
            You are an expert BEMS estimator working for DC Controls. Generate a Points List based on the description, mimicking the exact style of the "IDA Cavan" project.
            
            ENGINEERING RULES:
            {json.dumps(engineering_rules, indent=2)}
            
            TECHNICAL CATALOG:
            {json.dumps(component_catalog, indent=2)}
            
            CRITICAL FORMATTING INSTRUCTIONS:
            1. Create a HEADER ROW for each main equipment group. (e.g., Description: "Boiler", Quantity: 2, MCC: "MCB". Leave AI, AO, DI, DO, Labour blank).
            2. Below the header row, list its components based on the ENGINEERING RULES.
            3. CRITICAL CALCULATION: For each component, multiply base AI, AO, DI, DO, and Labour by the main equipment quantity.
            4. Use the exact Part No. from the catalog.
            5. Leave IOs or Labour as empty strings ("") if the value is 0.
            
            User description: "{project_description}"
            
            Return ONLY a JSON array matching the standard columns. DO NOT use markdown. You MUST use these exact keys:
            [
              {{
                "Description": "Boiler", "AI": "", "AO": "", "DI": "", "DO": "", "MCC": "MCB", "Quantity": 2, "Part No.": "", "Panel At 20%": "", "Parts At 0%": "", "Labour At 20%": ""
              }},
              {{
                "Description": "Boiler Enable", "AI": "", "AO": "", "DI": "", "DO": 2, "MCC": "", "Quantity": "", "Part No.": "Volt Free Contacts", "Panel At 20%": "", "Parts At 0%": "", "Labour At 20%": 100
              }}
            ]
            """
            
            try:
                response = model.generate_content(prompt)
                json_text = response.text.strip().replace("```json", "").replace("```", "")
                materials_data = json.loads(json_text)
                
                # --- EXCEL 1: DOCUMENTO COMPLETO (Intacto) ---
                buffer_full = crear_excel_formateado(materials_data, project_name, es_io_schedule=False)

                # --- EXCEL 2: I/O SCHEDULE (Filtrado y con filas ocultas) ---
                datos_io = []
                header_temporal = None
                
                for row in materials_data:
                    desc = str(row.get("Description", "")).strip()
                    part = str(row.get("Part No.", "")).strip()
                    
                    es_header = bool(desc) and not bool(part) and not tiene_puntos(row)
                    
                    if es_header:
                        header_temporal = row
                    elif tiene_puntos(row):
                        if header_temporal:
                            datos_io.append(header_temporal)
                            header_temporal = None 
                        datos_io.append(row)

                buffer_filtrado = crear_excel_formateado(datos_io, project_name, es_io_schedule=True)
                
                # --- GUARDAR EN MEMORIA ---
                st.session_state.generado = True
                st.session_state.buffer_full = buffer_full
                st.session_state.buffer_filtrado = buffer_filtrado
                st.session_state.materials_data = materials_data
                st.session_state.nombre_archivo = f"Quotation_{project_name.replace(' ', '_')}.xlsx"
                st.session_state.nombre_io = f"IO_Schedule_{project_name.replace(' ', '_')}.xlsx"

            except Exception as e:
                st.error(f"Error processing template: {e}")

# --- MOSTRAR RESULTADOS DESDE LA MEMORIA ---
if st.session_state.generado:
    st.success("Documents generated successfully!")
    
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        st.download_button(
            label="📄 Download Official Quotation (Full)",
            data=st.session_state.buffer_full,
            file_name=st.session_state.nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    with col_btn2:
        st.download_button(
            label="🔌 Download I/O Points Only (>0)",
            data=st.session_state.buffer_filtrado,
            file_name=st.session_state.nombre_io,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    st.dataframe(pd.DataFrame(st.session_state.materials_data), use_container_width=True)
