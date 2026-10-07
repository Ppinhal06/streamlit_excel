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
    model = genai.GenerativeModel('gemini-3.6-flash') 

# 2. Base de Conocimiento Expandida
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
    },
    "Storage Tank (Cold Water / Mains)": {
        "per_unit": ["Tank Low Low Level Status", "Tank Section 1 Low Level", "Tank Section 1 High Level", "Tank Section 2 Low Level", "Tank Section 2 High Level", "Tank Immersion Temperature Sensor", "Solenoid Valve 40mm", "Ultrasonic Level Transmitter"]
    },
    "Chiller": {
        "per_unit": ["Chiller Enable", "Chiller Status", "Chiller Flow Switch", "Chiller Flow Temp Sensor", "Chiller Return Temp Sensor"]
    },
    "Extract Fan": {
        "per_unit": ["Fan Enable", "Fan Current Switch"]
    },
    "Metering": {
        "mandatory": ["Gas Meter Pulsed Input", "Water Meter Pulsed Input", "Electricity Meter Pulsed Input"]
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
    "Frost Stat": {"Part": "DBET-23U", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Tank Low Low Level Status": {"Part": "LL13 (3M Cable)", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Tank Section 1 Low Level": {"Part": "LL13 (3M Cable)", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Tank Section 1 High Level": {"Part": "LL13 (3M Cable)", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Tank Section 2 Low Level": {"Part": "LL13 (3M Cable)", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Tank Section 2 High Level": {"Part": "LL13 (3M Cable)", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Tank Immersion Temperature Sensor": {"Part": "TI/Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Solenoid Valve 40mm": {"Part": "Solenoid Valve 40mm / ZS50", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Ultrasonic Level Transmitter": {"Part": "LS-MC", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Chiller Enable": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Chiller Status": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Chiller Flow Switch": {"Part": "FS 541", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Chiller Flow Temp Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Chiller Return Temp Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Fan Enable": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Fan Current Switch": {"Part": "RIBXKTF", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Gas Meter Pulsed Input": {"Part": "Device By Others", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Water Meter Pulsed Input": {"Part": "Device By Others", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Electricity Meter Pulsed Input": {"Part": "Device By Others", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50}
}

def tiene_puntos(row):
    for io_type in ["AI", "AO", "DI", "DO"]:
        val = str(row.get(io_type, "")).strip()
        if val and val.replace('.', '', 1).isdigit() and float(val) > 0:
            return True
    return False

# --- FUNCIÓN MAESTRA DIVIDIDA (QUOTATION vs I/O SCHEDULE) ---
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

    if not es_io_schedule:
        # ========================================================
        # LÓGICA 1: COTIZACIÓN OFICIAL (Mapeo Inteligente)
        # Aquí NO tocamos descripciones ni fuentes. Solo llenamos los números.
        # ========================================================
        
        # 1. Encontrar dónde termina la tabla de equipos
        req_row = ws.max_row
        for r in range(start_row, ws.max_row + 1):
            val = str(ws.cell(row=r, column=2).value).strip()
            if val in ["BMS Requirements", "Summary"]:
                req_row = r
                break
                
        # 2. LIMPIEZA PROFUNDA: Borrar "1s" fantasma y basura vieja (Solo columnas numéricas)
        for r in range(start_row, req_row):
            for c in [3, 4, 5, 6, 7, 8, 10, 11, 12]:
                ws.cell(row=r, column=c).value = ""

        # 3. Agrupar la data de la IA por "Padre" (Ej: Boiler, Storage Tank)
        ai_groups = {}
        current_ai_parent = None
        for item in datos:
            desc = str(item.get("Description", "")).strip().lower()
            part = str(item.get("Part No.", "")).strip()
            
            if desc and not part and not tiene_puntos(item):
                current_ai_parent = desc
                ai_groups[current_ai_parent] = {"parent": item, "children": []}
            else:
                if current_ai_parent:
                    ai_groups[current_ai_parent]["children"].append(item)
                    
        # 4. Mapear e inyectar SOLO los números en las filas correspondientes de la plantilla
        current_tpl_parent = None
        for r in range(start_row, req_row):
            desc_cell = ws.cell(row=r, column=2).value
            part_cell = ws.cell(row=r, column=9).value
            if not desc_cell:
                continue
                
            desc_str = str(desc_cell).strip().lower()
            part_str = str(part_cell).strip().lower() if part_cell else ""
            
            if desc_str and not part_str: 
                # Es un Encabezado Padre
                current_tpl_parent = desc_str
                if current_tpl_parent in ai_groups:
                    parent_data = ai_groups[current_tpl_parent]["parent"]
                    ws.cell(row=r, column=7).value = parent_data.get("MCC", "")
                    ws.cell(row=r, column=8).value = parent_data.get("Quantity", "")
            else:
                # Es un Componente Hijo
                if current_tpl_parent in ai_groups:
                    for child in ai_groups[current_tpl_parent]["children"]:
                        if str(child.get("Description", "")).strip().lower() == desc_str:
                            ws.cell(row=r, column=3).value = child.get("AI", "")
                            ws.cell(row=r, column=4).value = child.get("AO", "")
                            ws.cell(row=r, column=5).value = child.get("DI", "")
                            ws.cell(row=r, column=6).value = child.get("DO", "")
                            ws.cell(row=r, column=12).value = child.get("Labour At 20%", "")
                            break # Una vez inyectado, saltamos al siguiente renglón
    else:
        # ========================================================
        # LÓGICA 2: I/O SCHEDULE (Secuencial y Recortado)
        # ========================================================
        fuente_base = copy.copy(ws.cell(row=start_row, column=2).font)
        fuente_header = copy.copy(fuente_base)
        fuente_header.bold = True
        fuente_item = copy.copy(fuente_base)
        fuente_item.bold = False 
        borde_base = copy.copy(ws.cell(row=start_row, column=2).border)
        alineacion_base = copy.copy(ws.cell(row=start_row, column=2).alignment)

        # Inyectar secuencialmente (Aquí sí sobreescribimos todo)
        for idx, row_data in enumerate(datos):
            current_row = start_row + idx
            desc = str(row_data.get("Description", "")).strip()
            part = str(row_data.get("Part No.", "")).strip()
            es_header = bool(desc) and not bool(part) and not tiene_puntos(row_data)

            col_map = {
                2: row_data.get("Description", ""), 3: row_data.get("AI", ""), 4: row_data.get("AO", ""),
                5: row_data.get("DI", ""), 6: row_data.get("DO", ""), 7: row_data.get("MCC", ""),
                8: row_data.get("Quantity", ""), 9: row_data.get("Part No.", ""), 10: row_data.get("Panel At 20%", ""),
                11: row_data.get("Parts At 0%", ""), 12: row_data.get("Labour At 20%", "")
            }
            for col_num, val in col_map.items():
                try:
                    cell = ws.cell(row=current_row, column=col_num, value=val)
                    cell.font = fuente_header if es_header else fuente_item
                    cell.border = borde_base
                    cell.alignment = alineacion_base
                except AttributeError:
                    pass

        # Recortar celdas sobrantes para dejarlo limpio
        last_row = start_row + len(datos) - 1
        delete_start = last_row + 1
        ws.print_area = ""
        delete_amount = ws.max_row - delete_start + 1
        if delete_amount > 0:
            for mcr in list(ws.merged_cells.ranges):
                if mcr.min_row >= delete_start:
                    ws.merged_cells.remove(mcr)
            ws.delete_rows(idx=delete_start, amount=delete_amount)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

st.subheader("Project Details")
col1, col2 = st.columns([1, 2])
with col1:
    project_name = st.text_input("Project Name (Header):", placeholder="e.g. IDA Cavan - Building 2")
with col2:
    project_description = st.text_area(
        "Project Scope Description:", 
        placeholder="Example: We need a plant with 2 Boilers, 3 Primary Pumps, 1 Storage Tank, and Metering."
    )

st.markdown("---")

# 4. Generation Logic
if st.button("Generate Points List"):
    if not api_key or not project_description:
        st.warning("Please ensure you have entered your API Key and provided a project description.")
    else:
        with st.spinner("Engineering system points and cloning corporate styles..."):
            
            prompt = f"""
            You are an expert BEMS estimator working for DC Controls. Generate a Points List based on the description.
            
            ENGINEERING RULES (Known Systems):
            {json.dumps(engineering_rules, indent=2)}
            
            TECHNICAL CATALOG:
            {json.dumps(component_catalog, indent=2)}
            
            CRITICAL FORMATTING INSTRUCTIONS:
            1. Create a HEADER ROW for each main equipment group. (e.g., Description: "Boiler", Quantity: 2, MCC: "MCB". Leave AI, AO, DI, DO, Labour blank).
            2. Below the header row, list its components based on the ENGINEERING RULES.
            3. CRITICAL RULE: For components (child rows), "Quantity" and "MCC" MUST be left completely blank (""). Only the HEADER ROW should contain the Quantity and MCC.
            4. IF A SYSTEM IS NOT IN THE RULES, infer standard BEMS components for it (Enable DO, Status DI, Fault DI) and use "Device By Others" or "Volt Free Contacts" as the Part No.
            5. CRITICAL CALCULATION: For each component, multiply base AI, AO, DI, DO, and Labour by the main equipment quantity.
            6. Use the exact Part No. from the catalog.
            7. Leave IOs or Labour as empty strings ("") if the value is 0.
            
            User description: "{project_description}"
            
            Return ONLY a JSON array matching the standard columns. DO NOT use markdown. You MUST use these exact keys:
            [
              {{
                "Description": "Storage Tank (Cold Water / Mains)", "AI": "", "AO": "", "DI": "", "DO": "", "MCC": "MCB", "Quantity": 1, "Part No.": "", "Panel At 20%": "", "Parts At 0%": "", "Labour At 20%": ""
              }}
            ]
            """
            
            try:
                response = model.generate_content(prompt)
                json_text = response.text.strip().replace("```json", "").replace("```", "")
                materials_data = json.loads(json_text)
                
                # --- EXCEL 1: DOCUMENTO COMPLETO (Mapeo Inteligente) ---
                buffer_full = crear_excel_formateado(materials_data, project_name, es_io_schedule=False)

                # --- EXCEL 2: I/O SCHEDULE (Cortado y Filtrado) ---
                datos_io = []
                headers_pendientes = []
                
                for row in materials_data:
                    desc = str(row.get("Description", "")).strip()
                    part = str(row.get("Part No.", "")).strip()
                    qty = str(row.get("Quantity", "")).strip()
                    
                    es_header = bool(desc) and not bool(part) and not tiene_puntos(row)
                    
                    if es_header:
                        if qty == "0" or qty == 0:
                            pass # Ignorar encabezados vacíos
                        else:
                            headers_pendientes.append(row)
                    elif tiene_puntos(row):
                        for h in headers_pendientes:
                            datos_io.append(h)
                        headers_pendientes = [] 
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
            label="🔌 Download I/O Schedule Only (>0)",
            data=st.session_state.buffer_filtrado,
            file_name=st.session_state.nombre_io,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
