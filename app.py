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
st.markdown("Generador Profesional con Inyección Dinámica de Plantillas Excel.")

# --- SIDEBAR PROFESIONAL ---
with st.sidebar:
    st.header("⚙️ Configuración")
    api_key = st.text_input("Enter API Key (Gemini):", type="password")
    st.markdown("---")
    st.subheader("📁 Plantilla Corporativa")
    uploaded_template = st.file_uploader("Sube aquí el archivo '0001 DC Controls Points List Template...xlsx'", type=["xlsx"])
    st.info("💡 Al subir tu plantilla, la IA preservará tu logo, rellenará los headers (Project, Date) y clonará las fuentes y bordes exactos de la empresa.")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-3.6-flash') 

# 2. Knowledge Base (Real rules extracted from IDA Cavan Project)
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

# 3. Technical Catalog (Base Labour is 50 per point)
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

col1, col2 = st.columns([1, 2])
with col1:
    project_name = st.text_input("Project Name (Header):", placeholder="Ej: IDA Cavan - Building 2")
with col2:
    project_description = st.text_area(
        "Describe the project for the Points List:", 
        placeholder="Example: We need a plant with 2 Boilers, 3 Primary Pumps..."
    )

# 4. Generation Logic
if st.button("🚀 Generate Pro Points List"):
    if not api_key or not project_description or uploaded_template is None:
        st.warning("⚠️ Asegúrate de ingresar tu API Key, subir la plantilla en la barra lateral y escribir la descripción.")
    else:
        with st.spinner("Inyectando ingeniería y clonando estilos corporativos..."):
            
            prompt = f"""
            You are an expert BEMS estimator working for DC Controls. Generate a Points List based on the description, mimicking the exact style of the "IDA Cavan" project.
            
            ENGINEERING RULES (What components go in each system):
            {json.dumps(engineering_rules, indent=2)}
            
            TECHNICAL CATALOG (Exact Parts and I/O values):
            {json.dumps(component_catalog, indent=2)}
            
            CRITICAL FORMATTING INSTRUCTIONS (CAVAN PROJECT STYLE):
            1. Create a HEADER ROW for each main equipment group. (e.g., Description: "Boiler", Quantity: 2, MCC: "MCB". Leave AI, AO, DI, DO, Labour blank).
            2. Below the header row, list its components based on the ENGINEERING RULES.
            3. CRITICAL CALCULATION: For each component, lookup its base AI, AO, DI, DO, and Labour in the CATALOG. You MUST multiply these base values by the quantity of the main equipment.
            4. Use the exact Part No. from the catalog (e.g., "Volt Free Contacts", "0...10V dc", "TTI-S Brass Pocket").
            5. Leave IOs or Labour as empty strings ("") if the value is 0.
            
            User description: "{project_description}"
            
            Return ONLY a JSON array matching the standard DC Controls columns. DO NOT use markdown.
            [
              {{
                "Description": "Boiler", "AI": "", "AO": "", "DI": "", "DO": "", "MCC": "MCB", "Quantity": 2, "Part No.": "", "Panel At 20%": "", "Parts At 0%": "", "Labour At 20%": ""
              }}
            ]
            """
            
            try:
                response = model.generate_content(prompt)
                json_text = response.text.strip().replace("```json", "").replace("```", "")
                materials_data = json.loads(json_text)
                
                # --- MAGIA DE CLONACIÓN DE ESTILOS ---
                wb = load_workbook(uploaded_template)
                
                # Buscar la hoja correcta
                sheet_name = "Points List" if "Points List" in wb.sheetnames else wb.sheetnames[0]
                ws = wb[sheet_name]
                
                # Inyectar Metadata Automática en el Header de la empresa
                ws['C1'] = project_name
                ws['C4'] = datetime.now().strftime("%d/%m/%Y")
                ws['C5'] = "Generado por IA Estimator"
                
                start_row = 7 # Aquí inician tus datos reales
                
                # 1. CLONAR ESTILOS de la fila 7 para replicarlos
                reference_styles = {}
                for col in range(2, 13): # Columnas B a L
                    cell = ws.cell(row=start_row, column=col)
                    reference_styles[col] = {
                        'font': copy.copy(cell.font),
                        'border': copy.copy(cell.border),
                        'fill': copy.copy(cell.fill),
                        'alignment': copy.copy(cell.alignment)
                    }

                # 2. INYECTAR DATOS y APLICAR ESTILOS
                for idx, row_data in enumerate(materials_data):
                    current_row = start_row + idx
                    
                    col_map = {
                        2: row_data.get("Description", ""), 3: row_data.get("AI", ""), 4: row_data.get("AO", ""),
                        5: row_data.get("DI", ""), 6: row_data.get("DO", ""), 7: row_data.get("MCC", ""),
                        8: row_data.get("Quantity", ""), 9: row_data.get("Part No.", ""), 10: row_data.get("Panel At 20%", ""),
                        11: row_data.get("Parts At 0%", ""), 12: row_data.get("Labour At 20%", "")
                    }
                    
                    for col_num, val in col_map.items():
                        cell = ws.cell(row=current_row, column=col_num, value=val)
                        
                        # Pegarle el estilo clonado a cada celda nueva
                        if col_num in reference_styles:
                            cell.font = copy.copy(reference_styles[col_num]['font'])
                            cell.border = copy.copy(reference_styles[col_num]['border'])
                            cell.fill = copy.copy(reference_styles[col_num]['fill'])
                            cell.alignment = copy.copy(reference_styles[col_num]['alignment'])

                buffer = io.BytesIO()
                wb.save(buffer)

                st.success(f"¡Cotización para '{project_name}' generada con éxito!")
                st.download_button(
                    label="📥 Descargar Points List Oficial (Excel)",
                    data=buffer.getvalue(),
                    file_name=f"Quotation_{project_name.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.dataframe(pd.DataFrame(materials_data), use_container_width=True)

            except Exception as e:
                st.error(f"Error procesando la plantilla: {e}")
