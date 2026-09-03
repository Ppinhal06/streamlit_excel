import streamlit as st
import pandas as pd
import io
import json
import google.generativeai as genai
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# 1. UI Configuration & API
st.set_page_config(page_title="BEMS Estimator - DC Controls", layout="wide")
st.title("Automated BEMS Points List & Estimator")
st.markdown("Generador basado en el estándar de cotización de DC Controls (Formato IDA Cavan).")

api_key = st.text_input("Enter API Key (Gemini/OpenAI):", type="password")
if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-3.6-flash') 

# 2. Knowledge Base (Real rules extracted from IDA Cavan Project)
engineering_rules = {
    "Common LPHW/CHW Devices (System Level)": {
        "mandatory": [
            "Header Flow Immersion Temperature Sensor",
            "Header Return Immersion Temperature Sensor",
            "Outside Frost Thermostat",
            "Outside Temperature Sensor",
            "Immersion Frost Thermostat"
        ],
        "notes": "Include these ONLY ONCE per plant for the main headers."
    },
    "Boiler": {
        "per_unit": [
            "Boiler Enable",
            "Boiler Common Fault",
            "Boiler Control Signal",
            "Boiler Flow Immersion Temperature Sensor",
            "Boiler Return Immersion Temperature Sensor"
        ]
    },
    "Pump (Primary/Secondary)": {
        "per_unit": [
            "Pump Enable",
            "Pump Status",
            "Variable Speed Drive",
            "Flow Immersion Temperature Sensor",
            "3 Port Control Valve / Actuator"
        ]
    },
    "Pressurisation Unit": {
        "per_unit": [
            "Pressurisation Unit High Pressure",
            "Pressurisation Unit Low Pressure"
        ]
    },
    "Calorifier / Hot Water Generator": {
        "per_unit": [
            "Enable",
            "Common Fault",
            "Control Signal",
            "Immersion Temperature Sensor",
            "High Limit Thermostat (60-95 Man. Reset)"
        ]
    },
    "AHU (Air Handling Unit)": {
        "per_unit": [
            "Enable", "Status", "Control Signal", "Supply Air Temp Sensor", "Return Air Temp Sensor", "Frost Stat"
        ]
    },
    "FCU (Fan Coil Unit)": {
        "per_unit": [
            "Space Temperature Sensor", "Control Valve Actuator"
        ]
    }
}

# 3. Technical Catalog (Exact extraction from IDA Cavan)
# Base Labour is 50 per point. The AI will multiply this by the quantity.
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

project_description = st.text_area(
    "Describe the project for the Points List:", 
    placeholder="Example: We need a plant with 2 Boilers, 3 Primary Pumps, 2 LPHW Pressurisation Units, and 1 Calorifier."
)

# 4. Generation Logic
if st.button("Generate Points List (Excel)"):
    if not api_key or not project_description:
        st.warning("Please provide both the API Key and the project description.")
    else:
        with st.spinner("Engineering Points List, assigning real IOs, Part Numbers and calculating Labour..."):
            
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
               (e.g., If Boiler Qty is 2, and "Boiler Flow Immersion Temp Sensor" has 1 AI and 50 Labour, the output MUST be AI: 2, Quantity: 2, Labour At 20%: 100).
            4. Use the exact Part No. from the catalog (e.g., "Volt Free Contacts", "0...10V dc", "TTI-S Brass Pocket").
            5. Leave IOs or Labour as empty strings ("") if the value is 0.
            
            User description: "{project_description}"
            
            Return ONLY a JSON array matching the standard DC Controls Excel columns:
            [
              {{
                "Description": "Boiler", "AI": "", "AO": "", "DI": "", "DO": "", "MCC": "MCB", "Quantity": 2, "Part No.": "", "Panel At 20%": "", "Parts At 0%": "", "Labour At 20%": ""
              }},
              {{
                "Description": "Boiler Flow Immersion Temperature Sensor", "AI": 2, "AO": "", "DI": "", "DO": "", "MCC": "", "Quantity": 2, "Part No.": "TTI-S Brass Pocket", "Panel At 20%": "", "Parts At 0%": "", "Labour At 20%": 100
              }}
            ]
            """
            
            try:
                response = model.generate_content(prompt)
                json_text = response.text.strip().replace("```json", "").replace("```", "")
                materials_data = json.loads(json_text)
                
                df = pd.DataFrame(materials_data)
                
                # Reorder to match exact Excel template if possible
                expected_columns = ['Description', 'AI', 'AO', 'DI', 'DO', 'MCC', 'Quantity', 'Part No.', 'Panel At 20%', 'Parts At 0%', 'Labour At 20%']
                for col in expected_columns:
                    if col not in df.columns:
                        df[col] = ""
                df = df[expected_columns]

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Points List')
                    worksheet = writer.sheets['Points List']
                    
                    for col_num, col in enumerate(df.columns, 1):
                        max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                        col_letter = get_column_letter(col_num)
                        worksheet.column_dimensions[col_letter].width = max_len
                    
                    max_row = len(df) + 1
                    max_col = len(df.columns)
                    table_ref = f"A1:{get_column_letter(max_col)}{max_row}"
                    
                    tab = Table(displayName="PointsList_Table", ref=table_ref)
                    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                    tab.tableStyleInfo = style
                    worksheet.add_table(tab)

                st.success("Points List generated successfully!")
                st.download_button(
                    label="📥 Download Points List (Excel)",
                    data=buffer.getvalue(),
                    file_name="DC_Controls_Points_List_Priced.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.dataframe(df, use_container_width=True)

            except Exception as e:
                st.error(f"Error: {e}\nEnsure the AI returned a valid JSON format.")
