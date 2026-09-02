import streamlit as st
import pandas as pd
import io
import json
import google.generativeai as genai
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# 1. UI Configuration & API
st.set_page_config(page_title="BOM Generator - DC Controls", layout="centered")
st.title("Automated BEMS Component Generator")

api_key = st.text_input("Enter API Key (Gemini/OpenAI):", type="password")
if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-3.6-flash') 

# 2. Knowledge Base (Strict Engineering Rules based on DC Controls diagrams)
engineering_rules = {
    "Boiler System": {
        "system_level_components": [
            "Outside Temperature Sensor",
            "Outside Frost Stat",
            "Fire Alarm Interface Relay",
            "Boiler Flow Temperature Sensor",
            "Boiler Return Temperature Sensor",
            "Immersion Frost Stat"
        ],
        "per_boiler_components": [
            "Boiler Enable Relay",
            "Boiler Lockout Monitor"
        ],
        "notes": "Add ONE set of system-level components per plant. Multiply per-boiler components by the total number of boilers."
    },
    "Heat Pump System": {
        "system_level_components": [
            "Outside Temperature Sensor",
            "Outside Frost Stat",
            "Fire Alarm Interface Relay"
        ],
        "notes": "Primary pumps are built into the HP unit. Do not add external pump components for primary flow. System-level components added once per plant."
    },
    "Buffer Tank (LPHW or CHW)": {
        "mandatory_components": [
            "Flow Temperature Sensor",
            "Return Temperature Sensor",
            "Buffer High Temperature Sensor",
            "Buffer Low Temperature Sensor"
        ],
        "notes": "Applies to both Hot Water (LPHW) and Chilled Water (CHW) buffer cylinders."
    },
    "Pressurization Unit": {
        "mandatory_components": [
            "High Pressure Switch",
            "Low Pressure Switch"
        ],
        "notes": "If it is an LPHW unit, add 1x Immersion Frost Stat."
    },
    "Secondary Circuit (LPHW or CHW)": {
        "mandatory_components": [
            "Secondary Flow Temperature Sensor",
            "Secondary Return Temperature Sensor",
            "Pump Flow Switch", 
            "Pump Solid State Relay"
        ],
        "notes": "Applies to AHU, FCU, or DHW secondary pump circuits. Multiply by the number of circuits requested."
    },
    "AHU (Air Handling Unit)": {
        "mandatory_components": [
            "Outside Temperature Sensor", 
            "Room Temperature Sensor",
            "Supply Temperature Sensor",
            "Fresh Air Damper Actuator", 
            "Panel Filter DP Switch", 
            "Bag Filter DP Switch", 
            "Frost Stat", 
            "Heating Valve Actuator", 
            "Cooling Valve Actuator", 
            "Supply Fan VSD",
            "Fire Alarm Interface Relay"
        ],
        "notes": "Standard configuration based on Restaurant Ventilation schematic."
    },
    "FCU (Fan Coil Unit)": {
        "mandatory_components": [
            "Space Temperature Sensor", 
            "Space Humidity Sensor", 
            "Space CO2 Sensor", 
            "Control Valve Actuator"
        ],
        "notes": "Standard bedroom/zone configuration."
    },
    "Water Tank System (General)": {
        "system_level_components": [
            "Outside Temperature Sensor",
            "Outside Frost Stat",
            "Fire Alarm Interface Relay"
        ],
        "notes": "Add these ONLY ONCE if the project contains any Mains or Fire Water Tanks."
    },
    "Mains Water Tank": {
        "mandatory_components": [
            "High Level Sensor (x2)",
            "Low Level Sensor (x2)",
            "Tank Temperature Sensor",
            "Boiler/Booster Status Monitor"
        ],
        "notes": "Standard domestic mains tank. Multiplies by quantity of tanks."
    },
    "Fire Water Tank": {
        "mandatory_components": [
            "High Level Sensor",
            "Low Level Sensor",
            "Tank Temperature Sensor",
            "Boiler/Booster Status Monitor"
        ],
        "notes": "Standard fire reserve tank. Multiplies by quantity of tanks."
    }
}

project_description = st.text_area(
    "Describe the circuit equipment:", 
    placeholder="Example: We need to automate a project with: 2 Boilers, 3 LPHW Secondary Circuits, 2 CHW Secondary Circuits..."
)

# 3. Generation Logic
if st.button("Generate Bill of Materials (Excel)"):
    if not api_key or not project_description:
        st.warning("Please provide both the API Key and the project description.")
    else:
        with st.spinner("Analyzing requirements and applying engineering rules..."):
            
            prompt = f"""
            You are an expert BEMS engineer. Analyze the following circuit description and generate a Bill of Materials (BOM).
            
            STRICT ENGINEERING RULES:
            {json.dumps(engineering_rules, indent=2)}
            
            You must apply these rules mandatorily. 
            - If the user mentions "AHU", "FCU", "Secondary Circuit" (or LPHW/CHW circuits), "Buffer Tank", "Pressurization Unit", "Mains Water Tank", or "Fire Water Tank", add all mandatory components multiplied by the quantity requested.
            - If the user mentions "Boilers", "Heat Pump", or any "Water Tanks", apply the system rules carefully: add the system-level components only once per plant/system, and multiply the per-unit components by the number of units requested. Avoid duplicating system-level components (like Outside Temp Sensors) if they already exist in the overall plant.
            
            User description: "{project_description}"
            
            Return ONLY a JSON array with the exact following structure for each component (no markdown, no additional text):
            [
              {{"Main Equipment": "LPHW Secondary Circuit 1", "Component": "Secondary Flow Temperature Sensor", "Quantity": 1, "Notes": "Rule applied"}},
              {{"Main Equipment": "AHU 1", "Component": "Frost Stat", "Quantity": 1, "Notes": "Rule applied"}}
            ]
            """
            
            try:
                response = model.generate_content(prompt)
                json_text = response.text.strip().replace("```json", "").replace("```", "")
                materials_data = json.loads(json_text)
                
                df = pd.DataFrame(materials_data)
                buffer = io.BytesIO()
                
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Bill of Materials')
                    worksheet = writer.sheets['Bill of Materials']
                    
                    for col_num, col in enumerate(df.columns, 1):
                        max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                        col_letter = get_column_letter(col_num)
                        worksheet.column_dimensions[col_letter].width = max_len
                    
                    max_row = len(df) + 1
                    max_col = len(df.columns)
                    max_col_letter = get_column_letter(max_col)
                    table_ref = f"A1:{max_col_letter}{max_row}"
                    
                    tab = Table(displayName="BOM_Table", ref=table_ref)
                    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                    tab.tableStyleInfo = style
                    worksheet.add_table(tab)

                st.success("Excel generated successfully!")
                st.download_button(
                    label="Download Excel",
                    data=buffer.getvalue(),
                    file_name="Circuit_Materials.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.dataframe(df, use_container_width=True)

            except Exception as e:
                st.error(f"An error occurred during generation: {e}\nEnsure the AI returned a valid JSON format.")