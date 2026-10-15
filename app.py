import streamlit as st
import pandas as pd
import io
import json
import copy
from datetime import datetime
import google.generativeai as genai
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
import os

# --- 1. PAGE CONFIGURATION & MODERN SAAS MINIMALIST UI ---
st.set_page_config(
    page_title="DC Controls | BEMS Estimator", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    /* Global background and typography */
    .stApp {
        background-color: #0F172A;
        background: radial-gradient(circle at 50% 0%, #1E293B 0%, #0F172A 100%);
        font-family: -apple-system, BlinkMacSystemFont, "Inter", "Segoe UI", sans-serif;
    }
    
    /* Center container card */
    .block-container {
        max-width: 860px !important;
        padding-top: 3.5rem !important;
        padding-bottom: 4rem !important;
    }
    
    .saas-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 32px 36px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1);
        margin-bottom: 24px;
    }
    
    .brand-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        border-bottom: 1px solid #F1F5F9;
        padding-bottom: 18px;
        margin-bottom: 24px;
    }
    
    .brand-title {
        font-size: 20px;
        font-weight: 600;
        color: #0F172A;
        letter-spacing: -0.3px;
        margin: 0;
    }
    
    .brand-sub {
        font-size: 13px;
        color: #64748B;
        margin: 4px 0 0 0;
    }

    .badge {
        background-color: #F1F5F9;
        color: #002B49;
        font-size: 11px;
        font-weight: 600;
        padding: 4px 10px;
        border-radius: 9999px;
        border: 1px solid #E2E8F0;
        letter-spacing: 0.5px;
    }

    /* Form labels & inputs */
    label p {
        font-size: 13px !important;
        font-weight: 600 !important;
        color: #334155 !important;
        margin-bottom: 4px !important;
    }

    .stTextInput input, .stTextArea textarea {
        background-color: #F8FAFC !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
        color: #0F172A !important;
        font-size: 14px !important;
        padding: 10px 14px !important;
        box-shadow: none !important;
    }

    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #002B49 !important;
        background-color: #FFFFFF !important;
        box-shadow: 0 0 0 1px #002B49 !important;
    }

    /* Primary Action Button */
    .stButton>button {
        background-color: #002B49 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 8px !important;
        font-size: 14px !important;
        font-weight: 500 !important;
        height: 44px !important;
        width: 100% !important;
        transition: all 0.15s ease-in-out !important;
    }

    .stButton>button:hover {
        background-color: #0A3D62 !important;
        box-shadow: 0 4px 12px rgba(0, 43, 73, 0.25) !important;
    }

    /* Download Buttons */
    .stDownloadButton>button {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        height: 42px !important;
        width: 100% !important;
        transition: all 0.15s ease-in-out !important;
    }

    .stDownloadButton>button:hover {
        background-color: #F8FAFC !important;
        border-color: #94A3B8 !important;
    }

    /* Final clean button highlight */
    div[data-testid="stDownloadButton"]:has(button:contains("Clean")) button,
    .clean-btn>button {
        background-color: #0284C7 !important;
        color: #FFFFFF !important;
        border: none !important;
    }
    
    /* Result tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 16px;
        border-bottom: 1px solid #E2E8F0;
    }
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        background-color: transparent !important;
        border: none !important;
        color: #64748B;
        font-size: 13px;
        font-weight: 500;
        padding: 0 8px;
    }
    .stTabs [aria-selected="true"] {
        color: #002B49 !important;
        border-bottom: 2px solid #002B49 !important;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

if "generated" not in st.session_state:
    st.session_state.generated = False
if "excel_reviewed" not in st.session_state:
    st.session_state.excel_reviewed = False

def mark_downloaded():
    st.session_state.excel_reviewed = True

# --- 2. API KEY RESOLUTION ---
api_key = os.environ.get("GEMINI_API_KEY")
if not api_key:
    try:
        api_key = st.secrets.get("GEMINI_API_KEY")
    except Exception:
        api_key = None

if not api_key:
    with st.sidebar:
        st.subheader("Configuration")
        api_key = st.text_input("Gemini API Key:", type="password")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash')

# === SYSTEMS LIST ===
allowed_systems = [
    "Common LPHW Devices", "Hot Water Generator  No.1", "Boiler", "LPHW Pressurisation Unit",
    "Dosing Unit", "Primary LPHW Single Pumps", "Calorifier", "Gas Detection Panel",
    "Storage Tank - Mains Water", "Booster Unit - Mains Water", "Storage Tank - Cold Water",
    "Booster Unit - Cold Water", "Storage Tank - Fire Hose Reel", "Storage Tank - Rain Water",
    "Rain Water Harvesting", "Chiller", "Primary CHW Pump A", "Primary CHW Pump B", 
    "CHW Pressurisation Unit", "Chilled Water Buffer Tank", "AHU", "Extract Fan",
    "Natural Ventilation", "Metering", "Generator", "Unit Heater", "Air Curtain", "AC Units",
    "Radiator Circuit", "FCU (Fan Coil Unit)"
]

component_catalog = {
    "LPHW Header Flow Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "LPHW Header Return Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Outside Frost Thermostat": {"Part": "DBET-23U", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Outside Temperature Sensor": {"Part": "TB/TO", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Immersion Frost Thermostat": {"Part": "DBTV-2U", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Boiler Enable": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Boiler Common Fault": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Boiler Control Signal": {"Part": "0…10V dc", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50},
    "Boiler Flow Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Boiler Return Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Primary LPHW Pump": {"Part": "Intelligent Pumps", "AI": 0, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Enable": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Status": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Variable Speed Drive": {"Part": "VSD", "AI": 0, "AO": 1, "DI": 1, "DO": 1, "Labour": 50},
    "Flow Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "3 Port Control Valve / Actuator": {"Part": "Valve 40mm", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50},
    "LPHW Pressurisation Unit High Pressure": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "LPHW Pressurisation Unit Low Pressure": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Calorifier Enable and Fault": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 1, "Labour": 50},
    "Calorifier Immersion Temperature Sensor": {"Part": "TI/Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "High Limit Thermostat (60-95 Man. Reset)": {"Part": "RAK TW 1000B", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Return Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Space Temperature Sensor": {"Part": "RS-Temp", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Control Valve Actuator": {"Part": "MVC / DB_VZ", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50},
    "Supply Air Temp Sensor": {"Part": "Duct Temp Sensor", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Return Air Temp Sensor": {"Part": "Duct Temp Sensor", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Frost Stat": {"Part": "DBET-23U", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Chiller Enable and Status": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Chiller Flow Switch": {"Part": "FS 541", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Chiller Flow Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Chiller Return Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Enable and Current Switch": {"Part": "RIBXKTF", "AI": 0, "AO": 0, "DI": 1, "DO": 1, "Labour": 50},
    "Radiator Enable": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Radiator Flow Temp Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Radiator Return Temp Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Radiator Control Valve Actuator": {"Part": "Valve 25mm", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50}
}

def get_catalog_match(desc, part):
    d = desc.lower()
    for k, v in component_catalog.items():
        if k.lower() == d: return v
    if "immersion temperature" in d and "flow" in d: return component_catalog["LPHW Header Flow Immersion Temperature Sensor"]
    if "immersion temperature" in d and "return" in d: return component_catalog["LPHW Header Return Immersion Temperature Sensor"]
    if "frost stat" in d or "frost thermostat" in d: return component_catalog["Outside Frost Thermostat"]
    if "space temperature" in d or "room temp" in d: return component_catalog["Space Temperature Sensor"]
    if "control valve" in d or "actuator" in d: return {"Part": part if part else "Valve", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50}
    if "enable" in d: return {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50}
    if "fire alarm interface" in d: return {"Part": part if part else "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50}
    if "pulsed input" in d or "meter" in d: return {"Part": part if part else "Device By Others", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50}
    if "fault" in d or "status" in d or "switch" in d: return {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50}
    if "variable speed" in d or "vsd" in d: return {"Part": "VSD", "AI": 0, "AO": 1, "DI": 1, "DO": 1, "Labour": 50}
    if "control signal" in d: return {"Part": "0...10V dc", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50}
    return {"Part": part, "AI": 0, "AO": 0, "DI": 0, "DO": 0, "Labour": 50}

def match_native_row(sys_name, ws):
    s = sys_name.lower().strip()
    for r in range(7, ws.max_row):
        desc = str(ws.cell(row=r, column=2).value or "").strip().lower()
        part = str(ws.cell(row=r, column=9).value or "").strip()
        if not desc or part: continue
        if s == desc: return r
        
        if "primary chw pump a" in s and "primary chw pump a" in desc: return r
        if "primary chw pump b" in s and "primary chw pump b" in desc: return r
        if "boiler" in s and "boiler" == desc: return r
        if "pressurisation" in s and "pressurisation" in desc: return r
        if "calorifier" in s and "calorifier" == desc: return r
        if "lphw" in s and "common" in s and "lphw" in desc and "common" in desc: return r
        if "primary lphw" in s and "primary lphw" in desc: return r
        if "ahu" in s and ("ahu" == desc or "air handling unit" in desc): return r
        if "fcu" in s and ("fcu" in desc or "fan coil" in desc): return r
        if "chiller" in s and "chiller" == desc: return r
        if "extract fan" in s and "extract fan" in desc: return r
        if "metering" in s and "metering" in desc: return r
        if "radiator" in s and "radiator" == desc: return r
    return None

def clonar_estilo_columna_local(source_ws, target_ws, source_row, target_row, es_padre=False):
    for c in range(1, 15):
        try:
            s_cell = source_ws.cell(row=source_row, column=c)
            t_cell = target_ws.cell(row=target_row, column=c)
            if isinstance(t_cell, MergedCell) or isinstance(s_cell, MergedCell): continue
            fuente = copy.copy(source_ws.cell(row=source_row, column=9).font) if c == 8 else copy.copy(s_cell.font)
            if c == 2: fuente.bold = es_padre
            t_cell.font = fuente
            t_cell.border = copy.copy(s_cell.border)
            t_cell.alignment = copy.copy(s_cell.alignment)
            t_cell.fill = copy.copy(s_cell.fill)
            t_cell.number_format = s_cell.number_format
        except AttributeError: pass

def procesar_archivos(json_data, nombre_proyecto):
    try:
        if os.path.exists("template_2_2.xlsx"):
            wb_quot = load_workbook("template_2_2.xlsx")
            template_path = "template_2_2.xlsx"
        else:
            wb_quot = load_workbook("template_2.xlsx")
            template_path = "template_2.xlsx"
            
        ws_quot = wb_quot.active
        
        ws_quot['B1'] = f"Project: {nombre_proyecto}"
        ws_quot['B2'] = f"Quote Ref: QN{datetime.now().strftime('%y%m%d')}"
        ws_quot['B3'] = "Revision: 0"
        ws_quot['B4'] = f"Date: {datetime.now().strftime('%d/%m/%Y')}"
        ws_quot['B5'] = "Prepared By: Generated by AI Estimator"
        
        req_row = ws_quot.max_row
        for r in range(7, ws_quot.max_row):
            desc = str(ws_quot.cell(row=r, column=2).value or "").strip()
            if desc in ["BMS Requirements", "Summary"]:
                req_row = r
                break
            cell = ws_quot.cell(row=r, column=8)
            if not isinstance(cell, MergedCell): 
                cell.value = None

        has_chiller = False
        has_boiler = False
        for k, v in json_data.items():
            if v > 0:
                if "chiller" in k.lower(): has_chiller = True
                if "boiler" in k.lower(): has_boiler = True
                
        if has_chiller:
            json_data["Primary CHW Pump A"] = 1
            json_data["Primary CHW Pump B"] = 1
            
        if has_boiler:
            if "Common LPHW Devices" not in json_data or json_data["Common LPHW Devices"] == 0:
                json_data["Common LPHW Devices"] = 1

        for sys_name, qty in json_data.items():
            if qty <= 0: continue
            match_r = match_native_row(sys_name, ws_quot)
            if match_r:
                ws_quot.cell(row=match_r, column=8).value = qty

        buf_quot_full = io.BytesIO()
        wb_quot.save(buf_quot_full)

        # --- 2. PROCESS BOM ---
        wb_bom = load_workbook(template_path)
        ws_bom = wb_bom.active
        
        ws_bom['B1'] = f"Project: {nombre_proyecto}"
        ws_bom['B2'] = f"Quote Ref: QN{datetime.now().strftime('%y%m%d')}"
        ws_bom['B3'] = "Revision: 0"
        ws_bom['B4'] = f"Date: {datetime.now().strftime('%d/%m/%Y')}"
        ws_bom['B5'] = "Prepared By: Generated by AI Estimator - I/O Schedule"
        
        aggregated_io = {}
        for r in range(7, req_row):
            desc = str(ws_quot.cell(row=r, column=2).value or "").strip()
            part = str(ws_quot.cell(row=r, column=9).value or "").strip()
            qty_val = ws_quot.cell(row=r, column=8).value
            
            q = 0
            if qty_val is not None:
                try: q = int(float(qty_val))
                except: pass
            
            if desc and not part and q > 0:
                parent_qty = q
                for child_r in range(r + 1, req_row):
                    c_desc = str(ws_quot.cell(row=child_r, column=2).value or "").strip()
                    c_part = str(ws_quot.cell(row=child_r, column=9).value or "").strip()
                    c_qty_val = ws_quot.cell(row=child_r, column=8).value
                    
                    if not c_desc: break 
                    if not c_part:
                        c_q = 0
                        try:
                            if c_qty_val is not None: c_q = int(float(c_qty_val))
                        except: pass
                        if c_q > 0: break 
                        
                    if not c_part and not str(ws_quot.cell(row=child_r, column=6).value).startswith("="):
                        if not str(ws_quot.cell(row=child_r, column=3).value).startswith("="):
                            continue 
                        
                    cat_data = get_catalog_match(c_desc, c_part)
                    part_no = c_part if c_part else cat_data["Part"]
                    if not part_no: part_no = "Generic"
                    agg_key = part_no.strip().lower()
                    
                    if agg_key not in aggregated_io:
                        clean_desc = c_desc
                        if "volt free" in part_no.lower() or "ribxktf" in part_no.lower():
                            clean_desc = "Software/Relay Interfaces (Enable/Status/Fault)"
                        elif "by others" in part_no.lower():
                            clean_desc = "3rd Party Device Interface"
                        elif "tti" in part_no.lower() or "temp" in part_no.lower() or "dbet" in part_no.lower() or "dbtv" in part_no.lower():
                            clean_desc = "Temperature Sensor / Thermostat"
                        elif "valve" in part_no.lower() or "actuator" in part_no.lower():
                            clean_desc = "Control Valve / Actuator"
                            
                        aggregated_io[agg_key] = {
                            "Description": clean_desc, "Part No.": part_no, "Quantity": 0,
                            "AI": cat_data["AI"], "AO": cat_data["AO"], "DI": cat_data["DI"], "DO": cat_data["DO"],
                            "Labour_Base": cat_data["Labour"]
                        }
                    aggregated_io[agg_key]["Quantity"] += parent_qty
                    
            elif desc and part and q > 0:
                cat_data = get_catalog_match(desc, part)
                part_no = part
                agg_key = part_no.strip().lower()
                
                if agg_key not in aggregated_io:
                    clean_desc = desc
                    if "volt free" in part_no.lower() or "ribxktf" in part_no.lower():
                        clean_desc = "Software/Relay Interfaces (Enable/Status/Fault)"
                    elif "by others" in part_no.lower():
                        clean_desc = "3rd Party Device Interface"
                    elif "tti" in part_no.lower() or "temp" in part_no.lower() or "dbet" in part_no.lower() or "dbtv" in part_no.lower():
                        clean_desc = "Temperature Sensor / Thermostat"
                    elif "valve" in part_no.lower() or "actuator" in part_no.lower():
                        clean_desc = "Control Valve / Actuator"
                        
                    aggregated_io[agg_key] = {
                        "Description": clean_desc, "Part No.": part_no, "Quantity": 0,
                        "AI": cat_data["AI"], "AO": cat_data["AO"], "DI": cat_data["DI"], "DO": cat_data["DO"],
                        "Labour_Base": cat_data["Labour"]
                    }
                aggregated_io[agg_key]["Quantity"] += q
        
        for r in range(7, ws_bom.max_row):
            for c in range(2, 13):
                cell = ws_bom.cell(row=r, column=c)
                if not isinstance(cell, MergedCell): cell.value = None

        ws_bom.cell(row=7, column=2).value = "Consolidated Bill of Materials (BOM)"
        clonar_estilo_columna_local(wb_quot.active, ws_bom, 7, 7, es_padre=True)
        
        current_row = 8
        sorted_keys = sorted(aggregated_io.keys(), key=lambda k: aggregated_io[k]["Description"])
        
        for k in sorted_keys:
            data = aggregated_io[k]
            tot_qty = data["Quantity"]
            calc_labour = data["Labour_Base"] * tot_qty
            
            ws_bom.cell(row=current_row, column=2).value = data["Description"]
            ws_bom.cell(row=current_row, column=3).value = data["AI"] if data["AI"] > 0 else ""
            ws_bom.cell(row=current_row, column=4).value = data["AO"] if data["AO"] > 0 else ""
            ws_bom.cell(row=current_row, column=5).value = data["DI"] if data["DI"] > 0 else ""
            ws_bom.cell(row=current_row, column=6).value = data["DO"] if data["DO"] > 0 else ""
            ws_bom.cell(row=current_row, column=8).value = tot_qty
            ws_bom.cell(row=current_row, column=9).value = data["Part No."]
            ws_bom.cell(row=current_row, column=12).value = calc_labour
            
            clonar_estilo_columna_local(wb_quot.active, ws_bom, 8, current_row, es_padre=False)
            current_row += 1

        delete_start = current_row
        delete_amount = ws_bom.max_row - delete_start + 1
        if delete_amount > 0:
            for mcr in list(ws_bom.merged_cells.ranges):
                if mcr.min_row >= delete_start: ws_bom.merged_cells.remove(mcr)
            ws_bom.delete_rows(idx=delete_start, amount=delete_amount)

        buf_bom = io.BytesIO()
        wb_bom.save(buf_bom)

        # --- 3. FINAL QUOTATION CLEANING ---
        wb_clean = load_workbook(io.BytesIO(buf_quot_full.getvalue()))
        ws_clean = wb_clean.active

        main_system_rows = set()
        for sys in allowed_systems:
            r_match = match_native_row(sys, ws_clean)
            if r_match:
                main_system_rows.add(r_match)

        current_block_keep = True
        
        for r in range(7, req_row):
            c_b = ws_clean.cell(row=r, column=2)
            desc = str(c_b.value or "").strip()
            is_bold = c_b.font and c_b.font.bold
            
            prev_desc = str(ws_clean.cell(row=r-1, column=2).value or "").strip() if r > 7 else ""
            
            is_main_system = False
            if r in main_system_rows:
                is_main_system = True
            elif desc and is_bold and not prev_desc:
                is_main_system = True
                
            if "primary chw pump" in desc.lower():
                is_main_system = False
                
            if is_main_system:
                qty_val = ws_clean.cell(row=r, column=8).value
                current_block_keep = False 
                try:
                    if qty_val is not None and int(float(qty_val)) > 0:
                        current_block_keep = True
                except:
                    pass
            
            if "Fire Alarm Interface" in desc or "Metering" in desc or "Pulsed Input" in desc:
                current_block_keep = True
            
            if not current_block_keep:
                ws_clean.row_dimensions[r].hidden = True

        buf_quot_clean = io.BytesIO()
        wb_clean.save(buf_quot_clean)
        
        return buf_quot_full.getvalue(), buf_bom.getvalue(), buf_quot_clean.getvalue()
        
    except Exception as e:
        st.error(f"Error processing files: {e}")
        return None, None, None

# --- 4. MAIN USER INTERFACE CARD ---
st.markdown("""
<div class="saas-card">
    <div class="brand-header">
        <div>
            <h2 class="brand-title">DC Controls Estimator</h2>
            <p class="brand-sub">Building Energy Management Systems (BEMS)</p>
        </div>
        <div>
            <span class="badge">v2.4 PROD</span>
        </div>
    </div>
""", unsafe_allow_html=True)

# Form inside card
project_name = st.text_input("Project Reference", placeholder="e.g. DC-2026-Cavan")
project_description = st.text_area(
    "Engineering Scope & Plant Specifications", 
    placeholder="e.g. Plant room with 2 Boilers, 2 Chillers, 25 Radiator Circuits, and 4 FCUs.",
    height=120
)

st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

if st.button("Generate Quotation Package"):
    if not project_description:
        st.warning("Please provide the project engineering scope.")
    elif not api_key:
        st.error("Missing Gemini API Key. Configure GEMINI_API_KEY in Streamlit Secrets.")
    else:
        with st.status("Analyzing specifications and generating models...", expanded=False) as status:
            prompt = f'''
            You are a BEMS Estimator. Read the user description and count how many of the ALLOWED SYSTEMS are needed.
            
            ALLOWED SYSTEMS EXACT NAMES:
            {json.dumps(allowed_systems, indent=2)}
            
            RULES:
            1. ONLY output a JSON dictionary mapping the EXACT system name from the list to its integer quantity.
            2. DO NOT output children or points. JUST the system name and quantity.
            3. If the user mentions Boilers or LPHW Heating, add "Common LPHW Devices": 1 to the JSON.
            4. If the user mentions Chillers, ALWAYS add "Primary CHW Pump A": 1 AND "Primary CHW Pump B": 1 to the JSON.
            5. ONLY use names exactly as they appear in the ALLOWED SYSTEMS list. DO NOT invent names or categories.
            
            User Description: "{project_description}"
            
            RETURN ONLY VALID JSON:
            '''
            try:
                response = model.generate_content(prompt)
                json_text = response.text.strip().replace("```json", "").replace("```", "")
                system_quantities = json.loads(json_text)
                
                st.session_state.materials_data = system_quantities 
                
                buffer_full, buffer_bom, buffer_clean = procesar_archivos(system_quantities, project_name)
                
                if buffer_full and buffer_bom and buffer_clean:
                    st.session_state.generated = True
                    st.session_state.buffer_full = buffer_full
                    st.session_state.buffer_bom = buffer_bom
                    st.session_state.buffer_clean = buffer_clean
                    st.session_state.nombre_full = f"Quotation_Review_{project_name.replace(' ', '_')}.xlsx"
                    st.session_state.nombre_clean = f"Clean_Quotation_{project_name.replace(' ', '_')}.xlsx"
                    st.session_state.nombre_io = f"IO_Schedule_{project_name.replace(' ', '_')}.xlsx"
                    
                    st.session_state.excel_reviewed = False
            except Exception as e:
                st.error(f"Execution Error: {e}")

st.markdown("</div>", unsafe_allow_html=True)

# --- 5. RESULTS CARD ---
if st.session_state.generated:
    st.markdown("""
    <div class="saas-card">
        <h3 style="font-size: 16px; font-weight: 600; color: #0F172A; margin: 0 0 16px 0;">Deliverables</h3>
    """, unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["Review Package", "Client Deliverable"])
    
    with tab1:
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2, gap="medium")
        with col1:
            st.download_button(
                "Download Master Quotation (Full)",
                data=st.session_state.buffer_full,
                file_name=st.session_state.nombre_full,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                on_click=mark_downloaded
            )
        with col2:
            st.download_button(
                "Download I/O Schedule (BOM)",
                data=st.session_state.buffer_bom,
                file_name=st.session_state.nombre_io,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    with tab2:
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
        if st.session_state.get("excel_reviewed", False):
            st.download_button(
                label="Download Clean Quotation (Client Ready)",
                data=st.session_state.buffer_clean,
                file_name=st.session_state.nombre_clean,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Download the Master Quotation in Step 1 first to unlock the client-ready package.")
            
    st.markdown("</div>", unsafe_allow_html=True)
