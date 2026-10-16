import streamlit as st
import pandas as pd
import io
import json
import copy
from datetime import datetime
import google.generativeai as genai
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Border, Side
import openpyxl.worksheet.properties
import subprocess
import os

# --- 1. PAGE CONFIGURATION & STYLING ---
st.set_page_config(
    page_title="BEMS Estimator Pro | DC Controls", 
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main {
        background-color: #F8F9FA;
    }
    h1, h2, h3 {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        color: #002B49;
    }
    .stContainer {
        background-color: #FFFFFF;
        padding: 24px;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        border: 1px solid #E2E8F0;
    }
    .stButton>button {
        border-radius: 6px;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    [data-testid="stSidebar"] {
        background-color: #001F33;
        color: #FFFFFF;
    }
    [data-testid="stSidebar"] label {
        color: #E2E8F0 !important;
    }
    [data-testid="stSidebar"] .stTextInput input {
        background-color: #002B49;
        color: #FFFFFF;
        border: 1px solid #1E3A5F;
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

# --- 3. SIDEBAR ---
with st.sidebar:
    st.markdown("<br>", unsafe_allow_html=True)
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    elif os.path.exists("logo.jpg"):
        st.image("logo.jpg", use_container_width=True)
    else:
        st.markdown("<h2 style='text-align: center; color: #FFFFFF;'>DC CONTROLS</h2>", unsafe_allow_html=True)
        
    st.markdown("<p style='text-align: center; color: #94A3B8; font-size: 13px;'>BEMS Estimation Engine</p>", unsafe_allow_html=True)
    st.markdown("<hr style='border-color: #1E3A5F;'>", unsafe_allow_html=True)
    
    st.markdown("### System Access")
    if not api_key:
        api_key = st.text_input("Enter API Key (Gemini):", type="password", help="Requires authorized Gemini API token.")
    else:
        st.markdown("<p style='color: #38BDF8; font-size: 12px;'>API Key configured securely via environment.</p>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<p style='color: #94A3B8; font-size: 12px;'>Automatic point schedule generation, BOM consolidation, and commercial quotation export.</p>", unsafe_allow_html=True)

# --- 4. MAIN HEADER ---
st.markdown("<h1 style='text-align: center; color: #003366;'>Automated BEMS Quotation System</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 16px; color: #555;'>AI-Powered I/O Schedule and Cost Estimator</p>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-3.6-flash') 

# === SYSTEMS LIST ===
allowed_systems = [
    "Common LPHW Devices", "Hot Water Generator  No.1", "Boiler", "LPHW Pressurisation Unit",
    "Dosing Unit", "Primary LPHW Single Pumps", "Calorifier", "Gas Detection Panel",
    "Storage Tank - Mains Water", "Booster Unit - Mains Water", "Storage Tank - Cold Water",
    "Booster Unit - Cold Water", "Storage Tank - Fire Hose Reel", "Storage Tank - Rain Water",
    "Rain Water Harvesting", "Chiller", "Primary CHW Pump A", "Primary CHW Pump B", 
    "CHW Pressurisation Unit", "Chilled Water Buffer Tank", "AHU", "Extract Fan",
    "Natural Ventilation", "Metering", "Fire Alarm Interface", "Generator", "Unit Heater", "Air Curtain", "AC Units",
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
        
        if not desc: continue
        # Permitir que Fire Alarm coincida aunque tenga Part No. ("Volt Free Contacts")
        if part and "fire alarm" not in s: continue
        
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
        if "fire alarm" in s and "fire alarm" in desc: return r
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

def aplicar_bordes_profesionales(ws):
    borde_azul = Side(style='medium', color='003366')
    borde_tabla = Side(style='thin', color='B0C4DE')
    
    # 1. BORDES EXTERIORES DEL ENCABEZADO (B1:I5)
    for r in range(1, 6):
        for c in range(2, 10):
            cell = ws.cell(row=r, column=c)
            top = borde_azul if r == 1 else None
            bottom = borde_azul if r == 5 else None
            left = borde_azul if c == 2 else None
            right = borde_azul if c == 9 else None
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)
            
    # 2. BORDES COMPLETOS EN CADA FILA VISIBLE DE LA TABLA HASTA EL FINAL
    for r in range(6, ws.max_row + 1):
        if ws.row_dimensions[r].hidden:
            continue
            
        for c in range(2, 10):
            cell = ws.cell(row=r, column=c)
            if cell.border and cell.border.left and cell.border.left.style == 'medium':
                continue
            
            left_side = borde_azul if (c == 2 and r == 6) else (borde_tabla if c > 2 else Side(style='medium', color='003366'))
            right_side = borde_azul if (c == 9 and r == 6) else (borde_tabla if c < 9 else Side(style='medium', color='003366'))
            
            cell.border = Border(
                top=borde_tabla, 
                bottom=borde_tabla, 
                left=left_side, 
                right=right_side
            )

def procesar_archivos(json_data, nombre_proyecto, customer, contact, quote_ref_user):
    try:
        template_file = "template_2_2.xlsx" if os.path.exists("template_2_2.xlsx") else "template_2.xlsx"
        wb_quot = load_workbook(template_file)
        ws_quot = wb_quot.active
        
        # Inyección de encabezado
        ref_text = quote_ref_user.strip() if quote_ref_user and quote_ref_user.strip() else f"QN{datetime.now().strftime('%y%m%d')}"
        
        ws_quot['B1'] = f"Project: {nombre_proyecto}"
        ws_quot['B2'] = f"Quote Ref: {ref_text}"
        ws_quot['B3'] = "Revision: 0"
        ws_quot['B4'] = f"Date: {datetime.now().strftime('%d/%m/%Y')}"
        ws_quot['B5'] = "Prepared by: DC Controls"
        
        if customer and customer.strip():
            ws_quot['E2'] = f"Customer: {customer.strip()}"
        if contact and contact.strip():
            ws_quot['E3'] = f"Contact: {contact.strip()}"
        
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
                # Inyección en DI (columna E) para activar las fórmulas de costo de Fire Alarm
                if "fire alarm" in sys_name.lower():
                    ws_quot.cell(row=match_r, column=5).value = qty

        for r in range(7, req_row):
            cell = ws_quot.cell(row=r, column=2)
            if not isinstance(cell, MergedCell):
                val = str(cell.value or "").strip()
                if "XXXX" in val or "??" in val:
                    cell.value = val.split(" - ")[0].strip()

        buf_quot_full = io.BytesIO()
        wb_quot.save(buf_quot_full)

        # --- 2. PROCESS BOM ---
        wb_bom = load_workbook(template_file)
        ws_bom = wb_bom.active
        
        ws_bom['B1'] = f"Project: {nombre_proyecto}"
        ws_bom['B2'] = f"Quote Ref: {ref_text}"
        ws_bom['B3'] = "Revision: 0"
        ws_bom['B4'] = f"Date: {datetime.now().strftime('%d/%m/%Y')}"
        ws_bom['B5'] = "Prepared by: DC Controls"
        if customer and customer.strip():
            ws_bom['E2'] = f"Customer: {customer.strip()}"
        if contact and contact.strip():
            ws_bom['E3'] = f"Contact: {contact.strip()}"
        
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
        has_fire_alarm = any("fire alarm" in k.lower() for k, v in json_data.items() if v > 0)
        
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

            # Manejo de visibilidad específico para Fire Alarm
            if "fire alarm" in desc.lower():
                if has_fire_alarm:
                    ws_clean.row_dimensions[r].hidden = False
                    continue
                else:
                    ws_clean.row_dimensions[r].hidden = True
                    continue
            
            if not current_block_keep:
                ws_clean.row_dimensions[r].hidden = True

        # Aplicar balance de bordes
        aplicar_bordes_profesionales(ws_clean)

        # Configuración de página limpia para impresión
        ws_clean.print_options.gridLines = False
        if not ws_clean.sheet_properties.pageSetUpPr:
            ws_clean.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetUpPr()
        ws_clean.sheet_properties.pageSetUpPr.fitToPage = True
        ws_clean.page_setup.fitToWidth = 1
        ws_clean.page_setup.fitToHeight = 0

        buf_quot_clean = io.BytesIO()
        wb_clean.save(buf_quot_clean)
        
        return buf_quot_full.getvalue(), buf_bom.getvalue(), buf_quot_clean.getvalue()
        
    except Exception as e:
        st.error(f"Error processing files: {e}")
        return None, None, None

def convert_excel_to_pdf(excel_bytes):
    temp_excel = os.path.abspath("temp_clean_quotation.xlsx")
    temp_dir = os.path.abspath(".")
    
    with open(temp_excel, "wb") as f:
        f.write(excel_bytes)
    
    pdf_bytes = None
    
    if os.name != "nt":
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", temp_dir,
            temp_excel
        ]
        subprocess.run(cmd, capture_output=True, text=True)
        expected_pdf = os.path.splitext(temp_excel)[0] + ".pdf"
        
        if os.path.exists(expected_pdf):
            with open(expected_pdf, "rb") as f:
                pdf_bytes = f.read()
            os.remove(expected_pdf)
    else:
        temp_pdf = os.path.abspath("temp_clean_quotation.pdf")
        ps_script = f"""
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open('{temp_excel}')
        $wb.ExportAsFixedFormat(0, '{temp_pdf}')
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        """
        subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], capture_output=True)
        if os.path.exists(temp_pdf):
            with open(temp_pdf, "rb") as f:
                pdf_bytes = f.read()
            os.remove(temp_pdf)
            
    if os.path.exists(temp_excel):
        os.remove(temp_excel)
        
    return pdf_bytes

# --- 5. PROJECT DETAILS INPUT CONTAINER ---
st.markdown("### Project Details")
with st.container():
    col1, col2 = st.columns([1, 2])
    with col1:
        project_name = st.text_input("Project Name / Reference:", placeholder="e.g. IDA Cavan - Building 2")
        customer_name = st.text_input("Customer:", placeholder="e.g. John Doe / Company")
        contact_name = st.text_input("Contact:", placeholder="e.g. contact@domain.com")
    with col2:
        quote_ref = st.text_input("Quote Ref:", placeholder="e.g. QN-2026-001 (Leave empty for auto)")
        project_description = st.text_area(
            "Engineering Scope:", 
            placeholder="Example: We need a plant with 2 Boilers, 2 Chillers, 25 Radiator Circuits, and 4 FCUs. Include Metering and Fire Alarm Interface.",
            height=100
        )

st.write("") 

col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    btn_generate = st.button("Analyze Scope & Generate Quotation", type="primary", use_container_width=True)

if btn_generate:
    if not api_key:
        st.warning("Please provide an API Key in the sidebar or configure it in Secrets.")
    elif not project_description:
        st.warning("Please provide a project engineering description.")
    else:
        with st.status("AI Estimator is working...", expanded=True) as status:
            prompt = f'''
            You are a BEMS Estimator. Read the user description and count how many of the ALLOWED SYSTEMS are needed.
            
            ALLOWED SYSTEMS EXACT NAMES:
            {json.dumps(allowed_systems, indent=2)}
            
            RULES:
            1. ONLY output a JSON dictionary mapping the EXACT system name from the list to its integer quantity.
            2. DO NOT output children or points. JUST the system name and quantity.
            3. If the user mentions Boilers or LPHW Heating, add "Common LPHW Devices": 1 to the JSON.
            4. If the user mentions Chillers, ALWAYS add "Primary CHW Pump A": 1 AND "Primary CHW Pump B": 1 to the JSON.
            5. ONLY include "Metering" if the user explicitly specifies meters, energy metering, or pulsed inputs.
            6. ONLY include "Fire Alarm Interface" if the user explicitly mentions fire alarm shutdown, interlock, or interface.
            7. ONLY use names exactly as they appear in the ALLOWED SYSTEMS list. DO NOT invent names or categories.
            
            User Description: "{project_description}"
            
            RETURN ONLY VALID JSON:
            '''
            try:
                response = model.generate_content(prompt)
                json_text = response.text.strip().replace("```json", "").replace("```", "")
                system_quantities = json.loads(json_text)
                
                st.session_state.materials_data = system_quantities 
                
                buffer_full, buffer_bom, buffer_clean = procesar_archivos(
                    system_quantities, 
                    project_name, 
                    customer_name, 
                    contact_name, 
                    quote_ref
                )
                
                if buffer_full and buffer_bom and buffer_clean:
                    st.session_state.generated = True
                    st.session_state.buffer_full = buffer_full
                    st.session_state.buffer_bom = buffer_bom
                    st.session_state.buffer_clean = buffer_clean
                    st.session_state.nombre_full = f"Quotation_Review_{project_name.replace(' ', '_')}.xlsx"
                    st.session_state.nombre_clean = f"Clean_Quotation_{project_name.replace(' ', '_')}.xlsx"
                    st.session_state.nombre_io = f"IO_Schedule_{project_name.replace(' ', '_')}.xlsx"
                    
                    st.session_state.excel_reviewed = False
                    status.update(label="Process Complete", state="complete", expanded=False)

            except Exception as e:
                status.update(label="Process Failed", state="error", expanded=True)
                st.error(f"Error AI/JSON: {e}")

# --- 6. RESULTS & DOWNLOADS ---
if st.session_state.generated:
    st.markdown("---")
    st.markdown("### Step 1: Review Working Files")
    st.info("Download the full Quotation to review counting logic and the consolidated I/O Schedule.")
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button(
            "Download Quotation (Full / Review)", 
            data=st.session_state.buffer_full, 
            file_name=st.session_state.nombre_full, 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=mark_downloaded,
            use_container_width=True
        )
    with col_d2:
        st.download_button(
            "Download I/O Schedule (Excel)", 
            data=st.session_state.buffer_bom, 
            file_name=st.session_state.nombre_io, 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    if st.session_state.get("excel_reviewed", False):
        st.markdown("---")
        st.markdown("### Step 2: Client-Ready Documents")
        st.success("The clean quotation is ready. All unused systems have been automatically hidden while preserving formulas.")
        
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.download_button(
                label="DOWNLOAD CLEAN QUOTATION (EXCEL)",
                data=st.session_state.buffer_clean,
                file_name=st.session_state.nombre_clean,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_c2:
            with st.spinner("Rendering High-Fidelity Clean PDF..."):
                pdf_data = convert_excel_to_pdf(st.session_state.buffer_clean)
                
                if pdf_data:
                    st.download_button(
                        label="DOWNLOAD OFFICIAL CLIENT PDF",
                        data=pdf_data,
                        file_name=f"Approved_Quotation_{project_name.replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                else:
                    st.info("Direct PDF generation is ready once LibreOffice is deployed on the server via packages.txt.")
