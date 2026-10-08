import streamlit as st
import pandas as pd
import io
import json
import copy
import re
from difflib import SequenceMatcher
from datetime import datetime
import google.generativeai as genai
from openpyxl import load_workbook
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.cell.cell import MergedCell

st.set_page_config(page_title="BEMS Estimator Pro - DC Controls", layout="wide")
st.title("Automated BEMS Points List & Estimator")
st.markdown("---")

if "generado" not in st.session_state:
    st.session_state.generado = False

with st.sidebar:
    st.header("System Configuration")
    api_key = st.text_input("Enter API Key (Gemini):", type="password")
    st.markdown("---")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-3.6-flash') 

engineering_rules = {
    "Common Heating Devices (LPHW)": {"per_unit": ["Header Flow Immersion Temperature Sensor", "Header Return Immersion Temperature Sensor", "Outside Frost Thermostat", "Outside Temperature Sensor", "Immersion Frost Thermostat"]},
    "Common Cooling Devices (CHW)": {"per_unit": ["Header Flow Immersion Temperature Sensor", "Header Return Immersion Temperature Sensor", "Outside Frost Thermostat", "Outside Temperature Sensor"]},
    "Boiler": {"per_unit": ["Boiler Enable", "Boiler Common Fault", "Boiler Control Signal", "Boiler Flow Immersion Temperature Sensor", "Boiler Return Immersion Temperature Sensor"]},
    "CHW Pump (Primary/Secondary)": {"per_unit": ["Pump Enable", "Pump Status", "Variable Speed Drive", "Flow Immersion Temperature Sensor", "3 Port Control Valve / Actuator"]},
    "LPHW Pump (Primary/Secondary)": {"per_unit": ["Pump Enable", "Pump Status", "Variable Speed Drive", "Flow Immersion Temperature Sensor", "3 Port Control Valve / Actuator"]},
    "Pressurisation Unit": {"per_unit": ["Pressurisation Unit High Pressure", "Pressurisation Unit Low Pressure"]},
    "Calorifier / Hot Water Generator": {"per_unit": ["Enable", "Common Fault", "Control Signal", "Immersion Temperature Sensor", "High Limit Thermostat (60-95 Man. Reset)"]},
    "AHU (Air Handling Unit)": {"per_unit": ["Enable", "Status", "Control Signal", "Supply Air Temp Sensor", "Return Air Temp Sensor", "Frost Stat"]},
    "FCU (Fan Coil Unit)": {"per_unit": ["Space Temperature Sensor", "Control Valve Actuator"]},
    "Storage Tank (Cold Water / Mains)": {"per_unit": ["Tank Low Low Level Status", "Tank Section 1 Low Level", "Tank Section 1 High Level", "Tank Section 2 Low Level", "Tank Section 2 High Level", "Tank Immersion Temperature Sensor", "Solenoid Valve 40mm", "Ultrasonic Level Transmitter"]},
    "Chiller": {"per_unit": ["Chiller Enable and Status", "Chiller Flow Switch", "Chiller Flow Immersion Temperature Sensor", "Chiller Return Immersion Temperature Sensor"]},
    "Extract Fan": {"per_unit": ["Enable and Current Switch"]},
    "Metering": {"per_unit": ["Gas Meter Pulsed Input", "Water Meter Pulsed Input", "Electricity Meter Pulsed Input"]}
}

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
    "Chiller Enable and Status": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Chiller Flow Switch": {"Part": "FS 541", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Chiller Flow Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Chiller Return Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Enable and Current Switch": {"Part": "RIBXKTF", "AI": 0, "AO": 0, "DI": 1, "DO": 1, "Labour": 50},
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

def fuzzy_match_parent(ai_desc, tpl_desc):
    ai = str(ai_desc).lower()
    tpl = str(tpl_desc).lower().strip()
    
    if any(x in tpl for x in ["sensor", "actuator", "valve", "switch", "fault", "enable", "status", "mains", "burner", "wheel", "battery", "heater", "kw"]):
        if "pump" in tpl and "pump" not in ai: return False
        if "pump" in tpl and "pump" in ai: pass
        else: return False

    if "chw" in ai and "lphw" in tpl: return False
    if "lphw" in ai and "chw" in tpl: return False

    if "lphw" in ai and "lphw" in tpl and "common" in tpl: return True
    if "chw" in ai and "chw" in tpl and "common" in tpl: return True
    
    if "boiler" in ai and tpl == "boiler": return True
    if "pump" in ai and "pump" in tpl:
        if "recovery" not in tpl: return True
    if "pressurisation" in ai and "pressurisation unit" in tpl: return True
    if "calorifier" in ai and tpl == "calorifier": return True
    if "ahu" in ai and (tpl == "ahu" or tpl == "air handling units" or tpl == "air handling unit"): return True
    if "fcu" in ai and (tpl == "fcu" or "fan coil" in tpl): return True
    if "tank" in ai and "tank" in tpl: return True
    if "chiller" in ai and (tpl == "chiller" or tpl == "chillers"): return True
    if "extract fan" in ai and "extract fan" in tpl: return True
    if "metering" in ai and "metering" in tpl: return True
    return False

def fuzzy_match_strict(ai_str, tpl_str):
    ai_orig = str(ai_str).strip().lower().replace("temp ", "temperature ").replace("return air", "extract air")
    tpl_orig = str(tpl_str).strip().lower().replace("temp ", "temperature ").replace("return air", "extract air")
    
    ai_words = set(re.findall(r'[a-z0-9]+', ai_orig))
    tpl_words = set(re.findall(r'[a-z0-9]+', tpl_orig))
    if not ai_words or not tpl_words: return False

    if len(ai_words) <= 2:
        if ai_orig == tpl_orig: return True
        if ai_words == tpl_words: return True
        if len(tpl_words) <= 3 and ai_words.issubset(tpl_words): return True
        return False
        
    if ai_words == tpl_words: return True
    if ai_words.issubset(tpl_words): return True
    
    critical_keywords = ["immersion", "air", "water", "duct", "room", "space", "supply", "return", "extract", "flow", "pressure", "valve", "actuator"]
    for word in critical_keywords:
        if word in ai_words and word not in tpl_words: return False
        if word in tpl_words and word not in ai_words: return False
            
    if SequenceMatcher(None, ai_orig, tpl_orig).ratio() > 0.85: return True
    return False

def get_catalog_data(desc):
    for k, v in component_catalog.items():
        if fuzzy_match_strict(desc, k) or desc.strip().lower() == k.lower():
            return v
    return {}

def is_parent(item):
    desc = str(item.get("Description", "")).strip().lower()
    part = str(item.get("Part No.", "")).strip()
    if part: return False
    
    parents = ["common", "boiler", "pump", "pressurisation", "calorifier", "hot water generator", "ahu", "air handling", "fcu", "fan coil", "tank", "chiller", "extract fan", "metering"]
    for p in parents:
        if p in desc:
            if any(x in desc for x in ["sensor", "actuator", "valve", "switch", "fault", "enable", "status", "heater", "kw"]):
                return False
            return True
    return False

def actualizar_formulas(ws, green_idx, total_insert):
    def replacer(match):
        col = match.group(1)
        row_num = int(match.group(2))
        if row_num >= green_idx:
            return f"{col}{row_num + total_insert}"
        if row_num == green_idx - 1:
            return f"{col}{row_num + total_insert}"
        return match.group(0)

    for r in range(1, ws.max_row + 1):
        for c in range(1, 15):
            cell = ws.cell(row=r, column=c)
            if not isinstance(cell, MergedCell) and isinstance(cell.value, str) and cell.value.startswith("="):
                new_formula = re.sub(r'([A-Z]{1,2})([0-9]+)', replacer, cell.value)
                if cell.value != new_formula:
                    cell.value = new_formula

def reparar_celdas_combinadas(ws, green_idx, total_insert):
    new_merged = set()
    old_merged_list = list(ws.merged_cells.ranges)
    for mcr in old_merged_list:
        if mcr.min_row >= green_idx:
            new_mcr = CellRange(min_col=mcr.min_col, min_row=mcr.min_row + total_insert,
                                max_col=mcr.max_col, max_row=mcr.max_row + total_insert)
            new_merged.add(new_mcr)
            ws.merged_cells.remove(mcr)
        else:
            new_merged.add(mcr)
            ws.merged_cells.remove(mcr)
            
    for mcr in new_merged:
        ws.merged_cells.add(mcr)

def clonar_estilo_columna(ws, source_row, target_row, es_padre=False):
    for c in range(1, 15):
        try:
            source_cell = ws.cell(row=source_row, column=c)
            target_cell = ws.cell(row=target_row, column=c)
            
            if isinstance(target_cell, MergedCell) or isinstance(source_cell, MergedCell):
                continue
                
            if c == 8:
                fuente = copy.copy(ws.cell(row=source_row, column=9).font)
            else:
                fuente = copy.copy(source_cell.font)
                
            if c == 2: fuente.bold = es_padre
            
            target_cell.font = fuente
            target_cell.border = copy.copy(source_cell.border)
            target_cell.alignment = copy.copy(source_cell.alignment)
            target_cell.fill = copy.copy(source_cell.fill)
            target_cell.number_format = source_cell.number_format
        except AttributeError:
            pass

def limpiar_nombres_feos(ws, start_row, end_row):
    for r in range(start_row, end_row):
        cell = ws.cell(row=r, column=2)
        if not isinstance(cell, MergedCell):
            val = str(cell.value or "").strip()
            if "XXXX" in val or "??" in val:
                cell.value = val.split(" - ")[0].strip()

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
        req_row = ws.max_row
        for r in range(start_row, ws.max_row + 1):
            val = str(ws.cell(row=r, column=2).value).strip()
            if val in ["BMS Requirements", "Summary"]:
                req_row = r - 2 if (r - 2) > start_row else r
                break
                
        green_row_idx = req_row
        for r in range(start_row, req_row):
            fill = ws.cell(row=r, column=2).fill
            if fill and fill.fgColor and fill.fgColor.type == 'indexed' and fill.fgColor.indexed == 42:
                green_row_idx = r
                break
                
        if green_row_idx == ws.max_row:
            green_row_idx = req_row
                
        for r in range(start_row, req_row):
            for c in [3, 4, 5, 6, 7, 8]:
                cell = ws.cell(row=r, column=c)
                if not isinstance(cell, MergedCell):
                    cell.value = None 

        ai_groups = []
        current_ai_parent = None
        for item in datos:
            if is_parent(item):
                current_ai_parent = {"parent": item, "children": [], "best_match_row": None}
                ai_groups.append(current_ai_parent)
            else:
                if current_ai_parent:
                    current_ai_parent["children"].append(item)

        for group in ai_groups:
            ai_p_desc = group["parent"].get("Description", "")
            for r in range(start_row, req_row):
                desc_cell = ws.cell(row=r, column=2).value
                part_cell = ws.cell(row=r, column=9).value
                
                if desc_cell and not part_cell:
                    if fuzzy_match_parent(ai_p_desc, desc_cell):
                        if not any(g.get("best_match_row") == r for g in ai_groups):
                            group["best_match_row"] = r
                            break

        matched_groups = sorted([g for g in ai_groups if g["best_match_row"]], key=lambda x: x["best_match_row"])
        orphans_to_insert = []
        
        for i, group in enumerate(matched_groups):
            best_match_row = group["best_match_row"]
            next_boundary = matched_groups[i+1]["best_match_row"] if i + 1 < len(matched_groups) else green_row_idx
            
            p_item = group["parent"]
            ai_p_desc = p_item.get("Description", "")
            
            tpl_name = str(ws.cell(row=best_match_row, column=2).value or "")
            if "XXXX" in tpl_name or "??" in tpl_name:
                cell_name = ws.cell(row=best_match_row, column=2)
                if not isinstance(cell_name, MergedCell):
                    cell_name.value = ai_p_desc
            
            cell_qty = ws.cell(row=best_match_row, column=8)
            if not isinstance(cell_qty, MergedCell): cell_qty.value = p_item.get("Quantity", "")
            
            if "chiller" in tpl_name.lower():
                next_desc = str(ws.cell(row=best_match_row + 1, column=2).value or "").strip()
                if "chiller mains" in next_desc.lower():
                    cm_qty = ws.cell(row=best_match_row + 1, column=8)
                    if not isinstance(cm_qty, MergedCell):
                        cm_qty.value = p_item.get("Quantity", "")
            
            missing_children = []
            for child in group["children"]:
                ai_c_desc = child.get("Description", "")
                child_matched = False
                
                for cr in range(best_match_row + 1, next_boundary):
                    c_desc = str(ws.cell(row=cr, column=2).value or "").strip()
                    if c_desc:
                        if fuzzy_match_strict(ai_c_desc, c_desc):
                            if not ws.cell(row=cr, column=3).value and not ws.cell(row=cr, column=6).value:
                                c_ai = ws.cell(row=cr, column=3)
                                c_ao = ws.cell(row=cr, column=4)
                                c_di = ws.cell(row=cr, column=5)
                                c_do = ws.cell(row=cr, column=6)
                                
                                if not isinstance(c_ai, MergedCell): c_ai.value = child.get("AI")
                                if not isinstance(c_ao, MergedCell): c_ao.value = child.get("AO")
                                if not isinstance(c_di, MergedCell): c_di.value = child.get("DI")
                                if not isinstance(c_do, MergedCell): c_do.value = child.get("DO")
                                child_matched = True
                                break 
                                
                if not child_matched:
                    missing_children.append(child)
                    
            if missing_children:
                orphans_to_insert.append({"parent": p_item, "children": missing_children, "is_partial": True})
                
        for group in ai_groups:
            if not group.get("best_match_row"):
                orphans_to_insert.append({"parent": group["parent"], "children": group["children"], "is_partial": False})

        if orphans_to_insert:
            total_insert = sum(1 + len(org["children"]) for org in orphans_to_insert)
                
            if total_insert > 0:
                ws.insert_rows(green_row_idx, total_insert)
                actualizar_formulas(ws, green_row_idx, total_insert)
                reparar_celdas_combinadas(ws, green_row_idx, total_insert)
                
                current_insert_row = green_row_idx
                for org in orphans_to_insert:
                    p_item = org["parent"]
                    titulo = p_item.get("Description", "")
                    qty_str = str(p_item.get("Quantity", "1")).strip()
                    parent_qty = int(qty_str) if qty_str.isdigit() else 1
                    
                    cell_title = ws.cell(row=current_insert_row, column=2)
                    if not isinstance(cell_title, MergedCell): cell_title.value = titulo
                    
                    if not org["is_partial"]:
                        c_qty = ws.cell(row=current_insert_row, column=8)
                        if not isinstance(c_qty, MergedCell): c_qty.value = p_item.get("Quantity", "")
                        
                    clonar_estilo_columna(ws, green_row_idx - 1, current_insert_row, es_padre=True)
                    current_insert_row += 1
                    
                    for child in org["children"]:
                        ai_c_desc = child.get("Description", "")
                        ws.cell(row=current_insert_row, column=2, value=ai_c_desc)
                        ws.cell(row=current_insert_row, column=3, value=child.get("AI", ""))
                        ws.cell(row=current_insert_row, column=4, value=child.get("AO", ""))
                        ws.cell(row=current_insert_row, column=5, value=child.get("DI", ""))
                        ws.cell(row=current_insert_row, column=6, value=child.get("DO", ""))
                        
                        cat_item = get_catalog_data(ai_c_desc)
                        part_no = child.get("Part No.", "")
                        if not part_no and cat_item: part_no = cat_item.get("Part", "")
                        
                        base_labour = int(cat_item.get("Labour", 0)) if cat_item else 0
                        calc_labour = base_labour * parent_qty
                        
                        c_part = ws.cell(row=current_insert_row, column=9)
                        c_labour = ws.cell(row=current_insert_row, column=12)
                        if not isinstance(c_part, MergedCell) and part_no: c_part.value = part_no
                        if not isinstance(c_labour, MergedCell) and calc_labour > 0: c_labour.value = calc_labour
                        
                        clonar_estilo_columna(ws, green_row_idx - 1, current_insert_row, es_padre=False)
                        current_insert_row += 1

        limpiar_nombres_feos(ws, start_row, green_row_idx + (total_insert if orphans_to_insert else 0))

    else:
        for r in range(start_row, ws.max_row + 1):
            for c in range(2, 13):
                cell = ws.cell(row=r, column=c)
                if not isinstance(cell, MergedCell):
                    cell.value = None
        
        aggregated_io = {}
        ai_groups_io = []
        current_ai_parent_io = None
        
        for item in datos:
            if is_parent(item):
                current_ai_parent_io = {"parent": item, "children": []}
                ai_groups_io.append(current_ai_parent_io)
            else:
                if current_ai_parent_io:
                    current_ai_parent_io["children"].append(item)

        for group in ai_groups_io:
            p_item = group["parent"]
            qty_str = str(p_item.get("Quantity", "1")).strip()
            p_qty = int(qty_str) if qty_str.isdigit() else 1
            
            for child in group["children"]:
                desc = str(child.get("Description", "")).strip()
                cat_item = get_catalog_data(desc)
                part_no = str(child.get("Part No.", "")).strip()
                if not part_no and cat_item:
                    part_no = cat_item.get("Part", "")
                
                agg_key = part_no.strip().lower() if part_no.strip() else desc.strip().lower()
                if not agg_key:
                    continue
                    
                if agg_key not in aggregated_io:
                    clean_desc = desc
                    if part_no:
                        if part_no.lower() == "volt free contacts":
                            clean_desc = "Software/Relay Interfaces (Enable/Status/Fault)"
                        elif part_no.lower() == "device by others":
                            clean_desc = "3rd Party Device Interface"
                        else:
                            for k, v in component_catalog.items():
                                if v.get("Part", "").strip().lower() == part_no.strip().lower():
                                    clean_desc = k.replace("Boiler ", "").replace("Chiller ", "").replace("Header ", "").replace("Pump ", "")
                                    break
                    
                    def get_val(val):
                        v = str(val).strip()
                        return int(float(v)) if v.replace('.', '', 1).isdigit() else 0
                        
                    base_ai = cat_item.get("AI", 0) if cat_item else (1 if get_val(child.get("AI", 0)) > 0 else 0)
                    base_ao = cat_item.get("AO", 0) if cat_item else (1 if get_val(child.get("AO", 0)) > 0 else 0)
                    base_di = cat_item.get("DI", 0) if cat_item else (1 if get_val(child.get("DI", 0)) > 0 else 0)
                    base_do = cat_item.get("DO", 0) if cat_item else (1 if get_val(child.get("DO", 0)) > 0 else 0)
                    base_labour = cat_item.get("Labour", 0) if cat_item else 0
                                
                    aggregated_io[agg_key] = {
                        "Description": clean_desc,
                        "Part No.": part_no,
                        "Quantity": 0,
                        "AI": base_ai, "AO": base_ao, "DI": base_di, "DO": base_do,
                        "Labour_Base": base_labour
                    }
                
                def get_val(val):
                    v = str(val).strip()
                    return int(float(v)) if v.replace('.', '', 1).isdigit() else 0
                
                ai_tot = get_val(child.get("AI", 0))
                ao_tot = get_val(child.get("AO", 0))
                di_tot = get_val(child.get("DI", 0))
                do_tot = get_val(child.get("DO", 0))
                
                b_ai = aggregated_io[agg_key]["AI"]
                b_ao = aggregated_io[agg_key]["AO"]
                b_di = aggregated_io[agg_key]["DI"]
                b_do = aggregated_io[agg_key]["DO"]
                
                pieces = 0
                if b_ai > 0: pieces = ai_tot // b_ai
                elif b_ao > 0: pieces = ao_tot // b_ao
                elif b_di > 0: pieces = di_tot // b_di
                elif b_do > 0: pieces = do_tot // b_do
                
                if pieces == 0:
                    pieces = max(ai_tot, ao_tot, di_tot, do_tot)
                    if pieces == 0: pieces = p_qty
                    
                aggregated_io[agg_key]["Quantity"] += pieces

        datos_io_procesados = []
        datos_io_procesados.append({
            "Description": "Consolidated Bill of Materials (BOM)",
            "Part No.": "", "Piece_Qty": "", "AI": "", "AO": "", "DI": "", "DO": "", "MCC": "", "Labour": ""
        })
        
        sorted_keys = sorted(aggregated_io.keys(), key=lambda k: aggregated_io[k]["Description"])
        
        for k in sorted_keys:
            data = aggregated_io[k]
            tot_qty = data["Quantity"]
            base_labour = data["Labour_Base"]
            calc_labour = (base_labour * tot_qty) if base_labour > 0 and tot_qty > 0 else ""
            
            datos_io_procesados.append({
                "Description": data["Description"],
                "Part No.": data["Part No."],
                "AI": data["AI"] if data["AI"] > 0 else "",
                "AO": data["AO"] if data["AO"] > 0 else "",
                "DI": data["DI"] if data["DI"] > 0 else "",
                "DO": data["DO"] if data["DO"] > 0 else "",
                "MCC": "",
                "Piece_Qty": tot_qty if tot_qty > 0 else "",
                "Labour": calc_labour
            })

        for idx, row_data in enumerate(datos_io_procesados):
            current_row = start_row + idx
            desc = str(row_data.get("Description", "")).strip()
            part = str(row_data.get("Part No.", "")).strip()
            es_header = bool(desc) and not bool(part) and not any(str(row_data.get(x, "")).strip() for x in ["AI", "AO", "DI", "DO"])

            col_map = {
                2: row_data.get("Description", ""), 
                3: row_data.get("AI", ""), 4: row_data.get("AO", ""),
                5: row_data.get("DI", ""), 6: row_data.get("DO", ""), 
                7: row_data.get("MCC", ""),
                8: row_data.get("Piece_Qty", ""),
                9: row_data.get("Part No.", ""),
                12: row_data.get("Labour", "")
            }
            
            for col_num, val in col_map.items():
                try:
                    cell = ws.cell(row=current_row, column=col_num)
                    if not isinstance(cell, MergedCell):
                        cell.value = val
                except AttributeError:
                    pass
            
            clonar_estilo_columna(ws, start_row, current_row, es_padre=es_header)

        last_row = start_row + len(datos_io_procesados) - 1
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

if st.button("Generate Points List"):
    if not api_key or not project_description:
        st.warning("Please ensure you have entered your API Key and provided a project description.")
    else:
        with st.spinner("Engineering system points and intelligently routing to template..."):
            prompt = f"""
            You are an expert BEMS estimator working for DC Controls. Generate a Points List based ONLY on the explicit systems requested.
            
            ENGINEERING RULES (Known Systems):
            {json.dumps(engineering_rules, indent=2)}
            
            TECHNICAL CATALOG:
            {json.dumps(component_catalog, indent=2)}
            
            CRITICAL FORMATTING INSTRUCTIONS:
            1. Create a HEADER ROW for each main equipment group. YOU MUST USE THE EXACT KEY FROM THE ENGINEERING RULES DICTIONARY AS THE "Description". DO NOT INVENT NAMES.
            2. Below the header row, list its components based EXACTLY on the ENGINEERING RULES.
            3. CRITICAL RULE: "MCC" MUST be left completely blank ("") for ALL rows (both parent and child). Do NOT invent MCC values like 1.0kW.
            4. ONLY include equipment explicitly requested by the user. Do NOT add unrequested systems.
            5. IF A SYSTEM IS NOT IN THE RULES, infer standard BEMS components for it (Enable DO, Status DI, Fault DI).
            6. CRITICAL CALCULATION: For each component, multiply base AI, AO, DI, DO by the main equipment quantity. 
            7. You do NOT need to calculate or output "Part No." or "Labour At 20%". Python will do it automatically.
            8. Leave IOs as completely empty strings ("") if the value is 0. Do NOT output a 0.
            
            User description: "{project_description}"
            
            Return ONLY a JSON array.
            """
            try:
                response = model.generate_content(prompt)
                json_text = response.text.strip().replace("```json", "").replace("```", "")
                materials_data = json.loads(json_text)
                
                for item in materials_data:
                    for key, val in item.items():
                        if val == 0 or val == "0" or val == "0.0":
                            item[key] = ""
                
                buffer_full = crear_excel_formateado(materials_data, project_name, es_io_schedule=False)
                # SE PASA LA DATA PURA Y DURA DIRECTO AL BOM (Sin filtros estúpidos que borren los sensores)
                buffer_filtrado = crear_excel_formateado(materials_data, project_name, es_io_schedule=True)
                
                st.session_state.generado = True
                st.session_state.buffer_full = buffer_full
                st.session_state.buffer_filtrado = buffer_filtrado
                st.session_state.materials_data = materials_data
                st.session_state.nombre_archivo = f"Quotation_{project_name.replace(' ', '_')}.xlsx"
                st.session_state.nombre_io = f"IO_Schedule_{project_name.replace(' ', '_')}.xlsx"

            except Exception as e:
                st.error(f"Error processing template: {e}")

if st.session_state.generado:
    st.success("Documents generated successfully!")
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        st.download_button("📄 Download Official Quotation (Full)", data=st.session_state.buffer_full, file_name=st.session_state.nombre_archivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_btn2:
        st.download_button("🔌 Download I/O Schedule Only (>0)", data=st.session_state.buffer_filtrado, file_name=st.session_state.nombre_io, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
