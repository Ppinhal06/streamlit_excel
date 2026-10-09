import streamlit as st
import pandas as pd
import io
import json
import copy
import re
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

# === LA LISTA 100% PURA DE TU EXCEL ===
allowed_systems = [
    "Common LPHW Devices",
    "Hot Water Generator  No.1",
    "Boiler",
    "LPHW Pressurisation Unit",
    "Dosing Unit",
    "Primary LPHW Single Pumps",
    "Calorifier",
    "Gas Detection Panel",
    "Storage Tank - Mains Water",
    "Booster Unit - Mains Water",
    "Storage Tank - Cold Water",
    "Booster Unit - Cold Water",
    "Storage Tank - Fire Hose Reel",
    "Storage Tank - Rain Water",
    "Rain Water Harvesting",
    "Chiller",
    "Chilled Water Buffer Tank",
    "AHU",
    "Extract Fan",
    "Natural Ventilation",
    "Metering",
    "Generator",
    "Unit Heater",
    "Air Curtain",
    "AC Units",
    "Radiator Circuit",
    "FCU (Fan Coil Unit)"
]

# CATALOGO FIJO PARA EL BOM (Para calcular puntos en Python sin romper el Excel)
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
    "Chiller Flow Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Chiller Return Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Enable and Current Switch": {"Part": "RIBXKTF", "AI": 0, "AO": 0, "DI": 1, "DO": 1, "Labour": 50},
    "Radiator Enable": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Radiator Flow Temp Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Radiator Return Temp Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Radiator Control Valve Actuator": {"Part": "Valve 25mm", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50}
}

def get_catalog_match(desc, part):
    d = desc.lower()
    p = part.lower()
    for k, v in component_catalog.items():
        if k.lower() == d: return v
    if "immersion temperature" in d and "flow" in d: return component_catalog["LPHW Header Flow Immersion Temperature Sensor"]
    if "immersion temperature" in d and "return" in d: return component_catalog["LPHW Header Return Immersion Temperature Sensor"]
    if "frost stat" in d or "frost thermostat" in d: return component_catalog["Outside Frost Thermostat"]
    if "space temperature" in d or "room temp" in d: return component_catalog["Space Temperature Sensor"]
    if "control valve" in d or "actuator" in d: return {"Part": part if part else "Valve", "AI": 0, "AO": 1, "DI": 0, "DO": 0, "Labour": 50}
    if "enable" in d: return {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50}
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
        if "radiator" in s and "radiator" in desc: return r
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
        wb_quot = load_workbook("template_2.xlsx")
        ws_quot = wb_quot.active
        ws_quot['B1'] = f"Project: {nombre_proyecto}"
        ws_quot['C4'] = datetime.now().strftime("%d/%m/%Y")
        ws_quot['C5'] = "Generated by AI Estimator"
        
        req_row = ws_quot.max_row
        for r in range(7, ws_quot.max_row):
            desc = str(ws_quot.cell(row=r, column=2).value or "").strip()
            if desc in ["BMS Requirements", "Summary"]:
                req_row = r
                break
            cell = ws_quot.cell(row=r, column=8)
            if not isinstance(cell, MergedCell): cell.value = None

        for sys_name, qty in json_data.items():
            if qty <= 0: continue
            match_r = match_native_row(sys_name, ws_quot)
            if match_r:
                ws_quot.cell(row=match_r, column=8).value = qty
                if "chiller" in sys_name.lower():
                    next_desc = str(ws_quot.cell(row=match_r + 1, column=2).value or "").strip().lower()
                    if "chiller mains" in next_desc:
                        ws_quot.cell(row=match_r + 1, column=8).value = qty

        for r in range(7, req_row):
            cell = ws_quot.cell(row=r, column=2)
            if not isinstance(cell, MergedCell):
                val = str(cell.value or "").strip()
                if "XXXX" in val or "??" in val:
                    cell.value = val.split(" - ")[0].strip()

        buf_quot = io.BytesIO()
        wb_quot.save(buf_quot)

        wb_bom = load_workbook("template_2.xlsx")
        ws_bom = wb_bom.active
        ws_bom['B1'] = f"Project: {nombre_proyecto}"
        ws_bom['C4'] = datetime.now().strftime("%d/%m/%Y")
        ws_bom['C5'] = "Generated by AI Estimator - I/O Schedule Only"
        
        aggregated_io = {}
        for r in range(7, req_row):
            desc = str(ws_quot.cell(row=r, column=2).value or "").strip()
            part = str(ws_quot.cell(row=r, column=9).value or "").strip()
            qty_val = ws_quot.cell(row=r, column=8).value
            
            if desc and not part and qty_val and str(qty_val).isdigit() and int(qty_val) > 0:
                parent_qty = int(qty_val)
                for child_r in range(r + 1, req_row):
                    c_desc = str(ws_quot.cell(row=child_r, column=2).value or "").strip()
                    c_part = str(ws_quot.cell(row=child_r, column=9).value or "").strip()
                    c_qty = str(ws_quot.cell(row=child_r, column=8).value or "").strip()
                    
                    if not c_desc: break 
                    if not c_part and c_qty.isdigit(): break 
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
        ws_bom.print_area = ""
        delete_amount = ws_bom.max_row - delete_start + 1
        if delete_amount > 0:
            for mcr in list(ws_bom.merged_cells.ranges):
                if mcr.min_row >= delete_start: ws_bom.merged_cells.remove(mcr)
            ws_bom.delete_rows(idx=delete_start, amount=delete_amount)

        buf_bom = io.BytesIO()
        wb_bom.save(buf_bom)
        
        return buf_quot.getvalue(), buf_bom.getvalue()
        
    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
        return None, None

st.subheader("Project Details")
col1, col2 = st.columns([1, 2])
with col1:
    project_name = st.text_input("Project Name (Header):", placeholder="e.g. IDA Cavan - Building 2")
with col2:
    project_description = st.text_area(
        "Project Scope Description:", 
        placeholder="Example: We need a plant with 2 Boilers, 2 Chillers, 25 Radiator Circuits, and 4 FCUs."
    )

st.markdown("---")

if st.button("Generate Points List"):
    if not api_key or not project_description:
        st.warning("Please ensure you have entered your API Key and provided a project description.")
    else:
        with st.spinner("Analyzing project scope and assigning quantities..."):
            prompt = f'''
            You are a BEMS Estimator. Read the user description and count how many of the ALLOWED SYSTEMS are needed.
            
            ALLOWED SYSTEMS EXACT NAMES:
            {json.dumps(allowed_systems, indent=2)}
            
            RULES:
            1. ONLY output a JSON dictionary mapping the EXACT system name from the list to its integer quantity.
            2. DO NOT output children or points. JUST the system name and quantity.
            3. If a system requires common heating devices (e.g. Boilers), add "Common LPHW Devices": 1 to the JSON.
            4. ONLY use names exactly as they appear in the ALLOWED SYSTEMS list. DO NOT invent names or categories.
            
            User Description: "{project_description}"
            
            RETURN ONLY VALID JSON:
            '''
            try:
                response = model.generate_content(prompt)
                json_text = response.text.strip().replace("```json", "").replace("```", "")
                system_quantities = json.loads(json_text)
                
                # Guardamos las cantidades para usarlas en el PDF
                st.session_state.materials_data = system_quantities 
                
                buffer_full, buffer_filtrado = procesar_archivos(system_quantities, project_name)
                
                if buffer_full and buffer_filtrado:
                    st.session_state.generado = True
                    st.session_state.buffer_full = buffer_full
                    st.session_state.buffer_filtrado = buffer_filtrado
                    st.session_state.nombre_archivo = f"Quotation_{project_name.replace(' ', '_')}.xlsx"
                    st.session_state.nombre_io = f"IO_Schedule_{project_name.replace(' ', '_')}.xlsx"
                    st.session_state.autorizado = False # Resetea la autorización al generar uno nuevo

            except Exception as e:
                st.error(f"Error AI/JSON: {e}")

# --- SECCIÓN DE DESCARGAS Y VERIFICACIÓN PARA PDF ---
if st.session_state.generado:
    st.success("Documents generated successfully!")
    
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        st.download_button("Download Official Quotation (Excel)", data=st.session_state.buffer_full, file_name=st.session_state.nombre_archivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_btn2:
        st.download_button("Download Bill of Materials (Excel)", data=st.session_state.buffer_filtrado, file_name=st.session_state.nombre_io, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    st.markdown("---")
    
    st.subheader("🔒 Manager Approval & PDF Export")
    st.info("⚠️ Official PDF quotations cannot be printed without managerial authorization.")
    
    col_pass, col_btn_auth = st.columns([2, 1])
    with col_pass:
        manager_password = st.text_input("Enter Authorization Code (Manager Only):", type="password")
    
    # CONTRASEÑA DEL JEFE
    SECRET_PASSWORD = "Example1" 
    
    with col_btn_auth:
        st.write("") 
        st.write("")
        btn_authorize = st.button("Authorize & Unlock PDF")
        
    if btn_authorize:
        if manager_password == SECRET_PASSWORD:
            st.session_state.autorizado = True
            st.success("✅ Quotation Verified and Authorized by Management.")
        else:
            st.error("❌ Invalid Authorization Code. Access Denied.")
            st.session_state.autorizado = False

    if st.session_state.get("autorizado", False):
        try:
            from fpdf import FPDF
            
            pdf = FPDF(orientation="P", unit="mm", format="A4")
            pdf.add_page()
            
            # Header
            pdf.set_font("Arial", "B", 16)
            pdf.set_text_color(0, 51, 102) 
            pdf.cell(0, 10, "DC CONTROLS - OFFICIAL QUOTATION", ln=True, align="C")
            
            pdf.set_font("Arial", "", 10)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(0, 6, f"Project: {project_name}", ln=True)
            pdf.cell(0, 6, f"Date: {datetime.now().strftime('%B %d, %Y')}", ln=True)
            pdf.cell(0, 6, "Status: VERIFIED & APPROVED", ln=True)
            pdf.ln(10)
            
            # Tabla (Resumen de Sistemas)
            pdf.set_font("Arial", "B", 11)
            pdf.cell(100, 8, "System Description", border=1, fill=False)
            pdf.cell(40, 8, "Quantity", border=1, align="C", ln=True)
            
            pdf.set_font("Arial", "", 10)
            for sys_name, qty in st.session_state.materials_data.items():
                if qty > 0:
                    pdf.cell(100, 8, sys_name[:45], border=1) 
                    pdf.cell(40, 8, str(qty), border=1, align="C", ln=True)

            pdf.ln(15)
            pdf.set_font("Arial", "I", 8)
            pdf.cell(0, 5, "This is an authorized printable summary generated by the DC Controls BEMS Estimator.", ln=True, align="C")
            
            pdf_bytes = pdf.output(dest="S").encode("latin-1")
            
            st.download_button(
                label="Download Authorized PDF for Printing",
                data=pdf_bytes,
                file_name=f"Approved_Quotation_{project_name.replace(' ', '_')}.pdf",
                mime="application/pdf",
                type="primary"
            )
            
        except ImportError:
            st.error("⚠️ FPDF library is missing. Please ensure 'fpdf' is in your GitHub requirements.txt file.")
