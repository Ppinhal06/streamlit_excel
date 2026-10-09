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
import subprocess
import os

# --- 1. CONFIGURACIÓN DE PÁGINA (Debe ir hasta arriba) ---
st.set_page_config(page_title="BEMS Estimator Pro | DC Controls", layout="wide")

# Inicializar variables de estado
if "generado" not in st.session_state:
    st.session_state.generado = False
if "excel_revisado" not in st.session_state:
    st.session_state.excel_revisado = False
if "autorizado" not in st.session_state:
    st.session_state.autorizado = False

def marcar_descargado():
    st.session_state.excel_revisado = True

# --- 2. BARRA LATERAL (SIDEBAR) CORPORATIVA ---
with st.sidebar:
    # Lógica para cargar el logo automáticamente si existe en el servidor
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    elif os.path.exists("logo.jpg"):
        st.image("logo.jpg", use_container_width=True)
    else:
        # Texto de respaldo si no se encuentra la imagen
        st.markdown("<h2 style='text-align: center; color: #003366;'>DC CONTROLS</h2>", unsafe_allow_html=True)
        
    st.markdown("<p style='text-align: center; color: gray;'>BEMS Estimation Engine</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    st.header("System Access")
    api_key = st.text_input("Enter API Key (Gemini):", type="password", help="Requires authorized Gemini API token.")

# --- 3. ENCABEZADO PRINCIPAL ---
st.markdown("<h1 style='text-align: center; color: #003366;'>Automated BEMS Quotation System</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 18px; color: #555;'>AI-Powered I/O Schedule and Cost Estimator</p>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash') 

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

# CATALOGO FIJO PARA EL BOM
component_catalog = {
    "LPHW Header Flow Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "LPHW Header Return Immersion Temperature Sensor": {"Part": "TTI-S Brass Pocket", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Outside Frost Thermostat": {"Part": "DBET-23U", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Outside Temperature Sensor": {"Part": "TB/TO", "AI": 1, "AO": 0, "DI": 0, "DO": 0, "Labour": 50},
    "Immersion Frost Thermostat": {"Part": "DBTV-2U", "AI": 0, "AO": 0, "DI": 1, "DO": 0, "Labour": 50},
    "Boiler Enable": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI": 0, "DO": 1, "Labour": 50},
    "Boiler Common Fault": {"Part": "Volt Free Contacts", "AI": 0, "AO": 0, "DI":
