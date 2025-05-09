# Install needed packages before running:
# pip install streamlit pandas openpyxl

import streamlit as st
import pandas as pd
import json
from io import BytesIO

# Set Streamlit Page Config
st.set_page_config(page_title="Transformer Service Report App", layout="wide")

col1, col2 = st.columns([1, 5])

with col1:
    st.image("favicon.ico", width=80)

with col2:
    st.markdown("## Excellent Power Engineering - Transformer Report Generator")

# ---------------- Client Information ----------------
st.header("Client Information")
client_name = st.text_input("Client Name")
client_address = st.text_area("Client Address")
pincode = st.text_input("Pincode")
tr_number = st.text_input("TR Number")
date_of_test = st.date_input("Date of Test")
no_of_transformers = st.number_input("No. of Transformers", min_value=1, step=1)
no_with_oltc = st.number_input("No. of Transformers with OLTC", min_value=0, step=1, max_value=no_of_transformers)
no_without_oltc = no_of_transformers - no_with_oltc

st.markdown(f"**No. of Transformers without OLTC:** {no_without_oltc}")

st.divider()

# ---------------- Transformer Details ----------------
transformer_data = []

makes = [
    "ABB India", "BHEL", "Crompton Greaves", "Siemens India", "Kirloskar Electric", "Voltamp Transformers",
    "Schneider Electric", "Bharat Bijlee", "Areva", "Alstom", "Vijai Electricals",
    "UNIVERSAL POWER TRANSFORMERS", "ESSENNAR", "3si ecopowerllp", "Prolec", "Indotech"
]

capacities = [
    250, 315, 500, 630, 750, 800, 900, 1000, 1200, 1250, 1500, 1600, 1800, 2000,
    2200, 2500, 2800, 3000, 3500, 4000, 4500, 5000, 6000
]

oltc_makes = ["OLG", "CTR"]

st.header("Transformer Details")

for i in range(int(no_of_transformers)):
    st.subheader(f"Transformer {i+1}")

    with_oltc = st.selectbox(f"Does Transformer {i+1} have OLTC?", ("Yes", "No"), key=f"oltc_{i}")
    make = st.selectbox(f"Transformer Make {i+1}", makes, key=f"make_{i}")
    capacity = st.selectbox(f"Capacity (kVA) {i+1}", capacities, key=f"capacity_{i}")
    serial_number = st.text_input(f"Serial Number {i+1}", key=f"serial_{i}")
    year_of_manufacture = st.number_input(f"Year of Manufacture {i+1}", min_value=1900, max_value=2100, step=1, key=f"year_{i}")
    voltage_hv = st.number_input(f"Voltage HV (V) {i+1}", value=11000, key=f"voltage_hv_{i}")
    voltage_lv = st.number_input(f"Voltage LV (V) {i+1}", value=433, key=f"voltage_lv_{i}")
    impedance_percent = st.number_input(f"Impedance Voltage (%) {i+1}", key=f"impedance_{i}")
    oil_temperature = st.number_input(f"Oil Temperature (°C) {i+1}", value=32.0, key=f"oil_temp_{i}")
    electrode_gap = st.number_input(f"Electrode Gap (mm) {i+1}", value=2.5, key=f"gap_{i}")
    bdv_sample_1 = st.text_input(f"BDV Sample No 1 {i+1}", value="STOOD 40 kV PER MINUTE", key=f"bdv1_{i}")
    bdv_sample_2 = st.text_input(f"BDV Sample No 2 {i+1}", value="STOOD 40 kV PER MINUTE", key=f"bdv2_{i}")
    breakdown_voltage = st.text_input(f"Breakdown Voltage {i+1}", value="BROKE AT 60 kV", key=f"bdv_break_{i}")
    acidity_value = st.number_input(f"Acidity Value (mg of KOH/g) {i+1}", key=f"acid_{i}")
    permissible_limit = st.number_input(f"Permissible Limit (mg of KOH/g) {i+1}", value=0.30, key=f"acid_limit_{i}")

    oltc_details = {}
    if with_oltc == "Yes":
        st.markdown(f"**OLTC Details for Transformer {i+1}**")
        oltc_make = st.selectbox(f"OLTC Make {i+1}", oltc_makes, key=f"oltc_make_{i}")
        oltc_type = st.text_input(f"OLTC Type {i+1}", key=f"oltc_type_{i}")
        oltc_serial = st.text_input(f"OLTC Serial Number {i+1}", key=f"oltc_serial_{i}")
        oltc_year = st.number_input(f"OLTC Year of Manufacture {i+1}", min_value=1900, max_value=2100, step=1, key=f"oltc_year_{i}")
        oltc_voltage = st.number_input(f"OLTC Voltage HV (V) {i+1}", value=11000, key=f"oltc_voltage_{i}")
        oltc_current = st.number_input(f"OLTC Rated Current (A) {i+1}", key=f"oltc_current_{i}")
        oltc_oil_temp = st.number_input(f"OLTC Oil Temperature (°C) {i+1}", value=32.0, key=f"oltc_temp_{i}")
        oltc_gap = st.number_input(f"OLTC Electrode Gap (mm) {i+1}", value=2.5, key=f"oltc_gap_{i}")

        oltc_details = {
            "make": oltc_make,
            "type": oltc_type,
            "serial_number": oltc_serial,
            "year_of_manufacture": oltc_year,
            "voltage_hv": oltc_voltage,
            "rated_current_a": oltc_current,
            "oil_temperature_c": oltc_oil_temp,
            "electrode_gap_mm": oltc_gap
        }

    transformer_entry = {
        "transformer_id": f"Transformer {i+1}",
        "with_oltc": True if with_oltc == "Yes" else False,
        "make": make,
        "capacity_kva": capacity,
        "serial_number": serial_number,
        "year_of_manufacture": year_of_manufacture,
        "voltage_hv": voltage_hv,
        "voltage_lv": voltage_lv,
        "impedance_voltage_percent": impedance_percent,
        "oil_temperature_c": oil_temperature,
        "electrode_gap_mm": electrode_gap,
        "bdv_sample_1": bdv_sample_1,
        "bdv_sample_2": bdv_sample_2,
        "breakdown_voltage": breakdown_voltage,
        "acidity_value": acidity_value,
        "permissible_limit": permissible_limit,
        "oltc": oltc_details if with_oltc == "Yes" else None
    }
    
    transformer_data.append(transformer_entry)

# ---------------- Submit & Export ----------------
if st.button("✅ Submit & Export"):
    output = {
        "client_info": {
            "client_name": client_name,
            "client_address": client_address,
            "pincode": pincode,
            "tr_number": tr_number,
            "date_of_test": str(date_of_test),
            "no_of_transformers": no_of_transformers,
            "no_with_oltc": no_with_oltc,
            "no_without_oltc": no_without_oltc
        },
        "transformers": transformer_data
    }

    # JSON Export
    json_data = json.dumps(output, indent=4)
    st.download_button(
        label="📥 Download JSON",
        file_name="transformer_report_data.json",
        mime="application/json",
        data=json_data
    )

    # Excel Export
    client_info_df = pd.DataFrame([output["client_info"]])
    transformers_df = pd.json_normalize(output["transformers"])

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        client_info_df.to_excel(writer, sheet_name="Client Info", index=False)
        transformers_df.to_excel(writer, sheet_name="Transformers", index=False)
    buffer.seek(0)

    st.download_button(
        label="📥 Download Excel",
        file_name="transformer_report_data.xlsx",
        data=buffer,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

