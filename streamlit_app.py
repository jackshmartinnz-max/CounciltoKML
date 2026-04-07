import streamlit as st
import pandas as pd
import pyproj
import simplekml
from io import BytesIO

# --- CONFIGURATION & REGISTRY (Steps 2.1 & 2.2) ---
REGISTRY = {
    "all bores": {"type": "Boreholes", "color": "ffff0000"},       # Blue (aabbggrr)
    "bores": {"type": "Boreholes", "color": "ffff0000"},           # Blue
    "all consents": {"type": "Discharge consents", "color": "ff00ffff"}, # Yellow
    "consents": {"type": "Discharge consents", "color": "ff00ffff"},     # Yellow
    "all incidents": {"type": "Incidents / spills", "color": "ff0000ff"},# Red
    "incidents": {"type": "Incidents / spills", "color": "ff0000ff"},    # Red
    "hail": {"type": "HAIL sites", "color": "ff800080"}            # Purple
}

FALLBACK_COLORS = ["ffffff00", "ff00a5ff", "ff008000", "ffffc0cb", "ff2222a2", "ff808080"]
ICON_URL = "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"

# Setup Coordinate Transformer (NZTM2000 -> WGS84)
transformer = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)

def process_excel_to_kml(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    available_sheets = xl.sheet_names
    kml = simplekml.Kml()
    color_index = 0

    for sheet_name in available_sheets:
        df = xl.parse(sheet_name)
        lookup_name = sheet_name.strip().lower()
        
        # Routing logic
        if lookup_name in REGISTRY:
            info = REGISTRY[lookup_name]
            color, d_type = info["color"], info["type"]
        else:
            color, d_type = FALLBACK_COLORS[color_index % len(FALLBACK_COLORS)], "Other"
            color_index += 1

        folder = kml.newfolder(name=sheet_name)

        # Column Detection
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'eastings', 'x']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'northings', 'y']), None)
        addr_col = next((c for c in df.columns if 'address' in c.lower()), None)

        for _, row in df.iterrows():
            # HTML Table Styling (Step 6)
            border = "0" if d_type == "HAIL sites" else "1"
            html = f'<table border="{border}" style="font-family:sans-serif; font-size:12px; border-collapse:collapse; width:300px;">'
            for col, val in row.items():
                if pd.notna(val):
                    html += f'<tr><td style="padding:4px; background:#eee; font-weight:bold;">{col}</td>'
                    html += f'<td style="padding:4px;">{val}</td></tr>'
            html += '</table>'

            pnt = folder.newpoint(description=html)
            pnt.name = str(row.get('ID', row.get('Reference', 'Point')))
            pnt.style.iconstyle.icon.href = ICON_URL
            pnt.style.iconstyle.color = color
            pnt.style.iconstyle.scale = 0.8

            # Logic for Coordinates vs Address
            try:
                if e_col and n_col and pd.to_numeric(row[e_col], errors='coerce'):
                    lon, lat = transformer.transform(row[e_col], row[n_col])
                    pnt.coords = [(lon, lat)]
                elif addr_col and pd.notna(row[addr_col]):
                    pnt.address = f"{row[addr_col]}, Auckland, New Zealand"
            except:
                continue

    return kml.kml()

# --- STREAMLIT UI ---
st.set_page_config(page_title="Auckland Council KML Tool", page_icon="🌍")
st.title("🌍 Auckland Council KML Generator")
st.markdown("Convert multi-tab Excel files into color-coded Google Earth maps.")

uploaded_file = st.file_uploader("Upload Auckland Council Excel File", type="xlsx")

if uploaded_file:
    with st.spinner("Processing datasets..."):
        kml_data = process_excel_to_kml(uploaded_file)
        
        st.success("KML Generated Successfully!")
        st.download_button(
            label="📥 Download KML for Google Earth",
            data=kml_data,
            file_name=f"{uploaded_file.name.split('.')[0]}.kml",
            mime="application/vnd.google-earth.kml+xml"
        )

st.info("Note: This tool handles NZTM coordinates and addresses automatically.")