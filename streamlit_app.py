import streamlit as st
import pandas as pd
import pyproj
import simplekml
from io import BytesIO

# --- CONFIGURATION & REGISTRY ---
REGISTRY = {
    "all bores": {"type": "Boreholes", "color": "ffff0000"},       # Blue
    "bores": {"type": "Boreholes", "color": "ffff0000"},
    "all consents": {"type": "Discharge consents", "color": "ff00ffff"}, # Yellow
    "all incidents": {"type": "Incidents / spills", "color": "ff0000ff"},# Red
    "hail": {"type": "HAIL sites", "color": "ff800080"}            # Purple
}
FALLBACK_COLORS = ["ffffff00", "ff00a5ff", "ff008000", "ffffc0cb", "ff2222a2", "ff808080"]
ICON_URL = "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"

# --- THE FIX: FORCED AXIS ORDER ---
# EPSG:2193 is NZTM2000 (New). EPSG:27200 is NZMG (Old). 
# always_xy=True ensures we always treat input as (Easting, Northing)
nztm_transformer = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)

def process_excel_to_kml(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    kml = simplekml.Kml()
    color_index = 0

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        lookup_name = sheet_name.strip().lower()
        
        # Routing
        if lookup_name in REGISTRY:
            color, d_type = REGISTRY[lookup_name]["color"], REGISTRY[lookup_name]["type"]
        else:
            color, d_type = FALLBACK_COLORS[color_index % len(FALLBACK_COLORS)], "Other"
            color_index += 1

        folder = kml.newfolder(name=sheet_name)

        # Better Column Detection (looks for E/N or X/Y)
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'eastings', 'x', 'east']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'northings', 'y', 'north']), None)
        addr_col = next((c for c in df.columns if 'address' in c.lower()), None)

        for _, row in df.iterrows():
            # Build HTML Table
            border = "0" if d_type == "HAIL sites" else "1"
            html = f'<table border="{border}" style="font-family:sans-serif; font-size:12px; border-collapse:collapse; width:300px;">'
            for col, val in row.items():
                if pd.notna(val):
                    html += f'<tr><td style="padding:4px; background:#eee; font-weight:bold;">{col}</td><td>{val}</td></tr>'
            html += '</table>'

            pnt = folder.newpoint(description=html)
            pnt.name = str(row.get('ID', row.get('Reference', 'Point')))
            pnt.style.iconstyle.icon.href = ICON_URL
            pnt.style.iconstyle.color = color
            pnt.style.iconstyle.scale = 0.8

            # COORDINATE PROCESSING
            try:
                e_val = pd.to_numeric(row[e_col], errors='coerce') if e_col else None
                n_val = pd.to_numeric(row[n_col], errors='coerce') if n_col else None

                if e_val and n_val:
                    # TRANSFORM: Input Easting(X), Northing(Y) -> Output Lon, Lat
                    lon, lat = nztm_transformer.transform(e_val, n_val)
                    pnt.coords = [(lon, lat)]
                elif addr_col and pd.notna(row[addr_col]):
                    pnt.address = f"{row[addr_col]}, Auckland, New Zealand"
            except Exception:
                continue

    return kml.kml()

# --- STREAMLIT UI ---
st.title("🌍 Auckland KML Generator (Fixed Projection)")
uploaded_file = st.file_uploader("Upload Excel", type="xlsx")

if uploaded_file:
    kml_output = process_excel_to_kml(uploaded_file)
    st.download_button("📥 Download KML", kml_output, file_name="Auckland_Map.kml")