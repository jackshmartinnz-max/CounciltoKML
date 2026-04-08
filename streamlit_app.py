import streamlit as st
import pandas as pd
import pyproj
import simplekml
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import re

# --- CONFIGURATION ---
# The geographic "Box" for Auckland to prevent points jumping to other countries
AKL_BOUNDS = [(-37.3, 174.4), (-36.3, 175.3)] 

geolocator = Nominatim(user_agent="Auckland_Council_Project_Tool")
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=0.5, timeout=10)

COLOR_MAP = {
    "all bores": "ffff0000", "bores": "ffff0000",
    "all consents": "ff00ffff", "consents": "ff00ffff",
    "all incidents": "ff0000ff", "incidents": "ff0000ff",
    "hail": "ff800080"
}
FALLBACK_PALETTE = ["ffffff00", "ff00a5ff", "ff008000", "ffffc0cb", "ff2222a2", "ff808080"]

nztm = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)

def is_in_nz(lon, lat):
    return (170 < lon < 179) and (-38 < lat < -34)

def clean_address_for_search(addr):
    """Universal cleaner for NZ addresses."""
    if not addr or pd.isna(addr): return None
    s = str(addr)
    # Remove 4-digit postcodes (e.g., 0612) which often break searches
    s = re.sub(r'\b\d{4}\b', '', s)
    # Remove "Waitakere", "Manukau" etc if they follow "Auckland" to simplify
    s = s.replace("Waitakere", "").replace("Manukau", "").replace("North Shore", "")
    return s.strip()

def get_universal_title(row, sheet_name):
    id_cols = ['BORE_ID', 'CONSENT_NUMBER', 'INCIDENTNUMBER', 'SAPSiteID', 'Reference']
    for col in id_cols:
        if col in row and pd.notna(row[col]):
            return f"{row[col]}"
    addr_col = next((c for c in row.index if 'address' in c.lower() or 'location' in c.lower()), None)
    if addr_col:
        return str(row[addr_col]).split("Auckland")[0].strip()
    return "Map Point"

def process_excel_to_kml(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    kml = simplekml.Kml()
    color_index = 0

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        lookup = sheet_name.strip().lower()
        folder = kml.newfolder(name=sheet_name)
        
        target_color = COLOR_MAP.get(lookup, FALLBACK_PALETTE[color_index % len(FALLBACK_PALETTE)])
        if lookup not in COLOR_MAP: color_index += 1

        # Universal Column Detection
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'nztmxcoord', 'x', 'east']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'nztmycoord', 'y', 'north']), None)
        addr_col = next((c for c in df.columns if 'address' in c.lower() or 'location' in c.lower()), None)

        for _, row in df.iterrows():
            # Build Table
            html = '<table border="1" style="font-family:sans-serif; font-size:12px; border-collapse:collapse; width:300px;">'
            for col, val in row.items():
                if pd.notna(val) and str(val).lower() != 'n/a':
                    html += f'<tr><td style="padding:4px; background:#eee; font-weight:bold;">{col}</td><td>{val}</td></tr>'
            html += "</table>"

            lon, lat = None, None
            
            # 1. TRY COORDINATES (NZTM)
            try:
                e_val = float(str(row[e_col]).replace(',', '')) if e_col and pd.notna(row[e_col]) else None
                n_val = float(str(row[n_col]).replace(',', '')) if n_col and pd.notna(row[n_col]) else None
                if e_val and n_val:
                    lon, lat = nztm.transform(e_val, n_val)
            except: pass

            # 2. TRY CLEANED ADDRESS SEARCH
            if (not lon or not lat) and addr_col and pd.notna(row[addr_col]):
                search_addr = clean_address_for_search(row[addr_col])
                query = f"{search_addr}, Auckland, New Zealand"
                try:
                    location = geolocator.geocode(query, country_codes="nz", timeout=10)
                    if location:
                        lon, lat = location.longitude, location.latitude
                except: pass

            # 3. KML CREATION
            pnt = folder.newpoint(name=get_universal_title(row, sheet_name), description=html)
            if lon and lat and is_in_nz(lon, lat):
                pnt.coords = [(lon, lat)]
            elif addr_col and pd.notna(row[addr_col]):
                # Fallback: Let Google Earth geocode it
                pnt.address = f"{str(row[addr_col])}, Auckland, NZ"
                pnt.geometry = None 
            else:
                folder.features.remove(pnt)

            pnt.style.iconstyle.color = target_color
            pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
            pnt.style.iconstyle.scale = 0.8

    return kml.kml()

# --- STREAMLIT UI ---
st.title("🌍 Universal Auckland Council KML Tool")
file = st.file_uploader("Upload Excel", type="xlsx")

if file:
    with st.spinner("Processing... This version uses cleaned addresses to improve accuracy."):
        kml_str = process_excel_to_kml(file)
        st.success("Complete!")
        st.download_button("📥 Download KML", kml_str, file_name="Auckland_Project_Map.kml")