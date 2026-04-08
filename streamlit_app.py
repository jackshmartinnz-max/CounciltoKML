import streamlit as st
import pandas as pd
import pyproj
import simplekml
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import re

# --- CONFIGURATION ---
# Set the timeout globally here
geolocator = Nominatim(user_agent="Auckland_Universal_Mapper", timeout=10)
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.0)

# Colors for common Auckland Council sheets
COLOR_MAP = {
    "all bores": "ffff0000", "bores": "ffff0000",
    "all consents": "ff00ffff", "consents": "ff00ffff",
    "all incidents": "ff0000ff", "incidents": "ff0000ff",
    "hail": "ff800080"
}
FALLBACK_PALETTE = ["ffffff00", "ff00a5ff", "ff008000", "ffffc0cb", "ff2222a2", "ff808080"]

nztm = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    """Ensures coordinates stay within the NZ region."""
    return (160 < lon < 185) and (-48 < lat < -32)

def clean_address(addr):
    """Universal cleaner: Removes postcodes and tech noise."""
    if not addr or pd.isna(addr): return None
    s = str(addr)
    s = re.sub(r'\b\d{4}\b', '', s) # Removes 4-digit postcodes
    return s.strip()

def get_smart_title(row, sheet_name):
    """Finds the best ID or short address for the sidebar."""
    for col in ['BORE_ID', 'CONSENT_NUMBER', 'INCIDENTNUMBER', 'SAPSiteID', 'Reference']:
        if col in row and pd.notna(row[col]): return str(row[col])
    
    # Fallback to short street address
    addr_cols = [c for c in row.index if 'address' in c.lower() or 'location' in c.lower()]
    if addr_cols and pd.notna(row[addr_cols[0]]):
        return str(row[addr_cols[0]]).split(',')[0].replace("Auckland", "").strip()
    return "Point"

def process_excel(file):
    xl = pd.ExcelFile(file)
    kml = simplekml.Kml()
    color_index = 0

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        lookup = sheet_name.strip().lower()
        folder = kml.newfolder(name=sheet_name)
        
        target_color = COLOR_MAP.get(lookup, FALLBACK_PALETTE[color_index % len(FALLBACK_PALETTE)])
        if lookup not in COLOR_MAP: color_index += 1

        # Detect Columns
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'nztmxcoord', 'x', 'east']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'nztmycoord', 'y', 'north']), None)
        addr_col = next((c for c in df.columns if 'address' in c.lower() or 'location' in c.lower()), None)

        for _, row in df.iterrows():
            # Popup Content
            html = '<table border="1" style="font-family:sans-serif;font-size:12px;border-collapse:collapse;width:280px;">'
            for col, val in row.items():
                if pd.notna(val): html += f'<tr><td style="background:#eee;font-weight:bold;">{col}</td><td>{val}</td></tr>'
            html += "</table>"

            pnt = folder.newpoint(name=get_smart_title(row, sheet_name), description=html)
            
            lon, lat = None, None
            # 1. Try NZTM Math
            try:
                if e_col and n_col:
                    e_val = float(str(row[e_col]).replace(',', ''))
                    n_val = float(str(row[n_col]).replace(',', ''))
                    lon, lat = nztm.transform(e_val, n_val)
            except: pass

            # 2. Try Geocoding if math failed (Common for HAIL)
            if (not lon or not lat) and addr_col and pd.notna(row[addr_col]):
                search_q = f"{clean_address(row[addr_col])}, Auckland, New Zealand"
                try:
                    loc = geolocator.geocode(search_q)
                    if loc:
                        lon, lat = loc.longitude, loc.latitude
                except: pass

            # 3. Apply coordinates or set as searchable address
            if lon and lat and is_near_nz(lon, lat):
                pnt.coords = [(lon, lat)]
            elif addr_col and pd.notna(row[addr_col]):
                pnt.address = f"{row[addr_col]}, Auckland, NZ"
                pnt.geometry = None # Kills the 'Africa' 0,0 default
            else:
                folder.features.remove(pnt)

            pnt.style.iconstyle.color = target_color
            pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
            pnt.style.iconstyle.scale = 0.8

    return kml.kml()

st.title("🌍 Universal Auckland Council KML Tool")
file = st.file_uploader("Upload Excel", type="xlsx")
if file:
    with st.spinner("Processing... This takes about 1 second per address."):
        output = process_excel(file)
        st.success("Done!")
        st.download_button("📥 Download KML", output, file_name="Auckland_Project.kml")