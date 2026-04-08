import streamlit as st
import pandas as pd
import pyproj
import simplekml
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import re

# --- CONFIGURATION ---
# We set the timeout inside Nominatim, not the RateLimiter
geolocator = Nominatim(user_agent="Auckland_Council_Mapper_Final", timeout=10)
# RateLimiter adds a delay between requests to avoid being blocked
geocode_service = RateLimiter(geolocator.geocode, min_delay_seconds=0.8)

nztm = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    """Ensures coordinates stay within the Auckland/NZ region."""
    if lon is None or lat is None: return False
    return (160 < lon < 185) and (-48 < lat < -32)

def get_smart_title(row, sheet_name):
    """Finds the best ID or short address for the sidebar."""
    for col in ['BORE_ID', 'CONSENT_NUMBER', 'INCIDENTNUMBER', 'SAPSiteID', 'Reference']:
        if col in row and pd.notna(row[col]): return str(row[col])
    
    addr_cols = [c for c in row.index if 'address' in c.lower() or 'location' in c.lower()]
    if addr_cols and pd.notna(row[addr_cols[0]]):
        return str(row[addr_cols[0]]).split(',')[0].strip()
    return "Point"

def process_excel(file):
    xl = pd.ExcelFile(file)
    kml = simplekml.Kml()
    
    COLOR_MAP = {
        "all bores": "ffff0000", "bores": "ffff0000",
        "all consents": "ff00ffff", "consents": "ff00ffff",
        "all incidents": "ff0000ff", "incidents": "ff0000ff",
        "hail": "ff800080"
    }
    FALLBACK_PALETTE = ["ffffff00", "ff00a5ff", "ff008000", "ffffc0cb"]
    color_index = 0

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        lookup = sheet_name.strip().lower()
        folder = kml.newfolder(name=sheet_name)
        
        target_color = COLOR_MAP.get(lookup, FALLBACK_PALETTE[color_index % len(FALLBACK_PALETTE)])
        if lookup not in COLOR_MAP: color_index += 1

        # Column Detection
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'nztmxcoord', 'x', 'east']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'nztmycoord', 'y', 'north']), None)
        addr_col = next((c for c in df.columns if 'address' in c.lower() or 'location' in c.lower()), None)

        for _, row in df.iterrows():
            lon, lat = None, None
            
            # 1. TRY COORDINATE MATH FIRST (Instant)
            try:
                if e_col and n_col and pd.notna(row[e_col]):
                    e_val = float(str(row[e_col]).replace(',', ''))
                    n_val = float(str(row[n_col]).replace(',', ''))
                    lon, lat = nztm.transform(e_val, n_val)
            except: pass

            # 2. TRY GEOCODING IF NO MATH (The HAIL sites)
            if not is_near_nz(lon, lat) and addr_col and pd.notna(row[addr_col]):
                # Strip postcodes and add Auckland context
                clean_addr = re.sub(r'\b\d{4}\b', '', str(row[addr_col]))
                query = f"{clean_addr}, Auckland, New Zealand"
                try:
                    location = geocode_service(query)
                    if location:
                        lon, lat = location.longitude, location.latitude
                except: pass

            # 3. ONLY CREATE PIN IF WE HAVE VALID DATA
            if is_near_nz(lon, lat):
                html = '<table border="1" style="font-family:sans-serif;font-size:12px;border-collapse:collapse;width:280px;">'
                for col, val in row.items():
                    if pd.notna(val): html += f'<tr><td style="background:#eee;font-weight:bold;padding:3px;">{col}</td><td>{val}</td></tr>'
                html += "</table>"

                pnt = folder.newpoint(name=get_smart_title(row, sheet_name), description=html, coords=[(lon, lat)])
                pnt.style.iconstyle.color = target_color
                pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
                pnt.style.iconstyle.scale = 0.8
            else:
                # If it's not in NZ and we can't find the address, we SKIP it.
                # No more pins in Africa!
                continue

    return kml.kml()

st.title("🌍 Universal Auckland Council KML Tool")
st.info("Addresses (HAIL) are being looked up automatically. This may take a minute.")

file = st.file_uploader("Upload Excel", type="xlsx")
if file:
    with st.spinner("Geocoding addresses... please wait."):
        kml_output = process_excel(file)
        st.success("Complete! Only verified Auckland locations were mapped.")
        st.download_button("📥 Download KML", kml_output, file_name="Auckland_Project.kml")