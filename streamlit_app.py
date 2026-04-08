import streamlit as st
import pandas as pd
import pyproj
import simplekml
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# --- SETUP ---
# User agent is required for the address-finder to work
geolocator = Nominatim(user_agent="AucklandCouncilKMLTool")
rate_limited_geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)

COLOR_MAP = {
    "all bores": "ffff0000", "bores": "ffff0000",
    "all consents": "ff00ffff", "consents": "ff00ffff",
    "all incidents": "ff0000ff", "incidents": "ff0000ff",
    "hail": "ff800080"
}
FALLBACK_PALETTE = ["ffffff00", "ff00a5ff", "ff008000", "ffffc0cb", "ff2222a2", "ff808080"]

nztm = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)
nzmg = pyproj.Transformer.from_crs("epsg:27200", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    return (160 < lon < 185) and (-50 < lat < -30)

def get_smart_title(row, sheet_name):
    """Trims the title to be clean and relevant."""
    s = sheet_name.lower()
    val = ""
    
    if "hail" in s and 'PropertyAddress' in row:
        val = str(row['PropertyAddress']).split("Henderson")[0].strip()
    elif "bore" in s and 'BORE_ID' in row:
        val = f"Bore {row['BORE_ID']}"
    elif "incident" in s and 'INCIDENTNUMBER' in row:
        val = f"Incident {row['INCIDENTNUMBER']}"
    elif "consent" in s and 'CONSENT_NUMBER' in row:
        val = f"Consent {row['CONSENT_NUMBER']}"
    
    if not val:
        for col in ['PropertyAddress', 'Location', 'ID']:
            if col in row and pd.notna(row[col]): 
                val = str(row[col])
                break
    return val if val else "Site Detail"

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

        # Column Detection
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'nztmxcoord', 'x']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'nztmycoord', 'y']), None)
        addr_col = 'PropertyAddress' if 'PropertyAddress' in df.columns else \
                   next((c for c in df.columns if 'address' in c.lower() or 'location' in c.lower()), None)

        for _, row in df.iterrows():
            # Build Table (No border for HAIL)
            is_hail = "hail" in lookup
            html = f'<table border="{"0" if is_hail else "1"}" style="font-family:sans-serif; font-size:12px; border-collapse:collapse; width:300px;">'
            for col, val in row.items():
                if pd.notna(val) and str(val).lower() != 'n/a':
                    html += f'<tr><td style="padding:4px; background:#eee; font-weight:bold;">{col}</td><td>{val}</td></tr>'
            html += "</table>"

            lon, lat = None, None
            clean_addr = None

            # 1. TRY COORDINATES FIRST
            try:
                e_val = float(str(row[e_col]).replace(',', '')) if e_col and pd.notna(row[e_col]) else None
                n_val = float(str(row[n_col]).replace(',', '')) if n_col and pd.notna(row[n_col]) else None
                if e_val and n_val:
                    lon, lat = nztm.transform(e_val, n_val)
                    if not is_near_nz(lon, lat):
                        lon, lat = nzmg.transform(e_val, n_val)
            except: pass

            # 2. TRY ADDRESS GEOCODING IF NO COORDS (Like HAIL sites)
            if (not lon or not lat) and addr_col and pd.notna(row[addr_col]):
                clean_addr = str(row[addr_col]).strip()
                full_search = f"{clean_addr}, Auckland, New Zealand"
                try:
                    # Attempt to find the house on the map via Python
                    location = rate_limited_geocode(full_search)
                    if location:
                        lon, lat = location.longitude, location.latitude
                except: pass

            # 3. CREATE PLACEMARK
            if lon and lat and is_near_nz(lon, lat):
                pnt = folder.newpoint(name=get_smart_title(row, sheet_name), description=html, coords=[(lon, lat)])
            elif clean_addr:
                # Last resort: Add a point with NO coordinates so it doesn't go to Africa
                pnt = folder.newpoint(name=get_smart_title(row, sheet_name), description=html)
                pnt.address = f"{clean_addr}, Auckland, New Zealand"
                pnt.geometry = None # Kills the 0,0 default
            else:
                continue

            pnt.style.iconstyle.color = target_color
            pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
            pnt.style.iconstyle.scale = 0.8

    return kml.kml()

# --- STREAMLIT ---
st.title("🌍 Auckland Council KML Pro")
st.markdown("Geocoding address-only sites (like HAIL) to prevent points in Africa.")
file = st.file_uploader("Upload Excel", type="xlsx")
if file:
    with st.spinner("Geocoding addresses... This may take a moment."):
        kml_str = process_excel_to_kml(file)
        st.download_button(
            "📥 Download KML",
            kml_str.encode("utf-8"),
            file_name="Henderson_Report.kml",
            mime="application/vnd.google-earth.kml+xml",
        )