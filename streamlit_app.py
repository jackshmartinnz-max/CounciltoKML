import streamlit as st
import pandas as pd
import pyproj
import simplekml
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import re
import io

# --- CONFIGURATION ---
geolocator = Nominatim(user_agent="Auckland_Universal_Mapper_v6", timeout=10)
geocode_service = RateLimiter(geolocator.geocode, min_delay_seconds=0.8)
nztm = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    if lon is None or lat is None: return False
    return (160 < lon < 185) and (-48 < lat < -32)

def get_smart_title(row):
    id_cols = ['BORE_ID', 'CONSENT_NUMBER', 'INCIDENTNUMBER', 'SAPSiteID', 'Reference', 'ID']
    for col in id_cols:
        if col in row and pd.notna(row[col]): return str(row[col])
    addr_cols = [c for c in row.index if 'address' in c.lower() or 'location' in c.lower()]
    if addr_cols and pd.notna(row[addr_cols[0]]):
        return str(row[addr_cols[0]]).split(',')[0].strip()
    return "Point"

def get_sheet_color(sheet_name):
    s = sheet_name.lower()
    if "bore" in s: return "ffff0000"  # Blue
    if "consent" in s: return "ff00ffff"  # Yellow
    if "incident" in s or "pollution" in s: return "ff0000ff"  # Red
    if "hail" in s or "characteristic" in s: return "ff800080"  # Purple
    return "ff00a5ff"  # Orange

def find_actual_header(df_raw):
    """
    Council files often have a title row before the headers.
    This scans the first 5 rows for coordinate keywords.
    """
    keywords = {'xcoord', 'ycoord', 'easting', 'northing', 'x', 'y', 'east', 'north'}
    for i in range(min(5, len(df_raw))):
        row_values = [str(val).lower() for val in df_raw.iloc[i].values if pd.notna(val)]
        if any(key in row_values for key in keywords):
            return i
    return 0

def process_project(file):
    xl = pd.ExcelFile(file)
    kml = simplekml.Kml()
    failed_rows = []
    stats = {"math": 0, "address": 0}
    
    for sheet_name in xl.sheet_names:
        # 1. Load raw to find the real header row
        df_raw = xl.parse(sheet_name, header=None)
        header_row_index = find_actual_header(df_raw)
        
        # 2. Reload with the correct header
        df = xl.parse(sheet_name, skiprows=header_row_index)
        folder = kml.newfolder(name=sheet_name)
        target_color = get_sheet_color(sheet_name)
        
        # Universal Column Detection
        e_col = next((c for c in df.columns if str(c).lower() in ['easting', 'nztmxcoord', 'xcoord', 'x', 'east']), None)
        n_col = next((c for c in df.columns if str(c).lower() in ['northing', 'nztmycoord', 'ycoord', 'y', 'north']), None)
        addr_col = next((c for c in df.columns if 'address' in str(c).lower() or 'location' in str(c).lower()), None)

        for _, row in df.iterrows():
            lon, lat = None, None
            found_by = None
            
            # Coordinate Math
            try:
                if e_col and n_col and pd.notna(row[e_col]) and pd.notna(row[n_col]):
                    e_val = float(str(row[e_col]).replace(',', ''))
                    n_val = float(str(row[n_col]).replace(',', ''))
                    # Swap logic if accidentally reversed
                    if e_val > 3000000: e_val, n_val = n_val, e_val 
                    lon, lat = nztm.transform(e_val, n_val)
                    if is_near_nz(lon, lat): found_by = "math"
            except: pass

            # Geocoding Fallback
            if found_by != "math" and addr_col and pd.notna(row[addr_col]):
                clean_addr = re.sub(r'\b\d{4}\b', '', str(row[addr_col]))
                query = f"{clean_addr}, Auckland, NZ"
                try:
                    location = geocode_service(query)
                    if location:
                        lon, lat = location.longitude, location.latitude
                        if is_near_nz(lon, lat): found_by = "address"
                except: pass

            if found_by:
                stats[found_by] += 1
                html = '<table border="1" style="font-family:sans-serif;font-size:12px;border-collapse:collapse;width:280px;">'
                for col, val in row.items():
                    if pd.notna(val): html += f'<tr><td style="background:#eee;font-weight:bold;padding:3px;">{col}</td><td>{val}</td></tr>'
                html += "</table>"
                pnt = folder.newpoint(name=get_smart_title(row), description=html, coords=[(lon, lat)])
                pnt.style.iconstyle.color = target_color
                pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
            else:
                row_dict = row.to_dict()
                row_dict['Original_Sheet'] = sheet_name
                failed_rows.append(row_dict)

    failed_df = pd.DataFrame(failed_rows)
    output_excel = io.BytesIO()
    if not failed_df.empty:
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            failed_df.to_excel(writer, index=False)
    
    return kml.kml(), output_excel.getvalue(), stats, len(failed_rows)

# --- UI ---
st.set_page_config(page_title="AKL Universal Mapper", layout="centered")
st.title("🌍 Universal Auckland Council KML Tool")
st.info("Updated to handle multi-line headers (Whenuapai/Kauri Rd format).")

file = st.file_uploader("Upload Council Export", type="xlsx")
if file:
    with st.spinner("Processing... Finding coordinates in multi-line sheets."):
        kml_data, excel_data, stats, fail_count = process_project(file)
        st.success(f"Mapping Complete! Math: {stats['math']} | Address: {stats['address']}")
        
        c1, c2 = st.columns(2)
        with c1: st.download_button("📥 Download KML Map", kml_data, file_name="Auckland_Map.kml", use_container_width=True)
        with c2: 
            if fail_count > 0: st.download_button("⚠️ Download Failed", excel_data, file_name="Review.xlsx", use_container_width=True)