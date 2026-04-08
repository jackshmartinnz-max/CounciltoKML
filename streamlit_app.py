import streamlit as st
import pandas as pd
import pyproj
import simplekml
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import re
import io

# --- CONFIGURATION ---
geolocator = Nominatim(user_agent="Auckland_Universal_Mapper_v9", timeout=10)
geocode_service = RateLimiter(geolocator.geocode, min_delay_seconds=0.8)
nztm = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    if lon is None or lat is None: return False
    return (160 < lon < 185) and (-48 < lat < -32)

def get_smart_title(row):
    id_cols = ['BORE_ID', 'CONSENT_NUMBER', 'INCIDENTNUMBER', 'SAPSiteID', 'ConsentReference', 'Reference', 'ID']
    for col in row.index:
        if any(x == str(col).strip() for x in id_cols):
            if pd.notna(row[col]): return str(row[col])
    
    addr_cols = [c for c in row.index if any(x in str(c).lower() for x in ['address', 'location', 'primaryaddress'])]
    if addr_cols and pd.notna(row[addr_cols[0]]):
        return str(row[addr_cols[0]]).split(',')[0].strip()
    return "Point"

def get_sheet_color(sheet_name, current_headers):
    combined = (str(sheet_name) + " " + " ".join([str(h) for h in current_headers])).lower()
    if any(x in combined for x in ["bore", "well"]): return "ffff0000"
    if "consent" in combined: return "ff00ffff"
    if any(x in combined for x in ["incident", "pollution"]): return "ff0000ff"
    if any(x in combined for x in ["hail", "characteristic", "contaminated"]): return "ff800080"
    return "ff00a5ff"

def process_project(file):
    xl = pd.ExcelFile(file)
    kml = simplekml.Kml()
    failed_rows = []
    stats = {"math": 0, "address": 0}
    
    keywords = {'xcoord', 'ycoord', 'easting', 'northing', 'x', 'y', 'east', 'north', 'sapsiteid', 'consent_number', 'bore_id'}

    for sheet_name in xl.sheet_names:
        df_raw = xl.parse(sheet_name, header=None)
        folder = kml.newfolder(name=sheet_name)
        
        current_headers = None
        current_cols = {}

        for i, row in df_raw.iterrows():
            row_clean = [str(v).lower().strip() for v in row.values if pd.notna(v)]
            if any(key in row_clean for key in keywords):
                current_headers = [str(v).strip() if pd.notna(v) else f"Col_{idx}" for idx, v in enumerate(row.values)]
                current_cols = {str(h).lower().strip(): idx for idx, h in enumerate(current_headers)}
                continue
            
            if current_headers is None: continue

            lon, lat = None, None
            found_by = None
            
            e_key = next((k for k in ['easting', 'nztmxcoord', 'xcoord', 'x', 'east'] if k in current_cols), None)
            n_key = next((k for k in ['northing', 'nztmycoord', 'ycoord', 'y', 'north'] if k in current_cols), None)
            addr_key = next((k for k in current_cols if 'address' in k or 'location' in k), None)

            # 1. Math
            try:
                if e_key and n_key:
                    e_val = float(str(row[current_cols[e_key]]).replace(',', '').strip())
                    n_val = float(str(row[current_cols[n_key]]).replace(',', '').strip())
                    if e_val > 3000000: e_val, n_val = n_val, e_val 
                    lon, lat = nztm.transform(e_val, n_val)
                    if is_near_nz(lon, lat): found_by = "math"
            except: pass

            # 2. Address (Fixed Syntax Error here)
            if not found_by and addr_key:
                val = row[current_cols[addr_key]]
                if pd.notna(val) and len(str(val)) > 5:
                    address_clean = re.sub(r'\b\d{4}\b', '', str(val))
                    query = f"{address_clean}, Auckland, NZ"
                    try:
                        location = geocode_service(query)
                        if location:
                            lon, lat = location.longitude, location.latitude
                            if is_near_nz(lon, lat): found_by = "address"
                    except: pass

            if found_by:
                stats[found_by] += 1
                html = '<table border="1" style="font-family:sans-serif;font-size:12px;border-collapse:collapse;width:280px;">'
                for idx, col_name in enumerate(current_headers):
                    v = row[idx]
                    if pd.notna(v) and "Col_" not in str(col_name):
                        html += f'<tr><td style="background:#eee;font-weight:bold;padding:3px;">{col_name}</td><td>{v}</td></tr>'
                html += "</table>"
                
                temp_row = pd.Series(row.values, index=current_headers)
                pnt = folder.newpoint(name=get_smart_title(temp_row), description=html, coords=[(lon, lat)])
                pnt.style.iconstyle.color = get_sheet_color(sheet_name, current_headers)
                pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
            else:
                if row.dropna().empty: continue
                row_dict = {current_headers[idx]: val for idx, val in enumerate(row.values) if idx < len(current_headers)}
                row_dict['Original_Sheet'] = sheet_name
                failed_rows.append(row_dict)

    failed_df = pd.DataFrame(failed_rows)
    output_excel = io.BytesIO()
    if not failed_df.empty:
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            failed_df.to_excel(writer, index=False)
    
    return kml.kml(), output_excel.getvalue(), stats, len(failed_rows)

# --- UI ---
st.set_page_config(page_title="AKL Mapper v9", layout="centered")
st.title("🌍 Auckland Council Universal Mapper")

file = st.file_uploader("Upload Excel File", type="xlsx")
if file:
    with st.spinner("Processing stacked tables..."):
        kml_data, excel_data, stats, fail_count = process_project(file)
        st.success(f"Complete! Math: {stats['math']} | Address: {stats['address']}")
        
        c1, c2 = st.columns(2)
        with c1: st.download_button("📥 Download KML", kml_data, file_name="Auckland_Map.kml", use_container_width=True)
        with c2: 
            if fail_count > 0: st.download_button("⚠️ Review Failures", excel_data, file_name="Review.xlsx", use_container_width=True)