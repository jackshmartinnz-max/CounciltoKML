import streamlit as st
import pandas as pd
import pyproj
import simplekml
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import re
import io

# --- CONFIGURATION ---
geolocator = Nominatim(user_agent="Auckland_Universal_Mapper_v11", timeout=10)
geocode_service = RateLimiter(geolocator.geocode, min_delay_seconds=0.8)
nztm = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    if lon is None or lat is None: return False
    return (160 < lon < 185) and (-48 < lat < -32)

def is_nztm_val(val):
    """Check if a number looks like an Auckland NZTM coordinate."""
    try:
        f = float(str(val).replace(',', '').strip())
        # Easting range ~1.5M-1.9M, Northing range ~5.8M-6.0M
        return (1500000 < f < 2000000) or (5800000 < f < 6100000)
    except: return False

def get_smart_title(row):
    id_cols = ['BORE_ID', 'CONSENT_NUMBER', 'INCIDENTNUMBER', 'SAPSiteID', 'ID']
    for col in row.index:
        if any(x == str(col).strip() for x in id_cols):
            if pd.notna(row[col]) and str(row[col]).strip() != "": return str(row[col])
    return "Point"

def process_project(file):
    xl = pd.ExcelFile(file)
    kml = simplekml.Kml()
    failed_rows = []
    stats = {"math": 0, "address": 0}
    
    keywords = {'xcoord', 'ycoord', 'easting', 'northing', 'x', 'y', 'sapsiteid', 'bore_id', 'property_address'}

    for sheet_name in xl.sheet_names:
        df_raw = xl.parse(sheet_name, header=None)
        folder = kml.newfolder(name=sheet_name)
        current_headers = None
        current_cols = {}

        for _, row in df_raw.iterrows():
            # 1. HEADER DETECTION
            row_vals = [str(v).strip() for v in row.values]
            row_lower = [v.lower() for v in row_vals]
            if any(key in row_lower for key in keywords):
                current_headers = [v if v != "" else f"Col_{idx}" for idx, v in enumerate(row_vals)]
                current_cols = {v.lower(): idx for idx, v in enumerate(current_headers)}
                continue
            
            if current_headers is None: continue

            lon, lat, found_by = None, None, None
            
            # 2. COORDINATE MATH (VIA HEADERS)
            e_idx = next((current_cols[k] for k in ['easting', 'xcoord', 'nztmxcoord', 'x'] if k in current_cols), None)
            n_idx = next((current_cols[k] for k in ['northing', 'ycoord', 'nztmycoord', 'y'] if k in current_cols), None)
            
            try:
                if e_idx is not None and n_idx is not None:
                    e_val = float(str(row[e_idx]).replace(',', '').strip())
                    n_val = float(str(row[n_idx]).replace(',', '').strip())
                    if e_val > 3000000: e_val, n_val = n_val, e_val
                    lon, lat = nztm.transform(e_val, n_val)
                    if is_near_nz(lon, lat): found_by = "math"
            except: pass

            # 3. COORDINATE MATH (BRUTE FORCE FALLBACK for missing headers)
            if not found_by:
                numeric_vals = []
                for v in row.values:
                    if is_nztm_val(v): numeric_vals.append(float(str(v).replace(',', '')))
                if len(numeric_vals) >= 2:
                    # Sort so smaller is Easting, larger is Northing
                    e_val, n_val = min(numeric_vals), max(numeric_vals)
                    lon, lat = nztm.transform(e_val, n_val)
                    if is_near_nz(lon, lat): found_by = "math"

            # 4. ADDRESS FALLBACK
            if not found_by:
                addr_val = None
                for v in row.values:
                    if pd.notna(v) and len(str(v)) > 10:
                        if any(w in str(v).upper() for w in [' ROAD', ' STREET', ' AVE', ' RD', ' ST']):
                            addr_val = str(v)
                            break
                if addr_val:
                    clean = re.sub(r'\b\d{4}\b', '', addr_val)
                    try:
                        loc = geocode_service(f"{clean}, Auckland, NZ")
                        if loc:
                            lon, lat = loc.longitude, loc.latitude
                            found_by = "address"
                    except: pass

            if found_by:
                stats[found_by] += 1
                temp_series = pd.Series(row.values, index=current_headers)
                
                # Build Table
                html = '<table border="1" style="font-size:11px; border-collapse:collapse; width:300px;">'
                for idx, val in enumerate(row.values):
                    h = current_headers[idx]
                    if pd.notna(val) and "Col_" not in h:
                        html += f'<tr><td style="background:#eee; font-weight:bold;">{h}</td><td>{val}</td></tr>'
                html += "</table>"

                pnt = folder.newpoint(name=get_smart_title(temp_series), description=html, coords=[(lon, lat)])
                pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
            elif not row.dropna().empty:
                failed_rows.append({current_headers[idx]: v for idx, v in enumerate(row.values)})

    output_excel = io.BytesIO()
    if failed_rows:
        pd.DataFrame(failed_rows).to_excel(output_excel, index=False)
    
    return kml.kml(), output_excel.getvalue(), stats, len(failed_rows)

# --- UI ---
st.set_page_config(page_title="AKL Mapper v11", layout="wide")
st.title("🌍 Auckland Council Universal Mapper")
file = st.file_uploader("Upload Excel", type="xlsx")
if file:
    with st.spinner("Executing Brute-Force Coordinate Recovery..."):
        kml_data, excel_data, stats, fail_count = process_project(file)
        st.metric("Points Found via Math", stats['math'])
        st.metric("Points Found via Address", stats['address'])
        st.download_button("📥 Download KML", kml_data, file_name="Map.kml")
        if fail_count > 0:
            st.download_button("⚠️ Download Failures", excel_data, file_name="Fail.xlsx")