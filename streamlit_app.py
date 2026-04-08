import streamlit as st
import pandas as pd
import pyproj
import simplekml
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import re
import io

# --- CONFIGURATION ---
geolocator = Nominatim(user_agent="Auckland_Universal_Mapper_v12", timeout=10)
geocode_service = RateLimiter(geolocator.geocode, min_delay_seconds=0.8)
nztm = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    if lon is None or lat is None: return False
    return (160 < lon < 185) and (-48 < lat < -32)

def is_nztm_val(val):
    try:
        f = float(str(val).replace(',', '').strip())
        return (1500000 < f < 2000000) or (5800000 < f < 6100000)
    except: return False

def get_smart_title(row):
    id_cols = ['BORE_ID', 'CONSENT_NUMBER', 'INCIDENTNUMBER', 'SAPSiteID', 'ID']
    for col in row.index:
        if any(x == str(col).strip() for x in id_cols):
            if pd.notna(row[col]) and str(row[col]).strip() != "": return str(row[col])
    return "Point"

def get_muted_color(sheet_name, headers):
    """Returns 'Dull' professional KML colors (AABBGGRR)."""
    combined = (str(sheet_name) + " " + " ".join([str(h) for h in headers])).lower()
    
    # Bores -> Muted Blue (Steel Blue)
    if any(x in combined for x in ["bore", "well"]): 
        return "ffb08446" 
    # Consents -> Muted Teal/Cyan
    if "consent" in combined: 
        return "ffa0a000"
    # Incidents/Pollution -> Muted Red (Terracotta)
    if any(x in combined for x in ["incident", "pollution"]): 
        return "ff4b4bc8"
    # HAIL/Contaminated -> Muted Purple (Dusty Lavender)
    if any(x in combined for x in ["hail", "characteristic", "contaminated"]): 
        return "ff966496"
    
    # Default -> Muted Grey/Orange
    return "ff5a96f0"

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
            row_vals = [str(v).strip() for v in row.values]
            row_lower = [v.lower() for v in row_vals]
            
            if any(key in row_lower for key in keywords):
                current_headers = [v if v != "" else f"Col_{idx}" for idx, v in enumerate(row_vals)]
                current_cols = {v.lower(): idx for idx, v in enumerate(current_headers)}
                continue
            
            if current_headers is None: continue

            lon, lat, found_by = None, None, None
            
            # 1. Coordinate Math
            e_idx = next((current_cols[k] for k in ['easting', 'xcoord', 'nztmxcoord', 'x'] if k in current_cols), None)
            n_idx = next((current_cols[k] for k in ['northing', 'ycoord', 'nztmycoord', 'y'] if k in current_cols), None)
            
            try:
                if e_idx is not None and n_idx is not None:
                    e_val = float(str(row[e_idx]).replace(',', '').strip())
                    n_val = float(str(row[n_idx]).replace(',', '').strip())
                    if e_val > 0:
                        if e_val > 3000000: e_val, n_val = n_val, e_val
                        lon, lat = nztm.transform(e_val, n_val)
                        if is_near_nz(lon, lat): found_by = "math"
            except: pass

            # 2. Brute Force Fallback
            if not found_by:
                nums = [float(str(v).replace(',', '')) for v in row.values if is_nztm_val(v)]
                if len(nums) >= 2:
                    lon, lat = nztm.transform(min(nums), max(nums))
                    if is_near_nz(lon, lat): found_by = "math"

            # 3. Address Fallback
            if not found_by:
                addr_val = next((str(v) for v in row.values if pd.notna(v) and len(str(v)) > 10 and any(w in str(v).upper() for w in [' ROAD', ' STREET', ' AVE', ' RD', ' ST'])), None)
                if addr_val:
                    clean = re.sub(r'\b\d{4}\b', '', addr_val)
                    try:
                        loc = geocode_service(f"{clean}, Auckland, NZ")
                        if loc and is_near_nz(loc.longitude, loc.latitude):
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
                        html += f'<tr><td style="background:#eee; font-weight:bold; padding:2px;">{h}</td><td style="padding:2px;">{val}</td></tr>'
                html += "</table>"

                pnt = folder.newpoint(name=get_smart_title(temp_series), description=html, coords=[(lon, lat)])
                
                # Styling - Fix for colors
                pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
                pnt.style.iconstyle.color = get_muted_color(sheet_name, current_headers)
                pnt.style.labelstyle.scale = 0.7 # Makes the text on map slightly smaller/cleaner
            elif not row.dropna().empty:
                failed_rows.append({current_headers[idx]: v for idx, v in enumerate(row.values) if idx < len(current_headers)})

    output_excel = io.BytesIO()
    if failed_rows:
        pd.DataFrame(failed_rows).to_excel(output_excel, index=False)
    
    return kml.kml(), output_excel.getvalue(), stats, len(failed_rows)

# --- UI ---
st.set_page_config(page_title="AKL Mapper v12", layout="wide")
st.title("🌍 Auckland Council Universal Mapper")
file = st.file_uploader("Upload Combined Excel", type="xlsx")
if file:
    with st.spinner("Rendering Map with Muted Color Palette..."):
        kml_data, excel_data, stats, fail_count = process_project(file)
        c1, c2, c3 = st.columns(3)
        c1.metric("Math Points", stats['math'])
        c2.metric("Address Points", stats['address'])
        c3.metric("Review Required", fail_count)
        
        st.download_button("📥 Download Mapped KML", kml_data, file_name="Professional_Map.kml", use_container_width=True)
        if fail_count > 0:
            st.download_button("⚠️ Download Failed Rows", excel_data, file_name="Review.xlsx", use_container_width=True)