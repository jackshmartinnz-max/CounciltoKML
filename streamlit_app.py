import streamlit as st
import pandas as pd
import pyproj
import simplekml
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import re
import io

# --- CONFIGURATION ---
geolocator = Nominatim(user_agent="Auckland_Universal_Mapper_v5", timeout=10)
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
    return "Map Point"

def get_sheet_color(sheet_name):
    """
    Assigns colors based on keywords in the sheet name.
    KML Colors are AABBGGRR (Alpha, Blue, Green, Red).
    """
    s = sheet_name.lower()
    if "bore" in s:
        return "ffff0000"  # Blue
    if "consent" in s:
        return "ff00ffff"  # Yellow
    if "incident" in s or "pollution" in s:
        return "ff0000ff"  # Red
    if "hail" in s or "land use" in s:
        return "ff800080"  # Purple
    if "well" in s:
        return "ff00ff00"  # Green
    return "ff00a5ff"      # Orange (Default)

def process_project(file):
    xl = pd.ExcelFile(file)
    kml = simplekml.Kml()
    failed_rows = []
    stats = {"math": 0, "address": 0}
    
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        folder = kml.newfolder(name=sheet_name)
        
        # Get color based on sheet name keyword
        target_color = get_sheet_color(sheet_name)
        
        # Universal Column Detection
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'nztmxcoord', 'xcoord', 'x', 'east']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'nztmycoord', 'ycoord', 'y', 'north']), None)
        addr_col = next((c for c in df.columns if 'address' in c.lower() or 'location' in c.lower()), None)

        for _, row in df.iterrows():
            lon, lat = None, None
            found_by = None
            
            # 1. Try Coordinate Math (X/Y or Easting/Northing)
            try:
                if e_col and n_col and pd.notna(row[e_col]) and pd.notna(row[n_col]):
                    e_val = float(str(row[e_col]).replace(',', ''))
                    n_val = float(str(row[n_col]).replace(',', ''))
                    lon, lat = nztm.transform(e_val, n_val)
                    if is_near_nz(lon, lat):
                        found_by = "math"
            except: pass

            # 2. Try Geocoding
            if found_by != "math" and addr_col and pd.notna(row[addr_col]):
                clean_addr = re.sub(r'\b\d{4}\b', '', str(row[addr_col]))
                query = f"{clean_addr}, Auckland, New Zealand"
                try:
                    location = geocode_service(query)
                    if location:
                        lon, lat = location.longitude, location.latitude
                        if is_near_nz(lon, lat):
                            found_by = "address"
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
                pnt.style.iconstyle.scale = 0.8
            else:
                row_dict = row.to_dict()
                row_dict['Original_Sheet'] = sheet_name
                row_dict['Reason_Failed'] = "Invalid X/Y and Address not found"
                failed_rows.append(row_dict)

    failed_df = pd.DataFrame(failed_rows)
    output_excel = io.BytesIO()
    if not failed_df.empty:
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            failed_df.to_excel(writer, index=False)
    
    return kml.kml(), output_excel.getvalue(), stats, len(failed_rows)

# --- UI ---
st.set_page_config(page_title="AKL Council Mapper", layout="centered")
st.title("🌍 Universal Auckland Council KML Tool")

file = st.file_uploader("Upload Council Export (Excel)", type="xlsx")

if file:
    with st.spinner("Processing... Mapping via X/Y and Addresses."):
        kml_data, excel_data, stats, fail_count = process_project(file)
        
        st.success(f"Processing Complete!")
        st.write(f"✅ **{stats['math']}** located via X/Y Coordinates")
        st.write(f"📍 **{stats['address']}** located via Street Address")
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📥 Download KML Map", kml_data, file_name="Auckland_Map.kml", use_container_width=True)
            
        with col2:
            if fail_count > 0:
                st.download_button("⚠️ Download Failed Sites", excel_data, file_name="Review_Required.xlsx", use_container_width=True)