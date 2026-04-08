import streamlit as st
import pandas as pd
import pyproj
import simplekml
import geopandas as gpd
from shapely.geometry import Point
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import re
import io
import os

# --- 1. CONFIGURATION & GIS SETUP ---
# User agent for Nominatim geocoding
geolocator = Nominatim(user_agent="Auckland_Spatial_Suite_v14", timeout=10)
geocode_service = RateLimiter(geolocator.geocode, min_delay_seconds=0.8)

# NZTM to WGS84 Transformer
nztm_transformer = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    """Safety check to ensure points fall within the NZ region."""
    if lon is None or lat is None: return False
    return (160 < lon < 185) and (-48 < lat < -32)

def is_nztm_val(val):
    """Identifies if a number looks like an Auckland-region NZTM coordinate."""
    try:
        f = float(str(val).replace(',', '').strip())
        # Easting range ~1.5M-1.9M | Northing range ~5.8M-6.1M
        return (1500000 < f < 2000000) or (5800000 < f < 6100000)
    except: return False

def get_muted_color(sheet_name, headers):
    """Returns 'Dull' professional KML colors (Format: AABBGGRR)."""
    combined = (str(sheet_name) + " " + " ".join([str(h) for h in headers])).lower()
    if any(x in combined for x in ["bore", "well"]): return "ffb08446"      # Muted Steel Blue
    if "consent" in combined: return "ffa0a000"                           # Muted Teal
    if any(x in combined for x in ["incident", "pollution"]): return "ff4b4bc8" # Terracotta Red
    if any(x in combined for x in ["hail", "contaminated"]): return "ff966496" # Dusty Lavender
    return "ff5a96f0" # Muted Grey-Orange Default

def get_smart_title(row):
    """Tries to find a useful ID to name the map pin."""
    id_cols = ['BORE_ID', 'CONSENT_NUMBER', 'INCIDENTNUMBER', 'SAPSiteID', 'ID', 'ConsentReference']
    for col in row.index:
        if any(x.lower() == str(col).strip().lower() for x in id_cols):
            val = str(row[col]).strip()
            if val and val.lower() != "nan": return val
    return "Asset/Point"

# --- 2. PROCESSING ENGINE ---
def process_excel_to_spatial(file):
    xl = pd.ExcelFile(file)
    kml = simplekml.Kml()
    all_points_data = [] # Store for GeoPackage
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
            
            # Header Detection
            if any(key in row_lower for key in keywords):
                current_headers = [v if (pd.notna(v) and str(v) != "nan" and str(v).strip() != "") else f"Col_{idx}" for idx, v in enumerate(row_vals)]
                current_cols = {v.lower(): idx for idx, v in enumerate(current_headers)}
                continue
            
            if current_headers is None: continue

            lon, lat, found_by = None, None, None
            e_val, n_val = None, None
            
            # A. COORDINATE MATH (Header-based)
            e_idx = next((current_cols[k] for k in ['easting', 'xcoord', 'nztmxcoord', 'x'] if k in current_cols), None)
            n_idx = next((current_cols[k] for k in ['northing', 'ycoord', 'nztmycoord', 'y'] if k in current_cols), None)
            
            try:
                if e_idx is not None and n_idx is not None:
                    raw_e = float(str(row[e_idx]).replace(',', '').strip())
                    raw_n = float(str(row[n_idx]).replace(',', '').strip())
                    if raw_e > 3000000: raw_e, raw_n = raw_n, raw_e # Flip if swapped
                    lon, lat = nztm_transformer.transform(raw_e, raw_n)
                    if is_near_nz(lon, lat): 
                        found_by, e_val, n_val = "math", raw_e, raw_n
            except: pass

            # B. BRUTE FORCE RECOVERY (If headers failed)
            if not found_by:
                nums = [float(str(v).replace(',', '')) for v in row.values if is_nztm_val(v)]
                if len(nums) >= 2:
                    e_val, n_val = min(nums), max(nums)
                    lon, lat = nztm_transformer.transform(e_val, n_val)
                    if is_near_nz(lon, lat): found_by = "math"

            # C. ADDRESS GEOCODING (Fallback)
            if not found_by:
                addr_val = next((str(v) for v in row.values if pd.notna(v) and len(str(v)) > 10 and any(w in str(v).upper() for w in [' ROAD', ' STREET', ' AVE', ' RD', ' ST'])), None)
                if addr_val:
                    # Fix for pre-3.12 Python: regex outside f-string
                    clean_addr = re.sub(r'\b\d{4}\b', '', addr_val)
                    try:
                        loc = geocode_service(f"{clean_addr}, Auckland, NZ")
                        if loc and is_near_nz(loc.longitude, loc.latitude):
                            lon, lat = loc.longitude, loc.latitude
                            found_by = "address"
                    except: pass

            # D. DATA PACKAGING
            if found_by:
                stats[found_by] += 1
                row_dict = {current_headers[i]: row.values[i] for i in range(min(len(current_headers), len(row)))}
                row_dict['Sheet_Source'] = sheet_name
                
                # Setup KML Point
                pnt_name = get_smart_title(pd.Series(row_dict))
                html = '<table border="1" style="font-size:11px; border-collapse:collapse; width:300px;">'
                for k, v in row_dict.items():
                    if pd.notna(v) and "Col_" not in k:
                        html += f'<tr><td style="background:#eee; font-weight:bold; padding:2px;">{k}</td><td style="padding:2px;">{v}</td></tr>'
                html += "</table>"
                
                pnt = folder.newpoint(name=pnt_name, description=html, coords=[(lon, lat)])
                pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
                pnt.style.iconstyle.color = get_muted_color(sheet_name, current_headers)
                pnt.style.labelstyle.scale = 0.7
                
                # Setup GeoPackage Logic (NZTM is better for ArcGIS Pro)
                geom_pt = Point(e_val, n_val) if e_val else Point(lon, lat)
                row_dict['geometry'] = geom_pt
                all_points_data.append(row_dict)
            elif not row.dropna().empty:
                failed_rows.append({current_headers[i]: row.values[i] for i in range(min(len(current_headers), len(row)))})

    # Prepare Outputs
    gpkg_data = None
    if all_points_data:
        gdf = gpd.GeoDataFrame(all_points_data, crs="EPSG:2193" if e_val else "EPSG:4326")
        for col in gdf.columns:
            if gdf[col].dtype == 'object': gdf[col] = gdf[col].astype(str)
        gdf.to_file("output.gpkg", driver="GPKG")
        with open("output.gpkg", "rb") as f:
            gpkg_data = f.read()
        os.remove("output.gpkg")

    err_excel = io.BytesIO()
    if failed_rows:
        pd.DataFrame(failed_rows).to_excel(err_excel, index=False)

    return kml.kml(), gpkg_data, err_excel.getvalue(), stats, len(failed_rows)

# --- 3. UI LAYOUT ---
st.set_page_config(page_title="Auckland Spatial Suite", layout="wide")
st.title("🌍 Auckland Council Universal Spatial Suite")
st.markdown("Convert Council Excel extracts into **Google Earth (KML)** and **ArcGIS Pro (GeoPackage)** files.")

uploaded_file = st.file_uploader("Upload your Combined Council Excel (.xlsx)", type="xlsx")

if uploaded_file:
    with st.spinner("Processing spatial data and recovering coordinates..."):
        kml_str, gpkg_bytes, err_bytes, stats, fail_count = process_excel_to_spatial(uploaded_file)
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Mapped via Coordinates", stats['math'])
        m2.metric("Mapped via Address", stats['address'])
        m3.metric("Records Requiring Review", fail_count)

        st.divider()
        
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Google Earth Pro")
            st.download_button("📌 Download KML", kml_str, file_name="Council_Map.kml", use_container_width=True)
        with c2:
            st.subheader("ArcGIS Pro / QGIS")
            if gpkg_bytes:
                st.download_button("💾 Download GeoPackage (.gpkg)", gpkg_bytes, file_name="Council_Data.gpkg", use_container_width=True)
        
        if fail_count > 0:
            st.warning(f"{fail_count} rows could not be mapped. Download the report below to investigate.")
            st.download_button("⚠️ Download Review Report", err_bytes, file_name="Review_Needed.xlsx")