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

# --- CONFIGURATION ---
geolocator = Nominatim(user_agent="Auckland_Universal_Mapper_v13", timeout=10)
geocode_service = RateLimiter(geolocator.geocode, min_delay_seconds=0.8)
nztm_transformer = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    if lon is None or lat is None: return False
    return (160 < lon < 185) and (-48 < lat < -32)

def is_nztm_val(val):
    try:
        f = float(str(val).replace(',', '').strip())
        return (1500000 < f < 2000000) or (5800000 < f < 6100000)
    except: return False

def process_project(file):
    xl = pd.ExcelFile(file)
    kml = simplekml.Kml()
    all_points_data = [] # For GeoPackage
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
                current_headers = [v if v != "" and v != "nan" else f"Col_{idx}" for idx, v in enumerate(row_vals)]
                current_cols = {v.lower(): idx for idx, v in enumerate(current_headers)}
                continue
            
            if current_headers is None: continue

            lon, lat, found_by = None, None, None
            e_val, n_val = None, None
            
            # 1. Coordinate Math (NZTM)
            e_idx = next((current_cols[k] for k in ['easting', 'xcoord', 'nztmxcoord', 'x'] if k in current_cols), None)
            n_idx = next((current_cols[k] for k in ['northing', 'ycoord', 'nztmycoord', 'y'] if k in current_cols), None)
            
            try:
                if e_idx is not None and n_idx is not None:
                    e_val = float(str(row[e_idx]).replace(',', '').strip())
                    n_val = float(str(row[n_idx]).replace(',', '').strip())
                    if e_val > 3000000: e_val, n_val = n_val, e_val
                    lon, lat = nztm_transformer.transform(e_val, n_val)
                    if is_near_nz(lon, lat): found_by = "math"
            except: pass

            # 2. Brute Force Fallback
            if not found_by:
                nums = [float(str(v).replace(',', '')) for v in row.values if is_nztm_val(v)]
                if len(nums) >= 2:
                    e_val, n_val = min(nums), max(nums)
                    lon, lat = nztm_transformer.transform(e_val, n_val)
                    if is_near_nz(lon, lat): found_by = "math"

            # 3. Address Geocoding
            if not found_by:
                addr_val = next((str(v) for v in row.values if pd.notna(v) and len(str(v)) > 10 and any(w in str(v).upper() for w in [' ROAD', ' STREET', ' AVE', ' RD', ' ST'])), None)
                if addr_val:
                    try:
                        loc = geocode_service(f"{re.sub(r'\b\d{4}\b', '', addr_val)}, Auckland, NZ")
                        if loc:
                            lon, lat = loc.longitude, loc.latitude
                            found_by = "address"
                    except: pass

            if found_by:
                stats[found_by] += 1
                row_dict = {current_headers[i]: row.values[i] for i in range(len(current_headers))}
                row_dict['Sheet_Source'] = sheet_name
                row_dict['Found_Via'] = found_by
                
                # For GeoPackage (using NZTM geometry)
                if e_val and n_val:
                    row_dict['geometry'] = Point(e_val, n_val)
                else: # If found by address, reverse transform back to NZTM for ArcGIS consistency
                    # (Simplified for this example, usually stay in 4326)
                    row_dict['geometry'] = Point(lon, lat) 

                all_points_data.append(row_dict)
                
                # Add to KML as usual...
                folder.newpoint(name=str(row.values[0]), coords=[(lon, lat)])

    # Create GeoPackage
    gpkg_buffer = io.BytesIO()
    if all_points_data:
        gdf = gpd.GeoDataFrame(all_points_data)
        # Ensure dates are strings for GPKG compatibility in basic scripts
        for col in gdf.columns:
            if gdf[col].dtype == 'object': gdf[col] = gdf[col].astype(str)
        
        # Save to temporary file because GeoPandas needs a real path
        gdf.to_file("temp.gpkg", driver="GPKG")
        with open("temp.gpkg", "rb") as f:
            gpkg_buffer = f.read()
        os.remove("temp.gpkg")

    return kml.kml(), gpkg_buffer, stats

# --- UI ---
st.set_page_config(page_title="AKL Spatial Suite", layout="wide")
st.title("🌍 Auckland Council Spatial Suite (ArcGIS & Google Earth)")
file = st.file_uploader("Upload Excel", type="xlsx")

if file:
    kml_out, gpkg_out, stats = process_project(file)
    col1, col2 = st.columns(2)
    with col1:
        st.download_button("📌 Download for Google Earth (KML)", kml_out, file_name="Auckland_Map.kml")
    with col2:
        st.download_button("💾 Download for ArcGIS Pro (GeoPackage)", gpkg_out, file_name="Auckland_Data.gpkg")