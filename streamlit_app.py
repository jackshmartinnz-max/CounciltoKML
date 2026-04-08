import streamlit as st
import pandas as pd
import pyproj
import simplekml

# --- COORDINATE SYSTEM REGISTRY ---
# NZTM2000 (New/Standard)
nztm_transformer = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)
# NZMG (Old - used in legacy Council data)
nzmg_transformer = pyproj.Transformer.from_crs("epsg:27200", "epsg:4326", always_xy=True)

def is_near_auckland(lon, lat):
    """Safety check to ensure points aren't in Africa or the ocean."""
    return (170 < lon < 179) and (-39 < lat < -34)

def process_excel_to_kml(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    kml = simplekml.Kml()
    
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        folder = kml.newfolder(name=sheet_name)

        # Robust Column Search
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'eastings', 'x', 'east', 'nzte', 'mge']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'northings', 'y', 'north', 'nztn', 'mgn']), None)
        addr_col = next((c for c in df.columns if 'address' in c.lower()), None)

        for _, row in df.iterrows():
            # Create Table (HAIL logic included)
            border = "0" if "hail" in sheet_name.lower() else "1"
            html = f'<table border="{border}" style="font-family:sans-serif; font-size:12px; border-collapse:collapse; width:300px;">'
            for col, val in row.items():
                if pd.notna(val):
                    html += f'<tr><td style="padding:4px; background:#eee; font-weight:bold;">{col}</td><td>{val}</td></tr>'
            html += '</table>'

            pnt = folder.newpoint(description=html)
            pnt.name = str(row.get('ID', row.get('Reference', 'Point')))
            pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
            pnt.style.iconstyle.scale = 0.7

            # --- THE CONVERSION LOGIC ---
            lon, lat = None, None
            try:
                e_val = float(str(row[e_col]).replace(',', '')) if e_col else None
                n_val = float(str(row[n_col]).replace(',', '')) if n_col else None

                if e_val and n_val:
                    # Try NZTM First (Most likely)
                    lon, lat = nztm_transformer.transform(e_val, n_val)
                    
                    # If it lands in Africa, try the Old NZMG system
                    if not is_near_auckland(lon, lat):
                        lon, lat = nzmg_transformer.transform(e_val, n_val)
            except:
                pass

            # Apply Location
            if lon and lat and is_near_auckland(lon, lat):
                pnt.coords = [(lon, lat)]
            elif addr_col and pd.notna(row[addr_col]):
                pnt.address = f"{row[addr_col]}, Auckland, New Zealand"
            else:
                # If everything fails, delete this empty point to avoid 0,0 Africa errors
                folder.features.remove(pnt)

    return kml.kml()

# --- STREAMLIT UI ---
st.set_page_config(page_title="Auckland KML Pro", page_icon="🌍")
st.title("🌍 Auckland Council KML Generator")
st.write("Specialized for NZTM/NZMG Coordinate Conversions")

file = st.file_uploader("Upload Henderson/Auckland Excel", type="xlsx")

if file:
    with st.spinner("Converting..."):
        kml_str = process_excel_to_kml(file)
        st.success("Conversion Complete!")
        st.download_button("📥 Download KML", kml_str, file_name="Henderson_Project.kml")