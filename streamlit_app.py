import streamlit as st
import pandas as pd
import pyproj
import simplekml

# --- CONFIGURATION ---
# KML colors are aabbggrr (Alpha, Blue, Green, Red)
COLOR_MAP = {
    "all bores": "ffff0000",       # Blue
    "bores": "ffff0000",           # Blue
    "all consents": "ff00ffff",     # Yellow
    "consents": "ff00ffff",         # Yellow
    "all incidents": "ff0000ff",    # Red
    "incidents": "ff0000ff",        # Red
    "hail": "ff800080",            # Purple
}

# Fallback palette for unknown sheets (Cyan, Orange, Green, Pink, Brown, Grey)
FALLBACK_PALETTE = ["ffffff00", "ff00a5ff", "ff008000", "ffffc0cb", "ff2222a2", "ff808080"]

nztm_transformer = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)
nzmg_transformer = pyproj.Transformer.from_crs("epsg:27200", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    return (160 < lon < 180) and (-48 < lat < -32)

def process_excel_to_kml(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    kml = simplekml.Kml()
    color_index = 0

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        lookup = sheet_name.strip().lower()
        
        # 1. Determine Sheet Color (Step 5.2)
        target_color = COLOR_MAP.get(lookup, FALLBACK_PALETTE[color_index % len(FALLBACK_PALETTE)])
        if lookup not in COLOR_MAP:
            color_index += 1

        folder = kml.newfolder(name=sheet_name)

        # Column Detection
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'eastings', 'x', 'east']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'northings', 'y', 'north']), None)
        addr_col = next((c for c in df.columns if 'address' in c.lower() or 'location' in c.lower()), None)

        for _, row in df.iterrows():
            # 2. Build HTML Table (Step 6)
            is_hail = "hail" in lookup
            border = "0" if is_hail else "1"
            html = f'<table border="{border}" style="font-family:sans-serif; font-size:12px; border-collapse:collapse; width:300px;">'
            for col, val in row.items():
                if pd.notna(val):
                    html += f'<tr><td style="padding:4px; background:#eee; font-weight:bold;">{col}</td><td>{val}</td></tr>'
            html += "</table>"

            # 3. Create Placemark
            pnt = folder.newpoint(description=html)
            pnt.name = str(row.get('ID', row.get('Reference', 'Point')))
            
            # Apply Color and Icon
            pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
            pnt.style.iconstyle.color = target_color
            pnt.style.iconstyle.scale = 0.7

            # 4. Location Logic (Step 4.1 & 7.2)
            lon, lat = None, None
            try:
                e_val = float(str(row[e_col]).replace(',', '')) if e_col else None
                n_val = float(str(row[n_col]).replace(',', '')) if n_col else None
                
                if e_val and n_val:
                    # Try NZTM then NZMG
                    lon, lat = nztm_transformer.transform(e_val, n_val)
                    if not is_near_nz(lon, lat):
                        lon, lat = nzmg_transformer.transform(e_val, n_val)
            except:
                pass

            if lon and lat and is_near_nz(lon, lat):
                pnt.coords = [(lon, lat)]
            elif addr_col and pd.notna(row[addr_col]):
                # IMPORTANT: Clean the address for Google Earth
                clean_address = str(row[addr_col]).strip()
                if "auckland" not in clean_address.lower():
                    clean_address += ", Auckland, New Zealand"
                pnt.address = clean_address
            else:
                folder.features.remove(pnt)

    return kml.kml()

# --- STREAMLIT UI ---
st.set_page_config(page_title="Auckland Council Map Tool")
st.title("🌍 Auckland KML Pro")
st.markdown("Handles Coordinates, Addresses, and Color-Coding.")

file = st.file_uploader("Upload Excel", type="xlsx")

if file:
    with st.spinner("Processing Sheets and Geocoding..."):
        kml_str = process_excel_to_kml(file)
        st.success("KML Ready!")
        st.download_button("📥 Download KML", kml_str, file_name="Auckland_Project_Map.kml")