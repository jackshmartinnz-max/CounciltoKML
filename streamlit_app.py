import streamlit as st
import pandas as pd
import pyproj
import simplekml

# --- CONFIGURATION ---
COLOR_MAP = {
    "all bores": "ffff0000",       # Blue
    "bores": "ffff0000",
    "all consents": "ff00ffff",     # Yellow
    "consents": "ff00ffff",
    "all incidents": "ff0000ff",    # Red
    "incidents": "ff0000ff",
    "hail": "ff800080",            # Purple
}
FALLBACK_PALETTE = ["ffffff00", "ff00a5ff", "ff008000", "ffffc0cb", "ff2222a2", "ff808080"]

nztm_transformer = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)
nzmg_transformer = pyproj.Transformer.from_crs("epsg:27200", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    """Ensures points are within the NZ region."""
    return (160 < lon < 185) and (-50 < lat < -30)

def get_smart_title(row, sheet_name):
    """Selects the most relevant ID or Name for the point title."""
    s = sheet_name.lower()
    # Priority list of column names for titles
    candidates = [
        'BORE_ID', 'CONSENT_NUMBER', 'INCIDENTNUMBER', 'SAPSiteID', 
        'Reference', 'ID', 'SiteName', 'PropertyAddress', 'Location'
    ]
    for col in candidates:
        if col in row and pd.notna(row[col]):
            return str(row[col])
    return "Feature"

def process_excel_to_kml(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    kml = simplekml.Kml()
    color_index = 0

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        lookup = sheet_name.strip().lower()
        
        # Determine Color
        target_color = COLOR_MAP.get(lookup, FALLBACK_PALETTE[color_index % len(FALLBACK_PALETTE)])
        if lookup not in COLOR_MAP:
            color_index += 1

        folder = kml.newfolder(name=sheet_name)

        # Coordinate Column Search
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'eastings', 'x', 'east', 'nztmxcoord']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'northings', 'y', 'north', 'nztmycoord']), None)
        addr_col = next((c for c in df.columns if 'address' in c.lower() or 'location' in c.lower()), None)

        for _, row in df.iterrows():
            # Build Table
            border = "0" if "hail" in lookup else "1"
            html = f'<table border="{border}" style="font-family:sans-serif; font-size:12px; border-collapse:collapse; width:300px;">'
            for col, val in row.items():
                if pd.notna(val):
                    html += f'<tr><td style="padding:4px; background:#eee; font-weight:bold;">{col}</td><td>{val}</td></tr>'
            html += "</table>"

            # CONVERSION LOGIC
            lon, lat = None, None
            try:
                e_val = float(str(row[e_col]).replace(',', '')) if e_col and pd.notna(row[e_col]) else None
                n_val = float(str(row[n_col]).replace(',', '')) if n_col and pd.notna(row[n_col]) else None
                
                if e_val and n_val:
                    lon, lat = nztm_transformer.transform(e_val, n_val)
                    if not is_near_nz(lon, lat):
                        lon, lat = nzmg_transformer.transform(e_val, n_val)
            except:
                pass

            # PLACEMARK CREATION
            # If we have valid NZ coordinates:
            if lon and lat and is_near_nz(lon, lat):
                pnt = folder.newpoint(name=get_smart_title(row, sheet_name), description=html, coords=[(lon, lat)])
            # If no coordinates, try Address:
            elif addr_col and pd.notna(row[addr_col]):
                clean_addr = str(row[addr_col]).strip()
                if "auckland" not in clean_addr.lower():
                    clean_addr += ", Auckland, New Zealand"
                # Use newplacemark (generic) instead of newpoint to avoid Africa (0,0) defaults
                pnt = folder.newplacemark(name=get_smart_title(row, sheet_name), description=html)
                pnt.address = clean_addr
            else:
                continue # Skip rows with no location data

            # Apply Icon and Color
            pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
            pnt.style.iconstyle.color = target_color
            pnt.style.iconstyle.scale = 0.8

    return kml.kml()

# --- STREAMLIT UI ---
st.title("🌍 Auckland Council KML Pro")
file = st.file_uploader("Upload Excel", type="xlsx")

if file:
    with st.spinner("Processing..."):
        kml_str = process_excel_to_kml(file)
        st.success("KML Ready!")
        st.download_button("📥 Download KML", kml_str, file_name="Auckland_Project_Map.kml")