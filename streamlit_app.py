import streamlit as st
import pandas as pd
import pyproj
import simplekml

# --- CONFIGURATION & COLOURS ---
COLOR_MAP = {
    "all bores": "ffff0000",       # Blue (aabbggrr)
    "bores": "ffff0000",
    "all consents": "ff00ffff",     # Yellow
    "consents": "ff00ffff",
    "all incidents": "ff0000ff",    # Red
    "incidents": "ff0000ff",
    "hail": "ff800080",            # Purple
}
FALLBACK_PALETTE = ["ffffff00", "ff00a5ff", "ff008000", "ffffc0cb", "ff2222a2", "ff808080"]

# Projection setup
nztm_transformer = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)
nzmg_transformer = pyproj.Transformer.from_crs("epsg:27200", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    return (160 < lon < 185) and (-50 < lat < -30)

def get_smart_title(row, sheet_name):
    """Specific logic for Auckland Council column names."""
    # Try specific ID columns first based on the sheet type
    if "bore" in sheet_name.lower():
        if 'BORE_ID' in row and pd.notna(row['BORE_ID']): return f"Bore: {row['BORE_ID']}"
    if "consent" in sheet_name.lower():
        if 'CONSENT_NUMBER' in row and pd.notna(row['CONSENT_NUMBER']): return f"Consent: {row['CONSENT_NUMBER']}"
    if "incident" in sheet_name.lower():
        if 'INCIDENTNUMBER' in row and pd.notna(row['INCIDENTNUMBER']): return f"Incident: {row['INCIDENTNUMBER']}"
    if "hail" in sheet_name.lower():
        if 'PropertyAddress' in row and pd.notna(row['PropertyAddress']): return str(row['PropertyAddress'])
        if 'SAPSiteID' in row and pd.notna(row['SAPSiteID']): return f"HAIL: {row['SAPSiteID']}"

    # General fallback candidates
    for col in ['Reference', 'ID', 'SiteName', 'LOCATION', 'PropertyAddress']:
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
        
        target_color = COLOR_MAP.get(lookup, FALLBACK_PALETTE[color_index % len(FALLBACK_PALETTE)])
        if lookup not in COLOR_MAP:
            color_index += 1

        folder = kml.newfolder(name=sheet_name)

        # Coordinate Column Detection (including your specific 'NZTMXCOORD' columns)
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'eastings', 'x', 'east', 'nztmxcoord']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'northings', 'y', 'north', 'nztmycoord']), None)
        addr_col = next((c for c in df.columns if 'address' in c.lower() or 'location' in c.lower()), None)

        for _, row in df.iterrows():
            # HTML Table
            border = "0" if "hail" in lookup else "1"
            html = f'<table border="{border}" style="font-family:sans-serif; font-size:12px; border-collapse:collapse; width:300px;">'
            for col, val in row.items():
                if pd.notna(val) and str(val).strip().lower() != 'n/a':
                    html += f'<tr><td style="padding:4px; background:#eee; font-weight:bold;">{col}</td><td>{val}</td></tr>'
            html += "</table>"

            # Create the Placemark
            pnt = folder.newpoint(name=get_smart_title(row, sheet_name), description=html)
            
            # Icon styling
            pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"
            pnt.style.iconstyle.color = target_color
            pnt.style.iconstyle.scale = 0.8

            # COORDINATE MATH
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

            # PLACEMENT LOGIC
            if lon and lat and is_near_nz(lon, lat):
                pnt.coords = [(lon, lat)]
            elif addr_col and pd.notna(row[addr_col]) and str(row[addr_col]).lower() != 'n/a':
                # Use Address and REMOVE coordinates to prevent "Africa" issues
                clean_addr = str(row[addr_col]).strip()
                if "auckland" not in clean_addr.lower():
                    clean_addr += ", Auckland, New Zealand"
                pnt.address = clean_addr
                pnt.geometry = None # This is the magic line that prevents Africa!
            else:
                folder.features.remove(pnt)

    return kml.kml()

# --- STREAMLIT UI ---
st.title("🌍 Auckland Council KML Pro")
st.info("Now with Smart Titles and 'Africa-Free' Address Geocoding")
file = st.file_uploader("Upload Henderson Excel File", type="xlsx")

if file:
    with st.spinner("Processing..."):
        kml_str = process_excel_to_kml(file)
        st.success("KML Ready!")
        st.download_button("📥 Download KML", kml_str, file_name="Auckland_Project_Map.kml")