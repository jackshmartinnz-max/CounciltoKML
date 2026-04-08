import streamlit as st
import pandas as pd
import pyproj
import simplekml

# --- CONFIGURATION ---
COLOR_MAP = {
    "all bores": "ffff0000", "bores": "ffff0000",
    "all consents": "ff00ffff", "consents": "ff00ffff",
    "all incidents": "ff0000ff", "incidents": "ff0000ff",
    "hail": "ff800080"
}
FALLBACK_PALETTE = ["ffffff00", "ff00a5ff", "ff008000", "ffffc0cb", "ff2222a2", "ff808080"]

nztm = pyproj.Transformer.from_crs("epsg:2193", "epsg:4326", always_xy=True)
nzmg = pyproj.Transformer.from_crs("epsg:27200", "epsg:4326", always_xy=True)

def is_near_nz(lon, lat):
    return (160 < lon < 185) and (-50 < lat < -30)

def get_smart_title(row, sheet_name):
    """Refined title logic based on your specific Excel columns."""
    s = sheet_name.lower()
    if "hail" in s and 'PropertyAddress' in row:
        return str(row['PropertyAddress'])
    if "bore" in s and 'BORE_ID' in row:
        return f"Bore {row['BORE_ID']}"
    if "incident" in s and 'INCIDENTNUMBER' in row:
        return f"Incident {row['INCIDENTNUMBER']}"
    if "consent" in s and 'CONSENT_NUMBER' in row:
        return f"Consent {row['CONSENT_NUMBER']}"
    
    for col in ['PropertyAddress', 'Location', 'ID', 'Reference']:
        if col in row and pd.notna(row[col]): return str(row[col])
    return "Feature"

def process_excel_to_kml(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    kml = simplekml.Kml()
    color_index = 0

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        lookup = sheet_name.strip().lower()
        folder = kml.newfolder(name=sheet_name)
        
        target_color = COLOR_MAP.get(lookup, FALLBACK_PALETTE[color_index % len(FALLBACK_PALETTE)])
        if lookup not in COLOR_MAP: color_index += 1

        # Column Detection
        e_col = next((c for c in df.columns if c.lower() in ['easting', 'nztmxcoord', 'x']), None)
        n_col = next((c for c in df.columns if c.lower() in ['northing', 'nztmycoord', 'y']), None)
        addr_col = 'PropertyAddress' if 'PropertyAddress' in df.columns else \
                   next((c for c in df.columns if 'address' in c.lower() or 'location' in c.lower()), None)

        for _, row in df.iterrows():
            # Build Table
            border = "0" if "hail" in lookup else "1"
            html = f'<table border="{border}" style="font-family:sans-serif; font-size:12px; border-collapse:collapse; width:300px;">'
            for col, val in row.items():
                if pd.notna(val) and str(val).lower() != 'n/a':
                    html += f'<tr><td style="padding:4px; background:#eee; font-weight:bold;">{col}</td><td>{val}</td></tr>'
            html += "</table>"

            # Create Point
            pnt = folder.newpoint(name=get_smart_title(row, sheet_name), description=html)
            pnt.style.iconstyle.color = target_color
            pnt.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/wht-blank.png"

            # --- THE "AFRICA-PROOF" LOGIC ---
            lon, lat = None, None
            is_hail_sheet = "hail" in lookup

            # Only try coordinates if it's NOT a HAIL sheet (HAIL data uses addresses)
            if not is_hail_sheet:
                try:
                    e_val = float(str(row[e_col]).replace(',', '')) if e_col and pd.notna(row[e_col]) else None
                    n_val = float(str(row[n_col]).replace(',', '')) if n_col and pd.notna(row[n_col]) else None
                    if e_val and n_val:
                        lon, lat = nztm.transform(e_val, n_val)
                        if not is_near_nz(lon, lat):
                            lon, lat = nzmg.transform(e_val, n_val)
                except: pass

            # Placement
            if lon and lat and is_near_nz(lon, lat):
                pnt.coords = [(lon, lat)]
            elif addr_col and pd.notna(row[addr_col]):
                # If we are here, we are using the address (HAIL sites end up here)
                clean_addr = str(row[addr_col]).strip()
                if "auckland" not in clean_addr.lower():
                    clean_addr += ", Auckland, New Zealand"
                pnt.address = clean_addr
                # CRITICAL: If no coords, we must explicitly delete the 0,0 geometry
                pnt.geometry = None 
            else:
                folder.features.remove(pnt)

    return kml.kml()

# --- UI ---
st.title("🌍 Auckland Council KML Pro")
file = st.file_uploader("Upload Excel", type="xlsx")
if file:
    kml_str = process_excel_to_kml(file)
    st.download_button("📥 Download KML", kml_str, file_name="Henderson_Map.kml")