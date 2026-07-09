import streamlit as st
import pandas as pd
import datetime
import urllib.parse
import googlemaps
from streamlit_searchbox import st_searchbox
import openpyxl
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="Mileage Tracker Tool", layout="wide")

# --- CONSTANTS ---
MILEAGE_COLUMNS = [
    "Date", "Starting Location", "Destination", "Round Trip",
    "Purpose of Travel", "Odometer Start", "Odometer End", "Calculated Mileage", "Program Code"
]
IMPORT_MARKER = "Imported from template"
METERS_TO_MILES = 0.000621371

# --- API KEY & CLIENT INITIALIZATION ---
api_key = st.secrets.get("GOOGLE_MAPS_API_KEY", "")

@st.cache_resource
def get_gmaps_client(key):
    """Initialize Google Maps client with error handling."""
    if key and key != "YOUR_API_KEY":
        try:
            return googlemaps.Client(key=key)
        except Exception:
            return None
    return None

gmaps = get_gmaps_client(api_key)

# --- GOOGLE PLACES AUTOCOMPLETE ---
def search_google_places(search_term: str):
    """Fetch address predictions from Google Places API."""
    if not gmaps or not search_term:
        return []
    try:
        predictions = gmaps.places_autocomplete(search_term)
        return [p['description'] for p in predictions]
    except Exception:
        return []

# --- GOOGLE DISTANCE MATRIX API ---
def get_google_distance_miles(origin, destination):
    """Get driving distance in miles from Google Distance Matrix API."""
    if not gmaps:
        return 0.0
    try:
        result = gmaps.distance_matrix(origin, destination, mode="driving")
        element = result['rows'][0]['elements'][0]
        
        if element['status'] == 'OK':
            meters = element['distance']['value']
            return round(meters * METERS_TO_MILES, 1)
        else:
            st.error(f"Google couldn't calculate a route: {element['status']}")
            return 0.0
    except Exception as e:
        st.error(f"Error calling Distance Matrix API: {e}")
        return 0.0

# --- SESSION STATE INITIALIZATION ---
def init_session_state():
    """Initialize all required session state variables."""
    defaults = {
        "mileage_data": pd.DataFrame(columns=MILEAGE_COLUMNS),
        "employee_name": "",
        "date_range_str": "",
        "uploaded_file_bytes": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# --- UI: HEADER ---
st.title("🚗 Company Mileage Tracker")
st.markdown("---")

# --- UI: FILE UPLOAD & INGESTION ---
st.header("📂 Excel Template")
uploaded_template = st.file_uploader("Upload your mileage workbook (.xlsx)", type=["xlsx"])

if uploaded_template and st.session_state.uploaded_file_bytes is None:
    st.session_state.uploaded_file_bytes = uploaded_template.getvalue()
    
    try:
        wb = openpyxl.load_workbook(BytesIO(st.session_state.uploaded_file_bytes), data_only=True)
        
        # Extract Sheet 1 metadata
        sheet1 = wb.worksheets[0]
        st.session_state.employee_name = str(sheet1["D11"].value or "").strip()
        st.session_state.date_range_str = str(sheet1["D15"].value or "").strip()
        
        # Extract Sheet 3 journey entries
        if len(wb.worksheets) >= 3:
            sheet3 = wb.worksheets[2]
            existing_rows = []
            row_idx = 5
            
            while sheet3[f"B{row_idx}"].value:
                raw_date = sheet3[f"B{row_idx}"].value
                if isinstance(raw_date, (datetime.date, datetime.datetime)):
                    formatted_date = raw_date.strftime("%Y-%m-%d")
                else:
                    formatted_date = str(raw_date)[:10]
                
                existing_rows.append({
                    "Date": formatted_date,
                    "Starting Location": IMPORT_MARKER,
                    "Destination": str(sheet3[f"C{row_idx}"].value or ""),
                    "Round Trip": "Yes",
                    "Purpose of Travel": str(sheet3[f"E{row_idx}"].value or ""),
                    "Odometer Start": 0,
                    "Odometer End": 0,
                    "Calculated Mileage": float(sheet3[f"F{row_idx}"].value or 0.0),
                    "Program Code": str(sheet3[f"D{row_idx}"].value or "")
                })
                row_idx += 1
            
            if existing_rows:
                st.session_state.mileage_data = pd.DataFrame(existing_rows)
                st.toast(f"✅ Imported {len(existing_rows)} journey records", icon="📥")
        
        st.rerun()
    except Exception as e:
        st.error(f"Failed to parse workbook: {e}")

st.markdown("---")

# --- UI: COVER SHEET INFO ---
st.header("📋 Cover Sheet Information")
col1, col2 = st.columns(2)

with col1:
    employee_name = st.text_input(
        "Employee Name",
        value=st.session_state.employee_name,
        placeholder="John Doe",
        key="cs_employee_name"
    )
    st.session_state.employee_name = employee_name

with col2:
    date_range = st.text_input(
        "Time Period / Date Range",
        value=st.session_state.date_range_str,
        placeholder="e.g., July 1 - July 31, 2026",
        key="cs_date_range"
    )
    st.session_state.date_range_str = date_range

st.markdown("---")

# --- UI: ADD NEW JOURNEY FORM ---
st.header("📍 Add New Journey")

today = datetime.date.today()
col_date, col_start, col_dest = st.columns(3)

with col_date:
    travel_date = st.date_input("Date", value=today, key="journey_travel_date")

with col_start:
    st.write("**Starting Location**")
    start_loc = st_searchbox(
        search_google_places,
        key="start_location_search",
        placeholder="Type starting address..."
    )

with col_dest:
    st.write("**Destination**")
    dest_loc = st_searchbox(
        search_google_places,
        key="destination_search",
        placeholder="Type destination address..."
    )

col_purpose, col_prog_code, col_rt = st.columns([2, 1, 1])

with col_purpose:
    purpose = st.text_input("Purpose of Travel", key="journey_purpose")

with col_prog_code:
    program_code = st.text_input(
        "Program Code",
        placeholder="e.g., PROG-101",
        key="journey_prog_code"
    )

with col_rt:
    round_trip = st.selectbox("Round Trip?", ["No", "Yes"], key="journey_round_trip")

st.markdown("##### 🚗 Odometer Sync Settings")
col_odo_start, col_odo_end = st.columns(2)

with col_odo_start:
    odo_start_input = st.text_input(
        "Odometer Start",
        placeholder="e.g., 45100",
        key="journey_odo_start"
    )

with col_odo_end:
    odo_end_input = st.text_input(
        "Odometer End",
        placeholder="e.g., 45125",
        key="journey_odo_end"
    )

submit_button = st.button("Calculate & Add Entry", type="primary", key="journey_submit_btn", use_container_width=True)

if submit_button:
    if not start_loc or not dest_loc:
        st.error("⚠️ Please provide both a Starting Location and Destination.")
    else:
        google_miles = get_google_distance_miles(start_loc, dest_loc)
        calculated_miles = google_miles * 2 if round_trip == "Yes" else google_miles
        
        o_start = int(odo_start_input) if odo_start_input.strip().isdigit() else None
        o_end = int(odo_end_input) if odo_end_input.strip().isdigit() else None
        
        # Auto-fill odometer fields based on calculated distance
        if o_start is not None and o_end is None:
            o_end = int(o_start + calculated_miles)
        elif o_end is not None and o_start is None:
            o_start = int(o_end - calculated_miles)
        elif o_start is not None and o_end is not None:
            calculated_miles = o_end - o_start
        else:
            o_start, o_end = 0, int(calculated_miles)

        new_entry = {
            "Date": travel_date.strftime("%Y-%m-%d"),
            "Starting Location": start_loc,
            "Destination": f"{dest_loc} (RT)" if round_trip == "Yes" else dest_loc,
            "Round Trip": round_trip,
            "Purpose of Travel": purpose,
            "Odometer Start": o_start,
            "Odometer End": o_end,
            "Calculated Mileage": calculated_miles,
            "Program Code": program_code
        }
        st.session_state.mileage_data = pd.concat(
            [st.session_state.mileage_data, pd.DataFrame([new_entry])],
            ignore_index=True
        )
        st.success(f"✅ Added! Distance: {google_miles} miles")
        st.rerun()

st.markdown("---")

# --- UI: MILEAGE LOG TABLE ---
st.header("📊 Mileage Log")

if not st.session_state.mileage_data.empty:
    total_miles = st.session_state.mileage_data["Calculated Mileage"].sum()
    col_metric, col_spacer = st.columns([1, 3])
    with col_metric:
        st.metric("Total Period Mileage", f"{total_miles} miles")
    
    edited_df = st.data_editor(
        st.session_state.mileage_data,
        num_rows="dynamic",
        use_container_width=True,
        key="mileage_editor"
    )
    st.session_state.mileage_data = edited_df
else:
    st.info("📝 No mileage entries added yet. Use the form above to get started.")

st.markdown("---")

# --- UI: EXPORT TO EXCEL ---
if st.session_state.uploaded_file_bytes:
    st.subheader("💾 Export Back to Excel Template")
    st.markdown("Updates `Sheet 1` and `Sheet 3` while preserving formatting.")
    
    if st.button("Generate Updated Excel Document", type="secondary", use_container_width=True):
        try:
            output_wb = openpyxl.load_workbook(BytesIO(st.session_state.uploaded_file_bytes))
            
            # Update Sheet 1
            s1 = output_wb.worksheets[0]
            s1["D11"] = st.session_state.employee_name
            s1["D15"] = st.session_state.date_range_str

            # Update Sheet 3
            if len(output_wb.worksheets) >= 3:
                s3 = output_wb.worksheets[2]
                current_write_row = 5
                
                # Find next empty row
                while s3[f"B{current_write_row}"].value:
                    current_write_row += 1
                
                # Write only new entries (exclude imported ones)
                new_session_rows = st.session_state.mileage_data[
                    st.session_state.mileage_data["Starting Location"] != IMPORT_MARKER
                ]
                
                for _, row in new_session_rows.iterrows():
                    s3[f"B{current_write_row}"] = row["Date"]
                    s3[f"C{current_write_row}"] = row["Destination"]
                    s3[f"D{current_write_row}"] = row.get("Program Code", "")
                    s3[f"E{current_write_row}"] = row["Purpose of Travel"]
                    s3[f"F{current_write_row}"] = row["Calculated Mileage"]
                    current_write_row += 1
            
            # Generate download
            excel_stream = BytesIO()
            output_wb.save(excel_stream)
            excel_stream.seek(0)
            
            st.download_button(
                label="📥 Download Updated Excel File",
                data=excel_stream,
                file_name=f"Mileage_Report_{st.session_state.employee_name.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Failed to generate Excel file: {e}")

st.markdown("---")

# --- UI: ROUTE MAP & PRINT VIEW ---
st.header("🗺️ Route Map & Print View")

if not st.session_state.mileage_data.empty:
    new_app_entries = st.session_state.mileage_data[
        st.session_state.mileage_data["Starting Location"] != IMPORT_MARKER
    ]

    if not new_app_entries.empty:
        # Get the most recent entry
        last_entry = new_app_entries.iloc[-1]
        
        origin = last_entry["Starting Location"]
        destination = last_entry["Destination"]
        clean_dest = destination.replace(" (RT)", "") if isinstance(destination, str) else str(destination)
        
        trip_miles = last_entry["Calculated Mileage"]
        is_round_trip = last_entry["Round Trip"] == "Yes"
        entry_date = last_entry["Date"]
        entry_purpose = last_entry["Purpose of Travel"]
        
        # Format display labels
        if is_round_trip:
            route_label = f"{origin} ➡️ {clean_dest} 🔄 {origin}"
            mileage_label = f"{trip_miles} miles (Round Trip)"
        else:
            route_label = f"{origin} ➡️ {clean_dest}"
            mileage_label = f"{trip_miles} miles (One Way)"
        
        st.subheader("Current Route Detail")
        col_m1, col_m2 = st.columns(2)
        col_m1.metric("Route", route_label)
        col_m2.metric("Distance", mileage_label)
        
        # Encode URLs
        encoded_origin = urllib.parse.quote_plus(origin)
        encoded_destination = urllib.parse.quote_plus(clean_dest)
        
        # Build Google Maps URL
        if is_round_trip:
            direct_maps_url = f"https://www.google.com/maps/dir/{encoded_origin}/{encoded_destination}/{encoded_origin}/"
        else:
            direct_maps_url = f"https://www.google.com/maps/dir/{encoded_origin}/{encoded_destination}/"
        
        # Action buttons
        st.markdown("##### Actions")
        col_print, col_copy = st.columns(2)
        
        with col_print:
            st.link_button(
                "🖨️ Open in Google Maps",
                direct_maps_url,
                type="primary",
                use_container_width=True
            )
            st.caption("💡 Click **Print** icon (top-right) or press **Ctrl+P** / **Cmd+P**")

        with col_copy:
            text_to_copy = f"Date: {entry_date} | Purpose: {entry_purpose}"
            st.code(text_to_copy, language="text")
            st.caption("📋 Click the copy icon to save trip details")
        
        st.markdown("---")
        
        # Render embedded map if API is valid
        if gmaps:
            if is_round_trip:
                map_url = (
                    f"https://www.google.com/maps/embed/v1/directions"
                    f"?key={api_key}"
                    f"&origin={encoded_origin}"
                    f"&destination={encoded_origin}"
                    f"&waypoints={encoded_destination}"
                )
            else:
                map_url = (
                    f"https://www.google.com/maps/embed/v1/directions"
                    f"?key={api_key}"
                    f"&origin={encoded_origin}"
                    f"&destination={encoded_destination}"
                )
            
            st.components.v1.iframe(map_url, width=900, height=500)
        else:
            st.warning("🔑 Map preview unavailable. Please add a valid Google Maps API Key.")
    else:
        st.info("📅 Sheet template journeys are cached. Add a new manual entry above to view the map.")
else:
    st.write("Add an entry above to generate a live map route.")
