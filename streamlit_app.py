import streamlit as st
import pandas as pd
import datetime
import googlemaps # pip install googlemaps
from streamlit_searchbox import st_searchbox # pip install streamlit-searchbox
import openpyxl # pip install openpyxl
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="Company Mileage Tracker", layout="wide")

# --- API KEY MANAGEMENT ---
if "api_key" not in st.session_state:
    if "GOOGLE_MAPS_API_KEY" in st.secrets:
        st.session_state.api_key = st.secrets["GOOGLE_MAPS_API_KEY"]
    else:
        st.session_state.api_key = "YOUR_API_KEY"  # Fallback placeholder

# --- API VALIDATION FUNCTION ---
@st.cache_data(show_spinner=False)
def check_api_key_status(api_key):
    """Tests if the Google Maps API key is fundamentally valid."""
    if not api_key or api_key == "YOUR_API_KEY":
        return "Missing", "Please configure your Google Maps API key to enable address autocompletion."
    
    try:
        test_client = googlemaps.Client(key=api_key)
        test_client.places_autocomplete("A")
        return "Valid", "Google Maps API connected successfully! Autocomplete and routing are online."
    except Exception as e:
        error_msg = str(e)
        if "API key not valid" in error_msg:
            return "Invalid Key", "The API key provided is not recognized by Google. Check for typos."
        elif "The provided API key is expired" in error_msg:
            return "Expired Key", "This Google API key has expired."
        else:
            return "API Error", f"Connection failed: {error_msg}"

# --- INITIALIZE GOOGLE MAPS CLIENT ---
@st.cache_resource
def get_gmaps_client(api_key):
    if api_key and api_key != "YOUR_API_KEY":
        try:
            return googlemaps.Client(key=api_key)
        except:
            return None
    return None

gmaps = get_gmaps_client(st.session_state.api_key)

# --- AUTOCOMPLETE SEARCH FUNCTION ---
def search_google_places(search_term: str):
    """Fetches address predictions from Google Places API as the user types."""
    if not gmaps or not search_term:
        return []
    try:
        predictions = gmaps.places_autocomplete(search_term)
        return [p['description'] for p in predictions]
    except Exception:
        return []

# --- GOOGLE DISTANCE MATRIX API FETCH ---
def get_google_distance_miles(origin, destination):
    """Calls Google API to get actual driving distance in miles."""
    if not gmaps:
        return 0.0
    try:
        matrix_result = gmaps.distance_matrix(origin, destination, mode="driving")
        element = matrix_result['rows'][0]['elements'][0]
        
        if element['status'] == 'OK':
            meters = element['distance']['value']
            miles = meters * 0.000621371
            return round(miles, 1)
        else:
            st.error(f"Google couldn't calculate a route: {element['status']}")
            return 0.0
    except Exception as e:
        st.error(f"Error calling Distance Matrix API: {e}")
        return 0.0

# --- INITIALIZE SESSION STATE FOR DATA ---
if "mileage_data" not in st.session_state:
    st.session_state.mileage_data = pd.DataFrame(columns=[
        "Date", "Starting Location", "Destination", "Round Trip", 
        "Purpose of Travel", "Odometer Start", "Odometer End", "Calculated Mileage"
    ])

if "employee_name" not in st.session_state:
    st.session_state.employee_name = ""

if "date_range_str" not in st.session_state:
    st.session_state.date_range_str = ""

if "uploaded_file_bytes" not in st.session_state:
    st.session_state.uploaded_file_bytes = None

# --- HEADER SECTION ---
st.title("🚗 Company Mileage Tracker")

# --- API STATUS INDICATOR BAR (TOP OF PAGE) ---
api_status, api_message = check_api_key_status(st.session_state.api_key)

if api_status == "Valid":
    st.success(f"🟢 **API Status: Connected** | {api_message}")
elif api_status == "Missing":
    st.info(f"🔵 **API Status: Not Configured** | {api_message}")
else:
    st.error(f"🔴 **API Status: Error ({api_status})** | {api_message}")

# --- API DEBUG UTILITY EXPANDER ---
with st.expander("🛠️ API Key Debugger & Config"):
    st.markdown("Use this utility to dynamically test keys or read configuration diagnostic codes.")
    
    user_pasted_key = st.text_input(
        "Test a temporary Google API Key:", 
        value=st.session_state.api_key, 
        type="password",
        help="Paste your string here. It will instantly re-verify the token against Google Cloud endpoints."
    )
    
    if user_pasted_key != st.session_state.api_key:
        st.session_state.api_key = user_pasted_key
        st.cache_data.clear() 
        st.rerun()
        
    st.markdown(f"""
    **Diagnostic Checklist:**
    * Current Token Value: `{st.session_state.api_key[:5]}...{st.session_state.api_key[-5:] if len(st.session_state.api_key) > 5 else ''}`
    * Required Library Status: `googlemaps` integration verified.
    * Target Cloud Credentials Needed: **Places API**, **Maps Embed API**, & **Distance Matrix API**.
    """)

st.markdown("---")

# --- EXCEL TEMPLATE UPLOADER & INGESTION (UPDATED) ---
st.header("📂 Excel Template Synchronization")
uploaded_template = st.file_uploader("Upload your company mileage workbook (.xlsx)", type=["xlsx"])

if uploaded_template is not None and st.session_state.uploaded_file_bytes is None:
    # Save file bytes in state to maintain permanence during updates
    st.session_state.uploaded_file_bytes = uploaded_template.getvalue()
    
    try:
        wb = openpyxl.load_workbook(BytesIO(st.session_state.uploaded_file_bytes), data_only=True)
        
        # 1. Ingest Sheet 1 Metadata
        sheet1 = wb.worksheets[0]
        st.session_state.employee_name = str(sheet1["D11"].value or "")
        st.session_state.date_range_str = str(sheet1["D15"].value or "")
        
        # 2. Ingest Sheet 3 Existing Journey Entries
        if len(wb.worksheets) >= 3:
            sheet3 = wb.worksheets[2]
            existing_rows = []
            
            # CHANGED FROM 4 TO 5: Skip the header row ('Total Miles', etc.) and start at first data row
            row_idx = 5 
            while sheet3[f"B{row_idx}"].value is not None:
                # Handle Excel datetime objects vs raw strings safely
                raw_date = sheet3[f"B{row_idx}"].value
                if isinstance(raw_date, (datetime.date, datetime.datetime)):
                    formatted_date = raw_date.strftime("%Y-%m-%d")
                else:
                    formatted_date = str(raw_date)[:10]
                
                # Append rows mapping exactly to our internal DataFrame schema
                existing_rows.append({
                    "Date": formatted_date,
                    "Starting Location": "Imported from template", 
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
                # Synchronize the main DataFrame with the imported Excel records
                st.session_state.mileage_data = pd.DataFrame(existing_rows)
                st.toast(f"Imported {len(existing_rows)} existing journey records from sheet 3!", icon="📥")
                
        st.rerun()
    except Exception as e:
        st.error(f"Failed parsing workbook configuration parameters: {e}")


st.markdown("---")

# --- SECTION 1: COVER SHEET INFO ---
st.header("📋 Cover Sheet Information")
col1, col2 = st.columns(2)

with col1:
    employee_name = st.text_input("Employee Name (C3 / D11)", value=st.session_state.employee_name, placeholder="John Doe", key="cs_employee_name")
    st.session_state.employee_name = employee_name

with col2:
    date_range = st.text_input("Time Period / Date Range (D15)", value=st.session_state.date_range_str, placeholder="e.g., July 1 - July 31, 2026", key="cs_date_range")
    st.session_state.date_range_str = date_range

st.markdown("---")

# --- SECTION 2: USER INPUT FORM ---
st.header("📍 Add New Journey")

today = datetime.date.today()
col_date, col_start, col_dest = st.columns(3)
with col_date:
    travel_date = st.date_input("Date", value=today, key="journey_travel_date")

with col_start:
    st.write("**Starting Location**")
    if api_status == "Valid":
        start_loc = st_searchbox(search_google_places, key="start_location_search", placeholder="Type starting address...")
    else:
        start_loc = st.text_input("Starting Location (Fallback mode)", placeholder="Type address manually...", key="start_fallback")

with col_dest:
    st.write("**Destination**")
    if api_status == "Valid":
        dest_loc = st_searchbox(search_google_places, key="destination_search", placeholder="Type destination address...")
    else:
        dest_loc = st.text_input("Destination (Fallback mode)", placeholder="Type address manually...", key="dest_fallback")

col_purpose, col_prog_code, col_rt = st.columns([2, 1, 1])
with col_purpose:
    purpose = st.text_input("Purpose of Travel", key="journey_purpose")
with col_prog_code:
    program_code = st.text_input("Program Code (D4)", placeholder="e.g., PROG-101", key="journey_prog_code")
with col_rt:
    round_trip = st.selectbox("Round Trip?", ["No", "Yes"], key="journey_round_trip")

# --- ODOMETER INTERACTION BLOCK ---
st.markdown("##### 🚗 Odometer Sync Settings")
col_odo_start, col_odo_end = st.columns(2)
with col_odo_start:
    odo_start_input = st.text_input("Odometer Start", value="", placeholder="e.g., 45100", key="journey_odo_start")
with col_odo_end:
    odo_end_input = st.text_input("Odometer End", value="", placeholder="e.g., 45125", key="journey_odo_end")

submit_button = st.button("Calculate & Add Entry", type="primary", key="journey_submit_btn")

# --- FORM LOGIC ---
if submit_button:
    if not start_loc or not dest_loc:
        st.error("Error: Please provide both a Starting Location and Destination.")
    else:
        google_miles = get_google_distance_miles(start_loc, dest_loc)
        calculated_miles = google_miles * 2 if round_trip == "Yes" else google_miles
        
        o_start = int(odo_start_input) if odo_start_input.strip().isdigit() else None
        o_end = int(odo_end_input) if odo_end_input.strip().isdigit() else None
        
        if o_start is not None and o_end is None:
            o_end = int(o_start + calculated_miles)
        elif o_end is not None and o_start is None:
            o_start = int(o_end - calculated_miles)
        elif o_start is not None and o_end is not None:
            calculated_miles = o_end - o_start
        else:
            o_start, o_end = 0, int(calculated_miles)

        # Build structural row
        new_entry = {
            "Date": travel_date.strftime("%Y-%m-%d"),
            "Starting Location": start_loc,
            "Destination": f"{dest_loc} (RT)" if round_trip == "Yes" else dest_loc,
            "Round Trip": round_trip,
            "Purpose of Travel": purpose,
            "Odometer Start": o_start,
            "Odometer End": o_end,
            "Calculated Mileage": calculated_miles,
            "Program Code": program_code  # Retained specifically for Sheet 3 targeting
        }
        st.session_state.mileage_data = pd.concat([st.session_state.mileage_data, pd.DataFrame([new_entry])], ignore_index=True)
        st.success(f"Added! Google Route Distance: {google_miles} miles calculated.")
        st.rerun()

st.markdown("---")

# --- SECTION 3: DATA TABLE ---
st.header("📊 Mileage Log")

if not st.session_state.mileage_data.empty:
    total_miles = st.session_state.mileage_data["Calculated Mileage"].sum()
    st.metric(label="Total Period Mileage", value=f"{total_miles} miles")

    edited_df = st.data_editor(
        st.session_state.mileage_data, 
        num_rows="dynamic",
        use_container_width=True,
        key="mileage_editor"
    )
    st.session_state.mileage_data = edited_df
else:
    st.info("No mileage entries added yet.")

# --- NEW: WORKBOOK WRITE & GENERATION PIPELINE ---
if st.session_state.uploaded_file_bytes is not None:
    st.subheader("💾 Export Back to Uploaded Excel Template")
    st.markdown("Clicking the generator below safely builds a download bundle updating `Sheet 1` and `Sheet 3` without overriding formatting styles.")
    
    if st.button("Generate Updated Excel Document", type="secondary"):
        try:
            # Load active state stream object directly 
            output_wb = openpyxl.load_workbook(BytesIO(st.session_state.uploaded_file_bytes))
            
            # 1. Update Sheet 1 Cover Details
            s1 = output_wb.worksheets[0]
            s1["D11"] = st.session_state.employee_name
            s1["D15"] = st.session_state.date_range_str

            # 2. Update Sheet 3 Route Lines safely down rows
            if len(output_wb.worksheets) >= 3:
                s3 = output_wb.worksheets[2]
                
                # CHANGED FROM 4 TO 5: Start looking for empty rows under the headers
                current_write_row = 5
                while s3[f"B{current_write_row}"].value is not None:
                    current_write_row += 1
                
                # Write only the new rows added during this session
                # If we imported data, we look for rows that aren't placeholders
                new_session_rows = st.session_state.mileage_data[
                    st.session_state.mileage_data["Starting Location"] != "Imported from template"
                ]
                
                for _, row in new_session_rows.iterrows():
                    s3[f"B{current_write_row}"] = row["Date"]
                    s3[f"C{current_write_row}"] = row["Destination"]
                    s3[f"D{current_write_row}"] = row.get("Program Code", "")
                    s3[f"E{current_write_row}"] = row["Purpose of Travel"]
                    s3[f"F{current_write_row}"] = row["Calculated Mileage"]
                    current_write_row += 1
                    
            # 3. Stream back bytes array to download button container
            excel_stream = BytesIO()
            output_wb.save(excel_stream)
            excel_stream.seek(0)
            
            st.download_button(
                label="📥 Download Updated Excel File",
                data=excel_stream,
                file_name=f"Mileage_Report_{st.session_state.employee_name.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error compiling output modifications: {e}")

st.markdown("---")

# --- SECTION 4: GOOGLE MAPS ROUTE VISUALIZATION ---
# --- SECTION 4: GOOGLE MAPS ROUTE VISUALIZATION (COMPLETED) ---

st.header("🗺️ Route Map & Print View")



if not st.session_state.mileage_data.empty:

    # 1. Grab the most recent entry from the DataFrame

    last_entry = st.session_state.mileage_data.iloc[-1]

    origin = last_entry["Starting Location"]

    destination = last_entry["Destination"]

    trip_miles = last_entry["Calculated Mileage"]

    is_round_trip = last_entry["Round Trip"] == "Yes"

    entry_date = last_entry["Date"]

    entry_purpose = last_entry["Purpose of Travel"]

    

    # 2. Format labels dynamically based on Round Trip status

    if is_round_trip:

        route_label = f"{origin} ➡️ {destination} 🔄 {origin} (Round Trip)"

        mileage_label = f"{trip_miles} miles (Total Round Trip)"

    else:

        route_label = f"{origin} ➡️ {destination}"

        mileage_label = f"{trip_miles} miles (One Way)"

    

    st.subheader(f"Current Route Detail")

    col_m1, col_m2 = st.columns(2)

    col_m1.metric("Route Leg", route_label)

    col_m2.metric("Odometer Calculated Distance", mileage_label)

    

    # URL safe conversion for locations

    formatted_origin = origin.replace(" ", "+")

    formatted_destination = destination.replace(" ", "+")

    

    # 3. GENERATE THE PERFECT DIRECT PRINT LINK

    if is_round_trip:

        # Full web URL format: Origin -> Destination -> Back to Origin

        direct_maps_url = f"https://www.google.com/maps/dir/{formatted_origin}/{formatted_destination}/{formatted_origin}/"

    else:

        # Standard One Way

        direct_maps_url = f"https://www.google.com/maps/dir/{formatted_origin}/{formatted_destination}/"

        

    st.markdown("##### Actions")

    

    # Side-by-side uniform layout configuration

    action_col1, action_col2 = st.columns([1, 1])

    

    with action_col1:

        # Secure direct link mechanism targeting full maps app layout

        st.link_button("🖨️ Open Official Google Maps Print Layout", direct_maps_url, type="primary", use_container_width=True)

        st.caption("💡 *How to print:* In the new tab, press **Ctrl+P** (or **Cmd+P**) to trigger Google's clean print wizard.")



    with action_col2:

        # Build your dynamic target string

        text_to_copy = f"Date: {entry_date} | Purpose: {entry_purpose}"

        

        # Display via st.code to inherit native copy buttons safely without sandboxed JS errors

        st.code(text_to_copy, language="text")

        st.caption("📋 *Hover & click the icon on the right edge of the gray box above* to copy this text, then paste directly into Google's print notes context.")

                

    st.markdown("---")



    # 4. Render embedded visual route map

    if api_status == "Valid":

        if is_round_trip:

            # Multi-stop configuration: Start at A, end at A, waypoint via B

            map_url = (

                f"https://www.google.com/maps/embed/v1/directions"

                f"?key={st.session_state.api_key}"

                f"&origin={formatted_origin}"

                f"&destination={formatted_origin}"

                f"&waypoints={formatted_destination}"

                f"&mode=driving"

            )

        else:

            # Standard one-way visual trace mapping

            map_url = (

                f"https://www.google.com/maps/embed/v1/directions"

                f"?key={st.session_state.api_key}"

                f"&origin={formatted_origin}"

                f"&destination={formatted_destination}"

                f"&mode=driving"

            )

            

        st.components.v1.iframe(map_url, width=900, height=500)

    else:

        st.info("🔄 Map streaming paused because the API status is currently offline or invalid.")

        st.code(f"[Google Map Route Preview]\nRoute: {route_label}\nCalculated Distance: {mileage_label}")

else:

    st.write("Add an entry above to generate the live map route.")

st.header("🗺️ Route Map & Print View")


