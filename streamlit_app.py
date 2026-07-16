import streamlit as st
import pandas as pd
import datetime
import urllib.parse
import googlemaps
from streamlit_searchbox import st_searchbox
import openpyxl
from io import BytesIO

# --- PAGE CONFIG # [cite: 1]
st.set_page_config(page_title="GP Mileage Tool", layout="wide")

# --- CONSTANTS # [cite: 1, 2]
MILEAGE_COLUMNS = [
    "Date", "Starting Location", "Destination", "Round Trip",
    "Purpose of Travel", "Odometer Start", "Odometer End", "Calculated Mileage", "Program Code"
]

IMPORT_MARKER = "Imported from template"
METERS_TO_MILES = 0.000621371
DEFAULT_RATE_PER_MILE = 0.725

# Common locations to select from # [cite: 2]
COMMON_LOCATIONS = {
    "Custom / Type Address...": "",
    "Main Office": "1616 29th St, Bakersfield, CA 93301",
    "DEC": "1130 17th St, Bakersfield, CA 93301",
    "Delano": "1109 High St., Delano, CA 93215"
}

# --- API KEY & CLIENT INITIALIZATION # [cite: 2]
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

# --- GOOGLE PLACES AUTOCOMPLETE # [cite: 2, 3]
def search_google_places(query_text):
    if not query_text or not gmaps:
        return [] # [cite: 2, 3]
    try:
        bakersfield_coords = (35.3733, -119.0187) 
        search_radius = 50000 
        predictions = gmaps.places_autocomplete(
            input_text=query_text,
            location=bakersfield_coords,
            radius=search_radius,
            components={"country": "us"}
        )
        return [p['description'] for p in predictions]
    except Exception as e:
        print(f"Error fetching local places: {e}")
        return []

# --- GOOGLE DISTANCE MATRIX API # [cite: 4]
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

# --- HELPERS: TEMPLATE PROCESSING # [cite: 6]
def detect_and_extract_workbook(file_bytes, filename):
    """Detect template type and extract existing journey rows with header metadata."""
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
        sheet1 = wb.worksheets[0] 
        is_probation = bool(sheet1["C3"].value or sheet1["E3"].value or sheet1["E4"].value)
        
        template_type = "standard" 
        employee_name = "" 
        date_range_str = "" 
        rate_per_mile = DEFAULT_RATE_PER_MILE 
        existing_rows = []
        
        if is_probation:
            template_type = "at_promise"
            employee_name = str(sheet1["C3"].value or "").strip() 
            date_range_str = str(sheet1["E4"].value or "").strip() 
            try:
                rate_per_mile = float(sheet1["E3"].value or DEFAULT_RATE_PER_MILE) 
            except (ValueError, TypeError):
                rate_per_mile = DEFAULT_RATE_PER_MILE
                
            row_idx = 9 
            while sheet1[f"B{row_idx}"].value:
                raw_date = sheet1[f"B{row_idx}"].value 
                try:
                    formatted_date = pd.to_datetime(raw_date).strftime("%Y-%m-%d") 
                except Exception:
                    formatted_date = str(raw_date)[:10] if raw_date else "" 
                
                try:
                    odo_start = float(sheet1[f"F{row_idx}"].value or 0.0) 
                    odo_end = float(sheet1[f"G{row_idx}"].value or 0.0) 
                    calc_mileage = abs(odo_end - odo_start) if (odo_start and odo_end) else 0.0 
                except (ValueError, TypeError):
                    odo_start, odo_end, calc_mileage = 0.0, 0.0, 0.0
                    
                excel_start_loc = sheet1[f"C{row_idx}"].value 
                start_location_val = str(excel_start_loc).strip() if excel_start_loc else f"{IMPORT_MARKER} (Probation)" 
                    
                existing_rows.append({
                    "Date": formatted_date, 
                    "Starting Location": start_location_val, 
                    "Destination": str(sheet1[f"D{row_idx}"].value or ""), 
                    "Round Trip": "Yes", 
                    "Purpose of Travel": str(sheet1[f"E{row_idx}"].value or ""), 
                    "Odometer Start": odo_start, 
                    "Odometer End": odo_end, 
                    "Calculated Mileage": round(calc_mileage, 1), 
                    "Program Code": "", 
                    "_source_file": filename 
                })
                row_idx += 1 
        else:
            template_type = "standard" 
            employee_name = str(sheet1["D11"].value or "").strip() 
            date_range_str = str(sheet1["D15"].value or "").strip() 
            
            if len(wb.worksheets) >= 3: 
                sheet3 = wb.worksheets[2] 
                row_idx = 5 
                while sheet3[f"B{row_idx}"].value: 
                    raw_date = sheet3[f"B{row_idx}"].value 
                    try:
                        formatted_date = pd.to_datetime(raw_date).strftime("%Y-%m-%d")
                    except Exception:
                        formatted_date = str(raw_date)[:10] if raw_date else "" 
                        
                    existing_rows.append({
                        "Date": formatted_date, # [cite: 18]
                        "Starting Location": IMPORT_MARKER, # [cite: 18]
                        "Destination": str(sheet3[f"C{row_idx}"].value or ""),
                        "Round Trip": "Yes", # [cite: 19]
                        "Purpose of Travel": str(sheet3[f"E{row_idx}"].value or ""), 
                        "Odometer Start": 0, # [cite: 19, 20]
                        "Odometer End": 0, # [cite: 20]
                        "Calculated Mileage": float(sheet3[f"F{row_idx}"].value or 0.0), 
                        "Program Code": str(sheet3[f"D{row_idx}"].value or ""), 
                        "_source_file": filename 
                    })
                    row_idx += 1 # [cite: 21]

        has_uploaded_date = bool(date_range_str and date_range_str.strip()) 
        
        return {
            "filename": filename, # [cite: 21, 22]
            "bytes": file_bytes, # [cite: 22]
            "template_type": template_type, # [cite: 22]
            "employee_name": employee_name, # [cite: 22]
            "date_range_str": date_range_str, # [cite: 22]
            "has_uploaded_date": has_uploaded_date, # [cite: 22]
            "rate_per_mile": rate_per_mile, # [cite: 22]
            "imported_count": len(existing_rows), # [cite: 22, 23]
            "rows": existing_rows # [cite: 23]
        }
    except Exception as e:
        st.error(f"Error parsing structural configuration of {filename}: {e}") # [cite: 23]
        return None

# --- SESSION STATE INITIALIZATION # [cite: 23]
def init_session_state():
    defaults = {
        "mileage_data": pd.DataFrame(columns=MILEAGE_COLUMNS + ["_source_file"]), # [cite: 23]
        "employee_name": "", # [cite: 23]
        "date_range_str": "", # [cite: 23, 24]
        "rate_per_mile": DEFAULT_RATE_PER_MILE, # [cite: 24]
        "template_type": "standard",  
        "uploaded_files_registry": {},  
        "processed_file_hashes": set() 
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state() # [cite: 24]

# --- UI: HEADER
st.title("GP Mileage Tracker 🚗")
st.markdown("Easily calculate journeys, build route maps, and sync back to your formatted GP templates.")

# --- INITIALIZE STEP IN SESSION STATE
if "current_step" not in st.session_state:
    st.session_state.current_step = 1

# --- HELPER FUNCTION TO NAVIGATE
def go_to_step(step_number):
    st.session_state.current_step = step_number

# --- STEP PROGRESS INDICATOR
steps = {
    1: "📂 1. Upload Template",
    2: "📝 2. Cover Sheet",
    3: "📍 3. Log Journey",
    4: "📥 4. Review & Export"
}

cols = st.columns(len(steps))
for i, (step_num, step_name) in enumerate(steps.items(), 1):
    with cols[i-1]:
        if st.session_state.current_step == step_num:
            st.markdown(f"**🔵 {step_name}**")  # Active Step
        else:
            st.markdown(f"<span style='color: gray;'>⚪ {step_name}</span>", unsafe_allow_html=True)

st.markdown("---")
current_step = st.session_state.current_step

# Initialize a temporary state to hold the "calculated but uncommitted" preview trip
if "temp_preview_trip" not in st.session_state:
    st.session_state.temp_preview_trip = None
# ==========================================
# STEP 1: FILE UPLOAD & INGESTION
# ==========================================

if current_step == 1:
    st.header("Upload Mileage Excel Sheet")
    uploaded_templates = st.file_uploader(
        "Upload your mileage excel sheet (.xlsx)", # [cite: 24, 25]
        type=["xlsx"], 
        accept_multiple_files=True,
        key="excel_file_uploader"
    )

    # Clean up registry if files are deleted # [cite: 25]
    if "uploaded_files_registry" in st.session_state and st.session_state.uploaded_files_registry:
        current_uploader_names = [f.name for f in uploaded_templates] if uploaded_templates else [] # [cite: 25, 26]
        registered_names = list(st.session_state.uploaded_files_registry.keys()) # [cite: 26]
        removed_files = [name for name in registered_names if name not in current_uploader_names] # [cite: 26]
        
        if removed_files:
            for f_name in removed_files:
                st.session_state.uploaded_files_registry.pop(f_name, None) # [cite: 26]
                st.session_state.processed_file_hashes.discard(f_name) # [cite: 26]
                
                if not st.session_state.mileage_data.empty and "_source_file" in st.session_state.mileage_data.columns: # [cite: 26, 27]
                    st.session_state.mileage_data = st.session_state.mileage_data[
                        st.session_state.mileage_data["_source_file"] != f_name # [cite: 27]
                    ]
            
            if not st.session_state.uploaded_files_registry: # [cite: 27, 28]
                st.session_state.employee_name = "" # [cite: 28]
                st.session_state.date_range_str = "" # [cite: 28]
                st.session_state.has_uploaded_date = False # [cite: 28]
                st.session_state.pop("cs_employee_name", None) # [cite: 28]
                st.session_state.pop("cs_month_select", None) # [cite: 28, 29]
            else:
                remaining_file = list(st.session_state.uploaded_files_registry.values())[0] # [cite: 29]
                st.session_state.template_type = remaining_file["template_type"] # [cite: 29]
                st.session_state.employee_name = remaining_file["employee_name"] # [cite: 29]
                st.session_state.date_range_str = remaining_file["date_range_str"] # [cite: 29]
                st.session_state.rate_per_mile = remaining_file["rate_per_mile"] # [cite: 29, 30]
                
            st.toast("🗑️ File removed and database synchronized.") # [cite: 30]
            st.rerun()

    if uploaded_templates:
        state_updated = False # [cite: 30]
        for uploaded_file in uploaded_templates:
            if uploaded_file.name not in st.session_state.processed_file_hashes: # [cite: 30]
                file_bytes = uploaded_file.getvalue() # [cite: 30]
                parsed_result = detect_and_extract_workbook(file_bytes, uploaded_file.name) # [cite: 30]
                
                if parsed_result: # [cite: 31]
                    if parsed_result["rows"]: # [cite: 31]
                        for r in parsed_result["rows"]: # [cite: 31]
                            r["_source_file"] = uploaded_file.name # [cite: 31, 32]
                    
                    st.session_state.uploaded_files_registry[uploaded_file.name] = parsed_result # [cite: 32]
                    st.session_state.processed_file_hashes.add(uploaded_file.name) # [cite: 32]
                    
                    st.session_state.template_type = parsed_result["template_type"] # [cite: 32]
                    st.session_state.employee_name = parsed_result["employee_name"] # [cite: 32, 33]
                    st.session_state.date_range_str = parsed_result["date_range_str"] # [cite: 33]
                    st.session_state.rate_per_mile = parsed_result["rate_per_mile"] # [cite: 33]
                    
                    if parsed_result["rows"]: # [cite: 33]
                        new_df = pd.DataFrame(parsed_result["rows"]) # [cite: 33, 34]
                        st.session_state.mileage_data = pd.concat( # [cite: 34]
                            [st.session_state.mileage_data, new_df], # [cite: 34]
                            ignore_index=True # [cite: 34]
                        ) # [cite: 35]
                    st.toast(f" Imported {parsed_result['imported_count']} records from {uploaded_file.name} ({parsed_result['template_type'].upper()})") # [cite: 35]
                    state_updated = True # [cite: 35]
                    
        if state_updated: # [cite: 35]
            st.rerun() # [cite: 35]

            # --- DYNAMIC "NEXT" BUTTON
        if len(st.session_state.uploaded_files_registry) > 0:
            st.success("🎉 Base template(s) loaded successfully!")
            st.button("Continue to Cover Sheet Setup ➡️", on_click=go_to_step, args=(2,), use_container_width=True)
            
# ==========================================
# STEP 2: COVER SHEET METADATA
# ==========================================
elif current_step == 2:
    st.header("Cover Sheet Information")
    col1, col2, col3 = st.columns([2, 2, 1])

    with col1:
        # Properly synchronize session states without typing interference
        if "cs_employee_name" not in st.session_state:
            st.session_state["cs_employee_name"] = st.session_state.employee_name

        employee_name = st.text_input(
            "Employee Name",
            placeholder="John Doe",
            key="cs_employee_name"
        )
        st.session_state.employee_name = employee_name

    with col2:
        months = [
            "January", "February", "March", "April", "May", "June", 
            "July", "August", "September", "October", "November", "December"
        ]
        month_days = {
            "January": 31, "February": 28, "March": 31, "April": 30, 
            "May": 31, "June": 30, "July": 31, "August": 31, 
            "September": 30, "October": 31, "November": 30, "December": 31
        }

        def extract_month_name(date_str):
            if not date_str or not isinstance(date_str, str):
                return None
            date_str = date_str.strip()
            cleaned_str = date_str.replace("-", "/")
            if "/" in cleaned_str:
                parts = cleaned_str.split("/")
                if parts[0].isdigit():
                    month_num = int(parts[0])
                    if 1 <= month_num <= 12:
                        return months[month_num - 1]
            for m in months:
                if m.lower() in date_str.lower():
                    return m
            return None

        detected_months = []
        active_files = list(st.session_state.get("uploaded_files_registry", {}).values())

        for f in active_files:
            raw_date_field = f.get("date_range_str", "")
            parsed_month = extract_month_name(raw_date_field)
            if parsed_month:
                detected_months.append(parsed_month)

        should_lock = len(active_files) > 0 and len(detected_months) == len(active_files)

        if should_lock:
            locked_month = detected_months[0]
            days = month_days[locked_month]
            st.session_state.date_range_str = f"{locked_month} 1 - {locked_month} {days}, 2026"
            st.text_input(
                "Reporting Period Range (Locked from Import)", 
                value=st.session_state.date_range_str, 
                disabled=True, 
                key="static_period_display"
            )
        else:
            default_idx = 0
            current_state_month = extract_month_name(st.session_state.get("date_range_str", ""))
            if current_state_month in months:
                default_idx = months.index(current_state_month)

            selected_month = st.selectbox(
                "Select Reporting Month (2026)",
                options=months,
                index=default_idx,
                key="cs_month_select"
            )
            
            days_in_month = month_days[selected_month]
            st.session_state.date_range_str = f"{selected_month} 1 - {selected_month} {days_in_month}, 2026"
            st.caption(f" **Formatted Range:** {st.session_state.date_range_str}")

    with col3:
        if st.session_state.template_type == "at_promise":
            rate_per_mile = st.number_input(
                "Rate per Mile ($)",
                value=float(st.session_state.rate_per_mile),
                min_value=0.0,
                step=0.01,
                format="%.3f",
                key="cs_rate_per_mile"
            )
            st.session_state.rate_per_mile = rate_per_mile
        else:
            st.session_state.rate_per_mile = DEFAULT_RATE_PER_MILE

    # --- NAVIGATION BUTTONS
    col_back, col_next = st.columns(2)
    with col_back:
        st.button("⬅️ Back to Upload", on_click=go_to_step, args=(1,), use_container_width=True)
    with col_next:
        if st.session_state.employee_name:
            st.button("Continue to Journey Logger ➡️", on_click=go_to_step, args=(3,), use_container_width=True)
        else:
            st.warning("⚠️ Please enter your Employee Name to proceed.")


# ==========================================
# STEP 3: LOG NEW TRIPS & VIEW MAP ROUTES
# ==========================================
elif current_step == 3:
    st.header("📍 Track Your Journey")
    st.markdown("Calculate distances, add dynamic stops along your route, and track travel details below.")

    if "temp_preview_trip" not in st.session_state:
        st.session_state.temp_preview_trip = None

    if "form_generation" not in st.session_state:
        st.session_state.form_generation = 0 

    if "num_stops" not in st.session_state:
        st.session_state.num_stops = 0 

    gen = st.session_state.form_generation 
    today = datetime.date.today() 

    # --- SECTION 1: ROUTE & ADDRESS CONFIGURATION ---
    st.subheader("1. Route Details")
    
    # Date Selection Row (keeps it neat and concise)
    col_date, _ = st.columns([1, 2])
    with col_date:
        travel_date = st.date_input("Date of Travel", value=today, key=f"journey_travel_date_{gen}")

    # Start & Destination side-by-side card container
    route_container = st.container(border=True)
    with route_container:
        col_start, col_dest = st.columns(2)

        # Starting Location Column
        with col_start:
            st.markdown("**Departure Point**")
            selected_shortcut = st.selectbox( 
                "Quick select standard location:",
                options=list(COMMON_LOCATIONS.keys()),
                key=f"start_shortcut_{gen}"
            ) 
            
            if selected_shortcut == "Custom / Type Address...": 
                start_loc = st_searchbox( 
                    search_google_places, 
                    key=f"start_location_search_{gen}",
                    placeholder="🔍 Type starting address..."
                ) 
            else:
                start_loc = COMMON_LOCATIONS[selected_shortcut] 
                st.caption(f"📍 Selected: `{start_loc}`") 

        # Final Destination Column
        with col_dest:
            st.markdown("**Arrival Destination**")
            selected_dest_shortcut = st.selectbox( 
                "Quick select standard location:",
                options=list(COMMON_LOCATIONS.keys()), 
                key=f"dest_shortcut_{gen}"
            ) 

            if selected_dest_shortcut == "Custom / Type Address...": 
                dest_loc = st_searchbox( 
                    search_google_places,
                    key=f"destination_search_{gen}",
                    placeholder="🔍 Type destination address..."
                ) 
            else:
                dest_loc = COMMON_LOCATIONS[selected_dest_shortcut] 
                st.caption(f"📍 Selected: `{dest_loc}`")

    # --- INTERMEDIATE STOPS SUBSECTION ---
    st.markdown("##### Multi-Stop Route Planning")
    
    # If we have active stops, display them inside a cleanly bordered list
    if st.session_state.num_stops > 0:
        stops_container = st.container(border=True)
        with stops_container:
            additional_stops = [] 
            for i in range(st.session_state.num_stops): 
                stop = st_searchbox( 
                    search_google_places,
                    key=f"stop_search_{i}_{gen}",
                    placeholder=f"🔍 Search address for intermediate stop #{i+1}..."
                ) 
                if stop: 
                    additional_stops.append(stop)
    else:
        additional_stops = []

    # Dynamic Control Buttons (Pills/Inline look)
    col_btn1, col_btn2, _ = st.columns([1, 1, 2])
    with col_btn1:
        if st.button("✚ Add Stop", key=f"add_stop_btn_{gen}", use_container_width=True): 
            st.session_state.num_stops += 1 
            st.rerun() 
    with col_btn2:
        if st.button("▬ Remove", key=f"rem_stop_btn_{gen}", use_container_width=True) and st.session_state.num_stops > 0: 
            st.session_state.num_stops -= 1 
            st.rerun() 

    # --- SECTION 2: WORKBOOK & METADATA SECTION ---
    st.markdown("---")
    st.subheader("2. Travel Purpose & Metadata")
    
    # Determine fields based on configurations
    active_registry = st.session_state.get("uploaded_files_registry", {})
    has_standard = any(meta["template_type"] == "standard" for meta in active_registry.values())
    
    if active_registry:
        show_program_code = has_standard
    else:
        show_program_code = (st.session_state.template_type == "standard")
    
    program_code = ""
    round_trip = "Yes"

    metadata_container = st.container(border=True)
    with metadata_container:
        if not show_program_code:
            col_purpose, col_calc = st.columns([3, 1])
            with col_purpose:
                purpose = st.text_input("Purpose of Travel", key=f"journey_purpose_{gen}", placeholder="e.g., Client Visit / Site Audit")
        else: 
            col_purpose, col_prog_code, col_rt = st.columns([2, 1, 1])
            with col_purpose:
                purpose = st.text_input("Purpose of Travel", key=f"journey_purpose_{gen}", placeholder="e.g., Client Visit / Site Audit")
            with col_prog_code:
                program_code = st.text_input(
                    "Program Code",
                    placeholder="e.g., 101",
                    key=f"journey_prog_code_{gen}"
                )
            with col_rt:
                round_trip = st.selectbox("Round Trip?", ["Yes", "No"], key=f"journey_round_trip_{gen}")

        # Dynamic Odometer Fields
        if st.session_state.template_type == "at_promise": 
            st.markdown("**Odometer Count** *(Required for Probation Form)*") 
            col_odo_start, col_odo_end = st.columns(2) 

            with col_odo_start:
                odo_start_input = st.text_input( 
                    "Odometer Start",
                    placeholder="e.g., 45100",
                    key=f"journey_odo_start_{gen}"
                ) 
            with col_odo_end:
                odo_end_input = st.text_input( 
                    "Odometer End",
                    placeholder="e.g., 45125",
                    key=f"journey_odo_end_{gen}" 
                )
        else:
            odo_start_input = "" 
            odo_end_input = ""

    # --- SECTION 3: CALCULATION, PREVIEW & SAVE ---
    st.markdown("---")
    
    # Prominent Primary Action Button
    calculate_button = st.button("Calculate & Preview Route 🗺️", type="primary", key=f"journey_calc_btn_{gen}", use_container_width=True) 

    if calculate_button: 
        if not start_loc or not dest_loc: 
            st.error("⚠️ Please provide both a Starting Location and Destination.") 
        else:
            with st.spinner("Calculating distances..."):
                current_origin = start_loc 
                google_miles = 0.0 
                
                for stop in additional_stops: 
                    google_miles += get_google_distance_miles(current_origin, stop) 
                    current_origin = stop 
                    
                google_miles += get_google_distance_miles(current_origin, dest_loc) 
                calculated_miles = google_miles * 2 if round_trip == "Yes" else google_miles 
                calculated_miles = round(calculated_miles, 1) 

                if additional_stops: 
                    stops_str = " -> ".join(additional_stops) 
                    combined_destination = f"{stops_str} -> {dest_loc}" 
                else: 
                    combined_destination = dest_loc 
                    
                if round_trip == "Yes": 
                    combined_destination = f"{combined_destination} (RT)" 

                def parse_to_float(val): 
                    if not val or not str(val).strip():
                        return None
                    try:
                        return round(float(val.strip()), 1) 
                    except ValueError:
                        return None 

                o_start = parse_to_float(odo_start_input)
                o_end = parse_to_float(odo_end_input)
                
                if o_start is not None and o_end is None: 
                    o_end = round(o_start + calculated_miles, 1) 
                elif o_end is not None and o_start is None: 
                    st_val = round(o_end - calculated_miles, 1)
                    o_start = max(0.0, st_val)
                elif o_start is not None and o_end is not None: 
                    calculated_miles = round(o_end - o_start, 1) 
                else:
                    o_start, o_end = 0.0, round(calculated_miles, 1) 

                # Populate temporary preview details in state
                st.session_state.temp_preview_trip = {
                    "Date": travel_date.strftime("%Y-%m-%d") if isinstance(travel_date, (datetime.date, datetime.datetime)) else str(travel_date), 
                    "Starting Location": start_loc, 
                    "Destination": combined_destination, 
                    "Round Trip": round_trip, 
                    "Purpose of Travel": purpose, 
                    "Odometer Start": o_start, 
                    "Odometer End": o_end, 
                    "Calculated Mileage": calculated_miles, 
                    "Program Code": program_code, 
                    "_source_file": "manual_entry" 
                }
                st.rerun()

    # Map details and print view for the calculated draft route
    if st.session_state.temp_preview_trip is not None: 
        preview = st.session_state.temp_preview_trip
        origin = preview["Starting Location"] 
        raw_destination = str(preview["Destination"]) 
        clean_destination_chain = raw_destination.replace(" (RT)", "") 
        all_dest_legs = [leg.strip() for leg in clean_destination_chain.split(" -> ")] 
        final_destination = all_dest_legs[-1] 
        intermediate_waypoints = all_dest_legs[:-1] if len(all_dest_legs) > 1 else [] 
        trip_miles = preview["Calculated Mileage"] 
        is_round_trip = preview["Round Trip"] == "Yes" 
        entry_date = preview["Date"] 
        entry_purpose = preview["Purpose of Travel"] 

        st.subheader("🗺️ Live Route Preview") 
        with st.container(border=True): 
            st.markdown(f"**Date of Travel:** `{entry_date}` | **Purpose:** {entry_purpose}") 
            timeline_steps = [f"**{origin}** (Start)->"] 
            for wp in intermediate_waypoints: 
                timeline_steps.append(f"`Stop: {wp}`") 
            timeline_steps.append(f" **{final_destination}** (Destination) | ") 
            if is_round_trip: 
                timeline_steps.append(f" **{origin}** (RT)") 
            st.markdown(" ".join(timeline_steps)) 

        col_metric_dist, col_metric_type, col_metric_status = st.columns(3) 
        with col_metric_dist: 
            st.metric(label="Calculated Distance", value=f"{trip_miles} mi") 
        with col_metric_type: 
            st.metric(label="Trip Type", value="Round Trip" if is_round_trip else "One Way") 
        with col_metric_status: 
            st.metric( 
                label="Final Stop", 
                value=final_destination if not is_round_trip else "Returned Back", 
                help="The primary destination point of this recorded log." 
            )

        encoded_origin = urllib.parse.quote_plus(origin) 
        encoded_final_destination = urllib.parse.quote_plus(final_destination) 
        maps_url_legs = [encoded_origin] + [urllib.parse.quote_plus(wp) for wp in intermediate_waypoints] + [encoded_final_destination] 
        if is_round_trip: 
            maps_url_legs.append(encoded_origin) 
            
        direct_maps_url = f"https://www.google.com/maps/dir/{'/'.join(maps_url_legs)}/" 

        if gmaps: 
            embed_waypoints = list(intermediate_waypoints) 
            if is_round_trip: 
                embed_waypoints.append(final_destination) 
                embed_destination_target = origin 
            else: 
                embed_destination_target = final_destination 
                
            encoded_embed_waypoints = "|".join([urllib.parse.quote_plus(wp) for wp in embed_waypoints]) 
        
            # FIX: Used the correct and authorized Google Maps Embed API endpoint
            map_url = ( 
                f"https://www.google.com/maps/embed/v1/directions" 
                f"?key={api_key}" 
                f"&origin={encoded_origin}" 
                f"&destination={urllib.parse.quote_plus(embed_destination_target)}" 
            )
            if encoded_embed_waypoints: 
                map_url += f"&waypoints={encoded_embed_waypoints}" 
            st.components.v1.iframe(map_url, width=900, height=500) 
        else:
            st.warning("Please add a valid Google Maps API Key.") 
            
        st.caption("**Check the miles:** If the map above shows a different total than the calculation, adjust it below before submitting.") 
    
        adjusted_miles = st.number_input( 
            "Confirm/Override Mileage:",
            min_value=0.0,
            value=float(trip_miles),
            step=0.1,
            format="%.1f",
            key="preview_override_input"
        )
        if adjusted_miles != float(trip_miles): 
            st.session_state.temp_preview_trip["Calculated Mileage"] = round(adjusted_miles, 1) 
            
            odo_start = float(preview["Odometer Start"] or 0.0) 
            odo_end = float(preview["Odometer End"] or 0.0) 
            
            start_is_whole = odo_start.is_integer() 
            end_is_whole = odo_end.is_integer() 
            
            if start_is_whole and not end_is_whole: 
                st.session_state.temp_preview_trip["Odometer End"] = round(odo_start + adjusted_miles, 1) 
            elif end_is_whole and not start_is_whole: 
                st.session_state.temp_preview_trip["Odometer Start"] = round(max(0.0, odo_end - adjusted_miles), 1) 
            else: 
                st.session_state.temp_preview_trip["Odometer End"] = round(odo_start + adjusted_miles, 1) 
             
            st.rerun() 
        
        st.markdown("##### Actions") 
        col_copy, col_print = st.columns(2)

        with col_copy:
            raw_name = st.session_state.get("employee_name", "") 
            name_parts = raw_name.split() 
            initials = "".join([part[0].upper() for part in name_parts if part]) 
            initials_suffix = f" | Initials: {initials}" if initials else "" 
            text_to_copy = f"Date: {entry_date} | Purpose: {entry_purpose} | Miles: {adjusted_miles}{initials_suffix}" 
            st.code(text_to_copy, language="text") 
            st.caption("📋 Click the copy icon to save trip details") 
                
        with col_print:
            st.link_button( 
                "🔗 Open in Google Maps",
                direct_maps_url, 
                type="secondary",
                use_container_width=True
            )
            st.caption("press **Ctrl+P** to print out the Map Route") 

        # BUTTON 2: SUBMIT TRIP (Modified to safely redirect using on_click callback routing)
        st.write(" ")
        def handle_submit():
            st.session_state.mileage_data = pd.concat(
                [st.session_state.mileage_data, pd.DataFrame([st.session_state.temp_preview_trip])],
                ignore_index=True
            )
            st.session_state.temp_preview_trip = None
            st.session_state.form_generation += 1 
            st.session_state.num_stops = 0 
            go_to_step(4)

        submit_log_btn = st.button(
            "🎯 Submit Trip to Log & Continue", 
            type="primary", 
            use_container_width=True,
            on_click=handle_submit
        )

    else:
        st.info("💡 Fill out the form fields above and click 'Calculate & Preview Route' to generate your route map.")

    # Bottom Step Navigation
    st.markdown("---")
    col_back, col_next_fallback = st.columns(2)
    with col_back:
        st.button("⬅️ Back to Cover Sheet", on_click=go_to_step, args=(2,), use_container_width=True)
    with col_next_fallback:
        st.button("Skip to Review & Export ➡️", on_click=go_to_step, args=(4,), use_container_width=True)

# ==========================================
# STEP 4: MILEAGE LOG TABLE & EXPORT BACK TO EXCEL
# ==========================================
elif current_step == 4:
    st.header("Review & Export")

    if not st.session_state.mileage_data.empty:
        st.subheader("Edit/Verify Logs")
        st.caption("**Tip:** Double-click any calculated mileage cell to make manual corrections directly in the sheet tables below.")

        # If templates are uploaded, split them into separate tabs
        if st.session_state.uploaded_files_registry:
            workbook_names = list(st.session_state.uploaded_files_registry.keys())
            tabs = st.tabs([f"📋 {name}" for name in workbook_names])

            for idx, filename in enumerate(workbook_names):
                with tabs[idx]:
                    meta = st.session_state.uploaded_files_registry[filename]
                    is_probation = (meta["template_type"] == "at_promise")

                    # Filter: Show this workbook's imported rows OR any manually added entries
                    file_subset = st.session_state.mileage_data[
                        (st.session_state.mileage_data["_source_file"] == filename) |
                        (st.session_state.mileage_data["_source_file"] == "manual_entry")
                    ]

                    if file_subset.empty:
                        st.info("No logs generated or imported for this workbook yet.")
                        continue

                    # --- 1. NEW: SCOPED TALLY SECTION PER TAB ---
                    tab_total_miles = file_subset["Calculated Mileage"].sum()
                    tab_reimbursement_amt = tab_total_miles * st.session_state.rate_per_mile
                    
                    col_m1, col_m2, col_m3 = st.columns([1, 1, 1])
                    with col_m1:
                        st.metric("Total Workbook Mileage", f"{tab_total_miles:.1f} miles")
                    with col_m2:
                        st.metric("Reimbursement Rate", f"${st.session_state.rate_per_mile:.3f} / mi")
                    with col_m3:
                        st.metric("Estimated Reimbursement", f"${tab_reimbursement_amt:.2f}")

                    st.markdown("---")

                    # Set up columns based on the specific template type of this tab
                    if is_probation:
                        display_columns = [col for col in MILEAGE_COLUMNS if col != "Program Code"]
                        column_configuration = {
                            "Calculated Mileage": st.column_config.NumberColumn(
                                "Calculated Mileage",
                                help="Double-click to override with manual app miles if needed.",
                                format="%.1f",
                                min_value=0.0,
                                required=True
                            )
                        }
                    else:
                        display_columns = [col for col in MILEAGE_COLUMNS if col not in ["Odometer Start", "Odometer End"]]
                        column_configuration = {
                            "Calculated Mileage": st.column_config.NumberColumn(
                                "Calculated Mileage",
                                help="Double-click to override with manual app miles if needed.",
                                format="%.1f",
                                min_value=0.0,
                                required=True
                            )
                        }

                    # Render the interactive data editor for this specific workbook tab
                    edited_subset = st.data_editor(
                        file_subset,
                        column_order=display_columns, 
                        num_rows="dynamic",
                        use_container_width=True,
                        key=f"mileage_editor_{filename}",
                        column_config=column_configuration
                    )
                    
                    # Check for edits and sync them back to the main session DataFrame
                    if not edited_subset.equals(file_subset):
                        for sub_idx in edited_subset.index:
                            if sub_idx in st.session_state.mileage_data.index:
                                old_mileage = st.session_state.mileage_data.loc[sub_idx, "Calculated Mileage"]
                                new_mileage = edited_subset.loc[sub_idx, "Calculated Mileage"]
                                
                                if old_mileage != new_mileage:
                                    # Convert odometer targets to floats
                                    odo_start = float(edited_subset.loc[sub_idx, "Odometer Start"] or 0.0)
                                    odo_end = float(edited_subset.loc[sub_idx, "Odometer End"] or 0.0)
                                    
                                    start_has_decimal = "." in str(odo_start) and not str(odo_start).endswith(".0")
                                    end_has_decimal = "." in str(odo_end) and not str(odo_end).endswith(".0")
                                    
                                    if start_has_decimal and not end_has_decimal:
                                        edited_subset.loc[sub_idx, "Odometer End"] = round(odo_start + new_mileage, 1)
                                    elif end_has_decimal and not start_has_decimal:
                                        edited_subset.loc[sub_idx, "Odometer Start"] = round(max(0.0, odo_end - new_mileage), 1)
                                    else:
                                        edited_subset.loc[sub_idx, "Odometer End"] = round(odo_start + new_mileage, 1)
                        
                        # Update the global state with the edited subset using the matched index
                        st.session_state.mileage_data.update(edited_subset)
                        st.rerun()
        else:
            # Fallback if no files are uploaded yet (just display a clean default editor and global metrics)
            st.info("Upload Excel templates in Step 1 to organize logs into workbook tabs.")
            
            total_miles = st.session_state.mileage_data["Calculated Mileage"].sum()
            reimbursement_amt = total_miles * st.session_state.rate_per_mile
            
            col_m1, col_m2, col_m3 = st.columns([1, 1, 1])
            with col_m1:
                st.metric("Total Period Mileage", f"{total_miles:.1f} miles")
            with col_m2:
                st.metric("Reimbursement Rate", f"${st.session_state.rate_per_mile:.3f} / mi")
            with col_m3:
                st.metric("Estimated Reimbursement", f"${reimbursement_amt:.2f}")

            st.markdown("---")
            
            display_columns = [col for col in MILEAGE_COLUMNS if col not in ["Odometer Start", "Odometer End"]]
            edited_df = st.data_editor(
                st.session_state.mileage_data,
                column_order=display_columns, 
                num_rows="dynamic",
                use_container_width=True,
                key="mileage_editor_fallback"
            )
            if not edited_df.equals(st.session_state.mileage_data):
                st.session_state.mileage_data = edited_df
                st.rerun()
    else:
        st.info("No mileage entries added yet. Go back to Step 3 or upload a template in Step 1 to get started.")

    st.markdown("---")

    # Excel Exporter Section
    if st.session_state.uploaded_files_registry:
        st.subheader("Export Back to Excel Templates")
        st.markdown("This will dynamically append only new manual entries into your clean target workbook files while retaining original styling layouts.")
        
        if "_source_file" in st.session_state.mileage_data.columns:
            new_session_rows = st.session_state.mileage_data[
                st.session_state.mileage_data["_source_file"] == "manual_entry"
            ]
        else:
            new_session_rows = st.session_state.mileage_data[
                ~st.session_state.mileage_data["Starting Location"].str.contains(IMPORT_MARKER, na=False)
            ]
        
        for filename, meta in st.session_state.uploaded_files_registry.items():
            is_probation = (meta["template_type"] == "at_promise")
            template_display = "AT-PROMISE (Probation)" if is_probation else "Standard (GP)"
            form_label = "Probation" if is_probation else "GP"
            
            with st.expander(f"📥 Download Config: {filename} ({template_display})", expanded=True):
                if st.button(f"Generate Updated File for {filename}", key=f"gen_btn_{filename}"):
                    try:
                        output_wb = openpyxl.load_workbook(BytesIO(meta["bytes"]))
                        s1 = output_wb.worksheets[0]
                        
                        if is_probation:
                            s1["C3"] = st.session_state.employee_name
                            s1["E4"] = st.session_state.date_range_str
                            s1["E3"] = st.session_state.rate_per_mile
                            
                            current_write_row = 9 + meta["imported_count"]
                            for _, row in new_session_rows.iterrows():
                                s1[f"B{current_write_row}"] = row["Date"]
                                s1[f"C{current_write_row}"] = row["Starting Location"]
                                s1[f"D{current_write_row}"] = row["Destination"]
                                s1[f"E{current_write_row}"] = row["Purpose of Travel"]
                                s1[f"F{current_write_row}"] = row["Odometer Start"]
                                s1[f"G{current_write_row}"] = row["Odometer End"]
                                current_write_row += 1
                        else:
                            s1["D11"] = st.session_state.employee_name
                            s1["D15"] = st.session_state.date_range_str

                            if len(output_wb.worksheets) >= 3:
                                s3 = output_wb.worksheets[2]
                                current_write_row = 5
                                while s3[f"B{current_write_row}"].value:
                                    current_write_row += 1
                                
                                for _, row in new_session_rows.iterrows():
                                    s3[f"B{current_write_row}"] = row["Date"]
                                    s3[f"C{current_write_row}"] = row["Destination"]
                                    s3[f"D{current_write_row}"] = row.get("Program Code", "")
                                    s3[f"E{current_write_row}"] = row["Purpose of Travel"]
                                    s3[f"F{current_write_row}"] = row["Calculated Mileage"]
                                    current_write_row += 1
                        
                        excel_stream = BytesIO()
                        output_wb.save(excel_stream)
                        excel_stream.seek(0)
                        
                        today_str = datetime.date.today().strftime("%Y-%m-%d")

                        st.download_button(
                            label=f"📥 Download Updated {filename}",
                            data=excel_stream,
                            file_name=f"{form_label}_Mileage_{today_str}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_btn_{filename}",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Failed to append records to target workbook {filename}: {e}")
    else:
        st.info("Upload excel sheets in Step 1 to enable spreadsheet exports.")

    # --- NAVIGATION BUTTON
    st.markdown("---")
    st.button("⬅️ Back to Logger", on_click=go_to_step, args=(3,), use_container_width=True)

