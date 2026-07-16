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
    st.rerun()

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
    st.header("Cover Sheet Information") # [cite: 35]
    col1, col2, col3 = st.columns([2, 2, 1]) # [cite: 35]

    with col1:
        if st.session_state.employee_name and not st.session_state.get("cs_employee_name"): # [cite: 35]
            st.session_state["cs_employee_name"] = st.session_state.employee_name # [cite: 35, 36]

        employee_name = st.text_input( # [cite: 36]
            "Employee Name",
            placeholder="John Doe",
            key="cs_employee_name"
        ) # [cite: 36]
        st.session_state.employee_name = employee_name # [cite: 36]

    with col2:
        months = [
            "January", "February", "March", "April", "May", "June", 
            "July", "August", "September", "October", "November", "December"
        ] # [cite: 36]
        month_days = {
            "January": 31, "February": 28, "March": 31, "April": 30, # [cite: 36, 37]
            "May": 31, "June": 30, "July": 31, "August": 31, 
            "September": 30, "October": 31, "November": 30, "December": 31
        } # [cite: 37]

        def extract_month_name(date_str): # [cite: 37]
            if not date_str or not isinstance(date_str, str): # [cite: 37]
                return None
            date_str = date_str.strip() # [cite: 37]
            cleaned_str = date_str.replace("-", "/") # [cite: 38]
            if "/" in cleaned_str: # [cite: 38]
                parts = cleaned_str.split("/") # [cite: 38]
                if parts[0].isdigit(): # [cite: 38]
                    month_num = int(parts[0]) # [cite: 38, 39]
                    if 1 <= month_num <= 12: # [cite: 39]
                        return months[month_num - 1] # [cite: 39]
            for m in months: # [cite: 39, 40]
                if m.lower() in date_str.lower(): # [cite: 40]
                    return m # [cite: 40]
            return None # [cite: 40]

        detected_months = [] # [cite: 40]
        active_files = list(st.session_state.get("uploaded_files_registry", {}).values()) # [cite: 40]

        for f in active_files: # [cite: 40]
            raw_date_field = f.get("date_range_str", "") # [cite: 40]
            parsed_month = extract_month_name(raw_date_field) # [cite: 41]
            if parsed_month: # [cite: 41]
                detected_months.append(parsed_month) # [cite: 41]

        should_lock = len(active_files) > 0 and len(detected_months) == len(active_files) # [cite: 41]

        if should_lock: # [cite: 41]
            locked_month = detected_months[0] # [cite: 41, 42]
            days = month_days[locked_month] # [cite: 42]
            st.session_state.date_range_str = f"{locked_month} 1 - {locked_month} {days}, 2026" # [cite: 42]
            st.text_input( # [cite: 42]
                "Reporting Period Range (Locked from Import)", # [cite: 42]
                value=st.session_state.date_range_str, # [cite: 42, 43]
                disabled=True, 
                key="static_period_display"
            ) # [cite: 43]
        else:
            default_idx = 0 # [cite: 43]
            current_state_month = extract_month_name(st.session_state.get("date_range_str", "")) # [cite: 43, 44]
            if current_state_month in months: # [cite: 44]
                default_idx = months.index(current_state_month) # [cite: 44]

            selected_month = st.selectbox( # [cite: 44]
                "Select Reporting Month (2026)",
                options=months,
                index=default_idx,
                key="cs_month_select"
            ) # [cite: 44]
            
            days_in_month = month_days[selected_month] # [cite: 45]
            st.session_state.date_range_str = f"{selected_month} 1 - {selected_month} {days_in_month}, 2026" # [cite: 45]
            st.caption(f" **Formatted Range:** {st.session_state.date_range_str}") # [cite: 45]

    with col3:
        if st.session_state.template_type == "at_promise": # [cite: 45]
            rate_per_mile = st.number_input( # [cite: 45]
                "Rate per Mile ($)",
                value=float(st.session_state.rate_per_mile), # [cite: 45, 46]
                min_value=0.0, # [cite: 46]
                step=0.01,
                format="%.3f",
                key="cs_rate_per_mile"
            ) # [cite: 46]
            st.session_state.rate_per_mile = rate_per_mile # [cite: 46]
        else:
            st.session_state.rate_per_mile = DEFAULT_RATE_PER_MILE # [cite: 46]

    # --- NAVIGATION BUTTONS
    col_back, col_next = st.columns(2)
    with col_back:
        st.button("⬅️ Back to Upload", on_click=go_to_step, args=(1,), use_container_width=True)
    with col_next:
        # Show "Next" button only if they have filled in their name
        if st.session_state.employee_name:
            st.button("Continue to Journey Logger ➡️", on_click=go_to_step, args=(3,), use_container_width=True)
        else:
            st.warning("⚠️ Please enter your Employee Name to proceed.")

# ==========================================
# STEP 3: LOG NEW TRIPS & VIEW MAP ROUTES
# ==========================================
elif current_step == 3:
    st.header("Track Your Journey")

    # Add Journey Entry Form # [cite: 46]
    if "form_generation" not in st.session_state:
        st.session_state.form_generation = 0 # [cite: 46]

    gen = st.session_state.form_generation # [cite: 46]
    today = datetime.date.today() # [cite: 46, 47]
    col_date, col_start, col_dest = st.columns(3) # [cite: 47]

    with col_date:
        travel_date = st.date_input("Date", value=today, key=f"journey_travel_date_{gen}") # [cite: 47]

    with col_start:
        st.write("**Starting Location**") # [cite: 47]
        selected_shortcut = st.selectbox( # [cite: 47]
            "Quick Select Location",
            options=list(COMMON_LOCATIONS.keys()),
            key=f"start_shortcut_{gen}",
            label_visibility="collapsed"
        ) # [cite: 47]
        
        if selected_shortcut == "Custom / Type Address...": # [cite: 47]
            start_loc = st_searchbox( # [cite: 47]
                search_google_places, # [cite: 47, 48]
                key=f"start_location_search_{gen}",
                placeholder="Type custom starting address..."
            ) # [cite: 48]
        else:
            start_loc = COMMON_LOCATIONS[selected_shortcut] # [cite: 48]
            st.info(f"**Using:** {start_loc}") # [cite: 48]

    if "num_stops" not in st.session_state:
        st.session_state.num_stops = 0 # [cite: 48]

    with col_dest:
        st.write("**Final Destination**") # [cite: 48]
        selected_dest_shortcut = st.selectbox( # [cite: 48]
            "Quick Select Destination",
            options=list(COMMON_LOCATIONS.keys()), # [cite: 48]
            key=f"dest_shortcut_{gen}", # [cite: 48, 49]
            label_visibility="collapsed"
        ) # [cite: 49]

        if selected_dest_shortcut == "Custom / Type Address...": # [cite: 49]
            dest_loc = st_searchbox( # [cite: 49]
                search_google_places,
                key=f"destination_search_{gen}",
                placeholder="Type final destination address..."
            ) # [cite: 49]
        else:
            dest_loc = COMMON_LOCATIONS[selected_dest_shortcut] # [cite: 49, 50]
            st.info(f"**Using:** {dest_loc}") # [cite: 50]
            
        additional_stops = [] # [cite: 50]
        for i in range(st.session_state.num_stops): # [cite: 50]
            stop = st_searchbox( # [cite: 50]
                search_google_places,
                key=f"stop_search_{i}_{gen}",
                placeholder=f"Type stop address #{i+1}..."
            ) # [cite: 50]
            if stop: # [cite: 50]
                additional_stops.append(stop) # [cite: 51]
                
        c_add, c_rem = st.columns(2) # [cite: 51]
        with c_add:
            if st.button("✚ Add Stop", key=f"add_stop_btn_{gen}", use_container_width=True): # [cite: 51]
                st.session_state.num_stops += 1 # [cite: 51]
                st.rerun() # [cite: 51]
        with c_rem:
            if st.button("▬ Remove Stop", key=f"rem_stop_btn_{gen}", use_container_width=True) and st.session_state.num_stops > 0: # [cite: 51, 52]
                st.session_state.num_stops -= 1 # [cite: 52]
                st.rerun() # [cite: 52]

    if st.session_state.template_type == "at_promise": # [cite: 52]
        col_purpose, col_calc = st.columns([3, 1]) # [cite: 52]
        form_prog_code = "" # [cite: 52]
    else: 
        col_purpose, col_prog_code, col_rt = st.columns([2, 1, 1]) # [cite: 52]
        
        with col_purpose:
            purpose = st.text_input("Purpose of Travel", key=f"journey_purpose_{gen}") # [cite: 52]
        
        with col_prog_code:
            program_code = st.text_input( # [cite: 52]
                "Program Code", # [cite: 52, 53]
                placeholder="e.g., 101",
                key=f"journey_prog_code_{gen}"
            ) # [cite: 53]
        
        with col_rt:
            round_trip = st.selectbox("Round Trip?", ["Yes", "No"], key=f"journey_round_trip_{gen}") # [cite: 53]

    if st.session_state.template_type == "at_promise": # [cite: 53]
        st.markdown("##### Odometer Count (Probation Form ONLY) ") # [cite: 53]
        col_odo_start, col_odo_end = st.columns(2) # [cite: 53]

        with col_odo_start:
            odo_start_input = st.text_input( # [cite: 53, 54]
                "Odometer Start",
                placeholder="e.g., 45100",
                key=f"journey_odo_start_{gen}"
            ) # [cite: 54]

        with col_odo_end:
            odo_end_input = st.text_input( # [cite: 54]
                "Odometer End",
                placeholder="e.g., 45125",
                key=f"journey_odo_end_{gen}" # [cite: 54, 55]
            )
    else:
        odo_start_input = "" # [cite: 55]
        odo_end_input = "" # [cite: 55]

    submit_button = st.button("Calculate & Add Entry", type="primary", key=f"journey_submit_btn_{gen}", use_container_width=True) # [cite: 55]

    if submit_button: # [cite: 55]
        if not start_loc or not dest_loc: # [cite: 55]
            st.error("⚠️ Please provide both a Starting Location and Destination.") # [cite: 55]
        else:
            current_origin = start_loc # [cite: 55]
            google_miles = 0.0 # [cite: 55]
            
            for stop in additional_stops: # [cite: 55]
                google_miles += get_google_distance_miles(current_origin, stop) # [cite: 56]
                current_origin = stop # [cite: 56]
                
            google_miles += get_google_distance_miles(current_origin, dest_loc) # [cite: 56]
            calculated_miles = google_miles * 2 if round_trip == "Yes" else google_miles # [cite: 56]
            calculated_miles = round(calculated_miles, 1) # [cite: 56]

            if additional_stops: # [cite: 56]
                stops_str = " -> ".join(additional_stops) # [cite: 56, 57]
                combined_destination = f"{stops_str} -> {dest_loc}" # [cite: 57]
            else:
                combined_destination = dest_loc # [cite: 57]
                
            if round_trip == "Yes": # [cite: 57]
                combined_destination = f"{combined_destination} (RT)" # [cite: 57]

            def parse_to_float(val): # [cite: 57]
                try:
                    return round(float(val.strip()), 1) # [cite: 57, 58]
                except ValueError:
                    return None # [cite: 58]

            o_start = parse_to_float(odo_start_input) if odo_start_input.strip() else None # [cite: 58]
            o_end = parse_to_float(odo_end_input) if odo_end_input.strip() else None # [cite: 58]
            
            if o_start is not None and o_end is None: # [cite: 58]
                o_end = round(o_start + calculated_miles, 1) # [cite: 59]
            elif o_end is not None and o_start is None: # [cite: 59]
                o_start = round(o_end - calculated_miles, 1) # [cite: 59]
            elif o_start is not None and o_end is not None: # [cite: 59]
                calculated_miles = round(o_end - o_start, 1) # [cite: 59]
            else:
                o_start, o_end = 0.0, round(calculated_miles, 1) # [cite: 59, 60]

            new_entry = {
                "Date": travel_date.strftime("%Y-%m-%d") if isinstance(travel_date, (datetime.date, datetime.datetime)) else str(travel_date), # [cite: 60]
                "Starting Location": start_loc, # [cite: 60]
                "Destination": combined_destination, # [cite: 60]
                "Round Trip": round_trip, # [cite: 60]
                "Purpose of Travel": purpose, # [cite: 60]
                "Odometer Start": o_start, # [cite: 60, 61]
                "Odometer End": o_end, # [cite: 61]
                "Calculated Mileage": calculated_miles, # [cite: 61]
                "Program Code": program_code, # [cite: 61]
                "_source_file": "manual_entry" # [cite: 61]
            }
            st.session_state.mileage_data = pd.concat( # [cite: 61]
                [st.session_state.mileage_data, pd.DataFrame([new_entry])], # [cite: 61]
                ignore_index=True # [cite: 61, 62]
            )
            
            st.session_state.form_generation += 1 # [cite: 62]
            st.session_state.num_stops = 0 # [cite: 62]
            st.success(f"✅ Added! Distance: {google_miles} miles") # [cite: 62, 63]
            st.rerun() # [cite: 63]

    st.markdown("---")

    # Map details and print view for the current path # [cite: 63]
    if not st.session_state.mileage_data.empty: # [cite: 63]
        new_app_entries = st.session_state.mileage_data[ # [cite: 63]
            ~st.session_state.mileage_data["Starting Location"].str.contains(IMPORT_MARKER, na=False) # [cite: 63, 64]
        ]

        if not new_app_entries.empty: # [cite: 64]
            last_entry = new_app_entries.iloc[-1] # [cite: 64]
            origin = last_entry["Starting Location"] # [cite: 64]
            raw_destination = str(last_entry["Destination"]) # [cite: 64]
            clean_destination_chain = raw_destination.replace(" (RT)", "") # [cite: 64]
            all_dest_legs = [leg.strip() for leg in clean_destination_chain.split(" -> ")] # [cite: 64]
            final_destination = all_dest_legs[-1] # [cite: 64]
            intermediate_waypoints = all_dest_legs[:-1] if len(all_dest_legs) > 1 else [] # [cite: 64]
            trip_miles = last_entry["Calculated Mileage"] # [cite: 64]
            is_round_trip = last_entry["Round Trip"] == "Yes" # [cite: 64]
            entry_date = last_entry["Date"] # [cite: 64]
            entry_purpose = last_entry["Purpose of Travel"] # [cite: 64]

            st.subheader("Current Route Detail") # [cite: 64]
            with st.container(border=True): # [cite: 64, 65]
                st.markdown(f"**Date of Travel:** `{entry_date}` | **Purpose:** {entry_purpose}") # [cite: 65, 66]
                timeline_steps = [f"**{origin}** (Start)->"] # [cite: 66]
                for wp in intermediate_waypoints: # [cite: 66]
                    timeline_steps.append(f"`Stop: {wp}`") # [cite: 66]
                timeline_steps.append(f" **{final_destination}** (Destination) | ") # [cite: 66]
                if is_round_trip: # [cite: 66]
                    timeline_steps.append(f" **{origin}** (RT)") # [cite: 66, 67]
                st.markdown(" ".join(timeline_steps)) # [cite: 67]

            col_metric_dist, col_metric_type, col_metric_status = st.columns(3) # [cite: 67]
            with col_metric_dist: # [cite: 67]
                st.metric(label="Total Distance", value=f"{trip_miles} mi") # [cite: 67]
            with col_metric_type: # [cite: 67]
                st.metric(label="Trip Type", value="Round Trip" if is_round_trip else "One Way") # [cite: 67]
            with col_metric_status: # [cite: 67]
                st.metric( # [cite: 67]
                    label="Final Stop", # [cite: 67, 68]
                    value=final_destination if not is_round_trip else "Returned Back", # [cite: 68]
                    help="The primary destination point of this recorded log." # [cite: 68]
                )

            encoded_origin = urllib.parse.quote_plus(origin) # [cite: 68]
            encoded_final_destination = urllib.parse.quote_plus(final_destination) # [cite: 68]
            maps_url_legs = [encoded_origin] + [urllib.parse.quote_plus(wp) for wp in intermediate_waypoints] + [encoded_final_destination] # [cite: 68, 69]
            if is_round_trip: # [cite: 69]
                maps_url_legs.append(encoded_origin) # [cite: 69]
                
            direct_maps_url = f"https://www.google.com/maps/dir/{'/'.join(maps_url_legs)}/" # [cite: 69]

            if gmaps: # [cite: 69]
                embed_waypoints = list(intermediate_waypoints) # [cite: 69]
                if is_round_trip: # [cite: 69]
                    embed_waypoints.append(final_destination) # [cite: 69, 70]
                    embed_destination_target = origin # [cite: 70]
                else:
                    embed_destination_target = final_destination # [cite: 70]
                    
                encoded_embed_waypoints = "|".join([urllib.parse.quote_plus(wp) for wp in embed_waypoints]) # [cite: 70]
            
                map_url = ( # [cite: 70, 71]
                    f"https://www.google.com/maps/embed/v1/directions" # [cite: 71]
                    f"?key={api_key}" # [cite: 71]
                    f"&origin={encoded_origin}" # [cite: 71]
                    f"&destination={urllib.parse.quote_plus(embed_destination_target)}" # [cite: 71]
                )
                if encoded_embed_waypoints: # [cite: 71]
                    map_url += f"&waypoints={encoded_embed_waypoints}" # [cite: 71, 72]
                st.components.v1.iframe(map_url, width=900, height=500) # [cite: 72]
            else:
                st.warning("Please add a valid Google Maps API Key.") # [cite: 72]
                
            st.caption("**Check the miles:** If the map above shows a different total than the calculation, adjust it below before launching or printing your maps.") # [cite: 72]
        
            last_entry_index = new_app_entries.index[-1] # [cite: 72, 73]
            
            adjusted_miles = st.number_input( # [cite: 73]
                "Confirmed Logged Mileage:",
                min_value=0.0,
                value=float(trip_miles),
                step=0.1,
                format="%.1f",
                key=f"map_override_{last_entry_index}" # [cite: 73, 74]
            )
            if adjusted_miles != float(trip_miles): # [cite: 74]
                st.session_state.mileage_data.at[last_entry_index, "Calculated Mileage"] = round(adjusted_miles, 1) # [cite: 74]
                
                raw_start = st.session_state.mileage_data.at[last_entry_index, "Odometer Start"] # [cite: 74]
                raw_end = st.session_state.mileage_data.at[last_entry_index, "Odometer End"] # [cite: 74, 75]
                
                odo_start = float(raw_start or 0.0) # [cite: 75]
                odo_end = float(raw_end or 0.0) # [cite: 75]
                
                start_is_whole = odo_start.is_integer() # [cite: 75, 76]
                end_is_whole = odo_end.is_integer() # [cite: 76]
                
                if start_is_whole and not end_is_whole: # [cite: 76]
                    st.session_state.mileage_data.at[last_entry_index, "Odometer End"] = round(odo_start + adjusted_miles, 1) # [cite: 76, 77]
                elif end_is_whole and not start_is_whole: # [cite: 77]
                    st.session_state.mileage_data.at[last_entry_index, "Odometer Start"] = round(max(0.0, odo_end - adjusted_miles), 1) # [cite: 77, 78]
                else:
                    st.session_state.mileage_data.at[last_entry_index, "Odometer End"] = round(odo_start + adjusted_miles, 1) # [cite: 78]
                 
                st.rerun() # [cite: 78, 79]
            
            st.markdown("##### Actions") # [cite: 79]
            col_copy, col_print = st.columns(2) # [cite: 79]

            with col_copy:
                raw_name = st.session_state.get("employee_name", "") # [cite: 79]
                name_parts = raw_name.split() # [cite: 79, 80]
                initials = "".join([part[0].upper() for part in name_parts if part]) # [cite: 80]
                initials_suffix = f" | Initials: {initials}" if initials else "" # [cite: 80, 81]
                text_to_copy = f"Date: {entry_date} | Purpose: {entry_purpose} | Miles: {adjusted_miles}{initials_suffix}" # [cite: 81, 82]
                st.code(text_to_copy, language="text") # [cite: 82]
                st.caption("📋 Click the copy icon to save trip details") # [cite: 82]
                    
            with col_print:
                st.link_button( # [cite: 82]
                    "🔗 Open in Google Maps",
                    direct_maps_url, # [cite: 82, 83]
                    type="primary",
                    use_container_width=True
                )
                st.caption("press **Ctrl+P** to print out the Map Route") # [cite: 83]
        else:
            st.info("Sheet template journeys are uploaded. Add a new manual entry above to view the map.") # [cite: 83, 84]
    else:
        st.write("Add an entry above to generate a live map route.") # [cite: 84]

# ==========================================
# STEP 4: MILEAGE LOG TABLE & EXPORT BACK TO EXCEL
# ==========================================
elif current_step == 4:
    st.header("Review & Export")

    if not st.session_state.mileage_data.empty: # [cite: 84]
        # Running tally section for reassurance
        total_miles = st.session_state.mileage_data["Calculated Mileage"].sum() # [cite: 84]
        reimbursement_amt = total_miles * st.session_state.rate_per_mile
        
        col_m1, col_m2, col_m3 = st.columns([1, 1, 1])
        with col_m1:
            st.metric("Total Period Mileage", f"{total_miles:.1f} miles") # [cite: 84]
        with col_m2:
            st.metric("Reimbursement Rate", f"${st.session_state.rate_per_mile:.3f} / mi")
        with col_m3:
            st.metric("Estimated Reimbursement", f"${reimbursement_amt:.2f}")

        st.markdown("---")
        st.subheader("Edit/Verify Logs")
        st.caption("**Tip:** Double-click any calculated mileage cell to make manual corrections directly in the sheet table below.") # [cite: 84]

        if st.session_state.template_type == "at_promise": # [cite: 84]
            display_columns = [col for col in MILEAGE_COLUMNS if col != "Program Code"] # [cite: 85]
            column_configuration = {
                "Calculated Mileage": st.column_config.NumberColumn(
                    "Calculated Mileage",
                    help="Double-click to override with manual app miles if needed.",
                    format="%.1f",
                    min_value=0.0, # [cite: 85, 86]
                    required=True
                )
            }
        else:
            display_columns = [col for col in MILEAGE_COLUMNS if col not in ["Odometer Start", "Odometer End"]] # [cite: 86]
            column_configuration = {
                "Calculated Mileage": st.column_config.NumberColumn(
                    "Calculated Mileage", # [cite: 86, 87]
                    help="Double-click to override with manual app miles if needed.",
                    format="%.1f",
                    min_value=0.0,
                    required=True
                )
            } # [cite: 88]
        
        edited_df = st.data_editor( # [cite: 88]
            st.session_state.mileage_data,
            column_order=display_columns, 
            num_rows="dynamic",
            use_container_width=True,
            key="mileage_editor",
            column_config=column_configuration
        )

        if not edited_df.equals(st.session_state.mileage_data): # [cite: 88]
            for idx in edited_df.index: # [cite: 88]
                if idx in st.session_state.mileage_data.index: # [cite: 88, 89]
                    old_mileage = st.session_state.mileage_data.loc[idx, "Calculated Mileage"] # [cite: 89]
                    new_mileage = edited_df.loc[idx, "Calculated Mileage"] # [cite: 89]
                    
                    if old_mileage != new_mileage: # [cite: 89]
                        odo_start = float(edited_df.loc[idx, "Odometer Start"] or 0.0) # [cite: 89, 90]
                        odo_end = float(edited_df.loc[idx, "Odometer End"] or 0.0) # [cite: 90]
                        
                        start_has_decimal = "." in str(odo_start) and not str(odo_start).endswith(".0") # [cite: 90, 91]
                        end_has_decimal = "." in str(odo_end) and not str(odo_end).endswith(".0") # [cite: 91]
                        
                        if start_has_decimal and not end_has_decimal: # [cite: 91]
                            edited_df.loc[idx, "Odometer End"] = round(odo_start + new_mileage, 1) # [cite: 91, 92]
                        elif end_has_decimal and not start_has_decimal: # [cite: 92]
                            edited_df.loc[idx, "Odometer Start"] = round(max(0.0, odo_end - new_mileage), 1) # [cite: 92, 93]
                        else:
                            edited_df.loc[idx, "Odometer End"] = round(odo_start + new_mileage, 1) # [cite: 93, 94]
                            
            st.session_state.mileage_data = edited_df # [cite: 94]
            st.rerun() # [cite: 94]
    else:
        st.info("No mileage entries added yet. Go back to Step 3 or upload a template in Step 1 to get started.") # [cite: 94, 95]

    st.markdown("---")

    # Excel Exporter Section # [cite: 95]
    if st.session_state.uploaded_files_registry: # [cite: 95]
        st.subheader("Export Back to Excel Templates") # [cite: 95]
        st.markdown("This will dynamically append only new manual entries into your clean target workbook files while retaining original styling layouts.") # [cite: 95]
        
        if "_source_file" in st.session_state.mileage_data.columns: # [cite: 95]
            new_session_rows = st.session_state.mileage_data[ # [cite: 95]
                st.session_state.mileage_data["_source_file"] == "manual_entry" # [cite: 95]
            ]
        else:
            new_session_rows = st.session_state.mileage_data[ # [cite: 95, 96]
                ~st.session_state.mileage_data["Starting Location"].str.contains(IMPORT_MARKER, na=False) # [cite: 96]
            ]
        
        for filename, meta in st.session_state.uploaded_files_registry.items(): # [cite: 96]
            is_probation = (meta["template_type"] == "at_promise") # [cite: 96]
            template_display = "AT-PROMISE (Probation)" if is_probation else "Standard (GP)" # [cite: 96, 97]
            form_label = "Probation" if is_probation else "GP" # [cite: 97]
            
            with st.expander(f"📥 Download Config: {filename} ({template_display})", expanded=True): # [cite: 97]
                if st.button(f"Generate Updated File for {filename}", key=f"gen_btn_{filename}"): # [cite: 97]
                    try:
                        output_wb = openpyxl.load_workbook(BytesIO(meta["bytes"])) # [cite: 97, 98]
                        s1 = output_wb.worksheets[0] # [cite: 98]
                        
                        if is_probation: # [cite: 98]
                            s1["C3"] = st.session_state.employee_name # [cite: 98, 99]
                            s1["E4"] = st.session_state.date_range_str # [cite: 99]
                            s1["E3"] = st.session_state.rate_per_mile # [cite: 99]
                            
                            current_write_row = 9 + meta["imported_count"] # [cite: 99, 100]
                            for _, row in new_session_rows.iterrows(): # [cite: 100]
                                s1[f"B{current_write_row}"] = row["Date"] # [cite: 100]
                                s1[f"C{current_write_row}"] = row["Starting Location"] # [cite: 100, 101]
                                s1[f"D{current_write_row}"] = row["Destination"] # [cite: 101]
                                s1[f"E{current_write_row}"] = row["Purpose of Travel"] # [cite: 101]
                                s1[f"F{current_write_row}"] = row["Odometer Start"] # [cite: 101]
                                s1[f"G{current_write_row}"] = row["Odometer End"] # [cite: 101, 102]
                                current_write_row += 1 # [cite: 102]
                        else:
                            s1["D11"] = st.session_state.employee_name # [cite: 102, 103]
                            s1["D15"] = st.session_state.date_range_str # [cite: 103]

                            if len(output_wb.worksheets) >= 3: # [cite: 103]
                                s3 = output_wb.worksheets[2] # [cite: 103]
                                current_write_row = 5 # [cite: 103, 104]
                                while s3[f"B{current_write_row}"].value: # [cite: 104]
                                    current_write_row += 1 # [cite: 104]
                                
                                for _, row in new_session_rows.iterrows(): # [cite: 104, 105]
                                    s3[f"B{current_write_row}"] = row["Date"] # [cite: 105]
                                    s3[f"C{current_write_row}"] = row["Destination"] # [cite: 105, 106]
                                    s3[f"D{current_write_row}"] = row.get("Program Code", "") # [cite: 106]
                                    s3[f"E{current_write_row}"] = row["Purpose of Travel"] # [cite: 106]
                                    s3[f"F{current_write_row}"] = row["Calculated Mileage"] # [cite: 106, 107]
                                    current_write_row += 1 # [cite: 107]
                        
                        excel_stream = BytesIO() # [cite: 107]
                        output_wb.save(excel_stream) # [cite: 107, 108]
                        excel_stream.seek(0) # [cite: 108]
                        
                        today_str = datetime.date.today().strftime("%Y-%m-%d") # [cite: 108]

                        st.download_button( # [cite: 108]
                            label=f"📥 Download Updated {filename}", # [cite: 108, 109]
                            data=excel_stream, # [cite: 109]
                            file_name=f"{form_label}_Mileage_{today_str}.xlsx", # [cite: 109]
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", # [cite: 109]
                            key=f"dl_btn_{filename}", # [cite: 109, 110]
                            use_container_width=True # [cite: 110]
                        )
                    except Exception as e: # [cite: 110]
                        st.error(f"Failed to append records to target workbook {filename}: {e}") # [cite: 110, 111]
    else:
        st.info("Upload excel sheets in Step 1 to enable spreadsheet exports.") # [cite: 111]

    # --- NAVIGATION BUTTON
    st.markdown("---")
    st.button("⬅️ Back to Logger", on_click=go_to_step, args=(3,), use_container_width=True)

