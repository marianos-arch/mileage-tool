import streamlit as st
import pandas as pd
import datetime
import urllib.parse
import googlemaps
from streamlit_searchbox import st_searchbox
import openpyxl
from io import BytesIO

# --- PAGE CONFIG
st.set_page_config(page_title="GP Mileage Tool", layout="wide")

# --- CONSTANTS
MILEAGE_COLUMNS = [
    "Date", "Starting Location", "Destination", "Round Trip",
    "Purpose of Travel", "Odometer Start", "Odometer End", "Calculated Mileage", "Program Code"
]

IMPORT_MARKER = "Imported from template"
METERS_TO_MILES = 0.000621371
DEFAULT_RATE_PER_MILE = 0.725

# Common locations to select from
COMMON_LOCATIONS = {
    "Custom / Type Address...": "",
    "Main Office": "1616 29th St, Bakersfield, CA 93301",
    "DEC": "1130 17th St, Bakersfield, CA 93301",
    "Delano": "1109 High St., Delano, CA 93215"
}

# --- API KEY & CLIENT INITIALIZATION
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

# --- GOOGLE PLACES AUTOCOMPLETE
def search_google_places(query_text):
    if not query_text or not gmaps:
        return []
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

# --- GOOGLE DISTANCE MATRIX API 
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

# --- HELPERS: TEMPLATE PROCESSING
def detect_and_extract_workbook(file_bytes, filename):
    """Detect template type and extract existing journey rows with header metadata."""
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
        sheet1 = wb.worksheets[0]
        # Identify template explicitly by cell landmarks
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
                
            # Extract Probation journeys (row 9+)
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
            # Standard Template Target Processing
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
                        "Date": formatted_date,
                        "Starting Location": IMPORT_MARKER,
                        "Destination": str(sheet3[f"C{row_idx}"].value or ""),
                        "Round Trip": "Yes",
                        "Purpose of Travel": str(sheet3[f"E{row_idx}"].value or ""),
                        "Odometer Start": 0,
                        "Odometer End": 0,
                        "Calculated Mileage": float(sheet3[f"F{row_idx}"].value or 0.0),
                        "Program Code": str(sheet3[f"D{row_idx}"].value or ""),
                        "_source_file": filename
                    })
                    row_idx += 1

        # Check if the extracted date string actually has content
        has_uploaded_date = bool(date_range_str and date_range_str.strip())
        
        return {
            "filename": filename,
            "bytes": file_bytes,
            "template_type": template_type,
            "employee_name": employee_name,
            "date_range_str": date_range_str,
            "has_uploaded_date": has_uploaded_date,
            "rate_per_mile": rate_per_mile,
            "imported_count": len(existing_rows),
            "rows": existing_rows
        }
    except Exception as e:
        st.error(f"Error parsing structural configuration of {filename}: {e}")
        return None

# --- SESSION STATE INITIALIZATION
def init_session_state():
    defaults = {
        "mileage_data": pd.DataFrame(columns=MILEAGE_COLUMNS + ["_source_file"]),
        "employee_name": "",
        "date_range_str": "",
        "rate_per_mile": DEFAULT_RATE_PER_MILE,
        "template_type": "standard",  
        "uploaded_files_registry": {},  
        "processed_file_hashes": set() 
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# --- UI: HEADER
st.title("GP Mileage Tracker")
st.markdown("---")

# --- UI: FILE UPLOAD & INGESTION
st.header("Mileage Excel Template Upload 🗂️")
uploaded_templates = st.file_uploader(
    "Upload your mileage excel sheet (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_templates:
    state_updated = False
    for uploaded_file in uploaded_templates:
        if uploaded_file.name not in st.session_state.processed_file_hashes:
            file_bytes = uploaded_file.getvalue()
            parsed_result = detect_and_extract_workbook(file_bytes, uploaded_file.name)
            
            if parsed_result:
                st.session_state.uploaded_files_registry[uploaded_file.name] = parsed_result
                st.session_state.processed_file_hashes.add(uploaded_file.name)
                
                st.session_state.template_type = parsed_result["template_type"]
                st.session_state.employee_name = parsed_result["employee_name"]
                st.session_state.date_range_str = parsed_result["date_range_str"]
                st.session_state.rate_per_mile = parsed_result["rate_per_mile"]
                
                if parsed_result["rows"]:
                    new_df = pd.DataFrame(parsed_result["rows"])
                    st.session_state.mileage_data = pd.concat(
                        [st.session_state.mileage_data, new_df], 
                        ignore_index=True
                    )
                st.toast(f" Imported {parsed_result['imported_count']} records from {uploaded_file.name} ({parsed_result['template_type'].upper()})")
                state_updated = True
                
    if state_updated:
        st.rerun()

st.markdown("---")

# --- UI: COVER SHEET INFO
st.header("Mileage Cover Sheet Information")
col1, col2, col3 = st.columns([2, 2, 1])

with col1:
    if st.session_state.employee_name and not st.session_state.get("cs_employee_name"):
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

    # Initialize tracking variables safely
    if "has_uploaded_date" not in st.session_state:
        st.session_state.has_uploaded_date = False

    if st.session_state.has_uploaded_date and st.session_state.date_range_str:
        st.text_input(
            "Reporting Period Range", 
            value=st.session_state.date_range_str, 
            disabled=True, 
            key="static_period_display"
        )
    else:
        # 1. Define a clean callback function to handle updates safely
        def update_month_string():
            if "cs_month_select" in st.session_state:
                picked_month = st.session_state.cs_month_select
                days = month_days[picked_month]
                st.session_state.date_range_str = f"{picked_month} 1 - {picked_month} {days}, 2026"

        # 2. Render the select box safely WITHOUT feeding it an index derived from text
        # If it doesn't exist yet, default it to the first month
        if "cs_month_select" not in st.session_state:
            st.session_state.cs_month_select = "January"
            
        selected_month = st.selectbox(
            "Select Reporting Month (2026)",
            options=months,
            key="cs_month_select",
            on_change=update_month_string
        )
        
        # 3. Fallback sync to ensure date_range_str is never blank on initial render
        if not st.session_state.get("date_range_str"):
            days = month_days[selected_month]
            st.session_state.date_range_str = f"{selected_month} 1 - {selected_month} {days}, 2026"
            
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

st.markdown("---")

# --- UI: ADD NEW JOURNEY FORM 
st.header(" Add New Mileage Entry")

if "form_generation" not in st.session_state:
    st.session_state.form_generation = 0

gen = st.session_state.form_generation
today = datetime.date.today()
col_date, col_start, col_dest = st.columns(3)

with col_date:
    travel_date = st.date_input("Date", value=today, key=f"journey_travel_date_{gen}")

with col_start:
    st.write("**Starting Location**")
    selected_shortcut = st.selectbox(
        "Quick Select Location",
        options=list(COMMON_LOCATIONS.keys()),
        key=f"start_shortcut_{gen}",
        label_visibility="collapsed"
    )
    
    if selected_shortcut == "Custom / Type Address...":
        start_loc = st_searchbox(
            search_google_places,
            key=f"start_location_search_{gen}",
            placeholder="Type custom starting address..."
        )
    else:
        start_loc = COMMON_LOCATIONS[selected_shortcut]
        st.info(f"**Using:** {start_loc}")

if "num_stops" not in st.session_state:
    st.session_state.num_stops = 0

with col_dest:
    st.write("**Final Destination**")
    selected_dest_shortcut = st.selectbox(
        "Quick Select Destination",
        options=list(COMMON_LOCATIONS.keys()),
        key=f"dest_shortcut_{gen}",
        label_visibility="collapsed"
    )

    if selected_dest_shortcut == "Custom / Type Address...":
        dest_loc = st_searchbox(
            search_google_places,
            key=f"destination_search_{gen}",
            placeholder="Type final destination address..."
        )
    else:
        dest_loc = COMMON_LOCATIONS[selected_dest_shortcut]
        st.info(f"**Using:** {dest_loc}")
        
    additional_stops = []
    for i in range(st.session_state.num_stops):
        stop = st_searchbox(
            search_google_places,
            key=f"stop_search_{i}_{gen}",
            placeholder=f"Type stop address #{i+1}..."
        )
        if stop:
            additional_stops.append(stop)
            
    c_add, c_rem = st.columns(2)
    with c_add:
        if st.button("✚ Add Stop", key=f"add_stop_btn_{gen}", use_container_width=True):
            st.session_state.num_stops += 1
            st.rerun()
    with c_rem:
        if st.button("▬ Remove Stop", key=f"rem_stop_btn_{gen}", use_container_width=True) and st.session_state.num_stops > 0:
            st.session_state.num_stops -= 1
            st.rerun()

if st.session_state.template_type == "at_promise":
    col_purpose, col_calc = st.columns([3, 1])
    form_prog_code = "" # Safe empty string baseline for the database schema
else: 
    col_purpose, col_prog_code, col_rt = st.columns([2, 1, 1])
    
    with col_purpose:
        purpose = st.text_input("Purpose of Travel", key=f"journey_purpose_{gen}")
    
    with col_prog_code:
        program_code = st.text_input(
            "Program Code",
            placeholder="e.g., 101",
            key=f"journey_prog_code_{gen}"
        )
    
    with col_rt:
        round_trip = st.selectbox("Round Trip?", ["Yes", "No"], key=f"journey_round_trip_{gen}")

if st.session_state.template_type == "at_promise":
    st.markdown("##### Odometer Count (Probabtion Form ONLY) ")
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

submit_button = st.button("Calculate & Add Entry", type="primary", key=f"journey_submit_btn_{gen}", use_container_width=True)

if submit_button:
    if not start_loc or not dest_loc:
        st.error("⚠️ Please provide both a Starting Location and Destination.")
    else:
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
            try:
                return round(float(val.strip()), 1)
            except ValueError:
                return None

        o_start = parse_to_float(odo_start_input) if odo_start_input.strip() else None
        o_end = parse_to_float(odo_end_input) if odo_end_input.strip() else None
        
        if o_start is not None and o_end is None:
            o_end = round(o_start + calculated_miles, 1)
        elif o_end is not None and o_start is None:
            o_start = round(o_end - calculated_miles, 1)
        elif o_start is not None and o_end is not None:
            calculated_miles = round(o_end - o_start, 1)
        else:
            o_start, o_end = 0.0, round(calculated_miles, 1)

        new_entry = {
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
        st.session_state.mileage_data = pd.concat(
            [st.session_state.mileage_data, pd.DataFrame([new_entry])],
            ignore_index=True
        )
        
        st.session_state.form_generation += 1
        st.session_state.num_stops = 0
        st.success(f"✅ Added! Distance: {google_miles} miles")
        st.rerun()
        
st.markdown("---")

# --- UI: ROUTE MAP & PRINT VIEW
st.header("Route Map & Print View")

if not st.session_state.mileage_data.empty:
    new_app_entries = st.session_state.mileage_data[
        ~st.session_state.mileage_data["Starting Location"].str.contains(IMPORT_MARKER, na=False)
    ]

    if not new_app_entries.empty:
        last_entry = new_app_entries.iloc[-1]
        origin = last_entry["Starting Location"]
        raw_destination = str(last_entry["Destination"])
        clean_destination_chain = raw_destination.replace(" (RT)", "")
        all_dest_legs = [leg.strip() for leg in clean_destination_chain.split(" -> ")]
        final_destination = all_dest_legs[-1]
        intermediate_waypoints = all_dest_legs[:-1] if len(all_dest_legs) > 1 else []
        trip_miles = last_entry["Calculated Mileage"]
        is_round_trip = last_entry["Round Trip"] == "Yes"
        entry_date = last_entry["Date"]
        entry_purpose = last_entry["Purpose of Travel"]

        st.subheader("Current Route Detail")
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
            st.metric(label="Total Distance", value=f"{trip_miles} mi")
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
        st.markdown("---")

        if gmaps:
            embed_waypoints = list(intermediate_waypoints)
            if is_round_trip:
                embed_waypoints.append(final_destination)
                embed_destination_target = origin
            else:
                embed_destination_target = final_destination
                
            encoded_embed_waypoints = "|".join([urllib.parse.quote_plus(wp) for wp in embed_waypoints])
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
            
        st.caption("**Check the miles:** If the map above shows a different total than the calculation, adjust it below before launching or printing your maps.")
        last_entry_index = new_app_entries.index[-1]
        
        adjusted_miles = st.number_input(
            "Confirmed Logged Mileage:",
            min_value=0.0,
            value=float(trip_miles),
            step=0.1,
            format="%.1f",
            key=f"map_override_{last_entry_index}"
        )
        if adjusted_miles != float(trip_miles):
            # 1. Update the Calculated Mileage
            st.session_state.mileage_data.at[last_entry_index, "Calculated Mileage"] = round(adjusted_miles, 1)
            
            # 2. Extract current odometer values
            raw_start = st.session_state.mileage_data.at[last_entry_index, "Odometer Start"]
            raw_end = st.session_state.mileage_data.at[last_entry_index, "Odometer End"]
            
            odo_start = float(raw_start or 0.0)
            odo_end = float(raw_end or 0.0)
            
            # 3. Determine which value is a clean whole number (ends in .0 when cast to float)
            start_is_whole = odo_start.is_integer()
            end_is_whole = odo_end.is_integer()
            
            # 4. Lock the whole number, change the decimal number
            if start_is_whole and not end_is_whole:
                # Start is a whole number -> Lock Start, Change End
                st.session_state.mileage_data.at[last_entry_index, "Odometer End"] = round(odo_start + adjusted_miles, 1)
                
            elif end_is_whole and not start_is_whole:
                # End is a whole number -> Lock End, Change Start
                st.session_state.mileage_data.at[last_entry_index, "Odometer Start"] = round(max(0.0, odo_end - adjusted_miles), 1)
                
            else:
                # Fallback: If both are whole numbers or both are decimals, default to updating the End Odometer
                st.session_state.mileage_data.at[last_entry_index, "Odometer End"] = round(odo_start + adjusted_miles, 1)
                
            st.rerun()
        
        st.markdown("##### Actions")
        col_copy, col_print = st.columns(2)

        with col_copy:
            raw_name = st.session_state.get("employee_name", "")
            # take the first letter of each word, and capitalize it
            name_parts = raw_name.split()
            initials = "".join([part[0].upper() for part in name_parts if part])
            # append it only if initials were successfully generated
            initials_suffix = f" | Initials: {initials}" if initials else ""
            text_to_copy = f"Date: {entry_date} | Purpose: {entry_purpose} | Miles: {adjusted_miles}{initials_suffix}"
            st.code(text_to_copy, language="text")
            st.caption("📋 Click the copy icon to save trip details")
                
        with col_print:
            st.link_button(
                "🔗 Open in Google Maps",
                direct_maps_url,
                type="primary",
                use_container_width=True
            )
            st.caption("press **Ctrl+P** to print out the Map Route")
        st.markdown("---")
    else:
        st.info("Sheet template journeys are uploaded. Add a new manual entry above to view the map.")
else:
    st.write("Add an entry above to generate a live map route.")

st.markdown("---")

    
# --- UI: MILEAGE LOG TABLE
st.header("Mileage Log Table")

if not st.session_state.mileage_data.empty:
    total_miles = st.session_state.mileage_data["Calculated Mileage"].sum()
    col_metric1, col_metric2, col_spacer = st.columns([1, 1, 2])
    with col_metric1:
        st.metric("Total Period Mileage", f"{total_miles:.1f} miles")
    
    st.caption("**Tip:** Double-click the **Calculated Mileage** cell to type the exact number.")

    if st.session_state.template_type == "at_promise":
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
    
    # Render the interactive data table
    edited_df = st.data_editor(
        st.session_state.mileage_data,
        column_order=display_columns, 
        num_rows="dynamic",
        use_container_width=True,
        key="mileage_editor",
        column_config=column_configuration
    )
    

    if not edited_df.equals(st.session_state.mileage_data):
        for idx in edited_df.index:
            if idx in st.session_state.mileage_data.index:
                old_mileage = st.session_state.mileage_data.loc[idx, "Calculated Mileage"]
                new_mileage = edited_df.loc[idx, "Calculated Mileage"]
                
                if old_mileage != new_mileage:
                    # convert odometer targets to floats
                    odo_start = float(edited_df.loc[idx, "Odometer Start"] or 0.0)
                    odo_end = float(edited_df.loc[idx, "Odometer End"] or 0.0)
                    
                    start_has_decimal = "." in str(odo_start) and not str(odo_start).endswith(".0")
                    end_has_decimal = "." in str(odo_end) and not str(odo_end).endswith(".0")
                    
                    if start_has_decimal and not end_has_decimal:
                        # Odometer Start contains the true data -> adjust End
                        edited_df.loc[idx, "Odometer End"] = round(odo_start + new_mileage, 1)
                    elif end_has_decimal and not start_has_decimal:
                        # Odometer End contains the true data -> adjust Start
                        edited_df.loc[idx, "Odometer Start"] = round(max(0.0, odo_end - new_mileage), 1)
                    else:
                        # Fallback default
                        edited_df.loc[idx, "Odometer End"] = round(odo_start + new_mileage, 1)
                        
        # our modified calculations safely back to memory and trigger an visual update
        st.session_state.mileage_data = edited_df
        st.rerun()
else:
    st.info("No mileage entries added yet. Use the form above to get started.")

st.markdown("---")

# --- UI: EXPORT TO EXCEL
if st.session_state.uploaded_files_registry:
    st.subheader("Export Back to Excel Templates")
    st.markdown("Updates information and adds **only new manual entries** while preserving template formatting.")
    
    # Isolate newly recorded rows safely (handles old cached DataFrames without crashing)
    if "_source_file" in st.session_state.mileage_data.columns:
        new_session_rows = st.session_state.mileage_data[
            st.session_state.mileage_data["_source_file"] == "manual_entry"
        ]
    else:
        new_session_rows = st.session_state.mileage_data[
            ~st.session_state.mileage_data["Starting Location"].str.contains(IMPORT_MARKER, na=False)
        ]
    
    # Dynamic loop through all active sheet registries
    for filename, meta in st.session_state.uploaded_files_registry.items():
        # Read directly from the file's own meta dictionary, NOT the global st.session_state
        is_probation = (meta["template_type"] == "at_promise")
        template_display = "AT-PROMISE (Probation)" if is_probation else "Standard (GP)"
        form_label = "Probation" if is_probation else "GP"
        
        with st.expander(f"Export Workbook: {filename} ({template_display})", expanded=True):
            if st.button(f"Generate Updated File for {filename}", key=f"gen_btn_{filename}"):
                try:
                    output_wb = openpyxl.load_workbook(BytesIO(meta["bytes"]))
                    s1 = output_wb.worksheets[0]
                    
                    # Route formatting exclusively based on this specific file's internal type
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
    st.info("Upload excel sheets above to enable spreadsheet exports.")





