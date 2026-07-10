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
DEFAULT_RATE_PER_MILE = 0.725

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
        "rate_per_mile": DEFAULT_RATE_PER_MILE,
        "template_type": "standard",  # "standard" or "at_promise"
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# --- UI: HEADER ---
st.title("🚗 Company Mileage Tracker")
st.markdown("---")

# --- UI: FILE UPLOAD & INGESTION ---
st.header("📂 Excel Template Upload")
uploaded_template = st.file_uploader("Upload your mileage workbook (.xlsx)", type=["xlsx"])

if uploaded_template and st.session_state.uploaded_file_bytes is None:
    st.session_state.uploaded_file_bytes = uploaded_template.getvalue()
    
    try:
        wb = openpyxl.load_workbook(BytesIO(st.session_state.uploaded_file_bytes), data_only=True)
        
        # Detect template type by checking cell references
        sheet1 = wb.worksheets[0]
        is_at_promise = False
        
        # Check if cells match AT-PROMISE format (C3, E3, E4) vs Standard format (D11, D15)
        at_promise_check = sheet1["C3"].value or sheet1["E3"].value or sheet1["E4"].value
        standard_check = sheet1["D11"].value or sheet1["D15"].value
        
        if at_promise_check and not standard_check:
            is_at_promise = True
            st.session_state.template_type = "at_promise"
            
            # Extract AT-PROMISE metadata
            st.session_state.employee_name = str(sheet1["C3"].value or "").strip()
            st.session_state.date_range_str = str(sheet1["E4"].value or "").strip()
            
            try:
                st.session_state.rate_per_mile = float(sheet1["E3"].value or DEFAULT_RATE_PER_MILE)
            except (ValueError, TypeError):
                st.session_state.rate_per_mile = DEFAULT_RATE_PER_MILE
            
            # Extract AT-PROMISE journey entries (starting at row 9)
            existing_rows = []
            row_idx = 9
            
            while sheet1[f"B{row_idx}"].value:
                raw_date = sheet1[f"B{row_idx}"].value
                if isinstance(raw_date, (datetime.date, datetime.datetime)):
                    formatted_date = raw_date.strftime("%Y-%m-%d")
                else:
                    formatted_date = str(raw_date)[:10]
                
                try:
                    odo_start = float(sheet1[f"F{row_idx}"].value or 0.0)
                    odo_end = float(sheet1[f"G{row_idx}"].value or 0.0)
                    calc_mileage = abs(odo_end - odo_start) if (odo_start and odo_end) else 0.0
                except (ValueError, TypeError):
                    odo_start, odo_end = 0.0, 0.0
                    calc_mileage = 0.0
                
                existing_rows.append({
                    "Date": formatted_date,
                    "Starting Location": IMPORT_MARKER,
                    "Destination": str(sheet1[f"D{row_idx}"].value or ""),
                    "Round Trip": "Yes",
                    "Purpose of Travel": str(sheet1[f"E{row_idx}"].value or ""),
                    "Odometer Start": odo_start,
                    "Odometer End": odo_end,
                    "Calculated Mileage": round(calc_mileage, 1),
                    "Program Code": ""
                })
                row_idx += 1
            
            if existing_rows:
                st.session_state.mileage_data = pd.DataFrame(existing_rows)
                st.toast(f"✅ Imported {len(existing_rows)} journey records from AT-PROMISE template", icon="📥")
        else:
            # Standard template format
            st.session_state.template_type = "standard"
            st.session_state.employee_name = str(sheet1["D11"].value or "").strip()
            st.session_state.date_range_str = str(sheet1["D15"].value or "").strip()
            
            # Extract Sheet 3 journey entries (standard format)
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
                    st.toast(f"✅ Imported {len(existing_rows)} journey records from Standard template", icon="📥")
        
        st.rerun()
    except Exception as e:
        st.error(f"Failed to parse workbook: {e}")

st.markdown("---")

# --- UI: COVER SHEET INFO ---
st.header("📋 Cover Sheet Information")
col1, col2, col3 = st.columns([2, 2, 1])

with col1:
    employee_name = st.text_input(
        "Employee Name",
        value=st.session_state.employee_name,
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
    
    # Handle date range parsing safely
    try:
        current_month_name = st.session_state.date_range_str.split(" ")[0]
        default_idx = months.index(current_month_name) if current_month_name in months else 0
    except (IndexError, ValueError):
        default_idx = 0

    selected_month = st.selectbox(
        "Select Reporting Month (2026)",
        options=months,
        index=default_idx,
        key="cs_month_select"
    )
    
    days_in_month = month_days[selected_month]
    computed_range = f"{selected_month} 1 - {selected_month} {days_in_month}, 2026"
    
    st.session_state.date_range_str = computed_range
    st.caption(f"📅 **Formatted Range:** {computed_range}")

with col3:
    rate_per_mile = st.number_input(
        "Rate per Mile ($)",
        value=st.session_state.rate_per_mile,
        min_value=0.0,
        step=0.01,
        format="%.3f",
        key="cs_rate_per_mile"
    )
    st.session_state.rate_per_mile = rate_per_mile

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
    
if "num_stops" not in st.session_state:
    st.session_state.num_stops = 0

with col_dest:
    st.write("**Destinations / Stops**")
    
    dest_loc = st_searchbox(
        search_google_places,
        key="destination_search",
        placeholder="Type final destination address..."
    )
    
    additional_stops = []
    for i in range(st.session_state.num_stops):
        stop = st_searchbox(
            search_google_places,
            key=f"stop_search_{i}",
            placeholder=f"Type intermediate stop #{i+1}..."
        )
        if stop:
            additional_stops.append(stop)
            
    c_add, c_rem = st.columns(2)
    with c_add:
        if st.button("➕ Add Stop", key="add_stop_btn", use_container_width=True):
            st.session_state.num_stops += 1
            st.rerun()
    with c_rem:
        if st.button("➖ Remove Stop", key="rem_stop_btn", use_container_width=True) and st.session_state.num_stops > 0:
            st.session_state.num_stops -= 1
            st.rerun()

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
            "Date": travel_date.strftime("%Y-%m-%d"),
            "Starting Location": start_loc,
            "Destination": combined_destination,
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
    total_reimbursement = total_miles * st.session_state.rate_per_mile
    
    col_metric1, col_metric2, col_spacer = st.columns([1, 1, 2])
    with col_metric1:
        st.metric("Total Period Mileage", f"{total_miles} miles")
    with col_metric2:
        st.metric("Total Reimbursement", f"${total_reimbursement:.2f}")
    
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
    template_name = "AT-PROMISE" if st.session_state.template_type == "at_promise" else "Standard"
    st.markdown(f"Updates `Sheet 1` and `Sheet 3` ({template_name} format) while preserving formatting.")
    
    if st.button("Generate Updated Excel Document", type="secondary", use_container_width=True):
        try:
            output_wb = openpyxl.load_workbook(BytesIO(st.session_state.uploaded_file_bytes))
            
            # Update based on template type
            s1 = output_wb.worksheets[0]
            
            if st.session_state.template_type == "at_promise":
                # AT-PROMISE template updates
                s1["C3"] = st.session_state.employee_name
                s1["E4"] = st.session_state.date_range_str
                s1["E3"] = st.session_state.rate_per_mile
                
                # Write journey entries starting at row 9
                current_write_row = 9
                
                # Clear existing entries first
                for row_idx in range(9, 100):
                    s1[f"B{row_idx}"].value = None
                    s1[f"D{row_idx}"].value = None
                    s1[f"E{row_idx}"].value = None
                    s1[f"F{row_idx}"].value = None
                    s1[f"G{row_idx}"].value = None
                
                # Write only new entries (exclude imported ones)
                new_session_rows = st.session_state.mileage_data[
                    st.session_state.mileage_data["Starting Location"] != IMPORT_MARKER
                ]
                
                for _, row in new_session_rows.iterrows():
                    s1[f"B{current_write_row}"] = row["Date"]
                    s1[f"D{current_write_row}"] = row["Destination"]
                    s1[f"E{current_write_row}"] = row["Purpose of Travel"]
                    s1[f"F{current_write_row}"] = row["Odometer Start"]
                    s1[f"G{current_write_row}"] = row["Odometer End"]
                    current_write_row += 1
            else:
                # Standard template updates
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
        
        visual_chain = " ➡️ ".join(all_dest_legs)
        if is_round_trip:
            route_label = f"{origin} ➡️ {visual_chain} 🔄 {origin}"
            mileage_label = f"{trip_miles} miles (Round Trip)"
        else:
            route_label = f"{origin} ➡️ {visual_chain}"
            mileage_label = f"{trip_miles} miles (One Way)"
        
        st.subheader("Current Route Detail")
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Route", route_label)
        col_m2.metric("Distance", mileage_label)

        encoded_origin = urllib.parse.quote_plus(origin)
        encoded_final_destination = urllib.parse.quote_plus(final_destination)
        
        waypoints_joined = "|".join([urllib.parse.quote_plus(wp) for wp in intermediate_waypoints])
        
        maps_url_legs = [encoded_origin] + [urllib.parse.quote_plus(wp) for wp in intermediate_waypoints] + [encoded_final_destination]
        if is_round_trip:
            maps_url_legs.append(encoded_origin)
            
        direct_maps_url = f"https://www.google.com/maps/dir/{'/'.join(maps_url_legs)}/"
        
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
            text_to_copy = f"Date: {entry_date} | Purpose: {entry_purpose} | Miles: {trip_miles}"
            st.code(text_to_copy, language="text")
            st.caption("📋 Click the copy icon to save trip details")
            
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
            st.warning("🔑 Map preview unavailable. Please add a valid Google Maps API Key.")
    else:
        st.info("📅 Sheet template journeys are cached. Add a new manual entry above to view the map.")
else:
    st.write("Add an entry above to generate a live map route.")
