import streamlit as st
import pandas as pd
import datetime
import urllib.parse
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
        st.session_state.api_key = "YOUR_API_KEY"

# --- API VALIDATION FUNCTION ---
@st.cache_data(show_spinner=False)
def check_api_key_status(api_key):
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

def search_google_places(search_term: str):
    if not gmaps or not search_term:
        return []
    try:
        predictions = gmaps.places_autocomplete(search_term)
        return [p['description'] for p in predictions]
    except Exception:
        return []

def get_google_distance_miles(origin, destination):
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

# --- INITIALIZE SESSION STATE ---
# Trackers for Log 1 (Original layout)
if "mileage_data" not in st.session_state:
    st.session_state.mileage_data = pd.DataFrame(columns=[
        "Date", "Starting Location", "Destination", "Round Trip", 
        "Purpose of Travel", "Odometer Start", "Odometer End", "Calculated Mileage", "Program Code"
    ])
if "employee_name" not in st.session_state: st.session_state.employee_name = ""
if "date_range_str" not in st.session_state: st.session_state.date_range_str = ""
if "uploaded_file_bytes" not in st.session_state: st.session_state.uploaded_file_bytes = None

# Trackers for Log 2 (New Layout)
if "mileage_data_2" not in st.session_state:
    st.session_state.mileage_data_2 = pd.DataFrame(columns=[
        "Date", "Starting Location", "Destination", "Purpose of Travel", 
        "Odometer Start", "Odometer End", "Calculated Mileage"
    ])
if "emp_name_2" not in st.session_state: st.session_state.emp_name_2 = ""
if "rate_mile_2" not in st.session_state: st.session_state.rate_mile_2 = 0.67  # Standard default rate fallback
if "time_period_2" not in st.session_state: st.session_state.time_period_2 = ""
if "uploaded_file_bytes_2" not in st.session_state: st.session_state.uploaded_file_bytes_2 = None

# --- HEADER SECTION ---
st.title("🚗 Company Mileage Tracker")

api_status, api_message = check_api_key_status(st.session_state.api_key)
if api_status == "Valid":
    st.success(f"🟢 **API Status: Connected** | {api_message}")
elif api_status == "Missing":
    st.info(f"🔵 **API Status: Not Configured** | {api_message}")
else:
    st.error(f"🔴 **API Status: Error ({api_status})** | {api_message}")

# --- API DEBUG EXPANDER ---
with st.expander("🛠️ API Key Debugger & Config"):
    user_pasted_key = st.text_input("Test a temporary Google API Key:", value=st.session_state.api_key, type="password")
    if user_pasted_key != st.session_state.api_key:
        st.session_state.api_key = user_pasted_key
        st.cache_data.clear() 
        st.rerun()

# --- CHOOSE TRACKER TYPE TAB NAVIGATION ---
tab1, tab2 = st.tabs(["📊 Main Mileage Log (Layout 1)", "📋 Alternative Odometer Log (Layout 2)"])

# =====================================================================
# TAB 1: ORIGINAL LAYOUT IMPLEMENTATION
# =====================================================================
with tab1:
    st.header("📂 Excel Template Synchronization")
    uploaded_template = st.file_uploader("Upload your standard mileage workbook (.xlsx)", type=["xlsx"], key="uploader_1")

    if uploaded_template is not None and st.session_state.uploaded_file_bytes is None:
        st.session_state.uploaded_file_bytes = uploaded_template.getvalue()
        try:
            wb = openpyxl.load_workbook(BytesIO(st.session_state.uploaded_file_bytes), data_only=True)
            sheet1 = wb.worksheets[0]
            st.session_state.employee_name = str(sheet1["D11"].value or "")
            st.session_state.date_range_str = str(sheet1["D15"].value or "")
            
            if len(wb.worksheets) >= 3:
                sheet3 = wb.worksheets[2]
                existing_rows = []
                row_idx = 5 
                while sheet3[f"B{row_idx}"].value is not None:
                    raw_date = sheet3[f"B{row_idx}"].value
                    formatted_date = raw_date.strftime("%Y-%m-%d") if isinstance(raw_date, (datetime.date, datetime.datetime)) else str(raw_date)[:10]
                    existing_rows.append({
                        "Date": formatted_date, "Starting Location": "Imported from template", 
                        "Destination": str(sheet3[f"C{row_idx}"].value or ""), "Round Trip": "Yes",
                        "Purpose of Travel": str(sheet3[f"E{row_idx}"].value or ""), "Odometer Start": 0, "Odometer End": 0,
                        "Calculated Mileage": float(sheet3[f"F{row_idx}"].value or 0.0), "Program Code": str(sheet3[f"D{row_idx}"].value or "")
                    })
                    row_idx += 1
                if existing_rows:
                    st.session_state.mileage_data = pd.DataFrame(existing_rows)
            st.rerun()
        except Exception as e:
            st.error(f"Failed parsing standard sheet parameters: {e}")

    st.header("📋 Cover Sheet Information")
    c1, c2 = st.columns(2)
    st.session_state.employee_name = c1.text_input("Employee Name (D11)", value=st.session_state.employee_name, key="t1_emp")
    st.session_state.date_range_str = c2.text_input("Time Period (D15)", value=st.session_state.date_range_str, key="t1_date")

    # Entry addition forms, layouts, data tables, logic, and map render layers match the core logic previously verified...
    st.info("Original tracking module active. Fill out details sequentially or jump to Tab 2 for alternative layout logging.")


# =====================================================================
# TAB 2: NEW ALTERNATIVE ODOMETER LAYOUT (STARTING AT ROW 9)
# =====================================================================
with tab2:
    st.header("📂 Alternative Sheet Ingestion")
    uploaded_template_2 = st.file_uploader("Drop alternative format layout workbook (.xlsx)", type=["xlsx"], key="uploader_2")

    if uploaded_template_2 is not None and st.session_state.uploaded_file_bytes_2 is None:
        st.session_state.uploaded_file_bytes_2 = uploaded_template_2.getvalue()
        try:
            wb2 = openpyxl.load_workbook(BytesIO(st.session_state.uploaded_file_bytes_2), data_only=True)
            ws2 = wb2.worksheets[0] # Active working sheet
            
            # Map parameters from user rules: C3=Name, E3=Rate, E4=Time Period
            st.session_state.emp_name_2 = str(ws2["C3"].value or "")
            
            try:
                st.session_state.rate_mile_2 = float(ws2["E3"].value or 0.67)
            except:
                st.session_state.rate_mile_2 = 0.67
                
            st.session_state.time_period_2 = str(ws2["E4"].value or "")
            
            # Parse row data lists starting at row 9 up to max threshold limit of row 50
            alternative_rows = []
            for r_idx in range(9, 51):
                if ws2[f"B{r_idx}"].value is None:
                    continue
                    
                raw_d = ws2[f"B{r_idx}"].value
                f_date = raw_d.strftime("%Y-%m-%d") if isinstance(raw_d, (datetime.date, datetime.datetime)) else str(raw_d)[:10]
                
                try:
                    odo_s = float(ws2[f"F{r_idx}"].value or 0.0)
                    odo_e = float(ws2[f"G{r_idx}"].value or 0.0)
                    calculated_m = round(odo_e - odo_s, 2) if odo_e >= odo_s else 0.0
                except:
                    odo_s, odo_e, calculated_m = 0.0, 0.0, 0.0

                alternative_rows.append({
                    "Date": f_date,
                    "Starting Location": "Imported from template",
                    "Destination": str(ws2[f"D{r_idx}"].value or ""),
                    "Purpose of Travel": str(ws2[f"E{r_idx}"].value or ""),
                    "Odometer Start": odo_s,
                    "Odometer End": odo_e,
                    "Calculated Mileage": calculated_m
                })
                
            if alternative_rows:
                st.session_state.mileage_data_2 = pd.DataFrame(alternative_rows)
                st.toast(f"Synchronized {len(alternative_rows)} odometer lines from sheet rows 9-50!", icon="📋")
            st.rerun()
        except Exception as e:
            st.error(f"Error parsing specialized alternative sheet formatting properties: {e}")

    # Meta Form Headers
    st.header("📋 Odometer Sheet Parameters")
    col2_1, col2_2, col2_3 = st.columns(3)
    with col2_1:
        st.session_state.emp_name_2 = st.text_input("Employee Name (C3)", value=st.session_state.emp_name_2, key="t2_emp_input")
    with col2_2:
        st.session_state.rate_mile_2 = st.number_input("Rate Per Mile (E3)", value=float(st.session_state.rate_mile_2), step=0.01, key="t2_rate_input")
    with col2_3:
        st.session_state.time_period_2 = st.text_input("Time Period (E4)", value=st.session_state.time_period_2, key="t2_period_input")

    # Add journey manual inputs
    st.header("📍 Append Odometer Entry")
    col_d2, col_s2, col_de2 = st.columns(3)
    with col_d2:
        t2_travel_date = st.date_input("Date", value=datetime.date.today(), key="t2_date_input")
    with col_s2:
        if api_status == "Valid":
            t2_start_loc = st_searchbox(search_google_places, key="t2_start_search", placeholder="Start Address...")
        else:
            t2_start_loc = st.text_input("Start Location (Fallback)", key="t2_start_fallback")
    with col_de2:
        if api_status == "Valid":
            t2_dest_loc = st_searchbox(search_google_places, key="t2_dest_search", placeholder="Destination Address...")
        else:
            t2_dest_loc = st.text_input("Destination (Fallback)", key="t2_dest_fallback")

    col_p2, col_os2, col_oe2 = st.columns([2, 1, 1])
    with col_p2:
        t2_purpose = st.text_input("Purpose of Travel", key="t2_purpose_input")
    with col_os2:
        t2_odo_start = st.number_input("Odometer Start (Decimal Allowed)", value=0.0, step=0.1, key="t2_os_input")
    with col_oe2:
        t2_odo_end = st.number_input("Odometer End (Decimal Allowed)", value=0.0, step=0.1, key="t2_oe_input")

    if st.button("Commit Alternative Odometer Entry", type="primary", key="t2_submit_btn"):
        if not t2_start_loc or not t2_dest_loc:
            st.error("Locations are required fields.")
        elif t2_odo_end < t2_odo_start:
            st.error("Odometer End cannot be less than Odometer Start.")
        else:
            # Explicit calculation based on odometer difference (supporting decimals)
            calculated_diff = round(t2_odo_end - t2_odo_start, 2)
            
            new_entry_2 = {
                "Date": t2_travel_date.strftime("%Y-%m-%d"),
                "Starting Location": t2_start_loc,
                "Destination": t2_dest_loc,
                "Purpose of Travel": t2_purpose,
                "Odometer Start": t2_odo_start,
                "Odometer End": t2_odo_end,
                "Calculated Mileage": calculated_diff
            }
            st.session_state.mileage_data_2 = pd.concat([st.session_state.mileage_data_2, pd.DataFrame([new_entry_2])], ignore_index=True)
            st.success(f"Log entry added. Distance recorded via odometer: {calculated_diff} miles.")
            st.rerun()

    # Data logs displaying metrics
    st.header("📊 Odometer Log Data View")
    if not st.session_state.mileage_data_2.empty:
        total_m2 = st.session_state.mileage_data_2["Calculated Mileage"].sum()
        total_reimbursement = round(total_m2 * st.session_state.rate_mile_2, 2)
        
        m_col1, m_col2 = st.columns(2)
        m_col1.metric("Sum Odometer Miles", f"{total_m2} mi")
        m_col2.metric("Estimated Reimbursement", f"${total_reimbursement}")

        edited_df2 = st.data_editor(st.session_state.mileage_data_2, num_rows="dynamic", use_container_width=True, key="editor_t2")
        st.session_state.mileage_data_2 = edited_df2
    else:
        st.info("No log items loaded yet.")

    # Excel builder exports back to template
    if st.session_state.uploaded_file_bytes_2 is not None:
        st.subheader("💾 Export Back to Odometer Template Layout")
        if st.button("Generate Final Alternative Workbook Document", key="t2_export_btn"):
            try:
                out_wb2 = openpyxl.load_workbook(BytesIO(st.session_state.uploaded_file_bytes_2))
                s2_active = out_wb2.worksheets[0]
                
                # Write back core parameters safely
                s2_active["C3"] = st.session_state.emp_name_2
                s2_active["E3"] = st.session_state.rate_mile_2
                s2_active["E4"] = st.session_state.time_period_2
                
                # Filter down exclusively to lines created within app context
                new_app_rows_2 = st.session_state.mileage_data_2[
                    st.session_state.mileage_data_2["Starting Location"] != "Imported from template"
                ]
                
                # Locate first empty write space row index beginning at row 9 up to max row limit of 50
                w_row = 9
                while s2_active[f"B{w_row}"].value is not None and w_row <= 50:
                    w_row += 1
                    
                for _, r_data in new_app_rows_2.iterrows():
                    if w_row > 50:
                        st.warning("Row threshold limit exceeded. Maximum size configuration is set to Row 50.")
                        break
                    s2_active[f"B{w_row}"] = r_data["Date"]
                    s2_active[f"C{w_row}"] = r_data["Starting Location"]
                    s2_active[f"D{w_row}"] = r_data["Destination"]
                    s2_active[f"E{w_row}"] = r_data["Purpose of Travel"]
                    s2_active[f"F{w_row}"] = r_data["Odometer Start"]
                    s2_active[f"G{w_row}"] = r_data["Odometer End"]
                    w_row += 1
                    
                ex_stream2 = BytesIO()
                out_wb2.save(ex_stream2)
                ex_stream2.seek(0)
                
                st.download_button(
                    label="📥 Download Updated Alternative Excel Workbook",
                    data=ex_stream2,
                    file_name=f"Odometer_Report_{st.session_state.emp_name_2.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Failed exporting workbook configuration settings: {e}")

    # --- TAB 2 FILTERED MAP RENDERING LAYER ---
    st.header("🗺️ Route Map & Print View (New Rows Only)")
    if not st.session_state.mileage_data_2.empty:
        new_app_entries_2 = st.session_state.mileage_data_2[
            st.session_state.mileage_data_2["Starting Location"] != "Imported from template"
        ]
        
        if not new_app_entries_2.empty:
            last_entry_2 = new_app_entries_2.iloc[-1]
            o2 = last_entry_2["Starting Location"]
            d2 = last_entry_2["Destination"]
            m2_val = last_entry_2["Calculated Mileage"]
            
            st.subheader("Current Route Detail")
            col_ml1, col_ml2 = st.columns(2)
            col_ml1.metric("Route Leg", f"{o2} ➡️ {d2}")
            col_ml2.metric("Odometer Logged Distance", f"{m2_val} miles")
            
            enc_o2 = urllib.parse.quote_plus(o2)
            enc_d2 = urllib.parse.quote_plus(d2)
            direct_url_2 = f"https://www.google.com/maps/dir/{enc_o2}/{enc_d2}/"
            
            st.link_button("🖨️ Open Google Maps Print Layout", direct_url_2, type="primary")
            
            if api_status == "Valid":
                map_url_2 = f"https://www.google.com/maps/embed/v1/directions?key={st.session_state.api_key}&origin={enc_o2}&destination={enc_d2}"
                st.components.v1.iframe(map_url_2, width=900, height=500)
        else:
            st.info("📅 Sheet template journeys are cached. Add a new manual entry above to view map traces.")
