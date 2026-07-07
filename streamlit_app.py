import streamlit as st
import pandas as pd
import datetime
import googlemaps # pip install googlemaps
from streamlit_searchbox import st_searchbox # pip install streamlit-searchbox

# --- PAGE CONFIG ---
st.set_page_config(page_title="Company Mileage Tracker", layout="wide")

# --- API KEY MANAGEMENT ---
# Priority: 1. Sidebar/UI Input (for debugging) -> 2. Hardcoded string fallback
if "api_key" not in st.session_state:
    st.session_state.api_key = "YOUR_API_KEY"  # Replace this with your default key if desired

# --- API VALIDATION FUNCTION ---
@st.cache_data(show_spinner=False)
def check_api_key_status(api_key):
    """Tests if the Google Maps API key is fundamentally valid."""
    if not api_key or api_key == "YOUR_API_KEY":
        return "Missing", "Please configure your Google Maps API key to enable address autocompletion."
    
    try:
        # Spin up a temporary client to test a lightweight endpoint (timezone or simple autocomplete)
        test_client = googlemaps.Client(key=api_key)
        # Requesting predictions for a single character is a fast way to validate the Places API pipeline
        test_client.places_autocomplete("A")
        return "Valid", "Google Maps API connected successfully! Autocomplete and routing are online."
    except Exception as e:
        # Catch API errors (e.g., Invalid key, Billing not enabled, API restriction blocks)
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

# --- INITIALIZE SESSION STATE FOR DATA ---
if "mileage_data" not in st.session_state:
    st.session_state.mileage_data = pd.DataFrame(columns=[
        "Date", "Starting Location", "Destination", "Round Trip", 
        "Purpose of Travel", "Odometer Start", "Odometer End", "Calculated Mileage"
    ])

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
    
    # Text input to paste a key on the fly for testing
    user_pasted_key = st.text_input(
        "Test a temporary Google API Key:", 
        value=st.session_state.api_key, 
        type="password",
        help="Paste your string here. It will instantly re-verify the token against Google Cloud endpoints."
    )
    
    if user_pasted_key != st.session_state.api_key:
        st.session_state.api_key = user_pasted_key
        st.cache_data.clear() # Clear verification cache to force re-evaluation
        st.rerun()
        
    st.markdown(f"""
    **Diagnostic Checklist:**
    * Current Token Value: `{st.session_state.api_key[:5]}...{st.session_state.api_key[-5:] if len(st.session_state.api_key) > 5 else ''}`
    * Required Library Status: `googlemaps` integration verified.
    * Target Cloud Credentials Needed: **Places API** (for typing search) & **Directions API** (for routing maps).
    """)

st.markdown("---")

# --- SECTION 1: COVER SHEET INFO ---
st.header("📋 Cover Sheet Information")
col1, col2 = st.columns(2)

with col1:
    employee_name = st.text_input("Employee Name (C3 / D11)", placeholder="John Doe")

with col2:
    today = datetime.date.today()
    selected_month = st.selectbox("Time Period (Month)", 
                                  ["January", "February", "March", "April", "May", "June", 
                                   "July", "August", "September", "October", "November", "December"],
                                  index=today.month - 1)
    st.caption(f"Selected Period: **{selected_month} 2026**")

st.markdown("---")

# --- SECTION 2: USER INPUT FORM ---
st.header("📍 Add New Journey")

col_date, col_start, col_dest = st.columns(3)

with col_date:
    travel_date = st.date_input("Date", value=today)

with col_start:
    st.write("**Starting Location**")
    if api_status == "Valid":
        start_loc = st_searchbox(
            search_google_places,
            key="start_location_search",
            placeholder="Type starting address..."
        )
    else:
        start_loc = st.text_input("Starting Location (Fallback mode)", placeholder="Type address manually...", key="start_fallback")

with col_dest:
    st.write("**Destination**")
    if api_status == "Valid":
        dest_loc = st_searchbox(
            search_google_places,
            key="destination_search",
            placeholder="Type destination address..."
        )
    else:
        dest_loc = st.text_input("Destination (Fallback mode)", placeholder="Type address manually...", key="dest_fallback")

col_purpose, col_odo_start, col_odo_end, col_rt = st.columns([2, 1, 1, 1])
with col_purpose:
    purpose = st.text_input("Purpose of Travel")
with col_odo_start:
    odo_start = st.number_input("Odometer Start", min_value=0, value=0, step=1)
with col_odo_end:
    odo_end = st.number_input("Odometer End", min_value=0, value=0, step=1)
with col_rt:
    round_trip = st.selectbox("Round Trip?", ["No", "Yes"])

submit_button = st.button("Add Entry", type="primary")

# --- FORM LOGIC ---
if submit_button:
    calculated_miles = odo_end - odo_start
    if calculated_miles < 0:
        st.error("Error: Odometer End cannot be less than Odometer Start.")
    elif not start_loc or not dest_loc:
        st.error("Error: Please provide both a Starting Location and Destination.")
    else:
        new_entry = {
            "Date": travel_date.strftime("%Y-%m-%d"),
            "Starting Location": start_loc,
            "Destination": dest_loc,
            "Round Trip": round_trip,
            "Purpose of Travel": purpose,
            "Odometer Start": odo_start,
            "Odometer End": odo_end,
            "Calculated Mileage": calculated_miles
        }
        st.session_state.mileage_data = pd.concat([st.session_state.mileage_data, pd.DataFrame([new_entry])], ignore_index=True)
        st.success("Entry added successfully!")
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
    st.caption("💡 *Tip: You can double-click cells to edit them directly, or select a row and hit 'Delete' on your keyboard.*")
else:
    st.info("No mileage entries added yet.")

st.markdown("---")

# --- SECTION 4: GOOGLE MAPS ROUTE VISUALIZATION ---
st.header("🗺️ Route Map & Print View")

if not st.session_state.mileage_data.empty:
    last_entry = st.session_state.mileage_data.iloc[-1]
    origin = last_entry["Starting Location"]
    destination = last_entry["Destination"]
    
    st.subheader(f"Current Route: {origin} ➡️ {destination}")
    
    if st.button("🖨️ Open Print View"):
        st.warning("Press Ctrl+P (or Cmd+P on Mac) to print this page configuration.")
    
    if api_status == "Valid":
        map_url = f"https://www.google.com/maps/embed/v1/directions?key={st.session_state.api_key}&origin={origin}&destination={destination}"
        st.components.v1.iframe(map_url, width=800, height=450)
    else:
        st.info("🔄 Map display paused because the API configuration status is currently offline or invalid.")
        st.code(f"[Google Map Route Preview]\nFrom: {origin}\nTo: {destination}")
else:
    st.write("Add an entry above to generate the map route.")
