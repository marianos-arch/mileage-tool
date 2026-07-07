import streamlit as st
import pandas as pd
import datetime
import googlemaps # pip install googlemaps
from streamlit_searchbox import st_searchbox # pip install streamlit-searchbox

# --- PAGE CONFIG ---
st.set_page_config(page_title="Company Mileage Tracker", layout="wide")

# --- INITIALIZE GOOGLE MAPS CLIENT ---
# Enter your actual API key here. It must have the "Places API" and "Directions API" enabled.
API_KEY = "YOUR_API_KEY" 

@st.cache_resource
def get_gmaps_client(api_key):
    if api_key and api_key != "YOUR_API_KEY":
        return googlemaps.Client(key=api_key)
    return None

gmaps = get_gmaps_client(API_KEY)

# --- AUTOCOMPLETE SEARCH FUNCTION ---
def search_google_places(search_term: str):
    """Fetches address predictions from Google Places API as the user types."""
    if not gmaps or not search_term:
        return []
    
    try:
        # Call the Google Places autocomplete API
        predictions = gmaps.places_autocomplete(search_term)
        # Return a list of strings representing the suggested addresses
        return [p['description'] for p in predictions]
    except Exception as e:
        return [f"Error fetching suggestions: {str(e)}"]

# --- INITIALIZE SESSION STATE ---
if "mileage_data" not in st.session_state:
    st.session_state.mileage_data = pd.DataFrame(columns=[
        "Date", "Starting Location", "Destination", "Round Trip", 
        "Purpose of Travel", "Odometer Start", "Odometer End", "Calculated Mileage"
    ])

st.title("🚗 Company Mileage Tracker")
st.write("Input your travel details below to calculate mileage and visualize your route.")

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

# NOTE: We removed st.form because st_searchbox needs to rerun live to fetch API suggestions
col_date, col_start, col_dest = st.columns(3)

with col_date:
    travel_date = st.date_input("Date", value=today)

with col_start:
    st.write("**Starting Location**")
    if API_KEY != "YOUR_API_KEY":
        start_loc = st_searchbox(
            search_google_places,
            key="start_location_search",
            placeholder="Type starting address..."
        )
    else:
        start_loc = st.text_input("Starting Location (API Key Missing)", placeholder="e.g., 123 Main St", key="start_fallback")

with col_dest:
    st.write("**Destination**")
    if API_KEY != "YOUR_API_KEY":
        dest_loc = st_searchbox(
            search_google_places,
            key="destination_search",
            placeholder="Type destination address..."
        )
    else:
        dest_loc = st.text_input("Destination (API Key Missing)", placeholder="e.g., Client Office", key="dest_fallback")

col_purpose, col_odo_start, col_odo_end, col_rt = st.columns([2, 1, 1, 1])
with col_purpose:
    purpose = st.text_input("Purpose of Travel")
with col_odo_start:
    odo_start = st.number_input("Odometer Start", min_value=0, value=0, step=1)
with col_odo_end:
    odo_end = st.number_input("Odometer End", min_value=0, value=0, step=1)
with col_rt:
    round_trip = st.selectbox("Round Trip?", ["No", "Yes"])

# Submit action manual button since we aren't using st.form
submit_button = st.button("Add Entry", type="primary")

# --- FORM LOGIC ---
if submit_button:
    calculated_miles = odo_end - odo_start
    if calculated_miles < 0:
        st.error("Error: Odometer End cannot be less than Odometer Start.")
    elif not start_loc or not dest_loc:
        st.error("Error: Please select both a Starting Location and Destination from the dropdown suggestions.")
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
    
    if API_KEY != "YOUR_API_KEY":
        # Using Google Maps Embed API for the visual iframe route representation
        map_url = f"https://www.google.com/maps/embed/v1/directions?key={API_KEY}&origin={origin}&destination={destination}"
        st.components.v1.iframe(map_url, width=800, height=450)
    else:
        st.info("🔄 To see live autocomplete suggestions and the Google Map route, please replace `'YOUR_API_KEY'` with your actual Google Maps API token.")
        st.code(f"[Google Map Route Preview]\nFrom: {origin}\nTo: {destination}")
else:
    st.write("Add an entry above to generate the map route.")
