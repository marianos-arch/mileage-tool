import streamlit as st
import pandas as pd
import requests
from datetime import datetime
import io
from streamlit_searchbox import st_searchbox

st.set_page_config(page_title="Live Mileage Tracker", page_icon="🚗", layout="wide")

st.title("🚗 Smart Mileage Tracker (Autocomplete Ready)")
st.caption("Now with live typing predictions powered by Google Places API.")

# --- Fetch Secure API Key ---
try:
    API_KEY = st.secrets["GOOGLE_MAPS_API_KEY"]
except Exception:
    st.error("🔑 Google Maps API Key not found! Please check your `.streamlit/secrets.toml` file.")
    st.stop()

# --- Live Google Places Autocomplete Function ---
def search_google_places(search_term: str):
    """Triggers on keystroke to fetch matching addresses from Google."""
    if not search_term or len(search_term) < 3:  # Only look up after 3 characters to save API credits
        return []
    
    # Restricting results to US components to make lookups more accurate
    url = f"https://maps.googleapis.com/maps/api/place/autocomplete/json?input={search_term}&key={API_KEY}&components=country:us"
    try:
        response = requests.get(url).json()
        options = [prediction["description"] for prediction in response.get("predictions", [])]
        return options
    except Exception:
        return []

# --- Live Google Distance Calculation ---
def get_live_distance(origin, destination):
    """Queries Google Distance Matrix API for exact driving miles."""
    url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={origin}&destinations={destination}&units=imperial&key={API_KEY}"
    try:
        response = requests.get(url).json()
        if response["status"] == "OK":
            element = response["rows"][0]["elements"][0]
            if element["status"] == "OK":
                meters = element["distance"]["value"]
                miles = meters * 0.000621371
                return round(miles, 1)
        return None
    except Exception:
        return None

# --- Sidebar: Document Upload ---
with st.sidebar:
    st.header("1. Upload Template")
    uploaded_file = st.file_uploader("Upload your mileage_template.xlsx", type=["xlsx"])

# --- Main Interface ---
col1, col2 = st.columns([1, 1])

with col1:
    st.header("2. Trip Information")
    date_selected = st.date_input("Date of Travel", datetime.now())
    
    # 🌟 NEW: Live Google Places Autocomplete Searchboxes
    st.markdown("**From (Origin Address)**")
    from_address = st_searchbox(
        search_google_places,
        key="origin_search",
        placeholder="Start typing origin address..."
    )
    
    st.markdown("**To (Destination Address)**")
    to_address = st_searchbox(
        search_google_places,
        key="destination_search",
        placeholder="Start typing destination address..."
    )
    
    purpose = st.text_area("Purpose of Travel", placeholder="e.g., Client meeting regarding project setup")

with col2:
    st.header("3. Route & Calculations")
    
    if from_address and to_address:
        with st.spinner("Calculating live route constraints..."):
            calculated_miles = get_live_distance(from_address, to_address)
        
        if calculated_miles is not None:
            st.metric(label="Official Google Maps Distance", value=f"{calculated_miles} miles")
            
            # Static Map Preview Generation
            st.subheader("Route Map Preview")
            static_map_url = f"https://maps.googleapis.com/maps/api/staticmap?size=600x300&markers=color:red|label:A|{from_address}&markers=color:blue|label:B|{to_address}&key={API_KEY}"
            st.image(static_map_url, caption="Live Route Bounds")
        else:
            st.error("Could not calculate distance. Please check the address format.")
    else:
        st.info("Start typing and select full addresses to ping Google Maps APIs.")

st.divider()

# --- Processing & Excel Generation ---
st.header("4. Generate Report")

if uploaded_file is not None:
    if from_address and to_address and purpose and 'calculated_miles' in locals() and calculated_miles is not None:
        if st.button("🚀 Process & Update Sheet", type="primary"):
            df = pd.read_excel(uploaded_file)
            
            new_data = {
                "Date": [date_selected.strftime("%Y-%m-%d")],
                "Purpose of Travel": [purpose],
                "From": [from_address],
                "To": [to_address],
                "Miles": [calculated_miles]
            }
            
            updated_df = pd.concat([df, pd.DataFrame(new_data)], ignore_index=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                updated_df.to_excel(writer, index=False)
            buffer.seek(0)
            
            st.success("🎉 Data accurately compiled into your template layout!")
            st.download_button(
                label="📥 Download Updated Mileage Sheet",
                data=buffer,
                file_name=f"updated_mileage_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Please fully select search predictions and enter trip details to compile the output.")
else:
    st.info("Please upload your baseline tracking Excel template file in the sidebar to proceed.")
