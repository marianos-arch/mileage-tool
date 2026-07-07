import streamlit as st
import pandas as pd
import requests
from datetime import datetime
import io

st.set_page_config(page_title="Live Mileage Tracker", page_icon="🚗", layout="wide")

st.title("🚗 Smart Mileage Tracker (Debug Mode)")

# --- Fetch Secure API Key ---
try:
    API_KEY = st.secrets["GOOGLE_MAPS_API_KEY"]
except Exception:
    st.error("🔑 Google Maps API Key not found! Please check your `.streamlit/secrets.toml` file.")
    st.stop()

# --- Custom Search Logic ---
def get_predictions(text_input):
    if not text_input or len(text_input) < 3:
        return []
    
    url = f"https://maps.googleapis.com/maps/api/place/autocomplete/json?input={text_input}&key={API_KEY}&components=country:us"
    try:
        res = requests.get(url)
        data = res.json()
        
        # If Google rejects it, print the exact reason to screen
        if data.get("status") not in ["OK", "ZERO_RESULTS"]:
            st.error(f"Google API Error Status: {data.get('status')} - {data.get('error_message', 'No message detailed')}")
            return []
            
        return [pred["description"] for pred in data.get("predictions", [])]
    except Exception as e:
        st.error(f"Network error: {e}")
        return []

def get_live_distance(origin, destination):
    url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={origin}&destinations={destination}&units=imperial&key={API_KEY}"
    try:
        response = requests.get(url).json()
        if response["status"] == "OK":
            element = response["rows"][0]["elements"][0]
            if element["status"] == "OK":
                return round(element["distance"]["value"] * 0.000621371, 1)
        return None
    except Exception:
        return None

# --- Main Interface ---
col1, col2 = st.columns([1, 1])

with col1:
    st.header("1. Enter Trip Info")
    
    # --- From Address block ---
    from_raw = st.text_input("Type Origin Address (Type at least 3 letters):", key="from_raw")
    from_options = get_predictions(from_raw)
    from_address = st.selectbox("Confirm Origin Selection:", ["-- Select a verified address --"] + from_options, key="from_select")

    st.markdown("---")

    # --- To Address block ---
    to_raw = st.text_input("Type Destination Address (Type at least 3 letters):", key="to_raw")
    to_options = get_predictions(to_raw)
    to_address = st.selectbox("Confirm Destination Selection:", ["-- Select a verified address --"] + to_options, key="to_select")

    purpose = st.text_area("Purpose of Travel")

with col2:
    st.header("2. Route Calculation")
    
    # Only calculate if a valid choice was picked from both dropdown confirmations
    if from_address != "-- Select a verified address --" and to_address != "-- Select a verified address --":
        miles = get_live_distance(from_address, to_address)
        if miles:
            st.metric("Total Mileage", f"{miles} miles")
            static_map_url = f"https://maps.googleapis.com/maps/api/staticmap?size=600x300&markers=color:red|label:A|{from_address}&markers=color:blue|label:B|{to_address}&key={API_KEY}"
            st.image(static_map_url)
        else:
            st.warning("Could not calculate distance between these exact items.")
