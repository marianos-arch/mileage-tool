import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Mileage Tracker PoC", page_icon="🚗", layout="wide")

st.title("🚗 Smart Mileage Tracker (Proof of Concept)")
st.caption("A streamlined way to log trips and update your tracking sheets instantly.")

# --- Simulated Google Maps API Function ---
def get_mock_distance(origin, destination):
    """
    Placeholder for the Google Maps Distance Matrix API.
    In production, this will make a live network request.
    """
    if not origin or not destination:
        return 0.0
    # Just a mock calculation for the PoC wrapper
    return round(float(len(origin) + len(destination)) * 1.5, 1)

# --- Sidebar: Document Upload ---
with st.sidebar:
    st.header("1. Upload Template")
    uploaded_file = st.file_uploader("Upload your mileage_template.xlsx", type=["xlsx"])
    
    st.divider()
    st.markdown("### PoC Guardrails Active")
    st.info("Data is processed entirely in-memory. No files are stored permanently on a server.")

# --- Main Interface ---
col1, col2 = st.columns([1, 1])

with col1:
    st.header("2. Trip Information")
    
    # Input Fields
    date_selected = st.date_input("Date of Travel", datetime.now())
    from_address = st.text_input("From (Origin Address)", placeholder="e.g., 123 Main St, New York, NY")
    to_address = st.text_input("To (Destination Address)", placeholder="e.g., 456 Business Rd, Boston, MA")
    purpose = st.text_area("Purpose of Travel", placeholder="e.g., Q3 Client Strategy Meeting")

with col2:
    st.header("3. Route & Calculations")
    
    if from_address and to_address:
        # Calculate mock distance
        calculated_miles = get_mock_distance(from_address, to_address)
        
        # Display Metrics
        st.metric(label="Calculated Distance", value=f"{calculated_miles} miles")
        
        # Mocking the visual "Ctrl+P" map area
        st.subheader("Route Map Preview")
        st.warning("🗺️ [Google Maps Static Image will render here in production]")
        st.caption(f"Route: From {from_address} ➔ To {to_address}")
    else:
        st.info("Enter 'From' and 'To' addresses to generate route metrics and map previews.")

st.divider()

# --- Processing & Download Section ---
st.header("4. Generate Report")

if uploaded_file is not None:
    if from_address and to_address and purpose:
        if st.button("🚀 Process & Update Sheet", type="primary"):
            try:
                # 1. Read existing Excel file into memory
                df = pd.read_excel(uploaded_file)
                
                # 2. Create the new row data
                new_data = {
                    "Date": [date_selected.strftime("%Y-%m-%d")],
                    "Purpose of Travel": [purpose],
                    "From": [from_address],
                    "To": [to_address],
                    "Miles": [calculated_miles]
                }
                new_df = pd.DataFrame(new_data)
                
                # 3. Append new row to existing data
                updated_df = pd.concat([df, new_df], ignore_index=True)
                
                # 4. Save updated dataframe back to an in-memory bytes buffer
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    updated_df.to_excel(writer, index=False)
                buffer.seek(0)
                
                st.success("🎉 Data successfully compiled into the template!")
                
                # 5. Provide Download Button for the modified Excel
                st.download_button(
                    label="📥 Download Updated Mileage Sheet",
                    data=buffer,
                    file_name=f"updated_mileage_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"An error occurred while processing the Excel file: {e}")
    else:
        st.warning("Please fill out all Trip Information fields to activate sheet generation.")
else:
    st.info("Please upload your Excel template in the sidebar to enable report updates.")
