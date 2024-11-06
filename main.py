# Import required modules
import os
import streamlit as st
import pandas as pd
import datetime
import pytz
import xlsxwriter  # To save Excel files

# Streamlit app setup
st.title("MTN Outage Report Generator")

# Step 1: Upload NOC Matrix only once
if "noc_matrix" not in st.session_state:
    # Upload button for NOC Matrix file
    noc_matrix_file = st.file_uploader("Upload NOC Matrix Excel file (one-time)", type=["xlsx"], key="noc_matrix_upload")
    if noc_matrix_file:
        # Read and store the NOC matrix in session_state
        st.session_state.noc_matrix = pd.read_excel(noc_matrix_file, sheet_name='NOC Matrix')
        st.success("NOC Matrix uploaded successfully.")
else:
    st.write("Using the previously uploaded NOC Matrix.")

# Step 2: Upload the MTN Incident report raw data each time
uploaded_file = st.file_uploader("Upload MTN Incident report raw data", type=["xlsx"])

# Ensure the file is uploaded
if uploaded_file is not None and "noc_matrix" in st.session_state:
    # Read the uploaded MTN Incident report raw data into a DataFrame
    df = pd.read_excel(uploaded_file)
    st.write("DataFrame preview of MTN Incident report raw data:")
    st.dataframe(df.head())

    # Collapsible section for filtering steps
    with st.expander("View Filtering Steps"):
        # Step 1: Filter rows where 'Fault Recovery Time' is blank
        df_step1 = df[df['Fault Recovery Time(HW Process TT_faultrecoverytime)'].isna()]
        st.write("After Step 1 (Blank Fault Recovery Time):")
        st.dataframe(df_step1.head())

        # Step 2: Filter for rows with 'Yes' in the 'Site Outage' column
        df_step2 = df_step1[df_step1['Site Outage(Create TT)'] == 'Yes']
        st.write("After Step 2 (Site Outage = Yes):")
        st.dataframe(df_step2.head())

        # Step 3: Filter for rows with 'ATC' in the 'Passive Colo' column
        df_step3 = df_step2[df_step2['Passive Colo(Create TT_passivecolo)'] == 'ATC']
        st.write("After Step 3 (Passive Colo = ATC):")
        st.dataframe(df_step3.head())

        # Step 4: Filter rows for specific alarm names
        allowed_alarms = [
            'Alternating Current Diesel Generator - Generator faulty and need to be replaced', 'BTS Down', 'CSL Fault', 
            'Heartbeat Failure', 'NE Is Disconnected', 'NodeB Down', 'NodeB is out of service', 'NodeB Unavailable', 
            'OML Fault', 'Site Abis control link broken', 'The link between the server and the ME is broken', 
            'System Undervoltage', 'Cell Unavailable', 'Temperature abnormal', 'POWER_ABNORMAL', 'TEMP_ALARM', 
            'Power Disturbance', 'Battery Discharging', 'epsEnodeBUnreachable', 'gNodeB Out of Service', 
            'Service Unavailable', 'eNodeB is out of service', 'Cell is out of service', 'HW Partial Fault', 
            'Internal fault', 'System Overvoltage', 'S1 Interface Fault', 'Power Fail(Entity)', 'S1 link is broken'
        ]
        df_step4 = df_step3[df_step3['Alarm Name(Create TT_alarmname)'].isin(allowed_alarms)]
        st.write("After Step 4 (Allowed Alarms):")
        st.dataframe(df_step4.head())

    # Final filtered DataFrame after all steps
    df_final = df_step4

    # # Rename columns for report
    df_final = df_step4.rename(columns={
        'Ticket ID': 'Ticket ID',
        'Alarm Name(Create TT_alarmname)': 'Alarm Name',
        'Site ID(Create TT)': 'Site ID',
        'Fault Last Occur Time': 'Outage start time'
    })

    # --- Merging NOC Matrix columns ---
    noc_matrix = st.session_state.noc_matrix

    # Convert merge keys to strings to ensure compatibility during merging
    df_final['Site ID'] = df_final['Site ID'].astype(str)
    noc_matrix['MTN ID'] = noc_matrix['MTN ID'].astype(str)
    noc_matrix['Site Number'] = noc_matrix['Site Number'].astype(str)
    
    # Merge with NOC Matrix to get 'ATC ID', 'ISM', 'SDS', 'Anchor Site ID'
    try:
        df_final = pd.merge(df_final, noc_matrix[['MTN ID', 'Site Number']], left_on='Site ID', right_on='MTN ID', how='left')
        df_final.rename(columns={'Site Number': 'ATC ID'}, inplace=True)

         # Ensure 'ATC ID' is in plain numeric format without commas
        df_final['ATC ID'] = df_final['ATC ID'].apply(lambda x: x if pd.isnull(x) else str(int(float(x))))

        # Merge to get 'ISM'
        df_final = pd.merge(df_final, noc_matrix[['Site Number', 'SMPMS Vendor']], left_on='ATC ID', right_on='Site Number', how='left')
        df_final.rename(columns={'SMPMS Vendor': 'ISM'}, inplace=True)

        # Merge to get 'SDS'
        df_final = pd.merge(df_final, noc_matrix[['Site Number', 'Ops Person Responsible']], left_on='ATC ID', right_on='Site Number', how='left')
        df_final.rename(columns={'Ops Person Responsible': 'SDS'}, inplace=True)

        # Merge to get 'Anchor Site ID'
        df_final = pd.merge(df_final, noc_matrix[['Site Number', 'ANCHOR SITE ID']], left_on='ATC ID', right_on='Site Number', how='left')
        df_final.rename(columns={'ANCHOR SITE ID': 'Anchor Site ID'}, inplace=True)

        # Drop unnecessary columns after merging
        df_final.drop(columns=['Site Number_x', 'Site Number_y'], inplace=True, errors='ignore')

        # Add an empty 'RCA' column
        df_final['RCA'] = ''
        
        # Sort and prepare the final report
        mtn_outage_report = df_final[['Ticket ID', 'Alarm Name', 'Site ID', 'ATC ID', 'Outage start time', 'ISM', 'SDS', 'Anchor Site ID', 'RCA']].sort_values(by='ISM')

        # Display the final report
        st.write("MTN Outage Report:")
        st.dataframe(mtn_outage_report)

        # Step to save the report
        save_report = st.button("Save MTN Outage Report to Directory")

        if save_report:
            # Define filename and save location
            tz = pytz.timezone('Africa/Lagos')
            time_str = datetime.datetime.now(tz).strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"MTN_Outage_Report_{time_str}.xlsx"
            filepath = os.path.join(r"C:\Users\semilore.fadumila\OneDrive - American Tower\Desktop\MTN OUTAGE HALF HOURLY REPORT", filename)

            # Save the DataFrame to Excel
            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                mtn_outage_report.to_excel(writer, sheet_name='Outage Report', index=False)
            
            st.success(f"Report saved successfully as {filename} in the designated directory.")
    except Exception as e:
        st.error(f"Error during merging or report generation: {e}")

else:
    st.write("Please upload both the NOC Matrix and MTN Incident report raw data files.")