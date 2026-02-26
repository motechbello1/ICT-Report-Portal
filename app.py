import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import io
import plotly.express as px
from github import Github  
from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt

# --- Configuration ---
st.set_page_config(page_title="ICT Events Portal", layout="wide")
GITHUB_REPO = "motechbello1/ICT-Report-Portal" # Ensure this is correct

# File paths
EVENT_DATA_FILE = "ict_events_log.csv"
TRACKER_FILE = "last_event_report_date.txt"

# --- Improved GitHub Sync ---
def push_to_github(file_path):
    try:
        if "GITHUB_TOKEN" in st.secrets:
            g = Github(st.secrets["GITHUB_TOKEN"])
            repo = g.get_repo(GITHUB_REPO)
            with open(file_path, 'r') as file:
                content = file.read()
            
            try:
                contents = repo.get_contents(file_path)
                repo.update_file(contents.path, f"Auto-sync {file_path}", content, contents.sha)
                return True, "Synced to GitHub successfully!"
            except:
                repo.create_file(file_path, f"Auto-sync created {file_path}", content)
                return True, "Created and synced to GitHub successfully!"
        else:
            return False, "GITHUB_TOKEN not found in Secrets. Saved locally."
    except Exception as e:
        return False, f"GitHub Sync Error: {str(e)}"

# --- Smart Data Loader & Sanitizer ---
def load_and_clean_data(file_path, columns):
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        df = pd.read_csv(file_path)
        # SANITIZE DATA: Strip spaces and force Title Case for uniformity
        for col in df.columns:
            if col != "Attendee Count": # Don't capitalize numbers
                df[col] = df[col].astype(str).str.strip().str.title()
            else:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        return df
    else:
        # Auto-create the database if it doesn't exist
        df = pd.DataFrame(columns=columns)
        df.to_csv(file_path, index=False)
        return df

# --- Tracker for Report Dates ---
def get_last_report_date():
    if os.path.exists(TRACKER_FILE):
        with open(TRACKER_FILE, 'r') as f:
            try:
                return datetime.strptime(f.read().strip(), "%Y-%m-%d").date()
            except: pass
    return (datetime.now() - timedelta(days=7)).date()

def save_last_report_date(end_date):
    with open(TRACKER_FILE, 'w') as f:
        f.write(end_date.strftime("%Y-%m-%d"))

# ==============================================================================
# MAIN APPLICATION: ICT EVENTS DASHBOARD
# ==============================================================================
st.title("🎟️ Executive ICT Events Portal")
st.markdown("Centralized dashboard for logging, tracking, and reporting on all ICT-supported organizational events.")

# Handle memory messages globally
if "form_message" in st.session_state:
    if st.session_state.message_type == "success":
        st.success(st.session_state.form_message)
    else:
        st.warning(st.session_state.form_message)
    del st.session_state.form_message
    del st.session_state.message_type

# Load Data
event_columns = ["Date", "Day", "Time", "Event Name", "Location", "Coordinator", "Equipment Deployed", "Attendee Count", "Status", "Remarks"]
df_events = load_and_clean_data(EVENT_DATA_FILE, event_columns)

tab1, tab2, tab3, tab4 = st.tabs(["📝 Log Event Activity", "📊 Event Analytics", "🗄️ Master Database", "📑 Executive Smart Report"])

# --- TAB 1: DATA ENTRY ---
with tab1:
    with st.form("event_log_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            event_name = st.text_input("Event Name")
            location = st.text_input("Location/Venue")
            coordinator = st.text_input("ICT Coordinator")
            equipment = st.text_area("Equipment Deployed (PA, Projectors, Mics, etc.)")
        with col2:
            attendees = st.number_input("Estimated Attendee Count", min_value=0)
            status = st.selectbox("Event Status", ["Planned", "Ongoing", "Completed", "Cancelled", "Delayed"]) 
            remarks = st.text_area("Post-Event Remarks / Issues Faced")
            
        if st.form_submit_button("💾 Save Event & Sync to GitHub"):
            now = datetime.now()
            new_event = {
                "Date": now.strftime("%Y-%m-%d"), "Day": now.strftime("%A"), "Time": now.strftime("%H:%M:%S"),
                "Event Name": event_name, "Location": location, "Coordinator": coordinator, 
                "Equipment Deployed": equipment, "Attendee Count": attendees, 
                "Status": status, "Remarks": remarks
            }
            pd.DataFrame([new_event]).to_csv(EVENT_DATA_FILE, mode='a', header=not os.path.exists(EVENT_DATA_FILE), index=False)
            success, msg = push_to_github(EVENT_DATA_FILE)
            st.session_state.message_type = "success" if success else "warning"
            st.session_state.form_message = f"Event Log Status: {msg}"
            st.rerun()

# --- TAB 2: ANALYTICS ---
with tab2:
    if df_events.empty: 
        st.info("No events logged yet. Please add data in the 'Log Event Activity' tab.")
    else:
        st.subheader("Interactive Event Statistics")
        c1, c2 = st.columns(2)
        
        with c1:
            # Pie Chart
            col_metric = st.selectbox("Analyze Distribution By:", ["Status", "Location", "Coordinator"], index=0)
            counts = df_events[col_metric].value_counts().reset_index()
            counts.columns = [col_metric, "Count"]
            fig_pie = px.pie(counts, values="Count", names=col_metric, title=f"Events Distribution by {col_metric}", hole=0.3)
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with c2:
            # Bar Chart for Attendees
            top_events = df_events.sort_values(by="Attendee Count", ascending=False).head(10)
            fig_bar = px.bar(top_events, x="Event Name", y="Attendee Count", color="Location", title="Top 10 Largest Events by Attendance")
            st.plotly_chart(fig_bar, use_container_width=True)

# --- TAB 3: DATABASE & DOWNLOAD ---
with tab3:
    if df_events.empty:
        st.info("No data available.")
    else:
        st.markdown("### Search & Export Event Logs")
        col_s1, col_s2 = st.columns([3, 1])
        with col_s1:
            search_term = st.text_input("🔍 Search database by Event Name, Location, or Coordinator...").lower()
        with col_s2:
            selected_columns = st.multiselect("Visible Columns:", df_events.columns.tolist(), default=df_events.columns.tolist())
            
        filtered_df = df_events[selected_columns]
        if search_term:
            mask = filtered_df.astype(str).apply(lambda x: x.str.contains(search_term, case=False)).any(axis=1)
            filtered_df = filtered_df[mask]
            
        st.dataframe(filtered_df, use_container_width=True)
        
        # Download Button
        csv_data = filtered_df.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Data as CSV", csv_data, f"ICT_Events_Export_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")

# --- TAB 4: AUTOMATED SMART REPORT (.DOCX) ---
with tab4:
    st.subheader("Generate Factual Executive Event Report")
    
    last_report_start = get_last_report_date()
    col_d1, col_d2 = st.columns(2)
    with col_d1: start_date = st.date_input("Start Date", value=last_report_start)
    with col_d2: end_date = st.date_input("End Date", value=last_report_start + timedelta(days=7))

    if st.button("📄 Generate Executive Analysis Report"):
        if df_events.empty:
            st.error("Database is empty. No reports can be generated.")
        else:
            temp_df = df_events.copy()
            temp_df['Date'] = pd.to_datetime(temp_df['Date']).dt.date
            logged_dates = temp_df['Date'].unique()
            
            # STRICT FACTUAL CHECK
            if start_date not in logged_dates or end_date not in logged_dates:
                st.error(f"⚠️ Authenticity Error: One or both of the selected dates ({start_date} to {end_date}) have zero logged events. Please select dates that contain actual data to ensure reporting accuracy.")
            else:
                report_df = temp_df[(temp_df['Date'] >= start_date) & (temp_df['Date'] <= end_date)].copy()
                
                if report_df.empty:
                    st.warning("No records found in this date range.")
                else:
                    # Core Analytics
                    total_events = len(report_df)
                    total_attendees = report_df['Attendee Count'].sum()
                    completed = len(report_df[report_df['Status'] == 'Completed'])
                    comp_rate = (completed / total_events) * 100 if total_events > 0 else 0
                    
                    top_loc = report_df['Location'].value_counts().index[0]
                    top_loc_count = report_df['Location'].value_counts().iloc[0]
                    top_coord = report_df['Coordinator'].value_counts().index[0]

                    # Generate Document
                    doc = Document()
                    doc.add_heading('ICT Event Operations & Analytics Report', 0)
                    doc.add_paragraph(f"Reporting Period: {start_date} to {end_date}\nGenerated on: {datetime.now().strftime('%Y-%m-%d')}")
                    
                    doc.add_heading('Executive Summary', level=1)
                    doc.add_paragraph(f"During this period, the ICT department successfully supported {total_events} distinct events, facilitating technological requirements for an estimated total of {total_attendees} attendees across the organization.")

                    # 1. Operational Status
                    fig_stat, ax_stat = plt.subplots(figsize=(5, 3))
                    report_df['Status'].value_counts().plot(kind='pie', autopct='%1.1f%%', colors=['#28a745', '#ffc107', '#dc3545', '#17a2b8', '#6c757d'], ax=ax_stat)
                    ax_stat.set_ylabel("")
                    fig_stat.savefig("e_stat.png", bbox_inches='tight')
                    plt.close(fig_stat)

                    doc.add_heading('1. Event Completion & Status Distribution', level=1)
                    doc.add_picture("e_stat.png", width=Inches(4.5))
                    doc.add_paragraph(
                        f"Analysis: The team achieved a {comp_rate:.1f}% completion rate for scheduled events during this reporting window. "
                        f"Tracking cancellations, delays, and ongoing statuses helps identify bottlenecks in pre-event planning.\n\n"
                        f"Future Recommendation: If the cancellation or delay rate exceeds 15%, ICT should mandate a strict '48-hour hardware confirmation' policy with event hosts to ensure equipment isn't reserved and left unused."
                    )

                    # 2. Location Analysis
                    fig_loc, ax_loc = plt.subplots(figsize=(5, 3))
                    report_df['
