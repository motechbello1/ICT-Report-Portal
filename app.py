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
st.set_page_config(page_title="ICT Department Portal", layout="wide")
GITHUB_REPO = "motechbello1/ICT-Report-Portal" 

# File paths
ICT_DATA_FILE = "ict_master_log.csv"
EVENT_DATA_FILE = "ict_events_log.csv"
TRACKER_FILE = "last_report_date.txt"

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
# Forces everything to lowercase (e.g. "FELIX" -> "felix") for perfect matching
def load_and_clean_data(file_path, columns):
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        df = pd.read_csv(file_path)
        for col in df.columns:
            if col != "Attendee Count": # Don't lower-case numbers
                df[col] = df[col].astype(str).str.strip().str.lower()
            else:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        return df
    else:
        df = pd.DataFrame(columns=columns)
        df.to_csv(file_path, index=False)
        return df

# Helper to handle multiple comma-separated staff names
def explode_staff_counts(df, column_name):
    if df.empty or column_name not in df.columns:
        return pd.Series(dtype=int)
    # Split by comma, expand into rows, strip spaces, and count
    return df[column_name].dropna().astype(str).str.split(',').explode().str.strip().value_counts()

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

# ==========================================
# SIDEBAR NAVIGATION SWITCH
# ==========================================
st.sidebar.title("🎛️ System Navigation")
app_mode = st.sidebar.radio("Select Portal Module:", [
    "🛠️ ICT Departmental Activities", 
    "📅 ICT Events Dashboard",
    "📑 Master Combined Report"
])
st.sidebar.markdown("---")
st.sidebar.info("The system automatically maintains separate databases and intelligently handles multiple staff assignments.")

# Handle memory messages globally
if "form_message" in st.session_state:
    if st.session_state.message_type == "success":
        st.success(st.session_state.form_message)
    else:
        st.warning(st.session_state.form_message)
    del st.session_state.form_message
    del st.session_state.message_type


# ==============================================================================
# MODULE 1: ICT DEPARTMENTAL ACTIVITIES
# ==============================================================================
if app_mode == "🛠️ ICT Departmental Activities":
    st.title("💻 Executive ICT Departmental Activities Portal")
    
    ict_columns = ["Date", "Day", "Time", "Department", "Reported By", "System ID", "Description", "Problem", "Action", "Parts", "IT Staff", "Status", "Remarks"]
    df_ict = load_and_clean_data(ICT_DATA_FILE, ict_columns)

    tab1, tab2, tab3, tab4 = st.tabs(["📝 Log Support Activity", "📊 Analytics Dashboard", "🗄️ Master Database", "📑 Executive Smart Report"])

    with tab1:
        with st.form("ict_log_form", clear_on_submit=True):
            st.markdown("**📅 Timestamping (Change if logging a past event)**")
            t_col1, t_col2 = st.columns(2)
            with t_col1: log_date = st.date_input("Date", value=datetime.today())
            with t_col2: log_time = st.time_input("Time", value=datetime.now().time())
            st.markdown("---")
            
            col1, col2 = st.columns(2)
            with col1:
                department = st.text_input("Department")
                reported_by = st.text_input("Reported By")
                system_id = st.text_input("System ID")
                description = st.text_input("System Description")
                problem = st.text_area("Diagnosed Problem")
            with col2:
                action = st.text_area("Action Taken")
                parts = st.text_input("Replaced Parts (if any)")
                it_staff = st.text_input("IT Staff Name(s) - Separate with commas")
                status = st.selectbox("Status", ["resolved", "pending"]) 
                remarks = st.text_area("Remarks")
                
            if st.form_submit_button("💾 Save Activity Log & Sync"):
                new_data = {
                    "Date": log_date.strftime("%Y-%m-%d"), "Day": log_date.strftime("%A").lower(), "Time": log_time.strftime("%H:%M:%S"),
                    "Department": department.lower(), "Reported By": reported_by.lower(), "System ID": system_id.lower(), 
                    "Description": description.lower(), "Problem": problem.lower(), "Action": action.lower(),
                    "Parts": parts.lower(), "IT Staff": it_staff.lower(), "Status": status, "Remarks": remarks.lower()
                }
                
                # Append, Sort chronologically, and Save
                temp_df = pd.concat([df_ict, pd.DataFrame([new_data])], ignore_index=True)
                temp_df = temp_df.sort_values(by=["Date", "Time"]).reset_index(drop=True)
                temp_df.to_csv(ICT_DATA_FILE, index=False)
                
                success, msg = push_to_github(ICT_DATA_FILE)
                st.session_state.message_type = "success" if success else "warning"
                st.session_state.form_message = msg
                st.rerun()

    with tab2:
        if df_ict.empty: st.info("No activity data logged yet.")
        else:
            metric_mapping = {"Status": "Status", "Department": "Department", "IT Staff": "IT Staff", "Diagnosed Problem": "Problem"}
            sort_by = st.selectbox("Analyze Activity Metric:", list(metric_mapping.keys()))
            col_name = metric_mapping[sort_by]
            
            # Smart handling for multiple staff
            if col_name == "IT Staff":
                counts = explode_staff_counts(df_ict, "IT Staff").reset_index()
                counts.columns = ["IT Staff", "Count"]
            else:
                counts = df_ict[col_name].value_counts().reset_index()
                counts.columns = [col_name, "Count"]
            
            if col_name == "Status":
                fig = px.pie(counts, values="Count", names="Status", color="Status", color_discrete_map={"resolved": "#28a745", "pending": "#dc3545"})
            else:
                fig = px.bar(counts, x=col_name, y="Count", color=col_name)
            st.plotly_chart(fig, use_container_width=True)

    with tab3:
        if not df_ict.empty:
            csv_data = df_ict.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Download Activity Data as CSV", csv_data, f"ICT_Activities_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")
        st.dataframe(df_ict, use_container_width=True)

    with tab4:
        st.subheader("Generate In-Depth Executive Action Report")
        
        last_report_start = get_last_report_date()
        col_d1, col_d2 = st.columns(2)
        with col_d1: start_date = st.date_input("Start Date", value=last_report_start)
        with col_d2: end_date = st.date_input("End Date", value=last_report_start + timedelta(days=7))

        if st.button("📄 Generate Deep Analysis Report"):
            if df_ict.empty: st.error("Database is empty.")
            else:
                temp_df = df_ict.copy()
                temp_df['Date'] = pd.to_datetime(temp_df['Date']).dt.date
                weekly_df = temp_df[(temp_df['Date'] >= start_date) & (temp_df['Date'] <= end_date)].copy()
                
                if weekly_df.empty: st.warning("No records found in this date range.")
                else:
                    total_issues = len(weekly_df)
                    resolved = len(weekly_df[weekly_df['Status'] == 'resolved'])
                    res_rate = (resolved / total_issues) * 100 if total_issues > 0 else 0
                    
                    top_dept = weekly_df['Department'].value_counts().index[0]
                    staff_counts = explode_staff_counts(weekly_df, "IT Staff")
                    top_staff = staff_counts.index[0] if not staff_counts.empty else "N/A"
                    
                    problem_counts = weekly_df['Problem'].value_counts()
                    top_prob = problem_counts.index[0]

                    doc = Document()
                    doc.add_heading('ICT Executive Analytics Report', 0)
                    doc.add_paragraph(f"Reporting Period: {start_date} to {end_date}\nGenerated on: {datetime.now().strftime('%Y-%m-%d')}")
                    
                    fig_prob, ax_prob = plt.subplots(figsize=(6, 4))
                    problem_counts.head(5).plot(kind='barh', color='#ff6b6b', ax=ax_prob)
                    ax_prob.invert_yaxis()
                    ax_prob.set_title("Top Systemic Complaints")
                    fig_prob.savefig("c_prob.png", bbox_inches='tight')
                    plt.close(fig_prob)

                    doc.add_heading('1. Systemic Complaints', level=1)
                    doc.add_picture("c_prob.png", width=Inches(5.0))
                    doc.add_paragraph(f"Analysis: The most frequent issue was '{top_prob}'. Action is recommended to permanently patch or replace systems prone to this error.")

                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    if os.path.exists("c_prob.png"): os.remove("c_prob.png")
                    save_last_report_date(end_date)
                    
                    st.success(f"✅ Report generated for {start_date} to {end_date}!")
                    st.download_button("📥 Download Executive Report (.docx)", doc_buffer, f"ICT_Activities_{end_date}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ==============================================================================
# MODULE 2: ICT EVENTS DASHBOARD
# ==============================================================================
elif app_mode == "📅 ICT Events Dashboard":
    st.title("🎟️ ICT Event Activity Dashboard")
    
    event_columns = ["Date", "Day", "Time", "Event Name", "Location", "Coordinator", "Equipment Deployed", "Attendee Count", "Status", "Remarks"]
    df_events = load_and_clean_data(EVENT_DATA_FILE, event_columns)

    etab1, etab2, etab3, etab4 = st.tabs(["📝 Log Event Activity", "📊 Event Analytics", "🗄️ Event Database", "📑 Executive Event Report"])

    with etab1:
        with st.form("event_log_form", clear_on_submit=True):
            st.markdown("**📅 Timestamping (Change if logging a past event)**")
            t_col1, t_col2 = st.columns(2)
            with t_col1: log_date = st.date_input("Date", value=datetime.today())
            with t_col2: log_time = st.time_input("Time", value=datetime.now().time())
            st.markdown("---")

            col1, col2 = st.columns(2)
            with col1:
                event_name = st.text_input("Event Name")
                location = st.text_input("Location/Venue")
                coordinator = st.text_input("ICT Coordinator(s) - Separate with commas")
                equipment = st.text_area("Equipment Deployed")
            with col2:
                attendees = st.number_input("Estimated Attendee Count", min_value=0)
                status = st.selectbox("Event Status", ["planned", "ongoing", "completed", "cancelled"]) 
                remarks = st.text_area("Post-Event Remarks")
                
            if st.form_submit_button("💾 Save Event & Sync"):
                new_event = {
                    "Date": log_date.strftime("%Y-%m-%d"), "Day": log_date.strftime("%A").lower(), "Time": log_time.strftime("%H:%M:%S"),
                    "Event Name": event_name.lower(), "Location": location.lower(), "Coordinator": coordinator.lower(), 
                    "Equipment Deployed": equipment.lower(), "Attendee Count": attendees, 
                    "Status": status, "Remarks": remarks.lower()
                }
                
                temp_df = pd.concat([df_events, pd.DataFrame([new_event])], ignore_index=True)
                temp_df = temp_df.sort_values(by=["Date", "Time"]).reset_index(drop=True)
                temp_df.to_csv(EVENT_DATA_FILE, index=False)
                
                success, msg = push_to_github(EVENT_DATA_FILE)
                st.session_state.message_type = "success" if success else "warning"
                st.session_state.form_message = msg
                st.rerun()

    with etab2:
        if df_events.empty: st.info("No events logged yet.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                col_metric = st.selectbox("Analyze Events By:", ["Status", "Location", "Coordinator"])
                
                if col_metric == "Coordinator":
                    counts = explode_staff_counts(df_events, "Coordinator").reset_index()
                    counts.columns = ["Coordinator", "Count"]
                else:
                    counts = df_events[col_metric].value_counts().reset_index()
                    counts.columns = [col_metric, "Count"]
                    
                fig_pie = px.pie(counts, values="Count", names=col_metric, title=f"Events Distribution by {col_metric}", hole=0.3)
                st.plotly_chart(fig_pie, use_container_width=True)
            with c2:
                top_events = df_events.sort_values(by="Attendee Count", ascending=False).head(10)
                fig_bar = px.bar(top_events, x="Event Name", y="Attendee Count", color="Location", title="Top Events by Attendee Volume")
                st.plotly_chart(fig_bar, use_container_width=True)

    with etab3:
        if not df_events.empty:
            csv_data = df_events.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Download Event Data CSV", csv_data, f"ICT_Events_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")
        st.dataframe(df_events, use_container_width=True)

    with etab4:
        st.subheader("Generate Factual Executive Event Report")
        
        last_report_start = get_last_report_date()
        col_d1, col_d2 = st.columns(2)
        with col_d1: start_date = st.date_input("Start Date (Events)", value=last_report_start)
        with col_d2: end_date = st.date_input("End Date (Events)", value=last_report_start + timedelta(days=7))

        if st.button("📄 Generate Event Analysis Report"):
            if df_events.empty: st.error("Event Database is empty.")
            else:
                temp_df = df_events.copy()
                temp_df['Date'] = pd.to_datetime(temp_df['Date']).dt.date
                report_df = temp_df[(temp_df['Date'] >= start_date) & (temp_df['Date'] <= end_date)].copy()
                
                if report_df.empty: st.warning("No records found in this specific date range.")
                else:
                    total_events = len(report_df)
                    coord_counts = explode_staff_counts(report_df, "Coordinator")
                    top_coord = coord_counts.index[0] if not coord_counts.empty else "N/A"

                    doc = Document()
                    doc.add_heading('ICT Event Operations Report', 0)
                    doc.add_paragraph(f"Reporting Period: {start_date} to {end_date}\nGenerated on: {datetime.now().strftime('%Y-%m-%d')}")
                    
                    fig_stat, ax_stat = plt.subplots(figsize=(5, 3))
                    report_df['Status'].value_counts().plot(kind='pie', autopct='%1.1f%%', ax=ax_stat)
                    ax_stat.set_ylabel("")
                    fig_stat.savefig("e_stat.png", bbox_inches='tight')
                    plt.close(fig_stat)

                    doc.add_heading('1. Event Status Distribution', level=1)
                    doc.add_picture("e_stat.png", width=Inches(4.5))
                    doc.add_paragraph(f"Analysis: {top_coord} served as the primary coordinator for the highest volume of events. A total of {total_events} events were logged in this period.")

                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    if os.path.exists("e_stat.png"): os.remove("e_stat.png")
                    
                    st.success("✅ Executive Event Report generated!")
                    st.download_button("📥 Download Event Report (.docx)", doc_buffer, f"ICT_Events_Report_{end_date}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ==============================================================================
# MODULE 3: MASTER COMBINED REPORT
# ==============================================================================
elif app_mode == "📑 Master Combined Report":
    st.title("📑 Master ICT Consolidated Report")
    st.markdown("Generate a unified executive report merging analytics from **Departmental Activities** and **Events**.")

    ict_columns = ["Date", "Day", "Time", "Department", "Reported By", "System ID", "Description", "Problem", "Action", "Parts", "IT Staff", "Status", "Remarks"]
    event_columns = ["Date", "Day", "Time", "Event Name", "Location", "Coordinator", "Equipment Deployed", "Attendee Count", "Status", "Remarks"]
    
    df_ict = load_and_clean_data(ICT_DATA_FILE, ict_columns)
    df_events = load_and_clean_data(EVENT_DATA_FILE, event_columns)

    if df_ict.empty and df_events.empty:
        st.warning("Both databases are currently empty. Please log data first.")
    else:
        last_report_start = get_last_report_date()
        col_d1, col_d2 = st.columns(2)
        with col_d1: start_date = st.date_input("Start Date (Combined)", value=last_report_start)
        with col_d2: end_date = st.date_input("End Date (Combined)", value=last_report_start + timedelta(days=7))

        if st.button("📄 Generate Master Consolidated Report"):
            
            # Filter Data
            df_ict['Date'] = pd.to_datetime(df_ict['Date'], errors='coerce').dt.date
            df_events['Date'] = pd.to_datetime(df_events['Date'], errors='coerce').dt.date
            
            mask_ict = (df_ict['Date'] >= start_date) & (df_ict['Date'] <= end_date)
            mask_evt = (df_events['Date'] >= start_date) & (df_events['Date'] <= end_date)
            
            rep_ict = df_ict[mask_ict]
            rep_evt = df_events[mask_evt]

            if rep_ict.empty and rep_evt.empty:
                st.error("No data found in either database for the selected date range.")
            else:
                # 1. Calculate Combined Workload (Smart multiple staff handling)
                staff_act = explode_staff_counts(rep_ict, "IT Staff")
                staff_evt = explode_staff_counts(rep_evt, "Coordinator")
                
                # Combine and sum identical names
                combined_staff = pd.concat([staff_act, staff_evt], axis=1).fillna(0)
                combined_staff.columns = ['Activities', 'Events']
                combined_staff['Total Tasks'] = combined_staff['Activities'] + combined_staff['Events']
                combined_staff = combined_staff.sort_values(by='Total Tasks', ascending=False)

                # Initialize Document
                doc = Document()
                doc.add_heading('Master ICT Consolidated Operations Report', 0)
                doc.add_paragraph(f"Reporting Period: {start_date} to {end_date}\nGenerated on: {datetime.now().strftime('%Y-%m-%d')}")
                
                doc.add_heading('Executive Summary', level=1)
                doc.add_paragraph(
                    f"During this period, the ICT department executed a total of {len(rep_ict) + len(rep_evt)} core operations. "
                    f"This consists of {len(rep_ict)} departmental support activities and {len(rep_evt)} event technology deployments."
                )

                # 2. Workload Infographic
                if not combined_staff.empty:
                    fig_work, ax_work = plt.subplots(figsize=(7, 4))
                    combined_staff[['Activities', 'Events']].head(10).plot(kind='bar', stacked=True, color=['#4dabf7', '#ffc107'], ax=ax_work)
                    plt.title("Combined Staff Workload (Top 10)")
                    plt.xlabel("ICT Personnel")
                    plt.ylabel("Number of Tasks/Events")
                    plt.xticks(rotation=45, ha='right')
                    fig_work.savefig("m_work.png", bbox_inches='tight')
                    plt.close(fig_work)

                    doc.add_heading('1. Unified Personnel Workload Analysis', level=1)
                    doc.add_picture("m_work.png", width=Inches(5.5))
                    
                    top_performer = combined_staff.index[0]
                    doc.add_paragraph(
                        f"Analysis: {top_performer.title()} handled the highest overall volume of requests across both events and daily activities. "
                        f"Tracking combined metrics provides a factual basis for performance reviews and highlights potential bottleneck dependencies on specific staff members. "
                        f"Recommendation: Cross-train personnel to balance the workload displayed above."
                    )
                    if os.path.exists("m_work.png"): os.remove("m_work.png")

                # 3. Overall Completion Rate
                ict_resolved = len(rep_ict[rep_ict['Status'] == 'resolved']) if not rep_ict.empty else 0
                evt_completed = len(rep_evt[rep_evt['Status'] == 'completed']) if not rep_evt.empty else 0
                total_success = ict_resolved + evt_completed
                total_ops = len(rep_ict) + len(rep_evt)
                success_rate = (total_success / total_ops) * 100 if total_ops > 0 else 0

                doc.add_heading('2. Global Departmental Efficiency', level=1)
                doc.add_paragraph(
                    f"The consolidated success/completion rate for the ICT department stands at {success_rate:.1f}%. "
                    f"This accounts for all resolved hardware/software tickets and successfully executed event setups."
                )

                # Save and download
                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)
                
                st.success("✅ Master Consolidated Report Generated!")
                st.download_button("📥 Download Master Report (.docx)", doc_buffer, f"ICT_Master_Report_{end_date}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
