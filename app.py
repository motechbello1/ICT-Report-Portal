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
# This forces "FELIX", "felix", and "Felix" to all become perfectly matching "Felix"
def load_and_clean_data(file_path, columns):
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        df = pd.read_csv(file_path)
        # SANITIZE DATA: Strip spaces and force Title Case for uniformity
        for col in df.columns:
            # We don't want to capitalize numeric columns if they exist
            if col != "Attendee Count":
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

# ==========================================
# SIDEBAR NAVIGATION SWITCH
# ==========================================
st.sidebar.title("🎛️ System Navigation")
# Modifed the Helpdesk name to ICT Departmental Activities
app_mode = st.sidebar.radio("Select Portal Module:", ["🛠️ ICT Departmental Activities", "📅 ICT Events Dashboard"])
st.sidebar.markdown("---")
st.sidebar.info("The system automatically maintains separate databases for Activity logs and Event schedules.")

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
                it_staff = st.text_input("IT Staff Name")
                status = st.selectbox("Status", ["Resolved", "Pending"]) 
                remarks = st.text_area("Remarks")
                
            if st.form_submit_button("💾 Save Activity Log & Sync"):
                now = datetime.now()
                new_data = {
                    "Date": now.strftime("%Y-%m-%d"), "Day": now.strftime("%A"), "Time": now.strftime("%H:%M:%S"),
                    "Department": department, "Reported By": reported_by, "System ID": system_id, 
                    "Description": description, "Problem": problem, "Action": action,
                    "Parts": parts, "IT Staff": it_staff, "Status": status, "Remarks": remarks
                }
                pd.DataFrame([new_data]).to_csv(ICT_DATA_FILE, mode='a', header=not os.path.exists(ICT_DATA_FILE), index=False)
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
            counts = df_ict[col_name].value_counts().reset_index()
            counts.columns = [col_name, "Count"]
            
            if col_name == "Status":
                fig = px.pie(counts, values="Count", names="Status", color="Status", color_discrete_map={"Resolved": "#28a745", "Pending": "#dc3545"})
            else:
                fig = px.bar(counts, x=col_name, y="Count", color=col_name)
            st.plotly_chart(fig, use_container_width=True)

    with tab3:
        st.dataframe(df_ict, use_container_width=True)

    with tab4:
        st.subheader("Generate In-Depth Executive Action Report")
        
        last_report_start = get_last_report_date()
        col_d1, col_d2 = st.columns(2)
        with col_d1: start_date = st.date_input("Start Date", value=last_report_start)
        with col_d2: end_date = st.date_input("End Date", value=last_report_start + timedelta(days=7))

        if st.button("📄 Generate Deep Analysis Report"):
            if df_ict.empty:
                st.error("Database is empty.")
            else:
                temp_df = df_ict.copy()
                temp_df['Date'] = pd.to_datetime(temp_df['Date']).dt.date
                logged_dates = temp_df['Date'].unique()
                
                if start_date not in logged_dates or end_date not in logged_dates:
                    st.error("⚠️ Authenticity Error: Selected dates contain no data. Please select valid logging dates.")
                else:
                    weekly_df = temp_df[(temp_df['Date'] >= start_date) & (temp_df['Date'] <= end_date)].copy()
                    
                    # Core Analytics
                    total_issues = len(weekly_df)
                    resolved = len(weekly_df[weekly_df['Status'] == 'Resolved'])
                    res_rate = (resolved / total_issues) * 100 if total_issues > 0 else 0
                    
                    top_dept = weekly_df['Department'].value_counts().index[0]
                    top_staff = weekly_df['IT Staff'].value_counts().index[0]
                    
                    # Top Problem Analytics
                    problem_counts = weekly_df['Problem'].value_counts()
                    top_prob = problem_counts.index[0]
                    top_prob_count = problem_counts.iloc[0]
                    top_prob_percent = (top_prob_count / total_issues) * 100

                    # Generate Document
                    doc = Document()
                    doc.add_heading('ICT Executive Analytics & Systems Health Report', 0)
                    doc.add_paragraph(f"Reporting Period: {start_date} to {end_date}\nGenerated on: {datetime.now().strftime('%Y-%m-%d')}")
                    
                    # 1. Top Complaint Deep Dive
                    fig_prob, ax_prob = plt.subplots(figsize=(6, 4))
                    problem_counts.head(5).plot(kind='barh', color='#ff6b6b', ax=ax_prob)
                    ax_prob.invert_yaxis()
                    ax_prob.set_title("Top 5 Systemic Complaints")
                    fig_prob.savefig("c_prob.png", bbox_inches='tight')
                    plt.close(fig_prob)

                    doc.add_heading('1. Critical Vulnerability: Most Issued Complaint', level=1)
                    doc.add_picture("c_prob.png", width=Inches(5.0))
                    doc.add_paragraph(
                        f"Executive In-Depth Analysis: During this reporting window, the most frequent point of failure across the organization was '{top_prob}'. "
                        f"This specific issue accounted for {top_prob_count} separate incident reports, representing {top_prob_percent:.1f}% of our total departmental workload. "
                        f"Because this complaint is recurring across multiple users, it indicates a systemic hardware or software limitation rather than isolated user error. "
                        f"Strategic Recommendation: The ICT department advises a targeted audit of systems prone to '{top_prob}' to determine if proactive replacement or centralized software patching is more cost-effective than continuous reactive troubleshooting."
                    )

                    # 2. Department Analysis
                    fig_dept, ax_dept = plt.subplots(figsize=(5, 3))
                    weekly_df['Department'].value_counts().plot(kind='bar', color='#4dabf7', ax=ax_dept)
                    plt.xticks(rotation=45, ha='right')
                    fig_dept.savefig("c_dept.png", bbox_inches='tight')
                    plt.close(fig_dept)

                    doc.add_heading('2. Departmental Resource Consumption', level=1)
                    doc.add_picture("c_dept.png", width=Inches(5.0))
                    doc.add_paragraph(
                        f"Executive In-Depth Analysis: Resource allocation heavily skewed towards the {top_dept} department, which generated the highest volume of support tickets. "
                        f"This metric is vital for operational budgeting. If the {top_dept} department continues to dominate ICT support queues, it may necessitate assigning a dedicated liaison to that unit, or conducting specialized training to reduce user-generated errors. "
                        f"Conversely, departments with low ticket volumes are operating with stable infrastructure."
                    )

                    # 3. Resolution Efficiency
                    fig_stat, ax_stat = plt.subplots(figsize=(5, 3))
                    weekly_df['Status'].value_counts().plot(kind='pie', autopct='%1.1f%%', colors=['#20c997', '#ff8787'], ax=ax_stat)
                    ax_stat.set_ylabel("")
                    fig_stat.savefig("c_stat.png", bbox_inches='tight')
                    plt.close(fig_stat)

                    doc.add_heading('3. Team Resolution Efficiency', level=1)
                    doc.add_picture("c_stat.png", width=Inches(4.0))
                    doc.add_paragraph(
                        f"Executive In-Depth Analysis: The overarching performance of the ICT team is measured at a {res_rate:.1f}% resolution efficiency for this period. "
                        f"Out of {total_issues} reported incidents, {resolved} were successfully mitigated and closed. Commendations to {top_staff}, who personally closed the highest number of tickets during this cycle. "
                        f"The remaining pending tasks are being rolled over into the next sprint prioritization queue."
                    )

                    # Cleanup
                    for file in ["c_prob.png", "c_dept.png", "c_stat.png"]:
                        if os.path.exists(file): os.remove(file)

                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    save_last_report_date(end_date)
                    
                    st.success(f"✅ Deep Analytics Report generated for {start_date} to {end_date}!")
                    st.download_button("📥 Download Executive Report (.docx)", doc_buffer, f"ICT_Activities_Report_{end_date}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ==============================================================================
# MODULE 2: ICT EVENTS DASHBOARD
# ==============================================================================
elif app_mode == "📅 ICT Events Dashboard":
    st.title("🎟️ ICT Event Activity Dashboard")
    
    event_columns = ["Date", "Day", "Time", "Event Name", "Location", "Coordinator", "Equipment Deployed", "Attendee Count", "Status", "Remarks"]
    df_events = load_and_clean_data(EVENT_DATA_FILE, event_columns)

    # Added the 4th tab for Reports
    etab1, etab2, etab3, etab4 = st.tabs(["📝 Log Event Activity", "📊 Event Analytics", "🗄️ Event Database", "📑 Executive Event Report"])

    with etab1:
        with st.form("event_log_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                event_name = st.text_input("Event Name")
                location = st.text_input("Location/Venue")
                coordinator = st.text_input("ICT Coordinator")
                equipment = st.text_area("Equipment Deployed (PA, Projectors, etc.)")
            with col2:
                attendees = st.number_input("Estimated Attendee Count", min_value=0)
                status = st.selectbox("Event Status", ["Planned", "Ongoing", "Completed", "Cancelled"]) 
                remarks = st.text_area("Post-Event Remarks")
                
            if st.form_submit_button("💾 Save Event & Sync"):
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
                st.session_state.form_message = f"Event Data: {msg}"
                st.rerun()

    with etab2:
        if df_events.empty: st.info("No events logged yet.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                col_metric = st.selectbox("Analyze Events By:", ["Status", "Location", "Coordinator"])
                counts = df_events[col_metric].value_counts().reset_index()
                counts.columns = [col_metric, "Count"]
                fig_pie = px.pie(counts, values="Count", names=col_metric, title=f"Events Distribution by {col_metric}", hole=0.3)
                st.plotly_chart(fig_pie, use_container_width=True)
            with c2:
                # Added new chart for Event sizes
                top_events = df_events.sort_values(by="Attendee Count", ascending=False).head(10)
                fig_bar = px.bar(top_events, x="Event Name", y="Attendee Count", color="Location", title="Top Events by Attendee Volume")
                st.plotly_chart(fig_bar, use_container_width=True)

    with etab3:
        # Added Download button and Dataframe
        if not df_events.empty:
            csv_data = df_events.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Download Event Data as CSV", csv_data, f"ICT_Events_Export_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")
        st.dataframe(df_events, use_container_width=True)

    with etab4:
        st.subheader("Generate Factual Executive Event Report")
        
        last_report_start = get_last_report_date()
        col_d1, col_d2 = st.columns(2)
        with col_d1: start_date = st.date_input("Start Date (Events)", value=last_report_start)
        with col_d2: end_date = st.date_input("End Date (Events)", value=last_report_start + timedelta(days=7))

        if st.button("📄 Generate Event Analysis Report"):
            if df_events.empty:
                st.error("Event Database is empty.")
            else:
                temp_df = df_events.copy()
                temp_df['Date'] = pd.to_datetime(temp_df['Date']).dt.date
                logged_dates = temp_df['Date'].unique()
                
                if start_date not in logged_dates or end_date not in logged_dates:
                    st.error("⚠️ Authenticity Error: Selected dates contain no event data.")
                else:
                    report_df = temp_df[(temp_df['Date'] >= start_date) & (temp_df['Date'] <= end_date)].copy()
                    
                    if report_df.empty:
                        st.warning("No records found in this specific date range.")
                    else:
                        # Analytics
                        total_events = len(report_df)
                        total_attendees = report_df['Attendee Count'].sum()
                        completed = len(report_df[report_df['Status'] == 'Completed'])
                        comp_rate = (completed / total_events) * 100 if total_events > 0 else 0
                        top_loc = report_df['Location'].value_counts().index[0]
                        top_coord = report_df['Coordinator'].value_counts().index[0]

                        # Generate Document
                        doc = Document()
                        doc.add_heading('ICT Event Operations & Analytics Report', 0)
                        doc.add_paragraph(f"Reporting Period: {start_date} to {end_date}\nGenerated on: {datetime.now().strftime('%Y-%m-%d')}")
                        
                        doc.add_heading('Executive Summary', level=1)
                        doc.add_paragraph(f"During this period, the ICT department successfully supported {total_events} distinct events, facilitating technological requirements for an estimated total of {total_attendees} attendees across the organization.")

                        # 1. Operational Status
                        fig_stat, ax_stat = plt.subplots(figsize=(5, 3))
                        report_df['Status'].value_counts().plot(kind='pie', autopct='%1.1f%%', colors=['#28a745', '#ffc107', '#dc3545', '#17a2b8'], ax=ax_stat)
                        ax_stat.set_ylabel("")
                        fig_stat.savefig("e_stat.png", bbox_inches='tight')
                        plt.close(fig_stat)

                        doc.add_heading('1. Event Completion & Status Distribution', level=1)
                        doc.add_picture("e_stat.png", width=Inches(4.5))
                        doc.add_paragraph(
                            f"Analysis: The team achieved a {comp_rate:.1f}% event completion rate for scheduled activities during this reporting window. "
                            f"Tracking cancellations and ongoing statuses helps identify bottlenecks in pre-event planning.\n\n"
                            f"Future Recommendation: If the cancellation or delay rate exceeds 15%, ICT should mandate a strict '48-hour hardware confirmation' policy with event hosts to ensure equipment isn't reserved and left unused."
                        )

                        # 2. Location Analysis
                        fig_loc, ax_loc = plt.subplots(figsize=(5, 3))
                        report_df['Location'].value_counts().plot(kind='bar', color='#007bff', ax=ax_loc)
                        plt.xticks(rotation=45, ha='right')
                        fig_loc.savefig("e_loc.png", bbox_inches='tight')
                        plt.close(fig_loc)

                        doc.add_heading('2. Venue Utilization & Hardware Demand', level=1)
                        doc.add_picture("e_loc.png", width=Inches(4.5))
                        doc.add_paragraph(
                            f"Analysis: The highest concentration of event-related hardware deployment occurred at '{top_loc}'. "
                            f"This indicates that this specific venue is the primary hub for organizational gatherings.\n\n"
                            f"Future Recommendation: Given the high frequency of events at '{top_loc}', we recommend conducting a proactive audit of permanently installed AV equipment in this location. Upgrading built-in infrastructure here will significantly reduce the manual setup time currently required by our coordinators."
                        )
                        
                        doc.add_heading('3. Personnel Allocation', level=1)
                        doc.add_paragraph(
                            f"Analysis: {top_coord} served as the primary coordinator for the highest volume of events during this cycle. "
                            f"To prevent burnout and ensure equal skill distribution, management should consider shadowing programs to rotate event assignments among junior IT staff."
                        )

                        # Cleanup
                        for file in ["e_stat.png", "e_loc.png"]:
                            if os.path.exists(file): os.remove(file)

                        doc_buffer = io.BytesIO()
                        doc.save(doc_buffer)
                        doc_buffer.seek(0)
                        
                        st.success(f"✅ Executive Event Report generated for {start_date} to {end_date}!")
                        st.download_button("📥 Download Event Report (.docx)", doc_buffer, f"ICT_Events_Report_{end_date}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
