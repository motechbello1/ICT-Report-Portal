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
def load_and_clean_data(file_path, columns):
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        df = pd.read_csv(file_path)
        for col in df.columns:
            if col != "Attendee Count": 
                df[col] = df[col].astype(str).str.strip().str.lower()
            else:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        return df
    else:
        df = pd.DataFrame(columns=columns)
        df.to_csv(file_path, index=False)
        return df

def explode_staff_counts(df, column_name):
    if df.empty or column_name not in df.columns:
        return pd.Series(dtype=int)
    return df[column_name].dropna().astype(str).str.split(',').explode().str.strip().value_counts()

# --- Advanced Narrative Generator for Missing Dates ---
def get_idle_dates_narrative(start_date, end_date, active_dates_series):
    # Convert series to a set of unique python date objects
    active_dates = set(pd.to_datetime(active_dates_series).dt.date)
    delta = end_date - start_date
    all_dates = {start_date + timedelta(days=i) for i in range(delta.days + 1)}
    
    missing_dates = sorted(list(all_dates - active_dates))
    
    if not missing_dates:
        return "Operational tempo remained continuously high; activities were recorded on every single day within the selected reporting window."
    
    # Group consecutive missing dates
    groups = []
    current_group = [missing_dates[0]]
    for i in range(1, len(missing_dates)):
        if (missing_dates[i] - missing_dates[i-1]).days == 1:
            current_group.append(missing_dates[i])
        else:
            groups.append(current_group)
            current_group = [missing_dates[i]]
    groups.append(current_group)
    
    text_parts = []
    for g in groups:
        if len(g) == 1:
            text_parts.append(g[0].strftime("%B %d, %Y"))
        else:
            text_parts.append(f"between {g[0].strftime('%B %d, %Y')} and {g[-1].strftime('%B %d, %Y')}")
            
    date_str = ", ".join(text_parts)
    return f"It is highly important to note that no operations or incidents were recorded {date_str}. These documented idle periods represent strategic windows that management should actively utilize for scheduling proactive system maintenance, executing firmware updates, or conducting dedicated staff training sessions without risking disruption to standard organizational workflows."

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
                temp_df['Date_Obj'] = pd.to_datetime(temp_df['Date'], errors='coerce').dt.date
                weekly_df = temp_df[(temp_df['Date_Obj'] >= start_date) & (temp_df['Date_Obj'] <= end_date)].copy()
                
                if weekly_df.empty: 
                    st.warning(f"No records found between {start_date} and {end_date}.")
                else:
                    total_issues = len(weekly_df)
                    resolved = len(weekly_df[weekly_df['Status'] == 'resolved'])
                    pending = total_issues - resolved
                    res_rate = (resolved / total_issues) * 100 if total_issues > 0 else 0
                    
                    top_dept = weekly_df['Department'].value_counts().index[0].title()
                    staff_counts = explode_staff_counts(weekly_df, "IT Staff")
                    top_staff = staff_counts.index[0].title() if not staff_counts.empty else "N/A"
                    
                    problem_counts = weekly_df['Problem'].value_counts()
                    top_prob = problem_counts.index[0].title()
                    top_prob_count = problem_counts.iloc[0]

                    idle_narrative = get_idle_dates_narrative(start_date, end_date, weekly_df['Date'])

                    doc = Document()
                    doc.add_heading('Detailed ICT Departmental Operations & Health Report', 0)
                    doc.add_paragraph(f"Reporting Window: {start_date} to {end_date}\nGenerated on: {datetime.now().strftime('%B %d, %Y')}")
                    
                    doc.add_heading('1. Executive Overview & Operational Tempo', level=1)
                    doc.add_paragraph(
                        f"During this evaluation period, the ICT technical support division was engaged in {total_issues} recorded troubleshooting and deployment activities. "
                        f"The team successfully drove a {res_rate:.1f}% resolution rate, closing {resolved} distinct incident tickets. "
                        f"However, there remain {pending} pending tasks that require rollover attention in the immediate future. Unresolved tickets directly correlate with hardware downtime, which subsequently diminishes operational efficiency across the affected departments. "
                        f"\n\nRegarding the timeline of events: {idle_narrative}"
                    )

                    fig_prob, ax_prob = plt.subplots(figsize=(6, 4))
                    problem_counts.head(5).plot(kind='barh', color='#ff6b6b', ax=ax_prob)
                    ax_prob.invert_yaxis()
                    ax_prob.set_title("Top Systemic Complaints")
                    fig_prob.savefig("c_prob.png", bbox_inches='tight')
                    plt.close(fig_prob)

                    doc.add_heading('2. Critical System Vulnerabilities', level=1)
                    doc.add_picture("c_prob.png", width=Inches(5.0))
                    doc.add_paragraph(
                        f"A deep analytical dive into the specific technological failures reveals that '{top_prob}' is the most pervasive issue currently facing the infrastructure, accounting for {top_prob_count} separate emergency interventions. "
                        f"When a single issue manifests this frequently across the organization, it strongly indicates a systemic hardware decay or an overarching software vulnerability rather than isolated user negligence. "
                        f"It is the formal recommendation of the ICT department that management considers a targeted lifecycle replacement or network-wide software patching strategy specifically addressing '{top_prob}' to drastically reduce repetitive support bottlenecks."
                    )
                    
                    doc.add_heading('3. Resource Distribution & Team Dynamics', level=1)
                    doc.add_paragraph(
                        f"In terms of resource consumption, the {top_dept} department demanded the highest volume of ICT focus during this cycle. "
                        f"If this trajectory continues, it may be necessary to deploy dedicated departmental liaisons or implement specialized digital literacy training for {top_dept} staff to lower the volume of user-generated IT support tickets. "
                        f"\n\nInternally, {top_staff} executed the highest volume of ticket resolutions. Monitoring this human resource metric ensures fair task delegation and highlights personnel who are currently carrying the operational weight of the department."
                    )

                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    if os.path.exists("c_prob.png"): os.remove("c_prob.png")
                    save_last_report_date(end_date)
                    
                    st.success(f"✅ Deep Descriptive Report generated for {start_date} to {end_date}!")
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
                temp_df['Date_Obj'] = pd.to_datetime(temp_df['Date'], errors='coerce').dt.date
                report_df = temp_df[(temp_df['Date_Obj'] >= start_date) & (temp_df['Date_Obj'] <= end_date)].copy()
                
                if report_df.empty: 
                    st.warning(f"Absolutely no events or schedules were found recorded between {start_date} and {end_date}.")
                else:
                    total_events = len(report_df)
                    total_attendees = report_df['Attendee Count'].sum()
                    coord_counts = explode_staff_counts(report_df, "Coordinator")
                    top_coord = coord_counts.index[0].title() if not coord_counts.empty else "N/A"
                    top_loc = report_df['Location'].value_counts().index[0].title()
                    
                    idle_narrative = get_idle_dates_narrative(start_date, end_date, report_df['Date'])

                    doc = Document()
                    doc.add_heading('In-Depth ICT Event Logistics & Infrastructure Report', 0)
                    doc.add_paragraph(f"Analysis Timeline: {start_date} to {end_date}\nGenerated on: {datetime.now().strftime('%B %d, %Y')}")
                    
                    doc.add_heading('1. Scale of Impact and Audience Reach', level=1)
                    doc.add_paragraph(
                        f"The ICT events coordination team successfully engineered the technological backbones for {total_events} distinct organizational events during this review period. "
                        f"Crucially, the aggregate audience size that relied on our uninterrupted technological deployments was estimated at {total_attendees} attendees. "
                        f"Handling high-volume crowds requires meticulous pre-planning, rapid hardware setup, and dedicated monitoring to prevent public-facing technical failures that could harm the organization's reputation. "
                        f"\n\nReviewing the calendar distribution of these events: {idle_narrative}"
                    )
                    
                    fig_stat, ax_stat = plt.subplots(figsize=(5, 3))
                    report_df['Status'].value_counts().plot(kind='pie', autopct='%1.1f%%', ax=ax_stat)
                    ax_stat.set_ylabel("")
                    fig_stat.savefig("e_stat.png", bbox_inches='tight')
                    plt.close(fig_stat)

                    doc.add_heading('2. Event Success Metrics and Status Integrity', level=1)
                    doc.add_picture("e_stat.png", width=Inches(4.5))
                    doc.add_paragraph(
                        f"The infographic above delineates the operational lifecycle outcomes of the scheduled events. "
                        f"While completion numbers signal a strong delivery framework, any 'Cancelled' or excessively 'Ongoing/Delayed' statuses inherently translate to wasted labor hours and misallocated technological hardware. "
                        f"We advise implementing stricter 48-hour event-lock-in policies with hosts to ensure ICT personnel are not setting up complex audiovisual environments for events that ultimately fail to materialize."
                    )
                    
                    doc.add_heading('3. Venue Strain and Coordination Burden', level=1)
                    doc.add_paragraph(
                        f"Data indicates that '{top_loc}' was the most heavily utilized venue for hardware installations. "
                        f"Because this location is experiencing exceptionally high physical wear-and-tear on deployed projectors, PA systems, and cabling, we strongly urge organizational leadership to fund permanent, built-in audiovisual installations at '{top_loc}'. This shift from temporary, reactive setups to permanent infrastructure will immediately reclaim massive amounts of IT labor time. "
                        f"\n\nFinally, the human coordination metric shows {top_coord} at the forefront of event management. Relying extensively on a single staff member for high-stress, public-facing setups is a critical operational risk. Cross-departmental shadowing is mandatory moving forward."
                    )

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
    st.markdown("Generate a unified, highly descriptive executive report merging analytics from **Departmental Activities** and **Events**.")

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

        if st.button("📄 Generate Narrative Master Report"):
            
            df_ict['Date_Obj'] = pd.to_datetime(df_ict['Date'], errors='coerce').dt.date
            df_events['Date_Obj'] = pd.to_datetime(df_events['Date'], errors='coerce').dt.date
            
            mask_ict = (df_ict['Date_Obj'] >= start_date) & (df_ict['Date_Obj'] <= end_date)
            mask_evt = (df_events['Date_Obj'] >= start_date) & (df_events['Date_Obj'] <= end_date)
            
            rep_ict = df_ict[mask_ict]
            rep_evt = df_events[mask_evt]

            if rep_ict.empty and rep_evt.empty:
                st.error(f"Comprehensive Database Scan: Absolutely no data was found in either the Activities or Events databases between {start_date} and {end_date}.")
            else:
                staff_act = explode_staff_counts(rep_ict, "IT Staff")
                staff_evt = explode_staff_counts(rep_evt, "Coordinator")
                
                combined_staff = pd.concat([staff_act, staff_evt], axis=1).fillna(0)
                combined_staff.columns = ['Activities', 'Events']
                combined_staff['Total Tasks'] = combined_staff['Activities'] + combined_staff['Events']
                combined_staff = combined_staff.sort_values(by='Total Tasks', ascending=False)
                
                # Merge dates to find overall idle times
                all_recorded_dates = pd.concat([rep_ict['Date'], rep_evt['Date']])
                idle_narrative = get_idle_dates_narrative(start_date, end_date, all_recorded_dates)

                doc = Document()
                doc.add_heading('Consolidated Executive ICT Operations & Strategic Insights Report', 0)
                doc.add_paragraph(f"Audited Reporting Period: {start_date} to {end_date}\nGenerated by Systems Intelligence on: {datetime.now().strftime('%B %d, %Y')}")
                
                doc.add_heading('Section 1: Executive Overview & Macro-Operational Health', level=1)
                doc.add_paragraph(
                    f"During the extensive evaluation period spanning {start_date} to {end_date}, the Information and Communication Technology (ICT) department served as the critical operational backbone, executing a total of {len(rep_ict) + len(rep_evt)} core structural operations. "
                    f"This massive organizational footprint breaks down into {len(rep_ict)} highly specific, departmental technical support interventions, running parallel to the complex logistical facilitation of {len(rep_evt)} specialized organizational events. "
                    f"By successfully mitigating both routine technological degradation and managing high-stakes public deployments, the ICT department has directly preserved massive amounts of organizational capital and continuous productivity."
                )
                
                doc.add_heading('Section 2: Timeline Analysis and Strategic Idle Periods', level=1)
                doc.add_paragraph(
                    f"An exhaustive chronological mapping of the department's workflow reveals highly distinct rhythms in operational demands. {idle_narrative} "
                    f"Ignoring these cyclical lulls invites technical stagnation. Management is heavily advised to formalize a 'Dark Day' protocol, wherein days showing zero immediate support/event requests trigger automatic, mandatory backend server audits and preventative physical hardware cleaning regimens."
                )

                if not combined_staff.empty:
                    fig_work, ax_work = plt.subplots(figsize=(7, 4))
                    combined_staff[['Activities', 'Events']].head(10).plot(kind='bar', stacked=True, color=['#4dabf7', '#ffc107'], ax=ax_work)
                    plt.title("Consolidated Staff Operational Workload")
                    plt.xlabel("ICT Personnel")
                    plt.ylabel("Number of Complete Operations (Tasks + Events)")
                    plt.xticks(rotation=45, ha='right')
                    fig_work.savefig("m_work.png", bbox_inches='tight')
                    plt.close(fig_work)

                    doc.add_heading('Section 3: Unified Personnel Workload & Bottleneck Vulnerabilities', level=1)
                    doc.add_picture("m_work.png", width=Inches(5.5))
                    
                    top_performer = combined_staff.index[0].title()
                    doc.add_paragraph(
                        f"The visual data clearly establishes that {top_performer} absorbed the heaviest and most complex operational burden, aggressively managing workflows across both micro-support tickets and macro-event coordination spaces. "
                        f"While this output is highly commendable, relying on a solitary \"lynchpin\" staff member creates an incredibly dangerous point of failure for the entire department. Should {top_performer} require sudden medical leave or transition to a different role, institutional knowledge and task momentum will severely collapse. "
                        f"Strategic Mandate: We must immediately initiate a forced cross-training protocol, effectively distributing {top_performer}'s core event and support duties across junior and mid-level technicians to flatten this dangerously skewed workload distribution."
                    )
                    if os.path.exists("m_work.png"): os.remove("m_work.png")

                ict_resolved = len(rep_ict[rep_ict['Status'] == 'resolved']) if not rep_ict.empty else 0
                evt_completed = len(rep_evt[rep_evt['Status'] == 'completed']) if not rep_evt.empty else 0
                total_success = ict_resolved + evt_completed
                total_ops = len(rep_ict) + len(rep_evt)
                success_rate = (total_success / total_ops) * 100 if total_ops > 0 else 0

                doc.add_heading('Section 4: Global Success Metrics and Final Recommendations', level=1)
                doc.add_paragraph(
                    f"Aggregating all data streams yields a final, consolidated global success and completion rate of {success_rate:.1f}% for the ICT department across this audit window. "
                    f"This percentage reflects every completely resolved hardware/software emergency and every flawlessly executed event setup. "
                    f"To push this metric closer to perfection in the upcoming quarters, management must immediately transition from a 'reactive troubleshooting' mindset to a 'proactive infrastructure upgrading' philosophy. Targeting our most recurring physical problems and standardizing the hardware deployed at our most utilized venues will drastically reduce future emergency intervention requirements."
                )

                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)
                
                st.success("✅ Deep Narrative Master Consolidated Report Generated!")
                st.download_button("📥 Download Narrative Master Report (.docx)", doc_buffer, f"ICT_Master_Narrative_{end_date}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
