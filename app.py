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
import random

# --- Configuration ---
st.set_page_config(page_title="ICT Department Portal", layout="wide")
GITHUB_REPO = "motechbello1/ICT-Report-Portal" 

# File paths
ICT_DATA_FILE = "ict_master_log.csv"
EVENT_DATA_FILE = "ict_events_log.csv"
TRACKER_FILE = "last_report_date.txt"

# --- Dynamic Narrative Engine ---
# These functions randomly select and combine different phrasing so the report never looks exactly the same twice.

def get_dynamic_intro(start, end, total, resolved, pending, rate, report_type="ICT"):
    if report_type == "ICT":
        intros = [
            f"During the audited window spanning {start} to {end}, the ICT technical support division was engaged in {total} recorded troubleshooting and deployment activities. The team drove a {rate:.1f}% resolution rate, effectively closing {resolved} distinct incident tickets, leaving {pending} tasks pending.",
            f"This executive analysis covers operations from {start} to {end}. Over this timeline, {total} technical incidents were actively managed. Our engineers successfully remediated {resolved} of these requests (an operational efficiency of {rate:.1f}%), while {pending} items have been rolled over for ongoing investigation.",
            f"Analyzing the operational timeframe of {start} through {end}, the ICT department handled {total} discrete service requests. With {resolved} critical issues fully resolved and {pending} in the active queue, the department maintained a completion metric of {rate:.1f}%."
        ]
    elif report_type == "Event":
        intros = [
            f"Between {start} and {end}, the ICT logistics team engineered the technological frameworks for {total} organizational events. Crucially, the aggregate audience size relying on our infrastructure was estimated at {resolved} attendees.",
            f"This logistical report audits the timeframe from {start} to {end}. Our technicians were deployed to support {total} distinct events, facilitating uninterrupted audio-visual and network services for approximately {resolved} total participants.",
            f"Evaluating the period from {start} to {end}, ICT personnel executed complex technical setups for {total} scheduled gatherings. These deployments successfully supported a combined audience volume of {resolved} individuals."
        ]
    return random.choice(intros)

def get_dynamic_problem_analysis(top_prob, count):
    analyses = [
        f"A deep analytical dive reveals that '{top_prob}' is currently the most pervasive vulnerability, accounting for {count} separate emergency interventions. This frequency strongly indicates a systemic decay rather than isolated user error.",
        f"The data highlights '{top_prob}' as the primary disruption to daily workflows, triggering {count} distinct support tickets. Immediate managerial review is recommended to determine if hardware replacement or network-wide patching is required.",
        f"Diagnostically, '{top_prob}' emerged as the dominant technical failure this period, demanding intervention {count} times. We heavily advise addressing this root cause strategically to eliminate this recurring bottleneck."
    ]
    return random.choice(analyses)

def get_dynamic_staff_analysis(top_staff, type="support"):
    if type == "support":
        analyses = [
            f"Internally, {top_staff} executed the highest volume of ticket resolutions. Monitoring this metric ensures fair task delegation moving forward.",
            f"Personnel analytics show {top_staff} absorbed the heaviest troubleshooting workload this cycle. We must ensure this output is sustainable and consider cross-training others.",
            f"Operational tracking indicates {top_staff} was the primary responder for the majority of technical incidents. Distributing this load remains a key priority."
        ]
    else:
        analyses = [
            f"The human coordination metric identifies {top_staff} at the forefront of event management. Relying heavily on one coordinator poses an operational risk requiring load balancing.",
            f"Logistical data shows {top_staff} anchored the highest number of physical event deployments. Expanding our pool of trained AV coordinators is highly recommended.",
            f"Reviewing staff deployment, {top_staff} handled the most complex and frequent event schedules. Shadowing programs should be implemented to assist them."
        ]
    return random.choice(analyses)

def get_idle_dates_narrative(start_date, end_date, active_dates_series):
    active_dates = set(pd.to_datetime(active_dates_series).dt.date)
    delta = end_date - start_date
    all_dates = {start_date + timedelta(days=i) for i in range(delta.days + 1)}
    missing_dates = sorted(list(all_dates - active_dates))
    
    if not missing_dates:
        return "Operational tempo remained continuously high; activities were recorded on every single day within this reporting window."
    
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
            text_parts.append(g[0].strftime("%b %d"))
        else:
            text_parts.append(f"{g[0].strftime('%b %d')} to {g[-1].strftime('%b %d')}")
            
    date_str = ", ".join(text_parts)
    
    idle_phrases = [
        f"Notably, zero operational incidents or requests were logged on the following dates: {date_str}. These strategic lulls should be utilized for server maintenance.",
        f"The database reflects no emergency calls or event setups on {date_str}. Management is advised to schedule preventative physical hardware cleaning during these exact windows.",
        f"We recorded a complete pause in reactive support requirements on {date_str}. These idle periods represent vital opportunities for internal staff training."
    ]
    return random.choice(idle_phrases)

# --- GitHub Sync & Data Loaders ---
def push_to_github(file_path):
    try:
        if "GITHUB_TOKEN" in st.secrets:
            g = Github(st.secrets["GITHUB_TOKEN"])
            repo = g.get_repo(GITHUB_REPO)
            with open(file_path, 'r') as file: content = file.read()
            try:
                contents = repo.get_contents(file_path)
                repo.update_file(contents.path, f"Auto-sync {file_path}", content, contents.sha)
                return True, "Synced to GitHub successfully!"
            except:
                repo.create_file(file_path, f"Auto-sync created {file_path}", content)
                return True, "Created and synced to GitHub successfully!"
        return False, "GITHUB_TOKEN not found. Saved locally."
    except Exception as e: return False, f"GitHub Sync Error: {str(e)}"

def load_and_clean_data(file_path, columns):
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        df = pd.read_csv(file_path)
        for col in df.columns:
            if col != "Attendee Count": 
                df[col] = df[col].astype(str).str.strip().str.lower()
            else:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        return df
    df = pd.DataFrame(columns=columns)
    df.to_csv(file_path, index=False)
    return df

def explode_staff_counts(df, column_name):
    if df.empty or column_name not in df.columns: return pd.Series(dtype=int)
    return df[column_name].dropna().astype(str).str.split(',').explode().str.strip().value_counts()

def get_last_report_date():
    if os.path.exists(TRACKER_FILE):
        with open(TRACKER_FILE, 'r') as f:
            try: return datetime.strptime(f.read().strip(), "%Y-%m-%d").date()
            except: pass
    return (datetime.now() - timedelta(days=7)).date()

def save_last_report_date(end_date):
    with open(TRACKER_FILE, 'w') as f: f.write(end_date.strftime("%Y-%m-%d"))

# ==========================================
# UI & NAVIGATION
# ==========================================
st.sidebar.title("🎛️ System Navigation")
app_mode = st.sidebar.radio("Select Portal Module:", ["🛠️ ICT Departmental Activities", "📅 ICT Events Dashboard", "📑 Master Combined Report"])
st.sidebar.markdown("---")

if "form_message" in st.session_state:
    if st.session_state.message_type == "success": st.success(st.session_state.form_message)
    else: st.warning(st.session_state.form_message)
    del st.session_state.form_message
    del st.session_state.message_type

# ==============================================================================
# MODULE 1: ICT DEPARTMENTAL ACTIVITIES
# ==============================================================================
if app_mode == "🛠️ ICT Departmental Activities":
    st.title("NBTI ICT Departmental Activities Portal")
    
    ict_columns = ["Date", "Day", "Time", "Department", "Reported By", "System ID", "Description", "Problem", "Action", "Parts", "IT Staff", "Status", "Remarks"]
    df_ict = load_and_clean_data(ICT_DATA_FILE, ict_columns)

    tab1, tab2, tab3, tab4 = st.tabs(["📝 Log Support Activity", "📊 Analytics Dashboard", "🗄️ Master Database", "📑 Executive Smart Report"])

    with tab1:
        with st.form("ict_log_form", clear_on_submit=True):
            t_col1, t_col2 = st.columns(2)
            with t_col1: log_date = st.date_input("Date", value=datetime.today())
            with t_col2: log_time = st.time_input("Time", value=datetime.now().time())
            
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
        if not df_ict.empty:
            sort_by = st.selectbox("Analyze Metric:", ["Status", "Department", "IT Staff", "Problem"])
            col_name = "Problem" if sort_by == "Problem" else sort_by
            counts = explode_staff_counts(df_ict, col_name).reset_index() if col_name == "IT Staff" else df_ict[col_name].value_counts().reset_index()
            counts.columns = [col_name, "Count"]
            fig = px.pie(counts, values="Count", names=col_name) if col_name == "Status" else px.bar(counts, x=col_name, y="Count", color=col_name)
            st.plotly_chart(fig, use_container_width=True)

    with tab3:
        if not df_ict.empty: st.download_button("📥 Download CSV", df_ict.to_csv(index=False).encode('utf-8'), f"ICT_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")
        st.dataframe(df_ict, use_container_width=True)

    with tab4:
        st.subheader("Generate Dynamic Action Report")
        col_d1, col_d2 = st.columns(2)
        with col_d1: start_date = st.date_input("Start Date", value=get_last_report_date())
        with col_d2: end_date = st.date_input("End Date", value=get_last_report_date() + timedelta(days=7))

        if st.button("📄 Generate Report"):
            temp_df = df_ict.copy()
            temp_df['Date_Obj'] = pd.to_datetime(temp_df['Date'], errors='coerce').dt.date
            df_w = temp_df[(temp_df['Date_Obj'] >= start_date) & (temp_df['Date_Obj'] <= end_date)].copy()
            
            if df_w.empty: st.warning("No records found.")
            else:
                total_issues = len(df_w)
                resolved = len(df_w[df_w['Status'] == 'resolved'])
                pending = total_issues - resolved
                res_rate = (resolved / total_issues) * 100 if total_issues > 0 else 0
                
                doc = Document()
                doc.add_heading('ICT Operational Health Report', 0)
                doc.add_paragraph(f"Generated on: {datetime.now().strftime('%B %d, %Y')}")
                
                # Dynamic Narratives
                doc.add_heading('1. Executive Tempo', level=1)
                doc.add_paragraph(get_dynamic_intro(start_date, end_date, total_issues, resolved, pending, res_rate, "ICT"))
                doc.add_paragraph(get_idle_dates_narrative(start_date, end_date, df_w['Date']))

                # MULTIPLE INFOGRAPHICS
                # Graphic 1: Status Pie
                fig1, ax1 = plt.subplots(figsize=(4, 3))
                df_w['Status'].value_counts().plot(kind='pie', autopct='%1.1f%%', ax=ax1, colors=['#28a745', '#dc3545'])
                ax1.set_ylabel("")
                plt.title("Resolution Rate")
                fig1.savefig("t_stat.png", bbox_inches='tight')
                plt.close(fig1)
                
                # Graphic 2: Top Issues Bar
                fig2, ax2 = plt.subplots(figsize=(5, 3))
                prob_counts = df_w['Problem'].value_counts().head(5)
                prob_counts.plot(kind='barh', color='#ff6b6b', ax=ax2)
                ax2.invert_yaxis()
                plt.title("Systemic Complaints")
                fig2.savefig("t_prob.png", bbox_inches='tight')
                plt.close(fig2)

                doc.add_heading('2. Diagnostic Breakdown', level=1)
                doc.add_picture("t_stat.png", width=Inches(3.0))
                doc.add_picture("t_prob.png", width=Inches(4.5))
                doc.add_paragraph(get_dynamic_problem_analysis(prob_counts.index[0].title(), prob_counts.iloc[0]))

                # Graphic 3: Department Demand
                fig3, ax3 = plt.subplots(figsize=(5, 3))
                df_w['Department'].value_counts().head(5).plot(kind='bar', color='#4dabf7', ax=ax3)
                plt.xticks(rotation=45, ha='right')
                plt.title("Departmental IT Demand")
                fig3.savefig("t_dept.png", bbox_inches='tight')
                plt.close(fig3)

                doc.add_heading('3. Resource Allocation', level=1)
                doc.add_picture("t_dept.png", width=Inches(4.5))
                doc.add_paragraph(get_dynamic_staff_analysis(explode_staff_counts(df_w, "IT Staff").index[0].title(), "support"))

                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)
                for f in ["t_stat.png", "t_prob.png", "t_dept.png"]: 
                    if os.path.exists(f): os.remove(f)
                save_last_report_date(end_date)
                
                st.success("✅ Multi-Graphic Report Generated!")
                st.download_button("📥 Download Document", doc_buffer, f"ICT_{end_date}.docx")

# ==============================================================================
# MODULE 2: ICT EVENTS (Abridged for brevity - identical logic to above)
# ==============================================================================
elif app_mode == "📅 ICT Events Dashboard":
    st.title("🎟️ ICT Event Activity Dashboard")
    # ... (Form and Data loading identical to previous code)
    event_columns = ["Date", "Day", "Time", "Event Name", "Location", "Coordinator", "Equipment Deployed", "Attendee Count", "Status", "Remarks"]
    df_events = load_and_clean_data(EVENT_DATA_FILE, event_columns)

    etab1, etab2, etab3 = st.tabs(["📝 Log Event Activity", "🗄️ Database", "📑 Generate Report"])
    with etab1:
        with st.form("event_log_form", clear_on_submit=True):
            t_col1, t_col2 = st.columns(2)
            with t_col1: log_date = st.date_input("Date", value=datetime.today())
            with t_col2: log_time = st.time_input("Time", value=datetime.now().time())
            c1, c2 = st.columns(2)
            with c1:
                event_name = st.text_input("Event Name")
                location = st.text_input("Location")
                coordinator = st.text_input("Coordinator(s)")
                equipment = st.text_area("Equipment")
            with c2:
                attendees = st.number_input("Attendees", min_value=0)
                status = st.selectbox("Status", ["planned", "completed", "cancelled"]) 
                remarks = st.text_area("Remarks")
            if st.form_submit_button("💾 Save Event"):
                new_event = {"Date": log_date.strftime("%Y-%m-%d"), "Day": log_date.strftime("%A").lower(), "Time": log_time.strftime("%H:%M:%S"), "Event Name": event_name.lower(), "Location": location.lower(), "Coordinator": coordinator.lower(), "Equipment Deployed": equipment.lower(), "Attendee Count": attendees, "Status": status, "Remarks": remarks.lower()}
                temp_df = pd.concat([df_events, pd.DataFrame([new_event])], ignore_index=True)
                temp_df.to_csv(EVENT_DATA_FILE, index=False)
                st.rerun()
                
    with etab2: st.dataframe(df_events, use_container_width=True)
    with etab3:
        st.subheader("Generate Dynamic Event Report")
        col_d1, col_d2 = st.columns(2)
        with col_d1: start_date = st.date_input("Start Date", value=get_last_report_date())
        with col_d2: end_date = st.date_input("End Date", value=get_last_report_date() + timedelta(days=7))

        if st.button("📄 Generate Event Report"):
            temp_df = df_events.copy()
            temp_df['Date_Obj'] = pd.to_datetime(temp_df['Date'], errors='coerce').dt.date
            df_e = temp_df[(temp_df['Date_Obj'] >= start_date) & (temp_df['Date_Obj'] <= end_date)].copy()
            
            if not df_e.empty:
                doc = Document()
                doc.add_heading('ICT Event Logistics Report', 0)
                
                doc.add_heading('1. Logistics Overview', level=1)
                doc.add_paragraph(get_dynamic_intro(start_date, end_date, len(df_e), df_e['Attendee Count'].sum(), 0, 0, "Event"))
                doc.add_paragraph(get_idle_dates_narrative(start_date, end_date, df_e['Date']))
                
                # GRAPHIC 1: Locations
                fig1, ax1 = plt.subplots(figsize=(5, 3))
                df_e['Location'].value_counts().plot(kind='bar', color='#17a2b8', ax=ax1)
                plt.xticks(rotation=45, ha='right')
                plt.title("Venue Utilization Map")
                fig1.savefig("e_loc.png", bbox_inches='tight')
                plt.close(fig1)

                # GRAPHIC 2: Coordinators
                fig2, ax2 = plt.subplots(figsize=(5, 3))
                explode_staff_counts(df_e, "Coordinator").plot(kind='pie', ax=ax2, autopct='%1.0f%%')
                ax2.set_ylabel("")
                plt.title("Coordinator Workload")
                fig2.savefig("e_coord.png", bbox_inches='tight')
                plt.close(fig2)

                doc.add_heading('2. Venue & Personnel Analysis', level=1)
                doc.add_picture("e_loc.png", width=Inches(4.5))
                doc.add_picture("e_coord.png", width=Inches(3.5))
                doc.add_paragraph(get_dynamic_staff_analysis(explode_staff_counts(df_e, "Coordinator").index[0].title(), "event"))

                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)
                for f in ["e_loc.png", "e_coord.png"]: 
                    if os.path.exists(f): os.remove(f)
                
                st.success("✅ Multi-Graphic Report Generated!")
                st.download_button("📥 Download Report", doc_buffer, f"Events_{end_date}.docx")

# ==============================================================================
# MODULE 3: MASTER COMBINED REPORT
# ==============================================================================
elif app_mode == "📑 Master Combined Report":
    st.title("📑 Master ICT Consolidated Report")
    
    ict_columns = ["Date", "Day", "Time", "Department", "Reported By", "System ID", "Description", "Problem", "Action", "Parts", "IT Staff", "Status", "Remarks"]
    event_columns = ["Date", "Day", "Time", "Event Name", "Location", "Coordinator", "Equipment Deployed", "Attendee Count", "Status", "Remarks"]
    df_ict = load_and_clean_data(ICT_DATA_FILE, ict_columns)
    df_events = load_and_clean_data(EVENT_DATA_FILE, event_columns)

    col_d1, col_d2 = st.columns(2)
    with col_d1: start_date = st.date_input("Start Date", value=get_last_report_date())
    with col_d2: end_date = st.date_input("End Date", value=get_last_report_date() + timedelta(days=7))

    if st.button("📄 Generate Narrative Master Report"):
        df_ict['Date_Obj'] = pd.to_datetime(df_ict['Date'], errors='coerce').dt.date
        df_events['Date_Obj'] = pd.to_datetime(df_events['Date'], errors='coerce').dt.date
        
        rep_ict = df_ict[(df_ict['Date_Obj'] >= start_date) & (df_ict['Date_Obj'] <= end_date)]
        rep_evt = df_events[(df_events['Date_Obj'] >= start_date) & (df_events['Date_Obj'] <= end_date)]

        if not rep_ict.empty or not rep_evt.empty:
            staff_act = explode_staff_counts(rep_ict, "IT Staff")
            staff_evt = explode_staff_counts(rep_evt, "Coordinator")
            combined_staff = pd.concat([staff_act, staff_evt], axis=1).fillna(0)
            combined_staff.columns = ['Activities', 'Events']
            combined_staff['Total'] = combined_staff['Activities'] + combined_staff['Events']
            combined_staff = combined_staff.sort_values(by='Total', ascending=False)
            
            doc = Document()
            doc.add_heading('Consolidated Executive ICT Operations', 0)
            
            # Dynamic Intro
            intros_master = [
                f"This unified report merges {len(rep_ict)} daily technical interventions with the logistical coordination of {len(rep_evt)} organizational events.",
                f"Aggregating cross-departmental data, the ICT team successfully managed {len(rep_ict)} support tickets alongside {len(rep_evt)} full-scale event setups.",
                f"From {start_date} to {end_date}, operational output peaked at a combined total of {len(rep_ict) + len(rep_evt)} discrete operational missions."
            ]
            doc.add_heading('Section 1: Macro-Operational Health', level=1)
            doc.add_paragraph(random.choice(intros_master))
            doc.add_paragraph(get_idle_dates_narrative(start_date, end_date, pd.concat([rep_ict['Date'], rep_evt['Date']])))

            # GRAPHIC 1: Stacked Bar Workload
            if not combined_staff.empty:
                fig_work, ax_work = plt.subplots(figsize=(6, 4))
                combined_staff[['Activities', 'Events']].head(7).plot(kind='bar', stacked=True, color=['#4dabf7', '#ffc107'], ax=ax_work)
                plt.title("Unified Personnel Workload")
                plt.xticks(rotation=45, ha='right')
                fig_work.savefig("m_work.png", bbox_inches='tight')
                plt.close(fig_work)

            # GRAPHIC 2: Daily Tempo Line Chart
            tempo_df = pd.DataFrame({
                'Support': rep_ict['Date'].value_counts(),
                'Events': rep_evt['Date'].value_counts()
            }).fillna(0).sort_index()
            
            fig_line, ax_line = plt.subplots(figsize=(6, 3))
            tempo_df.plot(kind='line', marker='o', ax=ax_line, color=['#e83e8c', '#fd7e14'])
            plt.title("Daily Operational Tempo")
            plt.xticks(rotation=45, ha='right')
            fig_line.savefig("m_line.png", bbox_inches='tight')
            plt.close(fig_line)

            doc.add_heading('Section 2: Tempo & Personnel Tracking', level=1)
            doc.add_picture("m_line.png", width=Inches(5.0))
            doc.add_paragraph("The trendline above tracks the exact volume of incoming requests and event dates, mapping our daily operational pressure points.")
            
            if not combined_staff.empty:
                doc.add_picture("m_work.png", width=Inches(5.0))
                top_staff = combined_staff.index[0].title()
                staff_phrases = [
                    f"Visually, {top_staff} absorbed the heaviest cross-disciplinary load. Relying on one technician for both support and events is an urgent bottleneck.",
                    f"The data flags {top_staff} as the most heavily burdened staff member across both domains. To prevent burnout, tasks must be redistributed immediately.",
                    f"Analysis proves {top_staff} handled a disproportionate amount of total department operations. Cross-training juniors to relieve this load is mandatory."
                ]
                doc.add_paragraph(random.choice(staff_phrases))

            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            for f in ["m_work.png", "m_line.png"]: 
                if os.path.exists(f): os.remove(f)
            
            st.success("✅ Multi-Graphic Dynamic Report Generated!")
            st.download_button("📥 Download Master Report", doc_buffer, f"Master_{end_date}.docx")

