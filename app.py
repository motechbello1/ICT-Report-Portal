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

ICT_DATA_FILE = "ict_master_log.csv"
EVENT_DATA_FILE = "ict_events_log.csv"
TRACKER_FILE = "last_report_date.txt"

# --- Define the 7 Enterprise Sectors ---
ICT_SECTORS = [
    "Cybersecurity & Access Control", 
    "Network Infrastructure", 
    "Cloud & Server Operations", 
    "Hardware & Asset Management", 
    "Software & Applications", 
    "Database Management", 
    "Telecommunications & VoIP"
]

EVENT_SECTORS = [
    "Audio Engineering", 
    "Visual & Projection Systems", 
    "Staging & Power Distribution", 
    "Venue Network Provisioning", 
    "Vendor Logistics", 
    "Crowd & Access Control", 
    "General Coordination"
]

# ==========================================
# DYNAMIC NARRATIVE ENGINE (DEEP PARAGRAPHS)
# ==========================================

def get_dynamic_intro(start, end, total, resolved, pending, rate, report_type="ICT"):
    if report_type == "ICT":
        intros = [
            f"During the comprehensive reporting window spanning {start} to {end}, the Information and Communication Technology (ICT) technical support division was engaged in a high volume of crucial diagnostic and remediation activities. The department actively managed {total} distinct support tickets encompassing network failures, hardware degradation, and software troubleshooting. The engineering team successfully drove a {rate:.1f}% operational resolution rate, fully closing {resolved} of these critical incidents. However, {pending} complex tasks remain in the active queue, requiring rollover prioritization in the immediate future to prevent cascading disruptions to daily organizational workflows.",
            f"This executive operational audit evaluates the structural health and responsiveness of the ICT department from {start} to {end}. Over this critical timeline, our infrastructure absorbed and processed {total} technical service requests across various organizational wings. Demonstrating strong technical agility, the support team achieved a {rate:.1f}% completion metric, successfully remediating {resolved} system-halting issues. The remaining {pending} unresolved items have been documented and escalated for advanced troubleshooting. Maintaining this high rate of resolution is paramount to ensuring that baseline administrative continuity is never compromised by technological bottlenecks.",
        ]
    elif report_type == "Event":
        intros = [
            f"Between the operational dates of {start} and {end}, the ICT logistics and deployment team engineered the complex technological frameworks required for {total} distinct organizational events. Crucially, the aggregate audience size that relied directly on our uninterrupted audio-visual, networking, and presentation infrastructure was estimated at {resolved} attendees. Successfully managing public-facing deployments of this scale requires extreme precision, rapid on-site troubleshooting, and exhaustive pre-event preparation to safeguard the organization's professional image.",
        ]
    elif report_type == "Master":
        intros = [
            f"This unified executive operational report merges the comprehensive data streams from both daily technical interventions and high-stakes logistical event coordination. Analyzing the period from {start} to {end}, the ICT department successfully managed a massive operational footprint consisting of {resolved} core diagnostic support tickets running concurrently alongside {pending} full-scale event infrastructure setups. By aggressively mitigating routine technological degradation while simultaneously facilitating major organizational gatherings, the ICT department has directly preserved massive amounts of institutional capital, productivity, and operational momentum.",
        ]
    return random.choice(intros)

def get_sector_analysis(sector, total_sector_ops, resolved, pending):
    if total_sector_ops == 0:
        zero_phrases = [
            f"Analysis of the '{sector}' sector indicates absolute operational stability during this reporting cycle. Zero critical failure incidents or emergency support requests were logged. This uninterrupted 100% uptime reflects the success of our recent preventative maintenance protocols and the baseline durability of this specific infrastructure pillar.",
            f"The telemetry for the '{sector}' domain shows no registered anomalies, support tickets, or logistical bottlenecks for this operational window. Maintaining zero incident velocity in this sector allows our engineering staff to safely redirect human resources toward proactive system upgrades rather than reactive troubleshooting.",
            f"A deep audit of the '{sector}' vertical reveals completely optimal functioning. With zero reported defects or workflow interruptions, this sector successfully operated autonomously. Management should view this prolonged stability as an indicator of robust system health."
        ]
        return random.choice(zero_phrases)
    else:
        rate = (resolved / total_sector_ops) * 100
        active_phrases = [
            f"The '{sector}' sector was highly active during this cycle, generating {total_sector_ops} distinct operational demands. Our technicians successfully closed {resolved} of these items, achieving a {rate:.1f}% resolution efficiency within this specific domain. The {pending} tasks currently rolling over will require dedicated focus in the coming week to ensure this sector does not become an organizational bottleneck.",
            f"Auditing the '{sector}' vertical reveals a significant workflow load, with {total_sector_ops} isolated incidents or requests requiring direct intervention. The engineering and logistics teams mitigated {resolved} of these issues directly. With {pending} remaining unresolved, we must carefully monitor this sector's strain matrix to determine if additional capital investment or specialized staff training is required to permanently stabilize it.",
            f"Evaluating the structural demands placed on the '{sector}' domain, the data shows {total_sector_ops} active deployments or system failures. Through aggressive technical routing, {resolved} items were fully remediated (a {rate:.1f}% clearance rate). The remaining {pending} open tickets in this sector have been flagged for executive review, as continuous strain on this pillar directly impacts overarching network health."
        ]
        return random.choice(active_phrases)

def get_idle_dates_narrative(start_date, end_date, active_dates_series):
    active_dates = set(pd.to_datetime(active_dates_series).dt.date)
    delta = end_date - start_date
    all_dates = {start_date + timedelta(days=i) for i in range(delta.days + 1)}
    missing_dates = sorted(list(all_dates - active_dates))
    
    if not missing_dates:
        return "It is highly notable that operational tempo remained continuously aggressive; activities, emergencies, or deployments were actively recorded on every single calendar day within this reporting window. This indicates a heavily maximized operational environment."
    
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
            text_parts.append(f"the period stretching from {g[0].strftime('%B %d')} to {g[-1].strftime('%B %d, %Y')}")
            
    date_str = ", ".join(text_parts)
    return f"A chronological audit of the database reveals distinct operational lulls; specifically, absolutely zero emergency calls, technical interventions, or event setups were logged on {date_str}. Rather than viewing these as inactive days, management is strongly advised to officially designate these exact windows as 'Dark Days'—mandating preventative physical hardware cleaning, deep-level server audits, and automated firmware patching when organizational disruption risk is at its absolute lowest."

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
                return True, "Synced successfully!"
            except:
                repo.create_file(file_path, f"Auto-sync created {file_path}", content)
                return True, "Created and synced successfully!"
        return False, "GITHUB_TOKEN not found. Saved locally."
    except Exception as e: return False, f"GitHub Sync Error: {str(e)}"

def load_and_clean_data(file_path, columns):
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        df = pd.read_csv(file_path)
        # Backward compatibility for old databases missing the Sector column
        if "Sector" not in df.columns:
            df["Sector"] = "Unassigned / General"
        for col in df.columns:
            if col != "Attendee Count": 
                df[col] = df[col].astype(str).str.strip().str.title()
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
st.sidebar.title("System Navigation")
app_mode = st.sidebar.radio("Select Portal Module:", ["ICT Departmental Activities", "ICT Events Dashboard", "Master Combined Report"])
st.sidebar.markdown("---")

# ==============================================================================
# MODULE 1: ICT DEPARTMENTAL ACTIVITIES
# ==============================================================================
if app_mode == "ICT Departmental Activities":
    st.title("Executive ICT Departmental Activities Portal")
    
    ict_columns = ["Date", "Day", "Time", "Sector", "Department", "Reported By", "System ID", "Description", "Problem", "Action", "Parts", "IT Staff", "Status", "Remarks"]
    df_ict = load_and_clean_data(ICT_DATA_FILE, ict_columns)

    tab1, tab2, tab3 = st.tabs(["Log Support Activity", "Master Database", "Executive Smart Report"])

    with tab1:
        with st.form("ict_log_form", clear_on_submit=True):
            t_col1, t_col2 = st.columns(2)
            with t_col1: log_date = st.date_input("Date", value=datetime.today())
            with t_col2: log_time = st.time_input("Time", value=datetime.now().time())
            
            sector = st.selectbox("Assign to Technical Sector", ICT_SECTORS)
            
            col1, col2 = st.columns(2)
            with col1:
                department = st.text_input("Department Requiring Support")
                reported_by = st.text_input("Reported By")
                system_id = st.text_input("System ID")
                description = st.text_input("System Description")
                problem = st.text_area("Diagnosed Problem")
            with col2:
                action = st.text_area("Action Taken")
                parts = st.text_input("Replaced Parts (if any)")
                it_staff = st.text_input("IT Staff Name(s) - Separate with commas")
                status = st.selectbox("Status", ["Resolved", "Pending"]) 
                remarks = st.text_area("Remarks")
                
            if st.form_submit_button("Save Activity Log & Sync"):
                new_data = {
                    "Date": log_date.strftime("%Y-%m-%d"), "Day": log_date.strftime("%A"), "Time": log_time.strftime("%H:%M:%S"),
                    "Sector": sector, "Department": department, "Reported By": reported_by, "System ID": system_id, 
                    "Description": description, "Problem": problem, "Action": action,
                    "Parts": parts, "IT Staff": it_staff, "Status": status, "Remarks": remarks
                }
                temp_df = pd.concat([df_ict, pd.DataFrame([new_data])], ignore_index=True)
                temp_df = temp_df.sort_values(by=["Date", "Time"]).reset_index(drop=True)
                temp_df.to_csv(ICT_DATA_FILE, index=False)
                push_to_github(ICT_DATA_FILE)
                st.success("Data Saved Successfully.")
                st.rerun()

    with tab2:
        st.dataframe(df_ict, use_container_width=True)

    with tab3:
        st.subheader("Generate 7-Sector Dynamic Action Report")
        col_d1, col_d2 = st.columns(2)
        with col_d1: start_date = st.date_input("Start Date", value=get_last_report_date())
        with col_d2: end_date = st.date_input("End Date", value=get_last_report_date() + timedelta(days=7))

        if st.button("Generate Comprehensive Report"):
            temp_df = df_ict.copy()
            temp_df['Date_Obj'] = pd.to_datetime(temp_df['Date'], errors='coerce').dt.date
            df_w = temp_df[(temp_df['Date_Obj'] >= start_date) & (temp_df['Date_Obj'] <= end_date)].copy()
            
            total_issues = len(df_w)
            resolved = len(df_w[df_w['Status'] == 'Resolved'])
            pending = total_issues - resolved
            res_rate = (resolved / total_issues) * 100 if total_issues > 0 else 0
            
            doc = Document()
            doc.add_heading('Deep Executive ICT Operational Health Report', 0)
            doc.add_paragraph(f"Audited Timeline: {start_date} to {end_date}\nGenerated on: {datetime.now().strftime('%B %d, %Y')}")
            
            doc.add_heading('Section 1: Executive Summary & Global Tempo', level=1)
            doc.add_paragraph(get_dynamic_intro(start_date, end_date, total_issues, resolved, pending, res_rate, "ICT"))
            doc.add_paragraph(get_idle_dates_narrative(start_date, end_date, df_w['Date']))

            doc.add_heading('Section 2: Comprehensive 7-Sector Analytics', level=1)
            doc.add_paragraph("The following section breaks down departmental performance across our seven core technological pillars, identifying specific strain points and establishing a clear matrix of our operational stability.")

            img_paths = []
            
            for index, sector in enumerate(ICT_SECTORS, 1):
                doc.add_heading(f"2.{index} Domain: {sector}", level=2)
                sector_df = df_w[df_w['Sector'] == sector]
                sec_total = len(sector_df)
                sec_resolved = len(sector_df[sector_df['Status'] == 'Resolved'])
                sec_pending = sec_total - sec_resolved
                
                doc.add_paragraph(get_sector_analysis(sector, sec_total, sec_resolved, sec_pending))
                
                fig, ax = plt.subplots(figsize=(5, 2.5))
                if sec_total > 0:
                    sector_df['Status'].value_counts().plot(kind='barh', color=['#28a745', '#dc3545'], ax=ax)
                    plt.title(f"{sector} Resolution Status")
                else:
                    ax.text(0.5, 0.5, '100% Stability\nZero Incidents Logged', horizontalalignment='center', verticalalignment='center', fontsize=12, color='green')
                    ax.axis('off')
                    plt.title(f"{sector} Health Status")
                
                img_name = f"sec_ict_{index}.png"
                fig.savefig(img_name, bbox_inches='tight')
                plt.close(fig)
                doc.add_picture(img_name, width=Inches(4.5))
                img_paths.append(img_name)

            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            for f in img_paths: 
                if os.path.exists(f): os.remove(f)
            save_last_report_date(end_date)
            
            st.success("Deep Narrative 7-Sector Report Generated!")
            st.download_button("Download Executive Document", doc_buffer, f"ICT_Deep_Report_{end_date}.docx")

# ==============================================================================
# MODULE 2: ICT EVENTS DASHBOARD
# ==============================================================================
elif app_mode == "ICT Events Dashboard":
    st.title("ICT Event Activity Dashboard")
    event_columns = ["Date", "Day", "Time", "Sector", "Event Name", "Location", "Coordinator", "Equipment Deployed", "Attendee Count", "Status", "Remarks"]
    df_events = load_and_clean_data(EVENT_DATA_FILE, event_columns)

    etab1, etab2, etab3 = st.tabs(["Log Event Activity", "Database", "Generate Event Report"])
    with etab1:
        with st.form("event_log_form", clear_on_submit=True):
            t_col1, t_col2 = st.columns(2)
            with t_col1: log_date = st.date_input("Date", value=datetime.today())
            with t_col2: log_time = st.time_input("Time", value=datetime.now().time())
            
            sector = st.selectbox("Assign to Event Sector", EVENT_SECTORS)
            
            c1, c2 = st.columns(2)
            with c1:
                event_name = st.text_input("Event Name")
                location = st.text_input("Location")
                coordinator = st.text_input("Coordinator(s)")
                equipment = st.text_area("Equipment Deployed")
            with c2:
                attendees = st.number_input("Attendees", min_value=0)
                status = st.selectbox("Status", ["Planned", "Completed", "Cancelled"]) 
                remarks = st.text_area("Remarks")
            if st.form_submit_button("Save Event"):
                new_event = {"Date": log_date.strftime("%Y-%m-%d"), "Day": log_date.strftime("%A"), "Time": log_time.strftime("%H:%M:%S"), "Sector": sector, "Event Name": event_name, "Location": location, "Coordinator": coordinator, "Equipment Deployed": equipment, "Attendee Count": attendees, "Status": status, "Remarks": remarks}
                temp_df = pd.concat([df_events, pd.DataFrame([new_event])], ignore_index=True)
                temp_df.to_csv(EVENT_DATA_FILE, index=False)
                st.success("Event Saved Successfully.")
                st.rerun()
                
    with etab2: st.dataframe(df_events, use_container_width=True)
    
    with etab3:
        st.subheader("Generate Narrative Event Report")
        col_d1, col_d2 = st.columns(2)
        with col_d1: start_date = st.date_input("Start Date", value=get_last_report_date())
        with col_d2: end_date = st.date_input("End Date", value=get_last_report_date() + timedelta(days=7))

        if st.button("Generate Detailed Event Report"):
            temp_df = df_events.copy()
            temp_df['Date_Obj'] = pd.to_datetime(temp_df['Date'], errors='coerce').dt.date
            df_e = temp_df[(temp_df['Date_Obj'] >= start_date) & (temp_df['Date_Obj'] <= end_date)].copy()
            
            doc = Document()
            doc.add_heading('Comprehensive ICT Event Logistics Report', 0)
            doc.add_paragraph(f"Audited Timeline: {start_date} to {end_date}\nGenerated on: {datetime.now().strftime('%B %d, %Y')}")
            
            doc.add_heading('Section 1: Global Logistics & Impact Overview', level=1)
            doc.add_paragraph(get_dynamic_intro(start_date, end_date, len(df_e), df_e['Attendee Count'].sum() if not df_e.empty else 0, 0, 0, "Event"))
            
            doc.add_heading('Section 2: Comprehensive 7-Sector Event Logistics', level=1)
            img_paths = []
            
            for index, sector in enumerate(EVENT_SECTORS, 1):
                doc.add_heading(f"2.{index} Domain: {sector}", level=2)
                sector_df = df_e[df_e['Sector'] == sector] if not df_e.empty else pd.DataFrame()
                sec_total = len(sector_df)
                sec_completed = len(sector_df[sector_df['Status'] == 'Completed']) if not sector_df.empty else 0
                sec_pending = sec_total - sec_completed
                
                doc.add_paragraph(get_sector_analysis(sector, sec_total, sec_completed, sec_pending))
                
                fig, ax = plt.subplots(figsize=(5, 2.5))
                if sec_total > 0:
                    sector_df['Location'].value_counts().head(3).plot(kind='bar', color='#17a2b8', ax=ax)
                    plt.title(f"{sector} Deployment Locations")
                    plt.xticks(rotation=15, ha='right')
                else:
                    ax.text(0.5, 0.5, 'Zero Deployments Required', horizontalalignment='center', verticalalignment='center', fontsize=12, color='gray')
                    ax.axis('off')
                
                img_name = f"sec_evt_{index}.png"
                fig.savefig(img_name, bbox_inches='tight')
                plt.close(fig)
                doc.add_picture(img_name, width=Inches(4.0))
                img_paths.append(img_name)

            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            for f in img_paths: 
                if os.path.exists(f): os.remove(f)
            
            st.success("Deep Event Narrative Report Generated!")
            st.download_button("Download Report", doc_buffer, f"Events_Deep_{end_date}.docx")

# ==============================================================================
# MODULE 3: MASTER COMBINED REPORT
# ==============================================================================
elif app_mode == "Master Combined Report":
    st.title("Master ICT Consolidated Report")
    st.markdown("Generates a unified executive document merging all 14 analytical sectors (7 Technical + 7 Logistical) with workload mapping.")
    
    ict_columns = ["Date", "Day", "Time", "Sector", "Department", "Reported By", "System ID", "Description", "Problem", "Action", "Parts", "IT Staff", "Status", "Remarks"]
    event_columns = ["Date", "Day", "Time", "Sector", "Event Name", "Location", "Coordinator", "Equipment Deployed", "Attendee Count", "Status", "Remarks"]
    df_ict = load_and_clean_data(ICT_DATA_FILE, ict_columns)
    df_events = load_and_clean_data(EVENT_DATA_FILE, event_columns)

    col_d1, col_d2 = st.columns(2)
    with col_d1: start_date = st.date_input("Start Date", value=get_last_report_date())
    with col_d2: end_date = st.date_input("End Date", value=get_last_report_date() + timedelta(days=7))

    if st.button("Generate Deep Narrative Master Report"):
        df_ict['Date_Obj'] = pd.to_datetime(df_ict['Date'], errors='coerce').dt.date
        df_events['Date_Obj'] = pd.to_datetime(df_events['Date'], errors='coerce').dt.date
        
        rep_ict = df_ict[(df_ict['Date_Obj'] >= start_date) & (df_ict['Date_Obj'] <= end_date)]
        rep_evt = df_events[(df_events['Date_Obj'] >= start_date) & (df_events['Date_Obj'] <= end_date)]

        total_ops = len(rep_ict) + len(rep_evt)
        
        doc = Document()
        doc.add_heading('Consolidated Executive ICT Operations & Strategic Insights', 0)
        doc.add_paragraph(f"Audited Timeline: {start_date} to {end_date}\nGenerated by Systems Intelligence on: {datetime.now().strftime('%B %d, %Y')}")
        
        doc.add_heading('Section 1: Macro-Operational Health', level=1)
        doc.add_paragraph(get_dynamic_intro(start_date, end_date, total_ops, len(rep_ict), len(rep_evt), 0, "Master"))
        
        doc.add_heading('Section 2: Technical Infrastructure Domains (7 Pillars)', level=1)
        img_paths = []
        for index, sector in enumerate(ICT_SECTORS, 1):
            doc.add_heading(f"2.{index} {sector}", level=2)
            sector_df = rep_ict[rep_ict['Sector'] == sector] if not rep_ict.empty else pd.DataFrame()
            sec_total = len(sector_df)
            sec_resolved = len(sector_df[sector_df['Status'] == 'Resolved']) if not sector_df.empty else 0
            doc.add_paragraph(get_sector_analysis(sector, sec_total, sec_resolved, sec_total - sec_resolved))
            
            fig, ax = plt.subplots(figsize=(4, 2))
            if sec_total > 0:
                sector_df['Status'].value_counts().plot(kind='pie', ax=ax, colors=['#4dabf7', '#ff6b6b'])
                ax.set_ylabel('')
            else:
                ax.text(0.5, 0.5, 'Optimal Health', horizontalalignment='center', verticalalignment='center')
                ax.axis('off')
            img_name = f"m_ict_{index}.png"
            fig.savefig(img_name, bbox_inches='tight')
            plt.close(fig)
            doc.add_picture(img_name, width=Inches(3.0))
            img_paths.append(img_name)

        doc.add_heading('Section 3: Logistical Event Domains (7 Pillars)', level=1)
        for index, sector in enumerate(EVENT_SECTORS, 1):
            doc.add_heading(f"3.{index} {sector}", level=2)
            sector_df = rep_evt[rep_evt['Sector'] == sector] if not rep_evt.empty else pd.DataFrame()
            sec_total = len(sector_df)
            sec_resolved = len(sector_df[sector_df['Status'] == 'Completed']) if not sector_df.empty else 0
            doc.add_paragraph(get_sector_analysis(sector, sec_total, sec_resolved, sec_total - sec_resolved))

            fig, ax = plt.subplots(figsize=(4, 2))
            if sec_total > 0:
                sector_df['Status'].value_counts().plot(kind='bar', color='#ffc107', ax=ax)
                plt.xticks(rotation=0)
            else:
                ax.text(0.5, 0.5, 'Zero Logistical Strain', horizontalalignment='center', verticalalignment='center')
                ax.axis('off')
            img_name = f"m_evt_{index}.png"
            fig.savefig(img_name, bbox_inches='tight')
            plt.close(fig)
            doc.add_picture(img_name, width=Inches(3.0))
            img_paths.append(img_name)

        doc_buffer = io.BytesIO()
        doc.save(doc_buffer)
        doc_buffer.seek(0)
        for f in img_paths: 
            if os.path.exists(f): os.remove(f)
        
        st.success("Massive 14-Sector Master Report Generated!")
        st.download_button("Download Master Report", doc_buffer, f"Master_14_Sector_{end_date}.docx")
