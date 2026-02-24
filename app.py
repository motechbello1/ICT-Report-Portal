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
DATA_FILE = "ict_master_log.csv"
TRACKER_FILE = "last_report_date.txt"
GITHUB_REPO = "motechbello1/ICT-Report-Portal" # CHANGE THIS

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
            return False, "GITHUB_TOKEN not found in Streamlit Secrets. Saved locally."
    except Exception as e:
        return False, f"GitHub Sync Error: {str(e)}"

# --- Helper Function to Load & Clean Data ---
def load_data():
    if os.path.exists(DATA_FILE) and os.path.getsize(DATA_FILE) > 0:
        df = pd.read_csv(DATA_FILE)
        if 'Status' in df.columns:
            df['Status'] = df['Status'].astype(str).str.lower().str.strip()
        return df
    else:
        df = pd.DataFrame(columns=[
            "Date", "Day", "Time", "Department", "Reported By", "System ID", 
            "Description", "Problem", "Action", "Parts", "IT Staff", "Status", "Remarks"
        ])
        df.to_csv(DATA_FILE, index=False)
        return df

def get_last_report_date():
    if os.path.exists(TRACKER_FILE):
        with open(TRACKER_FILE, 'r') as f:
            try:
                return datetime.strptime(f.read().strip(), "%Y-%m-%d").date()
            except:
                pass
    # Default to 7 days ago if no previous report exists
    return (datetime.now() - timedelta(days=7)).date()

def save_last_report_date(end_date):
    with open(TRACKER_FILE, 'w') as f:
        f.write(end_date.strftime("%Y-%m-%d"))

df = load_data()

# --- Main App ---
st.title("NBTI ICT Department Activity Portal")

if "form_message" in st.session_state:
    if st.session_state.message_type == "success":
        st.success(st.session_state.form_message)
    else:
        st.warning(st.session_state.form_message)
    del st.session_state.form_message
    del st.session_state.message_type


tab1, tab2, tab3, tab4 = st.tabs(["📝 Log New Activity", "📊 Dynamic Dashboard", "🗄️ Database View", "📑 Weekly Report"])

# ==========================================
# TAB 1: DATA LOGGING INTERFACE
# ==========================================
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
            
        if st.form_submit_button("💾 Save Log & Sync"):
            now = datetime.now()
            new_data = {
                "Date": now.strftime("%Y-%m-%d"), "Day": now.strftime("%A"), "Time": now.strftime("%H:%M:%S"),
                "Department": department.lower().strip(), "Reported By": reported_by.lower().strip(), 
                "System ID": system_id.lower().strip(), "Description": description.lower().strip(), 
                "Problem": problem.lower().strip(), "Action": action.lower().strip(),
                "Parts": parts.lower().strip(), "IT Staff": it_staff.lower().strip(), 
                "Status": status.lower(), "Remarks": remarks.lower().strip()
            }
            new_df = pd.DataFrame([new_data])
            new_df.to_csv(DATA_FILE, mode='a', header=not os.path.exists(DATA_FILE), index=False)
            sync_success, sync_msg = push_to_github(DATA_FILE)
            
            if sync_success:
                st.session_state.message_type = "success"
                st.session_state.form_message = f"✅ Log safely added! {sync_msg}"
            else:
                st.session_state.message_type = "warning"
                st.session_state.form_message = f"⚠️ Saved locally, but GitHub sync failed: {sync_msg}"
            st.rerun() 

# ==========================================
# TAB 2 & 3: DASHBOARD & DATABASE
# ==========================================
with tab2:
    if df.empty:
        st.info("No data logged yet.")
    else:
        metric_mapping = {"Status": "Status", "Department": "Department", "IT Staff": "IT Staff", "Diagnosed Problem": "Problem", "Action Taken": "Action", "System ID": "System ID"}
        sort_by_display = st.selectbox("Select Metric to Analyze:", list(metric_mapping.keys()), index=0)
        sort_by_col = metric_mapping[sort_by_display]
        counts = df[sort_by_col].value_counts().reset_index()
        counts.columns = [sort_by_col, "Count"]
        
        if sort_by_col == "Status":
            fig = px.pie(counts, values="Count", names="Status", hole=0.4, color="Status", color_discrete_map={"resolved": "#28a745", "reserved": "#28a745", "pending": "#dc3545"})
        elif sort_by_col in ["Department", "IT Staff"]:
            fig = px.bar(counts, x=sort_by_col, y="Count", color=sort_by_col)
        else:
            fig = px.bar(counts, x="Count", y=sort_by_col, orientation='h')
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig, use_container_width=True)

with tab3:
    if not df.empty:
        search_term = st.text_input("🔍 Search database...").lower()
        selected_columns = st.multiselect("Columns:", df.columns.tolist(), default=df.columns.tolist())
        filtered_df = df[selected_columns]
        if search_term:
            mask = filtered_df.astype(str).apply(lambda x: x.str.contains(search_term, case=False)).any(axis=1)
            filtered_df = filtered_df[mask]
        st.dataframe(filtered_df, use_container_width=True)
        st.download_button("📥 Download Filtered Data (CSV)", filtered_df.to_csv(index=False).encode('utf-8'), "Filtered_Logs.csv", "text/csv")

# ==========================================
# TAB 4: AUTOMATED SMART REPORT (.DOCX)
# ==========================================
with tab4:
    st.subheader("Generate Executive Action Report")
    
    # Calculate Dates
    last_report_start = get_last_report_date()
    calculated_end = last_report_start + timedelta(days=7)
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        start_date = st.date_input("Report Start Date (Auto-detected from last report)", value=last_report_start)
    with col_d2:
        end_date = st.date_input("Report End Date (Auto-calculated +7 days)", value=calculated_end)

    if st.button("📄 Generate Analyzed Executive Report"):
        if df.empty:
            st.error("No data available in the master log.")
        else:
            temp_df = df.copy()
            temp_df['Date'] = pd.to_datetime(temp_df['Date']).dt.date
            
            # --- STRICT AUTHENTICITY & FACTUAL CHECK ---
            logged_dates = temp_df['Date'].unique()
            
            if start_date not in logged_dates:
                st.error(f"⚠️ Authenticity Error: The exact Start Date selected ({start_date}) does not exist in the master log. To guarantee factual reporting, please select a valid date with logged activities.")
            elif end_date not in logged_dates:
                st.error(f"⚠️ Authenticity Error: The exact End Date selected ({end_date}) does not exist in the master log. To guarantee factual reporting, please select a valid date with logged activities.")
            else:
                # Filter Data STRICTLY to the selected dates
                weekly_df = temp_df[(temp_df['Date'] >= start_date) & (temp_df['Date'] <= end_date)].copy()
                
                if weekly_df.empty:
                    st.warning(f"No records found between {start_date} and {end_date}.")
                else:
                    # Calculate Core Metrics
                    total_issues = len(weekly_df)
                    resolved = len(weekly_df[weekly_df['Status'].isin(['resolved', 'reserved'])])
                    pending = total_issues - resolved
                    res_rate = (resolved / total_issues) * 100 if total_issues > 0 else 0
                    
                    top_dept = weekly_df['Department'].value_counts().index[0].title() if not weekly_df['Department'].empty else "N/A"
                    top_dept_count = weekly_df['Department'].value_counts().iloc[0] if not weekly_df['Department'].empty else 0
                    
                    top_prob = weekly_df['Problem'].value_counts().index[0].capitalize() if not weekly_df['Problem'].empty else "N/A"
                    top_staff = weekly_df['IT Staff'].value_counts().index[0].title() if not weekly_df['IT Staff'].empty else "N/A"

                    # Generate Images via Matplotlib
                    doc = Document()
                    doc.add_heading('ICT Executive Summary & Analytics', 0)
                    doc.add_paragraph(f"Reporting Period: {start_date} to {end_date}\nGenerated on: {datetime.now().strftime('%Y-%m-%d')}")
                    
                    # --- CHART 1: STATUS PIE ---
                    fig1, ax1 = plt.subplots(figsize=(5, 3))
                    colors = [{'resolved': '#28a745', 'reserved': '#28a745', 'pending': '#dc3545'}.get(x, '#333') for x in weekly_df['Status'].value_counts().index]
                    weekly_df['Status'].value_counts().plot(kind='pie', autopct='%1.1f%%', colors=colors, ax=ax1)
                    ax1.set_ylabel("")
                    fig1.savefig("c1.png", bbox_inches='tight')
                    plt.close(fig1)

                    doc.add_heading('1. Resolution Efficiency', level=1)
                    doc.add_picture("c1.png", width=Inches(4.5))
                    doc.add_paragraph(f"Analysis: During this period, the ICT department processed {total_issues} total support requests. "
                                      f"The team successfully resolved {resolved} cases, resulting in an overall resolution rate of {res_rate:.1f}%. "
                                      f"Currently, {pending} cases remain pending and will be prioritized in the upcoming cycle.")

                    # --- CHART 2: DEPARTMENT BAR ---
                    fig2, ax2 = plt.subplots(figsize=(5, 3))
                    weekly_df['Department'].value_counts().plot(kind='bar', color='#007bff', ax=ax2)
                    plt.xticks(rotation=45, ha='right')
                    fig2.savefig("c2.png", bbox_inches='tight')
                    plt.close(fig2)

                    doc.add_heading('2. Departmental Demand', level=1)
                    doc.add_picture("c2.png", width=Inches(4.5))
                    doc.add_paragraph(f"Analysis: This chart illustrates where ICT resources are being consumed. The highest volume of requests "
                                      f"originated from the {top_dept} department, generating {top_dept_count} individual tickets. "
                                      f"Monitoring these trends allows us to identify if specific departments require new hardware or additional user-training.")

                    # --- CHART 3: TOP PROBLEMS (Horizontal Bar) ---
                    fig3, ax3 = plt.subplots(figsize=(5, 3))
                    weekly_df['Problem'].value_counts().nlargest(5).plot(kind='barh', color='#ffc107', ax=ax3)
                    ax3.invert_yaxis() # Highest at top
                    fig3.savefig("c3.png", bbox_inches='tight')
                    plt.close(fig3)

                    doc.add_heading('3. Primary Diagnosed Faults', level=1)
                    doc.add_picture("c3.png", width=Inches(4.5))
                    doc.add_paragraph(f"Analysis: Focusing on the top 5 reported issues, the most frequently diagnosed problem was '{top_prob}'. "
                                      f"By tracking specific hardware or software failures, the ICT department can shift from reactive maintenance "
                                      f"to proactive replacement strategies for failing system types.")

                    # --- CHART 4: IT STAFF PERFORMANCE ---
                    fig4, ax4 = plt.subplots(figsize=(5, 3))
                    weekly_df['IT Staff'].value_counts().plot(kind='bar', color='#17a2b8', ax=ax4)
                    plt.xticks(rotation=45, ha='right')
                    fig4.savefig("c4.png", bbox_inches='tight')
                    plt.close(fig4)

                    doc.add_heading('4. Staff Workload Distribution', level=1)
                    doc.add_picture("c4.png", width=Inches(4.5))
                    doc.add_paragraph(f"Analysis: This metric tracks the individual ticket completion rates of the ICT personnel. "
                                      f"For this period, {top_staff} managed the highest volume of system interventions. "
                                      f"This data is vital for ensuring balanced resource allocation and preventing technician burnout.")

                    # --- TABLE: ONLY THE FILTERED WINDOW ---
                    doc.add_heading(f'5. Activity Log ({start_date} to {end_date})', level=1)
                    table = doc.add_table(rows=1, cols=6)
                    table.style = 'Table Grid'
                    hdr = table.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text, hdr[4].text, hdr[5].text = 'Date', 'Dept', 'System ID', 'Problem', 'Staff', 'Status'

                    for _, row in weekly_df.sort_values(by='Date').iterrows():
                        row_cells = table.add_row().cells
                        row_cells[0].text = str(row['Date'])
                        row_cells[1].text = str(row['Department']).title()
                        row_cells[2].text = str(row['System ID']).upper()
                        row_cells[3].text = str(row['Problem']).capitalize()
                        row_cells[4].text = str(row['IT Staff']).title()
                        row_cells[5].text = str(row['Status']).title()

                    # Clean up temp images
                    for file in ["c1.png", "c2.png", "c3.png", "c4.png"]:
                        if os.path.exists(file): os.remove(file)

                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    
                    # Update the tracker file so the NEXT report starts exactly where this one ended
                    save_last_report_date(end_date)
                    
                    st.success(f"✅ Verified & Authentic Report generated for {start_date} to {end_date}!")
                    st.download_button("📥 Download Analyzed Report (.docx)", doc_buffer, f"ICT_Report_{end_date}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

