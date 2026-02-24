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
GITHUB_REPO = "yourusername/your-repo-name" # CHANGE THIS

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
                repo.update_file(contents.path, "Auto-sync updated log", content, contents.sha)
                return True, "Synced to GitHub successfully!"
            except:
                repo.create_file(file_path, "Auto-sync created log", content)
                return True, "Created and synced to GitHub successfully!"
        else:
            return False, "GITHUB_TOKEN not found in Streamlit Secrets. Saved locally."
    except Exception as e:
        return False, f"GitHub Sync Error: {str(e)}"

# --- Helper Function to Load & Clean Data ---
def load_data():
    if os.path.exists(DATA_FILE) and os.path.getsize(DATA_FILE) > 0:
        df = pd.read_csv(DATA_FILE)
        
        # FORCE ALL STATUSES TO LOWERCASE TO FIX CASE SENSITIVITY
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

df = load_data()

# --- Main App ---
st.title("💻 Executive ICT Activity Portal")

# --- MEMORY CHECK FOR SUCCESS MESSAGES ---
# This checks if a message was saved in memory before the page refreshed
if "form_message" in st.session_state:
    if st.session_state.message_type == "success":
        st.success(st.session_state.form_message)
    else:
        st.warning(st.session_state.form_message)
    # Delete the message from memory so it doesn't show up forever
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
            status = st.selectbox("Status", ["Resolved", "Pending"]) # Dropdown stays capitalized for UI
            remarks = st.text_area("Remarks")
            
        if st.form_submit_button("💾 Save Log & Sync"):
            now = datetime.now()
            
            new_data = {
                "Date": now.strftime("%Y-%m-%d"), 
                "Day": now.strftime("%A"), 
                "Time": now.strftime("%H:%M:%S"),
                "Department": department.lower().strip(), 
                "Reported By": reported_by.lower().strip(), 
                "System ID": system_id.lower().strip(),
                "Description": description.lower().strip(), 
                "Problem": problem.lower().strip(), 
                "Action": action.lower().strip(),
                "Parts": parts.lower().strip(), 
                "IT Staff": it_staff.lower().strip(), 
                "Status": status.lower(), # Saved as lowercase
                "Remarks": remarks.lower().strip()
            }
            
            new_df = pd.DataFrame([new_data])
            new_df.to_csv(DATA_FILE, mode='a', header=not os.path.exists(DATA_FILE), index=False)
            
            sync_success, sync_msg = push_to_github(DATA_FILE)
            
            # SAVE MESSAGE TO MEMORY BEFORE REFRESHING
            if sync_success:
                st.session_state.message_type = "success"
                st.session_state.form_message = f"✅ Log safely added to database! {sync_msg}"
            else:
                st.session_state.message_type = "warning"
                st.session_state.form_message = f"⚠️ Log added locally, but GitHub sync failed: {sync_msg}"
                
            st.rerun() # Refresh page to update charts

# ==========================================
# TAB 2: DYNAMIC INFOGRAPHIC DASHBOARD
# ==========================================
with tab2:
    if df.empty:
        st.info("No data logged yet.")
    else:
        st.subheader("Interactive Visualizations")
        
        metric_mapping = {
            "Status": "Status", 
            "Department": "Department", 
            "IT Staff": "IT Staff", 
            "Reported By": "Reported By", 
            "Diagnosed Problem": "Problem", 
            "Action Taken": "Action", 
            "System ID": "System ID", 
            "Replaced Parts": "Parts"
        }
        
        sort_by_display = st.selectbox("Select Metric to Analyze:", list(metric_mapping.keys()), index=0)
        sort_by_col = metric_mapping[sort_by_display]
        
        counts = df[sort_by_col].value_counts().reset_index()
        counts.columns = [sort_by_col, "Count"]
        
        if sort_by_col == "Status":
            # Maps "resolved" AND "reserved" to green, "pending" to red
            fig = px.pie(counts, values="Count", names="Status", hole=0.4, 
                         color="Status", color_discrete_map={"resolved": "#28a745", "reserved": "#28a745", "pending": "#dc3545"},
                         title="Overall Resolution Status")
        elif sort_by_col in ["Department", "IT Staff"]:
            fig = px.bar(counts, x=sort_by_col, y="Count", color=sort_by_col, title=f"Issues segmented by {sort_by_display}")
        else:
            fig = px.bar(counts, x="Count", y=sort_by_col, orientation='h', title=f"Breakdown of {sort_by_display}")
            fig.update_layout(yaxis={'categoryorder':'total ascending'})

        st.plotly_chart(fig, use_container_width=True)

# ==========================================
# TAB 3: SEARCHABLE DATABASE VIEW
# ==========================================
with tab3:
    st.subheader("Master Activity Log")
    if not df.empty:
        col_search, col_filter = st.columns([2, 1])
        with col_search:
            search_term = st.text_input("🔍 Search any keyword...").lower()
        with col_filter:
            selected_columns = st.multiselect("Select columns to view:", df.columns.tolist(), default=df.columns.tolist())
        
        filtered_df = df[selected_columns]
        if search_term:
            mask = filtered_df.astype(str).apply(lambda x: x.str.contains(search_term, case=False)).any(axis=1)
            filtered_df = filtered_df[mask]
            
        st.dataframe(filtered_df, use_container_width=True)
        csv_data = filtered_df.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Filtered Data (CSV)", csv_data, "Filtered_ICT_Logs.csv", "text/csv")

# ==========================================
# TAB 4: AUTOMATED WEEKLY REPORT (.DOCX)
# ==========================================
with tab4:
    st.subheader("Generate Executive Weekly Report")
    st.write("Click below to analyze the last 7 days of data and generate a professional Word Document.")
    
    if st.button("📄 Generate Report Document"):
        if df.empty:
            st.error("No data available to generate a report.")
        else:
            temp_df = df.copy()
            temp_df['Date'] = pd.to_datetime(temp_df['Date'])
            one_week_ago = datetime.now() - timedelta(days=7)
            weekly_df = temp_df[temp_df['Date'] >= one_week_ago].copy()
            
            if weekly_df.empty:
                st.warning("No records logged in the past 7 days.")
            else:
                total_issues = len(weekly_df)
                # Count both resolved and reserved as finished
                resolved = len(weekly_df[weekly_df['Status'].isin(['resolved', 'reserved'])])
                pending = total_issues - resolved
                top_dept = weekly_df['Department'].mode()[0].title() if not weekly_df['Department'].empty else "N/A"
                top_staff = weekly_df['IT Staff'].mode()[0].title() if not weekly_df['IT Staff'].empty else "N/A"

                chart_path = "temp_chart.png"
                fig, ax = plt.subplots(figsize=(6, 4))
                color_map = {'resolved': '#28a745', 'reserved': '#28a745', 'pending': '#dc3545'}
                colors = [color_map.get(x, '#333333') for x in weekly_df['Status'].value_counts().index]
                weekly_df['Status'].value_counts().plot(kind='pie', autopct='%1.1f%%', colors=colors, ax=ax)
                ax.set_title("Weekly Resolution Status")
                ax.set_ylabel("")
                fig.savefig(chart_path)
                plt.close(fig)

                doc = Document()
                doc.add_heading('ICT Department - Executive Weekly Summary', 0)
                
                doc.add_heading('1. Executive Narrative', level=1)
                now = datetime.now()
                narrative = (
                    f"During the week ending {now.strftime('%Y-%m-%d')}, the ICT department managed a total of {total_issues} reported incidents. "
                    f"The team successfully resolved {resolved} of these issues, leaving {pending} pending for follow-up. "
                    f"The highest volume of requests originated from the '{top_dept}' department. "
                    f"Special commendation to {top_staff} for leading the resolution metrics this week."
                )
                doc.add_paragraph(narrative)

                doc.add_heading('2. Status Infographic', level=1)
                doc.add_picture(chart_path, width=Inches(5.0))
                os.remove(chart_path) 

                doc.add_heading('3. Complete Weekly Activity Log', level=1)
                table = doc.add_table(rows=1, cols=6)
                table.style = 'Table Grid'
                
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'Date'
                hdr_cells[1].text = 'Dept'
                hdr_cells[2].text = 'System ID'
                hdr_cells[3].text = 'Problem'
                hdr_cells[4].text = 'Staff'
                hdr_cells[5].text = 'Status'

                for index, row in weekly_df.iterrows():
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(row['Date'].strftime('%Y-%m-%d'))
                    row_cells[1].text = str(row['Department']).title()
                    row_cells[2].text = str(row['System ID']).upper()
                    row_cells[3].text = str(row['Problem']).capitalize()
                    row_cells[4].text = str(row['IT Staff']).title()
                    row_cells[5].text = str(row['Status']).title()

                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)

                st.success("✅ Report generated successfully!")
                
                st.download_button(
                    label="📥 Download Weekly Report (.docx)",
                    data=doc_buffer,
                    file_name=f"ICT_Weekly_Report_{now.strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
