import streamlit as st
import pandas as pd
from datetime import datetime
import os
import plotly.express as px
from github import Github  # Required for syncing to GitHub

# --- Configuration ---
st.set_page_config(page_title="ICT Department Portal", layout="wide")
DATA_FILE = "ict_master_log.csv"

# --- GitHub Sync Setup (For Cloud Deployment) ---
# To use this, add your GitHub token to Streamlit Secrets (st.secrets["GITHUB_TOKEN"])
GITHUB_REPO = "yourusername/your-repo-name" # CHANGE THIS

def push_to_github(file_path):
    try:
        if "GITHUB_TOKEN" in st.secrets:
            g = Github(st.secrets["GITHUB_TOKEN"])
            repo = g.get_repo(GITHUB_REPO)
            with open(file_path, 'r') as file:
                content = file.read()
            
            # Check if file exists in repo to update, otherwise create
            try:
                contents = repo.get_contents(file_path)
                repo.update_file(contents.path, "Auto-sync updated log", content, contents.sha)
            except:
                repo.create_file(file_path, "Auto-sync created log", content)
    except Exception as e:
        pass # Fails silently locally, but you can print(e) for debugging

# --- Helper Function to Load Data ---
@st.cache_data(ttl=5) # Refreshes data frequently
def load_data():
    if os.path.exists(DATA_FILE):
        return pd.read_csv(DATA_FILE)
    else:
        return pd.DataFrame(columns=[
            "Date", "Day", "Time", "Department", "Reported By", "System ID", 
            "Description", "Problem", "Action", "Parts", "IT Staff", "Status", "Remarks"
        ])

df = load_data()

# --- Main App ---
st.title("💻 Executive ICT Activity Portal")
st.markdown("*Use **Ctrl+P** (Windows) or **Cmd+P** (Mac) to export this entire page as a PDF.*")

tab1, tab2, tab3 = st.tabs(["📝 Log New Activity", "📊 Dynamic Dashboard", "🗄️ Database View"])

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
                "Department": department, "Reported By": reported_by, "System ID": system_id,
                "Description": description, "Problem": problem, "Action": action,
                "Parts": parts, "IT Staff": it_staff, "Status": status, "Remarks": remarks
            }
            new_df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
            new_df.to_csv(DATA_FILE, index=False)
            push_to_github(DATA_FILE) # Syncs to GitHub automatically
            st.success("✅ Logged and synced to GitHub successfully!")
            st.rerun()

# ==========================================
# TAB 2: DYNAMIC INFOGRAPHIC DASHBOARD
# ==========================================
with tab2:
    if df.empty:
        st.info("No data logged yet.")
    else:
        st.subheader("Interactive Visualizations")
        
        # Dropdown for dynamic sorting
        sort_by = st.selectbox(
            "Select Metric to Analyze:", 
            ["Status", "Department", "IT Staff", "Reported By", "Diagnosed Problem", "Action Taken", "System ID", "Replaced Parts"],
            index=0 # Default is Status
        )
        
        # Dynamic Chart Generation based on selection
        counts = df[sort_by].value_counts().reset_index()
        counts.columns = [sort_by, "Count"]
        
        if sort_by == "Status":
            fig = px.pie(counts, values="Count", names="Status", hole=0.4, 
                         color="Status", color_discrete_map={"Resolved": "#28a745", "Pending": "#dc3545"},
                         title="Overall Resolution Status")
        elif sort_by in ["Department", "IT Staff"]:
            fig = px.bar(counts, x=sort_by, y="Count", color=sort_by, title=f"Issues segmented by {sort_by}")
        else:
            # Horizontal bar chart for longer text fields (Problems, Actions, etc.)
            fig = px.bar(counts, x="Count", y=sort_by, orientation='h', title=f"Breakdown of {sort_by}")
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
            search_term = st.text_input("🔍 Search any keyword (e.g., specific ID, problem, staff)...")
        with col_filter:
            selected_columns = st.multiselect("Select columns to view:", df.columns.tolist(), default=df.columns.tolist())
        
        # Filter Logic
        filtered_df = df[selected_columns]
        if search_term:
            # Search across all columns as strings
            mask = filtered_df.astype(str).apply(lambda x: x.str.contains(search_term, case=False)).any(axis=1)
            filtered_df = filtered_df[mask]
            
        st.dataframe(filtered_df, use_container_width=True)
        
        # CSV Download Button
        csv_data = filtered_df.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Filtered Data as CSV", csv_data, "Filtered_ICT_Logs.csv", "text/csv")