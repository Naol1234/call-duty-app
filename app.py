import streamlit as st
import pandas as pd
from datetime import datetime, date
import io
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Outbound Call Duty Dashboard", layout="wide", page_icon="📞")

# ─────────────────────────────────────────────
# AGENTS
# ─────────────────────────────────────────────
AGENTS = [
    "Fuad Rahimli", "Muhammad Taimoor", "Allyson Kawondera",
    "Sofiane Seddik", "Joseph Muwuduri", "Joseph Muteme",
    "Prince Ishimwe", "Naol Uso", "Wayne Foromozo", "Wilmah Mupa"
]

DELAY_REASONS = [
    "", "Different time zone", "Public holiday",
    "Impossible to contact customer", "Delayed from other department", "Other"
]

RESOLUTION_OPTIONS = ["", "Resolved", "Not Resolved", "Other"]

COLUMNS = [
    "Response", "Date of listing", "Case Number", "Date/Time",
    "Assigned to (agents)", "Recent Interaction Date", "Recent Interaction Notes",
    "Delayed from other department", "Impossible to contact (reason)",
    "Call Answered (Yes/No)", "Date of call if answered",
    "Resolution keywords", "Case Resolved (Yes/No)",
    "1st Callback attempt (date & time)", "2nd Callback attempt (date & time)",
    "Postponed to another day", "Resolved once and for all"
]

MANAGER_PASSWORD = "manager123"
SHEET_ID = "1p7jFcvQIKOJaPHn-CmvuIL9ucZMzheT-YpnlrItzd5w"

# ─────────────────────────────────────────────
# GOOGLE SHEETS CONNECTION
# ─────────────────────────────────────────────
@st.cache_resource
def get_gsheet_client():
    creds_dict = st.secrets["gcp_service_account"]
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

def load_from_gsheet():
    try:
        client = get_gsheet_client()
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet("Check Daily")
        data = ws.get_all_values()
        if len(data) < 2:
            return pd.DataFrame(columns=COLUMNS)
        headers = data[0]
        rows = data[1:]
        df = pd.DataFrame(rows, columns=headers)
        # Ensure all columns exist
        for c in COLUMNS:
            if c not in df.columns:
                df[c] = ""
        df = df[COLUMNS]
        df = df[df["Case Number"].notna() & (df["Case Number"] != "")]
        df["Case Number"] = df["Case Number"].apply(
            lambda x: str(int(float(x))) if str(x).replace('.','').isdigit() else str(x)
        ).str.strip()
        return df
    except Exception as e:
        st.error(f"Error loading from Google Sheets: {e}")
        return pd.DataFrame(columns=COLUMNS)

def save_to_gsheet(df):
    try:
        client = get_gsheet_client()
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet("Check Daily")
        df_clean = df.fillna("").astype(str)
        ws.clear()
        ws.update([df_clean.columns.tolist()] + df_clean.values.tolist())
        return True
    except Exception as e:
        st.error(f"Error saving to Google Sheets: {e}")
        return False

def save_row_to_gsheet(df, idx):
    """Save a single row update efficiently."""
    save_to_gsheet(df)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "tickets" not in st.session_state:
    st.session_state.tickets = load_from_gsheet()

if "manager_authenticated" not in st.session_state:
    st.session_state.manager_authenticated = False

if "last_refresh" not in st.session_state:
    st.session_state.last_refresh = datetime.now()

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def export_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Check Daily")
    return buf.getvalue()

def auto_assign(new_df, existing_df):
    counts = {a: len(existing_df[existing_df["Assigned to (agents)"] == a]) for a in AGENTS}
    assigned = []
    for _ in range(len(new_df)):
        agent_pick = min(counts, key=counts.get)
        assigned.append(agent_pick)
        counts[agent_pick] += 1
    new_df = new_df.copy()
    new_df["Assigned to (agents)"] = assigned
    return new_df

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
st.sidebar.image("https://img.icons8.com/fluency/96/phone-office.png", width=60)
st.sidebar.title("📞 Call Duty App")
st.sidebar.markdown("---")

# Refresh button
if st.sidebar.button("🔄 Refresh Data"):
    st.session_state.tickets = load_from_gsheet()
    st.session_state.last_refresh = datetime.now()
    st.rerun()

st.sidebar.caption(f"Last refresh: {st.session_state.last_refresh.strftime('%H:%M:%S')}")
st.sidebar.markdown("---")

role = st.sidebar.radio("I am a:", ["👤 Agent", "🛠️ Manager"])
st.sidebar.markdown("---")

if role == "👤 Agent":
    agent_name = st.sidebar.selectbox("Select your name:", AGENTS)
    pages = ["My Tickets", "Fill Ticket Details"]
    page = st.sidebar.radio("Navigate:", pages)

else:
    if not st.session_state.manager_authenticated:
        st.title("🔒 Manager Access")
        st.markdown("Please enter the manager password to continue.")
        pwd = st.text_input("Password:", type="password")
        if st.button("Login", type="primary"):
            if pwd == MANAGER_PASSWORD:
                st.session_state.manager_authenticated = True
                st.rerun()
            else:
                st.error("❌ Incorrect password. Try again.")
        st.stop()

    if st.sidebar.button("🔓 Logout"):
        st.session_state.manager_authenticated = False
        st.rerun()

    pages = ["Dashboard", "All Tickets", "Add Ticket", "Manage Tickets", "Import / Export"]
    page = st.sidebar.radio("Navigate:", pages)

# ─────────────────────────────────────────────
# AGENT PAGES
# ─────────────────────────────────────────────
if role == "👤 Agent":
    df = st.session_state.tickets

    if page == "My Tickets":
        st.title(f"🎫 My Tickets — {agent_name}")
        my = df[df["Assigned to (agents)"] == agent_name].copy()

        if my.empty:
            st.info("No tickets assigned to you yet.")
        else:
            unresolved = my[my["Case Resolved (Yes/No)"].isna() | (my["Case Resolved (Yes/No)"] == "") | (~my["Case Resolved (Yes/No)"].isin(["Resolved"]))]
            resolved   = my[my["Case Resolved (Yes/No)"] == "Resolved"]

            col1, col2, col3 = st.columns(3)
            col1.metric("Total Assigned", len(my))
            col2.metric("Pending", len(unresolved))
            col3.metric("Resolved", len(resolved))

            st.markdown("### 🔴 Pending Tickets")
            if unresolved.empty:
                st.success("All tickets resolved! 🎉")
            else:
                st.dataframe(
                    unresolved[["Case Number", "Date/Time", "Recent Interaction Date",
                                "Call Answered (Yes/No)", "Case Resolved (Yes/No)"]],
                    use_container_width=True, hide_index=True
                )

            st.markdown("### ✅ Resolved Tickets")
            if resolved.empty:
                st.info("No resolved tickets yet.")
            else:
                st.dataframe(
                    resolved[["Case Number", "Date/Time", "Call Answered (Yes/No)",
                               "Resolution keywords", "Case Resolved (Yes/No)"]],
                    use_container_width=True, hide_index=True
                )

    elif page == "Fill Ticket Details":
        st.title(f"✏️ Fill Ticket Details — {agent_name}")
        my = df[df["Assigned to (agents)"] == agent_name]

        if my.empty:
            st.info("No tickets assigned to you yet.")
        else:
            case_options = my["Case Number"].tolist()
            selected_case = st.selectbox("Select Case Number to update:", case_options)

            idx = df[df["Case Number"] == selected_case].index[0]
            row = df.loc[idx]

            with st.expander("📋 Ticket Info", expanded=True):
                col1, col2, col3 = st.columns(3)
                col1.text_input("Case Number 🔒", value=str(row['Case Number']), disabled=True)
                new_dt = col2.text_input("Date/Time", value=str(row['Date/Time']) if str(row['Date/Time']) not in ["None","nan"] else "")
                new_agent = col3.selectbox("Assigned to", AGENTS,
                    index=AGENTS.index(row["Assigned to (agents)"])
                          if row["Assigned to (agents)"] in AGENTS else 0)

            st.markdown("### Fill in the fields below")

            st.markdown("**Recent Interaction Date**")
            try:
                default_date = pd.to_datetime(row["Recent Interaction Date"]).date() \
                    if pd.notna(row["Recent Interaction Date"]) and str(row["Recent Interaction Date"]) not in ["", "None", "nan"] \
                    else date.today()
            except:
                default_date = date.today()
            ri_date = st.date_input("Date", value=default_date, label_visibility="collapsed")
            ri_notes = st.text_area("Recent Interaction Notes",
                value=str(row["Recent Interaction Notes"]) if pd.notna(row["Recent Interaction Notes"]) and str(row["Recent Interaction Notes"]) not in ["None", "nan"] else "",
                placeholder="Type your notes here...")

            delayed = st.selectbox("Delayed from other department?", ["", "Yes", "No"],
                index=["", "Yes", "No"].index(str(row["Delayed from other department"]))
                      if str(row["Delayed from other department"]) in ["Yes", "No"] else 0)

            impossible = st.selectbox("Impossible to contact (reason):", DELAY_REASONS,
                index=DELAY_REASONS.index(str(row["Impossible to contact (reason)"]))
                      if str(row["Impossible to contact (reason)"]) in DELAY_REASONS else 0)

            call_answered = st.selectbox("Call Answered?", ["", "Yes", "No"],
                index=["", "Yes", "No"].index(str(row["Call Answered (Yes/No)"]))
                      if str(row["Call Answered (Yes/No)"]) in ["Yes", "No"] else 0)

            st.markdown("---")
            st.markdown("#### 📞 Call Details")

            st.markdown("**Date of call (if answered)**")
            try:
                default_call_date = pd.to_datetime(row["Date of call if answered"]).date() \
                    if pd.notna(row["Date of call if answered"]) and str(row["Date of call if answered"]) not in ["", "None", "nan"] \
                    else date.today()
            except:
                default_call_date = date.today()
            call_date = st.date_input("Call date", value=default_call_date, label_visibility="collapsed")

            resolution_kw = st.text_input("📝 Resolution keywords:",
                value=str(row["Resolution keywords"]) if pd.notna(row["Resolution keywords"]) and str(row["Resolution keywords"]) not in ["None","nan"] else "",
                placeholder="e.g. Created WO, Scheduled visit, Sent email...")

            case_resolved = st.selectbox("Case Resolved?", RESOLUTION_OPTIONS,
                index=RESOLUTION_OPTIONS.index(str(row["Case Resolved (Yes/No)"]))
                      if str(row["Case Resolved (Yes/No)"]) in RESOLUTION_OPTIONS else 0)

            st.markdown("---")
            st.markdown("#### 🔁 Callback Details")
            cb1 = st.text_input("1st Callback attempt (date & time):",
                value=str(row["1st Callback attempt (date & time)"]) if pd.notna(row["1st Callback attempt (date & time)"]) and str(row["1st Callback attempt (date & time)"]) not in ["None","nan"] else "",
                placeholder="e.g. 2026-05-26 10:00")
            cb2 = st.text_input("2nd Callback attempt (date & time):",
                value=str(row["2nd Callback attempt (date & time)"]) if pd.notna(row["2nd Callback attempt (date & time)"]) and str(row["2nd Callback attempt (date & time)"]) not in ["None","nan"] else "",
                placeholder="e.g. 2026-05-27 14:00")
            postponed = st.selectbox("Postponed to another day?", ["", "Yes", "No"],
                index=["", "Yes", "No"].index(str(row["Postponed to another day"]))
                      if str(row["Postponed to another day"]) in ["Yes", "No"] else 0)

            st.markdown("---")
            st.markdown("#### ✅ Resolved Once and For All")
            col_check, col_notes = st.columns([1, 3])
            with col_check:
                resolved_all_check = st.checkbox("Mark as resolved once and for all",
                    value=str(row["Resolved once and for all"]) in ["True", "Yes", "1"])
            with col_notes:
                resolved_all_notes = st.text_input("Final notes (optional):",
                    placeholder="Any final notes about this case...")

            if st.button("💾 Save Changes", type="primary"):
                st.session_state.tickets.at[idx, "Date/Time"] = new_dt
                st.session_state.tickets.at[idx, "Assigned to (agents)"] = new_agent
                st.session_state.tickets.at[idx, "Recent Interaction Date"] = str(ri_date)
                st.session_state.tickets.at[idx, "Recent Interaction Notes"] = ri_notes
                st.session_state.tickets.at[idx, "Delayed from other department"] = delayed
                st.session_state.tickets.at[idx, "Impossible to contact (reason)"] = impossible
                st.session_state.tickets.at[idx, "Call Answered (Yes/No)"] = call_answered
                st.session_state.tickets.at[idx, "Date of call if answered"] = str(call_date)
                st.session_state.tickets.at[idx, "Resolution keywords"] = resolution_kw
                st.session_state.tickets.at[idx, "Case Resolved (Yes/No)"] = case_resolved
                st.session_state.tickets.at[idx, "1st Callback attempt (date & time)"] = cb1
                st.session_state.tickets.at[idx, "2nd Callback attempt (date & time)"] = cb2
                st.session_state.tickets.at[idx, "Postponed to another day"] = postponed
                st.session_state.tickets.at[idx, "Resolved once and for all"] = "Yes" if resolved_all_check else "No"
                with st.spinner("Saving to Google Sheets..."):
                    if save_to_gsheet(st.session_state.tickets):
                        st.success(f"✅ Ticket {selected_case} saved successfully!")
                    else:
                        st.error("Failed to save. Please try again.")

# ─────────────────────────────────────────────
# MANAGER PAGES
# ─────────────────────────────────────────────
else:
    df = st.session_state.tickets

    if page == "Dashboard":
        st.title("📊 Manager Dashboard")

        if df.empty:
            st.warning("No tickets loaded yet.")
        else:
            total = len(df)
            resolved_mask = df["Case Resolved (Yes/No)"] == "Resolved"
            pending_mask  = df["Case Resolved (Yes/No)"].isna() | (df["Case Resolved (Yes/No)"] == "")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Tickets", total)
            c2.metric("Resolved", resolved_mask.sum())
            c3.metric("Pending", pending_mask.sum())
            c4.metric("Call Answered", (df["Call Answered (Yes/No)"] == "Yes").sum())

            st.markdown("---")
            col_a, col_b = st.columns(2)

            with col_a:
                st.markdown("#### Tickets per Agent")
                agent_counts = df["Assigned to (agents)"].value_counts().reset_index()
                agent_counts.columns = ["Agent", "Tickets"]
                st.bar_chart(agent_counts.set_index("Agent"))

            with col_b:
                st.markdown("#### Resolved vs Pending per Agent")
                summary = df.groupby("Assigned to (agents)").apply(
                    lambda x: pd.Series({
                        "Resolved": (x["Case Resolved (Yes/No)"] == "Resolved").sum(),
                        "Pending":  (x["Case Resolved (Yes/No)"].isna() | (x["Case Resolved (Yes/No)"] == "")).sum()
                    })
                ).reset_index()
                st.dataframe(summary, use_container_width=True, hide_index=True)

    elif page == "All Tickets":
        st.title("📋 All Tickets")

        if df.empty:
            st.warning("No tickets loaded yet.")
        else:
            filter_agent = st.selectbox("Filter by agent:", ["All"] + AGENTS)
            filter_status = st.selectbox("Filter by resolution:", ["All", "Resolved", "Not Resolved", "Pending"])

            view_df = df.copy()
            if filter_agent != "All":
                view_df = view_df[view_df["Assigned to (agents)"] == filter_agent]
            if filter_status == "Resolved":
                view_df = view_df[view_df["Case Resolved (Yes/No)"] == "Resolved"]
            elif filter_status == "Not Resolved":
                view_df = view_df[view_df["Case Resolved (Yes/No)"] == "Not Resolved"]
            elif filter_status == "Pending":
                view_df = view_df[view_df["Case Resolved (Yes/No)"].isna() | (view_df["Case Resolved (Yes/No)"] == "")]

            st.write(f"Showing **{len(view_df)}** tickets")
            st.dataframe(view_df, use_container_width=True, hide_index=True)

    elif page == "Add Ticket":
        st.title("➕ Add Tickets")

        tab1, tab2, tab3 = st.tabs(["📋 Paste from Excel", "📁 Upload Excel File", "➕ Single Ticket"])

        with tab1:
            st.markdown("**Copy Case Number + Date/Time columns from Excel, paste below:**")
            pasted = st.text_area("Paste Excel data here:", height=200,
                placeholder="1414513\t2026-05-11 19:10\n1414517\t2026-05-11 19:38\n...")
            list_date_paste = st.date_input("Date of listing for all these tickets:", value=date.today(), key="paste_date")

            if st.button("➕ Add Pasted Tickets", type="primary"):
                if not pasted.strip():
                    st.error("Please paste some data first.")
                else:
                    try:
                        rows = []
                        skipped = []
                        existing_cases = st.session_state.tickets["Case Number"].values
                        for line in pasted.strip().split("\n"):
                            parts = line.strip().split("\t")
                            if len(parts) >= 1:
                                case = str(parts[0]).strip()
                                if not case:
                                    continue
                                try:
                                    case = str(int(float(case)))
                                except:
                                    pass
                                dt_val = parts[1].strip() if len(parts) >= 2 else ""
                                if case in existing_cases:
                                    skipped.append(case)
                                    continue
                                new_row = {c: "" for c in COLUMNS}
                                new_row["Case Number"] = case
                                new_row["Date of listing"] = str(list_date_paste)
                                new_row["Date/Time"] = dt_val
                                new_row["Response"] = "False"
                                rows.append(new_row)

                        if rows:
                            new_df = pd.DataFrame(rows)
                            new_df = auto_assign(new_df, st.session_state.tickets)
                            st.session_state.tickets = pd.concat(
                                [st.session_state.tickets, new_df], ignore_index=True
                            )
                            with st.spinner("Saving to Google Sheets..."):
                                save_to_gsheet(st.session_state.tickets)
                            st.success(f"✅ Added {len(rows)} tickets!")
                            st.dataframe(new_df[["Case Number", "Date/Time", "Assigned to (agents)"]],
                                use_container_width=True, hide_index=True)
                        if skipped:
                            st.warning(f"Skipped {len(skipped)} duplicates: {', '.join(skipped)}")
                    except Exception as e:
                        st.error(f"Error: {e}")

        with tab2:
            new_file = st.file_uploader("Upload Excel file with new tickets:", type=["xlsx"], key="bulk_upload")
            list_date_upload = st.date_input("Date of listing:", value=date.today(), key="upload_date")

            if new_file:
                try:
                    preview_df = pd.read_excel(new_file)
                    col_candidates = [c for c in preview_df.columns if "case" in c.lower()]
                    dt_candidates  = [c for c in preview_df.columns if "date" in c.lower() or "time" in c.lower()]

                    if not col_candidates:
                        st.error("Could not find a 'Case Number' column.")
                    else:
                        case_col = col_candidates[0]
                        dt_col   = dt_candidates[0] if dt_candidates else None
                        preview_df = preview_df.dropna(subset=[case_col])
                        preview_df[case_col] = preview_df[case_col].apply(
                            lambda x: str(int(float(x))) if str(x).replace('.','').isdigit() else str(x)
                        )
                        st.write(f"Found **{len(preview_df)}** tickets. Preview:")
                        st.dataframe(preview_df[[case_col] + ([dt_col] if dt_col else [])].head(10),
                            use_container_width=True, hide_index=True)

                        if st.button("➕ Add These Tickets", type="primary"):
                            existing_cases = st.session_state.tickets["Case Number"].values
                            rows = []
                            skipped = []
                            for _, r in preview_df.iterrows():
                                case = str(r[case_col]).strip()
                                if case in existing_cases:
                                    skipped.append(case)
                                    continue
                                new_row = {c: "" for c in COLUMNS}
                                new_row["Case Number"] = case
                                new_row["Date of listing"] = str(list_date_upload)
                                new_row["Date/Time"] = str(r[dt_col]) if dt_col else ""
                                new_row["Response"] = "False"
                                rows.append(new_row)

                            if rows:
                                new_df = pd.DataFrame(rows)
                                new_df = auto_assign(new_df, st.session_state.tickets)
                                st.session_state.tickets = pd.concat(
                                    [st.session_state.tickets, new_df], ignore_index=True
                                )
                                with st.spinner("Saving to Google Sheets..."):
                                    save_to_gsheet(st.session_state.tickets)
                                st.success(f"✅ Added {len(rows)} tickets!")
                            if skipped:
                                st.warning(f"Skipped {len(skipped)} duplicates.")
                except Exception as e:
                    st.error(f"Error: {e}")

        with tab3:
            case_num = st.text_input("Case Number *")
            list_date = st.date_input("Date of listing *", value=date.today(), key="single_date")
            dt = st.text_input("Date/Time (e.g. 5/11/2026 19:10)")
            auto = st.checkbox("Auto-assign agent", value=True)
            agent = "" if auto else st.selectbox("Assign to agent *", [""] + AGENTS)

            if st.button("➕ Add Ticket", type="primary"):
                if not case_num:
                    st.error("Case Number is required.")
                elif case_num in st.session_state.tickets["Case Number"].values:
                    st.error(f"Case {case_num} already exists.")
                else:
                    new_row = {c: "" for c in COLUMNS}
                    new_row["Case Number"] = case_num
                    new_row["Date of listing"] = str(list_date)
                    new_row["Date/Time"] = dt
                    new_row["Response"] = "False"
                    new_df = pd.DataFrame([new_row])
                    if auto:
                        new_df = auto_assign(new_df, st.session_state.tickets)
                        assigned = new_df["Assigned to (agents)"].iloc[0]
                    else:
                        new_df["Assigned to (agents)"] = agent
                        assigned = agent
                    st.session_state.tickets = pd.concat(
                        [st.session_state.tickets, new_df], ignore_index=True
                    )
                    with st.spinner("Saving..."):
                        save_to_gsheet(st.session_state.tickets)
                    st.success(f"✅ Ticket {case_num} assigned to {assigned}!")

    elif page == "Manage Tickets":
        st.title("🛠️ Manage Tickets")

        if df.empty:
            st.warning("No tickets loaded yet.")
        else:
            case_list = df["Case Number"].tolist()
            selected = st.selectbox("Select ticket to edit/delete:", case_list)
            idx = df[df["Case Number"] == selected].index[0]
            row = df.loc[idx]

            tab1, tab2 = st.tabs(["✏️ Edit", "🗑️ Delete"])

            with tab1:
                new_agent = st.selectbox("Reassign agent:", AGENTS,
                    index=AGENTS.index(row["Assigned to (agents)"])
                          if row["Assigned to (agents)"] in AGENTS else 0)
                new_date = st.text_input("Date/Time:", value=str(row["Date/Time"]))
                if st.button("💾 Save Edit", type="primary"):
                    st.session_state.tickets.at[idx, "Assigned to (agents)"] = new_agent
                    st.session_state.tickets.at[idx, "Date/Time"] = new_date
                    with st.spinner("Saving..."):
                        save_to_gsheet(st.session_state.tickets)
                    st.success("✅ Ticket updated!")

            with tab2:
                st.warning(f"Are you sure you want to delete ticket **{selected}**?")
                if st.button("🗑️ Confirm Delete", type="primary"):
                    st.session_state.tickets = df.drop(index=idx).reset_index(drop=True)
                    with st.spinner("Saving..."):
                        save_to_gsheet(st.session_state.tickets)
                    st.success(f"Ticket {selected} deleted.")
                    st.rerun()

    elif page == "Import / Export":
        st.title("📂 Import / Export")

        st.markdown("### 📥 Import Excel File")
        st.info("This will replace ALL current data with the Excel file contents.")
        uploaded = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])
        if uploaded:
            try:
                df_load = pd.read_excel(uploaded, sheet_name=0, header=1)
                df_load = df_load.dropna(how="all")
                col_map = {
                    "Response": "Response",
                    "Date of listing case in this file": "Date of listing",
                    "Case Number": "Case Number",
                    "Date/Time": "Date/Time",
                    "Assigned to (agents)": "Assigned to (agents)",
                    "Recent interaction (5 days)": "Recent Interaction Date",
                    "Delayed from other department ": "Delayed from other department",
                    "Is it  impossible to contact the customer: (due to different timezone, Public holidays, etc)": "Impossible to contact (reason)",
                    "Called answered (Yes or No)": "Call Answered (Yes/No)",
                    "Date of the call if answered": "Date of call if answered",
                    "if Yes, what is the resolution, key words": "Resolution keywords",
                    "Case resolved  (Yes or No)": "Case Resolved (Yes/No)",
                    "If no, 1st attempt time call back (date and time)": "1st Callback attempt (date & time)",
                    "If no, 2nd attempt call back (date abbd time)": "2nd Callback attempt (date & time)",
                    "to be postponed to another day (Yes, No)": "Postponed to another day",
                    "Resolved once and for all": "Resolved once and for all",
                }
                df_load = df_load.rename(columns=col_map)
                for c in COLUMNS:
                    if c not in df_load.columns:
                        df_load[c] = ""
                df_load = df_load[COLUMNS]
                df_load = df_load[df_load["Case Number"].notna()]
                df_load["Case Number"] = df_load["Case Number"].apply(
                    lambda x: str(int(float(x))) if str(x).replace('.','').isdigit() else str(x)
                ).str.strip()
                st.success(f"✅ Loaded {len(df_load)} tickets!")
                st.dataframe(df_load.head(10), use_container_width=True, hide_index=True)
                if st.button("✅ Confirm Import — Replace all data", type="primary"):
                    st.session_state.tickets = df_load
                    with st.spinner("Saving to Google Sheets..."):
                        save_to_gsheet(df_load)
                    st.success("✅ Data imported and saved to Google Sheets!")
                    st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

        st.markdown("---")
        st.markdown("### 📤 Export to Excel")
        if df.empty:
            st.info("No data to export yet.")
        else:
            excel_data = export_excel(df)
            st.download_button(
                label="⬇️ Download Excel",
                data=excel_data,
                file_name=f"outbound_call_duty_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
