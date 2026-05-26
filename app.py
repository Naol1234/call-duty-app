import streamlit as st
import pandas as pd
from datetime import datetime, date
import io
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Outbound Call Duty Dashboard", layout="wide", page_icon="📞")

AGENTS = [
    "FR", "MT", "AK",
    "SS", "JMW", "JMT",
    "PI", "NU", "WF", "WM"
]
DELAY_REASONS = ["", "Different time zone", "Public holiday",
    "Impossible to contact customer", "Delayed from other department", "Other"]
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
# GOOGLE SHEETS
# ─────────────────────────────────────────────
@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return gspread.authorize(creds)

def get_all_months():
    """Return list of all sheet tab names (months)."""
    try:
        sh = get_client().open_by_key(SHEET_ID)
        return [ws.title for ws in sh.worksheets()]
    except:
        return []

def load_month(month_name):
    """Load data from a specific month tab."""
    try:
        sh = get_client().open_by_key(SHEET_ID)
        ws = sh.worksheet(month_name)
        data = ws.get_all_values()
        if len(data) < 2:
            return pd.DataFrame(columns=COLUMNS)
        df = pd.DataFrame(data[1:], columns=data[0])
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
        st.error(f"Error loading {month_name}: {e}")
        return pd.DataFrame(columns=COLUMNS)

def save_month(month_name, df):
    """Save dataframe to a specific month tab."""
    try:
        sh = get_client().open_by_key(SHEET_ID)
        try:
            ws = sh.worksheet(month_name)
        except:
            ws = sh.add_worksheet(title=month_name, rows=2000, cols=20)
        df_clean = df.fillna("").astype(str)
        ws.clear()
        ws.update([df_clean.columns.tolist()] + df_clean.values.tolist())
        return True
    except Exception as e:
        st.error(f"Error saving: {e}")
        return False

def create_month_tab(month_name):
    """Create a new empty month tab."""
    try:
        sh = get_client().open_by_key(SHEET_ID)
        ws = sh.add_worksheet(title=month_name, rows=2000, cols=20)
        ws.update([COLUMNS])
        return True
    except Exception as e:
        st.error(f"Error creating tab: {e}")
        return False

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
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

def export_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
    return buf.getvalue()

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "manager_authenticated" not in st.session_state:
    st.session_state.manager_authenticated = False
if "tickets" not in st.session_state:
    st.session_state.tickets = pd.DataFrame(columns=COLUMNS)
if "active_month" not in st.session_state:
    months = get_all_months()
    st.session_state.active_month = months[-1] if months else None
if "months_list" not in st.session_state:
    st.session_state.months_list = get_all_months()

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
st.sidebar.image("https://img.icons8.com/fluency/96/phone-office.png", width=60)
st.sidebar.title("📞 Call Duty App")
st.sidebar.markdown("---")

# Month selector
months = st.session_state.months_list
if months:
    selected_month = st.sidebar.selectbox("📅 Select Month:", months, index=len(months)-1)
    if selected_month != st.session_state.active_month:
        st.session_state.active_month = selected_month
        st.session_state.tickets = load_month(selected_month)
        st.rerun()
    if st.session_state.tickets.empty or st.session_state.active_month != selected_month:
        st.session_state.tickets = load_month(selected_month)
else:
    st.sidebar.info("No months available yet.")
    selected_month = None

if st.sidebar.button("🔄 Refresh"):
    if selected_month:
        st.session_state.tickets = load_month(selected_month)
    st.session_state.months_list = get_all_months()
    st.rerun()

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
        pwd = st.text_input("Password:", type="password")
        if st.button("Login", type="primary"):
            if pwd == MANAGER_PASSWORD:
                st.session_state.manager_authenticated = True
                st.rerun()
            else:
                st.error("❌ Incorrect password.")
        st.stop()
    if st.sidebar.button("🔓 Logout"):
        st.session_state.manager_authenticated = False
        st.rerun()
    pages = ["Dashboard", "All Tickets", "Add Ticket", "Manage Tickets", "Import / Export"]
    page = st.sidebar.radio("Navigate:", pages)

df = st.session_state.tickets

# Month badge
if selected_month:
    st.sidebar.markdown(f"**Active month:** `{selected_month}`")

# ─────────────────────────────────────────────
# AGENT PAGES
# ─────────────────────────────────────────────
if role == "👤 Agent":

    if page == "My Tickets":
        st.title(f"🎫 My Tickets — {agent_name}")
        if selected_month:
            st.caption(f"📅 Viewing: {selected_month}")
        my = df[df["Assigned to (agents)"] == agent_name].copy()

        if my.empty:
            st.info("No tickets assigned to you yet.")
        else:
            unresolved = my[~(my["Case Resolved (Yes/No)"] == "Resolved")]
            resolved   = my[my["Case Resolved (Yes/No)"] == "Resolved"]

            col1, col2, col3 = st.columns(3)
            col1.metric("Total Assigned", len(my))
            col2.metric("Pending", len(unresolved))
            col3.metric("Resolved", len(resolved))

            st.markdown("### 🔴 Pending Tickets")
            if unresolved.empty:
                st.success("All tickets resolved! 🎉")
            else:
                st.dataframe(unresolved[["Case Number","Date/Time","Recent Interaction Date",
                    "Call Answered (Yes/No)","Case Resolved (Yes/No)"]],
                    use_container_width=True, hide_index=True)

            st.markdown("### ✅ Resolved Tickets")
            if resolved.empty:
                st.info("No resolved tickets yet.")
            else:
                st.dataframe(resolved[["Case Number","Date/Time","Call Answered (Yes/No)",
                    "Resolution keywords","Case Resolved (Yes/No)"]],
                    use_container_width=True, hide_index=True)

    elif page == "Fill Ticket Details":
        st.title(f"✏️ Fill Ticket Details — {agent_name}")
        if selected_month:
            st.caption(f"📅 Editing: {selected_month}")

        my = df[df["Assigned to (agents)"] == agent_name]
        if my.empty:
            st.info("No tickets assigned to you yet.")
        else:
            selected_case = st.selectbox("Select Case Number:", my["Case Number"].tolist())
            idx = df[df["Case Number"] == selected_case].index[0]
            row = df.loc[idx]

            with st.expander("📋 Ticket Info", expanded=True):
                col1, col2, col3 = st.columns(3)
                col1.text_input("Case Number 🔒", value=str(row['Case Number']), disabled=True)
                new_dt = col2.text_input("Date/Time", value=str(row['Date/Time']) if str(row['Date/Time']) not in ["None","nan"] else "")
                new_agent = col3.selectbox("Assigned to", AGENTS,
                    index=AGENTS.index(row["Assigned to (agents)"]) if row["Assigned to (agents)"] in AGENTS else 0)

            st.markdown("### Fill in the fields below")

            st.markdown("**Recent Interaction Date**")
            try:
                default_date = pd.to_datetime(row["Recent Interaction Date"]).date() \
                    if str(row["Recent Interaction Date"]) not in ["","None","nan"] else date.today()
            except:
                default_date = date.today()
            ri_date = st.date_input("Date", value=default_date, label_visibility="collapsed")
            ri_notes = st.text_area("Recent Interaction Notes",
                value=str(row["Recent Interaction Notes"]) if str(row["Recent Interaction Notes"]) not in ["None","nan"] else "",
                placeholder="Type your notes here...")

            delayed = st.selectbox("Delayed from other department?", ["","Yes","No"],
                index=["","Yes","No"].index(str(row["Delayed from other department"]))
                      if str(row["Delayed from other department"]) in ["Yes","No"] else 0)

            impossible = st.selectbox("Impossible to contact (reason):", DELAY_REASONS,
                index=DELAY_REASONS.index(str(row["Impossible to contact (reason)"]))
                      if str(row["Impossible to contact (reason)"]) in DELAY_REASONS else 0)

            call_answered = st.selectbox("Call Answered?", ["","Yes","No"],
                index=["","Yes","No"].index(str(row["Call Answered (Yes/No)"]))
                      if str(row["Call Answered (Yes/No)"]) in ["Yes","No"] else 0)

            st.markdown("---")
            st.markdown("#### 📞 Call Details")
            try:
                default_call = pd.to_datetime(row["Date of call if answered"]).date() \
                    if str(row["Date of call if answered"]) not in ["","None","nan"] else date.today()
            except:
                default_call = date.today()
            call_date = st.date_input("Date of call", value=default_call, label_visibility="collapsed")

            resolution_kw = st.text_input("📝 Resolution keywords:",
                value=str(row["Resolution keywords"]) if str(row["Resolution keywords"]) not in ["None","nan"] else "",
                placeholder="e.g. Created WO, Scheduled visit...")

            case_resolved = st.selectbox("Case Resolved?", RESOLUTION_OPTIONS,
                index=RESOLUTION_OPTIONS.index(str(row["Case Resolved (Yes/No)"]))
                      if str(row["Case Resolved (Yes/No)"]) in RESOLUTION_OPTIONS else 0)

            st.markdown("---")
            st.markdown("#### 🔁 Callback Details")
            cb1 = st.text_input("1st Callback attempt:",
                value=str(row["1st Callback attempt (date & time)"]) if str(row["1st Callback attempt (date & time)"]) not in ["None","nan"] else "",
                placeholder="e.g. 2026-05-26 10:00")
            cb2 = st.text_input("2nd Callback attempt:",
                value=str(row["2nd Callback attempt (date & time)"]) if str(row["2nd Callback attempt (date & time)"]) not in ["None","nan"] else "",
                placeholder="e.g. 2026-05-27 14:00")
            postponed = st.selectbox("Postponed to another day?", ["","Yes","No"],
                index=["","Yes","No"].index(str(row["Postponed to another day"]))
                      if str(row["Postponed to another day"]) in ["Yes","No"] else 0)

            st.markdown("---")
            st.markdown("#### ✅ Resolved Once and For All")
            col_c, col_n = st.columns([1,3])
            with col_c:
                resolved_all = st.checkbox("Resolved once and for all",
                    value=str(row["Resolved once and for all"]) in ["True","Yes","1"])
            with col_n:
                resolved_notes = st.text_input("Final notes:", placeholder="Any final notes...")

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
                st.session_state.tickets.at[idx, "Resolved once and for all"] = "Yes" if resolved_all else "No"
                with st.spinner("Saving..."):
                    if save_month(selected_month, st.session_state.tickets):
                        st.success(f"✅ Ticket {selected_case} saved!")

# ─────────────────────────────────────────────
# MANAGER PAGES
# ─────────────────────────────────────────────
else:

    if page == "Dashboard":
        st.title("📊 Manager Dashboard")
        if selected_month:
            st.caption(f"📅 Viewing: {selected_month}")

        if df.empty:
            st.warning("No tickets for this month yet.")
        else:
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Total Tickets", len(df))
            c2.metric("Resolved", (df["Case Resolved (Yes/No)"]=="Resolved").sum())
            c3.metric("Pending", (df["Case Resolved (Yes/No)"]!="Resolved").sum())
            c4.metric("Call Answered", (df["Call Answered (Yes/No)"]=="Yes").sum())

            st.markdown("---")
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("#### Tickets per Agent")
                ac = df["Assigned to (agents)"].value_counts().reset_index()
                ac.columns = ["Agent","Tickets"]
                st.bar_chart(ac.set_index("Agent"))
            with col_b:
                st.markdown("#### Resolved vs Pending per Agent")
                summary = df.groupby("Assigned to (agents)").apply(
                    lambda x: pd.Series({
                        "Resolved": (x["Case Resolved (Yes/No)"]=="Resolved").sum(),
                        "Pending": (x["Case Resolved (Yes/No)"]!="Resolved").sum()
                    })
                ).reset_index()
                st.dataframe(summary, use_container_width=True, hide_index=True)

            # All months summary
            st.markdown("---")
            st.markdown("#### 📅 All Months Overview")
            all_months = st.session_state.months_list
            if all_months:
                month_stats = []
                for m in all_months:
                    mdf = load_month(m)
                    if not mdf.empty:
                        month_stats.append({
                            "Month": m,
                            "Total": len(mdf),
                            "Resolved": (mdf["Case Resolved (Yes/No)"]=="Resolved").sum(),
                            "Pending": (mdf["Case Resolved (Yes/No)"]!="Resolved").sum(),
                        })
                if month_stats:
                    st.dataframe(pd.DataFrame(month_stats), use_container_width=True, hide_index=True)

    elif page == "All Tickets":
        st.title("📋 All Tickets")
        if selected_month:
            st.caption(f"📅 Month: {selected_month}")

        if df.empty:
            st.warning("No tickets for this month.")
        else:
            fa = st.selectbox("Filter by agent:", ["All"]+AGENTS)
            fs = st.selectbox("Filter by status:", ["All","Resolved","Not Resolved","Pending"])
            vdf = df.copy()
            if fa != "All":
                vdf = vdf[vdf["Assigned to (agents)"]==fa]
            if fs == "Resolved":
                vdf = vdf[vdf["Case Resolved (Yes/No)"]=="Resolved"]
            elif fs == "Not Resolved":
                vdf = vdf[vdf["Case Resolved (Yes/No)"]=="Not Resolved"]
            elif fs == "Pending":
                vdf = vdf[vdf["Case Resolved (Yes/No)"].isin(["","None"]) | vdf["Case Resolved (Yes/No)"].isna()]
            st.write(f"Showing **{len(vdf)}** tickets")
            st.dataframe(vdf, use_container_width=True, hide_index=True)

    elif page == "Add Ticket":
        st.title("➕ Add Tickets")
        if selected_month:
            st.caption(f"📅 Adding to: {selected_month}")

        tab1, tab2, tab3 = st.tabs(["📋 Paste from Excel", "📁 Upload Excel", "➕ Single Ticket"])

        with tab1:
            st.markdown("Copy **Case Number** + **Date/Time** from Excel and paste below:")
            pasted = st.text_area("Paste here:", height=200, placeholder="1414513\t2026-05-11 19:10\n...")
            lpdate = st.date_input("Date of listing:", value=date.today(), key="p_date")
            if st.button("➕ Add Pasted Tickets", type="primary"):
                if not pasted.strip():
                    st.error("Please paste data first.")
                else:
                    rows, skipped = [], []
                    existing = df["Case Number"].values
                    for line in pasted.strip().split("\n"):
                        parts = line.strip().split("\t")
                        if not parts: continue
                        case = str(parts[0]).strip()
                        try: case = str(int(float(case)))
                        except: pass
                        if case in existing:
                            skipped.append(case)
                            continue
                        dt_val = parts[1].strip() if len(parts)>=2 else ""
                        nr = {c:"" for c in COLUMNS}
                        nr["Case Number"]=case; nr["Date of listing"]=str(lpdate)
                        nr["Date/Time"]=dt_val; nr["Response"]="False"
                        rows.append(nr)
                    if rows:
                        ndf = auto_assign(pd.DataFrame(rows), df)
                        st.session_state.tickets = pd.concat([df, ndf], ignore_index=True)
                        with st.spinner("Saving..."):
                            save_month(selected_month, st.session_state.tickets)
                        st.success(f"✅ Added {len(rows)} tickets!")
                        st.dataframe(ndf[["Case Number","Date/Time","Assigned to (agents)"]], use_container_width=True, hide_index=True)
                    if skipped:
                        st.warning(f"Skipped {len(skipped)} duplicates.")

        with tab2:
            nfile = st.file_uploader("Upload Excel:", type=["xlsx"], key="bulk")
            ludate = st.date_input("Date of listing:", value=date.today(), key="u_date")
            if nfile:
                try:
                    pdf = pd.read_excel(nfile)
                    cc = [c for c in pdf.columns if "case" in c.lower()]
                    dc = [c for c in pdf.columns if "date" in c.lower() or "time" in c.lower()]
                    if not cc:
                        st.error("No Case Number column found.")
                    else:
                        ccol = cc[0]; dcol = dc[0] if dc else None
                        pdf = pdf.dropna(subset=[ccol])
                        pdf[ccol] = pdf[ccol].apply(lambda x: str(int(float(x))) if str(x).replace('.','').isdigit() else str(x))
                        st.write(f"Found **{len(pdf)}** tickets:")
                        st.dataframe(pdf[[ccol]+([dcol] if dcol else [])].head(10), use_container_width=True, hide_index=True)
                        if st.button("➕ Add These Tickets", type="primary"):
                            existing = df["Case Number"].values
                            rows, skipped = [], []
                            for _, r in pdf.iterrows():
                                case = str(r[ccol]).strip()
                                if case in existing:
                                    skipped.append(case); continue
                                nr = {c:"" for c in COLUMNS}
                                nr["Case Number"]=case; nr["Date of listing"]=str(ludate)
                                nr["Date/Time"]=str(r[dcol]) if dcol else ""; nr["Response"]="False"
                                rows.append(nr)
                            if rows:
                                ndf = auto_assign(pd.DataFrame(rows), df)
                                st.session_state.tickets = pd.concat([df, ndf], ignore_index=True)
                                with st.spinner("Saving..."):
                                    save_month(selected_month, st.session_state.tickets)
                                st.success(f"✅ Added {len(rows)} tickets!")
                            if skipped:
                                st.warning(f"Skipped {len(skipped)} duplicates.")
                except Exception as e:
                    st.error(f"Error: {e}")

        with tab3:
            cn = st.text_input("Case Number *")
            ld = st.date_input("Date of listing:", value=date.today(), key="s_date")
            dt = st.text_input("Date/Time:")
            auto_a = st.checkbox("Auto-assign agent", value=True)
            ag = "" if auto_a else st.selectbox("Agent:", [""]+AGENTS)
            if st.button("➕ Add Ticket", type="primary"):
                if not cn:
                    st.error("Case Number required.")
                elif cn in df["Case Number"].values:
                    st.error("Case already exists.")
                else:
                    nr = {c:"" for c in COLUMNS}
                    nr["Case Number"]=cn; nr["Date of listing"]=str(ld)
                    nr["Date/Time"]=dt; nr["Response"]="False"
                    ndf = pd.DataFrame([nr])
                    if auto_a:
                        ndf = auto_assign(ndf, df)
                        assigned = ndf["Assigned to (agents)"].iloc[0]
                    else:
                        ndf["Assigned to (agents)"]=ag; assigned=ag
                    st.session_state.tickets = pd.concat([df, ndf], ignore_index=True)
                    with st.spinner("Saving..."):
                        save_month(selected_month, st.session_state.tickets)
                    st.success(f"✅ Ticket {cn} assigned to {assigned}!")

    elif page == "Manage Tickets":
        st.title("🛠️ Manage Tickets")
        if df.empty:
            st.warning("No tickets loaded.")
        else:
            sel = st.selectbox("Select ticket:", df["Case Number"].tolist())
            idx = df[df["Case Number"]==sel].index[0]
            row = df.loc[idx]
            t1, t2 = st.tabs(["✏️ Edit", "🗑️ Delete"])
            with t1:
                na = st.selectbox("Reassign:", AGENTS,
                    index=AGENTS.index(row["Assigned to (agents)"]) if row["Assigned to (agents)"] in AGENTS else 0)
                nd = st.text_input("Date/Time:", value=str(row["Date/Time"]))
                if st.button("💾 Save", type="primary"):
                    st.session_state.tickets.at[idx,"Assigned to (agents)"]=na
                    st.session_state.tickets.at[idx,"Date/Time"]=nd
                    with st.spinner("Saving..."):
                        save_month(selected_month, st.session_state.tickets)
                    st.success("✅ Updated!")
            with t2:
                st.warning(f"Delete ticket **{sel}**?")
                if st.button("🗑️ Confirm Delete", type="primary"):
                    st.session_state.tickets = df.drop(index=idx).reset_index(drop=True)
                    with st.spinner("Saving..."):
                        save_month(selected_month, st.session_state.tickets)
                    st.success("Deleted!")
                    st.rerun()

    elif page == "Import / Export":
        st.title("📂 Import / Export")

        # Create new month
        st.markdown("### 📅 Create New Month")
        col1, col2 = st.columns([2,1])
        with col1:
            new_month = st.text_input("Month name (e.g. June 2026):")
        with col2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕ Create Month"):
                if new_month:
                    if new_month in st.session_state.months_list:
                        st.error("Month already exists!")
                    else:
                        with st.spinner("Creating..."):
                            if create_month_tab(new_month):
                                st.session_state.months_list = get_all_months()
                                st.session_state.active_month = new_month
                                st.session_state.tickets = pd.DataFrame(columns=COLUMNS)
                                st.success(f"✅ Month '{new_month}' created!")
                                st.rerun()

        st.markdown("---")
        st.markdown(f"### 📥 Import Excel into **{selected_month}**")
        st.warning("⚠️ This will replace all data in the current month tab only. Other months are safe!")
        uploaded = st.file_uploader("Upload Excel (.xlsx):", type=["xlsx"])
        if uploaded:
            try:
                dfl = pd.read_excel(uploaded, sheet_name=0, header=1)
                dfl = dfl.dropna(how="all")
                col_map = {
                    "Response":"Response",
                    "Date of listing case in this file":"Date of listing",
                    "Case Number":"Case Number","Date/Time":"Date/Time",
                    "Assigned to (agents)":"Assigned to (agents)",
                    "Recent interaction (5 days)":"Recent Interaction Date",
                    "Delayed from other department ":"Delayed from other department",
                    "Is it  impossible to contact the customer: (due to different timezone, Public holidays, etc)":"Impossible to contact (reason)",
                    "Called answered (Yes or No)":"Call Answered (Yes/No)",
                    "Date of the call if answered":"Date of call if answered",
                    "if Yes, what is the resolution, key words":"Resolution keywords",
                    "Case resolved  (Yes or No)":"Case Resolved (Yes/No)",
                    "If no, 1st attempt time call back (date and time)":"1st Callback attempt (date & time)",
                    "If no, 2nd attempt call back (date abbd time)":"2nd Callback attempt (date & time)",
                    "to be postponed to another day (Yes, No)":"Postponed to another day",
                    "Resolved once and for all":"Resolved once and for all",
                }
                dfl = dfl.rename(columns=col_map)
                for c in COLUMNS:
                    if c not in dfl.columns: dfl[c]=""
                dfl = dfl[COLUMNS]
                dfl = dfl[dfl["Case Number"].notna()]
                dfl["Case Number"] = dfl["Case Number"].apply(
                    lambda x: str(int(float(x))) if str(x).replace('.','').isdigit() else str(x)
                ).str.strip()
                st.success(f"✅ Found {len(dfl)} tickets!")
                st.dataframe(dfl.head(10), use_container_width=True, hide_index=True)
                if st.button("✅ Confirm Import", type="primary"):
                    st.session_state.tickets = dfl
                    with st.spinner("Saving to Google Sheets..."):
                        save_month(selected_month, dfl)
                    st.success(f"✅ Imported into {selected_month}!")
                    st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

        st.markdown("---")
        st.markdown(f"### 📤 Export **{selected_month}** to Excel")
        if df.empty:
            st.info("No data to export.")
        else:
            st.download_button("⬇️ Download Excel", export_excel(df),
                file_name=f"call_duty_{selected_month}_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
