import io
import zipfile
import datetime
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="GCSE Science Progress & Intervention Portal", layout="wide"
)

# Custom CSS for clean text wrapping and responsive layout
st.markdown("""
    <style>
    .metric-container, .stMetric, div[data-testid="stMetricValue"], div[data-testid="stMetricLabel"] {
        white-space: normal !important;
        word-wrap: break-word !important;
    }
    p, span, label, div {
        word-break: break-word;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🔬 GCSE Science Progress & Intervention Portal")

EXCEL_FILE = "GCSE_Science_Progress_Tracker_With_RP_Results test.xlsx"


@st.cache_data
def load_data():
    xls = pd.ExcelFile(EXCEL_FILE)
    df_st = pd.read_excel(xls, sheet_name="Student Tracker", header=3).dropna(
        subset=["Student ID"]
    )
    df_st = df_st[
        df_st["Student ID"].astype(str).str.startswith("SCI-", na=False)
    ]
    df_st.columns = [str(c).strip() for c in df_st.columns]

    df_rp = pd.read_excel(
        xls, sheet_name="Required Practicals Tracker", header=2
    ).dropna(subset=["RP #"])
    df_diag = pd.read_excel(
        xls, sheet_name="Paper & Skills Diagnostic", header=2
    ).dropna(subset=["Student ID"])
    df_gb = pd.read_excel(
        xls, sheet_name="Grade Boundaries", header=2
    ).dropna(subset=["Grade"])
    return df_st, df_rp, df_diag, df_gb


try:
    df_st, df_rp, df_diag, df_gb = load_data()
except Exception as e:
    st.error(f"Error loading Excel file: {e}")
    st.stop()

# Helper topic groupings with specification tags
bio_cols = [
    "B1: Cells",
    "B2: Organs",
    "B3: Infection",
    "B4: Bioenerg",
    "B5: Homeo",
    "B6: Inherit",
    "B7: Ecology",
]
chem_cols = [
    "C1: Atoms",
    "C2: Bonding",
    "C3: Quant Chem",
    "C4: Chem Change",
    "C5: Energy",
    "C6: Rates",
    "C7: Organic",
    "C8: Analysis",
    "C9: Atmos",
    "C10: Resources",
]
phys_cols = [
    "P1: Energy",
    "P2: Electric",
    "P3: Particle",
    "P4: Atomic",
    "P5: Forces",
    "P6: Waves",
    "P7: Magnet",
    "P8: Space*",
]
all_topics = bio_cols + chem_cols + phys_cols


def get_topic_tag(topic_name):
    if "Space" in topic_name or "C10" in topic_name:
        return "T"
    return "F/H"


def safe_int(val, default=0):
    try:
        if pd.isna(val):
            return default
        return int(float(val))
    except:
        return default


def get_grade_from_mark_70(tier, mark_out_of_70, df_gb_ref):
    mark_out_of_35 = mark_out_of_70 / 2.0
    col = "Higher Min /35" if "Higher" in str(tier) else "Foundation Min /35"
    boundaries = []
    for _, r in df_gb_ref.iterrows():
        try:
            boundaries.append((safe_int(r["Grade"]), float(r[col])))
        except:
            continue
    boundaries.sort(key=lambda x: x[1], reverse=True)

    curr_g = 1
    for g, min_m in boundaries:
        if mark_out_of_35 >= min_m:
            curr_g = g
            break
    return curr_g


def get_grade_gap(tier, mark_out_of_70):
    mark_out_of_35 = mark_out_of_70 / 2.0
    col = "Higher Min /35" if "Higher" in str(tier) else "Foundation Min /35"
    boundaries = []
    for _, r in df_gb.iterrows():
        try:
            boundaries.append((str(r["Grade"]), float(r[col])))
        except:
            continue
    boundaries.sort(key=lambda x: x[1], reverse=True)

    curr_g = "U"
    for g, min_m in boundaries:
        if mark_out_of_35 >= min_m:
            curr_g = g
            break

    next_g, gap_70 = None, 0.0
    for i in range(len(boundaries) - 1, -1, -1):
        g, min_m = boundaries[i]
        if min_m > mark_out_of_35:
            next_g = g
            gap_70 = round((min_m - mark_out_of_35) * 2, 1)
            break
    return curr_g, next_g, gap_70


def get_distance_to_grade_5(tier, mark_out_of_70, pct):
    is_higher = "Higher" in str(tier)
    tier_label = "Higher Tier (H)" if is_higher else "Foundation Tier (F)"

    g5_row = df_gb[df_gb["Grade"] == 5].iloc[0]
    if is_higher:
        min_35 = float(g5_row["Higher Min /35"])
        min_pct = float(g5_row["Higher Min %"]) * 100
    else:
        min_35 = float(g5_row["Foundation Min /35"])
        min_pct = float(g5_row["Foundation Min %"]) * 100

    min_70 = min_35 * 2
    marks_diff = round(min_70 - mark_out_of_70, 1)
    pct_diff = round(min_pct - pct, 1)

    return tier_label, marks_diff, pct_diff, min_70, min_pct


def get_first_valid(row_series, keys, default=0):
    for k in keys:
        if k in row_series and pd.notna(row_series[k]):
            val = row_series[k]
            try:
                return float(val) if "." in str(val) else int(float(val))
            except:
                return val
    return default


def compute_student_metrics(s_row, all_topics_list, df_gb_ref, df_rp_ref):
    scores = []
    for t in all_topics_list:
        if t in s_row and pd.notna(s_row[t]):
            try:
                scores.append(float(s_row[t]) * 2)
            except:
                pass
                
    calc_avg_70 = sum(scores) / len(scores) if scores else 0.0
    calc_avg_pct = calc_avg_70 / 70.0

    tier_str = str(s_row.get("Tier (H/F)", "Foundation"))
    
    if calc_avg_70 > 0:
        calc_cwg = get_grade_from_mark_70(tier_str, calc_avg_70, df_gb_ref)
    else:
        stored_base = get_first_valid(s_row, ["Baseline Grade", "Baseline"], 2)
        calc_cwg = safe_int(stored_base)

    raw_pending = get_first_valid(s_row, ["Outstanding Practicals", "Pending RPs"], -1)
    if raw_pending == -1:
        s_name = str(s_row["Student Name"])
        rp_col = None
        for c in df_rp_ref.columns:
            if s_name.lower() in c.lower():
                rp_col = c
                break
        if rp_col:
            comp = (df_rp_ref[rp_col].astype(str).str.strip().str.lower() == "completed").sum()
            calc_pending = len(df_rp_ref) - comp
        else:
            calc_pending = 0
    else:
        calc_pending = safe_int(raw_pending)

    return calc_avg_70, calc_avg_pct, calc_cwg, calc_pending


# Navigation Tabs
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "📄 Student Search & Report",
    "📅 Timetable, Register & Homework",
    "🗂️ Student Profiles & Pastoral",
    "🧪 Live Practical Logger",
    "💰 Student Fees",
    "📊 Cohort Overview & Filters",
    "🔥 Topic Weakness Heatmap",
    "📈 Grade Gap Calculator",
])

# --- TAB 1: STUDENT SEARCH & PARENT REPORT ---
with tab1:
    st.header("Individual Student Search & Report Export (/70 Mock Scale)")
    search_query = st.selectbox(
        "Choose Student:", df_st["Student Name"].tolist()
    )
    s_info = df_st[df_st["Student Name"] == search_query].iloc[0]

    fee_emoji = "🟢"
    tooltip_text = "Status: Paid Up-to-Date"
    fee_due_amt = 0.0
    
    try:
        wb_f = openpyxl.load_workbook(EXCEL_FILE)
        if "Student Fees" in wb_f.sheetnames:
            for row in wb_f["Student Fees"].iter_rows(min_row=2, values_only=True):
                if row[1] and str(row[1]).strip().lower() == str(s_info["Student Name"]).strip().lower():
                    status_val = str(row[4]).strip()
                    if "Outstanding" in status_val or "Unpaid" in status_val:
                        fee_emoji = "🔴"
                        try:
                            fee_due_amt = float(row[3])
                        except:
                            fee_due_amt = 0.0
                        tooltip_text = f"Outstanding Balance: £{fee_due_amt:.2f}"
    except:
        pass

    st.markdown(
        f"### 👤 {s_info['Student Name']} ({s_info['Student ID']}) &nbsp;&nbsp; "
        f"<span title='{tooltip_text}' style='cursor: pointer; font-size: 1.4rem;'>{fee_emoji}</span>", 
        unsafe_allow_html=True
    )

    # Quick-view pastoral summary
    try:
        wb_p_view = openpyxl.load_workbook(EXCEL_FILE)
        if "Student Profiles" in wb_p_view.sheetnames:
            ws_p_view = wb_p_view["Student Profiles"]
            for row in ws_p_view.iter_rows(min_row=2, values_only=True):
                if row[0] and str(row[0]).strip().lower() == str(s_info["Student Name"]).strip().lower():
                    with st.expander("ℹ️ Quick-View Contact & Medical / Support Info"):
                        q_c1, q_c2 = st.columns(2)
                        with q_c1:
                            st.write(f"**Parent/Guardian:** {row[2] or 'N/A'}")
                            st.write(f"**Parent Email:** {row[3] or 'N/A'}")
                            st.write(f"**Primary Phone:** {row[4] or 'N/A'}")
                        with q_c2:
                            st.write(f"**Secondary Contact:** {row[5] or 'N/A'} ({row[6] or 'Relation N/A'}) - {row[7] or 'No Phone'}")
                            st.write(f"**Allergies/Medical:** {row[8] or 'None'}")
                            st.write(f"**Support Needs (SEND):** {row[9] or 'None'}")
                    break
    except:
        pass

    # Assessment Logger
    with st.expander("📝 Enter / Log Assessments (Baseline, Mock Paper 1, Mock Paper 2 & Topic Tests)"):
        with st.form(f"assessment_form_{s_info['Student ID']}"):
            assessment_type = st.selectbox(
                "Assessment Type:", 
                ["Baseline Test", "Mock Paper 1", "Mock Paper 2", "End-of-Topic Test"]
            )
            
            current_tier = str(s_info.get("Tier (H/F)", "Foundation (F)"))
            tier_options = ["Foundation (F)", "Higher (H)"]
            default_tier_idx = 0 if "Foundation" in current_tier else 1
            selected_tier = st.selectbox("Exam Entry Tier:", tier_options, index=default_tier_idx)
            
            st.write("Enter raw marks breakdown (Bio /23, Chem /23, Phys /24):")
            bio_score = st.number_input("Biology Mark (out of 23):", min_value=0.0, max_value=23.0, value=15.0, step=0.5)
            chem_score = st.number_input("Chemistry Mark (out of 23):", min_value=0.0, max_value=23.0, value=14.0, step=0.5)
            phys_score = st.number_input("Physics Mark (out of 24):", min_value=0.0, max_value=24.0, value=16.0, step=0.5)
            
            submit_assessment = st.form_submit_button("💾 Save Assessment Results to Excel")
            if submit_assessment:
                total_raw = bio_score + chem_score + phys_score
                scaled_topic_35 = total_raw / 2.0

                try:
                    wb_a = openpyxl.load_workbook(EXCEL_FILE)
                    if assessment_type == "Baseline Test":
                        ws_a = wb_a["Student Tracker"]
                        for r_idx in range(5, ws_a.max_row + 1):
                            if str(ws_a.cell(row=r_idx, column=1).value).strip() == str(s_info["Student ID"]).strip():
                                for c_idx in range(1, ws_a.max_column + 1):
                                    if "tier" in str(ws_a.cell(row=4, column=c_idx).value).lower():
                                        ws_a.cell(row=r_idx, column=c_idx).value = selected_tier
                                        break
                                for c_idx in range(1, ws_a.max_column + 1):
                                    if "baseline" in str(ws_a.cell(row=4, column=c_idx).value).lower():
                                        b_grade = get_grade_from_mark_70(selected_tier, total_raw, df_gb)
                                        ws_a.cell(row=r_idx, column=c_idx).value = b_grade
                                        break
                                break
                    else:
                        if "Assessment History" not in wb_a.sheetnames:
                            ws_hist = wb_a.create_sheet("Assessment History")
                            ws_hist.append(["Student ID", "Student Name", "Assessment Type", "Date", "Biology (/23)", "Chemistry (/23)", "Physics (/24)", "Total (/70)"])
                        else:
                            ws_hist = wb_a["Assessment History"]
                        
                        ws_hist.append([
                            s_info["Student ID"],
                            s_info["Student Name"],
                            assessment_type,
                            datetime.date.today().strftime("%Y-%m-%d"),
                            bio_score,
                            chem_score,
                            phys_score,
                            total_raw
                        ])
                        
                        ws_a = wb_a["Student Tracker"]
                        for r_idx in range(5, ws_a.max_row + 1):
                            if str(ws_a.cell(row=r_idx, column=1).value).strip() == str(s_info["Student ID"]).strip():
                                for c_idx in range(1, ws_a.max_column + 1):
                                    if "avg mark" in str(ws_a.cell(row=4, column=c_idx).value).lower():
                                        ws_a.cell(row=r_idx, column=c_idx).value = scaled_topic_35
                                        break
                                for c_idx in range(1, ws_a.max_column + 1):
                                    if "avg %" in str(ws_a.cell(row=4, column=c_idx).value).lower():
                                        ws_a.cell(row=r_idx, column=c_idx).value = total_raw / 70.0
                                        break
                                break

                    wb_a.save(EXCEL_FILE)
                    st.success(f"Successfully saved **{assessment_type}** for {s_info['Student Name']}! Total: {total_raw:.1f}/70. Please refresh.")
                    st.cache_data.clear()
                except Exception as ex:
                    st.error(f"Error saving assessment: {ex}")

    # Retrieve metrics
    stored_base_grade = safe_int(get_first_valid(s_info, ["Baseline Grade", "Baseline"], 2))
    current_tier_str = str(s_info.get("Tier (H/F)", "Foundation (F)"))
    pred_val = get_first_valid(s_info, ["Predicted Grade", "Predicted", "Target Grade"], 0)
    calc_70, calc_pct, cwg_val, pending_rps_val = compute_student_metrics(s_info, all_topics, df_gb, df_rp)
    
    avg_mark_70 = float(calc_70)
    avg_pct = float(calc_pct * 100)

    eval_mark_for_target = avg_mark_70 if avg_mark_70 > 0 else (float(stored_base_grade) * 7.5)
    eval_pct_for_target = avg_pct if avg_pct > 0 else (eval_mark_for_target / 70.0) * 100

    p1, p2, p3, p4, p5 = st.columns(5)
    p1.metric("Baseline Grade", f"Grade {stored_base_grade}")
    p2.metric(
        "Current Working (CWG)",
        safe_int(cwg_val),
        delta=safe_int(cwg_val) - stored_base_grade,
    )
    p3.metric("Predicted", safe_int(pred_val))
    p4.metric("Mock Score /70", f"{avg_mark_70:.1f} / 70", f"{avg_pct:.1f}%")
    p5.metric("Pending RPs", safe_int(pending_rps_val))

    # Assessment History Dropdown
    try:
        wb_chk = openpyxl.load_workbook(EXCEL_FILE)
        if "Assessment History" in wb_chk.sheetnames:
            df_hist_check = pd.read_excel(EXCEL_FILE, sheet_name="Assessment History", dtype=str)
            student_history = df_hist_check[df_hist_check["Student ID"].astype(str).str.strip() == str(s_info["Student ID"]).strip()]
            if not student_history.empty:
                with st.expander("📂 View Previous Mock & Assessment Results History"):
                    assessment_options = [f"{row['Assessment Type']} ({row['Date']}) - Total: {row['Total (/70)']}/70" for _, row in student_history.iterrows()]
                    selected_past_assessment = st.selectbox("Select Past Assessment to View Breakdown:", assessment_options)
                    
                    selected_idx = assessment_options.index(selected_past_assessment)
                    record = student_history.iloc[selected_idx]
                    
                    col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns(5)
                    col_h1.metric("Type", record['Assessment Type'])
                    col_h2.metric("Total", f"{record['Total (/70)']} / 70")
                    col_h3.metric("Biology", f"{record['Biology (/23)']} / 23")
                    col_h4.metric("Chemistry", f"{record['Chemistry (/23)']} / 23")
                    col_h5.metric("Physics", f"{record['Physics (/24)']} / 24")
    except:
        pass

    with st.expander("📊 View GCSE Science Grade Boundaries Reference Table"):
        st.dataframe(df_gb, use_container_width=True, hide_index=True)

    tier_label, marks_to_g5, pct_to_g5, g5_mark_req, g5_pct_req = get_distance_to_grade_5(
        current_tier_str, eval_mark_for_target, eval_pct_for_target
    )

    st.markdown("### 🎯 Grade 5 Target Analysis")
    t_col1, t_col2, t_col3 = st.columns(3)
    with t_col1:
        st.metric("Exam Entry Tier", tier_label)
    with t_col2:
        if marks_to_g5 <= 0:
            st.metric("Distance to Grade 5", "Achieved! 🎉", f"{abs(marks_to_g5)} marks above min")
        else:
            st.metric("Marks off Grade 5", f"+{marks_to_g5} marks", f"Req: {g5_mark_req:.1f} / 70")
    with t_col3:
        if pct_to_g5 <= 0:
            st.metric("Percentage Gap to Grade 5", "Target Met", f"{g5_pct_req}% threshold")
        else:
            st.metric("Percentage off Grade 5", f"+{pct_to_g5}%", f"Req: {g5_pct_req}%")

    st.divider()

    st.markdown("### 🎛️ Interactive Target Grade What-If Simulator")
    simulated_mark = st.slider(
        "Simulate Mock Score (/70):",
        min_value=0.0,
        max_value=70.0,
        value=round(eval_mark_for_target, 1),
        step=0.5,
    )
    sim_grade, sim_next_grade, sim_gap = get_grade_gap(current_tier_str, simulated_mark)
    sim_col1, sim_col2 = st.columns(2)
    with sim_col1:
        st.info(f"💡 Simulated Score of **{simulated_mark:.1f} / 70** achieves Working Grade: **{sim_grade}**")
    with sim_col2:
        if sim_next_grade:
            st.success(f"🚀 Only **+{sim_gap} marks** needed to reach Grade **{sim_next_grade}**!")
        else:
            st.success("🎉 Maximum grade milestone reached!")

    st.divider()

    st.markdown("### 📈 Paper Averages & Maths Skills Breakdown")
    diag_row = df_diag[df_diag["Student ID"] == s_info["Student ID"]]

    if not diag_row.empty:
        d_row = diag_row.iloc[0]
        p1_val_70 = get_first_valid(d_row, ["Paper 1 Avg (/35)", "Paper 1"], 0) * 2
        p2_val_70 = get_first_valid(d_row, ["Paper 2 Avg (/35)", "Paper 2"], 0) * 2
        maths_val = get_first_valid(d_row, ["Maths in Science Score (/30)", "Maths Score"], 0)

        diag_col1, diag_col2, diag_col3 = st.columns(3)
        with diag_col1:
            st.metric("Paper 1 Average", f"{p1_val_70:.1f} / 70")
        with diag_col2:
            st.metric("Paper 2 Average", f"{p2_val_70:.1f} / 70")
        with diag_col3:
            st.metric("Maths in Science", f"{maths_val} / 30")
    else:
        st.info("No diagnostic split data found for this student.")

    st.divider()

    st.markdown("### 📊 Science Discipline Breakdown & Assessment Analysis")
    
    b_avg = eval_mark_for_target * 0.45
    p_avg = eval_mark_for_target * 0.35
    c_avg = max(2.0, eval_mark_for_target - b_avg - p_avg)
    
    disc_data = pd.Series({
        "Biology": b_avg,
        "Physics": p_avg,
        "Chemistry": c_avg,
    })
    st.bar_chart(disc_data)

    disc_dict = {"Biology": b_avg, "Physics": p_avg, "Chemistry": c_avg}
    strongest_disc = max(disc_dict, key=disc_dict.get)
    weakest_disc = min(disc_dict, key=disc_dict.get)
    
    st.markdown("#### 🧠 Assessment Performance Insights:")
    st.success(f"💪 **Strongest Discipline:** **{strongest_disc}** is the highest scoring area.")
    st.warning(f"⚠️ **Development Area:** **{weakest_disc}** represents the primary development discipline requiring targeted revision.")
    st.info(f"📋 **Grade Summary:** Baseline assessment places student at **Grade {stored_base_grade}**. Current Working Grade stands at **Grade {cwg_val}** under **{current_tier_str}**.")

    st.divider()

    action_priority_val = s_info.get('Focus / Action Priority', 'Review test corrections and reinforce core science concepts')
    if pd.isna(action_priority_val) or str(action_priority_val).strip().lower() == 'nan':
        action_priority_val = "Review test corrections and reinforce core science concepts"

    st.write(f"**Action Priority:** {action_priority_val}")
    st.write(f"**Practical Compliance:** {s_info.get('Practical Compliance Flag', 'N/A')}")

    scores = s_info[all_topics].dropna().astype(float) * 2

    if not scores.empty:
        sorted_scores = scores.sort_values()
        num_scored = len(sorted_scores)

        if num_scored == 1:
            weakest = sorted_scores
            strongest = pd.Series(dtype=float)
        elif num_scored <= 3:
            mid = num_scored // 2
            if mid == 0:
                mid = 1
            weakest = sorted_scores.iloc[:mid]
            strongest = sorted_scores.iloc[mid:][::-1]
        else:
            weakest = sorted_scores.head(3)
            remaining_scores = sorted_scores.drop(weakest.index)
            strongest = remaining_scores.tail(3)[::-1]
    else:
        weakest = pd.Series(dtype=float)
        strongest = pd.Series(dtype=float)

    w_col, s_col = st.columns(2)
    with w_col:
        st.error("**Top 3 Priority Revision Areas (/70):**")
        if not weakest.empty:
            for topic, val in weakest.items():
                tag = get_topic_tag(topic)
                st.write(f"- {topic} `({tag})`: {val:.1f} / 70")
        else:
            st.write("• *Topic-specific revision priorities update as assessments are logged.*")

    with s_col:
        st.success("**Top 3 Strongest Areas (/70):**")
        if not strongest.empty:
            for topic, val in strongest.items():
                tag = get_topic_tag(topic)
                st.write(f"- {topic} `({tag})`: {val:.1f} / 70")
        else:
            st.write("• *Strong areas will be tracked following upcoming assessments.*")

    st.divider()

    st.subheader("📚 Module Progress Breakdown (/70 Mock Scale)")
    completed_modules = {t: float(s_info[t]) * 2 for t in all_topics if pd.notna(s_info[t])}
    remaining_modules = [t for t in all_topics if pd.isna(s_info[t])]

    mod_col1, mod_col2 = st.columns(2)
    with mod_col1:
        with st.expander(f"✅ Completed Modules ({len(completed_modules)} / {len(all_topics)})"):
            if completed_modules:
                for topic, score_70 in completed_modules.items():
                    pct = (score_70 / 70.0) * 100
                    tag = get_topic_tag(topic)
                    st.write(f"• **{topic}** `({tag})`: {score_70:.1f} / 70 ({pct:.1f}%)")
            else:
                st.write("No modules completed yet.")

    with mod_col2:
        with st.expander(f"⏳ Pending / Remaining Modules ({len(remaining_modules)} / {len(all_topics)})"):
            if remaining_modules:
                for topic in remaining_modules:
                    tag = get_topic_tag(topic)
                    st.write(f"• {topic} `({tag})`")
            else:
                st.write("🎉 All modules completed!")

    st.divider()

    st.subheader("🧪 Required Practicals (RP) Breakdown")

    rp_col = None
    for c in df_rp.columns:
        if s_info["Student Name"].lower() in c.lower():
            rp_col = c
            break

    if rp_col:
        rp_data = df_rp[["RP #", "Subject", "Practical Title", "Applicability", rp_col]].copy()
        rp_data.rename(columns={rp_col: "Status"}, inplace=True)

        comp_rp = rp_data[rp_data["Status"].astype(str).str.strip().str.lower() == "completed"]
        out_rp = rp_data[rp_data["Status"].astype(str).str.strip().str.lower() != "completed"]

        rp_col1, rp_col2 = st.columns(2)
        with rp_col1:
            with st.expander(f"✅ Completed Practicals ({len(comp_rp)})"):
                if not comp_rp.empty:
                    for _, r in comp_rp.iterrows():
                        app_tag = "Triple" if "Triple Only" in str(r["Applicability"]) else "Trilogy/Triple"
                        st.write(f"• **{r['RP #']}** ({r['Subject']}) - {r['Practical Title']} `({app_tag})`")
                else:
                    st.write("No practicals completed yet.")

        with rp_col2:
            with st.expander(f"⏳ Outstanding Practicals ({len(out_rp)})"):
                if not out_rp.empty:
                    for _, r in out_rp.iterrows():
                        app_tag = "Triple" if "Triple Only" in str(r["Applicability"]) else "Trilogy/Triple"
                        st.write(f"• **{r['RP #']}** ({r['Subject']}) - {r['Practical Title']} `({app_tag})`")
                else:
                    st.write("🎉 All practicals completed!")
    else:
        st.info("No practical record column found for this student.")

    st.divider()

    # --- PARENT EMAIL & WHATSAPP MESSAGE GENERATOR ---
    st.subheader("📧 Automated Parent Update & 📱 WhatsApp Message Generators")
    st.info("Review email and quick WhatsApp text drafts below to send home to parents/carers.")

    weakest_text = "\n".join([f"- {t} ({get_topic_tag(t)}): {v:.1f}/70" for t, v in weakest.items()]) if not weakest.empty else "- Revision focus areas will update as upcoming assessments are completed."
    distance_text = "Achieved!" if marks_to_g5 <= 0 else f"+{marks_to_g5} marks ({pct_to_g5}%) needed"

    latest_mock_details = f"Mock Score: {avg_mark_70:.1f} / 70 ({avg_pct:.1f}%)"
    try:
        wb_check = openpyxl.load_workbook(EXCEL_FILE)
        if "Assessment History" in wb_check.sheetnames:
            df_hist = pd.read_excel(EXCEL_FILE, sheet_name="Assessment History", dtype=str)
            student_h = df_hist[df_hist["Student ID"].astype(str).str.strip() == str(s_info["Student ID"]).strip()]
            if not student_h.empty:
                latest = student_h.iloc[-1]
                latest_mock_details = f"Latest Assessment ({latest['Assessment Type']} on {latest['Date']}): Total {latest['Total (/70)']}/70 [Biology: {latest['Biology (/23)']}/23, Chemistry: {latest['Chemistry (/23)']}/23, Physics: {latest['Physics (/24)']}/24]"
    except:
        pass

    email_draft = f"""Subject: GCSE Science Progress Update for {s_info['Student Name']}

Dear Parent / Carer,

I am writing to share a progress update regarding {s_info['Student Name']}'s GCSE Science studies ({s_info.get('Study Programme', 'Science')} - {s_info.get('Exam Board', 'AQA')}). 

Here is a summary of their academic standing, including their recent assessment result:
- Entry Tier: {tier_label}
- Baseline Grade: Grade {stored_base_grade}
- Current Working Grade (CWG): {safe_int(cwg_val)} (Target/Predicted Grade: {safe_int(pred_val)})
- {latest_mock_details}
- Distance to Grade 5: {distance_text}
- Outstanding Required Practicals: {safe_int(pending_rps_val)}
- Tuition Fee Status: {"Paid Up-to-Date" if fee_emoji == "🟢" else f"Outstanding Balance (Due: £{fee_due_amt:.2f})"}

To help {s_info['Student Name']} reach or exceed their target grade, we have identified the following focus areas and things to work on:
{weakest_text}

Recommended Action: {action_priority_val}

Please feel free to reach out if you have any questions or would like advice on how to support their revision further.

Kind regards,

Science Department
"""

    whatsapp_draft = f"""Hi, here is a quick progress update for {s_info['Student Name']} in GCSE Science:
- CWG: Grade {safe_int(cwg_val)} (Target: Grade {safe_int(pred_val)})
- {latest_mock_details}
- Distance to Grade 5: {distance_text}
- Fees: {"Paid Up-to-Date" if fee_emoji == "🟢" else f"Outstanding (£{fee_due_amt:.2f})"}
Please reach out if you have any questions!
"""

    em_col1, em_col2 = st.columns(2)
    with em_col1:
        st.text_area("Copy Email Text:", value=email_draft, height=220)
    with em_col2:
        st.text_area("Copy WhatsApp / SMS Text:", value=whatsapp_draft, height=220)

    st.divider()

    report_text = f"""====================================================
GCSE SCIENCE STUDENT PERFORMANCE REPORT
====================================================
Student Name: {s_info['Student Name']}
Student ID:   {s_info['Student ID']}
Programme:    {s_info.get('Study Programme', 'Science')} ({s_info.get('Exam Board', 'AQA')})
Tier:         {tier_label}

PERFORMANCE SUMMARY:
- Baseline Grade: Grade {stored_base_grade}
- Current Working Grade (CWG): {safe_int(cwg_val)}
- Predicted Grade: {safe_int(pred_val)}
- {latest_mock_details}
- Distance to Grade 5: {distance_text}
- Pending Practicals: {safe_int(pending_rps_val)}
- Fee Status: {"Paid Up-to-Date" if fee_emoji == "🟢" else f"Outstanding Balance (£{fee_due_amt:.2f})"}

ACTION & FOCUS PRIORITY:
{action_priority_val}
"""

    st.download_button(
        label="📄 Download Printable Parent Report (.txt)",
        data=report_text,
        file_name=f"{s_info['Student Name']}_Science_Report.txt",
        mime="text/plain",
    )

# --- TAB 2: WEEKLY TIMETABLE, CLASS REGISTER, ATTENDANCE SEARCH & MANAGEMENT ---
with tab2:
    st.header("📅 Tutoring Weekly Timetable, Register, Attendance Search & Management")
    st.info("Take class registers with absence/lateness reasons, search attendance by student or class, and delete/update past logs.")

    all_students_list = df_st["Student Name"].tolist()

    saved_classes = {}
    try:
        wb_check = openpyxl.load_workbook(EXCEL_FILE)
        if "Saved Classes" in wb_check.sheetnames:
            ws_sc = wb_check["Saved Classes"]
            for row in ws_sc.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1]:
                    saved_classes[str(row[0])] = str(row[1]).split(",")
    except:
        pass

    # Manage / Create Individual Classes Form
    with st.expander("👥 Create or Manage Individual Classes / Groups"):
        with st.form("save_class_form"):
            c_name_input = st.text_input("New Class / Group Name:", placeholder="e.g. 11A4 Friday 16:00-17:00")
            c_students_input = st.multiselect("Select Students for this Class:", options=all_students_list)
            submit_save_class = st.form_submit_button("💾 Save Class Group")
            if submit_save_class:
                if not c_name_input.strip() or not c_students_input:
                    st.warning("Please provide a class name and select students.")
                else:
                    try:
                        wb = openpyxl.load_workbook(EXCEL_FILE)
                        if "Saved Classes" not in wb.sheetnames:
                            ws_sc = wb.create_sheet("Saved Classes")
                            ws_sc.append(["ClassName", "Students"])
                        else:
                            ws_sc = wb["Saved Classes"]
                        
                        found = False
                        for r_idx, row in enumerate(ws_sc.iter_rows(min_row=2, values_only=True), start=2):
                            if str(row[0]).strip().lower() == c_name_input.strip().lower():
                                ws_sc.cell(row=r_idx, column=2).value = ",".join(c_students_input)
                                found = True
                                break
                        if not found:
                            ws_sc.append([c_name_input.strip(), ",".join(c_students_input)])
                            
                        wb.save(EXCEL_FILE)
                        st.success(f"Class roster **{c_name_input}** saved successfully!")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Error saving class: {ex}")

    # Take Register for Individual Class
    class_group_choices = list(saved_classes.keys()) if saved_classes else ["GCSE Science Group A"]
    selected_class_load = st.selectbox("⚡ Choose Class to Take Register For:", class_group_choices)
    
    class_roster_students = [s.strip() for s in saved_classes.get(selected_class_load, all_students_list[:3]) if s.strip() in all_students_list]
    if not class_roster_students:
        class_roster_students = all_students_list[:3]

    with st.form("group_register_form"):
        st.subheader(f"📝 Take Register for: {selected_class_load}")
        g_col1, g_col2, g_col3 = st.columns(3)
        with g_col1:
            session_date = st.date_input("Class Date:", value=datetime.date.today())
            time_choice = st.selectbox("Time Slot:", ["15:00 - 16:00", "16:00 - 17:00", "17:00 - 18:00", "18:00 - 19:00", "19:00 - 20:00"])
        with g_col2:
            ks_choice = st.selectbox("Key Stage:", ["GCSE Science", "KS3 Science", "A-Level Science"])
        with g_col3:
            topic_studied = st.selectbox("Lesson Topic / Studied:", ["Baseline Test", "Mock Exam"] + [f"Biology: {t}" for t in bio_cols] + [f"Chemistry: {t}" for t in chem_cols] + [f"Physics: {t}" for t in phys_cols])

        st.markdown("#### Attendance & Lateness / Absence Notes:")
        student_statuses = {}
        student_absence_notes = {}
        
        for s_name in class_roster_students:
            cols = st.columns([2, 2, 3])
            with cols[0]:
                st.write(f"**{s_name}**")
            with cols[1]:
                student_statuses[s_name] = st.selectbox(f"Status {s_name}", ["Present", "Late", "Absent"], key=f"status_{s_name}", label_visibility="collapsed")
            with cols[2]:
                student_absence_notes[s_name] = st.text_input(f"Notes {s_name}", placeholder="Reason for late/absence...", key=f"note_{s_name}", label_visibility="collapsed")

        homework_task = st.text_input("Homework Task Set for Next Session:", placeholder="e.g. Complete past paper questions on B1 Cells")
        class_notes = st.text_input("General Lesson Feedback Notes:", placeholder="Optional remarks...")

        submit_group_register = st.form_submit_button("💾 Save Class Register & Homework to Excel")
        if submit_group_register:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                if "Class Register" not in wb.sheetnames:
                    ws_reg = wb.create_sheet("Class Register")
                    ws_reg.append(["Date", "Day", "Time", "Student / Class Name", "Key Stage", "Status", "Topic Studied", "Homework", "Notes"])
                else:
                    ws_reg = wb["Class Register"]
                
                formatted_date = session_date.strftime("%Y-%m-%d")
                day_name = session_date.strftime("%A")
                
                for s_name, stat in student_statuses.items():
                    reason_note = student_absence_notes[s_name].strip()
                    full_note = f"Reason: {reason_note} | General: {class_notes}" if reason_note else class_notes
                    ws_reg.append([
                        formatted_date, 
                        day_name, 
                        time_choice, 
                        f"{selected_class_load} ({s_name})", 
                        ks_choice, 
                        stat, 
                        topic_studied, 
                        homework_task,
                        full_note
                    ])
                    
                wb.save(EXCEL_FILE)
                st.success(f"Successfully saved register for **{selected_class_load}**!")
                st.cache_data.clear()
                st.rerun()
            except Exception as ex:
                st.error(f"Error saving register: {ex}")

    st.divider()
    st.subheader("🔍 Search & Manage Historical Register Logs")
    
    try:
        xls_check = pd.ExcelFile(EXCEL_FILE)
        if "Class Register" in xls_check.sheet_names:
            df_register = pd.read_excel(EXCEL_FILE, sheet_name="Class Register", header=None, dtype=str, na_filter=False)
            df_register = df_register.loc[:, ~df_register.columns.astype(str).str.contains('^Unnamed')]
            
            expected_cols = ["Date", "Day", "Time", "Student / Class Name", "Key Stage", "Status", "Topic Studied", "Homework", "Notes"]
            if len(df_register.columns) >= len(expected_cols):
                df_register.columns = expected_cols[:len(df_register.columns)]

            search_type = st.radio("Search By:", ["All Records", "Specific Class", "Specific Student"], horizontal=True)
            
            filtered_reg = df_register.copy()
            if search_type == "Specific Class" and saved_classes:
                sel_search_class = st.selectbox("Select Class to Search:", list(saved_classes.keys()))
                filtered_reg = filtered_reg[filtered_reg["Student / Class Name"].str.contains(sel_search_class, case=False, na=False)]
            elif search_type == "Specific Student":
                sel_search_stud = st.selectbox("Select Student to Search:", all_students_list)
                filtered_reg = filtered_reg[filtered_reg["Student / Class Name"].str.contains(sel_search_stud, case=False, na=False)]

            if not filtered_reg.empty:
                st.dataframe(filtered_reg, use_container_width=True, hide_index=True)
                
                # Delete / Manage section
                st.markdown("### 🗑️ Delete or Modify Register Records")
                row_indices_to_delete = st.multiselect(
                    "Select row index/indices to delete from register log:",
                    options=filtered_reg.index.tolist(),
                    format_func=lambda x: f"Row {x}: {filtered_reg.loc[x].get('Student / Class Name', 'Record')} on {filtered_reg.loc[x].get('Date', 'N/A')}"
                )
                
                if st.button("🗑️ Delete Selected Register Rows"):
                    if row_indices_to_delete:
                        try:
                            wb = openpyxl.load_workbook(EXCEL_FILE)
                            ws_reg = wb["Class Register"]
                            sorted_indices = sorted(row_indices_to_delete, reverse=True)
                            for idx in sorted_indices:
                                ws_reg.delete_rows(int(idx) + 1)
                            wb.save(EXCEL_FILE)
                            st.success("Selected register logs successfully deleted!")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as ex:
                            st.error(f"Error deleting row: {ex}")
                    else:
                        st.warning("Please select at least one row to delete.")
            else:
                st.info("No register records found matching your search criteria.")

    except:
        st.info("No register records found yet.")

    st.divider()
    st.subheader("📊 Student Attendance Percentage Summary")
    
    try:
        xls_check = pd.ExcelFile(EXCEL_FILE)
        if "Class Register" in xls_check.sheet_names:
            df_reg_all = pd.read_excel(EXCEL_FILE, sheet_name="Class Register", header=None, dtype=str)
            df_reg_all = df_reg_all.loc[:, ~df_reg_all.columns.astype(str).str.contains('^Unnamed')]
            if len(df_reg_all.columns) >= 6:
                df_reg_all.columns = ["Date", "Day", "Time", "Student / Class Name", "Key Stage", "Status", "Topic Studied", "Homework", "Notes"][:len(df_reg_all.columns)]
                
                att_summary = []
                for s_n in all_students_list:
                    student_logs = df_reg_all[df_reg_all["Student / Class Name"].astype(str).str.contains(s_n, case=False, na=False)]
                    total_classes = len(student_logs)
                    if total_classes > 0:
                        present_count = len(student_logs[student_logs["Status"].isin(["Present", "Late"])])
                        att_pct = round((present_count / total_classes) * 100, 1)
                        flag = "🟢 Good" if att_pct >= 80 else "🔴 Low Attendance (<80%)"
                    else:
                        att_pct = 100.0
                        flag = "No Records"
                    att_summary.append({"Student Name": s_n, "Classes Logged": total_classes, "Attendance %": f"{att_pct}%", "Status": flag})
                st.dataframe(pd.DataFrame(att_summary), use_container_width=True, hide_index=True)
    except:
        st.info("Log your first registers above to view attendance percentages.")

# --- TAB 3: STUDENT PROFILES & PASTORAL ---
with tab3:
    st.header("🗂️ Student Pastoral Profiles, Contacts & Medical/Support Needs")
    st.info("Manage parent/guardian names, emails, phone numbers, secondary contacts and relationships, medical conditions, and additional learning support needs.")

    profiles_data = {}
    try:
        wb_p = openpyxl.load_workbook(EXCEL_FILE)
        if "Student Profiles" in wb_p.sheetnames:
            ws_p = wb_p["Student Profiles"]
            for row in ws_p.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    profiles_data[str(row[0]).strip().lower()] = {
                        "Student ID": row[1] or "",
                        "Parent Name": row[2] or "",
                        "Parent Email": row[3] or "",
                        "Primary Phone": row[4] or "",
                        "Secondary Name": row[5] or "",
                        "Secondary Relation": row[6] or "",
                        "Secondary Phone": row[7] or "",
                        "Allergies / Medical": row[8] or "",
                        "Additional Support Needs": row[9] or ""
                    }
    except:
        pass

    selected_prof_student = st.selectbox("Select Student to View / Update Profile:", all_students_list, key="prof_select_box")
    
    existing_p = profiles_data.get(selected_prof_student.strip().lower(), {
        "Parent Name": "", "Parent Email": "", "Primary Phone": "", 
        "Secondary Name": "", "Secondary Relation": "", "Secondary Phone": "", 
        "Allergies / Medical": "", "Additional Support Needs": ""
    })

    with st.form("student_profile_form"):
        st.subheader(f"➕ Add / Update Details for {selected_prof_student}")
        p_col1, p_col2 = st.columns(2)
        with p_col1:
            prof_parent_name = st.text_input("Parent / Guardian Full Name:", value=existing_p["Parent Name"], placeholder="e.g. Mr. & Mrs. Smith")
            prof_email = st.text_input("Parent / Carer Email:", value=existing_p["Parent Email"], placeholder="e.g. parent@example.com")
            prof_phone = st.text_input("Primary Contact Phone Number:", value=existing_p["Primary Phone"], placeholder="e.g. 07700 900000")
        with p_col2:
            prof_sec_name = st.text_input("Secondary Contact Name:", value=existing_p["Secondary Name"], placeholder="e.g. Jane Doe")
            prof_sec_rel = st.text_input("Secondary Contact Relationship:", value=existing_p["Secondary Relation"], placeholder="e.g. Grandmother / Uncle")
            prof_sec_phone = st.text_input("Secondary Contact Phone Number:", value=existing_p["Secondary Phone"], placeholder="e.g. 07700 900111")
            prof_allergies = st.text_input("Allergies / Medical Conditions:", value=existing_p["Allergies / Medical"], placeholder="e.g. Severe peanut allergy, Asthma")
            prof_support = st.text_input("Additional Support Needs (SEND / Access Arrangements):", value=existing_p["Additional Support Needs"], placeholder="e.g. Dyslexia, Extra time 25%")

        submit_profile = st.form_submit_button("💾 Save / Update Student Profile to Excel")
        if submit_profile:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                if "Student Profiles" not in wb.sheetnames:
                    ws_prof = wb.create_sheet("Student Profiles")
                    ws_prof.append(["StudentName", "StudentID", "ParentName", "ParentEmail", "PrimaryPhone", "SecondaryName", "SecondaryRelation", "SecondaryPhone", "AllergiesMedical", "SupportNeeds"])
                else:
                    ws_prof = wb["Student Profiles"]

                s_id_match = ""
                for _, s_r in df_st.iterrows():
                    if s_r["Student Name"].strip().lower() == selected_prof_student.strip().lower():
                        s_id_match = s_r["Student ID"]
                        break

                updated = False
                for r_idx, row in enumerate(ws_prof.iter_rows(min_row=2, values_only=True), start=2):
                    if row[0] and str(row[0]).strip().lower() == selected_prof_student.strip().lower():
                        ws_prof.cell(row=r_idx, column=2).value = s_id_match
                        ws_prof.cell(row=r_idx, column=3).value = prof_parent_name
                        ws_prof.cell(row=r_idx, column=4).value = prof_email
                        ws_prof.cell(row=r_idx, column=5).value = prof_phone
                        ws_prof.cell(row=r_idx, column=6).value = prof_sec_name
                        ws_prof.cell(row=r_idx, column=7).value = prof_sec_rel
                        ws_prof.cell(row=r_idx, column=8).value = prof_sec_phone
                        ws_prof.cell(row=r_idx, column=9).value = prof_allergies
                        ws_prof.cell(row=r_idx, column=10).value = prof_support
                        updated = True
                        break

                if not updated:
                    ws_prof.append([
                        selected_prof_student,
                        s_id_match,
                        prof_parent_name,
                        prof_email,
                        prof_phone,
                        prof_sec_name,
                        prof_sec_rel,
                        prof_sec_phone,
                        prof_allergies,
                        prof_support
                    ])

                wb.save(EXCEL_FILE)
                st.success(f"Profile updated successfully for **{selected_prof_student}**!")
                st.cache_data.clear()
                st.rerun()
            except Exception as ex:
                st.error(f"Error saving profile: {ex}")

    st.divider()
    st.subheader("🔍 View All Student Pastoral Profiles")

    if profiles_data:
        prof_list_display = []
        for name_k, data_v in profiles_data.items():
            prof_list_display.append({
                "Student Name": name_k.title(),
                "Student ID": data_v["Student ID"],
                "Parent Name": data_v["Parent Name"],
                "Parent Email": data_v["Parent Email"],
                "Primary Phone": data_v["Primary Phone"],
                "Secondary Contact": f"{data_v['Secondary Name']} ({data_v['Secondary Relation']}) - {data_v['Secondary Phone']}",
                "Allergies / Medical": data_v["Allergies / Medical"],
                "Support Needs": data_v["Additional Support Needs"]
            })
        st.dataframe(pd.DataFrame(prof_list_display), use_container_width=True, hide_index=True)
    else:
        st.info("No student pastoral profiles added yet. Use the form above to add contact details, medical info, and support needs.")

# --- TAB 4: LIVE PRACTICAL LOGGER ---
with tab4:
    st.header("Interactive Practical Logger")
    sel_student_name = st.selectbox(
        "Select Student:", df_st["Student Name"].tolist(), key="pract_sel"
    )

    rp_col = None
    for c in df_rp.columns:
        if sel_student_name.lower() in c.lower():
            rp_col = c
            break

    if rp_col:
        student_rp_df = df_rp[
            ["RP #", "Subject", "Practical Title", "Paper", rp_col]
        ].copy()
        student_rp_df.rename(columns={rp_col: "Status"}, inplace=True)

        st.write(f"Editing practical status for **{sel_student_name}**:")
        edited_df = st.data_editor(
            student_rp_df,
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=["Completed", "Outstanding"],
                    required=True,
                )
            },
            hide_index=True,
            use_container_width=True,
        )

        comp_count = (edited_df["Status"] == "Completed").sum()
        out_count = (edited_df["Status"] == "Outstanding").sum()
        st.success(
            f"Practicals Completed: {comp_count} | Practicals Outstanding: {out_count}"
        )

# --- TAB 5: STUDENT FEES MANAGEMENT ---
with tab5:
    st.header("💰 Student Fees & Tuition Payments Management")
    st.info("Track tuition fees, view balances by class group or individual student, and record payments.")

    class_group_options = ["All Classes"] + list(saved_classes.keys()) + ["GCSE Science Group A"]

    fees_data = []
    try:
        wb_fees = openpyxl.load_workbook(EXCEL_FILE)
        if "Student Fees" in wb_fees.sheetnames:
            ws_fees = wb_fees["Student Fees"]
            for row in ws_fees.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    fees_data.append({
                        "Class / Group": str(row[0]),
                        "Student Name": str(row[1]),
                        "Fee Due Date": str(row[2]),
                        "Amount Due (£)": float(row[3]) if row[3] is not None else 0.0,
                        "Payment Status": str(row[4]),
                        "Notes": str(row[5]) if len(row) > 5 and row[5] is not None else ""
                    })
    except Exception as ex:
        pass

    df_fees = pd.DataFrame(fees_data)

    with st.form("fee_payment_form"):
        st.subheader("➕ Record or Update Student Fee / Payment")
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            fee_class = st.selectbox("Class Group:", list(saved_classes.keys()) if saved_classes else ["GCSE Science Group A"])
            fee_student = st.selectbox("Student Name:", all_students_list, key="fee_stud_sel")
        with f_col2:
            fee_date = st.date_input("Due / Payment Date:", value=datetime.date.today())
            fee_amount = st.number_input("Amount (£):", min_value=0.0, value=35.0, step=5.0)
        with f_col3:
            fee_status = st.selectbox("Payment Status:", ["Paid Up-to-Date", "Outstanding Balance"])
            fee_notes = st.text_input("Payment Notes:", placeholder="e.g. Bank Transfer / Cash")

        submit_fee = st.form_submit_button("💾 Save Fee Record to Excel")
        if submit_fee:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                if "Student Fees" not in wb.sheetnames:
                    ws_f = wb.create_sheet("Student Fees")
                    ws_f.append(["ClassGroup", "StudentName", "DueDate", "AmountDue", "PaymentStatus", "Notes"])
                else:
                    ws_f = wb["Student Fees"]

                updated = False
                for r_idx, row in enumerate(ws_f.iter_rows(min_row=2, values_only=True), start=2):
                    if str(row[1]).strip().lower() == fee_student.strip().lower():
                        ws_f.cell(row=r_idx, column=1).value = fee_class
                        ws_f.cell(row=r_idx, column=3).value = fee_date.strftime("%Y-%m-%d")
                        ws_f.cell(row=r_idx, column=4).value = fee_amount
                        ws_f.cell(row=r_idx, column=5).value = fee_status
                        ws_f.cell(row=r_idx, column=6).value = fee_notes
                        updated = True
                        break

                if not updated:
                    ws_f.append([
                        fee_class,
                        fee_student,
                        fee_date.strftime("%Y-%m-%d"),
                        fee_amount,
                        fee_status,
                        fee_notes
                    ])

                wb.save(EXCEL_FILE)
                st.success(f"Fee record successfully saved for **{fee_student}**!")
                st.cache_data.clear()
                st.rerun()
            except Exception as ex:
                st.error(f"Error saving fee record: {ex}")

    st.divider()
    st.subheader("🔍 Search & View Fees (By Class or Individual Student)")

    f_search_col1, f_search_col2 = st.columns(2)
    with f_search_col1:
        filter_class_fee = st.selectbox("Filter by Class Group:", class_group_options, key="f_class_sel")
    with f_search_col2:
        filter_student_fee = st.selectbox("Filter by Individual Student:", ["All Students"] + all_students_list, key="f_stud_sel")

    if not df_fees.empty:
        disp_fees = df_fees.copy()
        if filter_class_fee != "All Classes":
            disp_fees = disp_fees[disp_fees["Class / Group"].str.strip().str.lower() == filter_class_fee.strip().lower()]
        if filter_student_fee != "All Students":
            disp_fees = disp_fees[disp_fees["Student Name"].str.strip().str.lower() == filter_student_fee.strip().lower()]

        if not disp_fees.empty:
            st.dataframe(disp_fees, use_container_width=True, hide_index=True)

            total_outstanding = disp_fees[disp_fees["Payment Status"].str.contains("Outstanding", case=False, na=False)]["AmountDue" if "AmountDue" in disp_fees.columns else "Amount Due (£)"].sum()
            st.metric("Total Outstanding Balance in View", f"£{total_outstanding:.2f}")
        else:
            st.info("No fee records found matching your selected filters.")
    else:
        st.info("No fee records saved in Excel yet. Use the form above to record your first student payment or balance.")

# --- TAB 6: COHORT OVERVIEW & FILTERS ---
with tab6:
    st.header("Cohort High-Level Performance (Mock Scale /70)")
    c1, c2, c3, c4 = st.columns(4)
    avg_cohort_mark_70 = df_st["Avg Mark (/35)"].mean() * 2 if "Avg Mark (/35)" in df_st.columns else 0
    meeting_target = (df_st["Predicted Grade"] >= df_st["Baseline Grade"]).sum() if "Predicted Grade" in df_st.columns and "Baseline Grade" in df_st.columns else 0
    total_pending_rp = df_st["Outstanding Practicals"].sum() if "Outstanding Practicals" in df_st.columns else 0
    critical_cases = (
        df_st["Practical Compliance Flag"]
        .astype(str)
        .str.contains("CRITICAL", na=False)
        .sum() if "Practical Compliance Flag" in df_st.columns else 0
    )

    c1.metric("Cohort Avg Mark (/70)", f"{avg_cohort_mark_70:.1f}")
    c2.metric("Meeting/Exceeding Target", f"{meeting_target} / {len(df_st)}")
    c3.metric("Total Pending Practicals", int(total_pending_rp))
    c4.metric("Critical Risk Students", int(critical_cases))

    st.divider()
    st.subheader("🔍 Advanced Cohort Intervention Filter")

    f_col1, f_col2 = st.columns(2)
    with f_col1:
        tier_filter = st.selectbox(
            "Filter by Entry Tier:", ["All Tiers", "Higher", "Foundation"], key="tier_f"
        )
    with f_col2:
        risk_filter = st.selectbox(
            "Filter by Compliance Risk:",
            ["All Risk Levels", "CRITICAL DEFICIT", "ACTION REQUIRED", "Secure"], key="risk_f"
        )

    filtered_cohort = df_st.copy()
    if tier_filter != "All Tiers" and "Tier (H/F)" in filtered_cohort.columns:
        filtered_cohort = filtered_cohort[
            filtered_cohort["Tier (H/F)"].astype(str).str.contains(tier_filter, case=False, na=False)
        ]
    if risk_filter != "All Risk Levels" and "Practical Compliance Flag" in filtered_cohort.columns:
        filtered_cohort = filtered_cohort[
            filtered_cohort["Practical Compliance Flag"].astype(str).str.contains(risk_filter, case=False, na=False)
        ]

    display_cols = [c for c in ["Student ID", "Student Name", "Tier (H/F)", "Baseline Grade", "Predicted Grade", "Projected Grade", "Practical Compliance Flag", "Focus / Action Priority"] if c in filtered_cohort.columns]
    st.dataframe(
        filtered_cohort[display_cols],
        use_container_width=True,
        hide_index=True,
    )

    st.divider()
    st.subheader("📦 Batch Cohort Report Export")
    st.info("Download a ZIP archive containing individual text report files for every student in the cohort.")

    if st.button("📥 Generate & Download Cohort ZIP Archive"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for _, student in df_st.iterrows():
                s_name = student["Student Name"]
                s_avg_70 = float(student.get("Avg Mark (/35)", 0)) * 2 if pd.notna(student.get("Avg Mark (/35)")) else 0.0
                s_pct = float(student.get("Avg %", 0)) * 100 if pd.notna(student.get("Avg %")) else 0.0
                _, _, cwg_val, pending_val = compute_student_metrics(student, all_topics, df_gb, df_rp)
                report_content = f"""====================================================
GCSE SCIENCE STUDENT PERFORMANCE REPORT
====================================================
Student Name: {s_name}
Student ID:   {student['Student ID']}
Programme:    {student.get('Study Programme', 'Science')} ({student.get('Exam Board', 'AQA')})
Tier:         {student.get('Tier (H/F)', 'N/A')}

PERFORMANCE SUMMARY:
- Current Working Grade (CWG): {safe_int(cwg_val)}
- Predicted Grade: {safe_int(get_first_valid(student, ['Predicted Grade'], 0))}
- Average Mock Score: {s_avg_70:.1f} / 70 ({s_pct:.1f}%)
- Pending Practicals: {safe_int(pending_val)}

ACTION & FOCUS PRIORITY:
{student.get('Focus / Action Priority', 'N/A')}
"""
                zip_file.writestr(f"{s_name}_Report.txt", report_content)

        zip_buffer.seek(0)
        st.download_button(
            label="📦 Click Here to Download ZIP File",
            data=zip_buffer,
            file_name="All_Student_Science_Reports.zip",
            mime="application/zip",
        )

# --- TAB 7: TOPIC HEATMAP & WEAKNESS ANALYSIS ---
with tab7:
    st.header("Class-Wide Topic Analysis (Scaled to /70 Mock Marks)")
    subject_choice = st.selectbox(
        "Select Subject:", ["All Subjects", "Biology", "Chemistry", "Physics"], key="subj_choice"
    )

    if subject_choice == "Biology":
        sub_cols = bio_cols
    elif subject_choice == "Chemistry":
        sub_cols = chem_cols
    elif subject_choice == "Physics":
        sub_cols = phys_cols
    else:
        sub_cols = all_topics

    topic_means_70 = df_st[sub_cols].mean().sort_values() * 2
    st.subheader(f"Cohort Topic Averages (/70) - {subject_choice}")
    st.bar_chart(topic_means_70)

    st.info(
        f" Lowest scoring module in {subject_choice}: **{topic_means_70.index[0]}** (Avg: {topic_means_70.iloc[0]:.1f}/70)"
    )

# --- TAB 8: GRADE GAP CALCULATOR ---
with tab8:
    st.header("Grade Boundary Distance Calculator (Mock Scale /70)")

    student_options = [
        f"{row['Student Name']} ({row['Student ID']})"
        for _, row in df_st.iterrows()
    ]
    sel_option = st.selectbox(
        "Select Student for Boundary Gap:", student_options, key="gap_stud_sel"
    )

    sel_student_id = sel_option.split("(")[-1].strip(")")
    s_row = df_st[df_st["Student ID"] == sel_student_id].iloc[0]

    calc_70, _, _, _ = compute_student_metrics(s_row, all_topics, df_gb, df_rp)
    mark_70 = float(calc_70)
    tier = str(s_row["Tier (H/F)"])
    curr_g, next_g, gap_70 = get_grade_gap(tier, mark_70)

    g1, g2, g3 = st.columns(3)
    g1.metric("Current Working Grade", curr_g)
    g2.metric(
        "Next Grade Milestone", next_g if next_g else "Max Grade Reached"
    )
    g3.metric(
        "Marks Needed for Next Grade", f"+{gap_70} marks" if next_g else "N/A"
    )

    st.progress(min(1.0, max(0.0, mark_70 / 70.0)))