import streamlit as st
from docx import Document
from io import BytesIO
import re
import time
import calendar
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import pandas as pd

# --- CORE FUNCTIONS ---
def fill_template(template, data):
    doc = Document(template)
    def replace_text_in_paragraph(paragraph, data):
        for key, value in data.items():
            if key in paragraph.text:
                for run in paragraph.runs:
                    if key in run.text:
                        run.text = run.text.replace(str(key), str(value))
    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph, data)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text_in_paragraph(paragraph, data)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def format_suffix_date(d):
    day = d.day
    suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return d.strftime(f"{day}{suffix} %B %Y")

def number_to_word_format(n_str):
    words = {
        '1': 'One', '2': 'Two', '3': 'Three', '4': 'Four', '5': 'Five', '6': 'Six', '7': 'Seven', '8': 'Eight', '9': 'Nine', '10': 'Ten',
        '11': 'Eleven', '12': 'Twelve', '13': 'Thirteen', '14': 'Fourteen', '15': 'Fifteen', '16': 'Sixteen', '17': 'Seventeen', '18': 'Eighteen', '19': 'Nineteen', '20': 'Twenty',
        '21': 'Twenty-One', '22': 'Twenty-Two', '23': 'Twenty-Three', '24': 'Twenty-Four', '25': 'Twenty-Five', '26': 'Twenty-Six', '27': 'Twenty-Seven', '28': 'Twenty-Eight', '29': 'Twenty-Nine', '30': 'Thirty',
        '31': 'Thirty-One', '32': 'Thirty-Two', '33': 'Thirty-Three', '34': 'Thirty-Four', '35': 'Thirty-Five', '36': 'Thirty-Six', '37': 'Thirty-Seven', '38': 'Thirty-Eight', '39': 'Thirty-Nine', '40': 'Forty',
        '41': 'Forty-One', '42': 'Forty-Two', '43': 'Forty-Three', '44': 'Forty-Four', '45': 'Forty-Five', '46': 'Forty-Six', '47': 'Forty-Seven', '48': 'Forty-Eight', '49': 'Forty-Nine', '50': 'Fifty',
        '51': 'Fifty-One', '52': 'Fifty-Two', '53': 'Fifty-Three', '54': 'Fifty-Four', '55': 'Fifty-Five', '56': 'Fifty-Six', '57': 'Fifty-Seven', '58': 'Fifty-Eight', '59': 'Fifty-Nine', '60': 'Sixty'
    }
    word = words.get(n_str, n_str) 
    return f"{word} ({n_str})"

if 'generated' not in st.session_state:
    st.session_state.generated = False

def clear_form():
    st.session_state.generated = False
    for key in list(st.session_state.keys()):
        if key.startswith("in_"):
            if "date" in key or "expiry" in key:
                st.session_state[key] = datetime.today().date()
            else:
                st.session_state[key] = ""

st.set_page_config(page_title="LoE Generator", layout="centered")

# --- HELP SECTION ---
with st.expander("📖 Template & App Help"):
    st.markdown("""
    - **Font Preservation:** Format tags in **Times New Roman** in Word.
    - **Calculated Dates:** `{{PREV_EXPIRY}}` is 1 day before Trans Start.
    - **New Tags:** `{{TRANS_MONTH}}` (e.g. September 2026).
    """)

st.markdown("<h1 style='font-size: 36px; font-weight: bold;'>Step 1: Upload Templates</h1>", unsafe_allow_html=True)
col_t1, col_t2 = st.columns(2)
with col_t1:
    perm_template = st.file_uploader("Permanent Employment Template", type="docx", key="ft_perm")
with col_t2:
    fix_template = st.file_uploader("Fixed Term Template", type="docx", key="ftc_perm")

if perm_template or fix_template:
    st.divider()
    st.subheader("Step 2: Category & Details")
    emp_type = st.radio("Generate letter for:", ["Permanent Employment", "Fixed Term"], horizontal=True)
    selected_template = perm_template if emp_type == "Permanent Employment" else fix_template
    
    if not selected_template:
        st.error(f"❌ Please upload the {emp_type} template first.")
    else:
        col_main_l, col_main_r = st.columns(2)
        with col_main_l:
            name = st.text_input("Full Name*", key="in_name")
            name_valid = bool(name and re.fullmatch(r"^[a-zA-Z\s]+$", name))
            if name and not name_valid: st.error("⚠️ Full Name: Letters and spaces only.")

            ic_number = st.text_input("IC Number (12 Digits)*", max_chars=12, key="in_ic")
            ic_valid = bool(ic_number.isdigit() and len(ic_number) == 12)
            if ic_number and not ic_valid: st.error("⚠️ Exactly 12 digits required.")

            job_title = st.text_input("Job Title*", key="in_job")
            job_valid = bool(job_title and re.fullmatch(r"^[a-zA-Z\s]+$", job_title))
            if job_title and not job_valid: st.error("⚠️ Job Title: Letters and spaces only.")

            project = st.text_input("Project Name*", key="in_project")
            project_valid = bool(project and re.fullmatch(r"^[a-zA-Z\s]+$", project))
            if project and not project_valid: st.error("⚠️ Project Name: Letters and spaces only.")

            emp_grade = st.text_input("Employee Grade*", key="in_grade")
            grade_valid = bool(emp_grade)
            if emp_grade == "": st.warning("⚠️ Employee Grade is required.")

        with col_main_r:
            s1, s2, s3 = st.columns([0.5, 2, 2])
            s1.markdown("<br>RM", unsafe_allow_html=True)
            salary_raw = s2.text_input("Monthly Salary*", key="in_salary")
            salary_clean = salary_raw.replace(",", "")
            salary_valid = bool(re.match(r"^\d+(\.\d{1,2})?$", salary_clean)) if salary_clean else False
            if salary_raw and not salary_valid: st.error("⚠️ Monthly Salary: Format 1,000.00")
            
            manager = st.text_input("Reporting Manager (Optional)", key="in_manager")
            manager_valid = True if not manager else bool(re.fullmatch(r"^[a-zA-Z\s]+$", manager))
            if manager and not manager_valid: st.error("⚠️ Reporting Manager: Letters and spaces only.")
            
            ceo_name = st.text_input("CEO Name*", value="SAMANTHA TAN", key="in_ceo")
            ceo_valid = bool(ceo_name and re.fullmatch(r"^[a-zA-Z\s]+$", ceo_name))
            if ceo_name and not ceo_valid: st.error("⚠️ CEO Name: Letters and spaces only.")
            
            start_date = st.date_input("Start Date*", min_value=datetime.today().date() - timedelta(days=30), key="in_start_date")

        fix_valid = True
        if emp_type == "Fixed Term":
            st.markdown("#### Additional Fixed Term Details")
            fl, fr = st.columns(2)
            with fl:
                t1, t2, t3 = st.columns([2, 1.5, 1.5])
                c_term = t1.text_input("Term*", key="in_fix_term")
                t2.markdown("<br>Months", unsafe_allow_html=True)
                term_valid = c_term.isdigit()
                if c_term and not term_valid: st.error("⚠️ Term: Digits only.")
                
                p1, p2, p3 = st.columns([2, 1.5, 1.5])
                probation = p1.text_input("Probation*", key="in_fix_prob")
                p2.markdown("<br>Months", unsafe_allow_html=True)
                prob_valid = probation.isdigit()
                if probation and not prob_valid: st.error("⚠️ Probation: Digits only.")
                
                n1, n2, n3 = st.columns([2, 1.5, 1.5])
                n_prob = n1.text_input("Notice (Probation)*", key="in_fix_nprob")
                n2.markdown("<br>Months", unsafe_allow_html=True)
                n_prob_valid = n_prob.isdigit()
                if n_prob and not n_prob_valid: st.error("⚠️ Notice (Probation): Digits only.")

                nc1, nc2, nc3 = st.columns([2, 1.5, 1.5])
                n_conf = nc1.text_input("Notice (Confirmed)*", key="in_fix_nconf")
                nc2.markdown("<br>Months", unsafe_allow_html=True)
                n_conf_valid = n_conf.isdigit()
                if n_conf and not n_conf_valid: st.error("⚠️ Notice (Confirmed): Digits only.")

            with fr:
                al1, al2, al3 = st.columns([2, 1.5, 1.5])
                al_val = al1.text_input("Annual Leave*", key="in_fix_al")
                al2.markdown("<br>Days", unsafe_allow_html=True)
                al_valid = al_val.isdigit()
                if al_val and not al_valid: st.error("⚠️ Annual Leave: Digits only.")
                
                o1, o2, o3 = st.columns([0.5, 2, 2.5])
                o1.markdown("<br>RM", unsafe_allow_html=True)
                outpatient = o2.text_input("Outpatient*", key="in_fix_out")
                out_valid = bool(re.match(r"^\d+(\.\d{1,2})?$", value = 0, outpatient.replace(",",""))) if outpatient else False
                if outpatient and not out_valid: st.error("⚠️ Outpatient: Digits/Decimals only.")

                tr1, tr2, tr3 = st.columns([0.5, 2, 2.5])
                tr1.markdown("<br>RM", unsafe_allow_html=True)
                travel = tr2.text_input("Traveling Allowance*", value = 0, key="in_fix_travel")
                travel_valid = bool(re.match(r"^\d+(\.\d{1,2})?$", travel.replace(",",""))) if travel else False
                if travel and not travel_valid: st.error("⚠️ Traveling Allowance: Digits/Decimals only.")

                kp1, kp2, kp3 = st.columns([0.5, 2, 2.5])
                kp1.markdown("<br>RM", unsafe_allow_html=True)
                kpi = kp2.text_input("Max KPI Payout*",  value = 0, key="in_fix_kpi")
                kpi_valid = bool(re.match(r"^\d+(\.\d{1,2})?$", kpi.replace(",",""))) if kpi else False
                if kpi and not kpi_valid: st.error("⚠️ KPI: Digits/Decimals only.")

                # Transitional Selectors
                months_list = list(calendar.month_name)[1:]
                years_list = [datetime.today().year + i for i in range(-1, 3)]
                m1, y1 = st.columns(2)
                t_sm_name = m1.selectbox("Trans. Start Month", months_list, key="in_fix_tsm")
                t_sy = y1.selectbox("Start Year", years_list, index=1, key="in_fix_tyear")
                m2, y2 = st.columns(2)
                t_em_name = m2.selectbox("Trans. End Month", months_list, key="in_fix_tem")
                t_ey = y2.selectbox("End Year", years_list, index=1, key="in_fix_teyear")
                
                f_ts = datetime(t_sy, months_list.index(t_sm_name) + 1, 1)
                f_te = datetime(t_ey, months_list.index(t_em_name) + 1, calendar.monthrange(t_ey, months_list.index(t_em_name) + 1)[1])

                date_order_valid = f_te >= f_ts
                if not date_order_valid:
                    st.error("⚠️ Error: Transitional End Month must be same or after Transitional Start Month.")

                f_as = f_te + timedelta(days=1)
                prev_expiry_calc = f_ts - timedelta(days=1)
                trans_month_label = f_ts.strftime("%B %Y")
                
                fix_valid = term_valid and probation and n_prob_valid and n_conf_valid and al_valid and out_valid and travel_valid and kpi_valid

        is_ready = bool(name_valid and ic_valid and job_valid and project_valid and grade_valid and salary_valid and ceo_valid and fix_valid)

        st.divider()
        if is_ready:
            if st.button(f"🚀 Generate {emp_type} Letter", type="primary"):
                st.session_state.generated = True
            if st.session_state.generated:
                with st.expander("✅ Review Data & Download", expanded=True):
                    def fmt_rm(v):
                        clean = str(v).replace(",","")
                        if not clean or float(clean) == 0: return "RM TBA"
                        return f"RM {float(clean):,.2f}"

                    data_map = {
                        "{{TODAY}}": format_suffix_date(datetime.today()),
                        "{{NAME}}": name.upper(), "{{IC_NUMBER}}": ic_number,
                        "{{JOB_TITLE}}": job_title, "{{PROJECT}}": project.upper(),
                        "{{GRADE}}": emp_grade.upper(), "{{START_DATE}}": format_suffix_date(start_date),
                        "{{SALARY}}": fmt_rm(salary_clean), "{{MANAGER}}": manager.upper() if manager else "YOUR SUPERIOR",
                        "{{CEO_NAME}}": ceo_name.upper()
                    }
                    if emp_type == "Fixed Term":
                        data_map.update({
                            "{{TERM}}": number_to_word_format(c_term), "{{EXPIRY_DATE}}": format_suffix_date(start_date + relativedelta(months=int(c_term))),
                            "{{PROBATION}}": probation, "{{NOTICE_PROB}}": n_prob, "{{NOTICE_CONF}}": n_conf,
                            "{{AL}}": al_val, "{{OUTPATIENT}}": fmt_rm(outpatient), "{{TRAVEL}}": fmt_rm(travel), "{{KPI}}": fmt_rm(kpi),
                            "{{TRANS_START}}": format_suffix_date(f_ts), "{{TRANS_END}}": format_suffix_date(f_te), "{{AGREE_START}}": format_suffix_date(f_as),
                            "{{PREV_EXPIRY}}": format_suffix_date(prev_expiry_calc),
                            "{{TRANS_MONTH}}": trans_month_label
                        })

                    type_code = "PERMANENT" if emp_type[0]=='P' else "FIXED-TERM"
                    fname = f"LOE_{type_code}_CONTRACT_{name.upper().replace(' ','_')}_{start_date.strftime('%Y_%m_%d')}.docx"
                    st.table(pd.DataFrame(list(data_map.items()), columns=["Tag", "Value"]))
                    st.download_button(f"📥 Download {fname}", data=fill_template(selected_template, data_map), file_name=fname)
                    st.button("🗑️ Reset Form", on_click=clear_form)
        else:
            st.warning("⚠️ Form incomplete or contains errors. Please fix the red labels above.")
