import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

st.set_page_config(
    page_title="AIM Rapid Review Clinic Discharge Summary",
    page_icon="🏥",
    layout="wide"
)

st.title("🏥 Acute Internal Medicine — Rapid Review Clinic Discharge Summary")
st.caption("Customized for AIM patients & Rapid Review Clinic | Fast, Structured & HIPAA/PDPA Offline Ready")

# Tabs for input
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "1. Demographics",
    "2. Problem List & Comorbidities",
    "3. Progress, POCUS & Resuscitation",
    "4. Lab & Imaging Trends",
    "5. Medications",
    "6. Review & Export"
])

with tab1:
    st.header("1. Patient Demographics")
    col1, col2 = st.columns(2)
    with col1:
        patient_name = st.text_input("Patient Full Name", value="", key="patient_name")
        rn_mrn = st.text_input("RN / MRN", value="", key="rn_mrn")
        ic_passport = st.text_input("NRIC / Passport No.", value="", key="ic_passport")
        age_sex = st.text_input("Age/Gender", value="", key="age_sex")
        contact_num = st.text_input("Patient/Caretaker Contact No.", value="", key="contact_num")
        caretaker = st.text_input("Caretaker Name & Relationship", value="", key="caretaker")
    with col2:
        adm_date = st.date_input("Admission Date", key="adm_date")
        dis_date = st.date_input("Date of Discharge", key="dis_date")
        referral_place = st.selectbox(
            "Referral Source", 
            ["ED", "Private Hospital/Private Clinic", "Klinik Kesihatan (KK)", "Ambulance", "Transfer from Other Wards HBM", "Transfer from Hospitals Cluster Perai", "Transfer from Hospitals of different Clusters", "Others"],
            key="referral_place"
        )
        referral_other = st.text_input("Specify Source", value="", key="referral_other")
        ward_bed = st.text_input("Ward / Bed No.", value="", key="ward_bed")
        attending_dr = st.text_input("Attending Doctor / Team", value="", key="attending_dr")
        subspec_team = st.text_input("Subspecialty Team(s) Involved", value="", key="subspec_team")
        address = st.text_area("Patient Address", value="", key="address")

with tab2:
    st.header("2. Problem List & Clinical Diagnoses")
    st.info("💡 Sub-options appear dynamically ONLY when a problem is checked.")
    st.subheader("Primary Diagnosis/Problems")
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        p_ccf = st.checkbox("Decompensated Heart Failure", key="chk_p_ccf")
        p_apo = st.checkbox("Acute Pulmonary Edema", key="chk_p_apo")
        p_copd = st.checkbox("Acute Exacerbation of COPD / Asthma", key="chk_p_copd")
        
        p_bronch = st.checkbox("Acute Exacerbation of Bronchiectasis", key="chk_p_bronch")
        bronch_type = ""
        bronch_detail = ""
        if p_bronch:
            bronch_type = st.selectbox("Bronchiectasis Status", ["Infective", "Non-Infective"], key="sel_bronch_type")
            bronch_detail = st.text_input("Organism/source", value="", key="txt_bronch_detail")

        p_ild = st.checkbox("Acute Exacerbation of Interstitial Lung Disease (ILD)", key="chk_p_ild")
        ild_type = ""
        ild_detail = ""
        if p_ild:
            ild_type = st.selectbox("ILD Status", ["Infective", "Non-Infective/Acute Flare"], key="sel_ild_type")
            ild_detail = st.text_input("Subtype (e.g. post infectious/IPF / CTD-ILD)", value="", key="txt_ild_detail")
        
        p_sepsis = st.checkbox("Severe Sepsis/septic shock", key="chk_p_sepsis")
        sepsis_source = ""
        if p_sepsis:
            sepsis_source = st.text_input("Sepsis Source", value="", key="txt_sepsis_source")

        p_poisoning = st.checkbox("Acute Poisoning", key="chk_p_poisoning")
        poisoning_detail = ""
        if p_poisoning:
            poisoning_detail = st.text_input("Specify Agent/Poisoning Details", value="", key="txt_poisoning_detail")

        p_withdrawal = st.checkbox("Drug / Substance Withdrawal", key="chk_p_withdrawal")
        withdrawal_type = ""
        if p_withdrawal:
            withdrawal_type = st.text_input("substance type/amphetamine/opiod/toxidromes", value="", key="sel_withdrawal_type")

    with col_p2:
        p_pneu = st.checkbox("Severe Pneumonia", key="chk_p_pneu")
        pneu_type = ""
        pneu_detail = ""
        if p_pneu:
            pneu_type = st.selectbox("Pneumonia Etiology", ["Not Specified", "Viral", "Bacterial", "fungi", "tuberculosis", "Atypical", "Others"], key="sel_pneu_type")
            pneu_detail = st.text_input("Organism / source", value="", key="txt_pneu_detail")

        p_snake = st.checkbox("Snake Bite with Envenomation", key="chk_p_snake")
        snake_detail = ""
        if p_snake:
            snake_detail = st.text_input("Snake type/Local or Systemic Envenomation", value="", key="txt_snake_detail")

        p_acs = st.checkbox("Acute Coronary Syndrome (ACS)", key="chk_p_acs")
        acs_type = ""
        acs_timi = ""
        if p_acs:
            acs_type = st.selectbox("ACS Type", ["UA", "NSTEMI", "STEMI"], key="sel_acs_type")
            acs_timi = st.text_input("TIMI Score", value="", key="txt_acs_timi")

        s_acs = st.checkbox("Secondary Acute Coronary Syndrome", key="chk_s_acs")
        
        p_pe = st.checkbox("Pulmonary Embolism (PE)", key="chk_p_pe")
        pe_category = ""
        pe_resp_mod = ""
        if p_pe:
            pe_category = st.selectbox(
                "2026 Acute PE Clinical Category (A–E)", 
                [
                    "Category A: Low Risk - Subclinical / Asymptomatic",
                    "Category B1: Symptomatic — Subsegmental (Single/multiple)",
                    "Category B2: Symptomatic — Non-subsegmental",
                    "Category C1: Low-Intermediate Risk - Normal Biomarkers & No RV Dysfunction",
                    "Category C2: Intermediate-High Risk - Elevated Cardiac Biomarkers",
                    "Category D1: Incipient Cardiopulmonary Failure — Transient hypotension",
                    "Category D2: Incipient Cardiopulmonary Failure — Normotensive shock",
                    "Category E1: Cardiopulmonary Failure — Recurrent/persistent hypotension",
                    "Category E2: Cardiopulmonary Failure — Refractory shock or cardiac arrest",
                ],
                key="sel_pe_cat_2026"
            )
            pe_resp_mod = st.selectbox(
                "Respiratory Modifier (+/- R)",
                [
                    "None / Not Applicable",
                    "Modifier R (Category C: SpO2 <90%, RR ≥30, need supplemental O2)",
                    "Modifier R (Category D: >6L NC or NRB mask)",
                    "Modifier R (Category E: Hypoxemic or ventilatory failure)"
                ],
                key="sel_pe_resp_mod_2026"
            )

        p_stroke_pri = st.checkbox("Acute Stroke", key="chk_p_stroke_pri")
        stroke_type = ""
        stroke_subtype = ""
        nihss_score = ""
        if p_stroke_pri:
            stroke_type = st.selectbox("Stroke Classification", ["Ischemic", "Hemorrhagic"], key="sel_stroke_type_pri")
            if stroke_type == "Ischemic":
                stroke_subtype = st.selectbox("Ischemic Subtype (Oxford / BAMFORD)", ["LACI (Lacunar Infarct)", "PACI (Partial Anterior Circulation Infarct)", "TACI (Total Anterior Circulation Infarct)", "POCI (Posterior Circulation Infarct)"], key="sel_stroke_ischemic_sub")
            else:
                stroke_subtype = st.text_input("Hemorrhagic Stroke Details", value="", key="txt_stroke_hem_detail")
            nihss_score = st.text_input("NIHSS Score / Severity at Presentation", value="", key="txt_nihss_score")

        p_seizure = st.checkbox("Seizures / Status Epilepticus", key="chk_p_seizure")
        seizure_type = ""
        seizure_detail = ""
        if p_seizure:
            seizure_type = st.selectbox("Seizure Type", ["Generalized Tonic-Clonic (GTC)", "Focal Seizure", "Status Epilepticus", "Absence / Myoclonic", "Provoked", "New Onset / Unclassified"], key="sel_seizure_type")
            seizure_detail = st.text_input("Seizure Details / Meds", value="", key="txt_seizure_detail")

        other_primary = st.text_area("Others (Primary Diagnoses)", value="", key="txt_other_primary")

    st.subheader("Secondary Problems")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        s_aki = st.checkbox("Acute Kidney Injury (AKI)", key="chk_s_aki")
        aki_cause = ""
        if s_aki:
            aki_cause = st.text_input("AKI Etiology", value="", key="txt_aki_cause")
            
        s_ards = st.checkbox("Acute Respiratory Distress Syndrome (ARDS)", key="chk_s_ards")
        ards_severity = ""
        if s_ards:
            ards_severity = st.selectbox("ARDS Severity", ["Mild", "Moderate", "Severe"], key="sel_ards_sev")

        s_rv_dys = st.checkbox("Right Ventricular Dysfunction / Failure", key="chk_s_rv_dys")

    with col_s2:
        s_arrhythmia = st.checkbox("Cardiac Arrhythmia", key="chk_s_arrhythmia")
        arrhythmia_detail = ""
        if s_arrhythmia:
            arrhythmia_detail = st.text_input("Arrhythmia type", value="", key="txt_arrhythmia_detail")
            
        s_stroke = st.checkbox("Stroke (Ischemic/Hemorrhagic)", key="chk_s_stroke")
        s_stroke_detail = ""
        if s_stroke:
            s_stroke_detail = st.text_input("Stroke details", value="", key="txt_s_stroke_detail")

        s_pe_sec = st.checkbox("Pulmonary Embolism", key="chk_s_pe_sec")
        s_pe_detail = ""
        if s_pe_sec:
            s_pe_detail = st.text_input("Pulmonary Embolism details", value="", key="txt_s_pe_detail")

    other_secondary = st.text_area("Additional Problems", value="", key="txt_other_secondary")

    st.subheader("Comorbidities")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        c_ckd = st.checkbox("CKD / ESRF", key="chk_c_ckd")
        ckd_stage = ""
        if c_ckd:
            ckd_stage = st.text_input("CKD Stage & Baseline Creatinine", value="", key="txt_ckd_stage")

        c_hf = st.checkbox("HFrEF / HFpEF / HFimpEF", key="chk_c_hf")
        hf_detail = ""
        if c_hf:
            hf_detail = st.text_input("LVEF/Etiology/Echo", value="", key="txt_hf_detail")

        c_cad = st.checkbox("Coronary Artery Disease", key="chk_c_cad")
        cad_detail = ""
        if c_cad:
            cad_detail = st.text_input("ASCVD Risk Score", value="", key="txt_cad_detail")

    with col_c2:
        c_dm = st.checkbox("Diabetes Mellitus", key="chk_c_dm")
        dm_detail = ""
        if c_dm:
            dm_detail = st.text_input("HbA1c % & Target Organ Damage", value="", key="txt_dm_detail")

        c_htn = st.checkbox("Hypertension", key="chk_c_htn")

    other_comorb = st.text_area("Other Comorbidities", value="", key="txt_other_comorb")

    st.subheader("Functional Status & Functional Capacity")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        frailty_selected = st.multiselect(
            "Frailty / Baseline Mobility Status (Select all that apply)",
            ["Independent", "Partially dependent", "Homebound", "Wheelchair-bound", "Fully bedridden"],
            key="mult_frailty"
        )

    with col_f2:
        nyha_mets = st.selectbox(
            "Functional Capacity",
            [
                "Not Specified",
                "NYHA Class I (No limitation / > 4 METS)",
                "NYHA Class II (Slight limitation / ~4 METS)",
                "NYHA Class III (Marked limitation / < 4 METS)",
                "NYHA Class IV (Severe limitation / Symptoms at rest / < 2 METS)"
            ],
            key="sel_nyha_mets"
        )
        func_notes = st.text_input("Remarks", value="", key="txt_func_notes")

with tab3:
    st.header("3. Ward Clinical Progress, POCUS & Resuscitation Directives")
    st.subheader("Main Presentation & Acute Interventions")
    presentation = st.text_area("Main Presentation & Associated Symptoms/Signs", value="", key="txt_presentation")
        
    col_int1, col_int2 = st.columns(2)
    with col_int1:
        int_o2 = st.text_input("Acute Oxygen Therapy Given", value="", key="txt_int_o2")
        int_fluid = st.text_input("Fluid Bolus", value="", key="txt_int_fluid")
        int_inotropes = st.text_input("Inotropes/Vasopressors (Acute)", value="", key="txt_int_inotropes")

    with col_int2:
        int_scan_by = st.text_input("Admission Bedside Scan Performed By", value="", key="txt_int_scan_by")
        int_scan_findings = st.text_area("Scan Findings", value="", key="txt_int_scan_findings")

    st.subheader("Ward Clinical Progress & Hemodynamic Monitoring")
    col_prog1, col_prog2 = st.columns(2)
    with col_prog1:
        max_o2_modality = st.selectbox("Highest Oxygenation", ["Room Air", "Nasal Cannula", "High Flow Mask (HFM)", "HFNC", "CPAP / BiPAP", "Invasive Mechanical Ventilation"], key="sel_max_o2")
        max_fio2 = st.text_input("Highest FiO2 %", value="", key="txt_max_fio2")
        o2_duration = st.text_input("Duration of High Flow Oxygenation (HFNC/NIV/IMV)", value="", key="txt_o2_duration")

    with col_prog2:
        o2_conc_dis = st.selectbox("Need Temporary Oxygen Concentrator at Discharge?", ["No", "Yes"], key="sel_o2_conc")
        o2_conc_duration = ""
        if o2_conc_dis == "Yes":
            o2_conc_duration = st.text_input("If Yes, Duration Needed", value="", key="txt_o2_conc_dur")
            
        inotropes_ward = st.multiselect(
            "Inotropes / Vasopressors Required in Ward",
            ["IV Noradrenaline", "IV Dobutamine", "IV Vasopressin", "IV Adrenaline", "IV Dopamine"],
            key="mult_inotropes_ward"
        )

    st.subheader("Inpatient Medications & Treatment")
    st.caption("⚡ Select inpatient medications administered. Configure Dose, Frequency, Date Start, and Date Complete to dynamically compile into a summary box to expedite discharges!")

    inmed_list = []
    
    # Grid layout for Inpatient Meds
    ci_1, ci_2 = st.columns(2)
    
    with ci_1:
        # 1. IV Unasyn
        chk_unasyn = st.checkbox("IV Unasyn (Ampicillin/Sulbactam)", key="chk_in_unasyn")
        if chk_unasyn:
            c_d, c_f = st.columns(2)
            u_dose = c_d.selectbox("Dose", ["1.5g", "3g", "4.5g"], key="sel_in_unasyn_d")
            u_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS", "QID", "4 hrly", "3 hr extended infusion", "4 hr extended infusion"], key="sel_in_unasyn_f")
            c_s, c_c = st.columns(2)
            u_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_unasyn_s")
            u_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_unasyn_c")
            u_start = u_s_dt.strftime("%d/%m/%Y") if u_s_dt else ""
            u_comp = u_c_dt.strftime("%d/%m/%Y") if u_c_dt else ""
            inmed_list.append(f"• IV Unasyn {u_dose} {u_freq} [Started: {u_start if u_start else 'N/A'} | Completed: {u_comp if u_comp else 'N/A'}]")

        # 2. IV Augmentin
        chk_aug = st.checkbox("IV Augmentin (Amoxycillin/Clavulanate"), key="chk_in_aug")
        if chk_aug:
            c_d, c_f = st.columns(2)
            aug_dose = c_d.selectbox("Dose", ["600mg", "1.2g"], key="sel_in_aug_d")
            aug_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS"], key="sel_in_aug_f")
            c_s, c_c = st.columns(2)
            aug_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_aug_s")
            aug_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_aug_c")
            aug_start = aug_s_dt.strftime("%d/%m/%Y") if aug_s_dt else ""
            aug_comp = aug_c_dt.strftime("%d/%m/%Y") if aug_c_dt else ""
            inmed_list.append(f"• IV Augmentin {aug_dose} {aug_freq} [Started: {aug_start if aug_start else 'N/A'} | Completed: {aug_comp if aug_comp else 'N/A'}]")

        # 3. IV Meropenem
        chk_mero = st.checkbox("IV Meropenem", key="chk_in_mero")
        if chk_mero:
            c_d, c_f = st.columns(2)
            mero_dose = c_d.selectbox("Dose", ["500 mg", "1g", "2g"], key="sel_in_mero_d")
            mero_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS"], key="sel_in_mero_f")
            c_s, c_c = st.columns(2)
            mero_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_mero_s")
            mero_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_mero_c")
            mero_start = mero_s_dt.strftime("%d/%m/%Y") if mero_s_dt else ""
            mero_comp = mero_c_dt.strftime("%d/%m/%Y") if mero_c_dt else ""
            inmed_list.append(f"• IV Meropenem {mero_dose} {mero_freq} [Started: {mero_start if mero_start else 'N/A'} | Completed: {mero_comp if mero_comp else 'N/A'}]")

        # 4. Tazosin
        chk_tazo = st.checkbox("IV Tazosin (Piperacillin/Tazobactam)", key="chk_in_tazo")
        if chk_tazo:
            c_d, c_f = st.columns(2)
            tazo_dose = c_d.selectbox("Dose", ["2.25g", "3.375g", "4.5g"], key="sel_in_tazo_d")
            tazo_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS", "QID", "3 hr extended infusion", "4 extended infusion"], key="sel_in_tazo_f")
            c_s, c_c = st.columns(2)
            tazo_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_tazo_s")
            tazo_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_tazo_c")
            tazo_start = tazo_s_dt.strftime("%d/%m/%Y") if tazo_s_dt else ""
            tazo_comp = tazo_c_dt.strftime("%d/%m/%Y") if tazo_c_dt else ""
            inmed_list.append(f"• Tazosin {tazo_dose} {tazo_freq} [Started: {tazo_start if tazo_start else 'N/A'} | Completed: {tazo_comp if tazo_comp else 'N/A'}]")

        # 5. Cefepime
        chk_cefep = st.checkbox("IV Cefepime", key="chk_in_cefep")
        if chk_cefep:
            c_d, c_f = st.columns(2)
            cefep_dose = c_d.selectbox("Dose", ["500mg", "1g", "2g"], key="sel_in_cefep_d")
            cefep_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS"], key="sel_in_cefep_f")
            c_s, c_c = st.columns(2)
            cefep_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_cefep_s")
            cefep_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_cefep_c")
            cefep_start = cefep_s_dt.strftime("%d/%m/%Y") if cefep_s_dt else ""
            cefep_comp = cefep_c_dt.strftime("%d/%m/%Y") if cefep_c_dt else ""
            inmed_list.append(f"• Cefepime {cefep_dose} {cefep_freq} [Started: {cefep_start if cefep_start else 'N/A'} | Completed: {cefep_comp if cefep_comp else 'N/A'}]")

        # 6. Ceftriaxone
        chk_ceftri = st.checkbox("IV Ceftriaxone", key="chk_in_ceftri")
        if chk_ceftri:
            c_d, c_f = st.columns(2)
            ceftri_dose = c_d.selectbox("Dose", ["1g", "2g"], key="sel_in_ceftri_d")
            ceftri_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS"], key="sel_in_ceftri_f")
            c_s, c_c = st.columns(2)
            ceftri_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_ceftri_s")
            ceftri_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_ceftri_c")
            ceftri_start = ceftri_s_dt.strftime("%d/%m/%Y") if ceftri_s_dt else ""
            ceftri_comp = ceftri_c_dt.strftime("%d/%m/%Y") if ceftri_c_dt else ""
            inmed_list.append(f"• Ceftriaxone {ceftri_dose} {ceftri_freq} [Started: {ceftri_start if ceftri_start else 'N/A'} | Completed: {ceftri_comp if ceftri_comp else 'N/A'}]")

        # 7. Fortum
        chk_fortum = st.checkbox("IV Fortum (Ceftazidime)", key="chk_in_fortum")
        if chk_fortum:
            c_d, c_f = st.columns(2)
            fortum_dose = c_d.selectbox("Dose", ["1g", "2g"], key="sel_in_fortum_d")
            fortum_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS", "QID"], key="sel_in_fortum_f")
            c_s, c_c = st.columns(2)
            fortum_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_fortum_s")
            fortum_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_fortum_c")
            fortum_start = fortum_s_dt.strftime("%d/%m/%Y") if fortum_s_dt else ""
            fortum_comp = fortum_c_dt.strftime("%d/%m/%Y") if fortum_c_dt else ""
            inmed_list.append(f"• Fortum {fortum_dose} {fortum_freq} [Started: {fortum_start if fortum_start else 'N/A'} | Completed: {fortum_comp if fortum_comp else 'N/A'}]")

        # 8. Cipro
        chk_cipro = st.checkbox("IV Ciprofloxacin)", key="chk_in_cipro")
        if chk_cipro:
            c_d, c_f = st.columns(2)
            cipro_dose = c_d.selectbox("Dose", ["200 mg", "600 mg", "400 mg"], key="sel_in_cipro_d")
            cipro_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS"], key="sel_in_cipro_f")
            c_s, c_c = st.columns(2)
            cipro_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_cipro_s")
            cipro_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_cipro_c")
            cipro_start = cipro_s_dt.strftime("%d/%m/%Y") if cipro_s_dt else ""
            cipro_comp = cipro_c_dt.strftime("%d/%m/%Y") if cipro_c_dt else ""
            inmed_list.append(f"• Cipro {cipro_dose} {cipro_freq} [Started: {cipro_start if cipro_start else 'N/A'} | Completed: {cipro_comp if cipro_comp else 'N/A'}]")

    with ci_2:
        # 9. IV Clindamycin
        chk_clinda = st.checkbox("IV Clindamycin", key="chk_in_clinda")
        if chk_clinda:
            c_d, c_f = st.columns(2)
            clinda_dose = c_d.selectbox("Dose", ["300 mg", "600 mg", "900 mg"], key="sel_in_clinda_d")
            clinda_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS", "QID"], key="sel_in_clinda_f")
            c_s, c_c = st.columns(2)
            clinda_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_clinda_s")
            clinda_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_clinda_c")
            clinda_start = clinda_s_dt.strftime("%d/%m/%Y") if clinda_s_dt else ""
            clinda_comp = clinda_c_dt.strftime("%d/%m/%Y") if clinda_c_dt else ""
            inmed_list.append(f"• IV Clindamycin {clinda_dose} {clinda_freq} [Started: {clinda_start if clinda_start else 'N/A'} | Completed: {clinda_comp if clinda_comp else 'N/A'}]")

        # 10. Acyclovir
        chk_acyclo = st.checkbox("IV Acyclovir", key="chk_in_acyclo")
        if chk_acyclo:
            c_d, c_f = st.columns(2)
            acyclo_dose = c_d.selectbox("Dose", ["250 mg", "500 mg", "10mg/kg/dose"], key="sel_in_acyclo_d")
            acyclo_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS"], key="sel_in_acyclo_f")
            c_s, c_c = st.columns(2)
            acyclo_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_acyclo_s")
            acyclo_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_acyclo_c")
            acyclo_start = acyclo_s_dt.strftime("%d/%m/%Y") if acyclo_s_dt else ""
            acyclo_comp = acyclo_c_dt.strftime("%d/%m/%Y") if acyclo_c_dt else ""
            inmed_list.append(f"• Acyclovir {acyclo_dose} {acyclo_freq} [Started: {acyclo_start if acyclo_start else 'N/A'} | Completed: {acyclo_comp if acyclo_comp else 'N/A'}]")

        # 11. Hydrocortisone
        chk_hcort = st.checkbox("IV Hydrocortisone", key="chk_in_hcort")
        if chk_hcort:
            c_d, c_f = st.columns(2)
            hcort_dose = c_d.selectbox("Dose", ["25 mg", "50 mg", "100 mg"], key="sel_in_hcort_d")
            hcort_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS", "QID"], key="sel_in_hcort_f")
            c_s, c_c = st.columns(2)
            hcort_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_hcort_s")
            hcort_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_hcort_c")
            hcort_start = hcort_s_dt.strftime("%d/%m/%Y") if hcort_s_dt else ""
            hcort_comp = hcort_c_dt.strftime("%d/%m/%Y") if hcort_c_dt else ""
            inmed_list.append(f"• Hydrocortisone {hcort_dose} {hcort_freq} [Started: {hcort_start if hcort_start else 'N/A'} | Completed: {hcort_comp if hcort_comp else 'N/A'}]")

        # 12. Dexa
        chk_dexa_in = st.checkbox("IV Dexamethasone", key="chk_in_dexa")
        if chk_dexa_in:
            c_d, c_f = st.columns(2)
            dexain_dose = c_d.selectbox("Dose", ["2 mg", "4 mg", "6 mg", "8 mg"], key="sel_in_dexain_d")
            dexain_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS"], key="sel_in_dexain_f")
            c_s, c_c = st.columns(2)
            dexain_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_dexain_s")
            dexain_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_dexain_c")
            dexain_start = dexain_s_dt.strftime("%d/%m/%Y") if dexain_s_dt else ""
            dexain_comp = dexain_c_dt.strftime("%d/%m/%Y") if dexain_c_dt else ""
            inmed_list.append(f"• Dexa {dexain_dose} {dexain_freq} [Started: {dexain_start if dexain_start else 'N/A'} | Completed: {dexain_comp if dexain_comp else 'N/A'}]")

        # 13. IV Pantoprazole
        chk_panto_in = st.checkbox("IV Pantoprazole", key="chk_in_panto")
        if chk_panto_in:
            c_d, c_f = st.columns(2)
            panto_in_dose = c_d.selectbox("Dose", ["40 mg", "80 mg", "8mg/hr"], key="sel_in_panto_d")
            panto_in_freq = c_f.selectbox("Freq", ["OD", "BD", "Continuous Infusion"], key="sel_in_panto_f")
            c_s, c_c = st.columns(2)
            panto_in_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_panto_s")
            panto_in_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_panto_c")
            panto_in_start = panto_in_s_dt.strftime("%d/%m/%Y") if panto_in_s_dt else ""
            panto_in_comp = panto_in_c_dt.strftime("%d/%m/%Y") if panto_in_c_dt else ""
            inmed_list.append(f"• IV Pantoprazole {panto_in_dose} {panto_in_freq} [Started: {panto_in_start if panto_in_start else 'N/A'} | Completed: {panto_in_comp if panto_in_comp else 'N/A'}]")

        # 14. IV Azithromycin
        chk_azith = st.checkbox("IV Azithromycin", key="chk_in_azith")
        if chk_azith:
            c_d, c_f = st.columns(2)
            azith_dose = c_d.selectbox("Dose", ["500 mg"], key="sel_in_azith_d")
            azith_freq = c_f.selectbox("Freq", ["OD"], key="sel_in_azith_dur")
            c_s, c_c = st.columns(2)
            azith_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_azith_s")
            azith_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_azith_c")
            azith_start = azith_s_dt.strftime("%d/%m/%Y") if azith_s_dt else ""
            azith_comp = azith_c_dt.strftime("%d/%m/%Y") if azith_c_dt else ""
            inmed_list.append(f"• IV Azithromycin {azith_dose} OD ({azith_dur}) [Started: {azith_start if azith_start else 'N/A'} | Completed: {azith_comp if azith_comp else 'N/A'}]")

        # 15. Oral Tamiflu
        chk_tami = st.checkbox("Oral Tamiflu", key="chk_in_tami")
        if chk_tami:
            c_d, c_f = st.columns(2)
            tami_dose = c_d.selectbox("Dose", ["75 mg", "150 mg"], key="sel_in_tami_d")
            tami_freq = c_f.selectbox("Freq", ["OD", "EOD", "BD"], key="sel_in_tami_f")
            c_s, c_c = st.columns(2)
            tami_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_tami_s")
            tami_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_tami_c")
            tami_start = tami_s_dt.strftime("%d/%m/%Y") if tami_s_dt else ""
            tami_comp = tami_c_dt.strftime("%d/%m/%Y") if tami_c_dt else ""
            inmed_list.append(f"• Oral Tamiflu {tami_dose} {tami_freq} [Started: {tami_start if tami_start else 'N/A'} | Completed: {tami_comp if tami_comp else 'N/A'}]")

    st.markdown("**Nebulization**")
    cn1, cn2 = st.columns(2)
    with cn1:
        # Combivent
        chk_comb = st.checkbox("Neb Combivent (Ipratopium bromide 0.5 mg/albuterol 2.5 mg per 3 mL vial) 4 hrly", key="chk_in_comb")
        if chk_comb:
            c_s, c_c = st.columns(2)
            comb_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_comb_s")
            comb_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_comb_c")
            comb_start = comb_s_dt.strftime("%d/%m/%Y") if comb_s_dt else ""
            comb_comp = comb_c_dt.strftime("%d/%m/%Y") if comb_c_dt else ""
            inmed_list.append(f"• Neb Combivent 3 ml 4 hrly [Started: {comb_start if comb_start else 'N/A'} | Completed: {comb_comp if comb_comp else 'N/A'}]")

        # Salbutamol
        chk_salb = st.checkbox("Neb Salbutamol", key="chk_in_salb")
        if chk_salb:
            salb_freq = st.selectbox("Freq", ["2 hrly", "4 hrly"], key="sel_in_salb_f")
            c_s, c_c = st.columns(2)
            salb_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_salb_s")
            salb_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_salb_c")
            salb_start = salb_s_dt.strftime("%d/%m/%Y") if salb_s_dt else ""
            salb_comp = salb_c_dt.strftime("%d/%m/%Y") if salb_c_dt else ""
            inmed_list.append(f"• Neb Salbutamol 2.5mg {salb_freq} [Started: {salb_start if salb_start else 'N/A'} | Completed: {salb_comp if salb_comp else 'N/A'}]")

    with cn2:
        # Pulmicort
        chk_pulm = st.checkbox("Neb Pulmicort", key="chk_in_pulm")
        if chk_pulm:
            pulm_dose = st.selectbox("Dose", ["1 mg", "2 mg", "3 mg"], key="sel_in_pulm_d")
            c_s, c_c = st.columns(2)
            pulm_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_pulm_s")
            pulm_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_pulm_c")
            pulm_start = pulm_s_dt.strftime("%d/%m/%Y") if pulm_s_dt else ""
            pulm_comp = pulm_c_dt.strftime("%d/%m/%Y") if pulm_c_dt else ""
            inmed_list.append(f"• Neb Pulmicort {pulm_dose} BD [Started: {pulm_start if pulm_start else 'N/A'} | Completed: {pulm_comp if pulm_comp else 'N/A'}]")

        # Neb GTN
        chk_gtn = st.checkbox("Neb GTN (5 mg)", key="chk_in_gtn")
        if chk_gtn:
            c_s, c_c = st.columns(2)
            gtn_s_dt = c_s.date_input("Date Start", value=None, format="DD/MM/YYYY", key="txt_in_gtn_s")
            gtn_c_dt = c_c.date_input("Date Complete", value=None, format="DD/MM/YYYY", key="txt_in_gtn_c")
            gtn_start = gtn_s_dt.strftime("%d/%m/%Y") if gtn_s_dt else ""
            gtn_comp = gtn_c_dt.strftime("%d/%m/%Y") if gtn_c_dt else ""
            inmed_list.append(f"• Neb GTN 5 mg [Started: {gtn_start if gtn_start else 'N/A'} | Completed: {gtn_comp if gtn_comp else 'N/A'}]")

    inmed_custom = st.text_input("Other Inpatient Medications / Infusions", value="", key="txt_inmed_custom")
    if inmed_custom:
        inmed_list.append(f"• Others: {inmed_custom}")

    # Summary box for Inpatient Meds Regimen
    inpatient_regimen_text = "\n".join(inmed_list) if inmed_list else "Nil inpatient acute medications selected."
    st.text_area("Inpatient Medications & Treatment Regimen Summary Box (Copyable)", value=inpatient_regimen_text, height=150)

    
    
        # 4. Tazosin
        chk_tazo = st.checkbox("IV Tazosin (Piperacillin/Tazobactam)", key="chk_in_tazo")
        if chk_tazo:
            c_d, c_f = st.columns(2)
            tazo_dose = c_d.selectbox("Dose", ["2.25g", "3.375g", "4.5g"], key="sel_in_tazo_d")
            tazo_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS", "QID", "QID 3 hr infusion", "TDS 4 hr infusion"], key="sel_in_tazo_f")
            c_s, c_c = st.columns(2)
            tazo_start = c_s.text_input("Date Start", value="", key="txt_in_tazo_s")
            tazo_comp = c_c.text_input("Date Complete", value="", key="txt_in_tazo_c")
            inmed_list.append(f"• Tazosin {tazo_dose} {tazo_freq} [Started: {tazo_start if tazo_start else 'N/A'} | Completed: {tazo_comp if tazo_comp else 'N/A'}]")

        # 5. Cefepime
        chk_cefep = st.checkbox("IV Cefepime", key="chk_in_cefep")
        if chk_cefep:
            c_d, c_f = st.columns(2)
            cefep_dose = c_d.selectbox("Dose", ["1g", "2g"], key="sel_in_cefep_d")
            cefep_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS"], key="sel_in_cefep_f")
            c_s, c_c = st.columns(2)
            cefep_start = c_s.text_input("Date Start", value="", key="txt_in_cefep_s")
            cefep_comp = c_c.text_input("Date Complete", value="", key="txt_in_cefep_c")
            inmed_list.append(f"• Cefepime {cefep_dose} {cefep_freq} [Started: {cefep_start if cefep_start else 'N/A'} | Completed: {cefep_comp if cefep_comp else 'N/A'}]")

        # 6. Ceftriaxone
        chk_ceftri = st.checkbox("IV Ceftriaxone", key="chk_in_ceftri")
        if chk_ceftri:
            c_d, c_f = st.columns(2)
            ceftri_dose = c_d.selectbox("Dose", ["1g", "2g"], key="sel_in_ceftri_d")
            ceftri_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS"], key="sel_in_ceftri_f")
            c_s, c_c = st.columns(2)
            ceftri_start = c_s.text_input("Date Start", value="", key="txt_in_ceftri_s")
            ceftri_comp = c_c.text_input("Date Complete", value="", key="txt_in_ceftri_c")
            inmed_list.append(f"• Ceftriaxone {ceftri_dose} {ceftri_freq} [Started: {ceftri_start if ceftri_start else 'N/A'} | Completed: {ceftri_comp if ceftri_comp else 'N/A'}]")

        # 7. Fortum
        chk_fortum = st.checkbox("IV Fortum (Ceftazidime)", key="chk_in_fortum")
        if chk_fortum:
            c_d, c_f = st.columns(2)
            fortum_dose = c_d.selectbox("Dose", ["1g", "2g"], key="sel_in_fortum_d")
            fortum_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS", "QID"], key="sel_in_fortum_f")
            c_s, c_c = st.columns(2)
            fortum_start = c_s.text_input("Date Start", value="", key="txt_in_fortum_s")
            fortum_comp = c_c.text_input("Date Complete", value="", key="txt_in_fortum_c")
            inmed_list.append(f"• Fortum {fortum_dose} {fortum_freq} [Started: {fortum_start if fortum_start else 'N/A'} | Completed: {fortum_comp if fortum_comp else 'N/A'}]")

        # 8. Cipro
        chk_cipro = st.checkbox("Cipro (IV Ciprofloxacin)", key="chk_in_cipro")
        if chk_cipro:
            c_d, c_f = st.columns(2)
            cipro_dose = c_d.selectbox("Dose", ["400 mg"], key="sel_in_cipro_d")
            cipro_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS", "QID"], key="sel_in_cipro_f")
            c_s, c_c = st.columns(2)
            cipro_start = c_s.text_input("Date Start", value="", key="txt_in_cipro_s")
            cipro_comp = c_c.text_input("Date Complete", value="", key="txt_in_cipro_c")
            inmed_list.append(f"• Cipro {cipro_dose} {cipro_freq} [Started: {cipro_start if cipro_start else 'N/A'} | Completed: {cipro_comp if cipro_comp else 'N/A'}]")

    with ci_2:
        # 9. IV Clindamycin
        chk_clinda = st.checkbox("IV Clindamycin", key="chk_in_clinda")
        if chk_clinda:
            c_d, c_f = st.columns(2)
            clinda_dose = c_d.selectbox("Dose", ["300 mg", "600 mg", "900 mg"], key="sel_in_clinda_d")
            clinda_freq = c_f.selectbox("Freq", ["OD", "BD", "TDS", "QID"], key="sel_in_clinda_f")
            c_s, c_c = st.columns(2)
            clinda_start = c_s.text_input("Date Start", value="", key="txt_in_clinda_s")
            clinda_comp = c_c.text_input("Date Complete", value="", key="txt_in_clinda_c")
            inmed_list.append(f"• IV Clindamycin {clinda_dose} {clinda_freq} [Started: {clinda_start if clinda_start else 'N/A'} | Completed: {clinda_comp if clinda_comp else 'N/
    st.markdown("**Nebulization:**")
    cn1, cn2 = st.columns(2)
    with cn1:
        # Combivent
        chk_comb = st.checkbox("Neb Combivent (Ipratopium bromide 0.5 mg/albuterol 2.5 mg per 3 mL vial) 4 hrly", key="chk_in_comb")
        if chk_comb:
            c_s, c_c = st.columns(2)
            comb_start = c_s.text_input("Date Start", value="", key="txt_in_comb_s")
            comb_comp = c_c.text_input("Date Complete", value="", key="txt_in_comb_c")
            inmed_list.append(f"• Neb Combivent 3ml 4 hrly [Started: {comb_start if comb_start else 'N/A'} | Completed: {comb_comp if comb_comp else 'N/A'}]")

        # Salbutamol
        chk_salb = st.checkbox("Neb Salbutamol", key="chk_in_salb")
        if chk_salb:
            salb_freq = st.selectbox("Freq", ["2 hrly", "4 hrly"], key="sel_in_salb_f")
            c_s, c_c = st.columns(2)
            salb_start = c_s.text_input("Date Start", value="", key="txt_in_salb_s")
            salb_comp = c_c.text_input("Date Complete", value="", key="txt_in_salb_c")
            inmed_list.append(f"• Neb Salbutamol 2.5mg {salb_freq} [Started: {salb_start if salb_start else 'N/A'} | Completed: {salb_comp if salb_comp else 'N/A'}]")

    with cn2:
        # Pulmicort
        chk_pulm = st.checkbox("Neb Pulmicort", key="chk_in_pulm")
        if chk_pulm:
            pulm_dose = st.selectbox("Dose", ["1 mg", "2 mg", "3 mg"], key="sel_in_pulm_d")
            c_s, c_c = st.columns(2)
            pulm_start = c_s.text_input("Date Start", value="", key="txt_in_pulm_s")
            pulm_comp = c_c.text_input("Date Complete", value="", key="txt_in_pulm_c")
            inmed_list.append(f"• Neb Pulmicort {pulm_dose} BD [Started: {pulm_start if pulm_start else 'N/A'} | Completed: {pulm_comp if pulm_comp else 'N/A'}]")

        # Neb GTN
        chk_gtn = st.checkbox("Neb GTN (5 mg)", key="chk_in_gtn")
        if chk_gtn:
            c_s, c_c = st.columns(2)
            gtn_start = c_s.text_input("Date Start", value="", key="txt_in_gtn_s")
            gtn_comp = c_c.text_input("Date Complete", value="", key="txt_in_gtn_c")
            inmed_list.append(f"• Neb GTN 5 mg QID/PRN [Started: {gtn_start if gtn_start else 'N/A'} | Completed: {gtn_comp if gtn_comp else 'N/A'}]")

    inmed_custom = st.text_input("Other Inpatient Medications / Infusions", value="", key="txt_inmed_custom")
    if inmed_custom:
        inmed_list.append(f"• Others: {inmed_custom}")

    # Summary box for Inpatient Meds Regimen
    inpatient_regimen_text = "\n".join(inmed_list) if inmed_list else "Nil inpatient acute medications selected."
    st.text_area("Inpatient Medications & Treatment Regimen Summary Box (Copyable)", value=inpatient_regimen_text, height=150)

    st.subheader("POCUS & Diagnostic Investigations")
    pocus_lung = st.text_input("POCUS Lung", value="", key="txt_pocus_lung")
    pocus_heart = st.text_input("POCUS Cardiac", value="", key="txt_pocus_heart")
    pocus_plr = st.text_input("Passive Leg Raising (PLR) for Fluid Responsiveness", value="", key="txt_pocus_plr")
    pocus_vessels = st.text_input("POCUS Vessels", value="", key="txt_pocus_vessels")
    pocus_others = st.text_input("POCUS Others", value="", key="txt_pocus_others")
    pocus_interp = st.text_area("Clinical Interpretation", value="", key="txt_pocus_interp")
    
    ecg_findings = st.text_area("ECG", value="", key="txt_ecg_findings")
    cxr_findings = st.text_area("CXR", value="", key="txt_cxr_findings")
    other_ix = st.text_area("Other Imaging", value="", key="other_imaging")

    st.subheader("Inpatient Procedures")
    st.caption("💡 Check a procedure to enter specific details & lab/scan results.")
    col_proc1, col_proc2 = st.columns(2)
    with col_proc1:
        proc_pleural = st.checkbox("Pleural Aspiration / Drainage", key="chk_proc_pl")
        proc_pleural_date = ""
        proc_pleural_detail = ""
        if proc_pleural:
            cp1, cp2 = st.columns([1, 2.5])
            with cp1:
                proc_pleural_date = st.text_input("Date Performed", value="", key="txt_proc_pl_date")
            with cp2:
                proc_pleural_detail = st.text_input("Details & Results (e.g. volume, appearance, biochemistry, culture)", value="", key="txt_proc_pl_detail")

        proc_ascites = st.checkbox("Peritoneal Aspiration / Drainage", key="chk_proc_asc")
        proc_ascites_date = ""
        proc_ascites_detail = ""
        if proc_ascites:
            cp1, cp2 = st.columns([1, 2.5])
            with cp1:
                proc_ascites_date = st.text_input("Date Performed", value="", key="txt_proc_asc_date")
            with cp2:
                proc_ascites_detail = st.text_input("Paracentesis Details & Results (e.g. SAAG, cell count, culture)", value="", key="txt_proc_asc_detail")

        proc_lp = st.checkbox("Lumbar Puncture", key="chk_proc_lp")
        proc_lp_date = ""
        proc_lp_detail = ""
        if proc_lp:
            cp1, cp2 = st.columns([1, 2.5])
            with cp1:
                proc_lp_date = st.text_input("Date Performed", value="", key="txt_proc_lp_date")
            with cp2:
                proc_lp_detail = st.text_input("Lumbar Puncture Details & Results (e.g. opening pressure, CSF biochemistry, Gram stain)", value="", key="txt_proc_lp_detail")

    with col_proc2:
        proc_cvl = st.checkbox("Central Venous Line (CVL)", key="chk_proc_cvl")
        proc_cvl_detail = ""
        if proc_cvl:
            proc_cvl_detail = st.text_input("CVL Site & date of Insertion and Removed", value="", key="txt_proc_cvl_detail")

        proc_hd_cath = st.checkbox("Acute Dialysis Catheter", key="chk_proc_hd")
        proc_hd_detail = ""
        if proc_hd_cath:
            proc_hd_detail = st.text_input("Insertion Site, Date of Insertion and Removed & Line Status", value="", key="txt_proc_hd_detail")

        proc_others_chk = st.checkbox("Other Procedures", key="chk_proc_others")
        proc_others_detail = ""
        if proc_others_chk:
            proc_others_detail = st.text_input("Other Procedures Types and Results", value="", key="txt_proc_others_detail")

    dialysis_needed = st.selectbox("Acute Dialysis Required During Admission?", ["No", "Yes"], key="sel_dialysis")
    dialysis_mode = ""
    dialysis_date = ""
    renal_recovery = ""
    if dialysis_needed == "Yes":
        dialysis_mode = st.selectbox("Mode of Dialysis Support", ["Intermittent HD", "SLED / SCUF", "CVVH", "CVVHD", "CVVHDF", "PD"], key="sel_dialysis_mode")
        col_dial1, col_dial2 = st.columns(2)
        with col_dial1:
            dialysis_date = st.text_input("Indication & Date Commenced & Stopped", value="", key="txt_dialysis_date")
        with col_dial2:
            renal_recovery = st.selectbox("Any Renal Recovery at Discharge", ["No", "Yes"], key="txt_renal_rec")    

    st.subheader("Resuscitation Status & Ceiling of Care Directive")
    Advance_directive = st.selectbox("Any Advanced Directive/Patient Expressed Wish", ["No", "Yes"], key="advance_directive")
    resus_code_selected = st.multiselect("Resuscitation Code (Select all that apply)", ["Full Code", "Do-Not-Intubate (DNI)/CPR with Maximum Ceiling of Care", "Full Comfort/Palliative Care"], key="mult_resus_code")
    st.markdown("**Maximum Ceiling of Care Interventions Allowed (Select all applicable):**")
    col_rc1, col_rc2 = st.columns(2)
    with col_rc1:
        resus_hfnc = st.checkbox("HFNC (High Flow Nasal Cannula)", key="chk_resus_hfnc")
        resus_niv = st.checkbox("NIV (Non-Invasive Ventilation)", key="chk_resus_niv")
        resus_hfmo = st.checkbox("HFM / HFMO (High Flow Mask Oxygen)", key="chk_resus_hfmo")
    with col_rc2:
        resus_hd = st.checkbox("Acute Dialysis", key="chk_resus_hd")
        resus_inotropes = st.checkbox("Vasopressors/Inotropes", key="chk_resus_inotropes")
        resus_ab = st.checkbox("Ab (Antibiotics)", key="chk_resus_ab")
    
    resus_notes = st.text_area("Remarks/Family Discussion/Future Direction of Care", value="", key="txt_resus_notes")

with tab4:
    st.header("4. Investigations Trend Table (5-Column Format)")
    st.caption("5 Serial Date Timestamps: Admission → Random Ward Day 1 → Random Ward Day 2 → Random Ward Day 3 → Predischarge Date")
        
    col_d1, col_d2, col_d3, col_d4, col_d5 = st.columns(5)
    with col_d1:
        date_adm_str = st.text_input("Admission Date", value="14/10/2026", key="lbl_adm_date")
    with col_d2:
        date_mid1_str = st.text_input("Ward Day 1", value="15/10/2026", key="lbl_mid1_date")
    with col_d3:
        date_mid2_str = st.text_input("Ward Day 2", value="16/10/2026", key="lbl_mid2_date")
    with col_d4:
        date_mid3_str = st.text_input("Ward Day 3", value="17/10/2026", key="lbl_mid3_date")
    with col_d5:
        date_dis_str = st.text_input("Predischarge Date", value="19/10/2026", key="lbl_dis_date")

    st.markdown("---")
        
    # Headers
    h1, h2, h3, h4, h5, h6 = st.columns([2.5, 3, 3, 3, 3, 3])
    h1.markdown("**Parameter**")
    h2.markdown(f"**Admission ({date_adm_str})**")
    h3.markdown(f"**Ward Day 1 ({date_mid1_str})**")
    h4.markdown(f"**Ward Day 2 ({date_mid2_str})**")
    h5.markdown(f"**Ward Day 3 ({date_mid3_str})**")
    h6.markdown(f"**Predischarge ({date_dis_str})**")

    # Lab Dict
    lab_params = [
        ("Hb", "g/dL"),
        ("Hct", "%"),
        ("Wcc", "x10^9/L"),
        ("Lymp", "% / count"),
        ("Plt", "x10^9/L"),
        ("Urea", "mmol/L"),
        ("Creat", "umol/L"),
        ("Na", "mmol/L"),
        ("K", "mmol/L"),
        ("LDH", "U/L"),
        ("ALT", "U/L"),
        ("Coag", "PT/aPTT/INR"),
        ("CRP", "mg/L"),
        ("HbA1c", "%"),
        ("Trop I", "ug/L"),
        ("ProBNP", "pg/mL"),
        ("Others", "ABG/Notes")
    ]

    lab_inputs = {}
    for param, unit in lab_params:
        c1, c2, c3, c4, c5, c6 = st.columns([2.5, 3, 3, 3, 3, 3])
        c1.markdown(f"**{param}** <small>({unit})</small>", unsafe_allow_html=True)
        val_adm = c2.text_input(f"{param} Adm", value="", key=f"inp_{param}_adm", label_visibility="collapsed")
        val_mid1 = c3.text_input(f"{param} Mid1", value="", key=f"inp_{param}_mid1", label_visibility="collapsed")
        val_mid2 = c4.text_input(f"{param} Mid2", value="", key=f"inp_{param}_mid2", label_visibility="collapsed")
        val_mid3 = c5.text_input(f"{param} Mid3", value="", key=f"inp_{param}_mid3", label_visibility="collapsed")
        val_dis = c6.text_input(f"{param} Dis", value="", key=f"inp_{param}_dis", label_visibility="collapsed")
        lab_inputs[param] = (val_adm, val_mid1, val_mid2, val_mid3, val_dis)

    st.markdown("---")

with tab5:
    st.header("5. Discharge Medication Regimen & Steroid Regimen")
    st.caption("Check ONLY the medications prescribed at discharge.")
    st.subheader("Select Prescribed Discharge Medications")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        # 1. Pantoprazole
        m_panto = st.checkbox("Pantoprazole", value=True, key="chk_m_panto")
        val_panto = ""
        if m_panto:
            val_panto = "40mg OD (Fixed)"
                
        # 2. Calcium Carbonate (CaCO3)
        m_caco3 = st.checkbox("Calcium Carbonate (CaCO3)", value=False, key="chk_m_caco3")
        val_caco3 = ""
        if m_caco3:
            c1, c2 = st.columns(2)
            with c1:
                caco_dose = st.selectbox("Dose", ["500 mg", "1 g", "1.5 g", "2 g"], key="sel_caco_dose")
            with c2:
                caco_freq = st.selectbox("Freq", ["BD", "OD", "TDS", "PRN"], key="sel_caco_freq")
            val_caco3 = f"{caco_dose} {caco_freq}"
                
        # 3. Rocaltriol / Calcitriol
        m_rocal = st.checkbox("Rocaltriol / Calcitriol", value=False, key="chk_m_rocal")
        val_rocal = ""
        if m_rocal:
            c1, c2 = st.columns(2)
            with c1:
                rocal_dose = st.selectbox("Dose", ["0.25 mcg", "0.5 mcg", "0.75 mcg", "1 mcg"], key="sel_rocal_dose")
            with c2:
                rocal_freq = st.selectbox("Freq", ["OD", "BD", "PRN"], key="sel_rocal_freq")
            val_rocal = f"{rocal_dose} {rocal_freq}"
        
        # 4. SGLT2i
        m_sglt2 = st.checkbox("SGLT2i (Dapagliflozin / Empagliflozin)", value=True, key="chk_m_sglt2")
        val_sglt2 = ""
        if m_sglt2:
            sglt2_choice = st.selectbox("Regimen", ["Dapagliflozin 10mg OD (Fixed)", "Empagliflozin 10mg OD", "Empagliflozin 25mg OD"], key="sel_sglt2_choice")
            val_sglt2 = sglt2_choice

        # 5. Spironolactone
        m_spiro = st.checkbox("Spironolactone", value=False, key="chk_m_spiro")
        val_spiro = ""
        if m_spiro:
            c1, c2 = st.columns(2)
            with c1:
                spiro_dose = st.selectbox("Dose", ["25 mg", "12.5 mg", "37.5 mg", "50 mg"], key="sel_spiro_dose")
            with c2:
                spiro_freq = st.selectbox("Freq", ["OD", "BD", "PRN"], key="sel_spiro_freq")
            val_spiro = f"{spiro_dose} {spiro_freq}"

        # 6. Frusemide
        m_fruse = st.checkbox("Frusemide", value=True, key="chk_m_fruse")
        val_fruse = ""
        if m_fruse:
            c1, c2 = st.columns(2)
            with c1:
                fruse_dose = st.selectbox("Dose", ["40 mg", "20 mg", "60 mg", "80 mg"], key="sel_fruse_dose")
            with c2:
                fruse_freq = st.selectbox("Freq", ["OD", "BD", "TDS", "QID", "PRN"], key="sel_fruse_freq")
            val_fruse = f"{fruse_dose} {fruse_freq}"

        # 7. Amlodipine
        m_amlo = st.checkbox("Amlodipine", value=False, key="chk_m_amlo")
        val_amlo = ""
        if m_amlo:
            c1, c2 = st.columns(2)
            with c1:
                amlo_dose = st.selectbox("Dose", ["5 mg", "10 mg"], key="sel_amlo_dose")
            with c2:
                amlo_freq = st.selectbox("Freq", ["OD", "BD", "PRN"], key="sel_amlo_freq")
            val_amlo = f"{amlo_dose} {amlo_freq}"

        # 8. Aspirin / Glyprin
        m_glyprin = st.checkbox("Aspirin / Glyprin", value=True, key="chk_m_glyprin")
        val_glyprin = ""
        if m_glyprin:
            val_glyprin = "1 tablet (100mg) OD (Fixed)"

        # 9. Atorvastatin
        m_atorva = st.checkbox("Atorvastatin", value=True, key="chk_m_atorva")
        val_atorva = ""
        if m_atorva:
            c1, c2 = st.columns(2)
            with c1:
                atorva_dose = st.selectbox("Dose", ["40 mg", "80 mg"], key="sel_atorva_dose")
            with c2:
                atorva_freq = st.selectbox("Freq", ["ON (Night)", "OD", "PRN"], key="sel_atorva_freq")
            val_atorva = f"{atorva_dose} {atorva_freq}"
        
        # 10. Bisoprolol
        m_biso = st.checkbox("Bisoprolol", value=False, key="chk_m_biso")
        val_biso = ""
        if m_biso:
            c1, c2 = st.columns(2)
            with c1:
                biso_dose = st.selectbox("Dose", ["1.25 mg", "2.5 mg", "3.75 mg", "5 mg"], key="sel_biso_dose")
            with c2:
                biso_freq = st.selectbox("Freq", ["OD", "BD", "PRN"], key="sel_biso_freq")
            val_biso = f"{biso_dose} {biso_freq}"

    with col_m2:
        # 11. Perindopril
        m_perin = st.checkbox("Perindopril", value=False, key="chk_m_perin")
        val_perin = ""
        if m_perin:
            c1, c2 = st.columns(2)
            with c1:
                perin_dose = st.selectbox("Dose", ["2 mg", "4 mg", "6 mg", "8 mg"], key="sel_perin_dose")
            with c2:
                perin_freq = st.selectbox("Freq", ["OD", "BD", "PRN"], key="sel_perin_freq")
            val_perin = f"{perin_dose} {perin_freq}"

        # 12. ARNI (Sacubitril / Valsartan)
        m_arni = st.checkbox("ARNI (Sacubitril / Valsartan)", value=False, key="chk_m_arni")
        val_arni = ""
        if m_arni:
            c1, c2 = st.columns(2)
            with c1:
                arni_dose = st.selectbox("Dose", ["25 mg", "50 mg", "100 mg", "150 mg", "200 mg"], key="sel_arni_dose")
            with c2:
                arni_freq = st.selectbox("Freq", ["BD", "OD", "PRN"], key="sel_arni_freq")
            val_arni = f"{arni_dose} {arni_freq}"

        # 13. Metformin (IR & XR)
        m_metformin = st.checkbox("Metformin", value=False, key="chk_m_metformin")
        val_metformin = ""
        if m_metformin:
            c1, c2, c3 = st.columns([1.5, 1, 1])
            with c1:
                met_form = st.selectbox("Form", ["Metformin IR", "Metformin XR"], key="sel_met_form")
            with c2:
                met_dose = st.selectbox("Dose", ["500 mg", "1 g"], key="sel_met_dose")
            with c3:
                met_freq = st.selectbox("Freq", ["BD", "OD", "TDS", "PRN"], key="sel_met_freq")
            val_metformin = f"{met_form} {met_dose} {met_freq}"

        # 14. Gliclazide (Standard & MR)
        m_gliclazide = st.checkbox("Gliclazide", value=False, key="chk_m_gliclazide")
        val_gliclazide = ""
        if m_gliclazide:
            c1, c2, c3 = st.columns([1.5, 1, 1])
            with c1:
                glic_form = st.selectbox("Form", ["Gliclazide Standard", "Gliclazide MR"], key="sel_glic_form")
            with c2:
                if glic_form == "Gliclazide MR":
                    glic_dose = st.selectbox("Dose", ["30 mg", "60 mg", "90 mg", "120 mg"], key="sel_glic_dose_mr")
                else:
                    glic_dose = st.selectbox("Dose", ["40 mg", "80 mg", "120 mg", "160 mg"], key="sel_glic_dose_std")
            with c3:
                glic_freq = st.selectbox("Freq", ["OD", "BD", "PRN"], key="sel_glic_freq")
            val_gliclazide = f"{glic_form} {glic_dose} {glic_freq}"
                
        # 15. Subcutaneous Insulin Regimens
        m_insulin = st.checkbox("SC Insulin Regimens", value=False, key="chk_m_insulin")
        val_insulin = ""
        if m_insulin:
            c1, c2 = st.columns([1.5, 1])
            with c1:
                ins_type = st.selectbox("Regimen Type", ["SC Mixtard (BD)", "SC Insulatard BIDS (BD)", "SC Insulatard Basal-Bolus (ON)", "SC Actrapid (OD/BD/TDS)", "SC Glargine / Lantus (ON)"], key="sel_ins_type")
            with c2:
                ins_dose = st.text_input("Dose Details (e.g. 12u BD / 10u ON)", value="12u BD", key="txt_ins_dose")
            val_insulin = f"{ins_type}: {ins_dose}"

        # 16. Prazosin
        m_prazosin = st.checkbox("Prazosin", value=False, key="chk_m_prazosin")
        val_prazosin = ""
        if m_prazosin:
            c1, c2 = st.columns(2)
            with c1:
                prazo_dose = st.selectbox("Dose", ["0.5 mg", "1 mg", "1.5 mg", "2 mg"], key="sel_prazo_dose")
            with c2:
                prazo_freq = st.selectbox("Freq", ["TDS", "QID", "BD", "PRN"], key="sel_prazo_freq")
            val_prazosin = f"{prazo_dose} {prazo_freq}"

        # 17. Zincofer
        m_zincofer = st.checkbox("Zincofer", value=False, key="chk_m_zincofer")
        val_zincofer = ""
        if m_zincofer:
            c1, c2 = st.columns(2)
            with c1:
                zinc_dose = st.selectbox("Dose", ["1 tablet", "2 tablets"], key="sel_zinc_dose")
            with c2:
                zinc_freq = st.selectbox("Freq", ["OD", "BD", "PRN"], key="sel_zinc_freq")
            val_zincofer = f"{zinc_dose} {zinc_freq}"

        # 18. Vit B12 / Folate / Multivitamins
        m_vit = st.checkbox("Vit B12 / Folate / Multivitamins", value=False, key="chk_m_vit")
        val_vit = ""
        if m_vit:
            c1, c2 = st.columns(2)
            with c1:
                vit_dose = st.selectbox("Dose", ["1 tablet", "2 tablets"], key="sel_vit_dose")
            with c2:
                vit_freq = st.selectbox("Freq", ["OD", "BD", "PRN"], key="sel_vit_freq")
            val_vit = f"{vit_dose} {vit_freq}"

        # 19. Syp Morphine
        m_morphine = st.checkbox("Syp Morphine", value=False, key="chk_m_morphine")
        val_morphine = ""
        if m_morphine:
            c1, c2 = st.columns(2)
            with c1:
                morph_dose = st.text_input("Dosage (e.g. 5mg / 2.5ml)", value="5mg", key="txt_morph_dose")
            with c2:
                morph_freq = st.selectbox("Freq", ["4-hourly", "PRN", "OD", "BD", "TDS", "QID"], key="sel_morph_freq")
            val_morphine = f"{morph_dose} {morph_freq}"

        # 20. Paracetamol (PCM)
        m_pcm = st.checkbox("Paracetamol (PCM)", value=False, key="chk_m_pcm")
        val_pcm = ""
        if m_pcm:
            c1, c2 = st.columns(2)
            with c1:
                pcm_dose = st.selectbox("Dose", ["1 g (2 tablets)", "500 mg (1 tablet)"], key="sel_pcm_dose")
            with c2:
                pcm_freq = st.selectbox("Freq", ["QID PRN", "TDS PRN", "QID", "TDS", "PRN"], key="sel_pcm_freq")
            val_pcm = f"{pcm_dose} {pcm_freq}"

    med_custom = st.text_area("Additional Discharge Medications", value="", key="txt_med_custom")

    # Oral Steroid Regimen Section
    st.subheader("Oral Prednisolone & Dexamethasone Regimen")
    
    # Prednisolone
    chk_p_reg = st.checkbox("Prescribe Oral Prednisolone Regimen", key="chk_p_reg")
    pred_steps_list = []
    pred_regimen_summary = ""
    if chk_p_reg:
        p_ind = st.text_input("Indication", value="COPD / Asthma / ILD Flare", key="txt_p_ind")
        st.markdown("**Configure Step-by-Step Tapering Steps:**")
        
        c_s1, c_s2, c_s3 = st.columns(3)
        with c_s1:
            p_d1 = st.selectbox("Step 1 Dose", ["30 mg", "40 mg", "50 mg", "60 mg", "20 mg", "15 mg", "10 mg", "5 mg"], key="sel_p_d1")
            p_f1 = st.selectbox("Step 1 Freq", ["OD", "BD", "OFF"], key="sel_p_f1")
            p_dur1 = st.selectbox("Step 1 Duration", ["5 days", "3 days", "7 days", "10 days", "14 days"], key="sel_p_dur1")
            if p_f1 != "OFF":
                pred_steps_list.append(f"  - Step 1: Prednisolone {p_d1} {p_f1} for {p_dur1}")

        with c_s2:
            p_d2 = st.selectbox("Step 2 Dose", ["20 mg", "30 mg", "15 mg", "10 mg", "5 mg"], key="sel_p_d2")
            p_f2 = st.selectbox("Step 2 Freq", ["OD", "BD", "OFF"], key="sel_p_f2")
            p_dur2 = st.selectbox("Step 2 Duration", ["5 days", "3 days", "7 days", "10 days"], key="sel_p_dur2")
            if p_f2 != "OFF":
                pred_steps_list.append(f"  - Step 2: Then {p_d2} {p_f2} for {p_dur2}")

        with c_s3:
            p_d3 = st.selectbox("Step 3 Dose", ["10 mg", "15 mg", "5 mg"], key="sel_p_d3")
            p_f3 = st.selectbox("Step 3 Freq", ["OD", "BD", "OFF"], key="sel_p_f3")
            p_dur3 = st.selectbox("Step 3 Duration", ["5 days", "3 days", "7 days"], key="sel_p_dur3")
            if p_f3 != "OFF":
                pred_steps_list.append(f"  - Step 3: Then {p_d3} {p_f3} for {p_dur3}")

        c_s4, c_s5 = st.columns(2)
        with c_s4:
            p_d4 = st.selectbox("Step 4 Dose", ["5 mg", "10 mg"], key="sel_p_d4")
            p_f4 = st.selectbox("Step 4 Freq", ["OFF", "OD"], key="sel_p_f4")
            p_dur4 = st.selectbox("Step 4 Duration", ["5 days", "3 days", "Then OFF"], key="sel_p_dur4")
            if p_f4 != "OFF":
                pred_steps_list.append(f"  - Step 4: Then {p_d4} {p_f4} for {p_dur4}")

        pred_regimen_summary = f"Oral Prednisolone Regimen (Indication: {p_ind}):\n" + "\n".join(pred_steps_list)
        st.text_area("Oral Prednisolone Regimen Summary Box", value=pred_regimen_summary, height=130)

    # Dexamethasone
    chk_d_reg = st.checkbox("Prescribe Oral Dexamethasone Regimen", key="chk_d_reg")
    dexa_steps_list = []
    dexa_regimen_summary = ""
    if chk_d_reg:
        d_ind = st.text_input("Indication (Dexa)", value="Brain Edema / Flare", key="txt_d_ind")
        st.markdown("**Configure Step-by-Step Steps:**")
        
        cd_1, cd_2 = st.columns(2)
        with cd_1:
            d_d1 = st.selectbox("Step 1 Dose (Dexa)", ["6 mg", "8 mg", "4 mg", "2 mg"], key="sel_d_d1")
            d_f1 = st.selectbox("Step 1 Freq (Dexa)", ["OD", "BD", "OFF"], key="sel_d_f1")
            d_dur1 = st.selectbox("Step 1 Duration (Dexa)", ["5 days", "3 days", "7 days", "10 days"], key="sel_d_dur1")
            if d_f1 != "OFF":
                dexa_steps_list.append(f"  - Step 1: Dexamethasone {d_d1} {d_f1} for {d_dur1}")

        with cd_2:
            d_d2 = st.selectbox("Step 2 Dose (Dexa)", ["4 mg", "2 mg", "1 mg"], key="sel_d_d2")
            d_f2 = st.selectbox("Step 2 Freq (Dexa)", ["OD", "OFF"], key="sel_d_f2")
            d_dur2 = st.selectbox("Step 2 Duration (Dexa)", ["5 days", "3 days", "Then OFF"], key="sel_d_dur2")
            if d_f2 != "OFF":
                dexa_steps_list.append(f"  - Step 2: Then {d_d2} {d_f2} for {d_dur2}")

        dexa_regimen_summary = f"Oral Dexamethasone Regimen (Indication: {d_ind}):\n" + "\n".join(dexa_steps_list)
        st.text_area("Oral Dexamethasone Regimen Summary Box", value=dexa_regimen_summary, height=110)

with tab6:
    st.header("6. Final Review & Export Options")
    st.subheader("AIM Rapid Review Clinic & Follow-Up Plan")
    aim_clinic_plan = st.text_area("AIM Clinic Appointment & Issues to Review/Plan", value="Review symptoms, fluid balance, repeat renal profile & electrolyte monitoring.", key="txt_aim_clinic_plan")
    other_clinic_plan = st.text_area("Other Subspecialty Clinic Follow-Ups (Nephro/Cardio/Resp)", value="", key="txt_other_clinic_plan")
    red_flags = st.text_area("Red Flag Warning Signs Provided to Patient", value="Severe dyspnea, chest discomfort, acute syncope, marked decrease in urine output, or sudden swelling.", key="txt_red_flags")
    mo_name1 = st.text_input("Medical Officer Name", value="", key="txt_mo1_name")
    mo_name2 = st.text_input("Medical Officer Incharged", value="", key="txt_mo2_name")

    # Build Dynamic Plain Text Summary
    summary_text = (
        "================================================================================\n"
        " ACUTE INTERNAL MEDICINE — RAPID REVIEW CLINIC DISCHARGE SUMMARY\n"
        "================================================================================\n\n"
        "[ 1. PATIENT DEMOGRAPHICS & METADATA ]\n"
        f"Patient Name: {patient_name}    RN/MRN: {rn_mrn}\n"
        f"IC/Passport: {ic_passport}    Age/Sex: {age_sex}\n"
        f"Admission Date: {adm_date}    Discharge Date: {dis_date}\n"
        f"Referral Source: {referral_place} {f'({referral_other})' if referral_other else ''}\n"
        f"Ward/Bed: {ward_bed}    Attending Team: {attending_dr}\n"
        f"Subspecialty Team: {subspec_team}\n"
        f"Contact No: {contact_num}    Caretaker: {caretaker}\n"
        f"Address: {address}\n\n"
        "[ 2. PROBLEM LIST & DIAGNOSES ]\nPrimary Diagnoses:\n"
    )
    
    has_primary = False
    if p_ccf: summary_text += "• Decompensated Heart Failure\n"; has_primary = True
    if p_apo: summary_text += "• Acute Pulmonary Edema\n"; has_primary = True
    if p_copd: summary_text += "• Acute Exacerbation of COPD / Asthma\n"; has_primary = True
    if p_bronch: summary_text += f"• Bronchiectasis ({bronch_type}): {bronch_detail}\n"; has_primary = True
    if p_ild: summary_text += f"• Interstitial Lung Disease ({ild_type}): {ild_detail}\n"; has_primary = True
    if p_sepsis: summary_text += f"• Severe Sepsis secondary to: {sepsis_source}\n"; has_primary = True
    if p_poisoning: summary_text += f"• Poisoning: {poisoning_detail}\n"; has_primary = True
    if p_pneu: summary_text += f"• Severe Pneumonia ({pneu_type}): {pneu_detail}\n"; has_primary = True
    if p_snake: summary_text += f"• Snake Bite with Envenomation: {snake_detail}\n"; has_primary = True
    if p_acs: summary_text += f"• Acute Coronary Syndrome ({acs_type}) - TIMI: {acs_timi}\n"; has_primary = True
    if s_acs: summary_text += "• Secondary Acute Coronary Syndrome\n"; has_primary = True
    if p_pe: summary_text += f"• Pulmonary Embolism ({pe_category}){f' [Resp Mod: {pe_resp_mod}]' if pe_resp_mod != 'None / Not Applicable' else ''}\n"; has_primary = True
    if p_stroke_pri:
        summary_text += f"• Acute Stroke ({stroke_type}): {stroke_subtype}\n"
        if nihss_score: summary_text += f"  - NIHSS Score: {nihss_score}\n"
        has_primary = True
    if p_withdrawal: summary_text += f"• Drug Withdrawal ({withdrawal_type})\n"; has_primary = True
    if p_seizure: summary_text += f"• Seizures ({seizure_type}){f': {seizure_detail}' if seizure_detail else ''}\n"; has_primary = True
    if other_primary: summary_text += f"• Others: {other_primary}\n"; has_primary = True
    if not has_primary: summary_text += "• Nil acute primary diagnoses selected.\n"

    summary_text += "\nSecondary Problems:\n"
    has_sec = False
    if s_aki: summary_text += f"• AKI due to: {aki_cause if aki_cause else 'Diagnosed'}\n"; has_sec = True
    if s_ards: summary_text += f"• ARDS ({ards_severity if ards_severity else 'Diagnosed'})\n"; has_sec = True
    if s_rv_dys: summary_text += "• Right Ventricular Dysfunction / Failure\n"; has_sec = True
    if s_arrhythmia: summary_text += f"• Cardiac Arrhythmia: {arrhythmia_detail if arrhythmia_detail else 'Diagnosed'}\n"; has_sec = True
    if s_stroke: summary_text += f"• Stroke (Secondary){f': {s_stroke_detail}' if s_stroke_detail else ''}\n"; has_sec = True
    if s_pe_sec: summary_text += f"• Pulmonary Embolism (Secondary){f': {s_pe_detail}' if s_pe_detail else ''}\n"; has_sec = True
    if other_secondary: summary_text += f"• Others: {other_secondary}\n"; has_sec = True
    if not has_sec: summary_text += "• Nil acute secondary problems.\n"

    summary_text += "\nComorbidities:\n"
    if c_ckd: summary_text += f"• CKD Stage / Baseline: {ckd_stage if ckd_stage else 'Diagnosed'}\n"
    if c_hf: summary_text += f"• Heart Failure: {hf_detail if hf_detail else 'HFrEF/HFpEF/HFimpEF'}\n"
    if c_cad: summary_text += f"• CAD / ASCVD: {cad_detail if cad_detail else 'Diagnosed'}\n"
    if c_dm: summary_text += f"• Diabetes Mellitus: {dm_detail if dm_detail else 'Diagnosed'}\n"
    if c_htn: summary_text += "• Hypertension\n"
    if other_comorb: summary_text += f"• Other Comorbidities: {other_comorb}\n"

    summary_text += "\nFunctional Status & Capacity:\n"
    if frailty_selected: summary_text += f"• Mobility / Frailty: {', '.join(frailty_selected)}\n"
    if nyha_mets != "Not Specified": summary_text += f"• NYHA / METS Functional Capacity: {nyha_mets}\n"
    if func_notes: summary_text += f"• Functional / ADL Notes: {func_notes}\n"

    inotropes_str = ", ".join(inotropes_ward) if inotropes_ward else "None"

    summary_text += (
        "\n[ 3. CLINICAL PROGRESS & RESUSCITATION DIRECTIVE ]\n"
        f"Main Presentation: {presentation}\n"
        f"Acute Interventions: O2: {int_o2} | Fluid: {int_fluid} | Inotropes: {int_inotropes}\n"
        f"Initial Bedside Scan by {int_scan_by}: {int_scan_findings}\n\n"
        "Ward Progress:\n"
        f"• Highest Oxygen Modality: {max_o2_modality} (FiO2: {max_fio2}, Duration: {o2_duration})\n"
        f"• Concentrator Needed at Discharge: {o2_conc_dis} {f'(Duration: {o2_conc_duration})' if o2_conc_duration else ''}\n"
        f"• Inotropic Support Required: {inotropes_str}\n\n"
        "Inpatient Medications & Treatment Regimen:\n"
        f"{inpatient_regimen_text}\n\n"
        "POCUS & Diagnostics:\n"
        f"• Lung: {pocus_lung} | Cardiac: {pocus_heart}\n"
        f"• PLR: {pocus_plr} | Vessels: {pocus_vessels} | Others: {pocus_others}\n"
        f"• Interpretation: {pocus_interp}\n"
        f"• ECG: {ecg_findings} | CXR: {cxr_findings} | Other Imagings: {other_ix}\n\n"
    )

    proc_summary_lines = []
    if proc_pleural:
        dt_str = f" [{proc_pleural_date}]" if proc_pleural_date else ""
        det_str = f": {proc_pleural_detail}" if proc_pleural_detail else ""
        proc_summary_lines.append(f"  - Pleural Aspiration / Drainage{dt_str}{det_str}")
    if proc_ascites:
        dt_str = f" [{proc_ascites_date}]" if proc_ascites_date else ""
        det_str = f": {proc_ascites_detail}" if proc_ascites_detail else ""
        proc_summary_lines.append(f"  - Paracentesis{dt_str}{det_str}")
    if proc_lp:
        dt_str = f" [{proc_lp_date}]" if proc_lp_date else ""
        det_str = f": {proc_lp_detail}" if proc_lp_detail else ""
        proc_summary_lines.append(f"  - Lumbar Puncture{dt_str}{det_str}")
    if proc_cvl:
        det_str = f": {proc_cvl_detail}" if proc_cvl_detail else ""
        proc_summary_lines.append(f"  - Central Venous Line (CVL){det_str}")
    if proc_hd_cath:
        det_str = f": {proc_hd_detail}" if proc_hd_detail else ""
        proc_summary_lines.append(f"  - Acute Dialysis Catheter{det_str}")
    if proc_others_chk:
        det_str = f": {proc_others_detail}" if proc_others_detail else ""
        proc_summary_lines.append(f"  - Other Procedure{det_str}")

    if proc_summary_lines:
        summary_text += "• Procedures Performed:\n" + "\n".join(proc_summary_lines) + "\n"
    else:
        summary_text += "• Procedures Performed: None\n"

    if dialysis_needed == "Yes":
        dial_details = f"Mode: {dialysis_mode}"
        if dialysis_date: dial_details += f", Date: {dialysis_date}"
        if renal_recovery: dial_details += f", Recovery: {renal_recovery}"
        summary_text += f"• Acute Dialysis: Yes ({dial_details})\n"
    else:
        summary_text += "• Acute Dialysis: No\n"

    resus_str = ", ".join(resus_code_selected) if resus_code_selected else "Unspecified / See Notes"
    ceiling_items = []
    if resus_hfnc: ceiling_items.append("HFNC")
    if resus_niv: ceiling_items.append("NIV")
    if resus_hfmo: ceiling_items.append("HFM/HFMO")
    if resus_hd: ceiling_items.append("Acute Dialysis")
    if resus_inotropes: ceiling_items.append("Vasopressors/Inotropes")
    if resus_ab: ceiling_items.append("Antibiotics")
    ceiling_str = ", ".join(ceiling_items) if ceiling_items else "NIL specified"

    summary_text += (
        "\nRESUSCITATION DIRECTIVE & CEILING OF CARE:\n"
        f"• Advanced Directive / Wish: {Advance_directive}\n"
        f"• Resuscitation Status: {resus_str}\n"
        f"• Maximum Ceiling Interventions: {ceiling_str}\n"
        f"• Directives / Notes: {resus_notes}\n\n"
        "[ 4. LABORATORY & DIAGNOSTIC TRENDS ]\n"
        f"Flowsheet Trend Table ({date_adm_str} vs {date_mid1_str} vs {date_mid2_str} vs {date_mid3_str} vs {date_dis_str}):\n"
    )

    for param, unit in lab_params:
        v1, v2, v3, v4, v5 = lab_inputs[param]
        if v1 or v2 or v3 or v4 or v5:
            summary_text += f"  - {param:6s} ({unit:15s}):  {v1:8s} | {v2:8s} | {v3:8s} | {v4:8s} | {v5:8s}\n"

    summary_text += "\n[ 5. DISCHARGE MEDICATIONS ]\n"
    med_idx = 1
    if m_panto: summary_text += f"{med_idx}. Pantoprazole: {val_panto}\n"; med_idx += 1
    if m_caco3: summary_text += f"{med_idx}. Calcium Carbonate: {val_caco3}\n"; med_idx += 1
    if m_rocal: summary_text += f"{med_idx}. Rocaltriol: {val_rocal}\n"; med_idx += 1
    if m_sglt2: summary_text += f"{med_idx}. SGLT2i: {val_sglt2}\n"; med_idx += 1
    if m_spiro: summary_text += f"{med_idx}. Spironolactone: {val_spiro}\n"; med_idx += 1
    if m_fruse: summary_text += f"{med_idx}. Frusemide: {val_fruse}\n"; med_idx += 1
    if m_amlo: summary_text += f"{med_idx}. Amlodipine: {val_amlo}\n"; med_idx += 1
    if m_glyprin: summary_text += f"{med_idx}. Glyprin / Aspirin: {val_glyprin}\n"; med_idx += 1
    if m_atorva: summary_text += f"{med_idx}. Atorvastatin: {val_atorva}\n"; med_idx += 1
    if m_biso: summary_text += f"{med_idx}. Bisoprolol: {val_biso}\n"; med_idx += 1
    if m_perin: summary_text += f"{med_idx}. Perindopril: {val_perin}\n"; med_idx += 1
    if m_arni: summary_text += f"{med_idx}. ARNI: {val_arni}\n"; med_idx += 1
    if m_metformin: summary_text += f"{med_idx}. Metformin: {val_metformin}\n"; med_idx += 1
    if m_gliclazide: summary_text += f"{med_idx}. Gliclazide: {val_gliclazide}\n"; med_idx += 1
    if m_insulin: summary_text += f"{med_idx}. Insulin: {val_insulin}\n"; med_idx += 1
    if m_prazosin: summary_text += f"{med_idx}. Prazosin: {val_prazosin}\n"; med_idx += 1
    if m_zincofer: summary_text += f"{med_idx}. Zincofer: {val_zincofer}\n"; med_idx += 1
    if m_vit: summary_text += f"{med_idx}. Vitamins: {val_vit}\n"; med_idx += 1
    if m_morphine: summary_text += f"{med_idx}. Syp Morphine: {val_morphine}\n"; med_idx += 1
    if m_pcm: summary_text += f"{med_idx}. Paracetamol: {val_pcm}\n"; med_idx += 1
    if med_custom: summary_text += f"Additional Meds: {med_custom}\n"

    if chk_p_reg and pred_regimen_summary:
        summary_text += f"\n{pred_regimen_summary}\n"
    if chk_d_reg and dexa_regimen_summary:
        summary_text += f"\n{dexa_regimen_summary}\n"

    summary_text += (
        "\n[ 6. PLAN & SAFETY-NETTING ]\n"
        f"• AIM Clinic Plan: {aim_clinic_plan}\n"
        f"• Other Clinic Follow-ups: {other_clinic_plan}\n"
        f"• Red Flag Warning Signs: {red_flags}\n"
        f"• Discharging MO: {mo_name1} {f'(Incharge: {mo_name2})' if mo_name2 else ''}\n"
        "================================================================================\n"
    )

    st.subheader("Generated Summary Preview")
    st.text_area("Copyable Plain Text Summary (For Hospital EHR)", summary_text, height=350)

    # Word Generator Function
    def generate_docx():
        doc = docx.Document()
        
        # Style page
        section = doc.sections[0]
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

        # Title
        p_title = doc.add_paragraph()
        run_t = p_title.add_run("ACUTE INTERNAL MEDICINE — RAPID REVIEW CLINIC DISCHARGE SUMMARY")
        run_t.bold = True
        run_t.font.size = Pt(14)
        run_t.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Helper function for headings
        def add_sec_heading(title):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(title)
            run.bold = True
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

        # Section 1: Demographics
        add_sec_heading("[ 1. PATIENT DEMOGRAPHICS & WARD METADATA ]")
        t1 = doc.add_table(rows=5, cols=2)
        t1.style = 'Table Grid'
        t1.alignment = WD_TABLE_ALIGNMENT.CENTER
            
        row_data = [
            (f"Patient Name: {patient_name}", f"RN / MRN: {rn_mrn}"),
            (f"IC / Passport: {ic_passport}", f"Age / Sex: {age_sex}"),
            (f"Admission Date: {adm_date}", f"Discharge Date: {dis_date}"),
            (f"Referral Source: {referral_place} ({referral_other})", f"Ward / Bed: {ward_bed}"),
            (f"Attending Doctor: {attending_dr}", f"Subspecialty Team: {subspec_team}")
        ]
        for idx, (c1, c2) in enumerate(row_data):
            t1.rows[idx].cells[0].text = c1
            t1.rows[idx].cells[1].text = c2

        p_meta = doc.add_paragraph()
        p_meta.add_run(f"Contact No: {contact_num} | Caretaker: {caretaker}\nAddress: {address}").italic = True

        # Section 2: Problem List
        add_sec_heading("[ 2. PROBLEM LIST & CLINICAL DIAGNOSES ]")
        p_prob = doc.add_paragraph()
        p_prob.add_run("Primary / Acute Discharge Diagnoses:\n").bold = True
        if p_ccf: p_prob.add_run("• Decompensated Heart Failure\n")
        if p_apo: p_prob.add_run("• Acute Pulmonary Edema\n")
        if p_copd: p_prob.add_run("• Acute Exacerbation of COPD / Asthma\n")
        if p_bronch: p_prob.add_run(f"• Bronchiectasis ({bronch_type}): {bronch_detail}\n")
        if p_ild: p_prob.add_run(f"• Interstitial Lung Disease ({ild_type}): {ild_detail}\n")
        if p_sepsis: p_prob.add_run(f"• Severe Sepsis secondary to: {sepsis_source}\n")
        if p_poisoning: p_prob.add_run(f"• Poisoning: {poisoning_detail}\n")
        if p_pneu: p_prob.add_run(f"• Severe Pneumonia ({pneu_type}): {pneu_detail}\n")
        if p_snake: p_prob.add_run(f"• Snake Bite with Envenomation: {snake_detail}\n")
        if p_acs: p_prob.add_run(f"• ACS ({acs_type}) - TIMI: {acs_timi}\n")
        if s_acs: p_prob.add_run("• Secondary Acute Coronary Syndrome\n")
        if p_pe: p_prob.add_run(f"• Pulmonary Embolism ({pe_category}){f' [Resp Mod: {pe_resp_mod}]' if pe_resp_mod != 'None / Not Applicable' else ''}\n")
        if p_stroke_pri:
            p_prob.add_run(f"• Acute Stroke ({stroke_type}): {stroke_subtype}\n")
            if nihss_score: p_prob.add_run(f"  - NIHSS Score: {nihss_score}\n")
        if p_withdrawal: p_prob.add_run(f"• Drug Withdrawal ({withdrawal_type})\n")
        if p_seizure: p_prob.add_run(f"• Seizures ({seizure_type}){f': {seizure_detail}' if seizure_detail else ''}\n")
        if other_primary: p_prob.add_run(f"• Others: {other_primary}\n")

        p_prob.add_run("\nSecondary Problems & Comorbidities:\n").bold = True
        if s_aki: p_prob.add_run(f"• AKI due to: {aki_cause if aki_cause else 'Diagnosed'}\n")
        if s_ards: p_prob.add_run(f"• ARDS ({ards_severity if ards_severity else 'Diagnosed'})\n")
        if s_rv_dys: p_prob.add_run("• Right Ventricular Dysfunction / Failure\n")
        if s_arrhythmia: p_prob.add_run(f"• Cardiac Arrhythmia: {arrhythmia_detail if arrhythmia_detail else 'Diagnosed'}\n")
        if s_stroke: p_prob.add_run(f"• Stroke (Secondary){f': {s_stroke_detail}' if s_stroke_detail else ''}\n")
        if s_pe_sec: p_prob.add_run(f"• Pulmonary Embolism (Secondary){f': {s_pe_detail}' if s_pe_detail else ''}\n")
        if c_ckd: p_prob.add_run(f"• CKD Stage / Baseline: {ckd_stage if ckd_stage else 'Diagnosed'}\n")
        if c_hf: p_prob.add_run(f"• Heart Failure: {hf_detail if hf_detail else 'HFrEF/HFpEF/HFimpEF'}\n")
        if c_cad: p_prob.add_run(f"• CAD / ASCVD: {cad_detail if cad_detail else 'Diagnosed'}\n")
        if c_dm: p_prob.add_run(f"• Diabetes Mellitus: {dm_detail if dm_detail else 'Diagnosed'}\n")
        if c_htn: p_prob.add_run("• Hypertension\n")
        if other_comorb: p_prob.add_run(f"• Other Comorbidities: {other_comorb}\n")

        p_prob.add_run("\nFunctional Status & Capacity:\n").bold = True
        if frailty_selected: p_prob.add_run(f"• Mobility / Frailty: {', '.join(frailty_selected)}\n")
        if nyha_mets != "Not Specified": p_prob.add_run(f"• NYHA / METS Functional Capacity: {nyha_mets}\n")
        if func_notes: p_prob.add_run(f"• Functional / ADL Notes: {func_notes}\n")

        # Section 3: Progress & Resuscitation
        add_sec_heading("[ 3. CLINICAL PROGRESS, POCUS & CEILING OF CARE ]")
        p_prog = doc.add_paragraph()
        p_prog.add_run(f"Presentation: {presentation}\n")
        p_prog.add_run(f"Interventions: O2: {int_o2} | Fluid: {int_fluid} | Inotropes: {int_inotropes}\n")
        p_prog.add_run(f"Ward Progress: Highest O2: {max_o2_modality} (FiO2: {max_fio2}, Duration: {o2_duration}). Concentrator Needed: {o2_conc_dis} ({o2_conc_duration}). Inotropes: {inotropes_str}.\n\n")
        if inmed_list:
            p_prog.add_run("Inpatient Medications & Treatment Administered:\n").bold = True
            for imline in inmed_list:
                p_prog.add_run(f"{imline}\n")
            p_prog.add_run("\n")
        p_prog.add_run(f"POCUS Summary: Cardiac: {pocus_heart} | Lung: {pocus_lung} | Vessels: {pocus_vessels}\n")
        p_prog.add_run(f"Interpretation: {pocus_interp}\n")
        p_prog.add_run(f"Diagnostics: ECG: {ecg_findings} | CXR: {cxr_findings} | Other Imagings: {other_ix}\n")
        dial_info = f"Acute Dialysis: {dialysis_needed} {f'(Mode: {dialysis_mode}' + (f', Date: {dialysis_date}' if dialysis_date else '') + (f', Recovery: {renal_recovery}' if renal_recovery else '') + ')' if dialysis_needed == 'Yes' else ''}\n"
        p_prog.add_run(dial_info)
        if proc_summary_lines:
            p_prog.add_run("Inpatient Procedures Performed:\n").bold = True
            for pline in proc_summary_lines:
                p_prog.add_run(f"{pline}\n")

        # Resuscitation Highlight Box
        t_resus = doc.add_table(rows=1, cols=1)
        t_resus.style = 'Table Grid'
        cell = t_resus.rows[0].cells[0]
        
        shd = parse_xml(r'<w:shd {} w:fill="FFF2CC"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shd)
        
        p_r = cell.paragraphs[0]
        run_r_head = p_r.add_run("RESUSCITATION STATUS & CEILING OF CARE:\n")
        run_r_head.bold = True
        run_r_head.font.color.rgb = RGBColor(0x99, 0x00, 0x00)
        p_r.add_run(f"• Advanced Directive / Wish: {Advance_directive}\n")
        p_r.add_run(f"• Resuscitation Status: {resus_str}\n")
        p_r.add_run(f"• Maximum Ceiling Interventions: {ceiling_str}\n")
        p_r.add_run(f"• Directives / Notes: {resus_notes if resus_notes else 'NIL'}")

        # Section 4: 5-Column Lab Table
        add_sec_heading("[ 4. KEY DIAGNOSTIC & LABORATORY TRENDS ]")
        
        active_labs = []
        for param, unit in lab_params:
            v1, v2, v3, v4, v5 = lab_inputs[param]
            if v1 or v2 or v3 or v4 or v5:
                active_labs.append((param, unit, v1, v2, v3, v4, v5))
        
        num_rows = len(active_labs) + 1
        t_labs = doc.add_table(rows=num_rows, cols=6)
        t_labs.style = 'Table Grid'
        
        headers = ["Parameter", f"Admission ({date_adm_str})", f"Ward Day 1 ({date_mid1_str})", f"Ward Day 2 ({date_mid2_str})", f"Ward Day 3 ({date_mid3_str})", f"Predischarge ({date_dis_str})"]
        for i, h in enumerate(headers):
            cell = t_labs.rows[0].cells[i]
            cell.text = h
            cell.paragraphs[0].runs[0].font.bold = True
            shd = parse_xml(r'<w:shd {} w:fill="1B365D"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(shd)
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        for row_idx, (p_name, p_unit, v1, v2, v3, v4, v5) in enumerate(active_labs, start=1):
            t_labs.rows[row_idx].cells[0].text = f"{p_name} ({p_unit})"
            t_labs.rows[row_idx].cells[1].text = v1
            t_labs.rows[row_idx].cells[2].text = v2
            t_labs.rows[row_idx].cells[3].text = v3
            t_labs.rows[row_idx].cells[4].text = v4
            t_labs.rows[row_idx].cells[5].text = v5

        # Section 5: Meds & Plan
        add_sec_heading("[ 5. DISCHARGE MEDICATIONS & STEROID REGIMEN ]")
        p_meds = doc.add_paragraph()
        m_count = 1
        if m_panto: p_meds.add_run(f"{m_count}. Pantoprazole: {val_panto}\n"); m_count += 1
        if m_caco3: p_meds.add_run(f"{m_count}. Calcium Carbonate: {val_caco3}\n"); m_count += 1
        if m_rocal: p_meds.add_run(f"{m_count}. Rocaltriol: {val_rocal}\n"); m_count += 1
        if m_sglt2: p_meds.add_run(f"{m_count}. SGLT2i: {val_sglt2}\n"); m_count += 1
        if m_spiro: p_meds.add_run(f"{m_count}. Spironolactone: {val_spiro}\n"); m_count += 1
        if m_fruse: p_meds.add_run(f"{m_count}. Frusemide: {val_fruse}\n"); m_count += 1
        if m_amlo: p_meds.add_run(f"{m_count}. Amlodipine: {val_amlo}\n"); m_count += 1
        if m_glyprin: p_meds.add_run(f"{m_count}. Glyprin / Aspirin: {val_glyprin}\n"); m_count += 1
        if m_atorva: p_meds.add_run(f"{m_count}. Atorvastatin: {val_atorva}\n"); m_count += 1
        if m_biso: p_meds.add_run(f"{m_count}. Bisoprolol: {val_biso}\n"); m_count += 1
        if m_perin: p_meds.add_run(f"{m_count}. Perindopril: {val_perin}\n"); m_count += 1
        if m_arni: p_meds.add_run(f"{m_count}. ARNI: {val_arni}\n"); m_count += 1
        if m_metformin: p_meds.add_run(f"{m_count}. Metformin: {val_metformin}\n"); m_count += 1
        if m_gliclazide: p_meds.add_run(f"{m_count}. Gliclazide: {val_gliclazide}\n"); m_count += 1
        if m_insulin: p_meds.add_run(f"{m_count}. Insulin: {val_insulin}\n"); m_count += 1
        if m_prazosin: p_meds.add_run(f"{m_count}. Prazosin: {val_prazosin}\n"); m_count += 1
        if m_zincofer: p_meds.add_run(f"{m_count}. Zincofer: {val_zincofer}\n"); m_count += 1
        if m_vit: p_meds.add_run(f"{m_count}. Vitamins: {val_vit}\n"); m_count += 1
        if m_morphine: p_meds.add_run(f"{m_count}. Syp Morphine: {val_morphine}\n"); m_count += 1
        if m_pcm: p_meds.add_run(f"{m_count}. Paracetamol: {val_pcm}\n"); m_count += 1
        if med_custom: p_meds.add_run(f"Additional Meds: {med_custom}\n")

        if chk_p_reg and pred_regimen_summary:
            p_meds.add_run(f"\n{pred_regimen_summary}\n").bold = True
        if chk_d_reg and dexa_regimen_summary:
            p_meds.add_run(f"\n{dexa_regimen_summary}\n").bold = True

        add_sec_heading("[ 6. FOLLOW-UP PLAN & RED FLAGS ]")
        p_plan = doc.add_paragraph()
        p_plan.add_run(f"• AIM Clinic Plan: {aim_clinic_plan}\n")
        p_plan.add_run(f"• Other Clinic Follow-ups: {other_clinic_plan}\n")
        p_plan.add_run(f"• Red Flag Warnings: {red_flags}\n\n")
        p_plan.add_run(f"Medical Officer: {mo_name1} {f'(Incharge: {mo_name2})' if mo_name2 else ''}    Signature: _______________ Date: ________")

        target = io.BytesIO()
        doc.save(target)
        return target.getvalue()

    doc_bytes = generate_docx()
    st.download_button(
        label="📄 Download Formatted Word Document (.docx)",
        data=doc_bytes,
        file_name=f"AIM_Discharge_{rn_mrn if rn_mrn else 'Draft'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
