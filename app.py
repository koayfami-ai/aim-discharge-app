import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
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
        patient_name = st.text_input("Patient Full Name", value="")
        rn_mrn = st.text_input("RN / MRN", value="")
        ic_passport = st.text_input("NRIC / Passport No.", value="")
        age_sex = st.text_input("Age/Gender", value="")
        contact_num = st.text_input("Patient/Caretaker Contact No.", value="")
        caretaker = st.text_input("Caretaker Name & Relationship", value="")
    with col2:
        adm_date = st.date_input("Admission Date")
        dis_date = st.date_input("Date of Discharge")
        referral_place = st.selectbox("Referral Source", ["ED", "Private Hospital/Private Clinic", "Klinik Kesihatan (KK)", "Ambulance", "Transfer from Other Wards HBM", " Transfer from Hospitals Cluster Perai", " Transfer from Hospitals of different Clusters", "Others"])
        referral_other = st.text_input("Specify Source", value="")
        ward_bed = st.text_input("Ward / Bed No.", value="")
        attending_dr = st.text_input("Attending Doctor / Team", value="")
        subspec_team = st.text_input("Subspecialty Team(s) Involved", value="")
        address = st.text_area("Patient Address", value="")

with tab2:
    st.header("2. Problem List & Clinical Diagnoses")
    st.info("💡 Sub-options appear dynamically ONLY when a problem is checked.")
    st.subheader("Primary Diagnosis/Peoblems")
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

        p_withdrawal = st.checkbox(
            "Drug / Substance Withdrawal", key="chk_p_withdrawal")
        withdrawal_type = ""
        withdrawal_toxidrome = ""
        withdrawal_detail = ""
        if p_withdrawal:
            withdrawal_type = st.text_input ("substance type/amphetamine/opiod/toxidromes", value="", key="sel_withdrawal_type")

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
            snake_detail = st.text_input ("Snake type/Local or Systemic Envenomation", value="", key="txt_snake_detail")

        p_acs = st.checkbox("Acute Coronary Syndrome (ACS)", key="chk_p_acs")
        acs_type = ""
        acs_timi = ""
        if p_acs:
            acs_type = st.selectbox("ACS Type", ["UA", "NSTEMI", "STEMI"], key="sel_acs_type")
            acs_timi = st.text_input("TIMI Score", value="", key="txt_acs_timi")

        s_acs = st.checkbox(
            "Secondary Acute Coronary Syndrome )",
            key="chk_s_acs",)
        p_pe = st.checkbox(
            "Pulmonary Embolism (PE)", key="chk_p_pe")
        pe_category = ""
        pe_pert = False
        pe_anticoag = ""
        pe_details = ""

        if p_pe:
            pe_category = st.selectbox(
                "2026 Acute PE Clinical Category (A–E)", 
                [
                    "Category A: Low Risk - Subclinical / Asymptomatic)",
                    "Category B1: Symptomatic — Subsegmental (Single/multiple)",
                    "Category B2: Symptomatic — Non-subsegmental",
                    "Category C1: Low-Intermediate Risk - Normal Biomarkers & No RV Dysfunction",
                    "Category C2: Intermediate-High Risk - Elevated Cardiac",
                    "Category D1: Incipient Cardiopulmonary Failure — Transient hypotension",
                    "Category D2: Incipient Cardiopulmonary Failure — Normotensive shock",
                    "Category E1: Cardiopulmonary Failure — Recurrent/persistent hypotension with cardiogenic shock",
                    "Category E2: Cardiopulmonary Failure — Refractory cardiogenic shock or cardiac arrest",
                ],key="sel_pe_cat_2026",)
            
        # Tiered Respiratory Modifier (+/- R)
        pe_resp_mod = st.selectbox(
        "Respiratory Modifier (+/- R)",
        ["None / Not Applicable",
         "Modifier R (Category C: SpO2 <90%, RR ≥30, need supplemental O2)",
         "Modifier R (Category D: >6L NC or NRB mask)",
         "Modifier R (Category E: Hypoxemic or ventilatory failure)",
        ], key="sel_pe_resp_mod_2026",)
    
        p_stroke_pri = st.checkbox("Acute Stroke", key="chk_p_stroke_pri")
        stroke_type = ""
        stroke_subtype = ""
        nihss_score = ""

        if p_stroke_pri:
            stroke_type = st.selectbox(
                "Stroke Classification",
                ["Ischemic", "Hemorrhagic"], key="sel_stroke_type_pri",)

        # Ischemic -> Oxford/BAMFORD Dropdown | Hemorrhagic -> Direct Text Input
        if stroke_type == "Ischemic":
            stroke_subtype = st.selectbox(
                "Ischemic Subtype (Oxford / BAMFORD)",
                ["LACI (Lacunar Infarct)",
                 "PACI (Partial Anterior Circulation Infarct)",
                 "TACI (Total Anterior Circulation Infarct)",
                 "POCI (Posterior Circulation Infarct)",],
                key="sel_stroke_ischemic_sub",)
        else:
           stroke_subtype = st.text_input("Hemorrhagic Stroke Details", value="", key="txt_stroke_hem_detail")

        # Presentation NIHSS Score
        nihss_score = st.text_input(
        "NIHSS Score / Severity at Presentation", value="",
        key="txt_nihss_score",)

        p_seizure = st.checkbox("Seizures / Status Epilepticus", key="chk_p_seizure")
        seizure_type = ""
        seizure_detail = ""
        if p_seizure:
            seizure_type = st.selectbox(
                "Seizure Type",[
                    "Generalized Tonic-Clonic (GTC)",
                    "Focal Seizure",],key="sel_seizure_type",)
            seizure_detail = st.text_input(
                "Seizure Type",value="",key="txt_seizure_detail",)

        other_primary = st.text_area("Others", value="")

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
    ckd_stage = st.text_input(
        "CKD Stage & Baseline Creatinine", value="", key="txt_ckd_stage"
    )

  c_hf = st.checkbox("HFrEF / HFpEF / HFimpEF", key="chk_c_hf")
  hf_detail = ""
  if c_hf:
    hf_detail = st.text_input(
        "LVEF/Etiology/Echo",
        value="",
        key="txt_hf_detail",
    )

  c_cad = st.checkbox(
      "Coronary Artery Disease", key="chk_c_cad"
  )
  cad_detail = ""
  if c_cad:
    cad_detail = st.text_input(
        "ASCVD Risk Score", value="", key="txt_cad_detail"
    )

with col_c2:
  c_dm = st.checkbox("Diabetes Mellitus", key="chk_c_dm")
  dm_detail = ""
  if c_dm:
    dm_detail = st.text_input(
        "HbA1c % & Target Organ Damage", value="", key="txt_dm_detail"
    )

  c_htn = st.checkbox("Hypertension", key="chk_c_htn")

other_comorb = st.text_area(
    "Other Comorbidities", value="", key="txt_other_comorb"
)
st.subheader("Functional Status & Functional Capacity")
col_f1, col_f2 = st.columns(2)

with col_f1:
  # Multi-select allows checking multiple statuses (e.g. Homebound + Wheelchair-bound)
  frailty_selected = st.multiselect(
      "Frailty / Baseline Mobility Status (Select all that apply)",
      [
          "Independent",
          "Partially dependent",
          "Homebound",
          "Wheelchair-bound",
          "Fully bedridden",
      ],
      key="mult_frailty",
  )

with col_f2:
  nyha_mets = st.selectbox(
      "Functional Capacity",
      [
          "Not Specified",
          "NYHA Class I (No limitation / > 4 METS)",
          "NYHA Class II (Slight limitation / ~4 METS)",
          "NYHA Class III (Marked limitation / < 4 METS)",
          "NYHA Class IV (Severe limitation / Symptoms at rest / < 2 METS)",
      ],
      key="sel_nyha_mets",
  )
  func_notes = st.text_input(
      "Remarks", value="", key="txt_func_notes"
  )
    
with tab3:
st.header("3. Ward Clinical Progress, POCUS & Resuscitation Directives")
st.subheader("Main Presentation & Acute Interventions")
presentation = st.text_area("Main Presentation & Associated Symptoms/Signs", value="", key="txt_presentation")
    
col_int1, col_int2 = st.columns(2)
with col_int1:
    int_o2 = st.text_input("Acute Oxygen Therapy Given", value="", key="txt_int_o2")
    int_fluid = st.text_input("Fluid Bolus", value="", key="txt_int_fluid")
    int_inotropes = st.text_input("Inotropes/Vasopressors", value="", key="txt_int_inotropes")

with col_int2:
        int_scan_by = st.text_input("Admission Bedside Scan Performed By", value="", key="txt_int_scan_by")
        int_scan_findings = st.text_area("Scan Findings", value="", key="txt_int_scan_findings")

st.subheader("Ward Clinical Progress & Hemodynamic Monitoring")
col_prog1, col_prog2 = st.columns(2)
with col_prog1:
    max_o2_modality = st.selectbox("Highest Oxygenation", ["Room Air", "Nasal Cannula", "High Flow Mask (HFM)", "HFNC", "CPAP / BiPAP", "Invasive Mechanical Ventilation"], key="sel_max_o2")
    max_fio2 = st.text_input("Highest FiO2 %)", value="", key="txt_max_fio2")
    o2_duration = st.text_input("Duration of High Flow Oxygenation (HFNC/NIV/IMV)", value="", key="txt_o2_duration")

with col_prog2:
    o2_conc_dis = st.selectbox("Need Temporary Oxygen Concentrator at Discharge?", ["No", "Yes"], key="sel_o2_conc")
    o2_conc_duration = ""
    if o2_conc_dis == "Yes":
        o2_conc_duration = st.text_input("If Yes, Duration Needed", value="", key="txt_o2_conc_dur")
        
    inotropes_ward = st.selectbox ("Inotropes/Vasopressors", ["IV Noradrenaline", "IV dobutamine", "IV Vasopressin", "IV Adrenaline","IV Dopamine"], key="txt_inotropes_ward")

st.subheader("POCUS & Diagnostic Investigations")
pocus_lung = st.text_input("POCUS Lung (e.g. B-lines, pleural effusion, sliding)", value="", key="txt_pocus_lung")
pocus_heart = st.text_input("POCUS Heart (EF, RV strain, pericardial fluid)", value="", key="txt_pocus_heart")
pocus_plr = st.text_input("POCUS Passive Leg Raise (PLR) / Fluid Responsiveness", value="", key="txt_pocus_plr")
pocus_vessels = st.text_input("POCUS Vessels (IVC collapsibility, DVT scan)", value="", key="txt_pocus_vessels")
pocus_others = st.text_input("POCUS Others", value="", key="txt_pocus_others")
pocus_interp = st.text_area("Overall POCUS & Clinical Interpretation", value="", key="txt_pocus_interp")
    
ecg_findings = st.text_area("ECG Findings & Evolution", value="", key="txt_ecg_findings")

    
st.subheader("Inpatient Procedures Performed")
    col_proc1, col_proc2 = st.columns(2)
    with col_proc1:
        proc_pleural = st.checkbox("Pleural Aspiration / Drainage", key="chk_proc_pl")
        proc_ascites = st.checkbox("Peritoneal Aspiration / Paracentesis", key="chk_proc_asc")
        proc_lp = st.checkbox("Lumbar Puncture", key="chk_proc_lp")
    with col_proc2:
        proc_cvl = st.checkbox("Central Venous Line (CVL)", key="chk_proc_cvl")
        proc_hd_cath = st.checkbox("Acute Dialysis Catheter (VasCath)", key="chk_proc_hd")
        proc_others = st.text_input("Other Procedures", value="", key="txt_proc_others")
    
    dialysis_needed = st.selectbox("Acute Dialysis Required During Admission?", ["No", "Yes"], key="sel_dialysis")
    dialysis_duration = ""
    renal_recovery = ""
    if dialysis_needed == "Yes":
        dialysis_duration = st.text_input("Dialysis Duration & Sessions", value="", key="txt_dialysis_dur")
        renal_recovery = st.text_input("Renal Recovery Status at Discharge", value="", key="txt_renal_rec")

    st.subheader("Resuscitation Status & Ceiling of Care Directive")
    resus_status = st.radio("Resuscitation Code", ["Full Code", "Do-Not-Intubate (DNI)", "Full Comfort / Palliative Care"], key="rad_resus_status")
    resus_notes = st.text_area("Resuscitation Details (e.g. Continue HD/inotropes, Highest NIV/HFNC ceiling, Future direction discussed with family)", value="", key="txt_resus_notes")

with tab4:
    st.header("4. Concise Investigation Trend Table (3-Point Serial Format)")
    st.caption("Compare serial values: Admission vs. Midway vs. Discharge Date")
    
    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1:
        date_adm_str = st.text_input("Admission Date Label", value="14/10/2026", key="lbl_adm_date")
    with col_d2:
        date_mid_str = st.text_input("Midway Date Label", value="15/10/2026", key="lbl_mid_date")
    with col_d3:
        date_dis_str = st.text_input("Discharge Date Label", value="16/10/2026", key="lbl_dis_date")

    st.markdown("---")
    
    # Headers
    h1, h2, h3, h4 = st.columns([2.5, 3, 3, 3])
    h1.markdown("**Parameter**")
    h2.markdown(f"**Admission ({date_adm_str})**")
    h3.markdown(f"**Midway ({date_mid_str})**")
    h4.markdown(f"**Discharge ({date_dis_str})**")

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
        ("Others", "ABG/POCUS/Notes")
    ]

    lab_inputs = {}
    for param, unit in lab_params:
        c1, c2, c3, c4 = st.columns([2.5, 3, 3, 3])
        c1.markdown(f"**{param}** <small>({unit})</small>", unsafe_allow_html=True)
        val_adm = c2.text_input(f"{param} Adm", value="", key=f"inp_{param}_adm", label_visibility="collapsed")
        val_mid = c3.text_input(f"{param} Mid", value="", key=f"inp_{param}_mid", label_visibility="collapsed")
        val_dis = c4.text_input(f"{param} Dis", value="", key=f"inp_{param}_dis", label_visibility="collapsed")
        lab_inputs[param] = (val_adm, val_mid, val_dis)

    st.markdown("---")
    st.subheader("Imaging & Modality Summaries")
    cxr_summary = st.text_area("Chest X-Ray (CXR) Findings & Evolution", value="", key="txt_cxr")
    usg_summary = st.text_area("Ultrasound (USG Abdomen/KUB/Vessels)", value="", key="txt_usg")
    ct_mri_summary = st.text_area("CT / MRI Scans", value="", key="txt_ct_mri")
    other_img = st.text_area("Other Imaging / Microbiology", value="", key="txt_other_img")

with tab5:
    st.header("5. Discharge Medication Regimen & Steroid Tapering")
    st.caption("Check ONLY the medications prescribed at discharge.")
    
    st.subheader("Select Prescribed Discharge Medications")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        m_panto = st.checkbox("Pantoprazole", value=True, key="chk_m_panto")
        val_panto = ""
        if m_panto:
            val_panto = st.text_input("Pantoprazole Dose & Freq", value="40mg OD", key="txt_m_panto")
            
        m_caco3 = st.checkbox("Calcium Carbonate", value=False, key="chk_m_caco3")
        val_caco3 = ""
        if m_caco3:
            val_caco3 = st.text_input("CaCO3 Dose & Freq", value="500mg BD", key="txt_m_caco3")
            
        m_rocal = st.checkbox("Rocaltriol / Calcitriol", value=False, key="chk_m_rocal")
        val_rocal = ""
        if m_rocal:
            val_rocal = st.text_input("Rocaltriol Dose & Freq", value="0.25mcg OD", key="txt_m_rocal")

        m_sglt2 = st.checkbox("SGLT2i (Dapagliflozin / Empagliflozin)", value=True, key="chk_m_sglt2")
        val_sglt2 = ""
        if m_sglt2:
            val_sglt2 = st.text_input("SGLT2i Dose & Freq", value="Dapagliflozin 10mg OD", key="txt_m_sglt2")

        m_spiro = st.checkbox("Spironolactone", value=False, key="chk_m_spiro")
        val_spiro = ""
        if m_spiro:
            val_spiro = st.text_input("Spironolactone Dose & Freq", value="25mg OD", key="txt_m_spiro")

        m_fruse = st.checkbox("Frusemide", value=True, key="chk_m_fruse")
        val_fruse = ""
        if m_fruse:
            val_fruse = st.text_input("Frusemide Dose & Freq", value="40mg OD", key="txt_m_fruse")

    with col_m2:
        m_glyprin = st.checkbox("Aspirin / Glyprin", value=True, key="chk_m_glyprin")
        val_glyprin = ""
        if m_glyprin:
            val_glyprin = st.text_input("Glyprin Dose & Freq", value="100mg OD", key="txt_m_glyprin")

        m_atorva = st.checkbox("Atorvastatin", value=True, key="chk_m_atorva")
        val_atorva = ""
        if m_atorva:
            val_atorva = st.text_input("Atorvastatin Dose & Freq", value="40mg ON", key="txt_m_atorva")

        m_insulin = st.checkbox("Insulin Regimen", value=False, key="chk_m_insulin")
        val_insulin = ""
        if m_insulin:
            val_insulin = st.text_input("Insulin Details", value="Mixtard 12u BD", key="txt_m_insulin")

        m_oha = st.checkbox("Oral Hypoglycemic Agents (OHA)", value=False, key="chk_m_oha")
        val_oha = ""
        if m_oha:
            val_oha = st.text_input("OHA Details", value="Metformin 500mg BD", key="txt_m_oha")

        m_zincofer = st.checkbox("Zincofer", value=False, key="chk_m_zincofer")
        val_zincofer = ""
        if m_zincofer:
            val_zincofer = st.text_input("Zincofer Dose & Freq", value="1 tab OD", key="txt_m_zincofer")

        m_vit = st.checkbox("Multivitamins / Supplements", value=False, key="chk_m_vit")
        val_vit = ""
        if m_vit:
            val_vit = st.text_input("Vitamins Dose & Freq", value="1 tab OD", key="txt_m_vit")

    med_custom = st.text_area("Additional Prescribed Discharge Medications", value="", key="txt_med_custom")

    st.subheader("Steroid Tapering Schedule (e.g. Prednisolone)")
    steroid_active = st.checkbox("Patient Discharged on Steroid Tapering Plan", value=False, key="chk_steroid_active")
    if steroid_active:
        steroid_ind = st.text_input("Indication for Steroid Taper", value="COPD / Asthma Exacerbation", key="txt_steroid_ind")
        st1 = st.text_input("Step 1: Dose & Duration", value="Prednisolone 30mg OD for 5 days", key="txt_st1")
        st2 = st.text_input("Step 2: Dose & Duration", value="Then 20mg OD for 5 days", key="txt_st2")
        st3 = st.text_input("Step 3: Dose & Duration", value="Then 10mg OD for 5 days", key="txt_st3")
        st4 = st.text_input("Step 4: Dose & Off Date", value="Then 5mg OD for 5 days, then OFF", key="txt_st4")
    else:
        steroid_ind = ""
        st1 = st2 = st3 = st4 = ""

with tab6:
    st.header("6. Review & Export Options")
    
    st.subheader("AIM Review Clinic & Follow-Up Plan")
    aim_clinic_plan = st.text_area("AIM Clinic Appointment & Issues to Review/Plan", value="Review symptoms, fluid balance, repeat renal profile & electrolyte monitoring.", key="txt_aim_clinic_plan")
    other_clinic_plan = st.text_area("Other Subspecialty Clinic Follow-Ups (Nephro/Cardio/Resp)", value="", key="txt_other_clinic_plan")
    red_flags = st.text_area("Red Flag Warning Signs Provided to Patient", value="Severe dyspnea, chest discomfort, acute syncope, marked decrease in urine output, or sudden swelling.", key="txt_red_flags")
    mo_name = st.text_input("Discharging Medical Officer Name", value="", key="txt_mo_name")

    # Build Dynamic Plain Text Summary (ONLY CHECKED ITEMS ARE ADDED)
    summary_text = f"""
================================================================================
 ACUTE INTERNAL MEDICINE — RAPID REVIEW CLINIC DISCHARGE SUMMARY
================================================================================

[ 1. PATIENT DEMOGRAPHICS & METADATA ]
"""
    summary_text += f"Patient Name: {patient_name}    RN/MRN: {rn_mrn}\n"
    summary_text += f"IC/Passport: {ic_passport}    Age/Sex: {age_sex}\n"
    summary_text += f"Admission Date: {adm_date}    Discharge Date: {dis_date}\n"
    summary_text += f"Referral Source: {referral_place} {f'({referral_other})' if referral_other else ''}\n"
    summary_text += f"Ward/Bed: {ward_bed}    Attending Team: {attending_dr}\n"
    summary_text += f"Subspecialty Team: {subspec_team}\n"
    summary_text += f"Contact No: {contact_num}    Caretaker: {caretaker}\n"
    summary_text += f"Address: {address}\n"

    summary_text += "\n[ 2. PROBLEM LIST & DIAGNOSES ]\nPrimary Diagnoses:\n"
    has_primary = False
    if p_ccf: summary_text += "• Decompensated CCF\n"; has_primary = True
    if p_apo: summary_text += "• Acute Pulmonary Edema\n"; has_primary = True
    if p_copd: summary_text += "• Acute Exacerbation of COPD / Asthma\n"; has_primary = True
    if p_bronch: summary_text += f"• Bronchiectasis ({bronch_type}): {bronch_detail}\n"; has_primary = True
    if p_ild: summary_text += f"• Interstitial Lung Disease ({ild_type}): {ild_detail}\n"; has_primary = True
    if p_sepsis: summary_text += f"• Severe Sepsis secondary to: {sepsis_source}\n"; has_primary = True
    if p_poisoning: summary_text += f"• Poisoning: {poisoning_detail}\n"; has_primary = True
    if p_pneu: summary_text += f"• Severe Pneumonia ({pneu_type}): {pneu_detail}\n"; has_primary = True
    if p_snake: summary_text += f"• Snake Bite with Envenomation: {snake_detail}\n"; has_primary = True
    if p_acs: summary_text += f"• Acute Coronary Syndrome ({acs_type}) - TIMI: {acs_timi}\n"; has_primary = True
    if p_pe: summary_text += f"• Pulmonary Embolism (Risk: {pe_risk})\n"; has_primary = True
    if p_stroke_pri:
        summary_text += f"• Acute Stroke ({stroke_type}): {stroke_subtype}\n"
        if nihss_score: summary_text += f"  - NIHSS Score: {nihss_score}\n"
        has_primary = True
    if p_withdrawal: summary_text += f"• Drug Withdrawal ({withdrawal_type} - Toxidrome: {withdrawal_toxidrome}){f': {withdrawal_detail}' if withdrawal_detail else ''}\n"; has_primary = True
    if p_seizure: summary_text += f"• Seizures ({seizure_type}){f': {seizure_detail}' if seizure_detail else ''}\n"; has_primary = True
    if other_primary: summary_text += f"• Others: {other_primary}\n"; has_primary = True
    if not has_primary: summary_text += "• Nil acute primary diagnoses selected.\n"

    summary_text += "\nSecondary Problems:\n"
    has_sec = False
    if s_aki: summary_text += f"• AKI due to: {aki_cause}\n"; has_sec = True
    if s_ards: summary_text += f"• ARDS ({ards_severity})\n"; has_sec = True
    if s_rv_dys: summary_text += "• Right Ventricular Dysfunction / Failure\n"; has_sec = True
    if s_arrhythmia: summary_text += f"• Cardiac Arrhythmia: {arrhythmia_detail}\n"; has_sec = True
    if s_stroke: summary_text += "• Stroke\n"; has_sec = True
    if s_pe_sec: summary_text += "• Secondary Pulmonary Embolism\n"; has_sec = True
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

    summary_text += f"""
[ 3. CLINICAL PROGRESS & RESUSCITATION DIRECTIVE ]
Main Presentation: {presentation}
Acute Interventions: O2: {int_o2} | Fluid: {int_fluid} | Inotropes: {int_inotropes}
Initial Bedside Scan by {int_scan_by}: {int_scan_findings}

Ward Progress:
• Highest Oxygen Modality: {max_o2_modality} (FiO2: {max_fio2}, Duration: {o2_duration})
• Concentrator Needed at Discharge: {o2_conc_dis} {f'(Duration: {o2_conc_duration})' if o2_conc_duration else ''}
• Inotropic Support: {inotropes_ward}
• Procedures: {'Pleural Aspiration ' if proc_pleural else ''}{'Paracentesis ' if proc_ascites else ''}{'LP ' if proc_lp else ''}{'CVL ' if proc_cvl else ''}{'VasCath ' if proc_hd_cath else ''}{proc_others}
• Acute Dialysis: {dialysis_needed} {f'(Duration: {dialysis_duration}, Recovery: {renal_recovery})' if dialysis_needed == 'Yes' else ''}

POCUS & Diagnostics:
• Lung: {pocus_lung} | Heart: {pocus_heart}
• PLR: {pocus_plr} | Vessels: {pocus_vessels} | Others: {pocus_others}
• Interpretation: {pocus_interp}
• ECG: {ecg_findings}

RESUSCITATION DIRECTIVE & CEILING OF CARE:
• Status: {resus_status}
• Notes / Ceiling Rules: {resus_notes}

[ 4. LABORATORY & DIAGNOSTIC TRENDS ]
Flowsheet Trend Table ({date_adm_str} vs {date_mid_str} vs {date_dis_str}):
"""
    for param, unit in lab_params:
        v1, v2, v3 = lab_inputs[param]
        if v1 or v2 or v3:
            summary_text += f"  - {param:6s} ({unit:15s}):  {v1:10s} | {v2:10s} | {v3:10s}\n"

    summary_text += f"""
Imaging: CXR: {cxr_summary} | USG: {usg_summary} | CT/MRI: {ct_mri_summary} | Others: {other_img}

[ 5. DISCHARGE MEDICATIONS ]
"""
    med_idx = 1
    if m_panto: summary_text += f"{med_idx}. Pantoprazole: {val_panto}\n"; med_idx += 1
    if m_caco3: summary_text += f"{med_idx}. Calcium Carbonate: {val_caco3}\n"; med_idx += 1
    if m_rocal: summary_text += f"{med_idx}. Rocaltriol: {val_rocal}\n"; med_idx += 1
    if m_sglt2: summary_text += f"{med_idx}. SGLT2i: {val_sglt2}\n"; med_idx += 1
    if m_spiro: summary_text += f"{med_idx}. Spironolactone: {val_spiro}\n"; med_idx += 1
    if m_fruse: summary_text += f"{med_idx}. Frusemide: {val_fruse}\n"; med_idx += 1
    if m_glyprin: summary_text += f"{med_idx}. Glyprin / Aspirin: {val_glyprin}\n"; med_idx += 1
    if m_atorva: summary_text += f"{med_idx}. Atorvastatin: {val_atorva}\n"; med_idx += 1
    if m_insulin: summary_text += f"{med_idx}. Insulin: {val_insulin}\n"; med_idx += 1
    if m_oha: summary_text += f"{med_idx}. OHA: {val_oha}\n"; med_idx += 1
    if m_zincofer: summary_text += f"{med_idx}. Zincofer: {val_zincofer}\n"; med_idx += 1
    if m_vit: summary_text += f"{med_idx}. Vitamins: {val_vit}\n"; med_idx += 1
    if med_custom: summary_text += f"Additional Meds: {med_custom}\n"

    if steroid_active:
        summary_text += f"""
Steroid Tapering Plan (Indication: {steroid_ind}):
  - {st1}
  - {st2}
  - {st3}
  - {st4}
"""

    summary_text += f"""
[ 6. PLAN & SAFETY-NETTING ]
• AIM Clinic Plan: {aim_clinic_plan}
• Other Clinic Follow-ups: {other_clinic_plan}
• Red Flag Warning Signs: {red_flags}

Discharging MO: {mo_name}
================================================================================
"""

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
        if p_ccf: p_prob.add_run("• Decompensated CCF\n")
        if p_apo: p_prob.add_run("• Acute Pulmonary Edema\n")
        if p_copd: p_prob.add_run("• Acute Exacerbation of COPD / Asthma\n")
        if p_bronch: p_prob.add_run(f"• Bronchiectasis ({bronch_type}): {bronch_detail}\n")
        if p_ild: p_prob.add_run(f"• Interstitial Lung Disease ({ild_type}): {ild_detail}\n")
        if p_sepsis: p_prob.add_run(f"• Severe Sepsis secondary to: {sepsis_source}\n")
        if p_poisoning: p_prob.add_run(f"• Poisoning: {poisoning_detail}\n")
        if p_pneu: p_prob.add_run(f"• Severe Pneumonia ({pneu_type}): {pneu_detail}\n")
        if p_snake: p_prob.add_run(f"• Snake Bite with Envenomation: {snake_detail}\n")
        if p_acs: p_prob.add_run(f"• ACS ({acs_type}) - TIMI: {acs_timi}\n")
        if p_pe: p_prob.add_run(f"• Pulmonary Embolism (Risk: {pe_risk})\n")
        if p_stroke_pri:
            p_prob.add_run(f"• Acute Stroke ({stroke_type}): {stroke_subtype}\n")
            if nihss_score: p_prob.add_run(f"  - NIHSS Score: {nihss_score}\n")
        if p_withdrawal: p_prob.add_run(f"• Drug Withdrawal ({withdrawal_type} - Toxidrome: {withdrawal_toxidrome}){f': {withdrawal_detail}' if withdrawal_detail else ''}\n")
        if p_seizure: p_prob.add_run(f"• Seizures ({seizure_type}){f': {seizure_detail}' if seizure_detail else ''}\n")
        if other_primary: p_prob.add_run(f"• Others: {other_primary}\n")

        p_prob.add_run("\nSecondary Problems & Comorbidities:\n").bold = True
        if s_aki: p_prob.add_run(f"• AKI due to: {aki_cause}\n")
        if s_ards: p_prob.add_run(f"• ARDS ({ards_severity})\n")
        if s_rv_dys: p_prob.add_run("• Right Ventricular Dysfunction / Failure\n")
        if s_arrhythmia: p_prob.add_run(f"• Cardiac Arrhythmia: {arrhythmia_detail}\n")
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
        p_prog.add_run(f"Ward Progress: Highest O2: {max_o2_modality} (FiO2: {max_fio2}, Duration: {o2_duration}). Concentrator Needed: {o2_conc_dis} ({o2_conc_duration}).\n")
        p_prog.add_run(f"POCUS Summary: Heart: {pocus_heart} | Lung: {pocus_lung} | Vessels: {pocus_vessels}\n")
        p_prog.add_run(f"Interpretation: {pocus_interp}\n")

        # Resuscitation Highlight Box
        t_resus = doc.add_table(rows=1, cols=1)
        t_resus.style = 'Table Grid'
        cell = t_resus.rows[0].cells[0]
        
        # Yellow shading
        shd = parse_xml(r'<w:shd {} w:fill="FFF2CC"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shd)
        
        p_r = cell.paragraphs[0]
        run_r_head = p_r.add_run("RESUSCITATION STATUS & CEILING OF CARE:\n")
        run_r_head.bold = True
        run_r_head.font.color.rgb = RGBColor(0x99, 0x00, 0x00)
        p_r.add_run(f"• Status: {resus_status}\n")
        p_r.add_run(f"• Directives / Notes: {resus_notes}")

        # Section 4: 3-Column Lab Table
        add_sec_heading("[ 4. KEY DIAGNOSTIC & LABORATORY TRENDS ]")
        
        # Calculate active lab rows
        active_labs = []
        for param, unit in lab_params:
            v1, v2, v3 = lab_inputs[param]
            if v1 or v2 or v3:
                active_labs.append((param, unit, v1, v2, v3))
        
        num_rows = len(active_labs) + 1
        t_labs = doc.add_table(rows=num_rows, cols=4)
        t_labs.style = 'Table Grid'
        
        headers = ["Parameter", f"Admission ({date_adm_str})", f"Midway ({date_mid_str})", f"Discharge ({date_dis_str})"]
        for i, h in enumerate(headers):
            cell = t_labs.rows[0].cells[i]
            cell.text = h
            cell.paragraphs[0].runs[0].font.bold = True
            shd = parse_xml(r'<w:shd {} w:fill="1B365D"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(shd)
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        for row_idx, (p_name, p_unit, v_adm, v_mid, v_dis) in enumerate(active_labs, start=1):
            t_labs.rows[row_idx].cells[0].text = f"{p_name} ({p_unit})"
            t_labs.rows[row_idx].cells[1].text = v_adm
            t_labs.rows[row_idx].cells[2].text = v_mid
            t_labs.rows[row_idx].cells[3].text = v_dis

        # Section 5: Meds & Plan
        add_sec_heading("[ 5. DISCHARGE MEDICATIONS & STEROID TAPER ]")
        p_meds = doc.add_paragraph()
        m_count = 1
        if m_panto: p_meds.add_run(f"{m_count}. Pantoprazole: {val_panto}\n"); m_count += 1
        if m_caco3: p_meds.add_run(f"{m_count}. Calcium Carbonate: {val_caco3}\n"); m_count += 1
        if m_rocal: p_meds.add_run(f"{m_count}. Rocaltriol: {val_rocal}\n"); m_count += 1
        if m_sglt2: p_meds.add_run(f"{m_count}. SGLT2i: {val_sglt2}\n"); m_count += 1
        if m_spiro: p_meds.add_run(f"{m_count}. Spironolactone: {val_spiro}\n"); m_count += 1
        if m_fruse: p_meds.add_run(f"{m_count}. Frusemide: {val_fruse}\n"); m_count += 1
        if m_glyprin: p_meds.add_run(f"{m_count}. Glyprin / Aspirin: {val_glyprin}\n"); m_count += 1
        if m_atorva: p_meds.add_run(f"{m_count}. Atorvastatin: {val_atorva}\n"); m_count += 1
        if m_insulin: p_meds.add_run(f"{m_count}. Insulin: {val_insulin}\n"); m_count += 1
        if m_oha: p_meds.add_run(f"{m_count}. OHA: {val_oha}\n"); m_count += 1
        if m_zincofer: p_meds.add_run(f"{m_count}. Zincofer: {val_zincofer}\n"); m_count += 1
        if m_vit: p_meds.add_run(f"{m_count}. Vitamins: {val_vit}\n"); m_count += 1
        if med_custom: p_meds.add_run(f"Additional Meds: {med_custom}\n")

        if steroid_active:
            p_meds.add_run(f"\nSteroid Tapering Plan ({steroid_ind}):\n").bold = True
            p_meds.add_run(f"• {st1}\n• {st2}\n• {st3}\n• {st4}\n")

        add_sec_heading("[ 6. FOLLOW-UP PLAN & RED FLAGS ]")
        p_plan = doc.add_paragraph()
        p_plan.add_run(f"• AIM Clinic Plan: {aim_clinic_plan}\n")
        p_plan.add_run(f"• Other Clinic Follow-ups: {other_clinic_plan}\n")
        p_plan.add_run(f"• Red Flag Warnings: {red_flags}\n\n")
        p_plan.add_run(f"Medical Officer Name: {mo_name}    Signature: _______________ Date: ________")

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

