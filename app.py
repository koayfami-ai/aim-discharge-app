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
    "4. Lab & Imaging Trends (3-Point)",
    "5. Medications & Steroid Taper",
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
    st.subheader("Primary Diagnosis/Problems")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        p_ccf = st.checkbox("Decompensated CCF")
        p_apo = st.checkbox("Acute Pulmonary Edema")
        p_copd = st.checkbox("Acute Exacerbation of COPD / Asthma")
        p_bronchiec = st.checkbox("Acute Exacerbation of bronchiectasis")
        bronchiec_source = st.selectbox ("Infective", "Non-Infective")
        p_ILD = st.checkbox("Acute Exacerbation of ILD")
        ILD_source = st.selectbox("Infective", "Non-Infective")
        p_sepsis = st.checkbox("Severe Sepsis")
        sepsis_source = st.text_input("Sepsis Source (if checked)", value="")
        p_poisoning = st.checkbox("Poisoning")
        poisoning_source = st.text_input(" Type (if checked)", value="")
        
    with col_p2:
        p_pneu = st.checkbox("Severe Pneumonia")
        pneu_type = st.selectbox("Pneumonia Etiology", ["Not Specified", "Viral", "Bacterial", "Atypical / Others"])
        pneu_detail = st.text_input("Pneumonia Organism / Details", value="")
        p_snake = st.checkbox("Snake Bite with Envenomation")
        snake_detail = st.text_input("Local/Systemic Envenomation Details", value="")
        p_acs = st.checkbox("Acute Coronary Syndrome (ACS)")
        acs_type = st.selectbox("ACS Type & TIMI Score", ["UA", "NSTEMI", "STEMI / MI"])
        acs_timi = st.text_input("TIMI Score & Details", value="")
        p_pe_primary = st.checkbox("Pulmonary Embolism (Primary)")
        pe_risk = st.text_input("PE Risk Stratification (e.g. High/Int/Low)", value="")
    
    other_primary = st.text_area("Other Primary Diagnoses", value="")

    st.subheader("Secondary Problems")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        s_aki = st.checkbox("Acute Kidney Injury (AKI)")
        aki_cause = st.text_input("AKI Etiology", value="")
        s_ards = st.checkbox("ARDS")
        ards_severity = st.selectbox("ARDS Severity", ["Mild", "Moderate", "Severe"])
        s_rv_dys = st.checkbox("Right Ventricular Dysfunction / Failure")
    with col_s2:
        s_arrhythmia = st.checkbox("Cardiac Arrhythmia (AF/VT/SVT/etc.)")
        arrhythmia_detail = st.text_input("Arrhythmia Details", value="")
        s_stroke = st.checkbox("Stroke (Ischemic/Hemorrhagic)")
        s_pe_sec = st.checkbox("Pulmonary Embolism (Secondary)")
    other_secondary = st.text_area("Additional Secondary Problems", value="")

    st.subheader("Comorbidities")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        c_ckd = st.checkbox("CKD / ESRF")
        ckd_stage = st.text_input("CKD Stage & Baseline Creatinine", value="")
        c_hf = st.checkbox("HFrEF / HFpEF")
        hf_detail = st.text_input("LVEF % & Etiology (Ischemic/Non-Ischemic)", value="")
        c_cad = st.checkbox("Coronary Artery Disease (ASCVD)")
        cad_detail = st.text_input("CAD / ASCVD Details", value="")
    with col_c2:
        c_dm = st.checkbox("Diabetes Mellitus")
        dm_detail = st.text_input("HbA1c % & Target Organ Damage", value="")
        c_htn = st.checkbox("Hypertension")
        c_frailty = st.selectbox("Baseline Mobility & Frailty", ["Independent", "Wheelchair-bound", "Bedridden"])
    other_comorb = st.text_area("Other Comorbidities", value="")

with tab3:
    st.header("3. Clinical Progress, POCUS & Resuscitation Directives")
    st.subheader("Main Presentation & Acute Interventions")
    presentation = st.text_area("Main Presentation & Associated Symptoms/Signs", value="")
    
    col_int1, col_int2 = st.columns(2)
    with col_int1:
        int_o2 = st.text_input("Acute Oxygen Therapy Given", value="")
        int_fluid = st.text_input("Fluid Resuscitation / Restriction", value="")
        int_inotropes = st.text_input("Inotropic Support Given", value="")
    with col_int2:
        int_scan_by = st.text_input("Bedside Scan Performed By", value="")
        int_scan_findings = st.text_area("Initial Scan Findings", value="")

    st.subheader("Ward Progress & Respiratory Support")
    col_prog1, col_prog2 = st.columns(2)
    with col_prog1:
        max_o2_modality = st.selectbox("Highest Oxygen Modality Required", ["Room Air", "Nasal Cannula", "High Flow Mask (HFM)", "HFNC", "CPAP / BiPAP", "Invasive Mechanical Ventilation"])
        max_fio2 = st.text_input("Highest FiO2 / Flow Rate (%)", value="")
        o2_duration = st.text_input("Duration of Advanced Oxygen Support", value="")
    with col_prog2:
        o2_conc_dis = st.selectbox("Need Temporary Oxygen Concentrator at Discharge?", ["No", "Yes"])
        o2_conc_duration = st.text_input("If Yes, Duration Needed", value="")
        inotropes_ward = st.text_input("Ward Inotropic Support Details", value="")

    st.subheader("POCUS & Diagnostic Investigations")
    pocus_lung = st.text_input("POCUS Lung (e.g. B-lines, pleural effusion, sliding)", value="")
    pocus_heart = st.text_input("POCUS Heart (EF, RV strain, pericardial fluid)", value="")
    pocus_plr = st.text_input("POCUS Passive Leg Raise (PLR) / Fluid Responsiveness", value="")
    pocus_vessels = st.text_input("POCUS Vessels (IVC collapsibility, DVT scan)", value="")
    pocus_others = st.text_input("POCUS Others", value="")
    pocus_interp = st.text_area("Overall POCUS & Clinical Interpretation", value="")
    ecg_findings = st.text_area("ECG Findings & Evolution", value="")

    st.subheader("Inpatient Procedures Performed")
    col_proc1, col_proc2 = st.columns(2)
    with col_proc1:
        proc_pleural = st.checkbox("Pleural Aspiration / Drainage")
        proc_ascites = st.checkbox("Peritoneal Aspiration / Paracentesis")
        proc_lp = st.checkbox("Lumbar Puncture")
    with col_proc2:
        proc_cvl = st.checkbox("Central Venous Line (CVL)")
        proc_hd_cath = st.checkbox("Acute Dialysis Catheter (VasCath)")
        proc_others = st.text_input("Other Procedures", value="")
    
    dialysis_needed = st.selectbox("Acute Dialysis Required During Admission?", ["No", "Yes"])
    dialysis_duration = st.text_input("If Yes, Dialysis Duration & Sessions", value="")
    renal_recovery = st.text_input("Renal Recovery Status at Discharge", value="")

    st.subheader("Resuscitation Status & Ceiling of Care Directive")
    resus_status = st.radio("Resuscitation Code", ["Full Code", "Do-Not-Intubate (DNI)", "Full Comfort / Palliative Care"])
    resus_notes = st.text_area("Resuscitation Details (e.g. Continue HD/inotropes, Highest NIV/HFNC ceiling, Future direction discussed with family)", value="")

with tab4:
    st.header("4. Key Diagnostic & Laboratory Trends (3-Point Comparison)")
    st.caption("Compare Initial (Admission), Midway, and Pre-Discharge Lab Values")
    
    st.subheader("Arterial / Venous Blood Gas (ABG/VBG)")
    col_abg1, col_abg2, col_abg3 = st.columns(3)
    with col_abg1:
        st.markdown("**Admission / Initial**")
        abg_adm_ph = st.text_input("pH (Adm)", value="")
        abg_adm_pco2 = st.text_input("pCO2 (Adm)", value="")
        abg_adm_po2 = st.text_input("pO2 (Adm)", value="")
        abg_adm_hco3 = st.text_input("HCO3 (Adm)", value="")
        abg_adm_lac = st.text_input("Lactate (Adm)", value="")
    with col_abg2:
        st.markdown("**Midway Trend**")
        abg_mid_ph = st.text_input("pH (Mid)", value="")
        abg_mid_pco2 = st.text_input("pCO2 (Mid)", value="")
        abg_mid_po2 = st.text_input("pO2 (Mid)", value="")
        abg_mid_hco3 = st.text_input("HCO3 (Mid)", value="")
        abg_mid_lac = st.text_input("Lactate (Mid)", value="")
    with col_abg3:
        st.markdown("**Pre-Discharge / Latest**")
        abg_dis_ph = st.text_input("pH (Dis)", value="")
        abg_dis_pco2 = st.text_input("pCO2 (Dis)", value="")
        abg_dis_po2 = st.text_input("pO2 (Dis)", value="")
        abg_dis_hco3 = st.text_input("HCO3 (Dis)", value="")
        abg_dis_lac = st.text_input("Lactate (Dis)", value="")

    st.subheader("Renal & Electrolytes Profile")
    col_r1, col_r2, col_r3 = st.columns(3)
    with col_r1:
        st.markdown("**Admission**")
        r_adm_urea = st.text_input("Urea (Adm)", value="")
        r_adm_creat = st.text_input("Creatinine (Adm)", value="")
        r_adm_k = st.text_input("K+ (Adm)", value="")
        r_adm_na = st.text_input("Na+ (Adm)", value="")
    with col_r2:
        st.markdown("**Midway**")
        r_mid_urea = st.text_input("Urea (Mid)", value="")
        r_mid_creat = st.text_input("Creatinine (Mid)", value="")
        r_mid_k = st.text_input("K+ (Mid)", value="")
        r_mid_na = st.text_input("Na+ (Mid)", value="")
    with col_r3:
        st.markdown("**Pre-Discharge**")
        r_dis_urea = st.text_input("Urea (Dis)", value="")
        r_dis_creat = st.text_input("Creatinine (Dis)", value="")
        r_dis_k = st.text_input("K+ (Dis)", value="")
        r_dis_na = st.text_input("Na+ (Dis)", value="")

    st.subheader("Full Blood Count (FBC)")
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        st.markdown("**Admission**")
        f_adm_wbc = st.text_input("WBC (Adm)", value="")
        f_adm_lym = st.text_input("Lymphocyte Count & % (Adm)", value="")
        f_adm_hb = st.text_input("Hb (Adm)", value="")
        f_adm_plt = st.text_input("Platelets (Adm)", value="")
    with col_f2:
        st.markdown("**Midway**")
        f_mid_wbc = st.text_input("WBC (Mid)", value="")
        f_mid_lym = st.text_input("Lymphocyte Count & % (Mid)", value="")
        f_mid_hb = st.text_input("Hb (Mid)", value="")
        f_mid_plt = st.text_input("Platelets (Mid)", value="")
    with col_f3:
        st.markdown("**Pre-Discharge**")
        f_dis_wbc = st.text_input("WBC (Dis)", value="")
        f_dis_lym = st.text_input("Lymphocyte Count & % (Dis)", value="")
        f_dis_hb = st.text_input("Hb (Dis)", value="")
        f_dis_plt = st.text_input("Platelets (Dis)", value="")

    st.subheader("Cardiac & Inflammatory Markers")
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        trop_i = st.text_input("Troponin I / T", value="")
        pro_bnp = st.text_input("NT-proBNP / BNP", value="")
    with col_m2:
        crp_esr = st.text_input("CRP / ESR", value="")
        other_markers = st.text_input("Other Biomarkers (PCT, LFT, etc.)", value="")

    st.subheader("Imaging & Modality Summaries")
    cxr_summary = st.text_area("Chest X-Ray (CXR) Findings & Evolution", value="")
    usg_summary = st.text_area("Ultrasound (USG Abdomen/KUB/Vessels)", value="")
    ct_mri_summary = st.text_area("CT / MRI Scans", value="")
    other_img = st.text_area("Other Imaging / Microbiology", value="")

with tab5:
    st.header("5. Discharge Medication Regimen & Steroid Tapering")
    st.caption("Reconcile Routine Discharge Medications & Complex Tapering Schedules")
    
    st.subheader("Standard Discharge Medications")
    med_pantoprazole = st.text_input("Pantoprazole (Dose & Freq)", value="40mg OD")
    med_caco3 = st.text_input("Calcium Carbonate (Dose & Freq)", value="500mg BD")
    med_rocaltriol = st.text_input("Rocaltriol / Calcitriol", value="0.25mcg OD")
    med_sglt2 = st.text_input("SGLT2i (Dapagliflozin / Empagliflozin)", value="Dapagliflozin 10mg OD")
    med_spiro = st.text_input("Spironolactone", value="25mg OD")
    med_frusemide = st.text_input("Frusemide / Loop Diuretic", value="40mg OD")
    med_glyprin = st.text_input("Aspirin / Glyprin", value="100mg OD")
    med_atorva = st.text_input("Atorvastatin", value="40mg ON")
    med_insulin = st.text_input("Insulin Regimen", value="")
    med_oha = st.text_input("Oral Hypoglycemic Agents (OHA)", value="")
    med_zincofer = st.text_input("Zincofer", value="1 tab OD")
    med_vitamins = st.text_input("Multivitamins / Supplements", value="")
    med_custom = st.text_area("Additional Discharge Medications", value="")

    st.subheader("Steroid Tapering Schedule (e.g. Prednisolone)")
    steroid_active = st.checkbox("Patient Discharged on Steroid Tapering Plan", value=False)
    if steroid_active:
        steroid_ind = st.text_input("Indication for Steroid Taper", value="COPD / Asthma Exacerbation / Vasculitis / Glomerulonephritis")
        st1 = st.text_input("Step 1: Dose & Duration", value="Prednisolone 30mg OD for 5 days")
        st2 = st.text_input("Step 2: Dose & Duration", value="Then 20mg OD for 5 days")
        st3 = st.text_input("Step 3: Dose & Duration", value="Then 10mg OD for 5 days")
        st4 = st.text_input("Step 4: Dose & Off Date", value="Then 5mg OD for 5 days, then OFF")
    else:
        steroid_ind = ""
        st1 = st2 = st3 = st4 = ""

with tab6:
    st.header("6. Review & Export Options")
    
    st.subheader("AIM Review Clinic & Follow-Up Plan")
    aim_clinic_plan = st.text_area("AIM Clinic Appointment & Issues to Review/Plan", value="Review symptoms, fluid balance, repeat renal profile & electrolyte monitoring.")
    other_clinic_plan = st.text_area("Other Subspecialty Clinic Follow-Ups (Nephro/Cardio/Resp)", value="")
    red_flags = st.text_area("Red Flag Warning Signs Provided to Patient", value="Severe dyspnea, chest discomfort, acute syncope, marked decrease in urine output, or sudden swelling.")
    mo_name = st.text_input("Discharging Medical Officer Name", value="")

    # Build document string preview
    summary_text = f"""
================================================================================
 ACUTE INTERNAL MEDICINE — RAPID REVIEW CLINIC DISCHARGE SUMMARY
================================================================================

[ 1. PATIENT DEMOGRAPHICS & METADATA ]
Patient Name: {patient_name}    RN/MRN: {rn_mrn}
IC/Passport: {ic_passport}    Age/Sex: {age_sex}
Admission Date: {adm_date}    Discharge Date: {dis_date}
Referral Source: {referral_place} {f'({referral_other})' if referral_other else ''}
Ward/Bed: {ward_bed}    Attending Team: {attending_dr}
Subspecialty Team: {subspec_team}
Contact No: {contact_num}    Caretaker: {caretaker}
Address: {address}

[ 2. PROBLEM LIST & DIAGNOSES ]
Primary Diagnoses:
"""
    if p_ccf: summary_text += "• Decompensated CCF\n"
    if p_apo: summary_text += "• Acute Pulmonary Edema\n"
    if p_copd: summary_text += "• Acute Exacerbation of COPD / Asthma\n"
    if p_sepsis: summary_text += f"• Severe Sepsis secondary to: {sepsis_source}\n"
    if p_poisoning: summary_text += f"• Poisoning: {poisoning_detail}\n"
    if p_pneu: summary_text += f"• Severe Pneumonia ({pneu_type}): {pneu_detail}\n"
    if p_snake: summary_text += f"• Snake Bite with Envenomation: {snake_detail}\n"
    if p_acs: summary_text += f"• Acute Coronary Syndrome ({acs_type}) - TIMI: {acs_timi}\n"
    if p_pe_primary: summary_text += f"• Pulmonary Embolism (Risk: {pe_risk})\n"
    if other_primary: summary_text += f"• Others: {other_primary}\n"

    summary_text += "\nSecondary Problems:\n"
    if s_aki: summary_text += f"• AKI due to: {aki_cause}\n"
    if s_ards: summary_text += f"• ARDS ({ards_severity})\n"
    if s_rv_dys: summary_text += "• Right Ventricular Dysfunction / Failure\n"
    if s_arrhythmia: summary_text += f"• Cardiac Arrhythmia: {arrhythmia_detail}\n"
    if s_stroke: summary_text += "• Stroke\n"
    if s_pe_sec: summary_text += "• Secondary Pulmonary Embolism\n"
    if other_secondary: summary_text += f"• Others: {other_secondary}\n"

    summary_text += "\nComorbidities:\n"
    if c_ckd: summary_text += f"• CKD Stage / Baseline: {ckd_stage}\n"
    if c_hf: summary_text += f"• Heart Failure: {hf_detail}\n"
    if c_cad: summary_text += f"• CAD / ASCVD: {cad_detail}\n"
    if c_dm: summary_text += f"• Diabetes Mellitus: {dm_detail}\n"
    if c_htn: summary_text += "• Hypertension\n"
    summary_text += f"• Baseline Mobility: {c_frailty}\n"
    if other_comorb: summary_text += f"• Other Comorbidities: {other_comorb}\n"

    summary_text += f"""
[ 3. CLINICAL PROGRESS & RESUSCITATION DIRECTIVE ]
Main Presentation: {presentation}
Acute Interventions: O2: {int_o2} | Fluid: {int_fluid} | Inotropes: {int_inotropes}
Initial Bedside Scan by {int_scan_by}: {int_scan_findings}

Ward Progress:
• Highest Oxygen Modality: {max_o2_modality} (FiO2: {max_fio2}, Duration: {o2_duration})
• Concentrator Needed at Discharge: {o2_conc_dis} (Duration: {o2_conc_duration})
• Inotropic Support: {inotropes_ward}
• Procedures: {'Pleural Aspiration ' if proc_pleural else ''}{'Paracentesis ' if proc_ascites else ''}{'LP ' if proc_lp else ''}{'CVL ' if proc_cvl else ''}{'VasCath ' if proc_hd_cath else ''}{proc_others}
• Acute Dialysis: {dialysis_needed} (Duration: {dialysis_duration}, Recovery: {renal_recovery})

POCUS & Diagnostics:
• Lung: {pocus_lung} | Heart: {pocus_heart}
• PLR: {pocus_plr} | Vessels: {pocus_vessels} | Others: {pocus_others}
• Interpretation: {pocus_interp}
• ECG: {ecg_findings}

RESUSCITATION DIRECTIVE & CEILING OF CARE:
• Status: {resus_status}
• Notes / Ceiling Rules: {resus_notes}

[ 4. LABORATORY & DIAGNOSTIC TRENDS ]
ABG/VBG:
  - Adm: pH {abg_adm_ph}, pCO2 {abg_adm_pco2}, pO2 {abg_adm_po2}, HCO3 {abg_adm_hco3}, Lac {abg_adm_lac}
  - Mid: pH {abg_mid_ph}, pCO2 {abg_mid_pco2}, pO2 {abg_mid_po2}, HCO3 {abg_mid_hco3}, Lac {abg_mid_lac}
  - Dis: pH {abg_dis_ph}, pCO2 {abg_dis_pco2}, pO2 {abg_dis_po2}, HCO3 {abg_dis_hco3}, Lac {abg_dis_lac}

Renal Profile & Electrolytes:
  - Adm: Urea {r_adm_urea}, Creat {r_adm_creat}, K+ {r_adm_k}, Na+ {r_adm_na}
  - Mid: Urea {r_mid_urea}, Creat {r_mid_creat}, K+ {r_mid_k}, Na+ {r_mid_na}
  - Dis: Urea {r_dis_urea}, Creat {r_dis_creat}, K+ {r_dis_k}, Na+ {r_dis_na}

FBC:
  - Adm: WBC {f_adm_wbc}, Lymph {f_adm_lym}, Hb {f_adm_hb}, Plt {f_adm_plt}
  - Mid: WBC {f_mid_wbc}, Lymph {f_mid_lym}, Hb {f_mid_hb}, Plt {f_mid_plt}
  - Dis: WBC {f_dis_wbc}, Lymph {f_dis_lym}, Hb {f_dis_hb}, Plt {f_dis_plt}

Cardiac & Biomarkers: Trop I: {trop_i} | ProBNP: {pro_bnp} | CRP: {crp_esr} | Others: {other_markers}
Imaging: CXR: {cxr_summary} | USG: {usg_summary} | CT/MRI: {ct_mri_summary} | Others: {other_img}

[ 5. DISCHARGE MEDICATIONS ]
1. Pantoprazole: {med_pantoprazole}
2. Calcium Carbonate: {med_caco3}
3. Rocaltriol: {med_rocaltriol}
4. SGLT2i: {med_sglt2}
5. Spironolactone: {med_spiro}
6. Frusemide: {med_frusemide}
7. Glyprin: {med_glyprin}
8. Atorvastatin: {med_atorva}
9. Insulin: {med_insulin}
10. OHA: {med_oha}
11. Zincofer: {med_zincofer}
12. Vitamins: {med_vitamins}
Other Meds: {med_custom}
"""
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
        if p_sepsis: p_prob.add_run(f"• Severe Sepsis secondary to: {sepsis_source}\n")
        if p_poisoning: p_prob.add_run(f"• Poisoning: {poisoning_detail}\n")
        if p_pneu: p_prob.add_run(f"• Severe Pneumonia ({pneu_type}): {pneu_detail}\n")
        if p_snake: p_prob.add_run(f"• Snake Bite with Envenomation: {snake_detail}\n")
        if p_acs: p_prob.add_run(f"• ACS ({acs_type}) - TIMI: {acs_timi}\n")
        if p_pe_primary: p_prob.add_run(f"• Pulmonary Embolism (Risk: {pe_risk})\n")
        if other_primary: p_prob.add_run(f"• Others: {other_primary}\n")

        p_prob.add_run("\nSecondary Problems & Comorbidities:\n").bold = True
        if s_aki: p_prob.add_run(f"• AKI due to: {aki_cause}\n")
        if s_ards: p_prob.add_run(f"• ARDS ({ards_severity})\n")
        if s_rv_dys: p_prob.add_run("• Right Ventricular Dysfunction / Failure\n")
        if s_arrhythmia: p_prob.add_run(f"• Cardiac Arrhythmia: {arrhythmia_detail}\n")
        if c_ckd: p_prob.add_run(f"• CKD Stage / Baseline: {ckd_stage}\n")
        if c_hf: p_prob.add_run(f"• Heart Failure: {hf_detail}\n")
        if c_cad: p_prob.add_run(f"• CAD / ASCVD: {cad_detail}\n")
        if c_dm: p_prob.add_run(f"• Diabetes Mellitus: {dm_detail}\n")
        if c_htn: p_prob.add_run("• Hypertension\n")
        p_prob.add_run(f"• Baseline Mobility: {c_frailty}\n")

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
        t_labs = doc.add_table(rows=5, cols=4)
        t_labs.style = 'Table Grid'
        
        headers = ["Parameter", "Admission / Initial", "Midway Trend", "Pre-Discharge / Latest"]
        for i, h in enumerate(headers):
            cell = t_labs.rows[0].cells[i]
            cell.text = h
            cell.paragraphs[0].runs[0].font.bold = True
            shd = parse_xml(r'<w:shd {} w:fill="1B365D"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(shd)
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        lab_rows = [
            ("ABG/VBG (pH / pCO2 / pO2 / HCO3)", f"{abg_adm_ph} / {abg_adm_pco2} / {abg_adm_po2} / {abg_adm_hco3}", f"{abg_mid_ph} / {abg_mid_pco2} / {abg_mid_po2} / {abg_mid_hco3}", f"{abg_dis_ph} / {abg_dis_pco2} / {abg_dis_po2} / {abg_dis_hco3}"),
            ("Renal (Urea / Creat / K+ / Na+)", f"{r_adm_urea} / {r_adm_creat} / {r_adm_k} / {r_adm_na}", f"{r_mid_urea} / {r_mid_creat} / {r_mid_k} / {r_mid_na}", f"{r_dis_urea} / {r_dis_creat} / {r_dis_k} / {r_dis_na}"),
            ("FBC (WBC / Lymph% / Hb / Plt)", f"{f_adm_wbc} / {f_adm_lym} / {f_adm_hb} / {f_adm_plt}", f"{f_mid_wbc} / {f_mid_lym} / {f_mid_hb} / {f_mid_plt}", f"{f_dis_wbc} / {f_dis_lym} / {f_dis_hb} / {f_dis_plt}"),
            ("Biomarkers & Imaging", f"Trop I: {trop_i}", f"ProBNP: {pro_bnp} | CRP: {crp_esr}", f"CXR: {cxr_summary}")
        ]
        for row_idx, data in enumerate(lab_rows, start=1):
            for col_idx, text in enumerate(data):
                t_labs.rows[row_idx].cells[col_idx].text = text

        # Section 5: Meds & Plan
        add_sec_heading("[ 5. DISCHARGE MEDICATIONS & STEROID TAPER ]")
        p_meds = doc.add_paragraph()
        p_meds.add_run("1. Pantoprazole: ").bold = True
        p_meds.add_run(f"{med_pantoprazole}\n")
        p_meds.add_run("2. Calcium Carbonate: ").bold = True
        p_meds.add_run(f"{med_caco3}\n")
        p_meds.add_run("3. Rocaltriol: ").bold = True
        p_meds.add_run(f"{med_rocaltriol}\n")
        p_meds.add_run("4. SGLT2i: ").bold = True
        p_meds.add_run(f"{med_sglt2}\n")
        p_meds.add_run("5. Spironolactone: ").bold = True
        p_meds.add_run(f"{med_spiro}\n")
        p_meds.add_run("6. Frusemide: ").bold = True
        p_meds.add_run(f"{med_frusemide}\n")
        p_meds.add_run("7. Glyprin / Aspirin: ").bold = True
        p_meds.add_run(f"{med_glyprin}\n")
        p_meds.add_run("8. Atorvastatin: ").bold = True
        p_meds.add_run(f"{med_atorva}\n")
        p_meds.add_run("9. Insulin / OHA: ").bold = True
        p_meds.add_run(f"{med_insulin} | {med_oha}\n")
        p_meds.add_run("10. Zincofer / Vitamins: ").bold = True
        p_meds.add_run(f"{med_zincofer} | {med_vitamins}\n")
        if med_custom:
            p_meds.add_run(f"Additional Meds: {med_custom}\n")

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
