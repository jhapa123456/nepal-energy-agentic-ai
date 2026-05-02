"""Streamlit dashboard for the flat Nepal Energy Agentic AI demo."""
from pathlib import Path
import json
import pandas as pd
import streamlit as st

import main

ROOT = Path(__file__).resolve().parent

st.set_page_config(
    page_title="नेपाल विद्युत् प्राधिकरण\nकेन्द्रीय कार्यालय, दरबारमार्ग, काठमाडौं, नेपाल\nINTELLIGENT ENERGY COMMAND CENTER",
    page_icon="⚡",
    layout="wide"
)

st.markdown("""
<style>
.block-container {padding-top: 4.2rem;}
.metric-card {background:#f8fafc; padding:18px; border-radius:18px; border:1px solid #e5e7eb;}
.big-title {font-size:clamp(24px, 4vw, 42px); font-weight:900; color:#14385c; line-height:1.18; letter-spacing:-0.5px; white-space:normal; word-break:normal;}
.nepali-title {display:block; font-size:clamp(28px, 3.4vw, 38px); margin-bottom:6px; text-align:center;}
.nepali-subtitle {display:block; font-size:clamp(17px, 2.2vw, 24px); font-weight:700; color:#1e3a8a; margin-bottom:10px; text-align:center;}
.english-title {display:block; font-size:clamp(20px, 2.6vw, 30px); letter-spacing:0.8px; text-align:center;}
.title-box {background:linear-gradient(135deg,#eff6ff,#f8fafc); border:1px solid #dbeafe; border-radius:24px; padding:30px 26px 24px 26px; margin-top:24px; margin-bottom:18px; box-shadow:0 8px 24px rgba(15,23,42,0.08);}
.subtle {color:#64748b; font-size:15px; text-align:center; margin-top:10px;}
</style>
""", unsafe_allow_html=True)

st.markdown('''
<div class="title-box">
  <div class="big-title">
    <span class="nepali-title">⚡ नेपाल विद्युत् प्राधिकरण</span>
    <span class="nepali-subtitle">केन्द्रीय कार्यालय, दरबारमार्ग, काठमाडौं, नेपाल</span>
    <span class="english-title">INTELLIGENT ENERGY COMMAND CENTER</span>
  </div>
  <div class="subtle">
    10-agent AI command system for smart-meter intelligence, billing and revenue recovery, transformer/feeder health, outage prediction, theft detection, field inspection approval, governance, monitoring, and automated executive reporting.
  </div>
</div>
''', unsafe_allow_html=True)

with st.sidebar:
    st.header("Run Demo")
    st.write("Use this button locally or on Streamlit Community Cloud.")
    if st.button("Run / Refresh Autonomous Pipeline", type="primary"):
        with st.spinner("Running agents, RAG, evaluation, charts, PPTX and DOCX..."):
            result = main.run_autonomous_pipeline()
        st.success("Pipeline completed.")
        st.json(result)
    st.divider()
    st.caption("No API key required.")

# Auto-create outputs if missing.
if not Path(main.RISK_CSV).exists() or not Path(main.PPTX_FILE).exists():
    with st.spinner("First run: generating synthetic data and outputs..."):
        main.run_autonomous_pipeline()

risk = pd.read_csv(main.RISK_CSV)
feeder = pd.read_csv(main.FEEDER_PRIORITY_CSV)
rag_eval = pd.read_csv(main.RAG_EVAL_CSV)
impact = pd.read_csv(main.BUSINESS_IMPACT_CSV)
customers = pd.read_csv(main.CUSTOMERS_CSV)

vals = impact.set_index("metric")["value"]

c1, c2, c3, c4 = st.columns(4)
c1.metric("Scaled annual recovery", f"NPR {vals.get('Scaled annual recovery for 25,000-customer pilot', vals['Annual revenue recovery opportunity from high-risk customers']):,.0f}")
c2.metric("Monthly recovery", f"NPR {vals['Monthly revenue recovery opportunity']:,.0f}")
c3.metric("Manual capacity saved", f"{vals['Modeled maximum manual capacity saved']:.1f} FTE/mo")
c4.metric("Repetitive workload reduction", f"{vals['Modeled repetitive workload reduction']:.1f}%")

st.info("Employee impact is framed as reducing repetitive manual data-reading/reporting workload and redeploying staff to revenue recovery, field verification, customer service, governance, safety, and planning. Human approval remains required before major decisions.")

tabs = st.tabs(["Executive Summary", "10 AI Agents", "Risk Detection", "Feeder Priority", "Agentic RAG", "Evaluation", "Reports", "Deployment"])

with tabs[0]:
    st.subheader("Best use case")
    st.write("**NEA Intelligent Energy Command Center: 10-agent workflow for revenue recovery and manual workload reduction**")
    st.write("This is stronger than a chatbot because each agent has a job: collect data, analyze risk, estimate revenue, reduce manual review, require human approval, and generate stakeholder-ready outputs.")
    st.dataframe(impact, use_container_width=True)
    colA, colB = st.columns(2)
    with colA:
        st.image(str(main.CHART_WORKFORCE), caption="Manual employee workload reduction / capacity saved")
    with colB:
        st.image(str(main.CHART_REVENUE), caption="Revenue recovery opportunity")

with tabs[1]:
    st.subheader("10 Specialized AI Agents for NEA")
    st.write("Each agent has a clear NEA role: what data it reads, what analysis it performs, how it reduces repetitive employee workload, and how it supports revenue increase.")
    agent_df = main.agent_roles_dataframe()
    st.dataframe(agent_df, use_container_width=True)
    st.markdown("""
### Simple explanation
- **Smart Meter Intelligence Agent** collects meter readings/events and detects abnormal usage.
- **Billing and Revenue Agent** connects suspicious usage to lost kWh, arrears, and recoverable NPR.
- **Field Inspection and Human Approval Agent** makes sure AI recommendations are checked by people before action.
- **Governance, Security and Audit Agent** protects production deployment with login, role control, audit logs, privacy, and cybersecurity.

This can reduce employees' repetitive manual data-reading work and increase revenue by helping staff focus only on high-risk, high-value cases.
""")

with tabs[2]:
    st.subheader("Suspicious customer activity detection")
    st.image(str(main.CHART_RISK), caption="Risk score distribution")
    st.dataframe(risk[["customer_id", "district", "feeder_id", "tariff_category", "kwh_change_pct", "meter_event", "risk_score_0_100", "estimated_revenue_recovery_npr_year", "recommended_action"]].head(50), use_container_width=True)

with tabs[3]:
    st.subheader("Feeder loss and inspection priority")
    st.image(str(main.CHART_FEEDER), caption="Top feeder priority scores")
    st.dataframe(feeder.head(30), use_container_width=True)

with tabs[4]:
    st.subheader("Agentic RAG demo")
    query = st.text_input("Ask a Nepal energy strategy question", "How can NEA reduce non-technical losses with AI?")
    knowledge = pd.read_csv(main.KNOWLEDGE_CSV)
    idx = main.build_rag_index(knowledge)
    top = main.retrieve(idx, query, top_k=4)
    answer = main.generate_grounded_answer(query, top)
    st.text_area("Grounded answer", answer, height=260)
    st.dataframe(top[["chunk_id", "topic", "source_name", "similarity", "chunk_text", "source_url"]], use_container_width=True)

with tabs[5]:
    st.subheader("RAG quality evaluation")
    st.image(str(main.CHART_RAG), caption="RAG metrics")
    st.dataframe(rag_eval, use_container_width=True)
    st.markdown("""
- **MRR**: whether the expected evidence appears near the top.
- **Relevance**: similarity of the best retrieved chunk to the question.
- **Similarity**: average top-k retrieval similarity.
- **Groundedness**: estimated support from retrieved evidence.
- **Hallucination risk**: estimated unsupported content risk.
""")

with tabs[6]:
    st.subheader("Download generated stakeholder files")
    for file_path, label in [
        (main.PPTX_FILE, "Download PowerPoint"),
        (main.DOCX_FILE, "Download DOCX Report"),
        (main.BUSINESS_IMPACT_CSV, "Download Business Impact CSV"),
        (main.RAG_EVAL_CSV, "Download RAG Evaluation CSV"),
        (main.RISK_CSV, "Download Risk Scores CSV"),
        (main.AGENT_ROLES_CSV, "Download Agent Roles CSV"),
        (main.AGENT_LOG_JSON, "Download Agent Log JSON"),
    ]:
        if Path(file_path).exists():
            with open(file_path, "rb") as f:
                st.download_button(label, f, file_name=Path(file_path).name)

with tabs[7]:
    st.subheader("Deployment Plan for Nepal Electricity Authority (NEA)")

    st.markdown("""
### Streamlit Community Cloud Demo Deployment

This package is ready for public demo deployment on Streamlit Community Cloud. It uses synthetic data only and does **not** require an API key.

1. Create a GitHub repository, for example `nea-intelligent-energy-command-center`.
2. Unzip this project on your computer.
3. Upload all files inside the project folder to the repository root.
4. Go to Streamlit Community Cloud.
5. Click **Create app** / **New app**.
6. Select the GitHub repository.
7. Set **Main file path** to `streamlit_app.py`.
8. Deploy and share the public URL with NEA leadership and stakeholders.

---

### Real NEA Production Data Sources

For production, the system should connect only approved NEA sources:

- Smart meter data
- Customer billing records
- Transformer and feeder data
- Outage and maintenance reports
- Energy loss and theft-related records
- Substation and distribution network data
- Field inspection reports

---

### Production Controls Required

- Secure login and role-based access control
- NEA-approved data integration
- Audit logs for every user action
- Human approval before any major decision
- Monitoring dashboards for system health
- Data privacy and cybersecurity controls
- Backup and disaster recovery plan
- Regular model validation by NEA engineers and domain experts

---

### NEA Business Value

This 10-agent system can help NEA reduce repetitive employee workload for reading data, checking billing files, preparing inspection lists, and writing reports. Staff can then focus on higher-value work: field verification, customer service, revenue recovery, safety, planning, and governance.

It can support energy loss detection, suspicious usage analysis, smart meter monitoring, transformer overload prediction, outage risk identification, revenue leakage analysis, customer service improvement, distribution network planning, and automatic reports for management and field offices.

**Important:** This system supports NEA engineers and decision-makers. It should not replace human judgment, field verification, official approvals, legal review, or customer communication.
""")
