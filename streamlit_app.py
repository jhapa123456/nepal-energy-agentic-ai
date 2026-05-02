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
    <span class="english-title">AI-POWERED ENERGY INTELLIGENCE COMMAND CENTER</span>
  </div>
  <div class="subtle">
    Agentic AI, Agentic RAG, smart-meter analytics, and automated reporting platform for energy-loss detection, revenue recovery, grid reliability, workforce productivity, and data-driven decision support.
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
c3.metric("Redeployable capacity", f"{vals['Modeled maximum redeployable capacity']:.1f} FTE/mo")
c4.metric("Repetitive workload reduction", f"{vals['Modeled repetitive workload reduction']:.1f}%")

st.info("Employee impact is framed as repetitive workload reduction and staff capacity redeployment, not automatic layoffs. Human staff remain responsible for field work, approvals, governance, safety, and customer communication.")

tabs = st.tabs(["Executive Summary", "Risk Detection", "Feeder Priority", "Agentic RAG", "Evaluation", "Reports", "Deployment"])

with tabs[0]:
    st.subheader("Best use case")
    st.write("**NEA Energy Loss, Revenue Protection, and Workforce Productivity Command Center**")
    st.write("This is stronger than a chatbot because it demonstrates end-to-end action: data ingestion, risk scoring, feeder prioritization, evidence retrieval, quality evaluation, and stakeholder-ready reporting.")
    st.dataframe(impact, use_container_width=True)
    colA, colB = st.columns(2)
    with colA:
        st.image(str(main.CHART_WORKFORCE), caption="Workforce productivity / redeployment")
    with colB:
        st.image(str(main.CHART_REVENUE), caption="Revenue recovery opportunity")

with tabs[1]:
    st.subheader("Suspicious customer activity detection")
    st.image(str(main.CHART_RISK), caption="Risk score distribution")
    st.dataframe(risk[["customer_id", "district", "feeder_id", "tariff_category", "kwh_change_pct", "meter_event", "risk_score_0_100", "estimated_revenue_recovery_npr_year", "recommended_action"]].head(50), use_container_width=True)

with tabs[2]:
    st.subheader("Feeder loss and inspection priority")
    st.image(str(main.CHART_FEEDER), caption="Top feeder priority scores")
    st.dataframe(feeder.head(30), use_container_width=True)

with tabs[3]:
    st.subheader("Agentic RAG demo")
    query = st.text_input("Ask a Nepal energy strategy question", "How can NEA reduce non-technical losses with AI?")
    knowledge = pd.read_csv(main.KNOWLEDGE_CSV)
    idx = main.build_rag_index(knowledge)
    top = main.retrieve(idx, query, top_k=4)
    answer = main.generate_grounded_answer(query, top)
    st.text_area("Grounded answer", answer, height=260)
    st.dataframe(top[["chunk_id", "topic", "source_name", "similarity", "chunk_text", "source_url"]], use_container_width=True)

with tabs[4]:
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

with tabs[5]:
    st.subheader("Download generated stakeholder files")
    for file_path, label in [
        (main.PPTX_FILE, "Download PowerPoint"),
        (main.DOCX_FILE, "Download DOCX Report"),
        (main.BUSINESS_IMPACT_CSV, "Download Business Impact CSV"),
        (main.RAG_EVAL_CSV, "Download RAG Evaluation CSV"),
        (main.RISK_CSV, "Download Risk Scores CSV"),
        (main.AGENT_LOG_JSON, "Download Agent Log JSON"),
    ]:
        if Path(file_path).exists():
            with open(file_path, "rb") as f:
                st.download_button(label, f, file_name=Path(file_path).name)

with tabs[6]:
    st.subheader("Deployment Plan for Nepal Electricity Authority (NEA)")

    st.markdown("""
### How NEA Can Deploy This Demo

This demo can be deployed on any cloud platform or on NEA’s own secure internal server.  
For a simple public demonstration, Streamlit Community Cloud can be used.  
For a real NEA production system, deployment should be done inside an approved, secure NEA environment.
Share the public demo URL with NEA leadership, technical teams, and stakeholders.

---

### Option 1: Secure NEA Production Deployment

For real use inside Nepal Electricity Authority, the system should be connected to approved NEA data sources such as:

- Smart meter data
- Customer billing records
- Transformer and feeder data
- Outage and maintenance reports
- Energy loss and theft-related records
- Substation and distribution network data
- Field inspection reports

The production version should include:

- Secure login and role-based access control
- NEA-approved data integration
- Audit logs for every user action
- Human approval before any major decision
- Monitoring dashboards for system health
- Data privacy and cybersecurity controls
- Backup and disaster recovery plan
- Regular model validation by NEA engineers and domain experts

---

### What NEA Can Do With This System

NEA can use this AI-powered platform to support:

- Energy loss detection
- Suspicious electricity usage analysis
- Smart meter monitoring
- Transformer overload prediction
- Outage risk identification
- Revenue leakage analysis
- Customer service improvement
- Data-driven planning for distribution networks
- Automatic reports for management and field offices

---
### Important Note

For a real deployment, NEA should connect only approved internal data sources and add authentication, audit logs, monitoring, cybersecurity controls, and human approval gates.

This system should support NEA engineers and decision-makers.  
It should not replace human judgment, field verification, or official NEA approval processes.
""")
