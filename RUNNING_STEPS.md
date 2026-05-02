# Nepal Energy Agentic AI Demo - Simple Flat Version

This version is intentionally simple: no nested code folders, no package imports, and no secrets required.

## Files

- `main.py` - Generates data, runs agents, RAG, evaluation, charts, PPTX, DOCX, and CSV outputs.
- `streamlit_app.py` - Streamlit dashboard.
- `requirements.txt` - Python packages.
- `synthetic_*.csv` - Synthetic Nepal energy data.
- Generated files: `nepal_energy_agentic_ai_demo.pptx`, `nepal_energy_agentic_ai_report.docx`, `risk_scores.csv`, `rag_evaluation.csv`, charts, and logs.

## Local run on Windows PowerShell

```powershell
cd C:\Users\Parshuram\Downloads\nepal_energy_flat_demo
python -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python main.py
streamlit run streamlit_app.py
```

If activation gives trouble, skip activation and run:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py
```

## Streamlit Community Cloud deployment

1. Create a GitHub repository.
2. Upload all files from this folder to the repository root.
3. Go to Streamlit Community Cloud.
4. Select your GitHub repository.
5. Set the main file path to: `streamlit_app.py`.
6. Deploy.

No API key is required for demo mode.

## Important note

Employee impact is shown as repetitive workload reduction and capacity redeployment, not direct layoffs.
