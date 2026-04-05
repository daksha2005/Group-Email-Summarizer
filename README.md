# 📧 AI-Powered Group Email Synthesiser (Enron Dataset)

> Turn hundreds of chaotic, unstructured email threads into a sharp, structured intelligence dashboard — **zero API keys, runs fully locally with 100% data privacy.**

---

## 🎯 Assignment Deliverables

| Deliverable | Verification / Link |
| :--- | :--- |
| **1. Architecture Note** | Full mapping and logic is housed in **[ARCHITECTURE.md](./ARCHITECTURE.md)** |
| **2. Public Dashboard Link** | **[👉 LINK TO LIVE STREAMLIT APP (Replace with your link) 👈]()** |
| **3. Demo Video** | **[👉 LINK TO YOUTUBE/LOOM DEMO (Replace with your link) 👈]()** |
| **4. Working Code** | Code is contained in this repository. Run `streamlit run streamlit_app.py` or explore the [Jupyter Notebook](./email_summarizer_notebook.ipynb) |

---

## 🛑 The Problem Statement
In modern corporate environments, professionals waste countless hours scrolling through massive, unorganised email chains simply to answer basic questions:
* *What is the conclusion of this discussion?*
* *Are there action items required?*
* *Is this conversation escalating negatively?*
* *Who is supposed to execute the next step?*

**The Solution:**
This application ingests unstructured data dumps (utilising the 1.4GB Kaggle Enron Dataset), mathematically groups conversations by subject lineage, and pushes them through a completely local, open-source Machine Learning pipeline to extract the answers. 

---

## 📂 Project Structure & File Dictionary

This codebase follows professional Python standards, segregating the UI from the engineering pipelines.

```text
email_summarizer/
│
├── main.py                          ← CLI version: run the pipeline from terminal
├── streamlit_app.py                 ← Frontend: The interactive web dashboard
├── email_summarizer_notebook.ipynb  ← Research: Jupyter Notebook walkthrough
├── requirements.txt                 ← Dependencies
├── README.md                        ← You are here
├── ARCHITECTURE.md                  ← Detailed explanation of the NLP engine
├── DEMO_SCRIPT.md                   ← Guideline for recording the demo video
│
├── utils/                           ← Core Logic Modules
│   ├── __init__.py
│   ├── email_loader.py              ← Ingests CSVs, isolates headers, creates threads
│   ├── nlp_engine.py                ← Houses the heavy lifting: Sumy, KeyBERT, VADER, SpaCy
│   └── excel_exporter.py            ← Builder for the formatted Excel analytical reports
│
└── data/                            ← Database Folder
    ├── enron_sample_1.csv           ← Processable 50-email slice from Kaggle
    ├── enron_sample_2.csv           ← Processable 50-email slice from Kaggle
    ├── enron_sample_3.csv           ← Processable 50-email slice from Kaggle
    └── results.csv                  ← Export location of the final machine-readable data
```

---

## ⚡ Key Features
1. **Extractive Summarisation**: Condenses 50-email threads into 3-sentence overviews.
2. **Sentiment Radar**: Flags "Urgent" or "Negative" operations.
3. **Owner Identification**: Autonomously finds the most relevant person associated with tasks using NLP Named Entity Recognition.
4. **Action Item Tracking**: Deploys Regex intelligence to surface explicit requests and mandates.
5. **No-Setup Demonstration**: The web dashboard comes pre-loaded with 'Real Enron Snippets', allowing managers or recruiters to instantly test the system's live capabilities without needing to upload huge files.

---

## 💻 Setup and Installation

### 1. Install Dependencies
```bash
# Install required libraries
pip install -r requirements.txt

# Download the required spaCy language module
python -m spacy download en_core_web_sm
```

### 2. Launching the App
The primary interface is the Streamlit Web Application. Run the following command in your terminal:
```bash
streamlit run streamlit_app.py
```
*Navigate to `http://localhost:8501` to view your data.*

### 3. Usage & Environments
- **Streamlit Demo Mode**: Navigate the sidebar to test out "Demo Scenarios" or live "Real Enron Snippets".
- **Jupyter Build Process**: If you wish to see how the code is structured block-by-block, open `email_summarizer_notebook.ipynb`.

---
*Developed as an AI-powered architecture assessment.*
