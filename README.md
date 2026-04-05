# 📧 Group Email Summarizer (Task Option 2)

> A single dashboard to make sense of various conversations in an email group, including Tasks, Followups, Threads, and Topic Insights. Uses the **Enron Email Dataset**.

---

## 🎯 Assignment Deliverables

| Deliverable | Verification / Link |
| :--- | :--- |
| **1. Architecture Note** | Full mapping and logic is housed in **[ARCHITECTURE.md](./ARCHITECTURE.md)** |
| **2. Public Dashboard Link** | [👉 **Live Streamlit App** 👈](https://group-email-summarizer-j7fuvqzryd9tdrgkzcrzlv.streamlit.app/)|
| **3. Demo Video** | [👉 **Google Drive Demo** 👈](https://drive.google.com/file/d/1IizsskffZ8N8K9d-NpYmnCGuzD4aJYui/view?usp=sharing) |
|                   | [👉 **Direct GitHub Video** 👈](https://github.com/daksha2005/Group-Email-Summarizer/blob/main/Group%20Email%20Summarizer.mp4) |

| **4. Working Code** | Code is contained in this repository. Run `streamlit run streamlit_app.py` or explore the [Jupyter Notebook](./email_summarizer_notebook.ipynb) |
| **5. Computed Table Data** | **[👉 View the final CSV Table here natively](./data/results.csv)**. Alternatively, download the **[👉 Excel Dashboard (email_dashboard.xlsx)](./email_dashboard.xlsx)** to view the exact same data with advanced spreadsheet styling! |

---

## 🛑 The Problem Statement (Task Option 2)
In a company, there are many email groups running parallel threads. The task is to create a single dashboard to make sense of the various conversations in that group, specifically handling:
*   Reading the incoming emails of an email group
*   Creating a summary of the email threads
*   Creating a structured table containing: **Email Thread, Key Topic, Action Items, Owner**

**The Solution:**
This application seamlessly fulfills all requirements. It ingests unstructured email group dumps (from the Enron Dataset), mathematically groups parallel threads, and pushes them through a completely local, open-source Machine Learning pipeline to extract the exact requested columns into a unified dashboard.

---

## 📂 Project Structure

```text
email_summarizer/
│
├── main.py                          ← CLI version: run the pipeline from terminal
├── streamlit_app.py                 ← Frontend: The interactive web dashboard
├── email_summarizer_notebook.ipynb  ← Research: Jupyter Notebook walkthrough
├── requirements.txt                 ← Dependencies
├── README.md                        ← You are here
├── ARCHITECTURE.md                  ← Detailed explanation of the NLP tools used
│
├── utils/                           ← Core Logic Modules
│   ├── email_loader.py              ← Reads incoming emails of a group & creates threads
│   ├── nlp_engine.py                ← Houses the intelligence: Sumy, KeyBERT, VADER, SpaCy
│   └── excel_exporter.py            ← Builder for the formatted Excel analytical reports
│
└── data/                            ← Database Folder
    ├── enron_sample_1.csv           ← Processable 50-email dataset directly from Enron
    └── results.csv                  ← The final machine-readable table mapped to requirements
```

---

## ⚡ Task Features & Allowed Tools Used
1. **Create summary of threads**: Built using `sumy` (Extractive NLP Summarisation).
2. **Key Topic**: Extracted using the advanced AI library `KeyBERT`.
3. **Action Items & Tasks**: Isolated using complex continuous Regular Expressions (Regex).
4. **Owner**: Handled by `spaCy` Named Entity Recognition (NER) to find the primary responsible person.
5. **Single Dashboard**: The Streamlit Web Application provides an interactive space to view all components at once without needing an API.

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
The primary interface is the Streamlit Dashboard. Run the following command in your terminal:
```bash
streamlit run streamlit_app.py
```
*Navigate to `http://localhost:8501` to view your data.*
