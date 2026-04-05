# 🎬 Demo Video Recording Script

This script is highly structured to be exactly ~3 minutes long. It balances technical explanation with an impressive visual showcase of the tool.

### Prep Checklist 
- [ ] Have the Streamlit Dashboard loaded and running smoothly in your browser.
- [ ] Have your code IDE open in the background just in case.
- [ ] Have the `ARCHITECTURE.md` file or diagram open in a different tab.
- [ ] Take a deep breath! You've got this.

---

## 0:00 - 0:30 | The Introduction & Problem Statement
*Action: Start recording with your webcam on, looking directly into the camera or the Streamlit Dashboard.*

**Dialogue:**
"Hi everyone. Today I'm demonstrating the AI-Powered Group Email Synthesiser. The core problem this system solves is the massive operational bloat caused by long, unstructured email chains in corporate environments. Professionals waste hours manually scrolling through endless replies just to find out what happened, what the action items are, and who is responsible for them."

## 0:30 - 0:50 | The Architecture Setup
*Action: Switch your screen to show the `ARCHITECTURE.md` file (Specifically the system diagram).*

**Dialogue:**
"To solve this, I architected a completely local NLP pipeline. Using the infamous Enron Email dataset, the system parses thousands of RFC-2822 emails, cleans the text, and groups them dynamically into business threads. From there, it feeds them through an AI engine using `sumy` for extractive summaries, `KeyBERT` for topic modelling, `VADER` for sentiment analysis, and `spaCy` to extract task owners."

## 0:50 - 2:00 | The Streamlit Web Application
*Action: Swap over to the live Streamlit Dashboard. Keep the window on the main screen displaying the KPI cards and the charts.*

**Dialogue:**
"Let me show you the live dashboard. As you can see, when data is processed, it immediately generates clear KPIs—showing exactly how many threads have actionable items and which ones are marked structurally as Urgent."

*Action: Scroll down slightly to show the 'Thread Intelligence Table'.*
"The intelligence table converts unstructured chaos into exact data. Let's look closer at one of those..."

*Action: Click to expand one of the threads securely down in the 'Thread Details' section.*
"If I expand a thread, you can see how my system has taken multiple, multi-page emails and distilled them into a clean 3-sentence summary. It has also used Regex and Named Entity Recognition to pull out specific Action Items, Follow-ups, and the suspected Owner—all happening locally with zero external API calls."

## 2:00 - 2:40 | The "Real Enron Snippet" Demo (The Big Moment)
*Action: Move your mouse over to the Sidebar on the left. Click on 'Real Enron Snippets'.*

**Dialogue:**
"One of the best features I added for reviewers is the Built-in Sample Engine. By switching my data source here, to a 'Real Enron Snippet', the dashboard instantly reads a 50-email chunk taken cleanly right out of the actual Kaggle dataset."

*Action: Pick 'Snippet 2' or 'Snippet 3' from the dropdown menu and watch the app rebuild the dashboard instantly.*
"Watch as the system processes real, raw Enron data on the fly. It rebuilds the dataset, generates sentiment metrics, and uncovers real historical events from the database instantly!"

## 2:40 - 3:00 | Conclusion & Wrap Up
*Action: Scroll down to the bottom and click the 'Download Excel Dashboard' button. Open the Excel file briefly on the screen to show it.*

**Dialogue:**
"Finally, since business runs on spreadsheets, the system provides a pristine, styled Excel Dashboard generation. All pipeline logic, fallback parameters, and error handlers are fully documented in the source code. Thank you so much for your time, and I look forward to your feedback!"
