"""
streamlit_app.py
----------------
Impressive Streamlit dashboard for the Group Email Summarizer.

Run with:
    streamlit run streamlit_app.py

Features:
  ▸ KPI metric cards (total threads, urgent, with action items, email count)
  ▸ Interactive sentiment donut chart + thread size bar chart (Plotly)
  ▸ Searchable, filterable full thread table
  ▸ Thread detail expander with summary, action items, follow-ups, participants
  ▸ Live CSV upload to analyse your own Enron data slice
  ▸ Excel download button
  ▸ Dark-navy branded theme
"""
from __future__ import annotations

import io
import time
import pandas as pd
import streamlit as st

# ── Page config — MUST be first Streamlit call ────────────────────────────────
st.set_page_config(
    page_title="Email Group Intelligence",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Local imports ─────────────────────────────────────────────────────────────
from utils.email_loader   import load_emails, group_into_threads, get_sample_threads
from utils.nlp_engine     import analyse_thread
from utils.excel_exporter import export_to_excel

# ── Plotly (optional — degrades gracefully) ───────────────────────────────────
try:
    import plotly.graph_objects as go
    import plotly.express as px
    PLOTLY = True
except ImportError:
    PLOTLY = False


# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL CSS
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
/* ── Base ───────────────────────────────────────────────────────────────── */
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif !important;
}

/* ── Sidebar ─────────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: #0d1829 !important;
}
[data-testid="stSidebar"] * {
    color: #dde6f0 !important;
}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stTextInput label {
    color: #5a7a99 !important;
    font-size: 11px !important;
    text-transform: uppercase;
    letter-spacing: .5px;
}

/* ── KPI Cards ───────────────────────────────────────────────────────────── */
.kpi-card {
    background: #162032;
    border: 1px solid #2a3d55;
    border-radius: 10px;
    padding: 18px 20px;
    text-align: center;
}
.kpi-val {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 38px;
    font-weight: 600;
    line-height: 1;
}
.kpi-lbl {
    font-size: 11px;
    color: #5a7a99;
    text-transform: uppercase;
    letter-spacing: .5px;
    margin-top: 4px;
}

/* ── Thread table ────────────────────────────────────────────────────────── */
.thread-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12.5px;
}
.thread-table th {
    background: #1f4e79;
    color: #e8f4fd;
    padding: 10px 14px;
    text-align: left;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: .5px;
    white-space: nowrap;
    border-bottom: 2px solid #38bdf8;
}
.thread-table td {
    padding: 11px 14px;
    vertical-align: top;
    border-bottom: 1px solid #2a3d55;
    line-height: 1.5;
    color: #dde6f0;
}
.thread-table tr:hover td { background: #1e2d42; }

/* ── Sentiment badges ────────────────────────────────────────────────────── */
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: .4px;
    text-transform: uppercase;
}
.badge-urgent   { background: #f9731622; color: #f97316; }
.badge-negative { background: #ef444422; color: #ef4444; }
.badge-positive { background: #22c55e22; color: #22c55e; }
.badge-neutral  { background: #94a3b822; color: #94a3b8; }

/* ── Detail panel ────────────────────────────────────────────────────────── */
.detail-box {
    background: #162032;
    border: 1px solid #2a3d55;
    border-radius: 8px;
    padding: 16px 20px;
    margin: 8px 0;
}
.detail-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    color: #38bdf8;
    text-transform: uppercase;
    letter-spacing: .5px;
    margin-bottom: 8px;
}
.action-item {
    background: #1e2d42;
    border-left: 3px solid #818cf8;
    padding: 6px 12px;
    margin: 4px 0;
    border-radius: 0 4px 4px 0;
    font-size: 12px;
    color: #dde6f0;
}
.followup-item {
    background: #1e2d42;
    border-left: 3px solid #22c55e;
    padding: 6px 12px;
    margin: 4px 0;
    border-radius: 0 4px 4px 0;
    font-size: 12px;
    color: #dde6f0;
}

/* ── Header ──────────────────────────────────────────────────────────────── */
.app-header {
    background: linear-gradient(135deg, #0d1829 0%, #1f4e79 100%);
    border-radius: 12px;
    padding: 28px 32px;
    margin-bottom: 24px;
    border: 1px solid #2a3d55;
}
.app-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 24px;
    color: #38bdf8;
    margin: 0;
    letter-spacing: -0.5px;
}
.app-sub {
    font-size: 13px;
    color: #5a7a99;
    margin-top: 6px;
}

/* ── Section headers ─────────────────────────────────────────────────────── */
.section-hdr {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 12px;
    color: #38bdf8;
    text-transform: uppercase;
    letter-spacing: .5px;
    border-bottom: 1px solid #2a3d55;
    padding-bottom: 8px;
    margin: 24px 0 16px;
}

/* ── Hide Streamlit branding ──────────────────────────────────────────────── */
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE — CACHE RESULTS
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def get_sample_data() -> pd.DataFrame:
    """Load and analyse the built-in sample threads (cached)."""
    threads = get_sample_threads()
    results = [analyse_thread(k, v) for k, v in threads.items()]
    return pd.DataFrame(results)


@st.cache_data(show_spinner=False)
def analyse_csv(csv_bytes: bytes, nrows: int, top_n: int) -> pd.DataFrame:
    """Parse and analyse a user-uploaded Enron CSV (cached by content hash)."""
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
        f.write(csv_bytes)
        tmp_path = f.name
    try:
        df_emails = load_emails(csv_path=tmp_path, nrows=nrows)
        threads   = group_into_threads(df_emails, top_n=top_n)
        results   = []
        for k, v in threads.items():
            results.append(analyse_thread(k, v))
        return pd.DataFrame(results)
    finally:
        os.unlink(tmp_path)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("""
    <div style="font-family:'IBM Plex Mono',monospace;font-size:16px;
                color:#38bdf8;padding:12px 0 20px;border-bottom:1px solid #2a3d55;
                margin-bottom:20px">
        // EMAIL INTEL
    </div>
    """, unsafe_allow_html=True)

    st.markdown("**DATA SOURCE**")
    data_mode = st.radio(
        "Choose data source",
        ["Demo Scenarios (Instant)", "Real Enron Snippets", "Upload Custom CSV"],
        label_visibility="collapsed"
    )

    uploaded_file = None
    snippet_choice = None
    nrows_sel = 80_000
    top_n_sel = 30


    if data_mode == "Upload Custom CSV":
        uploaded_file = st.file_uploader(
            "Upload emails.csv",
            type=["csv"],
            help="Download from: kaggle.com/datasets/wcukierski/enron-email-dataset"
        )
        nrows_sel = st.slider("Rows to read", 10_000, 200_000, 80_000, 10_000)
        top_n_sel = st.slider("Top N threads", 10, 100, 30, 5)
    elif data_mode == "Real Enron Snippets":
        snippet_choice = st.selectbox(
            "Select Built-in Sample",
            [
                "Snippet 1 (50 real emails)",
                "Snippet 2 (50 real emails)",
                "Snippet 3 (50 real emails)",
                "Snippet 4 (50 real emails)",
                "Snippet 5 (50 real emails)"
            ]
        )

    st.markdown("---")
    st.markdown("**FILTERS**")
    sentiment_filter = st.multiselect(
        "Sentiment",
        ["urgent", "negative", "neutral", "positive"],
        default=["urgent", "negative", "neutral", "positive"]
    )
    search_query = st.text_input("Search threads / topics / owners", "")

    st.markdown("---")
    st.markdown("""
    <div style="font-size:10px;color:#3d5a70;line-height:1.8">
    NLP Stack<br>
    · sumy (extractive summary)<br>
    · KeyBERT (topic keyphrases)<br>
    · VADER (sentiment)<br>
    · spaCy NER (owner)<br>
    · Regex (action items)<br><br>
    Dataset: Enron Email (Kaggle)<br>
    No API keys required.
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def analyse_snippet(snippet_idx: int) -> pd.DataFrame:
    """Parse and analyse a built-in snippet CSV."""
    path = f"data/enron_sample_{snippet_idx}.csv"
    df_emails = load_emails(csv_path=path, nrows=80_000)
    threads   = group_into_threads(df_emails, top_n=30)
    results   = []
    for k, v in threads.items():
        results.append(analyse_thread(k, v))
    return pd.DataFrame(results)

if data_mode == "Upload Custom CSV" and uploaded_file:
    with st.spinner("📧 Parsing emails and running NLP analysis…"):
        try:
            df_raw = analyse_csv(uploaded_file.read(), nrows_sel, top_n_sel)
            st.success(f"✓ Analysed {len(df_raw)} threads from uploaded CSV")
        except Exception as e:
            st.error(f"Failed to parse uploaded file.\\n\\nError: {e}")
            st.stop()
elif data_mode == "Real Enron Snippets" and snippet_choice:
    # Extract the index from string (e.g. "Snippet 1 (50 real emails)" -> 1)
    idx = int(snippet_choice.split(" ")[1])
    with st.spinner(f"📧 Running NLP pipeline on Snippet {idx}…"):
        df_raw = analyse_snippet(idx)
        st.success(f"✓ Analysed pure Enron threads from Snippet {idx}")
else:
    with st.spinner("Loading demo scenarios…"):
        df_raw = get_sample_data()


# ─────────────────────────────────────────────────────────────────────────────
# APPLY FILTERS
# ─────────────────────────────────────────────────────────────────────────────

df = df_raw.copy()
if sentiment_filter:
    df = df[df["Sentiment"].isin(sentiment_filter)]
if search_query:
    q = search_query.lower()
    mask = (
        df["Email Thread"].str.lower().str.contains(q, na=False) |
        df["Key Topic"].str.lower().str.contains(q, na=False) |
        df["Owner"].str.lower().str.contains(q, na=False) |
        df["Action Items"].str.lower().str.contains(q, na=False) |
        df["Summary"].str.lower().str.contains(q, na=False)
    )
    df = df[mask]


# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="app-header">
  <div class="app-title">// GROUP EMAIL INTELLIGENCE DASHBOARD</div>
  <div class="app-sub">
    Enron Email Dataset &nbsp;·&nbsp; NLP: sumy · KeyBERT · VADER · spaCy &nbsp;·&nbsp;
    Zero external API calls &nbsp;·&nbsp; Real-time thread analysis
  </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# KPI CARDS
# ─────────────────────────────────────────────────────────────────────────────

sent_counts  = df_raw["Sentiment"].value_counts().to_dict()
action_count = int((df_raw["Action Items"] != "None identified").sum())
urgent_count = int(sent_counts.get("urgent", 0))
total_emails = int(df_raw["Email Count"].sum())

k1, k2, k3, k4, k5 = st.columns(5)

with k1:
    st.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-val" style="color:#38bdf8">{len(df_raw)}</div>
      <div class="kpi-lbl">Threads Analysed</div>
    </div>""", unsafe_allow_html=True)

with k2:
    st.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-val" style="color:#818cf8">{total_emails}</div>
      <div class="kpi-lbl">Total Emails</div>
    </div>""", unsafe_allow_html=True)

with k3:
    st.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-val" style="color:#22c55e">{action_count}</div>
      <div class="kpi-lbl">Have Action Items</div>
    </div>""", unsafe_allow_html=True)

with k4:
    st.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-val" style="color:#f97316">{urgent_count}</div>
      <div class="kpi-lbl">Urgent Threads</div>
    </div>""", unsafe_allow_html=True)

with k5:
    neg = int(sent_counts.get("negative", 0))
    st.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-val" style="color:#ef4444">{neg}</div>
      <div class="kpi-lbl">Negative Threads</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# CHARTS
# ─────────────────────────────────────────────────────────────────────────────

st.markdown('<div class="section-hdr">Analytics</div>', unsafe_allow_html=True)

if PLOTLY:
    ch1, ch2 = st.columns([1, 2])

    with ch1:
        SC = {"urgent":"#f97316","negative":"#ef4444","positive":"#22c55e","neutral":"#94a3b8"}
        labels = list(sent_counts.keys())
        values = [sent_counts[l] for l in labels]
        colors = [SC.get(l, "#94a3b8") for l in labels]

        fig_donut = go.Figure(go.Pie(
            labels=[l.capitalize() for l in labels],
            values=values,
            hole=0.65,
            marker_colors=colors,
            textinfo="percent+label",
            textfont_size=11,
        ))
        fig_donut.update_layout(
            title=dict(text="Sentiment Split", font=dict(size=13, color="#38bdf8")),
            showlegend=False,
            margin=dict(t=40, b=20, l=10, r=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#dde6f0"),
            height=260,
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    with ch2:
        sizes   = df_raw["Email Count"].sort_values(ascending=False).values
        threads = [df_raw.iloc[i]["Email Thread"][:25] for i in df_raw["Email Count"].argsort()[::-1]]

        fig_bar = go.Figure(go.Bar(
            x=threads,
            y=sizes,
            marker_color="#38bdf8",
            marker_opacity=0.7,
            marker_line_color="#38bdf8",
            marker_line_width=0.5,
        ))
        fig_bar.update_layout(
            title=dict(text="Thread Size (emails per thread)", font=dict(size=13, color="#38bdf8")),
            xaxis=dict(tickangle=-30, tickfont=dict(size=9, color="#5a7a99"), showgrid=False),
            yaxis=dict(tickfont=dict(size=9, color="#5a7a99"), gridcolor="#2a3d55"),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=40, b=60, l=40, r=10),
            height=260,
        )
        st.plotly_chart(fig_bar, use_container_width=True)
else:
    # Fallback: show as text metrics
    st.write({k: v for k, v in sent_counts.items()})


# ─────────────────────────────────────────────────────────────────────────────
# MAIN TABLE
# ─────────────────────────────────────────────────────────────────────────────

st.markdown('<div class="section-hdr">Thread Intelligence Table</div>', unsafe_allow_html=True)

# Filter status line
filter_info = f"Showing **{len(df)}** of **{len(df_raw)}** threads"
if search_query:
    filter_info += f" matching *\"{search_query}\"*"
st.markdown(filter_info)


def sentiment_badge(s: str) -> str:
    cls = f"badge-{s.lower()}" if s.lower() in ("urgent","negative","positive","neutral") else "badge-neutral"
    return f'<span class="badge {cls}">{s}</span>'


if df.empty:
    st.warning("No threads match your current filters. Try adjusting the sidebar.")
else:
    # Build HTML table
    rows_html = ""
    for _, row in df.iterrows():
        actions_raw = str(row.get("Action Items", ""))
        actions = actions_raw.split(" | ") if actions_raw and actions_raw != "None identified" else []
        ai_html = "".join(f"<div style='margin:2px 0;font-size:11px'>• {a}</div>" for a in actions[:3]) \
                  or "<span style='color:#3d5a70;font-size:11px'>None</span>"

        rows_html += f"""
        <tr>
          <td>
            <strong style="color:#e2e8f0">{row.get('Email Thread','')}</strong>
            <div style="font-size:10px;color:#5a7a99;margin-top:2px">
              {row.get('Email Count',0)} emails &nbsp;·&nbsp; {row.get('Latest Date','')}
            </div>
          </td>
          <td style="color:#94a3b8;font-size:11px">{row.get('Key Topic','')}</td>
          <td>{ai_html}</td>
          <td style="color:#818cf8;font-size:11px">{row.get('Owner','')}</td>
          <td>{sentiment_badge(str(row.get('Sentiment','neutral')))}</td>
        </tr>"""

    table_html = f"""
    <div style="overflow-x:auto">
    <table class="thread-table">
      <thead>
        <tr>
          <th style="min-width:180px">Email Thread</th>
          <th style="min-width:180px">Key Topic</th>
          <th style="min-width:220px">Action Items</th>
          <th style="min-width:140px">Owner</th>
          <th style="min-width:100px">Sentiment</th>
        </tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>
    </div>"""

    st.markdown(table_html, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# THREAD DETAIL EXPANDERS
# ─────────────────────────────────────────────────────────────────────────────

st.markdown('<div class="section-hdr">Thread Details</div>', unsafe_allow_html=True)
st.caption("Expand any thread to see full summary, action items, follow-ups, and participants.")

for _, row in df.iterrows():
    sent = str(row.get("Sentiment", "neutral")).lower()
    icon = {"urgent":"🔴","negative":"🟠","positive":"🟢","neutral":"⚪"}.get(sent, "⚪")
    label = f"{icon}  {row.get('Email Thread','')}  —  {row.get('Email Count',0)} emails  ·  {row.get('Latest Date','')}"

    with st.expander(label, expanded=False):
        d1, d2 = st.columns([3, 1])

        with d1:
            # Summary
            st.markdown('<div class="detail-title">Summary</div>', unsafe_allow_html=True)
            summary = str(row.get("Summary", "No summary available."))
            st.markdown(f'<div style="font-size:13px;color:#b0c4d8;line-height:1.7">{summary}</div>',
                        unsafe_allow_html=True)

            # Action items
            st.markdown('<div class="detail-title" style="margin-top:16px">Action Items</div>',
                        unsafe_allow_html=True)
            ai_list = row.get("Action List", [])
            if isinstance(ai_list, list) and ai_list:
                for item in ai_list:
                    st.markdown(f'<div class="action-item">▸ {item}</div>', unsafe_allow_html=True)
            else:
                actions_raw = str(row.get("Action Items", ""))
                if actions_raw and actions_raw not in ("None identified", "nan"):
                    for item in actions_raw.split(" | "):
                        if item.strip():
                            st.markdown(f'<div class="action-item">▸ {item.strip()}</div>',
                                        unsafe_allow_html=True)
                else:
                    st.markdown('<div style="color:#3d5a70;font-size:12px">No action items identified</div>',
                                unsafe_allow_html=True)

            # Follow-ups
            st.markdown('<div class="detail-title" style="margin-top:16px">Follow-ups</div>',
                        unsafe_allow_html=True)
            fu_list = row.get("Follow-up List", [])
            followups_raw = str(row.get("Follow-ups", ""))
            fu_items = fu_list if (isinstance(fu_list, list) and fu_list) else \
                       [f for f in followups_raw.split(" | ") if f.strip() and f != "None"]

            if fu_items:
                for item in fu_items:
                    st.markdown(f'<div class="followup-item">↪ {item.strip()}</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div style="color:#3d5a70;font-size:12px">No follow-ups identified</div>',
                            unsafe_allow_html=True)

        with d2:
            # Metadata card
            st.markdown(f"""
            <div class="detail-box">
              <div class="detail-title">Thread Info</div>
              <div style="font-size:11px;color:#5a7a99;margin-bottom:4px">Owner</div>
              <div style="font-size:13px;color:#818cf8;margin-bottom:12px">{row.get('Owner','—')}</div>
              <div style="font-size:11px;color:#5a7a99;margin-bottom:4px">Sentiment</div>
              <div style="margin-bottom:12px">{sentiment_badge(str(row.get('Sentiment','neutral')))}</div>
              <div style="font-size:11px;color:#5a7a99;margin-bottom:4px">Latest Email</div>
              <div style="font-size:12px;color:#dde6f0;margin-bottom:12px">{row.get('Latest Date','—')}</div>
              <div style="font-size:11px;color:#5a7a99;margin-bottom:4px">Email Count</div>
              <div style="font-size:22px;color:#38bdf8;font-family:'IBM Plex Mono',monospace;font-weight:600">
                {row.get('Email Count',0)}
              </div>
            </div>
            """, unsafe_allow_html=True)

            # Participants
            participants = str(row.get("Participants", ""))
            if participants and participants not in ("", "nan"):
                st.markdown(f"""
                <div class="detail-box" style="margin-top:8px">
                  <div class="detail-title">Participants</div>
                  {"".join(f'<div style="font-size:10px;color:#5a7a99;margin:2px 0">✉ {p.strip()}</div>'
                            for p in participants.split(",") if p.strip())}
                </div>
                """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# DOWNLOAD BUTTONS
# ─────────────────────────────────────────────────────────────────────────────

st.markdown('<div class="section-hdr">Export</div>', unsafe_allow_html=True)

dl1, dl2 = st.columns(2)

with dl1:
    csv_buffer = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️  Download Results CSV",
        data=csv_buffer,
        file_name="email_intelligence_results.csv",
        mime="text/csv",
        use_container_width=True,
    )

with dl2:
    # Build Excel in memory
    excel_buffer = io.BytesIO()
    try:
        from openpyxl import Workbook
        export_to_excel(df_raw, output_path="/tmp/email_dashboard.xlsx")
        with open("/tmp/email_dashboard.xlsx", "rb") as f:
            excel_bytes = f.read()
        st.download_button(
            label="⬇️  Download Excel Dashboard",
            data=excel_bytes,
            file_name="email_dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception:
        st.info("Install openpyxl to enable Excel export: pip install openpyxl")


# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<div style="margin-top:48px;padding-top:20px;border-top:1px solid #2a3d55;
            text-align:center;font-size:10px;color:#3d5a70;">
  Group Email Intelligence Dashboard &nbsp;·&nbsp;
  Enron Email Dataset (Kaggle) &nbsp;·&nbsp;
  NLP: sumy · KeyBERT · VADER · spaCy &nbsp;·&nbsp;
  No external API keys required
</div>
""", unsafe_allow_html=True)
