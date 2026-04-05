from __future__ import annotations
"""
main.py
-------
CLI entry point for the Group Email Summarizer.

Modes:
  python main.py                    → use built-in sample data (demo)
  python main.py --csv data/emails.csv  → use real Enron CSV
  python main.py --csv data/emails.csv --rows 100000 --top 60

Output:
  email_dashboard.xlsx
  data/results.csv          (machine-readable)
"""

import argparse
import sys
import pandas as pd
from pathlib import Path

# ── Local imports ─────────────────────────────────────────────────────────────
from utils.email_loader  import load_emails, group_into_threads, get_sample_threads
from utils.nlp_engine    import analyse_thread
from utils.excel_exporter import export_to_excel


BANNER = """
╔══════════════════════════════════════════════════════════════╗
║        GROUP EMAIL SUMMARIZER — Enron Dataset               ║
║        NLP Stack: sumy · KeyBERT · VADER · spaCy            ║
╚══════════════════════════════════════════════════════════════╝
"""


# ─────────────────────────────────────────────────────────────────────────────
# PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def run_pipeline(
    csv_path: str | None,
    nrows: int = 80_000,
    top_n: int = 50,
    output_excel: str = "email_dashboard.xlsx",
    output_csv: str = "data/results.csv",
) -> pd.DataFrame:
    """
    Full pipeline: load → group → analyse → export.
    Returns the results DataFrame.
    """
    print(BANNER)

    # ── Step 1: Load & group threads ─────────────────────────────────────────
    if csv_path and Path(csv_path).exists():
        print(f"[1/4] Loading Enron CSV: {csv_path}")
        email_df = load_emails(csv_path=csv_path, nrows=nrows)
        threads  = group_into_threads(email_df, min_emails=2, top_n=top_n)
        print(f"      → {len(threads)} threads selected for analysis")
    else:
        print("[1/4] No CSV provided — using built-in sample threads (demo mode)")
        threads = get_sample_threads()
        print(f"      → {len(threads)} sample threads loaded")

    # ── Step 2: NLP analysis ─────────────────────────────────────────────────
    print(f"\n[2/4] Analysing {len(threads)} threads with NLP pipeline…")
    results = []
    total   = len(threads)
    for i, (tkey, emails) in enumerate(threads.items(), 1):
        print(f"      [{i:02d}/{total}] {tkey[:55]}", end="\r")
        try:
            result = analyse_thread(tkey, emails)
            results.append(result)
        except Exception as e:
            print(f"\n      ⚠ Error on '{tkey[:40]}': {e}")

    print(f"\n      ✓ Analysis complete — {len(results)} threads processed")

    # ── Step 3: Build DataFrame ───────────────────────────────────────────────
    print("\n[3/4] Building results table…")
    df = pd.DataFrame(results)
    print(df[["Email Thread", "Key Topic", "Owner", "Sentiment", "Email Count"]].to_string(index=False))

    # Save CSV
    Path(output_csv).parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_csv, index=False)
    print(f"\n      ✓ Results CSV saved → {output_csv}")

    # ── Step 4: Export Excel ──────────────────────────────────────────────────
    print("\n[4/4] Writing Excel dashboard…")
    export_to_excel(df, output_path=output_excel)

    # ── Summary ───────────────────────────────────────────────────────────────
    sent_counts = df["Sentiment"].value_counts().to_dict()
    print("\n" + "═" * 62)
    print("  PIPELINE COMPLETE")
    print("═" * 62)
    print(f"  Threads analysed   : {len(df)}")
    print(f"  Total emails       : {int(df['Email Count'].sum())}")
    for s in ("urgent", "negative", "positive", "neutral"):
        n = sent_counts.get(s, 0)
        bar = "█" * n
        print(f"  {s.capitalize():<14}    : {n:>3}  {bar}")
    print(f"\n  📊 Excel  → {output_excel}")
    print(f"  📄 CSV    → {output_csv}")
    print("═" * 62)

    return df


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def _parse_args():
    p = argparse.ArgumentParser(
        description="Group Email Summarizer — Enron Dataset NLP Pipeline"
    )
    p.add_argument("--csv",    default=None,
                   help="Path to Enron emails.csv (omit to use sample data)")
    p.add_argument("--rows",   type=int, default=80_000,
                   help="Max rows to read from CSV (default: 80000)")
    p.add_argument("--top",    type=int, default=50,
                   help="Number of top threads to analyse (default: 50)")
    p.add_argument("--output", default="email_dashboard.xlsx",
                   help="Output Excel path (default: email_dashboard.xlsx)")
    return p.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    run_pipeline(
        csv_path=args.csv,
        nrows=args.rows,
        top_n=args.top,
        output_excel=args.output,
    )
