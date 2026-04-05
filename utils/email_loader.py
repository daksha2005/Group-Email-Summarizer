from __future__ import annotations
"""
utils/email_loader.py
---------------------
Load the Enron Email Dataset CSV (or a folder of raw .txt files),
parse RFC-2822 email headers/body, and group messages into threads
by normalising the Subject line.

Enron CSV columns: file | message
Each 'message' is a raw RFC-2822 email string.
"""

import re
import email as email_lib
import pandas as pd
from pathlib import Path
from collections import defaultdict


# ── Supported file paths ──────────────────────────────────────────────────────
ENRON_CSV_DEFAULT = "data/emails.csv"


# ─────────────────────────────────────────────────────────────────────────────
# PARSING
# ─────────────────────────────────────────────────────────────────────────────

def _parse_single_email(raw: str) -> dict:
    """
    Parse a raw RFC-2822 string into a clean dict.
    Strips quoted reply blocks (lines starting with '>') to keep
    only the original message body.
    """
    msg = email_lib.message_from_string(raw)

    # Extract body
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                payload = part.get_payload(decode=True)
                body = payload.decode("utf-8", errors="ignore") if payload else ""
                break
    else:
        payload = msg.get_payload(decode=True)
        if isinstance(payload, bytes):
            body = payload.decode("utf-8", errors="ignore")
        else:
            body = msg.get_payload() or ""

    # Strip quoted blocks and forwarded headers
    body = re.sub(r"(-{3,}.*|_{3,}.*)", "", body, flags=re.DOTALL)
    body = re.sub(r"^>.*$", "", body, flags=re.MULTILINE)
    body = re.sub(r"\n{3,}", "\n\n", body).strip()

    return {
        "message_id": msg.get("Message-ID", "").strip(),
        "date":       msg.get("Date", ""),
        "from_":      msg.get("From", "").strip(),
        "to":         msg.get("To", "").strip(),
        "cc":         msg.get("Cc", "").strip(),
        "subject":    msg.get("Subject", "(no subject)").strip(),
        "body":       body[:3000],   # cap body length
    }


# ─────────────────────────────────────────────────────────────────────────────
# SUBJECT NORMALISATION → THREAD GROUPING
# ─────────────────────────────────────────────────────────────────────────────

def _normalise_subject(subject: str) -> str:
    """
    Strip Re:/Fwd:/Fw: prefixes, lower-case, collapse whitespace.
    This groups all replies into the same thread.
    """
    s = re.sub(r"^(re|fwd?|aw)[:\s]+", "", subject.strip(), flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s or "(no subject)"


# ─────────────────────────────────────────────────────────────────────────────
# PUBLIC API
# ─────────────────────────────────────────────────────────────────────────────

def load_emails(csv_path: str = ENRON_CSV_DEFAULT, nrows: int = 80_000) -> pd.DataFrame:
    """
    Load the Enron CSV and parse every email into a structured DataFrame.

    Args:
        csv_path : Path to emails.csv from the Enron dataset.
        nrows    : Max rows to read (limit for speed; full dataset = 517k rows).

    Returns:
        DataFrame with columns:
        message_id | date | from_ | to | cc | subject | body | thread_key
    """
    path = Path(csv_path)
    if not path.exists():
        raise FileNotFoundError(
            f"Enron CSV not found at: {path.resolve()}\n"
            "Download from: https://www.kaggle.com/datasets/wcukierski/enron-email-dataset\n"
            "Place emails.csv inside the data/ folder."
        )

    print(f"Loading {nrows:,} emails from {path.name}…")
    try:
        raw_df = pd.read_csv(path, nrows=nrows)
    except Exception:
        # Fallback for manually split/truncated files (handles 'EOF inside string')
        raw_df = pd.read_csv(path, nrows=nrows, engine="python", on_bad_lines="skip")

    print("Parsing email headers and bodies…")
    parsed = raw_df["message"].apply(_parse_single_email)
    df = pd.DataFrame(list(parsed))
    df["thread_key"] = df["subject"].apply(_normalise_subject)

    print(f"✓ Parsed {len(df):,} emails into {df['thread_key'].nunique():,} unique threads.")
    return df


def group_into_threads(
    df: pd.DataFrame,
    min_emails: int = 2,
    top_n: int = 50
) -> dict[str, list[dict]]:
    """
    Group emails into threads and return the top N busiest threads.

    Args:
        df         : Parsed email DataFrame from load_emails().
        min_emails : Minimum emails per thread to be included.
        top_n      : Maximum number of threads to return.

    Returns:
        Dict { thread_key -> list of email dicts }, sorted by size (desc).
    """
    counts = df["thread_key"].value_counts()
    active_keys = counts[counts >= min_emails].index

    bundles: dict[str, list[dict]] = defaultdict(list)
    for _, row in df[df["thread_key"].isin(active_keys)].iterrows():
        bundles[row["thread_key"]].append(row.to_dict())

    # Sort by thread size, take top N
    sorted_threads = sorted(bundles.items(), key=lambda x: len(x[1]), reverse=True)
    return dict(sorted_threads[:top_n])


def get_sample_threads() -> dict[str, list[dict]]:
    """
    Return 10 realistic hand-crafted sample threads that simulate
    the Enron dataset. Used as a fallback when the CSV is not available,
    so the app still runs for demo purposes.
    """
    from datetime import datetime, timedelta
    import random

    random.seed(42)
    base_date = datetime(2001, 9, 1)

    SAMPLES = [
        {
            "key": "california power crisis response",
            "emails": [
                {"from_": "tim.belden@enron.com", "subject": "California Power Crisis Response",
                 "body": "Team — urgent action needed. FERC is requesting all trading logs by Friday. Please compile every ISO-NE position we held in August. Jeff, can you lead this? We need to draft a regulatory response immediately. Action: compile trading logs by Thursday EOD. Owner: Jeff Dasovich."},
                {"from_": "jeff.dasovich@enron.com", "subject": "Re: California Power Crisis Response",
                 "body": "Understood. I'll coordinate with legal. We also need to loop in the government affairs team. Follow up: schedule call with FERC liaison by Wednesday. I'll have the preliminary report ready by Thursday."},
                {"from_": "richard.shapiro@enron.com", "subject": "Re: California Power Crisis Response",
                 "body": "Agreed. Also please note that the Senate committee wants written testimony. We should prepare talking points. Need to send the draft to lobbyists by Tuesday."},
                {"from_": "tim.belden@enron.com", "subject": "Re: California Power Crisis Response",
                 "body": "Escalating to Ken Lay for visibility. This is now a board-level concern. Action: prepare executive briefing pack. Deadline is Monday morning."},
            ]
        },
        {
            "key": "q4 gas contract renewal",
            "emails": [
                {"from_": "louise.kitchen@enron.com", "subject": "Q4 Gas Contract Renewal",
                 "body": "The counterparty has come back with revised pricing on the Q4 gas supply contract. We need to review the term sheet. Please review and send comments by Friday. Action: Louise to circulate revised term sheet. Owner: Louise Kitchen."},
                {"from_": "john.arnold@enron.com", "subject": "Re: Q4 Gas Contract Renewal",
                 "body": "Reviewed the pricing. We should push back on the $3.40/MMBtu floor. Suggest counter at $3.20. Also need legal to review the force majeure clause. Follow up: John to draft counter-proposal."},
                {"from_": "sara.shackleton@enron.com", "subject": "Re: Q4 Gas Contract Renewal",
                 "body": "Legal review complete. The force majeure clause needs revision — current language is too broad. I'll send a redline version today. Action: Sara to provide redlined agreement by EOD."},
            ]
        },
        {
            "key": "broadband unit wind-down",
            "emails": [
                {"from_": "jeff.skilling@enron.com", "subject": "Broadband Unit Wind-Down",
                 "body": "After careful consideration, we are proceeding with the wind-down of EBS. Key actions: (1) notify all customers by Oct 15, (2) cancel vendor contracts, (3) coordinate with HR on affected staff. This is confidential until announcement. Owner: Jeff Skilling."},
                {"from_": "cindy.olson@enron.com", "subject": "Re: Broadband Unit Wind-Down",
                 "body": "HR is ready to support. We have 312 employees in EBS. WARN Act notices need to go out 60 days before termination date. I'll prepare the communications plan. Action: Cindy to prepare WARN notices and severance schedule."},
                {"from_": "mark.koenig@enron.com", "subject": "Re: Broadband Unit Wind-Down",
                 "body": "Communications strategy: we plan to announce alongside Q3 earnings. Press release draft is being prepared. Need approval from legal and Jeff before release. Follow up: Mark to circulate draft press release for review."},
                {"from_": "james.derrick@enron.com", "subject": "Re: Broadband Unit Wind-Down",
                 "body": "Legal review needed on customer contract termination clauses. Some have early termination penalties. I'll have the assessment done by Friday. Action: Legal to review all 47 customer contracts."},
            ]
        },
        {
            "key": "raptor spv disclosure",
            "emails": [
                {"from_": "andrew.fastow@enron.com", "subject": "Raptor SPV Disclosure",
                 "body": "We need to finalise the 10-Q disclosure language for the Raptor vehicles. Andy Andersen has flagged concerns about the current draft. Meeting scheduled for Tuesday with the audit committee. Action: finalise disclosure language. Owner: Andrew Fastow. URGENT."},
                {"from_": "richard.causey@enron.com", "subject": "Re: Raptor SPV Disclosure",
                 "body": "I've reviewed the draft with Arthur Andersen. They want additional footnote disclosures on the hedging arrangements. This needs board approval before filing. Deadline: Friday before market open. Action: Rick Causey to present revised disclosures to audit committee Tuesday."},
                {"from_": "james.derrick@enron.com", "subject": "Re: Raptor SPV Disclosure",
                 "body": "Outside counsel has reviewed. The current structure may require restatement of prior quarters. This is a material finding — Ken Lay must be briefed immediately. Follow up: James to brief CEO and General Counsel by Monday."},
            ]
        },
        {
            "key": "ferc investigation response",
            "emails": [
                {"from_": "rex.rogers@enron.com", "subject": "FERC Investigation Response",
                 "body": "FERC has issued a formal data request covering all California trading activity from May-August 2001. Response deadline is October 30. We need to compile all trading records, communications, and strategy documents. Action: Legal to coordinate document collection. Owner: Rex Rogers."},
                {"from_": "jeff.dasovich@enron.com", "subject": "Re: FERC Investigation Response",
                 "body": "Government affairs is engaged. We have a call with FERC staff on Wednesday. I recommend we proactively disclose the Death Star and Ricochet strategies before they find them. Follow up: prepare full disclosure memo by Tuesday."},
                {"from_": "tim.belden@enron.com", "subject": "Re: FERC Investigation Response",
                 "body": "Trading records have been pulled. There are some email chains that will be problematic. Legal hold has been placed on all relevant accounts. Action: Tim to provide annotated trading log by Friday."},
                {"from_": "rex.rogers@enron.com", "subject": "Re: FERC Investigation Response",
                 "body": "Engaged outside counsel at Vinson & Elkins. They will lead the regulatory response. Do NOT discuss specifics over email. All strategy discussions should be in-person or over the phone. Action: all team members to brief outside counsel by Wednesday."},
            ]
        },
        {
            "key": "enron online performance review",
            "emails": [
                {"from_": "louise.kitchen@enron.com", "subject": "Enron Online Performance Review",
                 "body": "EOL processed $2.8B in transactions last month — a new record. However, the gas desk is experiencing latency above 200ms during peak hours. IT needs to investigate. Action: IT team to diagnose latency issue by next Wednesday. Owner: Louise Kitchen."},
                {"from_": "kevin.hannon@enron.com", "subject": "Re: Enron Online Performance Review",
                 "body": "The latency issue is linked to the new credit check module. We can disable it temporarily during peak hours as a workaround. Full fix will take 2 weeks. Follow up: Kevin to implement temporary workaround by Thursday."},
                {"from_": "louise.kitchen@enron.com", "subject": "Re: Enron Online Performance Review",
                 "body": "Approved the workaround. Also please prepare the October metrics report for the board presentation. Action: October metrics report due November 5."},
            ]
        },
        {
            "key": "employee stock plan changes",
            "emails": [
                {"from_": "cindy.olson@enron.com", "subject": "Employee Stock Plan Changes",
                 "body": "Following the recent stock performance, we are revising the employee stock purchase plan. Key change: lockup period extended from 6 to 12 months for executive grants. HR will send all-hands communication by October 28. Action: Cindy to draft communications. Owner: Cindy Olson."},
                {"from_": "ken.lay@enron.com", "subject": "Re: Employee Stock Plan Changes",
                 "body": "I want to emphasise that this is a long-term investment in our people. Please ensure the messaging is positive and focuses on the company's strong fundamentals. Follow up: Ken to record video message for employees."},
                {"from_": "cindy.olson@enron.com", "subject": "Re: Employee Stock Plan Changes",
                 "body": "Draft communications attached. Also note that the 401k match is being temporarily suspended — this needs careful messaging. Action: legal review of communications by Friday."},
            ]
        },
        {
            "key": "risk committee var limit breach",
            "emails": [
                {"from_": "rick.buy@enron.com", "subject": "Risk Committee — VaR Limit Breach",
                 "body": "URGENT: The power trading desk breached its daily VaR limit by 34% on Thursday. Risk committee convening emergency session Monday 8am. All desk heads must attend. Action: Reset VaR limits and present remediation plan. Owner: Rick Buy. URGENT ESCALATION."},
                {"from_": "greg.whalley@enron.com", "subject": "Re: Risk Committee — VaR Limit Breach",
                 "body": "I've reviewed the positions. The breach was driven by the PG&E counterparty exposure. We need to unwind at least 40% of the position before Monday open. Action: Greg to coordinate position unwind with trading desk Sunday."},
                {"from_": "rick.buy@enron.com", "subject": "Re: Risk Committee — VaR Limit Breach",
                 "body": "Board audit committee has been notified per policy. External risk consultants are being engaged. This will be disclosed in the next 8-K. Follow up: file 8-K disclosure by Tuesday."},
            ]
        },
        {
            "key": "india dabhol power project",
            "emails": [
                {"from_": "rebecca.mark@enron.com", "subject": "India — Dabhol Power Project",
                 "body": "The Maharashtra government has formally stopped paying for power from the Dabhol plant. We are owed $64M in arrears. OPIC insurance claim being filed. Action: Rebecca to lead arbitration proceedings. Owner: Rebecca Mark."},
                {"from_": "james.derrick@enron.com", "subject": "Re: India — Dabhol Power Project",
                 "body": "OPIC claim filed. We also have bilateral investment treaty protections that can be invoked. The State Department has been briefed. Follow up: James to coordinate with State Dept by end of week."},
                {"from_": "ken.lay@enron.com", "subject": "Re: India — Dabhol Power Project",
                 "body": "I spoke with the US Ambassador in Delhi. There is political will to resolve this but it will take time. Recommend we explore selling our stake to a strategic buyer. Action: engage Goldman Sachs to explore sale options."},
            ]
        },
        {
            "key": "trading floor systems upgrade",
            "emails": [
                {"from_": "scott.yeager@enron.com", "subject": "Trading Floor Systems Upgrade",
                 "body": "The Q4 systems upgrade is on track. New workstations are being deployed this weekend. All traders must back up their local files by Friday. Action: IT to complete workstation deployment by Sunday night. Owner: Scott Yeager."},
                {"from_": "kevin.hannon@enron.com", "subject": "Re: Trading Floor Systems Upgrade",
                 "body": "Confirmed — backup protocol is in place. The new systems include upgraded Bloomberg terminals and a new order management interface. Training sessions scheduled for Monday-Tuesday. Follow up: all traders to attend mandatory training."},
                {"from_": "scott.yeager@enron.com", "subject": "Re: Trading Floor Systems Upgrade",
                 "body": "Failover testing completed successfully. Disaster recovery systems are operational. Final sign-off from CTO required before go-live. Action: CTO sign-off by Friday."},
            ]
        },
    ]

    threads = {}
    for i, sample in enumerate(SAMPLES):
        key = sample["key"]
        email_list = []
        for j, e in enumerate(sample["emails"]):
            dt = base_date + timedelta(days=i, hours=j * 3)
            email_list.append({
                "message_id": f"<{i}-{j}@enron.com>",
                "date":       dt.strftime("%a, %d %b %Y %H:%M:%S +0000"),
                "from_":      e["from_"],
                "to":         "enron-group@enron.com",
                "cc":         "",
                "subject":    e["subject"],
                "body":       e["body"],
                "thread_key": key,
            })
        threads[key] = email_list

    return threads
