from __future__ import annotations
"""
utils/nlp_engine.py
-------------------
Zero-API NLP pipeline for email thread analysis.

Uses only open-source, locally-running libraries:
  - sumy (LSA extractive summarisation)
  - KeyBERT (keyphrase extraction via sentence-transformers)
  - VADER (rule-based sentiment)
  - spaCy (NER for person/owner detection)
  - Regex patterns (action item detection)

Falls back to pure stdlib implementations when optional
heavy packages are not installed, so the project always runs.
"""

import re
import math
from collections import defaultdict, Counter
from datetime import datetime, date
from typing import Optional


# ─────────────────────────────────────────────────────────────────────────────
# OPTIONAL IMPORTS — graceful fallbacks
# ─────────────────────────────────────────────────────────────────────────────

try:
    from sumy.parsers.plaintext import PlaintextParser
    from sumy.nlp.tokenizers import Tokenizer
    from sumy.summarizers.lsa import LsaSummarizer
    from sumy.nlp.stemmers import Stemmer
    from sumy.utils import get_stop_words
    SUMY_AVAILABLE = True
except Exception:
    SUMY_AVAILABLE = False

try:
    from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
    _vader = SentimentIntensityAnalyzer()
    VADER_AVAILABLE = True
except Exception:
    VADER_AVAILABLE = False

try:
    import spacy
    _nlp = spacy.load("en_core_web_sm")
    SPACY_AVAILABLE = True
except Exception:
    SPACY_AVAILABLE = False

try:
    from keybert import KeyBERT
    _kw_model = KeyBERT()
    KEYBERT_AVAILABLE = True
except Exception:
    KEYBERT_AVAILABLE = False


# ─────────────────────────────────────────────────────────────────────────────
# STOPWORDS (built-in so no NLTK download needed)
# ─────────────────────────────────────────────────────────────────────────────

STOPWORDS = {
    "the","and","for","are","was","that","this","with","have","from",
    "they","will","been","were","also","more","some","when","your","which",
    "their","there","then","than","had","has","not","but","its","our",
    "you","all","one","can","she","him","his","her","may","who","any",
    "would","could","should","said","well","just","into","over","after",
    "about","email","please","enron","thanks","thank","regards","dear",
    "need","know","think","want","make","good","time","look","come",
    "send","review","update","attached","meeting","team","let","call",
    "get","per","per","see","use","sure","note","work","day","week",
}

URGENT_WORDS = re.compile(
    r"\b(urgent|asap|immediately|critical|emergency|deadline|overdue|"
    r"escalat|crisis|action\s+required|high\s+priority|time\s+sensitive|"
    r"respond\s+immediately|today|by\s+eod|by\s+cob)\b",
    re.IGNORECASE
)

ACTION_PATTERNS = [
    r"(?:please|pls|kindly)[,\s]+([^.!?\n]{10,120})",
    r"(?:action\s*(?:item|required|needed)?)[:\s]+([^.!?\n]{10,120})",
    r"(?:follow[\s\-]up)[:\s]+([^.!?\n]{8,100})",
    r"(?:could\s+you|can\s+you|would\s+you)[,\s]+([^.!?\n]{10,100})",
    r"(?:need\s+to|needs\s+to|must|should\s+be)[:\s]+([^.!?\n]{8,100})",
    r"(?:deadline|due\s+by|by\s+[A-Z][a-z]+day|by\s+end\s+of)[:\s]+([^.!?\n]{5,80})",
    r"(?:owner)[:\s]+([^.!?\n]{5,60})",
    r"(?:i\s+will|i'll|we\s+will|we'll)\s+([^.!?\n]{8,100})",
    r"(?:send|submit|review|approve|confirm|schedule|prepare|complete|"
    r"update|compile|draft|coordinate|engage|notify|brief)\s+([^.!?\n]{8,100})",
]


# ─────────────────────────────────────────────────────────────────────────────
# SUMMARISATION
# ─────────────────────────────────────────────────────────────────────────────

def summarise_text(text: str, n_sentences: int = 3) -> str:
    """
    Produce an extractive summary of the given text.
    Uses sumy LSA if available, else a TF-based fallback.
    """
    if not text or len(text.split()) < 15:
        return text.strip()[:400]

    if SUMY_AVAILABLE:
        try:
            parser = PlaintextParser.from_string(text, Tokenizer("english"))
            stemmer = Stemmer("english")
            summariser = LsaSummarizer(stemmer)
            summariser.stop_words = get_stop_words("english")
            sentences = summariser(parser.document, n_sentences)
            result = " ".join(str(s) for s in sentences)
            if result.strip():
                return result
        except Exception:
            pass

    # Fallback: score sentences by word frequency
    return _tfidf_summary(text, n_sentences)


def _tfidf_summary(text: str, n: int = 3) -> str:
    """
    Simple TF-based extractive summary.
    Scores each sentence by the sum of its word frequencies,
    penalising very short or very long sentences.
    """
    sentences = re.split(r"(?<=[.!?])\s+", text)
    sentences = [s.strip() for s in sentences if len(s.split()) >= 5]
    if not sentences:
        return text[:400]

    # Build word frequency table
    words = re.findall(r"\b[a-z]{3,}\b", text.lower())
    freq = Counter(w for w in words if w not in STOPWORDS)
    if not freq:
        return " ".join(sentences[:n])

    max_freq = max(freq.values())
    freq = {w: v / max_freq for w, v in freq.items()}

    # Score each sentence
    def score(sent: str) -> float:
        wds = re.findall(r"\b[a-z]{3,}\b", sent.lower())
        if not wds:
            return 0.0
        s = sum(freq.get(w, 0) for w in wds if w not in STOPWORDS)
        # Penalise very short sentences
        length_bonus = min(len(wds) / 20, 1.0)
        return s * length_bonus

    scored = sorted(enumerate(sentences), key=lambda x: score(x[1]), reverse=True)
    top_indices = sorted([i for i, _ in scored[:n]])
    return " ".join(sentences[i] for i in top_indices)


# ─────────────────────────────────────────────────────────────────────────────
# TOPIC EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────

def extract_topic(text: str) -> str:
    """
    Extract the key topic/keyphrase from a block of text.
    Uses KeyBERT if available, else TF-IDF bigram extraction.
    """
    if not text or len(text.split()) < 5:
        return "General Discussion"

    if KEYBERT_AVAILABLE:
        try:
            kws = _kw_model.extract_keywords(
                text[:2000],
                keyphrase_ngram_range=(1, 3),
                stop_words="english",
                top_n=3
            )
            if kws:
                return ", ".join(kw for kw, _ in kws)
        except Exception:
            pass

    return _tfidf_topic(text)


def _tfidf_topic(text: str, top_n: int = 3) -> str:
    """Extract top keywords using TF-IDF-like scoring."""
    # Build bigrams + unigrams
    words = re.findall(r"\b[a-z]{4,}\b", text.lower())
    words = [w for w in words if w not in STOPWORDS]

    unigrams = Counter(words)
    bigrams  = Counter(
        f"{words[i]} {words[i+1]}"
        for i in range(len(words) - 1)
        if words[i] not in STOPWORDS and words[i+1] not in STOPWORDS
    )

    # Weight bigrams higher than unigrams
    combined = {**{k: v * 1.5 for k, v in bigrams.items()}, **unigrams}
    top = sorted(combined, key=combined.get, reverse=True)[:top_n]
    return ", ".join(top) if top else "General Discussion"


# ─────────────────────────────────────────────────────────────────────────────
# ACTION ITEM EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────

def extract_action_items(text: str) -> list[str]:
    """
    Find action items / tasks / follow-ups in the email body.
    Returns a deduplicated list of up to 5 action strings.
    """
    items = []
    for pat in ACTION_PATTERNS:
        for m in re.finditer(pat, text, re.IGNORECASE):
            item = m.group(1).strip().rstrip(".,;:")
            # Filter noise
            if 8 < len(item) < 150 and len(item.split()) >= 2:
                items.append(item)

    # Deduplicate (keep first occurrence)
    seen, unique = set(), []
    for it in items:
        key = it.lower()[:50]
        if key not in seen:
            seen.add(key)
            unique.append(it)

    return unique[:5]


# ─────────────────────────────────────────────────────────────────────────────
# OWNER DETECTION
# ─────────────────────────────────────────────────────────────────────────────

def extract_owner(thread_emails: list[dict]) -> str:
    """
    Identify the most likely owner / responsible person for the thread.

    Priority:
      1. Most frequent sender (by display name)
      2. First PERSON entity found in email bodies (spaCy NER)
      3. Email username of most frequent sender
    """
    # Build sender name list
    sender_names = []
    for e in thread_emails:
        frm = e.get("from_", "")
        # Extract display name (before the email address)
        name_match = re.match(r'^"?([^"<@\n]{3,40})"?\s*<?', frm)
        if name_match:
            name = name_match.group(1).strip().rstrip(".,")
            if name and "@" not in name and len(name) > 2:
                # Convert "last.first" email username to "First Last"
                if re.match(r"^[a-z]+\.[a-z]+$", name.lower()):
                    parts = name.split(".")
                    name = " ".join(p.capitalize() for p in parts)
                sender_names.append(name)

    if sender_names:
        freq = Counter(sender_names)
        return freq.most_common(1)[0][0]

    # spaCy NER fallback
    if SPACY_AVAILABLE:
        combined = " ".join(e.get("body", "")[:300] for e in thread_emails[:3])
        try:
            doc = _nlp(combined[:1500])
            persons = [ent.text for ent in doc.ents if ent.label_ == "PERSON"]
            if persons:
                return Counter(persons).most_common(1)[0][0]
        except Exception:
            pass

    # Last resort: parse email username
    for e in thread_emails:
        frm = e.get("from_", "")
        m = re.search(r"([a-z]+)\.([a-z]+)@", frm, re.IGNORECASE)
        if m:
            return f"{m.group(1).capitalize()} {m.group(2).capitalize()}"

    return "Unassigned"


# ─────────────────────────────────────────────────────────────────────────────
# SENTIMENT ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

def get_sentiment(text: str) -> str:
    """
    Classify thread sentiment as: urgent | negative | positive | neutral
    VADER is used if available; falls back to keyword counting.
    """
    # Urgency check takes priority over all other sentiment
    if URGENT_WORDS.search(text):
        return "urgent"

    if VADER_AVAILABLE:
        try:
            score = _vader.polarity_scores(text)["compound"]
            if score >= 0.15:  return "positive"
            if score <= -0.15: return "negative"
            return "neutral"
        except Exception:
            pass

    # Simple keyword-based fallback
    text_lower = text.lower()
    positive_kw = ["great","excellent","pleased","good","thanks","congratulations",
                   "success","approve","confirmed","complete","done","agreed"]
    negative_kw = ["concern","problem","issue","delay","breach","risk","loss",
                   "penalty","fail","error","wrong","not","cannot","unable"]

    pos_count = sum(1 for w in positive_kw if w in text_lower)
    neg_count = sum(1 for w in negative_kw if w in text_lower)

    if neg_count > pos_count + 1: return "negative"
    if pos_count > neg_count + 1: return "positive"
    return "neutral"


# ─────────────────────────────────────────────────────────────────────────────
# FOLLOW-UP DETECTION
# ─────────────────────────────────────────────────────────────────────────────

def extract_followups(text: str) -> list[str]:
    """Extract explicit follow-up items from the text."""
    patterns = [
        r"follow[\s\-]?up[:\s]+([^.!?\n]{8,120})",
        r"(?:next\s+steps?)[:\s]+([^.!?\n]{8,120})",
        r"(?:pending)[:\s]+([^.!?\n]{8,100})",
        r"(?:waiting\s+(?:for|on))[:\s]+([^.!?\n]{8,100})",
        r"(?:will\s+revert|will\s+get\s+back|will\s+confirm)[^.!?\n]{0,80}",
    ]
    items = []
    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            try:
                item = m.group(1).strip().rstrip(".,;:")
            except IndexError:
                item = m.group(0).strip()
            if 8 < len(item) < 150:
                items.append(item)

    seen, unique = set(), []
    for it in items:
        k = it.lower()[:40]
        if k not in seen:
            seen.add(k)
            unique.append(it)
    return unique[:3]


# ─────────────────────────────────────────────────────────────────────────────
# PARTICIPANT EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────

def extract_participants(thread_emails: list[dict]) -> list[str]:
    """Return a deduplicated list of participant email addresses."""
    seen = set()
    participants = []
    for e in thread_emails:
        for field in ("from_", "to", "cc"):
            val = e.get(field, "")
            emails_found = re.findall(r"[\w.\-]+@[\w.\-]+", val)
            for em in emails_found:
                em = em.lower().strip()
                if em not in seen and "enron.com" in em:
                    seen.add(em)
                    participants.append(em)
    return participants[:8]


# ─────────────────────────────────────────────────────────────────────────────
# MASTER ANALYSIS FUNCTION
# ─────────────────────────────────────────────────────────────────────────────

def analyse_thread(thread_key: str, emails: list[dict]) -> dict:
    """
    Run the full NLP pipeline on one email thread.

    Args:
        thread_key : Normalised subject string (the thread ID).
        emails     : List of email dicts for this thread.

    Returns:
        Dict with all extracted intelligence fields.
    """
    # Combine all bodies with speaker attribution
    combined_parts = []
    for e in emails:
        sender_name = e.get("from_", "").split("@")[0].replace(".", " ").title()
        body = e.get("body", "").strip()
        if body:
            combined_parts.append(f"{sender_name}: {body}")

    combined_text = "\n\n".join(combined_parts)

    # Display title: title-case the thread key, cap length
    thread_title = thread_key.replace("-", " ").replace("_", " ").title()
    if len(thread_title) > 70:
        thread_title = thread_title[:67] + "…"

    # Run all extractions
    summary      = summarise_text(combined_text, n_sentences=3)
    topic        = extract_topic(combined_text)
    action_items = extract_action_items(combined_text)
    owner        = extract_owner(emails)
    sentiment    = get_sentiment(combined_text)
    follow_ups   = extract_followups(combined_text)
    participants = extract_participants(emails)

    # Derive thread date (most recent email date)
    dates = []
    for e in emails:
        try:
            from email.utils import parsedate_to_datetime
            dt = parsedate_to_datetime(e.get("date", ""))
            dates.append(dt)
        except Exception:
            pass
    latest_date = max(dates).strftime("%d %b %Y") if dates else "Unknown"

    return {
        "Thread Key"      : thread_key,
        "Email Thread"    : thread_title,
        "Key Topic"       : topic,
        "Summary"         : summary,
        "Action Items"    : " | ".join(action_items) if action_items else "None identified",
        "Action List"     : action_items,
        "Owner"           : owner,
        "Follow-ups"      : " | ".join(follow_ups) if follow_ups else "None",
        "Follow-up List"  : follow_ups,
        "Participants"    : ", ".join(participants),
        "Participant List": participants,
        "Sentiment"       : sentiment,
        "Email Count"     : len(emails),
        "Latest Date"     : latest_date,
    }
