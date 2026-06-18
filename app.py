"""
PubMed Miner — Streamlit Web Interface
Imports search, scoring, and export functions directly from pubmed_miner.py.

Author:      Marouane Takie
Affiliation: PhD Candidate | Physiology, Pharmacology & Phytochemistry
Institution: Université Sidi Mohamed Ben Abdellah (USMBA), Fès, Morocco
"""

import io
import json
import os
import sys
import tempfile
import time
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

# Import everything needed from the CLI script in the same folder
sys.path.insert(0, str(Path(__file__).parent))
from pubmed_miner import (
    _AI_PROMPT,
    _build_summary_df,
    ArticleCache,
    BATCH_SIZE,
    COLUMNS_BASE,
    COLUMNS_AI,
    export_excel,
    fetch_details,
    score_relevance,
    search_europepmc,
    search_pubmed,
    to_bibtex,
    to_ris,
)

# ---------------------------------------------------------------------------
# Page config — must be first Streamlit call
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="PubMed Miner",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
)


try:
    groq_key = st.secrets["GROQ_API_KEY"]
except Exception:
    groq_key = ""

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------

st.markdown("""
<style>
/* Paper cards */
.paper-card {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 8px;
    padding: 16px 18px;
    margin-bottom: 12px;
}
.paper-title {
    font-size: 1.0rem;
    font-weight: 600;
    color: #e6edf3;
    line-height: 1.4;
}
.paper-authors {
    color: #8b949e;
    font-size: 0.82rem;
    margin-top: 3px;
}
.paper-meta {
    margin-top: 6px;
}
.paper-abstract {
    color: #c9d1d9;
    font-size: 0.88rem;
    margin-top: 8px;
    line-height: 1.55;
}
.ai-box {
    margin-top: 10px;
    padding: 8px 12px;
    background: #0d1117;
    border: 1px solid #21262d;
    border-radius: 6px;
    font-size: 0.82rem;
    line-height: 1.6;
}
.ai-box-header {
    color: #8b949e;
    font-size: 0.78rem;
    margin-bottom: 4px;
}

/* Inline tags / badges */
.tag {
    display: inline-block;
    background: #21262d;
    border: 1px solid #30363d;
    border-radius: 4px;
    padding: 1px 7px;
    font-size: 0.75rem;
    color: #c9d1d9;
    margin: 2px 3px 2px 0;
    vertical-align: middle;
}
.tag-epmc { border-color: #388bfd44; color: #79c0ff; }
.tag-pm   { border-color: #3fb95044; color: #56d364; }

/* Relevance score */
.score-high { color: #3fb950; font-weight: 700; font-size: 1.05rem; }
.score-mid  { color: #d29922; font-weight: 700; font-size: 1.05rem; }
.score-low  { color: #f85149; font-weight: 700; font-size: 1.05rem; }

/* DOI / external links */
.ext-link {
    color: #58a6ff;
    text-decoration: none;
    font-size: 0.82rem;
}
.ext-link:hover { text-decoration: underline; }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# AI extraction helpers (Groq and Gemini via requests — no extra SDK needed)
# ---------------------------------------------------------------------------

_GROQ_ENDPOINT   = "https://api.groq.com/openai/v1/chat/completions"
_GROQ_MODEL      = "llama-3.3-70b-versatile"
_GEMINI_ENDPOINT = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "gemini-1.5-flash:generateContent"
)


def _parse_ai_json(text: str) -> dict:
    text = text.strip()
    if "```" in text:
        parts = text.split("```")
        text = parts[1] if len(parts) > 1 else text
        if text.startswith("json"):
            text = text[4:]
    try:
        return json.loads(text.strip())
    except (json.JSONDecodeError, ValueError):
        return {k: "" for k in ["species", "compounds", "activities", "key_results", "study_type"]}


def _build_ai_prompt(record: dict) -> str:
    return _AI_PROMPT.format(
        title=str(record.get("Title", ""))[:300],
        abstract=str(record.get("Abstract", ""))[:2000],
    )


def ai_extract_groq(record: dict, api_key: str) -> dict:
    if not record.get("Abstract"):
        return {k: "" for k in ["species", "compounds", "activities", "key_results", "study_type"]}
    try:
        r = requests.post(
            _GROQ_ENDPOINT,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": _GROQ_MODEL,
                "messages": [{"role": "user", "content": _build_ai_prompt(record)}],
                "max_tokens": 512,
                "temperature": 0.1,
            },
            timeout=30,
        )
        r.raise_for_status()
        return _parse_ai_json(r.json()["choices"][0]["message"]["content"])
    except Exception:
        return {k: "" for k in ["species", "compounds", "activities", "key_results", "study_type"]}


def ai_extract_gemini(record: dict, api_key: str) -> dict:
    if not record.get("Abstract"):
        return {k: "" for k in ["species", "compounds", "activities", "key_results", "study_type"]}
    try:
        r = requests.post(
            f"{_GEMINI_ENDPOINT}?key={api_key}",
            json={
                "contents": [{"parts": [{"text": _build_ai_prompt(record)}]}],
                "generationConfig": {"maxOutputTokens": 512, "temperature": 0.1},
            },
            timeout=30,
        )
        r.raise_for_status()
        text = r.json()["candidates"][0]["content"]["parts"][0]["text"]
        return _parse_ai_json(text)
    except Exception:
        return {k: "" for k in ["species", "compounds", "activities", "key_results", "study_type"]}


# ---------------------------------------------------------------------------
# Export helpers — write to NamedTemporaryFile, return bytes for download
# ---------------------------------------------------------------------------

def _df_to_excel_bytes(df: pd.DataFrame, queries: list[str], threshold: int) -> bytes:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        tmp = f.name
    try:
        export_excel(df, queries, threshold, tmp)
        return Path(tmp).read_bytes()
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


def _df_to_bibtex_bytes(df: pd.DataFrame) -> bytes:
    with tempfile.NamedTemporaryFile(suffix=".bib", delete=False, mode="w", encoding="utf-8") as f:
        tmp = f.name
    to_bibtex(df, tmp)
    data = Path(tmp).read_bytes()
    try:
        os.unlink(tmp)
    except OSError:
        pass
    return data


def _df_to_ris_bytes(df: pd.DataFrame) -> bytes:
    with tempfile.NamedTemporaryFile(suffix=".ris", delete=False, mode="w", encoding="utf-8") as f:
        tmp = f.name
    to_ris(df, tmp)
    data = Path(tmp).read_bytes()
    try:
        os.unlink(tmp)
    except OSError:
        pass
    return data


# ---------------------------------------------------------------------------
# Relevance colour helper
# ---------------------------------------------------------------------------

def _score_class(score: int) -> str:
    if score >= 70:
        return "score-high"
    if score >= 40:
        return "score-mid"
    return "score-low"


# ---------------------------------------------------------------------------
# Core search runner (called from the "Start Search" button)
# ---------------------------------------------------------------------------

def run_search(
    queries: list[str],
    source: str,
    max_results: int,
    year_from: int | None,
    ncbi_key: str | None,
    rel_keywords: str,
    ai_enabled: bool,
    ai_provider: str,
    ai_key: str,
    threshold: int,
    progress_bar,
    status_text,
) -> pd.DataFrame:

    uid_to_rec:     dict[str, dict]      = {}
    uid_to_queries: dict[str, list[str]] = {}
    doi_to_uid:     dict[str, str]       = {}
    sleep = 0.12 if ncbi_key else 0.4

    def _register(rec: dict, query: str):
        uid = rec["UID"]
        doi = (rec.get("DOI") or "").lower().strip()
        if doi and doi in doi_to_uid:
            uid = doi_to_uid[doi]
        elif doi:
            doi_to_uid[doi] = uid
        if uid in uid_to_rec:
            uid_to_queries[uid].append(query)
            new_score = score_relevance(uid_to_rec[uid], query)
            if new_score > uid_to_rec[uid]["Relevance_Score"]:
                uid_to_rec[uid]["Relevance_Score"] = new_score
        else:
            rec["Relevance_Score"] = score_relevance(rec, query)
            uid_to_rec[uid]        = rec
            uid_to_queries[uid]    = [query]

    # Count expected steps for the progress bar
    steps_per_q = (1 if source == "pubmed" else 0) + (1 if source in ("europepmc", "both") else 0)
    if source == "both":
        steps_per_q = 2
    total_steps = len(queries) * steps_per_q + (1 if ai_enabled and ai_key else 0)
    done_steps  = 0

    def _tick(msg: str):
        nonlocal done_steps
        done_steps += 1
        progress_bar.progress(min(done_steps / max(total_steps, 1), 0.99))
        status_text.text(msg)

    for query in queries:
        if source in ("pubmed", "both"):
            status_text.text(f"PubMed → {query[:70]} …")
            try:
                pmids = search_pubmed(query, max_results, year_from, ncbi_key)
                for i in range(0, len(pmids), BATCH_SIZE):
                    batch = pmids[i : i + BATCH_SIZE]
                    for rec in fetch_details(batch, ncbi_key):
                        _register(rec, query)
                    time.sleep(sleep)
            except Exception as exc:
                st.warning(f"PubMed error — {query!r}: {exc}")
            _tick(f"PubMed done: {query[:50]}")

        if source in ("europepmc", "both"):
            status_text.text(f"Europe PMC → {query[:70]} …")
            try:
                for rec in search_europepmc(query, max_results, year_from):
                    _register(rec, query)
                time.sleep(sleep)
            except Exception as exc:
                st.warning(f"Europe PMC error — {query!r}: {exc}")
            _tick(f"Europe PMC done: {query[:50]}")

    # Merge query strings into each record
    for uid, rec in uid_to_rec.items():
        rec["Queries"] = "; ".join(uid_to_queries.get(uid, []))

    # Relevance keyword re-scoring (boost, take max)
    if rel_keywords.strip():
        kw_query = " ".join(rel_keywords.replace("\n", " ").split())
        for rec in uid_to_rec.values():
            kw_score = score_relevance(rec, kw_query)
            if kw_score > rec["Relevance_Score"]:
                rec["Relevance_Score"] = min(100, kw_score)

    if not uid_to_rec:
        return pd.DataFrame(columns=COLUMNS_BASE)

    # AI extraction
    has_ai_cols = False
    if ai_enabled and ai_key:
        ai_fn = ai_extract_groq if ai_provider == "Groq" else ai_extract_gemini
        records = list(uid_to_rec.values())
        n = len(records)
        for i, rec in enumerate(records):
            status_text.text(f"🤖 {ai_provider} extraction {i + 1}/{n} …")
            ext = ai_fn(rec, ai_key)
            rec["AI_Species"]     = ext.get("species", "")
            rec["AI_Compounds"]   = ext.get("compounds", "")
            rec["AI_Activities"]  = ext.get("activities", "")
            rec["AI_Key_Results"] = ext.get("key_results", "")
            rec["AI_Study_Type"]  = ext.get("study_type", "")
            time.sleep(0.05)
        has_ai_cols = True
        _tick(f"AI extraction done ({n} abstracts)")

    cols = COLUMNS_BASE + (COLUMNS_AI if has_ai_cols else [])
    df = (
        pd.DataFrame(uid_to_rec.values())
        .reindex(columns=cols, fill_value="")
        .sort_values("Relevance_Score", ascending=False)
        .reset_index(drop=True)
    )

    progress_bar.progress(1.0)
    status_text.text(f"Done — {len(df)} unique records.")
    return df


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown("## ⚙️ Search Configuration")

    source_label = st.selectbox(
        "Literature source",
        ["PubMed", "Europe PMC", "Both"],
        index=0,
    )
    source_map = {"PubMed": "pubmed", "Europe PMC": "europepmc", "Both": "both"}

    st.divider()

    yr_col1, yr_col2 = st.columns(2)
    year_from_ui = yr_col1.number_input("From year", min_value=1950, max_value=2030, value=2015, step=1)
    yr_col2.number_input("To year",   min_value=1950, max_value=2030, value=2026, step=1)
    year_from_val = int(year_from_ui)

    max_results_ui = st.slider("Max results / query", 10, 200, 50, 10)
    rel_threshold  = st.slider("High-relevance threshold", 0, 100, 60, 5,
                               help="Minimum Relevance Score for the High Relevance Excel sheet.")

    ncbi_api_key = st.text_input(
        "NCBI API key (optional)",
        type="password",
        help="Raises the rate limit from 3 → 10 req/s. Register free at ncbi.nlm.nih.gov/account/",
    )

    st.divider()

    ai_enabled = st.toggle("🤖 AI Analysis", value=False,
                           help="Extract species, compounds, activities, and key results from abstracts.")
    ai_provider = "Groq"
    ai_api_key  = ""
    if ai_enabled:
        if groq_key:
            # Secret is available — no input needed
            ai_api_key = groq_key
            st.success("✓ AI analysis ready — powered by Groq (free)")
        else:
            # No secret — let the user supply a key manually
            ai_provider = st.selectbox("AI Provider", ["Groq", "Gemini"])
            ai_api_key  = st.text_input(
                f"{ai_provider} API key",
                type="password",
                help="Groq: console.groq.com  |  Gemini: aistudio.google.com",
            )
            if not ai_api_key:
                st.info("Enter an API key to enable AI extraction.")

    st.divider()
    st.caption(
        "**PubMed Miner v3.1**\n\n"
        "[GitHub ↗](https://github.com/marouanetakie/pubmed-miner)  ·  "
        "[NCBI E-utilities](https://www.ncbi.nlm.nih.gov/books/NBK25500/)  ·  "
        "[Europe PMC API](https://europepmc.org/RestfulWebService)"
    )


# ---------------------------------------------------------------------------
# Main — title & search inputs
# ---------------------------------------------------------------------------

st.title("🔬 PubMed Miner")
st.caption(
    "Search PubMed & Europe PMC · Score relevance · Export to Excel, BibTeX, RIS · 100% Free"
)

q_col, kw_col = st.columns([3, 2], gap="medium")

with q_col:
    queries_raw = st.text_area(
        "Search queries — one per line",
        value="",
        placeholder="Enter one search query per line\ne.g. antioxidant activity plant extract\ne.g. antimicrobial medicinal plants\ne.g. anti-inflammatory natural compounds\ne.g. biological activity phenolic compounds",
        height=190,
        help="Each line is sent as a separate query. Use PubMed/Europe PMC syntax (AND, OR, [MeSH]).",
    )

with kw_col:
    rel_keywords = st.text_area(
        "Relevance keywords (boost scoring)",
        placeholder=(
            "antioxidant\nIC50\nflavonoid\nphenolic\n"
            "cytotoxic\nMIC\ndocking\nUPLC"
        ),
        height=190,
        help=(
            "Articles matching these terms get a relevance score boost. "
            "Leave blank to score purely by query term overlap."
        ),
    )

queries = [q.strip() for q in queries_raw.splitlines() if q.strip()]

btn_col, _ = st.columns([1, 5])
start = btn_col.button(
    "🔍 Start Search",
    type="primary",
    disabled=not queries,
    use_container_width=True,
)

# ---------------------------------------------------------------------------
# Search execution
# ---------------------------------------------------------------------------

if start:
    prog  = st.progress(0.0)
    stext = st.empty()

    with st.spinner("Searching …"):
        df = run_search(
            queries      = queries,
            source       = source_map[source_label],
            max_results  = max_results_ui,
            year_from    = year_from_val,
            ncbi_key     = ncbi_api_key or None,
            rel_keywords = rel_keywords,
            ai_enabled   = ai_enabled,
            ai_provider  = ai_provider,
            ai_key       = ai_api_key or groq_key,
            threshold    = rel_threshold,
            progress_bar = prog,
            status_text  = stext,
        )

    st.session_state["df"]        = df
    st.session_state["queries"]   = queries
    st.session_state["threshold"] = rel_threshold
    st.session_state["source"]    = source_label

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------

if "df" in st.session_state:
    df        = st.session_state["df"]
    qs        = st.session_state["queries"]
    threshold = st.session_state["threshold"]

    if df.empty:
        st.info("No results found. Try broader queries or a wider year range.")
        st.stop()

    st.divider()

    # ── Stats cards ──────────────────────────────────────────────────────────
    total       = len(df)
    high_rel    = int((df["Relevance_Score"] >= threshold).sum())
    pm_count    = int((df["Source"] == "pubmed").sum())
    epmc_count  = int((df["Source"] == "europepmc").sum())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Papers",   total)
    m2.metric("High Relevance", high_rel, f"≥ {threshold} score")
    m3.metric("PubMed",         pm_count)
    m4.metric("Europe PMC",     epmc_count)

    # ── Download buttons ──────────────────────────────────────────────────────
    st.subheader("📥 Export")
    dl1, dl2, dl3 = st.columns(3)

    with dl1:
        try:
            xls = _df_to_excel_bytes(df, qs, threshold)
            st.download_button(
                "⬇ Excel (.xlsx)",
                data=xls,
                file_name="pubmed_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Excel export failed: {exc}")

    with dl2:
        try:
            bib = _df_to_bibtex_bytes(df)
            st.download_button(
                "⬇ BibTeX (.bib)",
                data=bib,
                file_name="pubmed_results.bib",
                mime="text/plain",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"BibTeX export failed: {exc}")

    with dl3:
        try:
            ris = _df_to_ris_bytes(df)
            st.download_button(
                "⬇ RIS (.ris)",
                data=ris,
                file_name="pubmed_results.ris",
                mime="text/plain",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"RIS export failed: {exc}")

    # ── Results tabs ──────────────────────────────────────────────────────────
    st.subheader("📄 Papers")
    tab_cards, tab_table = st.tabs(["Card View", "Table View"])

    # ── Table View ────────────────────────────────────────────────────────────
    with tab_table:
        display_cols = [c for c in df.columns if c not in ("Abstract", "MeSH_Terms")]
        st.dataframe(df[display_cols], use_container_width=True, height=520)

    # ── Card View ─────────────────────────────────────────────────────────────
    with tab_cards:
        n_show = st.slider(
            "Papers shown",
            min_value=5,
            max_value=min(200, total),
            value=min(20, total),
            step=5,
            key="cards_slider",
        )
        has_ai = "AI_Key_Results" in df.columns

        for _, row in df.head(n_show).iterrows():
            score    = int(row.get("Relevance_Score", 0))
            css      = _score_class(score)
            doi      = str(row.get("DOI", "")).strip()
            pmid     = str(row.get("PMID", "")).strip()
            src      = str(row.get("Source", ""))
            title    = str(row.get("Title", ""))
            authors  = str(row.get("Authors", ""))
            journal  = str(row.get("Journal", ""))
            year     = str(row.get("Year", ""))
            abstract = str(row.get("Abstract", ""))
            abstract_short = abstract[:420] + ("…" if len(abstract) > 420 else "")

            doi_html = (
                f'<a class="ext-link" href="https://doi.org/{doi}" target="_blank">'
                f'DOI ↗</a>'
                if doi else ""
            )
            pm_html = (
                f'<a class="ext-link" href="https://pubmed.ncbi.nlm.nih.gov/{pmid}/" target="_blank">'
                f'PubMed ↗</a>'
                if pmid else ""
            )
            src_tag = (
                '<span class="tag tag-pm">PubMed</span>'
                if src == "pubmed"
                else '<span class="tag tag-epmc">Europe PMC</span>'
            )

            ai_html = ""
            if has_ai and row.get("AI_Key_Results"):
                ai_html = f"""
<div class="ai-box">
  <div class="ai-box-header">🤖 AI Extraction</div>
  <b>Species:</b> {row.get('AI_Species','') or '—'}<br>
  <b>Compounds:</b> {row.get('AI_Compounds','') or '—'}<br>
  <b>Activities:</b> {row.get('AI_Activities','') or '—'}<br>
  <b>Key results:</b> {row.get('AI_Key_Results','') or '—'}<br>
  <b>Study type:</b> {row.get('AI_Study_Type','') or '—'}
</div>"""

            st.markdown(f"""
<div class="paper-card">
  <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;">
    <div class="paper-title" style="flex:1;">{title}</div>
    <div class="{css}" style="white-space:nowrap;">{score}&thinsp;/&thinsp;100</div>
  </div>
  <div class="paper-authors">{authors[:130]}{'…' if len(authors)>130 else ''}</div>
  <div class="paper-meta">
    <span class="tag">{journal[:50]}</span>
    <span class="tag">{year}</span>
    {src_tag}
    {doi_html}
    {pm_html}
  </div>
  <div class="paper-abstract">{abstract_short}</div>
  {ai_html}
</div>
""", unsafe_allow_html=True)
