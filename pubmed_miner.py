"""
PubMed Miner v2.0
Fetch, score, deduplicate, and export PubMed literature to a structured Excel workbook.

New in v2.0: relevance scoring, year filter, global deduplication,
5-sheet Excel output, full CLI argument support, NCBI API key support.

Author:      Marouane Takie
Affiliation: PhD Candidate | Physiology, Pharmacology & Phytochemistry
Institution: Université Sidi Mohamed Ben Abdellah (USMBA), Fès, Morocco
ORCID:       0009-0009-8621-8548
GitHub:      github.com/marouanetakie
"""

import argparse
import time
import xml.etree.ElementTree as ET
from datetime import datetime

import requests
import pandas as pd

NCBI_BASE = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"

DEFAULT_QUERIES = [
    "Erodium moschatum",
    "Reseda alba",
    "Moroccan medicinal plants antidiabetic",
    "silver nanoparticles green synthesis antioxidant",
    "UPLC-ESI-MS plant extract phenolic compounds",
]

BATCH_SIZE = 20


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=(
            "PubMed Miner v2.0 — fetch, score, and export PubMed literature "
            "to a 5-sheet Excel workbook."
        )
    )
    p.add_argument(
        "--queries", nargs="+", default=None, metavar="QUERY",
        help=(
            "One or more PubMed search queries (quote multi-word queries). "
            "Defaults to 5 built-in phytochemistry queries."
        ),
    )
    p.add_argument(
        "--max-results", type=int, default=50, metavar="N",
        help="Maximum results per query (default: 50).",
    )
    p.add_argument(
        "--year-from", type=int, default=None, metavar="YEAR",
        help="Restrict search to articles published from YEAR onwards (e.g. 2015).",
    )
    p.add_argument(
        "--relevance-threshold", type=int, default=60, metavar="N",
        help="Minimum relevance score (0–100) for the High Relevance sheet (default: 60).",
    )
    p.add_argument(
        "--api-key", default=None, metavar="KEY",
        help="NCBI API key — raises rate limit from 3 to 10 requests/second.",
    )
    p.add_argument(
        "--output", default=None, metavar="FILE",
        help="Output .xlsx filename (default: pubmed_results_YYYYMMDD_HHMMSS.xlsx).",
    )
    return p.parse_args()


# ---------------------------------------------------------------------------
# PubMed API helpers
# ---------------------------------------------------------------------------

def search_pubmed(
    query: str,
    max_results: int,
    year_from: int | None,
    api_key: str | None,
) -> list[str]:
    """Return PMIDs matching *query*, optionally filtered by publication year."""
    params: dict = {
        "db": "pubmed",
        "term": query,
        "retmax": max_results,
        "retmode": "json",
        "sort": "relevance",
    }
    if year_from is not None:
        params["datetype"] = "pdat"
        params["mindate"] = str(year_from)
        params["maxdate"] = "3000"
    if api_key:
        params["api_key"] = api_key
    resp = requests.get(NCBI_BASE + "esearch.fcgi", params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()["esearchresult"]["idlist"]


def fetch_details(pmids: list[str], api_key: str | None = None) -> list[dict]:
    """Fetch full records for *pmids* and return parsed article dicts."""
    if not pmids:
        return []
    params: dict = {
        "db": "pubmed",
        "id": ",".join(pmids),
        "retmode": "xml",
        "rettype": "abstract",
    }
    if api_key:
        params["api_key"] = api_key
    resp = requests.get(NCBI_BASE + "efetch.fcgi", params=params, timeout=60)
    resp.raise_for_status()
    return _parse_pubmed_xml(resp.text)


def _parse_pubmed_xml(xml_text: str) -> list[dict]:
    root = ET.fromstring(xml_text)
    articles = []

    for article in root.findall(".//PubmedArticle"):
        rec: dict = {}

        elem = article.find(".//PMID")
        rec["PMID"] = elem.text if elem is not None else ""

        elem = article.find(".//ArticleTitle")
        rec["Title"] = "".join(elem.itertext()).strip() if elem is not None else ""

        authors = []
        for author in article.findall(".//Author"):
            last = author.findtext("LastName", "")
            fore = author.findtext("ForeName", "")
            if last:
                authors.append(f"{last} {fore}".strip())
        rec["Authors"] = "; ".join(authors)

        elem = article.find(".//Journal/Title")
        rec["Journal"] = elem.text if elem is not None else ""

        # Year — fall back to MedlineDate when structured date is absent
        elem = article.find(".//PubDate/Year")
        if elem is None:
            elem = article.find(".//PubDate/MedlineDate")
        rec["Year"] = elem.text[:4] if elem is not None else ""

        doi = ""
        for id_elem in article.findall(".//ArticleId"):
            if id_elem.get("IdType") == "doi":
                doi = id_elem.text or ""
                break
        rec["DOI"] = doi

        # Abstract — supports structured sections with labels
        parts = []
        for text_elem in article.findall(".//AbstractText"):
            label = text_elem.get("Label", "")
            text = "".join(text_elem.itertext()).strip()
            parts.append(f"{label}: {text}" if label else text)
        rec["Abstract"] = " ".join(parts)

        kws = ["".join(kw.itertext()) for kw in article.findall(".//Keyword")]
        rec["Keywords"] = "; ".join(kws)

        mesh = [m.text for m in article.findall(".//MeshHeading/DescriptorName") if m.text]
        rec["MeSH_Terms"] = "; ".join(mesh)

        articles.append(rec)

    return articles


# ---------------------------------------------------------------------------
# Relevance scoring
# ---------------------------------------------------------------------------

def score_relevance(record: dict, query: str) -> int:
    """
    Return a 0–100 relevance score based on keyword overlap.
    Title hits count 3×, Keywords field 2×, Abstract 1×.
    Only query words longer than 3 characters are used as terms.
    """
    terms = [w.lower() for w in query.split() if len(w) > 3]
    if not terms:
        return 0
    title    = record.get("Title", "").lower()
    abstract = record.get("Abstract", "").lower()
    kw_text  = record.get("Keywords", "").lower()
    hits = (
        sum(3 for t in terms if t in title) +
        sum(2 for t in terms if t in kw_text) +
        sum(1 for t in terms if t in abstract)
    )
    max_hits = len(terms) * 6  # upper bound: all terms in all three fields (3+2+1)
    return min(100, round(hits / max_hits * 100))


# ---------------------------------------------------------------------------
# Summary sheet builder
# ---------------------------------------------------------------------------

def build_summary_df(df: pd.DataFrame, queries: list[str]) -> pd.DataFrame:
    rows = []
    for q in queries:
        sub = df[df["Queries"].str.contains(q, regex=False, na=False)]
        rows.append({
            "Query": q,
            "Records": len(sub),
            "Year_Min": sub["Year"].replace("", pd.NA).dropna().min() if len(sub) else "",
            "Year_Max": sub["Year"].replace("", pd.NA).dropna().max() if len(sub) else "",
            "Avg_Relevance_Score": round(sub["Relevance_Score"].mean(), 1) if len(sub) else 0,
        })
    rows.append({
        "Query": "TOTAL (unique records)",
        "Records": len(df),
        "Year_Min": df["Year"].replace("", pd.NA).dropna().min() if len(df) else "",
        "Year_Max": df["Year"].replace("", pd.NA).dropna().max() if len(df) else "",
        "Avg_Relevance_Score": round(df["Relevance_Score"].mean(), 1) if len(df) else 0,
    })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    args = parse_args()
    queries = args.queries or DEFAULT_QUERIES
    # NCBI allows 10 req/s with an API key, 3 req/s without
    sleep = 0.12 if args.api_key else 0.4

    # Global deduplication: one canonical record per PMID
    pmid_to_record: dict[str, dict] = {}
    pmid_to_queries: dict[str, list[str]] = {}

    for query in queries:
        print(f"\n[SEARCH] {query}")
        pmids = search_pubmed(query, args.max_results, args.year_from, args.api_key)
        print(f"  {len(pmids)} PMIDs returned")
        time.sleep(sleep)

        new_pmids   = [p for p in pmids if p not in pmid_to_record]
        known_pmids = [p for p in pmids if p in pmid_to_record]

        # For already-fetched records, record the additional query match and
        # update the relevance score if this query yields a higher value.
        for p in known_pmids:
            pmid_to_queries[p].append(query)
            new_score = score_relevance(pmid_to_record[p], query)
            if new_score > pmid_to_record[p]["Relevance_Score"]:
                pmid_to_record[p]["Relevance_Score"] = new_score

        if not new_pmids:
            print("  No new records (all already fetched).")
            continue

        for i in range(0, len(new_pmids), BATCH_SIZE):
            batch = new_pmids[i : i + BATCH_SIZE]
            end   = min(i + BATCH_SIZE, len(new_pmids))
            print(f"  Fetching records {i + 1}–{end} …")
            records = fetch_details(batch, args.api_key)
            for rec in records:
                rec["Relevance_Score"] = score_relevance(rec, query)
                pmid_to_record[rec["PMID"]] = rec
                pmid_to_queries[rec["PMID"]] = [query]
            time.sleep(sleep)

    if not pmid_to_record:
        print("\nNo records fetched — nothing to export.")
        return

    # Attach merged query list to each record
    for pmid, rec in pmid_to_record.items():
        rec["Queries"] = "; ".join(pmid_to_queries.get(pmid, []))

    COLUMNS = [
        "Queries", "PMID", "Title", "Authors", "Journal", "Year",
        "DOI", "Relevance_Score", "Abstract", "Keywords", "MeSH_Terms",
    ]
    all_df = (
        pd.DataFrame(pmid_to_record.values(), columns=COLUMNS)
        .sort_values("Relevance_Score", ascending=False)
        .reset_index(drop=True)
    )

    # Sheet 2: High Relevance
    high_rel_df = all_df[all_df["Relevance_Score"] >= args.relevance_threshold].copy()

    # Sheets 3 & 4: first two queries get dedicated per-query views (no "Queries" column)
    def query_sheet_df(q: str) -> pd.DataFrame:
        mask = all_df["Queries"].str.contains(q, regex=False, na=False)
        return (
            all_df[mask]
            .drop(columns=["Queries"])
            .reset_index(drop=True)
        )

    # Sheet 5: Summary statistics
    summary_df = build_summary_df(all_df, queries)

    output_path = args.output or f"pubmed_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Sheet 1 — full deduplicated dataset
        all_df.to_excel(writer, index=False, sheet_name="All Results")

        # Sheet 2 — high-relevance filter
        high_rel_df.to_excel(writer, index=False, sheet_name="High Relevance")

        # Sheets 3 & 4 — first two queries (primary research topics)
        for q in queries[:2]:
            sheet_name = q[:31]
            query_sheet_df(q).to_excel(writer, index=False, sheet_name=sheet_name)

        # Sheet 5 — per-query summary statistics
        summary_df.to_excel(writer, index=False, sheet_name="Summary")

    print(f"\nDone — {len(all_df)} unique records exported to {output_path}")
    print(f"  High Relevance (score ≥ {args.relevance_threshold}): {len(high_rel_df)} records")
    if args.year_from:
        print(f"  Year filter applied: ≥ {args.year_from}")


if __name__ == "__main__":
    main()
