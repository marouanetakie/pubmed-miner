"""
PubMed Miner — fetches article metadata for a set of research queries
and exports everything to a timestamped Excel workbook.
"""

import time
import xml.etree.ElementTree as ET
from datetime import datetime

import requests
import pandas as pd

NCBI_BASE = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"

QUERIES = [
    "Erodium moschatum",
    "Reseda alba",
    "Moroccan medicinal plants antidiabetic",
    "silver nanoparticles green synthesis antioxidant",
    "UPLC-ESI-MS plant extract phenolic compounds",
]

MAX_RESULTS_PER_QUERY = 50
BATCH_SIZE = 20
SLEEP_BETWEEN_REQUESTS = 0.4  # NCBI allows 3 req/s without API key


def search_pubmed(query: str, max_results: int = MAX_RESULTS_PER_QUERY) -> list[str]:
    """Return a list of PMIDs matching *query*."""
    params = {
        "db": "pubmed",
        "term": query,
        "retmax": max_results,
        "retmode": "json",
        "sort": "relevance",
    }
    resp = requests.get(NCBI_BASE + "esearch.fcgi", params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()["esearchresult"]["idlist"]


def fetch_details(pmids: list[str]) -> list[dict]:
    """Fetch full records for *pmids* and return parsed article dicts."""
    if not pmids:
        return []
    params = {
        "db": "pubmed",
        "id": ",".join(pmids),
        "retmode": "xml",
        "rettype": "abstract",
    }
    resp = requests.get(NCBI_BASE + "efetch.fcgi", params=params, timeout=60)
    resp.raise_for_status()
    return _parse_pubmed_xml(resp.text)


def _parse_pubmed_xml(xml_text: str) -> list[dict]:
    root = ET.fromstring(xml_text)
    articles = []

    for article in root.findall(".//PubmedArticle"):
        rec: dict = {}

        # PMID
        elem = article.find(".//PMID")
        rec["PMID"] = elem.text if elem is not None else ""

        # Title (handles mixed-content markup like <i>)
        elem = article.find(".//ArticleTitle")
        rec["Title"] = "".join(elem.itertext()).strip() if elem is not None else ""

        # Authors
        authors = []
        for author in article.findall(".//Author"):
            last = author.findtext("LastName", "")
            fore = author.findtext("ForeName", "")
            if last:
                authors.append(f"{last} {fore}".strip())
        rec["Authors"] = "; ".join(authors)

        # Journal
        elem = article.find(".//Journal/Title")
        rec["Journal"] = elem.text if elem is not None else ""

        # Year — fall back to MedlineDate when structured date is absent
        elem = article.find(".//PubDate/Year")
        if elem is None:
            elem = article.find(".//PubDate/MedlineDate")
        rec["Year"] = elem.text[:4] if elem is not None else ""

        # DOI
        doi = ""
        for id_elem in article.findall(".//ArticleId"):
            if id_elem.get("IdType") == "doi":
                doi = id_elem.text or ""
                break
        rec["DOI"] = doi

        # Abstract (supports structured abstracts with labelled sections)
        parts = []
        for text_elem in article.findall(".//AbstractText"):
            label = text_elem.get("Label", "")
            text = "".join(text_elem.itertext()).strip()
            parts.append(f"{label}: {text}" if label else text)
        rec["Abstract"] = " ".join(parts)

        # Author keywords
        kws = ["".join(kw.itertext()) for kw in article.findall(".//Keyword")]
        rec["Keywords"] = "; ".join(kws)

        # MeSH descriptor names
        mesh = [m.text for m in article.findall(".//MeshHeading/DescriptorName") if m.text]
        rec["MeSH_Terms"] = "; ".join(mesh)

        articles.append(rec)

    return articles


def main() -> None:
    all_records: list[dict] = []
    seen_pmids: set[str] = set()

    for query in QUERIES:
        print(f"\n[SEARCH] {query}")
        pmids = search_pubmed(query)
        print(f"  {len(pmids)} PMIDs returned")
        time.sleep(SLEEP_BETWEEN_REQUESTS)

        new_pmids = [p for p in pmids if p not in seen_pmids]
        seen_pmids.update(new_pmids)

        if not new_pmids:
            print("  No new records (all already fetched).")
            continue

        for i in range(0, len(new_pmids), BATCH_SIZE):
            batch = new_pmids[i : i + BATCH_SIZE]
            end = min(i + BATCH_SIZE, len(new_pmids))
            print(f"  Fetching records {i + 1}–{end} …")
            records = fetch_details(batch)
            for rec in records:
                rec["Query"] = query
            all_records.extend(records)
            time.sleep(SLEEP_BETWEEN_REQUESTS)

    if not all_records:
        print("\nNo records fetched — nothing to export.")
        return

    df = pd.DataFrame(
        all_records,
        columns=["Query", "PMID", "Title", "Authors", "Journal", "Year",
                 "DOI", "Abstract", "Keywords", "MeSH_Terms"],
    )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"pubmed_results_{timestamp}.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="All Results")
        for query in QUERIES:
            sheet_df = df[df["Query"] == query].drop(columns=["Query"])
            # Excel sheet names are limited to 31 characters
            sheet_name = query[:31]
            sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

    print(f"\nDone — {len(df)} records exported to {output_path}")


if __name__ == "__main__":
    main()
