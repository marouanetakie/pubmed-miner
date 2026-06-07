# PubMed Miner

Automated literature retrieval from PubMed for phytochemistry and related topics.
Uses the NCBI E-utilities API (no API key required).

## Queries

| # | Query |
|---|-------|
| 1 | Erodium moschatum phytochemistry |
| 2 | Reseda alba biological activity |
| 3 | Moroccan medicinal plants antidiabetic |
| 4 | silver nanoparticles green synthesis antioxidant |
| 5 | UPLC-ESI-MS plant extract phenolic compounds |

Up to **50 results per query** are fetched; duplicates across queries are deduplicated
(each PMID appears only once in the sheet for that query, and once in "All Results").

## Fields collected

- PMID, Title, Authors, Journal, Year, DOI
- Abstract (structured sections preserved)
- Author Keywords
- MeSH Terms

## Output

A timestamped Excel workbook: `pubmed_results_YYYYMMDD_HHMMSS.xlsx`

- **All Results** sheet — every record with a "Query" column
- One additional sheet per query (first 31 characters of the query string)

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
python pubmed_miner.py
```

The script respects NCBI's rate limit (≤ 3 requests/second without an API key).
To increase throughput, register for a free NCBI API key at
<https://www.ncbi.nlm.nih.gov/account/> and add it as the `api_key` parameter
in the request calls, which raises the limit to 10 req/s.

## Author

| | |
|---|---|
| **Name** | Marouane Takie |
| **Affiliation** | PhD Candidate \| Physiology, Pharmacology & Phytochemistry |
| **Institution** | Université Sidi Mohamed Ben Abdellah (USMBA), Fès, Morocco |
| **ORCID** | [0009-0009-8621-8548](https://orcid.org/0009-0009-8621-8548) |
| **GitHub** | [github.com/marouanetakie](https://github.com/marouanetakie) |
