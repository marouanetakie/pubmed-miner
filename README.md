# PubMed Miner v2.0

Automated literature retrieval from PubMed with relevance scoring, year filtering,
global deduplication, and structured Excel export.
Uses the NCBI E-utilities API — no API key required (optional for higher throughput).

## What's new in v2.0

| Feature | Detail |
|---|---|
| **Relevance scoring** | Each article receives a 0–100 score based on keyword overlap with the query (title weighted 3×, author keywords 2×, abstract 1×) |
| **Year filter** | `--year-from YEAR` restricts searches to articles from that year onwards |
| **Global deduplication** | Each PMID appears once across all queries; the *Queries* column lists every query that matched it |
| **5-sheet Excel** | All Results · High Relevance · *Query 1* · *Query 2* · Summary |
| **Full CLI** | All parameters configurable via command-line arguments |
| **API key support** | Pass `--api-key` to raise the NCBI rate limit from 3 to 10 req/s |

## Default queries

| # | Query |
|---|-------|
| 1 | Erodium moschatum |
| 2 | Reseda alba |
| 3 | Moroccan medicinal plants antidiabetic |
| 4 | silver nanoparticles green synthesis antioxidant |
| 5 | UPLC-ESI-MS plant extract phenolic compounds |

## Output — 5-sheet Excel workbook

| Sheet | Contents |
|---|---|
| **All Results** | All deduplicated records sorted by Relevance Score (descending); includes a *Queries* column listing every query that matched each article |
| **High Relevance** | Subset of All Results with Relevance Score ≥ threshold (default 60) |
| **Erodium moschatum** | Records matching the first query, without the *Queries* column |
| **Reseda alba** | Records matching the second query, without the *Queries* column |
| **Summary** | Per-query record counts, year range, and average relevance score |

> When using `--queries`, sheets 3 and 4 are named after the first two custom queries.

## Fields collected

- PMID, Title, Authors, Journal, Year, DOI
- Relevance_Score (0–100)
- Abstract (structured sections preserved)
- Author Keywords
- MeSH Terms
- Queries (All Results sheet only)

## Setup

```bash
pip install -r requirements.txt
```

## Usage

**Default run (5 built-in queries, all options at default):**
```bash
python pubmed_miner.py
```

**Filter to articles from 2015 onwards:**
```bash
python pubmed_miner.py --year-from 2015
```

**Custom queries with a year filter:**
```bash
python pubmed_miner.py --queries "Erodium moschatum phytochemistry" "Reseda alba flavonoids" --year-from 2018
```

**Raise the relevance threshold and set a custom output filename:**
```bash
python pubmed_miner.py --relevance-threshold 75 --output erodium_lit.xlsx
```

**Use an NCBI API key for higher throughput (10 req/s):**
```bash
python pubmed_miner.py --api-key YOUR_KEY_HERE
```

### All CLI options

| Argument | Default | Description |
|---|---|---|
| `--queries QUERY [QUERY ...]` | 5 built-in | PubMed search queries (quote multi-word queries) |
| `--max-results N` | 50 | Maximum results fetched per query |
| `--year-from YEAR` | none | Minimum publication year |
| `--relevance-threshold N` | 60 | Minimum score for the High Relevance sheet |
| `--api-key KEY` | none | NCBI API key |
| `--output FILE` | timestamped | Output `.xlsx` filename |

Register for a free NCBI API key at <https://www.ncbi.nlm.nih.gov/account/>.

## Author

| | |
|---|---|
| **Name** | Marouane Takie |
| **Affiliation** | PhD Candidate \| Physiology, Pharmacology & Phytochemistry |
| **Institution** | Université Sidi Mohamed Ben Abdellah (USMBA), Fès, Morocco |
| **ORCID** | [0009-0009-8621-8548](https://orcid.org/0009-0009-8621-8548) |
| **GitHub** | [github.com/marouanetakie](https://github.com/marouanetakie) |
