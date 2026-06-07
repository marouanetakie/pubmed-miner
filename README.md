# PubMed Miner v3.0

Automated literature retrieval from **PubMed** and **Europe PMC** with relevance scoring,
SQLite caching, Claude AI structured extraction, BibTeX/RIS export, and publication charts.
No API key required (optional for higher throughput).

## What's new in v3.0

| Feature | Detail |
|---|---|
| **Europe PMC dual-source** | `--source both` fetches from PubMed *and* Europe PMC simultaneously; DOI-based cross-source deduplication merges overlapping records |
| **SQLite result cache** | All fetched records persist in `pubmed_cache.db` — re-runs are instant for already-fetched PMIDs; `--no-cache` forces a fresh fetch |
| **Claude AI extraction** | `--ai-extract` calls `claude-haiku-4-5` to extract species, compounds, activities, key quantitative results, and study type from each abstract |
| **BibTeX / RIS export** | `--export-bibtex` and `--export-ris` write citation files ready for Zotero, Mendeley, or LaTeX |
| **Publication charts** | New **Charts** sheet in the Excel workbook: publications per year, top 10 journals, relevance-score distribution |

## What was new in v2.0

| Feature | Detail |
|---|---|
| **Relevance scoring** | 0–100 score based on keyword overlap with the query (title 3×, keywords 2×, abstract 1×) |
| **Year filter** | `--year-from YEAR` restricts searches to articles from that year onwards |
| **Global deduplication** | Each PMID appears once across all queries; the *Queries* column lists every query that matched it |
| **5-sheet Excel** | All Results · High Relevance · *Query 1* · *Query 2* · Summary |
| **Full CLI** | All parameters configurable via command-line arguments |

## Default queries

| # | Query |
|---|-------|
| 1 | Erodium moschatum |
| 2 | Reseda alba |
| 3 | Moroccan medicinal plants antidiabetic |
| 4 | silver nanoparticles green synthesis antioxidant |
| 5 | UPLC-ESI-MS plant extract phenolic compounds |

## Output — 6-sheet Excel workbook

| Sheet | Contents |
|---|---|
| **All Results** | All deduplicated records sorted by Relevance Score (descending); includes a *Queries* column listing every query that matched each article |
| **High Relevance** | Subset with Relevance Score ≥ threshold (default 60) |
| **Erodium moschatum** | Records matching the first query, without the *Queries* column |
| **Reseda alba** | Records matching the second query, without the *Queries* column |
| **Summary** | Per-query record counts, year range, and average relevance score |
| **Charts** | Publication trend (bar), top 10 journals (bar), relevance distribution (bar) |

> When using `--queries`, sheets 3 and 4 are named after the first two custom queries.

## Fields collected

### Always present
- `UID`, `Source` (pubmed / europepmc)
- `PMID`, `Title`, `Authors`, `Journal`, `Year`, `DOI`
- `Relevance_Score` (0–100)
- `Abstract` (structured sections preserved)
- `Keywords`, `MeSH_Terms`
- `Queries` (All Results sheet only)

### With `--ai-extract`
- `AI_Species` — plant/organism scientific names
- `AI_Compounds` — chemical compounds or compound classes identified
- `AI_Activities` — biological activities studied
- `AI_Key_Results` — key quantitative findings with units (e.g. IC50=12.3 μg/mL vs DPPH)
- `AI_Study_Type` — `in_vitro | in_vivo | in_silico | clinical | review | mixed`

## Setup

```bash
pip install -r requirements.txt

# For AI extraction only:
pip install anthropic
```

## Usage

**Default run (5 built-in queries, PubMed only):**
```bash
python pubmed_miner.py
```

**Fetch from both PubMed and Europe PMC:**
```bash
python pubmed_miner.py --source both
```

**Filter to articles from 2015 onwards:**
```bash
python pubmed_miner.py --year-from 2015
```

**Custom queries with a year filter:**
```bash
python pubmed_miner.py --queries "Erodium moschatum phytochemistry" "Reseda alba flavonoids" --year-from 2018
```

**Enable AI extraction (requires ANTHROPIC_API_KEY):**
```bash
$env:ANTHROPIC_API_KEY = "sk-ant-..."      # PowerShell
python pubmed_miner.py --source both --ai-extract
```

**Export BibTeX and RIS alongside Excel:**
```bash
python pubmed_miner.py --export-bibtex results.bib --export-ris results.ris
```

**Raise the relevance threshold and set a custom output filename:**
```bash
python pubmed_miner.py --relevance-threshold 75 --output erodium_lit.xlsx
```

**Use an NCBI API key for higher throughput (10 req/s):**
```bash
python pubmed_miner.py --api-key YOUR_KEY_HERE
```

**Skip the cache and re-fetch everything:**
```bash
python pubmed_miner.py --no-cache
```

### All CLI options

| Argument | Default | Description |
|---|---|---|
| `--queries QUERY [QUERY ...]` | 5 built-in | PubMed/Europe PMC search queries |
| `--source pubmed\|europepmc\|both` | `pubmed` | Literature source(s) to query |
| `--max-results N` | 50 | Maximum results fetched per query per source |
| `--year-from YEAR` | none | Minimum publication year |
| `--relevance-threshold N` | 60 | Minimum score for the High Relevance sheet |
| `--ai-extract` | off | Enable Claude AI structured extraction |
| `--api-key KEY` | none | NCBI API key |
| `--cache-db FILE` | `pubmed_cache.db` | SQLite cache file path |
| `--no-cache` | off | Bypass cache and re-fetch all records |
| `--output FILE` | timestamped | Output `.xlsx` filename |
| `--export-bibtex FILE` | none | Also export as BibTeX `.bib` |
| `--export-ris FILE` | none | Also export as RIS `.ris` |

Register for a free NCBI API key at <https://www.ncbi.nlm.nih.gov/account/>.

## Author

| | |
|---|---|
| **Name** | Marouane Takie |
| **Affiliation** | PhD Candidate \| Physiology, Pharmacology & Phytochemistry |
| **Institution** | Université Sidi Mohamed Ben Abdellah (USMBA), Fès, Morocco |
| **ORCID** | [0009-0009-8621-8548](https://orcid.org/0009-0009-8621-8548) |
| **GitHub** | [github.com/marouanetakie](https://github.com/marouanetakie) |
