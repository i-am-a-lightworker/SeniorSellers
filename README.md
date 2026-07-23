# Property Sale Signal Researcher

A local, API-key-free Streamlit tool for reviewing **property-level public sale signals** while preserving the output of a Cease and Desist compliance check.

## What it does

- Uploads Excel or CSV address lists.
- Accepts the export from the Cease and Desist Checker.
- Automatically blocks every row already marked `CEASE AND DESIST`.
- Creates manual search links for each address.
- Optionally fetches user-supplied public URLs when the site's `robots.txt` permits access.
- Scores only public sale-related terms such as `for sale`, `coming soon`, `open house`, and `FSBO`.
- Ignores sensitive hardship terms and does not use them for scoring.
- Exports a color-coded Excel report.

## Important limitation

This tool does **not** crawl the entire internet or scrape social-media profiles. It is intentionally limited to public property-level information and user-supplied URLs. Many websites prohibit automated scraping, and personal profiling for real-estate solicitation can create serious privacy, fair-housing, licensing, and platform-compliance risks.

## Installation on Windows

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Recommended workflow

1. Run your direct-mail list through the Cease and Desist Checker.
2. Upload that result here.
3. Map the address and compliance-status columns.
4. Optionally include a column of public property URLs you are authorized to review.
5. Run the review.
6. Manually inspect generated search links for rows without a supplied URL.
7. Never contact rows marked `CEASE AND DESIST`.

## Safe signal sources

- Public active listings
- Public coming-soon pages
- Public open-house announcements
- Owner-published FSBO pages
- Your own authorized MLS/CRM exports

## Do not use

- Personal social-media posts
- Health, family, ethnicity, religion, disability, age, or other protected information
- Divorce, death, bankruptcy, foreclosure, tax delinquency, job loss, or hardship signals
- Data obtained by bypassing login walls, robots.txt, CAPTCHAs, or website terms

This software is a research and workflow aid, not legal advice.
