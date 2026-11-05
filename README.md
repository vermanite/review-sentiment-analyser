# Review Sentiment Analyser

A Python desktop application built with Tkinter for analysing English text reviews using VADER and TextBlob.

## Features

- Sentiment Summary tab with a chart of positive, neutral and negative reviews.
- Review-level sentiment scores, search and sentiment filters.
- Keyword, bigram and trigram analysis with review boundaries preserved.
- Background analysis with progress updates.
- Excel export containing six report sheets.

## Setup

Use Python 3.10 or newer with Tkinter available. On Windows, use a Python installation that includes Tcl/Tk.

```sh
python -m venv .venv
```

Activate the environment in Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Install dependencies and launch:

```sh
python -m pip install -r requirements.txt
python sentiment_gui.py
```

On the first analysis, the app downloads missing NLTK stopwords and VADER data, which requires internet access. Subsequent analyses use the downloaded resources.

## Usage

1. Browse to a text file of English reviews.
2. Choose one review per line or reviews separated by blank lines.
3. Select the number of top keywords and n-grams, then click **Analyse reviews**.
4. Explore the summary chart and use the Reviews tab to search or filter. Double-click a review to read its full text.
5. Click **Export Excel** to save the complete analysis. Review filters affect the displayed table; the chart and exported workbook cover the full dataset.

The workbook contains Sentiment Summary, Figure 4 Data, Review-level Sentiment, Keywords, Bigrams and Trigrams.

## Files

- `sentiment_gui.py`: recommended app, including chart and review filters.
- `examples/sample_reviews.txt`: synthetic sample input.

Input datasets, generated Excel reports and environment files are excluded from version control.

## Interpretation

VADER labels use compound-score thresholds of +0.05 and -0.05. These are automated English-language sentiment estimates; inspect individual reviews when interpreting results. Keywords and n-grams are based on cleaned text with English stopwords removed.

## Credits

Developed by **Dr. Nitesh Kumar Verma**.  
Website: [www.niteshkverma.com](https://www.niteshkverma.com)

## License and permitted use

Copyright (c) 2026 Dr. Nitesh Kumar Verma.

This project is provided under the custom [Noncommercial Use and Redistribution License](LICENSE).

- Free to use for noncommercial purposes.
- Free to redistribute unmodified copies without charge, retaining the license, citation file, copyright notice, and author credits.
- Commercial use and paid redistribution are prohibited without separate written permission.
- Modification and redistribution of modified versions are not licensed; contact the author for permission.
- Academic and research use requires author acknowledgment and a software citation in resulting scholarly outputs.

These are custom source-available terms, not the MIT License. Third-party dependencies retain their own licenses. Consult [LICENSE](LICENSE) for the full terms.

## Required academic citation

If you use this software in academic or research work, you must credit **Dr. Nitesh Kumar Verma** and cite **Review Sentiment Analyser** in the resulting publication, thesis, report, presentation, or other scholarly output, as required by the license.

Suggested citation:

> Verma, N. K. (2026). *Review Sentiment Analyser* [Computer software]. https://www.niteshkverma.com/index

Use the metadata in [CITATION.cff](CITATION.cff) for citation tools. On GitHub, place this file in the repository root on the default branch to enable **Cite this repository**. Add the repository URL or an archived-release DOI to the citation metadata once available.

Suggested acknowledgment:

> Review sentiment analysis was conducted using Review Sentiment Analyser, developed by Dr. Nitesh Kumar Verma.
