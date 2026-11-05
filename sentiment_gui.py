import re
import statistics
import threading
import queue
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def read_text_file(file_path: str) -> str:
    """Read a text file using common encoding fallbacks."""
    encodings = ("utf-8-sig", "windows-1252")
    last_error = None

    for encoding in encodings:
        try:
            with open(file_path, "r", encoding=encoding) as file:
                return file.read()
        except UnicodeDecodeError as error:
            last_error = error

    raise UnicodeDecodeError(
        "unknown", b"", 0, 1, f"Unable to decode the input file: {last_error}"
    )


def split_reviews(text: str, one_review_per_line: bool) -> list[str]:
    """Split input text into individual reviews."""
    if one_review_per_line:
        reviews = [line.strip() for line in text.splitlines() if line.strip()]
    else:
        # Reviews are separated by one or more blank lines.
        reviews = [
            re.sub(r"\s+", " ", block).strip()
            for block in re.split(r"\n\s*\n", text)
            if block.strip()
        ]

    return reviews


def clean_text_for_keywords(text: str) -> str:
    """Clean text only for keyword/ngram analysis.

    Sentiment analysis must use the original review text because VADER uses
    punctuation, capitalisation and other contextual cues.
    """
    text = text.lower()
    text = re.sub(r"\d+", " ", text)
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"_", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def classify_vader(compound_score: float) -> str:
    """Classify a review using the standard VADER thresholds."""
    if compound_score >= 0.05:
        return "Positive"
    if compound_score <= -0.05:
        return "Negative"
    return "Neutral"


def safe_mean(values: list[float]) -> float:
    return statistics.fmean(values) if values else 0.0


def safe_median(values: list[float]) -> float:
    return statistics.median(values) if values else 0.0


def safe_sample_stdev(values: list[float]) -> float:
    return statistics.stdev(values) if len(values) > 1 else 0.0


def style_header(worksheet) -> None:
    for cell in worksheet[1]:
        cell.font = Font(bold=True)


def analyse_file(file_path, one_per_line, top_n, progress):
    global Workbook, Font
    try:
        import nltk
        from nltk import ngrams
        from nltk.collocations import BigramAssocMeasures, BigramCollocationFinder
        from nltk.corpus import stopwords
        from nltk.probability import FreqDist
        from nltk.sentiment import SentimentIntensityAnalyzer
        from openpyxl import Workbook
        from openpyxl.styles import Font
        from textblob import TextBlob
    except ImportError as error:
        raise RuntimeError(
            "Missing analysis packages. Install them with:\n"
            "python -m pip install nltk textblob openpyxl"
        ) from error

    progress("Preparing language resources…")
    for resource, location in (("stopwords", "corpora/stopwords"),
                               ("vader_lexicon", "sentiment/vader_lexicon.zip")):
        try:
            nltk.data.find(location)
        except LookupError:
            if not nltk.download(resource, quiet=True):
                raise RuntimeError(f"Could not download {resource}. Check your internet connection.")

    reviews = split_reviews(read_text_file(file_path), one_per_line)
    if not reviews:
        raise ValueError("No reviews were found in the selected file.")
    english_stopwords = set(stopwords.words("english"))
    frequency_distribution = FreqDist()
    bigram_frequency = FreqDist()
    trigram_frequency = FreqDist()
    for review in reviews:
        tokens = [word for word in clean_text_for_keywords(review).split()
                  if word.isalpha() and word not in english_stopwords]
        frequency_distribution.update(tokens)
        bigram_frequency.update(ngrams(tokens, 2))
        trigram_frequency.update(ngrams(tokens, 3))
    bigram_finder = BigramCollocationFinder(frequency_distribution, bigram_frequency)
    top_bigrams = bigram_finder.nbest(BigramAssocMeasures().mi_like, top_n)
    top_trigrams = trigram_frequency.most_common(top_n)
    analyzer = SentimentIntensityAnalyzer()
    review_results = []

    for review_number, review in enumerate(reviews, start=1):
        progress(f"Analysing review {review_number:,} of {len(reviews):,}")
        vader_scores = analyzer.polarity_scores(review)
        compound = vader_scores["compound"]
        vader_class = classify_vader(compound)

        blob_sentiment = TextBlob(review).sentiment

        review_results.append(
            {
                "review_number": review_number,
                "review_text": review,
                "vader_negative": vader_scores["neg"],
                "vader_neutral": vader_scores["neu"],
                "vader_positive": vader_scores["pos"],
                "vader_compound": compound,
                "vader_class": vader_class,
                "textblob_polarity": blob_sentiment.polarity,
                "textblob_subjectivity": blob_sentiment.subjectivity,
            }
        )

    total_reviews = len(review_results)
    class_counts = {
        "Negative": sum(r["vader_class"] == "Negative" for r in review_results),
        "Neutral": sum(r["vader_class"] == "Neutral" for r in review_results),
        "Positive": sum(r["vader_class"] == "Positive" for r in review_results),
    }
    class_proportions = {
        label: count / total_reviews for label, count in class_counts.items()
    }

    compound_scores = [r["vader_compound"] for r in review_results]
    vader_negative_scores = [r["vader_negative"] for r in review_results]
    vader_neutral_scores = [r["vader_neutral"] for r in review_results]
    vader_positive_scores = [r["vader_positive"] for r in review_results]
    textblob_polarity_scores = [r["textblob_polarity"] for r in review_results]
    textblob_subjectivity_scores = [
        r["textblob_subjectivity"] for r in review_results
    ]

    mean_compound = safe_mean(compound_scores)
    median_compound = safe_median(compound_scores)
    compound_stdev = safe_sample_stdev(compound_scores)
    overall_vader_class = classify_vader(mean_compound)

    # -----------------------------
    # Excel output
    # -----------------------------
    workbook = Workbook()

    # Sheet 1: Summary suitable for reporting and Figure 4.
    summary_sheet = workbook.active
    summary_sheet.title = "Sentiment Summary"
    summary_sheet.append(["Metric", "Count", "Proportion / Mean", "Percentage"])

    for label in ("Negative", "Neutral", "Positive"):
        summary_sheet.append(
            [
                f"{label} reviews",
                class_counts[label],
                class_proportions[label],
                class_proportions[label],
            ]
        )

    summary_sheet.append(["Mean VADER compound", "", mean_compound, ""])
    summary_sheet.append(["Median VADER compound", "", median_compound, ""])
    summary_sheet.append(["SD of VADER compound", "", compound_stdev, ""])
    summary_sheet.append(["Interpretation of mean compound", "", overall_vader_class, ""])
    summary_sheet.append(
        ["Mean VADER negative component", "", safe_mean(vader_negative_scores), ""]
    )
    summary_sheet.append(
        ["Mean VADER neutral component", "", safe_mean(vader_neutral_scores), ""]
    )
    summary_sheet.append(
        ["Mean VADER positive component", "", safe_mean(vader_positive_scores), ""]
    )
    summary_sheet.append(
        ["Mean TextBlob polarity", "", safe_mean(textblob_polarity_scores), ""]
    )
    summary_sheet.append(
        [
            "Mean TextBlob subjectivity",
            "",
            safe_mean(textblob_subjectivity_scores),
            "",
        ]
    )
    summary_sheet.append(["Total reviews analysed", total_reviews, "", ""])

    style_header(summary_sheet)
    summary_sheet.freeze_panes = "A2"
    summary_sheet.column_dimensions["A"].width = 36
    summary_sheet.column_dimensions["B"].width = 14
    summary_sheet.column_dimensions["C"].width = 22
    summary_sheet.column_dimensions["D"].width = 16

    # Decimal and percentage formatting.
    for row in range(2, summary_sheet.max_row + 1):
        summary_sheet.cell(row=row, column=3).number_format = "0.0000"
        summary_sheet.cell(row=row, column=4).number_format = "0.00%"

    # Sheet 2: Exact values to use when recreating Figure 4.
    figure_sheet = workbook.create_sheet("Figure 4 Data")
    figure_sheet.append(["Metric", "Value"])
    figure_sheet.append(["Negative review proportion", class_proportions["Negative"]])
    figure_sheet.append(["Neutral review proportion", class_proportions["Neutral"]])
    figure_sheet.append(["Positive review proportion", class_proportions["Positive"]])
    figure_sheet.append(["Mean VADER compound score", mean_compound])
    figure_sheet.append(
        ["Mean TextBlob subjectivity", safe_mean(textblob_subjectivity_scores)]
    )
    figure_sheet.append(["Mean TextBlob polarity", safe_mean(textblob_polarity_scores)])
    style_header(figure_sheet)
    figure_sheet.column_dimensions["A"].width = 34
    figure_sheet.column_dimensions["B"].width = 16
    for row in range(2, figure_sheet.max_row + 1):
        figure_sheet.cell(row=row, column=2).number_format = "0.0000"

    # Sheet 3: One row per review for verification and reproducibility.
    detail_sheet = workbook.create_sheet("Review-level Sentiment")
    detail_sheet.append(
        [
            "Review No.",
            "Review Text",
            "VADER Negative",
            "VADER Neutral",
            "VADER Positive",
            "VADER Compound",
            "VADER Classification",
            "TextBlob Polarity",
            "TextBlob Subjectivity",
        ]
    )

    for result in review_results:
        detail_sheet.append(
            [
                result["review_number"],
                result["review_text"],
                result["vader_negative"],
                result["vader_neutral"],
                result["vader_positive"],
                result["vader_compound"],
                result["vader_class"],
                result["textblob_polarity"],
                result["textblob_subjectivity"],
            ]
        )

    for row in detail_sheet.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].data_type = "s"

    style_header(detail_sheet)
    detail_sheet.freeze_panes = "A2"
    detail_sheet.auto_filter.ref = detail_sheet.dimensions
    detail_sheet.column_dimensions["A"].width = 12
    detail_sheet.column_dimensions["B"].width = 80
    for column in ("C", "D", "E", "F", "H", "I"):
        detail_sheet.column_dimensions[column].width = 18
    detail_sheet.column_dimensions["G"].width = 22

    for row in range(2, detail_sheet.max_row + 1):
        for column in (3, 4, 5, 6, 8, 9):
            detail_sheet.cell(row=row, column=column).number_format = "0.0000"

    # Sheet 4: Keywords.
    keyword_sheet = workbook.create_sheet("Keywords")
    keyword_sheet.append(["Rank", "Keyword", "Frequency"])
    for rank, (word, frequency) in enumerate(
        frequency_distribution.most_common(top_n), start=1
    ):
        keyword_sheet.append([rank, word, frequency])
    style_header(keyword_sheet)

    # Sheet 5: Bigrams.
    bigram_sheet = workbook.create_sheet("Bigrams")
    bigram_sheet.append(["Rank", "Bigram", "Frequency"])
    for rank, bigram in enumerate(top_bigrams, start=1):
        bigram_sheet.append(
            [rank, " ".join(bigram), bigram_finder.ngram_fd[bigram]]
        )
    style_header(bigram_sheet)

    # Sheet 6: Trigrams.
    trigram_sheet = workbook.create_sheet("Trigrams")
    trigram_sheet.append(["Rank", "Trigram", "Frequency"])
    for rank, (trigram, frequency) in enumerate(top_trigrams, start=1):
        trigram_sheet.append([rank, " ".join(trigram), frequency])
    style_header(trigram_sheet)

    return {"workbook": workbook, "reviews": review_results,
            "counts": class_counts, "mean": mean_compound,
            "keywords": frequency_distribution.most_common(top_n),
            "bigrams": [(" ".join(b), bigram_frequency[b]) for b in top_bigrams],
            "trigrams": [(" ".join(t), count) for t, count in top_trigrams]}


class SentimentApp:
    def __init__(self, root):
        self.root = root
        self.result = None
        self.busy = False
        self.events = queue.Queue()
        root.title("Review Sentiment Analyser — Charts and Filters")
        root.geometry("1100x740")
        root.minsize(800, 580)
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Segoe UI", 21, "bold"))
        style.configure("Summary.TLabel", font=("Segoe UI", 12))
        self.footer = ttk.Frame(root)
        self.footer.pack(side="bottom", fill="x")
        ttk.Separator(self.footer, orient="horizontal").pack(fill="x")
        self.credit_frame = ttk.Frame(self.footer, padding=(22, 10))
        self.credit_frame.pack(fill="x")
        ttk.Label(self.credit_frame, text="Developed by Dr. Nitesh Kumar Verma", font=("Segoe UI", 10, "bold")).pack(side="left")
        website = ttk.Label(self.credit_frame, text="www.niteshkverma.com", foreground="#1260a0", cursor="hand2", takefocus=True)
        website.pack(side="right")
        website.bind("<Button-1>", lambda event: webbrowser.open("https://www.niteshkverma.com"))
        website.bind("<Return>", lambda event: webbrowser.open("https://www.niteshkverma.com"))
        main = ttk.Frame(root, padding=22)
        main.pack(fill="both", expand=True)
        ttk.Label(main, text="Review Sentiment Analyser", style="Title.TLabel").pack(anchor="w")
        ttk.Label(main, text="Analyse text reviews with VADER and TextBlob, then export an Excel report.").pack(anchor="w", pady=(4, 18))
        options = ttk.LabelFrame(main, text="Input and options", padding=12)
        options.pack(fill="x")
        self.path = tk.StringVar()
        ttk.Label(options, text="Reviews file").grid(row=0, column=0, sticky="w")
        self.path_entry = ttk.Entry(options, textvariable=self.path)
        self.path_entry.grid(row=0, column=1, sticky="ew", padx=10)
        self.browse_button = ttk.Button(options, text="Browse…", command=self.browse)
        self.browse_button.grid(row=0, column=2)
        options.columnconfigure(1, weight=1)
        self.format = tk.StringVar(value="line")
        formats = ttk.Frame(options)
        formats.grid(row=1, column=0, columnspan=3, sticky="w", pady=12)
        ttk.Label(formats, text="Review format:").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(formats, text="One review per line", variable=self.format, value="line").pack(side="left")
        ttk.Radiobutton(formats, text="Separated by blank lines", variable=self.format, value="block").pack(side="left", padx=16)
        ttk.Label(options, text="Top keywords / n-grams").grid(row=2, column=0, sticky="w")
        self.top = tk.StringVar(value="30")
        ttk.Spinbox(options, from_=1, to=10000, textvariable=self.top, width=10).grid(row=2, column=1, sticky="w", padx=10)
        actions = ttk.Frame(main)
        actions.pack(fill="x", pady=14)
        self.analyse_button = ttk.Button(actions, text="Analyse reviews", command=self.start)
        self.analyse_button.pack(side="left")
        self.export_button = ttk.Button(actions, text="Export Excel…", command=self.export, state="disabled")
        self.export_button.pack(side="left", padx=10)
        self.summary = tk.StringVar(value="Select a text file to begin.")
        ttk.Label(main, textvariable=self.summary, style="Summary.TLabel", wraplength=1000).pack(anchor="w", pady=(0, 12))
        tabs = ttk.Notebook(main)
        self.tabs = tabs
        tabs.pack(fill="both", expand=True)
        self.tables = {}
        for name, columns, widths in [
            ("Sentiment Summary", ("Metric", "Count", "Proportion / Mean", "Percentage"), (360, 100, 180, 140)),
            ("Reviews", ("No.", "Review text", "VADER class", "Compound", "Polarity", "Subjectivity"), (55, 480, 110, 90, 90, 100)),
            ("Keywords", ("Keyword", "Frequency"), (600, 150)),
            ("Bigrams", ("Bigram", "Frequency"), (600, 150)),
            ("Trigrams", ("Trigram", "Frequency"), (600, 150)),
        ]:
            frame = ttk.Frame(tabs)
            tabs.add(frame, text=name)
            table_row = 0
            if name == "Sentiment Summary":
                self.chart = tk.Canvas(frame, height=165, background="#ffffff", highlightthickness=0)
                self.chart.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
                self.chart.bind("<Configure>", self.draw_chart)
                table_row = 1
            elif name == "Reviews":
                toolbar = ttk.Frame(frame, padding=(4, 8))
                toolbar.grid(row=0, column=0, columnspan=2, sticky="ew")
                ttk.Label(toolbar, text="Sentiment:").pack(side="left")
                self.sentiment_filter = tk.StringVar(value="All")
                ttk.Combobox(toolbar, textvariable=self.sentiment_filter, values=("All", "Positive", "Neutral", "Negative"), state="readonly", width=12).pack(side="left", padx=(6, 16))
                ttk.Label(toolbar, text="Search reviews:").pack(side="left")
                self.search_query = tk.StringVar()
                ttk.Entry(toolbar, textvariable=self.search_query, width=24).pack(side="left", padx=6, fill="x", expand=True)
                ttk.Button(toolbar, text="Clear", command=self.clear_filters).pack(side="left", padx=6)
                self.filter_count = tk.StringVar(value="0 reviews")
                ttk.Label(toolbar, textvariable=self.filter_count).pack(side="left", padx=6)
                self.search_query.trace_add("write", self.schedule_filter)
                self.sentiment_filter.trace_add("write", self.schedule_filter)
                self.filter_job = None
                table_row = 1
            table = ttk.Treeview(frame, columns=columns, show="headings")
            for col, width in zip(columns, widths):
                table.heading(col, text=col)
                table.column(col, width=width, minwidth=50)
            vertical = ttk.Scrollbar(frame, orient="vertical", command=table.yview)
            horizontal = ttk.Scrollbar(frame, orient="horizontal", command=table.xview)
            table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
            table.grid(row=table_row, column=0, sticky="nsew")
            vertical.grid(row=table_row, column=1, sticky="ns")
            horizontal.grid(row=table_row + 1, column=0, sticky="ew")
            frame.rowconfigure(table_row, weight=1)
            frame.columnconfigure(0, weight=1)
            self.tables[name] = table
        self.tables["Reviews"].bind("<Double-1>", self.show_review)
        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.pack(fill="x", pady=(14, 5))
        self.status = tk.StringVar(value="Ready · Double-click a result to read the full review.")
        ttk.Label(main, textvariable=self.status).pack(anchor="w")
        root.after(100, self.poll)

    def browse(self):
        path = filedialog.askopenfilename(parent=self.root, title="Select reviews", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            self.path.set(path)

    def start(self):
        if self.busy:
            return
        try:
            count = int(self.top.get())
            if count < 1:
                raise ValueError()
        except ValueError:
            messagebox.showerror("Invalid option", "Enter a positive whole number for top keywords.", parent=self.root)
            return
        path = Path(self.path.get().strip())
        if not path.is_file():
            messagebox.showerror("Select a file", "Choose an existing text file.", parent=self.root)
            return
        self.source = path
        self.busy = True
        self.result = None
        self.draw_chart()
        self.filter_count.set("0 reviews")
        for table in self.tables.values():
            table.delete(*table.get_children())
        self.summary.set("Analysis in progress…")
        self.analyse_button.configure(state="disabled")
        self.export_button.configure(state="disabled")
        self.progress.start(12)
        one_per_line = self.format.get() == "line"
        threading.Thread(target=self.worker, args=(str(path), one_per_line, count), daemon=True).start()

    def worker(self, path, one_per_line, count):
        try:
            result = analyse_file(path, one_per_line, count, lambda status: self.events.put(("status", status)))
            self.events.put(("done", result))
        except Exception as error:
            self.events.put(("error", str(error)))

    def poll(self):
        latest_status = None
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "status":
                    latest_status = payload
                    continue
                latest_status = None
                self.busy = False
                self.progress.stop()
                self.analyse_button.configure(state="normal")
                if kind == "error":
                    self.summary.set("Analysis failed. Update the input or options and try again.")
                    self.status.set("Analysis failed")
                    messagebox.showerror("Analysis failed", payload, parent=self.root)
                else:
                    self.result = payload
                    self.populate()
                    self.export_button.configure(state="normal")
                    self.status.set("Analysis complete · Export Excel to save all six report sheets.")
        except queue.Empty:
            pass
        if latest_status is not None:
            self.status.set(latest_status)
        self.root.after(100, self.poll)

    def populate(self):
        result = self.result
        counts = result["counts"]
        table = self.tables["Sentiment Summary"]
        for row in result["workbook"]["Sentiment Summary"].iter_rows(min_row=2, values_only=True):
            metric, count, value, percentage = row
            if isinstance(value, (int, float)):
                value = f"{value:.4f}"
            if isinstance(percentage, (int, float)):
                percentage = f"{percentage:.2%}"
            table.insert("", "end", values=(metric, count if count is not None else "", value if value is not None else "", percentage if percentage is not None else ""))
        self.tabs.select(0)
        self.summary.set(f"{len(result['reviews']):,} reviews   |   Positive: {counts['Positive']:,}   Neutral: {counts['Neutral']:,}   Negative: {counts['Negative']:,}   |   Mean compound: {result['mean']:.4f}")
        self.apply_filters()
        self.draw_chart()
        for name in ("Keywords", "Bigrams", "Trigrams"):
            for term, count in result[name.lower()]:
                self.tables[name].insert("", "end", values=(term, count))

    def draw_chart(self, event=None):
        canvas = self.chart
        canvas.delete("all")
        if not self.result:
            canvas.create_text(20, 30, anchor="w", text="Sentiment distribution will appear after analysis.", fill="#586578", font=("Segoe UI", 11))
            return
        total = len(self.result["reviews"])
        width = max(canvas.winfo_width(), 400)
        bar_start = 110
        bar_width = max(80, width - bar_start - 200)
        canvas.create_text(16, 18, anchor="w", text=f"Sentiment distribution · All {total:,} reviews", font=("Segoe UI", 11, "bold"))
        for index, (label, color) in enumerate((("Positive", "#21875a"), ("Neutral", "#687c94"), ("Negative", "#c84b4b"))):
            count = self.result["counts"][label]
            proportion = count / total if total else 0
            y = 44 + index * 38
            canvas.create_text(16, y + 10, anchor="w", text=label, font=("Segoe UI", 10))
            canvas.create_rectangle(bar_start, y, bar_start + bar_width, y + 21, fill="#eef1f5", outline="")
            if count:
                canvas.create_rectangle(bar_start, y, bar_start + bar_width * proportion, y + 21, fill=color, outline="", tags=("sentiment_bar", label))
            canvas.create_text(bar_start + bar_width + 12, y + 10, anchor="w", text=f"{proportion:.1%}  ({count:,})", font=("Segoe UI", 10))

    def schedule_filter(self, *args):
        if self.filter_job is not None:
            self.root.after_cancel(self.filter_job)
        self.filter_job = self.root.after(180, self.apply_filters)

    def clear_filters(self):
        self.search_query.set("")
        self.sentiment_filter.set("All")
        self.apply_filters()

    def apply_filters(self):
        if self.filter_job is not None:
            self.root.after_cancel(self.filter_job)
            self.filter_job = None
        table = self.tables["Reviews"]
        table.delete(*table.get_children())
        reviews = self.result["reviews"] if self.result else []
        query = self.search_query.get().strip().casefold()
        label = self.sentiment_filter.get()
        shown = 0
        for r in reviews:
            if label != "All" and r["vader_class"] != label:
                continue
            if query and query not in r["review_text"].casefold():
                continue
            table.insert("", "end", iid=str(r["review_number"]), values=(r["review_number"], r["review_text"], r["vader_class"], f"{r['vader_compound']:.4f}", f"{r['textblob_polarity']:.4f}", f"{r['textblob_subjectivity']:.4f}"))
            shown += 1
        self.filter_count.set(f"{shown:,} of {len(reviews):,} reviews")

    def show_review(self, event):
        selected = self.tables["Reviews"].identify_row(event.y)
        if not selected or not self.result:
            return
        review = self.result["reviews"][int(selected) - 1]
        window = tk.Toplevel(self.root)
        window.title(f"Review {selected} · {review['vader_class']}")
        window.geometry("700x400")
        from tkinter.scrolledtext import ScrolledText
        text = ScrolledText(window, wrap="word", font=("Segoe UI", 11), padx=12, pady=12)
        text.pack(fill="both", expand=True)
        text.insert("1.0", review["review_text"])
        text.configure(state="disabled")

    def export(self):
        if not self.result:
            return
        path = filedialog.asksaveasfilename(parent=self.root, title="Export analysis", defaultextension=".xlsx", initialfile=f"{self.source.stem}_review_level_analysis.xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            self.result["workbook"].save(path)
        except Exception as error:
            messagebox.showerror("Export failed", str(error), parent=self.root)
            return
        self.status.set(f"Saved: {path}")
        messagebox.showinfo("Export complete", f"Report saved to:\n{path}", parent=self.root)


if __name__ == "__main__":
    root = tk.Tk()
    app = SentimentApp(root)
    root.mainloop()
