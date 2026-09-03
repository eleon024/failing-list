# Failing list

A grade-report reader for teachers. Drop in a stack of Individual Student Report
PDFs, get back every grade under your cutoff with the class and teacher attached,
and export it to Excel or CSV.

**Live: https://eleon024.github.io/failing-list/**

## Student data never leaves the computer

The page has no backend. PDFs are read in the browser with pdf.js — nothing is
uploaded, nothing is stored, no analytics. Closing the tab clears everything.
That is the whole reason this is a static page rather than a web service: student
grades are education records, and a tool that shipped them to a server would put
that on whoever runs the server.

The site is served over HTTPS by GitHub Pages and loads two libraries from public
CDNs (pdf.js, SheetJS). Those fetch code *in*; they never send data *out*.

## Deploying your own copy

1. Fork this repo, or create a new one and drop `index.html` in the root.
2. **Settings → Pages → Source: Deploy from a branch**, branch `main`, folder `/ (root)`.
3. Save. First build takes about a minute; the URL appears at the top of that page.

That's the whole deploy. No build step, no dependencies to install — `index.html`
is self-contained.

To use a custom domain, add a `CNAME` file containing the domain and point a
CNAME record at `YOUR-USERNAME.github.io`.

## Using it

- Drop in one PDF, a hundred, or a whole folder at once.
- **Failing at or below** — the cutoff. Defaults to 79 (at-risk). Set it to 69
  for actual failures.
- **Term** — defaults to the most recent grading period for each class. Switch to
  a specific term or to all terms.
- **Download Excel** gives four sheets: Failing, By Student, By Teacher, All Grades.

### If a file comes back empty

The reader needs selectable text. If your export is a scan or an image, there is
nothing to read and the file gets listed as skipped. Re-export it from the
gradebook rather than scanning a printout.

## What it reads

Reports laid out like this, which is what PowerSchool/HISD exports look like:

```
Individual Student Report Lastname, Firstname
Class: 5(B) AP PRECal A (Semester 1) Teacher: Xiong
Final Grade
Rpt. Term Grade Percent Absent Tardy Missing Late Incomplete
P1 73 73% 1 0 0 2 0
```

A student's report can run several pages, and only the first page carries the
name. The reader holds the current student and carries it forward across page
breaks — and across file breaks, so it still works if each page was saved
separately.

Non-numeric grades (`I` for incomplete, and so on) are kept in the `grade_raw`
column but skipped by the cutoff, so an incomplete never gets counted as a zero.

If your district's layout differs, the patterns are the four regexes at the top of
the `parsing` section in `index.html`.

## Batch use

`scripts/scrape_grades.py` does the same job from the command line, for when you
have hundreds of files or want it on a schedule.

```bash
pip install pdfplumber pandas openpyxl
python scripts/scrape_grades.py ./reports -o ./out --threshold 79
```

Writes `failing_report.xlsx`, `failing_report.csv`, and a standalone
`failing_dashboard.html`. Run `--help` for term filtering and recursion.

## License

MIT. Use it, fork it, change the cutoff, rip out the parts you don't need.
