# Section Running Time & Detention

Upload the running-time workbook, get the dashboard. Nothing to configure.

![tabs: Overview · Sections · Runs · Data]

---

## Run it on your own machine

```bash
pip install -r requirements.txt
streamlit run app.py
```

It opens at <http://localhost:8501>. Drop the `.xlsx` into the uploader in the
sidebar and the whole dashboard builds itself.

To let others on your office network use it:

```bash
streamlit run app.py --server.address 0.0.0.0
```

Then share the "Network URL" it prints. They will need to be on the same network.

---

## Put it online (free)

1. Put these three files in a GitHub repository:
   `app.py`, `requirements.txt`, `.streamlit/config.toml`
2. Go to <https://share.streamlit.io>, sign in with GitHub, click **New app**
3. Pick the repository and set the main file to `app.py`

You get a permanent link. Anyone you share it with uploads their own file — the
data is never stored, it lives only for as long as the page is open.

---

## What the sheet has to look like

Exactly what you already have:

| DATE | LP CMS ID | TRIAN NO | TO | SCMN DEP | KTV | VZM | CHE | PSA |
|---|---|---|---|---|---|---|---|---|
| 01.09.26 | SCMN 2351 | TAK/BCN | 13:15 | 14:35 | 5:12 | 10:09 | 14:52 | 19:40 |

The last four columns hold the running time **from SCMN**, so the section times
are the differences between them:

```
SCMN → KTV   KTV → VZM   VZM → CHE   CHE → PSA
```

**Nothing is tied to a column letter.** The app looks for the row carrying the
stop names and works the rest out from there, so renaming a heading or inserting
a column will not break it. It also copes with:

- dates written `dd.mm.yy`, and dates Excel stored with day and month swapped —
  it spots those against the rest of the column and says so
- the `TRIAN NO` spelling
- durations formatted past 24 hours
- a run that never reached the end

**Blank or repeated running time.** If a cell is empty, or holds the same
running time as the station before it, the train never got that far. The run is
counted as ending at the last station it actually reached — PSA when it is
there, otherwise CHE, otherwise VZM.

### A different route

Open **Route** in the sidebar and change the start and the stop list. Add a stop
and a new section appears by itself — charts, colours, filters and all. The
names have to match the headings in the sheet.

---

## What you get

**Filters** — date range, train or LP ID, section, and a "losing at least"
threshold. Click a band on the top chart to keep only those runs. Everything
stacks: section plus band plus date range all at once.

**Score cards** — runs in view, worst section, total time lost, detention before
the run starts, end-to-end average, and the single biggest hit. Pick a section
and all of them switch to that section.

**Overview** — how long the whole run took, in four bands, and the station
spread: hours up the side, stations along the bottom, with the band showing
where most runs sit and red lines for the five that lost the most time. Where
the band fans out is where runs come apart.

**Sections** — where the time is being lost, average per section against the
normal run, a section-by-section table, and a day-by-day trend.

**Runs** — one stacked bar per run, then a picker for the full detail of any
single run: ordered time, departure, detention, each leg against a normal run,
and roughly when it reached each station.

**Data** — every run in view as a table, with a CSV download, plus an expander
showing which sheet column it read for what. Check that once on a new sheet.

---

## Where the numbers come from

**Normal run** — the median time for that section across the whole file, not
just what is on screen. It stays put while you filter, so "extra minutes" always
means the same thing.

**Time lost** — for each run and each section, how far above the normal run it
was, never below zero, added up.

**Detention before start** — departure from SCMN minus the ordered time, rolled
over if the train left after midnight.

---

## If something looks wrong

Open **Data → Which column did it read for what?**. That lists every field and
the heading it came from. If one is wrong, either rename the heading in the
sheet or adjust the route in the sidebar.

The warning strip under the score cards lists anything odd the app found while
reading — unreadable dates, dates that looked swapped, and runs that never
reached the end.
