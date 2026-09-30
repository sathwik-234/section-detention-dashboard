"""
SECTION RUNNING TIME & DETENTION  -  Streamlit app
===================================================

Upload the running-time workbook and the whole dashboard builds itself.

    pip install -r requirements.txt
    streamlit run app.py

The sheet is expected to look like this:

    row 1 :  DATE | LP CMS ID | TRIAN NO | TO | SCMN DEP | TIME TAKEN FROM SCMN
    row 2 :                                     KTV | VZM | CHE | PSA
    row 3+:  data

KTV / VZM / CHE / PSA hold the running time FROM SCMN (cumulative), so the four
section times are the differences between them:

    SCMN -> KTV ,  KTV -> VZM ,  VZM -> CHE ,  CHE -> PSA

Nothing is tied to a column letter. The app looks for the row carrying the stop
names and works the rest out from there, so renaming a heading or inserting a
column will not break it.

A blank running time, or one that repeats the station before it, means the train
never got that far: the run is counted as ending at the last station it actually
reached.
"""

from __future__ import annotations

import datetime as dt
import io
import statistics
from dataclasses import dataclass, field

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

try:                                    # openpyxl is what actually reads the file
    import openpyxl
except ImportError:                     # pragma: no cover
    openpyxl = None


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

ORIGIN = "SCMN"
STOPS = ["KTV", "VZM", "CHE", "PSA"]

#: end-to-end running time bands, in minutes
SLOTS = [
    ("Below 10 hours", 0, 600),
    ("10 – 12 hours", 600, 720),
    ("12 – 14 hours", 720, 840),
    ("14 hours and above", 840, 10 ** 9),
]

# validated palette - the same one the other dashboards use
SEC_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300"]
SLOT_COLORS = ["#1baf7a", "#3987e5", "#eda100", "#e34948"]
ACCENT, GOOD, WARN, BAD, SEQ = "#2a78d6", "#1baf7a", "#eda100", "#e34948", "#3987e5"


def sections() -> list[str]:
    return [f"{ORIGIN if i == 0 else STOPS[i-1]} → {s}" for i, s in enumerate(STOPS)]


def points() -> list[str]:
    return [ORIGIN] + STOPS


# ---------------------------------------------------------------------------
# Reading the workbook
# ---------------------------------------------------------------------------

@dataclass
class Run:
    row: int
    date: dt.date
    lp: str
    train: str
    ordered: int | None
    dep: int | None
    det: int | None
    legs: list[int | None]
    cum: list[int | None]
    total: int | None
    reached: int                       # how many stops it actually made

    @property
    def disp(self) -> str:
        return self.date.strftime("%d-%m-%Y")

    @property
    def dow(self) -> str:
        return self.date.strftime("%a")

    @property
    def name(self) -> str:
        return self.lp or self.train or f"Row {self.row}"

    @property
    def end_station(self) -> str:
        return points()[self.reached]


@dataclass
class Book:
    runs: list[Run]
    sheet: str
    filename: str
    bench: list[int] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    columns: dict[str, str] = field(default_factory=dict)


def _norm(v) -> str:
    return "".join(ch for ch in str(v or "").lower() if ch.isalnum())


def _minutes(v) -> int | None:
    """A clock or duration cell -> minutes."""
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.hour * 60 + v.minute
    if isinstance(v, dt.time):
        return v.hour * 60 + v.minute
    if isinstance(v, (int, float)):
        f = float(v)
        return int(round((f % 1) * 1440)) if 0 <= f < 10 else None
    s = str(v).strip()
    if not s:
        return None
    bits = s.replace(".", ":").split(":")
    try:
        if len(bits) >= 2:
            return int(bits[0]) * 60 + int(bits[1])
    except ValueError:
        pass
    return None


def _text_date(v) -> dt.date | None:
    """dd.mm.yy / dd-mm-yyyy / 17 Sep 2026 -> date."""
    s = str(v or "").strip()
    if not s:
        return None
    for sep in (".", "/", "-"):
        if sep in s:
            bits = [b for b in s.split(sep) if b]
            if len(bits) == 3:
                try:
                    a, b, c = int(bits[0]), int(bits[1]), int(bits[2])
                except ValueError:
                    break
                if a > 31:                       # yyyy-mm-dd
                    y, m, d = a, b, c
                else:                            # dd-mm-yy
                    d, m, y = a, b, c
                    if y < 100:
                        y += 2000
                    if m > 12 and d <= 12:
                        d, m = m, d
                try:
                    return dt.date(y, m, d)
                except ValueError:
                    break
    try:
        return pd.to_datetime(s, dayfirst=True).date()
    except Exception:
        return None


def _locate(grid: list[list]) -> dict:
    """Find the row carrying the stop names, then the other columns above it."""
    want = [_norm(s) for s in STOPS]
    sub, cum = -1, None

    for r in range(min(len(grid), 25)):
        cols, ok = [], True
        for target in want:
            hit = -1
            for c, cell in enumerate(grid[r]):
                if _norm(cell) == target:
                    hit = c
                    break
            if hit == -1:
                ok = False
                break
            cols.append(hit)
        if ok:
            sub, cum = r, cols
            break

    if sub == -1:
        raise ValueError(
            "Could not find a row holding " + ", ".join(STOPS) +
            ". Check the stop names in the sidebar against the headings in your sheet."
        )

    taken = set(cum)

    def find(test) -> int:
        for r in range(sub + 1):
            for c, cell in enumerate(grid[r]):
                if c in taken:
                    continue
                t = _norm(cell)
                if t and test(t):
                    return c
        return -1

    out = {"header_row": sub, "cum": cum}
    for key, test in [
        ("date", lambda t: "date" in t),
        ("lp", lambda t: "cms" in t or t.startswith("lp")),
        # the sheet spells it TRIAN NO, so both spellings are accepted
        ("train", lambda t: "train" in t or "trian" in t),
        ("dep", lambda t: "dep" in t),
        # the ordered time sits under a heading of just "TO"
        ("ordered", lambda t: t == "to" or "order" in t),
    ]:
        c = find(test)
        out[key] = c
        if c != -1:
            taken.add(c)

    if out["date"] == -1:
        raise ValueError(f"A date column could not be found above row {sub + 1}.")
    return out


@st.cache_data(show_spinner=False)
def read_book(data: bytes, filename: str, origin: str, stops: tuple[str, ...]) -> Book:
    """Parse the workbook. Cached on the file's bytes, so re-filtering is instant."""
    global ORIGIN, STOPS
    ORIGIN, STOPS = origin, list(stops)

    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)

    sheet, grid = None, None
    want = {_norm(s) for s in STOPS}
    for ws in wb.worksheets:                     # the tab that carries the stop names
        g = [[c.value for c in row] for row in ws.iter_rows()]
        if any(len(want & {_norm(c) for c in row}) >= len(want) for row in g[:25]):
            sheet, grid = ws, g
            break
    if sheet is None:
        sheet = wb.worksheets[0]
        grid = [[c.value for c in row] for row in sheet.iter_rows()]

    loc = _locate(grid)
    n = len(STOPS)
    warnings: list[str] = []

    # ---- pass one ----
    raw = []
    for r in range(loc["header_row"] + 1, len(grid)):
        row = grid[r]

        def cell(idx):
            return row[idx] if 0 <= idx < len(row) else None

        rawdate = cell(loc["date"])
        if rawdate is None or str(rawdate).strip() == "":
            continue

        # Text is unambiguous (the sheet writes dd.mm.yy). A cell already stored
        # as a date can have come in with day and month the other way round,
        # which is settled below against the rest of the column.
        alt = None
        if isinstance(rawdate, (dt.datetime, dt.date)):
            d = rawdate.date() if isinstance(rawdate, dt.datetime) else rawdate
            if d.day <= 12 and d.day != d.month:
                try:
                    alt = dt.date(d.year, d.day, d.month)
                except ValueError:
                    alt = None
        else:
            d = _text_date(rawdate)
        if d is None:
            warnings.append(f"Row {r + 1}: the date “{rawdate}” could not be read, row skipped.")
            continue

        # A blank cell, or one repeating the station before it, means the train
        # never got that far.
        cum: list[int | None] = []
        legs: list[int | None] = []
        prev, stop = 0, False
        for k in range(n):
            v = None if stop else _minutes(cell(loc["cum"][k]))
            if v is None or v == prev:
                cum.append(None)
                legs.append(None)
                stop = True
                continue
            if v < prev:                        # running time past midnight
                v += 1440
            cum.append(v)
            legs.append(v - prev)
            prev = v

        if legs[0] is None:
            continue

        ordered = _minutes(cell(loc["ordered"])) if loc["ordered"] != -1 else None
        dep = _minutes(cell(loc["dep"])) if loc["dep"] != -1 else None
        det = None
        if ordered is not None and dep is not None:
            det = dep - ordered
            if det < 0:
                det += 1440                     # left after midnight

        reached = sum(1 for v in legs if v is not None)
        done = [v for v in cum if v is not None]

        raw.append({
            "row": r + 1, "date": d, "alt": alt,
            "lp": " ".join(str(cell(loc["lp"]) or "").split()) if loc["lp"] != -1 else "",
            "train": " ".join(str(cell(loc["train"]) or "").split()) if loc["train"] != -1 else "",
            "ordered": ordered, "dep": dep, "det": det,
            "legs": legs, "cum": cum,
            "total": done[-1] if done else None, "reached": reached,
        })

    if not raw:
        raise ValueError(f"No data rows were found below row {loc['header_row'] + 1}.")

    # ---- pass two: settle the dates that could be read two ways ----
    sure = [x["date"] for x in raw if x["alt"] is None]
    if sure:
        lo, hi = min(sure), max(sure)
        for x in raw:
            if x["alt"] is None:
                continue
            if not (lo <= x["date"] <= hi) and lo <= x["alt"] <= hi:
                warnings.append(
                    f"Row {x['row']}: the date was stored as {x['date']:%d-%m-%Y}, which is "
                    f"outside the rest of the sheet, so it has been read as "
                    f"{x['alt']:%d-%m-%Y} (day and month the other way round)."
                )
                x["date"] = x["alt"]

    runs = [Run(row=x["row"], date=x["date"], lp=x["lp"], train=x["train"],
                ordered=x["ordered"], dep=x["dep"], det=x["det"],
                legs=x["legs"], cum=x["cum"], total=x["total"], reached=x["reached"])
            for x in raw]
    runs.sort(key=lambda t: (t.date, t.ordered or 0))

    for t in runs:
        if t.reached < n:
            warnings.append(
                f"Row {t.row} ({t.name}): no running time past {points()[t.reached]}, "
                f"so this run is counted as ending there."
            )

    # ---- the normal run per section, over the whole file ----
    bench = []
    for s in range(n):
        vals = sorted(t.legs[s] for t in runs if t.legs[s] is not None)
        bench.append(int(round(statistics.median(vals))) if vals else 0)

    labels = {}
    hdr = loc["header_row"]
    for key in ("date", "lp", "train", "ordered", "dep"):
        c = loc[key]
        if c == -1:
            labels[key] = "not found"
        else:
            above = grid[hdr - 1][c] if hdr > 0 and c < len(grid[hdr - 1]) else None
            labels[key] = str(above or grid[hdr][c] or f"column {c + 1}").strip()
    for i, s in enumerate(STOPS):
        labels[f"to {s}"] = f"column {loc['cum'][i] + 1}"

    return Book(runs=runs, sheet=sheet.title, filename=filename,
                bench=bench, warnings=warnings, columns=labels)


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

def hm(m) -> str:
    if m is None or (isinstance(m, float) and m != m):
        return "–"
    sign = "-" if m < 0 else ""
    m = abs(int(round(m)))
    return f"{sign}{m}m" if m < 60 else f"{sign}{m // 60}h {m % 60:02d}m"


def clock(m) -> str:
    if m is None:
        return "–"
    m = int(m) % 1440
    return f"{m // 60:02d}:{m % 60:02d}"


def excess(run: Run, s: int, bench: list[int]) -> int:
    v = run.legs[s]
    return 0 if v is None else max(0, v - bench[s])


def total_excess(run: Run, bench: list[int], only: int = -1) -> int:
    if only >= 0:
        return excess(run, only, bench)
    return sum(excess(run, s, bench) for s in range(len(bench)))


def worst_leg(run: Run, bench: list[int]) -> tuple[int, int]:
    best, bx = -1, -1
    for s in range(len(bench)):
        x = excess(run, s, bench)
        if x > bx:
            best, bx = s, x
    return best, bx


def slot_of(run: Run) -> int:
    if run.total is None:
        return -1
    for i, (_, lo, hi) in enumerate(SLOTS):
        if lo <= run.total < hi:
            return i
    return -1


def pctl(sorted_vals: list[float], p: float) -> float:
    if not sorted_vals:
        return 0.0
    i = (len(sorted_vals) - 1) * p
    lo, hi = int(i), min(int(i) + 1, len(sorted_vals) - 1)
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (i - lo)


# ---------------------------------------------------------------------------
# Chart helpers
# ---------------------------------------------------------------------------

def theme() -> dict:
    """Follow whatever theme Streamlit is showing."""
    mode = "light"
    try:
        mode = st.context.theme.type or "light"
    except Exception:
        pass
    if mode == "dark":
        return {"paper": "rgba(0,0,0,0)", "ink": "#e8e6df", "muted": "#8f8e85",
                "grid": "#333331", "surface": "#1a1a19"}
    return {"paper": "rgba(0,0,0,0)", "ink": "#0b0b0b", "muted": "#82817a",
            "grid": "#e8e6e1", "surface": "#ffffff"}


def style(fig: go.Figure, height: int, **kw) -> go.Figure:
    t = theme()
    fig.update_layout(
        height=height,
        margin=dict(l=8, r=8, t=8, b=8),
        paper_bgcolor=t["paper"], plot_bgcolor=t["paper"],
        font=dict(color=t["ink"], size=13,
                  family="IBM Plex Sans, -apple-system, Segoe UI, Roboto, sans-serif"),
        hoverlabel=dict(font_size=13),
        showlegend=kw.pop("legend", False),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0,
                    bgcolor="rgba(0,0,0,0)"),
        **kw,
    )
    fig.update_xaxes(gridcolor=t["grid"], zeroline=False, linecolor=t["grid"],
                     tickfont=dict(color=t["muted"]))
    fig.update_yaxes(gridcolor=t["grid"], zeroline=False, linecolor=t["grid"],
                     tickfont=dict(color=t["muted"]))
    return fig


def hour_ticks(top: float) -> tuple[list[float], list[str]]:
    """Ticks that always land on a whole hour."""
    step = 60
    for s in (30, 60, 120, 180, 240, 360, 720):
        step = s
        if top / s <= 8:
            break
    top = max(step, (int(top // step) + 1) * step)
    vals = list(range(0, int(top) + 1, step))
    return vals, ["0" if v == 0 else f"{v // 60}h" for v in vals]


# ---------------------------------------------------------------------------
# The app
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Section Running Time & Detention",
                   page_icon="🚂", layout="wide")

st.markdown("""
<style>
  .block-container { padding-top: 2.2rem; padding-bottom: 3rem; }
  [data-testid="stMetricValue"] { font-size: 1.7rem; }
  [data-testid="stMetricLabel"] { font-size: .78rem; letter-spacing: .04em; text-transform: uppercase; }
  /* the caption under a metric is context, not a change, so no arrow */
  [data-testid="stMetricDelta"] svg { display: none; }
  [data-testid="stMetricDelta"] { padding-left: 0; }
  h1 { font-size: 1.7rem !important; }
  h2 { font-size: 1.2rem !important; padding-top: .4rem; }
  h3 { font-size: 1.02rem !important; }
</style>
""", unsafe_allow_html=True)

st.title("Section Running Time & Detention")

if openpyxl is None:
    st.error("openpyxl is missing.  Run:  pip install -r requirements.txt")
    st.stop()

# ---- sidebar: the file and the route ----
with st.sidebar:
    st.header("Data")
    upload = st.file_uploader("Running-time workbook", type=["xlsx", "xlsm"],
                              help="The sheet with DATE, LP CMS ID, TO, departure "
                                   "and the running times to each stop.")
    with st.expander("Route", expanded=False):
        origin_in = st.text_input("Start", ORIGIN)
        stops_in = st.text_input("Stops, in order", ", ".join(STOPS),
                                 help="Must match the headings in the sheet. "
                                      "Add one and a new section appears by itself.")

if upload is None:
    st.info("Upload the workbook in the sidebar to build the dashboard.")
    st.markdown(f"""
The sheet should look like this, which is the layout you already have:

| DATE | LP CMS ID | TRIAN NO | TO | {ORIGIN} DEP | {' | '.join(STOPS)} |
|---|---|---|---|---|{'---|' * len(STOPS)}
| 01.09.26 | SCMN 2351 | TAK/BCN | 13:15 | 14:35 | 5:12 | 10:09 | 14:52 | 19:40 |

The last columns hold the running time **from {ORIGIN}**, so the section times are
the differences between them. Nothing is tied to a column letter — the app finds
the row carrying the stop names and works the rest out from there.

A blank running time, or one repeating the station before it, means the train
never got that far, and the run is counted as ending where it stopped.
""")
    st.stop()

origin = origin_in.strip() or "SCMN"
stops = tuple(s.strip() for s in stops_in.split(",") if s.strip()) or tuple(STOPS)

try:
    book = read_book(upload.getvalue(), upload.name, origin, stops)
except Exception as exc:                                  # noqa: BLE001
    st.error(f"The workbook could not be read.\n\n**{exc}**")
    st.stop()

ORIGIN, STOPS = origin, list(stops)
SECS, PTS, BENCH, N = sections(), points(), book.bench, len(stops)
runs_all = book.runs

st.caption(f"**{upload.name}** · tab **{book.sheet}** · **{len(runs_all)}** runs · "
           f"{runs_all[0].disp} to {runs_all[-1].disp}")

# ---- sidebar: the filters ----
if "slot" not in st.session_state:
    st.session_state.slot = -1

with st.sidebar:
    st.header("Filters")
    lo_d, hi_d = runs_all[0].date, runs_all[-1].date
    d_from, d_to = st.date_input("Date range", (lo_d, hi_d),
                                 min_value=lo_d, max_value=hi_d,
                                 format="DD-MM-YYYY") if lo_d != hi_d else (lo_d, hi_d)

    query = st.text_input("Train / LP ID", "", placeholder="e.g. SCMN1723 or N/BOX")
    sec_pick = st.selectbox("Section", list(range(-1, len(SECS))),
                            format_func=lambda i: "All sections" if i < 0 else SECS[i])
    sec = int(sec_pick)

    thr = st.select_slider("Only runs losing at least",
                           options=[0, 30, 60, 120, 180],
                           value=0, format_func=lambda v: "any" if v == 0 else hm(v))

    if st.button("Reset all filters", width="stretch"):
        st.session_state.slot = -1
        st.rerun()


def apply_filters(skip_slot: bool = False) -> list[Run]:
    q = query.strip().lower()
    out = []
    for t in runs_all:
        if not (d_from <= t.date <= d_to):
            continue
        if q and q not in f"{t.lp} {t.train}".lower():
            continue
        if sec >= 0 and t.legs[sec] is None:
            continue
        if thr > 0 and total_excess(t, BENCH, sec) < thr:
            continue
        if not skip_slot and st.session_state.slot >= 0 and slot_of(t) != st.session_state.slot:
            continue
        out.append(t)
    return out


rows = apply_filters()
band_base = apply_filters(skip_slot=True)

if not rows:
    st.warning("No runs match these filters. Widen the date range or clear the band.")
    st.stop()

# ---- score cards ----
lost = [sum(excess(t, s, BENCH) for t in rows) for s in range(N)]
worst_s = max(range(N), key=lambda s: lost[s])
dets = [t.det for t in rows if t.det is not None]
full = [t.total for t in rows if t.reached == N and t.total is not None]
hits = [(total_excess(t, BENCH, sec), t) for t in rows]
big = max(hits, key=lambda p: p[0]) if hits else (0, None)

c = st.columns(6)
c[0].metric("Runs in view", len(rows), f"of {len(runs_all)} in the sheet",
            delta_color="off")
if sec >= 0:
    vals = [t.legs[sec] for t in rows if t.legs[sec] is not None]
    c[1].metric("Section in view", SECS[sec], f"{hm(lost[sec])} lost", delta_color="off")
    c[2].metric("Average here", hm(sum(vals) / len(vals) if vals else 0),
                f"normal {hm(BENCH[sec])}", delta_color="off")
    c[3].metric("Worst run here", hm(max(vals) if vals else 0),
                f"best {hm(min(vals) if vals else 0)}", delta_color="off")
else:
    c[1].metric("Worst section", SECS[worst_s], f"{hm(lost[worst_s])} lost",
                delta_color="off")
    c[2].metric("Total time lost", hm(sum(lost)), "against a normal run",
                delta_color="off")
    c[3].metric(f"{PTS[0]} to {PTS[-1]}", hm(sum(full) / len(full)) if full else "–",
                f"{len(full)} full runs", delta_color="off")
c[4].metric("Detention at start", hm(sum(dets) / len(dets)) if dets else "–",
            f"worst {hm(max(dets))}" if dets else "", delta_color="off")
c[5].metric("Single biggest hit", hm(big[0]) if big[1] else "–",
            f"{big[1].name} · {big[1].disp}" if big[1] else "", delta_color="off")

if book.warnings:
    with st.expander(f"⚠️  {len(book.warnings)} thing(s) to know about the sheet"):
        for w in book.warnings[:20]:
            st.write("•", w)
        if len(book.warnings) > 20:
            st.caption(f"… and {len(book.warnings) - 20} more")

# Selection events do not fire for a chart inside st.tabs, and this chart is a
# filter for the whole page rather than tab content, so it sits above them.
st.subheader("How long the whole run took")
short = sum(1 for t in band_base if t.total is not None and t.reached < N)
st.caption(
    f"Time from {PTS[0]} to the last station the train actually reached — "
    f"**click a band to keep only those runs**."
    + (f" {short} run(s) never got to {PTS[-1]}, so they are counted to where they stopped."
       if short else "")
)

counts, mins_in, short_in = [0] * len(SLOTS), [[] for _ in SLOTS], [0] * len(SLOTS)
for t in band_base:
    i = slot_of(t)
    if i < 0:
        continue
    counts[i] += 1
    mins_in[i].append(t.total)
    if t.reached < N:
        short_in[i] += 1
sized = sum(counts)

labels = [s[0] for s in SLOTS]
picked = st.session_state.slot
_dim = "#3a3a38" if theme()["ink"] == "#e8e6df" else "#d9d8d3"
# Per-point colours rather than an opacity array: plotly restyles the bars on
# selection, and an array-valued marker.opacity breaks its bar-text renderer.
bar_colors = [SLOT_COLORS[i] if picked in (-1, i) else _dim
              for i in range(len(SLOTS))]
fig = go.Figure(go.Bar(
    x=labels, y=counts,
    marker_color=bar_colors,
    text=counts, textposition="outside",
    customdata=[[i, round(c * 100 / sized) if sized else 0,
                 hm(pctl(sorted(mins_in[i]), 0.5)) if mins_in[i] else "–",
                 short_in[i]] for i, c in enumerate(counts)],
    hovertemplate="<b>%{x}</b><br>%{y} runs (%{customdata[1]}%)"
                  "<br>Typical %{customdata[2]}"
                  "<br>%{customdata[3]} stopped short<extra></extra>",
))
# stop plotly dimming the bars you did not click
fig.update_traces(selected=dict(marker=dict(opacity=1)),
                  unselected=dict(marker=dict(opacity=1)))
fig.update_layout(bargap=0.45)
fig.update_yaxes(title=None, rangemode="tozero")
ev = st.plotly_chart(style(fig, 300), width="stretch",
                     on_select="rerun", selection_mode="points", key="bands")
# A plotly selection stays put across reruns, so act only when it changes -
# otherwise the click would toggle itself straight back off.
pts = (ev.get("selection") or {}).get("points") if ev else None
if pts:
    idx = pts[0].get("point_index")
    if idx is not None and idx != st.session_state.slot:
        st.session_state.slot = idx
        st.rerun()

if st.session_state.slot >= 0:
    if st.button(f"Showing **{SLOTS[st.session_state.slot][0]}** only  ·  "
                 f"click to show every band", key="clearband"):
        st.session_state.slot = -1
        st.rerun()

st.divider()

tab_over, tab_sec, tab_runs, tab_data = st.tabs(
    ["Overview", "Sections", "Runs", "Data"])


# ===========================================================================
#  Overview
# ===========================================================================
with tab_over:
    st.subheader("Where the runs come apart")
    st.caption("Hours up the side, stations along the bottom. The band is where most "
               "runs sit — **where it fans out is where runs come apart**.")

    at = [[0.0]] + [sorted(t.cum[k] for t in rows if t.cum[k] is not None)
                    for k in range(N)]
    q = [{"n": len(v), "p10": pctl(v, .10), "p25": pctl(v, .25), "p50": pctl(v, .50),
          "p75": pctl(v, .75), "p90": pctl(v, .90)} if v else None for v in at]
    q[0] = {"n": len(rows), "p10": 0, "p25": 0, "p50": 0, "p75": 0, "p90": 0}

    xs = list(range(N + 1))
    keep = [i for i in xs if q[i]]
    hot = sorted(rows, key=lambda t: -total_excess(t, BENCH, sec))[:5]
    hot = [t for t in hot if total_excess(t, BENCH, sec) > 0]

    fig = go.Figure()
    for lo_k, hi_k, alpha in (("p10", "p90", 0.13), ("p25", "p75", 0.26)):
        fig.add_trace(go.Scatter(
            x=[xs[i] for i in keep], y=[q[i][hi_k] for i in keep],
            mode="lines", line=dict(width=0), hoverinfo="skip", showlegend=False))
        fig.add_trace(go.Scatter(
            x=[xs[i] for i in keep], y=[q[i][lo_k] for i in keep],
            mode="lines", line=dict(width=0), fill="tonexty",
            fillcolor=f"rgba(57,135,229,{alpha})", hoverinfo="skip",
            name="Middle half of runs" if lo_k == "p25" else "8 runs out of 10"))
    fig.add_trace(go.Scatter(
        x=[xs[i] for i in keep], y=[q[i]["p50"] for i in keep],
        mode="lines+markers", name="Median run",
        line=dict(color=ACCENT, width=3), marker=dict(size=9),
        customdata=[[q[i]["n"], hm(q[i]["p25"]), hm(q[i]["p75"]), hm(q[i]["p90"])]
                    for i in keep],
        hovertemplate="<b>Reaching %{x}</b><br>Median %{y:.0f} min"
                      "<br>Middle half %{customdata[1]} – %{customdata[2]}"
                      "<br>Slowest 1 in 10 over %{customdata[3]}"
                      "<br>%{customdata[0]} runs<extra></extra>"))
    for t in hot:
        ks = [0] + [k + 1 for k in range(N) if t.cum[k] is not None]
        ys = [0] + [t.cum[k] for k in range(N) if t.cum[k] is not None]
        fig.add_trace(go.Scatter(
            x=ks, y=ys, mode="lines", name=t.name,
            line=dict(color=BAD, width=2), showlegend=False,
            hovertemplate=f"<b>{t.name}</b> · {t.disp}<br>%{{y:.0f}} min<extra></extra>"))

    top = max([q[i]["p90"] for i in keep] + [max((y for t in hot for y in t.cum if y), default=0)])
    tv, tt = hour_ticks(top)
    fig.update_yaxes(tickvals=tv, ticktext=tt, rangemode="tozero")
    fig.update_xaxes(tickvals=xs, ticktext=PTS)
    st.plotly_chart(style(fig, 430, legend=True), width="stretch", key="spread")
    if hot:
        st.caption("The red lines are the "
                   f"{len(hot)} run(s) that lost the most time.")


# ===========================================================================
#  Sections
# ===========================================================================
with tab_sec:
    left, right = st.columns(2)

    with left:
        st.subheader("Where is the time being lost")
        st.caption("Extra minutes against a normal run on the same section — "
                   "**the tallest bar is your problem section**.")
        fig = go.Figure(go.Bar(
            x=lost, y=SECS, orientation="h",
            marker_color=[BAD if s == worst_s else SEQ for s in range(N)],
            marker_opacity=[0.3 if (sec >= 0 and sec != s) else 1 for s in range(N)],
            text=[hm(v) for v in lost], textposition="outside",
            textfont=dict(color=theme()["ink"]),
            hovertemplate="<b>%{y}</b><br>%{text} lost<extra></extra>"))
        fig.update_yaxes(autorange="reversed")
        fig.update_xaxes(showticklabels=False, range=[0, max(lost) * 1.28 or 1])
        fig.update_layout(bargap=0.4)
        st.plotly_chart(style(fig, 260), width="stretch", key="lost")

    with right:
        st.subheader("Average running time per section")
        st.caption("Bar is the average for the runs in view; the marker is the "
                   "normal run (median of the whole sheet).")
        avg = []
        for s in range(N):
            v = [t.legs[s] for t in rows if t.legs[s] is not None]
            avg.append(sum(v) / len(v) if v else 0)
        fig = go.Figure(go.Bar(
            x=avg, y=SECS, orientation="h",
            marker_color=SEC_COLORS[:N],
            marker_opacity=[0.3 if (sec >= 0 and sec != s) else 1 for s in range(N)],
            text=[hm(v) for v in avg], textposition="outside",
            textfont=dict(color=theme()["ink"]),
            hovertemplate="<b>%{y}</b><br>average %{text}<extra></extra>"))
        fig.add_trace(go.Scatter(
            x=BENCH, y=SECS, mode="markers", name="Normal run",
            marker=dict(symbol="line-ns", size=20, line=dict(width=3,
                        color=theme()["ink"])),
            hovertemplate="normal run %{x:.0f} min<extra></extra>"))
        fig.update_yaxes(autorange="reversed")
        fig.update_xaxes(showticklabels=False,
                         range=[0, max(avg + BENCH) * 1.28 or 1])
        fig.update_layout(bargap=0.4)
        st.plotly_chart(style(fig, 260), width="stretch", key="avg")

    st.divider()
    st.subheader("Section by section")
    tbl = []
    for s in range(N):
        v = [t.legs[s] for t in rows if t.legs[s] is not None]
        over = sum(1 for t in rows if excess(t, s, BENCH) > 0)
        tbl.append({
            "Section": SECS[s],
            "Runs": len(v),
            "Normal": hm(BENCH[s]),
            "Average": hm(sum(v) / len(v)) if v else "–",
            "Best": hm(min(v)) if v else "–",
            "Worst": hm(max(v)) if v else "–",
            "Slower than normal": f"{over} of {len(v)}" if v else "–",
            "Time lost": hm(lost[s]),
        })
    st.dataframe(pd.DataFrame(tbl), width="stretch", hide_index=True)

    st.divider()
    st.subheader("Day by day")
    days = sorted({t.date for t in rows})
    fig = go.Figure()
    for s in range(N):
        ys = []
        for d in days:
            v = [t.legs[s] for t in rows if t.date == d and t.legs[s] is not None]
            ys.append(sum(v) / len(v) if v else None)
        fig.add_trace(go.Scatter(
            x=days, y=ys, mode="lines+markers", name=SECS[s],
            line=dict(color=SEC_COLORS[s % len(SEC_COLORS)], width=2),
            marker=dict(size=7),
            hovertemplate="<b>%{x|%d-%m-%Y}</b><br>" + SECS[s] +
                          " %{y:.0f} min<extra></extra>"))
    tv, tt = hour_ticks(max([y for tr in fig.data for y in tr.y if y] or [60]))
    fig.update_yaxes(tickvals=tv, ticktext=tt, rangemode="tozero")
    st.plotly_chart(style(fig, 320, legend=True), width="stretch", key="days")


# ===========================================================================
#  Runs
# ===========================================================================
with tab_runs:
    st.subheader("Every run, section by section")
    st.caption("One bar per run, split into the sections. A long block is a slow "
               "section. Hover for the timings; pick a run below for the full detail.")

    order = st.radio("Sort by", ["Worst detention first", "Longest run first", "By date"],
                     horizontal=True, label_visibility="collapsed")
    if order == "Worst detention first":
        shown = sorted(rows, key=lambda t: -total_excess(t, BENCH, sec))
    elif order == "Longest run first":
        shown = sorted(rows, key=lambda t: -(t.total or 0))
    else:
        shown = sorted(rows, key=lambda t: (t.date, t.ordered or 0))
    shown = shown[:60]

    names, seen = [], {}
    for t in shown:                                   # keep labels unique
        base = t.name
        seen[base] = seen.get(base, 0) + 1
        names.append(base if seen[base] == 1 else f"{base} ({t.disp})")

    fig = go.Figure()
    for s in range(N):
        fig.add_trace(go.Bar(
            y=names, x=[t.legs[s] or 0 for t in shown], orientation="h",
            name=SECS[s], marker_color=SEC_COLORS[s % len(SEC_COLORS)],
            marker_opacity=0.28 if (sec >= 0 and sec != s) else 1,
            customdata=[[hm(t.legs[s]), hm(excess(t, s, BENCH))] for t in shown],
            hovertemplate="<b>%{y}</b><br>" + SECS[s] +
                          " %{customdata[0]}<br>over normal %{customdata[1]}<extra></extra>"))
    fig.update_layout(barmode="stack", bargap=0.3,
                      legend=dict(traceorder="normal"))
    tv, tt = hour_ticks(max((t.total or 0) for t in shown))
    fig.update_xaxes(tickvals=tv, ticktext=tt, side="top")
    fig.update_yaxes(autorange="reversed", tickfont=dict(size=11))
    ev = st.plotly_chart(style(fig, max(320, 26 * len(shown)), legend=True),
                         width="stretch", on_select="rerun",
                         selection_mode="points", key="runs")
    if len(rows) > 60:
        st.caption(f"Showing the first 60 of {len(rows)} runs.")

    st.divider()
    st.subheader("Run detail")
    opts = sorted(rows, key=lambda t: -total_excess(t, BENCH, sec))

    # A click on the chart moves the picker, but only the first time that click
    # is seen - after that the picker is the user's to change.
    pts = (ev.get("selection") or {}).get("points") if ev else None
    if pts:
        i = pts[0].get("point_index")
        if i is not None and i != st.session_state.get("_run_click") and i < len(shown):
            st.session_state._run_click = i
            if shown[i] in opts:
                st.session_state.run_pick = opts.index(shown[i])

    if st.session_state.get("run_pick", 0) >= len(opts):
        st.session_state.run_pick = 0

    pick = st.selectbox(
        "Run", list(range(len(opts))), key="run_pick",
        format_func=lambda i: (
            f"{opts[i].name} · {opts[i].disp} · {hm(opts[i].total)}"
            + (f"  (+{hm(total_excess(opts[i], BENCH, sec))})"
               if total_excess(opts[i], BENCH, sec) else "")))
    chosen = opts[pick]

    m = st.columns(5)
    m[0].metric("Ordered", clock(chosen.ordered))
    m[1].metric(f"Left {PTS[0]}", clock(chosen.dep))
    m[2].metric("Detention", hm(chosen.det))
    m[3].metric("Total run", hm(chosen.total))
    m[4].metric("Time lost", hm(total_excess(chosen, BENCH)))

    detail, arrive = [], chosen.dep
    for s in range(N):
        v = chosen.legs[s]
        d = None if v is None else v - BENCH[s]
        at_t = clock(arrive + v) if (v is not None and arrive is not None) else "–"
        if v is not None and arrive is not None:
            arrive += v
        detail.append({
            "Section": SECS[s],
            "Time": hm(v),
            "Normal": hm(BENCH[s]),
            "Against normal": ("–" if d is None else
                               ("+" if d > 0 else "") + hm(d)),
            f"Reached": PTS[s + 1] if v is not None else "not reached",
            "About": at_t,
        })
    st.dataframe(pd.DataFrame(detail), width="stretch", hide_index=True)

    ws, wx = worst_leg(chosen, BENCH)
    if wx > 0:
        st.info(f"**{SECS[ws]}** is where this run lost the most time — "
                f"**{hm(wx)}** over a normal run.")
    st.caption(f"{chosen.train or 'no train number'} · ended at "
               f"{chosen.end_station} · sheet row {chosen.row}")


# ===========================================================================
#  Data
# ===========================================================================
with tab_data:
    st.subheader("All runs in view")
    recs = []
    for t in rows:
        ws, wx = worst_leg(t, BENCH)
        rec = {
            "Date": t.disp, "LP CMS ID": t.lp, "Train": t.train,
            "Ordered": clock(t.ordered), f"{PTS[0]} dep": clock(t.dep),
            "Detention": hm(t.det),
        }
        for s in range(N):
            rec[SECS[s].replace(" → ", "-")] = hm(t.legs[s])
        rec["Total"] = hm(t.total)
        rec["Ended at"] = t.end_station
        rec["Lost"] = hm(total_excess(t, BENCH)) if total_excess(t, BENCH) else "–"
        rec["Worst section"] = SECS[ws] if wx > 0 else "–"
        rec["Row"] = t.row
        recs.append(rec)
    df = pd.DataFrame(recs)
    st.dataframe(df, width="stretch", hide_index=True, height=560)

    st.download_button("Download this table as CSV",
                       df.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"section_running_time_{d_from}_to_{d_to}.csv",
                       mime="text/csv")

    with st.expander("Which column did it read for what?"):
        st.caption("If any of these is wrong, rename the heading in the sheet or "
                   "adjust the route in the sidebar.")
        st.dataframe(pd.DataFrame(
            [{"Field": k, "Read from": v} for k, v in book.columns.items()]),
            width="stretch", hide_index=True)
