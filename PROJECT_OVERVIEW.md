# HomeBP Insight — Project Overview & Reconstruction Notes

> **Purpose of this document.** It is a plain-language record of what this project is,
> how it is built, what clinical logic it encodes, and what decisions remain open.
> It exists so that the project can be understood and continued even if the original
> working chat session is lost. It is written for a clinician-informaticist owner
> (technical, but not a professional software developer) and for any AI assistant
> picking the work back up cold.
>
> Last reviewed: 2026-06-14.

---

## 1. What this tool is

**HomeBP Insight** is a small, privacy-first web application that turns a patient's
home blood-pressure log (an Excel or CSV export from a home BP device/app, e.g. Omron)
into an interactive timeline and analysis, plotted against the patient's medication
history.

It answers the question a clinician actually asks in clinic: *"How has this patient's
home BP responded to each change in their antihypertensive regimen?"*

**Audience / scope (decided 2026-06-14):** This is a **general-purpose tool for any
clinician managing any patient with hypertension** — not a single-patient personal
tool. That has an important consequence: any logic that assumes one specific patient's
drugs or doses is a defect to be removed (see §8, "Generalization backlog").

**Privacy model:** Everything runs inside the user's web browser. The spreadsheet is
read and analyzed locally; **no patient data is ever uploaded to a server.** The only
persistence is the browser's own local storage on the user's machine.

---

## 2. Two generations of the project

### Generation 1 — Python / PDF report generator (historical)
The project began (early–mid February 2026) as a series of Python + matplotlib scripts
that read the Excel export and produced a static PDF chart: two stacked panels
(Morning and Afternoon/Evening), color-coded BP bars, a heart-rate line, dose-shaded
medication bars, medication-hold stars, and a statistics page.

These files (`blood_pressure_analysis.py`, the `BP_Timeline_v1_0_*.py` series, the
`_v79`/`_v81` scripts, and `files-2/`) are **superseded** by the web app and are
**excluded from version control** (`.gitignore` ignores `*.py`, `*.pdf`, `*.xlsx`).
They remain on disk as a historical reference only.

> **Open question (#2 from the review):** Is the Python/PDF path formally retired, or
> should it be maintained alongside the web app? **Current assumption: retired** —
> the web app's print-to-PDF feature replaces it. Confirm if that's wrong.

### Generation 2 — the web application (current, version-controlled)
This is the live product. It is a static website with no build step and no backend —
just HTML and JavaScript files served directly. See §3.

---

## 3. Architecture of the web app

Three moving parts. The key design idea: **all the analytical logic lives in one shared
file, and two different "report templates" call into it.** This keeps the clinical rules
in a single place instead of being copy-pasted into each page.

| File | Role | Analogy |
|------|------|---------|
| [`bp-core.js`](bp-core.js) | The **engine**. Reads the spreadsheet, classifies each reading, parses the medication history from free-text notes, does all statistics and dose math. Contains **no** screen/display code. | The shared lab + rules module |
| [`clinicianBP.html`](clinicianBP.html) | The **clinician dashboard** (~3,200 lines). Full analysis. | The detailed consult note |
| [`patientBP.html`](patientBP.html) | The **patient view** (~985 lines). Simplified, large fonts. | The patient handout |
| [`index.html`](index.html) | A landing page linking to the two views. | The front door |

External dependency: **SheetJS (`xlsx`)**, loaded from a CDN, which reads the Excel file
format. That is the only third-party library.

### Data flow (what happens when a file is loaded)
1. User drops/selects an Excel or CSV file (or it auto-restores from local storage).
2. The engine finds the header row and detects which columns are Date, Time, Systolic,
   Diastolic, Pulse, and Notes — tolerant of many naming variants.
3. Each row becomes a **reading**: timestamp, SBP, DBP, pulse, notes, and a
   time-of-day window (Morning vs Afternoon/Evening).
4. Readings are cleaned: out-of-range values dropped, exact duplicates removed, and
   readings within 10 minutes of each other averaged into one.
5. Each reading is **classified** (normal / above-goal / severe / low) — see §4.
6. The **Notes column is parsed** to reconstruct the medication timeline — see §5.
7. **Hold days** (notes like "hold meds") are detected and the medication bars are
   split around them.
8. The two pages render charts, tables, and (clinician only) narrative insights.

---

## 4. Clinical logic & thresholds

All defined in [`bp-core.js`](bp-core.js) near the top.

### Reading classification
Each reading is sorted into one of four mutually-relevant categories:

- **Very High** — SBP ≥ 160 **or** DBP ≥ 100. (`SEVERE_SBP_MIN`, `SEVERE_DBP_MIN`)
  *Displayed as "Very High" — renamed from "Severe" so it isn't confused with the
  guideline 180/120 "severe/crisis" definition. The internal code identifiers
  (`r.severe`, `pctSevere`, `SEVERE_SBP_MIN`, etc.) keep the `severe` name.*
- **Low (hypotensive)** — SBP < 100 **or** DBP < 60. (`HYPO_SBP_MAX`, `HYPO_DBP_MAX`)
- **At goal (green)** — below the *selected goal* (see below), and not severe/low.
- **Above goal (red)** — everything else (not green, not severe, not low).

> **Important nuance:** "At goal" and "above goal" move with the **goal you select**,
> but "severe" and "low" are **fixed absolutes** (160/100 and 100/60). That is by
> design — severe and hypotensive are clinical safety flags, not goal-relative.

### Goal options
Selectable in the clinician view (`getGoal`):
- **US (default):** 130/80
- **WHO:** 140/90
- **Custom:** any values the clinician types (defaults 135/85 if blank)

The patient view is **hardcoded to US 130/80** with no selector (intentionally simple).

### Response tiers (how much improvement "counts")
When comparing one regimen phase to another, the change in BP is bucketed into tiers
(e.g. "significant / moderate / mild improvement", "flat", "increased"). The cutoffs
are roughly: a ≥10 mmHg drop = significant, ≥5 = moderate, any drop = mild. The
clinician version uses a finer 5-level scale ("minimal / modest / meaningful /
significant / major").

> **Decision (#3 from the review):** Thresholds should ship with **sensible defaults
> but be adjustable per patient.** Implemented for the **goal** and now for **Severe
> and Low** — the clinician view has two GOAL-style dropdowns (SEVERE and LOW, right
> after GOAL), each offering **Standard** (the built-in default) or **Custom**, where
> Custom reveals editable SBP/DBP inputs. Changing them
> flows through everything (bar colors, the legend's parenthetical definitions, the
> severe/low percentages, the clinical insights, the phase table, and the regimen
> score) via `reclassifyForGoal(readings, goal, thresholds)` in `bp-core.js`. Persisted
> in localStorage (`hbp_severeSys/Dia`, `hbp_lowSys/Dia`). The **response-tier** cutoffs
> (how much improvement counts as "significant" etc.) are still fixed in code — that
> remains on the backlog. The patient view keeps the fixed defaults (no config there).

---

## 5. Medication timeline parsing

This is the cleverest and most fragile part. The clinician records medication changes
as **free text in the Notes column** of the BP log (e.g. *"started valsartan 80 mg
qd"*, *"increased valsartan to 160"*, *"stopped candesartan"*, or a bare *"metoprolol
12.5 mg qd"*). The engine reads that prose and reconstructs a structured timeline:
which drug, what daily dose, from when to when.

It handles: start / stop / increase / decrease / titrate language; dose-frequency
abbreviations (qd, bid, tid, qid, qhs); converting "80 mg bid" into a 160 mg/day total;
and drug-name typos. Each medication is then drawn as a horizontal bar whose **color
identifies the drug** and whose **shading encodes dose intensity**: the shade is
anchored to each drug's **lowest observed dose** for this patient (lightest) and darkens
across a **4-fold dose range** (e.g. 75→150→225→300 mg → light→dark). Because it's
anchored to the lowest observed dose rather than a fixed starting dose, a dose
*decrease* correctly shows as a **lighter** band, so a clinician can eyeball dose
intensity — up or down — at a glance. In addition, every **dose change** is marked with
a divider tick and a **dose label** (shown as a callout above the bar when the segment
is too narrow to label inline, e.g. a short or very recent regimen).

> **Typo reconciliation (important):** a single-character misspelling of a drug name
> (e.g. "irbesart**e**n" vs "irbesart**a**n") would otherwise create a *phantom second
> drug* that never reconciles with the correctly-spelled one — so both the old and new
> dose appear active at the same time (a row reading "Irbesartan 225 + Irbesarten 150").
> The engine now folds drug names that differ by at most one edit (within a patient's
> own notes) into the more frequently used spelling, so a typo can't split one drug
> into two. See `canonicalizeDrugNames` / `withinOneEdit` in `bp-core.js`.

> **Generalization concern:** A couple of spots assume specific drugs/doses (see §8).
> Because the tool is now general-purpose, these need to come out.

---

## 6. The clinician dashboard, section by section

- **Timeline** — the main visualization: two BP panels (Morning, Afternoon/Evening)
  with color-coded bars, heart-rate line, goal/target lines, severe-reading labels,
  and hold-day stars; plus a medication-bar panel beneath, and a "recent N-week
  average" summary card. **Zoom:** scroll the wheel over any chart to zoom the time
  axis, centered on the cursor; the three panels stay aligned because they share one
  time window. (The page itself still scrolls when the pointer is in the margins
  outside the charts.) When zoomed in, **click-drag** pans the view sideways (cursor
  shows a grab hand), and **double-click** resets to the full range. (Available in both
  the clinician and patient views.)
- **Medication Response** (first section titled *"History of Response to Medication
  Regimens (Morning vs Afternoon/Evening)"*) — two tables:
  - a **phase table** showing BP statistics for each medication regimen period; and
  - a **Treatment Plan Comparison** that ranks the regimens by a score:
    **(%at goal) + ½ × (%near goal) − (%very high) − 3 × (hold count)**, higher = better.
    Near-goal readings earn half credit; very-high readings and medication holds are
    penalties (the 3× hold weight is a tunable heuristic, not a validated constant).
    The formula legend is shown top-right of the section, and the best-scoring regimen
    is flagged (with a "consider revisiting" note if it isn't the current one).
  - *Dose-change reconciliation:* the **current** regimen's statistics use the recent
    window **clipped to the most recent dose change**, so readings taken on the prior
    dose are never counted under the new regimen. (Historical phases were already
    clipped to their own date range.)
- **Clinical Insights** — a profile-aware narrative. The clinician can fill in a
  **patient profile** (age, sex, race, comorbidities — diabetes, CKD, heart failure,
  CAD, stroke/TIA, afib — plus conditions, intolerances, pregnancy). The engine then
  produces both a **description** of the BP pattern and **advisory suggestions** about
  drug classes to consider or avoid, each tied to a guideline/trial citation.
- **Embedded Rules** — a self-documenting page that prints out every threshold and
  rule the tool uses, so the logic is transparent and auditable.

> **Decision (#4 from the review):** The insights engine is intended to be **both
> descriptive and advisory.** Because it gives treatment-class advice, the advisory
> statements warrant clinical sign-off (see the two flagged items in §8).

### Citations
The insights cite real, verified PubMed IDs (SPRINT, the Ettehad 2016 Lancet
meta-analysis, PROGRESS, UKPDS, ALLHAT, KDIGO, PATHWAY-2, the HFrEF beta-blocker
trials, SHEP, the J-curve literature, MAPEC/Hygia, AHA/ACC 2017, etc.). Each links out
to PubMed.

---

## 7. The patient view

A deliberately stripped-down, large-font version for handing to patients:
- The same charts, but no goal selector, no profile, no narrative, no regimen tables.
- A big "recent 2-week average" with plain-language status (At Goal / Above Goal /
  Very High / Low).
- Hardcoded US 130/80 goal.
- Print stylesheet produces a clean one-page handout.

---

## 8. Known issues & backlog

### A. Generalization (because the tool is now general-purpose)
- **Valsartan dosing rule — FIXED (2026-06-14).** `medSigLabel` in `bp-core.js`
  previously assumed any valsartan dose >80 mg/day was given as "80 mg bid" — a
  one-patient assumption. It is now drug-agnostic: it shows the total daily dose
  ("160 mg/day") only as a fallback, and everywhere a dose label is shown the tool
  now prefers the actual frequency the clinician typed in the Notes (e.g. "20 mg bid").
- **`STARTING_DOSES`** table (valsartan, candesartan, irbesartan, metoprolol) is **not**
  patient data — a small reference table of standard starting doses. It is **no longer
  used for the med-bar shading** (shading now anchors to each drug's lowest *observed*
  dose, per §6), so `getStartingDose`/`STARTING_DOSES` are now effectively unused except
  for the Embedded Rules display. Could be removed in a future cleanup.

### B. Consistency issues ("the same rule written in two places")
- **Morning/evening cutoff:** the clinician page lets you configure when "morning"
  ends, but the timeline's hold-star placement always assumes noon, while the phase
  table respects the setting. If the cutoff is ever changed, the two disagree.
- **Patient page** recomputes "is this at goal?" separately from the engine. Both use
  130/80 today so it's not wrong, but it's a second copy that could drift.
- **"Near goal" (within 2 of goal)** logic is written out several times in the
  clinician page; should be a single shared helper.

### C. Leftover code
- One branch in the insights status card asks a question and then does the same thing
  either way — a vestige of a removed distinction. Harmless, should be cleaned.

### D. Clinical-judgment items to confirm (not bugs — decisions)
- **Bedtime-dosing recommendation** leans on MAPEC/Hygia, which are contested; the more
  recent **TIME trial found no benefit** and is not cited. Consider softening, or
  presenting the counter-evidence.
- **Hold-day penalty** in the regimen-ranking score weights each medication hold as
  heavily as 3 percentage-points of at-goal readings (set by the clinician owner,
  reduced from 5). This coefficient is a judgment call with no citation.

### E. Patient-view polish
- No "your data is N days old" staleness note (the "recent 2-week" window floats off
  the last reading, not today's date).
- "Near Goal" status appears in the AM/PM bar but isn't explained in the legend.
- Charts have no text/screen-reader alternative (accessibility gap).

---

## 9. Deployment

The site is hosted on **GitHub Pages** in two separate repositories — a development
site and a production site. Work happens on the `dev` branch.

- **Dev site** → https://tang75.github.io/HomeBP-dev/
  Push with: `git push dev-pages dev:main`
  (remote `dev-pages` = repo `HomeBP-dev`, Pages serves from its `main`)
- **Production site** → https://tang75.github.io/HomeBP/
  Promote with: `git checkout main && git merge dev && git push origin main && git checkout dev`
  (remote `origin` = repo `HomeBP`)

> Pushing `dev` to `origin` does **not** update the dev site — it must go to the
> `dev-pages` remote. Only update production when the owner explicitly asks.
>
> After any push, GitHub's cache can serve stale files; review with a cache-busting
> URL parameter, e.g. `…/clinicianBP.html?v=8` (increment the number each time).
> The HTML currently loads the engine as `bp-core.js?v=7` — bump that query number
> whenever `bp-core.js` changes so browsers fetch the new version.

---

## 10. Open decisions, summarized

| # | Topic | Status |
|---|-------|--------|
| 1 | Audience | **General-purpose, for any clinician/patient** (decided) |
| 2 | Python/PDF generation | Assumed **retired**; confirm |
| 3 | Clinical thresholds | Goal + **Severe/Low now adjustable** (clinician view); response-tiers still fixed |
| 4 | Insights engine role | **Both descriptive and advisory** (decided) — advisory items need sign-off |
| 5 | Immediate next step | This document; then cleanup vs. feature work, owner's call |
