# Grade 6 ELA — Peer Revision Task Practice Simulation

A browser-based practice simulation of the Georgia Milestones **Peer Revision Task (PRT)** for Grade 6 ELA, with an automated teacher reporting workbook for long-range instructional planning and individual intervention.

Students read a peer's flawed draft article and two sources, then answer 23 items about how to revise it. Results post automatically to a Google Sheets workbook that produces class analytics, item analysis, intervention groupings, and printable individual student reports.

---

## Contents

| Path | What it is |
|---|---|
| `index.html` | The complete student-facing simulation. Single file, no build step, no dependencies. |
| `apps-script/Code.gs` | Google Apps Script: results endpoint + report builder for the teacher workbook. |
| `docs/ALIGNMENT.md` | Item-by-item alignment map, blueprint evidence, and psychometric notes. |

---

## Alignment to the source-of-truth documents

Built against five GaDOE documents: the Grade 6 ELA Standards, the Grade 6 ELA Assessment Blueprint (Aug 2025), the Grade 6 ELA Achievement Level Descriptors (Oct 2025), the Grade 6 Classroom Peer Revision Guidance (Feb 2026), and the 8-Point 3-Trait Writing Rubric.

**Task model.** Per the Peer Revision Guidance, the PRT presents *a peer's draft text* and asks which revision would improve it. The draft in this simulation is a deliberately flawed Grade 6 student article with 21 numbered sentences, so every item can anchor to specific text — matching the state's sample item stems (`Which sentence should be added / revised / removed…`).

**Reporting categories.** Results report on **both** lenses teachers actually need:
- the **three rubric traits** (Purpose & Organization, Evidence & Elaboration, Language Usage & Conventions), and
- the **blueprint reporting categories** (Structure & Style, Techniques, Research & Analysis, Grammar Conventions, Vocabulary).

**Blueprint conformance.**

| Measure | Blueprint target | This form |
|---|---|---|
| DOK 1 | 3–10% of points | 2 pts — **6.7%** ✓ |
| DOK 2 | 45–60% of points | 16 pts — **53.3%** ✓ |
| DOK 3 | 35–50% of points | 12 pts — **40.0%** ✓ |
| Item mix | 1-pt SR/TEI + 2-pt TEI | 16 one-point + 7 two-point = **30 points** ✓ |
| Partial credit | 2-point TEIs | All seven 2-point items award partial credit ✓ |

**Scope note.** This is a *PRT-focused* form, not a full 60-point blueprint replica. A peer-revision task cannot assess Context (6.T.C) or Periods & Movements (6.T.PM), so those categories are intentionally absent. The form emphasizes the Constructing Texts claim, which is what the PRT measures.

---

## Scoring and achievement levels

Total: **30 points** across 23 items. Two-point items award partial credit:

| Item type | 2 points | 1 point | 0 points |
|---|---|---|---|
| Multi-select (choose 2) | both correct | exactly one correct | neither correct |
| 4-row matching (1:1) | all 4 rows | 2–3 rows | 0–1 rows |
| Drag & drop (3 slots) | all 3 slots | 2 slots | 0–1 slots |
| Two-part drop-down | both parts | one part | neither part |

### ⚠️ These use Milestones names but are not Milestones levels

Results are labeled **Distinguished / Proficient / Developing / Beginning Learner** at 80 / 65 / 50%.

The *names* match Georgia Milestones so they are familiar to students and families, but the **cut points are teacher-set practice thresholds**, not GaDOE cut scores. Official Milestones achievement levels come from scale scores set through standard setting, which the ALD document notes occurs *after* the first administration of a new assessment — so no percentage-to-level mapping exists to use. A disclaimer stating this appears on the student results screen and at the top of the Class Dashboard.

Adjust the thresholds in `CONFIG.BANDS` in `index.html` — and mirror any change in `BANDS` in `apps-script/Code.gs`, or the workbook will disagree with what students saw.

### Students do not see the answer key

By default the results screen shows each student their score, trait breakdown, and **which questions they earned points on — but not the correct answers or the explanations.** This keeps the form reusable across periods and years. Teachers get the full key and every student response in the workbook.

To change it, edit `CONFIG.REVIEW` in `index.html`:

```js
REVIEW: {
  showCorrectAnswer: false,   // true = show the key for missed items
  showRationale:     false,   // true = show the "Why" explanation
  showOwnResponse:   true     // students see what they chose
}
```

The student CSV download drops the `correctAnswer` column automatically whenever `showCorrectAnswer` is off, so downloading the file cannot be used to obtain the key.

> **Limitation worth knowing.** The item bank lives in `index.html`, so a student who opens View Source can read the answers. This setting stops casual answer-sharing, not a determined student. If that matters for your use, host the form where students cannot easily view source, or ask and I can obfuscate the key.

---

## Teacher setup

### 1. Publish the simulation
Host `index.html` anywhere that serves static files — GitHub Pages, Google Sites, your LMS, or a shared drive. It needs no server and runs entirely in the browser.

### 2. Create the results workbook
1. Create a new Google Sheet (this becomes your gradebook workbook).
2. **Extensions → Apps Script**, delete the starter code, paste all of `apps-script/Code.gs`, and save.
3. **Deploy → New deployment → Web app**
   - *Execute as*: **Me**
   - *Who has access*: **Anyone**
   - Deploy, authorize, and copy the `/exec` URL.
4. Reload the spreadsheet. A **📊 Peer Revision Reports** menu appears. Click **Set up workbook**.

### 3. Connect the two — the step most likely to bite you
`CONFIG.SCRIPT_URL` ships **empty**. Paste your own `/exec` URL between the quotes, near the top of the `<script>` block in `index.html`:

```js
SCRIPT_URL: "https://script.google.com/macros/s/AKfy…/exec",
```

Get it from **Deploy → Manage deployments** in the Apps Script editor — not from the browser address bar of the editor itself.

Until you do, the opening screen shows a red **"Not connected to a results sheet"** banner. That is deliberate: a URL belonging to *someone else's* Apps Script project fails silently — students see "submitted" while nothing reaches your sheet — so an empty value that fails loudly is the safer default.

**If scores are not arriving,** open **Executions** (⏱ in the Apps Script left sidebar) and look for `doPost` entries. No `doPost` at all means submissions are reaching a different script: your `SCRIPT_URL` is wrong. `doPost — Failed` means the deployed code is out of date; redeploy a new version.

> **Updating an existing deployment.** If you already have an Apps Script project for this task, replace its code with `apps-script/Code.gs` and use **Deploy → Manage deployments → Edit → New version**. That keeps the same `/exec` URL, so `index.html` needs no change.

> **Security.** The `/exec` URL is public in this repository and is necessarily reachable by any student browser — that is how the submission works. Treat the destination sheet as append-only and glance at the `Submissions` tab occasionally for junk rows. The **Clear all student data** menu item resets both data tabs. Do not store anything sensitive in this workbook beyond names and periods.

### 4. Fallback if submission fails
The results screen always offers **Download my results (CSV)** and **Print / Save as PDF**, so a network failure never costs a student their work. The CSV mirrors the `Item Responses` tab and can be pasted straight in — minus the `correctAnswer` column, which is withheld from students by default (see above).

---

## The teacher workbook

**Set up workbook** creates ten tabs. All analysis refreshes automatically after each submission, and on demand via **Refresh all reports**. Only each student's **most recent attempt** is analyzed, so retakes do not double-count.

| Tab | Use it to |
|---|---|
| **Class Dashboard** | See class mean/median, achievement-level distribution, and trait means with a recommended instructional tier for each. Start here. |
| **Trait Mastery** | Per-student percentages on the three rubric traits, color-banded, with each student's lowest trait identified. |
| **Intervention Groups** | Students auto-grouped by lowest trait, each group paired with the **exact 3-part guided questions and item stems from the GaDOE Peer Revision Guidance** for that trait. This is your mini-lesson plan. |
| **Item Analysis** | Difficulty (*p*), discrimination (*D*, upper vs. lower 27%), the **most common wrong answer** for each item, and quality flags. The most-common-wrong-answer column is the misconception to address first. |
| **Long-Range Planning** | Class mastery by standard, sorted lowest first, with a recommended tier and a concrete instructional move per standard. Re-run the task later and compare to measure growth. |
| **Vocabulary Flags** | Which terms students flagged as unfamiliar, ranked by frequency — a ready-made vocabulary mini-lesson priority list. |
| **Student x Item** | Heat-mapped matrix. Read down a column to find an item the class missed; read across a row for one student's pattern. |
| **Student Report** | Pick a student from a dropdown for a printable one-page profile — built for conferences, MTSS documentation, and family communication. |
| **Submissions** / **Item Responses** | Raw data. One row per attempt, and one row per student per item (the long format that powers everything else). |

### Reading the reports honestly

- **Trait scores rest on 9–11 points each.** That is enough to point a small group in a direction; it is not enough to diagnose an individual. Confirm a pattern in **Item Analysis** and **Student x Item** before making a placement decision.
- **Discrimination needs numbers.** With fewer than ~20 students, *D* is flagged `N too small for stable stats`. Treat it as a hint.
- **A negative *D* usually means the item, not the students.** Check the key and the wording before reteaching.
- **Vocabulary flags are self-reported** — they show where students *know* they are stuck. Pair them with Item Analysis to catch the words they did not realize they were missing.

---

## Accessibility and device support

- **Touch and keyboard first.** Drag-and-drop items work by tap-to-pick / tap-to-place, so they function on iPads and Chromebooks in tablet mode. Mouse dragging still works as an enhancement. Every interactive control is a real focusable button.
- **Screen readers.** Options carry `role="radio"` / `role="checkbox"` with live `aria-checked`; progress is an `aria-live` region; matching radios have descriptive labels.
- **Answers are never locked.** Students can change or clear any answer before submitting.
- **Autosave.** Progress saves to `localStorage` after every interaction. A refresh or accidental tab close offers a *Resume where I left off* banner.
- **Responsive.** Two-column on desktop, stacked on phones, with no horizontal page scroll.
- Fonts load from Google Fonts; on a network that blocks it the page falls back to system sans-serif and remains fully usable.

---

## Customizing

The item bank is a single `QUESTIONS` array in `index.html`. Each item declares its `type`, `points`, `dok`, `trait`, `bigIdea`, `std`, stem, options, and a `why` rationale shown in the answer review.

If you edit the bank, keep these invariants — they are what the reports depend on:
- Every `sr` item has exactly one `correct: true` and is worth 1 point.
- Every `ms`, `match`, and `dd` item is worth 2 points; a `multipart` item is worth one point per part.
- A `match` item must have the same number of rows and columns, with each row keyed to a **different** column (the 1:1 constraint is enforced in the UI).
- A `dd` item must have at least as many bank items as slots.
- Bump `CONFIG.STORAGE_KEY` and `CONFIG.FORM_VERSION` after changing items, so students do not resume into a stale saved state.

See `docs/ALIGNMENT.md` for the full item map and the DOK/blueprint recalculation if you change the mix.
