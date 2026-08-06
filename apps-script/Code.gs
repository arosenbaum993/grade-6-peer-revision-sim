/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  GRADE 6 ELA — PEER REVISION TASK PRACTICE
 *  Results collection + teacher reporting workbook
 * ───────────────────────────────────────────────────────────────────────────
 *  Deploy:  Extensions → Apps Script → paste this file → Deploy → New deployment
 *           → Web app → Execute as: Me → Who has access: Anyone → Deploy.
 *           Paste the /exec URL into CONFIG.SCRIPT_URL in index.html.
 *
 *  First run:  reload the spreadsheet, then use the
 *              "📊 Peer Revision Reports" menu → "Set up workbook".
 *
 *  Student results are written the instant they submit. The eight report tabs
 *  are rebuilt on demand via "Refresh all reports" — a full rebuild is slow
 *  enough that running it inside every submission made students wait and left
 *  tabs half-built. The Class Dashboard shows when new results are pending.
 * ═══════════════════════════════════════════════════════════════════════════
 */

/* ═══════════════════════════════════════════════════════════════════════
   SPREADSHEET_ID
   ───────────────────────────────────────────────────────────────────────
   Leave this EMPTY if this script lives inside the spreadsheet itself
   (you opened it with Extensions → Apps Script from within the Sheet).
   That is the recommended setup.

   If this is a STANDALONE script project, paste the destination
   spreadsheet's ID here. It is the long string in the sheet's URL:
   docs.google.com/spreadsheets/d/  ⟨THIS PART⟩  /edit
   ═══════════════════════════════════════════════════════════════════════ */
var SPREADSHEET_ID = '';

/**
 * Returns the workbook to write to, whether this script is bound to a
 * spreadsheet or standalone. Throws a readable error instead of failing
 * with "Cannot call method of null" if neither is available.
 */
function getSS_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'No spreadsheet found. This looks like a standalone Apps Script project. ' +
      'Open the destination Google Sheet, copy the ID from its URL ' +
      '(docs.google.com/spreadsheets/d/THIS_PART/edit), and paste it into ' +
      'the SPREADSHEET_ID variable at the top of this file.');
  }
  return ss;
}

/* ── Sheet names ───────────────────────────────────────────────────────── */
var SH = {
  SUB:    'Submissions',
  ITEMS:  'Item Responses',
  DASH:   'Class Dashboard',
  ITEMAN: 'Item Analysis',
  MATRIX: 'Student x Item',
  TRAIT:  'Trait Mastery',
  GROUPS: 'Intervention Groups',
  VOCAB:  'Vocabulary Flags',
  PLAN:   'Long-Range Planning',
  STU:    'Student Report'
};

/* ── Achievement labels (MUST mirror CONFIG.BANDS in index.html) ────────
   These use the Georgia Milestones level names, but the cut points are
   teacher-set practice thresholds — not GaDOE cut scores. ───────────── */
var BANDS = [
  {min: 80, label: 'Distinguished Learner'},
  {min: 65, label: 'Proficient Learner'},
  {min: 50, label: 'Developing Learner'},
  {min: 0,  label: 'Beginning Learner'}
];

/* ── Class-mastery tiers used by Item Analysis and Long-Range Planning ─── */
function tierFor(pct) {
  if (pct >= 80) return {tier: 'Maintain',            move: 'Spiral review only. Use as a warm-up or bell-ringer every few weeks.'};
  if (pct >= 65) return {tier: 'Targeted small group', move: 'Pull the students below 65% for a 10-minute small-group mini-lesson.'};
  if (pct >= 50) return {tier: 'Whole-class reteach',  move: 'Reteach to the whole class with modeling, then guided practice.'};
  return             {tier: 'Reteach + prerequisites', move: 'Reteach from the beginning and check prerequisite skills before reassessing.'};
}

/* ═══════════════════════════════════════════════════════════════════════
   GaDOE PEER REVISION GUIDANCE — trait → guided questions & item stems
   Source: "Grade 6: Classroom Peer Revision Guidance", GaDOE, February 2026.
   These drive the Intervention Groups sheet so the data points straight at
   the instructional moves the state document recommends.
   ═══════════════════════════════════════════════════════════════════════ */
var TRAIT_GUIDANCE = {
  't1': {
    name: 'Purpose & Organization',
    focus: 'Introduction and claim · organizational structure · style for purpose and audience · transitions and cohesion · conclusion.',
    guided: [
      'Does the introduction align with the research topic? Is the purpose clearly defined? Should the last sentence of the introduction be revised to state the claim clearly?',
      'What is the organizational pattern of the draft? Does it organize information effectively for the audience? Should it be revised?',
      'What is the purpose of the draft, and who is the target audience? What style fits them? Are there sentences that make the style too casual?',
      'What transition words and phrases are used? Do they connect ideas within and between paragraphs? Where should transitions be added to create cohesion?',
      'What is the purpose of the conclusion? Does it wrap up the main ideas? Should it add a call to act, think, or feel?'
    ],
    stems: [
      'Which sentence should be added to the END of the paragraph to strengthen the introduction?',
      'Which sentence, in the introduction, should be revised to effectively present the claim?',
      'Select TWO sentences that should be revised to maintain a style appropriate for the purpose and audience.',
      'Review transitions in and between paragraphs. Which transitions should be revised to BEST connect reasons and evidence?',
      'Which sentence should be added to provide the BEST conclusion paragraph?',
      'Which TWO sentences should be removed to make all the information relevant and connected to the stated claim?'
    ]
  },
  't2': {
    name: 'Evidence & Elaboration',
    focus: 'Reasons that support the claim · relevant evidence from multiple sources · elaboration techniques · counterclaim · crediting sources.',
    guided: [
      'Does the response provide reasons? Do the reasons support the claim? Which reasons need revision to better support it?',
      'Is the evidence relevant and well chosen for each reason? Are there reasons that need more evidence? Does the evidence come from multiple sources?',
      'Does the response acknowledge a counterclaim? What opposing viewpoint might someone have? How could the counterclaim be revised to elaborate effectively?',
      'What key details or explanations are already included? Does the paragraph need another example, explanation, or clarification to support the reason better?',
      'Does the response credit the sources when using facts, examples, statistics, or explanations? Which pieces of evidence should be revised to credit the source?'
    ],
    stems: [
      'Which TWO pieces of credible information could be added to support the reasons and help develop the claim?',
      'Add evidence from TWO sources that BEST support the reason and effectively elaborate on the ideas presented.',
      'Which sentence should be added to address a counterclaim and strengthen the argument?',
      'Which information should be added to the end of the paragraph to BEST elaborate on the ideas?',
      'Which evidence from Source #1 and Source #2 makes the BEST connection to the claim? Choose ONE piece of evidence from EACH source.'
    ]
  },
  't3': {
    name: 'Language Usage & Conventions',
    focus: 'Complete sentences and sentence variety · verb tense · subject-verb agreement · active voice · punctuation and usage.',
    guided: [
      'Which ideas are meant to be joined together? Which relationship connects them (addition, contrast, cause/effect)? Are any conjunctions missing, incorrect, or misused?',
      'Are there sentences that are too short, choppy, or repetitive? Which simple sentences with related ideas could be joined?',
      'What verb tense is used throughout the paragraph? Does the verb match the subject in number? Do any sentences shift tense without a clear reason?',
      'Does each sentence clearly show the subject doing the action? Which sentences would be clearer if revised to active voice?',
      'What punctuation marks are used in each sentence? Are commas, apostrophes, quotation marks, and end punctuation used correctly?'
    ],
    stems: [
      'Which TWO changes should you make to ensure conjunctions correctly join ideas in the draft?',
      'Revise or join sentences using correct conjunctions to maintain consistent sentence structure.',
      'Revise verb phrases to ensure correct subject and verb agreement.',
      'Which sentences should be revised to correct errors in verb tense?',
      'Which sentences should be revised to maintain active voice?',
      'Select and correct punctuation errors within sentences of the draft.'
    ]
  }
};

/* Standard-level reteach moves for the Long-Range Planning sheet. */
var STANDARD_MOVES = {
  '6.T.SS.1.a': 'Explain how authors modify organizational structures to convey meaning. Model with two texts that organize the same information differently.',
  '6.T.SS.1.c': 'Concept repetition and connected terms for transitions and cohesion. Use a transition sort, then revise a paragraph that has none.',
  '6.T.SS.1.d': 'Craft multi-paragraph texts with a coherent structure. Focus a mini-lesson on introductions and conclusions using the rubric bullets.',
  '6.T.SS.2.c': 'Compare formal and informal style. Have students highlight slang and contractions in a draft, then rewrite for a principal audience.',
  '6.T.T.2.d':  'Apply expository techniques to elaborate. Practice the "evidence sandwich": state it, explain it, connect it back to the reason.',
  '6.T.T.3.a':  'Recognize claim, evidence, counterclaim, and conclusion. Label the parts of two model arguments before revising any draft.',
  '6.T.T.3.c':  'Apply argumentative techniques. Focus on writing a claim that takes a position, and a counterclaim that is answered rather than dismissed.',
  '6.T.RA.2.a': 'Locate evidence in multiple sources to support a central idea. Sort source facts by which reason they actually support.',
  '6.T.RA.2.b': 'Use basic parenthetical citations when quoting or paraphrasing. Practice turning a vague "they say" into a named source.',
  '6.L.GC.1':   'Grammar, usage, and mechanics from the GUM chart: vague pronoun reference and commas in a series. Use daily sentence-editing routines.',
  '6.L.GC.2.b': 'Combine ideas with a variety of sentence structures, maintaining consistent verb tense. Practice sentence-combining with conjunctions.',
  '6.L.GC.2.c': 'Identify and use active voice, revising for subject-verb agreement. Drill subjects separated from verbs by prepositional phrases.',
  '6.L.V.1.b':  'Use precise academic vocabulary when constructing texts. Build a "vague word → precise word" anchor chart from student drafts.'
};

/* ═══════════════════════════════════════════════════════════════════════
   WEB APP ENDPOINT
   ═══════════════════════════════════════════════════════════════════════ */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var raw = (e && e.parameter && e.parameter.payload)
      ? e.parameter.payload
      : (e && e.postData ? e.postData.contents : null);
    if (!raw) return jsonOut({ok: false, error: 'no payload'});

    var p = JSON.parse(raw);
    if (!p || !p.submissionId || !p.items) return jsonOut({ok: false, error: 'malformed payload'});

    var ss = getSS_();
    ensureDataSheets_(ss);

    // Idempotency — a retry from the student's browser must not double-log.
    var sub = ss.getSheetByName(SH.SUB);
    var last = sub.getLastRow();
    if (last > 1) {
      var existing = sub.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (existing[i][0] === p.submissionId) {
          return jsonOut({ok: true, duplicate: true, message: 'already recorded'});
        }
      }
    }

    writeSubmission_(sub, p);
    writeItems_(ss.getSheetByName(SH.ITEMS), p);

    // Deliberately NOT rebuilding reports here. A full rebuild is ~350 Sheets
    // operations (10-20s). Running it inside every submit meant a class
    // submitting together queued on the 30s script lock, later submissions were
    // killed part-way through the rebuild, and tabs were left half-updated.
    // Data is written immediately; reports refresh from the menu.
    markDirty_(p.submissionId);

    return jsonOut({ok: true, submissionId: p.submissionId, points: p.pointsEarned});
  } catch (err) {
    return jsonOut({ok: false, error: String(err)});
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Record that new results have landed since the last report refresh. */
function markDirty_(id) {
  try {
    var props = PropertiesService.getScriptProperties();
    var n = Number(props.getProperty('pendingCount') || 0) + 1;
    props.setProperties({pendingCount: String(n), lastSubmissionAt: new Date().toISOString()});
  } catch (e) { /* tracking is a convenience, never block a submit for it */ }
}
function clearDirty_() {
  try {
    PropertiesService.getScriptProperties().setProperties(
      {pendingCount: '0', lastRefreshAt: new Date().toISOString()});
  } catch (e) {}
}
function pendingCount_() {
  try { return Number(PropertiesService.getScriptProperties().getProperty('pendingCount') || 0); }
  catch (e) { return 0; }
}

function doGet() {
  return jsonOut({ok: true, service: 'Grade 6 Peer Revision Task results endpoint'});
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ═══════════════════════════════════════════════════════════════════════
   DATA WRITERS
   ═══════════════════════════════════════════════════════════════════════ */
var SUB_HEAD = ['submissionId','submittedAt','studentName','studentId','period','formId','formVersion',
  'elapsedMinutes','pointsEarned','pointsPossible','percent','band',
  'T1 Purpose & Organization %','T2 Evidence & Elaboration %','T3 Language Usage & Conventions %',
  'Structure & Style %','Techniques %','Research & Analysis %','Grammar Conventions %','Vocabulary %',
  'DOK1 %','DOK2 %','DOK3 %','vocabFlagCount','vocabFlagTerms'];

var ITEM_HEAD = ['submissionId','submittedAt','studentName','studentId','period',
  'itemNumber','itemId','itemType','trait','standardGroup','standard','dok',
  'pointsEarned','pointsPossible','outcome','studentResponse','correctAnswer',
  'vocabFlagged','secondsToAnswer'];

function writeSubmission_(sh, p) {
  var t = p.traits || {}, b = p.bigIdeas || {}, d = p.dok || {};
  var pc = function (o, k) { return (o && o[k] && o[k].percent !== null && o[k].percent !== undefined) ? o[k].percent / 100 : ''; };
  var terms = (p.vocabFlags || []).map(function (v) { return String(v.term || '').split('—')[0].trim(); });
  sh.appendRow([
    p.submissionId, p.submittedAt ? new Date(p.submittedAt) : new Date(),
    p.studentName, p.studentId || '', p.period || '', p.formId || '', p.formVersion || '',
    p.elapsedSeconds ? Math.round(p.elapsedSeconds / 60 * 10) / 10 : '',
    p.pointsEarned, p.pointsPossible, (p.percent || 0) / 100, p.bandLabel || '',
    pc(t,'t1'), pc(t,'t2'), pc(t,'t3'),
    pc(b,'ss'), pc(b,'tt'), pc(b,'ra'), pc(b,'gc'), pc(b,'vo'),
    pc(d,'1'), pc(d,'2'), pc(d,'3'),
    terms.length, terms.join('; ')
  ]);
}

function writeItems_(sh, p) {
  var when = p.submittedAt ? new Date(p.submittedAt) : new Date();
  var rows = (p.items || []).map(function (it) {
    return [p.submissionId, when, p.studentName, p.studentId || '', p.period || '',
      it.n, it.id, it.type, it.traitLabel, it.bigIdeaLabel, it.standard, it.dok,
      it.pointsEarned, it.pointsPossible, it.outcome, it.response, it.key,
      it.vocabFlagged ? 'YES' : '', (it.secondsToAnswer === null || it.secondsToAnswer === undefined) ? '' : it.secondsToAnswer];
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, ITEM_HEAD.length).setValues(rows);
}

function ensureDataSheets_(ss) {
  var sub = ss.getSheetByName(SH.SUB);
  if (!sub) { sub = ss.insertSheet(SH.SUB); sub.appendRow(SUB_HEAD); styleHeader_(sub, SUB_HEAD.length); }
  var it = ss.getSheetByName(SH.ITEMS);
  if (!it) { it = ss.insertSheet(SH.ITEMS); it.appendRow(ITEM_HEAD); styleHeader_(it, ITEM_HEAD.length); }
}

/* ═══════════════════════════════════════════════════════════════════════
   MENU + SETUP
   ═══════════════════════════════════════════════════════════════════════ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Peer Revision Reports')
    .addItem('Set up workbook', 'setupWorkbook')
    .addItem('Refresh all reports (do this after students submit)', 'refreshReportsFromMenu')
    .addSeparator()
    .addItem('Clear all student data', 'clearAllData')
    .addToUi();
}

function setupWorkbook() {
  var ss = getSS_();
  ensureDataSheets_(ss);
  [SH.DASH, SH.ITEMAN, SH.MATRIX, SH.TRAIT, SH.GROUPS, SH.VOCAB, SH.PLAN, SH.STU].forEach(function (n) {
    if (!ss.getSheetByName(n)) ss.insertSheet(n);
  });
  // Put the teacher-facing tabs first.
  [SH.DASH, SH.TRAIT, SH.GROUPS, SH.ITEMAN, SH.PLAN, SH.VOCAB, SH.MATRIX, SH.STU, SH.SUB, SH.ITEMS]
    .forEach(function (n, i) {
      var s = ss.getSheetByName(n);
      if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); }
    });
  rebuildReports();
  ss.setActiveSheet(ss.getSheetByName(SH.DASH));
  SpreadsheetApp.getUi().alert('Workbook ready.\n\n' +
    'Student scores land in the Submissions and Item Responses tabs the moment they submit.\n\n' +
    'The eight report tabs are rebuilt when you choose "Refresh all reports" — not on every ' +
    'submission, because a full rebuild takes 10-20 seconds and would make students wait. ' +
    'The Class Dashboard tells you when new results are waiting to be included.');
}

function clearAllData() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert('Clear all student data?',
    'This permanently deletes every submission and item response in this workbook. ' +
    'Reports will be rebuilt empty. This cannot be undone.', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  var ss = getSS_();
  [SH.SUB, SH.ITEMS].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
  });
  rebuildReports();
  ui.alert('All student data cleared.');
}

/* ═══════════════════════════════════════════════════════════════════════
   REPORT BUILDER
   ═══════════════════════════════════════════════════════════════════════ */
function rebuildReports() {
  var ss = getSS_();
  ensureDataSheets_(ss);
  var data = readData_(ss);
  var steps = [
    [SH.DASH,   buildDashboard_],
    [SH.TRAIT,  buildTraitMastery_],
    [SH.GROUPS, buildInterventionGroups_],
    [SH.ITEMAN, buildItemAnalysis_],
    [SH.PLAN,   buildPlanning_],
    [SH.VOCAB,  buildVocab_],
    [SH.MATRIX, buildMatrix_],
    [SH.STU,    buildStudentReport_]
  ];
  var failures = [];
  steps.forEach(function (step) {
    // Each tab is isolated. Previously these ran in sequence with no guard, so
    // a single failure left every later tab showing stale data with no warning.
    try {
      step[1](ss, data);
      flushRules_(ss.getSheetByName(step[0]));
    } catch (err) {
      failures.push(step[0] + ' — ' + ((err && err.message) ? err.message : String(err)));
    }
  });
  if (failures.length) {
    console.error('rebuildReports: ' + failures.join(' | '));
    noteFailures_(ss, failures);
  } else {
    clearDirty_();
  }
  return failures;
}

/** Make a failed rebuild visible instead of silent. */
function noteFailures_(ss, failures) {
  var sh = ss.getSheetByName(SH.DASH);
  if (!sh) return;
  try {
    var r = sh.getLastRow() + 2;
    sh.getRange(r, 1).setValue('\u26a0 Some reports could not be rebuilt. Use "Refresh all reports"; if it persists, send these lines to support:')
      .setFontWeight('bold').setFontColor('#b3261e');
    failures.forEach(function (f, i) { sh.getRange(r + 1 + i, 1).setValue(f).setFontColor('#b3261e'); });
  } catch (e) { /* the dashboard itself may be the thing that failed */ }
}

/** Menu entry point: refresh, and say so plainly either way. */
function refreshReportsFromMenu() {
  var failures = rebuildReports();
  var ui = SpreadsheetApp.getUi();
  if (failures && failures.length) {
    ui.alert('Some reports failed to rebuild:\n\n' + failures.join('\n\n') +
             '\n\nThe details are also written at the bottom of the Class Dashboard.');
  } else {
    ui.alert('All reports refreshed.');
  }
}

function readData_(ss) {
  var subSh = ss.getSheetByName(SH.SUB), itSh = ss.getSheetByName(SH.ITEMS);
  var subs = [], items = [];
  if (subSh.getLastRow() > 1) {
    subSh.getRange(2, 1, subSh.getLastRow() - 1, SUB_HEAD.length).getValues().forEach(function (r) {
      // The level is recomputed from the percentage rather than read from the
      // stored column. That way rows logged under older labels or older cut
      // scores are re-labeled under the current BANDS instead of silently
      // dropping out of the distribution counts.
      subs.push({
        id: r[0], when: r[1], name: r[2], sid: r[3], period: r[4],
        minutes: r[7], earned: Number(r[8]) || 0, possible: Number(r[9]) || 0, pct: Number(r[10]) || 0,
        band: bandFor_(Number(r[10]) || 0), bandAsLogged: r[11],
        t1: r[12], t2: r[13], t3: r[14], flags: Number(r[23]) || 0, terms: r[24]
      });
    });
  }
  if (itSh.getLastRow() > 1) {
    itSh.getRange(2, 1, itSh.getLastRow() - 1, ITEM_HEAD.length).getValues().forEach(function (r) {
      items.push({
        sub: r[0], name: r[2], period: r[4], n: Number(r[5]), id: r[6], type: r[7],
        trait: r[8], group: r[9], std: r[10], dok: Number(r[11]),
        earned: Number(r[12]) || 0, possible: Number(r[13]) || 0, outcome: r[14],
        response: r[15], key: r[16], flagged: r[17] === 'YES', secs: r[18]
      });
    });
  }
  // Keep only each student's most recent attempt in the analysis.
  var latest = {};
  subs.forEach(function (s) {
    var k = (s.name || '') + '|' + (s.period || '');
    if (!latest[k] || new Date(s.when) > new Date(latest[k].when)) latest[k] = s;
  });
  var keepIds = {};
  Object.keys(latest).forEach(function (k) { keepIds[latest[k].id] = true; });
  return {
    subs: Object.keys(latest).map(function (k) { return latest[k]; })
              .sort(function (a, b) { return String(a.name).localeCompare(String(b.name)); }),
    items: items.filter(function (i) { return keepIds[i.sub]; }),
    allSubs: subs
  };
}

/* ── helpers ─────────────────────────────────────────────────────────── */
function styleHeader_(sh, cols) {
  sh.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff')
    .setVerticalAlignment('middle').setWrap(true);
  sh.setFrozenRows(1);
}
function resetSheet_(ss, name) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  // clear() does NOT remove merged regions. On a rebuild the row positions move
  // with the data, so a later merge() that only partially overlaps a leftover
  // merge throws — which used to abort the whole rebuild and leave every
  // subsequent tab stale. Break merges explicitly first. No-op if none exist.
  all.breakApart();
  sh.clear();
  sh.clearConditionalFormatRules();
  all.clearNote();
  return sh;
}
/**
 * Title and subtitle in rows 1-2.
 * Pass willFreezeColumns=true on sheets that call setFrozenColumns(): Sheets
 * refuses to freeze a column boundary that cuts through a merged range
 * ("you can't freeze columns which contain only part of a merged cell"), which
 * previously failed the whole report. On those sheets the text is left
 * unmerged and simply overflows across the empty cells to its right.
 */
function titleRow_(sh, text, sub, cols, willFreezeColumns) {
  sh.getRange(1, 1).setValue(text).setFontSize(14).setFontWeight('bold').setFontColor('#1a3a5c');
  sh.getRange(2, 1).setValue(sub).setFontSize(9).setFontColor('#666666')
    .setWrap(!willFreezeColumns);
  if (!willFreezeColumns) {
    sh.getRange(1, 1, 1, cols).merge();
    sh.getRange(2, 1, 1, cols).merge();
  }
}
function writeTable_(sh, startRow, head, rows) {
  sh.getRange(startRow, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff').setWrap(true);
  if (rows.length) sh.getRange(startRow + 1, 1, rows.length, head.length).setValues(rows);
  return startRow + rows.length + 1;
}
function pctBands_(sh, range) {
  // Rules are queued on the sheet object and written once, at the end of the
  // report. Each get/set pair is a network round trip; doing it per call was a
  // large share of the rebuild cost.
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.8)
      .setBackground('#d4efdc').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.65, 0.7999)
      .setBackground('#dbe9f7').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.5, 0.6499)
      .setBackground('#fdf1cf').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.5)
      .setBackground('#fadbd8').setRanges([range]).build()
  ];
  sh.__pendingRules = (sh.__pendingRules || []).concat(rules);
}

/** Write all queued conditional-format rules for a sheet in one call. */
function flushRules_(sh) {
  if (!sh || !sh.__pendingRules || !sh.__pendingRules.length) return;
  sh.setConditionalFormatRules(sh.__pendingRules);
  sh.__pendingRules = [];
}
function mean_(arr) { return arr.length ? arr.reduce(function (a, b) { return a + b; }, 0) / arr.length : 0; }
function bandFor_(pct) {
  for (var i = 0; i < BANDS.length; i++) if (pct * 100 >= BANDS[i].min) return BANDS[i].label;
  return BANDS[BANDS.length - 1].label;
}
function emptyNotice_(sh, cols) {
  sh.getRange(4, 1).setValue('No student submissions yet. This report fills in automatically as students submit.')
    .setFontColor('#999999').setFontStyle('italic');
  sh.getRange(4, 1, 1, cols).merge();
}

/* ── 1. CLASS DASHBOARD ──────────────────────────────────────────────── */
function buildDashboard_(ss, d) {
  var sh = resetSheet_(ss, SH.DASH);
  titleRow_(sh, 'Class Dashboard — Grade 6 Peer Revision Task',
    'PRACTICE DATA ONLY. The levels below use Georgia Milestones NAMES but are based on teacher-set practice cut scores. ' +
    'They are not official Georgia Milestones achievement levels and do not predict one. ' +
    'Milestones levels are based on scale scores set through GaDOE standard setting. Use these results to plan instruction, not to predict a Milestones score.', 8);
  if (!d.subs.length) { emptyNotice_(sh, 8); sh.setColumnWidth(1, 300); return; }

  var r = 4;
  var pending = pendingCount_();
  if (pending > 0) {
    sh.getRange(r, 1).setValue('\u21bb ' + pending + ' new submission' + (pending === 1 ? '' : 's') +
      ' since these reports were built. Use \u201cPeer Revision Reports \u2192 Refresh all reports\u201d to include them.')
      .setFontWeight('bold').setFontColor('#7a5a00').setBackground('#fff8e6').setWrap(true);
    sh.getRange(r, 1, 1, 5).merge();
    r += 2;
  }
  var pcts = d.subs.map(function (s) { return s.pct; });
  var sorted = pcts.slice().sort(function (a, b) { return a - b; });
  var median = sorted.length % 2 ? sorted[(sorted.length - 1) / 2]
    : (sorted[sorted.length / 2 - 1] + sorted[sorted.length / 2]) / 2;

  var overviewHeader = r;
  r = writeTable_(sh, r, ['Overview', 'Value'], [
    ['Students assessed', d.subs.length],
    ['Class mean', mean_(pcts)],
    ['Class median', median],
    ['Lowest / Highest', Math.round(Math.min.apply(null, pcts) * 100) + '% / ' + Math.round(Math.max.apply(null, pcts) * 100) + '%'],
    ['Median time on task (min)', (function () {
      var m = d.subs.map(function (s) { return Number(s.minutes) || 0; }).filter(function (x) { return x > 0; }).sort(function (a, b) { return a - b; });
      return m.length ? m[Math.floor(m.length / 2)] : '—';
    })()],
    ['Total attempts logged (incl. retakes)', d.allSubs.length]
  ]);
  // Relative to the table, not absolute: the stale-results banner above shifts
  // these rows, and a hard-coded range formatted "Students assessed" as 300%.
  sh.getRange(overviewHeader + 2, 2, 2, 1).setNumberFormat('0%');  // class mean, class median
  sh.getRange(overviewHeader + 1, 2, 1, 1).setNumberFormat('0');   // students assessed = a count
  sh.getRange(overviewHeader + 6, 2, 1, 1).setNumberFormat('0');   // total attempts = a count
  r += 1;

  // Band distribution
  var counts = {};
  BANDS.forEach(function (b) { counts[b.label] = 0; });
  d.subs.forEach(function (s) { counts[s.band] = (counts[s.band] || 0) + 1; });
  r = writeTable_(sh, r, ['Practice Level', 'Students', '% of Class', 'What it means'], BANDS.map(function (b) {
    return [b.label, counts[b.label] || 0, (counts[b.label] || 0) / d.subs.length,
      b.min >= 80 ? 'Ready for independent peer revision.' :
      b.min >= 65 ? 'Close. Needs targeted practice on one or two traits.' :
      b.min >= 50 ? 'Needs small-group reteaching on the weakest trait.' :
                    'Needs substantial support. Start with modeled revision.'];
  }));
  sh.getRange(r - BANDS.length, 3, BANDS.length, 1).setNumberFormat('0%');
  r += 1;

  // Trait means — the headline instructional signal
  var traitRows = ['t1', 't2', 't3'].map(function (k) {
    var vals = d.subs.map(function (s) { return Number(s[k]) || 0; });
    var m = mean_(vals);
    var t = tierFor(m * 100);
    return [TRAIT_GUIDANCE[k].name, m, vals.filter(function (v) { return v < 0.65; }).length, t.tier, t.move];
  });
  r = writeTable_(sh, r, ['Rubric Trait', 'Class Mean', 'Students below 65%', 'Recommended Tier', 'Instructional Move'], traitRows);
  var tr = sh.getRange(r - 3, 2, 3, 1);
  tr.setNumberFormat('0%'); pctBands_(sh, tr);
  r += 1;

  // Standard-group and DOK means, computed from item rows
  var byGroup = {}, byDok = {};
  d.items.forEach(function (i) {
    (byGroup[i.group] = byGroup[i.group] || []).push(i.earned / i.possible);
    (byDok['DOK ' + i.dok] = byDok['DOK ' + i.dok] || []).push(i.earned / i.possible);
  });
  r = writeTable_(sh, r, ['Standard Group (blueprint reporting category)', 'Class Mean', 'Items'],
    Object.keys(byGroup).sort().map(function (g) { return [g, mean_(byGroup[g]), byGroup[g].length / d.subs.length]; }));
  var gr = sh.getRange(r - Object.keys(byGroup).length, 2, Object.keys(byGroup).length, 1);
  gr.setNumberFormat('0%'); pctBands_(sh, gr);
  r += 1;

  r = writeTable_(sh, r, ['Depth of Knowledge', 'Class Mean', 'Blueprint share of this form'],
    Object.keys(byDok).sort().map(function (k) {
      var share = {'DOK 1': '2 of 30 pts (6.7%)', 'DOK 2': '16 of 30 pts (53.3%)', 'DOK 3': '12 of 30 pts (40.0%)'}[k] || '';
      return [k, mean_(byDok[k]), share];
    }));
  var dr = sh.getRange(r - Object.keys(byDok).length, 2, Object.keys(byDok).length, 1);
  dr.setNumberFormat('0%'); pctBands_(sh, dr);

  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 100); sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 150); sh.setColumnWidth(5, 420);
  sh.getRange(1, 1, sh.getLastRow(), 5).setVerticalAlignment('top');
}

/* ── 2. TRAIT MASTERY BY STUDENT ─────────────────────────────────────── */
function buildTraitMastery_(ss, d) {
  var sh = resetSheet_(ss, SH.TRAIT);
  titleRow_(sh, 'Trait Mastery by Student',
    'Each student\'s percentage on the three traits of the Georgia Milestones writing rubric. ' +
    'Trait scores come from 9–11 points each, so treat a single student\'s trait score as a starting point for a conversation, not a diagnosis. ' +
    'Look at the Item Analysis and Student x Item sheets to confirm a pattern before grouping.', 8, true);
  if (!d.subs.length) { emptyNotice_(sh, 8); return; }

  var rows = d.subs.map(function (s) {
    var t = {t1: Number(s.t1) || 0, t2: Number(s.t2) || 0, t3: Number(s.t3) || 0};
    var lowest = ['t1', 't2', 't3'].reduce(function (a, b) { return t[a] <= t[b] ? a : b; });
    return [s.name, s.sid, s.period, s.earned, s.pct, s.band,
      t.t1, t.t2, t.t3, TRAIT_GUIDANCE[lowest].name, s.flags];
  });
  var head = ['Student', 'ID', 'Period', 'Points /30', 'Overall %', 'Practice Level',
    'T1 Purpose & Organization', 'T2 Evidence & Elaboration', 'T3 Language & Conventions',
    'Lowest Trait (focus here)', 'Vocab Flags'];
  writeTable_(sh, 4, head, rows);

  sh.getRange(5, 5, rows.length, 1).setNumberFormat('0%');
  var traitRange = sh.getRange(5, 7, rows.length, 3);
  traitRange.setNumberFormat('0%');
  pctBands_(sh, traitRange);
  pctBands_(sh, sh.getRange(5, 5, rows.length, 1));
  sh.setFrozenRows(4); sh.setFrozenColumns(1);
  sh.setColumnWidth(1, 170); sh.setColumnWidth(10, 210);
  for (var c = 7; c <= 9; c++) sh.setColumnWidth(c, 130);
}

/* ── 3. INTERVENTION GROUPS ──────────────────────────────────────────── */
function buildInterventionGroups_(ss, d) {
  var sh = resetSheet_(ss, SH.GROUPS);
  titleRow_(sh, 'Intervention Groups — grouped by lowest rubric trait',
    'Students are grouped by the trait where they scored lowest. The guided questions and item stems under each group come directly from ' +
    '"Grade 6: Classroom Peer Revision Guidance" (GaDOE, February 2026) — use them to build the 3-part mini-lesson for that group. ' +
    'A student appears in only one group: their greatest area of need.', 6);
  if (!d.subs.length) { emptyNotice_(sh, 6); return; }

  var groups = {t1: [], t2: [], t3: []};
  d.subs.forEach(function (s) {
    var t = {t1: Number(s.t1) || 0, t2: Number(s.t2) || 0, t3: Number(s.t3) || 0};
    var lowest = ['t1', 't2', 't3'].reduce(function (a, b) { return t[a] <= t[b] ? a : b; });
    groups[lowest].push({name: s.name, period: s.period, pct: t[lowest], overall: s.pct});
  });

  var r = 4;
  ['t1', 't2', 't3'].forEach(function (k) {
    var g = TRAIT_GUIDANCE[k];
    var members = groups[k].sort(function (a, b) { return a.pct - b.pct; });

    sh.getRange(r, 1).setValue('GROUP: ' + g.name + '  (' + members.length + ' student' + (members.length === 1 ? '' : 's') + ')')
      .setFontWeight('bold').setFontSize(12).setFontColor('#ffffff')
      .setBackground(k === 't1' ? '#0891b2' : k === 't2' ? '#d97706' : '#059669');
    sh.getRange(r, 1, 1, 6).merge();
    r++;
    sh.getRange(r, 1).setValue('Focus: ' + g.focus).setFontStyle('italic').setFontColor('#444444').setWrap(true);
    sh.getRange(r, 1, 1, 6).merge();
    r += 2;

    if (members.length) {
      r = writeTable_(sh, r, ['Student', 'Period', 'Score on this trait', 'Overall %'],
        members.map(function (m) { return [m.name, m.period, m.pct, m.overall]; }));
      var mr = sh.getRange(r - members.length, 3, members.length, 2);
      mr.setNumberFormat('0%'); pctBands_(sh, mr);
    } else {
      sh.getRange(r, 1).setValue('No students have this trait as their lowest — no group needed.')
        .setFontStyle('italic').setFontColor('#999999');
      r++;
    }
    r += 1;

    sh.getRange(r, 1).setValue('3-part guided questions for this group (GaDOE Peer Revision Guidance):')
      .setFontWeight('bold').setFontColor('#1a3a5c');
    r++;
    g.guided.forEach(function (q) {
      sh.getRange(r, 1).setValue('•  ' + q).setWrap(true);
      sh.getRange(r, 1, 1, 6).merge();
      r++;
    });
    r++;
    sh.getRange(r, 1).setValue('Matching assessment item stems to practice with:')
      .setFontWeight('bold').setFontColor('#1a3a5c');
    r++;
    g.stems.forEach(function (s) {
      sh.getRange(r, 1).setValue('→  ' + s).setWrap(true).setFontColor('#555555');
      sh.getRange(r, 1, 1, 6).merge();
      r++;
    });
    r += 2;
  });
  sh.setColumnWidth(1, 560); sh.setColumnWidth(2, 80);
  sh.setColumnWidth(3, 140); sh.setColumnWidth(4, 100);
}

/* ── 4. ITEM ANALYSIS ────────────────────────────────────────────────── */
function buildItemAnalysis_(ss, d) {
  var sh = resetSheet_(ss, SH.ITEMAN);
  titleRow_(sh, 'Item Analysis',
    'Difficulty (p) is the average share of available points earned. Discrimination (D) compares the top 27% of scorers with the bottom 27%; ' +
    'higher is better, and a negative D usually means a miskeyed or confusing item. "Most common wrong answer" is the misconception to address first. ' +
    'With fewer than about 20 students, treat D as a hint rather than a statistic.', 10);
  if (!d.items.length) { emptyNotice_(sh, 10); return; }

  // rank students by total score for upper/lower groups
  var totals = {};
  d.items.forEach(function (i) { totals[i.sub] = (totals[i.sub] || 0) + i.earned; });
  var ranked = Object.keys(totals).sort(function (a, b) { return totals[b] - totals[a]; });
  var cut = Math.max(1, Math.round(ranked.length * 0.27));
  var upper = {}, lower = {};
  ranked.slice(0, cut).forEach(function (s) { upper[s] = true; });
  ranked.slice(-cut).forEach(function (s) { lower[s] = true; });
  var stable = ranked.length >= 20;

  var byItem = {};
  d.items.forEach(function (i) { (byItem[i.n] = byItem[i.n] || []).push(i); });

  var rows = Object.keys(byItem).map(Number).sort(function (a, b) { return a - b; }).map(function (n) {
    var list = byItem[n], first = list[0];
    var p = mean_(list.map(function (i) { return i.earned / i.possible; }));
    var up = list.filter(function (i) { return upper[i.sub]; }).map(function (i) { return i.earned / i.possible; });
    var lo = list.filter(function (i) { return lower[i.sub]; }).map(function (i) { return i.earned / i.possible; });
    var D = (up.length && lo.length) ? mean_(up) - mean_(lo) : '';

    // most common wrong response
    var wrong = {};
    list.filter(function (i) { return i.outcome !== 'full'; })
        .forEach(function (i) { var k = String(i.response); wrong[k] = (wrong[k] || 0) + 1; });
    var top = Object.keys(wrong).sort(function (a, b) { return wrong[b] - wrong[a]; })[0];
    var topLabel = top ? top.substring(0, 90) + (top.length > 90 ? '…' : '') + '  (' + wrong[top] + ')' : '—';

    var flags = [];
    if (p >= 0.95) flags.push('Very easy');
    if (p <= 0.30) flags.push('Very hard — check the key and whether it was taught');
    if (stable && D !== '' && D < 0) flags.push('NEGATIVE discrimination — review the key');
    else if (stable && D !== '' && D < 0.10) flags.push('Low discrimination');
    if (!stable) flags.push('N too small for stable stats');
    var flagged = list.filter(function (i) { return i.flagged; }).length;
    if (flagged >= Math.max(2, list.length * 0.25)) flags.push('Vocabulary barrier (' + flagged + ' flags)');

    return [n, first.id, first.type, first.trait, first.group, first.std, first.dok,
      p, D === '' ? '—' : D, topLabel, flags.join('; ') || 'OK'];
  });

  writeTable_(sh, 4, ['#', 'Item', 'Type', 'Trait', 'Standard Group', 'Standard', 'DOK',
    'Difficulty (p)', 'Discrimination (D)', 'Most common wrong answer (count)', 'Flags'], rows);
  sh.getRange(5, 8, rows.length, 1).setNumberFormat('0%');
  sh.getRange(5, 9, rows.length, 1).setNumberFormat('0.00');
  pctBands_(sh, sh.getRange(5, 8, rows.length, 1));
  sh.setFrozenRows(4);
  sh.setColumnWidth(1, 40); sh.setColumnWidth(4, 170); sh.setColumnWidth(5, 150);
  sh.setColumnWidth(10, 420); sh.setColumnWidth(11, 260);
  sh.getRange(5, 1, rows.length, 11).setVerticalAlignment('top').setWrap(true);
}

/* ── 5. LONG-RANGE PLANNING ──────────────────────────────────────────── */
function buildPlanning_(ss, d) {
  var sh = resetSheet_(ss, SH.PLAN);
  titleRow_(sh, 'Long-Range Instructional Planning',
    'Class mastery by standard, with a recommended tier and instructional move for each. Work down the list — the lowest-mastery standards ' +
    'are where whole-class time pays off most. Re-run this task later in the year and compare to see whether mastery moved.', 7);
  if (!d.items.length) { emptyNotice_(sh, 7); return; }

  var byStd = {};
  d.items.forEach(function (i) {
    var k = i.std;
    byStd[k] = byStd[k] || {vals: [], trait: i.trait, group: i.group, items: {}};
    byStd[k].vals.push(i.earned / i.possible);
    byStd[k].items[i.n] = true;
  });

  var rows = Object.keys(byStd).map(function (std) {
    var o = byStd[std], m = mean_(o.vals), t = tierFor(m * 100);
    var studentsBelow = {};
    d.items.filter(function (i) { return i.std === std && i.earned < i.possible; })
           .forEach(function (i) { studentsBelow[i.name] = true; });
    return [std, o.trait, o.group, Object.keys(o.items).sort(function (a, b) { return a - b; }).join(', '),
      m, t.tier, STANDARD_MOVES[std] || 'Reteach using the guided questions for this trait.',
      Object.keys(studentsBelow).length];
  }).sort(function (a, b) { return a[4] - b[4]; });   // lowest mastery first

  writeTable_(sh, 4, ['Standard', 'Rubric Trait', 'Standard Group', 'Items on this form',
    'Class Mastery', 'Recommended Tier', 'Instructional Move', 'Students missing ≥1 item'], rows);
  var mr = sh.getRange(5, 5, rows.length, 1);
  mr.setNumberFormat('0%'); pctBands_(sh, mr);
  sh.setFrozenRows(4);
  sh.setColumnWidth(1, 100); sh.setColumnWidth(2, 170); sh.setColumnWidth(3, 150);
  sh.setColumnWidth(7, 480); sh.setColumnWidth(8, 130);
  sh.getRange(5, 1, rows.length, 8).setWrap(true).setVerticalAlignment('top');
}

/* ── 6. VOCABULARY FLAGS ─────────────────────────────────────────────── */
function buildVocab_(ss, d) {
  var sh = resetSheet_(ss, SH.VOCAB);
  titleRow_(sh, 'Vocabulary Flags — student-reported lexical demand',
    'Students flagged these items as hard because of an unfamiliar word. This is self-reported, so it shows where students ' +
    'KNOW they are stuck — pair it with the Item Analysis to catch the words they did not realise they were missing. ' +
    'Teach the most-flagged terms first.', 6);
  if (!d.items.length) { emptyNotice_(sh, 6); return; }

  var byItem = {};
  d.items.filter(function (i) { return i.flagged; }).forEach(function (i) {
    byItem[i.n] = byItem[i.n] || {n: i.n, id: i.id, trait: i.trait, std: i.std, students: []};
    byItem[i.n].students.push(i.name);
  });
  var total = d.subs.length || 1;
  var rows = Object.keys(byItem).map(function (k) { return byItem[k]; })
    .sort(function (a, b) { return b.students.length - a.students.length; })
    .map(function (o) {
      return [o.n, o.id, o.trait, o.std, o.students.length, o.students.length / total, o.students.sort().join(', ')];
    });

  if (!rows.length) {
    sh.getRange(4, 1).setValue('No vocabulary flags were raised on this administration.')
      .setFontColor('#999999').setFontStyle('italic');
    sh.getRange(4, 1, 1, 6).merge();
    return;
  }
  writeTable_(sh, 4, ['Item #', 'Item', 'Trait', 'Standard', 'Students flagging', '% of class', 'Who flagged it'], rows);
  var pr = sh.getRange(5, 6, rows.length, 1);
  pr.setNumberFormat('0%');
  sh.setFrozenRows(4);
  sh.setColumnWidth(3, 170); sh.setColumnWidth(7, 420);
  sh.getRange(5, 1, rows.length, 7).setWrap(true).setVerticalAlignment('top');
}

/* ── 7. STUDENT x ITEM MATRIX ────────────────────────────────────────── */
function buildMatrix_(ss, d) {
  var sh = resetSheet_(ss, SH.MATRIX);
  titleRow_(sh, 'Student x Item Matrix',
    'Points earned on each item. Green = full credit, yellow = partial, red = none. ' +
    'Read down a column to find an item the whole class missed (reteach it). Read across a row to find a student\'s pattern.', 8, true);
  if (!d.items.length) { emptyNotice_(sh, 8); return; }

  var nums = [];
  d.items.forEach(function (i) { if (nums.indexOf(i.n) < 0) nums.push(i.n); });
  nums.sort(function (a, b) { return a - b; });

  var byStudent = {};
  d.items.forEach(function (i) {
    byStudent[i.name] = byStudent[i.name] || {};
    byStudent[i.name][i.n] = i;
  });
  var names = Object.keys(byStudent).sort();

  var head = ['Student', 'Total'].concat(nums.map(function (n) { return 'Q' + n; }));
  var rows = names.map(function (nm) {
    var tot = 0;
    var cells = nums.map(function (n) {
      var it = byStudent[nm][n];
      if (!it) return '';
      tot += it.earned;
      return it.earned;
    });
    return [nm, tot].concat(cells);
  });
  // trailing class-average row
  var avg = ['CLASS AVERAGE', ''].concat(nums.map(function (n) {
    var list = d.items.filter(function (i) { return i.n === n; });
    return Math.round(mean_(list.map(function (i) { return i.earned; })) * 100) / 100;
  }));
  rows.push(avg);

  writeTable_(sh, 4, head, rows);
  var body = sh.getRange(5, 3, rows.length, nums.length);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#fadbd8').setRanges([body]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(1).setBackground('#fdf1cf').setRanges([body]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(2).setBackground('#d4efdc').setRanges([body]).build()
  ];
  sh.__pendingRules = (sh.__pendingRules || []).concat(rules);
  // 1-point items: a 1 is full credit, so recolour those columns green
  var onePt = {};
  d.items.forEach(function (i) { if (i.possible === 1) onePt[i.n] = true; });
  nums.forEach(function (n, idx) {
    if (!onePt[n]) return;
    var col = sh.getRange(5, 3 + idx, rows.length, 1);
    sh.__pendingRules = (sh.__pendingRules || []).concat([
      SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(1).setBackground('#d4efdc').setRanges([col]).build()
    ]);
  });
  sh.setFrozenRows(4); sh.setFrozenColumns(2);
  sh.setColumnWidth(1, 170);
  sh.getRange(5 + rows.length - 1, 1, 1, head.length).setFontWeight('bold').setBackground('#eef3fb');
  for (var c = 3; c <= head.length; c++) sh.setColumnWidth(c, 42);
}

/* ── 8. INDIVIDUAL STUDENT REPORT ────────────────────────────────────── */
function buildStudentReport_(ss, d) {
  var sh = resetSheet_(ss, SH.STU);
  titleRow_(sh, 'Individual Student Report',
    'Pick a student in cell B4. This page is built for conferences, MTSS documentation, and family communication — ' +
    'use File → Print to produce a one-page PDF. Remember that a single 30-point practice task is one data point, not a placement decision.', 6);
  if (!d.subs.length) { emptyNotice_(sh, 6); return; }

  sh.getRange(4, 1).setValue('Select student:').setFontWeight('bold');
  var names = d.subs.map(function (s) { return s.name; });
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(names, true).setAllowInvalid(false).build();
  sh.getRange(4, 2).setDataValidation(rule).setValue(names[0])
    .setBackground('#eef3fb').setFontWeight('bold').setBorder(true, true, true, true, false, false);

  // Live formulas against the Submissions / Item Responses sheets.
  var q = "'" + SH.SUB + "'!";
  var rows = [
    ['Overall score',  '=IFERROR(INDEX(' + q + 'I:I,MATCH($B$4,' + q + 'C:C,0))&" / "&INDEX(' + q + 'J:J,MATCH($B$4,' + q + 'C:C,0))&" points","—")'],
    ['Overall percent','=IFERROR(INDEX(' + q + 'K:K,MATCH($B$4,' + q + 'C:C,0)),"—")'],
    ['Practice band',  '=IFERROR(INDEX(' + q + 'L:L,MATCH($B$4,' + q + 'C:C,0)),"—")'],
    ['Period',         '=IFERROR(INDEX(' + q + 'E:E,MATCH($B$4,' + q + 'C:C,0)),"—")'],
    ['Time on task (min)','=IFERROR(INDEX(' + q + 'H:H,MATCH($B$4,' + q + 'C:C,0)),"—")'],
    ['T1 Purpose & Organization','=IFERROR(INDEX(' + q + 'M:M,MATCH($B$4,' + q + 'C:C,0)),"—")'],
    ['T2 Evidence & Elaboration','=IFERROR(INDEX(' + q + 'N:N,MATCH($B$4,' + q + 'C:C,0)),"—")'],
    ['T3 Language Usage & Conventions','=IFERROR(INDEX(' + q + 'O:O,MATCH($B$4,' + q + 'C:C,0)),"—")'],
    ['Vocabulary words flagged','=IFERROR(INDEX(' + q + 'Y:Y,MATCH($B$4,' + q + 'C:C,0)),"—")']
  ];
  var r = writeTable_(sh, 6, ['Measure', 'Result'], rows.map(function (x) { return [x[0], x[1]]; }));
  sh.getRange(7, 2).setNumberFormat('@');
  sh.getRange(8, 2).setNumberFormat('0%');
  sh.getRange(12, 2, 3, 1).setNumberFormat('0%');
  pctBands_(sh, sh.getRange(12, 2, 3, 1));
  r += 1;

  sh.getRange(r, 1).setValue('Items this student did not earn full credit on')
    .setFontWeight('bold').setFontSize(11).setFontColor('#1a3a5c');
  sh.getRange(r, 1, 1, 6).merge();
  r++;
  // NOTE: inside QUERY, column letters are relative to the C:Q range, so
  // A=studentName, D=itemNumber, G=trait, I=standard, K=pointsEarned,
  // L=pointsPossible, M=outcome, N=studentResponse, O=correctAnswer.
  sh.getRange(r, 1).setFormula(
    '=IFERROR(QUERY(\'' + SH.ITEMS + '\'!C:Q,"select D,G,I,K,L,N,O where A = \'"&$B$4&"\' and M <> \'full\' order by D",1),"All items earned full credit.")');
  sh.getRange(r, 1, 1, 7).setFontWeight('bold');

  sh.setColumnWidth(1, 250); sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 110); sh.setColumnWidth(4, 90);
  sh.setColumnWidth(5, 90); sh.setColumnWidth(6, 340); sh.setColumnWidth(7, 340);
}
