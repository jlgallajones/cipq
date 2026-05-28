// ─────────────────────────────────────────────────────────────────────────────
// SURVEY MODULE  —  Standalone quantitative perception layer
// Analytically separate from the CIPQ engine. Survey data is NEVER fed into
// CIPQ indicators, weighted policy pressures, or severity analytics.
// ─────────────────────────────────────────────────────────────────────────────

'use strict';

// ── Constants ────────────────────────────────────────────────────────────────
const SURVEY_TABLE     = 'survey_responses';
const SURVEY_Q_TABLE   = 'survey_questions';
const LIKERT_LABELS    = { 1:'Strongly Disagree', 2:'Disagree', 3:'Neutral', 4:'Agree', 5:'Strongly Agree' };
const CIPQ_DOMAINS     = ['Creation', 'Production', 'Distribution', 'Access'];
const SURVEY_TYPES     = { LIKERT:'likert', OPEN:'open_ended', MULTIPLE_CHOICE:'multiple_choice', CHECKBOX:'checkbox' };

// ── In-memory store ──────────────────────────────────────────────────────────
let surveyQuestions  = [];   // { id, code, text, question_type, cipq_domain, category }
let surveyResponses  = [];   // { id, respondent_id, respondent_group, region, question_code, score, answer_text, source_id, recorded_at }

// ── Supabase helpers (shared client from app.js) ─────────────────────────────
function getSB() { return window.supabaseClient || null; }
function getSBUser() { return window.currentUser || null; }

function normalizeQuestionType(value) {
  const type = String(value || '').trim().toLowerCase().replace(/[\s-]+/g, '_');
  if (['open', 'open_ended', 'openended', 'text', 'sentence', 'qualitative', 'insight'].includes(type))
    return SURVEY_TYPES.OPEN;
  if (['multiple_choice', 'multiplechoice', 'mc', 'single_choice', 'singlechoice', 'select_one', 'select'].includes(type))
    return SURVEY_TYPES.MULTIPLE_CHOICE;
  if (['checkbox', 'check', 'multi_select', 'multiselect', 'select_all', 'select_multiple'].includes(type))
    return SURVEY_TYPES.CHECKBOX;
  return SURVEY_TYPES.LIKERT;
}

function normalizeSurveyDomain(value) {
  const domain = String(value || '').trim().toLowerCase();
  return CIPQ_DOMAINS.find(item => item.toLowerCase() === domain) || null;
}

function surveyQuestionTypeLabel(question) {
  const t = surveyQuestionType(question);
  if (t === SURVEY_TYPES.OPEN) return 'open_ended';
  if (t === SURVEY_TYPES.MULTIPLE_CHOICE) return 'multiple_choice';
  if (t === SURVEY_TYPES.CHECKBOX) return 'checkbox';
  return 'rating';
}

function surveyQuestionType(question) {
  return normalizeQuestionType(question?.question_type || question?.type || SURVEY_TYPES.LIKERT);
}

function isLikertQuestion(question) {
  return surveyQuestionType(question) === SURVEY_TYPES.LIKERT;
}

function isMultipleChoiceQuestion(question) {
  const t = surveyQuestionType(question);
  return t === SURVEY_TYPES.MULTIPLE_CHOICE || t === SURVEY_TYPES.CHECKBOX;
}

function questionForCode(code) {
  return surveyQuestions.find(q => q.code === code);
}

function numericScore(value) {
  const score = Number(value);
  return Number.isInteger(score) && score >= 1 && score <= 5 ? score : null;
}

function likertResponses() {
  return surveyResponses.filter(r => numericScore(r.score) !== null && isLikertQuestion(questionForCode(r.question_code)));
}

function multipleChoiceResponses() {
  return surveyResponses.filter(r => {
    const q = questionForCode(r.question_code);
    return isMultipleChoiceQuestion(q) && String(r.answer_text || '').trim();
  });
}

function openEndedResponses() {
  return surveyResponses.filter(r => {
    const q = questionForCode(r.question_code);
    if (!q) return false;
    const t = surveyQuestionType(q);
    return t === SURVEY_TYPES.OPEN && String(r.answer_text || '').trim();
  });
}

function formatSurveySaveError(error) {
  const message = error?.message || String(error || 'Unknown error.');
  if (/relation .*survey_|does not exist|schema cache|could not find the table/i.test(message)) {
    return `${message} Run the survey table section in supabase_schema.sql, then refresh this dashboard.`;
  }
  if (/question_type_check|violates check constraint/i.test(message)) {
    return `Question type not allowed by the database. Re-run supabase_schema.sql in Supabase to add support for multiple_choice and checkbox types, then try again.`;
  }
  return message;
}

// ─────────────────────────────────────────────────────────────────────────────
// DATA LOADING
// ─────────────────────────────────────────────────────────────────────────────

// Fetches ALL survey_responses rows in pages of 1,000 to bypass Supabase's
// default 1,000-row cap, which silently truncates large datasets on refresh.
async function fetchAllSurveyResponses(sb) {
  const PAGE_SIZE = 1000;
  let allRows = [];
  let from = 0;

  while (true) {
    const { data, error } = await sb
      .from(SURVEY_TABLE)
      .select('*')
      .order('recorded_at')
      .range(from, from + PAGE_SIZE - 1);

    if (error) throw error;
    if (!data || data.length === 0) break;

    allRows = allRows.concat(data);
    if (data.length < PAGE_SIZE) break; // last page — no more rows
    from += PAGE_SIZE;
  }

  return allRows;
}

async function loadSurveyData() {
  const sb = getSB();
  surveyShowStatus('Loading survey data…', false);
  if (!sb) {
    renderSurveyTab();
    surveyShowStatus('Supabase not configured — using local data only.', true);
    return;
  }

  try {
    const [qRes, allResponses] = await Promise.all([
      sb.from(SURVEY_Q_TABLE).select('*').order('code'),
      fetchAllSurveyResponses(sb)
    ]);
    if (qRes.error) throw qRes.error;
    surveyQuestions = qRes.data || [];
    surveyResponses = allResponses;
    renderSurveyTab();
    surveyShowStatus(`Survey loaded: ${surveyQuestions.length} questions, ${surveyResponses.length} responses.`, false);
  } catch (err) {
    surveyShowStatus(`Survey load error: ${err.message || err}`, true);
    renderSurveyTab();
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// CSV IMPORT — QUESTIONS
// Format: code, text, question_type (rating/open_ended), cipq_domain (optional), category (optional)
// ─────────────────────────────────────────────────────────────────────────────

function handleSurveyQuestionsImport(event) {
  if (!window.canWrite || !window.canWrite()) {
    surveyShowStatus('Sign in to import survey questions.', true);
    event.target.value = '';
    return;
  }
  const file = event.target.files[0];
  if (!file) return;
  event.target.value = '';

  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: async (results) => {
      if (!results.data.length) { surveyShowStatus('CSV appears empty.', true); return; }

      const firstRow = results.data[0];
      const cols = Object.keys(firstRow).map(k => k.toLowerCase().trim());
      if (!cols.includes('code') || !cols.includes('text')) {
        surveyShowStatus('Questions CSV must have at least "code" and "text" columns.', true);
        return;
      }

      const questions = results.data.map(row => {
        const get = (k) => row[Object.keys(row).find(r => r.toLowerCase().trim() === k)] || '';
        const domain = get('cipq_domain') || get('domain') || null;
        return {
          code:        (get('code') || '').trim().toUpperCase(),
          text:        (get('text') || '').trim(),
          question_type: normalizeQuestionType(get('question_type') || get('type')),
          cipq_domain: normalizeSurveyDomain(domain),
          category:    (get('category') || '').trim() || 'Imported'
        };
      }).filter(q => q.code && q.text);

      if (!questions.length) { surveyShowStatus('No valid question rows found.', true); return; }

      await saveSurveyQuestions(questions);
      renderSurveyTab();
      surveyShowStatus(`${questions.length} question${questions.length !== 1 ? 's' : ''} imported.`, false);
    },
    error: (err) => surveyShowStatus(`CSV parse error: ${err.message}`, true)
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// CSV IMPORT — RESPONSES (long format)
// Format: respondent_id, respondent_group, region, source_id, question_code, score, answer_text
// One row = one respondent's answer to one question.
// ─────────────────────────────────────────────────────────────────────────────

function handleSurveyResponsesImport(event) {
  if (!window.canWrite || !window.canWrite()) {
    surveyShowStatus('Sign in to import survey responses.', true);
    event.target.value = '';
    return;
  }
  const file = event.target.files[0];
  if (!file) return;
  event.target.value = '';

  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: async (results) => {
      if (!results.data.length) { surveyShowStatus('CSV appears empty.', true); return; }

      const get = (row, k) => {
        const key = Object.keys(row).find(r => r.toLowerCase().trim() === k);
        return key ? (row[key] || '').trim() : '';
      };

      const rows = [];
      const unknownQuestions = new Map();

      for (const row of results.data) {
        const question_code = (get(row, 'question_code') || get(row, 'question code') || get(row, 'code')).toUpperCase();
        const scoreRaw      = get(row, 'score');
        const answer_text   = get(row, 'answer_text') || get(row, 'answer') || get(row, 'insight') || get(row, 'response_text') || get(row, 'text_response');
        const score         = numericScore(scoreRaw);
        const existingQ     = questionForCode(question_code);
        const rawTypeHint   = get(row, 'question_type') || get(row, 'question_type_hint') || get(row, 'type');
        const typeHint      = rawTypeHint ? normalizeQuestionType(rawTypeHint) : '';
        const inferredType  = existingQ ? surveyQuestionType(existingQ) : (typeHint || (answer_text && score === null ? SURVEY_TYPES.OPEN : SURVEY_TYPES.LIKERT));

        if (!question_code) continue;

        if (inferredType === SURVEY_TYPES.LIKERT && score === null) continue;
        if (inferredType === SURVEY_TYPES.OPEN && !answer_text) continue;
        if ((inferredType === SURVEY_TYPES.MULTIPLE_CHOICE || inferredType === SURVEY_TYPES.CHECKBOX) && !answer_text) continue;

        if (!existingQ) {
          unknownQuestions.set(question_code, inferredType);
        }

        const isTextBased = inferredType === SURVEY_TYPES.OPEN || inferredType === SURVEY_TYPES.MULTIPLE_CHOICE || inferredType === SURVEY_TYPES.CHECKBOX;
        rows.push({
          respondent_id:    get(row, 'respondent_id'),
          respondent_group: get(row, 'respondent_group'),
          region:           get(row, 'region'),
          source_id:        get(row, 'source_id'),
          question_code,
          score: inferredType === SURVEY_TYPES.LIKERT ? score : null,
          answer_text: isTextBased ? answer_text : '',
          recorded_at: new Date().toISOString()
        });
      }

      if (!rows.length) {
        surveyShowStatus('No valid rows found. Use "score" (1-5) for rate questions or "answer_text" for open-ended questions.', true);
        return;
      }

      // Auto-create placeholder questions for any unrecognised codes
      if (unknownQuestions.size) {
        await saveSurveyQuestions([...unknownQuestions.entries()].map(([code, question_type]) => ({
          code, text: code, question_type, cipq_domain: null, category: 'Imported'
        })));
        surveyShowStatus(`Note: ${unknownQuestions.size} unknown question code(s) auto-created. Edit them in the question list.`, false);
      }

      await saveSurveyResponses(rows);
    },
    error: (err) => surveyShowStatus(`CSV parse error: ${err.message}`, true)
  });
}

// Keep the old combined handler as an alias (for backwards compat)
const handleSurveyImport = handleSurveyResponsesImport;

async function saveSurveyQuestions(questions) {
  const sb = getSB();
  const normalizedQuestions = questions.map(q => ({
    ...q,
    code: String(q.code || '').trim().toUpperCase(),
    question_type: normalizeQuestionType(q.question_type || q.type)
  }));
  if (!sb) { surveyQuestions.push(...normalizedQuestions); return; }
  const { data, error } = await sb.from(SURVEY_Q_TABLE).upsert(normalizedQuestions, { onConflict: 'code' }).select();
  if (error) { surveyShowStatus(`Could not save questions: ${formatSurveySaveError(error)}`, true); return; }
  const newCodes = new Set((data || []).map(q => q.code));
  surveyQuestions = [...surveyQuestions.filter(q => !newCodes.has(q.code)), ...(data || [])];
}

async function saveSurveyResponses(rows) {
  const sb = getSB();
  if (!sb) {
    surveyResponses.push(...rows);
    renderSurveyTab();
    surveyShowStatus(`${rows.length} responses imported (local only — not connected to Supabase).`, false);
    return;
  }

  const { data, error } = await sb.from(SURVEY_TABLE).insert(rows).select();
  if (error) {
    surveyShowStatus(`Could not save responses: ${formatSurveySaveError(error)}`, true);
    return;
  }
  surveyResponses.push(...(data || rows));
  renderSurveyTab();
  surveyShowStatus(`${rows.length} responses imported and saved.`, false);
}

// ─────────────────────────────────────────────────────────────────────────────
// MANUAL RESPONSE ENTRY
// ─────────────────────────────────────────────────────────────────────────────

async function addSurveyResponse() {
  if (!window.canWrite || !window.canWrite()) {
    surveyShowStatus('Sign in to add survey responses.', true);
    return;
  }
  const respondent_id    = (document.getElementById('sv_respondent_id')?.value || '').trim();
  const respondent_group = (document.getElementById('sv_respondent_group')?.value || '').trim();
  const region           = (document.getElementById('sv_region')?.value || '').trim();
  const source_id        = (document.getElementById('sv_source_id')?.value || '').trim();
  const question_code    = (document.getElementById('sv_question_code')?.value || '').trim();
  const scoreRaw         = numericScore(document.getElementById('sv_score')?.value || '');
  const answer_text      = (document.getElementById('sv_answer_text')?.value || '').trim();
  const selectedQuestion = questionForCode(question_code);
  const questionType     = surveyQuestionType(selectedQuestion);

  if (!question_code) { surveyShowStatus('Select or enter a question code.', true); return; }
  if (questionType === SURVEY_TYPES.LIKERT && scoreRaw === null) { surveyShowStatus('Score must be 1-5.', true); return; }
  if (questionType === SURVEY_TYPES.OPEN && !answer_text) { surveyShowStatus('Open-ended answers require participant insight text.', true); return; }
  if ((questionType === SURVEY_TYPES.MULTIPLE_CHOICE || questionType === SURVEY_TYPES.CHECKBOX) && !answer_text) { surveyShowStatus('Please enter the selected choice(s).', true); return; }

  // Ensure question exists
  if (!selectedQuestion) {
    await saveSurveyQuestions([{ code: question_code, text: question_code, question_type: SURVEY_TYPES.LIKERT, cipq_domain: null, category: 'Manual' }]);
  }

  await saveSurveyResponses([{
    respondent_id, respondent_group, region, source_id,
    question_code,
    score: questionType === SURVEY_TYPES.LIKERT ? scoreRaw : null,
    answer_text: (questionType === SURVEY_TYPES.OPEN || questionType === SURVEY_TYPES.MULTIPLE_CHOICE || questionType === SURVEY_TYPES.CHECKBOX) ? answer_text : '',
    recorded_at: new Date().toISOString()
  }]);

  // Clear entry fields
  ['sv_respondent_id','sv_score','sv_answer_text'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
}

async function addSurveyQuestion() {
  if (!window.canWrite || !window.canWrite()) {
    surveyShowStatus('Sign in to add survey questions.', true);
    return;
  }
  const code       = (document.getElementById('sq_code')?.value || '').trim().toUpperCase();
  const text       = (document.getElementById('sq_text')?.value || '').trim();
  const question_type = normalizeQuestionType(document.getElementById('sq_question_type')?.value || '');
  const cipq_domain = normalizeSurveyDomain(document.getElementById('sq_domain')?.value || '');
  const category   = (document.getElementById('sq_category')?.value || '').trim() || 'General';

  if (!code || !text) { surveyShowStatus('Question code and text are required.', true); return; }
  if (surveyQuestions.find(q => q.code === code)) { surveyShowStatus(`Question ${code} already exists.`, true); return; }

  await saveSurveyQuestions([{ code, text, question_type, cipq_domain, category }]);
  renderSurveyAnalyst();
  surveyShowStatus(`Question ${code} added.`, false);
  ['sq_code','sq_text'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
}

async function deleteSurveyQuestion(code) {
  if (!window.canWrite || !window.canWrite()) return;
  if (!confirm(`Delete question ${code} and all its responses?`)) return;
  const sb = getSB();
  if (sb) {
    await sb.from(SURVEY_TABLE).delete().eq('question_code', code);
    await sb.from(SURVEY_Q_TABLE).delete().eq('code', code);
  }
  surveyQuestions = surveyQuestions.filter(q => q.code !== code);
  surveyResponses = surveyResponses.filter(r => r.question_code !== code);
  renderSurveyTab();
  surveyShowStatus(`Question ${code} and its responses deleted.`, false);
}

async function clearAllSurveyData() {
  if (!window.canWrite || !window.canWrite()) return;
  if (!confirm('Delete ALL survey questions and responses? This cannot be undone.')) return;
  const sb = getSB();
  if (sb) {
    await sb.from(SURVEY_TABLE).delete().neq('id', '00000000-0000-0000-0000-000000000000');
    await sb.from(SURVEY_Q_TABLE).delete().neq('id', '00000000-0000-0000-0000-000000000000');
  }
  surveyQuestions = [];
  surveyResponses = [];
  renderSurveyTab();
  surveyShowStatus('All survey data cleared.', false);
}

// ─────────────────────────────────────────────────────────────────────────────
// STATISTICS ENGINE
// ─────────────────────────────────────────────────────────────────────────────

function calcStats(scores) {
  scores = scores.map(numericScore).filter(score => score !== null);
  if (!scores.length) return null;
  const n   = scores.length;
  const sum = scores.reduce((a, b) => a + b, 0);
  const mean = sum / n;
  const sorted = [...scores].sort((a, b) => a - b);
  const median = n % 2 === 0 ? (sorted[n/2-1] + sorted[n/2]) / 2 : sorted[Math.floor(n/2)];
  const variance = scores.reduce((a, b) => a + (b - mean) ** 2, 0) / n;
  const sd = Math.sqrt(variance);
  const freq = {1:0,2:0,3:0,4:0,5:0};
  scores.forEach(s => freq[s] = (freq[s]||0) + 1);
  const modeEntry = Object.entries(freq).sort((a,b) => b[1]-a[1])[0];
  const agree_pct = ((freq[4]+freq[5])/n*100).toFixed(1);
  const disagree_pct = ((freq[1]+freq[2])/n*100).toFixed(1);
  return { n, mean: mean.toFixed(2), median: median.toFixed(2), sd: sd.toFixed(2), mode: modeEntry[0], freq, agree_pct, disagree_pct };
}

function questionStats() {
  return surveyQuestions.filter(isLikertQuestion).map(q => {
    const scores = likertResponses().filter(r => r.question_code === q.code).map(r => r.score);
    return { ...q, stats: calcStats(scores) };
  }).filter(q => q.stats);
}

function domainStats() {
  return CIPQ_DOMAINS.map(domain => {
    const qCodes = new Set(surveyQuestions.filter(q => q.cipq_domain === domain && isLikertQuestion(q)).map(q => q.code));
    const scores = likertResponses().filter(r => qCodes.has(r.question_code)).map(r => r.score);
    return { domain, stats: calcStats(scores), qCount: qCodes.size };
  }).filter(d => d.stats);
}

function groupStats(groupField = 'respondent_group') {
  const ratedRows = likertResponses();
  const groups = [...new Set(ratedRows.map(r => r[groupField] || 'Unspecified'))];
  return groups.map(group => {
    const groupResps = ratedRows.filter(r => (r[groupField] || 'Unspecified') === group);
    const scores = groupResps.map(r => r.score);
    const respondents = new Set(groupResps.map(r => r.respondent_id).filter(Boolean)).size;
    return { group, respondents, stats: calcStats(scores) };
  }).filter(g => g.stats).sort((a,b) => parseFloat(b.stats.mean) - parseFloat(a.stats.mean));
}

function openEndedQuestionSummaries() {
  return surveyQuestions.filter(q => surveyQuestionType(q) === SURVEY_TYPES.OPEN).map(q => {
    const responses = openEndedResponses().filter(r => r.question_code === q.code);
    return { ...q, responses };
  }).filter(q => q.responses.length);
}

function multipleChoiceQuestionSummaries() {
  return surveyQuestions.filter(isMultipleChoiceQuestion).map(q => {
    const responses = multipleChoiceResponses().filter(r => r.question_code === q.code);
    // Tally answer frequencies
    const tally = {};
    responses.forEach(r => {
      const ans = String(r.answer_text || '').trim();
      if (ans) tally[ans] = (tally[ans] || 0) + 1;
    });
    const choices = Object.entries(tally)
      .sort((a, b) => b[1] - a[1])
      .map(([label, count]) => ({ label, count, pct: responses.length ? ((count / responses.length) * 100).toFixed(1) : '0.0' }));
    return { ...q, responses, choices, n: responses.length };
  }).filter(q => q.n > 0);
}

// ─────────────────────────────────────────────────────────────────────────────
// EXPORT
// ─────────────────────────────────────────────────────────────────────────────

function exportSurveyCSV() {
  if (!surveyResponses.length) { surveyShowStatus('No survey responses to export.', true); return; }
  const rows = surveyResponses.map(r => ({
    respondent_id:    r.respondent_id || '',
    respondent_group: r.respondent_group || '',
    region:           r.region || '',
    source_id:        r.source_id || '',
    question_code:    r.question_code,
    question_text:    questionForCode(r.question_code)?.text || '',
    question_type:    surveyQuestionTypeLabel(questionForCode(r.question_code)),
    cipq_domain:      questionForCode(r.question_code)?.cipq_domain || '',
    score:            numericScore(r.score) ?? '',
    score_label:      LIKERT_LABELS[numericScore(r.score)] || '',
    answer_text:      r.answer_text || '',
    recorded_at:      r.recorded_at || ''
  }));
  const csv = Papa.unparse(rows);
  downloadText(csv, 'survey_responses.csv', 'text/csv');
}

function exportSurveySummaryCSV() {
  const qs = questionStats();
  const openQs = openEndedQuestionSummaries();
  const mcQs = multipleChoiceQuestionSummaries();
  if (!qs.length && !openQs.length && !mcQs.length) { surveyShowStatus('No survey data to export.', true); return; }
  const rows = qs.map(q => ({
    code: q.code,
    text: q.text,
    question_type: 'rating',
    cipq_domain: q.cipq_domain || '',
    category: q.category || '',
    n: q.stats.n,
    mean: q.stats.mean,
    median: q.stats.median,
    sd: q.stats.sd,
    mode: q.stats.mode,
    agree_pct: q.stats.agree_pct,
    disagree_pct: q.stats.disagree_pct,
    freq_1: q.stats.freq[1], freq_2: q.stats.freq[2], freq_3: q.stats.freq[3],
    freq_4: q.stats.freq[4], freq_5: q.stats.freq[5],
    insight_count: '',
    choices: ''
  })).concat(openQs.map(q => ({
    code: q.code,
    text: q.text,
    question_type: SURVEY_TYPES.OPEN,
    cipq_domain: q.cipq_domain || '',
    category: q.category || '',
    n: '',
    mean: '',
    median: '',
    sd: '',
    mode: '',
    agree_pct: '',
    disagree_pct: '',
    freq_1: '', freq_2: '', freq_3: '', freq_4: '', freq_5: '',
    insight_count: q.responses.length,
    choices: ''
  }))).concat(mcQs.map(q => ({
    code: q.code,
    text: q.text,
    question_type: surveyQuestionType(q),
    cipq_domain: q.cipq_domain || '',
    category: q.category || '',
    n: q.n,
    mean: '',
    median: '',
    sd: '',
    mode: '',
    agree_pct: '',
    disagree_pct: '',
    freq_1: '', freq_2: '', freq_3: '', freq_4: '', freq_5: '',
    insight_count: '',
    choices: q.choices.map(choice => `${choice.label}: ${choice.count} (${choice.pct}%)`).join(' | ')
  })));
  const csv = Papa.unparse(rows);
  downloadText(csv, 'survey_summary.csv', 'text/csv');
}

function exportSurveyGroupSummaryCSV() {
  const respondentGroups = groupStats('respondent_group').map(item => ({
    grouping: 'respondent_group',
    group: item.group,
    respondents: item.respondents || '',
    responses: item.stats.n,
    mean: item.stats.mean,
    median: item.stats.median,
    sd: item.stats.sd,
    mode: item.stats.mode,
    agree_pct: item.stats.agree_pct,
    disagree_pct: item.stats.disagree_pct,
    freq_1: item.stats.freq[1],
    freq_2: item.stats.freq[2],
    freq_3: item.stats.freq[3],
    freq_4: item.stats.freq[4],
    freq_5: item.stats.freq[5]
  }));
  const regions = groupStats('region').map(item => ({
    grouping: 'region',
    group: item.group,
    respondents: item.respondents || '',
    responses: item.stats.n,
    mean: item.stats.mean,
    median: item.stats.median,
    sd: item.stats.sd,
    mode: item.stats.mode,
    agree_pct: item.stats.agree_pct,
    disagree_pct: item.stats.disagree_pct,
    freq_1: item.stats.freq[1],
    freq_2: item.stats.freq[2],
    freq_3: item.stats.freq[3],
    freq_4: item.stats.freq[4],
    freq_5: item.stats.freq[5]
  }));
  const rows = respondentGroups.concat(regions);
  if (!rows.length) { surveyShowStatus('No Likert responses available for group summaries.', true); return; }
  downloadText(Papa.unparse(rows), 'survey_group_summary.csv', 'text/csv');
}

function surveyReportDateStamp() {
  return new Date().toISOString().slice(0, 10);
}

function surveySentiment(meanValue) {
  const mean = parseFloat(meanValue);
  if (mean >= 4.0) return 'Positive perception';
  if (mean >= 3.0) return 'Neutral / mixed perception';
  return 'Negative perception';
}

function surveyWordCell(value, style = '') {
  return `<td style="border:1px solid #c8bfae;padding:6px;vertical-align:top;${style}">${escHtml(value ?? '')}</td>`;
}

function surveyWordHeader(labels) {
  return `<tr>${labels.map(label => `<th style="border:1px solid #c8bfae;padding:6px;background:#f5f0e8;text-align:left;">${escHtml(label)}</th>`).join('')}</tr>`;
}

function surveyLikertSummary(stats) {
  return [1, 2, 3, 4, 5]
    .map(score => `${score}: ${stats.freq[score] || 0}`)
    .join(' | ');
}

function surveyDistributionTable(stats) {
  return `<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin:8px 0 16px;">
    ${surveyWordHeader(['Score', 'Label', 'Responses', 'Percent'])}
    ${[1, 2, 3, 4, 5].map(score => {
      const count = stats.freq[score] || 0;
      const pct = stats.n ? ((count / stats.n) * 100).toFixed(1) : '0.0';
      return `<tr>
        ${surveyWordCell(score)}
        ${surveyWordCell(LIKERT_LABELS[score])}
        ${surveyWordCell(count)}
        ${surveyWordCell(`${pct}%`)}
      </tr>`;
    }).join('')}
  </table>`;
}

function buildSurveyReportText() {
  if (!surveyResponses.length) return '';

  const qs = questionStats();
  const openQs = openEndedQuestionSummaries();
  const mcQs = multipleChoiceQuestionSummaries();
  const dStats = domainStats();
  const gStats = groupStats('respondent_group');
  const totalRespondents = new Set(surveyResponses.map(r => r.respondent_id).filter(Boolean)).size;
  const ratedRows = likertResponses();
  const openRows = openEndedResponses();
  const overall = calcStats(ratedRows.map(r => r.score));

  const lines = [];
  lines.push('Survey Analytics Report');
  lines.push(`Generated: ${new Date().toISOString()}`);
  lines.push('');
  lines.push('Analytical Note');
  lines.push('Survey results are a separate perception layer. Likert ratings are summarized quantitatively; open-ended answers are preserved as participant insights. They are not encoded as CIPQ indicators and do not affect CIPQ weighted policy pressures, frequencies, or severity analytics.');
  lines.push('');
  lines.push('Dataset Totals');
  lines.push(`- Responses: ${surveyResponses.length}`);
  lines.push(`- Likert score responses: ${ratedRows.length}`);
  lines.push(`- Open-ended insights: ${openRows.length}`);
  lines.push(`- Respondents: ${totalRespondents || 'Not specified'}`);
  lines.push(`- Questions: ${surveyQuestions.length}`);
  if (overall) {
    lines.push(`- Overall mean: ${overall.mean} / 5.00`);
    lines.push(`- Sentiment: ${surveySentiment(overall.mean)}`);
    lines.push(`- Median: ${overall.median}`);
    lines.push(`- Standard deviation: ${overall.sd}`);
    lines.push(`- Agree/Strongly Agree: ${overall.agree_pct}%`);
    lines.push(`- Disagree/Strongly Disagree: ${overall.disagree_pct}%`);
    lines.push(`- Likert distribution: ${surveyLikertSummary(overall)}`);
  }

  if (dStats.length) {
    lines.push('');
    lines.push('Perception by Related CIPQ Domain');
    lines.push('Domain | Questions | Responses | Mean | Median | SD | Agree %');
    dStats.forEach(item => {
      lines.push(`${item.domain} | ${item.qCount} | ${item.stats.n} | ${item.stats.mean} | ${item.stats.median} | ${item.stats.sd} | ${item.stats.agree_pct}%`);
    });
    lines.push('Domain mapping is for interpretive comparison only.');
  }

  if (gStats.length) {
    lines.push('');
    lines.push('Respondent-Group Summary');
    lines.push('Group | Respondents | Responses | Mean | SD | Agree %');
    gStats.forEach(item => {
      lines.push(`${item.group} | ${item.respondents || '-'} | ${item.stats.n} | ${item.stats.mean} | ${item.stats.sd} | ${item.stats.agree_pct}%`);
    });
  }

  if (qs.length) {
    lines.push('');
    lines.push('Question-Level Statistics');
    lines.push('Code | Question | Related Domain | Category | N | Mean | Median | SD | Agree % | Distribution');
    qs.forEach(q => {
      lines.push(`${q.code} | ${q.text} | ${q.cipq_domain || '-'} | ${q.category || 'General'} | ${q.stats.n} | ${q.stats.mean} | ${q.stats.median} | ${q.stats.sd} | ${q.stats.agree_pct}% | ${surveyLikertSummary(q.stats)}`);
    });
  }

  if (mcQs.length) {
    lines.push('');
    lines.push('Multiple Choice and Checkbox Results');
    lines.push('Code | Question | Type | Related Domain | Responses | Choices');
    mcQs.forEach(q => {
      lines.push(`${q.code} | ${q.text} | ${surveyQuestionType(q)} | ${q.cipq_domain || '-'} | ${q.n} | ${q.choices.map(choice => `${choice.label}: ${choice.count} (${choice.pct}%)`).join('; ')}`);
    });
  }

  if (openQs.length) {
    lines.push('');
    lines.push('Open-Ended Participant Insights');
    openQs.forEach(q => {
      lines.push(`${q.code} | ${q.text} | ${q.responses.length} insight${q.responses.length !== 1 ? 's' : ''}`);
      q.responses.forEach(r => {
        const respondent = [r.respondent_id, r.respondent_group, r.region].filter(Boolean).join(' / ') || 'Unspecified respondent';
        lines.push(`- ${respondent}: ${r.answer_text}`);
      });
    });
  }

  return lines.join('\r\n');
}

function buildSurveyReportWordHtml() {
  const qs = questionStats();
  const openQs = openEndedQuestionSummaries();
  const mcQs = multipleChoiceQuestionSummaries();
  const dStats = domainStats();
  const gStats = groupStats('respondent_group');
  const totalRespondents = new Set(surveyResponses.map(r => r.respondent_id).filter(Boolean)).size;
  const ratedRows = likertResponses();
  const openRows = openEndedResponses();
  const overall = calcStats(ratedRows.map(r => r.score));

  const domainRows = dStats.map(item => `<tr>
    ${surveyWordCell(item.domain)}
    ${surveyWordCell(item.qCount)}
    ${surveyWordCell(item.stats.n)}
    ${surveyWordCell(item.stats.mean)}
    ${surveyWordCell(item.stats.median)}
    ${surveyWordCell(item.stats.sd)}
    ${surveyWordCell(`${item.stats.agree_pct}%`)}
    ${surveyWordCell(surveyLikertSummary(item.stats))}
  </tr>`).join('');

  const groupRows = gStats.map(item => `<tr>
    ${surveyWordCell(item.group)}
    ${surveyWordCell(item.respondents || '-')}
    ${surveyWordCell(item.stats.n)}
    ${surveyWordCell(item.stats.mean)}
    ${surveyWordCell(item.stats.sd)}
    ${surveyWordCell(`${item.stats.agree_pct}%`)}
    ${surveyWordCell(surveyLikertSummary(item.stats))}
  </tr>`).join('');

  const questionRows = qs.map(q => `<tr>
    ${surveyWordCell(q.code)}
    ${surveyWordCell(q.text)}
    ${surveyWordCell(q.cipq_domain || '-')}
    ${surveyWordCell(q.category || 'General')}
    ${surveyWordCell(q.stats.n)}
    ${surveyWordCell(q.stats.mean)}
    ${surveyWordCell(q.stats.median)}
    ${surveyWordCell(q.stats.sd)}
    ${surveyWordCell(`${q.stats.agree_pct}%`)}
    ${surveyWordCell(surveyLikertSummary(q.stats))}
  </tr>`).join('');

  const multipleChoiceRows = mcQs.map(q => `<tr>
    ${surveyWordCell(q.code)}
    ${surveyWordCell(q.text)}
    ${surveyWordCell(surveyQuestionType(q))}
    ${surveyWordCell(q.cipq_domain || '-')}
    ${surveyWordCell(q.n)}
    ${surveyWordCell(q.choices.map(choice => `${choice.label}: ${choice.count} (${choice.pct}%)`).join(' | '))}
  </tr>`).join('');

  const openQuestionRows = openQs.map(q => `
    <h3>${escHtml(q.code)} — ${escHtml(q.text)}</h3>
    <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:14px;">
      ${surveyWordHeader(['Respondent', 'Group', 'Region', 'Insight'])}
      ${q.responses.map(r => `<tr>
        ${surveyWordCell(r.respondent_id || '-')}
        ${surveyWordCell(r.respondent_group || '-')}
        ${surveyWordCell(r.region || '-')}
        ${surveyWordCell(r.answer_text || '')}
      </tr>`).join('')}
    </table>
  `).join('');

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Survey Analytics Report</title>
<style>
  body { font-family: Arial, sans-serif; color: #1a1a2e; line-height: 1.45; }
  h1, h2, h3 { font-family: Georgia, serif; color: #1a1a2e; }
  h1 { font-size: 24pt; margin-bottom: 4px; }
  h2 { font-size: 16pt; margin-top: 22px; border-bottom: 1px solid #c8bfae; padding-bottom: 4px; }
  h3 { font-size: 12pt; margin-top: 16px; }
  .meta { color: #7a7065; margin-bottom: 18px; }
  .note { background: #f5f0e8; border: 1px solid #c8bfae; padding: 10px; margin: 12px 0 18px; }
  table { font-size: 9.5pt; }
</style>
</head>
<body>
  <h1>Survey Analytics Report</h1>
  <div class="meta">Generated ${escHtml(new Date().toISOString())} | ${ratedRows.length} Likert responses | ${openRows.length} open-ended insights</div>
  <div class="note">Survey results are a separate perception layer. Likert ratings are summarized quantitatively; open-ended answers are preserved as participant insights. They are not encoded as CIPQ indicators and do not affect CIPQ weighted policy pressures, frequencies, or severity analytics.</div>

  <h2>Dataset Totals</h2>
  <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;">
    ${surveyWordHeader(['Metric', 'Value'])}
    <tr>${surveyWordCell('Responses')}${surveyWordCell(surveyResponses.length)}</tr>
    <tr>${surveyWordCell('Likert score responses')}${surveyWordCell(ratedRows.length)}</tr>
    <tr>${surveyWordCell('Open-ended insights')}${surveyWordCell(openRows.length)}</tr>
    <tr>${surveyWordCell('Respondents')}${surveyWordCell(totalRespondents || 'Not specified')}</tr>
    <tr>${surveyWordCell('Questions')}${surveyWordCell(surveyQuestions.length)}</tr>
    ${overall ? `
    <tr>${surveyWordCell('Overall mean')}${surveyWordCell(`${overall.mean} / 5.00`)}</tr>
    <tr>${surveyWordCell('Sentiment')}${surveyWordCell(surveySentiment(overall.mean))}</tr>
    <tr>${surveyWordCell('Median')}${surveyWordCell(overall.median)}</tr>
    <tr>${surveyWordCell('Standard deviation')}${surveyWordCell(overall.sd)}</tr>
    <tr>${surveyWordCell('Agree / Strongly Agree')}${surveyWordCell(`${overall.agree_pct}%`)}</tr>
    <tr>${surveyWordCell('Disagree / Strongly Disagree')}${surveyWordCell(`${overall.disagree_pct}%`)}</tr>` : ''}
  </table>

  ${overall ? `<h2>Overall Likert Distribution</h2>${surveyDistributionTable(overall)}` : ''}

  ${dStats.length ? `
  <h2>Perception by Related CIPQ Domain</h2>
  <p>Domain mapping is comparative only and remains separate from CIPQ aggregation.</p>
  <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;">
    ${surveyWordHeader(['Domain', 'Questions', 'Responses', 'Mean', 'Median', 'SD', 'Agree %', 'Distribution'])}
    ${domainRows}
  </table>` : ''}

  ${gStats.length ? `
  <h2>Respondent-Group Summary</h2>
  <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;">
    ${surveyWordHeader(['Group', 'Respondents', 'Responses', 'Mean', 'SD', 'Agree %', 'Distribution'])}
    ${groupRows}
  </table>` : ''}

  ${qs.length ? `
  <h2>Question-Level Statistics</h2>
  <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;">
    ${surveyWordHeader(['Code', 'Question', 'Related Domain', 'Category', 'N', 'Mean', 'Median', 'SD', 'Agree %', 'Distribution'])}
    ${questionRows}
  </table>` : ''}

  ${mcQs.length ? `
  <h2>Multiple Choice and Checkbox Results</h2>
  <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;">
    ${surveyWordHeader(['Code', 'Question', 'Type', 'Related Domain', 'Responses', 'Choices'])}
    ${multipleChoiceRows}
  </table>` : ''}

  ${openQs.length ? `
  <h2>Open-Ended Participant Insights</h2>
  ${openQuestionRows}` : ''}
</body>
</html>`;
}

function downloadSurveyReport() {
  if (!surveyResponses.length) {
    surveyShowStatus('No survey data to export yet.', true);
    return;
  }
  downloadText(`\ufeff${buildSurveyReportWordHtml()}`, `Survey_Analytics_Report_${surveyReportDateStamp()}.doc`, 'application/msword;charset=utf-8');
  surveyShowStatus('Survey Word report downloaded.', false);
}

async function copySurveyTextToClipboard(text) {
  if (navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return;
    } catch (error) {
      // Fall back for local file usage where the Clipboard API may be blocked.
    }
  }

  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly', '');
  textarea.style.position = 'fixed';
  textarea.style.left = '-9999px';
  document.body.appendChild(textarea);
  textarea.select();
  document.execCommand('copy');
  textarea.remove();
}

async function copySurveyReport() {
  if (!surveyResponses.length) {
    surveyShowStatus('No survey report to copy yet.', true);
    return;
  }
  try {
    await copySurveyTextToClipboard(buildSurveyReportText());
    surveyShowStatus('Survey report copied to clipboard.', false);
  } catch (error) {
    surveyShowStatus(`Could not copy survey report: ${error.message || error}`, true);
  }
}

function downloadQuestionsTemplate() {
  const rows = [
    { code: 'SQ_01', text: 'Printing costs significantly affect publishing sustainability.', question_type: 'rating', cipq_domain: 'Production', category: 'Production Costs' },
    { code: 'SQ_02', text: 'Authors have adequate access to publishing opportunities.', question_type: 'rating', cipq_domain: 'Access', category: 'Author Access' },
    { code: 'SQ_03', text: 'What are the typical payment terms your company experiences? Public Sector', question_type: 'multiple_choice', cipq_domain: 'Distribution', category: 'Payment Terms' },
    { code: 'SQ_04', text: 'Which distribution challenges does your company currently face? (Select all that apply)', question_type: 'checkbox', cipq_domain: 'Distribution', category: 'Distribution Challenges' },
    { code: 'SQ_05', text: 'What participant insight best explains current distribution challenges?', question_type: 'open_ended', cipq_domain: 'Distribution', category: 'Participant Insights' },
  ];
  downloadText(Papa.unparse(rows), 'survey_questions_template.csv', 'text/csv');
}

function downloadResponsesTemplate() {
  const likertCodes = surveyQuestions.filter(isLikertQuestion).map(q => q.code);
  const mcCodes     = surveyQuestions.filter(q => surveyQuestionType(q) === SURVEY_TYPES.MULTIPLE_CHOICE).map(q => q.code);
  const cbCodes     = surveyQuestions.filter(q => surveyQuestionType(q) === SURVEY_TYPES.CHECKBOX).map(q => q.code);
  const openCodes   = surveyQuestions.filter(q => surveyQuestionType(q) === SURVEY_TYPES.OPEN).map(q => q.code);

  const qCodes   = likertCodes.length ? likertCodes : ['SQ_01', 'SQ_02'];
  const mcCode   = mcCodes[0]   || 'SQ_03';
  const cbCode   = cbCodes[0]   || 'SQ_04';
  const openCode = openCodes[0] || 'SQ_05';

  const rows = [
    // ── RATING (likert): fill score 1-5, leave answer_text blank ──
    { respondent_id: 'R001', respondent_group: 'Publisher',  region: 'NCR',         source_id: 'SURVEY_01', question_code: qCodes[0],         score: 4,  answer_text: '',                                                              question_type_hint: 'rating' },
    { respondent_id: 'R001', respondent_group: 'Publisher',  region: 'NCR',         source_id: 'SURVEY_01', question_code: qCodes[1]||qCodes[0], score: 3, answer_text: '',                                                              question_type_hint: 'rating' },
    { respondent_id: 'R002', respondent_group: 'Author',     region: 'Region IV-A', source_id: 'SURVEY_01', question_code: qCodes[0],         score: 2,  answer_text: '',                                                              question_type_hint: 'rating' },
    // ── MULTIPLE CHOICE: leave score blank, write the chosen option in answer_text (one row per respondent) ──
    { respondent_id: 'R001', respondent_group: 'Publisher',  region: 'NCR',         source_id: 'SURVEY_01', question_code: mcCode,            score: '',  answer_text: '30-60 days',                                                    question_type_hint: 'multiple_choice' },
    { respondent_id: 'R002', respondent_group: 'Author',     region: 'Region IV-A', source_id: 'SURVEY_01', question_code: mcCode,            score: '',  answer_text: 'More than 90 days',                                             question_type_hint: 'multiple_choice' },
    // ── CHECKBOX: leave score blank, one row per selected option per respondent ──
    { respondent_id: 'R001', respondent_group: 'Publisher',  region: 'NCR',         source_id: 'SURVEY_01', question_code: cbCode,            score: '',  answer_text: 'High logistics cost',                                           question_type_hint: 'checkbox' },
    { respondent_id: 'R001', respondent_group: 'Publisher',  region: 'NCR',         source_id: 'SURVEY_01', question_code: cbCode,            score: '',  answer_text: 'Limited provincial reach',                                      question_type_hint: 'checkbox' },
    { respondent_id: 'R002', respondent_group: 'Author',     region: 'Region IV-A', source_id: 'SURVEY_01', question_code: cbCode,            score: '',  answer_text: 'Limited provincial reach',                                      question_type_hint: 'checkbox' },
    // ── OPEN-ENDED: leave score blank, write free text in answer_text ──
    { respondent_id: 'R002', respondent_group: 'Author',     region: 'Region IV-A', source_id: 'SURVEY_01', question_code: openCode,          score: '',  answer_text: 'Provincial access depends heavily on informal networks.',        question_type_hint: 'open_ended' },
  ];
  downloadText(Papa.unparse(rows), 'survey_responses_template.csv', 'text/csv');
}

// Legacy alias
function downloadSurveyTemplate() { downloadResponsesTemplate(); }

function downloadText(content, filename, mime) {
  const bom = mime === 'text/csv' ? '\uFEFF' : '';
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([bom + content], { type: mime + ';charset=utf-8' }));
  a.download = filename;
  a.click();
}

// ─────────────────────────────────────────────────────────────────────────────
// RENDERING HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function likertBar(freq, n) {
  if (!n) return '';
  const colors = ['#c0392b','#e67e22','#7f8c8d','#2980b9','#27ae60'];
  const labels = ['SD','D','N','A','SA'];
  const segs = [1,2,3,4,5].map((v,i) => {
    const pct = ((freq[v]||0)/n*100).toFixed(1);
    if (!freq[v]) return '';
    return `<span title="${labels[i]}: ${freq[v]} (${pct}%)" style="display:inline-block;width:${pct}%;min-width:${freq[v]?2:0}px;height:12px;background:${colors[i]};vertical-align:middle;"></span>`;
  }).join('');
  return `<div style="display:flex;width:100%;height:12px;border-radius:3px;overflow:hidden;gap:1px;">${segs}</div>`;
}

function meanDot(mean) {
  const v = parseFloat(mean);
  let color = '#7f8c8d';
  if (v >= 4.0) color = '#27ae60';
  else if (v >= 3.0) color = '#2980b9';
  else if (v < 2.5) color = '#c0392b';
  return `<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${color};margin-right:5px;vertical-align:middle;"></span>`;
}

function domainPill(domain) {
  const colors = { Creation:'#8b4f9e', Production:'#c94a2e', Distribution:'#2a6b6e', Access:'#3a7a3a' };
  const c = colors[domain] || '#7a7065';
  return domain ? `<span style="display:inline-block;padding:0.15rem 0.6rem;border-radius:999px;background:${c}22;color:${c};font-size:0.67rem;font-family:'IBM Plex Mono',monospace;letter-spacing:0.05em;border:1px solid ${c}44;">${domain}</span>` : '';
}

function escHtml(s) {
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function surveyShowStatus(msg, isError) {
  // Reuse the main status bar
  if (window.showStatus) { window.showStatus(msg, isError); return; }
  const bar = document.getElementById('statusBar');
  if (!bar) return;
  bar.style.display = 'block';
  bar.style.background = isError ? '#c0392b' : 'var(--teal)';
  bar.textContent = msg;
  setTimeout(() => { bar.style.display = 'none'; }, 5000);
}

function updateSurveyResponseInputMode() {
  const code = (document.getElementById('sv_question_code')?.value || '').trim();
  const q = questionForCode(code);
  const type = surveyQuestionType(q);
  const scoreWrap = document.getElementById('sv_score_group');
  const textWrap = document.getElementById('sv_answer_text_group');
  const isTextBased = type === SURVEY_TYPES.OPEN || type === SURVEY_TYPES.MULTIPLE_CHOICE || type === SURVEY_TYPES.CHECKBOX;
  if (scoreWrap) scoreWrap.style.display = isTextBased ? 'none' : '';
  if (textWrap) textWrap.style.display = isTextBased ? '' : 'none';
}

// ─────────────────────────────────────────────────────────────────────────────
// RENDER — CLIENT (read-only analytics)
// ─────────────────────────────────────────────────────────────────────────────

function renderSurveyClient() {
  const el = document.getElementById('surveyClientContent');
  if (!el) return;

  if (!surveyResponses.length) {
    el.innerHTML = `<div class="no-data-msg">No survey data yet. Import responses in Analyst View.</div>`;
    return;
  }

  const qs    = questionStats();
  const openQs = openEndedQuestionSummaries();
  const mcQs  = multipleChoiceQuestionSummaries();
  const dStats = domainStats();
  const gStats = groupStats('respondent_group');
  const totalRespondents = new Set(surveyResponses.map(r => r.respondent_id).filter(Boolean)).size;
  const ratedRows = likertResponses();
  const openRows = openEndedResponses();

  // ── Summary cards ──
  const typeSort = document.getElementById('surveyQuestionTypeSort')?.value || 'all';
  const showRatings = typeSort === 'all' || typeSort === SURVEY_TYPES.LIKERT;
  const showMultipleChoice = typeSort === 'all' || typeSort === SURVEY_TYPES.MULTIPLE_CHOICE;
  const showOpen = typeSort === 'all' || typeSort === SURVEY_TYPES.OPEN;
  const overallScores = ratedRows.map(r => r.score);
  const overall = calcStats(overallScores);
  const overallMean = overall ? parseFloat(overall.mean) : null;
  const sentiment = overallMean >= 4.0 ? 'Positive' : overallMean >= 3.0 ? 'Neutral / Mixed' : 'Negative';
  const sentColor = overall ? (overallMean >= 4.0 ? '#27ae60' : overallMean >= 3.0 ? '#2980b9' : '#c0392b') : 'var(--ink)';

  let html = `
  <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:1rem;margin-bottom:2rem;">
    ${[
      ['Responses', typeSort === SURVEY_TYPES.LIKERT ? ratedRows.length : typeSort === SURVEY_TYPES.MULTIPLE_CHOICE ? multipleChoiceResponses().length : typeSort === SURVEY_TYPES.OPEN ? openRows.length : surveyResponses.length, ''],
      ...(showRatings ? [['Rating Scores', ratedRows.length, '']] : []),
      ...(showMultipleChoice ? [['Choice Responses', multipleChoiceResponses().length, '']] : []),
      ...(showOpen ? [['Open Insights', openRows.length, '']] : []),
      ['Respondents', totalRespondents || '—', ''],
      ['Questions', surveyQuestions.length, ''],
      ...(showRatings ? [
        ['Overall Mean', overall ? overall.mean : '-', overall ? '/ 5.00' : ''],
        ['Sentiment', overall ? sentiment : '-', '']
      ] : [])
    ].map(([label, val, sub]) => `
    <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1.1rem 1.25rem;">
      <div style="font-family:'IBM Plex Mono',monospace;font-size:0.65rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:0.35rem;">${label}</div>
      <div style="font-family:'Playfair Display',serif;font-size:1.55rem;font-weight:700;color:${label==='Sentiment'?sentColor:'var(--ink)'};">${val}</div>
      ${sub ? `<div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;color:var(--muted);">${sub}</div>` : ''}
    </div>`).join('')}
  </div>`;

  // ── Overall Likert distribution ──
  if (showRatings && overall) {
  html += `
  <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1.5rem;margin-bottom:2rem;">
    <div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:0.75rem;">Overall Response Distribution</div>
    <div style="margin-bottom:0.5rem;">${likertBar(overall.freq, overall.n)}</div>
    <div style="display:flex;gap:1.25rem;flex-wrap:wrap;font-family:'IBM Plex Mono',monospace;font-size:0.72rem;color:var(--muted);">
      ${[[1,'#c0392b','SD'],[2,'#e67e22','D'],[3,'#7f8c8d','N'],[4,'#2980b9','A'],[5,'#27ae60','SA']].map(([v,c,l])=>`
      <span><span style="display:inline-block;width:10px;height:10px;background:${c};border-radius:2px;margin-right:4px;vertical-align:middle;"></span>${l}: ${overall.freq[v]||0} (${((overall.freq[v]||0)/overall.n*100).toFixed(0)}%)</span>`).join('')}
    </div>
    <div style="margin-top:0.65rem;font-family:'IBM Plex Mono',monospace;font-size:0.72rem;color:var(--muted);">
      Agree/Strongly Agree: <strong>${overall.agree_pct}%</strong> &nbsp;|&nbsp; Disagree/Strongly Disagree: <strong>${overall.disagree_pct}%</strong> &nbsp;|&nbsp; SD: ${overall.sd}
    </div>
  </div>`;
  }

  // ── Legend ──
  if (showRatings && overall) {
  html += `
  <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1rem 1.5rem;margin-bottom:2rem;">
    <div style="font-family:'IBM Plex Mono',monospace;font-size:0.65rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:0.75rem;">Likert Scale Legend</div>
    <div style="display:flex;flex-wrap:wrap;gap:0.6rem;">
      ${[
        [1,'#c0392b','Strongly Disagree'],
        [2,'#e67e22','Disagree'],
        [3,'#7f8c8d','Neutral'],
        [4,'#2980b9','Agree'],
        [5,'#27ae60','Strongly Agree']
      ].map(([v,c,l]) => `
        <div style="display:flex;align-items:center;gap:0.5rem;background:${c}11;border:1px solid ${c}44;border-radius:8px;padding:0.4rem 0.85rem;min-width:fit-content;">
          <span style="display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:50%;background:${c};color:#fff;font-family:'IBM Plex Mono',monospace;font-size:0.72rem;font-weight:700;flex-shrink:0;">${v}</span>
          <span style="font-family:'IBM Plex Mono',monospace;font-size:0.78rem;color:#3a3530;font-weight:500;">${l}</span>
        </div>`).join('')}
    </div>
    <div style="margin-top:0.75rem;font-family:'IBM Plex Mono',monospace;font-size:0.7rem;color:var(--muted);line-height:1.5;">
      <span style="color:#27ae60;font-weight:600;">Agree %</span> = scores 4 &amp; 5 &nbsp;·&nbsp;
      <span style="color:#c0392b;font-weight:600;">Disagree %</span> = scores 1 &amp; 2 &nbsp;·&nbsp;
      <span style="font-weight:500;">Mean dot color</span>: <span style="color:#27ae60;">●</span> ≥4.0 &nbsp;<span style="color:#2980b9;">●</span> ≥3.0 &nbsp;<span style="color:#c0392b;">●</span> &lt;2.5
    </div>
  </div>`;
  }

  // ── By CIPQ Domain comparison (only if questions are mapped) ──
  if (showRatings && dStats.length) {
    html += `
    <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1.5rem;margin-bottom:2rem;">
      <div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:1rem;">Perception by CIPQ Domain <span style="font-size:0.62rem;opacity:0.6;font-style:italic;">(comparative only — not fed into CIPQ analytics)</span></div>
      <div style="overflow-x:auto;">
      <table style="width:100%;border-collapse:collapse;font-family:'IBM Plex Mono',monospace;font-size:0.78rem;">
        <thead><tr style="border-bottom:2px solid var(--border);">
          <th style="text-align:left;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">Domain</th>
          <th style="text-align:center;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">N</th>
          <th style="text-align:center;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">Mean</th>
          <th style="text-align:center;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">Median</th>
          <th style="text-align:center;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">SD</th>
          <th style="text-align:center;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">Agree %</th>
          <th style="padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">Distribution</th>
        </tr></thead>
        <tbody>
          ${dStats.map(d => `
          <tr style="border-bottom:1px solid var(--border);">
            <td style="padding:0.65rem 0.75rem;">${domainPill(d.domain)}</td>
            <td style="text-align:center;padding:0.65rem 0.75rem;">${d.stats.n}</td>
            <td style="text-align:center;padding:0.65rem 0.75rem;">${meanDot(d.stats.mean)}${d.stats.mean}</td>
            <td style="text-align:center;padding:0.65rem 0.75rem;">${d.stats.median}</td>
            <td style="text-align:center;padding:0.65rem 0.75rem;">${d.stats.sd}</td>
            <td style="text-align:center;padding:0.65rem 0.75rem;color:#27ae60;font-weight:500;">${d.stats.agree_pct}%</td>
            <td style="padding:0.65rem 0.75rem;min-width:140px;">${likertBar(d.stats.freq, d.stats.n)}</td>
          </tr>`).join('')}
        </tbody>
      </table>
      </div>
    </div>`;
  }

  // ── By Respondent Group ──
  if (showRatings && gStats.length) {
    html += `
    <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1.5rem;margin-bottom:2rem;">
      <div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:1rem;">Perception by Respondent Group</div>
      <div style="overflow-x:auto;">
      <table style="width:100%;border-collapse:collapse;font-family:'IBM Plex Mono',monospace;font-size:0.78rem;">
        <thead><tr style="border-bottom:2px solid var(--border);">
          <th style="text-align:left;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">Group</th>
          <th style="text-align:center;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">Respondents</th>
          <th style="text-align:center;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">Responses</th>
          <th style="text-align:center;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">Mean</th>
          <th style="text-align:center;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">SD</th>
          <th style="text-align:center;padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">Agree %</th>
          <th style="padding:0.5rem 0.75rem;color:var(--muted);font-weight:500;">Distribution</th>
        </tr></thead>
        <tbody>
          ${gStats.map(g => `
          <tr style="border-bottom:1px solid var(--border);">
            <td style="padding:0.65rem 0.75rem;font-weight:500;">${escHtml(g.group)}</td>
            <td style="text-align:center;padding:0.65rem 0.75rem;">${g.respondents || '—'}</td>
            <td style="text-align:center;padding:0.65rem 0.75rem;">${g.stats.n}</td>
            <td style="text-align:center;padding:0.65rem 0.75rem;">${meanDot(g.stats.mean)}${g.stats.mean}</td>
            <td style="text-align:center;padding:0.65rem 0.75rem;">${g.stats.sd}</td>
            <td style="text-align:center;padding:0.65rem 0.75rem;color:#27ae60;font-weight:500;">${g.stats.agree_pct}%</td>
            <td style="padding:0.65rem 0.75rem;min-width:140px;">${likertBar(g.stats.freq, g.stats.n)}</td>
          </tr>`).join('')}
        </tbody>
      </table>
      </div>
    </div>`;
  }

  // ── Per-question table ──
  if (showRatings && qs.length) {
    // Group by category
    const categories = [...new Set(qs.map(q => q.category || 'General'))];
    html += `
    <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1.5rem;margin-bottom:2rem;">
      <div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:1rem;">Question-Level Statistics</div>`;

    for (const cat of categories) {
      const catQs = qs.filter(q => (q.category || 'General') === cat);
      html += `
      <div style="font-family:'IBM Plex Mono',monospace;font-size:0.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:0.07em;border-top:1px solid var(--border);padding:0.65rem 0 0.4rem;margin-top:0.5rem;">${escHtml(cat)}</div>
      <div style="overflow-x:auto;">
      <table style="width:100%;border-collapse:collapse;font-family:'IBM Plex Mono',monospace;font-size:0.75rem;margin-bottom:1rem;">
        <thead><tr style="border-bottom:1px solid var(--border);">
          <th style="text-align:left;padding:0.45rem 0.6rem;color:var(--muted);font-weight:500;min-width:55px;">Code</th>
          <th style="text-align:left;padding:0.45rem 0.6rem;color:var(--muted);font-weight:500;min-width:240px;">Question</th>
          <th style="text-align:left;padding:0.45rem 0.6rem;color:var(--muted);font-weight:500;">Domain</th>
          <th style="text-align:center;padding:0.45rem 0.6rem;color:var(--muted);font-weight:500;">N</th>
          <th style="text-align:center;padding:0.45rem 0.6rem;color:var(--muted);font-weight:500;">Mean</th>
          <th style="text-align:center;padding:0.45rem 0.6rem;color:var(--muted);font-weight:500;">Mdn</th>
          <th style="text-align:center;padding:0.45rem 0.6rem;color:var(--muted);font-weight:500;">SD</th>
          <th style="text-align:center;padding:0.45rem 0.6rem;color:var(--muted);font-weight:500;">Agree%</th>
          <th style="padding:0.45rem 0.6rem;color:var(--muted);font-weight:500;">Distribution</th>
        </tr></thead>
        <tbody>
          ${catQs.map(q => `
          <tr style="border-bottom:1px solid #f0ece4;">
            <td style="padding:0.5rem 0.6rem;font-weight:500;white-space:nowrap;">${escHtml(q.code)}</td>
            <td style="padding:0.5rem 0.6rem;line-height:1.4;">${escHtml(q.text)}</td>
            <td style="padding:0.5rem 0.6rem;">${domainPill(q.cipq_domain)}</td>
            <td style="text-align:center;padding:0.5rem 0.6rem;">${q.stats.n}</td>
            <td style="text-align:center;padding:0.5rem 0.6rem;">${meanDot(q.stats.mean)}${q.stats.mean}</td>
            <td style="text-align:center;padding:0.5rem 0.6rem;">${q.stats.median}</td>
            <td style="text-align:center;padding:0.5rem 0.6rem;">${q.stats.sd}</td>
            <td style="text-align:center;padding:0.5rem 0.6rem;color:#27ae60;font-weight:500;">${q.stats.agree_pct}%</td>
            <td style="padding:0.5rem 0.6rem;min-width:120px;">${likertBar(q.stats.freq, q.stats.n)}</td>
          </tr>`).join('')}
        </tbody>
      </table>
      </div>`;
    }
    html += `</div>`;
  }

  // ── Multiple Choice / Checkbox summary ──
  if (showMultipleChoice && mcQs.length) {
    html += `
    <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1.5rem;margin-bottom:2rem;">
      <div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:1rem;">Multiple Choice Responses</div>
      ${mcQs.map(q => `
        <div style="border-top:1px solid var(--border);padding-top:1rem;margin-top:1rem;">
          <div style="display:flex;align-items:center;gap:0.6rem;margin-bottom:0.3rem;flex-wrap:wrap;">
            <span style="font-family:'IBM Plex Mono',monospace;font-size:0.72rem;color:var(--muted);">${escHtml(q.code)}</span>
            <span style="font-family:'IBM Plex Mono',monospace;font-size:0.65rem;padding:0.1rem 0.5rem;border-radius:999px;background:#f0ece4;color:var(--muted);border:1px solid var(--border);">${surveyQuestionType(q) === SURVEY_TYPES.CHECKBOX ? 'checkbox' : 'multiple choice'}</span>
            ${domainPill(q.cipq_domain)}
            <span style="font-family:'IBM Plex Mono',monospace;font-size:0.68rem;color:var(--muted);margin-left:auto;">${q.n} response${q.n !== 1 ? 's' : ''}</span>
          </div>
          <div style="font-size:0.95rem;font-weight:600;margin-bottom:0.85rem;line-height:1.4;">${escHtml(q.text)}</div>
          <div style="display:grid;gap:0.45rem;">
            ${q.choices.map((choice, i) => `
              <div style="display:grid;grid-template-columns:minmax(140px,30%) 1fr auto;align-items:center;gap:0.75rem;">
                <div style="font-family:'IBM Plex Mono',monospace;font-size:0.78rem;color:#3a3530;font-weight:${i === 0 ? '600' : '400'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${escHtml(choice.label)}">${escHtml(choice.label)}</div>
                <div style="background:#f0ece4;border-radius:4px;overflow:hidden;height:14px;">
                  <div style="width:${choice.pct}%;height:100%;background:${i === 0 ? '#2a6b6e' : '#7f8c8d'};border-radius:4px;transition:width 0.3s;"></div>
                </div>
                <div style="font-family:'IBM Plex Mono',monospace;font-size:0.72rem;color:var(--muted);white-space:nowrap;text-align:right;">${choice.count} <span style="opacity:0.65;">(${choice.pct}%)</span></div>
              </div>`).join('')}
          </div>
        </div>`).join('')}
    </div>`;
  }

  if (showOpen && openQs.length) {
    html += `
    <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1.5rem;margin-bottom:2rem;">
      <div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:1rem;">Open-Ended Participant Insights</div>
      ${openQs.map(q => `
        <div style="border-top:1px solid var(--border);padding-top:0.85rem;margin-top:0.85rem;">
          <div style="font-family:'IBM Plex Mono',monospace;font-size:0.72rem;color:var(--muted);margin-bottom:0.35rem;">${escHtml(q.code)} - ${q.responses.length} insight${q.responses.length !== 1 ? 's' : ''}</div>
          <div style="font-size:0.95rem;font-weight:600;margin-bottom:0.75rem;line-height:1.4;">${escHtml(q.text)}</div>
          <div style="display:grid;gap:0.75rem;">
            ${q.responses.map(r => `
              <div style="border:1px solid #f0ece4;border-radius:8px;padding:0.8rem 0.9rem;background:#fdfbf7;">
                <div style="font-family:'IBM Plex Mono',monospace;font-size:0.68rem;color:var(--muted);margin-bottom:0.35rem;">${escHtml([r.respondent_id, r.respondent_group, r.region].filter(Boolean).join(' / ') || 'Unspecified respondent')}</div>
                <div style="font-size:0.9rem;line-height:1.5;">${escHtml(r.answer_text)}</div>
              </div>
            `).join('')}
          </div>
        </div>
      `).join('')}
    </div>`;
  }

  if (typeSort === SURVEY_TYPES.LIKERT && !overall) {
    html += `<div class="no-data-msg">No rating survey responses yet.</div>`;
  }

  if (typeSort === SURVEY_TYPES.OPEN && !openQs.length) {
    html += `<div class="no-data-msg">No open-ended survey responses yet.</div>`;
  }

  if (typeSort === SURVEY_TYPES.MULTIPLE_CHOICE && !mcQs.length) {
    html += `<div class="no-data-msg">No multiple choice or checkbox survey responses yet.</div>`;
  }

  el.innerHTML = html;
}

// ─────────────────────────────────────────────────────────────────────────────
// RENDER — ANALYST (data entry + management)
// ─────────────────────────────────────────────────────────────────────────────

function renderSurveyAnalyst() {
  const el = document.getElementById('surveyAnalystContent');
  if (!el) return;
  const canEdit = window.canWrite && window.canWrite();

  // Refresh the question dropdown in the static response entry form
  const qSelect = document.getElementById('sv_question_code');
  if (qSelect) {
    const current = qSelect.value;
    qSelect.innerHTML = surveyQuestions.length
      ? `<option value="">— Select question —</option>` +
        surveyQuestions.map(q => `<option value="${escHtml(q.code)}">${escHtml(q.code)} — ${escHtml(q.text.length > 55 ? q.text.substring(0,55)+'…' : q.text)}</option>`).join('')
      : `<option value="">— Add questions first —</option>`;
    if (current) qSelect.value = current;
    updateSurveyResponseInputMode();
  }

  const qRows = surveyQuestions.length
    ? surveyQuestions.map(q => {
        const n = surveyResponses.filter(r => r.question_code === q.code).length;
        return `
        <tr style="border-bottom:1px solid var(--border);">
          <td style="padding:0.6rem 0.85rem;font-family:'IBM Plex Mono',monospace;font-size:0.8rem;font-weight:500;white-space:nowrap;">${escHtml(q.code)}</td>
          <td style="padding:0.6rem 0.85rem;font-size:0.88rem;line-height:1.45;">${escHtml(q.text)}</td>
          <td style="padding:0.6rem 0.85rem;font-family:'IBM Plex Mono',monospace;font-size:0.76rem;color:var(--muted);">${isLikertQuestion(q) ? 'Rate 1–5' : isMultipleChoiceQuestion(q) ? (surveyQuestionType(q) === SURVEY_TYPES.CHECKBOX ? 'Checkbox' : 'Multiple choice') : 'Open ended'}</td>
          <td style="padding:0.6rem 0.85rem;">${domainPill(q.cipq_domain)}</td>
          <td style="padding:0.6rem 0.85rem;font-family:'IBM Plex Mono',monospace;font-size:0.76rem;color:var(--muted);">${escHtml(q.category || '—')}</td>
          <td style="padding:0.6rem 0.85rem;text-align:center;font-family:'IBM Plex Mono',monospace;font-size:0.8rem;">${n}</td>
          ${canEdit ? `<td style="padding:0.6rem 0.85rem;"><button class="btn btn-secondary" type="button" style="padding:0.35rem 0.8rem;min-height:unset;font-size:0.7rem;" onclick="deleteSurveyQuestion('${escHtml(q.code)}')">Delete</button></td>` : '<td></td>'}
        </tr>`;
      }).join('')
    : `<tr><td colspan="6" style="padding:2rem;text-align:center;color:var(--muted);font-family:'IBM Plex Mono',monospace;font-size:0.8rem;">No questions yet — add one above or import a questions CSV.</td></tr>`;

  el.innerHTML = `
  <div class="summary-shell">
    <div class="section-title" style="font-size:1.1rem;margin-top:0;border-bottom:1px solid var(--border);">
      Survey Questions
      <span>${surveyQuestions.length} question${surveyQuestions.length !== 1 ? 's' : ''} · ${surveyResponses.length} response${surveyResponses.length !== 1 ? 's' : ''}</span>
    </div>
    <div style="overflow-x:auto;">
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="border-bottom:2px solid var(--border);">
            <th style="text-align:left;padding:0.5rem 0.85rem;font-family:'IBM Plex Mono',monospace;font-size:0.68rem;text-transform:uppercase;letter-spacing:0.07em;color:var(--muted);font-weight:500;">Code</th>
            <th style="text-align:left;padding:0.5rem 0.85rem;font-family:'IBM Plex Mono',monospace;font-size:0.68rem;text-transform:uppercase;letter-spacing:0.07em;color:var(--muted);font-weight:500;">Question</th>
            <th style="text-align:left;padding:0.5rem 0.85rem;font-family:'IBM Plex Mono',monospace;font-size:0.68rem;text-transform:uppercase;letter-spacing:0.07em;color:var(--muted);font-weight:500;">Type</th>
            <th style="text-align:left;padding:0.5rem 0.85rem;font-family:'IBM Plex Mono',monospace;font-size:0.68rem;text-transform:uppercase;letter-spacing:0.07em;color:var(--muted);font-weight:500;">Domain</th>
            <th style="text-align:left;padding:0.5rem 0.85rem;font-family:'IBM Plex Mono',monospace;font-size:0.68rem;text-transform:uppercase;letter-spacing:0.07em;color:var(--muted);font-weight:500;">Category</th>
            <th style="text-align:center;padding:0.5rem 0.85rem;font-family:'IBM Plex Mono',monospace;font-size:0.68rem;text-transform:uppercase;letter-spacing:0.07em;color:var(--muted);font-weight:500;">Responses</th>
            <th></th>
          </tr>
        </thead>
        <tbody>${qRows}</tbody>
      </table>
    </div>
  </div>`;
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN RENDER DISPATCHER
// ─────────────────────────────────────────────────────────────────────────────

function renderSurveyTab() {
  renderSurveyClient();
  renderSurveyAnalyst();
}

// ─────────────────────────────────────────────────────────────────────────────
// GLOBAL EXPORTS
// ─────────────────────────────────────────────────────────────────────────────

Object.assign(window, {
  loadSurveyData,
  fetchAllSurveyResponses,
  renderSurveyTab,
  renderSurveyClient,
  renderSurveyAnalyst,
  handleSurveyImport,
  handleSurveyQuestionsImport,
  handleSurveyResponsesImport,
  updateSurveyResponseInputMode,
  addSurveyResponse,
  addSurveyQuestion,
  deleteSurveyQuestion,
  clearAllSurveyData,
  exportSurveyCSV,
  exportSurveySummaryCSV,
  exportSurveyGroupSummaryCSV,
  downloadSurveyReport,
  copySurveyReport,
  downloadSurveyTemplate,
  downloadQuestionsTemplate,
  downloadResponsesTemplate
});
