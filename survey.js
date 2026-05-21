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

// ── In-memory store ──────────────────────────────────────────────────────────
let surveyQuestions  = [];   // { id, code, text, cipq_domain, category }
let surveyResponses  = [];   // { id, respondent_id, respondent_group, region, question_code, score, source_id, recorded_at }

// ── Supabase helpers (shared client from app.js) ─────────────────────────────
function getSB() { return window.supabaseClient || null; }
function getSBUser() { return window.currentUser || null; }

// ─────────────────────────────────────────────────────────────────────────────
// DATA LOADING
// ─────────────────────────────────────────────────────────────────────────────

async function loadSurveyData() {
  const sb = getSB();
  surveyShowStatus('Loading survey data…', false);
  if (!sb) {
    renderSurveyTab();
    surveyShowStatus('Supabase not configured — using local data only.', true);
    return;
  }

  try {
    const [qRes, rRes] = await Promise.all([
      sb.from(SURVEY_Q_TABLE).select('*').order('code'),
      getSBUser()
        ? sb.from(SURVEY_TABLE).select('*').order('recorded_at')
        : sb.from(SURVEY_TABLE).select('*').order('recorded_at')
    ]);
    if (qRes.error) throw qRes.error;
    if (rRes.error) throw rRes.error;
    surveyQuestions = qRes.data || [];
    surveyResponses = rRes.data || [];
    renderSurveyTab();
    surveyShowStatus(`Survey loaded: ${surveyQuestions.length} questions, ${surveyResponses.length} responses.`, false);
  } catch (err) {
    surveyShowStatus(`Survey load error: ${err.message || err}`, true);
    renderSurveyTab();
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// CSV IMPORT
// Format: respondent_id, respondent_group, region, source_id, [Q_CODE]…
// Header row must contain the question codes as column names.
// ─────────────────────────────────────────────────────────────────────────────

function handleSurveyImport(event) {
  if (!window.canWrite || !window.canWrite()) {
    surveyShowStatus('Sign in to import survey data.', true);
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
      if (!results.data.length) {
        surveyShowStatus('CSV appears empty.', true);
        return;
      }

      const meta = ['respondent_id','respondent_group','region','source_id'];
      const qCodes = Object.keys(results.data[0]).filter(k => !meta.includes(k.toLowerCase()));

      if (!qCodes.length) {
        surveyShowStatus('No question columns found. Ensure column headers match question codes (e.g. SQ_01).', true);
        return;
      }

      // Auto-create any unknown questions as placeholder entries
      const knownCodes = new Set(surveyQuestions.map(q => q.code));
      const newQs = qCodes.filter(c => !knownCodes.has(c)).map(c => ({
        code: c,
        text: c,
        cipq_domain: null,
        category: 'Imported'
      }));
      if (newQs.length) await saveSurveyQuestions(newQs);

      const rows = [];
      for (const row of results.data) {
        const respondent_id    = row.respondent_id || row.Respondent_ID || '';
        const respondent_group = row.respondent_group || row.Respondent_Group || '';
        const region           = row.region || row.Region || '';
        const source_id        = row.source_id || row.Source_ID || '';
        for (const qCode of qCodes) {
          const raw = row[qCode];
          const score = parseInt(raw, 10);
          if (!isNaN(score) && score >= 1 && score <= 5) {
            rows.push({ respondent_id, respondent_group, region, source_id, question_code: qCode, score, recorded_at: new Date().toISOString() });
          }
        }
      }

      if (!rows.length) {
        surveyShowStatus('No valid Likert scores (1–5) found in CSV.', true);
        return;
      }

      await saveSurveyResponses(rows);
    },
    error: (err) => surveyShowStatus(`CSV parse error: ${err.message}`, true)
  });
}

async function saveSurveyQuestions(questions) {
  const sb = getSB();
  if (!sb) { surveyQuestions.push(...questions); return; }
  const { data, error } = await sb.from(SURVEY_Q_TABLE).upsert(questions, { onConflict: 'code' }).select();
  if (error) { surveyShowStatus(`Could not save questions: ${error.message}`, true); return; }
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
    surveyShowStatus(`Could not save responses: ${error.message}`, true);
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
  const scoreRaw         = parseInt(document.getElementById('sv_score')?.value || '', 10);

  if (!question_code) { surveyShowStatus('Select or enter a question code.', true); return; }
  if (isNaN(scoreRaw) || scoreRaw < 1 || scoreRaw > 5) { surveyShowStatus('Score must be 1–5.', true); return; }

  // Ensure question exists
  if (!surveyQuestions.find(q => q.code === question_code)) {
    await saveSurveyQuestions([{ code: question_code, text: question_code, cipq_domain: null, category: 'Manual' }]);
  }

  await saveSurveyResponses([{
    respondent_id, respondent_group, region, source_id,
    question_code, score: scoreRaw,
    recorded_at: new Date().toISOString()
  }]);

  // Clear entry fields
  ['sv_respondent_id','sv_score'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
}

async function addSurveyQuestion() {
  if (!window.canWrite || !window.canWrite()) {
    surveyShowStatus('Sign in to add survey questions.', true);
    return;
  }
  const code       = (document.getElementById('sq_code')?.value || '').trim().toUpperCase();
  const text       = (document.getElementById('sq_text')?.value || '').trim();
  const cipq_domain = (document.getElementById('sq_domain')?.value || '') || null;
  const category   = (document.getElementById('sq_category')?.value || '').trim() || 'General';

  if (!code || !text) { surveyShowStatus('Question code and text are required.', true); return; }
  if (surveyQuestions.find(q => q.code === code)) { surveyShowStatus(`Question ${code} already exists.`, true); return; }

  await saveSurveyQuestions([{ code, text, cipq_domain, category }]);
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
  return surveyQuestions.map(q => {
    const scores = surveyResponses.filter(r => r.question_code === q.code).map(r => r.score);
    return { ...q, stats: calcStats(scores) };
  }).filter(q => q.stats);
}

function domainStats() {
  return CIPQ_DOMAINS.map(domain => {
    const qCodes = new Set(surveyQuestions.filter(q => q.cipq_domain === domain).map(q => q.code));
    const scores = surveyResponses.filter(r => qCodes.has(r.question_code)).map(r => r.score);
    return { domain, stats: calcStats(scores), qCount: qCodes.size };
  }).filter(d => d.stats);
}

function groupStats(groupField = 'respondent_group') {
  const groups = [...new Set(surveyResponses.map(r => r[groupField]).filter(Boolean))];
  return groups.map(group => {
    const groupResps = surveyResponses.filter(r => r[groupField] === group);
    const scores = groupResps.map(r => r.score);
    const respondents = new Set(groupResps.map(r => r.respondent_id).filter(Boolean)).size;
    return { group, respondents, stats: calcStats(scores) };
  }).filter(g => g.stats).sort((a,b) => parseFloat(b.stats.mean) - parseFloat(a.stats.mean));
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
    question_text:    surveyQuestions.find(q => q.code === r.question_code)?.text || '',
    cipq_domain:      surveyQuestions.find(q => q.code === r.question_code)?.cipq_domain || '',
    score:            r.score,
    score_label:      LIKERT_LABELS[r.score] || '',
    recorded_at:      r.recorded_at || ''
  }));
  const csv = Papa.unparse(rows);
  downloadText(csv, 'survey_responses.csv', 'text/csv');
}

function exportSurveySummaryCSV() {
  const qs = questionStats();
  if (!qs.length) { surveyShowStatus('No survey data to export.', true); return; }
  const rows = qs.map(q => ({
    code: q.code,
    text: q.text,
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
    freq_4: q.stats.freq[4], freq_5: q.stats.freq[5]
  }));
  const csv = Papa.unparse(rows);
  downloadText(csv, 'survey_summary.csv', 'text/csv');
}

function downloadSurveyTemplate() {
  const qCodes = surveyQuestions.length
    ? surveyQuestions.map(q => q.code)
    : ['SQ_01','SQ_02','SQ_03'];
  const header = ['respondent_id','respondent_group','region','source_id', ...qCodes];
  const example = ['R001','Publisher','NCR','SURVEY_01', ...qCodes.map(() => '4')];
  const csv = Papa.unparse([Object.fromEntries(header.map((h,i) => [h, example[i]]))]);
  downloadText(csv, 'survey_import_template.csv', 'text/csv');
}

function downloadText(content, filename, mime) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([content], { type: mime }));
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
  const dStats = domainStats();
  const gStats = groupStats('respondent_group');
  const totalRespondents = new Set(surveyResponses.map(r => r.respondent_id).filter(Boolean)).size;

  // ── Summary cards ──
  const overallScores = surveyResponses.map(r => r.score);
  const overall = calcStats(overallScores);
  const overallMean = parseFloat(overall.mean);
  const sentiment = overallMean >= 4.0 ? 'Positive' : overallMean >= 3.0 ? 'Neutral / Mixed' : 'Negative';
  const sentColor = overallMean >= 4.0 ? '#27ae60' : overallMean >= 3.0 ? '#2980b9' : '#c0392b';

  let html = `
  <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:1rem;margin-bottom:2rem;">
    ${[
      ['Responses', surveyResponses.length, ''],
      ['Respondents', totalRespondents || '—', ''],
      ['Questions', surveyQuestions.length, ''],
      ['Overall Mean', overall.mean, '/ 5.00'],
      ['Sentiment', sentiment, '']
    ].map(([label, val, sub]) => `
    <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1.1rem 1.25rem;">
      <div style="font-family:'IBM Plex Mono',monospace;font-size:0.65rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:0.35rem;">${label}</div>
      <div style="font-family:'Playfair Display',serif;font-size:1.55rem;font-weight:700;color:${label==='Sentiment'?sentColor:'var(--ink)'};">${val}</div>
      ${sub ? `<div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;color:var(--muted);">${sub}</div>` : ''}
    </div>`).join('')}
  </div>`;

  // ── Overall Likert distribution ──
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

  // ── By CIPQ Domain comparison (only if questions are mapped) ──
  if (dStats.length) {
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
  if (gStats.length) {
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
  if (qs.length) {
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

  // ── Legend ──
  html += `
  <div style="font-family:'IBM Plex Mono',monospace;font-size:0.68rem;color:var(--muted);display:flex;flex-wrap:wrap;gap:1rem;margin-top:0.5rem;">
    ${[[1,'#c0392b','1 — Strongly Disagree'],[2,'#e67e22','2 — Disagree'],[3,'#7f8c8d','3 — Neutral'],[4,'#2980b9','4 — Agree'],[5,'#27ae60','5 — Strongly Agree']].map(([,c,l])=>`<span><span style="display:inline-block;width:10px;height:10px;background:${c};border-radius:2px;margin-right:4px;vertical-align:middle;"></span>${l}</span>`).join('')}
  </div>`;

  el.innerHTML = html;
}

// ─────────────────────────────────────────────────────────────────────────────
// RENDER — ANALYST (data entry + management)
// ─────────────────────────────────────────────────────────────────────────────

function renderSurveyAnalyst() {
  const el = document.getElementById('surveyAnalystContent');
  if (!el) return;
  const canEdit = window.canWrite && window.canWrite();

  // Question list
  const qRows = surveyQuestions.length
    ? surveyQuestions.map(q => {
        const n = surveyResponses.filter(r => r.question_code === q.code).length;
        return `
        <tr style="border-bottom:1px solid var(--border);">
          <td style="padding:0.55rem 0.75rem;font-family:'IBM Plex Mono',monospace;font-size:0.78rem;font-weight:500;">${escHtml(q.code)}</td>
          <td style="padding:0.55rem 0.75rem;font-size:0.85rem;">${escHtml(q.text)}</td>
          <td style="padding:0.55rem 0.75rem;">${domainPill(q.cipq_domain)}</td>
          <td style="padding:0.55rem 0.75rem;font-family:'IBM Plex Mono',monospace;font-size:0.75rem;color:var(--muted);">${escHtml(q.category || '')}</td>
          <td style="text-align:center;padding:0.55rem 0.75rem;font-family:'IBM Plex Mono',monospace;font-size:0.78rem;">${n}</td>
          ${canEdit ? `<td style="padding:0.55rem 0.75rem;"><button class="btn btn-secondary" type="button" style="padding:0.3rem 0.7rem;min-height:unset;font-size:0.68rem;" onclick="deleteSurveyQuestion('${escHtml(q.code)}')">Delete</button></td>` : '<td></td>'}
        </tr>`;
      }).join('')
    : `<tr><td colspan="6" style="padding:1.5rem;text-align:center;color:var(--muted);font-family:'IBM Plex Mono',monospace;font-size:0.8rem;">No questions defined yet.</td></tr>`;

  el.innerHTML = `
  ${canEdit ? `
  <!-- Add Question -->
  <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1.5rem;margin-bottom:1.5rem;">
    <div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:1rem;">Add Survey Question</div>
    <div class="form-grid" style="grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:1rem;">
      <div class="form-group">
        <label>Question Code</label>
        <input type="text" id="sq_code" placeholder="e.g. SQ_01" style="text-transform:uppercase;">
      </div>
      <div class="form-group" style="grid-column:span 2;">
        <label>Question Text</label>
        <input type="text" id="sq_text" placeholder="e.g. Printing costs significantly affect publishing sustainability.">
      </div>
      <div class="form-group">
        <label>Related CIPQ Domain <span style="font-size:0.65rem;color:var(--muted);">(comparison only)</span></label>
        <select id="sq_domain">
          <option value="">-- None --</option>
          ${CIPQ_DOMAINS.map(d=>`<option>${d}</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label>Category</label>
        <input type="text" id="sq_category" placeholder="e.g. Production Costs">
      </div>
    </div>
    <div class="actions-row" style="margin-top:1rem;">
      <button class="btn btn-primary" type="button" onclick="addSurveyQuestion()">Add Question</button>
    </div>
  </div>

  <!-- Single Response Entry -->
  <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1.5rem;margin-bottom:1.5rem;">
    <div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:1rem;">Add Individual Response</div>
    <div class="form-grid" style="grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:1rem;">
      <div class="form-group">
        <label>Respondent ID</label>
        <input type="text" id="sv_respondent_id" placeholder="e.g. R001">
      </div>
      <div class="form-group">
        <label>Respondent Group</label>
        <select id="sv_respondent_group">
          <option value="">-- Select --</option>
          <option>Author</option><option>Publisher</option><option>Printer</option>
          <option>Distributor</option><option>Bookseller</option><option>Librarian</option>
          <option>Reader</option><option>Government</option><option>Other</option>
        </select>
      </div>
      <div class="form-group">
        <label>Region</label>
        <select id="sv_region">
          <option value="">-- Optional --</option>
          <option>NCR</option><option>CAR</option><option>Region I</option><option>Region II</option>
          <option>Region III</option><option>Region IV-A</option><option>Region IV-B</option>
          <option>Region V</option><option>Region VI</option><option>Region VII</option>
          <option>Region VIII</option><option>Region IX</option><option>Region X</option>
          <option>Region XI</option><option>Region XII</option><option>Region XIII</option>
          <option>BARMM</option><option>Unknown</option>
        </select>
      </div>
      <div class="form-group">
        <label>Source ID</label>
        <input type="text" id="sv_source_id" placeholder="e.g. SURVEY_01">
      </div>
      <div class="form-group">
        <label>Question Code</label>
        <select id="sv_question_code">
          <option value="">-- Select --</option>
          ${surveyQuestions.map(q=>`<option value="${escHtml(q.code)}">${escHtml(q.code)} — ${escHtml(q.text.substring(0,50))}${q.text.length>50?'…':''}</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label>Score (1–5)</label>
        <select id="sv_score">
          <option value="">-- Select --</option>
          ${[1,2,3,4,5].map(v=>`<option value="${v}">${v} — ${LIKERT_LABELS[v]}</option>`).join('')}
        </select>
      </div>
    </div>
    <div class="actions-row" style="margin-top:1rem;">
      <button class="btn btn-primary" type="button" onclick="addSurveyResponse()">Add Response</button>
    </div>
  </div>` : ''}

  <!-- Question List -->
  <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:1.5rem;margin-bottom:1.5rem;">
    <div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-bottom:1rem;">
      Survey Questions (${surveyQuestions.length})
      <span style="margin-left:1rem;font-size:0.63rem;opacity:0.6;">${surveyResponses.length} total responses</span>
    </div>
    <div style="overflow-x:auto;">
    <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
      <thead><tr style="border-bottom:2px solid var(--border);">
        <th style="text-align:left;padding:0.45rem 0.75rem;font-family:'IBM Plex Mono',monospace;font-size:0.68rem;color:var(--muted);font-weight:500;">Code</th>
        <th style="text-align:left;padding:0.45rem 0.75rem;font-family:'IBM Plex Mono',monospace;font-size:0.68rem;color:var(--muted);font-weight:500;">Text</th>
        <th style="text-align:left;padding:0.45rem 0.75rem;font-family:'IBM Plex Mono',monospace;font-size:0.68rem;color:var(--muted);font-weight:500;">Domain</th>
        <th style="text-align:left;padding:0.45rem 0.75rem;font-family:'IBM Plex Mono',monospace;font-size:0.68rem;color:var(--muted);font-weight:500;">Category</th>
        <th style="text-align:center;padding:0.45rem 0.75rem;font-family:'IBM Plex Mono',monospace;font-size:0.68rem;color:var(--muted);font-weight:500;">Responses</th>
        <th></th>
      </tr></thead>
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
  renderSurveyTab,
  renderSurveyClient,
  renderSurveyAnalyst,
  handleSurveyImport,
  addSurveyResponse,
  addSurveyQuestion,
  deleteSurveyQuestion,
  clearAllSurveyData,
  exportSurveyCSV,
  exportSurveySummaryCSV,
  downloadSurveyTemplate
});