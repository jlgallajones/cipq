let dataset = [];
let currentSeverity = null;
let activeTraceCode = null;
let activeView = 'client';

const SEVERITY_LABELS = {
  1: 'Minimal',
  2: 'Low',
  3: 'Moderate',
  4: 'High',
  5: 'Critical'
};

const DOMAIN_CLASS = { Creation: 'creation', Production: 'production', Distribution: 'distribution', Access: 'access' };
const DOMAIN_PREFIX = { Creation: 'CREATE', Production: 'PROD', Distribution: 'DIST', Access: 'ACCESS' };
const DOMAIN_COLORS = { Creation: '#8b4f9e', Production: '#c94a2e', Distribution: '#2a6b6e', Access: '#3a7a3a', Unspecified: '#7a7065' };
const CHART_COLORS = {
  ink: '#1a1a2e',
  muted: '#7a7065',
  border: '#c8bfae',
  paper: '#fffdf8',
  cream: '#ede8dc',
  rust: '#c94a2e',
  gold: '#c8993a',
  teal: '#2a6b6e'
};
const VALID_DOMAINS = ['Creation', 'Production', 'Distribution', 'Access'];
const DOMAIN_TO_VALUE_CHAIN_STAGE = {
  Creation: 'Development',
  Production: 'Production',
  Distribution: 'Distribution',
  Access: 'Market Access'
};
const SIMULATION_DEPENDENCY_MATRIX = {
  Creation: { Production: 0.45, Distribution: 0.2, Access: 0.2, Creation: 0.15 },
  Production: { Distribution: 0.45, Access: 0.3, Creation: 0.15, Production: 0.1 },
  Distribution: { Access: 0.5, Production: 0.2, Creation: 0.15, Distribution: 0.15 },
  Access: { Creation: 0.3, Distribution: 0.25, Production: 0.2, Access: 0.25 }
};
const DEFAULT_SIMULATION_STATE = {
  relief: { Creation: 0, Production: 0, Distribution: 0, Access: 0 },
  shock: { Creation: 0, Production: 0, Distribution: 0, Access: 0 },
  redistribution: 35,
  balancing: 20,
  rounds: 3
};
const LEGACY_DOMAIN_MAP = { Governance: 'Production' };
const SCORING_CONFIDENCE_OPTIONS = ['low', 'medium', 'high'];
const CIPQ_SOURCE_TYPES = ['FGD', 'KII', 'Document'];
const VALUE_CHAIN_STAGES = ['Development', 'Production', 'Distribution', 'Market Access'];
const VALUE_CHAIN_ALIASES = { Access: 'Market Access', Market: 'Market Access' };
const PESTLE_TAGS = ['Political', 'Economic', 'Social', 'Technological', 'Legal', 'Environmental'];
const expandedSnippetIds = new Set();
let simulationState = JSON.parse(JSON.stringify(DEFAULT_SIMULATION_STATE));

const SUPABASE_URL = 'https://ueyyrugaynzczkcwnxbt.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVleXlydWdheW56Y3prY3dueGJ0Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY5NDEyOTIsImV4cCI6MjA5MjUxNzI5Mn0.x1XqiUs52OH4TKe070OSOK4c0gSGeYRSZJ30XZSBSuQ';
const SUPABASE_TABLE = 'segments';
const SINGLE_LOGIN_USERNAME = 'techfactorsnbdb';
const SINGLE_LOGIN_PASSWORD = 'pbianbdb';
const SINGLE_LOGIN_AUTH_EMAIL = `${SINGLE_LOGIN_USERNAME}@cipq.local`;
const supabaseReady = SUPABASE_URL.startsWith('https://') && !SUPABASE_URL.includes('YOUR_PROJECT_ID') && !!SUPABASE_ANON_KEY && !SUPABASE_ANON_KEY.includes('YOUR_SUPABASE');
const supabaseClient = supabaseReady ? window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true
  }
}) : null;
let currentSession = null;
let currentUser = null;
let isCloudSyncing = false;
const unavailableSupabaseColumns = new Set();

// Shared read-only handles for standalone modules such as survey.js.
window.supabaseClient = supabaseClient;
Object.defineProperty(window, 'currentUser', {
  configurable: true,
  get: () => currentUser
});

// Returns true if the current user can write (add/edit/delete) segments.
function canWrite() {
  return !!currentUser;
}

// Enforce write-only access: show a status message and abort if not logged in.
function requireAuth(action = 'perform this action') {
  if (!canWrite()) {
    showStatus(`You must be signed in to ${action}. Use the sign-in form in Analyst View.`, true);
    return false;
  }
  return true;
}

// Load all public segments for guest users (anon read via RLS policy).
async function loadPublicSegmentsForGuest() {
  if (!supabaseClient) {
    dataset = [];
    expandedSnippetIds.clear();
    refreshAll();
    return;
  }
  setCloudSyncing(true);
  let data = null;
  let error = null;
  try {
    const result = await supabaseClient
      .from(SUPABASE_TABLE)
      .select('*')
      .order('created_at', { ascending: true, nullsFirst: true })
      .order('encoded_at', { ascending: true, nullsFirst: true })
      .order('segment_id', { ascending: true });
    data = result.data;
    error = result.error;
  } catch (fetchError) {
    error = fetchError;
  }
  setCloudSyncing(false);
  if (error) {
    // Silently fail for guests — no segments to show is acceptable
    dataset = [];
    expandedSnippetIds.clear();
    refreshAll();
    return;
  }
  dataset = mapRowsToCipqSegments(data);
  expandedSnippetIds.clear();
  refreshAll();
  if (dataset.length) showStatus(`Viewing ${dataset.length} published segment${dataset.length !== 1 ? 's' : ''} (read-only). Sign in to encode.`, false);
}

const CODEBOOK = {
  Creation: [
    {
      code: 'CREATE_AUTHOR_SUPPORT',
      label: 'Author Support Constraints',
      indicators: [
        { code: 'CREATE_AUTHOR_SUPPORT_FINANCE', label: 'Limited financial support for authors' },
        { code: 'CREATE_AUTHOR_SUPPORT_GRANTS', label: 'Lack of grants or funding' },
        { code: 'CREATE_AUTHOR_SUPPORT_ROYALTY', label: 'Low compensation or royalties' }
      ]
    },
    {
      code: 'CREATE_CONTENT_DEV',
      label: 'Content Development Limitations',
      indicators: [
        { code: 'CREATE_CONTENT_DEV_EDITORIAL', label: 'Limited editorial development' },
        { code: 'CREATE_CONTENT_DEV_RESEARCH', label: 'Lack of research support' },
        { code: 'CREATE_CONTENT_DEV_PIPELINE', label: 'Weak pipeline for new authors' }
      ]
    },
    {
      code: 'CREATE_LABOR',
      label: 'Creative Labor Conditions',
      indicators: [
        { code: 'CREATE_LABOR_PRECARITY', label: 'Precarious creative work' },
        { code: 'CREATE_LABOR_INSTITUTION', label: 'Lack of institutional support' },
        { code: 'CREATE_LABOR_FREELANCE', label: 'Over-reliance on freelance work' }
      ]
    },
    {
      code: 'CREATE_IP',
      label: 'Intellectual Property Issues',
      indicators: [
        { code: 'CREATE_IP_PROTECTION', label: 'Weak IP protection' },
        { code: 'CREATE_IP_PIRACY', label: 'Piracy affecting authors' },
        { code: 'CREATE_IP_CONTRACT', label: 'Contractual imbalances' }
      ]
    },
    {
      code: 'CREATE_CONTENT_GAPS',
      label: 'Cultural Content Gaps',
      indicators: [
        { code: 'CREATE_CONTENT_GAPS_LOCAL', label: 'Lack of local or regional content' },
        { code: 'CREATE_CONTENT_GAPS_LANGUAGE', label: 'Underrepresentation of languages' },
        { code: 'CREATE_CONTENT_GAPS_GENRE', label: 'Genre imbalance' }
      ]
    }
  ],
  Production: [
    {
      code: 'PROD_PRINT_COST',
      label: 'Printing Cost Issues',
      indicators: [
        { code: 'PROD_PRINT_COST_LOCAL', label: 'High local printing cost' },
        { code: 'PROD_PRINT_COST_OFFSHORE', label: 'Offshore printing due to cost' },
        { code: 'PROD_PRINT_COST_MATERIAL', label: 'Material cost pressure' }
      ]
    },
    {
      code: 'PROD_CAPACITY',
      label: 'Production Capacity Constraints',
      indicators: [
        { code: 'PROD_CAPACITY_FACILITY', label: 'Limited printing facilities' },
        { code: 'PROD_CAPACITY_DELAY', label: 'Delays in production' },
        { code: 'PROD_CAPACITY_SCALE', label: 'Limited scale capability' }
      ]
    },
    {
      code: 'PROD_INFRA',
      label: 'Infrastructure Limitations',
      indicators: [
        { code: 'PROD_INFRA_EQUIPMENT', label: 'Outdated equipment' },
        { code: 'PROD_INFRA_INVESTMENT', label: 'Lack of investment' },
        { code: 'PROD_INFRA_SYSTEM', label: 'Weak production ecosystem' }
      ]
    },
    {
      code: 'PROD_MATERIAL_COST',
      label: 'Cost of Materials',
      indicators: [
        { code: 'PROD_MATERIAL_COST_PAPER', label: 'Paper cost volatility' },
        { code: 'PROD_MATERIAL_COST_IMPORT', label: 'Import dependency' },
        { code: 'PROD_MATERIAL_COST_SUPPLY', label: 'Supply instability' }
      ]
    },
    {
      code: 'PROD_QUALITY',
      label: 'Quality Control Issues',
      indicators: [
        { code: 'PROD_QUALITY_PRINT', label: 'Inconsistent print quality' },
        { code: 'PROD_QUALITY_STANDARD', label: 'Lack of standards' },
        { code: 'PROD_QUALITY_SKILL', label: 'Skilled labor gaps' }
      ]
    }
  ],
  Distribution: [
    {
      code: 'DIST_LOGISTICS',
      label: 'Logistics Constraints',
      indicators: [
        { code: 'DIST_LOGISTICS_COST', label: 'High shipping cost' },
        { code: 'DIST_LOGISTICS_DELAY', label: 'Slow delivery' },
        { code: 'DIST_LOGISTICS_REGIONAL', label: 'Regional distribution gaps' }
      ]
    },
    {
      code: 'DIST_BOOKSTORE',
      label: 'Bookstore Access Issues',
      indicators: [
        { code: 'DIST_BOOKSTORE_DECLINE', label: 'Decline of physical bookstores' },
        { code: 'DIST_BOOKSTORE_SPACE', label: 'Limited shelf space' },
        { code: 'DIST_BOOKSTORE_URBAN', label: 'Concentration in urban areas' }
      ]
    },
    {
      code: 'DIST_SUPPLY_CHAIN',
      label: 'Supply Chain Fragmentation',
      indicators: [
        { code: 'DIST_SUPPLY_CHAIN_COORD', label: 'Weak coordination' },
        { code: 'DIST_SUPPLY_CHAIN_INVENTORY', label: 'Inventory issues' },
        { code: 'DIST_SUPPLY_CHAIN_EFFICIENCY', label: 'Distribution inefficiency' }
      ]
    },
    {
      code: 'DIST_MARKET_ACCESS',
      label: 'Market Access Barriers',
      indicators: [
        { code: 'DIST_MARKET_ACCESS_ENTRY', label: 'Difficulty entering markets' },
        { code: 'DIST_MARKET_ACCESS_DOMINANCE', label: 'Dominance of large players' },
        { code: 'DIST_MARKET_ACCESS_EXPORT', label: 'Limited export channels' }
      ]
    },
    {
      code: 'DIST_DIGITAL',
      label: 'Digital Distribution Gaps',
      indicators: [
        { code: 'DIST_DIGITAL_EBOOK', label: 'Weak e-book infrastructure' },
        { code: 'DIST_DIGITAL_PLATFORM', label: 'Platform dependency' },
        { code: 'DIST_DIGITAL_REACH', label: 'Limited online reach' }
      ]
    }
  ],
  Access: [
    {
      code: 'ACCESS_AFFORDABILITY',
      label: 'Affordability Issues',
      indicators: [
        { code: 'ACCESS_AFFORDABILITY_PRICE', label: 'High book prices' },
        { code: 'ACCESS_AFFORDABILITY_INCOME', label: 'Income constraints' },
        { code: 'ACCESS_AFFORDABILITY_SUBSIDY', label: 'Limited subsidies' }
      ]
    },
    {
      code: 'ACCESS_AVAILABILITY',
      label: 'Availability Constraints',
      indicators: [
        { code: 'ACCESS_AVAILABILITY_COPIES', label: 'Limited copies in circulation' },
        { code: 'ACCESS_AVAILABILITY_STOCK', label: 'Stock shortages' },
        { code: 'ACCESS_AVAILABILITY_REGION', label: 'Regional unavailability' }
      ]
    },
    {
      code: 'ACCESS_LITERACY',
      label: 'Literacy and Engagement',
      indicators: [
        { code: 'ACCESS_LITERACY_CULTURE', label: 'Low reading culture' },
        { code: 'ACCESS_LITERACY_PROGRAM', label: 'Limited literacy programs' },
        { code: 'ACCESS_LITERACY_ENGAGE', label: 'Weak reader engagement' }
      ]
    },
    {
      code: 'ACCESS_LIBRARY',
      label: 'Library Access Issues',
      indicators: [
        { code: 'ACCESS_LIBRARY_FUND', label: 'Underfunded libraries' },
        { code: 'ACCESS_LIBRARY_COLLECTION', label: 'Limited collections' },
        { code: 'ACCESS_LIBRARY_ACCESS', label: 'Poor accessibility' }
      ]
    },
    {
      code: 'ACCESS_DIGITAL',
      label: 'Digital Access Inequality',
      indicators: [
        { code: 'ACCESS_DIGITAL_INTERNET', label: 'Limited internet access' },
        { code: 'ACCESS_DIGITAL_DEVICE', label: 'Device inequality' },
        { code: 'ACCESS_DIGITAL_PAYWALL', label: 'Platform paywalls' }
      ]
    }
  ]
};

const THEME_INDEX = {};
const INDICATOR_INDEX = {};
Object.entries(CODEBOOK).forEach(([domain, themes]) => {
  themes.forEach(theme => {
    THEME_INDEX[theme.code] = { ...theme, domain };
    theme.indicators.forEach(indicator => {
      INDICATOR_INDEX[indicator.code] = { ...indicator, domain, themeCode: theme.code, themeLabel: theme.label };
    });
  });
});

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function roundTo(value, digits = 2) {
  if (!Number.isFinite(value)) return 0;
  const factor = 10 ** digits;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function mean(values) {
  if (!values.length) return 0;
  return values.reduce((total, value) => total + value, 0) / values.length;
}

function sum(values) {
  return values.reduce((total, value) => total + value, 0);
}

function countDistinct(values) {
  return new Set(values.filter(value => value !== null && value !== undefined && String(value).trim() !== '')).size;
}

function percentile(values, percentileValue) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const index = (percentileValue / 100) * (sorted.length - 1);
  const lower = Math.floor(index);
  const upper = Math.ceil(index);
  if (lower === upper) return sorted[lower];
  const weight = index - lower;
  return sorted[lower] * (1 - weight) + sorted[upper] * weight;
}

function variance(values) {
  if (!values.length) return 0;
  const average = mean(values);
  return mean(values.map(value => Math.pow(value - average, 2)));
}

function normalizeValue(value, maxValue) {
  if (!maxValue) return 0;
  return Math.min(1, value / maxValue);
}

function toArray(value) {
  if (Array.isArray(value)) return value.filter(Boolean);
  if (!value) return [];
  return String(value).split(/[|,;]/).map(item => item.trim()).filter(Boolean);
}

function normalizeControlledOption(value, options) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  return options.find(option => option.toLowerCase() === raw.toLowerCase()) || raw;
}

function normalizeValueChainStage(value) {
  const normalized = normalizeControlledOption(value, VALUE_CHAIN_STAGES);
  return VALUE_CHAIN_ALIASES[normalized] || normalized;
}

function valueChainStageForAnalysis(record) {
  return record.Value_Chain_Stage || DOMAIN_TO_VALUE_CHAIN_STAGE[record.CIPQ_Domain] || '';
}

function normalizePestleTags(value) {
  const seen = new Set();
  return toArray(value)
    .map(tag => normalizeControlledOption(tag, PESTLE_TAGS))
    .filter(tag => PESTLE_TAGS.includes(tag) && !seen.has(tag) && seen.add(tag));
}

function parseBoolean(value) {
  if (typeof value === 'boolean') return value;
  return ['true', '1', 'yes', 'y'].includes(String(value ?? '').trim().toLowerCase());
}

function formatSupabaseError(error) {
  if (!error) return 'Unknown error.';
  const message = error.message || String(error);
  if (/failed to fetch|networkerror|network request failed|name_not_resolved|err_name_not_resolved|load failed/i.test(message)) {
    return 'Could not reach Supabase. Check your internet connection, DNS/VPN/firewall settings, or confirm that the Supabase project URL is active: ueyyrugaynzczkcwnxbt.supabase.co. You can still work locally and export your data.';
  }
  if (/column|schema|record_confidence|theme_code|stakeholder_group|quadrant_primary|indicator_label|value_chain_stage|pestle_tags/i.test(message)) {
    return `${message} Run the updated SQL migration in supabase_schema.sql so the richer CIPQ fields exist before using cloud sync.`;
  }
  return message;
}

function mostCommon(values, maxItems = 3) {
  const counts = new Map();
  values.filter(Boolean).forEach(value => counts.set(value, (counts.get(value) || 0) + 1));
  return [...counts.entries()]
    .sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])))
    .slice(0, maxItems)
    .map(([label, count]) => ({ label, count }));
}

function inferDomainFromCode(code) {
  const upper = String(code || '').toUpperCase();
  for (const domain of VALID_DOMAINS) {
    if (upper.startsWith(`${DOMAIN_PREFIX[domain]}_`)) return domain;
  }
  if (upper.startsWith('C')) return 'Creation';
  if (upper.startsWith('P') || upper.startsWith('G')) return 'Production';
  if (upper.startsWith('D')) return 'Distribution';
  if (upper.startsWith('A')) return 'Access';
  return '';
}

function getThemeLabel(themeCode) {
  return THEME_INDEX[themeCode]?.label || '';
}

function getIndicatorLabel(indicatorCode) {
  return INDICATOR_INDEX[indicatorCode]?.label || indicatorCode || '';
}

function themeCodeForIndicator(indicatorCode) {
  return INDICATOR_INDEX[indicatorCode]?.themeCode || '';
}

function ensureUniqueSegmentId(baseId, takenIds = null) {
  const ids = takenIds || new Set(dataset.map(record => record.Segment_ID));
  const safeBase = String(baseId || `SEG_${Date.now()}`).trim() || `SEG_${Date.now()}`;
  let candidate = safeBase;
  let suffix = 1;
  while (ids.has(candidate)) {
    candidate = `${safeBase}_${suffix}`;
    suffix += 1;
  }
  ids.add(candidate);
  return candidate;
}

function normalizeDuplicateText(value) {
  return String(value ?? '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function recordIdentityKey(record) {
  if (record.DB_ID) return `db:${record.DB_ID}`;
  return `segment:${normalizeDuplicateText(record.Segment_ID)}`;
}

function findDuplicateMatches(incomingRecord, records = dataset) {
  const incomingSnippet = normalizeDuplicateText(incomingRecord.Snippet);
  if (!incomingSnippet) return [];

  return records.map(existingRecord => {
    const reasons = [];
    if (incomingSnippet === normalizeDuplicateText(existingRecord.Snippet)) {
      reasons.push('same snippet');
    }
    return reasons.length ? { record: existingRecord, reasons } : null;
  }).filter(Boolean);
}

function removeRecordReferences(records, references) {
  if (!references.size) return records;
  return records.filter(record => !references.has(record));
}

function planSegmentReplacements(incomingRecords) {
  let finalDataset = [...dataset];
  let recordsToSave = [];
  const replacedExisting = new Map();
  const duplicateSummaries = [];

  incomingRecords.forEach(record => {
    const matches = findDuplicateMatches(record, finalDataset);
    if (matches.length) {
      duplicateSummaries.push({ record, matches });
      matches.forEach(match => {
        if (dataset.includes(match.record)) {
          replacedExisting.set(recordIdentityKey(match.record), match.record);
        }
      });
    }

    const matchedRecords = new Set(matches.map(match => match.record));
    finalDataset = removeRecordReferences(finalDataset, matchedRecords);
    recordsToSave = removeRecordReferences(recordsToSave, matchedRecords);
    finalDataset.push(record);
    recordsToSave.push(record);
  });

  return {
    finalDataset,
    recordsToSave,
    replacedExisting: [...replacedExisting.values()],
    duplicateSummaries
  };
}

function duplicateSummaryText(duplicateSummaries) {
  const reasonCounts = new Map();
  duplicateSummaries.forEach(summary => {
    summary.matches.forEach(match => {
      match.reasons.forEach(reason => reasonCounts.set(reason, (reasonCounts.get(reason) || 0) + 1));
    });
  });
  return [...reasonCounts.entries()]
    .map(([reason, count]) => `${count} ${reason}`)
    .join(', ');
}

function renderDuplicateDetailItem(summary, isHidden = false) {
  const reasons = [...new Set(summary.matches.flatMap(match => match.reasons))].join(', ');
  const snippet = summary.record.Snippet ? summary.record.Snippet.slice(0, 120) : summary.record.Segment_ID;
  return `<li${isHidden ? ' class="is-hidden"' : ''}><strong>${escapeHtml(summary.record.Segment_ID || 'New segment')}</strong> - ${escapeHtml(reasons)}<span>${escapeHtml(snippet || '')}</span></li>`;
}

function requestOverwriteConfirmation(duplicateSummaries, sourceLabel = 'segment') {
  if (!duplicateSummaries.length) return Promise.resolve(true);

  const modal = document.getElementById('overwriteModal');
  if (!modal) {
    const count = duplicateSummaries.length;
    return Promise.resolve(window.confirm(`${count} duplicate ${count === 1 ? 'record was' : 'records were'} found. Overwrite the existing and replace?`));
  }

  const title = document.getElementById('overwriteModalTitle');
  const body = document.getElementById('overwriteModalBody');
  const details = document.getElementById('overwriteModalDetails');
  const viewMoreButton = document.getElementById('overwriteViewMoreBtn');
  const confirmButton = document.getElementById('overwriteConfirmBtn');
  const cancelButton = document.getElementById('overwriteCancelBtn');

  const duplicateCount = duplicateSummaries.length;
  const replacedCount = duplicateSummaries.reduce((total, summary) => total + summary.matches.length, 0);
  title.textContent = duplicateCount === 1 ? 'Duplicate Segment Found' : 'Duplicate Segments Found';
  body.textContent = `${duplicateCount} ${sourceLabel}${duplicateCount === 1 ? '' : 's'} matched ${replacedCount} existing record${replacedCount === 1 ? '' : 's'}. Overwrite the existing and replace?`;
  details.innerHTML = duplicateSummaries
    .map((summary, index) => renderDuplicateDetailItem(summary, index >= 4))
    .join('');
  if (viewMoreButton) {
    const hiddenCount = Math.max(duplicateSummaries.length - 4, 0);
    viewMoreButton.textContent = hiddenCount ? `View More (${hiddenCount})` : 'View More';
    viewMoreButton.classList.toggle('show', hiddenCount > 0);
    viewMoreButton.setAttribute('aria-expanded', 'false');
  }

  return new Promise(resolve => {
    let settled = false;
    const close = confirmed => {
      if (settled) return;
      settled = true;
      modal.classList.remove('show');
      modal.setAttribute('aria-hidden', 'true');
      confirmButton.removeEventListener('click', onConfirm);
      cancelButton.removeEventListener('click', onCancel);
      viewMoreButton?.removeEventListener('click', onViewMore);
      modal.removeEventListener('click', onBackdrop);
      document.removeEventListener('keydown', onKeydown);
      resolve(confirmed);
    };
    const onConfirm = () => close(true);
    const onCancel = () => close(false);
    const onViewMore = () => {
      const isExpanded = viewMoreButton.getAttribute('aria-expanded') === 'true';
      details.querySelectorAll('.is-hidden').forEach(item => {
        item.style.display = isExpanded ? 'none' : 'block';
      });
      viewMoreButton.setAttribute('aria-expanded', String(!isExpanded));
      viewMoreButton.textContent = isExpanded ? `View More (${duplicateSummaries.length - 4})` : 'View Less';
    };
    const onBackdrop = event => {
      if (event.target === modal) close(false);
    };
    const onKeydown = event => {
      if (event.key === 'Escape') close(false);
    };

    confirmButton.addEventListener('click', onConfirm);
    cancelButton.addEventListener('click', onCancel);
    viewMoreButton?.addEventListener('click', onViewMore);
    modal.addEventListener('click', onBackdrop);
    document.addEventListener('keydown', onKeydown);
    modal.classList.add('show');
    modal.setAttribute('aria-hidden', 'false');
    confirmButton.focus();
  });
}

async function deleteReplacedSegmentsFromSupabase(records) {
  if (!records.length) return null;

  const dbIds = [...new Set(records.map(record => record.DB_ID).filter(Boolean))];
  if (dbIds.length) {
    return supabaseClient.from(SUPABASE_TABLE).delete().in('id', dbIds);
  }

  const segmentIds = [...new Set(records.map(record => record.Segment_ID).filter(Boolean))];
  if (!segmentIds.length) return null;
  return supabaseClient
    .from(SUPABASE_TABLE)
    .delete()
    .eq('user_id', currentUser.id)
    .in('segment_id', segmentIds);
}

function chunkItems(items, size) {
  const chunks = [];
  for (let i = 0; i < items.length; i += size) {
    chunks.push(items.slice(i, i + size));
  }
  return chunks;
}

function normalizeSegmentRecord(record) {
  const indicatorCode = record.Indicator_Code || record.indicator_code || record['indicator.code'] || record.Indicator || '';
  const themeCode = record.Theme_Code || record.theme_code || record['theme.code'] || themeCodeForIndicator(indicatorCode) || '';
  const rawPrimary = record.CIPQ_Domain || record.cipq_domain || record.Quadrant_Primary || record.quadrant_primary || record['quadrant.primary'] || record.Domain || inferDomainFromCode(themeCode) || inferDomainFromCode(indicatorCode);
  const primary = LEGACY_DOMAIN_MAP[rawPrimary] || rawPrimary || '';
  const rawSecondary = record.Secondary_Domain || record.secondary_domain || record.Quadrant_Secondary || record.quadrant_secondary || record['quadrant.secondary'] || '';
  const secondary = LEGACY_DOMAIN_MAP[rawSecondary] || rawSecondary || '';
  const severityValue = parseInt(record.Severity ?? record.severity ?? record['scoring.severity'], 10);
  const scoringConfidence = String(record.Scoring_Confidence || record.Record_Confidence || record.record_confidence || record.Confidence || record['scoring.confidence'] || 'medium').toLowerCase();
  const linkedQuadrants = toArray(record.Linked_Quadrants || record.linked_quadrants || record['analysis_flags.linked_quadrants']);
  const valueChainStage = normalizeValueChainStage(
    record.Value_Chain_Stage || record.value_chain_stage || record['context.value_chain_stage'] || record.ValueChainStage || record['Value Chain Stage'],
  );
  const pestleTags = normalizePestleTags(
    record.PESTLE_Tags || record.pestle_tags || record['context.pestle_tags'] || record.PESTLE || record['PESTLE Tags']
  );
  const createdAt = record.Created_At || record.created_at || record.Encoded_At || record.encoded_at || record['timestamps.encoded_at'] || new Date().toISOString();
  const updatedAt = record.Updated_At || record.updated_at || record['timestamps.updated_at'] || createdAt;

  return {
    Segment_ID: record.Segment_ID || record.segment_id || record.id || '',
    Snippet: record.Snippet || record.snippet || record.text_segment || record.Text || record.Quote || '',
    Theme: record.Theme || record.Theme_Label || record.theme_label || record['theme.label'] || getThemeLabel(themeCode) || '',
    Theme_Code: themeCode,
    Open_Code: record.Open_Code || record.open_code || '',
    CIPQ_Domain: primary,
    Secondary_Domain: secondary,
    Indicator_Code: indicatorCode,
    Indicator_Name: record.Indicator_Name || record.Indicator_Label || record.indicator_label || record['indicator.label'] || getIndicatorLabel(indicatorCode) || '',
    Severity: Number.isFinite(severityValue) ? severityValue : NaN,
    Stakeholder: record.Stakeholder || record.Stakeholder_Group || record.stakeholder || record.stakeholder_group || record['metadata.stakeholder_group'] || '',
    Respondent_Type: record.Respondent_Type || record.respondent_type || record['metadata.respondent_type'] || '',
    Region: record.Region || record.region || record['metadata.region'] || '',
    Source_Type: normalizeCipqSourceType(record.Source_Type || record.source_type || record.Source || record['metadata.source_type'] || ''),
    Source_ID: record.Source_ID || record.source_id || record['metadata.source_id'] || '',
    Value_Chain_Stage: VALUE_CHAIN_STAGES.includes(valueChainStage) ? valueChainStage : '',
    PESTLE_Tags: pestleTags,
    Scoring_Confidence: SCORING_CONFIDENCE_OPTIONS.includes(scoringConfidence) ? scoringConfidence : 'medium',
    Is_Cross_Quadrant: parseBoolean(record.Is_Cross_Quadrant ?? record.is_cross_quadrant ?? record['analysis_flags.is_cross_quadrant']) || !!secondary || linkedQuadrants.length > 1,
    Linked_Quadrants: linkedQuadrants.length ? linkedQuadrants : [primary, secondary].filter(Boolean),
    Analysis_Notes: record.Analysis_Notes || record.analysis_notes || record['analysis_flags.notes'] || '',
    Session_ID: record.Session_ID || record.session_id || '',
    DB_ID: record.DB_ID || record.db_id || null,
    Created_At: createdAt,
    Updated_At: updatedAt
  };
}

function isSurveySourceType(value) {
  return String(value || '').trim().toLowerCase() === 'survey';
}

function normalizeCipqSourceType(value) {
  const raw = String(value || '').trim();
  const lower = raw.toLowerCase();
  if (lower === 'fgd' || lower === 'focus group' || lower === 'focus group discussion') return 'FGD';
  if (lower === 'kii' || lower === 'key informant' || lower === 'key informant interview') return 'KII';
  if (lower === 'document' || lower === 'document review') return 'Document';
  if (lower === 'survey') return 'Survey';
  return raw;
}

function toInterpretiveRecord(record) {
  return {
    id: record.Segment_ID,
    text_segment: record.Snippet,
    theme: {
      label: record.Theme,
      code: record.Theme_Code
    },
    quadrant: {
      primary: record.CIPQ_Domain,
      secondary: record.Secondary_Domain || null
    },
    indicator: {
      label: record.Indicator_Name,
      code: record.Indicator_Code
    },
    metadata: {
      stakeholder_group: record.Stakeholder || null,
      respondent_type: record.Respondent_Type || null,
      region: record.Region || null,
      source_type: record.Source_Type || null,
      source_id: record.Source_ID || null
    },
    context: {
      value_chain_stage: record.Value_Chain_Stage || null,
      pestle_tags: record.PESTLE_Tags || []
    },
    scoring: {
      severity: record.Severity,
      confidence: record.Scoring_Confidence
    },
    analysis_flags: {
      is_cross_quadrant: !!record.Is_Cross_Quadrant,
      linked_quadrants: record.Linked_Quadrants || [],
      notes: record.Analysis_Notes || ''
    },
    timestamps: {
      encoded_at: record.Created_At,
      updated_at: record.Updated_At
    }
  };
}

function mapSegmentToDbRow(record) {
  return {
    user_id: currentUser?.id,
    segment_id: record.Segment_ID,
    snippet: record.Snippet,
    theme: record.Theme || null,
    theme_code: record.Theme_Code || null,
    theme_label: record.Theme || null,
    open_code: record.Open_Code || null,
    cipq_domain: record.CIPQ_Domain,
    quadrant_primary: record.CIPQ_Domain,
    secondary_domain: record.Secondary_Domain || null,
    quadrant_secondary: record.Secondary_Domain || null,
    indicator_code: record.Indicator_Code,
    indicator_name: record.Indicator_Name,
    indicator_label: record.Indicator_Name,
    severity: record.Severity,
    stakeholder: record.Stakeholder || null,
    stakeholder_group: record.Stakeholder || null,
    respondent_type: record.Respondent_Type || null,
    region: record.Region || null,
    source_type: record.Source_Type || null,
    source_id: record.Source_ID || null,
    value_chain_stage: record.Value_Chain_Stage || null,
    pestle_tags: record.PESTLE_Tags?.length ? record.PESTLE_Tags : null,
    record_confidence: record.Scoring_Confidence || null,
    is_cross_quadrant: !!record.Is_Cross_Quadrant,
    linked_quadrants: record.Linked_Quadrants?.length ? record.Linked_Quadrants : null,
    analysis_notes: record.Analysis_Notes || null,
    session_id: record.Session_ID || null,
    encoded_at: record.Created_At || new Date().toISOString(),
    updated_at: new Date().toISOString()
  };
}

function mapDbRowToSegment(row) {
  return normalizeSegmentRecord({
    Segment_ID: row.segment_id,
    Snippet: row.snippet,
    Theme: row.theme_label || row.theme,
    Theme_Code: row.theme_code,
    Open_Code: row.open_code,
    CIPQ_Domain: row.quadrant_primary || row.cipq_domain,
    Secondary_Domain: row.quadrant_secondary || row.secondary_domain,
    Indicator_Code: row.indicator_code,
    Indicator_Name: row.indicator_label || row.indicator_name,
    Severity: row.severity,
    Stakeholder: row.stakeholder_group || row.stakeholder,
    Respondent_Type: row.respondent_type,
    Region: row.region,
    Source_Type: row.source_type,
    Source_ID: row.source_id,
    Value_Chain_Stage: row.value_chain_stage,
    PESTLE_Tags: row.pestle_tags,
    Scoring_Confidence: row.record_confidence,
    Is_Cross_Quadrant: row.is_cross_quadrant,
    Linked_Quadrants: row.linked_quadrants,
    Analysis_Notes: row.analysis_notes,
    Session_ID: row.session_id,
    DB_ID: row.id,
    Created_At: row.encoded_at || row.created_at,
    Updated_At: row.updated_at || row.created_at
  });
}

function mapRowsToCipqSegments(rows) {
  return (rows || [])
    .map(mapDbRowToSegment)
    .filter(record => !isSurveySourceType(record.Source_Type));
}

function getMissingSupabaseColumn(error) {
  const message = error?.message || '';
  const quotedMatch = message.match(/'([a-z0-9_]+)' column/i);
  if (quotedMatch) return quotedMatch[1];
  const namedMatch = message.match(/column "?([a-z0-9_]+)"?/i);
  return namedMatch ? namedMatch[1] : '';
}

function stripUnavailableSupabaseColumns(row) {
  const cleaned = { ...row };
  unavailableSupabaseColumns.forEach(column => {
    delete cleaned[column];
  });
  return cleaned;
}

async function insertSegmentsToSupabase(records, options = {}) {
  const isBatch = Array.isArray(records);
  const recordList = isBatch ? records : [records];
  let attempts = 0;

  while (attempts < 20) {
    const rows = recordList.map(record => stripUnavailableSupabaseColumns(mapSegmentToDbRow(record)));
    let query = supabaseClient
      .from(SUPABASE_TABLE)
      .insert(isBatch ? rows : rows[0]);

    if (options.selectSingle) query = query.select().single();
    else if (options.select) query = query.select();

    const result = await query;
    const missingColumn = getMissingSupabaseColumn(result.error);
    if (missingColumn && !unavailableSupabaseColumns.has(missingColumn)) {
      unavailableSupabaseColumns.add(missingColumn);
      attempts += 1;
      continue;
    }

    return result;
  }

  return {
    data: null,
    error: { message: 'Supabase save failed after retrying without missing schema columns.' }
  };
}

function validateRecord(record, strict = false) {
  const issues = [];
  if (!record.Segment_ID) issues.push('Missing segment ID.');
  if (!record.Snippet) issues.push('Missing text segment.');
  if (!VALID_DOMAINS.includes(record.CIPQ_Domain)) issues.push('Primary quadrant must be one of Creation, Production, Distribution, or Access.');
  if (!Number.isFinite(record.Severity) || record.Severity < 1 || record.Severity > 5) issues.push('Severity must be between 1 and 5.');
  if (!record.Theme_Code) issues.push('Missing theme code.');
  if (!record.Indicator_Code) issues.push('Missing indicator code.');
  if (!record.Stakeholder) issues.push('Missing stakeholder group.');
  if (!record.Region) issues.push('Missing region.');
  if (!record.Source_Type) issues.push('Missing source type.');
  if (isSurveySourceType(record.Source_Type)) issues.push('Survey responses belong in the separate Survey Data module and are not encoded as CIPQ indicators.');
  if (record.Source_Type && !isSurveySourceType(record.Source_Type) && !CIPQ_SOURCE_TYPES.includes(record.Source_Type)) {
    issues.push('Source type must be FGD, KII, or Document for CIPQ narrative coding.');
  }
  if (!record.Source_ID) issues.push('Missing source ID.');
  if (record.Value_Chain_Stage && !VALUE_CHAIN_STAGES.includes(record.Value_Chain_Stage)) issues.push('Value chain stage must be Development, Production, Distribution, or Market Access.');
  const invalidPestleTags = (record.PESTLE_Tags || []).filter(tag => !PESTLE_TAGS.includes(tag));
  if (invalidPestleTags.length) issues.push(`Invalid PESTLE tag: ${invalidPestleTags[0]}.`);

  const prefix = DOMAIN_PREFIX[record.CIPQ_Domain];
  if (record.Theme_Code && prefix && !String(record.Theme_Code).toUpperCase().startsWith(`${prefix}_`)) {
    issues.push(`Theme code ${record.Theme_Code} does not match the ${record.CIPQ_Domain} prefix.`);
  }

  const theme = THEME_INDEX[record.Theme_Code];
  if (theme && theme.domain !== record.CIPQ_Domain) {
    issues.push(`Theme code ${record.Theme_Code} maps to ${theme.domain}, not ${record.CIPQ_Domain}.`);
  }

  const indicator = INDICATOR_INDEX[record.Indicator_Code];
  if (indicator && theme && indicator.themeCode !== theme.code) {
    issues.push(`Indicator code ${record.Indicator_Code} is not mapped to theme code ${record.Theme_Code}.`);
  }
  if (indicator && indicator.domain !== record.CIPQ_Domain) {
    issues.push(`Indicator code ${record.Indicator_Code} maps to ${indicator.domain}, not ${record.CIPQ_Domain}.`);
  }

  if (strict && !theme) issues.push(`Theme code ${record.Theme_Code} is not in the active codebook.`);
  if (strict && !indicator) issues.push(`Indicator code ${record.Indicator_Code} is not in the active codebook.`);

  return issues;
}

function buildThemeOptions() {
  const select = document.getElementById('f_theme_code');
  if (!select) return;
  select.innerHTML = '<option value="">-- Select Theme --</option>';
  VALID_DOMAINS.forEach(domain => {
    const optgroup = document.createElement('optgroup');
    optgroup.label = domain;
    CODEBOOK[domain].forEach(theme => {
      const option = document.createElement('option');
      option.value = theme.code;
      option.textContent = `${theme.code} | ${theme.label}`;
      optgroup.appendChild(option);
    });
    select.appendChild(optgroup);
  });
}

function updateThemeSelection() {
  const themeCode = document.getElementById('f_theme_code').value;
  const theme = THEME_INDEX[themeCode];
  document.getElementById('f_theme_label').value = theme?.label || '';
  document.getElementById('f_domain').value = theme?.domain || '';
  const indicatorSelect = document.getElementById('f_indicator');
  indicatorSelect.innerHTML = '<option value="">-- Select Indicator --</option>';
  if (theme) {
    theme.indicators.forEach(indicator => {
      const option = document.createElement('option');
      option.value = indicator.code;
      option.textContent = `${indicator.code} | ${indicator.label}`;
      indicatorSelect.appendChild(option);
    });
  } else {
    indicatorSelect.innerHTML = '<option value="">-- Select Theme First --</option>';
  }
  document.getElementById('f_indicator_label').value = '';
}

function updateIndicatorMetadata() {
  const indicatorCode = document.getElementById('f_indicator').value;
  document.getElementById('f_indicator_label').value = INDICATOR_INDEX[indicatorCode]?.label || getIndicatorLabel(indicatorCode) || '';
}

function authCredentials() {
  return {
    username: document.getElementById('auth_username').value.trim(),
    password: document.getElementById('auth_password').value
  };
}

function clearAuthForm() {
  document.getElementById('auth_username').value = '';
  document.getElementById('auth_password').value = '';
}

function setCloudSyncing(syncing) {
  isCloudSyncing = syncing;
  renderAuthUI();
}

// Locks or unlocks the Analyst View button and encoder tabs based on auth state.
// Guests are forced into Client View and the Analyst View button is hidden.
function enforceGuestView() {
  const encoderBtn = document.getElementById('encoderViewBtn');
  const viewSwitchNote = document.querySelector('.view-switch-note');
  const guestBanner = document.getElementById('guestBanner');
  if (!encoderBtn) return;

  if (!currentUser) {
    // Hide the Analyst View toggle button for guests
    encoderBtn.style.display = 'none';
    if (viewSwitchNote) viewSwitchNote.textContent = 'Sign in to access Analyst View and encode segments.';
    if (guestBanner) guestBanner.style.display = 'block';
    // Force client view
    if (activeView === 'encoder') setAppView('client');
  } else {
    encoderBtn.style.display = '';
    if (viewSwitchNote) viewSwitchNote.textContent = 'Analyst View keeps coding, validation, and dataset management together. Client View keeps CIPQ analytics central while adding lightweight Value Chain and PESTLE context for reporting.';
    if (guestBanner) guestBanner.style.display = 'none';
  }
}

function renderAuthUI() {
  const loginBtn      = document.getElementById('authLoginBtn');
  const usernameInput = document.getElementById('auth_username');
  const passwordInput = document.getElementById('auth_password');
  const signInBtn     = document.getElementById('authOpenModalBtn');
  const userBadge     = document.getElementById('userBadge');
  const badgeDot      = document.getElementById('userBadgeDot');
  const badgeEmail    = document.getElementById('userBadgeEmail');
  const dropdownEmail = document.getElementById('dropdownEmail');
  const dropdownNote  = document.getElementById('dropdownNote');

  const authInputsDisabled = !supabaseReady || isCloudSyncing;
  if (loginBtn)      loginBtn.disabled      = authInputsDisabled;
  if (usernameInput) usernameInput.disabled = authInputsDisabled || !!currentUser;
  if (passwordInput) passwordInput.disabled = authInputsDisabled || !!currentUser;

  if (isCloudSyncing) {
    if (signInBtn)  signInBtn.style.display  = 'none';
    if (userBadge)  userBadge.style.display  = 'flex';
    if (badgeDot)   badgeDot.className       = 'badge-dot syncing';
    if (badgeEmail) badgeEmail.textContent   = 'Syncing…';
    return;
  }

  if (currentUser) {
    if (signInBtn)      signInBtn.style.display  = 'none';
    if (userBadge)      userBadge.style.display  = 'flex';
    if (badgeDot)       badgeDot.className        = 'badge-dot';
    if (badgeEmail)     badgeEmail.textContent    = SINGLE_LOGIN_USERNAME;
    if (dropdownEmail)  dropdownEmail.textContent = SINGLE_LOGIN_USERNAME;
    if (dropdownNote)   dropdownNote.textContent  = `${dataset.length} segment${dataset.length !== 1 ? 's' : ''} loaded`;
    // Close the login modal if it's open
    const modal = document.getElementById('loginModal');
    if (modal) modal.classList.remove('open');
    return;
  }

  // Guest / signed-out state
  if (signInBtn) signInBtn.style.display = '';
  if (userBadge) userBadge.style.display = 'none';
  // Close dropdown if open
  const dd = document.getElementById('userDropdown');
  if (dd) dd.classList.remove('open');
}

function openLoginModal() {
  const modal = document.getElementById('loginModal');
  if (!modal) return;
  modal.classList.add('open');
  // Clear any previous error
  const err = document.getElementById('loginModalError');
  if (err) { err.style.display = 'none'; err.textContent = ''; }
  // Focus username field
  setTimeout(() => {
    const e = document.getElementById('auth_username');
    if (e && !e.disabled) e.focus();
  }, 80);
}

function closeLoginModal() {
  const modal = document.getElementById('loginModal');
  if (modal) modal.classList.remove('open');
}

function toggleUserDropdown() {
  const dd = document.getElementById('userDropdown');
  if (dd) dd.classList.toggle('open');
}

// Close dropdown when clicking outside
document.addEventListener('click', (e) => {
  const badge = document.getElementById('userBadge');
  const dd = document.getElementById('userDropdown');
  if (dd && dd.classList.contains('open') && badge && !badge.contains(e.target)) {
    dd.classList.remove('open');
  }
  // Close login modal on backdrop click
  const modal = document.getElementById('loginModal');
  if (modal && modal.classList.contains('open') && e.target === modal) {
    closeLoginModal();
  }
});

// Submit login on Enter key
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    closeLoginModal();
    const dd = document.getElementById('userDropdown');
    if (dd) dd.classList.remove('open');
  }
  if (e.key === 'Enter') {
    const modal = document.getElementById('loginModal');
    if (modal && modal.classList.contains('open')) signInUser();
  }
});

async function loadSegmentsFromSupabase(options = {}) {
  if (!supabaseClient || !currentUser) {
    dataset = [];
    expandedSnippetIds.clear();
    refreshAll();
    renderAuthUI();
    return;
  }

  setCloudSyncing(true);
  let data = null;
  let error = null;
  try {
    const result = await supabaseClient
      .from(SUPABASE_TABLE)
      .select('*')
      .order('created_at', { ascending: true, nullsFirst: true })
      .order('encoded_at', { ascending: true, nullsFirst: true })
      .order('segment_id', { ascending: true });
    data = result.data;
    error = result.error;
  } catch (fetchError) {
    error = fetchError;
  }

  if (error) {
    setCloudSyncing(false);
    showStatus(`Could not load saved segments: ${formatSupabaseError(error)}`, true);
    return;
  }

  dataset = mapRowsToCipqSegments(data);
  expandedSnippetIds.clear();
  refreshAll();
  setCloudSyncing(false);
  if (!options.silent) showStatus(`Loaded ${dataset.length} saved segment${dataset.length !== 1 ? 's' : ''} from Supabase.`, false);
}

async function handleSessionChange(session, options = {}) {
  currentSession = session || null;
  currentUser = session?.user || null;

  if (currentUser && String(currentUser.email || '').toLowerCase() !== SINGLE_LOGIN_AUTH_EMAIL.toLowerCase()) {
    currentSession = null;
    currentUser = null;
    try {
      await supabaseClient?.auth.signOut();
    } catch (error) {
      // Session will still be treated as signed out locally.
    }
    enforceGuestView();
    renderAuthUI();
    if (!options.silent) showStatus('This workspace now accepts only the single username account.', true);
    await loadPublicSegmentsForGuest();
    if (window.loadSurveyData) await window.loadSurveyData();
    return;
  }

  if (!currentUser) {
    // Guest mode: clear local dataset and load public read-only data
    dataset = [];
    expandedSnippetIds.clear();
    // Force client view for guests — no write access
    enforceGuestView();
    renderAuthUI();
    if (!options.silent) showStatus('Signed out. Viewing published segments in read-only mode.', false);
    await loadPublicSegmentsForGuest();
    if (window.loadSurveyData) await window.loadSurveyData();
    return;
  }

  // Logged-in: restore analyst view availability
  enforceGuestView();
  renderAuthUI();
  await loadSegmentsFromSupabase({ silent: !!options.silent });
  if (window.loadSurveyData) await window.loadSurveyData();
}

async function initializeAuth() {
  renderAuthUI();
  enforceGuestView();

  if (!supabaseClient) {
    // No Supabase configured — nothing to load, just enforce guest view
    return;
  }

  let data = null;
  let error = null;
  try {
    const result = await supabaseClient.auth.getSession();
    data = result.data;
    error = result.error;
  } catch (fetchError) {
    error = fetchError;
  }
  if (error) {
    showStatus(`Could not restore session: ${formatSupabaseError(error)}`, true);
    await loadPublicSegmentsForGuest();
    return;
  }

  await handleSessionChange(data?.session || null, { silent: true });
  try {
    supabaseClient.auth.onAuthStateChange((event, session) => {
      handleSessionChange(session, { silent: event === 'INITIAL_SESSION' });
    });
  } catch (fetchError) {
    showStatus(`Could not start Supabase auth listener: ${formatSupabaseError(fetchError)}`, true);
  }
}

async function signInUser() {
  if (!supabaseClient) {
    showLoginModalError('Supabase is not configured yet.');
    return;
  }
  const { username, password } = authCredentials();
  if (!username || !password) {
    showLoginModalError('Enter both username and password to sign in.');
    return;
  }
  if (username !== SINGLE_LOGIN_USERNAME || password !== SINGLE_LOGIN_PASSWORD) {
    showLoginModalError('Invalid username or password.');
    return;
  }

  const loginBtn = document.getElementById('authLoginBtn');
  if (loginBtn) { loginBtn.disabled = true; loginBtn.textContent = 'Signing in…'; }

  setCloudSyncing(true);
  let data = null;
  let error = null;
  try {
    const result = await supabaseClient.auth.signInWithPassword({ email: SINGLE_LOGIN_AUTH_EMAIL, password });
    data = result.data;
    error = result.error;
  } catch (fetchError) {
    error = fetchError;
  }
  setCloudSyncing(false);

  if (loginBtn) { loginBtn.disabled = false; loginBtn.textContent = 'Sign In'; }

  if (error) {
    showLoginModalError(`${formatSupabaseError(error)} In Supabase Auth, create ${SINGLE_LOGIN_AUTH_EMAIL} with this password and either disable email confirmation or mark the user as confirmed.`);
    return;
  }

  clearAuthForm();
  closeLoginModal();
  await handleSessionChange(data.session, { silent: true });
  showStatus('Signed in. Segment history loaded from cloud.', false);
}

function showLoginModalError(msg) {
  const err = document.getElementById('loginModalError');
  if (err) { err.textContent = msg; err.style.display = 'block'; }
}

async function signOutUser() {
  if (!supabaseClient || !currentUser) return;

  setCloudSyncing(true);
  let error = null;
  try {
    const result = await supabaseClient.auth.signOut();
    error = result.error;
  } catch (fetchError) {
    error = fetchError;
  }
  setCloudSyncing(false);
  if (error) {
    showStatus(formatSupabaseError(error), true);
    return;
  }

  clearAuthForm();
  await handleSessionChange(null, { silent: true });
  showStatus('Logged out. Cloud-backed history is hidden until you sign in again.', false);
}

function setAppView(view) {
  activeView = view;
  document.getElementById('encoderViewBtn').classList.toggle('active', view === 'encoder');
  document.getElementById('clientViewBtn').classList.toggle('active', view === 'client');

  const navButtons = [...document.querySelectorAll('#mainNav button')];
  navButtons.forEach(button => {
    button.style.display = button.dataset.view === view ? '' : 'none';
  });

  const activePanel = document.querySelector('.tab-panel.active');
  if (!activePanel || activePanel.dataset.view !== view) {
    const targetButton = navButtons.find(button => button.dataset.view === view);
    if (targetButton) switchTab(targetButton.dataset.tab, targetButton);
  } else {
    refreshAll();
  }
}

function switchTab(name, btn = null) {
  document.querySelectorAll('.tab-panel').forEach(panel => panel.classList.remove('active'));
  document.querySelectorAll('#mainNav button').forEach(button => button.classList.remove('active'));
  const panel = document.getElementById(`tab-${name}`);
  if (panel) panel.classList.add('active');
  if (btn) {
    btn.classList.add('active');
    btn.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
  } else {
    const fallbackButton = document.querySelector(`#mainNav button[data-tab="${name}"]`);
    if (fallbackButton) fallbackButton.classList.add('active');
  }

  if (name === 'dashboard') renderDashboard();
  if (name === 'indicators') renderIndicators();
  if (name === 'comparison') renderComparison();
  if (name === 'priority') renderPriority();
  if (name === 'simulator') renderSimulator();
  if (name === 'dataset') renderDataset();
  if (name === 'entry') renderEntryPreview();
  if (name === 'survey') { if (window.renderSurveyClient) window.renderSurveyClient(); }
  if (name === 'survey-analyst') { if (window.renderSurveyAnalyst) window.renderSurveyAnalyst(); }
}

function setSeverity(value) {
  currentSeverity = value;
  document.getElementById('f_severity').value = value;
  document.querySelectorAll('.sev-picker button').forEach((button, index) => {
    button.classList.toggle('selected', index + 1 <= value);
  });
}

function getMultiSelectValues(id) {
  const node = document.getElementById(id);
  if (!node) return [];
  if (node.tagName === 'SELECT') return [...node.selectedOptions].map(option => option.value).filter(Boolean);
  return [...node.querySelectorAll('input[type="checkbox"]:checked')].map(input => input.value).filter(Boolean);
}

function updatePestleSummary() {
  const summary = document.getElementById('f_pestle_summary');
  if (!summary) return;
  const values = getMultiSelectValues('f_pestle_tags');
  summary.textContent = values.length ? values.join(', ') : '-- Optional --';
}

function clearMultiSelect(id) {
  const node = document.getElementById(id);
  if (!node) return;
  if (node.tagName === 'SELECT') {
    [...node.options].forEach(option => { option.selected = false; });
    return;
  }
  [...node.querySelectorAll('input[type="checkbox"]')].forEach(input => { input.checked = false; });
  node.open = false;
}

function clearForm() {
  ['f_segid', 'f_theme_label', 'f_domain', 'f_indicator_label', 'f_respondent_type', 'f_source_id', 'f_session', 'f_snippet', 'f_analysis_notes'].forEach(id => {
    const node = document.getElementById(id);
    if (node) node.value = '';
  });
  ['f_theme_code', 'f_indicator', 'f_stakeholder', 'f_region', 'f_domain_secondary', 'f_source_type', 'f_value_chain_stage', 'f_confidence'].forEach(id => {
    const node = document.getElementById(id);
    if (node) node.value = id === 'f_confidence' ? 'medium' : '';
  });
  clearMultiSelect('f_pestle_tags');
  updatePestleSummary();
  document.getElementById('f_indicator').innerHTML = '<option value="">-- Select Theme First --</option>';
  currentSeverity = null;
  document.getElementById('f_severity').value = '';
  document.querySelectorAll('.sev-picker button').forEach(button => button.classList.remove('selected'));
}

async function addSegment() {
  if (!requireAuth('add segments')) return;
  const themeCode = document.getElementById('f_theme_code').value;
  const theme = THEME_INDEX[themeCode];
  const indicatorCode = document.getElementById('f_indicator').value;
  const indicator = INDICATOR_INDEX[indicatorCode];
  const requestedId = document.getElementById('f_segid').value.trim();
  const primaryQuadrant = document.getElementById('f_domain').value.trim();
  const secondaryQuadrant = document.getElementById('f_domain_secondary').value;
  const snippet = document.getElementById('f_snippet').value.trim();
  const severity = parseInt(document.getElementById('f_severity').value, 10);
  const segmentRecord = normalizeSegmentRecord({
    Segment_ID: requestedId || ensureUniqueSegmentId(`SEG_${Date.now()}`),
    Snippet: snippet,
    Theme: theme?.label || '',
    Theme_Code: themeCode,
    CIPQ_Domain: primaryQuadrant,
    Secondary_Domain: secondaryQuadrant,
    Indicator_Code: indicatorCode,
    Indicator_Name: indicator?.label || '',
    Severity: severity,
    Stakeholder: document.getElementById('f_stakeholder').value,
    Respondent_Type: document.getElementById('f_respondent_type').value.trim(),
    Region: document.getElementById('f_region').value,
    Source_Type: document.getElementById('f_source_type').value,
    Source_ID: document.getElementById('f_source_id').value.trim(),
    Value_Chain_Stage: document.getElementById('f_value_chain_stage').value,
    PESTLE_Tags: getMultiSelectValues('f_pestle_tags'),
    Scoring_Confidence: document.getElementById('f_confidence').value,
    Is_Cross_Quadrant: !!secondaryQuadrant,
    Linked_Quadrants: [primaryQuadrant, secondaryQuadrant].filter(Boolean),
    Analysis_Notes: document.getElementById('f_analysis_notes').value.trim(),
    Session_ID: document.getElementById('f_session').value.trim(),
    Created_At: new Date().toISOString(),
    Updated_At: new Date().toISOString()
  });

  const issues = validateRecord(segmentRecord, true);
  if (issues.length) {
    showStatus(`Could not add segment: ${issues[0]}`, true);
    return;
  }

  const replacementPlan = planSegmentReplacements([segmentRecord]);
  const shouldReplace = await requestOverwriteConfirmation(replacementPlan.duplicateSummaries, 'segment');
  if (!shouldReplace) {
    showStatus('Segment was not changed.', false);
    return;
  }

  if (supabaseClient && currentUser) {
    setCloudSyncing(true);
    const deleteResult = await deleteReplacedSegmentsFromSupabase(replacementPlan.replacedExisting);
    if (deleteResult?.error) {
      setCloudSyncing(false);
      showStatus(`Could not replace existing segment: ${formatSupabaseError(deleteResult.error)}`, true);
      return;
    }

    const { data, error } = await insertSegmentsToSupabase(segmentRecord, { selectSingle: true });
    setCloudSyncing(false);

    if (error) {
      showStatus(`Could not save segment to Supabase: ${formatSupabaseError(error)}`, true);
      return;
    }

    const savedRecord = mapDbRowToSegment(data);
    dataset = dataset
      .filter(record => !replacementPlan.replacedExisting.includes(record))
      .concat(savedRecord);
    showStatus(`Segment ${segmentRecord.Segment_ID} saved to Supabase. Total: ${dataset.length}`, false);
  } else {
    dataset = replacementPlan.finalDataset;
    showStatus(`Segment ${segmentRecord.Segment_ID} added locally. Sign in to keep it after refresh.`, false);
  }

  clearForm();
  refreshAll();
}

function updateUploadMeta(filename = 'No file uploaded yet', success = false) {
  const filenameNode = document.getElementById('uploadFilename');
  const pill = document.getElementById('uploadSuccessPill');
  const zone = document.getElementById('uploadZone');
  if (filenameNode) filenameNode.textContent = filename || 'No file uploaded yet';
  if (pill) pill.classList.toggle('show', !!success);
  if (zone) zone.classList.toggle('uploaded', !!success);
}

function handleFileUpload(event) {
  if (!requireAuth('import segments')) {
    event.target.value = '';
    return;
  }
  if (event.target.files[0]) {
    updateUploadMeta(event.target.files[0].name, false);
    parseCSV(event.target.files[0]);
  }
}

function parseCSV(file) {
  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: async results => {
      const rows = results.data || [];
      const importedSegments = [];
      let skippedSurveyRows = 0;
      const takenIds = new Set(dataset.map(record => record.Segment_ID));

      rows.forEach((row, index) => {
        const record = normalizeSegmentRecord(row);
        if (!record.CIPQ_Domain) record.CIPQ_Domain = inferDomainFromCode(record.Theme_Code || record.Indicator_Code);
        if (!record.Theme_Code) record.Theme_Code = themeCodeForIndicator(record.Indicator_Code);
        if (!record.Theme) record.Theme = getThemeLabel(record.Theme_Code);
        if (!record.Indicator_Name) record.Indicator_Name = getIndicatorLabel(record.Indicator_Code);
        if (!record.Source_ID) record.Source_ID = row.Session_ID || row.session_id || '';
        if (!record.Segment_ID) record.Segment_ID = ensureUniqueSegmentId(`IMP_${Date.now()}_${index + 1}`, takenIds);
        else takenIds.add(record.Segment_ID);
        if (isSurveySourceType(record.Source_Type)) {
          skippedSurveyRows += 1;
          return;
        }
        if (!VALID_DOMAINS.includes(record.CIPQ_Domain) || !record.Indicator_Code || !Number.isFinite(record.Severity)) return;
        importedSegments.push(record);
      });

      if (!importedSegments.length) {
        showStatus('No valid rows were found in the CSV import.', true);
        updateUploadMeta(file.name, false);
        return;
      }

      const replacementPlan = planSegmentReplacements(importedSegments);
      const shouldReplace = await requestOverwriteConfirmation(replacementPlan.duplicateSummaries, 'CSV segment');
      if (!shouldReplace) {
        showStatus('CSV import canceled. Existing segments were not changed.', false);
        updateUploadMeta(file.name, false);
        return;
      }

      if (supabaseClient && currentUser) {
        setCloudSyncing(true);
        let saveError = null;
        const deleteResult = await deleteReplacedSegmentsFromSupabase(replacementPlan.replacedExisting);
        if (deleteResult?.error) {
          saveError = deleteResult.error;
        }
        if (!saveError) {
          for (const batch of chunkItems(replacementPlan.recordsToSave, 200)) {
            const result = await insertSegmentsToSupabase(batch);
            if (result.error) {
              saveError = result.error;
              break;
            }
          }
        }
        setCloudSyncing(false);

        if (saveError) {
          showStatus(`CSV import failed to save to Supabase: ${formatSupabaseError(saveError)}`, true);
          updateUploadMeta(file.name, false);
          return;
        }

        await loadSegmentsFromSupabase({ silent: true });
      } else {
        dataset = replacementPlan.finalDataset;
      }

      updateUploadMeta(file.name, true);
      const syncLabel = supabaseClient && currentUser ? ' and saved to Supabase' : ' locally';
      const replaceLabel = replacementPlan.duplicateSummaries.length ? ` Replaced ${replacementPlan.duplicateSummaries.length} duplicate${replacementPlan.duplicateSummaries.length !== 1 ? 's' : ''}.` : '';
      const skippedLabel = skippedSurveyRows ? ` Skipped ${skippedSurveyRows} survey row${skippedSurveyRows !== 1 ? 's' : ''}; import those in Survey Data.` : '';
      showStatus(`Imported ${replacementPlan.recordsToSave.length} segments from CSV${syncLabel}. Total: ${dataset.length}.${replaceLabel}${skippedLabel}`, false);
      refreshAll();
    },
    error: () => showStatus('Error parsing CSV file.', true)
  });
}

function downloadBlob(filename, content, type) {
  const blob = new Blob([content], { type });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function exportCSV() {
  if (!dataset.length) {
    showStatus('No data to export.', true);
    return;
  }
  const csv = Papa.unparse(dataset.map(record => ({
    Segment_ID: record.Segment_ID,
    Snippet: record.Snippet,
    Theme: record.Theme || '',
    Theme_Code: record.Theme_Code || '',
    Open_Code: record.Open_Code || '',
    CIPQ_Domain: record.CIPQ_Domain,
    Secondary_Domain: record.Secondary_Domain || '',
    Indicator_Code: record.Indicator_Code,
    Indicator_Name: record.Indicator_Name,
    Severity: record.Severity,
    Scoring_Confidence: record.Scoring_Confidence,
    Stakeholder: record.Stakeholder || '',
    Respondent_Type: record.Respondent_Type || '',
    Region: record.Region || '',
    Source_Type: record.Source_Type || '',
    Source_ID: record.Source_ID || '',
    Value_Chain_Stage: record.Value_Chain_Stage || '',
    PESTLE_Tags: (record.PESTLE_Tags || []).join('|'),
    Is_Cross_Quadrant: record.Is_Cross_Quadrant,
    Linked_Quadrants: (record.Linked_Quadrants || []).join('|'),
    Analysis_Notes: record.Analysis_Notes || '',
    Session_ID: record.Session_ID || '',
    Encoded_At: record.Created_At || '',
    Updated_At: record.Updated_At || ''
  })));
  downloadBlob(`CIPQ_Dataset_${new Date().toISOString().slice(0, 10)}.csv`, csv, 'text/csv;charset=utf-8');
}

function reportDateStamp() {
  return new Date().toISOString().slice(0, 10);
}

function aggregateToSummaryRow(section, item, extra = {}) {
  return {
    Section: section,
    Key: item.key || item.quadrant || item.indicator || item.label || '',
    Label: item.label || item.quadrant || item.indicator || item.key || '',
    Frequency: item.frequency ?? '',
    Average_Severity: Number.isFinite(item.average_severity) ? roundTo(item.average_severity, 2) : '',
    Weighted_Score: Number.isFinite(item.weighted_score) ? roundTo(item.weighted_score, 2) : '',
    Stakeholder_Spread: item.stakeholder_spread ?? '',
    Regional_Spread: item.regional_spread ?? '',
    Source_Spread: item.source_spread ?? '',
    Severity_Sum: item.severity_sum ?? '',
    Max_Severity: item.max_severity ?? '',
    Min_Severity: item.min_severity ?? '',
    Confidence: item.confidence_label || '',
    Top_Indicators: (item.top_indicators || []).join(' | '),
    Notes: item.interpretation || item.summary_text || item.narrative || '',
    ...extra
  };
}

function buildSummaryTableRows(layer) {
  const sourceTypeCounts = mostCommon(dataset.map(record => record.Source_Type || 'Unspecified'), 20);
  const optionalValueChainCount = dataset.filter(record => record.Value_Chain_Stage).length;
  const valueChainAnalysisCount = dataset.filter(record => valueChainStageForAnalysis(record)).length;
  const optionalPestleCount = dataset.filter(record => record.PESTLE_Tags?.length).length;
  const severityValues = dataset.map(record => record.Severity).filter(Number.isFinite);
  const rows = [
    {
      Section: 'Dataset Overview',
      Key: 'total_segments',
      Label: 'Total coded segments',
      Frequency: dataset.length,
      Average_Severity: roundTo(mean(severityValues), 2),
      Weighted_Score: '',
      Stakeholder_Spread: countDistinct(dataset.map(record => record.Stakeholder)),
      Regional_Spread: countDistinct(dataset.map(record => record.Region)),
      Source_Spread: countDistinct(dataset.map(record => record.Source_Type)),
      Severity_Sum: sum(severityValues),
      Max_Severity: severityValues.length ? Math.max(...severityValues) : '',
      Min_Severity: severityValues.length ? Math.min(...severityValues) : '',
      Confidence: layer.structural_pressure_label,
      Top_Indicators: layer.priority_signals.slice(0, 5).map(signal => signal.indicator).join(' | '),
      Notes: `Active indicators: ${layer.aggregates.indicator_stats.length}; value chain analyzed: ${valueChainAnalysisCount}; explicit value chain tagged: ${optionalValueChainCount}; PESTLE tagged: ${optionalPestleCount}`
    },
    ...sourceTypeCounts.map(item => ({
      Section: 'Source Type Totals',
      Key: item.label,
      Label: item.label,
      Frequency: item.count,
      Average_Severity: '',
      Weighted_Score: '',
      Stakeholder_Spread: '',
      Regional_Spread: '',
      Source_Spread: '',
      Severity_Sum: '',
      Max_Severity: '',
      Min_Severity: '',
      Confidence: '',
      Top_Indicators: '',
      Notes: ''
    })),
    ...layer.aggregates.quadrant_stats.map(item => aggregateToSummaryRow('Quadrant Summary', item)),
    ...layer.aggregates.indicator_stats.map(item => aggregateToSummaryRow('Indicator Summary', item)),
    ...layer.aggregates.stakeholder_stats.map(item => aggregateToSummaryRow('Stakeholder Summary', item)),
    ...layer.aggregates.region_stats.map(item => aggregateToSummaryRow('Region Summary', item)),
    ...layer.aggregates.source_stats.map(item => aggregateToSummaryRow('Source Summary', item)),
    ...layer.context_summaries.value_chain.map(item => aggregateToSummaryRow('Value Chain Summary', item)),
    ...layer.context_summaries.pestle.map(item => aggregateToSummaryRow('PESTLE Summary', item)),
    ...layer.priority_signals.map(signal => aggregateToSummaryRow('Priority Signal', signal, {
      Key: signal.indicator_code,
      Label: signal.indicator,
      Notes: `${signal.classification}: ${signal.narrative}`
    })),
    {
      Section: 'Chart Explanation',
      Key: 'quadrant_frequency_chart',
      Label: 'Quadrant Frequency',
      Frequency: '',
      Average_Severity: '',
      Weighted_Score: '',
      Stakeholder_Spread: '',
      Regional_Spread: '',
      Source_Spread: '',
      Severity_Sum: '',
      Max_Severity: '',
      Min_Severity: '',
      Confidence: layer.chart_explanations.quadrant_frequency_chart.confidence_label,
      Top_Indicators: '',
      Notes: layer.chart_explanations.quadrant_frequency_chart.text
    },
    {
      Section: 'Chart Explanation',
      Key: 'indicator_severity_chart',
      Label: 'Indicator Severity',
      Frequency: '',
      Average_Severity: '',
      Weighted_Score: '',
      Stakeholder_Spread: '',
      Regional_Spread: '',
      Source_Spread: '',
      Severity_Sum: '',
      Max_Severity: '',
      Min_Severity: '',
      Confidence: layer.chart_explanations.indicator_severity_chart.confidence_label,
      Top_Indicators: '',
      Notes: layer.chart_explanations.indicator_severity_chart.text
    },
    {
      Section: 'Chart Explanation',
      Key: 'stakeholder_comparison_chart',
      Label: 'Stakeholder Comparison',
      Frequency: '',
      Average_Severity: '',
      Weighted_Score: '',
      Stakeholder_Spread: '',
      Regional_Spread: '',
      Source_Spread: '',
      Severity_Sum: '',
      Max_Severity: '',
      Min_Severity: '',
      Confidence: layer.chart_explanations.stakeholder_comparison_chart.confidence_label,
      Top_Indicators: '',
      Notes: layer.chart_explanations.stakeholder_comparison_chart.text
    }
  ];

  return rows;
}

function exportSummaryTablesCsv() {
  const layer = buildInterpretiveLayer();
  if (!layer) {
    showStatus('No data to export yet.', true);
    return;
  }

  const csv = Papa.unparse(buildSummaryTableRows(layer));
  downloadBlob(`CIPQ_Summary_Tables_${reportDateStamp()}.csv`, csv, 'text/csv;charset=utf-8');
}

function downloadImportTemplate() {
  const fields = [
    'Segment_ID',
    'Snippet',
    'Theme_Code',
    'Indicator_Code',
    'Severity',
    'Stakeholder',
    'Respondent_Type',
    'Region',
    'Source_Type',
    'Source_ID',
    'Session_ID',
    'Value_Chain_Stage',
    'PESTLE_Tags',
    'Secondary_Domain',
    'Scoring_Confidence',
    'Analysis_Notes'
  ];
  const csv = Papa.unparse({ fields, data: [] });
  downloadBlob(`CIPQ_Import_Template_${reportDateStamp()}.csv`, csv, 'text/csv;charset=utf-8');
}

function chartNumber(value, digits = 1) {
  if (!Number.isFinite(value)) return '0';
  return Number.isInteger(value) ? String(value) : String(roundTo(value, digits));
}

function recordCountLabel(count) {
  return `${count} record${count === 1 ? '' : 's'}`;
}

function truncateLabel(value, maxLength = 32) {
  const text = String(value || 'Unspecified');
  return text.length > maxLength ? `${text.slice(0, maxLength - 3)}...` : text;
}

function domainColor(domain) {
  return DOMAIN_COLORS[domain] || DOMAIN_COLORS.Unspecified;
}

function heatColor(value, maxValue) {
  if (!maxValue) return '#f5f0e8';
  const ratio = Math.max(0.12, Math.min(1, value / maxValue));
  const start = { r: 245, g: 240, b: 232 };
  const end = { r: 201, g: 74, b: 46 };
  const r = Math.round(start.r + (end.r - start.r) * ratio);
  const g = Math.round(start.g + (end.g - start.g) * ratio);
  const b = Math.round(start.b + (end.b - start.b) * ratio);
  return `rgb(${r}, ${g}, ${b})`;
}

function chartLegend(items) {
  return `<div class="chart-legend">${items.map(item => `
    <span class="legend-item"><span class="legend-swatch" style="background:${escapeHtml(item.color)}"></span>${escapeHtml(item.label)}</span>
  `).join('')}</div>`;
}

function chartHelp(reading, meaning) {
  if (!reading && !meaning) return '';
  return `<div class="chart-help">
    ${reading ? `<p><strong>How to read this:</strong> ${escapeHtml(reading)}</p>` : ''}
    ${meaning ? `<p><strong>What it means:</strong> ${escapeHtml(meaning)}</p>` : ''}
  </div>`;
}

function panel(title, note, chartHtml, className = '', reading = '', meaning = '') {
  return `<section class="chart-panel ${escapeHtml(className)}">
    <h3>${escapeHtml(title)}</h3>
    <div class="chart-note">${escapeHtml(note)}</div>
    ${chartHelp(reading, meaning)}
    ${chartHtml}
  </section>`;
}

function chartGuidance(layer) {
  const topQuadrant = layer.aggregates.quadrant_stats[0];
  const topStakeholder = layer.aggregates.stakeholder_stats[0];
  const topSignal = layer.priority_signals[0];
  const topValueChain = [...layer.context_summaries.value_chain].sort((a, b) => b.frequency - a.frequency || b.average_severity - a.average_severity)[0];
  const topPestle = [...layer.context_summaries.pestle].sort((a, b) => b.frequency - a.frequency || b.average_severity - a.average_severity)[0];
  const criticalLocalized = layer.aggregates.indicator_stats.find(item => item.frequency <= 2 && item.average_severity >= 4);
  const widespread = layer.aggregates.indicator_stats.find(item => item.frequency >= 3);

  return {
    quadrant: {
      read: 'Long colored bars show weighted severity, which combines frequency and seriousness. The smaller gold bars show raw frequency only.',
      meaning: topQuadrant
        ? `${topQuadrant.label} currently carries the strongest quadrant-level pressure, with ${topQuadrant.frequency} coded segment${topQuadrant.frequency !== 1 ? 's' : ''} and an average severity of ${chartNumber(topQuadrant.average_severity, 2)}.`
        : 'No quadrant pattern is available yet.'
    },
    stakeholder: {
      read: 'Each row shows a stakeholder group, its dominant CIPQ domain, average severity, and number of coded records.',
      meaning: topStakeholder
        ? `${topStakeholder.label} contributes the largest stakeholder cluster in the current dataset, so its dominant domain should be read as an important stakeholder-specific pressure signal.`
        : 'Stakeholder comparison needs stakeholder metadata in the encoded records.'
    },
    heatmap: {
      read: 'Darker cells and longer bars indicate higher weighted severity. Weighted severity rises when an indicator is both frequent and severe.',
      meaning: topSignal
        ? `${topSignal.indicator} is the highest-ranked priority signal, combining frequency ${topSignal.frequency} with average severity ${chartNumber(topSignal.average_severity, 2)}.`
        : 'No priority signal is available yet.'
    },
    valueChain: {
      read: 'The larger blocks mark stages where more coded pressures accumulate. The value printed inside each block is the record count. If a record has no explicit Value Chain Stage, the chart infers one from its CIPQ domain.',
      meaning: topValueChain?.frequency
        ? `${topValueChain.label} has the strongest value-chain accumulation, helping locate where policy discussion may need to focus operational attention.`
        : 'No value-chain pattern is available yet.'
    },
    pestle: {
      read: 'Bars count how often each PESTLE tag appears. Average severity beside each bar shows whether that context is mild, moderate, or urgent.',
      meaning: topPestle?.frequency
        ? `${topPestle.label} is the most common PESTLE context in the tagged records, framing the macro-level setting around the core CIPQ findings.`
        : 'PESTLE is optional, so this view remains empty until records are tagged.'
    },
    scatter: {
      read: 'Points farther right are more widespread. Points higher up are more severe. Upper-right points are broad and serious; upper-left points are serious but localized.',
      meaning: criticalLocalized
        ? `${criticalLocalized.label} appears as a critical localized pressure, while ${widespread?.label || 'the more frequent indicators'} show wider recurrence across the dataset.`
        : widespread
          ? `${widespread.label} is a more widespread pressure; the scatterplot helps compare it against less frequent but potentially severe issues.`
          : 'The scatterplot will become more useful as more indicators accumulate repeated records.'
    }
  };
}

function renderQuadrantDistributionChart(layer) {
  const guidance = chartGuidance(layer).quadrant;
  const stats = VALID_DOMAINS.map(domain => {
    const stat = layer.aggregates.quadrant_stats.find(item => item.key === domain);
    return stat || {
      key: domain,
      label: domain,
      frequency: 0,
      weighted_score: 0,
      average_severity: 0
    };
  });
  const maxWeighted = Math.max(...stats.map(item => item.weighted_score), 1);
  const maxFrequency = Math.max(...stats.map(item => item.frequency), 1);
  const rowHeight = 62;
  const height = 60 + stats.length * rowHeight;
  let svg = `<svg class="chart-svg" role="img" aria-label="Quadrant distribution chart" viewBox="0 0 720 ${height}">
    <rect width="720" height="${height}" fill="${CHART_COLORS.paper}"/>
    <text x="148" y="28" class="axis-label">weighted severity</text>
    <text x="570" y="28" class="axis-label">frequency</text>`;
  stats.forEach((item, index) => {
    const y = 52 + index * rowHeight;
    const weightedWidth = Math.round((item.weighted_score / maxWeighted) * 360);
    const frequencyWidth = Math.round((item.frequency / maxFrequency) * 110);
    svg += `
      <text x="24" y="${y + 19}" font-size="14" font-weight="700">${escapeHtml(item.label)}</text>
      <rect x="148" y="${y}" width="360" height="24" rx="4" fill="${CHART_COLORS.cream}"/>
      <rect x="148" y="${y}" width="${weightedWidth}" height="24" rx="4" fill="${domainColor(item.key)}"/>
      <rect x="570" y="${y + 4}" width="110" height="16" rx="3" fill="${CHART_COLORS.cream}"/>
      <rect x="570" y="${y + 4}" width="${frequencyWidth}" height="16" rx="3" fill="${CHART_COLORS.gold}"/>
      <text x="520" y="${y + 18}" class="value-label">${escapeHtml(chartNumber(item.weighted_score, 1))}</text>
      <text x="688" y="${y + 18}" class="value-label" text-anchor="end">${escapeHtml(item.frequency)}</text>
      <text x="148" y="${y + 43}" class="tiny-label">Avg severity ${escapeHtml(chartNumber(item.average_severity, 2))}</text>`;
  });
  svg += '</svg>';
  svg += chartLegend([
    { label: 'Weighted severity', color: CHART_COLORS.rust },
    { label: 'Frequency', color: CHART_COLORS.gold }
  ]);
  return panel('Quadrant Distribution', 'Frequency and weighted severity by CIPQ domain', svg, '', guidance.read, guidance.meaning);
}

function renderStakeholderComparisonChart(layer) {
  const guidance = chartGuidance(layer).stakeholder;
  const stakeholderStats = layer.aggregates.stakeholder_stats.slice(0, 10);
  if (!stakeholderStats.length) {
    return panel('Stakeholder Comparison', 'Dominant domain and average severity', '<div class="no-data-msg">No stakeholder metadata available yet.</div>', '', guidance.read, guidance.meaning);
  }
  const rowHeight = 54;
  const height = 54 + stakeholderStats.length * rowHeight;
  let svg = `<svg class="chart-svg" role="img" aria-label="Stakeholder comparison chart" viewBox="0 0 900 ${height}">
    <rect width="900" height="${height}" fill="${CHART_COLORS.paper}"/>
    <text x="260" y="28" class="axis-label">dominant domain</text>
    <text x="445" y="28" class="axis-label">average severity</text>
    <text x="835" y="28" class="axis-label" text-anchor="middle">records</text>`;
  stakeholderStats.forEach((item, index) => {
    const stakeholderRecords = item.records || dataset.filter(record => record.Stakeholder === item.key);
    const dominant = aggregateBy(stakeholderRecords, record => record.CIPQ_Domain, key => key)[0];
    const domain = dominant?.key || 'Unspecified';
    const y = 48 + index * rowHeight;
    const severityWidth = Math.round((item.average_severity / 5) * 260);
    svg += `
      <text x="24" y="${y + 19}" font-size="13" font-weight="700">${escapeHtml(truncateLabel(item.label, 28))}</text>
      <rect x="260" y="${y}" width="145" height="24" rx="12" fill="${domainColor(domain)}" opacity="0.16"/>
      <text x="332" y="${y + 17}" font-size="11" text-anchor="middle" font-weight="700" fill="${domainColor(domain)}">${escapeHtml(domain)}</text>
      <rect x="445" y="${y + 4}" width="260" height="16" rx="4" fill="${CHART_COLORS.cream}"/>
      <rect x="445" y="${y + 4}" width="${severityWidth}" height="16" rx="4" fill="${domainColor(domain)}"/>
      <text x="735" y="${y + 18}" class="value-label" text-anchor="middle">${escapeHtml(chartNumber(item.average_severity, 2))}</text>
      <text x="835" y="${y + 18}" class="value-label" text-anchor="middle">${escapeHtml(item.frequency)}</text>`;
  });
  svg += '</svg>';
  svg += chartLegend(VALID_DOMAINS.map(domain => ({ label: domain, color: domainColor(domain) })));
  return panel('Stakeholder Comparison', 'Dominant domain, average severity, and record count', svg, '', guidance.read, guidance.meaning);
}

function renderPrioritySignalHeatmap(layer) {
  const guidance = chartGuidance(layer).heatmap;
  const signals = layer.priority_signals.slice(0, 14);
  if (!signals.length) {
    return panel('Priority Signal Heatmap', 'Indicators by weighted severity', '<div class="no-data-msg">No indicator signals available yet.</div>', 'wide', guidance.read, guidance.meaning);
  }
  const maxWeighted = Math.max(...signals.map(item => item.weighted_score), 1);
  const rowHeight = 42;
  const height = 60 + signals.length * rowHeight;
  let svg = `<svg class="chart-svg" role="img" aria-label="Priority signal heatmap" viewBox="0 0 900 ${height}">
    <rect width="900" height="${height}" fill="${CHART_COLORS.paper}"/>
    <text x="24" y="30" class="axis-label">indicator</text>
    <text x="470" y="30" class="axis-label">weighted severity heat</text>
    <text x="650" y="30" class="axis-label">frequency</text>
    <text x="745" y="30" class="axis-label">avg severity</text>`;
  signals.forEach((signal, index) => {
    const y = 48 + index * rowHeight;
    const tileWidth = Math.max(20, Math.round((signal.weighted_score / maxWeighted) * 170));
    svg += `
      <text x="24" y="${y + 18}" font-size="12" font-weight="700">${escapeHtml(truncateLabel(signal.indicator, 48))}</text>
      <rect x="470" y="${y}" width="180" height="24" rx="4" fill="${CHART_COLORS.cream}"/>
      <rect x="470" y="${y}" width="${tileWidth}" height="24" rx="4" fill="${heatColor(signal.weighted_score, maxWeighted)}"/>
      <text x="560" y="${y + 17}" class="value-label" text-anchor="middle">${escapeHtml(chartNumber(signal.weighted_score, 1))}</text>
      <text x="682" y="${y + 17}" class="value-label" text-anchor="middle">${escapeHtml(signal.frequency)}</text>
      <text x="790" y="${y + 17}" class="value-label" text-anchor="middle">${escapeHtml(chartNumber(signal.average_severity, 2))}</text>
      <circle cx="850" cy="${y + 12}" r="6" fill="${domainColor(signal.domain)}"/>`;
  });
  svg += '</svg>';
  return panel('Priority Signal Heatmap', 'Top indicators ranked by weighted severity', svg, 'wide', guidance.read, guidance.meaning);
}

function renderValueChainFlowChart(layer) {
  const guidance = chartGuidance(layer).valueChain;
  const items = layer.context_summaries.value_chain;
  const maxFrequency = Math.max(...items.map(item => item.frequency), 1);
  let svg = `<svg class="chart-svg" role="img" aria-label="Value chain flow summary" viewBox="0 0 920 310">
    <defs>
      <marker id="flowArrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto" markerUnits="strokeWidth">
        <path d="M0,0 L0,6 L9,3 z" fill="${CHART_COLORS.border}"></path>
      </marker>
    </defs>
    <rect width="920" height="310" fill="${CHART_COLORS.paper}"/>`;
  for (let index = 0; index < items.length - 1; index += 1) {
    const x1 = 54 + index * 220 + 176;
    const x2 = 54 + (index + 1) * 220 - 16;
    svg += `<line x1="${x1}" y1="78" x2="${x2}" y2="78" stroke="${CHART_COLORS.border}" stroke-width="4" stroke-linecap="round" marker-end="url(#flowArrow)"/>`;
  }
  items.forEach((item, index) => {
    const x = 54 + index * 220;
    const y = 104;
    const pressureWidth = Math.round((item.frequency / maxFrequency) * 150);
    const active = item.frequency > 0;
    const fill = active ? CHART_COLORS.teal : '#dbe5e2';
    const textColor = active ? 'white' : CHART_COLORS.ink;
    const mutedTextColor = active ? 'white' : CHART_COLORS.muted;
    svg += `
      <rect x="${x}" y="${y}" width="172" height="120" rx="10" fill="${fill}"/>
      <text x="${x + 86}" y="${y + 30}" font-size="14" font-weight="700" text-anchor="middle" fill="${textColor}">${escapeHtml(item.label)}</text>
      <text x="${x + 86}" y="${y + 57}" font-size="12" font-family="IBM Plex Mono, monospace" text-anchor="middle" fill="${mutedTextColor}">${escapeHtml(recordCountLabel(item.frequency))}</text>
      <text x="${x + 86}" y="${y + 78}" font-size="12" font-family="IBM Plex Mono, monospace" text-anchor="middle" fill="${mutedTextColor}">Avg severity ${escapeHtml(chartNumber(item.average_severity, 2))}</text>
      <rect x="${x + 12}" y="${y + 94}" width="150" height="10" rx="3" fill="${active ? 'rgba(255,255,255,0.28)' : CHART_COLORS.cream}"/>
      <rect x="${x + 12}" y="${y + 94}" width="${pressureWidth}" height="10" rx="3" fill="${active ? CHART_COLORS.gold : CHART_COLORS.border}"/>
      <text x="${x + 86}" y="${y + 144}" class="tiny-label" text-anchor="middle">${escapeHtml(item.top_indicators.slice(0, 1).join('') || 'No top indicator yet')}</text>`;
  });
  svg += `<text x="460" y="276" class="axis-label" text-anchor="middle">Flow order follows the book value chain. Gold bars show relative pressure by record count.</text>`;
  svg += '</svg>';
  return panel('Value Chain Flow Summary', 'Where pressures accumulate across the book value chain', svg, '', guidance.read, guidance.meaning);
}

function renderPestleDistributionChart(layer) {
  const guidance = chartGuidance(layer).pestle;
  const items = layer.context_summaries.pestle;
  const maxFrequency = Math.max(...items.map(item => item.frequency), 1);
  const rowHeight = 42;
  const height = 46 + items.length * rowHeight;
  let svg = `<svg class="chart-svg" role="img" aria-label="PESTLE distribution chart" viewBox="0 0 680 ${height}">
    <rect width="680" height="${height}" fill="${CHART_COLORS.paper}"/>`;
  items.forEach((item, index) => {
    const y = 34 + index * rowHeight;
    const width = Math.round((item.frequency / maxFrequency) * 400);
    svg += `
      <text x="24" y="${y + 17}" font-size="13" font-weight="700">${escapeHtml(item.label)}</text>
      <rect x="170" y="${y}" width="400" height="22" rx="4" fill="${CHART_COLORS.cream}"/>
      <rect x="170" y="${y}" width="${width}" height="22" rx="4" fill="${CHART_COLORS.gold}"/>
      <text x="590" y="${y + 16}" class="value-label">${escapeHtml(item.frequency)}</text>
      <text x="630" y="${y + 16}" class="tiny-label">sev ${escapeHtml(chartNumber(item.average_severity, 2))}</text>`;
  });
  svg += '</svg>';
  return panel('PESTLE Distribution', 'Macro-context tag frequency with average severity', svg, '', guidance.read, guidance.meaning);
}

function renderSeverityFrequencyScatter(layer) {
  const guidance = chartGuidance(layer).scatter;
  const items = layer.aggregates.indicator_stats.slice(0, 22);
  if (!items.length) {
    return panel('Severity vs Frequency Scatterplot', 'Widespread vs acute localized pressures', '<div class="no-data-msg">No indicators available yet.</div>', 'wide', guidance.read, guidance.meaning);
  }
  const maxFrequency = Math.max(...items.map(item => item.frequency), 1);
  const labeledItemCodes = layer.priority_signals.slice(0, 8).map(signal => signal.indicator_code);
  const labeledItems = new Set(labeledItemCodes);
  const plot = { x: 70, y: 34, width: 700, height: 300 };
  let svg = `<svg class="chart-svg" role="img" aria-label="Severity versus frequency scatterplot" viewBox="0 0 860 390">
    <rect width="860" height="390" fill="${CHART_COLORS.paper}"/>
    <line x1="${plot.x}" y1="${plot.y + plot.height}" x2="${plot.x + plot.width}" y2="${plot.y + plot.height}" stroke="${CHART_COLORS.border}" stroke-width="1.5"/>
    <line x1="${plot.x}" y1="${plot.y}" x2="${plot.x}" y2="${plot.y + plot.height}" stroke="${CHART_COLORS.border}" stroke-width="1.5"/>
    <line x1="${plot.x}" y1="${plot.y + plot.height * 0.25}" x2="${plot.x + plot.width}" y2="${plot.y + plot.height * 0.25}" stroke="${CHART_COLORS.border}" stroke-dasharray="4 6"/>
    <line x1="${plot.x + plot.width * 0.5}" y1="${plot.y}" x2="${plot.x + plot.width * 0.5}" y2="${plot.y + plot.height}" stroke="${CHART_COLORS.border}" stroke-dasharray="4 6"/>
    <text x="${plot.x + plot.width / 2}" y="372" class="axis-label" text-anchor="middle">frequency</text>
    <text x="18" y="${plot.y + plot.height / 2}" class="axis-label" transform="rotate(-90 18 ${plot.y + plot.height / 2})" text-anchor="middle">average severity</text>
    <text x="${plot.x}" y="${plot.y + plot.height + 20}" class="tiny-label">0</text>
    <text x="${plot.x + plot.width}" y="${plot.y + plot.height + 20}" class="tiny-label" text-anchor="end">${escapeHtml(maxFrequency)}</text>
    <text x="${plot.x - 12}" y="${plot.y + 5}" class="tiny-label" text-anchor="end">5</text>
    <text x="${plot.x - 12}" y="${plot.y + plot.height}" class="tiny-label" text-anchor="end">0</text>`;
  items.forEach((item, index) => {
    const x = plot.x + (item.frequency / maxFrequency) * plot.width;
    const y = plot.y + plot.height - (item.average_severity / 5) * plot.height;
    const radius = 5 + Math.min(12, item.weighted_score / Math.max(maxFrequency, 1));
    const labelIndex = labeledItemCodes.indexOf(item.key);
    svg += `
      <circle cx="${roundTo(x, 1)}" cy="${roundTo(y, 1)}" r="${roundTo(radius, 1)}" fill="${domainColor(item.domain)}" opacity="0.82">
        <title>${escapeHtml(item.label)} | Frequency ${item.frequency} | Avg severity ${chartNumber(item.average_severity, 2)}</title>
      </circle>`;
    if (labelIndex >= 0) {
      svg += `<circle cx="${roundTo(x, 1)}" cy="${roundTo(y, 1)}" r="9" fill="${CHART_COLORS.paper}" stroke="${domainColor(item.domain)}" stroke-width="2"/>
        <text x="${roundTo(x, 1)}" y="${roundTo(y + 4, 1)}" font-size="10" font-weight="700" text-anchor="middle">${labelIndex + 1}</text>`;
    }
  });
  svg += '</svg>';
  svg += chartLegend(VALID_DOMAINS.map(domain => ({ label: domain, color: domainColor(domain) })));
  svg += `<div class="chart-legend">${layer.priority_signals.slice(0, 8).map((signal, index) => `
    <span class="legend-item"><strong>${index + 1}.</strong> ${escapeHtml(truncateLabel(signal.indicator, 34))}</span>
  `).join('')}</div>`;
  return panel('Severity vs Frequency Scatterplot', 'Position indicators by recurrence and seriousness', svg, 'wide', guidance.read, guidance.meaning);
}

function renderChartVisualizations(layer) {
  return `<div class="section-title">Graph Visualizations <span>Readable exports for report writing, presentations, and policy discussion</span></div>
    <div class="chart-grid">
      ${renderQuadrantDistributionChart(layer)}
      ${renderStakeholderComparisonChart(layer)}
      ${renderPrioritySignalHeatmap(layer)}
      ${renderValueChainFlowChart(layer)}
      ${renderPestleDistributionChart(layer)}
      ${renderSeverityFrequencyScatter(layer)}
    </div>`;
}

function buildChartPackHtml(layer) {
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>CIPQ Chart Pack</title>
<style>
  body { font-family: Arial, sans-serif; color: #1a1a2e; background: #f5f0e8; padding: 24px; }
  h1 { margin: 0 0 4px; font-family: Georgia, serif; }
  .meta { color: #7a7065; margin-bottom: 24px; }
  .chart-grid { display: block; }
  .chart-panel { page-break-inside: avoid; background: white; border: 1px solid #c8bfae; border-radius: 8px; padding: 18px; margin-bottom: 22px; }
  .chart-panel h3 { margin: 0 0 4px; font-family: Georgia, serif; }
  .chart-note, .axis-label, .tiny-label, .value-label, .chart-legend { font-family: Arial, sans-serif; color: #7a7065; }
  .chart-note { font-size: 11px; text-transform: uppercase; margin-bottom: 10px; }
  .chart-svg { width: 100%; height: auto; border: 1px solid #c8bfae; border-radius: 6px; }
  .chart-legend { font-size: 11px; margin-top: 8px; }
  .legend-item { display: inline-flex; align-items: center; gap: 4px; margin-right: 12px; }
  .legend-swatch { width: 10px; height: 10px; display: inline-block; }
</style>
</head>
<body>
  <h1>CIPQ Chart Pack</h1>
  <div class="meta">Generated ${escapeHtml(new Date().toISOString())} | ${dataset.length} coded segments</div>
  ${renderChartVisualizations(layer)}
</body>
</html>`;
}

function downloadChartPack() {
  const layer = buildInterpretiveLayer();
  if (!layer) {
    showStatus('No chart data to export yet.', true);
    return;
  }

  downloadBlob(`CIPQ_Chart_Pack_${reportDateStamp()}.html`, buildChartPackHtml(layer), 'text/html;charset=utf-8');
}

function getTopExamples(records, maxItems = 3) {
  return records
    .filter(record => record.Snippet)
    .slice()
    .sort((a, b) => (b.Severity - a.Severity) || String(b.Updated_At).localeCompare(String(a.Updated_At)))
    .slice(0, maxItems)
    .map(record => record.Snippet);
}

function getTopEvidence(records, maxItems = 3) {
  return records
    .filter(record => record.Snippet)
    .slice()
    .sort((a, b) => (b.Severity - a.Severity) || String(b.Updated_At).localeCompare(String(a.Updated_At)))
    .slice(0, maxItems)
    .map(record => ({
      text_segment: record.Snippet,
      severity: record.Severity,
      stakeholder_group: record.Stakeholder,
      region: record.Region,
      source_type: record.Source_Type,
      segment_id: record.Segment_ID
    }));
}

function getConfidenceLabel(score) {
  if (score >= 0.75) return 'high';
  if (score >= 0.45) return 'moderate';
  return 'emergent';
}

function applyConfidencePrefix(confidenceLabel, sentence) {
  if (!sentence) return '';
  if (confidenceLabel === 'high') return `The data clearly indicates that ${sentence}`;
  if (confidenceLabel === 'moderate') return `The pattern suggests that ${sentence}`;
  return `An early signal appears to show that ${sentence}`;
}

function aggregateBy(records, fieldAccessor, labelAccessor) {
  const groups = new Map();
  records.forEach(record => {
    const key = fieldAccessor(record) || 'Unspecified';
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(record);
  });

  const results = [...groups.entries()].map(([key, groupRecords]) => {
    const severities = groupRecords.map(record => record.Severity).filter(Number.isFinite);
    const averageSeverity = mean(severities);
    return {
      key,
      label: typeof labelAccessor === 'function' ? (labelAccessor(key, groupRecords[0]) || key) : key,
      frequency: groupRecords.length,
      average_severity: averageSeverity,
      weighted_score: groupRecords.length * averageSeverity,
      stakeholder_spread: countDistinct(groupRecords.map(record => record.Stakeholder)),
      regional_spread: countDistinct(groupRecords.map(record => record.Region)),
      source_spread: countDistinct(groupRecords.map(record => record.Source_Type)),
      severity_sum: sum(severities),
      max_severity: severities.length ? Math.max(...severities) : 0,
      min_severity: severities.length ? Math.min(...severities) : 0,
      top_examples: getTopExamples(groupRecords, 3),
      top_indicators: mostCommon(groupRecords.map(record => record.Indicator_Name), 3).map(item => item.label),
      evidence: getTopEvidence(groupRecords, 3),
      records: groupRecords,
      domain: countDistinct(groupRecords.map(record => record.CIPQ_Domain)) === 1 ? groupRecords[0].CIPQ_Domain : ''
    };
  });

  const maxFrequency = Math.max(...results.map(result => result.frequency), 0);
  const maxStakeholderSpread = Math.max(...results.map(result => result.stakeholder_spread), 0);
  const maxRegionalSpread = Math.max(...results.map(result => result.regional_spread), 0);
  const maxSourceSpread = Math.max(...results.map(result => result.source_spread), 0);

  results.forEach(result => {
    const confidenceScore =
      normalizeValue(result.frequency, maxFrequency) * 0.4 +
      normalizeValue(result.stakeholder_spread, maxStakeholderSpread) * 0.3 +
      normalizeValue(result.regional_spread, maxRegionalSpread) * 0.2 +
      normalizeValue(result.source_spread, maxSourceSpread) * 0.1;

    result.confidence_score = roundTo(confidenceScore, 2);
    result.confidence_label = getConfidenceLabel(confidenceScore);
  });

  return results.sort((a, b) => b.weighted_score - a.weighted_score || b.frequency - a.frequency || String(a.label).localeCompare(String(b.label)));
}

function inferStructuralReading(topQuadrant, secondQuadrant) {
  if (topQuadrant === 'Production' && secondQuadrant === 'Access') {
    return 'operational bottlenecks may be restricting public access to books and related cultural products.';
  }
  if (topQuadrant === 'Creation' && secondQuadrant === 'Production') {
    return 'constraints in content creation may be affecting downstream production capacity.';
  }
  if (topQuadrant === 'Distribution' && secondQuadrant === 'Access') {
    return 'circulation barriers may be shaping unequal access across markets or regions.';
  }
  return 'sectoral challenges are structurally distributed rather than isolated to one domain.';
}

function generateDashboardSummary(quadrantStats, indicatorStats) {
  if (!quadrantStats.length || !indicatorStats.length) {
    return { title: 'Policy Insights', sentences: [] };
  }

  const topQuadrant = quadrantStats[0];
  const secondQuadrant = quadrantStats[1] || quadrantStats[0];
  const topIndicator = indicatorStats[0];
  const highestSeverityIndicator = [...indicatorStats].sort((a, b) => b.average_severity - a.average_severity || b.frequency - a.frequency)[0];
  const margin = secondQuadrant && secondQuadrant.frequency ? (topQuadrant.frequency - secondQuadrant.frequency) / secondQuadrant.frequency : 1;

  const concentrationSentence = margin >= 0.15
    ? applyConfidencePrefix(topQuadrant.confidence_label, `concerns are most concentrated in ${topQuadrant.key}.`)
    : applyConfidencePrefix(topQuadrant.confidence_label, `concerns are shared mainly between ${topQuadrant.key} and ${secondQuadrant.key}.`);

  const severitySentence = highestSeverityIndicator.key !== topIndicator.key
    ? applyConfidencePrefix(highestSeverityIndicator.confidence_label, `${highestSeverityIndicator.label} appears less frequently than some other concerns, but it carries the highest average severity.`)
    : applyConfidencePrefix(topIndicator.confidence_label, `${topIndicator.label} combines high frequency and high severity, making it a major pressure point.`);

  const weightedSentence = applyConfidencePrefix(topIndicator.confidence_label, `${topIndicator.label} currently represents the strongest weighted policy pressure point once frequency and severity are considered together.`);
  const structuralSentence = applyConfidencePrefix(topQuadrant.confidence_label, inferStructuralReading(topQuadrant.key, secondQuadrant.key));

  return {
    title: 'Policy Insights',
    sentences: [concentrationSentence, severitySentence, weightedSentence, structuralSentence]
  };
}

function generateQuadrantSummaryText(quadrant, topIndicators) {
  const indicatorsText = topIndicators.length ? topIndicators.join(', ') : 'the currently coded indicators';
  if (quadrant === 'Creation') {
    return `Creation-related concerns are shaped mainly by ${indicatorsText}. This suggests pressure on authorship, content development, and the conditions that support cultural production.`;
  }
  if (quadrant === 'Production') {
    return `Production emerges as a major pressure point, driven by ${indicatorsText}. This indicates bottlenecks in turning creative work into material output.`;
  }
  if (quadrant === 'Distribution') {
    return `Distribution concerns are concentrated around ${indicatorsText}. This suggests structural barriers in circulation, logistics, or market reach.`;
  }
  return `Access-related pressures are linked to ${indicatorsText}. This indicates barriers in affordability, reach, literacy, or public availability.`;
}

function generateQuadrantCards(records, quadrantStats) {
  return VALID_DOMAINS.map(quadrant => {
    const quadrantRecords = records.filter(record => record.CIPQ_Domain === quadrant);
    const stat = quadrantStats.find(item => item.key === quadrant);
    const topIndicators = mostCommon(quadrantRecords.map(record => record.Indicator_Name), 3).map(item => item.label);
    const summaryText = quadrantRecords.length
      ? applyConfidencePrefix(stat?.confidence_label || 'emergent', generateQuadrantSummaryText(quadrant, topIndicators))
      : `No coded evidence is available yet for ${quadrant}.`;

    return {
      quadrant,
      frequency: quadrantRecords.length,
      average_severity: roundTo(mean(quadrantRecords.map(record => record.Severity).filter(Number.isFinite)), 2),
      top_indicators: topIndicators,
      summary_text: summaryText,
      evidence: getTopEvidence(quadrantRecords, 3),
      confidence_label: stat?.confidence_label || 'emergent',
      confidence_score: stat?.confidence_score || 0
    };
  });
}

function classifySignal(indicator, thresholds) {
  if (
    indicator.frequency >= thresholds.highFrequencyThreshold &&
    indicator.average_severity >= thresholds.highSeverityThreshold &&
    indicator.stakeholder_spread >= thresholds.highSpreadThreshold
  ) {
    return 'widespread_structural_issue';
  }
  if (indicator.frequency < thresholds.highFrequencyThreshold && indicator.average_severity >= 4.5) {
    return 'acute_critical_issue';
  }
  if (
    indicator.frequency >= thresholds.highFrequencyThreshold &&
    indicator.average_severity >= 3.0 &&
    indicator.stakeholder_spread < thresholds.highSpreadThreshold
  ) {
    return 'persistent_operational_issue';
  }
  return 'localized_or_emergent_issue';
}

function generateSignalNarrative(indicator, classification) {
  if (classification === 'widespread_structural_issue') {
    return `${indicator.label} appears to be systemic, given its recurrence across multiple stakeholder groups and its consistently high severity.`;
  }
  if (classification === 'acute_critical_issue') {
    return `${indicator.label} is less common than some other issues, but it becomes highly serious when it occurs.`;
  }
  if (classification === 'persistent_operational_issue') {
    return `${indicator.label} appears as a recurring operational bottleneck that may require targeted intervention.`;
  }
  return `${indicator.label} appears significant in specific contexts, but is not yet broadly distributed across the dataset.`;
}

function generatePrioritySignals(indicatorStats) {
  if (!indicatorStats.length) return [];
  const thresholds = {
    highFrequencyThreshold: percentile(indicatorStats.map(indicator => indicator.frequency), 75),
    highSeverityThreshold: 4.0,
    highSpreadThreshold: 3
  };

  return indicatorStats.map(indicator => {
    const classification = classifySignal(indicator, thresholds);
    return {
      indicator: indicator.label,
      indicator_code: indicator.key,
      domain: indicator.domain,
      frequency: indicator.frequency,
      average_severity: roundTo(indicator.average_severity, 2),
      weighted_score: roundTo(indicator.weighted_score, 2),
      stakeholder_spread: indicator.stakeholder_spread,
      classification,
      narrative: applyConfidencePrefix(indicator.confidence_label, generateSignalNarrative(indicator, classification)),
      evidence: indicator.evidence,
      confidence_label: indicator.confidence_label,
      confidence_score: indicator.confidence_score
    };
  }).sort((a, b) => b.weighted_score - a.weighted_score || b.average_severity - a.average_severity);
}

function proportionBy(records, fieldAccessor) {
  const proportions = {};
  if (!records.length) return proportions;
  records.forEach(record => {
    const key = fieldAccessor(record) || 'Unspecified';
    proportions[key] = (proportions[key] || 0) + 1;
  });
  Object.keys(proportions).forEach(key => {
    proportions[key] = proportions[key] / records.length;
  });
  return proportions;
}

function compareStakeholderToOverall(stakeholderRecords, overallRecords) {
  const stakeholderShare = proportionBy(stakeholderRecords, record => record.CIPQ_Domain);
  const overallShare = proportionBy(overallRecords, record => record.CIPQ_Domain);
  const differences = {};
  VALID_DOMAINS.forEach(quadrant => {
    differences[quadrant] = (stakeholderShare[quadrant] || 0) - (overallShare[quadrant] || 0);
  });

  let mostDistinctQuadrant = VALID_DOMAINS[0];
  let largestDifference = Math.abs(differences[mostDistinctQuadrant] || 0);
  VALID_DOMAINS.slice(1).forEach(quadrant => {
    const difference = Math.abs(differences[quadrant] || 0);
    if (difference > largestDifference) {
      largestDifference = difference;
      mostDistinctQuadrant = quadrant;
    }
  });

  if ((differences[mostDistinctQuadrant] || 0) >= 0.2) {
    return `Unlike the broader dataset, this stakeholder group places greater emphasis on ${mostDistinctQuadrant}.`;
  }
  return null;
}

function generateStakeholderInsights(records) {
  const grouped = new Map();
  records.filter(record => record.Stakeholder).forEach(record => {
    if (!grouped.has(record.Stakeholder)) grouped.set(record.Stakeholder, []);
    grouped.get(record.Stakeholder).push(record);
  });

  const insights = [];
  grouped.forEach((stakeholderRecords, stakeholder) => {
    const quadrantBreakdown = aggregateBy(stakeholderRecords, record => record.CIPQ_Domain, key => key);
    const topQuadrant = quadrantBreakdown[0];
    const topIndicators = mostCommon(stakeholderRecords.map(record => record.Indicator_Name), 3).map(item => item.label);
    const differenceNote = compareStakeholderToOverall(stakeholderRecords, records);
    const narrative = `${stakeholder} most strongly emphasizes issues in ${topQuadrant?.key || 'Unspecified'}, especially ${topIndicators.join(', ') || 'the coded indicators currently in the dataset'}.`;

    insights.push({
      stakeholder,
      top_quadrant: topQuadrant?.key || 'Unspecified',
      top_indicators: topIndicators,
      narrative: applyConfidencePrefix(topQuadrant?.confidence_label || 'emergent', narrative),
      difference_note: differenceNote,
      evidence: getTopEvidence(stakeholderRecords, 2),
      confidence_label: topQuadrant?.confidence_label || 'emergent'
    });
  });

  return insights.sort((a, b) => a.stakeholder.localeCompare(b.stakeholder));
}

function textIncludesAny(sourceText, keywords) {
  const text = String(sourceText || '').toLowerCase();
  return keywords.some(keyword => text.includes(keyword));
}

function generateCrossQuadrantReading(records) {
  const linkMap = [
    {
      from_indicator_keywords: ['printing', 'production cost', 'paper cost', 'offshore', 'print cost'],
      from_quadrant: 'Production',
      to_quadrant: 'Access',
      sentence: 'rising production costs may be contributing to affordability and access barriers.'
    },
    {
      from_indicator_keywords: ['author support', 'creative labor', 'content development', 'editorial'],
      from_quadrant: 'Creation',
      to_quadrant: 'Production',
      sentence: 'constraints in content creation may be affecting downstream production capacity.'
    },
    {
      from_indicator_keywords: ['distribution gap', 'bookstore', 'regional reach', 'logistics', 'shipping'],
      from_quadrant: 'Distribution',
      to_quadrant: 'Access',
      sentence: 'distribution barriers may be shaping unequal public access across regions or markets.'
    }
  ];

  const readings = [];
  linkMap.forEach(rule => {
    const matchingRecords = records.filter(record => {
      const searchable = [
        record.Theme,
        record.Theme_Code,
        record.Indicator_Name,
        record.Indicator_Code,
        record.Snippet,
        record.Analysis_Notes
      ].join(' ');

      const secondaryMatch = record.Secondary_Domain === rule.to_quadrant || (record.Linked_Quadrants || []).includes(rule.to_quadrant);
      return record.CIPQ_Domain === rule.from_quadrant && (textIncludesAny(searchable, rule.from_indicator_keywords) || secondaryMatch);
    });

    if (matchingRecords.length >= 3 || matchingRecords.some(record => record.Secondary_Domain === rule.to_quadrant)) {
      const aggregate = aggregateBy(matchingRecords, () => `${rule.from_quadrant}_${rule.to_quadrant}`, () => `${rule.from_quadrant} to ${rule.to_quadrant}`)[0];
      readings.push({
        sentence: applyConfidencePrefix(aggregate?.confidence_label || 'moderate', rule.sentence),
        confidence_label: aggregate?.confidence_label || 'moderate',
        evidence: getTopEvidence(matchingRecords, 3),
        matched_count: matchingRecords.length
      });
    }
  });

  if (!readings.length) {
    readings.push({
      sentence: 'The dataset suggests interdependence across quadrants rather than isolated policy failures.',
      confidence_label: 'moderate',
      evidence: getTopEvidence(records, 3),
      matched_count: 0
    });
  }

  return readings;
}

function explainFrequencyChart(stats) {
  if (!stats.length) return { text: 'Not enough data to explain the quadrant chart.', confidence_label: 'emergent' };
  const top = stats[0];
  const second = stats[1] || stats[0];
  const text = second.frequency && top.frequency >= second.frequency * 1.2
    ? `${top.key} is the most recurrent area of concern in the current dataset.`
    : 'concerns are distributed across multiple quadrants rather than dominated by one.';
  return {
    text: applyConfidencePrefix(top.confidence_label, text),
    confidence_label: top.confidence_label
  };
}

function explainSeverityChart(stats) {
  if (!stats.length) return { text: 'Not enough data to explain the indicator severity chart.', confidence_label: 'emergent' };
  const highestSeverity = [...stats].sort((a, b) => b.average_severity - a.average_severity || b.frequency - a.frequency)[0];
  const mostFrequent = [...stats].sort((a, b) => b.frequency - a.frequency || b.average_severity - a.average_severity)[0];
  const text = highestSeverity.key !== mostFrequent.key
    ? `${highestSeverity.label} is not the most frequent issue, but it has the highest average severity, suggesting concentrated urgency.`
    : `${highestSeverity.label} combines both prevalence and seriousness, making it a major policy pressure point.`;
  return {
    text: applyConfidencePrefix(highestSeverity.confidence_label, text),
    confidence_label: highestSeverity.confidence_label
  };
}

function explainComparisonChart(stats) {
  if (!stats.length) return { text: 'Not enough stakeholder data is available to explain the comparison chart.', confidence_label: 'emergent' };
  const spread = variance(stats.map(item => item.frequency));
  const average = mean(stats.map(item => item.frequency));
  const highSpread = spread > Math.max(1, average);
  const reference = [...stats].sort((a, b) => b.frequency - a.frequency)[0];
  const text = highSpread
    ? 'stakeholder groups experience the same ecosystem differently, suggesting the need for differentiated policy responses.'
    : 'stakeholder groups show broad agreement regarding the main areas of concern.';
  return {
    text: applyConfidencePrefix(reference.confidence_label, text),
    confidence_label: reference.confidence_label
  };
}

function generateChartExplanations(quadrantStats, indicatorStats, stakeholderStats) {
  return {
    quadrant_frequency_chart: explainFrequencyChart(quadrantStats),
    indicator_severity_chart: explainSeverityChart(indicatorStats),
    stakeholder_comparison_chart: explainComparisonChart(stakeholderStats)
  };
}

function fillContextStats(options, stats) {
  return options.map(option => {
    const stat = stats.find(item => item.key === option);
    return stat || {
      key: option,
      label: option,
      frequency: 0,
      average_severity: 0,
      weighted_score: 0,
      stakeholder_spread: 0,
      regional_spread: 0,
      source_spread: 0,
      severity_sum: 0,
      max_severity: 0,
      min_severity: 0,
      top_examples: [],
      top_indicators: [],
      evidence: [],
      records: [],
      domain: '',
      confidence_score: 0,
      confidence_label: 'emergent'
    };
  });
}

function generateContextInterpretation(kind, item) {
  if (!item.frequency) return `No coded evidence is available yet for ${item.label}.`;
  const topIndicators = item.top_indicators.length ? item.top_indicators.join(', ') : 'the coded indicators in this group';
  const severityText = item.average_severity >= 4
    ? 'a high-severity contextual pressure'
    : item.average_severity >= 3
      ? 'a moderate-to-high contextual pressure'
      : 'an emerging contextual signal';

  if (kind === 'value_chain') {
    return `${item.label} records cluster around ${topIndicators}. With an average severity of ${item.average_severity.toFixed(2)}, this reads as ${severityText} supporting the core CIPQ findings.`;
  }

  return `${item.label} context appears through ${topIndicators}. With an average severity of ${item.average_severity.toFixed(2)}, this tag helps frame the macro-context around the CIPQ pressure signals.`;
}

function buildContextSummaries(records) {
  const valueChainStats = fillContextStats(
    VALUE_CHAIN_STAGES,
    aggregateBy(records.filter(record => valueChainStageForAnalysis(record)), record => valueChainStageForAnalysis(record), key => key)
  );

  const pestleExpandedRecords = [];
  records.forEach(record => {
    (record.PESTLE_Tags || []).forEach(tag => {
      pestleExpandedRecords.push({ ...record, __pestleTag: tag });
    });
  });

  const pestleStats = fillContextStats(
    PESTLE_TAGS,
    aggregateBy(pestleExpandedRecords, record => record.__pestleTag, key => key)
  );

  return {
    value_chain: valueChainStats.map(item => ({
      ...item,
      interpretation: generateContextInterpretation('value_chain', item)
    })),
    pestle: pestleStats.map(item => ({
      ...item,
      interpretation: generateContextInterpretation('pestle', item)
    }))
  };
}

function serializeAggregateItem(item) {
  return {
    key: item.key,
    label: item.label,
    frequency: item.frequency,
    average_severity: roundTo(item.average_severity, 2),
    weighted_score: roundTo(item.weighted_score, 2),
    stakeholder_spread: item.stakeholder_spread,
    regional_spread: item.regional_spread,
    source_spread: item.source_spread,
    severity_sum: item.severity_sum,
    max_severity: item.max_severity,
    min_severity: item.min_severity,
    top_examples: item.top_examples,
    top_indicators: item.top_indicators,
    confidence_score: item.confidence_score,
    confidence_label: item.confidence_label,
    interpretation: item.interpretation || ''
  };
}

function buildInterpretiveExport(layer) {
  return {
    generated_at: new Date().toISOString(),
    records: dataset.map(toInterpretiveRecord),
    validation: layer.validation,
    dashboard_summary: layer.dashboard_summary,
    quadrant_cards: layer.quadrant_cards,
    priority_signals: layer.priority_signals,
    stakeholder_insights: layer.stakeholder_insights,
    cross_quadrant_reading: layer.cross_quadrant_reading,
    chart_explanations: {
      quadrant_frequency_chart: layer.chart_explanations.quadrant_frequency_chart.text,
      indicator_severity_chart: layer.chart_explanations.indicator_severity_chart.text,
      stakeholder_comparison_chart: layer.chart_explanations.stakeholder_comparison_chart.text
    },
    quadrant_summary: layer.aggregates.quadrant_stats.map(serializeAggregateItem),
    indicator_summary: layer.aggregates.indicator_stats.map(serializeAggregateItem),
    value_chain_pressure_mapping: layer.context_summaries.value_chain.map(serializeAggregateItem),
    pestle_context_summary: layer.context_summaries.pestle.map(serializeAggregateItem),
    client_view: {
      top_panel: layer.dashboard_summary,
      main_charts: ['quadrant_frequency_chart', 'indicator_severity_chart', 'stakeholder_comparison_chart'],
      chart_explanations: layer.chart_explanations,
      quadrant_cards: layer.quadrant_cards,
      graph_visualizations: [
        'quadrant_distribution',
        'stakeholder_comparison',
        'priority_signal_heatmap',
        'value_chain_flow_summary',
        'pestle_distribution',
        'severity_frequency_scatterplot'
      ],
      value_chain_pressure_mapping: true,
      pestle_context_summary: true,
      cross_quadrant_reading: layer.cross_quadrant_reading,
      priority_signals: layer.priority_signals,
      stakeholder_insights: layer.stakeholder_insights,
      evidence_popups: true,
      download_report_button: true
    },
    encoder_view: {
      data_entry_form: true,
      record_editor: true,
      severity_assignment: true,
      theme_code_manager: true,
      quadrant_assignment: true,
      raw_dataset_table: true,
      client_interpretation_panel: 'optional_preview_only'
    }
  };
}

function buildInterpretiveLayer() {
  if (!dataset.length) return null;

  const validationIssues = dataset
    .map(record => ({ id: record.Segment_ID, issues: validateRecord(record, false) }))
    .filter(item => item.issues.length);

  const quadrantStats = aggregateBy(dataset, record => record.CIPQ_Domain, key => key);
  const indicatorStats = aggregateBy(dataset, record => record.Indicator_Code, (key, sample) => sample.Indicator_Name || getIndicatorLabel(key) || key);
  const stakeholderStats = aggregateBy(dataset.filter(record => record.Stakeholder), record => record.Stakeholder, key => key);
  const regionStats = aggregateBy(dataset.filter(record => record.Region), record => record.Region, key => key);
  const sourceStats = aggregateBy(dataset.filter(record => record.Source_Type), record => record.Source_Type, key => key);

  const dashboardSummary = generateDashboardSummary(quadrantStats, indicatorStats);
  const quadrantCards = generateQuadrantCards(dataset, quadrantStats);
  const prioritySignals = generatePrioritySignals(indicatorStats);
  const stakeholderInsights = generateStakeholderInsights(dataset);
  const crossQuadrantReading = generateCrossQuadrantReading(dataset);
  const chartExplanations = generateChartExplanations(quadrantStats, indicatorStats, stakeholderStats);
  const contextSummaries = buildContextSummaries(dataset);
  const cipqIndex = roundTo(mean(quadrantStats.map(item => item.average_severity).filter(value => value > 0)), 2);
  const structuralPressureLabel = cipqIndex >= 4 ? 'Critical' : cipqIndex >= 3 ? 'High' : cipqIndex >= 2 ? 'Moderate' : 'Emergent';

  return {
    validation: {
      issue_count: validationIssues.length,
      issues: validationIssues
    },
    aggregates: {
      quadrant_stats: quadrantStats,
      indicator_stats: indicatorStats,
      stakeholder_stats: stakeholderStats,
      region_stats: regionStats,
      source_stats: sourceStats
    },
    dashboard_summary: dashboardSummary,
    quadrant_cards: quadrantCards,
    priority_signals: prioritySignals,
    stakeholder_insights: stakeholderInsights,
    cross_quadrant_reading: crossQuadrantReading,
    context_summaries: contextSummaries,
    chart_explanations: chartExplanations,
    cipq_index: cipqIndex,
    structural_pressure_label: structuralPressureLabel
  };
}

function exportInterpretiveJson() {
  const layer = buildInterpretiveLayer();
  if (!layer) {
    showStatus('No data to export yet.', true);
    return;
  }
  const payload = buildInterpretiveExport(layer);
  downloadBlob(`CIPQ_Interpretive_Layer_${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(payload, null, 2), 'application/json;charset=utf-8');
}

function buildClientReportText(layer) {
  const sourceTypeCounts = mostCommon(dataset.map(record => record.Source_Type || 'Unspecified'), 20);
  const optionalValueChainCount = dataset.filter(record => record.Value_Chain_Stage).length;
  const valueChainAnalysisCount = dataset.filter(record => valueChainStageForAnalysis(record)).length;
  const optionalPestleCount = dataset.filter(record => record.PESTLE_Tags?.length).length;
  const severityValues = dataset.map(record => record.Severity).filter(Number.isFinite);

  const lines = [];
  lines.push('CIPQ Client Report');
  lines.push(`Generated: ${new Date().toISOString()}`);
  lines.push('');
  lines.push('Dataset Totals');
  lines.push(`- Total coded segments: ${dataset.length}`);
  lines.push(`- Average severity: ${roundTo(mean(severityValues), 2)}`);
  lines.push(`- Severity range: ${severityValues.length ? `${Math.min(...severityValues)} to ${Math.max(...severityValues)}` : 'No severity values'}`);
  lines.push(`- Active indicators: ${layer.aggregates.indicator_stats.length}`);
  lines.push(`- Stakeholder groups: ${countDistinct(dataset.map(record => record.Stakeholder))}`);
  lines.push(`- Regions: ${countDistinct(dataset.map(record => record.Region))}`);
  lines.push(`- Value Chain analyzed records: ${valueChainAnalysisCount}`);
  lines.push(`- Explicit Value Chain tagged records: ${optionalValueChainCount}`);
  lines.push(`- PESTLE tagged records: ${optionalPestleCount}`);
  if (sourceTypeCounts.length) {
    lines.push(`- Source totals: ${sourceTypeCounts.map(item => `${item.label} ${item.count}`).join('; ')}`);
  }
  lines.push('');
  lines.push('Policy Insights');
  layer.dashboard_summary.sentences.forEach(sentence => lines.push(`- ${sentence}`));
  lines.push('');
  lines.push('Quadrant Summary Table');
  lines.push('Quadrant | Frequency | Avg Severity | Weighted Score | Stakeholder Spread | Top Indicators');
  layer.aggregates.quadrant_stats.forEach(item => {
    lines.push(`${item.label} | ${item.frequency} | ${roundTo(item.average_severity, 2)} | ${roundTo(item.weighted_score, 2)} | ${item.stakeholder_spread} | ${item.top_indicators.join(', ') || '-'}`);
  });
  lines.push('');
  lines.push('Indicator Summary Table');
  lines.push('Indicator Code | Indicator | Frequency | Avg Severity | Weighted Score | Stakeholder Spread | Confidence');
  layer.aggregates.indicator_stats.forEach(item => {
    lines.push(`${item.key} | ${item.label} | ${item.frequency} | ${roundTo(item.average_severity, 2)} | ${roundTo(item.weighted_score, 2)} | ${item.stakeholder_spread} | ${item.confidence_label}`);
  });
  lines.push('');
  lines.push('Quadrant Summary Cards');
  layer.quadrant_cards.forEach(card => {
    lines.push(`${card.quadrant}: ${card.summary_text}`);
    lines.push(`  Frequency: ${card.frequency} | Avg Severity: ${card.average_severity} | Confidence: ${card.confidence_label}`);
    if (card.top_indicators.length) lines.push(`  Top indicators: ${card.top_indicators.join(', ')}`);
  });
  lines.push('');
  lines.push('Priority Policy Signals');
  layer.priority_signals.slice(0, 10).forEach(signal => {
    lines.push(`- ${signal.indicator} (${signal.classification})`);
    lines.push(`  ${signal.narrative}`);
  });
  lines.push('');
  lines.push('Value Chain Pressure Mapping');
  layer.context_summaries.value_chain.forEach(item => {
    lines.push(`- ${item.label}: Frequency ${item.frequency} | Avg Severity ${roundTo(item.average_severity, 2)}`);
    if (item.top_indicators.length) lines.push(`  Top indicators: ${item.top_indicators.join(', ')}`);
    lines.push(`  ${item.interpretation}`);
  });
  lines.push('');
  lines.push('PESTLE Context Summary');
  layer.context_summaries.pestle.forEach(item => {
    lines.push(`- ${item.label}: Frequency ${item.frequency} | Avg Severity ${roundTo(item.average_severity, 2)}`);
    if (item.top_indicators.length) lines.push(`  Top indicators: ${item.top_indicators.join(', ')}`);
    lines.push(`  ${item.interpretation}`);
  });
  lines.push('');
  lines.push('Stakeholder Perspectives');
  layer.stakeholder_insights.forEach(insight => {
    lines.push(`- ${insight.stakeholder}: ${insight.narrative}`);
    if (insight.difference_note) lines.push(`  ${insight.difference_note}`);
  });
  lines.push('');
  lines.push('Cross-Quadrant Reading');
  layer.cross_quadrant_reading.forEach(reading => lines.push(`- ${reading.sentence}`));

  return lines.join('\r\n');
}

function wordCell(value, style = '') {
  return `<td style="border:1px solid #c8bfae;padding:6px;vertical-align:top;${style}">${escapeHtml(value ?? '')}</td>`;
}

function wordHeader(labels) {
  return `<tr>${labels.map(label => `<th style="border:1px solid #c8bfae;padding:6px;background:#ede8dc;text-align:left">${escapeHtml(label)}</th>`).join('')}</tr>`;
}

function wordExplainBlock(reading, meaning) {
  return `<p style="background:#f5f0e8;border:1px solid #c8bfae;padding:8px;margin:6px 0 10px">
    <strong>How to read this:</strong> ${escapeHtml(reading)}<br>
    <strong>What it means:</strong> ${escapeHtml(meaning)}
  </p>`;
}

function wordBar(value, maxValue, color) {
  const width = maxValue ? Math.max(3, Math.round((value / maxValue) * 100)) : 0;
  return `<div style="width:100%;background:#ede8dc;height:14px;border-radius:3px">
    <div style="width:${width}%;background:${color};height:14px;border-radius:3px"></div>
  </div>`;
}

function pressureCategory(item) {
  if (item.frequency >= 3 && item.average_severity >= 4) return 'Widespread high severity';
  if (item.frequency < 3 && item.average_severity >= 4) return 'Critical localized';
  if (item.frequency >= 3) return 'Widespread moderate';
  return 'Emergent/localized';
}

function buildWordVisualTables(layer) {
  const guidance = chartGuidance(layer);
  const quadrantItems = VALID_DOMAINS.map(domain => {
    const stat = layer.aggregates.quadrant_stats.find(item => item.key === domain);
    return stat || { key: domain, label: domain, frequency: 0, weighted_score: 0, average_severity: 0, stakeholder_spread: 0, top_indicators: [] };
  });
  const maxQuadrantWeighted = Math.max(...quadrantItems.map(item => item.weighted_score), 1);
  const maxQuadrantFrequency = Math.max(...quadrantItems.map(item => item.frequency), 1);
  const maxPestleFrequency = Math.max(...layer.context_summaries.pestle.map(item => item.frequency), 1);
  const maxValueChainFrequency = Math.max(...layer.context_summaries.value_chain.map(item => item.frequency), 1);
  const maxSignalWeighted = Math.max(...layer.priority_signals.map(item => item.weighted_score), 1);

  let html = '<h2>Word-Safe Visual Summary Tables</h2><p>The following tables repeat the graph findings in a Word-safe format so the report remains readable even if a Word version does not preserve SVG charts perfectly.</p>';

  html += `<h3>Quadrant Distribution</h3>${wordExplainBlock(guidance.quadrant.read, guidance.quadrant.meaning)}<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:16px">`;
  html += wordHeader(['Quadrant', 'Weighted Severity', 'Frequency', 'Avg Severity']);
  quadrantItems.forEach(item => {
    html += `<tr>${wordCell(item.label)}<td style="border:1px solid #c8bfae;padding:6px">${wordBar(item.weighted_score, maxQuadrantWeighted, domainColor(item.key))}<div>${escapeHtml(chartNumber(item.weighted_score, 1))}</div></td><td style="border:1px solid #c8bfae;padding:6px">${wordBar(item.frequency, maxQuadrantFrequency, CHART_COLORS.gold)}<div>${escapeHtml(item.frequency)}</div></td>${wordCell(chartNumber(item.average_severity, 2))}</tr>`;
  });
  html += '</table>';

  html += `<h3>Stakeholder Comparison</h3>${wordExplainBlock(guidance.stakeholder.read, guidance.stakeholder.meaning)}<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:16px">`;
  html += wordHeader(['Stakeholder', 'Dominant Domain', 'Avg Severity', 'Frequency']);
  layer.aggregates.stakeholder_stats.slice(0, 12).forEach(item => {
    const dominant = aggregateBy(item.records || [], record => record.CIPQ_Domain, key => key)[0];
    html += `<tr>${wordCell(item.label)}${wordCell(dominant?.key || 'Unspecified')}${wordCell(chartNumber(item.average_severity, 2))}${wordCell(item.frequency)}</tr>`;
  });
  html += '</table>';

  html += `<h3>Priority Signal Heatmap</h3>${wordExplainBlock(guidance.heatmap.read, guidance.heatmap.meaning)}<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:16px">`;
  html += wordHeader(['Indicator', 'Weighted Severity', 'Frequency', 'Avg Severity', 'Classification']);
  layer.priority_signals.slice(0, 15).forEach(signal => {
    html += `<tr>${wordCell(signal.indicator)}<td style="border:1px solid #c8bfae;padding:6px;background:${heatColor(signal.weighted_score, maxSignalWeighted)}">${escapeHtml(chartNumber(signal.weighted_score, 1))}</td>${wordCell(signal.frequency)}${wordCell(chartNumber(signal.average_severity, 2))}${wordCell(signal.classification.replace(/_/g, ' '))}</tr>`;
  });
  html += '</table>';

  html += `<h3>Value Chain Flow Summary</h3>${wordExplainBlock(guidance.valueChain.read, guidance.valueChain.meaning)}<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:16px">`;
  html += wordHeader(layer.context_summaries.value_chain.map(item => item.label));
  html += '<tr>';
  layer.context_summaries.value_chain.forEach(item => {
    html += `<td style="border:1px solid #c8bfae;padding:8px;vertical-align:top">${wordBar(item.frequency, maxValueChainFrequency, CHART_COLORS.teal)}<div><strong>${escapeHtml(recordCountLabel(item.frequency))}</strong></div><div>Avg severity ${escapeHtml(chartNumber(item.average_severity, 2))}</div><div>${escapeHtml(item.top_indicators.join(', ') || 'No indicators yet')}</div></td>`;
  });
  html += '</tr></table>';

  html += `<h3>PESTLE Distribution</h3>${wordExplainBlock(guidance.pestle.read, guidance.pestle.meaning)}<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:16px">`;
  html += wordHeader(['Tag', 'Frequency', 'Avg Severity', 'Bar']);
  layer.context_summaries.pestle.forEach(item => {
    html += `<tr>${wordCell(item.label)}${wordCell(item.frequency)}${wordCell(chartNumber(item.average_severity, 2))}<td style="border:1px solid #c8bfae;padding:6px">${wordBar(item.frequency, maxPestleFrequency, CHART_COLORS.gold)}</td></tr>`;
  });
  html += '</table>';

  html += `<h3>Severity vs Frequency Scatterplot Data</h3>${wordExplainBlock(guidance.scatter.read, guidance.scatter.meaning)}<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:16px">`;
  html += wordHeader(['Indicator', 'Frequency', 'Avg Severity', 'Policy Reading']);
  layer.aggregates.indicator_stats.slice(0, 18).forEach(item => {
    html += `<tr>${wordCell(item.label)}${wordCell(item.frequency)}${wordCell(chartNumber(item.average_severity, 2))}${wordCell(pressureCategory(item))}</tr>`;
  });
  html += '</table>';

  return html;
}

function buildReportNarrativeHtml(layer) {
  const severityValues = dataset.map(record => record.Severity).filter(Number.isFinite);
  const sourceTypeCounts = mostCommon(dataset.map(record => record.Source_Type || 'Unspecified'), 20);
  const sourceTotals = sourceTypeCounts.length
    ? sourceTypeCounts.map(item => `${item.label} ${item.count}`).join('; ')
    : 'No source metadata available yet';
  const valueChainAnalysisCount = dataset.filter(record => valueChainStageForAnalysis(record)).length;
  const explicitValueChainCount = dataset.filter(record => record.Value_Chain_Stage).length;

  let html = '<h2>Dataset Totals</h2>';
  html += `<p>The current CIPQ backbone contains <strong>${escapeHtml(dataset.length)}</strong> coded segment${dataset.length !== 1 ? 's' : ''}. Average severity is <strong>${escapeHtml(roundTo(mean(severityValues), 2))}</strong>${severityValues.length ? `, with a severity range of <strong>${escapeHtml(Math.min(...severityValues))} to ${escapeHtml(Math.max(...severityValues))}</strong>` : ''}. The dataset currently covers <strong>${escapeHtml(layer.aggregates.indicator_stats.length)}</strong> active indicators, <strong>${escapeHtml(countDistinct(dataset.map(record => record.Stakeholder)))}</strong> stakeholder group${countDistinct(dataset.map(record => record.Stakeholder)) !== 1 ? 's' : ''}, and <strong>${escapeHtml(countDistinct(dataset.map(record => record.Region)))}</strong> region${countDistinct(dataset.map(record => record.Region)) !== 1 ? 's' : ''}.</p>`;
  html += `<p><strong>Value Chain note:</strong> ${escapeHtml(valueChainAnalysisCount)} records are included in Value Chain analysis. ${escapeHtml(explicitValueChainCount)} have an explicit Value Chain Stage; records without that optional field are inferred from their CIPQ domain for reporting continuity.</p>`;
  html += `<p><strong>Source totals:</strong> ${escapeHtml(sourceTotals)}.</p>`;

  html += '<h2>Quadrant Summary</h2>';
  layer.quadrant_cards.forEach(card => {
    html += `<h3>${escapeHtml(card.quadrant)}</h3>`;
    html += `<p>${escapeHtml(card.summary_text)}</p>`;
    html += `<p><strong>Frequency:</strong> ${escapeHtml(card.frequency)} | <strong>Average severity:</strong> ${escapeHtml(chartNumber(card.average_severity, 2))} | <strong>Confidence:</strong> ${escapeHtml(card.confidence_label)}. ${card.top_indicators.length ? `<strong>Top indicators:</strong> ${escapeHtml(card.top_indicators.join(', '))}.` : ''}</p>`;
  });

  html += '<h2>Priority Policy Signals</h2>';
  layer.priority_signals.slice(0, 10).forEach(signal => {
    html += `<h3>${escapeHtml(signal.indicator)}</h3>`;
    html += `<p>${escapeHtml(signal.narrative)}</p>`;
    html += `<p><strong>Frequency:</strong> ${escapeHtml(signal.frequency)} | <strong>Average severity:</strong> ${escapeHtml(chartNumber(signal.average_severity, 2))} | <strong>Weighted score:</strong> ${escapeHtml(chartNumber(signal.weighted_score, 2))} | <strong>Classification:</strong> ${escapeHtml(signal.classification.replace(/_/g, ' '))}.</p>`;
  });

  html += '<h2>Value Chain Pressure Mapping</h2>';
  layer.context_summaries.value_chain.forEach(item => {
    html += `<h3>${escapeHtml(item.label)}</h3>`;
    html += `<p>${escapeHtml(item.interpretation)}</p>`;
  });

  html += '<h2>PESTLE Context Summary</h2>';
  layer.context_summaries.pestle.forEach(item => {
    html += `<h3>${escapeHtml(item.label)}</h3>`;
    html += `<p>${escapeHtml(item.interpretation)}</p>`;
  });

  html += '<h2>Stakeholder Perspectives</h2>';
  if (layer.stakeholder_insights.length) {
    layer.stakeholder_insights.forEach(insight => {
      html += `<h3>${escapeHtml(insight.stakeholder)}</h3>`;
      html += `<p>${escapeHtml(insight.narrative)}</p>`;
      if (insight.difference_note) html += `<p>${escapeHtml(insight.difference_note)}</p>`;
    });
  } else {
    html += '<p>No stakeholder comparison is available yet because stakeholder metadata has not been encoded.</p>';
  }

  html += '<h2>Cross-Quadrant Reading</h2>';
  layer.cross_quadrant_reading.forEach(reading => {
    html += `<p>${escapeHtml(reading.sentence)}</p>`;
  });

  return html;
}

function buildClientReportWordHtml(layer) {
  const summaryRows = buildSummaryTableRows(layer);
  const summaryTableRows = summaryRows.slice(0, 80).map(row => `<tr>
    ${wordCell(row.Section)}
    ${wordCell(row.Label)}
    ${wordCell(row.Frequency)}
    ${wordCell(row.Average_Severity)}
    ${wordCell(row.Weighted_Score)}
    ${wordCell(row.Stakeholder_Spread)}
    ${wordCell(row.Notes)}
  </tr>`).join('');
  const insights = layer.dashboard_summary.sentences.map(sentence => `<li>${escapeHtml(sentence)}</li>`).join('');

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>CIPQ Client Report</title>
<style>
  body { font-family: Arial, sans-serif; color: #1a1a2e; line-height: 1.45; }
  h1, h2, h3 { font-family: Georgia, serif; color: #1a1a2e; }
  h1 { font-size: 24pt; margin-bottom: 4px; }
  h2 { font-size: 16pt; margin-top: 22px; border-bottom: 1px solid #c8bfae; padding-bottom: 4px; }
  h3 { font-size: 12pt; margin-top: 16px; }
  .meta { color: #7a7065; margin-bottom: 18px; }
  table { font-size: 9.5pt; }
  li { margin-bottom: 5px; }
  .section-title { font-family: Georgia, serif; font-size: 16pt; margin-top: 22px; border-bottom: 1px solid #c8bfae; padding-bottom: 4px; }
  .section-title span { display: block; font-family: Arial, sans-serif; font-size: 9pt; color: #7a7065; font-weight: normal; }
  .chart-grid { display: block; }
  .chart-panel { page-break-inside: avoid; border: 1px solid #c8bfae; padding: 12px; margin-bottom: 16px; background: #fffdf8; }
  .chart-note { color: #7a7065; font-size: 8.5pt; text-transform: uppercase; margin-bottom: 8px; }
  .chart-help { background: #f5f0e8; border: 1px solid #c8bfae; padding: 8px; margin-bottom: 10px; }
  .chart-help p { margin: 0 0 4px; }
  .chart-svg { width: 100%; height: auto; border: 1px solid #c8bfae; }
  .chart-legend { font-size: 8.5pt; color: #7a7065; margin-top: 6px; }
  .legend-item { margin-right: 10px; }
  .legend-swatch { width: 10px; height: 10px; display: inline-block; }
</style>
</head>
<body>
  <h1>CIPQ Client Report</h1>
  <div class="meta">Generated ${escapeHtml(new Date().toISOString())} | ${dataset.length} coded segments</div>
  <h2>Policy Insights</h2>
  <ul>${insights}</ul>
  ${buildReportNarrativeHtml(layer)}
  <h2>Graph Visualizations</h2>
  <p>These charts are included for report writing and stakeholder presentation. If a Word version does not preserve the SVG graphics exactly, the Word-safe visual summary tables below contain the same findings in table form.</p>
  ${renderChartVisualizations(layer)}
  ${buildWordVisualTables(layer)}
  <h2>Summary Tables</h2>
  <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%">
    ${wordHeader(['Section', 'Label', 'Frequency', 'Avg Severity', 'Weighted', 'Stakeholder Spread', 'Notes'])}
    ${summaryTableRows}
  </table>
</body>
</html>`;
}

function cloneSimulationDefaults() {
  return JSON.parse(JSON.stringify(DEFAULT_SIMULATION_STATE));
}

function getSimulationBaseline(layer = null) {
  const activeLayer = layer || buildInterpretiveLayer();
  const stats = activeLayer?.aggregates?.quadrant_stats || [];
  const baseline = {};
  VALID_DOMAINS.forEach(domain => {
    const stat = stats.find(item => item.key === domain);
    baseline[domain] = roundTo(stat?.weighted_score || 0, 2);
  });
  return baseline;
}

function totalPressure(values) {
  return sum(Object.values(values).filter(Number.isFinite));
}

function balanceIndex(values) {
  const entries = Object.values(values).filter(Number.isFinite);
  if (!entries.length) return 0;
  const average = mean(entries);
  if (!average) return 100;
  const maxDeviation = Math.max(...entries.map(value => Math.abs(value - average)));
  return roundTo(Math.max(0, 100 - ((maxDeviation / average) * 100)), 1);
}

function strongestDomain(values) {
  return VALID_DOMAINS
    .map(domain => ({ domain, value: values[domain] || 0 }))
    .sort((a, b) => b.value - a.value)[0];
}

function redistributePressure(values, sourceDomain, amount) {
  const weights = SIMULATION_DEPENDENCY_MATRIX[sourceDomain] || {};
  Object.entries(weights).forEach(([targetDomain, weight]) => {
    values[targetDomain] = (values[targetDomain] || 0) + (amount * weight);
  });
}

function simulateEquilibrium(layer = null, state = simulationState) {
  const baseline = getSimulationBaseline(layer);
  const projected = {};
  let redistributedPool = 0;
  VALID_DOMAINS.forEach(domain => {
    const base = baseline[domain] || 0;
    const reliefRate = (state.relief[domain] || 0) / 100;
    const shockRate = (state.shock[domain] || 0) / 100;
    const relievedPressure = base * reliefRate;
    projected[domain] = Math.max(0, base - relievedPressure + (base * shockRate));
    redistributePressure(projected, domain, relievedPressure * ((state.redistribution || 0) / 100));
    redistributedPool += relievedPressure * ((state.redistribution || 0) / 100);
  });

  const rounds = Math.max(1, Math.min(6, parseInt(state.rounds, 10) || 1));
  const balancingRate = Math.max(0, Math.min(60, state.balancing || 0)) / 100;
  for (let round = 0; round < rounds; round += 1) {
    const average = mean(VALID_DOMAINS.map(domain => projected[domain] || 0));
    VALID_DOMAINS.forEach(domain => {
      const surplus = Math.max(0, (projected[domain] || 0) - average) * balancingRate * 0.25;
      if (!surplus) return;
      projected[domain] = Math.max(0, projected[domain] - surplus);
      redistributePressure(projected, domain, surplus);
      redistributedPool += surplus;
    });
  }

  VALID_DOMAINS.forEach(domain => {
    projected[domain] = roundTo(projected[domain] || 0, 2);
  });

  const rows = VALID_DOMAINS.map(domain => ({
    domain,
    baseline: baseline[domain] || 0,
    projected: projected[domain] || 0,
    change: roundTo((projected[domain] || 0) - (baseline[domain] || 0), 2),
    relief: state.relief[domain] || 0,
    shock: state.shock[domain] || 0
  }));

  return {
    baseline,
    projected,
    rows,
    total_baseline: roundTo(totalPressure(baseline), 2),
    total_projected: roundTo(totalPressure(projected), 2),
    balance_baseline: balanceIndex(baseline),
    balance_projected: balanceIndex(projected),
    strongest_baseline: strongestDomain(baseline),
    strongest_projected: strongestDomain(projected),
    redistributed_pressure: roundTo(redistributedPool, 2),
    state
  };
}

function simulationConsequenceText(result) {
  if (!result.total_baseline) return 'The simulator needs coded data before it can establish a baseline pressure state.';
  const totalChange = roundTo(result.total_projected - result.total_baseline, 2);
  const balanceChange = roundTo(result.balance_projected - result.balance_baseline, 1);
  const direction = totalChange < 0 ? 'reduces' : totalChange > 0 ? 'increases' : 'keeps';
  const balanceDirection = balanceChange > 0 ? 'more balanced' : balanceChange < 0 ? 'less balanced' : 'similarly balanced';
  return `This scenario ${direction} total modeled pressure by ${Math.abs(totalChange).toFixed(2)} weighted points and leaves the ecosystem ${balanceDirection}. The dominant projected pressure is ${result.strongest_projected.domain}, suggesting that policy consequences should be checked most carefully in that domain.`;
}

function simulationRedistributionText(result) {
  const rising = result.rows.filter(row => row.change > 0).sort((a, b) => b.change - a.change);
  const falling = result.rows.filter(row => row.change < 0).sort((a, b) => a.change - b.change);
  if (!rising.length && !falling.length) return 'No redistribution is visible yet because all scenario controls are at baseline.';
  const riseText = rising.length ? `pressure rises in ${rising.map(row => `${row.domain} (+${row.change.toFixed(2)})`).join(', ')}` : 'no domain absorbs additional pressure';
  const fallText = falling.length ? `pressure falls in ${falling.map(row => `${row.domain} (${row.change.toFixed(2)})`).join(', ')}` : 'no domain shows pressure relief';
  return `Under this scenario, ${fallText}, while ${riseText}. This is the redistribution pattern to discuss before treating an intervention as system-wide relief.`;
}

function renderSimulationChart(result) {
  const maxValue = Math.max(...result.rows.flatMap(row => [row.baseline, row.projected]), 1);
  const rowHeight = 70;
  const height = 64 + result.rows.length * rowHeight;
  let svg = `<svg class="chart-svg" role="img" aria-label="Scenario equilibrium simulation" viewBox="0 0 860 ${height}">
    <rect width="860" height="${height}" fill="${CHART_COLORS.paper}"/>
    <text x="190" y="28" class="axis-label">baseline pressure</text>
    <text x="510" y="28" class="axis-label">projected pressure</text>
    <text x="780" y="28" class="axis-label" text-anchor="middle">change</text>`;
  result.rows.forEach((row, index) => {
    const y = 52 + index * rowHeight;
    const baselineWidth = Math.round((row.baseline / maxValue) * 240);
    const projectedWidth = Math.round((row.projected / maxValue) * 240);
    const changeColor = row.change > 0 ? CHART_COLORS.rust : row.change < 0 ? CHART_COLORS.teal : CHART_COLORS.muted;
    svg += `
      <text x="24" y="${y + 20}" font-size="14" font-weight="700">${escapeHtml(row.domain)}</text>
      <rect x="190" y="${y}" width="240" height="18" rx="4" fill="${CHART_COLORS.cream}"/>
      <rect x="190" y="${y}" width="${baselineWidth}" height="18" rx="4" fill="${domainColor(row.domain)}" opacity="0.45"/>
      <text x="444" y="${y + 15}" class="value-label">${escapeHtml(chartNumber(row.baseline, 2))}</text>
      <rect x="510" y="${y}" width="240" height="18" rx="4" fill="${CHART_COLORS.cream}"/>
      <rect x="510" y="${y}" width="${projectedWidth}" height="18" rx="4" fill="${domainColor(row.domain)}"/>
      <text x="764" y="${y + 15}" class="value-label">${escapeHtml(chartNumber(row.projected, 2))}</text>
      <text x="820" y="${y + 15}" class="value-label" text-anchor="end" fill="${changeColor}">${row.change > 0 ? '+' : ''}${escapeHtml(chartNumber(row.change, 2))}</text>
      <text x="190" y="${y + 42}" class="tiny-label">support ${escapeHtml(row.relief)}% | shock ${escapeHtml(row.shock)}%</text>`;
  });
  svg += '</svg>';
  svg += chartLegend([
    { label: 'Pale bars: current baseline', color: CHART_COLORS.border },
    { label: 'Solid bars: simulated equilibrium', color: CHART_COLORS.teal }
  ]);
  return svg;
}

function simulationControlRow(domain) {
  return `<div class="sim-control">
    <div class="sim-control-head">
      <strong>${escapeHtml(domain)}</strong>
      <span>Support ${escapeHtml(simulationState.relief[domain])}% | Shock ${escapeHtml(simulationState.shock[domain])}%</span>
    </div>
    <div class="sim-slider-row">
      <label>Support</label>
      <input type="range" min="0" max="70" value="${escapeHtml(simulationState.relief[domain])}" oninput="updateSimulationControl('relief','${domain}',this.value)">
      <span>${escapeHtml(simulationState.relief[domain])}%</span>
    </div>
    <div class="sim-slider-row">
      <label>Shock</label>
      <input type="range" min="0" max="60" value="${escapeHtml(simulationState.shock[domain])}" oninput="updateSimulationControl('shock','${domain}',this.value)">
      <span>${escapeHtml(simulationState.shock[domain])}%</span>
    </div>
  </div>`;
}

function renderSimulator() {
  const mount = document.getElementById('simulatorContent');
  if (!mount) return;
  const layer = buildInterpretiveLayer();
  if (!layer) {
    mount.innerHTML = '<div class="no-data-msg">No data yet. Encode or import segments first, then run scenario simulations.</div>';
    return;
  }

  const result = simulateEquilibrium(layer, simulationState);
  let html = `<div class="sim-grid">
    <section class="sim-panel">
      <h3>Scenario Controls</h3>
      <div class="sim-note">Support reduces direct pressure in a domain. Shock increases pressure. Redistribution estimates how much relieved pressure reappears elsewhere in the ecosystem.</div>
      ${VALID_DOMAINS.map(simulationControlRow).join('')}
      <div class="sim-control">
        <div class="sim-control-head"><strong>System Coupling</strong><span>${escapeHtml(simulationState.redistribution)}%</span></div>
        <div class="sim-slider-row">
          <label>Redistrib.</label>
          <input type="range" min="0" max="80" value="${escapeHtml(simulationState.redistribution)}" oninput="updateSimulationControl('system','redistribution',this.value)">
          <span>${escapeHtml(simulationState.redistribution)}%</span>
        </div>
        <div class="sim-slider-row">
          <label>Balancing</label>
          <input type="range" min="0" max="60" value="${escapeHtml(simulationState.balancing)}" oninput="updateSimulationControl('system','balancing',this.value)">
          <span>${escapeHtml(simulationState.balancing)}%</span>
        </div>
        <div class="sim-slider-row">
          <label>Rounds</label>
          <input type="range" min="1" max="6" value="${escapeHtml(simulationState.rounds)}" oninput="updateSimulationControl('system','rounds',this.value)">
          <span>${escapeHtml(simulationState.rounds)}</span>
        </div>
      </div>
    </section>
    <section class="sim-panel">
      <h3>Projected Dynamic Equilibrium</h3>
      <div class="sim-note">This is a transparent heuristic model for scenario planning. It should support discussion, not replace causal policy evaluation.</div>
      <div class="sim-metric-grid">
        <div class="sim-metric"><div class="label">Total Pressure</div><span class="value">${escapeHtml(result.total_projected)}</span></div>
        <div class="sim-metric"><div class="label">Balance Index</div><span class="value">${escapeHtml(result.balance_projected)}%</span></div>
        <div class="sim-metric"><div class="label">Dominant Domain</div><span class="value">${escapeHtml(result.strongest_projected.domain)}</span></div>
        <div class="sim-metric"><div class="label">Redistributed</div><span class="value">${escapeHtml(result.redistributed_pressure)}</span></div>
      </div>
      ${renderSimulationChart(result)}
      <div class="sim-warning"><strong>Policy consequence reading:</strong> ${escapeHtml(simulationConsequenceText(result))}<br><br><strong>Redistribution reading:</strong> ${escapeHtml(simulationRedistributionText(result))}</div>
      <div class="table-wrap" style="margin-top:1rem">
        <table class="sim-table">
          <thead><tr><th>Domain</th><th>Baseline</th><th>Projected</th><th>Change</th><th>Support</th><th>Shock</th></tr></thead>
          <tbody>
            ${result.rows.map(row => `<tr>
              <td><span class="tag ${DOMAIN_CLASS[row.domain] || ''}">${escapeHtml(row.domain)}</span></td>
              <td>${escapeHtml(chartNumber(row.baseline, 2))}</td>
              <td>${escapeHtml(chartNumber(row.projected, 2))}</td>
              <td>${row.change > 0 ? '+' : ''}${escapeHtml(chartNumber(row.change, 2))}</td>
              <td>${escapeHtml(row.relief)}%</td>
              <td>${escapeHtml(row.shock)}%</td>
            </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </section>
  </div>`;
  mount.innerHTML = html;
}

function updateSimulationControl(kind, key, value) {
  const numeric = parseInt(value, 10) || 0;
  if (kind === 'relief' || kind === 'shock') {
    simulationState[kind][key] = numeric;
  } else {
    simulationState[key] = numeric;
  }
  renderSimulator();
}

function resetSimulation() {
  simulationState = cloneSimulationDefaults();
  renderSimulator();
}

function buildSimulationReportText() {
  const layer = buildInterpretiveLayer();
  if (!layer) return '';
  const result = simulateEquilibrium(layer, simulationState);
  const lines = [];
  lines.push('CIPQ Scenario-Based Equilibrium Simulation');
  lines.push(`Generated: ${new Date().toISOString()}`);
  lines.push('');
  lines.push('Model Note');
  lines.push('This is a heuristic scenario simulator. It estimates pressure redistribution and ecosystem balance from coded CIPQ weighted pressure; it is not a causal forecast.');
  lines.push('');
  lines.push('Scenario Controls');
  VALID_DOMAINS.forEach(domain => {
    lines.push(`- ${domain}: support ${simulationState.relief[domain]}%, shock ${simulationState.shock[domain]}%`);
  });
  lines.push(`- Redistribution sensitivity: ${simulationState.redistribution}%`);
  lines.push(`- Balancing strength: ${simulationState.balancing}%`);
  lines.push(`- Adjustment rounds: ${simulationState.rounds}`);
  lines.push('');
  lines.push('Projected Results');
  lines.push(`- Total pressure: ${result.total_baseline} -> ${result.total_projected}`);
  lines.push(`- Balance index: ${result.balance_baseline}% -> ${result.balance_projected}%`);
  lines.push(`- Dominant pressure: ${result.strongest_baseline.domain} -> ${result.strongest_projected.domain}`);
  lines.push(`- Redistributed pressure: ${result.redistributed_pressure}`);
  lines.push('');
  lines.push('Domain Table');
  result.rows.forEach(row => {
    lines.push(`- ${row.domain}: baseline ${row.baseline}, projected ${row.projected}, change ${row.change > 0 ? '+' : ''}${row.change}`);
  });
  lines.push('');
  lines.push(simulationConsequenceText(result));
  lines.push(simulationRedistributionText(result));
  return lines.join('\r\n');
}

function downloadSimulationReport() {
  const text = buildSimulationReportText();
  if (!text) {
    showStatus('No simulation data to export yet.', true);
    return;
  }
  downloadBlob(`CIPQ_Scenario_Simulation_${reportDateStamp()}.txt`, text, 'text/plain;charset=utf-8');
}

function downloadClientReport() {
  const layer = buildInterpretiveLayer();
  if (!layer) {
    showStatus('No data to export yet.', true);
    return;
  }

  downloadBlob(`CIPQ_Client_Report_${reportDateStamp()}.doc`, `\ufeff${buildClientReportWordHtml(layer)}`, 'application/msword;charset=utf-8');
}

async function copyTextToClipboard(text) {
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

async function copyClientReport() {
  const layer = buildInterpretiveLayer();
  if (!layer) {
    showStatus('No report to copy yet.', true);
    return;
  }

  try {
    await copyTextToClipboard(buildClientReportText(layer));
    showStatus('Client report copied to clipboard.', false);
  } catch (error) {
    showStatus(`Could not copy report: ${error.message || error}`, true);
  }
}

function renderConfidencePill(label) {
  return `<span class="confidence-pill ${escapeHtml(label)}">${escapeHtml(label)}</span>`;
}

function renderPressureLegend(activeLabel) {
  const items = [
    { label: 'Emergent', range: 'below 2.00', color: '#c94a2e' },
    { label: 'Moderate', range: '2.00-2.99', color: '#d9a441' },
    { label: 'High', range: '3.00-3.99', color: '#2a6b6e' },
    { label: 'Critical', range: '4.00-5.00', color: '#3a7a3a' }
  ];
  return `<div class="pressure-legend" aria-label="Structural pressure level legend">
    ${items.map(item => `
      <div>
        <span class="pressure-dot" style="background:${escapeHtml(item.color)}"></span>
        <span><strong>${escapeHtml(item.label)}${item.label === activeLabel ? ' (current)' : ''}</strong>: CIPQ index ${escapeHtml(item.range)}</span>
      </div>`).join('')}
  </div>`;
}

function sevDots(value) {
  let html = `<div class="severity-bar"><span>${escapeHtml(value)}</span><div class="sev-dots">`;
  for (let i = 1; i <= 5; i += 1) {
    html += `<div class="dot ${i <= value ? 'filled' : ''}"></div>`;
  }
  html += `</div><span>${escapeHtml(SEVERITY_LABELS[value] || '')}</span></div>`;
  return html;
}

function renderEvidenceHtml(evidence) {
  if (!evidence.length) return '<div class="muted-inline">No supporting excerpts yet.</div>';
  return `<div class="evidence-list">${evidence.map(item => `<blockquote>${escapeHtml(item.text_segment)}</blockquote>`).join('')}</div>`;
}

function renderValidationNotice(layer) {
  if (!currentUser || !layer.validation.issue_count) return '';
  const issuePreview = layer.validation.issues.slice(0, 5).map(item => `<li><strong>${escapeHtml(item.id)}</strong>: ${escapeHtml(item.issues[0])}</li>`).join('');
  return `<div class="validation-banner">
    <strong>Validation notes</strong>
    <p>${layer.validation.issue_count} record${layer.validation.issue_count !== 1 ? 's' : ''} contain schema or codebook warnings. Review the flagged records in Analyst View before final export.</p>
    <ul>${issuePreview}</ul>
  </div>`;
}

function renderAnalystValidationPanel() {
  if (!currentUser) return '';
  const layer = buildInterpretiveLayer();
  if (!dataset.length || !layer) {
    return `<div class="summary-shell">
      <h3>Validation Notes</h3>
      <p class="muted-inline">No encoded segments yet. Validation notes will appear here after records are added or imported.</p>
    </div>`;
  }

  if (!layer.validation.issue_count) {
    return `<div class="summary-shell">
      <h3>Validation Notes</h3>
      <p class="muted-inline">No validation warnings in the current encoded dataset.</p>
    </div>`;
  }

  const issueItems = layer.validation.issues.map(item => `
    <li><strong>${escapeHtml(item.id)}</strong>: ${escapeHtml(item.issues.join('; '))}</li>
  `).join('');
  return `<div class="validation-banner">
    <strong>Validation Notes</strong>
    <p>${layer.validation.issue_count} record${layer.validation.issue_count !== 1 ? 's' : ''} need analyst review before final export.</p>
    <ul>${issueItems}</ul>
  </div>`;
}

function openIndicatorTrace(code, tabName = null) {
  activeTraceCode = code;
  if (tabName) {
    const button = document.querySelector(`#mainNav button[data-tab="${tabName}"]`);
    if (button) {
      setAppView(button.dataset.view);
      switchTab(tabName, button);
      return;
    }
  }
  renderIndicators();
}

function closeIndicatorTrace() {
  activeTraceCode = null;
  renderIndicators();
}

function renderTraceability(targetId) {
  const mount = document.getElementById(targetId);
  if (!mount) return;
  if (!activeTraceCode) {
    mount.innerHTML = '';
    return;
  }

  const rows = dataset.filter(record => record.Indicator_Code === activeTraceCode);
  if (!rows.length) {
    mount.innerHTML = '<div class="trace-panel"><div class="trace-empty">No linked segments found for this indicator.</div></div>';
    return;
  }

  const sample = rows[0];
  let html = `<div class="trace-panel">
    <div class="trace-head">
      <div>
        <div class="section-title" style="font-size:1.05rem;margin-bottom:0;border-bottom:none;padding-bottom:0">Indicator Traceability <span>Original coded segments</span></div>
        <p><strong>${escapeHtml(sample.Indicator_Code)} | ${escapeHtml(sample.Indicator_Name)}</strong> | <span class="tag ${DOMAIN_CLASS[sample.CIPQ_Domain] || ''}">${escapeHtml(sample.CIPQ_Domain)}</span> | ${rows.length} linked segment${rows.length !== 1 ? 's' : ''}</p>
      </div>
      <button class="btn btn-secondary" type="button" onclick="closeIndicatorTrace()">Close</button>
    </div>
    <div class="trace-list">`;

  rows.forEach(record => {
    const secondaryTag = record.Secondary_Domain ? `<span class="tag ${DOMAIN_CLASS[record.Secondary_Domain] || ''}">Secondary: ${escapeHtml(record.Secondary_Domain)}</span>` : '';
    html += `<div class="trace-item">
      <div style="display:flex;justify-content:space-between;gap:1rem;flex-wrap:wrap;align-items:center">
        <code style="font-size:0.72rem">${escapeHtml(record.Segment_ID)}</code>
        ${sevDots(record.Severity)}
      </div>
      ${record.Theme ? `<p><strong>Theme:</strong> ${escapeHtml(record.Theme_Code || '')} | ${escapeHtml(record.Theme)}</p>` : ''}
      ${record.Analysis_Notes ? `<p><strong>Notes:</strong> ${escapeHtml(record.Analysis_Notes)}</p>` : ''}
      <div class="trace-quote">${escapeHtml(record.Snippet || '-')}</div>
      <div class="trace-meta">
        <span class="tag ${DOMAIN_CLASS[record.CIPQ_Domain] || ''}">${escapeHtml(record.CIPQ_Domain)}</span>
        ${secondaryTag}
        ${record.Stakeholder ? `<span class="tag">${escapeHtml(record.Stakeholder)}</span>` : ''}
        ${record.Region ? `<span class="tag">${escapeHtml(record.Region)}</span>` : ''}
        ${record.Value_Chain_Stage ? `<span class="tag">${escapeHtml(record.Value_Chain_Stage)}</span>` : ''}
        ${(record.PESTLE_Tags || []).map(tag => `<span class="tag">${escapeHtml(tag)}</span>`).join('')}
        ${record.Source_Type ? `<span class="tag">${escapeHtml(record.Source_Type)}</span>` : ''}
        ${record.Source_ID ? `<span class="tag">${escapeHtml(record.Source_ID)}</span>` : ''}
        ${record.Scoring_Confidence ? renderConfidencePill(record.Scoring_Confidence) : ''}
      </div>
    </div>`;
  });

  html += '</div></div>';
  mount.innerHTML = html;
}

function renderEntryPreview() {
  const recent = dataset.slice(-5).reverse();
  if (!recent.length) {
    document.getElementById('entryPreview').innerHTML = '';
    return;
  }

  let html = `<div class="section-title" style="font-size:1rem;margin-top:0.5rem;">Recent Entries <span>Last ${recent.length}</span></div>
    <div class="table-wrap">
      <table class="dataset-table">
        <colgroup>
          <col style="width:130px">
          <col style="width:120px">
          <col style="width:200px">
          <col style="width:100px">
          <col style="width:90px">
          <col style="width:120px">
          <col style="width:110px">
          <col style="width:80px">
        </colgroup>
        <thead>
          <tr>
            <th>Segment ID</th>
            <th>Quadrant</th>
            <th>Indicator</th>
            <th>Severity</th>
            <th>Stakeholder</th>
            <th>Value Chain</th>
            <th>Source</th>
            <th></th>
          </tr>
        </thead>
        <tbody>`;

  recent.forEach(record => {
    const sourceLabel = [record.Source_Type, record.Source_ID].filter(Boolean).join(' · ');
    html += `<tr>
      <td><code class="ds-segid">${escapeHtml(record.Segment_ID)}</code></td>
      <td><span class="tag ${DOMAIN_CLASS[record.CIPQ_Domain] || ''}">${escapeHtml(record.CIPQ_Domain || '-')}</span></td>
      <td>
        <span class="ds-indicator">${escapeHtml(record.Indicator_Name || record.Indicator_Code)}</span>
        <div class="ds-meta-sub">${escapeHtml(record.Indicator_Code)}</div>
      </td>
      <td>${sevDots(record.Severity)}</td>
      <td><span class="ds-cell-text">${escapeHtml(record.Stakeholder || '-')}</span></td>
      <td><span class="ds-cell-text">${escapeHtml(record.Value_Chain_Stage || '-')}</span></td>
      <td><span class="ds-cell-text">${escapeHtml(sourceLabel || '-')}</span></td>
      <td>${canWrite() ? `<button class="btn btn-secondary ds-delete-btn" type="button" onclick="deleteSegment('${record.Segment_ID}')">Delete</button>` : ''}</td>
    </tr>`;
  });

  html += '</tbody></table></div>';
  document.getElementById('entryPreview').innerHTML = html;
}

async function deleteSegment(id) {
  if (!requireAuth('delete segments')) return;
  const target = dataset.find(record => record.Segment_ID === id);
  if (!target) return;

  if (supabaseClient && currentUser && target.DB_ID) {
    setCloudSyncing(true);
    const { error } = await supabaseClient.from(SUPABASE_TABLE).delete().eq('id', target.DB_ID);
    setCloudSyncing(false);
    if (error) {
      showStatus(`Could not delete segment from Supabase: ${formatSupabaseError(error)}`, true);
      return;
    }
  }

  dataset = dataset.filter(record => record.Segment_ID !== id);
  expandedSnippetIds.delete(id);
  showStatus(supabaseClient && currentUser ? 'Segment removed from Supabase.' : 'Segment removed.', false);
  refreshAll();
}

function toggleDatasetSnippet(id) {
  if (expandedSnippetIds.has(id)) expandedSnippetIds.delete(id);
  else expandedSnippetIds.add(id);
  renderDataset();
}

async function clearAllSegments() {
  if (!requireAuth('clear all segments')) return;
  if (!dataset.length) {
    showStatus('No segments to clear.', true);
    return;
  }

  if (supabaseClient && currentUser) {
    setCloudSyncing(true);
    const { error } = await supabaseClient.from(SUPABASE_TABLE).delete().eq('user_id', currentUser.id);
    setCloudSyncing(false);
    if (error) {
      showStatus(`Could not clear Supabase history: ${formatSupabaseError(error)}`, true);
      return;
    }
  }

  dataset = [];
  expandedSnippetIds.clear();
  refreshAll();
  showStatus(supabaseClient && currentUser ? 'All saved segments cleared from Supabase.' : 'All local segments cleared.', false);
}

function buildComparisonTable(groupKey, groups, title) {
  if (!groups.length) return '';
  const indicators = [...new Set(dataset.map(record => record.Indicator_Code))].sort();
  let html = `<div class="section-title" style="font-size:1.1rem;margin-top:1.5rem">${escapeHtml(title)}</div><div class="table-wrap"><table class="comp-table"><thead><tr><th>Indicator</th>`;
  groups.forEach(group => {
    html += `<th style="text-align:center;font-size:0.68rem">${escapeHtml(group)}</th>`;
  });
  html += '</tr></thead><tbody>';

  indicators.forEach(code => {
    const indicatorName = dataset.find(record => record.Indicator_Code === code)?.Indicator_Name || code;
    html += `<tr><td><code style="font-size:0.75rem">${escapeHtml(code)}</code> ${escapeHtml(indicatorName)}</td>`;
    groups.forEach(group => {
      const rows = dataset.filter(record => record.Indicator_Code === code && record[groupKey] === group);
      if (!rows.length) {
        html += `<td class="comp-cell" style="color:var(--border)">-</td>`;
        return;
      }
      const averageSeverity = mean(rows.map(record => record.Severity));
      const cellClass = averageSeverity >= 4 ? 'high' : averageSeverity >= 3 ? 'mid' : 'low';
      html += `<td class="comp-cell ${cellClass}">${averageSeverity.toFixed(1)}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody></table></div>';
  return html;
}

function renderContextSummaryCards(items) {
  return `<div class="insight-grid">${items.map(item => `
    <article class="insight-card">
      <div class="card-head">
        <div>
          <h3>${escapeHtml(item.label)}</h3>
          <div class="meta-row">${renderConfidencePill(item.confidence_label)}</div>
        </div>
      </div>
      <div class="stat-pair">
        <div class="stat-block"><div class="label">Frequency</div><span class="value">${escapeHtml(item.frequency)}</span></div>
        <div class="stat-block"><div class="label">Avg Severity</div><span class="value">${escapeHtml(roundTo(item.average_severity, 2).toFixed(2))}</span></div>
      </div>
      <div class="muted-inline" style="margin-top:0.85rem">Top indicators: ${escapeHtml(item.top_indicators.join(', ') || 'None yet')}</div>
      <p style="margin-top:0.8rem">${escapeHtml(item.interpretation)}</p>
    </article>
  `).join('')}</div>`;
}

function renderDashboard() {
  const mount = document.getElementById('dashContent');
  if (!dataset.length) {
    mount.innerHTML = '<div class="no-data-msg">No data yet. Encode or import segments in Analyst View to generate client-facing insights.</div>';
    return;
  }

  const layer = buildInterpretiveLayer();
  const topIndicator = layer.aggregates.indicator_stats[0];
  const stakeholderCount = countDistinct(dataset.map(record => record.Stakeholder));
  const regionCount = countDistinct(dataset.map(record => record.Region));
  const sourceCount = countDistinct(dataset.map(record => record.Source_Type));
  const valueChainCount = dataset.filter(record => record.Value_Chain_Stage).length;
  const valueChainAnalysisCount = dataset.filter(record => valueChainStageForAnalysis(record)).length;
  const pestleCount = dataset.filter(record => record.PESTLE_Tags?.length).length;

  let html = renderValidationNotice(layer);
  html += `<div class="index-hero">
    <div>
      <div class="label">CIPQ Index</div>
      <div class="big-number">${layer.cipq_index.toFixed(2)}</div>
    </div>
    <div>
      <div class="label">Structural Pressure Level</div>
      <div style="font-family:'Playfair Display',serif;font-size:1.6rem;font-weight:700;color:var(--gold)">${escapeHtml(layer.structural_pressure_label)}</div>
      <div class="desc">Average of the active quadrant severity scores across Creation, Production, Distribution, and Access.</div>
      ${renderPressureLegend(layer.structural_pressure_label)}
    </div>
    <div style="font-family:'IBM Plex Mono',monospace;font-size:0.78rem;color:#c8bfae">
      <div>${dataset.length} segments coded</div>
      <div style="margin-top:0.3rem">${layer.aggregates.indicator_stats.length} indicators active</div>
      <div style="margin-top:0.3rem">${stakeholderCount} stakeholder groups</div>
      <div style="margin-top:0.3rem">${regionCount} regions | ${sourceCount} source types</div>
      <div style="margin-top:0.3rem">${valueChainAnalysisCount} value chain analyzed (${valueChainCount} explicit) | ${pestleCount} PESTLE tagged</div>
    </div>
  </div>`;

  html += `<div class="summary-shell">
    <h3>${escapeHtml(layer.dashboard_summary.title)}</h3>
    <p>The client dashboard turns coded evidence into readable policy interpretations while keeping traceability back to original excerpts.</p>
    <div class="summary-list">
      ${layer.dashboard_summary.sentences.map(sentence => `<div class="summary-item"><p>${escapeHtml(sentence)}</p></div>`).join('')}
    </div>
  </div>`;

  html += `<div class="section-title">Main Readings <span>Chart explanations and top-line meaning</span></div>
    <div class="explanation-grid">
      <div class="explanation-card">
        <h3>Quadrant Frequency</h3>
        <p>${escapeHtml(layer.chart_explanations.quadrant_frequency_chart.text)}</p>
      </div>
      <div class="explanation-card">
        <h3>Indicator Severity</h3>
        <p>${escapeHtml(layer.chart_explanations.indicator_severity_chart.text)}</p>
      </div>
      <div class="explanation-card">
        <h3>Stakeholder Comparison</h3>
        <p>${escapeHtml(layer.chart_explanations.stakeholder_comparison_chart.text)}</p>
      </div>
    </div>`;

  html += renderChartVisualizations(layer);

  html += `<div class="section-title">Quadrant Summary Cards <span>Frequency, severity, narrative, and evidence</span></div>
    <div class="insight-grid">`;
  layer.quadrant_cards.forEach(card => {
    html += `<article class="insight-card">
      <div class="card-head">
        <div>
          <h3>${escapeHtml(card.quadrant)}</h3>
          <div class="meta-row">
            <span class="tag ${DOMAIN_CLASS[card.quadrant] || ''}">${escapeHtml(card.quadrant)}</span>
            ${renderConfidencePill(card.confidence_label)}
          </div>
        </div>
      </div>
      <div class="stat-pair">
        <div class="stat-block">
          <div class="label">Mentions</div>
          <span class="value">${escapeHtml(card.frequency)}</span>
        </div>
        <div class="stat-block">
          <div class="label">Avg Severity</div>
          <span class="value">${escapeHtml(card.average_severity.toFixed(2))}</span>
        </div>
      </div>
      <p>${escapeHtml(card.summary_text)}</p>
      <div class="muted-inline" style="margin-top:0.85rem">Top indicators: ${escapeHtml(card.top_indicators.join(', ') || 'None yet')}</div>
      ${renderEvidenceHtml(card.evidence)}
    </article>`;
  });
  html += '</div>';

  html += `<div class="section-title">Priority Policy Signals <span>Ranked by weighted score</span></div>
    <div class="signal-grid">`;
  layer.priority_signals.slice(0, 6).forEach(signal => {
    html += `<article class="insight-card signal-card">
      <div class="signal-head">
        <div>
          <h3>${escapeHtml(signal.indicator)}</h3>
          <div class="meta-row">
            <span class="tag ${DOMAIN_CLASS[signal.domain] || ''}">${escapeHtml(signal.domain || 'Unspecified')}</span>
            <span class="priority-badge">${escapeHtml(signal.classification.replace(/_/g, ' '))}</span>
            ${renderConfidencePill(signal.confidence_label)}
          </div>
        </div>
      </div>
      <div class="stat-pair">
        <div class="stat-block"><div class="label">Frequency</div><span class="value">${escapeHtml(signal.frequency)}</span></div>
        <div class="stat-block"><div class="label">Avg Severity</div><span class="value">${escapeHtml(signal.average_severity.toFixed(2))}</span></div>
      </div>
      <p class="narrative">${escapeHtml(signal.narrative)}</p>
      ${renderEvidenceHtml(signal.evidence)}
    </article>`;
  });
  html += '</div>';

  html += `<div class="section-title">Value Chain Pressure Mapping <span>Contextual grouping for report interpretation</span></div>`;
  html += renderContextSummaryCards(layer.context_summaries.value_chain);

  html += `<div class="section-title">PESTLE Context Summary <span>Macro-context tags, not a separate scoring engine</span></div>`;
  html += renderContextSummaryCards(layer.context_summaries.pestle);

  html += `<div class="section-title">Cross-Quadrant Reading <span>Inference-based, not causal overclaiming</span></div>
    <div class="reading-list">`;
  layer.cross_quadrant_reading.forEach(reading => {
    html += `<div class="reading-item">
      <div class="meta-row">${renderConfidencePill(reading.confidence_label)}</div>
      <p>${escapeHtml(reading.sentence)}</p>
      ${renderEvidenceHtml(reading.evidence)}
    </div>`;
  });
  html += '</div>';

  html += `<div class="section-title">Stakeholder Perspectives <span>How different groups frame the problem</span></div>
    <div class="stakeholder-grid">`;
  layer.stakeholder_insights.forEach(insight => {
    html += `<article class="insight-card">
      <div class="card-head">
        <div>
          <h3>${escapeHtml(insight.stakeholder)}</h3>
          <div class="meta-row">
            <span class="tag ${DOMAIN_CLASS[insight.top_quadrant] || ''}">${escapeHtml(insight.top_quadrant)}</span>
            ${renderConfidencePill(insight.confidence_label)}
          </div>
        </div>
      </div>
      <p>${escapeHtml(insight.narrative)}</p>
      ${insight.difference_note ? `<p style="margin-top:0.65rem"><strong>Relative difference:</strong> ${escapeHtml(insight.difference_note)}</p>` : ''}
      <div class="muted-inline" style="margin-top:0.85rem">Top indicators: ${escapeHtml(insight.top_indicators.join(', ') || 'None yet')}</div>
      ${renderEvidenceHtml(insight.evidence)}
    </article>`;
  });
  html += '</div>';

  if (topIndicator) {
    html += `<div class="section-title">Top Indicator Snapshot <span>Highest weighted policy pressure</span></div>
      <div class="summary-shell">
        <h3>${escapeHtml(topIndicator.label)}</h3>
        <p>${escapeHtml(applyConfidencePrefix(topIndicator.confidence_label, `${topIndicator.label} currently has the strongest combination of frequency and severity in the dataset.`))}</p>
        <div class="stat-pair">
          <div class="stat-block"><div class="label">Frequency</div><span class="value">${escapeHtml(topIndicator.frequency)}</span></div>
          <div class="stat-block"><div class="label">Weighted Score</div><span class="value">${escapeHtml(topIndicator.weighted_score.toFixed(2))}</span></div>
        </div>
        ${renderEvidenceHtml(topIndicator.evidence)}
      </div>`;
  }

  mount.innerHTML = html;
}

function renderIndicators() {
  const mount = document.getElementById('indicatorContent');
  if (!dataset.length) {
    mount.innerHTML = '<div class="no-data-msg">No data yet.</div>';
    renderTraceability('indicatorTraceability');
    return;
  }

  const filterDomain = document.getElementById('indFilterDomain').value;
  const filterValueChain = document.getElementById('indFilterValueChain')?.value || '';
  const filterPestle = document.getElementById('indFilterPestle')?.value || '';
  let sourceRecords = dataset;
  if (filterDomain) sourceRecords = sourceRecords.filter(record => record.CIPQ_Domain === filterDomain);
  if (filterValueChain) sourceRecords = sourceRecords.filter(record => record.Value_Chain_Stage === filterValueChain);
  if (filterPestle) sourceRecords = sourceRecords.filter(record => (record.PESTLE_Tags || []).includes(filterPestle));
  const items = aggregateBy(sourceRecords, record => record.Indicator_Code, (key, sample) => sample.Indicator_Name || getIndicatorLabel(key) || key);
  const severityExplanation = explainSeverityChart(items);

  if (!items.length) {
    mount.innerHTML = '<div class="no-data-msg">No indicators match the current filters.</div>';
    renderTraceability('indicatorTraceability');
    return;
  }

  let html = `<div class="summary-shell"><h3>What this means</h3><p>${escapeHtml(severityExplanation.text)}</p></div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Code</th>
            <th>Indicator</th>
            <th>Quadrant</th>
            <th>Frequency</th>
            <th>Avg Severity</th>
            <th>Weighted</th>
            <th>Spread</th>
            <th>Confidence</th>
          </tr>
        </thead>
        <tbody>`;

  items.forEach(item => {
    html += `<tr>
      <td><code style="font-family:'IBM Plex Mono',monospace;font-weight:600">${escapeHtml(item.key)}</code></td>
      <td><button class="inline-link-btn" type="button" onclick="openIndicatorTrace('${item.key}')">${escapeHtml(item.label)}</button></td>
      <td><span class="tag ${DOMAIN_CLASS[item.domain] || ''}">${escapeHtml(item.domain || 'Unspecified')}</span></td>
      <td style="font-family:'IBM Plex Mono',monospace;text-align:center">${escapeHtml(item.frequency)}</td>
      <td style="font-family:'IBM Plex Mono',monospace;font-weight:600;color:var(--rust)">${escapeHtml(item.average_severity.toFixed(2))}</td>
      <td style="font-family:'IBM Plex Mono',monospace">${escapeHtml(item.weighted_score.toFixed(2))}</td>
      <td style="font-size:0.78rem">${escapeHtml(`${item.stakeholder_spread} stakeholders | ${item.regional_spread} regions | ${item.source_spread} sources`)}</td>
      <td>${renderConfidencePill(item.confidence_label)}</td>
    </tr>`;
  });

  html += '</tbody></table></div>';
  mount.innerHTML = html;
  renderTraceability('indicatorTraceability');
}

function renderComparison() {
  const mount = document.getElementById('compContent');
  if (!dataset.length) {
    mount.innerHTML = '<div class="no-data-msg">No data yet.</div>';
    return;
  }

  const layer = buildInterpretiveLayer();
  const stakeholders = [...new Set(dataset.map(record => record.Stakeholder).filter(Boolean))].sort();
  const regions = [...new Set(dataset.map(record => record.Region).filter(Boolean))].sort();

  let html = `<div class="summary-shell"><h3>What this means</h3><p>${escapeHtml(layer.chart_explanations.stakeholder_comparison_chart.text)}</p></div>`;
  html += `<div class="stakeholder-grid">`;
  layer.stakeholder_insights.forEach(insight => {
    html += `<article class="insight-card">
      <div class="card-head">
        <div>
          <h3>${escapeHtml(insight.stakeholder)}</h3>
          <div class="meta-row">
            <span class="tag ${DOMAIN_CLASS[insight.top_quadrant] || ''}">${escapeHtml(insight.top_quadrant)}</span>
            ${renderConfidencePill(insight.confidence_label)}
          </div>
        </div>
      </div>
      <p>${escapeHtml(insight.narrative)}</p>
      ${insight.difference_note ? `<p style="margin-top:0.65rem">${escapeHtml(insight.difference_note)}</p>` : ''}
    </article>`;
  });
  html += '</div>';

  html += buildComparisonTable('Stakeholder', stakeholders, 'By Stakeholder Group');
  html += buildComparisonTable('Region', regions, 'By Region');
  mount.innerHTML = html || '<div class="no-data-msg">Not enough metadata to compare. Add stakeholder and region fields.</div>';
}

function renderPriority() {
  const mount = document.getElementById('priorityContent');
  if (!dataset.length) {
    mount.innerHTML = '<div class="no-data-msg">No data yet.</div>';
    return;
  }

  const layer = buildInterpretiveLayer();
  const minSeverity = parseFloat(document.getElementById('minSev').value);
  const minFrequency = parseInt(document.getElementById('minCount').value, 10);
  const items = layer.priority_signals.filter(signal => signal.average_severity >= minSeverity && signal.frequency >= minFrequency);

  if (!items.length) {
    mount.innerHTML = '<div class="no-data-msg">No indicators meet the current thresholds. Try lowering the filters.</div>';
    return;
  }

  let html = `<div class="signal-grid">`;
  items.forEach(signal => {
    html += `<article class="insight-card signal-card">
      <div class="signal-head">
        <div>
          <h3>${escapeHtml(signal.indicator)}</h3>
          <div class="meta-row">
            <span class="tag ${DOMAIN_CLASS[signal.domain] || ''}">${escapeHtml(signal.domain || 'Unspecified')}</span>
            <span class="priority-badge">${escapeHtml(signal.classification.replace(/_/g, ' '))}</span>
            ${renderConfidencePill(signal.confidence_label)}
          </div>
        </div>
      </div>
      <div class="stat-pair">
        <div class="stat-block"><div class="label">Frequency</div><span class="value">${escapeHtml(signal.frequency)}</span></div>
        <div class="stat-block"><div class="label">Weighted</div><span class="value">${escapeHtml(signal.weighted_score.toFixed(2))}</span></div>
      </div>
      <p class="narrative">${escapeHtml(signal.narrative)}</p>
      ${renderEvidenceHtml(signal.evidence)}
    </article>`;
  });
  html += '</div>';
  mount.innerHTML = html;
}

function renderDataset() {
  const validationMount = document.getElementById('validationContent');
  if (validationMount) validationMount.innerHTML = renderAnalystValidationPanel();

  const filterDomain = document.getElementById('dsFilterDomain').value;
  const filterStakeholder = document.getElementById('dsFilterStakeholder').value;
  const filterValueChain = document.getElementById('dsFilterValueChain')?.value || '';
  const filterPestle = document.getElementById('dsFilterPestle')?.value || '';

  const stakeholderSelect = document.getElementById('dsFilterStakeholder');
  const currentStakeholder = stakeholderSelect.value;
  const stakeholderOptions = [...new Set(dataset.map(record => record.Stakeholder).filter(Boolean))].sort();
  stakeholderSelect.innerHTML = '<option value="">All</option>';
  stakeholderOptions.forEach(stakeholder => {
    stakeholderSelect.innerHTML += `<option${stakeholder === currentStakeholder ? ' selected' : ''}>${escapeHtml(stakeholder)}</option>`;
  });

  let rows = dataset;
  if (filterDomain) rows = rows.filter(record => record.CIPQ_Domain === filterDomain);
  if (filterStakeholder) rows = rows.filter(record => record.Stakeholder === filterStakeholder);
  if (filterValueChain) rows = rows.filter(record => record.Value_Chain_Stage === filterValueChain);
  if (filterPestle) rows = rows.filter(record => (record.PESTLE_Tags || []).includes(filterPestle));

  document.getElementById('datasetCount').textContent = `${rows.length} of ${dataset.length} segments`;
  if (!rows.length) {
    document.getElementById('datasetContent').innerHTML = '<div class="no-data-msg">No segments match current filters.</div>';
    return;
  }

  let html = `<div class="table-wrap">
    <table class="dataset-table">
      <colgroup>
        <col style="width:120px">
        <col style="width:240px">
        <col style="width:200px">
        <col style="width:100px">
        <col style="width:100px">
        <col style="width:110px">
        <col style="width:130px">
        <col style="width:150px">
        <col style="width:90px">
        <col style="width:80px">
        <col style="width:80px">
      </colgroup>
      <thead>
        <tr>
          <th>Segment ID</th>
          <th>Snippet</th>
          <th>Indicator</th>
          <th>Quadrant</th>
          <th>Severity</th>
          <th>Stakeholder</th>
          <th>Value Chain</th>
          <th>PESTLE</th>
          <th>Region</th>
          <th>Confidence</th>
          ${canWrite() ? '<th></th>' : ''}
        </tr>
      </thead>
      <tbody>`;

  rows.forEach(record => {
    const isExpanded = expandedSnippetIds.has(record.Segment_ID);
    const isLong = record.Snippet.length > 90;
    const snippet = isExpanded || !isLong ? record.Snippet : `${record.Snippet.slice(0, 90)}...`;
    const sourceLabel = [record.Source_Type, record.Source_ID].filter(Boolean).join(' · ');
    const respondentLabel = record.Respondent_Type ? `<div class="ds-meta-sub">${escapeHtml(record.Respondent_Type)}</div>` : '';
    html += `<tr>
      <td>
        <code class="ds-segid">${escapeHtml(record.Segment_ID)}</code>
        ${sourceLabel ? `<div class="ds-meta-sub">${escapeHtml(sourceLabel)}</div>` : ''}
      </td>
      <td>
        <button class="snippet-toggle${isExpanded ? ' expanded' : ''}" type="button" onclick="toggleDatasetSnippet('${record.Segment_ID}')">
          ${escapeHtml(snippet || '-')}
          ${isLong ? `<span class="snippet-hint">${isExpanded ? 'Collapse' : 'Expand'}</span>` : ''}
        </button>
      </td>
      <td>
        <button class="inline-link-btn ds-indicator" type="button" onclick="openIndicatorTrace('${record.Indicator_Code}','indicators')">${escapeHtml(record.Indicator_Name || record.Indicator_Code)}</button>
        <div class="ds-meta-sub">${escapeHtml(record.Indicator_Code)}</div>
      </td>
      <td>
        <span class="tag ${DOMAIN_CLASS[record.CIPQ_Domain] || ''}">${escapeHtml(record.CIPQ_Domain || '-')}</span>
        ${record.Secondary_Domain ? `<div class="ds-meta-sub">+ ${escapeHtml(record.Secondary_Domain)}</div>` : ''}
      </td>
      <td>${sevDots(record.Severity)}</td>
      <td>
        <span class="ds-cell-text">${escapeHtml(record.Stakeholder || '-')}</span>
        ${respondentLabel}
      </td>
      <td><span class="ds-cell-text">${escapeHtml(record.Value_Chain_Stage || '-')}</span></td>
      <td><span class="ds-cell-text">${escapeHtml((record.PESTLE_Tags || []).join(', ') || '-')}</span></td>
      <td><span class="ds-cell-text">${escapeHtml(record.Region || '-')}</span></td>
      <td>${renderConfidencePill(record.Scoring_Confidence || 'medium')}</td>
      <td>${canWrite() ? `<button class="btn btn-secondary ds-delete-btn" type="button" onclick="deleteSegment('${record.Segment_ID}')">Delete</button>` : ''}</td>
    </tr>`;
  });

  html += '</tbody></table></div>';
  document.getElementById('datasetContent').innerHTML = html;
}

function showStatus(message, isError) {
  const bar = document.getElementById('statusBar');
  bar.textContent = message;
  bar.className = isError ? 'error' : '';
  bar.style.display = 'block';
  setTimeout(() => { bar.style.display = 'none'; }, 4500);
}

function updateCounts() {
  document.getElementById('segCount').textContent = `${dataset.length} segment${dataset.length !== 1 ? 's' : ''} loaded`;
  document.getElementById('datasetCount').textContent = `${dataset.length} segments`;
  renderAuthUI();
}

function refreshAll() {
  updateCounts();
  renderEntryPreview();
  const activePanel = document.querySelector('.tab-panel.active');
  if (!activePanel) return;
  if (activePanel.id === 'tab-dashboard') renderDashboard();
  if (activePanel.id === 'tab-indicators') renderIndicators();
  if (activePanel.id === 'tab-comparison') renderComparison();
  if (activePanel.id === 'tab-priority') renderPriority();
  if (activePanel.id === 'tab-simulator') renderSimulator();
  if (activePanel.id === 'tab-dataset') renderDataset();
}

Object.assign(window, {
  setAppView,
  switchTab,
  updateThemeSelection,
  updateIndicatorMetadata,
  updatePestleSummary,
  setSeverity,
  clearForm,
  addSegment,
  handleFileUpload,
  exportCSV,
  exportSummaryTablesCsv,
  downloadImportTemplate,
  downloadChartPack,
  exportInterpretiveJson,
  downloadClientReport,
  copyClientReport,
  updateSimulationControl,
  resetSimulation,
  downloadSimulationReport,
  openIndicatorTrace,
  closeIndicatorTrace,
  deleteSegment,
  toggleDatasetSnippet,
  clearAllSegments,
  signInUser,
  signOutUser,
  openLoginModal,
  closeLoginModal,
  toggleUserDropdown,
  enforceGuestView,
  canWrite
});

window.addEventListener('DOMContentLoaded', () => {
  buildThemeOptions();
  updateUploadMeta();
  updatePestleSummary();
  updateCounts();
  renderEntryPreview();
  const uploadZone = document.getElementById('uploadZone');
  uploadZone.addEventListener('dragover', event => {
    event.preventDefault();
    uploadZone.classList.add('drag-over');
  });
  uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
  uploadZone.addEventListener('drop', event => {
    event.preventDefault();
    uploadZone.classList.remove('drag-over');
    if (event.dataTransfer.files[0]) parseCSV(event.dataTransfer.files[0]);
  });
  initializeAuth();
  setAppView('client');
  if (window.loadSurveyData) window.loadSurveyData();
});
