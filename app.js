let dataset = [];
let currentSeverity = null;
let activeTraceCode = null;
let activeView = 'encoder';

const SEVERITY_LABELS = {
  1: 'Minimal',
  2: 'Low',
  3: 'Moderate',
  4: 'High',
  5: 'Critical'
};

const DOMAIN_CLASS = { Creation: 'creation', Production: 'production', Distribution: 'distribution', Access: 'access' };
const DOMAIN_PREFIX = { Creation: 'CREATE', Production: 'PROD', Distribution: 'DIST', Access: 'ACCESS' };
const VALID_DOMAINS = ['Creation', 'Production', 'Distribution', 'Access'];
const LEGACY_DOMAIN_MAP = { Governance: 'Production' };
const SCORING_CONFIDENCE_OPTIONS = ['low', 'medium', 'high'];
const expandedSnippetIds = new Set();

const SUPABASE_URL = 'https://ueyyrugaynzczkcwnxbt.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVleXlydWdheW56Y3prY3dueGJ0Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY5NDEyOTIsImV4cCI6MjA5MjUxNzI5Mn0.x1XqiUs52OH4TKe070OSOK4c0gSGeYRSZJ30XZSBSuQ';
const SUPABASE_TABLE = 'segments';
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

function parseBoolean(value) {
  if (typeof value === 'boolean') return value;
  return ['true', '1', 'yes', 'y'].includes(String(value ?? '').trim().toLowerCase());
}

function formatSupabaseError(error) {
  if (!error) return 'Unknown error.';
  const message = error.message || String(error);
  if (/column|schema|record_confidence|theme_code|stakeholder_group|quadrant_primary|indicator_label/i.test(message)) {
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
    Source_Type: record.Source_Type || record.source_type || record.Source || record['metadata.source_type'] || '',
    Source_ID: record.Source_ID || record.source_id || record['metadata.source_id'] || '',
    Scoring_Confidence: SCORING_CONFIDENCE_OPTIONS.includes(scoringConfidence) ? scoringConfidence : 'medium',
    Is_Cross_Quadrant: parseBoolean(record.Is_Cross_Quadrant ?? record.is_cross_quadrant ?? record['analysis_flags.is_cross_quadrant']) || !!secondary || linkedQuadrants.length > 1,
    Linked_Quadrants: linkedQuadrants.length ? linkedQuadrants : [primary, secondary].filter(Boolean),
    Analysis_Notes: record.Analysis_Notes || record.analysis_notes || record['analysis_flags.notes'] || '',
    Session_ID: record.Session_ID || record.session_id || '',
    DB_ID: record.DB_ID || record.id || null,
    Created_At: createdAt,
    Updated_At: updatedAt
  };
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
  if (!record.Source_ID) issues.push('Missing source ID.');

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
    email: document.getElementById('auth_email').value.trim(),
    password: document.getElementById('auth_password').value
  };
}

function clearAuthForm() {
  document.getElementById('auth_email').value = '';
  document.getElementById('auth_password').value = '';
}

function setCloudSyncing(syncing) {
  isCloudSyncing = syncing;
  renderAuthUI();
}

function renderAuthUI() {
  const pill = document.getElementById('authStatePill');
  const email = document.getElementById('authUserEmail');
  const note = document.getElementById('authSyncNote');
  const loginBtn = document.getElementById('authLoginBtn');
  const signupBtn = document.getElementById('authSignupBtn');
  const logoutBtn = document.getElementById('authLogoutBtn');
  const emailInput = document.getElementById('auth_email');
  const passwordInput = document.getElementById('auth_password');
  if (!pill || !email || !note || !loginBtn || !signupBtn || !logoutBtn || !emailInput || !passwordInput) return;

  const authInputsDisabled = !supabaseReady || isCloudSyncing;
  emailInput.disabled = authInputsDisabled || !!currentUser;
  passwordInput.disabled = authInputsDisabled || !!currentUser;
  loginBtn.disabled = authInputsDisabled || !!currentUser;
  signupBtn.disabled = authInputsDisabled || !!currentUser;
  logoutBtn.disabled = !supabaseReady || isCloudSyncing || !currentUser;

  pill.className = 'auth-state-pill';
  if (!supabaseReady) {
    pill.textContent = 'Local Only';
    email.textContent = 'Supabase not configured';
    note.textContent = 'Add your Supabase URL and anon key, then run the updated supabase_schema.sql before using cloud sync.';
    return;
  }

  if (isCloudSyncing) {
    pill.textContent = 'Syncing';
    pill.classList.add('syncing');
    email.textContent = currentUser?.email || 'Preparing workspace';
    note.textContent = 'Syncing your encoded segment history from Supabase.';
    return;
  }

  if (currentUser) {
    pill.textContent = 'Cloud Sync On';
    pill.classList.add('connected');
    email.textContent = currentUser.email || 'Signed in';
    note.textContent = `Signed in. ${dataset.length} segment${dataset.length !== 1 ? 's' : ''} currently loaded from cloud storage.`;
    return;
  }

  pill.textContent = 'Ready To Sign In';
  email.textContent = 'No active user session';
  note.textContent = 'Sign in or create an account to save segment history to Supabase. Until then, changes stay only in this browser session.';
}

async function loadSegmentsFromSupabase(options = {}) {
  if (!supabaseClient || !currentUser) {
    dataset = [];
    expandedSnippetIds.clear();
    refreshAll();
    renderAuthUI();
    return;
  }

  setCloudSyncing(true);
  const { data, error } = await supabaseClient
    .from(SUPABASE_TABLE)
    .select('*')
    .order('created_at', { ascending: true, nullsFirst: true })
    .order('encoded_at', { ascending: true, nullsFirst: true })
    .order('segment_id', { ascending: true });

  if (error) {
    setCloudSyncing(false);
    showStatus(`Could not load saved segments: ${formatSupabaseError(error)}`, true);
    return;
  }

  dataset = (data || []).map(mapDbRowToSegment);
  expandedSnippetIds.clear();
  refreshAll();
  setCloudSyncing(false);
  if (!options.silent) showStatus(`Loaded ${dataset.length} saved segment${dataset.length !== 1 ? 's' : ''} from Supabase.`, false);
}

async function handleSessionChange(session, options = {}) {
  currentSession = session || null;
  currentUser = session?.user || null;

  if (!currentUser) {
    dataset = [];
    expandedSnippetIds.clear();
    refreshAll();
    renderAuthUI();
    if (!options.silent) showStatus('Signed out. Local browser memory is now empty.', false);
    return;
  }

  renderAuthUI();
  await loadSegmentsFromSupabase({ silent: !!options.silent });
}

async function initializeAuth() {
  renderAuthUI();
  if (!supabaseClient) return;

  const { data, error } = await supabaseClient.auth.getSession();
  if (error) {
    showStatus(`Could not restore session: ${formatSupabaseError(error)}`, true);
    return;
  }

  await handleSessionChange(data?.session || null, { silent: true });
  supabaseClient.auth.onAuthStateChange((event, session) => {
    handleSessionChange(session, { silent: event === 'INITIAL_SESSION' });
  });
}

async function signUpUser() {
  if (!supabaseClient) {
    showStatus('Supabase is not configured yet.', true);
    return;
  }
  const { email, password } = authCredentials();
  if (!email || !password) {
    showStatus('Enter both email and password to create an account.', true);
    return;
  }

  setCloudSyncing(true);
  const { data, error } = await supabaseClient.auth.signUp({ email, password });
  setCloudSyncing(false);
  if (error) {
    showStatus(formatSupabaseError(error), true);
    return;
  }

  clearAuthForm();
  renderAuthUI();
  if (data?.session) {
    await handleSessionChange(data.session, { silent: true });
    showStatus('Account created and signed in. Cloud-backed history is ready.', false);
    return;
  }

  showStatus('Account created. Check your inbox if email confirmation is enabled, then sign in.', false);
}

async function signInUser() {
  if (!supabaseClient) {
    showStatus('Supabase is not configured yet.', true);
    return;
  }
  const { email, password } = authCredentials();
  if (!email || !password) {
    showStatus('Enter both email and password to sign in.', true);
    return;
  }

  setCloudSyncing(true);
  const { data, error } = await supabaseClient.auth.signInWithPassword({ email, password });
  setCloudSyncing(false);
  if (error) {
    showStatus(formatSupabaseError(error), true);
    return;
  }

  clearAuthForm();
  await handleSessionChange(data.session, { silent: true });
  showStatus('Logged in. Saved segment history reloaded from Supabase.', false);
}

async function signOutUser() {
  if (!supabaseClient || !currentUser) return;

  setCloudSyncing(true);
  const { error } = await supabaseClient.auth.signOut();
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
  if (name === 'dataset') renderDataset();
  if (name === 'entry') renderEntryPreview();
}

function setSeverity(value) {
  currentSeverity = value;
  document.getElementById('f_severity').value = value;
  document.querySelectorAll('.sev-picker button').forEach((button, index) => {
    button.classList.toggle('selected', index + 1 <= value);
  });
}

function clearForm() {
  ['f_segid', 'f_theme_label', 'f_domain', 'f_indicator_label', 'f_respondent_type', 'f_source_id', 'f_session', 'f_snippet', 'f_analysis_notes'].forEach(id => {
    const node = document.getElementById(id);
    if (node) node.value = '';
  });
  ['f_theme_code', 'f_indicator', 'f_stakeholder', 'f_region', 'f_domain_secondary', 'f_source_type', 'f_confidence'].forEach(id => {
    const node = document.getElementById(id);
    if (node) node.value = id === 'f_confidence' ? 'medium' : '';
  });
  document.getElementById('f_indicator').innerHTML = '<option value="">-- Select Theme First --</option>';
  currentSeverity = null;
  document.getElementById('f_severity').value = '';
  document.querySelectorAll('.sev-picker button').forEach(button => button.classList.remove('selected'));
}

async function addSegment() {
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
    Segment_ID: ensureUniqueSegmentId(requestedId || `SEG_${Date.now()}`),
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

  if (supabaseClient && currentUser) {
    setCloudSyncing(true);
    const { data, error } = await supabaseClient
      .from(SUPABASE_TABLE)
      .insert(mapSegmentToDbRow(segmentRecord))
      .select()
      .single();
    setCloudSyncing(false);

    if (error) {
      showStatus(`Could not save segment to Supabase: ${formatSupabaseError(error)}`, true);
      return;
    }

    dataset.push(mapDbRowToSegment(data));
    showStatus(`Segment ${segmentRecord.Segment_ID} saved to Supabase. Total: ${dataset.length}`, false);
  } else {
    dataset.push(segmentRecord);
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
      const takenIds = new Set(dataset.map(record => record.Segment_ID));

      rows.forEach((row, index) => {
        const record = normalizeSegmentRecord(row);
        if (!record.CIPQ_Domain) record.CIPQ_Domain = inferDomainFromCode(record.Theme_Code || record.Indicator_Code);
        if (!record.Theme_Code) record.Theme_Code = themeCodeForIndicator(record.Indicator_Code);
        if (!record.Theme) record.Theme = getThemeLabel(record.Theme_Code);
        if (!record.Indicator_Name) record.Indicator_Name = getIndicatorLabel(record.Indicator_Code);
        if (!record.Source_ID) record.Source_ID = row.Session_ID || row.session_id || '';
        if (!record.Segment_ID) record.Segment_ID = ensureUniqueSegmentId(`IMP_${Date.now()}_${index + 1}`, takenIds);
        else record.Segment_ID = ensureUniqueSegmentId(record.Segment_ID, takenIds);
        if (!VALID_DOMAINS.includes(record.CIPQ_Domain) || !record.Indicator_Code || !Number.isFinite(record.Severity)) return;
        importedSegments.push(record);
      });

      if (!importedSegments.length) {
        showStatus('No valid rows were found in the CSV import.', true);
        updateUploadMeta(file.name, false);
        return;
      }

      if (supabaseClient && currentUser) {
        setCloudSyncing(true);
        let saveError = null;
        for (const batch of chunkItems(importedSegments, 200)) {
          const result = await supabaseClient.from(SUPABASE_TABLE).insert(batch.map(mapSegmentToDbRow));
          if (result.error) {
            saveError = result.error;
            break;
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
        dataset.push(...importedSegments);
      }

      updateUploadMeta(file.name, true);
      const syncLabel = supabaseClient && currentUser ? ' and saved to Supabase' : ' locally';
      showStatus(`Imported ${importedSegments.length} segments from CSV${syncLabel}. Total: ${dataset.length}`, false);
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
    Is_Cross_Quadrant: record.Is_Cross_Quadrant,
    Linked_Quadrants: (record.Linked_Quadrants || []).join('|'),
    Analysis_Notes: record.Analysis_Notes || '',
    Session_ID: record.Session_ID || '',
    Encoded_At: record.Created_At || '',
    Updated_At: record.Updated_At || ''
  })));
  downloadBlob(`CIPQ_Dataset_${new Date().toISOString().slice(0, 10)}.csv`, csv, 'text/csv;charset=utf-8');
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
    confidence_label: item.confidence_label
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
    client_view: {
      top_panel: layer.dashboard_summary,
      main_charts: ['quadrant_frequency_chart', 'indicator_severity_chart', 'stakeholder_comparison_chart'],
      chart_explanations: layer.chart_explanations,
      quadrant_cards: layer.quadrant_cards,
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

function downloadClientReport() {
  const layer = buildInterpretiveLayer();
  if (!layer) {
    showStatus('No data to export yet.', true);
    return;
  }

  const lines = [];
  lines.push('CIPQ Client Report');
  lines.push(`Generated: ${new Date().toISOString()}`);
  lines.push('');
  lines.push('Policy Insights');
  layer.dashboard_summary.sentences.forEach(sentence => lines.push(`- ${sentence}`));
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
  lines.push('Stakeholder Perspectives');
  layer.stakeholder_insights.forEach(insight => {
    lines.push(`- ${insight.stakeholder}: ${insight.narrative}`);
    if (insight.difference_note) lines.push(`  ${insight.difference_note}`);
  });
  lines.push('');
  lines.push('Cross-Quadrant Reading');
  layer.cross_quadrant_reading.forEach(reading => lines.push(`- ${reading.sentence}`));

  downloadBlob(`CIPQ_Client_Report_${new Date().toISOString().slice(0, 10)}.txt`, lines.join('\r\n'), 'text/plain;charset=utf-8');
}

function renderConfidencePill(label) {
  return `<span class="confidence-pill ${escapeHtml(label)}">${escapeHtml(label)}</span>`;
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
  if (!layer.validation.issue_count) return '';
  const issuePreview = layer.validation.issues.slice(0, 5).map(item => `<li><strong>${escapeHtml(item.id)}</strong>: ${escapeHtml(item.issues[0])}</li>`).join('');
  return `<div class="validation-banner">
    <strong>Validation notes</strong>
    <p>${layer.validation.issue_count} record${layer.validation.issue_count !== 1 ? 's' : ''} contain schema or codebook warnings. The client interpretation layer still renders, but you may want to review the flagged records in Encoder View.</p>
    <ul>${issuePreview}</ul>
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
      <table>
        <thead>
          <tr>
            <th>Segment ID</th>
            <th>Theme</th>
            <th>Quadrant</th>
            <th>Indicator</th>
            <th>Severity</th>
            <th>Stakeholder</th>
            <th>Source</th>
            <th></th>
          </tr>
        </thead>
        <tbody>`;

  recent.forEach(record => {
    html += `<tr>
      <td><code style="font-size:0.72rem">${escapeHtml(record.Segment_ID)}</code></td>
      <td style="font-size:0.8rem">${escapeHtml(record.Theme_Code || '')}<br>${escapeHtml(record.Theme || '-')}</td>
      <td><span class="tag ${DOMAIN_CLASS[record.CIPQ_Domain] || ''}">${escapeHtml(record.CIPQ_Domain || '-')}</span></td>
      <td style="font-size:0.8rem">${escapeHtml(record.Indicator_Code)} | ${escapeHtml(record.Indicator_Name)}</td>
      <td>${sevDots(record.Severity)}</td>
      <td style="font-size:0.8rem">${escapeHtml(record.Stakeholder || '-')}</td>
      <td style="font-size:0.8rem">${escapeHtml(record.Source_Type || '-')} ${record.Source_ID ? `| ${escapeHtml(record.Source_ID)}` : ''}</td>
      <td><button class="btn btn-secondary" style="padding:0.45rem 0.65rem;font-size:0.65rem;min-height:36px" type="button" onclick="deleteSegment('${record.Segment_ID}')">Delete</button></td>
    </tr>`;
  });

  html += '</tbody></table></div>';
  document.getElementById('entryPreview').innerHTML = html;
}

async function deleteSegment(id) {
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

function renderDashboard() {
  const mount = document.getElementById('dashContent');
  if (!dataset.length) {
    mount.innerHTML = '<div class="no-data-msg">No data yet. Encode or import segments in Encoder View to generate client-facing insights.</div>';
    return;
  }

  const layer = buildInterpretiveLayer();
  const topIndicator = layer.aggregates.indicator_stats[0];
  const stakeholderCount = countDistinct(dataset.map(record => record.Stakeholder));
  const regionCount = countDistinct(dataset.map(record => record.Region));
  const sourceCount = countDistinct(dataset.map(record => record.Source_Type));

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
    </div>
    <div style="font-family:'IBM Plex Mono',monospace;font-size:0.78rem;color:#c8bfae">
      <div>${dataset.length} segments coded</div>
      <div style="margin-top:0.3rem">${layer.aggregates.indicator_stats.length} indicators active</div>
      <div style="margin-top:0.3rem">${stakeholderCount} stakeholder groups</div>
      <div style="margin-top:0.3rem">${regionCount} regions | ${sourceCount} source types</div>
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

  const layer = buildInterpretiveLayer();
  const filterDomain = document.getElementById('indFilterDomain').value;
  let items = layer.aggregates.indicator_stats;
  if (filterDomain) items = items.filter(item => item.domain === filterDomain);

  let html = `<div class="summary-shell"><h3>What this means</h3><p>${escapeHtml(layer.chart_explanations.indicator_severity_chart.text)}</p></div>
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
  const filterDomain = document.getElementById('dsFilterDomain').value;
  const filterStakeholder = document.getElementById('dsFilterStakeholder').value;

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

  document.getElementById('datasetCount').textContent = `${rows.length} of ${dataset.length} segments`;
  if (!rows.length) {
    document.getElementById('datasetContent').innerHTML = '<div class="no-data-msg">No segments match current filters.</div>';
    return;
  }

  let html = `<div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>Segment ID</th>
          <th>Theme</th>
          <th>Snippet</th>
          <th>Quadrant</th>
          <th>Indicator</th>
          <th>Severity</th>
          <th>Stakeholder</th>
          <th>Respondent</th>
          <th>Region</th>
          <th>Source</th>
          <th>Confidence</th>
          <th></th>
        </tr>
      </thead>
      <tbody>`;

  rows.forEach(record => {
    const isExpanded = expandedSnippetIds.has(record.Segment_ID);
    const isLong = record.Snippet.length > 90;
    const snippet = isExpanded || !isLong ? record.Snippet : `${record.Snippet.slice(0, 90)}...`;
    html += `<tr>
      <td><code style="font-size:0.72rem">${escapeHtml(record.Segment_ID)}</code></td>
      <td style="font-size:0.8rem">${escapeHtml(record.Theme_Code || '')}<br>${escapeHtml(record.Theme || '-')}</td>
      <td style="font-size:0.82rem;max-width:${isExpanded ? '420px' : '260px'}">
        <button class="snippet-toggle${isExpanded ? ' expanded' : ''}" type="button" onclick="toggleDatasetSnippet('${record.Segment_ID}')">
          ${escapeHtml(snippet || '-')}
          ${isLong ? `<span class="snippet-hint">${isExpanded ? 'Click to collapse' : 'Click to expand'}</span>` : ''}
        </button>
      </td>
      <td>
        <span class="tag ${DOMAIN_CLASS[record.CIPQ_Domain] || ''}">${escapeHtml(record.CIPQ_Domain || '-')}</span>
        ${record.Secondary_Domain ? `<div class="muted-inline" style="margin-top:0.35rem">Secondary: ${escapeHtml(record.Secondary_Domain)}</div>` : ''}
      </td>
      <td style="font-size:0.8rem"><button class="inline-link-btn" type="button" onclick="openIndicatorTrace('${record.Indicator_Code}','indicators')">${escapeHtml(record.Indicator_Code)} | ${escapeHtml(record.Indicator_Name)}</button></td>
      <td>${sevDots(record.Severity)}</td>
      <td style="font-size:0.8rem">${escapeHtml(record.Stakeholder || '-')}</td>
      <td style="font-size:0.8rem">${escapeHtml(record.Respondent_Type || '-')}</td>
      <td style="font-size:0.8rem">${escapeHtml(record.Region || '-')}</td>
      <td style="font-size:0.75rem;color:var(--muted)">${escapeHtml(record.Source_Type || '-')} ${record.Source_ID ? `| ${escapeHtml(record.Source_ID)}` : ''}</td>
      <td>${renderConfidencePill(record.Scoring_Confidence || 'medium')}</td>
      <td><button class="btn btn-secondary" style="padding:0.45rem 0.65rem;font-size:0.65rem;min-height:36px" type="button" onclick="deleteSegment('${record.Segment_ID}')">Delete</button></td>
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
  if (activePanel.id === 'tab-dataset') renderDataset();
}

Object.assign(window, {
  setAppView,
  switchTab,
  updateThemeSelection,
  updateIndicatorMetadata,
  setSeverity,
  clearForm,
  addSegment,
  handleFileUpload,
  exportCSV,
  exportInterpretiveJson,
  downloadClientReport,
  openIndicatorTrace,
  closeIndicatorTrace,
  deleteSegment,
  toggleDatasetSnippet,
  clearAllSegments,
  signInUser,
  signUpUser,
  signOutUser
});

window.addEventListener('DOMContentLoaded', () => {
  buildThemeOptions();
  updateUploadMeta();
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
  setAppView('encoder');
});
