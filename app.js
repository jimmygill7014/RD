/* ==========================================================================
   PURE FINANCIAL ADVISORS — CLIENT INTAKE WORKFLOW
   Multi-stage intake with validation, prefill, and Schwab readiness
   ========================================================================== */

'use strict';

/* ---------- CURRENCY FORMATTING ---------- */
function formatCommas(val) {
  // Takes a number or numeric string, returns comma-formatted string (e.g. 1234567 → "1,234,567")
  const n = parseFloat(String(val).replace(/,/g, ''));
  if (isNaN(n)) return '';
  // Show decimals only if present
  return n % 1 === 0
    ? n.toLocaleString('en-US')
    : n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function parseCommas(val) {
  // Strips commas and returns a Number (or NaN)
  return parseFloat(String(val).replace(/,/g, ''));
}

function attachCurrencyBehavior(inp) {
  // Format on blur, strip on focus
  inp.addEventListener('focus', () => {
    const n = parseCommas(inp.value);
    inp.value = isNaN(n) ? '' : String(n);
  });
  inp.addEventListener('blur', () => {
    const n = parseCommas(inp.value);
    inp.value = isNaN(n) ? '' : formatCommas(n);
  });
  // Format initial value if already populated
  if (inp.value !== '') {
    const n = parseCommas(inp.value);
    if (!isNaN(n)) inp.value = formatCommas(n);
  }
}

/* ---------- FOCUS MANAGEMENT ---------- */
// After renderPQ() rebuilds the DOM, focus the first visible input in a section
// so the user's tab flow isn't thrown to the top of the page.
function focusFirstInput(sectionId) {
  requestAnimationFrame(() => {
    const sec = findSectionEl(sectionId);
    if (!sec) return;
    const content = sec.querySelector('.section-content');
    if (!content || content.classList.contains('collapsed')) return;
    // Find the first visible, non-hidden, non-readonly input
    const inp = content.querySelector(
      'input:not([type="hidden"]):not([type="radio"]):not([type="checkbox"]):not([readonly]), select, textarea'
    );
    if (inp && inp.offsetParent !== null) {
      inp.focus();
    }
  });
}

// After renderPQ() rebuilds the DOM from a checkbox/click, keep the section visible
function scrollSectionIntoView(sectionId) {
  requestAnimationFrame(() => {
    const sec = findSectionEl(sectionId);
    if (sec) sec.scrollIntoView({ behavior: 'instant', block: 'nearest' });
  });
}

// Shared helper to locate a section's DOM element by its data-path prefix
function findSectionEl(sectionId) {
  return Array.from(document.querySelectorAll('.form-section')).find(s =>
    s.querySelector(`[data-path^="${sectionId}."]`) != null
  );
}

/* ---------- CONSTANTS ---------- */
const STORE_KEY = 'pfa-intake-v2';
const STAGES = ['pq', 'pq-review', 'plan', 'iap', 'client'];
let currentScreen = 'pq';

/* ---------- UTILITY FUNCTIONS ---------- */
function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  Object.entries(attrs).forEach(([k, v]) => {
    if (k === 'className') node.className = v;
    else if (k === 'textContent') node.textContent = v;
    else if (k === 'innerHTML') node.innerHTML = v;
    else if (k.startsWith('on')) node.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === 'dataset') Object.entries(v).forEach(([dk, dv]) => node.dataset[dk] = dv);
    else if (k === 'style' && typeof v === 'object') Object.assign(node.style, v);
    else node.setAttribute(k, v);
  });
  (Array.isArray(children) ? children : [children]).forEach(c => {
    if (c == null) return;
    node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
  });
  return node;
}

function parsePath(path) {
  return path.replace(/\]/g, '').split(/\.|\[/g).filter(Boolean);
}

function getDeep(src, path) {
  if (!src || !path) return undefined;
  return parsePath(path).reduce((a, k) => (a == null ? undefined : a[k]), src);
}

function setDeep(target, path, value) {
  const keys = parsePath(path);
  let ref = target;
  keys.forEach((key, idx) => {
    if (idx === keys.length - 1) { ref[key] = value; return; }
    const next = keys[idx + 1];
    const nextIsIdx = next !== undefined && /^\d+$/.test(next);
    if (ref[key] === undefined) ref[key] = nextIsIdx ? [] : {};
    ref = ref[key];
  });
}

function deepClone(v) { return JSON.parse(JSON.stringify(v)); }

function deepMerge(target, source) {
  if (!source) return target;
  Object.keys(source).forEach(k => {
    if (source[k] && typeof source[k] === 'object' && !Array.isArray(source[k]) && target[k] && typeof target[k] === 'object' && !Array.isArray(target[k])) {
      deepMerge(target[k], source[k]);
    } else if (source[k] !== undefined) {
      target[k] = source[k];
    }
  });
  return target;
}

function hasData(v) {
  if (v == null) return false;
  if (Array.isArray(v)) return v.some(hasData);
  if (typeof v === 'object') return Object.values(v).some(hasData);
  return String(v).trim() !== '';
}

function fmt$(n) {
  if (n == null || isNaN(n)) return '$0';
  return '$' + Number(n).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function fmtPct(n) {
  if (n == null || isNaN(n)) return '0%';
  return Number(n).toFixed(1) + '%';
}

function calcAge(dob) {
  if (!dob) return '';
  const d = new Date(dob);
  if (isNaN(d.getTime())) return '';
  const today = new Date();
  let age = today.getFullYear() - d.getFullYear();
  const m = today.getMonth() - d.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < d.getDate())) age--;
  return age;
}

function getOwnerNameOptions(pqData) {
  const opts = [];
  const family = pqData?.family || {};
  const c1First = (family.client1FirstName || '').trim();
  const c1Last = (family.client1LastName || '').trim();
  const c2First = (family.client2FirstName || '').trim();
  const c2Last = (family.client2LastName || '').trim();

  const c1 = [c1First, c1Last].filter(Boolean).join(' ');
  const c2 = [c2First, c2Last].filter(Boolean).join(' ');
  if (c1) opts.push(c1);
  if ((pqState.hasSpouse || c2First || c2Last) && c2) opts.push(c2);
  if (c1 && c2) opts.push('Joint');

  return opts;
}

function uid() { return '_' + Math.random().toString(36).slice(2, 9); }

function sumArray(arr, key) {
  if (!Array.isArray(arr)) return 0;
  return arr.reduce((s, r) => s + (parseFloat(r?.[key]) || 0), 0);
}

/* ---------- CENTRAL STORE ---------- */
function getStore() {
  try {
    return JSON.parse(localStorage.getItem(STORE_KEY) || '{}');
  } catch { return {}; }
}

function saveStore(data) {
  localStorage.setItem(STORE_KEY, JSON.stringify(data));
  refreshConsole();
  updatePipeline();
}

function refreshConsole() {
  const pre = document.getElementById('console-json');
  if (pre) pre.textContent = JSON.stringify(getStore(), null, 2);
}

/* ---------- PQ SCHEMA ---------- */
/* The PQ stage is built from this schema + conditional logic in the renderer */

const pqSections = [
  /* ---- 1. FAMILY INFO ---- */
  {
    id: 'family',
    title: 'Household / Family',
    fields: [
      { key: 'client1FirstName', label: 'First Name', width: 'medium', required: true },
      { key: 'client1LastName', label: 'Last Name', width: 'medium', required: true },
      { key: 'client1Nickname', label: 'Nickname', width: 'medium' },
      { key: 'client1DOB', label: 'Date of Birth', type: 'date', width: 'medium', required: true },
      { key: 'client1Age', label: 'Age', type: 'number', width: 'field-2' },
    ],
    conditionalBlocks: [
      {
        id: 'spouse-block',
        toggleLabel: '+ Add Spouse / Partner',
        hideLabel: '- Remove Spouse',
        flag: 'hasSpouse',
        fields: [
          { key: 'client2FirstName', label: 'Spouse First Name', width: 'medium' },
          { key: 'client2LastName', label: 'Spouse Last Name', width: 'medium' },
          { key: 'client2Nickname', label: 'Spouse Nickname', width: 'medium' },
          { key: 'client2DOB', label: 'Spouse Date of Birth', type: 'date', width: 'medium' },
          { key: 'client2Age', label: 'Spouse Age', type: 'number', width: 'field-2' },
          { key: 'yearsMarried', label: 'Years Married', type: 'number', width: 'field' },
        ]
      }
    ],
    afterConditional: [
      {
        id: 'children-block',
        toggleLabel: '+ Add Children / Grandchildren',
        hideLabel: '- Remove Children Section',
        flag: 'hasChildren',
        tables: [
          {
            key: 'children',
            title: 'Children',
            columns: [
              { key: 'name', label: 'Name' },
              { key: 'age', label: 'Age', type: 'number' },
            ]
          },
          {
            key: 'grandchildren',
            title: 'Grandchildren',
            columns: [
              { key: 'name', label: 'Name' },
              { key: 'age', label: 'Age', type: 'number' },
            ]
          },
        ],
        summaryFields: [
          { key: 'numChildren', label: '# of Children', type: 'number', width: 'field' },
          { key: 'numGrandchildren', label: '# of Grandchildren', type: 'number', width: 'field' },
        ]
      }
    ],
    footerFields: [
      { key: 'additionalFamilyInfo', label: 'Additional Family Information', type: 'textarea' },
    ]
  },

  /* ---- 2. CONTACT / RESIDENCE ---- */
  {
    id: 'contact',
    title: 'Contact / Residence',
    fields: [
      { key: 'address1', label: 'Street Address', width: 'wide' },
      { key: 'address2', label: 'Address Line 2', width: 'wide' },
      { key: 'city', label: 'City', width: 'medium' },
      { key: 'state', label: 'State', width: 'field' },
      { key: 'zip', label: 'Zip Code', width: 'field' },
      { key: 'homePhone', label: 'Home Phone', width: 'medium' },
      { key: 'client1Cell', label: 'Client Cell Phone', width: 'medium' },
      { key: 'client1Email', label: 'Client Email', type: 'email', width: 'wide' },
      { key: 'client2Cell', label: 'Spouse Cell Phone', width: 'medium', showIf: 'hasSpouse' },
      { key: 'client2Email', label: 'Spouse Email', type: 'email', width: 'wide', showIf: 'hasSpouse' },
      { key: 'referredBy', label: 'Referred By', width: 'medium' },
    ]
  },

  /* ---- 3. EMPLOYMENT ---- */
  {
    id: 'employment',
    title: 'Employment',
    customRenderer: 'renderEmploymentSection'
  },

  /* ---- 4. GOALS ---- */
  {
    id: 'goals',
    title: 'Goals & Planning Facts',
    fields: [
      { key: 'goals', label: 'Goals', type: 'textarea' },
      { key: 'estatePlan', label: 'Estate Plan', type: 'multiselect', options: ['Trust', 'Will', 'FPOA', 'MPOA'], width: 'medium' },
      { key: 'estatePlanYear', label: 'Year Established / Updated', width: 'medium' },
    ]
  },

  /* ---- 5. PROFESSIONAL RELATIONSHIPS ---- */
  {
    id: 'relationships',
    title: 'Current Professional Relationships',
    tables: [
      {
        key: 'professionals',
        columns: [
          { key: 'role', label: 'Role' },
          { key: 'name', label: 'Name' },
          { key: 'firmName', label: 'Firm' },
          { key: 'city', label: 'City' },
          { key: 'state', label: 'State' },
        ],
        starterRows: [
          { role: 'Financial Advisor' },
          { role: 'Attorney' },
          { role: 'Accountant' },
        ]
      }
    ]
  },

  /* ---- 6. ASSETS ---- */
  {
    id: 'assets',
    title: 'Assets',
    customRenderer: 'renderAssetsSection'
  },

  /* ---- 7. LIABILITIES ---- */
  {
    id: 'liabilities',
    title: 'Liabilities',
    customRenderer: 'renderLiabilitiesSection'
  },

  /* ---- 8. INSURANCE ---- */
  {
    id: 'insurance',
    title: 'Insurance',
    tables: [
      {
        key: 'policies',
        columns: [
          { key: 'company', label: 'Company' },
          { key: 'type', label: 'Type' },
          { key: 'benefit', label: 'Death/Daily Benefit', type: 'currency' },
          { key: 'insured', label: 'Insured' },
          { key: 'owner', label: 'Owner' },
          { key: 'policyDate', label: 'Policy Date', type: 'date' },
          { key: 'annualPremium', label: 'Annual Premium', type: 'currency' },
          { key: 'cashValue', label: 'Cash Value', type: 'currency' },
          { key: 'beneficiary', label: 'Beneficiary' },
        ]
      }
    ]
  },

  /* ---- 9. INCOME ---- */
  {
    id: 'income',
    title: 'Income',
    customRenderer: 'renderIncomeSection'
  },

  /* ---- 10. TAXES & EXPENSES ---- */
  {
    id: 'taxesExpenses',
    title: 'Taxes & Expenses',
    customRenderer: 'renderTaxesExpensesSection'
  },
];

/* ---------- PLAN SCHEMA ---------- */
const planSections = [
  {
    id: 'documents',
    title: 'Planning Documents Checklist',
    fields: [
      { key: 'investmentStatementsOnFile', label: 'Investment Statements on File', type: 'select', options: ['Yes', 'No', 'Pending'] },
      { key: 'taxReturnOnFile', label: 'Tax Return on File', type: 'select', options: ['Yes', 'No', 'Pending'] },
      { key: 'planningLiteRequested', label: 'Planning Lite Requested?', type: 'select', options: ['Yes', 'No'] },
      { key: 'dateEmailRequest', label: 'Date E-mail Request', type: 'date' },
    ]
  },
  {
    id: 'householdConfirm',
    title: 'Household Facts to Confirm',
    note: 'These fields are prefilled from the PQ. Please confirm or update each one.',
    fields: [
      { key: 'clientName', label: 'Client Name', width: 'medium', prefillFrom: 'family.client1FirstName+family.client1LastName' },
      { key: 'spouseName', label: 'Spouse Name', width: 'medium', prefillFrom: 'family.client2FirstName+family.client2LastName' },
      { key: 'clientAge', label: 'Client Age', type: 'number', prefillFrom: '_computed_client1Age' },
      { key: 'spouseAge', label: 'Spouse Age', type: 'number', prefillFrom: '_computed_client2Age' },
      { key: 'clientEmployment', label: 'Client Employment', width: 'wide', prefillFrom: 'employment.client1.status' },
      { key: 'spouseEmployment', label: 'Spouse Employment', width: 'wide', prefillFrom: 'employment.client2.status' },
      { key: 'clientRetirementDate', label: 'Client Retirement Date', type: 'date', prefillFrom: 'employment.client1.retirementDate' },
      { key: 'spouseRetirementDate', label: 'Spouse Retirement Date', type: 'date', prefillFrom: 'employment.client2.retirementDate' },
      { key: 'eMoneyRorPre', label: 'eMoney RoR (pre-retirement)' },
      { key: 'eMoneyRorPost', label: 'eMoney RoR (post-retirement)' },
      { key: 'baseFacts', label: 'Notes - Base Facts', type: 'textarea' },
    ]
  },
  {
    id: 'observations',
    title: 'Planning Assumptions & Observations',
    fields: [
      { key: 'currentCashFlow', label: 'Current Cash-Flow', type: 'select', options: ['Surplus', 'Deficit', 'Break-even'] },
      { key: 'maxingRetirement', label: 'Maxing Employer Retirement Plan?', type: 'select', options: ['Yes', 'No', 'N/A'] },
      { key: 'excessCash', label: 'Where is Excess Cash Going?', width: 'wide' },
      { key: 'pullingFrom', label: 'Where Pulling Money From?', width: 'wide' },
      { key: 'cashReserves', label: 'How Much Extra Cash Reserves?', width: 'wide' },
      { key: 'onTrack', label: 'On Track to Retire?', type: 'select', options: ['Yes', 'No', 'Unknown'] },
      { key: 'observationNotes', label: 'Observation Notes', type: 'textarea', prefillFrom: 'goals.goals' },
    ]
  },
  {
    id: 'taxes',
    title: 'Tax Observations & Strategies',
    fields: [
      { key: 'filingStatus', label: 'Filing Status', type: 'select', options: ['Single', 'Married Filing Jointly', 'Married Filing Separately', 'Head of Household'] },
      { key: 'taxDifferences', label: 'Differences from Prior Tax Return', width: 'wide' },
      { key: 'taxBeforeRmd', label: 'Taxes Before RMD Start', type: 'select', options: ['Higher', 'Lower', 'Same'] },
      { key: 'taxAfterRmd', label: 'Taxes After RMD Start', type: 'select', options: ['Higher', 'Lower', 'Same'] },
      { key: 'charitablyInclined', label: 'Charitably Inclined?', type: 'select', options: ['Yes', 'No'] },
      { key: 'rothConversion', label: 'Roth Conversion Strategy', width: 'wide' },
      { key: 'capitalGains', label: 'Capital Gains Strategy', width: 'wide' },
      { key: 'charitableStrategies', label: 'Charitable Strategies', width: 'wide' },
      { key: 'otherTax', label: 'Other Tax Strategies', width: 'wide' },
    ]
  },
  {
    id: 'retirement',
    title: 'Cash Flow / Retirement Planning',
    fields: [
      { key: 'cashReservesStrategy', label: 'Cash Reserves Strategy', width: 'wide' },
      { key: 'ssClaiming', label: 'SS Claiming Strategy', width: 'wide' },
      { key: 'otherCashFlow', label: 'Other Cash Flow Strategies', width: 'wide' },
    ]
  },
  {
    id: 'investment',
    title: 'Investment Planning',
    fields: [
      { key: 'pureIps', label: 'Pure IPS' },
      { key: 'outside401k', label: 'Outside 401(k) Request', type: 'select', options: ['Requested', 'Not Requested'] },
      { key: 'annuityStrategy', label: 'Annuity Strategy', width: 'wide' },
      { key: 'otherInvestment', label: 'Other Investment Requests', width: 'wide' },
    ]
  },
  {
    id: 'insurancePlan',
    title: 'Insurance / Risk Management',
    fields: [
      { key: 'lifeInsurance', label: 'Life Insurance', width: 'wide' },
      { key: 'pAndC', label: 'P&C', width: 'wide' },
      { key: 'ltdDisability', label: 'LTD / Disability', width: 'wide' },
    ]
  },
  {
    id: 'estatePlan',
    title: 'Estate Planning',
    fields: [
      { key: 'estatePlan', label: 'Estate Plan', width: 'wide', prefillFrom: 'goals.estatePlan' },
      { key: 'estateDetails', label: 'Estate Planning Details', width: 'wide' },
      { key: 'estateOptions', label: 'Estate Plan Options', width: 'wide' },
    ]
  },
  {
    id: 'recommendations',
    title: 'Recommendations Summary',
    fields: [
      { key: 'recommendationNotes', label: 'Recommendations', type: 'textarea' },
    ]
  },
  {
    id: 'accountsToEstablish',
    title: 'Accounts to Establish / Use',
    note: 'Structured rows for accounts that need to be opened or used for this plan.',
    tables: [
      {
        key: 'accounts',
        columns: [
          { key: 'accountType', label: 'Account Type' },
          { key: 'owner', label: 'Owner' },
          { key: 'registration', label: 'Registration' },
          { key: 'custodian', label: 'Custodian' },
          { key: 'purpose', label: 'Purpose' },
          { key: 'fundingMethod', label: 'Funding Method' },
          { key: 'beneficiaryNeeded', label: 'Bene. Needed?', type: 'select', options: ['Yes', 'No'] },
          { key: 'notes', label: 'Notes' },
        ]
      }
    ]
  },
];

/* ==========================================================================
   RENDERING ENGINE
   ========================================================================== */

function buildFieldEl(field, sectionId, formType) {
  const widthMap = { 'wide': 'field-wide', 'medium': 'field-medium', 'field': 'field', 'field-2': 'field-2' };
  let cls = widthMap[field.width] || 'field';
  if (field.type === 'textarea') cls = 'field-full';
  if (field.type === 'computed') cls += ' field-computed';

  const wrapper = el('label', { className: cls + (formType === 'plan' ? ' flag-red' : '') });

  const row = el('div', { className: 'label-row' });
  row.appendChild(el('span', { className: 'label', textContent: field.label }));

  if (formType === 'plan' && field.prefillFrom) {
    row.appendChild(el('span', { className: 'field-tag tag-prefilled', textContent: 'From PQ' }));
  }

  let input;
  const path = `${sectionId}.${field.key}`;

  if (field.type === 'multiselect') {
    // Render as checkbox group, store as comma-separated string in a hidden input
    input = el('input', { type: 'hidden' });
    input.dataset.path = path;
    input.name = path;
    const cbGroup = el('div', { className: 'multiselect-group' });
    (field.options || []).forEach(opt => {
      const item = el('label', { className: 'multiselect-item' });
      const cb = el('input', { type: 'checkbox' });
      cb.dataset.msOption = opt;
      cb.addEventListener('change', () => {
        const checked = Array.from(cbGroup.querySelectorAll('input[type="checkbox"]:checked'))
          .map(c => c.dataset.msOption);
        input.value = checked.join(', ');
        input.dispatchEvent(new Event('input', { bubbles: true }));
      });
      item.append(cb, document.createTextNode(' ' + opt));
      cbGroup.appendChild(item);
    });
    wrapper.append(row, cbGroup, input);
    // Hydrate checkboxes from stored value after a tick (fillForm sets input.value)
    requestAnimationFrame(() => {
      const vals = (input.value || '').split(',').map(v => v.trim()).filter(Boolean);
      cbGroup.querySelectorAll('input[type="checkbox"]').forEach(cb => {
        cb.checked = vals.includes(cb.dataset.msOption);
      });
    });
    return wrapper;
  } else if (field.type === 'textarea') {
    input = el('textarea', { rows: '3' });
  } else if (field.type === 'select') {
    input = el('select');
    input.appendChild(el('option', { value: '', textContent: '— Select —' }));
    (field.options || []).forEach(o => input.appendChild(el('option', { value: o, textContent: o })));
  } else if (field.type === 'computed') {
    input = el('input', { type: 'text', readonly: 'readonly', tabindex: '-1' });
  } else if (field.type === 'currency') {
    input = el('input', { type: 'text' });
    input.dataset.currencyType = 'currency';
    input.inputMode = 'decimal';
  } else {
    input = el('input', { type: field.type || 'text' });
    if (field.type === 'number') input.step = 'any';
  }

  input.dataset.path = path;
  input.name = path;
  if (field.required) input.required = true;
  if (field.type === 'currency') attachCurrencyBehavior(input);

  wrapper.append(row, input);
  return wrapper;
}

function buildTableEl(sectionId, tableDef, existingData) {
  // Guard: existingData must be an array or null; purge sparse/null entries
  if (existingData && !Array.isArray(existingData)) existingData = null;
  if (existingData) existingData = existingData.filter(r => r != null);
  const block = el('div');

  const tools = el('div', { className: 'table-tools' });
  if (tableDef.title) tools.appendChild(el('div', { className: 'table-title', textContent: tableDef.title }));
  const addBtn = el('button', { type: 'button', className: 'btn-link', textContent: '+ Add Row' });
  tools.appendChild(addBtn);

  const wrap = el('div', { className: 'table-wrap' });
  const table = el('table');
  const thead = el('thead');
  const headRow = el('tr');
  tableDef.columns.forEach(c => {
    const th = el('th', { textContent: c.label });
    if (c.hidden) th.style.display = 'none';
    headRow.appendChild(th);
  });
  headRow.appendChild(el('th', { textContent: '', style: { width: '40px' } }));
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = el('tbody', { dataset: { tablePath: `${sectionId}.${tableDef.key}` } });
  table.appendChild(tbody);
  wrap.appendChild(table);
  block.append(tools, wrap);

  function addRow(seed = {}) {
    const idx = tbody.children.length;
    const tr = el('tr');
    tableDef.columns.forEach(col => {
      const td = el('td');
      let inp;
      if (col.hidden) {
        inp = el('input', { type: 'hidden' });
      } else if (col.type === 'select') {
        inp = el('select');
        inp.appendChild(el('option', { value: '', textContent: '—' }));
        (col.options || []).forEach(o => inp.appendChild(el('option', { value: o, textContent: o })));
      } else if (col.type === 'datalist') {
        // Combobox: free-text input backed by a datalist for suggestions
        const listId = `dl-${sectionId}-${tableDef.key}-${col.key}`;
        // Create datalist once in document if not present
        if (!document.getElementById(listId)) {
          const dl = document.createElement('datalist');
          dl.id = listId;
          (col.options || []).forEach(o => {
            const opt = document.createElement('option');
            opt.value = o;
            dl.appendChild(opt);
          });
          document.body.appendChild(dl);
        } else {
          // Refresh options in case names changed
          const dl = document.getElementById(listId);
          dl.innerHTML = '';
          (col.options || []).forEach(o => {
            const opt = document.createElement('option');
            opt.value = o;
            dl.appendChild(opt);
          });
        }
        inp = el('input', { type: 'text' });
        inp.setAttribute('list', listId);
      } else if (col.type === 'currency') {
        inp = el('input', { type: 'text' });
        inp.dataset.currencyType = 'currency';
        inp.inputMode = 'decimal';
      } else if (col.type === 'percent') {
        inp = el('input', { type: 'number' });
        inp.step = 'any';
        inp.min = '0';
      } else {
        inp = el('input', { type: col.type || 'text' });
      }
      inp.dataset.path = `${sectionId}.${tableDef.key}[${idx}].${col.key}`;
      inp.name = inp.dataset.path;
      if (seed[col.key] != null) {
        inp.value = inp.dataset.currencyType === 'currency' ? formatCommas(seed[col.key]) : seed[col.key];
      }
      if (inp.dataset.currencyType === 'currency') attachCurrencyBehavior(inp);
      if (col.hidden) td.style.display = 'none';

      // Wrap with prefix/suffix symbol if needed
      if (!col.hidden && (col.type === 'currency' || col.type === 'percent')) {
        const wrap = el('div', { className: 'input-symbol-wrap' });
        if (col.type === 'currency') wrap.appendChild(el('span', { className: 'input-symbol input-symbol--pre', textContent: '$' }));
        wrap.appendChild(inp);
        if (col.type === 'percent') wrap.appendChild(el('span', { className: 'input-symbol input-symbol--post', textContent: '%' }));
        td.appendChild(wrap);
      } else {
        td.appendChild(inp);
      }
      tr.appendChild(td);
    });
    const actTd = el('td', { className: 'row-actions' });
    actTd.appendChild(el('button', {
      type: 'button', className: 'btn-icon', textContent: '\u00d7',
      onClick: () => { tr.remove(); reindex(); recalcPQComputed(); }
    }));
    tr.appendChild(actTd);
    tbody.appendChild(tr);
  }

  function reindex() {
    Array.from(tbody.children).forEach((row, ri) => {
      const inputs = row.querySelectorAll('[data-path]');
      tableDef.columns.forEach((col, ci) => {
        if (inputs[ci]) {
          const p = `${sectionId}.${tableDef.key}[${ri}].${col.key}`;
          inputs[ci].dataset.path = p;
          inputs[ci].name = p;
        }
      });
    });
  }

  function editableCells(tr) {
    return Array.from(tr.querySelectorAll('input, select, textarea')).filter(node => {
      if (!node || node.disabled || node.readOnly) return false;
      if (node.tagName === 'INPUT' && node.type === 'hidden') return false;
      return true;
    });
  }

  function rowHasAnyData(tr) {
    return editableCells(tr).some(node => String(node.value || '').trim() !== '');
  }

  // Expose addRow/reindex so external sync logic can add/update rows
  tbody._addRow = addRow;
  tbody._reindex = reindex;
  tbody._tableDef = tableDef;

  // Enter key moves to next editable cell; at row end, add a new row.
  tbody.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    const target = e.target;
    if (!(target instanceof HTMLElement)) return;
    const tr = target.closest('tr');
    if (!tr || !tbody.contains(tr)) return;

    const cells = editableCells(tr);
    const idx = cells.indexOf(target);
    if (idx === -1) return;

    e.preventDefault();
    if (idx < cells.length - 1) {
      cells[idx + 1].focus();
      return;
    }

    addRow();
    reindex();
    const nextRow = tbody.lastElementChild;
    if (!nextRow) return;
    const nextCells = editableCells(nextRow);
    if (nextCells[0]) nextCells[0].focus();
  });

  // Remove rows that are left completely blank when focus leaves that row.
  tbody.addEventListener('focusout', (e) => {
    const target = e.target;
    if (!(target instanceof HTMLElement)) return;
    const tr = target.closest('tr');
    if (!tr || !tbody.contains(tr)) return;

    setTimeout(() => {
      if (!tbody.contains(tr)) return;
      if (tr.contains(document.activeElement)) return;
      if (rowHasAnyData(tr)) return;
      tr.remove();
      reindex();
    }, 0);
  });

  // Seed rows
  const starters = tableDef.starterRows || [];
  const dataRows = existingData || [];
  const rows = dataRows.length ? dataRows : starters;
  rows.forEach(r => addRow(r));
  if (!rows.length) addRow();

  addBtn.addEventListener('click', () => addRow());
  return block;
}

/* ---------- COLLECT / FILL HELPERS ---------- */
function collectForm(formEl) {
  const data = {};
  formEl.querySelectorAll('[data-path]').forEach(inp => {
    // Skip checkboxes (handled by pqState) and radio buttons
    if (inp.type === 'checkbox' || inp.type === 'radio') return;
    const raw = inp.value;
    if (raw === '' || raw == null) return;
    let val;
    if (inp.dataset.currencyType === 'currency') {
      const n = parseCommas(raw);
      val = isNaN(n) ? raw : n;
    } else if (inp.type === 'number') {
      val = Number(raw);
    } else {
      val = raw;
    }
    setDeep(data, inp.dataset.path, val);
  });
  return data;
}

function fillForm(formEl, data) {
  if (!data) return;
  formEl.querySelectorAll('[data-path]').forEach(inp => {
    const v = getDeep(data, inp.dataset.path);
    if (v == null) return;
    if (inp.dataset.currencyType === 'currency') {
      const n = parseFloat(v);
      inp.value = isNaN(n) ? v : formatCommas(n);
    } else {
      inp.value = v;
    }
  });
}

/* ==========================================================================
   STAGE 1: PQ DISCOVERY INTAKE
   ========================================================================== */

const pqState = {
  hasSpouse: false,
  hasChildren: false,
  assetChecks: {
    ownHome: false, secondaryHome: false, additionalRE: false, ownBusiness: false,
    taxDeferred: false, roth: false, taxable: false, cashCd: false, plan529: false, hsa: false,
  },
  employmentClient1: 'Employed',
  employmentClient2: 'Employed',
  // Track which sections are expanded (by section id)
  openSections: { family: true },
};

// Temp holder for in-flight PQ data (not yet submitted)
let _pqDraft = {};

function snapshotPQForm() {
  const form = document.getElementById('pq-form');
  if (!form) return;
  const collected = collectForm(form);
  // Merge into draft (don't lose data from sections not currently rendered)
  Object.keys(collected).forEach(k => {
    if (typeof collected[k] === 'object' && !Array.isArray(collected[k]) && _pqDraft[k] && typeof _pqDraft[k] === 'object') {
      _pqDraft[k] = { ..._pqDraft[k], ...collected[k] };
    } else {
      _pqDraft[k] = collected[k];
    }
  });
}

/* ---------- REAL-ESTATE → LIABILITIES LIVE SYNC ---------- */
function syncREToLiabilities() {
  const form = document.getElementById('pq-form');
  if (!form) return;
  const reBody  = form.querySelector('tbody[data-table-path="assets.realEstate"]');
  const liabBody = form.querySelector('tbody[data-table-path="liabilities.items"]');
  if (!reBody || !liabBody) return;

  // Helper: get field value from a table row by column key suffix
  const rowVal = (tr, key) => tr.querySelector(`[data-path$=".${key}"]`)?.value ?? '';

  // Helper: set a value on a liab row input, formatting currency if needed
  const setRowVal = (tr, key, val) => {
    const inp = tr.querySelector(`[data-path$=".${key}"]`);
    if (!inp) return;
    if (inp.dataset.currencyType === 'currency') {
      const n = parseCommas(val);
      inp.value = isNaN(n) || n === 0 ? '' : formatCommas(n);
    } else {
      inp.value = val ?? '';
    }
  };

  Array.from(reBody.querySelectorAll('tr')).forEach(reTr => {
    const autoKey    = rowVal(reTr, '_autoKey');
    const desc       = rowVal(reTr, 'description');
    const loan       = parseCommas(rowVal(reTr, 'remainingLoan')) || 0;
    const payment    = rowVal(reTr, 'payment');
    const intRate    = rowVal(reTr, 'interestRate');
    const term       = rowVal(reTr, 'term');

    if (!autoKey) return;

    // Find matching liability row by hidden _reKey
    const liabRows = Array.from(liabBody.querySelectorAll('tr'));
    let liabTr = liabRows.find(tr => rowVal(tr, '_reKey') === autoKey);

    if (loan > 0) {
      const liabDesc = desc ? desc + ' Mortgage' : 'RE Loan Mortgage';
      if (!liabTr) {
        // Try to reuse an empty row (no _reKey and all visible fields blank)
        const liabCols = liabBody._tableDef?.columns || [];
        const visibleKeys = liabCols.filter(c => !c.hidden && c.key !== '_reKey').map(c => c.key);
        const emptyRow = Array.from(liabBody.querySelectorAll('tr')).find(tr => {
          if (rowVal(tr, '_reKey')) return false; // already tagged to an RE asset
          return visibleKeys.every(k => {
            const v = rowVal(tr, k);
            return v === '' || v === '—';
          });
        });

        if (emptyRow) {
          // Claim the empty row by writing _reKey and filling it
          setRowVal(emptyRow, '_reKey', autoKey);
          setRowVal(emptyRow, 'description', liabDesc);
          setRowVal(emptyRow, 'amount', String(loan));
          setRowVal(emptyRow, 'payment', payment);
          setRowVal(emptyRow, 'interestRate', intRate);
          setRowVal(emptyRow, 'term', term);
          liabTr = emptyRow;
        } else {
          // No empty row available — add a new one
          liabBody._addRow({ _reKey: autoKey, description: liabDesc, amount: loan, payment, interestRate: intRate, term });
          liabBody._reindex();
          const newRows = Array.from(liabBody.querySelectorAll('tr'));
          liabTr = newRows.find(tr => rowVal(tr, '_reKey') === autoKey);
        }
      } else {
        // Update in place
        setRowVal(liabTr, 'description', liabDesc);
        setRowVal(liabTr, 'amount', String(loan));
        setRowVal(liabTr, 'payment', payment);
        setRowVal(liabTr, 'interestRate', intRate);
        setRowVal(liabTr, 'term', term);
      }
    } else if (liabTr) {
      // Loan cleared — remove the row
      liabTr.remove();
      liabBody._reindex();
    }
  });

  // Trigger expense sync and dashboard recalc so totals stay current
  syncToExpenses();
  recalcPQComputed();
}

function setupRELiabSync() {
  const form = document.getElementById('pq-form');
  const reBody = form?.querySelector('tbody[data-table-path="assets.realEstate"]');
  if (!reBody) return;
  // Listen for any input on the RE table and sync to liabilities
  reBody.addEventListener('input', syncREToLiabilities);
  reBody.addEventListener('change', syncREToLiabilities);
}

/* ---------- LIABILITIES & INSURANCE → EXPENSES LIVE SYNC ---------- */
function syncToExpenses() {
  const form = document.getElementById('pq-form');
  if (!form) return;
  const liabBody = form.querySelector('tbody[data-table-path="liabilities.items"]');
  const insBody  = form.querySelector('tbody[data-table-path="insurance.policies"]');
  const expBody  = form.querySelector('tbody[data-table-path="taxesExpenses.expenses"]');
  if (!expBody) return;

  const rowVal = (tr, key) => tr.querySelector(`[data-path$=".${key}"]`)?.value ?? '';
  const setRowVal = (tr, key, val) => {
    const inp = tr.querySelector(`[data-path$=".${key}"]`);
    if (!inp) return;
    if (inp.dataset.currencyType === 'currency') {
      const n = parseCommas(val);
      inp.value = isNaN(n) || n === 0 ? '' : formatCommas(n);
    } else {
      inp.value = val ?? '';
    }
  };

  // Collect expected auto-sourced expenses from liabilities
  const expected = [];
  if (liabBody) {
    Array.from(liabBody.querySelectorAll('tr')).forEach((tr, i) => {
      const desc = rowVal(tr, 'description');
      const payment = parseCommas(rowVal(tr, 'payment')) || 0;
      if (payment > 0) {
        expected.push({ _source: 'liability', _syncKey: 'liab_' + i, description: desc || 'Loan Payment', amount: payment * 12, notes: 'From liabilities' });
      }
    });
  }

  // Collect expected auto-sourced expenses from insurance
  if (insBody) {
    Array.from(insBody.querySelectorAll('tr')).forEach((tr, i) => {
      const company = rowVal(tr, 'company');
      const premium = parseCommas(rowVal(tr, 'annualPremium')) || 0;
      if (premium > 0) {
        expected.push({ _source: 'insurance', _syncKey: 'ins_' + i, description: 'Insurance Premium - ' + (company || 'Unknown'), amount: premium, notes: 'From insurance' });
      }
    });
  }

  // Current expense rows
  const expRows = Array.from(expBody.querySelectorAll('tr'));

  // Remove stale auto-sourced rows (rows with _source that no longer match)
  const existingAutoRows = expRows.filter(tr => {
    const src = rowVal(tr, '_source');
    return src === 'liability' || src === 'insurance';
  });
  existingAutoRows.forEach(tr => {
    tr.remove();
  });

  // Add all expected auto-sourced rows
  expected.forEach(exp => {
    expBody._addRow(exp);
  });
  expBody._reindex();

  recalcPQComputed();
}

function setupExpenseSync() {
  const form = document.getElementById('pq-form');
  if (!form) return;
  const liabBody = form.querySelector('tbody[data-table-path="liabilities.items"]');
  const insBody  = form.querySelector('tbody[data-table-path="insurance.policies"]');
  if (liabBody) {
    liabBody.addEventListener('input', syncToExpenses);
    liabBody.addEventListener('change', syncToExpenses);
  }
  if (insBody) {
    insBody.addEventListener('input', syncToExpenses);
    insBody.addEventListener('change', syncToExpenses);
  }
}

/* ---------- ASSETS (EBT) → OTHER INCOME LIVE SYNC ---------- */
function syncAssetToOtherIncome() {
  const form = document.getElementById('pq-form');
  if (!form) return;
  const reBody = form.querySelector('tbody[data-table-path="assets.realEstate"]');
  const otherBody = form.querySelector('tbody[data-table-path="income.other"]');
  if (!otherBody) return;

  const rowVal = (tr, key) => tr.querySelector(`[data-path$=".${key}"]`)?.value ?? '';
  const setRowVal = (tr, key, val) => {
    const inp = tr.querySelector(`[data-path$=".${key}"]`);
    if (!inp) return;
    if (inp.dataset.currencyType === 'currency') {
      const n = parseCommas(val);
      inp.value = isNaN(n) || n === 0 ? '' : formatCommas(n);
    } else {
      inp.value = val ?? '';
    }
  };

  // Collect expected auto-sourced rows from assets with EBT income
  const expected = [];
  if (reBody) {
    Array.from(reBody.querySelectorAll('tr')).forEach(tr => {
      const desc = rowVal(tr, 'description');
      const ebt = parseCommas(rowVal(tr, 'incomeEBT')) || 0;
      const ownership = rowVal(tr, 'ownership');
      if (ebt > 0) {
        expected.push({ _source: 'asset', owner: ownership, description: desc || 'Asset Income', annualAmount: ebt, notes: 'From assets' });
      }
    });
  }

  // Remove existing auto-sourced rows
  Array.from(otherBody.querySelectorAll('tr')).forEach(tr => {
    const src = tr.querySelector('input[data-path$="._source"]');
    if (src?.value === 'asset') tr.remove();
  });

  // Fill blank rows first, then add new rows for the rest
  const tableDef = otherBody._tableDef;
  const visibleKeys = tableDef?.columns?.filter(c => !c.hidden).map(c => c.key) || [];

  expected.forEach(exp => {
    // Find a blank row (no _source and all visible fields empty)
    const blankRow = Array.from(otherBody.querySelectorAll('tr')).find(tr => {
      const src = rowVal(tr, '_source');
      if (src) return false;
      return visibleKeys.every(k => {
        const v = rowVal(tr, k);
        return v === '' || v === '—';
      });
    });

    if (blankRow) {
      // Fill the blank row
      Object.entries(exp).forEach(([k, v]) => setRowVal(blankRow, k, String(v)));
    } else {
      otherBody._addRow(exp);
    }
  });

  otherBody._reindex();
  recalcPQComputed();
}

function setupAssetIncomeSync() {
  const form = document.getElementById('pq-form');
  if (!form) return;
  const reBody = form.querySelector('tbody[data-table-path="assets.realEstate"]');
  if (reBody) {
    reBody.addEventListener('input', syncAssetToOtherIncome);
    reBody.addEventListener('change', syncAssetToOtherIncome);
  }
}

/* ---------- FAMILY/EMPLOYMENT → EMPLOYMENT INCOME LIVE SYNC ---------- */
function syncEmploymentFields() {
  const form = document.getElementById('pq-form');
  if (!form) return;
  const empBody = form.querySelector('tbody[data-table-path="income.employment"]');
  if (!empBody) return;

  const c1Name = [
    (form.querySelector('[data-path="family.client1FirstName"]')?.value || '').trim(),
    (form.querySelector('[data-path="family.client1LastName"]')?.value || '').trim()
  ].filter(Boolean).join(' ');
  const c2Name = [
    (form.querySelector('[data-path="family.client2FirstName"]')?.value || '').trim(),
    (form.querySelector('[data-path="family.client2LastName"]')?.value || '').trim()
  ].filter(Boolean).join(' ');
  const c1Employer = (form.querySelector('[data-path="employment.client1.employer"]')?.value || '').trim();
  const c2Employer = (form.querySelector('[data-path="employment.client2.employer"]')?.value || '').trim();

  empBody.querySelectorAll('tr').forEach(row => {
    const srcInp = row.querySelector('input[data-path$="._source"]');
    const ownerInp = row.querySelector('input[data-path$=".owner"]');
    const descInp = row.querySelector('input[data-path$=".description"]');
    if (!srcInp || !ownerInp) return;
    if (srcInp.value === 'client1') {
      ownerInp.value = c1Name;
      if (descInp) descInp.value = c1Employer;
    } else if (srcInp.value === 'client2') {
      ownerInp.value = c2Name;
      if (descInp) descInp.value = c2Employer;
    }
  });
}

function setupEmploymentNameSync() {
  const form = document.getElementById('pq-form');
  if (!form) return;
  const watchPaths = [
    'family.client1FirstName', 'family.client1LastName',
    'family.client2FirstName', 'family.client2LastName',
    'employment.client1.employer', 'employment.client2.employer'
  ];
  watchPaths.forEach(path => {
    const inp = form.querySelector(`[data-path="${path}"]`);
    if (inp) {
      inp.addEventListener('input', syncEmploymentFields);
    }
  });
}

function renderPQ() {
  // Snapshot any in-flight form data before rebuilding DOM
  snapshotPQForm();

  const host = document.getElementById('pq-form-sections');
  host.innerHTML = '';
  const store = getStore();
  const saved = store.pqTemplate || {};
  // Merge: saved store data as base, then overlay the in-flight draft
  const pq = deepMerge(deepClone(saved), _pqDraft);

  // Restore state from existing data
  if (pq.family?.client2FirstName || pq.family?.hasSpouse) pqState.hasSpouse = true;
  if (
    pq.family?.hasChildren ||
    (pq.family?.children && pq.family.children.length) ||
    (pq.family?.grandchildren && pq.family.grandchildren.length)
  ) pqState.hasChildren = true;
  if (pq.employment?.client1?.status) pqState.employmentClient1 = pq.employment.client1.status;
  if (pq.employment?.client2?.status) pqState.employmentClient2 = pq.employment.client2.status;
  // Restore asset checks from dedicated namespace (never conflicts with table data)
  if (pq._assetChecks) {
    Object.keys(pqState.assetChecks).forEach(k => {
      if (typeof pq._assetChecks[k] === 'boolean') pqState.assetChecks[k] = pq._assetChecks[k];
    });
  }

  pqSections.forEach((section) => {
    if (section.customRenderer) {
      host.appendChild(buildPQCustomSection(section, pq));
    } else {
      host.appendChild(buildPQSection(section, pq));
    }
  });

  // Hydrate form with saved data
  const form = document.getElementById('pq-form');
  fillForm(form, pq);
  recalcPQComputed();
  renderPQDashboard();
  // Wire live RE → Liabilities sync after DOM is fully built
  setupRELiabSync();
  // Wire live Liabilities & Insurance → Expenses sync
  setupExpenseSync();
  // Wire live Assets (EBT) → Other Income sync
  setupAssetIncomeSync();
  // Wire live Family Names → Employment Income owner sync
  setupEmploymentNameSync();
  // Recalc dashboard on any form input (debounced to avoid excessive redraws)
  setupFormRecalc();
}

let _recalcTimer = null;
function setupFormRecalc() {
  const form = document.getElementById('pq-form');
  if (!form || form._recalcWired) return;
  form._recalcWired = true;
  form.addEventListener('input', () => {
    clearTimeout(_recalcTimer);
    _recalcTimer = setTimeout(recalcPQComputed, 150);
  });
}

function buildPQSection(section, pqData) {
  const isOpen = !!pqState.openSections[section.id];
  const wrapper = el('section', { className: 'form-section' });

  const header = el('div', { className: 'section-header' });
  header.appendChild(el('h2', { textContent: section.title }));
  const toggleBtn = el('button', { type: 'button', className: 'toggle-btn', textContent: isOpen ? '\u2212' : '+' });
  header.appendChild(toggleBtn);
  wrapper.appendChild(header);

  const content = el('div', { className: 'section-content' + (isOpen ? '' : ' collapsed') });

  if (section.note) content.appendChild(el('p', { className: 'section-note', textContent: section.note }));

  // Regular fields
  if (section.fields?.length) {
    const grid = el('div', { className: 'grid' });
    section.fields.forEach(f => {
      if (f.showIf && !pqState[f.showIf]) return;
      grid.appendChild(buildFieldEl(f, section.id, 'pq'));
    });
    content.appendChild(grid);
  }

  // Conditional blocks (spouse, children)
  if (section.conditionalBlocks) {
    section.conditionalBlocks.forEach(block => {
      const containerClass = block.id === 'spouse-block' ? 'subsection subsection-spouse' : 'subsection';
      const container = el('div', { className: containerClass, id: `pq-${block.id}` });
      const btn = el('button', {
        type: 'button', className: 'btn-mini',
        textContent: pqState[block.flag] ? block.hideLabel : block.toggleLabel
      });
      const blockContent = el('div', {
        className: 'conditional-section' + (pqState[block.flag] ? '' : ' hidden'),
        style: { marginTop: '10px' }
      });
      const grid = el('div', { className: 'grid' });
      block.fields.forEach(f => grid.appendChild(buildFieldEl(f, section.id, 'pq')));
      blockContent.appendChild(grid);

      btn.addEventListener('click', () => {
        pqState[block.flag] = !pqState[block.flag];
        btn.textContent = pqState[block.flag] ? block.hideLabel : block.toggleLabel;
        blockContent.classList.toggle('hidden', !pqState[block.flag]);
        pqState.openSections[section.id] = true; // keep parent open
        renderPQ();
        focusFirstInput(section.id);
      });

      container.append(btn, blockContent);
      content.appendChild(container);
    });
  }

  // After conditional (children table)
  if (section.afterConditional) {
    section.afterConditional.forEach(block => {
      const container = el('div', { className: 'subsection', id: `pq-${block.id}` });
      const btn = el('button', {
        type: 'button', className: 'btn-mini',
        textContent: pqState[block.flag] ? block.hideLabel : block.toggleLabel
      });
      const blockContent = el('div', {
        className: 'conditional-section' + (pqState[block.flag] ? '' : ' hidden'),
        style: { marginTop: '10px' }
      });

      if (block.tables?.length) {
        const tablesWrap = el('div', { className: 'split-tables two-up' });
        block.tables.forEach(tbl => {
          let existingRows = getDeep(pqData, `${section.id}.${tbl.key}`);
          // Legacy compatibility: older data stored both children/grandchildren
          // in family.children with a relationship field.
          if (section.id === 'family') {
            const legacyRows = getDeep(pqData, 'family.children') || [];
            const hasLegacyRelationship = Array.isArray(legacyRows) && legacyRows.some(r => r && Object.prototype.hasOwnProperty.call(r, 'relationship'));
            if (tbl.key === 'children') {
              if ((!existingRows || !existingRows.length) || hasLegacyRelationship) {
                existingRows = legacyRows.filter(r => r?.relationship !== 'Grandchild');
              }
            } else if (tbl.key === 'grandchildren') {
              if (!existingRows || !existingRows.length) {
                existingRows = legacyRows.filter(r => r?.relationship === 'Grandchild');
              }
            }
          }
          tablesWrap.appendChild(buildTableEl(section.id, tbl, existingRows));
        });
        blockContent.appendChild(tablesWrap);
      } else if (block.table) {
        const existingRows = getDeep(pqData, `${section.id}.${block.table.key}`);
        blockContent.appendChild(buildTableEl(section.id, block.table, existingRows));
      }
      if (block.summaryFields) {
        if (section.id === 'family' && block.id === 'children-block' && block.tables?.length === 2) {
          const totalsWrap = el('div', { className: 'split-tables two-up children-totals', style: { marginTop: '10px' } });
          const left = el('div');
          const right = el('div');
          const numChildrenField = block.summaryFields.find(f => f.key === 'numChildren');
          const numGrandchildrenField = block.summaryFields.find(f => f.key === 'numGrandchildren');
          if (numChildrenField) left.appendChild(buildFieldEl(numChildrenField, section.id, 'pq'));
          if (numGrandchildrenField) right.appendChild(buildFieldEl(numGrandchildrenField, section.id, 'pq'));
          totalsWrap.append(left, right);
          blockContent.appendChild(totalsWrap);
        } else {
          const grid = el('div', { className: 'grid', style: { marginTop: '10px' } });
          block.summaryFields.forEach(f => grid.appendChild(buildFieldEl(f, section.id, 'pq')));
          blockContent.appendChild(grid);
        }
      }

      btn.addEventListener('click', () => {
        pqState[block.flag] = !pqState[block.flag];
        btn.textContent = pqState[block.flag] ? block.hideLabel : block.toggleLabel;
        blockContent.classList.toggle('hidden', !pqState[block.flag]);
      });

      container.append(btn, blockContent);
      content.appendChild(container);
    });
  }

  // Footer fields
  if (section.footerFields) {
    const grid = el('div', { className: 'grid', style: { marginTop: '10px' } });
    section.footerFields.forEach(f => grid.appendChild(buildFieldEl(f, section.id, 'pq')));
    content.appendChild(grid);
  }

  // Tables
  if (section.tables) {
    const ownerOptions = getOwnerNameOptions(pqData);
    section.tables.forEach(t => {
      const existing = getDeep(pqData, `${section.id}.${t.key}`);
      const tableDef = section.id === 'insurance'
        ? {
            ...t,
            columns: (t.columns || []).map(col => (
              col.key === 'owner'
                ? { ...col, type: 'datalist', options: ownerOptions }
                : col
            )),
          }
        : t;
      content.appendChild(buildTableEl(section.id, tableDef, existing));
    });
  }

  header.addEventListener('click', () => {
    const isCollapsed = content.classList.contains('collapsed');
    content.classList.toggle('collapsed', !isCollapsed);
    toggleBtn.textContent = isCollapsed ? '\u2212' : '+';
    pqState.openSections[section.id] = isCollapsed; // track state
  });

  wrapper.appendChild(content);
  return wrapper;
}

/* ---------- COLLAPSIBLE SECTION HELPER ---------- */
function buildCollapsibleShell(sectionId, title) {
  const isOpen = !!pqState.openSections[sectionId];
  const wrapper = el('section', { className: 'form-section' });
  const header = el('div', { className: 'section-header' });
  header.appendChild(el('h2', { textContent: title }));
  const toggleBtn = el('button', { type: 'button', className: 'toggle-btn', textContent: isOpen ? '\u2212' : '+' });
  header.appendChild(toggleBtn);
  wrapper.appendChild(header);

  const content = el('div', { className: 'section-content' + (isOpen ? '' : ' collapsed') });

  header.addEventListener('click', () => {
    const isCollapsed = content.classList.contains('collapsed');
    content.classList.toggle('collapsed', !isCollapsed);
    toggleBtn.textContent = isCollapsed ? '\u2212' : '+';
    pqState.openSections[sectionId] = isCollapsed;
  });

  wrapper.appendChild(content);
  return { wrapper, content };
}

/* ---------- CUSTOM PQ SECTION RENDERERS ---------- */

function buildPQCustomSection(section, pqData) {
  const renderers = {
    renderEmploymentSection: () => renderEmploymentSection(section, pqData),
    renderAssetsSection: () => renderAssetsSection(section, pqData),
    renderLiabilitiesSection: () => renderLiabilitiesSection(section, pqData),
    renderIncomeSection: () => renderIncomeSection(section, pqData),
    renderTaxesExpensesSection: () => renderTaxesExpensesSection(section, pqData),
  };
  return renderers[section.customRenderer]();
}

/* ----- EMPLOYMENT ----- */
function renderEmploymentSection(section, pqData) {
  const { wrapper, content } = buildCollapsibleShell(section.id, section.title);

  // Client 1 employment
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Client 1 Employment' }));
  const radio1 = el('div', { className: 'radio-group' });
  ['Employed', 'Retired', 'Not Employed'].forEach(opt => {
    const item = el('label', { className: 'radio-item' });
    const inp = el('input', { type: 'radio', name: 'employment.client1.status', value: opt });
    inp.dataset.path = 'employment.client1.status';
    if (pqState.employmentClient1 === opt) inp.checked = true;
    inp.addEventListener('change', () => {
      pqState.employmentClient1 = opt;
      pqState.openSections[section.id] = true;
      renderPQ();
      focusFirstInput(section.id);
    });
    item.append(inp, document.createTextNode(opt));
    radio1.appendChild(item);
  });
  content.appendChild(radio1);

  const grid1 = el('div', { className: 'grid' });
  if (pqState.employmentClient1 === 'Employed') {
    grid1.appendChild(buildFieldEl({ key: 'client1.jobTitle', label: 'Job Title', width: 'medium' }, 'employment', 'pq'));
    grid1.appendChild(buildFieldEl({ key: 'client1.employer', label: 'Employer', width: 'medium' }, 'employment', 'pq'));
    grid1.appendChild(buildFieldEl({ key: 'client1.years', label: '# of Years', type: 'number', width: 'field' }, 'employment', 'pq'));
    grid1.appendChild(buildFieldEl({ key: 'client1.retirementDate', label: 'Retirement Date', type: 'date', width: 'medium' }, 'employment', 'pq'));
    grid1.appendChild(buildFieldEl({ key: 'client1.businessType', label: 'Type of Business', width: 'medium' }, 'employment', 'pq'));
  } else {
    grid1.appendChild(buildFieldEl({ key: 'client1.lastJob', label: 'Last Job Title', width: 'medium' }, 'employment', 'pq'));
    grid1.appendChild(buildFieldEl({ key: 'client1.lastEmployer', label: 'Last Employer', width: 'medium' }, 'employment', 'pq'));
    grid1.appendChild(buildFieldEl({ key: 'client1.yearsWorked', label: 'Years Worked', type: 'number', width: 'field' }, 'employment', 'pq'));
  }
  content.appendChild(grid1);

  // Client 2 employment (if spouse)
  if (pqState.hasSpouse) {
    content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Spouse Employment', style: { marginTop: '16px' } }));
    const radio2 = el('div', { className: 'radio-group' });
    ['Employed', 'Retired', 'Not Employed'].forEach(opt => {
      const item = el('label', { className: 'radio-item' });
      const inp = el('input', { type: 'radio', name: 'employment.client2.status', value: opt });
      inp.dataset.path = 'employment.client2.status';
      if (pqState.employmentClient2 === opt) inp.checked = true;
      inp.addEventListener('change', () => {
        pqState.employmentClient2 = opt;
        pqState.openSections[section.id] = true;
        renderPQ();
        focusFirstInput(section.id);
      });
      item.append(inp, document.createTextNode(opt));
      radio2.appendChild(item);
    });
    content.appendChild(radio2);

    const grid2 = el('div', { className: 'grid' });
    if (pqState.employmentClient2 === 'Employed') {
      grid2.appendChild(buildFieldEl({ key: 'client2.jobTitle', label: 'Job Title', width: 'medium' }, 'employment', 'pq'));
      grid2.appendChild(buildFieldEl({ key: 'client2.employer', label: 'Employer', width: 'medium' }, 'employment', 'pq'));
      grid2.appendChild(buildFieldEl({ key: 'client2.years', label: '# of Years', type: 'number', width: 'field' }, 'employment', 'pq'));
      grid2.appendChild(buildFieldEl({ key: 'client2.retirementDate', label: 'Retirement Date', type: 'date', width: 'medium' }, 'employment', 'pq'));
      grid2.appendChild(buildFieldEl({ key: 'client2.businessType', label: 'Type of Business', width: 'medium' }, 'employment', 'pq'));
    } else {
      grid2.appendChild(buildFieldEl({ key: 'client2.lastJob', label: 'Last Job Title', width: 'medium' }, 'employment', 'pq'));
      grid2.appendChild(buildFieldEl({ key: 'client2.lastEmployer', label: 'Last Employer', width: 'medium' }, 'employment', 'pq'));
      grid2.appendChild(buildFieldEl({ key: 'client2.yearsWorked', label: 'Years Worked', type: 'number', width: 'field' }, 'employment', 'pq'));
    }
    content.appendChild(grid2);
  }

  return wrapper;
}

/* ----- ASSETS ----- */
function renderAssetsSection(section, pqData) {
  const { wrapper, content } = buildCollapsibleShell(section.id, section.title);

  // --- Real Estate / Business subsection ---
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Real Estate / Business' }));
  const reChecks = el('div', { className: 'checkbox-group' });
  const reCheckDefs = [
    { key: 'ownHome', label: 'Own your home?' },
    { key: 'secondaryHome', label: 'Secondary home(s)?' },
    { key: 'additionalRE', label: 'Additional real estate?' },
    { key: 'ownBusiness', label: 'Own a business?' },
  ];
  reCheckDefs.forEach(def => {
    const item = el('label', { className: 'checkbox-item' });
    const cb = el('input', { type: 'checkbox' });
    cb.checked = pqState.assetChecks[def.key];
    cb.dataset.path = `assets.${def.key}`;
    cb.addEventListener('change', () => {
      pqState.assetChecks[def.key] = cb.checked;
      _pqDraft._assetChecks = _pqDraft._assetChecks || {};
      _pqDraft._assetChecks[def.key] = cb.checked;
      pqState.openSections[section.id] = true;
      renderPQ();
      scrollSectionIntoView(section.id);
    });
    item.append(cb, document.createTextNode(def.label));
    reChecks.appendChild(item);
  });
  content.appendChild(reChecks);

  // Real estate table – smart merge: add rows for checked boxes, remove for
  // unchecked boxes only when no data has been entered beyond description.
  const checkToDesc = [
    { key: 'ownHome',       desc: 'Primary Residence' },
    { key: 'secondaryHome', desc: 'Secondary Residence', legacyDescs: ['Secondary Home'] },
    { key: 'additionalRE',  desc: 'Additional Real Estate' },
    { key: 'ownBusiness',   desc: 'Business' },
  ];

  const ownershipOptions = getOwnerNameOptions(pqData);

  const reColumns = [
    { key: '_autoKey',      label: '',               hidden: true },
    { key: 'description',   label: 'Description' },
    { key: 'marketValue',   label: 'Market Value',   type: 'currency' },
    { key: 'remainingLoan', label: 'Remaining Loan', type: 'currency' },
    { key: 'interestRate',  label: 'Int. Rate',      type: 'percent' },
    { key: 'term',          label: 'Term (years)' },
    { key: 'payment',       label: 'Payment',        type: 'currency' },
    { key: 'yearAcquired',  label: 'Year Acquired' },
    { key: 'costBasis',     label: 'Cost Basis',     type: 'currency' },
    { key: 'ownership',     label: 'Ownership',      type: 'datalist', options: ownershipOptions },
    { key: 'incomeEBT',     label: 'Income (EBT)',   type: 'currency' },
  ];
  const dataKeys = reColumns.map(c => c.key).filter(k => k !== 'description' && k !== '_autoKey');

  let existingRE = getDeep(pqData, 'assets.realEstate');
  if (existingRE && !Array.isArray(existingRE)) existingRE = null;
  const mergedRows = existingRE ? [...existingRE] : [];

  const normalizeDesc = v => String(v || '').trim().toLowerCase();
  const valueHasData = v => {
    if (v == null) return false;
    if (typeof v === 'string') return v.trim() !== '';
    return true;
  };
  const rowHasData = row => dataKeys.some(k => valueHasData(row?.[k]));
  // One-time migration for older data that predates _autoKey.
  // Only migrate when there is exactly one unkeyed candidate for a checkbox.
  checkToDesc.forEach(def => {
    const alreadyMapped = mergedRows.some(r => r?._autoKey === def.key);
    if (alreadyMapped) return;
    const targets = [def.desc, ...(def.legacyDescs || [])].map(normalizeDesc);
    const candidates = mergedRows
      .map((row, idx) => ({ row, idx }))
      .filter(({ row }) => !row?._autoKey && targets.includes(normalizeDesc(row?.description)));
    if (candidates.length === 1) {
      const idx = candidates[0].idx;
      mergedRows[idx] = { ...mergedRows[idx], _autoKey: def.key };
    }
  });

  // Repair prior buggy states: keep only one row per auto key.
  // Prefer the row with the most non-description user data.
  const scoreRowData = row => dataKeys.reduce((n, k) => n + (valueHasData(row?.[k]) ? 1 : 0), 0);
  checkToDesc.forEach(def => {
    const keyed = mergedRows
      .map((row, idx) => ({ row, idx }))
      .filter(({ row }) => row?._autoKey === def.key);
    if (keyed.length <= 1) return;
    let keepIdx = keyed[0].idx;
    let bestScore = scoreRowData(keyed[0].row);
    keyed.slice(1).forEach(({ row, idx }) => {
      const s = scoreRowData(row);
      if (s > bestScore) {
        bestScore = s;
        keepIdx = idx;
      }
    });
    for (let i = mergedRows.length - 1; i >= 0; i--) {
      if (mergedRows[i]?._autoKey === def.key && i !== keepIdx) mergedRows.splice(i, 1);
    }
  });

  const findMappedRowIndex = key => mergedRows.findIndex(r => r?._autoKey === key);

  checkToDesc.forEach(def => {
    const idx = findMappedRowIndex(def.key);
    if (pqState.assetChecks[def.key]) {
      if (idx === -1) {
        mergedRows.push({ _autoKey: def.key, description: def.desc });
      } else {
        mergedRows[idx] = { ...mergedRows[idx], _autoKey: def.key, description: def.desc };
      }
    } else if (idx !== -1 && !rowHasData(mergedRows[idx])) {
      mergedRows.splice(idx, 1);
    }
  });

  const anyREChecked = checkToDesc.some(({ key }) => pqState.assetChecks[key]);

  // Write mergedRows back so fillForm (which runs after all sections) uses
  // the correct indexed data instead of the stale pre-merge snapshot.
  if (!pqData.assets) pqData.assets = {};
  pqData.assets.realEstate = mergedRows;
  if (!_pqDraft.assets) _pqDraft.assets = {};
  _pqDraft.assets.realEstate = mergedRows;

  if (anyREChecked || mergedRows.length) {
    const reTable = {
      key: 'realEstate',
      title: 'Real Estate & Business Assets',
      columns: reColumns,
      starterRows: [],
    };
    content.appendChild(buildTableEl('assets', reTable, mergedRows.length ? mergedRows : null));
  }

  // --- Investment Accounts subsection ---
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Investment Accounts', style: { marginTop: '18px' } }));
  const invChecks = el('div', { className: 'checkbox-group' });
  const invCheckDefs = [
    { key: 'taxDeferred', label: 'Tax-Deferred Retirement' },
    { key: 'roth', label: 'Tax-Free Roth' },
    { key: 'taxable', label: 'Taxable Non-Retirement' },
    { key: 'cashCd', label: "Cash & CD's" },
    { key: 'plan529', label: '529 Plans' },
    { key: 'hsa', label: 'HSA' },
  ];
  invCheckDefs.forEach(def => {
    const item = el('label', { className: 'checkbox-item' });
    const cb = el('input', { type: 'checkbox' });
    cb.checked = pqState.assetChecks[def.key];
    cb.dataset.path = `assets.${def.key}`;
    cb.addEventListener('change', () => {
      pqState.assetChecks[def.key] = cb.checked;
      _pqDraft._assetChecks = _pqDraft._assetChecks || {};
      _pqDraft._assetChecks[def.key] = cb.checked;
      pqState.openSections[section.id] = true;
      renderPQ();
      scrollSectionIntoView(section.id);
    });
    item.append(cb, document.createTextNode(def.label));
    invChecks.appendChild(item);
  });
  content.appendChild(invChecks);

  // Retirement accounts table
  if (pqState.assetChecks.taxDeferred) {
    content.appendChild(buildTableEl('assets', {
      key: 'taxDeferred',
      title: 'Tax-Deferred Retirement Accounts',
      columns: [
        { key: 'custodian', label: 'Custodian' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'personalAdditions', label: 'Personal Additions', type: 'currency' },
        { key: 'companyMatch', label: 'Company Match', type: 'currency' },
        { key: 'type', label: 'Type' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
        { key: 'beneficiary', label: 'Beneficiary' },
      ]
    }, getDeep(pqData, 'assets.taxDeferred')));
  }

  if (pqState.assetChecks.roth) {
    content.appendChild(buildTableEl('assets', {
      key: 'roth',
      title: 'Tax-Free Roth Accounts',
      columns: [
        { key: 'custodian', label: 'Custodian' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'personalAdditions', label: 'Personal Additions', type: 'currency' },
        { key: 'companyMatch', label: 'Company Match', type: 'currency' },
        { key: 'type', label: 'Type' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
        { key: 'beneficiary', label: 'Beneficiary' },
      ]
    }, getDeep(pqData, 'assets.roth')));
  }

  if (pqState.assetChecks.taxable) {
    content.appendChild(buildTableEl('assets', {
      key: 'taxable',
      title: 'Taxable Non-Retirement Accounts',
      columns: [
        { key: 'custodian', label: 'Custodian' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'personalAdditions', label: 'Personal Additions', type: 'currency' },
        { key: 'costBasis', label: 'Cost Basis', type: 'currency' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
        { key: 'variableFixed', label: 'Variable / Fixed' },
        { key: 'issueDate', label: 'Issue Date', type: 'date' },
        { key: 'beneficiary', label: 'Beneficiary' },
      ]
    }, getDeep(pqData, 'assets.taxable')));
  }

  // Cash/CD, 529, HSA — simpler tables
  if (pqState.assetChecks.cashCd) {
    content.appendChild(buildTableEl('assets', {
      key: 'cashCd',
      title: "Cash & CD's",
      columns: [
        { key: 'description', label: 'Description' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'interestRate', label: 'Interest Rate' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
      ]
    }, getDeep(pqData, 'assets.cashCd')));
  }

  if (pqState.assetChecks.plan529) {
    content.appendChild(buildTableEl('assets', {
      key: 'plan529',
      title: '529 Plans',
      columns: [
        { key: 'description', label: 'Description' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'interestRate', label: 'Interest Rate' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
      ]
    }, getDeep(pqData, 'assets.plan529')));
  }

  if (pqState.assetChecks.hsa) {
    content.appendChild(buildTableEl('assets', {
      key: 'hsa',
      title: 'HSA Accounts',
      columns: [
        { key: 'description', label: 'Description' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'interestRate', label: 'Interest Rate' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
      ]
    }, getDeep(pqData, 'assets.hsa')));
  }

  return wrapper;
}

/* ----- LIABILITIES ----- */
function renderLiabilitiesSection(section, pqData) {
  const { wrapper, content } = buildCollapsibleShell(section.id, section.title);
  content.appendChild(el('p', { className: 'section-note', textContent: 'Known loan liabilities from the Assets section are prefilled below. Add any additional liabilities.' }));

  // Build starter rows from real estate loans; tag each with _reKey for live sync
  const starterRows = [];
  const reData = getDeep(pqData, 'assets.realEstate') || [];
  reData.forEach(r => {
    if (r?.remainingLoan && parseFloat(r.remainingLoan) > 0) {
      starterRows.push({
        _reKey: r._autoKey || '',
        description: (r.description || 'RE Loan') + ' Mortgage',
        payment: r.payment || '',
        amount: r.remainingLoan,
        interestRate: r.interestRate || '',
        term: r.term || '',
      });
    }
  });

  // Merge: preserve any existing liab rows but update RE-sourced rows from current data
  let existingLiab = getDeep(pqData, 'liabilities.items');
  if (existingLiab && !Array.isArray(existingLiab)) existingLiab = null;
  // Purge sparse/null/undefined entries that arise from collectForm on empty rows
  if (existingLiab) existingLiab = existingLiab.filter(r => r != null && typeof r === 'object');

  // For existing data: update RE-keyed rows in place, keep manual rows untouched
  let seedRows;
  if (existingLiab?.length) {
    // Clone so we can mutate
    const merged = existingLiab.map(r => ({ ...r }));
    starterRows.forEach(sr => {
      const idx = merged.findIndex(r => r._reKey && r._reKey === sr._reKey);
      if (idx !== -1) {
        // Update the RE-sourced row with latest RE values
        merged[idx] = { ...merged[idx], ...sr };
      } else {
        merged.push(sr);
      }
    });
    // Remove RE-keyed rows whose RE source no longer has a loan
    const activeReKeys = new Set(starterRows.map(r => r._reKey).filter(Boolean));
    const cleaned = merged.filter(r => !r._reKey || activeReKeys.has(r._reKey));
    seedRows = cleaned.length ? cleaned : null;
  } else {
    seedRows = starterRows.length ? starterRows : null;
  }

  content.appendChild(buildTableEl('liabilities', {
    key: 'items',
    title: 'All Liabilities',
    columns: [
      { key: '_reKey', label: '', hidden: true },
      { key: 'description', label: 'Description' },
      { key: 'payment', label: 'Payment', type: 'currency' },
      { key: 'amount', label: 'Amount', type: 'currency' },
      { key: 'interestRate', label: 'Int. Rate', type: 'percent' },
      { key: 'term', label: 'Term (years)' },
    ],
    starterRows: [],
  }, seedRows));

  return wrapper;
}

/* ----- INCOME ----- */
function renderIncomeSection(section, pqData) {
  const { wrapper, content } = buildCollapsibleShell(section.id, section.title);
  const ownerOptions = getOwnerNameOptions(pqData);

  // ── 1. Employment Income ──
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Employment Income' }));
  const family = pqData?.family || {};
  const c1Name = [(family.client1FirstName || '').trim(), (family.client1LastName || '').trim()].filter(Boolean).join(' ');
  const c2Name = [(family.client2FirstName || '').trim(), (family.client2LastName || '').trim()].filter(Boolean).join(' ');

  // Auto-sourced employment rows
  const autoEmp = [];
  if (pqState.employmentClient1 === 'Employed') {
    autoEmp.push({ _source: 'client1', owner: c1Name, description: pqData?.employment?.client1?.employer || '' });
  }
  if (pqState.hasSpouse && pqState.employmentClient2 === 'Employed') {
    autoEmp.push({ _source: 'client2', owner: c2Name, description: pqData?.employment?.client2?.employer || '' });
  }

  // Merge: keep auto-sourced rows fresh (update owner/description), preserve manual rows and user-entered amounts
  let existingEmp = getDeep(pqData, 'income.employment');
  let mergedEmp;
  if (existingEmp?.length) {
    const manualRows = existingEmp.filter(r => r && !r._source);
    // For auto rows, preserve user-entered fields (annualAmount, retirementDate, cola)
    const freshAuto = autoEmp.map(auto => {
      const prev = existingEmp.find(r => r?._source === auto._source);
      if (prev) {
        return { ...prev, owner: auto.owner, description: auto.description || prev.description };
      }
      return auto;
    });
    mergedEmp = [...freshAuto, ...manualRows];
  } else {
    mergedEmp = autoEmp.length ? autoEmp : null;
  }

  content.appendChild(buildTableEl('income', {
    key: 'employment',
    title: 'Employment Income',
    columns: [
      { key: '_source', label: '', hidden: true },
      { key: 'owner', label: 'Owner', type: 'datalist', options: ownerOptions },
      { key: 'description', label: 'Description' },
      { key: 'annualAmount', label: 'Annual Amount', type: 'currency' },
      { key: 'retirementDate', label: 'Retirement Date', type: 'date' },
      { key: 'cola', label: 'COLA %', type: 'percent' },
    ],
    starterRows: [],
  }, mergedEmp));

  // ── 2. Social Security ──
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Social Security', style: { marginTop: '18px' } }));
  const existingSS = getDeep(pqData, 'income.socialSecurity');
  content.appendChild(buildTableEl('income', {
    key: 'socialSecurity',
    title: 'Social Security',
    columns: [
      { key: 'owner', label: 'Owner', type: 'datalist', options: ownerOptions },
      { key: 'startingAge', label: 'Starting Age', type: 'number' },
      { key: 'annualAmount', label: 'Annual Amount', type: 'currency' },
      { key: 'fullRetAge', label: 'Full Ret. Age', type: 'number' },
      { key: 'cola', label: 'COLA %', type: 'percent' },
    ],
    starterRows: [],
  }, existingSS));

  // ── 3. Pension ──
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Pension', style: { marginTop: '18px' } }));
  const existingPension = getDeep(pqData, 'income.pension');
  content.appendChild(buildTableEl('income', {
    key: 'pension',
    title: 'Pension',
    columns: [
      { key: 'owner', label: 'Owner', type: 'datalist', options: ownerOptions },
      { key: 'startDate', label: 'Start Date', type: 'date' },
      { key: 'annualAmount', label: 'Annual Amount', type: 'currency' },
      { key: 'survivorBenefit', label: 'Survivor Benefit', type: 'currency' },
      { key: 'cola', label: 'COLA %', type: 'percent' },
    ],
    starterRows: [],
  }, existingPension));

  // ── 4. Other Income ──
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Other Income', style: { marginTop: '18px' } }));

  // Auto-source from assets with Income (EBT)
  const autoOther = [];
  const reData = getDeep(pqData, 'assets.realEstate') || [];
  reData.forEach(r => {
    if (r?.incomeEBT && parseFloat(r.incomeEBT) > 0) {
      autoOther.push({ _source: 'asset', owner: r.ownership || '', description: r.description || 'Asset Income', annualAmount: parseFloat(r.incomeEBT), notes: 'From assets' });
    }
  });

  let existingOther = getDeep(pqData, 'income.other');
  let mergedOther;
  if (existingOther?.length) {
    const manualRows = existingOther.filter(r => r && !r._source);
    mergedOther = [...autoOther, ...manualRows];
  } else {
    mergedOther = autoOther.length ? autoOther : null;
  }

  content.appendChild(buildTableEl('income', {
    key: 'other',
    title: 'Other Income',
    columns: [
      { key: '_source', label: '', hidden: true },
      { key: 'owner', label: 'Owner', type: 'datalist', options: ownerOptions },
      { key: 'description', label: 'Description' },
      { key: 'annualAmount', label: 'Annual Amount', type: 'currency' },
      { key: 'startDate', label: 'Start Date', type: 'date' },
      { key: 'endDate', label: 'End Date', type: 'date' },
      { key: 'cola', label: 'COLA %', type: 'percent' },
      { key: 'notes', label: 'Notes' },
    ],
    starterRows: [],
  }, mergedOther));

  // Total income display
  const totalRow = el('div', { className: 'grid', style: { marginTop: '10px' } });
  totalRow.appendChild(buildFieldEl({ key: 'totalIncome', label: 'Total Income', type: 'computed', width: 'medium' }, 'income', 'pq'));
  content.appendChild(totalRow);

  return wrapper;
}

/* ----- TAXES & EXPENSES ----- */
function renderTaxesExpensesSection(section, pqData) {
  const { wrapper, content } = buildCollapsibleShell(section.id, section.title);

  // Tax subsection
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Prior Year Tax Information' }));
  const taxGrid = el('div', { className: 'grid' });
  const taxFields = [
    { key: 'capLossCarryForward', label: 'Cap Loss Carry Forward', type: 'currency', width: 'medium' },
    { key: 'taxableIncome', label: 'Taxable Income', type: 'currency', width: 'medium' },
    { key: 'standardItemDeduction', label: 'Stnd/Item Deduction', type: 'currency', width: 'medium' },
    { key: 'federalTax', label: 'Federal Taxes Paid', type: 'currency', width: 'medium' },
    { key: 'stateTax', label: 'State Taxes Paid', type: 'currency', width: 'medium' },
    { key: 'ficaTax', label: 'FICA Taxes Paid', type: 'currency', width: 'medium' },
    { key: 'totalTaxesPaid', label: 'Total Taxes Paid', type: 'computed', width: 'medium' },
  ];
  taxFields.forEach(f => taxGrid.appendChild(buildFieldEl(f, 'taxesExpenses', 'pq')));
  content.appendChild(taxGrid);

  // Expenses subsection
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Expenses', style: { marginTop: '18px' } }));
  content.appendChild(el('p', { className: 'section-note', textContent: 'Known loan payments are prefilled from Liabilities.' }));

  // Build auto-sourced expense rows from Liabilities and Insurance
  const autoExpenses = [];
  const liabData = getDeep(pqData, 'liabilities.items') || [];
  liabData.forEach(l => {
    if (l?.payment && parseFloat(l.payment) > 0) {
      autoExpenses.push({ _source: 'liability', description: l.description || 'Loan Payment', amount: parseFloat(l.payment) * 12, notes: 'From liabilities' });
    }
  });

  const insuranceData = getDeep(pqData, 'insurance.policies') || [];
  insuranceData.forEach(p => {
    if (p?.annualPremium && parseFloat(p.annualPremium) > 0) {
      autoExpenses.push({ _source: 'insurance', description: 'Insurance Premium - ' + (p.company || 'Unknown'), amount: parseFloat(p.annualPremium), notes: 'From insurance' });
    }
  });

  // Merge: always keep auto-sourced rows fresh, preserve manual rows
  let existingExp = getDeep(pqData, 'taxesExpenses.expenses');
  let mergedExp;
  if (existingExp?.length) {
    // Keep only manually-entered rows (no _source)
    const manualRows = existingExp.filter(r => r && !r._source);
    mergedExp = [...autoExpenses, ...manualRows];
  } else {
    // First time: auto rows + Living Expenses default
    mergedExp = [...autoExpenses, { description: 'Living Expenses' }];
  }

  content.appendChild(buildTableEl('taxesExpenses', {
    key: 'expenses',
    title: 'Expense Lines',
    columns: [
      { key: '_source', label: '', hidden: true },
      { key: 'description', label: 'Description' },
      { key: 'amount', label: 'Annual Amount', type: 'currency' },
      { key: 'startDate', label: 'Start Date', type: 'date' },
      { key: 'endDate', label: 'End Date', type: 'date' },
      { key: 'cola', label: 'COLA %', type: 'percent' },
      { key: 'notes', label: 'Notes' },
    ],
    starterRows: [],
  }, mergedExp));

  const totalRow = el('div', { className: 'grid', style: { marginTop: '10px' } });
  totalRow.appendChild(buildFieldEl({ key: 'expenseTotal', label: 'Total Expenses', type: 'computed', width: 'medium' }, 'taxesExpenses', 'pq'));
  content.appendChild(totalRow);

  return wrapper;
}

/* ---------- PQ COMPUTED FIELDS ---------- */
function recalcPQComputed() {
  const form = document.getElementById('pq-form');
  if (!form) return;

  // Ages
  const dob1 = form.querySelector('[data-path="family.client1DOB"]');
  const age1 = form.querySelector('[data-path="family.client1Age"]');
  if (dob1 && age1) {
    const v = calcAge(dob1.value);
    if (v !== '') age1.value = v;
  }

  const dob2 = form.querySelector('[data-path="family.client2DOB"]');
  const age2 = form.querySelector('[data-path="family.client2Age"]');
  if (dob2 && age2) {
    const v = calcAge(dob2.value);
    if (v !== '') age2.value = v;
  }

  // Children/grandchildren count from separate tables
  const countRowsWithData = (prefix) => {
    const rowMap = new Map();
    form.querySelectorAll(`[data-path^="${prefix}["]`).forEach(inp => {
      const m = inp.dataset.path.match(/\[(\d+)\]\.(name|age)$/);
      if (!m) return;
      const idx = m[1];
      const key = m[2];
      const row = rowMap.get(idx) || { hasName: false, hasAge: false };
      if (key === 'name' && String(inp.value || '').trim() !== '') row.hasName = true;
      if (key === 'age' && String(inp.value || '').trim() !== '') row.hasAge = true;
      rowMap.set(idx, row);
    });
    let count = 0;
    rowMap.forEach(r => { if (r.hasName || r.hasAge) count++; });
    return count;
  };
  const childCount = countRowsWithData('family.children');
  const grandCount = countRowsWithData('family.grandchildren');
  const numC = form.querySelector('[data-path="family.numChildren"]');
  const numG = form.querySelector('[data-path="family.numGrandchildren"]');
  if (numC && !document.activeElement?.isSameNode(numC)) numC.value = childCount || numC.value;
  if (numG && !document.activeElement?.isSameNode(numG)) numG.value = grandCount || numG.value;

  // Total income (same future-date exclusion logic as dashboard)
  const thisYear = new Date().getFullYear();
  const isFuture = (dateStr) => {
    if (!dateStr) return false;
    const d = new Date(dateStr);
    return !isNaN(d.getTime()) && d.getFullYear() > thisYear;
  };
  let totalIncome = 0;

  // Employment — always included
  form.querySelectorAll('[data-path^="income.employment["]').forEach(inp => {
    if (inp.dataset.path.endsWith('.annualAmount')) totalIncome += parseCommas(inp.value) || 0;
  });

  // Social Security — exclude if startingAge > current age of owner
  const curAge1 = parseFloat(form.querySelector('[data-path="family.client1Age"]')?.value) || 0;
  const curAge2 = parseFloat(form.querySelector('[data-path="family.client2Age"]')?.value) || 0;
  const name1 = [
    (form.querySelector('[data-path="family.client1FirstName"]')?.value || '').trim(),
    (form.querySelector('[data-path="family.client1LastName"]')?.value || '').trim()
  ].filter(Boolean).join(' ');
  const name2 = [
    (form.querySelector('[data-path="family.client2FirstName"]')?.value || '').trim(),
    (form.querySelector('[data-path="family.client2LastName"]')?.value || '').trim()
  ].filter(Boolean).join(' ');
  const ssTb = form.querySelector('tbody[data-table-path="income.socialSecurity"]');
  if (ssTb) {
    ssTb.querySelectorAll('tr').forEach(row => {
      const amt = parseCommas(row.querySelector('input[data-path$=".annualAmount"]')?.value) || 0;
      const startAge = parseFloat(row.querySelector('input[data-path$=".startingAge"]')?.value) || 0;
      const owner = (row.querySelector('input[data-path$=".owner"]')?.value || '').trim();
      let ownerAge = curAge1;
      if (name2 && owner === name2) ownerAge = curAge2;
      if (startAge > 0 && startAge > ownerAge) return;
      totalIncome += amt;
    });
  }

  // Pension — exclude if startDate is future year
  const penTb = form.querySelector('tbody[data-table-path="income.pension"]');
  if (penTb) {
    penTb.querySelectorAll('tr').forEach(row => {
      const amt = parseCommas(row.querySelector('input[data-path$=".annualAmount"]')?.value) || 0;
      const sd = row.querySelector('input[data-path$=".startDate"]')?.value || '';
      if (isFuture(sd)) return;
      totalIncome += amt;
    });
  }

  // Other — exclude if startDate is future year
  const otherTb = form.querySelector('tbody[data-table-path="income.other"]');
  if (otherTb) {
    otherTb.querySelectorAll('tr').forEach(row => {
      const amt = parseCommas(row.querySelector('input[data-path$=".annualAmount"]')?.value) || 0;
      const sd = row.querySelector('input[data-path$=".startDate"]')?.value || '';
      if (isFuture(sd)) return;
      totalIncome += amt;
    });
  }

  const totalIncEl = form.querySelector('[data-path="income.totalIncome"]');
  if (totalIncEl) totalIncEl.value = fmt$(totalIncome);

  // Total taxes
  const fedTax = parseCommas(form.querySelector('[data-path="taxesExpenses.federalTax"]')?.value) || 0;
  const stTax = parseCommas(form.querySelector('[data-path="taxesExpenses.stateTax"]')?.value) || 0;
  const ficaTax = parseCommas(form.querySelector('[data-path="taxesExpenses.ficaTax"]')?.value) || 0;
  const totalTax = fedTax + stTax + ficaTax;
  const totalTaxEl = form.querySelector('[data-path="taxesExpenses.totalTaxesPaid"]');
  if (totalTaxEl) totalTaxEl.value = fmt$(totalTax);

  // Total expenses
  const expInputs = form.querySelectorAll('[data-path^="taxesExpenses.expenses["]');
  let totalExp = 0;
  expInputs.forEach(inp => {
    if (inp.dataset.path.endsWith('.amount')) totalExp += parseCommas(inp.value) || 0;
  });
  const totalExpEl = form.querySelector('[data-path="taxesExpenses.expenseTotal"]');
  if (totalExpEl) totalExpEl.value = fmt$(totalExp);

  renderPQDashboard();
}

/* ---------- PQ FLOATING DASHBOARD ---------- */
function renderPQDashboard() {
  const host = document.getElementById('pq-dashboard');
  if (!host) return;
  host.innerHTML = '';
  const form = document.getElementById('pq-form');
  if (!form) return;

  // ── Gather ALL data live from the DOM (parseCommas handles formatted currency) ──

  // Helper: sum all inputs matching a selector where the path ends with a given key
  const sumInputs = (selector, keySuffix) => {
    let total = 0;
    form.querySelectorAll(selector).forEach(inp => {
      if (!keySuffix || inp.dataset.path.endsWith(keySuffix)) {
        total += parseCommas(inp.value) || 0;
      }
    });
    return total;
  };

  // Total Assets = RE market values + all investment account market values
  let totalAssets = 0;
  totalAssets += sumInputs('[data-path^="assets.realEstate["]', '.marketValue');
  totalAssets += sumInputs('[data-path^="assets.taxDeferred["]', '.marketValue');
  totalAssets += sumInputs('[data-path^="assets.roth["]', '.marketValue');
  totalAssets += sumInputs('[data-path^="assets.taxable["]', '.marketValue');
  totalAssets += sumInputs('[data-path^="assets.cashCd["]', '.marketValue');
  totalAssets += sumInputs('[data-path^="assets.plan529["]', '.marketValue');
  totalAssets += sumInputs('[data-path^="assets.hsa["]', '.marketValue');

  // Total Savings = personal additions + company match across retirement accounts
  let totalSavings = 0;
  totalSavings += sumInputs('[data-path$=".personalAdditions"]');
  totalSavings += sumInputs('[data-path$=".companyMatch"]');

  // Total Liabilities
  const totalLiabilities = sumInputs('[data-path^="liabilities.items["]', '.amount');

  // Total Income (across all 4 sub-tables, excluding future-dated entries)
  const currentYear = new Date().getFullYear();
  let totalIncome = 0;

  // Helper: check if a date string is in a future year
  const isFutureDate = (dateStr) => {
    if (!dateStr) return false;
    const d = new Date(dateStr);
    return !isNaN(d.getTime()) && d.getFullYear() > currentYear;
  };

  // Employment — always included (no start date filter)
  totalIncome += sumInputs('[data-path^="income.employment["]', '.annualAmount');

  // Social Security — exclude if startingAge > current age of the owner
  const ssTbody = form.querySelector('tbody[data-table-path="income.socialSecurity"]');
  if (ssTbody) {
    const age1Val = parseFloat(form.querySelector('[data-path="family.client1Age"]')?.value) || 0;
    const age2Val = parseFloat(form.querySelector('[data-path="family.client2Age"]')?.value) || 0;
    const c1NameVal = [
      (form.querySelector('[data-path="family.client1FirstName"]')?.value || '').trim(),
      (form.querySelector('[data-path="family.client1LastName"]')?.value || '').trim()
    ].filter(Boolean).join(' ');
    const c2NameVal = [
      (form.querySelector('[data-path="family.client2FirstName"]')?.value || '').trim(),
      (form.querySelector('[data-path="family.client2LastName"]')?.value || '').trim()
    ].filter(Boolean).join(' ');
    ssTbody.querySelectorAll('tr').forEach(row => {
      const amt = parseCommas(row.querySelector('input[data-path$=".annualAmount"]')?.value) || 0;
      const startAge = parseFloat(row.querySelector('input[data-path$=".startingAge"]')?.value) || 0;
      const owner = (row.querySelector('input[data-path$=".owner"]')?.value || '').trim();
      // Determine current age of the owner
      let ownerAge = age1Val; // default to client 1
      if (c2NameVal && owner === c2NameVal) ownerAge = age2Val;
      if (startAge > 0 && startAge > ownerAge) return; // future — skip
      totalIncome += amt;
    });
  }

  // Pension — exclude if startDate is in a future year
  const penTbody = form.querySelector('tbody[data-table-path="income.pension"]');
  if (penTbody) {
    penTbody.querySelectorAll('tr').forEach(row => {
      const amt = parseCommas(row.querySelector('input[data-path$=".annualAmount"]')?.value) || 0;
      const startDate = row.querySelector('input[data-path$=".startDate"]')?.value || '';
      if (isFutureDate(startDate)) return; // future — skip
      totalIncome += amt;
    });
  }

  // Other — exclude if startDate is in a future year
  const otherTbody = form.querySelector('tbody[data-table-path="income.other"]');
  if (otherTbody) {
    otherTbody.querySelectorAll('tr').forEach(row => {
      const amt = parseCommas(row.querySelector('input[data-path$=".annualAmount"]')?.value) || 0;
      const startDate = row.querySelector('input[data-path$=".startDate"]')?.value || '';
      if (isFutureDate(startDate)) return; // future — skip
      totalIncome += amt;
    });
  }

  // Total Taxes
  let totalTax = 0;
  ['federalTax', 'stateTax', 'ficaTax'].forEach(k => {
    totalTax += parseCommas(form.querySelector(`[data-path="taxesExpenses.${k}"]`)?.value) || 0;
  });

  // Total Expenses — split by source
  const totalExpenses = sumInputs('[data-path^="taxesExpenses.expenses["]', '.amount');

  // Liability Payments: sum expense amounts where hidden _source = 'liability'
  let liabilityPayments = 0;
  let nonLiabilityExpenses = 0;
  const expTbody = form.querySelector('tbody[data-table-path="taxesExpenses.expenses"]');
  if (expTbody) {
    expTbody.querySelectorAll('tr').forEach(row => {
      const sourceInp = row.querySelector('input[data-path$="._source"]');
      const amountInp = row.querySelector('input[data-path$=".amount"]');
      const amt = parseCommas(amountInp?.value) || 0;
      if (sourceInp?.value === 'liability') {
        liabilityPayments += amt;
      } else {
        nonLiabilityExpenses += amt;
      }
    });
  }

  // Derived metrics
  const totalNetWorth = totalAssets - totalLiabilities;
  const dti = totalIncome > 0 ? (totalLiabilities / totalIncome * 100) : 0;
  const savingsRatio = totalIncome > 0 ? (totalSavings / totalIncome * 100) : 0;
  const cashFlow = totalIncome - totalSavings - liabilityPayments - totalTax - nonLiabilityExpenses;

  // ── Balance Sheet card ──
  const card1 = el('div', { className: 'dashboard-card' });
  card1.appendChild(el('h3', { textContent: 'Balance Sheet' }));
  const bsMetrics = el('ul', { className: 'metric-list' });
  [
    ['Total Assets', fmt$(totalAssets), ''],
    ['Total Liabilities', fmt$(totalLiabilities), ''],
    ['Total Net Worth', fmt$(totalNetWorth), totalNetWorth >= 0 ? 'positive' : 'negative'],
  ].forEach(([label, value, colorClass]) => {
    const li = el('li');
    li.appendChild(el('span', { textContent: label }));
    li.appendChild(el('strong', { className: `metric-value ${colorClass}`.trim(), textContent: value }));
    bsMetrics.appendChild(li);
  });
  card1.appendChild(bsMetrics);

  // ── Yearly Cash Flow card ──
  const card2 = el('div', { className: 'dashboard-card' });
  card2.appendChild(el('h3', { textContent: 'Yearly Cash Flow' }));
  const cfMetrics = el('ul', { className: 'metric-list' });

  const addCFRow = (label, value, colorClass) => {
    const li = el('li');
    li.appendChild(el('span', { textContent: label }));
    li.appendChild(el('strong', { className: `metric-value ${colorClass || ''}`.trim(), textContent: value }));
    cfMetrics.appendChild(li);
  };

  addCFRow('Income', fmt$(totalIncome));
  addCFRow('Savings', fmt$(totalSavings));
  addCFRow('Liability Payments', fmt$(liabilityPayments));
  addCFRow('Taxes', fmt$(totalTax));
  addCFRow('Non-Liability Expenses', fmt$(nonLiabilityExpenses));

  // Divider before Cash Flow
  const dividerLi = el('li', { style: { borderTop: '1px solid var(--border)', paddingTop: '6px', marginTop: '4px' } });
  dividerLi.appendChild(el('span', { textContent: 'Cash Flow', style: { fontWeight: '700' } }));
  dividerLi.appendChild(el('strong', { className: `metric-value ${cashFlow >= 0 ? 'positive' : 'negative'}`, textContent: fmt$(cashFlow) }));
  cfMetrics.appendChild(dividerLi);

  card2.appendChild(cfMetrics);

  // ── Key Ratios card ──
  const cardRatios = el('div', { className: 'dashboard-card' });
  cardRatios.appendChild(el('h3', { textContent: 'Key Ratios' }));
  const ratios = el('ul', { className: 'metric-list' });

  const addRatio = (label, value, colorClass) => {
    const li = el('li');
    li.appendChild(el('span', { textContent: label }));
    li.appendChild(el('strong', { className: `metric-value ${colorClass}`, textContent: value }));
    ratios.appendChild(li);
  };

  addRatio('Debt-to-Income', fmtPct(dti), dti > 40 ? 'negative' : dti > 20 ? 'warning' : 'positive');
  addRatio('Savings Ratio', fmtPct(savingsRatio), savingsRatio >= 20 ? 'positive' : savingsRatio >= 10 ? 'warning' : 'negative');

  cardRatios.appendChild(ratios);

  // Completeness Score
  const card3 = el('div', { className: 'dashboard-card' });
  card3.appendChild(el('h3', { textContent: 'Completeness' }));

  const allInputs = form.querySelectorAll('input[data-path], select[data-path], textarea[data-path]');
  let filled = 0, total = 0;
  allInputs.forEach(inp => {
    if (inp.type === 'hidden' || inp.readOnly) return;
    total++;
    if (inp.value && inp.value.trim() !== '' && inp.value !== '— Select —') filled++;
  });
  const pct = total > 0 ? Math.round(filled / total * 100) : 0;
  const colorClass = pct < 33 ? 'low' : pct < 66 ? 'mid' : 'high';
  const colorMap = { low: '#dc3545', mid: '#d9a708', high: '#198754' };
  const strokeColor = colorMap[colorClass];

  // SVG circular progress
  const ringSize = 90;
  const strokeWidth = 7;
  const radius = (ringSize - strokeWidth) / 2;
  const circumference = 2 * Math.PI * radius;
  const offset = circumference - (pct / 100) * circumference;

  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('class', 'completeness-ring');
  svg.setAttribute('width', ringSize);
  svg.setAttribute('height', ringSize);
  svg.setAttribute('viewBox', `0 0 ${ringSize} ${ringSize}`);

  const bgCircle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
  bgCircle.setAttribute('cx', ringSize / 2);
  bgCircle.setAttribute('cy', ringSize / 2);
  bgCircle.setAttribute('r', radius);
  bgCircle.setAttribute('fill', 'none');
  bgCircle.setAttribute('stroke', '#e5eaef');
  bgCircle.setAttribute('stroke-width', strokeWidth);

  const fgCircle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
  fgCircle.setAttribute('cx', ringSize / 2);
  fgCircle.setAttribute('cy', ringSize / 2);
  fgCircle.setAttribute('r', radius);
  fgCircle.setAttribute('fill', 'none');
  fgCircle.setAttribute('stroke', strokeColor);
  fgCircle.setAttribute('stroke-width', strokeWidth);
  fgCircle.setAttribute('stroke-linecap', 'round');
  fgCircle.setAttribute('stroke-dasharray', circumference);
  fgCircle.setAttribute('stroke-dashoffset', offset);
  fgCircle.setAttribute('transform', `rotate(-90 ${ringSize / 2} ${ringSize / 2})`);
  fgCircle.style.transition = 'stroke-dashoffset 400ms ease';

  const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  text.setAttribute('x', '50%');
  text.setAttribute('y', '50%');
  text.setAttribute('dominant-baseline', 'central');
  text.setAttribute('text-anchor', 'middle');
  text.setAttribute('class', 'ring-pct-text');
  text.textContent = pct + '%';

  svg.append(bgCircle, fgCircle, text);

  const ringWrapper = el('div', { className: 'completeness-score' });
  ringWrapper.appendChild(svg);
  ringWrapper.appendChild(el('div', { className: 'score-label', textContent: 'Fields completed' }));
  card3.appendChild(ringWrapper);

  host.append(card1, card2, cardRatios);

  // Left panel — Completeness + Advisor Notes
  const notesHost = document.getElementById('pq-notes-panel');
  if (notesHost) {
    notesHost.innerHTML = '';

    // Completeness card
    notesHost.appendChild(card3);

    // Advisor Notes card
    const notesCard = el('div', { className: 'dashboard-card notes-panel' });
    notesCard.appendChild(el('h3', { textContent: 'Advisor Notes' }));
    const notesTA = el('textarea', { placeholder: 'Free-form notes...' });
    notesTA.dataset.path = '_advisorNotes';
    const store = getStore();
    if (store._advisorNotes) notesTA.value = store._advisorNotes;
    notesTA.addEventListener('input', () => {
      const s = getStore();
      s._advisorNotes = notesTA.value;
      saveStore(s);
    });
    notesCard.appendChild(notesTA);
    notesHost.appendChild(notesCard);
  }
}

/* ==========================================================================
   PQ REVIEW SCREEN
   ========================================================================== */

function renderPQReview() {
  const host = document.getElementById('pq-review-content');
  host.innerHTML = '';
  const store = getStore();
  const pq = store.pqTemplate || {};

  const categories = {
    'Ready for Planning': [],
    'Missing for Planning': [],
    'Missing for Account Opening': [],
    'Missing for Schwab Application': [],
  };

  // Classify fields
  const planningRequired = ['family.client1FirstName', 'family.client1LastName', 'family.client1DOB', 'goals.goals'];
  const accountOpeningRequired = ['contact.address1', 'contact.city', 'contact.state', 'contact.zip', 'contact.client1Cell', 'contact.client1Email'];
  const schwabRequired = ['family.client1FirstName', 'family.client1LastName', 'family.client1DOB', 'contact.address1', 'contact.city', 'contact.state', 'contact.zip'];

  function checkField(path, label) {
    const val = getDeep(pq, path);
    const has = val != null && String(val).trim() !== '';
    if (has) {
      categories['Ready for Planning'].push({ label, value: String(val).slice(0, 60) });
    } else {
      if (planningRequired.includes(path)) categories['Missing for Planning'].push({ label });
      if (accountOpeningRequired.includes(path)) categories['Missing for Account Opening'].push({ label });
      if (schwabRequired.includes(path)) categories['Missing for Schwab Application'].push({ label });
    }
  }

  // Walk through key fields
  checkField('family.client1FirstName', 'Client First Name');
  checkField('family.client1LastName', 'Client Last Name');
  checkField('family.client1DOB', 'Client DOB');
  checkField('family.client2FirstName', 'Spouse First Name');
  checkField('family.client2LastName', 'Spouse Last Name');
  checkField('contact.address1', 'Address');
  checkField('contact.city', 'City');
  checkField('contact.state', 'State');
  checkField('contact.zip', 'Zip');
  checkField('contact.client1Cell', 'Client Cell');
  checkField('contact.client1Email', 'Client Email');
  checkField('employment.client1.status', 'Client Employment Status');
  checkField('goals.goals', 'Goals');
  checkField('goals.estatePlan', 'Estate Plan');

  // Count asset / income / expense tables
  const tableChecks = [
    ['assets.realEstate', 'Real Estate Assets'],
    ['assets.taxDeferred', 'Tax-Deferred Accounts'],
    ['assets.roth', 'Roth Accounts'],
    ['assets.taxable', 'Taxable Accounts'],
    ['liabilities.items', 'Liabilities'],
    ['insurance.policies', 'Insurance'],
    ['income.employment', 'Employment Income'],
    ['income.socialSecurity', 'Social Security'],
    ['income.pension', 'Pension'],
    ['income.other', 'Other Income'],
    ['taxesExpenses.expenses', 'Expenses'],
  ];
  tableChecks.forEach(([path, label]) => {
    const arr = getDeep(pq, path);
    if (Array.isArray(arr) && arr.length && arr.some(hasData)) {
      categories['Ready for Planning'].push({ label, value: `${arr.length} row(s)` });
    } else {
      categories['Missing for Planning'].push({ label });
    }
  });

  const grid = el('div', { className: 'review-grid' });
  Object.entries(categories).forEach(([title, items]) => {
    const pane = el('div', { className: 'review-pane' });
    const badgeClass = title.includes('Ready') ? 'badge-green' : title.includes('Schwab') ? 'badge-red' : 'badge-yellow';
    const hdr = el('div', { className: 'pane-header' });
    hdr.appendChild(el('span', { textContent: title }));
    hdr.appendChild(el('span', { className: `review-count-badge ${badgeClass}`, textContent: String(items.length) }));
    pane.appendChild(hdr);

    const body = el('div', { className: 'pane-body' });
    if (!items.length) {
      body.appendChild(el('p', { className: 'section-note', textContent: 'None' }));
    } else {
      const list = el('ul', { className: 'review-list' });
      items.forEach(item => {
        const li = el('li');
        li.appendChild(el('span', { className: 'item-label', textContent: item.label }));
        if (item.value) li.appendChild(el('span', { className: 'item-value', textContent: item.value }));
        list.appendChild(li);
      });
      body.appendChild(list);
    }
    pane.appendChild(body);
    grid.appendChild(pane);
  });

  host.appendChild(grid);
}

/* ==========================================================================
   STAGE 2: PLAN DESIGN
   ========================================================================== */

const planPrefillBaseline = {};

function renderPlan() {
  const host = document.getElementById('plan-form-sections');
  host.innerHTML = '';
  const store = getStore();
  const pq = store.pqTemplate || {};

  planSections.forEach((section, si) => {
    const wrapper = el('section', { className: 'form-section' });
    const header = el('div', { className: 'section-header' });
    header.appendChild(el('h2', { textContent: section.title }));
    const toggleBtn = el('button', { type: 'button', className: 'toggle-btn', textContent: si < 3 ? '\u2212' : '+' });
    header.appendChild(toggleBtn);
    wrapper.appendChild(header);

    const content = el('div', { className: 'section-content' + (si < 3 ? '' : ' collapsed') });
    if (section.note) content.appendChild(el('p', { className: 'section-note', textContent: section.note }));

    if (section.fields?.length) {
      const grid = el('div', { className: 'grid' });
      section.fields.forEach(f => grid.appendChild(buildFieldEl(f, section.id, 'plan')));
      content.appendChild(grid);
    }
    if (section.tables) {
      section.tables.forEach(t => {
        const existing = getDeep(store.planOrder, `${section.id}.${t.key}`);
        content.appendChild(buildTableEl(section.id, t, existing));
      });
    }

    header.addEventListener('click', () => {
      const c = content.classList.contains('collapsed');
      content.classList.toggle('collapsed', !c);
      toggleBtn.textContent = c ? '\u2212' : '+';
    });

    wrapper.appendChild(content);
    host.appendChild(wrapper);
  });

  // Prefill from PQ
  prefillPlanFromPQ();
}

function prefillPlanFromPQ() {
  const store = getStore();
  const pq = store.pqTemplate || {};
  const plan = store.planOrder || {};
  const form = document.getElementById('plan-form');
  Object.keys(planPrefillBaseline).forEach(k => delete planPrefillBaseline[k]);

  // Fill already-saved plan data first
  fillForm(form, plan);

  // Now attempt prefill for fields that have prefillFrom
  planSections.forEach(section => {
    (section.fields || []).forEach(field => {
      if (!field.prefillFrom) return;
      const path = `${section.id}.${field.key}`;
      const inp = form.querySelector(`[data-path="${path}"]`);
      if (!inp) return;

      // If plan already has a saved value, mark green
      const saved = getDeep(plan, path);
      if (saved != null && String(saved).trim() !== '') {
        inp.value = saved;
        setPlanFlag(path, 'flag-green');
        return;
      }

      // Try to compute prefill
      let val;
      if (field.prefillFrom === '_computed_client1Age') {
        val = calcAge(pq.family?.client1DOB);
      } else if (field.prefillFrom === '_computed_client2Age') {
        val = calcAge(pq.family?.client2DOB);
      } else if (field.prefillFrom.includes('+')) {
        // Concatenate fields
        val = field.prefillFrom.split('+').map(p => getDeep(pq, p)).filter(Boolean).join(' ');
      } else {
        val = getDeep(pq, field.prefillFrom);
      }

      if (val != null && String(val).trim() !== '') {
        inp.value = val;
        planPrefillBaseline[path] = String(val);
        setPlanFlag(path, 'flag-yellow');
      } else {
        setPlanFlag(path, 'flag-red');
      }
    });
  });

  // Color remaining fields
  form.querySelectorAll('[data-path]').forEach(inp => {
    const path = inp.dataset.path;
    const wrapper = inp.closest('.field, .field-medium, .field-wide, .field-full, .field-2');
    if (!wrapper) return;
    if (wrapper.classList.contains('flag-yellow') || wrapper.classList.contains('flag-green')) return;
    if (inp.value && inp.value.trim() !== '') {
      setPlanFlag(path, 'flag-green');
    } else {
      setPlanFlag(path, 'flag-red');
    }
  });
}

function setPlanFlag(path, flag) {
  const inp = document.querySelector(`#plan-form [data-path="${path}"]`);
  if (!inp) return;
  const w = inp.closest('.field, .field-medium, .field-wide, .field-full, .field-2');
  if (!w) return;
  w.classList.remove('flag-red', 'flag-yellow', 'flag-green');
  w.classList.add(flag);
}

function onPlanInput(e) {
  const inp = e.target;
  if (!inp?.dataset?.path) return;
  const path = inp.dataset.path;
  const val = String(inp.value || '');
  const baseline = planPrefillBaseline[path];
  if (val === '') { setPlanFlag(path, 'flag-red'); return; }
  if (baseline && val === baseline) { setPlanFlag(path, 'flag-yellow'); return; }
  setPlanFlag(path, 'flag-green');
}

/* ==========================================================================
   STAGE 3: INVESTMENT ACTION PLAN
   ========================================================================== */

function renderIAP() {
  const host = document.getElementById('iap-content');
  host.innerHTML = '';
  const store = getStore();
  const iap = store.investmentActionPlan || {};

  // Summary cards
  const cards = el('div', { className: 'iap-summary-cards' });
  const cardDefs = [
    { label: 'Target Allocation', key: 'targetAllocation', default: 'Not Set' },
    { label: 'Est. Total Assets', key: 'estimatedTotalAssets', format: true },
    { label: 'Est. Total Transfers', key: 'estimatedTotalTransfers', format: true },
    { label: 'Beneficiary Instruction', key: 'beneficiaryInstruction', default: 'TBD' },
  ];
  cardDefs.forEach(def => {
    const card = el('div', { className: 'iap-card' });
    card.appendChild(el('div', { className: 'card-label', textContent: def.label }));
    let val = iap[def.key] || def.default || '$0';
    if (def.format && typeof val === 'number') val = fmt$(val);
    card.appendChild(el('div', { className: 'card-value', textContent: val || '$0' }));
    cards.appendChild(card);
  });
  host.appendChild(cards);

  // Editable summary fields
  const summarySection = el('section', { className: 'form-section' });
  const summaryHeader = el('div', { className: 'section-header static-header' });
  summaryHeader.appendChild(el('h2', { textContent: 'Investment Action Plan Details' }));
  summarySection.appendChild(summaryHeader);
  const summaryContent = el('div', { className: 'section-content' });

  const summaryGrid = el('div', { className: 'grid' });
  [
    { key: 'targetAllocation', label: 'Target Asset Allocation', width: 'medium' },
    { key: 'estimatedTotalAssets', label: 'Estimated Total Assets', type: 'number', width: 'medium' },
    { key: 'estimatedTotalTransfers', label: 'Estimated Total Transfers', type: 'number', width: 'medium' },
    { key: 'beneficiaryInstruction', label: 'Beneficiary Designation Instruction', width: 'wide' },
    { key: 'capitalGainsNotes', label: 'Capital Gains / Tax Notes', type: 'textarea' },
  ].forEach(f => summaryGrid.appendChild(buildFieldEl(f, 'iap', 'iap')));
  summaryContent.appendChild(summaryGrid);
  summarySection.appendChild(summaryContent);
  host.appendChild(summarySection);

  // Action rows by category
  const categories = ['Qualified Accounts', 'Roth Accounts', 'Non-Qualified Accounts', 'Other / Insurance / Cash'];
  const catKeys = ['qualified', 'roth', 'nonQualified', 'other'];

  categories.forEach((catLabel, ci) => {
    const section = el('section', { className: 'form-section' });
    const hdr = el('div', { className: 'section-header' });
    hdr.appendChild(el('h2', { textContent: catLabel }));
    const toggleBtn = el('button', { type: 'button', className: 'toggle-btn', textContent: '\u2212' });
    hdr.appendChild(toggleBtn);
    section.appendChild(hdr);

    const content = el('div', { className: 'section-content' });
    const existingRows = getDeep(iap, `actionRows.${catKeys[ci]}`);
    content.appendChild(buildTableEl(`iap.actionRows`, {
      key: catKeys[ci],
      columns: [
        { key: 'owner', label: 'Owner' },
        { key: 'sourceAccount', label: 'Source Account' },
        { key: 'totalAmount', label: 'Total Amount', type: 'number' },
        { key: 'feeOrSurrender', label: 'Fee/Surrender', type: 'number' },
        { key: 'action', label: 'Action' },
        { key: 'transferAmount', label: 'Transfer Amt', type: 'number' },
        { key: 'receivingAccount', label: 'Receiving Account' },
        { key: 'newAccount', label: 'New Acct?', type: 'select', options: ['No', 'Yes'] },
        { key: 'needsStatement', label: 'Stmt Needed?', type: 'select', options: ['No', 'Yes'] },
        { key: 'status', label: 'Status', type: 'select', options: ['Pending', 'Validated', 'Rejected', 'TBD'] },
        { key: 'notes', label: 'Notes' },
      ]
    }, existingRows));

    hdr.addEventListener('click', () => {
      const c = content.classList.contains('collapsed');
      content.classList.toggle('collapsed', !c);
      toggleBtn.textContent = c ? '\u2212' : '+';
    });

    section.appendChild(content);
    host.appendChild(section);
  });

  // Fill saved data
  fillForm(host, iap);

  // Validation warnings
  renderIAPWarnings(host);
}

function renderIAPWarnings(host) {
  const warningSection = el('section', { className: 'form-section' });
  const hdr = el('div', { className: 'section-header static-header' });
  hdr.appendChild(el('h2', { textContent: 'Validation Warnings' }));
  warningSection.appendChild(hdr);

  const content = el('div', { className: 'section-content' });
  const warnings = collectIAPWarnings();

  if (!warnings.length) {
    content.appendChild(el('p', { className: 'section-note', textContent: 'No warnings — all action rows look valid.' }));
  } else {
    const list = el('ul', { className: 'blocker-list' });
    warnings.forEach(w => {
      const li = el('li');
      li.appendChild(el('span', { className: 'blocker-icon', textContent: w.level === 'error' ? '\u26d4' : '\u26a0' }));
      li.appendChild(el('span', { textContent: w.message }));
      list.appendChild(li);
    });
    content.appendChild(list);
  }

  warningSection.appendChild(content);
  host.appendChild(warningSection);
}

function collectIAPWarnings() {
  const warnings = [];
  const form = document.getElementById('iap-content');
  if (!form) return warnings;

  // Check each action row for issues
  const rows = form.querySelectorAll('tbody tr');
  rows.forEach((row, i) => {
    const inputs = row.querySelectorAll('[data-path]');
    const vals = {};
    inputs.forEach(inp => {
      const key = inp.dataset.path.split('.').pop();
      vals[key] = inp.value;
    });

    if (vals.action && (vals.action.toLowerCase().includes('transfer') || vals.action.toLowerCase().includes('rollover'))) {
      if (!vals.receivingAccount || vals.receivingAccount.trim() === '') {
        warnings.push({ level: 'error', message: `Row ${i + 1}: Transfer/rollover action but no receiving account specified.` });
      }
    }
    if (vals.newAccount === 'Yes') {
      warnings.push({ level: 'warning', message: `Row ${i + 1}: New account must be opened — ensure account-opening instructions exist.` });
    }
    if (vals.needsStatement === 'Yes') {
      warnings.push({ level: 'warning', message: `Row ${i + 1}: Statement needed — mark as resolved before final submission.` });
    }
    if (vals.status === 'TBD') {
      warnings.push({ level: 'warning', message: `Row ${i + 1}: Status is TBD — needs resolution.` });
    }
  });

  return warnings;
}

/* ==========================================================================
   STAGE 4: CLIENT INFO SHEET / SCHWAB READINESS
   ========================================================================== */

function renderClientSheet() {
  const host = document.getElementById('client-content');
  host.innerHTML = '';
  const store = getStore();
  const pq = store.pqTemplate || {};

  // Build structured review sections
  const sections = [
    {
      title: 'Client 1 Legal Identity',
      fields: [
        { key: 'legalFirstName', label: 'Legal First Name', source: 'family.client1FirstName', required: true },
        { key: 'middleInitial', label: 'Middle Initial' },
        { key: 'legalLastName', label: 'Legal Last Name', source: 'family.client1LastName', required: true },
        { key: 'dob', label: 'Date of Birth', type: 'date', source: 'family.client1DOB', required: true },
        { key: 'ssn', label: 'Social Security Number', masked: true, required: true },
        { key: 'citizenship', label: 'Citizenship', required: true },
        { key: 'dlState', label: "Driver's License State", required: true },
        { key: 'dlNumber', label: "Driver's License Number", masked: true, required: true },
        { key: 'dlExpiration', label: 'DL Expiration', type: 'date', required: true },
        { key: 'cellPhone', label: 'Cell Phone', source: 'contact.client1Cell', required: true },
        { key: 'email', label: 'Email', type: 'email', source: 'contact.client1Email', required: true },
      ]
    },
    {
      title: 'Client 2 Legal Identity',
      showIf: () => pqState.hasSpouse || pq.family?.client2FirstName,
      fields: [
        { key: 'c2LegalFirstName', label: 'Legal First Name', source: 'family.client2FirstName', required: true },
        { key: 'c2MiddleInitial', label: 'Middle Initial' },
        { key: 'c2LegalLastName', label: 'Legal Last Name', source: 'family.client2LastName', required: true },
        { key: 'c2Dob', label: 'Date of Birth', type: 'date', source: 'family.client2DOB', required: true },
        { key: 'c2Ssn', label: 'Social Security Number', masked: true, required: true },
        { key: 'c2Citizenship', label: 'Citizenship', required: true },
        { key: 'c2DlState', label: "Driver's License State", required: true },
        { key: 'c2DlNumber', label: "Driver's License Number", masked: true, required: true },
        { key: 'c2DlExpiration', label: 'DL Expiration', type: 'date', required: true },
        { key: 'c2CellPhone', label: 'Cell Phone', source: 'contact.client2Cell', required: true },
        { key: 'c2Email', label: 'Email', type: 'email', source: 'contact.client2Email', required: true },
      ]
    },
    {
      title: 'Contact & Residence',
      fields: [
        { key: 'address1', label: 'Address', source: 'contact.address1', required: true, width: 'wide' },
        { key: 'address2', label: 'Address Line 2', source: 'contact.address2', width: 'wide' },
        { key: 'city', label: 'City', source: 'contact.city', required: true, width: 'medium' },
        { key: 'state', label: 'State', source: 'contact.state', required: true },
        { key: 'zip', label: 'Zip', source: 'contact.zip', required: true },
        { key: 'homePhone', label: 'Home Phone', source: 'contact.homePhone' },
        { key: 'fax', label: 'Fax' },
      ]
    },
    {
      title: 'Employment — Client 1',
      fields: [
        { key: 'c1Employer', label: 'Employer', source: 'employment.client1.employer', required: true },
        { key: 'c1EmployerAddress', label: 'Employer Address', width: 'wide' },
        { key: 'c1EmployerCSZ', label: 'City / State / Zip', width: 'medium' },
        { key: 'c1Title', label: 'Title / Position', source: 'employment.client1.jobTitle' },
        { key: 'c1BusinessType', label: 'Type of Business', source: 'employment.client1.businessType' },
        { key: 'c1WorkPhone', label: 'Work Phone' },
        { key: 'c1Retired', label: 'Retired?', type: 'select', options: ['No', 'Yes'] },
      ]
    },
    {
      title: 'Employment — Client 2',
      showIf: () => pqState.hasSpouse || pq.family?.client2FirstName,
      fields: [
        { key: 'c2Employer', label: 'Employer', source: 'employment.client2.employer' },
        { key: 'c2EmployerAddress', label: 'Employer Address', width: 'wide' },
        { key: 'c2EmployerCSZ', label: 'City / State / Zip', width: 'medium' },
        { key: 'c2Title', label: 'Title / Position', source: 'employment.client2.jobTitle' },
        { key: 'c2BusinessType', label: 'Type of Business', source: 'employment.client2.businessType' },
        { key: 'c2WorkPhone', label: 'Work Phone' },
        { key: 'c2Retired', label: 'Retired?', type: 'select', options: ['No', 'Yes'] },
      ]
    },
    {
      title: 'Children / Beneficiaries',
      table: {
        key: 'beneficiaries',
        columns: [
          { key: 'firstName', label: 'First Name' },
          { key: 'lastName', label: 'Last Name' },
          { key: 'relationship', label: 'Relationship' },
          { key: 'dob', label: 'Date of Birth', type: 'date' },
          { key: 'ssn', label: 'SSN' },
          { key: 'childOf', label: 'Child of (Client 1/2)' },
        ]
      }
    },
  ];

  sections.forEach(section => {
    if (section.showIf && !section.showIf()) return;

    const wrapper = el('section', { className: 'form-section' });
    const hdr = el('div', { className: 'section-header' });
    hdr.appendChild(el('h2', { textContent: section.title }));

    // Completion meter
    let fieldCount = 0, filledCount = 0;
    if (section.fields) {
      section.fields.forEach(f => {
        if (f.required) {
          fieldCount++;
          const clientVal = getDeep(store.clientSheet, `${section.title}.${f.key}`);
          const sourceVal = f.source ? getDeep(pq, f.source) : null;
          if ((clientVal && String(clientVal).trim()) || (sourceVal && String(sourceVal).trim())) filledCount++;
        }
      });
    }
    if (fieldCount > 0) {
      const meter = el('div', { className: 'completion-meter' });
      const pct = Math.round(filledCount / fieldCount * 100);
      const bar = el('div', { className: 'completion-bar' });
      bar.appendChild(el('div', { className: 'completion-bar-fill', style: { width: pct + '%', background: pct === 100 ? '#198754' : pct > 50 ? '#d9a708' : '#dc3545' } }));
      meter.append(bar, document.createTextNode(` ${pct}%`));
      hdr.appendChild(meter);
    }

    wrapper.appendChild(hdr);
    const content = el('div', { className: 'section-content' });

    if (section.fields) {
      const grid = el('div', { className: 'grid' });
      section.fields.forEach(f => {
        const fieldEl = buildFieldEl(f, 'clientSheet.' + section.title, 'client');
        const inp = fieldEl.querySelector('[data-path]');

        // Prefill from PQ
        if (f.source) {
          const srcVal = getDeep(pq, f.source);
          if (srcVal) {
            inp.value = srcVal;
            fieldEl.classList.add('flag-yellow');
            // Add source tag
            const labelRow = fieldEl.querySelector('.label-row');
            if (labelRow) labelRow.appendChild(el('span', { className: 'field-tag tag-source', textContent: 'From PQ' }));
          }
        }

        // Override with client sheet saved data
        const saved = getDeep(store.clientSheet, `${section.title}.${f.key}`);
        if (saved != null && String(saved).trim() !== '') {
          inp.value = saved;
          fieldEl.classList.remove('flag-yellow');
          fieldEl.classList.add('flag-green');
        }

        // Mark required but empty as red
        if (f.required && (!inp.value || inp.value.trim() === '')) {
          fieldEl.classList.add('flag-red');
        }

        // Select field setup
        if (f.type === 'select' && f.options) {
          // Already handled in buildFieldEl
        }

        grid.appendChild(fieldEl);
      });
      content.appendChild(grid);
    }

    if (section.table) {
      const existingBene = getDeep(store.clientSheet, 'beneficiaries') || [
        ...(pq.family?.children || []).map(c => ({ firstName: c.name, relationship: 'Child' })),
        ...(pq.family?.grandchildren || []).map(c => ({ firstName: c.name, relationship: 'Grandchild' })),
      ];
      content.appendChild(buildTableEl('clientSheet', section.table, existingBene));
    }

    wrapper.appendChild(content);
    host.appendChild(wrapper);
  });

  // Schwab Application Packet Readiness
  renderSchwabReadiness(host, store, pq);

  // Certification box
  const certBox = el('div', { className: 'certification-box' });
  const certLabel = el('label');
  const certCB = el('input', { type: 'checkbox', id: 'schwab-cert-checkbox' });
  certCB.addEventListener('change', () => {
    document.getElementById('client-submit-btn').disabled = !certCB.checked;
  });
  certLabel.append(certCB, document.createTextNode('I have reviewed all client information and investment instructions for Schwab readiness. All required fields are complete and accurate.'));
  certBox.appendChild(certLabel);
  host.appendChild(certBox);
}

function renderSchwabReadiness(host, store, pq) {
  const section = el('section', { className: 'form-section' });
  const hdr = el('div', { className: 'section-header static-header' });
  hdr.appendChild(el('h2', { textContent: 'Schwab Application Packet Readiness' }));
  section.appendChild(hdr);

  const content = el('div', { className: 'section-content' });
  const grid = el('div', { className: 'readiness-grid' });

  const checks = [
    { label: 'Identity Complete', check: () => !!(pq.family?.client1FirstName && pq.family?.client1LastName && pq.family?.client1DOB) },
    { label: 'Contact Complete', check: () => !!(pq.contact?.address1 && pq.contact?.city && pq.contact?.state && pq.contact?.zip) },
    { label: 'Employment Complete', check: () => !!(pq.employment?.client1?.status) },
    { label: 'Beneficiaries on File', check: () => !!(pq.family?.children?.length || getDeep(store, 'clientSheet.beneficiaries')?.length) },
    { label: 'Account Opening Instructions', check: () => !!(getDeep(store, 'planOrder.accountsToEstablish.accounts')?.length) },
    { label: 'Transfer Instructions', check: () => !!(getDeep(store, 'investmentActionPlan.actionRows')) },
    { label: 'Statements Collected', check: () => { const warnings = collectIAPWarnings(); return !warnings.some(w => w.message.includes('Statement')); } },
    { label: 'No Unresolved TBDs', check: () => { const warnings = collectIAPWarnings(); return !warnings.some(w => w.message.includes('TBD')); } },
  ];

  const blockers = [];
  checks.forEach(c => {
    const pass = c.check();
    const item = el('div', { className: 'readiness-item' });
    item.appendChild(el('div', { className: `readiness-icon ${pass ? 'pass' : 'fail'}`, textContent: pass ? '\u2713' : '\u2717' }));
    const info = el('div');
    info.appendChild(el('div', { className: 'readiness-label', textContent: c.label }));
    info.appendChild(el('div', { className: 'readiness-detail', textContent: pass ? 'Complete' : 'Incomplete — required' }));
    item.appendChild(info);
    grid.appendChild(item);
    if (!pass) blockers.push(c.label);
  });

  content.appendChild(grid);

  if (blockers.length) {
    content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Blockers', style: { marginTop: '14px', color: '#dc3545' } }));
    const list = el('ul', { className: 'blocker-list' });
    blockers.forEach(b => {
      const li = el('li');
      li.appendChild(el('span', { className: 'blocker-icon', textContent: '\u26d4' }));
      li.appendChild(el('span', { textContent: `${b} is incomplete. This must be resolved before Schwab submission.` }));
      list.appendChild(li);
    });
    content.appendChild(list);
  }

  section.appendChild(content);
  host.appendChild(section);
}

/* ==========================================================================
   PIPELINE / NAVIGATION
   ========================================================================== */

function setActiveScreen(id) {
  currentScreen = id;
  STAGES.forEach(s => {
    const node = document.getElementById(`screen-${s}`);
    if (!node) return;
    const show = s === id;
    node.classList.toggle('is-active', show);
    node.setAttribute('aria-hidden', String(!show));
  });
  updatePipeline();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function setPipelineStage(id, stateClass, statusText, metaText, disabled) {
  const stage = document.getElementById(`step-btn-${id}`);
  const status = document.getElementById(`stage-status-${id}`);
  const meta = document.getElementById(`stage-meta-${id}`);
  if (!stage) return;
  stage.classList.remove('is-current', 'is-complete', 'is-ready', 'is-locked');
  stage.classList.add(stateClass);
  stage.disabled = Boolean(disabled);
  if (status) status.textContent = statusText;
  if (meta) meta.textContent = metaText;
}

function fmtTs(ts) {
  if (!ts) return 'Not started';
  const d = new Date(ts);
  if (isNaN(d.getTime())) return 'Not started';
  return d.toLocaleString([], { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' });
}

function updatePipeline() {
  const store = getStore();
  const w = store.workflow || {};

  const pqCurrent = currentScreen === 'pq' || currentScreen === 'pq-review';
  setPipelineStage('pq',
    pqCurrent ? 'is-current' : (w.pqSubmittedAt ? 'is-complete' : 'is-ready'),
    w.pqSubmittedAt ? 'Complete' : 'Open',
    w.pqSubmittedAt ? `Submitted ${fmtTs(w.pqSubmittedAt)}` : 'Not submitted',
    false
  );

  setPipelineStage('plan',
    !w.pqSubmittedAt ? 'is-locked' : (currentScreen === 'plan' ? 'is-current' : (w.planSubmittedAt ? 'is-complete' : 'is-ready')),
    !w.pqSubmittedAt ? 'Waiting on PQ' : (w.planSubmittedAt ? 'Complete' : 'Open'),
    w.planSubmittedAt ? `Submitted ${fmtTs(w.planSubmittedAt)}` : 'Not started',
    !w.pqSubmittedAt
  );

  setPipelineStage('iap',
    !w.planSubmittedAt ? 'is-locked' : (currentScreen === 'iap' ? 'is-current' : (w.iapValidatedAt ? 'is-complete' : 'is-ready')),
    !w.planSubmittedAt ? 'Waiting on Plan' : (w.iapValidatedAt ? 'Validated' : 'Open'),
    w.iapValidatedAt ? `Validated ${fmtTs(w.iapValidatedAt)}` : 'Not started',
    !w.planSubmittedAt
  );

  setPipelineStage('client',
    !w.iapValidatedAt ? 'is-locked' : (currentScreen === 'client' ? 'is-current' : (w.clientCertifiedAt ? 'is-complete' : 'is-ready')),
    !w.iapValidatedAt ? 'Waiting on IAP' : (w.clientCertifiedAt ? 'Certified' : 'Open'),
    w.clientCertifiedAt ? `Certified ${fmtTs(w.clientCertifiedAt)}` : 'Not started',
    !w.iapValidatedAt
  );
}

/* ==========================================================================
   EVENT BINDING
   ========================================================================== */

function bindEvents() {
  // PQ form
  const pqForm = document.getElementById('pq-form');
  pqForm.addEventListener('input', e => {
    if (e.target?.dataset?.path === '_advisorNotes') return;
    recalcPQComputed();
  });
  pqForm.addEventListener('submit', e => {
    e.preventDefault();
    const data = collectForm(pqForm);
    // Save checkbox states
    data.family = data.family || {};
    data.family.hasSpouse = pqState.hasSpouse;
    data.family.hasChildren = pqState.hasChildren;
    // Save asset checkbox states
    data.assets = data.assets || {};
    Object.keys(pqState.assetChecks).forEach(k => {
      data.assets[k] = pqState.assetChecks[k];
    });
    // Save employment status from radio buttons
    data.employment = data.employment || {};
    data.employment.client1 = data.employment.client1 || {};
    data.employment.client1.status = pqState.employmentClient1;
    if (pqState.hasSpouse) {
      data.employment.client2 = data.employment.client2 || {};
      data.employment.client2.status = pqState.employmentClient2;
    }

    const store = getStore();
    store.pqTemplate = data;
    store.workflow = store.workflow || {};
    store.workflow.pqSubmittedAt = new Date().toISOString();
    saveStore(store);
    _pqDraft = {}; // Clear draft after successful submit
    renderPQReview();
    setActiveScreen('pq-review');
  });

  document.getElementById('pq-reset-btn').addEventListener('click', () => {
    if (!confirm('Clear all PQ form data?')) return;
    pqForm.querySelectorAll('input, select, textarea').forEach(inp => {
      if (inp.tagName === 'SELECT') inp.selectedIndex = 0;
      else if (inp.type === 'checkbox' || inp.type === 'radio') inp.checked = false;
      else inp.value = '';
    });
    pqState.hasSpouse = false;
    pqState.hasChildren = false;
    Object.keys(pqState.assetChecks).forEach(k => pqState.assetChecks[k] = false);
    pqState.employmentClient1 = 'Employed';
    pqState.employmentClient2 = 'Employed';
    _pqDraft = {}; // Clear draft on reset
    renderPQ();
  });

  // PQ Review
  document.getElementById('back-to-pq-edit-btn').addEventListener('click', () => setActiveScreen('pq'));
  document.getElementById('continue-to-plan-btn').addEventListener('click', () => {
    const store = getStore();
    if (!store.workflow?.pqSubmittedAt) { alert('Submit PQ first.'); return; }
    renderPlan();
    setActiveScreen('plan');
  });

  // Plan form
  const planForm = document.getElementById('plan-form');
  planForm.addEventListener('input', onPlanInput);
  planForm.addEventListener('submit', e => {
    e.preventDefault();
    const store = getStore();
    store.planOrder = collectForm(planForm);
    store.workflow = store.workflow || {};
    store.workflow.planSubmittedAt = new Date().toISOString();
    saveStore(store);
    renderIAP();
    setActiveScreen('iap');
  });

  document.getElementById('back-to-pq-review-btn').addEventListener('click', () => {
    renderPQReview();
    setActiveScreen('pq-review');
  });

  // IAP
  document.getElementById('back-to-plan-btn2').addEventListener('click', () => {
    renderPlan();
    setActiveScreen('plan');
  });
  document.getElementById('iap-submit-btn').addEventListener('click', () => {
    const store = getStore();
    const iapHost = document.getElementById('iap-content');
    store.investmentActionPlan = collectForm(iapHost);
    store.workflow = store.workflow || {};
    store.workflow.iapValidatedAt = new Date().toISOString();
    saveStore(store);
    renderClientSheet();
    setActiveScreen('client');
  });

  // Client Sheet
  document.getElementById('back-to-iap-btn').addEventListener('click', () => {
    renderIAP();
    setActiveScreen('iap');
  });
  document.getElementById('client-submit-btn').addEventListener('click', () => {
    const store = getStore();
    const clientHost = document.getElementById('client-content');
    store.clientSheet = collectForm(clientHost);
    store.workflow = store.workflow || {};
    store.workflow.clientCertifiedAt = new Date().toISOString();
    saveStore(store);
    alert('Schwab readiness certified! Client info sheet is complete.');
    updatePipeline();
  });

  // Pipeline navigation
  document.getElementById('step-btn-pq').addEventListener('click', () => {
    const store = getStore();
    if (store.workflow?.pqSubmittedAt) { renderPQReview(); setActiveScreen('pq-review'); }
    else setActiveScreen('pq');
  });
  document.getElementById('step-btn-plan').addEventListener('click', () => {
    const store = getStore();
    if (!store.workflow?.pqSubmittedAt) { alert('Submit PQ to unlock Plan Design.'); return; }
    renderPlan();
    setActiveScreen('plan');
  });
  document.getElementById('step-btn-iap').addEventListener('click', () => {
    const store = getStore();
    if (!store.workflow?.planSubmittedAt) { alert('Submit Plan to unlock IAP.'); return; }
    renderIAP();
    setActiveScreen('iap');
  });
  document.getElementById('step-btn-client').addEventListener('click', () => {
    const store = getStore();
    if (!store.workflow?.iapValidatedAt) { alert('Validate IAP to unlock Client Info Sheet.'); return; }
    renderClientSheet();
    setActiveScreen('client');
  });

  // Data console
  document.getElementById('open-console-btn').addEventListener('click', () => {
    refreshConsole();
    const con = document.getElementById('data-console');
    con.classList.add('is-open');
    con.setAttribute('aria-hidden', 'false');
  });
  document.getElementById('close-console-btn').addEventListener('click', () => {
    const con = document.getElementById('data-console');
    con.classList.remove('is-open');
    con.setAttribute('aria-hidden', 'true');
  });

  // Reset demo
  document.getElementById('reset-demo-btn').addEventListener('click', () => {
    if (!confirm('Reset all data and start over?')) return;
    localStorage.removeItem(STORE_KEY);
    location.reload();
  });
}

/* ==========================================================================
   INITIALIZATION
   ========================================================================== */

function init() {
  renderPQ();
  bindEvents();
  refreshConsole();

  // Restore screen based on workflow state
  const store = getStore();
  if (store.workflow?.clientCertifiedAt) { renderClientSheet(); setActiveScreen('client'); }
  else if (store.workflow?.iapValidatedAt) { renderIAP(); setActiveScreen('iap'); }
  else if (store.workflow?.planSubmittedAt) { renderPlan(); setActiveScreen('plan'); }
  else if (store.workflow?.pqSubmittedAt) { renderPQReview(); setActiveScreen('pq-review'); }
  else setActiveScreen('pq');
}

init();
