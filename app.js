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
/* PQ-only app — single stage */

/* ---------- UTILITY FUNCTIONS ---------- */
function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  Object.entries(attrs).forEach(([k, v]) => {
    if (k === 'className') node.className = v;
    else if (k === 'textContent') node.textContent = v;
    // innerHTML intentionally omitted for Lightning Web Security compatibility
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

function calcAgeAtDate(dob, targetDate) {
  if (!dob || !targetDate) return '';
  const d = new Date(dob);
  const t = new Date(targetDate);
  if (isNaN(d.getTime()) || isNaN(t.getTime())) return '';
  let age = t.getFullYear() - d.getFullYear();
  const m = t.getMonth() - d.getMonth();
  if (m < 0 || (m === 0 && t.getDate() < d.getDate())) age--;
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
  opts.push('Trust');

  return opts;
}

function uid() { return '_' + Math.random().toString(36).slice(2, 9); }

function clearChildren(node) {
  while (node.firstChild) node.removeChild(node.firstChild);
}

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
    colorTheme: 'blue',
    title: 'Household / Family',
    fields: [
      { key: 'client1FirstName', label: 'First Name', width: 'medium', required: true },
      { key: 'client1LastName', label: 'Last Name', width: 'medium', required: true },
      { key: 'client1Nickname', label: 'Nickname', width: 'medium' },
      { key: 'client1DOB', label: 'Date of Birth', type: 'date', width: 'medium' },
      { key: 'client1Age', label: 'Age', type: 'number', width: 'field-2' },
    ],
    conditionalBlocks: [
      {
        id: 'spouse-block',
        toggleLabel: '+ Add Spouse / Partner',
        hideLabel: '- Remove Spouse',
        flag: 'hasSpouse',
        fields: [
          { key: 'client2FirstName', label: 'First Name', width: 'medium' },
          { key: 'client2LastName', label: 'Last Name', width: 'medium' },
          { key: 'client2Nickname', label: 'Nickname', width: 'medium' },
          { key: 'client2DOB', label: 'Date of Birth', type: 'date', width: 'medium' },
          { key: 'client2Age', label: 'Age', type: 'number', width: 'field-2' },
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
    colorTheme: 'teal',
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
    colorTheme: 'violet',
    title: 'Employment',
    customRenderer: 'renderEmploymentSection'
  },

  /* ---- 4. ESTATE PLAN ---- */
  {
    id: 'goals',
    colorTheme: 'amber',
    title: 'Estate Plan',
    fields: [
      { key: 'estatePlan', label: 'Estate Plan', type: 'multiselect', options: ['Trust', 'Will', 'FPOA', 'MPOA'], width: 'medium' },
      { key: 'estatePlanYear', label: 'Year Established / Updated', width: 'medium' },
    ]
  },

  /* ---- 5. PROFESSIONAL RELATIONSHIPS ---- */
  {
    id: 'relationships',
    colorTheme: 'emerald',
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
    colorTheme: 'indigo',
    title: 'Assets',
    customRenderer: 'renderAssetsSection'
  },

  /* ---- 7. LIABILITIES ---- */
  {
    id: 'liabilities',
    colorTheme: 'rose',
    title: 'Liabilities',
    customRenderer: 'renderLiabilitiesSection'
  },

  /* ---- 8. INSURANCE ---- */
  {
    id: 'insurance',
    colorTheme: 'sky',
    title: 'Insurance',
    tables: [
      {
        key: 'policies',
        columns: [
          { key: 'company', label: 'Company' },
          { key: 'type', label: 'Type', type: 'select', options: ['Life', 'LTC', 'Disability', 'Umbrella/Liability'] },
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
    colorTheme: 'green',
    title: 'Income',
    customRenderer: 'renderIncomeSection'
  },

  /* ---- 10. TAXES & EXPENSES ---- */
  {
    id: 'taxesExpenses',
    colorTheme: 'orange',
    title: 'Taxes & Expenses',
    customRenderer: 'renderTaxesExpensesSection'
  },
];

/* Plan, IAP, and Client schemas removed — PQ-only app */

/* ==========================================================================
   RENDERING ENGINE
   ========================================================================== */

function buildFieldEl(field, sectionId, formType) {
  const widthMap = { 'wide': 'field-wide', 'medium': 'field-medium', 'field': 'field', 'field-2': 'field-2' };
  let cls = widthMap[field.width] || 'field';
  if (field.type === 'textarea') cls = 'field-full';
  if (field.type === 'computed') cls += ' field-computed';
  if (field.emphasis === 'total') cls += ' field-total';

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
  if (tableDef.title) {
    const titleEl = el('div', { className: 'table-title' + (tableDef.titleClass ? ' ' + tableDef.titleClass : '') });
    if (tableDef.titleClass) titleEl.appendChild(el('span', { className: 'table-title-dot', ariaHidden: 'true' }));
    titleEl.appendChild(document.createTextNode(tableDef.title));
    tools.appendChild(titleEl);
  }
  const addBtn = el('button', { type: 'button', className: 'btn-link', textContent: '+ Add Row' });
  tools.appendChild(addBtn);

  const wrap = el('div', { className: 'table-wrap' });
  const table = el('table');
  const thead = el('thead');
  const headRow = el('tr');
  tableDef.columns.forEach(c => {
    const th = el('th');
    th.appendChild(document.createTextNode(c.label));
    if (c.labelInfo) {
      const info = el('span', { className: 'th-info', textContent: 'i', tabIndex: 0 });
      info.setAttribute('title', c.labelInfo);
      info.setAttribute('aria-label', c.labelInfo);
      th.appendChild(info);
    }
    if (c.hidden) th.style.display = 'none';
    if (c.type === 'percent') th.classList.add('th-percent');
    if (c.colClass) th.classList.add(c.colClass);
    headRow.appendChild(th);
  });
  headRow.appendChild(el('th', { textContent: '', style: { width: '40px' } }));
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = el('tbody', { dataset: { tablePath: `${sectionId}.${tableDef.key}` } });
  table.appendChild(tbody);

  // Optional totals footer for currency columns
  let tfoot = null;
  if (tableDef.showTotals) {
    tfoot = el('tfoot');
    const footRow = el('tr');
    tableDef.columns.forEach((c, ci) => {
      const td = el('td');
      if (c.hidden) { td.style.display = 'none'; }
      else if (ci === 0 && c.type !== 'currency') { td.textContent = 'Total'; td.style.fontWeight = '600'; }
      else if (c.type === 'currency') { td.classList.add('tfoot-total'); td.dataset.totalCol = c.key; }
      footRow.appendChild(td);
    });
    footRow.appendChild(el('td')); // action column spacer
    tfoot.appendChild(footRow);
    table.appendChild(tfoot);
  }

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
          clearChildren(dl);
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
      if (col.colClass) td.classList.add(col.colClass);

      // Mark field read-only if listed in seed._readOnly
      const roKeys = seed._readOnly;
      if (Array.isArray(roKeys) && roKeys.includes(col.key) && !col.hidden) {
        inp.readOnly = true;
        inp.tabIndex = -1;
        inp.classList.add('field-readonly');
      }

      // Wrap with prefix/suffix symbol if needed
      if (!col.hidden && (col.type === 'currency' || col.type === 'percent')) {
        const wrap = el('div', { className: 'input-symbol-wrap' + (col.type === 'percent' ? ' input-wrap-percent' : '') });
        if (col.type === 'currency') wrap.appendChild(el('span', { className: 'input-symbol input-symbol--pre', textContent: '$' }));
        wrap.appendChild(inp);
        if (col.type === 'percent') wrap.appendChild(el('span', { className: 'input-symbol input-symbol--post', textContent: '%' }));
        if (col.type === 'percent') td.classList.add('td-percent');
        td.appendChild(wrap);
      } else {
        td.appendChild(inp);
      }
      tr.appendChild(td);
    });
    const actTd = el('td', { className: 'row-actions' });
    const hasReadOnly = Array.isArray(seed._readOnly) && seed._readOnly.length > 0;
    if (hasReadOnly) {
      // Auto-sourced row — hide the delete button
      actTd.style.visibility = 'hidden';
    } else {
      actTd.appendChild(el('button', {
        type: 'button', className: 'btn-icon', textContent: '\u00d7',
        onClick: () => { tr.remove(); reindex(); recalcTotals(); recalcPQComputed(); }
      }));
    }
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
    // Check all visible inputs (including read-only) for any data
    return Array.from(tr.querySelectorAll('input, select, textarea')).some(node => {
      if (!node || (node.tagName === 'INPUT' && node.type === 'hidden')) return false;
      return String(node.value || '').trim() !== '';
    });
  }

  // Recalculate tfoot totals for currency columns
  function recalcTotals() {
    if (!tfoot) return;
    const currCols = tableDef.columns.filter(c => c.type === 'currency');
    currCols.forEach(col => {
      let sum = 0;
      Array.from(tbody.children).forEach(tr => {
        const inp = tr.querySelector(`[data-path$=".${col.key}"]`);
        if (inp) sum += parseCommas(inp.value) || 0;
      });
      const cell = tfoot.querySelector(`[data-total-col="${col.key}"]`);
      if (cell) cell.textContent = sum ? formatCommas(sum) : '';
    });
  }

  if (tfoot) {
    tbody.addEventListener('input', recalcTotals);
  }

  // Expose addRow/reindex so external sync logic can add/update rows
  tbody._addRow = addRow;
  tbody._reindex = reindex;
  tbody._recalcTotals = recalcTotals;
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
      recalcTotals();
    }, 0);
  });

  // Seed rows
  const starters = tableDef.starterRows || [];
  const dataRows = existingData || [];
  const rows = dataRows.length ? dataRows : starters;
  rows.forEach(r => addRow(r));
  if (!rows.length) addRow();

  addBtn.addEventListener('click', () => { addRow(); recalcTotals(); });
  recalcTotals(); // initial calculation after seed rows
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
    ownHome: false, secondaryHome: false, additionalRE: false,
    taxDeferred: false, roth: false, taxable: false, cashCd: false, plan529: false, hsa: false,
  },
  employmentClient1: 'Employed',
  employmentClient2: 'Employed',
  openSections: { family: true },
  _initialized: false,
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
          // Mark synced fields read-only
          ['description', 'payment', 'amount', 'interestRate', 'term'].forEach(k => {
            const inp = emptyRow.querySelector(`[data-path$=".${k}"]`);
            if (inp) { inp.readOnly = true; inp.tabIndex = -1; inp.classList.add('field-readonly'); }
          });
          // Hide delete button
          const actTd = emptyRow.querySelector('.row-actions');
          if (actTd) actTd.style.visibility = 'hidden';
          liabTr = emptyRow;
        } else {
          // No empty row available — add a new one
          liabBody._addRow({ _reKey: autoKey, description: liabDesc, amount: loan, payment, interestRate: intRate, term, _readOnly: ['description', 'payment', 'amount', 'interestRate', 'term'] });
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
        expected.push({ _source: 'liability', _syncKey: 'liab_' + i, description: desc || 'Loan Payment', amount: payment * 12, notes: 'From liabilities', _readOnly: ['description', 'amount', 'notes'] });
      }
    });
  }

  // Collect expected auto-sourced expenses from insurance
  if (insBody) {
    Array.from(insBody.querySelectorAll('tr')).forEach((tr, i) => {
      const company = rowVal(tr, 'company');
      const premium = parseCommas(rowVal(tr, 'annualPremium')) || 0;
      if (premium > 0) {
        expected.push({ _source: 'insurance', _syncKey: 'ins_' + i, description: 'Insurance Premium - ' + (company || 'Unknown'), amount: premium, notes: 'From insurance', _readOnly: ['description', 'amount', 'notes'] });
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
        expected.push({ _source: 'asset', owner: ownership, description: desc || 'Asset Income', annualAmount: ebt * 12, notes: 'From assets', _readOnly: ['owner', 'description', 'annualAmount', 'notes'] });
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
      Object.entries(exp).forEach(([k, v]) => {
        if (k === '_readOnly') return;
        setRowVal(blankRow, k, String(v));
      });
      // Mark synced fields read-only
      if (Array.isArray(exp._readOnly)) {
        exp._readOnly.forEach(k => {
          const inp = blankRow.querySelector(`[data-path$=".${k}"]`);
          if (inp) { inp.readOnly = true; inp.tabIndex = -1; inp.classList.add('field-readonly'); }
        });
        const actTd = blankRow.querySelector('.row-actions');
        if (actTd) actTd.style.visibility = 'hidden';
      }
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
  const c1RetDate = form.querySelector('[data-path="employment.client1.retirementDate"]')?.value || '';
  const c2RetDate = form.querySelector('[data-path="employment.client2.retirementDate"]')?.value || '';

  empBody.querySelectorAll('tr').forEach(row => {
    const srcInp = row.querySelector('input[data-path$="._source"]');
    const ownerInp = row.querySelector('input[data-path$=".owner"]');
    const descInp = row.querySelector('input[data-path$=".description"]');
    const retDateInp = row.querySelector('input[data-path$=".retirementDate"]');
    if (!srcInp || !ownerInp) return;
    if (srcInp.value === 'client1') {
      ownerInp.value = c1Name;
      if (descInp) descInp.value = c1Employer;
      if (retDateInp) retDateInp.value = c1RetDate;
    } else if (srcInp.value === 'client2') {
      ownerInp.value = c2Name;
      if (descInp) descInp.value = c2Employer;
      if (retDateInp) retDateInp.value = c2RetDate;
    }
  });
}

function refreshOwnerDataLists() {
  const form = document.getElementById('pq-form');
  if (!form) return;
  // Build current owner options from live field values
  const c1 = [
    (form.querySelector('[data-path="family.client1FirstName"]')?.value || '').trim(),
    (form.querySelector('[data-path="family.client1LastName"]')?.value || '').trim()
  ].filter(Boolean).join(' ');
  const c2 = [
    (form.querySelector('[data-path="family.client2FirstName"]')?.value || '').trim(),
    (form.querySelector('[data-path="family.client2LastName"]')?.value || '').trim()
  ].filter(Boolean).join(' ');
  const opts = [];
  if (c1) opts.push(c1);
  if ((pqState.hasSpouse || c2) && c2) opts.push(c2);
  if (c1 && c2) opts.push('Joint');
  opts.push('Trust');

  // Refresh every datalist whose id contains owner, ownership, or insured
  document.querySelectorAll('datalist').forEach(dl => {
    if (/owner|ownership|insured|beneficiary/i.test(dl.id)) {
      clearChildren(dl);
      opts.forEach(o => {
        const opt = document.createElement('option');
        opt.value = o;
        dl.appendChild(opt);
      });
    }
  });
}

function setupEmploymentNameSync() {
  const form = document.getElementById('pq-form');
  if (!form) return;
  const watchPaths = [
    'family.client1FirstName', 'family.client1LastName',
    'family.client2FirstName', 'family.client2LastName',
    'employment.client1.employer', 'employment.client2.employer',
    'employment.client1.retirementDate', 'employment.client2.retirementDate'
  ];
  watchPaths.forEach(path => {
    const inp = form.querySelector(`[data-path="${path}"]`);
    if (inp) {
      inp.addEventListener('input', syncEmploymentFields);
      inp.addEventListener('change', syncEmploymentFields);
      inp.addEventListener('input', refreshOwnerDataLists);
    }
  });
}

function clearSpouseData() {
  const stripClient2 = (obj) => {
    if (!obj || typeof obj !== 'object') return;
    Object.keys(obj).forEach(k => {
      if (k === 'client2') delete obj[k];
      else if (typeof k === 'string' && k.startsWith('client2')) delete obj[k];
      else if (obj[k] && typeof obj[k] === 'object') stripClient2(obj[k]);
    });
  };
  const clearForForm = (formEl) => {
    if (!formEl) return;
    formEl.querySelectorAll('[data-path]').forEach(inp => {
      const p = inp.dataset.path || '';
      if (p.includes('client2') || p === 'family.yearsMarried') {
        if (inp.type === 'checkbox' || inp.type === 'radio') inp.checked = false;
        else inp.value = '';
      }
    });
  };

  stripClient2(_pqDraft);
  if (_pqDraft?.family) delete _pqDraft.family.yearsMarried;

  const store = getStore();
  if (store.pqTemplate) {
    stripClient2(store.pqTemplate);
    if (store.pqTemplate.family) delete store.pqTemplate.family.yearsMarried;
    saveStore(store);
  }

  clearForForm(document.getElementById('pq-form'));
  pqState.employmentClient2 = null;
}

function renderPQ() {
  // Snapshot any in-flight form data before rebuilding DOM
  snapshotPQForm();

  const host = document.getElementById('pq-form-sections');
  clearChildren(host);
  const store = getStore();
  const saved = store.pqTemplate || {};
  // Merge: saved store data as base, then overlay the in-flight draft
  const pq = deepMerge(deepClone(saved), _pqDraft);

  // Restore state from existing data — only on initial load so explicit toggles (e.g. Remove Spouse) aren't undone
  if (!pqState._initialized) {
    if (pq.family?.client2FirstName || pq.family?.hasSpouse) pqState.hasSpouse = true;
    if (
      pq.family?.hasChildren ||
      (pq.family?.children && pq.family.children.length) ||
      (pq.family?.grandchildren && pq.family.grandchildren.length)
    ) pqState.hasChildren = true;
    if (pq.employment?.client1?.status) pqState.employmentClient1 = pq.employment.client1.status;
    if (pq.employment?.client2?.status) pqState.employmentClient2 = pq.employment.client2.status;
  }
  pqState._initialized = true;
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
  requestAnimationFrame(updateSectionProgress);
}

let _recalcTimer = null;
function setupFormRecalc() {
  const form = document.getElementById('pq-form');
  if (!form || form._recalcWired) return;
  form._recalcWired = true;
  form.addEventListener('input', () => {
    clearTimeout(_recalcTimer);
    _recalcTimer = setTimeout(() => { recalcPQComputed(); updateSectionProgress(); }, 150);
  });
}

function updateSectionProgress() {
  document.querySelectorAll('.form-section').forEach(section => {
    const label = section.querySelector('.section-progress-label');
    const fill = section.querySelector('.section-progress-fill');
    const progressEl = section.querySelector('.section-progress');
    if (!label || !fill || !progressEl) return;

    const content = section.querySelector('.section-content');
    if (!content) return;

    const inputs = Array.from(content.querySelectorAll(
      'input:not([type=radio]):not([type=checkbox]):not([type=hidden]):not([readonly]):not([tabindex="-1"]),' +
      'select, textarea:not([readonly])'
    ));

    const total = inputs.length;
    if (total === 0) {
      label.textContent = '';
      fill.style.width = '0%';
      progressEl.classList.remove('is-complete');
      return;
    }

    const filled = inputs.filter(inp => String(inp.value || '').trim() !== '').length;
    const pct = Math.round((filled / total) * 100);
    fill.style.width = pct + '%';

    if (pct === 100) {
      label.textContent = '\u2713';
      progressEl.classList.add('is-complete');
    } else {
      label.textContent = filled + '/' + total;
      progressEl.classList.remove('is-complete');
    }
  });
}

function buildPQSection(section, pqData) {
  const wrapper = el('section', { className: `form-section theme-${section.colorTheme || 'default'}` });

  const header = el('div', { className: 'section-header' });
  header.appendChild(el('h2', { textContent: section.title }));
  const progressEl = el('div', { className: 'section-progress' });
  progressEl.appendChild(el('span', { className: 'section-progress-label' }));
  const progressTrack = el('div', { className: 'section-progress-track' });
  progressTrack.appendChild(el('div', { className: 'section-progress-fill' }));
  progressEl.appendChild(progressTrack);
  header.appendChild(progressEl);
  wrapper.appendChild(header);

  const content = el('div', { className: 'section-content' });

  if (section.note) content.appendChild(el('p', { className: 'section-note', textContent: section.note }));

  // Family section uses a side-by-side Client + Spouse split
  const familySplit = section.id === 'family' ? el('div', { className: 'family-split' }) : null;
  const getFamilyCol = (isSpouse) => {
    const sel = isSpouse ? '.family-split-spouse-col' : '.family-split-client-col';
    let col = familySplit.querySelector(sel);
    if (!col) {
      col = el('div', { className: 'family-split-col ' + (isSpouse ? 'family-split-spouse-col' : 'family-split-client-col') });
      col.appendChild(el('div', { className: 'family-split-label', textContent: isSpouse ? 'Spouse / Partner' : 'Client' }));
      familySplit.appendChild(col);
    }
    return col;
  };

  // Regular fields
  if (section.fields?.length) {
    const grid = el('div', { className: 'grid' });
    section.fields.forEach(f => {
      if (f.showIf && !pqState[f.showIf]) return;
      grid.appendChild(buildFieldEl(f, section.id, 'pq'));
    });
    if (familySplit) getFamilyCol(false).appendChild(grid);
    else content.appendChild(grid);
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
        const willEnable = !pqState[block.flag];
        pqState[block.flag] = willEnable;
        btn.textContent = willEnable ? block.hideLabel : block.toggleLabel;
        blockContent.classList.toggle('hidden', !willEnable);
        pqState.openSections[section.id] = true; // keep parent open
        // When disabling the spouse block, wipe all spouse-owned data from draft + store
        if (!willEnable && block.id === 'spouse-block') clearSpouseData();
        renderPQ();
        focusFirstInput(section.id);
      });

      if (familySplit && block.id === 'spouse-block') {
        // Hide the "Spouse / Partner" label while the spouse block is collapsed so the
        // Add button sits exactly where the heading will appear. When the spouse is
        // active, the heading returns and the Remove button aligns to its right.
        const col = getFamilyCol(true);
        const labelEl = col.querySelector('.family-split-label');
        if (labelEl) {
          labelEl.classList.add('family-split-label--with-action');
          labelEl.classList.toggle('family-split-label--collapsed', !pqState[block.flag]);
          btn.classList.add('family-split-header-btn');
          if (pqState[block.flag]) btn.classList.add('family-split-header-btn--remove');
          else btn.classList.add('family-split-header-btn--add');
          labelEl.appendChild(btn);
        }
        container.appendChild(blockContent);
        col.appendChild(container);
      } else {
        container.append(btn, blockContent);
        content.appendChild(container);
      }
    });
  }

  if (familySplit) content.appendChild(familySplit);

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
              col.key === 'owner' || col.key === 'insured' || col.key === 'beneficiary'
                ? { ...col, type: 'datalist', options: ownerOptions }
                : col
            )),
          }
        : t;
      content.appendChild(buildTableEl(section.id, tableDef, existing));
    });
  }

  wrapper.appendChild(content);
  return wrapper;
}

/* ---------- SECTION SHELL HELPER ---------- */
function buildCollapsibleShell(sectionId, title) {
  const themeMap = {
    employment: 'violet', assets: 'indigo', liabilities: 'rose',
    income: 'green', taxesExpenses: 'orange',
  };
  const wrapper = el('section', { className: `form-section theme-${themeMap[sectionId] || 'default'}` });
  const header = el('div', { className: 'section-header' });
  header.appendChild(el('h2', { textContent: title }));
  const progressEl = el('div', { className: 'section-progress' });
  progressEl.appendChild(el('span', { className: 'section-progress-label' }));
  const progressTrack = el('div', { className: 'section-progress-track' });
  progressTrack.appendChild(el('div', { className: 'section-progress-fill' }));
  progressEl.appendChild(progressTrack);
  header.appendChild(progressEl);
  wrapper.appendChild(header);

  const content = el('div', { className: 'section-content' });

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
    grid1.appendChild(buildFieldEl({ key: 'client1.retirementAge', label: 'Retirement Age', type: 'number', width: 'field' }, 'employment', 'pq'));
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
      grid2.appendChild(buildFieldEl({ key: 'client2.retirementAge', label: 'Retirement Age', type: 'number', width: 'field' }, 'employment', 'pq'));
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

  // --- Real Estate subsection ---
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Real Estate' }));
  const reChecks = el('div', { className: 'checkbox-group' });
  const reCheckDefs = [
    { key: 'ownHome', label: 'Own your home?' },
    { key: 'secondaryHome', label: 'Secondary home(s)?' },
    { key: 'additionalRE', label: 'Additional real estate?' },
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
  ];

  const ownershipOptions = getOwnerNameOptions(pqData);

  const reColumns = [
    { key: '_autoKey',      label: '',               hidden: true },
    { key: 'description',   label: 'Description',    colClass: 'col-wide' },
    { key: 'marketValue',   label: 'Market Value',   type: 'currency' },
    { key: 'remainingLoan', label: 'Remaining Loan', type: 'currency' },
    { key: 'interestRate',  label: 'Int. Rate',      type: 'percent' },
    { key: 'term',          label: 'Term',           colClass: 'col-narrow', labelInfo: 'In years' },
    { key: 'payment',       label: 'Payment (P&I)',   type: 'currency', labelInfo: 'Monthly' },
    { key: 'yearAcquired',  label: 'Year Acquired',  colClass: 'col-narrow' },
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
      title: '',
      columns: reColumns,
      starterRows: [],
      showTotals: true,
    };
    content.appendChild(buildTableEl('assets', reTable, mergedRows.length ? mergedRows : null));
  }

  // --- Business and Other subsection ---
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Business and Other', style: { marginTop: '18px' } }));
  content.appendChild(buildTableEl('assets', {
    key: 'businessOther',
    title: '',
    showTotals: true,
    columns: [
      { key: 'description', label: 'Description' },
      { key: 'marketValue', label: 'Market Value', type: 'currency' },
      { key: 'costBasis', label: 'Cost Basis', type: 'currency' },
      { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
    ]
  }, getDeep(pqData, 'assets.businessOther')));

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
      title: 'Tax-Deferred Retirement',
      titleClass: 'tone-tax-deferred',
      showTotals: true,
      columns: [
        { key: 'custodian', label: 'Custodian' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'personalAdditions', label: 'Personal Additions', type: 'currency', labelInfo: 'Yearly' },
        { key: 'companyMatch', label: 'Company Match', type: 'currency' },
        { key: 'type', label: 'Type' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
        { key: 'beneficiary', label: 'Beneficiary', type: 'datalist', options: ownershipOptions },
      ]
    }, getDeep(pqData, 'assets.taxDeferred')));
  }

  if (pqState.assetChecks.roth) {
    content.appendChild(buildTableEl('assets', {
      key: 'roth',
      title: 'Tax-Free Roth',
      titleClass: 'tone-tax-free',
      showTotals: true,
      columns: [
        { key: 'custodian', label: 'Custodian' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'personalAdditions', label: 'Personal Additions', type: 'currency', labelInfo: 'Yearly' },
        { key: 'companyMatch', label: 'Company Match', type: 'currency' },
        { key: 'type', label: 'Type' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
        { key: 'beneficiary', label: 'Beneficiary', type: 'datalist', options: ownershipOptions },
      ]
    }, getDeep(pqData, 'assets.roth')));
  }

  if (pqState.assetChecks.taxable) {
    content.appendChild(buildTableEl('assets', {
      key: 'taxable',
      title: 'Taxable Non-Retirement',
      titleClass: 'tone-taxable',
      showTotals: true,
      columns: [
        { key: 'custodian', label: 'Custodian' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'personalAdditions', label: 'Personal Additions', type: 'currency', labelInfo: 'Yearly' },
        { key: 'costBasis', label: 'Cost Basis', type: 'currency' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
        { key: 'variableFixed', label: 'Variable / Fixed' },
        { key: 'issueDate', label: 'Issue Date', type: 'date' },
        { key: 'beneficiary', label: 'Beneficiary', type: 'datalist', options: ownershipOptions },
      ]
    }, getDeep(pqData, 'assets.taxable')));
  }

  // Cash/CD, 529, HSA — simpler tables
  if (pqState.assetChecks.cashCd) {
    content.appendChild(buildTableEl('assets', {
      key: 'cashCd',
      title: "Cash & CD's",
      showTotals: true,
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
      showTotals: true,
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
      showTotals: true,
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
        _readOnly: ['description', 'payment', 'amount', 'interestRate', 'term'],
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
    // Tag RE-keyed rows as read-only
    merged.forEach(r => {
      if (r._reKey) r._readOnly = ['description', 'payment', 'amount', 'interestRate', 'term'];
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
    title: '',
    columns: [
      { key: '_reKey', label: '', hidden: true },
      { key: 'description', label: 'Description' },
      { key: 'payment', label: 'Payment (P&I)', type: 'currency', labelInfo: 'Monthly' },
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
  const family = pqData?.family || {};
  const c1Name = [(family.client1FirstName || '').trim(), (family.client1LastName || '').trim()].filter(Boolean).join(' ');
  const c2Name = [(family.client2FirstName || '').trim(), (family.client2LastName || '').trim()].filter(Boolean).join(' ');

  // Auto-sourced employment rows
  const autoEmp = [];
  const c1RetDate = pqData?.employment?.client1?.retirementDate || '';
  const c2RetDate = pqData?.employment?.client2?.retirementDate || '';
  if (pqState.employmentClient1 === 'Employed') {
    autoEmp.push({ _source: 'client1', owner: c1Name, description: pqData?.employment?.client1?.employer || '', retirementDate: c1RetDate, _readOnly: ['owner', 'description', 'retirementDate'] });
  }
  if (pqState.hasSpouse && pqState.employmentClient2 === 'Employed') {
    autoEmp.push({ _source: 'client2', owner: c2Name, description: pqData?.employment?.client2?.employer || '', retirementDate: c2RetDate, _readOnly: ['owner', 'description', 'retirementDate'] });
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
        return { ...prev, owner: auto.owner, description: auto.description || prev.description, retirementDate: auto.retirementDate, _readOnly: auto._readOnly };
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
  const existingPension = getDeep(pqData, 'income.pension');
  content.appendChild(buildTableEl('income', {
    key: 'pension',
    title: 'Pension',
    columns: [
      { key: 'owner', label: 'Owner', type: 'datalist', options: ownerOptions },
      { key: 'startDate', label: 'Start Date', type: 'date' },
      { key: 'annualAmount', label: 'Annual Amount', type: 'currency' },
      { key: 'survivorBenefit', label: 'Survivor Benefit', type: 'percent' },
      { key: 'cola', label: 'COLA %', type: 'percent' },
    ],
    starterRows: [],
  }, existingPension));

  // ── 4. Other Income ──
  // Auto-source from assets with Income (EBT)
  const autoOther = [];
  const reData = getDeep(pqData, 'assets.realEstate') || [];
  reData.forEach(r => {
    if (r?.incomeEBT && parseFloat(r.incomeEBT) > 0) {
      autoOther.push({ _source: 'asset', owner: r.ownership || '', description: r.description || 'Asset Income', annualAmount: parseFloat(r.incomeEBT) * 12, notes: 'From assets', _readOnly: ['owner', 'description', 'annualAmount', 'notes'] });
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
  totalRow.appendChild(buildFieldEl({ key: 'totalIncome', label: 'Total Income', type: 'computed', width: 'medium', emphasis: 'total' }, 'income', 'pq'));
  content.appendChild(totalRow);

  return wrapper;
}

/* ----- TAXES & EXPENSES ----- */
function renderTaxesExpensesSection(section, pqData) {
  const { wrapper, content } = buildCollapsibleShell(section.id, section.title);

  // Helper: build an inline group label row with a right-aligned total badge
  const groupRow = (labelText, totalPath, hintId) => {
    const row = el('div', { className: 'taxexp-group-row' });
    row.appendChild(el('div', { className: 'taxexp-group-label', textContent: labelText }));
    if (totalPath) {
      const badge = el('div', { className: 'taxexp-total-badge' });
      const val = el('input', { type: 'text', className: 'taxexp-total-val', readOnly: true });
      val.dataset.path = totalPath;
      val.name = totalPath;
      badge.appendChild(val);
      if (hintId) badge.appendChild(el('span', { className: 'taxexp-total-hint', id: hintId }));
      row.appendChild(badge);
    }
    return row;
  };

  // ─── PRIOR YEAR TAX INFORMATION ───
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Prior Year Tax Information' }));

  content.appendChild(groupRow('Income & Deductions'));
  const incGrid = el('div', { className: 'grid' });
  [
    { key: 'taxableIncome', label: 'Taxable Income', type: 'currency', width: 'medium' },
    { key: 'standardItemDeduction', label: 'Std/Item Deduction', type: 'currency', width: 'medium' },
    { key: 'capLossCarryForward', label: 'Cap Loss Carry Forward', type: 'currency', width: 'medium' },
  ].forEach(f => incGrid.appendChild(buildFieldEl(f, 'taxesExpenses', 'pq')));
  content.appendChild(incGrid);

  content.appendChild(groupRow('Taxes Paid'));
  const paidGrid = el('div', { className: 'grid' });
  [
    { key: 'federalTax', label: 'Federal', type: 'currency', width: 'medium' },
    { key: 'stateTax', label: 'State', type: 'currency', width: 'medium' },
    { key: 'ficaTax', label: 'FICA', type: 'currency', width: 'medium' },
  ].forEach(f => paidGrid.appendChild(buildFieldEl(f, 'taxesExpenses', 'pq')));
  content.appendChild(paidGrid);

  // ─── EXPENSE DATA PREP ───
  const isLivingExpenseRow = (r) => r && typeof r.description === 'string' && r.description.trim().toLowerCase() === 'living expenses';
  let livingExpVal = getDeep(pqData, 'taxesExpenses.livingExpenses');
  const legacyExisting = getDeep(pqData, 'taxesExpenses.expenses');
  if (Array.isArray(legacyExisting)) {
    for (let i = legacyExisting.length - 1; i >= 0; i--) {
      const row = legacyExisting[i];
      if (!isLivingExpenseRow(row)) continue;
      if ((livingExpVal == null || livingExpVal === '') && row.amount) {
        livingExpVal = row.amount;
        setDeep(pqData, 'taxesExpenses.livingExpenses', livingExpVal);
      }
      legacyExisting.splice(i, 1);
    }
  }

  const autoExpenses = [];
  const liabData = getDeep(pqData, 'liabilities.items') || [];
  liabData.forEach(l => {
    if (l?.payment && parseFloat(l.payment) > 0) {
      autoExpenses.push({ _source: 'liability', description: l.description || 'Loan Payment', amount: parseFloat(l.payment) * 12, notes: 'From liabilities', _readOnly: ['description', 'amount', 'notes'] });
    }
  });

  const insuranceData = getDeep(pqData, 'insurance.policies') || [];
  insuranceData.forEach(p => {
    if (p?.annualPremium && parseFloat(p.annualPremium) > 0) {
      autoExpenses.push({ _source: 'insurance', description: 'Insurance Premium - ' + (p.company || 'Unknown'), amount: parseFloat(p.annualPremium), notes: 'From insurance', _readOnly: ['description', 'amount', 'notes'] });
    }
  });

  let existingExp = getDeep(pqData, 'taxesExpenses.expenses');
  let mergedExp;
  if (existingExp?.length) {
    const manualRows = existingExp.filter(r => r && !r._source && !isLivingExpenseRow(r));
    mergedExp = [...autoExpenses, ...manualRows];
  } else {
    mergedExp = [...autoExpenses];
  }

  // ─── ANNUAL EXPENSES ───
  content.appendChild(el('h3', { className: 'subsection-title', textContent: 'Annual Expenses', style: { marginTop: '14px' } }));

  content.appendChild(buildTableEl('taxesExpenses', {
    key: 'expenses',
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

  const livingRow = el('div', { className: 'grid', style: { marginTop: '10px' } });
  livingRow.appendChild(buildFieldEl({ key: 'livingExpenses', label: 'Living Expenses', type: 'currency', width: 'medium' }, 'taxesExpenses', 'pq'));
  livingRow.appendChild(buildFieldEl({ key: 'expenseTotal', label: 'Total Expenses', type: 'computed', width: 'medium', emphasis: 'total' }, 'taxesExpenses', 'pq'));
  content.appendChild(livingRow);

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

  // Retirement ages (auto-calc from DOB + retirement date)
  const retDate1 = form.querySelector('[data-path="employment.client1.retirementDate"]');
  const retAge1 = form.querySelector('[data-path="employment.client1.retirementAge"]');
  if (dob1 && retDate1 && retAge1 && retDate1.value) {
    const v = calcAgeAtDate(dob1.value, retDate1.value);
    if (v !== '') retAge1.value = v;
  }

  const retDate2 = form.querySelector('[data-path="employment.client2.retirementDate"]');
  const retAge2 = form.querySelector('[data-path="employment.client2.retirementAge"]');
  if (dob2 && retDate2 && retAge2 && retDate2.value) {
    const v = calcAgeAtDate(dob2.value, retDate2.value);
    if (v !== '') retAge2.value = v;
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

  // Effective tax rate hint
  const taxableInc = parseCommas(form.querySelector('[data-path="taxesExpenses.taxableIncome"]')?.value) || 0;
  const effHint = document.getElementById('taxexp-effective-hint');
  if (effHint) {
    if (taxableInc > 0 && totalTax > 0) {
      const rate = (totalTax / taxableInc) * 100;
      effHint.textContent = rate.toFixed(1) + '% effective rate on taxable income';
    } else {
      effHint.textContent = 'Federal + State + FICA';
    }
  }

  // Total expenses (separate Living Expenses field + expense line rows)
  const expInputs = form.querySelectorAll('[data-path^="taxesExpenses.expenses["]');
  let totalExp = 0;
  expInputs.forEach(inp => {
    if (inp.dataset.path.endsWith('.amount')) totalExp += parseCommas(inp.value) || 0;
  });
  const livingExpInput = form.querySelector('[data-path="taxesExpenses.livingExpenses"]');
  if (livingExpInput) totalExp += parseCommas(livingExpInput.value) || 0;
  const totalExpEl = form.querySelector('[data-path="taxesExpenses.expenseTotal"]');
  if (totalExpEl) totalExpEl.value = fmt$(totalExp);

  renderPQDashboard();
}

/* ---------- PQ FLOATING DASHBOARD ---------- */
// Persistent notes textarea — survives dashboard re-renders
let _notesTA = null;
let _goalsTA = null;
let _notesDeferPending = false;
let _notesDrawerWired = false;

function ensureNotesDrawer() {
  const drawer = document.getElementById('notes-drawer');
  if (!drawer) return;
  _goalsTA = document.getElementById('drawer-goals');
  _notesTA = document.getElementById('drawer-notes');
  if (!_goalsTA || !_notesTA) return;

  const store = getStore();
  if (_goalsTA.value === '' && store.pqTemplate?.goals?.goals) _goalsTA.value = store.pqTemplate.goals.goals;
  if (_notesTA.value === '' && store._advisorNotes) _notesTA.value = store._advisorNotes;

  if (_notesDrawerWired) return;
  _notesDrawerWired = true;

  _goalsTA.addEventListener('input', () => {
    const s = getStore();
    if (!s.pqTemplate) s.pqTemplate = {};
    if (!s.pqTemplate.goals) s.pqTemplate.goals = {};
    s.pqTemplate.goals.goals = _goalsTA.value;
    saveStore(s);
  });
  _notesTA.addEventListener('input', () => {
    const s = getStore();
    s._advisorNotes = _notesTA.value;
    saveStore(s);
  });
  [_goalsTA, _notesTA].forEach(ta => {
    ta.addEventListener('keydown', e => { if (e.key === 'Enter') e.stopPropagation(); });
  });

  const tab = document.getElementById('notes-drawer-tab');
  const closeBtn = document.getElementById('notes-drawer-close');

  const open = () => {
    drawer.classList.add('is-open');
    drawer.setAttribute('aria-hidden', 'false');
  };
  const close = () => {
    drawer.classList.remove('is-open');
    drawer.setAttribute('aria-hidden', 'true');
  };
  tab.addEventListener('click', (e) => {
    e.stopPropagation();
    drawer.classList.contains('is-open') ? close() : open();
  });
  closeBtn.addEventListener('click', close);
  document.addEventListener('mousedown', (e) => {
    if (!drawer.classList.contains('is-open')) return;
    if (drawer.contains(e.target)) return;
    close();
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && drawer.classList.contains('is-open')) close();
  });
}

function renderPQDashboard() {
  const host = document.getElementById('pq-dashboard');
  if (!host) return;

  // If either persistent textarea is focused, defer the rebuild so we don't steal focus
  if ((_notesTA && document.activeElement === _notesTA) ||
      (_goalsTA && document.activeElement === _goalsTA)) {
    if (!_notesDeferPending) {
      _notesDeferPending = true;
      document.activeElement.addEventListener('blur', function () {
        _notesDeferPending = false;
        renderPQDashboard();
      }, { once: true });
    }
    return;
  }

  clearChildren(host);
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
  totalAssets += sumInputs('[data-path^="assets.businessOther["]', '.marketValue');

  // Total Savings = personal additions only (excludes company match)
  let totalSavings = 0;
  totalSavings += sumInputs('[data-path$=".personalAdditions"]');

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
  // Separate Living Expenses field rolls into Non-Liability Expenses
  const livingExpDashInput = form.querySelector('[data-path="taxesExpenses.livingExpenses"]');
  nonLiabilityExpenses += parseCommas(livingExpDashInput?.value) || 0;

  // Derived metrics
  const totalNetWorth = totalAssets - totalLiabilities;
  const dti = totalIncome > 0 ? (totalLiabilities / totalIncome * 100) : 0;
  const savingsRatio = totalIncome > 0 ? (totalSavings / totalIncome * 100) : 0;
  const cashFlow = totalIncome - totalSavings - liabilityPayments - totalTax - nonLiabilityExpenses;

  // ── Balance Sheet metrics (used in combo card) ──
  // Placeholder — actual balance sheet rows are built later where breakdown values exist.
  const bsMetrics = el('div', { className: 'bs-metrics-placeholder' });

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

  // Ratios appended into the same card, separated by a divider
  const addRatioRow = (label, value, colorClass, isFirst) => {
    const li = el('li');
    if (isFirst) { li.style.borderTop = '1px solid var(--border)'; li.style.paddingTop = '6px'; li.style.marginTop = '4px'; }
    li.appendChild(el('span', { textContent: label }));
    li.appendChild(el('strong', { className: `metric-value ${colorClass}`, textContent: value }));
    cfMetrics.appendChild(li);
  };
  addRatioRow('Debt-to-Income', fmtPct(dti), dti > 40 ? 'negative' : dti > 20 ? 'warning' : 'positive', true);
  addRatioRow('Savings Ratio', fmtPct(savingsRatio), savingsRatio >= 20 ? 'positive' : savingsRatio >= 10 ? 'warning' : 'negative', false);

  card2.appendChild(cfMetrics);

  // Goals + Advisor Notes now live in the slide-out drawer (see initNotesDrawer)
  ensureNotesDrawer();

  // ── Tax Triangle data ──
  const taxFree = sumInputs('[data-path^="assets.roth["]', '.marketValue')
                + sumInputs('[data-path^="assets.hsa["]', '.marketValue');
  const taxDeferred = sumInputs('[data-path^="assets.taxDeferred["]', '.marketValue');
  const taxable = sumInputs('[data-path^="assets.taxable["]', '.marketValue')
                + sumInputs('[data-path^="assets.cashCd["]', '.marketValue');

  // ── Combined Balance Sheet + Tax Triangle card ──
  const comboCard = el('div', { className: 'dashboard-card balance-triangle-card' });
  const comboInner = el('div', { className: 'balance-triangle-inner' });

  // Left side: Balance Sheet metrics
  const bsCol = el('div', { className: 'balance-col' });
  bsCol.appendChild(el('h3', { textContent: 'Balance Sheet' }));

  // Asset subtotals (investable = everything minus real estate & business/other)
  const realEstateAssets = sumInputs('[data-path^="assets.realEstate["]', '.marketValue');
  const businessOtherAssets = sumInputs('[data-path^="assets.businessOther["]', '.marketValue');
  const investableAssets = totalAssets - realEstateAssets - businessOtherAssets;

  // Breakdown wrapper — groups the three asset categories as siblings
  const bsBreakdown = el('div', { className: 'bs-breakdown' });
  const makeAssetRow = (label, value, swatch) => {
    const row = el('div', { className: 'bs-asset-row' });
    const left = el('div', { className: 'bs-asset-left' });
    left.appendChild(el('span', { className: `bs-swatch bs-swatch-${swatch}` }));
    left.appendChild(el('span', { className: 'bs-asset-label', textContent: label }));
    row.appendChild(left);
    row.appendChild(el('strong', { className: 'bs-asset-value', textContent: fmt$(value) }));
    return row;
  };
  bsBreakdown.appendChild(makeAssetRow('Investable Assets', investableAssets, 'invest'));
  bsBreakdown.appendChild(makeAssetRow('Real Estate', realEstateAssets, 'realestate'));
  bsBreakdown.appendChild(makeAssetRow('Business / Other', businessOtherAssets, 'business'));
  bsCol.appendChild(bsBreakdown);

  // Total Assets — prominent summary row directly beneath the breakdown
  const totalAssetsBox = el('div', { className: 'bs-total-assets' });
  totalAssetsBox.appendChild(el('span', { className: 'bs-total-label', textContent: 'Total Assets' }));
  totalAssetsBox.appendChild(el('strong', { className: 'bs-total-value', textContent: fmt$(totalAssets) }));
  bsCol.appendChild(totalAssetsBox);

  // Liabilities line
  const liabLine = el('div', { className: 'bs-liab-line' });
  liabLine.appendChild(el('span', { textContent: 'Total Liabilities' }));
  liabLine.appendChild(el('strong', { className: 'metric-value', textContent: fmt$(totalLiabilities) }));
  bsCol.appendChild(liabLine);

  // Net worth highlight
  const nwBox = el('div', { className: 'bs-net-worth' });
  nwBox.appendChild(el('span', { textContent: 'Net Worth' }));
  nwBox.appendChild(el('strong', { className: `metric-value ${totalNetWorth >= 0 ? 'positive' : 'negative'}`, textContent: fmt$(totalNetWorth) }));
  bsCol.appendChild(nwBox);

  comboInner.appendChild(bsCol);

  // Right side: Tax Triangle — circles at vertices
  const triCol = el('div', { className: 'triangle-col' });
  const triWrapper = el('div', { className: 'tax-triangle-wrapper' });
  triWrapper.appendChild(buildTaxTriangleSVG(taxFree, taxable, taxDeferred));
  triCol.appendChild(triWrapper);
  comboInner.appendChild(triCol);
  comboCard.appendChild(comboInner);

  host.append(comboCard, card2);

  // Clear left panel (no longer used)
  const notesHost = document.getElementById('pq-notes-panel');
  if (notesHost) clearChildren(notesHost);
}

/* Plan Review, Plan Design, IAP, and Client Info Sheet stages removed — PQ-only app */

/* ---------- TAX TRIANGLE SVG BUILDER ---------- */
function buildTaxTriangleSVG(taxFree, taxable, taxDeferred, opts = {}) {
  const svgNS = 'http://www.w3.org/2000/svg';
  const idSuffix = opts.idSuffix || '';
  const svg = document.createElementNS(svgNS, 'svg');
  svg.setAttribute('viewBox', '0 0 320 280');
  svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');

  const defs = document.createElementNS(svgNS, 'defs');

  const filter = document.createElementNS(svgNS, 'filter');
  const filterId = 'triShadow' + idSuffix;
  filter.setAttribute('id', filterId);
  filter.setAttribute('x', '-20%'); filter.setAttribute('y', '-20%');
  filter.setAttribute('width', '140%'); filter.setAttribute('height', '140%');
  const feGauss = document.createElementNS(svgNS, 'feGaussianBlur');
  feGauss.setAttribute('in', 'SourceAlpha'); feGauss.setAttribute('stdDeviation', '3');
  const feOff = document.createElementNS(svgNS, 'feOffset');
  feOff.setAttribute('dx', '0'); feOff.setAttribute('dy', '2');
  const feMerge = document.createElementNS(svgNS, 'feMerge');
  const fm1 = document.createElementNS(svgNS, 'feMergeNode');
  const fm2 = document.createElementNS(svgNS, 'feMergeNode');
  fm2.setAttribute('in', 'SourceGraphic');
  feMerge.append(fm1, fm2);
  filter.append(feGauss, feOff, feMerge);
  defs.appendChild(filter);

  const makeGrad = (id, c1, c2) => {
    const g = document.createElementNS(svgNS, 'linearGradient');
    g.setAttribute('id', id); g.setAttribute('x1', '0'); g.setAttribute('y1', '0');
    g.setAttribute('x2', '0'); g.setAttribute('y2', '1');
    const s1 = document.createElementNS(svgNS, 'stop');
    s1.setAttribute('offset', '0%'); s1.setAttribute('stop-color', c1);
    const s2 = document.createElementNS(svgNS, 'stop');
    s2.setAttribute('offset', '100%'); s2.setAttribute('stop-color', c2);
    g.append(s1, s2);
    defs.appendChild(g);
  };
  const gradFreeId = 'gradFree' + idSuffix;
  const gradTaxableId = 'gradTaxable' + idSuffix;
  const gradDeferredId = 'gradDeferred' + idSuffix;
  makeGrad(gradFreeId, '#1a8fc4', '#1578a8');
  makeGrad(gradTaxableId, '#2b5ea7', '#1e4a8a');
  makeGrad(gradDeferredId, '#0d6efd', '#0a58ca');

  svg.appendChild(defs);

  const topCx = 160, topCy = 80;
  const blCx = 82, blCy = 210;
  const brCx = 238, brCy = 210;
  const cR = 62;

  const lines = [
    [topCx, topCy + cR, blCx + cR * 0.65, blCy - cR * 0.65],
    [blCx + cR, blCy, brCx - cR, brCy],
    [brCx - cR * 0.65, brCy - cR * 0.65, topCx, topCy + cR],
  ];
  lines.forEach(([x1, y1, x2, y2]) => {
    const line = document.createElementNS(svgNS, 'line');
    line.setAttribute('x1', x1); line.setAttribute('y1', y1);
    line.setAttribute('x2', x2); line.setAttribute('y2', y2);
    line.setAttribute('stroke', '#c0cdd8');
    line.setAttribute('stroke-width', '1.5');
    line.setAttribute('stroke-dasharray', '4,3');
    svg.appendChild(line);
  });

  const total = taxFree + taxable + taxDeferred;
  const pct = v => total > 0 ? Math.round((v / total) * 100) : 0;

  const addCircleNode = (cx, cy, gradId, label, value, percent) => {
    const shadow = document.createElementNS(svgNS, 'circle');
    shadow.setAttribute('cx', cx); shadow.setAttribute('cy', cy);
    shadow.setAttribute('r', cR);
    shadow.setAttribute('filter', `url(#${filterId})`);
    shadow.setAttribute('fill', 'white');
    svg.appendChild(shadow);

    const circle = document.createElementNS(svgNS, 'circle');
    circle.setAttribute('cx', cx); circle.setAttribute('cy', cy);
    circle.setAttribute('r', cR);
    circle.setAttribute('fill', `url(#${gradId})`);
    svg.appendChild(circle);

    const glow = document.createElementNS(svgNS, 'circle');
    glow.setAttribute('cx', cx); glow.setAttribute('cy', cy);
    glow.setAttribute('r', cR - 3);
    glow.setAttribute('fill', 'none');
    glow.setAttribute('stroke', 'rgba(255,255,255,0.2)');
    glow.setAttribute('stroke-width', '1');
    svg.appendChild(glow);

    const lbl = document.createElementNS(svgNS, 'text');
    lbl.setAttribute('x', cx); lbl.setAttribute('y', cy - 16);
    lbl.setAttribute('text-anchor', 'middle');
    lbl.setAttribute('class', 'tri-node-label');
    lbl.textContent = label;
    svg.appendChild(lbl);

    const val = document.createElementNS(svgNS, 'text');
    val.setAttribute('x', cx); val.setAttribute('y', cy + 4);
    val.setAttribute('text-anchor', 'middle');
    val.setAttribute('class', 'tri-node-value');
    val.textContent = fmt$(value);
    svg.appendChild(val);

    const pctText = document.createElementNS(svgNS, 'text');
    pctText.setAttribute('x', cx); pctText.setAttribute('y', cy + 24);
    pctText.setAttribute('text-anchor', 'middle');
    pctText.setAttribute('class', 'tri-node-pct');
    pctText.textContent = percent + '%';
    svg.appendChild(pctText);
  };

  addCircleNode(topCx, topCy, gradFreeId,     'TAX-FREE',     taxFree,     pct(taxFree));
  addCircleNode(blCx,  blCy,  gradTaxableId,  'TAXABLE',      taxable,     pct(taxable));
  addCircleNode(brCx,  brCy,  gradDeferredId, 'TAX-DEFERRED', taxDeferred, pct(taxDeferred));

  return svg;
}


/* ==========================================================================
   PRESENTATION MODE
   ========================================================================== */

function getPresData() {
  const form = document.getElementById('pq-form');
  const liveData = form ? collectForm(form) : {};
  const store = getStore();
  const saved = store.pqTemplate || {};
  const pq = deepMerge(deepClone(saved), liveData);
  pq.employment = pq.employment || {};
  pq.employment.client1 = pq.employment.client1 || {};
  pq.employment.client1.status = pqState.employmentClient1;
  if (pqState.hasSpouse) {
    pq.employment.client2 = pq.employment.client2 || {};
    pq.employment.client2.status = pqState.employmentClient2;
  }
  // Ensure goals text is current
  if (_goalsTA) {
    pq.goals = pq.goals || {};
    pq.goals.goals = _goalsTA.value || pq.goals.goals || '';
  }
  return pq;
}

function showPresentation() {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('is-active'));
  document.getElementById('screen-presentation').classList.add('is-active');
  document.body.classList.add('pres-mode');
  renderPresentation();
}

function hidePresentation() {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('is-active'));
  document.getElementById('screen-pq').classList.add('is-active');
  document.body.classList.remove('pres-mode');
}

/* ---------- PRES HELPERS ---------- */

function makePresSection(title, badge) {
  const sec = el('div', { className: 'pres-section' });
  const hdr = el('div', { className: 'pres-section-header' });
  hdr.appendChild(el('h2', { className: 'pres-section-title', textContent: title }));
  if (badge) hdr.appendChild(el('span', { className: 'pres-section-badge', textContent: badge }));
  sec.appendChild(hdr);
  const body = el('div', { className: 'pres-section-body' });
  sec.appendChild(body);
  return { el: sec, body };
}

function buildPresTable(columns, dataRows, footTotals) {
  const wrap = el('div', { className: 'pres-table-wrap' });
  const table = el('table', { className: 'pres-table' });

  // Head
  const thead = el('thead');
  const headRow = el('tr');
  columns.forEach(col => headRow.appendChild(el('th', {
    textContent: col.label,
    className: col.right ? 'th-right' : '',
  })));
  thead.appendChild(headRow);
  table.appendChild(thead);

  // Body
  const tbody = el('tbody');
  dataRows.forEach((row, i) => {
    const tr = el('tr', { className: i % 2 !== 0 ? 'pres-row-alt' : '' });
    columns.forEach(col => {
      const td = el('td', { className: col.right ? 'td-right td-mono' : '' });
      if (col.pill) {
        const span = el('span', { className: `pres-pill ${row[col.pill] || ''}`, textContent: row[col.key] || '—' });
        td.appendChild(span);
      } else {
        td.textContent = row[col.key] != null && row[col.key] !== '' ? row[col.key] : '—';
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  // Optional footer totals
  if (footTotals) {
    const tfoot = el('tfoot');
    const footRow = el('tr');
    columns.forEach((col, i) => {
      const td = el('td', { className: col.right ? 'td-right' : '' });
      if (footTotals[col.key] != null) td.textContent = footTotals[col.key];
      else if (i === 0) td.textContent = 'Total';
      footRow.appendChild(td);
    });
    tfoot.appendChild(footRow);
    table.appendChild(tfoot);
  }

  wrap.appendChild(table);
  return wrap;
}

function presRow(label, value, opts = {}) {
  const row = el('div', { className: 'pres-bs-row' + (opts.total ? ' is-total' : '') + (opts.indent ? ' pres-bs-indent' : '') });
  row.appendChild(el('span', { className: 'pres-bs-label', textContent: label }));
  const valEl = el('span', { className: 'pres-bs-value' + (opts.cls ? ' ' + opts.cls : ''), textContent: value });
  row.appendChild(valEl);
  return row;
}

/* ---------- MAIN RENDER ---------- */

function renderPresentation() {
  const pq = getPresData();
  const host = document.getElementById('pres-content');
  clearChildren(host);

  const family    = pq.family        || {};
  const contact   = pq.contact       || {};
  const emp       = pq.employment    || {};
  const assets    = pq.assets        || {};
  const liabs     = pq.liabilities   || {};
  const ins       = pq.insurance     || {};
  const income    = pq.income        || {};
  const taxes     = pq.taxesExpenses || {};
  const rels      = pq.relationships || {};
  const goalsData = pq.goals         || {};

  const num  = v => parseCommas(String(v || '')) || 0;
  const sumA = (arr, key) => (arr || []).filter(Boolean).reduce((s, r) => s + num(r[key]), 0);
  const cleanRows = arr => (arr || []).filter(r => r && Object.values(r).some(v => v != null && String(v).trim() !== '' && !String(v).startsWith('_')));
  const dash = v => (v == null || v === '') ? '—' : v;

  // ── Names / header ──
  const c1First = (family.client1FirstName || '').trim();
  const c1Last  = (family.client1LastName || '').trim();
  const c2First = (family.client2FirstName || '').trim();
  const c2Last  = (family.client2LastName || '').trim();
  const c1Name = [c1First, c1Last].filter(Boolean).join(' ') || 'Client';
  const c2Name = [c2First, c2Last].filter(Boolean).join(' ');
  const clientDisplay = c2Name ? `${c1Name} & ${c2Name}` : c1Name;
  const age1 = family.client1Age ? ` (age ${family.client1Age})` : '';
  const age2 = family.client2Age ? ` (age ${family.client2Age})` : '';
  const clientAges = c2Name ? `${c1Name}${age1} & ${c2Name}${age2}` : `${c1Name}${age1}`;
  const cityState = [contact.city, contact.state].filter(Boolean).join(', ');

  // ── Financial totals ──
  const reAssets    = sumA(assets.realEstate,    'marketValue');
  const busOther    = sumA(assets.businessOther, 'marketValue');
  const taxDef      = sumA(assets.taxDeferred,   'marketValue');
  const rothVal     = sumA(assets.roth,          'marketValue');
  const taxable     = sumA(assets.taxable,       'marketValue');
  const cashCd      = sumA(assets.cashCd,        'marketValue');
  const plan529     = sumA(assets.plan529,       'marketValue');
  const hsaVal      = sumA(assets.hsa,           'marketValue');
  const investable  = taxDef + rothVal + taxable + cashCd + plan529 + hsaVal;
  const totalAssets = investable + reAssets + busOther;
  const totalLiab   = sumA(liabs.items, 'amount');
  const netWorth    = totalAssets - totalLiab;

  const empInc   = sumA(income.employment,     'annualAmount');
  const ssInc    = sumA(income.socialSecurity, 'annualAmount');
  const pensInc  = sumA(income.pension,        'annualAmount');
  const otherInc = sumA(income.other,          'annualAmount');
  const totalIncome = empInc + ssInc + pensInc + otherInc;

  const fedTax   = num(taxes.federalTax);
  const stTax    = num(taxes.stateTax);
  const ficaTax  = num(taxes.ficaTax);
  const totalTax = fedTax + stTax + ficaTax;

  const totalSavings = sumA(assets.taxDeferred, 'personalAdditions')
    + sumA(assets.roth,    'personalAdditions')
    + sumA(assets.taxable, 'personalAdditions');

  const debtPayments = (liabs.items || []).filter(Boolean)
    .reduce((s, r) => s + num(r.payment) * 12, 0);

  const expRows = (taxes.expenses || []).filter(Boolean);
  const livingExpField = num(taxes.livingExpenses);
  const livingExp = livingExpField + expRows.filter(r => r._source !== 'liability' && r._source !== 'insurance')
    .reduce((s, r) => s + num(r.amount), 0);
  const totalExpenses = livingExpField + expRows.reduce((s, r) => s + num(r.amount), 0);

  const netCF = totalIncome - totalTax - totalSavings - debtPayments - livingExp;
  const savingsRatio = totalIncome > 0 ? (totalSavings / totalIncome) * 100 : 0;
  const dtiRatio     = totalIncome > 0 ? (debtPayments / totalIncome) * 100 : 0;

  // Helper: info grid
  const buildInfoGrid = (items) => {
    const grid = el('div', { className: 'pres-info-grid' });
    items.forEach(it => {
      if (it.hidden) return;
      const cell = el('div', { className: 'pres-info-cell' + (it.span ? ` span-${it.span}` : '') });
      cell.appendChild(el('div', { className: 'pres-info-label', textContent: it.label }));
      cell.appendChild(el('div', { className: 'pres-info-value', textContent: dash(it.value) }));
      grid.appendChild(cell);
    });
    return grid;
  };

  // ══════════════════════════════════════════════════
  // 1. CLIENT HEADER
  // ══════════════════════════════════════════════════
  const header = el('div', { className: 'pres-client-header' });
  const nameBlock = el('div');
  nameBlock.appendChild(el('div', { className: 'pres-client-name', textContent: clientDisplay }));
  nameBlock.appendChild(el('div', { className: 'pres-client-sub', textContent: clientAges + (cityState ? '  ·  ' + cityState : '') }));
  const prepBlock = el('div', { className: 'pres-prepared' });
  prepBlock.appendChild(el('strong', { textContent: 'Pure Financial Advisors' }));
  prepBlock.appendChild(document.createTextNode('Prepared: ' + new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })));
  header.append(nameBlock, prepBlock);
  host.appendChild(header);

  // ══════════════════════════════════════════════════
  // 2. KPI CARDS
  // ══════════════════════════════════════════════════
  const kpiRow = el('div', { className: 'pres-kpi-row' });
  [
    { label: 'Net Worth',         value: fmt$(netWorth),    color: '#1f3f67', cls: netWorth >= 0 ? 'positive' : 'negative' },
    { label: 'Total Assets',      value: fmt$(totalAssets), color: '#4f46e5', cls: '' },
    { label: 'Total Liabilities', value: fmt$(totalLiab),   color: '#e11d48', cls: '' },
    { label: 'Annual Income',     value: fmt$(totalIncome), color: '#059669', cls: '' },
    { label: 'Net Cash Flow',     value: fmt$(netCF),       color: netCF >= 0 ? '#059669' : '#e11d48', cls: netCF >= 0 ? 'positive' : 'negative' },
  ].forEach(k => {
    const card = el('div', { className: 'pres-kpi-card' });
    card.style.setProperty('--kpi-color', k.color);
    card.appendChild(el('div', { className: 'pres-kpi-label', textContent: k.label }));
    card.appendChild(el('div', { className: `pres-kpi-value ${k.cls}`, textContent: k.value }));
    kpiRow.appendChild(card);
  });
  host.appendChild(kpiRow);

  // ══════════════════════════════════════════════════
  // 3. FAMILY INFORMATION
  // ══════════════════════════════════════════════════
  const famSec = makePresSection('Family Information');

  const hasSpouseRow = c2Name || c2First || c2Last || family.client2Nickname || family.client2DOB || family.client2Age;
  const hasAddrRow = contact.address1 || contact.address2 || contact.city || contact.state || contact.zip;

  const ssGrid = el('div', { className: 'pres-ss-grid pres-ss-family' });
  const mkCell = (label, value, opts = {}) => {
    const cell = el('div', { className: 'pres-ss-cell' + (opts.span ? ` span-${opts.span}` : '') + (opts.empty ? ' is-empty' : '') });
    if (!opts.empty) {
      cell.appendChild(el('div', { className: 'pres-ss-label', textContent: label }));
      cell.appendChild(el('div', { className: 'pres-ss-value', textContent: dash(value) }));
    }
    ssGrid.appendChild(cell);
  };

  // Row 1: Client — matches spreadsheet layout
  mkCell('Your Name',  c1First);
  mkCell('Last Name',  c1Last);
  mkCell('Nickname',   family.client1Nickname);
  mkCell('Birth Date', family.client1DOB);
  mkCell('Age',        family.client1Age);

  // Row 2: Spouse
  if (hasSpouseRow) {
    mkCell("Spouse's Name", c2First);
    mkCell('Last Name',     c2Last);
    mkCell('Nickname',      family.client2Nickname);
    mkCell('Birth Date',    family.client2DOB);
    mkCell('Age',           family.client2Age);
  }

  // Row 3: Residence Address (spans cols 1–2), City, State, Zip
  if (hasAddrRow) {
    mkCell('Residence Address', [contact.address1, contact.address2].filter(Boolean).join(', '), { span: 2 });
    mkCell('City',     contact.city);
    mkCell('State',    contact.state);
    mkCell('Zip Code', contact.zip);
  }

  // Referred By as a trailing full-width cell
  if (contact.referredBy) {
    mkCell('Referred By', contact.referredBy, { span: 5 });
  }

  const childRows = (family.children || []).filter(c => c && (c.name || c.age));
  const grandRows = (family.grandchildren || []).filter(c => c && (c.name || c.age));
  const numChildren = parseInt(family.numChildren, 10) || 0;
  const numGrand = parseInt(family.numGrandchildren, 10) || 0;
  const showChildren = childRows.length > 0 || numChildren > 0;
  const showGrand = grandRows.length > 0 || numGrand > 0;
  const hasKids = showChildren || showGrand;

  if (hasKids) {
    const famLayout = el('div', { className: 'pres-family-layout' });
    famLayout.appendChild(ssGrid);
    const kidsPanel = el('div', { className: 'pres-kids-panel' });
    kidsPanel.classList.add(showChildren && showGrand ? 'pres-kids-panel--split' : 'pres-kids-panel--single');

    const buildKidGroup = (title, countLabel, rows) => {
      const group = el('div', { className: 'pres-kids-group' });
      const heading = el('div', { className: 'pres-kids-heading' });
      heading.appendChild(el('span', { className: 'pres-kids-title', textContent: title }));
      if (countLabel) heading.appendChild(el('span', { className: 'pres-kids-count', textContent: countLabel }));
      group.appendChild(heading);
      if (rows.length) {
        const list = el('ul', { className: 'pres-kids-list' });
        rows.forEach(r => {
          const li = el('li', { className: 'pres-kids-item' });
          li.appendChild(el('span', { className: 'pres-kids-name', textContent: dash(r.name) }));
          const ageVal = (r.age || r.age === 0) ? String(r.age) : '';
          li.appendChild(el('span', { className: 'pres-kids-age', textContent: ageVal || '—' }));
          list.appendChild(li);
        });
        group.appendChild(list);
      } else {
        group.appendChild(el('div', { className: 'pres-kids-empty', textContent: 'No names recorded' }));
      }
      return group;
    };

    if (showChildren) {
      const count = numChildren || childRows.length;
      kidsPanel.appendChild(buildKidGroup('Children', count ? String(count) : '', childRows));
    }
    if (showGrand) {
      const count = numGrand || grandRows.length;
      kidsPanel.appendChild(buildKidGroup('Grandchildren', count ? String(count) : '', grandRows));
    }
    famLayout.appendChild(kidsPanel);
    famSec.body.appendChild(famLayout);
  } else {
    famSec.body.appendChild(ssGrid);
  }
  if (family.additionalFamilyInfo) {
    famSec.body.appendChild(el('div', { className: 'pres-subsec-label', textContent: 'Additional Family Information' }));
    famSec.body.appendChild(el('p', { className: 'pres-notes-text', textContent: family.additionalFamilyInfo }));
  }
  host.appendChild(famSec.el);

  // ══════════════════════════════════════════════════
  // 4. OCCUPATION
  // ══════════════════════════════════════════════════
  const c1Emp = emp.client1 || {};
  const c2Emp = emp.client2 || {};
  const hasEmp = c1Emp.status || c2Emp.status || c1Emp.jobTitle || c2Emp.jobTitle || c1Emp.employer || c2Emp.employer || c1Emp.lastJob || c2Emp.lastJob;
  if (hasEmp) {
    const occSec = makePresSection('Occupation');
    occSec.el.classList.add('is-compact');
    const mkEmpBlock = (who, e) => {
      occSec.body.appendChild(el('div', { className: 'pres-subsec-label', textContent: who + (e.status ? ' · ' + e.status : '') }));
      const isWorking = e.status === 'Employed';
      occSec.body.appendChild(buildInfoGrid([
        { label: 'Job Title',                              value: isWorking ? e.jobTitle : e.lastJob },
        { label: 'Employer' + (isWorking ? '' : ' (Last)'), value: isWorking ? e.employer : e.lastEmployer },
        { label: '# of Years',                             value: isWorking ? e.years : e.yearsWorked },
        { label: 'Retirement Date',                        value: e.retirementDate,  hidden: !isWorking },
        { label: 'Retirement Age',                         value: e.retirementAge,   hidden: !isWorking },
        { label: 'Type of Business',                       value: e.businessType,    hidden: !isWorking },
      ]));
    };
    mkEmpBlock(c1Name, c1Emp);
    if (c2Name || c2Emp.status) mkEmpBlock(c2Name || 'Spouse', c2Emp);
    host.appendChild(occSec.el);
  }

  // ══════════════════════════════════════════════════
  // 5. CURRENT PROFESSIONAL RELATIONSHIPS
  // ══════════════════════════════════════════════════
  const relRows = cleanRows(rels.professionals);
  if (relRows.length) {
    const relSec = makePresSection('Current Professional Relationships');
    relSec.el.classList.add('is-compact');
    relSec.body.appendChild(buildPresTable([
      { label: 'Role',  key: 'role' },
      { label: 'Name',  key: 'name' },
      { label: 'Firm',  key: 'firmName' },
      { label: 'City',  key: 'city' },
      { label: 'State', key: 'state' },
    ], relRows));
    host.appendChild(relSec.el);
  }

  // ══════════════════════════════════════════════════
  // 6. GOALS
  // ══════════════════════════════════════════════════
  const goalsText = (goalsData.goals || '').trim();
  if (goalsText) {
    const gSec = makePresSection('Goals');
    gSec.el.classList.add('is-compact');
    gSec.body.appendChild(el('p', { className: 'pres-goals-text', textContent: goalsText }));
    host.appendChild(gSec.el);
  }

  // ══════════════════════════════════════════════════
  // 7. REAL ESTATE ASSETS
  // ══════════════════════════════════════════════════
  const reRows = cleanRows(assets.realEstate);
  if (reRows.length) {
    const reSec = makePresSection('Real Estate Assets', fmt$(reAssets));
    const reLoanTot = sumA(reRows, 'remainingLoan');
    const rePmtTot  = sumA(reRows, 'payment');
    const reEbtTot  = sumA(reRows, 'incomeEBT');
    const reCbTot   = sumA(reRows, 'costBasis');
    reSec.body.appendChild(buildPresTable([
      { label: 'Description',  key: 'description' },
      { label: 'Market Value', key: '_mv',   right: true },
      { label: 'Loan Amount',  key: '_loan', right: true },
      { label: 'Int. Rate',    key: '_rate', right: true },
      { label: 'Term',         key: 'term',  right: true },
      { label: 'Pmt (P&I)',    key: '_pmt',  right: true },
      { label: 'Income EBT',   key: '_ebt',  right: true },
      { label: 'Year Acq.',    key: 'yearAcquired' },
      { label: 'Cost Basis',   key: '_cb',   right: true },
      { label: 'Ownership',    key: 'ownership' },
    ], reRows.map(r => ({
      ...r,
      _mv:   fmt$(num(r.marketValue)),
      _loan: num(r.remainingLoan) ? fmt$(num(r.remainingLoan)) : '—',
      _rate: r.interestRate ? r.interestRate + '%' : '—',
      _pmt:  num(r.payment) ? fmt$(num(r.payment)) : '—',
      _ebt:  num(r.incomeEBT) ? fmt$(num(r.incomeEBT)) : '—',
      _cb:   num(r.costBasis) ? fmt$(num(r.costBasis)) : '—',
    })), {
      _mv:   fmt$(reAssets),
      _loan: reLoanTot ? fmt$(reLoanTot) : '—',
      _pmt:  rePmtTot  ? fmt$(rePmtTot)  : '—',
      _ebt:  reEbtTot  ? fmt$(reEbtTot)  : '—',
      _cb:   reCbTot   ? fmt$(reCbTot)   : '—',
    }));
    host.appendChild(reSec.el);
  }

  // ══════════════════════════════════════════════════
  // 8. INVESTMENT ACCOUNTS + TAX TREATMENT (consolidated)
  // ══════════════════════════════════════════════════
  const tdRows   = cleanRows(assets.taxDeferred);
  const rothRows = cleanRows(assets.roth);
  const txRows   = cleanRows(assets.taxable);
  const taxFreeTotal     = rothVal + hsaVal;
  const taxableTotal     = taxable + cashCd;
  const taxDeferredTotal = taxDef;
  const triTotal = taxFreeTotal + taxableTotal + taxDeferredTotal;

  if (tdRows.length || rothRows.length || txRows.length) {
    const invTotal = taxDef + rothVal + taxable;
    const invSec = makePresSection('Investment Accounts', fmt$(invTotal));
    const layout = el('div', { className: 'pres-invest-layout' });
    const accountsCol = el('div', { className: 'pres-invest-accounts' });

    const mkGroupTitle = (label, amount, variant) => {
      const h = el('div', { className: `pres-invest-group-title ${variant}` });
      h.appendChild(el('span', { className: 'pres-invest-group-label', textContent: label }));
      h.appendChild(el('span', { className: 'pres-invest-group-amount', textContent: fmt$(amount) }));
      return h;
    };

    if (tdRows.length) {
      accountsCol.appendChild(mkGroupTitle('Tax-Deferred Retirement', taxDef, 'td'));
      const tdAddTot   = sumA(tdRows, 'personalAdditions');
      const tdMatchTot = sumA(tdRows, 'companyMatch');
      accountsCol.appendChild(buildPresTable([
        { label: 'Custodian',          key: 'custodian' },
        { label: 'Market Value',       key: '_mv',    right: true },
        { label: 'Personal Additions', key: '_add',   right: true },
        { label: 'Company Match',      key: '_match', right: true },
        { label: 'Type',               key: 'type' },
        { label: 'Ownership',          key: 'ownership' },
        { label: 'Beneficiary',        key: 'beneficiary' },
      ], tdRows.map(r => ({
        ...r,
        _mv:    fmt$(num(r.marketValue)),
        _add:   num(r.personalAdditions) ? fmt$(num(r.personalAdditions)) : '—',
        _match: num(r.companyMatch) ? fmt$(num(r.companyMatch)) : '—',
      })), {
        _mv:    fmt$(taxDef),
        _add:   tdAddTot   ? fmt$(tdAddTot)   : '—',
        _match: tdMatchTot ? fmt$(tdMatchTot) : '—',
      }));
    }

    if (rothRows.length) {
      accountsCol.appendChild(mkGroupTitle('Tax-Free Roth', rothVal, 'roth'));
      const rothAddTot   = sumA(rothRows, 'personalAdditions');
      const rothMatchTot = sumA(rothRows, 'companyMatch');
      accountsCol.appendChild(buildPresTable([
        { label: 'Custodian',          key: 'custodian' },
        { label: 'Market Value',       key: '_mv',    right: true },
        { label: 'Personal Additions', key: '_add',   right: true },
        { label: 'Company Match',      key: '_match', right: true },
        { label: 'Roth Type',          key: 'type' },
        { label: 'Ownership',          key: 'ownership' },
        { label: 'Beneficiary',        key: 'beneficiary' },
      ], rothRows.map(r => ({
        ...r,
        _mv:    fmt$(num(r.marketValue)),
        _add:   num(r.personalAdditions) ? fmt$(num(r.personalAdditions)) : '—',
        _match: num(r.companyMatch) ? fmt$(num(r.companyMatch)) : '—',
      })), {
        _mv:    fmt$(rothVal),
        _add:   rothAddTot   ? fmt$(rothAddTot)   : '—',
        _match: rothMatchTot ? fmt$(rothMatchTot) : '—',
      }));
    }

    if (txRows.length) {
      accountsCol.appendChild(mkGroupTitle('Taxable Non-Retirement', taxable, 'taxable'));
      const txAddTot = sumA(txRows, 'personalAdditions');
      const txCbTot  = sumA(txRows, 'costBasis');
      accountsCol.appendChild(buildPresTable([
        { label: 'Custodian',    key: 'custodian' },
        { label: 'Market Value', key: '_mv',  right: true },
        { label: 'Additions',    key: '_add', right: true },
        { label: 'Cost Basis',   key: '_cb',  right: true },
        { label: 'Ownership',    key: 'ownership' },
        { label: 'Var/Fixed',    key: 'variableFixed' },
        { label: 'Issue Date',   key: 'issueDate' },
        { label: 'Beneficiary',  key: 'beneficiary' },
      ], txRows.map(r => ({
        ...r,
        _mv:  fmt$(num(r.marketValue)),
        _add: num(r.personalAdditions) ? fmt$(num(r.personalAdditions)) : '—',
        _cb:  num(r.costBasis) ? fmt$(num(r.costBasis)) : '—',
      })), {
        _mv:  fmt$(taxable),
        _add: txAddTot ? fmt$(txAddTot) : '—',
        _cb:  txCbTot  ? fmt$(txCbTot)  : '—',
      }));
    }

    layout.appendChild(accountsCol);

    if (triTotal) {
      const triCol = el('div', { className: 'pres-invest-triangle' });
      triCol.appendChild(el('div', { className: 'pres-invest-tri-title', textContent: 'Tax Treatment Mix' }));
      const triChart = el('div', { className: 'pres-tri-chart' });
      triChart.appendChild(buildTaxTriangleSVG(taxFreeTotal, taxableTotal, taxDeferredTotal, { idSuffix: 'Pres' }));
      triCol.appendChild(triChart);
      triCol.appendChild(el('div', { className: 'pres-invest-tri-footnote', textContent: 'Includes HSA & cash in their tax categories.' }));
      layout.appendChild(triCol);
    }

    invSec.body.appendChild(layout);
    host.appendChild(invSec.el);
  }

  // ══════════════════════════════════════════════════
  // 12. CASH & CDs  |  BUSINESS & OTHER  (side-by-side)
  // ══════════════════════════════════════════════════
  const cashRows   = cleanRows(assets.cashCd);
  const plan529Rs  = cleanRows(assets.plan529);
  const hsaRows    = cleanRows(assets.hsa);
  const busRows    = cleanRows(assets.businessOther);
  if (cashRows.length || plan529Rs.length || hsaRows.length || busRows.length) {
    const sec = makePresSection("Cash, CD's, Business & Other", fmt$(cashCd + plan529 + hsaVal + busOther));
    const split = el('div', { className: 'pres-split-col' });

    if (cashRows.length || plan529Rs.length || hsaRows.length) {
      const leftCol = el('div');
      if (cashRows.length) {
        leftCol.appendChild(el('div', { className: 'pres-col-title', textContent: `Cash & CD's · ${fmt$(cashCd)}` }));
        leftCol.appendChild(buildPresTable([
          { label: 'Description',  key: 'description' },
          { label: 'Market Value', key: '_mv', right: true },
          { label: 'Int. Rate',    key: '_rate' },
          { label: 'Owner',        key: 'ownership' },
        ], cashRows.map(r => ({ ...r, _mv: fmt$(num(r.marketValue)), _rate: r.interestRate ? r.interestRate + '%' : '—' })),
          { _mv: fmt$(cashCd) }));
      }
      if (plan529Rs.length) {
        leftCol.appendChild(el('div', { className: 'pres-col-title pres-col-title-sub', textContent: `529 Plans · ${fmt$(plan529)}` }));
        leftCol.appendChild(buildPresTable([
          { label: 'Description',  key: 'description' },
          { label: 'Market Value', key: '_mv', right: true },
          { label: 'Owner',        key: 'ownership' },
        ], plan529Rs.map(r => ({ ...r, _mv: fmt$(num(r.marketValue)) })), { _mv: fmt$(plan529) }));
      }
      if (hsaRows.length) {
        leftCol.appendChild(el('div', { className: 'pres-col-title pres-col-title-sub', textContent: `HSA · ${fmt$(hsaVal)}` }));
        leftCol.appendChild(buildPresTable([
          { label: 'Description',  key: 'description' },
          { label: 'Market Value', key: '_mv', right: true },
          { label: 'Owner',        key: 'ownership' },
        ], hsaRows.map(r => ({ ...r, _mv: fmt$(num(r.marketValue)) })), { _mv: fmt$(hsaVal) }));
      }
      split.appendChild(leftCol);
    }

    if (busRows.length) {
      const rightCol = el('div');
      rightCol.appendChild(el('div', { className: 'pres-col-title', textContent: `Business & Other · ${fmt$(busOther)}` }));
      const busCbTot = sumA(busRows, 'costBasis');
      rightCol.appendChild(buildPresTable([
        { label: 'Description',  key: 'description' },
        { label: 'Market Value', key: '_mv', right: true },
        { label: 'Cost Basis',   key: '_cb', right: true },
        { label: 'Owner',        key: 'ownership' },
      ], busRows.map(r => ({ ...r, _mv: fmt$(num(r.marketValue)), _cb: num(r.costBasis) ? fmt$(num(r.costBasis)) : '—' })),
        { _mv: fmt$(busOther), _cb: busCbTot ? fmt$(busCbTot) : '—' }));
      split.appendChild(rightCol);
    }

    sec.body.appendChild(split);
    host.appendChild(sec.el);
  }

  // ══════════════════════════════════════════════════
  // 13. OTHER LIABILITIES  +  TOTAL NET WORTH
  // ══════════════════════════════════════════════════
  const liabDetail = cleanRows(liabs.items);
  const liabSec = makePresSection('Other Liabilities', fmt$(totalLiab));
  if (liabDetail.length) {
    const liabPmtTot = sumA(liabDetail, 'payment');
    liabSec.body.appendChild(buildPresTable([
      { label: 'Description',  key: 'description' },
      { label: 'Payment (mo)', key: '_pmt',  right: true },
      { label: 'Amount',       key: '_amt',  right: true },
      { label: 'Int. Rate',    key: '_rate', right: true },
      { label: 'Term',         key: 'term' },
    ], liabDetail.map(r => ({
      ...r,
      _pmt:  num(r.payment) ? fmt$(num(r.payment)) : '—',
      _amt:  fmt$(num(r.amount)),
      _rate: r.interestRate ? r.interestRate + '%' : '—',
    })), {
      _pmt: liabPmtTot ? fmt$(liabPmtTot) : '—',
      _amt: fmt$(totalLiab),
    }));
  } else {
    liabSec.body.appendChild(el('p', { className: 'pres-empty-note', textContent: 'No liabilities reported.' }));
  }

  const nwFooter = el('div', { className: 'pres-networth-footer' });
  const nwLeft = el('div', { className: 'pres-networth-col' });
  nwLeft.appendChild(presRow('Total Assets',      fmt$(totalAssets)));
  nwLeft.appendChild(presRow('Total Liabilities', fmt$(totalLiab)));
  const nwRight = el('div', { className: 'pres-networth-total' });
  nwRight.appendChild(el('span', { className: 'pres-networth-label', textContent: 'Total Net Worth' }));
  nwRight.appendChild(el('span', { className: 'pres-networth-value ' + (netWorth >= 0 ? 'positive' : 'negative'), textContent: fmt$(netWorth) }));
  nwFooter.append(nwLeft, nwRight);
  liabSec.body.appendChild(nwFooter);
  host.appendChild(liabSec.el);

  // ══════════════════════════════════════════════════
  // 14. INSURANCE INFORMATION
  // ══════════════════════════════════════════════════
  const insRows = cleanRows(ins.policies);
  if (insRows.length) {
    const totalPrem    = sumA(ins.policies, 'annualPremium');
    const totalCV      = sumA(ins.policies, 'cashValue');
    const totalBenefit = sumA(ins.policies, 'benefit');
    const insSec = makePresSection('Insurance Information');
    insSec.body.appendChild(buildPresTable([
      { label: 'Company',             key: 'company' },
      { label: 'Type',                key: 'type' },
      { label: 'Death/Daily Benefit', key: '_benefit', right: true },
      { label: 'Insured',             key: 'insured' },
      { label: 'Owner',               key: 'owner' },
      { label: 'Policy Date',         key: 'policyDate' },
      { label: 'Annual Premium',      key: '_prem', right: true },
      { label: 'Cash Value',          key: '_cv',   right: true },
      { label: 'Beneficiary',         key: 'beneficiary' },
    ], insRows.map(r => ({
      ...r,
      _benefit: num(r.benefit)       ? fmt$(num(r.benefit))       : '—',
      _cv:      num(r.cashValue)     ? fmt$(num(r.cashValue))     : '—',
      _prem:    num(r.annualPremium) ? fmt$(num(r.annualPremium)) : '—',
    })), {
      _benefit: totalBenefit ? fmt$(totalBenefit) : '—',
      _prem:    fmt$(totalPrem),
      _cv:      fmt$(totalCV),
    }));
    host.appendChild(insSec.el);
  }

  // ══════════════════════════════════════════════════
  // 15. INCOME & TAX INFORMATION
  // ══════════════════════════════════════════════════
  const empRowsIn   = cleanRows(income.employment);
  const ssRowsIn    = cleanRows(income.socialSecurity);
  const pensRowsIn  = cleanRows(income.pension);
  const otherRowsIn = cleanRows(income.other);
  const hasIncome = empRowsIn.length || ssRowsIn.length || pensRowsIn.length || otherRowsIn.length;
  const hasTaxSummary = num(taxes.capLossCarryForward) || num(taxes.taxableIncome) || num(taxes.standardItemDeduction) || totalTax;

  if (hasIncome || hasTaxSummary) {
    const incSec = makePresSection('Income & Tax Information', fmt$(totalIncome));
    const incLayout = el('div', { className: 'pres-income-layout' });

    const incMain = el('div', { className: 'pres-income-main' });

    if (empRowsIn.length) {
      incMain.appendChild(el('div', { className: 'pres-col-title', textContent: `Employment Income · ${fmt$(empInc)}` }));
      incMain.appendChild(buildPresTable([
        { label: 'Owner',           key: 'owner' },
        { label: 'Description',     key: 'description' },
        { label: 'Annual Amount',   key: '_amt',  right: true },
        { label: 'Retirement Date', key: 'retirementDate' },
        { label: 'COLA %',          key: '_cola', right: true },
      ], empRowsIn.map(r => ({ ...r, _amt: fmt$(num(r.annualAmount)), _cola: r.cola ? r.cola + '%' : '—' })),
        { _amt: fmt$(empInc) }));
    }

    if (ssRowsIn.length) {
      incMain.appendChild(el('div', { className: 'pres-col-title pres-col-title-sub', textContent: `Social Security · ${fmt$(ssInc)}` }));
      incMain.appendChild(buildPresTable([
        { label: 'Owner',         key: 'owner' },
        { label: 'Starting Age',  key: 'startingAge', right: true },
        { label: 'Annual Amount', key: '_amt',        right: true },
        { label: 'Full Ret. Age', key: 'fullRetAge',  right: true },
        { label: 'COLA %',        key: '_cola',       right: true },
      ], ssRowsIn.map(r => ({ ...r, _amt: fmt$(num(r.annualAmount)), _cola: r.cola ? r.cola + '%' : '—' })),
        { _amt: fmt$(ssInc) }));
    }

    if (pensRowsIn.length) {
      incMain.appendChild(el('div', { className: 'pres-col-title pres-col-title-sub', textContent: `Pension · ${fmt$(pensInc)}` }));
      incMain.appendChild(buildPresTable([
        { label: 'Owner',            key: 'owner' },
        { label: 'Start Date',       key: 'startDate' },
        { label: 'Annual Amount',    key: '_amt',  right: true },
        { label: 'Survivor Benefit', key: '_surv', right: true },
        { label: 'COLA %',           key: '_cola', right: true },
      ], pensRowsIn.map(r => ({ ...r, _amt: fmt$(num(r.annualAmount)), _surv: r.survivorBenefit ? r.survivorBenefit + '%' : '—', _cola: r.cola ? r.cola + '%' : '—' })),
        { _amt: fmt$(pensInc) }));
    }

    if (otherRowsIn.length) {
      incMain.appendChild(el('div', { className: 'pres-col-title pres-col-title-sub', textContent: `Other Income · ${fmt$(otherInc)}` }));
      incMain.appendChild(buildPresTable([
        { label: 'Owner',         key: 'owner' },
        { label: 'Description',   key: 'description' },
        { label: 'Annual Amount', key: '_amt', right: true },
        { label: 'Start',         key: 'startDate' },
        { label: 'End',           key: 'endDate' },
        { label: 'COLA %',        key: '_cola', right: true },
      ], otherRowsIn.map(r => ({ ...r, _amt: fmt$(num(r.annualAmount)), _cola: r.cola ? r.cola + '%' : '—' })),
        { _amt: fmt$(otherInc) }));
    }

    if (hasIncome) {
      const totalIncRow = el('div', { className: 'pres-income-total' });
      totalIncRow.appendChild(el('span', { textContent: 'Current Total' }));
      totalIncRow.appendChild(el('span', { textContent: fmt$(totalIncome) }));
      incMain.appendChild(totalIncRow);
    }

    incLayout.appendChild(incMain);

    if (hasTaxSummary) {
      const taxPanel = el('aside', { className: 'pres-tax-panel' });
      taxPanel.appendChild(el('div', { className: 'pres-tax-panel-title', textContent: 'Prior Year Tax Summary' }));
      const mkTaxRow = (label, value, emph) => {
        const row = el('div', { className: 'pres-tax-row' + (emph ? ' is-total' : '') });
        row.appendChild(el('span', { textContent: label }));
        row.appendChild(el('span', { textContent: value }));
        return row;
      };
      if (num(taxes.capLossCarryForward))   taxPanel.appendChild(mkTaxRow('Cap Loss Carry Forward', fmt$(num(taxes.capLossCarryForward))));
      if (num(taxes.taxableIncome))         taxPanel.appendChild(mkTaxRow('Taxable Income',         fmt$(num(taxes.taxableIncome))));
      if (num(taxes.standardItemDeduction)) taxPanel.appendChild(mkTaxRow('Stnd/Item Deduction',    fmt$(num(taxes.standardItemDeduction))));
      if (totalTax) {
        taxPanel.appendChild(el('div', { className: 'pres-tax-subhead', textContent: 'Taxes Paid' }));
        if (fedTax)  taxPanel.appendChild(mkTaxRow('Federal Tax', fmt$(fedTax)));
        if (stTax)   taxPanel.appendChild(mkTaxRow('State Tax',   fmt$(stTax)));
        if (ficaTax) taxPanel.appendChild(mkTaxRow('FICA Tax',    fmt$(ficaTax)));
        taxPanel.appendChild(mkTaxRow('Total Taxes Paid', fmt$(totalTax), true));
        if (totalIncome > 0) taxPanel.appendChild(mkTaxRow('Effective Rate', ((totalTax / totalIncome) * 100).toFixed(1) + '%'));
      }
      incLayout.appendChild(taxPanel);
    }

    incSec.body.appendChild(incLayout);
    host.appendChild(incSec.el);
  }

  // ══════════════════════════════════════════════════
  // 16. EXPENSES & CASH FLOW
  // ══════════════════════════════════════════════════
  const hasExpenses = livingExpField || expRows.length;
  if (hasExpenses || totalIncome) {
    const expSec = makePresSection('Expenses & Cash Flow');

    if (livingExpField) {
      expSec.body.appendChild(el('div', { className: 'pres-subsec-label', textContent: 'Living Expenses' }));
      expSec.body.appendChild(buildInfoGrid([
        { label: 'Annual Amount', value: fmt$(livingExpField) },
      ]));
    }

    if (expRows.length) {
      expSec.body.appendChild(el('div', { className: 'pres-subsec-label', textContent: 'Other Expenses & Personal Liabilities' }));
      expSec.body.appendChild(buildPresTable([
        { label: 'Description',   key: 'description' },
        { label: 'Annual Amount', key: '_amt',  right: true },
        { label: 'Start',         key: 'startDate' },
        { label: 'End',           key: 'endDate' },
        { label: 'COLA %',        key: '_cola', right: true },
        { label: 'Notes',         key: 'notes' },
      ], expRows.map(r => ({ ...r, _amt: fmt$(num(r.amount)), _cola: r.cola ? r.cola + '%' : '—' })),
        { _amt: fmt$(expRows.reduce((s, r) => s + num(r.amount), 0)) }));
    }

    expSec.body.appendChild(el('div', { className: 'pres-subsec-label', textContent: 'Annual Cash Flow' }));
    const cfLayout = el('div', { className: 'pres-cf-layout' });

    const cfTable = el('table', { className: 'pres-cf-table' });
    const cfBody = el('tbody');
    const fedStateTax = totalTax - ficaTax;
    const cfLines = [
      { label: 'Gross Income',                  val: totalIncome,  minus: false },
      { label: 'Less: Federal & State Tax',     val: fedStateTax,  minus: true, skip: fedStateTax === 0 },
      { label: 'Less: FICA Tax',                val: ficaTax,      minus: true, skip: ficaTax === 0 },
      { label: 'Less: Savings & Contributions', val: totalSavings, minus: true, skip: totalSavings === 0 },
      { label: 'Less: Debt Payments',           val: debtPayments, minus: true, skip: debtPayments === 0 },
      { label: 'Less: Living Expenses',         val: livingExp,    minus: true, skip: livingExp === 0 },
    ];
    cfLines.forEach(line => {
      if (line.skip) return;
      const tr = el('tr');
      tr.appendChild(el('td', { className: 'cf-label', textContent: line.label }));
      const valTd = el('td', { className: 'cf-value' + (line.minus && line.val > 0 ? ' cf-negative' : '') });
      valTd.textContent = (line.minus && line.val > 0 ? '(' : '') + fmt$(line.val) + (line.minus && line.val > 0 ? ')' : '');
      tr.appendChild(valTd);
      cfBody.appendChild(tr);
    });
    cfTable.appendChild(cfBody);
    cfLayout.appendChild(cfTable);

    const netPanel = el('div', { className: 'pres-cf-net' + (netCF >= 0 ? ' is-positive' : ' is-negative') });
    netPanel.appendChild(el('div', { className: 'pres-cf-net-label', textContent: 'Net Cash Flow' }));
    netPanel.appendChild(el('div', { className: 'pres-cf-net-value', textContent: fmt$(netCF) }));
    netPanel.appendChild(el('div', { className: 'pres-cf-net-hint', textContent: netCF >= 0 ? 'Annual surplus' : 'Annual shortfall' }));

    if (totalIncome > 0) {
      const ratios = el('div', { className: 'pres-cf-ratios' });
      const mkRatio = (label, pct, hint) => {
        const r = el('div', { className: 'pres-ratio' });
        r.appendChild(el('div', { className: 'pres-ratio-label', textContent: label }));
        r.appendChild(el('div', { className: 'pres-ratio-value', textContent: pct.toFixed(1) + '%' }));
        if (hint) r.appendChild(el('div', { className: 'pres-ratio-hint', textContent: hint }));
        return r;
      };
      ratios.appendChild(mkRatio('Savings Ratio',        savingsRatio, 'Total Savings / Gross Income'));
      ratios.appendChild(mkRatio('Debt-to-Income Ratio', dtiRatio,     'Annual Debt Payments / Gross Income'));
      netPanel.appendChild(ratios);
    }

    cfLayout.appendChild(netPanel);
    expSec.body.appendChild(cfLayout);

    host.appendChild(expSec.el);
  }

  // ══════════════════════════════════════════════════
  // 17. ESTATE PLAN
  // ══════════════════════════════════════════════════
  const ep = goalsData.estatePlan;
  const epYear = goalsData.estatePlanYear;
  if ((ep && ep.length) || epYear) {
    const epSec = makePresSection('Estate Plan');
    epSec.body.appendChild(buildInfoGrid([
      { label: 'Plan Components',            value: Array.isArray(ep) ? ep.join(', ') : ep, span: 2 },
      { label: 'Year Established / Updated', value: epYear },
    ]));
    host.appendChild(epSec.el);
  }
}

/* ==========================================================================
   PLAN MODE
   ========================================================================== */

const PLAN_ROR_OPTS = ['','100/0','90/10','80/20','70/30','60/40','50/50','40/60','30/70','20/80','10/90','0/100'];

const planSections = [
  { id: 'documents', title: 'Documents', color: 'blue', fields: [
    { key: 'investmentStatementOnFile', label: 'Investment Statement on File?',       type: 'checkbox' },
    { key: 'taxReturnOnFile',           label: 'Tax Return on File?',                 type: 'checkbox' },
    { key: 'planningLifeRequested',     label: 'Planning Life Requested?',            type: 'checkbox' },
    { key: 'documentsEmailDate',        label: 'Date E-mail Request for Documents Sent', type: 'date' },
  ]},
  { id: 'baseFacts', title: 'Base Facts', color: 'teal', fields: [
    { key: 'clientRetirementDate', label: 'Client Retirement Date',     type: 'date' },
    { key: 'spouseRetirementDate', label: 'Spouse Retirement Date',     type: 'date' },
    { key: 'eMoneyRorPre',  label: 'eMoney RoR Pre-Retirement',  type: 'select', options: PLAN_ROR_OPTS },
    { key: 'eMoneyRorPost', label: 'eMoney RoR Post-Retirement', type: 'select', options: PLAN_ROR_OPTS },
    { key: 'baseFactsNotes', label: 'Base Facts Notes', type: 'textarea' },
  ]},
  { id: 'observations', title: 'Observations', color: 'amber', fields: [
    { key: 'currentCashFlow',    label: 'Current Cash Flow',                  type: 'select', options: ['','Surplus','Deficit','Break-Even'] },
    { key: 'maxingEmployerPlan', label: 'Maxing Employer Retirement Plan?',   type: 'select', options: ['','Yes','No','N/A'] },
    { key: 'excessCashGoing',    label: 'Where is Excess Cash Going?',        type: 'select', options: ['','Cash','Qualified','Roth','Non-Qualified','Other'] },
    { key: 'extraCashReserves',  label: 'How Much Extra Cash Reserves?',      type: 'text' },
    { key: 'onTrackToRetire',    label: 'Are They on Track to Retire?',       type: 'select', options: ['','Yes','No','Unsure'] },
    { key: 'observationNotes',   label: 'Observation Notes',                  type: 'textarea' },
  ]},
  { id: 'taxes', title: 'Taxes', color: 'violet', fields: [
    { key: 'filingStatus',         label: 'Filing Status This Year',                  type: 'select', options: ['','Married Filing Jointly','Married Filing Separately','Single','Head of Household'] },
    { key: 'taxReturnDifferences', label: 'Differences From Prior Tax Return',        type: 'textarea' },
    { key: 'taxBracket',           label: 'What Tax Bracket Are They in Now?',        type: 'textarea' },
    { key: 'charitablyInclined',   label: 'Charitably Inclined?',                     type: 'textarea' },
    { key: 'taxesBeforeRMDs',      label: 'Taxes: Higher or Lower Before RMDs Start', type: 'select', options: ['','Higher','Lower','Same'] },
    { key: 'taxesAfterRMDs',       label: 'Taxes: Higher or Lower After RMDs Start',  type: 'select', options: ['','Higher','Lower','Same'] },
    { key: 'highLevelStrategies',  label: 'High Level Strategies For This Year',      type: 'textarea' },
    { key: 'unrealizedCapGains',   label: 'Unrealized Capital Gains?',                type: 'textarea' },
  ]},
  { id: 'strategiesCF', title: 'Strategies \u2014 Cash Flow & Retirement', color: 'emerald', fields: [
    { key: 'cashReservesStrategy',    label: 'Cash Reserves Strategy',     type: 'textarea' },
    { key: 'ssClaimingStrategy',      label: 'SS Claiming Strategy',       type: 'textarea' },
    { key: 'otherCashFlowStrategies', label: 'Other Cash Flow Strategies', type: 'textarea' },
  ]},
  { id: 'strategiesTax', title: 'Strategies \u2014 Taxes', color: 'indigo', fields: [
    { key: 'rothConversionStrategy',    label: 'Roth Conversion Strategy',           type: 'textarea' },
    { key: 'charitableStrategies',      label: 'Charitable Strategies',              type: 'textarea' },
    { key: 'capitalGains',              label: 'Capital Gains',                      type: 'textarea' },
    { key: 'capitalGainsWorksheet',     label: 'Capital Gains Worksheet Requested?', type: 'select', options: ['','Yes','No','N/A'] },
    { key: 'employerPlanContributions', label: 'Employer Plan Contributions',        type: 'textarea' },
    { key: 'rothIraContributions',      label: 'Roth / IRA Contributions',           type: 'textarea' },
    { key: 'otherTaxStrategies',        label: 'Other Tax Strategies',               type: 'textarea' },
  ]},
  { id: 'strategiesInvest', title: 'Strategies \u2014 Investment Planning', color: 'sky', fields: [
    { key: 'pureIPS',             label: 'Pure IPS',                                   type: 'text' },
    { key: 'outside401kRequest',  label: 'Outside 401(k) Request',                     type: 'select', options: ['','Requested','Not Requested','Link Outside Retirement Plan to Pure'] },
    { key: 'annuityStrategy',     label: 'Annuity Strategy',                           type: 'textarea' },
    { key: 'morningstarEntered',  label: 'Morningstar Entered During Assessment?',      type: 'select', options: ['','Yes','No'] },
    { key: 'otherInvestRequests', label: 'Other Investment Requests',                  type: 'textarea' },
  ]},
  { id: 'strategiesInsurance', title: 'Strategies \u2014 Insurance & Risk', color: 'rose', fields: [
    { key: 'lifeInsurance', label: 'Life Insurance',  type: 'textarea' },
    { key: 'pAndC',         label: 'P&C',             type: 'textarea' },
    { key: 'ltcDisability', label: 'LTC / Disability', type: 'textarea' },
  ]},
  { id: 'strategiesEstate', title: 'Strategies \u2014 Estate Planning', color: 'orange', fields: [
    { key: 'estatePlanStatus',      label: 'Estate Plan',            type: 'select', options: ['','Has Current Estate Plan','Needs Updated Estate Plan','Needs Estate Plan','No Estate Plan','No Estate Plan Needed'] },
    { key: 'estatePlanningDetails', label: 'Estate Planning Details', type: 'textarea' },
  ]},
  { id: 'altScenarios', title: 'Alternative Scenarios', color: 'green', fields: [
    { key: 'alternativeScenarios', label: 'Alternative Scenarios', type: 'select', options: ['','Yes','No'] },
  ]},
];

/* ---------- Plan Mode navigation ---------- */

function showPlanMode() {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('is-active'));
  document.getElementById('screen-plan').classList.add('is-active');
  document.body.classList.add('plan-mode');
  renderPlanMode();
}

function hidePlanMode() {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('is-active'));
  document.getElementById('screen-pq').classList.add('is-active');
  document.body.classList.remove('plan-mode');
}

/* ---------- Plan Mode render ---------- */

function renderPlanMode() {
  renderPlanSummary();
  renderPlanForm();
}

function _planFinancials() {
  const pq = getPresData();
  const assets = pq.assets || {};
  const liabs  = pq.liabilities || {};
  const income = pq.income || {};
  const taxes  = pq.taxesExpenses || {};
  const num  = v => parseCommas(String(v || '')) || 0;
  const sumA = (arr, key) => (arr || []).filter(Boolean).reduce((s, r) => s + num(r[key]), 0);

  const taxDef   = sumA(assets.taxDeferred, 'marketValue');
  const rothVal  = sumA(assets.roth,        'marketValue');
  const taxable  = sumA(assets.taxable,     'marketValue');
  const cashCd   = sumA(assets.cashCd,      'marketValue');
  const plan529  = sumA(assets.plan529,     'marketValue');
  const hsaVal   = sumA(assets.hsa,         'marketValue');
  const reAssets = sumA(assets.realEstate,  'marketValue');
  const busOther = sumA(assets.businessOther,'marketValue');
  const investable   = taxDef + rothVal + taxable + cashCd + plan529 + hsaVal;
  const totalAssets  = investable + reAssets + busOther;
  const totalLiab    = sumA(liabs.items, 'amount');
  const netWorth     = totalAssets - totalLiab;

  const empInc   = sumA(income.employment,    'annualAmount');
  const ssInc    = sumA(income.socialSecurity,'annualAmount');
  const pensInc  = sumA(income.pension,       'annualAmount');
  const otherInc = sumA(income.other,         'annualAmount');
  const totalIncome = empInc + ssInc + pensInc + otherInc;

  const fedTax  = num(taxes.federalTax);
  const stTax   = num(taxes.stateTax);
  const ficaTax = num(taxes.ficaTax);
  const totalTax = fedTax + stTax + ficaTax;

  const totalSavings = sumA(assets.taxDeferred,'personalAdditions')
    + sumA(assets.roth,'personalAdditions') + sumA(assets.taxable,'personalAdditions');
  const debtPayments = (liabs.items || []).filter(Boolean)
    .reduce((s, r) => s + num(r.payment) * 12, 0);
  const expRows  = (taxes.expenses || []).filter(Boolean);
  const livingExp = num(taxes.livingExpenses) + expRows.filter(r => r._source !== 'liability' && r._source !== 'insurance')
    .reduce((s, r) => s + num(r.amount), 0);
  const netCF = totalIncome - totalTax - totalSavings - debtPayments - livingExp;

  return { pq, assets, taxDef, rothVal, taxable, cashCd, plan529, hsaVal,
           investable, reAssets, busOther, totalAssets, totalLiab, netWorth,
           empInc, ssInc, pensInc, otherInc, totalIncome,
           fedTax, stTax, ficaTax, totalTax, totalSavings, debtPayments, livingExp, netCF };
}

function renderPlanSummary() {
  const host = document.getElementById('plan-summary');
  clearChildren(host);

  const f = _planFinancials();
  const { pq, investable, reAssets, busOther, totalAssets, totalLiab, netWorth,
          empInc, ssInc, pensInc, otherInc, totalIncome,
          taxDef, rothVal, taxable, cashCd, plan529, hsaVal,
          fedTax, stTax, ficaTax, totalTax, totalSavings, debtPayments, livingExp, netCF } = f;

  const family  = pq.family        || {};
  const contact = pq.contact       || {};
  const assets  = pq.assets        || {};
  const liabs   = pq.liabilities   || {};
  const income  = pq.income        || {};
  const taxes   = pq.taxesExpenses || {};
  const ins     = pq.insurance     || {};
  const rels    = pq.relationships || {};
  const emp     = pq.employment    || {};
  const goals   = pq.goals         || {};

  const num = v => parseCommas(String(v || '')) || 0;
  const cleanRows = arr => (arr || []).filter(r => r && Object.values(r).some(v => v != null && String(v).trim() !== '' && !String(v).startsWith('_')));

  const c1Name = [family.client1FirstName, family.client1LastName].filter(Boolean).join(' ') || 'Client 1';
  const c2Name = [family.client2FirstName, family.client2LastName].filter(Boolean).join(' ');
  const clientDisplay = c2Name ? `${c1Name} & ${c2Name}` : c1Name;
  const age1 = family.client1Age ? ` (age ${family.client1Age})` : '';
  const age2 = family.client2Age ? ` (age ${family.client2Age})` : '';
  const ageDisplay = c2Name ? `${c1Name}${age1} & ${c2Name}${age2}` : `${c1Name}${age1}`;
  const cityState = [contact.city, contact.state].filter(Boolean).join(', ');

  /* ── builders ── */
  const mkRow = (label, value, total = false) => {
    const d = el('div', { className: 'plan-sum-row' + (total ? ' is-total' : '') });
    d.appendChild(el('span', { className: 'plan-sum-label', textContent: label }));
    d.appendChild(el('span', { className: 'plan-sum-value', textContent: value }));
    return d;
  };
  const mkAcctRow = (label, sub, value, indent = false) => {
    const d = el('div', { className: 'plan-sum-row plan-sum-row-multi' + (indent ? ' plan-sum-row-indent' : '') });
    const leftCol = el('div', { className: 'plan-sum-row-main' });
    leftCol.appendChild(el('span', { className: 'plan-sum-label', textContent: label }));
    if (sub) leftCol.appendChild(el('span', { className: 'plan-sum-sublabel', textContent: sub }));
    d.appendChild(leftCol);
    if (value) d.appendChild(el('span', { className: 'plan-sum-value', textContent: value }));
    return d;
  };
  const mkGroupHead = (label, value) => {
    const d = el('div', { className: 'plan-sum-grouphead' });
    d.appendChild(el('span', { className: 'plan-sum-grouphead-label', textContent: label }));
    if (value) d.appendChild(el('span', { className: 'plan-sum-grouphead-value', textContent: value }));
    return d;
  };
  const mkSection = title => {
    const s = el('div', { className: 'plan-sum-section' });
    s.appendChild(el('div', { className: 'plan-sum-title', textContent: title }));
    return s;
  };

  /* ── Client header ── */
  const hdr = el('div', { className: 'plan-sum-header' });
  hdr.appendChild(el('div', { className: 'plan-sum-client', textContent: clientDisplay }));
  if (ageDisplay !== clientDisplay) hdr.appendChild(el('div', { className: 'plan-sum-ages', textContent: ageDisplay }));
  if (cityState) hdr.appendChild(el('div', { className: 'plan-sum-location', textContent: cityState }));
  host.appendChild(hdr);

  /* ── Contact ── */
  if (contact.email || contact.phone || contact.addressLine1 || contact.city) {
    const s = mkSection('Contact');
    if (contact.email) s.appendChild(mkRow('Email', contact.email));
    if (contact.phone) s.appendChild(mkRow('Phone', contact.phone));
    const addr = [contact.addressLine1, contact.city, contact.state, contact.zip].filter(Boolean).join(', ');
    if (addr) s.appendChild(mkRow('Address', addr));
    host.appendChild(s);
  }

  /* ── Family ── */
  const children = cleanRows(family.children);
  if (family.maritalStatus || children.length) {
    const s = mkSection('Family');
    if (family.maritalStatus) s.appendChild(mkRow('Marital Status', family.maritalStatus));
    if (children.length) {
      s.appendChild(mkRow('Children', children.length.toString(), true));
      children.forEach(c => {
        const cn = [c.firstName, c.lastName].filter(Boolean).join(' ') || 'Child';
        const sub = [c.age ? 'age ' + c.age : '', c.dateOfBirth ? 'DOB ' + c.dateOfBirth : ''].filter(Boolean).join(' · ');
        s.appendChild(mkAcctRow(cn, sub, '', true));
      });
    }
    host.appendChild(s);
  }

  /* ── Employment ── */
  const addPerson = (sec, name, status, empData) => {
    sec.appendChild(mkRow(name, status || '—', true));
    const sub = [empData.occupation, empData.employer].filter(Boolean).join(' @ ');
    if (sub) sec.appendChild(mkAcctRow('Occupation', '', sub, true));
    if (empData.annualIncome) sec.appendChild(mkAcctRow('Annual Income', '', fmt$(num(empData.annualIncome)), true));
    if (empData.retirementDate) sec.appendChild(mkAcctRow('Retirement Date', '', empData.retirementDate, true));
  };
  const empSec = mkSection('Employment');
  addPerson(empSec, c1Name, pqState.employmentClient1, emp.client1 || {});
  if (pqState.hasSpouse) addPerson(empSec, c2Name || 'Spouse', pqState.employmentClient2, emp.client2 || {});
  host.appendChild(empSec);

  /* ── Net Worth ── */
  const nwSec = mkSection('Net Worth');
  nwSec.appendChild(el('div', { className: 'plan-sum-big ' + (netWorth >= 0 ? 'plan-sum-positive' : 'plan-sum-negative'), textContent: fmt$(netWorth) }));
  nwSec.appendChild(mkRow('Total Assets', fmt$(totalAssets)));
  nwSec.appendChild(mkRow('Total Liabilities', fmt$(totalLiab), true));
  host.appendChild(nwSec);

  /* ── Investment Accounts (detailed) ── */
  const acctGroups = [
    { key: 'taxDeferred', label: 'Tax-Deferred',  total: taxDef },
    { key: 'roth',        label: 'Roth',          total: rothVal },
    { key: 'taxable',     label: 'Taxable',       total: taxable },
    { key: 'cashCd',      label: 'Cash / CD',     total: cashCd },
    { key: 'plan529',     label: '529 Plan',      total: plan529 },
    { key: 'hsa',         label: 'HSA',           total: hsaVal },
  ];
  const hasAnyAcct = acctGroups.some(g => cleanRows(assets[g.key]).length > 0);
  if (hasAnyAcct) {
    const s = mkSection('Investment Accounts');
    acctGroups.forEach(g => {
      const rows = cleanRows(assets[g.key]);
      if (!rows.length) return;
      s.appendChild(mkGroupHead(g.label, fmt$(g.total)));
      rows.forEach(r => {
        const lbl = r.custodian || r.description || 'Account';
        const subParts = [r.type, r.ownership, r.beneficiary ? 'Bene: ' + r.beneficiary : ''].filter(Boolean);
        if (num(r.personalAdditions) > 0) subParts.push('Adds: ' + fmt$(num(r.personalAdditions)) + '/yr');
        s.appendChild(mkAcctRow(lbl, subParts.join(' · '), fmt$(num(r.marketValue)), true));
      });
    });
    s.appendChild(mkRow('Investable Total', fmt$(investable), true));
    host.appendChild(s);
  }

  /* ── Real Estate ── */
  const reRows = cleanRows(assets.realEstate);
  if (reRows.length) {
    const s = mkSection('Real Estate');
    reRows.forEach(r => {
      const subParts = [r.ownership];
      if (num(r.remainingLoan)) subParts.push('Loan: ' + fmt$(num(r.remainingLoan)));
      if (num(r.payment)) subParts.push('Pmt: ' + fmt$(num(r.payment)) + '/mo');
      if (r.interestRate) subParts.push(r.interestRate + '%');
      s.appendChild(mkAcctRow(r.description || 'Property', subParts.filter(Boolean).join(' · '), fmt$(num(r.marketValue))));
    });
    s.appendChild(mkRow('Total', fmt$(reAssets), true));
    host.appendChild(s);
  }

  /* ── Business & Other ── */
  const busRows = cleanRows(assets.businessOther);
  if (busRows.length) {
    const s = mkSection('Business & Other Assets');
    busRows.forEach(r => {
      const subParts = [r.ownership, num(r.costBasis) ? 'Basis: ' + fmt$(num(r.costBasis)) : ''].filter(Boolean);
      s.appendChild(mkAcctRow(r.description || 'Asset', subParts.join(' · '), fmt$(num(r.marketValue))));
    });
    s.appendChild(mkRow('Total', fmt$(busOther), true));
    host.appendChild(s);
  }

  /* ── Liabilities ── */
  const liabRows = cleanRows(liabs.items);
  if (liabRows.length) {
    const s = mkSection('Liabilities');
    liabRows.forEach(r => {
      const subParts = [];
      if (num(r.payment)) subParts.push('Pmt: ' + fmt$(num(r.payment)) + '/mo');
      if (r.interestRate) subParts.push(r.interestRate + '%');
      if (r.term) subParts.push(r.term);
      s.appendChild(mkAcctRow(r.description || 'Liability', subParts.join(' · '), fmt$(num(r.amount))));
    });
    s.appendChild(mkRow('Total Liabilities', fmt$(totalLiab), true));
    host.appendChild(s);
  }

  /* ── Income Sources (detailed) ── */
  const incGroups = [
    { key: 'employment',     label: 'Employment' },
    { key: 'socialSecurity', label: 'Social Security' },
    { key: 'pension',        label: 'Pension' },
    { key: 'other',          label: 'Other' },
  ];
  const hasAnyInc = incGroups.some(g => cleanRows(income[g.key]).length > 0);
  if (hasAnyInc) {
    const s = mkSection('Income Sources');
    incGroups.forEach(g => {
      const rows = cleanRows(income[g.key]);
      if (!rows.length) return;
      const total = rows.reduce((sum, r) => sum + num(r.annualAmount), 0);
      s.appendChild(mkGroupHead(g.label, fmt$(total)));
      rows.forEach(r => {
        const label = r.description || r.owner || g.label;
        const subParts = [];
        if (r.owner && r.description) subParts.push(r.owner);
        if (r.startDate) subParts.push('Start: ' + r.startDate);
        if (r.cola) subParts.push('COLA: ' + r.cola + '%');
        s.appendChild(mkAcctRow(label, subParts.join(' · '), fmt$(num(r.annualAmount)), true));
      });
    });
    s.appendChild(mkRow('Total Income', fmt$(totalIncome), true));
    host.appendChild(s);
  }

  /* ── Annual Cash Flow ── */
  const cfSec = mkSection('Annual Cash Flow');
  cfSec.appendChild(mkRow('Gross Income', fmt$(totalIncome)));
  if (totalTax)      cfSec.appendChild(mkRow('Less: Tax',       '(' + fmt$(totalTax) + ')'));
  if (totalSavings)  cfSec.appendChild(mkRow('Less: Savings',   '(' + fmt$(totalSavings) + ')'));
  if (debtPayments)  cfSec.appendChild(mkRow('Less: Debt Pmts', '(' + fmt$(debtPayments) + ')'));
  if (livingExp)     cfSec.appendChild(mkRow('Less: Living Exp','(' + fmt$(livingExp) + ')'));
  const netRow = mkRow('Net Cash Flow', fmt$(netCF), true);
  netRow.querySelector('.plan-sum-value').classList.add(netCF >= 0 ? 'plan-sum-positive' : 'plan-sum-negative');
  cfSec.appendChild(netRow);
  host.appendChild(cfSec);

  /* ── Taxes ── */
  if (totalTax > 0) {
    const s = mkSection('Taxes');
    if (fedTax)  s.appendChild(mkRow('Federal', fmt$(fedTax)));
    if (stTax)   s.appendChild(mkRow('State',   fmt$(stTax)));
    if (ficaTax) s.appendChild(mkRow('FICA',    fmt$(ficaTax)));
    s.appendChild(mkRow('Total Tax', fmt$(totalTax), true));
    if (totalIncome > 0) s.appendChild(mkRow('Effective Rate', ((totalTax / totalIncome) * 100).toFixed(1) + '%'));
    host.appendChild(s);
  }

  /* ── Expenses ── */
  const expRows = cleanRows(taxes.expenses);
  if (expRows.length) {
    const s = mkSection('Annual Expenses');
    expRows.forEach(r => {
      s.appendChild(mkAcctRow(r.description || 'Expense', r.notes || '', fmt$(num(r.amount))));
    });
    const totalExp = expRows.reduce((sum, r) => sum + num(r.amount), 0);
    s.appendChild(mkRow('Total Expenses', fmt$(totalExp), true));
    host.appendChild(s);
  }

  /* ── Insurance ── */
  const insRows = cleanRows(ins.policies);
  if (insRows.length) {
    const s = mkSection('Insurance');
    insRows.forEach(r => {
      const label = [r.company, r.type].filter(Boolean).join(' \u2014 ') || 'Policy';
      const subParts = [];
      if (r.insured) subParts.push('Insured: ' + r.insured);
      if (r.owner)   subParts.push('Owner: ' + r.owner);
      if (num(r.benefit))   subParts.push('Benefit: ' + fmt$(num(r.benefit)));
      if (num(r.cashValue)) subParts.push('CV: ' + fmt$(num(r.cashValue)));
      if (r.beneficiary)    subParts.push('Bene: ' + r.beneficiary);
      const val = num(r.annualPremium) ? fmt$(num(r.annualPremium)) + '/yr' : '';
      s.appendChild(mkAcctRow(label, subParts.join(' · '), val));
    });
    host.appendChild(s);
  }

  /* ── Professional Relationships ── */
  const relRows = cleanRows(rels.professionals);
  if (relRows.length) {
    const s = mkSection('Professional Relationships');
    relRows.forEach(r => {
      const label = r.role || 'Professional';
      const subParts = [r.firmName, [r.city, r.state].filter(Boolean).join(', ')].filter(Boolean);
      s.appendChild(mkAcctRow(label, subParts.join(' · '), r.name || '—'));
    });
    host.appendChild(s);
  }

  /* ── Estate Plan & Goals ── */
  const ep = goals.estatePlan;
  const goalsTxt = (goals.goals || '').trim();
  if ((ep && ep.length) || goals.estatePlanYear || goalsTxt) {
    const s = mkSection('Estate Plan & Goals');
    if (ep && ep.length)         s.appendChild(mkRow('Plan Components', Array.isArray(ep) ? ep.join(', ') : ep));
    if (goals.estatePlanYear)    s.appendChild(mkRow('Year Established', goals.estatePlanYear));
    if (goalsTxt)                s.appendChild(el('div', { className: 'plan-sum-notes', textContent: goalsTxt }));
    host.appendChild(s);
  }
}

function renderPlanForm() {
  const host = document.getElementById('plan-form-sections');
  clearChildren(host);

  const saved = (getStore().planOrder) || {};
  const f = _planFinancials();

  // Smart prefills
  const netCF = f.netCF;
  const cashFlowPrefill = saved.currentCashFlow ||
    (f.totalIncome > 0 ? (netCF > 50 ? 'Surplus' : netCF < -50 ? 'Deficit' : 'Break-Even') : '');
  const prefill = {
    clientRetirementDate: f.pq.employment?.client1?.retirementDate || '',
    spouseRetirementDate: f.pq.employment?.client2?.retirementDate || '',
    currentCashFlow: cashFlowPrefill,
  };

  planSections.forEach(section => {
    const card = el('div', { className: `plan-form-card plan-card-${section.color}` });
    card.appendChild(el('div', { className: 'plan-card-header', textContent: section.title }));
    const body = el('div', { className: 'plan-card-body' });

    section.fields.forEach(field => {
      const val = saved[field.key] !== undefined ? saved[field.key] : (prefill[field.key] || '');
      const row = el('div', { className: 'plan-field-row' + (field.type === 'checkbox' ? ' plan-field-row--check' : '') });

      if (field.type === 'checkbox') {
        const inp = el('input', { type: 'checkbox', id: `plan-${field.key}`, name: field.key });
        if (val === true || val === 'true') inp.checked = true;
        row.appendChild(inp);
        row.appendChild(el('label', { htmlFor: `plan-${field.key}`, textContent: field.label }));
      } else {
        row.appendChild(el('label', { htmlFor: `plan-${field.key}`, className: 'plan-field-label', textContent: field.label }));
        if (field.type === 'select') {
          const sel = el('select', { id: `plan-${field.key}`, name: field.key, className: 'plan-field-input' });
          (field.options || []).forEach(opt => {
            const o = el('option', { value: opt, textContent: opt || '\u2014 Select \u2014' });
            if (val === opt) o.selected = true;
            sel.appendChild(o);
          });
          row.appendChild(sel);
        } else if (field.type === 'textarea') {
          const ta = el('textarea', { id: `plan-${field.key}`, name: field.key, className: 'plan-field-input plan-field-ta', rows: 3 });
          ta.value = val || '';
          row.appendChild(ta);
        } else {
          const inp = el('input', { type: field.type || 'text', id: `plan-${field.key}`, name: field.key, className: 'plan-field-input' });
          inp.value = val || '';
          row.appendChild(inp);
        }
      }
      body.appendChild(row);
    });

    card.appendChild(body);
    host.appendChild(card);
  });
}

function collectPlanForm() {
  const form = document.getElementById('plan-form');
  const data = {};
  if (!form) return data;
  form.querySelectorAll('input, select, textarea').forEach(inp => {
    if (!inp.name) return;
    data[inp.name] = inp.type === 'checkbox' ? inp.checked : inp.value;
  });
  return data;
}

function savePlanOrder() {
  const data = collectPlanForm();
  const store = getStore();
  store.planOrder = data;
  store.workflow = store.workflow || {};
  store.workflow.planOrderSavedAt = new Date().toISOString();
  saveStore(store);
  const btn = document.getElementById('plan-save-btn');
  const orig = btn.textContent;
  btn.textContent = 'Sent \u2713';
  btn.disabled = true;
  setTimeout(() => { btn.textContent = orig; btn.disabled = false; }, 1500);
}

/* ==========================================================================
   EVENT BINDING
   ========================================================================== */

function doSavePQ(pqForm, { silent = false } = {}) {
  const data = collectForm(pqForm);
  data.family = data.family || {};
  data.family.hasSpouse = pqState.hasSpouse;
  data.family.hasChildren = pqState.hasChildren;
  data.assets = data.assets || {};
  data._assetChecks = { ...pqState.assetChecks };
  data.employment = data.employment || {};
  data.employment.client1 = data.employment.client1 || {};
  data.employment.client1.status = pqState.employmentClient1;
  if (pqState.hasSpouse) {
    data.employment.client2 = data.employment.client2 || {};
    data.employment.client2.status = pqState.employmentClient2;
  }

  const store = getStore();
  const existing = store.pqTemplate || {};
  // Preserve asset arrays for sections whose sub-form isn't currently rendered
  // (checkbox off ⇒ collectForm returns nothing for that path, so merging the
  //  prior saved arrays keeps them intact instead of wiping them).
  Object.keys(pqState.assetChecks).forEach(k => {
    if (!pqState.assetChecks[k] && Array.isArray(existing.assets?.[k])) {
      data.assets[k] = existing.assets[k];
    }
  });
  const merged = deepMerge(deepClone(existing), data);
  store.pqTemplate = merged;
  store.workflow = store.workflow || {};
  store.workflow.pqSavedAt = new Date().toISOString();
  saveStore(store);
  _pqDraft = {};

  if (window.parent !== window) {
    window.parent.postMessage({ type: 'PQ_SAVE', payload: merged }, '*');
  }

  if (!silent) {
    const btn = document.getElementById('pq-save-btn');
    const topBtn = document.getElementById('topbar-save-btn');
    const orig = 'Save';
    btn.textContent = 'Saved \u2713';
    btn.disabled = true;
    topBtn.textContent = 'Saved \u2713';
    topBtn.disabled = true;
    setTimeout(() => {
      btn.textContent = orig; btn.disabled = false;
      topBtn.textContent = orig; topBtn.disabled = false;
    }, 1500);
  }
  return data;
}

let _autosaveTimer = null;
let _autosaveStatusTimer = null;
function setAutosaveStatus(state) {
  const el = document.getElementById('autosave-status');
  if (!el) return;
  clearTimeout(_autosaveStatusTimer);
  if (state === 'saving') {
    el.textContent = 'Saving\u2026';
    el.className = 'autosave-status is-saving';
  } else if (state === 'saved') {
    el.textContent = 'Saved \u2713';
    el.className = 'autosave-status is-saved';
    _autosaveStatusTimer = setTimeout(() => {
      el.textContent = '';
      el.className = 'autosave-status';
    }, 2000);
  } else {
    el.textContent = '';
    el.className = 'autosave-status';
  }
}
function scheduleAutosave(pqForm) {
  clearTimeout(_autosaveTimer);
  setAutosaveStatus('saving');
  _autosaveTimer = setTimeout(() => {
    try {
      doSavePQ(pqForm, { silent: true });
      setAutosaveStatus('saved');
    } catch (err) {
      console.error('Autosave failed', err);
      setAutosaveStatus('');
    }
  }, 800);
}

function bindEvents() {
  // PQ form
  const pqForm = document.getElementById('pq-form');
  pqForm.addEventListener('input', e => {
    scheduleAutosave(pqForm);
    if (e.target?.dataset?.path === '_advisorNotes') return;
    recalcPQComputed();
  });
  pqForm.addEventListener('change', () => scheduleAutosave(pqForm));
  pqForm.addEventListener('submit', e => {
    e.preventDefault();
    clearTimeout(_autosaveTimer);
    doSavePQ(pqForm, { silent: false });
    setAutosaveStatus('');
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
    pqState._initialized = false;
    _pqDraft = {}; // Clear draft on reset
    renderPQ();
  });

  // Presentation mode
  document.getElementById('topbar-save-btn').addEventListener('click', () => document.getElementById('pq-save-btn').click());
  document.getElementById('plan-mode-btn').addEventListener('click', showPlanMode);
  document.getElementById('plan-edit-btn').addEventListener('click', hidePlanMode);
  document.getElementById('plan-save-btn').addEventListener('click', savePlanOrder);
  document.getElementById('presentation-mode-btn').addEventListener('click', showPresentation);
  document.getElementById('pres-edit-btn').addEventListener('click', hidePresentation);
  document.getElementById('pres-print-btn').addEventListener('click', () => window.print());

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
}

init();
