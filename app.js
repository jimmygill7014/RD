const pqSchema = [
  {
    id: 'family',
    title: 'Family Information',
    note: 'Core household and referral details from the PQ template.',
    fields: [
      { key: 'yourName', label: 'First Name', width: 'medium' },
      { key: 'yourLastName', label: 'Last Name', width: 'medium' },
      { key: 'yourNickname', label: 'Nickname', width: 'medium' },
      { key: 'yourBirthDate', label: 'Birth Date', type: 'date' },
      { key: 'yourAge', label: 'Age', type: 'number' },
      { key: 'spouseName', label: "Spouse's Name", width: 'medium' },
      { key: 'spouseLastName', label: 'Spouse Last Name', width: 'medium' },
      { key: 'yearsMarried', label: 'Years Married', type: 'number' },
      { key: 'spouseNickname', label: 'Spouse Nickname', width: 'medium' },
      { key: 'spouseBirthDate', label: 'Spouse Birth Date', type: 'date' },
      { key: 'spouseAge', label: 'Spouse Age', type: 'number' },
      { key: 'residenceAddress', label: 'Residence Address', width: 'wide' },
      { key: 'city', label: 'City' },
      { key: 'state', label: 'State' },
      { key: 'zipCode', label: 'Zip Code' },
      { key: 'referredBy', label: 'Referred By', width: 'medium' },
      { key: 'referralSource', label: 'Client Name / Source', width: 'medium' },
      { key: 'totalGrandchildren', label: 'Total # of Grandchildren', type: 'number' },
      { key: 'additionalFamilyInfo', label: 'Additional Family Information', type: 'textarea' }
    ],
    tables: [
      {
        key: 'children',
        title: 'Children & Grandchildren',
        columns: [
          { key: 'childNameAge', label: "Child's Name & Age" },
          { key: 'notes', label: 'Notes' }
        ],
        minRows: 4
      }
    ]
  },
  {
    id: 'occupationGoals',
    title: 'Occupation & Goals',
    fields: [
      { key: 'yourJobTitle', label: 'Your Job Title', width: 'medium' },
      { key: 'yourEmployer', label: 'Employer (Last, if Retired)', width: 'wide' },
      { key: 'yourYears', label: '# of Years', type: 'number' },
      { key: 'yourRetirementDate', label: 'Retirement Date', type: 'date' },
      { key: 'spouseJobTitle', label: "Spouse's Job Title", width: 'medium' },
      { key: 'spouseEmployer', label: 'Employer (Last, if Retired)', width: 'wide' },
      { key: 'spouseYears', label: '# of Years', type: 'number' },
      { key: 'spouseRetirementDate', label: 'Retirement Date', type: 'date' },
      { key: 'goals', label: 'Goals', type: 'textarea' },
      { key: 'additionalNotes', label: 'Additional Notes', type: 'textarea' },
      { key: 'estatePlan', label: 'Estate Plan' },
      { key: 'estatePlanYear', label: 'Year Established / Last Updated' }
    ]
  },
  {
    id: 'assetsLiabilities',
    title: 'Assets & Liabilities',
    note: 'All major account, asset, and liability grids from the worksheet.',
    tables: [
      {
        key: 'realEstateAssets',
        title: 'Real Estate Assets',
        columns: [
          { key: 'description', label: 'Description' },
          { key: 'marketValue', label: 'Market Value', type: 'number' },
          { key: 'liabilities', label: 'Liabilities', type: 'number' },
          { key: 'interestRate', label: 'Int. Rate' },
          { key: 'term', label: 'Term' },
          { key: 'paymentPI', label: 'Pmt (P&I)', type: 'number' },
          { key: 'incomeEBT', label: 'Income EBT', type: 'number' },
          { key: 'acquisitionYear', label: 'Acquisition Year' },
          { key: 'acquisitionHow', label: 'How Acquired' },
          { key: 'costBasis', label: 'Cost Basis', type: 'number' },
          { key: 'ownership', label: 'Ownership' }
        ],
        minRows: 2
      },
      {
        key: 'taxDeferred',
        title: 'Tax-Deferred Retirement Accounts',
        columns: [
          { key: 'custodian', label: 'Custodian' },
          { key: 'marketValue', label: 'Market Value', type: 'number' },
          { key: 'personalAdditions', label: 'Personal Additions', type: 'number' },
          { key: 'companyMatch', label: 'Company Match', type: 'number' },
          { key: 'type', label: 'Type' },
          { key: 'ownership', label: 'Ownership' },
          { key: 'beneficiary', label: 'Beneficiary' }
        ]
      },
      {
        key: 'rothAccounts',
        title: 'Tax-Free Roth Accounts',
        columns: [
          { key: 'custodian', label: 'Custodian' },
          { key: 'marketValue', label: 'Market Value', type: 'number' },
          { key: 'personalAdditions', label: 'Personal Additions', type: 'number' },
          { key: 'companyMatch', label: 'Company Match', type: 'number' },
          { key: 'rothType', label: 'Roth Type' },
          { key: 'yearEstablished', label: 'Year Est.' },
          { key: 'ownership', label: 'Ownership' },
          { key: 'beneficiary', label: 'Beneficiary' }
        ]
      },
      {
        key: 'taxableAccounts',
        title: 'Taxable Non-Retirement Accounts',
        columns: [
          { key: 'custodian', label: 'Custodian' },
          { key: 'marketValue', label: 'Market Value', type: 'number' },
          { key: 'additions', label: 'Additions', type: 'number' },
          { key: 'costBasis', label: 'Cost Basis', type: 'number' },
          { key: 'ownership', label: 'Ownership' },
          { key: 'variableFixed', label: 'Variable/Fixed' },
          { key: 'issueDate', label: 'Issue Date', type: 'date' },
          { key: 'beneficiary', label: 'Beneficiary' }
        ]
      },
      {
        key: 'cashCds',
        title: "Cash & CD's",
        columns: [
          { key: 'description', label: 'Description' },
          { key: 'marketValue', label: 'Market Value', type: 'number' },
          { key: 'interestRate', label: 'Int. Rate' },
          { key: 'owner', label: 'Owner' }
        ]
      },
      {
        key: 'businessOther',
        title: 'Business & Other',
        columns: [
          { key: 'description', label: 'Description' },
          { key: 'marketValue', label: 'Market Value', type: 'number' },
          { key: 'costBasis', label: 'Cost Basis', type: 'number' },
          { key: 'owner', label: 'Owner' }
        ]
      },
      {
        key: 'otherLiabilities',
        title: 'Other Liabilities',
        columns: [
          { key: 'description', label: 'Description' },
          { key: 'payment', label: 'Payment', type: 'number' },
          { key: 'amount', label: 'Amount', type: 'number' },
          { key: 'interestRate', label: 'Int. Rate' },
          { key: 'term', label: 'Term' }
        ]
      },
      {
        key: 'insurance',
        title: 'Insurance Information',
        columns: [
          { key: 'company', label: 'Company' },
          { key: 'type', label: 'Type' },
          { key: 'benefit', label: 'Death/Daily Benefit', type: 'number' },
          { key: 'insured', label: 'Insured' },
          { key: 'owner', label: 'Owner' },
          { key: 'policyDate', label: 'Policy Date', type: 'date' },
          { key: 'annualPremium', label: 'Annual Premium', type: 'number' },
          { key: 'cashValue', label: 'Cash Value', type: 'number' },
          { key: 'beneficiary', label: 'Beneficiary' }
        ]
      }
    ]
  },
  {
    id: 'incomeTax',
    title: 'Income & Tax Information',
    fields: [
      { key: 'capLossCarryForward', label: 'Cap Loss Carry Forward', type: 'number' },
      { key: 'taxesPaid', label: 'Taxes Paid', type: 'number' },
      { key: 'federalTax', label: 'Federal Tax', type: 'number' },
      { key: 'stateTax', label: 'State Tax', type: 'number' },
      { key: 'taxableIncome', label: 'Taxable Income', type: 'number' },
      { key: 'standardItemizedDeduction', label: 'Stnd/Item Ded.', type: 'number' },
      { key: 'ficaTax', label: 'FICA Tax', type: 'number' },
      { key: 'totalTaxesPaid', label: 'Total Taxes Paid', type: 'number' },
      { key: 'currentTotalIncome', label: 'Current Total Income', type: 'number' },
      { key: 'savingsRatio', label: 'Savings Ratio (%)', type: 'number' },
      { key: 'debtToIncomeRatio', label: 'Debt to Income Ratio (%)', type: 'number' }
    ]
  },
  {
    id: 'expenses',
    title: 'Expenses & Liabilities',
    fields: [
      { key: 'currentExpenseTotal', label: 'Current Total', type: 'number' },
      { key: 'totalSavings', label: 'Total Savings', type: 'number' },
      { key: 'totalCashFlow', label: 'Total Cash Flow +/-', type: 'number' }
    ],
    tables: [
      {
        key: 'expensesLiabilities',
        title: 'Expense Lines',
        columns: [
          { key: 'description', label: 'Description' },
          { key: 'amount', label: 'Amount', type: 'number' },
          { key: 'startDate', label: 'Start Date', type: 'date' },
          { key: 'endDate', label: 'End Date', type: 'date' },
          { key: 'cola', label: 'COLA %', type: 'number' },
          { key: 'notes', label: 'Notes' }
        ],
        minRows: 4,
        starterRows: [
          { description: 'Living Expenses' },
          { description: 'Personal Liabilities' },
          { description: 'Other Liability 1' },
          { description: 'Other Liability 2' }
        ]
      }
    ]
  },
  {
    id: 'relationships',
    title: 'Current Professional Relationships',
    tables: [
      {
        key: 'professionalRelationships',
        title: 'Professional Contacts',
        columns: [
          { key: 'role', label: 'Role' },
          { key: 'name', label: 'Name' },
          { key: 'firmName', label: 'Firm Name' },
          { key: 'city', label: 'City' },
          { key: 'state', label: 'State' }
        ],
        minRows: 3,
        starterRows: [
          { role: 'Financial Advisor' },
          { role: 'Attorney' },
          { role: 'Accountant' }
        ]
      }
    ]
  }
];

const planSchema = [
  {
    id: 'documents',
    title: 'Documents',
    fields: [
      { key: 'investmentStatementsOnFile', label: 'Investment Statements on File', type: 'select', options: ['Yes', 'No', 'Pending'] },
      { key: 'taxReturnOnFile', label: 'Tax Return on File', type: 'select', options: ['Yes', 'No', 'Pending'] },
      { key: 'dateEmailRequestForPlanningLite', label: 'Date E-mail Request for Planning Lite', type: 'date' },
      { key: 'planningLiteRequested', label: 'Planning Lite Requested?', type: 'select', options: ['Yes', 'No'] }
    ]
  },
  {
    id: 'clientInfo',
    title: 'Client Info',
    fields: [
      { key: 'clientName', label: 'Client Name', width: 'medium' },
      { key: 'spouseName', label: 'Spouse Name', width: 'medium' },
      { key: 'clientAge', label: 'Client Age', type: 'number' },
      { key: 'spouseAge', label: 'Spouse Age', type: 'number' },
      { key: 'clientTypeEmployment', label: 'Client Type of Employment' },
      { key: 'spouseTypeEmployment', label: 'Spouse Type of Employment' },
      { key: 'clientEmploymentDetails', label: 'Client Employment Details', width: 'wide' },
      { key: 'spouseEmploymentDetails', label: 'Spouse Employment Details', width: 'wide' },
      { key: 'clientRetirementDate', label: 'Client Retirement Date', type: 'date' },
      { key: 'spouseRetirementDate', label: 'Spouse Retirement Date', type: 'date' },
      { key: 'eMoneyRorPreRetirement', label: 'eMoney RoR (pre-retirement)' },
      { key: 'eMoneyRorPostRetirement', label: 'eMoney RoR (post-retirement)' },
      { key: 'notesBaseFacts', label: 'Notes - Base Facts', type: 'textarea' }
    ]
  },
  {
    id: 'observations',
    title: 'Observations',
    fields: [
      { key: 'currentCashFlow', label: 'Current Cash-Flow', type: 'select', options: ['Surplus', 'Deficit', 'Break-even'] },
      { key: 'maxingEmployerRetirementPlan', label: 'Maxing Employer Retirement Plan?', type: 'select', options: ['Yes', 'No', 'N/A'] },
      { key: 'whereExcessCashGoing', label: 'Where is Excess Cash Going?', width: 'wide' },
      { key: 'wherePullingMoneyFrom', label: 'Where Are They Pulling Money From?', width: 'wide' },
      { key: 'extraCashReserves', label: 'How Much Extra Cash Reserves?', width: 'wide' },
      { key: 'onTrackToRetire', label: 'Are They On Track to Retire?', type: 'select', options: ['Yes', 'No', 'Unknown'] },
      { key: 'observationNotes', label: 'Observation Notes', type: 'textarea' }
    ]
  },
  {
    id: 'observationTaxes',
    title: 'Observation Taxes',
    fields: [
      { key: 'filingStatusThisYear', label: 'Filing Status This Year', type: 'select', options: ['Single', 'Married Filing Jointly', 'Married Filing Separately', 'Head of Household'] },
      { key: 'differencesFromPriorTaxReturn', label: 'Differences from Prior Tax Return' },
      { key: 'taxesHigherOrLowerBeforeRmd', label: 'Taxes: Higher or Lower Before RMD Start', type: 'select', options: ['Higher', 'Lower', 'Same'] },
      { key: 'charitablyInclined', label: 'Charitably Inclined?', type: 'select', options: ['Yes', 'No'] },
      { key: 'taxesHigherOrLowerAfterRmd', label: 'Taxes: Higher or Lower After RMD Start', type: 'select', options: ['Higher', 'Lower', 'Same'] },
      { key: 'highLevelStrategiesThisYear', label: 'High Level Strategies For This Year', width: 'wide' },
      { key: 'unrealizedCapGains', label: 'Unrealized Cap Gains', width: 'wide' }
    ]
  },
  {
    id: 'cashFlowRetirementPlanning',
    title: 'Cash Flow / Retirement Planning',
    fields: [
      { key: 'cashReservesStrategy', label: 'Cash Reserves Strategy', width: 'wide' },
      { key: 'otherCashFlowStrategies', label: 'Other Cash Flow Strategies', width: 'wide' },
      { key: 'ssClaimingStrategy', label: 'SS Claiming Strategy', width: 'wide' }
    ]
  },
  {
    id: 'taxes',
    title: 'Taxes',
    fields: [
      { key: 'rothConversionStrategy', label: 'Roth Conversion Strategy', width: 'wide' },
      { key: 'capitalGains', label: 'Capital Gains', width: 'wide' },
      { key: 'charitableStrategies', label: 'Charitable Strategies', width: 'wide' },
      { key: 'capitalGainsWorksheetRequest', label: 'Capital Gains Worksheet Request?', type: 'select', options: ['Yes', 'No'] },
      { key: 'rothIraContributions', label: 'Roth IRA or IRA Contributions', width: 'wide' },
      { key: 'employerPlanRetirementContributions', label: 'Employer Plan Retirement Contributions', width: 'wide' },
      { key: 'otherTaxStrategies', label: 'Other Tax Strategies', width: 'wide' }
    ]
  },
  {
    id: 'investmentPlanning',
    title: 'Investment Planning',
    fields: [
      { key: 'pureIps', label: 'Pure IPS' },
      { key: 'outside401kRequest', label: 'Outside 401(k) Request', type: 'select', options: ['Requested', 'Not Requested'] },
      { key: 'annuityStrategy', label: 'Annuity Strategy', width: 'wide' },
      { key: 'morningstarEnteredDuringAssessment', label: 'Morningstar Entered During Assessment?', type: 'select', options: ['Yes', 'No'] },
      { key: 'otherInvestmentRequests', label: 'Other Investment Requests', width: 'wide' }
    ]
  },
  {
    id: 'insuranceRiskManagement',
    title: 'Insurance / Risk Management',
    fields: [
      { key: 'lifeInsurance', label: 'Life Insurance', width: 'wide' },
      { key: 'pAndC', label: 'P&C', width: 'wide' },
      { key: 'ltdDisability', label: 'LTD/Disability', width: 'wide' }
    ]
  },
  {
    id: 'estatePlanning',
    title: 'Estate Planning',
    fields: [
      { key: 'estatePlan', label: 'Estate Plan', width: 'wide' },
      { key: 'estatePlanningDetails', label: 'Estate Planning Details', width: 'wide' },
      { key: 'estatePlanOptions', label: 'Estate Plan Options', width: 'wide' }
    ]
  },
  {
    id: 'alternativeScenarios',
    title: 'Alternative Scenarios',
    fields: [
      { key: 'alternativeScenarios', label: 'Alternative Scenarios', type: 'select', options: ['No', 'Yes'] },
      { key: 'alternativeScenarioDescription', label: 'Alternative Scenario Description', type: 'textarea' }
    ]
  }
];

const CENTRAL_STORE_KEY = 'central-advisor-workflow-v1';
let currentScreen = 'pq';
const planPrefillBaseline = {};

const planPrefillMap = {
  'clientInfo.clientName': (pq) => joinName(getDeep(pq, 'family.yourName'), getDeep(pq, 'family.yourLastName')),
  'clientInfo.spouseName': (pq) => joinName(getDeep(pq, 'family.spouseName'), getDeep(pq, 'family.spouseLastName')),
  'clientInfo.clientAge': (pq) => getDeep(pq, 'family.yourAge'),
  'clientInfo.spouseAge': (pq) => getDeep(pq, 'family.spouseAge'),
  'clientInfo.clientTypeEmployment': (pq) => getDeep(pq, 'occupationGoals.yourJobTitle'),
  'clientInfo.spouseTypeEmployment': (pq) => getDeep(pq, 'occupationGoals.spouseJobTitle'),
  'clientInfo.clientEmploymentDetails': (pq) => getDeep(pq, 'occupationGoals.yourEmployer'),
  'clientInfo.spouseEmploymentDetails': (pq) => getDeep(pq, 'occupationGoals.spouseEmployer'),
  'clientInfo.clientRetirementDate': (pq) => getDeep(pq, 'occupationGoals.yourRetirementDate'),
  'clientInfo.spouseRetirementDate': (pq) => getDeep(pq, 'occupationGoals.spouseRetirementDate'),
  'clientInfo.notesBaseFacts': (pq) => getDeep(pq, 'occupationGoals.additionalNotes'),
  'documents.planningLiteRequested': (pq) => (getDeep(pq, 'family.referralSource') ? 'Yes' : undefined),
  'observations.observationNotes': (pq) => getDeep(pq, 'occupationGoals.goals'),
  'observationTaxes.filingStatusThisYear': (pq) => getDeep(pq, 'family.spouseName') ? 'Married Filing Jointly' : 'Single',
  'estatePlanning.estatePlan': (pq) => getDeep(pq, 'occupationGoals.estatePlan'),
  'estatePlanning.estatePlanningDetails': (pq) => getDeep(pq, 'occupationGoals.estatePlanYear')
};

const CLIENT_LABEL_OVERRIDES = {
  'pqTemplate.family.yourName': 'First Name',
  'pqTemplate.family.yourLastName': 'Last Name',
  'pqTemplate.family.spouseName': 'Spouse First Name',
  'pqTemplate.family.spouseLastName': 'Spouse Last Name'
};

function joinName(first, last) {
  return [first, last].filter(Boolean).join(' ').trim();
}

function parsePath(path) {
  return path.replace(/\]/g, '').split(/\.|\[/g).filter(Boolean);
}

function getDeep(source, path) {
  return parsePath(path).reduce((acc, key) => (acc == null ? undefined : acc[key]), source);
}

function setDeep(target, path, value) {
  const keys = parsePath(path);
  let ref = target;
  keys.forEach((key, idx) => {
    const isLast = idx === keys.length - 1;
    const nextKey = keys[idx + 1];
    const nextIsIndex = nextKey !== undefined && /^\d+$/.test(nextKey);
    if (isLast) {
      ref[key] = value;
      return;
    }
    if (ref[key] === undefined) ref[key] = nextIsIndex ? [] : {};
    ref = ref[key];
  });
}

function deepClone(value) {
  return JSON.parse(JSON.stringify(value));
}

function hasData(value) {
  if (value === null || value === undefined) return false;
  if (Array.isArray(value)) return value.some((item) => hasData(item));
  if (typeof value === 'object') return Object.values(value).some((child) => hasData(child));
  return String(value).trim() !== '';
}

function getCentralStore() {
  try {
    const parsed = JSON.parse(localStorage.getItem(CENTRAL_STORE_KEY) || '{}');
    return {
      pqTemplate: parsed.pqTemplate || {},
      planOrder: parsed.planOrder || {},
      workflow: parsed.workflow || {}
    };
  } catch {
    return { pqTemplate: {}, planOrder: {}, workflow: {} };
  }
}

function saveCentralStore(store) {
  localStorage.setItem(CENTRAL_STORE_KEY, JSON.stringify(store));
  refreshDataConsole();
  updatePipeline();
}

function formatTimestamp(ts) {
  if (!ts) return 'Not submitted';
  const date = new Date(ts);
  if (Number.isNaN(date.getTime())) return 'Not submitted';
  return date.toLocaleString([], { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' });
}

function buildSection(section, isOpen, formType) {
  const wrapper = document.createElement('section');
  wrapper.className = 'form-section';

  const header = document.createElement('div');
  header.className = 'section-header';
  header.innerHTML = `<h2>${section.title}</h2><button type=\"button\">${isOpen ? '-' : '+'}</button>`;

  const content = document.createElement('div');
  content.className = 'section-content';
  if (!isOpen) content.style.display = 'none';

  if (section.note) {
    const note = document.createElement('p');
    note.className = 'section-note';
    note.textContent = section.note;
    content.appendChild(note);
  }

  if (section.fields?.length) content.appendChild(buildFields(section, formType));
  if (section.tables?.length) section.tables.forEach((tableDef) => content.appendChild(buildTable(section.id, tableDef)));

  header.addEventListener('click', () => {
    const shown = content.style.display !== 'none';
    content.style.display = shown ? 'none' : 'block';
    header.querySelector('button').textContent = shown ? '+' : '-';
  });

  wrapper.append(header, content);
  return wrapper;
}

function buildFields(section, formType) {
  const grid = document.createElement('div');
  grid.className = 'grid';

  section.fields.forEach((field) => {
    const label = document.createElement('label');
    const widthClass = field.width === 'wide' ? 'field-wide' : field.width === 'medium' ? 'field-medium' : 'field';
    label.className = `${widthClass}${field.type === 'textarea' ? ' field-full' : ''}`;
    if (formType === 'plan') label.classList.add('flag-red');

    const row = document.createElement('div');
    row.className = 'label-row';

    const title = document.createElement('span');
    title.className = 'label';
    title.textContent = field.label;
    row.appendChild(title);

    if (formType === 'plan') {
      const tag = document.createElement('span');
      tag.className = 'prefill-tag';
      tag.textContent = 'Prefilled';
      row.appendChild(tag);
    }

    let input;
    if (field.type === 'textarea') {
      input = document.createElement('textarea');
      input.rows = 3;
    } else if (field.type === 'select') {
      input = document.createElement('select');
      const blank = document.createElement('option');
      blank.value = '';
      blank.textContent = 'Select';
      input.appendChild(blank);
      (field.options || []).forEach((value) => {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = value;
        input.appendChild(option);
      });
    } else {
      input = document.createElement('input');
      input.type = field.type || 'text';
      if (field.type === 'number') input.step = 'any';
    }

    input.dataset.path = `${section.id}.${field.key}`;
    input.name = input.dataset.path;

    label.append(row, input);
    grid.appendChild(label);
  });

  return grid;
}

function buildTable(sectionId, tableDef) {
  const block = document.createElement('div');
  const tools = document.createElement('div');
  tools.className = 'table-tools';
  tools.innerHTML = `<div class=\"table-title\">${tableDef.title}</div>`;

  const addBtn = document.createElement('button');
  addBtn.type = 'button';
  addBtn.className = 'btn-link';
  addBtn.textContent = '+ Add Row';
  tools.appendChild(addBtn);

  const wrap = document.createElement('div');
  wrap.className = 'table-wrap';
  const table = document.createElement('table');
  const head = document.createElement('thead');
  const headRow = document.createElement('tr');
  tableDef.columns.forEach((col) => {
    const th = document.createElement('th');
    th.textContent = col.label;
    headRow.appendChild(th);
  });
  const action = document.createElement('th');
  action.textContent = 'Actions';
  headRow.appendChild(action);
  head.appendChild(headRow);
  table.appendChild(head);

  const body = document.createElement('tbody');
  body.dataset.tablePath = `${sectionId}.${tableDef.key}`;
  table.appendChild(body);
  wrap.appendChild(table);
  block.append(tools, wrap);

  const starterRows = tableDef.starterRows?.length ? tableDef.starterRows : [];
  starterRows.forEach((rowData) => addTableRow(body, sectionId, tableDef, rowData));

  const existing = body.children.length;
  const required = Math.max(tableDef.minRows || 1, existing || 0);
  for (let i = existing; i < required; i += 1) addTableRow(body, sectionId, tableDef);

  addBtn.addEventListener('click', () => addTableRow(body, sectionId, tableDef));
  return block;
}

function addTableRow(tbody, sectionId, tableDef, seed = {}) {
  const row = document.createElement('tr');
  const rowIndex = tbody.children.length;

  tableDef.columns.forEach((col) => {
    const td = document.createElement('td');
    const input = document.createElement('input');
    input.type = col.type || 'text';
    if (col.type === 'number') input.step = 'any';
    input.dataset.path = `${sectionId}.${tableDef.key}[${rowIndex}].${col.key}`;
    input.name = input.dataset.path;
    if (seed[col.key] !== undefined) input.value = seed[col.key];
    td.appendChild(input);
    row.appendChild(td);
  });

  const actionTd = document.createElement('td');
  actionTd.className = 'row-actions';
  const del = document.createElement('button');
  del.type = 'button';
  del.className = 'btn-link';
  del.textContent = 'Remove';
  del.addEventListener('click', () => {
    row.remove();
    reindexTableRows(tbody, sectionId, tableDef);
  });
  actionTd.appendChild(del);
  row.appendChild(actionTd);
  tbody.appendChild(row);
}

function reindexTableRows(tbody, sectionId, tableDef) {
  Array.from(tbody.children).forEach((row, idx) => {
    tableDef.columns.forEach((col, colIdx) => {
      const input = row.children[colIdx].querySelector('input');
      const path = `${sectionId}.${tableDef.key}[${idx}].${col.key}`;
      input.dataset.path = path;
      input.name = path;
    });
  });
}

function collectFormData(formEl) {
  const data = {};
  formEl.querySelectorAll('[data-path]').forEach((el) => {
    const raw = el.value;
    if (raw === '') return;
    const val = el.type === 'number' ? Number(raw) : raw;
    setDeep(data, el.dataset.path, val);
  });
  return data;
}

function fillFormData(formEl, data) {
  formEl.querySelectorAll('[data-path]').forEach((el) => {
    const value = getDeep(data, el.dataset.path);
    if (value !== undefined && value !== null) el.value = value;
  });
}

function hydrateTableRows(formEl, schema, data) {
  schema.forEach((section) => {
    (section.tables || []).forEach((tableDef) => {
      const rows = getDeep(data, `${section.id}.${tableDef.key}`);
      if (!Array.isArray(rows)) return;
      const tbody = formEl.querySelector(`[data-table-path=\"${section.id}.${tableDef.key}\"]`);
      if (!tbody) return;
      const current = tbody.children.length;
      for (let i = current; i < rows.length; i += 1) addTableRow(tbody, section.id, tableDef);
      reindexTableRows(tbody, section.id, tableDef);
    });
  });
}

function clearForm(formEl) {
  formEl.querySelectorAll('input, textarea, select').forEach((el) => {
    if (el.tagName === 'SELECT') el.selectedIndex = 0;
    else el.value = '';
  });
}

function setPlanFieldFlag(path, flag) {
  const input = document.querySelector(`#plan-form [data-path=\"${path}\"]`);
  if (!input) return;
  const field = input.closest('.field, .field-medium, .field-wide, .field-full');
  if (!field) return;
  field.classList.remove('flag-red', 'flag-yellow', 'flag-green');
  field.classList.add(flag);
}

function initializePlanFormFromCentral() {
  const store = getCentralStore();
  const planForm = document.getElementById('plan-form');
  clearForm(planForm);
  Object.keys(planPrefillBaseline).forEach((k) => delete planPrefillBaseline[k]);

  planForm.querySelectorAll('.field, .field-medium, .field-wide, .field-full').forEach((field) => {
    field.classList.remove('flag-red', 'flag-yellow', 'flag-green');
    field.classList.add('flag-red');
  });

  fillFormData(planForm, store.planOrder || {});

  Object.keys(planPrefillMap).forEach((path) => {
    const input = planForm.querySelector(`[data-path=\"${path}\"]`);
    if (!input) return;
    const existingPlan = getDeep(store.planOrder || {}, path);
    if (existingPlan !== undefined && existingPlan !== null && existingPlan !== '') {
      input.value = existingPlan;
      setPlanFieldFlag(path, 'flag-green');
      return;
    }
    const mapped = planPrefillMap[path](store.pqTemplate || {});
    if (mapped !== undefined && mapped !== null && mapped !== '') {
      input.value = mapped;
      planPrefillBaseline[path] = String(mapped);
      setPlanFieldFlag(path, 'flag-yellow');
    } else {
      setPlanFieldFlag(path, 'flag-red');
    }
  });

  planForm.querySelectorAll('[data-path]').forEach((input) => {
    const path = input.dataset.path;
    if (!planPrefillMap[path]) {
      if (input.value !== '') setPlanFieldFlag(path, 'flag-green');
      else setPlanFieldFlag(path, 'flag-red');
    }
  });
}

function onPlanInput(event) {
  const input = event.target;
  if (!input?.dataset?.path) return;
  const path = input.dataset.path;
  const value = String(input.value || '');
  const baseline = planPrefillBaseline[path];
  if (value === '') {
    setPlanFieldFlag(path, 'flag-red');
    return;
  }
  if (baseline !== undefined && value === baseline) {
    setPlanFieldFlag(path, 'flag-yellow');
    return;
  }
  setPlanFieldFlag(path, 'flag-green');
}

function renderSchema(hostId, schema, formType) {
  const host = document.getElementById(hostId);
  schema.forEach((section, index) => {
    const isOpen = formType === 'plan' ? index < 3 : index === 0;
    host.appendChild(buildSection(section, isOpen, formType));
  });
}

function setActiveScreen(screenId) {
  currentScreen = screenId;
  const screens = ['pq', 'pq-review', 'plan', 'client'];
  screens.forEach((id) => {
    const node = document.getElementById(`screen-${id}`);
    if (!node) return;
    const show = id === screenId;
    node.classList.toggle('is-active', show);
    node.setAttribute('aria-hidden', String(!show));
  });
  updatePipeline();
}

function setPipelineStage(stageId, stateClass, statusText, metaText, disabled) {
  const stage = document.getElementById(`step-btn-${stageId}`);
  const status = document.getElementById(`stage-status-${stageId}`);
  const meta = document.getElementById(`stage-meta-${stageId}`);
  if (!stage || !status || !meta) return;
  stage.classList.remove('is-current', 'is-complete', 'is-ready', 'is-locked');
  stage.classList.add(stateClass);
  stage.disabled = Boolean(disabled);
  status.textContent = statusText;
  meta.textContent = metaText;
}

function updatePipeline() {
  const store = getCentralStore();
  const pqSubmitted = Boolean(store.workflow.pqSubmittedAt);
  const planSubmitted = Boolean(store.workflow.planSubmittedAt);
  const clientUpdated = Boolean(store.workflow.clientUpdatedAt);

  const pqCurrent = currentScreen === 'pq' || currentScreen === 'pq-review';
  setPipelineStage(
    'pq',
    pqCurrent ? 'is-current' : (pqSubmitted ? 'is-complete' : 'is-ready'),
    pqSubmitted ? 'Complete' : 'Open',
    pqSubmitted ? `Submitted ${formatTimestamp(store.workflow.pqSubmittedAt)}` : 'Intake not submitted',
    false
  );

  const planClass = !pqSubmitted ? 'is-locked' : (currentScreen === 'plan' ? 'is-current' : (planSubmitted ? 'is-complete' : 'is-ready'));
  const planStatus = !pqSubmitted ? 'Waiting on PQ' : (planSubmitted ? 'Complete' : 'Open');
  const planMeta = planSubmitted ? `Submitted ${formatTimestamp(store.workflow.planSubmittedAt)}` : 'Plan not submitted';
  setPipelineStage('plan', planClass, planStatus, planMeta, !pqSubmitted);

  const clientClass = !planSubmitted ? 'is-locked' : (currentScreen === 'client' ? 'is-current' : (clientUpdated ? 'is-complete' : 'is-ready'));
  const clientStatus = !planSubmitted ? 'Waiting on Plan' : (clientUpdated ? 'Complete' : 'Open');
  const clientMeta = clientUpdated ? `Updated ${formatTimestamp(store.workflow.clientUpdatedAt)}` : 'No updates submitted';
  setPipelineStage('client', clientClass, clientStatus, clientMeta, !planSubmitted);
}

function refreshDataConsole() {
  document.getElementById('console-json').textContent = JSON.stringify(getCentralStore(), null, 2);
}

function summarizePqForReview() {
  const store = getCentralStore();
  const pq = store.pqTemplate || {};
  const collected = [];
  const missing = [];

  pqSchema.forEach((section) => {
    (section.fields || []).forEach((field) => {
      const path = `${section.id}.${field.key}`;
      const value = getDeep(pq, path);
      const label = `${section.title} -> ${field.label}`;
      if (hasData(value)) collected.push(`${label}: ${String(value).slice(0, 80)}`);
      else missing.push(label);
    });
    (section.tables || []).forEach((tableDef) => {
      const rows = getDeep(pq, `${section.id}.${tableDef.key}`);
      const label = `${section.title} -> ${tableDef.title}`;
      if (Array.isArray(rows) && rows.length > 0 && rows.some((r) => hasData(r))) {
        collected.push(`${label}: ${rows.length} row(s) entered`);
      } else {
        missing.push(`${label}: no rows entered`);
      }
    });
  });

  const collectedHost = document.getElementById('pq-collected-list');
  const missingHost = document.getElementById('pq-missing-list');
  collectedHost.innerHTML = '';
  missingHost.innerHTML = '';

  if (!collected.length) {
    const li = document.createElement('li');
    li.textContent = 'No fields were submitted.';
    collectedHost.appendChild(li);
  } else {
    collected.forEach((text) => {
      const li = document.createElement('li');
      li.textContent = text;
      collectedHost.appendChild(li);
    });
  }

  if (!missing.length) {
    const li = document.createElement('li');
    li.textContent = 'All tracked PQ fields were collected.';
    missingHost.appendChild(li);
  } else {
    missing.forEach((text) => {
      const li = document.createElement('li');
      li.textContent = text;
      missingHost.appendChild(li);
    });
  }
}

function flattenLeaves(obj, prefix = '') {
  if (obj == null || typeof obj !== 'object' || Array.isArray(obj)) return [{ path: prefix, value: obj }];
  const out = [];
  Object.entries(obj).forEach(([key, value]) => {
    const next = prefix ? `${prefix}.${key}` : key;
    if (value != null && typeof value === 'object' && !Array.isArray(value)) out.push(...flattenLeaves(value, next));
    else if (Array.isArray(value)) {
      if (!value.length) out.push({ path: next, value: [] });
      value.forEach((row, idx) => {
        if (row != null && typeof row === 'object') out.push(...flattenLeaves(row, `${next}[${idx}]`));
        else out.push({ path: `${next}[${idx}]`, value: row });
      });
    } else out.push({ path: next, value });
  });
  return out;
}

function getFriendlyClientLabel(path) {
  if (CLIENT_LABEL_OVERRIDES[path]) return CLIENT_LABEL_OVERRIDES[path];

  const sourceKey = path.startsWith('pqTemplate.') ? 'pqTemplate' : (path.startsWith('planOrder.') ? 'planOrder' : null);
  if (!sourceKey) return path;

  const schema = sourceKey === 'pqTemplate' ? pqSchema : planSchema;
  const tokens = parsePath(path.replace(`${sourceKey}.`, ''));
  const [sectionId, second, third, fourth] = tokens;
  const section = schema.find((s) => s.id === sectionId);
  if (!section) return path;

  if (tokens.length === 2) {
    const field = (section.fields || []).find((f) => f.key === second);
    if (field) return field.label;
  }

  if (tokens.length >= 4 && /^\d+$/.test(third)) {
    const table = (section.tables || []).find((t) => t.key === second);
    if (!table) return path;
    const column = (table.columns || []).find((c) => c.key === fourth);
    const rowNumber = Number(third) + 1;
    if (column) return `${column.label} (Row ${rowNumber})`;
    return `${table.title} (Row ${rowNumber})`;
  }

  return section.title;
}

function renderClientEditor() {
  const store = getCentralStore();
  const host = document.getElementById('client-editor-grid');
  host.innerHTML = '';
  const combined = {
    pqTemplate: deepClone(store.pqTemplate || {}),
    planOrder: deepClone(store.planOrder || {})
  };

  const leaves = flattenLeaves(combined).filter((entry) => entry.path);
  if (!leaves.length) {
    const note = document.createElement('p');
    note.className = 'section-note';
    note.textContent = 'No submitted data found yet.';
    host.appendChild(note);
    return;
  }

  leaves.forEach((entry) => {
    const field = document.createElement('label');
    field.className = 'field-wide client-field';
    const row = document.createElement('div');
    row.className = 'label-row';
    const label = document.createElement('span');
    label.className = 'label';
    label.textContent = getFriendlyClientLabel(entry.path);
    row.appendChild(label);

    const input = document.createElement('input');
    input.type = typeof entry.value === 'number' ? 'number' : 'text';
    if (input.type === 'number') input.step = 'any';
    input.value = entry.value ?? '';
    input.dataset.clientPath = entry.path;
    input.dataset.baseValue = entry.value ?? '';
    input.dataset.valueType = typeof entry.value;
    input.addEventListener('input', () => {
      const changed = String(input.value) !== String(input.dataset.baseValue);
      field.classList.toggle('pending', changed);
      if (changed) field.classList.remove('committed');
    });
    field.append(row, input);
    host.appendChild(field);
  });
}

function commitClientEdits() {
  const store = getCentralStore();
  const fields = document.querySelectorAll('#client-editor-grid [data-client-path]');
  fields.forEach((input) => {
    const path = input.dataset.clientPath;
    const valueType = input.dataset.valueType;
    let val = input.value;
    if (valueType === 'number' && val !== '') val = Number(val);
    setDeep(store, path, val);
    input.dataset.baseValue = String(val ?? '');
    const container = input.closest('.client-field');
    if (container) {
      container.classList.remove('pending');
      container.classList.add('committed');
    }
  });
  store.workflow = store.workflow || {};
  store.workflow.clientUpdatedAt = new Date().toISOString();
  saveCentralStore(store);
}

function bindPqEvents() {
  const pqForm = document.getElementById('pq-form');
  pqForm.addEventListener('submit', (event) => {
    event.preventDefault();
    const payload = collectFormData(pqForm);
    const store = getCentralStore();
    store.pqTemplate = payload;
    store.workflow = store.workflow || {};
    store.workflow.pqSubmittedAt = new Date().toISOString();
    saveCentralStore(store);
    summarizePqForReview();
    initializePlanFormFromCentral();
    setActiveScreen('pq-review');
  });

  document.getElementById('pq-reset-btn').addEventListener('click', () => {
    clearForm(pqForm);
  });

  document.getElementById('back-to-pq-edit-btn').addEventListener('click', () => setActiveScreen('pq'));
  document.getElementById('continue-to-plan-btn').addEventListener('click', () => {
    const store = getCentralStore();
    if (!store.workflow?.pqSubmittedAt) {
      alert('Submit Discovery Intake before moving to Planning Build.');
      return;
    }
    initializePlanFormFromCentral();
    setActiveScreen('plan');
  });
}

function bindPlanEvents() {
  const planForm = document.getElementById('plan-form');
  planForm.addEventListener('input', onPlanInput);
  planForm.addEventListener('submit', (event) => {
    event.preventDefault();
    const store = getCentralStore();
    if (!store.workflow?.pqSubmittedAt) {
      alert('Submit Discovery Intake first.');
      setActiveScreen('pq');
      return;
    }
    store.planOrder = collectFormData(planForm);
    store.workflow = store.workflow || {};
    store.workflow.planSubmittedAt = new Date().toISOString();
    saveCentralStore(store);
    renderClientEditor();
    setActiveScreen('client');
  });

  document.getElementById('back-to-pq-review-btn').addEventListener('click', () => {
    summarizePqForReview();
    setActiveScreen('pq-review');
  });
}

function bindClientEvents() {
  document.getElementById('back-to-plan-btn').addEventListener('click', () => setActiveScreen('plan'));
  document.getElementById('refresh-client-sheet-btn').addEventListener('click', renderClientEditor);
  document.getElementById('client-update-btn').addEventListener('click', commitClientEdits);
}

function bindPipelineNavigation() {
  document.getElementById('step-btn-pq').addEventListener('click', () => {
    const store = getCentralStore();
    setActiveScreen(store.workflow?.pqSubmittedAt ? 'pq-review' : 'pq');
  });

  document.getElementById('step-btn-plan').addEventListener('click', () => {
    const store = getCentralStore();
    if (!store.workflow?.pqSubmittedAt) {
      alert('Submit Discovery Intake to unlock Planning Build.');
      return;
    }
    initializePlanFormFromCentral();
    setActiveScreen('plan');
  });

  document.getElementById('step-btn-client').addEventListener('click', () => {
    const store = getCentralStore();
    if (!store.workflow?.planSubmittedAt) {
      alert('Submit Plan Order Form to unlock Client Confirmation.');
      return;
    }
    renderClientEditor();
    setActiveScreen('client');
  });
}

function bindDataConsole() {
  document.getElementById('reset-demo-btn').addEventListener('click', () => {
    if (!confirm('Reset the full demo and clear all stored data?')) return;
    const keys = [CENTRAL_STORE_KEY, 'pq-template-draft-v1', 'plan-order-draft-v1', 'client-info-sheet-snapshot-v1', 'pq-template-updated-at', 'plan-order-updated-at', 'client-info-updated-at'];
    keys.forEach((key) => localStorage.removeItem(key));
    location.reload();
  });

  const consoleEl = document.getElementById('data-console');
  document.getElementById('open-console-btn').addEventListener('click', () => {
    refreshDataConsole();
    consoleEl.classList.add('is-open');
    consoleEl.setAttribute('aria-hidden', 'false');
  });
  document.getElementById('close-console-btn').addEventListener('click', () => {
    consoleEl.classList.remove('is-open');
    consoleEl.setAttribute('aria-hidden', 'true');
  });
}

function hydrateFromCentral() {
  const store = getCentralStore();
  const pqForm = document.getElementById('pq-form');
  if (hasData(store.pqTemplate)) {
    hydrateTableRows(pqForm, pqSchema, store.pqTemplate);
    fillFormData(pqForm, store.pqTemplate);
    summarizePqForReview();
  }
  initializePlanFormFromCentral();
  renderClientEditor();
}

function init() {
  renderSchema('pq-form-sections', pqSchema, 'pq');
  renderSchema('plan-form-sections', planSchema, 'plan');

  bindPipelineNavigation();
  bindPqEvents();
  bindPlanEvents();
  bindClientEvents();
  bindDataConsole();
  hydrateFromCentral();
  refreshDataConsole();

  const store = getCentralStore();
  if (store.workflow?.planSubmittedAt) setActiveScreen('client');
  else if (store.workflow?.pqSubmittedAt) setActiveScreen('pq-review');
  else setActiveScreen('pq');
}

init();
