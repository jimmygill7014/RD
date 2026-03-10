const PQ_DRAFT_KEY = 'pq-template-draft-v1';
const PLAN_DRAFT_KEY = 'plan-order-draft-v1';
const CLIENT_SNAPSHOT_KEY = 'client-info-sheet-snapshot-v1';
const PQ_UPDATED_KEY = 'pq-template-updated-at';
const PLAN_UPDATED_KEY = 'plan-order-updated-at';
const CLIENT_UPDATED_KEY = 'client-info-updated-at';

let currentScreen = 'pq';

const pqSchema = [
  {
    id: 'family',
    title: 'Family Information',
    note: 'Core household and referral details from the PQ template.',
    fields: [
      { key: 'yourName', label: 'Your Name', width: 'medium' },
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

    const labelRow = document.createElement('div');
    labelRow.className = 'label-row';

    const title = document.createElement('span');
    title.className = 'label';
    title.textContent = field.label;

    labelRow.appendChild(title);

    if (formType === 'plan') {
      const tag = document.createElement('span');
      tag.className = 'prefill-tag';
      tag.textContent = 'Prefilled';
      labelRow.appendChild(tag);
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
      (field.options || []).forEach((optionValue) => {
        const option = document.createElement('option');
        option.value = optionValue;
        option.textContent = optionValue;
        input.appendChild(option);
      });
    } else {
      input = document.createElement('input');
      input.type = field.type || 'text';
      if (field.type === 'number') input.step = 'any';
    }

    input.name = `${section.id}.${field.key}`;
    input.dataset.path = `${section.id}.${field.key}`;

    label.append(labelRow, input);
    grid.appendChild(label);
  });

  return grid;
}

function buildTable(sectionId, tableDef) {
  const block = document.createElement('div');
  const toolbar = document.createElement('div');
  toolbar.className = 'table-tools';
  toolbar.innerHTML = `<div class=\"table-title\">${tableDef.title}</div>`;

  const addBtn = document.createElement('button');
  addBtn.type = 'button';
  addBtn.className = 'btn-link';
  addBtn.textContent = '+ Add Row';
  toolbar.appendChild(addBtn);

  const wrap = document.createElement('div');
  wrap.className = 'table-wrap';

  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  tableDef.columns.forEach((col) => {
    const th = document.createElement('th');
    th.textContent = col.label;
    headRow.appendChild(th);
  });
  const actionTh = document.createElement('th');
  actionTh.textContent = 'Actions';
  headRow.appendChild(actionTh);

  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  tbody.dataset.tablePath = `${sectionId}.${tableDef.key}`;
  table.appendChild(tbody);

  wrap.appendChild(table);
  block.append(toolbar, wrap);

  const starterRows = tableDef.starterRows?.length ? tableDef.starterRows : [];
  if (starterRows.length) starterRows.forEach((rowData) => addTableRow(tbody, sectionId, tableDef, rowData));

  const existingRows = tbody.children.length;
  const requiredRows = Math.max(tableDef.minRows || 1, existingRows || 0);
  for (let i = existingRows; i < requiredRows; i += 1) addTableRow(tbody, sectionId, tableDef);

  addBtn.addEventListener('click', () => addTableRow(tbody, sectionId, tableDef));

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

  const actionCell = document.createElement('td');
  actionCell.className = 'row-actions';
  const deleteBtn = document.createElement('button');
  deleteBtn.type = 'button';
  deleteBtn.className = 'btn-link';
  deleteBtn.textContent = 'Remove';
  deleteBtn.addEventListener('click', () => {
    row.remove();
    reindexTableRows(tbody, sectionId, tableDef);
    updatePreview(document.getElementById('pq-form'), 'pq-json-preview');
  });
  actionCell.appendChild(deleteBtn);
  row.appendChild(actionCell);

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
    const ref = getDeep(data, el.dataset.path);
    if (ref !== undefined && ref !== null) el.value = ref;
  });
}

function hydrateTableRows(formEl, schema, data) {
  schema.forEach((section) => {
    (section.tables || []).forEach((tableDef) => {
      const path = `${section.id}.${tableDef.key}`;
      const rows = getDeep(data, path);
      if (!Array.isArray(rows)) return;

      const tbody = formEl.querySelector(`[data-table-path="${path}"]`);
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

function updatePreview(formEl, previewId) {
  const payload = collectFormData(formEl);
  document.getElementById(previewId).textContent = JSON.stringify(payload, null, 2);
}

function markPrefilled(planFormEl, planPath, isPrefilled) {
  const input = planFormEl.querySelector(`[data-path="${planPath}"]`);
  if (!input) return;
  const field = input.closest('.field, .field-medium, .field-wide, .field-full');
  if (!field) return;
  if (isPrefilled) field.classList.add('is-prefilled');
  else field.classList.remove('is-prefilled');
}

function applyPrefillFromPq(planFormEl, pqData, force = false) {
  if (!pqData || Object.keys(pqData).length === 0) return;

  Object.entries(planPrefillMap).forEach(([planPath, resolver]) => {
    const mappedValue = resolver(pqData);
    if (mappedValue === undefined || mappedValue === null || mappedValue === '') return;

    const input = planFormEl.querySelector(`[data-path="${planPath}"]`);
    if (!input) return;

    const currentValue = input.value;
    const shouldApply = force || currentValue === '';
    if (shouldApply) {
      input.value = mappedValue;
    }
    if (shouldApply || String(currentValue) === String(mappedValue)) markPrefilled(planFormEl, planPath, true);
  });
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
  const pqScreen = document.getElementById('screen-pq');
  const planScreen = document.getElementById('screen-plan');
  const clientScreen = document.getElementById('screen-client');

  const showPq = screenId === 'pq';
  const showPlan = screenId === 'plan';
  const showClient = screenId === 'client';

  pqScreen.classList.toggle('is-active', showPq);
  pqScreen.setAttribute('aria-hidden', String(!showPq));
  planScreen.classList.toggle('is-active', showPlan);
  planScreen.setAttribute('aria-hidden', String(!showPlan));
  clientScreen.classList.toggle('is-active', showClient);
  clientScreen.setAttribute('aria-hidden', String(!showClient));
  updatePipeline();
}

function persistPq() {
  const pqPayload = collectFormData(document.getElementById('pq-form'));
  localStorage.setItem(PQ_DRAFT_KEY, JSON.stringify(pqPayload));
  localStorage.setItem(PQ_UPDATED_KEY, String(Date.now()));
  updatePipeline();
  return pqPayload;
}

function persistPlan() {
  const planPayload = collectFormData(document.getElementById('plan-form'));
  localStorage.setItem(PLAN_DRAFT_KEY, JSON.stringify(planPayload));
  localStorage.setItem(PLAN_UPDATED_KEY, String(Date.now()));
  updatePipeline();
  return planPayload;
}

function buildCombinedPayload() {
  return {
    pq: collectFormData(document.getElementById('pq-form')),
    planOrder: collectFormData(document.getElementById('plan-form'))
  };
}

function formatReviewValue(value) {
  if (value === undefined || value === null || value === '') return '<span class="empty-note">Not provided</span>';
  if (Array.isArray(value)) return value.length ? `${value.length} item(s)` : '<span class="empty-note">Not provided</span>';
  if (typeof value === 'object') return JSON.stringify(value);
  return String(value);
}

function hasData(value) {
  if (value === null || value === undefined) return false;
  if (Array.isArray(value)) return value.some((item) => hasData(item));
  if (typeof value === 'object') return Object.values(value).some((child) => hasData(child));
  return String(value).trim() !== '';
}

function formatTimestamp(timestampValue) {
  if (!timestampValue) return '';
  const parsed = Number(timestampValue);
  const date = Number.isNaN(parsed) ? new Date(timestampValue) : new Date(parsed);
  if (Number.isNaN(date.getTime())) return '';
  return date.toLocaleString([], { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' });
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
  const pqData = collectFormData(document.getElementById('pq-form'));
  const planData = collectFormData(document.getElementById('plan-form'));
  const hasPqData = hasData(pqData);
  const hasPlanData = hasData(planData);
  const hasClientSnapshot = Boolean(localStorage.getItem(CLIENT_SNAPSHOT_KEY));

  const pqStamp = formatTimestamp(localStorage.getItem(PQ_UPDATED_KEY));
  const planStamp = formatTimestamp(localStorage.getItem(PLAN_UPDATED_KEY));
  const clientStamp = formatTimestamp(localStorage.getItem(CLIENT_UPDATED_KEY));

  const pqState = currentScreen === 'pq' ? 'is-current' : (hasPqData ? 'is-complete' : 'is-current');
  const pqStatus = currentScreen === 'pq' ? 'In Progress' : (hasPqData ? 'Completed' : 'In Progress');
  const pqMeta = pqStamp ? `Updated ${pqStamp}` : 'Capture initial client profile';
  setPipelineStage('pq', pqState, pqStatus, pqMeta, false);

  let planState = 'is-locked';
  let planStatus = 'Waiting on PQ';
  if (hasPqData) {
    if (currentScreen === 'plan') {
      planState = 'is-current';
      planStatus = 'In Progress';
    } else if (hasPlanData) {
      planState = 'is-complete';
      planStatus = 'Completed';
    } else {
      planState = 'is-ready';
      planStatus = 'Ready for Next Phase';
    }
  }
  const planMeta = planStamp ? `Updated ${planStamp}` : 'Advisor assumptions and analysis';
  setPipelineStage('plan', planState, planStatus, planMeta, !hasPqData);

  let clientState = 'is-locked';
  let clientStatus = 'Waiting on Plan';
  if (hasPlanData) {
    if (currentScreen === 'client') {
      clientState = 'is-current';
      clientStatus = 'In Review';
    } else if (hasClientSnapshot) {
      clientState = 'is-complete';
      clientStatus = 'Completed';
    } else {
      clientState = 'is-ready';
      clientStatus = 'Ready for Confirmation';
    }
  }
  const clientMeta = clientStamp ? `Updated ${clientStamp}` : 'Final validation summary';
  setPipelineStage('client', clientState, clientStatus, clientMeta, !hasPlanData);
}

function appendReviewSection(host, sourceLabel, sourceClass, section, sourceData) {
  const block = document.createElement('section');
  block.className = 'review-block';
  block.innerHTML = `<h3>${section.title}</h3>`;

  const grid = document.createElement('div');
  grid.className = 'review-grid';

  (section.fields || []).forEach((field) => {
    const path = `${section.id}.${field.key}`;
    const value = getDeep(sourceData, path);
    const item = document.createElement('article');
    item.className = `review-item ${sourceClass}`;
    item.innerHTML = `<div class="review-item-label">${field.label}</div><div class="review-item-value">${formatReviewValue(value)}</div><div class="review-source">${sourceLabel}</div>`;
    grid.appendChild(item);
  });

  (section.tables || []).forEach((tableDef) => {
    const rows = getDeep(sourceData, `${section.id}.${tableDef.key}`) || [];
    const item = document.createElement('article');
    item.className = `review-item ${sourceClass}`;
    const preview = Array.isArray(rows) && rows.length ? rows.slice(0, 2).map((row) => JSON.stringify(row)).join('\n') : '';
    item.innerHTML = `<div class="review-item-label">${tableDef.title}</div><div class="review-item-value">${rows.length ? `${rows.length} row(s)` : '<span class="empty-note">No rows entered</span>'}${preview ? `\n${preview}` : ''}</div><div class="review-source">${sourceLabel}</div>`;
    grid.appendChild(item);
  });

  if (!grid.children.length) return;

  block.appendChild(grid);
  host.appendChild(block);
}

function refreshClientInfoSheet() {
  const combined = buildCombinedPayload();
  const host = document.getElementById('client-review-sections');
  host.innerHTML = '';

  pqSchema.forEach((section) => appendReviewSection(host, 'From PQ Template', 'source-pq', section, combined.pq));
  planSchema.forEach((section) => appendReviewSection(host, 'From Plan Order Form', 'source-plan', section, combined.planOrder));

  document.getElementById('client-json-preview').textContent = JSON.stringify(combined, null, 2);
  updatePipeline();
  return combined;
}

function bindPqEvents() {
  const pqForm = document.getElementById('pq-form');

  pqForm.addEventListener('input', () => {
    updatePreview(pqForm, 'pq-json-preview');
    updatePipeline();
  });

  document.getElementById('pq-save-btn').addEventListener('click', () => {
    persistPq();
    updatePreview(pqForm, 'pq-json-preview');
    alert('PQ draft saved to localStorage.');
  });

  document.getElementById('pq-reset-btn').addEventListener('click', () => {
    if (!confirm('Clear all PQ values and remove saved PQ draft?')) return;
    localStorage.removeItem(PQ_DRAFT_KEY);
    localStorage.removeItem(PQ_UPDATED_KEY);
    localStorage.removeItem(PLAN_DRAFT_KEY);
    localStorage.removeItem(PLAN_UPDATED_KEY);
    localStorage.removeItem(CLIENT_SNAPSHOT_KEY);
    localStorage.removeItem(CLIENT_UPDATED_KEY);
    location.reload();
  });

  document.getElementById('go-to-plan-btn').addEventListener('click', () => {
    const pqPayload = persistPq();
    const planForm = document.getElementById('plan-form');
    applyPrefillFromPq(planForm, pqPayload, false);
    updatePreview(planForm, 'plan-json-preview');
    setActiveScreen('plan');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });
}

function bindPlanEvents() {
  const planForm = document.getElementById('plan-form');

  planForm.addEventListener('input', () => {
    updatePreview(planForm, 'plan-json-preview');
    updatePipeline();
  });

  document.getElementById('plan-save-btn').addEventListener('click', () => {
    persistPlan();
    updatePreview(planForm, 'plan-json-preview');
    alert('Plan Order draft saved to localStorage.');
  });

  document.getElementById('plan-reset-btn').addEventListener('click', () => {
    if (!confirm('Reset Plan Order Form draft fields?')) return;
    localStorage.removeItem(PLAN_DRAFT_KEY);
    localStorage.removeItem(PLAN_UPDATED_KEY);
    clearForm(planForm);
    planForm.querySelectorAll('.is-prefilled').forEach((el) => el.classList.remove('is-prefilled'));
    const pqPayload = collectFormData(document.getElementById('pq-form'));
    applyPrefillFromPq(planForm, pqPayload, true);
    updatePreview(planForm, 'plan-json-preview');
    updatePipeline();
  });

  planForm.addEventListener('submit', (event) => {
    event.preventDefault();
    const payload = persistPlan();
    updatePreview(planForm, 'plan-json-preview');
    alert('Plan Order Form submitted. Payload saved locally for prototype testing.');
    console.log('Plan Order Form payload', payload);
  });

  document.getElementById('back-to-pq-btn').addEventListener('click', () => {
    persistPlan();
    setActiveScreen('pq');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  document.getElementById('go-to-client-btn').addEventListener('click', () => {
    const pqPayload = persistPq();
    const planPayload = persistPlan();
    if (!hasData(pqPayload)) {
      alert('Complete PQ Template before opening Client Info Sheet.');
      setActiveScreen('pq');
      return;
    }
    if (!hasData(planPayload)) {
      alert('Add details on Plan Order Form before opening Client Info Sheet.');
      setActiveScreen('plan');
      return;
    }
    refreshClientInfoSheet();
    setActiveScreen('client');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });
}

function bindClientEvents() {
  document.getElementById('back-to-plan-btn').addEventListener('click', () => {
    setActiveScreen('plan');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  document.getElementById('client-refresh-btn').addEventListener('click', () => {
    refreshClientInfoSheet();
  });

  document.getElementById('client-save-btn').addEventListener('click', () => {
    const combined = refreshClientInfoSheet();
    localStorage.setItem(CLIENT_SNAPSHOT_KEY, JSON.stringify(combined));
    localStorage.setItem(CLIENT_UPDATED_KEY, String(Date.now()));
    updatePipeline();
    alert('Client Info Sheet snapshot saved to localStorage.');
  });
}

function bindStepper() {
  document.getElementById('step-btn-pq').addEventListener('click', () => {
    persistPlan();
    setActiveScreen('pq');
  });

  document.getElementById('step-btn-plan').addEventListener('click', () => {
    const pqPayload = persistPq();
    if (!hasData(pqPayload)) {
      alert('Complete at least part of PQ Template before moving to Planning Build.');
      setActiveScreen('pq');
      return;
    }
    applyPrefillFromPq(document.getElementById('plan-form'), pqPayload, false);
    setActiveScreen('plan');
    updatePreview(document.getElementById('plan-form'), 'plan-json-preview');
  });

  document.getElementById('step-btn-client').addEventListener('click', () => {
    const pqPayload = persistPq();
    const planPayload = persistPlan();
    if (!hasData(pqPayload)) {
      alert('Complete PQ Template before opening Client Info Sheet.');
      setActiveScreen('pq');
      return;
    }
    if (!hasData(planPayload)) {
      alert('Complete Plan Order Form before opening Client Info Sheet.');
      setActiveScreen('plan');
      return;
    }
    refreshClientInfoSheet();
    setActiveScreen('client');
  });
}

function hydrateDrafts() {
  const pqForm = document.getElementById('pq-form');
  const planForm = document.getElementById('plan-form');

  const pqDraft = localStorage.getItem(PQ_DRAFT_KEY);
  if (pqDraft) {
    try {
      const parsedPq = JSON.parse(pqDraft);
      hydrateTableRows(pqForm, pqSchema, parsedPq);
      fillFormData(pqForm, parsedPq);
    } catch (error) {
      console.warn('Unable to parse PQ draft', error);
    }
  }

  const planDraft = localStorage.getItem(PLAN_DRAFT_KEY);
  if (planDraft) {
    try {
      fillFormData(planForm, JSON.parse(planDraft));
    } catch (error) {
      console.warn('Unable to parse Plan draft', error);
    }
  }

  const pqPayload = collectFormData(pqForm);
  applyPrefillFromPq(planForm, pqPayload, false);

  updatePreview(pqForm, 'pq-json-preview');
  updatePreview(planForm, 'plan-json-preview');
  refreshClientInfoSheet();
}

function init() {
  renderSchema('pq-form-sections', pqSchema, 'pq');
  renderSchema('plan-form-sections', planSchema, 'plan');

  bindStepper();
  bindPqEvents();
  bindPlanEvents();
  bindClientEvents();
  hydrateDrafts();

  setActiveScreen('pq');
}

init();
