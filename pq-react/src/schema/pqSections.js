// Ported verbatim from legacy app.js `pqSections`.
// Sections marked `customRenderer` (employment, assets, liabilities, income,
// taxesExpenses) need bespoke React components — see TODOs in PQScreen.

export const pqSections = [
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
        ],
      },
    ],
    afterConditional: [
      {
        id: 'children-block',
        toggleLabel: '+ Add Children / Grandchildren',
        hideLabel: '- Remove Children Section',
        flag: 'hasChildren',
        tables: [
          { key: 'children', title: 'Children', columns: [
            { key: 'name', label: 'Name' },
            { key: 'age', label: 'Age', type: 'number' },
          ] },
          { key: 'grandchildren', title: 'Grandchildren', columns: [
            { key: 'name', label: 'Name' },
            { key: 'age', label: 'Age', type: 'number' },
          ] },
        ],
        summaryFields: [
          { key: 'numChildren', label: '# of Children', type: 'number', width: 'field' },
          { key: 'numGrandchildren', label: '# of Grandchildren', type: 'number', width: 'field' },
        ],
      },
    ],
    footerFields: [
      { key: 'additionalFamilyInfo', label: 'Additional Family Information', type: 'textarea' },
    ],
  },
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
    ],
  },
  {
    id: 'employment',
    colorTheme: 'violet',
    title: 'Employment',
    customRenderer: 'renderEmploymentSection',
  },
  {
    id: 'goals',
    colorTheme: 'amber',
    title: 'Estate Plan',
    fields: [
      { key: 'estatePlan', label: 'Estate Plan', type: 'multiselect', options: ['Trust', 'Will', 'FPOA', 'MPOA'], width: 'medium' },
      { key: 'estatePlanYear', label: 'Year Established / Updated', width: 'medium' },
    ],
  },
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
        ],
      },
    ],
  },
  { id: 'assets', colorTheme: 'indigo', title: 'Assets', customRenderer: 'renderAssetsSection' },
  { id: 'liabilities', colorTheme: 'rose', title: 'Liabilities', customRenderer: 'renderLiabilitiesSection' },
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
        ],
      },
    ],
  },
  { id: 'income', colorTheme: 'green', title: 'Income', customRenderer: 'renderIncomeSection' },
  { id: 'taxesExpenses', colorTheme: 'orange', title: 'Taxes & Expenses', customRenderer: 'renderTaxesExpensesSection' },
];
