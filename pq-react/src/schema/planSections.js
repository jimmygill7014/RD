// Ported from legacy app.js `planSections`. Plan-mode fields persist as a
// flat map under data.planOrder.{key} (not nested per section).

const PLAN_ROR_OPTS = [
  '', '100/0', '90/10', '80/20', '70/30', '60/40', '50/50',
  '40/60', '30/70', '20/80', '10/90', '0/100',
];

export const planSections = [
  { id: 'documents', title: 'Documents', color: 'blue', fields: [
    { key: 'investmentStatementOnFile', label: 'Investment Statement on File?',          type: 'checkbox' },
    { key: 'taxReturnOnFile',           label: 'Tax Return on File?',                    type: 'checkbox' },
    { key: 'planningLifeRequested',     label: 'Planning Life Requested?',               type: 'checkbox' },
    { key: 'documentsEmailDate',        label: 'Date E-mail Request for Documents Sent', type: 'date' },
  ]},
  { id: 'baseFacts', title: 'Base Facts', color: 'teal', fields: [
    { key: 'clientRetirementDate', label: 'Client Retirement Date',     type: 'date' },
    { key: 'spouseRetirementDate', label: 'Spouse Retirement Date',     type: 'date' },
    { key: 'eMoneyRorPre',         label: 'eMoney RoR Pre-Retirement',  type: 'select', options: PLAN_ROR_OPTS },
    { key: 'eMoneyRorPost',        label: 'eMoney RoR Post-Retirement', type: 'select', options: PLAN_ROR_OPTS },
    { key: 'baseFactsNotes',       label: 'Base Facts Notes',           type: 'textarea' },
  ]},
  { id: 'observations', title: 'Observations', color: 'amber', fields: [
    { key: 'currentCashFlow',    label: 'Current Cash Flow',                type: 'select', options: ['', 'Surplus', 'Deficit', 'Break-Even'] },
    { key: 'maxingEmployerPlan', label: 'Maxing Employer Retirement Plan?', type: 'select', options: ['', 'Yes', 'No', 'N/A'] },
    { key: 'excessCashGoing',    label: 'Where is Excess Cash Going?',      type: 'select', options: ['', 'Cash', 'Qualified', 'Roth', 'Non-Qualified', 'Other'] },
    { key: 'extraCashReserves',  label: 'How Much Extra Cash Reserves?',    type: 'text' },
    { key: 'onTrackToRetire',    label: 'Are They on Track to Retire?',     type: 'select', options: ['', 'Yes', 'No', 'Unsure'] },
    { key: 'observationNotes',   label: 'Observation Notes',                type: 'textarea' },
  ]},
  { id: 'taxes', title: 'Taxes', color: 'violet', fields: [
    { key: 'filingStatus',         label: 'Filing Status This Year',                  type: 'select', options: ['', 'Married Filing Jointly', 'Married Filing Separately', 'Single', 'Head of Household'] },
    { key: 'taxReturnDifferences', label: 'Differences From Prior Tax Return',        type: 'textarea' },
    { key: 'taxBracket',           label: 'What Tax Bracket Are They in Now?',        type: 'textarea' },
    { key: 'charitablyInclined',   label: 'Charitably Inclined?',                     type: 'textarea' },
    { key: 'taxesBeforeRMDs',      label: 'Taxes: Higher or Lower Before RMDs Start', type: 'select', options: ['', 'Higher', 'Lower', 'Same'] },
    { key: 'taxesAfterRMDs',       label: 'Taxes: Higher or Lower After RMDs Start',  type: 'select', options: ['', 'Higher', 'Lower', 'Same'] },
    { key: 'highLevelStrategies',  label: 'High Level Strategies For This Year',      type: 'textarea' },
    { key: 'unrealizedCapGains',   label: 'Unrealized Capital Gains?',                type: 'textarea' },
  ]},
  { id: 'strategiesCF', title: 'Strategies — Cash Flow & Retirement', color: 'emerald', fields: [
    { key: 'cashReservesStrategy',    label: 'Cash Reserves Strategy',     type: 'textarea' },
    { key: 'ssClaimingStrategy',      label: 'SS Claiming Strategy',       type: 'textarea' },
    { key: 'otherCashFlowStrategies', label: 'Other Cash Flow Strategies', type: 'textarea' },
  ]},
  { id: 'strategiesTax', title: 'Strategies — Taxes', color: 'indigo', fields: [
    { key: 'rothConversionStrategy',    label: 'Roth Conversion Strategy',           type: 'textarea' },
    { key: 'charitableStrategies',      label: 'Charitable Strategies',              type: 'textarea' },
    { key: 'capitalGains',              label: 'Capital Gains',                      type: 'textarea' },
    { key: 'capitalGainsWorksheet',     label: 'Capital Gains Worksheet Requested?', type: 'select', options: ['', 'Yes', 'No', 'N/A'] },
    { key: 'employerPlanContributions', label: 'Employer Plan Contributions',        type: 'textarea' },
    { key: 'rothIraContributions',      label: 'Roth / IRA Contributions',           type: 'textarea' },
    { key: 'otherTaxStrategies',        label: 'Other Tax Strategies',               type: 'textarea' },
  ]},
  { id: 'strategiesInvest', title: 'Strategies — Investment Planning', color: 'sky', fields: [
    { key: 'pureIPS',             label: 'Pure IPS',                              type: 'text' },
    { key: 'outside401kRequest',  label: 'Outside 401(k) Request',                type: 'select', options: ['', 'Requested', 'Not Requested', 'Link Outside Retirement Plan to Pure'] },
    { key: 'annuityStrategy',     label: 'Annuity Strategy',                      type: 'textarea' },
    { key: 'morningstarEntered',  label: 'Morningstar Entered During Assessment?', type: 'select', options: ['', 'Yes', 'No'] },
    { key: 'otherInvestRequests', label: 'Other Investment Requests',             type: 'textarea' },
  ]},
  { id: 'strategiesInsurance', title: 'Strategies — Insurance & Risk', color: 'rose', fields: [
    { key: 'lifeInsurance', label: 'Life Insurance',   type: 'textarea' },
    { key: 'pAndC',         label: 'P&C',              type: 'textarea' },
    { key: 'ltcDisability', label: 'LTC / Disability', type: 'textarea' },
  ]},
  { id: 'strategiesEstate', title: 'Strategies — Estate Planning', color: 'orange', fields: [
    { key: 'estatePlanStatus',      label: 'Estate Plan',             type: 'select', options: ['', 'Has Current Estate Plan', 'Needs Updated Estate Plan', 'Needs Estate Plan', 'No Estate Plan', 'No Estate Plan Needed'] },
    { key: 'estatePlanningDetails', label: 'Estate Planning Details', type: 'textarea' },
  ]},
  { id: 'altScenarios', title: 'Alternative Scenarios', color: 'green', fields: [
    { key: 'alternativeScenarios', label: 'Alternative Scenarios', type: 'select', options: ['', 'Yes', 'No'] },
  ]},
];
