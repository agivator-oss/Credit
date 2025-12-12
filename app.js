// Static, browser-only Excel-style Credit Checklist
// - No external libraries
// - Works when opening index.html
// - Uses fetch for data/*.json when possible, with embedded fallbacks for file://

const CHECKLIST_URL = './data/checklist.credit.json';
const SAMPLE_DEAL_URL = './data/sample.deal.json';

// Embedded fallbacks (so file:// works even if fetch is blocked)
const FALLBACK_CHECKLIST = {
  sheet: 'Credit Checklist',
  version: '1.0',
  columns: ['Item/Description', 'Status', 'CHECK', 'COMMENTS'],
  rows: [
    { id: 'deal_presentation_deal_presentation', section: 'Deal Presentation', label: 'Deal Presentation', statusDefault: 'PENDING' },
    { id: 'deal_presentation_sim_sheet', section: 'Deal Presentation', label: 'Supplementary Information Memorandum Sheet', statusDefault: 'PENDING' },

    { id: 'id_current_drivers_licence_held', section: 'Driver’s License, Passport & Medicare', label: 'Current Driver’s Licence held', statusDefault: 'PENDING' },
    { id: 'id_current_passport_held', section: 'Driver’s License, Passport & Medicare', label: 'Current Passport held', statusDefault: 'PENDING' },
    { id: 'id_current_medicare_held', section: 'Driver’s License, Passport & Medicare', label: 'Current Medicare held', statusDefault: 'PENDING' },
    { id: 'id_alternate_id_100_points', section: 'Driver’s License, Passport & Medicare', label: 'Alternate acceptable ID held up to 100 points', statusDefault: 'PENDING' },

    { id: 'loi_all_signatures_initials', section: 'Letter of Intent', label: 'All signatures & Initials', statusDefault: 'PENDING' },
    { id: 'loi_numbers_match_outline_conditions', section: 'Letter of Intent', label: 'Numbers as per loan outline and special conditions', statusDefault: 'PENDING' },

    { id: 'app_all_sections_completed', section: 'Application Form', label: 'All sections fully completed', statusDefault: 'PENDING' },
    { id: 'app_business_purpose_box_ticked', section: 'Application Form', label: 'Business purpose box ticked', statusDefault: 'PENDING' },
    { id: 'app_signature_matches_id', section: 'Application Form', label: 'Does the signature match ID', statusDefault: 'PENDING' },
    { id: 'app_exit_strategy_loan_purpose_provided', section: 'Application Form', label: 'Exit Strategy & Loan Purpose provided', statusDefault: 'PENDING' },
    { id: 'app_security_offered_details_correct', section: 'Application Form', label: 'Check Security Offered details are correct and filled', statusDefault: 'PENDING' },
    { id: 'app_borrowers_al_correct', section: 'Application Form', label: 'Check Borrower’s A/L are correct and filled', statusDefault: 'PENDING' },
    { id: 'app_accountant_solicitor_broker_verification', section: 'Application Form', label: 'Accountant, solicitor and broker verification', statusDefault: 'PENDING' },
    { id: 'app_directors_match_company_file', section: 'Application Form', label: 'Do the directors match the company file', statusDefault: 'PENDING' },

    { id: 'abn_check_abn_status', section: 'ABN Search', label: 'Check ABN status', statusDefault: 'PENDING' },
    { id: 'abn_registered_for_gst_or_not', section: 'ABN Search', label: 'Registered for GST or not', statusDefault: 'PENDING' },

    { id: 'asic_adverse_information', section: 'ASIC Search', label: 'Check for any adverse information', statusDefault: 'PENDING' },

    { id: 'borrower_background_provided', section: 'Borrower', label: 'Borrower background provided', statusDefault: 'PENDING' },
    { id: 'borrower_aleras_search_completed', section: 'Borrower', label: 'Aleras search completed', statusDefault: 'PENDING' },
    { id: 'borrower_registered_company_trust', section: 'Borrower', label: 'Is the borrower a registered company/trust', statusDefault: 'PENDING' },
    { id: 'borrower_trust_deed_fully_completed', section: 'Borrower', label: 'Check trust deed is fully completed', statusDefault: 'PENDING' },
    { id: 'borrower_company_search_conducted', section: 'Borrower', label: 'Has company search been conducted', statusDefault: 'PENDING' },
    { id: 'borrower_no_adverse_on_company_file', section: 'Borrower', label: 'No adverse evident on company file', statusDefault: 'PENDING' },
    { id: 'borrower_company_file_matches_id', section: 'Borrower', label: 'Does company file match ID', statusDefault: 'PENDING' },

    { id: 'equifax_enquiry_under_archer_wealth_pty_ltd', section: 'Equifax Search', label: 'Enquiry made under Archer Wealth Pty Ltd, not Archer Wealth Brokerage', statusDefault: 'PENDING' },
    { id: 'equifax_directors_in_place_gt_6_months', section: 'Equifax Search', label: 'Have the directors been in place for > 6 months', statusDefault: 'PENDING' },
    { id: 'equifax_scores_gt_600', section: 'Equifax Search', label: 'Are all scores > 600', statusDefault: 'PENDING' },
    { id: 'equifax_clear_of_adverse', section: 'Equifax Search', label: 'Check if clear of adverse information on report', statusDefault: 'PENDING' },

    { id: 'guarantors_no_borrowers_over_65', section: 'Guarantors', label: 'NO borrowers > 65 yrs old', statusDefault: 'PENDING' },
    { id: 'guarantors_no_third_party_guarantors', section: 'Guarantors', label: 'NO 3rd party Guarantors', statusDefault: 'PENDING' },
    { id: 'guarantors_aleras_search_completed', section: 'Guarantors', label: 'Aleras search completed', statusDefault: 'PENDING' },
    { id: 'guarantors_bankruptcy_search_completed', section: 'Guarantors', label: 'Bankruptcy search completed', statusDefault: 'PENDING' },
    { id: 'guarantors_fraud_google_search_completed', section: 'Guarantors', label: 'Fraud google search completed', statusDefault: 'PENDING' },
    { id: 'guarantors_strong_als', section: 'Guarantors', label: 'Do Guarantors have strong A&Ls', statusDefault: 'PENDING' },

    { id: 'accountant_letter_registration_checked', section: 'Accountant Letter (When interest is not fully capitalised)', label: 'Check the registration of accountant', statusDefault: 'PENDING' },
    { id: 'accountant_letter_letterhead', section: 'Accountant Letter (When interest is not fully capitalised)', label: 'Check the letter is sent on company letterhead', statusDefault: 'PENDING' },
    { id: 'accountant_letter_dated_addressed_aw', section: 'Accountant Letter (When interest is not fully capitalised)', label: 'Check the letter is dated and addressed to AW', statusDefault: 'PENDING' },
    { id: 'accountant_letter_signed', section: 'Accountant Letter (When interest is not fully capitalised)', label: 'Check the letter is signed by the accountant', statusDefault: 'PENDING' },
    { id: 'accountant_letter_serviceability_calculator_completed', section: 'Accountant Letter (When interest is not fully capitalised)', label: 'Serviceability calculator completed', statusDefault: 'PENDING' },

    { id: 'ato_guarantor_tax_portals_provided', section: 'ATO Tax Portals Company & Guarantors clearance', label: 'Guarantor Tax Portals Provided', statusDefault: 'PENDING' },
    { id: 'ato_company_tax_portals_provided', section: 'ATO Tax Portals Company & Guarantors clearance', label: 'Company Tax Portals Provided', statusDefault: 'PENDING' },
    { id: 'ato_any_director_penalties', section: 'ATO Tax Portals Company & Guarantors clearance', label: 'Any Director Penalties', statusDefault: 'PENDING' },
    { id: 'ato_director_id_provided', section: 'ATO Tax Portals Company & Guarantors clearance', label: 'Director ID Provided', statusDefault: 'PENDING' },

    { id: 'securities_title_search_conducted', section: 'Securities', label: 'Has title search been conducted', statusDefault: 'PENDING' },
    { id: 'securities_no_restrictions_on_use_of_land', section: 'Securities', label: 'NO restrictions on the use of land', statusDefault: 'PENDING' },
    { id: 'securities_all_investment_properties', section: 'Securities', label: 'Are all securities investment properties', statusDefault: 'PENDING' },
    { id: 'securities_no_caveats_or_easements', section: 'Securities', label: 'NO caveats or easements on title', statusDefault: 'PENDING' },
    { id: 'securities_all_pages_initialled_executed', section: 'Securities', label: 'All pages initialled and executed', statusDefault: 'PENDING' },

    { id: 'offer_signature_matches_id', section: 'Offer Letter', label: 'Does signature match ID', statusDefault: 'PENDING' },
    { id: 'offer_commitment_fees_paid', section: 'Offer Letter', label: 'Have commitment fees been paid', statusDefault: 'PENDING' },

    { id: 'valuation_any_changes_to_deal_metrics', section: 'Valuation and Appraisals', label: 'Any Changes to Deal Metrics', statusDefault: 'PENDING' },
    { id: 'valuation_dated_within_90_days', section: 'Valuation and Appraisals', label: 'Is the valuation dated within 90 Days', statusDefault: 'PENDING' },
    { id: 'valuation_addressed_to_archer_wealth', section: 'Valuation and Appraisals', label: 'Are all Valuations addressed to Archer Wealth Pty Ltd (not ARW01)', statusDefault: 'PENDING' },
    { id: 'valuation_review_completed', section: 'Valuation and Appraisals', label: 'Has valuation review been completed', statusDefault: 'PENDING' },
    { id: 'valuation_reliability_scorecard_completed', section: 'Valuation and Appraisals', label: 'Has valuation reliability scorecard been completed', statusDefault: 'PENDING' },
    { id: 'valuation_new_build_gst_included', section: 'Valuation and Appraisals', label: '(For new builds) GST has been included in the valuation', statusDefault: 'PENDING' },

    { id: 'cos_fully_executed_obtained', section: 'Contract of sale (Purchase deals only) or (When Exit Strategy is Sale)', label: 'Fully executed document is obtained', statusDefault: 'PENDING' },
    { id: 'cos_address_correct', section: 'Contract of sale (Purchase deals only) or (When Exit Strategy is Sale)', label: 'Check address is correct', statusDefault: 'PENDING' },
    { id: 'cos_purchaser_names_signed_dated', section: 'Contract of sale (Purchase deals only) or (When Exit Strategy is Sale)', label: 'Purchaser name(s) and is signed and dated', statusDefault: 'PENDING' },
    { id: 'cos_deposit_paid_amounts', section: 'Contract of sale (Purchase deals only) or (When Exit Strategy is Sale)', label: 'Deposit paid amounts', statusDefault: 'PENDING' },

    { id: 'cos_additional_check_for_debts', section: 'Contract of sale – additional checks', label: 'Check for any debts', statusDefault: 'PENDING' },

    { id: 'refinance_mortgage_statements_6_months', section: 'Mortgage Loan Statements (Refinance Deals)', label: 'Check 6 months mortgage statements and history', statusDefault: 'PENDING' },
    { id: 'refinance_check_address_correct', section: 'Mortgage Loan Statements (Refinance Deals)', label: 'Check address is correct', statusDefault: 'PENDING' },

    { id: 'docs_executed_application_form', section: 'Document Checklist', label: 'Executed Application form', statusDefault: 'PENDING' },
    { id: 'docs_executed_offer_letter', section: 'Document Checklist', label: 'Executed Offer Letter', statusDefault: 'PENDING' },
    { id: 'docs_valid_licence_or_passport', section: 'Document Checklist', label: 'Valid Licence or Passport provided', statusDefault: 'PENDING' },
    { id: 'docs_director_id_provided', section: 'Document Checklist', label: 'Director ID Provided', statusDefault: 'PENDING' },
    { id: 'docs_contract_of_sale', section: 'Document Checklist', label: 'Contract of Sale', statusDefault: 'PENDING' },
    { id: 'docs_trust_deed', section: 'Document Checklist', label: 'Trust deed', statusDefault: 'PENDING' },
    { id: 'docs_servicing_docs', section: 'Document Checklist', label: 'Servicing docs', statusDefault: 'PENDING' },
    { id: 'docs_refinance_statements', section: 'Document Checklist', label: 'Refinance statements', statusDefault: 'PENDING' },
    { id: 'docs_rates', section: 'Document Checklist', label: 'Rates', statusDefault: 'PENDING' },
    { id: 'docs_company_search', section: 'Document Checklist', label: 'Company search', statusDefault: 'PENDING' },
    { id: 'docs_title_search', section: 'Document Checklist', label: 'Title search', statusDefault: 'PENDING' },
    { id: 'docs_credit_report', section: 'Document Checklist', label: 'Credit report', statusDefault: 'PENDING' },
    { id: 'docs_valuation_and_appraisals', section: 'Document Checklist', label: 'Valuation and Appraisals', statusDefault: 'PENDING' },
    { id: 'docs_rental_income_evidence', section: 'Document Checklist', label: 'Rental income evidence', statusDefault: 'PENDING' },
    { id: 'docs_deed_of_priority', section: 'Document Checklist', label: 'Deed of priority', statusDefault: 'PENDING' },
    { id: 'docs_cross_collateralisation_deed', section: 'Document Checklist', label: 'Cross collateralisation deed', statusDefault: 'PENDING' },
    { id: 'docs_solicitor_instructions', section: 'Document Checklist', label: 'Solicitor instructions', statusDefault: 'PENDING' },
    { id: 'docs_insurance_certificate_full_term', section: 'Document Checklist', label: 'Insurance certificate of currency covering full loan term', statusDefault: 'PENDING' },
    { id: 'docs_verified_deposit_paid_receipt', section: 'Document Checklist', label: '(Purchase/Sales) Verified deposit paid receipt', statusDefault: 'PENDING' },
    { id: 'docs_land_tax_notices_all_security', section: 'Document Checklist', label: '(If applicable) Land tax notices for all security properties', statusDefault: 'PENDING' }
  ]
};

const FALLBACK_SAMPLE_DEAL = {
  dealPresentationProvided: true,
  simSheetProvided: true,
  driversLicenceHeld: true,
  passportHeld: false,
  medicareHeld: true,
  alternateId100Points: false
};

const STATUS_VALUES = ['', 'PENDING', 'COMPLETE', 'N/A'];

// Explicit mapping: deal JSON keys -> checklist item ids
// (Supports both the new sample keys and the earlier example keys.)
const DEAL_KEY_TO_ITEM_ID = {
  // Deal Presentation
  dealPresentationProvided: 'deal_presentation_deal_presentation',
  simSheetProvided: 'deal_presentation_sim_sheet',

  // IDs
  driversLicenceHeld: 'id_current_drivers_licence_held',
  driversLicenseHeld: 'id_current_drivers_licence_held', // legacy spelling
  passportHeld: 'id_current_passport_held',
  medicareHeld: 'id_current_medicare_held',
  alternateId100Points: 'id_alternate_id_100_points',

  // LOI
  loiSignaturesInitials: 'loi_all_signatures_initials',
  loiNumbersMatch: 'loi_numbers_match_outline_conditions',

  // Application Form
  applicationAllSectionsCompleted: 'app_all_sections_completed',
  businessPurposeBoxTicked: 'app_business_purpose_box_ticked',
  applicationSignatureMatchesId: 'app_signature_matches_id',
  exitStrategyLoanPurposeProvided: 'app_exit_strategy_loan_purpose_provided',
  securityOfferedDetailsCorrect: 'app_security_offered_details_correct',
  borrowersALCorrect: 'app_borrowers_al_correct',
  accountantSolicitorBrokerVerification: 'app_accountant_solicitor_broker_verification',
  directorsMatchCompanyFile: 'app_directors_match_company_file',

  // ABN Search
  abnStatusChecked: 'abn_check_abn_status',
  abnChecked: 'abn_check_abn_status', // legacy example key
  gstRegistered: 'abn_registered_for_gst_or_not',

  // ASIC Search
  asicAdverseChecked: 'asic_adverse_information',

  // Borrower
  borrowerBackgroundProvided: 'borrower_background_provided',
  borrowerAlerasSearchCompleted: 'borrower_aleras_search_completed',
  borrowerRegisteredCompanyTrust: 'borrower_registered_company_trust',
  trustDeedFullyCompleted: 'borrower_trust_deed_fully_completed',
  companySearchConducted: 'borrower_company_search_conducted',
  companySearchDone: 'borrower_company_search_conducted', // legacy example key
  noAdverseOnCompanyFile: 'borrower_no_adverse_on_company_file',
  companyFileMatchesId: 'borrower_company_file_matches_id',

  // Equifax
  equifaxEnquiryCorrectEntity: 'equifax_enquiry_under_archer_wealth_pty_ltd',
  equifaxDirectorsInPlaceGt6Months: 'equifax_directors_in_place_gt_6_months',
  equifaxScoresGt600: 'equifax_scores_gt_600',
  equifaxClearOfAdverse: 'equifax_clear_of_adverse',
  equifaxClear: 'equifax_clear_of_adverse', // legacy example key

  // Guarantors
  noBorrowersOver65: 'guarantors_no_borrowers_over_65',
  noThirdPartyGuarantors: 'guarantors_no_third_party_guarantors',
  guarantorAlerasSearchCompleted: 'guarantors_aleras_search_completed',
  bankruptcySearchCompleted: 'guarantors_bankruptcy_search_completed',
  fraudGoogleSearchCompleted: 'guarantors_fraud_google_search_completed',
  guarantorsStrongALs: 'guarantors_strong_als',

  // Accountant letter
  accountantRegistrationChecked: 'accountant_letter_registration_checked',
  accountantLetterhead: 'accountant_letter_letterhead',
  accountantLetterDatedAddressedAW: 'accountant_letter_dated_addressed_aw',
  accountantLetterSigned: 'accountant_letter_signed',
  serviceabilityCalculatorCompleted: 'accountant_letter_serviceability_calculator_completed',

  // ATO portals
  guarantorTaxPortalsProvided: 'ato_guarantor_tax_portals_provided',
  companyTaxPortalsProvided: 'ato_company_tax_portals_provided',
  anyDirectorPenalties: 'ato_any_director_penalties',
  directorIdProvided: 'ato_director_id_provided',

  // Securities
  titleSearchConducted: 'securities_title_search_conducted',
  noRestrictionsOnUseOfLand: 'securities_no_restrictions_on_use_of_land',
  allSecuritiesInvestmentProperties: 'securities_all_investment_properties',
  noCaveatsOrEasements: 'securities_no_caveats_or_easements',
  allPagesInitialledExecuted: 'securities_all_pages_initialled_executed',

  // Offer
  offerSignatureMatchesId: 'offer_signature_matches_id',
  commitmentFeesPaid: 'offer_commitment_fees_paid',

  // Valuation & appraisals
  anyChangesToDealMetrics: 'valuation_any_changes_to_deal_metrics',
  valuationDatedWithin90Days: 'valuation_dated_within_90_days',
  valuationsAddressedToArcherWealth: 'valuation_addressed_to_archer_wealth',
  valuationReviewCompleted: 'valuation_review_completed',
  valuationReliabilityScorecardCompleted: 'valuation_reliability_scorecard_completed',
  newBuildGstIncluded: 'valuation_new_build_gst_included',

  // Contract of sale
  fullyExecutedContractObtained: 'cos_fully_executed_obtained',
  contractAddressCorrect: 'cos_address_correct',
  purchaserNamesSignedDated: 'cos_purchaser_names_signed_dated',
  depositPaidAmounts: 'cos_deposit_paid_amounts',
  contractDebtsChecked: 'cos_additional_check_for_debts',

  // Refinance statements
  mortgageStatements6Months: 'refinance_mortgage_statements_6_months',
  mortgageStatementsAddressCorrect: 'refinance_check_address_correct',

  // Document checklist
  docExecutedApplicationForm: 'docs_executed_application_form',
  docExecutedOfferLetter: 'docs_executed_offer_letter',
  docValidLicenceOrPassport: 'docs_valid_licence_or_passport',
  docDirectorIdProvided: 'docs_director_id_provided',
  docContractOfSale: 'docs_contract_of_sale',
  docTrustDeed: 'docs_trust_deed',
  docServicingDocs: 'docs_servicing_docs',
  docRefinanceStatements: 'docs_refinance_statements',
  docRates: 'docs_rates',
  docCompanySearch: 'docs_company_search',
  docTitleSearch: 'docs_title_search',
  docCreditReport: 'docs_credit_report',
  docValuationAndAppraisals: 'docs_valuation_and_appraisals',
  docRentalIncomeEvidence: 'docs_rental_income_evidence',
  docDeedOfPriority: 'docs_deed_of_priority',
  docCrossCollateralisationDeed: 'docs_cross_collateralisation_deed',
  docSolicitorInstructions: 'docs_solicitor_instructions',
  docInsuranceCertificateFullTerm: 'docs_insurance_certificate_full_term',
  docVerifiedDepositPaidReceipt: 'docs_verified_deposit_paid_receipt',
  docLandTaxNoticesAllSecurity: 'docs_land_tax_notices_all_security'
};

const state = {
  checklist: null,
  rows: new Map(), // id -> { id, section, label, status, comments }
  lastDeal: null,
};

function $(id){
  return document.getElementById(id);
}

function setMeta(text){
  const el = $('sheetMeta');
  if (el) el.textContent = text;
}

async function loadJson(url, fallback){
  try{
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return await res.json();
  } catch (_err){
    return fallback;
  }
}

function computeCheckResult(status){
  if (status === 'PENDING') return 'Incomplete';
  if (status === 'COMPLETE') return 'Complete';
  return '';
}

function statusClass(status){
  if (status === 'PENDING') return 'status-pending';
  if (status === 'COMPLETE') return 'status-complete';
  if (status === 'N/A') return 'status-na';
  return '';
}

function createSectionRow(section){
  const tr = document.createElement('tr');
  tr.className = 'tr-section';

  const td = document.createElement('td');
  td.colSpan = 4;
  td.textContent = section;
  tr.appendChild(td);
  return tr;
}

function createChecklistRow(row){
  const tr = document.createElement('tr');
  tr.dataset.rowId = row.id;

  // Item/Description
  const tdItem = document.createElement('td');
  tdItem.textContent = row.label;
  tr.appendChild(tdItem);

  // Status
  const tdStatus = document.createElement('td');
  tdStatus.className = 'status-cell';

  const select = document.createElement('select');
  select.className = 'status-select';
  select.setAttribute('aria-label', `Status for ${row.label}`);
  for (const v of STATUS_VALUES){
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    select.appendChild(opt);
  }
  select.value = row.status || '';

  select.addEventListener('change', () => {
    updateRowState(row.id, { status: select.value });
    renderRowCells(row.id);
    setMeta(`Edited: ${row.label}`);
  });

  tdStatus.appendChild(select);
  tr.appendChild(tdStatus);

  // CHECK (computed)
  const tdCheck = document.createElement('td');
  tdCheck.className = 'check-cell';
  tdCheck.dataset.cell = 'check';
  tdCheck.textContent = computeCheckResult(row.status || '');
  tr.appendChild(tdCheck);

  // COMMENTS
  const tdComments = document.createElement('td');
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'comment-input';
  input.placeholder = '';
  input.value = row.comments || '';
  input.setAttribute('aria-label', `Comments for ${row.label}`);
  input.addEventListener('input', () => {
    updateRowState(row.id, { comments: input.value });
  });
  input.addEventListener('blur', () => {
    setMeta('Ready');
  });
  input.addEventListener('focus', () => {
    setMeta(`Editing comments: ${row.label}`);
  });

  tdComments.appendChild(input);
  tr.appendChild(tdComments);

  renderRowCells(row.id, tr);
  return tr;
}

function updateRowState(rowId, patch){
  const current = state.rows.get(rowId);
  if (!current) return;
  state.rows.set(rowId, { ...current, ...patch });
}

function renderRowCells(rowId, trEl){
  const row = state.rows.get(rowId);
  if (!row) return;
  const tr = trEl || document.querySelector(`tr[data-row-id="${cssEscape(rowId)}"]`);
  if (!tr) return;

  const statusTd = tr.children[1];
  const checkTd = tr.querySelector('[data-cell="check"]');

  // conditional formatting on Status cell
  statusTd.classList.remove('status-pending','status-complete','status-na');
  const cls = statusClass(row.status);
  if (cls) statusTd.classList.add(cls);

  // computed check
  if (checkTd) checkTd.textContent = computeCheckResult(row.status);
}

function cssEscape(value){
  // minimal escape for attribute selector
  return String(value).replace(/"/g, '\\"');
}

function initRowsFromChecklist(checklist){
  state.rows.clear();
  for (const r of checklist.rows){
    state.rows.set(r.id, {
      id: r.id,
      section: r.section,
      label: r.label,
      status: r.statusDefault || 'PENDING',
      comments: ''
    });
  }
}

function verifyRowCounts(){
  const banner = $('rowCountBanner');
  if (!banner || !state.checklist) return;

  const expected = Array.isArray(state.checklist.rows) ? state.checklist.rows.length : 0;
  const rendered = document.querySelectorAll('tr[data-row-id]').length;

  banner.classList.remove('ok','bad');
  if (rendered === expected){
    banner.classList.add('ok');
    banner.textContent = `All rows loaded: ${rendered}/${expected}`;
  } else {
    banner.classList.add('bad');
    banner.textContent = `Missing rows: rendered ${rendered}, expected ${expected}`;
  }
}

function renderTable(checklist){
  const body = $('gridBody');
  body.innerHTML = '';

  let currentSection = null;
  for (const r of checklist.rows){
    if (r.section !== currentSection){
      currentSection = r.section;
      body.appendChild(createSectionRow(currentSection));
    }
    const rowState = state.rows.get(r.id);
    body.appendChild(createChecklistRow(rowState));
  }

  setMeta('Ready');
  verifyRowCounts();
}

function applyDealToChecklist(deal){
  if (!state.checklist) return;

  state.lastDeal = deal;

  for (const [dealKey, itemId] of Object.entries(DEAL_KEY_TO_ITEM_ID)){
    if (!(dealKey in deal)) continue; // missing => leave default

    const v = deal[dealKey];
    if (v === true) updateRowState(itemId, { status: 'COMPLETE' });
    else if (v === false) updateRowState(itemId, { status: 'PENDING' });
    else updateRowState(itemId, { status: '' });

    renderRowCells(itemId);
    const tr = document.querySelector(`tr[data-row-id="${cssEscape(itemId)}"]`);
    if (tr){
      const sel = tr.querySelector('select.status-select');
      if (sel) sel.value = state.rows.get(itemId).status || '';
    }
  }

  setMeta('Autofill applied');
}

function exportChecklist(){
  const exportedRows = [];

  for (const r of state.checklist.rows){
    const s = state.rows.get(r.id);
    exportedRows.push({
      id: s.id,
      section: s.section,
      label: s.label,
      status: s.status || '',
      checkResult: computeCheckResult(s.status || ''),
      comments: s.comments || ''
    });
  }

  const payload = {
    sheet: state.checklist.sheet,
    version: state.checklist.version,
    exportedAt: new Date().toISOString(),
    lastDeal: state.lastDeal || null,
    rows: exportedRows
  };

  downloadJson('credit-checklist.export.json', payload);
  setMeta('Exported');
}

function downloadJson(filename, obj){
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

async function handleDealUpload(file){
  if (!file) return;
  const text = await file.text();
  const deal = JSON.parse(text);
  applyDealToChecklist(deal);
}

function initTabs(){
  document.querySelectorAll('.tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      // MVP: only Credit Checklist is implemented.
      const name = btn.getAttribute('data-sheet') || '';
      if (name !== 'Credit Checklist'){
        alert(`"${name}" is a placeholder tab in this MVP. Only "Credit Checklist" is implemented.`);
        return;
      }
    });
  });
}

async function main(){
  initTabs();

  // wire buttons
  $('btnUploadDeal').addEventListener('click', () => $('dealFileInput').click());
  $('dealFileInput').addEventListener('change', async (e) => {
    const file = e.target.files && e.target.files[0];
    try{
      await handleDealUpload(file);
    } catch (err){
      alert('Could not read deal JSON file.');
      console.error(err);
    } finally {
      e.target.value = '';
    }
  });

  $('btnLoadSample').addEventListener('click', async () => {
    const deal = await loadJson(SAMPLE_DEAL_URL, FALLBACK_SAMPLE_DEAL);
    applyDealToChecklist(deal);
  });

  $('btnExportChecklist').addEventListener('click', exportChecklist);

  // load checklist template
  setMeta('Loading checklist…');
  const checklist = await loadJson(CHECKLIST_URL, FALLBACK_CHECKLIST);
  state.checklist = checklist;
  initRowsFromChecklist(checklist);
  renderTable(checklist);
  setMeta('Ready');
}

main().catch((err) => {
  console.error(err);
  alert('Failed to load the checklist app.');
});
