// Static, browser-only Credit Checklist workspace
// - No backend, no build step
// - Runs by opening index.html
// - Uses localStorage for deal list, checklist state, attachment metadata, and audit trail

const CHECKLIST_URL = './data/checklist.credit.json';
const SAMPLE_DEAL_URL = './data/sample.deal.json';

// ---------- LocalStorage keys ----------
const LS_DEALS_INDEX = 'cc:deals:index';
const LS_CHECKLIST_PREFIX = 'creditChecklist:'; // per deal
const LS_ATTACH_PREFIX = 'cc:attachments:'; // metadata only
const LS_HISTORY_PREFIX = 'cc:history:';

// ---------- Checklist fallbacks (file:// fetch-safe) ----------
// Many browsers block fetch() for file://. To keep the app runnable by opening index.html,
// we embed the checklist template + sample deal as fallbacks.
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

const FALLBACK_DEAL = {
  dealId: 'sample-001',
  dealPresentationProvided: true,
  simSheetProvided: true,
  driversLicenceHeld: true,
  passportHeld: false,
  medicareHeld: true,
  alternateId100Points: false,

  loiSignaturesInitials: true,
  loiNumbersMatch: true,

  applicationAllSectionsCompleted: true,
  businessPurposeBoxTicked: true,
  applicationSignatureMatchesId: true,
  exitStrategyLoanPurposeProvided: true,
  securityOfferedDetailsCorrect: true,
  borrowersALCorrect: true,
  accountantSolicitorBrokerVerification: true,
  directorsMatchCompanyFile: true,

  abnStatusChecked: true,
  gstRegistered: false,

  asicAdverseChecked: true,

  borrowerBackgroundProvided: true,
  borrowerAlerasSearchCompleted: true,
  borrowerRegisteredCompanyTrust: true,
  trustDeedFullyCompleted: true,
  companySearchConducted: true,
  noAdverseOnCompanyFile: true,
  companyFileMatchesId: true,

  equifaxEnquiryCorrectEntity: true,
  equifaxDirectorsInPlaceGt6Months: true,
  equifaxScoresGt600: true,
  equifaxClearOfAdverse: true,

  noBorrowersOver65: true,
  noThirdPartyGuarantors: true,
  guarantorAlerasSearchCompleted: true,
  bankruptcySearchCompleted: true,
  fraudGoogleSearchCompleted: true,
  guarantorsStrongALs: true,

  accountantRegistrationChecked: true,
  accountantLetterhead: true,
  accountantLetterDatedAddressedAW: true,
  accountantLetterSigned: true,
  serviceabilityCalculatorCompleted: true,

  guarantorTaxPortalsProvided: true,
  companyTaxPortalsProvided: true,
  anyDirectorPenalties: false,
  directorIdProvided: true,

  titleSearchConducted: true,
  noRestrictionsOnUseOfLand: true,
  allSecuritiesInvestmentProperties: true,
  noCaveatsOrEasements: true,
  allPagesInitialledExecuted: true,

  offerSignatureMatchesId: true,
  commitmentFeesPaid: true,

  anyChangesToDealMetrics: false,
  valuationDatedWithin90Days: true,
  valuationsAddressedToArcherWealth: true,
  valuationReviewCompleted: true,
  valuationReliabilityScorecardCompleted: true,
  newBuildGstIncluded: true,

  fullyExecutedContractObtained: true,
  contractAddressCorrect: true,
  purchaserNamesSignedDated: true,
  depositPaidAmounts: true,
  contractDebtsChecked: true,

  mortgageStatements6Months: true,
  mortgageStatementsAddressCorrect: true,

  docExecutedApplicationForm: true,
  docExecutedOfferLetter: true,
  docValidLicenceOrPassport: true,
  docDirectorIdProvided: true,
  docContractOfSale: true,
  docTrustDeed: true,
  docServicingDocs: true,
  docRefinanceStatements: true,
  docRates: true,
  docCompanySearch: true,
  docTitleSearch: true,
  docCreditReport: true,
  docValuationAndAppraisals: true,
  docRentalIncomeEvidence: true,
  docDeedOfPriority: false,
  docCrossCollateralisationDeed: false,
  docSolicitorInstructions: true,
  docInsuranceCertificateFullTerm: true,
  docVerifiedDepositPaidReceipt: true,
  docLandTaxNoticesAllSecurity: false
};

// ---------- Status + mapping ----------
const STATUS_VALUES = ['', 'PENDING', 'COMPLETE', 'N/A'];

// Deal JSON keys -> checklist item ids
const DEAL_KEY_TO_ITEM_ID = {
  dealPresentationProvided: 'deal_presentation_deal_presentation',
  simSheetProvided: 'deal_presentation_sim_sheet',

  driversLicenceHeld: 'id_current_drivers_licence_held',
  driversLicenseHeld: 'id_current_drivers_licence_held',
  passportHeld: 'id_current_passport_held',
  medicareHeld: 'id_current_medicare_held',
  alternateId100Points: 'id_alternate_id_100_points',

  loiSignaturesInitials: 'loi_all_signatures_initials',
  loiNumbersMatch: 'loi_numbers_match_outline_conditions',

  applicationAllSectionsCompleted: 'app_all_sections_completed',
  businessPurposeBoxTicked: 'app_business_purpose_box_ticked',
  applicationSignatureMatchesId: 'app_signature_matches_id',
  exitStrategyLoanPurposeProvided: 'app_exit_strategy_loan_purpose_provided',
  securityOfferedDetailsCorrect: 'app_security_offered_details_correct',
  borrowersALCorrect: 'app_borrowers_al_correct',
  accountantSolicitorBrokerVerification: 'app_accountant_solicitor_broker_verification',
  directorsMatchCompanyFile: 'app_directors_match_company_file',

  abnStatusChecked: 'abn_check_abn_status',
  abnChecked: 'abn_check_abn_status',
  gstRegistered: 'abn_registered_for_gst_or_not',

  asicAdverseChecked: 'asic_adverse_information',

  borrowerBackgroundProvided: 'borrower_background_provided',
  borrowerAlerasSearchCompleted: 'borrower_aleras_search_completed',
  borrowerRegisteredCompanyTrust: 'borrower_registered_company_trust',
  trustDeedFullyCompleted: 'borrower_trust_deed_fully_completed',
  companySearchConducted: 'borrower_company_search_conducted',
  companySearchDone: 'borrower_company_search_conducted',
  noAdverseOnCompanyFile: 'borrower_no_adverse_on_company_file',
  companyFileMatchesId: 'borrower_company_file_matches_id',

  equifaxEnquiryCorrectEntity: 'equifax_enquiry_under_archer_wealth_pty_ltd',
  equifaxDirectorsInPlaceGt6Months: 'equifax_directors_in_place_gt_6_months',
  equifaxScoresGt600: 'equifax_scores_gt_600',
  equifaxClearOfAdverse: 'equifax_clear_of_adverse',
  equifaxClear: 'equifax_clear_of_adverse',

  noBorrowersOver65: 'guarantors_no_borrowers_over_65',
  noThirdPartyGuarantors: 'guarantors_no_third_party_guarantors',
  guarantorAlerasSearchCompleted: 'guarantors_aleras_search_completed',
  bankruptcySearchCompleted: 'guarantors_bankruptcy_search_completed',
  fraudGoogleSearchCompleted: 'guarantors_fraud_google_search_completed',
  guarantorsStrongALs: 'guarantors_strong_als',

  accountantRegistrationChecked: 'accountant_letter_registration_checked',
  accountantLetterhead: 'accountant_letter_letterhead',
  accountantLetterDatedAddressedAW: 'accountant_letter_dated_addressed_aw',
  accountantLetterSigned: 'accountant_letter_signed',
  serviceabilityCalculatorCompleted: 'accountant_letter_serviceability_calculator_completed',

  guarantorTaxPortalsProvided: 'ato_guarantor_tax_portals_provided',
  companyTaxPortalsProvided: 'ato_company_tax_portals_provided',
  anyDirectorPenalties: 'ato_any_director_penalties',
  directorIdProvided: 'ato_director_id_provided',

  titleSearchConducted: 'securities_title_search_conducted',
  noRestrictionsOnUseOfLand: 'securities_no_restrictions_on_use_of_land',
  allSecuritiesInvestmentProperties: 'securities_all_investment_properties',
  noCaveatsOrEasements: 'securities_no_caveats_or_easements',
  allPagesInitialledExecuted: 'securities_all_pages_initialled_executed',

  offerSignatureMatchesId: 'offer_signature_matches_id',
  commitmentFeesPaid: 'offer_commitment_fees_paid',

  anyChangesToDealMetrics: 'valuation_any_changes_to_deal_metrics',
  valuationDatedWithin90Days: 'valuation_dated_within_90_days',
  valuationsAddressedToArcherWealth: 'valuation_addressed_to_archer_wealth',
  valuationReviewCompleted: 'valuation_review_completed',
  valuationReliabilityScorecardCompleted: 'valuation_reliability_scorecard_completed',
  newBuildGstIncluded: 'valuation_new_build_gst_included',

  fullyExecutedContractObtained: 'cos_fully_executed_obtained',
  contractAddressCorrect: 'cos_address_correct',
  purchaserNamesSignedDated: 'cos_purchaser_names_signed_dated',
  depositPaidAmounts: 'cos_deposit_paid_amounts',
  contractDebtsChecked: 'cos_additional_check_for_debts',

  mortgageStatements6Months: 'refinance_mortgage_statements_6_months',
  mortgageStatementsAddressCorrect: 'refinance_check_address_correct',

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

// ---------- Utilities ----------
const $ = (id) => document.getElementById(id);

function safeLocalStorage(){
  try{
    if (!('localStorage' in window)) return null;
    const k = '__cc_test__';
    localStorage.setItem(k, '1');
    localStorage.removeItem(k);
    return localStorage;
  } catch (_err){
    return null;
  }
}
const LS = safeLocalStorage();

function nowIso(){ return new Date().toISOString(); }

function setMeta(text){ const el = $('sheetMeta'); if (el) el.textContent = text; }
function setAutofillBanner(text){ const el = $('autofillBanner'); if (el) el.textContent = text || ''; }

function randomId(prefix='deal'){
  if (globalThis.crypto?.randomUUID) return `${prefix}-${crypto.randomUUID()}`;
  return `${prefix}-${Math.random().toString(16).slice(2)}-${Date.now().toString(16)}`;
}

function fnv1a32(str){
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++){
    h ^= str.charCodeAt(i);
    h = (h * 0x01000193) >>> 0;
  }
  return ('00000000' + h.toString(16)).slice(-8);
}

function downloadText(filename, content, mime){
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function toCsv(rows){
  const esc = (v) => {
    const s = String(v ?? '');
    if (/[\",\n\r]/.test(s)) return `\"${s.replace(/\"/g,'\"\"')}\"`;
    return s;
  };
  return rows.map(r => r.map(esc).join(',')).join('\n');
}

function computeCheckResult(status){
  if (!status) return '';
  if (status === 'PENDING') return 'Incomplete';
  if (status === 'COMPLETE') return 'Complete';
  return '';
}

function runCheckColumnTests(){
  const cases = [
    { status: '', expected: '' },
    { status: 'PENDING', expected: 'Incomplete' },
    { status: 'COMPLETE', expected: 'Complete' },
    { status: 'N/A', expected: '' },
  ];
  for (const c of cases){
    const got = computeCheckResult(c.status);
    if (got !== c.expected){
      throw new Error(`CHECK test failed: status=${JSON.stringify(c.status)} expected=${JSON.stringify(c.expected)} got=${JSON.stringify(got)}`);
    }
  }
}

function escapeText(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
function escapeAttr(s){
  return escapeText(s).replace(/\"/g,'&quot;');
}

// ---------- In-memory file store (session only) ----------
const attachmentMemory = new Map(); // dealId -> Map(attachmentId -> File)

// ---------- App state ----------
const state = {
  checklist: null,
  checklistVersionHash: null,
  rows: new Map(),
  deals: [],
  activeDealId: 'local-demo',
  lastDealJson: null,
  aiPreview: null,
};

// ---------- Deal index ----------
function loadDealsIndex(){
  if (!LS) return [{ dealId: 'local-demo', name: 'Local Demo', createdAt: nowIso(), updatedAt: nowIso() }];
  const raw = LS.getItem(LS_DEALS_INDEX);
  if (!raw){
    const seed = [{ dealId: 'local-demo', name: 'Local Demo', createdAt: nowIso(), updatedAt: nowIso() }];
    LS.setItem(LS_DEALS_INDEX, JSON.stringify(seed));
    return seed;
  }
  try{
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed) || parsed.length === 0) throw new Error('bad');
    return parsed;
  } catch (_err){
    const seed = [{ dealId: 'local-demo', name: 'Local Demo', createdAt: nowIso(), updatedAt: nowIso() }];
    LS.setItem(LS_DEALS_INDEX, JSON.stringify(seed));
    return seed;
  }
}

function saveDealsIndex(){
  if (!LS) return;
  LS.setItem(LS_DEALS_INDEX, JSON.stringify(state.deals));
}

function upsertDealMeta(patch){
  const idx = state.deals.findIndex(d => d.dealId === patch.dealId);
  if (idx >= 0) state.deals[idx] = { ...state.deals[idx], ...patch };
  else state.deals.unshift({ ...patch });
  saveDealsIndex();
}

function deleteDealMeta(dealId){
  state.deals = state.deals.filter(d => d.dealId !== dealId);
  saveDealsIndex();
}

function setActiveDealId(dealId){
  state.activeDealId = dealId || 'local-demo';
  const el = $('dealIdLabel');
  if (el) el.textContent = state.activeDealId;
}

function checklistStorageKey(dealId){ return `${LS_CHECKLIST_PREFIX}${dealId}`; }
function attachStorageKey(dealId){ return `${LS_ATTACH_PREFIX}${dealId}`; }
function historyStorageKey(dealId){ return `${LS_HISTORY_PREFIX}${dealId}`; }

// ---------- Checklist state persistence ----------
function saveChecklistState({ silent = false } = {}){
  if (!LS) return;

  const dealId = state.activeDealId || 'local-demo';
  const key = checklistStorageKey(dealId);

  const rowsOut = {};
  for (const [id, r] of state.rows.entries()){
    rowsOut[id] = { status: r.status || '', comments: r.comments || '' };
  }

  const dealMeta = state.deals.find(d => d.dealId === dealId);

  const payload = {
    dealId,
    dealName: dealMeta?.name || null,
    savedAt: nowIso(),
    checklistVersion: state.checklistVersionHash,
    rows: rowsOut
  };

  LS.setItem(key, JSON.stringify(payload));
  upsertDealMeta({ dealId, name: dealMeta?.name || dealId, createdAt: dealMeta?.createdAt || nowIso(), updatedAt: payload.savedAt });
  renderDealList();

  if (!silent) setMeta(`Saved (${dealId})`);
}

function loadChecklistState(dealId){
  if (!LS) return null;
  const raw = LS.getItem(checklistStorageKey(dealId));
  if (!raw) return null;
  try{ return JSON.parse(raw); } catch (_err){ return null; }
}

let autosaveTimer = null;
function scheduleAutosave(){
  if (!LS) return;
  clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(() => {
    try{ saveChecklistState({ silent: true }); } catch (_err) {}
  }, 500);
}

function resetToDefaults(){
  if (!state.checklist) return;
  for (const def of state.checklist.rows){
    const r = state.rows.get(def.id);
    if (!r) continue;
    state.rows.set(def.id, { ...r, status: def.statusDefault || 'PENDING', comments: '' });
  }
  syncDomFromState();
  scheduleAutosave();
  setMeta('Reset to defaults');
}

function clearSavedState(){
  if (!LS) return;
  const dealId = state.activeDealId || 'local-demo';
  LS.removeItem(checklistStorageKey(dealId));
  setMeta('Cleared saved checklist state');
}

// ---------- Attachments (metadata persisted, file in memory) ----------
function loadAttachmentMeta(dealId){
  if (!LS) return [];
  const raw = LS.getItem(attachStorageKey(dealId));
  if (!raw) return [];
  try{ return Array.isArray(JSON.parse(raw)) ? JSON.parse(raw) : []; } catch (_err){ return []; }
}

function saveAttachmentMeta(dealId, metaList){
  if (!LS) return;
  LS.setItem(attachStorageKey(dealId), JSON.stringify(metaList));
}

function addAttachments(files){
  const dealId = state.activeDealId || 'local-demo';
  const meta = loadAttachmentMeta(dealId);
  let mem = attachmentMemory.get(dealId);
  if (!mem){ mem = new Map(); attachmentMemory.set(dealId, mem); }

  for (const f of files){
    const attachmentId = randomId('att');
    mem.set(attachmentId, f);
    meta.unshift({ attachmentId, name: f.name, type: f.type || 'application/octet-stream', size: f.size, lastModified: f.lastModified, addedAt: nowIso() });
  }

  saveAttachmentMeta(dealId, meta);
  renderAttachments();
}

function removeAttachment(attachmentId){
  const dealId = state.activeDealId || 'local-demo';
  const meta = loadAttachmentMeta(dealId).filter(m => m.attachmentId !== attachmentId);
  saveAttachmentMeta(dealId, meta);
  attachmentMemory.get(dealId)?.delete(attachmentId);
  renderAttachments();
}

function clearAttachmentMetadata(){
  if (!LS) return;
  const dealId = state.activeDealId || 'local-demo';
  LS.removeItem(attachStorageKey(dealId));
  attachmentMemory.delete(dealId);
  renderAttachments();
}

// ---------- Audit trail ----------
function loadHistory(dealId){
  if (!LS) return [];
  const raw = LS.getItem(historyStorageKey(dealId));
  if (!raw) return [];
  try{ return Array.isArray(JSON.parse(raw)) ? JSON.parse(raw) : []; } catch (_err){ return []; }
}

function saveHistory(dealId, entries){
  if (!LS) return;
  LS.setItem(historyStorageKey(dealId), JSON.stringify(entries));
}

function appendHistory(entry){
  if (!LS) return;
  const dealId = state.activeDealId || 'local-demo';
  const entries = loadHistory(dealId);
  entries.unshift(entry);
  if (entries.length > 5000) entries.length = 5000;
  saveHistory(dealId, entries);
}

// ---------- Rendering ----------
function statusClass(status){
  if (status === 'PENDING') return 'status-pending';
  if (status === 'COMPLETE') return 'status-complete';
  if (status === 'N/A') return 'status-na';
  return '';
}

function cssEscape(v){
  return String(v).replace(/\"/g, '\\\"');
}

function createSectionRow(section){
  const tr = document.createElement('tr');
  tr.className = 'tr-section';
  tr.dataset.section = section;
  tr.id = `section-${fnv1a32(section)}`;

  const td = document.createElement('td');
  td.colSpan = 4;
  td.textContent = section;
  tr.appendChild(td);
  return tr;
}

function createItemRow(row){
  const tr = document.createElement('tr');
  tr.dataset.rowId = row.id;
  tr.dataset.section = row.section;

  const tdLabel = document.createElement('td');
  tdLabel.textContent = row.label;

  const tdStatus = document.createElement('td');
  tdStatus.className = 'status-cell';

  const sel = document.createElement('select');
  sel.className = 'status-select';
  for (const v of STATUS_VALUES){
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    sel.appendChild(opt);
  }
  sel.value = row.status || '';
  sel.addEventListener('change', () => {
    const prev = state.rows.get(row.id)?.status || '';
    const next = sel.value;
    if (prev !== next){
      state.rows.set(row.id, { ...state.rows.get(row.id), status: next });
      appendHistory({ timestamp: nowIso(), itemId: row.id, field: 'status', oldValue: prev, newValue: next });
    }
    renderRowCells(row.id);
    scheduleAutosave();
  });
  sel.addEventListener('focus', () => {
    const ss = $('sectionSelect');
    if (ss) ss.value = row.section;
  });

  tdStatus.appendChild(sel);

  const tdCheck = document.createElement('td');
  tdCheck.dataset.cell = 'check';

  const tdComments = document.createElement('td');
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'comment-input';
  input.value = row.comments || '';
  input.addEventListener('input', () => {
    const prev = state.rows.get(row.id)?.comments || '';
    const next = input.value;
    if (prev !== next){
      state.rows.set(row.id, { ...state.rows.get(row.id), comments: next });
      appendHistory({ timestamp: nowIso(), itemId: row.id, field: 'comment', oldValue: prev, newValue: next });
    }
    scheduleAutosave();
  });
  input.addEventListener('focus', () => {
    const ss = $('sectionSelect');
    if (ss) ss.value = row.section;
  });

  tdComments.appendChild(input);

  tr.appendChild(tdLabel);
  tr.appendChild(tdStatus);
  tr.appendChild(tdCheck);
  tr.appendChild(tdComments);

  renderRowCells(row.id, tr);
  return tr;
}

function renderRowCells(itemId, trEl){
  const row = state.rows.get(itemId);
  if (!row) return;
  const tr = trEl || document.querySelector(`tr[data-row-id=\"${cssEscape(itemId)}\"]`);
  if (!tr) return;

  const statusTd = tr.children[1];
  const checkTd = tr.querySelector('[data-cell="check"]');

  statusTd.classList.remove('status-pending','status-complete','status-na');
  const cls = statusClass(row.status);
  if (cls) statusTd.classList.add(cls);

  if (checkTd) checkTd.textContent = computeCheckResult(row.status);
}

function verifyRowCounts(){
  const banner = $('rowCountBanner');
  if (!banner || !state.checklist) return;
  const expected = state.checklist.rows.length;
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

function renderTable(){
  const body = $('gridBody');
  body.innerHTML = '';

  let currentSection = null;
  for (const def of state.checklist.rows){
    if (def.section !== currentSection){
      currentSection = def.section;
      body.appendChild(createSectionRow(currentSection));
    }
    body.appendChild(createItemRow(state.rows.get(def.id)));
  }

  verifyRowCounts();
  renderProgressSummary();
}

function syncDomFromState(){
  for (const [id, r] of state.rows.entries()){
    const tr = document.querySelector(`tr[data-row-id=\"${cssEscape(id)}\"]`);
    if (!tr) continue;
    const sel = tr.querySelector('select.status-select');
    const input = tr.querySelector('input.comment-input');
    if (sel) sel.value = r.status || '';
    if (input) input.value = r.comments || '';
    renderRowCells(id);
  }
  verifyRowCounts();
  renderProgressSummary();
}

function renderSectionSelect(){
  const sel = $('sectionSelect');
  if (!sel || !state.checklist) return;
  const sections = [];
  const seen = new Set();
  for (const r of state.checklist.rows){
    if (!seen.has(r.section)){
      seen.add(r.section);
      sections.push(r.section);
    }
  }
  sel.innerHTML = '<option value="">Section…</option>' + sections.map(s => `<option value=\"${escapeAttr(s)}\">${escapeText(s)}</option>`).join('');
}

function computeProgress(){
  const perSection = new Map();
  let totalApplicable = 0;
  let totalComplete = 0;

  for (const def of state.checklist.rows){
    const r = state.rows.get(def.id);
    const section = r.section;
    const status = r.status || '';

    if (!perSection.has(section)) perSection.set(section, { section, applicable: 0, complete: 0 });

    if (status !== 'N/A'){
      perSection.get(section).applicable += 1;
      totalApplicable += 1;
      if (status === 'COMPLETE'){
        perSection.get(section).complete += 1;
        totalComplete += 1;
      }
    }
  }

  const sections = [...perSection.values()].map(s => ({ ...s, pct: s.applicable ? (s.complete / s.applicable) : 0 }));
  return { overall: { applicable: totalApplicable, complete: totalComplete, pct: totalApplicable ? (totalComplete / totalApplicable) : 0 }, sections };
}

function scrollToSection(section){
  const anchor = document.getElementById(`section-${fnv1a32(section)}`);
  if (anchor) anchor.scrollIntoView({ block: 'start' });
}

function renderProgressSummary(){
  const root = $('progressSummary');
  if (!root || !state.checklist) return;

  const p = computeProgress();
  const overallPct = Math.round(p.overall.pct * 100);

  root.innerHTML = '';

  const overall = document.createElement('div');
  overall.className = 'overall';
  overall.innerHTML = `<div><strong>Overall</strong> — ${overallPct}% (${p.overall.complete}/${p.overall.applicable})</div>`;

  const bar = document.createElement('div');
  bar.className = 'bar';
  const fill = document.createElement('div');
  fill.style.width = `${overallPct}%`;
  bar.appendChild(fill);

  root.appendChild(overall);
  root.appendChild(bar);

  const sectionsEl = document.createElement('div');
  sectionsEl.className = 'sections';

  p.sections.forEach(s => {
    const pct = Math.round(s.pct * 100);
    const row = document.createElement('div');
    row.className = 'section-row';
    row.innerHTML = `
      <div class="head">
        <div class="name">${escapeText(s.section)}</div>
        <div class="muted">${pct}% (${s.complete}/${s.applicable})</div>
      </div>
      <div class="bar"><div style="width:${pct}%; background:#cfe2ff;"></div></div>
    `;
    row.addEventListener('click', () => scrollToSection(s.section));
    sectionsEl.appendChild(row);
  });

  root.appendChild(sectionsEl);
}

// ---------- Deals list rendering ----------
function renderDealList(){
  const list = $('dealList');
  const count = $('dealCount');
  const q = ($('dealSearch')?.value || '').toLowerCase().trim();

  const deals = state.deals.slice().sort((a,b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')));
  const filtered = q ? deals.filter(d => (d.name || '').toLowerCase().includes(q) || (d.dealId || '').toLowerCase().includes(q)) : deals;

  if (count) count.textContent = `${filtered.length}`;
  list.innerHTML = '';

  filtered.forEach(d => {
    const div = document.createElement('div');
    div.className = 'deal-item' + (d.dealId === state.activeDealId ? ' active' : '');
    div.setAttribute('role','option');
    div.tabIndex = 0;
    div.innerHTML = `
      <div><strong>${escapeText(d.name || d.dealId)}</strong></div>
      <div class="meta">${escapeText(d.dealId)} • Updated ${escapeText((d.updatedAt||'').slice(0,19).replace('T',' '))}</div>
    `;
    div.addEventListener('click', () => selectDeal(d.dealId));
    div.addEventListener('keydown', (e) => { if (e.key === 'Enter') selectDeal(d.dealId); });
    list.appendChild(div);
  });
}

// ---------- Attachments rendering ----------
function renderAttachments(){
  const root = $('attachmentList');
  if (!root) return;

  const dealId = state.activeDealId || 'local-demo';
  const meta = loadAttachmentMeta(dealId);
  const mem = attachmentMemory.get(dealId);

  root.innerHTML = '';
  if (meta.length === 0){
    const empty = document.createElement('div');
    empty.className = 'attachment-item';
    empty.textContent = 'No attachments.';
    root.appendChild(empty);
    return;
  }

  meta.forEach(m => {
    const div = document.createElement('div');
    div.className = 'attachment-item';
    const inMem = mem?.has(m.attachmentId);

    div.innerHTML = `
      <div><strong>${escapeText(m.name)}</strong></div>
      <div class="meta">${escapeText(m.type)} • ${Math.round((m.size||0)/1024)} KB • ${inMem ? 'In memory' : 'Metadata only (re-upload needed)'}</div>
      <div class="actions">
        <button class="btn" data-action="remove">Remove</button>
      </div>
    `;

    div.querySelector('[data-action="remove"]').addEventListener('click', () => removeAttachment(m.attachmentId));
    root.appendChild(div);
  });
}

// ---------- Deal selection + CRUD ----------
function initRowsFromChecklist(){
  state.rows.clear();
  for (const def of state.checklist.rows){
    state.rows.set(def.id, { id: def.id, section: def.section, label: def.label, status: def.statusDefault || 'PENDING', comments: '' });
  }
}

function applyChecklistPayloadToState(payload){
  if (!payload || typeof payload !== 'object' || !payload.rows) return;
  for (const [id, patch] of Object.entries(payload.rows)){
    if (!state.rows.has(id)) continue;
    const current = state.rows.get(id);
    const status = typeof patch.status === 'string' ? patch.status : current.status;
    const comments = typeof patch.comments === 'string' ? patch.comments : current.comments;
    state.rows.set(id, { ...current, status, comments });
  }
}

function loadDealIntoWorkspace(dealId){
  setActiveDealId(dealId);
  initRowsFromChecklist();

  const saved = loadChecklistState(dealId);
  if (saved) applyChecklistPayloadToState(saved);

  renderTable();
  renderSectionSelect();
  renderAttachments();
  renderDealList();

  setAutofillBanner('');
  setMeta('Ready');
}

function selectDeal(dealId){
  if (!dealId) return;
  loadDealIntoWorkspace(dealId);
}

function newDeal(){
  const dealId = randomId('deal');
  const name = prompt('Deal name?', 'New Deal') || 'New Deal';
  const createdAt = nowIso();
  upsertDealMeta({ dealId, name, createdAt, updatedAt: createdAt });
  loadDealIntoWorkspace(dealId);
  scheduleAutosave();
}

function duplicateDeal(){
  const srcId = state.activeDealId || 'local-demo';
  const srcMeta = state.deals.find(d => d.dealId === srcId);
  const newId = randomId('deal');
  const name = prompt('Duplicate deal name?', `${srcMeta?.name || srcId} (Copy)`) || `${srcMeta?.name || srcId} (Copy)`;
  const createdAt = nowIso();

  if (LS){
    const saved = loadChecklistState(srcId);
    if (saved){
      const cloned = { ...saved, dealId: newId, dealName: name, savedAt: nowIso() };
      LS.setItem(checklistStorageKey(newId), JSON.stringify(cloned));
    }
    const att = loadAttachmentMeta(srcId);
    saveAttachmentMeta(newId, att.map(a => ({ ...a, attachmentId: randomId('att') })));
  }

  upsertDealMeta({ dealId: newId, name, createdAt, updatedAt: createdAt });
  loadDealIntoWorkspace(newId);
}

function deleteDeal(){
  const dealId = state.activeDealId || 'local-demo';
  if (!confirm(`Delete deal ${dealId}? (local only)`)) return;

  if (LS){
    LS.removeItem(checklistStorageKey(dealId));
    LS.removeItem(attachStorageKey(dealId));
    LS.removeItem(historyStorageKey(dealId));
  }
  attachmentMemory.delete(dealId);

  deleteDealMeta(dealId);

  if (!state.deals.find(d => d.dealId === 'local-demo')){
    state.deals.unshift({ dealId: 'local-demo', name: 'Local Demo', createdAt: nowIso(), updatedAt: nowIso() });
    saveDealsIndex();
  }

  const next = state.deals[0]?.dealId || 'local-demo';
  loadDealIntoWorkspace(next);
}

// ---------- Deal JSON autofill ----------
function computeAutofillCoverage(deal){
  const mappingKeys = Object.keys(DEAL_KEY_TO_ITEM_ID);
  const found = [];
  const missing = [];
  for (const k of mappingKeys){
    if (k in deal) found.push(k); else missing.push(k);
  }
  const unmapped = [];
  for (const k of Object.keys(deal)){
    if (!(k in DEAL_KEY_TO_ITEM_ID)) unmapped.push(k);
  }
  return { found, missing, unmapped };
}

function applyDealAutofill(deal){
  state.lastDealJson = deal;

  const incomingDealId = (deal && deal.dealId) ? String(deal.dealId) : null;
  if (incomingDealId){
    if (!state.deals.find(d => d.dealId === incomingDealId)){
      upsertDealMeta({ dealId: incomingDealId, name: incomingDealId, createdAt: nowIso(), updatedAt: nowIso() });
    }
    loadDealIntoWorkspace(incomingDealId);
  }

  for (const [dealKey, itemId] of Object.entries(DEAL_KEY_TO_ITEM_ID)){
    if (!(dealKey in deal)) continue;
    if (!state.rows.has(itemId)) continue;

    const v = deal[dealKey];
    const prev = state.rows.get(itemId).status || '';
    const next = (v === true) ? 'COMPLETE' : (v === false) ? 'PENDING' : '';

    if (prev !== next){
      state.rows.set(itemId, { ...state.rows.get(itemId), status: next });
      appendHistory({ timestamp: nowIso(), itemId, field: 'status', oldValue: prev, newValue: next });
    }
  }

  syncDomFromState();

  const coverage = computeAutofillCoverage(deal);
  setAutofillBanner(
    `Autofill coverage — mapped found: ${coverage.found.length}, mapped missing: ${coverage.missing.length}. ` +
    `Missing keys: ${coverage.missing.slice(0, 12).join(', ')}${coverage.missing.length > 12 ? '…' : ''} ` +
    `| Unmapped deal keys: ${coverage.unmapped.slice(0, 12).join(', ')}${coverage.unmapped.length > 12 ? '…' : ''}`
  );

  scheduleAutosave();
}

// ---------- AI Autofill demo with preview ----------
function openModal(id){ const m = $(id); if (m) m.setAttribute('aria-hidden','false'); }
function closeModal(id){ const m = $(id); if (m) m.setAttribute('aria-hidden','true'); }

function buildAiProposal(){
  const itemIds = [...new Set(Object.values(DEAL_KEY_TO_ITEM_ID))].filter(id => state.rows.has(id));
  const targetCount = Math.min(itemIds.length, 15 + Math.floor(Math.random() * 11));
  const shuffled = itemIds.slice().sort(() => Math.random() - 0.5);
  const chosen = shuffled.slice(0, targetCount);

  const proposed = new Map();
  const changes = [];

  for (const id of chosen){
    const oldStatus = state.rows.get(id).status || '';
    const newStatus = Math.random() < 0.75 ? 'COMPLETE' : 'PENDING';
    if (oldStatus !== newStatus){
      proposed.set(id, newStatus);
      changes.push({ itemId: id, label: state.rows.get(id).label, oldStatus, newStatus });
    }
  }

  return { proposed, changes };
}

async function aiAutofillDemo(){
  setMeta('AI Autofill (Demo)…');
  const out = $('aiOutput');
  if (out) out.textContent = 'Simulating extraction…';
  openModal('aiModal');
  state.aiPreview = null;

  await new Promise(r => setTimeout(r, 1000 + Math.floor(Math.random() * 1000)));

  const proposal = buildAiProposal();
  state.aiPreview = proposal;

  if (!out) return;
  if (proposal.changes.length === 0){
    out.textContent = 'No changes proposed.';
    return;
  }

  out.textContent = proposal.changes.map(c => `- ${c.itemId}: ${c.oldStatus || '(blank)'} -> ${c.newStatus}`).join('\n');
}

function applyAiProposal(){
  const proposal = state.aiPreview;
  if (!proposal) return;

  for (const [id, newStatus] of proposal.proposed.entries()){
    const prev = state.rows.get(id).status || '';
    if (prev !== newStatus){
      state.rows.set(id, { ...state.rows.get(id), status: newStatus });
      appendHistory({ timestamp: nowIso(), itemId: id, field: 'status', oldValue: prev, newValue: newStatus });
    }
  }

  syncDomFromState();
  scheduleAutosave();
  setMeta('AI Autofill applied');
  closeModal('aiModal');
}

// ---------- Audit + History ----------
function computeChecklistAudit(){
  const expectedIds = state.checklist.rows.map(r => r.id);
  const renderedEls = [...document.querySelectorAll('tr[data-row-id]')];
  const renderedIds = renderedEls.map(el => el.getAttribute('data-row-id') || '');

  const renderedSet = new Set(renderedIds);
  const missing = expectedIds.filter(id => !renderedSet.has(id));

  const counts = new Map();
  for (const id of renderedIds) counts.set(id, (counts.get(id) || 0) + 1);
  const duplicates = [...counts.entries()].filter(([,c]) => c > 1).map(([id, c]) => `${id} (x${c})`);

  return { expected: expectedIds.length, rendered: renderedEls.length, missing, duplicates };
}

function openAudit(){
  const out = $('auditOutput');
  const a = computeChecklistAudit();
  out.textContent =
    `Total items expected (JSON): ${a.expected}\n` +
    `Total items rendered (DOM): ${a.rendered}\n\n` +
    `Missing item IDs (${a.missing.length}):\n` +
    (a.missing.length ? a.missing.map(x => `- ${x}`).join('\n') : '- (none)') +
    `\n\nDuplicate DOM item IDs (${a.duplicates.length}):\n` +
    (a.duplicates.length ? a.duplicates.map(x => `- ${x}`).join('\n') : '- (none)') +
    `\n`;
  openModal('auditModal');
}

function renderHistory(){
  const out = $('historyOutput');
  const filter = ($('historyFilter')?.value || '').toLowerCase().trim();
  const field = $('historyField')?.value || '';

  const entries = loadHistory(state.activeDealId || 'local-demo');
  const filtered = entries.filter(e => {
    if (filter && !String(e.itemId || '').toLowerCase().includes(filter)) return false;
    if (field && String(e.field || '') !== field) return false;
    return true;
  });

  out.textContent = filtered.slice(0, 500).map(e => {
    return `${e.timestamp}  ${e.itemId}  ${e.field}  ${String(e.oldValue)} -> ${String(e.newValue)}`;
  }).join('\n') || '(no entries)';
}

function openHistory(){
  renderHistory();
  openModal('historyModal');
}

function exportHistory(){
  const dealId = state.activeDealId || 'local-demo';
  const entries = loadHistory(dealId);
  downloadText(`history.${dealId}.json`, JSON.stringify({ dealId, exportedAt: nowIso(), entries }, null, 2), 'application/json');
}

function clearHistory(){
  if (!LS) return;
  const dealId = state.activeDealId || 'local-demo';
  LS.removeItem(historyStorageKey(dealId));
  renderHistory();
}

// ---------- Bulk actions ----------
function markSection(section, status){
  if (!section) return;
  for (const def of state.checklist.rows){
    if (def.section !== section) continue;
    const prev = state.rows.get(def.id).status || '';
    if (prev !== status){
      state.rows.set(def.id, { ...state.rows.get(def.id), status });
      appendHistory({ timestamp: nowIso(), itemId: def.id, field: 'status', oldValue: prev, newValue: status });
    }
  }
  syncDomFromState();
  scheduleAutosave();
}

function markAll(status){
  for (const def of state.checklist.rows){
    const prev = state.rows.get(def.id).status || '';
    if (prev !== status){
      state.rows.set(def.id, { ...state.rows.get(def.id), status });
      appendHistory({ timestamp: nowIso(), itemId: def.id, field: 'status', oldValue: prev, newValue: status });
    }
  }
  syncDomFromState();
  scheduleAutosave();
}

// ---------- Import/export ----------
function exportChecklistJson(){
  const dealId = state.activeDealId || 'local-demo';
  const meta = state.deals.find(d => d.dealId === dealId);

  const rows = [];
  for (const def of state.checklist.rows){
    const r = state.rows.get(def.id);
    rows.push({ id: r.id, section: r.section, label: r.label, status: r.status || '', check: computeCheckResult(r.status || ''), comments: r.comments || '' });
  }

  const payload = { dealId, dealName: meta?.name || null, exportedAt: nowIso(), checklistVersion: state.checklistVersionHash, rows };
  downloadText(`credit-checklist.${dealId}.json`, JSON.stringify(payload, null, 2), 'application/json');
}

function exportChecklistCsv(){
  const rows = [['Section','Label','Status','Check','Comments']];
  for (const def of state.checklist.rows){
    const r = state.rows.get(def.id);
    rows.push([r.section, r.label, r.status || '', computeCheckResult(r.status || ''), r.comments || '']);
  }
  downloadText(`credit-checklist.${state.activeDealId || 'local-demo'}.csv`, toCsv(rows), 'text/csv');
}

async function importChecklistJsonFile(file){
  const text = await file.text();
  const payload = JSON.parse(text);
  if (!payload || !Array.isArray(payload.rows)) throw new Error('Invalid import payload.');

  const dealId = payload.dealId || randomId('deal');
  const name = payload.dealName || dealId;
  upsertDealMeta({ dealId, name, createdAt: nowIso(), updatedAt: nowIso() });

  loadDealIntoWorkspace(dealId);

  for (const row of payload.rows){
    if (!row || !row.id || !state.rows.has(row.id)) continue;
    const current = state.rows.get(row.id);
    const status = typeof row.status === 'string' ? row.status : current.status;
    const comments = typeof row.comments === 'string' ? row.comments : current.comments;
    state.rows.set(row.id, { ...current, status, comments });
  }

  syncDomFromState();
  scheduleAutosave();
  setMeta('Imported checklist JSON');
}

// ---------- Keyboard navigation ----------
function initKeyboardNav(){
  document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd+S
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's'){
      e.preventDefault();
      saveChecklistState({ silent: false });
      return;
    }

    // Esc closes modals
    if (e.key === 'Escape'){
      for (const id of ['auditModal','historyModal','aiModal']){
        const m = $(id);
        if (m && m.getAttribute('aria-hidden') === 'false'){
          closeModal(id);
          e.preventDefault();
          return;
        }
      }
    }

    const active = document.activeElement;
    if (!active) return;

    const isStatus = active.classList.contains('status-select');
    const isComment = active.classList.contains('comment-input');
    if (!isStatus && !isComment) return;

    const tr = active.closest('tr[data-row-id]');
    if (!tr) return;

    const rows = [...document.querySelectorAll('tr[data-row-id]')];
    const rowIndex = rows.indexOf(tr);
    if (rowIndex < 0) return;

    const focusCell = (rIdx, which) => {
      const t = rows[Math.max(0, Math.min(rows.length - 1, rIdx))];
      if (!t) return;
      const target = which === 'status' ? t.querySelector('select.status-select') : t.querySelector('input.comment-input');
      if (target){ target.focus(); target.select?.(); }
    };

    if (e.key === 'ArrowUp'){ e.preventDefault(); focusCell(rowIndex - 1, isStatus ? 'status' : 'comment'); return; }
    if (e.key === 'ArrowDown'){ e.preventDefault(); focusCell(rowIndex + 1, isStatus ? 'status' : 'comment'); return; }
    if (e.key === 'ArrowLeft'){ e.preventDefault(); focusCell(rowIndex, 'status'); return; }
    if (e.key === 'ArrowRight'){ e.preventDefault(); focusCell(rowIndex, 'comment'); return; }

    if (e.key === 'Enter'){
      e.preventDefault();
      if (isStatus) focusCell(rowIndex, 'comment');
      else focusCell(rowIndex + 1, 'status');
    }
  });
}

// ---------- Loading external JSON (best-effort) ----------
async function loadChecklistTemplate(){
  try{
    const res = await fetch(CHECKLIST_URL, { cache: 'no-store' });
    if (!res.ok) throw new Error('fetch failed');
    const j = await res.json();
    if (!j || !Array.isArray(j.rows) || j.rows.length === 0) throw new Error('bad json');
    return j;
  } catch (_err){
    return FALLBACK_CHECKLIST;
  }
}

async function loadSampleDeal(){
  try{
    const res = await fetch(SAMPLE_DEAL_URL, { cache: 'no-store' });
    if (!res.ok) throw new Error('fetch failed');
    const j = await res.json();
    if (!j || typeof j !== 'object') throw new Error('bad json');
    return j;
  } catch (_err){
    return FALLBACK_DEAL;
  }
}

function initTabs(){
  document.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => {
      const name = btn.getAttribute('data-sheet') || '';
      if (name !== 'Credit Checklist'){
        alert(`"${name}" is a placeholder tab. Only "Credit Checklist" is implemented.`);
      }
    });
  });
}

// ---------- UI wiring ----------
function wireUi(){
  $('dealSearch').addEventListener('input', renderDealList);

  $('btnNewDeal').addEventListener('click', newDeal);
  $('btnDuplicateDeal').addEventListener('click', duplicateDeal);
  $('btnDeleteDeal').addEventListener('click', deleteDeal);

  $('btnAddAttachment').addEventListener('click', () => $('attachmentsInput').click());
  $('attachmentsInput').addEventListener('change', (e) => {
    const files = [...(e.target.files || [])];
    if (files.length) addAttachments(files);
    e.target.value = '';
  });
  $('btnClearAttachments').addEventListener('click', () => {
    if (confirm('Clear attachment metadata for this deal?')) clearAttachmentMetadata();
  });

  $('btnUploadDeal').addEventListener('click', () => $('dealFileInput').click());
  $('dealFileInput').addEventListener('change', async (e) => {
    const file = e.target.files && e.target.files[0];
    try{
      if (file){
        const text = await file.text();
        applyDealAutofill(JSON.parse(text));
      }
    } catch (err){
      alert('Could not read deal JSON file.');
      console.error(err);
    } finally {
      e.target.value = '';
    }
  });

  $('btnLoadSample').addEventListener('click', async () => {
    const deal = await loadSampleDeal();
    applyDealAutofill(deal);
  });

  $('btnAiAutofill').addEventListener('click', () => aiAutofillDemo().catch(console.error));
  $('btnAiApply').addEventListener('click', applyAiProposal);
  $('btnAiCancel').addEventListener('click', () => closeModal('aiModal'));
  $('btnAiClose').addEventListener('click', () => closeModal('aiModal'));

  $('btnSave').addEventListener('click', () => saveChecklistState({ silent: false }));

  $('btnExportChecklist').addEventListener('click', exportChecklistJson);
  $('btnExportCsv').addEventListener('click', exportChecklistCsv);

  $('btnImportChecklist').addEventListener('click', () => $('checklistImportInput').click());
  $('checklistImportInput').addEventListener('change', async (e) => {
    const file = e.target.files && e.target.files[0];
    try{
      if (file) await importChecklistJsonFile(file);
    } catch (err){
      alert('Could not import checklist JSON.');
      console.error(err);
    } finally {
      e.target.value = '';
    }
  });

  $('btnAudit').addEventListener('click', openAudit);
  $('btnAuditClose').addEventListener('click', () => closeModal('auditModal'));
  $('auditModal').addEventListener('click', (e) => { if (e.target.id === 'auditModal') closeModal('auditModal'); });

  $('btnHistory').addEventListener('click', openHistory);
  $('btnHistoryClose').addEventListener('click', () => closeModal('historyModal'));
  $('historyModal').addEventListener('click', (e) => { if (e.target.id === 'historyModal') closeModal('historyModal'); });
  $('historyFilter').addEventListener('input', renderHistory);
  $('historyField').addEventListener('change', renderHistory);
  $('btnHistoryExport').addEventListener('click', exportHistory);
  $('btnHistoryClear').addEventListener('click', () => { if (confirm('Clear history for this deal?')) clearHistory(); });

  $('btnMarkSectionComplete').addEventListener('click', () => {
    const section = $('sectionSelect').value;
    if (!section) return alert('Select a section first.');
    markSection(section, 'COMPLETE');
  });
  $('btnMarkSectionPending').addEventListener('click', () => {
    const section = $('sectionSelect').value;
    if (!section) return alert('Select a section first.');
    markSection(section, 'PENDING');
  });
  $('btnMarkAllComplete').addEventListener('click', () => markAll('COMPLETE'));
  $('btnMarkAllPending').addEventListener('click', () => markAll('PENDING'));
  $('btnReset').addEventListener('click', () => { if (confirm('Reset to defaults and clear comments?')) resetToDefaults(); });
  $('btnClearSaved').addEventListener('click', () => { if (confirm('Clear saved checklist state for this deal?')) clearSavedState(); });
}

// ---------- Boot ----------
async function main(){
  runCheckColumnTests();
  initKeyboardNav();
  initTabs();

  state.checklist = await loadChecklistTemplate();
  state.checklistVersionHash = fnv1a32(JSON.stringify(state.checklist.rows));

  state.deals = loadDealsIndex();
  if (!state.deals.find(d => d.dealId === 'local-demo')){
    state.deals.unshift({ dealId: 'local-demo', name: 'Local Demo', createdAt: nowIso(), updatedAt: nowIso() });
    saveDealsIndex();
  }

  wireUi();

  const firstDeal = state.deals[0]?.dealId || 'local-demo';
  loadDealIntoWorkspace(firstDeal);
}

main().catch((err) => {
  console.error(err);
  alert('Failed to load app.');
});
