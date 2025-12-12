(() => {
  'use strict';

  // ------------------------------
  // In-memory state (no storage)
  // ------------------------------
  const appState = {
    mode: 'offline_mock',
    confidenceThreshold: 0.75,
    file: null,
    rawText: '',
    extractionResult: null,
    suggestions: new Map(),
    issues: [],
    lastFocusId: null,
  };

  // Explicitly expose for debugging/inspection (still memory only).
  window.privateCreditFormData = appState;

  // ------------------------------
  // DOM helpers
  // ------------------------------
  const $ = (id) => document.getElementById(id);

  function nowStamp(){
    return new Date().toISOString().replace('T',' ').replace('Z','Z');
  }

  function log(level, message){
    const el = $('log');
    const tag = String(level || 'info').toUpperCase();
    const line = `[${nowStamp()}] ${tag}: ${message}`;
    el.value = (el.value ? (el.value + '\n') : '') + line;
    el.scrollTop = el.scrollHeight;
    if (tag === 'ERROR') console.error(line);
    else if (tag === 'WARN') console.warn(line);
    else console.log(line);
  }

  function safeText(input){
    return String(input ?? '')
      .replace(/[\u0000-\u001f\u007f]/g, ' ')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;');
  }

  function isBlank(v){
    return v === null || v === undefined || String(v).trim() === '';
  }

  function clamp01(n){
    const x = Number(n);
    if (!Number.isFinite(x)) return 0;
    return Math.max(0, Math.min(1, x));
  }

  function parseNumberLoose(v){
    if (v === null || v === undefined) return null;
    if (typeof v === 'number' && Number.isFinite(v)) return v;
    const s = String(v).trim();
    if (!s) return null;

    const lowered = s.toLowerCase();
    const hasPercent = lowered.includes('%');

    let mult = 1;
    if (/\b\d+(?:\.\d+)?\s*(?:mm|m)\b/.test(lowered) || /\b(million|millions)\b/.test(lowered)) mult = 1_000_000;
    if (/\b\d+(?:\.\d+)?\s*(?:bn|b)\b/.test(lowered) || /\b(billion|billions)\b/.test(lowered)) mult = 1_000_000_000;
    if (/\b\d+(?:\.\d+)?\s*k\b/.test(lowered) || /\b(thousand|thousands)\b/.test(lowered)) mult = 1_000;

    const cleaned = lowered
      .replace(/[$,]/g,'')
      .replace(/\bmm\b/g,'')
      .replace(/\bbn\b/g,'')
      .replace(/\bk\b/g,'')
      .replace(/\b(million|millions|billion|billions|thousand|thousands)\b/g,'')
      .replace(/%/g,'')
      .trim();

    const num = Number(cleaned);
    if (!Number.isFinite(num)) return null;
    return hasPercent ? num : (num * mult);
  }

  function detectUnitsScale(text){
    const t = String(text || '').toLowerCase();
    if (/(\$\s*in\s*millions|\$\s*\(in\s*millions\)|in\s*millions|\$mm\b|\bmm\b)/.test(t)) return 1_000_000;
    if (/(\$\s*in\s*thousands|\$\s*\(in\s*thousands\)|in\s*thousands|\$000s|\b000s\b)/.test(t)) return 1_000;
    return 1;
  }

  function detectCurrency(text){
    const t = String(text || '');
    if (/\bUSD\b/.test(t)) return 'USD';
    if (/\bAUD\b/.test(t)) return 'AUD';
    if (/\bCAD\b/.test(t)) return 'CAD';
    if (/\bEUR\b/.test(t)) return 'EUR';
    if (/\bGBP\b/.test(t)) return 'GBP';
    // Symbol-based (weak)
    if (t.includes('$')) return 'USD';
    return null;
  }

  function findFirstPageRef(text, index){
    // Best-effort: find last [pN] marker before match.
    const before = text.slice(0, Math.max(0, index));
    const matches = [...before.matchAll(/\[p(\d+)\]/gi)];
    if (matches.length === 0) return ['p?'];
    const last = matches[matches.length - 1];
    return [`p${last[1]}`];
  }

  function makeField(value, confidence, evidence, pageRefs){
    const conf = clamp01(confidence);
    if (value === null || value === undefined || conf <= 0){
      return { value: null, confidence: 0, evidence: null, page_refs: [] };
    }
    const ev = (typeof evidence === 'string' && evidence.trim()) ? evidence.slice(0, 240) : null;
    const pr = Array.isArray(pageRefs) ? pageRefs.map(String) : [];
    if (!ev || pr.length === 0){
      return { value: null, confidence: 0, evidence: null, page_refs: [] };
    }
    return { value, confidence: conf, evidence: ev, page_refs: pr };
  }

  // ------------------------------
  // Static sample (mocked data)
  // ------------------------------
  const SAMPLE_CIM_TEXT = `
[p1] CONFIDENTIAL INFORMATION MEMORANDUM
[p1] Borrower: ExampleCo Holdings, LLC
[p1] Industry: Vertical SaaS (B2B)
[p2] All financial information is presented in USD and $ in millions unless otherwise noted.
[p8] Summary Financials (LTM ended 2025-06-30)
[p8] Revenue: $125.0mm
[p8] EBITDA: $32.5mm
[p8] EBITDA Margin: 26.0%
[p9] Capitalization
[p9] Total Debt: $150.0mm
[p9] Net Debt: $140.0mm
[p9] Total Leverage: 4.6x
[p9] Net Leverage: 4.3x
[p10] Proposed Debt Structure
[p10] Unitranche (Senior Secured) Commitment: $175.0mm Drawn: $150.0mm Maturity: 2030-07-01 Margin: 550 bps
[p11] Coverage Metrics
[p11] Interest Coverage: 2.1x
[p11] Fixed Charge Coverage: 1.6x
`.trim();

  // ------------------------------
  // Mock extractor (deterministic)
  // ------------------------------
  function mockExtractFromText(documentTitle, rawText){
    const text = String(rawText || '');
    const lower = text.toLowerCase();

    const unitsScale = detectUnitsScale(lower);
    const currency = detectCurrency(text);

    const matchLabelValue = (labelRegex, valueRegex) => {
      const re = new RegExp(labelRegex.source + '[^\n]{0,80}?' + valueRegex.source, 'i');
      const m = re.exec(text);
      if (!m) return null;
      const full = m[0];
      // last capture group should be the value
      const valStr = m[m.length - 1];
      const idx = m.index;
      return { full, valStr, idx };
    };

    const amountRe = /([\$€£]?\s*\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*(?:mm|m|bn|b|k)?)/i;
    const percentRe = /(\d{1,3}(?:\.\d+)?\s*%?)/i;
    const multipleRe = /(\d{1,2}(?:\.\d+)?\s*x)/i;
    const dateRe = /(\d{4}-\d{2}-\d{2})/i;
    const bpsRe = /(\d{2,4})\s*bps/i;

    // Borrower
    let borrower = null;
    {
      const re = /(Borrower|Company|Issuer)\s*:\s*([^\n]{2,90})/i;
      const m = re.exec(text);
      if (m){
        borrower = { value: m[2].trim(), confidence: 0.9, evidence: m[0].trim(), page_refs: findFirstPageRef(text, m.index) };
      }
    }

    // Industry
    let industry = null;
    {
      const re = /Industry\s*:\s*([^\n]{2,90})/i;
      const m = re.exec(text);
      if (m){
        industry = { value: m[1].trim(), confidence: 0.85, evidence: m[0].trim(), page_refs: findFirstPageRef(text, m.index) };
      }
    }

    // LTM As-of date
    let asOfDate = null;
    {
      const re = /(LTM\s*(?:ended)?\s*)(\d{4}-\d{2}-\d{2})/i;
      const m = re.exec(text);
      if (m){
        asOfDate = m[2];
      }
    }

    // Revenue, EBITDA, Margin
    const revMatch = matchLabelValue(/\b(Revenue|Sales)\b\s*[:\-]?\s*/i, amountRe);
    const ebitdaMatch = matchLabelValue(/\bEBITDA\b\s*[:\-]?\s*/i, amountRe);
    const marginMatch = matchLabelValue(/\bEBITDA\s*Margin\b\s*[:\-]?\s*/i, percentRe);

    const revenueVal = revMatch ? parseNumberLoose(revMatch.valStr) : null;
    const ebitdaVal = ebitdaMatch ? parseNumberLoose(ebitdaMatch.valStr) : null;
    const marginVal = marginMatch ? parseNumberLoose(marginMatch.valStr) : null;

    const revenueField = revMatch ? makeField(
      revenueVal,
      revenueVal !== null ? 0.9 : 0,
      revMatch.full.trim(),
      findFirstPageRef(text, revMatch.idx)
    ) : makeField(null, 0, null, []);

    const ebitdaField = ebitdaMatch ? makeField(
      ebitdaVal,
      ebitdaVal !== null ? 0.9 : 0,
      ebitdaMatch.full.trim(),
      findFirstPageRef(text, ebitdaMatch.idx)
    ) : makeField(null, 0, null, []);

    let marginField = makeField(null, 0, null, []);
    if (marginMatch){
      marginField = makeField(
        marginVal,
        marginVal !== null ? 0.85 : 0,
        marginMatch.full.trim(),
        findFirstPageRef(text, marginMatch.idx)
      );
    } else if (revenueVal && ebitdaVal && revenueVal !== 0){
      const implied = (ebitdaVal / revenueVal) * 100;
      const ev = `Computed EBITDA margin ≈ ${(implied).toFixed(1)}% from Revenue and EBITDA lines.`;
      marginField = makeField(Number(implied.toFixed(1)), 0.6, ev, ['p?']);
    }

    // Total/Net debt
    const totalDebtMatch = matchLabelValue(/\bTotal\s*Debt\b\s*[:\-]?\s*/i, amountRe);
    const netDebtMatch = matchLabelValue(/\bNet\s*Debt\b\s*[:\-]?\s*/i, amountRe);

    const totalDebtField = totalDebtMatch ? makeField(
      parseNumberLoose(totalDebtMatch.valStr),
      0.85,
      totalDebtMatch.full.trim(),
      findFirstPageRef(text, totalDebtMatch.idx)
    ) : makeField(null, 0, null, []);

    const netDebtField = netDebtMatch ? makeField(
      parseNumberLoose(netDebtMatch.valStr),
      0.85,
      netDebtMatch.full.trim(),
      findFirstPageRef(text, netDebtMatch.idx)
    ) : makeField(null, 0, null, []);

    // Leverage & coverage
    const totalLevMatch = matchLabelValue(/\bTotal\s*Leverage\b\s*[:\-]?\s*/i, multipleRe);
    const netLevMatch = matchLabelValue(/\bNet\s*Leverage\b\s*[:\-]?\s*/i, multipleRe);
    const icovMatch = matchLabelValue(/\bInterest\s*Coverage\b\s*[:\-]?\s*/i, multipleRe);
    const fccrMatch = matchLabelValue(/\bFixed\s*Charge\s*Coverage\b\s*[:\-]?\s*/i, multipleRe);

    const multipleToNum = (s) => {
      if (!s) return null;
      const m = /([0-9]+(?:\.[0-9]+)?)\s*x/i.exec(String(s));
      if (!m) return null;
      const n = Number(m[1]);
      return Number.isFinite(n) ? n : null;
    };

    const totalLevField = totalLevMatch ? makeField(multipleToNum(totalLevMatch.valStr), 0.85, totalLevMatch.full.trim(), findFirstPageRef(text, totalLevMatch.idx)) : makeField(null,0,null,[]);
    const netLevField = netLevMatch ? makeField(multipleToNum(netLevMatch.valStr), 0.85, netLevMatch.full.trim(), findFirstPageRef(text, netLevMatch.idx)) : makeField(null,0,null,[]);
    const icovField = icovMatch ? makeField(multipleToNum(icovMatch.valStr), 0.8, icovMatch.full.trim(), findFirstPageRef(text, icovMatch.idx)) : makeField(null,0,null,[]);
    const fccrField = fccrMatch ? makeField(multipleToNum(fccrMatch.valStr), 0.8, fccrMatch.full.trim(), findFirstPageRef(text, fccrMatch.idx)) : makeField(null,0,null,[]);

    // Debt tranche (single)
    let tranche = {
      tranche_name: makeField(null,0,null,[]),
      tranche_type: makeField(null,0,null,[]),
      commitment: makeField(null,0,null,[]),
      drawn_amount: makeField(null,0,null,[]),
      maturity_date: makeField(null,0,null,[]),
      rate_type: makeField(null,0,null,[]),
      base_rate: makeField(null,0,null,[]),
      margin_bps: makeField(null,0,null,[]),
      all_in_rate_bps: makeField(null,0,null,[])
    };

    {
      const trancheLineRe = /(Unitranche|Term\s*Loan\s*A|Revolver|Senior\s*Secured)[^\n]{0,200}/i;
      const m = trancheLineRe.exec(text);
      if (m){
        const line = m[0];
        const pageRefs = findFirstPageRef(text, m.index);
        tranche.tranche_name = makeField(m[1].trim(), 0.8, line.trim(), pageRefs);
        // tranche type in parens
        const typeM = /\(([^)]+)\)/.exec(line);
        if (typeM) tranche.tranche_type = makeField(typeM[1].trim(), 0.75, line.trim(), pageRefs);

        const commM = /Commitment\s*:\s*([^\s][^\n]{0,25})/i.exec(line);
        if (commM) tranche.commitment = makeField(parseNumberLoose(commM[1]), 0.8, line.trim(), pageRefs);

        const drawnM = /Drawn\s*:\s*([^\s][^\n]{0,25})/i.exec(line);
        if (drawnM) tranche.drawn_amount = makeField(parseNumberLoose(drawnM[1]), 0.8, line.trim(), pageRefs);

        const matM = new RegExp('Maturity\\s*:\\s*' + dateRe.source, 'i').exec(line);
        if (matM) tranche.maturity_date = makeField(matM[1], 0.85, line.trim(), pageRefs);

        const bpsM = bpsRe.exec(line);
        if (bpsM) tranche.margin_bps = makeField(Number(bpsM[1]), 0.85, line.trim(), pageRefs);
      }
    }

    const unitsField = makeField(unitsScale, (unitsScale !== 1 ? 0.8 : 0.5),
      unitsScale !== 1 ? 'Detected scale from text (e.g., “in millions/thousands”).' : 'Units scale not explicit; defaulted to 1.',
      ['p?']
    );

    const currencyField = currency ? makeField(currency, 0.75, `Detected currency token: ${currency}`, ['p?']) : makeField(null,0,null,[]);

    const result = {
      metadata: {
        document_title: documentTitle || 'Untitled',
        borrower_name: borrower ? makeField(borrower.value, borrower.confidence, borrower.evidence, borrower.page_refs) : makeField(null,0,null,[]),
        industry: industry ? makeField(industry.value, industry.confidence, industry.evidence, industry.page_refs) : makeField(null,0,null,[]),
        headquarters: makeField(null,0,null,[]),
        reporting_currency: currencyField,
        units_scale: unitsField,
      },
      financials: {
        income_statement: {
          periods: [
            {
              period_label: 'LTM',
              period_end_date: asOfDate || null,
              period_months: 12,
              revenue: revenueField,
              gross_profit: makeField(null,0,null,[]),
              operating_income: makeField(null,0,null,[]),
              ebitda: ebitdaField,
              ebitda_margin_percent: marginField,
              net_income: makeField(null,0,null,[]),
            }
          ]
        },
        balance_sheet: { periods: [ { period_label: null, period_end_date: null, cash: makeField(null,0,null,[]), total_assets: makeField(null,0,null,[]), total_liabilities: makeField(null,0,null,[]), total_debt: makeField(null,0,null,[]), shareholders_equity: makeField(null,0,null,[]) } ] },
        cash_flow: { periods: [ { period_label: null, period_end_date: null, cash_from_operations: makeField(null,0,null,[]), capex: makeField(null,0,null,[]), free_cash_flow: makeField(null,0,null,[]) } ] },
      },
      credit_metrics: {
        ltm: {
          as_of_date: asOfDate || null,
          leverage: {
            total_debt: totalDebtField,
            net_debt: netDebtField,
            total_leverage_x: totalLevField,
            net_leverage_x: netLevField,
          },
          coverage: {
            interest_coverage_x: icovField,
            fixed_charge_coverage_x: fccrField,
          }
        }
      },
      debt_structure: {
        tranches: [ tranche ]
      },
      quality: {
        notes: [
          'Offline mock extraction uses regex + heuristics (no network).',
          'Any field without explicit evidence remains null with confidence 0.'
        ],
        missing_critical_fields: []
      }
    };

    // Critical fields check
    const critical = [
      ['metadata.borrower_name', result.metadata.borrower_name],
      ['financials.income_statement.periods[0].revenue', result.financials.income_statement.periods[0].revenue],
      ['financials.income_statement.periods[0].ebitda', result.financials.income_statement.periods[0].ebitda],
    ];
    for (const [name, f] of critical){
      if (!f || f.value === null) result.quality.missing_critical_fields.push(name);
    }

    return result;
  }

  // ------------------------------
  // UI wiring
  // ------------------------------
  const fieldOrder = [
    'dealname','borrowername','industry','currency','unitsscale','asofdate',
    'periodend','periodmonths','revenue','ebitda','ebitdamargin','netincome',
    'tranchename','tranchetype','commitment','drawn','maturity','marginbps',
    'totaldebt','netdebt','totallev','netlev','interestcov','fccr'
  ];

  const validationRules = {
    text: (v, required) => required && isBlank(v) ? 'Required.' : null,
    currency_code: (v) => {
      if (isBlank(v)) return null;
      const s = String(v).trim().toUpperCase();
      if (!/^[A-Z]{3}$/.test(s)) return 'Use a 3-letter ISO code (e.g., USD).';
      return null;
    },
    units_scale: (v) => {
      if (isBlank(v)) return null;
      const n = Number(v);
      if (![1,1000,1000000].includes(n)) return 'Units scale must be 1, 1000, or 1000000.';
      return null;
    },
    date: (v) => {
      if (isBlank(v)) return null;
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(v))) return 'Use YYYY-MM-DD.';
      return null;
    },
    int: (v) => {
      if (isBlank(v)) return null;
      const n = Number(v);
      if (!Number.isInteger(n)) return 'Must be an integer.';
      return null;
    },
    money: (v, required) => {
      if (required && isBlank(v)) return 'Required.';
      if (isBlank(v)) return null;
      const n = parseNumberLoose(v);
      if (n === null) return 'Enter a numeric amount (supports 10m, 2,500,000).';
      return null;
    },
    percent: (v) => {
      if (isBlank(v)) return null;
      const n = parseNumberLoose(v);
      if (n === null) return 'Enter a numeric percent (e.g., 26.0).';
      if (n < -100 || n > 100) return 'Percent out of range.';
      return null;
    },
    multiple: (v) => {
      if (isBlank(v)) return null;
      const n = parseNumberLoose(v);
      if (n === null) return 'Enter a numeric multiple (e.g., 4.2).';
      if (n < 0 || n > 50) return 'Multiple out of range.';
      return null;
    },
    bps: (v) => {
      if (isBlank(v)) return null;
      const n = parseNumberLoose(v);
      if (n === null) return 'Enter basis points (e.g., 550).';
      if (n < 0 || n > 5000) return 'Bps out of range.';
      return null;
    }
  };

  function setActiveTab(tabId){
    document.querySelectorAll('.tab').forEach(b => b.setAttribute('aria-selected', String(b.getAttribute('data-tab') === tabId)));
    document.querySelectorAll('.section').forEach(s => s.classList.toggle('active', s.id === tabId));
  }

  document.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => setActiveTab(btn.getAttribute('data-tab')));
  });

  function focusField(fieldId){
    const el = $(fieldId);
    if (!el) return;
    const section = el.closest('.section');
    if (section && !section.classList.contains('active')) setActiveTab(section.id);
    el.focus();
    el.select?.();
    appState.lastFocusId = fieldId;
  }

  function validateField(el){
    const type = el.getAttribute('data-type') || 'text';
    const required = el.getAttribute('data-required') === 'true' || el.required;
    const fn = validationRules[type] || validationRules.text;
    const err = fn(el.value, required);

    const vEl = $(el.id + '-validation');
    el.classList.remove('valid','invalid');
    vEl.classList.remove('ok','error');

    if (err){
      el.classList.add('invalid');
      el.setAttribute('aria-invalid','true');
      vEl.textContent = err;
      vEl.classList.add('error');
    } else {
      el.removeAttribute('aria-invalid');
      const hasVal = !isBlank(el.value);
      const required2 = el.getAttribute('data-required') === 'true' || el.required;
      if (hasVal || !required2){
        if (hasVal) el.classList.add('valid');
        vEl.textContent = hasVal ? 'Looks good.' : '';
        if (hasVal) vEl.classList.add('ok');
      } else {
        vEl.textContent = '';
      }
    }
    return err;
  }

  let validateTimer = null;
  function requestValidateAll(){
    clearTimeout(validateTimer);
    validateTimer = setTimeout(validateAll, 280);
  }

  function validateAll(){
    const issues = [];

    for (const id of fieldOrder){
      const el = $(id);
      if (!el) continue;
      const err = validateField(el);
      if (err){
        const required = el.getAttribute('data-required') === 'true' || el.required;
        issues.push({
          severity: required ? 'error' : 'warn',
          fieldId: id,
          title: required ? 'Missing/invalid required field' : 'Invalid field',
          message: `${id}: ${err}`
        });
      }
    }

    // Consistency: implied margin
    const revenue = parseNumberLoose($('revenue')?.value);
    const ebitda = parseNumberLoose($('ebitda')?.value);
    const margin = parseNumberLoose($('ebitdamargin')?.value);
    if (Number.isFinite(revenue) && Number.isFinite(ebitda) && revenue !== 0){
      const implied = (ebitda / revenue) * 100;
      if (Number.isFinite(margin)){
        const diff = Math.abs(margin - implied);
        if (diff > 3){
          issues.push({
            severity:'warn',
            fieldId:'ebitdamargin',
            title:'EBITDA margin mismatch',
            message:`Implied margin is ${implied.toFixed(1)}%, entered is ${margin.toFixed(1)}% (Δ ${diff.toFixed(1)}%).`
          });
        }
      } else {
        issues.push({
          severity:'warn',
          fieldId:'ebitdamargin',
          title:'Missing EBITDA margin',
          message:`You can compute EBITDA margin ≈ ${implied.toFixed(1)}% from Revenue and EBITDA.`
        });
      }
    }

    appState.issues = issues;
    renderIssues();
    renderCompletion();
    renderReviewSummary();
    return issues;
  }

  function renderCompletion(){
    const requiredIds = fieldOrder.filter(id => {
      const el = $(id);
      return el && (el.getAttribute('data-required') === 'true' || el.required);
    });
    const filled = requiredIds.filter(id => !isBlank($(id).value)).length;
    const pct = requiredIds.length ? Math.round((filled / requiredIds.length) * 100) : 0;

    $('completion-percent').textContent = `${pct}%`;
    $('completion-bar').style.width = `${pct}%`;

    const dot = $('completion-dot');
    dot.classList.remove('success','warning','danger','primary');
    if (pct === 100) dot.classList.add('success');
    else if (pct >= 60) dot.classList.add('warning');
    else dot.classList.add('primary');
  }

  function renderIssues(){
    const issues = appState.issues || [];
    $('issues-count').textContent = String(issues.length);
    $('pillIssues').textContent = String(issues.length);

    const dot = $('issues-dot');
    dot.classList.remove('success','warning','danger');
    if (issues.length === 0) dot.classList.add('success');
    else if (issues.some(i => i.severity === 'error')) dot.classList.add('danger');
    else dot.classList.add('warning');

    const rail = $('issueRail');
    const body = $('issueRailBody');
    body.innerHTML = '';

    if (issues.length === 0){
      rail.style.display = 'none';
      return;
    }

    rail.style.display = 'block';
    issues.forEach((issue) => {
      const div = document.createElement('div');
      div.className = 'issue';
      div.setAttribute('role','button');
      div.tabIndex = 0;
      div.innerHTML = `
        <div class="title">
          <span>${safeText(issue.title)}</span>
          <span class="pill ${issue.severity === 'error' ? 'error' : 'warn'}">${issue.severity.toUpperCase()}</span>
        </div>
        <div class="meta">${safeText(issue.message)}</div>
      `;
      div.addEventListener('click', () => focusField(issue.fieldId));
      div.addEventListener('keydown', (e) => { if (e.key === 'Enter') focusField(issue.fieldId); });
      body.appendChild(div);
    });
  }

  let issueCursor = -1;
  function focusNextIssue(direction){
    const issues = appState.issues || [];
    if (issues.length === 0) return;
    if (issueCursor === -1) issueCursor = 0;
    else issueCursor += (direction || 1);
    if (issueCursor < 0) issueCursor = issues.length - 1;
    if (issueCursor >= issues.length) issueCursor = 0;
    const issue = issues[issueCursor];
    if (issue) focusField(issue.fieldId);
  }

  function focusNextMissingRequired(){
    for (const id of fieldOrder){
      const el = $(id);
      if (!el) continue;
      const required = el.getAttribute('data-required') === 'true' || el.required;
      if (!required) continue;
      if (isBlank(el.value)){
        focusField(id);
        return true;
      }
    }
    return false;
  }

  // Attach validation events
  for (const id of fieldOrder){
    const el = $(id);
    if (!el) continue;
    el.addEventListener('blur', () => { validateField(el); validateAll(); });
    el.addEventListener('input', () => { validateField(el); requestValidateAll(); });
    el.addEventListener('focus', () => { appState.lastFocusId = id; });
  }

  // ------------------------------
  // Confidence rendering + mapping
  // ------------------------------
  function confidenceLabel(c){
    if (c >= 0.85) return { label: 'High', color: 'rgba(40,167,69,.9)' };
    if (c >= 0.65) return { label: 'Med', color: 'rgba(255,193,7,.9)' };
    if (c > 0) return { label: 'Low', color: 'rgba(220,53,69,.9)' };
    return { label: '—', color: 'rgba(255,255,255,.25)' };
  }

  function renderConfidence(fieldId, conf){
    const root = $(fieldId + '-confidence');
    if (!root) return;
    root.innerHTML = '';
    const c = clamp01(conf);
    const meta = confidenceLabel(c);

    const bar = document.createElement('span');
    bar.className = 'bar';
    const fill = document.createElement('div');
    fill.style.width = `${Math.round(c * 100)}%`;
    fill.style.background = meta.color;
    bar.appendChild(fill);

    const txt = document.createElement('span');
    txt.textContent = `${meta.label} (${Math.round(c*100)}%)`;

    root.appendChild(bar);
    root.appendChild(txt);
  }

  function applyValue(fieldId, value){
    const el = $(fieldId);
    if (!el) return;
    if (el.tagName === 'SELECT') el.value = String(value);
    else el.value = value === null || value === undefined ? '' : String(value);
    validateField(el);
  }

  function mapExtractionToForm(result){
    appState.suggestions.clear();

    const threshold = Number(appState.confidenceThreshold || 0.75);
    const remember = (fieldId, fObj) => {
      if (!fObj || typeof fObj !== 'object') return;
      const conf = clamp01(fObj.confidence);
      if (fObj.value === null || conf <= 0) return;
      appState.suggestions.set(fieldId, { value: fObj.value, confidence: conf, evidence: fObj.evidence, page_refs: fObj.page_refs });

      // Auto-apply: high confidence AND evidence present
      const hasEvidence = typeof fObj.evidence === 'string' && fObj.evidence.trim() && Array.isArray(fObj.page_refs) && fObj.page_refs.length;
      if (conf >= threshold && hasEvidence) applyValue(fieldId, fObj.value);
      renderConfidence(fieldId, conf);
    };

    remember('borrowername', result?.metadata?.borrower_name);
    remember('industry', result?.metadata?.industry);
    remember('currency', result?.metadata?.reporting_currency);

    // units_scale is a field object
    if (result?.metadata?.units_scale) remember('unitsscale', result.metadata.units_scale);

    const income0 = result?.financials?.income_statement?.periods?.[0];
    if (income0){
      if (income0.period_end_date) applyValue('periodend', income0.period_end_date);
      if (income0.period_months) applyValue('periodmonths', String(income0.period_months));
      remember('revenue', income0.revenue);
      remember('ebitda', income0.ebitda);
      remember('ebitdamargin', income0.ebitda_margin_percent);
      remember('netincome', income0.net_income);
    }

    if (result?.credit_metrics?.ltm?.as_of_date) applyValue('asofdate', result.credit_metrics.ltm.as_of_date);
    const lev = result?.credit_metrics?.ltm?.leverage;
    const cov = result?.credit_metrics?.ltm?.coverage;
    remember('totaldebt', lev?.total_debt);
    remember('netdebt', lev?.net_debt);
    remember('totallev', lev?.total_leverage_x);
    remember('netlev', lev?.net_leverage_x);
    remember('interestcov', cov?.interest_coverage_x);
    remember('fccr', cov?.fixed_charge_coverage_x);

    const tr0 = result?.debt_structure?.tranches?.[0];
    if (tr0){
      remember('tranchename', tr0.tranche_name);
      remember('tranchetype', tr0.tranche_type);
      remember('commitment', tr0.commitment);
      remember('drawn', tr0.drawn_amount);
      remember('maturity', tr0.maturity_date);
      remember('marginbps', tr0.margin_bps);
    }

    renderCompletion();
    validateAll();
  }

  function renderReviewSummary(){
    const out = [];
    const issues = appState.issues || [];
    const errors = issues.filter(i => i.severity === 'error');
    const warns = issues.filter(i => i.severity !== 'error');

    out.push(`Completion: ${$('completion-percent').textContent}`);
    out.push(`Errors: ${errors.length}`);
    out.push(`Warnings: ${warns.length}`);

    const threshold = Number(appState.confidenceThreshold || 0.75);
    const needsReview = [];
    for (const [fieldId, s] of appState.suggestions.entries()){
      const autoApplied = s.confidence >= threshold && s.evidence && (s.page_refs||[]).length;
      if (!autoApplied){
        needsReview.push(`${fieldId}: suggested=${String(s.value)} (conf ${Math.round(s.confidence*100)}%)`);
      }
    }
    if (needsReview.length){
      out.push('');
      out.push('Suggestions requiring review:');
      out.push(...needsReview.slice(0, 30));
    }

    $('reviewSummary').value = out.join('\n');
  }

  // ------------------------------
  // Command palette
  // ------------------------------
  const cmdModal = $('cmdModal');
  const cmdInput = $('cmdInput');
  const cmdList = $('cmdList');
  let cmdItems = [];
  let cmdIndex = 0;

  const baseCommands = [
    { key: 'extract', name: 'Extract (Mock)', hint: 'Run offline mock extraction', action: () => runExtraction() },
    { key: 'validate', name: 'Validate', hint: 'Re-run validation rules', action: () => validateAll() },
    { key: 'export', name: 'Export', hint: 'Download JSON + CSV', action: () => exportData() },
    { key: 'clear', name: 'Clear', hint: 'Clear all data (memory only)', action: () => clearAll() },
  ];

  function buildCommands(){
    const fieldCmds = fieldOrder.map(id => ({
      key: `field:${id}`,
      name: `Jump: ${id}`,
      hint: 'Focus this field',
      action: () => focusField(id)
    }));
    return [...baseCommands, ...fieldCmds];
  }

  const allCommands = buildCommands();

  function openCmdPalette(){
    cmdModal.setAttribute('aria-hidden','false');
    cmdInput.value = '';
    cmdIndex = 0;
    renderCmdList('');
    setTimeout(() => cmdInput.focus(), 0);
  }

  function closeCmdPalette(){
    cmdModal.setAttribute('aria-hidden','true');
    if (appState.lastFocusId) focusField(appState.lastFocusId);
  }

  function renderCmdList(query){
    const q = String(query || '').trim().toLowerCase();
    cmdItems = allCommands
      .filter(c => !q || c.key.toLowerCase().includes(q) || c.name.toLowerCase().includes(q))
      .slice(0, 40);

    cmdIndex = Math.min(cmdIndex, Math.max(0, cmdItems.length - 1));

    cmdList.innerHTML = '';
    cmdItems.forEach((c, i) => {
      const div = document.createElement('div');
      div.className = 'cmditem';
      div.innerHTML = `
        <div class="left">
          <strong>${safeText(c.name)}</strong>
          <span>${safeText(c.hint || '')}</span>
        </div>
        <span class="kbd">Enter</span>
      `;
      if (i === cmdIndex) div.style.outline = '2px solid rgba(44,90,160,.75)';
      div.addEventListener('click', () => { closeCmdPalette(); c.action(); });
      cmdList.appendChild(div);
    });
  }

  $('btn-cmd').addEventListener('click', openCmdPalette);
  $('cmdClose').addEventListener('click', closeCmdPalette);
  cmdInput.addEventListener('input', () => renderCmdList(cmdInput.value));

  cmdModal.addEventListener('keydown', (e) => {
    if (cmdModal.getAttribute('aria-hidden') === 'true') return;
    if (e.key === 'Escape'){ e.preventDefault(); closeCmdPalette(); return; }
    if (e.key === 'ArrowDown'){ e.preventDefault(); cmdIndex = Math.min(cmdIndex + 1, cmdItems.length - 1); renderCmdList(cmdInput.value); return; }
    if (e.key === 'ArrowUp'){ e.preventDefault(); cmdIndex = Math.max(cmdIndex - 1, 0); renderCmdList(cmdInput.value); return; }
    if (e.key === 'Enter'){
      e.preventDefault();
      const item = cmdItems[cmdIndex];
      if (item){ closeCmdPalette(); item.action(); }
    }
  });

  // ------------------------------
  // Keyboard shortcuts
  // ------------------------------
  document.addEventListener('keydown', (e) => {
    const isCmdPaletteOpen = cmdModal.getAttribute('aria-hidden') === 'false';
    if (isCmdPaletteOpen) return;

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k'){
      e.preventDefault();
      openCmdPalette();
      return;
    }

    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter'){
      e.preventDefault();
      runExtraction();
      return;
    }

    if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'Enter'){
      e.preventDefault();
      if (!focusNextMissingRequired()) log('info', 'No missing required fields.');
      return;
    }

    if (e.key === 'F8'){
      e.preventDefault();
      validateAll();
      focusNextIssue(e.shiftKey ? -1 : 1);
      return;
    }

    // Enter / Shift+Enter navigation (skip textareas)
    const active = document.activeElement;
    if (!active) return;
    if (active.tagName === 'TEXTAREA') return;
    if (!(active.tagName === 'INPUT' || active.tagName === 'SELECT')) return;

    if (e.key === 'Enter'){
      e.preventDefault();
      const currentId = active.id;
      const idx = fieldOrder.indexOf(currentId);
      const nextIdx = idx < 0 ? 0 : idx + (e.shiftKey ? -1 : 1);
      const clamped = Math.max(0, Math.min(fieldOrder.length - 1, nextIdx));
      focusField(fieldOrder[clamped]);
    }
  });

  // ------------------------------
  // Ingestion: offline file stub
  // ------------------------------
  $('fileInput').addEventListener('change', async (e) => {
    const f = e.target.files && e.target.files[0];
    appState.file = f || null;

    const vEl = $('fileInput-validation');
    vEl.classList.remove('ok','error');

    if (!f){
      vEl.textContent = 'Upload a PDF, Excel, CSV, or TXT.';
      return;
    }

    if (f.size > 100 * 1024 * 1024){
      vEl.textContent = 'File too large (max 100MB).';
      vEl.classList.add('error');
      return;
    }

    const name = f.name.toLowerCase();
    const isTexty = name.endsWith('.txt') || name.endsWith('.csv');

    if (isTexty){
      const txt = await f.text();
      $('rawText').value = txt;
      appState.rawText = txt;
      vEl.textContent = 'Loaded text from file into Raw text.';
      vEl.classList.add('ok');
      log('info', `Loaded ${f.name} into raw text (${txt.length.toLocaleString()} chars).`);
      return;
    }

    // PDF/XLSX parsing is intentionally stubbed for offline MVP.
    vEl.textContent = 'Binary file uploaded. Offline parsing is stubbed — paste extracted text into Raw text.';
    vEl.classList.add('ok');
    log('warn', `Offline parsing stub: ${f.name}. Use copy/paste text for now.`);
  });

  // ------------------------------
  // Extraction flow (mock)
  // ------------------------------
  $('confidenceThreshold').addEventListener('change', (e) => {
    appState.confidenceThreshold = Number(e.target.value);
  });

  $('mode').addEventListener('change', (e) => {
    appState.mode = e.target.value;
  });

  $('rawText').addEventListener('input', (e) => {
    appState.rawText = e.target.value;
  });

  $('btn-load-sample').addEventListener('click', () => {
    $('rawText').value = SAMPLE_CIM_TEXT;
    appState.rawText = SAMPLE_CIM_TEXT;
    log('info', 'Loaded sample CIM text.');
  });

  $('btn-extract').addEventListener('click', runExtraction);

  function runExtraction(){
    const raw = $('rawText').value || '';
    if (isBlank(raw)){
      alert('Paste CIM text into Raw text (or load the sample) before extracting.');
      return;
    }

    $('btn-extract').disabled = true;
    $('btn-extract').textContent = 'Extracting…';

    try{
      log('info', 'Running offline mock extraction…');
      const title = appState.file ? appState.file.name : 'PastedText';
      const result = mockExtractFromText(title, raw);
      appState.extractionResult = result;

      const json = JSON.stringify(result, null, 2);
      $('aiJson').value = json;

      mapExtractionToForm(result);

      log('info', `Extraction complete. Missing critical: ${result.quality.missing_critical_fields.length}`);
      focusNextMissingRequired();
    } catch (err){
      const msg = (err && err.message) ? err.message : String(err);
      log('error', `Extraction failed: ${msg}`);
      alert(`Extraction failed\n\n${msg}`);
    } finally {
      $('btn-extract').disabled = false;
      $('btn-extract').textContent = 'Extract (Mock) Ctrl+Enter';
    }
  }

  // ------------------------------
  // Export
  // ------------------------------
  function collectFormData(){
    const get = (id) => $(id)?.value ?? '';
    return {
      metadata: {
        deal_name: get('dealname') || null,
        borrower_name: get('borrowername') || null,
        industry: get('industry') || null,
        reporting_currency: get('currency') || null,
        units_scale: Number(get('unitsscale') || 1),
      },
      financials: {
        income_statement: {
          period_end_date: get('periodend') || null,
          period_months: Number(get('periodmonths') || 12),
          revenue: parseNumberLoose(get('revenue')),
          ebitda: parseNumberLoose(get('ebitda')),
          ebitda_margin_percent: parseNumberLoose(get('ebitdamargin')),
          net_income: parseNumberLoose(get('netincome')),
        }
      },
      debt_structure: {
        tranche: {
          tranche_name: get('tranchename') || null,
          tranche_type: get('tranchetype') || null,
          commitment: parseNumberLoose(get('commitment')),
          drawn_amount: parseNumberLoose(get('drawn')),
          maturity_date: get('maturity') || null,
          margin_bps: parseNumberLoose(get('marginbps')),
        }
      },
      credit_metrics: {
        ltm: {
          as_of_date: get('asofdate') || null,
          total_debt: parseNumberLoose(get('totaldebt')),
          net_debt: parseNumberLoose(get('netdebt')),
          total_leverage_x: parseNumberLoose(get('totallev')),
          net_leverage_x: parseNumberLoose(get('netlev')),
          interest_coverage_x: parseNumberLoose(get('interestcov')),
          fixed_charge_coverage_x: parseNumberLoose(get('fccr')),
        }
      },
      extraction: {
        mode: appState.mode,
        confidence_threshold: appState.confidenceThreshold,
        result: appState.extractionResult,
        raw_text_chars: (appState.rawText || '').length,
      },
      validation: { issues: appState.issues },
      export_meta: { exported_at: new Date().toISOString() }
    };
  }

  function buildExportJsonString(){
    return JSON.stringify(collectFormData(), null, 2);
  }

  function buildExportCsv(){
    const data = collectFormData();
    const rows = [];
    const add = (k, v) => rows.push([k, (v === null || v === undefined) ? '' : String(v)]);

    add('deal_name', data.metadata.deal_name);
    add('borrower_name', data.metadata.borrower_name);
    add('industry', data.metadata.industry);
    add('reporting_currency', data.metadata.reporting_currency);
    add('units_scale', data.metadata.units_scale);

    add('period_end_date', data.financials.income_statement.period_end_date);
    add('period_months', data.financials.income_statement.period_months);
    add('revenue', data.financials.income_statement.revenue);
    add('ebitda', data.financials.income_statement.ebitda);
    add('ebitda_margin_percent', data.financials.income_statement.ebitda_margin_percent);
    add('net_income', data.financials.income_statement.net_income);

    add('tranche_name', data.debt_structure.tranche.tranche_name);
    add('tranche_type', data.debt_structure.tranche.tranche_type);
    add('commitment', data.debt_structure.tranche.commitment);
    add('drawn_amount', data.debt_structure.tranche.drawn_amount);
    add('maturity_date', data.debt_structure.tranche.maturity_date);
    add('margin_bps', data.debt_structure.tranche.margin_bps);

    add('ltm_as_of_date', data.credit_metrics.ltm.as_of_date);
    add('ltm_total_debt', data.credit_metrics.ltm.total_debt);
    add('ltm_net_debt', data.credit_metrics.ltm.net_debt);
    add('ltm_total_leverage_x', data.credit_metrics.ltm.total_leverage_x);
    add('ltm_net_leverage_x', data.credit_metrics.ltm.net_leverage_x);
    add('ltm_interest_coverage_x', data.credit_metrics.ltm.interest_coverage_x);
    add('ltm_fixed_charge_coverage_x', data.credit_metrics.ltm.fixed_charge_coverage_x);

    const escape = (s) => {
      const str = String(s ?? '');
      if (/[\",\n]/.test(str)) return '"' + str.replace(/\"/g,'""') + '"';
      return str;
    };
    return rows.map(r => r.map(escape).join(',')).join('\n');
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
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  async function copyToClipboard(text){
    try{
      await navigator.clipboard.writeText(text);
      log('info', 'Copied to clipboard.');
    } catch (_){
      log('warn', 'Clipboard blocked (often on file://). Use exported downloads.');
    }
  }

  function exportData(){
    validateAll();
    setActiveTab('tab-review');

    const jsonStr = buildExportJsonString();
    const csvStr = buildExportCsv();

    const name = ($('dealname').value || 'deal').replace(/[^a-z0-9-_]+/gi,'_').slice(0, 60);
    downloadText(`${name}.json`, jsonStr, 'application/json');
    downloadText(`${name}.csv`, csvStr, 'text/csv');

    log('info', 'Exported JSON + CSV (downloaded).');
  }

  // ------------------------------
  // Buttons
  // ------------------------------
  $('btn-validate').addEventListener('click', () => { validateAll(); focusNextIssue(1); });
  $('btn-export').addEventListener('click', exportData);

  $('btn-copy-json').addEventListener('click', () => copyToClipboard(buildExportJsonString()));
  $('btn-copy-csv').addEventListener('click', () => copyToClipboard(buildExportCsv()));

  $('btn-hide-issues').addEventListener('click', () => { $('issueRail').style.display = 'none'; });

  $('btn-clear').addEventListener('click', clearAll);

  function clearAll(){
    if (!confirm('Clear all in-memory data?')) return;

    for (const id of fieldOrder){
      const el = $(id);
      if (!el) continue;
      if (el.tagName === 'SELECT'){
        if (id === 'unitsscale') el.value = '1';
        else if (id === 'periodmonths') el.value = '12';
        else el.value = '';
      } else {
        el.value = '';
      }

      el.classList.remove('valid','invalid');
      const vEl = $(id + '-validation');
      if (vEl) vEl.textContent = '';
      const cEl = $(id + '-confidence');
      if (cEl) cEl.innerHTML = '';
    }

    $('rawText').value = '';
    $('aiJson').value = '';
    $('log').value = '';

    appState.file = null;
    appState.rawText = '';
    appState.extractionResult = null;
    appState.suggestions.clear();
    appState.issues = [];
    issueCursor = -1;

    $('fileInput').value = '';

    renderIssues();
    renderCompletion();
    renderReviewSummary();

    log('info', 'Cleared.');
    focusField('dealname');
  }

  // ------------------------------
  // Initialize
  // ------------------------------
  $('fileInput-validation').textContent = 'Upload a PDF, Excel, CSV, or TXT.';
  renderCompletion();
  validateAll();
  focusField('dealname');

})();
