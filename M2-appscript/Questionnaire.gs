/**
 * Questionnaire.gs — Populate questionnaire slides with client answers.
 *
 * Matches questions on slides to questionnaire columns and fills answers.
 * Also removes slides for goals the client didn't select.
 */

// ── Answer subcaptions (sourced from questionnaire form) ────
const ANSWER_SUBCAPTIONS = {
  'actively working':                    'Engaged in a full-time or part-time job, business, or self-employed with regular active income.',
  'soon to be retiring (within 5 yrs)':  'Planning to retire within the next few years',
  'soon to be retiring':                 'Planning to retire within the next few years',
  'retired early':                       'Not currently working by choice, but financially independent.',
  'retired':                             'No longer in active employment; dependent on pension, savings, or investments.',
  'active income only':                  'Regular earnings from salary, freelancing, or business.',
  'active + passive income':             'A mix of regular job/business income and recurring passive streams.',
  'active + passive':                    'A mix of regular job/business income and recurring passive streams.',
  'passive income only':                 'Recurring income from house rental, dividends, interest etc.',
  'pension income only':                 'Monthly pension received after retirement.',
  'passive + pension':                   'Combination of investment-based income and pension inflows.',
  'no regular source':                   'No income inflow.',
  'none':                                'No loans or dependents',
  'financial liabilities only':          'Loans/EMIs but no dependents.',
  'dependent liabilities only':          'People depend on your income, no loans.',
  'both financial & dependent':          'Loans/EMIs and dependents rely on your income.',
  'both financial and dependent':        'Loans/EMIs and dependents rely on your income.',
  'yes - comfortably':                   'I have enough surplus, no stress.',
  'just about':                          "I manage, but it's tight some months.",
  'no - struggling':                     'I often find it difficult to meet liabilities.',
  'short-term goals':                    'Less than 3 years',
  'medium-term goals':                   '3-5 years',
  'medium to long-term goals':           '5-8 years',
  'long-term wealth creation':           'More than 8 years',
  'exit all investments immediately':    'To prevent further loss',
  'exit partially':                      'Shift to safer options',
  'stay invested':                       "I'm comfortable with market fluctuations",
  'invest more':                         'I will average my cost',
};


/**
 * Get the answer for a questionnaire question from the client's data.
 */
function _getAnswer(questionText, qRow, context) {
  const q = questionText.toLowerCase().trim();

  if (q.includes('your age'))           return _safeStr(qRow['Age']);
  if (q.includes('employment status'))  return _safeStr(qRow['Employment Status']);
  if (q.includes('source of income'))   return _safeStr(qRow['Income Source']);

  if (q.includes('reason for investing') || q.includes('investing in mutual'))
    return _safeStr(qRow['Goals']);
  if (q.includes('types of liabilities'))
    return _safeStr(qRow['Liability Type']);
  if (q.includes('comfortably meet'))
    return _safeStr(qRow['Liability Followup Answer']);

  if (q.includes('emergency fund'))     return _safeStr(qRow['Emergency Fund']);
  if (q.includes('portfolio to grow') || (q.includes('prefer') && q.includes('portfolio')))
    return _safeStr(qRow['Portfolio Preference']);
  if (q.includes('investment horizon')) return _safeStr(qRow['Investment Horizon']);

  if (q.includes('investments fall') || q.includes('fall by 20'))
    return _safeStr(qRow['Fall Reaction']);
  if (q.includes('lumpsum'))
    return _safeInr(qRow['Lumpsum Amount (with Infinite)'] || 0);
  if (q.includes('monthly sip') && q.includes('amount'))
    return _safeInr(qRow['Monthly SIP Amount (with Infinite)'] || 0);

  // Retirement
  if (q.includes('monthly income') && q.includes('expense')) {
    const inc = qRow['Ret: Monthly Income'] || 0;
    const exp = qRow['Ret: Monthly Expenses'] || 0;
    return 'Income: ' + _safeInr(inc) + ' ; Expenses: ' + _safeInr(exp);
  }
  if (q.includes('change in expenses'))
    return _safePct(qRow['Ret: Expense Change %']);
  if (q.includes('current monthly investment'))
    return _safeInr(qRow['Ret: Monthly Investment'] || 0);
  if (q.includes('year-on-year') || q.includes('yoy'))
    return _safePct(qRow['Ret: YoY Investment Increase %']);
  if (q.includes('financial investments apart'))
    return _safeStr(qRow['Other Investments Value']);

  // Home purchase
  if (context !== 'vehicle') {
    if (q.includes('when do you want to purchase'))
      return _safeStr(qRow['Home: Purchase Year']);
    if (q.includes('flexibility to shift')) {
      const v = qRow['Home: Flexibility Yrs'];
      return v ? v + ' years' : '-';
    }
    if (q.includes('down payment'))
      return _safePct(qRow['Home: Down Payment %']);
    if (q.includes('debt financing'))
      return _safeStr(qRow['Home: Loan Y/N']);
  }
  if (q.includes('value of home'))
    return _safeInr(qRow['Home: Value'] || 0);

  // Children's Education
  if (q.includes('number of children') && q.includes('education')) {
    let count = 0;
    for (let i = 1; i <= 4; i++) {
      const v = qRow['Edu: Child ' + i + ' UG Year'];
      if (v && String(v).trim() !== '' && String(v) !== 'NaN') count++;
    }
    return String(Math.max(count, 1));
  }

  // Children's Marriage
  if (q.includes('number of children') && q.includes('marriage')) {
    const names = [];
    for (let i = 1; i <= 4; i++) {
      const n = qRow['Marriage: Child ' + i + ' Name'];
      if (n && String(n).trim()) names.push(String(n).trim());
    }
    const count = Math.max(names.length, 1);
    return names.length ? count + ' - ' + names.join(', ') : String(count);
  }

  // Vehicle
  if (q.includes('vehicle') || context === 'vehicle') {
    if (q.includes('flexibility') || q.includes('shift')) {
      const v = qRow['Vehicle: Flexibility Yrs'];
      return v ? v + ' years' : '-';
    }
    if (q.includes('down payment'))
      return _safePct(qRow['Vehicle: Down Payment %']);
    if (q.includes('when do you want'))
      return _safeStr(qRow['Vehicle: Purchase Year']);
    if (q.includes('value of vehicle') || (q.includes('value') && context === 'vehicle'))
      return _safeInr(qRow['Vehicle: Value'] || 0);
    if (q.includes('debt financing') || q.includes('financing'))
      return _safeStr(qRow['Vehicle: Loan Y/N']);
  }

  return null;
}


// ── Safe formatting helpers ─────────────────────────────────

function _safeStr(val) {
  if (val === null || val === undefined) return '-';
  if (typeof val === 'number' && isNaN(val)) return '-';
  const s = String(val).trim();
  return (s && s !== 'NaN') ? s : '-';
}

function _safeInr(val) {
  if (val === null || val === undefined) return '-';
  if (typeof val === 'number' && isNaN(val)) return '-';
  try {
    const fv = parseFloat(val);
    const d = fmtInrDisplay(fv);
    return d || '-';
  } catch (e) {
    return String(val).trim() || '-';
  }
}

function _safePct(val) {
  if (val === null || val === undefined) return '-';
  if (typeof val === 'number' && isNaN(val)) return '-';
  const s = String(val).trim();
  if (s.includes('%')) return s;
  try {
    let fv = parseFloat(s);
    if (Math.abs(fv) > 0 && Math.abs(fv) < 1) fv *= 100;
    return fv.toFixed(0) + '%';
  } catch (e) {
    return s || '-';
  }
}


/**
 * Populate questionnaire slides with the client's answers.
 * Also remove slides for goals the client didn't select,
 * and renumber remaining slides (X/total).
 */
function doQuestionnaire(presentation, goals, qRow) {
  if (!qRow) {
    Logger.log('Questionnaire: no qRow — skipping');
    return;
  }

  const norm = new Set();
  for (const g of goals) {
    const gl = g.toLowerCase();
    if (gl.includes('retirement'))  norm.add('Retirement Planning');
    if (gl.includes('home'))        norm.add('Home Purchase');
    if (gl.includes('education'))   norm.add("Children's Education");
    if (gl.includes('marriage'))    norm.add("Children's Marriage");
    if (gl.includes('vehicle'))     norm.add('Vehicle Purchase');
  }

  const GOAL_KW_SLIDES = {
    'Home Purchase':        'Home Purchase',
    "Children's Education": "Children's Education",
    "Children's Marriage":  "Children's Marriage",
    'Vehicle Purchase':     'Vehicle Purchase',
    'Vehicle':              'Vehicle Purchase',
  };

  // Find questionnaire slides (those containing "Infinite Questionnaire")
  const allSlides = presentation.getSlides();
  const qIndices = [];
  for (let i = 0; i < allSlides.length; i++) {
    for (const shape of allSlides[i].getShapes()) {
      if (shape.getText().asString().includes('Infinite Questionnaire')) {
        qIndices.push(i);
        break;
      }
    }
  }

  Logger.log('Questionnaire: found ' + qIndices.length + ' slides, goals=' + Array.from(norm));

  // Step 1: Populate answers
  for (const idx of qIndices) {
    _populateQuestionnaireSlide(allSlides[idx], qRow);
  }

  // Step 2: Remove slides for unselected goals
  const toDelete = [];
  for (const idx of qIndices) {
    const slide = allSlides[idx];
    let title = '';
    for (const shape of slide.getShapes()) {
      if (shape.getText().asString().includes('Infinite Questionnaire')) {
        title = shape.getText().asString().trim();
        break;
      }
    }

    for (const [kw, goalName] of Object.entries(GOAL_KW_SLIDES)) {
      if (title.includes(kw) && !norm.has(goalName)) {
        toDelete.push(idx);
        Logger.log('  Remove: "' + title + '"');
        break;
      }
    }
  }

  // Remove slides (highest index first)
  const slidesToRemove = [...new Set(toDelete)].sort((a, b) => b - a);
  const currentSlides = presentation.getSlides();
  for (const idx of slidesToRemove) {
    if (idx < currentSlides.length) {
      currentSlides[idx].remove();
    }
  }

  // Step 3: Renumber (X/total)
  const finalSlides = presentation.getSlides();
  const qSlidesFinal = [];
  for (let i = 0; i < finalSlides.length; i++) {
    for (const shape of finalSlides[i].getShapes()) {
      if (shape.getText().asString().includes('Infinite Questionnaire')) {
        qSlidesFinal.push({ idx: i, shape: shape });
        break;
      }
    }
  }

  const total = qSlidesFinal.length;
  for (let seq = 0; seq < total; seq++) {
    const shape = qSlidesFinal[seq].shape;
    const oldText = shape.getText().asString();
    const newText = oldText.replace(/\(\d+\/\d+\)/, '(' + (seq + 1) + '/' + total + ')');
    if (newText !== oldText) {
      setShapeText(shape, newText);
    }
  }

  Logger.log('Questionnaire: ' + toDelete.length + ' removed, ' + total + ' remaining');
}


/**
 * Fill answers on a single questionnaire slide.
 * Looks at group shapes: child 1 = question (larger font), child 2 = answer (smaller font).
 */
function _populateQuestionnaireSlide(slide, qRow) {
  // Detect slide context
  let context = '';
  for (const shape of slide.getShapes()) {
    const t = shape.getText().asString();
    if (t.includes('Vehicle Purchase') || (t.includes('Vehicle') && t.includes('Purchase'))) {
      context = 'vehicle'; break;
    }
    if (t.includes('Home Purchase') || (t.includes('Home') && t.includes('Purchase'))) {
      context = 'home'; break;
    }
  }

  // Process groups
  for (const group of slide.getGroups()) {
    try {
      const children = group.getChildren();
      const textChildren = [];
      for (const child of children) {
        try {
          const shape = child.asShape();
          if (shape.getText().asString().trim()) {
            textChildren.push(shape);
          }
        } catch (e) {}
      }

      if (textChildren.length < 2) continue;

      // First text child = question, second = answer
      const qShape = textChildren[0];
      const aShape = textChildren[1];
      const qText = qShape.getText().asString().trim();

      const answer = _getAnswer(qText, qRow, context);
      if (answer !== null) {
        // Set answer text
        aShape.getText().setText(answer);

        // Set subcaption if available
        const subcap = ANSWER_SUBCAPTIONS[answer.toLowerCase().trim()] || '';
        // If the answer shape has multiple paragraphs, set subcap in the second
        const paras = aShape.getText().getParagraphs();
        if (paras.length >= 2 && subcap) {
          paras[1].getRange().setText(subcap);
        }

        Logger.log('  Q: "' + qText.substring(0, 55) + '" -> "' + answer.substring(0, 40) + '"');
      }
    } catch (e) {
      // Group processing failed — skip
    }
  }
}
