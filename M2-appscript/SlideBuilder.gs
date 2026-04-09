/**
 * SlideBuilder.gs — Google Slides manipulation for M2 deck.
 *
 * Works directly with the Slides API (SlidesApp + Advanced Slides Service).
 * No PPTX conversion needed — operates on native Google Slides.
 *
 * Shape matching: The Python version uses shape names like "Google Shape;165;p19".
 * In Slides API, we match shapes by their objectId or by text content patterns.
 */


// ── Text replacement helpers ────────────────────────────────

/**
 * Find a shape on a slide whose text contains the given substring.
 * Returns the first matching shape, or null.
 */
function findShapeByText(slide, substring) {
  const shapes = slide.getShapes();
  for (const shape of shapes) {
    if (shape.getText().asString().includes(substring)) {
      return shape;
    }
  }
  return null;
}


/**
 * Find all shapes on a slide whose text contains the given substring.
 */
function findShapesByText(slide, substring) {
  return slide.getShapes().filter(s => s.getText().asString().includes(substring));
}


/**
 * Replace the entire text content of a shape, preserving the first text style.
 */
function setShapeText(shape, newText) {
  const tf = shape.getText();
  // Get style of the first character (if any) to preserve formatting
  let style = null;
  try {
    if (tf.asString().length > 1) {
      style = tf.getRange(0, 1).getTextStyle();
    }
  } catch (e) { style = null; }
  tf.setText(String(newText));
  // Only apply styling if there's actual text content (length > 1 because of trailing \n)
  if (style && tf.asString().length > 1) {
    try {
      const newRange = tf.getRange(0, tf.asString().length - 1);
      if (style.getFontFamily()) newRange.getTextStyle().setFontFamily(style.getFontFamily());
      if (style.getFontSize())   newRange.getTextStyle().setFontSize(style.getFontSize());
      if (style.getForegroundColor() && style.getForegroundColor().asRgbColor()) {
        newRange.getTextStyle().setForegroundColor(style.getForegroundColor());
      }
      newRange.getTextStyle().setBold(style.isBold());
    } catch (e) {
      // Style application failed — text is still set, just unstyled
    }
  }
}


/**
 * Replace text matching a pattern within a shape's text.
 */
function replaceInShape(shape, oldText, newText) {
  shape.getText().replaceAllText(oldText, newText);
}


// ── Slide 1: Title ──────────────────────────────────────────

function doSlide1(presentation, fullName) {
  const slide = presentation.getSlides()[0];
  const shapes = slide.getShapes();
  for (const shape of shapes) {
    const text = shape.getText().asString().toLowerCase();
    if (text.includes('with') && text.length < 80) {
      setShapeText(shape, 'with ' + fullName);
      // Set white text
      const range = shape.getText().getRange(0, shape.getText().asString().length - 1);
      range.getTextStyle().setForegroundColor('#FFFFFF');
      Logger.log('Slide 1: title -> "with ' + fullName + '"');
      return;
    }
  }
  Logger.log('Slide 1: WARNING - name placeholder not found');
}


// ── Slide 2: Welcome ────────────────────────────────────────

function doSlide2(presentation, firstName) {
  const slide = presentation.getSlides()[1];
  for (const shape of slide.getShapes()) {
    const text = shape.getText().asString();
    if (text.includes('Welcome')) {
      // The name is in the second paragraph (or second line)
      const paras = shape.getText().getParagraphs();
      if (paras.length >= 2) {
        // Clear paragraph 1 and set firstName
        const p1 = paras[1];
        p1.getRange().setText(firstName);
      }
      Logger.log('Slide 2: welcome -> "' + firstName + '"');
      return;
    }
  }
  Logger.log('Slide 2: WARNING - name placeholder not found');
}


// ── Slide 3: You at a Glance ────────────────────────────────

function doSlide3(presentation, qRow, riskProfile) {
  const slide = presentation.getSlides()[2];
  const goals    = parseGoals(qRow['Goals'] || '');
  const horizon  = getHorizon(qRow['Investment Horizon'] || '');
  const age      = qRow['Age'] || '';
  const lumpVal  = qRow['Lumpsum Amount (with Infinite)'] || 0;
  const sipVal   = qRow['Monthly SIP Amount (with Infinite)'] || 0;
  const stepUpRaw = qRow['Ret: YoY Investment Increase %'] || 0;

  const lumpStr = fmtInrDisplay(lumpVal) || 'INR 0';
  const sipStr  = fmtInrDisplay(sipVal)  || 'INR 0';

  let stepUp = 0;
  try {
    stepUp = parseFloat(String(stepUpRaw).replace('%', '')) || 0;
  } catch (e) { stepUp = 0; }
  const hasStepup = stepUp > 0;
  const sipLabel = hasStepup ? 'monthly SIP*' : 'monthly SIP';

  // Strategy: find shapes by their current text content pattern and replace
  for (const shape of slide.getShapes()) {
    const txt = shape.getText().asString().trim();

    // Goals shape (contains goal-like text)
    if (txt.includes('Wealth') || txt.includes('Financial') || txt.includes('Retirement')) {
      if (shape.getLeft() > 100 && shape.getTop() < 3000000) {
        const primary = goals[0] || 'Wealth Creation';
        const secondary = goals.slice(1).join(', ');
        setShapeText(shape, secondary ? primary + '\n' + secondary : primary);
        Logger.log('Slide 3: goals -> "' + primary + '"');
      }
    }

    // Horizon (short text containing "Years")
    if ((txt.includes('Years') || txt.includes('years')) && txt.length < 30) {
      setShapeText(shape, horizon);
      Logger.log('Slide 3: horizon -> "' + horizon + '"');
    }

    // Risk profile text
    if (txt === 'Balanced' || txt === 'Aggressive' || txt === 'Conservative' ||
        txt === 'Very Aggressive' || txt === 'Very Conservative') {
      setShapeText(shape, riskProfile);
      Logger.log('Slide 3: risk -> "' + riskProfile + '"');
    }

    // Risk profile + "Investor"
    if (txt.includes('Investor') && txt.length < 40) {
      setShapeText(shape, riskProfile + ' Investor');
    }

    // Investment text (contains "with" and INR amounts)
    if (txt.includes('with') && (txt.includes('INR') || txt.includes('₹'))) {
      setShapeText(shape, lumpStr + ' with ' + sipStr + ' ' + sipLabel);
    }

    // Age
    if (txt.includes('Current Age') && txt.length < 30) {
      if (age) setShapeText(shape, 'Current Age: ' + age);
    }

    // SIP step-up
    if (txt.includes('SIP Step') || txt.includes('step-Up') || txt.includes('Step-Up')) {
      if (hasStepup) {
        setShapeText(shape, 'SIP Step-Up every year: ' + stepUp.toFixed(0) + '%');
      } else {
        // Remove the shape text (effectively hiding it)
        shape.getText().setText('');
      }
    }
  }
}


// ── Slide 4: Portfolio Snapshot ──────────────────────────────

function doSlide4(presentation, pfRow, rgAgg, riskProfile) {
  const slide = presentation.getSlides()[3];
  const cv   = pfRow.PF_CURRENT_VALUE || 0;
  const iv   = pfRow.INVESTED_VALUE || 0;
  const xirr = pfRow.PF_XIRR || 0;
  const bxir = pfRow.BM_XIRR || 0;
  const pg   = pfRow.PF_GAINS || 0;
  const bg   = (pfRow.BM_CURRENT_VALUE || iv) - iv;
  const sm   = ((pfRow.SMALL || 0) + (pfRow.MID || 0)) * (pfRow.EQUITY || 0) * 100;
  const pfRisk   = portfolioRisk(sm);
  const matches  = pfRisk === riskProfile;

  // Replace metric values by matching text patterns
  for (const shape of slide.getShapes()) {
    const txt = shape.getText().asString().trim();

    // Current value (large number, typically ₹XX.XCr)
    // We identify shapes by their relative position and text content
    if (txt.includes('Portfolio gains')) {
      setShapeText(shape, 'Portfolio gains: ' + fmtInrRupee(pg));
      Logger.log('Slide 4: PF gains -> ' + fmtInrRupee(pg));
    }
    if (txt.includes('Benchmark gains')) {
      setShapeText(shape, 'Benchmark gains: ' + fmtInrRupee(bg));
      Logger.log('Slide 4: BM gains -> ' + fmtInrRupee(bg));
    }
    if (txt.includes('Small') && txt.includes('Mid')) {
      setShapeText(shape, 'Small + Mid Allocation: ' + sm.toFixed(0) + '%');
    }

    // Risk match/mismatch text
    if (txt.includes('risk profile') && (txt.includes('Matches') || txt.includes("Doesn't"))) {
      const newText = matches ? 'Matches your risk profile' : "Doesn't match your risk profile";
      const color   = matches ? '#2A9C4A' : '#CC0000';
      setShapeText(shape, newText);
      const range = shape.getText().getRange(0, shape.getText().asString().length - 1);
      range.getTextStyle().setForegroundColor(color);
    }
  }

  // Find shapes that show the 4 key metric values by object ID pattern
  // Since we can't rely on shape names from Python, we use a positional approach:
  // Replace placeholder values that look like numbers/percentages
  _replaceMetricValues(slide, cv, iv, xirr, bxir, pfRisk);

  // Generate pie chart
  _updatePieChart(slide, rgAgg, presentation);
}


/**
 * Replace the 4 key metric values on slide 4.
 * Searches for shapes with short numeric-looking text.
 */
function _replaceMetricValues(slide, cv, iv, xirr, bxir, pfRisk) {
  // Collect all shapes with short text that could be metric placeholders
  const candidates = [];
  for (const shape of slide.getShapes()) {
    const txt = shape.getText().asString().trim();
    // Metric value shapes typically show ₹XX or XX%
    if (txt.length < 15 && (txt.includes('₹') || txt.includes('%') || txt.includes('Rs'))) {
      candidates.push({
        shape: shape,
        text: txt,
        left: shape.getLeft(),
        top: shape.getTop(),
      });
    }
  }

  // Sort by position (top-to-bottom, left-to-right) to identify which is which
  candidates.sort((a, b) => a.top - b.top || a.left - b.left);

  // Expected order based on base deck layout:
  // Shape 165 (CV), 163 (IV), 161 (XIRR), 159 (BM XIRR)
  // These are typically in the top portion of the slide, left to right
  const values = [
    fmtInrRupee(cv),
    fmtInrRupee(iv),
    (xirr * 100).toFixed(1) + '%',
    (bxir * 100).toFixed(1) + '%',
  ];

  for (let i = 0; i < Math.min(candidates.length, values.length); i++) {
    setShapeText(candidates[i].shape, values[i]);
    Logger.log('Slide 4: metric ' + i + ' -> ' + values[i]);
  }

  // Portfolio risk text
  for (const shape of slide.getShapes()) {
    const txt = shape.getText().asString().trim();
    if (RISK_SCALE.includes(txt)) {
      setShapeText(shape, pfRisk);
      Logger.log('Slide 4: portfolio risk -> ' + pfRisk);
      break;
    }
  }
}


// ── Slide 6: What's working well ────────────────────────────

function doSlide6(presentation, pfRow, riskProfile) {
  const slide = presentation.getSlides()[5];

  const yearsRaw = pfRow.YEARS_SINCE_FIRST_TRANSACTION || 0;
  let yearsInt;
  try { yearsInt = parseInt(yearsRaw) || 0; } catch (e) { yearsInt = 0; }

  const cv = pfRow.PF_CURRENT_VALUE || 0;
  const cvStr = fmtInrRupee(cv, '');

  const pfXirr = pfRow.PF_XIRR || 0;
  const bmXirr = pfRow.BM_XIRR || 0;
  const useCompetitive = pfXirr > bmXirr;
  const diff = ((pfXirr - bmXirr) * 100).toFixed(1);

  for (const shape of slide.getShapes()) {
    const txt = shape.getText().asString();

    // Box 01 description — mentions "SIPs" or "lump sum" and "corpus"
    if (txt.includes('SIPs') || txt.includes('lump sum') || txt.includes('corpus')) {
      setShapeText(shape,
        'SIPs & lump sum over ' + yearsInt + ' years building a corpus of ' + cvStr);
    }

    // Box 02 title — "Delivering" or "Aligned"
    if (txt.includes('Delivering competitive') || txt.includes('Aligned to your risk')) {
      if (useCompetitive) {
        setShapeText(shape, 'Delivering competitive returns');
      } else {
        setShapeText(shape, 'Aligned to your risk profile');
      }
    }

    // Box 02 description
    if (txt.includes('Portfolio performance') || txt.includes('Your portfolio reflects')) {
      if (useCompetitive) {
        setShapeText(shape,
          "Portfolio performance has edged past benchmark by " + diff + "%, you're on the right track");
      } else {
        setShapeText(shape,
          'Your portfolio reflects your preferred risk level: ' + riskProfile + '*');
      }
    }
  }

  Logger.log('Slide 6: ' + yearsInt + 'y, ' + cvStr +
    ', variant=' + (useCompetitive ? 'competitive' : 'aligned'));
}


// ── Pie chart (Slide 4) ─────────────────────────────────────

/**
 * Generate a donut chart using Google Sheets Charts API and insert into the slide.
 *
 * Strategy: Create a temporary spreadsheet with the chart data,
 * build a chart there, then link/embed it into the slide.
 *
 * Alternative: Use the Slides API to insert a Sheets-linked chart.
 */
function _updatePieChart(slide, rgAgg, presentation) {
  if (!rgAgg || rgAgg.length === 0) {
    Logger.log('Slide 4: no riskgroup data - skipping chart');
    return;
  }

  // Build chart data
  const parts = [];
  for (const row of rgAgg) {
    const g = row.group;
    const p = row.pctOfPF;
    if (!p || p <= 0) continue;
    parts.push({
      label: CHART_LABELS[g] || g,
      pct:   p * 100,
      color: CHART_COLORS[g] || '#808080',
    });
  }
  if (parts.length === 0) return;

  // Create a temporary spreadsheet for the chart
  const tempSS = SpreadsheetApp.create('_M2_TempChart_' + new Date().getTime());
  const sheet = tempSS.getActiveSheet();

  // Write data
  sheet.getRange(1, 1).setValue('Category');
  sheet.getRange(1, 2).setValue('Percentage');
  for (let i = 0; i < parts.length; i++) {
    sheet.getRange(i + 2, 1).setValue(parts[i].label);
    sheet.getRange(i + 2, 2).setValue(parts[i].pct);
  }

  // Build donut chart
  const chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(1, 1, parts.length + 1, 2))
    .setOption('title', '')
    .setOption('legend', { position: 'none' })
    .setOption('pieHole', 0.65)
    .setOption('pieSliceBorderColor', 'white')
    .setOption('backgroundColor', 'transparent')
    .setOption('chartArea', { left: '5%', top: '5%', width: '90%', height: '90%' })
    .setPosition(1, 1, 0, 0);

  // Set colours
  const colors = parts.map(p => p.color);
  chartBuilder.setOption('colors', colors);

  const chart = chartBuilder.build();
  sheet.insertChart(chart);

  // Get chart as image blob
  const charts = sheet.getCharts();
  if (charts.length > 0) {
    const chartBlob = charts[0].getBlob();

    // Remove old pie chart image from the slide (large images in the middle area)
    const images = slide.getImages();
    for (const img of images) {
      if (img.getWidth() > 300 && img.getHeight() > 200 && img.getTop() < 200) {
        img.remove();
        Logger.log('Slide 4: removed old pie chart image');
        break;
      }
    }

    // Insert new chart image
    const newImg = slide.insertImage(chartBlob);
    // Position to match the original pie chart area
    newImg.setLeft(206);   // ~2667449 EMU ÷ 12700
    newImg.setTop(82);     // ~1043625 EMU ÷ 12700
    newImg.setWidth(406);  // ~5159751 EMU ÷ 12700
    newImg.setHeight(251); // ~3190450 EMU ÷ 12700

    // Send behind other shapes (z-order)
    newImg.sendToBack();

    Logger.log('Slide 4: donut chart inserted (' + parts.length + ' segments)');
  }

  // Clean up temp spreadsheet
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);

  // Update legend text shapes
  _updateLegendText(slide, parts);
}


/**
 * Update legend percentage text on slide 4.
 */
function _updateLegendText(slide, parts) {
  const pctMap = {};
  for (const p of parts) pctMap[p.label] = p.pct;

  const eqLabels = ['Aggressive', 'Balanced', 'Conservative'];
  const eqTotal = parts.filter(p => eqLabels.includes(p.label))
                       .reduce((s, p) => s + p.pct, 0);

  // Search for legend text shapes and update percentages
  for (const shape of slide.getShapes()) {
    const txt = shape.getText().asString().trim();

    // Category percentage shapes typically show "XX%" next to category labels
    for (const [label, pct] of Object.entries(pctMap)) {
      if (txt === label && shape.getLeft() < 200) {
        // This is a label shape — check adjacent shape for percentage
        continue;
      }
    }

    // Equity total percentage
    if (txt.includes('Equity') && txt.length < 20) {
      // Could be "Equity" label or percentage — leave label, update pct
    }
  }
  // Note: Fine-grained legend manipulation is complex in Slides API.
  // The chart image itself includes the data; the legend shapes in the
  // template may need manual adjustment depending on template structure.
}
