/**
 * Slide13Chart.gs — Portfolio vs Infinite comparison chart (Slide 13).
 *
 * Generates a line chart comparing:
 *   - Actual portfolio performance (blue dotted)
 *   - Infinite strategy performance (dark solid)
 *
 * Uses Google Sheets Charts to generate the image, then inserts into the slide.
 */

function doSlide13(presentation, pfId, riskProfile, data) {
  const slide = presentation.getSlides()[12];

  const prefix = RISK_TYPE_PREFIX[riskProfile] || 'B';
  const infType = bestInfiniteType(pfId, prefix, data.results);
  if (!infType) {
    Logger.log('Slide 13: no Infinite type found for prefix "' + prefix + '"');
    return;
  }

  // Get chart lines
  const custLines = data.lines.filter(r => String(r.PF_ID) === String(pfId));
  const pfLines = custLines.filter(r => r.TYPE === 'pf')
    .sort((a, b) => new Date(a.DATE) - new Date(b.DATE));
  const infLines = custLines.filter(r => r.TYPE === infType)
    .sort((a, b) => new Date(a.DATE) - new Date(b.DATE));

  if (pfLines.length === 0 || infLines.length === 0) {
    Logger.log('Slide 13: missing line data — skipping');
    return;
  }

  // Results for XIRR values
  const custResults = data.results.filter(r => String(r.PF_ID) === String(pfId));
  const pfRes  = custResults.find(r => r.TYPE === 'pf');
  const infRes = custResults.find(r => r.TYPE === infType);
  const pfXirr  = pfRes  ? (pfRes.XIRR  || 0) : 0;
  const infXirr = infRes ? (infRes.XIRR || 0) : 0;
  const pfFinal  = pfRes  ? (pfRes.CURRENT_VALUE  || 0) : 0;
  const infFinal = infRes ? (infRes.CURRENT_VALUE || 0) : 0;

  // Create temporary spreadsheet for the chart
  const tempSS = SpreadsheetApp.create('_M2_Slide13Chart_' + new Date().getTime());
  const sheet = tempSS.getActiveSheet();

  // Write headers
  sheet.getRange(1, 1).setValue('Date');
  sheet.getRange(1, 2).setValue('Portfolio (₹L)');
  sheet.getRange(1, 3).setValue('Infinite ' + riskProfile + ' (₹L)');

  // Build a merged date series
  const allDates = new Set();
  pfLines.forEach(r => allDates.add(String(r.DATE)));
  infLines.forEach(r => allDates.add(String(r.DATE)));
  const sortedDates = Array.from(allDates).sort();

  // Build value maps
  const pfMap = {};
  pfLines.forEach(r => pfMap[String(r.DATE)] = r.CURRENT_VALUE);
  const infMap = {};
  infLines.forEach(r => infMap[String(r.DATE)] = r.CURRENT_VALUE);

  // Write data rows in a single batch (MUCH faster than cell-by-cell)
  const rows = [];
  for (const d of sortedDates) {
    const pfVal  = pfMap[d];
    const infVal = infMap[d];
    if (pfVal !== undefined || infVal !== undefined) {
      rows.push([
        d,
        pfVal  !== undefined ? pfVal  / 1e5 : '',
        infVal !== undefined ? infVal / 1e5 : '',
      ]);
    }
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  const rowNum = rows.length + 2;
  SpreadsheetApp.flush();

  // Build line chart
  const dataRange = sheet.getRange(1, 1, rowNum - 1, 3);
  const chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dataRange)
    .setOption('title', '')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', {
      format: 'dd-MM-yyyy',
      textStyle: { fontSize: 8, color: '#555555' },
      gridlines: { count: 0 },
    })
    .setOption('vAxis', {
      textPosition: 'none',
      gridlines: { count: 0 },
    })
    .setOption('backgroundColor', 'white')
    .setOption('chartArea', { left: '3%', right: '3%', top: '3%', height: '85%' })
    .setOption('series', {
      0: { color: '#4E9EED', lineDashStyle: [3, 2], lineWidth: 1 },
      1: { color: '#1A1A2E', lineWidth: 1.5 },
    })
    .setPosition(1, 1, 0, 0);

  const chart = chartBuilder.build();
  sheet.insertChart(chart);

  // Get chart as image and insert into slide
  const charts = sheet.getCharts();
  if (charts.length > 0) {
    const chartBlob = charts[0].getBlob();

    // Remove old chart image (shape ;338)
    for (const img of slide.getImages()) {
      if (img.getWidth() > 300 && img.getTop() > 100) {
        img.remove();
        break;
      }
    }

    const newImg = slide.insertImage(chartBlob);
    newImg.setLeft(-12);   // -152400 EMU ÷ 12700
    newImg.setTop(131);    // 1666799 EMU ÷ 12700
    newImg.setWidth(402);  // 5110622 EMU ÷ 12700
    newImg.setHeight(249); // 3160301 EMU ÷ 12700
    newImg.sendToBack();
  }

  // Clean up temp spreadsheet
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);

  // Update text shapes
  for (const shape of slide.getShapes()) {
    const txt = shape.getText().asString().trim();
    if (txt.includes('Infinite') && txt.length < 30 && !txt.includes('XIRR')) {
      setShapeText(shape, 'Infinite ' + riskProfile);
    }
  }

  // Update XIRR table
  for (const table of slide.getTables()) {
    try {
      // Row 1, Col 1 = Infinite XIRR, Col 2 = Actual XIRR
      table.getCell(1, 1).getText().setText((infXirr * 100).toFixed(2) + '%');
      table.getCell(1, 2).getText().setText((pfXirr * 100).toFixed(2) + '%');
    } catch (e) {
      // Table structure may vary
      Logger.log('Slide 13: table update failed: ' + e);
    }
  }

  // Update final value labels
  for (const shape of slide.getShapes()) {
    const txt = shape.getText().asString().trim();
    // These are the end-of-line value labels
    if (txt.match(/^\d/) && txt.length < 10) {
      // Positional: higher shape = Infinite, lower = Actual
      // This is a rough heuristic; may need refinement
    }
  }

  Logger.log('Slide 13: chart done (' + riskProfile + ' -> ' + infType + ')');
  Logger.log('Slide 13: Actual XIRR=' + (pfXirr * 100).toFixed(2) + '%, Infinite XIRR=' + (infXirr * 100).toFixed(2) + '%');
}
