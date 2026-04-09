/**
 * Appendix.gs — Scheme appendix slides.
 *
 * Builds scheme-data slides showing fund details per subcategory.
 * Uses the template slides (indices 22-25) as blueprints, then clones
 * and fills with scheme data.
 *
 * Key difference from Python: In Apps Script, we duplicate slides natively
 * using SlidesApp, which is much simpler than python-pptx XML manipulation.
 */

/**
 * Build and fill appendix scheme slides.
 * Returns the number of slides created.
 */
function doAppendix(presentation, pfId, data) {
  const schemes = data.schemeLevel.filter(r => String(r.PF_ID) === String(pfId));
  if (schemes.length === 0) {
    Logger.log('Appendix: no schemes - removing template slides');
    _removeTemplateSlides(presentation);
    return 0;
  }

  // Load categorization (from the categorization data in schemeLevel itself)
  // In the Python version this comes from a separate Excel file, but the
  // UPDATED_SUBCATEGORY column is already on the scheme data.

  // Group schemes by UPDATED_SUBCATEGORY
  const groups = _groupSchemes(schemes);

  // Build slide specs (which template to use per group)
  const specs = _buildSlideSpecs(groups);
  Logger.log('Appendix: ' + schemes.length + ' schemes -> ' + specs.length + ' slides');

  if (specs.length === 0) {
    _removeTemplateSlides(presentation);
    return 0;
  }

  const slides = presentation.getSlides();

  // Template slide indices (0-based): 22=4row, 23=3row, 24=2row, 25=1row
  // Clone templates for each spec
  const tplMap = { 4: 22, 3: 23, 2: 24, 1: 25 };

  // Clone the needed slides (insert after templates)
  const newSlideIndices = [];
  for (const spec of specs) {
    const tplIdx = tplMap[Math.min(spec.rowCount, 4)] || tplMap[4];

    // Duplicate the template slide
    const tplSlide = slides[tplIdx];
    const newSlide = presentation.appendSlide(tplSlide);
    newSlideIndices.push(presentation.getSlides().length - 1);
  }

  // Remove the 4 original template slides (indices 22-25)
  // Must remove from highest index first to preserve indices
  const allSlides = presentation.getSlides();
  for (let i = 25; i >= 22; i--) {
    if (i < allSlides.length) {
      allSlides[i].remove();
    }
  }

  // Move new slides to position 22 onwards
  const currentSlides = presentation.getSlides();
  const totalSlides = currentSlides.length;
  const newStart = totalSlides - specs.length;
  for (let i = 0; i < specs.length; i++) {
    currentSlides[newStart + i].move(22 + i + 1); // 1-based position
  }

  // Fill each new slide with scheme data
  const filledSlides = presentation.getSlides();
  for (let i = 0; i < specs.length; i++) {
    _fillSchemeSlide(filledSlides[22 + i], specs[i]);
  }

  Logger.log('Appendix: ' + specs.length + ' slides created & filled');
  return specs.length;
}


function _removeTemplateSlides(presentation) {
  const slides = presentation.getSlides();
  for (let i = Math.min(25, slides.length - 1); i >= 22; i--) {
    slides[i].remove();
  }
}


/**
 * Group schemes by subcategory, sorted by categorization sort order.
 */
function _groupSchemes(schemes) {
  const grouped = {};
  for (const s of schemes) {
    const subcat = s.UPDATED_SUBCATEGORY;
    if (!subcat) continue;
    if (!grouped[subcat]) {
      grouped[subcat] = {
        cat:  s.UPDATED_BROAD_CATEGORY_GROUP || subcat.replace(/_/g, ' '),
        disp: subcat.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase()),
        rows: [],
      };
    }
    grouped[subcat].rows.push(s);
  }

  // Sort rows within each group by current value descending
  for (const g of Object.values(grouped)) {
    g.rows.sort((a, b) => (b.CURRENT_VALUE || 0) - (a.CURRENT_VALUE || 0));
  }

  return Object.values(grouped);
}


/**
 * Build slide specs from grouped schemes.
 * Each spec = { rowCount, cat, disp, rows[] }
 */
function _buildSlideSpecs(groups) {
  const specs = [];
  for (const g of groups) {
    let rows = g.rows.slice();
    while (rows.length > 0) {
      const n = Math.min(4, rows.length);
      specs.push({
        rowCount: n,
        cat:  g.cat,
        disp: g.disp,
        rows: rows.slice(0, n),
      });
      rows = rows.slice(n);
    }
  }
  return specs;
}


/**
 * Fill a scheme slide with fund data.
 */
function _fillSchemeSlide(slide, spec) {
  // Update category and subcategory labels
  for (const shape of slide.getShapes()) {
    const txt = shape.getText().asString().trim();

    // Category labels are short text on the left side
    if (shape.getLeft() < 200 && txt.length < 45) {
      // Known category names in the template
      const knownCats = new Set([
        'Equity', 'Hybrid', 'Debt', 'Gold & Silver', 'Fund of Funds',
        'Global Funds', 'Solution Oriented',
      ]);
      const knownSubcats = new Set([
        'Flexi Cap', 'Mid Cap', 'Small Cap', 'Large Cap',
        'Value & Contra', 'ELSS', 'Focused Fund', 'Multi Cap',
        'Large & Mid', 'Short Duration', 'Liquid',
      ]);

      if (knownCats.has(txt)) {
        setShapeText(shape, spec.cat);
      } else if (knownSubcats.has(txt)) {
        setShapeText(shape, spec.disp);
      }
    }
  }

  // Fill table data
  const tables = slide.getTables();
  // Tables are typically one per scheme row, each with 1 data row
  // Sort by vertical position
  const tableSorted = tables.slice().sort((a, b) => a.getTop() - b.getTop());

  for (let ri = 0; ri < tableSorted.length && ri < spec.rows.length; ri++) {
    const table = tableSorted[ri];
    const sr = spec.rows[ri];

    try {
      // Each table typically has 1 row (the data row)
      // Columns: Fund Name | Rating | Value | XIRR | Missed Gains
      const row = table.getRow(0);
      const numCells = table.getNumColumns();

      if (numCells >= 1) {
        table.getCell(0, 0).getText().setText(
          sr.FUND_NAME || sr.FUND_STANDARD_NAME || ''
        );
      }
      if (numCells >= 2) {
        // Rating column — just clear it (rating image handled separately)
        table.getCell(0, 1).getText().setText('');
      }
      if (numCells >= 3) {
        table.getCell(0, 2).getText().setText(
          fmtSchemeVal(sr.CURRENT_VALUE || 0, sr['% of PF'] || 0)
        );
      }
      if (numCells >= 4) {
        table.getCell(0, 3).getText().setText(
          fmtXirrPair(sr.XIRR_VALUE, sr.BM_XIRR)
        );
      }
      if (numCells >= 5) {
        table.getCell(0, 4).getText().setText(
          fmtMissed(sr.MG_AS_ON_APP || 0)
        );
      }
    } catch (e) {
      Logger.log('Appendix: table fill error at row ' + ri + ': ' + e);
    }
  }

  // Clear excess tables (template has rows we don't need)
  for (let ri = spec.rows.length; ri < tableSorted.length; ri++) {
    try {
      const table = tableSorted[ri];
      for (let c = 0; c < table.getNumColumns(); c++) {
        table.getCell(0, c).getText().setText('');
      }
    } catch (e) {}
  }

  // Insert rating images
  for (let ri = 0; ri < spec.rows.length; ri++) {
    const sr = spec.rows[ri];
    const rating = sr.POWERRATING;
    if (rating && RATING_IMAGE_IDS[rating]) {
      try {
        const imgBlob = DriveApp.getFileById(RATING_IMAGE_IDS[rating]).getBlob();
        if (ri < tableSorted.length) {
          const table = tableSorted[ri];
          const img = slide.insertImage(imgBlob);
          // Position near the rating column of this row
          img.setLeft(263);  // X_RATING_IMG ÷ 12700
          const rowH = (ri + 1 < tableSorted.length)
            ? tableSorted[ri + 1].getTop() - table.getTop()
            : 43; // ~552450 EMU default
          img.setTop(table.getTop() + (rowH - 20) / 2);
          img.setWidth(20);   // ~255600 EMU
          img.setHeight(20);
        }
      } catch (e) {
        Logger.log('Appendix: rating image error: ' + e);
      }
    }
  }
}
