/**
 * RiskReward.gs — Insert risk-reward slides from a separate template deck.
 *
 * The risk-reward deck has 4 slides per risk profile (5 profiles x 4 = 20 slides).
 * We copy the 4 slides matching the client's risk profile into the main deck
 * at positions 15-18 (indices 14-17).
 *
 * In Apps Script, this is much simpler than Python — we can directly copy slides
 * between presentations using SlidesApp.
 */

function doRiskRewardSlides(presentation, riskProfile, goals) {
  goals = goals || [];

  let rrPres;
  try {
    rrPres = SlidesApp.openById(M2_RISK_REWARD_DECK_ID);
  } catch (e) {
    Logger.log('Risk Reward: could not open deck — ' + e);
    return;
  }

  const startIdx = RISK_REWARD_IDX[riskProfile];
  if (startIdx === undefined) {
    Logger.log('Risk Reward: unknown profile "' + riskProfile + '"');
    return;
  }

  const rrSlides = rrPres.getSlides();
  const mainSlides = presentation.getSlides();

  let count = 0;
  for (let offset = 0; offset < 4; offset++) {
    const dstIdx = 14 + offset;
    const srcIdx = startIdx + offset;

    if (dstIdx >= mainSlides.length || srcIdx >= rrSlides.length) break;

    try {
      // Strategy: Replace content of the destination slide with source slide content.
      // SlidesApp doesn't have a direct "replace slide content" method,
      // so we append the source slide, then remove the old one.

      // Append source slide from risk-reward deck
      const srcSlide = rrSlides[srcIdx];
      presentation.appendSlide(srcSlide);

      // The new slide is now at the end — move it to dstIdx + 1 (1-based)
      const updatedSlides = presentation.getSlides();
      const newSlide = updatedSlides[updatedSlides.length - 1];
      newSlide.move(dstIdx + 1); // 1-based position

      // Remove the old slide that was at this position (now shifted to dstIdx + 1)
      const afterMove = presentation.getSlides();
      afterMove[dstIdx + 1].remove();

      // Fill goals on the newly inserted slide
      if (goals.length > 0) {
        _fillRRGoals(presentation.getSlides()[dstIdx], goals);
      }

      count++;
    } catch (e) {
      Logger.log('Risk Reward: WARNING slide ' + (srcIdx + 1) + ': ' + e);
    }
  }

  Logger.log('Risk Reward: replaced ' + count + ' slides for "' + riskProfile +
    '" (source idx ' + startIdx + '-' + (startIdx + count - 1) + ')');
}


/**
 * Fill goal placeholders on a risk-reward slide.
 */
function _fillRRGoals(slide, goals) {
  const primary   = goals[0] || 'Wealth Creation';
  const secondary = goals.slice(1).join(', ');

  for (const shape of slide.getShapes()) {
    const txt = shape.getText().asString();

    // Placeholder pattern: {{main_goal}} / {{secondary_goal}}
    if (txt.includes('{{main_goal}}') || txt.includes('{{secondary_goal}}')) {
      shape.getText().replaceAllText('{{main_goal}}', primary);
      shape.getText().replaceAllText('{{secondary_goal}}', secondary);
      Logger.log('Risk Reward: goals placeholder -> "' + primary + '" / "' + secondary + '"');
    }

    // Hardcoded goal text patterns
    if (txt.trim() === 'Financial Freedom' || txt.trim() === 'Wealth Growth') {
      if (secondary) {
        setShapeText(shape, primary + '\n' + secondary);
      } else {
        setShapeText(shape, primary);
      }
      Logger.log('Risk Reward: hardcoded goals -> "' + primary + '"');
    }
  }

  // Also check inside groups
  for (const group of slide.getGroups()) {
    try {
      for (const child of group.getChildren()) {
        if (child.asShape) {
          const shape = child.asShape();
          const txt = shape.getText().asString();
          if (txt.includes('{{main_goal}}') || txt.includes('{{secondary_goal}}')) {
            shape.getText().replaceAllText('{{main_goal}}', primary);
            shape.getText().replaceAllText('{{secondary_goal}}', secondary);
          }
        }
      }
    } catch (e) {}
  }
}
