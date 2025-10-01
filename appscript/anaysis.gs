function analyzeGames(gamesToAnalyze) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const analysisSheet = ss.getSheetByName(SHEETS.ANALYSIS);
  const derivedSheet = ss.getSheetByName(SHEETS.DERIVED);

  let successCount = 0, errorCount = 0;

  for (let i = 0; i < gamesToAnalyze.length; i++) {
    try {
      ss.toast(`Analyzing game ${i + 1} of ${gamesToAnalyze.length}...`, '♟️', -1);

      const game = gamesToAnalyze[i];
      const derivedRow = findDerivedRow(game.gameId, derivedSheet);
      if (!derivedRow) throw new Error('Derived data missing for game: ' + game.gameId);

      const moveList = parseMoveList(derivedRow[21]);
      const moveTimes = parseTimes(derivedRow[23]);
      const baseTime = Number(derivedRow[8]) || 0;
      const increment = Number(derivedRow[9]) || 0;
      const oppRating = Number(derivedRow[4]) || Number(game.blackRating) || Number(game.oppRating) || 1500;

      const avgMoveTime = moveTimes.length > 0 ? moveTimes.reduce((a, b) => a + b, 0) / moveTimes.length : 0;
      const timeVariance = moveTimes.length > 0
        ? moveTimes.reduce((a, b) => a + Math.pow(b - avgMoveTime, 2), 0) / moveTimes.length
        : 0;
      const criticalMoves = moveTimes
        .map((t, idx) => (t > avgMoveTime * CONFIG.CRITICAL_TIME_FACTOR ? idx + 1 : null))
        .filter(x => x !== null);

      const phaseRanges = detectPhases(moveList);

      const analysisResult = analyzeGameWithStockfish(moveList, game, phaseRanges);

      const estELO = estimateELO(
        analysisResult.accuracy,
        avgMoveTime,
        timeVariance,
        oppRating,
        analysisResult.complexityScore
      );

      const analysisRow = [
        game.gameId, game.gameUrl, CONFIG.USERNAME, game.myColor,
        game.opponent, game.outcome, analysisResult.opening,
        analysisResult.accuracy, analysisResult.avgCPLoss, estELO,
        moveList.length,
        phaseRanges.opening.join('-'), phaseRanges.middlegame.join('-'), phaseRanges.endgame.join('-'),
        avgMoveTime.toFixed(2), timeVariance.toFixed(2),
        criticalMoves.length > 0 ? criticalMoves.join(', ') : '',
        analysisResult.blunders, analysisResult.mistakes, analysisResult.inaccuracies, analysisResult.bestMoves
      ];
      const lastRow = analysisSheet.getLastRow();
      analysisSheet.getRange(lastRow + 1, 1, 1, analysisRow.length).setValues([analysisRow]);

      gamesSheet.getRange(game.row, 13).setValue(true); // Mark as analyzed
      successCount++;
    } catch (error) {
      Logger.log(`Error analyzing game ${gamesToAnalyze[i]?.gameId}: ${error}`);
      errorCount++;
    }
  }
  ss.toast(`✅ Analyzed: ${successCount}, Errors: ${errorCount}`, '✅', 5);
}

function findDerivedRow(gameId, derivedSheet) {
  if (!derivedSheet) return null;
  const data = derivedSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(gameId)) return data[i];
  }
  return null;
}

function parseMoveList(moveListStr) {
  return moveListStr
    .split(',')
    .map(m => m.trim())
    .filter(m => m.length > 0);
}

function parseTimes(timesStr) {
  return timesStr
    .split(',')
    .map(t => Number(t.trim()))
    .filter(t => !isNaN(t));
}

function detectPhases(moveList) {
  const openingEnd = Math.min(moveList.length, 12);
  let endgameStart = moveList.length;
  for (let i = 0; i < moveList.length; i++) {
    if (i > 24) {
      endgameStart = i + 1;
      break;
    }
  }
  return {
    opening: [1, openingEnd],
    middlegame: [openingEnd + 1, endgameStart - 1],
    endgame: [endgameStart, moveList.length]
  };
}

function analyzeGameWithStockfish(moveList, game, phaseRanges) {
  // MOCK: Replace with real analysis engine output
  return {
    opening: "Sicilian Defense: Najdorf Variation",
    accuracy: 92.5,
    avgCPLoss: 22,
    complexityScore: 1.2,
    blunders: 1,
    mistakes: 2,
    inaccuracies: 3,
    bestMoves: 15
  };
}

function estimateELO(accuracy, avgMoveTime, variance, oppRating, complexityScore) {
  let baseElo = 1000 + accuracy * 10;
  baseElo += oppRating * 0.2;
  baseElo += (complexityScore - 1) * 50;
  baseElo -= variance * 2;
  return Math.round(baseElo);
}
