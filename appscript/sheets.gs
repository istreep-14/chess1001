function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let analysisSheet = ss.getSheetByName(SHEETS.ANALYSIS);
  if (!analysisSheet) {
    analysisSheet = ss.insertSheet(SHEETS.ANALYSIS);
    const headers = [
      'Game ID', 'Game URL', 'Player', 'Color', 'Opponent', 'Outcome', 'Opening',
      'Accuracy %', 'Avg CP Loss', 'Estimated ELO', 'Total Moves',
      'Opening Range', 'Middlegame Range', 'Endgame Range',
      'Avg Move Time (s)', 'Time Variance', 'Critical Moves',
      'Blunders', 'Mistakes', 'Inaccuracies', 'Best Moves'
    ];
    analysisSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    analysisSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0f9d58')
      .setFontColor('#ffffff');
    analysisSheet.setFrozenRows(1);
  }
  // Add similar setup for other sheets (GAMES, DERIVED, CALLBACK, RATINGS_TIMELINE) as in your original code.
}
