function analyzeLast10() { analyzeLastNUnanalyzed(10); }
function analyzeLast25() { analyzeLastNUnanalyzed(25); }
function analyzeLast50() { analyzeLastNUnanalyzed(50); }

function analyzeLastNUnanalyzed(count) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const analysisSheet = ss.getSheetByName(SHEETS.ANALYSIS);

  if (!gamesSheet || !analysisSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" first!');
    return;
  }

  const unanalyzedGames = getUnanalyzedGames(count);
  if (unanalyzedGames.length === 0) {
    SpreadsheetApp.getUi().alert('✅ No unanalyzed games found!');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Analyze ${unanalyzedGames.length} game(s)?`,
    `This will use Stockfish to analyze each position.\nEstimated time: ${Math.ceil(unanalyzedGames.length * 2)} minutes.\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  analyzeGames(unanalyzedGames);
}

function getUnanalyzedGames(maxCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const data = gamesSheet.getDataRange().getValues();
  const games = [];
  for (let i = data.length - 1; i >= 1 && games.length < maxCount; i--) {
    if (data[i][12] === false) {
      games.push({
        row: i + 1,
        gameUrl: data[i][0],
        myColor: data[i][3],
        opponent: data[i][4],
        outcome: data[i][5],
        gameId: data[i][11],
        blackRating: data[i][9]
      });
    }
  }
  return games.reverse();
}
