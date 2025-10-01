// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
  USERNAME: 'frankscobey',
  STOCKFISH_DEPTH: 13,
  STOCKFISH_THREADS: 1,
  STOCKFISH_HASH: 128,
  MAX_GAMES_PER_BATCH: 3,
  BOOK_MOVES_TO_SKIP: 10,
  EVALUATION_TIMEOUT: 10000,
  AUTO_ANALYZE_NEW_GAMES: true, // Automatically analyze new games when fetched
  AUTO_FETCH_CALLBACK_DATA: true // Automatically fetch callback data for new games
};

const SHEETS = {
  GAMES: 'Games',
  ANALYSIS: 'Analysis',
  CALLBACK: 'Callback',
  RATINGS_TIMELINE: 'Ratings Timeline',
  DERIVED: 'Derived Data'
};

const STOCKFISH_API = {
  URL: 'https://chess-api.com/v1',
  METHOD: 'POST',
  USE_LICHESS: false
};

// Result to outcome mapping
const RESULT_MAP = {
  'win': 'Win',
  'checkmated': 'Loss',
  'agreed': 'Draw',
  'repetition': 'Draw',
  'timeout': 'Loss',
  'resigned': 'Loss',
  'stalemate': 'Draw',
  'lose': 'Loss',
  'insufficient': 'Draw',
  '50move': 'Draw',
  'abandoned': 'Loss',
  'kingofthehill': 'Loss',
  'threecheck': 'Loss',
  'timevsinsufficient': 'Draw',
  'bughousepartnerlose': 'Loss'
};

// Termination descriptions
const TERMINATION_MAP = {
  'win': 'Win',
  'checkmated': 'Checkmate',
  'agreed': 'Agreement',
  'repetition': 'Repetition',
  'timeout': 'Timeout',
  'resigned': 'Resignation',
  'stalemate': 'Stalemate',
  'lose': 'Loss',
  'insufficient': 'Insufficient material',
  '50move': '50-move rule',
  'abandoned': 'Abandoned',
  'kingofthehill': 'King of the Hill',
  'threecheck': 'Three-check',
  'timevsinsufficient': 'Timeout vs insufficient',
  'bughousepartnerlose': 'Bughouse partner lost'
};

// ============================================
// MAIN MENU
// ============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('♟️ Chess Analyzer')
    .addItem('1️⃣ Setup Sheets', 'setupSheets')
    .addItem('2️⃣ Initial Fetch (All Games)', 'fetchAllGamesInitial')
    .addItem('3️⃣ Update Recent Games', 'fetchChesscomGames')
    .addSeparator()
    .addItem('📊 Analyze Last 10 Unanalyzed', 'analyzeLast10')
    .addItem('📊 Analyze Last 25 Unanalyzed', 'analyzeLast25')
    .addItem('📊 Analyze Last 50 Unanalyzed', 'analyzeLast50')
    .addSeparator()
    .addItem('📋 Fetch Callback Last 10', 'fetchCallbackLast10')
    .addItem('📋 Fetch Callback Last 25', 'fetchCallbackLast25')
    .addItem('📋 Fetch Callback Last 50', 'fetchCallbackLast50')
    .addSeparator()
    .addItem('📈 Build Ratings Timeline', 'buildRatingsTimeline')
    .addItem('📈 View Analysis Summary', 'showAnalysisSummary')
    .addSeparator()
    .addItem('🧪 Test Stockfish', 'testStockfish')
    .addItem('⚙️ Settings', 'showSettings')
    .addToUi();
}

// ============================================
// SETUP
// ============================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  if (!gamesSheet) {
    gamesSheet = ss.insertSheet(SHEETS.GAMES);
    const headers = [
      'Game URL', 'End Date', 'End Time', 'My Color', 'Opponent',
      'Outcome', 'Termination', 'Format',
      'My Rating', 'Opp Rating', 'Last Rating',
      'Game ID', 'Analyzed', 'Callback Fetched'
    ];
    gamesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    gamesSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    gamesSheet.setFrozenRows(1);
    gamesSheet.setColumnWidth(1, 200);
    
    // Format date and time columns
    gamesSheet.getRange('B:B').setNumberFormat('m"/"d"/"yy');
    gamesSheet.getRange('C:C').setNumberFormat('h:mm AM/PM');
    
    // Add conditional formatting for My Color
    const colorRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('white')
      .setBackground('#FFFFFF')
      .setFontColor('#000000')
      .setRanges([gamesSheet.getRange('D2:D')])
      .build();
    const colorRule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('black')
      .setBackground('#333333')
      .setFontColor('#FFFFFF')
      .setRanges([gamesSheet.getRange('D2:D')])
      .build();
    
    // Conditional formatting for Outcome
    const winRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Win')
      .setBackground('#d9ead3')
      .setRanges([gamesSheet.getRange('F2:F')])
      .build();
    const lossRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Loss')
      .setBackground('#f4cccc')
      .setRanges([gamesSheet.getRange('F2:F')])
      .build();
    const drawRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Draw')
      .setBackground('#fff2cc')
      .setRanges([gamesSheet.getRange('F2:F')])
      .build();
    
    // Conditional formatting for Termination
    const checkmateRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Checkmate')
      .setBackground('#b6d7a8')
      .setRanges([gamesSheet.getRange('G2:G')])
      .build();
    const resignRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Resignation')
      .setBackground('#fce5cd')
      .setRanges([gamesSheet.getRange('G2:G')])
      .build();
    const timeoutRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Timeout')
      .setBackground('#f4cccc')
      .setRanges([gamesSheet.getRange('G2:G')])
      .build();
    const stalemateRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Stalemate')
      .setBackground('#d9d9d9')
      .setRanges([gamesSheet.getRange('G2:G')])
      .build();
    const agreementRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Agreement')
      .setBackground('#d9d9d9')
      .setRanges([gamesSheet.getRange('G2:G')])
      .build();
    const repetitionRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Repetition')
      .setBackground('#d9d9d9')
      .setRanges([gamesSheet.getRange('G2:G')])
      .build();
    
    // Conditional formatting for Format
    const bulletRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Bullet')
      .setBackground('#ea9999')
      .setRanges([gamesSheet.getRange('H2:H')])
      .build();
    const blitzRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Blitz')
      .setBackground('#f9cb9c')
      .setRanges([gamesSheet.getRange('H2:H')])
      .build();
    const rapidRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Rapid')
      .setBackground('#ffe599')
      .setRanges([gamesSheet.getRange('H2:H')])
      .build();
    const dailyRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Daily')
      .setBackground('#b6d7a8')
      .setRanges([gamesSheet.getRange('H2:H')])
      .build();
    const live960Rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('live960')
      .setBackground('#a4c2f4')
      .setRanges([gamesSheet.getRange('H2:H')])
      .build();
    const daily960Rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('daily960')
      .setBackground('#9fc5e8')
      .setRanges([gamesSheet.getRange('H2:H')])
      .build();
    
    // Conditional formatting for booleans
    const trueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=M2=TRUE')
      .setBackground('#d9ead3')
      .setRanges([gamesSheet.getRange('M2:N')])
      .build();
    
    const rules = gamesSheet.getConditionalFormatRules();
    rules.push(colorRule, colorRule2, winRule, lossRule, drawRule, 
               checkmateRule, resignRule, timeoutRule, stalemateRule, agreementRule, repetitionRule,
               bulletRule, blitzRule, rapidRule, dailyRule, live960Rule, daily960Rule,
               trueRule);
    gamesSheet.setConditionalFormatRules(rules);
  }
  
  let derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
  if (!derivedSheet) {
    derivedSheet = ss.insertSheet(SHEETS.DERIVED);
    const headers = [
      'Game ID', 'White Username', 'Black Username', 'White Rating', 'Black Rating',
      'Time Class', 'Time Control', 'Type', 'Base Time', 'Increment', 'Correspondence Time',
      'ECO', 'ECO URL', 'Rated',
      'End', 'Start', 'Start Date', 'Start Time', 'Duration (s)', 'Ply Count', 'Moves',
      'Move List', 'Move Clocks', 'Move Times'
    ];
    derivedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    derivedSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#666666')
      .setFontColor('#ffffff');
    derivedSheet.setFrozenRows(1);
    
    // Format date/time columns
    derivedSheet.getRange('O:O').setNumberFormat('m"/"d"/"yy h:mm AM/PM'); // End
    derivedSheet.getRange('P:P').setNumberFormat('m"/"d"/"yy h:mm AM/PM'); // Start
    derivedSheet.getRange('Q:Q').setNumberFormat('m"/"d"/"yy'); // Start Date
    derivedSheet.getRange('R:R').setNumberFormat('h:mm AM/PM'); // Start Time
    
    // Format Move Clocks and Move Times columns as text to prevent scientific notation
    derivedSheet.getRange('V:V').setNumberFormat('@STRING@'); // Move List
    derivedSheet.getRange('W:W').setNumberFormat('@STRING@'); // Move Clocks
    derivedSheet.getRange('X:X').setNumberFormat('@STRING@'); // Move Times
    
    // Hide the derived sheet
    derivedSheet.hideSheet();
  }
  
  // Apply formatting to existing Games sheet if it exists
  if (gamesSheet) {
    gamesSheet.getRange('B:B').setNumberFormat('m"/"d"/"yy');
    gamesSheet.getRange('C:C').setNumberFormat('h:mm AM/PM');
  }
  
  let analysisSheet = ss.getSheetByName(SHEETS.ANALYSIS);
  if (!analysisSheet) {
    analysisSheet = ss.insertSheet(SHEETS.ANALYSIS);
    const headers = [
      'Game ID', 'Game URL', 'Player', 'Color', 'Opponent', 'Outcome',
      'ECO',
      'Accuracy %', 'Avg CP Loss', 'Estimated ELO',
      'Total Moves', 'Book Moves', 'Analyzed Moves',
      'Best (!!)','Excellent (!)', 'Good', 'Inaccuracy (?!)', 'Mistake (?)', 'Blunder (??)',
      'T1 Rate %', 'T3 Rate %', 'Win→Draw', 'Win→Loss', 'Draw→Loss',
      'Opponent Rating', 'Time Class', 'Time Control', 'Date Analyzed'
    ];
    analysisSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    analysisSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0f9d58')
      .setFontColor('#ffffff');
    analysisSheet.setFrozenRows(1);
  }
  
  let callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  if (!callbackSheet) {
    callbackSheet = ss.insertSheet(SHEETS.CALLBACK);
    const headers = [
      'Game ID', 'Game URL', 'Callback URL', 'End Time', 'My Color', 'Time Class',
      'My Rating', 'Opp Rating', 'My Rating Change', 'Opp Rating Change',
      'My Rating Before', 'Opp Rating Before',
      'Base Time', 'Time Increment', 'Move Timestamps',
      'My Username', 'My Country', 'My Membership', 'My Member Since',
      'My Default Tab', 'My Post Move Action', 'My Location',
      'Opp Username', 'Opp Country', 'Opp Membership', 'Opp Member Since',
      'Opp Default Tab', 'Opp Post Move Action', 'Opp Location',
      'Date Fetched'
    ];
    callbackSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    callbackSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f4b400')
      .setFontColor('#ffffff');
    callbackSheet.setFrozenRows(1);
    
    // Format Move Timestamps column as text to prevent scientific notation
    callbackSheet.getRange('K:K').setNumberFormat('@STRING@');
  }
  
  let ratingsTimelineSheet = ss.getSheetByName(SHEETS.RATINGS_TIMELINE);
  if (!ratingsTimelineSheet) {
    ratingsTimelineSheet = ss.insertSheet(SHEETS.RATINGS_TIMELINE);
    
    // Build compact headers: Rtg, W, L, D, Gms, Dur, ΔR, Perf
    const formats = ['Bul', 'Blz', 'Rpd', 'Dly', '960L', '960D', 'KotH', 'Bug', 'Crzy', '3Chk'];
    const fullFormats = ['Bullet', 'Blitz', 'Rapid', 'Daily', 'live960', 'daily960', 'kingofthehill', 'bughouse', 'crazyhouse', 'threecheck'];
    const headers = ['Date'];
    
    // Add columns for each format (8 columns per format)
    for (let i = 0; i < formats.length; i++) {
      const fmt = formats[i];
      headers.push(
        fmt + ' Rtg',    // Rating
        fmt + ' W',      // Wins
        fmt + ' L',      // Losses
        fmt + ' D',      // Draws
        fmt + ' Gms',    // Games
        fmt + ' Dur',    // Duration
        fmt + ' ΔR',     // Rating change
        fmt + ' Perf'    // Performance ELO
      );
    }
    
    // Add total columns (main 3 formats)
    headers.push('Tot W', 'Tot L', 'Tot D', 'Tot Gms', 'Tot Dur', 'Tot ΔR', 'Tot Perf');
    
    ratingsTimelineSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Header row formatting
    const headerRange = ratingsTimelineSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold')
      .setBackground('#9900ff')
      .setFontColor('#ffffff')
      .setFontSize(9);
    
    ratingsTimelineSheet.setFrozenRows(1);
    ratingsTimelineSheet.setFrozenColumns(1);
    
    // Format date column
    ratingsTimelineSheet.getRange('A:A').setNumberFormat('m"/"d"/"yy');
    
    // Add alternating row colors for readability
    const maxRows = 1000;
    const dataRange = ratingsTimelineSheet.getRange(2, 1, maxRows, headers.length);
    
    // Color coding by format groups (lighter backgrounds)
    // Bullet columns (light red)
    ratingsTimelineSheet.getRange(2, 2, maxRows, 8).setBackground('#fce5cd');
    // Blitz columns (light orange) 
    ratingsTimelineSheet.getRange(2, 10, maxRows, 8).setBackground('#fff2cc');
    // Rapid columns (light yellow)
    ratingsTimelineSheet.getRange(2, 18, maxRows, 8).setBackground('#ffffd9');
    // Daily columns (light green)
    ratingsTimelineSheet.getRange(2, 26, maxRows, 8).setBackground('#d9ead3');
    // Total columns (light purple)
    ratingsTimelineSheet.getRange(2, headers.length - 6, maxRows, 7).setBackground('#d9d2e9');
  }
  
  SpreadsheetApp.getUi().alert('✅ Sheets setup complete!');
}

// ============================================
// TEST STOCKFISH CONNECTION
// ============================================
function testStockfish() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Testing Stockfish connection...', '🧪', -1);
  
  const testFen = 'rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq - 0 1';
  
  try {
    const result = evaluatePositionWithStockfish(testFen);
    
    if (result !== null) {
      SpreadsheetApp.getUi().alert(
        '✅ Stockfish Test Successful!',
        `Test evaluation: ${result.toFixed(2)} pawns\n\n` +
        `Starting position evaluated successfully.\n` +
        `Stockfish is ready to analyze games!`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        '⚠️ Stockfish Connection Issue',
        'Could not get evaluation from Stockfish.\n\n' +
        'Please check your STOCKFISH_API configuration in the script.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      '❌ Stockfish Error',
      `Error: ${error.message}\n\n` +
      'Check the script logs (Extensions → Apps Script → Executions) for details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    Logger.log(`Stockfish test error: ${error}`);
  }
}

// ============================================
// FETCH CHESS.COM GAMES
// ============================================

// INITIAL FETCH: Get all games from all archives
function fetchAllGamesInitial() {
  const username = CONFIG.USERNAME;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!gamesSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" first!');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Initial Full Fetch',
    'This will fetch ALL games from your Chess.com history.\n' +
    'This may take several minutes depending on how many games you have.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    ss.toast('Fetching all game archives...', '⏳', -1);
    
    const archivesUrl = `https://api.chess.com/pub/player/${username}/games/archives`;
    const archivesResponse = UrlFetchApp.fetch(archivesUrl);
    const archives = JSON.parse(archivesResponse.getContentText()).archives;
    
    ss.toast(`Found ${archives.length} archives. Fetching games...`, '⏳', -1);
    
    const allGames = [];
    const props = PropertiesService.getScriptProperties();
    
    const now = new Date();
    const currentYearMonth = `${now.getFullYear()}_${String(now.getMonth() + 1).padStart(2, '0')}`;
    
    for (let i = 0; i < archives.length; i++) {
      ss.toast(`Fetching archive ${i + 1} of ${archives.length}...`, '⏳', -1);
      Utilities.sleep(500);
      
      const response = fetchWithETag(archives[i], null);
      if (response.data) {
        allGames.push(...response.data.games);
        
        // Only store ETag if this is the current month (mutable archive)
        const archiveKey = archiveUrlToKey(archives[i]);
        if (archiveKey === currentYearMonth) {
          props.setProperty('etag_current', response.etag);
        }
        // Don't store ETags for completed months - they're immutable
      }
      
      Logger.log(`Archive ${i + 1}/${archives.length}: ${response.data?.games?.length || 0} games`);
    }
    
    // Filter out duplicates before processing
    const existingGameIds = new Set();
    if (gamesSheet.getLastRow() > 1) {
      const existingData = gamesSheet.getDataRange().getValues();
      for (let i = 1; i < existingData.length; i++) {
        existingGameIds.add(existingData[i][11]); // Game ID column (index 11)
      }
    }
    
    const newGames = allGames.filter(game => {
      const gameId = game.url.split('/').pop();
      return !existingGameIds.has(gameId);
    });
    
    ss.toast(`Processing ${newGames.length} games...`, '⏳', -1);
    const rows = processGamesData(newGames, username);
    
    if (rows.length > 0) {
      const lastRow = gamesSheet.getLastRow();
      gamesSheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
      
      // Find and store the most recent game URL (latest end_time)
      let mostRecentGame = newGames[0];
      for (const game of newGames) {
        if (game.end_time > mostRecentGame.end_time) {
          mostRecentGame = game;
        }
      }
      props.setProperty('LAST_GAME_URL', mostRecentGame.url);
      props.setProperty('INITIAL_FETCH_COMPLETE', 'true');
      
      ss.toast(`✅ Fetched ${newGames.length} games!`, '✅', 5);
      
      // Don't auto-process on initial fetch (too many games)
      // User can manually run "Build Ratings Timeline" and analyze games after
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// UPDATE FETCH: Check most recent archive(s) for new games
function fetchChesscomGames() {
  const username = CONFIG.USERNAME;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!gamesSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" first!');
    return;
  }
  
  const props = PropertiesService.getScriptProperties();
  const initialFetchComplete = props.getProperty('INITIAL_FETCH_COMPLETE');
  
  if (!initialFetchComplete) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'No Initial Fetch Detected',
      'It looks like you haven\'t done an initial full fetch yet.\n\n' +
      'Would you like to do a quick recent fetch (current month)?\n',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
  }
  
  try {
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = String(now.getMonth() + 1).padStart(2, '0');
    const currentYearMonth = `${currentYear}_${currentMonth}`;
    
    // Calculate current month archive URL directly (no API call needed)
    const currentArchiveUrl = `https://api.chess.com/pub/player/${username}/games/${currentYear}/${currentMonth}`;
    
    const archivesToCheck = [];
    const lastKnownGameUrl = props.getProperty('LAST_GAME_URL');
    const allGames = [];
    let foundLastKnownGame = false;
    
    // Check if we need to finalize previous month
    // If last_finalized_month is not the previous month, we should fetch it once
    const prevMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const prevYear = prevMonthDate.getFullYear();
    const prevMonth = String(prevMonthDate.getMonth() + 1).padStart(2, '0');
    const prevYearMonth = `${prevYear}_${prevMonth}`;
    const lastFinalizedMonth = props.getProperty('LAST_FINALIZED_MONTH') || '';
    
    // If previous month hasn't been finalized and it's not current month, fetch it
    if (lastFinalizedMonth !== prevYearMonth && prevYearMonth !== currentYearMonth) {
      const prevArchiveUrl = `https://api.chess.com/pub/player/${username}/games/${prevYear}/${prevMonth}`;
      archivesToCheck.push({url: prevArchiveUrl, key: prevYearMonth, isCurrent: false});
      ss.toast('Finalizing previous month...', '⏳', -1);
    }
    
    // Always check current month
    archivesToCheck.push({url: currentArchiveUrl, key: currentYearMonth, isCurrent: true});
    
    for (const archive of archivesToCheck) {
      Utilities.sleep(500);
      
      // Only use ETag for current month
      const storedETag = archive.isCurrent ? props.getProperty('etag_current') : null;
      
      const response = fetchWithETag(archive.url, storedETag);
      
      if (!response.data) {
        Logger.log(`Archive ${archive.url} not modified (ETag match)`);
        continue;
      }
      
      // Only store ETag if this is current month
      if (archive.isCurrent) {
        props.setProperty('etag_current', response.etag);
      } else {
        // Mark previous month as finalized
        props.setProperty('LAST_FINALIZED_MONTH', archive.key);
      }
      
      const gamesData = response.data.games;
      
      if (lastKnownGameUrl) {
        // Iterate from end to start (newest to oldest)
        for (let i = gamesData.length - 1; i >= 0; i--) {
          const game = gamesData[i];
          
          if (game.url === lastKnownGameUrl) {
            foundLastKnownGame = true;
            break;
          }
          
          allGames.unshift(game); // Add to front to maintain order
        }
        
        if (foundLastKnownGame) break;
      } else {
        allGames.push(...gamesData);
      }
    }
    
    if (allGames.length === 0) {
      ss.toast('No new games found!', 'ℹ️', 3);
      return;
    }
    
    // Filter out duplicates before processing
    const existingGameIds = new Set();
    if (gamesSheet.getLastRow() > 1) {
      const existingData = gamesSheet.getDataRange().getValues();
      for (let i = 1; i < existingData.length; i++) {
        existingGameIds.add(existingData[i][11]); // Game ID column (index 11)
      }
    }
    
    const newGames = allGames.filter(game => {
      const gameId = game.url.split('/').pop();
      return !existingGameIds.has(gameId);
    });
    
    if (newGames.length === 0) {
      ss.toast('No new games found (all were duplicates)!', 'ℹ️', 3);
      return;
    }
    
    ss.toast(`Found ${newGames.length} new games. Processing...`, '⏳', -1);
    const rows = processGamesData(newGames, username);
    
    if (rows.length > 0) {
      const lastRow = gamesSheet.getLastRow();
      gamesSheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
      
      // Find the most recent game from the new games fetched
      let mostRecentGame = newGames[0];
      for (const game of newGames) {
        if (game.end_time > mostRecentGame.end_time) {
          mostRecentGame = game;
        }
      }
      props.setProperty('LAST_GAME_URL', mostRecentGame.url);
      
      ss.toast(`✅ Added ${rows.length} new games!`, '✅', 5);
      
      // Process auto-analysis and callback data for new games
      const gamesToProcess = newGames.map(g => {
        const gameId = g.url.split('/').pop();
        return {
          row: findGameRow(gameId),
          gameId: gameId,
          gameUrl: g.url,
          white: g.white?.username || '',
          black: g.black?.username || '',
          timeClass: getTimeClass(g.time_class),
          outcome: getGameOutcome(g, CONFIG.USERNAME),
          pgn: g.pgn || ''
        };
      }).filter(g => g.row > 0 && g.gameId && g.white && g.black);
      
      processNewGamesAutoFeatures(gamesToProcess);
      
      // Update ratings timeline
      updateRatingsTimeline();
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// Fetch with ETag support
function fetchWithETag(url, etag) {
  const options = {
    muteHttpExceptions: true,
    headers: {}
  };
  
  if (etag) {
    options.headers['If-None-Match'] = etag;
  }
  
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  
  if (code === 304) {
    // Not modified
    return { data: null, etag: etag };
  }
  
  if (code === 200) {
    const newETag = response.getHeaders()['ETag'] || response.getHeaders()['etag'] || '';
    const data = JSON.parse(response.getContentText());
    return { data: data, etag: newETag };
  }
  
  throw new Error(`HTTP ${code}: ${response.getContentText()}`);
}

// Convert archive URL to storage key
function archiveUrlToKey(url) {
  // Extract YYYY/MM from URL like https://api.chess.com/pub/player/username/games/2024/09
  const match = url.match(/(\d{4})\/(\d{2})$/);
  return match ? `${match[1]}_${match[2]}` : url;
}

// Parse time control string into components
function parseTimeControl(timeControl, timeClass) {
  const result = {
    type: timeClass === 'daily' ? 'Daily' : 'Live',
    baseTime: null,
    increment: null,
    correspondenceTime: null
  };
  
  if (!timeControl) return result;
  
  const tcStr = String(timeControl);
  
  // Check if correspondence/daily format (1/value)
  if (tcStr.includes('/')) {
    const parts = tcStr.split('/');
    if (parts.length === 2) {
      result.correspondenceTime = parseInt(parts[1]) || null;
    }
  }
  // Check if live format with increment (value+value)
  else if (tcStr.includes('+')) {
    const parts = tcStr.split('+');
    if (parts.length === 2) {
      result.baseTime = parseInt(parts[0]) || null;
      result.increment = parseInt(parts[1]) || null;
    }
  }
  // Simple live format (just value)
  else {
    result.baseTime = parseInt(tcStr) || null;
    result.increment = 0;
  }
  
  return result;
}

// Helper function to process games data
function processGamesData(games, username) {
  const rows = [];
  const derivedRows = [];
  
  // Sort games by timestamp (oldest first) to ensure Last Rating fills correctly
  const sortedGames = games.slice().sort((a, b) => a.end_time - b.end_time);
  
  // Pre-load existing games data once for performance
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
  let existingGames = [];
  
  if (gamesSheet && gamesSheet.getLastRow() > 1) {
    const data = gamesSheet.getDataRange().getValues();
    // Build lookup map: format -> array of {timestamp, rating}
    for (let i = 1; i < data.length; i++) {
      try {
        const format = data[i][7]; // Format column (index 7)
        const endDate = data[i][1]; // End Date column
        const endTime = data[i][2]; // End Time column
        const myRating = data[i][8]; // My Rating column (index 8)
        const timestamp = new Date(endDate + ' ' + endTime).getTime() / 1000;
        
        existingGames.push({
          format: format,
          timestamp: timestamp,
          rating: myRating
        });
      } catch (error) {
        // Skip malformed rows
        continue;
      }
    }
  }
  
  for (const game of sortedGames) {
    try {
      if (!game || !game.url || !game.end_time) {
        Logger.log('Skipping game with missing data');
        continue;
      }
      
      const endDate = new Date(game.end_time * 1000);
      const gameId = game.url.split('/').pop();
      const eco = extractECOFromPGN(game.pgn);
      const ecoUrl = extractECOUrlFromPGN(game.pgn);
      const outcome = getGameOutcome(game, username);
      const termination = getGameTermination(game, username);
      const format = getGameFormat(game);
      const timeClass = getTimeClass(game.time_class);
      const duration = extractDurationFromPGN(game.pgn);
      
      // Determine my color and opponent
      const isWhite = game.white?.username.toLowerCase() === username.toLowerCase();
      const myColor = isWhite ? 'white' : 'black';
      const opponent = isWhite ? game.black?.username : game.white?.username;
      const myRating = isWhite ? game.white?.rating : game.black?.rating;
      const oppRating = isWhite ? game.black?.rating : game.white?.rating;
      
      // Calculate Last Rating from pre-loaded data AND games processed in this batch
      let lastRating = null;
      let lastGameTime = 0;
      
      // Check existing games from sheet
      for (const existingGame of existingGames) {
        if (existingGame.format === format && 
            existingGame.timestamp < game.end_time && 
            existingGame.timestamp > lastGameTime) {
          lastGameTime = existingGame.timestamp;
          lastRating = existingGame.rating;
        }
      }
      
      // Parse time control
      const tcParsed = parseTimeControl(game.time_control, game.time_class);
      
      // Extract moves with clocks and times
      const moveData = extractMovesWithClocks(game.pgn, tcParsed.baseTime, tcParsed.increment);
      
      // Create proper date/time objects
      // End Date: Set to midnight of the game's date (no time component)
      const endDateObj = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
      
      // End Time: Create a date with just time component (use epoch date + time)
      const endTimeObj = new Date(1970, 0, 1, endDate.getHours(), endDate.getMinutes(), endDate.getSeconds());
      
      // Combined End DateTime for derived sheet
      const endDateTime = new Date(endDate.getTime());
      
      // Calculate start time from duration
      let startDateTime = null;
      let startDateObj = null;
      let startTimeObj = null;
      
      if (duration && duration > 0) {
        startDateTime = new Date(endDateTime.getTime() - (duration * 1000));
        startDateObj = new Date(startDateTime.getFullYear(), startDateTime.getMonth(), startDateTime.getDate());
        startTimeObj = new Date(1970, 0, 1, startDateTime.getHours(), startDateTime.getMinutes(), startDateTime.getSeconds());
      }
      
      rows.push([
        game.url, endDateObj, endTimeObj, myColor, opponent || 'Unknown',
        outcome, termination, format,
        myRating || 'N/A', oppRating || 'N/A', lastRating || 'N/A',
        gameId, false, false
      ]);
      
      // Calculate Moves (ply count / 2 rounded up)
      const movesCount = moveData.plyCount > 0 ? Math.ceil(moveData.plyCount / 2) : 0;
      
      // Store derived data in hidden sheet
      derivedRows.push([
        gameId,
        game.white?.username || 'Unknown',
        game.black?.username || 'Unknown',
        game.white?.rating || 'N/A',
        game.black?.rating || 'N/A',
        timeClass,
        game.time_control || '',
        tcParsed.type,
        tcParsed.baseTime,
        tcParsed.increment,
        tcParsed.correspondenceTime,
        eco,
        ecoUrl,
        game.rated !== undefined ? game.rated : true,
        endDateTime,
        startDateTime,
        startDateObj,
        startTimeObj,
        duration,
        moveData.plyCount,
        movesCount,
        moveData.moveList,
        moveData.clocks,
        moveData.times
      ]);
      
      // Add this game to existingGames for subsequent games in this batch
      existingGames.push({
        format: format,
        timestamp: game.end_time,
        rating: myRating
      });
      
    } catch (error) {
      Logger.log(`Error processing game ${game?.url}: ${error.message}`);
      continue;
    }
  }
  
  // Write derived data to hidden sheet
  if (derivedSheet && derivedRows.length > 0) {
    const lastRow = derivedSheet.getLastRow();
    derivedSheet.getRange(lastRow + 1, 1, derivedRows.length, derivedRows[0].length).setValues(derivedRows);
  }
  
  return rows;
}

// Get game format based on rules and time control
function getGameFormat(game) {
  const rules = game.rules || 'chess';
  const timeClass = game.time_class || '';
  
  if (rules === 'chess') {
    // Use time class for standard chess (Bullet, Blitz, Rapid, Daily)
    return getTimeClass(timeClass);
  } else if (rules === 'chess960') {
    return timeClass === 'daily' ? 'daily960' : 'live960';
  } else {
    // For other variants, return the rules name
    return rules;
  }
}

// Remove duplicate games based on Game ID
function removeDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!gamesSheet) {
    SpreadsheetApp.getUi().alert('❌ Games sheet not found!');
    return;
  }
  
  const data = gamesSheet.getDataRange().getValues();
  const header = data[0];
  const gameIdCol = 11; // Game ID column (index 11)
  
  const seen = new Set();
  const rowsToKeep = [header];
  let duplicateCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const gameId = data[i][gameIdCol];
    
    if (!seen.has(gameId)) {
      seen.add(gameId);
      rowsToKeep.push(data[i]);
    } else {
      duplicateCount++;
    }
  }
  
  if (duplicateCount > 0) {
    gamesSheet.clear();
    gamesSheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
    
    gamesSheet.getRange(1, 1, 1, header.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    gamesSheet.setFrozenRows(1);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Removed ${duplicateCount} duplicate(s)`, 
      '🗑️', 
      3
    );
    Logger.log(`Removed ${duplicateCount} duplicates`);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('No duplicates found!', 'ℹ️', 2);
  }
}

// ============================================
// ANALYZE UNANALYZED GAMES
// ============================================
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
    `This will use Stockfish to analyze each position.\n` +
    `Estimated time: ${Math.ceil(unanalyzedGames.length * 2)} minutes.\n\nContinue?`,
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
  
  // Iterate from newest to oldest (reverse order)
  for (let i = data.length - 1; i >= 1 && games.length < maxCount; i--) {
    if (data[i][12] === false) { // Analyzed column (index 12)
      const myColor = data[i][3];
      const opponent = data[i][4];
      const gameId = data[i][11];
      
      // Get derived data for this game
      const derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
      let pgn = '';
      let timeClass = '';
      let timeControl = '';
      
      if (derivedSheet) {
        const derivedData = derivedSheet.getDataRange().getValues();
        for (let j = 1; j < derivedData.length; j++) {
          if (derivedData[j][0] === gameId) {
            timeClass = derivedData[j][5];
            timeControl = derivedData[j][6];
            // Get move list from derived
            const moveList = derivedData[j][21] || ''; // Move List column
            pgn = moveList; // Use moves list for analysis
            break;
          }
        }
      }
      
      games.push({
        row: i + 1,
        gameUrl: data[i][0],
        endDate: data[i][1],
        endTime: data[i][2],
        white: myColor === 'white' ? CONFIG.USERNAME : opponent,
        black: myColor === 'black' ? CONFIG.USERNAME : opponent,
        outcome: data[i][5],
        termination: data[i][6],
        format: data[i][7],
        timeClass: timeClass,
        timeControl: timeControl,
        whiteRating: myColor === 'white' ? data[i][8] : data[i][9],
        blackRating: myColor === 'black' ? data[i][8] : data[i][9],
        duration: 0, // Will get from derived if needed
        eco: '',
        pgn: pgn,
        gameId: gameId
      });
    }
  }
  
  return games.reverse(); // Return in chronological order (oldest first)
}

function analyzeGames(gamesToAnalyze) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const analysisSheet = ss.getSheetByName(SHEETS.ANALYSIS);
  
  let successCount = 0;
  let errorCount = 0;
  
  for (let i = 0; i < gamesToAnalyze.length; i++) {
    const game = gamesToAnalyze[i];
    
    try {
      ss.toast(`Analyzing game ${i + 1} of ${gamesToAnalyze.length}...`, '♟️', -1);
      
      const analysis = analyzeGameBatch(game);
      
      if (analysis) {
        const playerColor = game.white.toLowerCase() === CONFIG.USERNAME.toLowerCase() ? 'white' : 'black';
        const opponent = playerColor === 'white' ? game.black : game.white;
        const opponentRating = playerColor === 'white' ? game.blackRating : game.whiteRating;
        
        const analysisRow = [
          game.gameId, game.gameUrl, CONFIG.USERNAME, playerColor,
          opponent, game.outcome, game.eco,
          analysis.accuracy, analysis.avgCPLoss, analysis.estimatedELO,
          analysis.totalMoves, analysis.bookMoves, analysis.analyzedMoves,
          analysis.best, analysis.excellent, analysis.good,
          analysis.inaccuracy, analysis.mistake, analysis.blunder,
          analysis.t1Rate, analysis.t3Rate,
          analysis.winToDraw, analysis.winToLoss, analysis.drawToLoss,
          opponentRating, game.timeClass, game.timeControl, new Date()
        ];
        
        const lastRow = analysisSheet.getLastRow();
        analysisSheet.getRange(lastRow + 1, 1, 1, analysisRow.length).setValues([analysisRow]);
        
        const accCell = analysisSheet.getRange(lastRow + 1, 8);
        if (analysis.accuracy >= 90) accCell.setBackground('#d9ead3');
        else if (analysis.accuracy >= 80) accCell.setBackground('#fff2cc');
        else if (analysis.accuracy >= 70) accCell.setBackground('#fce5cd');
        else accCell.setBackground('#f4cccc');
        
        gamesSheet.getRange(game.row, 13).setValue(true); // Analyzed column (index 13)
        
        successCount++;
      }
      
    } catch (error) {
      Logger.log(`Error analyzing game ${game.gameId}: ${error}`);
      Logger.log(`Stack: ${error.stack}`);
      errorCount++;
    }
  }
  
  ss.toast(`✅ Analyzed: ${successCount}, Errors: ${errorCount}`, '✅', 5);
}

// ============================================
// BATCH ANALYSIS ENGINE
// ============================================
function analyzeGameBatch(game) {
  const chess = new ChessGame();
  const moves = extractMovesFromPGN(game.pgn);
  
  Logger.log(`=== Analyzing Game ${game.gameId} ===`);
  Logger.log(`Total moves in PGN: ${moves.length}`);
  
  const playerColor = game.white.toLowerCase() === CONFIG.USERNAME.toLowerCase() ? 'white' : 'black';
  
  const allPositions = [];
  allPositions.push({
    fen: chess.fen(), 
    move: 'start', 
    moveNum: 0, 
    turn: null,
    positionIndex: 0
  });
  
  for (let i = 0; i < moves.length; i++) {
    const moveNum = Math.floor(i / 2) + 1;
    const moveTurn = i % 2 === 0 ? 'white' : 'black';
    
    if (!chess.move(moves[i])) {
      Logger.log(`❌ Failed to apply move ${i}: "${moves[i]}"`);
      break;
    }
    
    allPositions.push({
      fen: chess.fen(), 
      move: moves[i],
      moveNum: moveNum,
      turn: moveTurn,
      positionIndex: i + 1
    });
  }
  
  Logger.log(`Built ${allPositions.length} positions`);
  
  const playerPositions = allPositions.filter(p => 
    p.turn === playerColor && p.moveNum > CONFIG.BOOK_MOVES_TO_SKIP
  );
  
  Logger.log(`Player positions to analyze: ${playerPositions.length}`);
  
  if (playerPositions.length === 0) {
    Logger.log('⚠️ No positions to analyze');
    return null;
  }
  
  const fensToEvaluate = [];
  const evaluationMap = new Map();
  
  for (const playerMove of playerPositions) {
    const posIdx = allPositions.findIndex(p => p.positionIndex === playerMove.positionIndex);
    if (posIdx === -1 || posIdx === 0) continue;
    
    const fenBefore = allPositions[posIdx - 1].fen;
    const fenAfter = allPositions[posIdx].fen;
    
    if (!evaluationMap.has(fenBefore)) {
      fensToEvaluate.push(fenBefore);
      evaluationMap.set(fenBefore, null);
    }
    if (!evaluationMap.has(fenAfter)) {
      fensToEvaluate.push(fenAfter);
      evaluationMap.set(fenAfter, null);
    }
  }
  
  Logger.log(`Unique positions to evaluate: ${fensToEvaluate.length}`);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Evaluating ${fensToEvaluate.length} positions with Stockfish...`, 
    '🔍', 
    -1
  );
  
  for (let i = 0; i < fensToEvaluate.length; i++) {
    const fen = fensToEvaluate[i];
    
    if (i % 5 === 0) {
      Logger.log(`Evaluating ${i+1}/${fensToEvaluate.length}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Evaluating position ${i+1}/${fensToEvaluate.length}...`, 
        '🔍 Stockfish', 
        -1
      );
    }
    
    const evaluation = evaluatePositionWithStockfish(fen);
    evaluationMap.set(fen, evaluation);
    
    if (i < fensToEvaluate.length - 1) {
      Utilities.sleep(300);
    }
  }
  
  Logger.log('✅ All positions evaluated');
  
  const stats = {
    totalCPLoss: 0,
    moveCount: 0,
    best: 0, excellent: 0, good: 0,
    inaccuracy: 0, mistake: 0, blunder: 0,
    t1Count: 0, t3Count: 0,
    winToDraw: 0, winToLoss: 0, drawToLoss: 0
  };
  
  for (const playerMove of playerPositions) {
    const posIdx = allPositions.findIndex(p => p.positionIndex === playerMove.positionIndex);
    if (posIdx === -1 || posIdx === 0) continue;
    
    const fenBefore = allPositions[posIdx - 1].fen;
    const fenAfter = allPositions[posIdx].fen;
    
    const evalBefore = evaluationMap.get(fenBefore);
    const evalAfter = evaluationMap.get(fenAfter);
    
    if (evalBefore === null || evalAfter === null || 
        evalBefore === undefined || evalAfter === undefined ||
        isNaN(evalBefore) || isNaN(evalAfter)) {
      Logger.log(`⚠️ Skipping move ${playerMove.moveNum} - missing evaluation`);
      continue;
    }
    
    const cpLoss = calculateCentipawnLoss(evalBefore, evalAfter, playerColor);
    
    Logger.log(`Move ${playerMove.moveNum}: ${playerMove.move} - Loss: ${cpLoss.toFixed(0)}cp (before: ${evalBefore.toFixed(2)}, after: ${evalAfter.toFixed(2)})`);
    
    stats.totalCPLoss += cpLoss;
    stats.moveCount++;
    
    classifyMove(cpLoss, stats);
    analyzePositionalChange(evalBefore, evalAfter, playerColor, stats);
    checkTopMoveRate(cpLoss, stats);
  }
  
  const totalPlayerMoves = allPositions.filter(p => p.turn === playerColor).length;
  
  const avgCPLoss = stats.moveCount > 0 ? stats.totalCPLoss / stats.moveCount : 0;
  const accuracy = isNaN(avgCPLoss) ? 0 : calculateAccuracy(avgCPLoss);
  const estimatedELO = isNaN(avgCPLoss) ? 950 : estimateELO(avgCPLoss);
  const t1Rate = stats.moveCount > 0 ? (stats.t1Count / stats.moveCount * 100).toFixed(1) : 0;
  const t3Rate = stats.moveCount > 0 ? (stats.t3Count / stats.moveCount * 100).toFixed(1) : 0;
  
  Logger.log(`Final: Accuracy ${accuracy.toFixed(1)}%, Avg CP Loss: ${avgCPLoss.toFixed(1)}, ELO: ${estimatedELO}, Analyzed: ${stats.moveCount} moves`);
  
  if (stats.moveCount === 0) {
    Logger.log('⚠️ No moves could be analyzed');
    return null;
  }
  
  return {
    accuracy: Math.round(accuracy * 10) / 10,
    avgCPLoss: Math.round(avgCPLoss),
    estimatedELO,
    totalMoves: totalPlayerMoves,
    bookMoves: CONFIG.BOOK_MOVES_TO_SKIP,
    analyzedMoves: stats.moveCount,
    best: stats.best,
    excellent: stats.excellent,
    good: stats.good,
    inaccuracy: stats.inaccuracy,
    mistake: stats.mistake,
    blunder: stats.blunder,
    t1Rate,
    t3Rate,
    winToDraw: stats.winToDraw,
    winToLoss: stats.winToLoss,
    drawToLoss: stats.drawToLoss
  };
}

// ============================================
// STOCKFISH EVALUATION
// ============================================
function evaluatePositionWithStockfish(fen) {
  try {
    const url = STOCKFISH_API.URL;
    
    const payload = {
      fen: fen,
      depth: 15
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      headers: {
        'Accept': 'application/json'
      }
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data.eval !== undefined && data.eval !== null) {
        const evaluation = data.eval / 100;
        return Math.max(-10, Math.min(10, evaluation));
      }
      
      if (data.evaluation !== undefined) return parseFloat(data.evaluation);
      if (data.score !== undefined) return parseFloat(data.score) / 100;
      
      Logger.log(`Unknown response: ${response.getContentText().substring(0, 200)}`);
    } else {
      Logger.log(`API error ${response.getResponseCode()}: ${response.getContentText().substring(0, 200)}`);
    }
    
    return evaluateMaterial(fen);
    
  } catch (error) {
    Logger.log(`Evaluation error: ${error.message}`);
    return evaluateMaterial(fen);
  }
}

// ============================================
// FALLBACK: MATERIAL EVALUATION
// ============================================
function evaluateMaterial(fen) {
  const pieces = fen.split(' ')[0];
  const pieceValues = {
    'P': 1, 'N': 3, 'B': 3.25, 'R': 5, 'Q': 9,
    'p': -1, 'n': -3, 'b': -3.25, 'r': -5, 'q': -9
  };
  
  let score = 0;
  for (const char of pieces) {
    if (pieceValues[char]) {
      score += pieceValues[char];
    }
  }
  
  return score;
}

// ============================================
// CALCULATION FUNCTIONS
// ============================================
function calculateCentipawnLoss(evalBefore, evalAfter, playerColor) {
  if (evalBefore === null || evalAfter === null || 
      evalBefore === undefined || evalAfter === undefined ||
      isNaN(evalBefore) || isNaN(evalAfter)) {
    return 0;
  }
  
  let cpBefore = evalBefore * 100;
  let cpAfter = evalAfter * 100;
  
  if (playerColor === 'black') {
    cpBefore = -cpBefore;
    cpAfter = -cpAfter;
  }
  
  cpBefore = Math.max(-1000, Math.min(1000, cpBefore));
  cpAfter = Math.max(-1000, Math.min(1000, cpAfter));
  
  const loss = cpBefore - cpAfter;
  return Math.max(0, loss);
}

function classifyMove(cpLoss, stats) {
  if (cpLoss <= 10) stats.best++;
  else if (cpLoss <= 25) stats.excellent++;
  else if (cpLoss <= 50) stats.good++;
  else if (cpLoss <= 100) stats.inaccuracy++;
  else if (cpLoss <= 300) stats.mistake++;
  else stats.blunder++;
}

function analyzePositionalChange(evalBefore, evalAfter, playerColor, stats) {
  let before = evalBefore * 100;
  let after = evalAfter * 100;
  
  if (playerColor === 'black') {
    before = -before;
    after = -after;
  }
  
  if (before > 200 && Math.abs(after) < 50) stats.winToDraw++;
  if (before > 150 && after < -50) stats.winToLoss++;
  if (Math.abs(before) < 50 && after < -150) stats.drawToLoss++;
}

function checkTopMoveRate(cpLoss, stats) {
  if (cpLoss <= 10) stats.t1Count++;
  if (cpLoss <= 25) stats.t3Count++;
}

function calculateAccuracy(avgCPLoss) {
  const accuracy = 103.1668 * Math.exp(-0.04354 * avgCPLoss);
  return Math.max(0, Math.min(100, accuracy));
}

function estimateELO(avgCPLoss) {
  if (avgCPLoss < 10) return 2600;
  if (avgCPLoss < 20) return 2400;
  if (avgCPLoss < 30) return 2200;
  if (avgCPLoss < 40) return 2000;
  if (avgCPLoss < 55) return 1850;
  if (avgCPLoss < 70) return 1700;
  if (avgCPLoss < 90) return 1550;
  if (avgCPLoss < 115) return 1400;
  if (avgCPLoss < 145) return 1250;
  if (avgCPLoss < 180) return 1100;
  return 950;
}

// Calculate performance ELO based on results against opponents
function calculatePerformanceELO(wins, losses, draws, avgOppRating) {
  if (!avgOppRating || avgOppRating === 'N/A') return null;
  
  const totalGames = wins + losses + draws;
  if (totalGames === 0) return null;
  
  const score = (wins + 0.5 * draws) / totalGames;
  
  // Performance rating formula: Avg Opp Rating + 400 * (2 * score - 1) / sqrt(games)
  // Simplified: Perf = Avg Opp + dp where dp depends on score percentage
  let dp = 0;
  if (score >= 1.0) dp = 800;
  else if (score >= 0.99) dp = 677;
  else if (score >= 0.9) dp = 366;
  else if (score >= 0.8) dp = 240;
  else if (score >= 0.7) dp = 149;
  else if (score >= 0.6) dp = 72;
  else if (score >= 0.5) dp = 0;
  else if (score >= 0.4) dp = -72;
  else if (score >= 0.3) dp = -149;
  else if (score >= 0.2) dp = -240;
  else if (score >= 0.1) dp = -366;
  else if (score >= 0.01) dp = -677;
  else dp = -800;
  
  return Math.round(parseFloat(avgOppRating) + dp);
}

// ============================================
// HELPER FUNCTIONS
// ============================================
function extractMovesFromPGN(pgn) {
  const moveSection = pgn.split(/\n\n/)[1] || pgn;
  
  return moveSection
    .replace(/\{[^}]*\}/g, '')
    .replace(/\([^)]*\)/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\$\d+/g, '')
    .replace(/\d+\.{3}/g, '')
    .replace(/\d+\./g, '')
    .replace(/[!?+#]+/g, '')
    .trim()
    .split(/\s+/)
    .filter(m => m && m !== '*' && !m.match(/^(1-0|0-1|1\/2-1\/2)$/));
}

function extractECOFromPGN(pgn) {
  if (!pgn) return '';
  const match = pgn.match(/\[ECO "([^"]+)"\]/);
  return match ? match[1] : '';
}

// Extract ECO URL from PGN
function extractECOUrlFromPGN(pgn) {
  if (!pgn) return '';
  const match = pgn.match(/\[ECOUrl "([^"]+)"\]/);
  return match ? match[1] : '';
}

// Extract moves with clock times from PGN
function extractMovesWithClocks(pgn, baseTime, increment) {
  if (!pgn) return { moves: [], clocks: [], times: [] };
  
  const moveSection = pgn.split(/\n\n/)[1] || pgn;
  const moves = [];
  const clocks = [];
  const times = [];
  
  // Regex to match move and its clock: "e4 {[%clk 0:02:59.9]}"
  const movePattern = /([NBRQK]?[a-h]?[1-8]?x?[a-h][1-8](?:=[NBRQK])?|O-O(?:-O)?)\s*\{?\[%clk\s+(\d+):(\d+):(\d+)(?:\.(\d+))?\]?\}?/g;
  
  let match;
  let prevClock = [baseTime || 0, baseTime || 0]; // [white, black] previous clocks
  let moveIndex = 0;
  
  while ((match = movePattern.exec(moveSection)) !== null) {
    const move = match[1];
    const hours = parseInt(match[2]) || 0;
    const minutes = parseInt(match[3]) || 0;
    const seconds = parseInt(match[4]) || 0;
    const deciseconds = parseInt(match[5]) || 0;
    
    // Convert clock to total seconds
    const clockSeconds = hours * 3600 + minutes * 60 + seconds + deciseconds / 10;
    
    moves.push(move);
    clocks.push(clockSeconds);
    
    // Calculate time spent on this move
    const playerIndex = moveIndex % 2; // 0 = white, 1 = black
    const prevPlayerClock = prevClock[playerIndex];
    
    // Time spent = previous clock - current clock + increment
    let timeSpent = prevPlayerClock - clockSeconds + (increment || 0);
    
    // Minimum move time is 0.1 seconds (Chess.com enforces this)
    if (timeSpent < 0.1) timeSpent = 0.1;
    
    times.push(Math.round(timeSpent * 10) / 10); // Round to 1 decimal
    
    // Update previous clock for this player
    prevClock[playerIndex] = clockSeconds;
    
    moveIndex++;
  }
  
  return { 
    moveList: moves.join(', '), 
    clocks: clocks.join(', '), 
    times: times.join(', '),
    plyCount: moves.length
  };
}

function extractDurationFromPGN(pgn) {
  if (!pgn) return null;
  
  const dateMatch = pgn.match(/\[UTCDate "([^"]+)"\]/);
  const timeMatch = pgn.match(/\[UTCTime "([^"]+)"\]/);
  const endDateMatch = pgn.match(/\[EndDate "([^"]+)"\]/);
  const endTimeMatch = pgn.match(/\[EndTime "([^"]+)"\]/);
  
  if (!dateMatch || !timeMatch || !endDateMatch || !endTimeMatch) {
    return null;
  }
  
  try {
    const startDateParts = dateMatch[1].split('.');
    const startTimeParts = timeMatch[1].split(':');
    const startDate = new Date(Date.UTC(
      parseInt(startDateParts[0]),
      parseInt(startDateParts[1]) - 1,
      parseInt(startDateParts[2]),
      parseInt(startTimeParts[0]),
      parseInt(startTimeParts[1]),
      parseInt(startTimeParts[2])
    ));
    
    const endDateParts = endDateMatch[1].split('.');
    const endTimeParts = endTimeMatch[1].split(':');
    const endDate = new Date(Date.UTC(
      parseInt(endDateParts[0]),
      parseInt(endDateParts[1]) - 1,
      parseInt(endDateParts[2]),
      parseInt(endTimeParts[0]),
      parseInt(endTimeParts[1]),
      parseInt(endTimeParts[2])
    ));
    
    const durationMs = endDate.getTime() - startDate.getTime();
    return Math.round(durationMs / 1000);
  } catch (error) {
    Logger.log(`Error parsing duration: ${error.message}`);
    return null;
  }
}

function getGameOutcome(game, username) {
  if (!game || !game.white || !game.black) return 'Unknown';
  
  const isWhite = game.white.username?.toLowerCase() === username.toLowerCase();
  const myResult = isWhite ? game.white.result : game.black.result;
  
  if (!myResult) return 'Unknown';
  
  return RESULT_MAP[myResult] || 'Unknown';
}

function getGameTermination(game, username) {
  if (!game || !game.white || !game.black) return 'Unknown';
  
  const isWhite = game.white.username?.toLowerCase() === username.toLowerCase();
  const myResult = isWhite ? game.white.result : game.black.result;
  const opponentResult = isWhite ? game.black.result : game.white.result;
  
  if (!myResult) return 'Unknown';
  
  // If I won, use opponent's result for termination
  if (myResult === 'win') {
    return TERMINATION_MAP[opponentResult] || opponentResult;
  }
  
  // Otherwise use my result
  return TERMINATION_MAP[myResult] || myResult;
}

function getTimeClass(timeClass) {
  if (timeClass === 'bullet') return 'Bullet';
  if (timeClass === 'blitz') return 'Blitz';
  if (timeClass === 'rapid') return 'Rapid';
  if (timeClass === 'daily') return 'Daily';
  return timeClass || 'Unknown';
}

// ============================================
// UI FUNCTIONS
// ============================================
function showAnalysisSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analysisSheet = ss.getSheetByName(SHEETS.ANALYSIS);
  
  if (!analysisSheet) {
    SpreadsheetApp.getUi().alert('❌ No analysis data found.');
    return;
  }
  
  const data = analysisSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('⚠️ No games analyzed yet.');
    return;
  }
  
  const totalGames = data.length - 1;
  const avgAccuracy = data.slice(1).reduce((sum, row) => sum + row[7], 0) / totalGames;
  const avgCPLoss = data.slice(1).reduce((sum, row) => sum + row[8], 0) / totalGames;
  const avgELO = data.slice(1).reduce((sum, row) => sum + row[9], 0) / totalGames;
  const avgT1Rate = data.slice(1).reduce((sum, row) => sum + parseFloat(row[19]), 0) / totalGames;
  
  let totalBlunders = 0, totalMistakes = 0, totalInaccuracies = 0;
  for (let i = 1; i < data.length; i++) {
    totalBlunders += data[i][18];
    totalMistakes += data[i][17];
    totalInaccuracies += data[i][16];
  }
  
  const message = `📊 Analysis Summary\n\n` +
    `Games Analyzed: ${totalGames}\n` +
    `━━━━━━━━━━━━━━━━━━━━━━\n` +
    `Average Accuracy: ${avgAccuracy.toFixed(1)}%\n` +
    `Average CP Loss: ${avgCPLoss.toFixed(1)}\n` +
    `Average Est. ELO: ${avgELO.toFixed(0)}\n` +
    `Top-1 Move Rate: ${avgT1Rate.toFixed(1)}%\n` +
    `━━━━━━━━━━━━━━━━━━━━━━\n` +
    `Total Blunders: ${totalBlunders}\n` +
    `Total Mistakes: ${totalMistakes}\n` +
    `Total Inaccuracies: ${totalInaccuracies}`;
  
  SpreadsheetApp.getUi().alert(message);
}

function showSettings() {
  const props = PropertiesService.getScriptProperties();
  const initialFetchComplete = props.getProperty('INITIAL_FETCH_COMPLETE') || 'false';
  const lastGameUrl = props.getProperty('LAST_GAME_URL') || 'None';
  
  const lastFinalizedMonth = props.getProperty('LAST_FINALIZED_MONTH') || 'None';
  
  const message = `⚙️ Current Settings:\n\n` +
    `Username: ${CONFIG.USERNAME}\n` +
    `Stockfish Depth: ${CONFIG.STOCKFISH_DEPTH}\n` +
    `Book Moves to Skip: ${CONFIG.BOOK_MOVES_TO_SKIP}\n` +
    `Max Games per Batch: ${CONFIG.MAX_GAMES_PER_BATCH}\n` +
    `━━━━━━━━━━━━━━━━━━━━━━\n` +
    `Initial Fetch Complete: ${initialFetchComplete}\n` +
    `Last Game URL: ${lastGameUrl.substring(0, 50)}...\n` +
    `Last Finalized Month: ${lastFinalizedMonth}\n` +
    `━━━━━━━━━━━━━━━━━━━━━━\n` +
    `Engine: Stockfish via API\n` +
    `API Endpoint: ${STOCKFISH_API.URL}\n` +
    `ETag Support: Current Month Only\n\n` +
    `To change settings, edit the CONFIG object in the script.`;
  
  SpreadsheetApp.getUi().alert(message);
}

// ============================================
// CHESS GAME ENGINE
// ============================================
function ChessGame() {
  let board = [];
  let turn = 'w';
  let castling = {w: {k: true, q: true}, b: {k: true, q: true}};
  let enPassant = null;
  let halfmove = 0;
  let fullmove = 1;
  
  this.reset = function() {
    board = [
      ['r','n','b','q','k','b','n','r'],
      ['p','p','p','p','p','p','p','p'],
      [null,null,null,null,null,null,null,null],
      [null,null,null,null,null,null,null,null],
      [null,null,null,null,null,null,null,null],
      [null,null,null,null,null,null,null,null],
      ['P','P','P','P','P','P','P','P'],
      ['R','N','B','Q','K','B','N','R']
    ];
    turn = 'w';
    castling = {w: {k: true, q: true}, b: {k: true, q: true}};
    enPassant = null;
    halfmove = 0;
    fullmove = 1;
  };
  
  this.reset();
  
  this.fen = function() {
    let fen = '';
    
    for (let rank = 0; rank < 8; rank++) {
      let empty = 0;
      for (let file = 0; file < 8; file++) {
        const piece = board[rank][file];
        if (piece === null) {
          empty++;
        } else {
          if (empty > 0) {
            fen += empty;
            empty = 0;
          }
          fen += piece;
        }
      }
      if (empty > 0) fen += empty;
      if (rank < 7) fen += '/';
    }
    
    fen += ' ' + turn;
    
    let castleStr = '';
    if (castling.w.k) castleStr += 'K';
    if (castling.w.q) castleStr += 'Q';
    if (castling.b.k) castleStr += 'k';
    if (castling.b.q) castleStr += 'q';
    fen += ' ' + (castleStr || '-');
    
    fen += ' ' + (enPassant || '-');
    fen += ' ' + halfmove + ' ' + fullmove;
    
    return fen;
  };
  
  this.turn = function() {
    return turn;
  };
  
  this.move = function(moveStr) {
    try {
      const move = parseMove(moveStr);
      if (!move) {
        Logger.log(`Failed to parse: ${moveStr}`);
        return false;
      }
      
      const {from, to, promotion, piece, isCastle} = move;
      const movingPiece = board[from.rank][from.file];
      
      if (!movingPiece) {
        Logger.log(`No piece at ${coordToAlgebraic(from)}`);
        return false;
      }
      
      if (isCastle) {
        const isKingside = to.file > from.file;
        if (turn === 'w') {
          board[7][4] = null;
          board[7][to.file] = 'K';
          if (isKingside) {
            board[7][7] = null;
            board[7][5] = 'R';
          } else {
            board[7][0] = null;
            board[7][3] = 'R';
          }
        } else {
          board[0][4] = null;
          board[0][to.file] = 'k';
          if (isKingside) {
            board[0][7] = null;
            board[0][5] = 'r';
          } else {
            board[0][0] = null;
            board[0][3] = 'r';
          }
        }
      } else {
        if (piece === 'p' && enPassant && coordToAlgebraic(to) === enPassant && board[to.rank][to.file] === null) {
          const captureRank = turn === 'w' ? to.rank + 1 : to.rank - 1;
          board[captureRank][to.file] = null;
        }
        
        board[to.rank][to.file] = promotion || movingPiece;
        board[from.rank][from.file] = null;
      }
      
      enPassant = null;
      if (piece === 'p' && Math.abs(to.rank - from.rank) === 2) {
        const epRank = turn === 'w' ? from.rank - 1 : from.rank + 1;
        enPassant = coordToAlgebraic({rank: epRank, file: from.file});
      }
      
      if (piece === 'k') {
        if (turn === 'w') {
          castling.w.k = false;
          castling.w.q = false;
        } else {
          castling.b.k = false;
          castling.b.q = false;
        }
      }
      if (piece === 'r') {
        if (turn === 'w') {
          if (from.file === 7 && from.rank === 7) castling.w.k = false;
          if (from.file === 0 && from.rank === 7) castling.w.q = false;
        } else {
          if (from.file === 7 && from.rank === 0) castling.b.k = false;
          if (from.file === 0 && from.rank === 0) castling.b.q = false;
        }
      }
      
      const isCapture = board[to.rank][to.file] !== movingPiece;
      if (piece === 'p' || isCapture) {
        halfmove = 0;
      } else {
        halfmove++;
      }
      
      turn = turn === 'w' ? 'b' : 'w';
      if (turn === 'w') fullmove++;
      
      return true;
    } catch (e) {
      Logger.log(`Move error: ${e.message} for ${moveStr}`);
      return false;
    }
  };
  
  function parseMove(moveStr) {
    moveStr = moveStr.replace(/[+#!?]/g, '').trim();
    
    if (moveStr === 'O-O' || moveStr === '0-0') {
      return parseCastling(true);
    }
    if (moveStr === 'O-O-O' || moveStr === '0-0-0') {
      return parseCastling(false);
    }
    
    let workingMove = moveStr;
    let promotion = null;
    
    if (workingMove.includes('=')) {
      const parts = workingMove.split('=');
      workingMove = parts[0];
      const promoPiece = parts[1][0].toLowerCase();
      promotion = turn === 'w' ? promoPiece.toUpperCase() : promoPiece;
    }
    
    const toSquare = workingMove.slice(-2);
    const to = algebraicToCoord(toSquare);
    if (!to) {
      Logger.log(`Invalid destination square: ${toSquare}`);
      return null;
    }
    
    let piece = 'p';
    let fromHints = {file: null, rank: null};
    
    if (/^[NBRQK]/.test(moveStr)) {
      piece = moveStr[0].toLowerCase();
      
      let hintStr = moveStr.substring(1).replace('x', '').slice(0, -2);
      
      if (/[a-h]/.test(hintStr)) {
        const fileMatch = hintStr.match(/[a-h]/);
        fromHints.file = fileMatch[0].charCodeAt(0) - 'a'.charCodeAt(0);
      }
      
      if (/[1-8]/.test(hintStr)) {
        const rankMatch = hintStr.match(/[1-8]/);
        fromHints.rank = 8 - parseInt(rankMatch[0]);
      }
    } 
    else if (/^[a-h]/.test(moveStr)) {
      piece = 'p';
      if (moveStr.includes('x')) {
        fromHints.file = moveStr[0].charCodeAt(0) - 'a'.charCodeAt(0);
      }
    }
    
    const from = findSourceSquare(piece, to, fromHints);
    if (!from) {
      Logger.log(`Could not find source for ${moveStr}: piece=${piece}, to=${toSquare}, hints=${JSON.stringify(fromHints)}`);
      return null;
    }
    
    return {from, to, promotion, piece, isCastle: false};
  }
  
  function findSourceSquare(piece, to, hints) {
    const targetPiece = turn === 'w' ? piece.toUpperCase() : piece.toLowerCase();
    const candidates = [];
    
    for (let rank = 0; rank < 8; rank++) {
      for (let file = 0; file < 8; file++) {
        if (board[rank][file] === targetPiece) {
          if (hints.file !== null && file !== hints.file) continue;
          if (hints.rank !== null && rank !== hints.rank) continue;
          
          if (canPieceReach(piece, {rank, file}, to)) {
            candidates.push({rank, file});
          }
        }
      }
    }
    
    if (candidates.length > 1) {
      Logger.log(`Ambiguous move: ${candidates.length} candidates for ${piece} to ${coordToAlgebraic(to)}`);
    }
    
    return candidates.length > 0 ? candidates[0] : null;
  }
  
  function canPieceReach(piece, from, to) {
    const dr = to.rank - from.rank;
    const df = to.file - from.file;
    const targetSquare = board[to.rank][to.file];
    
    if (targetSquare !== null) {
      const targetColor = targetSquare === targetSquare.toUpperCase() ? 'w' : 'b';
      if (targetColor === turn) return false;
    }
    
    switch (piece) {
      case 'p':
        if (turn === 'w') {
          if (df === 0 && dr === -1 && targetSquare === null) return true;
          if (df === 0 && dr === -2 && from.rank === 6 && targetSquare === null && board[to.rank + 1][to.file] === null) return true;
          if (Math.abs(df) === 1 && dr === -1 && (targetSquare !== null || coordToAlgebraic(to) === enPassant)) return true;
        } else {
          if (df === 0 && dr === 1 && targetSquare === null) return true;
          if (df === 0 && dr === 2 && from.rank === 1 && targetSquare === null && board[to.rank - 1][to.file] === null) return true;
          if (Math.abs(df) === 1 && dr === 1 && (targetSquare !== null || coordToAlgebraic(to) === enPassant)) return true;
        }
        return false;
        
      case 'n':
        return (Math.abs(dr) === 2 && Math.abs(df) === 1) || (Math.abs(dr) === 1 && Math.abs(df) === 2);
        
      case 'b':
        if (Math.abs(dr) !== Math.abs(df)) return false;
        return isPathClear(from, to);
        
      case 'r':
        if (dr !== 0 && df !== 0) return false;
        return isPathClear(from, to);
        
      case 'q':
        if (dr !== 0 && df !== 0 && Math.abs(dr) !== Math.abs(df)) return false;
        return isPathClear(from, to);
        
      case 'k':
        return Math.abs(dr) <= 1 && Math.abs(df) <= 1;
    }
    
    return false;
  }
  
  function isPathClear(from, to) {
    const dr = Math.sign(to.rank - from.rank);
    const df = Math.sign(to.file - from.file);
    
    let r = from.rank + dr;
    let f = from.file + df;
    
    while (r !== to.rank || f !== to.file) {
      if (board[r][f] !== null) return false;
      r += dr;
      f += df;
    }
    
    return true;
  }
  
  function parseCastling(isKingside) {
    if (turn === 'w') {
      return {
        from: {rank: 7, file: 4},
        to: {rank: 7, file: isKingside ? 6 : 2},
        piece: 'k',
        promotion: null,
        isCastle: true
      };
    } else {
      return {
        from: {rank: 0, file: 4},
        to: {rank: 0, file: isKingside ? 6 : 2},
        piece: 'k',
        promotion: null,
        isCastle: true
      };
    }
  }
  
  function algebraicToCoord(square) {
    if (!square || square.length !== 2) return null;
    const file = square.charCodeAt(0) - 'a'.charCodeAt(0);
    const rank = 8 - parseInt(square[1]);
    if (file < 0 || file > 7 || rank < 0 || rank > 7 || isNaN(rank)) return null;
    return {rank, file};
  }
  
  function coordToAlgebraic(coord) {
    const file = String.fromCharCode('a'.charCodeAt(0) + coord.file);
    const rank = 8 - coord.rank;
    return file + rank;
  }
}

// ============================================
// CALLBACK DATA FETCHING
// ============================================
function fetchCallbackData(game) {
  // Validate game object has required fields
  if (!game || !game.gameId || !game.timeClass || !game.white || !game.black) {
    Logger.log(`Skipping callback fetch - incomplete game data: ${JSON.stringify(game)}`);
    return null;
  }
  
  const gameId = game.gameId;
  const timeClass = game.timeClass.toLowerCase();
  const gameType = timeClass === 'daily' ? 'daily' : 'live';
  const callbackUrl = `https://www.chess.com/callback/${gameType}/game/${gameId}`;
  
  try {
    const response = UrlFetchApp.fetch(callbackUrl, {muteHttpExceptions: true});
    
    if (response.getResponseCode() !== 200) {
      Logger.log(`Callback API error for game ${gameId}: ${response.getResponseCode()}`);
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data || !data.game) {
      Logger.log(`Invalid callback data for game ${gameId}`);
      return null;
    }
    
    const gameData = data.game;
    const players = data.players || {};
    const topPlayer = players.top || {};
    const bottomPlayer = players.bottom || {};
    
    // Determine my color and player data
    const isWhite = game.white.toLowerCase() === CONFIG.USERNAME.toLowerCase();
    const myColor = isWhite ? 'white' : 'black';
    
    let myRatingChange = isWhite ? gameData.ratingChangeWhite : gameData.ratingChangeBlack;
    let oppRatingChange = isWhite ? gameData.ratingChangeBlack : gameData.ratingChangeWhite;
    
    // If rating change is 0, it's likely an error (unless edge case draw)
    // Set to null to indicate unreliable data
    if (myRatingChange === 0) myRatingChange = null;
    if (oppRatingChange === 0) oppRatingChange = null;
    
    // Get player data (top/bottom can be either color)
    let whitePlayer, blackPlayer;
    if (topPlayer.color === 'white') {
      whitePlayer = topPlayer;
      blackPlayer = bottomPlayer;
    } else {
      whitePlayer = bottomPlayer;
      blackPlayer = topPlayer;
    }
    
    // Determine my player and opponent player
    const myPlayer = isWhite ? whitePlayer : blackPlayer;
    const oppPlayer = isWhite ? blackPlayer : whitePlayer;
    
    // Get ratings from callback
    const myRating = myPlayer.rating || null;
    const oppRating = oppPlayer.rating || null;
    
    // Calculate "before" ratings by subtracting rating change
    let myRatingBefore = null;
    let oppRatingBefore = null;
    
    if (myRating !== null && myRatingChange !== null) {
      myRatingBefore = myRating - myRatingChange;
    }
    if (oppRating !== null && oppRatingChange !== null) {
      oppRatingBefore = oppRating - oppRatingChange;
    }
    
    return {
      gameId: gameId,
      gameUrl: game.gameUrl,
      callbackUrl: callbackUrl,
      endTime: gameData.endTime,
      myColor: myColor,
      timeClass: game.timeClass,
      myRating: myRating,
      oppRating: oppRating,
      myRatingChange: myRatingChange,
      oppRatingChange: oppRatingChange,
      myRatingBefore: myRatingBefore,
      oppRatingBefore: oppRatingBefore,
      baseTime: gameData.baseTime1 || 0,
      timeIncrement: gameData.timeIncrement1 || 0,
      moveTimestamps: gameData.moveTimestamps ? String(gameData.moveTimestamps) : '',
      myUsername: myPlayer.username || '',
      myCountry: myPlayer.countryName || '',
      myMembership: myPlayer.membershipCode || '',
      myMemberSince: myPlayer.memberSince || 0,
      myDefaultTab: myPlayer.defaultTab || null,
      myPostMoveAction: myPlayer.postMoveAction || '',
      myLocation: myPlayer.location || '',
      oppUsername: oppPlayer.username || '',
      oppCountry: oppPlayer.countryName || '',
      oppMembership: oppPlayer.membershipCode || '',
      oppMemberSince: oppPlayer.memberSince || 0,
      oppDefaultTab: oppPlayer.defaultTab || null,
      oppPostMoveAction: oppPlayer.postMoveAction || '',
      oppLocation: oppPlayer.location || ''
    };
    
  } catch (error) {
    Logger.log(`Error fetching callback data for game ${gameId}: ${error.message}`);
    return null;
  }
}

function saveCallbackData(callbackData) {
  if (!callbackData) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  
  if (!callbackSheet) return;
  
  const row = [
    callbackData.gameId,
    callbackData.gameUrl,
    callbackData.callbackUrl,
    callbackData.endTime,
    callbackData.myColor,
    callbackData.timeClass,
    callbackData.myRating,
    callbackData.oppRating,
    callbackData.myRatingChange,
    callbackData.oppRatingChange,
    callbackData.myRatingBefore,
    callbackData.oppRatingBefore,
    callbackData.baseTime,
    callbackData.timeIncrement,
    callbackData.moveTimestamps,
    callbackData.myUsername,
    callbackData.myCountry,
    callbackData.myMembership,
    callbackData.myMemberSince,
    callbackData.myDefaultTab,
    callbackData.myPostMoveAction,
    callbackData.myLocation,
    callbackData.oppUsername,
    callbackData.oppCountry,
    callbackData.oppMembership,
    callbackData.oppMemberSince,
    callbackData.oppDefaultTab,
    callbackData.oppPostMoveAction,
    callbackData.oppLocation,
    new Date()
  ];
  
  const lastRow = callbackSheet.getLastRow();
  callbackSheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
}

function processNewGamesAutoFeatures(newGames) {
  if (!newGames || newGames.length === 0) return;
  
  // Auto-fetch callback data
  if (CONFIG.AUTO_FETCH_CALLBACK_DATA && newGames.length <= CONFIG.MAX_GAMES_PER_BATCH) {
    fetchCallbackForGames(newGames);
  }
  
  // Auto-analyze new games
  if (CONFIG.AUTO_ANALYZE_NEW_GAMES && newGames.length <= CONFIG.MAX_GAMES_PER_BATCH) {
    analyzeGames(newGames);
  }
}

// ============================================
// FETCH CALLBACK DATA FOR GAMES
// ============================================
function fetchCallbackLast10() { fetchCallbackLastN(10); }
function fetchCallbackLast25() { fetchCallbackLastN(25); }
function fetchCallbackLast50() { fetchCallbackLastN(50); }

function fetchCallbackLastN(count) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  
  if (!gamesSheet || !callbackSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" first!');
    return;
  }
  
  const gamesWithoutCallback = getGamesWithoutCallback(count);
  
  if (gamesWithoutCallback.length === 0) {
    SpreadsheetApp.getUi().alert('✅ No games need callback data!');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Fetch callback data for ${gamesWithoutCallback.length} game(s)?`,
    `This will fetch detailed game data from Chess.com.\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  fetchCallbackForGames(gamesWithoutCallback);
}

function getGamesWithoutCallback(maxCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const data = gamesSheet.getDataRange().getValues();
  const games = [];
  
  // Iterate from newest to oldest (reverse order)
  for (let i = data.length - 1; i >= 1 && games.length < maxCount; i--) {
    if (data[i][13] === false) { // Callback Fetched column (index 13)
      const myColor = data[i][3];
      const opponent = data[i][4];
      const gameId = data[i][11];
      
      // Get time class from derived data
      const derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
      let timeClass = '';
      
      if (derivedSheet) {
        const derivedData = derivedSheet.getDataRange().getValues();
        for (let j = 1; j < derivedData.length; j++) {
          if (derivedData[j][0] === gameId) {
            timeClass = derivedData[j][5];
            break;
          }
        }
      }
      
      games.push({
        row: i + 1,
        gameId: gameId,
        gameUrl: data[i][0],
        white: myColor === 'white' ? CONFIG.USERNAME : opponent,
        black: myColor === 'black' ? CONFIG.USERNAME : opponent,
        timeClass: timeClass
      });
    }
  }
  
  return games.reverse(); // Return in chronological order (oldest first)
}

function fetchCallbackForGames(gamesToFetch) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  let successCount = 0;
  let errorCount = 0;
  
  ss.toast('Fetching callback data...', '📋', -1);
  
  for (let i = 0; i < gamesToFetch.length; i++) {
    const game = gamesToFetch[i];
    
    try {
      ss.toast(`Fetching callback ${i + 1} of ${gamesToFetch.length}...`, '📋', -1);
      
      const callbackData = fetchCallbackData(game);
      if (callbackData) {
        saveCallbackData(callbackData);
        
        // Mark callback as fetched
        if (game.row) {
          gamesSheet.getRange(game.row, 14).setValue(true); // Callback Fetched column (index 14)
        }
        
        successCount++;
      } else {
        errorCount++;
      }
      
      Utilities.sleep(300); // Rate limiting
      
    } catch (error) {
      Logger.log(`Error fetching callback for game ${game.gameId}: ${error}`);
      errorCount++;
    }
  }
  
  ss.toast(`✅ Callback fetched: ${successCount}, Errors: ${errorCount}`, '📋', 5);
  
  // Update ratings timeline for affected dates
  if (successCount > 0) {
    updateTimelineForGames(gamesToFetch);
  }
}

function updateTimelineForGames(games) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timelineSheet = ss.getSheetByName(SHEETS.RATINGS_TIMELINE);
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  
  if (!timelineSheet || !gamesSheet || games.length === 0) return;
  
  ss.toast('Updating timeline for affected dates...', '📈', -1);
  
  // Get all dates affected by these games
  const affectedDates = new Set();
  const gamesData = gamesSheet.getDataRange().getValues();
  
  for (const game of games) {
    for (let i = 1; i < gamesData.length; i++) {
      if (gamesData[i][11] === game.gameId) {
        const endDate = gamesData[i][1];
        let dateStr;
        if (typeof endDate === 'string') {
          dateStr = endDate;
        } else if (endDate instanceof Date) {
          dateStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          dateStr = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        affectedDates.add(dateStr);
        break;
      }
    }
  }
  
  if (affectedDates.size === 0) return;
  
  Logger.log(`Updating ${affectedDates.size} affected dates in timeline`);
  
  // Get earliest affected date - need to recalc from there forward for forward-fill
  const sortedDates = Array.from(affectedDates).sort();
  const earliestAffected = sortedDates[0];
  
  // Get callback data
  const callbackData = callbackSheet ? callbackSheet.getDataRange().getValues() : [];
  const callbackLookup = {};
  for (let i = 1; i < callbackData.length; i++) {
    const gameId = callbackData[i][0];
    callbackLookup[gameId] = {
      myRating: callbackData[i][6],
      myRatingChange: callbackData[i][8],
      myRatingBefore: callbackData[i][10]
    };
  }
  
  // Get timeline data
  const timelineData = timelineSheet.getDataRange().getValues();
  
  // Find the row for earliest affected date
  let startRow = -1;
  for (let i = 1; i < timelineData.length; i++) {
    const dateInSheet = timelineData[i][0];
    const dateStr = typeof dateInSheet === 'string' ? dateInSheet : 
                    Utilities.formatDate(new Date(dateInSheet), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (dateStr === earliestAffected) {
      startRow = i + 1; // Convert to 1-based
      break;
    }
  }
  
  if (startRow === -1) {
    Logger.log(`Earliest affected date ${earliestAffected} not found in timeline`);
    return;
  }
  
  // Get previous day's ratings to start forward fill
  const formats = ['Bullet', 'Blitz', 'Rapid', 'Daily', 'live960', 'daily960', 'kingofthehill', 'bughouse', 'crazyhouse', 'threecheck'];
  const lastKnownRating = {};
  
  if (startRow > 2) {
    const prevDayRatings = timelineSheet.getRange(startRow - 1, 2, 1, 10).getValues()[0];
    formats.forEach((format, index) => {
      lastKnownRating[format] = prevDayRatings[index] || null;
    });
  } else {
    formats.forEach(format => lastKnownRating[format] = null);
  }
  
  // Build game events index by date for affected dates and forward
  const gamesByDate = {};
  for (let i = 1; i < gamesData.length; i++) {
    const endDate = gamesData[i][1];
    let dateStr;
    if (typeof endDate === 'string') {
      dateStr = endDate;
    } else if (endDate instanceof Date) {
      dateStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      dateStr = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
    if (dateStr < earliestAffected) continue; // Skip dates before affected range
    
    const format = gamesData[i][7];
    const myRating = gamesData[i][8];
    const gameId = gamesData[i][11];
    
    if (!endDate || myRating === 'N/A') continue;
    
    if (!gamesByDate[dateStr]) gamesByDate[dateStr] = [];
    gamesByDate[dateStr].push({ format, myRating, gameId });
  }
  
  // Update timeline from startRow to end
  const rowsToUpdate = [];
  for (let i = startRow; i <= timelineData.length; i++) {
    const dateInSheet = timelineData[i - 1][0];
    const dateStr = typeof dateInSheet === 'string' ? dateInSheet : 
                    Utilities.formatDate(new Date(dateInSheet), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // Process games for this date
    const gamesOnDate = gamesByDate[dateStr] || [];
    for (const game of gamesOnDate) {
      if (callbackLookup[game.gameId]) {
        const callback = callbackLookup[game.gameId];
        if (callback.myRatingChange !== null && callback.myRatingBefore !== null) {
          lastKnownRating[game.format] = callback.myRatingBefore + callback.myRatingChange;
        } else if (callback.myRating !== null) {
          lastKnownRating[game.format] = callback.myRating;
        } else {
          lastKnownRating[game.format] = game.myRating;
        }
      } else {
        lastKnownRating[game.format] = game.myRating;
      }
    }
    
    // Build row with forward-filled ratings
    const row = [];
    formats.forEach(format => row.push(lastKnownRating[format] || ''));
    rowsToUpdate.push(row);
  }
  
  // Write updated rows back to sheet
  if (rowsToUpdate.length > 0) {
    timelineSheet.getRange(startRow, 2, rowsToUpdate.length, 10).setValues(rowsToUpdate);
    Logger.log(`Updated ${rowsToUpdate.length} rows in timeline starting from row ${startRow}`);
  }
  
  ss.toast(`✅ Updated timeline for ${affectedDates.size} affected dates`, '📈', 3);
}

function findGameRow(gameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const data = gamesSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === gameId) { // Game ID column (index 11)
      return i + 1;
    }
  }
  return -1;
}

// ============================================
// RATINGS TIMELINE
// ============================================
function buildRatingsTimeline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  const timelineSheet = ss.getSheetByName(SHEETS.RATINGS_TIMELINE);
  
  if (!gamesSheet || !timelineSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" first!');
    return;
  }
  
  ss.toast('Building ratings timeline...', '📈', -1);
  
  // Get all games data
  const gamesData = gamesSheet.getDataRange().getValues();
  const callbackData = callbackSheet ? callbackSheet.getDataRange().getValues() : [];
  
  Logger.log(`Games sheet has ${gamesData.length} rows (including header)`);
  if (gamesData.length > 1) {
    Logger.log(`Sample game row: ${JSON.stringify(gamesData[1])}`);
  }
  
  // Build callback lookup by game ID
  const callbackLookup = {};
  for (let i = 1; i < callbackData.length; i++) {
    const gameId = callbackData[i][0];
    callbackLookup[gameId] = {
      myRating: callbackData[i][6],
      myRatingChange: callbackData[i][8],
      myRatingBefore: callbackData[i][10]
    };
  }
  
  // Find earliest date (first game or account creation)
  let earliestDate = new Date();
  for (let i = 1; i < gamesData.length; i++) {
    const gameDate = new Date(gamesData[i][1]); // End Date column
    if (gameDate < earliestDate) {
      earliestDate = gameDate;
    }
  }
  
  // Get today's date
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  earliestDate.setHours(0, 0, 0, 0);
  
  // Build timeline data structure
  // Use exact format names as they appear in the Format column
  const formats = ['Bullet', 'Blitz', 'Rapid', 'Daily', 'live960', 'daily960', 'kingofthehill', 'bughouse', 'crazyhouse', 'threecheck'];
  const timeline = {};
  
  // Initialize all dates from earliest to today
  const currentDate = new Date(earliestDate);
  while (currentDate <= today) {
    const dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    timeline[dateStr] = {};
    formats.forEach(format => {
      timeline[dateStr][format] = {
        rating: null,
        wins: 0,
        losses: 0,
        draws: 0,
        games: 0,
        duration: 0,
        oppRatings: []
      };
    });
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  Logger.log(`Timeline initialized: ${Object.keys(timeline).length} dates from ${earliestDate} to ${today}`);
  
  // Build derived data lookup for performance
  const derivedLookup = {};
  const derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
  if (derivedSheet) {
    const derivedData = derivedSheet.getDataRange().getValues();
    for (let j = 1; j < derivedData.length; j++) {
      const gameId = derivedData[j][0];
      derivedLookup[gameId] = {
        duration: derivedData[j][18] || 0
      };
    }
  }
  
  Logger.log(`Built derived lookup with ${Object.keys(derivedLookup).length} games`);
  
  // Process each game to build events with full data
  const gameEvents = [];
  for (let i = 1; i < gamesData.length; i++) {
    const endDate = gamesData[i][1]; // End Date column (index 1)
    const endTime = gamesData[i][2]; // End Time column (index 2)
    const format = gamesData[i][7]; // Format column (index 7)
    const myRating = gamesData[i][8]; // My Rating column (index 8)
    const oppRating = gamesData[i][9]; // Opp Rating column (index 9)
    const outcome = gamesData[i][5]; // Outcome column (index 5)
    const gameId = gamesData[i][11]; // Game ID column (index 11)
    
    // Get duration from lookup (much faster than nested loop)
    const duration = derivedLookup[gameId] ? derivedLookup[gameId].duration : 0;
    
    if (i === 1) {
      Logger.log(`First game: endDate=${endDate}, format=${format}, myRating=${myRating}, gameId=${gameId}`);
    }
    
    if (!endDate || myRating === 'N/A') {
      if (i <= 5) Logger.log(`Skipping row ${i}: endDate=${endDate}, myRating=${myRating}`);
      continue;
    }
    
    // Handle both string and Date object formats for endDate
    let dateStr;
    if (typeof endDate === 'string') {
      dateStr = endDate;
    } else if (endDate instanceof Date) {
      dateStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      dateStr = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
    // Build proper timestamp
    const timestamp = new Date(dateStr + ' ' + endTime).getTime();
    
    gameEvents.push({
      dateStr: dateStr,
      timestamp: timestamp,
      format: format,
      myRating: myRating,
      oppRating: oppRating,
      outcome: outcome,
      duration: duration || 0,
      gameId: gameId
    });
  }
  
  // Sort by timestamp
  gameEvents.sort((a, b) => a.timestamp - b.timestamp);
  
  Logger.log(`Found ${gameEvents.length} game events to process`);
  if (gameEvents.length > 0) {
    Logger.log(`Sample game event: ${JSON.stringify(gameEvents[0])}`);
  }
  
  // Apply ratings to timeline with forward fill
  const lastKnownRating = {};
  formats.forEach(format => lastKnownRating[format] = null);
  
  for (const dateStr of Object.keys(timeline).sort()) {
    // Apply any games that happened on this date
    const gamesOnDate = gameEvents.filter(g => g.dateStr === dateStr);
    
    for (const game of gamesOnDate) {
      const format = game.format;
      const gameId = game.gameId;
      
      // Update rating with priority: callback > game rating
      if (callbackLookup[gameId]) {
        const callback = callbackLookup[gameId];
        if (callback.myRatingChange !== null && callback.myRatingBefore !== null) {
          lastKnownRating[format] = callback.myRatingBefore + callback.myRatingChange;
        } else if (callback.myRating !== null) {
          lastKnownRating[format] = callback.myRating;
        } else {
          lastKnownRating[format] = game.myRating;
        }
      } else {
        lastKnownRating[format] = game.myRating;
      }
      
      // Update stats for this format on this date
      const stats = timeline[dateStr][format];
      stats.rating = lastKnownRating[format];
      stats.games++;
      stats.duration += game.duration;
      
      // Track opponent ratings
      if (game.oppRating && game.oppRating !== 'N/A') {
        stats.oppRatings.push(parseFloat(game.oppRating));
      }
      
      // Count W-L-D
      if (game.outcome === 'Win') stats.wins++;
      else if (game.outcome === 'Loss') stats.losses++;
      else if (game.outcome === 'Draw') stats.draws++;
    }
    
    // Forward fill ratings for formats with no games today
    for (const format of formats) {
      if (timeline[dateStr][format].rating === null) {
        timeline[dateStr][format].rating = lastKnownRating[format];
      }
    }
  }
  
  // Build rows for sheet with all stats
  const rows = [];
  const sortedDates = Object.keys(timeline).sort();
  
  // Track previous day's ratings for delta calculation
  const prevDayRating = {};
  formats.forEach(format => prevDayRating[format] = null);
  
  for (let idx = 0; idx < sortedDates.length; idx++) {
    const dateStr = sortedDates[idx];
    const dateObj = new Date(dateStr);
    const row = [dateObj];
    
    // Track totals for main 3 formats (Bullet, Blitz, Rapid)
    let totalWins = 0, totalLosses = 0, totalDraws = 0, totalGames = 0, totalDuration = 0;
    const totalOppRatings = [];
    
    for (const format of formats) {
      const stats = timeline[dateStr][format];
      const rating = stats.rating;
      
      // Calculate rating delta
      const ratingDelta = (rating !== null && prevDayRating[format] !== null) 
        ? rating - prevDayRating[format] : null;
      
      // Calculate performance ELO
      const avgOppRating = stats.oppRatings.length > 0 
        ? stats.oppRatings.reduce((a, b) => a + b, 0) / stats.oppRatings.length 
        : null;
      const perfELO = calculatePerformanceELO(stats.wins, stats.losses, stats.draws, avgOppRating);
      
      // Add to row: Rating, W, L, D, Games, Duration, Δ Rating, Perf ELO
      row.push(
        rating || '',
        stats.wins || '',
        stats.losses || '',
        stats.draws || '',
        stats.games || '',
        stats.duration || '',
        ratingDelta || '',
        perfELO || ''
      );
      
      // Update previous day rating
      if (rating !== null) {
        prevDayRating[format] = rating;
      }
      
      // Add to totals if main 3 formats
      if (format === 'Bullet' || format === 'Blitz' || format === 'Rapid') {
        totalWins += stats.wins;
        totalLosses += stats.losses;
        totalDraws += stats.draws;
        totalGames += stats.games;
        totalDuration += stats.duration;
        totalOppRatings.push(...stats.oppRatings);
      }
    }
    
    // Add total columns (main 3 formats only)
    const totalAvgOppRating = totalOppRatings.length > 0 
      ? totalOppRatings.reduce((a, b) => a + b, 0) / totalOppRatings.length 
      : null;
    const totalPerfELO = calculatePerformanceELO(totalWins, totalLosses, totalDraws, totalAvgOppRating);
    
    // Calculate total rating delta (sum of all 3 formats)
    let totalRatingDelta = null;
    const ratingDeltas = [];
    for (const format of ['Bullet', 'Blitz', 'Rapid']) {
      const stats = timeline[dateStr][format];
      if (stats.rating !== null && prevDayRating[format] !== null) {
        ratingDeltas.push(stats.rating - prevDayRating[format]);
      }
    }
    if (ratingDeltas.length > 0) {
      totalRatingDelta = ratingDeltas.reduce((a, b) => a + b, 0);
    }
    
    row.push(
      totalWins || '',
      totalLosses || '',
      totalDraws || '',
      totalGames || '',
      totalDuration || '',
      totalRatingDelta || '',
      totalPerfELO || ''
    );
    
    rows.push(row);
  }
  
  // Clear existing data (except header)
  if (timelineSheet.getLastRow() > 1) {
    timelineSheet.getRange(2, 1, timelineSheet.getLastRow() - 1, timelineSheet.getLastColumn()).clear();
  }
  
  // Write timeline data
  if (rows.length > 0) {
    timelineSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    Logger.log(`Timeline written: ${rows.length} rows`);
    if (rows.length > 0) {
      Logger.log(`Sample row: ${JSON.stringify(rows[0])}`);
    }
  } else {
    Logger.log('No timeline rows to write!');
  }
  
  ss.toast(`✅ Built timeline: ${rows.length} days`, '📈', 5);
}

function updateRatingsTimeline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timelineSheet = ss.getSheetByName(SHEETS.RATINGS_TIMELINE);
  const props = PropertiesService.getScriptProperties();
  
  if (!timelineSheet) return;
  
  const initialFetchComplete = props.getProperty('INITIAL_FETCH_COMPLETE');
  const timelineLastBuilt = props.getProperty('TIMELINE_LAST_BUILT');
  
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const lastRow = timelineSheet.getLastRow();
  
  if (lastRow < 2) {
    // No data yet, build from scratch
    buildRatingsTimeline();
    props.setProperty('TIMELINE_LAST_BUILT', today);
    return;
  }
  
  // If initial fetch is complete, we know all dates before last fetched game are immutable
  // Only rebuild if:
  // 1. Today's date doesn't exist yet (new day)
  // 2. Timeline was never built
  // 3. New games were added (could affect today or recent dates)
  
  const lastDateInSheet = timelineSheet.getRange(lastRow, 1).getValue();
  const lastDateStr = Utilities.formatDate(new Date(lastDateInSheet), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  if (lastDateStr !== today || !timelineLastBuilt) {
    // Need to add new dates or rebuild
    buildRatingsTimeline();
    props.setProperty('TIMELINE_LAST_BUILT', today);
  } else {
    // Today already exists and timeline was built - just update today's row
    updateTodayRatings();
  }
}

function updateTodayRatings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timelineSheet = ss.getSheetByName(SHEETS.RATINGS_TIMELINE);
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  
  if (!timelineSheet || !gamesSheet) return;
  
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const lastRow = timelineSheet.getLastRow();
  const lastDateInSheet = timelineSheet.getRange(lastRow, 1).getValue();
  const lastDateStr = Utilities.formatDate(new Date(lastDateInSheet), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  if (lastDateStr !== today) return; // Today not in sheet yet
  
  // Get previous day's ratings
  const prevDayRatings = lastRow > 2 ? timelineSheet.getRange(lastRow - 1, 2, 1, 10).getValues()[0] : [];
  
  // Get today's games to update ratings
  const gamesData = gamesSheet.getDataRange().getValues();
  const callbackData = callbackSheet ? callbackSheet.getDataRange().getValues() : [];
  
  const callbackLookup = {};
  for (let i = 1; i < callbackData.length; i++) {
    const gameId = callbackData[i][0];
    callbackLookup[gameId] = {
      myRating: callbackData[i][6],
      myRatingChange: callbackData[i][8],
      myRatingBefore: callbackData[i][10]
    };
  }
  
  const formats = ['Bullet', 'Blitz', 'Rapid', 'Daily', 'live960', 'daily960', 'kingofthehill', 'bughouse', 'crazyhouse', 'threecheck'];
  const todayRatings = {};
  
  // Initialize with previous day's ratings
  formats.forEach((format, index) => {
    todayRatings[format] = prevDayRatings[index] || null;
  });
  
  // Process today's games
  for (let i = 1; i < gamesData.length; i++) {
    const endDate = gamesData[i][1];
    let dateStr;
    if (typeof endDate === 'string') {
      dateStr = endDate;
    } else if (endDate instanceof Date) {
      dateStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      dateStr = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
    if (dateStr !== today) continue;
    
    const format = gamesData[i][7];
    const myRating = gamesData[i][8];
    const gameId = gamesData[i][11];
    
    if (myRating === 'N/A') continue;
    
    // Use callback data if available
    if (callbackLookup[gameId]) {
      const callback = callbackLookup[gameId];
      if (callback.myRatingChange !== null && callback.myRatingBefore !== null) {
        todayRatings[format] = callback.myRatingBefore + callback.myRatingChange;
      } else if (callback.myRating !== null) {
        todayRatings[format] = callback.myRating;
      } else {
        todayRatings[format] = myRating;
      }
    } else {
      todayRatings[format] = myRating;
    }
  }
  
  // Update today's row
  const row = [today];
  formats.forEach(format => row.push(todayRatings[format] || ''));
  timelineSheet.getRange(lastRow, 1, 1, row.length).setValues([row]);
}
