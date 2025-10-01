// CONFIGURATION
const CONFIG = {
  USERNAME: 'frankscobey',
  STOCKFISH_DEPTH: 13,
  STOCKFISH_THREADS: 1,
  STOCKFISH_HASH: 128,
  MAX_GAMES_PER_BATCH: 3,
  BOOK_MOVES_TO_SKIP: 10,
  EVALUATION_TIMEOUT: 10000,
  AUTO_ANALYZE_NEW_GAMES: true,
  AUTO_FETCH_CALLBACK_DATA: true,
  CRITICAL_TIME_FACTOR: 2.0 // Tag moves as "critical" if time spent > (factor * avgMoveTime)
};

const SHEETS = {
  GAMES: 'Games',
  ANALYSIS: 'Analysis',
  DERIVED: 'Derived Data',
  CALLBACK: 'Callback',
  RATINGS_TIMELINE: 'Ratings Timeline'
};

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
