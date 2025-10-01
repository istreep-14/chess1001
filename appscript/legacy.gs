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
  if (myResult === 'win') {
    return TERMINATION_MAP[opponentResult] || opponentResult;
  }
  return TERMINATION_MAP[myResult] || myResult;
}

function getTimeClass(timeClass) {
  if (timeClass === 'bullet') return 'Bullet';
  if (timeClass === 'blitz') return 'Blitz';
  if (timeClass === 'rapid') return 'Rapid';
  if (timeClass === 'daily') return 'Daily';
  return timeClass || 'Unknown';
}
