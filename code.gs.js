// ============================================================
// WHEEL BET APP — Apps Script Backend v10 (patched)
// Refactored from v6: batched writes, column constants, unique
// bet IDs (col P), LockService, score throttle, DRY helpers.
// Preserves submitted_by (col N) and submitted_at (col O).
// ============================================================

// ── CONFIG (move to PropertiesService when ready) ──
var API_KEY     = 'zTNTf9l2VG75AtHlMWTVsw';

var GROUPME_BOT_ID = 'aa666f31040e31b9e435664693';
var EVENT_ID    = '12274879400071151280';
var TIMEZONE    = 'America/Chicago';

// ── COLUMN INDICES: DailyBets (A–P) ──
var B = {
  DATE: 0, W_P1: 1, W_P2: 2, C_P1: 3, C_P2: 4,
  W_FRONT: 5, W_BACK: 6, C_FRONT: 7, C_BACK: 8,
  FRONT_RESULT: 9, BACK_RESULT: 10, AMOUNT_STR: 11,
  STATUS: 12, SUBMITTED_BY: 13, SUBMITTED_AT: 14, BET_ID: 15
};
var BETS_COL_COUNT = 16;

// ── COLUMN INDICES: Rounds ──
var R = {
  DATE: 0, ROUND_ID: 1, NAME: 2, FRONT: 3, BACK: 4,
  TOTAL: 5, THRU: 6, COL_H: 7, HANDICAP: 8
};

var FRONT_PAR = 36;
var BACK_PAR  = 36;
var REFRESH_COOLDOWN_SECS = 15; // 15s throttle — prevents double-tap but allows quick retries
var STAKE = 5; // $ per nine — move to Config in v15

// ── SPREADSHEET BINDING ──
// DEV:  set DEV_SHEET_ID to your dev sheet ID and IS_DEV = true
// PROD: set IS_DEV = false (uses the bound spreadsheet — no ID needed)
var IS_DEV       = false;
var DEV_SHEET_ID = '1oAeRzTIAT7qUge9jUd15gWkQB_JvMHHrja9DVWZPpp8';

function SS() {
  return IS_DEV
    ? SpreadsheetApp.openById(DEV_SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

// ── MENU ──
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Wheel Bet')
    .addItem('🔄 Refresh Rounds List', 'refreshRoundsList')
    .addItem('🏌️ Trigger Make Pipeline', 'triggerMake')
    .addItem('🏆 Calculate Results', 'calculateBets')
    .addItem('🛠 Setup Sheets', 'setupSheets')
    .addToUi();
}


// ── COLUMN INDICES: PairingBets (A–W) ──
var PB = {
  DATE: 0, BET_ID: 1, PARENT_ID: 2, TEAMS_A: 3, TEAMS_B: 4, AMOUNT: 5,
  FORMAT: 6, HI_PCT: 7, START_HOLE: 8, SCOPE: 9, STATUS: 10,
  FRONT_RESULT: 11, BACK_RESULT: 12, LONG_RESULT: 13, AMOUNT_STR: 14,
  CREATED_BY: 15, CREATED_AT: 16, AUTO_PRESS: 17, PRESS_THRESHOLD: 18,
  GROUP_ID: 19,    // shared by front/back/long created together; presses inherit
  CURRENT_DIFF: 20, // running stroke diff: positive = Team A winning
  LAST_HOLE: 21,   // last hole computed (1-indexed, 0 = not started)
  FINAL_DIFF: 22,  // final stroke margin when scope completes
  POPS: 23         // JSON pops config: SU={playerName:[holes]} or HI={exclude:[holes]}
};
var PB_COL_COUNT = 24;

// ── ROUTER ──
var ACTIONS = {
  getPlayers:       function(p) { return getPlayers(); },
  getTodaysBets:    function(p) { return getTodaysBets(); },
  submitBet:        function(p) { return submitBetFromApp(p); },
  deleteBet:        function(p) { return deleteBetById(p); },
  addBet:           function(p) { return addBetManually(p); },
  getResults:       function(p) { return getResultsData(); },
  getSkins:         function(p) { return getSkinsData(); },
  getRoundsList:    function(p) { return getRoundsListData(); },
  loadRoundDirect:  function(p) { return loadRoundDirect(p); },
  syncPlayers:      function(p) { return syncPlayers(p); },
  refreshScores:    function(p) { return refreshScores(); },
  calculateResults: function(p) { calculateBets(); return {success: true}; },
  getSeasonLedger:  function(p) { return getSeasonLedger(); },
  getRoundTotals:   function(p) { return getRoundTotals(p); },
  getSeasonByRound: function(p) { return getSeasonByRound(); },
  getSettlement:    function(p) { return getSettlement(p); },
  getAllBets:        function(p) { return getAllBetsForManage(); },
  getRoundStatus:   function(p) { return getRoundStatus(); },
  lockRound:        function(p) { return lockRound(p); },
  unlockRound:      function(p) { return unlockRound(p); },
  finalizeRound:    function(p) { return finalizeRound(p); },
  getLeaderboard:   function(p) { return getLeaderboard(); },
  getHoleByHole:    function(p) { return getHoleByHole(); },
  getBetBanner:     function(p) { return {value: getConfigValue('bet_banner', '')}; },
  openRound:        function(p) { return openRound(p); },
  deleteBetSelf:    function(p) { return deleteBetSelf(p); },
  submitIndie:      function(p) { return submitIndie(p); },
  getIndies:        function(p) { return getIndies(); },
  deleteIndie:      function(p) { return deleteIndie(p); },
  getConfig:        function(p) { return {value: getConfigValue(p.key, p.default || '')}; },
  getAllConfig:      function(p) { return getAllConfig(); },
  validatePin:    function(p) { return validatePin(p); },
  getEverything:  function(p) { return getEverything(); },
  notifyBetsOpen:   function(p) { return notifyBetsOpen(); },
  notifyFrontDone:  function(p) { return notifyFrontDone(); },
  notifyRoundFinal: function(p) { return notifyRoundFinal(); },
  sendTestGroupMe:  function(p) { return sendTestGroupMe(); },
  setConfig:        function(p) { return setConfigValue(p.key, p.value); },
  getScores:        function(p) { return getScores(); },
  getPortalUrl:     function(p) { return getPortalUrl(); },
  changeWheel:      function(p) { return changeWheel(p); },
  removePlayer:     function(p) { return removePlayer(p); },
  deleteWheel:      function(p) { return deleteWheel(p); },
  getWheelTeams:    function(p) { var allData=SS().getSheetByName('DailyBets').getDataRange().getValues(); var today=getBetsForDate(getRoundDate(),allData); return {wheels:findWheelTeams(today)}; },
  refreshRoundsList: function(p) { return refreshRoundsListWeb(); },
  getRoundFormat:   function(p) { return {format: getConfigValue('round_format','nassau')}; },
  getBBHoles:       function(p) { return getBBHoles(); },
  createPairingBet: function(p) { return createPairingBet(p); },
  getPairingBets:   function(p) { return getPairingBets(p); },
  deletePairingBet: function(p) { return deletePairingBet(p); },
  addManualPress:   function(p) { return addManualPress(p); },
  detectPresses:    function(p) { return detectPresses(p); },
  getPairingResults:function(p) { return getPairingResults(p); }
};

function validatePin(p) {
  try {
    var submitted = String(p.pin || '').trim();
    if (!submitted) return { valid: false };
    var storedPin = getConfigValue('admin_pin', '4321');
    return { valid: submitted === String(storedPin).trim() };
  } catch(err) {
    Logger.log('validatePin error: ' + err);
    return { valid: false, error: err.toString() };
  }
}

function getEverything() {
  try {
    var playersData = getPlayers();
    var statusData  = getRoundStatus();
    var format      = getConfigValue('round_format', 'nassau');
    var coursePars  = getConfigValue('course_pars', '');
    var skinsBuyin  = getConfigValue('skins_buyin', '10');
    return {
      players:   playersData.players || [],
      locked:    statusData.locked   || false,
      betsFinal: statusData.betsFinal || false,
      roundDate: statusData.roundDate || '',
      lockedAt:  statusData.lockedAt  || '',
      status:    statusData.status    || '',
      format:    format,
      coursePars: coursePars,
      skinsBuyin: skinsBuyin
    };
  } catch(err) {
    Logger.log('getEverything error: ' + err);
    return { players:[], locked:false, betsFinal:false, roundDate:'', format:'nassau', error:err.toString() };
  }
}

function doGet(e) {
  var action = e.parameter.action;
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('wheelbet')
      .setTitle('Golf Addicts · Wheel Bet')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var result;
  try {
    var handler = ACTIONS[action];
    result = handler ? handler(e.parameter) : {error: 'Unknown action: ' + action};
  } catch (err) {
    Logger.log('doGet error [' + action + ']: ' + err);
    logToSheet(action, 'ERROR: ' + err.toString());
    result = {error: err.toString()};
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ── SETUP ──
function setupSheets() {
  var ss = SS();
  var specs = {
    'SeasonWheel':['date','player','wheel_amount'],
    'SeasonSkins':['date','player','skins_amount'],
    'SeasonIndie':['date','player','indie_amount','note'],
    'RoundStatus':['date','round_name','status','locked_at'],
    'Config':['key','value'],
    'AppLog':['timestamp','action','message'],
    'BBHoles':['date','player','h1','h2','h3','h4','h5','h6','h7','h8','h9','h10','h11','h12','h13','h14','h15','h16','h17','h18'],
    'PairingBets':['date','bet_id','parent_id','teams_a','teams_b','amount','format','hi_pct','start_hole','scope','status','front_result','back_result','long_result','amount_str','created_by','created_at','auto_press','press_threshold','group_id','current_diff','last_hole','final_diff']
  };
  Object.keys(specs).forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); sh.getRange(1,1,1,specs[name].length).setValues([specs[name]]).setFontWeight('bold'); }
  });
  // Seed default config values if Config is empty
  var cfg = ss.getSheetByName('Config');
  if (cfg.getLastRow() < 2) {
    cfg.appendRow(['skins_buyin', '10']);
    cfg.appendRow(['course_pars', '5,4,5,3,4,4,3,4,4,4,4,3,5,3,4,5,3,5']); // Crown Colony default
    cfg.appendRow(['course_name', 'Crown Colony Country Club']);
  }
}

/** Read a config value from Config sheet. Returns string or default. */
function sendGroupMe(message) {
  try {
    var payload = JSON.stringify({bot_id: GROUPME_BOT_ID, text: message});
    UrlFetchApp.fetch('https://api.groupme.com/v3/bots/post', {
      method: 'post',
      contentType: 'application/json',
      payload: payload,
      muteHttpExceptions: true
    });
    logToSheet('sendGroupMe', 'Sent: ' + message.substring(0,60));
  } catch(e) {
    logToSheet('sendGroupMe', 'Error: ' + e.toString());
  }
}

var BETCADDIE_URL = 'https://wagercaddie.golf';

function notifyBetsOpen() {
  var roundDate = getRoundDate();
  if (!roundDate) return {success:false, error:'No active round'};
  
  // Get current bet count
  var betCount = 0;
  try {
    var allData = SS().getSheetByName('DailyBets').getDataRange().getValues();
    var todaysBets = getBetsForDate(roundDate, allData);
    var challBets = todaysBets.filter(function(b) { return b[2] !== 'WHEEL'; });
    betCount = challBets.length;
  } catch(e) {}
  
  var countStr = betCount > 0 ? '\n' + betCount + ' bet' + (betCount !== 1 ? 's' : '') + ' in so far.' : '';
  sendGroupMe('⛳ WagerCaddie is live! Round ' + roundDate + ' is ready.' + countStr + '\nEnter your bets now → ' + BETCADDIE_URL);
  return {success:true};
}

function notifyFrontDone() {
  sendGroupMe('🏌️ Front 9 complete. Check results and enter back 9 bets.\n→ ' + BETCADDIE_URL);
  return {success:true};
}

function notifyRoundFinal() {
  sendGroupMe('🏆 Round is locked! Final results posted. Check your settlement.\n→ ' + BETCADDIE_URL);
  return {success:true};
}

function sendTestGroupMe() {
  sendGroupMe('✅ WagerCaddie is connected!\n→ ' + BETCADDIE_URL);
  return {success:true};
}

function getAllConfig() {
  var keys = ['skins_buyin','round_format','press_threshold','indie_buyin',
               'wheel_amount','chall_amount','hi_pct','auto_press',
               'admin_pin','round_date','round_name',
               'bet_banner','bet_banner_color','gg_portal_prefix','gg_portal_url'];
  var result = {};
  keys.forEach(function(k) {
    try { result[k] = getConfigValue(k, ''); } catch(e) { result[k] = ''; }
  });
  return {success:true, config: result};
}

function getConfigValue(key, defaultVal) {
  var sh = SS().getSheetByName('Config');
  if (!sh) return defaultVal || '';
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return (data[i][1] || '').toString();
  }
  return defaultVal || '';
}

/** Write an error/info line to AppLog sheet (max 500 rows kept) */
function logToSheet(action, message) {
  try {
    var ss = SS();
    var sh = ss.getSheetByName('AppLog');
    if (!sh) return;
    sh.appendRow([nowTimestamp(), action, message]);
    if (sh.getLastRow() > 501) sh.deleteRows(2, sh.getLastRow() - 501);
  } catch(e) {}
}

/** Update or insert a config value */
function setConfigValue(key, value) {
  var ss = SS();
  var sh = ss.getSheetByName('Config');
  if (!sh) { setupSheets(); sh = ss.getSheetByName('Config'); }
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i+1, 2).setValue(value); return {success:true}; }
  }
  sh.appendRow([key, value]);
  return {success:true};
}

// ── HELPERS ──
function generateBetId() { return Utilities.getUuid().substring(0, 8); }
function nowTimestamp() { return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss'); }
function fmtDate(val) { try { return Utilities.formatDate(new Date(val), TIMEZONE, 'yyyy-MM-dd'); } catch(e) { return null; } }

function getRoundDate(roundsData) {
  if (!roundsData) roundsData = SS().getSheetByName('Rounds').getDataRange().getValues();
  if (roundsData.length < 2) return null;
  return Utilities.formatDate(new Date(roundsData[1][R.DATE]), TIMEZONE, 'yyyy-MM-dd');
}

function getBetsForDate(roundDate, allData) {
  var results = [];
  for (var i = 1; i < allData.length; i++) {
    if (!allData[i][B.DATE]) continue;
    var rd = fmtDate(allData[i][B.DATE]);
    if (rd === roundDate) results.push({row: allData[i], idx: i + 1});
  }
  return results;
}

function findWheelTeam(betsForDate) {
  // Legacy single-wheel lookup — returns first wheel found
  var teams = findWheelTeams(betsForDate);
  return teams.length ? teams[0] : null;
}

function findWheelTeams(betsForDate) {
  // Returns all wheel teams for the round
  var teams = [];
  betsForDate.forEach(function(b) {
    var s = (b.row[B.STATUS]||'').toString();
    if (s === 'WHEEL_PENDING' || s === 'WHEEL_FINAL') {
      teams.push({
        p1: b.row[B.W_P1],
        p2: b.row[B.W_P2],
        betId: (b.row[B.BET_ID]||'').toString(),
        status: s
      });
    }
  });
  return teams;
}

// Normalize wheel key for duplicate detection — sorted last names
function wheelKey(p1, p2) {
  var ln = function(n) { return (n||'').split(',')[0].trim().toLowerCase(); };
  return [ln(p1), ln(p2)].sort().join('|');
}

function buildScoreMap(roundsData, roundDate) {
  var scores = {};
  for (var i = 1; i < roundsData.length; i++) {
    if (fmtDate(roundsData[i][R.DATE]) === roundDate) scores[roundsData[i][R.NAME]] = {front: roundsData[i][R.FRONT], back: roundsData[i][R.BACK]};
  }
  return scores;
}

// ── LOAD ROUND DIRECT — replaces Make.com webhook entirely ──
function loadRoundDirect(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var roundId = params.round_id, roundDate = params.round_date;
    var roundFormat = (params.round_format || 'nassau').toLowerCase();
    if (!roundId || !roundDate) return {success:false, error:'Missing round_id or round_date'};
    // Store format in Config — drives calculateBets and getHoleByHole
    setConfigValue('round_format', roundFormat);
    logToSheet('loadRoundDirect', 'Format set to: ' + roundFormat);

    var ss = SS();

    // Clear existing Rounds and DailyBets data
    ['Rounds','DailyBets'].forEach(function(n) {
      var sh = ss.getSheetByName(n);
      if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow()-1);
    });
    var resp = ss.getSheetByName('Form Responses 1');
    if (resp && resp.getLastRow() > 1) resp.deleteRows(2, resp.getLastRow()-1);

    // Fetch player list from GolfGenius Nassau tournament
    var tournsUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments?holes_scope=nassau';
    var tourns = JSON.parse(UrlFetchApp.fetch(tournsUrl, {muteHttpExceptions:true}).getContentText());
    if (!tourns || !tourns.length) return {success:false, error:'No Nassau tournament found for this round'};
    var nassauId = tourns[0].event.id;

    // Fetch scores/aggregates to get player list with handicaps
    var scoresUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments/'+nassauId+'.json';
    var scoresData = JSON.parse(UrlFetchApp.fetch(scoresUrl, {muteHttpExceptions:true}).getContentText());
    var aggregates = scoresData.event.scopes[0].aggregates;
    if (!aggregates || !aggregates.length) return {success:false, error:'No players found in this round'};

    // Write players to Rounds sheet (cols: date, round_id, name, front, back, total, thru, col_h, handicap)
    // Use Utilities.formatDate to create a proper date string, then wrap in new Date()
    // This ensures Google Sheets stores it as a date cell, not text
    var roundsSheet = ss.getSheetByName('Rounds');
    // Parse the date parts directly to avoid timezone issues
    var dateParts = roundDate.split('-');
    var roundDateObj = new Date(parseInt(dateParts[0]), parseInt(dateParts[1])-1, parseInt(dateParts[2]), 12, 0, 0);
    // Try to get handicap indexes from GolfGenius event players endpoint
    var handicapMap = {};
    try {
      var playersUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/players';
      var playersResp = UrlFetchApp.fetch(playersUrl, {muteHttpExceptions:true});
      var playersText = playersResp.getContentText();
      if (playersText.charAt(0) === '[' || playersText.charAt(0) === '{') {
        var playersData = JSON.parse(playersText);
        var playerList = Array.isArray(playersData) ? playersData : (playersData.players || []);
        playerList.forEach(function(p) {
          var name = p.name || (p.player && p.player.name);
          var hdcp = p.playing_handicap || p.course_handicap || p.handicap_index || p.handicap || null;
          if (name && hdcp !== null) handicapMap[name] = Math.round(parseFloat(hdcp));
        });
        logToSheet('loadRoundDirect', 'Players endpoint: '+Object.keys(handicapMap).length+' handicaps loaded');
      } else {
        logToSheet('loadRoundDirect', 'Players endpoint returned non-JSON (HTTP '+playersResp.getResponseCode()+')');
      }
    } catch(e) {
      logToSheet('loadRoundDirect', 'Players endpoint error: '+e.toString());
    }

    var rows = aggregates.map(function(agg) {
      // Get totals for this player
      var g = agg.totals && agg.totals.gross_scores ? agg.totals.gross_scores.total : null;
      var n = agg.totals && agg.totals.net_scores ? agg.totals.net_scores.total : null;
      // Derive handicap
      var handicap = '';
      if (handicapMap[agg.name] !== undefined) {
        handicap = handicapMap[agg.name];
      } else if (g !== null && g !== '' && n !== null && n !== '') {
        handicap = Math.round(g - n);
      } else if (agg.net_scores && agg.gross_scores) {
        // Try summing top-level score arrays
        var gSum = 0, nSum = 0, gCount = 0, nCount = 0;
        if (Array.isArray(agg.gross_scores)) {
          agg.gross_scores.forEach(function(s){if(s!==null&&s!==''&&!isNaN(s)){gSum+=parseInt(s);gCount++;}});
        }
        if (Array.isArray(agg.net_scores)) {
          agg.net_scores.forEach(function(s){if(s!==null&&s!==''&&!isNaN(s)){nSum+=parseInt(s);nCount++;}});
        }
        if (gCount > 0 && nCount > 0) handicap = Math.round(gSum - nSum);
      }
      return [
        roundDateObj, roundId, agg.name,
        agg.totals.net_scores.out !== null ? agg.totals.net_scores.out : '',
        agg.totals.net_scores.in !== null ? agg.totals.net_scores.in : '',
        n !== null ? n : '',
        agg.thru || '',
        '',
        handicap
      ];
    });
    if (rows.length) roundsSheet.getRange(roundsSheet.getLastRow()+1, 1, rows.length, 9).setValues(rows);

    // Write ACTIVE status for this new round (prevents inheriting prior locked status)
    setupSheets();
    var statusSheet = ss.getSheetByName('RoundStatus'), statusData = statusSheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < statusData.length; i++) {
      if (fmtDate(statusData[i][0]) === roundDate) {
        statusSheet.getRange(i+1, 3).setValue('ACTIVE'); found = true; break;
      }
    }
    if (!found) {
      statusSheet.appendRow([roundDateObj, roundDate, 'ACTIVE', '']);
    }

    var cleanDate = Utilities.formatDate(roundDateObj, TIMEZONE, 'yyyy-MM-dd');
    logToSheet('loadRoundDirect', 'Loaded '+rows.length+' players for '+cleanDate);
    return {success:true, players:rows.length, roundDate:cleanDate};
  } catch(e) {
    logToSheet('loadRoundDirect', 'ERROR: '+e.toString());
    return {success:false, error:e.toString()};
  } finally { lock.releaseLock(); }
}

function clearAndTrigger(roundId, roundDate) {

// ── SYNC PLAYERS — add/remove players without resetting bets ──
function syncPlayers(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var ss = SS();
    var roundsSheet = ss.getSheetByName('Rounds');
    var roundsData = roundsSheet.getDataRange().getValues();
    if (roundsData.length < 2) return {success:false, error:'No active round. Load a round first.'};

    var roundDate = getRoundDate(roundsData);
    var roundId = roundsData[1][R.ROUND_ID].toString();

    // Fetch current player list from GolfGenius
    var tournsUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments?holes_scope=nassau';
    var tourns = JSON.parse(UrlFetchApp.fetch(tournsUrl, {muteHttpExceptions:true}).getContentText());
    if (!tourns || !tourns.length) return {success:false, error:'No Nassau tournament found'};
    var nassauId = tourns[0].event.id;

    var scoresUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments/'+nassauId+'.json';
    var scoresData = JSON.parse(UrlFetchApp.fetch(scoresUrl, {muteHttpExceptions:true}).getContentText());
    var aggregates = scoresData.event.scopes[0].aggregates;
    if (!aggregates || !aggregates.length) return {success:false, error:'No players found'};

    // Build set of GolfGenius player names
    var ggNames = {};
    aggregates.forEach(function(agg) { ggNames[agg.name] = agg; });

    // Build set of existing players in Rounds sheet
    var existingNames = {};
    var existingRows = []; // {name, rowIndex}
    for (var i = 1; i < roundsData.length; i++) {
      var rd = fmtDate(roundsData[i][R.DATE]);
      if (rd === roundDate) {
        existingNames[roundsData[i][R.NAME]] = true;
        existingRows.push({name: roundsData[i][R.NAME], idx: i + 1});
      }
    }

    // Try to get handicaps from players endpoint
    var handicapMap = {};
    try {
      var playersUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/players';
      var playersResp = UrlFetchApp.fetch(playersUrl, {muteHttpExceptions:true});
      var playersText = playersResp.getContentText();
      if (playersText.charAt(0) === '[' || playersText.charAt(0) === '{') {
        var playersData = JSON.parse(playersText);
        var playerList = Array.isArray(playersData) ? playersData : (playersData.players || []);
        playerList.forEach(function(p) {
          var name = p.name || (p.player && p.player.name);
          var hdcp = p.playing_handicap || p.course_handicap || p.handicap_index || p.handicap || null;
          if (name && hdcp !== null) handicapMap[name] = Math.round(parseFloat(hdcp));
        });
      }
    } catch(e) {}

    // Find players to ADD (in GolfGenius but not in Rounds)
    var added = [];
    var dateParts = roundDate.split('-');
    var roundDateObj = new Date(parseInt(dateParts[0]), parseInt(dateParts[1])-1, parseInt(dateParts[2]), 12, 0, 0);

    var newRows = [];
    Object.keys(ggNames).forEach(function(name) {
      if (!existingNames[name]) {
        var agg = ggNames[name];
        var g = agg.totals && agg.totals.gross_scores ? agg.totals.gross_scores.total : null;
        var n = agg.totals && agg.totals.net_scores ? agg.totals.net_scores.total : null;
        var handicap = '';
        if (handicapMap[name] !== undefined) {
          handicap = handicapMap[name];
        } else if (g !== null && g !== '' && n !== null && n !== '') {
          handicap = Math.round(g - n);
        }
        newRows.push([
          roundDateObj, roundId, name,
          agg.totals.net_scores.out !== null ? agg.totals.net_scores.out : '',
          agg.totals.net_scores.in !== null ? agg.totals.net_scores.in : '',
          n !== null ? n : '',
          agg.thru || '',
          '',
          handicap
        ]);
        added.push(name);
      }
    });

    if (newRows.length) {
      roundsSheet.getRange(roundsSheet.getLastRow() + 1, 1, newRows.length, 9).setValues(newRows);
    }

    // Find players to REMOVE (in Rounds but not in GolfGenius)
    var removed = [];
    // Delete rows in reverse order to avoid index shifting
    var toDelete = [];
    existingRows.forEach(function(er) {
      if (!ggNames[er.name]) {
        toDelete.push(er.idx);
        removed.push(er.name);
      }
    });
    toDelete.sort(function(a,b) { return b - a; }); // reverse order
    toDelete.forEach(function(rowIdx) { roundsSheet.deleteRow(rowIdx); });

    var msg = [];
    if (added.length) msg.push('Added: ' + added.join(', '));
    if (removed.length) msg.push('Removed: ' + removed.join(', '));
    if (!msg.length) msg.push('No changes — player list is already current');

    logToSheet('syncPlayers', msg.join('. ') + ' (round ' + roundDate + ')');
    return {success:true, added:added, removed:removed, message:msg.join('. '), totalPlayers:Object.keys(ggNames).length};
  } catch(e) {
    logToSheet('syncPlayers', 'ERROR: ' + e.toString());
    return {success:false, error:e.toString()};
  } finally { lock.releaseLock(); }
}


  // Legacy — kept for fallback only. Use loadRoundDirect instead.
  var ss = SS();
  ['Rounds','DailyBets'].forEach(function(n) { var sh = ss.getSheetByName(n); if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow()-1); });
  var resp = ss.getSheetByName('Form Responses 1');
  if (resp && resp.getLastRow() > 1) resp.deleteRows(2, resp.getLastRow()-1);
}

function formatVsPar(score, par) { var d = score - par; return (d === 0 ? 'E' : (d > 0 ? '+' : '') + d) + ' (' + score + ')'; }

// ── ROUND STATUS ──
function getRoundStatus() {
  var roundDate = getRoundDate();
  if (!roundDate) return {status:'no_round', locked:false};
  var sh = SS().getSheetByName('RoundStatus');
  if (!sh) return {status:'active', locked:false, roundDate:roundDate};
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (fmtDate(data[i][0]) === roundDate) {
      var status = (data[i][2]||'').toString();
      return {status:status, locked:status==='LOCKED', betsFinal:status==='BETS_FINAL'||status==='LOCKED', roundDate:roundDate, roundName:data[i][1], lockedAt:(data[i][3]||'').toString()};
    }
  }
  return {status:'active', locked:false, roundDate:roundDate};
}

function lockRound(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var roundDate = getRoundDate();
    if (!roundDate) return {success:false, error:'No active round'};
    var ss = SS(); setupSheets();

    // ── DOUBLE-LOCK GUARD ──
    var sh0 = ss.getSheetByName('RoundStatus'), d0 = sh0.getDataRange().getValues();
    for (var i = 1; i < d0.length; i++) {
      if (fmtDate(d0[i][0]) === roundDate && (d0[i][2]||'').toString() === 'LOCKED') {
        return {success:false, error:'Round is already locked'};
      }
    }

    // Write season totals — fully isolated so API failure doesn't break lock
    var seasonError = null;
    try { writeSeasonTotals(roundDate); } catch(e) { seasonError = e.toString(); Logger.log('writeSeasonTotals error: '+e); }

    // Always write LOCKED status regardless of season total outcome
    var sh = ss.getSheetByName('RoundStatus'), data = sh.getDataRange().getValues();
    var lockedAt = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm'), found = false;
    for (var i = 1; i < data.length; i++) {
      if (fmtDate(data[i][0]) === roundDate) { sh.getRange(i+1,3,1,2).setValues([['LOCKED',lockedAt]]); found = true; break; }
    }
    if (!found) {
      var dp2 = roundDate.split('-');
      var lockDateObj = new Date(parseInt(dp2[0]), parseInt(dp2[1])-1, parseInt(dp2[2]), 12, 0, 0);
      sh.appendRow([lockDateObj, (params&&params.roundName)||roundDate, 'LOCKED', lockedAt]);
    }
    return {success:true, roundDate:roundDate, seasonError:seasonError};
  } finally { lock.releaseLock(); }
}

function unlockRound(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var roundDate = getRoundDate();
    if (!roundDate) return {success:false, error:'No active round'};
    var sh = SS().getSheetByName('RoundStatus');
    if (!sh) return {success:false, error:'RoundStatus sheet not found'};
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (fmtDate(data[i][0]) === roundDate) { sh.getRange(i+1,3).setValue('ACTIVE'); return {success:true}; }
    }
    return {success:false, error:'Round not found in status sheet'};
  } finally { lock.releaseLock(); }
}

// ── FINALIZE BETS — freeze submissions, scores still live ──
function finalizeRound(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var roundDate = getRoundDate();
    if (!roundDate) return {success:false, error:'No active round'};
    var ss = SS(); setupSheets();

    // ── AUTO-DEDUP: remove mirror bets before finalizing ──
    var dedupCount = 0;
    try {
      var betsSheet = ss.getSheetByName('DailyBets');
      var betsData = betsSheet.getDataRange().getValues();
      var today = getBetsForDate(roundDate, betsData);
      // Build map of canonical pair → {idx, submitted_at}
      var pairMap = {}; // key: "A|B" (sorted) → {row1, row2}
      var toDelete = []; // row indices to delete (1-based), in reverse order
      today.forEach(function(b) {
        var s = (b.row[B.STATUS]||'').toString();
        if (s === 'WHEEL_PENDING' || s === 'WHEEL_FINAL') return;
        var p1 = (b.row[B.C_P1]||'').toString().trim();
        var p2 = (b.row[B.C_P2]||'').toString().trim();
        if (!p1 || !p2) return;
        // Canonical key — include wheel team so same challenger pairing against different wheels aren't deduped
        var w1 = (b.row[B.W_P1]||'').toString().trim();
        var w2 = (b.row[B.W_P2]||'').toString().trim();
        var key = w1+'&'+w2+'|'+[p1,p2].sort().join('|');
        if (!pairMap[key]) {
          pairMap[key] = {idx: b.idx, submittedAt: (b.row[B.SUBMITTED_AT]||'').toString()};
        } else {
          // Keep whichever was submitted first, delete the later one
          var existing = pairMap[key];
          var existingTime = existing.submittedAt;
          var newTime = (b.row[B.SUBMITTED_AT]||'').toString();
          if (newTime >= existingTime) {
            // current b is later — delete it
            toDelete.push(b.idx);
          } else {
            // existing is later — delete existing, keep current
            toDelete.push(existing.idx);
            pairMap[key] = {idx: b.idx, submittedAt: newTime};
          }
          dedupCount++;
        }
      });
      // Delete in reverse order to preserve row indices
      toDelete.sort(function(a,b){return b-a;});
      toDelete.forEach(function(rowIdx) { betsSheet.deleteRow(rowIdx); });
      if (dedupCount > 0) logToSheet('finalizeRound', 'Deduped '+dedupCount+' mirror bet(s) for '+roundDate);
    } catch(e) { logToSheet('finalizeRound', 'Dedup error (non-fatal): '+e.toString()); }

    var sh = ss.getSheetByName('RoundStatus'), data = sh.getDataRange().getValues();
    var finalizedAt = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm'), found = false;
    for (var i = 1; i < data.length; i++) {
      if (fmtDate(data[i][0]) === roundDate) {
        sh.getRange(i+1,3,1,2).setValues([['BETS_FINAL', finalizedAt]]); found = true; break;
      }
    }
    if (!found) {
      var dp = roundDate.split('-');
      var rdo = new Date(parseInt(dp[0]),parseInt(dp[1])-1,parseInt(dp[2]),12,0,0);
      sh.appendRow([rdo, roundDate, 'BETS_FINAL', finalizedAt]);
    }
    logToSheet('finalizeRound', 'Bets finalized for '+roundDate);
    return {success:true, roundDate:roundDate, finalizedAt:finalizedAt, dedupCount:dedupCount};
  } finally { lock.releaseLock(); }
}

// ── OPEN ROUND — reverse finalize, back to ACTIVE ──
function openRound(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var roundDate = getRoundDate();
    if (!roundDate) return {success:false, error:'No active round'};
    var sh = SS().getSheetByName('RoundStatus');
    if (!sh) return {success:false, error:'RoundStatus sheet not found'};
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (fmtDate(data[i][0]) === roundDate) {
        sh.getRange(i+1,3).setValue('ACTIVE'); return {success:true};
      }
    }
    return {success:false, error:'Round not found'};
  } finally { lock.releaseLock(); }
}

// ── DELETE BET SELF — player deletes own bet before bets finalized ──
function deleteBetSelf(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var betId = (params.betId||'').toString();
    var playerName = (params.playerName||'').toString();
    if (!betId || !playerName) return {success:false, error:'Missing bet ID or player name'};
    var rs = getRoundStatus();
    if (rs.locked) return {success:false, error:'Round is locked. Contact admin.'};
    if (rs.betsFinal) return {success:false, error:'Bets are finalized. Contact admin to remove.'};
    var sheet = SS().getSheetByName('DailyBets');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][B.BET_ID]||'').toString() !== betId) continue;
      // Allow delete if player is C_P1 or C_P2
      var c1 = (data[i][B.C_P1]||'').toString();
      var c2 = (data[i][B.C_P2]||'').toString();
      if (c1 !== playerName && c2 !== playerName) {
        return {success:false, error:'You can only remove your own bets.'};
      }
      sheet.deleteRow(i+1);
      logToSheet('deleteBetSelf', playerName+' removed bet '+betId);
      return {success:true};
    }
    return {success:false, error:'Bet not found'};
  } finally { lock.releaseLock(); }
}

// ── WRITE SEASON TOTALS ──
function writeSeasonTotals(roundDate) {
  var ss = SS(); setupSheets();
  ['SeasonWheel','SeasonSkins'].forEach(function(shName) {
    var sh = ss.getSheetByName(shName), data = sh.getDataRange().getValues();
    for (var i = data.length-1; i >= 1; i--) { if (fmtDate(data[i][0]) === roundDate) sh.deleteRow(i+1); }
  });
  var betsData = ss.getSheetByName('DailyBets').getDataRange().getValues();
  var wt = {};
  for (var i = 1; i < betsData.length; i++) {
    if (!betsData[i][B.DATE] || (betsData[i][B.STATUS]||'').toString() !== 'FINAL' || !betsData[i][B.AMOUNT_STR]) continue;
    if (fmtDate(betsData[i][B.DATE]) !== roundDate) continue;
    var sum = betsData[i][B.AMOUNT_STR].toString().split(' | ')[2]||'', amt = parseInt(sum.replace(/[^0-9]/g,''))||0;
    var w1=betsData[i][B.W_P1],w2=betsData[i][B.W_P2],c1=betsData[i][B.C_P1],c2=betsData[i][B.C_P2];
    if (!w1||!w2||!c1||!c2||w1==='TBD') continue;
    [w1,w2,c1,c2].forEach(function(p){if(!wt[p])wt[p]=0;});
    if (sum.indexOf('WHEEL WINS')>-1) { wt[w1]+=amt;wt[w2]+=amt;wt[c1]-=amt;wt[c2]-=amt; }
    else if (sum.indexOf('CHALLENGER WINS')>-1) { wt[w1]-=amt;wt[w2]-=amt;wt[c1]+=amt;wt[c2]+=amt; }
  }
  var wSheet = ss.getSheetByName('SeasonWheel'), wRows = [];
  var rdp = roundDate.split('-');
  var roundDateObj = new Date(parseInt(rdp[0]), parseInt(rdp[1])-1, parseInt(rdp[2]), 12, 0, 0);
  Object.keys(wt).forEach(function(p) { wRows.push([roundDateObj,p,wt[p]]); });
  if (wRows.length) wSheet.getRange(wSheet.getLastRow()+1,1,wRows.length,3).setValues(wRows);
  try {
    var sk = getSkinsData(), active = (sk.players||[]).filter(function(p){return p.disposition!=='DNS';}), buyin=parseInt(getConfigValue('skins_buyin','10'));
    var pot=active.length*buyin, winners=active.filter(function(p){return parseInt(p.total)>0;});
    var ts=winners.reduce(function(s,p){return s+parseInt(p.total);},0), ps=ts>0?pot/ts:0;
    var sSheet=ss.getSheetByName('SeasonSkins'), sRows=[];
    active.forEach(function(p){var pay=parseInt(p.total)>0?Math.round(parseInt(p.total)*ps):0;sRows.push([roundDateObj,p.name,pay-buyin]);});
    if(sRows.length) sSheet.getRange(sSheet.getLastRow()+1,1,sRows.length,3).setValues(sRows);
  } catch(e) { Logger.log('Error writing skins totals: '+e); }
}

// ── GET PLAYERS ──
function getPlayers() {
  var rd = SS().getSheetByName('Rounds').getDataRange().getValues();
  if (rd.length < 2) return {players:[], date:null};
  var roundDate = getRoundDate(rd), players = [];
  for (var i = 1; i < rd.length; i++) { if (fmtDate(rd[i][R.DATE])===roundDate) players.push({name:rd[i][R.NAME], handicap:rd[i][R.HANDICAP]||0}); }
  var rs = getRoundStatus();
  return {players:players, date:roundDate, locked:rs.locked, betsFinal:rs.betsFinal||false, roundStatus:rs.status};
}

// ── GET TODAY'S BETS ──
function getTodaysBets() {
  var allData = SS().getSheetByName('DailyBets').getDataRange().getValues();
  var roundDate = getRoundDate(), today = getBetsForDate(roundDate, allData);
  var bets = [], wheelTeams = [];
  today.forEach(function(b) {
    var s = (b.row[B.STATUS]||'').toString(), isW = s==='WHEEL_PENDING'||s==='WHEEL_FINAL';
    if (isW && b.row[B.W_P1]) wheelTeams.push({p1:b.row[B.W_P1], p2:b.row[B.W_P2], betId:(b.row[B.BET_ID]||'').toString()});
    bets.push({betId:(b.row[B.BET_ID]||'').toString(), rowIndex:b.idx, p1:b.row[B.W_P1], p2:b.row[B.W_P2], opp1:b.row[B.C_P1]||'', opp2:b.row[B.C_P2]||'', isWheel:isW, status:s, canDelete:true});
  });
  var rs2 = getRoundStatus();
  // Legacy single wheelTeam for backward compat
  return {bets:bets, wheelTeam:wheelTeams.length?wheelTeams[0]:null, wheelTeams:wheelTeams, locked:rs2.locked, betsFinal:rs2.betsFinal||false};
}

// ── GET ALL BETS FOR MANAGE ──
function getAllBetsForManage() {
  var allData = SS().getSheetByName('DailyBets').getDataRange().getValues();
  var roundDate = getRoundDate(), today = getBetsForDate(roundDate, allData), bets = [];
  today.forEach(function(b) {
    var s = (b.row[B.STATUS]||'').toString();
    if (s==='WHEEL_PENDING'||s==='WHEEL_FINAL') return;
    bets.push({betId:(b.row[B.BET_ID]||'').toString(), rowIndex:b.idx,
      wheel_p1:b.row[B.W_P1]||'', wheel_p2:b.row[B.W_P2]||'',
      chall_p1:b.row[B.C_P1], chall_p2:b.row[B.C_P2], status:s,
      result:b.row[B.AMOUNT_STR]?(b.row[B.AMOUNT_STR].toString().split(' | ')[2]||''):'',
      submitted_by:b.row[B.SUBMITTED_BY]||'', submitted_at:b.row[B.SUBMITTED_AT]||''});
  });
  return {bets:bets};
}

// ── DELETE BET ──
function deleteBetById(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sheet = SS().getSheetByName('DailyBets');
    var betId = (params.betId||'').toString();
    if (betId) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) { if ((data[i][B.BET_ID]||'').toString()===betId) { sheet.deleteRow(i+1); return {success:true}; } }
      return {success:false, error:'Bet not found (ID: '+betId+')'};
    }
    var ri = parseInt(params.rowIndex);
    if (!ri||ri<2) return {success:false, error:'Invalid row'};
    if (ri > sheet.getLastRow()) return {success:false, error:'Row not found'};
    sheet.deleteRow(ri); return {success:true};
  } finally { lock.releaseLock(); }
}

// ── ADD BET MANUALLY ──
function addBetManually(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var c1=params.chall_p1, c2=params.chall_p2, by=params.submitted_by||'Admin';
    if (!c1||!c2) return {success:false, error:'Missing player names'};
    var ss = SS(), betsSheet = ss.getSheetByName('DailyBets');
    var roundDate = getRoundDate();
    if (!roundDate) return {success:false, error:'No active round'};
    var allData = betsSheet.getDataRange().getValues(), today = getBetsForDate(roundDate, allData);
    for (var i = 0; i < today.length; i++) {
      var s = (today[i].row[B.STATUS]||'').toString();
      if (s==='WHEEL_PENDING'||s==='WHEEL_FINAL') continue;
      if (today[i].row[B.C_P1]===c1 && today[i].row[B.C_P2]===c2) return {success:false, error:c1+' & '+c2+' already have a bet today!'};
    }
    var wt = findWheelTeam(today), w1=wt?wt.p1:'', w2=wt?wt.p2:'';
    betsSheet.appendRow([roundDate,w1,w2,c1,c2,'','','','','','','','PENDING',by,nowTimestamp(),generateBetId()]);
    return {success:true};
  } finally { lock.releaseLock(); }
}

// ── SUBMIT BET FROM APP ──
function submitBetFromApp(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var role=params.role, player=params.player, partner=params.partner||'', by=params.submitted_by||player;
    var opponents = params.opponents ? params.opponents.split('|') : [];
    if (!role||!player) return {success:false, error:'Missing required fields'};
    var ss = SS(), betsSheet = ss.getSheetByName('DailyBets');
    var roundDate = getRoundDate();
    if (!roundDate) return {success:false, error:'No active round. Ask admin to load the round first.'};
    var rs = getRoundStatus();
    if (rs.locked) return {success:false, error:'This round is locked. Contact admin.'};
    if (rs.betsFinal) return {success:false, error:'Bets are finalized. Contact admin to make changes.'};
    var now = nowTimestamp(), allData = betsSheet.getDataRange().getValues();
    var today = getBetsForDate(roundDate, allData), existingWheel = findWheelTeam(today);

    if (role === 'wheel') {
      if (!partner) return {success:false, error:'Partner required'};
      // Multi-wheel: check max 4 and no duplicate team (by sorted last names)
      var existingWheels = findWheelTeams(today);
      if (existingWheels.length >= 4) return {success:false, error:'Maximum 4 wheel teams per round.'};
      var newKey = wheelKey(player, partner);
      for (var i = 0; i < existingWheels.length; i++) {
        if (wheelKey(existingWheels[i].p1, existingWheels[i].p2) === newKey) {
          return {success:false, error:existingWheels[i].p1+' & '+existingWheels[i].p2+' are already a Wheel team!'};
        }
      }
      // Also check player isn't already on a wheel team
      for (var i = 0; i < existingWheels.length; i++) {
        if (existingWheels[i].p1 === player || existingWheels[i].p2 === player) {
          return {success:false, error:player+' is already on a Wheel team!'};
        }
        if (existingWheels[i].p1 === partner || existingWheels[i].p2 === partner) {
          return {success:false, error:partner+' is already on a Wheel team!'};
        }
      }
      var newWheelBetId = generateBetId();
      betsSheet.appendRow([roundDate,player,partner,'','','','','','','','','','WHEEL_PENDING',by,now,newWheelBetId]);
      return {success:true, message:player+' & '+partner+' are the Wheel! 🎡', wheelTeam:{p1:player, p2:partner, betId:newWheelBetId}};
    } else {
      // Per-player duplicate check removed — players can submit multiple times
      // Per-pair check below still prevents exact duplicates
      // Only flag duplicates against the SAME wheel team
      var targetWheelId = params.wheel_id || '';
      var dupes = [];
      opponents.forEach(function(opp) {
        for (var i = 0; i < today.length; i++) {
          var s = (today[i].row[B.STATUS]||'').toString();
          if (s==='WHEEL_PENDING'||s==='WHEEL_FINAL') continue;
          if (today[i].row[B.C_P1]!==player || today[i].row[B.C_P2]!==opp) continue;
          // Same challenger pairing — only a duplicate if against the same wheel
          if (targetWheelId) {
            var existingBetId = (today[i].row[B.BET_ID]||'').toString();
            // Find the wheel this existing bet is against
            var existingWheel = null;
            for (var j = 0; j < today.length; j++) {
              var ws = (today[j].row[B.STATUS]||'').toString();
              if (ws!=='WHEEL_PENDING'&&ws!=='WHEEL_FINAL') continue;
              if (today[j].row[B.W_P1]===today[i].row[B.W_P1] && today[j].row[B.W_P2]===today[i].row[B.W_P2]) {
                existingWheel = (today[j].row[B.BET_ID]||'').toString();
                break;
              }
            }
            if (existingWheel && existingWheel !== targetWheelId) continue; // different wheel — not a dupe
          }
          dupes.push(opp);
        }
      });
      if (dupes.length) return {success:false, error:'Duplicate bet detected: '+dupes.join(', ')};
      // Multi-wheel: use wheel_id to target specific wheel, fallback to first wheel
      var targetWheel = null;
      var allWheels = findWheelTeams(today);
      if (params.wheel_id) {
        for (var i = 0; i < allWheels.length; i++) {
          if (allWheels[i].betId === params.wheel_id) { targetWheel = allWheels[i]; break; }
        }
      }
      if (!targetWheel) targetWheel = allWheels.length ? allWheels[0] : null;
      var w1=targetWheel?targetWheel.p1:'TBD', w2=targetWheel?targetWheel.p2:'TBD';
      var newRows = opponents.map(function(opp) { return [roundDate,w1,w2,player,opp,'','','','','','','','PENDING',by,now,generateBetId()]; });
      if (newRows.length) betsSheet.getRange(betsSheet.getLastRow()+1, 1, newRows.length, BETS_COL_COUNT).setValues(newRows);
      var msg = opponents.length+' bet'+(opponents.length>1?'s':'')+' submitted'+(targetWheel?'':' (Wheel TBD)')+'! ⚔️';
      return {success:true, message:msg, wheelTeam:targetWheel, submittedBy:by, bets:opponents.map(function(o){return{player:player,partner:o};})};
    }
  } finally { lock.releaseLock(); }
}

// ── GET RESULTS DATA ──
function getResultsData() {
  var betsData = SS().getSheetByName('DailyBets').getDataRange().getValues();
  var roundDate = getRoundDate(), teams = {}, playerTotals = {};
  var hasPartial = false; // track if any bets are still mid-round
  for (var i = 1; i < betsData.length; i++) {
    var s = (betsData[i][B.STATUS]||'').toString();
    if (s!=='FINAL'&&s!=='PARTIAL') continue;
    if (!betsData[i][B.AMOUNT_STR]) continue;
    var w1=betsData[i][B.W_P1],w2=betsData[i][B.W_P2],c1=betsData[i][B.C_P1],c2=betsData[i][B.C_P2];
    if (!w1||!w2||!c1||!c2) continue;
    var str=betsData[i][B.AMOUNT_STR].toString(), parts=str.split(' | '), sum=parts[2]||'';
    var isPartial = s==='PARTIAL';
    if (isPartial) hasPartial = true;
    // For PARTIAL bets: tally front-nine result as locked-in projection
    // sum is empty for PARTIAL — derive amount from front result only
    var effectiveSum = sum;
    if (isPartial && !sum) {
      var fr = (betsData[i][B.FRONT_RESULT]||'').toString();
      if (fr==='WHEEL') effectiveSum = 'WHEEL WINS $' + STAKE;
      else if (fr==='CHALLENGER') effectiveSum = 'CHALLENGER WINS $' + STAKE;
      else effectiveSum = 'PUSH $0';
    }
    var wk=w1+' & '+w2, ck=c1+' & '+c2;
    if (!teams[wk]) teams[wk]={role:'WHEEL',bets:[],total:0};
    if (!teams[ck]) teams[ck]={role:'CHALLENGER',bets:[],total:0};
    teams[wk].bets.push({opp:ck,detail:str,frontResult:betsData[i][B.FRONT_RESULT]||'',backResult:betsData[i][B.BACK_RESULT]||'',summary:effectiveSum,isPartial:isPartial});
    teams[ck].bets.push({opp:wk,detail:str,frontResult:betsData[i][B.FRONT_RESULT]||'',backResult:betsData[i][B.BACK_RESULT]||'',summary:effectiveSum,isPartial:isPartial});
    var amt=parseInt(effectiveSum.replace(/[^0-9]/g,''))||0;
    if (effectiveSum.indexOf('WHEEL WINS')>-1){teams[wk].total+=amt;teams[ck].total-=amt;}
    else if (effectiveSum.indexOf('CHALLENGER WINS')>-1){teams[wk].total-=amt;teams[ck].total+=amt;}
  }
  Object.keys(teams).forEach(function(tk){tk.split(' & ').forEach(function(p){p=p.trim();if(!playerTotals[p])playerTotals[p]=0;playerTotals[p]+=teams[tk].total;});});
  var indieTotals = getIndieTotalsForDate(roundDate);
  Object.keys(indieTotals).forEach(function(p){if(!playerTotals[p])playerTotals[p]=0;});
  // Merge pairing bet totals into playerTotals
  try {
    var pr = getPairingResults({date:roundDate});
    Object.keys(pr.playerTotals||{}).forEach(function(p) {
      if (!playerTotals[p]) playerTotals[p] = 0;
      // Store separately for 2-down column
    });
    return {teams:teams, playerTotals:playerTotals, indieTotals:indieTotals, roundDate:roundDate, hasPartial:hasPartial, pairingTotals:pr.playerTotals||{}, pairingGroups:pr.betGroups||{}};
  } catch(e) {
    return {teams:teams, playerTotals:playerTotals, indieTotals:indieTotals, roundDate:roundDate, hasPartial:hasPartial, pairingTotals:{}, pairingGroups:{}};
  }
}

// ── INDIE BETS ──
function getIndies() {
  var ss = SS(); setupSheets();
  var data = ss.getSheetByName('SeasonIndie').getDataRange().getValues(), entries = [];
  for (var i = 1; i < data.length; i++) { if (!data[i][0]) continue; var rd=fmtDate(data[i][0]); if(rd) entries.push({date:rd,player:data[i][1],amount:data[i][2],note:data[i][3]||''}); }
  return {entries:entries.slice(-50)};
}

function submitIndie(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var ej=params.entries, note=params.note||'', date=params.date||getRoundDate();
    if (!ej||!date) return {success:false, error:'Missing data'};
    var entries; try{entries=JSON.parse(ej);}catch(e){return{success:false,error:'Invalid entries format'};}
    var total=0; for(var i=0;i<entries.length;i++) total+=(parseInt(entries[i].amount)||0);
    if (total!==0) return {success:false, error:'Entries must net to $0. Current total: '+(total>0?'+':'')+total};
    var ss=SS(); setupSheets();
    var sh=ss.getSheetByName('SeasonIndie'), rows=[];
    entries.forEach(function(e){if(e.player&&e.amount!==0) rows.push([date,e.player,parseInt(e.amount),note]);});
    if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,4).setValues(rows);
    return {success:true, count:entries.length};
  } finally { lock.releaseLock(); }
}

function deleteIndie(params) {
  var date = (params.date || '').toString().trim();
  var note = (params.note || '').toString().trim();
  if (!date) return {success: false, error: 'Missing date'};
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sh = SS().getSheetByName('SeasonIndie');
    if (!sh) return {success: false, error: 'SeasonIndie sheet not found'};
    var lastRow = sh.getLastRow();
    var deleted = 0;
    for (var i = lastRow; i >= 1; i--) {
      var cellVal = sh.getRange(i, 1).getValue();
      var rowNote = sh.getRange(i, 4).getValue().toString().trim();
      // Format date object to YYYY-MM-DD for comparison
      var rowDate = '';
      if (cellVal instanceof Date) {
        var y = cellVal.getFullYear();
        var m = String(cellVal.getMonth() + 1).padStart(2, '0');
        var d = String(cellVal.getDate()).padStart(2, '0');
        rowDate = y + '-' + m + '-' + d;
      } else {
        rowDate = cellVal.toString().trim();
      }
      if (rowDate === date && rowNote === note) {
        sh.deleteRow(i);
        deleted++;
      }
    }
    return {success: true, deleted: deleted};
  } finally { lock.releaseLock(); }
}

function getIndieTotalsForDate(roundDate) {
  if (!roundDate) return {};
  var sh = SS().getSheetByName('SeasonIndie');
  if (!sh) return {};
  var data = sh.getDataRange().getValues(), totals = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (fmtDate(data[i][0])===roundDate) { var p=data[i][1],a=parseInt(data[i][2])||0; if(!totals[p])totals[p]=0; totals[p]+=a; }
  }
  return totals;
}


// ── GET BEST BALL TOURNAMENT ID ──
// Finds the Best Ball tournament in the round's tournament list by name.
function getBestBallId(roundId) {
  var tournsUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments';
  var tourns = JSON.parse(UrlFetchApp.fetch(tournsUrl, {muteHttpExceptions:true}).getContentText());
  for (var i = 0; i < tourns.length; i++) {
    var name = (tourns[i].event.name || '').toLowerCase();
    if (name.indexOf('best ball') > -1 || name.indexOf('bestball') > -1) {
      return tourns[i].event.id;
    }
  }
  return null;
}


// ── GET BB HOLES — returns per-player hole arrays from BBHoles sheet ──
// Used by frontend to render the BB team breakdown grid.
function getBBHoles() {
  var ss = SS();
  var sh = ss.getSheetByName('BBHoles');
  if (!sh) return {available: false, error: 'BBHoles sheet not found'};
  var roundDate = getRoundDate();
  if (!roundDate) return {available: false, error: 'No active round'};
  var data = sh.getDataRange().getValues();
  var players = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    // Robust date comparison: handle Date objects, numbers, and strings
    var cellDate = data[i][0];
    var cellDateStr;
    if (cellDate instanceof Date) {
      cellDateStr = Utilities.formatDate(cellDate, TIMEZONE, 'yyyy-MM-dd');
    } else if (typeof cellDate === 'number') {
      // Spreadsheet serial date — convert properly
      cellDateStr = Utilities.formatDate(new Date(Math.round((cellDate - 25569) * 86400000)), 'UTC', 'yyyy-MM-dd');
    } else {
      cellDateStr = fmtDate(cellDate);
    }
      if (cellDateStr !== roundDate) continue;
    var name = data[i][1];
    var holes = [];
    for (var h = 0; h < 18; h++) {
      var v = data[i][2 + h];
      holes.push((v !== '' && v !== null && v !== undefined) ? Number(v) : null);
    }
    players[name] = holes;
  }
  return {available: Object.keys(players).length > 0, players: players, roundDate: roundDate};
}

// ── BUILD BB SCORE MAP — per-player hole arrays from BBHoles sheet ──
function buildBBScoreMap(roundDate) {
  var ss = SS();
  var sh = ss.getSheetByName('BBHoles');
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  var map = {}; // player -> {holes:[h1..h18], front:sum(h1-9), back:sum(h10-18), thru:N}
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var cellDate2 = data[i][0];
    var cellDateStr2 = cellDate2 instanceof Date
      ? Utilities.formatDate(cellDate2, TIMEZONE, 'yyyy-MM-dd')
      : typeof cellDate2 === 'number'
        ? Utilities.formatDate(new Date(Math.round((cellDate2 - 25569) * 86400000)), 'UTC', 'yyyy-MM-dd')
        : fmtDate(cellDate2);
    if (cellDateStr2 !== roundDate) continue;
    var player = data[i][1];
    var holes = [];
    for (var h = 0; h < 18; h++) {
      var v = data[i][2 + h];
      holes.push((v !== '' && v !== null && v !== undefined) ? Number(v) : null);
    }
    // Count thru (non-null holes)
    var thru = 0;
    for (var h = 0; h < 18; h++) { if (holes[h] !== null) thru++; else break; }
    // Front = holes 0-8, back = holes 9-17 (only if fully complete)
    var frontDone = holes.slice(0,9).every(function(v){return v!==null;});
    var backDone  = holes.slice(9,18).every(function(v){return v!==null;});
    var front = frontDone ? holes.slice(0,9).reduce(function(s,v){return s+v;},0) : null;
    var back  = backDone  ? holes.slice(9,18).reduce(function(s,v){return s+v;},0) : null;
    map[player] = {holes:holes, front:front, back:back, thru:thru};
  }
  return map;
}

// ── COMPUTE BEST BALL TEAM SCORE — min per hole across two players ──
// Returns {front, back, holes, thru} for a 2-player team.
// front/back are null if the nine isn't complete for both players.
function computeBBTeam(p1Map, p2Map) {
  if (!p1Map || !p2Map) return null;
  var bbHoles = [];
  var thru = 0;
  for (var h = 0; h < 18; h++) {
    var v1 = p1Map.holes[h], v2 = p2Map.holes[h];
    if (v1 !== null && v2 !== null) {
      bbHoles.push(Math.min(v1, v2));
      thru++;
    } else {
      bbHoles.push(null);
    }
  }
  var frontDone = bbHoles.slice(0,9).every(function(v){return v!==null;});
  var backDone  = bbHoles.slice(9,18).every(function(v){return v!==null;});
  var front = frontDone ? bbHoles.slice(0,9).reduce(function(s,v){return s+v;},0) : null;
  var back  = backDone  ? bbHoles.slice(9,18).reduce(function(s,v){return s+v;},0) : null;
  return {front:front, back:back, holes:bbHoles, thru:thru};
}

// ── REFRESH SCORES — with throttle ──
function refreshScores() {
  var ss = SS();
  if (getRoundStatus().locked) return {success:false, error:'Round is locked'};
  var cache = CacheService.getScriptCache(), last = cache.get('lastScoreRefresh'), now = new Date().getTime();
  if (last && (now-parseInt(last)) < REFRESH_COOLDOWN_SECS*1000) { calculateBets(); return {success:true, throttled:true, message:'Using cached scores'}; }
  var roundsSheet = ss.getSheetByName('Rounds'), roundsData = roundsSheet.getDataRange().getValues();
  if (roundsData.length < 2) return {success:false, error:'No active round'};
  var roundDate = getRoundDate(roundsData), roundId = roundsData[1][R.ROUND_ID].toString();
  var roundFormat = getConfigValue('round_format', 'nassau');
  var isBestBall = roundFormat === 'bestball';

  // ── Fetch Nassau scores (always — used for Rounds sheet + individual leaderboard) ──
  var tournsUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments?holes_scope=nassau';
  var tourns = JSON.parse(UrlFetchApp.fetch(tournsUrl).getContentText());
  if (!tourns||!tourns.length) return {success:false, error:'No Nassau tournament found'};
  var nassauId = tourns[0].event.id;
  var scoresUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments/'+nassauId+'.json';
  var agg = JSON.parse(UrlFetchApp.fetch(scoresUrl).getContentText()).event.scopes[0].aggregates;
  var scoreMap = {};
  agg.forEach(function(a) {
    var g=a.totals.gross_scores.total, n=a.totals.net_scores.total;
    scoreMap[a.name] = {front:a.totals.net_scores.out, back:a.totals.net_scores.in, total:n, thru:a.thru, handicap:(g!==null&&n!==null)?Math.round(g-n):null};
  });
  // BATCHED score writes (cols D–I per row) — always from Nassau
  for (var i = 1; i < roundsData.length; i++) {
    var name = roundsData[i][R.NAME];
    if (scoreMap[name]) {
      var s = scoreMap[name];
      roundsSheet.getRange(i+1, R.FRONT+1, 1, 6).setValues([[s.front!==null?s.front:'', s.back!==null?s.back:'', s.total!==null?s.total:'', s.thru||'', roundsData[i][R.COL_H]||'', s.handicap!==null?s.handicap:'']]);
    }
  }

  // ── Fetch Best Ball hole arrays and write to BBHoles sheet ──
  if (isBestBall) {
    try {
      var bbId = null;
      for (var i = 0; i < tourns.length; i++) {
        var tname = (tourns[i].event.name || '').toLowerCase();
        if (tname.indexOf('best ball') > -1 || tname.indexOf('bestball') > -1) { bbId = tourns[i].event.id; break; }
      }
      if (bbId) {
        var bbUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments/'+bbId+'.json';
        var bbAgg = JSON.parse(UrlFetchApp.fetch(bbUrl).getContentText()).event.scopes[0].aggregates;
        var bbSheet = ss.getSheetByName('BBHoles');
        if (!bbSheet) { setupSheets(); bbSheet = ss.getSheetByName('BBHoles'); }

        // Delete existing rows for this round date
        var bbData = bbSheet.getDataRange().getValues();
        var rowsToDelete = [];
        for (var i = bbData.length - 1; i >= 1; i--) {
          if (!bbData[i][0]) continue;
          var dc = bbData[i][0];
          var ds = dc instanceof Date ? Utilities.formatDate(dc, TIMEZONE, 'yyyy-MM-dd')
                 : typeof dc === 'number' ? Utilities.formatDate(new Date(Math.round((dc-25569)*86400000)),'UTC','yyyy-MM-dd')
                 : fmtDate(dc);
          if (ds === roundDate) rowsToDelete.push(i + 1);
        }
        rowsToDelete.forEach(function(r) { bbSheet.deleteRow(r); });

        // Write fresh BB hole data
        var rdParts = roundDate.split('-');
        var rdObj = new Date(parseInt(rdParts[0]), parseInt(rdParts[1])-1, parseInt(rdParts[2]), 12, 0, 0);
        var bbRows = [];
        bbAgg.forEach(function(a) {
          if (!a.name) return;
          var holes = Array.isArray(a.net_scores) ? a.net_scores : [];
          // Pad to 18 if short (mid-round)
          while (holes.length < 18) holes.push(null);
          // Replace 0 with null for unplayed holes (GG sometimes returns 0 for unplayed)
          // Actually keep 0 — a net score of 0 is theoretically possible (though rare)
          var row = [rdObj, a.name].concat(holes.slice(0, 18));
          bbRows.push(row);
        });
        if (bbRows.length) {
          bbSheet.getRange(bbSheet.getLastRow()+1, 1, bbRows.length, 20).setValues(bbRows);
        }
        logToSheet('refreshScores', 'BB holes written for ' + bbRows.length + ' players');
      } else {
        logToSheet('refreshScores', 'WARNING: Best Ball format set but no BB tournament found in round');
      }
    } catch(e) {
      logToSheet('refreshScores', 'BB fetch error (non-fatal): ' + e.toString());
    }
  }

  // Reset FINAL/PARTIAL to PENDING
  var betsSheet = ss.getSheetByName('DailyBets'), betsData = betsSheet.getDataRange().getValues();
  for (var i = 1; i < betsData.length; i++) {
    var st = (betsData[i][B.STATUS]||'').toString();
    if (st==='FINAL'||st==='PARTIAL') betsSheet.getRange(i+1, B.STATUS+1).setValue('PENDING');
  }
  cache.put('lastScoreRefresh', now.toString(), REFRESH_COOLDOWN_SECS*2);
  calculateBets();
  try { detectPresses({date:roundDate}); } catch(e) { logToSheet('refreshScores', 'Press detection error (non-fatal): '+e.toString()); }
  return {success:true, updated:Object.keys(scoreMap).length, roundDate:roundDate, format:roundFormat};
}

// ── GET SKINS DATA ──
function getSkinsData() {
  var rd = SS().getSheetByName('Rounds').getDataRange().getValues();
  var roundDate = getRoundDate(rd), roundId = null;
  for (var i = 1; i < rd.length; i++) { if (fmtDate(rd[i][R.DATE])===roundDate) { roundId=rd[i][R.ROUND_ID]; break; } }
  if (!roundId) return {players:[], error:'No active round'};

  // Fetch tournaments list with error handling
  var tournsUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments';
  var tournsResp = UrlFetchApp.fetch(tournsUrl, {muteHttpExceptions:true});
  var tournsText = tournsResp.getContentText();
  if (tournsText.charAt(0) !== '[' && tournsText.charAt(0) !== '{') {
    logToSheet('getSkinsData', 'Tournaments API returned non-JSON: '+tournsText.substring(0,200));
    return {players:[], error:'GolfGenius API error on tournaments list (HTTP '+tournsResp.getResponseCode()+')'};
  }
  var tourns = JSON.parse(tournsText);

  // Log all tournament formats for debugging
  var formats = tourns.map(function(t){return t.event.name+':'+t.event.score_format;}).join(', ');
  logToSheet('getSkinsData', 'Round '+roundId+' tournaments: '+formats);

  // Find skins tournament — check multiple possible format names
  var skinsId = null;
  for (var i = 0; i < tourns.length; i++) {
    var fmt = (tourns[i].event.score_format||'').toLowerCase();
    var name = (tourns[i].event.name||'').toLowerCase();
    if (fmt==='skins' || fmt==='skin' || name.indexOf('skin')>-1) { skinsId=tourns[i].event.id; break; }
  }
  if (!skinsId) return {players:[], error:'No skins tournament found. Formats: '+formats};

  // Fetch skins data with error handling
  var skinsUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments/'+skinsId+'.json';
  var skinsResp = UrlFetchApp.fetch(skinsUrl, {muteHttpExceptions:true});
  var skinsText = skinsResp.getContentText();
  if (skinsText.charAt(0) !== '{') {
    logToSheet('getSkinsData', 'Skins API returned non-JSON: '+skinsText.substring(0,200));
    return {players:[], error:'GolfGenius API error on skins data (HTTP '+skinsResp.getResponseCode()+')'};
  }
  var resp = JSON.parse(skinsText);
  return {players:resp.event.scopes[0].aggregates.map(function(a){return{name:a.name,total:a.total||'0',details:a.details||'',disposition:a.disposition||''};}), roundDate:roundDate};
}

// ── GET ROUNDS LIST ──
function getRoundsListData() {
  var sheet = SS().getSheetByName('Rounds List');
  if (!sheet) return {rounds:[]};
  var data = sheet.getDataRange().getValues(), rounds = [];
  for (var i = 1; i < data.length; i++) { if (data[i][0]) rounds.push({name:data[i][0], date:data[i][1], id:data[i][2].toString()}); }
  return {rounds:rounds};
}

// ── GET SETTLEMENT ──
function getSettlement(params) {
  var buyin = params&&params.buyin?parseInt(params.buyin):parseInt(getConfigValue('skins_buyin','10'));
  var rd = getResultsData(), sk = getSkinsData(), pt = {};
  // Track individual components
  var wheelBy = {}, skinsBy = {}, indieBy = {};
  Object.keys(rd.playerTotals||{}).forEach(function(p){pt[p]=rd.playerTotals[p]; wheelBy[p]=rd.playerTotals[p];});
  var active = (sk.players||[]).filter(function(p){return p.disposition!=='DNS';}), pot=active.length*buyin;
  var winners = active.filter(function(p){return parseInt(p.total)>0;});
  var ts = winners.reduce(function(s,p){return s+parseInt(p.total);},0), ps=ts>0?pot/ts:0;
  active.forEach(function(p){var pay=parseInt(p.total)>0?Math.round(parseInt(p.total)*ps):0;var net=pay-buyin;if(!pt[p.name])pt[p.name]=0;pt[p.name]+=net;skinsBy[p.name]=net;});
  var roundDate = getRoundDate(), it = getIndieTotalsForDate(roundDate);
  Object.keys(it).forEach(function(p){if(!pt[p])pt[p]=0;pt[p]+=it[p];indieBy[p]=it[p];});
  // Add pairing bet totals to settlement
  try {
    var pr = getPairingResults({});
    Object.keys(pr.playerTotals||{}).forEach(function(p) {
      if (!pt[p]) pt[p] = 0;
      pt[p] += pr.playerTotals[p];
    });
  } catch(e) { logToSheet('getSettlement', 'Pairing totals error (non-fatal): '+e.toString()); }

  var owers=[],owees=[];
  Object.keys(pt).forEach(function(p){var a=pt[p];if(a < -0.49)owers.push({name:p,amount:Math.abs(a)});else if(a > 0.49)owees.push({name:p,amount:a});});
  owers.sort(function(a,b){return b.amount-a.amount;}); owees.sort(function(a,b){return b.amount-a.amount;});
  var payments=[],oi=0,ej=0;
  while(oi<owers.length&&ej<owees.length){var pay=Math.min(owers[oi].amount,owees[ej].amount);if(pay>0)payments.push({from:owers[oi].name,to:owees[ej].name,amount:pay});owers[oi].amount-=pay;owees[ej].amount-=pay;if(owers[oi].amount===0)oi++;if(owees[ej].amount===0)ej++;}
  return {payments:payments, playerTotals:pt, wheelBy:wheelBy, skinsBy:skinsBy, indieBy:indieBy, hasPartial:rd.hasPartial||false};
}

// ── GET ROUND TOTALS — for a specific locked round date ──
function getRoundTotals(params) {
  var roundDate = params && params.date ? params.date : null;
  if (!roundDate) return {players: [], error: 'No date provided'};
  var ss = SS();
  var wd = ss.getSheetByName('SeasonWheel').getDataRange().getValues();
  var sd = ss.getSheetByName('SeasonSkins').getDataRange().getValues();
  var id = ss.getSheetByName('SeasonIndie').getDataRange().getValues();
  var wt = {}, st = {}, it = {};
  for (var i = 1; i < wd.length; i++) {
    if (!wd[i][0]) continue;
    if (fmtDate(wd[i][0]) === roundDate) { var p=wd[i][1],a=parseInt(wd[i][2])||0; if(!wt[p])wt[p]=0; wt[p]+=a; }
  }
  for (var i = 1; i < sd.length; i++) {
    if (!sd[i][0]) continue;
    if (fmtDate(sd[i][0]) === roundDate) { var p=sd[i][1],a=parseInt(sd[i][2])||0; if(!st[p])st[p]=0; st[p]+=a; }
  }
  for (var i = 1; i < id.length; i++) {
    if (!id[i][0]) continue;
    if (fmtDate(id[i][0]) === roundDate) { var p=id[i][1],a=parseInt(id[i][2])||0; if(!it[p])it[p]=0; it[p]+=a; }
  }
  var all = {};
  [wt,st,it].forEach(function(m){Object.keys(m).forEach(function(p){all[p]=true;});});
  var players = Object.keys(all).map(function(p) {
    var w=wt[p]||0,s=st[p]||0,ind=it[p]||0;
    return {name:p, wheel:w, skinsNet:s, indie:ind, net:w+s+ind};
  }).sort(function(a,b){return b.net-a.net;});
  return {players: players, roundDate: roundDate};
}

// ── GET SEASON BY ROUND — per-player per-round breakdown ──
function getSeasonByRound() {
  var ss = SS();
  var wd = ss.getSheetByName('SeasonWheel').getDataRange().getValues();
  var sd = ss.getSheetByName('SeasonSkins').getDataRange().getValues();
  var id = ss.getSheetByName('SeasonIndie').getDataRange().getValues();
  // byRound[date][player] = {wheel, skins, indie}
  var byRound = {};
  function ensure(d,p){if(!byRound[d])byRound[d]={};if(!byRound[d][p])byRound[d][p]={wheel:0,skins:0,indie:0};}
  for (var i = 1; i < wd.length; i++) {
    if (!wd[i][0]) continue;
    var d=fmtDate(wd[i][0]),p=wd[i][1],a=parseInt(wd[i][2])||0;
    if(d&&p){ensure(d,p);byRound[d][p].wheel+=a;}
  }
  for (var i = 1; i < sd.length; i++) {
    if (!sd[i][0]) continue;
    var d=fmtDate(sd[i][0]),p=sd[i][1],a=parseInt(sd[i][2])||0;
    if(d&&p){ensure(d,p);byRound[d][p].skins+=a;}
  }
  for (var i = 1; i < id.length; i++) {
    if (!id[i][0]) continue;
    var d=fmtDate(id[i][0]),p=id[i][1],a=parseInt(id[i][2])||0;
    if(d&&p){ensure(d,p);byRound[d][p].indie+=a;}
  }
  return {byRound: byRound};
}

// ── GET SEASON LEDGER ──
function getSeasonLedger() {
  var ss = SS(); setupSheets();
  var wd=ss.getSheetByName('SeasonWheel').getDataRange().getValues(), sd=ss.getSheetByName('SeasonSkins').getDataRange().getValues(), id=ss.getSheetByName('SeasonIndie').getDataRange().getValues();
  var wt={},st={},it={},rds={};
  for(var i=1;i<wd.length;i++){if(!wd[i][0])continue;var r=fmtDate(wd[i][0]);if(r){rds[r]=true;var p=wd[i][1],a=parseInt(wd[i][2])||0;if(!wt[p])wt[p]=0;wt[p]+=a;}}
  for(var i=1;i<sd.length;i++){if(!sd[i][0])continue;var p=sd[i][1],a=parseInt(sd[i][2])||0;if(!st[p])st[p]=0;st[p]+=a;}
  for(var i=1;i<id.length;i++){if(!id[i][0])continue;var p=id[i][1],a=parseInt(id[i][2])||0;if(!it[p])it[p]=0;it[p]+=a;}
  var all={};[wt,st,it].forEach(function(m){Object.keys(m).forEach(function(p){all[p]=true;});});
  var players=Object.keys(all).map(function(p){var w=wt[p]||0,s=st[p]||0,ind=it[p]||0;return{name:p,wheel:w,skins:s,indie:ind,total:w+s+ind};}).sort(function(a,b){return b.total-a.total;});
  return {players:players, rounds:Object.keys(rds).length, roundDates:Object.keys(rds)};
}

// ── CALCULATE BETS — Nassau and Best Ball ──
function calculateBets() {
  var ss = SS();
  var roundFormat = getConfigValue('round_format', 'nassau');
  var isBestBall = roundFormat === 'bestball';
  var roundsSheet=ss.getSheetByName('Rounds'), betsSheet=ss.getSheetByName('DailyBets'), ledgerSheet=ss.getSheetByName('Ledger');
  var roundsData = roundsSheet.getDataRange().getValues(), roundDate = getRoundDate(roundsData);
  if (!roundDate) { try{SpreadsheetApp.getUi().alert('No round data.');}catch(e){} return; }

  // Score sources differ by format
  var nassauScores = buildScoreMap(roundsData, roundDate);  // always available
  var bbScores = isBestBall ? buildBBScoreMap(roundDate) : null;

  var betsData = betsSheet.getDataRange().getValues(), ledgerRows = [];
  for (var i = 1; i < betsData.length; i++) {
    var rowStatus = (betsData[i][B.STATUS]||'').toString();
    if (rowStatus!=='PENDING'&&rowStatus!=='PARTIAL') continue;
    var w1=betsData[i][B.W_P1],w2=betsData[i][B.W_P2],c1=betsData[i][B.C_P1],c2=betsData[i][B.C_P2];
    if (!w1||!w2||!c1||!c2) continue;

    var wFt, wBt, cFt, cBt, bc;

    if (isBestBall) {
      // ── BEST BALL SCORING ──
      // Need BB hole data for all 4 players
      if (!bbScores[w1]||!bbScores[w2]||!bbScores[c1]||!bbScores[c2]) continue;
      var wTeam = computeBBTeam(bbScores[w1], bbScores[w2]);
      var cTeam = computeBBTeam(bbScores[c1], bbScores[c2]);
      if (!wTeam || !cTeam) continue;
      // Front nine required; back nine optional (PARTIAL if mid-round)
      if (wTeam.front === null || cTeam.front === null) continue;
      wFt = wTeam.front; cFt = cTeam.front;
      bc = wTeam.back !== null && cTeam.back !== null;
      wBt = bc ? wTeam.back : null;
      cBt = bc ? cTeam.back : null;
    } else {
      // ── NASSAU SCORING ──
      if (!nassauScores[w1]||!nassauScores[w2]||!nassauScores[c1]||!nassauScores[c2]) continue;
      var wF=nassauScores[w1].front,wF2=nassauScores[w2].front,cF=nassauScores[c1].front,cF2=nassauScores[c2].front;
      var wB=nassauScores[w1].back,wB2=nassauScores[w2].back,cB=nassauScores[c1].back,cB2=nassauScores[c2].back;
      if ([wF,wF2,cF,cF2].some(function(s){return s===null||s===''||s===undefined||isNaN(Number(s));})) continue;
      bc=[wB,wB2,cB,cB2].every(function(s){return s!==null&&s!==''&&s!==undefined&&!isNaN(Number(s));});
      wFt=Number(wF)+Number(wF2); cFt=Number(cF)+Number(cF2);
      wBt=bc?Number(wB)+Number(wB2):null; cBt=bc?Number(cB)+Number(cB2):null;
    }

    var fr=wFt<cFt?'WHEEL':wFt>cFt?'CHALLENGER':'PUSH';
    var br=!bc?'PUSH':wBt<cBt?'WHEEL':wBt>cBt?'CHALLENGER':'PUSH';
    var net=(fr==='WHEEL'?STAKE:0)+(br==='WHEEL'?STAKE:0)-((fr==='CHALLENGER'?STAKE:0)+(br==='CHALLENGER'?STAKE:0));
    var formatTag = isBestBall ? 'BB ' : '';
    var fStr='Front: '+formatTag+'WHEEL '+formatVsPar(wFt,FRONT_PAR)+' vs CHALL '+formatVsPar(cFt,FRONT_PAR)+' → '+(fr==='PUSH'?'PUSH':fr+' +$'+STAKE);
    var bStr=bc?'Back: '+formatTag+'WHEEL '+formatVsPar(wBt,BACK_PAR)+' vs CHALL '+formatVsPar(cBt,BACK_PAR)+' → '+(br==='PUSH'?'PUSH':br+' +$'+STAKE):'Back: In progress';
    var sumStr=net>0?'WHEEL WINS $'+net:net<0?'CHALLENGER WINS $'+Math.abs(net):'PUSH $0';
    var amtStr=fStr+' | '+bStr+' | '+sumStr, newStatus=bc?'FINAL':'PARTIAL';
    betsSheet.getRange(i+1, B.W_FRONT+1, 1, 8).setValues([[wFt,bc?wBt:'',cFt,bc?cBt:'',fr,br,amtStr,newStatus]]);
    if (newStatus==='FINAL') ledgerRows.push([roundDate,w1+' & '+w2+' vs '+c1+' & '+c2,fr,br,amtStr]);
  }
  if (ledgerRows.length) ledgerSheet.getRange(ledgerSheet.getLastRow()+1,1,ledgerRows.length,5).setValues(ledgerRows);
  try{SpreadsheetApp.getUi().alert('Results calculated!');}catch(e){}
}

// ── TRIGGER MAKE ──
function refreshRoundsListWeb() {
  try {
    var rounds = JSON.parse(UrlFetchApp.fetch(
      'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds',
      {muteHttpExceptions:true}
    ).getContentText());
    if (!Array.isArray(rounds)) return {success:false, error:'Unexpected GG response'};
    var ss = SS();
    var sheet = ss.getSheetByName('Rounds List') || ss.insertSheet('Rounds List');
    sheet.clearContents();
    sheet.getRange(1,1,1,3).setValues([['Round Name','Date','Round ID']]).setFontWeight('bold');
    var rows = [];
    rounds.forEach(function(r) {
      var ro = r.round;
      if (ro && ro.date && ro.id) rows.push([ro.name+' - '+ro.date, ro.date, "'"+ro.id]);
    });
    if (rows.length) sheet.getRange(2,1,rows.length,3).setValues(rows);
    logToSheet('refreshRoundsList', 'Refreshed — '+rows.length+' rounds found');
    return {success:true, count:rows.length};
  } catch(e) {
    return {success:false, error:e.toString()};
  }
}

function refreshRoundsList() {
  var rounds = JSON.parse(UrlFetchApp.fetch('https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds').getContentText());
  var ss=SS(), sheet=ss.getSheetByName('Rounds List')||ss.insertSheet('Rounds List');
  sheet.clearContents(); sheet.getRange(1,1,1,3).setValues([['Round Name','Date','Round ID']]).setFontWeight('bold');
  var rows=[]; rounds.forEach(function(r){var ro=r.round;if(ro.date&&ro.id)rows.push([ro.name+' - '+ro.date,ro.date,"'"+ro.id]);});
  if(rows.length) sheet.getRange(2,1,rows.length,3).setValues(rows);
  var trig=ss.getSheetByName('TriggerRound');
  if(trig) trig.getRange('A2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(rows.map(function(r){return r[0];}),true).build());
  SpreadsheetApp.getUi().alert('Rounds list refreshed! '+rows.length+' rounds found.');
}

function triggerMake() {
  var ss=SS(), sel=ss.getSheetByName('TriggerRound').getRange('A2').getValue();
  if(!sel){SpreadsheetApp.getUi().alert('Select a round first.');return;}
  var rlData=ss.getSheetByName('Rounds List').getDataRange().getValues(), roundId=null, roundDate=null;
  for(var i=1;i<rlData.length;i++){if(rlData[i][0]==sel){roundDate=Utilities.formatDate(new Date(rlData[i][1]),TIMEZONE,'yyyy-MM-dd');roundId=rlData[i][2].toString();break;}}
  if(!roundId){SpreadsheetApp.getUi().alert('Round not found.');return;}
  clearAndTrigger(roundId, roundDate);
  SpreadsheetApp.getUi().alert('Triggered for '+sel+'!\nCheck Rounds tab in ~45 seconds.');
}

// ── GET LEADERBOARD — stroke-play leaderboard from Nassau tournament ──
function getLeaderboard() {
  var rd = SS().getSheetByName('Rounds').getDataRange().getValues();
  var roundDate = getRoundDate(rd), roundId = null;
  for (var i = 1; i < rd.length; i++) { if (fmtDate(rd[i][R.DATE])===roundDate) { roundId=rd[i][R.ROUND_ID].toString(); break; } }
  if (!roundId) return {players:[], error:'No active round'};
  var roundFormat = getConfigValue('round_format', 'nassau');
  var isBestBall = roundFormat === 'bestball';
  try {
    var tournsUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments?holes_scope=nassau';
    var tourns = JSON.parse(UrlFetchApp.fetch(tournsUrl, {muteHttpExceptions:true}).getContentText());
    if (!tourns||!tourns.length) return {players:[], error:'No tournaments found'};
    // For Best Ball rounds use the BB tournament (100% HI net scores)
    var targetId = tourns[0].event.id; // default Nassau
    if (isBestBall) {
      for (var i = 0; i < tourns.length; i++) {
        var tn = (tourns[i].event.name || '').toLowerCase();
        if (tn.indexOf('best ball') > -1 || tn.indexOf('bestball') > -1) { targetId = tourns[i].event.id; break; }
      }
    }
    var scoresUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments/'+targetId+'.json';
    var scoresData = JSON.parse(UrlFetchApp.fetch(scoresUrl, {muteHttpExceptions:true}).getContentText());
    var agg = scoresData.event.scopes[0].aggregates;
    var players = agg.map(function(a) {
      var gross = a.totals && a.totals.gross_scores ? a.totals.gross_scores.total : null;
      var net   = a.totals && a.totals.net_scores   ? a.totals.net_scores.total   : null;
      return { name:a.name, gross:gross!==null?gross:'', net:net!==null?net:'', thru:a.thru||'' };
    });
    // Sort by net ascending (lower is better), not-started last
    players.sort(function(a,b) {
      if (a.net===''&&b.net==='') return 0;
      if (a.net==='') return 1; if (b.net==='') return -1;
      return parseInt(a.net)-parseInt(b.net);
    });
    // Assign positions with ties
    for (var i = 0; i < players.length; i++) {
      if (i>0&&players[i].net!==''&&players[i].net===players[i-1].net) players[i].position=players[i-1].position;
      else players[i].position = i+1;
    }
    return {players:players, roundDate:roundDate};
  } catch(e) {
    logToSheet('getLeaderboard','ERROR: '+e.toString());
    return {players:[], error:e.toString()};
  }
}

// ── GET HOLE BY HOLE — per-player net/gross score arrays ──
function getHoleByHole() {
  var rd = SS().getSheetByName('Rounds').getDataRange().getValues();
  var roundDate = getRoundDate(rd), roundId = null;
  for (var i = 1; i < rd.length; i++) { if (fmtDate(rd[i][R.DATE])===roundDate) { roundId=rd[i][R.ROUND_ID].toString(); break; } }
  if (!roundId) return {available:false, error:'No active round'};
  var roundFormat = getConfigValue('round_format', 'nassau');
  var isBestBall = roundFormat === 'bestball';
  try {
    var tournsUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments?holes_scope=nassau';
    var tourns = JSON.parse(UrlFetchApp.fetch(tournsUrl, {muteHttpExceptions:true}).getContentText());
    if (!tourns||!tourns.length) return {available:false, error:'No tournaments found'};
    // For Best Ball, use the BB tournament for hole-by-hole data (100% HI)
    // For Nassau, use tourns[0] (Nassau tournament, 80% HI)
    var targetId = null;
    if (isBestBall) {
      for (var i = 0; i < tourns.length; i++) {
        var tn = (tourns[i].event.name || '').toLowerCase();
        if (tn.indexOf('best ball') > -1 || tn.indexOf('bestball') > -1) { targetId = tourns[i].event.id; break; }
      }
      if (!targetId) { logToSheet('getHoleByHole', 'BB format but no BB tournament found, falling back to Nassau'); targetId = tourns[0].event.id; }
    } else {
      targetId = tourns[0].event.id;
    }
    var nassauId = targetId;
    var scoresUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments/'+nassauId+'.json';
    var scoresData = JSON.parse(UrlFetchApp.fetch(scoresUrl, {muteHttpExceptions:true}).getContentText());
    var agg = scoresData.event.scopes[0].aggregates;
    // Detect hole array location from first player
    var sample = agg[0] || {};
    var getNetArr = function(a) {
      if (Array.isArray(a.net_scores) && a.net_scores.length > 0) return a.net_scores;
      if (a.totals && a.totals.net_scores && Array.isArray(a.totals.net_scores.holes)) return a.totals.net_scores.holes;
      return null;
    };
    var getGrossArr = function(a) {
      if (Array.isArray(a.gross_scores) && a.gross_scores.length > 0) return a.gross_scores;
      if (a.totals && a.totals.gross_scores && Array.isArray(a.totals.gross_scores.holes)) return a.totals.gross_scores.holes;
      return null;
    };
    if (!getNetArr(sample)) return {available:false, error:'Hole-by-hole data not available in payload yet'};
    // Fetch Best Ball tournament to get accurate course handicaps (100% HI)
    var bbHcpMap = {};
    try {
      var bbId = null;
      for (var ti = 0; ti < tourns.length; ti++) {
        var tn = (tourns[ti].event.name||'').toLowerCase();
        if (tn.indexOf('best ball') > -1 || tn.indexOf('bestball') > -1) {
          bbId = tourns[ti].event.id; break;
        }
      }
      if (bbId && bbId !== nassauId) {
        var bbUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments/'+bbId+'.json';
        var bbData = JSON.parse(UrlFetchApp.fetch(bbUrl, {muteHttpExceptions:true}).getContentText());
        var bbAgg = bbData.event.scopes[0].aggregates;
        bbAgg.forEach(function(a) {
          var bbNet = getNetArr(a), bbGross = getGrossArr(a);
          if (!bbNet || !bbGross) return;
          var gSum = bbGross.reduce(function(s,v){return s+(v&&!isNaN(parseInt(v))?parseInt(v):0);},0);
          var nSum = bbNet.reduce(function(s,v){return s+(v&&!isNaN(parseInt(v))?parseInt(v):0);},0);
          if (gSum > 0 && nSum > 0) bbHcpMap[a.name] = Math.round(gSum - nSum);
        });
        logToSheet('getHoleByHole', 'Loaded course handicaps from Best Ball for '+Object.keys(bbHcpMap).length+' players');
      }
    } catch(e) {
      logToSheet('getHoleByHole', 'Could not load Best Ball HCP: '+e.toString());
    }

    var players = {};
    agg.forEach(function(a) {
      var net = getNetArr(a), gross = getGrossArr(a);
      if (!net) return;
      // Use Best Ball tournament for course handicap (most accurate)
      // Fallback to gross-net difference from current tournament
      var hdcp = bbHcpMap[a.name] || null;
      if (hdcp === null) {
        hdcp = a.playing_handicap || a.course_handicap || a.handicap_index || null;
        if (hdcp === null && gross && gross.length > 0 && net && net.length > 0) {
          var gSum = gross.reduce(function(s,v){return s+(v?parseInt(v):0);},0);
          var nSum = net.reduce(function(s,v){return s+(v?parseInt(v):0);},0);
          if (gSum > 0 && nSum > 0) hdcp = Math.round(gSum - nSum);
        }
      }
      players[a.name] = {net:net, gross:gross||[], thru:a.thru||0, handicap:hdcp||0};
    });
    logToSheet('getHoleByHole','Returned hole data for '+Object.keys(players).length+' players');
    return {available:true, players:players, roundDate:roundDate};
  } catch(e) {
    logToSheet('getHoleByHole','ERROR: '+e.toString());
    return {available:false, error:e.toString()};
  }
}



// ── CHANGE WHEEL — replace wheel team, update all pending bet rows ──
// Admin only. Only valid before BETS_FINAL.
// Requires wheel_id (betId of the WHEEL_PENDING row) to target specific wheel in multi-wheel rounds.
function changeWheel(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var p1 = (params.p1 || '').toString().trim();
    var p2 = (params.p2 || '').toString().trim();
    var wheelId = (params.wheel_id || '').toString().trim();
    if (!p1 || !p2) return {success: false, error: 'Both players required'};
    if (p1 === p2) return {success: false, error: 'Players must be different'};

    var rs = getRoundStatus();
    if (rs.locked) return {success: false, error: 'Round is locked'};
    if (rs.betsFinal) return {success: false, error: 'Bets are finalized — cannot change wheel'};

    var ss = SS();
    var betsSheet = ss.getSheetByName('DailyBets');
    var data = betsSheet.getDataRange().getValues();
    var roundDate = getRoundDate();
    if (!roundDate) return {success: false, error: 'No active round'};

    // Find the target wheel team's current players (to update only its challenger rows)
    var oldW1 = '', oldW2 = '';
    for (var i = 1; i < data.length; i++) {
      if (!data[i][B.DATE]) continue;
      if (fmtDate(data[i][B.DATE]) !== roundDate) continue;
      var s0 = (data[i][B.STATUS]||'').toString();
      if (s0 !== 'WHEEL_PENDING' && s0 !== 'WHEEL_FINAL') continue;
      // If wheel_id provided, match by betId; otherwise use first wheel found
      var rowBetId = (data[i][B.BET_ID]||'').toString();
      if (!wheelId || rowBetId === wheelId) {
        oldW1 = (data[i][B.W_P1]||'').toString();
        oldW2 = (data[i][B.W_P2]||'').toString();
        break;
      }
    }
    if (!oldW1) return {success: false, error: 'Wheel team not found'};

    var updated = 0;
    for (var i = 1; i < data.length; i++) {
      if (!data[i][B.DATE]) continue;
      if (fmtDate(data[i][B.DATE]) !== roundDate) continue;
      var status = (data[i][B.STATUS] || '').toString();
      var rowW1 = (data[i][B.W_P1]||'').toString();
      var rowW2 = (data[i][B.W_P2]||'').toString();
      // Only update rows belonging to this specific wheel team
      if (rowW1 !== oldW1 || rowW2 !== oldW2) continue;
      if (status === 'WHEEL_PENDING' || status === 'WHEEL_FINAL' ||
          status === 'PENDING' || status === 'PARTIAL') {
        betsSheet.getRange(i + 1, B.W_P1 + 1, 1, 2).setValues([[p1, p2]]);
        updated++;
      }
    }

    logToSheet('changeWheel', 'Wheel changed from '+oldW1+' & '+oldW2+' to ' + p1 + ' & ' + p2 + ' — ' + updated + ' rows updated');
    return {success: true, p1: p1, p2: p2, updated: updated, oldW1: oldW1, oldW2: oldW2};
  } finally { lock.releaseLock(); }
}

// ── GET SCORES — individual player net scores from the Rounds sheet ──
// Returns each player's front, back, total and thru for the current round.
// These are stored locally when loadRoundDirect runs and updated on every
// refreshScores call, so this works even for locked rounds.
function getScores() {
  var rd = SS().getSheetByName('Rounds').getDataRange().getValues();
  var roundDate = getRoundDate(rd);
  if (!roundDate) return {players: [], error: 'No active round'};
  var players = [];
  for (var i = 1; i < rd.length; i++) {
    if (fmtDate(rd[i][R.DATE]) !== roundDate) continue;
    var front = rd[i][R.FRONT];
    var back  = rd[i][R.BACK];
    var total = rd[i][R.TOTAL];
    var thru  = rd[i][R.THRU];
    players.push({
      name:     rd[i][R.NAME],
      front:    (front !== '' && front !== null && front !== undefined) ? Number(front) : null,
      back:     (back  !== '' && back  !== null && back  !== undefined) ? Number(back)  : null,
      total:    (total !== '' && total !== null && total !== undefined) ? Number(total) : null,
      thru:     (thru  !== '' && thru  !== null && thru  !== undefined) ? thru  : null,
      handicap: rd[i][R.HANDICAP] || 0
    });
  }
  return {players: players, roundDate: roundDate};
}

// ── GET PORTAL URL — auto-constructs GolfGenius TV/portal URL ──
// Requires gg_portal_prefix in Config sheet (set once).
// The TV display URL is always current and embeds cleanly without login
// once the portal page is set to Public in GolfGenius.
// Example: if prefix = "golfaddicts", URL = golfaddicts.golfgenius.com/tv
function getPortalUrl() {
  var prefix = getConfigValue('gg_portal_prefix', '');
  if (!prefix) return {url: '', error: 'gg_portal_prefix not set in Config sheet'};
  var rd = SS().getSheetByName('Rounds').getDataRange().getValues();
  if (rd.length < 2) return {url: '', error: 'No active round'};
  var roundId = (rd[1][R.ROUND_ID] || '').toString();
  var tvUrl     = 'https://' + prefix + '.golfgenius.com/tv';
  var portalUrl = 'https://' + prefix + '.golfgenius.com/rounds/' + roundId + '/results';
  return {url: tvUrl, tvUrl: tvUrl, portalUrl: portalUrl, prefix: prefix, roundId: roundId};
}



// ── REMOVE PLAYER — drop a player mid-round ──
// Removes player from Rounds sheet, deletes their challenger bets from DailyBets,
// clears their BBHoles rows. Skins are left intact (they may have won skins already).
// If player is on the Wheel team, returns a warning instead of removing.

// ── DELETE WHEEL TEAM — cascade deletes wheel row + all its challenger bets ──
function deleteWheel(params) {
  var betId = (params.betId||'').toString();
  if (!betId) return {success:false, error:'No wheel betId provided'};
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var ss = SS(), sheet = ss.getSheetByName('DailyBets');
    var data = sheet.getDataRange().getValues();
    var roundDate = getRoundDate();
    // Find wheel team names from the wheel row
    var w1 = '', w2 = '';
    for (var i = 1; i < data.length; i++) {
      if ((data[i][B.BET_ID]||'').toString() === betId) {
        w1 = (data[i][B.W_P1]||'').toString();
        w2 = (data[i][B.W_P2]||'').toString();
        break;
      }
    }
    if (!w1) return {success:false, error:'Wheel team not found'};
    // Delete wheel row + all challenger bets for this wheel team
    var toDelete = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][B.DATE]) continue;
      if (fmtDate(data[i][B.DATE]) !== roundDate) continue;
      var rowW1 = (data[i][B.W_P1]||'').toString();
      var rowW2 = (data[i][B.W_P2]||'').toString();
      if (rowW1 === w1 && rowW2 === w2) toDelete.push(i + 1);
    }
    toDelete.sort(function(a,b){return b-a;});
    toDelete.forEach(function(r) { sheet.deleteRow(r); });
    logToSheet('deleteWheel', 'Deleted wheel '+w1+' & '+w2+' and '+toDelete.length+' associated rows');
    return {success:true, deleted:toDelete.length, w1:w1, w2:w2};
  } finally { lock.releaseLock(); }
}

function removePlayer(params) {
  var name = (params.name || '').toString().trim();
  if (!name) return {success:false, error:'No player name provided'};
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var ss = SS();
    var roundDate = getRoundDate();
    if (!roundDate) return {success:false, error:'No active round'};

    // ── Check if player is on the Wheel team ──
    var betsSheet = ss.getSheetByName('DailyBets');
    var betsData = betsSheet.getDataRange().getValues();
    for (var i = 1; i < betsData.length; i++) {
      if (fmtDate(betsData[i][B.DATE]) !== roundDate) continue;
      var s = (betsData[i][B.STATUS]||'').toString();
      if (s !== 'WHEEL_PENDING' && s !== 'WHEEL_FINAL') continue;
      if (betsData[i][B.W_P1] === name || betsData[i][B.W_P2] === name) {
        return {success:false, error:'Cannot remove wheel player. Delete their wheel team first from Admin, then remove them.'};
      }
    }

    // ── Remove from Rounds sheet ──
    var roundsSheet = ss.getSheetByName('Rounds');
    var roundsData = roundsSheet.getDataRange().getValues();
    var roundsDeleted = 0;
    for (var i = roundsData.length - 1; i >= 1; i--) {
      if (fmtDate(roundsData[i][R.DATE]) === roundDate && roundsData[i][R.NAME] === name) {
        roundsSheet.deleteRow(i + 1);
        roundsDeleted++;
      }
    }

    // ── Delete challenger bets ──
    betsData = betsSheet.getDataRange().getValues(); // re-read after potential changes
    var betsDeleted = 0;
    for (var i = betsData.length - 1; i >= 1; i--) {
      if (fmtDate(betsData[i][B.DATE]) !== roundDate) continue;
      var s2 = (betsData[i][B.STATUS]||'').toString();
      if (s2 === 'WHEEL_PENDING' || s2 === 'WHEEL_FINAL') continue;
      if (betsData[i][B.C_P1] === name || betsData[i][B.C_P2] === name) {
        betsSheet.deleteRow(i + 1);
        betsDeleted++;
      }
    }

    // ── Remove BBHoles rows ──
    var bbDeleted = 0;
    var bbSheet = ss.getSheetByName('BBHoles');
    if (bbSheet) {
      var bbData = bbSheet.getDataRange().getValues();
      for (var i = bbData.length - 1; i >= 1; i--) {
        if (!bbData[i][0]) continue;
        var dc = bbData[i][0];
        var ds = dc instanceof Date ? Utilities.formatDate(dc, TIMEZONE, 'yyyy-MM-dd')
               : typeof dc === 'number' ? Utilities.formatDate(new Date(Math.round((dc-25569)*86400000)),'UTC','yyyy-MM-dd')
               : fmtDate(dc);
        if (ds === roundDate && bbData[i][1] === name) {
          bbSheet.deleteRow(i + 1);
          bbDeleted++;
        }
      }
    }

    logToSheet('removePlayer', 'Removed ' + name + ' — rounds:' + roundsDeleted + ' bets:' + betsDeleted + ' bbHoles:' + bbDeleted);
    return {
      success: true,
      name: name,
      roundsDeleted: roundsDeleted,
      betsDeleted: betsDeleted,
      bbDeleted: bbDeleted,
      message: name + ' removed. ' + betsDeleted + ' bet(s) voided. Skins unchanged.'
    };
  } finally { lock.releaseLock(); }
}



// ── SCORING UTILITIES FOR 2-DOWN BETS ──

// Crown Colony stroke index per hole (1-indexed: hole 1=SI3, hole 2=SI11, etc.)
var STROKE_INDEX = [3,11,9,15,7,5,17,13,1, 2,12,14,10,18,6,8,16,4];

// Compute net score for one player on one hole given their handicap
// hiPct: 0=gross, 80=80% HI, 100=full HI
// pops: array of 0-indexed hole numbers where this player gets a pop (SU bets only)
function computeNetScore(grossScore, handicap, holeIndex, hiPct, pops) {
  if (grossScore === null || grossScore === undefined || grossScore === '') return null;
  var g = parseInt(grossScore);
  if (isNaN(g) || g <= 0) return null;

  // Apply pops for SU bets (pops override hiPct when present)
  if (pops && pops.length > 0) {
    var popStroke = pops.indexOf(holeIndex) > -1 ? 1 : 0;
    return g - popStroke;
  }

  if (hiPct === 0) return g;

  // 80% or 100% HI: apply fractional handicap
  var effectiveHcp = Math.round(Math.abs(handicap) * hiPct / 100);
  var si = STROKE_INDEX[holeIndex]; // 1-indexed SI
  var stroke = (si <= effectiveHcp) ? 1 : 0;
  // Extra stroke if handicap > 18
  if (effectiveHcp > 18 && si <= (effectiveHcp - 18)) stroke = 2;
  return g - stroke;
}

// Best ball for a team on a given hole
// holeScores: {playerName: {gross:[], net:[], handicap:N}}
// Returns minimum net/gross score for the team, or null if no data
// popsMap: {playerName: [holeIndex, ...]} for SU bets with manual pops
// hiExclude: [holeIndex, ...] holes to exclude from HI distribution (80%/100% bets)
// baseHcp: lowest handicap in the ENTIRE bet — used for delta HI calculation
function bestBallScore(teamPlayers, holeIndex, holeScores, hiPct, popsMap, hiExclude, baseHcp) {
  var scores = [];
  for (var i = 0; i < teamPlayers.length; i++) {
    var p = teamPlayers[i];
    var pd = holeScores[p];
    if (!pd) continue;
    var gross = pd.gross && pd.gross.length > holeIndex ? pd.gross[holeIndex] : null;
    if (gross === null || gross === '' || isNaN(parseInt(gross))) continue;
    var handicap = Math.abs(pd.handicap || 0);

    var score;
    if (hiPct === 0 && popsMap) {
      // SU bet: use manual pops grid
      var playerPops = popsMap[p] || [];
      score = computeNetScore(gross, handicap, holeIndex, 0, playerPops);
    } else if (hiPct > 0) {
      // Delta HI: player gets strokes based on (their HI - lowest HI in bet) * hiPct/100
      // This means lowest HI player always gets 0 strokes
      var base = (baseHcp !== null && baseHcp !== undefined) ? baseHcp : 0;
      var deltaHcp = Math.max(0, Math.round((handicap - base) * hiPct / 100));
      var si = STROKE_INDEX[holeIndex];

      // Apply hiExclude: shift strokes past excluded holes
      if (hiExclude && hiExclude.length > 0 && hiExclude.indexOf(holeIndex) > -1) {
        // This hole excluded — no stroke regardless
        score = parseInt(gross);
      } else {
        // Adjust SI rank by counting excluded holes with lower SI
        var adjustedSi = si;
        if (hiExclude && hiExclude.length > 0) {
          var excludedBelow = 0;
          hiExclude.forEach(function(exH) {
            if (STROKE_INDEX[exH] < si) excludedBelow++;
          });
          adjustedSi = si - excludedBelow;
        }
        var stroke = (adjustedSi <= deltaHcp) ? 1 : 0;
        if (deltaHcp > 18 && adjustedSi <= (deltaHcp - 18)) stroke = 2;
        score = parseInt(gross) - stroke;
      }
    } else {
      score = parseInt(gross);
    }

    if (score !== null && !isNaN(score)) scores.push(score);
  }
  if (!scores.length) return null;
  return Math.min.apply(null, scores);
}

// Compute full running differential for a bet from startHole to current
// Returns {diff, lastHole, complete} where diff>0 = Team A winning
function computeBetDiff(teamsA, teamsB, startHole, endHole, holeScores, hiPct, popsMap, hiExclude) {
  var diff = 0;
  var lastHole = 0;

  // Compute minThru across ALL players in both teams combined
  // This ensures all scopes (front/back/long) stop at the same hole
  var allPlayers = teamsA.concat(teamsB);
  var minThru = endHole; // start at max
  allPlayers.forEach(function(p) {
    var pd = holeScores[p];
    if (!pd) { minThru = 0; return; }
    // Find how many holes this player has completed
    var arr = (hiPct === 0) ? pd.gross : pd.net;
    if (!arr) { minThru = 0; return; }
    var playerThru = 0;
    for (var i = 0; i < endHole; i++) {
      if (arr[i] === null || arr[i] === undefined || arr[i] === '') break;
      playerThru++;
    }
    if (playerThru < minThru) minThru = playerThru;
  });

  // Cap endHole at minThru so all scopes agree on last completed hole
  var effectiveEnd = Math.min(endHole, minThru);

  // Compute base handicap (lowest in entire bet) for delta HI
  var baseHcp = null;
  if (hiPct > 0) {
    allPlayers.forEach(function(p) {
      var pd = holeScores[p];
      if (!pd) return;
      var hcp = Math.abs(pd.handicap || 0);
      if (baseHcp === null || hcp < baseHcp) baseHcp = hcp;
    });
    if (baseHcp === null) baseHcp = 0;
  }

  for (var h = startHole - 1; h < effectiveEnd; h++) {
    var aScore = bestBallScore(teamsA, h, holeScores, hiPct, popsMap, hiExclude, baseHcp);
    var bScore = bestBallScore(teamsB, h, holeScores, hiPct, popsMap, hiExclude, baseHcp);
    if (aScore === null || bScore === null) break;
    diff += (bScore - aScore);
    lastHole = h + 1;
  }
  var complete = (lastHole === endHole);
  
  return {diff: diff, lastHole: lastHole, complete: complete};
}

// ════════════════════════════════════════════════════════
// PAIRING BETS (2-Down Tab)
// ════════════════════════════════════════════════════════

function getPairingSheet() {
  var ss = SS();
  var sh = ss.getSheetByName('PairingBets');
  if (!sh) { setupSheets(); sh = ss.getSheetByName('PairingBets'); }
  return sh;
}

function getPairingBetsForDate(roundDate) {
  var sh = getPairingSheet();
  var data = sh.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][PB.DATE]) continue;
    var d = data[i][PB.DATE] instanceof Date
      ? Utilities.formatDate(data[i][PB.DATE], TIMEZONE, 'yyyy-MM-dd')
      : fmtDate(data[i][PB.DATE]);
    if (d === roundDate) results.push({row: data[i], idx: i + 1});
  }
  return results;
}

// ── CREATE PAIRING BET ──
function createPairingBet(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var roundDate = getRoundDate();
    if (!roundDate) return {success:false, error:'No active round'};
    var rs = getRoundStatus();
    if (rs.locked) return {success:false, error:'Round is locked'};

    // Validate required params
    var teamsA = (params.teams_a||'').toString().trim(); // pipe-separated player names
    var teamsB = (params.teams_b||'').toString().trim();
    var amount = parseFloat(params.amount||'0');
    var amountB = parseFloat(params.amount_b||params.amount||'0');
    if (!amountB || amountB <= 0) amountB = amount;
    var format = (params.format||'fb').toString(); // 'fb' or 'fbl'
    var hiPct = parseInt(params.hi_pct||'100');
    var autoPress = params.auto_press === 'true' || params.auto_press === true;
    var pressThreshold = parseInt(params.press_threshold||'2');
    var playersA0 = (params.teams_a||'').toString().split('|').map(function(p){return p.trim();}).filter(Boolean);
    var createdBy = (params.created_by||playersA0[0]||'').toString().trim();

    if (!teamsA || !teamsB) return {success:false, error:'Both teams required'};
    if (amount <= 0) return {success:false, error:'Amount must be greater than 0'};

    var playersA = teamsA.split('|').map(function(p){return p.trim();}).filter(Boolean);
    var playersB = teamsB.split('|').map(function(p){return p.trim();}).filter(Boolean);
    if (!playersA.length || !playersB.length) return {success:false, error:'Each team needs at least one player'};

    // Check no player on both teams
    var overlap = playersA.filter(function(p){ return playersB.indexOf(p) > -1; });
    if (overlap.length) return {success:false, error:overlap[0]+' cannot be on both teams'};

    var sh = getPairingSheet();
    var betId = 'PB-' + generateBetId();
    var now = nowTimestamp();
    var rdp = roundDate.split('-');
    var rdObj = new Date(parseInt(rdp[0]), parseInt(rdp[1])-1, parseInt(rdp[2]), 12, 0, 0);

    // Create front bet
    // Store amounts as "amtA|amtB" so both team amounts are preserved
    var amountStr = amount + (amountB !== amount ? '|' + amountB : '');
    // betId is the GROUP ID — back and long rows use betId as their parent_id
    // so cascade delete on betId removes all related rows
    // group_id = betId (front row's ID) shared by all rows in this bet group
    var groupId = betId;
    var popsStr = params.pops ? params.pops.toString() : '';
    
        sh.appendRow([rdObj, betId, '', teamsA, teamsB, amountStr, format, hiPct, 1, 'front', 'ACTIVE', '', '', '', '', createdBy, now, autoPress ? 'true' : 'false', pressThreshold, groupId, '', 0, '', popsStr]);

    // Create back bet — parent_id empty (original row), group_id = groupId
    var backId = 'PB-' + generateBetId();
    sh.appendRow([rdObj, backId, '', teamsA, teamsB, amountStr, format, hiPct, 10, 'back', 'ACTIVE', '', '', '', '', createdBy, now, autoPress ? 'true' : 'false', pressThreshold, groupId, '', 0, '', popsStr]);

    // Create long bet if fbl format — parent_id empty (original row), group_id = groupId
    var longId = null;
    if (format === 'fbl') {
      longId = 'PB-' + generateBetId();
      sh.appendRow([rdObj, longId, '', teamsA, teamsB, amountStr, format, hiPct, 1, 'long', 'ACTIVE', '', '', '', '', createdBy, now, 'false', pressThreshold, groupId, '', 0, '', popsStr]);
    }

    logToSheet('createPairingBet', createdBy+' created '+format.toUpperCase()+' $'+amount+'/player: '+teamsA+' vs '+teamsB);
    // Auto-run detectPresses immediately after bet creation
    // This handles both live and post-round bet entry
    try { detectPresses({}); } catch(e) { logToSheet('createPairingBet', 'detectPresses error: '+e.toString()); }

    return {success:true, betId:betId, backId:backId, longId:longId, roundDate:roundDate};
  } finally { lock.releaseLock(); }
}

// ── GET PAIRING BETS ──
function getPairingBets(params) {
  var roundDate = params && params.date ? params.date : getRoundDate();
  if (!roundDate) return {bets:[], roundDate:null};
  var rows = getPairingBetsForDate(roundDate);
  // Build a quick lookup: betId -> scope for press detection
  var betScopeMap = {};
  rows.forEach(function(r) {
    betScopeMap[(r.row[PB.BET_ID]||'').toString()] = (r.row[PB.SCOPE]||'').toString();
  });

  var bets = rows.map(function(r) {
    var myScope = (r.row[PB.SCOPE]||'').toString();
    var myParentId = (r.row[PB.PARENT_ID]||'').toString();
    var myGroupId = (r.row[PB.GROUP_ID]||'').toString();
    // isPress = has a parent AND parent has the SAME scope
    var parentScope = myParentId ? (betScopeMap[myParentId]||'') : '';
    var isPress = myParentId && parentScope === myScope;
    
    return {
      betId: (r.row[PB.BET_ID]||'').toString(),
      parentId: isPress ? myParentId : '', // only set parentId for true presses
      groupId: myGroupId,
      teamsA: (r.row[PB.TEAMS_A]||'').toString(),
      teamsB: (r.row[PB.TEAMS_B]||'').toString(),
      amount: (function(){ var a=(r.row[PB.AMOUNT]||'0').toString(); var p=a.split('|'); return parseFloat(p[0]||0); })(),
      amountB: (function(){ var a=(r.row[PB.AMOUNT]||'0').toString(); var p=a.split('|'); return p.length>1?parseFloat(p[1]):parseFloat(p[0]||0); })(),
      format: (r.row[PB.FORMAT]||'fb').toString(),
      hiPct: r.row[PB.HI_PCT]!==''&&r.row[PB.HI_PCT]!==null&&r.row[PB.HI_PCT]!==undefined ? parseInt(r.row[PB.HI_PCT]) : 100,
      startHole: parseInt(r.row[PB.START_HOLE]||1),
      scope: (r.row[PB.SCOPE]||'front').toString(),
      status: (r.row[PB.STATUS]||'ACTIVE').toString(),
      frontResult: (r.row[PB.FRONT_RESULT]||'').toString(),
      backResult: (r.row[PB.BACK_RESULT]||'').toString(),
      longResult: (r.row[PB.LONG_RESULT]||'').toString(),
      amountStr: (r.row[PB.AMOUNT_STR]||'').toString(),
      createdBy: (r.row[PB.CREATED_BY]||'').toString(),
      autoPress: (r.row[PB.AUTO_PRESS]||'false').toString() === 'true',
      pressThreshold: parseInt(r.row[PB.PRESS_THRESHOLD]||2),
      currentDiff: r.row[PB.CURRENT_DIFF]!==''&&r.row[PB.CURRENT_DIFF]!==null ? parseInt(r.row[PB.CURRENT_DIFF]) : null,
      lastHole: parseInt(r.row[PB.LAST_HOLE]||0),
      finalDiff: r.row[PB.FINAL_DIFF]!==''&&r.row[PB.FINAL_DIFF]!==null ? parseInt(r.row[PB.FINAL_DIFF]) : null,
      pops: (function(){ 
        try { 
          var raw = (r.row[PB.POPS]||'{}').toString();
          var parsed = JSON.parse(raw);
          if (typeof parsed === 'string') parsed = JSON.parse(parsed);
          return parsed;
        } catch(e) { return {}; } 
      })(),
      rowIndex: r.idx
    };
  });
  return {bets:bets, roundDate:roundDate};
}

// ── DELETE PAIRING BET ──
// Players can delete their own bets; admin can delete any
function deletePairingBet(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var betId = (params.betId||'').toString();
    var requestedBy = (params.requestedBy||'').toString();
    var isAdmin = params.isAdmin === 'true' || params.isAdmin === true;
    var pressOnly = params.pressOnly === 'true' || params.pressOnly === true;

    // pressOnly: delete just this single row (a press), not the whole group
    if (pressOnly && betId) {
      var sh = getPairingSheet();
      var data = sh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if ((data[i][PB.BET_ID]||'').toString() === betId) {
          sh.deleteRow(i + 1);
          logToSheet('deletePairingBet', 'Deleted press ' + betId);
          return {success: true};
        }
      }
      return {success: false, error: 'Press not found'};
    }

    if (!betId) return {success:false, error:'No bet ID'};

    var sh = getPairingSheet();
    var data = sh.getDataRange().getValues();
    var toDelete = [];

    // Find the group_id from the target bet row
    var targetGroupId = '';
    var roundDate = getRoundDate();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][PB.BET_ID]||'').toString() === betId) {
        targetGroupId = (data[i][PB.GROUP_ID]||betId).toString();
        toDelete.push(i + 1);
        break;
      }
    }
    // If not found by betId, look for group via teams+date
    var teamsAParam = (params.teamsA||'').toString();
    var teamsBParam = (params.teamsB||'').toString();
    if (!targetGroupId && teamsAParam && teamsBParam) {
      for (var i = 1; i < data.length; i++) {
        if (!data[i][PB.DATE]) continue;
        var rd = data[i][PB.DATE] instanceof Date
          ? Utilities.formatDate(data[i][PB.DATE], TIMEZONE, 'yyyy-MM-dd')
          : fmtDate(data[i][PB.DATE]);
        if (rd !== roundDate) continue;
        if ((data[i][PB.TEAMS_A]||'').toString() === teamsAParam &&
            (data[i][PB.TEAMS_B]||'').toString() === teamsBParam) {
          targetGroupId = (data[i][PB.GROUP_ID]||data[i][PB.BET_ID]||'').toString();
          toDelete.push(i + 1);
          break;
        }
      }
    }
    if (!targetGroupId) return {success:false, error:'Bet not found'};

    // Delete ALL rows sharing this group_id (front + back + long + all presses)
    for (var i = 1; i < data.length; i++) {
      if (toDelete.indexOf(i+1) > -1) continue;
      var rowGroupId = (data[i][PB.GROUP_ID]||'').toString();
      var rowBetId = (data[i][PB.BET_ID]||'').toString();
      var rowParentId = (data[i][PB.PARENT_ID]||'').toString();
      // Match by group_id OR by parent chain for legacy rows
      if (rowGroupId === targetGroupId ||
          rowBetId === targetGroupId ||
          rowParentId === targetGroupId) {
        toDelete.push(i + 1);
      }
    }
    // Legacy team+date fallback removed — use group_id only

    toDelete.sort(function(a,b){return b-a;});
    toDelete.forEach(function(r){ sh.deleteRow(r); });
    logToSheet('deletePairingBet', requestedBy+' deleted bet '+betId+' and '+( toDelete.length-1)+' press(es)');
    return {success:true, deleted:toDelete.length};
  } finally { lock.releaseLock(); }
}

// ── ADD MANUAL PRESS ──
function addManualPress(params) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var parentId = (params.parentId||'').toString();
    var startHole = parseInt(params.startHole||0);
    var scope = (params.scope||'front').toString();
    var requestedBy = (params.requestedBy||'').toString();
    if (!parentId || !startHole) return {success:false, error:'Parent bet ID and start hole required'};

    var sh = getPairingSheet();
    var data = sh.getDataRange().getValues();
    var parent = null;
    for (var i = 1; i < data.length; i++) {
      if ((data[i][PB.BET_ID]||'').toString() === parentId) { parent = data[i]; break; }
    }
    if (!parent) return {success:false, error:'Parent bet not found'};

    // Validate requester is in the bet
    var teamsA = (parent[PB.TEAMS_A]||'').toString();
    var teamsB = (parent[PB.TEAMS_B]||'').toString();
    // Anyone can add a press

    var roundDate = getRoundDate();
    var rdp = roundDate.split('-');
    var rdObj = new Date(parseInt(rdp[0]), parseInt(rdp[1])-1, parseInt(rdp[2]), 12, 0, 0);
    var pressId = 'PB-' + generateBetId();
    var now = nowTimestamp();

    var pressGroupId = (parent[PB.GROUP_ID]||parent[PB.BET_ID]||'').toString();
    sh.appendRow([
      rdObj, pressId, parentId, teamsA, teamsB,
      parent[PB.AMOUNT], parent[PB.FORMAT], parent[PB.HI_PCT],
      startHole, scope, 'ACTIVE', '', '', '', '',
      requestedBy, now, 'false', parent[PB.PRESS_THRESHOLD], pressGroupId, '', 0, ''
    ]);

    logToSheet('addManualPress', requestedBy+' added press on '+parentId+' starting hole '+startHole);
    // Auto-run detectPresses so manual press gets diffs immediately
    try { detectPresses({}); } catch(e) { logToSheet('addManualPress', 'detectPresses error: '+e.toString()); }
    return {success:true, pressId:pressId};
  } finally { lock.releaseLock(); }
}

// ── DETECT AUTO PRESSES AND UPDATE RUNNING DIFFS ──
// Called after each score refresh.
// 1. Fetches live hole-by-hole scores from GG (gross + net + handicap per player)
// 2. Computes current_diff and last_hole for ALL active bets and writes to sheet
// 3. For the MOST RECENTLY OPENED bet per scope: if |diff| >= threshold, creates new press
// 4. Marks completed scopes as FINAL with final_diff
function detectPresses(params) {
  var roundDate = getRoundDate();
  if (!roundDate) return {success:false, error:'No active round'};
  var rows = getPairingBetsForDate(roundDate);
  if (!rows.length) return {success:true, pressed:0, message:'No pairing bets'};
  var sh = getPairingSheet();

  // Fetch live hole scores from GG — gross + handicap per player
  var holeData = getHoleByHole();
  if (!holeData || !holeData.available || !holeData.players) {
    return {success:true, pressed:0, message:'No hole data available yet'};
  }
  var holeScores = holeData.players; // {name: {gross:[], net:[], handicap:N, thru:N}}

  var rdp = roundDate.split('-');
  var rdObj = new Date(parseInt(rdp[0]), parseInt(rdp[1])-1, parseInt(rdp[2]), 12, 0, 0);
  var pressed = 0;
  var newPresses = [];

  // ── Pass 1: Update current_diff and last_hole for ALL active bets ──
  // Also collect bets by pairing+scope to find most-recent for press detection
  var pairingScopes = {}; // key -> [{row, idx, startHole, betId}]

  rows.forEach(function(r) {
    var bet = r.row;
    var status = (bet[PB.STATUS]||'').toString();
    // Process ALL rows including FINAL — always recompute with latest GG data

    var scope = (bet[PB.SCOPE]||'').toString();
    var startHole = parseInt(bet[PB.START_HOLE]||1);
    var endHole = scope === 'front' ? 9 : 18;
    var hiPct = parseInt(bet[PB.HI_PCT]);
    if (isNaN(hiPct)) hiPct = 100;
    var teamsA = (bet[PB.TEAMS_A]||'').toString().split('|').map(function(p){return p.trim();});
    var teamsB = (bet[PB.TEAMS_B]||'').toString().split('|').map(function(p){return p.trim();});

    // Compute diff from startHole to current
    var betPopsRaw = (bet[PB.POPS]||'{}').toString();
var betPopsConfig = {};
try { 
  betPopsConfig = JSON.parse(betPopsRaw);
  if (typeof betPopsConfig === 'string') betPopsConfig = JSON.parse(betPopsConfig);
} catch(e) {}
// If this bet has no pops (press row), inherit from group's original bet
if (hiPct === 0 && Object.keys(betPopsConfig).length === 0) {
  var groupId0 = (bet[PB.GROUP_ID]||'').toString();
  if (groupId0) {
    rows.forEach(function(r2) {
      if ((r2.row[PB.BET_ID]||'').toString() === groupId0) {
        var raw2 = (r2.row[PB.POPS]||'{}').toString();
        try {
          var p2 = JSON.parse(raw2);
          if (typeof p2 === 'string') p2 = JSON.parse(p2);
          if (Object.keys(p2).length > 0) betPopsConfig = p2;
        } catch(e2) {}
      }
    });
  }
}
var betPopsMap = (hiPct === 0) ? betPopsConfig : null;
var betHiExclude = (hiPct > 0 && betPopsConfig.exclude) ? betPopsConfig.exclude : null;
var result = computeBetDiff(teamsA, teamsB, startHole, endHole, holeScores, hiPct, betPopsMap, betHiExclude);
    

    // Write current_diff and last_hole back to sheet
    // Always write current diffs based on latest GG data
    sh.getRange(r.idx, PB.CURRENT_DIFF+1).setValue(result.lastHole > 0 ? result.diff : '');
    sh.getRange(r.idx, PB.LAST_HOLE+1).setValue(result.lastHole);
    

    // If scope complete, write final_diff and mark FINAL
    if (result.complete) {
      sh.getRange(r.idx, PB.FINAL_DIFF+1).setValue(result.diff);
      sh.getRange(r.idx, PB.STATUS+1).setValue('FINAL');
      var winner = result.diff > 0 ? 'A' : result.diff < 0 ? 'B' : 'PUSH';
      var amtRaw = (bet[PB.AMOUNT]||'0').toString().split('|');
      var amtA = parseFloat(amtRaw[0]||0);
      var amtStr = 'A: '+teamsA.join('/')+' vs B: '+teamsB.join('/')+' → '+(winner==='PUSH'?'PUSH':winner+' WINS by '+Math.abs(result.diff))+' ($'+amtA+'/player)';
      sh.getRange(r.idx, PB.FRONT_RESULT+1).setValue(winner);
      sh.getRange(r.idx, PB.AMOUNT_STR+1).setValue(amtStr);
      // NOTE: do NOT return here — still need to add to pairingScopes for press detection
    }

    // Add to pairingScopes for press detection (front/back only, auto-press only)
    // Include complete scopes — post-round bets need press detection even when FINAL
    var autoPress = (bet[PB.AUTO_PRESS]||'false').toString() === 'true';
    if (!autoPress || scope === 'long') return;

    var tA = teamsA.slice().sort().join('|');
    var tB = teamsB.slice().sort().join('|');
    var groupId1 = (bet[PB.GROUP_ID]||'').toString();
    var key = tA + '||' + tB + '||' + scope + '||' + groupId1;
    if (!pairingScopes[key]) pairingScopes[key] = [];
    pairingScopes[key].push({
      row: bet, idx: r.idx, startHole: startHole,
      betId: (bet[PB.BET_ID]||'').toString(),
      diff: result.diff, lastHole: result.lastHole,
      threshold: parseInt(bet[PB.PRESS_THRESHOLD]||2),
      hiPct: hiPct, endHole: endHole
    });
  });

  // ── Pass 2: Press detection — full scan for all missing presses ──
  // Works for both live and post-round bets
  // Simulate hole-by-hole and create ALL presses that should exist
  Object.keys(pairingScopes).forEach(function(key) {
    var betsForScope = pairingScopes[key];
    betsForScope.sort(function(a,b){ return a.startHole - b.startHole; }); // chronological

    var scopeInfo = betsForScope[0];
    var teamsA = (scopeInfo.row[PB.TEAMS_A]||'').toString().split('|').map(function(p){return p.trim();});
    var teamsB = (scopeInfo.row[PB.TEAMS_B]||'').toString().split('|').map(function(p){return p.trim();});
    var hiPct = scopeInfo.hiPct;
    var threshold = scopeInfo.threshold;
    var scope = (scopeInfo.row[PB.SCOPE]||'').toString();
    var endHole = scope === 'front' ? 9 : 18;
    var groupId = (scopeInfo.row[PB.GROUP_ID]||scopeInfo.betId).toString();

    // Build set of existing start holes for this scope
    var existingStartHoles = {};
    betsForScope.forEach(function(b){ existingStartHoles[b.startHole] = true; });

    // Compute minThru for this scope to ensure consistent hole counting
    var simMinThru = endHole;
    teamsA.concat(teamsB).forEach(function(p) {
      var pd = holeScores[p];
      if (!pd || !pd.gross) { simMinThru = 0; return; }
      var playerThru = 0;
      for (var i = 0; i < endHole; i++) {
        if (pd.gross[i] === null || pd.gross[i] === undefined || pd.gross[i] === '') break;
        playerThru++;
      }
      if (playerThru < simMinThru) simMinThru = playerThru;
    });
    var simEndHole = Math.min(endHole, simMinThru);

    // Simulate hole by hole — track the most recent active bet's running diff
    var activeBets = [];
    var origBet = betsForScope[0];
    // Parse pops config for this scope
    var popsRaw = (origBet.row[PB.POPS]||'{}').toString();
    var popsConfig = {};
    try { 
      popsConfig = JSON.parse(popsRaw);
      if (typeof popsConfig === 'string') popsConfig = JSON.parse(popsConfig);
    } catch(e) {}
    var scopePopsMap = (hiPct === 0) ? popsConfig : null;
    var scopeHiExclude = (hiPct > 0 && popsConfig.exclude) ? popsConfig.exclude : null;
    
    // Compute base HCP for delta HI
    var scopeBaseHcp = null;
    if (hiPct > 0) {
      teamsA.concat(teamsB).forEach(function(p) {
        var pd = holeScores[p];
        if (!pd) return;
        var hcp = Math.abs(pd.handicap || 0);
        if (scopeBaseHcp === null || hcp < scopeBaseHcp) scopeBaseHcp = hcp;
      });
      if (scopeBaseHcp === null) scopeBaseHcp = 0;
    }

    activeBets.push({startHole: origBet.startHole, diff: 0, holesPlayed: 0,
                     betId: origBet.betId, parentId: origBet.betId, groupId: groupId});

    for (var h = origBet.startHole - 1; h < simEndHole; h++) {
      var hole = h + 1;
      var aBB = bestBallScore(teamsA, h, holeScores, hiPct, scopePopsMap, scopeHiExclude, scopeBaseHcp);
      var bBB = bestBallScore(teamsB, h, holeScores, hiPct, scopePopsMap, scopeHiExclude, scopeBaseHcp);
      if (aBB === null || bBB === null) break; // scores not available — minThru handles this

      var holeDiff = bBB - aBB; // positive = A winning

      // Update all active bets that have started
      activeBets.forEach(function(ab) {
        if (ab.startHole <= hole) {
          ab.diff += holeDiff;
          ab.holesPlayed++;
        }
      });

      // Check most recent bet for press trigger
      // Require at least 1 hole played on this bet before it can trigger a press
      var latest = activeBets[activeBets.length - 1];
      if (latest.startHole <= hole && latest.holesPlayed >= 1 &&
          Math.abs(latest.diff) >= threshold && hole < endHole) {
        var pressHole = hole + 1;
        if (!existingStartHoles[pressHole]) {
          // Press needed at pressHole — create it
          existingStartHoles[pressHole] = true;
          var pressId = 'PB-' + generateBetId();
          sh.appendRow([
            rdObj, pressId, latest.betId,
            scopeInfo.row[PB.TEAMS_A], scopeInfo.row[PB.TEAMS_B],
            scopeInfo.row[PB.AMOUNT], scopeInfo.row[PB.FORMAT], hiPct,
            pressHole, scope, 'ACTIVE', '', '', '', '',
            'auto', nowTimestamp(), 'true', threshold, groupId, '', 0, '', ''
          ]);
          activeBets.push({startHole: pressHole, diff: 0, holesPlayed: 0,
                           betId: pressId, parentId: latest.betId, groupId: groupId});
          pressed++;
          logToSheet('detectPresses', 'Auto-press created: '+scope+' h'+pressHole+' for '+teamsA.join('/')+' vs '+teamsB.join('/'));
        } else {
          // Press already exists — find it and add to activeBets tracker
          var existingPress = betsForScope.filter(function(b){ return b.startHole === pressHole; })[0];
          if (existingPress && !activeBets.some(function(ab){ return ab.startHole === pressHole; })) {
            activeBets.push({startHole: pressHole, diff: 0, holesPlayed: 0,
                             betId: existingPress.betId, parentId: latest.betId, groupId: groupId});
          }
        }
      }
    }
  });

  // ── Pass 3: Re-fetch rows (new presses may have been added) and update all diffs ──
  if (pressed > 0) {
    rows = getPairingBetsForDate(roundDate);
    rows.forEach(function(r) {
      var bet = r.row;
      var scope = (bet[PB.SCOPE]||'').toString();
      var startHole = parseInt(bet[PB.START_HOLE]||1);
      var endHole = scope === 'front' ? 9 : 18;
      var hiPct = parseInt(bet[PB.HI_PCT]); if (isNaN(hiPct)) hiPct = 100;
      var teamsA = (bet[PB.TEAMS_A]||'').toString().split('|').map(function(p){return p.trim();});
      var teamsB = (bet[PB.TEAMS_B]||'').toString().split('|').map(function(p){return p.trim();});
      var result = computeBetDiff(teamsA, teamsB, startHole, endHole, holeScores, hiPct);
      sh.getRange(r.idx, PB.CURRENT_DIFF+1).setValue(result.lastHole > 0 ? result.diff : '');
      sh.getRange(r.idx, PB.LAST_HOLE+1).setValue(result.lastHole);
      if (result.complete) {
        sh.getRange(r.idx, PB.FINAL_DIFF+1).setValue(result.diff);
        sh.getRange(r.idx, PB.STATUS+1).setValue('FINAL');
        var winner = result.diff > 0 ? 'A' : result.diff < 0 ? 'B' : 'PUSH';
        var amtRaw = (bet[PB.AMOUNT]||'0').toString().split('|');
        var amtA = parseFloat(amtRaw[0]||0);
        sh.getRange(r.idx, PB.FRONT_RESULT+1).setValue(winner);
        sh.getRange(r.idx, PB.AMOUNT_STR+1).setValue('A: '+teamsA.join('/')+' vs B: '+teamsB.join('/')+' → '+(winner==='PUSH'?'PUSH':winner+' WINS by '+Math.abs(result.diff))+' ($'+amtA+'/player)');
      }
    });
  }

  // Write new presses
  if (newPresses.length) {
    newPresses.forEach(function(p) {
      var pressId = 'PB-' + generateBetId();
      sh.appendRow([
        rdObj, pressId, p.parentId, p.teamsA, p.teamsB,
        p.amount, p.format, p.hiPct, p.startHole, p.scope,
        'ACTIVE', '', '', '', '', 'auto', nowTimestamp(),
        'true', p.pressThreshold, p.groupId||p.parentId, '', 0, ''
      ]);
    });
    logToSheet('detectPresses', 'Round '+roundDate+': updated diffs, created '+newPresses.length+' press(es)');
  } else {
    logToSheet('detectPresses', 'Round '+roundDate+': updated diffs for '+rows.length+' bets');
  }

  return {success:true, pressed:pressed};
}

// ── BEST BALL FOR HOLE — legacy wrapper, delegates to bestBallScore ──
function bestBallForHole(teamPlayers, holeIndex, holeScores, hiPct) {
  return bestBallScore(teamPlayers, holeIndex, holeScores, hiPct);
}

// ── GET HOLE SCORING DATA — reads BBHoles for BB rounds, GG API for Nassau ──
function getHoleScoringData(roundDate) {
  var ss = SS();

  // Try BBHoles sheet first (Best Ball rounds)
  var sh = ss.getSheetByName('BBHoles');
  if (sh) {
    var data = sh.getDataRange().getValues();
    var result = {};
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var cellDate = data[i][0];
      var ds = cellDate instanceof Date
        ? Utilities.formatDate(cellDate, TIMEZONE, 'yyyy-MM-dd')
        : fmtDate(cellDate);
      if (ds !== roundDate) continue;
      var player = data[i][1];
      var net = [];
      for (var h = 0; h < 18; h++) {
        var v = data[i][2 + h];
        net.push((v !== '' && v !== null && v !== undefined) ? Number(v) : null);
      }
      result[player] = {net: net, gross: net};
    }
    if (Object.keys(result).length) return result;
  }

  // Fallback: fetch live hole data from GG API (Nassau rounds)
  try {
    var roundId = null;
    var roundsSheet = ss.getSheetByName('Rounds');
    if (roundsSheet) {
      var rdata = roundsSheet.getDataRange().getValues();
      for (var i = 1; i < rdata.length; i++) {
        if (!rdata[i][0]) continue;
        var rd = fmtDate(rdata[i][0]);
        if (rd === roundDate && rdata[i][1]) { roundId = rdata[i][1].toString(); break; }
      }
    }
    if (!roundId) return null;

    // Get tournaments for this round
    var tournsUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments.json';
    var tournsData = JSON.parse(UrlFetchApp.fetch(tournsUrl).getContentText());
    var tourns = tournsData.event ? tournsData.event.tournaments || [] : [];

    // Find Nassau/stroke tournament
    var nassauId = null;
    for (var i = 0; i < tourns.length; i++) {
      var tname = (tourns[i].event ? tourns[i].event.name || '' : '').toLowerCase();
      if (tname.indexOf('nassau') > -1 || tname.indexOf('stroke') > -1) {
        nassauId = tourns[i].event ? tourns[i].event.id : null;
        break;
      }
    }
    // Fallback: use first tournament
    if (!nassauId && tourns.length) nassauId = tourns[0].event ? tourns[0].event.id : null;
    if (!nassauId) return null;

    var aggUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments/'+nassauId+'.json';
    var aggData = JSON.parse(UrlFetchApp.fetch(aggUrl).getContentText());
    var aggregates = aggData.event && aggData.event.scopes && aggData.event.scopes[0]
      ? aggData.event.scopes[0].aggregates || [] : [];

    var liveResult = {};
    aggregates.forEach(function(a) {
      if (!a.name) return;
      var gross = Array.isArray(a.gross_scores) ? a.gross_scores : [];
      var net = Array.isArray(a.net_scores) ? a.net_scores : [];
      // Pad to 18
      while (gross.length < 18) gross.push(null);
      while (net.length < 18) net.push(null);
      liveResult[a.name] = {
        gross: gross.map(function(v){ return (v !== null && v !== undefined && v !== '') ? Number(v) : null; }),
        net: net.map(function(v){ return (v !== null && v !== undefined && v !== '') ? Number(v) : null; })
      };
    });
    return Object.keys(liveResult).length ? liveResult : null;
  } catch(e) {
    logToSheet('getHoleScoringData', 'Live fetch error: '+e.toString());
    return null;
  }
}

// ── CALCULATE PAIRING BET RESULTS ──
function calculatePairingResults(roundDate) {
  if (!roundDate) roundDate = getRoundDate();
  if (!roundDate) return;
  var rows = getPairingBetsForDate(roundDate);
  if (!rows.length) return;
  var sh = getPairingSheet();
  // Use getHoleByHole for gross scores (correct for 0% HI straight-up bets)
  var holeData = getHoleByHole();
  if (!holeData || !holeData.available || !holeData.players) return;
  var holeScores = holeData.players;

  var parPerHole = getConfigValue('course_pars', '5,4,5,3,4,4,3,4,4,4,4,3,5,3,4,5,3,5')
    .split(',').map(Number);

  rows.forEach(function(r) {
    var bet = r.row;
    var status = (bet[PB.STATUS]||'').toString();
    var teamsA = (bet[PB.TEAMS_A]||'').toString().split('|').map(function(p){return p.trim();});
    var teamsB = (bet[PB.TEAMS_B]||'').toString().split('|').map(function(p){return p.trim();});
    var startHole = parseInt(bet[PB.START_HOLE]||1) - 1; // 0-indexed
    var scope = (bet[PB.SCOPE]||'front').toString();
    var hiPct = parseInt(bet[PB.HI_PCT]||100);
    var amountRaw = (bet[PB.AMOUNT]||'0').toString().split('|');
    var amount = parseFloat(amountRaw[0]||0);   // Team A per-player amount
    var amountB = amountRaw.length > 1 ? parseFloat(amountRaw[1]) : amount; // Team B per-player amount
    var endHole = scope === 'front' ? 9 : scope === 'back' ? 18 : 18;

    // Sum team best balls from startHole to endHole
    var aTotal = 0, bTotal = 0, allPlayed = true;
    for (var h = startHole; h < endHole; h++) {
      var aBB = bestBallForHole(teamsA, h, holeScores, hiPct);
      var bBB = bestBallForHole(teamsB, h, holeScores, hiPct);
      if (aBB === null || bBB === null) { allPlayed = false; break; }
      aTotal += aBB;
      bTotal += bBB;
    }
    if (!allPlayed) return; // not complete yet

    // Determine result
    var result = aTotal < bTotal ? 'A' : aTotal > bTotal ? 'B' : 'PUSH';
    var aPar = parPerHole.slice(startHole, endHole).reduce(function(s,v){return s+v;},0);
    var bPar = aPar;
    var margin = Math.abs(aTotal - bTotal);
    var resultStr = result === 'PUSH' ? 'PUSH' : (result === 'A' ? 'A WINS by '+margin : 'B WINS by '+margin);
    var amtStr = 'A: '+aTotal+' vs B: '+bTotal+' → '+resultStr+' ($'+amount+'/player)';

    // Update the row
    sh.getRange(r.idx, PB.FRONT_RESULT+1).setValue(result);
    sh.getRange(r.idx, PB.AMOUNT_STR+1).setValue(amtStr);
    sh.getRange(r.idx, PB.STATUS+1).setValue('FINAL');
  });
}

// ── GET PAIRING RESULTS — for Results/Settlement integration ──
function getPairingResults(params) {
  var roundDate = params && params.date ? params.date : getRoundDate();
  if (!roundDate) return {bets:[], playerTotals:{}};

  // detectPresses handles all calculations — calculatePairingResults removed
  // to prevent it overwriting detectPresses results with wrong net scores
  var rows = getPairingBetsForDate(roundDate);
  var playerTotals = {};
  var betGroups = {}; // key = sorted(teamsA,teamsB) -> array of bet results

  rows.forEach(function(r) {
    var bet = r.row;
    var status = (bet[PB.STATUS]||'').toString();
    var isPartial = status === 'PARTIAL';
    if (status !== 'FINAL' && !isPartial) return;

    var teamsA = (bet[PB.TEAMS_A]||'').toString().split('|').map(function(p){return p.trim();});
    var teamsB = (bet[PB.TEAMS_B]||'').toString().split('|').map(function(p){return p.trim();});
    var amountRaw = (bet[PB.AMOUNT]||'0').toString().split('|');
    var amount = parseFloat(amountRaw[0]||0);
    var amountBv = amountRaw.length > 1 ? parseFloat(amountRaw[1]) : amount;
    // For PARTIAL: use front result as projection
    var result = (bet[PB.FRONT_RESULT]||'').toString();
    if (!result && isPartial) return; // front not done yet
    var amtStr = (bet[PB.AMOUNT_STR]||'').toString();
    var startHole = parseInt(bet[PB.START_HOLE]||1);
    var scope = (bet[PB.SCOPE]||'').toString();
    var parentId = (bet[PB.PARENT_ID]||'').toString();

    // Tally per-player results — each player wins/loses their own per-player amount
    if (result === 'A') {
      teamsA.forEach(function(p){if(!playerTotals[p])playerTotals[p]=0; playerTotals[p]+=amount;});
      teamsB.forEach(function(p){if(!playerTotals[p])playerTotals[p]=0; playerTotals[p]-=amountBv;});
    } else if (result === 'B') {
      teamsB.forEach(function(p){if(!playerTotals[p])playerTotals[p]=0; playerTotals[p]+=amountBv;});
      teamsA.forEach(function(p){if(!playerTotals[p])playerTotals[p]=0; playerTotals[p]-=amount;});
    }

    // Group bets for scorecard display
    var groupKey = [teamsA.sort().join('|'), teamsB.sort().join('|')].join('||');
    if (!betGroups[groupKey]) betGroups[groupKey] = {teamsA:teamsA, teamsB:teamsB, bets:[]};
    betGroups[groupKey].bets.push({
      betId:(bet[PB.BET_ID]||'').toString(),
      parentId:parentId,
      startHole:startHole,
      scope:scope,
      result:result,
      amount:amount,
      amtStr:amtStr
    });
  });

  // Sort each group's bets by startHole for scorecard display
  Object.keys(betGroups).forEach(function(k) {
    betGroups[k].bets.sort(function(a,b){return a.startHole-b.startHole;});
  });

  return {betGroups:betGroups, playerTotals:playerTotals, roundDate:roundDate};
}

function debugBestBallPayload() {
  var ss = SS();
  var rd = ss.getSheetByName('Rounds').getDataRange().getValues();
  if (rd.length < 2) { Logger.log('No active round'); return; }
  var roundId = rd[1][R.ROUND_ID].toString();
  Logger.log('Round ID: ' + roundId);

  // Get all tournaments
  var tournsUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments?holes_scope=nassau';
  var tourns = JSON.parse(UrlFetchApp.fetch(tournsUrl, {muteHttpExceptions:true}).getContentText());
  Logger.log('All tournaments: ' + tourns.map(function(t){return t.event.name+':'+t.event.score_format+':'+t.event.id;}).join(', '));

  // Find Best Ball tournament
  var bbId = null;
  for (var i = 0; i < tourns.length; i++) {
    if ((tourns[i].event.name||'').toLowerCase().indexOf('best ball') > -1) {
      bbId = tourns[i].event.id;
      Logger.log('Found Best Ball tournament: id=' + bbId);
      break;
    }
  }
  if (!bbId) { Logger.log('No Best Ball tournament found'); return; }

  // Fetch BB payload
  var url = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments/'+bbId+'.json';
  var raw = UrlFetchApp.fetch(url, {muteHttpExceptions:true}).getContentText();
  var data = JSON.parse(raw);
  var agg = data.event.scopes[0].aggregates;

  Logger.log('Aggregate count: ' + agg.length);
  if (agg.length > 0) {
    // Log first aggregate's top-level keys
    Logger.log('First agg keys: ' + Object.keys(agg[0]).join(', '));
    Logger.log('First agg name: ' + agg[0].name);

    // Log totals structure
    if (agg[0].totals) {
      Logger.log('totals keys: ' + Object.keys(agg[0].totals).join(', '));
      if (agg[0].totals.net_scores) {
        Logger.log('net_scores keys: ' + Object.keys(agg[0].totals.net_scores).join(', '));
        Logger.log('net_scores.out (front): ' + agg[0].totals.net_scores.out);
        Logger.log('net_scores.in (back): ' + agg[0].totals.net_scores.in);
        Logger.log('net_scores.total: ' + agg[0].totals.net_scores.total);
        // Check for hole array
        var holes = agg[0].totals.net_scores.holes;
        if (holes) {
          Logger.log('net_scores.holes length: ' + holes.length);
          Logger.log('net_scores.holes (first 9): ' + holes.slice(0,9).join(', '));
        } else {
          Logger.log('net_scores.holes: NOT PRESENT');
        }
      }
      if (agg[0].totals.gross_scores) {
        Logger.log('gross_scores.out: ' + agg[0].totals.gross_scores.out);
        Logger.log('gross_scores.in: ' + agg[0].totals.gross_scores.in);
      }
    }

    // Check top-level net_scores (alternate location)
    if (agg[0].net_scores) {
      Logger.log('top-level net_scores type: ' + typeof agg[0].net_scores);
      if (Array.isArray(agg[0].net_scores)) {
        Logger.log('top-level net_scores (array, first 9): ' + agg[0].net_scores.slice(0,9).join(', '));
      }
    }

    Logger.log('thru: ' + agg[0].thru);

    // Log second aggregate if exists (check if teams or individuals)
    if (agg.length > 1) {
      Logger.log('Second agg name: ' + agg[1].name);
    }
    // If names contain '&' or '/' they are team entries, otherwise individual
    var hasAmpersand = agg.some(function(a){ return (a.name||'').indexOf('&') > -1 || (a.name||'').indexOf('/') > -1; });
    Logger.log('Aggregates appear to be TEAMS: ' + hasAmpersand);
  }
}

function debugGolfGeniusPayload() {
  // Get any active round ID from your RoundsList sheet
  var ss = SS();
  var roundId = ss.getSheetByName('Rounds List').getRange('C2').getValue();
  // Note: use www.golfgenius.com, not api.golfgenius.com
  var tournsUrl = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments?holes_scope=nassau';
  var tourns = JSON.parse(UrlFetchApp.fetch(tournsUrl, {muteHttpExceptions:true}).getContentText());
  if (!tourns||!tourns.length) { Logger.log('No tournaments found'); return; }
  var nassauId = tourns[0].event.id;
  var url = 'https://www.golfgenius.com/api_v2/'+API_KEY+'/events/'+EVENT_ID+'/rounds/'+roundId+'/tournaments/'+nassauId+'.json';
  var response = UrlFetchApp.fetch(url, {muteHttpExceptions:true});
  Logger.log(response.getContentText());
}

function debugIndieSheet() {
  var sh = SS().getSheetByName('SeasonIndie');
  var data = sh.getDataRange().getValues();
  return {rows: data.slice(0,3).map(function(r){ return [typeof r[0], r[0].toString(), r[1], r[2], r[3]]; })};
}

function testIndieDelete() {
  var sh = SS().getSheetByName('SeasonIndie');
  var lastRow = sh.getLastRow();
  for (var i = 1; i <= lastRow; i++) {
    var d = sh.getRange(i, 1).getValue();
    Logger.log('Row ' + i + ': type=' + typeof d + ' value=' + d + ' toString=' + d.toString());
  }
}
