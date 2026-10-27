// ================================================================
// CFB FUN BETS HUB — Google Apps Script Backend
// ================================================================
// SETUP:
//   1. In your Google Sheet: Extensions > Apps Script
//   2. Delete the default code, paste this entire file
//   3. Deploy > New Deployment > Web App
//      - Execute as: Me
//      - Who has access: Anyone
//   4. Copy the Web App URL into the site's Setup screen
//   5. Use the site's Setup > Initialize to set your admin password

const SS_ID = '12ocyOtb8mrr977WpvjXpvscY96mCeNsR96SaZ7Xm6Xw';

// ── Sheet names from your existing workbook ─────────────────────
const BET_SHEETS = [
  { name: 'Bet 1',     type: 'game'  },
  { name: 'Bet 2',     type: 'prop'  },
  { name: 'Bet 3',     type: 'game'  },
  { name: 'Bet 4',     type: 'game'  },
  { name: 'Bet 5',     type: 'prop'  },
  { name: 'Bet 6',     type: 'game'  },
  { name: 'Bet 7',     type: 'game'  },
  { name: 'Bonus Bet', type: 'bonus' },
  { name: 'Local Bet', type: 'local' }
];

// Known participants — used to filter filler rows from bet sheets
const KNOWN_PARTICIPANTS = [
  'evan s.','levi b.','james t.','grant h.','hayden l',
  'hayden l.','ace p.','cole m.','frankie b.','brody s.',
  'aidan b.','seth c.','seth c','layne s.','cole c.','mason l.'
];

// ── Routing ─────────────────────────────────────────────────────
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getAll';
  var out;
  try {
    switch (action) {
      case 'getAll':       out = getAllData();                                              break;
      case 'getStandings': out = getStandings();                                           break;
      case 'getBets':      out = getCurrentBets();                                         break;
      case 'getHistory':   out = getHistory();                                             break;
      case 'getConfig':    out = getConfig();                                              break;
      case 'hasPin':       out = { hasPin: hasPin(e.parameter.participant || '') };        break;
      case 'getPlayoff':   out = getPlayoffData();                                         break;
      default:             out = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    out = { error: err.message, stack: err.stack };
  }
  return jsonOut(out);
}

function doPost(e) {
  var body, out;
  try {
    body = JSON.parse(e.postData.contents);

    // Participant actions — verified by PIN not admin password
    if (body.action === 'placeBet' || body.action === 'cancelBet' || body.action === 'setPin') {
      switch (body.action) {
        case 'placeBet':  out = placeBet(body);             break;
        case 'cancelBet': out = cancelBet(body);            break;
        case 'setPin':    out = setParticipantPin(body);    break;
      }
      return jsonOut(out);
    }

    // init is allowed without a password only when no password is set yet
    if (body.action === 'init') {
      var existingHash = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH');
      if (!existingHash) {
        out = initScript(body);
        return jsonOut(out);
      }
      // If already initialized, fall through to password check
    }

    if (!verifyPassword(body.password)) {
      return jsonOut({ error: 'auth' });
    }

    switch (body.action) {
      case 'ping':             out = { ok: true };                  break;
      case 'init':             out = initScript(body);              break;
      case 'resolveBet':       out = resolveBet(body);              break;
      case 'adjustBalance':    out = adjustBalance(body);           break;
      case 'addBonus':         out = addBonus(body);                break;
      case 'markInactive':     out = markInactive(body);            break;
      case 'reinstate':        out = reinstateParticipant(body);    break;
      case 'addParticipant':   out = addParticipant(body);          break;
      case 'addBet':           out = addBet(body);                  break;
      case 'updateConfig':     out = updateConfig(body);            break;
      case 'setPassword':      out = setPassword(body);             break;
      case 'resetPin':         out = resetParticipantPin(body);     break;
      default:                 out = { error: 'Unknown action: ' + body.action };
    }
  } catch (err) {
    out = { error: err.message, stack: err.stack };
  }
  return jsonOut(out);
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Auth ─────────────────────────────────────────────────────────
function hashPw(pw) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(pw),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0');
  }).join('');
}

function verifyPassword(pw) {
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH');
  return !!stored && hashPw(pw || '') === stored;
}

function setPassword(body) {
  if (!body.newPassword) return { error: 'newPassword required' };
  PropertiesService.getScriptProperties().setProperty('ADMIN_HASH', hashPw(body.newPassword));
  return { ok: true };
}

function initScript(body) {
  if (!body.newPassword) return { error: 'newPassword required' };
  var props = PropertiesService.getScriptProperties();
  // Allow re-init only if called with correct current password (handled in doPost routing)
  props.setProperty('ADMIN_HASH', hashPw(body.newPassword));
  ensureWDSheets();
  return { ok: true, message: 'Initialized successfully' };
}

// ── WD_ Sheet setup ───────────────────────────────────────────────
function getSS() {
  return SpreadsheetApp.openById(SS_ID);
}

function ensureWDSheets() {
  var ss = getSS();
  var sheets = {
    'WD_Config': ['key', 'value'],
    'WD_Results': ['bet_id','participant','week','outcome','amount_change','resolved_at','notes'],
    'WD_Bonuses': ['participant','week','bonus_type','amount','awarded_at','notes'],
    'WD_InactiveLog': ['participant','marked_at','reinstated_at','status','notes'],
    'WD_Bets':       ['bet_id','week','type','title','description','entry_min','bet_max','option_a','option_b','payout_a','payout_b','status','result'],
    'WD_Pins':       ['participant','pin_hash','set_at'],
    'WD_PlayerBets': ['bet_id','participant','week','pick','amount','placed_at','status']
  };

  Object.keys(sheets).forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
    }
  });

  // Seed default config if WD_Config is empty (only header row)
  var cfg = ss.getSheetByName('WD_Config');
  if (cfg && cfg.getLastRow() <= 1) {
    var defaults = [
      ['season_name',      '2024 CFB Season'],
      ['season_end',       '2024-12-12T12:00:00'],
      ['week_lock_day',    '5'],
      ['week_lock_time',   '19:00'],
      ['current_week',     '9'],
      ['start_balance',    '500000'],
      ['min_bets_per_week','2'],
      ['miss_penalty',     '10000'],
      ['inactivity_weeks', '4']
    ];
    defaults.forEach(function(row) { cfg.appendRow(row); });
  }
}

// ── Config ────────────────────────────────────────────────────────
function getConfig() {
  var sheet = getSS().getSheetByName('WD_Config');
  if (!sheet || sheet.getLastRow() < 2) return getDefaultConfig();
  var rows = sheet.getDataRange().getValues().slice(1);
  var cfg = {};
  rows.forEach(function(r) { if (r[0]) cfg[String(r[0])] = formatSheetValue(r[1]); });
  return cfg;
}

function getDefaultConfig() {
  return {
    season_name: '2024 CFB Season',
    season_end: '2024-12-12T12:00:00',
    week_lock_day: '5',
    week_lock_time: '19:00',
    current_week: '9',
    start_balance: '500000',
    min_bets_per_week: '2',
    miss_penalty: '10000',
    inactivity_weeks: '4'
  };
}

function updateConfig(body) {
  ensureWDSheets();
  var sheet = getSS().getSheetByName('WD_Config');
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(body.key)) {
      sheet.getRange(i + 1, 2).setValue(body.value);
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow([body.key, body.value]);
  return { ok: true };
}

// ── Helpers ───────────────────────────────────────────────────────
// Google Sheets auto-converts times like "19:00" into Date objects.
// This extracts them back to "HH:MM" strings.
function formatSheetValue(val) {
  if (val instanceof Date) {
    var h = val.getHours();
    var m = val.getMinutes();
    if (h === 0 && m === 0) return '';
    return String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
  }
  return String(val === null || val === undefined ? '' : val);
}

function parseAmount(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return Math.round(val);
  return parseInt(String(val).replace(/[$,\s]/g, '')) || 0;
}

function isParticipant(name) {
  if (!name || typeof name !== 'string') return false;
  var n = name.trim().toLowerCase();
  if (!n) return false;
  // Check known list first
  if (KNOWN_PARTICIPANTS.indexOf(n) !== -1) return true;
  // Fallback: matches "FirstName L." or "FirstName L" pattern
  return /^[a-z][a-z]+ [a-z]\.?$/.test(n);
}

// ── Participants & Standings ──────────────────────────────────────
function getParticipants() {
  var sheet = getSS().getSheetByName('People and Money');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var participants = [];

  for (var i = 0; i < rows.length; i++) {
    var name = String(rows[i][0] || '').trim();
    if (!isParticipant(name)) continue;
    var balance = parseAmount(rows[i][4]);
    var rank = parseInt(rows[i][2]) || 0;
    // Guard: only include rows that have either a rank or non-zero balance
    if (rank > 0 || balance > 0) {
      participants.push({ name: name, rank: rank, balance: balance, status: 'active' });
    }
  }

  // Mark inactive participants
  var inactiveSheet = getSS().getSheetByName('WD_InactiveLog');
  if (inactiveSheet && inactiveSheet.getLastRow() > 1) {
    var inactiveRows = inactiveSheet.getDataRange().getValues().slice(1);
    // Build a map of latest status per participant
    var statusMap = {};
    inactiveRows.forEach(function(r) {
      var p = String(r[0] || '').trim().toLowerCase();
      if (p) statusMap[p] = String(r[3] || 'inactive');
    });
    participants.forEach(function(p) {
      if (statusMap[p.name.toLowerCase()] === 'inactive') p.status = 'inactive';
    });
  }

  return participants;
}

function getStandings() {
  var participants = getParticipants();
  var history = getHistory();
  var bonuses = getBonuses();
  var config = getConfig();
  var totalWeeks = parseInt(config.current_week) || 1;

  // Accumulate per-participant stats
  var stats = {};
  participants.forEach(function(p) {
    stats[p.name] = { correct: 0, incorrect: 0, total: 0, weeksBet: {} };
  });

  history.forEach(function(r) {
    if (!stats[r.participant]) return;
    stats[r.participant].total++;
    stats[r.participant].weeksBet[r.week] = true;
    if (r.outcome === 'win') stats[r.participant].correct++;
    else if (r.outcome === 'loss') stats[r.participant].incorrect++;
  });

  // Compute Heisman score
  participants.forEach(function(p) {
    var s = stats[p.name] || { correct: 0, incorrect: 0, total: 0, weeksBet: {} };
    var weeksParticipated = Object.keys(s.weeksBet).length;
    var consistency = weeksParticipated / totalWeeks;
    p.correctBets     = s.correct;
    p.incorrectBets   = s.incorrect;
    p.totalBets       = s.total;
    p.weeksParticipated = weeksParticipated;
    p.consistency     = Math.round(consistency * 100);
    // Heisman Score: correct picks * (1 + consistency bonus)
    // Rewards showing up AND picking correctly — labeled as estimate on site
    p.heismanScore    = Math.round(s.correct * (1 + consistency));
    p.accuracy        = s.total > 0 ? Math.round((s.correct / s.total) * 100) : 0;
  });

  // Sort for money race
  participants.sort(function(a, b) { return b.balance - a.balance; });

  // Assign Heisman rank separately
  var heismanSorted = participants.slice().sort(function(a, b) { return b.heismanScore - a.heismanScore; });
  heismanSorted.forEach(function(p, i) { p.heismanRank = i + 1; });

  return participants;
}

// ── Bet Sheet Parsing ─────────────────────────────────────────────
function parseBetSheet(sheetName, betType) {
  var ss = getSS();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;
  var raw = sheet.getDataRange().getValues();
  if (raw.length < 3) return null;

  var title = String(raw[0][0] || '').trim();
  var entryMin = 0, betMax = 0;
  var payoutA = '', payoutB = '', payoutLose = '';

  // Scan first 6 rows for metadata
  for (var r = 0; r < Math.min(6, raw.length); r++) {
    var rowStr = raw[r].map(function(c) { return String(c || ''); }).join('|');

    var enterMatch = rowStr.match(/Bet to enter[:\s\$]+([\d,]+)/i);
    if (enterMatch) entryMin = parseInt(enterMatch[1].replace(/,/g, ''));

    var limitMatch = rowStr.match(/Bet limit[:\s\$]+([\d,]+)/i);
    if (limitMatch) betMax = parseInt(limitMatch[1].replace(/,/g, ''));

    // Look for payout descriptions in individual cells
    raw[r].forEach(function(cell) {
      var c = String(cell || '').trim();
      if (!c || c.toLowerCase().includes('levi') || c.toLowerCase().includes('gay')) return;
      if (c.match(/\+([\$\d]|wins|over|under)/i) && !payoutA) payoutA = c;
      else if (c.match(/\+([\$\d]|wins|over|under)/i) && !payoutB) payoutB = c;
      else if (c.match(/x[\d.]+/i) && !payoutA) payoutA = c;
      else if (c.match(/x[\d.]+/i) && !payoutB) payoutB = c;
      if (c.toLowerCase().includes('loses') && !payoutLose) payoutLose = c;
    });
  }

  // Find header row (Name: column)
  var headerRow = -1;
  for (var r = 0; r < raw.length; r++) {
    if (String(raw[r][0] || '').trim().toLowerCase().startsWith('name')) {
      headerRow = r;
      break;
    }
  }
  if (headerRow === -1) return null;

  // Parse participant bets
  var bets = [];
  for (var r = headerRow + 1; r < raw.length; r++) {
    var name = String(raw[r][0] || '').trim();
    if (!isParticipant(name)) continue;

    var amount = parseAmount(raw[r][1]);
    var pick   = String(raw[r][2] || '').trim();

    if (pick && pick.toLowerCase() !== 'none') {
      bets.push({ participant: name, amount: amount, pick: pick });
    }
  }

  return {
    sheetName: sheetName,
    betType:   betType || 'game',
    title:     title,
    entryMin:  entryMin,
    betMax:    betMax,
    payoutA:   payoutA,
    payoutB:   payoutB,
    payoutLose: payoutLose,
    bets:      bets
  };
}

function getCurrentBets() {
  var result = [];
  BET_SHEETS.forEach(function(s) {
    var parsed = parseBetSheet(s.name, s.type);
    if (parsed) result.push(parsed);
  });
  // Also pull admin-added bets from WD_Bets
  var wdBets = getWDBets();
  result = result.concat(wdBets);
  return result;
}

function getWDBets() {
  var sheet = getSS().getSheetByName('WD_Bets');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getDataRange().getValues().slice(1);
  return rows.filter(function(r) { return !!r[0]; }).map(function(r) {
    return {
      bet_id:   String(r[0]),
      week:     parseInt(r[1]) || 0,
      betType:  String(r[2] || 'game'),
      title:    String(r[3] || ''),
      description: String(r[4] || ''),
      entryMin: parseAmount(r[5]),
      betMax:   parseAmount(r[6]),
      optionA:  String(r[7] || ''),
      optionB:  String(r[8] || ''),
      payoutA:  String(r[9] || ''),
      payoutB:  String(r[10] || ''),
      status:   String(r[11] || 'open'),
      result:   String(r[12] || ''),
      bets:     []
    };
  });
}

// ── History & Bonuses ─────────────────────────────────────────────
function getHistory() {
  var sheet = getSS().getSheetByName('WD_Results');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getDataRange().getValues().slice(1);
  return rows.filter(function(r) { return !!r[0]; }).map(function(r) {
    return {
      bet_id:        String(r[0]),
      participant:   String(r[1]),
      week:          parseInt(r[2]) || 0,
      outcome:       String(r[3]),
      amount_change: parseAmount(r[4]),
      resolved_at:   r[5] ? new Date(r[5]).toISOString() : '',
      notes:         String(r[6] || '')
    };
  });
}

function getBonuses() {
  var sheet = getSS().getSheetByName('WD_Bonuses');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getDataRange().getValues().slice(1);
  return rows.filter(function(r) { return !!r[0]; }).map(function(r) {
    return {
      participant: String(r[0]),
      week:        parseInt(r[1]) || 0,
      bonus_type:  String(r[2] || ''),
      amount:      parseAmount(r[3]),
      awarded_at:  r[4] ? new Date(r[4]).toISOString() : '',
      notes:       String(r[5] || '')
    };
  });
}

// ── Admin: Resolve Bet ────────────────────────────────────────────
function resolveBet(body) {
  // body.results = [{ participant, outcome:'win'|'loss', amount_change }]
  ensureWDSheets();
  var sheet = getSS().getSheetByName('WD_Results');
  var now = new Date().toISOString();

  body.results.forEach(function(r) {
    sheet.appendRow([
      body.bet_id || body.bet_title || 'UNKNOWN',
      r.participant,
      parseInt(body.week) || 0,
      r.outcome,
      r.amount_change,
      now,
      body.notes || ''
    ]);
  });

  // Batch-update balances in People and Money
  updateBalancesFromResults(body.results);
  return { ok: true, resolved: body.results.length };
}

function updateBalancesFromResults(results) {
  var pmSheet = getSS().getSheetByName('People and Money');
  if (!pmSheet) return;
  var rows = pmSheet.getDataRange().getValues();

  results.forEach(function(r) {
    var targetName = String(r.participant || '').trim().toLowerCase();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim().toLowerCase() === targetName) {
        var current = parseAmount(rows[i][4]);
        var updated = current + (parseInt(r.amount_change) || 0);
        pmSheet.getRange(i + 1, 5).setValue(updated);
        rows[i][4] = updated; // update local cache to avoid double-applying
        break;
      }
    }
  });
}

// ── Admin: Balance Adjustment ─────────────────────────────────────
function adjustBalance(body) {
  var pmSheet = getSS().getSheetByName('People and Money');
  if (!pmSheet) return { error: 'People and Money sheet not found' };
  var rows = pmSheet.getDataRange().getValues();
  var target = String(body.participant || '').trim().toLowerCase();

  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim().toLowerCase() === target) {
      var oldBalance = parseAmount(rows[i][4]);
      var newBalance = oldBalance + (parseInt(body.amount) || 0);
      pmSheet.getRange(i + 1, 5).setValue(newBalance);

      // Log it
      ensureWDSheets();
      var logSheet = getSS().getSheetByName('WD_Results');
      if (logSheet) {
        logSheet.appendRow([
          'MANUAL_ADJUST',
          body.participant,
          parseInt(body.week) || 0,
          parseInt(body.amount) >= 0 ? 'adjustment_add' : 'adjustment_deduct',
          parseInt(body.amount) || 0,
          new Date().toISOString(),
          body.reason || 'Manual adjustment'
        ]);
      }
      return { ok: true, oldBalance: oldBalance, newBalance: newBalance };
    }
  }
  return { error: 'Participant not found: ' + body.participant };
}

// ── Admin: Bonuses ────────────────────────────────────────────────
function addBonus(body) {
  ensureWDSheets();
  var sheet = getSS().getSheetByName('WD_Bonuses');
  sheet.appendRow([
    body.participant,
    parseInt(body.week) || 0,
    body.bonus_type || 'custom',
    parseInt(body.amount) || 0,
    new Date().toISOString(),
    body.notes || ''
  ]);
  return adjustBalance({
    participant: body.participant,
    amount: parseInt(body.amount) || 0,
    week: body.week,
    reason: 'Bonus: ' + (body.bonus_type || 'custom')
  });
}

// ── Admin: Inactive Management ────────────────────────────────────
function markInactive(body) {
  ensureWDSheets();
  var logSheet = getSS().getSheetByName('WD_InactiveLog');
  logSheet.appendRow([
    body.participant,
    new Date().toISOString(),
    '',
    'inactive',
    'Week ' + (body.week || '?')
  ]);
  return adjustBalance({
    participant: body.participant,
    amount: -10000,
    week: body.week,
    reason: 'Inactivity penalty'
  });
}

function reinstateParticipant(body) {
  ensureWDSheets();
  var logSheet = getSS().getSheetByName('WD_InactiveLog');
  if (!logSheet || logSheet.getLastRow() < 2) return { error: 'No inactive records' };
  var rows = logSheet.getDataRange().getValues();
  var target = String(body.participant || '').trim().toLowerCase();

  // Find the most recent 'inactive' row for this participant
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0] || '').trim().toLowerCase() === target && rows[i][3] === 'inactive') {
      logSheet.getRange(i + 1, 3).setValue(new Date().toISOString());
      logSheet.getRange(i + 1, 4).setValue('reinstated');
      return { ok: true };
    }
  }
  return { error: 'No inactive record found for ' + body.participant };
}

// ── Admin: Add Participant ────────────────────────────────────────
function addParticipant(body) {
  var pmSheet = getSS().getSheetByName('People and Money');
  if (!pmSheet) return { error: 'People and Money sheet not found' };
  var config = getConfig();
  var startBalance = parseInt(body.startBalance) || parseInt(config.start_balance) || 500000;
  var nextRank = pmSheet.getLastRow(); // approximate rank
  pmSheet.appendRow([body.name, '', nextRank, '', startBalance]);
  return { ok: true, name: body.name, startBalance: startBalance };
}

// ── Admin: Add Bet ────────────────────────────────────────────────
function addBet(body) {
  ensureWDSheets();
  var sheet = getSS().getSheetByName('WD_Bets');
  var bet_id = 'BET_' + Date.now();
  sheet.appendRow([
    bet_id,
    parseInt(body.week) || 0,
    body.type || 'game',
    body.title || '',
    body.description || '',
    parseInt(body.entry_min) || 0,
    parseInt(body.bet_max) || 0,
    body.option_a || '',
    body.option_b || '',
    body.payout_a || '',
    body.payout_b || '',
    'open',
    ''
  ]);
  return { ok: true, bet_id: bet_id };
}

// ── PIN System ────────────────────────────────────────────────────
function hashPin(pin) {
  return hashPw(String(pin || ''));
}

function hasPin(name) {
  var sheet = getSS().getSheetByName('WD_Pins');
  if (!sheet || sheet.getLastRow() < 2) return false;
  var rows = sheet.getDataRange().getValues().slice(1);
  var target = String(name).trim().toLowerCase();
  return rows.some(function(r) {
    return String(r[0] || '').trim().toLowerCase() === target;
  });
}

function verifyParticipantPin(name, pin) {
  var sheet = getSS().getSheetByName('WD_Pins');
  if (!sheet || sheet.getLastRow() < 2) return false;
  var rows = sheet.getDataRange().getValues().slice(1);
  var target = String(name).trim().toLowerCase();
  var hashed = hashPin(pin);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim().toLowerCase() === target) {
      return String(rows[i][1]) === hashed;
    }
  }
  return false;
}

function setParticipantPin(body) {
  if (!body.participant) return { error: 'participant required' };
  if (!body.pin || String(body.pin).length < 4) return { error: 'PIN must be at least 4 digits' };

  ensureWDSheets();
  var sheet = getSS().getSheetByName('WD_Pins');
  var rows = sheet.getDataRange().getValues();
  var target = String(body.participant).trim().toLowerCase();
  var hashed = hashPin(body.pin);

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim().toLowerCase() === target) {
      // Already has PIN — require old PIN to change
      if (!body.oldPin) return { error: 'pin_exists' };
      if (String(rows[i][1]) !== hashPin(body.oldPin)) return { error: 'wrong_old_pin' };
      sheet.getRange(i + 1, 2).setValue(hashed);
      sheet.getRange(i + 1, 3).setValue(new Date().toISOString());
      return { ok: true, updated: true };
    }
  }
  // First time — create PIN
  sheet.appendRow([body.participant, hashed, new Date().toISOString()]);
  return { ok: true, created: true };
}

function resetParticipantPin(body) {
  // Admin only — clears PIN so participant can set a new one
  if (!body.participant) return { error: 'participant required' };
  ensureWDSheets();
  var sheet = getSS().getSheetByName('WD_Pins');
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, message: 'No PIN to reset' };
  var rows = sheet.getDataRange().getValues();
  var target = String(body.participant).trim().toLowerCase();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0] || '').trim().toLowerCase() === target) {
      sheet.deleteRow(i + 1);
      return { ok: true, message: body.participant + ' PIN cleared. They can now set a new one.' };
    }
  }
  return { ok: true, message: 'No PIN found for ' + body.participant };
}

// ── Player Bets (web-placed) ──────────────────────────────────────
function getPlayerBets() {
  var sheet = getSS().getSheetByName('WD_PlayerBets');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return r[0] && r[6] !== 'cancelled'; })
    .map(function(r) {
      return {
        bet_id:      String(r[0]),
        participant: String(r[1]),
        week:        parseInt(r[2]) || 0,
        pick:        String(r[3]),
        amount:      parseAmount(r[4]),
        placed_at:   r[5] ? new Date(r[5]).toISOString() : '',
        status:      String(r[6] || 'open')
      };
    });
}

function placeBet(body) {
  if (!body.participant || !body.pin) return { error: 'participant and PIN required' };
  if (!verifyParticipantPin(body.participant, body.pin)) return { error: 'wrong_pin' };
  if (!body.bet_id || !body.pick)  return { error: 'bet_id and pick required' };

  var amount = parseInt(body.amount) || 0;
  if (amount <= 0) return { error: 'Amount must be positive' };

  // Check participant balance
  var pmSheet = getSS().getSheetByName('People and Money');
  if (!pmSheet) return { error: 'Sheet not found' };
  var pmRows = pmSheet.getDataRange().getValues();
  var target  = String(body.participant).trim().toLowerCase();
  var balance = 0;
  var found   = false;
  for (var i = 0; i < pmRows.length; i++) {
    if (String(pmRows[i][0] || '').trim().toLowerCase() === target) {
      balance = parseAmount(pmRows[i][4]);
      found = true;
      break;
    }
  }
  if (!found)        return { error: 'Participant not found' };
  if (amount > balance) return { error: 'Not enough balance. You have $' + balance.toLocaleString() };
  if (body.entry_min && amount < parseInt(body.entry_min)) {
    return { error: 'Minimum bet is $' + parseInt(body.entry_min).toLocaleString() };
  }
  if (body.bet_max && amount > parseInt(body.bet_max)) {
    return { error: 'Maximum bet is $' + parseInt(body.bet_max).toLocaleString() };
  }

  ensureWDSheets();
  var sheet = getSS().getSheetByName('WD_PlayerBets');

  // Check for duplicate
  if (sheet.getLastRow() > 1) {
    var existing = sheet.getDataRange().getValues().slice(1);
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]) === String(body.bet_id) &&
          String(existing[i][1] || '').trim().toLowerCase() === target &&
          existing[i][6] !== 'cancelled') {
        return { error: 'You already have a bet on this. Cancel it first to change.' };
      }
    }
  }

  sheet.appendRow([body.bet_id, body.participant, parseInt(body.week)||0,
    body.pick, amount, new Date().toISOString(), 'open']);
  return { ok: true, message: 'Bet placed: ' + body.pick + ' for $' + amount.toLocaleString() };
}

function cancelBet(body) {
  if (!body.participant || !body.pin) return { error: 'participant and PIN required' };
  if (!verifyParticipantPin(body.participant, body.pin)) return { error: 'wrong_pin' };

  var sheet = getSS().getSheetByName('WD_PlayerBets');
  if (!sheet || sheet.getLastRow() < 2) return { error: 'No bets found' };
  var rows   = sheet.getDataRange().getValues();
  var target = String(body.participant).trim().toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(body.bet_id) &&
        String(rows[i][1] || '').trim().toLowerCase() === target &&
        rows[i][6] === 'open') {
      sheet.getRange(i + 1, 7).setValue('cancelled');
      return { ok: true };
    }
  }
  return { error: 'Bet not found or already resolved/cancelled' };
}

// ── Playoff & Heisman Extra ───────────────────────────────────────
function getPlayoffData() {
  var sheet = getSS().getSheetByName('Playoff');
  if (!sheet) return null;
  var raw = sheet.getDataRange().getValues();
  var title = String(raw[0] ? raw[0][0] : 'Playoffs').trim();
  var seeds = [];
  raw.forEach(function(row) {
    row.forEach(function(cell) {
      var c = String(cell || '').trim();
      if (!c) return;
      var match = c.match(/^(\d+)\.\s*(.+)/);
      if (match) {
        var seed = parseInt(match[1]);
        var name = match[2].trim().replace(/\./g,'').trim();
        if (seed >= 1 && seed <= 16 && name.length > 1) {
          seeds.push({ seed: seed, name: name });
        }
      }
    });
  });
  seeds.sort(function(a,b) { return a.seed - b.seed; });
  return { title: title, seeds: seeds };
}

function getHeismanBetterData() {
  var sheet = getSS().getSheetByName('Heisman Better');
  if (!sheet) return null;
  var raw = sheet.getDataRange().getValues();
  var playerOfWeek = '';
  var rankings = [];

  raw.forEach(function(row) {
    // Look for "Player of the week:" in any cell
    row.forEach(function(cell) {
      var c = String(cell || '').trim();
      if (c.toLowerCase().includes('player of the week:')) {
        playerOfWeek = c.replace(/player of the week:\s*/i, '').trim();
      }
    });
    // Ranked entries: "1. Evan S.", "2. Levi B." etc.
    var name = String(row[0] || '').trim();
    var rankMatch = name.match(/^(\d+)\.\s+(.+)/);
    if (rankMatch) {
      var delta = row[2] !== '' && row[2] !== null ? String(row[2]).trim() : null;
      rankings.push({ rank: parseInt(rankMatch[1]), name: rankMatch[2].trim(), delta: delta });
    }
  });

  return { playerOfWeek: playerOfWeek, rankings: rankings };
}

// ── Merge web-placed bets into getCurrentBets ─────────────────────
// (patch getCurrentBets to pull from WD_PlayerBets too)
var _origGetCurrentBets = getCurrentBets;
getCurrentBets = function() {
  var result = [];
  BET_SHEETS.forEach(function(s) {
    var parsed = parseBetSheet(s.name, s.type);
    if (parsed) result.push(parsed);
  });
  var wdBets = getWDBets();
  result = result.concat(wdBets);

  // Merge web-placed bets
  var playerBets = getPlayerBets();
  result.forEach(function(bet) {
    var betKey = bet.sheetName || bet.bet_id || '';
    playerBets.forEach(function(pb) {
      if (pb.bet_id !== betKey) return;
      var already = bet.bets.some(function(b) {
        return b.participant.trim().toLowerCase() === pb.participant.trim().toLowerCase();
      });
      if (!already) {
        bet.bets.push({ participant: pb.participant, amount: pb.amount, pick: pb.pick, source: 'web' });
      }
    });
  });
  return result;
};

// ── Aggregated ────────────────────────────────────────────────────
function getAllData() {
  return {
    standings:    getStandings(),
    currentBets:  getCurrentBets(),
    history:      getHistory(),
    bonuses:      getBonuses(),
    config:       getConfig(),
    playoff:      getPlayoffData(),
    heismanExtra: getHeismanBetterData(),
    playerBets:   getPlayerBets(),
    timestamp:    new Date().toISOString()
  };
}
