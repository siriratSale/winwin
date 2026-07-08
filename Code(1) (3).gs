/**
 * Binary Option Personal Goal Journal
 * Google Apps Script Web App + Google Sheets storage
 * UI language: Thai
 * Version: Futuristic Console / compact trade entry
 */

const APP = {
  SPREADSHEET_PROP: 'TRADING_JOURNAL_SPREADSHEET_ID',
  SPREADSHEET_NAME: 'Binary Option Trading Journal - Personal Stats',
  SHEET_TRADES: 'Trades',
  SHEET_SETTINGS: 'Settings',
  TZ: Session.getScriptTimeZone() || 'Asia/Bangkok'
};

const TRADE_HEADERS = [
  'id', 'datetime', 'date', 'time', 'stake', 'result',
  'payoutPct', 'pnl', 'emotion', 'createdAt'
];

const DEFAULT_SETTINGS = {
  initialCapital: '3000',
  currency: 'THB',
  dailyProfitTarget: '2000',
  dailyMaxLoss: '2000',
  weeklyProfitTarget: '10000',
  monthlyProfitTarget: '40000',
  dailyMaxTrades: '20',
  maxConsecutiveLosses: '3',
  stopAfterLosses: '2',
  tradingStartTime: '00:00',
  tradingEndTime: '23:59',
  defaultPayoutPct: '80'
};

const SETTING_LABELS = {
  initialCapital: 'เงินทุนเริ่มต้น',
  currency: 'หน่วยเงิน',
  dailyProfitTarget: 'เป้ากำไรรายวัน',
  dailyMaxLoss: 'ขาดทุนสูงสุดต่อวัน',
  weeklyProfitTarget: 'เป้ากำไรรายสัปดาห์',
  monthlyProfitTarget: 'เป้ากำไรรายเดือน',
  dailyMaxTrades: 'จำนวนไม้สูงสุดต่อวัน',
  maxConsecutiveLosses: 'จำนวนไม้แพ้ติดกันสูงสุด',
  stopAfterLosses: 'ห้ามเทรดหลังแพ้ติดกันกี่ไม้',
  tradingStartTime: 'เวลาเริ่มเทรด',
  tradingEndTime: 'เวลาสิ้นสุดเทรด',
  defaultPayoutPct: 'Payout เริ่มต้น (%)'
};

function doGet() {
  setup();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Binary Option Control Console')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setup() {
  const ss = getOrCreateSpreadsheet_();
  initSheets_(ss);
  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    timezone: APP.TZ
  };
}

function getSheetUrl() {
  const ss = getOrCreateSpreadsheet_();
  return ss.getUrl();
}

function getAppData(filters) {
  const ss = getOrCreateSpreadsheet_();
  initSheets_(ss);
  const settings = getSettings_(ss);
  const trades = getTrades_(ss);
  const history = filterTrades_(trades, filters || {});
  return buildAppPayload_(settings, trades, history, ss.getUrl());
}

function addTrade(trade) {
  const ss = getOrCreateSpreadsheet_();
  initSheets_(ss);
  validateTrade_(trade);

  const settings = getSettings_(ss);
  const payoutDefault = toNumber_(settings.defaultPayoutPct, 80);
  const now = new Date();
  const tradeDate = cleanString_(trade.tradeDate) || formatDateKey_(now);
  const tradeTime = cleanString_(trade.tradeTime) || Utilities.formatDate(now, APP.TZ, 'HH:mm');
  const stake = toNumber_(trade.stake, 0);
  const result = cleanString_(trade.result).toUpperCase();
  const payoutPct = toNumber_(trade.payoutPct, payoutDefault);
  let pnl = toNumber_(trade.pnl, NaN);

  if (isNaN(pnl)) {
    if (result === 'WIN') pnl = roundMoney_(stake * payoutPct / 100);
    else if (result === 'LOSS') pnl = roundMoney_(-stake);
    else pnl = 0;
  }

  const row = [
    Utilities.getUuid(),
    tradeDate + ' ' + tradeTime,
    tradeDate,
    tradeTime,
    stake,
    result,
    payoutPct,
    pnl,
    cleanString_(trade.emotion),
    Utilities.formatDate(now, APP.TZ, 'yyyy-MM-dd HH:mm:ss')
  ];

  ss.getSheetByName(APP.SHEET_TRADES).appendRow(row);
  
  // สำคัญ: บังคับเขียนข้อมูลลง Sheet ทันทีเพื่อป้องกันปัญหา UI ไม่อัปเดต
  SpreadsheetApp.flush();
  
  return getAppData({});
}

function saveSettings(settings) {
  const ss = getOrCreateSpreadsheet_();
  initSheets_(ss);
  const merged = Object.assign({}, DEFAULT_SETTINGS, settings || {});

  merged.initialCapital = String(toNumber_(merged.initialCapital, 3000));
  merged.dailyProfitTarget = String(toNumber_(merged.dailyProfitTarget, 2000));
  merged.dailyMaxLoss = String(toNumber_(merged.dailyMaxLoss, 2000));
  merged.weeklyProfitTarget = String(toNumber_(merged.weeklyProfitTarget, 10000));
  merged.monthlyProfitTarget = String(toNumber_(merged.monthlyProfitTarget, 40000));
  merged.dailyMaxTrades = String(Math.max(0, Math.floor(toNumber_(merged.dailyMaxTrades, 20))));
  merged.maxConsecutiveLosses = String(Math.max(0, Math.floor(toNumber_(merged.maxConsecutiveLosses, 3))));
  merged.stopAfterLosses = String(Math.max(0, Math.floor(toNumber_(merged.stopAfterLosses, 2))));
  merged.defaultPayoutPct = String(toNumber_(merged.defaultPayoutPct, 80));
  merged.currency = cleanString_(merged.currency || 'THB').toUpperCase();
  merged.tradingStartTime = normalizeTime_(merged.tradingStartTime || '00:00');
  merged.tradingEndTime = normalizeTime_(merged.tradingEndTime || '23:59');

  const sheet = ss.getSheetByName(APP.SHEET_SETTINGS);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['key', 'value', 'label']]);
  const rows = Object.keys(DEFAULT_SETTINGS).map(function(key) {
    return [key, merged[key], SETTING_LABELS[key] || key];
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.setFrozenRows(1);

  // สำคัญ: บังคับเขียนข้อมูลลง Sheet ทันที
  SpreadsheetApp.flush();

  return getAppData({});
}

function getHistory(filters) {
  const ss = getOrCreateSpreadsheet_();
  initSheets_(ss);
  const settings = getSettings_(ss);
  const trades = getTrades_(ss);
  const history = filterTrades_(trades, filters || {});
  return buildAppPayload_(settings, trades, history, ss.getUrl());
}

function deleteTrade(tradeId) {
  const ss = getOrCreateSpreadsheet_();
  initSheets_(ss);
  const sheet = ss.getSheetByName(APP.SHEET_TRADES);
  const id = cleanString_(tradeId);
  if (!id || sheet.getLastRow() < 2) return getAppData({});

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      sheet.deleteRow(i + 2);
      break;
    }
  }
  
  // สำคัญ: บังคับลบข้อมูลลง Sheet ทันที
  SpreadsheetApp.flush();
  
  return getAppData({});
}

function resetDemoDataKeepSettings() {
  const ss = getOrCreateSpreadsheet_();
  initSheets_(ss);
  const sheet = ss.getSheetByName(APP.SHEET_TRADES);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, TRADE_HEADERS.length).setValues([TRADE_HEADERS]);
  sheet.setFrozenRows(1);
  
  // สำคัญ: บังคับลบข้อมูลลง Sheet ทันที
  SpreadsheetApp.flush();
  
  return getAppData({});
}

function getOrCreateSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(APP.SPREADSHEET_PROP);
  if (existingId) {
    try {
      return SpreadsheetApp.openById(existingId);
    } catch (err) {
      props.deleteProperty(APP.SPREADSHEET_PROP);
    }
  }
  const ss = SpreadsheetApp.create(APP.SPREADSHEET_NAME);
  props.setProperty(APP.SPREADSHEET_PROP, ss.getId());
  return ss;
}

function initSheets_(ss) {
  let trades = ss.getSheetByName(APP.SHEET_TRADES);
  if (!trades) trades = ss.insertSheet(APP.SHEET_TRADES);
  if (trades.getLastRow() === 0) {
    trades.getRange(1, 1, 1, TRADE_HEADERS.length).setValues([TRADE_HEADERS]);
    trades.setFrozenRows(1);
  } else {
    migrateTradeSheetIfNeeded_(trades);
  }

  let settings = ss.getSheetByName(APP.SHEET_SETTINGS);
  if (!settings) settings = ss.insertSheet(APP.SHEET_SETTINGS);
  if (settings.getLastRow() === 0) {
    settings.getRange(1, 1, 1, 3).setValues([['key', 'value', 'label']]);
    const rows = Object.keys(DEFAULT_SETTINGS).map(function(key) {
      return [key, DEFAULT_SETTINGS[key], SETTING_LABELS[key] || key];
    });
    settings.getRange(2, 1, rows.length, 3).setValues(rows);
    settings.setFrozenRows(1);
  }

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 2 && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
  }
}

function migrateTradeSheetIfNeeded_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), TRADE_HEADERS.length);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return cleanString_(h);
  });
  const current = TRADE_HEADERS.join('|');
  const existing = header.slice(0, TRADE_HEADERS.length).join('|');
  const hasRemovedColumns = header.indexOf('asset') !== -1 || header.indexOf('direction') !== -1 || header.indexOf('strategy') !== -1 || header.indexOf('note') !== -1;
  if (existing === current && !hasRemovedColumns) return;

  const lastRow = sheet.getLastRow();
  const allValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const index = {};
  header.forEach(function(name, i) {
    if (name) index[name] = i;
  });

  const rows = allValues.map(function(row) {
    const obj = {};
    Object.keys(index).forEach(function(key) { obj[key] = row[index[key]]; });
    return [
      cleanString_(obj.id) || Utilities.getUuid(),
      cleanString_(obj.datetime) || (cleanString_(obj.date) + ' ' + cleanString_(obj.time)).trim(),
      cleanString_(obj.date),
      cleanString_(obj.time),
      toNumber_(obj.stake, 0),
      cleanString_(obj.result).toUpperCase(),
      toNumber_(obj.payoutPct, 80),
      toNumber_(obj.pnl, 0),
      cleanString_(obj.emotion),
      cleanString_(obj.createdAt)
    ];
  }).filter(function(row) {
    return row[0] && row[5];
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, TRADE_HEADERS.length).setValues([TRADE_HEADERS]);
  if (rows.length) sheet.getRange(2, 1, rows.length, TRADE_HEADERS.length).setValues(rows);
  sheet.setFrozenRows(1);
}

function getSettings_(ss) {
  const sheet = ss.getSheetByName(APP.SHEET_SETTINGS);
  const settings = Object.assign({}, DEFAULT_SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return settings;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  values.forEach(function(row) {
    const key = cleanString_(row[0]);
    if (key) settings[key] = cleanString_(row[1]);
  });
  return settings;
}

function getTrades_(ss) {
  const sheet = ss.getSheetByName(APP.SHEET_TRADES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return cleanString_(h);
  });
  const index = {};
  header.forEach(function(name, i) {
    if (name) index[name] = i;
  });
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  const trades = values.map(function(row) {
    const get = function(key) {
      return index[key] === undefined ? '' : row[index[key]];
    };
    
    // 👉 DEV FIX: ดักจับและบังคับแปลง Date Object จาก Google Sheets กลับเป็น String
    let dateVal = get('date');
    let dateStr = '';
    if (Object.prototype.toString.call(dateVal) === '[object Date]') {
      dateStr = Utilities.formatDate(dateVal, APP.TZ, 'yyyy-MM-dd');
    } else {
      dateStr = cleanString_(dateVal).substring(0, 10);
    }
    
    let timeVal = get('time');
    let timeStr = '';
    if (Object.prototype.toString.call(timeVal) === '[object Date]') {
      timeStr = Utilities.formatDate(timeVal, APP.TZ, 'HH:mm');
    } else {
      timeStr = cleanString_(timeVal).substring(0, 5);
    }

    let dtVal = get('datetime');
    let dtStr = '';
    if (Object.prototype.toString.call(dtVal) === '[object Date]') {
      dtStr = Utilities.formatDate(dtVal, APP.TZ, 'yyyy-MM-dd HH:mm');
    } else {
      dtStr = cleanString_(dtVal);
    }

    return {
      id: String(get('id') || ''),
      datetime: dtStr,
      date: dateStr,
      time: timeStr,
      stake: toNumber_(get('stake'), 0),
      result: String(get('result') || '').toUpperCase(),
      payoutPct: toNumber_(get('payoutPct'), 0),
      pnl: toNumber_(get('pnl'), 0),
      emotion: String(get('emotion') || ''),
      createdAt: String(get('createdAt') || '')
    };
  }).filter(function(t) { return t.id; });

  trades.sort(function(a, b) {
    return String(b.datetime).localeCompare(String(a.datetime));
  });
  return trades;
}

function buildAppPayload_(settings, allTrades, historyTrades, spreadsheetUrl) {
  const today = formatDateKey_(new Date());
  const weekStart = getWeekStartKey_();
  const monthStart = today.substring(0, 8) + '01';

  const todayTrades = allTrades.filter(function(t) { return t.date === today; });
  const weekTrades = allTrades.filter(function(t) { return t.date >= weekStart && t.date <= today; });
  const monthTrades = allTrades.filter(function(t) { return t.date >= monthStart && t.date <= today; });

  const stats = {
    today: buildStats_(todayTrades),
    week: buildStats_(weekTrades),
    month: buildStats_(monthTrades),
    all: buildStats_(allTrades),
    history: buildStats_(historyTrades || allTrades)
  };

  const initialCapital = toNumber_(settings.initialCapital, 3000);
  const dailyProfitTarget = toNumber_(settings.dailyProfitTarget, 2000);
  const dailyMaxLoss = toNumber_(settings.dailyMaxLoss, 2000);

  // DEV ADD: ส่งยอดเป้าคงเหลือไปให้หน้าเว็บ เพื่อช่วยคำนวณราคาลงไม้ถัดไป
  stats.today.dailyProfitTarget = dailyProfitTarget;
  stats.today.dailyMaxLoss = dailyMaxLoss;
  stats.today.targetRemaining = roundMoney_(Math.max(dailyProfitTarget - stats.today.pnl, 0));

  stats.currentCapital = roundMoney_(initialCapital + stats.all.pnl);
  stats.equityCurve = buildEquityCurve_(allTrades, initialCapital);
  stats.alerts = buildAlerts_(settings, todayTrades, stats.today);
  stats.best = buildBestInsights_(allTrades);

  return {
    ok: true,
    timezone: APP.TZ,
    spreadsheetUrl: spreadsheetUrl,
    settings: settings,
    stats: stats,
    recentTrades: allTrades.slice(0, 12),
    historyTrades: (historyTrades || allTrades).slice(0, 300),
    options: {
      emotions: ['มั่นใจ', 'นิ่ง', 'กลัว', 'รีบ', 'แก้มือ', 'ลังเล', 'โลภ', 'เหนื่อย'],
      results: ['WIN', 'LOSS']
    }
  };
}

function buildStats_(trades) {
  const chronological = trades.slice().sort(function(a, b) {
    return String(a.datetime).localeCompare(String(b.datetime));
  });
  const total = trades.length;
  const wins = trades.filter(function(t) { return t.result === 'WIN'; }).length;
  const losses = trades.filter(function(t) { return t.result === 'LOSS'; }).length;
  const pnl = roundMoney_(trades.reduce(function(sum, t) { return sum + toNumber_(t.pnl, 0); }, 0));
  const stake = roundMoney_(trades.reduce(function(sum, t) { return sum + toNumber_(t.stake, 0); }, 0));
  const dates = unique_(trades.map(function(t) { return t.date; }).filter(Boolean));

  return {
    totalTrades: total,
    wins: wins,
    losses: losses,
    winRate: total ? roundMoney_(wins / total * 100) : 0,
    pnl: pnl,
    stake: stake,
    avgPnlPerTrade: total ? roundMoney_(pnl / total) : 0,
    avgPnlPerTradingDay: dates.length ? roundMoney_(pnl / dates.length) : 0,
    maxConsecutiveWins: maxConsecutive_(chronological, 'WIN'),
    maxConsecutiveLosses: maxConsecutive_(chronological, 'LOSS'),
    currentConsecutiveLosses: currentConsecutive_(chronological, 'LOSS'),
    currentConsecutiveWins: currentConsecutive_(chronological, 'WIN')
  };
}

function buildBestInsights_(trades) {
  return {
    bestHour: bestByPnl_(trades.map(function(t) {
      const copy = Object.assign({}, t);
      copy.hour = t.time ? t.time.substring(0, 2) + ':00' : 'ไม่ระบุ';
      return copy;
    }), 'hour'),
    worstEmotion: worstByPnl_(trades, 'emotion'),
    mostLossEmotion: mostLossByCount_(trades, 'emotion')
  };
}

function buildEquityCurve_(trades, initialCapital) {
  let balance = toNumber_(initialCapital, 3000);
  const chronological = trades.slice().sort(function(a, b) {
    return String(a.datetime).localeCompare(String(b.datetime));
  });
  const curve = [{ label: 'เริ่มต้น', balance: roundMoney_(balance) }];
  chronological.forEach(function(t) {
    balance += toNumber_(t.pnl, 0);
    curve.push({
      label: t.date + ' ' + t.time,
      balance: roundMoney_(balance)
    });
  });
  return curve.slice(-140);
}

function buildAlerts_(settings, todayTrades, todayStats) {
  const alerts = [];
  const dailyProfitTarget = toNumber_(settings.dailyProfitTarget, 2000);
  const dailyMaxLoss = toNumber_(settings.dailyMaxLoss, 2000);
  const dailyMaxTrades = Math.floor(toNumber_(settings.dailyMaxTrades, 0));
  const maxConsecutiveLosses = Math.floor(toNumber_(settings.maxConsecutiveLosses, 0));
  const stopAfterLosses = Math.floor(toNumber_(settings.stopAfterLosses, 0));

  if (dailyProfitTarget > 0 && todayStats.pnl >= dailyProfitTarget) {
    alerts.push({ level: 'success', text: 'หยุดเทรด วันนี้ถึงเป้ากำไรแล้ว' });
  }
  if (dailyMaxLoss > 0 && todayStats.pnl <= -dailyMaxLoss) {
    alerts.push({ level: 'danger', text: 'หยุดเทรด วันนี้แตะขาดทุนสูงสุดแล้ว' });
  } else if (dailyMaxLoss > 0 && todayStats.pnl <= -(dailyMaxLoss * 0.8)) {
    alerts.push({ level: 'warning', text: 'ใกล้แตะขาดทุนสูงสุด ระวังอย่าแก้มือ' });
  }
  if (dailyMaxTrades > 0 && todayStats.totalTrades >= dailyMaxTrades) {
    alerts.push({ level: 'warning', text: 'วันนี้เทรดครบจำนวนไม้สูงสุดแล้ว' });
  }
  if (maxConsecutiveLosses > 0 && todayStats.currentConsecutiveLosses >= maxConsecutiveLosses) {
    alerts.push({ level: 'danger', text: 'แพ้ติดกันเกินกำหนด ควรหยุดเทรดทันที' });
  } else if (stopAfterLosses > 0 && todayStats.currentConsecutiveLosses >= stopAfterLosses) {
    alerts.push({ level: 'danger', text: 'เข้ากฎห้ามเทรดหลังแพ้ติดกัน ควรหยุดก่อน' });
  }
  if (!isWithinTradingTime_(settings.tradingStartTime, settings.tradingEndTime)) {
    alerts.push({ level: 'warning', text: 'ตอนนี้อยู่นอกเวลาที่ตั้งไว้สำหรับเทรด' });
  }
  if (!alerts.length) {
    alerts.push({ level: 'info', text: 'ยังอยู่ในกรอบวินัย ตรวจอารมณ์ก่อนเข้าไม้ถัดไป' });
  }
  return alerts;
}

function filterTrades_(trades, filters) {
  filters = filters || {};
  const dateFrom = cleanString_(filters.dateFrom);
  const dateTo = cleanString_(filters.dateTo);
  const result = cleanString_(filters.result).toUpperCase();
  const emotion = cleanString_(filters.emotion).toLowerCase();

  return trades.filter(function(t) {
    if (dateFrom && t.date < dateFrom) return false;
    if (dateTo && t.date > dateTo) return false;
    if (result && t.result !== result) return false;
    if (emotion && String(t.emotion).toLowerCase().indexOf(emotion) === -1) return false;
    return true;
  });
}

function validateTrade_(trade) {
  if (!trade) throw new Error('ไม่พบข้อมูลออเดอร์');
  const stake = toNumber_(trade.stake, 0);
  if (stake <= 0) throw new Error('กรุณาใส่จำนวนเงินต่อไม้มากกว่า 0');
  const result = cleanString_(trade.result).toUpperCase();
  if (['WIN', 'LOSS'].indexOf(result) === -1) throw new Error('ผลลัพธ์ต้องเป็น WIN หรือ LOSS');
}

function bestByPnl_(trades, key) {
  const map = groupByKey_(trades, key);
  let best = null;
  Object.keys(map).forEach(function(name) {
    const list = map[name];
    const pnl = roundMoney_(list.reduce(function(sum, t) { return sum + toNumber_(t.pnl, 0); }, 0));
    const wins = list.filter(function(t) { return t.result === 'WIN'; }).length;
    const total = list.length;
    const item = { name: name, pnl: pnl, trades: total, winRate: total ? roundMoney_(wins / total * 100) : 0 };
    if (!best || item.pnl > best.pnl) best = item;
  });
  return best || { name: '-', pnl: 0, trades: 0, winRate: 0 };
}

function worstByPnl_(trades, key) {
  const map = groupByKey_(trades, key);
  let worst = null;
  Object.keys(map).forEach(function(name) {
    const list = map[name];
    const pnl = roundMoney_(list.reduce(function(sum, t) { return sum + toNumber_(t.pnl, 0); }, 0));
    const losses = list.filter(function(t) { return t.result === 'LOSS'; }).length;
    const item = { name: name, pnl: pnl, trades: list.length, losses: losses };
    if (!worst || item.pnl < worst.pnl) worst = item;
  });
  return worst || { name: '-', pnl: 0, trades: 0, losses: 0 };
}

function mostLossByCount_(trades, key) {
  const losses = trades.filter(function(t) { return t.result === 'LOSS'; });
  const map = groupByKey_(losses, key);
  let worst = null;
  Object.keys(map).forEach(function(name) {
    const item = { name: name, losses: map[name].length };
    if (!worst || item.losses > worst.losses) worst = item;
  });
  return worst || { name: '-', losses: 0 };
}

function groupByKey_(trades, key) {
  const map = {};
  trades.forEach(function(t) {
    const name = cleanString_(t[key]) || 'ไม่ระบุ';
    if (!map[name]) map[name] = [];
    map[name].push(t);
  });
  return map;
}

function maxConsecutive_(trades, result) {
  let max = 0;
  let current = 0;
  trades.forEach(function(t) {
    if (t.result === result) {
      current += 1;
      if (current > max) max = current;
    } else {
      current = 0;
    }
  });
  return max;
}

function currentConsecutive_(trades, result) {
  let count = 0;
  for (let i = trades.length - 1; i >= 0; i--) {
    if (trades[i].result === result) count += 1;
    else break;
  }
  return count;
}

function isWithinTradingTime_(start, end) {
  const now = Utilities.formatDate(new Date(), APP.TZ, 'HH:mm');
  start = normalizeTime_(start || '00:00');
  end = normalizeTime_(end || '23:59');
  if (start <= end) return now >= start && now <= end;
  return now >= start || now <= end;
}

function getWeekStartKey_() {
  const now = new Date();
  const day = Number(Utilities.formatDate(now, APP.TZ, 'u')); // 1 Monday - 7 Sunday
  const start = new Date(now.getTime());
  start.setDate(now.getDate() - day + 1);
  return formatDateKey_(start);
}

function formatDateKey_(date) {
  return Utilities.formatDate(date, APP.TZ, 'yyyy-MM-dd');
}

function normalizeTime_(value) {
  value = cleanString_(value);
  const match = value.match(/^(\d{1,2}):(\d{1,2})/);
  if (!match) return '00:00';
  let hh = Math.max(0, Math.min(23, Number(match[1])));
  let mm = Math.max(0, Math.min(59, Number(match[2])));
  return String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0');
}

function toNumber_(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback;
  const n = Number(String(value).replace(/,/g, ''));
  return isNaN(n) ? fallback : n;
}

function roundMoney_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function cleanString_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function unique_(arr) {
  const out = [];
  const seen = {};
  arr.forEach(function(item) {
    if (!seen[item]) {
      seen[item] = true;
      out.push(item);
    }
  });
  return out;
}
