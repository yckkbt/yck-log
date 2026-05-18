// YCK授業ログ + スケジュール - Google Apps Script (コード.gs)

const SHEET_LOGS     = '授業ログ';
const SHEET_MASTER   = 'マスタ';
const SHEET_CONFIG   = '設定';
const SHEET_SHIFTS   = 'シフト';
const SHEET_EVENTS   = '行事';
const SHEET_NOTICES  = 'お知らせ';

function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  const body = (() => { try { return JSON.parse(e.postData?.contents || '{}'); } catch { return {}; } })();
  const action = (e.parameter||{}).action || body.action;
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    let result;
    switch (action) {
      // 既存
      case 'auth':        result = auth(body.password);                break;
      case 'getLogs':     result = getLogs();                          break;
      case 'addLog':      result = addLog(body.log);                   break;
      case 'getMaster':   result = getMaster();                        break;
      case 'setMaster':   result = setMaster(body.master);             break;
      case 'changePass':  result = changePass(body.oldPw, body.newPw); break;
      case 'dedupMaster': result = dedupMaster();                      break;
      // スケジュール
      case 'getShifts':   result = getShifts();                        break;
      case 'addShift':    result = addShift(body.shift);               break;
      case 'updateShift': result = updateShift(body.shift);            break;
      case 'deleteShift': result = deleteShift(body.id);               break;
      case 'getEvents':   result = getEvents();                        break;
      case 'addEvent':    result = addEvent(body.event);               break;
      case 'deleteEvent': result = deleteEvent(body.id);               break;
      case 'getNotices':  result = getNotices();                       break;
      case 'setNotice':   result = setNotice(body.content);            break;
      default:            result = { ok: false, error: 'unknown action' };
    }
    out.setContent(JSON.stringify(result));
  } catch(err) {
    out.setContent(JSON.stringify({ ok: false, error: err.toString() }));
  }
  return out;
}

// ── 認証 ──
function auth(pw) {
  return { ok: pw === getConfig().getRange(1,2).getValue() };
}

// ── ログ ──
function getLogs() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOGS);
  if (!sheet) return { ok: true, logs: [] };
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, logs: [] };
  const rawHeaders = rows[0].map(h => String(h).trim());
  const HEADER_MAP = {
    '日付':'date','生徒':'student','講師':'teacher','科目':'subject',
    '前計':'preTotal','後計':'postTotal','タイプ':'postType','代講':'daiko','メモ':'memo',
    '授業内容':'content',
    'id':'id','date':'date','student':'student','teacher':'teacher',
    'subject':'subject','content':'content','daiko':'daiko',
    'preU':'preU','preO':'preO','preS':'preS','preTotal':'preTotal','preType':'preType',
    'postU':'postU','postO':'postO','postS':'postS','postTotal':'postTotal','postType':'postType',
    'scoreUp':'scoreUp','memo':'memo'
  };
  const headers = rawHeaders.map(h => HEADER_MAP[h] || h);
  const logs = rows.slice(1).map(r => {
    const o = {};
    headers.forEach((h,i) => { o[h] = r[i]; });
    if (!('daiko' in o)) o.daiko = '';
    o.preTotal  = isNaN(Number(o.preTotal))  ? 0 : Number(o.preTotal);
    o.postTotal = isNaN(Number(o.postTotal)) ? 0 : Number(o.postTotal);
    if (!o.postType || String(o.postType).trim() === '') {
      const t = o.postTotal;
      o.postType = t>=9?'加速型':t>=7?'育成型':t>=4?'分解型':'危険型';
    }
    if (!o.preType || String(o.preType).trim() === '') {
      const t = o.preTotal;
      o.preType = t>=9?'加速型':t>=7?'育成型':t>=4?'分解型':'危険型';
    }
    if (!o.daiko)   o.daiko   = '';
    if (!o.memo)    o.memo    = '';
    if (!o.subject) o.subject = '';
    if (!o.date)    o.date    = '';
    return o;
  });
  return { ok: true, logs };
}

function addLog(log) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOGS);
    sheet.appendRow(['日付','生徒','講師','科目','前計','後計','タイプ','代講','メモ','授業内容','preU','preO','preS','postU','postO','postS']);
  }
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h=>String(h).trim());
  const HEADER_MAP = {
    '日付':'date','生徒':'student','講師':'teacher','科目':'subject',
    '前計':'preTotal','後計':'postTotal','タイプ':'postType','代講':'daiko','メモ':'memo',
    '授業内容':'content',
    'id':'id','date':'date','student':'student','teacher':'teacher',
    'subject':'subject','content':'content','daiko':'daiko',
    'preU':'preU','preO':'preO','preS':'preS','preTotal':'preTotal','preType':'preType',
    'postU':'postU','postO':'postO','postS':'postS','postTotal':'postTotal','postType':'postType',
    'scoreUp':'scoreUp','memo':'memo'
  };
  const logData = {
    id:log.id,date:log.date,student:log.student,teacher:log.teacher,
    subject:log.subject||'',content:log.content||'',daiko:log.daiko||'',
    preU:log.preU,preO:log.preO,preS:log.preS,preTotal:log.preTotal,preType:log.preType,
    postU:log.postU,postO:log.postO,postS:log.postS,postTotal:log.postTotal,postType:log.postType,
    scoreUp:Number(log.postTotal)-Number(log.preTotal),memo:log.memo||''
  };
  const row = headers.map(h => {
    const key = HEADER_MAP[h] || h;
    return logData[key] !== undefined ? logData[key] : '';
  });
  sheet.appendRow(row);
  return { ok: true };
}

// ── マスタ ──
function getMaster() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreate(ss, SHEET_MASTER, ['type','value','extra1','extra2']);
  const rows  = sheet.getDataRange().getValues().slice(1);
  const master = { students:[], teachers:[], subjects:[] };
  rows.forEach(r => {
    const [type,value,ex1,ex2] = r;
    if (type==='student') master.students.push({
      name:value,
      teachers:ex1?String(ex1).split('|').filter(Boolean):[],
      subjects:ex2?String(ex2).split('|').filter(Boolean):[]
    });
    if (type==='teacher') master.teachers.push(value);
    if (type==='subject') master.subjects.push(value);
  });
  if (!master.teachers.length) master.teachers=['中村 先生','渡辺 先生','小林 先生'];
  if (!master.subjects.length) master.subjects=['数学','英語','国語','理科','社会'];
  return { ok:true, master };
}

function setMaster(master) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreate(ss, SHEET_MASTER, ['type','value','extra1','extra2']);
  sheet.clearContents();
  sheet.appendRow(['type','value','extra1','extra2']);
  (master.students||[]).forEach(s => {
    const name = typeof s==='string'?s:s.name;
    sheet.appendRow(['student',name,(s.teachers||[]).join('|'),(s.subjects||[]).join('|')]);
  });
  (master.teachers||[]).forEach(v=>sheet.appendRow(['teacher',v,'','']));
  (master.subjects||[]).forEach(v=>sheet.appendRow(['subject',v,'','']));
  return { ok:true };
}

function changePass(oldPw,newPw) {
  const config=getConfig();
  if (oldPw!==config.getRange(1,2).getValue()) return {ok:false,error:'現在のパスワードが違います'};
  config.getRange(1,2).setValue(newPw);
  return {ok:true};
}

function dedupMaster() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(SHEET_MASTER);
  if (!sheet) return {ok:true,removed:0};
  const rows=sheet.getDataRange().getValues();
  if (rows.length<=1) return {ok:true,removed:0};
  const header=rows[0];
  const seen=new Set();
  const unique=[header];
  rows.slice(1).forEach(r=>{
    const key=r[0]+'|'+r[1];
    if (!seen.has(key)){seen.add(key);unique.push(r);}
  });
  sheet.clearContents();
  unique.forEach((r,i)=>sheet.getRange(i+1,1,1,r.length).setValues([r]));
  return {ok:true,removed:rows.length-unique.length};
}

// ── シフト ──
function getShifts() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreate(ss, SHEET_SHIFTS, ['id','teacher','date','startTime','slots','students','note']);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length<=1) return {ok:true,shifts:[]};
  const headers=rows[0];
  return {ok:true,shifts:rows.slice(1).map(r=>{
    const o={};headers.forEach((h,i)=>o[h]=r[i]);return o;
  })};
}

function addShift(shift) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreate(ss, SHEET_SHIFTS, ['id','teacher','date','startTime','slots','students','note']);
  sheet.appendRow([shift.id,shift.teacher,shift.date,shift.startTime,shift.slots,shift.students||'',shift.note||'']);
  return {ok:true};
}

function updateShift(shift) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SHIFTS);
  if (!sheet) return {ok:false,error:'sheet not found'};
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][0])===String(shift.id)) {
      sheet.getRange(i+1,1,1,7).setValues([[shift.id,shift.teacher,shift.date,shift.startTime,shift.slots,shift.students||'',shift.note||'']]);
      return {ok:true};
    }
  }
  return {ok:false,error:'not found'};
}

function deleteShift(id) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SHIFTS);
  if (!sheet) return {ok:true};
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][0])===String(id)) {sheet.deleteRow(i+1);return {ok:true};}
  }
  return {ok:true};
}

// ── 行事 ──
function getEvents() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreate(ss, SHEET_EVENTS, ['id','startDate','endDate','type','content','school']);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length<=1) return {ok:true,events:[]};
  const headers=rows[0];
  return {ok:true,events:rows.slice(1).map(r=>{
    const o={};headers.forEach((h,i)=>o[h]=r[i]);return o;
  })};
}

function addEvent(ev) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreate(ss, SHEET_EVENTS, ['id','startDate','endDate','type','content','school']);
  sheet.appendRow([ev.id,ev.startDate,ev.endDate||ev.startDate,ev.type,ev.content,ev.school||'']);
  return {ok:true};
}

function deleteEvent(id) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_EVENTS);
  if (!sheet) return {ok:true};
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][0])===String(id)) {sheet.deleteRow(i+1);return {ok:true};}
  }
  return {ok:true};
}

// ── お知らせ ──
function getNotices() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreate(ss, SHEET_NOTICES, ['content']);
  const rows  = sheet.getDataRange().getValues().slice(1);
  return {ok:true,notices:rows.map(r=>({content:String(r[0]||'')})).filter(n=>n.content)};
}

function setNotice(content) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreate(ss, SHEET_NOTICES, ['content']);
  sheet.clearContents();
  sheet.appendRow(['content']);
  if (content && content.trim()) sheet.appendRow([content]);
  return {ok:true};
}

// ── ユーティリティ ──
function getConfig() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const s=ss.getSheetByName(SHEET_CONFIG)||ss.insertSheet(SHEET_CONFIG);
  if (s.getLastRow()===0) s.appendRow(['password','yck2024']);
  return s;
}

function getOrCreate(ss,name,headers) {
  let s=ss.getSheetByName(name);
  if (!s){s=ss.insertSheet(name);s.appendRow(headers);}
  return s;
}
