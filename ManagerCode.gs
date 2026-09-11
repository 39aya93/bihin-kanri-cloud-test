/**
 * 備品管理：DB管理用 Apps Script（1つで複数DB管理方式・Apps Script API不使用）
 */
const MANAGER_SHEET_NAME = 'DB管理台帳';
const DEFAULT_LOCATIONS = ['工具棚A','工具棚B','ガレージ','物置','パントリー','倉庫','車庫'];
const DEFAULT_STORES = ['カインズ','コーナン','Amazon','楽天市場','モノタロウ','その他'];
const ITEMS_SHEET = 'Items';
const SETTINGS_SHEET = 'Settings';
const ITEM_HEADERS = ['id','name','spec','quantity','image','location','obtained','expiry','store','price','memo','createdAt','updatedAt'];

function assertManagerUser_() {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  return email;
}

function setupManager() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('REGISTRY_SHEET_ID');
  if (!id) {
    const ss = SpreadsheetApp.create(MANAGER_SHEET_NAME);
    id = ss.getId();
    props.setProperty('REGISTRY_SHEET_ID', id);
  }
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheets()[0];
  sh.setName(MANAGER_SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,7).setValues([['DB名','Spreadsheet ID','Spreadsheet URL','Script ID','WebアプリURL','編集者','作成日時']]);
    sh.setFrozenRows(1);
  }
  return id;
}

function doGet(e) {
  assertManagerUser_();
  const p = e && e.parameter ? e.parameter : {};
  const op = p.op || 'list';
  let result;
  try {
    if (op === 'list') result = {ok:true, databases:listDatabases_()};
    else if (op === 'all') {
      const ssId = String(p.spreadsheetId || '').trim();
      if (!ssId) throw new Error('DBが指定されていません。');
      result = getAll_(ssId);
    } else if (op === 'health') result = {ok:true, service:'BihinKanri DB Manager',time:new Date().toISOString()};
    else throw new Error('不明な操作です: '+op);
  } catch (err) {
    result = {ok:false,error:String(err && err.message ? err.message : err)};
  }
  return output_(result,p.prefix);
}

function doPost(e) {
  let result;
  try {
    let payload = {};
    if (e.postData && e.postData.contents) {
      const ct = String(e.postData.type || '').toLowerCase();
      if (ct.includes('json')) {
        payload = JSON.parse(e.postData.contents);
      } else {
        const params = new URLSearchParams(e.postData.contents);
        const payloadStr = params.get('payload');
        if (payloadStr) payload = JSON.parse(payloadStr);
      }
    }
    const action = payload.action || '';
    const ssId = String(payload.spreadsheetId || '').trim();

    if (action === 'createDb') {
      const name = String(payload.name || '').trim();
      const editors = Array.isArray(payload.editors) ? payload.editors : [];
      result = createDatabase_(name, editors);
    } else if (!ssId) {
      throw new Error('DBが指定されていません。');
    } else if (action === 'saveItem') result = saveItem_(ssId, payload.item);
    else if (action === 'deleteItem') result = deleteItem_(ssId, String(payload.id));
    else if (action === 'saveSettings') result = saveSettings_(ssId, payload.settings || {});
    else if (action === 'saveAll') result = saveAll_(ssId, payload);
    else throw new Error('不明な書き込み操作です: '+action);
  } catch(err) {
    result = {ok:false,error:String(err && err.message ? err.message : err)};
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doOptions(e) {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

function listDatabases_() {
  setupManager();
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('REGISTRY_SHEET_ID'));
  const sh = ss.getSheetByName(MANAGER_SHEET_NAME);
  const values = sh.getDataRange().getValues();
  return values.slice(1).filter(r=>r[0] && r[1]).map(r=>({
    name:String(r[0]), spreadsheetId:String(r[1]), spreadsheetUrl:String(r[2]||('https://docs.google.com/spreadsheets/d/'+r[1]+'/edit')),
    scriptId:String(r[3]||''), webAppUrl:String(r[4]||''), editors:String(r[5]||''), createdAt:r[6] ? new Date(r[6]).toISOString() : ''
  }));
}

function createDatabase_(name, editors) {
  setupManager();
  const existing = listDatabases_();
  if (existing.some(x=>x.name === name)) throw new Error('同名のDBがすでにあります。別のDB名にしてください。');

  const ss = SpreadsheetApp.create(name);
  const spreadsheetId = ss.getId();
  initializeDatabase_(ss);

  editors.forEach(email=>{
    try { DriveApp.getFileById(spreadsheetId).addEditor(email); } catch(err) {}
  });

  const reg = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('REGISTRY_SHEET_ID')).getSheetByName(MANAGER_SHEET_NAME);
  reg.appendRow([name,spreadsheetId,ss.getUrl(),'','',editors.join(','),new Date()]);

  return {
    name, spreadsheetId, spreadsheetUrl:ss.getUrl(),
    scriptId:'', webAppUrl:'', editors
  };
}

function getAll_(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sh = ss.getSheetByName(ITEMS_SHEET);
  const settings = readSettings_(ss);
  if (!sh || sh.getLastRow() < 2) return {ok:true,dbName:ss.getName(),spreadsheetId:ss.getId(),settings:settings,items:[]};
  const values = sh.getRange(2,1,sh.getLastRow()-1,ITEM_HEADERS.length).getValues();
  const items = values.filter(r=>r[0] !== '').map(rowToItem_);
  return {ok:true,dbName:ss.getName(),spreadsheetId:ss.getId(),settings:settings,items:items};
}

function rowToItem_(r) {
  const o = {};
  ITEM_HEADERS.forEach((h,i)=>o[h]=r[i] === null || r[i] === undefined ? '' : r[i]);
  if (o.quantity !== '' && o.quantity !== null) o.quantity = Number(o.quantity);
  if (o.price !== '' && o.price !== null) o.price = Number(o.price);
  return o;
}

function itemToRow_(item) {
  return ITEM_HEADERS.map(h=>item && item[h] !== undefined ? item[h] : '');
}

function saveItem_(spreadsheetId, item) {
  if (!item || !item.id) throw new Error('備品IDがありません。');
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const sh = getItemsSheet_(spreadsheetId);
    const id = String(item.id);
    const last = sh.getLastRow();
    if (last >= 2) {
      const ids = sh.getRange(2,1,last-1,1).getValues();
      for (let i=0;i<ids.length;i++) {
        if (String(ids[i][0]) === id) {
          sh.getRange(i+2,1,1,ITEM_HEADERS.length).setValues([itemToRow_(item)]);
          return {ok:true,updated:true,id:id};
        }
      }
    }
    sh.appendRow(itemToRow_(item));
    return {ok:true,updated:false,id:id};
  } finally { lock.releaseLock(); }
}

function deleteItem_(spreadsheetId, id) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const sh = getItemsSheet_(spreadsheetId);
    const last = sh.getLastRow();
    if (last < 2) return {ok:true,deleted:false};
    const ids = sh.getRange(2,1,last-1,1).getValues();
    for (let i=0;i<ids.length;i++) {
      if (String(ids[i][0]) === id) { sh.deleteRow(i+2); return {ok:true,deleted:true}; }
    }
    return {ok:true,deleted:false};
  } finally { lock.releaseLock(); }
}

function saveSettings_(spreadsheetId, settings) {
  const sh = getSettingsSheet_(spreadsheetId);
  sh.getRange('A1:B3').setValues([
    ['key','value'],
    ['locations',JSON.stringify(Array.isArray(settings.locations)?settings.locations:[])],
    ['stores',JSON.stringify(Array.isArray(settings.stores)?settings.stores:[])]
  ]);
  return {ok:true};
}

function saveAll_(spreadsheetId, payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = getItemsSheet_(spreadsheetId);
    if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,ITEM_HEADERS.length).clearContent();
    const items = Array.isArray(payload.items) ? payload.items : [];
    if (items.length) sh.getRange(2,1,items.length,ITEM_HEADERS.length).setValues(items.map(itemToRow_));
    saveSettings_(spreadsheetId, payload.settings || {});
    return {ok:true,count:items.length};
  } finally { lock.releaseLock(); }
}

function readSettings_(ss) {
  const sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh || sh.getLastRow() < 3) return {locations:[],stores:[]};
  const rows = sh.getRange(2,1,Math.min(2,sh.getLastRow()-1),2).getValues();
  const out = {locations:[],stores:[]};
  rows.forEach(r=>{
    try {
      const val = JSON.parse(String(r[1] || '[]'));
      if (r[0] === 'locations') out.locations = Array.isArray(val)?val:[];
      if (r[0] === 'stores') out.stores = Array.isArray(val)?val:[];
    } catch(err) {}
  });
  return out;
}

function getItemsSheet_(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sh = ss.getSheetByName(ITEMS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ITEMS_SHEET);
    sh.getRange(1,1,1,ITEM_HEADERS.length).setValues([ITEM_HEADERS]);
  }
  return sh;
}

function getSettingsSheet_(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) sh = ss.insertSheet(SETTINGS_SHEET);
  return sh;
}

function initializeDatabase_(ss) {
  let items = ss.getSheetByName(ITEMS_SHEET);
  if (!items) items = ss.insertSheet(ITEMS_SHEET);
  items.clear();
  const headers = ['id','name','spec','quantity','image','location','obtained','expiry','store','price','memo','createdAt','updatedAt'];
  items.getRange(1,1,1,headers.length).setValues([headers]);
  items.setFrozenRows(1);
  items.getRange(1,1,1,headers.length).setFontWeight('bold');

  let settings = ss.getSheetByName(SETTINGS_SHEET);
  if (!settings) settings = ss.insertSheet(SETTINGS_SHEET);
  settings.clear();
  settings.getRange('A1:B3').setValues([
    ['key','value'],
    ['locations',JSON.stringify(DEFAULT_LOCATIONS)],
    ['stores',JSON.stringify(DEFAULT_STORES)]
  ]);
  settings.setFrozenRows(1);
  const first = ss.getSheetByName('Sheet1');
  if (first && first.getSheetId() !== items.getSheetId() && first.getLastRow() === 0) ss.deleteSheet(first);
}

function output_(obj, prefix) {
  const json = JSON.stringify(obj);
  if (prefix) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(prefix)) throw new Error('prefix is invalid');
    return ContentService.createTextOutput(prefix+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
