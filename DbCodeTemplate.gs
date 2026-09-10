/**
 * 備品管理：DBごとのスプレッドシートに自動配置されるCode.gs
 * このファイルはDB管理用スクリプトがApps Script APIで各Spreadsheetへコピーします。
 */

const ITEMS_SHEET = 'Items';
const SETTINGS_SHEET = 'Settings';
const ITEM_HEADERS = ['id','name','spec','quantity','image','location','obtained','expiry','store','price','memo','createdAt','updatedAt'];
const ALLOWED_USERS = __ALLOWED_USERS__;

function assertUser_() {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('Googleアカウントを確認できません。DBのWebアプリにはGoogleアカウントでログインしてアクセスしてください。');
  if (ALLOWED_USERS.length && !ALLOWED_USERS.map(x=>String(x).toLowerCase()).includes(email)) throw new Error('このDBを利用する権限がありません：'+email);
  return email;
}

function doGet(e) {
  assertUser_();
  const p = e && e.parameter ? e.parameter : {};
  const op = p.op || 'all';
  let result;
  try {
    if (op === 'all') result = getAll_();
    else if (op === 'health') result = {ok:true,db:SpreadsheetApp.getActiveSpreadsheet().getName(),time:new Date().toISOString()};
    else throw new Error('不明な操作です: '+op);
  } catch(err) {
    result = {ok:false,error:String(err && err.message ? err.message : err)};
  }
  return output_(result,p.prefix);
}

function doPost(e) {
  let result;
  try { assertUser_(); } catch (authErr) { return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(authErr.message||authErr)})).setMimeType(ContentService.MimeType.JSON); }
  try {
    const p = e && e.parameter ? e.parameter : {};
    const payload = p.payload ? JSON.parse(p.payload) : p;
    const action = payload.action || '';
    if (action === 'saveItem') result = saveItem_(payload.item);
    else if (action === 'deleteItem') result = deleteItem_(String(payload.id));
    else if (action === 'saveSettings') result = saveSettings_(payload.settings || {});
    else if (action === 'saveAll') result = saveAll_(payload);
    else throw new Error('不明な書き込み操作です: '+action);
  } catch(err) {
    result = {ok:false,error:String(err && err.message ? err.message : err)};
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function getAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(ITEMS_SHEET);
  const settings = readSettings_();
  if (!sh || sh.getLastRow() < 2) return {ok:true,dbName:ss.getName(),settings:settings,items:[]};
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

function saveItem_(item) {
  if (!item || !item.id) throw new Error('備品IDがありません。');
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const sh = getItemsSheet_();
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

function deleteItem_(id) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const sh = getItemsSheet_();
    const last = sh.getLastRow();
    if (last < 2) return {ok:true,deleted:false};
    const ids = sh.getRange(2,1,last-1,1).getValues();
    for (let i=0;i<ids.length;i++) {
      if (String(ids[i][0]) === id) { sh.deleteRow(i+2); return {ok:true,deleted:true}; }
    }
    return {ok:true,deleted:false};
  } finally { lock.releaseLock(); }
}

function saveSettings_(settings) {
  const sh = getSettingsSheet_();
  sh.getRange('A1:B3').setValues([
    ['key','value'],
    ['locations',JSON.stringify(Array.isArray(settings.locations)?settings.locations:[])],
    ['stores',JSON.stringify(Array.isArray(settings.stores)?settings.stores:[])]
  ]);
  return {ok:true};
}

function saveAll_(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = getItemsSheet_();
    if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,ITEM_HEADERS.length).clearContent();
    const items = Array.isArray(payload.items) ? payload.items : [];
    if (items.length) sh.getRange(2,1,items.length,ITEM_HEADERS.length).setValues(items.map(itemToRow_));
    saveSettings_(payload.settings || {});
    return {ok:true,count:items.length};
  } finally { lock.releaseLock(); }
}

function readSettings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
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

function getItemsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ITEMS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ITEMS_SHEET);
    sh.getRange(1,1,1,ITEM_HEADERS.length).setValues([ITEM_HEADERS]);
  }
  return sh;
}

function getSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) sh = ss.insertSheet(SETTINGS_SHEET);
  return sh;
}

function output_(obj,prefix) {
  const json = JSON.stringify(obj);
  if (prefix) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(prefix)) throw new Error('prefix is invalid');
    return ContentService.createTextOutput(prefix+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
