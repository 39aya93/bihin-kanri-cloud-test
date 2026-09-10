/**
 * 備品管理：DB管理用 Apps Script
 *
 * 役割
 *  - 新しいGoogleスプレッドシートを作成
 *  - そのスプレッドシートにCode.gsをバインド
 *  - Apps Script APIでCode.gsを埋め込み
 *  - バージョン作成→Webアプリとしてデプロイ
 *  - DB名、Spreadsheet ID、WebアプリURLを管理
 *
 * 初回設定
 *  1. このコードを「DB管理」というApps Scriptプロジェクトに貼る
 *  2. appsscript.json の oauthScopes を設定
 *  3. Apps Script APIをGoogle Cloudプロジェクトで有効化
 *  4. 初回だけ setupManager() を実行して認証
 *  5. Webアプリとしてデプロイ（自分として実行／ログインユーザー全員）
 */

const MANAGER_SHEET_NAME = 'DB管理台帳';
const DEFAULT_LOCATIONS = ['工具棚A','工具棚B','ガレージ','物置','パントリー','倉庫','車庫'];
const DEFAULT_STORES = ['カインズ','コーナン','Amazon','楽天市場','モノタロウ','その他'];
// DB管理用Webアプリを使えるGoogleアカウント。空配列なら制限なし。
const ALLOWED_MANAGER_USERS = [];

function assertManagerUser_() {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (ALLOWED_MANAGER_USERS.length && !ALLOWED_MANAGER_USERS.map(x=>String(x).toLowerCase()).includes(email)) throw new Error('DB管理権限がありません：'+(email||'Googleアカウント不明'));
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
    else if (op === 'create') {
      const name = String(p.name || '').trim();
      if (!name) throw new Error('DB名がありません。');
      const editors = String(p.editors || '').split(',').map(x=>x.trim()).filter(Boolean);
      result = {ok:true, database:createDatabase_(name, editors)};
    } else if (op === 'health') result = {ok:true, service:'BihinKanri DB Manager',time:new Date().toISOString()};
    else throw new Error('不明な操作です: '+op);
  } catch (err) {
    result = {ok:false,error:String(err && err.message ? err.message : err)};
  }
  return output_(result,p.prefix);
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

  const ownerEmail = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const allowedUsers = Array.from(new Set([ownerEmail].concat(editors).filter(Boolean)));
  const project = apiCreateProject_(name, spreadsheetId);
  apiUpdateContent_(project.scriptId, getDbCodeTemplate_(allowedUsers), name);
  const version = apiCreateVersion_(project.scriptId, '初期DB版');
  const deployment = apiCreateDeployment_(project.scriptId, version.versionNumber, name);
  const webAppUrl = deployment.entryPoints && deployment.entryPoints[0] && deployment.entryPoints[0].webApp
    ? deployment.entryPoints[0].webApp.url : '';

  const reg = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('REGISTRY_SHEET_ID')).getSheetByName(MANAGER_SHEET_NAME);
  reg.appendRow([name,spreadsheetId,ss.getUrl(),project.scriptId,webAppUrl,editors.join(','),new Date()]);
  return {name,spreadsheetId,spreadsheetUrl:ss.getUrl(),scriptId:project.scriptId,webAppUrl,editors};
}

function initializeDatabase_(ss) {
  let items = ss.getSheetByName('Items');
  if (!items) items = ss.insertSheet('Items');
  items.clear();
  const headers = ['id','name','spec','quantity','image','location','obtained','expiry','store','price','memo','createdAt','updatedAt'];
  items.getRange(1,1,1,headers.length).setValues([headers]);
  items.setFrozenRows(1);
  items.getRange(1,1,1,headers.length).setFontWeight('bold');

  let settings = ss.getSheetByName('Settings');
  if (!settings) settings = ss.insertSheet('Settings');
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

function apiFetch_(url, method, body) {
  const params = {method:method || 'get', headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}, muteHttpExceptions:true};
  if (body !== undefined) {
    params.contentType = 'application/json';
    params.payload = JSON.stringify(body);
  }
  const res = UrlFetchApp.fetch(url, params);
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code < 200 || code >= 300) throw new Error('Apps Script API '+code+': '+text);
  return JSON.parse(text);
}

function apiCreateProject_(title, parentId) {
  return apiFetch_('https://script.googleapis.com/v1/projects','post',{title:title,parentId:parentId});
}

function apiUpdateContent_(scriptId, code, title) {
  const manifest = {
    timeZone: Session.getScriptTimeZone() || 'Asia/Tokyo',
    dependencies: {},
    exceptionLogging: 'STACKDRIVER',
    runtimeVersion: 'V8',
    webapp: {executeAs:'USER_DEPLOYING', access:'ANYONE'}
  };
  const files = [
    {name:'appsscript',type:'JSON',source:JSON.stringify(manifest)},
    {name:'Code',type:'SERVER_JS',source:code}
  ];
  return apiFetch_('https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptId)+'/content','put',{files:files});
}

function apiCreateVersion_(scriptId, description) {
  return apiFetch_('https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptId)+'/versions','post',{description:description});
}

function apiCreateDeployment_(scriptId, versionNumber, description) {
  return apiFetch_('https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptId)+'/deployments','post',{
    versionNumber:versionNumber,
    manifestFileName:'appsscript',
    description:description
  });
}

function output_(obj, prefix) {
  const json = JSON.stringify(obj);
  if (prefix) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(prefix)) throw new Error('prefix is invalid');
    return ContentService.createTextOutput(prefix+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function getDbCodeTemplate_(allowedUsers) {
  return `/**
 * 備品管理：DBごとのスプレッドシートに自動配置されるCode.gs
 * このファイルはDB管理用スクリプトがApps Script APIで各Spreadsheetへコピーします。
 */

const ITEMS_SHEET = 'Items';
const SETTINGS_SHEET = 'Settings';
const ITEM_HEADERS = ['id','name','spec','quantity','image','location','obtained','expiry','store','price','memo','createdAt','updatedAt'];

function doGet(e) {
  assertManagerUser_();
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
`.replace('__ALLOWED_USERS__', JSON.stringify(allowedUsers || []));
}