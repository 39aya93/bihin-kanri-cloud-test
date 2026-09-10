/*
 * 備品管理アプリの接続設定
 *
 * 1) MANAGER_URL に「DB管理用Apps Script」のWebアプリURLを入れます。
 * 2) DB_LIST はDB作成後に「DB一覧を更新」ボタンで自動更新できます。
 *    手動で編集する必要は基本ありません。
 */
const APP_CONFIG = {
  MANAGER_URL: 'https://script.google.com/macros/s/AKfycbyV0nMIoA0AwbA-JjOOvvCPGzBwmoPKf9iVqJUmF7p1fTj1KIAx_SRi23gpZPb-1-8/exec',
  DB_LIST: [],
  ACTIVE_DB_ID: ''
};
