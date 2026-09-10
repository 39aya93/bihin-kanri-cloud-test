/*
 * 備品管理アプリの接続設定
 *
 * 1) MANAGER_URL に「DB管理用Apps Script」のWebアプリURLを入れます。
 * 2) DB_LIST はDB作成後に「DB一覧を更新」ボタンで自動更新できます。
 *    手動で編集する必要は基本ありません。
 */
const APP_CONFIG = {
  MANAGER_URL: 'https://script.google.com/macros/s/AKfycbzMC9b2NRUvFbSvA7NfB301gl5jNVBanY5dWFkZLLL1QMq9f32Rx9M7pa7Ax2bSbRbm/exec',
  DB_LIST: [],
  ACTIVE_DB_ID: ''
};
