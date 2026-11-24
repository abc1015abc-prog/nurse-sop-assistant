/* =====================================================================
 * 示範版「我的提問紀錄」儲存（history_store.js）
 *   正式版：紀錄存在伺服器資料庫 S_coach.mdb 的 qlog 表（S_coach_api.asp 寫入）
 *   示範版：紀錄只存在使用者自己瀏覽器的 localStorage，不會上傳
 *
 *   每筆紀錄：
 *   { id, time:'YYYY/MM/DD HH:mm', ts, mode:'ask'|'search', q,
 *     result:'ok'|'ok+AI'|'low'|'找不到'|'無關鍵詞'|'n筆',
 *     refs:[{code,title}], aiText, resp:<與 act=ask 相同格式的完整回答，查看回答用> }
 *
 *   localStorage 在無痕模式、封鎖網站資料或容量已滿時可能無法使用，
 *   所有讀寫都包在 try/catch 裡；失敗時只是不保存，頁面照常運作。
 * ===================================================================== */
(function () {
  'use strict';
  var KEY = 'sc_demo_history_v1', MAX = 50;

  function available() {
    try {
      var k = '__sc_demo_test__';
      window.localStorage.setItem(k, '1');
      window.localStorage.removeItem(k);
      return true;
    } catch (e) { return false; }
  }

  function load() {
    try {
      var a = JSON.parse(window.localStorage.getItem(KEY) || '[]');
      return Array.isArray(a) ? a.filter(function (r) { return r && typeof r === 'object' && r.id; }) : [];
    } catch (e) { return []; }
  }

  // 寫入；容量不足時從最舊的開始捨棄再試
  function store(list) {
    list = list.slice(0, MAX);
    while (true) {
      try { window.localStorage.setItem(KEY, JSON.stringify(list)); return true; }
      catch (e) {
        if (list.length <= 1) return false;
        list = list.slice(0, list.length - 1);
      }
    }
  }

  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function fmt(d) { return d.getFullYear() + '/' + pad(d.getMonth() + 1) + '/' + pad(d.getDate()) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()); }

  var lastId = 0;
  function newId(list) {
    var id = Date.now();
    if (id <= lastId) id = lastId + 1;
    while (list.some(function (r) { return r.id === id; })) id++;
    lastId = id;
    return id;
  }

  window.DEMO_HISTORY = {
    MAX: MAX,
    available: available,
    list: load,                          // 新的在前
    get: function (id) {
      id = +id;
      var a = load();
      for (var i = 0; i < a.length; i++) if (a[i].id === id) return a[i];
      return null;
    },
    // 新增一筆，回傳紀錄 id（無法保存時仍回傳 id，讓畫面的回饋按鈕可以運作）
    add: function (rec) {
      var a = load(), now = new Date();
      rec.id = newId(a); rec.ts = now.getTime(); rec.time = fmt(now);
      a.unshift(rec);
      store(a);
      return rec.id;
    },
    update: function (id, patch) {
      id = +id;
      var a = load(), hit = false;
      a.forEach(function (r) { if (r.id === id) { for (var k in patch) r[k] = patch[k]; hit = true; } });
      if (hit) store(a);
      return hit;
    },
    remove: function (id) {
      id = +id;
      var a = load(), b = a.filter(function (r) { return r.id !== id; });
      if (b.length !== a.length) store(b);
    },
    clear: function () {
      try { window.localStorage.removeItem(KEY); return true; } catch (e) { return false; }
    }
  };
})();
