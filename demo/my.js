/* 示範版「我的提問紀錄」頁面（my.js）
 * 畫面與判斷比照正式版 src/asp/S_coach_my.asp；資料改由 history_store.js 從 localStorage 讀取。
 * localStorage 無法使用時：顯示提示、清單為空，其餘畫面照常。 */
(function () {
  'use strict';
  var HIST = window.DEMO_HISTORY;
  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function icon(id, s) { s = s || 14; return '<svg width="' + s + '" height="' + s + '"><use href="#' + id + '"/></svg>'; }
  function toast(m) { var t = $('toast'); if (!t) return; t.textContent = m; t.classList.add('show'); clearTimeout(toast.t); toast.t = setTimeout(function () { t.classList.remove('show'); }, 2400); }
  function safe(fn, fb) { try { return HIST ? fn() : fb; } catch (e) { return fb; } }

  var canStore = safe(function () { return HIST.available(); }, false);
  if ($('maxN') && HIST) $('maxN').textContent = HIST.MAX;
  if (!canStore) { $('noStore').hidden = false; $('clearBtn').disabled = true; }

  // 期間：同正式版 ?d=30 / 90 / 365，預設 90 天
  var m = location.search.match(/[?&]d=(\d+)/), days = m ? +m[1] : 90;
  if ([30, 90, 365].indexOf(days) < 0) days = 90;
  $('dayBadge').textContent = '近 ' + days + ' 天';
  Array.prototype.forEach.call(document.querySelectorAll('#dayChips .chip'), function (c) { c.classList.toggle('on', +c.dataset.d === days); });

  // 結果標籤：同正式版 MyBadge()
  function badge(r, mode) {
    if (mode !== 'ask') return '<span class="badge neutral">關鍵字查詢</span>';
    if (r === 'pending') return '<span class="badge info pulse">AI 整理中</span>';
    if (r === '找不到' || r === '無關鍵詞') return '<span class="badge danger">找不到</span>';
    if (r === 'low') return '<span class="badge warning">可能相關</span>';
    if (String(r).indexOf('AI') >= 0) return '<span class="badge success">AI 整理</span>';
    return '<span class="badge info">標準原文</span>';
  }

  function row(r) {
    var q = encodeURIComponent(r.q || ''), refs = (r.refs || []).map(function (x) {
      return '<span class="rel-i" style="padding:4px 8px;margin:2px 0"><span class="tag code">' + esc(x.code) + '</span><span>' + esc(x.title) + '</span></span>';
    }).join('');
    var t = String(r.time || '').split(' '), act;
    if (r.mode === 'ask' && r.resp) {
      act = '<a class="lnk primary" href="index.html?log=' + r.id + '#chat">' + icon('i-file') + '查看回答</a> ' +
            '<a class="lnk" href="index.html?q=' + q + '#chat" title="重新查詢">' + icon('i-refresh') + '</a>';
    } else if (r.mode === 'ask') {
      act = '<a class="lnk primary" href="index.html?q=' + q + '#chat">' + icon('i-refresh') + '再問一次</a>';
    } else {
      act = '<a class="lnk" href="index.html?kw=' + q + '#kw">' + icon('i-search') + '再查一次</a>';
    }
    var sum = r.aiText ? '<span class="cell-sub" style="margin-top:4px;line-height:1.6">' + esc(String(r.aiText).slice(0, 90)) + '…</span>' : '';
    return '<tr data-s="' + esc(((r.q || '') + ' ' + (r.refs || []).map(function (x) { return x.code + ' ' + x.title; }).join(' ')).toLowerCase()) + '">' +
      '<td class="num" style="text-align:left">' + esc(t[0]) + '<span class="cell-sub">' + esc(t[1] || '') + '</span></td>' +
      '<td class="q-cell"><b>' + esc(r.q) + '</b>' + sum + '</td>' +
      '<td class="hide-m">' + refs + '</td>' +
      '<td>' + badge(r.result, r.mode) + '</td>' +
      '<td style="text-align:right">' + act + '</td></tr>';
  }

  function render() {
    var since = Date.now() - days * 86400000;
    var list = safe(function () { return HIST.list(); }, []).filter(function (r) { return !r.ts || r.ts >= since; });
    var ask = list.filter(function (r) { return r.mode === 'ask'; });
    $('nAsk').textContent = ask.length;
    $('nKw').textContent = list.length - ask.length;
    $('nMiss').textContent = ask.filter(function (r) { return r.result === '找不到' || r.result === 'low' || r.result === '無關鍵詞'; }).length;
    $('tb').innerHTML = list.length ? list.map(row).join('') :
      '<tr><td colspan="5" class="empty">' + (canStore ? '這段期間還沒有紀錄，到「問助理」問第一個問題吧！' : '這個瀏覽器無法保存提問紀錄。') + '</td></tr>';
    $('clearBtn').disabled = !canStore || !safe(function () { return HIST.list().length; }, 0);
    filter();
  }

  function filter() {
    var q = $('tf').value.trim().toLowerCase(), rows = document.querySelectorAll('#tb tr[data-s]'), n = 0;
    Array.prototype.forEach.call(rows, function (r) { var ok = !q || r.dataset.s.indexOf(q) >= 0; r.style.display = ok ? '' : 'none'; if (ok) n++; });
    $('noRow').style.display = (rows.length && !n) ? 'block' : 'none';
  }
  $('tf').oninput = filter;

  $('clearBtn').onclick = function () {
    if (!window.confirm('確定要清除這個瀏覽器中所有的提問紀錄嗎？清除後無法復原。')) return;
    var ok = safe(function () { return HIST.clear(); }, false);
    render();
    toast(ok ? '已清除提問紀錄' : '清除失敗：無法存取 localStorage');
  };

  // 右上角「回內網首頁」：示範版沒有
  document.addEventListener('click', function (e) {
    if (e.target.closest && e.target.closest('[data-demo-off]')) { e.preventDefault(); toast('示範版未提供此功能'); }
  });

  // 側欄的標準數量（以 file:// 開啟讀不到 JSON 時顯示 —）
  function setN(v) { Array.prototype.forEach.call(document.querySelectorAll('[data-demo-n]'), function (el) { el.textContent = v; }); }
  fetch('sample_standards.json', { cache: 'no-store' }).then(function (r) { if (!r.ok) throw 0; return r.json(); })
    .then(function (d) { var n = (d.standards || []).length; setN(n); var b = $('demoKbBar'); if (b) b.style.width = n ? '100%' : '0'; })
    .catch(function () { setN('—'); });

  // 其他分頁新增紀錄時同步更新
  window.addEventListener('storage', function (e) { if (!e.key || e.key.indexOf('sc_demo_history') === 0) render(); });

  render();
})();
