/* =====================================================================
 * 護理標準查詢助理  示範版  本機模擬 API（mock_api.js）
 *
 *   正式版：前端 post() → S_coach_api.asp → 檢索 S_coach.mdb（由 amend_work.mdb
 *           與標準 PDF 建立的索引）→ 呼叫院內 Ollama
 *   示範版：前端 post() → DEMO.post()（本檔）→ 檢索 sample_standards.json
 *           → 以規則「模擬」AI 整理（直接摘錄原文句子，不是真正的語言模型）
 *
 *   回傳的 JSON 欄位與 S_coach_api.asp 相同，所以 app.js 的畫面程式不用改。
 *   提問紀錄（正式版 qlog）改存瀏覽器 localStorage，見 history_store.js；act=getlog 由此讀回。
 *   檢索邏輯是簡化版（關鍵詞＋同義詞比對），結果排序不保證與正式版一致。
 * ===================================================================== */
(function () {
  'use strict';

  // 與正式版 S_coach_types.asp 的 SC_TYPE_DEF 相同（CPT 必須排在 CP 前面）
  var TYPE_DEF = 'CPT=照護準則|CP=護理技術|CR=護理常規|CE=儀器設備|EP=檢查護理|DG=紀錄表單|IG=資訊系統作業|NP=專科護理師技術|WM=病室管理|MM=資材管理';
  var TYPES = TYPE_DEF.split('|').map(function (d) { var p = d.split('='); return [p[0], p[1]]; });

  // 與正式版 config.asp 的檢索參數相同
  var TOPK = 6, PER_STD = 2, MIN_COVER = 0.35;

  // 與正式版 S_coach_api.asp 的問題意圖判斷相同
  var INTENTS = [
    ['注意事項', ['注意事項', '注意', '小心', '禁忌', '避免'], ['注意事項', '注意']],
    ['步驟', ['步驟', '流程', '怎麼做', '如何做', '操作', '執行方式', '方法'], ['步驟', '操作', '流程', '方法']],
    ['用物', ['用物', '準備', '物品', '器材', '需要什麼東西'], ['用物', '準備']],
    ['目的', ['目的', '為什麼', '原因', '原則'], ['目的', '原則']],
    ['更換', ['多久', '頻率', '幾天', '幾小時', '更換', '間隔'], ['更換', '頻率', '小時', '每日']],
    ['評估', ['評估', '觀察', '監測'], ['評估', '觀察']]
  ];
  // 與正式版 S_coach_syn.asp 的停用詞相同
  var STOP = ['請問', '怎麼', '如何', '什麼', '多久', '多少', '要', '嗎', '呢', '的', '是否', '需要', '可以', '病人', '注意事項', '步驟', '?', '？', '，', ','];

  var DATA = null, CHUNKS = [], STDS = {}, VOCAB = [], GROUPS = [], CORPUS = '';

  function norm(s) { return String(s || '').toLowerCase().replace(/[\s　 ]+/g, ''); }
  function typeOf(code) {
    var c = String(code || '').toUpperCase();
    var seg = c.split('-');
    if (seg.length >= 3) for (var j = 0; j < TYPES.length; j++) if (TYPES[j][0] === seg[1]) return seg[1];
    for (var i = 0; i < TYPES.length; i++) {
      var k = TYPES[i][0];
      if (c.indexOf(k) === 0 && !(k === 'CP' && c.indexOf('CPT') === 0)) return k;
    }
    return '';
  }
  function typeName(k) { for (var i = 0; i < TYPES.length; i++) if (TYPES[i][0] === k) return TYPES[i][1]; return k || '其他'; }
  function unitOf(code) { var p = String(code).toUpperCase().split('-'); return p.length >= 3 && typeOf(code) === p[1] ? p[0] : ''; }
  function plainTitle(t) { return String(t || '').replace(/（示範）|\(示範\)/g, ''); }

  // 同義詞群組：一個詞 → 同組所有寫法
  function variantsOf(w) {
    var n = norm(w);
    for (var i = 0; i < GROUPS.length; i++) {
      if (GROUPS[i].some(function (v) { return norm(v) === n; })) return GROUPS[i].slice();
    }
    return [w];
  }

  function prepare(data) {
    DATA = data; CHUNKS = []; STDS = {}; GROUPS = (data.synonyms || []).map(function (g) { return g.slice(); });
    var vocab = {};
    (data.standards || []).forEach(function (s) {
      STDS[s.code] = s;
      vocab[plainTitle(s.title)] = 1;
      (s.keywords || []).forEach(function (k) { vocab[k] = 1; });
      (s.pages || []).forEach(function (p) {
        (p.sections || [p.text || '']).forEach(function (txt) {
          CHUNKS.push({ id: CHUNKS.length + 1, code: s.code, page: p.page, text: txt, ntext: norm(txt) });
        });
      });
    });
    GROUPS.forEach(function (g) { g.forEach(function (v) { vocab[v] = 1; }); });
    CORPUS = CHUNKS.map(function (c) { return c.ntext; }).join('|') + '|' + (data.standards || []).map(function (s) { return norm(s.title); }).join('|');
    VOCAB = Object.keys(vocab).filter(function (v) { return v.length >= 2; })
      .sort(function (a, b) { return b.length - a.length; });   // 長詞優先
  }

  function scopeOk(code, types, src) {
    if (types) { var t = typeOf(code); if (types.split(',').indexOf(t) < 0) return false; }
    if (src) {
      var u = unitOf(code);
      if (src === 'ND' && u) return false;
      if (src === 'SUS' && !u) return false;
      if (src.indexOf('SUS:') === 0 && u !== src.slice(4)) return false;
    }
    return true;
  }

  // 問句 → 查詢重點（concepts）與意圖
  function analyze(q) {
    var ql = String(q || '').toLowerCase(), intent = '', hk = [];
    INTENTS.forEach(function (it) {
      it[1].forEach(function (w) {
        if (ql.indexOf(w) >= 0) { if (!intent) { intent = it[0]; hk = it[2]; } ql = ql.split(w).join(' '); }
      });
    });
    var concepts = [];
    // 1) 詞庫（標準名稱、關鍵詞、同義詞）長詞優先比對
    VOCAB.forEach(function (v) {
      var lv = v.toLowerCase();
      if (ql.indexOf(lv) >= 0) {
        var vs = variantsOf(v);
        if (!concepts.some(function (c) { return c.vs.indexOf(v) >= 0; })) concepts.push({ w: vs[0], vs: vs });
        ql = ql.split(lv).join(' ');
      }
    });
    // 2) 剩下的字：去停用詞後，取出「標準原文中有出現」的 2～4 字片段（長的優先）；
    //    都沒有出現時保留原片段，讓結果顯示「標準中找不到」
    STOP.forEach(function (w) { ql = ql.split(w).join(' '); });
    var raw = [], found = [];
    ql.split(/[^0-9a-z㐀-鿿]+/).forEach(function (seg) {
      if (seg.length < 2 || /^\d+$/.test(seg)) return;
      raw.push(seg);
      if (CORPUS.indexOf(seg) >= 0) { found.push(seg); return; }
      for (var len = Math.min(4, seg.length - 1); len >= 2; len--) {
        for (var i = 0; i + len <= seg.length; i++) {
          var sub = seg.substr(i, len);
          if (CORPUS.indexOf(sub) >= 0) { found.push(sub); seg = seg.slice(0, i) + ' ' + seg.slice(i + len); }
        }
      }
    });
    (found.length || concepts.length ? found : raw).forEach(function (w) {
      if (!concepts.some(function (c) { return c.w === w; })) concepts.push({ w: w, vs: [w] });
    });
    return { concepts: concepts, intent: intent, hk: hk };
  }

  function hitMask(ntext, concepts) {
    var m = [];
    concepts.forEach(function (c, i) { if (c.vs.some(function (v) { return ntext.indexOf(norm(v)) >= 0; })) m.push(i); });
    return m;
  }

  // ---------- act=ask ----------
  function ask(f) {
    var a = analyze(f.q), concepts = a.concepts;
    var base = { ok: 1, act: 'ask', q: f.orig || f.q, qa: null, kw: f.orig && f.orig !== f.q ? f.q : '', logId: 0,
      aiEnabled: true, aiFmt: 'ollama', aiErr: '', intent: a.intent, ai: null };
    var conJ = concepts.map(function (c) { return { w: c.w, syn: c.vs.length > 1, wt: 1 }; });
    if (!concepts.length) return extend(base, { status: 'noterm', concepts: [], sources: [], stds: [], terms: [], related: [], cover: 0, primary: null });

    var nK = concepts.length, scored = [], perStd = {};
    CHUNKS.forEach(function (ch) {
      if (!scopeOk(ch.code, f.types, f.src)) return;
      var std = STDS[ch.code], m = hitMask(ch.ntext + norm(plainTitle(std.title)), concepts);
      var inBody = hitMask(ch.ntext, concepts), head = ch.text.split('\n')[0];
      var hkHead = a.hk.some(function (h) { return head.indexOf(h) >= 0; });   // 段落標題就是問的章節（注意事項、步驟…）
      if (!inBody.length && !(m.length && hkHead)) return;
      var sc = m.length + (hkHead ? 0.8 : a.hk.some(function (h) { return ch.text.indexOf(h) >= 0; }) ? 0.3 : 0);
      scored.push({ ch: ch, sc: sc, mask: m });
      var ps = perStd[ch.code] || (perStd[ch.code] = { mask: {}, best: 0, rows: [] });
      m.forEach(function (i) { ps.mask[i] = 1; });
      ps.best = Math.max(ps.best, sc); ps.rows.push({ ch: ch, sc: sc });
    });
    var terms = []; concepts.forEach(function (c) { c.vs.forEach(function (v) { if (terms.indexOf(v) < 0) terms.push(v); }); });
    if (!scored.length) return extend(base, { status: 'nohit', concepts: conJ, sources: [], stds: [], terms: terms, related: [], cover: 0, primary: null });

    scored.sort(function (x, y) { return y.sc - x.sc || x.ch.id - y.ch.id; });
    var docs = Object.keys(perStd).map(function (code) {
      var ps = perStd[code]; return { code: code, cov: Object.keys(ps.mask).length / nK, best: ps.best, ps: ps };
    }).sort(function (x, y) { return y.cov - x.cov || y.best - x.best; });
    var primCode = docs[0].code, cover = docs[0].cov;

    // 原文段落：最相關的標準優先，每份最多 PER_STD 段，共 TOPK 段
    var pick = [], cnt = {};
    [primCode].concat(docs.slice(1).map(function (d) { return d.code; })).forEach(function (code) {
      scored.forEach(function (s) {
        if (s.ch.code !== code || pick.length >= TOPK || (cnt[code] || 0) >= PER_STD) return;
        cnt[code] = (cnt[code] || 0) + 1; pick.push(s.ch);
      });
    });
    var sources = pick.map(function (ch, i) { return srcRow(ch, i + 1); });

    var stds = [], related = [];
    docs.forEach(function (d, i) {
      var st = STDS[d.code], info = stdInfo(st);
      if (i === 0 || d.cov >= 0.3) {
        var rows = d.ps.rows.slice().sort(function (x, y) { return y.sc - x.sc; });
        var pages = []; rows.forEach(function (r) { if (pages.indexOf(r.ch.page) < 0) pages.push(r.ch.page); });
        pages.sort(function (x, y) { return x - y; });
        stds.push(extend(info, { cov: Math.round(d.cov * 100) / 100, pages: pages,
          hits: Object.keys(d.ps.mask).map(Number).sort(), ex: rows.slice(0, 3).map(function (r) { return { id: r.ch.id, page: r.ch.page, text: r.ch.text }; }) }));
      } else related.push(info);
    });

    return extend(base, { status: cover < MIN_COVER ? 'low' : 'ok', cover: Math.round(cover * 100) / 100,
      concepts: conJ, terms: terms, primary: stdInfo(STDS[primCode]), stds: stds, related: related, sources: sources });
  }

  function stdInfo(st) { var t = typeOf(st.code); return { code: st.code, title: st.title, cat: st.cat || '', video: st.video || '', type: t, typeName: typeName(t) }; }
  function srcRow(ch, n) { return extend(stdInfo(STDS[ch.code]), { n: n, id: ch.id, page: ch.page, text: ch.text }); }
  function extend(o, p) { for (var k in p) o[k] = p[k]; return o; }

  // 模擬 AI：從引用段落中挑出含查詢重點的句子，原文照錄並加上 [來源n]
  function fakeAI(d) {
    var a = analyze(d.q), lines = ['以下依示範標準摘錄整理重點（示範版以規則模擬 AI 輸出，正式版由院內 Ollama 產生）：'];
    var byStd = {}, order = [];
    d.sources.forEach(function (s) { if (!byStd[s.code]) { byStd[s.code] = []; order.push(s.code); } byStd[s.code].push(s); });
    order.forEach(function (code) {
      var st = STDS[code], out = [];
      var has = function (txt) { var nt = norm(txt); return a.concepts.some(function (c) { return c.vs.some(function (v) { return nt.indexOf(norm(v)) >= 0; }); }); };
      byStd[code].forEach(function (s) {
        var lines2 = s.text.split(/\n/), head = lines2[0];
        // 段落標題符合（例：「五、濕洗手步驟」「六、注意事項」）→ 整段重點都列；否則只列含查詢重點的句子
        var whole = has(head) || a.hk.some(function (h) { return head.indexOf(h) >= 0; });
        lines2.slice(1).forEach(function (sent) {
          var t = sent.replace(/^\s*\d+[.、)]\s*/, '').trim();
          if (!t || out.length >= 5) return;
          if (whole || has(t)) out.push('- ' + t + ' [來源' + s.n + ']');
        });
      });
      if (!out.length) out.push('- ' + byStd[code][0].text.split('\n').slice(1).join('').slice(0, 60) + ' [來源' + byStd[code][0].n + ']');
      lines.push('### 【' + typeName(typeOf(code)) + '】' + code + ' ' + st.title);
      lines = lines.concat(out);
    });
    return { message: { role: 'assistant', content: lines.join('\n') } };
  }

  // ---------- act=search（關鍵字查詢）----------
  function search(f) {
    var kws = String(f.q || '').trim().split(/\s+/).filter(Boolean);
    if (!kws.length) return { ok: 0, err: '請輸入關鍵字' };
    var kVar = kws.map(variantsOf), terms = [];
    kVar.forEach(function (vs) { vs.forEach(function (v) { if (terms.indexOf(v) < 0) terms.push(v); }); });
    var results = [];
    (DATA.standards || []).forEach(function (st) {
      if (!scopeOk(st.code, f.types, f.src)) return;
      var nt = norm(st.title), matched = 0, hits = 0, titleHit = false, snips = [];
      kVar.forEach(function (vs) {
        var got = false;
        vs.forEach(function (v) {
          var nv = norm(v);
          if (nt.indexOf(nv) >= 0) { got = true; titleHit = true; }
          CHUNKS.forEach(function (ch) {
            if (ch.code !== st.code) return;
            var pos = ch.ntext.indexOf(nv);
            if (pos < 0) return;
            got = true; hits += ch.ntext.split(nv).length - 1;
            if (snips.length < 3 && !snips.some(function (x) { return x.p === ch.page; })) {
              var flat = ch.text.replace(/\s+/g, ' '), p2 = flat.toLowerCase().indexOf(v.toLowerCase());
              if (p2 < 0) p2 = 0;
              snips.push({ p: ch.page, t: (p2 > 30 ? '…' : '') + flat.slice(Math.max(0, p2 - 30), p2 + 50) + '…' });
            }
          });
        });
        if (got) matched++;
      });
      if (!matched) return;
      var t = typeOf(st.code);
      results.push({ code: st.code, typeName: typeName(t), title: st.title, cat: st.cat || '', video: st.video || '',
        hits: hits, titleHit: titleHit, full: matched >= kws.length, indexed: true, snips: snips, _m: matched });
    });
    results.sort(function (a, b) { return b._m - a._m || (b.titleHit - a.titleHit) || b.hits - a.hits; });
    return { ok: 1, act: 'search', nk: kws.length, terms: terms, results: results };
  }

  function delay(v, ms) { return new Promise(function (r) { setTimeout(function () { r(v); }, ms); }); }

  // ---------- 提問紀錄（存在瀏覽器 localStorage，見 history_store.js）----------
  // history_store.js 沒載入或 localStorage 無法使用時，所有紀錄操作都安靜地略過
  var HIST = window.DEMO_HISTORY;
  function safe(fn, fallback) { try { return HIST ? fn() : fallback; } catch (e) { return fallback; } }
  function resultOf(r) { return r.status === 'nohit' ? '找不到' : r.status === 'noterm' ? '無關鍵詞' : r.status; }   // 同正式版 qlog.result
  function refsOf(r) {
    var seen = {}, out = [];
    [r.primary].concat(r.stds || [], r.sources || []).forEach(function (s) {
      if (s && s.code && !seen[s.code]) { seen[s.code] = 1; out.push({ code: s.code, title: s.title || '' }); }
    });
    return out.slice(0, 5);
  }
  function plainAI(ai) {
    var t = ai && ai.message ? String(ai.message.content || '') : '';
    t = t.replace(/^以下依示範標準[^\n]*\n/, '');   // 略過示範版固定的開頭說明，讓紀錄摘要直接顯示重點
    return t.replace(/#{1,4}\s*/g, '').replace(/\[來源\s*\d+\]/g, '').replace(/\s*\n\s*/g, ' ').trim();
  }

  // ---------- 給 app.js 用的介面 ----------
  window.DEMO = {
    types: TYPES,
    aiOn: true,
    std: function () {   // [編號, 名稱, 分類, 影片, 已建索引, 修改日, 單位]
      return (DATA.standards || []).map(function (s) { return [s.code, s.title, s.cat || '', s.video || '', 1, s.rev_date || '', s.unit || '']; });
    },
    post: function (f) {
      var r;
      try {
        if (f.act === 'ask') {
          r = ask(f);
          var isAI = String(f.ai) === '1' && r.status === 'ok';
          if (isAI) r.ai = fakeAI(r);
          // 第 2 步（AI 整理，nolog=1）：更新第 1 步建立的同一筆紀錄（同正式版 SaveResp）
          if (String(f.nolog) === '1') {
            r.logId = +f.logid || 0;
            return delay(r, isAI ? 1200 : 300).then(function (x) {
              if (isAI && x.logId) safe(function () { HIST.update(x.logId, { resp: x, aiText: plainAI(x.ai), result: 'ok+AI' }); });
              return x;
            });
          }
          // 第 1 步：建立紀錄；AI 改寫關鍵詞後重查（replog）時，刪掉前一筆（同正式版）
          if (+f.replog) safe(function () { HIST.remove(f.replog); });
          return delay(r, isAI ? 1200 : 300).then(function (x) {
            x.logId = safe(function () {
              return HIST.add({ mode: 'ask', q: x.q, result: isAI ? 'ok+AI' : resultOf(x), refs: refsOf(x),
                aiText: isAI ? plainAI(x.ai) : '',
                resp: x.status === 'ok' || x.status === 'low' ? x : null });   // 同正式版：找不到的題目不保存回答，紀錄頁顯示「再問一次」
            }, Date.now());
            return x;
          });
        }
        if (f.act === 'search') {
          return delay(search(f), 250).then(function (x) {
            if (x.ok) safe(function () {
              HIST.add({ mode: 'search', q: String(f.q || '').trim(), result: x.results.length + '筆', aiText: '', resp: null,
                refs: x.results.slice(0, 3).map(function (s) { return { code: s.code, title: s.title }; }) });
            });
            return x;
          });
        }
        if (f.act === 'kwx') return delay({ ok: 0, err: '示範版沒有 AI 關鍵詞判讀' }, 100);
        if (f.act === 'fb') return delay({ ok: 1 }, 150);
        if (f.act === 'getlog') {
          var g = safe(function () { return HIST.get(f.id); }, null);
          if (!g) return delay({ ok: 0, err: safe(function () { return HIST.available(); }, false)
            ? '找不到這筆紀錄（可能已清除）' : '這個瀏覽器無法使用 localStorage，示範版無法保存提問紀錄' }, 100);
          return delay({ ok: 1, q: g.q, time: g.time, result: g.result, resp: g.mode === 'ask' ? (g.resp || null) : null }, 100);
        }
      } catch (e) { return Promise.reject(e); }
      return delay({ ok: 0, err: 'unknown act' }, 50);
    },
    // 正式版開啟 S_coach_pdf.asp?code=…#page=…；示範版改為在頁面中顯示該標準的原文
    pdfUrl: function (code, page) { return '#demo-pdf=' + encodeURIComponent(code) + (page ? '&p=' + page : ''); }
  };

  // ---------- 「開啟 PDF」示範視窗 ----------
  function openDoc(code, page) {
    var st = STDS[code]; if (!st) return;
    var box = document.getElementById('demoDoc');
    var esc = function (s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
    box.querySelector('.dd-t').innerHTML = '<span class="tag code">' + esc(st.code) + '</span> ' + esc(st.title);
    box.querySelector('.dd-b').innerHTML = st.pages.map(function (p) {
      return '<section class="dd-pg' + (+p.page === +page ? ' on' : '') + '" data-p="' + p.page + '"><div class="dd-pn">第 ' + p.page + ' 頁</div>' +
        (p.sections || [p.text || '']).map(function (s) { return '<p>' + esc(s).replace(/\n/g, '<br>') + '</p>'; }).join('') + '</section>';
    }).join('');
    box.hidden = false;
    var on = box.querySelector('.dd-pg.on');
    if (on) on.scrollIntoView({ block: 'start' }); else box.querySelector('.dd-b').scrollTop = 0;
  }
  document.addEventListener('click', function (e) {
    var a = e.target.closest && e.target.closest('a[href^="#demo-pdf="]');
    if (a) {
      e.preventDefault(); e.stopPropagation();
      var m = a.getAttribute('href').match(/^#demo-pdf=([^&]+)(?:&p=(\d+))?/);
      if (m) openDoc(decodeURIComponent(m[1]), m[2] ? +m[2] : 0);
      return;
    }
    if (e.target.closest && e.target.closest('[data-demo-off]')) { e.preventDefault(); var t = document.getElementById('toast'); if (t) { t.textContent = '示範版未提供此功能'; t.classList.add('show'); setTimeout(function () { t.classList.remove('show'); }, 2000); } }
    if (e.target.closest && e.target.closest('[data-dd-close]')) document.getElementById('demoDoc').hidden = true;
  }, true);
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') { var b = document.getElementById('demoDoc'); if (b) b.hidden = true; } });

  // ---------- 啟動：讀 sample_standards.json，成功後才載入畫面程式 app.js ----------
  function start(data) {
    prepare(data);
    var n = (data.standards || []).length;
    Array.prototype.forEach.call(document.querySelectorAll('[data-demo-n]'), function (el) { el.textContent = n; });
    var bar = document.getElementById('demoKbBar'); if (bar) bar.style.width = n ? '100%' : '0';
    var pk = document.getElementById('demoPick'); if (pk) pk.remove();
    var s = document.createElement('script'); s.src = 'app.js'; document.body.appendChild(s);
  }
  function showPicker(err) {
    var box = document.createElement('section');
    box.className = 'card'; box.id = 'demoPick'; box.style.cssText = 'padding:20px;margin:0 0 16px';
    box.innerHTML = '<div class="banner warn"><div><b>瀏覽器不允許直接讀取同資料夾的 sample_standards.json</b>（以 file:// 開啟時常見）。<br>' +
      '請按下方按鈕選擇 <code>demo/sample_standards.json</code>；或在 <code>demo/</code> 資料夾內執行 <code>python -m http.server 8000</code>，改開 <code>http://localhost:8000/</code>。</div></div>' +
      '<input type="file" accept=".json,application/json" id="demoFile">' +
      '<div class="note" style="margin-top:8px">錯誤訊息：' + String(err && err.message || err).replace(/</g, '&lt;') + '</div>';
    var c = document.querySelector('.main .container'); c.insertBefore(box, c.firstChild);
    document.getElementById('demoFile').onchange = function () {
      var f = this.files[0]; if (!f) return;
      var rd = new FileReader();
      rd.onload = function () { try { start(JSON.parse(rd.result)); } catch (e) { alert('JSON 格式錯誤：' + e.message); } };
      rd.readAsText(f, 'utf-8');
    };
  }
  fetch('sample_standards.json', { cache: 'no-store' })
    .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
    .then(start).catch(showPicker);
})();
