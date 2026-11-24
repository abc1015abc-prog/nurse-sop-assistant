/* =====================================================================
 * 護理標準查詢助理  示範版  多模態功能（multimodal.js）
 *   1. 語音提問：Web Speech API SpeechRecognition（zh-TW），說完自動送出
 *   2. 朗讀回答：speechSynthesis（優先 zh-TW 語音，語速 0.95），可設定自動朗讀
 *   3. 拍照查標準：transformers.js + CLIP（Xenova/clip-vit-base-patch32）在瀏覽器內
 *      做 zero-shot 圖片分類 → 帶入問題，讓使用者確認後再送出
 *
 *   只操作 index.html 的「問助理」提問列（#q、#send、#thread），不修改 app.js。
 *   注意：Chrome / Edge 的語音辨識是把聲音送到 Google / Microsoft 的雲端辨識，
 *         示範版使用虛構資料沒有問題；正式版若要語音輸入，需改用院內語音辨識。
 * ===================================================================== */
(function () {
  'use strict';

  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  var toastT;
  function toast(m) { var t = $('toast'); if (!t) return; t.textContent = m; t.classList.add('show'); clearTimeout(toastT); toastT = setTimeout(function () { t.classList.remove('show'); }, 3200); }
  // localStorage 在無痕模式或封鎖網站資料時會丟錯，一律 try/catch
  function lsGet(k) { try { return window.localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { window.localStorage.setItem(k, v); } catch (e) { /* 無法保存設定，不影響使用 */ } }

  var qEl = $('q'), sendBtn = $('send'), thread = $('thread');
  function setQuestion(t) { qEl.value = t; qEl.dispatchEvent(new Event('input')); }   // 觸發 app.js 的自動調整高度

  /* =================================================================
   * 1. 語音提問
   * ================================================================= */
  var SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  var micBtn = $('mmMic'), rec = null, recOn = false;

  if (!SR) {
    micBtn.hidden = true;                 // Firefox 等不支援：隱藏麥克風，顯示提示
    $('mmMicHint').hidden = false;
  } else {
    micBtn.onclick = function () {
      if (recOn) { try { rec.stop(); } catch (e) { } return; }   // 錄音中再按一次 = 停止
      startListening();
    };
  }

  function startListening() {
    stopSpeaking();
    var finalTxt = '', ph = qEl.placeholder;
    rec = new SR();
    rec.lang = 'zh-TW';
    rec.interimResults = true;            // 邊說邊顯示在輸入框
    rec.continuous = false;               // 停頓即結束
    rec.maxAlternatives = 1;
    rec.onstart = function () {
      recOn = true;
      micBtn.classList.add('mm-rec'); micBtn.setAttribute('aria-pressed', 'true');
      micBtn.title = '錄音中…再按一次停止';
      qEl.placeholder = '請說出你的問題…';
    };
    rec.onresult = function (e) {
      var interim = '';
      for (var i = e.resultIndex; i < e.results.length; i++) {
        var t = e.results[i][0].transcript;
        if (e.results[i].isFinal) finalTxt += t; else interim += t;
      }
      setQuestion((finalTxt + interim).trim());
    };
    rec.onerror = function (e) {
      var m = {
        'not-allowed': '麥克風權限被拒絕，請在網址列旁允許使用麥克風',
        'service-not-allowed': '瀏覽器不允許語音辨識（請確認網址是 https 或 localhost）',
        'no-speech': '沒有聽到聲音，請再試一次',
        'audio-capture': '找不到麥克風',
        'network': '語音辨識需要網路連線',
        'aborted': ''
      }[e.error];
      if (m !== '') toast(m || ('語音辨識錯誤：' + e.error));
    };
    rec.onend = function () {
      recOn = false;
      micBtn.classList.remove('mm-rec'); micBtn.setAttribute('aria-pressed', 'false');
      micBtn.title = '語音提問（說完自動送出）';
      qEl.placeholder = ph;
      var t = finalTxt.trim();
      if (!t) return;
      setQuestion(t);
      if (sendBtn.disabled) { toast('上一題還在處理，問題已填入，請稍後按「詢問助理」'); return; }
      sendBtn.click();                    // 說完自動送出
    };
    try { rec.start(); } catch (e) { toast('無法啟動語音辨識：' + e.message); }
  }

  /* =================================================================
   * 2. 朗讀回答
   * ================================================================= */
  var TTS = 'speechSynthesis' in window && 'SpeechSynthesisUtterance' in window;
  var AUTO_KEY = 'sc_demo_autoread', voice = null, speakingBtn = null, speakSeq = 0;
  var autoBox = $('mmAuto');

  function pickVoice() {
    var vs = window.speechSynthesis.getVoices() || [];
    function find(re) { for (var i = 0; i < vs.length; i++) if (re.test(vs[i].lang)) return vs[i]; return null; }
    voice = find(/^zh[-_]TW$/i) || find(/^zh[-_](HK|Hant)/i) || find(/^(zh|cmn)/i) || null;   // 優先台灣中文
  }

  if (!TTS) {
    $('mmAutoWrap').hidden = true;
  } else {
    pickVoice();
    if (window.speechSynthesis.addEventListener) window.speechSynthesis.addEventListener('voiceschanged', pickVoice);
    else window.speechSynthesis.onvoiceschanged = pickVoice;
    autoBox.checked = lsGet(AUTO_KEY) === '1';
    autoBox.onchange = function () { lsSet(AUTO_KEY, autoBox.checked ? '1' : '0'); if (!autoBox.checked) stopSpeaking(); };
  }

  // 回答文字 → 朗讀用文字：去掉出處（CP902 p.2）、標準編號、Markdown 與 ⚠ 等符號
  function textForSpeech(box) {
    var c = box.querySelector('.answer').cloneNode(true);
    Array.prototype.forEach.call(c.querySelectorAll('.cite,.banner'), function (n) { n.remove(); });
    Array.prototype.forEach.call(c.querySelectorAll('li,div'), function (n) { n.appendChild(document.createTextNode('\n')); });
    return c.textContent
      .replace(/^以下依示範標準[^\n]*\n/, '')                   // 示範版固定開頭說明不朗讀
      .replace(/\[來源\s*\d+\]/g, '')
      .replace(/\b[A-Z]{2,4}\d{2,4}(?:-\d+)?\b/g, '')         // 標準編號，例 CP902
      .replace(/\bp\.\s*\d+/gi, '')                           // 頁碼
      .replace(/[⚠#*_`>|【】\[\]]/g, ' ')
      .replace(/（示範）/g, '')
      .replace(/[ \t]+/g, ' ')
      .replace(/\s*\n\s*/g, '\n')
      .trim();
  }
  // Chrome 朗讀太長的句子會中斷，切成約 120 字以內的段落依序朗讀
  function chunks(t) {
    var out = [];
    t.replace(/([。；！？\n])/g, '$1\u0000').split('\u0000').forEach(function (s) {   // 不用 lookbehind，舊版 Safari 才不會整支程式失效
      s = s.trim(); if (!s) return;
      if (!/[。；！？，、：]$/.test(s)) s += '。';   // 小標題、條列結尾補句號，朗讀時才會停頓
      while (s.length > 120) { out.push(s.slice(0, 120)); s = s.slice(120); }
      if (out.length && out[out.length - 1].length + s.length < 120) out[out.length - 1] += s; else out.push(s);
    });
    return out;
  }
  function setBtn(b, on) {
    if (!b) return;
    b.classList.toggle('on', on);
    b.innerHTML = on ? '⏹ 停止' : '🔊 朗讀';
    b.setAttribute('aria-pressed', on ? 'true' : 'false');
  }
  function stopSpeaking() {
    speakSeq++;
    if (TTS) window.speechSynthesis.cancel();
    setBtn(speakingBtn, false); speakingBtn = null;
  }
  function speak(box, btn) {
    stopSpeaking();
    var parts = chunks(textForSpeech(box)), seq = speakSeq;
    if (!parts.length) return;
    speakingBtn = btn; setBtn(btn, true);
    setTimeout(function () {             // cancel 後立刻 speak 在 Chrome 偶爾會無聲，稍等一下
      if (seq !== speakSeq) return;
      parts.forEach(function (p, i) {
        var u = new SpeechSynthesisUtterance(p);
        u.lang = 'zh-TW'; u.rate = 0.95;
        if (voice) u.voice = voice;
        if (i === parts.length - 1) u.onend = function () { if (seq === speakSeq) { setBtn(btn, false); speakingBtn = null; } };
        u.onerror = function (e) { if (seq === speakSeq && e.error !== 'interrupted' && e.error !== 'canceled') { setBtn(btn, false); speakingBtn = null; } };
        window.speechSynthesis.speak(u);
      });
    }, 60);
  }

  // AI 回答出現時（app.js 動態產生 .ai-box）加上「🔊 朗讀」按鈕；開啟自動朗讀時直接念
  function decorate(isNew) {
    if (!TTS || !thread) return;
    Array.prototype.forEach.call(thread.querySelectorAll('.ai-box'), function (box) {
      if (box.dataset.mm) return;
      box.dataset.mm = '1';
      var title = box.querySelector('.ai-title'); if (!title || !box.querySelector('.answer')) return;
      var b = document.createElement('button');
      b.type = 'button'; b.className = 'btn btn-ghost btn-sm mm-speak'; b.title = '朗讀這則回答（再按一次停止）';
      setBtn(b, false);
      b.onclick = function () { if (speakingBtn === b) stopSpeaking(); else speak(box, b); };
      title.appendChild(b);
      var qa = box.closest('.qa'), who = qa && qa.querySelector('.who');
      var hist = !!who && /歷史紀錄/.test(who.textContent);   // 從「我的提問紀錄」重開的回答不自動朗讀
      if (isNew && autoBox.checked && !hist) speak(box, b);
    });
  }
  if (thread && window.MutationObserver) new MutationObserver(function () { decorate(true); }).observe(thread, { childList: true, subtree: true });
  // 「新的提問」或送出新問題時停止朗讀
  var nc = $('newChat'); if (nc) nc.addEventListener('click', stopSpeaking);
  sendBtn.addEventListener('click', function () { if (!recOn) stopSpeaking(); });

  /* =================================================================
   * 3. 拍照查標準
   * ================================================================= */
  var TF_URL = 'https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.8.1/dist/transformers.min.js';
  var MODEL_ID = 'Xenova/clip-vit-base-patch32';
  var MIN_CONF = 0.25;                    // 最高信心低於 25% → 不猜
  var clfPromise = null, onProgress = null;

  // 第一次使用時才載入 transformers.js 與 CLIP 模型（約 150MB，下載後由瀏覽器 Cache Storage 快取）
  function loadClassifier() {
    if (!clfPromise) {
      var files = {};
      clfPromise = import(TF_URL).then(function (tf) {
        tf.env.allowLocalModels = false;  // 模型一律從 Hugging Face Hub 下載，不找本機 /models/
        return tf.pipeline('zero-shot-image-classification', MODEL_ID, {
          dtype: 'q8', device: 'wasm',
          progress_callback: function (p) {
            if (p.status !== 'progress' || !p.file || !p.total) return;
            files[p.file] = { l: p.loaded || 0, t: p.total };
            var L = 0, T = 0; for (var k in files) { L += files[k].l; T += files[k].t; }
            if (onProgress) onProgress(L, T);
          }
        });
      }).catch(function (e) { clfPromise = null; throw e; });
    }
    return clfPromise;
  }

  /**
   * 圖片 → 辨識結果
   *
   * 示範版：在瀏覽器內用 CLIP 做 zero-shot 分類。候選描述來自 sample_standards.json 的
   *   clip_labels（每份標準 4～6 句英文）與 clip_other_labels（與護理無關的照片，例如人臉、
   *   文件、食物），照片不會離開瀏覽器。CLIP 對所有候選描述的機率加總為 1，
   *   同一份標準的描述機率相加 = 該標準的信心。
   *
   * 正式版：改為把圖片 POST 到院內 ASP（例如 S_coach_api.asp?act=img），由院內 Ollama
   *   視覺模型（支援 images 的多模態模型）辨識照片內容並對應到護理標準，回傳與本函式
   *   相同格式的結果；照片只在院內網路傳輸，不送到外部服務。
   *
   * @param {File|Blob} file 使用者拍攝或選擇的圖片
   * @returns {Promise<{top:Array<{code:string,title:string,score:number,label:string}>,
   *                    best:Object|null, other:number, unsure:boolean}>}
   *   top：信心最高的前 3 份標準；unsure：最高信心 < 25% 或「與護理無關」的機率更高
   */
  function recognizeImage(file) {
    return DEMO.ready.then(function (data) {
      var labels = [], owner = [];
      (data.standards || []).forEach(function (s) {
        (s.clip_labels || []).forEach(function (t) { labels.push(t); owner.push(s); });
      });
      (data.clip_other_labels || []).forEach(function (t) { labels.push(t); owner.push(null); });
      if (!labels.length) throw new Error('範例資料中沒有 clip_labels');
      return loadClassifier().then(function (clf) {
        var url = URL.createObjectURL(file);
        return clf(url, labels, { hypothesis_template: 'a photo of {}.' })
          .then(function (out) { URL.revokeObjectURL(url); return out; },
                function (e) { URL.revokeObjectURL(url); throw e; });
      }).then(function (out) {
        var by = {}, other = 0;
        out.forEach(function (o) {
          var s = owner[labels.indexOf(o.label)];
          if (!s) { other += o.score; return; }
          var r = by[s.code] || (by[s.code] = { code: s.code, title: s.title.replace(/（示範）/g, ''), score: 0, label: o.label });
          r.score += o.score;
        });
        var top = Object.keys(by).map(function (k) { return by[k]; }).sort(function (a, b) { return b.score - a.score; });
        var best = top[0] || null;
        return { top: top.slice(0, 3), best: best, other: other, unsure: !best || best.score < MIN_CONF || other > best.score };
      });
    });
  }
  window.recognizeImage = recognizeImage;   // 方便測試與日後替換

  var camBtn = $('mmCam'), fileIn = $('mmFile'), panel = $('mmPanel'), thumb = $('mmThumb'), res = $('mmRes'), thumbUrl = null, camSeq = 0;
  camBtn.onclick = function () { fileIn.click(); };   // 手機會出現「拍照／相簿」選項，電腦則是選檔
  $('mmClose').onclick = closePanel;
  function closePanel() {
    camSeq++; panel.hidden = true; res.innerHTML = '';
    if (thumbUrl) { URL.revokeObjectURL(thumbUrl); thumbUrl = null; }
    thumb.removeAttribute('src');
  }
  fileIn.onchange = function () {
    var f = fileIn.files && fileIn.files[0];
    fileIn.value = '';                    // 同一張圖可以再選一次
    if (!f) return;
    if (!/^image\//.test(f.type)) { toast('請選擇圖片檔'); return; }
    var seq = ++camSeq;
    if (thumbUrl) URL.revokeObjectURL(thumbUrl);
    thumbUrl = URL.createObjectURL(f); thumb.src = thumbUrl;
    panel.hidden = false;
    res.innerHTML = '<div class="mm-st" id="mmSt">辨識中…</div>';
    var shown = false;
    onProgress = function (L, T) {
      if (seq !== camSeq) return;
      var pct = Math.min(100, Math.round(L / T * 100));
      if (!shown) { shown = true; res.innerHTML = '<div class="mm-st" id="mmSt"></div><div class="mm-prog"><i id="mmBar"></i></div><div class="note">第一次使用需下載辨識模型，之後會使用瀏覽器快取。</div>'; }
      $('mmSt').textContent = pct >= 100 ? '模型下載完成，辨識中…' : '下載圖片辨識模型 ' + pct + '%（' + (L / 1048576).toFixed(0) + ' / ' + (T / 1048576).toFixed(0) + ' MB）';
      $('mmBar').style.width = pct + '%';
    };
    var t0 = Date.now();
    recognizeImage(f).then(function (r) {
      if (seq !== camSeq) return;
      showResult(r, Date.now() - t0);
    }).catch(function (e) {
      if (seq !== camSeq) return;
      res.innerHTML = '<div class="banner err"><div>無法完成圖片辨識（' + esc(e && e.message || e) + '）。第一次使用需連網下載模型；也可以直接打字提問。</div></div>';
    });
  };

  function showResult(r, ms) {
    var h = '<div class="mm-st">圖片辨識結果（前 3 名）</div><ul class="mm-top">';
    r.top.forEach(function (x, i) {
      var pct = Math.round(x.score * 100);
      h += '<li class="' + (i === 0 && !r.unsure ? 'first' : '') + '" title="最接近的描述：' + esc(x.label) + '"><b>' + esc(x.title) + '<span class="note"> ' + esc(x.code) + '</span></b>' +
           '<span class="bar"><i style="width:' + pct + '%"></i></span><span class="pct">' + pct + '%</span></li>';
    });
    h += '</ul>';
    if (r.unsure) {
      h += '<div class="banner warn"><div>看不太出來，請換一張或直接打字提問。</div></div>';
    } else {
      var q = '這張照片相關的是【' + r.best.title + '】，請問操作重點？';
      setQuestion(q);
      h += '<div class="note" style="margin-top:6px">已把問題帶入輸入框，確認或修改後再按「詢問助理」。</div>';
      setTimeout(function () { qEl.focus(); }, 30);
    }
    res.innerHTML = h;
    res.setAttribute('data-ms', ms);
  }
})();
