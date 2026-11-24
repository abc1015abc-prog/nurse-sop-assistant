/* 護理標準查詢助理  示範版畫面程式（app.js）
 * 由 src/asp/S_coach.asp 的 <script> 原樣複製；與正式版不同之處都標有 [DEMO]。
 * 正式版前端修改後，請把同樣的修改同步到這裡。需在 mock_api.js 載入資料後才執行。 */
var STD = DEMO.std();   // [DEMO] 正式版由伺服器從 std_list 輸出   // [編號, 名稱, 分類, 影片, 已建索引, 修改日, 單位]
var API = null;   // [DEMO] 正式版為 'S_coach_api.asp'；示範版查詢由 mock_api.js 處理
var TYPES = DEMO.types;   // [DEMO]   // [代碼, 名稱]，依編號開頭分類
// 次專科標準編號：單位-類別-流水號（ICU-CPT-01）；護理部共通的 CP062-1 不算
function susType(code){var p=String(code||'').toUpperCase().split('-');if(p.length<3)return '';for(var i=0;i<TYPES.length;i++)if(TYPES[i][0]===p[1])return p[1];return '';}
function unitOf(code){return susType(code)?String(code).split('-')[0].toUpperCase():'';}
function unitTag(code){var u=unitOf(code);return u?'<span class="tag sus" title="病房單位次專科標準">次專科・'+esc(u)+'</span>':'';}
function srcOk(code){if(!SRC)return true;var u=unitOf(code);if(SRC==='ND')return !u;if(SRC==='SUS')return !!u;return u===SRC.slice(4);}
function typeOf(code){var sg=susType(code);if(sg)return sg;code=String(code||'').toUpperCase();for(var i=0;i<TYPES.length;i++){var k=TYPES[i][0];if(code.indexOf(k)===0&&!(k==='CP'&&code.indexOf('CPT')===0))return k;}return '';}
function typeName(k){for(var i=0;i<TYPES.length;i++)if(TYPES[i][0]===k)return TYPES[i][1];return k||'其他';}
function typeRank(k){for(var i=0;i<TYPES.length;i++)if(TYPES[i][0]===k)return i;return 99;}

/* ---------- 查詢範圍（類別篩選，可複選；記在瀏覽器） ---------- */
var SCOPE=[], SRC='';
try{SRC=localStorage.getItem('sc_src')||'';}catch(e){SRC='';}
try{SCOPE=JSON.parse(localStorage.getItem('sc_scope')||'[]')||[];}catch(e){SCOPE=[];}
function scopeStr(){return SCOPE.join(',');}
function drawScope(){
  var cnt={}, units={}; STD.forEach(function(s){var u=unitOf(s[0]);if(u)units[u]=(units[u]||0)+1;if(!srcOk(s[0]))return;var k=typeOf(s[0]);cnt[k]=(cnt[k]||0)+1;});
  var ul=Object.keys(units).sort();
  if(SRC.indexOf('SUS:')===0&&!units[SRC.slice(4)]) SRC='';
  var sel=ul.length?'<select class="scope-sel" data-src title="標準來源"><option value="">全部來源</option><option value="ND"'+(SRC==='ND'?' selected':'')+'>護理部共通標準</option>'+
    '<option value="SUS"'+(SRC==='SUS'?' selected':'')+'>單位次專科（全部單位）</option>'+
    ul.map(function(u){return '<option value="SUS:'+esc(u)+'"'+(SRC==='SUS:'+u?' selected':'')+'>次專科・'+esc(u)+'（'+units[u]+'）</option>';}).join('')+'</select>':'';
  var html='<span class="scope-l">查詢範圍</span>'+sel+'<button type="button" class="chip'+(SCOPE.length?'':' on')+'" data-t="">全部標準</button>'+
    TYPES.filter(function(t){return cnt[t[0]];}).map(function(t){return '<button type="button" class="chip'+(SCOPE.indexOf(t[0])>=0?' on':'')+'" data-t="'+t[0]+'" title="'+t[0]+'">'+esc(t[1])+'<span class="n">'+cnt[t[0]]+'</span></button>';}).join('');
  $$('[data-scope]').forEach(function(box){
    box.innerHTML=html;
    $$('.chip',box).forEach(function(b){b.onclick=function(){
      var t=b.dataset.t;
      if(!t) SCOPE=[]; else { var i=SCOPE.indexOf(t); if(i>=0) SCOPE.splice(i,1); else SCOPE.push(t); }
      try{localStorage.setItem('sc_scope',JSON.stringify(SCOPE));}catch(e){}
      drawScope();
    };});
    var sl=box.querySelector('[data-src]');
    if(sl) sl.onchange=function(){SRC=sl.value;try{localStorage.setItem('sc_src',SRC);}catch(e){}drawScope();if(window.scOvDraw)window.scOvDraw();};
  });
}
var AI_ON = DEMO.aiOn;   // [DEMO] 示範版以規則模擬 AI

function $(s,r){return (r||document).querySelector(s);}
function $$(s,r){return Array.prototype.slice.call((r||document).querySelectorAll(s));}
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}
function pdfUrl(code,page){return DEMO.pdfUrl(code,page);}   // [DEMO] 正式版：S_coach_pdf.asp?code=…#page=…
function hl(html,terms){
  terms=(terms||[]).filter(function(t){return t&&t.length>=2;}).sort(function(a,b){return b.length-a.length;});
  if(!terms.length) return html;
  var pats=terms.map(function(t){return esc(t).split('').map(function(c){return c.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}).join('\\s*');});
  return html.replace(new RegExp('('+pats.join('|')+')','gi'),'<mark>$1</mark>');
}
function post(data){
  if((data.act==='ask'||data.act==='search')&&data.src===undefined) data.src=SRC;
  return DEMO.post(data);   // [DEMO] 正式版：fetch(API, POST x-www-form-urlencoded)
}
var toastT;
function toast(m){var t=$('#toast');t.textContent=m;t.classList.add('show');clearTimeout(toastT);toastT=setTimeout(function(){t.classList.remove('show');},2400);}
function icon(id,s){s=s||16;return '<svg width="'+s+'" height="'+s+'"><use href="#'+id+'"/></svg>';}

/* ---------- 分頁（同頁切換，網址 #chat / #kw / #all 可直接連結） ---------- */
var VIEWS={chat:1,kw:1,all:1};
function showView(v,noHash){
  if(!VIEWS[v]) v='chat';
  $$('.nav-item[data-view]').forEach(function(b){var on=b.dataset.view===v;b.classList.toggle('active',on);if(on)b.setAttribute('aria-current','page');else b.removeAttribute('aria-current');});
  $$('.view').forEach(function(s){s.classList.toggle('on',s.id==='v-'+v);});
  $('#composer').classList.toggle('hide',v!=='chat');
  if(window.scNav) window.scNav(false);
  if(!noHash && history.replaceState) history.replaceState(null,'','#'+v);
  document.title=({chat:'問助理',kw:'關鍵字查詢',all:'標準總覽'})[v]+'｜護理標準查詢助理';
  window.scrollTo({top:0});
  if(v==='kw') setTimeout(function(){$('#kw').focus();},50);
}
$$('.nav-item[data-view]').forEach(function(b){b.addEventListener('click',function(e){e.preventDefault();showView(b.dataset.view);});});
window.addEventListener('hashchange',function(){showView(location.hash.slice(1),true);});
window.scKwSearch=function(q){$('#kw').value=q;showView('kw');kwSearch();};

/* ---------- Stepper ---------- */
var STEP_N=['輸入問題','檢索標準','AI 整理','核對出處'];
function setStep(cur,opt){
  opt=opt||{};
  $$('#steps .step').forEach(function(s,i){
    var done=i<cur||(opt.allDone);
    s.classList.toggle('done',done); s.classList.toggle('current',!opt.allDone&&i===cur); s.classList.toggle('busy',!!opt.busy&&i===cur);
    s.classList.toggle('skip',i===2&&!!opt.skipAI);
    s.querySelector('.step-dot').innerHTML=done?icon('i-check',18):(i+1);
  });
  var k=opt.allDone?3:cur;
  $('#mStep').textContent=STEP_N[k]+(opt.busy?'中…':''); $('#mStepNo').textContent=(k+1)+' / 4'; $('#mBar').style.width=((k+1)*25)+'%';
}

/* ---------- 問助理 ---------- */
var thread=$('#thread'), qEl=$('#q'), sendBtn=$('#send'), askSeq=0;
qEl.addEventListener('keydown',function(e){if(e.key==='Enter'&&!e.shiftKey&&!e.isComposing){e.preventDefault();ask();}});
qEl.addEventListener('input',function(){qEl.style.height='auto';qEl.style.height=Math.min(qEl.scrollHeight,140)+'px';});
sendBtn.onclick=function(){ask();};
$$('.sugg button').forEach(function(b){b.onclick=function(){qEl.value=b.dataset.q;ask();};});
$('#newChat').onclick=function(){thread.innerHTML='';$('#welcome').style.display='';setStep(0);resetSummary();qEl.focus();window.scrollTo({top:0,behavior:'smooth'});};

function badge(cls,txt,pulse){return '<span class="badge '+cls+(pulse?' pulse':'')+'">'+txt+'</span>';}

function ask(){
  var q=qEl.value.trim(); if(!q) return;
  qEl.value=''; qEl.style.height='auto';
  sendBtn.disabled=true; sendBtn.classList.add('loading');
  $('#welcome').style.display='none';
  var id='a'+(++askSeq);
  var wrap=document.createElement('article'); wrap.className='qa';
  wrap.innerHTML='<div class="q"><div>'+esc(q)+'</div></div>'+
    '<section class="card a"><div class="a-head"><div class="who"><i>'+icon('i-spark',16)+'</i>助理回覆</div><div class="badges" id="'+id+'-b">'+badge('info','檢索中',true)+'</div></div>'+
    '<div class="a-body" id="'+id+'-body"><div class="loading-row"><span class="dots"><i></i><i></i><i></i></span>正在比對護理標準全文…</div></div></section>';
  thread.appendChild(wrap); wrap.scrollIntoView({behavior:'smooth',block:'start'});
  setStep(1,{busy:true});
  var body=document.getElementById(id+'-body'), bdg=document.getElementById(id+'-b');

  // 第 1 步：馬上取得標準原文；找不到時請 AI 把口語／縮寫轉成標準用詞再找一次
  var qs=q;
  post({act:'ask',q:q,ai:0,types:scopeStr()}).then(function(d){
    var weak=d.ok&&(d.status==='nohit'||d.status==='noterm'||d.status==='low'||!(d.sources&&d.sources.length));
    if(!weak||!AI_ON||d.qa) return d;
    body.innerHTML='<div class="loading-row"><span class="dots"><i></i><i></i><i></i></span>標準中沒有直接對到，正在請 AI 判讀關鍵詞後再找一次…</div>';
    return post({act:'kwx',q:q}).then(function(k){
      if(!k.ok||!k.kw) return d;
      return post({act:'ask',q:k.kw,orig:q,replog:d.logId||0,ai:0,types:scopeStr()}).then(function(d2){
        var better=d2.ok&&d2.sources&&d2.sources.length&&(!(d.sources&&d.sources.length)||(d2.cover||0)>(d.cover||0));
        if(better){qs=k.kw;return d2;}
        d.kwTried=k.kw; d.logId=d2.logId||d.logId; return d;
      });
    }).catch(function(){return d;});
  }).then(function(d){
    if(!d.ok){body.innerHTML='<div class="banner err">'+esc(d.err||'發生錯誤')+'</div>';bdg.innerHTML=badge('danger','錯誤');setStep(0);return;}
    var qaOk=!!(d.qa&&d.qa.status==='ok');
    var pending=AI_ON&&d.status==='ok'&&!qaOk;
    body.innerHTML=renderAsk(d,id,pending);
    bindSrc(body,id);
    bdg.innerHTML=statusBadges(d,pending);
    updateSummary(d);
    var rb=body.querySelector('[data-runai]');
    if(rb) rb.onclick=function(){rb.parentNode.innerHTML='';runAI(d,id,q,qs,bdg);};
    if(!d.sources||!d.sources.length){setStep(d.qa?3:0,d.qa?{allDone:true,skipAI:true}:{});return;}
    if(!pending){setStep(3,{allDone:true,skipAI:true});return;}
    runAI(d,id,q,qs,bdg);
  }).catch(function(e){body.innerHTML='<div class="banner err">'+esc(e.message)+'</div>';bdg.innerHTML=badge('danger','錯誤');setStep(0);})
  .then(function(){sendBtn.disabled=false;sendBtn.classList.remove('loading');qEl.focus();});
}

// 第 2 步：背景請 AI 依同樣的原文整理（離開頁面也會存到提問紀錄）
function runAI(d,id,q,qs,bdg){
  setStep(2,{busy:true});
  var slot=document.getElementById(id+'-ai');
  slot.innerHTML='<div class="ai-wait"><span class="dots"><i></i><i></i><i></i></span><span>院內 AI 正在依下方標準整理重點，約需 1 分鐘，可以先點「查看原文」或「開啟 PDF」。離開本頁也沒關係，整理好會自動存到「我的提問紀錄」</span><span class="shim"></span></div>';
  bdg.innerHTML=statusBadges(d,true);
  post({act:'ask',q:qs,orig:q,ai:1,nolog:1,logid:d.logId||0,types:scopeStr()}).then(function(d2){
    var html=d2.ok?aiBlock(d2,id):'';
    slot.innerHTML=html||'<div class="banner info">'+icon('i-info')+'<div>AI 整理暫時無法使用，請點下方標準的「查看原文」或「開啟 PDF」。</div></div>';
    bindSrc(slot,id);
    var ok=!!slot.querySelector('.ai-box');
    bdg.innerHTML=(d.qa?badge('success','已審核答案')+' ':'')+(ok?badge('success','AI 已整理'):badge('neutral','僅原文'));
    setStep(3,{allDone:true,skipAI:!ok});
    if(ok) toast('AI 已依標準原文整理完成');
  }).catch(function(e){
    slot.innerHTML='<div class="banner info">'+icon('i-info')+'<div>AI 整理暫時無法使用（'+esc(e.message)+'），請直接參考下方原文。</div></div>';
    bdg.innerHTML=badge('neutral','僅原文'); setStep(3,{allDone:true,skipAI:true});
  });
}

// 已審核答案的簡易排版（### 小標、- 條列、**粗體**）
function mdLite(t){
  var html='', inList=null;
  String(t||'').split(/\n/).forEach(function(line){
    var raw=line.trim();
    if(!raw){if(inList){html+='</'+inList+'>';inList=null;}return;}
    var li=raw.match(/^[-*•]\s+(.*)$/), ol=raw.match(/^\d+[.、)]\s*(.*)$/), hd=raw.match(/^#{1,4}\s+(.*)$/);
    var body=esc(li?li[1]:ol?ol[1]:hd?hd[1]:raw).replace(/\*\*(.+?)\*\*/g,'<b>$1</b>');
    if(li||ol){var tag=li?'ul':'ol';if(inList!==tag){if(inList)html+='</'+inList+'>';html+='<'+tag+'>';inList=tag;}html+='<li>'+body+'</li>';return;}
    if(inList){html+='</'+inList+'>';inList=null;}
    html+=(hd||/^【[^】]+】/.test(raw))?'<div class="ai-h">'+body+'</div>':'<div>'+body+'</div>';
  });
  if(inList) html+='</'+inList+'>';
  return html;
}
// 臨床教師已審核的答案
function qaCard(qa,q){
  if(!qa) return '';
  var re=qa.status==='recheck';
  var h='<div class="qa-card'+(re?' recheck':'')+'"><div class="qa-h"><span class="qa-ico">'+icon('i-shield',16)+'</span><b>臨床教師已審核的答案</b>'+
        '<span class="note">審核 '+esc(qa.reviewer)+'・'+esc(String(qa.rtime).split(' ')[0])+'</span></div>';
  if(re) h+='<div class="banner warn" style="margin:0 0 10px">'+icon('i-info')+'<div>引用的標準已改版，這個答案正在等待臨床教師重新確認，請以最新版標準原文為準。</div></div>';
  if(q&&qa.question&&qa.question.replace(/\s/g,'')!==q.replace(/\s/g,'')) h+='<div class="qa-sim">相近問題：「'+esc(qa.question)+'」</div>';
  h+='<div class="answer">'+mdLite(qa.answer)+'</div>';
  if(qa.refs&&qa.refs.length){
    h+='<div class="qa-refs"><span class="qa-rk">出處</span>';
    qa.refs.forEach(function(r){
      h+='<a class="qa-ref" target="_blank" href="'+pdfUrl(r.code,(r.pages||[])[0])+'"><span class="tag code">'+esc(r.code)+'</span>'+esc(r.title||'')+
         (r.pages&&r.pages.length?'<span class="note">第 '+r.pages.join('、')+' 頁</span>':'')+icon('i-file',14)+'</a>';
    });
    h+='</div>';
  }
  return h+'</div>';
}
// 回答回饋列
function fbBar(d){
  if(!d.logId) return '';
  var rs=['找錯標準','內容與標準不符','不完整','看不懂','其他'];
  return '<div class="fb" data-log="'+d.logId+'" data-qa="'+(d.qa?d.qa.id:0)+'"><div class="fb-row"><span class="fb-q">這個回答有幫助嗎？</span>'+
    '<button type="button" class="fb-b up" data-r="1">'+icon('i-up',16)+'有幫助</button>'+
    '<button type="button" class="fb-b down" data-r="-1">'+icon('i-down',16)+'需要改進</button><span class="fb-done note"></span></div>'+
    '<div class="fb-form" hidden><div class="fb-l">哪裡需要改進？（可複選）</div><div class="chips">'+rs.map(function(r){return '<button type="button" class="chip" data-reason="'+r+'">'+r+'</button>';}).join('')+'</div>'+
    '<textarea class="input" rows="2" maxlength="250" placeholder="補充說明（選填）：例如正確應參考哪一份標準、哪裡寫錯"></textarea>'+
    '<div class="fb-acts"><button type="button" class="btn btn-primary btn-sm" data-send>送出回饋</button><button type="button" class="btn btn-ghost btn-sm" data-cancel>取消</button></div></div></div>';
}
function bindFb(root){
  $$('.fb',root).forEach(function(f){
    var form=f.querySelector('.fb-form'), done=f.querySelector('.fb-done');
    function send(r,reason,note){
      return post({act:'fb',logid:f.dataset.log,qaid:f.dataset.qa,rating:r,reason:reason||'',note:note||''}).then(function(x){
        if(!x.ok){toast(x.err||'送出失敗');return;}
        $$('.fb-b',f).forEach(function(b){b.classList.toggle('on',b.dataset.r===String(r));});
        form.hidden=true;
        done.textContent=r===1?'謝謝回饋！':'已送出，臨床教師會檢視並修正答案，謝謝！';
      }).catch(function(e){toast(e.message);});
    }
    f.querySelector('.fb-b.up').onclick=function(){send(1);};
    f.querySelector('.fb-b.down').onclick=function(){form.hidden=false;done.textContent='';form.querySelector('textarea').focus();};
    $$('[data-reason]',f).forEach(function(c){c.onclick=function(){c.classList.toggle('on');};});
    f.querySelector('[data-cancel]').onclick=function(){form.hidden=true;};
    f.querySelector('[data-send]').onclick=function(){
      var rsn=$$('[data-reason].on',f).map(function(c){return c.dataset.reason;}).join('、');
      var note=form.querySelector('textarea').value.trim();
      if(!rsn&&!note){toast('請選擇原因或填寫說明');return;}
      send(-1,rsn,note);
    };
  });
}

// 開啟「我的提問紀錄」裡存下來的回答
function openLog(lid){
  $('#welcome').style.display='none';
  var id='a'+(++askSeq);
  var wrap=document.createElement('article'); wrap.className='qa';
  wrap.innerHTML='<div class="q"><div>…</div></div><section class="card a"><div class="a-head"><div class="who"><i>'+icon('i-spark',16)+'</i>助理回覆（歷史紀錄）</div><div class="badges" id="'+id+'-b"></div></div><div class="a-body" id="'+id+'-body"><div class="loading-row"><span class="dots"><i></i><i></i><i></i></span>讀取紀錄中…</div></div></section>';
  thread.appendChild(wrap);
  var body=document.getElementById(id+'-body'), bdg=document.getElementById(id+'-b');
  post({act:'getlog',id:lid}).then(function(g){
    if(!g.ok){body.innerHTML='<div class="banner err">'+esc(g.err||'讀取失敗')+'</div>';return;}
    wrap.querySelector('.q div').textContent=g.q;
    bdg.innerHTML='<span class="note">'+esc(g.time)+'</span>';
    if(!g.resp){
      body.innerHTML=g.result==='pending'
        ?'<div class="banner info">'+icon('i-info')+'<div>AI 還在整理這一題，請稍後重新整理此頁。</div></div>'
        :'<div class="banner warn">'+icon('i-info')+'<div>這筆紀錄沒有保存回答內容（可能是更新前的舊紀錄或標準中找不到），可以按下方重新提問。</div></div>';
      body.innerHTML+='<a class="btn btn-primary btn-sm" href="index.html?q='+encodeURIComponent(g.q)+'#chat">再問一次</a>';
      return;
    }
    if(!g.resp.logId) g.resp.logId=+lid;
    body.innerHTML=renderAsk(g.resp,id,false,true); bindSrc(body,id); updateSummary(g.resp);
    bdg.innerHTML+=' '+(g.resp.ai?badge('success','AI 已整理'):badge('neutral','標準原文'));
    setStep(3,{allDone:true,skipAI:!g.resp.ai});
  }).catch(function(e){body.innerHTML='<div class="banner err">'+esc(e.message)+'</div>';});
}

function statusBadges(d,pending){
  var qb=d.qa?badge(d.qa.status==='ok'?'success':'warning',d.qa.status==='ok'?'已審核答案':'審核答案待確認')+' ':'';
  if(d.status==='noindex') return badge('danger','尚未建立索引');
  if(d.qa&&(!d.sources||!d.sources.length)) return qb;
  if(!d.sources||!d.sources.length) return badge('danger','標準中找不到');
  var b=qb;
  if(d.status==='low') b+=badge('warning','可能相關');
  b+=pending?badge('info','AI 整理中',true):badge('neutral','標準原文');
  return b;
}

// AI 整理結果
function aiBlock(d,id){
  var aiText='';
  if(d.ai){
    if(d.ai.message&&d.ai.message.content!=null) aiText=String(d.ai.message.content);                         // Ollama
    else if(d.ai.choices&&d.ai.choices[0]) aiText=String((d.ai.choices[0].message||{}).content||'');           // OpenAI 相容
    else if(d.ai.content) aiText=d.ai.content.filter(function(c){return c.type==='text';}).map(function(c){return c.text;}).join('\n'); // Claude
    aiText=aiText.replace(/<think>[\s\S]*?<\/think>/g,'').replace(/^[\s\S]*?<\/think>/,'').trim();
  }
  if(aiText) return '<div class="ai-box arrive"><div class="ai-title"><svg class="done-check" viewBox="0 0 20 20"><circle cx="10" cy="10" r="9"/><path d="m6 10.5 2.6 2.6L14 7.6"/></svg>AI 依下方標準原文整理的重點</div>'+
                    '<div class="answer">'+renderAI(aiText,d.sources,id)+'</div></div>';
  if(d.ai&&d.ai.message&&d.ai.message.thinking) return '<div class="banner info">'+icon('i-info')+'<div>AI 這次沒有產生答案，請直接參考下方原文。</div></div>';
  if(d.aiErr) return '<div class="banner info">'+icon('i-info')+'<div>AI 整理暫時無法使用（'+esc(d.aiErr)+'），請直接參考下方原文。</div></div>';
  return '';
}

function renderAsk(d,id,pending,hist){
  var h=qaCard(d.qa,d.q);
  if(d.status==='noindex') return h+'<div class="banner warn">'+icon('i-info')+'<div>標準全文索引尚未建立，請通知管理者執行「建立索引」。</div></div>';
  if(d.status==='noterm'||d.status==='nohit'||!d.sources.length)
    return h+(d.qa?'':'<div class="banner warn">'+icon('i-info')+'<div><b>護理標準中找不到與這個問題相關的內容。</b>'+(d.kwTried?'（AI 判讀的關鍵詞「'+esc(d.kwTried)+'」也找不到）':'')+'<br>可以改用標準上的正式名稱（例如「留置導尿管護理」）、改用「關鍵字查詢」，或直接詢問臨床教師／單位護理長。</div></div>')+fbBar(d);
  if(d.status==='low')
    h+='<div class="banner warn">'+icon('i-info')+'<div>標準中沒有找到明確符合的內容，以下是「可能相關」的段落，請自行判斷或詢問臨床教師。</div></div>';
  h+=understandLine(d);
  // 最相關的標準（先給明確結論：答案在哪一份標準）
  if(d.primary&&d.primary.code){
    var pc=d.sources.filter(function(s){return s.code===d.primary.code;});
    h+='<div class="prim"><div class="prim-l"><span class="prim-k">最相關的標準</span><b>'+esc(d.primary.title)+'</b>'+
       '<div class="prim-m"><span class="tag code">'+esc(d.primary.code)+'</span>'+unitTag(d.primary.code)+(d.primary.cat?'<span class="tag">'+esc(d.primary.cat)+'</span>':'')+
       (d.intent?'<span class="badge info">已優先找「'+esc(d.intent)+'」章節</span>':'')+
       (pc.length?'<span class="note">相關段落在第 '+pc.map(function(s){return s.page;}).filter(function(v,i,a){return a.indexOf(v)===i;}).join('、')+' 頁</span>':'')+'</div></div>'+
       '<div class="prim-r"><a class="lnk primary" target="_blank" href="'+pdfUrl(d.primary.code,pc.length?pc[0].page:0)+'">'+icon('i-file',14)+'開啟標準 PDF</a>'+
       (d.primary.video?'<a class="lnk" target="_blank" href="'+esc(d.primary.video)+'">'+icon('i-play',14)+'教學影片</a>':'')+'</div></div>';
  }
  h+='<div id="'+id+'-ai">'+(pending?'<div class="ai-wait"><span class="dots"><i></i><i></i><i></i></span><span>院內 AI 正在依下方標準整理重點，約需 1 分鐘，可以先點「查看原文」或「開啟 PDF」。離開本頁也沒關係，整理好會自動存到「我的提問紀錄」</span><span class="shim"></span></div>':(!d.ai&&d.qa&&d.qa.status==='ok'&&AI_ON&&d.status==='ok'&&!hist?'<div class="ai-more"><button type="button" class="btn btn-secondary btn-sm" data-runai>'+icon('i-spark',16)+'也請 AI 依標準原文整理</button></div>':aiBlock(d,id)))+'</div>';
  h+=d.stds&&d.stds.length?renderStds(d,id):renderSources(d,id);
  var inSrc={}; d.sources.forEach(function(x){inSrc[x.code]=1;}); (d.stds||[]).forEach(function(x){inSrc[x.code]=1;});
  var rel=(d.related||[]).filter(function(r){return !inSrc[r.code];});
  if(rel.length){
    h+='<div class="sec-title" style="margin-top:16px">'+icon('i-book',14)+'其他可能相關的標準</div><div class="rel">';
    rel.forEach(function(r){h+='<a class="rel-i" target="_blank" href="'+pdfUrl(r.code)+'"><span class="tag code">'+esc(r.code)+'</span><span>'+esc(r.title)+'</span>'+icon('i-right',14)+'</a>';});
    h+='</div>';
  }
  return h+fbBar(d);
}

// 系統怎麼理解這個問題：辨識到的關鍵詞、查詢意圖、AI 判讀的關鍵詞
function understandLine(d){
  var c=d.concepts||[]; if(!c.length&&!d.kw) return '';
  var h='<div class="und"><span class="und-k">'+icon('i-search',14)+'查詢重點</span>';
  c.forEach(function(x){h+='<span class="tag'+(x.syn?' syn':'')+'"'+(x.syn?' title="已含同義詞、縮寫一起找"':'')+'>'+esc(x.w)+'</span>';});
  if(d.intent) h+='<span class="badge info">想知道：'+esc(d.intent)+'</span>';
  if(d.kw) h+='<span class="badge warning" title="原本的問法找不到，由 AI 改寫成標準用詞後查詢">AI 判讀：'+esc(d.kw)+'</span>';
  return h+'</div>';
}
// 符合的關鍵詞（去掉被較長詞包含的雙字詞）
function hitWords(d,st){
  var c=d.concepts||[], w=(st.hits||[]).map(function(i){return c[i]?c[i].w:'';}).filter(Boolean);
  return w.filter(function(x){return !w.some(function(y){return y!==x&&y.length>x.length&&y.indexOf(x)>=0;});});
}
// 找到的標準：預設收合，一行一份（名稱、編號、頁碼、符合度），依關鍵詞符合度排序
function renderStds(d,id){
  var bySrc={}; (d.sources||[]).forEach(function(s){bySrc[s.id]=s;});
  var h='<div class="sx-head"><div class="sec-title" style="margin:0">'+icon('i-file',14)+'找到 '+d.stds.length+' 份相關標準・依關鍵詞符合度排序</div>'+
        '<button type="button" class="btn btn-ghost btn-sm" data-allx>全部展開</button></div><div class="sx-list">';
  d.stds.forEach(function(st,i){
    // 原文段落：AI 引用的段落（有來源編號）＋最相關段落，依頁碼排序
    var seen={}, parts=[];
    (d.sources||[]).forEach(function(s){if(s.code===st.code){seen[s.id]=1;parts.push({n:s.n,page:s.page,text:s.text});}});
    (st.ex||[]).forEach(function(e){if(!seen[e.id]&&parts.length<4){seen[e.id]=1;parts.push({page:e.page,text:e.text});}});
    parts.sort(function(a,b){return a.page-b.page;});
    var pct=Math.round((+st.cov||0)*100), words=hitWords(d,st), prim=d.primary&&st.code===d.primary.code;
    h+='<div class="sx'+(prim?' prim-x':'')+'" data-code="'+esc(st.code)+'">'+
       '<div class="sx-row">'+
         '<span class="sx-no">'+(i+1)+'</span>'+
         '<div class="sx-main"><div class="sx-t"><b>'+esc(st.title)+'</b><span class="tag code">'+esc(st.code)+'</span>'+unitTag(st.code)+
           '<span class="tag">'+esc(st.typeName||typeName(st.type))+'</span>'+(prim?'<span class="badge info">最相關</span>':'')+'</div>'+
           '<div class="sx-m"><span class="sx-pg">第 '+st.pages.join('、')+' 頁</span>'+
           '<span class="sx-meter" title="問題關鍵詞在這份標準中的符合度"><i style="width:'+pct+'%"></i></span><span class="sx-pct">'+pct+'%</span>'+
           (words.length?'<span class="sx-hits">符合：'+words.map(esc).join('、')+'</span>':'')+'</div></div>'+
         '<div class="sx-acts"><button type="button" class="lnk" data-x aria-expanded="false">'+icon('i-right',14)+'查看原文</button>'+
           '<a class="lnk primary" target="_blank" href="'+pdfUrl(st.code,st.pages[0])+'">'+icon('i-file',14)+'開啟 PDF</a>'+
           (st.video?'<a class="lnk" target="_blank" href="'+esc(st.video)+'">'+icon('i-play',14)+'影片</a>':'')+'</div>'+
       '</div><div class="sx-body" hidden>';
    parts.forEach(function(p){
      h+='<div class="ex"'+(p.n?' id="'+id+'-s'+p.n+'"':'')+'><div class="ex-h">'+(p.n?'<span class="no">'+p.n+'</span>':'')+'第 '+p.page+' 頁'+
         (p.n?'<span class="badge neutral">AI 參考段落</span>':'')+
         '<a target="_blank" href="'+pdfUrl(st.code,p.page)+'">在 PDF 看這頁 ›</a></div>'+
         '<div class="txt">'+hl(esc(p.text),d.terms)+'</div><button class="more" hidden>展開全文</button></div>';
    });
    h+='</div></div>';
  });
  return h+'</div>';
}
function setupMore(root){
  $$('.ex',root).forEach(function(s){
    var t=s.querySelector('.txt'), b=s.querySelector('.more');
    if(t&&b&&!b.dataset.ok&&t.offsetParent&&t.scrollHeight>t.clientHeight+4){b.dataset.ok=1;b.hidden=false;b.onclick=function(){t.classList.toggle('open');b.textContent=t.classList.contains('open')?'收合':'展開全文';};}
  });
}
function openSx(sx,on){
  var b=sx.querySelector('[data-x]'), body=sx.querySelector('.sx-body');
  if(on===undefined) on=body.hidden;
  body.hidden=!on; sx.classList.toggle('open',on); b.setAttribute('aria-expanded',on);
  b.innerHTML=icon('i-right',14)+(on?'收合原文':'查看原文');
  if(on) setupMore(sx);
}

// 標準原文：同一類別放一起；同一份標準只出現一次，多個段落收在同一張卡片
function groupSources(sources){
  var groups=[], gi={};
  sources.forEach(function(s){
    var t=s.type||typeOf(s.code);
    if(!(t in gi)){gi[t]=groups.length;groups.push({type:t,name:s.typeName||typeName(t),stds:[],si:{}});}
    var g=groups[gi[t]];
    if(!(s.code in g.si)){g.si[s.code]=g.stds.length;g.stds.push({code:s.code,title:s.title,cat:s.cat,video:s.video,parts:[]});}
    g.stds[g.si[s.code]].parts.push(s);
  });
  return groups;
}
function renderSources(d,id){
  var groups=groupSources(d.sources), h='', nStd=0;
  groups.forEach(function(g){nStd+=g.stds.length;});
  h+='<div class="sec-title">'+icon('i-file',14)+'標準原文・'+nStd+' 份標準、'+d.sources.length+' 段</div>';
  groups.forEach(function(g){
    h+='<div class="tgroup"><div class="tgroup-h"><span class="tdot"></span>'+esc(g.name)+'<span class="note">'+esc(g.type)+'・'+g.stds.length+' 份</span></div><div class="srcs">';
    g.stds.forEach(function(st){
      var pages=st.parts.map(function(p){return p.page;}).filter(function(v,i,a){return a.indexOf(v)===i;});
      h+='<div class="src std"><div class="hd"><b>'+esc(st.title)+'</b><span class="tag code">'+esc(st.code)+'</span>'+
         '<span class="tag">第 '+pages.join('、')+' 頁</span>'+
         '<span class="acts"><a class="lnk primary" target="_blank" href="'+pdfUrl(st.code,pages[0])+'">'+icon('i-file',14)+'開啟 PDF</a>'+
         (st.video?'<a class="lnk" target="_blank" href="'+esc(st.video)+'">'+icon('i-play',14)+'教學影片</a>':'')+'</span></div>';
      st.parts.forEach(function(p){
        h+='<div class="ex" id="'+id+'-s'+p.n+'"><div class="ex-h"><span class="no">'+p.n+'</span>第 '+p.page+' 頁'+
           '<a target="_blank" href="'+pdfUrl(st.code,p.page)+'">看原文 ›</a></div>'+
           '<div class="txt">'+hl(esc(p.text),d.terms)+'</div><button class="more" hidden>展開全文</button></div>';
      });
      h+='</div>';
    });
    h+='</div></div>';
  });
  return h;
}

// AI 回答：簡易 Markdown、[來源n] 轉為可點出處、檢查數值是否出現在引用原文
function renderAI(t,sources,id){
  var byN={}; sources.forEach(function(s){byN[s.n]=s;});
  var allTxt=sources.map(function(s){return s.text;}).join('\n').replace(/\s+/g,'');
  var anyCite=/\[來源\s*\d+\]/.test(t), unverified=0, html='', inList=null;
  t.split(/\n/).forEach(function(line){
    var raw=line.trim(); if(!raw){ if(inList){html+='</'+inList+'>';inList=null;} return; }
    var ns=[], m, re=/\[來源\s*(\d+)\]/g; while((m=re.exec(raw))){ns.push(+m[1]);}
    var ref=ns.filter(function(n){return byN[n];}).map(function(n){return byN[n].text;}).join('\n').replace(/\s+/g,'');
    if(!ref) ref=allTxt;
    var li=raw.match(/^[-*•]\s+(.*)$/), ol=raw.match(/^\d+[.、)]\s*(.*)$/);
    if(li) raw=li[1]; else if(ol) raw=ol[1];
    var isH=/^#{1,4}\s+/.test(raw)||/^【[^】]+】\s*[A-Z]{2,3}[\w\-.]*/.test(raw);
    raw=raw.replace(/^#+\s*/,'');
    var body='', last=0, tk, tokRe=/\[來源\s*(\d+)\]|\d+(?:\.\d+)?/g;
    while((tk=tokRe.exec(raw))){
      body+=esc(raw.slice(last,tk.index));
      if(tk[1]){
        var n=+tk[1];
        body+=byN[n]?'<span class="cite" data-to="'+id+'-s'+n+'" title="跳到原文">'+esc(byN[n].code)+' p.'+byN[n].page+'</span>':'<span class="cite bad" title="無此出處">出處?</span>';
      }else if(!isH && !/[A-Za-z\-]$/.test(raw.slice(0,tk.index)) && ref.indexOf(tk[0])<0){
        unverified++; body+='<span class="numwarn" title="此數值在引用的標準原文中找不到，請核對原文">'+tk[0]+'⚠</span>';
      }else body+=tk[0];
      last=tokRe.lastIndex;
    }
    body=(body+esc(raw.slice(last))).replace(/\*\*(.+?)\*\*/g,'<b>$1</b>');
    if(isH&&!li&&!ol){
      if(inList){html+='</'+inList+'>';inList=null;}
      html+='<div class="ai-h">'+body+'</div>';
    }else if(li||ol){
      var tag=li?'ul':'ol';
      if(inList!==tag){ if(inList) html+='</'+inList+'>'; html+='<'+tag+'>'; inList=tag; }
      html+='<li>'+body+'</li>';
    }else{
      if(inList){html+='</'+inList+'>';inList=null;}
      html+='<div>'+body+'</div>';
    }
  });
  if(inList) html+='</'+inList+'>';
  var warn='';
  if(!anyCite && !/未提及/.test(t)) warn+='<div class="banner warn">'+icon('i-info')+'<div>此回答沒有附上出處，請以下方標準原文為準。</div></div>';
  if(unverified) warn+='<div class="banner warn">'+icon('i-info')+'<div>有 '+unverified+' 個數值（標 ⚠）在引用的原文中找不到，請務必核對原文。</div></div>';
  return warn+html;
}

function bindSrc(root,id){
  bindFb(root);
  $$('.cite[data-to]',root).forEach(function(c){
    c.onclick=function(){var el=document.getElementById(c.dataset.to); if(!el) return; var sx=el.closest('.sx'); if(sx) openSx(sx,true); el.scrollIntoView({behavior:'smooth',block:'center'}); el.classList.remove('flash'); void el.offsetWidth; el.classList.add('flash');};
  });
  $$('.sx [data-x]',root).forEach(function(b){b.onclick=function(){openSx(b.closest('.sx'));};});
  $$('[data-allx]',root).forEach(function(b){b.onclick=function(){
    var list=$$('.sx',b.closest('.sx-head').nextElementSibling), on=b.textContent==='全部展開';
    list.forEach(function(sx){openSx(sx,on);}); b.textContent=on?'全部收合':'全部展開';
  };});
  $$('.src:not(.std),.ex',root).forEach(function(s){
    var t=s.querySelector('.txt'), b=s.querySelector('.more');
    if(t&&b&&t.scrollHeight>t.clientHeight+4){b.hidden=false;b.onclick=function(){t.classList.toggle('open');b.textContent=t.classList.contains('open')?'收合':'展開全文';};}
  });
}

/* ---------- 右欄摘要 ---------- */
function resetSummary(){
  $('#ringPct').textContent='—'; $('#ringProg').style.strokeDashoffset=188.5;
  $('#sumHint').textContent='提問後顯示關鍵詞命中率與引用標準'; $('#sumBadge').innerHTML='';
  $('#refs').innerHTML='<li class="empty-s">尚無</li>';
}
function updateSummary(d){
  var c=Math.max(0,Math.min(1,+d.cover||0)), pct=d.sources&&d.sources.length?Math.round(c*100):0;
  $('#ringPct').textContent=pct+'%'; $('#ringProg').style.strokeDashoffset=188.5*(1-pct/100);
  $('#sumHint').textContent='問題關鍵詞在標準中的命中率';
  $('#sumBadge').innerHTML=!d.sources||!d.sources.length?badge('danger','找不到'):d.status==='low'?badge('warning','可能相關'):badge('success','明確相關');
  var seen={}, list=[], h='';
  if(d.primary&&d.primary.code) list.push(d.primary);
  (d.sources||[]).forEach(function(s){list.push(s);});
  (d.related||[]).forEach(function(s){list.push(s);});
  list=list.filter(function(s){if(seen[s.code])return false;seen[s.code]=1;return true;});
  var primT=d.primary?typeOf(d.primary.code):'';
  list.forEach(function(s,i){s._i=i;});
  list.sort(function(a,b){
    var ta=typeOf(a.code), tb=typeOf(b.code), ra=ta===primT?-1:typeRank(ta), rb=tb===primT?-1:typeRank(tb);
    if(ra!==rb) return ra-rb;
    return a._i-b._i;
  });
  var lastT=null;
  list.forEach(function(s){
    var t=typeOf(s.code);
    if(t!==lastT){h+='<li class="ref-h">'+esc(typeName(t))+'</li>';lastT=t;}
    h+='<li><span class="tag code">'+esc(s.code)+'</span><div class="tt"><b title="'+esc(s.title)+'">'+esc(s.title)+'</b>'+
       (d.primary&&s.code===d.primary.code?'<span style="color:var(--primary-700);font-weight:600">最相關</span>':'<span>'+esc(s.cat||'')+'</span>')+'</div>'+
       '<a class="lnk" target="_blank" href="'+pdfUrl(s.code)+'">PDF</a></li>';
  });
  $('#refs').innerHTML=h||'<li class="empty-s">無</li>';
}

/* ---------- 關鍵字查詢 ---------- */
var kwEl=$('#kw'), kwRes=$('#kwres'), kwBtn=$('#kwgo');
kwBtn.onclick=kwSearch;
kwEl.addEventListener('keydown',function(e){if(e.key==='Enter'&&!e.isComposing)kwSearch();});
function kwSearch(){
  var q=kwEl.value.trim(); if(!q) return;
  kwBtn.classList.add('loading'); kwBtn.disabled=true;
  kwRes.innerHTML='<div class="card empty" style="margin-top:20px"><span class="dots"><i></i><i></i><i></i></span><div style="margin-top:8px">搜尋中…</div></div>';
  post({act:'search',q:q,types:scopeStr()}).then(function(d){
    if(!d.ok){kwRes.innerHTML='<div class="banner err" style="margin-top:20px">'+esc(d.err)+'</div>';return;}
    if(!d.results.length){kwRes.innerHTML='<div class="card empty" style="margin-top:20px"><div class="hero-ico">'+icon('i-search',24)+'</div>找不到含「'+esc(q)+'」的護理標準<br><span class="note">試試較短的關鍵字或同義詞</span></div>';return;}
    var full=d.results.filter(function(r){return r.full;}), part=d.results.filter(function(r){return !r.full;});
    var h='<div class="result-meta">'+badge('info','找到 '+full.length+' 份標準')+(d.nk>1&&part.length?'<span class="note">另有 '+part.length+' 份只符合部分關鍵字</span>':'')+'</div><div class="res-list">';
    function one(r){
      var s='<section class="card res"><div class="t"><b>'+hl(esc(r.title),d.terms)+'</b><span class="tag code">'+esc(r.code)+'</span>'+(r.typeName?'<span class="tag">'+esc(r.typeName)+'</span>':'')+
        (r.cat?'<span class="tag">'+esc(r.cat)+'</span>':'')+(r.hits?badge('success','全文 '+r.hits+' 處'):'')+(r.indexed?'':badge('neutral','僅比對名稱'))+
        '<span class="acts"><a class="lnk primary" target="_blank" href="'+pdfUrl(r.code)+'">'+icon('i-file',14)+'開啟 PDF</a>'+
        (r.video?'<a class="lnk" target="_blank" href="'+esc(r.video)+'">'+icon('i-play',14)+'教學影片</a>':'')+'</span></div>';
      r.snips.forEach(function(x){s+='<div class="sn"><a target="_blank" href="'+pdfUrl(r.code,x.p)+'">第 '+x.p+' 頁 ›</a><span>'+hl(esc(x.t),d.terms)+'</span></div>';});
      return s+'</section>';
    }
    full.forEach(function(r){h+=one(r);});
    if(part.length&&d.nk>1){h+='<div class="sec-title" style="margin-top:12px">只符合部分關鍵字</div>';part.slice(0,20).forEach(function(r){h+=one(r);});}
    kwRes.innerHTML=h+'</div>';
  }).catch(function(e){kwRes.innerHTML='<div class="banner err" style="margin-top:20px">'+esc(e.message)+'</div>';})
  .then(function(){kwBtn.classList.remove('loading');kwBtn.disabled=false;});
}

/* ---------- 標準總覽 ---------- */
(function(){
  var cats={}, curCat='', curType='';
  var sc=document.createElement('div'); sc.className='chips'; sc.id='srcChips'; sc.style.marginBottom='8px';
  var hasSus=STD.some(function(s){return !!unitOf(s[0]);});
  var tc={}; STD.forEach(function(s){var k=typeOf(s[0]);tc[k]=(tc[k]||0)+1;});
  var tChips=$('#typeChips');
  tChips.innerHTML='<button class="chip on" data-t="">全部類別</button>'+TYPES.filter(function(t){return tc[t[0]];}).map(function(t){return '<button class="chip" data-t="'+t[0]+'">'+esc(t[1])+'<span class="n">'+tc[t[0]]+'</span></button>';}).join('');
  $$('.chip',tChips).forEach(function(b){b.onclick=function(){curType=b.dataset.t;$$('.chip',tChips).forEach(function(x){x.classList.toggle('on',x===b);});draw();};});
  STD.forEach(function(s){(s[2]||'未分類').split(',').forEach(function(c){cats[c]=(cats[c]||0)+1;});});
  var chips=$('#catChips'), names=Object.keys(cats).sort(function(a,b){return cats[b]-cats[a];});
  chips.innerHTML='<button class="chip on" data-c="">全部</button>'+names.map(function(c){return '<button class="chip" data-c="'+esc(c)+'">'+esc(c)+'</button>';}).join('');
  $$('.chip',chips).forEach(function(b){b.onclick=function(){curCat=b.dataset.c;$$('.chip',chips).forEach(function(x){x.classList.toggle('on',x===b);});draw();};});
  var flt=$('#flt'), tb=$('#lst');
  function draw(){
    var f=flt.value.trim().toLowerCase(), n=0, h='';
    STD.forEach(function(s){
      if(!srcOk(s[0])) return;
      if(curCat&&(s[2]||'未分類').split(',').indexOf(curCat)<0) return;
      if(curType&&typeOf(s[0])!==curType) return;
      if(f&&(s[0]+s[1]).toLowerCase().indexOf(f)<0) return;
      n++;
      h+='<tr><td><span class="tag code">'+esc(s[0])+'</span></td><td class="ttl">'+esc(s[1])+unitTag(s[0])+'</td><td>'+(s[2]?'<span class="tag">'+esc(s[2])+'</span>':'')+'</td><td>'+esc(s[5])+'</td>'+
        '<td style="white-space:nowrap;text-align:right"><a class="lnk primary" target="_blank" href="'+pdfUrl(s[0])+'">'+icon('i-file',14)+'PDF</a> '+
        (s[3]?'<a class="lnk" target="_blank" href="'+esc(s[3])+'">'+icon('i-play',14)+'影片</a>':'')+'</td></tr>';
    });
    tb.innerHTML=h||'<tr><td colspan="5" class="empty">沒有符合的標準</td></tr>';
    $('#cnt').textContent='共 '+n+' 項';
  }
  flt.oninput=draw; draw();
  window.scOvDraw=draw;
  if(hasSus){
    tChips.parentNode.insertBefore(sc,tChips);
    var drawSrc=function(){
      sc.innerHTML=[['','全部來源'],['ND','護理部共通'],['SUS','單位次專科']].map(function(o){return '<button class="chip'+((SRC===o[0]||(o[0]==='SUS'&&SRC.indexOf('SUS:')===0))?' on':'')+'" data-s="'+o[0]+'">'+o[1]+'</button>';}).join('');
      $$('.chip',sc).forEach(function(b){b.onclick=function(){SRC=b.dataset.s;try{localStorage.setItem('sc_src',SRC);}catch(e){}drawSrc();drawScope();draw();};});
    };
    drawSrc();
    var od=window.scOvDraw; window.scOvDraw=function(){drawSrc();od();};
  }
})();
setStep(0);
drawScope();
(function(){
  var m=location.search.match(/[?&]kw=([^&]*)/);
  if(m){window.scKwSearch(decodeURIComponent(m[1].replace(/\+/g,' ')));return;}
  var ml=location.search.match(/[?&]log=(\d+)/);
  if(ml){showView('chat',true);openLog(ml[1]);return;}
  var mq=location.search.match(/[?&]q=([^&]*)/);
  if(mq){showView('chat',true);qEl.value=decodeURIComponent(mq[1].replace(/\+/g,' '));if(history.replaceState)history.replaceState(null,'','index.html#chat');ask();return;}
  showView(location.hash.slice(1)||'chat',true);
})();
