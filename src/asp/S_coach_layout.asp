<%
'=====================================================================
' 護理標準查詢助理  共用版面（S_coach_layout.asp）
'   SC_Top "chat", "頁面標題"   → 輸出 <head>、Header、Sidebar，並開啟主內容區
'   SC_Bottom                    → 關閉主內容區、輸出共用 JS
'   key：chat / kw / all（前台三個畫面）、build / log / review / syn / aitest / setup（管理）
'   需先 include S_coach_conn.asp
'=====================================================================
Const SC_APP_NAME = "護理標準查詢助理"

Function SC_NavCls(key, cur)
  If key = cur Then SC_NavCls = "nav-item active" Else SC_NavCls = "nav-item"
End Function
Function SC_Cur(key, cur)
  If key = cur Then SC_Cur = " aria-current=""page""" Else SC_Cur = ""
End Function

Sub SC_Top(cur, pageTitle)
  Dim kc, kNI, kNA, kPct, isAdm, kCode
  kNI = 0: kNA = 0
  On Error Resume Next
  Set kc = SC_Open()
  If Err.Number = 0 Then
    kNA = kc.Execute("SELECT COUNT(*) FROM std_list")(0)
    kNI = kc.Execute("SELECT COUNT(*) FROM std_list WHERE status='已建索引'")(0)
    kc.Close
  End If
  Err.Clear
  On Error GoTo 0
  If kNA > 0 Then kPct = Int(kNI / kNA * 100) Else kPct = 0
  isAdm = SC_IsAdmin()
%><!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><%=H(pageTitle)%>｜<%=SC_APP_NAME%></title>
<link rel="stylesheet" href="css/S_coach.css">
</head>
<body>
<svg width="0" height="0" style="position:absolute" aria-hidden="true"><defs>
  <linearGradient id="ringGrad" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#1677FF"/><stop offset="1" stop-color="#13C2C2"/></linearGradient>
  <symbol id="i-cross" viewBox="0 0 24 24"><path fill="currentColor" d="M9.5 3h5a1 1 0 0 1 1 1v4.5H20a1 1 0 0 1 1 1v5a1 1 0 0 1-1 1h-4.5V20a1 1 0 0 1-1 1h-5a1 1 0 0 1-1-1v-4.5H4a1 1 0 0 1-1-1v-5a1 1 0 0 1 1-1h4.5V4a1 1 0 0 1 1-1Z"/></symbol>
  <symbol id="i-menu" viewBox="0 0 24 24"><path stroke="currentColor" stroke-width="2" stroke-linecap="round" d="M4 7h16M4 12h16M4 17h16"/></symbol>
  <symbol id="i-search" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></g></symbol>
  <symbol id="i-help" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M9.1 9a3 3 0 0 1 5.8 1c0 2-3 2.5-3 4.5"/><path d="M12 17.5h.01"/></g></symbol>
  <symbol id="i-home" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" d="m3 10 9-7 9 7v10a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1Z"/></symbol>
  <symbol id="i-chat" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a8 8 0 0 1-11.6 7.1L4 20l1-4.4A8 8 0 1 1 21 12Z"/><path d="M8.5 12h.01M12 12h.01M15.5 12h.01"/></g></symbol>
  <symbol id="i-key" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6h10M4 12h7M4 18h5"/><circle cx="17" cy="15" r="3.5"/><path d="m19.5 17.5 2 2"/></g></symbol>
  <symbol id="i-book" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5.5A2.5 2.5 0 0 1 6.5 3H20v15H6.5A2.5 2.5 0 0 0 4 20.5Z"/><path d="M4 20.5A2.5 2.5 0 0 0 6.5 23H20v-5"/><path d="M8 7h8M8 11h6"/></g></symbol>
  <symbol id="i-db" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><ellipse cx="12" cy="5.5" rx="8" ry="2.5"/><path d="M4 5.5v13c0 1.4 3.6 2.5 8 2.5s8-1.1 8-2.5v-13M4 12c0 1.4 3.6 2.5 8 2.5s8-1.1 8-2.5"/></g></symbol>
  <symbol id="i-list" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M9 6h11M9 12h11M9 18h11"/><path d="M4.5 6h.01M4.5 12h.01M4.5 18h.01"/></g></symbol>
  <symbol id="i-pulse" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" d="M3 12h4l2-6 4 12 2-6h6"/></symbol>
  <symbol id="i-file" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8Z"/><path d="M14 3v5h5M9 13h6M9 17h4"/></g></symbol>
  <symbol id="i-play" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="3"/><path d="m10 9 5 3-5 3Z"/></g></symbol>
  <symbol id="i-send" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 2 11 13"/><path d="M22 2 15 22l-4-9-9-4Z"/></g></symbol>
  <symbol id="i-spark" viewBox="0 0 24 24"><path fill="currentColor" d="M12 2l1.9 5.6L19.5 9.5l-5.6 1.9L12 17l-1.9-5.6L4.5 9.5l5.6-1.9Zm7 11 .9 2.6 2.6.9-2.6.9L19 21l-.9-2.6-2.6-.9 2.6-.9Z"/></symbol>
  <symbol id="i-shield" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3 4 6v6c0 5 3.4 8.3 8 9 4.6-.7 8-4 8-9V6Z"/><path d="m9 12 2 2 4-4"/></g></symbol>
  <symbol id="i-check" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" d="m5 12.5 4.5 4.5L19 7.5"/></symbol>
  <symbol id="i-info" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5h.01"/></g></symbol>
  <symbol id="i-plus" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" d="M12 5v14M5 12h14"/></symbol>
  <symbol id="i-right" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" d="m9 18 6-6-6-6"/></symbol>
  <symbol id="i-user" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/></g></symbol>
  <symbol id="i-lungs" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v8m0 0-3 2m3-2 3 2"/><path d="M8.5 7C6 8 3 12 3 17c0 2 1.5 3 3 3 2 0 3-1.5 3-3V9.5M15.5 7C18 8 21 12 21 17c0 2-1.5 3-3 3-2 0-3-1.5-3-3V9.5"/></g></symbol>
  <symbol id="i-tube" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M7 3v6a5 5 0 0 0 10 0V3"/><path d="M12 14v4a3 3 0 0 0 6 0v-1"/></g></symbol>
  <symbol id="i-syringe" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="m18 2 4 4M17 7l-9.5 9.5-3 .5.5-3L14.5 4.5M14 3l7 7M9 10l2 2M7 12l2 2M2 22l3.5-3.5"/></g></symbol>
  <symbol id="i-brain" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M9 4a3 3 0 0 0-3 3 3 3 0 0 0-2 5 3 3 0 0 0 2 5 3 3 0 0 0 6 1V5a2 2 0 0 0-3-1ZM15 4a3 3 0 0 1 3 3 3 3 0 0 1 2 5 3 3 0 0 1-2 5 3 3 0 0 1-6 1"/></g></symbol>
  <symbol id="i-clock" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></g></symbol>
  <symbol id="i-refresh" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 0 1-15.5 6.2L3 16"/><path d="M3 12a9 9 0 0 1 15.5-6.2L21 8"/><path d="M21 3v5h-5M3 21v-5h5"/></g></symbol>
  <symbol id="i-bolt" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" d="M13 2 4 14h7l-1 8 9-12h-7Z"/></symbol>
  <symbol id="i-alert" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/><path d="M12 9v4M12 17h.01"/></g></symbol>
  <symbol id="i-folder" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/></symbol>
  <symbol id="i-cpu" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><rect x="6" y="6" width="12" height="12" rx="2"/><path d="M10 10h4v4h-4zM9 2v4M15 2v4M9 18v4M15 18v4M2 9h4M2 15h4M18 9h4M18 15h4"/></g></symbol>
  <symbol id="i-gear" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1Z"/></g></symbol>
  <symbol id="i-users" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0M16 4.5a3.5 3.5 0 0 1 0 7M18 14a6.5 6.5 0 0 1 3.5 6"/></g></symbol>
  <symbol id="i-up" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M7 10v12"/><path d="M15 5.88 14 10h5.83a2 2 0 0 1 1.92 2.56l-2.33 8A2 2 0 0 1 17.5 22H4a2 2 0 0 1-2-2v-8a2 2 0 0 1 2-2h2.76a2 2 0 0 0 1.79-1.11L12 2a3.13 3.13 0 0 1 3 3.88Z"/></g></symbol>
  <symbol id="i-down" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M17 14V2"/><path d="M9 18.12 10 14H4.17a2 2 0 0 1-1.92-2.56l2.33-8A2 2 0 0 1 6.5 2H20a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2h-2.76a2 2 0 0 0-1.79 1.11L12 22a3.13 3.13 0 0 1-3-3.88Z"/></g></symbol>
  <symbol id="i-tag" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M20.6 13.4 13.4 20.6a2 2 0 0 1-2.8 0L3 13V3h10l7.6 7.6a2 2 0 0 1 0 2.8Z"/><path d="M7.5 7.5h.01"/></g></symbol>
  <symbol id="i-x" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" d="M6 6l12 12M18 6 6 18"/></symbol>
</defs></svg>
<header class="header">
  <button class="icon-btn menu-btn" id="menuBtn" aria-label="開啟選單"><svg width="20" height="20"><use href="#i-menu"/></svg></button>
  <a class="brand" href="S_coach.asp" style="color:inherit">
    <div class="brand-mark"><svg width="18" height="18"><use href="#i-cross"/></svg></div>
    <div class="brand-text"><b><%=SC_APP_NAME%></b><span>Nursing Standards Assistant</span></div>
  </a>
  <label class="search" title="在 PDF 全文中搜尋">
    <svg width="18" height="18"><use href="#i-search"/></svg>
    <input id="gsearch" placeholder="快速搜尋標準全文，例如：PICC 換藥" aria-label="搜尋標準全文">
    <span class="kbd">Ctrl K</span>
  </label>
  <div class="header-right">
    <% If SC_AI_ENABLED Then %><span class="env-pill on"><i></i>院內 AI 已啟用</span><% Else %><span class="env-pill off"><i></i>原文模式</span><% End If %>
    <a class="icon-btn" href="<%=SC_HOME_URL%>" aria-label="回內網首頁" title="回內網首頁"><svg width="20" height="20"><use href="#i-home"/></svg></a>
    <div class="user">
      <div class="avatar"><svg width="18" height="18"><use href="#i-user"/></svg></div>
      <div class="user-meta"><b><%=H(SC_Worknb())%></b><span><% If isAdm Then Response.Write "管理者" Else Response.Write H(Request.Cookies("division") & "") %></span></div>
    </div>
  </div>
</header>

<aside class="sidebar" id="sidebar" aria-label="主選單">
  <div class="nav-label">查詢</div>
  <a class="<%=SC_NavCls("chat", cur)%>" href="S_coach.asp#chat" data-view="chat"<%=SC_Cur("chat", cur)%>><svg width="20" height="20"><use href="#i-chat"/></svg>問助理</a>
  <a class="<%=SC_NavCls("kw", cur)%>" href="S_coach.asp#kw" data-view="kw"<%=SC_Cur("kw", cur)%>><svg width="20" height="20"><use href="#i-key"/></svg>關鍵字查詢</a>
  <a class="<%=SC_NavCls("all", cur)%>" href="S_coach.asp#all" data-view="all"<%=SC_Cur("all", cur)%>><svg width="20" height="20"><use href="#i-book"/></svg>標準總覽<span class="count"><%=kNA%></span></a>
  <a class="<%=SC_NavCls("my", cur)%>" href="S_coach_my.asp"<%=SC_Cur("my", cur)%>><svg width="20" height="20"><use href="#i-clock"/></svg>我的提問紀錄</a>
  <% If isAdm Then %>
  <div class="nav-label">管理</div>
  <a class="<%=SC_NavCls("build", cur)%>" href="S_coach_build.asp"<%=SC_Cur("build", cur)%>><svg width="20" height="20"><use href="#i-db"/></svg>索引管理<% If kNA > kNI Then %><span class="count"><%=kNA - kNI%></span><% End If %></a>
  <a class="<%=SC_NavCls("log", cur)%>" href="S_coach_log.asp"<%=SC_Cur("log", cur)%>><svg width="20" height="20"><use href="#i-list"/></svg>提問紀錄</a>
  <a class="<%=SC_NavCls("review", cur)%>" href="S_coach_review.asp"<%=SC_Cur("review", cur)%>><svg width="20" height="20"><use href="#i-up"/></svg>回饋與審核</a>
  <a class="<%=SC_NavCls("syn", cur)%>" href="S_coach_syn.asp"<%=SC_Cur("syn", cur)%>><svg width="20" height="20"><use href="#i-tag"/></svg>同義詞管理</a>
  <a class="<%=SC_NavCls("aitest", cur)%>" href="S_coach_aitest.asp"<%=SC_Cur("aitest", cur)%>><svg width="20" height="20"><use href="#i-pulse"/></svg>AI 連線測試</a>
  <a class="<%=SC_NavCls("setup", cur)%>" href="S_coach_setup.asp"<%=SC_Cur("setup", cur)%>><svg width="20" height="20"><use href="#i-gear"/></svg>系統初始化</a>
  <% End If %>
  <div class="sidebar-foot">
    <div class="kb">
      <div class="t">知識庫</div>
      <div class="v"><%=kNI%><small>/ <%=kNA%> 份標準</small></div>
      <div class="bar"><i style="width:<%=kPct%>%"></i></div>
      <ul>
        <li><svg width="14" height="14"><use href="#i-check"/></svg>只引用本部護理標準</li>
        <li><svg width="14" height="14"><use href="#i-check"/></svg>每段附標準編號與頁碼</li>
        <li><svg width="14" height="14"><use href="#i-check"/></svg>資料不離開醫院</li>
      </ul>
    </div>
  </div>
</aside>
<div class="scrim" id="scrim"></div>
<main class="main"><div class="container">
<%
End Sub

Sub SC_Bottom()
%>
</div></main>
<div class="toast" id="toast" role="status" aria-live="polite"></div>
<script>
(function(){
  var b=document.body, m=document.getElementById('menuBtn'), s=document.getElementById('scrim'), g=document.getElementById('gsearch');
  function nav(o){b.classList.toggle('nav-open',o);}
  m.onclick=function(){nav(!b.classList.contains('nav-open'));};
  s.onclick=function(){nav(false);};
  window.scNav=nav;
  document.addEventListener('keydown',function(e){
    if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){e.preventDefault();if(g.offsetParent)g.focus();else location.href='S_coach.asp#kw';}
    if(e.key==='Escape')nav(false);
  });
  g.addEventListener('keydown',function(e){
    var q=g.value.trim(); if(e.key!=='Enter'||e.isComposing||!q)return;
    if(window.scKwSearch) window.scKwSearch(q); else location.href='S_coach.asp?kw='+encodeURIComponent(q)+'#kw';
  });
})();
</script>
</body>
</html>
<%
End Sub
%>