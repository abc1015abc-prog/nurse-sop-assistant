<%@ Language="VBScript" CodePage="65001" EnableSessionState="False" %>
<!--#include file="S_coach_conn.asp"-->
<!--#include file="S_coach_layout.asp"-->
<!--#include file="S_coach_types.asp"-->
<%
'=====================================================================
' 同義詞管理（S_coach_syn.asp）－ 讓臨床縮寫、俗稱能對到標準上的正式用詞
'   t=list   同義詞清單（新增／修改／刪除、拆詞測試）
'   t=dict   匯入臨床縮寫詞庫（S_coach_syn_dict.txt）
'   t=mine   從標準 PDF 內文自動找「中文(英文/縮寫)」對照
'   t=mesh   MeSH 醫學主題詞查英文同義詞（伺服器需能連外）
'   t=miss   找不到的提問 → 補同義詞
'=====================================================================
Response.Charset = "utf-8"
Server.ScriptTimeout = 900
SC_CheckAdmin

Dim conn, act, t, msg, msgCls
Set conn = SC_Open()
SC_EnsureSynCols conn
act = Request("act")
t = Request("t")
If t <> "dict" And t <> "mine" And t <> "mesh" And t <> "miss" Then t = "list"
msg = "": msgCls = "ok"

'---------------------------------------------------------------------
' MeSH 查詢（轉送 NLM 的 JSON 給瀏覽器）
'---------------------------------------------------------------------
Function HttpGet(url)
  Dim x
  HttpGet = ""
  On Error Resume Next
  Set x = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
  x.setTimeouts 5000, 8000, 10000, 20000
  If SC_PROXY <> "" Then x.setProxy 2, SC_PROXY, "" Else x.setProxy 1
  x.open "GET", url, False
  x.setRequestHeader "Accept", "application/json"
  x.send
  If Err.Number <> 0 Then
    HttpGet = "{""ok"":0,""err"":" & JE("無法連線 MeSH（" & Err.Description & "）。伺服器可能無法連外，可改用詞庫或標準自動擷取。") & "}"
  ElseIf x.status <> 200 Then
    HttpGet = "{""ok"":0,""err"":" & JE("MeSH 回應錯誤 HTTP " & x.status) & "}"
  Else
    HttpGet = x.responseText
  End If
  Err.Clear
  On Error GoTo 0
End Function

If act = "mesh" Or act = "meshd" Then
  Dim mq, re0
  Response.ContentType = "application/json"
  mq = Trim(Request("q") & "")
  If act = "mesh" Then
    If mq = "" Or Len(mq) > 80 Then Response.Write "{""ok"":0,""err"":""請輸入英文詞""}": Response.End
    Response.Write HttpGet("https://id.nlm.nih.gov/mesh/lookup/descriptor?label=" & Server.URLEncode(mq) & "&match=contains&limit=12")
  Else
    Set re0 = New RegExp: re0.Pattern = "^D\d{6,9}$"
    If Not re0.Test(mq) Then Response.Write "{""ok"":0,""err"":""代碼錯誤""}": Response.End
    Response.Write HttpGet("https://id.nlm.nih.gov/mesh/lookup/details?descriptor=" & mq & "&includes=terms")
  End If
  conn.Close
  Response.End
End If

'---------------------------------------------------------------------
' 同義詞合併：任一個詞已存在某組 → 併入那組；否則新增一組
'---------------------------------------------------------------------
Dim gTerm, gWord, gAlias, nAdd, nMerge, nSkip
Set gTerm = Server.CreateObject("Scripting.Dictionary")
Set gWord = Server.CreateObject("Scripting.Dictionary")
Set gAlias = Server.CreateObject("Scripting.Dictionary")
nAdd = 0: nMerge = 0: nSkip = 0

Sub LoadSyn()
  Dim r, sid, w, a
  gTerm.RemoveAll: gWord.RemoveAll: gAlias.RemoveAll
  Set r = conn.Execute("SELECT id, word, alias FROM syn ORDER BY id")
  Do Until r.EOF
    sid = CLng(r("id").Value)
    gWord(sid) = r("word") & ""
    gAlias(sid) = r("alias") & ""
    w = SC_Norm(r("word"))
    If w <> "" And Not gTerm.Exists(w) Then gTerm.Add w, sid
    For Each a In SplitAlias(r("alias") & "")
      a = SC_Norm(a)
      If a <> "" And Not gTerm.Exists(a) Then gTerm.Add a, sid
    Next
    r.MoveNext
  Loop
  r.Close
End Sub

' 別名可用 , ， 、 ; 分隔
Function SplitAlias(s)
  Dim x
  x = Replace(Replace(Replace(Replace(s & "", "，", ","), "、", ","), ";", ","), "；", ",")
  SplitAlias = Split(x, ",")
End Function

Function CleanList(s)
  Dim a, r, seen
  Set seen = Server.CreateObject("Scripting.Dictionary")
  r = ""
  For Each a In SplitAlias(s)
    a = Trim(SC_Half(a))
    If a <> "" And Not seen.Exists(SC_Norm(a)) Then
      seen.Add SC_Norm(a), 1
      If r <> "" Then r = r & ","
      r = r & a
    End If
  Next
  CleanList = r
End Function

Function MergeGroup(word, aliasStr, srcName)
  Dim terms, a, sid, newA, nw, rsI, added
  word = Trim(SC_Half(word))
  aliasStr = CleanList(aliasStr)
  MergeGroup = 0
  If word = "" Then Exit Function
  ' 找現有群組
  sid = 0
  If gTerm.Exists(SC_Norm(word)) Then sid = gTerm(SC_Norm(word))
  If sid = 0 Then
    For Each a In SplitAlias(aliasStr)
      If SC_Norm(a) <> "" And gTerm.Exists(SC_Norm(a)) Then sid = gTerm(SC_Norm(a)): Exit For
    Next
  End If
  If sid = 0 Then
    conn.Execute "INSERT INTO syn (word, alias, src) VALUES ('" & SQ(word) & "','" & SQ(aliasStr) & "','" & SQ(srcName) & "')"
    Set rsI = conn.Execute("SELECT @@IDENTITY")
    sid = CLng(rsI(0).Value): rsI.Close
    gWord(sid) = word: gAlias(sid) = aliasStr
    gTerm(SC_Norm(word)) = sid
    For Each a In SplitAlias(aliasStr)
      If SC_Norm(a) <> "" And Not gTerm.Exists(SC_Norm(a)) Then gTerm.Add SC_Norm(a), sid
    Next
    nAdd = nAdd + 1
  Else
    newA = gAlias(sid): added = False
    nw = Array(word)
    For Each a In nw
      If Not gTerm.Exists(SC_Norm(a)) Then
        If newA <> "" Then newA = newA & ","
        newA = newA & a: gTerm.Add SC_Norm(a), sid: added = True
      End If
    Next
    For Each a In SplitAlias(aliasStr)
      a = Trim(a)
      If SC_Norm(a) <> "" And Not gTerm.Exists(SC_Norm(a)) Then
        If newA <> "" Then newA = newA & ","
        newA = newA & a: gTerm.Add SC_Norm(a), sid: added = True
      End If
    Next
    If added Then
      conn.Execute "UPDATE syn SET alias='" & SQ(newA) & "' WHERE id=" & sid
      gAlias(sid) = newA
      nMerge = nMerge + 1
    Else
      nSkip = nSkip + 1
    End If
  End If
  MergeGroup = sid
End Function

'---------------------------------------------------------------------
' 動作
'---------------------------------------------------------------------
Dim fid, fw, fa, lines, ln, pp, sel, dictPath, fso0
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
  If act = "save" Then
    fid = CLng("0" & Request.Form("id"))
    fw = Trim(SC_Half(Request.Form("word")))
    fa = CleanList(Request.Form("alias"))
    If fw = "" Then
      msg = "請輸入標準用詞": msgCls = "err"
    ElseIf fid > 0 Then
      conn.Execute "UPDATE syn SET word='" & SQ(fw) & "', alias='" & SQ(fa) & "' WHERE id=" & fid
      msg = "已更新「" & fw & "」"
    Else
      LoadSyn
      MergeGroup fw, fa, "手動"
      If nAdd > 0 Then
        msg = "已新增「" & fw & "」"
      ElseIf nMerge > 0 Then
        msg = "「" & fw & "」已有相同的詞，已把新別名併入原本那一組"
      Else
        msg = "這些詞都已經在同義詞中了"
      End If
    End If
  ElseIf act = "del" Then
    fid = CLng("0" & Request.Form("id"))
    If fid > 0 Then conn.Execute "DELETE FROM syn WHERE id=" & fid: msg = "已刪除一組同義詞"
  ElseIf act = "import" Then
    t = "dict"
    dictPath = Server.MapPath("S_coach_syn_dict.txt")
    Set fso0 = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso0.FileExists(dictPath) Then
      msg = "找不到詞庫檔 S_coach_syn_dict.txt，請放在與本頁相同的資料夾": msgCls = "err"
    Else
      LoadSyn
      lines = Split(Replace(SC_ReadUtf8(dictPath), vbCr, ""), vbLf)
      For Each ln In lines
        ln = Trim(ln)
        If ln <> "" And Left(ln, 1) <> "#" And InStr(ln, "|") > 1 Then
          pp = Split(ln, "|")
          MergeGroup pp(0), pp(1), "詞庫"
        End If
      Next
      msg = "匯入完成：新增 " & nAdd & " 組、併入既有 " & nMerge & " 組、已存在略過 " & nSkip & " 組"
    End If
  ElseIf act = "addsel" Or act = "addgroup" Then
    LoadSyn
    If act = "addsel" Then
      t = "mine"
      For Each sel In Request.Form("c")
        pp = Split(sel & "", "|")
        If UBound(pp) = 1 Then MergeGroup pp(0), pp(1), "標準擷取"
      Next
    Else
      t = "mesh"
      MergeGroup Request.Form("word"), Request.Form("alias"), "MeSH"
    End If
    msg = "已加入：新增 " & nAdd & " 組、併入既有 " & nMerge & " 組" & IIf3(nSkip > 0, "、" & nSkip & " 組原本就有", "")
  End If
End If

Function IIf3(c, a, b)
  If c Then IIf3 = a Else IIf3 = b
End Function

Function TabCls(k)
  If k = t Then TabCls = "chip on" Else TabCls = "chip"
End Function

Function SrcBadge(s)
  Select Case s & ""
    Case "詞庫"
      SrcBadge = "<span class='badge info'>詞庫</span>"
    Case "標準擷取"
      SrcBadge = "<span class='badge success'>標準擷取</span>"
    Case "MeSH"
      SrcBadge = "<span class='badge warning'>MeSH</span>"
    Case "手動"
      SrcBadge = "<span class='badge neutral'>手動</span>"
    Case Else
      SrcBadge = "<span class='badge neutral'>預設</span>"
  End Select
End Function

Dim nSyn
nSyn = conn.Execute("SELECT COUNT(*) FROM syn")(0)

SC_Top "syn", "同義詞管理"
%>
<style>
.syn-chips{display:flex;flex-wrap:wrap;gap:4px}
.syn-chips .tag{font-weight:500}
.syn-edit{display:none;margin-top:8px}
tr.editing .syn-edit{display:flex}
tr.editing .syn-view{display:none}
.syn-test{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px;min-height:24px}
.syn-test .hit{background:var(--primary-50);color:var(--primary-700)}
.syn-test .rest{background:#FFF7E6;color:#AD6800}
.cand td{vertical-align:top}
.mesh-res button{display:block;width:100%;text-align:left;border:1px solid var(--line);background:#fff;border-radius:10px;padding:10px 12px;margin:0 0 6px;cursor:pointer;font:inherit}
.mesh-res button:hover,.mesh-res button.on{border-color:var(--primary);background:var(--primary-50)}
.mesh-terms label{display:flex;gap:8px;align-items:flex-start;padding:4px 0;font-size:13.5px}
.form-grid{display:grid;grid-template-columns:200px 1fr auto;gap:8px;align-items:center}
@media (max-width:760px){.form-grid{grid-template-columns:1fr}
  .lst.stack thead{display:none}
  .lst.stack tr{display:grid;grid-template-columns:1fr auto;gap:6px 10px;padding:12px 16px;border-bottom:1px solid var(--line)}
  .lst.stack td{display:block;padding:0!important;border:0!important;width:auto!important}
  .lst.stack td:nth-child(1){grid-row:1;grid-column:1}.lst.stack td:nth-child(4){grid-row:1;grid-column:2}.lst.stack td:nth-child(2){grid-row:2;grid-column:1/-1}
  .lst.stack td.hide-m{display:none}}
</style>
<nav class="crumbs">管理<svg width="14" height="14"><use href="#i-right"/></svg><span>同義詞管理</span></nav>
<div class="page-head">
  <div>
    <div class="page-title"><h1>同義詞管理</h1><span class="badge neutral"><%=nSyn%> 組</span></div>
    <p class="page-sub">護理師常用縮寫或俗稱（如 Foley、on endo、NG、i/o），要先對到標準上的正式用詞，才找得到正確段落。同一組的詞在查詢時會一起找。</p>
  </div>
  <div class="head-actions">
    <a class="btn btn-secondary" href="S_coach_log.asp?f=miss"><svg width="18" height="18"><use href="#i-list"/></svg>提問紀錄</a>
  </div>
</div>

<% If msg <> "" Then %>
<div class="banner <%=IIf3(msgCls = "err", "err", "ok")%>" style="margin-bottom:16px"><svg width="18" height="18"><use href="#<%=IIf3(msgCls = "err", "i-alert", "i-check")%>"/></svg><div><%=H(msg)%></div></div>
<% End If %>

<div class="chips" style="margin-bottom:16px">
  <a class="<%=TabCls("list")%>" href="S_coach_syn.asp?t=list">同義詞清單</a>
  <a class="<%=TabCls("dict")%>" href="S_coach_syn.asp?t=dict">臨床縮寫詞庫</a>
  <a class="<%=TabCls("mine")%>" href="S_coach_syn.asp?t=mine">從標準自動擷取</a>
  <a class="<%=TabCls("mesh")%>" href="S_coach_syn.asp?t=mesh">MeSH 查詢</a>
  <a class="<%=TabCls("miss")%>" href="S_coach_syn.asp?t=miss">找不到的提問</a>
</div>

<%
'=====================================================================
If t = "list" Then
'=====================================================================
Dim rs, grpJ, srcN
grpJ = ""
Set rs = conn.Execute("SELECT id, word, alias, src FROM syn ORDER BY word")
%>
<section class="card" style="margin-bottom:16px">
  <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-search"/></svg></div><div><h2>拆詞測試</h2><p>輸入護理師可能會問的話，看系統認得哪些詞。藍色＝有對到同義詞；黃色＝沒有對到，可以補進來。</p></div></div>
  <div class="card-body">
    <label class="big-input"><svg width="18" height="18"><use href="#i-search"/></svg><input id="tq" placeholder="例如：on foley 多久要換？病人 NG 反抽多少要停灌？"></label>
    <div class="syn-test" id="tr"></div>
  </div>
</section>

<section class="card" style="margin-bottom:16px" id="addCard">
  <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-plus"/></svg></div><div><h2>新增一組同義詞</h2><p>左邊填標準上的正式用詞，右邊填縮寫、俗稱、英文（用逗號分隔）。如果其中有詞已存在，會自動併入原本那一組。</p></div></div>
  <div class="card-body">
    <form method="post" action="S_coach_syn.asp?t=list" class="form-grid">
      <input type="hidden" name="act" value="save">
      <input class="input" name="word" id="fWord" placeholder="標準用詞，例：導尿" required>
      <input class="input" name="alias" id="fAlias" placeholder="別名，例：foley, on foley, 尿管">
      <button class="btn btn-primary"><svg width="18" height="18"><use href="#i-plus"/></svg>加入</button>
    </form>
  </div>
</section>

<section class="card">
  <div class="toolbar">
    <label class="big-input"><svg width="18" height="18"><use href="#i-search"/></svg><input id="tf" placeholder="篩選同義詞"></label>
  </div>
  <div class="table-wrap">
  <table class="lst admin stack">
    <thead><tr><th style="width:200px">標準用詞</th><th>縮寫／俗稱／英文</th><th style="width:100px" class="hide-m">來源</th><th style="width:130px"></th></tr></thead>
    <tbody id="tb">
<%
If rs.EOF Then Response.Write "<tr><td colspan='4' class='empty'>還沒有同義詞，先到「臨床縮寫詞庫」匯入吧</td></tr>"
Do Until rs.EOF
  If grpJ <> "" Then grpJ = grpJ & ","
  grpJ = grpJ & "[" & JE(rs("word") & "")
  For Each fa In SplitAlias(rs("alias") & "")
    If Trim(fa) <> "" Then grpJ = grpJ & "," & JE(Trim(fa))
  Next
  grpJ = grpJ & "]"
%>
<tr data-s="<%=H(LCase(rs("word") & " " & rs("alias")))%>">
  <td><b><%=H(rs("word"))%></b></td>
  <td>
    <div class="syn-view syn-chips"><% For Each fa In SplitAlias(rs("alias") & "")
      If Trim(fa) <> "" Then Response.Write "<span class='tag'>" & H(Trim(fa)) & "</span>"
    Next %></div>
    <form class="syn-edit inline-form" method="post" action="S_coach_syn.asp?t=list" style="flex-wrap:wrap">
      <input type="hidden" name="act" value="save"><input type="hidden" name="id" value="<%=rs("id")%>">
      <input class="input input-sm" name="word" value="<%=H(rs("word"))%>" style="max-width:160px">
      <input class="input input-sm" name="alias" value="<%=H(rs("alias"))%>" style="flex:1">
      <button class="btn btn-primary btn-sm">儲存</button>
      <button type="button" class="btn btn-ghost btn-sm" onclick="this.closest('tr').classList.remove('editing')">取消</button>
    </form>
  </td>
  <td class="hide-m"><%=SrcBadge(rs("src"))%></td>
  <td style="text-align:right;white-space:nowrap">
    <button type="button" class="lnk" onclick="this.closest('tr').classList.add('editing')">修改</button>
    <form method="post" action="S_coach_syn.asp?t=list" style="display:inline" onsubmit="return confirm('確定刪除「<%=H(Replace(rs("word") & "", "'", ""))%>」這一組？')">
      <input type="hidden" name="act" value="del"><input type="hidden" name="id" value="<%=rs("id")%>">
      <button class="lnk">刪除</button>
    </form>
  </td>
</tr>
<%
  rs.MoveNext
Loop
rs.Close
%>
    </tbody>
  </table>
  </div>
  <div class="empty" id="noRow" style="display:none">沒有符合的同義詞</div>
</section>
<script>
(function(){
  var G=[<%=grpJ%>];
  function norm(s){return (s||'').toLowerCase().replace(/[！-～]/g,function(c){return String.fromCharCode(c.charCodeAt(0)-0xfee0);}).replace(/[\s　]+/g,'');}
  function isLat(s){return /^[a-z0-9\-\.\s\/]+$/.test(s);}
  function esc(s){return s.replace(/[.*+?^${}()|[\]\\\/]/g,'\\$&');}
  function has(q,v){
    if(isLat(v)){
      if(new RegExp('(^|[^a-z0-9])'+esc(v)+'([^a-z0-9]|$)').test(q))return true;
      return v.length>=6 && q.replace(/ /g,'').indexOf(v)>=0;
    }
    return q.indexOf(v)>=0;
  }
  function h(s){return s.replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
  var tq=document.getElementById('tq'), tr=document.getElementById('tr');
  tq.oninput=function(){
    var q=tq.value.toLowerCase().replace(/[！-～]/g,function(c){return String.fromCharCode(c.charCodeAt(0)-0xfee0);}), out=[];
    G.forEach(function(g){
      var vs=g.map(norm), hit=vs.some(function(v){return v.length>=2&&has(q,v);});
      if(hit){out.push('<span class="tag hit" title="'+h(g.join('、'))+'">'+h(g[0])+' ＝ '+h(g.slice(1,6).join('、'))+(g.length>6?'…':'')+'</span>');
        vs.slice().sort(function(a,b){return b.length-a.length;}).forEach(function(v){if(v.length>=2&&has(q,v))q=isLat(v)?q.replace(new RegExp('(^|[^a-z0-9])'+(v.length>=6?v.split('').map(esc).join('\\s*'):esc(v))+'(?=[^a-z0-9]|$)','g'),'$1 '):q.split(v).join(' ');});}
    });
    ['請問','怎麼','如何','什麼','多久','多少','要','嗎','呢','的','是否','需要','可以','病人','注意事項','步驟','?','？','，',','].forEach(function(w){q=q.split(w).join(' ');});
    (q.match(/[㐀-鿿]{2,}|[a-z][a-z0-9\/\-]+/g)||[]).forEach(function(w){out.push('<span class="tag rest" title="沒有對到同義詞，會直接用字面去找">'+h(w)+'</span>');});
    tr.innerHTML=out.join('')||'<span class="note">輸入後即時顯示</span>';
  };
  tq.oninput();
  var tf=document.getElementById('tf'), rows=[].slice.call(document.querySelectorAll('#tb tr[data-s]'));
  tf.oninput=function(){
    var q=tf.value.trim().toLowerCase(), n=0;
    rows.forEach(function(r){var ok=!q||r.dataset.s.indexOf(q)>=0;r.style.display=ok?'':'none';if(ok)n++;});
    document.getElementById('noRow').style.display=(rows.length&&!n)?'block':'none';
  };
  var m=location.search.match(/[?&]add=([^&]*)/);
  if(m){document.getElementById('fAlias').value=decodeURIComponent(m[1].replace(/\+/g,' '));document.getElementById('fWord').focus();document.getElementById('addCard').scrollIntoView();}
})();
</script>
<%
'=====================================================================
ElseIf t = "dict" Then
'=====================================================================
Dim nDict, dl
nDict = 0
dictPath = Server.MapPath("S_coach_syn_dict.txt")
Set fso0 = Server.CreateObject("Scripting.FileSystemObject")
If fso0.FileExists(dictPath) Then
  For Each dl In Split(Replace(SC_ReadUtf8(dictPath), vbCr, ""), vbLf)
    dl = Trim(dl)
    If dl <> "" And Left(dl, 1) <> "#" And InStr(dl, "|") > 1 Then nDict = nDict + 1
  Next
End If
%>
<section class="card">
  <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-book"/></svg></div><div><h2>臨床縮寫詞庫</h2><p>內建約 <%=nDict%> 組臨床常用縮寫與俗稱（呼吸道、管路、給藥、檢驗、急重症、婦兒、感控、基本照護…），對應到標準上的中文用詞。</p></div></div>
  <div class="card-body">
    <ul style="margin:0 0 16px;padding-left:20px;line-height:1.9;color:var(--text-2)">
      <li>詞庫檔：<code>S_coach_syn_dict.txt</code>（與本頁同一資料夾，UTF-8），一行一組：<code>標準用詞|別名1,別名2</code></li>
      <li>匯入時，只要某個詞已經在某一組，就併入那一組，不會重複；可以放心重複匯入。</li>
      <li>容易誤判的兩字母縮寫（PE、PT、OR、AD…）刻意沒收，貴部常用的話請自行在清單中加入。</li>
      <li>單位慣用語不同時，直接修改詞庫檔再匯入，或在「同義詞清單」修改。</li>
    </ul>
    <% If nDict = 0 Then %>
      <div class="banner err"><svg width="18" height="18"><use href="#i-alert"/></svg><div>找不到詞庫檔或檔案是空的，請把 S_coach_syn_dict.txt 放到本頁同一個資料夾。</div></div>
    <% Else %>
      <form method="post" action="S_coach_syn.asp?t=dict" onsubmit="var b=this.querySelector('button');b.classList.add('loading');b.disabled=true;">
        <input type="hidden" name="act" value="import">
        <button class="btn btn-primary"><span class="spin"></span><svg width="18" height="18" class="lbl-ico"><use href="#i-plus"/></svg>匯入 <%=nDict%> 組詞庫</button>
      </form>
    <% End If %>
  </div>
</section>
<%
'=====================================================================
ElseIf t = "mine" Then
'=====================================================================
%>
<section class="card" style="margin-bottom:16px">
  <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-file"/></svg></div><div><h2>從標準內文自動擷取</h2><p>標準裡常寫「周邊置入中心靜脈導管（PICC）」這類對照。這裡把全部標準掃一遍，找出中英文／縮寫對照，勾選後加入同義詞。這些詞直接來自本部標準，最貼近實際用語。</p></div></div>
  <div class="card-body">
    <% If act <> "scan" Then %>
    <a class="btn btn-primary" href="S_coach_syn.asp?t=mine&act=scan" onclick="this.classList.add('loading')"><span class="spin"></span><svg width="18" height="18" class="lbl-ico"><use href="#i-search"/></svg>開始掃描標準</a>
    <span class="note" style="margin-left:8px">約需 10～60 秒</span>
    <% End If %>
  </div>
</section>
<%
If act = "scan" Then
  Dim rsC, reA, reB, mc, mm, zh, enRaw, enArr, enOk, en, k, cands, candCode, cn, cc, keyC, maxN, bucket, existW, allIn, firstSid, sid2, x, units, reU, reCode, reNum
  LoadSyn
  Set cands = Server.CreateObject("Scripting.Dictionary")
  Set candCode = Server.CreateObject("Scripting.Dictionary")
  Set reA = New RegExp: reA.Global = True
  reA.Pattern = "([㐀-鿿]{2,16})\s*[（(]\s*([A-Za-z][A-Za-z0-9 ,;\.\-/&']{1,80}?)\s*[)）]"
  Set reB = New RegExp: reB.Global = True
  reB.Pattern = "([A-Za-z][A-Za-z0-9\-/]{1,9})\s*[（(]\s*([㐀-鿿]{2,16})\s*[)）]"
  Set reU = New RegExp: reU.IgnoreCase = True
  reU.Pattern = "^(mg|ml|cc|cm|mm|kg|g|gm|mcg|meq|iu|u|l|hr|hrs|min|mmhg|cmh2o|ml/hr|ml/kg|mg/kg|note|ps|e\.g\.?|i\.e\.?|etc\.?|figure|fig|table|yes|no|or|and|for|if|the|of|to)$"
  Set reCode = New RegExp: reCode.IgnoreCase = True
  reCode.Pattern = "^[a-z]{1,4}\d{2,}$|\d{3,}"
  Set rsC = conn.Execute("SELECT code, content FROM chunk")
  Do Until rsC.EOF
    x = Replace(rsC("content") & "", vbLf, " ")
    ' 中文（英文, 縮寫）
    For Each mm In reA.Execute(x)
      zh = CleanZh(mm.SubMatches(0))
      enRaw = mm.SubMatches(1)
      enOk = ""
      For Each en In Split(Replace(enRaw, ";", ","), ",")
        en = Trim(en)
        If Len(en) >= 2 And Len(en) <= 60 And Not reU.Test(en) And Not reCode.Test(en) Then
          If enOk <> "" Then enOk = enOk & ","
          enOk = enOk & en
        End If
      Next
      If zh <> "" And enOk <> "" Then AddCand zh, enOk, rsC("code") & ""
    Next
    ' 縮寫（中文）
    For Each mm In reB.Execute(x)
      en = mm.SubMatches(0)
      zh = CleanZh(mm.SubMatches(1))
      If zh <> "" And Not reU.Test(en) And Not reCode.Test(en) Then AddCand zh, en, rsC("code") & ""
    Next
    rsC.MoveNext
  Loop
  rsC.Close

  ' 依出現次數排序（次數多的先）
  maxN = 0
  For Each keyC In cands.Keys
    If cands(keyC) > maxN Then maxN = cands(keyC)
  Next
  cn = 0
%>
<form method="post" action="S_coach_syn.asp?t=mine">
<input type="hidden" name="act" value="addsel">
<section class="card">
  <div class="toolbar">
    <b>找到 <%=cands.Count%> 組對照</b><span class="note">已勾選出現 2 次以上的；「已在同義詞中」的不會列出。請把不是同義詞的取消勾選。</span>
    <span style="flex:1"></span>
    <button type="button" class="btn btn-ghost btn-sm" onclick="document.querySelectorAll('.cand input').forEach(function(c){c.checked=true})">全選</button>
    <button type="button" class="btn btn-ghost btn-sm" onclick="document.querySelectorAll('.cand input').forEach(function(c){c.checked=false})">全不選</button>
    <button class="btn btn-primary btn-sm"><svg width="16" height="16"><use href="#i-plus"/></svg>加入勾選的</button>
  </div>
  <div class="table-wrap">
  <table class="lst admin cand">
    <thead><tr><th style="width:40px"></th><th style="width:220px">中文（標準用詞）</th><th>英文／縮寫</th><th style="width:80px;text-align:right">次數</th><th class="hide-m" style="width:200px">出現的標準</th><th class="hide-m" style="width:160px">處理方式</th></tr></thead>
    <tbody>
<%
  For bucket = maxN To 1 Step -1
    For Each keyC In cands.Keys
      If cands(keyC) = bucket And cn < 600 Then
        pp = Split(keyC, "|")
        ' 是否已全部存在於同一組
        firstSid = -1: allIn = True: existW = ""
        For Each en In Array(pp(0))
          If gTerm.Exists(SC_Norm(en)) Then firstSid = gTerm(SC_Norm(en)) Else allIn = False
        Next
        For Each en In Split(pp(1), ",")
          If gTerm.Exists(SC_Norm(en)) Then
            sid2 = gTerm(SC_Norm(en))
            If firstSid = -1 Then firstSid = sid2
            If sid2 <> firstSid Then allIn = False
          Else
            allIn = False
          End If
        Next
        If Not allIn Then
          cn = cn + 1
          If firstSid > 0 Then existW = gWord(firstSid)
%>
<tr>
  <td><input type="checkbox" name="c" value="<%=H(keyC)%>"<% If bucket >= 2 Then Response.Write " checked" %> aria-label="選取"></td>
  <td><b><%=H(pp(0))%></b></td>
  <td><div class="syn-chips"><% For Each en In Split(pp(1), ",")
    Response.Write "<span class='tag'>" & H(en) & "</span>"
  Next %></div></td>
  <td class="num"><%=bucket%></td>
  <td class="hide-m"><span class="note"><%=H(candCode(keyC))%></span></td>
  <td class="hide-m"><% If existW <> "" Then %><span class="badge info">併入「<%=H(existW)%>」</span><% Else %><span class="badge success">新增一組</span><% End If %></td>
</tr>
<%
        End If
      End If
    Next
  Next
  If cn = 0 Then Response.Write "<tr><td colspan='6' class='empty'>沒有新的對照（可能都已經加入了，或還沒建立索引）</td></tr>"
%>
    </tbody>
  </table>
  </div>
</section>
</form>
<%
End If

'=====================================================================
ElseIf t = "mesh" Then
'=====================================================================
%>
<div class="two-col" style="display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:16px">
<section class="card">
  <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-search"/></svg></div><div><h2>查詢 MeSH 主題詞</h2><p>輸入英文（例：urinary catheter、tracheostomy、pressure ulcer），從美國國家醫學圖書館的 MeSH 找出正式名稱與所有英文同義寫法。</p></div></div>
  <div class="card-body">
    <form class="inline-form" id="mf"><input class="input" id="mq" placeholder="英文醫學名詞" style="flex:1"><button class="btn btn-primary"><svg width="18" height="18"><use href="#i-search"/></svg>查詢</button></form>
    <div class="mesh-res" id="mr" style="margin-top:12px"></div>
    <p class="note" style="margin-top:12px">MeSH 只有英文，沒有中文與台灣臨床慣用縮寫（如 on endo、NG 反抽），適合補「英文全名的各種寫法」。伺服器需能連到 id.nlm.nih.gov，不能連外時請改用詞庫或標準自動擷取。</p>
  </div>
</section>
<section class="card">
  <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-plus"/></svg></div><div><h2>加入同義詞</h2><p>勾選要的英文寫法，填上標準上的中文用詞後加入。</p></div></div>
  <div class="card-body">
    <form method="post" action="S_coach_syn.asp?t=mesh" id="mAdd">
      <input type="hidden" name="act" value="addgroup"><input type="hidden" name="alias" id="mAlias">
      <div class="mesh-terms" id="mt"><span class="note">先在左邊選一個主題詞</span></div>
      <div class="inline-form" style="margin-top:12px">
        <input class="input" name="word" id="mWord" placeholder="標準用詞（中文），例：導尿" required style="flex:1">
        <button class="btn btn-primary"><svg width="18" height="18"><use href="#i-plus"/></svg>加入</button>
      </div>
    </form>
  </div>
</section>
</div>
<script>
(function(){
  var mr=document.getElementById('mr'), mt=document.getElementById('mt');
  function h(s){return String(s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
  // "Catheterization, Central Venous" → "Central Venous Catheterization"
  function flip(s){var p=s.split(', ');return p.length===2?p[1]+' '+p[0]:s;}
  function get(u){return fetch(u,{credentials:'same-origin'}).then(function(r){return r.json();});}
  document.getElementById('mf').onsubmit=function(e){
    e.preventDefault(); var q=document.getElementById('mq').value.trim(); if(!q)return;
    mr.innerHTML='<div class="loading-row"><span class="dots"><i></i><i></i><i></i></span>查詢中…</div>';
    get('S_coach_syn.asp?act=mesh&q='+encodeURIComponent(q)).then(function(d){
      if(!Array.isArray(d)){mr.innerHTML='<div class="banner err">'+h(d.err||'查詢失敗')+'</div>';return;}
      if(!d.length){mr.innerHTML='<div class="note">MeSH 中找不到，換個英文寫法試試</div>';return;}
      mr.innerHTML=d.map(function(x){var id=x.resource.split('/').pop();return '<button type="button" data-id="'+h(id)+'"><b>'+h(x.label)+'</b> <span class="note">'+h(id)+'</span></button>';}).join('');
    }).catch(function(e){mr.innerHTML='<div class="banner err">'+h(e.message)+'</div>';});
  };
  mr.onclick=function(e){
    var b=e.target.closest('button'); if(!b)return;
    [].forEach.call(mr.children,function(c){c.classList.toggle('on',c===b);});
    mt.innerHTML='<div class="loading-row"><span class="dots"><i></i><i></i><i></i></span>讀取同義寫法…</div>';
    get('S_coach_syn.asp?act=meshd&q='+encodeURIComponent(b.dataset.id)).then(function(d){
      if(!d.terms){mt.innerHTML='<div class="banner err">'+h(d.err||'讀取失敗')+'</div>';return;}
      var seen={}, list=[];
      d.terms.forEach(function(x){var v=flip(x.label); var k=v.toLowerCase(); if(!seen[k]&&v.length<=60){seen[k]=1;list.push(v);}});
      mt.innerHTML=list.map(function(v){return '<label><input type="checkbox" checked value="'+h(v)+'">'+h(v)+'</label>';}).join('');
    }).catch(function(e){mt.innerHTML='<div class="banner err">'+h(e.message)+'</div>';});
  };
  document.getElementById('mAdd').onsubmit=function(e){
    var v=[].map.call(mt.querySelectorAll('input:checked'),function(c){return c.value.replace(/,/g,' ');});
    if(!v.length){e.preventDefault();alert('請至少勾選一個英文寫法');return;}
    document.getElementById('mAlias').value=v.join(',');
  };
})();
</script>
<style>@media (max-width:900px){.two-col{grid-template-columns:1fr!important}}</style>
<%
'=====================================================================
ElseIf t = "miss" Then
'=====================================================================
Dim rsM, mDict, mLast, mq2, mk
Set mDict = Server.CreateObject("Scripting.Dictionary")
Set mLast = Server.CreateObject("Scripting.Dictionary")
Set rsM = conn.Execute("SELECT TOP 500 question, qtime FROM qlog WHERE qmode='ask' AND (result='找不到' OR result='low' OR result='無關鍵詞') AND qtime >= Date()-90 ORDER BY qtime DESC")
Do Until rsM.EOF
  mq2 = Trim(rsM("question") & "")
  If mq2 <> "" Then
    If mDict.Exists(mq2) Then
      mDict(mq2) = mDict(mq2) + 1
    Else
      mDict.Add mq2, 1: mLast.Add mq2, rsM("qtime").Value
    End If
  End If
  rsM.MoveNext
Loop
rsM.Close
%>
<section class="card">
  <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-alert"/></svg></div><div><h2>近 90 天找不到／可能相關的提問</h2><p>常常是用了縮寫或口語。按「補同義詞」把詞加進來，再按「測試」確認找得到。若標準本來就沒寫，則列入標準修訂參考。</p></div></div>
  <div class="table-wrap">
  <table class="lst admin">
    <thead><tr><th>提問</th><th style="width:80px;text-align:right">次數</th><th class="hide-m" style="width:120px">最近一次</th><th style="width:190px"></th></tr></thead>
    <tbody>
<%
If mDict.Count = 0 Then Response.Write "<tr><td colspan='4' class='empty'>近 90 天沒有找不到的提問 🎉</td></tr>"
For Each mk In mDict.Keys
%>
<tr>
  <td class="q-cell"><%=H(mk)%></td>
  <td class="num"><%=mDict(mk)%></td>
  <td class="hide-m"><span class="note"><%=FormatDateTime(mLast(mk), 2)%></span></td>
  <td style="text-align:right;white-space:nowrap">
    <a class="lnk primary" href="S_coach_syn.asp?t=list&add=<%=Server.URLEncode(mk)%>"><svg width="14" height="14"><use href="#i-plus"/></svg>補同義詞</a>
    <a class="lnk" target="_blank" href="S_coach.asp?q=<%=Server.URLEncode(mk)%>#chat"><svg width="14" height="14"><use href="#i-chat"/></svg>測試</a>
  </td>
</tr>
<%
Next
%>
    </tbody>
  </table>
  </div>
</section>
<%
End If
conn.Close
SC_Bottom

'---------------------------------------------------------------------
' 擷取用的小工具（放最後；VBScript 函式可在任何位置宣告）
'---------------------------------------------------------------------
' 中文詞去頭：只保留最後一個分隔字之後的部分，並去掉常見動詞開頭
Function CleanZh(z)
  Dim sepC, i, c, r, pre, again
  r = z
  sepC = "的之及與或和、，；：於在為予以由將對並等"
  For i = Len(r) To 1 Step -1
    c = Mid(r, i, 1)
    If InStr(sepC, c) > 0 Then r = Mid(r, i + 1): Exit For
  Next
  pre = Array("執行", "進行", "使用", "給予", "協助", "採用", "需", "應", "可", "若", "如", "含", "包括", "經", "再", "依", "即", "請", "有", "無", "是")
  again = True
  Do While again
    again = False
    For Each c In pre
      If Left(r, Len(c)) = c And Len(r) - Len(c) >= 2 Then r = Mid(r, Len(c) + 1): again = True
    Next
  Loop
  If Len(r) < 2 Or Len(r) > 12 Then r = ""
  CleanZh = r
End Function

Sub AddCand(z, e, cd)
  Dim kk
  kk = z & "|" & e
  If cands.Exists(kk) Then
    cands(kk) = cands(kk) + 1
    If InStr(" " & candCode(kk) & " ", " " & cd & " ") = 0 And Len(candCode(kk)) < 60 Then candCode(kk) = candCode(kk) & " " & cd
  Else
    cands.Add kk, 1
    candCode.Add kk, cd
  End If
End Sub
%>
