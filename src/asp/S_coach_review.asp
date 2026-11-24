<%@ Language="VBScript" CodePage="65001" EnableSessionState="False" %>
<!--#include file="S_coach_conn.asp"-->
<!--#include file="S_coach_layout.asp"-->
<!--#include file="S_coach_types.asp"-->
<%
'=====================================================================
' 回饋與審核（S_coach_review.asp）
'   t=fb    護理師的 👍／👎 回饋（待處理優先）
'   t=qa    已審核答案（有效／需重新確認／停用）
'   t=edit  建立或修改審核答案（&log=提問紀錄 或 &id=答案編號）
'   審核答案引用的標準改版（修訂日或 PDF 檔案時間不同）→ 自動變成「需重新確認」
'=====================================================================
Response.Charset = "utf-8"
SC_CheckAdmin
Dim conn, act, t, msg, msgCls, me_wk
Set conn = SC_Open()
SC_EnsureLogCols conn
SC_EnsureFbTables conn
me_wk = SC_Worknb()
act = Request("act")
t = Request("t")
If t <> "qa" And t <> "edit" Then t = "fb"
msg = "": msgCls = "ok"
Select Case Request("msg")
  Case "saved"
    msg = "審核答案已儲存，之後有人問相近的問題會優先顯示"
  Case "done"
    msg = "已更新"
End Select

Function RIf(c, a, b)
  If c Then RIf = a Else RIf = b
End Function

'---------------------------------------------------------------------
' 動作
'---------------------------------------------------------------------
Dim pid, rsE, newId, pLog
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
  pid = CLng("0" & Request.Form("id"))
  If act = "qasave" Then
    Dim fQ, fK, fA, fR
    fQ = Left(Trim(SC_Half(Request.Form("question"))), 255)
    fK = Left(Trim(Replace(Replace(SC_Half(Request.Form("keywords")), "，", ","), "、", ",")), 255)
    fA = Trim(Request.Form("answer") & "")
    fR = Left(SC_CleanRefs(Request.Form("refs")), 255)
    pLog = CLng("0" & Request.Form("logid"))
    If fQ = "" Or fA = "" Then
      msg = "請填寫問題與答案": msgCls = "err": t = "edit"
    Else
      Set rsE = Server.CreateObject("ADODB.Recordset")
      rsE.Open "SELECT * FROM qa WHERE id=" & pid, conn, 1, 3
      If rsE.EOF Then
        rsE.AddNew
        rsE("hits") = 0
        rsE("src_log") = pLog
      End If
      rsE("question") = fQ
      rsE("keywords") = fK
      rsE("answer") = fA
      rsE("refs") = fR
      rsE("ref_ver") = SC_StdVer(conn, fR)
      rsE("reviewer") = me_wk
      rsE("rtime") = Now()
      rsE("status") = "ok"
      rsE.Update
      newId = CLng(rsE("id").Value)
      rsE.Close
      conn.Execute "UPDATE fb SET status='done', fix_qa=" & newId & " WHERE status='new' AND (" & RIf(pLog > 0, "logid=" & pLog & " OR ", "") & "qaid=" & newId & ")"
      conn.Close
      Response.Redirect "S_coach_review.asp?t=qa&msg=saved"
    End If
  ElseIf act = "fbdone" Then
    conn.Execute "UPDATE fb SET status='" & RIf(Request.Form("st") = "new", "new", "ignore") & "' WHERE id=" & pid
    conn.Close: Response.Redirect "S_coach_review.asp?t=fb&f=" & Server.URLEncode(Request.Form("f")) & "&msg=done"
  ElseIf act = "qastat" Then
    If Request.Form("st") = "off" Then
      conn.Execute "UPDATE qa SET status='off' WHERE id=" & pid
    Else
      ' 啟用／確認仍正確：記下目前標準版本
      Dim rr
      Set rr = conn.Execute("SELECT refs FROM qa WHERE id=" & pid)
      If Not rr.EOF Then
        conn.Execute "UPDATE qa SET status='ok', ref_ver='" & SQ(SC_StdVer(conn, rr("refs") & "")) & "', reviewer='" & SQ(me_wk) & "', rtime=Now() WHERE id=" & pid
      End If
      rr.Close
    End If
    conn.Close: Response.Redirect "S_coach_review.asp?t=qa&f=" & Server.URLEncode(Request.Form("f")) & "&msg=done"
  ElseIf act = "qadel" Then
    conn.Execute "DELETE FROM qa WHERE id=" & pid
    conn.Close: Response.Redirect "S_coach_review.asp?t=qa&msg=done"
  End If
End If

'---------------------------------------------------------------------
' 標準改版檢查：有效答案引用的標準版本變了 → 需重新確認
'---------------------------------------------------------------------
Dim rsV, vIds, vi
vIds = ""
Set rsV = conn.Execute("SELECT id, refs, ref_ver FROM qa WHERE status='ok'")
Do Until rsV.EOF
  If SC_StdVer(conn, rsV("refs") & "") <> rsV("ref_ver") & "" Then vIds = vIds & "," & rsV("id")
  rsV.MoveNext
Loop
rsV.Close
If vIds <> "" Then conn.Execute "UPDATE qa SET status='recheck' WHERE id IN (" & Mid(vIds, 2) & ")"

Dim nUp, nDown, nNew, nQa, nRe, rate
nUp = conn.Execute("SELECT COUNT(*) FROM fb WHERE rating=1 AND ftime >= Date()-30")(0)
nDown = conn.Execute("SELECT COUNT(*) FROM fb WHERE rating=-1 AND ftime >= Date()-30")(0)
nNew = conn.Execute("SELECT COUNT(*) FROM fb WHERE rating=-1 AND status='new'")(0)
nQa = conn.Execute("SELECT COUNT(*) FROM qa WHERE status='ok'")(0)
nRe = conn.Execute("SELECT COUNT(*) FROM qa WHERE status='recheck'")(0)
If nUp + nDown > 0 Then rate = Round(nUp * 100 / (nUp + nDown)) & "<small>%</small>" Else rate = "—"

Function TabCls(k)
  If k = t Or (k = "qa" And t = "edit") Then TabCls = "chip on" Else TabCls = "chip"
End Function

Dim stitle
stitle = "回饋與審核"
SC_Top "review", stitle
%>
<style>
.rv-reason{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px}
.rv-note{margin-top:6px;padding:8px 10px;border-radius:10px;background:var(--surface-2);color:var(--text-2);font-size:13px;white-space:pre-wrap}
.ed-grid{display:grid;grid-template-columns:minmax(0,1.15fr) minmax(0,1fr);gap:16px;align-items:start}
.ed-f label{display:block;font-size:13px;font-weight:700;color:var(--text-2);margin:14px 0 6px}
.ed-f label:first-child{margin-top:0}
.ed-f .hint{font-weight:400;color:var(--text-3);margin-left:6px}
.ed-f textarea.input{min-height:260px;line-height:1.7;resize:vertical}
.ed-ref .ex{border:1px solid var(--line-2);border-radius:12px;margin-bottom:10px;padding:10px 12px}
.ed-ref .ex .txt{max-height:12em;overflow:auto;white-space:pre-wrap;font-size:13px}
.ed-ai{white-space:pre-wrap;font-size:13.5px;background:var(--primary-50);border-radius:12px;padding:12px 14px;margin-bottom:14px}
@media (max-width:980px){.ed-grid{grid-template-columns:1fr}}
@media (max-width:760px){
  .lst.stack thead{display:none}
  .lst.stack tr{display:grid;grid-template-columns:1fr auto;gap:6px 10px;padding:12px 16px;border-bottom:1px solid var(--line)}
  .lst.stack td{display:block;padding:0!important;border:0!important;width:auto!important;text-align:left!important}
  .lst.stack td.hide-m{display:none}
  .lst.stack td.full{grid-column:1/-1}
}
</style>
<nav class="crumbs">管理<svg width="14" height="14"><use href="#i-right"/></svg><span>回饋與審核</span></nav>
<div class="page-head">
  <div>
    <div class="page-title"><h1>回饋與審核</h1><% If nNew > 0 Then %><span class="badge danger"><%=nNew%> 筆待處理</span><% End If %></div>
    <p class="page-sub">護理師對回答按「需要改進」後會出現在這裡。由臨床教師修正並存成「已審核答案」，之後有人問相近問題會優先顯示；引用的標準改版時會自動標示「需重新確認」。</p>
  </div>
</div>

<% If msg <> "" Then %>
<div class="banner <%=RIf(msgCls = "err", "err", "ok")%>" style="margin-bottom:16px"><svg width="18" height="18"><use href="#<%=RIf(msgCls = "err", "i-alert", "i-check")%>"/></svg><div><%=H(msg)%></div></div>
<% End If %>

<div class="stats s5">
  <section class="card stat"><div class="k"><i class="ok"><svg width="16" height="16"><use href="#i-up"/></svg></i>有幫助（30 天）</div><div class="v"><%=nUp%><small>次</small></div></section>
  <section class="card stat"><div class="k"><i class="warn"><svg width="16" height="16"><use href="#i-down"/></svg></i>需要改進（30 天）</div><div class="v"><%=nDown%><small>次</small></div></section>
  <section class="card stat"><div class="k"><i><svg width="16" height="16"><use href="#i-check"/></svg></i>好評率（30 天）</div><div class="v"><%=rate%></div></section>
  <section class="card stat"><div class="k"><i class="teal"><svg width="16" height="16"><use href="#i-shield"/></svg></i>已審核答案</div><div class="v"><%=nQa%><small>則</small></div></section>
  <section class="card stat"><div class="k"><i class="warn"><svg width="16" height="16"><use href="#i-alert"/></svg></i>標準改版待確認</div><div class="v"><%=nRe%><small>則</small></div></section>
</div>

<div class="chips" style="margin:16px 0">
  <a class="<%=TabCls("fb")%>" href="S_coach_review.asp?t=fb">護理師回饋<% If nNew > 0 Then %><span class="n"><%=nNew%></span><% End If %></a>
  <a class="<%=TabCls("qa")%>" href="S_coach_review.asp?t=qa">已審核答案<% If nRe > 0 Then %><span class="n"><%=nRe%></span><% End If %></a>
</div>

<%
'=====================================================================
If t = "fb" Then
'=====================================================================
Dim f, wh, rsF, rsn
f = Request("f")
Select Case f
  Case "all"
    wh = "1=1"
  Case "up"
    wh = "rating=1"
  Case "down"
    wh = "rating=-1"
  Case Else
    f = "new": wh = "rating=-1 AND status='new'"
End Select
%>
<section class="card">
  <div class="toolbar">
    <div class="chips">
      <a class="chip<%=RIf(f = "new", " on", "")%>" href="S_coach_review.asp?t=fb&f=new">待處理</a>
      <a class="chip<%=RIf(f = "down", " on", "")%>" href="S_coach_review.asp?t=fb&f=down">需要改進</a>
      <a class="chip<%=RIf(f = "up", " on", "")%>" href="S_coach_review.asp?t=fb&f=up">有幫助</a>
      <a class="chip<%=RIf(f = "all", " on", "")%>" href="S_coach_review.asp?t=fb&f=all">全部</a>
    </div>
  </div>
  <div class="table-wrap">
  <table class="lst admin stack">
    <thead><tr><th style="width:140px">時間</th><th>提問與回饋內容</th><th class="hide-m" style="width:160px">當時找到的標準</th><th style="width:110px">評價</th><th style="width:220px"></th></tr></thead>
    <tbody>
<%
Set rsF = conn.Execute("SELECT TOP 300 * FROM fb WHERE " & wh & " ORDER BY ftime DESC")
If rsF.EOF Then Response.Write "<tr><td colspan='5' class='empty'>" & RIf(f = "new", "沒有待處理的回饋 🎉", "沒有資料") & "</td></tr>"
Do Until rsF.EOF
%>
<tr>
  <td class="num" style="text-align:left"><%=FormatDateTime(rsF("ftime"), 2)%><span class="cell-sub"><%=H(rsF("unit"))%>・<%=H(rsF("worknb"))%></span></td>
  <td class="full"><b><%=H(rsF("question"))%></b>
    <% If rsF("reason") & "" <> "" Then %><div class="rv-reason"><% For Each rsn In Split(rsF("reason") & "", "、")
      Response.Write "<span class='badge warning'>" & H(rsn) & "</span>"
    Next %></div><% End If %>
    <% If rsF("note") & "" <> "" Then %><div class="rv-note"><%=H(rsF("note"))%></div><% End If %>
    <% If CLng("0" & rsF("qaid")) > 0 Then %><div class="note" style="margin-top:4px">當時顯示的是已審核答案 #<%=rsF("qaid")%></div><% End If %>
  </td>
  <td class="hide-m"><span class="note"><%=H(rsF("codes"))%></span></td>
  <td><% If rsF("rating") = 1 Then %><span class="badge success">有幫助</span><% Else %><span class="badge danger">需要改進</span><% End If %>
    <% If rsF("status") = "done" And rsF("rating") = -1 Then %><span class="cell-sub">已修正</span><% ElseIf rsF("status") = "ignore" Then %><span class="cell-sub">已略過</span><% End If %></td>
  <td style="text-align:right;white-space:nowrap">
    <a class="lnk" target="_blank" href="S_coach.asp?log=<%=rsF("logid")%>#chat"><svg width="14" height="14"><use href="#i-chat"/></svg>查看回答</a>
    <% If rsF("rating") = -1 Then %>
      <% If CLng("0" & rsF("qaid")) > 0 Then %>
        <a class="lnk primary" href="S_coach_review.asp?t=edit&id=<%=rsF("qaid")%>&log=<%=rsF("logid")%>"><svg width="14" height="14"><use href="#i-shield"/></svg>修改審核答案</a>
      <% Else %>
        <a class="lnk primary" href="S_coach_review.asp?t=edit&log=<%=rsF("logid")%>"><svg width="14" height="14"><use href="#i-shield"/></svg>建立審核答案</a>
      <% End If %>
      <form method="post" action="S_coach_review.asp" style="display:inline">
        <input type="hidden" name="act" value="fbdone"><input type="hidden" name="id" value="<%=rsF("id")%>"><input type="hidden" name="f" value="<%=H(f)%>">
        <% If rsF("status") = "new" Then %><button class="lnk" title="不需修正">略過</button><% Else %><input type="hidden" name="st" value="new"><button class="lnk">改回待處理</button><% End If %>
      </form>
    <% End If %>
  </td>
</tr>
<%
  rsF.MoveNext
Loop
rsF.Close
%>
    </tbody>
  </table>
  </div>
</section>
<%
'=====================================================================
ElseIf t = "qa" Then
'=====================================================================
Dim qf, wq, rsQ, pr2, cd2
qf = Request("f")
Select Case qf
  Case "ok"
    wq = "status='ok'"
  Case "recheck"
    wq = "status='recheck'"
  Case "off"
    wq = "status='off'"
  Case Else
    qf = "": wq = "1=1"
End Select
%>
<section class="card">
  <div class="toolbar">
    <label class="big-input"><svg width="18" height="18"><use href="#i-search"/></svg><input id="tf" placeholder="搜尋問題、關鍵詞、標準編號"></label>
    <span style="flex:1"></span>
    <div class="chips">
      <a class="chip<%=RIf(qf = "", " on", "")%>" href="S_coach_review.asp?t=qa">全部</a>
      <a class="chip<%=RIf(qf = "ok", " on", "")%>" href="S_coach_review.asp?t=qa&f=ok">有效</a>
      <a class="chip<%=RIf(qf = "recheck", " on", "")%>" href="S_coach_review.asp?t=qa&f=recheck">需重新確認</a>
      <a class="chip<%=RIf(qf = "off", " on", "")%>" href="S_coach_review.asp?t=qa&f=off">停用</a>
    </div>
    <a class="btn btn-primary btn-sm" href="S_coach_review.asp?t=edit"><svg width="16" height="16"><use href="#i-plus"/></svg>新增</a>
  </div>
  <div class="table-wrap">
  <table class="lst admin stack">
    <thead><tr><th>問題／比對關鍵詞</th><th class="hide-m" style="width:200px">出處</th><th style="width:120px">狀態</th><th class="hide-m" style="width:80px;text-align:right">使用</th><th class="hide-m" style="width:120px">審核</th><th style="width:200px"></th></tr></thead>
    <tbody id="tb">
<%
Set rsQ = conn.Execute("SELECT id, question, keywords, refs, reviewer, rtime, status, hits FROM qa WHERE " & wq & " ORDER BY IIf(status='recheck',0,IIf(status='ok',1,2)), rtime DESC")
If rsQ.EOF Then Response.Write "<tr><td colspan='6' class='empty'>還沒有審核答案。從「護理師回饋」按「建立審核答案」，或按右上角「新增」</td></tr>"
Do Until rsQ.EOF
%>
<tr data-s="<%=H(LCase(rsQ("question") & " " & rsQ("keywords") & " " & rsQ("refs")))%>">
  <td class="full"><b><%=H(rsQ("question"))%></b><% If rsQ("keywords") & "" <> "" Then %><span class="cell-sub">關鍵詞：<%=H(rsQ("keywords"))%></span><% End If %></td>
  <td class="hide-m"><% For Each pr2 In Split(rsQ("refs") & "", ";")
      cd2 = pr2
      If InStr(cd2, ":") > 0 Then cd2 = Left(cd2, InStr(cd2, ":") - 1)
      If cd2 <> "" Then Response.Write "<a class='tag code' style='margin:2px' target='_blank' href='S_coach_pdf.asp?code=" & Server.URLEncode(cd2) & "'>" & H(pr2) & "</a>"
    Next %></td>
  <td><% Select Case rsQ("status") & ""
      Case "ok"
        Response.Write "<span class='badge success'>有效</span>"
      Case "recheck"
        Response.Write "<span class='badge warning'>標準改版，需重新確認</span>"
      Case Else
        Response.Write "<span class='badge neutral'>停用</span>"
    End Select %></td>
  <td class="num hide-m"><%=CLng("0" & rsQ("hits"))%></td>
  <td class="hide-m"><span class="note"><%=H(rsQ("reviewer"))%><br><%=FormatDateTime(rsQ("rtime"), 2)%></span></td>
  <td style="text-align:right;white-space:nowrap">
    <a class="lnk primary" href="S_coach_review.asp?t=edit&id=<%=rsQ("id")%>">編輯</a>
    <form method="post" action="S_coach_review.asp" style="display:inline">
      <input type="hidden" name="act" value="qastat"><input type="hidden" name="id" value="<%=rsQ("id")%>"><input type="hidden" name="f" value="<%=H(qf)%>">
      <% If rsQ("status") = "recheck" Then %>
        <button class="lnk" title="已核對新版標準，內容仍正確">確認仍正確</button>
      <% ElseIf rsQ("status") = "off" Then %>
        <button class="lnk">啟用</button>
      <% Else %>
        <input type="hidden" name="st" value="off"><button class="lnk">停用</button>
      <% End If %>
    </form>
    <form method="post" action="S_coach_review.asp" style="display:inline" onsubmit="return confirm('確定刪除這則審核答案？')">
      <input type="hidden" name="act" value="qadel"><input type="hidden" name="id" value="<%=rsQ("id")%>"><button class="lnk">刪除</button>
    </form>
  </td>
</tr>
<%
  rsQ.MoveNext
Loop
rsQ.Close
%>
    </tbody>
  </table>
  </div>
  <div class="empty" id="noRow" style="display:none">沒有符合的答案</div>
</section>
<script>
(function(){
  var tf=document.getElementById('tf'), rows=[].slice.call(document.querySelectorAll('#tb tr[data-s]'));
  tf.oninput=function(){
    var q=tf.value.trim().toLowerCase(), n=0;
    rows.forEach(function(r){var ok=!q||r.dataset.s.indexOf(q)>=0;r.style.display=ok?'':'none';if(ok)n++;});
    document.getElementById('noRow').style.display=(rows.length&&!n)?'block':'none';
  };
})();
</script>
<%
'=====================================================================
Else   ' t = "edit"
'=====================================================================
Dim eId, eLog, eQ, eK, eA, eR, eSt, respTxt, aiTxt, rsL, isNew
eId = CLng("0" & Request("id"))
eLog = CLng("0" & Request("log"))
eQ = "": eK = "": eA = "": eR = "": eSt = "": respTxt = "": aiTxt = ""
If eId > 0 Then
  Set rsL = conn.Execute("SELECT question, keywords, answer, refs, status, src_log FROM qa WHERE id=" & eId)
  If Not rsL.EOF Then
    eQ = rsL("question") & "": eK = rsL("keywords") & "": eA = rsL("answer") & "": eR = rsL("refs") & "": eSt = rsL("status") & ""
    If eLog = 0 Then eLog = CLng("0" & rsL("src_log"))
  Else
    eId = 0
  End If
  rsL.Close
End If
If eLog > 0 Then
  Set rsL = conn.Execute("SELECT question, resp, ai_text FROM qlog WHERE id=" & eLog)
  If Not rsL.EOF Then
    If eQ = "" Then eQ = rsL("question") & ""
    respTxt = rsL("resp") & ""
    aiTxt = rsL("ai_text") & ""
  End If
  rsL.Close
End If
If msgCls = "err" Then
  eQ = Request.Form("question"): eK = Request.Form("keywords"): eA = Request.Form("answer"): eR = Request.Form("refs")
End If
isNew = (eId = 0)
If respTxt = "" Then respTxt = "null"
respTxt = Replace(respTxt, "</", "<\/")
%>
<div class="ed-grid">
<section class="card">
  <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-shield"/></svg></div><div><h2><%=RIf(isNew, "建立審核答案", "修改審核答案 #" & eId)%></h2><p>請對照右側標準原文修正。答案只能寫標準上有的內容，並寫明出處標準與頁碼。<% If eSt = "recheck" Then %><b style="color:var(--warning-700)">引用的標準已改版，請核對新版後儲存。</b><% End If %></p></div></div>
  <div class="card-body">
    <form method="post" action="S_coach_review.asp" class="ed-f" id="edf">
      <input type="hidden" name="act" value="qasave"><input type="hidden" name="id" value="<%=eId%>"><input type="hidden" name="logid" value="<%=eLog%>">
      <label>問題<span class="hint">用護理師常見的問法</span></label>
      <input class="input" name="question" id="eQ" value="<%=H(eQ)%>" maxlength="250" required>
      <label>比對關鍵詞<span class="hint">用逗號分隔；新問題「全部都有」這些詞時就會顯示此答案（縮寫、同義詞也算）</span></label>
      <input class="input" name="keywords" id="eK" value="<%=H(eK)%>" maxlength="250" placeholder="例：鼻胃管, 胃殘餘量">
      <label>答案<span class="hint">可用「- 」條列、「### 」小標；建議在句尾寫出處，如（CP016 第6頁）</span></label>
      <textarea class="input" name="answer" id="eA" required><%=H(eA)%></textarea>
      <label>出處<span class="hint">格式：標準編號:頁碼，多份用分號，例 CP016:6,7;CP018:2</span></label>
      <input class="input" name="refs" id="eR" value="<%=H(eR)%>" maxlength="250">
      <div style="display:flex;gap:8px;margin-top:18px;flex-wrap:wrap">
        <button class="btn btn-primary"><svg width="18" height="18"><use href="#i-check"/></svg>儲存並設為有效</button>
        <a class="btn btn-ghost" href="S_coach_review.asp?t=<%=RIf(eLog > 0 And isNew, "fb", "qa")%>">取消</a>
        <% If eQ <> "" Then %><a class="btn btn-ghost" target="_blank" href="S_coach.asp?q=<%=Server.URLEncode(eQ)%>#chat" title="儲存後可用這個測試"><svg width="18" height="18"><use href="#i-chat"/></svg>測試提問</a><% End If %>
      </div>
    </form>
  </div>
</section>
<section class="card ed-ref">
  <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-file"/></svg></div><div><h2>當時的回答與標準原文</h2><p>核對用；點標準可開啟 PDF。</p></div></div>
  <div class="card-body" id="refBox"><div class="note">沒有當時的回答紀錄（舊紀錄或手動新增），請直接開啟標準 PDF 核對。</div></div>
</section>
</div>
<script>
(function(){
  var R=<%=respTxt%>, AI=<%=Replace(JE(aiTxt), "</", "<\/")%>, isNew=<%=RIf(isNew, "true", "false")%>;
  function h(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
  function pdf(c,p){return 'S_coach_pdf.asp?code='+encodeURIComponent(c)+(p?'#page='+p:'');}
  if(!R) return;
  var byN={}; (R.sources||[]).forEach(function(s){byN[s.n]=s;});
  // AI 文字的 [來源n] 換成（編號 第p頁）
  var aiPlain=String(AI||'').replace(/\[來源\s*(\d+)\]/g,function(m,n){var s=byN[+n];return s?'（'+s.code+' 第'+s.page+'頁）':'';});
  var box=document.getElementById('refBox'), x='';
  if(aiPlain) x+='<div class="sec-title">'+'AI 當時的整理</div><div class="ed-ai">'+h(aiPlain)+'</div>';
  if(R.qa) x+='<div class="sec-title">當時顯示的審核答案 #'+R.qa.id+'</div><div class="ed-ai" style="background:var(--success-50)">'+h(R.qa.answer)+'</div>';
  var src=R.sources||[];
  if(src.length){
    x+='<div class="sec-title">AI 參考的標準原文</div>';
    src.forEach(function(s){x+='<div class="ex"><div class="ex-h"><span class="tag code">'+h(s.code)+'</span>'+h(s.title)+'・第 '+s.page+' 頁<a target="_blank" href="'+pdf(s.code,s.page)+'">開啟 PDF ›</a></div><div class="txt">'+h(s.text)+'</div></div>';});
  }
  if(R.stds&&R.stds.length){
    x+='<div class="sec-title">當時找到的標準</div><div class="rel">';
    R.stds.forEach(function(s){x+='<a class="rel-i" target="_blank" href="'+pdf(s.code,s.pages[0])+'"><span class="tag code">'+h(s.code)+'</span><span>'+h(s.title)+'　第 '+s.pages.join('、')+' 頁</span></a>';});
    x+='</div>';
  }
  if(x) box.innerHTML=x;
  if(!isNew) return;
  // 新建：自動帶入答案、出處、關鍵詞（請再核對修改）
  var eA=document.getElementById('eA'), eR=document.getElementById('eR'), eK=document.getElementById('eK');
  if(!eA.value&&aiPlain) eA.value=aiPlain.replace(/^#{1,4}\s*/gm,'### ');
  if(!eR.value){
    var used={}, order=[];
    String(AI||'').replace(/\[來源\s*(\d+)\]/g,function(m,n){var s=byN[+n];if(s){if(!used[s.code]){used[s.code]=[];order.push(s.code);}if(used[s.code].indexOf(s.page)<0)used[s.code].push(s.page);}return m;});
    if(!order.length&&R.primary&&R.primary.code){order.push(R.primary.code);used[R.primary.code]=src.filter(function(s){return s.code===R.primary.code;}).map(function(s){return s.page;});}
    eR.value=order.map(function(c){return c+(used[c].length?':'+used[c].sort(function(a,b){return a-b;}).join(','):'');}).join(';');
  }
  if(!eK.value&&R.concepts){
    var c=R.concepts.slice().sort(function(a,b){return (b.syn-a.syn)||(b.wt-a.wt);}).map(function(x){return x.w;});
    c=c.filter(function(w,i){return w.length>=2&&!c.some(function(y,j){return j<i&&y.indexOf(w)>=0;});}).slice(0,3);
    eK.value=c.join(', ');
  }
})();
</script>
<%
End If
conn.Close
SC_Bottom
%>
