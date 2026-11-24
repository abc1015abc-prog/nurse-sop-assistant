<%@ Language="VBScript" CodePage="65001" EnableSessionState="False" %>
<!--#include file="S_coach_conn.asp"-->
<!--#include file="S_coach_layout.asp"-->
<!--#include file="S_coach_types.asp"-->
<%
'=====================================================================
' 我的提問紀錄（S_coach_my.asp）：只顯示登入者自己（cookie worknb）的紀錄
'=====================================================================
Response.Charset = "utf-8"
SC_CheckLogin
Dim conn, rs, me_wk, days, base, nAsk, nKw, nMiss, rsT, titles, cd, codesArr, res, mode
Set conn = SC_Open()
SC_EnsureLogCols conn
me_wk = SC_Worknb()
days = CLng("0" & Request("d")): If days <> 30 And days <> 90 And days <> 365 Then days = 90
base = "worknb='" & SQ(me_wk) & "' AND qtime >= Date()-" & days

nAsk = conn.Execute("SELECT COUNT(*) FROM qlog WHERE " & base & " AND qmode='ask'")(0)
nKw = conn.Execute("SELECT COUNT(*) FROM qlog WHERE " & base & " AND qmode='search'")(0)
nMiss = conn.Execute("SELECT COUNT(*) FROM qlog WHERE " & base & " AND qmode='ask' AND (result='找不到' OR result='low' OR result='無關鍵詞')")(0)

' 標準名稱對照（顯示「找到的標準」用）
Set titles = Server.CreateObject("Scripting.Dictionary")
Set rsT = conn.Execute("SELECT code, title FROM std_list")
Do Until rsT.EOF
  titles(rsT("code") & "") = rsT("title") & ""
  rsT.MoveNext
Loop
rsT.Close

Function DayChip(n)
  If n = days Then DayChip = "chip on" Else DayChip = "chip"
End Function

Function MyBadge(r, md)
  If md <> "ask" Then
    MyBadge = "<span class='badge neutral'>關鍵字查詢</span>"
  ElseIf r = "pending" Then
    MyBadge = "<span class='badge info pulse'>AI 整理中</span>"
  ElseIf r = "找不到" Or r = "無關鍵詞" Then
    MyBadge = "<span class='badge danger'>找不到</span>"
  ElseIf r = "low" Then
    MyBadge = "<span class='badge warning'>可能相關</span>"
  ElseIf InStr(r, "AI") > 0 Then
    MyBadge = "<span class='badge success'>AI 整理</span>"
  Else
    MyBadge = "<span class='badge info'>標準原文</span>"
  End If
End Function

SC_Top "my", "我的提問紀錄"
%>
<nav class="crumbs">查詢<svg width="14" height="14"><use href="#i-right"/></svg><span>我的提問紀錄</span></nav>
<div class="page-head">
  <div>
    <div class="page-title"><h1>我的提問紀錄</h1><span class="badge neutral">近 <%=days%> 天</span></div>
    <p class="page-sub">這裡只會顯示你自己（員工代號 <%=H(me_wk)%>）問過的問題與查過的關鍵字。點「查看回答」可以重看當時助理的整理與原文；離開頁面時 AI 還沒整理完也沒關係，完成後會自動存在這裡。</p>
  </div>
  <div class="head-actions">
    <a class="btn btn-primary" href="S_coach.asp#chat"><svg width="18" height="18"><use href="#i-chat"/></svg>問新的問題</a>
  </div>
</div>

<div class="stats" style="grid-template-columns:repeat(3,minmax(0,1fr))">
  <section class="card stat"><div class="k"><i><svg width="16" height="16"><use href="#i-chat"/></svg></i>提問</div><div class="v"><%=nAsk%><small>次</small></div></section>
  <section class="card stat"><div class="k"><i class="teal"><svg width="16" height="16"><use href="#i-key"/></svg></i>關鍵字查詢</div><div class="v"><%=nKw%><small>次</small></div></section>
  <section class="card stat"><div class="k"><i class="warn"><svg width="16" height="16"><use href="#i-alert"/></svg></i>標準中找不到</div><div class="v"><%=nMiss%><small>次</small></div><div class="d">可改用標準上的正式名稱再試，或詢問臨床教師</div></section>
</div>

<section class="card">
  <div class="toolbar">
    <label class="big-input"><svg width="18" height="18"><use href="#i-search"/></svg><input id="tf" placeholder="在我的紀錄中搜尋"></label>
    <span style="flex:1"></span>
    <div class="chips">
      <a class="<%=DayChip(30)%>" href="S_coach_my.asp?d=30">30 天</a>
      <a class="<%=DayChip(90)%>" href="S_coach_my.asp?d=90">90 天</a>
      <a class="<%=DayChip(365)%>" href="S_coach_my.asp?d=365">1 年</a>
    </div>
  </div>
  <div class="table-wrap">
  <table class="lst admin">
    <thead><tr><th style="width:140px">時間</th><th>問題 / AI 整理摘要</th><th class="hide-m">當時找到的標準</th><th style="width:110px">結果</th><th style="width:150px"></th></tr></thead>
    <tbody id="tb">
<%
Set rs = conn.Execute("SELECT TOP 500 * FROM qlog WHERE " & base & " ORDER BY qtime DESC")
If rs.EOF Then
  Response.Write "<tr><td colspan='5' class='empty'>這段期間還沒有紀錄，到「問助理」問第一個問題吧！</td></tr>"
End If
Do Until rs.EOF
  mode = rs("qmode") & ""
  res = rs("result") & ""
%>
<tr data-s="<%=H(LCase(rs("question") & " " & rs("top_codes")))%>">
  <td class="num" style="text-align:left"><%=FormatDateTime(rs("qtime"), 2)%><span class="cell-sub"><%=FormatDateTime(rs("qtime"), 4)%></span></td>
  <td class="q-cell"><b><%=H(rs("question"))%></b>
    <% If rs("ai_text") & "" <> "" Then %><span class="cell-sub" style="margin-top:4px;line-height:1.6"><%=H(Left(Replace(Replace(rs("ai_text") & "", "###", ""), vbLf, " "), 90))%>…</span><% End If %></td>
  <td class="hide-m"><%
    For Each cd In Split(Trim(rs("top_codes") & ""), " ")
      If cd <> "" Then
        Response.Write "<a class='rel-i' style='padding:4px 8px;margin:2px 0' target='_blank' href='S_coach_pdf.asp?code=" & Server.URLEncode(cd) & "'><span class='tag code'>" & H(cd) & "</span><span>"
        If titles.Exists(cd) Then Response.Write H(titles(cd))
        Response.Write "</span></a>"
      End If
    Next %></td>
  <td><%=MyBadge(res, mode)%></td>
  <td style="text-align:right">
    <% If mode = "ask" And rs("resp") & "" <> "" Then %>
      <a class="lnk primary" href="S_coach.asp?log=<%=rs("id")%>#chat"><svg width="14" height="14"><use href="#i-file"/></svg>查看回答</a>
      <a class="lnk" href="S_coach.asp?q=<%=Server.URLEncode(rs("question") & "")%>#chat" title="用最新版標準重新查詢"><svg width="14" height="14"><use href="#i-refresh"/></svg></a>
    <% ElseIf mode = "ask" Then %>
      <a class="lnk primary" href="S_coach.asp?q=<%=Server.URLEncode(rs("question") & "")%>#chat"><svg width="14" height="14"><use href="#i-refresh"/></svg>再問一次</a>
    <% Else %>
      <a class="lnk" href="S_coach.asp?kw=<%=Server.URLEncode(rs("question") & "")%>#kw"><svg width="14" height="14"><use href="#i-search"/></svg>再查一次</a>
    <% End If %>
  </td>
</tr>
<%
  rs.MoveNext
Loop
rs.Close: conn.Close
%>
    </tbody>
  </table>
  </div>
  <div class="empty" id="noRow" style="display:none">沒有符合的紀錄</div>
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
<% SC_Bottom %>
