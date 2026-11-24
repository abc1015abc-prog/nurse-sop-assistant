<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="S_coach_conn.asp"-->
<!--#include file="S_coach_layout.asp"-->
<!--#include file="S_coach_types.asp"-->
<%
Response.Charset = "utf-8"
SC_CheckAdmin
Dim conn, rs, f, days, base, where, nAll, nAsk, nMiss, nUser, nAI, missPct, res, isMiss, rb, cd
Set conn = SC_Open()
SC_EnsureLogCols conn
f = Request("f"): If f <> "ask" And f <> "miss" And f <> "kw" Then f = ""
days = CLng("0" & Request("d")): If days <> 7 And days <> 30 And days <> 90 And days <> 365 Then days = 30
base = "qtime >= Date()-" & days
Const MISS_SQL = "(result='找不到' OR result='low' OR result='無關鍵詞')"

nAll = conn.Execute("SELECT COUNT(*) FROM qlog WHERE " & base)(0)
nAsk = conn.Execute("SELECT COUNT(*) FROM qlog WHERE " & base & " AND qmode='ask'")(0)
nMiss = conn.Execute("SELECT COUNT(*) FROM qlog WHERE " & base & " AND qmode='ask' AND " & MISS_SQL)(0)
nAI = conn.Execute("SELECT COUNT(*) FROM qlog WHERE " & base & " AND qmode='ask' AND result LIKE '%AI%'")(0)
nUser = conn.Execute("SELECT COUNT(*) FROM (SELECT DISTINCT worknb FROM qlog WHERE " & base & ") AS t")(0)
If nAsk > 0 Then missPct = Int(nMiss / nAsk * 100) Else missPct = 0

where = base
If f = "ask" Then where = where & " AND qmode='ask'"
If f = "kw" Then where = where & " AND qmode='search'"
If f = "miss" Then where = where & " AND qmode='ask' AND " & MISS_SQL

Function Lnk(nf, nd)
  Lnk = "S_coach_log.asp?f=" & nf & "&d=" & nd
End Function
Function ChipCls(a, b)
  If CStr(a) = CStr(b) Then ChipCls = "chip on" Else ChipCls = "chip"
End Function

Function ResBadge(r, mode)
  If mode <> "ask" Then
    ResBadge = "<span class='badge neutral'>" & H(r) & "</span>"
  ElseIf r = "pending" Then
    ResBadge = "<span class='badge info'>AI 整理中</span>"
  ElseIf r = "找不到" Or r = "無關鍵詞" Then
    ResBadge = "<span class='badge danger'>找不到</span>"
  ElseIf r = "low" Then
    ResBadge = "<span class='badge warning'>可能相關</span>"
  ElseIf InStr(r, "AI") > 0 Then
    ResBadge = "<span class='badge success'>AI 整理</span>"
  Else
    ResBadge = "<span class='badge info'>原文</span>"
  End If
End Function

SC_Top "log", "提問紀錄"
%>
<nav class="crumbs">管理<svg width="14" height="14"><use href="#i-right"/></svg><span>提問紀錄</span></nav>
<div class="page-head">
  <div>
    <div class="page-title"><h1>提問紀錄</h1><span class="badge neutral">近 <%=days%> 天</span></div>
    <p class="page-sub">「找不到」的問題代表標準未涵蓋，或護理師的用語與標準不同。可請臨床教師回覆、列入標準修訂參考，或到 <a href="S_coach_syn.asp?t=miss">同義詞管理</a> 補上縮寫與俗稱。</p>
  </div>
  <div class="head-actions">
    <a class="btn btn-secondary" href="S_coach_build.asp"><svg width="18" height="18"><use href="#i-db"/></svg>索引管理</a>
  </div>
</div>

<div class="stats">
  <section class="card stat"><div class="k"><i><svg width="16" height="16"><use href="#i-chat"/></svg></i>提問次數</div><div class="v"><%=nAsk%><small>次</small></div><div class="d">另有關鍵字查詢 <%=nAll - nAsk%> 次</div></section>
  <section class="card stat"><div class="k"><i class="teal"><svg width="16" height="16"><use href="#i-users"/></svg></i>使用人數</div><div class="v"><%=nUser%><small>人</small></div><div class="d">依員工代號計算</div></section>
  <section class="card stat"><div class="k"><i class="ok"><svg width="16" height="16"><use href="#i-spark"/></svg></i>AI 整理</div><div class="v"><%=nAI%><small>次</small></div><div class="d">其餘僅顯示標準原文</div></section>
  <section class="card stat"><div class="k"><i class="err"><svg width="16" height="16"><use href="#i-alert"/></svg></i>找不到 / 可能相關</div><div class="v"><%=nMiss%><small><%=missPct%>%</small></div><div class="bar"><i style="width:<%=missPct%>%;background:linear-gradient(90deg,#F59E0B,#DC2626)"></i></div></section>
</div>

<section class="card">
  <div class="toolbar">
    <div class="chips">
      <a class="<%=ChipCls(f, "")%>" href="<%=Lnk("", days)%>">全部</a>
      <a class="<%=ChipCls(f, "ask")%>" href="<%=Lnk("ask", days)%>">提問</a>
      <a class="<%=ChipCls(f, "kw")%>" href="<%=Lnk("kw", days)%>">關鍵字查詢</a>
      <a class="<%=ChipCls(f, "miss")%>" href="<%=Lnk("miss", days)%>">只看找不到的</a>
    </div>
    <span style="flex:1"></span>
    <div class="chips">
      <a class="<%=ChipCls(days, 7)%>" href="<%=Lnk(f, 7)%>">7 天</a>
      <a class="<%=ChipCls(days, 30)%>" href="<%=Lnk(f, 30)%>">30 天</a>
      <a class="<%=ChipCls(days, 90)%>" href="<%=Lnk(f, 90)%>">90 天</a>
      <a class="<%=ChipCls(days, 365)%>" href="<%=Lnk(f, 365)%>">1 年</a>
    </div>
  </div>
  <div class="table-wrap">
  <table class="lst admin">
    <thead><tr><th style="width:150px">時間</th><th class="hide-m" style="width:130px">提問者</th><th>問題 / 關鍵字</th><th class="hide-m">找到的標準</th><th class="hide-m" style="width:90px;text-align:right">命中率</th><th style="width:110px">結果</th></tr></thead>
    <tbody>
<%
Set rs = conn.Execute("SELECT TOP 1000 * FROM qlog WHERE " & where & " ORDER BY qtime DESC")
If rs.EOF Then
  Response.Write "<tr><td colspan='6' class='empty'>這段期間沒有紀錄</td></tr>"
End If
Do Until rs.EOF
  rb = ResBadge(rs("result") & "", rs("qmode") & "")
%>
<tr>
  <td class="num" style="text-align:left"><%=FormatDateTime(rs("qtime"), 2)%><span class="cell-sub"><%=FormatDateTime(rs("qtime"), 4)%></span></td>
  <td class="hide-m"><%=H(rs("worknb"))%><span class="cell-sub"><%=H(rs("unit"))%></span></td>
  <td class="q-cell"><% If rs("qmode") & "" = "ask" Then %><span class="badge info" style="margin-right:6px">提問</span><% Else %><span class="badge neutral" style="margin-right:6px">關鍵字</span><% End If %><%=H(rs("question"))%><% If rs("qmode") & "" = "ask" Then %> <a class="note" href="S_coach.asp?log=<%=rs("id")%>#chat" target="_blank">查看回答 ›</a><% End If %></td>
  <td class="hide-m"><% For Each cd In Split(Trim(rs("top_codes") & ""), " ")
       If cd <> "" Then Response.Write "<span class='tag code' style='margin:2px'>" & H(cd) & "</span>"
     Next %></td>
  <td class="hide-m num"><% If rs("qmode") & "" = "ask" Then Response.Write rs("cover") & "%" %></td>
  <td><%=rb%></td>
</tr>
<%
  rs.MoveNext
Loop
rs.Close: conn.Close
%>
    </tbody>
  </table>
  </div>
</section>
<% SC_Bottom %>
