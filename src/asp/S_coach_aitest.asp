<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="S_coach_conn.asp"-->
<!--#include file="S_coach_layout.asp"-->
<%
Response.Charset = "utf-8"
Server.ScriptTimeout = 400
SC_CheckAdmin

Dim baseUrl, x, t0, body, model
Dim s1State, s1Msg, s1Secs, s2State, s2Msg, s2Secs, s2Raw, s2Answer, genSpd, readSpd, loadSecs
model = Trim(Request("model")): If model = "" Then model = SC_MODEL
baseUrl = SC_AI_URL
If InStr(baseUrl, "/api/") > 0 Then baseUrl = Left(baseUrl, InStr(baseUrl, "/api/") - 1)
If InStr(baseUrl, "/v1/") > 0 Then baseUrl = Left(baseUrl, InStr(baseUrl, "/v1/") - 1)
s2State = ""

' 會「先思考」的模型關掉思考模式（與 S_coach_api.asp 相同規則）
Function TkOpt(md)
  Dim m: m = LCase(md)
  TkOpt = ""
  If InStr(m, "instruct") > 0 Then Exit Function
  If InStr(m, "gpt-oss") > 0 Then
    TkOpt = ",""think"":""low"""
  ElseIf InStr(m, "qwen3") > 0 Or InStr(m, "deepseek") > 0 Or InStr(m, "gemma4") > 0 Or InStr(m, "gemma-4") > 0 Or InStr(m, "magistral") > 0 Or InStr(m, "reasoning") > 0 Then
    TkOpt = ",""think"":false"
  End If
End Function

Function IIf3(c, a, b)
  If c Then IIf3 = a Else IIf3 = b
End Function

Function ReNum(js, pat)
  Dim re, m
  Set re = New RegExp: re.Pattern = pat
  Set m = re.Execute(js)
  If m.Count > 0 Then ReNum = CDbl(m(0).SubMatches(0)) Else ReNum = -1
End Function

' 從 Ollama 回傳內容取出速度與答案
Sub ParseOllama(js)
  Dim ec, ed, pc, pd, ld, re, m
  ec = ReNum(js, """eval_count"":(\d+)"): ed = ReNum(js, """eval_duration"":(\d+)")
  pc = ReNum(js, """prompt_eval_count"":(\d+)"): pd = ReNum(js, """prompt_eval_duration"":(\d+)")
  ld = ReNum(js, """load_duration"":(\d+)")
  If ec > 0 And ed > 0 Then genSpd = Round(ec / (ed / 1e9), 1)
  If pc > 0 And pd > 0 Then readSpd = Round(pc / (pd / 1e9), 0)
  If ld >= 0 Then loadSecs = Round(ld / 1e9, 1)
  Set re = New RegExp: re.Pattern = """content"":""((?:[^""\\]|\\.)*)"""
  Set m = re.Execute(js)
  If m.Count > 0 Then s2Answer = Replace(Replace(m(0).SubMatches(0), "\n", vbLf), "\""", """")
End Sub

Function NewX(timeoutSec)
  Dim o
  Set o = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
  If SC_PROXY <> "" Then o.setProxy 2, SC_PROXY, "" Else o.setProxy 1
  o.setTimeouts 5000, 5000, 10000, timeoutSec * 1000
  Set NewX = o
End Function

' 步驟 1：連線
On Error Resume Next
Set x = NewX(10)
t0 = Timer
x.open "GET", baseUrl & "/", False
x.send
s1Secs = Round(Timer - t0, 1)
If Err.Number <> 0 Then
  s1State = "err"
  s1Msg = H(Err.Description)
Else
  s1State = "ok"
  s1Msg = "HTTP " & x.status & "「" & H(Left(x.responseText, 60)) & "」"
End If
Err.Clear

' 步驟 2：請模型回答一題
If s1State = "ok" And Request("go") = "1" Then
  Set x = NewX(SC_AI_TIMEOUT)
  body = "{""model"":" & JE(model) & ",""stream"":false"
  body = body & TkOpt(model)
  body = body & ",""messages"":[{""role"":""user"",""content"":""請用 3 點條列說明洗手的目的，請用繁體中文" & IIf3(InStr(LCase(model), "qwen3") > 0 And InStr(LCase(model), "instruct") = 0, " /no_think", "") & """}],""options"":{""num_predict"":200" & _
         IIf3(SC_AI_THREADS > 0, ",""num_thread"":" & SC_AI_THREADS, "") & "}}"
  t0 = Timer
  x.open "POST", baseUrl & "/api/chat", False
  x.setRequestHeader "content-type", "application/json; charset=utf-8"
  x.send body
  s2Secs = Round(Timer - t0, 1)
  If Err.Number <> 0 Then
    s2State = "err"
    s2Msg = H(Err.Description) & "（等了 " & Int(s2Secs) & " 秒）。若等滿 " & SC_AI_TIMEOUT & " 秒，代表模型太慢或正在載入：請改用較小的模型，並設定 OLLAMA_KEEP_ALIVE=-1。"
  ElseIf x.status <> 200 Then
    s2State = "err"
    s2Msg = "HTTP " & x.status & "：" & H(Left(x.responseText, 300)) & "<br>若顯示 model not found，請在伺服器執行 <code>ollama pull " & H(model) & "</code>"
  Else
    s2Raw = x.responseText
    ParseOllama s2Raw
    If InStr(s2Raw, """content"":""""") > 0 Then
      s2State = "warn"
      s2Msg = "模型只輸出思考過程（thinking），沒有產生答案。請確認已更新最新版程式（會自動關閉思考模式）；若仍如此，請改用其他模型。"
    Else
      s2State = "ok"
      s2Msg = "模型已回答"
    End If
  End If
  Err.Clear
End If
On Error GoTo 0

SC_Top "aitest", "AI 連線測試"
%>
<nav class="crumbs">管理<svg width="14" height="14"><use href="#i-right"/></svg><span>AI 連線測試</span></nav>
<div class="page-head">
  <div>
    <div class="page-title"><h1>AI 連線測試</h1>
      <% If Not SC_AI_ENABLED Then %><span class="badge neutral">AI 未啟用</span><% ElseIf s1State = "ok" Then %><span class="badge success">Ollama 運作中</span><% Else %><span class="badge danger">連不到 AI</span><% End If %>
    </div>
    <p class="page-sub">檢查網站能否連到院內 Ollama，並實測模型回答速度。第一次測試包含載入模型的時間，馬上再測一次的秒數，才是平常每題要等的時間。</p>
  </div>
  <div class="head-actions">
    <a class="btn btn-secondary" href="S_coach.asp#chat"><svg width="18" height="18"><use href="#i-chat"/></svg>前往問助理</a>
  </div>
</div>

<div class="two-col">
  <section class="card">
    <div class="test-step">
      <span class="n <%=s1State%>"><% If s1State = "ok" Then %><svg width="18" height="18"><use href="#i-check"/></svg><% Else %><svg width="18" height="18"><use href="#i-x"/></svg><% End If %></span>
      <div class="body">
        <b>① 連線</b>
        <% If s1State = "ok" Then %>
          <span class="badge success">連得到</span> <span class="note"><%=s1Msg%>・<%=s1Secs%> 秒</span>
        <% Else %>
          <span class="badge danger">連不到</span> <span class="note"><%=s1Msg%>・<%=s1Secs%> 秒</span>
          <div class="banner warn" style="margin-top:10px"><svg width="16" height="16"><use href="#i-info"/></svg><div>請檢查：伺服器上的 Ollama 是否在執行（工作排程器的 ollama serve）、<code>SC_AI_URL</code> 是否正確；若 Ollama 裝在其他電腦，需設定 OLLAMA_HOST=0.0.0.0 並開放防火牆 11434 埠。</div></div>
        <% End If %>
      </div>
    </div>
    <div class="test-step">
      <span class="n <%=IIf3(s2State = "ok", "ok", IIf3(s2State <> "", "err", ""))%>"><% If s2State = "ok" Then %><svg width="18" height="18"><use href="#i-check"/></svg><% Else %>2<% End If %></span>
      <div class="body">
        <b>② 模型回答</b>
        <form class="inline-form" style="margin:6px 0 4px;flex-wrap:wrap" onsubmit="var b=this.querySelector('button');b.classList.add('loading');b.disabled=true;">
          <input class="input" name="model" value="<%=H(model)%>" style="flex:1;min-width:220px" aria-label="模型名稱">
          <input type="hidden" name="go" value="1">
          <button class="btn btn-primary" <% If s1State <> "ok" Then Response.Write "disabled" %>><span class="spin"></span><svg width="18" height="18" class="lbl-ico"><use href="#i-bolt"/></svg>開始測試</button>
        </form>
        <div class="note">需先在伺服器執行 <code>ollama pull 模型名稱</code>；可輸入其他模型名稱比較速度。按下後最久等 <%=SC_AI_TIMEOUT%> 秒。</div>

        <% If s2State <> "" Then %>
          <div style="margin-top:14px">
          <% If s2State = "ok" Then %><span class="badge success">回答成功</span><% ElseIf s2State = "warn" Then %><span class="badge warning">沒有產生答案</span><% Else %><span class="badge danger">失敗</span><% End If %>
          <span class="note"><%=s2Msg%></span>
          </div>
          <% If s2State <> "err" Then %>
          <div class="metric-row">
            <div class="metric"><span>總耗時</span><b><%=s2Secs%></b> 秒</div>
            <div class="metric"><span>生成速度</span><b><%=genSpd%></b> 詞/秒</div>
            <div class="metric"><span>讀題速度</span><b><%=readSpd%></b> 詞/秒</div>
            <div class="metric"><span>載入模型</span><b><%=loadSecs%></b> 秒</div>
          </div>
          <% If s2Answer <> "" Then %><div class="code-box" style="font-family:inherit;font-size:14px"><%=H(s2Answer)%></div><% End If %>
          <details style="margin-top:10px"><summary class="note" style="cursor:pointer">查看原始回應</summary><div class="code-box"><%=H(Left(s2Raw, 1500))%></div></details>
          <% Else %>
          <div class="banner err" style="margin-top:10px"><svg width="16" height="16"><use href="#i-alert"/></svg><div><%=s2Msg%></div></div>
          <% End If %>
        <% End If %>
      </div>
    </div>
  </section>

  <aside class="aside">
    <section class="card">
      <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-gear"/></svg></div><div><h2>目前設定</h2><p>在 config.asp 修改</p></div></div>
      <div class="card-body">
        <table class="kv-table">
          <tr><th>AI 開關</th><td><% If SC_AI_ENABLED Then %><span class="badge success">啟用</span><% Else %><span class="badge neutral">關閉</span><% End If %> <span class="note"><%=H(SC_AI_PROVIDER)%></span></td></tr>
          <tr><th>網址</th><td><code><%=H(SC_AI_URL)%></code></td></tr>
          <tr><th>模型</th><td><code><%=H(SC_MODEL)%></code></td></tr>
          <tr><th>逾時 / 執行緒</th><td><%=SC_AI_TIMEOUT%> 秒 ／ <%=SC_AI_THREADS%></td></tr>
          <tr><th>每題原文段數</th><td><%=SC_TOPK%> 段</td></tr>
        </table>
      </div>
    </section>
    <section class="card help">
      <div class="card-ico teal"><svg width="18" height="18"><use href="#i-info"/></svg></div>
      <div><b>速度判斷</b><p>每題約讀 2,500 詞、寫 300～500 詞。生成 ≥ 10 詞/秒、讀題 ≥ 80 詞/秒，一題約 1 分鐘內。</p></div>
    </section>
  </aside>
</div>
<% SC_Bottom %>
