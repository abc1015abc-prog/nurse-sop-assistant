<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="S_coach_conn.asp"-->
<!--#include file="S_coach_layout.asp"-->
<%
Response.Charset = "utf-8"
Server.ScriptTimeout = 300
SC_CheckAdmin

Dim fso, dbPath, msg, conn, cat, doneOK
doneOK = False
Set fso = Server.CreateObject("Scripting.FileSystemObject")
dbPath = Server.MapPath(SC_DB)
msg = ""

Sub AddMsg(s)
  msg = msg & "<li><svg width='16' height='16'><use href='#i-check'/></svg><span>" & s & "</span></li>"
End Sub

Function TableExists(c, t)
  Dim rs
  Set rs = c.OpenSchema(20, Array(Empty, Empty, t, "TABLE"))
  TableExists = Not rs.EOF
  rs.Close
End Function

If Request.Form("go") = "1" Then
  '--- 1. 建立資料庫 ---
  If Not fso.FileExists(dbPath) Then
    Set cat = Server.CreateObject("ADOX.Catalog")
    cat.Create "Provider=" & SC_PROVIDER & ";Data Source=" & dbPath
    Set cat = Nothing
    AddMsg "已建立資料庫 " & H(SC_DB)
  Else
    AddMsg "資料庫已存在：" & H(SC_DB)
  End If
  If Not fso.FolderExists(Server.MapPath(SC_TMP)) Then fso.CreateFolder Server.MapPath(SC_TMP): AddMsg "已建立暫存資料夾 " & H(SC_TMP)

  Set conn = SC_Open()

  '--- 2. 建立資料表 ---
  If Not TableExists(conn, "std_list") Then
    conn.Execute "CREATE TABLE std_list (code TEXT(30) CONSTRAINT pk_std PRIMARY KEY, num TEXT(10), title TEXT(255), cat TEXT(100), has_eval TEXT(4), has_video TEXT(4), video_url TEXT(255), eff_date TEXT(20), rev_date TEXT(20), unit TEXT(50), editor TEXT(100), pdf_file TEXT(255), pdf_manual TEXT(255), pdf_mtime DATETIME, pages LONG, chars LONG, built_at DATETIME, status TEXT(100))"
    AddMsg "建立資料表 std_list（標準清單）"
  End If
  If Not TableExists(conn, "chunk") Then
    conn.Execute "CREATE TABLE chunk (id COUNTER CONSTRAINT pk_chunk PRIMARY KEY, code TEXT(30), title TEXT(255), cat TEXT(100), page LONG, seq LONG, content MEMO, ctext MEMO)"
    conn.Execute "CREATE INDEX ix_chunk_code ON chunk (code)"
    AddMsg "建立資料表 chunk（PDF 全文段落）"
  End If
  If Not TableExists(conn, "syn") Then
    conn.Execute "CREATE TABLE syn (id COUNTER CONSTRAINT pk_syn PRIMARY KEY, word TEXT(50), alias MEMO, src TEXT(20))"
    Dim seeds, sd, p
    seeds = Array("抽痰|suction", "鼻胃管|ng,n-g,nasogastric", "中心靜脈導管|cvc,cvp", "周邊置入中心靜脈導管|picc", _
                  "心肺復甦術|cpr,急救", "氣管內管|endo,ett,氣管插管", "導尿|foley,尿管", "胰島素|insulin", _
                  "昏迷指數|gcs", "灌食|feeding,管灌", "氣切|tracheostomy,trach", "血糖|bs,glucose", _
                  "輸血|blood transfusion", "壓傷|壓瘡,褥瘡", "跌倒|fall", "約束|保護性約束", "化療|化學治療", _
                  "靜脈注射|iv,點滴", "肌肉注射|im", "皮下注射|sc", "傷口|wound", "管路|tube,導管")
    For Each sd In seeds
      p = Split(sd, "|")
      conn.Execute "INSERT INTO syn (word, alias, src) VALUES ('" & SQ(p(0)) & "','" & SQ(p(1)) & "','預設')"
    Next
    AddMsg "建立資料表 syn（同義詞），預設 " & (UBound(seeds) + 1) & " 組，可到「同義詞管理」匯入臨床縮寫詞庫"
  End If
  If Not TableExists(conn, "qlog") Then
    conn.Execute "CREATE TABLE qlog (id COUNTER CONSTRAINT pk_qlog PRIMARY KEY, qtime DATETIME, worknb TEXT(20), unit TEXT(30), qmode TEXT(10), question TEXT(255), top_codes TEXT(255), cover LONG, result TEXT(20))"
    AddMsg "建立資料表 qlog（提問紀錄）"
  End If

  '--- 3. 匯入 / 更新標準清單 ---
  Dim txt, lines, ln, f, n, rs, isNew
  Dim listPath: listPath = Server.MapPath("S_coach_list.txt")
  If Not fso.FileExists(listPath) Then
    AddMsg "未找到 S_coach_list.txt（" & H(listPath) & "），略過「系統分類」匯入；標準清單會在 S_coach_build.asp 按①時從 amend_work.mdb 同步。"
  Else
    txt = SC_ReadUtf8(listPath)
    If Left(txt, 1) = ChrW(&HFEFF) Then txt = Mid(txt, 2)
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    n = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    For Each ln In lines
      If Trim(ln) <> "" Then
        f = Split(ln, vbTab)
        If UBound(f) >= 10 Then
          rs.Open "SELECT * FROM std_list WHERE code='" & SQ(f(0)) & "'", conn, 1, 3
          If rs.EOF Then rs.AddNew: rs("code") = f(0)
          rs("num") = f(1): rs("title") = f(2): rs("cat") = f(3)
          rs("has_eval") = f(4): rs("has_video") = f(5): rs("video_url") = f(6)
          rs("eff_date") = f(7): rs("rev_date") = f(8): rs("unit") = f(9): rs("editor") = f(10)
          rs.Update: rs.Close
          n = n + 1
        End If
      End If
    Next
    AddMsg "匯入 / 更新標準清單 " & n & " 筆（已對應的 PDF 與索引不會被清除）"
  End If
  conn.Close
  doneOK = True
End If
%>
<% SC_Top "setup", "系統初始化" %>
<nav class="crumbs">管理<svg width="14" height="14"><use href="#i-right"/></svg><span>系統初始化</span></nav>
<div class="page-head">
  <div>
    <div class="page-title"><h1>系統初始化</h1><% If fso.FileExists(dbPath) Then %><span class="badge success">資料庫已建立</span><% Else %><span class="badge warning">尚未建立</span><% End If %></div>
    <p class="page-sub">第一次安裝時執行。會建立本系統資料庫與資料表，並匯入 <code>S_coach_list.txt</code> 的系統分類；重複執行是安全的，只會補上缺少的部分。</p>
  </div>
</div>
<div class="two-col">
  <section class="card">
    <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-gear"/></svg></div><div><h2>執行初始化</h2><p>完成後到「索引管理」依序執行 ① 同步、② 建立索引。</p></div></div>
    <div class="card-body">
      <table class="kv-table">
        <tr><th>本系統資料庫</th><td><code><%=H(SC_DB)%></code></td></tr>
        <tr><th>標準資料庫</th><td><code><%=H(SC_CATALOG)%></code></td></tr>
        <tr><th>PDF 資料夾</th><td><code><%=H(SC_PDF_ROOT)%></code></td></tr>
        <tr><th>pdftotext</th><td><code><%=H(SC_PDFTOTEXT)%></code> <% If fso.FileExists(SC_PDFTOTEXT) Then %><span class="badge success">可用</span><% Else %><span class="badge danger">找不到</span><% End If %></td></tr>
      </table>
      <form method="post" style="margin-top:18px" onsubmit="var b=this.querySelector('button');b.classList.add('loading');b.disabled=true;">
        <input type="hidden" name="go" value="1">
        <button class="btn btn-primary btn-lg"><span class="spin"></span><svg width="18" height="18" class="lbl-ico"><use href="#i-bolt"/></svg>執行初始化</button>
      </form>
      <% If msg <> "" Then %>
      <div class="sec-title" style="margin-top:22px">執行結果</div>
      <ul class="guard" style="gap:10px"><%=msg%></ul>
      <% If doneOK Then %><div class="banner ok" style="margin-top:16px"><svg width="16" height="16"><use href="#i-check"/></svg><div>初始化完成。下一步：<a href="S_coach_build.asp"><b>前往索引管理</b></a>，依序執行 ① 同步、② 建立索引。</div></div><% End If %>
      <% End If %>
    </div>
  </section>
  <aside class="aside">
    <section class="card help">
      <div class="card-ico teal"><svg width="18" height="18"><use href="#i-info"/></svg></div>
      <div><b>建立的資料表</b><p>std_list 標準清單、chunk 全文段落、syn 同義詞、qlog 提問紀錄。同義詞可直接在 Access 增修。</p></div>
    </section>
  </aside>
</div>
<% SC_Bottom %>
