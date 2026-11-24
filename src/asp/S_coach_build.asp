<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="S_coach_conn.asp"-->
<!--#include file="S_coach_layout.asp"-->
<!--#include file="S_coach_types.asp"-->
<%
Response.Charset = "utf-8"
Server.ScriptTimeout = 900
SC_CheckAdmin

Const BATCH = 8   ' 每次處理幾份 PDF，處理完自動接續下一批（避免逾時）

Dim conn, fso, act, msg
Set conn = SC_Open()
SC_EnsureStdSrc conn
Set fso = Server.CreateObject("Scripting.FileSystemObject")
act = Request("act")
msg = ""

'==================== PDF 檔案清單 ====================
Dim fPath(), fName(), fTime(), fTop(), fRel(), fCnt
fCnt = 0
ReDim fPath(3000): ReDim fName(3000): ReDim fTime(3000): ReDim fTop(3000): ReDim fRel(3000)

' 遞迴掃描；記下每個檔案所在的第一層資料夾（CP、CPT、CR…＝標準類型）
Sub ScanFolder(fld, rel, top)
  Dim f, sf, r, n
  For Each f In fld.Files
    If LCase(fso.GetExtensionName(f.Name)) = "pdf" And Left(f.Name, 2) <> "~$" Then
      If fCnt > UBound(fPath) Then
        n = fCnt + 1000
        ReDim Preserve fPath(n): ReDim Preserve fName(n): ReDim Preserve fTime(n): ReDim Preserve fTop(n): ReDim Preserve fRel(n)
      End If
      If rel = "" Then r = f.Name Else r = rel & "\" & f.Name
      fPath(fCnt) = r
      fName(fCnt) = UCase(fso.GetBaseName(f.Name))
      fRel(fCnt) = UCase(Left(r, Len(r) - 4))
      fTop(fCnt) = UCase(top)
      fTime(fCnt) = f.DateLastModified
      fCnt = fCnt + 1
    End If
  Next
  For Each sf In fld.SubFolders
    If rel = "" Then ScanFolder sf, sf.Name, sf.Name Else ScanFolder sf, rel & "\" & sf.Name, top
  Next
End Sub

' 拆編號：CP062-1 → letters=CP, mainNo=62, suf=[-_ ]1
Function CodeParts(code, ByRef letters, ByRef mainNo, ByRef suf)
  Dim re, m
  Set re = New RegExp
  re.Pattern = "^([A-Z]+)0*([0-9]+)(.*)$"
  Set m = re.Execute(UCase(code))
  CodeParts = (m.Count > 0)
  If m.Count = 0 Then Exit Function
  letters = m(0).SubMatches(0): mainNo = m(0).SubMatches(1): suf = m(0).SubMatches(2)
  suf = Replace(suf, ".", "\.")
  suf = Replace(suf, "-", "[-_ ]")
End Function

Const NO_MORE = "(?![0-9]|[-_.][0-9])"   ' 後面不能再接數字或 -數字（CP062 不會配到 CP062-1、CP0621）

Function Pick(re, onlyTop, useName)
  Dim i, best, bestPen, pen, target
  best = -1: bestPen = 9
  For i = 0 To fCnt - 1
    If onlyTop = "" Or fTop(i) = onlyTop Then
      If useName Then target = fName(i) Else target = fRel(i)
      If re.Test(target) Then
        pen = Penalty(fName(i))
        If Better(i, pen, best, bestPen) Then best = i: bestPen = pen
      End If
    End If
  Next
  Pick = best
End Function

Function Penalty(nm)
  Dim re
  Set re = New RegExp: re.IgnoreCase = True
  re.Pattern = "評值|評估表|檢核|稽核|舊版|作廢|OLD|BAK|備份"
  If re.Test(nm) Then Penalty = 1 Else Penalty = 0
End Function

' 同一標準有多個候選檔：先選非評值表/舊版，再選最新修改者
Function Better(i, pen, best, bestPen)
  If best = -1 Then
    Better = True
  ElseIf pen < bestPen Then
    Better = True
  ElseIf pen = bestPen Then
    Better = (fTime(i) > fTime(best))
  Else
    Better = False
  End If
End Function

' 為一個標準找最適合的 PDF；回傳索引或 -1
Function FindPdf(code, title, num, numUnique)
  Dim re, i, best, pen, bestPen, t, letters, mainNo, suf
  FindPdf = -1
  If Not CodeParts(code, letters, mainNo, suf) Then Exit Function
  Set re = New RegExp: re.IgnoreCase = True

  ' 1. 在同類型資料夾（例 CP001 → CP\）中，路徑含完整編號：CP\CP001抽痰術.pdf、CP\CP001\xxx.pdf
  re.Pattern = "(^|[^A-Z])" & letters & "[-_ ]?0*" & mainNo & suf & NO_MORE
  best = Pick(re, letters, False)
  ' 2. 在同類型資料夾中，檔名以號碼開頭：CP\001.pdf、CP\001抽痰術.pdf、CP\062-1.pdf
  If best = -1 Then
    re.Pattern = "^(" & letters & ")?[-_ ]?0*" & mainNo & suf & NO_MORE
    best = Pick(re, letters, True)
  End If
  ' 3. 其他資料夾中，路徑含完整編號
  If best = -1 Then
    re.Pattern = "(^|[^A-Z])" & letters & "[-_ ]?0*" & mainNo & suf & NO_MORE
    best = Pick(re, "", False)
  End If
  ' 4. 檔名包含完整標準名稱（先找同類型資料夾）
  If best = -1 And Len(title) >= 3 Then
    t = SC_Norm(title): bestPen = 9
    For i = 0 To fCnt - 1
      If InStr(SC_Norm(fName(i)), t) > 0 Then
        pen = Penalty(fName(i))
        If fTop(i) <> letters Then pen = pen + 2
        If Better(i, pen, best, bestPen) Then best = i: bestPen = pen
      End If
    Next
  End If
  ' 5. 檔名就是【流水號】（僅限不重複的流水號）
  If best = -1 And numUnique And num <> "" Then
    For i = 0 To fCnt - 1
      If fName(i) = num Then best = i: Exit For
    Next
  End If
  FindPdf = best
End Function

'==================== PDF → 文字 → 段落 ====================
Dim reMarker, rePageNo
Set reMarker = New RegExp
reMarker.Pattern = "^([\(（]?[0-9０-９]{1,2}[\.\)、）．]|[\(（][一二三四五六七八九十0-9]+[\)）]|[一二三四五六七八九十]+[、．.]|[壹貳參肆伍陸柒捌玖拾]+[、．.]|[•●○■□◆◇▲►※★☆\-‧．])"
Set rePageNo = New RegExp
rePageNo.Pattern = "^(第\s*[0-9]+\s*頁.*|[0-9]+\s*/\s*[0-9]+|-?\s*[0-9]+\s*-?|page\s*[0-9]+.*|[0-9]+)$"
rePageNo.IgnoreCase = True

Function IsAlnum(ch)
  Dim c
  IsAlnum = False
  If ch = "" Then Exit Function
  c = AscW(ch)
  IsAlnum = (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

' 把一頁文字的折行接回成「段落/條列項目」陣列
Function PageItems(pg, rep)
  Dim lines, ln, t, cur, items, n, blank
  ReDim items(200): n = 0
  cur = "": blank = False
  lines = Split(Replace(pg, vbCr, ""), vbLf)
  For Each ln In lines
    t = Trim(Replace(Replace(ln, ChrW(&H3000), " "), vbTab, " "))
    If t = "" Then
      blank = True
    ElseIf rep.Exists(SC_Norm(t)) Or rePageNo.Test(t) Then
      ' 每頁重複的頁首頁尾、頁碼 → 略過
    Else
      If cur = "" Then
        cur = t
      ElseIf blank Or reMarker.Test(t) Then
        If n > UBound(items) Then ReDim Preserve items(n + 200)
        items(n) = cur: n = n + 1: cur = t
      ElseIf IsAlnum(Right(cur, 1)) And IsAlnum(Left(t, 1)) Then
        cur = cur & " " & t
      Else
        cur = cur & t
      End If
      blank = False
    End If
  Next
  If cur <> "" Then
    If n > UBound(items) Then ReDim Preserve items(n + 1)
    items(n) = cur: n = n + 1
  End If
  If n = 0 Then PageItems = Array() Else ReDim Preserve items(n - 1): PageItems = items
End Function

Function BuildOne(code, title, cat, pdfPath)
  Dim tmp, inF, outF, sh, rc, txt, pages, i, rep, cnt, k, ln, seen, items, it, acc, rs, seq, nChars, nPages, chunks(), cn, j
  tmp = Server.MapPath(SC_TMP)
  If Not fso.FolderExists(tmp) Then fso.CreateFolder tmp
  inF = tmp & "\sc_in.pdf": outF = tmp & "\sc_out.txt"
  If fso.FileExists(outF) Then fso.DeleteFile outF, True
  If Not fso.FileExists(pdfPath) Then BuildOne = "找不到檔案": Exit Function
  fso.CopyFile pdfPath, inF, True     ' 複製成英文檔名，避免中文路徑造成 pdftotext 失敗

  Set sh = Server.CreateObject("WScript.Shell")
  rc = sh.Run("""" & SC_PDFTOTEXT & """ -q -enc UTF-8 -eol unix """ & inF & """ """ & outF & """", 0, True)
  If Not fso.FileExists(outF) Then BuildOne = "轉檔失敗(代碼" & rc & ")": Exit Function

  txt = SC_ReadUtf8(outF)
  pages = Split(txt, Chr(12))
  nPages = UBound(pages) + 1
  If nPages > 1 And Trim(pages(nPages - 1)) = "" Then nPages = nPages - 1

  ' 找出出現在 60% 以上頁面的行（頁首、頁尾、機構名稱）
  Set cnt = Server.CreateObject("Scripting.Dictionary")
  Set rep = Server.CreateObject("Scripting.Dictionary")
  If nPages >= 3 Then
    For i = 0 To nPages - 1
      Set seen = Server.CreateObject("Scripting.Dictionary")
      For Each ln In Split(Replace(pages(i), vbCr, ""), vbLf)
        k = SC_Norm(ln)
        If k <> "" And Len(k) <= 60 And Not seen.Exists(k) Then
          seen.Add k, 1
          If cnt.Exists(k) Then cnt(k) = cnt(k) + 1 Else cnt.Add k, 1
        End If
      Next
    Next
    For Each k In cnt.Keys
      If cnt(k) >= nPages * 0.6 Then rep.Add k, 1
    Next
  End If

  conn.Execute "DELETE FROM chunk WHERE code='" & SQ(code) & "'"
  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open "SELECT * FROM chunk WHERE 1=0", conn, 1, 3
  seq = 0: nChars = 0
  For i = 0 To nPages - 1
    items = PageItems(pages(i), rep)
    ReDim chunks(50): cn = 0: acc = ""
    For Each it In items
      If acc <> "" And Len(acc) + Len(it) > SC_CHUNK_LEN Then
        If cn > UBound(chunks) Then ReDim Preserve chunks(cn + 50)
        chunks(cn) = acc: cn = cn + 1: acc = ""
      End If
      ' 單一段落過長 → 硬切
      Do While Len(it) > SC_CHUNK_LEN * 1.5
        If acc <> "" Then
          If cn > UBound(chunks) Then ReDim Preserve chunks(cn + 50)
          chunks(cn) = acc: cn = cn + 1: acc = ""
        End If
        If cn > UBound(chunks) Then ReDim Preserve chunks(cn + 50)
        chunks(cn) = Left(it, SC_CHUNK_LEN): cn = cn + 1
        it = Mid(it, SC_CHUNK_LEN + 1)
      Loop
      If acc = "" Then acc = it Else acc = acc & vbLf & it
    Next
    If acc <> "" Then
      If cn > 0 And Len(acc) < 40 Then
        chunks(cn - 1) = chunks(cn - 1) & vbLf & acc      ' 頁尾零碎文字併入前段
      Else
        If cn > UBound(chunks) Then ReDim Preserve chunks(cn + 1)
        chunks(cn) = acc: cn = cn + 1
      End If
    End If
    For j = 0 To cn - 1
      seq = seq + 1
      rs.AddNew
      rs("code") = code: rs("title") = title: rs("cat") = cat
      rs("page") = i + 1: rs("seq") = seq
      rs("content") = SC_CleanText(chunks(j))
      rs("ctext") = SC_Norm(SC_CleanText(chunks(j)))
      rs.Update
      nChars = nChars + Len(chunks(j))
    Next
  Next
  rs.Close

  conn.Execute "UPDATE std_list SET pages=" & nPages & ", chars=" & nChars & ", built_at=Now(), status='" & _
               SQ(IIf(nChars < 50, "無文字(可能為掃描檔)", "已建索引")) & "' WHERE code='" & SQ(code) & "'"
  If fso.FileExists(outF) Then fso.DeleteFile outF, True
  If fso.FileExists(inF) Then fso.DeleteFile inF, True
  BuildOne = "OK " & nPages & "頁/" & seq & "段"
End Function

Function IIf(c, a, b)
  If c Then IIf = a Else IIf = b
End Function

Sub SyncFail(what, detail)
  SC_Top "build", "索引管理"
  Response.Write "<nav class='crumbs'>管理<svg width='14' height='14'><use href='#i-right'/></svg><span>索引管理</span></nav>" & _
    "<div class='page-head'><div><div class='page-title'><h1>同步失敗</h1><span class='badge danger'>" & H(what) & "</span></div></div></div>" & _
    "<section class='card'><div class='card-body'><table class='kv-table'><tr><th>資料庫</th><td><code>" & H(SC_CatalogPath()) & "</code></td></tr>" & _
    "<tr><th>錯誤訊息</th><td><b>" & H(detail) & "</b></td></tr><tr><th>常見原因</th><td>IIS 帳號對該資料夾沒有讀取權限；catalog 是「連結資料表」而來源檔不在；資料庫正被 Access 以獨佔方式開啟。</td></tr></table>" & _
    "<div style='margin-top:16px'><a class='btn btn-secondary' href='S_coach_build.asp'>返回索引管理</a></div></div></section>"
  SC_Bottom
  Response.End
End Sub

'==================== 動作 ====================
Dim rs, i, idx, numCount, startAt, mode, list, done, path, needBuild, built, res, nOk
Dim isBuilding, progTot, progCur, progDone, progNext
isBuilding = False

If act = "sync" Then
  ' 直接讀護理標準系統的 amend_work.mdb → catalog（全部標準以 catalog 為準；hidden_YN=YES 為停用，不納入）
  Dim cconn, crs, seen, code, fn, rel, oldRel, v, nNew, nUpd, nGone, nNoFile, nNoPdf, gone
  nNew = 0: nUpd = 0: nGone = 0: nNoFile = 0: nNoPdf = 0
  Set cconn = Server.CreateObject("ADODB.Connection")
  If Not fso.FileExists(SC_CatalogPath()) Then SyncFail "找不到資料庫檔案", "請確認 SC_CATALOG 路徑"
  On Error Resume Next
  cconn.Mode = 1                  ' 唯讀開啟，不鎖定、不寫入標準資料庫
  cconn.CursorLocation = 3        ' 用戶端游標：整列先讀進來，備忘欄位不會讀到空白
  cconn.Open "Provider=" & SC_PROVIDER & ";Data Source=" & SC_CatalogPath()
  If Err.Number <> 0 Then SyncFail "開啟資料庫失敗", Err.Description
  Dim errMsg
  ' catalog 是 SQL 保留字 → 加中括號；仍失敗就改用「直接開資料表」，再不行改用一般模式重開
  Set crs = cconn.Execute("SELECT * FROM [catalog]")
  If Err.Number <> 0 Then
    errMsg = "SQL：" & Hex(Err.Number) & " " & Err.Description
    Err.Clear
    Set crs = Server.CreateObject("ADODB.Recordset")
    crs.Open "catalog", cconn, 0, 1, 2          ' adOpenForwardOnly, adLockReadOnly, adCmdTable
    If Err.Number <> 0 Then
      errMsg = errMsg & "／資料表：" & Hex(Err.Number) & " " & Err.Description
      Err.Clear
      cconn.Close
      Set cconn = Server.CreateObject("ADODB.Connection")
      cconn.CursorLocation = 3
      cconn.Open "Provider=" & SC_PROVIDER & ";Data Source=" & SC_CatalogPath()
      Set crs = cconn.Execute("SELECT * FROM [catalog]")
      If Err.Number <> 0 Then SyncFail "讀取 catalog 資料表失敗", errMsg & "／一般模式：" & Hex(Err.Number) & " " & Err.Description
    End If
  End If
  On Error GoTo 0
  Set seen = Server.CreateObject("Scripting.Dictionary")
  Set rs = Server.CreateObject("ADODB.Recordset")
  Dim fv, fi, cPress, cSDNB, cTitle, cAuth, cAmend, cEditor, cFolder, cFile, cFile1, cFile2, cEform, cHidden, sampleFn, cand, cnd
  sampleFn = ""
  Do Until crs.EOF
    ' 依欄位順序一次讀完（避免備忘欄位跳讀後變空白）
    Set fv = Server.CreateObject("Scripting.Dictionary")
    For fi = 0 To crs.Fields.Count - 1
      fv(LCase(crs.Fields(fi).Name)) = crs.Fields(fi).Value & ""
    Next
    cPress = fv("pressmark"): cSDNB = fv("sdnb"): cTitle = fv("denominate")
    cAuth = fv("authenticate"): cAmend = fv("amend_date"): cEditor = fv("castigator_unit")
    cFolder = Trim(fv("fileway")): If cFolder = "" Then cFolder = Trim(fv("castigator"))   ' fileway＝資料夾名稱（例 CP）
    cFolder = Replace(Replace(cFolder, "/", "\"), "paper\", "")
    If Right(cFolder, 1) = "\" Then cFolder = Left(cFolder, Len(cFolder) - 1)
    cFile = fv("filename"): cFile1 = fv("filename1"): cFile2 = fv("filename2"): cEform = fv("eform"): cHidden = fv("hidden_yn")
    code = UCase(Trim(cPress))
    If UCase(Trim(cHidden)) = "YES" Then code = ""   ' hidden_YN=YES 為停用
    If code <> "" And Not seen.Exists(code) Then
      seen.Add code, 1
      rs.Open "SELECT * FROM std_list WHERE code='" & SQ(code) & "'", conn, 1, 3
      If rs.EOF Then
        rs.AddNew: rs("code") = code: nNew = nNew + 1
      Else
        nUpd = nUpd + 1
      End If
      rs("num") = Left(cSDNB, 10)
      rs("title") = Left(cTitle, 255)
      rs("eff_date") = Left(cAuth, 20)
      rs("rev_date") = Left(cAmend, 20)
      rs("editor") = Left(cEditor, 100)
      v = Trim(cEform)
      If LCase(Left(v, 4)) = "http" Then
        rs("video_url") = Left(v, 255): rs("has_video") = "有"
      Else
        rs("video_url") = Null: rs("has_video") = "無"
      End If
      ' PDF＝paper\fileway\filename1（filename 是 Word）；filename1 空白或找不到時，再看 filename2/filename 中副檔名為 .pdf 的
      fn = "": rel = ""
      For Each cand In Array(Trim(cFile1), Trim(cFile2), Trim(cFile))
        If rel = "" And LCase(Right(cand, 4)) = ".pdf" Then
          cnd = Replace(cand, "/", "\")
          If InStr(cnd, "\") = 0 Then cnd = Trim(cFolder) & "\" & cnd       ' 只有檔名 → 類型資料夾\檔名
          If fn = "" Then fn = cnd                                           ' 記下第一個候選
          If InStr(cnd, ":\") > 0 Then
            If fso.FileExists(cnd) Then rel = cnd
          ElseIf fso.FileExists(SC_PDF_ROOT & "\" & cnd) Then
            rel = cnd
          End If
        End If
      Next
      If rel = "" Then rel = fn
      If Len(sampleFn) < 300 Then sampleFn = sampleFn & H(code) & "：fileway「" & H(cFolder) & "」filename1「" & H(cFile1) & "」filename「" & H(cFile) & "」　"
      If rel <> "" Then
        oldRel = rs("pdf_file") & ""
        rs("pdf_file") = rel
        If LCase(oldRel) <> LCase(rel) Then rs("built_at") = Null   ' 換了檔案 → 需重建
        If InStr(rel, ":\") > 0 Then cnd = rel Else cnd = SC_PDF_ROOT & "\" & rel
        If Not fso.FileExists(cnd) And Trim(rs("pdf_manual") & "") = "" Then
          rs("status") = "找不到檔案": nNoFile = nNoFile + 1
        ElseIf rs("status") & "" <> "已建索引" Or IsNull(rs("built_at")) Then
          rs("status") = "已對應(未建索引)"
        End If
      Else
        rs("pdf_file") = Null
        If Trim(rs("pdf_manual") & "") = "" Then rs("status") = "無PDF(僅影片/簡報)": nNoPdf = nNoPdf + 1
      End If
      rs.Update: rs.Close
    End If
    crs.MoveNext
  Loop
  crs.Close: cconn.Close
  ' 標準系統已刪除或隱藏的 → 從本系統移除（次專科標準另外同步，不在這裡處理）
  Set crs = conn.Execute("SELECT code FROM std_list WHERE src Is Null OR src<>'SUS'")
  gone = ""
  Do Until crs.EOF
    If Not seen.Exists(UCase(crs("code") & "")) Then gone = gone & crs("code") & "|"
    crs.MoveNext
  Loop
  crs.Close
  For Each code In Split(gone, "|")
    If code <> "" Then
      conn.Execute "DELETE FROM chunk WHERE code='" & SQ(code) & "'"
      conn.Execute "DELETE FROM std_list WHERE code='" & SQ(code) & "'"
      nGone = nGone + 1
    End If
  Next
  msg = "<div class='banner ok'>已從 amend_work.mdb 同步：使用中標準 " & seen.Count & " 項（新增 " & nNew & "、更新 " & nUpd & "、移除停用 " & nGone & "）。" & _
        "無 PDF（僅影片/簡報）" & nNoPdf & " 項；資料庫有檔名但資料夾找不到檔案 " & nNoFile & " 項。</div>"
  If nNoPdf > seen.Count / 2 Then msg = msg & "<div class='banner err'>大部分標準讀不到 PDF 檔名，前幾筆讀到的 filename：" & sampleFn & "</div>"
  If seen.Count = 0 Then msg = "<div class='banner err'>catalog 沒有讀到資料，請確認 SC_CATALOG 路徑。</div>"

ElseIf act = "syncsus" Then
  ' 病房單位次專科標準：讀 D_sus.mdb 的 eodata（hidden_YN=YES 為停用）；PDF 檔名＝file_1
  Dim sconn, srs, sSeen, sCode, sFile, sPath, sDict, sScanned, sNew, sUpd, sGone, sNoFile, sNoPdf, sMiss, sk, sGoneList, sUnit
  sNew = 0: sUpd = 0: sGone = 0: sNoFile = 0: sNoPdf = 0: sMiss = "": sScanned = False
  If Trim(SC_SUS_DB) = "" Then SyncFail "未設定次專科資料庫", "請在 config.asp 設定 SC_SUS_DB"
  If Not fso.FileExists(SC_SUS_DB) Then SyncFail "找不到次專科資料庫", "請確認 config.asp 的 SC_SUS_DB：" & SC_SUS_DB
  Set sconn = Server.CreateObject("ADODB.Connection")
  sconn.Mode = 1: sconn.CursorLocation = 3
  sconn.Open "Provider=" & SC_PROVIDER & ";Data Source=" & SC_SUS_DB
  Set srs = sconn.Execute("SELECT punit, SNB, SNB_1, Sclass, Sname, file_1, A_day, U_day, system_assort, gerent, hidden_YN FROM eodata ORDER BY hidden_YN, ID")
  Set sSeen = Server.CreateObject("Scripting.Dictionary")
  Set sDict = Server.CreateObject("Scripting.Dictionary")
  Set rs = Server.CreateObject("ADODB.Recordset")
  Do Until srs.EOF
    sCode = UCase(Trim(srs("SNB") & ""))
    sUnit = Trim(srs("punit") & "")
    If UCase(Trim(srs("hidden_YN") & "")) <> "YES" And sCode <> "" And SC_SusType(sCode) <> "" And Not sSeen.Exists(sCode) Then
      sSeen.Add sCode, 1
      ' 找 PDF：資料夾\單位\檔名 → 資料夾\檔名 → 往子資料夾搜尋同檔名
      sFile = Trim(srs("file_1") & ""): sPath = ""
      If LCase(Right(sFile, 4)) = ".pdf" Then
        If fso.FileExists(SC_SUS_PDF_ROOT & "\" & sUnit & "\" & sFile) Then
          sPath = SC_SUS_PDF_ROOT & "\" & sUnit & "\" & sFile
        ElseIf fso.FileExists(SC_SUS_PDF_ROOT & "\" & sFile) Then
          sPath = SC_SUS_PDF_ROOT & "\" & sFile
        Else
          If Not sScanned Then
            sScanned = True
            If fso.FolderExists(SC_SUS_PDF_ROOT) Then
              fCnt = 0
              ScanFolder fso.GetFolder(SC_SUS_PDF_ROOT), "", ""
              For sk = 0 To fCnt - 1
                ' Revision（修訂中／舊版）資料夾不採用，避免抓到非現行版本
                If InStr("\" & UCase(fPath(sk)), "\REVISION\") = 0 Then
                  If Not sDict.Exists(LCase(fso.GetFileName(fPath(sk)))) Then sDict.Add LCase(fso.GetFileName(fPath(sk))), SC_SUS_PDF_ROOT & "\" & fPath(sk)
                End If
              Next
            End If
          End If
          If sDict.Exists(LCase(sFile)) Then sPath = sDict(LCase(sFile))
        End If
      End If
      rs.Open "SELECT * FROM std_list WHERE code='" & SQ(sCode) & "'", conn, 1, 3
      If rs.EOF Then
        rs.AddNew: rs("code") = sCode: sNew = sNew + 1
      Else
        sUpd = sUpd + 1
      End If
      rs("src") = "SUS"
      rs("num") = Left(srs("SNB_1") & "", 10)
      rs("title") = Left(srs("Sname") & "", 255)
      rs("cat") = Left(srs("system_assort") & "", 100)
      rs("eff_date") = Left(srs("A_day") & "", 20)
      rs("rev_date") = Left(srs("U_day") & "", 20)
      rs("unit") = Left(sUnit, 50)
      rs("editor") = Left(srs("gerent") & "", 100)
      rs("video_url") = Null: rs("has_video") = "無"
      If sPath <> "" Then
        If LCase(rs("pdf_file") & "") <> LCase(sPath) Then rs("built_at") = Null
        rs("pdf_file") = sPath
        If rs("status") & "" <> "已建索引" Or IsNull(rs("built_at")) Then rs("status") = "已對應(未建索引)"
      ElseIf sFile <> "" Then
        rs("pdf_file") = Null
        If Trim(rs("pdf_manual") & "") = "" Then
          rs("status") = "找不到檔案": sNoFile = sNoFile + 1
          If Len(sMiss) < 400 Then sMiss = sMiss & H(sCode) & "（" & H(sFile) & "）　"
        End If
      Else
        rs("pdf_file") = Null
        If Trim(rs("pdf_manual") & "") = "" Then rs("status") = "無PDF(僅影片/簡報)": sNoPdf = sNoPdf + 1
      End If
      rs.Update: rs.Close
    End If
    srs.MoveNext
  Loop
  srs.Close: sconn.Close
  ' 已停用或刪除的次專科標準 → 移除
  Set srs = conn.Execute("SELECT code FROM std_list WHERE src='SUS'")
  sGoneList = ""
  Do Until srs.EOF
    If Not sSeen.Exists(UCase(srs("code") & "")) Then sGoneList = sGoneList & srs("code") & "|"
    srs.MoveNext
  Loop
  srs.Close
  For Each sk In Split(sGoneList, "|")
    If sk <> "" Then
      conn.Execute "DELETE FROM chunk WHERE code='" & SQ(sk) & "'"
      conn.Execute "DELETE FROM std_list WHERE code='" & SQ(sk) & "'"
      sGone = sGone + 1
    End If
  Next
  msg = "<div class='banner ok'>已從 D_sus.mdb 同步次專科標準 " & sSeen.Count & " 項（新增 " & sNew & "、更新 " & sUpd & "、移除停用 " & sGone & "）。" & _
        "無 PDF " & sNoPdf & " 項；找不到 PDF 檔案 " & sNoFile & " 項。接著請按「建立索引」。</div>"
  If sNoFile > 0 Then msg = msg & "<div class='banner err'>找不到 PDF（已在 " & H(SC_SUS_PDF_ROOT) & " 及其子資料夾搜尋）：" & sMiss & _
        IIf(sNoFile > sSeen.Count / 2, "<br><b>大部分都找不到，請修改 config.asp 的 SC_SUS_PDF_ROOT 為次專科 PDF 所在資料夾。</b>", "") & "</div>"

ElseIf act = "scan" Then
  If Not fso.FolderExists(SC_PDF_ROOT) Then
    msg = "<div class='banner err'>找不到 PDF 資料夾：" & H(SC_PDF_ROOT) & "（請修改 config.asp 的 SC_PDF_ROOT）</div>"
  Else
    ScanFolder fso.GetFolder(SC_PDF_ROOT), "", ""
    Set numCount = Server.CreateObject("Scripting.Dictionary")
    Set rs = conn.Execute("SELECT num FROM std_list")
    Do Until rs.EOF
      If numCount.Exists(rs("num") & "") Then numCount(rs("num") & "") = 2 Else numCount.Add rs("num") & "", 1
      rs.MoveNext
    Loop
    rs.Close
    Dim used: Set used = Server.CreateObject("Scripting.Dictionary")
    Dim nMatch, nMiss: nMatch = 0: nMiss = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM std_list", conn, 1, 3
    Do Until rs.EOF
      If Trim(rs("pdf_manual") & "") = "" Then
        idx = FindPdf(rs("code") & "", rs("title") & "", rs("num") & "", numCount(rs("num") & "") = 1)
        If idx >= 0 Then
          rs("pdf_file") = fPath(idx): rs("pdf_mtime") = fTime(idx)
          If rs("status") & "" = "" Or rs("status") & "" = "找不到PDF" Then rs("status") = "已對應(未建索引)"
          used(LCase(fPath(idx))) = 1: nMatch = nMatch + 1
        Else
          rs("pdf_file") = Null: rs("status") = "找不到PDF": nMiss = nMiss + 1
        End If
        rs.Update
      Else
        used(LCase(rs("pdf_manual") & "")) = 1: nMatch = nMatch + 1
      End If
      rs.MoveNext
    Loop
    rs.Close
    msg = "<div class='banner ok'>掃描到 " & fCnt & " 個 PDF；已對應 " & nMatch & " 項標準，找不到 " & nMiss & " 項。</div>"
    Dim unusedList: unusedList = ""
    For i = 0 To fCnt - 1
      If Not used.Exists(LCase(fPath(i))) Then unusedList = unusedList & "<li>" & H(fPath(i)) & "</li>"
    Next
    If unusedList <> "" Then msg = msg & "<details><summary>未對應到任何標準的 PDF（可複製檔名到下方手動指定）</summary><ul>" & unusedList & "</ul></details>"
  End If

ElseIf act = "setfile" Then
  conn.Execute "UPDATE std_list SET pdf_manual=" & IIf(Trim(Request.Form("f")) = "", "Null", "'" & SQ(Trim(Request.Form("f"))) & "'") & _
               ", status='已對應(未建索引)' WHERE code='" & SQ(Request.Form("code")) & "'"
  msg = "<div class='banner ok'>已更新 " & H(Request.Form("code")) & " 的 PDF 檔案。</div>"

ElseIf act = "build" Then
  ' mode=all 全部重建；mode=changed 只建新的或 PDF 有更新的
  mode = Request("mode"): If mode <> "all" Then mode = "changed"
  startAt = CLng("0" & Request("start"))
  nOk = CLng("0" & Request("ok"))
  Set rs = conn.Execute("SELECT code, title, cat, built_at FROM std_list WHERE (pdf_file Is Not Null AND pdf_file<>'') OR (pdf_manual Is Not Null AND pdf_manual<>'') ORDER BY code")
  If rs.EOF Then
    rs.Close
    Response.Redirect "S_coach_build.asp"
  End If
  list = rs.GetRows(): rs.Close
  done = ""
  i = startAt: built = 0
  Do While i <= UBound(list, 2) And built < BATCH
    path = SC_PdfPath(conn, list(0, i))
    needBuild = True
    If mode = "changed" And Not IsNull(list(3, i)) And path <> "" Then
      If fso.FileExists(path) Then
        If fso.GetFile(path).DateLastModified <= list(3, i) Then needBuild = False
      End If
    End If
    If needBuild Then
      On Error Resume Next
      res = BuildOne(list(0, i) & "", list(1, i) & "", list(2, i) & "", path)
      If Err.Number <> 0 Then res = "錯誤：" & Err.Description: Err.Clear
      On Error GoTo 0
      If Left(res, 2) <> "OK" Then conn.Execute "UPDATE std_list SET status='" & SQ(Left(res, 100)) & "' WHERE code='" & SQ(list(0, i)) & "'" Else nOk = nOk + 1
      done = done & "<li>" & IIf(Left(res, 2) = "OK", "<span class='badge success'>完成</span>", "<span class='badge danger'>失敗</span>") & _
             "<span class='tag code'>" & H(list(0, i)) & "</span>" & H(list(1, i)) & "<span class='note'>" & H(res) & "</span></li>"
      built = built + 1
    End If
    i = i + 1
  Loop
  progTot = UBound(list, 2) + 1: progCur = i: progDone = done
  If i <= UBound(list, 2) Then
    isBuilding = True
    progNext = "S_coach_build.asp?act=build&mode=" & mode & "&start=" & i & "&ok=" & nOk
  Else
    msg = "<div class='banner ok'>索引建立完成，本次共處理 " & nOk & " 份。</div>"
  End If
End If

'==================== 狀態表 ====================
Dim stat, tot, nIdx, nChunk, nTodo, nNoPdf2, nErr, pdfOk, grp, catOk, cntJ

Function StGroup(st)
  If st = "已建索引" Then
    StGroup = "ok"
  ElseIf st = "已對應(未建索引)" Or st = "" Then
    StGroup = "todo"
  ElseIf Left(st, 5) = "無PDF(" Then
    StGroup = "nopdf"
  Else
    StGroup = "err"
  End If
End Function

Function StBadge(st)
  Select Case StGroup(st)
    Case "ok":    StBadge = "<span class='badge success'>已建索引</span>"
    Case "todo":  StBadge = "<span class='badge info'>待建索引</span>"
    Case "nopdf": StBadge = "<span class='badge neutral'>無 PDF</span>"
    Case Else
      If InStr(st, "掃描") > 0 Then StBadge = "<span class='badge warning'>掃描檔無文字</span>" Else StBadge = "<span class='badge danger'>" & H(st) & "</span>"
  End Select
End Function

tot = conn.Execute("SELECT COUNT(*) FROM std_list")(0)
nIdx = conn.Execute("SELECT COUNT(*) FROM std_list WHERE status='已建索引'")(0)
nTodo = conn.Execute("SELECT COUNT(*) FROM std_list WHERE status='已對應(未建索引)' OR status Is Null OR status=''")(0)
nNoPdf2 = conn.Execute("SELECT COUNT(*) FROM std_list WHERE status LIKE '無PDF%'")(0)
nErr = tot - nIdx - nTodo - nNoPdf2
nChunk = conn.Execute("SELECT COUNT(*) FROM chunk")(0)
pdfOk = fso.FileExists(SC_PDFTOTEXT)
catOk = fso.FileExists(SC_CatalogPath())

SC_Top "build", "索引管理"
%>
<% If isBuilding Then %><meta http-equiv="refresh" content="1;url=<%=progNext%>"><% End If %>
<nav class="crumbs">管理<svg width="14" height="14"><use href="#i-right"/></svg><span>索引管理</span></nav>
<div class="page-head">
  <div>
    <div class="page-title"><h1>索引管理</h1>
      <% If isBuilding Then %><span class="badge info pulse">建立索引中</span><% ElseIf nTodo > 0 Then %><span class="badge warning"><%=nTodo%> 份待建索引</span><% Else %><span class="badge success">索引已是最新</span><% End If %>
    </div>
    <p class="page-sub">從護理標準系統同步標準清單與 PDF 檔名，再把 PDF 全文轉成可搜尋的段落。標準改版後，依序按 ① 同步、② 建立索引即可。</p>
    <div class="meta-row">
      <span><svg width="16" height="16"><use href="#i-db"/></svg>標準資料庫 <% If catOk Then %><span class="badge success">已連線</span><% Else %><span class="badge danger">找不到</span><% End If %></span>
      <span><svg width="16" height="16"><use href="#i-folder"/></svg><code><%=H(SC_PDF_ROOT)%></code></span>
      <span><svg width="16" height="16"><use href="#i-cpu"/></svg>pdftotext <% If pdfOk Then %><span class="badge success">可用</span><% Else %><span class="badge danger">找不到</span><% End If %></span>
    </div>
  </div>
  <div class="head-actions">
    <a class="btn btn-secondary" href="S_coach_log.asp"><svg width="18" height="18"><use href="#i-list"/></svg>提問紀錄</a>
    <a class="btn btn-primary" href="S_coach.asp#chat"><svg width="18" height="18"><use href="#i-chat"/></svg>前往問助理</a>
  </div>
</div>

<% If msg <> "" Then %><div class="msg"><%=msg%></div><% End If %>

<% If isBuilding Then %>
<section class="card progress-card">
  <div class="top"><div class="card-ico"><svg width="20" height="20"><use href="#i-refresh"/></svg></div>
    <div><b>正在建立索引…</b><div class="note">每批處理 <%=BATCH%> 份，頁面會自動接續，請勿關閉</div></div>
    <span class="pct"><%=Int(progCur / progTot * 100)%>%</span></div>
  <div class="bar"><i style="width:<%=Int(progCur / progTot * 100)%>%"></i></div>
  <div class="note" style="margin-top:6px"><%=progCur%> / <%=progTot%> 份</div>
  <ul><%=progDone%></ul>
</section>
<% End If %>

<div class="stats">
  <section class="card stat"><div class="k"><i><svg width="16" height="16"><use href="#i-book"/></svg></i>標準總數</div><div class="v"><%=tot%><small>份</small></div><div class="d">依 catalog（hidden_YN≠YES）</div></section>
  <section class="card stat"><div class="k"><i class="ok"><svg width="16" height="16"><use href="#i-check"/></svg></i>已建索引</div><div class="v"><%=nIdx%><small><% If tot > 0 Then Response.Write Int(nIdx / tot * 100) & "%" %></small></div><div class="bar"><i style="width:<% If tot > 0 Then Response.Write Int(nIdx / tot * 100) Else Response.Write 0 %>%"></i></div></section>
  <section class="card stat"><div class="k"><i class="warn"><svg width="16" height="16"><use href="#i-alert"/></svg></i>待建 / 需處理</div><div class="v"><%=nTodo%><small>／ <%=nErr%></small></div><div class="d">需處理 = 找不到檔案、掃描檔、轉檔失敗</div></section>
  <section class="card stat"><div class="k"><i class="teal"><svg width="16" height="16"><use href="#i-file"/></svg></i>全文段落</div><div class="v"><%=FormatNumber(nChunk, 0)%><small>段</small></div><div class="d">無 PDF（僅影片/簡報）<%=nNoPdf2%> 份</div></section>
</div>

<section class="card" style="margin-bottom:24px">
  <div class="card-head"><div class="card-ico"><svg width="20" height="20"><use href="#i-refresh"/></svg></div><div><h2>更新流程</h2><p>標準有新增、改版或停用時，依序執行這兩步。</p></div></div>
  <div class="card-body">
    <div class="flow-cards">
      <div class="flow-card"><span class="n">1</span><div><b>同步標準清單與檔名</b><p>讀取 amend_work.mdb 的 catalog，更新標準名稱、PDF 路徑與教學影片；停用的標準會自動移除。</p>
        <div class="row"><a class="btn btn-primary" href="S_coach_build.asp?act=sync" onclick="this.classList.add('loading')"><span class="spin"></span><svg width="18" height="18" class="lbl-ico"><use href="#i-db"/></svg>同步護理部共通標準</a>
        <% If Trim(SC_SUS_DB) <> "" Then %><a class="btn btn-secondary" href="S_coach_build.asp?act=syncsus" onclick="this.classList.add('loading')"><span class="spin"></span><svg width="18" height="18" class="lbl-ico"><use href="#i-db"/></svg>同步單位次專科標準</a><% End If %></div></div></div>
      <div class="flow-card"><span class="n">2</span><div><b>建立全文索引</b><p>只處理新的或 PDF 有更新的標準；把 PDF 轉成文字並切成段落，供問答與關鍵字查詢使用。</p>
        <div class="row"><a class="btn btn-primary" href="S_coach_build.asp?act=build&mode=changed" onclick="this.classList.add('loading')"><span class="spin"></span><svg width="18" height="18" class="lbl-ico"><use href="#i-bolt"/></svg>建立索引</a></div></div></div>
    </div>
    <div class="tools-row">
      <span class="note">進階：</span>
      <a class="btn btn-warn btn-sm" href="S_coach_build.asp?act=build&mode=all" onclick="return confirm('全部重建約需數分鐘，確定？')"><svg width="16" height="16"><use href="#i-refresh"/></svg>全部重建</a>
      <a class="btn btn-ghost btn-sm" href="S_coach_build.asp?act=scan" onclick="return confirm('備用功能：不看資料庫，改用檔名自動猜測對應（會覆蓋自動對應的檔名）。確定？')"><svg width="16" height="16"><use href="#i-search"/></svg>備用：掃描資料夾猜測對應</a>
    </div>
  </div>
</section>

<section class="card">
  <div class="toolbar">
    <label class="big-input"><svg width="18" height="18"><use href="#i-search"/></svg><input id="tf" placeholder="篩選編號、名稱或檔名"></label>
    <div class="chips" id="stChips">
      <button class="chip on" data-g="">全部<span class="n"><%=tot%></span></button>
      <button class="chip" data-g="ok">已建索引<span class="n"><%=nIdx%></span></button>
      <button class="chip" data-g="todo">待建索引<span class="n"><%=nTodo%></span></button>
      <button class="chip" data-g="err">需處理<span class="n"><%=nErr%></span></button>
      <button class="chip" data-g="nopdf">無 PDF<span class="n"><%=nNoPdf2%></span></button>
    </div>
  </div>
  <div class="table-wrap">
  <table class="lst admin">
    <thead><tr><th style="width:110px">編號</th><th>標準名稱</th><th class="hide-m">PDF 檔案</th><th class="hide-m" style="width:90px;text-align:right">頁 / 字數</th><th style="width:130px">狀態</th><th class="hide-m" style="width:300px">手動指定 PDF</th></tr></thead>
    <tbody id="tb">
<%
Set rs = conn.Execute("SELECT * FROM std_list ORDER BY IIf(status='已建索引',1,0), code")
Do Until rs.EOF
  stat = rs("status") & ""
  grp = StGroup(stat)
%>
<tr data-g="<%=grp%>" data-s="<%=H(LCase(rs("code") & " " & rs("title") & " " & rs("pdf_file") & " " & rs("pdf_manual")))%>">
  <td><span class="tag code"><%=H(rs("code"))%></span></td>
  <td class="ttl"><%=H(rs("title"))%><span class="cell-sub"><%=H(rs("cat"))%></span></td>
  <td class="hide-m"><% If rs("pdf_manual") & "" <> "" Then %><span class="badge info" style="margin-bottom:4px">手動</span> <span class="mono"><%=H(rs("pdf_manual"))%></span><% Else %><span class="mono"><%=H(rs("pdf_file"))%></span><% End If %></td>
  <td class="hide-m num"><% If Not IsNull(rs("pages")) Then Response.Write rs("pages") & " 頁<span class='cell-sub'>" & FormatNumber(rs("chars"), 0) & " 字</span>" %></td>
  <td><%=StBadge(stat)%></td>
  <td class="hide-m"><form class="inline-form" method="post" action="S_coach_build.asp"><input type="hidden" name="act" value="setfile"><input type="hidden" name="code" value="<%=H(rs("code"))%>">
    <input class="input input-sm" type="text" name="f" value="<%=H(rs("pdf_manual"))%>" placeholder="例 CP\CP001_1001_抽痰術.pdf"><button class="btn btn-secondary btn-sm">儲存</button></form></td>
</tr>
<%
  rs.MoveNext
Loop
rs.Close: conn.Close
%>
    </tbody>
  </table>
  </div>
  <div class="empty" id="noRow" style="display:none">沒有符合條件的標準</div>
</section>

<script>
(function(){
  var g='', tf=document.getElementById('tf'), rows=[].slice.call(document.querySelectorAll('#tb tr'));
  function draw(){
    var q=tf.value.trim().toLowerCase(), n=0;
    rows.forEach(function(r){var ok=(!g||r.dataset.g===g)&&(!q||r.dataset.s.indexOf(q)>=0);r.style.display=ok?'':'none';if(ok)n++;});
    document.getElementById('noRow').style.display=n?'none':'block';
  }
  [].slice.call(document.querySelectorAll('#stChips .chip')).forEach(function(c){c.onclick=function(){g=c.dataset.g;[].slice.call(document.querySelectorAll('#stChips .chip')).forEach(function(x){x.classList.toggle('on',x===c);});draw();};});
  tf.oninput=draw;
})();
</script>
<% SC_Bottom %>
