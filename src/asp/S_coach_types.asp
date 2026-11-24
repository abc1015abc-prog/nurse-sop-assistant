<%
'=====================================================================
' 標準類別設定（S_coach_types.asp）
'   依編號開頭分類；名稱可自行修改，順序＝畫面上顯示的順序
'   注意：CPT 必須排在 CP 前面（CPT 也是 CP 開頭）
'=====================================================================
Const SC_TYPE_DEF = "CPT=照護準則|CP=護理技術|CR=護理常規|CE=儀器設備|EP=檢查護理|DG=紀錄表單|IG=資訊系統作業|NP=專科護理師技術|WM=病室管理|MM=資材管理"

' 依編號判斷類別代碼，例 CPT28 → CPT、CP001 → CP
Function SC_TypeOf(code)
  Dim defs, d, k, seg
  SC_TypeOf = ""
  ' 次專科標準編號：單位-類別-流水號，例 ICU-CPT-01、OR-CP-GS-41（護理部共通的 CP062-1 不算）
  seg = SC_SusType(code)
  If seg <> "" Then SC_TypeOf = seg: Exit Function
  defs = Split(SC_TYPE_DEF, "|")
  For Each d In defs
    k = Left(d, InStr(d, "=") - 1)
    If UCase(Left(code, Len(k))) = k Then
      If k = "CP" And UCase(Left(code, 3)) = "CPT" Then
        ' CPT 已在前面判斷過
      Else
        SC_TypeOf = k
        Exit Function
      End If
    End If
  Next
End Function

Function SC_TypeName(k)
  Dim defs, d
  SC_TypeName = k
  defs = Split(SC_TYPE_DEF, "|")
  For Each d In defs
    If Left(d, InStr(d, "=") - 1) = k Then
      SC_TypeName = Mid(d, InStr(d, "=") + 1)
      Exit Function
    End If
  Next
End Function

' 類別排序（數字越小越前面）
Function SC_TypeRank(k)
  Dim defs, i
  SC_TypeRank = 99
  defs = Split(SC_TYPE_DEF, "|")
  For i = 0 To UBound(defs)
    If Left(defs(i), InStr(defs(i), "=") - 1) = k Then
      SC_TypeRank = i
      Exit Function
    End If
  Next
End Function

' 前端用的類別清單 JSON：[["CPT","照護準則"],...]
Function SC_TypesJson()
  Dim defs, d, r
  r = ""
  defs = Split(SC_TYPE_DEF, "|")
  For Each d In defs
    If r <> "" Then r = r & ","
    r = r & "[""" & Left(d, InStr(d, "=") - 1) & """,""" & Mid(d, InStr(d, "=") + 1) & """]"
  Next
  SC_TypesJson = "[" & r & "]"
End Function

' 把前端傳來的類別（例 "CP,CPT"）轉成 SQL 條件；空白＝不限
Function SC_TypeSql(types, fld)
  Dim t, r, re
  r = ""
  Set re = New RegExp
  re.Pattern = "^[A-Z]{2,4}$"
  For Each t In Split(UCase(types & ""), ",")
    t = Trim(t)
    If re.Test(t) Then
      If r <> "" Then r = r & " OR "
      If t = "CP" Then
        r = r & fld & " LIKE 'CP[0-9]%' OR " & fld & " LIKE '%-CP-%'"
      Else
        r = r & fld & " LIKE '" & t & "%' OR " & fld & " LIKE '%-" & t & "-%'"
      End If
    End If
  Next
  If r <> "" Then r = "(" & r & ")"
  SC_TypeSql = r
End Function

Function SC_TypeAllowed(types, code)
  Dim t
  If Trim(types & "") = "" Then
    SC_TypeAllowed = True
    Exit Function
  End If
  SC_TypeAllowed = False
  For Each t In Split(UCase(types), ",")
    If Trim(t) = SC_TypeOf(code) Then
      SC_TypeAllowed = True
      Exit Function
    End If
  Next
End Function

' 去掉 PDF 每頁重複的表頭雜訊（文件名稱、機密等級、頁碼/總頁數、文件編號、版次）
Function SC_CleanText(s)
  Dim re, t
  t = Replace(s & "", vbCr, "")
  Set re = New RegExp
  re.Global = True
  re.Multiline = True
  re.Pattern = "^.*(文件名稱|文件編號|機密等級|頁碼/總頁數|■一般|□敏感).*$"
  t = re.Replace(t, "")
  re.Pattern = "^\s*(N\d{3}-[\w\-]+|V\d{1,3}|\d{1,3}/\d{1,3})\s*$"
  t = re.Replace(t, "")
  re.Pattern = "\n{2,}"
  t = re.Replace(t, vbLf)
  re.Pattern = "^\n+|\n+$"
  t = re.Replace(t, "")
  SC_CleanText = t
End Function

' 提問紀錄加欄位：resp（完整回答 JSON，供日後重看）、ai_text（AI 摘要純文字，供列表預覽）
Sub SC_EnsureLogCols(c)
  Dim tr
  On Error Resume Next
  Set tr = c.Execute("SELECT TOP 1 resp, ai_text FROM qlog")
  If Err.Number <> 0 Then
    Err.Clear
    c.Execute "ALTER TABLE qlog ADD COLUMN resp MEMO"
    Err.Clear
    c.Execute "ALTER TABLE qlog ADD COLUMN ai_text MEMO"
  Else
    tr.Close
  End If
  Err.Clear
  On Error GoTo 0
End Sub

' 全形英數轉半形（ＣＶＣ → CVC、ｉ／ｏ → i/o），全形空白轉半形
Function SC_Half(s)
  Dim i, c, r, t
  t = s & "": r = ""
  For i = 1 To Len(t)
    c = AscW(Mid(t, i, 1))
    If c < 0 Then c = c + 65536
    If c >= &HFF01 And c <= &HFF5E Then
      r = r & ChrW(c - &HFEE0)
    ElseIf c = &H3000 Then
      r = r & " "
    Else
      r = r & Mid(t, i, 1)
    End If
  Next
  SC_Half = r
End Function

' 同義詞表加欄位：src（來源：預設／詞庫／標準自動／MeSH／手動）；alias 改成備忘欄位以容納較多別名
Sub SC_EnsureSynCols(c)
  Dim tr
  On Error Resume Next
  Set tr = c.Execute("SELECT TOP 1 src FROM syn")
  If Err.Number <> 0 Then
    Err.Clear
    c.Execute "ALTER TABLE syn ADD COLUMN src TEXT(20)"
  Else
    tr.Close
  End If
  Err.Clear
  Set tr = c.Execute("SELECT TOP 1 alias FROM syn")
  If Err.Number = 0 Then
    If tr.Fields("alias").Type <> 203 Then
      tr.Close
      c.Execute "ALTER TABLE syn ALTER COLUMN alias MEMO"
    Else
      tr.Close
    End If
  End If
  Err.Clear
  On Error GoTo 0
End Sub

' 回饋（fb）與已審核答案（qa）資料表，沒有就自動建立（建立失敗會顯示原因）
Function SC_HasTable(c, tname)
  Dim rt
  SC_HasTable = False
  Set rt = c.OpenSchema(20)          ' adSchemaTables
  Do Until rt.EOF
    If LCase(rt("TABLE_NAME") & "") = LCase(tname) Then SC_HasTable = True: Exit Do
    rt.MoveNext
  Loop
  rt.Close
End Function

Sub SC_MakeTable(c, tname, ddl)
  Dim em
  If SC_HasTable(c, tname) Then Exit Sub
  On Error Resume Next
  c.Execute ddl
  em = ""
  If Err.Number <> 0 Then em = Err.Description & "（" & Hex(Err.Number) & "）"
  Err.Clear
  On Error GoTo 0
  If em <> "" Then
    Response.Clear
    Response.Write "<p style='font-family:sans-serif;padding:2em'>無法建立資料表 " & tname & "：" & Server.HTMLEncode(em) & _
      "<br>請確認 IIS 帳號對 S_coach.mdb 所在資料夾有寫入權限，或把此訊息截圖給系統管理者。</p>"
    Response.End
  End If
End Sub

Sub SC_EnsureFbTables(c)
  SC_MakeTable c, "fb", "CREATE TABLE [fb] ([id] COUNTER CONSTRAINT pk_fb PRIMARY KEY, [logid] LONG, [qaid] LONG, [worknb] TEXT(20), [unit] TEXT(30), [ftime] DATETIME, " & _
    "[rating] SHORT, [reason] TEXT(100), [note] TEXT(255), [question] TEXT(255), [codes] TEXT(255), [status] TEXT(10), [fix_qa] LONG)"
  SC_MakeTable c, "qa", "CREATE TABLE [qa] ([id] COUNTER CONSTRAINT pk_qa PRIMARY KEY, [question] TEXT(255), [keywords] TEXT(255), [answer] MEMO, [refs] TEXT(255), " & _
    "[ref_ver] MEMO, [reviewer] TEXT(20), [rtime] DATETIME, [status] TEXT(10), [hits] LONG, [src_log] LONG)"
End Sub

' 出處字串「CP016:6,7;CP018:2」→ 只保留合法字元
Function SC_CleanRefs(s)
  Dim re, parts, p, cd, pg, r
  Set re = New RegExp: re.Global = True
  r = ""
  parts = Split(Replace(Replace(Replace(s & "", "；", ";"), "：", ":"), vbLf, ";"), ";")
  For Each p In parts
    p = Trim(p)
    If p <> "" Then
      If InStr(p, ":") > 0 Then
        cd = Trim(Left(p, InStr(p, ":") - 1)): pg = Mid(p, InStr(p, ":") + 1)
      Else
        cd = p: pg = ""
      End If
      re.Pattern = "[^A-Za-z0-9\-\.]": cd = UCase(re.Replace(cd, ""))
      re.Pattern = "[^0-9,]": pg = re.Replace(Replace(Replace(pg, "、", ","), "，", ","), "")
      re.Pattern = ",{2,}": pg = re.Replace(pg, ",")
      If Left(pg, 1) = "," Then pg = Mid(pg, 2)
      If Right(pg, 1) = "," Then pg = Left(pg, Len(pg) - 1)
      If cd <> "" Then
        If r <> "" Then r = r & ";"
        r = r & cd & IIf4(pg <> "", ":" & pg, "")
      End If
    End If
  Next
  SC_CleanRefs = r
End Function

Function IIf4(c, a, b)
  If c Then IIf4 = a Else IIf4 = b
End Function

' 出處標準目前的版本（修訂日＋PDF 檔案時間）；標準改版後與存下來的不同 → 需重新確認
Function SC_StdVer(c, refs)
  Dim p, cd, r, rv
  r = ""
  For Each p In Split(refs & "", ";")
    cd = Trim(p)
    If InStr(cd, ":") > 0 Then cd = Left(cd, InStr(cd, ":") - 1)
    If cd <> "" Then
      Set rv = c.Execute("SELECT rev_date, pdf_mtime FROM std_list WHERE code='" & Replace(cd, "'", "''") & "'")
      If r <> "" Then r = r & ";"
      If rv.EOF Then
        r = r & cd & "=(無此標準)"
      Else
        r = r & cd & "=" & rv("rev_date") & "@" & rv("pdf_mtime")
      End If
      rv.Close
    End If
  Next
  SC_StdVer = r
End Function

' 是否為次專科標準：第二段是類別代碼（ICU-CPT-01 → CPT）；不是則回傳空字串
Function SC_SusType(code)
  Dim p, d
  SC_SusType = ""
  If InStr(code & "", "-") = 0 Then Exit Function
  p = Split(UCase(code & ""), "-")
  If UBound(p) < 2 Then Exit Function
  For Each d In Split(SC_TYPE_DEF, "|")
    If Left(d, InStr(d, "=") - 1) = p(1) Then SC_SusType = p(1): Exit Function
  Next
End Function

Function SC_SusCond(fld)
  Dim d, r
  r = ""
  For Each d In Split(SC_TYPE_DEF, "|")
    If r <> "" Then r = r & " OR "
    r = r & fld & " LIKE '%-" & Left(d, InStr(d, "=") - 1) & "-%'"
  Next
  SC_SusCond = "(" & r & ")"
End Function

' 標準來源篩選：""＝全部、ND＝護理部共通、SUS＝單位次專科（全部單位）、SUS:ICU＝某單位次專科
' 次專科標準編號為「單位-類別-流水號」（例 ICU-CPT-01）
Function SC_SrcSql(src, fld)
  Dim s, u, re
  s = UCase(Trim(src & ""))
  SC_SrcSql = ""
  If s = "ND" Then
    SC_SrcSql = "(NOT " & SC_SusCond(fld) & ")"
  ElseIf s = "SUS" Then
    SC_SrcSql = SC_SusCond(fld)
  ElseIf Left(s, 4) = "SUS:" Then
    Set re = New RegExp: re.Global = True: re.Pattern = "[^A-Z0-9]"
    u = re.Replace(Mid(s, 5), "")
    If u <> "" Then SC_SrcSql = "(" & fld & " LIKE '" & u & "-%' AND " & SC_SusCond(fld) & ")"
  End If
End Function

' std_list 加「來源」欄（SUS＝次專科），同步時用來區分
Sub SC_EnsureStdSrc(c)
  Dim tr
  On Error Resume Next
  Set tr = c.Execute("SELECT TOP 1 src FROM std_list")
  If Err.Number <> 0 Then
    Err.Clear
    c.Execute "ALTER TABLE std_list ADD COLUMN src TEXT(10)"
  Else
    tr.Close
  End If
  Err.Clear
  On Error GoTo 0
End Sub
%>
