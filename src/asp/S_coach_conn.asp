<!--#include file="config.asp"-->
<%
'=====================================================================
' 護理標準學習教練  共用設定與函式  (S_coach_conn.asp)
' 本系統所有檔案皆為 UTF-8；只設定「本頁」Response.CodePage，
' 不動 Session.CodePage，不影響部內其他 BIG5 網頁。
'=====================================================================

'---------- 所有伺服器路徑、資料庫、AI 模型、管理者等設定都在 config.asp ----------
'           （新環境：複製 config.sample.asp 為 config.asp 再修改）

'=====================================================================
Function SC_Open()
  Dim c
  Set c = Server.CreateObject("ADODB.Connection")
  c.Open "Provider=" & SC_PROVIDER & ";Data Source=" & Server.MapPath(SC_DB)
  Set SC_Open = c
End Function

Function SC_Worknb()
  SC_Worknb = Trim(Request.Cookies("worknb") & "")
End Function

Function SC_IsAdmin()
  Dim w: w = UCase(SC_Worknb())
  SC_IsAdmin = (w <> "" And InStr(UCase(SC_ADMINS), "," & w & ",") > 0)
End Function

Sub SC_CheckLogin()
  If SC_REQUIRE_LOGIN And SC_Worknb() = "" Then
    Response.Redirect SC_HOME_URL
    Response.End
  End If
End Sub

Sub SC_CheckAdmin()
  If Not SC_IsAdmin() Then
    Response.Write "<p style='font-family:sans-serif;padding:2em'>此頁僅限管理者使用（請在 config.asp 的 SC_ADMINS 加入您的員工代號）。</p>"
    Response.End
  End If
End Sub

Function SQ(s)
  SQ = Replace(s & "", "'", "''")
End Function

Function H(s)
  If IsNull(s) Then H = "" Else H = Server.HTMLEncode(s & "")
End Function

' JSON 字串跳脫
Function JE(ByVal s)
  Dim re
  If IsNull(s) Then s = ""
  s = s & ""
  s = Replace(s, "\", "\\")
  s = Replace(s, """", "\""")
  s = Replace(s, vbCr, "")
  s = Replace(s, vbLf, "\n")
  s = Replace(s, vbTab, " ")
  Set re = New RegExp: re.Global = True
  re.Pattern = "[\x00-\x1F]"
  s = re.Replace(s, " ")
  JE = """" & s & """"
End Function

' 移除所有空白（檢索用正規化文字）
Function SC_Norm(ByVal s)
  Dim re
  Set re = New RegExp: re.Global = True
  re.Pattern = "[\s\u3000\u00A0]+"
  SC_Norm = re.Replace(LCase(s & ""), "")
End Function

' 讀 UTF-8 文字檔
Function SC_ReadUtf8(path)
  Dim st, fs
  Set fs = Server.CreateObject("Scripting.FileSystemObject")
  If Not fs.FileExists(path) Then
    Response.Write "<p style='color:#b00020;font-family:sans-serif'>找不到檔案：" & Server.HTMLEncode(path) & "</p>"
    Response.End
  End If
  Set st = Server.CreateObject("ADODB.Stream")
  st.Type = 2: st.Charset = "utf-8": st.Open
  st.LoadFromFile path
  SC_ReadUtf8 = st.ReadText
  st.Close
End Function

Function SC_CatalogPath()
  If InStr(SC_CATALOG, ":\") > 0 Or Left(SC_CATALOG, 2) = "\\" Then SC_CatalogPath = SC_CATALOG Else SC_CatalogPath = Server.MapPath(SC_CATALOG)
End Function

' 取得標準 PDF 的實體路徑（手動指定優先）
Function SC_PdfPath(conn, code)
  Dim rs, f
  SC_PdfPath = ""
  Set rs = conn.Execute("SELECT pdf_file, pdf_manual FROM std_list WHERE code='" & SQ(code) & "'")
  If Not rs.EOF Then
    f = Trim(rs("pdf_manual") & "")
    If f = "" Then f = Trim(rs("pdf_file") & "")
    If f <> "" Then
      If InStr(f, ":\") > 0 Or Left(f, 2) = "\\" Then SC_PdfPath = f Else SC_PdfPath = SC_PDF_ROOT & "\" & f
    End If
  End If
  rs.Close
End Function
%>
