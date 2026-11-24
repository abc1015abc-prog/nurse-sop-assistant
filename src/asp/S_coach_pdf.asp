<%@ Language="VBScript" CodePage="65001" EnableSessionState="False" %>
<!--#include file="S_coach_conn.asp"-->
<%
' 開啟標準 PDF：S_coach_pdf.asp?code=CP001#page=3 （瀏覽器 PDF 檢視器會直接跳到第 3 頁）
SC_CheckLogin
Dim conn, code, path, fso, st, sz, pos
code = Trim(Request.QueryString("code"))
Set conn = SC_Open()
path = SC_PdfPath(conn, code)
conn.Close
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If path = "" Or Not fso.FileExists(path) Then
  Response.Charset = "utf-8"
  Response.Write "<p style='font-family:sans-serif;padding:2em'>找不到此標準的 PDF 檔（" & H(code) & "），請通知管理者。</p>"
  Response.End
End If
Response.Clear          ' 清掉 include 產生的換行，避免破壞 PDF
Response.ContentType = "application/pdf"
Response.AddHeader "Content-Disposition", "inline; filename=""" & code & ".pdf"""
Set st = Server.CreateObject("ADODB.Stream")
st.Type = 1: st.Open
st.LoadFromFile path
sz = st.Size
Response.AddHeader "Content-Length", CStr(sz)
pos = 0
Do While pos < sz And Response.IsClientConnected
  Response.BinaryWrite st.Read(262144)
  Response.Flush
  pos = pos + 262144
Loop
st.Close
Response.End
%>