<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incMD5.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim scriptCtrl : SiteTitle = "考核项目API"
Dim strParm, arrParm : strParm = Trim(Request("Parm")) : arrParm = Split(strParm, "/")
If HR_isNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim str_data : ErrMsg = "未知错误"
	sql = "select * from HR_Class where ClassType=1 And Template='TempTableA'"
	Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.Open sql, conn, 1, 1
		If Not(rs.BOF And rs.EOF) Then
			i = 0
			Do While Not rs.EOF
				If i>0 Then str_data = str_data & ","
				str_data = str_data & "{""item"":""" & rs("ClassName") & """, ""id"":" & HR_Clng(rs("ClassID")) & ", ""Tips"":""" & HR_HTMLEncode(rs("Tips")) & """}"
				rs.MoveNext
				i = i + 1
			Loop
			ErrMsg = "有数据"
		Else
			ErrMsg = "没有数据"
		End If
	Set rs = Nothing
	strTmp = "{""err"":true, ""errcode"":500, ""errmsg"":""" & ErrMsg & """, ""icon"":2, ""data"":[" & str_data & "]}"
	Response.Write strTmp
End Sub
%>