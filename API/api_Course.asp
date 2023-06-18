<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incMD5.asp"-->
<!--#include file="../Core/incVerify.asp"-->
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
	Dim itemID : itemID = HR_Clng(Request.QueryString("item"))
	Dim arrItem : arrItem = GetTableDataQuery("HR_Class", "", 1, "ClassID=" & itemID & "")		'取绩效考核项目信息
	Dim tTable : tTable = "HR_Sheet_" & itemID
	Call ChkUserLogin()		'验证登陆信息

	If HR_Clng(UserYGDM)=0 Or ChkDataTable(tTable, false)=False Then
		strTmp = "{""err"":true, ""errcode"":500, ""errmsg"":""考核项目不存在或未登陆"", ""icon"":2, ""data"":[" & str_data & "]}"
		Response.Write strTmp : Exit Sub
	End If

	sql = "select * from " & tTable & " where VA1=" & HR_Clng(UserYGDM) & ""
	Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.Open sql, conn, 1, 1
		If Not(rs.BOF And rs.EOF) Then
			i = 0
			Do While Not rs.EOF
				If i>0 Then str_data = str_data & ","
				str_data = str_data & "{""id"":" & rs("ID") & ", ""ygdm"":""" & rs("VA1") & """, ""ygxm"":""" & Trim(rs("VA2")) & """, ""item"":""" & itemID & """"
				str_data = str_data & ",""teach_time"":""" & FormatDate(ConvertNumDate(rs("VA4")),2) & """, ""course"":""" & Trim(rs("VA8")) & """, ""year"":""" & Trim(rs("scYear")) & """}"
				rs.MoveNext
				i = i + 1
			Loop
			ErrMsg = "有数据"
		Else
			ErrMsg = "没有数据" & sql
		End If
	Set rs = Nothing
	strTmp = "{""err"":true, ""errcode"":500, ""errmsg"":""" & ErrMsg & """, ""icon"":2, ""data"":[" & str_data & "]}"
	Response.Write strTmp
End Sub
%>