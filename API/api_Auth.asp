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

Dim scriptCtrl
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case Else Response.Write "{""err"":true,""errcode"":500,""errmsg"":""接口参数错误"",""icon"":3}"
End Select

Sub MainBody()
	'Response.Write "{""err"":true,""errcode"":500,""errmsg"":""未被授权的访问"",""icon"":3}"
	Call Collect()
End Sub

Sub Collect()
	Dim tYear : tYear = HR_CLng(Request("year"))
	Dim tKey : tKey = Trim(Request("key"))
	If tYear = 0 Then tYear = DefYear

	If HR_IsNull(tKey) Then
		Response.Write "{""err"":true,""errcode"":404,""errmsg"":""参数key不能为空"",""icon"":3}"
		Exit Sub
	End If

	Set rsTmp = Conn.Execute("Select * From HR_Interface Where ApiKey='" & tKey & "'")
		If rsTmp.BOF And rsTmp.EOF Then
			Response.Write "{""err"":true,""errcode"":501,""errmsg"":""未被授权的访问"",""icon"":3}"
			Exit Sub
		End If
	Set rsTmp = Nothing

	SiteTitle = "项目学时汇总"
	Dim reJson, rows
	Dim tSumVA3, noSumVA3, CountYGDM, tItemName, tSheetName
	Set rs = Conn.Execute("Select * From HR_Class Where ModuleID=1001 Order By ClassType ASC, RootID ASC, OrderID ASC")
		If Not(rs.BOF And rs.EOF) Then
			i = 0
			Do While Not rs.EOF
				If i > 0 Then rows = rows & ","

				tSumVA3 = 0 : noSumVA3 = 0 : CountYGDM=0
				tItemName = Trim(rs("ClassName"))
				rows = rows & "{""item"":""" & tItemName & ""","

				If rs("Child") > 0 Then
					rows = rows & """parent"":true,""passed"":0,""nopass"":0,""total"":0,""teacher"":0"
				Else
					tSheetName = "HR_Sheet_" & rs("ClassID")
					If ChkTable(tSheetName) Then
						Set rsTmp = Conn.Execute("Select Sum(VA3) From " & tSheetName & " Where scYear=" & tYear & " And Passed=" & HR_True)
							tSumVA3 = rsTmp(0)
						Set rsTmp = Nothing
						Set rsTmp = Conn.Execute("Select Sum(VA3) From " & tSheetName & " Where scYear=" & tYear & " And Passed=" & HR_False)
							noSumVA3 = rsTmp(0)
						Set rsTmp = Nothing
						Set rsTmp = Server.CreateObject("ADODB.RecordSet")
							rsTmp.Open("Select Count(VA1) From " & tSheetName & " Where scYear=" & tYear & " Group By VA1"), Conn, 1, 1
							CountYGDM = rsTmp.Recordcount
						Set rsTmp = Nothing
					End If
					rows = rows & """parent"":false,""passed"":" & HR_CDbl(tSumVA3) & ","
					rows = rows & """nopass"":" & HR_CDbl(noSumVA3) & ","
					rows = rows & """total"":" & HR_CDbl(tSumVA3) + HR_CDbl(noSumVA3) & ","
					rows = rows & """teacher"":" & CountYGDM & ""
				End If
				rows = rows & "}"
				rs.MoveNext
				i = i + 1
			Loop
		End If
	Set rs = Nothing
	reJson = "{""err"":false,""errcode"":0,""errmsg"":""查询成功！"",""action"":""summary"",""year"":""" & tYear & """,""data"":[" & rows & "]}"
	Response.Write reJson
End Sub
%>