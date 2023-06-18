<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
SiteTitle = "数据导出"
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "ExportAll" Call ExportAll()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Response.Write UserID & UserName
End Sub

Sub ExportAll()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))

	Dim tItemName, tTemplate, lenField, tSheetName
	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = "业绩考核项目" & tItemName & "不存在！<br>"
		End If
	Set rsTmp = Nothing
	If ChkTable(tSheetName) = False Then ErrMsg = ErrMsg & "数据表 " & tSheetName & " 不存在！<br>"	'检查数据表是否存在

	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From " & tSheetName & " Order By ID"), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			TotalPut = rsTmp.Recordcount
		End If
	Set rsTmp = Nothing
	Response.Write UserID & "/" & tItemName & "共有" & TotalPut & " 条数据，导出会花费一定的时间，请稍侯"
End Sub
%>