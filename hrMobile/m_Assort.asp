<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="incCommon.asp"-->
<%
SiteTitle = "更多功能"
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "WinAssort" Call WinAssort()			'弹窗分类选择
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	ErrMsg = "更多功能开发中"
	Response.Write GetErrBody(0)
End Sub

Sub WinAssort()
	Dim tmpID : tmpID = Trim(Request("PID"))
	Dim tModuleID : tModuleID = HR_CLng(Request("ModuleID"))
	Response.Write GetAssortList(tmpID)
End Sub
%>