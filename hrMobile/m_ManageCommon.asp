<%
'If UserRank = 0 Then Response.Write GetManagePermit()

'======== 返回无管理权限提示 ========
Function GetManagePermit()
	Dim strFun : strFun = ""
	If UserRank=0 Then
		strFun = "<div class=""hr-permit-tips""><dl class=""hr-rows tipsBox""><dt><i class=""hr-icon"">&#xf05e;</i></dt><dd>您没有访问权限！<br><b>Access to the page is denied!</b></dd></dl>"
		strFun = strFun & "<h4 class=""back-btn""><a href=""" & ParmPath & "Index.html"" class=""hr-btn"">返回首页</a></h4></div>"
	End If
	GetManagePermit = strFun
End Function

%>