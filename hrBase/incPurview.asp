<%
If Not(ChkUserLogin()) Then			'由本系统登陆身份识别
	Response.Redirect InstallDir & "Desktop/Login.html"
	Response.End
End If

'重新判断是否有管理权限
Dim UserStuType
Set rs = Conn.Execute("Select Top 1 * From HR_User Where YGDM=" & UserYGDM & " Order By UserID ASC")
	If Not(rs.BOF And rs.EOF) Then
		UserID = HR_Clng(rs("UserID"))
		UserRank = HR_Clng(rs("ManageRank"))
		UserFace = Trim(rs("UserFace"))
		UserStuType = FilterArrNull(rs("StuType"), ",")
	End If
Set rs = Nothing

If UserRank = 0 Then		'判断是否为管理员
	ErrMsg = "您没有访问权限！"
	Response.Write GetErrBody(1) : Response.End
End If
%>