<%
If HR_CBool(Request("ssoLogin")) And HR_Clng(Request("chk"))=1 Then
	'认证成功
	UserYGDM = HR_Clng(Request("ygdm"))
	Dim rsPurv
	Set rsPurv = Conn.Execute("Select * From HR_Teacher Where TeacherID>0 And YGDM='" & UserYGDM & "'")
		If Not(rsPurv.BOF And rsPurv.EOF) Then
			Response.Cookies(Site_Sn)("UserName") = ""
			Response.Cookies(Site_Sn)("YGDM") = UserYGDM
			Response.Cookies(Site_Sn)("UserPass") = rsPurv("LoginPass")
			Response.Cookies(Site_Sn)("RndCode") = "e46560952dce2deb"
		Else
			ErrMsg = "SSO登陆失败，您的员工信息可能未同步！<br /><a href=""" & InstallDir & "Desktop/Login.html"">帐号登陆</a>"
			Response.Write ErrMsg : Response.End
		End If
	Set rsPurv = Nothing
End If

If Not(ChkUserLogin()) Then			'由本系统登陆身份识别
	Response.Redirect ParmPath & "Login.html"
End If

Set rs = Conn.Execute("Select Top 1 * From HR_User Where YGDM=" & UserYGDM & " Order By UserID ASC")
	If Not(rs.BOF And rs.EOF) Then
		UserID = HR_Clng(rs("UserID"))
		UserRank = HR_Clng(rs("ManageRank"))
	End If
Set rs = Nothing
'//判断密码是否为初始密码
Function ChkInitPass()
	Dim tPassWord
	Dim ArrField : ArrField = GetTableDataQuery("HR_Teacher", "", 1, "YGDM='" & UserYGDM & "'")			'取教师信息
	tPassWord = ArrField(15,1)
	If tPassWord = "83aa400af464c76d" Then
		ErrMsg = "您的登陆密码为初始密码" : strTmp = "为了安全，请点击“返回”立即修改"
		ErrHref = "" & ParmPath & "Setup/ModiPass.html"
		Response.Write GetErrBody(0):Response.End
	End If
End Function
%>