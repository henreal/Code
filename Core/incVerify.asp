<%
'******************* 验证及安全检测 *******************
'=====================================================================
'函数名：ChkUserLogin
'作  用：判断会员是否已经登陆
'返回值：True/False
'=====================================================================
Function ChkUserLogin()
	Dim ChkPass, ChkRandCode
    Dim rsChk, sqlChk
    UserID = 0
    ChkUserLogin = False
	UserYGDM = HR_Clng(Request.Cookies(Site_Sn)("YGDM"))
    ChkPass = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserPass")))
    ChkRandCode = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("RndCode")))

    If UserYGDM <> "" And ChkPass <> "" And ChkRandCode <> "" Then
		sqlChk = "Select YGDM,YGXM From HR_Teacher Where TeacherID>0"
		sqlChk = sqlChk & " And YGDM='" & UserYGDM & "' And LoginPass='" & ChkPass & "'"
		Set rsChk = Conn.Execute(sqlChk)
			If Not(rsChk.BOF And rsChk.EOF) Then
				UserID = 0
				UserYGXM = Trim(rsChk("YGXM"))
				UserRank = 0
				ChkUserLogin = True
			End If
		Set rsChk = Nothing
	Else
		UserID = 0 : UserRank = 0
		UserYGDM = "0" : UserYGXM = ""
	End If
End Function

'=====================================================================
'函数名：ChkComeUrl
'作  用：验证来源网址是否为合法，是否跨域提交
'返回值：True/False
'=====================================================================
Function ChkComeUrl()
	Dim cServer_Name:cServer_Name = Request.ServerVariables("SERVER_NAME")
	ChkComeUrl = False
	If Instr(LCase(ComeUrl), LCase(cServer_Name)) > 0 Then ChkComeUrl = True
End Function

'=====================================================================
'函数名：IsValidEmail		【检查邮箱是否正确返回:True|False】
'=====================================================================
Function IsValidEmail(Email)
	IsValidEmail = False
    regEx.Pattern = "^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$"
	'regEx.Pattern = "^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$"		'第二种正则
	If Len(Email) > 0 Then IsValidEmail = regEx.Test(Email)
End Function

'=====================================================================
'函数名：IsValidUserName		【检查用户名是否正确返回:True|False】
'=====================================================================
Function IsValidUserName(fName)
	IsValidUserName = False
    regEx.Pattern = "^[a-zA-Z][a-zA-Z0-9\-\_]{5,17}$"
	If HR_IsNull(fName)=False Then IsValidUserName = regEx.Test(fName)
End Function

'=====================================================================
'函数名：IsValidURL			【检查URL是否合法，返回:True|False】
'=====================================================================
Function IsValidURL(iURL)
	IsValidURL = False
	Dim fUrl : fUrl = Trim(iURL)
	If Len(fUrl) > 0 Then
		If Left(LCase(fUrl), 7) <> "http://" Then fUrl = "http://" & LCase(fUrl)
		regEx.Pattern = "^((https?|ftp|rtsp|mms|file|ms-help):((//)|(\\\\))+[\w\d:#@%/;$()~_?\+-=\\\.&]*)$"
		'regEx.Pattern = "^((http|https|ftp):(\/\/|\\\\)((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)$"	'第二种
		IsValidURL = regEx.Test(fUrl)
	End If
End Function

'=====================================================================
'函数名：IsValidPhone		【判断是否为电话号码，返回布尔值】
'=====================================================================
Function IsValidPhone(fPhoneNum, fType)
	IsValidPhone = False
	If HR_Clng(fType) = 1 Then
		regEx.Pattern = "^[1][345789]\d{9}$"	'验证手机
	ElseIf HR_Clng(fType) = 2 Then
		regEx.Pattern = "(^[0-9]{7,8}$)|(^0\d{2,3}\d{7,8}$)|(^0\d{2,3}-\d{7,8}$)"	'验证固话，包括：带区号、带分隔线、不带区号【禁止3位、5位客服号及分机号】
	Else
		regEx.Pattern = "(^[0-9]{7,8}$)|(^0\d{2,3}\d{7,8}$)|(^0\d{2,3}-\d{7,8}$)|(^[1][345789]\d{9}$)"		'固话和手机正则
	End If
	If HR_IsNull(fPhoneNum) = False Then IsValidPhone = regEx.Test(fPhoneNum)
End Function

'=====================================================================
'函数名：IsValidPassword	【判断密码是否符合要求，返回布尔值】
'注：密码由6-22位字母、数学及!@#$%^&_-+组成
'=====================================================================
Function IsValidPassword(iPassword)			'明文
	IsValidPassword = False
	regEx.Pattern = "^[\@A-Za-z0-9\!\#\$\%\^\&\_\-\+]{6,22}$"
	'regExpS="(?<one>[0-9])|(?<two>[a-z])|(?<four>[A-Z])|(?<three>[~!@#$%^&*_])";
	If Len(iPassword) > 0 Then IsValidPassword = regEx.Test(iPassword)
End Function

'=====================================================================
'函数名：IsValidIDCard	【验证身份证是否合法，返回布尔值】
'=====================================================================
Function IsValidIDCard(iIDCard)
	IsValidIDCard = False
	regEx.Pattern = "^\d{6}(18|19|20)?\d{2}(0[1-9]|1[012])(0[1-9]|[12]\d|3[01])\d{3}(\d|[xX])$"
	If Len(iIDCard) > 0 Then IsValidIDCard = regEx.Test(iIDCard)
End Function

%>
