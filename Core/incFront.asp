<%
'********** 公用标签处理及相关调用函数解析 **********
' Powered By：Henreal Studio
' Update：Henreal SMCS V1.0.23 Build 20170208
' Website：http://www.henreal.com
' Weixin：Henreal-Net【恒锐网络科技】
' Tel：0831-8239995 / 13700999995
'----------------------------------------------------



'=====================================================================
' 函数名：ReplaceCommonLabel()			作用：替换全局通用标签[R1]
'=====================================================================
Function ReplaceCommonLabel(iHtml)
	If IsNull(iHtml) Or iHtml = "" Then Exit Function
	'MetaKeywords = XmlText("Config", "MetaKeywords", "")
	'MetaDescription = XmlText("Config", "MetaDescription", "")
	
	Dim CommFunLabel, strFunLabel, arrFunLabel, TmpFunLabel, FunMatches, x
	
	'=====> L0级 全站通用标签
	regEx.Pattern = "\[\@ShowLinks\_(.*?)\]"			'----- 替换友情链接函数式标签
	strFunLabel = ""
	Set FunMatches = regEx.Execute(iHtml)
		For Each x In FunMatches
			arrFunLabel = Split(x.SubMatches(0), "_") : strFunLabel = GetListLinks(arrFunLabel(0), arrFunLabel(1), arrFunLabel(2))
			iHtml = Replace(iHtml, x.Value, strFunLabel)
		Next
	Set FunMatches = Nothing
	
	regEx.Pattern = "\[\@HeadHtml\_(.*?)\]"				'----- 替换顶部通用HTML
	strFunLabel = ""
	Set FunMatches = regEx.Execute(iHtml)
		For Each x In FunMatches
			arrFunLabel = Split(x.SubMatches(0), "_") : strFunLabel = GetHeadHtml(arrFunLabel(0), arrFunLabel(1))
			iHtml = Replace(iHtml, x.Value, strFunLabel)
		Next
	Set FunMatches = Nothing
	
	regEx.Pattern = "\[\@FootHtml\_(.*?)\]"				'----- 替换底部通用HTML
	strFunLabel = ""
	Set FunMatches = regEx.Execute(iHtml)
		For Each x In FunMatches
			arrFunLabel = Split(x.SubMatches(0), "_") : strFunLabel = GetFootHtml(arrFunLabel(0), arrFunLabel(1))
			iHtml = Replace(iHtml, x.Value, strFunLabel)
		Next
	Set FunMatches = Nothing

	regEx.Pattern = "\[\@Head\_(.*?)\]"				'----- 替换顶部导航HTML
	strFunLabel = ""
	Set FunMatches = regEx.Execute(iHtml)
		For Each x In FunMatches
			arrFunLabel = Split(x.SubMatches(0), "_")
			If Ubound(arrFunLabel) > 0 Then strFunLabel = GetHeaderNav(arrFunLabel(1))
			iHtml = Replace(iHtml, x.Value, strFunLabel)
		Next
	Set FunMatches = Nothing

	regEx.Pattern = "\[\@Foot\_(.*?)\]"				'----- 替换底部导航HTML
	strFunLabel = ""
	Set FunMatches = regEx.Execute(iHtml)
		For Each x In FunMatches
			arrFunLabel = Split(x.SubMatches(0), "_") : strFunLabel = GettFootNav(arrFunLabel(1))
			iHtml = Replace(iHtml, x.Value, strFunLabel)
		Next
	Set FunMatches = Nothing
	
	iHtml = Replace(iHtml, "[@Site_Title]", SiteTitle)
	iHtml = Replace(iHtml, "[@Site_Name]", SiteName)
	iHtml = Replace(iHtml, "[@Site_Url]", SiteUrl)
	iHtml = Replace(iHtml, "[@Web_Dir]", InstallDir)
	iHtml = Replace(iHtml, "[@Manage_Dir]", ManageDir)
	iHtml = Replace(iHtml, "[@UploadDir]", UploadDir)
	iHtml = Replace(iHtml, "[@VerifyCode]", VerifyCodeFile)
	iHtml = Replace(iHtml, "[@Page_Keywords]", MetaKeywords)
	iHtml = Replace(iHtml, "[@Page_Description]", MetaDescription)
	
	Copyright = XmlText("Config", "CopyRight", "")
	iHtml = Replace(iHtml, "[@FootCopyRight]", Copyright)
	iHtml = Replace(iHtml, "[@MyCompany]", XmlText("Contact", "Company", ""))
	iHtml = Replace(iHtml, "[@MyCompanyEN]", XmlText("Contact", "CompanyEN", ""))
	iHtml = Replace(iHtml, "[@ShowPhone_1]", XmlText("Contact", "Tel1", ""))
	iHtml = Replace(iHtml, "[@ShowPhone_2]", XmlText("Contact", "Tel2", ""))
	iHtml = Replace(iHtml, "[@FaxPhone]", XmlText("Contact", "FaxPhone", ""))
	iHtml = Replace(iHtml, "[@Mail_1]", XmlText("Contact", "eMail1", ""))
	iHtml = Replace(iHtml, "[@Mail_2]", XmlText("Contact", "eMail2", ""))
	iHtml = Replace(iHtml, "[@ContactADD]", XmlText("Contact", "Address", ""))
	iHtml = Replace(iHtml, "[@ContactQQ_1]", XmlText("Contact", "TenQQ1", ""))
	iHtml = Replace(iHtml, "[@ContactQQ_2]", XmlText("Contact", "TenQQ2", ""))
	iHtml = Replace(iHtml, "[@SiteMII]", XmlText("Contact", "MII", ""))
	iHtml = Replace(iHtml, "[@SiteVer]", XmlText("Config", "Ver", ""))
	ReplaceCommonLabel = iHtml
End Function


'=====================================================================
' 函数：ReplaceNewsSub()			作用：替换新闻函数标签
'=====================================================================
Sub ReplaceNewsSub()
	Dim strFunLabel, FunMatches, x, arrFunLabel
	regEx.Pattern = "\[\@ListNews\_(.*?)\]"			'共13个参数
	strFunLabel = ""
	Set FunMatches = regEx.Execute(strHtml)
		For Each x In FunMatches
			arrFunLabel = Split(x.SubMatches(0), "_")
			If Ubound(arrFunLabel) = 12 Then strFunLabel = GetListNews(arrFunLabel(0), arrFunLabel(1), arrFunLabel(2), arrFunLabel(3), arrFunLabel(4), arrFunLabel(5), arrFunLabel(6), arrFunLabel(7), arrFunLabel(8), arrFunLabel(9), arrFunLabel(10), arrFunLabel(11), arrFunLabel(12) )
			strHtml = Replace(strHtml, x.Value, strFunLabel)
		Next
	Set FunMatches = Nothing
End Sub

'=====================================================================
' 函数：ShowBodyCopy(iShow)			作用：显示欢迎页通用版权信息[R1]
'=====================================================================
Function ShowBodyCopy(iShow)
	Dim strFun
	strFun = "<div class=""BodyCopy"">"
	strFun = strFun & "<li class=""KF"">客服　QQ：146579</li>"
	strFun = strFun & "<li class=""Tel"">客服电话：0831-8239995 / 13700999995</li>"
	strFun = strFun & "<li class=""Site"" id=""Site"">官方网址：<a href=""http://www.henreal.com"" target=""_blank"">恒锐网络科技</a></li>"
	strFun = strFun & "</div>" & vbCrlf
	ShowBodyCopy = strFun
End Function

'=====================================================================
' 函数：HR_GB2UTF(字符串)			作用：GB2312转UTF-8
'=====================================================================
Function HR_U2UTF8(Byval a_iNum)
    Dim fResult:fResult = ""
    Dim fTemp, fHexNum
    fHexNum = Trim(a_iNum)
    If HR_IsNull(fHexNum) Then Exit Function
    If (fHexNum < 128) Then
        fResult = fResult & fHexNum
    ElseIf (fHexNum < 2048) Then
        fResult = ChrB(&H80 + (fHexNum And &H3F))
        fHexNum = fHexNum \ &H40
        fResult = ChrB(&HC0 + (fHexNum And &H1F)) & fResult
    ElseIf (fHexNum < 65536) Then
        fResult = ChrB(&H80 + (fHexNum And &H3F))
        fHexNum = fHexNum \ &H40
        fResult = ChrB(&H80 + (fHexNum And &H3F)) & fResult
        fHexNum = fHexNum \ &H40
        fResult = ChrB(&HE0 + (fHexNum And &HF)) & fResult
    End If
    HR_U2UTF8 = fResult
End Function
Function HR_GB2UTF(Byval a_sStr)
    Dim sGB, fResult, sTemp
    Dim iLen, iUnicode, iTemp, iArr
    sGB = Trim(a_sStr)
    iLen = Len(sGB)
	If HR_IsNull(sGB) Then Exit Function
    For iArr = 1 To iLen
         sTemp = Mid(sGB, iArr, 1)
         iTemp = Asc(sTemp)
         If (iTemp > 127 OR iTemp < 0) Then
             iUnicode = AscW(sTemp)
             If iUnicode < 0 Then iUnicode = iUnicode + 65536
        Else
            iUnicode = iTemp
        End If
        fResult = fResult & HR_U2UTF8(iUnicode)
    Next
    HR_GB2UTF = fResult
End Function


'=====================================================================
' 函数：GetStrImg(iStr)			作用：仅取编辑器中的图片[R1]
'=====================================================================
Function GetStrImg(iStr)
	Dim regF1, regF2, MatchF1, MatchF2, MatchesF1, MatchesF2
	Dim tStr1, tStr2, ImgNum, tPic
	GetStrImg = "":tStr1 = "":ImgNum = 1
	If Len(iStr) > 0 Then
		Set regF1 = New Regexp
			regF1.IgnoreCase = True
			regF1.Global = True
			regF1.Pattern = "src="".+?"""
			Set MatchesF1 = regF1.Execute(iStr)
				For Each MatchF1 in MatchesF1  
					tStr2 = Mid(MatchF1.Value, 6, Len(MatchF1.Value) - 6)
					tStr2 = Replace(tStr2, "/UploadFiles/", "")
					tStr1 = tStr1 & "," & tStr2
				Next
			Set MatchF1 = Nothing
		Set regF1 = Nothing
		tStr1 = FilterArrNull(DelRightComma(tStr1), ",")
		GetStrImg = Replace(tStr1, ",", "||")
	Else
		GetStrImg = ""
	End If
End Function

'=====================================================================
'函数名：GetClassHtmlDir(参数)		取栏目目录
'=====================================================================
Function GetClassHtmlDir(iModuleID, iClassID)
	Dim strFun, rsFun, sqlFun
	Dim ArrFPath, iFun, strParentDir, ModuleDir
	sqlFun = "Select * From HR_Class Where ModuleID=" & HR_Clng(iModuleID) & " And ClassID=" & HR_Clng(iClassID)
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			'ModuleDir = GetTypeName("HR_Module", "HtmlDir", "ModuleID", iModuleID)		'频道目录
			'If Len(ModuleDir) > 0 Then ModuleDir = ModuleDir & "/"
			If Len(rsFun("ParentPath")) > 0 Then
				ArrFPath = Split(rsFun("ParentPath"), ",")
				For iFun = 0 To Ubound(ArrFPath)
					strParentDir = GetTypeName("HR_Class", "ClassDir", "ClassID", HR_Clng(ArrFPath(iFun)))
					If Len(strParentDir) > 0 Then strParentDir = strParentDir & "/"
					strFun = strFun & strParentDir
				Next
			End If
			If Len(rsFun("ClassDir")) > 0 Then strFun = strFun & rsFun("ClassDir") & "/"
		End If
	Set rsFun = Nothing
	GetClassHtmlDir = ModuleDir & strFun
End Function

'=====================================================================
'函数名：GetClassPath(参数)		取栏目路径
'=====================================================================


'=====================================================================
'函数名：GetListNews(参数)		列表新闻
'iShowType：1.显示栏目，2.标题省略号，3.显示时间(列表)
'=====================================================================

'=====================================================================
'函数名：ChkTable		【检查表是否存在，返回布尔值】
'=====================================================================
Function ChkTable(tTableName)
	on error resume next
	Dim rsFun : ChkTable = False
	If isNull(tTableName) Or tTableName = "" Then Exit Function
	Set rsFun = Server.CreateObject("ADODB.RecordSet")
		rsFun.open "Select Top 1 * From " & tTableName, Conn, 1, 1
		If Not Err.Number=0 Then Err.Clear : Exit Function
		ChkTable = True
		rsFun.Close
	Set rsFun = Nothing 
End Function

'=====================================================================
'函数名：ConvertNumDate		【时间戳返回日期】
'=====================================================================
Function ConvertNumDate(timeStamp)
	If IsEmpty(timeStamp) or Not IsNumeric(timeStamp) Then
        ConvertNumDate = FormatDate(Now(), 1)
        Exit Function
    End If
	ConvertNumDate = DateAdd("d", timeStamp-2, "1900-01-01 00:00:00")		'减2调整时间差【因为PHPExcel未设置格式】
End Function
Function ConvertDateToNum(fTime)		'将日期转为时间戳
	If IsEmpty(fTime) Or Not IsDate(fTime) Or fTime="" Then fTime = FormatDate(Now(), 2) & " 23:59:59"
	ConvertDateToNum = DateDiff("d","1900-01-01 00:00:00", fTime)
End Function


'=====================================================================
'函数名：FilterHtmlToText(参数)		【过滤Html为文本，支付JSON】
'=====================================================================
Function FilterHtmlToText(fHtml)
	fHtml = Trim(nohtml(fHtml))
	If HR_IsNull(fHtml) = False Then
		fHtml = Replace(fHtml, chr(9), "")
		fHtml = Replace(fHtml, "\", "\\")			'去除TAB键
		fHtml = Replace(fHtml, chr(10), "")
		fHtml = Replace(fHtml, CHR(13), "")			'去除回车及换行符
	End If
	FilterHtmlToText = fHtml
End Function
%>
