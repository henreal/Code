<%
'******************* 公用过程函数 *******************

Call GetConfig()
Call InitMain()

Sub GetConfig()		'获取系统配置信息
    'On Error Resume Next
    Dim rsConfig, sqlConfig
    sqlConfig = "Select Top 1 * From HR_Config Where ID=" & ConfigID
    sqlConfig = sqlConfig & " Order By ID ASC"
    
    Set rsConfig = Conn.Execute(sqlConfig)
    If rsConfig.BOF And rsConfig.EOF Then
        rsConfig.Close
        Set rsConfig = Nothing
        Response.Write "<div style=""padding:20px;text-align:center;color:red;"">系统配置参数丢失，系统无法正常运行！</div>"
        Response.End
        Exit Sub
    End If
	SiteName = rsConfig("SiteName")
	SiteTitle = rsConfig("SiteTitle")
	SiteUrl = rsConfig("SiteUrl")
	InstallDir = rsConfig("InstallDir")
	ManageDir = rsConfig("RearBase")
	UploadDir = rsConfig("UploadDir")
	Copyright = rsConfig("Copyright")
	If Len(rsConfig("MetaKeywords")) > 0 Then MetaKeywords = rsConfig("MetaKeywords")
	If Len(rsConfig("MetaDescription")) > 0 Then MetaDescription = rsConfig("MetaDescription")
    
    objName_FSO = rsConfig("FSO_Script")
    MailObject = rsConfig("MailObject")
    MailServer = rsConfig("MailServer")
    MailServerUserName = rsConfig("MailServerUserName")
    MailServerPassWord = rsConfig("MailServerPassword")
    MailDomain = rsConfig("MailDomain")
	If Right(InstallDir, 1) <> "/" Then InstallDir = InstallDir & "/"
	If Right(ManageDir, 1) <> "/" Then ManageDir = ManageDir & "/"
	If Right(UploadDir, 1) <> "/" Then UploadDir = UploadDir & "/"
	
    If rsConfig("SiteUrlType") = 1 Then
        If Right(SiteUrl, 1) = "/" Then SiteUrl = Left(SiteUrl, Len(SiteUrl) - 1)
        If Left(SiteUrl, 7) <> "http://" Then SiteUrl = "http://" & SiteUrl
        strInstallDir = SiteUrl & InstallDir
    Else
        strInstallDir = InstallDir
    End If
    VerifyCodeFile = InstallDir & "Extend/" & rsConfig("VerifyCodeFile")

    rsConfig.Close
    Set rsConfig = Nothing
End Sub

Sub InitMain()
	CurrentPage = 1
    If HR_CLng(LCase(Request("Page"))) > 0 Then CurrentPage = HR_CLng(LCase(Request("Page")))
    MaxPerPage = HR_CLng(Trim(Request("MaxPerPage")))
    SoType = HR_CLng(Trim(Request("SoType")))
    SoField = ReplaceBadChar(Trim(Request("SoField")))
    SoWord = ReplaceBadChar(Trim(Request("SoWord")))
        
    ComeUrl = FilterJs(Trim(Request("ComeUrl")))
    If ComeUrl = "" Then
        ComeUrl = FilterJs(Trim(Request.ServerVariables("HTTP_REFERER")))
    End If
    Action = ReplaceBadChar(Trim(Request("Action")))

    Site_Sn = Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME") & InstallDir), "/", ""), ".", "")
    Set XmlDoc = CreateObject("Microsoft.XMLDOM")
    XmlDoc.async = False
    
    If XmlDoc.Load(Server.MapPath(XMLDataPath)) = False Then	'加载XML数据包
		Response.Write "<div style=""margin:20px;text-align:center;font-size:14px;color:red;"">读取XML数据出错！请检查系统配置。</div>"
		Exit Sub
	End If
	If MaxPerPage <= 0 Then MaxPerPage = HR_Clng(XmlText("Config", "MaxPerPage", "30"))
	CurrentVer = XmlText("Config", "Ver", "Henreal CMS Ver 1.0")
	DefYear = HR_CLng(XmlText("Common", "Year", "2019"))

	ObjInstalled_FSO = IsObjInstalled(objName_FSO)
    If ObjInstalled_FSO = True Then
		Set FSO = Server.CreateObject(objName_FSO)
    Else
		Response.Write "<div style=""margin:20px;text-align:center;font-size:14px;color:red;"">FSO组件不可用，各种与FSO相关的功能都将出错！请配置好FSO组件名称。</div>"
		Exit Sub
	End If
End Sub

'=====================================================================
'函数名：ShowPage	作用：显示"上一页 下一页"等信息【继承父容器】
'参  数：sFileName  ----链接地址		TotalNumber ----总数量
'        MaxPerPage  ----每页数量		CurrentPage ----当前页
'        ShowTotal   ----是否显示总数量
'        ShowAllPages ---是否用下拉列表显示所有页面以供跳转。
'        strUnit     ----计数单位		ShowMaxPerPage  ----是否显示每页信息量选项框
'返回值："上一页 下一页"等信息的HTML代码
'=====================================================================
Function ShowPage(sfilename, totalnumber, MaxPerPage, CurrentPage, ShowTotal, ShowAllPages, strUnit, ShowMaxPerPage)
    Dim TotalPage, strTemp, strUrl, i
    If totalnumber = 0 Or MaxPerPage = 0 Or IsNull(MaxPerPage) Then
        ShowPage = ""
        Exit Function
    End If
    If totalnumber Mod MaxPerPage = 0 Then
        TotalPage = totalnumber \ MaxPerPage
    Else
        TotalPage = totalnumber \ MaxPerPage + 1
    End If
    If CurrentPage > TotalPage Then CurrentPage = TotalPage
    strTemp = "<div class=""pagin"">"
    If ShowTotal = True Then strTemp = strTemp & "<div class=""pageTips"">共<i class=""blue"">" & totalnumber & "</i>" & strUnit & "，当前显示第&nbsp;<i class=""blue"">" & CurrentPage & "&nbsp;</i>页</div>"
    
	strUrl = JoinChar(sfilename)
    If HR_CBool(ShowMaxPerPage) Then strUrl = JoinChar(sfilename) & "limit=" & MaxPerPage & "&"
	
	strTemp = strTemp & "<ul class=""paginList"">"
    If CurrentPage = 1 Then
        strTemp = strTemp & "<li class=""paginItem""><span class=""pagepre""><i class=""hr-icon"">&#xed10;</i></span></li>"
    Else
        strTemp = strTemp & "<li class=""paginItem""><a href=""" & strUrl & "page=" & (CurrentPage - 1) & """ title=""上一页"" class=""pagepre""><i class=""hr-icon"">&#xed10;</i></a></li>"
    End If

    If HR_CBool(ShowAllPages) Then
        Dim Jmaxpages
        If (CurrentPage - 4) <= 0 Or TotalPage < 6 Then
            Jmaxpages = 1
            Do While (Jmaxpages < 6)
                If Jmaxpages = CurrentPage Then
                    strTemp = strTemp & "<li class=""paginItem current""><span>" & Jmaxpages & "</span></li>"
				Else
					If strUrl <> "" Then strTemp = strTemp & "<li class=""paginItem""><a href=""" & strUrl & "page=" & Jmaxpages & """>" & Jmaxpages & "</a></li>"
				End If
				If Jmaxpages = TotalPage Then Exit Do
				Jmaxpages = Jmaxpages + 1
            Loop
			If TotalPage >= 6 Then
				If TotalPage > 6 Then strTemp = strTemp & "<li class=""paginItem""><span><i class=""hr-icon"">&#xee9b;</i></span></li>"
				strTemp = strTemp & "<li class=""paginItem""><a href=""" & strUrl & "page=" & TotalPage & """>" & TotalPage & "</a></li>"
            End If
        ElseIf (CurrentPage + 4) >= TotalPage Then
            Jmaxpages = TotalPage - 4
            strTemp = strTemp & "<li class=""paginItem""><a href=""" & strUrl & "page=1"">1</a></li>"
			If TotalPage > 6 Then strTemp = strTemp & "<li class=""paginItem""><span><i class=""hr-icon"">&#xee9b;</i></span></li>"
			Do While (Jmaxpages <= TotalPage)
                If Jmaxpages = CurrentPage Then
                    strTemp = strTemp & "<li class=""paginItem current""><span>" & Jmaxpages & "</span></li>"
                Else
                    If strUrl <> "" Then strTemp = strTemp & "<li class=""paginItem""><a href=""" & strUrl & "page=" & Jmaxpages & """>" & Jmaxpages & "</a></li>"
                End If
                Jmaxpages = Jmaxpages + 1
            Loop
        Else
            Jmaxpages = CurrentPage - 2
			strTemp = strTemp & "<li class=""paginItem""><a href=""" & strUrl & "page=1"">1</a></li>"
            strTemp = strTemp & "<li class=""paginItem""><span><i class=""hr-icon"">&#xee9b;</i></span></li>"
			Do While (Jmaxpages < CurrentPage + 3)
				If Jmaxpages = CurrentPage Then
					strTemp = strTemp & "<li class=""paginItem current""><a href=""javascript:void(0);"">" & Jmaxpages & "</a></li>"
				Else
					If strUrl <> "" Then strTemp = strTemp & "<li class=""paginItem""><a href=""" & strUrl & "page=" & Jmaxpages & """>" & Jmaxpages & "</a></li>"
				End If
				Jmaxpages = Jmaxpages + 1
            Loop
            strTemp = strTemp & "<li class=""paginItem""><a href=""javascript:void(0);""><i class=""hr-icon"">&#xee9b;</i></a></li>"
			strTemp = strTemp & "<li class=""paginItem""><a href=""" & strUrl & "page=" & TotalPage & """>" & TotalPage & "</a></li>"
        End If
    End If
    If CurrentPage >= TotalPage Then
		strTemp = strTemp & "<li class=""paginItem noPage""><span class=""pagenxt""><i class=""hr-icon"">&#xed11;</i></span></li>"
    Else
		strTemp = strTemp & "<li class=""paginItem""><a href=""" & strUrl & "page=" & (CurrentPage + 1) & """ class=""pagenxt""><i class=""hr-icon"">&#xed11;</i></a></li>"
    End If
	strTemp = strTemp & "</ul>"

	If HR_CBool(ShowMaxPerPage) Then
		'strTemp = strTemp & "<span class=""MaxPer""><input type=""text"" name=""MaxPerPage"" size=""3"" maxlength=""4"" value=""" & MaxPerPage & """ onKeyPress=""if (event.keyCode==13) window.location='" & JoinChar(sfilename) & "Page=" & CurrentPage & "&MaxPerPage=" & "'+this.value;"" />" & strUnit & "/页</span>"
		'strTemp = strTemp & "<span class=""Locat"">转到第<Input type=""text"" name=""Page"" size=""3"" maxlength=""5"" value=""" & CurrentPage & """ onKeyPress=""if (event.keyCode==13) window.location='" & strUrl & "Page=" & "'+this.value;"" />页</span>"
	Else
		'strTemp = strTemp & "<span class=""MaxPer""><b>" & MaxPerPage & "</b>" & strUnit & "/页</span>"
    End If

    strTemp = strTemp & "</div>"
    ShowPage = strTemp
End Function

'=====================================================================
'函数名：ShowPageMobile	作用：显示"上一页 下一页"等信息【触屏版】
'参  数：sFileName  ----链接地址		TotalNumber ----总数量
'        MaxPerPage  ----每页数量		CurrentPage ----当前页
'        ShowTotal   ----是否显示总数量
'        ShowAllPages ---是否用下拉列表显示所有页面以供跳转。
'        strUnit     ----计数单位		ShowMaxPerPage  ----是否显示每页信息量选项框
'返回值："上一页 下一页"等信息的HTML代码
'=====================================================================
Function ShowPageMobile(sfilename, totalnumber, MaxPerPage, CurrentPage, ShowTotal, ShowAllPages, strUnit, ShowMaxPerPage)
    Dim TotalPage, strFun, strUrl, i
    If totalnumber = 0 Or MaxPerPage = 0 Or IsNull(MaxPerPage) Then
        ShowPage = ""
        Exit Function
    End If
    If totalnumber Mod MaxPerPage = 0 Then
        TotalPage = totalnumber \ MaxPerPage
    Else
        TotalPage = totalnumber \ MaxPerPage + 1
    End If
    If CurrentPage > TotalPage Then CurrentPage = TotalPage
    strFun = "<div class=""ShowPage"">"
    If ShowTotal = True Then strFun = strFun & "<span class=""Total"">共<b>" & totalnumber & "</b>" & strUnit & "</span>"
    
	strUrl = JoinChar(sfilename)
    If HR_CBool(ShowMaxPerPage) Then strUrl = JoinChar(sfilename) & "MaxPerPage=" & MaxPerPage & "&"

    If CurrentPage = 1 Then
        strFun = strFun & "<span class=""Prev""><i class=""ion-ios-arrow-back""></i></span>"
    Else
        strFun = strFun & "<span class=""Prev""><a href=""" & strUrl & "Page=" & (CurrentPage - 1) & """ title=""上一页""><i class=""ion-ios-arrow-back""></i></a></span>"
    End If

    If HR_CBool(ShowAllPages) Then
        Dim Jmaxpages
        If (CurrentPage - 4) <= 0 Or TotalPage < 6 Then
            Jmaxpages = 1
            Do While (Jmaxpages < 6)
                If Jmaxpages = CurrentPage Then
                    strFun = strFun & "<span class=""Current"" title=""当前页""><i>" & Jmaxpages & "</i></span>"
				Else
					If strUrl <> "" Then strFun = strFun & "<span class=""CurrentA""><a href=""" & strUrl & "Page=" & Jmaxpages & """><i>" & Jmaxpages & "</i></a></span>"
				End If
				If Jmaxpages = TotalPage Then Exit Do
				Jmaxpages = Jmaxpages + 1
            Loop
			If TotalPage >= 6 Then
				If TotalPage > 6 Then strFun = strFun & "<span class=""More""><i>…</i></span>"
				strFun = strFun & "<span class=""Last""><a href=""" & strUrl & "Page=" & TotalPage & """><i>" & TotalPage & "</i></a></span>"
            End If
        ElseIf (CurrentPage + 4) >= TotalPage Then
            Jmaxpages = TotalPage - 4
            strFun = strFun & "<span class=""First""><a href=""" & strUrl & "Page=1""><i>1</i></a></span>"
			If TotalPage > 6 Then strFun = strFun & "<span class=""More""><i>…</i></span>"
			Do While (Jmaxpages <= TotalPage)
                If Jmaxpages = CurrentPage Then
                    strFun = strFun & "<span class=""Current"" title=""当前页""><i>" & Jmaxpages & "</i></span>"
                Else
                    If strUrl <> "" Then strFun = strFun & "<span class=""CurrentA""><a href=""" & strUrl & "Page=" & Jmaxpages & """><i>" & Jmaxpages & "</i></a></span>"
                End If
                Jmaxpages = Jmaxpages + 1
            Loop
        Else
            Jmaxpages = CurrentPage - 2
			strFun = strFun & "<span class=""First""><a href=""" & strUrl & "Page=1""><i>1</i></a></span>"
            strFun = strFun & "<span class=""More""><i>…</i></span>"
			Do While (Jmaxpages < CurrentPage + 3)
				If Jmaxpages = CurrentPage Then
					strFun = strFun & "<span class=""Current"" title=""当前页""><i>" & Jmaxpages & "</i></span>"
				Else
					If strUrl <> "" Then strFun = strFun & "<span class=""CurrentA""><a href=""" & strUrl & "Page=" & Jmaxpages & """><i>" & Jmaxpages & "</i></a></span>"
				End If
				Jmaxpages = Jmaxpages + 1
            Loop
            strFun = strFun & "<span class=""More""><i>…</i></span>"
			strFun = strFun & "<span class=""Last""><a href=""" & strUrl & "Page=" & TotalPage & """><i>" & TotalPage & "</i></a></span>"
        End If
    End If
	strFun = strFun & "<span class=""MaxPer""><b>" & CurrentPage & "</b>/<b>" & TotalPage & "</b></span>"
    If CurrentPage >= TotalPage Then
		strFun = strFun & "<span class=""Next""><i class=""ion-ios-arrow-forward""></i></span>"
    Else
        strFun = strFun & "<span class=""Next""><a href=""" & strUrl & "Page=" & (CurrentPage + 1) & """ title=""下一页""><i class=""ion-ios-arrow-forward""></i></a></span>"
    End If
	If HR_CBool(ShowMaxPerPage) Then
        'strFun = strFun & "<span class=""MaxPer""><input type=""text"" name=""MaxPerPage"" size=""3"" maxlength=""4"" value=""" & MaxPerPage & """ onKeyPress=""if (event.keyCode==13) window.location='" & JoinChar(sfilename) & "Page=" & CurrentPage & "&MaxPerPage=" & "'+this.value;"" />" & strUnit & "/页</span>"
		'strFun = strFun & "<span class=""Locat"">转到第<Input type=""text"" name=""Page"" size=""3"" maxlength=""5"" value=""" & CurrentPage & """ onKeyPress=""if (event.keyCode==13) window.location='" & strUrl & "Page=" & "'+this.value;"" />页</span>"
	Else
        'strFun = strFun & "<span class=""MaxPer""><b>" & MaxPerPage & "</b>" & strUnit & "/页</span>"
    End If

    strFun = strFun & "</div>"
    ShowPageMobile = strFun
End Function

Sub RecordFrontLog(sModule, sHTTP_REFERER, strLog, sPassed, sParam1)
	Dim rsLog
	Set rsLog = Server.CreateObject("Adodb.RecordSet")
		rsLog.Open("Select * From HR_LogFront"), Conn, 1, 3
		rsLog.AddNew
		rsLog("LogID") = GetNewID("HR_LogFront", "LogID")
		rsLog("ModuleID") = HR_Clng(sModule)
		rsLog("HttpReferer") = sHTTP_REFERER
		rsLog("LogContent") = strLog
		rsLog("Catalog") = HR_Clng(sPassed)
		rsLog("RecordTime") = Now()
		rsLog("UserName") = UserYGXM
		rsLog("TrueIP") = UserTrueIP
		rsLog.Update
		rsLog.Close
	Set rsLog = Nothing
End Sub

'******************* 以下为业绩考核专用 *******************

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
'函数名：ReplaceSQLChar		【替换SQL专用字符为全角】
'=====================================================================
Function ReplaceSQLChar(tSqlChar)
	Dim strFun
	If HR_IsNull(tSqlChar) Then Exit Function
	strFun = Replace(tSqlChar, "+", "＋")
	strFun = Replace(strFun, "'", "‘")
	strFun = Replace(strFun, "%", "％")
	strFun = Replace(strFun, "^", "’")
	strFun = Replace(strFun, "&", "＆")
	strFun = Replace(strFun, "?", "？")
	strFun = Replace(strFun, "(", "（")
	strFun = Replace(strFun, ")", "（")
	strFun = Replace(strFun, "<", "〈")
	strFun = Replace(strFun, ">", "〉")
	strFun = Replace(strFun, "[", "【")
	strFun = Replace(strFun, "]", "】")
	strFun = Replace(strFun, "{", "｛")
	strFun = Replace(strFun, "}", "｝")
	strFun = Replace(strFun, "/", "∕")
	strFun = Replace(strFun, "\", "＼")
	strFun = Replace(strFun, ";", "；")
	strFun = Replace(strFun, ":", "：")
	strFun = Replace(strFun, Chr(9), "")		'去除TAB键
	strFun = Replace(strFun, Chr(34), "")
	strFun = Replace(strFun, Chr(0), "")
	ReplaceSQLChar = strFun
End Function

'=====================================================================
'函数名：ConvertNumDate		【时间戳返回日期】
'=====================================================================
Function ConvertNumDate(timeStamp)
	If HR_Clng(timeStamp) > 10000 Then
		ConvertNumDate = DateAdd("d", timeStamp-2, "1900-01-01 00:00:00")		'减2调整时间差【因为PHPExcel未设置格式】
    End If	
End Function
Function ConvertDateToNum(fTime)		'将日期转为时间戳
	If HR_IsNull(fTime) = False Then
		If IsDate(fTime) Then ConvertDateToNum = DateDiff("d","1900-01-01 00:00:00", fTime)
	End If	
End Function

'=====================================================================
'函数名：GetItemOption		【返回项目下拉菜单项】
'=====================================================================
Function GetItemOption(fType, fItemID, fIsTempA)			'取项目下拉菜单项
	Dim rsFun, sqlFun, strFun, fItemName
	sqlFun = "Select * From HR_CLass Where ModuleID=1001"
	If HR_CBool(fIsTempA) Then sqlFun = sqlFun & " And Template='TempTableA'"
	If HR_CLng(fType) > 0 Then sqlFun = sqlFun & " And ClassType=" & HR_CLng(fType)
	sqlFun = sqlFun & " Order By ClassType, RootID, OrderID"
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			Do While Not rsFun.EOF
				fItemName = Trim(rsFun("ClassName"))
				strFun = strFun & "<option value=""" & rsFun("ClassID") & """"
				If HR_CLng(rsFun("Depth")) > 0 Then
					fItemName = "　" & Trim(rsFun("ClassName"))
				End If
				If HR_CLng(rsFun("ClassID")) = HR_CLng(fItemID) Then strFun = strFun & " selected"
				If HR_CLng(rsFun("Child")) > 0 Then strFun = strFun & " disabled"
				strFun = strFun & ">" & fItemName & "</option>"
				rsFun.MoveNext
			Loop
		End If
	Set rsFun = Nothing
	GetItemOption = strFun
End Function

'=====================================================================
'函数名：GetTemplateOption		【返回模板下拉】
'=====================================================================
Function GetTemplateOption(fTypeID, fTemp)
	Dim rsFun, iFun, arrTemplate, strFun
	arrTemplate = Split(XmlText("Common", "Template", ""), "|")
	For iFun = 0 To Ubound(arrTemplate)
		strFun = strFun & "<option value=""" & arrTemplate(iFun) & """"
		If Trim(fTemp) = arrTemplate(iFun) Then strFun = strFun & " selected"
		strFun = strFun & ">" & arrTemplate(iFun) & "</option>"
	Next
	GetTemplateOption = strFun
End Function

'=====================================================================
'函数名：GetPeriodOption		【取节次下拉】
'=====================================================================
Function GetPeriodOption(fCampus, fPeriod, fType)
	Dim sqlFun, rsFun, strFun : strFun = ""
	sqlFun = "Select * From HR_Period Where Campus='" & Trim(fCampus) & "'"
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			Do While Not rsFun.EOF
				strFun = strFun & "<option value=""" & rsFun("Period") & """ name=""" & rsFun("PeriodID") & """"
				If Trim(fPeriod) = rsFun("Period") Then strFun = strFun & " selected"
				strFun = strFun & ">"
				strFun = strFun & rsFun("Period") & "</option>"
				rsFun.MoveNext
			Loop
		End If
	Set rsFun = Nothing
	GetPeriodOption = strFun
End Function

'=====================================================================
'函数名：GetPeriodTime		【根据节次返回时间】
'=====================================================================
Function GetPeriodTime(fCampus, fPeriod, fType)
	Dim strFun, rsFun, fArr, strArr, stTime, enTime
	fCampus = Trim(fCampus) : fPeriod = Trim(fPeriod)
	If fCampus <> "" And fPeriod <> "" Then
		If HR_Clng(fPeriod) > 0 Then
			Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fPeriod))
				If Not(rsFun.BOF And rsFun.EOF) Then
					strFun = Trim(rsFun("StartTime")) & " - " & Trim(rsFun("EndTime"))
				End If
			Set rsFun = Nothing
		ElseIf Instr(fPeriod, "-") Then
			fArr = Split(fPeriod, "-")
			If Ubound(fArr) = 1 Then
				Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fArr(0)))
					If Not(rsFun.BOF And rsFun.EOF) Then
						stTime = Trim(rsFun("StartTime"))
					End If
				Set rsFun = Nothing
				Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fArr(1)))
					If Not(rsFun.BOF And rsFun.EOF) Then
						enTime = Trim(rsFun("EndTime"))
					End If
				Set rsFun = Nothing
				strFun = Trim(stTime) & " - " & Trim(enTime)
			End If
		End If
	End If
	GetPeriodTime = strFun
End Function

'=====================================================================
'函数名：GetSubmoduleOption		【取级别下拉】
'=====================================================================
Function GetSubmoduleOption(fFieldID, strField)
	Dim sqlFun, rsFun, strFun
	sqlFun = "Select * From HR_ItemModel Where ClassID=" & HR_Clng(fFieldID)
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			strFun = ""
			Do While Not rsFun.EOF
				strFun = strFun & "<option value=""" & rsFun("FieldName") & """ name=""" & rsFun("ClassID") & """"
				If Trim(strField) = rsFun("FieldName") Then strFun = strFun & " selected"
				strFun = strFun & ">"
				strFun = strFun & rsFun("FieldName") & "</option>"
				rsFun.MoveNext
			Loop
		End If
	Set rsFun = Nothing
	GetSubmoduleOption = strFun
End Function
Function GetSubmoduleSelect(fFieldID, strField)
	Dim sqlFun, rsFun, strFun, iFun
	sqlFun = "Select * From HR_ItemModel Where ClassID=" & HR_Clng(fFieldID)
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			strFun = "" : iFun = 0
			Do While Not rsFun.EOF
				If iFun > 0 Then strFun = strFun & ","
				strFun = strFun & """" & rsFun("FieldName") & """"
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		End If
	Set rsFun = Nothing
	GetSubmoduleSelect = strFun
End Function

'=====================================================================
'函数名：GetItemGradeOption		【取等级下拉】
'=====================================================================
Function GetItemGradeOption(fItemID, fLevel, strField)
	Dim sqlFun, rsFun, strFun, fLevelID
	Set rsFun = Conn.Execute("Select Top 1 ID From HR_ItemModel Where ClassID=" & HR_Clng(fItemID) & " And FieldName='" & Trim(fLevel) & "'")
		fLevelID = HR_Clng(rsFun(0))
	Set rsFun = Nothing
	sqlFun = "Select * From HR_ItemGrade Where LevelID=" & HR_Clng(fLevelID)
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			strFun = ""
			Do While Not rsFun.EOF
				strFun = strFun & "<option value=""" & rsFun("Grade") & """ name=""" & rsFun("ID") & """"
				If Trim(strField) = rsFun("Grade") Then strFun = strFun & " selected"
				strFun = strFun & ">"
				strFun = strFun & rsFun("Grade") & "</option>"
				rsFun.MoveNext
			Loop
		End If
	Set rsFun = Nothing
	GetItemGradeOption = strFun
End Function
Function GetItemGradeSelect(fItemID, fLevel, strField)
	Dim sqlFun, rsFun, strFun, fLevelID, iFun
	Set rsFun = Conn.Execute("Select Top 1 ID From HR_ItemModel Where ClassID=" & HR_Clng(fItemID) & " And FieldName='" & Trim(fLevel) & "'")
		fLevelID = HR_Clng(rsFun(0))
	Set rsFun = Nothing
	sqlFun = "Select * From HR_ItemGrade Where LevelID=" & HR_Clng(fLevelID)
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			strFun = "" : iFun = 0
			Do While Not rsFun.EOF
				If iFun > 0 Then strFun = strFun & ","
				strFun = strFun & """" & rsFun("Grade") & """"
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		End If
	Set rsFun = Nothing
	GetItemGradeSelect = strFun
End Function

'=====================================================================
'函数名：GetFieldOption		【取指定字段下拉】
'=====================================================================
Function GetFieldOption(fSheetName, fField, fValue)
	on error resume next
	Dim rsFun, strFun : strFun = ""
	If isNull(fSheetName) Or Instr(fSheetName, "R_") = 0 Then Exit Function
	If isNull(fField) Or Trim(fField) = "" Then Exit Function
	Set rsFun = Conn.Execute("Select " & Trim(fField) & " From " & fSheetName & " Group By " & Trim(fField))
		If Not Err.Number=0 Then Err.Clear : Exit Function
		If Not(rsFun.BOF And rsFun.EOF) Then
			Do While Not rsFun.EOF
				strFun = strFun & "<option value=""" & rsFun(fField) & """"
				If Trim(fValue) = Trim(rsFun(fField)) Then strFun = strFun & " selected"
				strFun = strFun & ">"
				strFun = strFun & rsFun(fField) & "</option>"
				rsFun.MoveNext
			Loop
		End If
	Set rsFun = Nothing
	GetFieldOption = strFun
End Function
Function GetFieldSelect(fSheetName, fField, fValue)
	on error resume next
	Dim rsFun, strFun : strFun = ""
	If isNull(fSheetName) Or Instr(fSheetName, "R_") = 0 Then Exit Function
	If isNull(fField) Or Trim(fField) = "" Then Exit Function
	Set rsFun = Conn.Execute("Select " & Trim(fField) & " From " & fSheetName & " Group By " & Trim(fField))
		If Not Err.Number=0 Then Err.Clear : Exit Function
		If Not(rsFun.BOF And rsFun.EOF) Then
			iFun = 0
			Do While Not rsFun.EOF
				If iFun > 0 Then strFun = strFun & ","
				strFun = strFun & """" & rsFun(fField) & """"
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		End If
	Set rsFun = Nothing
	GetFieldSelect = strFun
End Function


'=====================================================================
'函数名：BackSQLData		【备份数据库，返回备份文件】
'=====================================================================
Function BackSQLData(fType, fPath)
	Dim fTime : fTime = DateAdd("h", -8, Now()) : fTime = DateDiff("s","1970-01-01 00:00:00", fTime)
	Dim fBackPath : fBackPath = Server.MapPath("/BackData/")
	Dim fBackFile : fBackFile = "\HR-SQL" & fTime & GetRndPassword(4) & ".dat"
	'on error resume next
	Conn.Execute("backup database " & SqlDatabaseName & " to disk='" & fBackPath & fBackFile &"'")
	If Not Err.Number=0 Then Response.Write Err.Descripting : Response.End
	BackSQLData = "/BackData/" & Replace(fBackFile, "\", "")
End Function

'=====================================================================
'函数名：GetStudentType		【获取学生分类ID】
'=====================================================================
Function GetStudentType(strType)
	Dim fArrStuType : fArrStuType = Split(XmlText("Common", "StudentType", ""), "|")
	Dim iFun : GetStudentType = 0
	For iFun = 0 To Ubound(fArrStuType)
		If Trim(strType) = fArrStuType(iFun) Then GetStudentType = iFun + 1
	Next
End Function
Function GetStudentOption(strType)
	Dim fArrStuType : fArrStuType = Split(XmlText("Common", "StudentType", ""), "|")
	Dim iFun, strFun
	For iFun = 0 To Ubound(fArrStuType)
		strFun = strFun & "<option value=""" & fArrStuType(iFun) & """"
		If Trim(strType) = fArrStuType(iFun) Then strFun = strFun & " selected"
		strFun = strFun & ">" & fArrStuType(iFun) & "</option>"
	Next
	GetStudentOption = strFun
End Function

'=====================================================================
'函数名：GetSemesterOption		【取学期/学年下拉】
'=====================================================================
Function GetSemesterOption(fType, fVal)
	Dim strFun, iFun, sYear, eYear
	eYear = Year(Date())
	If Month(Date()) > 6 Then eYear = eYear + 1
	sYear = eYear-3
	For iFun = eYear To sYear Step -1
		If iFun > sYear Then
			If HR_Clng(fType) <> 2 Then
				strFun = strFun & "<option value=""" & iFun-1 & "-" & iFun & """"
				If Trim(fVal) = iFun-1 & "-" & iFun Then strFun = strFun & " selected"
				strFun = strFun & ">" & iFun-1 & "-" & iFun & "</option>"
			End If
			strFun = strFun & "<option value=""" & iFun-1 & "-" & iFun & "-1"""
			If Trim(fVal) = iFun-1 & "-" & iFun & "-1" Then strFun = strFun & " selected"
			strFun = strFun & ">" & iFun-1 & "-" & iFun & "-1</option>"

			strFun = strFun & "<option value=""" & iFun-1 & "-" & iFun & "-2"""
			If Trim(fVal) = iFun-1 & "-" & iFun & "-2" Then strFun = strFun & " selected"
			strFun = strFun & ">" & iFun-1 & "-" & iFun & "-2</option>"
		End If
	Next
	GetSemesterOption = strFun
End Function

'=====================================================================
'函数名：GetItemClassID		【返回取考核项目ID数组串】
'=====================================================================
Function GetItemClassID(fLimit)
	Dim rsFun, sqlFun, iFun, strFun : strFun = ""
	sqlFun = "Select ClassID From HR_Class Where ModuleID=1001 And Child=0"
	If HR_IsNull(fLimit) = False Then sqlFun = sqlFun & " " & fLimit
	Set rsFun = Server.CreateObject("ADODB.RecordSet")
		rsFun.Open(sqlFun), Conn, 1, 1
		If Not(rsFun.BOF And rsFun.EOF) Then
			iFun = 0
			Do While Not rsFun.EOF
				If iFun > 0 Then strFun = strFun & ","
				strFun = strFun & rsFun(0)
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		End If
	Set rsFun = Nothing
	GetItemClassID = strFun
End Function

'=====================================================================
'函数名：GetStatisTableField		【返回汇总表中的统计字段字符组串】
'返回值：string
'=====================================================================
Function GetStatisTableField()		'取统计表字段
	Dim rsFun, sqlFun, strFun, iFun, fStuType, fArrStuType
	sqlFun = "Select * From HR_Class Order By ClassType,RootID,OrderID"
	Set rsFun = Conn.Execute(sqlFun)
		Do While Not rsFun.EOF
			If rsFun("Child") = 0 Then	'有子类跳过
				fStuType = FilterArrNull(Trim(rsFun("StudentType")), ",")
				If HR_IsNull(fStuType) = False Then
					fArrStuType = Split(fStuType, ",")
					For iFun = 0 To Ubound(fArrStuType)
						strFun = strFun & "F" & rsFun("ClassID") & "_" & GetStudentType(fArrStuType(iFun)) & "||"
					Next
				Else
					strFun = strFun & "F" & rsFun("ClassID") & "||"
				End If
			End If
			rsFun.MoveNext
		Loop
	Set rsFun = Nothing
	strFun = FilterArrNull(strFun, "||")
	GetStatisTableField = strFun
End Function

'=====================================================================
'函数名：GetItemCourseOption		【返回指定项目课程下拉】
'=====================================================================
Function GetItemCourseOption(fItem, fCourse, fYGDM, tParam)
	Dim rsFun, sqlFun, strFun, iFun, fSheetName, fDate
	fSheetName = "HR_Sheet_" & fItem
	If ChkTable(fSheetName) Then
		sqlFun = "Select * From " & fSheetName & " Where scYear=" & DefYear
		If HR_CLng(fYGDM) > 0 Then sqlFun = sqlFun & " And VA1=" & HR_CLng(fYGDM)
		sqlFun = sqlFun & " Order By VA4 DESC"
		Set rsFun = Conn.Execute(sqlFun)
			If Not(rsFun.BOF And rsFun.EOF) Then
				Do While Not rsFun.EOF
					fDate = FormatDate(ConvertNumDate(rsFun("VA4")), 2)
					strFun = strFun & "<option value=""" & rsFun("ID") & """"
					If HR_CLng(rsFun("ID")) = HR_CLng(fCourse) Then strFun = strFun & " selected"
					strFun = strFun & ">" & rsFun("VA8") & "_" & fDate & "</option>"
					rsFun.MoveNext
				Loop
			End If
		Set rsFun = Nothing

	End If
	GetItemCourseOption = strFun
End Function
'=====================================================================
'函数名：GetCourseOption		【返回课程名称下拉】
'=====================================================================
Function GetCourseOption(fCourse, fType)
	Dim rsFun, iFun, arrCourse, strFun
	arrCourse = Split(XmlText("Common", "Course", ""), "|")
	For iFun = 0 To Ubound(arrCourse)
		strFun = strFun & "<option value=""" & arrCourse(iFun) & """"
		If Trim(fCourse) = Trim(arrCourse(iFun)) Then strFun = strFun & " selected"
		strFun = strFun & ">" & arrCourse(iFun) & "</option>"
	Next
	GetCourseOption = strFun
End Function

'=====================================================================
'函数名：GetClassroomOption		【返回授课地点下拉】
'=====================================================================
Function GetClassroomOption(fValue, fType)
	Dim rsFun, iFun, arrCourse, strFun
	arrCourse = Split(XmlText("Common", "Classroom", ""), "|")
	For iFun = 0 To Ubound(arrCourse)
		strFun = strFun & "<option value=""" & arrCourse(iFun) & """"
		If Trim(fValue) = Trim(arrCourse(iFun)) Then strFun = strFun & " selected"
		strFun = strFun & ">" & arrCourse(iFun) & "</option>"
	Next
	GetClassroomOption = strFun
End Function

'=====================================================================
'函数名：GetDepartmentOption		【取科室下拉】
'=====================================================================
Function GetDepartmentOption(fParentID, fDepartID, isRoot)
	Dim strFun, rsFun, sqlFun, fOrder, tDeptName
	sqlFun = "Select * From HR_Department Where RootID>0"
	If HR_CBool(isRoot) Then sqlFun = sqlFun & " And ParentID=0"
	If HR_Clng(fParentID) > 0 Then sqlFun = sqlFun & " And SJKS=" & HR_Clng(fParentID)
	sqlFun = sqlFun & " Order By RootID, OrderID ASC"
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			strFun = ""
			Do While Not rsFun.EOF
				tDeptName = Trim(rsFun("KSMC"))
				If HR_Clng(rsFun("ParentID")) > 0 Then
					'Set rsTmp = Conn.Execute("Select Max(OrderID) From HR_Department Where RootID=" & rsFun("RootID") & " And SJKS=" & HR_Clng(rsFun("SJKS")))
					'	fOrder = HR_Clng(rsTmp(0))
					'Set rsTmp = Nothing
					tDeptName = " " & tDeptName
				End If

				strFun = strFun & "<option value=""" & rsFun("KSDM") & """"
				If HR_Clng(fDepartID) = rsFun("KSDM") Then strFun = strFun & " selected"
				strFun = strFun & ">"
				strFun = strFun & "" & tDeptName & "</option>"
				rsFun.MoveNext
			Loop
		End If
	Set rsFun = Nothing
	GetDepartmentOption = strFun
End Function
Function GetDeptOption(fParentID, fDepart, isRoot)
	Dim strFun, rsFun, sqlFun, fOrder, tDeptName
	sqlFun = "Select * From HR_Department Where DepartmentID>0"
	If HR_CBool(isRoot) Then sqlFun = sqlFun & " And SJKS=KSDM"
	If HR_Clng(fParentID) > 0 Then sqlFun = sqlFun & " And SJKS=" & HR_Clng(fParentID)
	sqlFun = sqlFun & " Order By PXBH ASC"
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			strFun = ""
			Do While Not rsFun.EOF
				tDeptName = Trim(rsFun("KSMC"))
				strFun = strFun & "<option value=""" & rsFun("KSDM") & """"
				If HR_Clng(fDepart) = rsFun("KSDM") Then strFun = strFun & " selected"
				strFun = strFun & ">"
				strFun = strFun & "" & tDeptName & "</option>"
				rsFun.MoveNext
			Loop
		End If
	Set rsFun = Nothing
	GetDeptOption = strFun
End Function
Function GetTeacherDeptName(fYGDM)
	Dim strFun, rsFun : fYGDM = Trim(ReplaceBadChar(fYGDM))
	If HR_IsNull(fYGDM) Then Exit Function
	Set rsFun = Conn.Execute("Select Top 1 b.KSMC From HR_Teacher a Inner Join HR_Department b On a.KSDM=b.KSDM Where a.YGDM='" & fYGDM & "'")
		If Not(rsFun.BOF And rsFun.EOF) Then
			strFun = Trim(rsFun("KSMC"))
		End If
	Set rsFun = Nothing
	GetTeacherDeptName = strFun
End Function

'=====================================================================
'函数名：GetCampusOption		【取校(院)区下拉】
'=====================================================================
Function GetCampusOption(fCampus, fType)
	Dim strFun, iFun, fArrCampus : fArrCampus = Split(XmlText("Common", "Campus", ""), "|")
	For iFun = 0 To Ubound(fArrCampus)
		strFun = strFun & "<option value=""" & fArrCampus(iFun) & """"
		If Trim(fCampus) = Trim(fArrCampus(iFun)) Then strFun = strFun & " selected"
		strFun = strFun & ">" & fArrCampus(iFun) & "</option>"
	Next
	GetCampusOption = strFun
End Function

'=====================================================================
'函数名：GetClassOption		【取分类下拉】
'=====================================================================
Function GetClassOption(fModuleID, fClassID, fType)
	Dim rsFun, strFun, fParentID
	fParentID = 0
	Set rsFun = Conn.Execute("Select * From HR_Class Where ModuleID=" & fModuleID & " And ParentID=" & fParentID & " Order By ClassType ASC, RootID, OrderID")
		If Not(rsFun.BOF And rsFun.EOF) Then
			strFun = ""
			Do While Not rsFun.EOF
				strFun = strFun & "<option value=""" & rsFun("ClassID") & """"
				If HR_Clng(fClassID) = rsFun("ClassID") Then strFun = strFun & " selected"
				strFun = strFun & ">"
				If fParentID > 0 Then
					If rsFun("NextID") > 0 Then
						strFun = strFun & "├&nbsp;"
					Else
						strFun = strFun & "└&nbsp;"
					End If
				End If
				strFun = strFun & "" & rsFun("ClassName") & "</option>"
				rsFun.MoveNext
			Loop
		End If
	Set rsFun = Nothing
	GetClassOption = strFun
End Function

'=====================================================================
'函数名：GetYearOption		【取学年下拉】
'=====================================================================
Function GetYearOption(fType, fVal)
	Dim strFun, iFun, fYear : fYear = HR_Clng(Year(Date()))
	If Month(Date()) > 6 Then fYear = fYear + 1
	For iFun = fYear To fYear-10 Step -1
		strFun = strFun & "<option value=""" & iFun & """"
		If HR_Clng(fVal) = iFun Then strFun = strFun & " selected"
		strFun = strFun & ">" & iFun-1 & "-" & iFun & "年度</option>"
	Next
	GetYearOption = strFun
End Function

'=====================================================================
'函数名：GetSchoolYear		【取学年】fType:1（日期）,2:学年/学期,3:取学期（1上学期，2下学期）
'=====================================================================
Function GetSchoolYear(fDate, fType)
	Dim fArr, funYear : funYear = 0
	If HR_IsNull(fDate) = False Then
		If IsDate(fDate) Then			'传入为日期
			If HR_Clng(fType) = 3 Then	'取学期
				funYear = 2
				If Month(fDate)>6 Then funYear = 1
			Else
				funYear = HR_Clng(Year(fDate))
				If Month(fDate)>6 Then funYear = funYear + 1
			End If
		Else
			fArr = Split(fDate, "-")
			If Ubound(fArr) > 0 Then
				If HR_Clng(HR_Clng(fArr(1))) > 1970 And HR_Clng(HR_Clng(fArr(1))) < 2030 Then funYear = HR_Clng(fArr(1))
			End If
			If Ubound(fArr) = 2 And HR_Clng(fType) = 3 Then funYear = HR_Clng(fArr(2))		'取学期
		End If
	End If
	GetSchoolYear = HR_CLng(funYear)
End Function

'=====================================================================
'函数名：GetAttachIcon		【取附件图标】
'=====================================================================
Function GetAttachIcon(fExtname)
	Dim fArrExtname, fArrIcon, iFun
	GetAttachIcon = "&#xec15;"
	fArrExtname = Split("jpg,jpeg,png,bmp,gif,xls,xlsx,pdf,doc,docx,txt,rar,zip",",")
	fArrIcon = Split("&#xf1c5;,&#xf1c5;,&#xf1c5;,&#xf1c5;,&#xf1c5;,&#xf1c3;,&#xf1c3;,&#xf1c1;,&#xf1c2;,&#xf1c2;,&#xf0f6;,&#xec1c;,&#xec1c;",",")
	If HR_IsNull(fExtname) = False Then
		For iFun = 0 To Ubound(fArrExtname)
			If Trim(LCase(fExtname)) = Trim(fArrExtname(iFun)) Then
				GetAttachIcon = Trim(fArrIcon(iFun))
			End If
		Next
	End If
End Function
Function GetCountAttach(fFiles)		'取附件总数
	GetCountAttach = 0
	If HR_IsNull(fFiles) = False Then
		fFiles = FilterArrNull(fFiles, "|")
		GetCountAttach = Ubound(Split(fFiles, "|")) + 1
	End If
End Function

'=====================================================================
'函数名：SendMessage		【发送消息】
'fType：0事务通知|1申请修改|2退回|3确认|4审核
'=====================================================================
Function SendMessage(fType, fItemID, fCourseID, fReceiverID, fTitle, fMessage, fViewUrl)
	Dim rsFun
	SendMessage = False
	If HR_IsNull(fMessage) Then Exit Function
	Set rsFun = Server.CreateObject("ADODB.RecordSet")
		rsFun.Open("Select * From HR_Message"), Conn, 1, 3
		rsFun.AddNew
		rsFun("ID") = GetNewID("HR_Message", "ID")
		rsFun("MsgType") = HR_CLng(fType)
		rsFun("ItemID") = HR_CLng(fItemID)
		rsFun("CourseID") = HR_CLng(fCourseID)
		rsFun("ReceiverID") = HR_CLng(fReceiverID)
		rsFun("IsSend") = HR_False
		rsFun("Title") = Trim(fTitle)
		rsFun("Message") = Trim(fMessage)
		rsFun("SenderID") = HR_CLng(UserYGDM)
		rsFun("SendTime") = Now()
		rsFun("isRead") = False
		rsFun("ViewUrl") = Trim(fViewUrl)
		rsFun.Update
	Set rsFun = Nothing
	SendMessage = True
End Function

'=====================================================================
'函数名：FormatAPIDate		【格式化接口中的日期格式】
'=====================================================================
Function FormatAPIDate(fStrDate, fType)
	Dim strFun : strFun = Null
	If HR_IsNull(fStrDate) = False Then
		fStrDate = Replace(fStrDate, "T00:00:00", "")
		If fStrDate <> "" And IsDate(fStrDate) Then strFun = fStrDate
	End If
	FormatAPIDate = strFun
End Function

'=====================================================================
'函数名：GetModelIncItem()		【获取模型对应的项目ID】
'=====================================================================
Function GetModelIncItem(fModelName)
	Dim strFun, rsFun, iFun
	Set rsFun = Server.CreateObject("ADODB.RecordSet")
		rsFun.Open("Select ClassID From HR_Class Where Template='" & Trim(fModelName) & "'"), Conn, 1, 1
		If Not(rsFun.BOF And rsFun.EOF) Then
			iFun = 0
			Do While Not rsFun.EOF
				If iFun > 0 Then strFun = strFun & ","
				strFun = strFun & rsFun(0)
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		End If
	Set rsFun = Nothing
	GetModelIncItem = strFun
End Function

'=====================================================================
'函数名：GetAllManagerYGDM()		【获取所有管理员员工代码】
'返回值：String	(***|***|***)
'=====================================================================
Function GetAllManagerYGDM()
	Dim strFun, rsFun, iFun
	Set rsFun = Conn.Execute("Select * From HR_User Where ManageRank>0")
		If Not(rsFun.BOF And rsFun.EOF) Then
			iFun = 0
			Do While Not rsFun.EOF
				If iFun > 0 Then strFun = strFun & "|"
				strFun = strFun & rsFun("YGDM")
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		End If
	Set rsFun = Nothing
	GetAllManagerYGDM = strFun
End Function

'======== 发送文本卡片消息 ========
Function SentWechatMSG_QYCard(sTouser, sTitle, sURL, sContent)		'发送文本卡片消息
	Dim postJson, strSub : strSub = ""
	If HR_IsNull(sTouser) = False And HR_IsNull(sTitle) = False And HR_IsNull(sURL) = False And HR_IsNull(sContent) = False Then
		sContent = Replace(sContent, """", "\""")
		postJson = "{""touser"":""" & sTouser & """,""msgtype"":""textcard"",""agentid"":" & boAgentId & ",""textcard"":{""title"":""" & sTitle & """,""description"":""" & sContent & """,""url"":""" & sURL & """,""btntxt"":""查看详情""}}"
		strSub = PostWechatMessageQY(postJson, 1)
	End If
	SentWechatMSG_QYCard = strSub
End Function
'=====================================================================
'函数名：PostWechatMessageQY()	【企业微信发送会员消息】信息播报
'返回值：String
'=====================================================================
Function PostWechatMessageQY(fPostJson, fPostType)
	Dim strFun, fPostHttp, fPostUrl
	If Not(ChkTokenBobao) Then		'判断信息播报Access Token
		PostWechatMessageQY = "{""errcode"":500, ""errmsg"":""Access Token 已过期"", ""invaliduser"":""""}"
		Exit Function
	End If
	fPostUrl = "https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=" & boToken
	Set fPostHttp = CreateObject("Msxml2.ServerXMLHTTP")
		With fPostHttp
			.Open "Post", fPostUrl, False
			.setRequestHeader "Content-Type","application/xml;charset=UTF-8"
			.Send fPostJson
			strFun = .ResponseText
		End With
	Set fPostHttp = Nothing
	PostWechatMessageQY = strFun
End Function
%>
