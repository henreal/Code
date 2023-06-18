<%
'======== 触屏版公用部分【顶部、底部、导航等】 ========
Dim ParmPath : ParmPath = InstallDir & "Touch/"
Dim HeadFace

'----- 获取超管UserID数组arrManager()
Dim rsManager
Set rsManager = Server.CreateObject("ADODB.RecordSet")
	rsManager.Open("Select * From HR_User Where ManageRank=2"), Conn, 1, 1
	Redim arrManager(rsManager.RecordCount)
	If Not(rsManager.BOF And rsManager.EOF) Then
		k = 0
		Do While Not rsManager.EOF
			arrManager(k) = rsManager("UserID")
			rsManager.MoveNext
			k = k + 1
		Loop
	End If
Set rsManager = Nothing

Function getPageHead(iType)
	Dim strFun
	strFun = "<!DOCTYPE html>" & vbCrlf
	strFun = strFun & "<html lang=""zh-cn"">" & vbCrlf
	strFun = strFun & "<head>" & vbCrlf
	strFun = strFun & "	<meta charset=""utf-8"">" & vbCrlf
	strFun = strFun & "	<title>[@Site_Title]</title>" & vbCrlf
	strFun = strFun & "	<meta name=""renderer"" content=""webkit"">" & vbCrlf
	strFun = strFun & "	<meta http-equiv=""X-UA-Compatible"" content=""IE=edge,chrome=1"">" & vbCrlf
	strFun = strFun & "	<meta name=""viewport"" content=""width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=0, minimal-ui"">" & vbCrlf
	strFun = strFun & "	<link type=""text/css"" href=""[@Web_Dir]Static/weui/lib/weui.min.css?v=1.1.3"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strFun = strFun & "	<link type=""text/css"" href=""[@Web_Dir]Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strFun = strFun & "	<link type=""text/css"" href=""[@Web_Dir]Static/css/hr.common.css?v=1.0.3"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strFun = strFun & "	<link type=""text/css"" href=""[@Web_Dir]Static/css/touch.common.css?v=1.0.1&r=" & GetRndString(5) & """ rel=""stylesheet"" media=""all"">" & vbCrlf
	strFun = strFun & "	<!--[if lt IE 9]>" & vbCrlf
	strFun = strFun & "		<script src=""https://cdn.staticfile.org/html5shiv/r29/html5.min.js""></script>" & vbCrlf
	strFun = strFun & "		<script src=""https://cdn.staticfile.org/respond.js/1.4.2/respond.min.js""></script>" & vbCrlf
	strFun = strFun & "	<![endif]-->" & vbCrlf
	strFun = strFun & "	<script type=""text/javascript"" src=""[@Web_Dir]Static/js/jquery.min.js?v=1.11.2""></script>" & vbCrlf
	strFun = strFun & "	<script type=""text/javascript"" src=""[@Web_Dir]Static/weui/js/jquery-weui.min.js?v=1.2.1""></script>" & vbCrlf

	strFun = strFun & "	[@HeadStyle]" & vbCrlf
	strFun = strFun & "	[@HeadScript]" & vbCrlf
	strFun = strFun & "</head>" & vbCrlf
	strFun = strFun & "<body>" & vbCrlf
	getPageHead = strFun
End Function

Function getPageFoot(iType)
	Dim strFun
	strFun = "</body>" & vbCrlf
	strFun = strFun & "</html>" & vbCrlf
	strFun = strFun & "[@FootScript]" & vbCrlf
	getPageFoot = strFun
End Function

Function getHeadNav(iType)
	Dim strFun
	strFun = strFun & "<header class=""hr-rows hr-header"">" & vbCrlf
	strFun = strFun & "	<nav class=""navBack""><em><i class=""hr-icon"">&#xec58;</i></em></nav>" & vbCrlf
	strFun = strFun & "	<nav class=""navTitle""><span>" & SiteTitle & "</span></nav>" & vbCrlf
	strFun = strFun & "	<nav class=""navMenu""><em><i class=""hr-icon"">&#xeef7;</i></em></nav>" & vbCrlf
	strFun = strFun & "</header>" & vbCrlf
	strFun = strFun & "<div class=""nctouch-nav-layout layerNav"" style=""display: none;"">" & vbCrlf
	strFun = strFun & "	<div class=""nctouch-nav-menu"">" & vbCrlf
	strFun = strFun & "		<span class=""arrow""></span>" & vbCrlf
	strFun = strFun & "		<ul><li><a href=""" & ParmPath & "Index.html""><i class=""hr-icon hr-icon-top"">&#xebf1;</i>首　页</a></li>" & vbCrlf
    'strFun = strFun & "			<li class=""addNew""><a href=""javascript:void(0)""><i class=""hr-icon hr-icon-top"">&#xf3c0;</i>添加课程</a><sup></sup></li>" & vbCrlf
	strFun = strFun & "			<li><a href=""" & ParmPath & "Achieve/Index.html?A=List""><i class=""hr-icon hr-icon-top"">&#xec8d;</i>查看业绩<sup></sup></a></li>" & vbCrlf
	strFun = strFun & "			<li><a href=""" & ParmPath & "myCenter/Message.html""><i class=""hr-icon hr-icon-top"">&#xeea0;</i>我的消息<sup></sup></a></li>" & vbCrlf
	If UserRank > 0 Then strFun = strFun & "			<li><a href=""" & ParmPath & "Manage/Index.html""><i class=""hr-icon hr-icon-top"">&#xe948;</i>管理面板<sup></sup></a></li>" & vbCrlf
	strFun = strFun & "			<li><a href=""" & ParmPath & "Login/Logout.html?noBind=1""><i class=""hr-icon hr-icon-top"">&#xeca7;</i>退出登陆</a><sup></sup></li>" & vbCrlf
	strFun = strFun & "		</ul>" & vbCrlf
	strFun = strFun & "	</div>" & vbCrlf
	strFun = strFun & "</div>" & vbCrlf
	getHeadNav = strFun
End Function

Function GetErrBody(fType)
	Dim funHtml, funStr
	If ErrMsg = "" Then ErrMsg = "非常遗憾，您访问的页面不存在！"
	funStr = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	funStr = funStr & "		body {background-color:#f1f1f1;} .error{margin:0;width:100%;}" & vbCrlf
	funStr = funStr & "		.hr-err-box {position:fixed;left:20px;top:15%;right:20px;}" & vbCrlf
	funStr = funStr & "		.hr-err-box .error {width:100%;padding:20px;box-sizing:border-box;border:1px solid #ccc;background-color:rgba(255,255,255,0.8);}" & vbCrlf
	funStr = funStr & "		.error .errorInfo {flex-grow:2;padding: 0 10px;font-size:1rem;} .error .errorInfo h2 {font-size:1.3rem;}" & vbCrlf
	funStr = funStr & "		.error .errorIcon i {font-size:4rem;color:#f60;}" & vbCrlf
	funStr = funStr & "		.reindex {padding-top:20px;}" & vbCrlf
	funStr = funStr & "	</style>" & vbCrlf
	funHtml = getPageHead(1)
	funHtml = Replace(funHtml, "[@HeadStyle]", funStr)
	funHtml = Replace(funHtml, "[@HeadScript]", "")
	funHtml = funHtml & "<div class=""hr-err-box"">" & vbCrlf
	funHtml = funHtml & "	<div class=""hr-rows error"">" & vbCrlf
	funHtml = funHtml & "		<div class=""errorIcon""><i class=""hr-icon"">&#xe811;</i></div>" & vbCrlf
	funHtml = funHtml & "		<div class=""errorInfo"">" & vbCrlf
	funHtml = funHtml & "			<h2>" & ErrMsg & "</h2>" & vbCrlf
	funHtml = funHtml & "			<p>您可以按提示操作或联系管理员以解决此问题!</p>" & vbCrlf
	funHtml = funHtml & "		</div>" & vbCrlf
	funHtml = funHtml & "	</div>" & vbCrlf
	funHtml = funHtml & "	<div class=""reindex""><a class=""weui-btn weui-btn_warn"" href=""" & ParmPath & "Index/Start.html"">返回首页</a></div>" & vbCrlf
	funHtml = funHtml & "</div>" & vbCrlf
	funHtml = funHtml & getPageFoot(1)
	funHtml = Replace(funHtml, "[@FootScript]", "")
	funHtml = ReplaceCommonLabel(funHtml)
	GetErrBody = funHtml
End Function

Function GetCourseSelect(fField, fDefValue)		'取课程jqweui select
	Dim iFun, arrCourse, strFun
	arrCourse = Split(XmlText("Common", "Course", ""), "|")
	For iFun = 0 To Ubound(arrCourse)
		If iFun > 0 Then strFun = strFun & ","
		strFun = strFun & """" & arrCourse(iFun) & """"
	Next
	GetCourseSelect = strFun
End Function

Function getFieldSelect(fItemID, fField, fDefValue)		'取指定字段jqweui select
	on error resume next
	Dim fSheetName : fSheetName = "HR_Sheet_" & HR_Clng(fItemID)
	Dim rsFun, iFun, strFun : strFun = ""
	If ChkTable(fSheetName) And HR_IsNull(fField) = False Then
		Set rsFun = Conn.Execute("Select " & Trim(fField) & " From " & fSheetName & " Where " & Trim(fField) & "<>'' Group By " & Trim(fField))
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
	End If
	getFieldSelect = strFun
End Function

Function GetCampusSelect(fCampus, fDefValue)		'取校区jqweui select
	Dim iFun, arrCampus, strFun
	arrCampus = Split(XmlText("Common", "Campus", ""), "|")
	For iFun = 0 To Ubound(arrCampus)
		If iFun > 0 Then strFun = strFun & ","
		strFun = strFun & """" & arrCampus(iFun) & """"
	Next
	GetCampusSelect = strFun
End Function

Function GetClassRoomSelect(fClassRoom, fDefValue)		'取校区jqweui select
	Dim iFun, arrCampus, strFun
	arrCampus = Split(XmlText("Common", "Classroom", ""), "|")
	For iFun = 0 To Ubound(arrCampus)
		If iFun > 0 Then strFun = strFun & ","
		strFun = strFun & """" & arrCampus(iFun) & """"
	Next
	GetClassRoomSelect = strFun
End Function

'=====================================================================
'函数名：GetAssortList(param)	【弹窗选择科室】
'返回值：str
'=====================================================================
Function GetAssortList(fParent)
	Dim rsFun, sqlFun, strFun
	sqlFun = "Select * From HR_DepartAssort Where Parent='" & Trim(fParent) & "'"
	sqlFun = sqlFun & " Order By ID ASC"
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			Do While Not rsFun.EOF
				strFun = strFun & "<li data-id=""" & rsFun("ID") & """"
				strFun = strFun & "><span>" & Trim(rsFun("ChildAssort")) & "</span></li>" & vbCrlf
				rsFun.MoveNext
			Loop
		End If
	Set rsFun = Nothing
	GetAssortList = strFun
End Function

'=====================================================================
'函数名：GetManagerID(fStuType, fAll)		【取管理员ID按学生类别】
'=====================================================================
Function GetManagerID(fStuType, fAll)
	Dim strFun, rsFun, iFun :strFun = ""
	If HR_IsNull(fStuType) = False Then
		Set rsFun = Conn.Execute("Select * From HR_User Where ManageRank=1")
			If Not(rsFun.BOF And rsFun.EOF) Then
				iFun = 0
				Do While Not rsFun.EOF
					If FoundInArr(Trim(rsFun("StuType")), fStuType, ",") Then
						strFun = strFun & rsFun("UserID") & ","
					End If
					rsFun.MoveNext
					iFun = iFun + 1
				Loop
				strFun = FilterArrNull(strFun, ",")
			End If
		Set rsFun = Nothing
	End If
	GetManagerID = strFun
End Function

'=====================================================================
'函数名：PassProcess(fStep, fPass)		【审核状态】0待审1代课确认2教研主任审核3教学处审核4教辅审核
'=====================================================================
Function PassProcess(fStep, fPass)
	Dim strFun : strFun = "<span class=""step0"">待确认</span>"
	fStep = HR_CLng(fStep) : fPass = HR_CLng(fPass)
	If fStep = 1 Then
		If fPass=1 Then
			strFun = "<span class=""step1"">已确认</span>"
		ElseIf fPass=2 Then
			strFun = "<span class=""step2"">拒绝</span>"
		End If
	ElseIf fStep = 2 Then
		If fPass=1 Then
			strFun = "<span class=""step1"">同意</span>"
		ElseIf fPass=2 Then
			strFun = "<span class=""step2"">拒绝</span>"
		Else
			strFun = "<span class=""step0"">待审</span>"
		End If
	ElseIf fStep = 3 Then
		If fPass=1 Then
			strFun = "<span class=""step1"">同意</span>"
		ElseIf fPass=2 Then
			strFun = "<span class=""step2"">拒绝</span>"
		Else
			strFun = "<span class=""step0"">待审</span>"
		End If
	ElseIf fStep = 4 Then
		If fPass=1 Then
			strFun = "<span class=""step1"">审核完毕</span>"
		ElseIf fPass=2 Then
			strFun = "<span class=""step2"">拒绝</span>"
		Else
			strFun = "<span class=""step0"">待审</span>"
		End If
	ElseIf fStep = 5 Then
		strFun = "<span class=""step2"">撤销</span>"
	End If
	PassProcess = strFun
End Function
%>