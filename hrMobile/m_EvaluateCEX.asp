<%
Sub CEX()
	SiteTitle = "mini-CEX<sup>plus</sup>记录"
	Dim rsList, strList, tParm, tmpID, ArrDepart
	If Ubound(arrParm) > 1 Then
		tParm = Trim(arrParm(2))
		tmpID = HR_Clng(Request("ID"))
		Select Case tParm
			Case "ViewCEX" Call ViewCEX()
			Case "EditCEX" Call EditCEX()
			Case "SaveCEX" Call SaveCEX()
			Case "DelCEX" Call DeleteCEX()
		End Select
		Exit Sub
	End If
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-float-btn {width:55px;height:55px;text-align:center;font-size:2.6rem;position: fixed;right: 1rem;bottom: 3rem;z-index: 99;}" & vbCrlf
	strHtml = strHtml & "		.hr-float-btn i {color:#814ee2;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", "mini-CEX plus 记录")
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Set rsList = Conn.Execute("Select * From HR_EvaluateCEX Where TeacherID=" & UserYGDM)
		If Not(rsList.BOF And rsList.EOF) Then
			Do While Not rsList.EOF
				strList = strList & "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "/Evaluate/CEX/ViewCEX.html?ID=" & rsList("ID") & """>" & vbCrlf
				strList = strList & "		<div class=""weui-cell__hd iconTit""><i class=""hr-icon"">&#xead1;</i></div><div class=""weui-cell__bd weui-cell_primary viewMSG"" data-id=""1""><p>" & rsList("Student") & " " & rsList("EvaluateAdd") & "<br>评价时间：" & FormatDate(rsList("EvaluateTime"),2) & "</p></div>" & vbCrlf
				strList = strList & "		<div class=""weui-cell__ft""></div>" & vbCrlf
				strList = strList & "	</a>" & vbCrlf
				rsList.MoveNext
			Loop
		Else
			strList = "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef7b;</i></h2><h3>您暂时还没有发表评价！</h3></div>" & vbCrlf
		End If
	Set rsList = Nothing
	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf
	Response.Write " " & strList
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-float-btn"">" & vbCrlf
	Response.Write "	<a href=""" & ParmPath & "Evaluate/CEX/EditCEX.html?AddNew=True"" class=""addBtn""><i class=""hr-icon"">&#xf3c0;</i></a>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub EditCEX()
	Dim tDepart, tParent
	Set rs = Conn.Execute("Select Parent From HR_DepartAssort Group By Parent")
		If Not(rs.BOF And rs.EOF) Then
			Do While Not rs.EOF
				If i=0 Then
					tParent = rs("Parent")
					tDepart = tDepart & "<li class=""opt-this""><span>" & Trim(rs("Parent")) & "</span></li>"
				Else
					tDepart = tDepart & "<li><span>" & Trim(rs("Parent")) & "</span></li>"
				End If
				rs.MoveNext
				i = i + 1
			Loop
		End If
	Set rs = Nothing

	Dim tmpID : tmpID = HR_CLng(Request("ID"))
	Dim isModify : isModify = HR_CBool(Request("Modify"))

	Dim sqlGet, rsGet, tTechJob, tStudent, tMajor, tSutType, tEvaluateTime, tEvaluateAdd, tPatientAge, tPatientGender, tPatientType, tPatientKSMC, tImpression
	Dim tTreat, tComplexity, tDifficulty, tFocus, tEvaluate(7), tSwitch(7), tScore(7)
	Dim tTotalScore, tDuration, tBackTime, tRraise, tMend, tMeans, ArrDepart
	For i = 1 To 7
		tScore(i) = 0
	Next
	sqlGet = "Select a.* From HR_EvaluateCEX a Where a.ID=" & tmpID
	Set rsGet = Conn.Execute(sqlGet)
		If Not(rsGet.BOF And rsGet.EOF) Then
			isModify = True
			tTechJob = rsGet("TeacherJob")
			tStudent = rsGet("Student")
			tMajor = rsGet("Major")
			tSutType = rsGet("SutType")
			tEvaluateTime = rsGet("EvaluateTime")
			tEvaluateAdd = rsGet("EvaluateAdd")
			tPatientAge = rsGet("PatientAge")
			tPatientGender = rsGet("PatientGender")
			tPatientType = rsGet("PatientType")
			tPatientKSMC = rsGet("PatientKSMC")
			tImpression = rsGet("Impression")
			tTreat = rsGet("Treat")
			tComplexity = rsGet("Complexity")
			tDifficulty = rsGet("Difficulty")
			tFocus = rsGet("Focus")
			For i = 1 To 7
				tEvaluate(i) = rsGet("Evaluate" & i)
				tSwitch(i) = ""
				tScore(i) = HR_CLng(rsGet("Score" & i))
				If HR_CBool(rsGet("Switch" & i)) Then
					tSwitch(i) = " checked"
					tScore(i) = 0
				End If
			Next
			tTotalScore = rsGet("TotalScore")
			tDuration = rsGet("Duration")
			tBackTime = rsGet("BackTime")
			tRraise = rsGet("Rraise")
			tMend = rsGet("Mend")
			tMeans = rsGet("Means")
		End If
	Set rsGet = Nothing

	'取科室下拉数组
	Set rs = Conn.Execute("Select * From HR_Department Order By RootID ASC,OrderID ASC")
		If Not(rs.BOF And rs.EOF) Then
			Do While Not rs.EOF
				ArrDepart = ArrDepart & """" & rs("KSMC") & ""","
				rs.MoveNext
			Loop
		End If
	Set rs = Nothing
	ArrDepart = FilterArrNull(ArrDepart, ",")

	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/lib/weui.min.css?v=1.1.3"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		div,li,ul{box-sizing:border-box;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .weui-cell__ft {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .popWin i {font-size:1.2rem;position:relative;top:2px;color:#4ce}" & vbCrlf
	strHtml = strHtml & "		.weui-toast {margin-left: auto;} .weui-textarea{font-size:1rem}" & vbCrlf

	strHtml = strHtml & "		.hr-tips {border-top:1px solid #eee;padding:5px} .hr-tips .tipsIcon {color:#f30;padding-right:5px;font-size:1.5rem;}" & vbCrlf
	strHtml = strHtml & "		.title {border-top:1px solid #eee;padding:10px} .title h3 {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.remarkbar {padding:5px 10px;border-top:1px solid #eee;color:#444;background-color:#f2f2f2}" & vbCrlf
	strHtml = strHtml & "		.hr-pop-btn {padding:10px; position:fixed; bottom: 0px; width: 100%; box-sizing: border-box; background: #eee; border-top: 1px solid #f90;z-index:10;}" & vbCrlf
	strHtml = strHtml & "		.weui-count .Score{font-size:1rem;width:2.3rem}" & vbCrlf
	strHtml = strHtml & "		.hr-opt-floor {position:fixed;top:0;left:100%;background-color:#eee;border:5px solid #009688;box-sizing:border-box;width:100%;height:100%;z-index:1000}" & vbCrlf
	strHtml = strHtml & "		.hr-r-page {box-sizing:border-box;display:flex;width:100%;height:100%;}" & vbCrlf		'分类选择层
	strHtml = strHtml & "		.hr-r-page .hr-optbox {width:100px;border-right:1px solid #fff;box-sizing:border-box;flex-shrink: 0;}" & vbCrlf
	strHtml = strHtml & "		.hr-r-page .opt-menu li {line-height:40px;border-bottom:1px solid #fff;word-break:break-word;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding:0 5px}" & vbCrlf
	strHtml = strHtml & "		.hr-r-page .opt-menu li.opt-this {background:#fff}" & vbCrlf
	strHtml = strHtml & "		.hr-r-page .hr-opt-item {box-sizing:border-box;background:#fff;flex-grow: 2;}" & vbCrlf
	strHtml = strHtml & "		.hr-r-page .optlist {box-sizing:border-box;background:#fff;padding:0 10px;display:flex;flex-wrap:wrap;}" & vbCrlf
	strHtml = strHtml & "		.hr-r-page .optlist li {line-height:35px;padding:5px;width:50%}" & vbCrlf
	strHtml = strHtml & "		.hr-r-page .optlist li span {border:1px solid #ddd;display:block;word-break:break-word;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;box-sizing:border-box;padding:0 5px}" & vbCrlf

	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", "mini-CEX plus 记录")
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write "<header class=""hr-rows hr-header"">" & vbCrlf
	Response.Write "	<nav class=""navBack""><em><i class=""hr-icon"">&#xf320;</i></em></nav>" & vbCrlf
	Response.Write "	<nav class=""navTitle""><span>" & SiteTitle & "</span></nav>" & vbCrlf
	Response.Write "	<nav class=""navMenu""><em><i class=""hr-icon"">&#xf32a;</i></em></nav>" & vbCrlf
	Response.Write "</header>" & vbCrlf
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评教师：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Teacher"" class=""weui-input"" id=""Teacher"" type=""text"" value=""" & UserYGXM & """ data-key=""Teacher"" data-value=""TeacherID"" placeholder="""">" & vbCrlf
	Response.Write "			<input name=""TeacherID"" class=""weui-input"" id=""TeacherID"" type=""hidden"" value=""" & UserYGDM & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	'Response.Write "		<div class=""weui-cell__ft popWin"" data-id=""Teacher""><i class=""hr-icon"">&#xeeed;</i>选择</div>" & vbCrlf	'教师不用选
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">职　务：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""TeacherJob"" class=""weui-input"" id=""TeacherJob"" type=""text"" value=""" & tTechJob & """ placeholder=""点此选择"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">学生姓名：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Student"" class=""weui-input"" id=""Student"" type=""text"" value=""" & tStudent & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">学生专业：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Major"" class=""weui-input"" id=""Major"" type=""text"" value=""" & tMajor & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">类　别：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""SutType"" class=""weui-input"" id=""SutType"" type=""text"" value=""" & tSutType & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评时间：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""EvaluateTime"" class=""weui-input"" id=""EvaluateTime"" type=""text"" value=""" & tEvaluateTime & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评地点：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""EvaluateAdd"" class=""weui-input"" id=""EvaluateAdd"" type=""text"" value=""" & tEvaluateAdd & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>病人基本资料</h3></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">年龄：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""PatientAge"" class=""weui-input"" id=""PatientAge"" type=""number"" value=""" & tPatientAge & """ placeholder=""输入年龄"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">性别：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""PatientGender"" class=""weui-input"" id=""PatientGender"" type=""text"" value=""" & tPatientGender & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">类别：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""PatientType"" class=""weui-input"" id=""PatientType"" type=""text"" value=""" & tPatientType & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">科室：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""PatientKSMC"" class=""weui-input"" id=""PatientKSMC"" type=""text"" value=""" & tPatientKSMC & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><i class=""hr-icon"">&#xef8d;</i></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>病人初步诊断（或主要问题）：</h3></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Impression"" id=""Impression"" placeholder=""请输入内容"" rows=""5"">" & tImpression & "</textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">操作名称：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Treat"" class=""weui-input"" id=""Treat"" type=""text"" value=""" & tTreat & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">病情复杂程度：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Complexity"" class=""weui-input"" id=""Complexity"" type=""text"" value=""" & tComplexity & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">操作难度：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Difficulty"" class=""weui-input"" id=""Difficulty"" type=""text"" value=""" & tDifficulty & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评重点：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Focus"" class=""weui-input"" id=""Focus"" type=""text"" value=""" & tFocus & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""hr-rows hr-tips"">" & vbCrlf
	Response.Write "		<em class=""tipsIcon""><i class=""hr-icon"">&#xf06a;</i></em>" & vbCrlf
	Response.Write "		<em class=""hr-row-fill tipstxt"">测评标准：1-5不合格/6合格/7-8良好/9-10优秀</em>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">1.医疗问诊：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate1"" class=""weui-input"" id=""Evaluate1"" type=""text"" value=""" & tEvaluate(1) & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch1"" class=""weui-switch"" id=""Switch1"" value=""true"" type=""checkbox""" & tSwitch(1) & ">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score1"" id=""Score1"" class=""weui-count__number Score"" type=""number"" value=""" & tScore(1) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">2.体格检查：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate2"" class=""weui-input"" id=""Evaluate2"" type=""text"" value=""" & tEvaluate(2) & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch2"" class=""weui-switch"" id=""Switch2"" value=""true"" type=""checkbox""" & tSwitch(2) & ">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score2"" id=""Score2"" class=""weui-count__number Score"" type=""number"" value=""" & tScore(2) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">3.临床操作：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate3"" class=""weui-input"" id=""Evaluate3"" type=""text"" value=""" & tEvaluate(3) & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch3"" class=""weui-switch"" id=""Switch3"" value=""true"" type=""checkbox""" & tSwitch(3) & ">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score3"" id=""Score3"" class=""weui-count__number Score"" type=""number"" value=""" & tScore(3) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">4.临床思维与治疗：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate4"" class=""weui-input"" id=""Evaluate4"" type=""text"" value=""" & tEvaluate(4) & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch4"" class=""weui-switch"" id=""Switch4"" value=""true"" type=""checkbox""" & tSwitch(4) & ">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score4"" id=""Score4"" class=""weui-count__number Score"" type=""number"" value=""" & tScore(4) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">5.医疗咨询与宣教：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate5"" class=""weui-input"" id=""Evaluate5"" type=""text"" value=""" & tEvaluate(5) & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch5"" class=""weui-switch"" id=""Switch5"" value=""true"" type=""checkbox""" & tSwitch(5) & ">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score5"" id=""Score5"" class=""weui-count__number Score"" type=""number"" value=""" & tScore(5) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">6.沟通技能与人文关怀：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate6"" class=""weui-input"" id=""Evaluate6"" type=""text"" value=""" & tEvaluate(6) & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch6"" class=""weui-switch"" id=""Switch6"" value=""true"" type=""checkbox""" & tSwitch(6) & ">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score6"" id=""Score6"" class=""weui-count__number Score"" type=""number"" value=""" & tScore(6) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">7.整体表现：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate7"" class=""weui-input"" id=""Evaluate7"" type=""text"" value=""" & tEvaluate(7) & """ placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch7"" class=""weui-switch"" id=""Switch7"" value=""true"" type=""checkbox""" & tSwitch(7) & ">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score7"" id=""Score7"" class=""weui-count__number Score"" type=""number"" value=""" & tScore(7) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	'Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd""><p>得分：</p></div>" & vbCrlf
	'Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	'Response.Write "			<div class=""weui-count""><input name=""TotalScore"" id=""TotalScore"" class=""weui-count__number"" type=""number"" value=""0"" placeholder=""自动计算"" readonly /></div>" & vbCrlf
	'Response.Write "		</div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>本次测评时间：</h3></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">直接观察：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><input name=""Duration"" id=""Duration"" class=""weui-input"" type=""number"" value=""" & tDuration & """ /></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">建议15-20分钟</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">反　馈：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><input name=""BackTime"" id=""BackTime"" class=""weui-input"" type=""number"" value=""" & tBackTime & """ /></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">建议5-10分钟</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>教师评语：</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">值得肯定：</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Rraise"" id=""Rraise"" placeholder=""请输入内容"" rows=""2"">" & tRraise & "</textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">需要改进：</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Mend"" id=""Mend"" placeholder=""请输入内容"" rows=""2"">" & tMend & "</textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">下一步措施：</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Means"" id=""Means"" placeholder=""请输入内容"" rows=""2"">" & tMeans & "</textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""hr-pop-btn""><em class=""weui-btn weui-btn_primary"" id=""subPost"">提交保存</em></div>" & vbCrlf
	If isModify Then Response.Write "	<input name=""Modify"" type=""hidden"" value=""True""><input name=""ID"" type=""hidden"" value=""" & tmpID & """>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "<div id=""full"" class=""hr-popup"">" & vbCrlf
	Response.Write "	<iframe src=""about:bank"" name=""listFrame"" id=""listFrame"" title=""ListFrame"" width=""100%"" height=""100%"" frameborder=""0""></iframe>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-opt-floor"" id=""Assort"">" & vbCrlf
	Response.Write "	<div class=""hr-r-page"" id=""Assort"">" & vbCrlf
	Response.Write "		<div class=""hr-optbox""><ul class=""opt-menu"">" & tDepart & "</ul></div>" & vbCrlf
	Response.Write "		<div class=""hr-opt-item""><ul class=""optlist"">" & GetAssortList(tParent) & "</ul></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ location.href=""" & ParmPath & "Evaluate/CEX.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf

	strHtml = strHtml & "	$("".popWin"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		$(""#full"").show();var obj=$(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "		$(""#listFrame"").attr(""src"",""" & ParmPath & "Directories/SelectTeacher.html?Type=3&reObjTxt="" + $(""#""+obj).data(""key"") + ""&reObjValue="" +  $(""#""+obj).data(""value""));" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	$(""#TeacherJob"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择职务"",items:[""主任医师"", ""副主任医师"", ""主治医师"", ""完成住培的住院医师""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#SutType"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择类别"",items:[""实习医师"", ""住培/硕士研究生"", ""住院医师"", ""专培/博士研究生""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#EvaluateTime"").datetimePicker({max: '" & FormatDate(Now(), 2) & "'});" & vbCrlf
	strHtml = strHtml & "	$(""#EvaluateAdd"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择地点"",items:[""病房"", ""门诊"", ""急诊"", ""ICU"", ""临床技能中心"", ""其他""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#PatientGender"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择性别"",items:[""男"", ""女""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#PatientType"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择类别"",items:[""新接触患者"", ""已接触患者""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#Complexity"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择复杂程度"",items:[""低"", ""中"", ""高""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	
	strHtml = strHtml & "	SelectAssort();" & vbCrlf
	strHtml = strHtml & "	$(""#PatientKSMC"").on(""click"",function(){" & vbCrlf		'弹窗选择科室
	strHtml = strHtml & "		$(""#Assort"").animate({left:0});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$("".opt-menu li"").on(""click"",function(){" & vbCrlf		'选择科室大类
	strHtml = strHtml & "		var sortid = $(this).text();" & vbCrlf
	strHtml = strHtml & "		$("".opt-menu li"").removeClass(""opt-this"");" & vbCrlf
	strHtml = strHtml & "		$(this).addClass(""opt-this"");" & vbCrlf
	strHtml = strHtml & "		$.get(""" & ParmPath & "Assort/WinAssort.html"", {ModuleID:7, PID:sortid}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "			$("".optlist"").html(rsStr);SelectAssort();" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	function SelectAssort(){" & vbCrlf			'弹窗选中科室并返回值
	strHtml = strHtml & "		$("".optlist li"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "			var sortid = $(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "			var sortname = $(this).find(""span"").text();" & vbCrlf
	strHtml = strHtml & "			$(""#Assort"").animate({left:""100%""}); $(""#PatientKSMC"").val(sortname);" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	}" & vbCrlf


	strHtml = strHtml & "	$(""#Difficulty"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择难度"",items:[""低"", ""中"", ""高""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#Focus"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择测评重点"",multi:true,items:[""医疗问诊"", ""体格检查"", ""临床操作"", ""医疗咨询及宣教"", ""临床思维与治疗""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	$(""#Evaluate1"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择"",multi: true,items:[""正确称呼患者"", ""自我介绍"", ""向患者说明目的"", ""尽可能让患者自己陈述，适时给患者支持、鼓励"", ""耐心倾听患者陈述"", ""与患者有适当的眼神、言语、肢体的交流"", ""适时引导患者，以充分获取正确资料"", ""问诊逻辑清晰、条理清楚"", ""采用易懂语言"", ""重点突出，信息收集完整"", ""必要的记录""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#Evaluate2"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择"",multi: true,items:[""准备必需的体检用物"", ""注意保护患者的隐私，必要时，请其他人员在旁"", ""清洁双手"", ""按病情需要进行检查，顺序合理，及时处理患者在体检中出现的不适"", ""检查手法规范"", ""检查内容全面""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#Evaluate3"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择"",multi: true,items:[""了解适应证及相关解剖知识"",""取得患者同意（口头或书面）"",""操作前准备"",""适当的止痛或镇静"",""操作能力"",""无菌技术"",""适时寻求帮助"",""术后处理""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#Evaluate4"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择"",multi: true,items:[""能对病史与体检内容进行整合、分析"",""能解释相关的检查结果"",""临床分析具有逻辑性"",""有一定的诊断、鉴别诊断能力，"",""治疗方案合理可行""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#Evaluate5"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择"",multi: true,items:[""解释检查或处置的基本理由"",""各种治疗方案的利弊比较"",""患者用药指导"",""生活方式及注意事项的宣教""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#Evaluate6"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择"",multi: true,items:[""仪表端正，态度和蔼，口齿清楚"",""尊重患者与家属，具有同情心"",""获得患者与家属的信任"",""注意患者的舒适度，适时正确处理患者出现的不适"",""适当解释患者及家属提出的问题""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#Evaluate7"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择"",multi: true,items:[""对患者及家属态度"",""时间控制得当，过程简洁精炼"",""有整合资料与判断能力"",""能按优先顺序进行正确处理"",""整体效率高""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	var maxNum = 10, minNum = 0;" & vbCrlf
	strHtml = strHtml & "	$('.weui-count__decrease').click(function (e) {" & vbCrlf
	strHtml = strHtml & "		var $input = $(e.currentTarget).parent().find('.weui-count__number');" & vbCrlf
	strHtml = strHtml & "		var number = parseInt($input.val() || ""0"") - 1" & vbCrlf
	strHtml = strHtml & "		if (number < minNum) number = minNum;" & vbCrlf
	strHtml = strHtml & "		$input.val(number);" & vbCrlf
	'strHtml = strHtml & "		CountTotalScore();" & vbCrlf
	strHtml = strHtml & "	})" & vbCrlf
	strHtml = strHtml & "	$('.weui-count__increase').click(function (e) {" & vbCrlf
	strHtml = strHtml & "		var $input = $(e.currentTarget).parent().find('.weui-count__number');" & vbCrlf
	strHtml = strHtml & "		var number = parseInt($input.val() || ""0"") + 1" & vbCrlf
	strHtml = strHtml & "		if (number > maxNum) number = maxNum;" & vbCrlf
	strHtml = strHtml & "		$input.val(number);" & vbCrlf
	'strHtml = strHtml & "		CountTotalScore();" & vbCrlf
	strHtml = strHtml & "	})" & vbCrlf

	strHtml = strHtml & "	$("".weui-switch"").each(function(index, res){" & vbCrlf
	strHtml = strHtml & "		$(this).click(function(){" & vbCrlf
	strHtml = strHtml & "			if($(this).prop(""checked"")){ $(""#Score"" + (index+1)).val(""0"");}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	})" & vbCrlf

	strHtml = strHtml & "	$(""#subPost"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var TeacherID = $(""#TeacherID"").val();" & vbCrlf
	strHtml = strHtml & "		if(TeacherID==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""请选择测评教师"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else{" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Evaluate/CEX/SaveCEX.html"", $(""#EditForm"").serialize(), function(rsStr){" & vbCrlf
	strHtml = strHtml & "				$.alert(rsStr.reMessge, function(){" & vbCrlf
	strHtml = strHtml & "					$.closePopup();" & vbCrlf
	strHtml = strHtml & "					if(rsStr.Return){location.href=""" & ParmPath & "Evaluate/CEX.html"";}" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			},""json"");" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	function CountTotalScore(){" & vbCrlf			'统计总分
	strHtml = strHtml & "		var total=0;" & vbCrlf
	strHtml = strHtml & "		$("".Score"").each(function(){" & vbCrlf
	strHtml = strHtml & "			total = total + parseInt($(this).val());" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#TotalScore"").val(total);" & vbCrlf
	strHtml = strHtml & "	}" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub SaveCEX()
	Dim rsSave : ErrMsg = ""
	Dim tmpID : tmpID = HR_CLng(Request("ID"))
	Dim isModify : isModify = HR_CBool(Request("Modify"))

	If HR_Clng(Request("TeacherID")) = 0 Then ErrMsg = "您没有选择测评教师!"
	If HR_IsNull(Request("EvaluateTime")) Then ErrMsg = "您没有选择评测时间！"
	If Len(Request("Rraise")) < 7 Then ErrMsg = "肯定评语至少要输入8个字符！"
	If Len(Request("Mend")) < 8 Then ErrMsg = "改进内容至少要输入8个字符！"
	If Len(Request("Means")) < 8 Then ErrMsg = "下一步措施至少要输入8个字符！"
	If HR_IsNull(ErrMsg) = False Then Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """}" : Exit Sub

	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open("Select * From HR_EvaluateCEX Where ID>0 And ID=" & tmpID), Conn, 1, 3
		If rsSave.BOF And rsSave.EOF Then
			rsSave.AddNew
			rsSave("ID") = GetNewID("HR_EvaluateCEX", "ID")
			rsSave("Teacher") = Trim(Request("Teacher"))
			rsSave("TeacherID") = HR_Clng(Request("TeacherID"))
		End If
		rsSave("TeacherJob") = Trim(Request("TeacherJob"))

		rsSave("Student") = Trim(Request("Student"))
		rsSave("Major") = Trim(Request("Major"))
		rsSave("SutType") = Trim(Request("SutType"))
		rsSave("EvaluateTime") = SaveDate(Request("EvaluateTime"))
		rsSave("EvaluateAdd") = Trim(Request("EvaluateAdd"))
		rsSave("PatientAge") = HR_Clng(Request("PatientAge"))
		rsSave("PatientGender") = Trim(Request("PatientGender"))
		rsSave("PatientType") = Trim(Request("PatientType"))
		rsSave("PatientKSMC") = Trim(Request("PatientKSMC"))		
		rsSave("Impression") = Trim(Request("Impression"))
		rsSave("Treat") = Trim(Request("Treat"))
		rsSave("Complexity") = Trim(Request("Complexity"))
		rsSave("Difficulty") = Trim(Request("Difficulty"))
		rsSave("Focus") = Trim(Request("Focus"))

		rsSave("Evaluate1") = Trim(Request("Evaluate1"))
		rsSave("Switch1") = HR_CBool(Request("Switch1"))
		rsSave("Score1") = HR_Clng(Request("Score1"))

		rsSave("Evaluate2") = Trim(Request("Evaluate2"))
		rsSave("Switch2") = HR_CBool(Request("Switch2"))
		rsSave("Score2") = HR_Clng(Request("Score2"))

		rsSave("Evaluate3") = Trim(Request("Evaluate3"))
		rsSave("Switch3") = HR_CBool(Request("Switch3"))
		rsSave("Score3") = HR_Clng(Request("Score3"))

		rsSave("Evaluate4") = Trim(Request("Evaluate4"))
		rsSave("Switch4") = HR_CBool(Request("Switch4"))
		rsSave("Score4") = HR_Clng(Request("Score4"))

		rsSave("Evaluate5") = Trim(Request("Evaluate5"))
		rsSave("Switch5") = HR_CBool(Request("Switch5"))
		rsSave("Score5") = HR_Clng(Request("Score5"))

		rsSave("Evaluate6") = Trim(Request("Evaluate6"))
		rsSave("Switch6") = HR_CBool(Request("Switch6"))
		rsSave("Score6") = HR_Clng(Request("Score6"))

		rsSave("Evaluate7") = Trim(Request("Evaluate7"))
		rsSave("Switch7") = HR_CBool(Request("Switch7"))
		rsSave("Score7") = HR_Clng(Request("Score7"))

		rsSave("TotalScore") = HR_Clng(Request("TotalScore"))
		rsSave("Duration") = HR_Clng(Request("Duration"))
		rsSave("BackTime") = HR_Clng(Request("BackTime"))
		rsSave("Rraise") = Trim(Request("Rraise"))
		rsSave("Mend") = Trim(Request("Mend"))
		rsSave("Means") = Trim(Request("Means"))
		rsSave("CreateTime") = Now()
		rsSave.Update
	Set rsSave = Nothing
	ErrMsg = "记录已提交成功！"
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
End Sub

Sub ViewCEX()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tScore, tEvalu

	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.title {border-top:1px solid #eee;padding:10px} .title h3 {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.hr-tips {border-top:1px solid #eee;padding:5px} .hr-tips .tipsIcon {color:#f30;padding-right:5px;font-size:1.5rem;}" & vbCrlf
	strHtml = strHtml & "		.remarkbar {padding:5px 10px;border-top:1px solid #eee;color:#444;background-color:#f2f2f2}" & vbCrlf
	strHtml = strHtml & "		.cex-item {margin-left:1rem;} .cex-item li {list-style-type:disc;display:list-item;}" & vbCrlf
	strHtml = strHtml & "		.weui-label {width:auto;}" & vbCrlf
	strHtml = strHtml & "		.hr-pop-btn {padding:10px; position:fixed; bottom:0px; width:100%; box-sizing:border-box; background:#eee; border-top:1px solid #f90; z-index:10}" & vbCrlf
	strHtml = strHtml & "		.hr-pop-btn em {width:40%;} .weui-btn + .weui-btn {margin-top:0}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", "mini-CEX plus 记录")
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Set rsTmp = Conn.Execute("Select * From HR_EvaluateCEX Where TeacherID=" & UserYGDM & " And ID=" & tmpID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评教师：</label></div><div class=""weui-cell__bd"">" & UserYGXM & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">职　务：</label></div><div class=""weui-cell__bd"">" & rsTmp("TeacherJob") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">学生姓名：</label></div><div class=""weui-cell__bd"">" & rsTmp("Student") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">学生专业：</label></div><div class=""weui-cell__bd"">" & rsTmp("Major") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">类　别：</label></div><div class=""weui-cell__bd"">" & rsTmp("SutType") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评时间：</label></div><div class=""weui-cell__bd"">" & FormatDate(rsTmp("EvaluateTime"), 4) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评地点：</label></div><div class=""weui-cell__bd"">" & rsTmp("EvaluateAdd") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""title""><h3>病人基本资料</h3></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">年龄：</label></div><div class=""weui-cell__bd"">" & rsTmp("PatientAge") & " 岁</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">性别：</label></div><div class=""weui-cell__bd"">" & rsTmp("PatientGender") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">类别：</label></div><div class=""weui-cell__bd"">" & rsTmp("PatientType") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""title""><h3>病人初步诊断（或主要问题）：</h3></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & rsTmp("Impression") & "</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">操作名称：</label></div><div class=""weui-cell__bd"">" & rsTmp("Treat") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">病情复杂程度：</label></div><div class=""weui-cell__bd"">" & rsTmp("Complexity") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">操作难度：</label></div><div class=""weui-cell__bd"">" & rsTmp("Difficulty") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Dim tFocus
			If Trim(rsTmp("Focus")) <> "" Then tFocus = "<ul class=""cex-item""><li>" & Replace(rsTmp("Focus"), ",", "</li><li>") & "</li></ul>"
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评重点：</label></div><div class=""weui-cell__bd"">" & tFocus & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""hr-rows hr-tips"">" & vbCrlf
			Response.Write "		<em class=""tipsIcon""><i class=""hr-icon"">&#xf06a;</i></em>" & vbCrlf
			Response.Write "		<em class=""hr-row-fill tipstxt"">测评标准：1-5不合格/6合格/7-8良好/9-10优秀</em>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">1.医疗问诊：</label></div>" & vbCrlf
			tScore = HR_CLng(rsTmp("Score1")) & " 分" : If HR_CBool(rsTmp("Switch1")) Then tScore = "[未评测]"
			tEvalu = Replace(rsTmp("Evaluate1"), ",", "</li><li>")
			Response.Write "		<div class=""weui-cell__bd"">" & tScore & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			If HR_CBool(rsTmp("Switch1")) = False Then Response.Write "	<div class=""weui-cell""><ul class=""cex-item""><li>" & tEvalu & "</li></ul></div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">2.体格检查：</label></div>" & vbCrlf
			tScore = HR_CLng(rsTmp("Score2")) & " 分" : If HR_CBool(rsTmp("Switch2")) Then tScore = "[未评测]"
			tEvalu = Replace(rsTmp("Evaluate2"), ",", "</li><li>")
			Response.Write "		<div class=""weui-cell__bd"">" & tScore & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			If HR_CBool(rsTmp("Switch2")) = False Then Response.Write "	<div class=""weui-cell""><ul class=""cex-item""><li>" & tEvalu & "</li></ul></div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">3.临床操作：</label></div>" & vbCrlf
			tScore = HR_CLng(rsTmp("Score3")) & " 分" : If HR_CBool(rsTmp("Switch3")) Then tScore = "[未评测]"
			tEvalu = Replace(rsTmp("Evaluate3"), ",", "</li><li>")
			Response.Write "		<div class=""weui-cell__bd"">" & tScore & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			If HR_CBool(rsTmp("Switch3")) = False Then Response.Write "	<div class=""weui-cell""><ul class=""cex-item""><li>" & tEvalu & "</li></ul></div>" & vbCrlf
			'Response.Write "	<div class=""weui-cell"">" & vbCrlf
			'Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div><div class=""weui-cell__bd"">" & rsTmp("Score3") & " 分</div>" & vbCrlf
			'Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">4.临床思维与治疗：</label></div>" & vbCrlf
			tScore = HR_CLng(rsTmp("Score4")) & " 分" : If HR_CBool(rsTmp("Switch4")) Then tScore = "[未评测]"
			tEvalu = Replace(rsTmp("Evaluate4"), ",", "</li><li>")
			Response.Write "		<div class=""weui-cell__bd"">" & tScore & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			If HR_CBool(rsTmp("Switch4")) = False Then Response.Write "	<div class=""weui-cell""><ul class=""cex-item""><li>" & tEvalu & "</li></ul></div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">5.医疗咨询与宣教：</label></div>" & vbCrlf
			tScore = HR_CLng(rsTmp("Score5")) & " 分" : If HR_CBool(rsTmp("Switch5")) Then tScore = "[未评测]"
			tEvalu = Replace(rsTmp("Evaluate5"), ",", "</li><li>")
			Response.Write "		<div class=""weui-cell__bd"">" & tScore & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			If HR_CBool(rsTmp("Switch5")) = False Then Response.Write "	<div class=""weui-cell""><ul class=""cex-item""><li>" & tEvalu & "</li></ul></div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">6.沟通技能与人文关怀：</label></div>" & vbCrlf
			tScore = HR_CLng(rsTmp("Score6")) & " 分" : If HR_CBool(rsTmp("Switch6")) Then tScore = "[未评测]"
			tEvalu = Replace(rsTmp("Evaluate6"), ",", "</li><li>")
			Response.Write "		<div class=""weui-cell__bd"">" & tScore & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			If HR_CBool(rsTmp("Switch6")) = False Then Response.Write "	<div class=""weui-cell""><ul class=""cex-item""><li>" & tEvalu & "</li></ul></div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">7.整体表现：</label></div>" & vbCrlf
			tScore = HR_CLng(rsTmp("Score7")) & " 分" : If HR_CBool(rsTmp("Switch7")) Then tScore = "[未评测]"
			tEvalu = Replace(rsTmp("Evaluate7"), ",", "</li><li>")
			Response.Write "		<div class=""weui-cell__bd"">" & tScore & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			If HR_CBool(rsTmp("Switch7")) = False Then Response.Write "	<div class=""weui-cell""><ul class=""cex-item""><li>" & tEvalu & "</li></ul></div>" & vbCrlf

			'Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			'Response.Write "	<div class=""weui-cell"">" & vbCrlf
			'Response.Write "		<div class=""weui-cell__bd""><p>得分：</p></div><div class=""weui-cell__bd"">" & rsTmp("TotalScore") & " 分</div>" & vbCrlf
			'Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""title""><h3>本次测评时间：</h3></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">直接观察：</label></div><div class=""weui-cell__bd"">" & rsTmp("Duration") & " 分钟</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">反　馈：</label></div><div class=""weui-cell__bd"">" & rsTmp("BackTime") & " 分钟</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""title""><h3>教师评语：</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">值得肯定：</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & Replace(rsTmp("Rraise"), chr(13), "<br>") & "</div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">需要改进：</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & Replace(rsTmp("Mend"), chr(13), "<br>") & "</div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">下一步措施：</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & Replace(rsTmp("Means"), chr(13), "<br>") & "</div>" & vbCrlf
			Response.Write "	<div class=""hr-rows hr-pop-btn""><em class=""weui-btn weui-btn_primary"" data-event=""edit"" data-id=""" & tmpID & """>修改</em><em class=""weui-btn weui-btn_warn"" data-event=""delete"" data-id=""" & tmpID & """>删除</em></div>" & vbCrlf
		Else
			Response.Write "	<div class=""title""><h3>没有记录</h3></div>" & vbCrlf
		End If
	Set rsTmp = Nothing

	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-header-hide""></div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	strHtml = strHtml & "	$("".hr-pop-btn em"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var event=$(this).data(""event"");" & vbCrlf
	strHtml = strHtml & "		if(event==""edit""){" & vbCrlf
	strHtml = strHtml & "			location.href=""" & ParmPath & "Evaluate/CEX/EditCEX.html?AddNew=True&ID=" & tmpID & """;" & vbCrlf
	strHtml = strHtml & "		}else if(event==""delete""){" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "Evaluate/CEX/DelCEX.html"", {ID:""" & tmpID & """,Delete:""True""}, function(rsData){" & vbCrlf
	strHtml = strHtml & "				$.toast(rsData.reMessge,function(){" & vbCrlf
	strHtml = strHtml & "					location.href=""" & ParmPath & "Evaluate/CEX.html"";" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub DeleteCEX()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tDelete : tDelete = HR_CBool(Request("Delete"))

	Dim rsDel : ErrMsg = ""
	Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.Open("Select * From HR_EvaluateCEX Where TeacherID=" & UserYGDM & " And ID=" & tmpID), Conn, 1, 3
		If Not(rs.BOF And rs.EOF) Then
			rs.Delete
			ErrMsg = "删除成功"
		Else
			ErrMsg = "删除失败，数据不存在！"
		End If
	Set rs = Nothing
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
End Sub
%>