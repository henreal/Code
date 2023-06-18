<%
Sub CEX()
	SiteTitle = "mini-CEX<sup>plus</sup>记录"
	Dim rsList, strList, tParm, tmpID
	If Ubound(arrParm) > 1 Then
		tParm = Trim(arrParm(2))
		tmpID = HR_Clng(Request("ID"))
		Select Case tParm
			Case "ViewCEX" Call ViewCEX()
			Case "EditCEX" Call EditCEX()
			Case "SaveCEX" Call SaveCEX()
		End Select
		Exit Sub
	End If
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-header {background-color:#061;}" & vbCrlf
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
	Set rsList = Conn.Execute("Select * From HR_EvaluateCEX Where TeacherID>0")
		If Not(rsList.BOF And rsList.EOF) Then
			Do While Not rsList.EOF
				strList = strList & "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManageEvaluate/CEX/ViewCEX.html?ID=" & rsList("ID") & """>" & vbCrlf
				strList = strList & "		<div class=""weui-cell__hd iconTit""><i class=""hr-icon"">&#xead1;</i></div><div class=""weui-cell__bd weui-cell_primary viewMSG"" data-id=""1""><p>测评教师：" & rsList("Teacher") & " 地点：" & rsList("EvaluateAdd") & "<br>学生：" & rsList("Student") & "　评价时间：" & FormatDate(rsList("EvaluateTime"), 2) & "</p></div>" & vbCrlf
				strList = strList & "		<div class=""weui-cell__ft""></div>" & vbCrlf
				strList = strList & "	</a>" & vbCrlf
				rsList.MoveNext
			Loop
		Else
			strList = "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef7b;</i></h2><h3>暂时还没有学生发表过评价！</h3></div>" & vbCrlf
		End If
	Set rsList = Nothing
	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf
	Response.Write " " & strList
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub EditCEX()
	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/lib/weui.min.css?v=1.1.3"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .weui-cell__ft {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .popWin i {font-size:1.2rem;position:relative;top:2px;color:#4ce}" & vbCrlf
	strHtml = strHtml & "		.weui-toast {margin-left: auto;} .weui-textarea{font-size:1rem}" & vbCrlf

	strHtml = strHtml & "		.hr-tips {border-top:1px solid #eee;padding:5px} .hr-tips .tipsIcon {color:#f30;padding-right:5px;font-size:1.5rem;}" & vbCrlf
	strHtml = strHtml & "		.title {border-top:1px solid #eee;padding:10px} .title h3 {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.remarkbar {padding:5px 10px;border-top:1px solid #eee;color:#444;background-color:#f2f2f2}" & vbCrlf
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

	Response.Write "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评教师：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Teacher"" class=""weui-input"" id=""Teacher"" type=""text"" value=""" & UserYGXM & """ data-key=""Teacher"" data-value=""TeacherID"" placeholder="""">" & vbCrlf
	Response.Write "			<input name=""TeacherID"" class=""weui-input"" id=""TeacherID"" type=""hidden"" value=""" & UserYGDM & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft popWin"" data-id=""Teacher""><i class=""hr-icon"">&#xeeed;</i>选择</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">职　务：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""TeacherJob"" class=""weui-input"" id=""TeacherJob"" type=""text"" value="""" placeholder=""点此选择"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">学生姓名：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Student"" class=""weui-input"" id=""Student"" type=""text"" value="""">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">学生专业：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Major"" class=""weui-input"" id=""Major"" type=""text"" value="""">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">类　别：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""SutType"" class=""weui-input"" id=""SutType"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评时间：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""EvaluateTime"" class=""weui-input"" id=""EvaluateTime"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评地点：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""EvaluateAdd"" class=""weui-input"" id=""EvaluateAdd"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>病人基本资料</h3></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">年龄：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""PatientAge"" class=""weui-input"" id=""PatientAge"" type=""number"" value="""" placeholder=""输入年龄"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">性别：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""PatientGender"" class=""weui-input"" id=""PatientGender"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">类别：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""PatientType"" class=""weui-input"" id=""PatientType"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>病人初步诊断（或主要问题）：</h3></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Impression"" id=""Impression"" placeholder=""请输入内容"" rows=""5""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">操作名称：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Treat"" class=""weui-input"" id=""Treat"" type=""text"" value="""">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">病情复杂程度：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Complexity"" class=""weui-input"" id=""Complexity"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">操作难度：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Difficulty"" class=""weui-input"" id=""Difficulty"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评重点：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Focus"" class=""weui-input"" id=""Focus"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
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
	Response.Write "			<input name=""Evaluate1"" class=""weui-input"" id=""Evaluate1"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch1"" class=""weui-switch"" id=""Switch1"" type=""checkbox"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score1"" id=""Score1"" class=""weui-count__number Score"" type=""number"" value=""0"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">2.体格检查：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate2"" class=""weui-input"" id=""Evaluate2"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch2"" class=""weui-switch"" id=""Switch2"" type=""checkbox"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score2"" id=""Score2"" class=""weui-count__number Score"" type=""number"" value=""0"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">3.临床操作：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate3"" class=""weui-input"" id=""Evaluate3"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch3"" class=""weui-switch"" id=""Switch3"" type=""checkbox"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score3"" id=""Score3"" class=""weui-count__number Score"" type=""number"" value=""0"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">4.临床思维与治疗：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate4"" class=""weui-input"" id=""Evaluate4"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch4"" class=""weui-switch"" id=""Switch4"" type=""checkbox"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score4"" id=""Score4"" class=""weui-count__number Score"" type=""number"" value=""0"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">5.医疗咨询与宣教：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate5"" class=""weui-input"" id=""Evaluate5"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch5"" class=""weui-switch"" id=""Switch5"" type=""checkbox"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score5"" id=""Score5"" class=""weui-count__number Score"" type=""number"" value=""0"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">6.沟通技能与人文关怀：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate6"" class=""weui-input"" id=""Evaluate6"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch6"" class=""weui-switch"" id=""Switch6"" type=""checkbox"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score6"" id=""Score6"" class=""weui-count__number Score"" type=""number"" value=""0"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">7.整体表现：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Evaluate7"" class=""weui-input"" id=""Evaluate7"" type=""text"" value="""" placeholder=""点此选择"" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell weui-cell_switch"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">未评测：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<input name=""Switch7"" class=""weui-switch"" id=""Switch7"" type=""checkbox"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score7"" id=""Score7"" class=""weui-count__number Score"" type=""number"" value=""0"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>得分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><input name=""TotalScore"" id=""TotalScore"" class=""weui-count__number"" type=""number"" value=""0"" placeholder=""自动计算"" readonly /></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>本次测评时间：</h3></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">直接观察：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><input name=""Duration"" id=""Duration"" class=""weui-input"" type=""number"" value=""0"" /></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">建议15-20分钟</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">反　馈：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><input name=""BackTime"" id=""BackTime"" class=""weui-input"" type=""number"" value=""0"" /></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">建议5-10分钟</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>教师评语：</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">值得肯定：</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Rraise"" id=""Rraise"" placeholder=""请输入内容"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">需要改进：</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Mend"" id=""Mend"" placeholder=""请输入内容"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">下一步措施：</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Means"" id=""Means"" placeholder=""请输入内容"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-btn-area""><em class=""weui-btn weui-btn_primary"" id=""subPost"">提交保存</em></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-header-hide""></div>" & vbCrlf
	Response.Write "<div id=""full"" class=""hr-popup"">" & vbCrlf
	Response.Write "	<iframe src=""about:bank"" name=""listFrame"" id=""listFrame"" title=""ListFrame"" width=""100%"" height=""100%"" frameborder=""0""></iframe>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf

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
	strHtml = strHtml & "	$(""#EvaluateTime"").calendar({dateFormat: 'yyyy年mm月dd日'});" & vbCrlf
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
	strHtml = strHtml & "	$(""#Difficulty"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择难度"",items:[""低"", ""中"", ""高""]," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#Focus"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择测评重点"",items:[""医疗问诊"", ""体格检查"", ""临床操作"", ""医疗咨询及宣教"", ""临床思维与治疗""]," & vbCrlf
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
	strHtml = strHtml & "		CountTotalScore();" & vbCrlf
	strHtml = strHtml & "	})" & vbCrlf
	strHtml = strHtml & "	$('.weui-count__increase').click(function (e) {" & vbCrlf
	strHtml = strHtml & "		var $input = $(e.currentTarget).parent().find('.weui-count__number');" & vbCrlf
	strHtml = strHtml & "		var number = parseInt($input.val() || ""0"") + 1" & vbCrlf
	strHtml = strHtml & "		if (number > maxNum) number = maxNum;" & vbCrlf
	strHtml = strHtml & "		$input.val(number);" & vbCrlf
	strHtml = strHtml & "		CountTotalScore();" & vbCrlf
	strHtml = strHtml & "	})" & vbCrlf

	strHtml = strHtml & "	$(""#subPost"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var TeacherID = $(""#TeacherID"").val();" & vbCrlf
	strHtml = strHtml & "		if(TeacherID==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""请选择测评教师"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else{" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Evaluate/CEX/SaveCEX.html"", $(""#EditForm"").serialize(), function(rsStr){" & vbCrlf
	strHtml = strHtml & "				$.toast(rsStr.reMessge, function(){ $.closePopup();location.href=""" & ParmPath & "/Evaluate/CEX.html""; });" & vbCrlf
	strHtml = strHtml & "			},""json"");" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	function CountTotalScore(){" & vbCrlf
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
	If HR_Clng(Request("TeacherID")) = 0 Then ErrMsg = "您没有选择测评教师"
	If HR_IsNull(ErrMsg) = False Then Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """}" : Exit Sub

	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open("Select * From HR_EvaluateCEX"), Conn, 1, 3
		rsSave.AddNew
		rsSave("ID") = GetNewID("HR_EvaluateCEX", "ID")
		rsSave("Teacher") = Trim(Request("Teacher"))
		rsSave("TeacherID") = HR_Clng(Request("TeacherID"))
		rsSave("TeacherJob") = Trim(Request("TeacherJob"))

		rsSave("Student") = Trim(Request("Student"))
		rsSave("Major") = Trim(Request("Major"))
		rsSave("SutType") = Trim(Request("SutType"))
		rsSave("EvaluateTime") = Trim(Request("EvaluateTime"))
		rsSave("EvaluateAdd") = Trim(Request("EvaluateAdd"))
		rsSave("PatientAge") = HR_Clng(Request("PatientAge"))
		rsSave("PatientGender") = Trim(Request("PatientGender"))
		rsSave("PatientType") = Trim(Request("PatientType"))
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

	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.title {border-top:1px solid #eee;padding:10px} .title h3 {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.hr-tips {border-top:1px solid #eee;padding:5px} .hr-tips .tipsIcon {color:#f30;padding-right:5px;font-size:1.5rem;}" & vbCrlf
	strHtml = strHtml & "		.remarkbar {padding:5px 10px;border-top:1px solid #eee;color:#444;background-color:#f2f2f2}" & vbCrlf
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

	Set rsTmp = Conn.Execute("Select * From HR_EvaluateCEX Where ID=" & tmpID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评教师：</label></div><div class=""weui-cell__bd"">" & rsTmp("Teacher") & "</div>" & vbCrlf
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
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">测评重点：</label></div><div class=""weui-cell__bd"">" & rsTmp("Focus") & "</div>" & vbCrlf
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
			If HR_CBool(rsTmp("Switch1")) Then Response.Write "[未评测]"
			Response.Write "		</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & Replace(rsTmp("Evaluate1"), ",", "<br>") & "</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div><div class=""weui-cell__bd"">" & rsTmp("Score1") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">2.体格检查：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
			If HR_CBool(rsTmp("Switch2")) Then Response.Write "[未评测]"
			Response.Write "		</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & Replace(rsTmp("Evaluate2"), ",", "<br>") & "</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div><div class=""weui-cell__bd"">" & rsTmp("Score2") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf


			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">3.临床操作：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
			If HR_CBool(rsTmp("Switch3")) Then Response.Write "[未评测]"
			Response.Write "		</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & Replace(rsTmp("Evaluate3"), ",", "<br>") & "</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div><div class=""weui-cell__bd"">" & rsTmp("Score3") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">4.临床思维与治疗：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
			If HR_CBool(rsTmp("Switch4")) Then Response.Write "[未评测]"
			Response.Write "		</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & Replace(rsTmp("Evaluate4"), ",", "<br>") & "</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div><div class=""weui-cell__bd"">" & rsTmp("Score4") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">5.医疗咨询与宣教：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
			If HR_CBool(rsTmp("Switch5")) Then Response.Write "[未评测]"
			Response.Write "		</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & Replace(rsTmp("Evaluate5"), ",", "<br>") & "</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div><div class=""weui-cell__bd"">" & rsTmp("Score5") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">6.沟通技能与人文关怀：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
			If HR_CBool(rsTmp("Switch6")) Then Response.Write "[未评测]"
			Response.Write "		</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & Replace(rsTmp("Evaluate6"), ",", "<br>") & "</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div><div class=""weui-cell__bd"">" & rsTmp("Score6") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">7.整体表现：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
			If HR_CBool(rsTmp("Switch7")) Then Response.Write "[未评测]"
			Response.Write "		</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & Replace(rsTmp("Evaluate7"), ",", "<br>") & "</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>测评结果：</p></div><div class=""weui-cell__bd"">" & rsTmp("Score7") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>得分：</p></div><div class=""weui-cell__bd"">" & rsTmp("TotalScore") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

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
		Else
			Response.Write "	<div class=""title""><h3>没有记录</h3></div>" & vbCrlf
		End If
	Set rsTmp = Nothing

	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-header-hide""></div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub
%>