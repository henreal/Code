<%

Sub ListTeacher()
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"

	Dim listType : listType = HR_Clng(Request("Type"))
	Dim reObjTxt : reObjTxt = Trim(Request("reObjTxt"))
	Dim reObjValue : reObjValue = Trim(Request("reObjValue"))

	Dim sqlList, rsList, strList
	SiteTitle = "教师列表"
	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/weui.min.css?v=1.1.3"" rel=""stylesheet"" media=""all"">" & vbCrlf
	If GetUserAgent() = "iPhone" Then
		strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
		strHtml = strHtml & "	.sort_box {height:600px;overflow-y: auto;}" & vbCrlf
		strHtml = strHtml & "	</style>" & vbCrlf
	End If
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@Head_style]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	'strHtml = strHtml & "		$.showLoading();" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Head_script]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)

	Response.Write "<header class=""fixed""><div class=""header"">教师列表</div></header>" & vbCrlf
	Response.Write "<div id=""letter""></div>" & vbCrlf
	Response.Write "<div class=""sort_box"">" & vbCrlf
	Response.Write "	<div class=""weui-loadmore""><i class=""weui-loading""></i><span class=""weui-loadmore__tips"">正在加载</span></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""initials""><ul><li><i class=""hr-icon"">&#xef3e;</i></li></ul></div>" & vbCrlf


	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.charfirst.pinyin.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/sort.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(""html,body"").height($(window).height());" & vbCrlf
	strHtml = strHtml & "	$.get(""" & ParmPath & "/Course/GetTeacher.html"",{Type:" & listType & "}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "		$("".sort_box"").html(rsStr);" & vbCrlf
	strHtml = strHtml & "		initList();" & vbCrlf
	strHtml = strHtml & "		FastClick.attach(document.body);" & vbCrlf
	strHtml = strHtml & "		$("".num_name"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "			var ygdm = $(this).data(""ygdm"");" & vbCrlf
	strHtml = strHtml & "			var elParent = $(window.parent.document);" & vbCrlf
	strHtml = strHtml & "			elParent.find(""#full"").hide();" & vbCrlf
	strHtml = strHtml & "			elParent.find(""#" & reObjTxt & """).val($(this).data(""ygxm""));" & vbCrlf
	strHtml = strHtml & "			elParent.find(""#" & reObjValue & """).val(ygdm);" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	'strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@Foot_script]", strHtml)
	Response.Write strFootHtml
End Sub

Sub GetTeacherData()
	Dim sqlList, rsList
	Dim listType : listType = HR_Clng(Request("Type"))
	sqlList = "Select Top 1000 * From HR_Teacher Where Cast(YGDM As int)>0"
	If listType > 0 Then sqlList = sqlList & " And ApiType=" & listType
	sqlList = sqlList & " Order By YGXM collate Chinese_PRC_CS_AS_KS_WS"
	Set rsList = Server.CreateObject("ADODB.RecordSet")
		rsList.Open(sqlList), Conn, 1, 1
		If Not(rsList.BOF And rsList.EOF) Then
			m = 0
			Do While Not rsList.EOF
				'If m > 0 Then strList = strList & ","
				Response.Write "	<div class=""sort_list"">" & vbCrlf
				Response.Write "		<div class=""num_name"" data-ygdm=""" & rsList("YGDM") & """ data-ygxm=""" & rsList("YGXM") & """>" & rsList("YGXM") & "<b>[" & rsList("KSMC") & "]</b></div>" & vbCrlf
				Response.Write "	</div>" & vbCrlf
				rsList.MoveNext
				m = m + 1
			Loop
			'Response.Write "	<div class=""sort_list"">" & vbCrlf
			'Response.Write "		<div class=""num_name"" data-ygdm="""">&nbsp;" & Timer()-BeginTime & "</div>" & vbCrlf
			'Response.Write "	</div>" & vbCrlf
		End If
	Set rsList = Nothing
End Sub

Sub getItemGradeArr()		'取等级AJAX
	Dim tItemID : tItemID = HR_Clng(Request("Item"))
	Dim tLevel : tLevel = Trim(ReplaceBadChar(Request("Field")))
	Dim tValue : tValue = Trim(ReplaceBadChar(Request("value")))
	Dim rsGet, sqlGet, fStr, fLevelID
	If tItemID > 0 And HR_IsNull(tLevel) = False Then
		Set rsGet = Conn.Execute("Select Top 1 ID From HR_ItemModel Where ClassID=" & HR_Clng(tItemID) & " And FieldName='" & Trim(tLevel) & "'")
			fLevelID = HR_Clng(rsGet(0))
		Set rsGet = Nothing
		sqlGet = "Select * From HR_ItemGrade Where LevelID=" & fLevelID
		Set rsGet = Conn.Execute(sqlGet)
			If Not(rsGet.BOF And rsGet.EOF) Then
				fStr = "" : m = 0
				Do While Not rsGet.EOF
					If m > 0 Then fStr = fStr & ","
					fStr = fStr & "{""title"":""" & rsGet("Grade") & """, ""value"":""" & rsGet("ID") & """}"
					rsGet.MoveNext
					m = m + 1
				Loop
			End If
		Set rsGet = Nothing
	End If
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""取等级成功"",""reData"":[" & fStr & "]}"
End Sub

Sub SaveEdit()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim TempID : TempID = HR_Clng(Request("ID"))
	Dim tYGDM : tYGDM = HR_Clng(Request("ygdm"))
	Dim tYGXM : tYGXM = Trim(ReplaceBadChar(Request("ygxm")))
	Dim StudentType : StudentType = Trim(ReplaceBadChar(Request("StudentType")))
	Dim tAttach : tAttach = Trim(ReplaceBadUrl(Request("UploadAttach")))

	Dim IsModify : IsModify = HR_CBool(Request("Modify"))
	Dim SubButTxt : ErrMsg = "" : SubButTxt = "添加"

	Dim tItemName, tSheetName, tStuType, tTemplate, numField
	If tAttach <> "" Then tAttach = FilterArrNull(tAttach, "|")

	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tSheetName = "HR_Sheet_" & tItemID
			tStuType = Trim(rsTmp("StudentType"))
			tTemplate = Trim(rsTmp("Template"))
			numField = HR_Clng(rsTmp("FieldLen"))
		Else
			ErrMsg = "考核项目[ID:" & tItemID & "]不存在！<br>"
		End If
	Set rsTmp = Nothing

	If Not(ChkTable(tSheetName)) Then ErrMsg = "数据表 " & tSheetName & " 不存在！<br>"
	
	If UserYGDM <> "" And UserRank=0 Then tYGDM = UserYGDM		'非管理员限制仅添加本人
	If UserYGXM <> "" And UserRank=0 Then tYGXM = UserYGXM
	Dim tIsDate : tIsDate = False
	Dim tVA4 : tVA4 = Trim(Request("VA4"))
	If tTemplate = "TempTableA" Or tTemplate = "TempTableC" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE" Then
		If HR_IsNull(tVA4) Then
			ErrMsg = "日期还没有填写"
		ElseIf IsDate(tVA4) Then
			tVA4 = ConvertDateToNum(tVA4) + 2		'处理时间戳误差(非导入时转时间戳必须减2)
			tIsDate = True
		Else
			ErrMsg = "日期格式不正确！【" & tVA4 & "】"
		End If
	Else
		If HR_IsNull(tVA4) Then ErrMsg = "学期（学年）还没有填写！"
	End If

	Dim rsChk, sqlChk, sqlAdd, rsAdd
	sqlAdd = "Select * From " & tSheetName & " Where ItemID=" & tItemID & " And ID=" & TempID
	If StudentType <> "" Then sqlAdd = sqlAdd & " And StudentType='" & StudentType & "'"
	'判断数据是否重复
	
	sqlChk = "Select * From " & tSheetName & " Where ItemID=" & tItemID & ""
	If HR_IsNull(StudentType) = False Then sqlChk = sqlChk & " And StudentType='" & Trim(StudentType) & "'"
	If HR_Clng(tYGDM) > 0 Then sqlChk = sqlChk & " And VA1=" & HR_Clng(tYGDM)
	If HR_IsNull(tVA4) = False Then
		If tIsDate Then
			sqlChk = sqlChk & " And VA4=" & tVA4
		Else
			sqlChk = sqlChk & " And VA4='" & tVA4 & "'"
		End If
	End If
	Select Case tTemplate
		Case "TempTableA"
			sqlChk = sqlChk & " And VA7='" & Trim(Request("VA7")) & "' And VA8='" & Trim(Request("VA8")) & "' "		'判断工号、姓名、日期、节次、课程名称
			If HR_IsNull(Request("VA7")) Or HR_IsNull(Request("VA8")) Or HR_IsNull(Request("VA11")) Then
				ErrMsg = ErrMsg & "节次、课程名称、校区都不能为空！<br>"
			End If
		Case "TempTableB"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA5")) & "'"		'判断工号、姓名、学年(学期)、项目名称
			If Trim(Request("VA6")) <> "" Then sqlChk = sqlChk & " And Cast(VA6 as nvarchar)='" & Trim(Request("VA6")) & "'"
			If HR_IsNull(Request("VA5")) Then
				ErrMsg = ErrMsg & "学年(学期)或项目名称不能为空！<br>"
			End If
		Case "TempTableC"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA" & 5)) & "' And VA6='" & Trim(Request("VA" & 6)) & "'"
			If Trim(Request("VA" & 7)) <> "" Then sqlChk = sqlChk & " And Cast(VA7 As nvarchar)='" & Trim(Request("VA" & 7)) & "'"		'判断工号、姓名、学期、项目名称、备注
			If HR_IsNull(Request("VA5")) Or HR_IsNull(Request("VA6")) Then
				ErrMsg = ErrMsg & "学年(学期)或项目名称不能为空！<br>"
			End If
		Case "TempTableD"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA" & 5)) & "' And VA6='" & Trim(Request("VA" & 6)) & "'"		'判断工号、姓名、学期、教材
			If HR_IsNull(Request("VA5")) Or HR_IsNull(Request("VA6")) Then
				ErrMsg = ErrMsg & "学年(学期)或项目名称不能为空！<br>"
			End If
		Case "TempTableE"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA" & 5)) & "' And VA6='" & Trim(Request("VA" & 6)) & "'"		'判断工号、姓名、学期、案例、级别
			If Trim(Request("VA" & 7)) <> "" Then sqlChk = sqlChk & " And VA7='" & Trim(Request("VA" & 7)) & "'"
			If HR_IsNull(Request("VA5")) Or HR_IsNull(Request("VA6")) Then
				ErrMsg = ErrMsg & "学年(学期)、项目名称及级别不能为空！<br>"
			End If
		Case "TempTableF"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA" & 5)) & "' And VA6='" & Trim(Request("VA" & 6)) & "'"		'判断工号、姓名、学期、项目名称、级别
			If Trim(Request("VA" & 7)) <> "" Then sqlChk = sqlChk & " And VA7='" & Trim(Request("VA" & 7)) & "'"
			If HR_IsNull(Request("VA5")) Or HR_IsNull(Request("VA6")) Then
				ErrMsg = ErrMsg & "学年(学期)、项目名称及级别不能为空！<br>"
			End If
		Case "TempTableG"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA" & 5)) & "' And VA6='" & Trim(Request("VA" & 6)) & "'"		'判断工号、姓名、学期、项目名称、级别
			If HR_IsNull(Request("VA5")) Or HR_IsNull(Request("VA6")) Then
				ErrMsg = ErrMsg & "学年(学期)、项目名称不能为空！<br>"
			End If
		Case Else
			ErrMsg = ErrMsg & "您填写的数据与系统所有模型都不匹配！“" & tItemName & "”": ChkPass = False
	End Select
	If TempID > 0 Then sqlChk = sqlChk & " And ID Not In(" & TempID & ")"

	Set rsAdd = Conn.Execute(sqlChk)
		If Not(rsAdd.BOF And rsAdd.EOF) Then
			ErrMsg = ErrMsg & "教师 " & tYGXM & "[" & tYGDM & "] 的该条数据已经存在，不能重复添加！<br>"
		End If
	Set rsAdd = Nothing
	If ErrMsg <> "" Then
		Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}"
		Exit Sub
	End If
	Dim tPXXH
	Set rsAdd = Server.CreateObject("ADODB.RecordSet")
		rsAdd.Open(sqlAdd), Conn, 1, 3
		If rsAdd.BOF And rsAdd.EOF Then
			rsAdd.AddNew
			TempID = GetNewID(tSheetName, "ID")
			tPXXH = GetNewID(tSheetName, "VA0")
			rsAdd("ID") = TempID
			rsAdd("ItemID") = tItemID
			rsAdd("StudentType") = StudentType
			rsAdd("UserID") = 0
			rsAdd("AppendTime") = Now()
			rsAdd("scYear") = DefYear
			rsAdd("VA0") = tPXXH
			rsAdd("VA1") = HR_Clng(tYGDM)
			rsAdd("VA2") = Trim(tYGXM)
			rsAdd("State") = 1						'添加时状态0
		Else
			IsModify = True
			SubButTxt = "修改"
			If HR_Clng(rsAdd("UserID")) = 0 Then rsAdd("UserID") = UserID
			tPXXH = rsAdd("VA0")
		End If
		rsAdd("VA3") = HR_CDbl(Request("VA3"))
		rsAdd("VA4") = tVA4
		rsAdd("Passed") = False
		
		For i = 5 To numField-1
			rsAdd("VA" & i) = Trim(Request("VA" & i))
		Next
		rsAdd("Explain") = tAttach	'保存附件
		rsAdd.Update

	Set rsAdd = Nothing

	'tUpKPI = UpdateKPIField()		'此处更新业绩表字段
	'tUpKPI = ChkTeacherKPI(tYGDM)	'添加员工信息至业绩表
	'tUpKPI = UpdateTeacherKPI(tItemID, tYGDM, StudentType)	'更新本项目员工统计数据
	'tUpKPI = UpdateTeacherTotalKPI(tYGDM)	'更新员工总计数据

	Dim SendMsgManager, tArrSender
	Dim logUrl : logUrl = "Course/View.html?ItemID=" & tItemID & "&ID=" & TempID
	If UserYGXM <> "" And UserID=0 Then			'发送消息给管理员
		If HR_IsNull(StudentType) = False Then
			SendMsgManager = GetManagerID(StudentType, 0)
			tArrSender = Split(SendMsgManager, ",")
		Else
			tArrSender = arrManager
		End If
		For i = 0 To Ubound(tArrSender)
			ErrMsg = UserYGXM & SubButTxt & "了课程业绩需要您审核，考核项目：" & tItemName & "，序号：" & tPXXH & "，教师" & tYGXM & "[工号 " & tYGDM & "]，时间：" & FormatDate(Now(), 1)
			ErrMsg = ErrMsg & " <a href=""" & InstallDir & ManageDir & "Course.html?ItemID=" & tItemID & "&SearchWord=" & tYGDM & """>【查看】</a>"
			Call SendMessage(0, tItemID, TempID, tArrSender(i), UserYGXM & SubButTxt & "了课程业绩需要您审核", ErrMsg, logUrl)
		Next
	ElseIf UserID > 0 Then
		ErrMsg = tYGXM & "老师，系统" & SubButTxt & "了您的课程业绩需要您审阅，考核项目：" & tItemName & "，序号：" & tPXXH & "，教师" & tYGXM & "[工号 " & tYGDM & "]，时间：" & FormatDate(Now(), 1)
		ErrMsg = ErrMsg & " <a href=""" & InstallDir & ManageDir & "Course.html?ItemID=" & tItemID & "&SearchWord=" & tYGDM & """>【查看】</a>"
		Call SendMessage(0, tItemID, TempID, tYGDM, "您有新的课程业绩，请审阅", ErrMsg, logUrl)
	End If
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""保存成功"",""reData"":[" & tYGDM & "], ""id"":" & TempID & "}"
End Sub

Sub ViewAttach()
	SiteTitle = "查看附件" : ErrMsg = ""
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))

	Dim tItemName, tSheetName, tTemplate, tStuType
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tSheetName = "HR_Sheet_" & tItemID
			tStuType = Trim(rsTmp("StudentType"))
			tTemplate = Trim(rsTmp("Template"))
		Else
			ErrMsg = "考核项目[ID:" & tItemID & "]不存在！<br>"
		End If
	Set rsTmp = Nothing
	If Not(ChkTable(tSheetName)) Then ErrMsg = "数据表 " & tSheetName & " 不存在！<br>"

	If ErrMsg<>"" Then Response.Write GetErrBody(0) : Exit Sub
	Dim tAttach, picArr, strAttach, tArrAttach, AttachNum : AttachNum = 0
	Set rsTmp = Conn.Execute("Select * From " & tSheetName & " Where ID=" & tmpID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tAttach = rsTmp("Explain")		'取附件
		End If
	Set rsTmp = Nothing
	If HR_IsNull(tAttach) = False Then
		tArrAttach = Split(tAttach, "|")
		AttachNum = Ubound(tArrAttach) + 1
		For i = 0 To Ubound(tArrAttach)
			strAttach = strAttach & "<span class='pic_look' data-img='" & tArrAttach(i) & "' style='background-image: url(" & tArrAttach(i) & ")'><em id='delete_pic'>-</em></span>"
			If i> 0 Then picArr = picArr & ","
			picArr = picArr & """" & tArrAttach(i) & """"
		Next
	End If

	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<link type=""text/css"" href=""" & InstallDir & "Static/h5Upload/h5Upload.css"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "	.weui-dialog__bd {text-align:left;}" & vbCrlf
	strHtml = strHtml & "	.release_up_pic .tit{padding:12px;font-size:1.4rem;color:#999}" & vbCrlf
	strHtml = strHtml & "	.release_up_pic .tit h4{font-weight:400}" & vbCrlf
	strHtml = strHtml & "	.release_up_pic .tit h4 em{font-size:1.1rem}" & vbCrlf
	strHtml = strHtml & "	.release_up_pic .up_pic{background-color:#fff;padding:15px 12px;font-size:0;margin-left:-3.33333%;padding-bottom:3px;border-bottom:1px solid #e7e7e7;border-top:1px solid #e7e7e7}" & vbCrlf
	strHtml = strHtml & "	.release_up_pic .up_pic .pic_look{width:30%;height:80px;display:inline-block;background-size:cover;background-position:center center;background-repeat:no-repeat;box-sizing:border-box;margin-left:3.3333%;margin-bottom:12px;position:relative}" & vbCrlf
	strHtml = strHtml & "	.release_up_pic .up_pic .pic_look em{position:absolute;display:inline-block;width:25px;height:25px;background-color:red;color:#fff;font-size:18px;right:5px;top:5px;text-align:center;line-height:22px;border-radius:50%;font-weight:700}" & vbCrlf
	strHtml = strHtml & "	#chose_pic_btn {width:30%;height:80px;position:relative;display:inline-block;background:#eee url(" & InstallDir & "Static/images/upload.png) center no-repeat;box-sizing:border-box;background-size:30px 30px;border:1px solid #dbdbdb;margin-left:3.3333%;margin-bottom:12px}" & vbCrlf
	strHtml = strHtml & "	#chose_pic_btn input{position:absolute;left:0;top:0;opacity:0;width:100%;height:100%}" & vbCrlf
	strHtml = strHtml & "	.release_btn{padding:0 24px;margin-top:70px}" & vbCrlf
	strHtml = strHtml & "	.release_btn button{width:100%;background-color:#2c87af;font-size:1.4rem;color:#fff;border:0;border-radius:3px;height:45px;outline:0}" & vbCrlf
	strHtml = strHtml & "	.release_btn button.none_btn{background-color:#f2f2f2;color:#2c87af;border:1px solid #2c87af;margin-top:15px}" & vbCrlf
	strHtml = strHtml & "	.upbtn {box-sizing:border-box;padding:10px;} .upbtn em{width:50%;text-align:center;box-sizing:border-box;padding:0 10px;}" & vbCrlf
	strHtml = strHtml & "	#show1 {word-break: break-all;word-wrap: break-word;white-space: pre-wrap;}" & vbCrlf
	strHtml = strHtml & "	#loading1 {display:none;position:absolute;left:0;top:0;background:rgba(0,0,0,0.5) url(" & InstallDir & "Static/layui/css/modules/layer/default/loading-1.gif) center no-repeat;width:100%;height:100%;z-index:1000}" & vbCrlf
	strHtml = strHtml & "	.weui-photo-browser-modal {z-index:1000}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<div class=""release_up_pic"">" & vbCrlf
	Response.Write "	<div class=""up_pic"">" & vbCrlf
	Response.Write "		" & strAttach & vbCrlf
	Response.Write "		<span id=""chose_pic_btn"" style=""""><input type=""file"" accept=""image/*""></span>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div id=""show1""></div>" & vbCrlf
	Response.Write "<div class=""hr-rows upbtn""><em><button type=""button"" class=""weui-btn weui-btn_primary save"">保存</button></em><em><button type=""button"" class=""weui-btn weui-btn_primary preview"">预览</button></em></div>" & vbCrlf
	Response.Write "<div id=""loading1""></div>" & vbCrlf

	strHtml = "	<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/js/swiper.min.js?v=3.3.1""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Upload/localResizeIMG.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Upload/mobileBUGFix.mini.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ window.history.back() });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".addNew a"").css(""display"",""none"");" & vbCrlf
	strHtml = strHtml & "	$("".viewMSG"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		$.get(""" & ParmPath & "myCenter/viewMSG.html"",{id:$(this).data(""id"")}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "			$.alert(rsStr, ""查看消息"");" & vbCrlf
	strHtml = strHtml & "			$("".ShowCourse"").css(""display"",""none"");" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	var picArr = new Array(" & picArr & ");" & vbCrlf			'存储图片
	strHtml = strHtml & "	$(""input:file"").localResizeIMG({" & vbCrlf
	strHtml = strHtml & "		width:1920," & vbCrlf			'宽度
	strHtml = strHtml & "		quality: 0.6," & vbCrlf			'压缩参数 1 不压缩 越小清晰度越低
	strHtml = strHtml & "		success: function (result) {" & vbCrlf
	strHtml = strHtml & "			var img = new Image();" & vbCrlf
	strHtml = strHtml & "			img.src = result.base64;" & vbCrlf
	strHtml = strHtml & "			$.showLoading();" & vbCrlf			'上传提示
	strHtml = strHtml & "			$.ajax({" & vbCrlf
	strHtml = strHtml & "				url:""" & InstallDir & "API/UploadBase.htm"",type: ""POST"",data:{formFile:img.src,UploadDir:""Attach""}," & vbCrlf
	strHtml = strHtml & "				dataType: ""HTML"",timeout: 20000,error: function(){alert(""上传超时"");},success: function(reUrl){" & vbCrlf
	'strHtml = strHtml & "					$(""#show1"").html(result);" & vbCrlf
	strHtml = strHtml & "					var _str = ""<span class='pic_look' data-img='""+ reUrl + ""' style='background-image: url(""+ reUrl + "")'><em id='delete_pic'>-</em></span>""" & vbCrlf
	strHtml = strHtml & "					$('#chose_pic_btn').before(_str);" & vbCrlf
	strHtml = strHtml & "					$.hideLoading();" & vbCrlf				'关闭提示
	strHtml = strHtml & "					var _i =  picArr.length;" & vbCrlf
	strHtml = strHtml & "					picArr[_i] = reUrl;" & vbCrlf
	'strHtml = strHtml & "					picArr[_i] = _i;" & vbCrlf
	strHtml = strHtml & "				}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	// 删除" & vbCrlf
	strHtml = strHtml & "	$(document).on('click', '#delete_pic', function(event){" & vbCrlf
	strHtml = strHtml & "		var aa = $(this).parents("".pic_look"").index();" & vbCrlf
	strHtml = strHtml & "		picArr.splice(aa,1);" & vbCrlf
	strHtml = strHtml & "		$(this).parents("".pic_look"").remove();" & vbCrlf
	strHtml = strHtml & "		console.log(picArr);" & vbCrlf
	strHtml = strHtml & "	});	" & vbCrlf
	strHtml = strHtml & "	$(document).on('click', '.save', function(event){" & vbCrlf
	strHtml = strHtml & "		console.log(picArr);" & vbCrlf
	strHtml = strHtml & "		$.post(""" & ParmPath & "Course/SaveAttach.html"",{pic:picArr.join(""|""), ItemID:" & tItemID & ", ID:" & tmpID & "},function(reStr){" & vbCrlf
	strHtml = strHtml & "			$.toast(reStr.errmsg,function(){" & vbCrlf
	strHtml = strHtml & "				if(!reStr.err){ location.href=""" & ParmPath & "Course/View.html?ItemID=" & tItemID & "&ID=" & tmpID & """; }" & vbCrlf
	strHtml = strHtml & "			});	" & vbCrlf
	strHtml = strHtml & "		},""json"");" & vbCrlf
	strHtml = strHtml & "	});	" & vbCrlf
	strHtml = strHtml & "	$(document).on('click', '.preview', function(event){" & vbCrlf
	strHtml = strHtml & "		var pb1 = $.photoBrowser({ items:[" & picArr & "]});pb1.open(2);" & vbCrlf
	strHtml = strHtml & "		console.log(picArr);" & vbCrlf
	strHtml = strHtml & "	});	" & vbCrlf

	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub
Sub SaveAttach()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tPic : tPic = Trim(Request("pic"))
	sql = "Select * From HR_Sheet_" & tItemID & " Where ID=" & tmpID
	Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.Open sql, Conn, 1, 3
		If Not(rs.BOF And rs.EOF) Then
			rs("Explain") = tPic
			rs.Update
		End If
	Set rs = Nothing
	Response.Write "{""err"":false, ""errcode"":0, ""errmsg"":""附件保存成功！"", ""ItemID"":" & tItemID & ", ""CourseID"":" & tmpID & ", ""PicRows"":""" & tPic & """}"
End Sub

Function getFieldOption(fItem, fField, fValue)
	Dim rsGet, sqlGet, fStr
	If fItem > 0 Then
		If ChkTable("HR_Sheet_" & fItem) And HR_IsNull(fField) = False Then
			sqlGet = "Select " & Trim(fField) & " From HR_Sheet_" & fItem & " Group By " & Trim(fField)
			Set rsGet = Conn.Execute(sqlGet)
				If Not(rsGet.BOF And rsGet.EOF) Then
					m = 0
					Do While Not rsGet.EOF
						If m > 0 Then fStr = fStr & ","
						fStr = fStr & """" & rsGet(0) & """"
						rsGet.MoveNext
						m = m + 1
					Loop
				End If
			Set rsGet = Nothing
		End If
	End If
	getFieldOption = fStr
End Function

Sub ApplyModi()
	SiteTitle = "申请修改" : ErrMsg = ""
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim isModify : isModify = HR_CBool(Request("Modify"))

	'------ 取考核项目
	Dim tItemName, tSheetName, tTemplate, tStuType
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tSheetName = "HR_Sheet_" & tItemID
			tStuType = Trim(rsTmp("StudentType"))
			tTemplate = Trim(rsTmp("Template"))
		Else
			ErrMsg = "考核项目[ID:" & tItemID & "]不存在！<br>"
		End If
	Set rsTmp = Nothing

	'------ 取数据记录
	Dim StudentType, tYGDM, tYGXM, tPXXH, tSendUserID
	If ChkTable(tSheetName) Then
		Set rsTmp = Conn.Execute("Select * From " & tSheetName & " Where ID=" & tmpID)
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				tYGDM = HR_Clng(rsTmp("VA1"))					'取当前记录员工代码
				tYGXM = Trim(ReplaceBadChar(rsTmp("VA2")))		'取当前记录员工姓名
				StudentType = Trim(ReplaceBadChar(rsTmp("StudentType")))		'取当前记录类别
				tPXXH = HR_Clng(rsTmp("VA0"))
				tSendUserID = HR_Clng(rsTmp("UserID"))
			End If
		Set rsTmp = Nothing
	Else
		ErrMsg = "数据表 " & tSheetName & " 不存在！<br>"
	End If

	If isModify Then			'发送消息
		Dim logUrl : logUrl = "Course.html?ItemID=" & tItemID & "&YGDM=" & tYGDM & "&ID=" & tmpID
		Call SendMessage(1, tItemID, tmpID, 0, tYGXM & "申请修改[" & tItemName & "]课程业绩", Trim(Request("Reason")), logUrl)		'1申请修改课程,2退回课程修改，0其他消息
		Response.Write "{""Return"":true,""Err"":0,""reMessge"":""修改课程业绩申请提交成功！"",""msgTitle"":""操作失败""}"
		Exit Sub
	Else
		If ErrMsg<>"" Then Response.Write GetErrBody(0) : Exit Sub
	End If

	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/weui.min.css?v=1.1.3"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "	.weui-dialog__bd {text-align:left;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<div class=""weui-cells__title"">申请理由</div>" & vbCrlf
	Response.Write "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" placeholder=""请输入修改原因"" name=""Reason"" rows=""5""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<input type=""hidden"" name=""ItemID"" value=""" & tItemID & """><input type=""hidden"" name=""ID"" value=""" & tmpID & """><input type=""hidden"" name=""Modify"" value=""True"">" & vbCrlf
	Response.Write "<div class=""weui-btn-area"">" & vbCrlf
	Response.Write "	<a class=""weui-btn weui-btn_primary"" href=""javascript:"" id=""appPost"">确定</a>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "</form>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "/Course/View.html?ItemID=" & tItemID & "&ID=" & tmpID & """; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".addNew a"").css(""display"",""none"");" & vbCrlf

	strHtml = strHtml & "	$(""#appPost"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "		$.getJSON(""" & ParmPath & "Course/ApplyModi.html"", $(""#EditForm"").serialize(), function(strForm){" & vbCrlf
	strHtml = strHtml & "			$.alert(strForm.reMessge,function(){" & vbCrlf
	strHtml = strHtml & "				if(strForm.Return){ location.href=""" & ParmPath & "Course/View.html?ItemID=" & tItemID & "&ID=" & tmpID & """; }" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub Affirm()
	SiteTitle = "确认提交" : ErrMsg = ""
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim isModify : isModify = HR_CBool(Request("Modify"))

	'------ 取考核项目
	Dim tItemName, tSheetName, tTemplate, tStuType
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tSheetName = "HR_Sheet_" & tItemID
			tStuType = Trim(rsTmp("StudentType"))
			tTemplate = Trim(rsTmp("Template"))
		Else
			ErrMsg = "考核项目[ID:" & tItemID & "]不存在！<br>"
		End If
	Set rsTmp = Nothing
	'------ 取数据记录
	Dim tYGDM, tYGXM, tPXXH, StudentType
	If ChkTable(tSheetName) = False Then ErrMsg = "数据表 " & tSheetName & " 不存在！<br>"
	If ErrMsg <> "" Then Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""reTitle"":""操作失败""}" : Exit Sub

	Set rsTmp = Conn.Execute("Select * From " & tSheetName & " Where VA1=" & HR_Clng(UserYGDM) & " And ID=" & tmpID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tYGDM = HR_Clng(rsTmp("VA1"))						'取当前记录员工代码
			tYGXM = Trim(ReplaceBadChar(rsTmp("VA2")))			'取当前记录员工姓名
			StudentType = Trim(ReplaceBadChar(rsTmp("StudentType")))		'取当前记录类别
			tPXXH = HR_Clng(rsTmp("VA0"))
			Conn.Execute("Update " & tSheetName & " Set State=1 Where ID=" & rsTmp("ID"))
			ErrMsg = "您选择中的课程业绩已确认提交！"
		Else
			ErrMsg = "您选择中的课程业绩不在存！"
		End If
	Set rsTmp = Nothing
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """,""reTitle"":""操作成功""}"
End Sub

Sub Swap()	'调换课程
	SiteTitle = "申请调换课程" : ErrMsg = ""
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))

	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "	.weui-dialog__bd {text-align:left;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<div class=""weui-cells__title"">申请理由</div>" & vbCrlf
	Response.Write "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" placeholder=""请输入修改原因"" rows=""5""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<input type=""hidden"" name=""ItemID"" value=""" & tItemID & """><input type=""hidden"" name=""ID"" value=""" & tmpID & """><input type=""hidden"" name=""Modify"" value=""True"">" & vbCrlf
	Response.Write "<div class=""weui-btn-area"">" & vbCrlf
	Response.Write "	<a class=""weui-btn weui-btn_primary"" href=""javascript:"" id=""appPost"">确定</a>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "</form>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "/Course/List.html?ItemID=" & tItemID & """; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf

	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub ShowCourse()		'显示课程详情，仅用于AJAX异步读取，返回string
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))		'项目ID
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim isShow : isShow = False
	SiteTitle = "课程详情" : ErrMsg = ""
	Dim strShow, sqlShow, rsShow, tItemName, tSheetName, tTemplate, tStuType, tUnit, tFieldLen, tFieldHead, tArrHead
	Dim tVA4, tmpTime

	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)			'取项目信息
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tSheetName = "HR_Sheet_" & tItemID
			tStuType = Trim(rsTmp("StudentType"))
			tTemplate = Trim(rsTmp("Template"))
			tUnit = rsTmp("Unit")
			tFieldLen = HR_CLng(rsTmp("FieldLen"))
			tFieldHead = Trim(rsTmp("FieldHead"))
			If Not(ChkTable(tSheetName)) Then ErrMsg = ErrMsg & "未找到数据表 " & tSheetName & "！<br>"
		Else
			ErrMsg = ErrMsg & "<li>考核项目[ID:" & tItemID & "]不存在！</li>"
		End If
	Set rsTmp = Nothing
	If HR_IsNull(tFieldHead) = False Then
		tFieldHead = FilterArrNull(tFieldHead, ",")
		tArrHead = Split(tFieldHead, ",")
		If Ubound(tArrHead) <> tFieldLen Then Redim Preserve tArrHead(tFieldLen)
	Else
		Redim tArrHead(tFieldLen)
	End If

	If ErrMsg<>"" Then Response.Write GetErrBody(0) : Exit Sub
	SiteTitle = tItemName & " 详情"

	'式样及Head/Foot均为演示，调用时由调用页定义UI
	If isShow Then
		tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
		tmpHtml = tmpHtml & "		.viewPanel li.hr-gap-20 {background-color:#eee;}" & vbCrlf
		tmpHtml = tmpHtml & "	</style>" & vbCrlf
		strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
		strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", tmpHtml)

		tmpHtml = vbCrlf & "	<script type=""text/javascript"">" & vbCrlf
		tmpHtml = tmpHtml & "		$(document).ready(function(){});" & vbCrlf
		tmpHtml = tmpHtml & "	</script>" & vbCrlf
		strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", tmpHtml)
		Response.Write ReplaceCommonLabel(strHeadHtml & getHeadNav(0))
		Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	End If
	sqlShow = "Select a.*,b.KSMC,b.PRZC From " & tSheetName & " a Inner Join HR_Teacher b on a.VA1=b.YGDM Where a.ID=" & tmpID
	Set rsShow = Server.CreateObject("ADODB.RecordSet")
		rsShow.Open(sqlShow), Conn, 1, 1
		If Not(rsShow.BOF And rsShow.EOF) Then
			tVA4 = Trim(rsShow("VA4"))
			If HR_CLng(tVA4) > 1000 Then tVA4 = FormatDate(ConvertNumDate(tVA4), 2)
			strShow = strShow & "<ul class=""viewPanel"">" & vbCrlf
			strShow = strShow & "	<li class=""hr-gap-20""></li>" & vbCrlf
			strShow = strShow & "	<li class=""info listItem""><b>" & tArrHead(2) & "：</b><em>" & rsShow("VA2") & " [" & rsShow("VA1") & "]</em></li>" & vbCrlf
			strShow = strShow & "	<li class=""info""><b>科室：</b><em>" & rsShow("KSMC") & "</em></li>" & vbCrlf
			strShow = strShow & "	<li class=""info""><b>职称：</b><em>" & rsShow("PRZC") & "</em></li>" & vbCrlf
			strShow = strShow & "	<li class=""info""><b>考核项目：</b><em>" & tItemName & "</em></li>" & vbCrlf
			strShow = strShow & "	<li class=""info""><b>" & tArrHead(3) & "：</b><em>" & rsShow("VA3") & " " & tUnit & "</em></li>" & vbCrlf
			strShow = strShow & "	<li class=""hr-gap-20""></li>" & vbCrlf
			strShow = strShow & "	<li class=""info""><b>" & tArrHead(4) & "：</b><em>" & tVA4 & "</em></li>" & vbCrlf
			If tTemplate = "TempTableA" Then
				strShow = strShow & "	<li class=""info""><b>" & tArrHead(5) & "：</b><em>" & rsShow("VA5") & " [星期" & rsShow("VA6") & "]</em></li>" & vbCrlf
				tmpTime = GetPeriodTime(Trim(rsShow("VA11")), rsShow("VA7"), 0)		'计算节次时间
				strShow = strShow & "	<li class=""info""><b>" & tArrHead(7) & "：</b><em>第" & rsShow("VA7") & "节 " & tmpTime & "</em></li>" & vbCrlf
			Else
				strShow = strShow & "	<li class=""info""><b>" & tArrHead(5) & "：</b><em>" & rsShow("VA5") & "</em></li>" & vbCrlf
				strShow = strShow & "	<li class=""info""><b>" & tArrHead(6) & "：</b><em>" & rsShow("VA6") & "</em></li>" & vbCrlf
			End If
			If tTemplate <> "TempTableB" Then
				strShow = strShow & "	<li class=""info""><b>" & tArrHead(7) & "：</b><em>" & rsShow("VA7") & "</em></li>" & vbCrlf
			End If
			If tTemplate = "TempTableA" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE" Or tTemplate = "TempTableF" Then
				strShow = strShow & "	<li class=""info""><b>" & tArrHead(8) & "：</b><em>" & rsShow("VA8") & "</em></li>" & vbCrlf
			End If
			If tTemplate = "TempTableA" Or tTemplate = "TempTableE" Then
				strShow = strShow & "	<li class=""info""><b>" & tArrHead(9) & "：</b><em>" & rsShow("VA9") & "</em></li>" & vbCrlf
			End If
			If tTemplate = "TempTableA" Then
				strShow = strShow & "	<li class=""info""><b>" & tArrHead(10) & "：</b><em>" & rsShow("VA10") & "</em></li>" & vbCrlf
				strShow = strShow & "	<li class=""info""><b>" & tArrHead(11) & "：</b><em>" & rsShow("VA11") & "</em></li>" & vbCrlf
				strShow = strShow & "	<li class=""info""><b>" & tArrHead(12) & "：</b><em>" & rsShow("VA12") & "</em></li>" & vbCrlf
			End If

			strShow = strShow & "	<li class=""hr-gap-20""></li>" & vbCrlf
			strShow = strShow & "</ul>" & vbCrlf
		Else
			strShow = "课程不在存" & vbCrlf
		End If
	Set rsShow = Nothing
	Response.Write strShow
	If isShow Then
		strFootHtml = Replace(strFootHtml, "[@FootScript]", "")
		Response.Write ReplaceCommonLabel(strFootHtml)
	End If
End Sub
%>