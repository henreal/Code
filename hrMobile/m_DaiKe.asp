<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./m_ManageCommon.asp"-->
<!--#include file="../Core/classItem.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl : SiteTitle = "填写代课申请"
If ChkWechatTokenQY() = False Then Call GetWechatTokenQY()		'提前检查企业微信Token是否过期
If ChkTokenBobao() = False Then Call GetTokenBobao()			'检查信息播报Token

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "getItemJson" Call getItemJson()
	Case "Step2" Call Step2()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim arrCourse : arrCourse = GetTableDataQuery("HR_Swap", "", 1, "ID=" & tmpID & "")		'// 取代换课信息
	
	If HR_Clng(arrCourse(4, 1)) <> UserYGDM Then	' 判断是否本人的代课申请
		ErrMsg = "不是您的代课申请"
		'Response.Write GetErrBody(1) : Response.End
	End If

	Dim tItemID : tItemID = HR_Clng(arrCourse(1, 1))		'// 取项目ID
	Dim tTableName : tTableName = "HR_Sheet_" & tItemID		'// 取项目表名

	Dim tItemName : tItemName = GetTypeName("HR_Class", "ClassName", "ClassID", tItemID)
	Dim tTemplate : tTemplate = GetTypeName("HR_Class", "Template", "ClassID", tItemID)		'// 取项目模板类型
	
	Dim tCourseID : tCourseID = HR_Clng(arrCourse(2, 1))		'// 取原课程信息ID
	Dim tCourse : tCourse = ""

	'// 原课程信息
	Dim arrOldCourse, oldCourseDate, oldContents, tCourseDate, tContents, oldStudent, tStudent
	Dim IsModify : IsModify = False

	If HR_clng(arrCourse(0, 1)) > 0 Then
		IsModify = True
		arrOldCourse = GetTableDataQuery(tTableName, "", 1, "ID=" & tCourseID & "")		'// 取原课程信息
		oldCourseDate = FormatDate(arrCourse(6, 1), 2)		'// 原上课时间
		tCourseDate = FormatDate(arrCourse(18, 1), 2)		'// 新上课时间
		tCourse = FormatDate(arrCourse(18, 1), 2)
		tCourse = tCourse & " " & arrCourse(22, 1) & "_第" & arrCourse(21, 1) & "节"	'// 新课程节次

		oldContents = arrCourse(11, 1)
		tContents = arrCourse(23, 1)
		oldStudent = arrCourse(12, 1)
		tStudent = arrCourse(24, 1)
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-header {z-index:1}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-cells_form .weui-cell__ft {font-size:1.1rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-cells_form .popWin i {font-size:1.2rem;position:relative;top:2px;color:#4ce}" & vbCrlf

	tmpHtml = tmpHtml & "		.old-box h3 {padding:10px; border-bottom:1px solid #4fb74e;}" & vbCrlf
	tmpHtml = tmpHtml & "		.old-box ul {padding:10px; display:flex; flex-direction:column;}" & vbCrlf
	tmpHtml = tmpHtml & "		.old-box li {padding:10px; border-bottom:1px solid #ddd; display:flex;}" & vbCrlf
	tmpHtml = tmpHtml & "		.old-box li tt {width:5.2rem;color:#999;flex-shrink:0;}" & vbCrlf
	tmpHtml = tmpHtml & "		.old-box li em {font-size:1.1rem;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells__title"">第一步：选择原课程</div>" & vbCrlf
	Response.Write "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">申请人：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Applyer"" class=""weui-input"" id=""Applyer"" type=""text"" value=""" & UserYGXM & """ readonly>" & vbCrlf
	Response.Write "			<input name=""ApplyID"" class=""weui-input"" id=""ApplyID"" type=""hidden"" value=""" & UserYGDM & """ data-values=""" & UserYGDM & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">选择项目：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Item"" class=""weui-input opt1"" id=""Item"" type=""text"" value=""" & tItemName & """ data-values=""" & tItemID & """ readonly>" & vbCrlf
	Response.Write "			<input name=""ItemID"" class=""weui-input"" id=""ItemID"" type=""hidden"" value=""" & tItemID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><i class=""hr-icon"">&#xf321;</i></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">选择课程：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Course"" class=""weui-input opt2"" id=""Course"" type=""text"" value=""" & tCourse & """ data-values=""" & tCourseID & """>" & vbCrlf
	Response.Write "			<input name=""CourseID"" class=""weui-input"" id=""CourseID"" type=""hidden"" value=""" & tCourseID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><i class=""hr-icon"">&#xf321;</i></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20 hr-gapbg""></div>" & vbCrlf

	Dim show : show = ""
	If IsModify = False Then show = " style=""display:none"""
	Response.Write "	<div class=""old-box""" & show & "><h3>课程信息：</h3><ul>" & vbCrlf 	'//原课程信息
	Response.Write "	<li><tt>授课日期：</tt><em>" & oldCourseDate & "</em></li>" & vbCrlf
	Response.Write "	<li><tt>节次：</tt><em>" & Trim(arrCourse(9, 1)) & "</em></li>" & vbCrlf
	Response.Write "	<li><tt>学时：</tt><em>" & Trim(arrCourse(5, 1)) & "</em></li>" & vbCrlf
	Response.Write "	<li><tt>授课对象：</tt><em>" & oldStudent & "</em></li>" & vbCrlf
	Response.Write "	<li><tt>授课内容：</tt><em>" & oldContents & "</em></li>" & vbCrlf
	Response.Write "	<li><tt>校(院)区：</tt><em>" & Trim(arrCourse(13, 1)) & "</em></li>" & vbCrlf
	Response.Write "	</ul>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20 hr-gapbg""></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	If tmpID > 0 Then Response.Write "<input name=""ID"" type=""hidden"" value=""" & tmpID & """>" & vbCrlf
	Response.Write "<div class=""weui-btn-area""><em class=""weui-btn weui-btn_primary"" id=""subPost"">提交申请</em></div>" & vbCrlf
	
	Response.Write "</div>" & vbCrlf
	Response.Write "</form>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf

	'//选择考核项目
	tmpHtml = tmpHtml & "	var arrItem =[" & GetSelectOptionItem() & "];" & vbCrlf		'// 取考核项目
	tmpHtml = tmpHtml & "	$(""#Item"").select({" & vbCrlf
	tmpHtml = tmpHtml & "		title: ""请选择"",items:arrItem," & vbCrlf
	tmpHtml = tmpHtml & "		onChange:function(res){" & vbCrlf
	'tmpHtml = tmpHtml & "			console.log(res);" & vbCrlf
	tmpHtml = tmpHtml & "		}," & vbCrlf
	tmpHtml = tmpHtml & "		onClose:function(e){" & vbCrlf
	'tmpHtml = tmpHtml & "			console.log(e.data.values);" & vbCrlf
	tmpHtml = tmpHtml & "			var tid = $(""#Item"").data(""values"");" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#ItemID"").val(tid);" & vbCrlf
	tmpHtml = tmpHtml & "			$.get(""" & ParmPath & "Substitute/getItemCourse.html"",{Item:tid}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "				var reData = eval(""("" + strForm + "")"");" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#Course"").select(""update"", reData);" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#Course"").val("""");" & vbCrlf
	tmpHtml = tmpHtml & "				$("".old-box ul"").html("""");" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#Student"").val(""""); $(""#VA12"").val(""""); $(""#VA11"").val("""");" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#Contents"").val(""""); $(""#CourseDate"").val("""");" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#VA7"").val(""""); $(""#VA3"").val(""""); $(""#VA5"").val("""");" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#VA8"").val(""""); $(""#VA6"").val(""""); " & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		}" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	'===== 选择课程 =====
	tmpHtml = tmpHtml & "	$(""#Course"").select({" & vbCrlf
	tmpHtml = tmpHtml & "		title:""选择课程"",items:[{title:""暂无课程"",value:""""}]," & vbCrlf
	tmpHtml = tmpHtml & "		onChange:function(res){" & vbCrlf
	'tmpHtml = tmpHtml & "			console.log(res);" & vbCrlf
	tmpHtml = tmpHtml & "		}," & vbCrlf
	tmpHtml = tmpHtml & "		onOpen:function(e){" & vbCrlf			'打开时回调
	'tmpHtml = tmpHtml & "			console.log(e.config.items); return false;" & vbCrlf
	tmpHtml = tmpHtml & "		}," & vbCrlf
	tmpHtml = tmpHtml & "		onClose:function(){" & vbCrlf
	tmpHtml = tmpHtml & "			$.showLoading();" & vbCrlf 	' Load提示框
	tmpHtml = tmpHtml & "			var str1="""",cid = $(""#Course"").data(""values""), itemid = $(""#Item"").data(""values"");" & vbCrlf
	tmpHtml = tmpHtml & "			if(cid==0){return false;};" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#CourseID"").val(cid);" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "Substitute/getCourse.html"",{ID:cid, ItemID:itemid}, function(redata){" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#Student"").val(redata.Student); $(""#VA12"").val(redata.VA12); $(""#VA11"").val(redata.VA11);" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#Contents"").val(redata.Contents); $(""#CourseDate"").val(redata.CourseDate);" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#VA7"").val(redata.Period); $(""#VA3"").val(redata.VA3); $(""#VA5"").val(redata.VA5);" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#VA8"").val(redata.Course); $(""#VA6"").val(redata.VA6); " & vbCrlf

	tmpHtml = tmpHtml & "				str1 = ""<li><tt>授课日期：</tt><em>""+ redata.CourseDate +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>节次：</tt><em>""+ redata.Period +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>学时：</tt><em>""+ redata.VA3 +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>授课对象：</tt><em>""+ redata.Student +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>授课内容：</tt><em>""+ redata.Contents +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>校(院)区：</tt><em>""+ redata.VA11 +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				$("".old-box ul"").html(str1);$("".old-box"").show();" & vbCrlf
	tmpHtml = tmpHtml & "				$.hideLoading();" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		}" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	'===== 提交表单 =====
	tmpHtml = tmpHtml & "	$(""#subPost"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		let applyid = parseInt($(""#ApplyID"").val());" & vbCrlf
	tmpHtml = tmpHtml & "		let itemid = parseInt($(""#ItemID"").val());" & vbCrlf
	tmpHtml = tmpHtml & "		let courseid = parseInt($(""#CourseID"").val());" & vbCrlf
	tmpHtml = tmpHtml & "		if(applyid==0){ $.toptip('获取申请人员工代码失败！', 'error'); return false; }" & vbCrlf
	tmpHtml = tmpHtml & "		if(itemid==0){ $.toptip('请选择考核项目', 'error'); return false; }" & vbCrlf
	tmpHtml = tmpHtml & "		if(courseid==0){ $.toptip('请选择课程', 'error'); return false; }" & vbCrlf
	'tmpHtml = tmpHtml & "		console.log(courseid);" & vbCrlf
	tmpHtml = tmpHtml & "		location.href=""" & ParmPath & "Daike/Step2.html?tid="" + itemid + ""&cid="" + courseid + """";" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "</script>" & vbCrlf
	
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub
Sub Step2()		'第二步
	Dim tid : tid = HR_Clng(Request("tid"))
	Dim cid : cid = HR_Clng(Request("cid"))
	Dim arrItem : arrItem = GetTableDataQuery("HR_Class", "", 1, "ClassID=" & tid & "")		'// 取考核项目信息
	If Hr_Clng(arrItem(0, 1)) = 0 Then
		ErrMsg = "考核项目不存在"
		Response.Write GetErrBody(0) : Response.End
	End If
	Dim tTable : tTable = "HR_Sheet_" & tid
	Dim arrCourse : arrCourse = GetTableDataQuery(tTable, "", 1, "ID=" & cid & "")		'// 取课程信息
	If Hr_Clng(arrCourse(0, 1)) = 0 Then
		ErrMsg = "课程不存在"
		Response.Write GetErrBody(0) : Response.End
	End If
	If UserYGDM = 0 Then
		ErrMsg = "您的登陆已经失效"
		Response.Write GetErrBody(0) : Response.End
	End If

	SiteTitle = "第二步：选择方式"
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids {border-bottom:1px solid #ddd;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item {width:33.3%;box-sizing:border-box;text-align:center;padding:8px 0;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item em:first-child {width:50px;height:50px;line-height:50px;text-align:center;margin:0 auto;background-color:#f90;color:#fff;border-radius: 40px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item i {font-size:1.5rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item em {font-size:1rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item:nth-child(6n+1) em:first-child {background-color:#0bf;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item:nth-child(6n+2) em:first-child {background-color:#2da;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item:nth-child(6n+3) em:first-child {background-color:#b6c;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item:nth-child(6n+4) em:first-child {background-color:#ca4;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item:nth-child(6n+5) em:first-child {background-color:#5b6;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)
	Response.Write getHeadNav(0)

	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""hr-grids"">" & vbCrlf
	Response.Write "		<a class=""grid-item"" href=""" & ParmPath & "DaikeChangeTime/Index.html?tid=" & tid & "&cid=" & cid & """><em><i class=""hr-icon"">&#xe91d;</i></em><em>改授课时间</em></a>" & vbCrlf
	Response.Write "		<a class=""grid-item"" href=""" & ParmPath & "DaikeChangeTeacher/Index.html?tid=" & tid & "&cid=" & cid & """><em><i class=""hr-icon"">&#xf2f2;</i></em><em>换授课人</em></a>" & vbCrlf
	Response.Write "		<a class=""grid-item"" href=""" & ParmPath & "DaikeExchange/Index.html?tid=" & tid & "&cid=" & cid & """><em><i class=""hr-icon"">&#xf0be;</i></em><em>与别人换课</em></a>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<input name=""ItemID"" id=""ItemID"" type=""hidden"" value=""" & tid & """>" & vbCrlf
	Response.Write "	<input name=""CourseID"" id=""CourseID"" type=""hidden"" value=""" & cid & """>" & vbCrlf
	Response.Write "</form>" & vbCrlf


	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Daike/Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(1) : strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml

End Sub

Function GetSelectOptionItem()				'// 取考核项目下拉
	Dim iFun, funItem, rsFun, sqlFun
	sqlFun = "Select ClassID, ClassName From HR_Class Where ModuleID=1001 And Child=0 And Template='TempTableA'"
	sqlFun = sqlFun & " Order By RootID, OrderID"
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			iFun = 0
			Do While Not rsFun.EOF
				If iFun > 0 Then funItem = funItem & ","
				funItem = funItem & "{title:""" & rsFun("ClassName") & """,value:""" & rsFun("ClassID") & """}"
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		End If
	Set rsFun = Nothing
	GetSelectOptionItem = funItem
End Function
%>