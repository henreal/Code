<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incMD5.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl : SiteTitle = "个人中心"

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "Message" Call Message()
	Case "viewMSG" Call viewMSG()
	Case "Pass" Call Pass()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	SiteTitle = "我的信息"
	Dim arrSex : arrSex = Split(XmlText("Config", "Sex", ""), "|")

	Dim tKSMC, tPRZC, tYGXB, tXZZW, tYGZT
	Set rsTmp = Conn.Execute("Select * From HR_Teacher Where YGDM='" & UserYGDM & "'")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tKSMC = Trim(rsTmp("KSMC"))
			tYGXB = Trim(rsTmp("YGXB"))
			tPRZC = Trim(rsTmp("PRZC"))
			tXZZW = Trim(rsTmp("XZZW"))
			tYGZT = Trim(rsTmp("YGZT"))
		End If
	Set rsTmp = Nothing

	Dim tEmail, tMobile, tGender, tName_EN, tQR_code
	Dim jsonOBJ, strJson : strJson = GetWechatUserInfoQY(UserYGDM)
	Set jsonOBJ = parseJSON(strJson)
		If jsonOBJ.errcode = 0 Then
			If HR_IsNull(HeadFace) Then HeadFace = Trim(jsonOBJ.avatar)
			tEmail = Trim(jsonOBJ.email)
			tMobile = Trim(jsonOBJ.mobile)
			tGender = arrSex(HR_Clng(jsonOBJ.gender))
			tName_EN = Trim(jsonOBJ.english_name)
			tQR_code = Trim(jsonOBJ.qr_code)
		End If
	Set jsonOBJ = Nothing
	If HR_IsNull(HeadFace) Then HeadFace = InstallDir & "Static/images/nopic.png"
	If HR_IsNull(tQR_code) Then
		tQR_code = "<i class=""hr-icon"">&#xec29;</i>"
	Else
		tQR_code = "<em class=""hr-qrcode""><img src=""" & tQR_code & """></em>"
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background-color:#eee;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
	Response.Write getHeadNav(0)

	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""weui-form-preview"">" & vbCrlf
	Response.Write "	<div class=""weui-form-preview__hd"">" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item""><label class=""weui-form-preview__label"">姓名</label><em class=""weui-form-preview__value"">" & UserYGXM & "</em></div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-form-preview__hd"">" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item""><label class=""weui-form-preview__label"">工号</label><em class=""weui-form-preview__value"">" & UserYGDM & "</em></div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf
	Response.Write "	<div class=""weui-form-preview__bd"">" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item"">" & vbCrlf
	Response.Write "			<label class=""weui-form-preview__label"">科室</label>" & vbCrlf
	Response.Write "			<span class=""weui-form-preview__value"">" & tKSMC & "</span>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item"">" & vbCrlf
	Response.Write "			<label class=""weui-form-preview__label"">职称</label>" & vbCrlf
	Response.Write "			<span class=""weui-form-preview__value"">" & tPRZC & "</span>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item"">" & vbCrlf
	Response.Write "			<label class=""weui-form-preview__label"">职务</label>" & vbCrlf
	Response.Write "			<span class=""weui-form-preview__value"">" & tXZZW & "</span>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item"">" & vbCrlf
	Response.Write "			<label class=""weui-form-preview__label"">性别</label>" & vbCrlf
	Response.Write "			<span class=""weui-form-preview__value"">" & tYGXB & "</span>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item"">" & vbCrlf
	Response.Write "			<label class=""weui-form-preview__label"">状态</label>" & vbCrlf
	Response.Write "			<span class=""weui-form-preview__value"">" & tYGZT & "</span>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf
	Response.Write "	<div class=""weui-form-preview__bd"">" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item"">" & vbCrlf
	Response.Write "			<label class=""weui-form-preview__label"">邮箱</label>" & vbCrlf
	Response.Write "			<span class=""weui-form-preview__value"">" & tEmail & "</span>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item"">" & vbCrlf
	Response.Write "			<label class=""weui-form-preview__label"">手机</label>" & vbCrlf
	Response.Write "			<span class=""weui-form-preview__value"">" & tMobile & "</span>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item"">" & vbCrlf
	Response.Write "			<label class=""weui-form-preview__label"">性别</label>" & vbCrlf
	Response.Write "			<span class=""weui-form-preview__value"">" & tGender & "</span>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item"">" & vbCrlf
	Response.Write "			<label class=""weui-form-preview__label"">姓名拼音</label>" & vbCrlf
	Response.Write "			<span class=""weui-form-preview__value"">" & tName_EN & "</span>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item"">" & vbCrlf
	Response.Write "			<label class=""weui-form-preview__label"">我的二维码</label>" & vbCrlf
	Response.Write "			<span class=""weui-form-preview__value"">" & tQR_code & "</span>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""weui-form-preview__ft"">" & vbCrlf
	Response.Write "		<a class=""weui-form-preview__btn weui-form-preview__btn_primary"" href=""" & ParmPath & "/Index.html"">返　回</a>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "/Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".addNew a"").css(""display"",""none"");" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub

Sub Message()
	SiteTitle = "我的消息"

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-cell__hd {padding-right:5px;font-size:1.4rem;} .weui-cell {align-items:stretch;}" & vbCrlf
	tmpHtml = tmpHtml & "		.unread .weui-cell__hd {color:#f30;} .unread h3 {color:#f30;} .viewMSG h5 {color:#999;font-size:0.9rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-dialog__hd {padding:8px 0;border-bottom:1px solid #eee;color:#f30}" & vbCrlf	'预览窗口
	tmpHtml = tmpHtml & "		.weui-dialog__bd {padding:8px;font-size:1.2rem;color:#000;text-align:left;min-height:20rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-dialog {max-width:initial;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
	Response.Write getHeadNav(0)

	Dim newMsgNum : newMsgNum = 0
	Set rsTmp = Conn.Execute("Select count(ID) From HR_Message Where isRead=" & HR_False & " And ReceiverID=" & HR_CLng(UserYGDM))
		newMsgNum = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing

	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""weui-form-preview"">" & vbCrlf
	Response.Write "	<div class=""weui-form-preview__hd"">" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item""><label class=""weui-form-preview__label"">新消息</label><em class=""weui-form-preview__value"">" & newMsgNum & "</em></div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""weui-cells__title"">全部消息</div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf

	Dim tmpStyle, tmpIcon, tSheetName, tVA4
	sqlTmp = "Select * From HR_Message Where ReceiverID=" & HR_CLng(UserYGDM)
	sqlTmp = sqlTmp & " Order By isRead ASC, SendTime DESC"
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Do While Not rsTmp.EOF
				tmpStyle = " unread" : tmpIcon = "&#xea11;"
				If HR_CBool(rsTmp("isRead")) Then tmpStyle = "" : tmpIcon = "&#xea12;"
				Response.Write "	<a class=""weui-cell weui-cell_access" & tmpStyle & """ href=""javascript:;"">" & vbCrlf
				Response.Write "		<div class=""weui-cell__hd""><i class=""hr-icon"">" & tmpIcon & "</i></div>" & vbCrlf
				Response.Write "		<div class=""weui-cell__bd viewMSG"" data-id=""" & rsTmp("ID") & """><h3>" & rsTmp("Title") & "</h3>"
				tSheetName = "HR_Sheet_" & HR_CLng(rsTmp("ItemID"))
				If rsTmp("ItemID")>0 And rsTmp("CourseID")>0 And ChkTable(tSheetName) Then
					sql = "Select a.*,(Select ClassName From HR_Class Where ClassID=" & rsTmp("ItemID") & ") As ItemName"
					sql = sql & ",(Select Template From HR_Class Where ClassID=" & rsTmp("ItemID") & ") As Template"
					sql = sql & " From " & tSheetName & " a Where a.ID=" & rsTmp("CourseID")
					Set rs = Conn.Execute(sql)
						If Not(rs.BOF And rs.EOF) Then
							tVA4 = "学年：" & Trim(rs("VA4"))
							If rs("Template") = "TempTableA" Or rs("Template") = "TempTableC" Or rs("Template") = "TempTableD" Or rs("Template") = "TempTableE" Then
								tVA4 = "日期：" & FormatDate(ConvertNumDate(tVA4), 4)
							End If
							Response.Write "<h4>考核项目：" & rs("ItemName") & "</h4>"
							Response.Write "<h4>" & tVA4 & "</h4>"
						End If
					Set rs = Nothing
				End If
				Response.Write "<h5>发送时间：" & FormatDate(rsTmp("SendTime"), 10) & "</h5></div>" & vbCrlf
				Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
				Response.Write "	</a>" & vbCrlf
				rsTmp.MoveNext
			Loop
		Else
			Response.Write "	<a class=""weui-cell weui-cell_access"" href=""javascript:;"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>暂时没有消息</p></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
			Response.Write "	</a>" & vbCrlf
		End If
	Set rsTmp = Nothing

	Response.Write "</div>" & vbCrlf
	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".addNew a"").css(""display"",""none"");" & vbCrlf
	tmpHtml = tmpHtml & "	$("".viewMSG"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		$.get(""" & ParmPath & "myCenter/viewMSG.html"",{id:$(this).data(""id"")}, function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "			$.alert(rsStr, ""查看消息"",function(){ location.reload(); });" & vbCrlf
	tmpHtml = tmpHtml & "			$("".ShowCourse"").css(""display"",""none"");" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub
Sub viewMSG()
	Dim tmpID : tmpID = HR_Clng(Request("id"))
	Dim tSheetName, tVA4, tMessage : ErrMsg = "" : strTmp = ""
	If tmpID > 0 Then
		Set rsTmp = Conn.Execute("Select * From HR_Message Where ID=" & tmpID )
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				tSheetName = "HR_Sheet_" & HR_CLng(rsTmp("ItemID"))
				If rsTmp("ItemID")>0 And rsTmp("CourseID")>0 And ChkTable(tSheetName) Then
					sql = "Select a.*,(Select ClassName From HR_Class Where ClassID=" & rsTmp("ItemID") & ") As ItemName"
					sql = sql & ",(Select Template From HR_Class Where ClassID=" & rsTmp("ItemID") & ") As Template"
					sql = sql & " From " & tSheetName & " a Where a.ID=" & rsTmp("CourseID")
					Set rs = Conn.Execute(sql)
						If Not(rs.BOF And rs.EOF) Then
							tVA4 = "学年：" & Trim(rs("VA4"))
							If rs("Template") = "TempTableA" Or rs("Template") = "TempTableC" Or rs("Template") = "TempTableD" Or rs("Template") = "TempTableE" Then
								tVA4 = "日期：" & FormatDate(ConvertNumDate(tVA4), 4)
							End If
							strTmp = strTmp & "<h4>考核项目：" & rs("ItemName") & "</h4>"
							strTmp = strTmp & "<h4>" & tVA4 & "</h4>"
						End If
					Set rs = Nothing
				End If
				tMessage = Trim(rsTmp("Message"))
				If HR_IsNull(tMessage) = False Then
					tMessage = Replace(tMessage, "Manage/Course", "Touch/Course/List")
				End If
				If HR_CBool(rsTmp("isRead")) = False Then Conn.Execute("Update HR_Message Set isRead=" & HR_True & ",ReadTime=getdate() Where ID=" & rsTmp("ID"))
				strTmp = strTmp & tMessage
			End If
		Set rsTmp = Nothing
	End If
	Response.Write strTmp
End Sub

Sub Pass()
	SiteTitle = "修改密码"
	If Ubound(arrParm) > 1 Then
		Select Case Trim(arrParm(2))
			Case "SavePass" Call SavePass()
		End Select
		Exit Sub
	End If
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-header {z-index:8;} .weui-toast {margin-left: auto;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-cells_form .weui-cell__ft {font-size:1.5rem;color:#f30;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
	Response.Write getHeadNav(0)

	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label for="""" class=""weui-label"">原密码：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><input class=""weui-input"" name=""oldpass"" id=""oldpass"" type=""password"" value="""" placeholder=""请输入原密码""></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft weui-cell_warn""><i class=""hr-icon"">&#xe947;</i></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label for="""" class=""weui-label"">新密码：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><input class=""weui-input"" name=""newpass"" id=""newpass"" type=""password"" value="""" placeholder=""请输入新密码""></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft weui-cell_warn""><i class=""hr-icon"">&#xe947;</i></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label for="""" class=""weui-label"">新密码确认：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><input class=""weui-input"" name=""newpass1"" id=""newpass1"" type=""password"" value="""" placeholder=""请再次输入新密码""></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft weui-cell_warn""><i class=""hr-icon"">&#xe947;</i></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-btn-area"">" & vbCrlf
	Response.Write "		<button class=""weui-btn weui-btn_primary"" type=""button"" name=""SendForm"" id=""SendForm"">修改密码</button>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "/Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#SendForm"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var oldpass = $(""#oldpass"").val(), newpass = $(""#newpass"").val(), newpass1 = $(""#newpass1"").val();" & vbCrlf
	tmpHtml = tmpHtml & "		if(oldpass==""""){ $.toast(""请输入原密码！"", ""cancel"", function(){ return false; }); }" & vbCrlf
	tmpHtml = tmpHtml & "		if(newpass==""""){ $.toast(""新密码必须填写！"", ""cancel"", function(){ return false; }); }" & vbCrlf
	tmpHtml = tmpHtml & "		if(newpass!=newpass1){ $.toast(""您两次输入的新密码不一致！"", ""cancel"", function(){ return false; }); }" & vbCrlf
	tmpHtml = tmpHtml & "		$.post(""" & ParmPath & "/myCenter/Pass/SavePass.html"", {FormerlyPass:oldpass, NewPass:newpass}, function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "			if(rsStr.Return){" & vbCrlf
	tmpHtml = tmpHtml & "				$.toast(rsStr.reMessge, function(){ location.reload(); });" & vbCrlf
	tmpHtml = tmpHtml & "			}else{" & vbCrlf
	tmpHtml = tmpHtml & "				$.toast(rsStr.reMessge, ""cancel"", function(){ return false; });" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		},""json"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub
Sub SavePass()
	Dim FormerlyPass : FormerlyPass = Trim(Request("FormerlyPass"))
	Dim NewPass : NewPass = Trim(Request("NewPass"))
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select LoginPass From HR_Teacher Where YGDM='" & UserYGDM & "'"), Conn, 1, 3
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			If MD5(FormerlyPass, 16) = rsTmp("LoginPass") Then
				rsTmp("LoginPass") = MD5(NewPass, 16)
				rsTmp.Update
				ErrMsg = "密码修改成功！"
				Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
			Else
				ErrMsg = "原密码不正确！"
				Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """}"
			End If
		End If
	Set rsTmp = Nothing
	
End Sub
%>