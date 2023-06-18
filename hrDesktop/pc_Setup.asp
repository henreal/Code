<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incMD5.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="incCommon.asp"-->
<!--#include file="incPurview.asp"-->
<%
Dim scriptCtrl : SiteTitle = "修改密码"
Dim strParm, arrParm : strParm = Trim(Request("Parm")) : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "ModiPass" Call ModiPass()
	Case "SaveModify" Call SaveModify()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Call ChkInitPass()	'//检查初始密码
End Sub
Sub ModiPass()
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead("Index", 1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml) : strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml) : Response.Write strHtml

	tmpHtml = "<div class=""hr-workZones hr-shrink-x20"">" & vbCrlf
	tmpHtml = tmpHtml & "<fieldset class=""layui-elem-field site-demo-button"" style=""margin:5px;"">"
	tmpHtml = tmpHtml & "<legend>修改密码</legend>"
	
	tmpHtml = tmpHtml & "<form class=""layui-form layui-form-pane"" id=""FloatForm"" name=""FloatForm"" lay-filter=""FloatForm"" action="""">" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""layer-hr-box"" id=""editBody"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">原密码</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><input type=""password"" name=""old_pass"" value="""" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">新密码</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><input type=""password"" name=""new_pass"" value="""" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">再次确认密码</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><input type=""password"" name=""confirm"" value="""" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""hr-send"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""btn-group""><button type=""button"" class=""layui-btn layui-btn-radius"" id=""SendBtn"">提交</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf

	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</form>" & vbCrlf
	tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	Response.Write tmpHtml : tmpHtml = ""

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#SendBtn"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var load = layer.load(1);" & vbCrlf
	tmpHtml = tmpHtml & "		$.post(""SaveModify.html"", $(""#FloatForm"").serialize(),function(res){" & vbCrlf
	tmpHtml = tmpHtml & "			layer.close(load);" & vbCrlf
	tmpHtml = tmpHtml & "			layer.alert(res.errmsg,{icon:res.icon,btn:'关闭'},function(inx){" & vbCrlf
	tmpHtml = tmpHtml & "				if(!res.err){ location.reload(); }" & vbCrlf
	tmpHtml = tmpHtml & "				layer.close(inx);" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		},""json"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot("Index", 0) & "" & vbCrlf
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub
Sub SaveModify()
	Dim tNow : tNow = Now() : ErrMsg = "" : strTmp = ""
	Dim oldPass : oldPass = Trim(Request.Form("old_pass"))
	Dim newPass : newPass = Trim(Request.Form("new_pass"))
	Dim confPass : confPass = Trim(Request.Form("confirm"))

	Dim tTimes : tTimes = HR_Clng(Request.Cookies(Site_Sn)("TIMES"))
	Dim tErrTime : tErrTime = Trim(Request.Cookies(Site_Sn)("err_time"))
	Dim tDateDiff : tDateDiff = 0
	If isdate(tErrTime) Then tDateDiff = HR_Clng(DateDiff("s", Now(), tErrTime))	'//获取已经过去的时间

	Dim ArrField : ArrField = GetTableDataQuery("HR_Teacher", "", 1, "YGDM='" & UserYGDM & "'")			'取教师信息
	If IsValidPassword(newPass) = False Then ErrMsg = "新密码必须由6-22位字母、数字或!@#$%^&_-+等组成！"
	If newPass<>confPass Then ErrMsg = "两次密码不一致！"
	'Response.Cookies(Site_Sn)("TIMES") = 0
	'Response.Cookies(Site_Sn)("err_time") = ""

	If tDateDiff > 10 Then
		ErrMsg = "请" & HR_Clng(tDateDiff/60) & "分钟后再试"
	Else
		If HR_isNull(oldPass) Then
			ErrMsg = "请输入原密码！"
		Else
			If tTimes > 4 Then
				ErrMsg = "您的原密码已错误5次，请15分钟后再试"
				Response.Cookies(Site_Sn)("err_time") = DateAdd("s", 900, tNow)
			Else
				If MD5(oldPass,16)=Trim(ArrField(15,1)) Then
					Response.Cookies(Site_Sn)("TIMES") = 0
					Response.Cookies(Site_Sn)("err_time") = ""
				Else
					tTimes = tTimes+1
					Response.Cookies(Site_Sn)("TIMES") = tTimes
					ErrMsg = "原密码不正确！已错误" & tTimes & "次<br>您最多可以试5次"
				End If
			End If
		End If
	End If
	If Not(HR_IsNull(ErrMsg)) Then Response.Write "{""err"":true, ""errcode"":500, ""errmsg"":""" & ErrMsg & """, ""icon"":2, ""YGDM"":""" & UserYGDM & """}" : Exit Sub

	sql = "Update HR_Teacher Set LoginPass='" & MD5(newPass,16) & "' Where TeacherID=" & HR_Clng(ArrField(0, 1))
	Conn.Execute(sql)
	Response.Cookies(Site_Sn)("YGDM") = ""
	Response.Cookies(Site_Sn)("UserPass") = ""
	Response.Cookies(Site_Sn)("RndCode") = ""
	ErrMsg = "密码修改成功，请用新密码重新登陆"
	Response.Write "{""err"":false, ""errcode"":0, ""errmsg"":""" & ErrMsg & """, ""icon"":1}"
End Sub
%>