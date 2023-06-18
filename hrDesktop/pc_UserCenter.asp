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
SiteTitle = "会员中心"
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "ModifyPass" Call ModifyPass()
	Case "SavePass" Call SavePass()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Response.Write UserID & UserName
End Sub

Sub ModifyPass()
	

	tmpHtml = "<link type=""text/css"" href=""[@Web_Dir]Static/css/rb.common.css?v=1.0.3"" rel=""stylesheet"" media=""all"">" & vbCrlf
	tmpHtml = tmpHtml & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .layui-form-label {width:100px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .layui-input-inline {margin-left:8px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead("Desktop", 1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<form class=""layui-form"" id=""EditForm"" name=""EditForm"" lay-filter=""EditForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layui-form-item""><label class=""layui-form-label"">原密码：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""password"" name=""oldpass"" id=""oldpass"" value="""" lay-verify=""required"" placeholder=""请输入原密码"" class=""layui-input""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux morebtn""><span class=""tips"">必填项</span></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item""><label class=""layui-form-label"">新密码：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""password"" name=""newpass"" id=""newpass"" value="""" lay-verify=""required"" placeholder=""请输入新的密码"" class=""layui-input""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux morebtn""><span class=""tips"">新密码由16个字母或数字组成</span></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item""><label class=""layui-form-label"">密码确认：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""password"" name=""confpass"" id=""confpass"" value="""" lay-verify=""required"" placeholder=""再次输入新密码"" class=""layui-input""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux morebtn""><span class=""tips"">必填项</span></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""hr-pop-fix"">" & vbCrlf
	Response.Write "			<div class=""hr-grids hr-btn-group"">" & vbCrlf
	Response.Write "				<em><button type=""button"" class=""layui-btn hr-btn_deon"" id=""EditPost"" lay-filter=""EditPost"" lay-submit title=""保存""><i class=""hr-icon"">&#xf0c7;</i>保存</button></em>" & vbCrlf
	Response.Write "				<em><button type=""button"" class=""layui-btn layui-btn-primary"" id=""refresh"" data-event=""refresh"" title=""刷新""><i class=""hr-icon"">&#xf343;</i></button></em>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-place-h50""></div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#refresh"").on(""click"", function(){location.reload();});" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""form""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		form.on(""submit(EditPost)"", function(data){" & vbCrlf			'保存
	tmpHtml = tmpHtml & "			$.post(""" & ParmPath & "UserCenter/SavePass.html"",$(""#EditForm"").serialize(), function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.msg(reData.errmsg, {icon:reData.icon}, function(){" & vbCrlf
	tmpHtml = tmpHtml & "					if(!reData.err){" & vbCrlf
	tmpHtml = tmpHtml & "						var index1 = parent.layer.getFrameIndex(window.name);" & vbCrlf
	tmpHtml = tmpHtml & "						parent.location.reload();" & vbCrlf						'刷新
	tmpHtml = tmpHtml & "						parent.layer.close(index1);" & vbCrlf					'关闭[在iframe页面]
	tmpHtml = tmpHtml & "					}" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			},""json"");" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	
	strHtml = getPageFoot("Desktop", 1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub SavePass()
	Dim tmpJson, rsGet, sqlGet : ErrMsg = ""
	Dim OldPass : OldPass = Trim(Request("oldpass"))
	Dim NewPass : NewPass = Trim(Request("newpass"))
	Dim ConfPass : ConfPass = Trim(Request("confpass"))

	If HR_IsNull(OldPass) Then ErrMsg = "请输入原密码！"
	If HR_IsNull(NewPass) Then ErrMsg = "新密码不能为空！"
	If HR_IsNull(ConfPass) Then ErrMsg = "请您再次输入新密码！"
	If NewPass<>ConfPass Then ErrMsg = "您两次输入的新密码不一致！"
	If HR_IsNull(ErrMsg) = False Then Response.Write "{""err"":true,""icon"":2,""errcode"":500,""errmsg"":""" & ErrMsg & """}" : Exit Sub

	sqlGet = "Select Top 1 * From HR_Teacher Where LoginPass='" & MD5(OldPass, 16) & "' And YGDM='" & UserYGDM & "'"
	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 3
		If Not(rsGet.BOF And rsGet.EOF) Then
			rsGet("LoginPass") = MD5((NewPass), 16)
			rsGet.Update
			Response.Write "{""err"":false,""icon"":1,""errcode"":0,""errmsg"":""密码修改成功，请重新登陆！""}" : Exit Sub
		Else
			Response.Write "{""err"":true,""icon"":2,""errcode"":500,""errmsg"":""您输入的原密码不正确！""}" : Exit Sub
		End If
	Set rsGet = Nothing
End Sub
%>