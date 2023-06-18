<%@Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Response.CodePage = 65001
Response.CharSet = "UTF-8"					'根据根据设置：GB2312|UTF-8|ISO-8859-1，请与编码申明一致，如：GB2312则为CodePage="936"；
Response.Buffer = True

Dim GetID : GetID = Request("ID")

Call ShowHead()

	Response.Write "<div class=""ShowRequest"">ID：" & GetID & "<br /><a href=""URLRewrite-" & GetID & ".html"" target=""_blank"">打开重写地址</a></div>" & vbCrlf
	
	Response.Write "<div class=""Tips""><ul><li>" & GetID & "</li>" & vbCrlf
	Response.Write "		<li>" & GetID & "</li>" & vbCrlf
	Response.Write "	</ul>" & vbCrlf
	Response.Write "</div>" & vbCrlf

Call ShowFoot()

Sub ShowHead()
	Response.Write "<!doctype html>" & vbCrlf
	Response.Write "<html>" & vbCrlf
	Response.Write "<head>" & vbCrlf
	Response.Write "	<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" & vbCrlf
	Response.Write "	<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />" & vbCrlf
	Response.Write "	<meta name=""viewport"" content=""initial-scale=1,maximum-scale=1,minimum-scale=1"" />" & vbCrlf
	Response.Write "	<meta name=""viewport"" content=""width=device-width, initial-scale=1, user-scalable=no"" />" & vbCrlf
	Response.Write "	<title>URL Rewrite TEST</title>" & vbCrlf
	Response.Write "	<style type=""text/css"">" & vbCrlf
	Response.Write "		body {margin:0 auto;width:100%;height:100%}" & vbCrlf
	Response.Write "		.ShowRequest {display:block;text-align:center;padding:5px;}" & vbCrlf
	Response.Write "		.GetImg img {width:160px;border:1px solid #888;}" & vbCrlf
	Response.Write "		.Gap {display:block;clear:both;margin:10px auto;line-height:1.5rem;width:90%;}" & vbCrlf
	Response.Write "		#ShowSchedule {background:url(/Images/LoadC.gif) right center no-repeat;display:block;width:90%;line-height:30px;margin:10px auto;color:red;}" & vbCrlf
	Response.Write "	</style>" & vbCrlf
	Response.Write "	<script type=""text/javascript"" src=""/JS/jquery.min.js""></script>" & vbCrlf
	Response.Write "</head>" & vbCrlf
	Response.Write "<body>" & vbCrlf
End Sub

Sub ShowFoot()
	Response.Write "" & vbCrlf
	Response.Write "" & vbCrlf
	Response.Write "</body>" & vbCrlf
	Response.Write "</html>" & vbCrlf
End Sub

%>