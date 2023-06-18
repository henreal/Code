<!--#include file="../Lead.asp"-->
<!--#include file="../Custom/incPublic.asp"-->
<!--#include file="../Custom/incKernel.asp"-->
<!--#include file="../Custom/incMD5.asp"-->
<%
Dim Password, isLong
Password = Trim(Request("Word"))
isLong = HR_CBool(Request("Long"))


Response.Write MainHead()
Response.Write "<p></p><div class=""BodyM1"">" & vbCrlf
Response.Write "	<div class=""M1"">" & vbCrlf
If isLong Then
	Response.Write Password & "：" & MD5(Password, 32)
Else
	Response.Write Password & "：" & MD5(Password, 16)
End If
Response.Write "<form action="""" method=""post"" name=""getForm"">" & vbCrlf
Response.Write "	<input type=""text"" name=""Word"" placeholder=""请输入密码"" value="""" /><input type=""submit"" name=""submit"" value=""提交"" /><br />" & vbCrlf
Response.Write "	<input type=""checkbox"" name=""Long"" value=""True"" />32位" & vbCrlf
Response.Write "</form>" & vbCrlf
Response.Write "	</div>" & vbCrlf
Response.Write "</div><br />" & vbCrlf
Response.Write MainFoot()

Function MainHead()
	'参数：iType_5:生成HTML
	Dim fTmp
	fTmp = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbCrlf
	fTmp = fTmp & "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrlf
	fTmp = fTmp & "<head>" & vbCrlf
	fTmp = fTmp & "	<title>系统信息</title>" & vbCrlf
	fTmp = fTmp & "	<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"" />" & vbCrlf
	fTmp = fTmp & "	<meta http-equiv=""Content-Language"" content=""zh-CN"" />" & vbCrlf
	fTmp = fTmp & "	<meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0"" />" & vbCrlf
	fTmp = fTmp & "	<meta name=""robots"" content=""none"" />" & vbCrlf
	fTmp = fTmp & "	<style>body{margin:0;padding:0;font-size:14px;font-family:微软雅黑;background:#F2F2F2;}dl,dt,dd{margin:0;padding:0;}.BodyM1{clear:both;margin:0 auto;padding:0px;width:95%;border:1px solid #CCC;background:#FFF;line-height:2em;}"
	fTmp = fTmp & ".M1{padding:20px;}"
	fTmp = fTmp & "	</style>" & vbCrlf
	fTmp = fTmp & "</head>" & vbCrlf
	fTmp = fTmp & "<body>" & vbCrlf
	MainHead = fTmp
End Function

Function MainFoot()
	Dim fTmp
	fTmp = "</body>" & vbCrlf
	fTmp = fTmp & "</html>" & vbCrlf
	MainFoot = fTmp
End Function

%>