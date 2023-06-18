<%
'======== PC版公用部分【顶部、底部、导航等】 ========
Dim ParmPath : ParmPath = InstallDir & "Desktop/"
Dim HeadFace, rsManager
'----- 获取超管UserID数组arrManager()
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

Function getPageHead(fModule, fType)
	Dim strFun
	strFun = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbCrlf
	strFun = strFun & "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrlf
	strFun = strFun & "<head>" & vbCrlf
	strFun = strFun & "	<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"" />" & vbCrlf
	strFun = strFun & "	<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />" & vbCrlf
	strFun = strFun & "	<title>[@Site_Title]</title>" & vbCrlf
	strFun = strFun & "	<link href=""[@Web_Dir]Static/Admin/css/style.css"" rel=""stylesheet"" type=""text/css"" />" & vbCrlf
	strFun = strFun & "	<!--[if IE]>" & vbCrlf
	strFun = strFun & "		<script src=""[@Web_Dir]Static/js/html5shiv.min.js""></script>" & vbCrlf
	strFun = strFun & "	<![endif]-->" & vbCrlf
	strFun = strFun & "	<script type=""text/javascript"" src=""[@Web_Dir]Static/js/jquery.min.js?v=1.11.2""></script>"
	strFun = strFun & "[@HeadStyle][@HeadScript]" & vbCrlf
	strFun = strFun & "</head>" & vbCrlf
	strFun = strFun & "<body>" & vbCrlf
	If HR_Clng(fType) = 1 Then
		strFun = "<!DOCTYPE html>" & vbCrlf
		strFun = strFun & "<html lang=""zh-cn"">" & vbCrlf
		strFun = strFun & "<head>" & vbCrlf
		strFun = strFun & "	<meta charset=""utf-8"" />" & vbCrlf
		strFun = strFun & "	<title>[@Site_Title]</title>" & vbCrlf
		strFun = strFun & "	<meta name=""renderer"" content=""webkit"" />" & vbCrlf
		strFun = strFun & "	<meta http-equiv=""X-UA-Compatible"" content=""IE=edge,chrome=1"" />" & vbCrlf
		strFun = strFun & "	<meta name=""viewport"" content=""width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=0, minimal-ui"">" & vbCrlf
		strFun = strFun & "	<link rel=""stylesheet"" type=""text/css"" href=""[@Web_Dir]Static/layui/css/layui.css?v=layui-v2.4.5"" />" & vbCrlf
		strFun = strFun & "	<link type=""text/css"" href=""[@Web_Dir]Static/css/hr.common.css?v=1.0.3"" rel=""stylesheet"" media=""all"">" & vbCrlf
		strFun = strFun & "	<link type=""text/css"" href=""[@Web_Dir]Static/css/rb.common.css?v=1.0.3"" rel=""stylesheet"" media=""all"">"
		strFun = strFun & "[@HeadStyle]" & vbCrlf
		strFun = strFun & "	<!--[if lt IE 9]>" & vbCrlf
		strFun = strFun & "		<script src=""https://cdn.staticfile.org/html5shiv/r29/html5.min.js""></script>" & vbCrlf
		strFun = strFun & "		<script src=""https://cdn.staticfile.org/respond.js/1.4.2/respond.min.js""></script>" & vbCrlf
		strFun = strFun & "	<![endif]-->" & vbCrlf
	
		strFun = strFun & "	<script type=""text/javascript"" src=""[@Web_Dir]Static/js/jquery.min.js?v=1.11.2""></script>" & vbCrlf
		strFun = strFun & "	<script type=""text/javascript"" src=""[@Web_Dir]Static/layui/layui.js?v=layui-v2.4.5""></script>" & vbCrlf
		strFun = strFun & "	<script type=""text/javascript"">" & vbCrlf
		strFun = strFun & "		$(document).ready(function(){  });" & vbCrlf
		strFun = strFun & "		layui.use([""element"", ""layer""], function(){ var layer=layui.layer, element=layui.element; layer.config({skin:""layer-hr""});});" & vbCrlf
		strFun = strFun & "	</script>"
		strFun = strFun & "[@HeadScript]" & vbCrlf
		strFun = strFun & "</head>" & vbCrlf
		strFun = strFun & "<body>" & vbCrlf
	End If
	getPageHead = strFun
End Function

Function getPageFoot(fModule, fType)
	Dim strFun
	strFun = "</body>" & vbCrlf
	strFun = strFun & "</html>" & vbCrlf
	strFun = strFun & "[@FootScript]" & vbCrlf
	getPageFoot = strFun
End Function

Function GetErrBody(fType)
	Dim funHtml, funStr
	fType = HR_Clng(fType)

	If HR_isNull(ErrMsg) Then ErrMsg = "非常遗憾，您访问的页面不存在！"
	If HR_isNull(strTmp) Then strTmp = "您可以按提示操作或联系管理员以解决此问题!"
	If HR_isNull(ErrHref) Then ErrHref = "" & ParmPath & "Index/Start.html"

	funStr = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	funStr = funStr & "		body {background-color:#f1f1f1;} .error{margin:0;width:100%;}" & vbCrlf
	funStr = funStr & "		.hr-err-box {position:fixed;left:50%;top:15%;margin-left:-225px;}" & vbCrlf
	funStr = funStr & "	</style>"

	funHtml = getPageHead("Index", 1)
	funHtml = Replace(funHtml, "[@HeadStyle]", funStr)
	funHtml = Replace(funHtml, "[@HeadScript]", "")

	funHtml = funHtml & "<div class=""hr-err-box"">" & vbCrlf
	funHtml = funHtml & "	<div class=""hr-rows error"">" & vbCrlf
	funHtml = funHtml & "		<div class=""errorIcon""><i class=""hr-icon"">&#xe811;</i></div>" & vbCrlf
	funHtml = funHtml & "		<div class=""errorInfo"">" & vbCrlf
	funHtml = funHtml & "			<h2>" & ErrMsg & "</h2>" & vbCrlf
	funHtml = funHtml & "			<p>" & strTmp & "</p>" & vbCrlf
	funHtml = funHtml & "			<div class=""reindex""><a href=""" & ErrHref & """>返回</a>"
	If fType = 1 Then funHtml = funHtml & "<a href=""javascript:;"">取消</a>"
	funHtml = funHtml & "</div>" & vbCrlf
	funHtml = funHtml & "		</div>" & vbCrlf
	funHtml = funHtml & "	</div>" & vbCrlf
	funHtml = funHtml & "</div>" & vbCrlf

	funHtml = funHtml & getPageFoot("Index", 0)
	funHtml = Replace(funHtml, "[@FootScript]", "")
	funHtml = ReplaceCommonLabel(funHtml)
	GetErrBody = funHtml
End Function

Function getFrameNav(iType)
	Dim strFun
	strFun = "<header class=""hr-rows iframe-nav"">" & vbCrlf
	strFun = strFun & "	<nav class=""navPath""><i class=""hr-icon"">&#xf101;</i>我的位置：</nav>" & vbCrlf
	strFun = strFun & "	<nav class=""hr-row-fill""><hgroup class=""layui-breadcrumb"" lay-separator=""/""><a href=""" & ParmPath & "Index/Start.html"">开始</a>[@Module_Path]</hgroup></nav>" & vbCrlf
	strFun = strFun & "	<nav class=""navBtn""><a href=""javascript:void(0);"" class=""navLayer""><i class=""hr-icon"">&#xf141;</i></a></nav>" & vbCrlf
	strFun = strFun & "</header>" & vbCrlf
	getFrameNav = strFun
End Function

%>