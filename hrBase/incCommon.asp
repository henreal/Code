<%
'***************** 后台管理通用函数 *****************
' Powered By：Henreal Studio
' Update：Henreal SMCS V1.0.23 Build 20170208
' Website：http://www.henreal.com
' Wechat：Henreal-Net【恒锐网络科技】
' Tel：0831-8239995 / 13700999995
'----------------------------------------------------

Dim ParmPath : ParmPath = InstallDir & ManageDir
Dim ModuleID, ClassID, UserFace
Dim MapLngLat : MapLngLat = XmlText("Common", "MapLngLat", "")
Dim apiHost : apiHost = "http://" & Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
Dim strINI : strINI = "[ssoAPI]" & vbCrlf & "ssurl=" & wmu2Api
Call WriteToFile("/sso/ssoconfig.ini", strINI, "UTF-8", 1)

'----- 获取超管UserID数组arrManager()
Dim rsManager
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

'=====================================================================
'函数名：getNavSubClass		【取侧导航子类】
'=====================================================================
Function getNavSubClass(fModuleID, fParentID, fType)
	Dim rsFun, strFun
	Set rsFun = Conn.Execute("Select * From HR_Class Where ModuleID=" & fModuleID & " And ParentID=" & HR_Clng(fParentID) & " Order By RootID, OrderID")
		If Not(rsFun.BOF And rsFun.EOF) Then
			strFun = "					<dl class=""layui-nav-child"">" & vbCrlf
			Do While Not rsFun.EOF
				strFun = strFun & "						<dd><a href=""javascript:;"" data-id=""" & rsFun("ClassID") & """>"
				If rsFun("NextID") > 0 Then
                    strFun = strFun & "├&nbsp;"
				Else
					strFun = strFun & "└&nbsp;"
				End If
				strFun = strFun & "" & rsFun("ClassName") & "</a></dd>" & vbCrlf
				rsFun.MoveNext
			Loop
			strFun = strFun & "					</dl>" & vbCrlf
		End If
	Set rsFun = Nothing
	getNavSubClass = strFun
End Function

Function getFrameNav(iType)
	Dim strFun
	strFun = "<header class=""hr-rows iframe-nav"">" & vbCrlf
	strFun = strFun & "	<nav class=""hr-row navPath""><i class=""hr-icon"">&#xf101;</i>我的位置：</nav>" & vbCrlf
	strFun = strFun & "	<nav class=""hr-row hr-row-fill""><hgroup class=""layui-breadcrumb"" lay-separator=""/""><a href=""[@Web_Dir][@Manage_Dir]Index/Start.html"">开始</a>[@Module_Path]</hgroup></nav>" & vbCrlf
	strFun = strFun & "	<nav class=""hr-row navBtn""><a href=""javascript:void(0);"" class=""navLayer""><i class=""hr-icon"">&#xf141;</i></a></nav>" & vbCrlf
	strFun = strFun & "</header>" & vbCrlf
	getFrameNav = strFun
End Function

Function getPageHead(iType)
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
	strFun = strFun & "	<script type=""text/javascript"" src=""[@Web_Dir]Static/js/jquery.min.js?v=1.11.2""></script>" & vbCrlf
	strFun = strFun & "	[@HeadStyle]" & vbCrlf
	strFun = strFun & "	[@HeadScript]" & vbCrlf
	strFun = strFun & "</head>" & vbCrlf
	strFun = strFun & "<body>" & vbCrlf
	If HR_Clng(iType) = 1 Then
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
		strFun = strFun & "		layui.use([""element"", ""layer""], function(){ var layer=layui.layer, element=layui.element; layer.config({moveOut:true,skin:""layer-hr""});});" & vbCrlf
		strFun = strFun & "	</script>"
		strFun = strFun & "[@HeadScript]" & vbCrlf
		strFun = strFun & "</head>" & vbCrlf
		strFun = strFun & "<body>" & vbCrlf
	End If
	getPageHead = strFun
End Function

Function getPageFoot(iType)
	Dim strFun
	strFun = "</body>" & vbCrlf
	strFun = strFun & "</html>" & vbCrlf
	strFun = strFun & "[@FootScript]" & vbCrlf
	getPageFoot = strFun
End Function

'=====================================================================
'函数名：GetErrBody		【取错误页】
'参数:type 0:默认显示错误信息/1:不显示导航路径/2:返回JSON数据
'=====================================================================
Function GetErrBody(fType)
	Dim funHtml, funStr
	If ErrMsg = "" Then ErrMsg = "非常遗憾，您访问的页面不存在！"
	funStr = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	funStr = funStr & "		body {background-color:#f1f1f1;}" & vbCrlf
	funStr = funStr & "		.hr-err-box {position:fixed;left:50%;top:15%;margin-left:-225px;}" & vbCrlf
	funStr = funStr & "	</style>"
	funHtml = getPageHead(1)
	funHtml = Replace(funHtml, "[@HeadStyle]", funStr)
	funHtml = Replace(funHtml, "[@HeadScript]", "")
	If HR_CLng(fType) = 0 Then
		funHtml = funHtml & getFrameNav(0)
		funHtml = Replace(funHtml, "[@Module_Path]", "<a><cite>错误提示</cite></a>")
	End If
	funHtml = funHtml & "<div class=""hr-err-box"">" & vbCrlf
	funHtml = funHtml & "	<div class=""hr-rows error"">" & vbCrlf
	funHtml = funHtml & "		<div class=""errorIcon""><i class=""hr-icon"">&#xeba5;</i></div>" & vbCrlf
	funHtml = funHtml & "		<div class=""errorInfo"">" & vbCrlf
	funHtml = funHtml & "			<h2>" & ErrMsg & "</h2>" & vbCrlf
	funHtml = funHtml & "			<p>您可以按提示操作或联系管理员以解决此问题!</p>" & vbCrlf
	funHtml = funHtml & "			<div class=""reindex""><a href=""" & ParmPath & "Index/Start.html"">返回首页</a><a href=""javascript:;"" class=""contact"">联系管理人员</a></div>" & vbCrlf
	funHtml = funHtml & "		</div>" & vbCrlf
	funHtml = funHtml & "	</div>" & vbCrlf
	funHtml = funHtml & "</div>" & vbCrlf
	funHtml = funHtml & getPageFoot(0)
	funHtml = Replace(funHtml, "[@FootScript]", "")
	funHtml = ReplaceCommonLabel(funHtml)
	If HR_CLng(fType) = 2 Then
		funHtml = "{""err"":false,""errcode"":500,""errmsg"":""" & ErrMsg & """,""icon"":0}"
	End If
	GetErrBody = funHtml
End Function

'=====================================================================
'函数名：GetTeacherOption		【取教师下拉】
'=====================================================================
Function GetTeacherOption(fTypeID, fTeacher)
	Dim rsFun, strFun : strFun = ""
	Set rsFun = Conn.Execute("Select * From HR_Teacher Where TeacherID>0 Order BY YGDM ASC")
		If Not(rsFun.BOF And rsFun.EOF) Then
			Do While Not rsFun.EOF
				strFun = strFun & "<option value=""" & rsFun("YGDM") & """"
				If Trim(fTeacher) = rsFun("YGDM") Then strFun = strFun & " selected"
				strFun = strFun & ">"
				strFun = strFun & rsFun("YGXM") & "</option>"
				rsFun.MoveNext
			Loop
		End If
	Set rsFun = Nothing
	GetTeacherOption = strFun
End Function
Function GetTeacherYGDMItems(fKSDM, fType)		'【根据科室名称或科室代码取员工代码组】
	Dim rsFun, sqlFun, iFun, strFun : strFun = ""
	sqlFun = "Select * From HR_Teacher Where TeacherID>0"
	If HR_IsNull(fKSDM) = False And HR_Clng(fKSDM) = 0 Then sqlFun = sqlFun & " And KSMC like '%" & fKSDM & "%'"
	If HR_Clng(fKSDM) > 0 Then sqlFun = sqlFun & " And KSDM=" & HR_Clng(fKSDM) & ""
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			iFun = 0
			Do While Not rsFun.EOF
				If iFun > 0 Then strFun = strFun & ","
				strFun = strFun & rsFun("YGDM")
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		End If
	Set rsFun = Nothing
	If HR_IsNull(strFun) = False Then strFun = FilterArrNull(strFun, ",")
	GetTeacherYGDMItems = strFun
End Function

'=====================================================================
'函数名：ChkUserName		【判断登陆帐号是否合法，返回布尔值】
'=====================================================================
Function ChkUserName(fName)
	regEx.Pattern ="^[A-Za-z0-9]{2,16}$"
	ChkUserName	 = False
	If HR_IsNull(fName) = False Then ChkUserName = regEx.Test(fName)
End Function

'=====================================================================
'函数名：ReplaceAPIStr		【替换接口中的特别字符】
'=====================================================================
Function ReplaceAPIStr(fStr)
	Dim strFun : strFun = ""
	If Trim(fStr) <> "" Then
		fStr = Replace(fStr, chr(9), "")
		fStr = Replace(fStr, chr(10), "")
		fStr = Replace(fStr, chr(13), "")
		If fStr <> "" Then strFun = fStr
	End If
	ReplaceAPIStr = strFun
End Function


'=====================================================================
'函数名：GetManagerID(fStuType, fAll)		【取管理员ID按学生类别】
'=====================================================================
Function GetManagerID(fStuType, fAll)
	Dim strFun, rsFun, iFun :strFun = ""
	If HR_IsNull(fStuType) = False Then
		Set rsFun = Conn.Execute("Select * From HR_User Where ManageRank=1")
			If Not(rsFun.BOF And rsFun.EOF) Then
				iFun = 0
				Do While Not rsFun.EOF
					If FoundInArr(Trim(rsFun("StuType")), fStuType, ",") Then
						strFun = strFun & rsFun("UserID") & ","
					End If
					rsFun.MoveNext
					iFun = iFun + 1
				Loop
				strFun = FilterArrNull(strFun, ",")
			End If
		Set rsFun = Nothing
	End If
	GetManagerID = strFun
End Function

%>