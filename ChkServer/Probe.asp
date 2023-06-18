<%@Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.Buffer = True

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim Url1 : Url1 = GetCurrentPath(2)

Response.Write MainHead()
Call SrvInfo()
Response.Write MainFoot()

Sub SrvInfo()
	Response.Write "<div id=""ShowPath""><dl><dt>当前位置：<a href=""http://www.henreal.com"" target=""_blank"">恒锐网络</a> &gt;&gt; 系统检测</dt><dd>&nbsp;</dd></dl></div>" & vbCrlf
	Response.Write "<div class=""BodyM1"">" & vbCrlf
	Response.Write "	<div class=""M1Tit"">" & vbCrlf
	Response.Write "		<div class=""M1On"">运行环境检测结果信息</div><div class=""M1Right"">&nbsp;</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""M1"">" & vbCrlf
	Response.Write "		<div class=""MyRemind""><ul>" & vbCrlf
	Dim getNetVer:getNetVer = GetHttpPage(Url1 & "Get.NetVer.aspx", 1)		'取ASP.Net版本
	Dim getPHPVer:getPHPVer = GetHttpPage(Url1 & "GetPHPVer.php", 1)		'取PHP版本

	Dim getJPEG, JpegTime
	If IsObjInstalled("Persits.Jpeg") Then
		Set getJPEG = Server.CreateObject("Persits.JPEG")
			JpegTime = getJPEG.Expires
		Set getJPEG = Nothing
	End If

	Response.Write "<li><b class=""Tit"">域　　名：</b>http://" & Request.ServerVariables("SERVER_NAME") & "　【IP：" & Request.ServerVariables("LOCAL_ADDR") & "】</li>"
	Response.Write "<li><b class=""Tit"">服务器端口：</b>" & Request.ServerVariables("Server_Port") & "　【脚本：" & GetCurrentPath(1) & "】</li>"
	Response.Write "<li><b class=""Tit"">IIS版本：</b>" & Request.ServerVariables("SERVER_SOFTWARE") & "</li>"
	Response.Write "<li><b class=""Tit"">脚本解释引擎：</b>" & ScriptEngine & "/" & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion & "  【<a href='http://www.microsoft.com/downloads/release.asp?ReleaseID=33136' target='_blank'>请点此更新</a>】</li>"
	Response.Write "<li><b class=""Tit"">物理路径：</b>" & Request.ServerVariables("APPL_PHYSICAL_PATH") & "</li>"
	Response.Write "<li><b class=""Tit"">ADO支持：</b>" & GetShowBit(IsObjInstalled("ADODB.Connection"), 0) & "</li>"
	Response.Write "<li><b class=""Tit"">FSO读写：</b>" & GetShowBit(IsObjInstalled("Scripting.FileSystemObject"), 0) & "</li>"
	Response.Write "<li><b class=""Tit"">数据流(无组件)读写：</b>" & GetShowBit(IsObjInstalled("ADODB.Stream"), 0) & "</li>"
	Response.Write "<li><b class=""Tit"">XMLHTTP支持：</b>" & GetShowBit(IsObjInstalled("Microsoft.XMLHTTP"), 0) & "</li>"
	Response.Write "<li><b class=""Tit"">XMLDOM支持：</b>" & GetShowBit(IsObjInstalled("Microsoft.XMLDOM"), 0) & "</li>"
	Response.Write "<li><b class=""Tit"">MSXML2.XMLHTTP支持：</b>" & GetShowBit(IsObjInstalled("MSXML2.XMLHTTP"), 0) & "</li>"
	Response.Write "<li><b class=""Tit"">MSXML2.XMLHTTP.3.0支持：</b>" & GetShowBit(IsObjInstalled("MSXML2.XMLHTTP.3.0"), 0) & "</li>"
	Response.Write "<li><b class=""Tit"">MSXML2.XMLHTTP.4.0支持：</b>" & GetShowBit(IsObjInstalled("MSXML2.XMLHTTP.4.0"), 0) & "</li>"
	Response.Write "<li><b class=""Tit"">MSXML2.XMLHTTP.5.0支持：</b>" & GetShowBit(IsObjInstalled("MSXML2.XMLHTTP.5.0"), 0) & "</li>"
	Response.Write "<li><b class=""Tit"">MSXML2.XMLHTTP.6.0支持：</b>" & GetShowBit(IsObjInstalled("MSXML2.XMLHTTP.6.0"), 0) & "</li>"
	Response.Write "<li><b class=""Tit"">ASP.Net支持：</b>Ver " & getNetVer & "</li>"
	Response.Write "<li><b class=""Tit"">PHP支持：</b>Ver " & getPHPVer & "  【<a href='phpinfo.php' target='_blank'>PHP Info</a>】</li>"
	Response.Write "<li><b class=""Tit"">AspJpeg支持：</b>" & GetShowBit(IsObjInstalled("Persits.Jpeg"), 0)  & "【到期时间：" & FormatDate(JpegTime, 4) & "】</li>"
	Response.Write "<li><b class=""Tit"">Jmail支持：</b>" & GetShowBit(IsObjInstalled("JMail.SMTPMail"), 0) & GetObjectVer("JMail.SMTPMail") & "</li>"
	Response.Write "<li><b class=""Tit"">CDONTS支持：</b>" & GetShowBit(IsObjInstalled("CDONTS.NewMail"), 0) & GetObjectVer("CDONTS.NewMail") & "</li>"
	Response.Write "<li><b class=""Tit"">AspEmail支持：</b>" & GetShowBit(IsObjInstalled("Persits.MailSender"), 0) & GetObjectVer("Persits.MailSender") & "</li>"
	Response.Write "<li><b class=""Tit"">AspUpload支持：</b>" & GetShowBit(IsObjInstalled("Persits.Upload"), 0) & GetObjectVer("Persits.Upload") & "</li>"
	Response.Write "<li><b class=""Tit"">SA-FileUp支持：</b>" & GetShowBit(IsObjInstalled("SoftArtisans.FileUp"), 0) & GetObjectVer("SoftArtisans.FileUp") & "</li>"
	Response.Write "<li><b class=""Tit"">DvFile-Up支持：</b>" & GetShowBit(IsObjInstalled("DvFile.Upload"), 0) & GetObjectVer("DvFile.Upload") & "</li>"
	Response.Write "<li><b class=""Tit"">CreatePreviewImage支持：</b>" & GetShowBit(IsObjInstalled("CreatePreviewImage.cGvbox"), 0) & GetObjectVer("CreatePreviewImage.cGvbox") & "</li>"
	Response.Write "<li><b class=""Tit"">SA-ImgWriter支持：</b>" & GetShowBit(IsObjInstalled("SoftArtisans.ImageGen"), 0) & GetObjectVer("SoftArtisans.ImageGen") & "</li>"
	
	Response.Write "<li><b class=""Tit"">W3Image.Image支持：</b>" & GetShowBit(IsObjInstalled("Persits.MailSender"), 0) & GetObjectVer("Persits.MailSender") & "</li>"
	Response.Write "<li><b class=""Tit"">动易组件版本：</b>" & GetShowBit(IsObjInstalled("PE_Common6.GetVersion"), 0) & GetObjectVer("PE_Common6.GetVersion") & "</li>"
	Response.Write "<li><b class=""Tit"">URL Rewrite Check：</b><a href=""URLRewrite.asp?ID=123456"" target=""_blank"">检查</a></li>"
	Response.Write "		</ul></div>" & vbCrlf
	Response.Write "		<div class=""Info""><a href=""http://www.henreal.com/Soft/Probe/Index.html"" target=""_blank"">Henreal</a> Probe V1.09.21 Build 20170921</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div><br />" & vbCrlf
End Sub


'==================================================
'函数名：IsObjInstalled		作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：布尔值
'==================================================
Function IsObjInstalled(strClassString)
    On Error Resume Next
    IsObjInstalled = False
    Err = 0
    Dim xTestObj
    Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
    Set xTestObj = Nothing
    Err = 0
End Function

'==================================================
'函数名：GetHttpPage		作  用：获取网页源码
'参  数：HttpUrl ------网页地址
'==================================================
Function GetHttpPage(HttpUrl, Coding)
    On Error Resume Next
    If IsNull(HttpUrl) = True Or Len(HttpUrl) < 18 Or HttpUrl = "" Then
        GetHttpPage = ""
        Exit Function
    End If
    Dim Http
	Set Http = Server.CreateObject("MSXML2.XMLHTTP")
	'Set Http = Server.CreateObject("MSXML2.XMLHTTP.3.0")
    Http.Open "GET", HttpUrl, False
    Http.Send
    If Http.Readystate <> 4 Then
        GetHttpPage = ""
        Exit Function
    End If
    If Coding = 1 Then
        GetHttpPage = BytesToBstr(Http.ResponseBody, "UTF-8")
    ElseIf Coding = 2 Then
        GetHttpPage = BytesToBstr(Http.ResponseBody, "Big5")
    Else
        GetHttpPage = BytesToBstr(Http.ResponseBody, "GB2312")
    End If
    
    Set Http = Nothing
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Function

'==================================================
'函数名：BytesToBstr
'作  用：将获取的源码转换为中文
'参  数：Body ------要转换的变量
'参  数：Cset ------要转换的类型
'==================================================
Function BytesToBstr(Body, Cset)
   Dim Objstream
   Set Objstream = Server.CreateObject("Adodb.Stream")
   objstream.Type = 1
   objstream.Mode = 3
   objstream.Open
   objstream.Write body
   objstream.Position = 0
   objstream.Type = 2
   objstream.Charset = Cset
   BytesToBstr = objstream.ReadText 
   objstream.Close
   set objstream = nothing
End Function

'==================================================
'函数名：GetObjectVer		作  用：取OBJ版本号
'==================================================
Function GetObjectVer(strObjName)
	Dim TestObj, TestVer
	If IsObjInstalled(strObjName) Then
		Set TestObj = Server.CreateObject(strObjName)
		If strObjName = "PE_Common6.GetVersion" Then
			TestVer = TestObj.strVersion
		Else
			TestVer = TestObj.version
		End If
		GetObjectVer = "（版本：<b class=""b2"">" & TestVer & "</b>）"
	End If
End Function

'==================================================
'函数名：GetCurrentPath		作  用：返回当前路径
'1：http://www.***.com/***/***.***
'2：http://www.***.com/***/
'3：/***/
'0：/***/***.***
'==================================================
Function GetCurrentPath(fType)
    Dim fPath:fPath = LCase(Request.ServerVariables("URL"))
	Dim fPORT:fPORT = Request.ServerVariables("SERVER_PORT")
	Dim rUrlLen:rUrlLen = InstrRev(fPath, "/")
	Dim lUrl:lUrl = Left(fPath, rUrlLen)
	Dim fFile:fFile = Mid(fPath, rUrlLen + 1)
	If HR_CLng(fPORT) > 0 And HR_CLng(fPORT) <> 80 Then
		fPath = ":" & fPORT & fPath
		lUrl = ":" & fPORT & lUrl
	End If
	
    Select Case fType
		Case 1
			fPath = "http://" & Request.ServerVariables("Server_Name") & fPath
		Case 2
			fPath = "http://" & Request.ServerVariables("Server_Name") & lUrl
		Case 3
			fPath = lUrl
    End Select
	GetCurrentPath = fPath
End Function

Function FormatDate(DateAndTime, para)
	On Error Resume Next
	Dim y, m, d, h, mi, s, strDateTime
	FormatDate = DateAndTime
	If Not IsNumeric(para) Then Exit Function
	If Not IsDate(DateAndTime) Then Exit Function
	y = CStr(Year(DateAndTime))
	m = CStr(Month(DateAndTime))
	If Len(m) = 1 Then m = "0" & m
	d = CStr(Day(DateAndTime))
	If Len(d) = 1 Then d = "0" & d
	h = CStr(Hour(DateAndTime))
	If Len(h) = 1 Then h = "0" & h
	mi = CStr(Minute(DateAndTime))
	If Len(mi) = 1 Then mi = "0" & mi
	s = CStr(Second(DateAndTime))
	If Len(s) = 1 Then s = "0" & s
	Select Case para
		Case "1" strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
		Case "2" strDateTime = y & "-" & m & "-" & d
		Case "3" strDateTime = y & "/" & m & "/" & d
		Case "4" strDateTime = y & "年" & m & "月" & d & "日"
		Case "5" strDateTime = m & "-" & d
		Case "6" strDateTime = m & "/" & d
		Case "7" strDateTime = m & "月" & d & "日"
		Case "8" strDateTime = y & "年" & m & "月"
		Case "9" strDateTime = y & "-" & m
		Case "10" strDateTime = y & "/" & m
		Case "11" strDateTime = y  & m & d
		Case "12" strDateTime = y  & m & d & h & mi & s
		Case "13" strDateTime = y & m
		Case "14" strDateTime = y
		Case Else strDateTime = DateAndTime
	End Select
	FormatDate = strDateTime
End Function

Function HR_CBool(strBool)
	HR_CBool = False
    If strBool = True Or LCase(Trim(strBool)) = "true" Or LCase(Trim(strBool)) = "yes" Or Trim(strBool) = "1" Then HR_CBool = True
End Function
Function HR_CLng(ByVal str1)
	HR_CLng = 0 : If IsNumeric(str1) Then HR_CLng = CLng(str1)
End Function

Function GetShowBit(vBit, ShowType)
	Dim strTmp
	If ShowType = 1 Then
		strTmp = "<b class=""ShowTrue"">否</b>"
		If HR_CBool(vBit) Then strTmp = "<b class=""ShowFalse"">是</b>"
	Else
		strTmp = "<b class=""ShowFalse"">×</b>"
		If HR_CBool(vBit) Then strTmp = "<b class=""ShowTrue"">√</b>"
	End If
	GetShowBit = strTmp
End Function

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
	fTmp = fTmp & "#ShowPath {margin:0 auto;width:95%;height:50px;line-height:50px;font-size:13px;overflow:hidden;}#ShowPath dt {height:50px;line-height:50px;}#ShowPath dd {overflow:hidden;display:none;}"
	fTmp = fTmp & ".M1Tit{height:30px;line-height:30px;color:#F00;overflow:hidden;text-align:center;background:#CCC;font-size:16px;}b.Tit{color:#555;}.MyRemind li {color:#07B;}"
	fTmp = fTmp & ".Info{text-align:right;padding:0 20px 0 0;color:#777;}"
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