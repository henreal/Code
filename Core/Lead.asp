<%@Language="VBScript" CodePage="65001" %>
<%
Option Explicit
Response.CodePage = 65001
Response.CharSet = "UTF-8"					'根据根据设置：GB2312|UTF-8|ISO-8859-1，请与编码申明一致，如：GB2312则为CodePage="936"；
Response.Buffer = True

Const DataType = "SQL"
Const DBFileName = "/Core/Data/WMU2$2018.mdb"
Const XMLDataPath = "/Core/Data/Param.xml"			'字典数据
Const SqlUsername = "Henreal"							'SQL数据库用户名
Const SqlPassword = "hrsql995"						'SQL数据库用户密码
Const SqlDatabaseName = "WMU2_V2"						'SQL数据库名
Const SqlHostIP = "(local)"							'SQL主机IP地址。可用“(local)”或“127.0.0.1”
Const wmu2Api = "http://wap.wzhealth.com/feyservice/interface.asmx?wsdl"	'API接口外网地址
'Const wmu2Api = "http://10.20.1.73:83/feyservice/interface.asmx?wsdl"	'API接口内网地址

'----- 基本参数设置
Const IsChkRndCode = False						'是否检查多人异地登陆

Dim Conn, BeginTime, UserTrueIP, ScriptName, Site_Sn, InstallDir, ManageDir, UploadDir, strInstallDir, ComeUrl, Action, SoType, SoField, SoWord, ConfigID
Dim HR_True, HR_False, HR_Now, HR_OrderType, HR_DatePart_D, HR_DatePart_Y, HR_DatePart_M, HR_DatePart_W, HR_DatePart_H
Dim rs, sql, rsTmp, sqlTmp, strTmp, i, k, m

Dim sPath, SiteName, SiteTitle, SiteUrl, MetaKeywords, MetaDescription, objName_FSO, FSO, ObjInstalled_FSO, Copyright, CurrentVer, NewVer
Dim MailObject, MailServer, MailServerUserName, MailServerPassWord, MailDomain

Dim UserID, UserName, UserPass, GroupID, UserRank, arrGroup, LoginTimes
Dim FileName, strFileName, MaxPerPage, CurrentPage, totalPut, UpdatePages
Dim xmlDoc, XmlDOM, Node, VerifyCodeFile, ErrMsg, ErrHref
Dim TempFile, strHtml, tmpHtml, UserYGDM, UserYGXM, DefYear

BeginTime = Timer()
ConfigID = 1									'系统配置默认值
Call OpenConn()

'----- 正则表达式相关的变量
Dim regEx, Match, Match2, Matches, Matches2
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True
regEx.MultiLine = True

ScriptName = Trim(Request.ServerVariables("SCRIPT_NAME"))
UserTrueIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If UserTrueIP = "" Then UserTrueIP = Request.ServerVariables("REMOTE_ADDR")
UserTrueIP = ReplaceBadChar(UserTrueIP)
UpdatePages = 3												'列表更新页数

Sub OpenConn()
	On Error Resume Next
	Dim ConnStr
	If DataType = "SQL" Then
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlHostIP & ";"
	Else
		ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DBFileName)
	End If
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.open ConnStr
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "<div style=""margin:20px;text-align:center;font-size:14px;color:red;"">数据库连接出错，请检查数据库参数设置。</div>"
		Response.End
	End If
	If DataType = "SQL" Then
		HR_True = "1"
		HR_False = "0"
		HR_Now = "GetDate()"
		HR_OrderType = " desc"
		HR_DatePart_D = "d"
		HR_DatePart_Y = "yyyy"
		HR_DatePart_M = "m"
		HR_DatePart_W = "ww"
		HR_DatePart_H = "hh"
	Else
		HR_True = "True"
		HR_False = "False"
		HR_Now = "Now()"
		HR_OrderType = " asc"
		HR_DatePart_D = "'d'"
		HR_DatePart_Y = "'yyyy'"
		HR_DatePart_M = "'m'"
		HR_DatePart_W = "'ww'"
		HR_DatePart_H = "'h'"
	End If
End Sub

Sub CloseConn()
    On Error Resume Next
    If IsObject(Conn) Then
        Conn.Close
        Set Conn = Nothing
    End If
	Set regEx = Nothing
End Sub
%>
