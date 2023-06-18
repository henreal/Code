<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incMD5.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Server.ScriptTimeout=240			'超时4分钟（预留1分钟缓存，避免服务器进入死循环）

If ChkTokenBobao() = False Then Call GetTokenBobao()
If ChkWechatTokenQY() = False Then Call GetWechatTokenQY()

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl, jsonOBJ

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index", "A1" Call MainBody()
	Case Else Response.Write GetErrBody()
End Select

Sub MainBody()
	Dim tmpJson : ErrMsg = ""
	'检查Token是否过期
	If ChkTokenBobao() = False Then
		Call GetTokenBobao()
		Call GetWechatTokenQY()
		If HR_CLng(Request("reload")) = 0 Then
			Response.Redirect InstallDir & "API/TimedTask.html?reload=1"		'刷新一次本页
			Exit Sub
		Else
			Response.Write "{""err"":true,""errcode"":500,""errmsg"":""获取企业微信Token失败！"",""icon"":2,""tips"":""已重试1次""}"
			Exit Sub
		End If
	End If
	Call SendCourseRemind()		'上课提醒
	'Response.Write "暂停发送消息！"
	'Response.Write "{""err"":false,""errcode"":0,""errmsg"":""暂停发送消息！"",""icon"":1,""redata"":[]}"
End Sub

'发送上课提醒
Sub SendCourseRemind()
	Dim postJson, isSend, tExpired, tDiffTime, tDataTable, reSendMsg, tStartTime
	Dim fLastDay : fLastDay = Day(DateAdd("m", 1, FormatDate(Date(),9) + "-1") -1)		'本月最后一天
	Dim tRemindTime : tRemindTime = XmlText("Common", "RemindTime", "")					'提醒时间点（轮巡周期为5分钟）
	Dim tBeforeTime : tBeforeTime = HR_Clng(XmlText("Common", "BeforeTime", ""))		'提前时间（秒）
	Dim SendUser : SendUser = "810000"
	Dim isBatch : isBatch = HR_CBool(Request("Batch"))									'批量提醒一周内

	sqlTmp = GetRemainCourse(" And Passed=" & HR_True)
	sqlTmp = "Select * Into #tmpTable From (" & sqlTmp & ") a"
	Conn.Execute(sqlTmp)
	sqlTmp = "Select * From #tmpTable Order By VA4 ASC,VA7 ASC"
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open sqlTmp, Conn, 1, 1
		'Response.Write "{""err"":false,""errcode"":0,""errmsg"":""共有" & rsTmp.RecordCount & "条数据"",""icon"":1,""redata"":[{""SQL"":""" & sqlTmp & """}]}"
		'Exit Sub
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0 : m = 0
			Do While Not rsTmp.EOF
				'判断是否符合发送条件
				isSend = False
				tDataTable = "HR_Sheet_" & rsTmp("ItemID")
				tStartTime = GetPeriodTime(rsTmp("VA11"), rsTmp("VA7"), 1)
				If HR_IsNull(tStartTime) = False Then tRemindTime = tStartTime
				tExpired = FormatDate(rsTmp("ClassDate"), 2) & " " & tRemindTime & ":00"		'到期时间

				If IsDate(tExpired) And ChkTable(tDataTable) Then				'判断到期时间格式及数据表
					tDiffTime = DateDiff("s", Now(), tExpired)					'取到期时间与当前时间差（秒）
					If tDiffTime >= 5390 And tDiffTime <= 5705 Then			'提前一个半小时提醒5400
						isSend = True
					ElseIf tDiffTime > 1200 And tDiffTime < 605105 And isBatch Then		'提醒本周内所有老师（测试用）
						isSend = True
					ElseIf tDiffTime >= 86390 And tDiffTime <= 86705 Then		'提前24小时提醒86400
						isSend = True
					ElseIf tDiffTime >= 604790 And tDiffTime <= 605105 Then	'提前1周提醒604800
						isSend = True
					End If
					'Response.Write "<br>" & isSend & "【" & rsTmp("ClassDate") & "】" & rsTmp("VA2") & "时间：距" & tExpired & " 有" & tDiffTime
				End If

				If isSend Then
					SendUser = "810000"
					postJson = "{""touser"" : """ & SendUser & """,""msgtype"":""textcard"",""agentid"" : " & boAgentId & ",""textcard"":{""title"":""【上课提醒】" & rsTmp("VA2") & " 老师：您于" & FormatDate(tExpired, 4) & " 第" & rsTmp("VA5") & "周 " & rsTmp("VA7") & "节有未授课程"","
					postJson = postJson & """description"":""课程名称：" & rsTmp("ItemName") & " " & rsTmp("VA8") & "<br>授课时间：第" & rsTmp("VA7") & "节 " & FormatDate(tExpired, 10)
					postJson = postJson & "<br>授课内容：" & rsTmp("VA9") & "<br>授课对象：" & rsTmp("VA10")
					postJson = postJson & "<br>授课地点：" & rsTmp("VA11") & " " & rsTmp("VA12")
					postJson = postJson & "<br>本次消息发送时间：" & Formatdate(Now(), 1) & ""","
					postJson = postJson & """url"":""" & SiteUrl & "/Touch/Remain/Index.html"",""btntxt"":""查看详情""}}"
					reSendMsg = PostWechatMessageQY(postJson, 1)		'发送

					SendUser = HR_Clng(rsTmp("VA1"))
					postJson = "{""touser"" : """ & SendUser & """,""msgtype"":""textcard"",""agentid"" : " & boAgentId & ",""textcard"":{""title"":""【上课提醒】" & rsTmp("VA2") & " 老师：您于" & FormatDate(tExpired, 4) & " 第" & rsTmp("VA5") & "周 " & rsTmp("VA7") & "节有未授课程"","
					postJson = postJson & """description"":""课程名称：" & rsTmp("ItemName") & " " & rsTmp("VA8") & "<br>授课时间：第" & rsTmp("VA7") & "节 " & FormatDate(tExpired, 10)
					postJson = postJson & "<br>授课内容：" & rsTmp("VA9") & "<br>授课对象：" & rsTmp("VA10")
					postJson = postJson & "<br>授课地点：" & rsTmp("VA11") & " " & rsTmp("VA12")
					postJson = postJson & "<br>本次消息发送时间：" & Formatdate(Now(), 1) & ""","
					postJson = postJson & """url"":""" & SiteUrl & "/Touch/Remain/Index.html"",""btntxt"":""查看详情""}}"
					
					reSendMsg = PostWechatMessageQY(postJson, 1)		'发送
					If Instr(reSendMsg, "errmsg") > 0 Then				'发送结果
						Set jsonOBJ = parseJSON(reSendMsg)
							If jsonOBJ.errmsg = "ok" Then			'发送成功
								'Response.Write "<br>" & reSendMsg
								m = m + 1
							End If
							Call WriteTextCronLog(jsonOBJ.errmsg & "【上课提醒/" & SendUser & "】" & rsTmp("VA2") & "[" & rsTmp("VA1") & "]老师：" & FormatDate(tExpired, 10) & " 第" & rsTmp("VA7") & "节|ItemID:" & rsTmp("ItemID") & "/" & rsTmp("ID") )		'存入日志
						Set jsonOBJ = Nothing
					Else
						'Response.Write "<br>" & reSendMsg	'发送结果
					End If						
		
				End If
				rsTmp.MoveNext
				i = i + 1
			Loop
			If HR_CLng(i)>0 Then
				Response.Write "已发送" & HR_CLng(m) & "/" & HR_CLng(i) & "条上课提醒<br>"' & tExpired
			End If
		End If
	Set rsTmp = Nothing
	Conn.Execute("Drop Table #tmpTable")			'删除临时表
	'Response.Write "<br>" & sqlTmp
End Sub

'发送未确认提醒
Sub SendFallbackRemind()
	
End Sub

Function GetItemClassID(fLimit)		'取考核项目ID
	Dim rsFun, sqlFun, iFun, strFun : strFun = ""
	sqlFun = "Select ClassID From HR_Class Where ModuleID=1001 And Child=0"
	If HR_IsNull(fLimit) = False Then sqlFun = sqlFun & " " & fLimit
	Set rsFun = Server.CreateObject("ADODB.RecordSet")
		rsFun.Open(sqlFun), Conn, 1, 1
		If Not(rsFun.BOF And rsFun.EOF) Then
			iFun = 0
			Do While Not rsFun.EOF
				If iFun > 0 Then strFun = strFun & ","
				strFun = strFun & rsFun(0)
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		End If
	Set rsFun = Nothing
	GetItemClassID = strFun
End Function

Function GetRemainCourse(fLimit)				'取未上课程(8天内所有课程)
	Dim iFun, funItem, arrItem, strFun : strFun = ""
	funItem = GetItemClassID(" And Template in('TempTableA') ")		'取考核项ID
	If HR_IsNull(funItem) = False Then
		arrItem = Split(FilterArrNull(funItem, ","), ",")
		iFun = 0
		For m = 0 To Ubound(arrItem)
			If iFun > 0 Then strFun = strFun & " union all "
			strFun = strFun & "(Select a.ID,a.ItemID,b.ClassName As ItemName,a.VA1,a.VA2,a.VA4,a.VA5,a.VA7,a.VA8,a.VA9,a.VA10,a.VA11,a.VA12,DATEADD(d,a.VA4-2,'1900-01-01') As ClassDate From HR_Sheet_" & arrItem(m) & " a Inner Join HR_CLass b on a.ItemID=b.ClassID Where a.VA4>0 And a.VA4>DATEDIFF(d,'1900-01-01',getDate())+1 And a.VA4<=DATEDIFF(d,'1900-01-01',getDate())+9"
			If HR_IsNull(fLimit) = False Then strFun = strFun & " " & fLimit
			strFun = strFun & ")"
			iFun = iFun + 1
		Next
	End If
	GetRemainCourse = strFun
End Function

'=====================================================================
'函数名：ChkTable		【检查表是否存在，返回布尔值】
'=====================================================================
Function ChkTable(tTableName)
	on error resume next
	Dim rsFun : ChkTable = False
	If isNull(tTableName) Or tTableName = "" Then Exit Function
	Set rsFun = Server.CreateObject("ADODB.RecordSet")
		rsFun.open "Select Top 1 * From " & tTableName, Conn, 1, 1
		If Not Err.Number=0 Then Err.Clear : Exit Function
		ChkTable = True
		rsFun.Close
	Set rsFun = Nothing 
End Function

'=====================================================================
'函数名：ConvertNumDate		【时间戳返回日期】
'=====================================================================
Function ConvertNumDate(timeStamp)
	If IsEmpty(timeStamp) or Not IsNumeric(timeStamp) Then
        ConvertNumDate = FormatDate(Now(), 1)
        Exit Function
    End If
	ConvertNumDate = DateAdd("d", timeStamp-2, "1900-01-01 00:00:00")		'减2调整时间差【因为PHPExcel未设置格式】
End Function
Function ConvertDateToNum(fTime)		'将日期转为时间戳
	If IsEmpty(fTime) Or Not IsDate(fTime) Or fTime="" Then fTime = FormatDate(Now(), 2) & " 23:59:59"
	ConvertDateToNum = DateDiff("d","1900-01-01 00:00:00", fTime)
End Function
'=====================================================================
'函数名：GetPeriodTime		【根据节次返回时间】
'fType：1只取开始时间/0上课时间段
'=====================================================================
Function GetPeriodTime(fCampus, fPeriod, fType)
	Dim strFun, rsFun, fArr, strArr, stTime, enTime
	fCampus = Trim(fCampus) : fPeriod = Trim(fPeriod)
	If fCampus <> "" And fPeriod <> "" Then
		If HR_Clng(fPeriod) > 0 Then
			Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fPeriod))
				If Not(rsFun.BOF And rsFun.EOF) Then
					strFun = Trim(rsFun("StartTime")) & " - " & Trim(rsFun("EndTime"))
					If HR_CLng(fType) = 1 Then strFun = Trim(rsFun("StartTime"))
				End If
			Set rsFun = Nothing
		ElseIf Instr(fPeriod, "-") Then
			fArr = Split(fPeriod, "-")
			If Ubound(fArr) = 1 Then
				Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fArr(0)))
					If Not(rsFun.BOF And rsFun.EOF) Then
						stTime = Trim(rsFun("StartTime"))
					End If
				Set rsFun = Nothing
				Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fArr(1)))
					If Not(rsFun.BOF And rsFun.EOF) Then
						enTime = Trim(rsFun("EndTime"))
					End If
				Set rsFun = Nothing
				strFun = Trim(stTime) & " - " & Trim(enTime)
				If HR_CLng(fType) = 1 Then strFun = Trim(stTime)
			End If
		End If
	End If
	GetPeriodTime = strFun
End Function

Function WriteTextCronLog(fLog)
	Dim ftxt, fFileName : fFileName = Server.MapPath("/Upload/CronLog/Cron_" & FormatDate(Now(),11) & ".txt")
	Call CreateMultiFolder("/Upload/CronLog")
	If fso.FileExists(fFileName) = False Then
		Set ftxt = fso.CreateTextFile(fFileName, false)
		Set ftxt = Nothing
	End If
	Set ftxt = fso.OpenTextFile(fFileName, 8, true)
		ftxt.WriteLine(Now() & "：" & fLog)
	Set ftxt = Nothing
End Function
%>
