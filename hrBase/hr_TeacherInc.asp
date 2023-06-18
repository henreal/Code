<%
Sub ImportAll()
	Server.ScriptTimeout=900
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.ImportTips {text-align:center;line-height:50px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.ImportTips b {color:#f30}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	'tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	'tmpHtml = tmpHtml & "		layui.use([""layer""], function(){ layer.load(1); });" & vbCrlf
	'tmpHtml = tmpHtml & "	</script>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	strHtml = strHtml & "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	strHtml = strHtml & "<fieldset class=""layui-elem-field site-demo-button"">" & vbCrlf
	strHtml = strHtml & "	<legend>导入员工数据</legend>" & vbCrlf
	strHtml = strHtml & "	<div class=""hr-shrink-x10"">" & vbCrlf
	strHtml = strHtml & "		<div class=""layui-progress layui-progress-big"" lay-showpercent=""true"" lay-filter=""demo"">" & vbCrlf
	strHtml = strHtml & "			<div class=""layui-progress-bar layui-bg-red"" lay-percent=""0%""></div>" & vbCrlf
	strHtml = strHtml & "		</div>" & vbCrlf
	strHtml = strHtml & "	</div>" & vbCrlf
	Dim apiStr : apiStr = GetHttpPage(apiHost & "/API/API.htm?Type=GetYgDict", 1)
	strHtml = strHtml & "	<div class=""ImportTips"" id=""ImportTips""><i class='layui-icon layui-anim layui-anim-rotate layui-anim-loop'>&#xe63d;</i>系统正在检查教师数据，可能会持续几分钟　<b>数据准备中，请稍候…</b></div>" & vbCrlf
	strHtml = strHtml & "</fieldset>" & vbCrlf
	strHtml = strHtml & "</div>" & vbCrlf
	Response.Write strHtml

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""layer"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var layer = layui.layer , element = layui.element;layer.load(1);" & vbCrlf
	strHtml = strHtml & "		UpdateImportData(0);" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf

	strHtml = strHtml & "	function UpdateImportData(numBegin){" & vbCrlf
	strHtml = strHtml & "		$.getJSON(""" & ParmPath & "Teacher/PostImportAll.html?Count=""+ numBegin, function(strForm){" & vbCrlf
	strHtml = strHtml & "			if(strForm.Return){" & vbCrlf
	strHtml = strHtml & "				$(""#ImportTips b"").html(strForm.End + "" / "" + strForm.total);" & vbCrlf
	strHtml = strHtml & "				var p1 = (strForm.End/strForm.total)*100;" & vbCrlf
	strHtml = strHtml & "				element.progress('demo', p1.toFixed(2) + '%');" & vbCrlf
	strHtml = strHtml & "				if(strForm.End>=strForm.total){" & vbCrlf
	strHtml = strHtml & "					layer.closeAll(""loading""); $(""#ImportTips b"").html(""完成！"");return false;" & vbCrlf
	strHtml = strHtml & "				}else{" & vbCrlf
	strHtml = strHtml & "					UpdateImportData(strForm.End);layer.load(1);" & vbCrlf
	strHtml = strHtml & "				}" & vbCrlf
	strHtml = strHtml & "			}else{" & vbCrlf
	strHtml = strHtml & "				$(""#ImportTips b"").html(strForm.reMessge);" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	}" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1) & strHtml
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub

Sub PostImportAll()
	Server.ScriptTimeout = 900
	Dim Count1 : Count1 = HR_Clng(Request("Count"))
	If Count1 = 0 Then session("Count") = 0

	Dim getStr, jsonObj, iYGDM, iYGLB, iYGZSM, iYGXM, iKSDM, iKSMC, iXMJP, iYGZT, iXZZW, iPRZC, iZFPB, str1
	Dim rsSave, sqlSave, beginLen, endLen, totalLen
	getStr = GetHttpPage(apiHost & "/Upload/Teacher.txt", 1)
	Set jsonObj = parseJSON(getStr)
		totalLen = jsonObj.reData.length-1
		If totalLen>0 Then
			str1 = "" : i = HR_Clng(session("Count"))
			beginLen = Count1 : endLen = beginLen + 200
			If endLen > totalLen Then endLen = totalLen
			For m=beginLen To endLen
				iYGDM = Trim(jsonObj.reData.get(m).YGDM)	'员工代码
				iYGXM = Trim(jsonObj.reData.get(m).YGXM)	'员工姓名
				iKSDM = Trim(jsonObj.reData.get(m).KSDM)	'科室代码
				iXZZW = Trim(jsonObj.reData.get(m).YGZW)	'职务
				iPRZC = Trim(jsonObj.reData.get(m).YGZC)	'职称
				iXMJP = Trim(jsonObj.reData.get(m).PYDM)	'拼音代码
				iYGZT = Trim(jsonObj.reData.get(m).YGLB)	'状态
				iZFPB = HR_Clng(jsonObj.reData.get(m).ZFPB)	'作废
				iKSMC = GetTypeName("HR_Department", "KSMC", "KSDM", iKSDM)
				If HR_Clng(iYGDM) > 0 And HR_IsNull(iYGXM) = False Then
					sqlSave = "Select * From HR_Teacher Where YGXM='" & iYGXM & "' And YGDM='" & iYGDM & "'"
					Set rsSave = Server.CreateObject("ADODB.RecordSet")
						rsSave.Open(sqlSave), Conn, 1, 3
						If rsSave.BOF And rsSave.EOF Then
							rsSave.AddNew
							rsSave("TeacherID") = GetNewID("HR_Teacher", "TeacherID")
							rsSave("ApiType") = 2
							rsSave("LoginPass") = "83aa400af464c76d"
							rsSave("UpdateTime") = Now()
							rsSave("YGZT") = iYGZT
							rsSave("XMJP") = iXMJP

							rsSave("YGDM") = iYGDM
							rsSave("YGXM") = iYGXM
							rsSave("KSDM") = iKSDM
							rsSave("KSMC") = iKSMC
							rsSave("XZZW") = iXZZW
							rsSave("PRZC") = iPRZC

							rsSave("PXXH") = GetNewID("HR_Teacher", "PXXH")
							rsSave("ZFPB") = iZFPB
							
							i = i + 1
							session("Count") = i
						End If
						rsSave("ZFPB") = iZFPB
						rsSave.Update

					Set rsSave = Nothing
				End If
				'str1 = str1 & "<li>" & iYGXM & "[" & iYGDM & "]：" & iKSMC & "[" & iKSDM & "]</li>"
			Next
			'str1 = "<br><ul>" & str1 & "</ul>"
			ErrMsg = "{""Return"":true,""Err"":0,""reMessge"":""共更新“" & i & "/" & HR_Clng(session("Count")) & "”名员工！" & str1 & """,""Begin"":" & beginLen & ",""End"":" & endLen & ",""total"":" & totalLen & "}"
		Else
			ErrMsg = "{""Return"":false,""Err"":400,""reMessge"":""导入的数据为空！" & str1 & """,""ReStr"":[]}"
		End If
	Set jsonObj = Nothing
	Response.Write ErrMsg
End Sub

Sub ResetPass()		'重置密码
	Dim tmpID :tmpID = HR_Clng(Request("ID"))
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_Teacher Where TeacherID>0 And TeacherID=" & tmpID), Conn, 1, 3
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			rsTmp("LoginPass") = "83aa400af464c76d"
			rsTmp.Update
			ErrMsg = "{""Return"":true,""Err"":0,""reMessge"":""" & rsTmp("YGXM") & "老师[工号：" & rsTmp("YGDM") & "]的登陆密码已经被重置为“12345678”！<br />请登陆后及时修改！"",""ReStr"":[]}"
		Else
			ErrMsg = "{""Return"":false,""Err"":500,""reMessge"":""该教师不存在或已被删除[ID:" & tmpID & "]"",""ReStr"":[]}"
		End If
	Set rsTmp = Nothing
	Response.Write ErrMsg
End Sub

Sub UpdatePinyin()		'更新姓名拼音
	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-navmenu-main u {display:inline;padding:0 5px;font-style:normal;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Dim tName
	Set rs = Conn.Execute("Select * From HR_Teacher Where (ascii(XMJP) between 48 and 57) Or ltrim(rtrim(XMJP))='' Order By TeacherID DESC")
		Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
		Response.Write "	<div class=""hr-fix"">" & vbCrlf
		If Not(rs.BOF And rs.EOF) Then
			Do While Not rs.EOF
				tName = rs("YGXM")
				tName = ConvertCnToPy(tName, 1)
				If HR_Clng(rs("XMJP")) > 0 Or HR_IsNull(rs("XMJP")) Then
					Response.Write "" & rs("YGXM") & " | " & rs("XMJP") & " | " & tName & "		<br>" & vbCrlf
					Conn.Execute("Update HR_Teacher Set XMJP='" & tName & "' Where TeacherID=" & rs("TeacherID"))
				End If
				rs.MoveNext
			Loop
		End If
		Response.Write "	</div>" & vbCrlf
		Response.Write "</div>" & vbCrlf & Timer()-BeginTime
	Set rs = Nothing

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element;" & vbCrlf

	tmpHtml = tmpHtml & "		layer.closeAll(""loading"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

'============================================
'函数名：ConvertCnToPy()【汉字转拼音，共20850个汉字】
'ftype：为1是取首字母
'提示：预加载拼音数据, 并定义数组变量名为：arrPinYin
'数据文件：Static/js/PinyinData.js，注意替换js字符
'--------------------------------------------
Function ConvertCnToPy(fStr, ftype)
	Dim strFun, fstr1
	For i = 1 To Len(fStr)
		fstr1 = Mid(fStr, i, 1)
		For m = 0 To Ubound(arrPinYin)
			If fstr1 = Left(arrPinYin(m), 1) Then
				If HR_Clng(ftype) = 1 Then
					strFun = strFun & "" & UCase(Left(Replace(arrPinYin(m), fstr1, ""), 1))
				Else
					strFun = strFun & "" & Replace(arrPinYin(m), fstr1, "")
				End If
			End If
		Next
	Next
	ConvertCnToPy = strFun
End Function
%>