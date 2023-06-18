<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./incKPI.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim scriptCtrl, SubButTxt : SiteTitle = "数据导入"
Dim IsImport : IsImport = HR_CBool(XmlText("Common", "ImportSwitch", "0"))

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Teacher" Call ImportTeacher()
	Case "PostTeacher" Call PostTeacher()
	Case "UpdateTeacherData" Call UpdateTeacherData()
	Case Else Response.Write GetErrBody(1)
End Select

Sub ImportTeacher()
	Server.ScriptTimeout = 900
	Dim tStuType : ErrMsg = ""
	Dim apiDataType : apiDataType = Trim(ReplaceBadChar(Request("apiAction")))
	Dim tmpHtml, xlsUrl, getStr, jsonObj, st1
	'getStr = GetHttpPage(apiHost & "/API/API.htm?Type=GetRyxxForJXGL", 1)

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.im-box {min-height:180px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-tips_dl dt {padding-right:0;} .hr-tips_dl dd {padding-left:15px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.normal dt {color:#060; width:50px; height:50px; line-height:50px;} .normal dd h4{color:#060;}" & vbCrlf

	tmpHtml = tmpHtml & "	</style>"
	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	tmpHtml = vbCrlf & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>"
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	tmpHtml = "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	tmpHtml = tmpHtml & "	<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	tmpHtml = tmpHtml & "		<legend title=""导入教师数据"">导入教师库数据</legend>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""hr-shrink-x10 hr-align_c""><button class=""layui-btn layui-btn-sm"" type=""button"" name=""ImportPost"" id=""ImportPost"">导入</button><button type=""button"" class=""layui-btn layui-btn-sm"" name=""ViewBtn"" id=""ViewBtn"">查看</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""im-box"" id=""ImportData"">" & vbCrlf
	tmpHtml = tmpHtml & "			<dl class=""hr-tips_dl""><dt><i class=""hr-icon"">&#xefd5;</i></dt><dd><h4>操作提示：</h4><p>请点击上面导入按钮</p></dd></dl>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""xlsData"" id=""xlsData""></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</fieldset>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	Response.Write tmpHtml

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#ImportPost"").on(""click"", function(){" & vbCrlf		'从API读取数据存到本地
	tmpHtml = tmpHtml & "			$("".hr-tips_dl"").addClass(""normal""); $("".hr-tips_dl dt"").addClass(""layui-anim layui-anim-rotate layui-anim-loop"");" & vbCrlf
	tmpHtml = tmpHtml & "			$("".hr-tips_dl dt"").html(""<i class=\""hr-icon\"">&#xefe3;</i>""); $("".hr-tips_dl dd"").find(""h4"").text(""正在与远程服务器通讯中"");" & vbCrlf
	tmpHtml = tmpHtml & "			$("".hr-tips_dl dd"").find(""p"").text(""可能需要几分钟，请稍候…"");" & vbCrlf
	tmpHtml = tmpHtml & "			var posturl = """ & ParmPath & "ImportTmp/PostTeacher.html"";" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(posturl, {Action:""GetRyxxForJXGL""}, function(res){" & vbCrlf
	tmpHtml = tmpHtml & "				if(res.err){" & vbCrlf
	tmpHtml = tmpHtml & "					$(""#ImportData"").html(""<dl class=\""hr-tips_dl\""><dt><i class=\""hr-icon\"">&#xef61;</i></dt><dd><h4>导入中止！</h4><p>与远程数据通讯时发生错误</p></dd></dl>"");" & vbCrlf
	tmpHtml = tmpHtml & "				}else{" & vbCrlf
	tmpHtml = tmpHtml & "					$(""#ImportData"").html(""<dl class=\""hr-tips_dl normal\""><dt><i class=\""hr-icon\"">&#xefe3;</i></dt><dd><h4>"" + res.errmsg + ""</h4><p>"" + res.tips + ""</p></dd></dl>"");" & vbCrlf
	'tmpHtml = tmpHtml & "					$(""#xlsData"").html(res.errmsg);" & vbCrlf
	tmpHtml = tmpHtml & "					updateTeacher(0, 50);" & vbCrlf
	tmpHtml = tmpHtml & "				};" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		$(""#viewTpl"").on(""click"", function(){" & vbCrlf			'预览模板
	tmpHtml = tmpHtml & "			var Temp1 = $(""#TempName"").val(), itemName = $(""#ItemName"").val();" & vbCrlf
	tmpHtml = tmpHtml & "			parent.layer.open({type:2,content:""" & ParmPath & "Import/ExcelTemp.html?ItemID=1&itemName=1"",title:[""查看Excel模板"",""font-size:16""],area:[""95%"", ""72%""], moveOut:true, maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	function updateTeacher(fStartID, fLen){" & vbCrlf
	tmpHtml = tmpHtml & "		$.getJSON(""" & ParmPath & "ImportTmp/UpdateTeacherData.html"", {StartID:fStartID, Len:fLen}, function(res){" & vbCrlf
	'tmpHtml = tmpHtml & "			$("".hr-tips_dl dt"").html(""<i class=\""hr-icon\"">&#xefe3;</i>"");" & vbCrlf
	tmpHtml = tmpHtml & "			$("".hr-tips_dl dd"").find(""h4"").text(res.errmsg); $("".hr-tips_dl dd"").find(""p"").text(res.tips);" & vbCrlf
	'tmpHtml = tmpHtml & "			$(""#ImportData"").html(""<dl class=\""hr-tips_dl normal\""><dt class=\""layui-anim layui-anim-rotate layui-anim-loop\""><i class=\""hr-icon\"">&#xefe3;</i></dt><dd><h4>"" + res.errmsg + ""</h4><p>"" + res.tips + ""</p></dd></dl>"");" & vbCrlf
	tmpHtml = tmpHtml & "			if(!res.err){" & vbCrlf
	tmpHtml = tmpHtml & "				$("".hr-tips_dl dt"").addClass(""layui-anim layui-anim-rotate layui-anim-loop"");" & vbCrlf
	tmpHtml = tmpHtml & "				setTimeout(updateTeacher(res.Start, 50), 100);" & vbCrlf
	tmpHtml = tmpHtml & "			}else{" & vbCrlf
	tmpHtml = tmpHtml & "				$("".hr-tips_dl"").removeClass(""normal"");" & vbCrlf
	tmpHtml = tmpHtml & "				$("".hr-tips_dl dt"").html(""<i class=\""hr-icon\"">&#xef8a;</i>"");" & vbCrlf
	tmpHtml = tmpHtml & "				$("".hr-tips_dl dt"").removeClass(""layui-anim layui-anim-rotate layui-anim-loop"");" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	}" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(0)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub PostTeacher()
	Server.ScriptTimeout = 900		'超时900秒(15分钟)
	Dim tAction : tAction = Trim(ReplaceBadChar(Request("Action")))
	Dim jsonObj, getStr
	getStr = GetHttpPage(apiHost & "/API/API.htm?Type=GetRyxxForJXGL", 1)		'参数Type:GetAllRyxx所有员工

	If Instr(getStr, "reMessge") > 0 Then
		Set jsonObj = parseJSON(getStr)
			Response.Write "{""err"":false, ""errcode"":0,""icon"":1, ""errmsg"":""" & jsonObj.reMessge & """, ""tips"":""" & jsonObj.ReStr & """}"
		Set jsonObj = Nothing
	Else
		Response.Write "{""err"":true, ""errcode"":500,""icon"":2, ""errmsg"":""导入中止！"", ""tips"":""与远程数据通讯时发生错误""}"
	End If
End Sub
Sub UpdateTeacherData()
	Dim tStartID : tStartID = HR_CLng(Request("StartID"))
	Dim tLen : tLen = HR_CLng(Request("Len"))
	Dim tmpNow : tmpNow = Now()
	Dim tCountNum : tCountNum = 0
	Dim tEndID : tEndID = tStartID + tLen

	Dim apiDataType : apiDataType = Trim(ReplaceBadChar(Request("apiAction")))
	apiDataType = "GetRyxxForJXGL"
	Dim dataFile : dataFile = "Teacher08.txt"
	If apiDataType = "GetAllRyxx" Then dataFile = "AllTeacher.txt"
	Dim jsonObj, getStr : getStr = GetHttpPage(apiHost & "/Upload/" & dataFile, 1)

	If Instr(getStr, "reData") = 0 Then			'没有数据
		Response.Write "{""err"":true, ""errcode"":500,""icon"":2, ""errmsg"":""导入错误！"", ""tips"":""没有员工数据""}" : Exit Sub
	End If

	Dim rsUpdate, j, k, tYGDM, tYGXM, tYGXB, tUpKPI
	Set jsonObj = parseJSON(getStr)
		tCountNum = HR_CLng(jsonObj.reData.length)-1		'总数
		If tEndID > tCountNum Then tEndID = tCountNum
		j = 0 : k = 0
		For m = tStartID To tEndID
			tYGDM = Trim(jsonObj.reData.get(m).YGDM)
			tYGXM = Trim(jsonObj.reData.get(m).YGXM)
			If HR_IsNull(tYGDM) = False And HR_IsNull(tYGXM) = False Then
				Set rsUpdate = Server.CreateObject("ADODB.RecordSet")
					rsUpdate.Open("Select Top 1 * From HR_Teacher Where YGXM='" & tYGXM & "' And YGDM='" & tYGDM & "'"), Conn, 1, 3
					If rsUpdate.BOF And rsUpdate.EOF Then
						rsUpdate.AddNew
						rsUpdate("TeacherID") = GetNewID("HR_Teacher", "TeacherID")
						rsUpdate("YGDM") = tYGDM
						rsUpdate("YGXM") = tYGXM
						rsUpdate("PXXH") = GetNewID("HR_Teacher", "PXXH")
						rsUpdate("LoginPass") = "83aa400af464c76d"
						rsUpdate("ImportTime") = tmpNow
						j = j + 1
					Else
						k = k + 1
					End If
					rsUpdate("Explain") = Trim(Request("Explain"))				'注释
					rsUpdate("UpdateTime") = tmpNow
					rsUpdate("KSDM") = HR_Clng(jsonObj.reData.get(m).KSDM)		'科室代码

					If apiDataType = "GetRyxxForJXGL" Then
						rsUpdate("ApiType") = 3
						rsUpdate("KSMC") = Trim(jsonObj.reData.get(m).KSMC)			'科室名称
						rsUpdate("XMJP") = Trim(jsonObj.reData.get(m).XMJP)			'姓名简拼
						rsUpdate("YGZT") = Trim(jsonObj.reData.get(m).YGZT)			'员工状态
						rsUpdate("YGXB") = Trim(jsonObj.reData.get(m).YGXB)			'员工性别
						rsUpdate("ZCBM") = HR_Clng(jsonObj.reData.get(m).ZCBM)		'职称编码
						rsUpdate("PRZC") = Trim(jsonObj.reData.get(m).PRZC)
						rsUpdate("ZWBM") = HR_Clng(jsonObj.reData.get(m).ZWBM)		'职务编码
						rsUpdate("XZZW") = Trim(jsonObj.reData.get(m).XZZW)			'职务
						rsUpdate("YGNM") = tYGDM
					ElseIf apiDataType = "GetAllRyxx" Then
						rsUpdate("ApiType") = 1
						rsUpdate("YGNM") = Trim(jsonObj.reData.get(m).YGNM)
						tYGXB = HR_Clng(jsonObj.reData.get(m).YGXB)
						rsUpdate("YGXB") = arrSex(tYGXB)		'员工性别
						rsUpdate("KSMC") = Trim(jsonObj.reData.get(m).KSMC)			'科室名称
						rsUpdate("RYRQ") = FormatAPIDate(jsonObj.reData.get(m).RYRQ, 0)		'入院原日期
						rsUpdate("YGZT") = Trim(jsonObj.reData.get(m).YGZT)
						rsUpdate("PDZC") = Trim(jsonObj.reData.get(m).PDZC)
						rsUpdate("PDRQ") = FormatAPIDate(jsonObj.reData.get(m).PDRQ, 0)
						rsUpdate("PRZC") = Trim(jsonObj.reData.get(m).PRZC)
						rsUpdate("PRRQ") = FormatAPIDate(jsonObj.reData.get(m).PRRQ, 0)
						rsUpdate("YGXW") = Trim(jsonObj.reData.get(m).YGXW)
						rsUpdate("YGXL") = Trim(jsonObj.reData.get(m).YGXL)
						rsUpdate("YGXZ") = Trim(jsonObj.reData.get(m).YGXZ)
						rsUpdate("BYXX") = Trim(jsonObj.reData.get(m).BYXX)
						rsUpdate("BYZY") = Trim(jsonObj.reData.get(m).BYZY)
						rsUpdate("RXRQ") = FormatAPIDate(jsonObj.reData.get(m).RXRQ, 0)
						rsUpdate("BYRQ") = FormatAPIDate(jsonObj.reData.get(m).BYRQ, 0)		'毕业日期
						rsUpdate("CSRQ") = FormatAPIDate(jsonObj.reData.get(m).CSRQ, 0)
						rsUpdate("JG") = Trim(jsonObj.reData.get(m).JG)
						rsUpdate("ZJH") = Trim(jsonObj.reData.get(m).ZJH)
						rsUpdate("GZRQ") = FormatAPIDate(jsonObj.reData.get(m).GZRQ, 0)
						rsUpdate("XMJP") = Trim(jsonObj.reData.get(m).XMJP)
						rsUpdate("ZZMM") = Trim(jsonObj.reData.get(m).ZZMM)
						rsUpdate("SJHM") = Trim(jsonObj.reData.get(m).SJHM)
						rsUpdate("DH") = Trim(jsonObj.reData.get(m).DH)
						rsUpdate("HLHSKSSJ") = Trim(jsonObj.reData.get(m).HLHSKSSJ)
						rsUpdate("HL") = Trim(jsonObj.reData.get(m).HL)
						rsUpdate("PYJG") = Trim(jsonObj.reData.get(m).PYJG)
						rsUpdate("XZZW") = Trim(jsonObj.reData.get(m).XZZW)
						rsUpdate("RMRQ") = FormatAPIDate(jsonObj.reData.get(m).RMRQ, 0)
						rsUpdate("RZJSRQ") = Trim(jsonObj.reData.get(m).RZJSRQ)
					Else
						rsUpdate("ApiType") = 2
						rsUpdate("YGZT") = Trim(jsonObj.reData.get(m).YGLB)
						rsUpdate("XMJP") = Trim(jsonObj.reData.get(m).PYDM)
						rsUpdate("ZFPB") = HR_Clng(jsonObj.reData.get(m).ZFPB)
						rsUpdate("SIGN") = Trim(jsonObj.reData.get(m).SIGN)
						rsUpdate("PRZC") = Trim(jsonObj.reData.get(m).YGZC)
						rsUpdate("XZZW") = Trim(jsonObj.reData.get(m).YGZW)
					End If
					rsUpdate.Update
					tUpKPI = ChkTeacherKPI(tYGDM)	'添加员工信息至业绩表
				Set rsUpdate = Nothing
			End If
		Next
	Set jsonObj = Nothing


	If tStartID<tCountNum Then
		Response.Write "{""err"":false, ""errcode"":0,""icon"":1, ""errmsg"":""数据更新中"", ""tips"":""正在核对 " & tEndID & "/" & tCountNum & """, ""Start"":" & tEndID & "}"
	Else
		Response.Write "{""err"":true, ""errcode"":500,""icon"":2, ""errmsg"":""导入完成！"", ""tips"":""结束""}"
	End If
End Sub

%>