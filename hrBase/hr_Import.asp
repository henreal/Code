<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./incKPI.asp"-->
<%
Dim SubButTxt : SiteTitle = "数据导入"
Dim IsImport : IsImport = HR_CBool(XmlText("Common", "ImportSwitch", "0"))

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Excel" Call ImportExcel()
	Case "SaveExcel" Call SaveExcel()
	Case "ShowResult" Call ShowResult()

	Case "ExcelTemp" Call ExcelTemp()		'预览Excel模板文件
	Case Else Response.Write GetErrBody(1)
End Select

Sub ImportExcel()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tSheetName, tItemName, tTemplate, tStuType, stuType, tArr2
	ErrMsg = ""
	If Not(IsImport) Then		'判断是否允许导入
		ErrMsg = "非常遗憾，导入功能已关闭！"
		Response.Write GetErrBody(1) : Exit Sub
	End If
	tSheetName = "HR_Sheet_" & tItemID

	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tStuType = Trim(rsTmp("StudentType"))
			tTemplate = Trim(rsTmp("Template"))
		Else
			ErrMsg = "非常遗憾，该考核项目不存在！"
			Response.Write GetErrBody(1) : Exit Sub
		End If
	Set rsTmp = Nothing

	If ChkTable(tSheetName) = False Then
		ErrMsg = "考核项目 " & tItemName & " 对应的数据表未建立！"
		Response.Write GetErrBody(1) : Exit Sub
	End If


	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.tipsStep {font-size:18px;} .tipsStep i {font-size:32px;color:#92b}" & vbCrlf
	tmpHtml = tmpHtml & "		.tipsWarn {color:#999;padding-left:2em;} .tipsWarn h4 {color:#f00;} .tipsWarn h4 i {color:#f30;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>"

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	tmpHtml = vbCrlf & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>"
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf

	tmpHtml = "	<fieldset class=""layui-elem-field"">" & vbCrlf
	tmpHtml = tmpHtml & "		<legend title=""" & tSheetName & """>导入课程数据【" & tItemName & "】</legend>" & vbCrlf
	tmpHtml = tmpHtml & "		<form class=""layui-form layui-form-pane"" id=""ImportForm"" name=""ImportForm"" action=""" & ParmPath & "Import/SaveExcel.html"">" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layer-hr-box"" id=""xlsBox"">" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""hr-grids tipsStep""><em><i class=""hr-icon"">&#xe90a;</i></em><em>上传Excel数据文件：</em></div>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-form-item tipsWarn""><h4><i class=""hr-icon"">&#xeb85;</i>重要提示：请上传Excel文件时务必检查是否与模板一致，否则会出错！</h4>"
	tmpHtml = tmpHtml & "<p>如果多次出现格式错误的提示，请新建Excel文件，然后将复制数据到新文件中，再导入。</p>"
	tmpHtml = tmpHtml & "<h5>操作步骤：上传 → 保存 → 提交 → 模版</h5>"
	tmpHtml = tmpHtml & "</p></div>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-form-item""><label class=""layui-form-label"">Excel文件：</label><div class=""layui-input-block""><input type=""text"" name=""xlsUrl"" id=""xlsUrl"" value="""" placeholder=""请上传Excel文件"" lay-verify=""required"" class=""layui-input""></div></div>" & vbCrlf
	If tStuType <> "" Then
		tmpHtml = tmpHtml & "				<div class=""layui-form-item"" pane>" & vbCrlf
		tmpHtml = tmpHtml & "					<label class=""layui-form-label"">学生类别：</label>" & vbCrlf
		tmpHtml = tmpHtml & "					<div class=""layui-input-block"" id=""stuType"">" & vbCrlf
		tStuType = FilterArrNull(tStuType, ",")
		tArr2 = Split(tStuType, ",")
		For i = 0 To Ubound(tArr2)
			tmpHtml = tmpHtml & "						<input type=""radio"" name=""StudentType"" value=""" & tArr2(i) & """ title=""" & tArr2(i) & """ lay-skin=""primary"""
			If FoundInArr(UserStuType, tArr2(i), ",") = False And UserRank=1 Then tmpHtml = tmpHtml & " disabled="""""
			tmpHtml = tmpHtml & ">" & vbCrlf
		Next
		tmpHtml = tmpHtml & "					</div>" & vbCrlf
		tmpHtml = tmpHtml & "				</div>" & vbCrlf
	End If
	tmpHtml = tmpHtml & "				<input type=""hidden"" name=""ItemID"" id=""ItemID"" value=""" & tItemID & """>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-form-item soBox"">" & vbCrlf
	tmpHtml = tmpHtml & "					<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-btn-danger"" id=""upExcel""><i class=""hr-icon hr-icon-top"">&#xedd3;</i>上传</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "					<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn"" lay-submit lay-filter=""stepNext""><i class=""hr-icon hr-icon-top"">&#xf051;</i>保存</button>"
	tmpHtml = tmpHtml & "<button type=""button"" class=""layui-btn layui-btn-normal"" id=""viewTpl"" title=""查看Excel模板""><i class=""hr-icon hr-icon-top"">&#xf0ce;</i>模板</button>"
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "				</div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "		</form>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""xlsData"" id=""xlsData""></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</fieldset>" & vbCrlf
	Response.Write tmpHtml

	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""form"", ""upload"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, element = layui.element, form = layui.form, upload = layui.upload;" & vbCrlf
	tmpHtml = tmpHtml & "		upload.render({" & vbCrlf		'上传文件
	tmpHtml = tmpHtml & "			elem:'#upExcel', url: '" & InstallDir & "API/UploadFile.htm?UploadDir=Excel'" & vbCrlf
	tmpHtml = tmpHtml & "			,multiple:false, accept:'file',acceptMime:'application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', exts:'xls|xlsx',done: function(res, index){" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#xlsUrl"").val(res.data.src);" & vbCrlf
	tmpHtml = tmpHtml & "			}, error:function (index, upload){ console.log(index); }" & vbCrlf	'错误时显示提示
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		form.on(""submit(stepNext)"", function(PostData){" & vbCrlf
	tmpHtml = tmpHtml & "			var xls1 = PostData.field.xlsUrl, stuType = PostData.field.StudentType, loadImport = layer.load(2);" & vbCrlf
	tmpHtml = tmpHtml & "			if(xls1==""""){layer.tips(""您还没有上传Excel数据文件！"",""#xlsUrl"",{tips: [1, ""#393D49""]});layer.close(loadImport);return false;}" & vbCrlf
	If tStuType <> "" Then tmpHtml = tmpHtml & "			if(stuType==undefined || stuType==""""){layer.tips(""您还没有选择学生类别！"",""#stuType"",{tips:[1, '#393D49']});layer.close(loadImport);return false;}" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#ImportForm"").submit(); return false;" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#viewTpl"").on(""click"", function(){" & vbCrlf			'预览模板
	tmpHtml = tmpHtml & "			var Temp1 = $(""#TempName"").val(), itemName = $(""#ItemName"").val();" & vbCrlf
	tmpHtml = tmpHtml & "			parent.layer.open({type:2,content:""" & ParmPath & "Import/ExcelTemp.html?ItemID=" & tItemID & "&itemName=" & tItemName & """,title:[""查看Excel模板"",""font-size:16""],area:[""95%"", ""72%""],moveOut:true,maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(0)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub SaveExcel()
	Server.ScriptTimeout = 900		'超时900秒(15分钟)

	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tUploadFile : tUploadFile = Trim(ReplaceBadUrl(Request("xlsUrl")))
	Dim tType : tType = Trim(Request("StudentType"))

	'tUploadFile = InstallDir & "Upload/Excel/HR-1550708878x9FHn3.xls"

	Dim strRead, tArrRow, tSheetName, tItemName, tStuType, tTemplate, tFieldLen, TempExcel
	strRead = GetHttpStr(apiHost & "/API/ReadExcel.htm?type=1&xlsFile=" & tUploadFile, "", 1, 10)
	If HR_IsNull(strRead) Then ErrMsg = "Excel数据文件 " & tUploadFile & " 没有记录！" : Response.Write GetErrBody(1) : Exit Sub

	tArrRow = Split(strRead, "@@")		'取行数，第一行为项目名称，第二行为标题，第三行起才是数据
	If Ubound(tArrRow) < 2 Then ErrMsg = "上传的文件没有数据！<br>第三行开始才是业绩数据" : Response.Write GetErrBody(1) : Exit Sub

	tSheetName = "HR_Sheet_" & tItemID
	sqlTmp = "Select a.*,b.FieldsLen From HR_Class a Left Join HR_DataModel b On a.Template=b.ModelName Where a.Template<>'' And a.ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tStuType = Trim(rsTmp("StudentType"))
			tTemplate = Trim(rsTmp("Template"))
			tFieldLen = HR_CLng(rsTmp("FieldsLen"))
			TempExcel = Trim(rsTmp("ExcelFile"))
		Else
			ErrMsg = "非常遗憾，该考核项目不存在！" : Response.Write GetErrBody(1) : Exit Sub
		End If
	Set rsTmp = Nothing

	If ChkTable(tSheetName) = False Then
		ErrMsg = "考核项目 " & tItemName & " 对应的数据表未建立！" : Response.Write GetErrBody(1) : Exit Sub
	End If

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-tmpbox h3 {text-align:center;font-size:1.3rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.err-tips, .succ-tips {display:none;}" & vbCrlf
	tmpHtml = tmpHtml & "		.import-tips {width:75%;margin: 0 auto;padding:15px;border-radius: 5px;position: relative;background-color:#ffffd0;border:1px solid #f19d65;top:2rem;font-size:1.3rem}" & vbCrlf
	tmpHtml = tmpHtml & "		.import-tips span {padding:0 3px;color:#060; font-weight:bold;} .import-tips b {padding:0 3px;color:#f30}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>"

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	tmpHtml = vbCrlf & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>"
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	If Instr(tArrRow(1), "工号") = 0 Then		'上传数据中第一行是否包含标题，以工号为关键字
		Response.Write "<dl class=""hr-tips_dl""><dt><i class=""hr-icon"">&#xef61;</i></dt><dd><h4>数据格式不正确！</h4><p>第一行为项目名称，第二行为标题。<br>正确格式如本页所示！</p></dd></dl>"
		TempExcel = InstallDir & "Upload/ExcelTemp/" & TempExcel
		Response.Write ShowTempExcel(TempExcel)
		Exit Sub
	End If

	Dim colsLen, tArr1
	tArr1 = Split(tArrRow(1), "||")
	colsLen = Ubound(Split(tArrRow(1), "||")) + 1
	If colsLen < tFieldLen Then		'当列数少时
		Response.Write "<dl class=""hr-tips_dl""><dt><i class=""hr-icon"">&#xef61;</i></dt><dd><h4>上传数据不是项目“" & tItemName & "”的数据！</h4><p>您是不是选错考核项了？？</p></dd></dl>"
		Exit Sub
	ElseIf colsLen > tFieldLen Then				'当列数大时（过滤空列）
		If HR_IsNull(tArr1(tFieldLen)) = False Then		'根据最大列的后一列数据为不空时，即不是该项目数据
			Response.Write "<dl class=""hr-tips_dl""><dt><i class=""hr-icon"">&#xef61;</i></dt><dd><h4>上传数据不是项目“" & tItemName & "”的数据！</h4><p>您是不是选错考核项了？？</p></dd></dl>"
			Exit Sub
		End If
	End If

	Dim tYGDM, tYGXM, tKSMC, tKSDM, tYGXB, tPRZC, tNewID, tVA3, tVA4, tVA4Num, tErrData, tSuccData, tEduYear, tTerm
	Dim sqlAdd, rsAdd, strCols, tError, numErr, numAdd
	If Ubound(tArrRow) > 800 Then
		Response.Write "<dl class=""hr-tips_dl""><dt><i class=""hr-icon"">&#xef61;</i></dt><dd><h4>您导入的课程数据超过了800条，请分批上传！</h4><p>共“" & Ubound(tArrRow) + 1 & "”条数据！为保证数据的准确，请分批次导入。<br>辛苦您了.谢谢！<h5>提示：请仔细检查导入的数据中是否有空行数据</h5></p></dd></dl>" : Exit Sub
	End If
	For i = 2 To Ubound(tArrRow)					'开始导入数据
		tError = False
		strCols = ReplaceSQLChar(tArrRow(i))		'替换SQL字符
		tArr1 = Split(strCols, "||")
		tYGDM = HR_CLng(tArr1(1))		'员工代码
		tYGXM = Trim(tArr1(2))			'员工姓名
		tVA3 = HR_CDbl(tArr1(3))		'学时
		If HR_CLng(tArr1(4)) > 20000 Then			'检查是否为时间戳
			tVA4 = FormatDate(ConvertNumDate(tArr1(4)), 2)
			tVA4Num = HR_CLng(tArr1(4))
		ElseIf IsDate(tArr1(4)) Then				'若为日期转为时间戳
			tVA4 = FormatDate(tArr1(4), 2)
			tVA4Num = HR_CLng(ConvertDateToNum(tArr1(4)))
		Else
			tVA4 = Trim(tArr1(4))		'学年的格式必须为2017-2018，学年则为2018【即2018学年表示为：2017-2018，时间段为2017-07-01至2018-06-30】
			tVA4Num = Trim(tArr1(4))
		End If
		tEduYear = GetSchoolYear(tVA4, 2)		'取学年
		tTerm = GetSchoolYear(tVA4, 3)			'取学期

		Set rsAdd = Conn.Execute("Select * From HR_Teacher Where YGDM='" & tYGDM & "'")
			If Not(rsAdd.BOF And rsAdd.EOF) Then		'未找到则从接口读取
				tKSDM = HR_Clng(rsAdd("KSDM"))
				tKSMC = Trim(rsAdd("KSMC"))
				tPRZC = Trim(rsAdd("PRZC"))
				tYGXB = Trim(rsAdd("YGXB"))
			Else
				tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：教师库中没有该名教师，请核对后再试！</li>"
				tError = True
			End If
		Set rsAdd = Nothing

		'检查数据是否符合规定
		
		If tEduYear <> DefYear Then
			tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：该数据不属于本学年[" & tVA4 & "]</li>"
			tError = True
		End If
		If HR_CLng(tArr1(4)) = 0 And Instr(tArr1(4), "-") = 0 Then
			tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：该数据学年格式[" & tArr1(4) & "]不正确</li>"
			tError = True
		End If
		If HR_CLng(tArr1(1)) < 10000 Then
			tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：该数据员工代码不正确[" & tArr1(1) & "]</li>"
			tError = True
		End If
		If HR_IsNull(tArr1(2)) Then
			tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：该数据员工姓名不能为空</li>"
			tError = True
		End If

		'开始判断是否有重复数据
		sqlAdd = "Select * From " & tSheetName & " Where VA1=" & tYGDM & " And VA2='" & Trim(tYGXM) & "' And VA3=" & tVA3
		If HR_IsNull(tType) = False Then sqlAdd = sqlAdd & " And StudentType='" & tType & "'"		'判断是否有学生类别

			Select Case tTemplate
				Case "TempTableA"
					sqlAdd = sqlAdd & " And VA4=" & HR_Clng(tArr1(4)) & " And VA5=" & HR_CLng(tArr1(5)) & " And VA7='" & Trim(tArr1(7)) & "' And VA8='" & Trim(tArr1(8)) & "' "		'判断日期、周次、节次、课程名称
					If HR_IsNull(tArr1(9)) = False Then sqlAdd = sqlAdd & " And VA9='" & Trim(tArr1(9)) & "'"		'判断授课内容
					sqlAdd = sqlAdd & " And VA11='" & Trim(tArr1(11)) & "'"											'判断校区

					If HR_CLng(tArr1(5)) = 0 Or HR_IsNull(tArr1(7)) Or HR_IsNull(tArr1(8)) Or HR_IsNull(tArr1(11)) Then
						tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：该数据周次、节次、课程名称、校区都不能为空</li>"
						tError = True
					End If

				Case "TempTableB"
					sqlAdd = sqlAdd & " And VA4='" & Trim(tArr1(4)) & "' And VA5='" & Trim(tArr1(5)) & "'"			'判断学年(学期)、项目名称
					If HR_IsNull(tArr1(6)) = False Then sqlAdd = sqlAdd & " And Cast(VA6 As nvarchar)='" & Trim(tArr1(6)) & "'"		'判断备注

					If HR_IsNull(tArr1(5)) Then
						tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：该数据项目名称不能为空</li>"
						tError = True
					End If

				Case "TempTableC"
					sqlAdd = sqlAdd & " And VA4=" & HR_Clng(tArr1(4)) & " And VA5='" & Trim(tArr1(5)) & "' And VA6='" & Trim(tArr1(6)) & "'"		'判断日期、学期、项目名称
					If HR_IsNull(tArr1(7)) = False Then sqlAdd = sqlAdd & " And Cast(VA7 As nvarchar)='" & Trim(tArr1(7)) & "'"		'判断备注

					If Instr(tArr1(5), "-") = 0 Or HR_IsNull(tArr1(6)) Then
						tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：该数据学年格式不对或项目名称不能为空</li>"
						tError = True
					End If

				Case "TempTableD"
					sqlAdd = sqlAdd & " And VA4=" & HR_Clng(tArr1(4)) & " And VA5='" & Trim(tArr1(5)) & "' And VA6='" & Trim(tArr1(6)) & "' And VA7='" & Trim(tArr1(7)) & "'"		'判断日期、学期、教材(论文)名称、级别
					If HR_IsNull(tArr1(8)) = False Then sqlAdd = sqlAdd & " And Cast(VA8 As nvarchar)='" & Trim(tArr1(8)) & "'"		'判断备注

					If Instr(tArr1(5), "-") = 0 Or HR_IsNull(tArr1(6)) Then
						tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：该数据学年格式不对或教材(论文)名称不能为空</li>"
						tError = True
					End If

				Case "TempTableE"
					sqlAdd = sqlAdd & " And VA4=" & HR_Clng(tArr1(4)) & " And VA5='" & Trim(tArr1(5)) & "' And VA6='" & Trim(tArr1(6)) & "' And VA7='" & Trim(tArr1(7)) & "'"		'判断日期、学期、项目名称、级别
					If HR_IsNull(tArr1(8)) = False Then sqlAdd = sqlAdd & " And VA8='" & Trim(tArr1(8)) & "'"		'判断等级

					If Instr(tArr1(5), "-") = 0 Or HR_IsNull(tArr1(6)) Then
						tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：该数据学年格式不对或项目名称不能为空</li>"
						tError = True
					End If

				Case "TempTableF"
					sqlAdd = sqlAdd & " And VA4='" & Trim(tArr1(4)) & "' And VA5='" & Trim(tArr1(5)) & "' And VA6='" & Trim(tArr1(6)) & "'"		'判断学期学年、项目名称、级别
					If HR_IsNull(tArr1(7)) = False Then sqlAdd = sqlAdd & " And VA7='" & Trim(tArr1(7)) & "'"		'判断等级

					If HR_IsNull(tArr1(5)) Or HR_IsNull(tArr1(6)) Then
						tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：该数据项目名称或级别不能为空</li>"
						tError = True
					End If

				Case "TempTableG"
					sqlAdd = sqlAdd & " And VA4='" & Trim(tArr1(4)) & "' And VA5='" & Trim(tArr1(5)) & "' And VA6='" & Trim(tArr1(6)) & "'"		'判断学期、项目名称、级别
					If HR_IsNull(tArr1(5)) Or HR_IsNull(tArr1(6)) Then
						tErrData = tErrData & "<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "　错误：该数据项目名称或级别不能为空</li>"
						tError = True
					End If
			End Select
		If tError Then
			numErr = numErr + 1			'计算错误条数
		Else
			Set rsAdd = Server.CreateObject("ADODB.RecordSet")
				rsAdd.Open sqlAdd, Conn, 1, 3
				If Not(rsAdd.BOF And rsAdd.EOF) Then		'数据已经存在
					tErrData = tErrData & "	<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "【" & tVA4 & "】　错误：该数据已经存在</li>"
					tError = True
					numErr = numErr + 1
				Else
					rsAdd.AddNew
					tNewID = GetNewID(tSheetName, "ID")
					rsAdd("ID") = tNewID
					rsAdd("ItemID") = tItemID
					rsAdd("StudentType") = tType
					rsAdd("Passed") = HR_False
					rsAdd("UserID") = UserID
					rsAdd("AppendTime") = Now()
					rsAdd("State") = 0								'导入时状态为0
					rsAdd("KSMC") = tKSMC
					rsAdd("KSDM") = tKSDM
					rsAdd("YGXB") = tYGXB
					rsAdd("PRZC") = tPRZC
					rsAdd("scYear") = tEduYear						'学年，根据VA4计算
					rsAdd("scTerm") = tTerm							'学期

					rsAdd("VA0") = HR_Clng(tArr1(0))				'序号
					rsAdd("VA1") = HR_Clng(tYGDM)					'工号
					rsAdd("VA2") = Trim(tYGXM)						'教师
					rsAdd("VA3") = HR_CDbl(tArr1(3))				'分值
					
					rsAdd("VA4") = tVA4Num							'日期判断
					For m = 5 To tFieldLen-1
						rsAdd("VA" & m) = Trim(tArr1(m))
					Next
					rsAdd.Update
					numAdd = numAdd + 1
					tSuccData = tSuccData & "	<li>序号：" & tArr1(0) & " 工号：" & tArr1(1) & "【" & tVA4 & "】　导入成功！</li>"

					Call UpdateKPIField()		'此处更新业绩表字段
					Call ChkTeacherKPI(tYGDM)	'添加员工信息至业绩表
					Call UpdateTeacherKPI(tItemID, tYGDM, tType)	'更新本项目员工统计数据
					Call UpdateTeacherTotalKPI(tYGDM)	'更新员工总计数据

				End If
			Set rsAdd = Nothing
		End If
	Next
	Response.Write "<h3 class=""import-tips"">成功导入<span>" & HR_CLng(numAdd) & "</span>条数据"
	If HR_CLng(numErr) > 0 Then Response.Write "，其中错误数据<b>" & HR_CLng(numErr) & "</b>条"
	Response.Write "。</h3>" & vbCrlf
	Response.Write "	<ul class=""succ-tips"">" & vbCrlf
	Response.Write tSuccData
	Response.Write "	</ul>" & vbCrlf
	Response.Write "	<ul class=""err-tips"">" & vbCrlf
	Response.Write tErrData
	Response.Write "	</ul>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-pop-fix"">" & vbCrlf
	Response.Write "	<div class=""layui-inline""><button type=""button"" class=""layui-btn layui-btn-sm hr-btn_darkred"" id=""ShowResult"">查看导入结果</button></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-shrink-x20""></div>"

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""element"", ""layer""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#ShowResult"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			sessionStorage.setItem(""succ_data"", $("".succ-tips"").html());" & vbCrlf		'本地缓存成功结果
	tmpHtml = tmpHtml & "			sessionStorage.setItem(""err_data"", $("".err-tips"").html());" & vbCrlf
	'tmpHtml = tmpHtml & "			console.log(window.sessionStorage);" & vbCrlf
	tmpHtml = tmpHtml & "			window.open(""" & ParmPath & "Import/ShowResult.html?key1=succ-tips&key2=err-tips"",""_blank"");" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(0)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub ShowResult()
	SiteTitle = "查看数据导入结果"
	Dim Key1 : Key1 = Trim(ReplaceBadChar(Request("key1")))
	Dim Key2 : Key2 = Trim(ReplaceBadChar(Request("key2")))

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	'tmpHtml = tmpHtml & "		body {background-color:#eee;}"
	tmpHtml = tmpHtml & "	</style>"

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-tab"">" & vbCrlf
	Response.Write "		<ul class=""layui-tab-title"">" & vbCrlf
	Response.Write "			<li class=""layui-this"">成功导入</li>" & vbCrlf
	Response.Write "			<li>导入失败</li>" & vbCrlf
	Response.Write "		</ul>" & vbCrlf
	Response.Write "		<div class=""layui-tab-content"">" & vbCrlf
	Response.Write "			<div class=""layui-tab-item layui-show"">" & vbCrlf
	Response.Write "				<div class=""hr-result-bar""><ul class=""show-succ""></ul></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-tab-item"">" & vbCrlf
	Response.Write "				<div class=""hr-result-bar""><ul class=""show-err""></ul></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	var succ_msg = sessionStorage.getItem(""succ_data"");" & vbCrlf
	tmpHtml = tmpHtml & "	var err_msg = sessionStorage.getItem(""err_data"");" & vbCrlf
	tmpHtml = tmpHtml & "	console.log(window.sessionStorage);" & vbCrlf
	tmpHtml = tmpHtml & "	$("".show-succ"").html(succ_msg);" & vbCrlf
	tmpHtml = tmpHtml & "	$("".show-err"").html(err_msg);" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(0)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub ExcelTemp()
	Dim tClassID : tClassID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tClassName, tExcelFile, tFileName, tTemplate, strData, j
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tClassID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tClassName = rsTmp("ClassName")
			tTemplate = Trim(rsTmp("Template"))
			tFileName = Trim(rsTmp("ExcelFile"))
		End If
	Set rsTmp = Nothing

	Dim hrDate : hrDate = False		'判断VA4是否为日期
	If tTemplate = "TempTableA" Or tTemplate = "TempTableC" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE" Then hrDate = True

	Dim ModelFile : ModelFile = strGetTypeName("HR_DataModel", "ExcelFile", "ModelName", tTemplate)
	If HR_IsNull(tFileName) Then
		ErrMsg = "" & tClassName & " 项目的模板文件未设置！<br>"
		Response.Write GetErrBody(2) : Exit Sub
	Else
		tExcelFile = "/Upload/ExcelTemp/" & tFileName
	End If
	If fso.FileExists(Server.MapPath(tExcelFile)) = False Then
		ErrMsg = "没有找到 " & tClassName & " 项目的模板文件！<br>"
		Response.Write GetErrBody(1) : Exit Sub
	End If

	strData = ShowTempExcel(tExcelFile)

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.downTips {color:#900;line-height:30px;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.Export {color:#f60;font-size:18px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.Export i {font-size: 18px!important;position: relative;top:2px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.tplView th, .tplView td {white-space:nowrap;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	Response.Write "	<legend>" & tClassName & " Excel模板管理</legend>" & vbCrlf
	If UserRank = 2 Then
		Response.Write "		<div class=""hr-shrink-x10"">" & vbCrlf
		Response.Write "			<div><button class=""layui-btn layui-btn-sm layui-btn-normal"" id=""AddBtn"" title=""修改Excel模板文件"" data-cid=""" & tClassID & """><i class=""layui-icon"">&#xe642;</i>设置</button></div>" & vbCrlf
		Response.Write "		</div>" & vbCrlf
	End If
	Response.Write "		<div class=""xlsData"" id=""xlsData"">" & strData & "</div>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf
	Response.Write "<div class=""layer-hr-box ExportBox"">" & vbCrlf
	Response.Write "	<div class=""downTips"">请点击鼠标右键后选择“另存为”</div>" & vbCrlf
	Response.Write "	<div class=""Export""><i class=""hr-icon"">&#xf019;</i><a href=""" & tExcelFile & """ id=""ExcelFile"">" & tExcelFile & "</a></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "	layui.use([""form"", ""table"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var form = layui.form, table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element; layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "		$(""#AddBtn"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			var cid = $(this).data(""cid"");" & vbCrlf
	strHtml = strHtml & "			layer.prompt({title:""修改Excel模板文件"", value:""" & tFileName & """}, function(value, index, elem){" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "KaoItem/SaveExcelName.html"",{ItemID:cid, FileName:value, Field:""AddNew""}, function(reData){" & vbCrlf
	strHtml = strHtml & "					if(reData.Return){window.location.reload();}else{" & vbCrlf
	strHtml = strHtml & "						layer.alert(reData.reMessge, function(index){layer.close(index); return false;});" & vbCrlf
	strHtml = strHtml & "					}" & vbCrlf
	strHtml = strHtml & "				}); return false;" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1) & strHtml
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub

'------ 显示模板文件
Function ShowTempExcel(fExcelFile)		'注意模板文件为全路径
	Dim tmpFun, funArr, funRows, funCols, iFun, strFun : strFun = ""
	If FSO.FileExists(Server.MapPath(fExcelFile)) Then
		tmpFun = GetHttpStr(apiHost & "/API/ReadExcel.htm?type=1&xlsFile=" & fExcelFile, "", 1, 10)
		funRows = Split(tmpFun, "@@")
		If Ubound(funRows) > 0 Then
			funArr = Split(funRows(1), "||")
			funCols = Ubound(funArr) + 1
			strFun = "<div class=""hr-tmpbox"">"
			strFun = strFun & "<h3>" & Replace(funRows(0), "|", "") & "</h3>"
			strFun = strFun & "<table class=""layui-table hr-table-temp"" cellspacing=""0"" cellpadding=""0"" border=""0"" id=""tmptable"">"
			strFun = strFun & "<thead><tr>"
			For iFun = 0 To Ubound(funArr)
				strFun = strFun & "<th>" & funArr(iFun) & "</th>"
			Next
			strFun = strFun & "</tr></thead>"
			If Ubound(funRows) > 1 Then
				strFun = strFun & "<tbody>"
				For m = 2 To Ubound(funRows)
					funArr = Split(funRows(m), "||")
					For iFun = 0 To Ubound(funArr)
						strFun = strFun & "<td>" & funArr(iFun) & "</td>"
					Next
					strFun = strFun & "</tr>"
				Next
				strFun = strFun & "</tbody>"
			End If
			strFun = strFun & "</table></div>"
		Else
			strFun = "<dl class=""hr-tips_dl""><dt><i class=""hr-icon"">&#xef61;</i></dt><dd><h4>模板文件格式不正确！</h4><p>模板文件第一行为项目名称，第二行为标题！</p></dd></dl>"
		End If
	Else
		strFun = "<dl class=""hr-tips_dl""><dt><i class=""hr-icon"">&#xef61;</i></dt><dd><h4>请联系管理员检查模板文件！</h4><p>没有找到“" & fExcelFile & "”文件！</p></dd></dl>"
	End If
	ShowTempExcel = strFun
End Function
%>