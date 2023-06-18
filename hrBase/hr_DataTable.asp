<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="incCommon.asp"-->
<!--#include file="incPurview.asp"-->
<%
SiteTitle = "数据表"
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "jsonData" Call GetJsonData()

	Case Else Response.Write GetErrBody(0, "", False)
End Select

Sub MainBody()
	Dim tTable : tTable = Trim(Request("Table"))
	tmpHtml = "<link type=""text/css"" href=""[@Web_Dir]Static/css/rb.common.css?v=1.0.3"" rel=""stylesheet"" media=""all"">" & vbCrlf
	tmpHtml = tmpHtml & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.table-name {border-bottom:1px solid #ccc;padding:10px 0 5px}" & vbCrlf
	tmpHtml = tmpHtml & "		.table-name em {font-weight:bold;color:#900} .table-name tt {padding-left:15px;color:#999}" & vbCrlf
	tmpHtml = tmpHtml & "		.field-box {flex-wrap:wrap;justify-content:initial;padding:5px 0} .field-box li {width:160px;padding:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.field-box dl {display:flex;flex-direction:column;border-radius:5px;border:1px solid #ddd;}" & vbCrlf
	tmpHtml = tmpHtml & "		.field-box dt {background-color:#eee;text-align:center;padding:3px 0} .field-box dd {text-align:center;padding:2px;line-height:40px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.table-href {border:1px solid #ccc;padding:8px;position:fixed;background-color:#eee;border-radius:3px;top:10px;right:15px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.table-href .golist {display:none;position:absolute;top:35px;right:-1px;background-color:#fff;border:1px solid #ccc;width:130px;height:90vh;padding:3px;overflow-y:auto;overflow-x:hidden;}" & vbCrlf
	tmpHtml = tmpHtml & "		.table-href .goto {padding:3px 0} .table-href .goto:hover {color:#f30;cursor:pointer;} .table-href p {cursor:pointer;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
	Response.Write ReplaceCommonLabel(strHtml)
	Response.Write "<div class=""hr-workZones hr-shrink-x20 table-wrap"">" & vbCrlf
	Response.Write "	<div class=""table-box""></div>" & vbCrlf
	Response.Write "	<div class=""table-href""><p><i class=""hr-icon"">&#xf329;</i></p><ul class=""golist""></ul></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.js?v=3.7.6""></script>" & vbCrlf
	tmpHtml = tmpHtml & "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""layer"", ""table"", ""form""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var layer=layui.layer, table = layui.table, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		var load1 = layer.load(3);" & vbCrlf
	'tmpHtml = tmpHtml & "		var scroll1 = $("".table-href"").niceScroll();" & vbCrlf
	tmpHtml = tmpHtml & "		$.getJSON(""" & ParmPath & "DataTable/jsonData.html"",{Table:""" & tTable & """},function(res){" & vbCrlf
	tmpHtml = tmpHtml & "			layer.close(load1); $("".table-box"").html(res.FieldData); $("".golist"").html(res.TableName); $(""body"").getNiceScroll().resize(); $("".golist"").getNiceScroll().resize();" & vbCrlf
	tmpHtml = tmpHtml & "			$(document).on(""click"","".goto"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "				var elname=$(this).data(""name""),el=$(""."" + elname), el_scroll=el.offset().top;" & vbCrlf
	tmpHtml = tmpHtml & "				$(document).scrollTop(el_scroll);" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "			$(document).on(""mouseenter"","".table-href"",function(){ $(this).children("".golist"").show();});" & vbCrlf	'显示
	tmpHtml = tmpHtml & "			$(document).on(""mouseleave"","".table-href"",function(){ $(this).children("".golist"").hide();});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	var rollopts = {""horizrailenabled"":false,""cursorwidth"":""10px""}, scroll1 = $(""body"").niceScroll(rollopts), scroll2 = $("".golist"").niceScroll(rollopts);" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
		
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub GetJsonData()
	Dim tField, tArrTable, tTable : tTable = Trim(Request("Table"))
	Dim tCount, tTableName, tFieldData
	sql = "select name,object_id as objID from sys.objects where type ='U' Order By name"
	If HR_IsNull(tTable) = False Then
		tArrTable = Split(tTable, "|") : tTable = ""
		For i=0 To Ubound(tArrTable)
			If i>0 Then tTable = tTable & ","
			tTable = tTable & "'" & Trim(tArrTable(i)) & "'"
		Next
		sql = "select name,object_id as objID from sys.objects where type ='U' And name in (" & tTable & ") Order By name"
	End If
	Set rs = Conn.Execute(sql)
		If Not(rs.BOF And rs.EOF) Then
			i = 1
			Do While Not rs.EOF
				Set rsTmp = Conn.Execute("Select Count(0) From " & rs("name"))
					tCount = HR_Clng(rsTmp(0))
				Set rsTmp = Nothing
				tTableName = tTableName & "<li class=\""goto\"" data-name=\""t" & Trim(rs("objID")) & "\"">" & rs("name") & "</li>"
				tFieldData = tFieldData & "<div class=\""hr-rows table-name t" & Trim(rs("objID")) & "\""><em>" & i & "、" & rs("name") & "</em><tt class=\""hr-grow\"">" & rs("objID") & "【" & tCount & "】</tt></div>"
				sqlTmp = "Select * From " & rs("name") & ""
				Set rsTmp = Server.CreateObject("ADODB.RecordSet")
					rsTmp.Open sqlTmp, Conn, 1, 1
					tFieldData = tFieldData & "<ul class=\""hr-rows field-box\"">"
					k = 0
					For Each tField in rsTmp.Fields
						tFieldData = tFieldData & "<li><dl><dt>" & k & "</dt><dd class=\""hr-ellip\"">" & tField.Name & "</dd></dl></li>"
						k = k + 1
					Next
					tFieldData = tFieldData & "</ul>"
				Set rsTmp = Nothing
				rs.MoveNext
				i = i + 1
			Loop
			Response.Write "{""err"":false,""errcode"":0,""errmsg"":""查询成功"",""icon"":1,""FieldData"":""" & tFieldData & ""","
			Response.Write """TableName"":""" & tTableName & """}" & vbCrlf
		Else
			Response.Write "{""err"":true,""errcode"":500,""errmsg"":""查暂无数据表"",""icon"":2}"
		End If
	Set rs = Nothing
End Sub
%>