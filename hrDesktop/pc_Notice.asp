<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="incCommon.asp"-->
<!--#include file="incPurview.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "查看通知"

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "jsonList" Call GetJsonList()
	Case "Details" Call Details()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim layUrl : layUrl = ParmPath & "Notice/jsonList.html"
		
	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table-cell b {color:#f30;font-weight: normal;}" & vbCrlf	'搜索关键高光
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead("Desktop", 1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	tmpHtml = vbCrlf & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	tmpHtml = "<a href=""" & ParmPath & "Notice/Index.html"">" & SiteTitle & "</a><a><cite>列表</cite></a>"
	strHtml = strHtml & getFrameNav(1) : strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<form name=""SearchForm"" id=""SearchForm"" class=""layui-form soBox"" action="""" method=""get""><div class=""layui-inline"">筛选：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input type=""text"" class=""layui-input"" name=""soWord"" value=""" & soWord & """ id=""soWord"" placeholder=""搜索通知标题"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-btn-normal"" data-type=""reload"" id=""SearchBtn""><i class=""hr-icon hr-icon-top"">&#xeba1;</i>搜索</button></div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf

	Response.Write "	<table class=""layui-table"" id=""TableList"" lay-filter=""TableList""></table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""details"" title=""查看详情""><i class=""hr-icon"">&#xefb9;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "		var tableIns = table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#TableList"",id:""layList"",height:'full-115',page:true,limit:30,skin:'line',limits:[10,15,20,30,50,100,200]" & vbCrlf
	tmpHtml = tmpHtml & "			,text:{none:'暂时没有通知'},cols: [[" & vbCrlf				'设置表头
	tmpHtml = tmpHtml & "				{type:'checkbox',unresize:true,align:'center',width:50}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'ID',title:'序号',sort:true,width:60,align:'center',unresize:true}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Title',title:'标题',width:250}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Intro',title:'内容',minWidth:300,event:'viewIntro',style:'cursor: pointer;'}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'PublishesTime',title:'发布时间',align:'center',width:150}" & vbCrlf
	tmpHtml = tmpHtml & "				,{title:'操作',align:'center',unresize:true,width:80, toolbar:'#barTable'}" & vbCrlf
	tmpHtml = tmpHtml & "			]],url:""" & layUrl & """" & vbCrlf		'设置异步接口
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#SearchBtn"").on(""click"", function(){" & vbCrlf	'搜索
	tmpHtml = tmpHtml & "			var arrForm = $(""#SearchForm"").serializeArray(), postStr={};" & vbCrlf
	tmpHtml = tmpHtml & "			$.each(arrForm, function(key, val){ postStr[this.name]=this.value; });" & vbCrlf		'表单序列转json
	tmpHtml = tmpHtml & "			table.reload(""layList"", {" & vbCrlf
	tmpHtml = tmpHtml & "				url:""" & layUrl & """,where: postStr" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf		'监听工具栏
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""details""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2, id:""viewWin"", content:""" & ParmPath & "Notice/Details.html?ID="" + data.ID, title:""通知内容"", area:[""640px"", ""82%""] });" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot("Desktop", 0) & "" & vbCrlf
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub GetJsonList()
	Dim tmpJson, rsGet, sqlGet
	Dim tLimit : tLimit = HR_Clng(Request("limit"))
	Dim tPage : tPage = HR_Clng(Request("page"))
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))

	sqlGet = "Select *,DATEDIFF(""d"", PublishesTime, getDate()) As Day From HR_Notice Where ID>0"
	If HR_IsNull(soWord) = False Then sqlGet = sqlGet & " And Title like '%" & soWord & "%'"
	sqlGet = sqlGet & " Order By PublishesTime DESC"
	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			i = 0 : CurrentPage = 1 : MaxPerPage = tLimit
			If tPage > 0 Then CurrentPage = tPage
			If MaxPerPage <= 0 Then MaxPerPage = 20
			TotalPut = rsGet.Recordcount
			If TotalPut > 0 Then
				If CurrentPage < 1 Then CurrentPage = 1
				If (CurrentPage - 1) * MaxPerPage > TotalPut Then
					If (TotalPut Mod MaxPerPage) = 0 Then
						CurrentPage = TotalPut \ MaxPerPage
					Else
						CurrentPage = TotalPut \ MaxPerPage + 1
					End If
				End If
				If CurrentPage > 1 Then
					If (CurrentPage - 1) * MaxPerPage < TotalPut Then
						rsGet.Move (CurrentPage - 1) * MaxPerPage
					Else
						CurrentPage = 1
					End If
				End If
			End If
			Dim tTitle, tIntro
			Do While Not rsGet.EOF
				tIntro = nohtml(rsGet("Content")) : tIntro = Replace(nohtml(tIntro), chr(10), "") : tIntro = Replace(nohtml(tIntro), "&nbsp;", "") : tIntro = GetSubStr(tIntro, 110, True)
				tTitle = Trim(rsGet("Title")) : tTitle = Replace(tTitle, soWord, "<b>" & soWord & "</b>")
				If i > 0 Then tmpJson = tmpJson & ","
				tmpJson = tmpJson & "{""ID"":" & rsGet("ID") & ",""Title"":""" & tTitle & """,""Intro"":""" & tIntro & """,""PubDay"":""" & HR_Clng(rsGet("Day")) & """"
				tmpJson = tmpJson & ",""KeyWord"":""" & Trim(rsGet("KeyWord")) & """,""Hits"":" & HR_Clng(rsGet("Hits")) & ",""PublishesTime"":""" & FormatDate(rsGet("PublishesTime"), 1) & """}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""查询成功"",""count"":" & HR_Clng(TotalPut) & ",""data"":[" & tmpJson & "],""limit"":""" & MaxPerPage & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub

Sub Details()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tTitle, tContent, tPublishesTime
	Set rs = Conn.Execute("Select * From HR_Notice Where ID=" & tmpID)
		If Not(rs.BOF And rs.EOF) Then
			tTitle = Trim(rs("Title"))
			tContent = Trim(rs("Content")) : tContent = Replace(tContent, chr(10),"<br>")
			tPublishesTime = FormatDate(rs("PublishesTime"), 1)
		Else
			tTitle = "通知不存在！" & tmpID
		End If
	Set rs = Nothing

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.ViewBox .title {text-align:center;padding:1rem 0;border-bottom:1px solid #ccc;}" & vbCrlf	'搜索关键高光
	tmpHtml = tmpHtml & "		.ViewBox .Content {padding:10px 0;font-size:1.2rem;line-height:1.5;min-height:300px}" & vbCrlf	'搜索关键高光
	tmpHtml = tmpHtml & "		.ViewBox .Content img {max-width:100%;}" & vbCrlf	'搜索关键高光
	tmpHtml = tmpHtml & "		.ViewBox .date {padding:10px;color:#999;border-top:1px solid #ccc;}" & vbCrlf	'搜索关键高光
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead("Desktop", 1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	tmpHtml = vbCrlf & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""ViewBox"">" & vbCrlf
	Response.Write "		<h2 class=""title"">" & tTitle & "</h2>" & vbCrlf
	Response.Write "		<div class=""Content"">" & tContent & "</div>" & vbCrlf
	Response.Write "		<h3 class=""date""><i class=""hr-icon"">&#xe95c;</i>" & tPublishesTime & "</h3>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot("Desktop", 1) & "" & vbCrlf
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
%>