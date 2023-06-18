<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->

<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl, strNode, tXmlNode
SiteTitle = "数据字典管理"

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "Base" Call MainBody()
	Case "Contact" Call SetContact()
	Case "Common" Call SetCommon()
	Case "Wechat" Call WechatBody()				'微信公众平台接口
	Case "WechatQY" Call WechatBodyQY()			'企业微信接口
	Case "TenQQ" Call TenQQ()					'QQ接口
	Case "SaveConfig" Call SaveConfigXML()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	tXmlNode = "Config"
	tmpHtml = "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.flex-top{align-items:stretch;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	
	tmpHtml = "<a href=""" & ParmPath & "DataDict/Index.html"">" & SiteTitle & "</a><a><cite>通用类参数</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Dim strNode, arrNode : arrNode = Split(XmlText(tXmlNode, "ManageRank", ""), "|")

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Call MenuNavBtn()
	Response.Write "	<div class=""hr-shrink-x10""></div>" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field"">" & vbCrlf
	Response.Write "		<legend>系统字典 - 通用类</legend>" & vbCrlf
	Response.Write "		<div class=""hr-setup-box hr-shrink-x20"">" & vbCrlf
	strNode = XmlText(tXmlNode, "UserLevel", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">会员等级[UserLevel]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""UserLevel"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue"" data-name=""UserLevel""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	strNode = XmlText(tXmlNode, "UserGroup", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar""><dt class=""hr-setup-title"">管理员等级[ManageRank]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""ManageRank"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue"" data-name=""ManageRank""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	strNode = XmlText(tXmlNode, "Module", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar""><dt class=""hr-setup-title"">应用模型[Module]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""Module"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue"" data-name=""Module""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">分页数[MaxPerPage]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "MaxPerPage", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""MaxPerPage"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	strNode = XmlText(tXmlNode, "UserGroup", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar""><dt class=""hr-setup-title"">会员组[UserGroup]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""UserGroup"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	strNode = XmlText(tXmlNode, "Sex", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar""><dt class=""hr-setup-title"">性　别[Sex]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""Sex"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">默认点击量[MaxHits]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "MaxHits", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""MaxHits"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">SEO关键字[MetaKeywords]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "MetaKeywords", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""MetaKeywords"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">SEO描述[MetaDescription]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "MetaDescription", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""MetaDescription"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar flex-top"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">底部版权[CopyRight]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "CopyRight", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""CopyRight"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">版本号[Ver]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "Ver", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""Ver"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "		</div>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Call CommScript()
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub SetContact()
	tXmlNode = "Contact"
	tmpHtml = "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.flex-top{align-items:stretch;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	
	tmpHtml = "<a href=""" & ParmPath & "DataDict/Index.html"">" & SiteTitle & "</a><a><cite>设置联系方式</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Call MenuNavBtn()
	Response.Write "	<div class=""hr-shrink-x10""></div>" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field"">" & vbCrlf
	Response.Write "		<legend>系统字典 - 联系方式</legend>" & vbCrlf
	Response.Write "		<div class=""hr-setup-box hr-shrink-x20"">" & vbCrlf
	
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">单位名称[Company]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "Company", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""Company"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue"" data-name=""UserLevel""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar""><dt class=""hr-setup-title"">英文名称[CompanyEN]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "CompanyEN", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""CompanyEN"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue"" data-name=""ManageRank""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar""><dt class=""hr-setup-title"">简称[CompanyShort]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "CompanyShort", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""CompanyShort"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue"" data-name=""Module""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">地址[Address]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "Address", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""Address"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar""><dt class=""hr-setup-title"">客服电话[Tel1]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "Tel1", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""Tel1"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar""><dt class=""hr-setup-title"">联系电话1[Tel2]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "Tel2", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""Tel2"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">腾讯QQ1[TenQQ1]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "TenQQ1", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""TenQQ1"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">腾讯QQ2[TenQQ2]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "TenQQ2", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""TenQQ2"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">客服邮箱[eMail1]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "eMail1", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""eMail1"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">客服邮箱[eMail2]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "eMail2", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""eMail2"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">传值电话[FaxPhone]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "FaxPhone", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""FaxPhone"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">网站备案号[MII]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "MII", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""MII"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "		</div>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Call CommScript()
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub SetCommon()
	tXmlNode = "Common"

	tmpHtml = "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.flex-top{align-items:stretch;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	
	tmpHtml = "<a href=""" & ParmPath & "DataDict/Index.html"">" & SiteTitle & "</a><a><cite>公用类参数</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Call MenuNavBtn()
	Response.Write "	<div class=""hr-shrink-x10""></div>" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field"">" & vbCrlf
	Response.Write "		<legend>系统字典 - 公用类</legend>" & vbCrlf
	Response.Write "		<div class=""hr-setup-box hr-shrink-x20"">" & vbCrlf

	strNode = XmlText(tXmlNode, "District", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">区域[District]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""District"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	strNode = XmlText(tXmlNode, "Nation", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar flex-top"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">民族[Nation]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""Nation"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	strNode = XmlText(tXmlNode, "NewsStatus", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">新闻状态[NewsStatus]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""NewsStatus"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">百度地图KEY[BaiduMapKey]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "BaiduMapKey", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""BaiduMapKey"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">地图中心点[MapLngLat]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "MapLngLat", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""MapLngLat"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	strNode = XmlText(tXmlNode, "EduLevel", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">教育程度[EduLevel]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""EduLevel"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	strNode = XmlText(tXmlNode, "Bank", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">银行[Bank]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""Bank"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	strNode = XmlText(tXmlNode, "BankCategory", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">银行卡分类[BankCategory]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""BankCategory"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	strNode = XmlText(tXmlNode, "LoanSort", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">借贷类别[LoanSort]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""LoanSort"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	strNode = XmlText(tXmlNode, "AccCategory", "") : strNode = Replace(strNode, "|", "&nbsp;")
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">帐务类别[AccCategory]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & strNode & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""AccCategory"" data-type=""1""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">百度统计Key[BaiduTongjiKey]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "BaiduTongjiKey", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""BaiduTongjiKey"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">友盟统计ID[CnzzID]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "CnzzID", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""CnzzID"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "		</div>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Call CommScript()
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub WechatBody()
	tXmlNode = "WechatConfig"

	tmpHtml = "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.flex-top{align-items:stretch;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	
	tmpHtml = "<a href=""" & ParmPath & "DataDict/Index.html"">" & SiteTitle & "</a><a><cite>微信公众平台接口参数</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Call MenuNavBtn()
	Response.Write "	<div class=""hr-shrink-x10""></div>" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field"">" & vbCrlf
	Response.Write "		<legend>系统字典 - 微信公众平台接口</legend>" & vbCrlf
	Response.Write "		<div class=""hr-setup-box hr-shrink-x20"">" & vbCrlf

	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">公众号名称[wxName]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "wxName", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""wxName"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">微信号[wxCode]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "wxCode", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""wxCode"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">原始ID[wxID]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "wxID", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""wxID"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">开发者ID[wxAppID]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "wxAppID", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""wxAppID"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">开发者密码[wxAppSecret]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "wxAppSecret", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""wxAppSecret"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">令牌[wxToken]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "wxToken", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""wxToken"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">消息密钥[wxAESKey]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "wxAESKey", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""wxAESKey"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">服务器地址[wxSrvURL]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "wxSrvURL", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""wxSrvURL"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">消息回调地址[wxBackURL]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "wxBackURL", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""wxBackURL"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">微信凭证[wxAccessToken]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "wxAccessToken", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""wxAccessToken"" data-type=""0""><button class=""layui-btn layui-btn-sm layui-btn-disabled"" type=""button""><i class=""hr-icon"">&#xee39;</i></button></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">凭证有效期[wxTokenExpires]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "wxTokenExpires", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""wxTokenExpires"" data-type=""0""><button class=""layui-btn layui-btn-sm layui-btn-disabled"" type=""button""><i class=""hr-icon"">&#xee39;</i></button></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf

	Response.Write "		</div>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Call CommScript()
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub WechatBodyQY()
	tXmlNode = "WechatConfig"
	
	tmpHtml = "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.flex-top{align-items:stretch;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	
	tmpHtml = "<a href=""" & ParmPath & "DataDict/Index.html"">" & SiteTitle & "</a><a><cite>企业微信接口参数</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Call MenuNavBtn()
	Response.Write "	<div class=""hr-shrink-x10""></div>" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field"">" & vbCrlf
	Response.Write "		<legend>系统字典 - 企业微信接口</legend>" & vbCrlf
	Response.Write "		<div class=""hr-setup-box hr-shrink-x20"">" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">企业微信Id[qyid]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "qyid", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""qyid"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">应用AgentId[AgentId]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "qyAgentId", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""qyAgentId"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">应用名称[qyAppName]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "qyAppName", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""qyAppName"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">应用Secret[qySecret]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "qySecret", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""qySecret"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">企业微信凭证[qyAccessToken]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "qyAccessToken", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""qyAccessToken"" data-type=""0""><button class=""layui-btn layui-btn-sm layui-btn-disabled"" type=""button""><i class=""hr-icon"">&#xee39;</i></button></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">凭证有效期[qyExpires]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "qyExpires", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""qyExpires"" data-type=""0""><button class=""layui-btn layui-btn-sm layui-btn-disabled"" type=""button""><i class=""hr-icon"">&#xee39;</i></button></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Call CommScript()
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub TenQQ()
	tXmlNode = "WechatConfig"
	
	tmpHtml = "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.flex-top{align-items:stretch;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	
	tmpHtml = "<a href=""" & ParmPath & "DataDict/Index.html"">" & SiteTitle & "</a><a><cite>QQ接口参数</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Call MenuNavBtn()
	Response.Write "	<div class=""hr-shrink-x10""></div>" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field"">" & vbCrlf
	Response.Write "		<legend>系统字典 - QQ接口参数</legend>" & vbCrlf
	Response.Write "		<div class=""hr-setup-box hr-shrink-x20"">" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">QQ AppID[qqAppID]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "qqAppID", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""qqAppID"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">QQ AppKey[qqAppKey]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "qqAppKey", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""qqAppKey"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "			<dl class=""hr-setup-bar"">" & vbCrlf
	Response.Write "				<dt class=""hr-setup-title"">回调地址[qqBackURL]：</dt>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-value""><span>" & XmlText(tXmlNode, "qqBackURL", "") & "</span></dd>" & vbCrlf
	Response.Write "				<dd class=""hr-setup-btn"" data-name=""qqBackURL"" data-type=""0""><a class=""layui-btn layui-btn-sm hr-btn_skyblue""><i class=""hr-icon"">&#xee39;</i></a></dd>" & vbCrlf
	Response.Write "			</dl>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Call CommScript()
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub SaveConfigXML()
	Dim tNode : tNode= Trim(ReplaceBadChar(Request("xmlnode")))
	Dim tName : tName = Trim(ReplaceBadChar(Request("name")))
	Dim tValue : tValue = Trim(Request("value"))
	If HR_IsNull(tNode) = False And HR_IsNull(tName) = False Then
		Call UpdateXmlText(tNode, tName, tValue)
		Response.Write "{""code"":0, ""msg"":""配置数据保存成功！""}"
	Else
		Response.Write "{""code"":500, ""msg"":""Err:Node or NodeName is Null!""}"
	End If
End Sub

Sub MenuNavBtn()
	Response.Write "	<div class=""layui-form soBox"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn""><button class=""layui-btn layui-btn-normal base"" data-type=""base"" id=""base"" title=""基本配置""><i class=""hr-icon hr-icon-top"">&#xee37;</i>基本配置</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-bg-cyan contact"" data-type=""contact"" id=""contact"" title=""联系信息配置""><i class=""hr-icon"">&#xe0ba;</i>联系信息配置</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_deon common"" data-type=""common"" id=""common"" title=""公用参数配置""><i class=""hr-icon"">&#xf260;</i>公用参数配置</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn tenqq"" data-type=""tenqq"" id=""tenqq"" title=""QQ接口参数""><i class=""hr-icon"">&#xf1d6;</i>QQ接口</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_peru wechat"" data-type=""wechat"" id=""wechat"" title=""微信接口参数配置""><i class=""hr-icon"">&#xf1d7;</i>微信接口</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_olive wechat"" data-type=""wechatqy"" id=""wechatqy"" title=""企业微信接口参数""><i class=""hr-icon"">&#xec60;</i>企业微信</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_fuch refresh"" data-type=""refresh"" id=""refresh"" title=""刷新本页""><i class=""hr-icon"">&#xebbb;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
End Sub

Sub CommScript()
	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element"", ""form""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table; element = layui.element, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		$("".searchBtn button"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var btnEvent = $(this).data(""type"");" & vbCrlf
	tmpHtml = tmpHtml & "			if(btnEvent==""base""){;" & vbCrlf
	tmpHtml = tmpHtml & "				location.href=""" & ParmPath & "DataDict/Base.html"";" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(btnEvent==""contact""){" & vbCrlf
	tmpHtml = tmpHtml & "				location.href=""" & ParmPath & "DataDict/Contact.html"";" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(btnEvent==""common""){" & vbCrlf
	tmpHtml = tmpHtml & "				location.href=""" & ParmPath & "DataDict/Common.html"";" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(btnEvent==""wechat""){" & vbCrlf
	tmpHtml = tmpHtml & "				location.href=""" & ParmPath & "DataDict/Wechat.html"";" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(btnEvent==""wechatqy""){" & vbCrlf
	tmpHtml = tmpHtml & "				location.href=""" & ParmPath & "DataDict/WechatQY.html"";" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(btnEvent==""tenqq""){" & vbCrlf
	tmpHtml = tmpHtml & "				location.href=""" & ParmPath & "DataDict/TenQQ.html"";" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(btnEvent==""refresh""){ location.reload(); };" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".hr-setup-btn a"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var node=$(this).parent().data(""name""), prevValue = $(this).parent().prev().find(""span"").html(), nodeType=$(this).parent().data(""type"");" & vbCrlf
	tmpHtml = tmpHtml & "			if(parseInt(nodeType)==1){prevValue=prevValue.replace(/&nbsp;/g,""|"");};" & vbCrlf		'将&nbsp;替换为|
	tmpHtml = tmpHtml & "			layer.prompt({formType:2,value:prevValue,title:""修改字典值"" + node, area:['560px', '300px']},function(value, index, elem){" & vbCrlf
	tmpHtml = tmpHtml & "				console.log(index);" & vbCrlf
	tmpHtml = tmpHtml & "				$.post(""" & ParmPath & "DataDict/SaveConfig.html"",{xmlnode:""" & tXmlNode & """, name:node, value:value}, function(formResult){" & vbCrlf
	tmpHtml = tmpHtml & "					var reData = eval(""("" + formResult + "")"");" & vbCrlf
	tmpHtml = tmpHtml & "					if(reData.code==0){layer.close(index);location.reload();return false;}" & vbCrlf
	tmpHtml = tmpHtml & "					layer.msg(reData.msg,{btn:""关闭"",icon:2});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
End Sub
%>