﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'┌┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┐
'┊　　　　　　　　　　　　　　　　　　　　　　　　　┊
'┊　　　　　　艾迪创意企业网站管理系统　　　　　　　┊
'┊　　　　　　　　　　　　　　　　　　　　　　　　　┊
'
'　　设计制作: 艾迪创意技术部
'　　　　　　　Tel:15925772269
'　　　　　　　E-mai1:1036345101@qq.com
'　　　　　　　Q Q:1036345101
'
'　　网    址: http://www.ideanet.cn
'┊　　　　　　　　　　　　　　　　　　　　　　　　　┊
'└┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┘
%> 
<% Option Explicit %>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2008-2018 - www.ideanet.cn" />
<META NAME="Author" CONTENT="艾迪创意,www.ideanet.cn" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>网站信息设置</TITLE>
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="Js/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|11,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'判断是否具有管理权限
%>
<body>
<table id="bodycontent" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
<table cellpadding="0" cellspacing="0">
  <tr>
    <th>系统管理：添加，修改站点的相关信息</th>
  </tr>
  <tr>
    <td height="24" align="center">【系统主参数设置】</td>
  </tr>
</table>
<br>
<table cellpadding="0" cellspacing="0">
  <tr>
    <td height="24">
<br>
<table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td class="LangShowI"><b>生成<%=LangWebI%>静态页面列表</b></td>
	<td class="LangShowD"><b>生成<%=LangWebD%>静态页面列表</b></td>
    <td class="LangShowE"><b>生成<%=LangWebE%>静态页面列表</b></td>
  </tr>
  <tr>
	<td class="LangShowI"><li><a href="Admin_HtmlContent.asp?DataStr=IdeaNet_About&LangData=LangI&MuLu=&UrlNameDiy=<%=AboutNameDiy%>&MuluFrom=&FileFrom=About.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebI%>企业信息列表")'>生成<%=LangWebI%>企业信息列表</a></li></td>
	<td class="LangShowD"><li><a href="Admin_HtmlContent.asp?DataStr=IdeaNet_About&LangData=LangD&MuLu=Cn/&UrlNameDiy=<%=AboutNameDiy%>&MuluFrom=Cn/&FileFrom=About.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebD%>企业信息列表")'>生成<%=LangWebD%>企业信息列表</a></li></td>
	<td class="LangShowE"><li><a href="Admin_HtmlContent.asp?DataStr=IdeaNet_About&LangData=LangE&MuLu=Tr/&UrlNameDiy=<%=AboutNameDiy%>&MuluFrom=Tr/&FileFrom=About.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebE%>企业信息列表")'>生成<%=LangWebE%>企业信息列表</a></li></td>
  </tr>
  <tr>
	<td class="LangShowI"><li><a href="Admin_HtmlSort.asp?DataStr=IdeaNet_News&SortDataStr=IdeaNet_NewsSort&LangData=LangI&MuLu=&UrlNameDiy=<%=NewsSortNameDiy%>&MuluFrom=&FileFrom=News.Asp&PageSizeNum=<%=NewsPageDiy%>" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebI%>新闻分类页面")'>生成<%=LangWebI%>新闻分类页面</a></li></td>
	<td class="LangShowD"><li><a href="Admin_HtmlSort.asp?DataStr=IdeaNet_News&SortDataStr=IdeaNet_NewsSort&LangData=LangD&MuLu=Cn/&UrlNameDiy=<%=NewsSortNameDiy%>&MuluFrom=Cn/&FileFrom=News.Asp&PageSizeNum=<%=NewsPageDiy%>" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebD%>新闻分类页面")'>生成<%=LangWebD%>新闻分类页面</a></li></td>
	<td class="LangShowE"><li><a href="Admin_HtmlSort.asp?DataStr=IdeaNet_News&SortDataStr=IdeaNet_NewsSort&LangData=LangE&MuLu=Tr/&UrlNameDiy=<%=NewsSortNameDiy%>&MuluFrom=Tr/&FileFrom=News.Asp&PageSizeNum=<%=NewsPageDiy%>" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebE%>新闻分类页面")'>生成<%=LangWebE%>新闻分类页面</a></li></td>
  <tr>
	<td class="LangShowI"><li><a href="Admin_HtmlDetails.asp?DataStr=IdeaNet_News&LangData=LangI&MuLu=&UrlNameDiy=<%=NewsNameDiy%>&MuluFrom=&FileFrom=NewsView.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebI%>新闻详细页面")'>生成<%=LangWebI%>新闻详细页面</a></li></td>
	<td class="LangShowD"><li><a href="Admin_HtmlDetails.asp?DataStr=IdeaNet_News&LangData=LangD&MuLu=Cn/&UrlNameDiy=<%=NewsNameDiy%>&MuluFrom=Cn/&FileFrom=NewsView.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebD%>新闻详细页面")'>生成<%=LangWebD%>新闻详细页面</a></li></td>
	<td class="LangShowE"><li><a href="Admin_HtmlDetails.asp?DataStr=IdeaNet_News&LangData=LangE&MuLu=Tr/&UrlNameDiy=<%=NewsNameDiy%>&MuluFrom=Tr/&FileFrom=NewsView.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebE%>新闻详细页面")'>生成<%=LangWebE%>新闻详细页面</a></li></td>
  </tr>
  <tr>
	<td class="LangShowI"><li><a href="Admin_HtmlSort.asp?DataStr=IdeaNet_Products&SortDataStr=IdeaNet_ProductsSort&LangData=LangI&MuLu=&UrlNameDiy=<%=ProductsSortNameDiy%>&MuluFrom=&FileFrom=Products.Asp&PageSizeNum=<%=ProductsPageDiy%>" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebI%>商品/作品分类页面")'>生成<%=LangWebI%>商品/作品分类页面</a></li></td>
	<td class="LangShowD"><li><a href="Admin_HtmlSort.asp?DataStr=IdeaNet_Products&SortDataStr=IdeaNet_ProductsSort&LangData=LangD&MuLu=Cn/&UrlNameDiy=<%=ProductsSortNameDiy%>&MuluFrom=Cn/&FileFrom=Products.Asp&PageSizeNum=<%=ProductsPageDiy%>" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebD%>商品/作品分类页面")'>生成<%=LangWebD%>商品/作品分类页面</a></li></td>
	<td class="LangShowE"><li><a href="Admin_HtmlSort.asp?DataStr=IdeaNet_Products&SortDataStr=IdeaNet_ProductsSort&LangData=LangE&MuLu=Tr/&UrlNameDiy=<%=ProductsSortNameDiy%>&MuluFrom=Tr/&FileFrom=Products.Asp&PageSizeNum=<%=ProductsPageDiy%>" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebE%>商品/作品分类页面")'>生成<%=LangWebE%>商品/作品分类页面</a></li></td>
  <tr>
	<td class="LangShowI"><li><a href="Admin_HtmlDetails.asp?DataStr=IdeaNet_Products&LangData=LangI&MuLu=&UrlNameDiy=<%=ProductsNameDiy%>&MuluFrom=&FileFrom=ProductsView.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebI%>商品/作品详细页面")'>生成<%=LangWebI%>商品/作品详细页面</a></li></td>
	<td class="LangShowD"><li><a href="Admin_HtmlDetails.asp?DataStr=IdeaNet_Products&LangData=LangD&MuLu=Cn/&UrlNameDiy=<%=ProductsNameDiy%>&MuluFrom=Cn/&FileFrom=ProductsView.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebD%>商品/作品详细页面")'>生成<%=LangWebD%>商品/作品详细页面</a></li></td>
	<td class="LangShowE"><li><a href="Admin_HtmlDetails.asp?DataStr=IdeaNet_Products&LangData=LangE&MuLu=Tr/&UrlNameDiy=<%=ProductsNameDiy%>&MuluFrom=Tr/&FileFrom=ProductsView.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebE%>商品/作品详细页面")'>生成<%=LangWebE%>商品/作品详细页面</a></li></td>
  </tr>
  <tr>
	<td class="LangShowI"><li><a href="Admin_HtmlSort.asp?DataStr=IdeaNet_Info&SortDataStr=IdeaNet_InfoSort&LangData=LangI&MuLu=&UrlNameDiy=<%=InfoSortNameDiy%>&MuluFrom=&FileFrom=Info.Asp&PageSizeNum=<%=InfoPageDiy%>" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebI%>信息中心分类页面")'>生成<%=LangWebI%>信息中心分类页面</a></li></td>
	<td class="LangShowD"><li><a href="Admin_HtmlSort.asp?DataStr=IdeaNet_Info&SortDataStr=IdeaNet_InfoSort&LangData=LangD&MuLu=Cn/&UrlNameDiy=<%=InfoSortNameDiy%>&MuluFrom=Cn/&FileFrom=Info.Asp&PageSizeNum=<%=InfoPageDiy%>" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebD%>信息中心分类页面")'>生成<%=LangWebD%>信息中心分类页面</a></li></td>
	<td class="LangShowE"><li><a href="Admin_HtmlSort.asp?DataStr=IdeaNet_Info&SortDataStr=IdeaNet_InfoSort&LangData=LangE&MuLu=Tr/&UrlNameDiy=<%=InfoSortNameDiy%>&MuluFrom=Tr/&FileFrom=Info.Asp&PageSizeNum=<%=InfoPageDiy%>" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebE%>信息中心分类页面")'>生成<%=LangWebE%>信息中心分类页面</a></li></td>
  <tr>
	<td class="LangShowI"><li><a href="Admin_HtmlDetails.asp?DataStr=IdeaNet_Info&LangData=LangI&MuLu=&UrlNameDiy=<%=InfoNameDiy%>&MuluFrom=&FileFrom=InfoView.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebI%>信息中心详细页面")'>生成<%=LangWebI%>信息中心详细页面</a></li></td>
	<td class="LangShowD"><li><a href="Admin_HtmlDetails.asp?DataStr=IdeaNet_Info&LangData=LangD&MuLu=Cn/&UrlNameDiy=<%=InfoNameDiy%>&MuluFrom=Cn/&FileFrom=InfoView.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebD%>信息中心详细页面")'>生成<%=LangWebD%>信息中心详细页面</a></li></td>
	<td class="LangShowE"><li><a href="Admin_HtmlDetails.asp?DataStr=IdeaNet_Info&LangData=LangE&MuLu=Tr/&UrlNameDiy=<%=InfoNameDiy%>&MuluFrom=Tr/&FileFrom=InfoView.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebE%>信息中心详细页面")'>生成<%=LangWebE%>信息中心详细页面</a></li></td>
  </tr>
  <tr>
	<td class="LangShowI"><li><a href="Admin_HtmlContent.asp?DataStr=IdeaNet_Content&LangData=LangI&MuLu=&UrlNameDiy=<%=ContentNameDiy%>&MuluFrom=&FileFrom=Content.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebI%>单页面信息列表")'>生成<%=LangWebI%>单页面信息列表</a></li></td>
	<td class="LangShowD"><li><a href="Admin_HtmlContent.asp?DataStr=IdeaNet_Content&LangData=LangD&MuLu=Cn/&UrlNameDiy=<%=ContentNameDiy%>&MuluFrom=Cn/&FileFrom=Content.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebD%>单页面信息列表")'>生成<%=LangWebD%>单页面信息列表</a></li></td>
	<td class="LangShowE"><li><a href="Admin_HtmlContent.asp?DataStr=IdeaNet_Content&LangData=LangE&MuLu=Tr/&UrlNameDiy=<%=ContentNameDiy%>&MuluFrom=Tr/&FileFrom=Content.Asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebE%>单页面信息列表")'>生成<%=LangWebE%>单页面信息列表</a></li></td>
  </tr>
  <tr>
	<td class="LangShowI"><li><a href="Admin_HtmlindexLangI.asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebI%>首页及其他主页面")'>生成<%=LangWebI%>首页及其他主页面</a></li></td>
	<td class="LangShowD"><li><a href="Admin_HtmlindexLangD.asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebD%>首页及其他主页面")'>生成<%=LangWebD%>首页及其他主页面</a></li></td>
	<td class="LangShowE"><li><a href="Admin_HtmlindexLangE.asp" target="mainFrame" onClick='changeAdminFlag("生成<%=LangWebE%>首页及其他主页面")'>生成<%=LangWebE%>首页及其他主页面</a></li></td>
  </tr>
</table>
<br>
    </td>
  </tr>
</table>
</td></tr></table>
</body>
</html>