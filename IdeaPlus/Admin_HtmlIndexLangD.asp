<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="Admin_htmlconfig.asp"-->
<link rel="stylesheet" href="Css/Content.css">
<br />
<table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="300">
    <table width="100%" border="0" cellspacing="0" cellpadding="1">
      <tr>
        <td style="border-bottom: #ccc 1px solid; border-top: #ccc 1px solid; border-left: #ccc 1px solid; border-right: #ccc 1px solid"><img src="Images/Survey_1.gif" width="0" height="16" id="bar_img" name="bar_img" align="absmiddle"></td>
      </tr>
    </table>
    </td>
    <td><span id="bar_txt2" name="bar_txt2" style="font-size:12px; color:red;"></span><span id="bar_txt1" name="bar_txt1" style="font-size:12px">0</span><span style="font-size:12px">%</span></td>
  </tr>
</table>
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
<%
call htmll("/Cn/","/","Index."&HtmlName&"","Cn/Index.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((1/8)*300)&";bar_txt1.innerHTML=""成功生成首页。完成比例" & formatnumber(1/8*100) & """;</script>"
Response.Flush
call htmll("/Cn/","/",""&AboutNameDiy&"."&HtmlName&"","Cn/About.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((2/8)*300)&";bar_txt1.innerHTML=""成功生成“企业信息列表”静态页面。完成比例：" & formatnumber(2/8*100) & """;</script>"
Response.Flush
call htmll("/Cn/","/",""&NewsSortNameDiy&"."&HtmlName&"","Cn/News.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((3/8)*300)&";bar_txt1.innerHTML=""成功生成“新闻列表”静态页面。完成比例：" & formatnumber(3/8*100) & """;</script>"
Response.Flush
call htmll("/Cn/","/",""&ProductsSortNameDiy&"."&HtmlName&"","Cn/Products.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((4/8)*300)&";bar_txt1.innerHTML=""成功生成“产品列表”静态页面。完成比例：" & formatnumber(4/8*100) & """;</script>"
Response.Flush
call htmll("/Cn/","/",""&InfoSortNameDiy&"."&HtmlName&"","Cn/Info.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((6/8)*300)&";bar_txt1.innerHTML=""成功生成“信息中心列表”静态页面。完成比例：" & formatnumber(6/8*100) & """;</script>"
Response.Flush
call htmll("/Cn/","/",""&ContentNameDiy&"."&HtmlName&"","Cn/Content.asp","","","","")
Response.Write "<script>bar_img.width="&Fix((7/8)*300)&";bar_txt1.innerHTML=""成功生成“单页内容列表”静态页面。完成比例：" & formatnumber(7/8*100) & """;</script>"
Response.Flush
Response.Write "<script>bar_img.width=300;bar_txt2.innerHTML='';bar_txt1.innerHTML=""成功生成相关静态文件。完成比例：100"";</script>"
Response.End
%>