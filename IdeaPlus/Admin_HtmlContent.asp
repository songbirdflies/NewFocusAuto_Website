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
<br>
<!-- LangI -->
<!-- Admin_htmlContent.asp?DataStr=IdeaNet_About&LangData=LangI&MuLu=&UrlNameDiy=About&MuluFrom=&FileFrom=About.Asp -->
<!-- LangD -->
<!-- Admin_htmlContent.asp?DataStr=IdeaNet_About&LangData=LangD&MuLu=Cn/&UrlNameDiy=About&MuluFrom=Cn/&FileFrom=About.Asp -->
<!-- LangE -->
<!-- Admin_htmlContent.asp?DataStr=IdeaNet_About&LangData=LangE&MuLu=Tr/&UrlNameDiy=About&MuluFrom=Tr/&FileFrom=About.Asp -->
<%
DataStr=Request("DataStr")
LangData=Request("LangData")
MuLu=Request("Mulu")
UrlNameDiy=Request("UrlNameDiy")
MuluFrom=Request("MuluFrom")
FileFrom=Request("FileFrom")
totalrec=Conn.Execute("select count(*) from "&DataStr&" where ViewFlag"&LangData&"")(0)
sql="Select * from "&DataStr&" where ViewFlag"&LangData&" order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
HtmlUrlName=rs("HtmlUrlName"&LangData)
If HtmlUrlName<>"" then
call htmll(""&MuLu&"",""&MuLuFrom&"",""&HtmlUrlName&"."&HtmlName&"",""&FileFrom&"","ID=",ID,"","")
else
call htmll(""&MuLu&"",""&MuLuFrom&"",""&UrlNameDiy&""&Separated&""&ID&"."&HtmlName&"",""&FileFrom&"","ID=",ID,"","")
end if
Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush
rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing
conn.close
set conn=Nothing
%>