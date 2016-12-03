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
<!-- Admin_htmlSort.asp?DataStr=IdeaNet_News&SortDataStr=IdeaNet_NewsSort&LangData=LangI&MuLu=&UrlNameDiy=Newslist&MuluFrom=&FileFrom=News.Asp&PageSizeNum=10 -->
<!-- LangD -->
<!-- Admin_htmlSort.asp?DataStr=IdeaNet_News&SortDataStr=IdeaNet_NewsSort&LangData=LangD&MuLu=Cn/&UrlNameDiy=Newslist&MuluFrom=Cn/&FileFrom=News.Asp&PageSizeNum=10 -->
<!-- LangE -->
<!-- Admin_htmlSort.asp?DataStr=IdeaNet_News&SortDataStr=IdeaNet_NewsSort&LangData=LangE&MuLu=Tr/&UrlNameDiy=Newslist&MuluFrom=Tr/&FileFrom=News.Asp&PageSizeNum=10 -->
<%
DataStr=Request("DataStr")
SortDataStr=Request("SortDataStr")
LangData=Request("LangData")
MuLu=Request("Mulu")
UrlNameDiy=Request("UrlNameDiy")
MuluFrom=Request("MuluFrom")
FileFrom=Request("FileFrom")
PageSizeNum=Request("PageSizeNum")
totalrec=Conn.Execute("Select count(*) from "&DataStr&" Where ViewFlag"&LangData&"")(0)
totalpage=int(totalrec/PageSizeNum)
If (totalpage * PageSizeNum)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll(""&MuLu&"",""&MuLuFrom&"",""&UrlNameDiy&""&Separated&"1."&HtmlName&"",""&FileFrom&"","Page=",1,"","")
else
for i=1 to totalpage
call htmll(""&MuLu&"",""&MuLuFrom&"",""&UrlNameDiy&""&Separated&""&i&"."&HtmlName&"",""&FileFrom&"","Page=",i,"","")
next
end if
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from "&SortDataStr&" where ViewFlag"&LangData&" order by ID desc"
rs.open sql,conn,1,1
If rs.eof Then
	Class_Num=0
Else
	Class_Num=1
do while not rs.eof
ID=rs("ID")
HtmlUrlName=rs("HtmlUrlName"&LangData)
If DataStr="IdeaNet_Products" then
totalrec=Conn.Execute("Select count(*) from "&DataStr&" where ViewFlag"&LangData&" and Instr(SortClass,',"&ID&",')>0")(0)
ElseIf DataStr="IdeaNet_Info" then
totalrec=Conn.Execute("Select count(*) from "&DataStr&" where ViewFlag"&LangData&" and Instr(SortClass,',"&ID&",')>0")(0)
Else
totalrec=Conn.Execute("Select count(*) from "&DataStr&" where ViewFlag"&LangData&" and SortID="&ID&"")(0)
End If
totalpage=int(totalrec/PageSizeNum)
If (totalpage * PageSizeNum)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
If HtmlUrlName<>"" then
call htmll(""&MuLu&"",""&MuLuFrom&"",""&HtmlUrlName&""&Separated&""&ID&""&Separated&"1."&HtmlName&"",""&FileFrom&"","SortID=",ID,"Page=",1)
else
call htmll(""&MuLu&"",""&MuLuFrom&"",""&UrlNameDiy&""&Separated&""&ID&""&Separated&"1."&HtmlName&"",""&FileFrom&"","SortID=",ID,"Page=",1)
end if
else
for i=1 to totalpage
If HtmlUrlName<>"" then
call htmll(""&MuLu&"",""&MuLuFrom&"",""&HtmlUrlName&""&Separated&""&ID&""&Separated&""&i&"."&HtmlName&"",""&FileFrom&"","SortID=",ID,"Page=",i)
else
call htmll(""&MuLu&"",""&MuLuFrom&"",""&UrlNameDiy&""&Separated&""&ID&""&Separated&""&i&"."&HtmlName&"",""&FileFrom&"","SortID=",ID,"Page=",i)
end if
next
end If
Response.Write "<script>bar_img.width="&Fix((Class_Num/rs.recordcount)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & formatnumber(Class_Num/rs.recordcount*100) & """;</script>"
Response.Flush
rs.movenext
Class_Num=Class_Num+1
Loop
End If
rs.close
set rs=nothing
conn.close
set conn=nothing
%>