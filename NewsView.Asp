<!--#include file="Top.Asp"-->
<%call TitleSortSeo("IdeaNet_News","IdeaNet_NewsSort",""&Request("ID")&"","ViewFlag"&LangData&" and ID="&Request("ID")&"","Error","Error")%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=SeoTitle%> - <%=Keywords%>,<%=SiteTitle%></title>
<meta name="keywords" content="<%=SeoTitle%>,<%=Keywords%>,<%=SiteTitle%>" />
<meta name="description" content="<%=SeoTitle%>,<%=Descriptions%>" />
<meta name="generator" content="<%=SiteTitle%>" />
<meta name="author" content="<%=SiteUrl%>" />
<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.3)" />
<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.3)" />
<!--#include file="Head.Asp"-->
<%if Request.QueryString("SortID")<>"" or SortId<>"" then%>
<%dim SortAddatafrom
  SortAddatafrom="IdeaNet_NewsSort"%>
<!--#include file="SortAd.Asp"-->
<%else%>
<!--#include file="Ad2.Asp"-->
<%end if%>
<!--#include file="Lang.Asp"-->
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0" class="margintop10">
  <tr>
    <td height="400" valign="top">
    <table width="249" border="0" cellpadding="0" cellspacing="0">
 	  <tr>
   		<td height="34" background="Images/l_01_01.gif" class="paddingleft15 f15 c2 bold">News</td>
 	  </tr>
      <tr>
   		<td height="150" valign="top" background="Images/l_01_02.gif">
		<img src="Images/l01.gif" />
		<div align="right"><a href="javascript: void(0);" onmouseover="ScrollDivUpLeft_T=setInterval(ScrollDivUpLeft,10);" onmouseout="clearInterval(ScrollDivUpLeft_T)"><img src="Images/up.gif" /></a></div>
        <div id="itemsleft" style=" widows:249px; height:287px; overflow:hidden">
        <%call ShowList("IdeaNet_News","","ViewFlag"&LangData&"","ViewFlag"&LangData&"","OrderNum,ID desc","TitleListLeft","TitleListLeft",1,25,0,25,0,"left_menu","left_menu",""&NewsNameDiy&"")%>
        </div>
        <div align="right"><a href="javascript: void(0);" onmouseover="ScrollDivDownLeft_T=setInterval(ScrollDivDownLeft,10);" onmouseout="clearInterval(ScrollDivDownLeft_T)"><img src="Images/down.gif" /></a></div>
        <br></td>
 	  </tr>
      <!-- <tr>
   		<td><a href="http://nfaqcyp.tmall.com/"><img src="Images/l_01_03.gif" /></a></td>
 	  </tr> -->
	</table>
    <!--#include file="Left.Asp"-->
    </td>
    <td width="720" valign="top">
    <table width="100%" height="40" border="0" cellpadding="0" cellspacing="0" background="Images/r_01_02.gif">
      <tr>
        <td class="paddingleft10 c2"> current position:<a href="Index.<%=HtmlName%>" class="white">Home</a>  ><a href="<%=NewsSortNameDiy%>.<%=HtmlName%>" class="white">News</a><%if Request("Id")<>"" then%>  ><font class="c2 bold"><%=SeoSort%></font><%end if%></td>
      </tr>
    </table>
    <br><%call Content("IdeaNet_News",""&request.QueryString("ID")&"","ViewFlag"&LangData&"","ViewFlag"&LangData&" and ID="&request.QueryString("ID")&"","OrderNum,ID","Content","class='c4 f16 padding5'")%>
	</td>
  </tr>
</table>
<!--#include file="Foot.Asp"-->