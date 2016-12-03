<!--#include file="Top.Asp"-->
<%call TitleSeo("IdeaNet_ProductsSort",""&request.QueryString("SortID")&"","ViewFlag"&LangData&" and ID="&request.QueryString("SortID")&"","SortName","Concise","PicTitle","Products")%>
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
  SortAddatafrom="IdeaNet_ProductsSort"%>
<!--#include file="SortAd.Asp"-->
<%else%>
<!--#include file="Ad2.Asp"-->
<%end if%>
<!--#include file="ProLang.Asp"-->
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0" class="margintop10">
  <tr>
    <td height="400" valign="top">
    <table width="249" border="0" cellpadding="0" cellspacing="0">
 	  <tr>
   		<td height="34" background="Images/l_01_01.gif" class="paddingleft15 f15 c2 bold">Products</td>
 	  </tr>
      <tr>
   		<td height="150" valign="top" background="Images/l_01_02.gif"><%call ShowList("IdeaNet_ProductsSort",12,"ViewFlag"&LangData&"","ViewFlag"&LangData&" and ParentID=0","OrderNum,ID","SortListLeft","SortListLeft",1,40,10,40,10,"left_menu","left_menu",""&ProductsSortNameDiy&"")%></td>
 	  </tr>
      <!-- <tr>
   		<td><a href="http://nfaqcyp.tmall.com/"><img src="Images/l_01_03.gif" /></a></td>
 	  </tr> -->
	</table>
    <!--#include file="Left.Asp"-->    </td>
    <td width="725" valign="top">
    <table width="100%" height="69" border="0" cellpadding="0" cellspacing="0" background="Images/r_01_00.gif">
      <tr>
        <td width="225">&nbsp;</td>
        <td width="500" valign="top" class="paddingtop8 c2"> current position:<a href="Index.<%=HtmlName%>" class="white">Home</a>  ><a href="<%=ProductsSortNameDiy%>.<%=HtmlName%>" class="white">Products</a>
          <%if Request("SortId")<>"" then%>  ><font class="c2 bold"><%=SeoTitle%></font><%end if%></td>
      </tr>
    </table><br />
	<%call ShowPageList("IdeaNet_Products","",""&Request("SortId")&"","ViewFlag"&LangData&"","ViewFlag"&LangData&" and Instr(SortClass,',"&Request("SortId")&",')>0","OrderNum,ID Desc","Concise",1,50,200,"class='f14 bold white'",""&ProductsNameDiy&"",""&PageUrlName&"",""&ProductsPageDiy&"")%>
    </td>
  </tr>
</table>
<!--#include file="Foot.Asp"-->

