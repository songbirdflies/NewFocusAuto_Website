<!--#include file="TopAsp.Asp"-->
<%SeoTitle="Search"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=SeoTitle%> - <%=Keywords%>,<%=SiteTitle%></title>
<meta name="keywords" content="<%=SeoTitle%>,<%=Keywords%>,<%=SiteTitle%>" />
<meta name="description" content="<%=SeoTitle%>,<%=Descriptions%>" />
<meta name="generator" content="<%=SiteTitle%>" />
<meta name="author" content="<%=SiteUrl%>" />
<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.3)" />
<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.3)" />
<!--#include file="Head.Asp"-->
<table width="970" border="0" align="center" cellpadding="0" cellspacing="0" class="margintop10">
  <tr>
    <td width="250" height="400" valign="top">
    <table width="231" border="0" cellpadding="0" cellspacing="0">
 	  <tr>
   		<td height="34" background="Images/l_01_01.gif" class="paddingleft15 f15 c2 bold">Products</td>
 	  </tr>
      <tr>
   		<td height="150" valign="top" background="Images/l_01_02.gif"><%call ShowList("IdeaNet_ProductsSort",12,"ViewFlag"&LangData&"","ViewFlag"&LangData&" and ParentID=0","OrderNum,ID","SortListLeft","SortListLeft",1,25,10,25,10,"left_menu","left_menu",""&ProductsSortNameDiy&"")%></td>
 	  </tr>
      <!-- <tr>
   		<td><a href="#"><img src="Images/l_01_03.gif" width="231" height="44" /></a></td>
 	  </tr> -->
	</table>
    <!--#include file="Left.Asp"-->    </td>
    <td width="720" valign="top">
    <table width="100%" height="40" border="0" cellpadding="0" cellspacing="0" background="Images/r_01_01.gif">
      <tr>
        <td class="paddingleft10"><font class="f14 bold c2"><%=SeoTitle%></font></td>
        <td width="450" align="right" class="paddingright10 c2"> current position:<a href="Index.Asp" class="white">Home</a>  >Keyword search:
<font class="f14 c2 bold"><%=searchkey%></font></td>
      </tr>
    </table>
    <div class="padding5">
    <%If searchkey="Enter the keyword" Then searchkey=""%>
	<%call ShowPageList("IdeaNet_Products","",""&searchkey&"","ViewFlag"&LangData&"","ViewFlag"&LangData&" and ProductNo like '%"&searchkey&"%' or TitleName"&LangData&" like '%"&searchkey&"%' or orderNum like '%"&searchkey&"%'","OrderNum,ID desc","PicTitle",2,50,200,"class='f14 bold white'",""&ProductsNameDiy&"","ProductsView.Asp",12)%>
    </div></td>
  </tr>
</table>
<!--#include file="Foot.Asp"-->