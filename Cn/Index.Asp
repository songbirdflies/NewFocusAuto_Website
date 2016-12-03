<!--#include file="Top.Asp"-->
<%SeoTitle="首页"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=Keywords%>,<%=SiteTitle%> - <%=SeoTitle%></title>
<meta name="keywords" content="<%=SeoTitle%>,<%=Keywords%>,<%=SiteTitle%>" />
<meta name="description" content="<%=SeoTitle%>,<%=Descriptions%>" />
<meta name="generator" content="<%=SiteTitle%>" />
<meta name="author" content="<%=SiteUrl%>" />
<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.3)" />
<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.3)" />
<link rel="stylesheet" type="text/css" href="Css/StyleIndex.Css" />
<script type="text/javascript" src="Js/Style.js"></script>
<script type="text/javascript" src="Js/JqueryLR.js"></script>
<script type="text/javascript" src="Js/LRscroll.js"></script>
<script type="text/javascript" src="Js/Marquee.js"></script>
<script type="text/javascript" src="Js/JQh_jquery.js"></script>
<script type="text/javascript" src="Js/JQh_commonv1.1.js"></script>
<script type="text/javascript" src="Js/JQh_commonv1.2.js"></script>
<!--[if lt IE 7]>
<script defer type="text/javascript" src="Js/pngfix_inline.js"></script>
<script defer type="text/javascript" src="Js/pngfix_bg.js"></script>
<![endif]-->

</head>
<body>
<table height="99" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="275"><a href="http://<%=SiteUrl%>"><img src="<%=FilePath%><%=SystemPic%>" border="0"></a></td>
    <%set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.open "Select * from IdeaNet_NavigationMenu where ViewFlag"&LangData&" and ParentId=0 order by OrderNum asc,ID asc",conn,1,1
    ii=1
    do while not Rs.eof%>
    <%if ii=7 then%>
    <%if Rs("SmallPic")<>"" then%>
    <td width="100">
    <div id="menu_nav"><a href="<%=Rs("GotoUrl")%>"><img src="<%=FilePath%><%=Rs("SmallPic")%>" onMouseOver="src='<%=FilePath%><%=Rs("bigPic")%>';" onmouseleave="src='<%=FilePath%><%=Rs("SmallPic")%>';" border="0"></a></div></td>
    <%else%>
    <td width="100"><div id="menu_nav"><a href="<%=Rs("GotoUrl")%>"><%=Rs("SortName"&LangData)%></a></div></td>
    <%end if%>
    <%else%>
    <%if Rs("SmallPic")<>"" then%>
    <td width="100">
    <div id="menu_nav"><a href="<%=Rs("GotoUrl")%>"><img src="<%=FilePath%><%=Rs("SmallPic")%>" onMouseOver="src='<%=FilePath%><%=Rs("bigPic")%>';" onmouseleave="src='<%=FilePath%><%=Rs("SmallPic")%>';" border="0"></a></div></td>
    <td width="2" class="paddingtop20"><img src="Images/menu03.gif" /></td>
    <%else%>
    <td width="100"><div id="menu_nav"><a href="<%=Rs("GotoUrl")%>"><%=Rs("SortName"&LangData)%></a></div></td>
    <td width="2" class="paddingtop20"><img src="Images/menu03.gif" /></td>
    <%end if%>
    <%end if%>
    <%Rs.movenext
	ii=ii+1
    loop
    Rs.close
    set Rs=nothing%>
    <td>&nbsp;</td>
  </tr>
</table>
<!--#include file="Ad1.Asp"-->
<div style="background:url(Images/topbg_01_02.gif)">
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="55" align="right"><a href="../"><img src="Images/En.gif" width="69" height="24" /></a></td>
  </tr>
</table>
</div>
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0" class="margintop10 marginbottom10">
  <tr>
    <td valign="top" class="paddingtop10"><!--#include file="Ad3.Asp"--></td>
    <td width="339" height="150" valign="top">
    <table width="329" border="0" cellspacing="0" cellpadding="0" class="margintop10">
      <tr>
        <td height="32" background="Images/t_01_01.gif" class="paddingleft15 c2 bold">公司简介</td>
      </tr>
      <tr>
        <td class="padding5"><%call Content("IdeaNet_About","","ViewFlag"&LangData&"","ViewFlag"&LangData&" and Id=1","OrderNum,ID","Concise","")%></td>
      </tr>
    </table></td>
    <td width="220" valign="top"><img src="Images/m01.gif" width="100" height="136" class="marginleft10 margintop10" /><img src="Images/m03.gif" width="100" height="136" class="marginleft10 margintop10" /></td>
  </tr>
</table>
<!--#include file="Foot.Asp"-->