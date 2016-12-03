<link rel="stylesheet" type="text/css" href="Css/Style.Css" />
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