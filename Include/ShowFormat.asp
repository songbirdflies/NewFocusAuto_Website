<%
Sub ShowFormat(Str,LinkStr,TitleNameLength,ConciseLength,CssStr,Rs)%>
<%If Str="MenuTitleName" Then%>
<img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else if Rs("HtmlUrlName"&LangData&"")<>"" then Response.Write(""&LinkStr&"."&HtmlName&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a>
<%end if%>

<%If Str="MenuSortName" Then%>
<div class="lh30"><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&""&Separated&"1."&HtmlName&"")%>" <%=CssStr%>><%=left(Rs("SortName"&LangData),TitleNameLength)%></a></div>
<%end if%>

<%If Str="SortName" Then%>
<a href="<%=Rs("GotoUrl")%>" <%=CssStr%>><%=Rs("SortName"&LangData)%></a>
<%end if%>

<%If Str="TitleListLeft" Then%>
<%if Rs("SmallPic")<>"" then%>
<table width="98%" align="center" border="0" cellspacing="0" cellpadding="0" class="borderbottom margintop5">
  <tr>
    <td align="center" class="paddingbottom5"><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else if Rs("HtmlUrlName"&LangData&"")<>"" then Response.Write(""&LinkStr&"."&HtmlName&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" class="margin10" border="0"></a><br>
    <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else if Rs("HtmlUrlName"&LangData&"")<>"" then Response.Write(""&LinkStr&"."&HtmlName&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><b><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%>
    </b></a>
    </td>
  </tr>
</table>
<%else%>
<div <%if int(Request("Id"))=Rs("Id") then Response.Write("Id='"&CssStr&"2'") else Response.Write("Id='"&CssStr&"1'")%>><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else if Rs("HtmlUrlName"&LangData&"")<>"" then Response.Write(""&LinkStr&"."&HtmlName&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></div>
<%end if%>
<%end if%>

<%If Str="SortListLeft" Then%>
<%if Rs("SmallPic")<>"" then%>
<table width="98%" align="center" border="0" cellspacing="0" cellpadding="0" class="borderbottom margintop5">
  <tr>
    <td align="center" class="paddingbottom5"><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&""&Separated&"1."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" class="margin10" border="0"></a><br>
    <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&""&Separated&"1."&HtmlName&"")%>"><b><%=left(Rs("SortName"&LangData),TitleNameLength)%></b></a>
    </td>
  </tr>
</table>
<%else%>
<div <%if int(Request("SortId"))=Rs("Id") Or SortId=Rs("Id") then Response.Write("Id='"&CssStr&"2'") else Response.Write("Id='"&CssStr&"1'")%>><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&""&Separated&"1."&HtmlName&"")%>"><%=left(Rs("SortName"&LangData),TitleNameLength)%></a></div>
<%end if%>
<%end if%>

<%If Str="Title" Then%>
<div <%=CssStr%>><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a>
</div>
<%end if%>

<%If Str="Pic" Then%>
<a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="155" height="125" class="border margin5" border="0"></a>
<%end if%>

<%If Str="Concise" Then%>
<%=Rs("Concise"&LangData)%>
<%end if%>

<%If Str="TitleTime" Then%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a></td>
    <td width="45" class="c5">[<%=Split(getShortDate(Rs("AddTime")),"-")(1)+"/"+Split(getShortDate(Rs("AddTime")),"-")(2)%>]</td>
  </tr>
</table>
<%end if%>

<%If Str="TitleConcise" Then%>
<div <%=CssStr%>><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a></div>
<%=left(Rs("Concise"&LangData),ConciseLength)%></a>
<img src="Images/Ico/Line01.gif" width="98%" height="3" border="0">
<%end if%>

<%If Str="TitleConciseTime" Then%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" class="borderbottom margintop5">
  <tr>
    <td height="27" valign="top" class="paddingbottom5"><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a>
	<%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"")%></td>
    <td width="145" valign="top" class="paddingbottom5 c5">/ <%=Rs("AddTime")%> /</td>
  </tr>
</table>
<%end if%>

<%If Str="TimeTitleConcise" Then%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" class="borderbottom margintop5">
  <tr>
    <td width="145" valign="top" class="paddingbottom5"><img src="Images/Ico/icon_clock.gif" border="0"/> <%=Rs("AddTime")%></td>
    <td height="27" valign="top" class="paddingbottom5"><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a>
	<%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"")%></td>
  </tr>
</table>
<%end if%>

<%If Str="PicTitle" Then%>
<table width="315" border="0" cellspacing="0" cellpadding="0" class="marginbottom10">
  <tr>
    <td><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="313" height="275" class="border" border="0"></a></td>
  </tr>
  <tr>
    <td height="31" align="center" background="Images/pt01.gif"><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a></td>
  </tr>
</table>
<%end if%>

<%If Str="PicConcise" Then%>
<%if Rs.RecordCount=1 then%>
<div class="ContentImg"><%=Rs("Content"&LangData)%></div>
<%else%>
<table border="0" cellspacing="0" cellpadding="0" class="margin5">
  <tr>
    <td><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="155" height="125" class="border" border="0"></a></td>
  </tr>
  <tr>
    <td height="27"><%=left(Rs("Concise"&LangData),""&ConciseLength&"")%></td>
  </tr>
</table>
<%end if%>
<%end if%>

<%If Str="PicTitleConcise" Then%>
<%if Rs.RecordCount=1 then%>
<div class="ContentImg"><%=Rs("Content"&LangData)%></div>
<%else%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" class="margintop5">
  <tr>
    <td width="90" valign="top" class="paddingbottom5">
	<%if Rs("SmallPic")<>"" then%>
    <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="90" class="border margin5" border="0"></a><br>
	<%end if%><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a><%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"")%></td>
  </tr>
</table>
<%end if%>
<%end if%>

<%If Str="TitlePicConcise" Then%>
<table border="0" cellspacing="0" cellpadding="0" class="margin5">
  <tr>
    <td valign="top">
    <img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a><br>
    <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="297" class="margintop5 border" border="0"></a>
    <%if Rs("Concise"&LangData)<>"" then Response.Write("<Div>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"</div>")%>
    </td>
  </tr>
</table>
<%end if%>

<%If Str="LeftPicConcise" Then%>
<%if Rs.RecordCount=1 then%>
<div class="ContentImg"><%=Rs("Content"&LangData)%></div>
<%else%>
<%if Rs("SmallPic")<>"" then%>
<div style="float:left;"><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="105" class="border margin5" border="0"></a></div><%end if%>
<div style="float:left;"><%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"<br>")%></div>
<img src="Images/Ico/Line01.gif" width="98%" height="3" border="0">
<%end if%>
<%end if%>

<%If Str="LeftPicTitleConcise" Then%>
<%if Rs.RecordCount=1 then%>
<div class="ContentImg"><%=Rs("Content"&LangData)%></div>
<%else%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" class="borderbottom margintop5">
  <tr>
    <%if Rs("SmallPic")<>"" then%>
    <td width="120" valign="top" class="paddingbottom5"><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="100" class="border margin5" border="0"></a></td>
	<%end if%>
    <td valign="top" class="paddingbottom5"><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a><%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"<br>")%></td>
  </tr>
</table>
<%end if%>
<%end if%>

<%If Str="IndexLeftPicTitleConcise" Then%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" class="margintop5">
  <tr>
    <%if Rs("SmallPic")<>"" then%>
    <td width="160" valign="top" class="paddingbottom5"><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="150" class="border margin5" border="0"></a></td>
	<%end if%>
    <td valign="top" class="paddingbottom5"><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a><%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"<br>")%></td>
  </tr>
</table>
<%end if%>

<%If Str="LeftPicTitleConciseIndex" Then%>
<table width="320" height="161" border="0" cellspacing="0" cellpadding="0" class="margintop5" style="background: url(Images/t_01_02.gif) repeat-x top">
  <tr>
    <%if Rs("SmallPic")<>"" then%>
    <td width="88" valign="top" class="padding5"><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="78" class="margin5" border="0"></a></td>
	<%end if%>
    <td valign="top" class="padding5"><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a> <img src="Images/ico04.gif" /><%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"<br>")%><a href="ContactUs.html"><img src="Images/ico01.png" width="76" height="29" border="0" /></a><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="Images/ico02.png" width="76" height="29" border="0"  /></a></td>
  </tr>
</table>
<%end if%>

<%If Str="LeftTitlePicConcise" Then%>
<%if Rs.RecordCount=1 then%>
<div class="ContentImg"><%=Rs("Content"&LangData)%></div>
<%else%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" class="borderbottom margintop5">
  <tr>
    <td><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a></td>
  </tr>
  <%if Rs("SmallPic")<>"" then%>
  <tr>
    <td><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" class="border margin5" border="0"></a></td>
  </tr>
  <%end if%>
  <tr>
    <td class="paddingbottom5"><%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"<br>")%></td>
  </tr>
</table>
<%end if%>
<%end if%>

<%If Str="LeftPicTitleConciseTime" Then%>
<%if Rs.RecordCount=1 then%>
<div class="ContentImg"><%=Rs("Content"&LangData)%></div>
<%else%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" class="borderbottom margintop5">
  <tr>
    <%if Rs("SmallPic")<>"" then%>
    <td width="115" valign="top" class="paddingbottom5"><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="105" class="border margin5" border="0"></a></td>
	<%end if%>
    <td valign="top" class="paddingbottom5"><div style="float:left;"><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a><%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"<br>")%></div><div style="float:right;" class="c5">[<%=Split(getShortDate(Rs("AddTime")),"-")(0)+"/"+Split(getShortDate(Rs("AddTime")),"-")(1)+"/"+Split(getShortDate(Rs("AddTime")),"-")(2)%>]</div></td>
  </tr>
</table>
<%end if%>
<%end if%>

<%If Str="Content" Then%>
<div class="ContentImg"><%=Rs("Content"&LangData)%></div>
<%end if%>

<%If Str="TitleContent" Then%>
<div <%=CssStr%>><%=Rs("TitleName"&LangData)%></div>
<img src="Images/Ico/Line01.gif" width="100%" height="3" border="0">
<div class="ContentImg">
<%if Rs("BigPic")<>"" then%>
<img src="<%=SysRootDir%><%=Rs("BigPic")%>" border="0"><br>
<%end if%>
<%=Rs("Content"&LangData)%>
</div>
<%end if%>

<%If Str="TitleTimeContent" Then%>
<div <%=CssStr%>><%=Rs("TitleName"&LangData)%></div>
<div align="center"> <%if Rs("Source"&LangData)<>"" then Response.Write(""&Rs("Source"&LangData)&"&nbsp;&nbsp;&nbsp;/&nbsp;&nbsp;&nbsp;")%><%=Rs("Addtime")%></div>
<img src="Images/Ico/Line01.gif" width="100%" height="3" border="0">
<div class="ContentImg">
<%if Rs("BigPic")<>"" then%>
<img src="<%=SysRootDir%><%=Rs("BigPic")%>" border="0"><br>
<%end if%>
<%=Rs("Content"&LangData)%>
</div>
<%end if%>

<%If Str="ProList" Then%>
<div <%=CssStr%>><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a></div>
<%end if%>

<%If Str="ProListLeft" Then%>
<div <%if int(Request("Id"))=Rs("Id") then Response.Write("Id='"&CssStr&"2'") else Response.Write("Id='"&CssStr&"1'")%>><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a></div>
<%end if%>

<%If Str="ProGrid" Then%>
<a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="155" height="125" class="border margin5" border="0"></a><br>
<div <%=CssStr%>><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a></div>
<%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"")%>
<%end if%>

<%If Str="ProGridLeft" Then%>
<%if Rs.RecordCount=1 then%>
<div class="ContentImg"><%=Rs("Content"&LangData)%></div>
<%else%>
<%if Rs("SmallPic")<>"" then%>
<div style="float:left;"><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="155" class="border margin5" border="0"></a></div><%end if%>
<div style="float:left;"><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a>
<%if Rs("ProductNo")<>"" then Response.Write("<br><img src='Images/Ico/pro-ico01.png' width=44 height=20 />"&Rs("ProductNo")&"")%>
<%if Rs("Parameters"&LangData)<>"" then Response.Write("<br><img src='Images/Ico/pro-ico02.png' width=44 height=20 />"&Rs("Parameters"&LangData)&"")%>
<%if Rs("Price"&LangData)<>"" then Response.Write("<br><img src='Images/Ico/pro-ico03.png' width=44 height=20 />"&Rs("Price"&LangData)&"")%>
<%if Rs("Concise"&LangData)<>"" then Response.Write("<br><img src='Images/Ico/pro-ico04.png' width=44 height=20 />"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"")%></div>
<img src="Images/Ico/Line01.gif" width="98%" height="3" border="0">
<%end if%>
<%end if%>

<%If Str="ProText" Then%>
<img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a>
<%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"")%>
<img src="Images/Ico/Line01.gif" width="98%" height="3" border="0">
<%end if%>

<%If Str="ProView" Then%>
<div class="ContentImg"><%=Rs("Content"&LangData)%></div>
<%end if%>

<%If Str="DownloadList" Then%>
<%if Rs.RecordCount=1 then%>
<font <%=CssStr%>><%=Rs("TitleName"&LangData)%></font>&nbsp;&nbsp;&nbsp;<font class="c5">/ <%=Rs("AddTime")%> /</font></td><br>
<br>
<img src="Images/Ico/download-ico02.png" width="61" height="20" border="0" /><br>
<a href="<%=SysRootDir%><%=Rs("FileUrl")%>" target="_blank"><img src="Images/Ico/download-ico01.png" width="56" height="51" border="0" /></a><br>
<font class="c4"><img src="Images/Ico/download-ico03.png" width="61" height="20" border="0" /><%=rs("FileSize")%> kb<img src="Images/Ico/Line01.gif" width="98%" height="3" border="0" /><br>
<img src="Images/Ico/download-ico04.png" width="61" height="20" border="0" /><%=rs("System")%></font><br>
<%=Rs("Concise"&LangData)%><br>
<img src="Images/Ico/Line01.gif" width="98%" height="3" border="0">
<div class="ContentImg"><%=Rs("Content"&LangData)%></div>
<%else%>
<%if Rs("SmallPic")<>"" then%>
<div style="float:left;"><a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="105" class="border margin5" border="0"></a></div><%end if%>
<div style="float:left;"><img src="Images/Ico/Title-ico01.gif" border="0"> <a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a><br>
<font class="c4"><img src="Images/Ico/download-ico03.png"  border="0" /><%=rs("FileSize")%> kb<br><img src="Images/Ico/download-ico04.png"  border="0" /><%=rs("System")%></font>
<%if Rs("Concise"&LangData)<>"" then Response.Write("<br>"&left(Rs("Concise"&LangData),""&ConciseLength&"")&"<br>")%></div>
<div style="float:right;"><img src="Images/Ico/icon_clock.gif" border="0"/> <%=Rs("AddTime")%></div>
<img src="Images/Ico/Line01.gif" width="98%" height="3" border="0">
<%end if%>
<%end if%>

<%If Str="DownloadView" Then%>
<font <%=CssStr%>><%=Rs("TitleName"&LangData)%></font>&nbsp;&nbsp;&nbsp;<img src="Images/Ico/icon_clock.gif" /> <%=Rs("AddTime")%><br>
<img src="Images/Ico/Line01.gif" width="98%" height="3" border="0"><br>
<img src="Images/Ico/download-ico02.png"  border="0" /><br>
<a href="<%=SysRootDir%><%=Rs("FileUrl")%>" target="_blank"><img src="Images/Ico/download-ico01.png"  border="0" /></a><br>
<font class="c4"><img src="Images/Ico/download-ico03.png"  border="0" /><%=rs("FileSize")%> kb<br>
<img src="Images/Ico/download-ico04.png"  border="0" /><%=rs("System")%></font><br><%=Rs("Concise"&LangData)%><br>
<img src="Images/Ico/Line01.gif" width="98%" height="3" border="0">
<div class="ContentImg"><%=Rs("Content"&LangData)%></div>
<%end if%>

<%If Str="MesList1" Then%>
<div <%=CssStr%>><img src="Images/ico/ico03.gif" border="0" /> <a href="<%=LinkStr%>?Id=<%=Rs("Id")%>"><%=Rs("TitleName")%></a></div>
<%end if%>

<%If Str="MesList2" Then%>
<table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="25%"><img src="Images/Ico/mes-ico1.png" /><b><%=Rs("TitleName")%></b></td>
    <td align="right"><%=left(Rs("UserName"),6)%>&nbsp;&nbsp;&nbsp;<img src="Images/Ico/mes-ico3.png" /><%=getShortDate(Rs("AddTime"))%></td>
  </tr>
</table>
<table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
    <img src="Images/Ico/Line01.gif" width="100%" height="3" border="0">
    <%=left(Rs("Content"),ConciseLength)%>
	<%if Rs("ReplyContent")<>"" then%><br><img src="Images/Ico/mes-ico4.png" /><br><%=left(Rs("ReplyContent"),50)%><br><%end if%>
    <br><img src="Images/Ico/Line01.gif" width="100%" height="3" border="0">
    </td>
  </tr>
</table>
<%end if%>

<%If Str="MesView" Then%>
<div style="float:left;"><img src="Images/Ico/mes-ico1.png" /><b><%=Rs("TitleName")%></b></div>
<div style="float:right;"><img src="Images/Ico/mes-ico2.png" /><%=left(Rs("UserName"),6)%><img src="Images/Ico/mes-ico3.png" /><%=getShortDate(Rs("AddTime"))%></div>
<img src="Images/Ico/Line01.gif" width="98%" height="3" border="0">
<%=Rs("Content")%>
<%if Rs("ReplyContent")<>"" then%><br><img src="Images/Ico/mes-ico4.png" /><br><%=Rs("ReplyContent")%><br><%end if%>
<br><img src="Images/Ico/Line01.gif" width="98%" height="3" border="0">
<%end if%>

<%If Str="FriendList" Then%>
<%if Rs("Content")<>"" then%><a href="<%=Rs("GotoUrl")%>"><img src="<%=SysRootDir%><%=Rs("Content")%>" border="0" style="margin:5px;"></a><br><%end if%>
<a href="<%=Rs("GotoUrl")%>" <%=CssStr%>><%if len(Rs("TitleName"&LangData))>TitleNameLength then Response.Write left(Rs("TitleName"&LangData),TitleNameLength)&"..." else Response.Write Rs("TitleName"&LangData) end if%></a>
<%end if%>

<%If Str="SortSmallPic" Then%>
<%if Rs("SmallPic")<>"" then%><a href="<%=Rs("GotoUrl")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>"  border="0"></a><%end if%>
<%end if%>

<%If Str="SmallPic" Then%>
<%if Rs("SmallPic")<>"" then%>
<a href="<%if Rs("GotoUrl")<>"" then Response.Write(""&Rs("GotoUrl")&"") else Response.Write(""&LinkStr&""&Separated&""&Rs("Id")&"."&HtmlName&"")%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" class="border" border="0"></a><br><img src="Images/t_04_04.gif" class="marginbottom10" /><%end if%>
<%end if%>

<%If Str="SortBigPic" Then%>
<%if Rs("BigPic")<>"" then%><a href="<%=Rs("GotoUrl")%>"><img src="<%=SysRootDir%><%=Rs("BigPic")%>"  border="0"></a><%end if%>
<%end if%>

<%If Str="BigPic" Then%>
<%if Rs("BigPic")<>"" then%>
<img src="<%=SysRootDir%><%=Rs("BigPic")%>" border="0"></a>
<%end if%>
<%end if%>
<%end sub%>