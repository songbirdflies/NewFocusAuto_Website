<%
set Rs=Server.CreateObject("ADODB.RECORDSET")
Sql="Select top 2 * from IdeaNet_Ads where ViewFlag"&LangData&" and SortID=3 order by OrderNum Desc,ID Desc"
Rs.open Sql,conn,1,1
do while not Rs.eof%>
<a target="_blank" href="<%=Rs("GotoUrl")%>"><img alt="" src="<%=SysRootDir%><%=Rs("Content")%>" width="206" height="135" class="marginright10" /></a>
<%Rs.movenext
loop
Rs.close
set Rs=nothing%>