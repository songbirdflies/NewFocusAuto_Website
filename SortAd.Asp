<%
set Rs=Server.CreateObject("ADODB.RECORDSET")
if SortID<>"" then
Sql="Select * from "&SortAddatafrom&" where ViewFlag"&LangData&" and ID="&SortID&""
else
Sql="Select * from "&SortAddatafrom&" where ViewFlag"&LangData&" and ID="&Request.QueryString("SortID")&""
end if
Rs.open Sql,conn,1,1
If not Rs.eof then%>
<div style="height:550px; background:url(<%=SysRootDir%><%=Rs("BigPic")%>) repeat-x top;cursor:pointer;" onclick="window.location.href='<%=Rs("GotoUrl")%>';"></div>
<%
End If
Rs.close
set Rs=nothing%>