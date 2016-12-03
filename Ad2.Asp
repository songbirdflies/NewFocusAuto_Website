<div align="center" style="clear:both;">
 <ul id="Ad1">
   <%
   set Rs=Server.CreateObject("ADODB.RECORDSET")
   Sql="Select * from IdeaNet_Ads where ViewFlag"&LangData&" and SortID=2 order by OrderNum Desc,ID Desc"
   Rs.open Sql,conn,1,1
   do while not Rs.eof%>
   <li style="background:url(<%=SysRootDir%><%=Rs("Content")%>) repeat-x top;cursor:pointer;" onclick="window.location.href='<%=Rs("GotoUrl")%>';"></li>
   <%Rs.movenext
   loop
   Rs.close
   set Rs=nothing%>
 </ul>
</div>
<script type="text/javascript">new dk_slideplayer2("#Ad1",{width:"100%",height:"440px",fontsize:"12px",time:"3000"});</script>