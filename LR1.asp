<table width="737" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="16"><img src="Images/left.gif"border="0" style="cursor:pointer;" onClick="a1.ScrollLeft();" /></td>
    <td><div id="LR1" style="overflow:hidden;">
      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <%set Rs=server.CreateObject("adodb.recordset")
	  Rs.open "select top 10 * from IdeaNet_Info where ViewFlag"&LangData&" and Instr(SortClass,',13,')>0 order by OrderNum,Id Desc",conn,1,1
	  While Not Rs.eof
	  If Rs("HtmlUrlName"&LangData)<>"" then
	  HtmlUrlName=Rs("HtmlUrlName"&LangData)
	  else
	  HtmlUrlName=""&InfoNameDiy&""
	  end if
	  %>
          <td width="141" align="center">
          <div style="width:141px;">
          <a href="<%=HtmlUrlName%><%=Separated%><%=Rs("Id")%>.<%=HtmlName%>"><img src="<%=SysRootDir%><%=Rs("SmallPic")%>" width="119" height="85" border="0" class="border3" /> </a><br><a href="<%=HtmlUrlName%><%=Separated%><%=Rs("Id")%>.<%=HtmlName%>"><%=left(Rs("TitleName"&LangData),16)%> </a></div>
          </td>
          <%Rs.movenext
		  Wend
		  Rs.close
		  set Rs=nothing%>
        </tr>
      </table>
    </div>
      <script type="text/javascript">
	var a1=new scrollImg("LR1",5,145);//滚动对象ID，步长，每次滚动长度［就是图片(width+margin)*n］ 本例为(121+20)*3=423;	stime属性为滚动间隔时间(单位：毫秒),speed属性为步长。
	a1.stime=4500;
	a1.Initial();
    </script>    </td>
    <td width="16" align="right"><img src="Images/right.gif" border="0" style="cursor:pointer;" onclick="a1.ScrollRight();" /></td>
  </tr>
</table>
