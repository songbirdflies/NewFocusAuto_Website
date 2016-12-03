<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<body style="margin:0px 15px 0px 0px; padding:0px; padding-top:20px;background:url(Images/bodybg.jpg) repeat-x top;">
<table width="100%" border="0" cellpadding="0" cellspacing="0" style="font-family:微软雅黑;font-size:12px;color:#fff; background:url(Images/toptitlebg.jpg) repeat-x;">
  <tr>
  	<td width="5%"><img src="Images/title_left.jpg"></td>
    <td width="15%">系统授权：艾迪创意</td>
    <td>管理员：<%=session("AdminName")%>[<%=session("UserName")%>]</td>
    <td width="20%" align="right">位置：后台管理首页</td>
    <td width="5%" align="right"><img src="Images/titleright.jpg"></td>
  </tr>
  </tr>
</table>
</body>