<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'┌┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┐
'┊　　　　　　　　　　　　　　　　　　　　　　　　　┊
'┊　　　　　　艾迪创意企业网站管理系统　　　　　　　┊
'┊　　　　　　　　　　　　　　　　　　　　　　　　　┊
'
'　　设计制作: 艾迪创意技术部
'　　　　　　　Tel:15925772269
'　　　　　　　E-mai1:1036345101@qq.com
'　　　　　　　Q Q:1036345101
'
'　　网    址: http://www.ideanet.cn
'┊　　　　　　　　　　　　　　　　　　　　　　　　　┊
'└┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┘
%> 
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>欢迎进入系统后台</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2008-2018 - www.ideanet.cn" />
<META NAME="Author" CONTENT="艾迪创意,www.ideanet.cn" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Css/Content.css">
</HEAD>
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<BODY>
<table id="bodycontent" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">


<table border="0" cellpadding="0" cellspacing="0">
  <tr>
    <th colspan="2"><%=Request.ServerVariables("local_addr")%>【系统信息】</th>
  </tr>
  <tr> 
    <td width="50%">用户名：<%=Session("AdminName")%></td>
    <td width="50%">登录IP：<%=Request.ServerVariables("REMOTE_ADDR")%></td>
  </tr>
  <tr> 
    <td width="50%">身份过期：<%=Session.timeout%> 分钟</font></td>
    <td width="50%">现在时间：<%=year(now())%>年<%=month(now())%>月<%=day(now())%>日<%=hour(now())%>:<%=minute(now())%></font></td>
  </tr>
  <tr>
    <td width="50%">服务器域名：<%=Request.ServerVariables("server_name")%> / <%=Request.ServerVariables("Http_HOST")%></td>
    <td width="50%">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td>站点物理路径：<%=request.ServerVariables("APPL_PHYSICAL_PAth")%></td>
    <td> CDONTS组件：
      <%
	  On Error Resume Next
	  Server.CreateObject("CDONTS.NewMail")
	  if err=0 then 
		 response.write("<b><font color=""red"">√</font></b>")
	  else
         response.write("<b><font color=""red"">×</font></b>")
	  end if
	  err=0
      %>
      Jmail邮箱组件：
    <%
	  If Not IsObjInstalled(theInstalledObjects(13)) Then
         response.write("<b><font color=""red"">×</font></b>") 
      else
         response.write("<b><font color=""red"">√</font></b>") 
      end if
      %></td>
  </tr>
  <tr>
    <td>FSO文本读写：
      <%If Not testObject("scripting.filesystemobject") Then%>
      <b><font color="#FF0000">×</font></b>
      <%else%>
      <b><font color="#FF0000">√</font></b>
    <%end if%></td>
    <td>脚本超时时间：<%=Server.ScriptTimeout%>秒</td>
  </tr>
  <tr>
    <td>客户端操作系统：
    <%
      dim thesoft,vOS
      thesoft=Request.ServerVariables("HTTP_USER_AGENT")
      if instr(thesoft,"Windows NT 5.0") then
	     vOS="Microsoft Windows 2000"
      elseif instr(thesoft,"Windows NT 5.2") then
	     vOs="Microsoft Windows 2003"
      elseif instr(thesoft,"Windows NT 5.1") then
         vOs="Microsoft Windows XP"
      elseif instr(thesoft,"Windows NT") then
       	 vOs="Microsoft Windows NT"
      elseif instr(thesoft,"Windows 9") then
	     vOs="Microsoft Windows 9x"
      elseif instr(thesoft,"unix") or instr(thesoft,"linux") or instr(thesoft,"SunOS") or instr(thesoft,"BSD") then
	     vOs="类Unix"
      elseif instr(thesoft,"Mac") then
	     vOs="Mac"
      else
     	vOs="Other"
      end if
      response.Write(vOs)
      %></td>
    <td>返回服务器处理请求的端口：<%=Request.ServerVariables("SERVER_PORT")%></td>
  </tr>
</table>
<br />
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <th colspan="2">【版本信息】</th>
  </tr>
  <tr>
    <td width="50%">艾迪创意 2008-2018</td>
    <td width="50%">程序开发：<span style="color:#FF0000;">艾迪创意</span>技术部</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font color="red">艾迪咨讯：</font></td>
  </tr>
</table>

</td>
  </tr>
</table>
<br>
<table width="100%" style="font-family:Arial, Helvetica, sans-serif; font-size:12px;" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center">Copyright &copy; 2008-2018 <a href="http://www.ideanet.cn" title="公司理念：客户便捷、客户实惠、客户满意" target="_blank"><font face="Verdana, arial, helvetica, sans-serif" size="1"><b><font color="#CC0000">www.ideanet.cn</font></b></font></a> All Rights Reserved.</td>
  </tr>
</table>
</BODY>
</HTML>
