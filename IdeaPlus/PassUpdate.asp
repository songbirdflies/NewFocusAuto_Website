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
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2008-2018 - www.ideanet.cn" />
<META NAME="Author" CONTENT="艾迪创意,www.ideanet.cn" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>修改密码</TITLE>
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="Js/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Md5.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|101,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'判断是否具有管理权限
%>
<body>
<table id="bodycontent" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">

<%
select case request.QueryString("Action")
  case "ModifyPass"
    SaveNewPass
  case else
end select
%>
<table cellpadding="0" cellspacing="0">
  <tr>
    <th>系统管理：添加，修改站点的相关信息</th>
  </tr>
  <tr>
    <td height="24" align="center">
	【管理员密码修改】</td>
  </tr>
</table>
<br>
<table cellpadding="0" cellspacing="0">
  <form name="editForm" method="post" action="PassUpdate.asp?Action=ModifyPass&LoginName=<%=session("AdminName")%>" >
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table class="noborder" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="220" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">登&nbsp;录&nbsp;名：</td>
        <td><input name="AdminName" type="text" class="textfield" id="AdminName" style="WIDTH: 120;" value="<%=session("AdminName")%>" maxlength="16" readonly>&nbsp;3-10位字符，不可修改</td>
      </tr>
      <tr>
        <td height="20" align="right">新&nbsp;密&nbsp;码：</td>
        <td><input name="NewPassword" type="password" class="textfield" id="NewPassword" maxlength="20" style="WIDTH: 120;">&nbsp;*&nbsp;注意字母大小写</td>
      </tr>
      <tr>
        <td height="20" align="right">确认密码：</td>
        <td><input name="vNewPassword" type="password" class="textfield" id="vNewPassword" maxlength="20" style="WIDTH: 120;">&nbsp;*</td>
      </tr>

      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 60;" ></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>

</td></tr></table>
</body>
</html>
<%
function SaveNewPass()
  dim LoginName,rs,sql 
  LoginName=request.QueryString("LoginName")
  set rs = server.createobject("adodb.recordset")
  sql="select * from IdeaNet_Admin where AdminName='"&LoginName&"'"
  rs.open sql,conn,1,3
  if rs.bof and rs.eof then
    response.write "读取数据库记录出错！"
    response.end
  else
	if len(trim(Request.Form("NewPassword")))<6 or len(trim(Request.Form("NewPassword")))>20  then
      response.write "<script language=javascript> alert('管理员密码必填，且字符数为6-20位！');history.back(-1);</script>"
      response.end
    end if
	if Request.Form("NewPassword")<>Request.Form("vNewPassword") then 
      response.write "<script language=javascript> alert('两次输入的密码不一样！');history.back(-1);</script>"
      response.end
	end if
	rs("Password")=Md5(Request.Form("NewPassword"))  
    rs.update
    rs.close
    set rs=nothing 
  end if
  response.write "你的管理密码已成功修改，请牢记[&nbsp;<font color='red'>"&trim(Request.Form("NewPassword"))&"</font>&nbsp;]！"
  response.end
end function
%>
