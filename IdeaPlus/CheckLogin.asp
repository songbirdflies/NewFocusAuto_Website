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
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2008-2018 - www.ideanet.cn" />
<META NAME="Author" CONTENT="企业网站管理系统,www.ideanet.cn" />
<META NAME="Keywords" CONTENT="企业网站管理系统,www.ideanet.cn" />
<META NAME="Description" CONTENT="企业网站管理系统,www.ideanet.cn" />
<TITLE>检查管理登录</TITLE>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Md5.asp"-->
<BODY>
<%
If cstr(session("getcode"))<>cstr(trim(request("verifycode"))) Then
   response.write "<script language=javascript> alert('您输入验证码错误，请返回重新登录！');location.replace('Login.asp');</script>"
   response.end
end if

dim LoginName,LoginPassword,AdminName,Password,AdminPurview,Working,UserName,rs,sql
LoginName=trim(request.form("LoginName"))
LoginPassword=Md5(request.form("LoginPassword"))
set rs = server.createobject("adodb.recordset")
sql="select * from IdeaNet_Admin where AdminName='"&LoginName&"'"
rs.open sql,conn,1,3

if rs.eof then
   response.write "<script language=javascript> alert('管理员名称不正确，请重新输入。');location.replace('Login.asp');</script>"
   response.end
else
   AdminName=rs("AdminName")
   Password=rs("Password")
   AdminPurview=rs("AdminPurview")
   Working=rs("Working")
   UserName=rs("UserName")
end if

if LoginPassword<>Password then
   response.write "<script language=javascript> alert('管理员密码不正确，请重新输入。');location.replace('Login.asp');</script>"
   response.end
end if 

if not Working then
   response.write "<script language=javascript> alert('不能登录，此管理员帐号已被锁定。');location.replace('Login.asp');</script>"
   response.end
end if 
 
if LoginName=AdminName and LoginPassword=Password then
   rs("LastLoginTime")=now()
   rs("LastLoginIP")=Request.ServerVariables("Remote_Addr")
   rs.update
   rs.close
   set rs=nothing 
   session("AdminName")=AdminName
   session("UserName")=UserName
   session("AdminPurview")=AdminPurview
   session("LoginSystem")="Succeed"
   session.timeout=60
   '==================================
   dim LoginIP,LoginTime,LoginSoft
   LoginIP=Request.ServerVariables("Remote_Addr")
   LoginSoft=Request.ServerVariables("Http_USER_AGENT")
   LoginTime=now()
   '====================================
   set rs = server.createobject("adodb.recordset")
   sql="select * from IdeaNet_AdminLog"
   rs.open sql,conn,1,3
   rs.addnew
   rs("AdminName")=AdminName
   rs("UserName")=UserName
   rs("LoginIP")=LoginIP
   rs("LoginSoft")=LoginSoft
   rs("LoginTime")=LoginTime
   rs.update
   rs.close
   set rs=nothing 
   '========================================
   response.redirect "main.asp"
   response.end
end if
%>
</BODY>
</HTML>