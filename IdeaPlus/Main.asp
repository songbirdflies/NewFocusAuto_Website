<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="" />
<META NAME="Author" CONTENT="" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>企业后台管理系统</TITLE>
<link rel="stylesheet" href="css/master.css">
<Script language=JavaScript>
function switchSysBar()
{
   if (switchPoint.innerText==3)
   {
      switchPoint.innerText=4
      document.all("frameTitle").style.display="none"
   }
   else
   {
      switchPoint.innerText=3
      document.all("frameTitle").style.display=""
   }
}
</Script>
</HEAD>
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<BODY scroll="no" topmargin="0" bottom="0" leftmargin="0" rightmargin="0">
<TABLE height="100%" cellSpacing="0" cellPadding="0" border="0" width="100%">
  <TR>
    <TD colSpan="3">
	<iframe style="z-index:1; visibility:inherit; width:100%; height:77px;" name="topFrame" id="topFrame" marginWidth="0" marginHeight="0" src="SysTop.asp" frameBorder="0" noResize scrolling="no"></iframe></TD>
  </TR>
  <TR>
    <TD width="200" height="100%" rowspan="2" align="middle" id="frameTitle">
    <iframe style="z-index:2; visibility:inherit; width:200; height:100%" name="leftFrame" id="leftFrame"  marginWidth="0" marginHeight="0" src="SysLeft.asp" frameBorder="0" noResize scrolling="no"></iframe>
	</TD>
    <TD width="10" height="100%" rowspan="2" align="center" style="background:url(Images/midimg.jpg) no-repeat;" onClick="switchSysBar()">
	<FONT style="FONT-SIZE: 10px; CURSOR: hand; COLOR: #ffffff; FONT-FAMILY: Webdings">
	  <SPAN id="switchPoint">3</SPAN>
	</FONT>
	</TD>
    <TD height="50">
	<iframe style="z-index:3; visibility:inherit; width:100%; height:50" name="headFrame" id="mainFrame" marginwidth="16" marginheight="3" src="SysHead.asp" frameborder="0"  scrolling="no">
	</iframe>
	</TD>
  </TR>
  <TR>
    <TD height="100%" bgcolor="#ade1f6">
	<iframe style="z-index:4; visibility:inherit; width:100%; height:100%" name="mainFrame" id="mainFrame" marginwidth="16" marginheight="16" src="SysCome.asp" frameborder="0" noresize scrolling="yes">
	</iframe>	
	</TD>
  </TR>
</TABLE>
</BODY>
</HTML>
