<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<title>后台登陆</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script language="javascript" src="Js/Admin.js"></script>
<!--[if lt IE 7]>
<script defer type="text/javascript" src="Js/pngfix_inline.js"></script>
<script defer type="text/javascript" src="Js/pngfix_bg.js"></script>
<![endif]-->
<style type="text/css">
<!--
body {
	margin:0px;
	padding:0px;
	background: url(Images/loginbg01.gif) repeat-x top;
	background-color:#d2ecfa;
	font-size: 12px;
}
.input01{ width:200px; height:20px; background-color:#ffffff; border:0px;}
.input02{ width:100px; height:20px; background-color:#ffffff; margin-top:5px; border:0px;}
-->
</style>
</head>

<body>
<table width="580" border="0" align="center" cellpadding="0" cellspacing="0" background="Images/loginbg02.gif">
  <tr>
    <td height="462" valign="top">
    <table width="480" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:225px;">
    <form action="CheckLogin.asp" method="post" name="AdminLogin" id="AdminLogin"  onSubmit="return CheckAdminLogin()" onkeydown="if(event.keyCode == 13){this.submit();}">
      <tr>
        <td width="95" align="center"><img src="Images/lt01.gif" border="0"/></td>
        <td width="241" height="40" align="center" background="Images/login01.gif"><input name="LoginName" class="input01" type="text" id="LoginName" /></td>
        <td rowspan="5" align="center" valign="top"><a href="#" onclick="AdminLogin.submit(); return false;"><img src="Images/login03.gif" width="120" height="145" border="0" /></a></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        </tr>
      <tr>
        <td align="center"><img src="Images/lt02.gif" border="0"/></td>
        <td height="41" align="center" background="Images/login01.gif"><input name="LoginPassword" class="input01" type="password" id="LoginPassword" /></td>
        </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        </tr>
      <tr>
        <td align="center"><img src="Images/lt03.gif" border="0"/></td>
        <td>
        <table border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="141" height="36" align="center" background="Images/login02.gif"><input name="VerifyCode" class="input02" type="text" id="VerifyCode" maxlength="10" /></td>
            <td width="80" align="center"><img width="55" height="20" src="VerifyCode.asp" onclick="this.src='VerifyCode.asp'" style="cursor:pointer;" title="看不清楚请点击刷新！"></td>
          </tr>
        </table>
        </td>
        </tr>
        </form>
    </table></td>
  </tr>
</table>
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="right"><a href="http://www.ideanet.cn" target="_blank"><img src="Images/copyrightlogo.png" width="291" height="41" border="0" /></a></td>
  </tr>
</table>
</body>
</html>