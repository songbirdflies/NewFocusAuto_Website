<!--#include file="TopAsp.Asp"-->
<%SeoTitle="Message"%>
<%If Request("Action")="Save" Then  savepl%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=SeoTitle%> - <%=Keywords%>,<%=SiteTitle%>,<%=ComName%></title>
<meta name="keywords" content="<%=SeoTitle%>,<%=Keywords%>,<%=SiteTitle%>" />
<meta name="description" content="<%=SeoTitle%>,<%=Descriptions%>" />
<meta name="generator" content="<%=SiteTitle%>" />
<meta name="author" content="<%=SiteUrl%>" />
<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.3)" />
<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.3)" />
<!--[if lt IE 7]>
<script defer type="text/javascript" src="Js/pngfix_inline.js"></script>
<script defer type="text/javascript" src="Js/pngfix_bg.js"></script>
<![endif]-->
<style type="text/css">
<!--
body{
	font-family:'微软雅黑',Verdana, Arial, Helvetica, sans-serif;
	font-size:12px;
	margin:0px;
	line-height:150%;}

td{font-family: Verdana, Arial, Helvetica, sans-serif; font-size:12px; color:#919191; line-height:150%;}
#nextpage{font-size:12px; font-family:Arial, Helvetica, sans-serif; clear:both; margin:5px; text-align:center;}
.border{border: solid 1px #c2b08e;}

a:link {color: #919191; text-decoration: none}
a:visited {color: #919191; text-decoration: none}
a:hover {color: #009999; text-decoration: none;}
a:active {color: #009999; text-decoration: none}
-->
</style>
</head>

<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="5">
  <form name="form" method="post" action="?Action=Save" >
  <tr>
    <td width="15%" align="right">姓名：</td>
    <td><input name="name" type="text" id="name" size="40" />
    <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="right">性别：</td>
    <td><select name="sex" id="sex">
      <option value="先生" selected="selected">先生</option>
      <option value="女士">女士</option>
    </select>    </td>
  </tr>、
  <tr>
    <td align="right">电子邮件：</td>
    <td><input name="email" type="text" id="email" size="40" /></td>
  </tr>
  <tr>
    <td align="right">手机号码：</td>
    <td><input name="phone" type="text" id="phone" size="40" />
    <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="right">联系地址：</td>
    <td><input name="address" type="text" id="address" size="40" /></td>
  </tr>
  <tr>
    <td align="right">留言内容：</td>
    <td>
    <textarea name="content" cols="65" rows="5" id="content"></textarea>
    <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="right">验证码：</td>
  <td>
  <input name="verifycode" type="text" id="verifycode"  style="width:80px; height:18px;color:#999" value="验证码" onFocus="if(this.value.indexOf('验证码')==0){this.value='';};this.style.color='#000';" onBlur="if(this.value==''){this.value='验证码';this.style.color='#999';};">
  <img width="55" height="16" src="Include/VerifyCode.asp" onclick="this.src='Include/VerifyCode.asp'" style="cursor:pointer;" class="border" title="看不清楚请点击刷新！"></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><br><a href="javascript:document.form.submit();"><img src="Images/message01.gif" border="0" /></a></td>
  </tr></form>
</table><br><br><br>
</body>
</html>
<%
sub savepl
If cstr(session("getcode"))<>cstr(trim(Request("verifycode"))) Then
response.Write("<script>alert('验证码不正确!');history.back(-1);</script>")
response.End()
end if

dim Mname,Msex,Memail,Mphone,Maddress,Mcontent
Mname=trim(Request.form("name"))
Msex=trim(Request.form("sex"))
Memail=trim(Request.form("email"))
Mphone=trim(Request.form("phone"))
Maddress=trim(Request.form("address"))
Mcontent=trim(Request.form("content"))

if Mname="" or Mphone="" or Mcontent="" then
response.Write("<script>alert('请将表单打*的填写完毕!');history.back(-1);</script>")
response.End()
end if

dim sendname,SystemEmail,mailbody
sendname = "Bangdaitong"
email= "1036345101@qq.com"
SystemEmail= "ray_ocean@126.com" 
mailbody="<html>"
mailbody=mailbody & "<title>客户"&Mname&"向你发来留言</title>"
mailbody=mailbody & "<body>"
mailbody=mailbody & "<TABLE border=0 width='95%' align=center><TR>"
mailbody=mailbody & "<TD valign=middle align=top>"
mailbody=mailbody & "尊敬的("&sendname&")您好：<br><br>"
mailbody=mailbody & "下面是客户"&Mname&"给您的留言信息:<br>"
mailbody=mailbody & "客户姓名："&Mname&"<br>"
mailbody=mailbody & "性别："&Msex&"<br>"
mailbody=mailbody & "电子邮件："&Memail&"<br>"
mailbody=mailbody & "客户手机："&Mphone&"<br>"
mailbody=mailbody & "客户地址："&Maddress&"<br>"
mailbody=mailbody & "留言内容："&Mcontent&"<br>"
mailbody=mailbody & "请您及时联系客户"
mailbody=mailbody & "<br><br>"
mailbody=mailbody & "</TD></TR></TABLE><br><hr width=95% size=1>"
mailbody=mailbody & "</body>" 
mailbody=mailbody & "</html>" 

dim jmail,SendMailed
Set jmail = Server.CreateObject("JMail.Message") 
jmail.silent = true 
jmail.Logging = true 
jmail.Charset = "" 
JMail.ContentType = "text/html" 
jmail.MailServerUserName ="ray_ocean@126.com" '输入smtp服务器验证登陆名 （邮局中任何一个用户的Email地址） 
jmail.MailServerPassword = "515187227" '输入smtp服务器验证密码 （用户Email帐号对应的密码） 
jmail.From = SystemEmail '发件人Email
jmail.FromName = sendname '发件人姓名
jmail.AddRecipient (email) '收件人Email
jmail.Subject = "客户"&Mname&"向你发来留言信息" '信件主题
jmail.Body = mailbody '正文
jmail.Send ("smtp.126.com") 'smtp服务器地址（企业邮局地址） 
set jmail = nothing
SendMailed = "OK"
Response.Write("<script>alert('提交成功，我们客服马上会与您取得联系',location='Message.Asp')</script>")
End Sub%>