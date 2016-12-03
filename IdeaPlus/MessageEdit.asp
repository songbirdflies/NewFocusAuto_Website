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
''
'　　网    址:  http://www.ideanet.cn
'┊　　　　　　　　　　　　　　　　　　　　　　　　　┊
'└┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┘
%> 
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2008-2018 - www.ideanet.cn" />
<META NAME="Author" CONTENT="艾迪创意,www.ideanet.cn" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>编辑留言</TITLE>
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="Js/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|92,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'判断是否具有管理权限
%>
<BODY>
<table id="bodycontent" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">

<% 
dim Result
Result=request.QueryString("Result")
dim ID,TitleName,Content,UserName,Company,Address,Tel,Mobile,Email,ViewFlag,AddTime,ReplyTime,ReplyName,ReplyContent
ID=request.QueryString("ID")
call MessageEdit() 
%>
<table cellpadding="0" cellspacing="0">
  <tr>
    <th>留言检索及分类查看：添加，修改、审核，删除留言</th>
  </tr>
  <tr>
    <td align="center">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改、审核<%End If%>留言】</td>
  </tr>
</table>
<br>
<table cellpadding="0" cellspacing="0">
  <form name="editForm" method="post" action="MessageEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24"><table  class="noborder" cellpadding="0" cellspacing="0">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr class="LangShowI">
        <td height="20" align="right">留言标题：</td>
        <td><input name="TitleName" type="text" class="textfield" id="TitleName" style="width: 240;" value="<%=TitleName%>" maxlength="100">
        &nbsp;*</td>
      </tr>     
      <tr>
        <td height="20" align="right">留言内容：</td>
        <td><textarea name="Content" rows="3" class="textfield" id="Content" style="width: 86%;"><%=Content%></textarea></td>
      </tr>
	  <tr>
      <td height="20" align="right">姓　　名：</td>
      <td><input name="UserName" type="text" class="textfield" style="width: 240;" value="<%=UserName%>" maxlength="180">&nbsp;</td>
    </tr>
	<tr>
      <td height="20" align="right">单位名称：</td>
      <td><input name="Company" type="text" class="textfield" style="width: 240;" value="<%=Company%>" maxlength="250">&nbsp;</td>
    </tr>
	<tr>
      <td height="20" align="right">地　　址：</td>
      <td><input name="Address" type="text" class="textfield" style="width: 240;" value="<%=Address%>" maxlength="250">&nbsp;</td>
    </tr>
	<tr>
      <td height="20" align="right">联系电话：</td>
      <td><input name="Tel" type="text" class="textfield" style="width: 240;" value="<%=Tel%>" maxlength="180">&nbsp;</td>
    </tr>
	<tr>
      <td height="20" align="right">手　　机：</td>
      <td><input name="Mobile" type="text" class="textfield" style="width: 240;" value="<%=Mobile%>" maxlength="180">&nbsp;</td>
    </tr>
	<tr>
      <td height="20" align="right">邮　　箱：</td>
      <td><input name="Email" type="text" class="textfield" style="width: 240;" value="<%=Email%>" maxlength="180">&nbsp;</td>
    </tr>
	  <tr>
        <td height="20" align="right">审核状态：</td>
        <td>
		<input name="ViewFlag" type="radio" value="1" <%if ViewFlag then response.write ("checked=checked")%> />审核通过
        <input name="ViewFlag" type="radio" value="0" <%if not ViewFlag then response.write ("checked=checked")%>/>未审核</td>
	  </tr>
	  <tr>
      <td align="right">留言时间：</td>
      <td><input name="AddTime" type="text" class="textfield" style="width: 240;" value="<%=AddTime%>" maxlength="180"></td>
    </tr>
	<tr>
      <td align="right">回复时间：</td>
      <td><input name="ReplyTime" type="text" class="textfield" style="width: 240;" value="<%=ReplyTime%>" maxlength="180"></td>
    </tr>
	<tr>
      <td align="right">回 复 人：</td>
      <td><input name="ReplyName" type="text" class="textfield" style="width: 240;" value="<%=ReplyName%>" maxlength="180"></td>
    </tr>
	<tr>
      <td align="right">回复内容：</td>
      <td><textarea name="ReplyContent" rows="3" class="textfield" id="ReplyContent" style="width: 86%;"><%=ReplyContent%></textarea>&nbsp;</td>
    </tr>
      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value=" 保存 " style="width: 80;" ></td>
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
</BODY>
</HTML>
<%
sub MessageEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  
  if Action="SaveEdit" then '保存编辑留言信息
    set rs = server.createobject("adodb.recordset")
	if len(trim(request.Form("TitleName")))<1 then
	   response.write ("<script language=javascript> alert('留言标题为必填！');history.back(-1);</script>")
	   response.end
     end if

	
    if Result="Add" then '添加留言
	  sql="select * from IdeaNet_Message"
      rs.open sql,conn,1,3
      rs.addnew
      rs("TitleName")=trim(Request.Form("TitleName"))
      rs("Content")=trim(Request.Form("Content"))
	  rs("UserName")=trim(Request.Form("UserName"))
      rs("Company")=trim(Request.Form("Company"))
	  rs("Address")=trim(Request.Form("Address"))
	  rs("Tel")=trim(Request.Form("Tel"))
	  rs("Mobile")=trim(Request.Form("Mobile"))
	  rs("Email")=trim(Request.Form("Email"))
	  rs("ViewFlag")=request.Form("ViewFlag")
	  rs("AddTime")=now()
	  if not (trim(request.Form("ReplyContent"))="" or trim(request.Form("ReplyTime"))<>"") then
	  rs("ReplyTime")=now()
	  end if
	  rs("ReplyName")=trim(Request.Form("ReplyName"))
	  rs("ReplyContent")=trim(Request.Form("ReplyContent"))
	  rs("OrderNum")=9999
	end if  
	
	
	if Result="Modify" then '修改留言
      sql="select * from IdeaNet_Message where ID="&ID
      rs.open sql,conn,1,3
      rs("TitleName")=trim(Request.Form("TitleName"))
      rs("Content")=trim(Request.Form("Content"))
	  rs("UserName")=trim(Request.Form("UserName"))
      rs("Company")=trim(Request.Form("Company"))
	  rs("Address")=trim(Request.Form("Address"))
	  rs("Tel")=trim(Request.Form("Tel"))
	  rs("Mobile")=trim(Request.Form("Mobile"))
	  rs("Email")=trim(Request.Form("Email"))
	  rs("ViewFlag")=request.Form("ViewFlag")
	  rs("AddTime")=trim(Request.Form("AddTime"))
	  if not (trim(request.Form("ReplyContent"))="" or trim(request.Form("ReplyTime"))<>"") then
	  rs("ReplyTime")=now()
	  end if
	  rs("ReplyName")=trim(Request.Form("ReplyName"))
	  rs("ReplyContent")=trim(Request.Form("ReplyContent"))
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑、审核留言！');location.replace('MessageList.asp');</script>"
	
  else '提取留言
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from IdeaNet_Message where ID="& ID
      rs.open sql,conn,1,1
      if rs.bof and rs.eof then
        response.write ("数据库读取记录出错！")
        response.end
      end if
	  TitleName=rs("TitleName")
      Content=rs("Content")
	  UserName=rs("UserName")
      Company=rs("Company")
	  Address=rs("Address")
	  Tel=rs("Tel")
	  Mobile=rs("Mobile")
	  Email=rs("Email")
	  ViewFlag=rs("ViewFlag")
	  AddTime=rs("AddTime")
	  ReplyTime=rs("ReplyTime")
	  ReplyName=rs("ReplyName")
	  ReplyContent=rs("ReplyContent")
	  rs.close
      set rs=nothing 
    end if
  end if
end sub
%>