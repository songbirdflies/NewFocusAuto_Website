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
<meta name="copyright" content="Copyright 2008-2018 - IdeaNet.cn - all rights reserved. - 艾迪创意">
<META NAME="Author" CONTENT="艾迪创意,www.IdeaNet.cn" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>编辑会员组别信息</TITLE>
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|22,")=0 then 
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
dim ID,GroupID,GroupName,GroupLevel,Explain,AddTime,RanNum
ID=request.QueryString("ID")
randomize timer
RanNum=Int((8999)*Rnd +1009)
if ID="" then GroupID=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&RanNum
if Result<>"" then
  call MemGroupEdit()
end if
%>
		<table cellpadding="0" cellspacing="0">
		  <tr>
			<th>会员组别信息：添加，修改会员组别相关的内容</th>
		  </tr>
		  <tr>
			<td align="center">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>会员组别信息】</td>
		  </tr>
		</table>
		<br>
		<table border="0" cellpadding="0" cellspacing="0">
		  <form name="editForm" method="post" action="MemGroupEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
		  <tr>
			<td>
			<table border="0" align="center" class="noborder" cellpadding="0" cellspacing="0">
			  <tr>
				<td width="100" height="20" align="right">&nbsp;</td>
				<td>&nbsp;</td>
			  </tr>
			  <tr>
				<td height="20" align="right">ID：</td>
				<td>
                <input name="ID" type="text" class="textfield" id="ID" style="WIDTH: 240;" value="<%if ID="" then response.write ("自动") else response.write (ID) end if%>" maxlength="100" readonly>&nbsp;</td>
			  </tr>
			  <tr>
				<td height="20" align="right">组别号：</td>
				<td><input name="GroupID" type="text" id="GroupID" style="width: 240;" value="<%=GroupID%>" maxlength="18" readonly> <font color="red">*</font>&nbsp; </td>
			  </tr>
			  <tr>
				<td height="20" align="right">组别名称：</td>
				<td><input name="GroupName" type="text" id="GroupName" style="width: 240" value="<%=GroupName%>"> <font color="red">*</font>&nbsp;</td>
			  </tr>
			  <tr>
				<td height="20" align="right">权限值：</td>
				<td><input name="GroupLevel" type="text" id="GroupLevel" style="width: 240" onChange="if(/\D/.test(this.value)){alert('只能在权限值内输入整数！');}" value="<%=GroupLevel%>" maxlength="5"> 
				<font color="red">*必须为整数</font>&nbsp;</td>
			  </tr>
			  <tr>
				<td height="20" align="right">备注：</td>
				<td><textarea name="Explain" rows="2" class="textfield" id="Explain" style="WIDTH: 86%;"><%=Explain%></textarea>
                </td>
			  </tr>
			  <tr>
				<td height="30" align="right">&nbsp;</td>
				<td><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="  保存  " />&nbsp;</td>
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
sub MemGroupEdit()
  dim Action,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if Result="Add" then
	  sql="select * from IdeaNet_MemGroup"
      rs.open sql,conn,1,3
      rs.addnew
      if len(trim(Request.Form("GroupName")))<3 or len(trim(Request.Form("GroupName")))>16  then
        response.write "<script language='javascript'>alert('请填写会员组别名称(6-16个字符或3-8个汉字)！');history.back(-1);</script>"
        response.end
      end if
	  rs("GroupID")=Request.Form("GroupID")
	  rs("GroupName")=trim(Request.Form("GroupName"))
	  rs("GroupLevel")=trim(Request.Form("GroupLevel"))
	  rs("Explain")=trim(Request.Form("Explain"))
	  rs("AddTime")=now()
	end if
	if Result="Modify" then
      sql="select * from IdeaNet_MemGroup where ID="&ID
      rs.open sql,conn,1,3
      if len(trim(Request.Form("GroupName")))<3 or len(trim(Request.Form("GroupName")))>16  then
        response.write "<script language='javascript'>alert('请填写会员组别名称(6-16个字符或3-8个汉字)！');history.back(-1);</script>"
        response.end
      end if
	  rs("GroupName")=trim(Request.Form("GroupName"))
	  rs("GroupLevel")=trim(Request.Form("GroupLevel"))
	  rs("Explain")=trim(Request.Form("Explain"))
      conn.execute("Update IdeaNet_Members set GroupName='"&trim(Request.Form("GroupName"))&"' where GroupID='"&trim(Request.Form("GroupID"))&"'")
	end if
	  rs("OrderNum")=99
	rs.update
	rs.close
    set rs=nothing
    response.write "<script language='javascript'>alert('成功编辑会员组别信息！');location.replace('MemGroupList.Asp');</script>"
  else
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from IdeaNet_MemGroup where ID="& ID
      rs.open sql,conn,1,1
	  if rs.RecordCount=0 then
        response.write "<script language='javascript'>alert('无此记录！');history.back(-1)</script>"
        response.end
	  end if
	  ID=rs("ID")
      GroupID=rs("GroupID")
	  GroupName=rs("GroupName")
	  GroupLevel=rs("GroupLevel")
	  Explain=rs("Explain")
	  rs.close
      set rs=nothing
	end if
  end if
End Sub
%>