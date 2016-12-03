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
<TITLE>编辑友情链接</TITLE>
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="Js/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|16,")=0 then 
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
dim ID,TitleNameLangI,TitleNameLangD,TitleNameLangE,ViewFlagLangI,ViewFlagLangD,ViewFlagLangE,Content,GotoUrl,ConciseLangI,ConciseLangD,ConciseLangE
ID=request.QueryString("ID")
call FriendSiteEdit() 
%>
	
		<table cellpadding="0" cellspacing="0">
		  <tr>
			<th>友情链接：添加，修改友情链接相关的内容</th>
		  </tr>
		  <tr>
			<td align="center">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>友情链接】</td>
		  </tr>
		</table>
		<br>
		<table border="0" cellpadding="0" cellspacing="0">
		  <form name="editForm" method="post" action="FriendSiteEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
		  <tr>
			<td>
			<table border="0" align="center" class="noborder" cellpadding="0" cellspacing="0">
			  <tr>
				<td width="100" height="20" align="right">&nbsp;</td>
				<td>&nbsp;</td>
			  </tr>
			  <tr class="LangShowI">
				<td height="20" align="right"><%=LangWebI%>名称：</td>
				<td><input name="TitleNameLangI" type="text" class="textfield" id="TitleNameLangI" style="WIDTH: 240;" value="<%=TitleNameLangI%>" maxlength="100">&nbsp;显示：<input name="ViewFlagLangI" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlagLangI then response.write ("checked")%>>
		&nbsp;不少于1个字符</td>
			  </tr>
			  <tr class="LangShowD">
				<td height="20" align="right"><%=LangWebD%>名称：</td>
				<td><input name="TitleNameLangD" type="text" class="textfield" id="TitleNameLangD" style="WIDTH: 240;" value="<%=TitleNameLangD%>" maxlength="100">&nbsp;显示：<input name="ViewFlagLangD" type="checkbox" value="1" style='HEIGHT: 13px;WIDTH: 13px;' <%if ViewFlagLangD then response.write ("checked")%>></td>
			  </tr>
			  <tr class="LangShowE">
				<td height="20" align="right"><%=LangWebE%>名称：</td>
				<td><input name="TitleNameLangE" type="text" class="textfield" id="TitleNameLangE" style="WIDTH: 240;" value="<%=TitleNameLangE%>" maxlength="100">&nbsp;显示：<input name="ViewFlagLangE" type="checkbox" value="1" style='HEIGHT: 13px;WIDTH: 13px;' <%if ViewFlagLangE then response.write ("checked")%>></td>
			  </tr>
			  <tr>
				<td height="20" align="right">链接地址：</td>
				<td><textarea name="Content" rows="1" class="textfield" style="WIDTH: 56%;"><%=Content%></textarea>
				&nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=Content',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
			  </tr>
			  <tr>
				<td height="20" align="right">链接网址：</td>
				<td><input name="GotoUrl" type="text" class="textfield" id="GotoUrl" style="WIDTH: 240;" value="<%=GotoUrl%>" maxlength="100">
				<img src="Images/memo.gif" alt="无链接地址请为空" /></td>
			  </tr>
			  <tr class="LangShowI">
				<td height="20" align="right" valign="top"><%=LangWebI%>简要概述：</td>
				<td><textarea name="ConciseLangI" rows="2" class="textfield" id="ConciseLangI" style="WIDTH: 86%;"><%=ConciseLangI%></textarea>&nbsp;</td>
			  </tr>
			  <tr class="LangShowD">
				<td height="20" align="right" valign="top"><%=LangWebD%>简要概述：</td>
				<td><textarea name="ConciseLangD" rows="2" class="textfield" id="ConciseLangD" style="WIDTH: 86%;"><%=ConciseLangD%></textarea>&nbsp;</td>
			  </tr>
			  <tr class="LangShowE">
				<td height="20" align="right" valign="top"><%=LangWebE%>简要概述：</td>
				<td><textarea name="ConciseLangE" rows="2" class="textfield" id="ConciseLangE" style="WIDTH: 86%;"><%=ConciseLangE%></textarea>&nbsp;</td>
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
sub FriendSiteEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑信息
    set rs = server.createobject("adodb.recordset")
    if request.Form("ViewFlagLangI")=1 then
	  if len(trim(request.Form("TitleNameLangI")))<1 then
	  response.write ("<script language=javascript> alert('你已经选择了"""&LangWebI&"""显示，因此"&LangWebI&"名称必填！');history.back(-1);</script>")
	  response.end
	  end if
	end if
    if request.Form("ViewFlagLangD")=1 then
      if len(trim(request.Form("TitleNameLangD")))<1 then
      response.write ("<script language=javascript> alert('您已经选择了"""&LangWebD&"""显示，因此"&LangWebD&"名称必填！');history.back(-1);</script>")
      response.end
      end if
    end if
    if request.Form("ViewFlagLangE")=1 then
      if len(trim(request.Form("TitleNameLangE")))<1 then
      response.write ("<script language=javascript> alert('您已经选择了"""&LangWebE&"""显示，因此"&LangWebE&"名称必填！');history.back(-1);</script>")
      response.end
      end if
    end if
	
    if Result="Add" then '添加
	  sql="select * from IdeaNet_FriendSite"
      rs.open sql,conn,1,3
      rs.addnew
      rs("TitleNameLangI")=trim(Request.Form("TitleNameLangI"))
      rs("TitleNameLangD")=trim(Request.Form("TitleNameLangD"))
      rs("TitleNameLangE")=trim(Request.Form("TitleNameLangE"))
	  if Request.Form("ViewFlagLangI")=1 then
        rs("ViewFlagLangI")=Request.Form("ViewFlagLangI")
	  else
        rs("ViewFlagLangI")=0
	  end if
	  if Request.Form("ViewFlagLangD")=1 then
        rs("ViewFlagLangD")=Request.Form("ViewFlagLangD")
	  else
        rs("ViewFlagLangD")=0
	  end if
	  if Request.Form("ViewFlagLangE")=1 then
        rs("ViewFlagLangE")=Request.Form("ViewFlagLangE")
	  else
        rs("ViewFlagLangE")=0
	  end if
	  rs("Content")=trim(Request.Form("Content"))  
	  rs("GotoUrl")=Request.Form("GotoUrl")
	  rs("ConciseLangI")=Request.Form("ConciseLangI")
	  rs("ConciseLangD")=Request.Form("ConciseLangD")
	  rs("ConciseLangE")=Request.Form("ConciseLangE")
	  rs("OrderNum")=99
	  rs("AddTime")=now()
	end if
	
	if Result="Modify" then '修改
      sql="select * from IdeaNet_FriendSite where ID="&ID
      rs.open sql,conn,1,3
      rs("TitleNameLangI")=trim(Request.Form("TitleNameLangI"))
      rs("TitleNameLangD")=trim(Request.Form("TitleNameLangD"))
      rs("TitleNameLangE")=trim(Request.Form("TitleNameLangE"))
	  if Request.Form("ViewFlagLangI")=1 then
        rs("ViewFlagLangI")=Request.Form("ViewFlagLangI")
	  else
        rs("ViewFlagLangI")=0
	  end if
	  if Request.Form("ViewFlagLangD")=1 then
        rs("ViewFlagLangD")=Request.Form("ViewFlagLangD")
	  else
        rs("ViewFlagLangD")=0
	  end if
	  if Request.Form("ViewFlagLangE")=1 then
        rs("ViewFlagLangE")=Request.Form("ViewFlagLangE")
	  else
        rs("ViewFlagLangE")=0
	  end if
	  rs("Content")=trim(Request.Form("Content")) 
	  rs("GotoUrl")=Request.Form("GotoUrl")
	  rs("ConciseLangI")=Request.Form("ConciseLangI")
	  rs("ConciseLangD")=Request.Form("ConciseLangD")
	  rs("ConciseLangE")=Request.Form("ConciseLangE")
	  rs("OrderNum")=99
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑友情链接！');location.replace('FriendSiteList.asp');</script>"
	
  else '提取友情链接
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from IdeaNet_FriendSite where ID="& ID
      rs.open sql,conn,1,1
	  TitleNameLangI=rs("TitleNameLangI")
	  TitleNameLangD=rs("TitleNameLangD")
	  TitleNameLangE=rs("TitleNameLangE")
	  ViewFlagLangI=rs("ViewFlagLangI")
	  ViewFlagLangD=rs("ViewFlagLangD")
	  ViewFlagLangE=rs("ViewFlagLangE")
	  Content=rs("Content")
	  GotoUrl=rs("GotoUrl")
	  ConciseLangI=rs("ConciseLangI")
      ConciseLangD=rs("ConciseLangD")
      ConciseLangE=rs("ConciseLangE")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
%>
