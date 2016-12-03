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
<META NAME="Author" CONTENT="艾迪创意,www.ideanet.cn" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>留言列表</TITLE>
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="Js/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|91,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'判断是否具有管理权限
%>
<BODY>
<table id="bodycontent" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">

<table cellpadding="0" cellspacing="0">
  <tr>
    <th><strong>留言信息：添加，修改、审核留言相关的内容</strong></th>
  </tr>
  <tr>
    <td height="24" align="center">【留言信息管理】</td>    
  </tr>
</table>
<br>
<table border="0" cellpadding="0" cellspacing="0">
  <form action="DelContent.asp?Result=Message" method="post" name="formDel" >
    <tr>
      <th width="30">ID</th>
      <th>留言标题</th>
	  <th width="60">审核权限</th>
      <th width="60">显示顺序</th>
      <th width="125">发布时间</th>
      <th colspan="2" width="140">操作
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="全" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="反" style="HEIGHT: 18px;WIDTH: 16px;">      </th>
    </tr>
	<% MessageList() %>
  </form>
</table>
<% if request.QueryString("Result")="ModifyOrderNum" then call ModifyOrderNum() %>
<% if request.QueryString("Result")="SaveOrderNum" then call SaveOrderNum() %>

</td></tr></table>
</body>
</html>
<%
'-----------------------------------------------------------
function MessageList()
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数
  dim page'页码
      page=clng(request("Page"))
  dim pagenc '每页显示的分页页码数量=pagenc*2+1
      pagenc=2
  dim pagenmax '每页显示的分页的最大页码
  dim pagenmin '每页显示的分页的最小页码
  dim datafrom'数据表名
      datafrom="IdeaNet_Message"
  dim datawhere'数据条件
      datawhere=""
  dim sqlid'本页需要用到的id
  dim Myself,PATH_INFO,QUERY_STRING'本页地址和参数
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" or Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
	    Myself = PATH_INFO & "?"
	  else
	    Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
	  end if
  dim taxis'排序的语句
      taxis="order by OrderNum,Id asc"
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(ID) as idCount from ["& datafrom &"]" & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,0,1
  idCount=rs("idCount")
  '获取记录总数

  if(idcount>0) then'如果记录总数=0,则不处理
    if(idcount mod pages=0)then'如果记录总数除以每页条数有余数,则=记录总数/每页条数+1
	  pagec=int(idcount/pages)'获取总页数
   	else
      pagec=int(idcount/pages)+1'获取总页数
    end if
	'获取本页需要用到的id============================================
    '读取所有记录的id数值,因为只有id所以速度很快
    sql="select id from ["& datafrom &"] " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    rs.pagesize = pages '每页显示记录数
    if page < 1 then page = 1
    if page > pagec then page = pagec
    if pagec > 0 then rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("id")
	  else
	    sqlid=sqlid &","&rs("id")
	  end if
	  rs.movenext
    next
  '获取本页需要用到的id结束============================================
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  if(idcount>0 and sqlid<>"") then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    sql="select * from ["& datafrom &"] where id in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,0,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#ffffff'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>&nbsp;"&rs("ID")&"</td>" & vbCrLf
      Response.Write "<td nowrap>&nbsp;"&rs("TitleName")& vbCrLf  
	  Response.Write "<td nowrap>" & vbCrLf
	  if Not rs("ViewFlag") then Response.Write "<font color='red'>未审核</font>"
	  if rs("ViewFlag") then Response.Write "<font color='green'>审核通过</font>"
	  Response.Write "&nbsp;</td>"
      Response.Write "<td nowrap><font color='blue'>"&rs("OrderNum")&"</font></td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("AddTime")&"</td>" & vbCrLf
      Response.Write "<td width=140 nowrap><a href='MessageEdit.asp?Result=Add' onClick='changeAdminFlag(""添加留言信息"")'><font color='#330099'>添加</font></a>&nbsp;&nbsp;<a href='MessageEdit.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""修改留言信息"")'><font color='#330099'>修改、审核</font></a>&nbsp;&nbsp;<a href='MessageList.asp?Result=ModifyOrderNum&ID="&rs("ID")&"' onClick='changeAdminFlag(""排序留言信息"")'><font color='#330099'>排序</font></a></td>" & vbCrLf
 	  Response.Write "<td width='20' nowrap><input name='selectID' type='checkbox' value='"&rs("ID")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='5' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td colspan='2' nowrap  bgcolor='#EBF2F9'><input name='submitDelSelect' type='button' class='button'  id='submitDelSelect' value='删除所选' onClick='ConfirmDel(""您真的要删除这些留言信息吗？"");'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td height='50' align='center' colspan='12' nowrap  bgcolor='#EBF2F9'>暂无留言信息</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='12'>" & vbCrLf
  Response.Write "<table align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td style='border:0px;'>共计：<font color='#ff6600'>"&idcount&"</font>条记录&nbsp;页次：<font color='#ff6600'>"&page&"</font></b>/"&pagec&"&nbsp;每页：<font color='#ff6600'>"&pages&"</font>条</td>" & vbCrLf
  Response.Write "<td align='right' style='border:0px;'>" & vbCrLf
  '设置分页页码开始===============================
  pagenmin=page-pagenc '计算页码开始值
  pagenmax=page+pagenc '计算页码结束值
  if(pagenmin<1) then pagenmin=1 '如果页码开始值小于1则=1
  if(page>1) then response.write ("<a href='"& myself &"Page=1'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>9</font></a>&nbsp;") '如果页码大于1则显示(第一页)
  if(pagenmin>1) then response.write ("<a href='"& myself &"Page="& page-(pagenc*2+1) &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>7</font></a>&nbsp;") '如果页码开始值大于1则显示(更前)
  if(pagenmax>pagec) then pagenmax=pagec '如果页码结束值大于总页数,则=总页数
  for i = pagenmin to pagenmax'循环输出页码
	if(i=page) then
	  response.write ("&nbsp;<font color='#ff6600'>"& i &"</font>&nbsp;")
	else
	  response.write ("[<a href="& myself &"Page="& i &">"& i &"</a>]")
	end if
  next
  if(pagenmax<pagec) then response.write ("&nbsp;<a href='"& myself &"Page="& page+(pagenc*2+1) &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>8</font></a>&nbsp;") '如果页码结束值小于总页数则显示(更后)
  if(page<pagec) then response.write ("<a href='"& myself &"Page="& pagec &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>:</font></a>&nbsp;") '如果页码小于总页数则显示(最后页)	
  '设置分页页码结束===============================
  Response.Write "跳到：第&nbsp;<input name='SkipPage' onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('只能在跳转目标页框内输入整数！');this.value='"&Page&"';}"" style='HEIGHT: 18px;WIDTH: 40px;'  type='text' class='textfield' value='"&Page&"'>&nbsp;页" & vbCrLf
  Response.Write "<input style='HEIGHT: 18px;WIDTH: 20px;' name='submitSkip' type='button' class='button' onClick='GoPage("""&Myself&""")' value='GO'>" & vbCrLf
  Response.Write "</td>" & vbCrLf
  Response.Write "</tr>" & vbCrLf
  Response.Write "</table>" & vbCrLf
  rs.close
  set rs=nothing
  Response.Write "</td>" & vbCrLf  
  Response.Write "</tr>" & vbCrLf
'-----------------------------------------------------------
'-----------------------------------------------------------
end function 
%>
<%
sub ModifyOrderNum()
  dim rs,sql,ID,TitleName,OrderNum
  ID=request.QueryString("ID")
  set rs = server.createobject("adodb.recordset")
  sql="select * from IdeaNet_Message where ID="& ID
  rs.open sql,conn,1,1
  TitleName=rs("TitleName")
  OrderNum=rs("OrderNum")
  rs.close
  set rs=nothing
  response.write "<br>"
  response.write "<table width='100%' border='0' cellpadding='1' cellspacing='1' bgcolor='#6298E1'>"
  response.write "<form action='MessageList.asp?Result=SaveOrderNum' method='post' name='formOrderNum'>"
  response.write "<tr>"
  response.write "<td height='24' align='center' nowrap  bgcolor='#EBF2F9'>ID：<input name='ID' type='text' class='textfield'  style='WIDTH: 30;' value='"&ID&"' maxlength='4' readonly>&nbsp;留言信息名称：<input name='TitleName' type='text' class='textfield' id='TitleName' style='WIDTH: 240;' value='"&TitleName&"' maxlength='36' >&nbsp;排序号：<input name='OrderNum' type='text' class='textfield' style='WIDTH: 60;' value='"&OrderNum&"' maxlength='4'  onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('只能在排序号内输入整数！');this.value='"&OrderNum&"';}"">*(数字越小越靠前&nbsp;&nbsp;<input name='submitOrderNum' type='submit' class='button' value='保存' style='WIDTH: 60;' ></td>"
  response.write "</tr>"
  response.write "</form>"
  response.write "</table>"
end sub

sub SaveOrderNum()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from IdeaNet_Message where ID="& request.form("ID")
  rs.open sql,conn,1,3
  rs("TitleName")=request.form("TitleName")
  rs("OrderNum")=request.form("OrderNum")
  rs.update
  rs.close
  set rs=nothing
  response.redirect "MessageList.asp"
end sub
%>