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
<TITLE>信息中心列表</TITLE>
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="Js/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|62,")=0 then 
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
dim Result,StartDate,EndDate,Keyword,SortID,SortPath
Result=request.QueryString("Result")
StartDate=request.QueryString("StartDate")
EndDate=request.QueryString("EndDate")
Keyword=request.QueryString("Keyword")
SortID=request.QueryString("SortID")
SortPath=request.QueryString("SortPath")
function PlaceFlag()
  if Result="Search" then
    Response.Write "信息中心：列表&nbsp;->&nbsp;检索&nbsp;->&nbsp;添加时间[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，关键字[<font color='red'>"&Keyword&"</font>]"
  else
    if SortPath<>"" then
      Response.Write "信息中心：列表&nbsp;->&nbsp<a href='InfoList.asp'>全部</a>"
	  TextPath(SortID)
	else
      Response.Write "信息中心：列表&nbsp;->&nbsp全部"
	end if
  end if
end function  
%>

<table cellpadding="0" cellspacing="0">
  <tr>
    <th>信息中心检索及分类查看：添加，修改，删除信息中心</th>
  </tr>
  <tr>
    <td height="36" align="center"><table class="noborder" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="Search.asp?Result=Info">
          <td nowrap>从
          <script language=javascript> 
          var myDate=new dateSelector(); 
          myDate.year--; 
		  myDate.date; 
          myDate.inputName='start_date';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。 
          myDate.display(); 
          </script>
          &nbsp;到
          <script language=javascript> 
          myDate.year++; 
          myDate.inputName='end_date';  //注意这里设置输入框的name，同一页中的日期输入框，不能出现重复的name。 
          myDate.display(); 
          </script>
          &nbsp;&nbsp;关键字：<input name="Keyword" type="text" class="textfield" value="<%=Keyword%>" size="18">
          <input name="submitSearch" type="submit" class="button" value="检索">
          </td>
        </form>
        <td align="right" nowrap><a href="InfoList.asp" onClick='changeAdminFlag("信息中心列表")'>全部信息中心</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="InfoSort.asp" onClick='changeAdminFlag("选择查看分类")'>其他类别信息</a></td>
      </tr>
    </table></td>    
  </tr>
</table>
<table cellspacing="0" cellpadding="0">
  <tr>
    <td height="30"><%PlaceFlag()%></td>
  </tr>
</table>

<table cellpadding="0" cellspacing="0">
  <form action="DelContent.asp?Result=Info" method="post" name="formDel" >
  <tr>
    <th width="30">ID</th>
	<th width="100">所属类别</th>
    <th>信息中心标题</th>
	<th width="100">属性、权限</th>
    <th width="60">显示顺序</th>
    <th width="125">发布时间</th>
    <th width="60">点击量</th>
    <th colspan="2" width="120">操作
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="全" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="反" style="HEIGHT: 18px;WIDTH: 16px;">	</th>
  </tr>
  <% OthersList() %>
  </form>
</table>
<% if request.QueryString("Result")="ModifyOrderNum" then call ModifyOrderNum() %>
<% if request.QueryString("Result")="SaveOrderNum" then call SaveOrderNum() %>

</td></tr></table>
</BODY>
</HTML>
<%
'-----------------------------------------------------------
function OthersList()
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数
  dim page'页码
      page=clng(request("Page"))
  dim pagenc'每页显示的分页页码数量=pagenc*2+1
      pagenc=2
  dim pagenmax'每页显示的分页的最大页码
  dim pagenmin'每页显示的分页的最小页码
  dim datafrom'数据表名
      datafrom="IdeaNet_Info"
  dim datawhere'数据条件
      if Result="Search" then
	     datawhere="where ( TitleNameLangI like '%" & Keyword &_
		           "%' or SourceLangI like '%" & Keyword &_
		           "%') and AddTime >= #" & StartDate & " # and AddTime <= #" & EndDate & "#"
	  else
	    if SortPath<>"" then'是否查看的分类信息
		  datawhere="where Instr(SortPath,'"&SortPath&"')>0 "
        else
		  datawhere=""
		end if
	  end if
  dim sqlid'本页需要用到的id
  dim Myself,PATH_INFO,QUERY_STRING'本页地址和参数
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" or Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
	    Myself = PATH_INFO & "?"
	  else
	    Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
	  end if
  dim taxis'排序的语句 asc,desc
      taxis="order by OrderNum,Id desc"
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
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ID")&"</td>" & vbCrLf
	  Response.Write "<td nowrap><a href='InfoList.asp?SortID="&rs("ID")&"&SortPath="&rs("SortPath")&"' onClick='changeAdminFlag(""信息中心列表"")'>"&SortText(rs("SortID"))&"</a></td>" & vbCrLf
	  if StrLen((rs("TitleNameLangI")))>30 then
        Response.Write "<td nowrap><a href='InfoEdit.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""信息中心列表"")'>"&StrLeft(rs("TitleNameLangI"),30)&"</a></td>" & vbCrLf
      else
        Response.Write "<td nowrap><a href='InfoEdit.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""信息中心列表"")'>"&rs("TitleNameLangI")&"</a></td>" & vbCrLf
      end if 
	  Response.Write "<td nowrap>" & vbCrLf
	  if rs("NewFlag") then Response.Write "<font color='red'>最新</font>"
	  if rs("TjFlag") then Response.Write "<font color='green'>推荐</font>"
	  if rs("HotFlag") then Response.Write "<font color='yellow'>热门</font>"
	  Response.Write "&nbsp;</td>"
      Response.Write "<td nowrap>"&rs("OrderNum")&"&nbsp;</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("AddTime")&"&nbsp;</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Hits")&"&nbsp;</td>" & vbCrLf
      Response.Write "<td width=100 nowrap><a href='InfoEdit.asp?Result=Add' onClick='changeAdminFlag(""添加信息中心"")'><font color='#330099'>添加</font></a>&nbsp;&nbsp;<a href='InfoEdit.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""修改信息中心"")'><font color='#330099'>修改</font></a>&nbsp;&nbsp;<a href='InfoList.asp?Result=ModifyOrderNum&ID="&rs("ID")&"' onClick='changeAdminFlag(""排序信息中心"")'><font color='#330099'>排序</font></a></td>" & vbCrLf
      Response.Write "<td width='20' nowrap><input name='selectID' type='checkbox' value='"&rs("ID")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='7' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td colspan='2' nowrap  bgcolor='#EBF2F9'><input name='submitDelSelect' type='button' class='button'  id='submitDelSelect' value='删除所选' onClick='ConfirmDel(""您真的要删除这些信息中心吗？"");'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td height='50' align='center' colspan='11' nowrap  bgcolor='#EBF2F9'>暂无信息中心</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='12' nowrap  bgcolor='#D7E4F7'>" & vbCrLf
  Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td style='border:0px;'>共计：<font color='#ff6600'>"&idcount&"</font>条记录&nbsp;页次：<font color='#ff6600'>"&page&"</font></strong>/"&pagec&"&nbsp;每页：<font color='#ff6600'>"&pages&"</font>条</td>" & vbCrLf
  Response.Write "<td style='border:0px;' align='right'>" & vbCrLf
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
 if instr(myself,"sortpath")>0 THen
   Response.Write "<input style='HEIGHT: 18px;WIDTH: 20px;' name='submitSkip' type='button' class='button' onClick='GoPage("""&Myself&""")' value='GO'>" &  vbCrLf
  Response.Write "</td>" & vbCrLf
  else
  
  Response.Write "<input style='HEIGHT: 18px;WIDTH: 20px;' name='submitSkip' type='button' class='button' onClick='GoPage("""&Myself&"sortpath="&sortpath&"&sortid="&SortID&"&"")' value='GO'>" &  vbCrLf
  Response.Write "</td>" & vbCrLf
   end if
  Response.Write "</tr>" & vbCrLf
  Response.Write "</table>" & vbCrLf
  rs.close
  set rs=nothing
  Response.Write "</td>" & vbCrLf  
  Response.Write "</tr>" & vbCrLf
'-----------------------------------------------------------
'-----------------------------------------------------------
end function 
'生成节点文字路径--------------------------
Function TextPath(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From IdeaNet_InfoSort where ID="&ID
  rs.open sql,conn,1,1
  TextPath="&nbsp;->&nbsp;<a href=InfoList.asp?SortID="&rs("ID")&"&SortPath="&rs("SortPath")&">"&rs("SortNameLangI")&"</a>"
  if rs("ParentID")<>0 then TextPath rs("ParentID")
  response.write(TextPath)
End Function
%>
<%
'生成所属类别--------------------------
Function SortText(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From IdeaNet_InfoSort where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortNameLangI")
  rs.close
  set rs=nothing
End Function
%>


<%
sub ModifyOrderNum()
  dim rs,sql,ID,TitleNameLangI,OrderNum,Hits
  ID=request.QueryString("ID")
  set rs = server.createobject("adodb.recordset")
  sql="select * from IdeaNet_Info where ID="& ID
  rs.open sql,conn,1,1
  TitleNameLangI=rs("TitleNameLangI")
  OrderNum=rs("OrderNum")
  Hits=rs("Hits")
  rs.close
  set rs=nothing
  response.write "<br>"
  response.write "<table width='100%' border='0' cellpadding='1' cellspacing='1' bgcolor='#6298E1'>"
  response.write "<form action='InfoList.asp?Result=SaveOrderNum' method='post' name='formOrderNum'>"
  response.write "<tr>"
  response.write "<td height='24' align='center' nowrap  bgcolor='#EBF2F9'>ID：<input name='ID' type='text' class='textfield'  style='WIDTH: 30;' value='"&ID&"' maxlength='4' readonly>&nbsp;信息中心名称：<input name='TitleNameLangI' type='text' class='textfield' id='TitleNameLangI' style='WIDTH: 240;' value='"&TitleNameLangI&"' maxlength='36' >&nbsp;排序号：<input name='OrderNum' type='text' class='textfield' style='WIDTH: 60;' value='"&OrderNum&"' maxlength='4'  onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('只能在排序号内输入整数！');this.value='"&OrderNum&"';}"">*(数字越小越靠前)&nbsp;&nbsp;点击量：<input name='Hits' type='text' class='textfield' style='WIDTH: 60;' value='"&Hits&"' maxlength='8'  onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('只能在排序号内输入整数！');this.value='"&Hits&"';}"">&nbsp;&nbsp;<input name='submitOrderNum' type='submit' class='button' value='保存' style='WIDTH: 60;' ></td>"
  response.write "</tr>"
  response.write "</form>"
  response.write "</table>"
end sub

sub SaveOrderNum()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from IdeaNet_Info where ID="& request.form("ID")
  rs.open sql,conn,1,3
  rs("TitleNameLangI")=request.form("TitleNameLangI")
  rs("OrderNum")=request.form("OrderNum")
  rs("Hits")=request.form("Hits")
  rs.update
  rs.close
  set rs=nothing
  response.redirect "InfoList.asp"
end sub
%>