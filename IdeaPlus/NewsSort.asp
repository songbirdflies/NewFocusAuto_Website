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
<TITLE>新闻分类</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2008-2018 - www.ideanet.cn" />
<META NAME="Author" CONTENT="艾迪创意,www.ideanet.cn" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="Js/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|31,")=0 then 
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
Dim OrderNum,Action
Action=request.QueryString("Action")
Select Case Action
  Case "Add"
	addFolder
  	CallFolderView()
  Case "Del"
    Dim rs,sql,SortPath
    Set rs=server.CreateObject("adodb.recordset")
    sql="Select * From IdeaNet_NewsSort where ID="&request.QueryString("id")
    rs.open sql,conn,1,1
	SortPath=rs("SortPath")
	conn.execute("delete from  IdeaNet_NewsSort where Instr(SortPath,'"&SortPath&"')>0")
    conn.execute("delete from  IdeaNet_News where Instr(SortPath,'"&SortPath&"')>0")
    response.write ("<script language=javascript> alert('成功删除本类、子类及所有下属信息条目，点击确定查看类别树！');location.replace('NewsSort.asp');</script>")
  Case "Save"
	saveFolder ()
  Case "Edit"
	editFolder
  	CallFolderView()	
  Case "Move"
	moveFolderForm ()
  	CallFolderView()
  Case "MoveSave"
	saveMoveFolder ()
  Case Else
	CallFolderView()
End Select
%>

</td></tr></table>
</BODY>
</HTML>
<%
'调用显示节点------------------------------
Function CallFolderView()
%>
<table cellpadding="0" cellspacing="0">
  <tr>
    <th>类别树查看管理：</th>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="NewsSort.asp?Action=Add&ParentID=0">添加一级分类</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NewsList.asp" onClick='changeAdminFlag("新闻列表")'>查看所有新闻</a></td>
  </tr>
  <tr>
    <td height="24" nowrap  bgcolor="#EBF2F9"><% Folder(0) %></td>
  </tr>
</table>
<%
End Function
'列出所有节点------------------------------
Function Folder(id)
  Dim rs,sql,i,ChildCount,FolderType,FolderName,onMouseUp,ListType,OrderNum
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From IdeaNet_NewsSort where ParentID="&id&" order by id"
  rs.open sql,conn,1,1
  if id=0 and rs.recordcount=0 then
    response.write ("暂无分类!</td></tr></table>")
    response.end
  end if  
  i=1
  response.write("<table class='noborder' cellspacing='0' cellpadding='0'>")
  while not rs.eof
    ChildCount=conn.execute("select count(*) from IdeaNet_NewsSort where ParentID="&rs("id"))(0)
    if ChildCount=0 then
	  if i=rs.recordcount then
	    FolderType="SortFileEnd"
	  else
	    FolderType="SortFile"
	  end if
	  FolderName=rs("SortNameLangI")
	  onMouseUp=""
    else
	  if i=rs.recordcount then
	 	FolderType="SortEndFolderClose"
		ListType="SortEndListline"
		onMouseUp="EndSortChange('a"&rs("id")&"','b"&rs("id")&"');"
	  else
		FolderType="SortFolderClose"
		ListType="SortListline"
		onMouseUp="SortChange('a"&rs("id")&"','b"&rs("id")&"');"
	  end if
	  FolderName=rs("SortNameLangI")
    end if
    response.write("<tr>")
    response.write("<td nowrap id='b"&rs("id")&"' class='"&FolderType&"' onMouseUp="&onMouseUp&"></td><td nowrap>"&FolderName&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")	
    response.write("<font color='#FF0000'>分类：</font><a href='NewsSort.asp?Action=Add&ParentID="&rs("id")&"'>添加</a>")
    response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href='NewsSort.asp?Action=Edit&ID="&rs("id")&"'>修改</a>")
    response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href='NewsSort.asp?Action=Move&ID="&rs("id")&"&ParentID="&rs("Parentid")&"&SortNameLangI="&rs("SortNameLangI")&"&SortPath="&rs("SortPath")&"'>移</a>")
    response.write("→<a href='#' onclick='SortFromTo.rows[4].cells[0].innerHTML=""→&nbsp;"&rs("SortNameLangI")&""";MoveForm.toID.value="&rs("ID")&";MoveForm.toParentID.value="&rs("ParentID")&";MoveForm.toSortPath.value="""&rs("SortPath")&""";'>至</a>")
	response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href=javascript:ConfirmDelSort('NewsSort',"&rs("id")&")>删除</a>")
    response.write("&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>新闻：</font><a href='NewsEdit.asp?Result=Add' onClick='changeAdminFlag(""添加新闻"")'>添加</a>")
    response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href='NewsList.asp?SortID="&rs("ID")&"&SortPath="&rs("SortPath")&"' onClick='changeAdminFlag(""新闻列表"")'>列表</a>&nbsp;|&nbsp;排序&nbsp;"&rs("OrderNum")&"")
    response.write("</td></tr>")
    if ChildCount>0 then
%>
      <tr id="a<%= rs("id")%>" style="display:yes"><td class="<%= ListType%>" nowrap></td><td ><% Folder(rs("id")) %></td></tr>
<%
    end if
    rs.movenext
    i=i+1
  wend
  response.write("</table>")
  rs.close
  set rs=nothing
end function
'添加节点---------------------------------
Function addFolder()
  Dim ParentID
  ParentID=request.QueryString("ParentID")
  addFolderForm ParentID
end function
'添加节点表单------------------------------
Function addFolderForm(ParentID)
  Dim ParentPath,SortTextPath,rs,sql
  if ParentID=0 then
    ParentPath="0,"
	SortTextPath=""
  else 
    Set rs=server.CreateObject("adodb.recordset")
    sql="Select * From IdeaNet_NewsSort where ID="&ParentID
    rs.open sql,conn,1,1
	ParentPath=rs("SortPath")
  end if
%>
<table cellpadding="0" cellspacing="0">
<form name="editForm" method="post" action="NewsSort.asp?Action=Save&From=Add">
  <tr>
    <th>添加类别：通过"显示"可控制每种分类是否在相应语言版网站里显示出来。</th>
  </tr>
  <tr>
    <td height="24">|&nbsp;根类&nbsp;→&nbsp;<% if ParentID<>0 then TextPath(ParentID)%></td>
  </tr>
  <tr>
    <td height="24">
	<table class="noborder" width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
		<td width="330" colspan="4"></td>
        <td width="120">父类ID：<input readonly name="ParentID" type="text" class="textfield" id="ParentID" size="6" value="<%=ParentID %>"></td>
        <td>父类数字路径：<input readonly name="ParentPath" type="text" class="textfield" id="ParentPath" size="8" value="<%=ParentPath%>">&nbsp;&nbsp;排序 <input   name="OrderNum" type="text" class="textfield" id="OrderNum" size="4" value="<%=OrderNum%>"></td>		
		</tr>
      <tr class="LangShowI">
        <td width="100" align="right"><%=LangWebI%>名称：</td>
        <td width="140"><input name="SortNameLangI" type="text" class="textfield" id="SortNameLangI" size="20"></td>
        <td width="40">显示：</td>
        <td width="90">
		<input name="ViewFlagLangI" type="radio" value="1" checked="checked" />是
		<input name="ViewFlagLangI" type="radio" value="0" />否</td>
		<td colspan="2"></td>
	  </tr>
	  <tr class="LangShowD">
        <td align="right"><%=LangWebD%>名称：</td>
        <td><input name="SortNameLangD" type="text" class="textfield" id="SortNameLangD" size="20"></td>
        <td>显示：</td>
        <td>
		<input name="ViewFlagLangD" type="radio" value="1" />是
		<input name="ViewFlagLangD" type="radio" value="0" checked="checked" />否</td>
		<td colspan="2"></td>
	  </tr>
      <tr class="LangShowE">
        <td align="right"><%=LangWebE%>名称：</td>
        <td><input name="SortNameLangE" type="text" class="textfield" id="SortNameLangE" size="20"></td>
        <td>显示：</td>
        <td>
		<input name="ViewFlagLangE" type="radio" value="1" />是
		<input name="ViewFlagLangE" type="radio" value="0" checked="checked" />否</td>
		<td colspan="2"></td>
	  </tr>
      <tr class="LangShowI">
		<td align="right"><%=LangWebI%>静态文件名：</td>
		<td colspan="5"><input name="HtmlUrlNameLangI" type="text" class="textfield" id="HtmlUrlNameLangI" size="20">&nbsp;不填则由系统自动按<%=LangWebI%>名称生成(建议)，不可有重复，一旦填好不可更改</td>
	  </tr>
	  <tr class="LangShowD">
		<td align="right"><%=LangWebD%>静态文件名：</td>
		<td colspan="5"><input name="HtmlUrlNameLangD" type="text" class="textfield" id="HtmlUrlNameLangD" size="20">&nbsp;不填则由系统自动按<%=LangWebD%>名称生成(建议)，不可有重复，一旦填好不可更改</td>
	  </tr>
	  <tr class="LangShowE">
		<td align="right"><%=LangWebE%>静态文件名：</td>
		<td colspan="5"><input name="HtmlUrlNameLangE" type="text" class="textfield" id="HtmlUrlNameLangE" size="20">&nbsp;不填则由系统自动按<%=LangWebE%>名称生成(建议)，不可有重复，一旦填好不可更改</td>
	  </tr>	
      <tr>
        <td align="right">显示方式：</td>
        <td colspan="5">
		<select name="SitListshow" id="SitListshow" style="width:270;">
          <option value="Title">标题列表</option>
          <option value="Pic">图片列表</option>
          <option value="Concise">简要概述列表</option>
          <option value="TitleTime">标题+时间列表</option>
          <option value="TitleConcise">标题+简要概述列表</option>
          <option value="TitleConciseTime">标题+简要概述+时间列表</option>
          <option value="TimeTitleConcise">时间+标题+简要概述列表</option>
          <option value="PicTitle">纵向图片+标题列表</option>
          <option value="PicConcise">纵向图片+简要概述列表</option>
          <option value="PicTitleConcise">纵向图片+标题+简要概述列表</option>
          <option value="TitlePicConcise">纵向标题+图片+简要概述列表</option>
          <option value="LeftPicConcise">横向图片+简要概述列表</option>
          <option value="LeftPicTitleConcise" selected="selected">横向图片+标题+简要概述列表</option>
          <option value="LeftTitlePicConcise">横向标题+图片+简要概述列表</option>
          <option value="LeftPicTitleConciseTime">横向图片+标题+简要概述+时间列表</option>
        </select>
		</td>
      </tr>
      <tr>
		<td align="right">每行显示条数：</td>
        <td colspan="5"><input name="TdNumstr" type="text" class="TdNumstr" value="1" size="20" maxlength="5"></td>
	  </tr>	
	  <tr>
        <td align="right">分类小图：</td>
        <td colspan="5"><input name="SmallPic" type="text" class="textfield" style="width:270;" maxlength="100">&nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=SmallPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
      <tr>
        <td align="right">分类大图：</td>
        <td colspan="5"><input name="BigPic" type="text" class="textfield" style="width:270;" maxlength="100">&nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=BigPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
	  <tr>
        <td align="right">链接网址：</td>
        <td colspan="5"><input name="GotoUrl" type="text" class="textfield" style="width:270;" maxlength="100"></td>
      </tr>
	  <tr class="LangShowI">
        <td align="right"><%=LangWebI%>简要概述：</td>
        <td colspan="5"><textarea id="ConciseLangI" name="ConciseLangI" rows="1" cols="60"></textarea>&nbsp;</td>
      </tr>
	  <tr class="LangShowD">
        <td align="right"><%=LangWebD%>简要概述：</td>
        <td colspan="5"><textarea id="ConciseLangD" name="ConciseLangD" rows="1" cols="60"></textarea>&nbsp;</td>
      </tr>
	  <tr class="LangShowE">
        <td align="right"><%=LangWebE%>简要概述：</td>
        <td colspan="5"><textarea id="ConciseLangE" name="ConciseLangE" rows="1" cols="60"></textarea>&nbsp;</td>
      </tr>  	  
	  <tr>
	    <td align="right">&nbsp;</td>
	    <td colspan="5"><input name="submitSave" type="submit" class="button" id="保存" value=" 保存 "></td>
	  </tr>
    </table>
	</td>
  </tr>
</form>
</table>
<br>
<%
End Function
'生成节点文字路径--------------------------
Function TextPath(ID)
  Dim rs,sql,SortTextPath
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From IdeaNet_NewsSort where ID="&ID
  rs.open sql,conn,1,1
  SortTextPath=rs("SortNameLangI")&"&nbsp;→&nbsp;"
  if rs("ParentID")<>0 then TextPath rs("ParentID")
  response.write(SortTextPath)
End Function
'保存添加、修改节点-------------------------
Function saveFolder
  if request.Form("ViewFlagLangI")=1 then
    if len(trim(request.Form("SortNameLangI")))=0 then
	  response.write ("<script language=javascript> alert('你已经选择了"""&LangWebI&"""显示，因此"&LangWebI&"不名称必填！');history.back(-1);</script>")
	  response.end
	  end if
  end if
  if request.Form("ViewFlagLangD")=1 then
    if len(trim(request.Form("SortNameLangD")))=0 then
      response.write ("<script language=javascript> alert('您已经选择了"""&LangWebD&"""显示，因此"&LangWebD&"类别名必填！');history.back(-1);</script>")
      response.end
    end if
  end if
  if request.Form("ViewFlagLangE")=1 then
    if len(trim(request.Form("SortNameLangE")))=0 then
      response.write ("<script language=javascript> alert('您已经选择了"""&LangWebE&"""显示，因此"&LangWebE&"类别名必填！');history.back(-1);</script>")
      response.end
    end if
  end if
  Dim From,Action,rs,sql,SortTextPath,OrderNum
  From=request.QueryString("From")
  Set rs=server.CreateObject("adodb.recordset")
  if From="Add" then 
    sql="Select * From IdeaNet_NewsSort"
    rs.open sql,conn,1,3
    rs.addnew
	Action="添加类别"
    rs("SortPath")=request.Form("ParentPath") & rs("ID") &","
  else
    sql="Select * From IdeaNet_NewsSort where ID="&request.QueryString("ID")
    rs.open sql,conn,1,3
	Action="修改类别"
    rs("SortPath")=request.Form("SortPath")
  end if
  rs("SortNameLangI")=request.Form("SortNameLangI")
  rs("ViewFlagLangI")=request.Form("ViewFlagLangI")
  rs("SortNameLangD")=request.Form("SortNameLangD")
  rs("ViewFlagLangD")=request.Form("ViewFlagLangD")
  rs("SortNameLangE")=request.Form("SortNameLangE")
  rs("ViewFlagLangE")=request.Form("ViewFlagLangE")
  rs("ConciseLangI")=request.Form("ConciseLangI")
  rs("ConciseLangD")=request.Form("ConciseLangD")
  rs("ConciseLangE")=request.Form("ConciseLangE")
  rs("HtmlUrlNameLangI")=trim(Request.Form("HtmlUrlNameLangI"))
  rs("HtmlUrlNameLangD")=trim(Request.Form("HtmlUrlNameLangD"))
  rs("HtmlUrlNameLangE")=trim(Request.Form("HtmlUrlNameLangE"))
  rs("SmallPic")=request.Form("SmallPic")
  rs("BigPic")=request.Form("BigPic")
  rs("GotoUrl")=request.Form("GotoUrl")
  rs("SitListshow")=request.Form("SitListshow")
  rs("TdNumstr")=request.Form("TdNumstr")
  rs("ParentID")=request.Form("ParentID")
  if cstr(request.Form("OrderNum"))="" then
  rs("OrderNum")=0
  else
  rs("OrderNum")=request.Form("OrderNum")
  end if
  rs.update 
  response.write ("<script language=javascript> alert('"&Action&"保存成功，点击确定查看类别树！');location.replace('NewsSort.asp');</script>")
End Function 
'修改节点---------------------------------
Function editFolder()
  Dim ID
  ID=request.QueryString("ID")
  editFolderForm ID
end function
'修改节点表单------------------------------
Function editFolderForm(ID)
  Dim SortNameLangI,ViewFlagLangI,SortNameLangE,ViewFlagLangE,SortNameLangD,ViewFlagLangD,ParentID,SortPath,rs,sql,OrderNum,HtmlUrlNameLangI,HtmlUrlNameLangD,HtmlUrlNameLangE,SmallPic,BigPic,ConciseLangI,ConciseLangD,ConciseLangE,GotoUrl,SitListshow,TdNumstr
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From IdeaNet_NewsSort where ID="&ID
  rs.open sql,conn,1,1
  SortNameLangI=rs("SortNameLangI")
  ViewFlagLangI=rs("ViewFlagLangI")
  SortNameLangD=rs("SortNameLangD")
  ViewFlagLangD=rs("ViewFlagLangD")  
  SortNameLangE=rs("SortNameLangE")
  ViewFlagLangE=rs("ViewFlagLangE")
  ConciseLangI=rs("ConciseLangI")
  ConciseLangD=rs("ConciseLangD")
  ConciseLangE=rs("ConciseLangE")
  HtmlUrlNameLangI=rs("HtmlUrlNameLangI")
  HtmlUrlNameLangD=rs("HtmlUrlNameLangD")
  HtmlUrlNameLangE=rs("HtmlUrlNameLangE")
  SmallPic=rs("SmallPic")
  BigPic=rs("BigPic")
  GotoUrl=rs("GotoUrl")
  SitListshow=rs("SitListshow")
  TdNumstr=rs("TdNumstr")
  ParentID=rs("ParentID")
  SortPath=rs("SortPath")
  OrderNum=rs("OrderNum")
%>
<table cellpadding="0" cellspacing="0">
<form name="editForm" method="post" action="NewsSort.asp?Action=Save&From=Edit&ID=<%=ID%>">
  <tr>
    <th>修改类别：通过"显示"可控制每种分类是否在相应语言版网站里显示出来。</th>
  </tr>
  <tr>
    <td height="24">|&nbsp;根类&nbsp;→&nbsp;<% if ParentID<>0 then TextPath(ParentID)%></td>
  </tr>
  <tr>
    <td>
	<table class="noborder" cellpadding="0" cellspacing="0">
		<tr>
		<td colspan="4"></td>
        <td width="120">父类ID：<input readonly name="ParentID" type="text" class="textfield" id="ParentID" size="6" value="<%=ParentID %>"></td>
        <td>本类数字路径：<input readonly name="SortPath" type="text" class="textfield" id="SortPath" size="8" value="<%=SortPath%>"> 排序 <input   name="OrderNum" type="text" class="textfield" id="OrderNum" size="4" value="<%=OrderNum%>"></td>
      <tr class="LangShowI">
        <td width="100" align="right"><%=LangWebI%>名称：</td>
        <td width="140"><input name="SortNameLangI" type="text" class="textfield" id="SortNameLangI" size="20" value="<%=SortNameLangI%>"></td>
        <td width="40">显示：</td>
        <td width="90">
		<input name="ViewFlagLangI" type="radio" value="1" <%if ViewFlagLangI then response.write ("checked=checked")%> />是
		<input name="ViewFlagLangI" type="radio" value="0" <%if not ViewFlagLangI then response.write ("checked=checked")%>/>否</td>
        <td colspan="2"></td>
	  </tr>
      <tr class="LangShowD">
        <td align="right"><%=LangWebD%>名称：</td>
        <td><input name="SortNameLangD" type="text" class="textfield" id="SortNameLangD" size="20" value="<%=SortNameLangD%>"></td>
        <td>显示：</td>
        <td>
		<input name="ViewFlagLangD" type="radio" value="1" <%if ViewFlagLangD then response.write ("checked=checked")%> />是
		<input name="ViewFlagLangD" type="radio" value="0" <%if not ViewFlagLangD then response.write ("checked=checked")%> />否</td>
        <td colspan="2"></td>
	  </tr>
      <tr class="LangShowE">
        <td align="right"><%=LangWebE%>名称：</td>
        <td><input name="SortNameLangE" type="text" class="textfield" id="SortNameLangE" size="20" value="<%=SortNameLangE%>"></td>
        <td>显示：</td>
        <td>
		<input name="ViewFlagLangE" type="radio" value="1" <%if ViewFlagLangE then response.write ("checked=checked")%> />是
		<input name="ViewFlagLangE" type="radio" value="0" <%if not ViewFlagLangE then response.write ("checked=checked")%> />否</td>
        <td colspan="2"></td>
	  </tr>
      <tr class="LangShowI">
		<td align="right"><%=LangWebI%>静态文件名：</td>
		<td colspan="5"><input name="HtmlUrlNameLangI" type="text" class="textfield" id="HtmlUrlNameLangI" size="20" value="<%=HtmlUrlNameLangI%>">&nbsp;不填则由系统自动按<%=LangWebI%>名称生成(建议)，不可有重复，一旦填好不可更改</td>
	  </tr>
	  <tr class="LangShowD">
		<td align="right"><%=LangWebD%>静态文件名：</td>
		<td colspan="5"><input name="HtmlUrlNameLangD" type="text" class="textfield" id="HtmlUrlNameLangD" size="20" value="<%=HtmlUrlNameLangD%>">&nbsp;不填则由系统自动按<%=LangWebD%>名称生成(建议)，不可有重复，一旦填好不可更改</td>
	  </tr>
	  <tr class="LangShowE">
		<td align="right"><%=LangWebE%>静态文件名：</td>
		<td colspan="5"><input name="HtmlUrlNameLangE" type="text" class="textfield" id="HtmlUrlNameLangE" size="20" value="<%=HtmlUrlNameLangE%>">&nbsp;不填则由系统自动按<%=LangWebE%>名称生成(建议)，不可有重复，一旦填好不可更改</td>
	  </tr>
      <tr>
        <td align="right">显示方式：</td>
        <td colspan="5">
        <select name="SitListshow" id="SitListshow" style="width:270;">
          <option value="Title" <%if SitListshow="Title" then response.Write("selected='selected'")%>>标题列表</option>
          <option value="Pic" <%if SitListshow="Pic" then response.Write("selected='selected'")%>>图片列表</option>
          <option value="Concise" <%if SitListshow="Concise" then response.Write("selected='selected'")%>>简要概述列表</option>
          <option value="TitleTime" <%if SitListshow="TitleTime" then response.Write("selected='selected'")%>>标题+时间列表</option>
          <option value="TitleConcise" <%if SitListshow="TitleConcise" then response.Write("selected='selected'")%>>标题+简要概述列表</option>
          <option value="TitleConciseTime" <%if SitListshow="TitleConciseTime" then response.Write("selected='selected'")%>>标题+简要概述+时间列表</option>
          <option value="TimeTitleConcise" <%if SitListshow="TimeTitleConcise" then response.Write("selected='selected'")%>>时间+标题+简要概述列表</option>
          <option value="PicTitle" <%if SitListshow="PicTitle" then response.Write("selected='selected'")%>>纵向图片+标题列表</option>
          <option value="PicConcise" <%if SitListshow="PicTitleConcise" then response.Write("selected='selected'")%>>纵向图片+简要概述列表</option>
          <option value="PicTitleConcise" <%if SitListshow="PicTitleConcise" then response.Write("selected='selected'")%>>纵向图片+标题+简要概述列表</option>
          <option value="TitlePicConcise" <%if SitListshow="TitlePicConcise" then response.Write("selected='selected'")%>>纵向标题+图片+简要概述列表</option>
          <option value="LeftPicConcise" <%if SitListshow="LeftPicConcise" then response.Write("selected='selected'")%>>横向图片+简要概述列表</option>
          <option value="LeftPicTitleConcise" <%if SitListshow="LeftPicTitleConcise" then response.Write("selected='selected'")%>>横向图片+标题+简要概述列表</option>
          <option value="LeftTitlePicConcise" <%if SitListshow="LeftTitlePicConcise" then response.Write("selected='selected'")%>>横向标题+图片+简要概述列表</option>
          <option value="LeftPicTitleConciseTime" <%if SitListshow="LeftPicTitleConciseTime" then response.Write("selected='selected'")%>>横向图片+标题+简要概述+时间列表</option>
        </select>
		</td>
      </tr>
      <tr>
		<td align="right">每行显示条数：</td>
        <td colspan="5"><input name="TdNumstr" type="text" class="TdNumstr" value="<%=TdNumstr%>" size="20" maxlength="2"></td>
	  </tr>	
	  <tr>
        <td align="right">分类小图：</td>
        <td colspan="5"><input name="SmallPic" type="text" class="textfield" style="width:270;" value="<%=SmallPic%>" maxlength="100">&nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=SmallPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
      <tr>
        <td align="right">分类大图：</td>
        <td colspan="5"><input name="BigPic" type="text" class="textfield" style="width:270;" value="<%=BigPic%>" maxlength="100">&nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=BigPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
	  <tr>
        <td align="right">链接网址：</td>
        <td colspan="5"><input name="GotoUrl" type="text" class="textfield" style="width:270;" value="<%=GotoUrl%>" maxlength="100"></td>
      </tr>
	  <tr class="LangShowI">
        <td align="right"><%=LangWebI%>简要概述：</td>
        <td colspan="5"><textarea id="ConciseLangI" name="ConciseLangI" rows="1" cols="60"><%=ConciseLangI%></textarea></td>
      </tr>
	  <tr class="LangShowD">
        <td align="right"><%=LangWebD%>简要概述：</td>
        <td colspan="5"><textarea id="ConciseLangD" name="ConciseLangD" rows="1" cols="60"><%=ConciseLangD%></textarea></td>
      </tr>
	  <tr class="LangShowE">
        <td align="right"><%=LangWebE%>简要概述：</td>
        <td colspan="5"><textarea id="ConciseLangE" name="ConciseLangE" rows="1" cols="60"><%=ConciseLangE%></textarea></td>
      </tr>
	  <tr>
	    <td align="right">&nbsp;</td>
	    <td colspan="5"><input name="submitSave" type="submit" class="button" id="保存" value=" 保存 "></td>
	   </tr>
    </table>
	</td>
  </tr>
</form>
</table>
<br>
<%
End Function
'转移节点表单------------------------------
Function moveFolderForm()
  Dim ID,ParentID,SortNameLangI,SortPath
  ID=request.QueryString("ID")
  ParentID=request.QueryString("ParentID")
  SortNameLangI=request.QueryString("SortNameLangI")
  SortPath=request.QueryString("SortPath")
%>
<table id="SortFromTo" cellpadding="0" cellspacing="0">
<form name="MoveForm" method="post" action="NewsSort.asp?Action=MoveSave">
  <tr>
    <th height="24" colspan="3" nowrap>类别移动：通过点击分类树中类别对应的"移"可重新选择将要作移动的类别，包括本类、子类及所有下属信息条目将一起被移动。</th>
  </tr>
  <tr>
    <td height="24" colspan="3">→&nbsp;<% response.write (SortNameLangI) %></td>
  </tr>
  <tr>
    <td>移动类ID：<input readonly name="ID" type="text" class="textfield" id="ID" size="14" value="<%=ID%>"></td>
    <td>移动类父ID：<input readonly name="ParentID" type="text" class="textfield" id="ParentID" size="14" value="<%=ParentID%>"></td>
    <td>移动类数字路径：<input readonly name="SortPath" type="text" class="textfield" id="SortPath" size="30" value="<%=SortPath%>"></td>
  </tr>
  <tr><td colspan="3"></td></tr>
  <tr>
    <th height="24" colspan="3">目标位置：通过点击"至"选择将要放置到的类别。</th>
  </tr>
  <tr>
    <td height="24" colspan="3">→&nbsp;请选择…</td>
  </tr>
  <tr>
    <td>目标类ID：<input readonly name="toID" type="text" class="textfield" id="toID" size="14" value=""></td>
    <td>目标类父ID：<input readonly name="toParentID" type="text" class="textfield" id="toParentID" size="14" value=""></td>
    <td>目标类数字路径：<input readonly name="toSortPath" type="text" class="textfield" id="toSortPath" size="30" value=""></td>
  </tr>
  <tr>
    <td height="40" colspan="3"><input name="submitMove" type="submit" class="button" id="转移" value="  转移  "></td>
  </tr>
</form>
</table>
<br>
<%
End Function
'保存转移节点------------------------------
Function saveMoveFolder()
  Dim rs,sql,fromID,fromParentID,fromSortPath,toID,toParentID,toSortPath,fromParentSortPath
  fromID=request.Form("ID")
  fromParentID=request.Form("ParentID")
  fromSortPath=request.Form("SortPath")
  toID=request.Form("toID")
  toParentID=request.Form("toParentID")
  toSortPath=request.Form("toSortPath")

  if toID="" or toParentID="" or toSortPath="" then
    response.write ("<script language=javascript> alert('没有选择移动的目标位置，请返回选择！');history.back(-1);</script>")
    response.end
  end if
  if fromParentID=0 then
    response.write ("<script language=javascript> alert('一级分类不能被移动，请返回选择！');history.back(-1);</script>")
    response.end
  end if
  if fromSortPath=toSortPath then
    response.write ("<script language=javascript> alert('选择的移动类别和目标位置相同了，请返回重新选择！');history.back(-1);</script>")
    response.end
  end if
  if Instr(toSortPath,fromSortPath)>0 or fromParentID=toID then
    response.write ("<script language=javascript> alert('不能将类别移动到本类或下属类里，请返回重新选择！');history.back(-1);</script>")
    response.end
  end if
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From IdeaNet_NewsSort where ID="&fromParentID
  rs.open sql,conn,0,1
  fromParentSortPath=rs("SortPath")
  conn.execute("update IdeaNet_NewsSort set SortPath='"&toSortPath&"'+Mid(SortPath,Len('"&fromParentSortPath&"')+1) where Instr(SortPath,'"&fromSortPath&"')>0")'更新类别数字路径
  conn.execute("update IdeaNet_NewsSort set ParentID='"&toID&"' where ID="&fromID)'更新类别父类ID
  conn.execute("update IdeaNet_News set SortPath='"&toSortPath&"'+Mid(SortPath,Len('"&fromParentSortPath&"')+1) where Instr(SortPath,'"&fromSortPath&"')>0")'更新信息数字路径
  response.write ("<script language=javascript> alert('移动类别成功，点击确定查看类别树！');location.replace('NewsSort.asp');</script>")
End Function
%>
