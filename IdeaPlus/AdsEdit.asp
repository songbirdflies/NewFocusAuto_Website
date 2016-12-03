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
<TITLE>编辑广告</TITLE>
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="Js/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|83,")=0 then 
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
dim ID,TitleNameLangI,TitleNameLangD,TitleNameLangE,ViewFlagLangI,ViewFlagLangD,ViewFlagLangE,SortNameLangI,SortID,SortPath
dim Types,Content,GotoUrl,Width,Height
dim ConciseLangI,ConciseLangD,ConciseLangE,ContentLangI,ContentlangD,ContentlangE,AddTime
ID=request.QueryString("ID")
call AdsEdit() 
%>
<table cellpadding="0" cellspacing="0">
  <tr>
    <th>广告检索及分类查看：添加，修改，删除广告</th>
  </tr>
  <tr>
    <td align="center">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>广告】</td>
  </tr>
</table>
<br>
<table cellpadding="0" cellspacing="0">
  <form name="editForm" method="post" action="AdsEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24"><table class="noborder" cellpadding="0" cellspacing="0">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr class="LangShowI">
        <td height="20" align="right"><%=LangWebI%>标题：</td>
        <td><input name="TitleNameLangI" type="text" class="textfield" id="TitleNameLangI" style="width: 240;" value="<%=TitleNameLangI%>" maxlength="100">&nbsp;显示：<input name="ViewFlagLangI" type="checkbox" style='height: 13px;width: 13px;' value="1" <%if ViewFlagLangI then response.write ("checked")%>>&nbsp;不少于1个字符</td>
      </tr>
      <tr class="LangShowD">
        <td height="20" align="right"><%=LangWebD%>标题：</td>
        <td><input name="TitleNameLangD" type="text" class="textfield" id="TitleNameLangD" style="width: 240;" value="<%=TitleNameLangD%>" maxlength="100">&nbsp;显示：<input name="ViewFlagLangD" type="checkbox" value="1" style='height: 13px;width: 13px;' <%if ViewFlagLangD then response.write ("checked")%>></td>
      </tr>
      <tr class="LangShowE">
        <td height="20" align="right"><%=LangWebE%>标题：</td>
        <td><input name="TitleNameLangE" type="text" class="textfield" id="TitleNameLangE" style="width: 240;" value="<%=TitleNameLangE%>" maxlength="100">&nbsp;显示：<input name="ViewFlagLangE" type="checkbox" value="1" style='height: 13px;width: 13px;' <%if ViewFlagLangE then response.write ("checked")%>></td>
      </tr>
      <tr>
        <td height="20" align="right">所属类别：</td>
        <td><input name="SortNameLangI" type="text" class="textfield" id="SortNameLangI" value="<%=SortNameLangI%>" style="width: 240;background-color:#EBF2F9;" readonly>&nbsp;<a href="javaScript:OpenScript('SelectSort.asp?Result=Ads',500,500,'')"><img src="Images/Select.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
      <tr>
        <td height="20" align="right">类别数字：</td>
        <td><input name="SortID" type="text" class="textfield" id="SortID" style="width: 40;background-color:#EBF2F9;" value="<%=SortID%>" readonly><input name="SortPath" type="text" class="textfield" id="SortPath" style="width: 200;background-color:#EBF2F9;" value="<%=SortPath%>" readonly>&nbsp;*</td>
      </tr>
	  <tr>
      <td height="20" align="right">展现方式：</td>
	    <td><input value="1" name="Types" type="radio" id="rblPic" <%If Types="" Or Types="1" Then response.write "checked"%> />图片&nbsp;&nbsp;<input value="2" name="Types" type="radio" id="RbJs" <%If  Types="2" Then response.write "checked"%> />Js&nbsp;&nbsp;<input value="3" name="Types" type="radio" id="RbFlash" <%If  Types="3" Then response.write "checked"%> />Flash
	    </td>
       </tr>
	   <tr>
        <td height="20" align="right">广告地址：</td>
        <td><textarea name="Content" rows="1" class="textfield" style="width: 56%;"><%=Content%></textarea>
        &nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=Content',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a> <img src="Images/memo.gif" alt="广告地址中为广告图,或JS代码或flash地址" /></td>
      </tr>
	  <tr>
        <td height="20" align="right">链接网址：</td>
        <td><input name="GotoUrl" type="text" class="textfield" id="GotoUrl" style="width: 240;" value="<%=GotoUrl%>" maxlength="100"> 
        <img src="Images/memo.gif" alt="无链接地址请为空" /></td>
	  </tr>
	  <tr>
	    <td height="20" align="right">广告宽度：</td>
	    <td>
		<input name="Width" type="text" maxlength="25" id="Width" class="textfield" style="width:40px;" value="<%=Width%>" dataType="Require"  msg="请填写宽度" /> Px &nbsp;&nbsp; (0为不限宽)
	    </td>
      </tr>
	  <tr>
	    <td height="20" align="right">广告高度：</td>
	    <td>
		<input name="Height" type="text" maxlength="25" id="Height" class="textfield" style="width:40px;" value="<%=Height%>" dataType="Require"  msg="请填写高度" /> Px &nbsp;&nbsp; (0为不限高,可更改)
	    </td>
      </tr>
	  <tr class="LangShowI">
        <td height="20" align="right"><%=LangWebI%>简要概述：</td>
        <td><textarea name="ConciseLangI" rows="2" class="textfield" id="ConciseLangI" style="width: 86%;"><%=ConciseLangI%></textarea>&nbsp;</td>
	  </tr>
	  <tr class="LangShowD">
        <td height="20" align="right"><%=LangWebD%>简要概述：</td>
        <td><textarea name="ConciseLangD" rows="2" class="textfield" id="ConciseLangD" style="width: 86%;"><%=ConciseLangD%></textarea>&nbsp;</td>
	  </tr>
	  <tr class="LangShowE">
        <td height="20" align="right"><%=LangWebE%>简要概述：</td>
        <td><textarea name="ConciseLangE" rows="2" class="textfield" id="ConciseLangE" style="width: 86%;"><%=ConciseLangE%></textarea>&nbsp;</td>
	  </tr>
      <tr class="LangShowI">
		<td height="20" align="right" valign="top"><%=LangWebI%>详细内容：</td>
        <td><textarea name="ContentLangI" id="ContentLangI" style="display:none"><%=Server.HTMLEncode((ContentLangI))%></textarea>
        <iframe ID="eWebEditor1" src="../IdeEditor/Ewebeditor.htm?id=ContentLangI&style=coolblue" frameborder="0" scrolling="no" width="86%" height="450"></iframe>&nbsp;</td>
      </tr>
      <tr class="LangShowD">
		<td height="20" align="right" valign="top"><%=LangWebD%>详细内容：</td>
        <td><textarea name="ContentLangD" id="ContentLangD" style="display:none"><%=Server.HTMLEncode((ContentLangD))%></textarea>
        <iframe ID="eWebEditor1" src="../IdeEditor/Ewebeditor.htm?id=ContentLangD&style=coolblue" frameborder="0" scrolling="no" width="86%" height="450"></iframe>&nbsp;</td>
	  </tr>
	  <tr class="LangShowE">
		<td height="20" align="right" valign="top"><%=LangWebE%>详细内容：</td>
        <td><textarea name="ContentLangE" id="ContentLangE" style="display:none"><%=Server.HTMLEncode((ContentLangE))%></textarea>
        <iframe ID="eWebEditor1" src="../IdeEditor/Ewebeditor.htm?id=ContentLangE&style=coolblue" frameborder="0" scrolling="no" width="86%" height="450"></iframe>&nbsp;</td>
      </tr>
      <tr>
		<td height="20" align="right">发布时间：</td>
		<td><input name="AddTime" type="text" class="textfield" id="AddTime" style="WIDTH: 240;" value="<%if AddTime<>"" then Response.Write(""&AddTime&"") else Response.Write(""&now()&"")%>" maxlength="100">&nbsp;请按 <%=now%> 格式正确输入</td>
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
sub AdsEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑产品信息
    set rs = server.createobject("adodb.recordset")
    if request.Form("ViewFlagLangI")=1 then
		if len(trim(request.Form("TitleNameLangI")))<1 then
		  response.write ("<script language=javascript> alert('你选择了"""&LangWebI&"""显示，因此"&LangWebI&"名称必填！');history.back(-1);</script>")
		  response.end
		end if
	end if
    if request.Form("ViewFlagLangD")=1 then
      if len(trim(request.Form("TitleNameLangD")))<1 then
      response.write ("<script language=javascript> alert('您已经选择了"""&LangWebD&"""显示，因此"&LangWebD&"标题必填！');history.back(-1);</script>")
      response.end
      end if
    end if
    if request.Form("ViewFlagLangE")=1 then
      if len(trim(request.Form("TitleNameLangE")))<1 then
      response.write ("<script language=javascript> alert('您已经选择了"""&LangWebE&"""显示，因此"&LangWebE&"标题必填！');history.back(-1);</script>")
      response.end
      end if
    end if
	
	
    if Result="Add" then '添加广告
	  sql="select * from IdeaNet_Ads"
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
	  if Request.Form("SortID")<>"" and Request.Form("SortPath")<>"" then
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  else
        response.write ("<script language=javascript> alert('请选择所属分类！');history.back(-1);</script>")
        response.end
	  end if
	  rs("Types")=trim(Request.Form("Types"))
	  rs("Content")=trim(Request.Form("Content"))
	  rs("GotoUrl")=trim(Request.Form("GotoUrl"))
	  rs("Width")=trim(Request.Form("Width"))
	  rs("Height")=trim(Request.Form("Height"))
	  rs("ConciseLangI")=Request.Form("ConciseLangI")
	  rs("ConciseLangD")=Request.Form("ConciseLangD")
	  rs("ConciseLangE")=Request.Form("ConciseLangE")
	  rs("ContentLangI")=trim(Request.Form("ContentLangI"))
	  rs("ContentlangD")=trim(Request.Form("ContentlangD"))
	  rs("ContentlangE")=trim(Request.Form("ContentlangE"))
	  rs("OrderNum")=9999
	  if Request.Form("AddTime")<>"" then
        rs("AddTime")=Request.Form("AddTime")
	  else
        rs("AddTime")=now()
	  end if
	end if  
	
	
	if Result="Modify" then '修改广告
      sql="select * from IdeaNet_Ads where ID="&ID
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
	  if Request.Form("SortID")<>"" and Request.Form("SortPath")<>"" then
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  else
        response.write ("<script language=javascript> alert('请选择所属分类！');history.back(-1);</script>")
        response.end
	  end if
	  rs("Types")=trim(Request.Form("Types"))
	  rs("Content")=trim(Request.Form("Content"))
	  rs("GotoUrl")=trim(Request.Form("GotoUrl"))
	  rs("Width")=trim(Request.Form("Width"))
	  rs("Height")=trim(Request.Form("Height"))
	  rs("ConciseLangI")=Request.Form("ConciseLangI")
	  rs("ConciseLangD")=Request.Form("ConciseLangD")
	  rs("ConciseLangE")=Request.Form("ConciseLangE")
	  rs("ContentLangI")=trim(Request.Form("ContentLangI"))
	  rs("ContentlangD")=trim(Request.Form("ContentlangD"))
	  rs("ContentlangE")=trim(Request.Form("ContentlangE"))
	  if Request.Form("AddTime")<>"" then
        rs("AddTime")=Request.Form("AddTime")
	  else
        rs("AddTime")=now()
	  end if
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑广告！');location.replace('AdsList.asp');</script>"
  else '提取广告
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from IdeaNet_Ads where ID="& ID
      rs.open sql,conn,1,1
      if rs.bof and rs.eof then
        response.write ("数据库读取记录出错！")
        response.end
      end if
	  TitleNameLangI=rs("TitleNameLangI")
	  TitleNameLangD=rs("TitleNameLangD")
	  TitleNameLangE=rs("TitleNameLangE")
	  ViewFlagLangI=rs("ViewFlagLangI")
	  ViewFlagLangD=rs("ViewFlagLangD")
	  ViewFlagLangE=rs("ViewFlagLangE")
	  SortNameLangI=SortText(rs("SortID"))
	  SortID=rs("SortID")
	  SortPath=rs("SortPath")
	  Types=rs("Types")
	  Content=rs("Content")
	  GotoUrl=rs("GotoUrl")
	  Width=rs("Width")
	  Height=rs("Height")
	  ConciseLangI=rs("ConciseLangI")
      ConciseLangD=rs("ConciseLangD")
      ConciseLangE=rs("ConciseLangE")
      ContentLangI=rs("ContentLangI")
      ContentlangD=rs("ContentlangD")
      ContentlangE=rs("ContentlangE")
	  AddTime=rs("AddTime")
	  rs.close
      set rs=nothing 
    end if
  end if
end sub
%>

<%
'生成所属类别--------------------------
Function SortText(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From IdeaNet_AdsSort where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortNameLangI")
  rs.close
  set rs=nothing
End Function
%>