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
<TITLE>编辑管理员</TITLE>
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="Js/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Md5.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|103,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'判断是否具有管理权限
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,AdminName,Working,Password,vPassword,UserName,Purview,Explain,AddTime
ID=request.QueryString("ID")
if ID="" then ID=0
call AdminEdit() 
%>
<table id="bodycontent" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">

		<table cellpadding="0" cellspacing="0">
		  <tr>
			<th>网站管理员：添加，修改管理员信息</th>
		  </tr>
		  <tr>
			<td align="center" bgcolor="#EBF2F9"><a href="AdminEdit.asp?Result=Add" onClick='changeAdminFlag("添加管理员")'>添加管理员</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="AdminList.asp" onClick='changeAdminFlag("网站管理员")'>查看所有管理员</a></td>    
		  </tr>
		</table>
		<br>
		
<table class="edit" border="0" cellpadding="0" cellspacing="0">
  <form name="editForm" method="post" action="AdminEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" onSubmit="return CheckAdminEdit()">
  <tr>
    <td height="24"><table border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">登&nbsp;录&nbsp;名：</td>
        <td><input name="AdminName" type="text" class="textfield" id="AdminName" style="WIDTH: 120;" value="<%=AdminName%>" maxlength="16" <%if Result="Modify" then response.write ("readonly")%>>&nbsp;*&nbsp;3-10位字符，不可修改</td>
      </tr>
      <tr>
        <td height="20" align="right">生　　效：</td>
        <td><input name="Working" type="checkbox" id="ViewFlagSi" value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if Working then response.write ("checked")%>></td>
      </tr>
      <tr>
        <td height="20" align="right">密　　码：</td>
        <td><input name="Password" type="password" class="textfield" id="Password" maxlength="20" style="WIDTH: 120;">&nbsp;*&nbsp;6-16位字符，不填表未修改密码</td>
      </tr>
      <tr>
        <td height="20" align="right">确认密码：</td>
        <td><input name="vPassword" type="password" class="textfield" id="vPassword" maxlength="20" style="WIDTH: 120;">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">管理员名：</td>
        <td><input name="UserName" type="text" class="textfield" id="UserName" style="WIDTH: 120;" value="<%=UserName%>"></td>
      </tr>
	  <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">操作权限：</td>
        <td nowrap>
		  <input name="Purview11" type="checkbox" value="|11," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|11,")>0 then response.write ("checked")%>>&nbsp;系统主参数设置
		  <input name="Purview12" type="checkbox" value="|12," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|12,")>0 then response.write ("checked")%>>&nbsp;导航栏目列表
		  <input name="Purview13" type="checkbox" value="|13," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|13,")>0 then response.write ("checked")%>>&nbsp;编辑导航栏目
		  <input name="Purview14" type="checkbox" value="|14," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|14,")>0 then response.write ("checked")%>>&nbsp;数据库操作
		  <input name="Purview15" type="checkbox" value="|15," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|15,")>0 then response.write ("checked")%>>&nbsp;友情链接列表
		  <input name="Purview16" type="checkbox" value="|16," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|16,")>0 then response.write ("checked")%>>&nbsp;编辑友情链接</td>
      </tr>
	  	  <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td nowrap><input name="Purview21" type="checkbox" value="|21," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|21,")>0 then response.write ("checked")%>>&nbsp;企业列表
          <input name="Purview22" type="checkbox" value="|22," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|22,")>0 then response.write ("checked")%>>&nbsp;编辑企业</td>
      </tr>
	  <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td nowrap><input name="Purview31" type="checkbox" value="|31," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|31,")>0 then response.write ("checked")%>>&nbsp;新闻类别
          <input name="Purview32" type="checkbox" value="|32," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|32,")>0 then response.write ("checked")%>>&nbsp;新闻列表
          <input name="Purview33" type="checkbox" value="|33," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|33,")>0 then response.write ("checked")%>>&nbsp;编辑新闻</td>
      </tr>
	  <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td nowrap><input name="Purview41" type="checkbox" value="|41," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|41,")>0 then response.write ("checked")%>>&nbsp;商品类别
          <input name="Purview42" type="checkbox" value="|42," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|42,")>0 then response.write ("checked")%>>&nbsp;商品列表
          <input name="Purview43" type="checkbox" value="|43," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|43,")>0 then response.write ("checked")%>>&nbsp;编辑商品</td>
      </tr>
	  <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td nowrap><input name="Purview51" type="checkbox" value="|51," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|51,")>0 then response.write ("checked")%>>&nbsp;下载类别
          <input name="Purview52" type="checkbox" value="|52," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|52,")>0 then response.write ("checked")%>>&nbsp;下载列表
          <input name="Purview53" type="checkbox" value="|53," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|53,")>0 then response.write ("checked")%>>&nbsp;编辑下载</td>
      </tr>
      <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td><input name="Purview61" type="checkbox" value="|61," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|61,")>0 then response.write ("checked")%>>&nbsp;信息类别
          <input name="Purview62" type="checkbox" value="|62," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|62,")>0 then response.write ("checked")%>>&nbsp;信息列表
          <input name="Purview63" type="checkbox" value="|63," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|63,")>0 then response.write ("checked")%>>&nbsp;编辑信息 </td>
      </tr>
      <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td><input name="Purview71" type="checkbox" value="|71," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|71,")>0 then response.write ("checked")%>>&nbsp;单页列表
          <input name="Purview72" type="checkbox" value="|72," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|72,")>0 then response.write ("checked")%>>&nbsp;编辑单页</td>
      </tr>
	  <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td><input name="Purview81" type="checkbox" value="|81," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|81,")>0 then response.write ("checked")%>>&nbsp;广告类别
          <input name="Purview82" type="checkbox" value="|82," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|82,")>0 then response.write ("checked")%>>&nbsp;广告列表
          <input name="Purview83" type="checkbox" value="|83," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|83,")>0 then response.write ("checked")%>>&nbsp;编辑广告 </td>
      </tr>
      <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td><input name="Purview91" type="checkbox" value="|91," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|91,")>0 then response.write ("checked")%>>&nbsp;留言列表
		  <input name="Purview92" type="checkbox" value="|92," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|92,")>0 then response.write ("checked")%>>&nbsp;编辑留言</td>
      </tr>
      <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td><input name="Purview101" type="checkbox" value="|101," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|101,")>0 then response.write ("checked")%>>&nbsp;管理员密码修改
          <input name="Purview102" type="checkbox" value="|102," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|102,")>0 then response.write ("checked")%>>&nbsp;网站管理员列表
		  <input name="Purview103" type="checkbox" value="|103," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|103,")>0 then response.write ("checked")%>>&nbsp;编辑网站管理员</td>
      </tr>
      <tr <%if ID<>1 then response.write ("style=display:none")%>>
        <td height="20" align="right">操作权限：</td>
        <td nowrap><font color="#FF0000">内置超级管理员帐号，不可修改！</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">备注说明：</td>
        <td><textarea name="Explain" cols="88" rows="3" class="textfield" id="Explain" style="WIDTH: 580;" ><%=Explain%></textarea></td>
      </tr>

      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 60;" ></td>
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
sub AdminEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑管理员信息
    set rs = server.createobject("adodb.recordset")
    if Result="Add" then '创建网站管理员
      set rsCheckAdd = conn.execute("select AdminName from IdeaNet_Admin where AdminName='" & trim(Request.Form("AdminName")) & "'")
      if not (rsCheckAdd.bof and rsCheckAdd.eof) then '判断此管理员名是否存在
        response.write "<script language=javascript> alert('" & trim(Request.Form("AdminName")) & "管理员已经存在，请换一个登录名再试试！');history.back(-1);</script>"
        response.end
      end if  
	  sql="select * from IdeaNet_Admin"
      rs.open sql,conn,1,3
      rs.addnew
      if len(trim(Request.Form("AdminName")))<3 or len(trim(Request.Form("Password")))>10  then
        response.write "<script language=javascript> alert('管理员登录名必填，且字符数为3-10位！');history.back(-1);</script>"
        response.end
      end if	  
      if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>16  then
        response.write "<script language=javascript> alert('管理员密码必填，且字符数为6-16位！');history.back(-1);</script>"
        response.end
      end if
	  if Request.Form("Password")<>Request.Form("vPassword") then 
        response.write "<script language=javascript> alert('两次输入的密码不一样！');history.back(-1);</script>"
        response.end
	  end if
      rs("AdminName")=trim(Request.Form("AdminName"))
	  if Request.Form("Working")=1 then
        rs("Working")=Request.Form("Working")
	  else
        rs("Working")=0
	  end if
	  rs("Password")=Md5(Request.Form("Password"))
	  rs("UserName")=trim(Request.Form("UserName"))
	  rs("AdminPurview")=Request.Form("Purview11") & Request.Form("Purview12") & Request.Form("Purview13") & Request.Form("Purview14") & Request.Form("Purview15") & Request.Form("Purview16") &_
	                     Request.Form("Purview21") & Request.Form("Purview22") &_
	                     Request.Form("Purview31") & Request.Form("Purview32") & Request.Form("Purview33") &_
	                     Request.Form("Purview41") & Request.Form("Purview42") & Request.Form("Purview43") &_
	                     Request.Form("Purview51") & Request.Form("Purview52") & Request.Form("Purview53") &_
	                     Request.Form("Purview61") & Request.Form("Purview62") & Request.Form("Purview63") &_
	                     Request.Form("Purview71") & Request.Form("Purview72") &_
	                     Request.Form("Purview81") & Request.Form("Purview82") & Request.Form("Purview83") &_
	                     Request.Form("Purview91") & Request.Form("Purview92") &_
						 Request.Form("Purview101") & Request.Form("Purview102") & Request.Form("Purview103")
	  rs("Explain")=trim(Request.Form("Explain"))
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '修改网站管理员
      sql="select * from IdeaNet_Admin where ID="&ID
      rs.open sql,conn,1,3
      rs("AdminName")=trim(Request.Form("AdminName"))
	  if Request.Form("Working")=1 then
        rs("Working")=Request.Form("Working")
	  else
        rs("Working")=0
	  end if
      if trim(Request.Form("Password"))<>"" then
	    if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>20  then
          response.write "<script language=javascript> alert('管理员密码必填，且字符数为6-20位！');history.back(-1);</script>"
          response.end
        end if
	    if Request.Form("Password")<>Request.Form("vPassword") then 
          response.write "<script language=javascript> alert('两次输入的密码不一样！');history.back(-1);</script>"
          response.end
	    end if
	    rs("Password")=Md5(Request.Form("Password"))
	  end if
	  rs("UserName")=trim(Request.Form("UserName"))
	  rs("AdminPurview")=Request.Form("Purview11") & Request.Form("Purview12") & Request.Form("Purview13") & Request.Form("Purview14") & Request.Form("Purview15") & Request.Form("Purview16") &_
	                     Request.Form("Purview21") & Request.Form("Purview22") &_
	                     Request.Form("Purview31") & Request.Form("Purview32") & Request.Form("Purview33") &_
	                     Request.Form("Purview41") & Request.Form("Purview42") & Request.Form("Purview43") &_
	                     Request.Form("Purview51") & Request.Form("Purview52") & Request.Form("Purview53") &_
	                     Request.Form("Purview61") & Request.Form("Purview62") & Request.Form("Purview63") &_
	                     Request.Form("Purview71") & Request.Form("Purview72") &_
	                     Request.Form("Purview81") & Request.Form("Purview82") & Request.Form("Purview83") &_
	                     Request.Form("Purview91") & Request.Form("Purview92") &_
						 Request.Form("Purview101") & Request.Form("Purview102") & Request.Form("Purview103")
	  rs("Explain")=trim(Request.Form("Explain"))
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑网站管理员！');location.replace('AdminList.asp');</script>"
  else '提取管理员信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from IdeaNet_Admin where ID="& ID
      rs.open sql,conn,1,1
	  AdminName=rs("AdminName")
	  Working=rs("Working")
	  UserName=rs("UserName")
	  Purview=rs("AdminPurview")
	  Explain=rs("Explain")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
  
%>
