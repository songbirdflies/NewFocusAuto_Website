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
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->

<%
dim Result,Selectid
Result=request.QueryString("Result")
SelectID=request.Form("SelectID")
select case Result
  case "Administrators"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_Admin where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "LoginLog"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_AdminLog where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "About"
    if SelectID<>"" then  
		if instr(SelectID,("28"))>0 then
			response.Write("<script language=""javascript"">alert('所选项目中有系统ID不能删除！');</script>")
		elseif instr(SelectID,("34"))>0 then
			response.Write("<script language=""javascript"">alert('所选项目中有系统ID不能删除！');</script>")
		elseif instr(SelectID,("38"))>0 then
			response.Write("<script language=""javascript"">alert('所选项目中有系统ID不能删除！');</script>")
		else
			conn.execute "delete from IdeaNet_About where id in ("&SelectID&")"
		end if
	end if
    response.write("<script language=""javascript"">window.location.href='AboutList.asp';</script>")
  case "Products"
    if SelectID<>"" then
		conn.execute "delete from IdeaNet_Products where id in ("&SelectID&")"
		conn.execute "delete from IdeaNet_TmpProperty where showID in ("&SelectID&") "
	end if
    response.redirect request.servervariables("http_referer")
  case "News"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_News where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "DownLoad"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_Download where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Ads"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_Ads where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Content"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_Content where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Message"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_Message where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Supply"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_Supply where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Talents"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_Talents where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Format"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_Format where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "FriendSite"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_FriendSite where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Info"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_Info where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Vote"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_vote where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "NoHackSql"
    if SelectID<>"" then  conn.execute "delete from IdeaNet_NoHackSql where SqlIn_ID in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case else	
end select
%>
