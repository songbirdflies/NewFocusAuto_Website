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
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2008-2018 - www.ideanet.cn" />
<META NAME="Author" CONTENT="艾迪创意,www.ideanet.cn" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>网站信息设置</TITLE>
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="Js/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|11,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'判断是否具有管理权限
%>
<body>
<table id="bodycontent" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
<%
dim ID,SiteTitleLangI,SiteTitleLangD,SiteTitleLangE,SiteUrl,ComNameLangI,ComNameLangD,ComNameLangE,AddressLangI,AddressLangD,AddressLangE
dim ZipCode,Telephone,Fax,Email,EmailUrl,Modetel,SystemPicLangI,SystemPicLangD,SystemPicLangE
dim KeywordsLangI,DescriptionsLangI,KeywordsLangD,DescriptionsLangD,KeywordsLangE,DescriptionsLangE
dim GonggaoLangI,GonggaoLangD,GonggaoLangE,CopyrightLangI,CopyrightLangD,CopyrightLangE,IcpNumber,SystemSn,MesViewFlag
dim fso,hf,ReplaceConstChar

select case request.QueryString("Action")
  case "Save"
    SaveSiteInfo
  case "SaveConst"
    SaveConstInfo
  case else
    ViewSiteInfo
end select
%>
<table cellpadding="0" cellspacing="0">
  <tr>
    <th>系统管理：添加，修改站点的相关信息</th>
  </tr>
  <tr>
    <td height="24" align="center">【系统主参数设置】</td>
  </tr>
</table>
<br>
<table cellpadding="0" cellspacing="0">
  <tr>
    <td height="24">
	<table cellpadding="0" cellspacing="0" class="noborder" id=editProduct idth="100%">
    <form name="editForm" method="post" action="?Action=Save" >
      <tr>
        <td width="150" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr class="LangShowI">
        <td height="20" align="right"><%=LangWebI%>网站标题：</td>
        <td><input name="SiteTitleLangI" type="text" class="textfield" id="SiteTitleLangI" style="width:400;" value="<%=SiteTitleLangI%>" >&nbsp;</td>
      </tr>
      <tr class="LangShowD">
        <td height="20" align="right"><%=LangWebD%>网站标题：</td>
        <td><input name="SiteTitleLangD" type="text" class="textfield" id="SiteTitleLangD" style="width:400;" value="<%=SiteTitleLangD%>" >&nbsp;</td>
      </tr>
      <tr class="LangShowE">
        <td height="20" align="right"><%=LangWebE%>网站标题：</td>
        <td><input name="SiteTitleLangE" type="text" class="textfield" id="SiteTitleLangE" style="width:400;" value="<%=SiteTitleLangE%>" >&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">网　　址：</td>
        <td><input name="SiteUrl" type="text" class="textfield" id="SiteUrl" style="width:400;" value="<%=SiteUrl%>">&nbsp;</td>
      </tr>
      <tr class="LangShowI">
        <td height="20" align="right">公司<%=LangWebI%>名称：</td>
        <td><input name="ComNameLangI" type="text" class="textfield" id="ComNameLangI" style="width:400;" value="<%=ComNameLangI%>" >&nbsp;</td>
      </tr>
      <tr class="LangShowD">
        <td height="20" align="right">公司<%=LangWebD%>名称：</td>
        <td><input name="ComNameLangD" type="text" class="textfield" id="ComNameLangD" style="width:400;" value="<%=ComNameLangD%>" >&nbsp;</td>
      </tr>
      <tr class="LangShowE">
        <td height="20" align="right">公司<%=LangWebE%>名称：</td>
        <td><input name="ComNameLangE" type="text" class="textfield" id="ComNameLangE" style="width:400;" value="<%=ComNameLangE%>" >&nbsp;</td>
      </tr>
      <tr class="LangShowI">
        <td height="20" align="right"><%=LangWebI%>地址：</td>
        <td><input name="AddressLangI" type="text" class="textfield" id="AddressLangI" style="width:400;" value="<%=AddressLangI%>" >&nbsp;</td>
      </tr>
      <tr class="LangShowD">
        <td height="20" align="right"><%=LangWebD%>地址：</td>
        <td><input name="AddressLangD" type="text" class="textfield" id="AddressLangD" style="width:400;" value="<%=AddressLangD%>" >&nbsp;</td>
      </tr>
      <tr class="LangShowE">
        <td height="20" align="right"><%=LangWebE%>地址：</td>
        <td><input name="AddressLangE" type="text" class="textfield" id="AddressLangE" style="width:400;" value="<%=AddressLangE%>" >&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">邮　　编：</td>
        <td><input name="ZipCode" type="text" class="textfield" id="ZipCode" style="width:200;" value="<%=ZipCode%>">&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">电　　话：</td>
        <td><input name="Telephone" type="text" class="textfield" id="Telephone" style="width:200;" value="<%=Telephone%>">&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">传　　真：</td>
        <td><input name="Fax" type="text" class="textfield" id="Fax" style="width:200;" value="<%=Fax%>" >&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">E-mail：</td>
        <td><input name="Email" type="text" class="textfield" id="Email" style="width:200;" value="<%=Email%>">&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">电子邮件发送网址：</td>
        <td><input name="EmailUrl" type="text" class="textfield" id="EmailUrl" style="width:200;" value="<%=EmailUrl%>" >&nbsp;</td>
      </tr>
	  <tr>
        <td height="20" align="right">24小时手机热线：</td>
        <td><input name="Modetel" type="text" class="textfield" id="Modetel" style="width:200;" value="<%=Modetel%>">&nbsp;</td>
      </tr>
	  <tr class="LangShowI">
        <td height="20" align="right"><%=LangWebI%>Logo图片：</td>
        <td><input name="SystemPicLangI" type="text" class="textfield" id="SystemPicLangI" style="width:500;" value="<%=SystemPicLangI%>" >
          <a href="javaScript:OpenScript('UpFileForm.asp?Result=SystemPicLangI',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
      </tr>
      <tr class="LangShowD">
        <td height="20" align="right"><%=LangWebD%>Logo图片：</td>
        <td><input name="SystemPicLangD" type="text" class="textfield" id="SystemPicLangD" style="width:500;" value="<%=SystemPicLangD%>" >
          <a href="javaScript:OpenScript('UpFileForm.asp?Result=SystemPicLangD',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
      </tr>
      <tr class="LangShowE">
        <td height="20" align="right"><%=LangWebE%>Logo图片：</td>
        <td><input name="SystemPicLangE" type="text" class="textfield" id="SystemPicLangE" style="width:500;" value="<%=SystemPicLangE%>" >
          <a href="javaScript:OpenScript('UpFileForm.asp?Result=SystemPicLangE',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
      </tr>
	  <tr class="LangShowI">
        <td height="20" align="right" valign="top"><%=LangWebI%>关 键 字：</td>
        <td><textarea name="KeywordsLangI" rows="1"  class="textfield" id="KeywordsLangI" style="width:500;"><%=KeywordsLangI%></textarea>
          关键字设置有利于网站的搜索</td>
      </tr>
      <tr class="LangShowI">
        <td height="20" align="right" valign="top"><%=LangWebI%>网站描述：</td>
        <td><textarea name="DescriptionsLangI" rows="1" class="textfield" id="DescriptionsLangI" style="width:500;"><%=DescriptionsLangI%></textarea>&nbsp;网站描述设置有利于网站的搜索</td>
      </tr>
	  <tr class="LangShowD">
        <td height="20" align="right" valign="top"><%=LangWebD%>关 键 字：</td>
        <td><textarea name="KeywordsLangD" rows="1"  class="textfield" id="KeywordsLangD" style="width:500;"><%=KeywordsLangD%></textarea>&nbsp;</td>
      </tr>
      <tr class="LangShowD">
        <td height="20" align="right" valign="top"><%=LangWebD%>网站描述：</td>
        <td><textarea name="DescriptionsLangD" rows="1" class="textfield" id="DescriptionsLangD" style="width:500;"><%=DescriptionsLangD%></textarea>&nbsp;</td>
      </tr>
	  <tr class="LangShowE">
        <td height="20" align="right" valign="top"><%=LangWebE%>关 键 字：</td>
        <td><textarea name="KeywordsLangE" rows="1"  class="textfield" id="KeywordsLangE" style="width:500;"><%=KeywordsLangE%></textarea>&nbsp;</td>
      </tr>
      <tr class="LangShowE">
        <td height="20" align="right" valign="top"><%=LangWebE%>网站描述：</td>
        <td><textarea name="DescriptionsLangE" rows="1" class="textfield" id="DescriptionsLangE" style="width:500;"><%=DescriptionsLangE%></textarea>&nbsp;</td>
      </tr>
	  <tr class="LangShowI">
        <td height="20" align="right" valign="top"><%=LangWebI%>公告：</td>
        <td><textarea name="GonggaoLangI" rows="2" class="textfield" id="GonggaoLangI" style="width:500;"><%=GonggaoLangI%></textarea>&nbsp;</td>
      </tr>
	  <tr class="LangShowD">
        <td height="20" align="right" valign="top"><%=LangWebD%>公告：</td>
        <td><textarea name="GonggaoLangD" rows="2" class="textfield" id="GonggaoLangD" style="width:500;"><%=GonggaoLangD%></textarea>&nbsp;</td>
      </tr>
	  <tr class="LangShowE">
        <td height="20" align="right" valign="top"><%=LangWebE%>公告：</td>
        <td><textarea name="GonggaoLangE" rows="2" class="textfield" id="GonggaoLangE" style="width:500;"><%=GonggaoLangE%></textarea>&nbsp;</td>
      </tr>
	  <tr class="LangShowI">
        <td height="20" align="right" valign="top"><%=LangWebI%>版权信息：<br>(支持HTML代码)</td>
        <td><textarea name="CopyrightLangI" rows="3" class="textfield" id="CopyrightLangI" style="width:500;"><%=CopyrightLangI%></textarea>&nbsp;</td>
      </tr>
	  <tr class="LangShowD">
        <td height="20" align="right" valign="top"><%=LangWebD%>版权信息：<br>(支持HTML代码)</td>
        <td><textarea name="CopyrightLangD" rows="3" class="textfield" id="CopyrightLangD" style="width:500;"><%=CopyrightLangD%></textarea>&nbsp;</td>
      </tr>
	  <tr class="LangShowE">
        <td height="20" align="right" valign="top"><%=LangWebE%>版权信息：<br>(支持HTML代码)</td>
        <td><textarea name="CopyrightLangE" rows="3" class="textfield" id="CopyrightLangE" style="width:500;"><%=CopyrightLangE%></textarea>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">ICP&nbsp;备案：</td>
        <td><input name="IcpNumber" type="text" class="textfield" id="IcpNumber" style="width:200;" value="<%=IcpNumber%>"></td>
      </tr>
	  <tr>
        <td height="20" align="right">授&nbsp;权&nbsp;号：</td>
        <td><input name="SystemSn" type="text" class="textfield" id="SystemSn" style="width:200;" value="<%=SystemSn%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">留&nbsp;言&nbsp;簿：</td>
        <td><input name="MesViewFlag" type="checkbox" id="MesViewFlag" value="1" style="HEIGHT: 13px;width: 13px;" <%if MesViewFlag then response.write ("checked")%>>&nbsp;自动通过审核</td>
      </tr>
      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="保存主参数设置" />&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom">&nbsp;</td>
      </tr>
      </form>
    </table>
    <table cellpadding="0" cellspacing="0" class="noborder" id=editProduct idth="100%">
    <form name="SysForm" method="post" action="?Action=SaveConst" >
      <tr>
        <td width="150" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">系统安装目录：</td>
        <td><input name="SysRootDir" type="text" id="SysRootDir" style="width:200" value="<%=SysRootDir%>"> <font color="red">*</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">数据库目录：</td>
        <td><input name="SiteDataPath" type="text" id="SiteDataPath" style="width:200" value="<%=SiteDataPath%>"> <font color="red">*</font>        </td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">页面防刷新时间：</td>
        <td><input name="Refresh" type="text" id="Refresh" style="width:80" value="<%=Refresh%>"> 秒 <font color="red">*</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">生成静态页面后缀：</td>
        <td><input name="HtmlName" type="text" id="HtmlName" style="width:80" value="<%=HtmlName%>"> <font color="red">*</font> 如：New.html中“html”即为后缀。可设置如：html、htm、Asp</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">分隔符：</td>
        <td><input name="Separated" type="text" id="Separated" style="width:80" value="<%=Separated%>"> <font color="red">*</font> 如：New-1.html中的“-”即为分隔符</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">企业信息生成前缀：</td>
        <td><input name="AboutNameDiy" type="text" id="AboutNameDiy" style="width:200" value="<%=AboutNameDiy%>"> <font color="red">*</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">新闻分类生成前缀：</td>
        <td><input name="NewsSortNameDiy" type="text" id="NewsSortNameDiy" style="width:200" value="<%=NewsSortNameDiy%>"> <font color="red">*</font> 如：News-1.html中“News”即为前缀</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">新闻分类页显示条数：</td>
        <td><input name="NewsPageDiy" type="text" id="NewsPageDiy" style="width:200" value="<%=NewsPageDiy%>"> <font color="red">*</font> 如：默认为30</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">新闻详细页生成前缀：</td>
        <td><input name="NewsNameDiy" type="text" id="NewsNameDiy" style="width:200" value="<%=NewsNameDiy%>"> <font color="red">*</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">商品/作品分类生成前缀：</td>
        <td><input name="ProductsSortNameDiy" type="text" id="ProductsSortNameDiy" style="width:200" value="<%=ProductsSortNameDiy%>"> <font color="red">*</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">商品/作品分类页显示条数：</td>
        <td><input name="ProductsPageDiy" type="text" id="ProductsPageDiy" style="width:200" value="<%=ProductsPageDiy%>"> <font color="red">*</font> 如：默认为30</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">商品/作品详细页生成前缀：</td>
        <td><input name="ProductsNameDiy" type="text" id="ProductsNameDiy" style="width:200" value="<%=ProductsNameDiy%>"> <font color="red">*</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">下载分类生成前缀：</td>
        <td><input name="DownLoadSortNameDiy" type="text" id="DownLoadSortNameDiy" style="width:200" value="<%=DownLoadSortNameDiy%>"> <font color="red">*</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">下载分类页显示条数：</td>
        <td><input name="DownLoadPageDiy" type="text" id="DownLoadPageDiy" style="width:200" value="<%=DownLoadPageDiy%>"> <font color="red">*</font> 如：默认为30</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">下载详细页生成前缀：</td>
        <td><input name="DownLoadNameDiy" type="text" id="DownLoadNameDiy" style="width:200" value="<%=DownLoadNameDiy%>"> <font color="red">*</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">信息中心分类生成前缀：</td>
        <td><input name="InfoSortNameDiy" type="text" id="InfoSortNameDiy" style="width:200" value="<%=InfoSortNameDiy%>"> <font color="red">*</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">信息中心分类页显示条数：</td>
        <td><input name="InfoPageDiy" type="text" id="InfoPageDiy" style="width:200" value="<%=InfoPageDiy%>"> <font color="red">*</font> 如：默认为30</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">信息中心详细页生成前缀：</td>
        <td><input name="InfoNameDiy" type="text" id="InfoNameDiy" style="width:200" value="<%=InfoNameDiy%>"> <font color="red">*</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">单页内容生成前缀：</td>
        <td><input name="ContentNameDiy" type="text" id="ContentNameDiy" style="width:200" value="<%=ContentNameDiy%>"> <font color="red">*</font></td>
      </tr>
      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td><input name="submitSaveEdit" type="submit" class="button" id="submitSaveEdit" value="保存附加参数设置" />&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom">&nbsp;</td>
      </tr>
      </form>
    </table>
    </td>
  </tr>
</table>
</td></tr></table>
</body>
</html>
<%
function SaveSiteInfo
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from IdeaNet_Site"
  rs.open sql,conn,1,3
  rs("SiteTitleLangI")=trim(Request.Form("SiteTitleLangI"))
  rs("SiteTitleLangD")=trim(Request.Form("SiteTitleLangD"))
  rs("SiteTitleLangE")=trim(Request.Form("SiteTitleLangE"))
  rs("SiteUrl")=trim(Request.Form("SiteUrl"))
  rs("ComNameLangI")=trim(Request.Form("ComNameLangI"))
  rs("ComNameLangD")=trim(Request.Form("ComNameLangD"))
  rs("ComNameLangE")=trim(Request.Form("ComNameLangE"))
  rs("AddressLangI")=trim(Request.Form("AddressLangI"))
  rs("AddressLangD")=trim(Request.Form("AddressLangD"))
  rs("AddressLangE")=trim(Request.Form("AddressLangE"))
  rs("ZipCode")=trim(Request.Form("ZipCode"))
  rs("Telephone")=trim(Request.Form("Telephone"))
  rs("Fax")=trim(Request.Form("Fax"))
  rs("Email")=trim(Request.Form("Email"))
  rs("EmailUrl")=trim(Request.Form("EmailUrl"))
  rs("Modetel")=trim(Request.Form("Modetel"))
  rs("SystemPicLangI")=trim(Request.Form("SystemPicLangI"))
  rs("SystemPicLangD")=trim(Request.Form("SystemPicLangD"))
  rs("SystemPicLangE")=trim(Request.Form("SystemPicLangE"))
  rs("KeywordsLangI")=trim(Request.Form("KeywordsLangI"))
  rs("DescriptionsLangI")=trim(Request.Form("DescriptionsLangI"))
  rs("KeywordsLangD")=trim(Request.Form("KeywordsLangD"))
  rs("DescriptionsLangD")=trim(Request.Form("DescriptionsLangD"))
  rs("KeywordsLangE")=trim(Request.Form("KeywordsLangE"))
  rs("DescriptionsLangE")=trim(Request.Form("DescriptionsLangE"))
  rs("GonggaoLangI")=trim(Request.Form("GonggaoLangI"))
  rs("GonggaoLangD")=trim(Request.Form("GonggaoLangD"))
  rs("GonggaoLangE")=trim(Request.Form("GonggaoLangE"))
  rs("CopyrightLangI")=trim(Request.Form("CopyrightLangI"))
  rs("CopyrightLangD")=trim(Request.Form("CopyrightLangD"))
  rs("CopyrightLangE")=trim(Request.Form("CopyrightLangE"))
  rs("IcpNumber")=trim(Request.Form("IcpNumber"))
  
  if Request.Form("MesViewFlag")=1 then
  rs("MesViewFlag")=1
  else
  rs("MesViewFlag")=0
  end if
  rs.update
  rs.close
  set rs=nothing 
  response.write "<script language=javascript> alert('系统主参数设置成功！');location.replace('SetSite.asp');</script>"
end function

function ViewSiteInfo
  dim rs,sql 
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from IdeaNet_Site"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write "读取数据库记录出错！"
    response.end
  else
    SiteTitleLangI=rs("SiteTitleLangI")
    SiteTitleLangD=rs("SiteTitleLangD")
    SiteTitleLangE=rs("SiteTitleLangE")
    SiteUrl=rs("SiteUrl")
    ComNameLangI=rs("ComNameLangI")
    ComNameLangD=rs("ComNameLangD")
    ComNameLangE=rs("ComNameLangE")
    AddressLangI=rs("AddressLangI")
    AddressLangD=rs("AddressLangD")
    AddressLangE=rs("AddressLangE")
	GonggaoLangI=rs("GonggaoLangI")
    GonggaoLangD=rs("GonggaoLangD")
    GonggaoLangE=rs("GonggaoLangE")
	CopyrightLangI=rs("CopyrightLangI")
    CopyrightLangD=rs("CopyrightLangD")
    CopyrightLangE=rs("CopyrightLangE")
    ZipCode=rs("ZipCode")
    Telephone=rs("Telephone")
    Fax=rs("Fax")
    Email=rs("Email")
	EmailUrl=rs("EmailUrl")
	Modetel=rs("Modetel")
    SystemPicLangI=rs("SystemPicLangI")
	SystemPicLangD=rs("SystemPicLangD")
	SystemPicLangE=rs("SystemPicLangE")
    KeywordsLangI=rs("KeywordsLangI")
    DescriptionsLangI=rs("DescriptionsLangI")
	KeywordsLangD=rs("KeywordsLangD")
    DescriptionsLangD=rs("DescriptionsLangD")
	KeywordsLangE=rs("KeywordsLangE")
    DescriptionsLangE=rs("DescriptionsLangE")
    IcpNumber=rs("IcpNumber")
    SystemSn=rs("SystemSn")
	MesViewFlag=rs("MesViewFlag")
    rs.close
    set rs=nothing 
  end if
end function

Function SaveConstInfo()
 set fso=Server.CreateObject("Scripting.FileSystemObject")
 set hf=fso.CreateTextFile(Server.mappath("../Include/Const.Asp"),true)
 hf.write "<" & "%" & vbcrlf
 hf.write "Const SysRootDir = " & chr(34) & trim(request("SysRootDir")) & chr(34) & "" & vbcrlf
 hf.write "Const SiteDataPath = " & chr(34) & trim(request("SiteDataPath")) & chr(34) & "" & vbcrlf
 hf.write "Const Refresh = " & trim(request("Refresh")) & "" & vbcrlf
 hf.write "Const HtmlName = " & chr(34) & trim(request("HtmlName")) & chr(34) & "" & vbcrlf
 hf.write "Const Separated = " & chr(34) & trim(request("Separated")) & chr(34) & "" & vbcrlf
 hf.write "Const AboutNameDiy = " & chr(34) & trim(request("AboutNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const NewsSortNameDiy = " & chr(34) & trim(request("NewsSortNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const NewsPageDiy = " & trim(request("NewsPageDiy")) & "" & vbcrlf
 hf.write "Const NewsNameDiy = " & chr(34) & trim(request("NewsNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const ProductsSortNameDiy = " & chr(34) & trim(request("ProductsSortNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const ProductsPageDiy = " & trim(request("ProductsPageDiy")) & "" & vbcrlf
 hf.write "Const ProductsNameDiy = " & chr(34) & trim(request("ProductsNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const DownLoadSortNameDiy = " & chr(34) & trim(request("DownLoadSortNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const DownLoadPageDiy = " & trim(request("DownLoadPageDiy")) & "" & vbcrlf
 hf.write "Const DownLoadNameDiy = " & chr(34) & trim(request("DownLoadNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const InfoSortNameDiy = " & chr(34) & trim(request("InfoSortNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const InfoPageDiy = " & trim(request("InfoPageDiy")) & "" & vbcrlf
 hf.write "Const InfoNameDiy = " & chr(34) & trim(request("InfoNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "Const ContentNameDiy = " & chr(34) & trim(request("ContentNameDiy")) & chr(34) & "" & vbcrlf
 hf.write "%" & ">"
 hf.close
 set hf=nothing
 set fso=nothing
 response.Write "<script language=javascript>alert('系统默认附加参数设置成功！');location.href='SetSite.Asp';</script>"
End Function
%>
