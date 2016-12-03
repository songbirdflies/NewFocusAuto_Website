<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<%Response.charset="utf-8"%>
<!--#include file="Include/Const.asp" -->
<!--#include file="Include/ConnSiteData.asp" -->
<!--#include file="Include/NoSqlHack.asp" -->
<%
dim LangPack,LangData,FilePath
LangPack="Chinese"
LangData="LangI"
FilePath=""
%>
<%
dim rs,sql,rss,SiteTitle,SiteUrl,ComName,Address,Gonggao,ZipCode,Telephone,Fax,Email,EmailUrl,Modetel,SystemPic,Keywords,Descriptions,Copyright,IcpNumber,MesViewFlag,i,ii,SortId,ParentID,SortPath,SeoTitle,SeoSort,SortName,SortHtmlUrlName,HtmlUrlName,PageUrlName,HtmlId,SitListshow,TdNumstr,searchkey
set rs = server.createobject("adodb.recordset")
sql="select top 1 * from IdeaNet_Site"
rs.open sql,conn,1,1
SiteTitle=rs("SiteTitle"&LangData)'网站标题
SiteUrl=rs("SiteUrl")'网站网址
ComName=rs("ComName"&LangData)'公司名称
Address=rs("Address"&LangData)'地址
ZipCode=rs("ZipCode")'邮编
Telephone=rs("Telephone")'电话
Fax=rs("Fax")'传真
Email=rs("Email")'邮箱
EmailUrl=rs("EmailUrl")'电子邮件发送网址
Modetel=rs("Modetel")'24小时手机热线
Keywords=rs("Keywords"&LangData)'关键字
Descriptions=rs("Descriptions"&LangData)'网站描述
SystemPic=rs("SystemPic"&LangData)'系统Banner
Gonggao=rs("Gonggao"&LangData)'公告
Copyright=rs("Copyright"&LangData)'底部版权信息
IcpNumber=rs("IcpNumber")'ICP 备案
MesViewFlag=rs("MesViewFlag")'自动通过审核
rs.close
set rs=Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="Include/NextPage.Asp" -->
<!--#include file="Include/ShowFunction.asp" -->