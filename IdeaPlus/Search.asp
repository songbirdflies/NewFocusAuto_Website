<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->

<%
dim Result
  Result=request.QueryString("Result")
dim StartDate,EndDate,Keyword
  StartDate=request.form("Start_Date")
  EndDate=request.form("End_Date")
  Keyword=request.form("Keyword")
select case Result
  case "News"
    response.redirect ("NewsList.asp?Result=Search&StartDate="&StartDate&"&EndDate="&EndDate&"&Keyword="&server.urlencode(Keyword)&"&Page=1")
  case "Down"
    response.redirect ("DownList.asp?Result=Search&StartDate="&StartDate&"&EndDate="&EndDate&"&Keyword="&server.urlencode(Keyword)&"&Page=1")
  case "Products"
    response.redirect ("ProductsList.asp?Result=Search&StartDate="&StartDate&"&EndDate="&EndDate&"&Keyword="&server.urlencode(Keyword)&"&Page=1")
  case "Ads"
    response.redirect ("AdsList.asp?Result=Search&StartDate="&StartDate&"&EndDate="&EndDate&"&Keyword="&server.urlencode(Keyword)&"&Page=1")	
  case "Jobs"
    response.redirect ("JobsList.asp?Result=Search&StartDate="&StartDate&"&EndDate="&EndDate&"&Keyword="&server.urlencode(Keyword)&"&Page=1")	
  case "Message"
    response.redirect ("MessageList.asp?Result=Search&StartDate="&StartDate&"&EndDate="&EndDate&"&Keyword="&server.urlencode(Keyword)&"&Page=1")		
  case "Supply"
    response.redirect ("SupplyList.asp?Result=Search&StartDate="&StartDate&"&EndDate="&EndDate&"&Keyword="&server.urlencode(Keyword)&"&Page=1")	
  case "Talents"
    response.redirect ("TalentsList.asp?Result=Search&StartDate="&StartDate&"&EndDate="&EndDate&"&Keyword="&server.urlencode(Keyword)&"&Page=1")	
  case "Others"
    response.redirect ("InfoList.asp?Result=Search&StartDate="&StartDate&"&EndDate="&EndDate&"&Keyword="&server.urlencode(Keyword)&"&Page=1")	
  case else
end select
%>
