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
<TITLE>后台管理导航</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2008-2018 - www.ideanet.cn" />
<META NAME="Author" CONTENT="艾迪创意,www.ideanet.cn" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="css/master.css">
<script language="javascript" src="Js/Admin.js"></script>
<script>
function closewin() {
   if (opener!=null && !opener.closed) {
      opener.window.newwin=null;
      opener.openbutton.disabled=false;
      opener.closebutton.disabled=true;
   }
}

var count=0;//做计数器
var limit=new Array();//用于记录当前显示的哪几个菜单
var countlimit=1;//同时打开菜单数目，可自定义

function expandIt(el) {
   obj = eval("sub" + el);
   if (obj.style.display == "none") {
      obj.style.display = "block";//显示子菜单
      if (count<countlimit) {//限制2个
         limit[count]=el;//录入数组
         count++;
      }
      else {
         eval("sub" + limit[0]).style.display = "none";
         for (i=0;i<limit.length-1;i++) {limit[i]=limit[i+1];}//数组去掉头一位，后面的往前挪一位
         limit[limit.length-1]=el;
      }
   }
   else {
      obj.style.display = "none";
      var j;
      for (i=0;i<limit.length;i++) {if (limit[i]==el) j=i;}//获取当前点击的菜单在limit数组中的位置
      for (i=j;i<limit.length-1;i++) {limit[i]=limit[i+1];}//j以后的数组全部往前挪一位
      limit[limit.length-1]=null;//删除数组最后一位
      count--;
   }
}
</script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<BODY style="padding:20px 0px 0px 6px;background:url(Images/bodybg.jpg) repeat-x #003b80;" onmouseover="self.status='艾迪网络 - 您最佳网络服务提供商! www.ideanet.cn';return true">
<div id="leftnav">
	<table width="193" height="32" border="0" cellpadding="0" cellspacing="0">
	  <tr>
		<td class="SystemLeft"><img src="Images/lefttitle.jpg"></td>
	  </tr>
	</table>
	<div id="main1" class="paddiv" onclick=expandIt(1)>
	<table width="193" height="24" border="0" cellpadding="0" cellspacing="0">
	  <tr style="cursor: hand;">
		<td width="7"></td>
		<td class="SystemLefttitle">系统管理</td>
		<td width="7"></td>
	  </tr>
	</table>
	</div>
	<div id="sub1" class="SystemLeftlist" style="display:none">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	  	<td width="15"></td>
		<td class="list"><li><a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("网站参数设置")'>网站参数设置</a></li></td>
	  	<td width="10"></td>
	  </tr>
	  <tr>
	  	<td width="15"></td>
		<td class="list"><li><a href="NavigationMenu.asp" target="mainFrame" onClick='changeAdminFlag("导航栏目列表")'>导航栏目</a></li></td>
	  	<td width="10"></td>
	  </tr>
	  <tr>
	  	<td width="15"></td>
		<td class="list"><li><a href="FriendSiteList.asp" target="mainFrame" onClick='changeAdminFlag("显示格式管理")'>友情链接列表</a></li></td>
	  	<td width="10"></td>
	  </tr>  
	  <tr>
	  	<td width="15"></td>
		<td class="list"><li><a href="FriendSiteEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("友情链接添加")'>添加友情链接</a></li></td>
	  	<td width="10"></td>
	  </tr>
	  <tr>
	  	<td colspan="3" height="7"></td>
	  </tr>
	  </table><img src="Images/listbottom.jpg"></div>
      
    
    <div id="main11" class="paddiv" onclick=expandIt(11)>
	<table width="193" height="24" border="0" cellpadding="0" cellspacing="0">
	  <tr style="cursor: hand;">
		<td width="7" ></td>
		<td class="SystemLefttitle">静态页面管理</td>
		<td width="7" ></td>
	  </tr>
	</table>
	</div>
	
	<div id="sub11" class="SystemLeftlist" style="display:none">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr class="LangShowI">
		<td width="7" ></td>
		<td class="list"><li><a href="Admin_htmldo.asp" target="mainFrame" onClick='changeAdminFlag("生成静态页面列表")'>生成静态页面列表</a></li></td>
		<td width="7" ></td>
	  </tr>
	  <tr>
	  	<td colspan="3" height="7"></td>
	  </tr>
	</table><img src="Images/listbottom.jpg"></div>
    
    
	<div id="main2" class="paddiv" onclick=expandIt(2)>
	<table width="193" height="24" border="0" cellpadding="0" cellspacing="0">
	  <tr style="cursor: hand;">
		<td width="7" ></td>
		<td class="SystemLefttitle">企业信息管理</td>
		<td width="7"></td>
	  </tr>
	</table>
	</div>
	<div id="sub2" class="SystemLeftlist" style="display:none">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td width="15"></td>
		<td class="list"><li><a href="AboutList.asp" target="mainFrame" onClick='changeAdminFlag("企业信息列表")'>企业信息列表</a></li></td>
	  	<td width="10"></td>
	  </tr>
	  <tr>
		<td width="15"></td>
		<td class="list"><li><a href="AboutEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("添加企业信息")'>添加企业信息</a></li></td>
	  	<td width="10"></td>
	  </tr>
	  <tr>
	  	<td colspan="3" height="7"></td>
	  </tr>
	</table><img src="Images/listbottom.jpg"></div>
	
	<div id="main3" class="paddiv" onclick=expandIt(3)>
	<table width="193" height="24" border="0" cellpadding="0" cellspacing="0">
	  <tr style="cursor: hand;">
		<td width="7" ></td>
		<td class="SystemLefttitle">新闻资讯管理</td>
		<td width="7" ></td>
	  </tr>
	</table>
	</div>
	
	<div id="sub3" class="SystemLeftlist" style="display:none">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td width="7"></td>
		<td class="list"><li><a href="NewsSort.asp?Action=Add&ParentID=0" target="mainFrame" onClick='changeAdminFlag("新闻类别")'>新闻类别</a></li></td>
		<td width="7"></td>
	  </tr>
	  <tr>
		<td width="7"></td>
		<td class="list"><li><a href="NewsList.asp" target="mainFrame" onClick='changeAdminFlag("新闻列表")'>新闻列表</a></li></td>
		<td width="7"></td>
	  </tr>
	  <tr>
		<td width="7"></td>
		<td class="list"><li><a href="NewsEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("添加新闻")'>添加新闻</a></li></td>
		<td width="7"></td>
	  </tr>
	  <tr>
	  	<td colspan="3" height="7"></td>
	  </tr>
	</table><img src="Images/listbottom.jpg"></div>
	
	<div id="main4" class="paddiv" onclick=expandIt(4)>
	<table width="193" height="24" border="0" cellpadding="0" cellspacing="0">
	  <tr style="cursor: hand;">
		<td width="7" ></td>
		<td class="SystemLefttitle">商品/作品信息管理</td>
		<td width="7" ></td>
	  </tr>
	</table>
	</div>
	<div id="sub4" class="SystemLeftlist" style="display:none">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td width="7"></td>
		<td class="list"><li><a href="ProductsSort.asp?Action=Add&ParentID=0" target="mainFrame" onClick='changeAdminFlag("商品/作品类别")'>商品/作品类别</a></li></td>
		<td width="7"></td>
	  </tr>
	  <tr>
		<td width="7"></td>
		<td class="list"><li><a href="ProductsList.asp" target="mainFrame" onClick='changeAdminFlag("商品/作品列表")'>商品/作品列表</a></li></td>
		<td width="7"></td>
	  </tr>
	  <tr>
		<td width="7"></td>
		<td class="list"><li><a href="ProductsEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("添加商品/作品")'>添加商品/作品</a></li></td>
		<td width="7"></td>
	  </tr>
	  <tr>
	  	<td colspan="3" height="7"></td>
	  </tr>
	</table><img src="Images/listbottom.jpg"></div>
	
	
	
	<div id="main6" class="paddiv" onclick=expandIt(6)>
	<table width="193" height="24" border="0" cellpadding="0" cellspacing="0">
	  <tr style="cursor: hand;">
		<td width="7" ></td>
		<td class="SystemLefttitle">信息中心管理</td>
		<td width="7" ></td>
	  </tr>
	</table>
	</div>
	
	<div id="sub6" class="SystemLeftlist" style="display:none">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td width="7" ></td>
		<td class="list"><li><a href="InfoSort.asp?Action=Add&ParentID=0" target="mainFrame" onClick='changeAdminFlag("信息中心类别")'>信息类别</a></li></td>
		<td width="7" ></td>
	  </tr>
	  <tr>
		<td width="7" ></td>
		<td class="list"><li><a href="InfoList.asp" target="mainFrame" onClick='changeAdminFlag("信息中心列表")'>信息列表</a></li></td>
		<td width="7" ></td>
	  </tr>
	  <tr>
		<td width="7" ></td>
		<td class="list"><li><a href="InfoEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("添加信息中心")'>添加信息</a></li></td>
		<td width="7" ></td>
	   </tr>
	  <tr>
	  	<td colspan="3" height="7"></td>
	  </tr>
	</table><img src="Images/listbottom.jpg"></div>
	
	
	<div id="main7" class="paddiv" onclick=expandIt(7)>
	<table width="193" height="24" border="0" cellpadding="0" cellspacing="0">
	  <tr style="cursor: hand;">
		<td width="7" ></td>
		<td class="SystemLefttitle">单页面管理</td>
		<td width="7" ></td>
	  </tr>
	</table>
	</div>
	
	<div id="sub7" class="SystemLeftlist" style="display:none">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td width="7"></td>
		<td class="list"><li><a href="ContentList.asp" target="mainFrame" onClick='changeAdminFlag("单页面信息列表")'>单页面列表</a></li></td>
		<td width="7"></td>
	  </tr>
	  <tr>
		<td width="7"></td>
		<td class="list"><li><a href="ContentEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("添加单页面信息")'>添加单页面</a></li></td>
		<td width="7"></td>
	  </tr>	
	  <tr>
	  	<td colspan="3" height="7"></td>
	  </tr>
	</table><img src="Images/listbottom.jpg"></div>
	
	
	<div id="main8" class="paddiv" onclick=expandIt(8)>
	<table width="193" height="24" border="0" cellpadding="0" cellspacing="0">
	  <tr style="cursor: hand;">
		<td width="7" ></td>
		<td class="SystemLefttitle">广告信息管理</td>
		<td width="7" ></td>
	  </tr>
	</table>
	</div>
	
	<div id="sub8" class="SystemLeftlist" style="display:none">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td width="7" ></td>
		<td class="list"><li><a href="AdsSort.asp?Action=Add&ParentID=0" target="mainFrame" onClick='changeAdminFlag("广告信息类别")'>广告类别</a></li></td>
		<td width="7" ></td>
	  </tr>
	  <tr>
		<td width="7" ></td>
		<td class="list"><li><a href="AdsList.asp" target="mainFrame" onClick='changeAdminFlag("广告信息列表")'>广告列表</a></li></td>
		<td width="7" ></td>
	  </tr>
	  <tr>
		<td width="7" ></td>
		<td class="list"><li><a href="AdsEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("添加广告信息")'>添加广告</a></li></td>
		<td width="7" ></td>
	   </tr>
	  <tr>
	  	<td colspan="3" height="7"></td>
	  </tr>
	</table><img src="Images/listbottom.jpg"></div>
    
    
    <div id="main9" class="paddiv" onclick=expandIt(9)>
	<table width="193" height="24" border="0" cellpadding="0" cellspacing="0">
	  <tr style="cursor: hand;">
		<td width="7" ></td>
		<td class="SystemLefttitle">留言信息管理</td>
		<td width="7" ></td>
	  </tr>
	</table>
	</div>
	
	<div id="sub9" class="SystemLeftlist" style="display:none">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td width="7" ></td>
		<td class="list"><li><a href="MessageList.asp" target="mainFrame" onClick='changeAdminFlag("留言信息列表")'>留言列表</a></li></td>
		<td width="7" ></td>
	  </tr>
	  <tr>
		<td width="7" ></td>
		<td class="list"><li><a href="MessageEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("添加留言信息")'>添加留言</a></li></td>
		<td width="7" ></td>
	  </tr>
	  <tr>
	  	<td colspan="3" height="7"></td>
	  </tr>
	</table><img src="Images/listbottom.jpg"></div>
	
	
	<div id="main10" class="paddiv" onclick=expandIt(10)>
	<table width="193" height="24" border="0" cellpadding="0" cellspacing="0">
	  <tr style="cursor: hand;">
		<td width="7" ></td>
		<td class="SystemLefttitle">用户管理</td>
		<td width="7" ></td>
	  </tr>
	</table>
	</div>
	
	<div id="sub10" class="SystemLeftlist" style="display:none">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	   <tr>
		<td width="7" ></td>
		<td class="list"><li><a href="AdminList.asp" target="mainFrame" onClick='changeAdminFlag("后台管理员")'>后台管理员</a></li></td>
		<td width="7" ></td>
	  </tr>
      <tr>
		<td width="7" ></td>
		<td class="list"><li><a href="MemList.asp" target="mainFrame" onClick='changeAdminFlag("前台会员")'>前台会员</a></li></td>
		<td width="7" ></td>
	  </tr>
      <tr>
		<td width="7" ></td>
		<td class="list"><li><a href="MemGroupList.Asp" target="mainFrame" onClick='changeAdminFlag("会员组别")'>会员组别</a></li></td>
		<td width="7" ></td>
	  </tr>
	  <tr>
	  	<td colspan="3" height="7"></td>
	  </tr>
	</table><img src="Images/listbottom.jpg"></div>
	
	<div style="margin:10px 0px;"></div>
</div></BODY></HTML>
<script>
function Creathtml()
{
var bln=confirm("注意：添加、修改、删除相关数据时会自动生成、更新、删除所生成的静态文件。\n如果您没有对模板作过修改，不需要批量生成所有商品或新闻详细页面！\n如果您仅对商品/作品、新闻、下载、信息中心等分类页面作过修改，只需要生成相关分类页面。\n\n请确定是否操作？");
return bln;
}
</script>