<!--#include file="Top.Asp"-->
<%SeoTitle="Home"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=Keywords%>,<%=SiteTitle%> - <%=SeoTitle%></title>
<meta name="keywords" content="<%=SeoTitle%>,<%=Keywords%>,<%=SiteTitle%>" />
<meta name="description" content="<%=SeoTitle%>,<%=Descriptions%>" />
<meta name="generator" content="<%=SiteTitle%>" />
<meta name="author" content="<%=SiteUrl%>" />
<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.3)" />
<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.3)" />
<link rel="stylesheet" type="text/css" href="Css/StyleIndex.Css" />
<script type="text/javascript" src="Js/Style.js"></script>
<script type="text/javascript" src="Js/JqueryLR.js"></script>
<script type="text/javascript" src="Js/LRscroll.js"></script>
<script type="text/javascript" src="Js/Marquee.js"></script>
<script type="text/javascript" src="Js/JQh_jquery.js"></script>
<script type="text/javascript" src="Js/JQh_commonv1.1.js"></script>
<script type="text/javascript" src="Js/JQh_commonv1.2.js"></script>
<!--[if lt IE 7]>
<script defer type="text/javascript" src="Js/pngfix_inline.js"></script>
<script defer type="text/javascript" src="Js/pngfix_bg.js"></script>
<![endif]-->

<script type="text/javascript" src="javascript/jquery.js"></script>
    <script type="text/javascript" src="javascript/jquery.easing.1.3.min.js"></script>
    <script type="text/javascript" src="javascript/jquery.wt-rotator.min.js"></script>
   <script type="text/javascript" src="javascript/jMyCarousel.js"></script>


</head>
<body>
<table height="99" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="275"><a href="http://<%=SiteUrl%>"><img src="<%=FilePath%><%=SystemPic%>" border="0"></a></td>
    <%set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.open "Select * from IdeaNet_NavigationMenu where ViewFlag"&LangData&" and ParentId=0 order by OrderNum asc,ID asc",conn,1,1
    ii=1
    do while not Rs.eof%>
    <%if ii=7 then%>
    <%if Rs("SmallPic")<>"" then%>
    <td width="210">
    <div id="menu_nav"><a href="<%=Rs("GotoUrl")%>"><img src="<%=FilePath%><%=Rs("SmallPic")%>" onMouseOver="src='<%=FilePath%><%=Rs("bigPic")%>';" onmouseleave="src='<%=FilePath%><%=Rs("SmallPic")%>';" border="0"></a></div></td>
    <%else%>
    <td width="210"><div id="menu_nav"><a href="<%=Rs("GotoUrl")%>"><%=Rs("SortName"&LangData)%></a></div></td>
    <%end if%>
    <%else%>
    <%if Rs("SmallPic")<>"" then%>

    <td width="100">
    <div id="menu_nav"><a href="<%=Rs("GotoUrl")%>"><img src="<%=FilePath%><%=Rs("SmallPic")%>" onMouseOver="src='<%=FilePath%><%=Rs("bigPic")%>';" onmouseleave="src='<%=FilePath%><%=Rs("SmallPic")%>';" border="0"></a></div></td>
    <td width="2" class="paddingtop20"><img src="Images/menu03.gif" /></td>
    <%else%>
    <td width="100"><div id="menu_nav"><a href="<%=Rs("GotoUrl")%>"><%=Rs("SortName"&LangData)%></a></div></td>
    <td width="2" class="paddingtop20"><img src="Images/menu03.gif" /></td>
    <%end if%>
    <%end if%>
    <%Rs.movenext
	ii=ii+1
    loop
    Rs.close
    set Rs=nothing%>

		<td width="150"><div id="menu_bar"><p></p>&nbsp;&nbsp;&nbsp;<img src="Images/facebook.jpg" width="25"> &nbsp;<img src="Images/twitter.jpg" width="25"> 
		&nbsp;<a href="http://e.weibo.com/u/2948823690" target="_blank"><img src="Images/sina.jpg" width="25"></a>
		&nbsp;<a href="http://t.qq.com/nfa2012" target="_blank"><img src="Images/logo_tencent.png" width="25"></a>
	</div></td>
	<td width="2" class="paddingtop20"><img src="Images/menu03.gif" /></td>
        <td>&nbsp;</td>
  </tr>
</table>

<div class="container">
<div class="wt-rotator">
    	<div class="screen">
           <script type="text/javascript">
    	$(document).ready(	
			function() {
				$(".container").wtRotator({
					width:1140,
					height:255,
					thumb_width:24,
            		thumb_height:24,
					button_width:24,
					button_height:24,
					button_margin:5,
					auto_start:true,
					delay:4000,
					play_once:false,
					transition:"fade",
					transition_speed:1000,
					auto_center:true,
					easing:"",
					cpanel_position:"inside",
					cpanel_align:"BR",
					timer_align:"top",
					display_thumbs:false,
					display_dbuttons:false,
					display_playbutton:false,
					display_thumbimg:false,
           			display_side_buttons:false,
					display_numbers:false,
					display_timer:false,
					mouseover_pause:false,
					cpanel_mouseover:false,
					text_mouseover:false,
					text_effect:"fade",
					text_sync:true,
					tooltip_type:"image",
					lock_tooltip:true,
					shuffle:false,
					block_size:75,
					vert_size:55,
					horz_size:50,
					block_delay:25,
					vstripe_delay:75,
					hstripe_delay:180			
				});
			}
		);
    </script>   
 <noscript>
            	<!-- placeholder 1st image when javascript is off -->
                <img src="Images/banner_snow.png"/>
            </noscript>
      	</div>
        <div class="c-panel">
      		<div class="thumbnails">
                <ul>
                  
					<li>
					    <a href="Images/banner_inverter.png"><img src="images/thumbs/banner_inverter.jpg"/></a>	
                                         	                           
                    </li>
					<li>
                        <a href="Images/banner_charger.png"><img src="images/thumbs/banner_charger.jpg"/></a>	
                                         	                           
                    </li>  
					     	<li>
                        <a href="Images/banner_light.png"><img src="images/thumbs/banner_light.jpg"/></a>
                       	
                            
                                         	                           
                    </li> 
					  <li>
                        <a href="Images/banner_jump.png"><img src="images/thumbs/banner_jump.jpg"/></a>
                       	
					<li>
                    	<a href="Images/banner_snow.png"><img src="images/thumbs/banner-snow.jpg"/></a>
                                              
                    </li>
                    <li>
                        <a href="Images/indoorsoutdoors.png"><img src="images/thumbs/indoorsoutdoors.jpg"/></a>
                                             

                    </li>

                            
                                         	                           
                    </li>    


              	</ul>
          	</div>  <!-- thumbnails -->   
		</div> <!-- c panel -->
	</div> <!-- wt rotator -->
</div> <!-- .container -->



<div style="background:url(Images/topbg_01_02.gif)">
<table height="55" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="960" align="right"><a href="Cn/"><img src="Images/cn.gif" width="69" height="24" /></a></td>
  </tr>
</table>
</div>

<table width="1100" border="0" align="center" cellpadding="0" cellspacing="0" class="margintop10 marginbottom10">
  <tr>
    <td valign="top" class="paddingtop10" align="center" cellpadding="5"><!--#include file="Ad3.Asp"--></td>
    <td width="339" height="150" valign="top">
    <table width="329" border="0" cellspacing="0" cellpadding="0" class="margintop10">
      <tr>
        <td height="32" background="Images/t_01_01.gif" class="paddingleft15 c2 bold">Company Introduction</td>
      </tr>
      <tr>
        <td class="padding5"><%call Content("IdeaNet_About","","ViewFlag"&LangData&"","ViewFlag"&LangData&" and Id=1","OrderNum,ID","Concise","")%><p>
		
		</td>
      </tr>
    </table></td>
    <td width="200" valign="top"><a href="http://www.newfocusauto.com/News.Asp"><img src="Images/michelin.jpg" height="135" class="marginleft10 margintop10" alt="News"></ br><p><b><center>NFA Signs Contract with Michelin</center></b></a></td>
  </tr>
</table>
<!--#include file="Foot.Asp"-->