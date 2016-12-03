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
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Css/Content.css">
<title>文件选择</title>
</head> 
<!--#include file="CheckAdmin.asp"-->
<body>
<table id="bodycontent" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
	<table class="noborder" width="400" align="center" cellpadding="12" cellspacing="1" >
	<form action="UpFileSave.asp?Result=<%=request.QueryString("Result")%>" method="post" enctype="multipart/form-data" name="formUpload">
	<tr>
    <td>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="60" height="30" nowrap>选择文件：</td>
        <td><input name="FromFile" type="file" class="textfield" id="FromFile" size="41"></td>
      </tr>
      <tr>
        <td height="30">上传位置：</td>
        <td>
		<select name="SaveToPath" class="textfield">
		  <option value="../Upload/SysFiles/" selected>系统文件 Upload/SysFiles</option>
		  <option value="../Upload/AboutFiles/">企业文件 Upload/AboutFiles</option>
		  <option value="../Upload/NewsFiles/">新闻文件 Upload/NewsFiles</option>
		  <option value="../Upload/PicFiles/">商品/作品文件 Upload/PicFiles</option>
          <option value="../Upload/DownLoadFiles/">下载文件 Upload/DownLoadFiles</option>
          <option value="../Upload/InfoFiles/">信息文件 Upload/InfoFiles</option>
          <option value="../Upload/ContentFiles/">单页面文件 Upload/ContentFiles</option>
		  <option value="../Upload/AdsFiles/">广告文件 Upload/AdsFiles</option>
        </select>
	    </td>
      </tr>
      <tr>
	  <td height="36" colspan="2" align="center" valign="bottom">
	  <input name="reset" type="reset" class="button" value=" 重置 ">
	  <input name="Submit" type="submit" class="button" value=" 上传 ">
	  </td>
	  </tr>
    </table>
	</td>
  </tr>
  </form>
</table>
</td></tr></table>
</body>
</html>

