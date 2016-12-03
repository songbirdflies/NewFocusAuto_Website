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
<TITLE>编辑新闻</TITLE>
<link rel="stylesheet" href="Css/Content.css">
<script language="javascript" src="Js/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="Config.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|33,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'判断是否具有管理权限
%>
<script language = "JavaScript">
//列表添加
	function AddItem(ObjName, DesName)
	{
	  //GET OBJECT ID AND DESTINATION OBJECT
	  ObjID    = GetObjID(ObjName);
	  DesObjID = GetObjID(DesName);
	//  window.alert(document.form.elements[DesObjID].length);
	  k=0;
	  i = document.editForm.elements[ObjID].options.length;
	  if (i==0)
		return;
	  maxselected=0
	  for (h=0; h<i; h++)
		 if (document.editForm.elements[ObjID].options[h].selected ) {
			 k=k+1;
			 maxselected=h+1;
			 }
	  if (maxselected>=i)
		 maxselected=0;
	//  if ( document.form.elements[DesObjID].length + k >4 ) {
	//	window.alert("最多可选择4条！");
	//	return;
	//	}
	
	  if ( ObjID != -1 && DesObjID != -1  )
	  {
		i = document.editForm.elements[ObjID].options.length;
		j = document.editForm.elements[DesObjID].options.length;
		for (h=0; h<i; h++)
		{ if (document.editForm.elements[ObjID].options[h].selected )
		  {  Code = document.editForm.elements[ObjID].options[h].value;
			 Text = document.editForm.elements[ObjID].options[h].text;
			 j = document.editForm.elements[DesObjID].options.length;
			 if (Text.indexOf('－－请选择－－')!=-1) {
				continue;
			 }
			 HasSelected = false;
			 for (k=0; k<j; k++ ) {
			   if (document.editForm.elements[DesObjID].options[k].value == Code)
			   {  HasSelected = true;
				  break;
			   }
			 }
		 //alert("HasSelected="+HasSelected);
			if ( HasSelected == false)
			{
				document.editForm.elements[DesObjID].options[j] = new Option(Text, Code);
			}//if
			document.editForm.elements[ObjID].options[h].selected =false;
		   }//if
		}//for
		document.editForm.elements[ObjID].options[maxselected].selected =true;
	  }//if
	}//end of function


//	删除
function DeleteItem(ObjName)
{
  ObjID = GetObjID(ObjName);
  minselected=0;
  if ( ObjID != -1 )
  {
    for (i=window.editForm.elements[ObjID].length-1; i>=0; i--)
    {  if (window.editForm.elements[ObjID].options[i].selected)
       { // window.alert(i);
          if (minselected==0 || i<minselected)
            minselected=i;
          window.editForm.elements[ObjID].options[i] = null;
       }
    }
    i=window.editForm.elements[ObjID].length;

    if (i>0)  {
        if (minselected>=i)
           minselected=i-1;
        window.editForm.elements[ObjID].options[minselected].selected=true;
        }
  }
}

//列表添加	删除___附属
function GetObjID(ObjName)
{
  for (var ObjID=0; ObjID < window.editForm.elements.length; ObjID++)
    if ( window.editForm.elements[ObjID].name == ObjName )
    {  return(ObjID);
       break;
    }
  return(-1);
}

//参数比较
	function gotoCSCompare()
	{
		var objCarId = document.editForm.SortClass;  		
		SelectCarIds();
		//alert(objCarId.value);
		//document.form.submit();
	}

//参数比较————负
function SelectCarIds()
{
  //alert("SelectCarIds");
  ObjID = GetObjID("SortClass");
  if (ObjID != -1)
  { for (i=0; i<document.editForm.elements[ObjID].length; i++)
      document.editForm.elements[ObjID].options[i].selected = true;
  }
  //return false;
}
//-->	



//列表添加	删除___附属
function GetObjID(ObjName)
{
  for (var ObjID=0; ObjID < window.editForm.elements.length; ObjID++)
    if ( window.editForm.elements[ObjID].name == ObjName )
    {  return(ObjID);
       break;
    }
  return(-1);
}
function selectall(){
	var i;
	var j;
	j=document.editForm.SortClass.options.length;
	for(i=0; i<j;i++){
		document.editForm.SortClass.options[i].selected=true;
	}
	return true;
}
</script>
<BODY>
<table id="bodycontent" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">

<% 
dim Result
Result=request.QueryString("Result")
dim i,ID,TitleNameLangI,TitleNameLangD,TitleNameLangE,ViewFlagLangI,ViewFlagLangD,ViewFlagLangE,HtmlUrlNameLangI,HtmlUrlNameLangD,HtmlUrlNameLangE,SearchKeywordsLangI,SearchKeywordsLangD,SearchKeywordsLangE,SortNameLangI,SeoKeywordsLangI,SeoKeywordsLangD,SeoKeywordsLangE,SeoDescriptionLangI,SeoDescriptionLangD,SeoDescriptionLangE,SortID,SortPath,SourceLangI,SourceLangD,SourceLangE,OrderNum
dim NewFlag,TjFlag,HotFlag
dim SmallPic,BigPic,GotoUrl,ConciseLangI,ConciseLangD,ConciseLangE,ContentLangI,ContentLangD,ContentLangE,AddTime,SortClassID,SortClass
dim GroupID,GroupIdName,Exclusive
ID=request.QueryString("ID")
call NewsEdit() 
%>
<table cellpadding="0" cellspacing="0">
  <tr>
    <th>新闻检索及分类查看：添加，修改，删除新闻信息</th>
  </tr>
  <tr>
    <td height="24" align="center">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>新闻信息】</td>
  </tr>
</table>
<br>
<table cellpadding="0" cellspacing="0">
  <form name="editForm" method="post" onSubmit="return selectall();" action="NewsEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24">
	<table class="noborder" cellpadding="0" cellspacing="0">
      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr class="LangShowI">
        <td height="20" align="right"><%=LangWebI%>名称：</td>
        <td><input name="TitleNameLangI" type="text" class="textfield" id="TitleNameLangI" style="width:240px;" value="<%=TitleNameLangI%>" maxlength="100">&nbsp;显示：<input name="ViewFlagLangI" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlagLangI then response.write ("checked")%>>&nbsp;不少于1个字符</td>
      </tr>
      <tr class="LangShowD">
        <td height="20" align="right"><%=LangWebD%>名称：</td>
        <td><input name="TitleNameLangD" type="text" class="textfield" id="TitleNameLangD" style="width:240px;" value="<%=TitleNameLangD%>" maxlength="100">&nbsp;显示：<input name="ViewFlagLangD" type="checkbox" value="1" style='HEIGHT: 13px;WIDTH: 13px;' <%if ViewFlagLangD then response.write ("checked")%>></td>
      </tr>
      <tr class="LangShowE">
        <td height="20" align="right"><%=LangWebE%>名称：</td>
        <td><input name="TitleNameLangE" type="text" class="textfield" id="TitleNameLangE" style="width:240px;" value="<%=TitleNameLangE%>" maxlength="100">&nbsp;显示：<input name="ViewFlagLangE" type="checkbox" value="1" style='HEIGHT: 13px;WIDTH: 13px;' <%if ViewFlagLangE then response.write ("checked")%>></td>
      </tr>
      <tr class="LangShowI">
		<td height="20" align="right"><%=LangWebI%>静态文件名：</td>
		<td><input name="HtmlUrlNameLangI" type="text" class="textfield" id="HtmlUrlNameLangI" style="WIDTH: 240px;" value="<%=HtmlUrlNameLangI%>" maxlength="100">&nbsp;不填则由系统自动按<%=LangWebI%>名称生成(建议)，不可有重复，一旦填好不可更改</td>
	  </tr>
	  <tr class="LangShowD">
		<td height="20" align="right"><%=LangWebD%>静态文件名：</td>
		<td><input name="HtmlUrlNameLangD" type="text" class="textfield" id="HtmlUrlNameLangD" style="WIDTH: 240px;" value="<%=HtmlUrlNameLangD%>" maxlength="100">&nbsp;不填则由系统自动按<%=LangWebD%>名称生成(建议)，不可有重复，一旦填好不可更改</td>
	  </tr>
	  <tr class="LangShowE">
		<td height="20" align="right"><%=LangWebE%>静态文件名：</td>
		<td><input name="HtmlUrlNameLangE" type="text" class="textfield" id="HtmlUrlNameLangE" style="WIDTH: 240px;" value="<%=HtmlUrlNameLangE%>" maxlength="100">&nbsp;不填则由系统自动按<%=LangWebE%>名称生成(建议)，不可有重复，一旦填好不可更改</td>
	  </tr>
      <tr class="LangShowI">
      <td height="20" align="right"><%=LangWebI%>搜索匹配关键字：</td>
      <td><textarea name="SearchKeywordsLangI" rows="1" class="textfield" id="SearchKeywordsLangI" style="width:500px;"><%=SearchKeywordsLangI%></textarea>
      &nbsp;如：/a/b/</td>
    </tr>
    <tr class="LangShowD">
      <td height="20" align="right"><%=LangWebD%>搜索匹配关键字：</td>
      <td><textarea name="SearchKeywordsLangD" rows="1" class="textfield" id="SearchKeywordsLangD" style="width:500px;"><%=SearchKeywordsLangD%></textarea>
      &nbsp;如：/a/b/</td>
    </tr>
    <tr class="LangShowE">
      <td height="20" align="right"><%=LangWebE%>搜索匹配关键字：</td>
      <td><textarea name="SearchKeywordsLangE" rows="1" class="textfield" id="SearchKeywordsLangE" style="width:500px;"><%=SearchKeywordsLangE%></textarea>
      &nbsp;如：/a/b/</td>
    </tr>
	  <tr class="LangShowI">
      <td height="20" align="right"><%=LangWebI%>关键字：</td>
      <td><textarea name="SeoKeywordsLangI" rows="1" class="textfield" id="SeoKeywordsLangI" style="width:500px;"><%=SeoKeywordsLangI%></textarea>&nbsp;*小于等于 80 个字符</td>
    </tr>
    <tr class="LangShowI">
      <td height="20" align="right"><%=LangWebI%>网站描述：</td>
      <td><textarea name="SeoDescriptionLangI" rows="1" class="textfield" id="SeoDescriptionLangI" style="width:500px;"><%=SeoDescriptionLangI%></textarea>&nbsp;*小于等于 200 个字符</td>
    </tr>
	<tr class="LangShowD">
      <td height="20" align="right"><%=LangWebD%>关键字：</td>
      <td><textarea name="SeoKeywordsLangD" rows="1" class="textfield" id="SeoKeywordsLangD" style="width:500px;"><%=SeoKeywordsLangD%></textarea>&nbsp;</td>
    </tr>
    <tr class="LangShowD">
      <td height="20" align="right"><%=LangWebD%>网站描述：</td>
      <td><textarea name="SeoDescriptionLangD" rows="1" class="textfield" id="SeoDescriptionLangD" style="width:500px;"><%=SeoDescriptionLangD%></textarea>&nbsp;</td>
    </tr>
	<tr class="LangShowE">
      <td height="20" align="right"><%=LangWebE%>关键字：</td>
      <td><textarea name="SeoKeywordsLangE" rows="1" class="textfield" id="SeoKeywordsLangE" style="width:500px;"><%=SeoKeywordsLangE%></textarea>&nbsp;</td>
    </tr>
    <tr class="LangShowE">
      <td height="20" align="right"><%=LangWebE%>网站描述：</td>
      <td><textarea name="SeoDescriptionLangE" rows="1" class="textfield" id="SeoDescriptionLangE" style="width:500px;"><%=SeoDescriptionLangE%></textarea>&nbsp;</td>
    </tr>
      <tr>
        <td height="20" align="right">所属主类别：</td>
        <td><input name="SortNameLangI" type="text" class="textfield" id="SortNameLangI" value="<%=SortNameLangI%>" style="width:240px;background-color:#EBF2F9;" readonly>&nbsp;<a href="javaScript:OpenScript('SelectSort.asp?Result=News',500,500,'')"><img src="Images/Select.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td><input name="SortID" type="text" class="textfield" id="SortID" style="WIDTH: 40;background-color:#EBF2F9;" value="<%=SortID%>" readonly><input name="SortPath" type="text" class="textfield" id="SortPath" style="WIDTH: 200;background-color:#EBF2F9;" value="<%=SortPath%>" readonly>&nbsp;*</td>
      </tr>
      <tr>
      	<td height="20" align="right">所属附类别：</td>
      	<td align="left" style="padding:0px; margin:0px;">
		<table border="0" cellpadding="0" cellspacing="0">
		<Tr><td width="240" align="left">
				  <select name="SortClass" size="14" id="SortClass" multiple style="width:240px;background-color:#EBF2F9;" >
							 <%
			dim sa
			If SortClass<>"" Then
		   sa=split(SortClass,",")
			For i = LBound(sa) To UBound(sa)
			 If Trim(sa(i))<>"" Then
			 dim rspt,sqlpt
			    set rspt=Server.CreateObject("ADODB.RECORDSET")
				sqlpt="Select * from IdeaNet_NewsSort where Id="&Trim(sa(i))&""
				rspt.open sqlpt,conn,1,1
				If Not rspt.eof Then
				%>
				<option value="<%=rspt("Id")%>"><%=rspt("SortNameLangI")%></option>
				<%
				End if
			 End if
			Next		  
		  End if%>
				  </select>
		
		</td>
		<td align=center valign="middle">
		<input type="button" id="btnBack"  OnClick="JavaScript:AddItem('Party_all','SortClass')" value="<- 添加" class="button" />
		<br />
		<br />
		 <input type="button" id="btnBack"  OnClick="JavaScript:DeleteItem('SortClass')" value="-> 删除" class="button" />		</td>
		<td>
		<select name="Party_all" size=14 id="Party_all" style="width:240px;background-color:#EBF2F9;">
		<%call printSort(0,"")%>
		</select>
		<%
		sub printSort(id,str)
			dim sql,rs
			sql="select * from IdeaNet_NewsSort where ParentID="&id
			set rs=conn.execute(sql)
			do while not rs.eof
				response.Write("<option value="""&rs("id")&""">"&str&rs("SortNameLangI")&"</option>")
				call printSort(rs("id"),str&"　")
				rs.movenext
			loop
		end sub	
		%>* 请同时选择主类别名称到附类别中
        </td>
		</tr>
		</table></td>
      	</tr>
      <%
dim rs,sql,SiteTitlelangI,SiteTitlelangD,SiteTitlelangE
set rs = server.createobject("adodb.recordset")
sql="select top 1 * from IdeaNet_Site"
rs.open sql,conn,1,1
SiteTitlelangI=rs("SiteTitlelangI")
SiteTitlelangD=rs("SiteTitlelangD")
SiteTitlelangE=rs("SiteTitlelangE")
rs.close
set rs=Nothing
%>
      <tr class="LangShowI">
        <td height="20" align="right"><%=LangWebI%>来源：</td>
        <td><input name="SourceLangI" type="text" class="textfield" id="MakerLangI" style="width: 240;" value="<%=SourceLangI%>" maxlength="100">&nbsp;<a href="javascript:" onClick="document.editForm.SourceLangI.value='<%=SiteTitlelangI%>'"><%=SiteTitlelangI%></a>&nbsp;</td>
      </tr>
      <tr class="LangShowD">
        <td height="20" align="right"><%=LangWebD%>来源：</td>
        <td><input name="SourceLangD" type="text" class="textfield" id="MakerLangD" style="width: 240;" value="<%=SourceLangD%>" maxlength="100">&nbsp;<a href="javascript:" onClick="document.editForm.SourceLangD.value='<%=SiteTitlelangD%>'"><%=SiteTitlelangD%></a>&nbsp;</td>
      </tr>
      <tr class="LangShowE">
        <td height="20" align="right"><%=LangWebE%>来源：</td>
        <td><input name="SourceLangE" type="text" class="textfield" id="MakerLangE" style="width: 240;" value="<%=SourceLangE%>" maxlength="100">&nbsp;<a href="javascript:" onClick="document.editForm.SourceLangE.value='<%=SiteTitlelangE%>'"><%=SiteTitlelangE%></a>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">属性、权限：</td>
        <td>
		<input name="NewFlag" type="checkbox" value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if NewFlag then response.write ("checked")%>>&nbsp;最新
		<input name="TjFlag" type="checkbox" style="HEIGHT: 13px;WIDTH: 13px;" value="1" <%if TjFlag then response.write ("checked")%>>&nbsp;推荐
		<input name="HotFlag" type="checkbox" style="HEIGHT: 13px;WIDTH: 13px;" value="1" <%if HotFlag then response.write ("checked")%>>&nbsp;热门</td>
      </tr>
	  <tr>
        <td height="20" align="right">缩 略 图：</td>
        <td><input name="SmallPic" type="text" class="textfield" style="width:240px;" value="<%=SmallPic%>" maxlength="100">&nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=SmallPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
      <tr>
        <td height="20" align="right">主　　图：</td>
        <td><input name="BigPic" type="text" class="textfield" style="width:240px;" value="<%=BigPic%>" maxlength="100">&nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=BigPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
	  <tr>
        <td height="20" align="right">链接网址：</td>
        <td><input name="GotoUrl" type="text" class="textfield" id="GotoUrl" style="width:240px;" value="<%=GotoUrl%>" maxlength="100"></td>
	  </tr>
      <tr>
        <td height="20" align="right">查看权限：</td>
        <td><select name="GroupID" class="textfield">
          <% call SelectGroup() %>
          </select>
          <input name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>> 隶属<input type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">专属（隶属：权限值≥可查看，专属：权限值＝可查看）</td>
      </tr>
	  <tr class="LangShowI">
        <td height="20" align="right"><%=LangWebI%>简要概述：</td>
        <td><textarea name="ConciseLangI" rows="2" class="textfield" id="ConciseLangI" style="WIDTH: 86%;"><%=ConciseLangI%></textarea>&nbsp;</td>
	  </tr>
	  <tr class="LangShowD">
        <td height="20" align="right"><%=LangWebD%>简要概述：</td>
        <td><textarea name="ConciseLangD" rows="2" class="textfield" id="ConciseLangD" style="WIDTH: 86%;"><%=ConciseLangD%></textarea>&nbsp;</td>
	  </tr>
	  <tr class="LangShowE">
        <td height="20" align="right"><%=LangWebE%>简要概述：</td>
        <td><textarea name="ConciseLangE" rows="2" class="textfield" id="ConciseLangE" style="WIDTH: 86%;"><%=ConciseLangE%></textarea>&nbsp;</td>
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
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;" ></td>
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
sub NewsEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑新闻信息
    set rs = server.createobject("adodb.recordset")
    if request.Form("ViewFlagLangI")=1 then
	  if len(trim(request.Form("TitleNameLangI")))<1 then
	  response.write ("<script language=javascript> alert('你选择了"""&LangWebI&"""显示，因此"&LangWebI&"名称为必填！');history.back(-1);</script>")
	  response.end
      end if
	end if
    if request.Form("ViewFlagLangD")=1 then
      if len(trim(request.Form("TitleNameLangD")))<1 then
      response.write ("<script language=javascript> alert('您已经选择了"""&LangWebD&"""显示，因此"&LangWebD&"名称必填！');history.back(-1);</script>")
      response.end
      end if
    end if
    if request.Form("ViewFlagLangE")=1 then
      if len(trim(request.Form("TitleNameLangE")))<1 then
      response.write ("<script language=javascript> alert('您已经选择了"""&LangWebE&"""显示，因此"&LangWebE&"名称必填！');history.back(-1);</script>")
      response.end
      end if
    end if
    if Result="Add" then '创建新闻信息
	  sql="select * from IdeaNet_News"
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
	  rs("HtmlUrlNameLangI")=trim(Request.Form("HtmlUrlNameLangI"))
      rs("HtmlUrlNameLangD")=trim(Request.Form("HtmlUrlNameLangD"))
      rs("HtmlUrlNameLangE")=trim(Request.Form("HtmlUrlNameLangE"))
	  rs("SearchKeywordsLangI")=trim(Request.Form("SearchKeywordsLangI"))
	  rs("SearchKeywordsLangD")=trim(Request.Form("SearchKeywordsLangD"))
	  rs("SearchKeywordsLangE")=trim(Request.Form("SearchKeywordsLangE"))
	  rs("SeoKeywordsLangI")=trim(Request.Form("SeoKeywordsLangI"))
	  rs("SeoDescriptionLangI")=trim(Request.Form("SeoDescriptionLangI"))
	  rs("SeoKeywordsLangD")=trim(Request.Form("SeoKeywordsLangD"))
	  rs("SeoDescriptionLangD")=trim(Request.Form("SeoDescriptionLangD"))
	  rs("SeoKeywordsLangE")=trim(Request.Form("SeoKeywordsLangE"))
	  rs("SeoDescriptionLangE")=trim(Request.Form("SeoDescriptionLangE"))
	  if Request.Form("SortID")="" and Request.Form("SortPath")="" then
        response.write ("<script language=javascript> alert('请选择所属分类！');history.back(-1);</script>")
        response.end
	  else
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  end if
	  rs("SourceLangI")=trim(Request.Form("SourceLangI"))
	  rs("SourceLangD")=trim(Request.Form("SourceLangD"))
	  rs("SourceLangE")=trim(Request.Form("SourceLangE"))
	  if Request.Form("NewFlag")=1 then
        rs("NewFlag")=Request.Form("NewFlag")
	  else
        rs("NewFlag")=0
	  end if
	  if Request.Form("TjFlag")=1 then
        rs("TjFlag")=Request.Form("TjFlag")
	  else
        rs("TjFlag")=0
	  end if
	  if Request.Form("HotFlag")=1 then
        rs("HotFlag")=Request.Form("HotFlag")
	  else
        rs("HotFlag")=0
	  end if
	  rs("SmallPic")=trim(Request.Form("SmallPic"))
	  rs("BigPic")=trim(Request.Form("BigPic"))
	  rs("GotoUrl")=trim(Request.Form("GotoUrl"))
	  GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("ConciseLangI")=Request.Form("ConciseLangI")
	  rs("ConciseLangD")=Request.Form("ConciseLangD")
	  rs("ConciseLangE")=Request.Form("ConciseLangE")
	  rs("ContentLangI")=Request.Form("ContentLangI")
	  rs("ContentLangD")=Request.Form("ContentLangD")
	  rs("ContentLangE")=Request.Form("ContentLangE")
	  rs("OrderNum")=9999
	  if Request.Form("AddTime")<>"" then
        rs("AddTime")=Request.Form("AddTime")
	  else
        rs("AddTime")=now()
	  end if
	end if  
	
	if Result="Modify" then '修改新闻信息
      sql="select * from IdeaNet_News where ID="&ID
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
	  rs("HtmlUrlNameLangI")=trim(Request.Form("HtmlUrlNameLangI"))
      rs("HtmlUrlNameLangD")=trim(Request.Form("HtmlUrlNameLangD"))
      rs("HtmlUrlNameLangE")=trim(Request.Form("HtmlUrlNameLangE"))
	  rs("SearchKeywordsLangI")=trim(Request.Form("SearchKeywordsLangI"))
	  rs("SearchKeywordsLangD")=trim(Request.Form("SearchKeywordsLangD"))
	  rs("SearchKeywordsLangE")=trim(Request.Form("SearchKeywordsLangE"))
	  rs("SeoKeywordsLangI")=trim(Request.Form("SeoKeywordsLangI"))
	  rs("SeoDescriptionLangI")=trim(Request.Form("SeoDescriptionLangI"))
	  rs("SeoKeywordsLangD")=trim(Request.Form("SeoKeywordsLangD"))
	  rs("SeoDescriptionLangD")=trim(Request.Form("SeoDescriptionLangD"))
	  rs("SeoKeywordsLangE")=trim(Request.Form("SeoKeywordsLangE"))
	  rs("SeoDescriptionLangE")=trim(Request.Form("SeoDescriptionLangE"))
	  if Request.Form("SortID")<>"" and Request.Form("SortPath")<>"" then
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  else
        response.write ("<script language=javascript> alert('请选择所属分类！');history.back(-1);</script>")
        response.end
	  end if	
	  rs("SourceLangI")=trim(Request.Form("SourceLangI"))
	  rs("SourceLangD")=trim(Request.Form("SourceLangD"))
	  rs("SourceLangE")=trim(Request.Form("SourceLangE"))
	  if Request.Form("NewFlag")=1 then
        rs("NewFlag")=Request.Form("NewFlag")
	  else
        rs("NewFlag")=0
	  end if
	  if Request.Form("TjFlag")=1 then
        rs("TjFlag")=Request.Form("TjFlag")
	  else
        rs("TjFlag")=0
	  end if
	  if Request.Form("HotFlag")=1 then
        rs("HotFlag")=Request.Form("HotFlag")
	  else
        rs("HotFlag")=0
	  end if
	  rs("SmallPic")=trim(Request.Form("SmallPic"))
	  rs("BigPic")=trim(Request.Form("BigPic"))
	  rs("GotoUrl")=trim(Request.Form("GotoUrl"))
	  GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("ConciseLangI")=Request.Form("ConciseLangI")
	  rs("ConciseLangD")=Request.Form("ConciseLangD")
	  rs("ConciseLangE")=Request.Form("ConciseLangE")
	  rs("ContentLangI")=Request.Form("ContentLangI")
	  rs("ContentLangD")=Request.Form("ContentLangD")
	  rs("ContentLangE")=Request.Form("ContentLangE")
	  if Request.Form("AddTime")<>"" then
        rs("AddTime")=Request.Form("AddTime")
	  else
        rs("AddTime")=now()
	  end if
	end if
	  if Request.Form("SortClass")<>"" then
	    sa=split(Request("SortClass"),",") 
		rs("SortClassID")=sa(0)
	    rs("SortClass")=","&replace(request("SortClass")," ","")&","  '多个分类
	  else
        response.write ("<script language=javascript> alert('对不起,请至少选择一个分类！');history.back(-1);</script>")
        response.end
	  end if
	rs.update
	rs.close
    set rs=nothing
    response.write "<script language=javascript> alert('成功编辑新闻信息！');location.replace('NewsList.asp');</script>"
  else '提取新闻信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from IdeaNet_News where ID="& ID
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
	  HtmlUrlNameLangI=rs("HtmlUrlNameLangI")
	  HtmlUrlNameLangD=rs("HtmlUrlNameLangD")
	  HtmlUrlNameLangE=rs("HtmlUrlNameLangE")
	  SearchKeywordsLangI=rs("SearchKeywordsLangI")
	  SearchKeywordsLangD=rs("SearchKeywordsLangD")
	  SearchKeywordsLangE=rs("SearchKeywordsLangE")
	  SeoKeywordsLangI=rs("SeoKeywordsLangI")
	  SeoDescriptionLangI=rs("SeoDescriptionLangI")
	  SeoKeywordsLangD=rs("SeoKeywordsLangD")
	  SeoDescriptionLangD=rs("SeoDescriptionLangD")
	  SeoKeywordsLangE=rs("SeoKeywordsLangE")
	  SeoDescriptionLangE=rs("SeoDescriptionLangE")
	  SortNameLangI=SortText(rs("SortID"))
	  SortID=rs("SortID")
	  SortPath=rs("SortPath")
	  SourceLangI=rs("SourceLangI")
	  SourceLangD=rs("SourceLangD")
	  SourceLangE=rs("SourceLangE")
	  NewFlag=rs("NewFlag")
	  TjFlag=rs("TjFlag")
	  HotFlag=rs("HotFlag")
	  SmallPic=rs("SmallPic")
	  BigPic=rs("BigPic")
	  GotoUrl=rs("GotoUrl")
	  GroupID=rs("GroupID")
	  Exclusive=rs("Exclusive")
	  ConciseLangI=rs("ConciseLangI")
      ConciseLangD=rs("ConciseLangD")
      ConciseLangE=rs("ConciseLangE")
      ContentLangI=rs("ContentLangI")
      ContentLangD=rs("ContentLangD")
      ContentLangE=rs("ContentLangE")
	  AddTime=rs("AddTime")
	  SortClassID=rs("SortClassID")
	  SortClass=rs("SortClass")
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
  sql="Select * From IdeaNet_NewsSort where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortNameLangI")
  rs.close
  set rs=nothing
End Function
%>

<% 
sub SelectGroup()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from IdeaNet_MemGroup"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write("未设组别")
  end if
  while not rs.eof
    response.write("<option value='"&rs("GroupID")&"┎╂┚"&rs("GroupName")&"'")
    if GroupID=rs("GroupID") then response.write ("selected")
    response.write(">"&rs("GroupName")&"</option>")
    rs.movenext
  wend
  rs.close
  set rs=nothing
end sub
%>