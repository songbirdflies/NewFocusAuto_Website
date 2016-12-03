<%
'得到时间
function getShortDate(datestr)
Dim yearstr,daystr,monthstr
yearstr=Year(datestr)
daystr=Day(datestr)
if len(daystr)<2 then daystr="0"&daystr
monthstr=Month(datestr)
if len(monthstr)<2 then monthstr="0"&monthstr
getShortDate=""&yearstr&"-"&monthstr&"-"&daystr&""
end function


'str输出字符串过滤，
function FormatStr(str)
	str=replace(str,"'","‘")
	str=replace(str,"<","&lt;")
	str=replace(str,">","&gt;")
	str=replace(str,"""","“")
	if LangData="LangI" then str=replace(str," ","&nbsp;")
	str=replace(str,chr(13),"<br>")
	FormatStr=str
end function

'补长度
function adleng(str,lengthn)
Dim i
if len(str)<lengthn then
	for i=1 to lengthn-len(str)
		str=str&"　"
	next
else
	str=left(str,lengthn)
end if
adleng=str
end function


' @desc    截取数组字符串
' @param   String  str     要截取的字符串
' @param   String  dstr    分隔符
' @param   Integer len     要截取的数组元素的个数
' @param   Integer start   要截取的数组元素的开始位置
' @param   Boolean endwith 是否在末尾加上分隔符
' @return  String
Function cutArrStr(str, dstr, len, start, endwith)
	Dim tmp_arr, tmp_len, arr(), i

	tmp_arr = Split(str, dstr)
	tmp_len = UBound(tmp_arr)

	If len > tmp_len Then
		len = tmp_len
	End If
	If len < 1 Then
		len = 1
	End If
	
	Redim arr(len)
	
	
	If start < 0 Then
		i = 0
	End If
	If start > tmp_len Then
		i = tmp_len
	End If
	
	Do While i<len
		arr(i) = tmp_arr(i)
		i = i + 1
	Loop

	str = Join(arr, dstr)

	If endwith Then
		str = str + dstr
	End If

	cutArrStr = str
End Function
'USAGE
'str = "0,10,43,"
'str = cutArrStr(str, ",", 2, 0, false)
'Response.write(str)
'-------------------------------------------------------------------------------------------------------------------------



'-------------------------------------------------------------------------------------------------------------------------
'多级分类显示公共函数
'数据库DataStr; '取值路径ParentID; '排序方式OrderStr; '显示字符长度SortNameLength; '链接地址LinkStr

Sub IdeanetSort(DataStr,ParentID,OrderStr,SortNameLength,LinkStr)
    Dim Rs,Sql,i,ChildCount
    Sql="Select * From "&DataStr&" where ViewFlag"&LangData&" and ParentId="&ParentID&" order by "&OrderStr
    set Rs=conn.execute(Sql)
    response.Write("<table>")
	If ParentId=0 And Rs.recordcount=0 Then
        response.Write ("<tr><td align=""center"">&nbsp;</td></tr>")
        Exit Sub
    End If
    i=1
    While Not Rs.EOF
        ChildCount = conn.Execute("select count(*) from "&DataStr&" where ViewFlag"&LangData&" and ParentId="&Rs("Id"))(0)
        response.Write("	<tr height=""23"">")
        response.Write("		<td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td><img src=""Images/arr2.gif"" border=""0""> <a href="""&LinkStr&"?SortId="&Rs("Id")&"&SortPath="&Rs("SortPath")&""">"&left(Rs("SortName"&LangData),SortNameLength)&"</a></td>")
        response.Write("	</tr>")
        If ChildCount>0 Then%>
<tr>
  <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td><%call IdeanetSort(DataStr,Rs("Id"),OrderStr,SortNameLength,LinkStr)%></td>
</tr>
<%End If
Rs.movenext
i=i+1
Wend
response.Write("</table>")
Rs.Close
Set Rs = Nothing
End Sub
'-------------------------------------------------------------------------------------------------------------------------



'-------------------------------------------------------------------------------------------------------------------------
'单行多列同级分类显示公共函数
'数据库DataStr; '显示前N条; '取值条件WhereStr; '排序方式OrderStr; '分列显示数TdNumstr默认为填1; '显示字符长度SortNameLength; 'Css样式CssStr; '链接地址LinkStr

sub IdeaNetMenu(DataStr,N,WhereStr,OrderStr,TdNumstr,SortNameLength,CssStr,LinkStr)
Dim Rs,Sql,TdNum
if WhereStr="" then WhereStr="ParentId=0"
if OrderStr="" then OrderStr="OrderNum desc,ID"
if TdNumstr="" or TdNumstr=0 then TdNumstr=1 
TdNum=100/TdNumstr
if trim(LinkStr)="" then LinkStr=""
if N="" then
Sql="select * from "&DataStr&" where "&WhereStr&" order by "&OrderStr
else
Sql="select top "&N&" * from "&DataStr&" where "&WhereStr&" order by "&OrderStr
end if
set Rs=conn.execute(Sql)
Response.Write "<table width=""98%"">" & vbCrLf
If Rs.EOF Then
        Response.Write "<tr>"&vbCrLf
		Response.Write "<td align=""center"">&nbsp;</td>"&vbCrLf
	    Response.Write "</tr>"&vbCrLf
Else
While Not Rs.eof
Response.Write "<tr>" & vbCrLf
For i=1 To TdNumstr
If Not Rs.eof then
Response.Write "<td width="""&TdNum&"%"">" & vbCrLf
   Call ShowFormat(ShowStr,LinkStr,TitleNameLength,ConciseLength,CssStr,Rs)
Response.Write "</td>" & vbCrLf
Rs.movenext
else
Response.Write "<td width="""&TdNum&"%"">&nbsp;</td>" & vbCrLf
End If
next 
Response.Write "</tr>" & vbCrLf
Wend
End If
Response.Write "</table>"&vbCrLf
Rs.Close
Set Rs=Nothing
end sub
'-------------------------------------------------------------------------------------------------------------------------



'-------------------------------------------------------------------------------------------------------------------------
'信息名称Seo优化: TitleSeo()
'数据库DataStr; '当前数据ID值IDStr; '动态取值条件WhereStr如：ViewFlag"&LangData&" and Id="&request.QueryString("Id")&"; '取值字段TNameStr如：TitleName或SortName; '默认名称TStr如：关于我们;
sub TitleSeo(DataStr,IDStr,WhereStr,TNameStr,HtmlUrlNameDiy,SitListshowDiy,TStr)
Dim Rs,Sql
if IDStr="" then
SeoTitle=""&TStr&""
TdNumstr=1
PageUrlName=""&HtmlUrlNameDiy&""
SitListshow=""&SitListshowDiy&""
elseif not IsNumeric(IDStr) then
SeoTitle="Error!"
elseif conn.execute("select * from "&DataStr&" Where "&WhereStr&"").eof then
SeoTitle="Error!"
else
set rs = server.createobject("adodb.recordset")
sql="select * from "&DataStr&" where "&WhereStr&""
rs.open sql,conn,1,1
SeoTitle=rs(""&TNameStr&""&LangData)
If rs("HtmlUrlName"&LangData)<>"" then
PageUrlName=rs("HtmlUrlName"&LangData)
else
PageUrlName=""&HtmlUrlNameDiy&""
end if
If TNameStr="SortName" then SitListshow=rs("SitListshow")
If TNameStr="SortName" then TdNumstr=rs("TdNumstr")
rs.close
set rs=nothing
end if
end sub
'-------------------------------------------------------------------------------------------------------------------------



'-------------------------------------------------------------------------------------------------------------------------
'分类名称Seo优化: SortSeo()
'数据库DataStr; '当前数据ID值IDStr; '动态取值条件WhereStr如：ViewFlag"&LangData&" and Id="&request.QueryString("SortID")&"; '取值字段TNameStr如：SortName; '默认名称TStr如：商品/作品展示;
sub SortSeo(DataStr,IDStr,WhereStr,TNameStr,TStr)
Dim Rs,Sql
if IDStr="" then
SeoSort=""&TStr&""
elseif not IsNumeric(IDStr) then
SeoSort="Error!"
elseif conn.execute("select * from "&DataStr&" Where "&WhereStr&"").eof then
SeoSort="Error!"
else
set rs = server.createobject("adodb.recordset")
sql="select * from "&DataStr&" where "&WhereStr&""
rs.open sql,conn,1,1
SeoSort=Rs(""&TNameStr&""&LangData)
rs.close
set rs=nothing
end if
end sub
'-------------------------------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------------------------------
'信息名称Seo优化: TitleSeo()
'数据库DataStr; '分类数据库SortDataStr; '当前数据ID值IDStr; '动态取值条件WhereStr如：ViewFlag"&LangData&" and Id="&request.QueryString("Id")&"; '默认名称TStr如：关于我们; '默认分类名称SStr如：新闻动态
sub TitleSortSeo(DataStr,SortDataStr,IDStr,WhereStr,TStr,SStr)
Dim Rs,Sql
if IDStr="" then
SeoTitle=""&TStr&""
SeoSort=""&SStr&""
elseif not IsNumeric(IDStr) then
SeoTitle="Error!"
SeoSort="Error!"
elseif conn.execute("select * from "&DataStr&" Where "&WhereStr&"").eof then
SeoTitle="Error!"
SeoSort="Error!"
else
set rs = server.createobject("adodb.recordset")
sql="select * from "&DataStr&" where "&WhereStr&""
rs.open sql,conn,1,1
SeoTitle=rs("TitleName"&LangData)
SortId=rs("SortId")
rs.close
set rs=nothing

set rs = server.createobject("adodb.recordset")
sql="select * from "&SortDataStr&" where ViewFlag"&LangData&" and Id="&SortId&""
rs.open sql,conn,1,1
SeoSort=rs("SortName"&LangData)
rs.close
set rs=nothing
end if
end sub
'-------------------------------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------------------------------
'一级分类名称取值: BigSortName()
'数据库SortDataStr; '当前数据SortID值SortIDStr; '动态取值条件WhereStr如：ViewFlag"&LangData&" and Id="&request.QueryString("Id")&"; '默认名称PStr如：ParentID=0; '默认分类名称SStr如：综合信息
sub BigSortName(SortDataStr,SortIDStr,WhereStr,PStr,HtmlUrlNameDiy,SStr)
if SortIDStr="" then
ParentID=""&PStr&""
SortName=""&SStr&""
else
Dim Rs,Sql
set Rs=server.CreateObject("adodb.recordset")
sql="select * from "&SortDataStr&" where "&WhereStr&""
rs.open sql,conn,1,1
SortPath=Rs("SortPath")
rs.close
set rs=nothing

set Rs=server.CreateObject("adodb.recordset")
sql="select * from "&SortDataStr&" where ViewFlag"&LangData&" and SortPath='"&CutArrStr(""&SortPath&"", ",", 2, 0, false)&"'"
rs.open sql,conn,1,1
ParentID=Rs("ID")
SortName=Rs("SortName"&LangData)
If Rs("HtmlUrlName"&LangData)<>"" then
SortHtmlUrlName=Rs("HtmlUrlName"&LangData)
else
SortHtmlUrlName=""&HtmlUrlNameDiy&""
end if
rs.close
set rs=nothing
end if
end sub
'-------------------------------------------------------------------------------------------------------------------------



'-------------------------------------------------------------------------------------------------------------------------
'信息单列表公共函数: ShowTrList()
'数据库DataStr; '显示前n条N; '默认取值条件MrWhereStr; '动态取值条件WhereStr如：ViewFlag"&LangData&" and Instr(SortPath,'"&SortPath&"')>0; '排序方式OrderStr; '第一条显示方式TopShowStr; '显示方式ShowStr; '第一条标题显示字符长度TopTitleNameLength; '第一条简述显示字符长度TopConciseLength; '标题显示字符长度TitleNameLength; '简述显示字符长度ConciseLength; '第一条Css样式TopCssStr; 'Css样式CssStr; '链接地址LinkStr

sub ShowTrList(DataStr,N,MrWhereStr,WhereStr,OrderStr,TopShowStr,ShowStr,TopTitleNameLength,TopConciseLength,TitleNameLength,ConciseLength,TopCssStr,CssStr,LinkStr)
Dim Rs,Sql,TdNum
if WhereStr="" then WhereStr=""&MrWhereStr&""
if OrderStr="" then OrderStr="OrderNum,ID desc"
if TopShowStr="" then TopShowStr="text"
if ShowStr="" then ShowStr="text"
if trim(LinkStr)="" then LinkStr=""
if N="" then
Sql="select * from "&DataStr&" where "&WhereStr&" order by "&OrderStr
else
Sql="select top "&N&" * from "&DataStr&" where "&WhereStr&" order by "&OrderStr
end if
set Rs=conn.Execute(Sql)
If not Rs.EOF Then
i=1
While Not Rs.eof
if i=1 then
   Call ShowFormat(TopShowStr,LinkStr,TitleNameLength,ConciseLength,TopCssStr,Rs)
else
   Call ShowFormat(ShowStr,LinkStr,TitleNameLength,ConciseLength,CssStr,Rs)
end if
Rs.movenext
i=i+1
Wend
End If
Rs.Close
Set Rs=Nothing
end sub
'-------------------------------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------------------------------
'信息列表公共函数: ShowList()
'数据库DataStr; '显示前n条N; '默认取值条件MrWhereStr; '动态取值条件WhereStr如：ViewFlag"&LangData&" and Instr(SortPath,'"&SortPath&"')>0; '排序方式OrderStr; '第一条显示方式TopShowStr; '显示方式ShowStr; '分列显示数TdNumstr默认为填1; '第一条标题显示字符长度TopTitleNameLength; '第一条简述显示字符长度TopConciseLength; '标题显示字符长度TitleNameLength; '简述显示字符长度ConciseLength; '第一条Css样式TopCssStr; 'Css样式CssStr; '链接地址LinkStr

sub ShowList(DataStr,N,MrWhereStr,WhereStr,OrderStr,TopShowStr,ShowStr,TdNumstr,TopTitleNameLength,TopConciseLength,TitleNameLength,ConciseLength,TopCssStr,CssStr,LinkStr)
Dim Rs,Sql,TdNum
if WhereStr="" then WhereStr=""&MrWhereStr&""
if OrderStr="" then OrderStr="OrderNum,ID desc"
if TopShowStr="" then TopShowStr="text"
if ShowStr="" then ShowStr="text"
if TdNumstr="" or TdNumstr=0 then TdNumstr=1 
TdNum=100/TdNumstr
if N="" then
Sql="select * from "&DataStr&" where "&WhereStr&" order by "&OrderStr
else
Sql="select top "&N&" * from "&DataStr&" where "&WhereStr&" order by "&OrderStr
end if
set Rs=conn.Execute(Sql)
Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
If Rs.EOF Then
        Response.Write "<tr>"&vbCrLf
		Response.Write "<td>&nbsp;</td>"&vbCrLf
	    Response.Write "</tr>"&vbCrLf
Else
Dim f
f=1'第一条
While Not Rs.eof
Response.Write "<tr>" & vbCrLf
For i=1 To TdNumstr
If Not Rs.eof then
Response.Write "<td width="""&TdNum&"%"" valign=""top"">" & vbCrLf
If Rs("HtmlUrlName"&LangData)<>"" then
HtmlUrlName=Rs("HtmlUrlName"&LangData)
else
HtmlUrlName=""&LinkStr&""
end if
if f=1 then   
   Call ShowFormat(TopShowStr,""&HtmlUrlName&"",TitleNameLength,TopConciseLength,TopCssStr,Rs)
else
   Call ShowFormat(ShowStr,""&HtmlUrlName&"",TitleNameLength,ConciseLength,CssStr,Rs)
end if
Response.Write "</td>" & vbCrLf
Rs.movenext
else
Response.Write "<td width="""&TdNum&"%"">&nbsp;</td>" & vbCrLf
End If
f=f+1
next
Response.Write "</tr>" & vbCrLf
Wend
End If
Response.Write "</table>"&vbCrLf
Rs.Close
Set Rs=Nothing
end sub
'-------------------------------------------------------------------------------------------------------------------------



'-------------------------------------------------------------------------------------------------------------------------
'信息分页列表公共函数: ShowPageList()
'数据库DataStr; '显示前n条N; 'ID取值Qz; '默认取值条件MrWhereStr; '动态取值条件WhereStr如：ViewFlag"&LangData&" and Instr(SortPath,'"&SortPath&"')>0; '排序方式OrderStr; '显示方式ShowStr; '分列显示数TdNumstr默认为填1; '标题显示字符长度TitleNameLength; '简述显示字符长度ConciseLength; '标题Css样式CssStr; '链接地址LinkStr; '分页Html名称HtmlUrlName; '每页显示信息条数PageSizeNum

Sub ShowPageList(DataStr,N,Qz,MrWhereStr,WhereStr,OrderStr,ShowStr,TdNumstr,TitleNameLength,ConciseLength,CssStr,LinkStr,PageUrlName,PageSizeNum)
Dim Rs,Sql,TdNum,Page
page=clng(request("page"))
PageUrlName=""&PageUrlName&""
HtmlId=Request("SortId")
if WhereStr="" or Qz="" then WhereStr=""&MrWhereStr&""
if OrderStr="" then OrderStr="OrderNum,ID desc"
if ShowStr="" then ShowStr="text"
if TdNumstr="" or TdNumstr=0 then TdNumstr=1 
TdNum=100/TdNumstr
if N="" then
Sql="select * from "&DataStr&" where "&WhereStr&" order by "&OrderStr
else
Sql="select top "&N&" * from "&DataStr&" where "&WhereStr&" order by "&OrderStr
end if
set Rs=server.CreateObject("adodb.recordset")
Rs.open Sql,conn,1,1
Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
If Rs.EOF Then
        Response.Write "<tr>"&vbCrLf
		Response.Write "<td>&nbsp;</td>"&vbCrLf
	    Response.Write "</tr>"&vbCrLf
Else
if page<1 or page="" then page=1
		Rs.pagesize=PageSizeNum
		if page>Rs.pagecount then page=Rs.pagecount
		Rs.absolutepage=page
Dim p
p=1
While Not Rs.eof and p<=Rs.pagesize
Response.Write "<tr>" & vbCrLf
For i=1 To TdNumstr
If Not Rs.eof then
Response.Write "<td width="""&TdNum&"%"" valign=""top"">" & vbCrLf
If Rs("HtmlUrlName"&LangData)<>"" then
HtmlUrlName=Rs("HtmlUrlName"&LangData)
else
HtmlUrlName=""&LinkStr&""
end if
   Call ShowFormat(ShowStr,""&HtmlUrlName&"",TitleNameLength,ConciseLength,CssStr,Rs)
Response.Write "</td>" & vbCrLf
Rs.movenext
else
Response.Write "<td width="""&TdNum&"%"">&nbsp;</td>" & vbCrLf
End If
p=p+1
next 
Response.Write "</tr>" & vbCrLf
Wend
End If
Response.Write "</table>"&vbCrLf
if Rs.recordcount>1 then
call NextPage(Page,Rs.pagecount,PageUrlName,HtmlId,3)
end if
Rs.Close
Set Rs=Nothing
end sub
'-------------------------------------------------------------------------------------------------------------------------



'-------------------------------------------------------------------------------------------------------------------------
'信息内容公共函数: Content()
'数据库DataStr; 'ID取值IDStr如："&request.QueryString("Id")&"; '默认取值条件MrWhereStr; '动态取值条件WhereStr如：ViewFlag"&LangData&" and Id="&request.QueryString("Id")&"; '排序方式OrderStr; '显示方式ShowStr; '标题Css样式CssStr;

Sub Content(DataStr,IDStr,MrWhereStr,WhereStr,OrderStr,ShowStr,CssStr)
Dim Rs,Sql
if IDStr="" then WhereStr=""&MrWhereStr&""
if OrderStr="" then OrderStr="OrderNum,ID"
if ShowStr="" then ShowStr="text"
if IDStr="" then
Sql="select top 1 * from "&DataStr&" where "&WhereStr&" order by "&OrderStr
else
Sql="select * from "&DataStr&" where "&WhereStr&""
end if
set Rs=server.CreateObject("adodb.recordset")
Rs.open Sql,conn,1,1
If Rs.EOF Then
		Response.Write "<center>&nbsp;</center>"&vbCrLf
Else
   Call ShowFormat(ShowStr,"","","",CssStr,Rs)
End If
Rs.Close
Set Rs=Nothing
end sub%>
<!--#include file="ShowFormat.Asp"-->