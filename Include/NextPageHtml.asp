<%
sub nextpage(page,fpagecount,HtmlUrlName,HtmlId,skip)'page当前页，fpagecount页数,skip中间数字标
Const HtmlName="html"
dim nextstr
nextstr=nextstr&"<div id='nextpage'>"
dim j
if skip<>"" then
	if fpagecount<=skip then
	for i=1 to fpagecount
		if HtmlId>0 then 
		    if page=i then
			   nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&HtmlId&"-"&i&"."&HtmlName&"' class='f14'> "&i&" </a>"
			else
			   nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&HtmlId&"-"&i&"."&HtmlName&"'> ["&i&"]</a>"
			end if
		else
			if page=i then
			   nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&i&"."&HtmlName&"' class='f14'> "&i&" </a>"
			else
			   nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&i&"."&HtmlName&"'> ["&i&"] </a>"
			end if
		end if
	next
    else
	if page+int(skip/2)>fpagecount then j=fpagecount-skip else j=page-int(skip/2)
		if j<1 then j=1
		if j=1 then
		    if HtmlId>0 then 
				nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&HtmlId&"-"&page-1&"."&HtmlName&"' class='f14 Webdingss margin5'> 9 </a>"
			else
				nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&page-1&"."&HtmlName&"' class='f14 Webdingss margin5'> 9 </a>"
			end if	
		else	
		    if HtmlId>0 then 
				nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&HtmlId&"-"&page-1&"."&HtmlName&"' class='f14 Webdingss margin5'> 9 </a>"
				nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&HtmlId&"-1."&HtmlName&"'>[ 1 ]</a>"
			else
				nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&page-1&"."&HtmlName&"' class='f14 Webdingss margin5'> 9 </a>"
				nextstr=nextstr&"<a href='"&HtmlUrlName&"-1."&HtmlName&"'>[ 1 ]</a>"
			end if	
		    if j-1<>1 then nextstr=nextstr+"..."
		end if
			
		for i=1 to skip
			if HtmlId>0 then 
				if page=j then
					nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&HtmlId&"-"&j&"."&HtmlName&"' class='f14'> "&j&" </a>"
				else
					nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&HtmlId&"-"&j&"."&HtmlName&"'> ["&j&"] </a>"
				end if
			else
				if page=j then
					nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&j&"."&HtmlName&"' class='f14'> "&j&" </a>"
				else
					nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&j&"."&HtmlName&"'> ["&j&"] </a>"
				end if
			end if
			j=j+1
		next
		
		if j-1=fpagecount then
			nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&HtmlId&"-"&page+1&"."&HtmlName&"' class='f14 Webdingss margin5'> : </a>"
		else
			if j<fpagecount then nextstr=nextstr+"..."
			if HtmlId>0 then
				if page=j then
					nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&HtmlId&"-"&fpagecount&"."&HtmlName&"' class='f14'> "&fpagecount&" </a>"
				else
					nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&HtmlId&"-"&fpagecount&"."&HtmlName&"'> ["&fpagecount&"] </a>"
				end if
				nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&HtmlId&"-"&page+1&"."&HtmlName&"' class='f14 Webdingss margin5'> : </a>"
			else
				if page=j then
					nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&fpagecount&"."&HtmlName&"' class='f14'> "&fpagecount&" </a>"
				else
					nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&fpagecount&"."&HtmlName&"'> ["&fpagecount&"] </a>"
				end if					
				nextstr=nextstr&"<a href='"&HtmlUrlName&"-"&page+1&"."&HtmlName&"' class='f14 Webdingss margin5'> : </a>"
			end if
	    end if
    end if
	nextstr=nextstr&"</div>"
    response.Write(nextstr)
end if
end sub
%>