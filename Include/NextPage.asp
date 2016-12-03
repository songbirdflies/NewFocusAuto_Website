<%
'sub nextpage(page,fpagecount,skip)'page当前页，fpagecount页数,skip中间数字标
sub nextpage(page,fpagecount,HtmlUrlName,HtmlId,skip)'page当前页，fpagecount页数,skip中间数字标
dim nextstr
Dim query, a, x, temp, action,url
action="http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
query=Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
        a=Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
           temp=temp & a(0) & "=" & a(1) & "&"
        End If
    Next
	url=action&"?"&temp
	nextstr=nextstr&"<div id='nextpage'>"

	dim j
	if skip<>"" then				
		if fpagecount<=skip then
			for i=1 to fpagecount
				if InStr(1,url,"?")<>0 then 
					if page=i then
						nextstr=nextstr&"<a href='"&url&"page="&i&"' class='f14'> "&i&" </a>"
					else
						nextstr=nextstr&"<a href='"&url&"page="&i&"'> ["&i&"] </a>"
					end if
				else
					if page=i then
						nextstr=nextstr&"<a href='"&url&"page="&i&"' class='f14'> "&i&" </a>"
					else
						nextstr=nextstr&"<a href='"&url&"?page="&i&"'> ["&i&"] </a>"
					end if
				end if
			next
		else
			if page+int(skip/2)>fpagecount then j=fpagecount-skip else j=page-int(skip/2)
			if j<1 then j=1
			if j=1 then
				nextstr=nextstr&"<a href='"&url&"page="&page-1&"' class='f14 Webdingss margin5'> 9 </a>"
			else	
				if InStr(1,url,"?")<>0 then 
					nextstr=nextstr&"<a href='"&url&"page="&page-1&"' class='f14 Webdingss margin5'> 9 </a>"
					nextstr=nextstr&"<a href='"&url&"page=1'>[ 1 ]</a>"
				else
					nextstr=nextstr&"<a href='"&url&"?page="&page-1&"' class='f14 Webdingss margin5'> 9 </a>"
					nextstr=nextstr&"<a href='"&url&"?page=1'>[ 1 ]</a>"
				end if	
				if j-1<>1 then nextstr=nextstr+"..."
			end if
			
			for i=1 to skip
				if InStr(1,url,"?")<>0 then 
					if page=j then
						nextstr=nextstr&"<a href='"&url&"page="&j&"' class='f14'> "&j&" </a>"
					else
						nextstr=nextstr&"<a href='"&url&"page="&j&"'> ["&j&"] </a>"
					end if
				else
					if page=j then
						nextstr=nextstr&"<a href='"&url&"page="&j&"' class='f14'> "&j&" </a>"
					else
						nextstr=nextstr&"<a href='"&url&"?page="&j&"'> ["&j&"] </a>"
					end if
				end if
				j=j+1
			next
			if j-1=fpagecount then
				nextstr=nextstr&"<a href='"&url&"page="&page+1&"' class='f14 Webdingss margin5'> : </a>"
			else
				if j<fpagecount then nextstr=nextstr+"..."
				if InStr(1,url,"?")<>0 then
				    if page=j then
					nextstr=nextstr&"<a href='"&url&"page="&fpagecount&"' class='f14'> "&fpagecount&" </a>"
					else
					nextstr=nextstr&"<a href='"&url&"page="&fpagecount&"'> ["&fpagecount&"] </a>"
					end if
					nextstr=nextstr&"<a href='"&url&"page="&page+1&"' class='f14 Webdingss margin5'> : </a>"
				else
				    if page=j then
					nextstr=nextstr&"<a href='"&url&"?page="&fpagecount&"' class='f14'> "&fpagecount&" </a>"
					else
					nextstr=nextstr&"<a href='"&url&"?page="&fpagecount&"'> ["&fpagecount&"] </a>"
					end if					
					nextstr=nextstr&"<a href='"&url&"?page="&page+1&"' class='f14 Webdingss margin5'> : </a>"
				end if
			end if
		end if
		nextstr=nextstr&"</div>"
		response.Write(nextstr)
	end if
end sub
%>