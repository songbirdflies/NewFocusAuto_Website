<%
dim searchkey
Searchkey=Trim(request("keyword"))
If searchkey="Enter the keyword" Then searchkey=""
Response.redirect "Search.Asp?keyword="&searchkey
%>