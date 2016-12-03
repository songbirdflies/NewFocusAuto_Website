<%
dim Language
Language=request.queryString("Language")
select case Language
  case "Simplified"
    response.cookies("Language")="Si"
  case "English"
    response.cookies("Language")="En"
end select
response.cookies("Language").expires = DateAdd("m",1,now)
%>
