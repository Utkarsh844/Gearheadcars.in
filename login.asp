<% @ language="vbscript"%>
<% option explicit %>
<% dim nm,pd
nm=Request.QueryString("name")
pd=Request.QueryString("password")
if nm="vinay" AND pd= "vinay" then

Response.Redirect("Welcome.html")
else
Response.Write("Invalid User!")
end if
%>