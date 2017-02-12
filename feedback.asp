<% @ Language=VBScript %>
<% OPtion Explicit %>
<%
dim strConnectionString, con
set con=server.CreateObject("ADODB.Connection")
con.ConnectionString="Driver={Microsoft Access Driver (*.mdb)};Dbq="&server.mappath("feedback_database.mdb")

con.open
dim fname,lname,gender
dim occ,email,tele,address,city,pincode
dim state,country,comments
fname=Request.Form("fname")
lname=Request.Form("lname")
gender=Request.Form("gender")
occ=Request.Form("occupation")
email=Request.Form("email")
tele=Request.Form ("telephone")
address=Request.Form ("address")
city=Request.Form ("city")
pincode=Request.Form ("pincode")
state=Request.Form ("state")
country=Request.Form ("country")
comments=Request.Form("comments")

dim sql
sql="insert into feedback(fname, lname, gender, occ, email, tele, address, city, pincode, state, country, comments) VALUES('"&fname&"','"&lname&"','"&gender&"','"&occ&"','"&email&"','"&tele&"','"&address&"','"&city&"',"&pincode&",'"&state&"','"&country&"','"&comments&"')"
con.Execute(sql)
Response.Write "<h2>"&"Feedback details successfully submitted to database"
con.close
set con=Nothing
%>

