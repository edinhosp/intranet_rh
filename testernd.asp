<%@ Language=VBScript %>
<!-- #Include file="ADOVBS.INC" -->
<!-- #Include file="funcoesclear.inc" -->
<html>
<head>
</head>
<body>


<%
randomize timer
response.write rnd
response.write "<br>"
response.write int(rnd()*10)
response.write "<br>"
response.write int(rnd()*100)


%>