<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a75")="N" or session("a75")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Gerador de etiquetas</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script type="text/javascript" src="../jquery.js"></script>
<SCRIPT LANGUAGE="JavaScript" type="text/javascript" src="../selectbox.js"></SCRIPT>
<script language="JavaScript" type="text/javascript"><!--
function library1() {
	temp=form.id_etiq.value
	tipo=temp.substring(0,1)
	temp=temp.substring(0,temp.length)
	form.textosql.value=temp
}	
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsquery=server.createobject ("ADODB.Recordset")
set rsquery.ActiveConnection = conexao

pixel=96/2.54
point=72/2.54
pointp=72.27/2.54
vlinha=13

pixel=96/2.54
point=72/2.54
pointp=72.27/2.54

colunas     =request.form("colunas")      '3
linhas      =request.form("linhas")       '10
msuperior   =request.form("msuperior")    '1,27
mdireita    =request.form("mdireita")     '0,4
mesquerda   =request.form("mesquerda")    '0,4
altura      =request.form("altura")       '2,54
largura     =request.form("largura")      '6,67
espacolinha =request.form("espacolinha")  '0
espacocoluna=request.form("espacocoluna") '0,3
msuperiorp  =request.form("msuperiorp")   '0,2
mesquerdap  =request.form("mesquerdap")   '0,2
sqltexto    =request.form("sqltexto")
sql=sqltexto
rs.Open sql, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
folha=colunas*linhas
paginas=total/folha
if paginas=int(paginas) then paginas=paginas else paginas=int(paginas)+1
itens=rs.fields.count-1

for a=1 to paginas

response.write "<table border='0' bordercolor='#000000' cellpadding='0' cellspacing='0' style='border-collapse: collapse'>"
for l=1 to linhas
	response.write "<tr>"
	for c=1 to colunas
%>
	<td width='<%=largura*pixel%>px' height='<%=altura*pixel%>px' valign=top>
	<font style="color:black">
	<%
	if not rs.eof then
	for i=0 to rs.fields.count-1
		response.write "<span style='font-size:11px'>" & rs.fields(i).value & "</span>"
		if i<rs.fields.count-1 then response.write "<br>"
	next 'l
	if not rs.eof then rs.movenext
	end if
	%>
	</td>
<%	
	if cint(c)<cint(colunas) then
		acrescimo=0
		'if c=2 then acrescimo=40 else acrescimo=0
%>
		<td width='<%=espacocoluna*pixel+acrescimo%>px'>&nbsp;</td>
<%
	end if
	next 'c
	response.write "</tr>"
next 'l
response.write "</table>"
if a<paginas then response.write "<DIV style=""page-break-after:always""></DIV>"

next 'a

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>