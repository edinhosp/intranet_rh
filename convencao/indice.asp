<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a45")="N" or session("a45")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Indices</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:40px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

sql="select * from indices"
rs.Open sql, ,adOpenStatic, adLockReadOnly

%>
<p class=titulo>Prévia da correção salarial - Março/2012</p>
<%
'*************** inicio teste **********************
response.write "<table border='0' bordercolor='#000000' cellpadding='3' cellspacing='0' style='border-collapse:collapse'>"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	nome=ucase(rs.fields(a).name)
	if a>0 and a<13 then nome=mid(nome,2,2) & "/" & mid(nome,4,2)
	response.write "<td class=titulo style='border: 1px solid #000000'>" & nome & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	valor=rs.fields(a)
	if a=13 then total1=total1+valor
	if a=14 then total2=total2+valor
	if rs.fields(a).type=5 and valor<>"" then valor=formatpercent(valor,2)
	response.write "<td class=campo nowrap style='border: 1px solid #000000'>" & valor & "</td>"
	
next
response.write "</tr>"
rs.movenext
loop
response.write "<tr><td colspan=13 class=titulo>Média aritmética atual</td>"
response.write "<td class=""campol"">" & formatpercent(total1/3,2) & "</td>"
response.write "<td class=""campol"" align=""center"" style='border:2 dotted #000000'><b>" & formatpercent(total2/3,2) & "</td>"
indice=total2/3
indice=int(indice*10000+0.5)/10000
if indice>0.0999 then
	dissidio=0.0999
	convencao=indice-dissidio
else
	dissidio=indice
	convencao=0
end if
%>
<%
response.write "</table>"
response.write "<p>Indice aritmético: <b>" & formatpercent(indice,2) & "</b>"
response.write "<br>Correção Salarial: <b>" & formatpercent(dissidio,2) & "</b>"
response.write "<br>Indice para discussão: " & formatpercent(convencao,2)
response.write "<p>"
'*************** fim teste **********************
%>
<!--
<script language='JavaScript' src='http://www.debit.com.br/resumogratuito.php?info1=inflacao'></script>
-->

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td class=titulop>5. Reajuste salarial em 2012
<tr><td class=campo style="text-align:justify">Em 1º de março de 2012, as MANTENEDORAS deverão aplicar sobre os salários devidos em 1º de março de 2011, o percentual definido pela média aritmética dos índices inflacionários do período compreendido entre 1º de março de 2011 e 29 de fevereiro de 2012, apurados pelo IBGE (INPC), FIPE (IPC) e DIEESE (ICV), até o limite de 6,5% (seis e meio por cento).
<br><b>Parágrafo primeiro</b> – Caso o limite de 6,5% (seis e meio por cento) seja ultrapassado, as entidades signatárias negociarão, no prazo de 90 (noventa) dias a contar de 1º de abril de 2012, o pagamento da diferença entre a média aritmética dos índices inflacionários e 6,5%, sendo certo que, para base de cálculo de março de 2013, será considerada a média aritmética dos índices inflacionários, sem o limite estabelecido no caput.
</table>
 

<%
'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>