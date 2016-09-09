<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a87")="N" or session("a87")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Relação de Estacionamento da Brasil Park</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form="" then
%>

<p class=titulo>Seleção para impressão de recibo de cartão B.Park</p>
<form method="POST" action="brasilparkrecibo.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=350>
<tr>
	<td class=titulo>Pessoa</td>
</tr>
<tr>
	<td class=titulo>
	<select size="1" name="chapa">
		<option value="Todos">Todos</option>
		<option value="Novos">Novas inclusões</option>
<%
sql1="select chapa, nome from dc_professor where codsituacao<>'D' order by nome "
sql1="SELECT va.chapa, va.bp, f.nome " & _
"FROM veiculos v, veiculos_a va, " & _
"(select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' and codtipo in ('N','T','A') " & _
"union all select chapa collate database_default, nome collate database_default from grades_novos) f " & _
"WHERE v.chapa=va.chapa AND v.chapa=f.chapa " & _
"AND getdate() Between inicio And termino GROUP BY va.chapa, f.nome, va.bp HAVING va.bp=1 ORDER BY f.nome "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof
%>
<option value="<%=rs("chapa")%>"><%=rs("nome")%></option>
<%
rs.movenext
loop
rs.close
%>			
	</select>
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=350>
<tr><td align="center" class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3"></td></tr>
</table>
</form>
<hr>

<%
end if 'request.form


if request.form<>"" then
chapa=request.form("chapa")
if chapa="Todos" then
	sql2=""
elseif chapa="Novos" then
	sql2=" and status='A' and pabp=0 and bp=1 "
else
	sql2=" AND n.chapa='" & chapa & "' "
end if
sql1="select c.chapa, n.nome, n.descricao, n.codsindicato from " & _
"(select chapa, pabp, bp, status from veiculos_a where bp=1 and getdate() between inicio and termino group by chapa, pabp, bp, status) c, " & _
"(select chapa, nome, min(descricao) as descricao, min(codsindicato) as codsindicato from (select chapa, nome, codsecao as descricao, codsindicato from grades_novos union all select f.chapa collate database_default, f.nome collate database_default, f.secao collate database_default, f.codsindicato collate database_default from qry_funcionarios f where f.codsituacao<>'D') t group by chapa, nome) n " & _
"where c.chapa=n.chapa " & sql2 & " order by n.nome "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<!-- table pagina -->
<table border="0" width=620 height="450">
<tr><td valign="top" class=campo>
<!-- table recibo -->
<%
rs.movefirst
do while not rs.eof
for a=1 to 2
%>
<div align="right">
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=620>
<tr>
	<td class=campo colspan=1 align="left"><img src="../images/logo_centro_universitario_unifieo_big.jpg" width="180" border="0"></td>
	<td class="campop" colspan=1 align="center"><b>RECIBO DE ENTREGA<br>Cartão de Estacionamento</td>
	<td class=campo colspan=1 align="right" style="color:blue;font-size:20pt"><b>BrasilPark</td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="5" cellspacing="0" style="border-collapse: collapse" width=620>
<tr>
	<td class="campop"><b>Nome: <%=rs("nome")%></td><td class="campop" width=100>Reg.: <%=rs("chapa")%></td>
</tr>
<tr>
	<td class="campop" colspan=2>Local: <%=rs("descricao")%></td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="5" cellspacing="0" style="border-collapse: collapse" width=620>
<tr>
	<td class="campop" style="border: 1px solid"><br>
Estou recebendo o cartão nº ______________ para uso no estacionamento localizado no bloco Branco
e gerenciado pela BrasilPark.<br>
<br>
Estou ciente de que:
<ol>
	<li>O cartão é provisório, para ser utilizado durante a vigência do contrato de trabalho ou enquanto as normas de
	utilização estiverem sendo satisfeitas (nº de aulas atribuídas na graduação, local de trabalho, etc).</li>
	<li>No caso de esquecimento do cartão, ficará sob meu encargo o valor da diária avulsa.</li>
	<li>No caso de perda, que deverá ser comunicado imediatamente à BrasilPark ou à FIEO, arcarei com o custo de
	emissão de um novo cartão, atualmente fixado em R$ 20,00 (vinte reais).</li>
</ol>
Osasco, 
<br>
<br>
<br>
_____________________________________________________
<br>
<br>
<br>	
	</td>
</tr>
<%
if a=1 then via="FIEO" else via="Cópia BrasilPark"
%>
<tr><td class=campo align="right"><%=via%> - <%=rs.absoluteposition%></td></tr>
</table>
<%if a=1 then response.write "<hr>"%>
</div>
<%
next

sql2="select marca, modelo, placa from veiculos where chapa='" & rs("chapa") & "' and dttermino is null"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
	response.write "<hr>"
	response.write "<table border=1 bordercolor=#000000 cellpadding=3 cellspacing=1 style='border-collapse: collapse'>"
	response.write "<tr>"
	do while not rs2.eof
		response.write "<td class=titulo style='border-left:3 solid'>Marca</td><td class=titulo>Modelo</td><td class=titulo>Placa</td>"
	rs2.movenext
	loop
	response.write "</tr>"
	response.write "<tr>"
	rs2.movefirst
	do while not rs2.eof
			response.write "<td class=campo style='border-left:3 solid'>" & rs2("marca") & "</td><td class=campo>" & rs2("modelo") & "</td><td class=campo>" & rs2("placa") & "</td>"
	rs2.movenext
	loop
	response.write "</tr>"
	rs2.close
	response.write "</table>"
end if
if rs.recordcount>1 and rs.recordcount<>rs.absoluteposition then response.write "<DIV style=""page-break-after:always""></DIV>"
rs.movenext
loop
rs.close
%>

<!-- table recibo -->
</td></tr>
</table>
<!-- table pagina -->
<%
end if 'request.form

set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>