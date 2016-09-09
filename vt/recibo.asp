<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a83")="N" or session("a83")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Recibos de Vale-Transporte</title>
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

if request.form("ver")="" then
ano=0:mes=0:per=0
%>
<form action="recibo.asp" method="post" name="form">
<p class=realce>Emissão de Recibo de Entrega de VT</p>
Selecionar data: <select size="1" name="data" onchange="javascript:submit()">
<option value="0">Selecione uma data</option>
<%
sql2="select data from vt_saldo where deletada=0 and id_tipo=4 group by data order by data desc"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
if cdate(request.form("data"))=rs2("data") then tempsel="selected" else tempsel=""
%>
	<option value="<%=rs2("data")%>" <%=tempsel%>><%=rs2("data")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>
<br>
Funcionário: <select name="chapa">
<option value="0">Todos</option>
<%
	vartemp=request.form("data")
sql2="select v.chapa, nome from vt_saldo v, corporerm.dbo.pfunc f " & _
"where data='" & dtaccess(vartemp) & "' and deletada=0 and v.chapa=f.chapa collate database_default group by v.chapa, nome order by nome"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("chapa")%>"><%=rs2("nome")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>
<br>

<input type="submit" name="ver" value="Visualizar" class=button>
</form>
<%

else 'request.form
	vartemp=request.form("data")
	meiapagina=1
	pagina=0
	inicio=1
	chapa=request.form("chapa")
	numero=0
if chapa="0" then
	sql2="select v.chapa, nome from vt_saldo v, corporerm.dbo.pfunc f " & _
	"where data='" & dtaccess(vartemp) & "' and deletada=0 and v.chapa=f.chapa collate database_default group by v.chapa, nome order by nome"
	'response.write sql2
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	rs2.movefirst:	do while not rs2.eof
	redim preserve chapai(numero)
	chapai(numero)=rs2("chapa")
	rs2.movenext
	numero=numero+1
	loop
	rs2.close
else
	redim preserve chapai(0)
	chapai(0)=chapa
end if
for a=0 to ubound(chapai)
	chapap=chapai(a)
	sql="SELECT v.chapa, f.nome " & _
"FROM vt_saldo v, corporerm.dbo.pfunc f WHERE v.chapa=f.chapa collate database_default " & _
"and v.chapa='" & chapap & "' order BY f.nome "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<!-- table pagina -->
<table border="0" width=650 height="450">
<tr><td valign="top" class=campo>
<!-- table recibo -->
<table border="0" cellspacing="0" width="650" cellpadding="0" style="border-collapse: collapse">
<tr><td class=campo align="center"><font size=3><b>Recibo de Entrega de Vale Transporte</b></font></td></tr>
<tr><td class=campo align="center"><font size=2>FIEO - Fundação Instituto de Ensino para Osasco</font></td></tr>
<table>

<table border="1" bordercolor="#000000" cellspacing="0" width="650" cellpadding="3" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop">Matrícula<br><%=rs("chapa")%></td>
	<td valign=top class="campop">Nome<br><b><%=rs("nome")%></b></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellspacing="0" width="650" cellpadding="4" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop">
	Recebi da FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO-FIEO a quantidade de Vales-Transportes 
	abaixo discriminada, autorizando a descontar de meus vencimentos, os valores dos referidos 
	vales-transporte até o limite de 6% de meu salário.
	</td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellspacing="0" width="650" cellpadding="4" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop" colspan=5 align="center">Discriminação das Parcelas</td>
</tr>
<tr>
	<td valign=top class="campop" align="center">Cód.</td>
	<td valign=top class="campop" align="center">Discriminação</td>
	<td valign=top class="campop" align="center">Tarifa</td>
	<td valign=top class="campop" align="center">Quantidade</td>
	<td valign=top class="campop" align="center">Total</td>
</tr>
<%
totalp=0
if meiapagina=0 then meiapagina=1 else meiapagina=0

sql2="SELECT vt_saldo.codigo, PTARIFA.DESCRICAO, vt_saldo.tarifa, vt_saldo.quantidade, vt_saldo.total " & _
"FROM vt_saldo INNER JOIN corporerm.dbo.PTARIFA PTARIFA ON vt_saldo.codigo = PTARIFA.CODIGO collate database_default " & _
"WHERE vt_saldo.chapa='" & chapap & "' AND vt_saldo.data='" & dtaccess(vartemp) & "' AND id_tipo=4 AND deletada=0 " & _
"and vt_saldo.data between ptarifa.iniciovigencia and ptarifa.finalvigencia "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
totalp=totalp+cdbl(rs2("total"))
%>
<tr>
	<td valign=top class="campop"><%=rs2("codigo")%></td>
	<td valign=top class="campop"><%=rs2("descricao")%></td>
	<td valign=top class="campop" align="right"><%=formatnumber(rs2("tarifa"),2)%>&nbsp;</td>
	<td valign=top class="campop" align="right"><%=formatnumber(rs2("quantidade"),0)%>&nbsp;</td>
	<td valign=top class="campop" align="right"><%=formatnumber(rs2("total"),2)%>&nbsp;</td>
</tr>
<%
rs2.movenext
loop
rs2.close
%>
<tr>
	<td valign=top class="campop" colspan=4>Total</td>
	<td valign=top class="campop" align="right"><%=formatnumber(totalp,2)%>&nbsp;</td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellspacing="0" width="650" cellpadding="4" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop" width=100>Data<br>&nbsp;</td>
	<td valign=top class="campop">Assinatura<br>&nbsp;</td>
</tr>
</table>

</td></tr>
</table>
<!-- table pagina -->
<%
response.write "<p style='margin-top:0; margin-bottom:0'><font size=1>Recursos Humanos - FIEO"
response.write "<hr>"
if meiapagina=1 then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 

rs.close
'if a<>ubound(chapai) then response.write "<DIV style=""page-break-after:always""></DIV>"

next
end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>