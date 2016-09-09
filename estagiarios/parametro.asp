<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")="N" or session("a72")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Parâmetros de Horário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript" src="../date.js"></script>
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome1()		{form.chapa.value=form.nome.value;}
function chapa1()		{form.nome.value=form.chapa.value;}
function descricao1()	{form.codigo.value=form.descricao.value;}
function codigo1()		{form.descricao.value=form.codigo.value;}

function mand_ini1(muda) {
	temp=form.dtinigozo.value;
	inicio=new Date(temp.substr(6),temp.substr(3,2)-1,temp.substr(0,2));
	temp2=form.dtfimgozo.value;
	termino=new Date(temp2.substr(6),temp2.substr(3,2)-1,temp2.substr(0,2));
	dinicio=montharray[inicio.getMonth()]+" "+inicio.getDate()+", "+inicio.getFullYear()
	dfinal=montharray[termino.getMonth()]+" "+termino.getDate()+", "+termino.getFullYear()
	dias=(Math.round((Date.parse(dfinal)-Date.parse(dinicio))/(24*60*60*1000))*1)+1
	document.form.dias.value=dias
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
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1

	if request.form("inicio")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe o início do período!');</script>"
	if request.form("fim")=""    then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe o fim do período!');</script>"
	if request.form("ano")=""    then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe o ano!');</script>"
	if request.form("mes")=""    then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe o mês!');</script>"

	sql="UPDATE est_parametro SET "
	sql=sql & "ano=" & request.form("ano")
	sql=sql & ", mes=" & request.form("mes")
	sql=sql & ", descricao='" & request.form("descricao") & "'"
	sql=sql & ", inicio='" & dtaccess(request.form("inicio")) & "'"
	sql=sql & ", fim='" & dtaccess(request.form("fim")) & "'"
	sql=sql & ", limite=" & request.form("limite")
	'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
	'sql=sql & ",dataa   =getdate() "
	'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
	'response.write sql
	if tudook=1 then conexao.Execute sql, , adCmdText
end if


sql="select top 1 * from est_parametro "
rs.Open sql, ,adOpenStatic, adLockReadOnly

if request.form("ano")=""       then ano      =rs("ano")       else ano=request.form("ano")
if request.form("mes")=""       then mes      =rs("mes")       else mes=request.form("mes")
if request.form("descricao")="" then descricao=rs("descricao") else descricao=request.form("descricao")
if request.form("inicio")=""    then inicio   =rs("inicio")    else inicio=request.form("inicio")
if request.form("fim")=""       then fim      =rs("fim")       else fim=request.form("fim")
if request.form("limite")=""    then limite   =rs("limite")    else limite=request.form("limite")

%>
<form method="POST" action="parametro.asp" name="form">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Parâmetros de Horários - Estagiário</td></tr>
</table>

<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Ano</td>
	<td class=titulo>Mês</td>
	<td class=titulo>Descrição</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="ano" size="4" value="<%=ano%>"></td>
	<td class=titulo><input type="text" name="mes" size="2" value="<%=mes%>"></td>
	<td class=titulo><input type="text" name="descricao" size="50" value="<%=descricao%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Inicio</td>
	<td class=titulo>Fim</td>
	<td class=titulo>Limite min</td>
	<td class=titulo width="60%">&nbsp;</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="inicio" size="8" value="<%=inicio%>"></td>
	<td class=titulo><input type="text" name="fim" size="8" value="<%=fim%>"></td>
	<td class=titulo><input type="text" name="limite" size="1" value="<%=limite%>"></td>
	<td class=titulo>&nbsp;</td>
</tr>
</table>


<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser alterado!');</script>"
	end if
end if

conexao.close
set conexao=nothing
%>
</body>
</html>