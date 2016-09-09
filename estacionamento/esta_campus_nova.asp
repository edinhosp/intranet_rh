<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a87")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Vaga de Estacionamento</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript" src="../date.js"></script>
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
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
<script src="../coolmenu/coolmenus_frame.js" type="text/javascript"></script>
<%
dim conexao, conexao2, chapach, rs, rs2, ok
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao

if request.form<>"" then
		if request.form("bt_salvar")<>"" then
		tudook=1
		'if request.form("salvar")="1" then

if request.form("campus_trabalho")="0" or request.form("campus_estudo")="0" or request.form("periodo")="0" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe as datas/campus!');</script>"
end if
		sqla = "INSERT INTO veiculos_alunosfunc (chapa, matricula, campus_trabalho, campus_estudo, periodo, "
		sqla = sqla & "validade, modelo, cor, placa, dtauto, sequencia, obs, usuarioc, datac "
		sqla = sqla & " )"
		
		sqlb = " SELECT '" & request.form("chapa") & "'"
		sqlb = sqlb & ",'" & request.form("matricula") & "'"
		sqlb = sqlb & ",'" & request.form("campus_trabalho") & "'"
		sqlb = sqlb & ",'" & request.form("campus_estudo") & "'"
		sqlb = sqlb & ",'" & request.form("periodo") & "'"
		sqlb = sqlb & ",'" & request.form("validade") & "'"
		sqlb = sqlb & ",'" & request.form("modelo") & "'"
		sqlb = sqlb & ",'" & request.form("cor") & "'"
		sqlb = sqlb & ",'" & request.form("placa") & "'"
		if request.form("dtauto")<>"" then sqlb=sqlb & ",'" & dtaccess(request.form("dtauto")) & "' " else sqlb=sqlb & ",null"
		sqlb = sqlb & ",'" & ucase(request.form("sequencia")) & "'"
		sqlb = sqlb & ",'" & request.form("obs") & "'"
		sqlb = sqlb & ",'" & session("usuariomaster") & "'"
		sqlb = sqlb & ",getdate()"
		sql = sqla & sqlb
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
		'end if
		end if 'request btsalvar
	else 'request.form=""
	end if

'if request.form("bt_salvar")="" then
if (request.form("bt_salvar")="") or (request.form("bt_salvar")<>"" and tudook=0) then
%>
<form method="POST" action="esta_campus_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Inclusão de Vaga de Estacionamento</td></tr>
</table>
<%
'for each strItem in Request.form
'	Response.write stritem & " = " & request.form(stritem) & " "
'next
if request("codigo")<>"" then
	chapa=request("codigo")
elseif request.form("chapa")<>"" then
	chapa=request.form("chapa") 
else
	chapa=""
end if
%>
<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo>0</td>
	<td class=fundo><input type="text" value="<%=chapa%>" name="chapa" size="5" onchange="chapa1()" onfocus="javascript:window.status='Informe o chapa do funcionário'"></td>
	<td class=fundo>
		<select size="1" name="nome" onchange="nome1()" onfocus="javascript:window.status='Selecione o Nome do Funcionário'" >
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' order by nome "
'if session("dp_chapa")<>"" then sql2=sql2 & "and chapa='" & session("dp_chapa") & "'" else sql2=sql2 & "order by nome"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o funcionário....</option>"
rsc.movefirst:do while not rsc.eof
if chapa=rsc("chapa") then temp="selected" else temp=""
%>
		<option value="<%=rsc("chapa")%>" <%=temp%>><%=rsc("nome")%></option>
<%
rsc.movenext:loop
rsc.close
%>
		</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Matricula</td>
	<td class=titulor>Campus<br>trabalho</td>
	<td class=titulor>Campus<br>estudo</td>
	<td class=titulor>Período<br>estudo</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="matricula" size="8" value=""></td>
<%if request.form("campus_trabalho")<>"" then tipo=request.form("campus_trabalho") else tipo=""%>
	<td class=fundo><select size="1" name="campus_trabalho">
		<option value="0">...</option>
		<option value="VY" <%if tipo="VY" then response.write "Selected"%>>Vila Yara</option>
		<option value="NS" <%if tipo="NS" then response.write "Selected"%>>Narciso</option>
		<option value="JW" <%if tipo="JW" then response.write "Selected"%>>Jd.Wilson</option>
		</select>
	</td>

<%if request.form("campus_estudo")<>"" then tipo=request.form("campus_trabalho") else tipo=""%>
	<td class=fundo><select size="1" name="campus_estudo">
		<option value="0">...</option>
		<option value="VY" <%if tipo="VY" then response.write "Selected"%>>Vila Yara</option>
		<option value="NS" <%if tipo="NS" then response.write "Selected"%>>Narciso</option>
		<option value="JW" <%if tipo="JW" then response.write "Selected"%>>Jd.Wilson</option>
		</select>
	</td>

<%if request.form("periodo")<>"" then tipo=request.form("periodo") else tipo=""%>
	<td class=fundo><select size="1" name="periodo">
		<option value="0">...</option>
		<option value="M" <%if tipo="M" then response.write "Selected"%>>Matutino</option>
		<option value="V" <%if tipo="V" then response.write "Selected"%>>Vespertino</option>
		<option value="N" <%if tipo="N" then response.write "Selected"%>>Noturno</option>
		</select>
	</td>
</tr>
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Validade</td>
	<td class=titulo>Veículo</td>
	<td class=titulo>Cor</td>
	<td class=titulo>Placa</td>
	<td class=titulo>Dt.Autor.</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="validade" size="7" value=""></td>
	<td class=fundo><input type="text" name="modelo" size="15" value=""></td>
	<td class=fundo><input type="text" name="cor" size="5" value=""></td>
	<td class=fundo><input type="text" name="placa" size="7" value=""></td>
	<td class=fundo><input type="text" name="dtauto" size="7" value=""></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Sequência</td>
	<td class=titulo>Observação</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="sequencia" size="5" value=""></td>
	<td class=fundo><input type="text" name="obs" size="45" value=""></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2" onfocus="javascript:window.status='Clique para desfazer e limpar a tela'"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()" onfocus="javascript:window.status='Clique aqui para fechar sem salvar'"></td>
</tr>
</table>
</form>
<%
'else
'rs.close
set rs=nothing
end if   'request.form=""
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<!--
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if

%>
</body>
</html>