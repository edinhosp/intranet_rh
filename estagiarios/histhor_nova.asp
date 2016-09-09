<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Histórico de Horário</title>
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
dim conexao, chapach, rs
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

if request.form("dtmudanca")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe a data da mudança!');</script>"
if request.form("codigo")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe o código do horário!');</script>"
if request.form("chapa")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe a chapa do funcionário!');</script>"

		sqla = "INSERT INTO est_histhor (chapa, codigo, dtmudanca, dia )"
		sqlb = " SELECT '" & request.form("chapa") & "'"
		sqlb=sqlb & ", '" & request.form("codigo") & "' "
		sqlb=sqlb & ", '" & dtaccess(request.form("dtmudanca")) & "' "
		sqlb=sqlb & ", " & request.form("dia") & " "
		'sqlb=sqlb & ",'" & session("usuariomaster") & "'"
		'sqlb=sqlb & ",getdate()"
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
<form method="POST" action="histhor_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Inclusão de Histórico de Horário</td></tr>
</table>
<%
'for each strItem in Request.form
'	Response.write stritem & " = " & request.form(stritem) & " "
'next
if request("codigoh")<>"" then
	chapa=request("codigoh")
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
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' and codtipo='T' order by nome "
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
	<td class=titulo>Horário</td>
</tr>
<tr>
	<td class=fundo>
		<input type="text" value="<%=request.form("codigo")%>" name="codigo" size="5" onchange="codigo1();form.submit();" onfocus="javascript:window.status='Informe o codigo do horário';">
		<select size="1" name="descricao" onchange="descricao1();form.submit();" onfocus="javascript:window.status='Selecione o horário';" >
<%
sql2="select codigo, descricao from est_cadhorario order by descricao "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o horário....</option>"
rsc.movefirst:do while not rsc.eof
if request.form("codigo")=rsc("codigo") then temp="selected" else temp=""
%>
		<option value="<%=rsc("codigo")%>" <%=temp%>><%=rsc("descricao")%></option>
<%
rsc.movenext:loop
rsc.close
%>
		</select></td>
</tr>
</table>
<!--
		<input type="text" name="dtfimper" onchange="mand_ini1(1)" size="9" value="" onFocus="javascript:vDateType='3'" onKeyUp="DateFormat(this,this.value,event,false,'3')" onBlur="DateFormat(this,this.value,event,true,'3')" >
-->

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Data Mudança</td>
	<td class=titulo>Indice</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="dtmudanca" size="9" value="<%=request.form("dtmudanca")%>"></td>
	<td class=fundo>
		<select size="1" name="dia" onchange="" onfocus="javascript:window.status='Selecione o dia de indice'" >
<%
if request.form("codigo")="" then codigo=0 else codigo=request.form("codigo")
sql2="select dia, comp, [desc] from est_cadhorario_marcacoes where codigo='" & codigo & "' order by dia "
response.write sql2
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
rsc.movefirst:do while not rsc.eof
if request.form("dia")=rsc("dia") then temp="selected" else temp=""
%>
		<option value="<%=rsc("dia")%>" <%=temp%>><%=rsc("dia")%></option>
<%
rsc.movenext:loop
end if 'recordcount>0
rsc.close
%>
		</select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
	</td>
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