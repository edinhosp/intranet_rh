<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a36")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Controle de Férias</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<!-- <script language="JavaScript" type="text/javascript" src="../date.js"></script> -->
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

if request.form("dtinigozo")="" or request.form("dtfimgozo")="" or request.form("dtfimper")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe as datas do periodo aquisitivo e do gozo!');</script>"
end if
		dias=request.form("dias")
		if request.form("tipo")="P" then dias=0
		if dias<0 then dias=cdate(request.form("dtfimgozo"))-cdate(request.form("dtinigozo"))+1
		sqla = "INSERT INTO ferias (chapa, "
		if request.form("dtfimper")<>"" then sqla=sqla & "dtfimper, "
		if request.form("dtinigozo")<>"" then sqla=sqla & "dtinigozo, "
		if request.form("dtfimgozo")<>"" then sqla=sqla & "dtfimgozo, "
		sqla = sqla & "dias, obs, tipo "
		sqla = sqla & " )"
		
		sqlb = " SELECT '" & request.form("chapa") & "'"
		if request.form("dtfimper")<>"" then sqlb=sqlb & ",'" & dtaccess(request.form("dtfimper")) & "' "
		if request.form("dtinigozo")<>"" then sqlb=sqlb & ",'" & dtaccess(request.form("dtinigozo")) & "' "
		if request.form("dtfimgozo")<>"" then sqlb=sqlb & ",'" & dtaccess(request.form("dtfimgozo")) & "' "
		sqlb=sqlb & "," & dias & ""
		sqlb=sqlb & ",'" & request.form("obs") & "' "
		sqlb=sqlb & ",'" & request.form("tipo") & "' "
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
<form method="POST" action="fer_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Inclusão de Controle de Férias</td></tr>
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
	<td class=titulo>Tipo</td>
	<td class=titulo>Venc.PAq</td>
	<td class=titulo>Obs</td>
</tr>
<tr>
<%if request.form("tipo")<>"" then tipo=request.form("tipo") else tipo=""%>
	<td class=fundo><select size="1" name="tipo">
		<option value="C" <%if tipo="C" then response.write "Selected"%>>Crédito</option>
		<option value="D" <%if tipo="D" then response.write "Selected"%>>Débito</option>
		<option value="P" <%if tipo="P" then response.write "Selected"%>>Prev.Gozo</option>
		</select>
	</td>
	<td class=fundo>
	
		<input type="text" name="dtfimper" size="9" value="" onchange="mand_ini1(1)" > 
<%if teste=1 then %>		
		<select size="1" name="dtfimper">
		<option value="<%=m%>" <%if m=cint(mes) then response.write "Selected"%>><%=m%></option>
		</select>
<% end if%>
	</td>
	<td class=fundo><input type="text" name="obs" size="20" value=""></td>
</tr>
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Início Gozo</td>
	<td class=titulo>Término Gozo</td>
	<td class=titulo>Dias</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="dtinigozo" size="9" value="" onchange="mand_ini1(1)"   ></td>
	<td class=fundo><input type="text" name="dtfimgozo" size="9" value="" onchange="mand_ini1(1)"   ></td>
	<td class=fundo><input type="text" name="dias" size="8" value="" onfocus="this.blur()" ></td>
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