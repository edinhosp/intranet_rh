<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a47")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Lançamento Folha</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome1()	{	form.chapa.value=form.nome.value;	}
function chapa1()	{	form.nome.value=form.chapa.value;	}
function evento1()	{	form.codevento.value=form.evento.value;	form.submit();	}
function codigo1()	{	form.evento.value=form.codevento.value;	form.submit();	}
--></script>
</head>
<body>
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
		valor=request.form("valor")
		if request.form("tipolanc")="H" then
			valor1=replace(valor,":",",")
			valor1=replace(valor1,".",",")
			hora=int(valor1)
			minuto=(valor1-hora)*100
			valor=hora*60 + minuto
		else
			valor=valor
		end if

		sqla = "INSERT INTO apont_adm (chapa, ano, mes, codevento, valor, nrovezes "
		sqla = sqla & " )"
		
		sqlb = " SELECT '" & request.form("chapa") & "'"
		sqlb=sqlb & ",'" & request.form("ano") & "'"
		sqlb=sqlb & "," & request.form("mes")
		sqlb=sqlb & ",'" & request.form("codevento") & "'"
		sqlb=sqlb & "," & nraccess(valor)
		sqlb=sqlb & "," & nraccess(request.form("nrovezes"))
		'sqlb=sqlb & ",'" & session("usuariomaster") & "'"
		'sqlb=sqlb & ",getdate()"
		sql = sqla & sqlb
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if 'request btsalvar
else 'request.form=""
end if
if request.form("bt_salvar")<>"" then
else
end if	

'if request.form("bt_salvar")="" then
%>
<form method="POST" action="apont_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
	<tr><td class=grupo>Inclusão de Lançamento Folha</td></tr>
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
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo>0</td>
	<td class=fundo><input type="text" value="<%=chapa%>" name="chapa" size="5" onfocus="javascript:window.status='Informe o chapa do funcionário'" onchange="chapa1()"></td>
	<td class=fundo>
		<select size="1" name="nome" onfocus="javascript:window.status='Selecione o Nome do Funcionário'" onchange="nome1()">
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

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Ano</td>
	<td class=titulo>Mês</td>
	<td class=titulo>Evento</td>
</tr>
<tr>
<%if request.form("ano")<>"" then ano=request.form("ano") else ano=year(now)%>
<%if request.form("mes")<>"" then mes=request.form("mes") else mes=month(now)%>
	<td class=fundo><select size="1" name="ano">
	<%for a=year(now)-1 to year(now)+1%>
		<option value="<%=a%>" <%if a=cint(ano) then response.write "Selected"%>><%=a%></option>
	<%next%>
		</select>
	</td>
	<td class=fundo><select size="1" name="mes">
	<%for m=1 to 13%>
		<option value="<%=m%>" <%if m=cint(mes) then response.write "Selected"%>><%=m%></option>
	<%next%>
		</select>
	</td>
	<td class=fundo>
<%if request.form("codevento")<>"" then codevento=request.form("codevento") else codevento="0"%>
	<input type="text" name="evento" size="3" value="<%=codevento%>" onchange="evento1()">
	<select size="1" name="codevento" onfocus="javascript:window.status='Selecione o evento'" onchange="codigo1()">
<%
sqla="select codigo, descricao from apont_adm_eventos order by descricao "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if codevento=rsd("codigo") then tempc="selected" else tempc=""
%>
		<option value="<%=rsd("codigo")%>" <%=tempc%>><%=rsd("descricao")%></option>
<%
rsd.movenext:loop
rsd.close
%>
	</select></td>
</tr>
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Tipo</td>
	<td class=titulo>Valor/Hora/Dia</td>
	<td class=titulo>Nº Vezes</td>
</tr>
<tr>
<%
sql="select valhordiaref, fator, provdescbase from apont_adm_eventos where codigo='" & codevento & "' "
rsd.Open sql, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
	tplanc=rsd("valhordiaref")
	tpeve=rsd("provdescbase")
	fator=rsd("fator")
end if
rsd.close
if request.form("valor")<>"" then 
	valor=request.form("valor"):primeira=1
else
	valor=0:primeira=0
end if
if tplanc="D" then
	'if isdate(valor)=true then valor=hour(valor)
	if request.form("tipolanc")="H" then valor=hour(valor)+(minute(valor)/100)
	valord=valor
elseif tplanc="H" then
	if primeira=0 then 
		valord=formatdatetime((cdbl(valor)/60)/24,4)
	else
		valor1=replace(valor,":",",")
		valor1=replace(valor1,".",",")
		valor1=valor1
		hora=int(valor1)
		minuto=cint((valor1-hora)*100)
		if minuto>60 then minuto2=minuto:minuto=minuto2-(int(minuto2/60)*60):hora=hora+int(minuto2/60)
		valord=hora & ":" & numzero(minuto,2)
	end if
elseif tplanc="V" then
	if request.form("tipolanc")="H" then valor=hour(valor)+(minute(valor)/100)
	valord=formatnumber(valor,2)
else
	valord=valor
end if
if request.form("nrovezes")<>"" then nrovezes=request.form("nrovezes") else nrovezes=1
%>
<input type="hidden" name="tipolanc" size="5" value="<%=tplanc%>">
	<td class=fundo><%=tplanc%></td>
	<td class=fundo><input type="text" name="valor" class=vr size="8" value="<%=valord%>" ></td>
	<td class=fundo><input type="text" name="nrovezes" class=vr size="8" value="<%=formatnumber(nrovezes,0)%>" ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
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
'end if   'request.form=""
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;</script>"
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