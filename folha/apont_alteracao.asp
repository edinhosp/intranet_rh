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
<title>Alteração de Lançamento em Folha</title>
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
<script src="../coolmenu/coolmenus_frame.js" type="text/javascript"></script>
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
		sql="UPDATE apont_adm SET "
		sql=sql & "ano        ='" & request.form("ano") & "' "
		sql=sql & ",mes       = " & request.form("mes")
		sql=sql & ",chapa     ='" & request.form("chapa") & "' "
		sql=sql & ",codevento ='" & request.form("codevento") & "' "
		sql=sql & ",valor     =" & nraccess(valor)
		sql=sql & ",nrovezes  =" & nraccess(request.form("nrovezes"))
		'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
		'sql=sql & ",dataa   =getdate() "
		sql=sql & " WHERE id_adm=" & session("id_alt_adm")
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="UPDATE apont_adm set deletada=-1 WHERE id_adm=" & session("id_alt_adm")
		sql="DELETE FROM apont_adm WHERE id_adm=" & session("id_alt_adm")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
	if request("codigo")="" then
		id_adm=session("id_alt_adm")
		id_adm=request.form("id_adm")
	else
		id_adm=request("codigo")
	end if
	sql="select * from apont_adm where id_adm=" & id_adm
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
session("id_alt_adm")=rs("id_adm")
sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
'response.write request.form
%>
<form method="POST" action="apont_alteracao.asp" name="form">
<input type="hidden" name="id_adm" size="4" value="<%=rs("id_adm")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
	<tr><td class=grupo>Alteração de Lançamento em Folha <%=rs("id_adm")%></td></tr>
</table>

<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
<input type="hidden" name="chapa" size="5" value="<%=rs("chapa")%>">
	<td class=fundo><%=rs("id_adm")%></td>
	<td class=fundo><%=rs("chapa")%></td>
	<td class=fundo><%=rsnome("nome")%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo>Ano</td>
	<td class=titulo>Mês</td>
	<td class=titulo>Evento</td>
</tr>
<tr>
<%if request.form("ano")<>"" then ano=request.form("ano") else ano=rs("ano")%>
<%if request.form("mes")<>"" then mes=request.form("mes") else mes=rs("mes")%>
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
<%if request.form("codevento")<>"" then codevento=request.form("codevento") else codevento=rs("codevento")%>
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
tplanc=rsd("valhordiaref")
tpeve=rsd("provdescbase")
fator=rsd("fator")
rsd.close
if request.form("valor")<>"" then 
	valor=request.form("valor"):primeira=1
else
	valor=rs("valor"):primeira=0
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
		minuto=(valor1-hora)*100
		if minuto>60 then minuto2=minuto:minuto=minuto2-(int(minuto2/60)*60):hora=hora+int(minuto2/60)
		valord=hora & ":" & numzero(minuto,2)
	end if
elseif tplanc="V" then
	if request.form("tipolanc")="H" then valor=hour(valor)+(minute(valor)/100)
	valord=formatnumber(valor,2)
else
	valord=valor
end if
if request.form("nrovezes")<>"" then nrovezes=request.form("nrovezes") else nrovezes=rs("nrovezes")
%>
<input type="hidden" name="tipolanc" size="5" value="<%=tplanc%>">
	<td class=fundo><%=tplanc%></td>
	<td class=fundo><input type="text" name="valor" class=vr size="8" value="<%=valord%>" ></td>
	<td class=fundo><input type="text" name="nrovezes" class=vr size="8" value="<%=formatnumber(nrovezes,0)%>" ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser alterado!');</script>"
	end if
	'Response.write "<p>Registro atualizado.<br>"
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<!--
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>