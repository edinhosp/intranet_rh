<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a64")="N" or session("a64")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Convênio de Bolsas</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1()	{	form.chapa.value=form.nome.value;	}
function chapa1()	{	form.nome.value=form.chapa.value;	}
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, ok, varpro(2), varper(2)
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
		sqla = "INSERT INTO bolsistas_lanc (id_bolsa, situacao, ano_letivo, renovacao, validade, " & _
		"observacao, id_faculdade, curso, periodo, protocolo )"
		sqlb = " SELECT " & request.form("id_bolsa") & " " & _
		",'" & request.form("situacao") & "' " & _
		",'" & request.form("ano_letivo") & "' "
		if request.form("renovacao")="" then sqlb=sqlb & ",null" else sqlb=sqlb & ",'" & dtaccess(request.form("renovacao")) & "'"
		if request.form("validade")="" then sqlb=sqlb & ",null" else sqlb=sqlb & ",'" & dtaccess(request.form("validade")) & "'"
		sqlb=sqlb & ",'" & request.form("observacao") & "' " & _
		", " & request.form("id_faculdade") & " " & _
		",'" & request.form("curso") & "' " & _
		",'" & request.form("periodo") & "' " & _
		", " & request.form("protocolo") & " " 
		'sqlb=sqlb & ",'" & session("usuariomaster") & "'"
		'sqlb=sqlb & ",now()"
		sql = sqla & sqlb
		'response.write sql
		if tudook=1 then 
			conexao.Execute sql, , adCmdText
			id_bolsa1=request.form("id_bolsa")
			sql2="select faculdade from rhconveniobe where id_faculdade=" & request.form("id_faculdade")
			rs.Open sql2, ,adOpenStatic, adLockReadOnly
			if rs.recordcount>0 then faculdade=rs("faculdade") else faculdade="-"
			rs.close
		sql1="select top 1 renovacao from bolsistas_lanc where id_bolsa=" & id_bolsa1 & " order by renovacao desc"
		rsc.Open sql1, ,adOpenStatic, adLockReadOnly
		if rsc.recordcount>0 then
			ultimolanc=rsc("renovacao")
			if cdate(ultimolanc)=cdate(request.form("renovacao")) then
			sql2="update bolsistas set situacao='" & request.form("situacao") & "', instituicao='" & faculdade & "', curso='" & request.form("curso") & "' where id_bolsa=" & id_bolsa1
			conexao.Execute sql2, , adCmdText
			end if
		end if
		rsc.close
			
		end if
	end if 'request btsalvar
else 'request.form=""
end if

'if request.form("bt_salvar")="" then
if (request.form("bt_salvar")="") or (request.form("bt_salvar")<>"" and tudook=0) then

situacao=request.form("situacao"):if situacao="" then situacao="M"
ano_letivo=request.form("ano_letivo"):if ano_letivo="" then ano_letivo=year(now)&"/"
renovacao=request.form("renovacao"): if renovacao="" then renovacao=formatdatetime(now(),2)
validade=request.form("validade")
observacao=request.form("observacao")
protocolo=request.form("protocolo"): if protocolo="" then protocolo=session("64protocolo")
faculdade=request.form("id_faculdade"): if faculdade="" then faculdade=session("64faculdade")
curso=request.form("curso"): if curso="" then curso=session("64curso")
periodo=request.form("periodo"): if periodo="" then periodo=session("64periodo")

if request("codigo")="" or isnull(request("codigo")) then
	codigo=request.form("id_bolsa")
else
	codigo=request("codigo")
end if

%>
<form method="POST" action="convenio_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<input type="hidden" name="id_bolsa" size="4" value="<%=codigo%>" >  
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
	<tr><td class=grupo>Inclusão de Convênio de Bolsas</td></tr>
</table>
<%
'for each strItem in Request.form
'	Response.write stritem & " = " & request.form(stritem) & " "
'next
%>
<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Data Renovação</td>
	<td class=titulo>Data Validade</td>
	<td class=titulo>Situação do Lançamento</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="renovacao" value="<%=renovacao%>" size="10"></td>
	<td class=fundo><input type="text" name="validade" value="<%=validade%>" size="10"></td>
	<td class=fundo><select size="1" name="situacao">
<%
sqla="SELECT * from bolsistas_situacao ORDER by descricao"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if rsc("id_sit")=situacao then temps="selected" else temps=""
%>
		<option value="<%=rsc("id_sit")%>" <%=temps%>><%=rsc("descricao")%></option>
<%
rsc.movenext
loop
rsc.close
%>
	</select></td>
</tr>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Periodo Letivo</td>
	<td class=titulo>Observação</td>
	<td class=titulo>Tipo Emissão</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="ano_letivo"  size="15" value="<%=ano_letivo%>" class=a></td>
	<td class=fundo><input type="text" name="observacao"  size="30" value="<%=observacao%>"></td>
	<td class=fundo><select size="1" name="protocolo" class=a>
          <option value="-1">Selecione</option>
<%
varpro(0)="Inscr.Vestibular"
varpro(1)="Matrícula"
varpro(2)="Rematrícula"
for a=0 to 2
	if cdbl(protocolo)=cdbl(a+1) then tempp="selected" else tempp=""
%>
	<option value="<%=a+1%>" <%=tempp%>><%=varpro(a)%></option>
<%
next
%>
	</select></td>
</tr>
</table>
<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>Faculdade</td>
	<td class=titulo>Período</td>
</tr>
<tr>
	<td class=fundo>
		<select size="1" name="id_faculdade" onchange="javascript:submit()">
		<option value="0">Selecione uma faculdade</option>
<%
sql2="select id_faculdade, faculdade from rhconveniobe order by faculdade "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if cint(faculdade)=cint(rsc("id_faculdade")) then temp1="selected" else temp1=""
%>
		<option value="<%=rsc("id_faculdade")%>" <%=temp1%>><%=rsc("faculdade")%></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select></td>
	<td class=fundo><select size="1" name="periodo">
          <option value=""></option>
<%
varper(0)="Matutino"
varper(1)="Vespertino"
varper(2)="Noturno"
for a=0 to 2
	if periodo=varper(a) then tempp="selected" else tempp=""
%>
	<option value="<%=varper(a)%>" <%=tempp%>><%=varper(a)%></option>
<%
next
%>
	</select></td>
</tr>	
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>Curso</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="curso">
		<option value="0">Selecione um curso</option>
<%
if faculdade="" then faculdade=0
sql2="select cursos, id_curso from rhconveniobec where id_faculdade=" & faculdade & " order by cursos "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then 
rsc.movefirst:do while not rsc.eof
if curso=rsc("cursos") then temp1="selected" else temp1=""
%>
		<option value="<%=rsc("cursos")%>" <%=temp1%>><%=rsc("cursos")%></option>
<%
rsc.movenext:loop
end if 'recordcount
rsc.close
%>
	</select>	
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
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
set rsc=nothing
set rds=nothing
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if
%>
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