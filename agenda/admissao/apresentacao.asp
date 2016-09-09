<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")="N" or session("a48")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Emissão de Apresentação (Mini-Curriculum)</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"><!--
function nome1() {	form.chapa.value=form.nome.value; }
function chapa1() {	form.nome.value=form.chapa.value; }
--></script>
</head>
<body style="margin-left:20px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
sessao=session.sessionid
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

espacamento=5
if request.form="" then
sql="select p.chapa, p.nome, p.dataadmissao from corporerm.dbo.pfunc p where p.chapa<'10000' and p.codtipo='N' and codsituacao<>'D' order by p.dataadmissao, p.chapa, p.nome "
sql="select CHAPA, NOME, Tipo, Data from ( " & _
"Select p.chapa, p.nome, 'Tipo'='E', 'Data'=p.dataadmissao from corporerm.dbo.pfunc p where p.chapa<'10000' and p.codtipo='N' and codsituacao<>'D' " & _
"union " & _
"Select p.chapa, p.nome, 'Tipo'='S', 'Data'=p.DATADEMISSAO from corporerm.dbo.pfunc p where p.chapa<'10000' and p.codtipo='N' and codsituacao='D' " & _
") p order by p.Data desc, p.chapa, p.nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form name="form" action="apresentacao.asp" method="post">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Seleção de Funcionário para emissão de Apresentação (Mini-Curriculum)</td>
</tr>
<tr>
	<td class=campo>Funcionário</td>
	<td class=campo>
		<select name="chapa" class=a size=20 multiple>
		<option value="0"> Selecione o funcionário</option>
<%
rs.movefirst
do while not rs.eof
%>
		<option value="<%=rs("chapa")%>"> <%=rs("tipo") & " - " & rs("chapa") & " - " & rs("nome") & " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(" & rs("data") & ")"%></option>
<%
rs.movenext
loop
rs.close
%>
		</select>
	</td>
</tr>
<tr>
	<td class=campo colspan=3>&nbsp;
	<input type="radio" name="ordem" value="D">Por ordem de data
	<input type="radio" name="ordem" value="T">Por ordem de tipo
	</td>
</tr>
<tr>
	<td class=campo colspan=3>&nbsp;
		<input type="submit" value="Visualizar" class=button name="B1">
	</td>
</tr>
</table>
</form>

<%
else
'response.write request.form("chapa").count
chapas=request.form("chapa").count
coluna=1
if request.form("ordem")="T" then ordemsql=" order by case when codsituacao='D' then 'S' else 'E' end" else ordemsql=" order by case when codsituacao='D' then datademissao else dataadmissao end "
%>
<table border="0" cellpadding="1" cellspacing="0" width="930" height=470 style="border-collapse: collapse">
<%
'****
for a=1 to chapas
'****
if coluna=1 then response.write "<tr><td>"
if culuna=2 then response.write "</td><td>"
chapa=request.form("chapa").item(a)
sql="select f.chapa, f.nome, f.dataadmissao, f.datademissao, p.dtnascimento, p.bairro, p.cidade, s.descricao, p.sexo, p.apelido, f.codsecao, f.codsindicato, f.codsituacao, f.tipodemissao " & _
", 'ordem'=case when codsituacao='D' then 'S' else 'E' end, 'data'=case when codsituacao='D' then datademissao else dataadmissao end " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.psecao s " & _
"where f.codpessoa=p.codigo and f.codsecao=s.codigo and f.chapa='" & chapa & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs("sexo")="F" then s1="a" else s1="o"
if rs("sexo")="F" then s2="a" else s2=""
if rs("codsindicato")="03" then
	titulo="Prof"
else
	titulo="Sr"
end if
sufixo=" no "
cs=rs("codsecao"):ds=rs("descricao")
if cs="01.1.002" or cs="03.1.002" or cs="01.1.099" or cs="01.1.997" or cs="03.1.998" or cs="01.1.995" or cs="03.1.996" or cs="03.1.999" or cs="01.1.007" or cs="03.1.007" then sufixo=" em "
if cs="03.2.042" or cs="01.2.005" or cs="03.2.005" or cs="01.2.006" or cs="03.2.006" or cs="03.2.011" or cs="03.1.021" or cs="04.2.332" or cs="03.2.332" or cs="04.2.331" or cs="01.1.013" or cs="03.1.013" or cs="01.1.003" or cs="03.1.003" or cs="01.2.003" or cs="03.2.003" or cs="01.1.005" or cs="03.1.005" or cs="01.2.008" or cs="03.2.008" or cs="01.1.006" or cs="03.1.006" or cs="03.1.022" or cs="03.1.097" or cs="03.2.200" or cs="04.1.008" or cs="01.1.008" or cs="03.1.008" or cs="03.2.104" or cs="03.2.102" or cs="03.2.108" or cs="03.2.106" or cs="04.1.020" or cs="03.1.020" or cs="01.2.001" or cs="03.2.001" or cs="03.2.100" or cs="01.2.059" or cs="01.2.055" or cs="03.2.055" or cs="03.2.067" or cs="01.2.050" or cs="03.2.071" or cs="03.2.070" or cs="03.2.074" or cs="01.2.120" or cs="03.2.065" or cs="01.2.051" or cs="03.2.051" or cs="04.2.250" or cs="03.2.250" or cs="01.2.021" or cs="01.2.010" or cs="03.2.010" or cs="01.2.007" or cs="03.2.007" or cs="03.2.105" or cs="03.2.103" or cs="03.2.109" or cs="03.2.107" or cs="01.2.002" or cs="03.2.002" or cs="03.2.101" or cs="01.2.004" or cs="03.2.004" or cs="01.1.012" or cs="03.1.012" then sufixo=" na "
if rs("codsituacao")="D" then
	textocab="Despedida" 
	tipoap="S"
else 
	textocab="Apresentação"
	tipoap="E"
end if
if rs("tipodemissao")="4" then textosaida=" por vontade própria." else textosaida="."
%>
<table border="0" cellpadding="1" cellspacing="0" width="450" height=220 style="border-collapse: collapse">
<tr><td colspan=2 class=titulop align="center" style="border: 1px solid #000000;border-bottom:3 double #000000">
	<font size="3"><%=textocab%>
	</td>
</tr>

<tr><td class=campo valign="top" style="border-left: 1px solid #000000;border-bottom: 1px solid #000000">
<img border="0" src="../func_foto.asp?chapa=<%=rs("chapa")%>"  height="150px" width="112px">
</td><td class="campop" valign="top" style="border-bottom: 1px solid #000000;border-right: 1px solid #000000">
<p style="margin-top:0;margin-bottom:0;text-align:justify">
<!--
Em <%=rs("dataadmissao")%>, ingressou em nosso quadro de funcionários, <%=s1%> Sr<%=s2%>. <%=rs("nome")%>,
que desenvolverá suas atividades <%=sufixo%> <%=rs("descricao")%>.
-->
<%if tipoap="E" then%>
Em <%=rs("dataadmissao")%>, ingressou <%=s1%>&nbsp;<%=titulo%><%=s2%>. <%=rs("nome")%> (<%=rs("descricao")%>).

<%else%>
Em <%=rs("datademissao")%>, <%=s1%>&nbsp;<%=titulo%><%=s2%>. <%=rs("nome")%>, que estava desde <%=rs("dataadmissao")%> em <%=rs("descricao")%> saiu do quadro de funcionários<%=textosaida%>

<%end if%>

<br>
<br>
<%
if rs("bairro")="" or isnull(rs("bairro")) then bairro="" else bairro=", no bairro " & rs("bairro") & ""
%>
<!--
<%=ucase(s1)%> Sr<%=s2%>. <%=rs("apelido")%> reside em <%=rs("cidade")%><%=bairro%>.
-->
Reside em <%=rs("cidade")%>.

<br>
<br>
<!--
Desejamos-lhe sucesso nesta nova etapa de sua carreira.
Sucesso nesta etapa de sua carreira.
-->
</td>
</tr>
</table>
<%
rs.close
if coluna=1 then response.write "</td><td>"
if coluna=2 then response.write "</td></tr>"
'if int(a/2)*2=a then response.write "</tr>"
if coluna=1 then coluna=2 else coluna=1

if a/6-int(a/6)=0 then 
	response.write "</table>"
	response.write "<DIV style=""page-break-after:always""></DIV>"
	response.write "<table border=0 cellpadding=1 cellspacing=0 width=930 height=470 style='border-collapse:collapse'>"
end if
'****
next
'****
response.write "</table>"
end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>