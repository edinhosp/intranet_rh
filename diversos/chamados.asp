<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Chamados Ouvidoria</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<%

dim conexao, conexao2, chapach
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

ipacesso=Request.ServerVariables("REMOTE_ADDR")
if ipacesso="10.0.1.91" or ipacesso="10.0.1.91" or ipacesso="127.0.0.1" then
if request("chapa")="" then chapa=" and f.codsituacao<>'D' " else chapa=" and chapa='" & request("chapa") & "' "
%>
<p class=titulo>Elogios/Críticas a Funcionários
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=690>
<tr>
	<td class=titulop align="center">Chapa</td>
	<td class=titulop align="center">Nome</td>
	<td class=titulop align="center">Data</td>
	<td class=titulop align="center">Texto</td>
	<td class=titulop align="center">Chamado</td>
</tr>
<%
sql="select f.codpessoa, f.chapa, f.nome, a.texto, a.dtanotacao, a.tipo from corporerm.dbo.panotac a, corporerm.dbo.pfunc f " & _
"where a.codpessoa=f.codpessoa and (f.chapa<'10000' or f.chapa>'90000') and a.tipo in (19,20) " & chapa & _
"order by f.nome, f.chapa, a.dtanotacao "

rsc.Open sql, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
chamado=trim(right(rsc("texto"),10))
chamado2=""
for a=1 to len(chamado)
	letra=mid(chamado,a,1)
	if isnumeric(letra)=true then chamado2=chamado2 & letra
next
chamado=chamado2
sqll="select count(codpessoa) as total from corporerm.dbo.panotac where codpessoa=" & rsc("codpessoa") & " and tipo in (19,20) "
rs.Open sqll, ,adOpenStatic, adLockReadOnly
linhas=rs("total")
rs.close
%>
<tr>
<%
if rsc("chapa")=lastchapa then
else
%>
	<td class=campo align="center" style="border-top:2 solid #000000" rowspan=<%=linhas%>><%=rsc("chapa")%></td>
	<td class=campo align="left" style="border-top:2 solid #000000" rowspan=<%=linhas%>><%=rsc("nome")%></td>
<%end if%>
	<td class=campo align="center" style="border-top:2 solid #000000"><%=rsc("dtanotacao")%></td>
	<td class=campo align="left"><%=rsc("texto")%></td>
	<td class=campo><a href="http://intranet.unifieo.br/legado/intranet/ouvidoria/visualiza.php?chamado=<%=chamado%>" target="_blank"><%=chamado%></a></td>
</tr>
<%
lastchapa=rsc("chapa")
rsc.movenext
loop
rsc.close
%>
</table>
<%

end if 'ipacesso
%>
</body>
</html>
<%
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%>