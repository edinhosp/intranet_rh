<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a46")="N" or session("a46")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Autorização de Extra para Atrasos</title>
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
	
if request.form="" then
%>
<p class=titulo>Compensação de Atrasos para Extras Executadas
<form method="POST" action="n1_compensa.asp">
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo>Verificar Atrasos e Extras</td></tr>
<%
hoje=int(now())
diasem=weekday(hoje)
d2=hoje - (diasem-1)
d1=d2-6
%>
<tr>
	<td class=titulo>de <input type="text" name="d1" value="<%=d1%>" size="9"> até <input type="text" name="d2" value="<%=d2%>" size="9"></td>
</tr>
<tr><td colspan=3 class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3">
</td></tr>
</table>
</form>
<hr>
<%
else 'request.form <>''
	datai=request.form("d1")
	dataf=request.form("d2")
	linha=0:pagina=0
	sqld="select h.CHAPA, f.NOME, h.DATA, h.ATRASO, h.EXTRAEXECUTADO, h.EXTRAAUTORIZADO " & _
"from corporerm.dbo.AAFHTFUN h inner join corporerm.dbo.PFUNC f on f.CHAPA=h.CHAPA " & _
"where h.DATA between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "' " & _
"and f.CODSINDICATO<>'03' and h.ATRASO>0 and h.EXTRAEXECUTADO>0 and h.atraso<>h.extraautorizado and h.extraautorizado<>h.extraexecutado " & _
"and h.chapa not in ('00099','00554','02297','02538','02653') " & _
"order by h.CHAPA, h.DATA "
rs.Open sqld, ,adOpenStatic, adLockReadOnly
totalpag=int(rs.recordcount/65)+1
do while not rs.eof
if linha=0 or linha>64 then
	if linha<>0 then
		pagina=pagina+1
		response.write "<tr><td class=""campor"" colspan=7 style='border-top:1px solid #000000'>Página " & pagina & "/" & totalpag & " - " & now() & "</td></tr>"
		response.write "</table>"
		response.write "<DIV style=""page-break-after:always""></DIV>"
	end if
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td class=titulo colspan=7 align="center">Relatório de Extras a Autorizar para compensar Atrasos - De <%=datai%> a <%=dataf%></td></td>
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Data</td>
	<td class=titulo>Atraso</td>
	<td class=titulo>Extra Exec.</td>
	<td class=titulo>Extra Aut.</td>
	<td class=titulo>Obs.</td>
</tr>
<%
	if linha<>0 then linha=0
end if 'linha
if rs("atraso")>rs("extraexecutado") then comp="T" else comp="P"
if comp="P" then obs="autorizar " & horaload(rs("atraso"),1) else obs=""
if rs("chapa")<>ultchapa then cab=1 else cab=0
'obs=rs.absoluteposition & "-" & obs 
%>
<tr>
<%if cab=1 then%>
	<td class=campo style="border-top:1px solid #000000"><%=rs("chapa")%></td>
	<td class=campo style="border-top:1px solid #000000"><%=rs("nome")%></td>
<%else%>
	<td class=campo colspan=2>&nbsp;</td>
<%
end if
if cab=1 then estilo="border-top:1px solid #000000" else estilo="border-top:0px solid #000000"
%>
	<td class=campo style="<%=estilo%>" ><%=rs("data")%></td>
	<td class=campo style="<%=estilo%>" align="center">&nbsp;<%=horaload(rs("atraso"),1)%></td>
	<td class=campo style="<%=estilo%>" align="center">&nbsp;<%=horaload(rs("extraexecutado"),1)%></td>
	<td class=campo style="<%=estilo%>" align="center">&nbsp;<%=horaload(rs("extraautorizado"),1)%></td>
	<td class=campo style="<%=estilo%>" ><%=obs%></td>
</tr>
<%
linha=linha+1
ultchapa=rs("chapa")
rs.movenext
loop
rs.close
pagina=pagina+1
%>
<tr><td class="campor" colspan=7 style='border-top:1px solid #000000'>Página <%=pagina & "/" & totalpag%> - <%=now()%></td></tr>
</table>

<%
end if ' request.form	
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>