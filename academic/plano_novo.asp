<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a93")="N" or session("a93")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>UNIFIEO - Plano de Ensino</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1

if request.form("justificativa")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe a Justificativa!');</script>"
if request.form("ementa")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe a Ementa!');</script>"
if request.form("objetivos_gerais")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe os Objetivos Gerais!');</script>"

		sql="UPDATE grades_plano SET "
		sql=sql & "justificativa     ='" & request.form("justificativa")      & "', "
		sql=sql & "ementa            ='" & request.form("ementa")             & "', "
		sql=sql & "objetivos_gerais  ='" & request.form("objetivos_gerais")   & "', "
		sql=sql & "unidades_tematicas='" & request.form("unidades_tematicas") & "', "
		sql=sql & "metodologia       ='" & request.form("metodologia")        & "', "
		sql=sql & "avaliacao         ='" & request.form("avaliacao")          & "', "
		'sql=sql & "bibliografia      ='" & request.form("bibliografia")       & "', "
		'sql=sql & "bibliografiac     ='" & request.form("bibliografiac")       & "', "
		sql=sql & "novo=0, pa=0, "
		sql=sql & "usuarioc          ='" & session("usuariomaster")           & "', "
		sql=sql & "datac=getdate() "
		sql=sql & " WHERE id_plano=" & request.form("id_plano") & " "		
		'response.write "<font size='2'>" & sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if request("codigo")<>"" then id_plano=request("codigo")
if request.form("id_plano")<>"" then id_plano=request.form("id_plano")

if request.form="" or (request.form<>"" and tudook=0) then
sql="select p.*, m.materia from grades_plano p, corporerm.dbo.umaterias m where m.codmat collate database_default=p.codmat and id_plano=" & id_plano
rs.Open sql, ,adOpenStatic, adLockReadOnly

%>
<form method="POST" action="plano_novo.asp" name="form" >
<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=grupo>Inclusão de Plano de Ensino</td></tr>
</table>
<input type="hidden" name="id_plano" value="<%=id_plano%>">

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=fundo>Disciplina</td>
	<td class=fundop><%=rs("codmat")%> - <b><%=rs("materia")%></b> da Grade <%=rs("grade")%> no curso <%=rs("coddoc")%>, Período Letivo <%=rs("perlet")%>.
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">JUSTIFICATIVA</font></b></td></tr>
<%
len1=int(len(request.form("justificativa"))/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len1%>" name="justificativa" cols="80" style="background-color: #FFFFCC"><%=request.form("justificativa")%></textarea>
	</td>
</tr>

<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">EMENTA</font></b></td></tr>
<%
len2=int(len(request.form("ementa"))/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len2%>" name="ementa" cols="80" style="background-color: #FFFFCC"><%=request.form("ementa")%></textarea>
	</td>
</tr>

<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">OBJETIVOS GERAIS</font></b></td></tr>
<%
len3=int(len(request.form("objetivos_gerais"))/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len3%>" name="objetivos_gerais" cols="80" style="background-color: #FFFFCC"><%=request.form("objetivos_gerais")%></textarea>
	</td>
</tr>

<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">UNIDADES TEMÁTICAS</font></b></td></tr>
<%
len4=int(len(request.form("unidades_tematicas"))/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len4%>" name="unidades_tematicas" cols="80" style="background-color: #FFFFCC"><%=request.form("unidades_tematicas")%></textarea>
	</td>
</tr>

<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">METODOLOGIA</font></b></td></tr>
<%
len5=int(len(request.form("metodologia"))/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len5%>" name="metodologia" cols="80" style="background-color: #FFFFCC"><%=request.form("metodologia")%></textarea>
	</td>
</tr>

<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">AVALIAÇÃO</font></b></td></tr>
<%
len6=int(len(request.form("avaliacao"))/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len6%>" name="avaliacao" cols="80" style="background-color: #FFFFCC"><%=request.form("avaliacao")%></textarea>
	</td>
</tr>

<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">BIBLIOGRAFIA BÁSICA/COMPLEMENTAR</font></b></td></tr>
<tr>
	<td class=titulo>Para inserir a bibliografia salve o plano de ensino.
	</td>
</tr>


</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar" class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
rs.close

else
set rs=nothing
end if
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		'response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.document.form.submit();self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if

'	Response.write "<p>Registro salvo.<br>"
	'response.write '<script>javascript:top.window.close();</script>
%>
<script language="Javascript">javascript:window.opener.document.form.submit()</script>
<!-- <script language="Javascript">javascript:window.opener.location.refresh</script>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script> -->
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