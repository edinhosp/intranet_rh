<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a47")="N" or session("a47")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Lançamentos Folha</title>
<script language="javascript" type="text/javascript">
<!--
/****************************************************
     Author: Eric King
     Url: http://redrival.com/eak/index.shtml
     This script is free to use as long as this info is left in
     Featured on Dynamic Drive script library (http://www.dynamicdrive.com)
****************************************************/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body topmargin="5" leftmargin="5">
<%
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
'conexao.Open Application("consql")
conexao.open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")

if request.form("B2")<>"" then
 	Session("PrimeiraVez47")="Sim"

	if request.form("localizar")="" then session("loc47")="" else session("loc47")=request.form("localizar")

	if isnumeric(session("loc47"))=true then session("loc47")=numzero(session("loc47"),5)

	if session("loc47")<>"" then
		if isnumeric(session("loc47")) then
			session("sql47d")="AND (f.chapa like '%" & session("loc47") & "%') "
		else
			session("sql47d")="AND (f.nome like '%" & session("loc47") & "%') "
		end if
	else
		session("sql47d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")

sqla="SELECT f.chapa, f.nome, s.DESCRICAO, Max(cast(ano as char)+'/'+cast(mes as char)) AS ultima " & _
"FROM apont_adm c, corporerm.dbo.pfunc f, corporerm.dbo.pcodsituacao s " & _
"where c.chapa=f.chapa collate database_default and f.codsituacao=s.codcliente "
sqlc="GROUP BY f.chapa, f.nome, descricao order by f.nome "
sql47=sqla & session("sql47d") & sqlc
	
if Session("PrimeiraVez47")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql47, ,adOpenStatic, adLockReadOnly
	Session("Pagina")=1
	MostraDados
	Session("PrimeiraVez47")="Nao"
else
	if request("folha")="" then pagina=1
	if request.form("pagina")<>"" then pagina=request.form("pagina")
	if request("folha")<>"" then pagina=request("folha")
	Session("Pagina")=pagina
	conexao.cursorlocation = 3 'aduseclient
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql47, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 	MostraDados
end if	
	
Sub MostraDados()
	if rs.recordcount>0 then 	rs.AbsolutePage=Session("Pagina") 'vai para o número da pagina armazenado na sessão
End Sub
%>
<form method="POST" name="form" action="apontadm.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Lançamentos em Folha - Administrativos</p>
<table border="0" width="650" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="55%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""apontadm.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""apontadm.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
end if

response.write "&nbsp;<b>"
response.write "<select size='1' name='pagina' onchange='javascript:submit()'>"
for selpag=1 to rs.pagecount
	if selpag=atual then selpag1="selected" else selpag1=""
	response.write "<option value=" & selpag & " " & selpag1 & ">" & selpag & "</option>"
next
response.write "</select>"
response.write "/" & rs.pagecount & "</b>&nbsp;"

if atual=rs.pagecount or rs.pagecount=0 then
response.write "<img src='../images/setanext0.gif' border='0'>"
response.write "<img src='../images/setalast0.gif' border='0'>"
else
response.write "<a href=""apontadm.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""apontadm.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
    <td class=campo width="30%" valign="top" align="center">
<% if session("a47")="T" then %>
<a href="apont_nova.asp?codigo=<%=chapa%>" onclick="NewWindow(this.href,'InclusaoFolha','420','200','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo teto">
<font size="1">inserir novo lançamento</font></a>
<% end if %>

	</td>
    <td class=campo width="15%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="650" cellpadding="0" style="border-collapse: collapse">
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="center">Nome</td>
	<td class=titulo align="center">Situação</td>
	<td class=titulo align="center">Ultima</td>
</tr>
<%
linha=1
'rs.movefirst
'do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to Session("RegistrosPorPagina")
%>
<tr>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    <a class=r href="apontadm_ver.asp?chapa=<%=rs("chapa")%>&nome=<%=rs("nome")%>" onclick="NewWindow(this.href,'controlelanc','600','430','yes','center');return false" onfocus="this.blur()">
	&nbsp;<%=rs("chapa")%></a></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    &nbsp;<%=rs("nome")%></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%>>
    <%=rs("descricao")%></td>
	<td class=campo <%if linha=0 then response.write "bgcolor='#FFFFCC'"%> align="center">
    &nbsp;<%=rs("ultima")%></td>
</tr>
<%
if linha=1 then linha=0 else linha=1
rs.movenext
if rs.eof then exit for
'loop
Next
%>
</table>
<%
else 'sem registros
%>
<p><b><font color="#FF0000">Esta seleção não mostra nenhum registro.</font></b></p>
<%
end if 'sem registros
%>
<hr>
<br>
Localizar por nome/chapa: <input type="text" name="localizar" size=25 value="<%=session("loc47")%>">
Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
</font>
<input name="B2" type="submit" class="button" value="Clique para Filtrar"></p>
</form>
<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>