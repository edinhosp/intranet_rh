<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a56")="N" or session("a56")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Convênio com IES - Candidatos Recebidos</title>
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
<body>
<%
dim conexao, conexao2, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
'conexao.Open Application("consql")

if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"

	if request.form("status")="" then session("sel56")="Todos" else session("sel56")=request.form("status")
	if request.form("faculdade")="" then session("emp56")="Todas" else session("emp56")=request.form("faculdade")
	if request.form("perlet")="" then session("per56")="Todos" else session("per56")=request.form("perlet")

	if request.form("localizar")="" then session("loc56")="" else session("loc56")=request.form("localizar")
		
	if isnumeric(session("loc56"))=true then session("loc56")=session("loc56")

	if session("sel56")<>"Todos" then
		session("sql56b")="AND (c.bolsa=" & session("sel56") & ") "
	else
		session("sql56b")=""
	end if

	if session("emp56")<>"0" then
		session("sql56c")="AND (f.id_faculdade=" & session("emp56") & ") "
	else
		session("sql56c")=""
	end if

	if session("per56")<>"Todos" then
		session("sql56e")="AND (c.perlet='" & session("per56") & "') "
	else
		session("sql56e")=""
	end if

	if session("loc56")<>"" then
		if isnumeric(session("loc56")) then
			session("sql56d")="AND (c.inscricao like '%" & session("loc56") & "%') "
		else
			session("sql56d")="AND (c.nome like '%" & session("loc56") & "%') "
		end if
	else
		session("sql56d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")

sqla="SELECT c.*, f.faculdade " & _
"FROM rhconveniadosfac c, rhconveniobe f " & _
"WHERE c.id>0 and f.id_faculdade=c.id_faculdade "
sqlb=""
sqlc="ORDER BY c.nome, c.perlet "

sql1=sqla & sqlb & session("sql56b") & session("sql56d") & session("sql56c") & sqlc
sql1=sqla & sqlb & session("sql56b") & session("sql56d") & session("sql56c") & session("sql56e") & sqlc

if Session("PrimeiraVez")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	conexao.open Application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	Session("Pagina")=1
	MostraDados
	Session("PrimeiraVez")="Nao"
else
	if request("folha")="" then pagina=1
	if request.form("pagina")<>"" then pagina=request.form("pagina")
	if request("folha")<>"" then pagina=request("folha")
	Session("Pagina")=pagina
	conexao.cursorlocation = 3 'aduseclient
	conexao.open Application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 	MostraDados
end if	

Sub MostraDados()
	if rs.recordcount>0 then 	rs.AbsolutePage=Session("Pagina") 'vai para o número da pagina armazenado na sessão
End Sub
%>
<p class=titulo style="margin-top: 0; margin-bottom: 0">Convênio IES - Candidatos Recebidos</p>
<form method="POST" name="form" action="recebidos.asp">
<table border="0" width="690" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="70%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
	response.write "<img src='../images/setafirst0.gif' border='0'>"
	response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
	response.write "<a href=""recebidos.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
	response.write "<a href=""recebidos.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
end if

response.write "&nbsp;<b>"
response.write "<select size='1' name='pagina' onChange='javascript:submit()'>"
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
	response.write "<a href=""recebidos.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
	response.write "<a href=""recebidos.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
	<td class=campo width="30%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="690" cellpadding="2" style="border-collapse: collapse">
<tr>
	<td class=titulor align="center">Nome do candidato</td>
	<td class=titulor align="center">Faculdade   </td>
	<td class=titulor align="center">Curso       </td>
	<td class=titulor align="center">Matrícula   </td>
	<td class=titulor align="center">Entrada     </td>
	<td class=titulor align="center">Status      </td>
	<td class=titulor align="center">Bolsa      </td>
	<td class=titulor align="center">Situação    </td>
	<td class=titulor align="center">A           </td>
</tr>
<%
linha=1
'rs.movefirst:do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
%>
<tr>
	<td class="campor">
		<a class=r href="recebidosver.asp?codigo=<%=rs("id")%>" onclick="NewWindow(this.href,'Pesquisa_ver','595','500','yes','center');return false" onfocus="this.blur()">
		<%=rs("nome")%></a></td>
	<td class="campor"><%=rs("faculdade")%></td>
	<td class="campor"><%=rs("curso") %></td>
	<td class="campor"><%=rs("inscricao") %></td>
	<td class="campor"><%=rs("perlet") %></td>
	<td class="campor"><%=rs("status")%></td>
	<td class="campor" align="center"><%if rs("bolsa")=0 then response.write "<img src='../images/bullet.gif'>" else response.write "<img src='../images/bullet_hl.gif'>"%></td>
	<td class="campor"><%%></td>
	<td class="campor">
<% if session("a56")="T" then %>
	<a href="recebidos_alteracao.asp?codigo=<%=rs("id")%>" onclick="NewWindow(this.href,'Requisicao_alterar','635','355','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Write.gif" border="0" height=14 alt="Clique para alterar"></a>
<% end if %>
	</td>
</tr>
<%
if linha=1 then linha=0 else linha=1
rs.movenext
if rs.eof then exit for
'loop
Next

else 'sem registros
%>
<td class=grupo colspan=9>Esta seleção não mostra nenhum registro.</td>
<%
end if

set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
%>
</table>
<% if session("a56")="T" then %>
<a href="recebidos_nova.asp" onclick="NewWindow(this.href,'candidato_nova','440','340','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
<font size="1">inserir novo candidato</font></a><br>
<% end if %>

<font size="1">
<%
'sql2="SELECT * from ifip_wstatus"
'rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Status: <select size="1" name="status">
	<option value="Todos" <%if session("sel56")="Todos" then response.write "selected"%>>Todos Status</option>
<%
'rs2.movefirst:do while not rs2.eof
%>
	<option value="-1" <%if session("sel56")="-1" then response.write "selected"%> >Bolsistas</td>
	<option value="0" <%if session("sel56")="0" then response.write "selected"%> >Sem bolsa</td>
<%
'rs2.movenext:loop
'rs2.close
%>
</select>
&nbsp;&nbsp;&nbsp;
<%
sql2="select faculdade, id_faculdade from rhconveniobe order by faculdade"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Faculdade: <select size="1" name="faculdade">
	<option value="0" <%if session("emp56")="0" then response.write "selected"%>>Todas Faculdades</option>
<%
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("id_faculdade")%>" <%if session("emp56")=cstr(rs2("id_faculdade")) then response.write "selected"%>><%=rs2("faculdade")%></option>
<%
rs2.movenext:loop
end if
rs2.close
%>
</select>
&nbsp;&nbsp;&nbsp;
<%
sql2="select perlet from rhconveniadosfac group by perlet"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Período: <select size="1" name="perlet">
	<option value="Todos" <%if session("per56")="Todos" then response.write "selected"%>>Todos Periodos</option>
<%
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("perlet")%>" <%if session("per56")=rs2("perlet") then response.write "selected"%>><%=rs2("perlet")%></option>
<%
rs2.movenext:loop
end if
rs2.close
%>
</select>

<br>
Localizar por candidato: <input type="text" name="localizar" size=35 value="<%=session("loc56")%>">
Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
</font>
<input name="B2" type="submit" class="button" value="Clique para Filtrar">
</form>

</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>