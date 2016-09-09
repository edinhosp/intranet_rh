<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a30")="N" or session("a15")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Processos IFIP</title>
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
session("emp20")="Todas"
dim conexao, conexao2, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
'conexao.Open Application("consql")
conexao.open Application("conexao")

if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"

	if request.form("status")="" then session("sel20")="Todos" else session("sel20")=request.form("status")
	if request.form("pessoa")="" then session("emp20")="Todas" else session("emp20")=request.form("pessoa")

	if request.form("localizar")="" then session("loc20")="" else session("loc20")=request.form("localizar")
		
	if isnumeric(session("loc20"))=true then session("loc20")=session("loc20")

	if session("sel20")<>"Todos" then
		session("sql20b")="AND (i.status='" & session("sel20") & "') "
	else
		session("sql20b")=""
	end if

	if session("emp20")<>"Todas" then
		session("sql20c")="AND (t.chapa='" & session("emp20") & "') "
	else
		session("sql20c")=""
	end if

	if session("loc20")<>"" then
		if isnumeric(session("loc20")) then
			session("sql20d")="AND (i.num_processo like '%" & session("loc20") & "%') "
		else
			session("sql20d")="AND (i.titulo_pesquisa like '%" & session("loc20") & "%') "
		end if
	else
		session("sql20d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")

if session("emp20")<>"Todas" then	
	sqla="SELECT i.*, s.desc_status, t.chapa " & _
	"FROM (ifip_cadastro AS i LEFT JOIN ifip_wstatus s ON i.status=s.id_status) INNER JOIN ifip_titulares t ON i.id_ifip=t.id_ifip " & _
	"WHERE i.id_ifip>0 "
	sqla="SELECT i.*, s.desc_status, f.NOME, t.chapa " & _
	"FROM (ifip_cadastro AS i INNER JOIN ifip_wstatus AS s ON i.status=s.id_status) LEFT JOIN (corporerm.dbo.pfunc AS f RIGHT JOIN (select id_ifip, chapa from ifip_titulares where tp_docente in ('T','C') ) AS t ON f.CHAPA collate database_default=t.chapa) ON i.id_ifip=t.id_ifip " & _
	"WHERE i.id_ifip>0 "
else
	sqla="SELECT i.*, s.desc_status " & _
	"FROM ifip_cadastro AS i, ifip_wstatus s WHERE i.status=s.id_status " & _
	"AND i.id_ifip>0 "
	sqla="SELECT i.*, s.desc_status, f.nome " & _
	"FROM ifip_cadastro AS i, ifip_wstatus s, pfunc f, ifip_titulares t WHERE i.status=s.id_status and f.chapa=t.chapa and t.id_ifip=i.id_ifip " & _
	"AND i.id_ifip>0 "
	sqla="SELECT i.*, s.desc_status, f.NOME " & _
	"FROM (ifip_cadastro AS i INNER JOIN ifip_wstatus AS s ON i.status=s.id_status) LEFT JOIN (corporerm.dbo.pfunc AS f RIGHT JOIN (select id_ifip, chapa from ifip_titulares where tp_docente='T') AS t ON f.CHAPA collate database_default=t.chapa) ON i.id_ifip=t.id_ifip " & _
	"WHERE i.id_ifip>0 "
end if

sqlb=""
sqlc="ORDER BY i.titulo_pesquisa "
sqlc="ORDER BY i.num_processo "

sql1=sqla & sqlb & session("sql20b") & session("sql20d") & session("sql20c") & sqlc

if Session("PrimeiraVez")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
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
<p class=titulo style="margin-top: 0; margin-bottom: 0">Processos IFIP</p>
<form method="POST" name="form" action="pesquisa.asp">
<table border="0" width="690" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="70%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
	response.write "<img src='../images/setafirst0.gif' border='0'>"
	response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
	response.write "<a href=""pesquisa.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
	response.write "<a href=""pesquisa.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
	response.write "<a href=""pesquisa.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
	response.write "<a href=""pesquisa.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
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

<table border="1" cellspacing="0" width="690" cellpadding="2" style="border-collapse: collapse">
<tr>
    <td class=titulo align="center">Processo          </td>
    <td class=titulo align="center">Nome          </td>
    <td class=titulo align="center">Título da Pesquisa</td>
    <td class=titulo align="center">Status            </td>
    <td class=titulo align="center">Início            </td>
    <td class=titulo align="center">Término           </td>
    <td class=titulo align="center">Vigência          </td>
    <td class=titulo align="center">Horas              </td>
    <td class=titulo align="center">A           </td>
</tr>
<%
linha=1
'rs.movefirst:do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
%>
<tr>
	<td class=campo>
		<a class=r href="pesquisaver.asp?codigo=<%=rs("id_ifip")%>" onclick="NewWindow(this.href,'Pesquisa_ver','595','500','yes','center');return false" onfocus="this.blur()">
		<%=rs("num_processo")%></a></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo><%=rs("titulo_pesquisa")%></td>
	<td class=campo><%=rs("desc_status") %></td>
	<td class=campo align="center"><%=rs("dt_entrada") %></td>
	<td class=campo align="center"><%=rs("dt_termino") %></td>
	<td class=campo align="center"><%=rs("vigencia")%></td>
	<td class=campo align="center"><%=rs("horas_semanais")%></td>
	<td class=campo align="center">
<% if session("a30")="T" then %>
	<a href="pesquisa_alteracao.asp?codigo=<%=rs("id_ifip")%>" onclick="NewWindow(this.href,'Requisicao_alterar','635','355','yes','center');return false" onfocus="this.blur()">
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
<% if session("a30")="T" then %>
<a href="pesquisa_nova.asp" onclick="NewWindow(this.href,'pesquisa_nova','635','355','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
<font size="1">inserir novo processo</font></a><br>
<% end if %>

<font size="1">
<%
sql2="SELECT * from ifip_wstatus"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Status: <select size="1" name="status">
	<option value="Todos" <%if session("sel20")="Todos" then response.write "selected"%>>Todos Status</option>
<%
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("id_status")%>" <%if session("sel20")=rs2("id_status") then response.write "selected"%>><%=rs2("desc_status")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>
&nbsp;&nbsp;&nbsp;
<%
sql2="select i.chapa, f.nome from ifip_titulares i, corporerm.dbo.pfunc f where f.chapa collate database_default=i.chapa group by i.chapa, f.nome order by f.nome "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Pessoa: <select size="1" name="pessoa">
	<option value="Todas" <%if session("emp20")="Todas" then response.write "selected"%>>Todas Pessoas</option>
<%
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("chapa")%>" <%if session("emp20")=rs2("chapa") then response.write "selected"%>><%=rs2("chapa") & " - " & rs2("nome")%></option>
<%
rs2.movenext:loop
end if
rs2.close
%>
</select>
<br>
Localizar por descrição: <input type="text" name="localizar" size=35 value="<%=session("loc20")%>">
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