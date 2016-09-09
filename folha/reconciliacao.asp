<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")="N" or session("a48")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Reconciliação Contábil</title>
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

	if request.form("tipo")="" then session("sel48")="0" else session("sel48")=request.form("tipo")
	if request.form("competencia")="" then session("emp48")="Todas" else session("emp48")=request.form("competencia")
	if request.form("valor1")="" then session("va48")="0" else session("va48")=request.form("valor1")
	if request.form("valor2")="" then session("vb48")="0" else session("vb48")=request.form("valor2")

	if request.form("localizar")="" then session("loc48")="" else session("loc48")=request.form("localizar")
		
	if isnumeric(session("loc48"))=true then session("loc48")=session("loc48")

	if session("sel48")<>"0" then
		session("sql48b")="AND (r.id_tipo=" & session("sel48") & ") "
	else
		session("sql48b")=""
	end if

	if session("emp48")<>"Todas" then
		ano=left(session("emp48"),4)
		mes=right(session("emp48"),2)
		session("sql48c")="AND (r.anocomp=" & ano & " and r.mescomp=" & mes & ") "
	else
		session("sql48c")=""
	end if

	if session("loc48")<>"" then
		if isnumeric(session("loc48")) then
			session("sql48d")="AND (r.obs like '%" & session("loc48") & "%') "
		else
			session("sql48d")="AND (r.obs like '%" & session("loc48") & "%') "
		end if
	else
		session("sql48d")=""
	end if
	
	if session("va48")<>"" or session("vb48")<>"" then
		if (session("va48")="" or isnull(session("va48")) or session("va48")="0") and session("vb48")<>"" then session("va48")="0"
		if (session("vb48")="" or isnull(session("vb48")) or session("vb48")="0") and session("va48")<>"" then session("vb48")="999999999"
		session("sql48v")="AND (r.valor between " & nraccess(session("va48")) & " and " & nraccess(session("vb48")) & ") "
	else
		session("sql48v")=""	
	end if

	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")

sqla="SELECT c.*, f.faculdade " & _
"FROM rhconveniadosfac c, rhconveniobe f " & _
"WHERE c.id>0 and f.id_faculdade=c.id_faculdade "
sqla="select r.*, t.tipo from reconciliacao r, reconciliacao_eventos t where t.id_tipo=r.id_tipo "
sqlb=""
sqlc="ORDER BY r.anocomp, r.mescomp, r.data "

sql1=sqla & sqlb & session("sql48b") & session("sql48d") & session("sql48c") & session("sql48v") & sqlc

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
<form method="POST" name="form" action="reconciliacao.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Reconciliação de Pagamentos</p>
<table border="0" width="690" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="55%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
	response.write "<img src='../images/setafirst0.gif' border='0'>"
	response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
	response.write "<a href=""reconciliacao.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
	response.write "<a href=""reconciliacao.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
	response.write "<a href=""reconciliacao.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
	response.write "<a href=""reconciliacao.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
	<td class=campo width="30%" valign="top" align="center">
<% if session("a48")="T" then %>
<a href="reconciliacao_nova.asp" onclick="NewWindow(this.href,'reconciliacao_nova','440','200','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
<font size="1">inserir novo lançamento</font></a><br>
<% end if %>

	</td>
	<td class=campo width="15%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="690" cellpadding="2" style="border-collapse: collapse">
<tr>
	<td class=titulor align="center">Tipo de Pagamento</td>
	<td class=titulor align="center">Data             </td>
	<td class=titulor align="center">Valor            </td>
	<td class=titulor align="center">Competência      </td>
	<td class=titulor align="center">Observação       </td>
	<td class=titulor align="center">A           </td>
</tr>
<%
linha=1
'rs.movefirst:do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
%>
<tr>
	<td class="campor"><%=rs("tipo")%></td>
	<td class="campor"><%=rs("data")%></td>
	<td class="campor"><%=formatnumber(rs("valor"),2) %></td>
	<td class="campor"><%=rs("mescomp") & " / " & rs("anocomp") %></td>
	<td class="campor"><%=rs("obs") %></td>
	<td class="campor">
<% if session("a48")="T" then %>
	<a href="reconciliacao_alteracao.asp?codigo=<%=rs("id_rec")%>" onclick="NewWindow(this.href,'Reconciliacao_alterar','440','200','yes','center');return false" onfocus="this.blur()">
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
<br>
<font size="1">
<%
'sql2="SELECT * from ifip_wstatus"
'rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Tipo: <select size="1" name="tipo">
	<option value="0" <%if cint(session("sel48"))=0 then response.write "selected"%>>Todos Tipos</option>
<%
sql2="select id_tipo, tipo from reconciliacao_eventos order by tipo"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("id_tipo")%>" <%if cint(session("sel48"))=rs2("id_tipo") then response.write "selected"%>><%=rs2("tipo")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>
&nbsp;&nbsp;&nbsp;
<%
sql2="select anocomp, mescomp from reconciliacao group by anocomp, mescomp order by anocomp desc, mescomp desc"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Competência: <select size="1" name="competencia">
	<option value="Todas" <%if session("emp48")="Todas" then response.write "selected"%>>Todas competências</option>
<%
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
valor=numzero(rs2("anocomp"),4) & numzero(rs2("mescomp"),2)
%>
	<option value="<%=valor%>" <%if session("emp48")=valor then response.write "selected"%>><%=rs2("mescomp")&"/"&rs2("anocomp")%></option>
<%
rs2.movenext:loop
end if
rs2.close
%>
</select>
&nbsp;&nbsp;&nbsp;
Valor entre <input type="text" name="valor1" size=6 value="<%=session("va48")%>"> e <input type="text" name="valor2" size=6 value="<%=session("vb48")%>">
<br>
Localizar por observação: <input type="text" name="localizar" size=35 value="<%=session("loc48")%>">
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