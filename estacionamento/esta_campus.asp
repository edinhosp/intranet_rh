<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a87")="N" or session("a87")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cadastro de Estacionamento</title>
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
dim conexao, conexao2
dim rs, rs2
set conexao=server.createobject ("ADODB.Connection")
'conexao.Open Application("consql")

if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"

	if request.form("secao")="" then
		session("sel87")="Todas"
	else
		session("sel87")=request.form("secao")
	end if

	if request.form("localizar")="" then
		session("loc87")=""
	else
		session("loc87")=request.form("localizar")
	end if
		
	'if isnumeric(session("loc87"))=true then session("loc87")=numzero(session("loc87"),5)

	if session("sel87")<>"Todas" then
		session("sql87b")="AND (a.tipo_prestacao='" & session("sel87") & "') "
	else
		session("sql87b")=""
	end if

	if session("loc87")<>"" then
		session("sql87d")="AND (f.nome like '%" & session("loc87") & "%') "
	else
		session("sql87d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if

registros=Session("RegistrosPorPagina")

sqla="SELECT v.*, f.nome " & _
"FROM veiculos_alunosfunc v, corporerm.dbo.pfunc f " & _
"WHERE v.chapa=f.chapa collate database_default and v.id_esta>0 "
sqlb=""
sqlc="ORDER BY validade DESC, campus_estudo, f.nome "

sql1=sqla & sqlb & session("sql87b") & session("sql87d") & sqlc

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
	if rs.recordcount>0 then 	rs.AbsolutePage=Session("Pagina") 'vai para o n�mero da pagina armazenado na sess�o
End Sub
%>
<form method="POST" name="form" action="esta_campus.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Controle de Estacionamento para funcion�rios estudando</p>
<table border="0" width="650" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="70%" valign="top" align="left">P�gina: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
	response.write "<img src='../images/setafirst0.gif' border='0'>"
	response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
	response.write "<a href=""esta_campus.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
	response.write "<a href=""esta_campus.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
	response.write "<a href=""esta_campus.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
	response.write "<a href=""esta_campus.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
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

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="690" cellpadding="1" style="border-collapse: collapse">
<tr>
	<td class=titulor align="center" rowspan=2>Chapa </td>
	<td class=titulor align="center" rowspan=2>Funcion�rio</td>
	<td class=titulor align="center" rowspan=2>Matricula</td>
	<td class=titulor align="center" colspan=2>Campus</td>
	<td class=titulor align="center" rowspan=2>Per�odo</td>
	<td class=titulor align="center" colspan=3>Ve�culo</td>
	<td class=titulor align="center" rowspan=2>Dt.Autor.</td>
	<td class=titulor align="center" rowspan=2>Validade</td>
	<td class=titulor align="center" rowspan=2>&nbsp;        </td>
	<td class=fundor align="center" rowspan=2>&nbsp;         </td>
	<td class=titulor align="center" rowspan=2>#</td>
</tr>
<tr>
	<td class=titulor align="center">trabalho</td>
	<td class=titulor align="center">estudo</td>
	<td class=titulor align="center">modelo</td>
	<td class=titulor align="center">cor</td>
	<td class=titulor align="center">placa</td>
</tr>
<%
linha=1
'rs.movefirst:do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
%>
<tr>
	<td class="campor"><%=rs("chapa") %></td>
	<td class="campor"><%=rs("nome")%></td>
	<td class="campor"><a class=r href="..\ouvidoria\alunos_ver.asp?matricula=<%=rs("matricula")%>" onclick="NewWindow(this.href,'Ouvidoria','690','500','yes','center');return false" onfocus="this.blur()">
	<%=rs("matricula")%></a></td>
	<td class="campor" align="center"><%=rs("campus_trabalho")%></td>
	<td class="campor" align="center"><%=rs("campus_estudo")%></td>
	<td class="campor" align="center"><%=rs("periodo")%></td>
	<td class="campor"><%=rs("modelo")%></td>
	<td class="campor"><%=rs("cor")%></td>
	<td class="campor"><%=rs("placa")%></td>
	<td class="campor" align="center"><%=rs("dtauto")%></td>
	<td class="campor" align="center"><%=rs("validade")%></td>

	<td class="campor" align="center">
	<% if session("a87")="T" or session("a87")="C" then %>
		<a href="esta_campus_alteracao.asp?codigo=<%=rs("id_esta")%>" onclick="NewWindow(this.href,'AlteracaoEstacionamento','420','260','no','center');return false" onfocus="this.blur()">
		<img src="../images/Folder95O.gif" border="0" width=13 alt="Alterar os dados cadastrais"></a>
	<% end if %>
	</td>
	<td class="campor" align="center">
	<a class=r href="esta_cracha.asp?chapa=<%=rs("chapa")%>&ano=<%=rs("validade")%>" onclick="NewWindow(this.href,'ImpressaoCracha','565','500','yes','center');return false" onfocus="this.blur()">
	<img src='../images/truck.gif' border='0'></a>
	</td>
	<td class="campor" align="center"><%=rs("sequencia")%></td>
</tr>
<%
if linha=1 then linha=0 else linha=1
rs.movenext
if rs.eof then exit for
'loop
Next

else 'sem registros
%>
<td class=grupo colspan=12>Esta sele��o n�o mostra nenhum registro.</td>
<%
end if
%>
</table>
<%
if session("a87")="T" then
%>
<a href="esta_campus_nova.asp" onclick="NewWindow(this.href,'InclusaoEstacionamento','420','260','no','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif" alt="Cadastrar novo controle">
<font size="1">inserir novo controle</font></a><br>
<%
end if
%>
<font size="1">
<!--
<%
sql2="select tipo_prestacao as servico from autonomo group by tipo_prestacao"
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Campus: <select size="1" name="secao">
<option value="Todas" <%if session("sel87")="Todas" then response.write "selected"%>>Todos servi�os</option>
<%
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
%>
    <option value="<%=rs2("servico")%>" <%if session("sel87")=rs2("servico") then response.write "selected"%>><%=rs2("servico")%></option>
<%
rs2.movenext
loop
end if 'rs2.recordcount
rs2.close
%>
</select>
<br>
-->
Localizar por nome: <input type="text" name="localizar" size=35 value="<%=session("loc87")%>">
Registros/P�gina: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
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