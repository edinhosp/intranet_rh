<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a24")="N" or session("a24")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Nomeados por Portaria</title>
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
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
t1="UPDATE n_indicacoes SET n_indicacoes.temp = Right([portaria],2);"
conexao.execute t1
portaria=""
	
if request("codigo")<>"" or request.form<>"" then
	if request.form="" then idnomeacao=request("codigo")
	if request("codigo")="" then idnomeacao=request.form("D1")
	sqlc="SELECT Status=case when mand_fim<getdate() then 'Vencidas' else 'Ativas' end "
	sqlc="SELECT " & _
	"i.id_nomeacao, n.nomeacao, i.PORTARIA, i.id_indicado, i.CHAPA, i.NOME, i.complemento, " & _
	"i.CARGO, i.codeve, i.MAND_INI, i.MAND_FIM, i.alunos, i.CH, i.OBS, i.CONTRATO, i.temp " & _
	"FROM n_indicacoes as i, n_nomeacoes as n"
	sqld=" where (i.id_nomeacao=n.id_nomeacao) "
	if idnomeacao="todas" then 
		sqld=sqld & " "
		portaria="Ativas/Vencidas"
	elseif idnomeacao="ativas" then
		sqld=sqld & " and (mand_fim>getdate() or mand_fim is null ) "
		portaria="Ativas"
	elseif idnomeacao="semvenc" then
		sqld=sqld & " and ( mand_fim is null ) "
		portaria="Sem vencimento"
	else
		sqld=sqld & " and i.portaria='" & idnomeacao & "' "
	end if
	sqle=" order by n.nomeacao, i.nome, i.portaria "
	sqlb=sqlc & sqld & sqle
	'response.write sqlb
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
	if portaria="" then portaria=rs("portaria")
	temp=0
else
	temp=1
end if
'	session("nomeacao_chapa")=""
'	session("nomeacao_id")=""
'	session("nomeacao_descr")=""

if temp=1 then
	sqla="SELECT PORTARIA FROM n_indicacoes GROUP BY temp, PORTARIA order by temp, portaria "
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<p class=titulo>Nomeações <%=portaria %>
<form method="POST" action="nomeados_portaria.asp" name="form">
	<p><select size="1" name="D1" style="font-size: 8 pt">
	<option value="todas">Todas portarias</option>
	<option value="ativas">Portarias Ativas</option>
	<option value="semvenc">Portarias sem vencimento</option>
<%
rs.movefirst:do while not rs.eof
%>
	<option value="<%=rs("portaria")%>"><%=rs("portaria")%></option>
<%
rs.movenext:loop
rs.close
%>
	</select>
	<br>
	<input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>
<p style="margin-top: 0; margin-bottom: 0"><font color="#FF0000">Configure a página do seu navegador (Internet
Explorer, Netscape, Mozilla, etc) no sentido PAISAGEM.</font></p>
<%
else ' temp=0
%>
<p class=titulo>Nomeações <%=portaria %>
<table border="1" cellpadding="0" cellspacing="1" style="border-collapse: collapse" width="1000">
<tr>
	<td class=titulor align="center">Nomeação</td>
<!--	<td class=titulor align="center">Chapa</td> -->
	<td class=titulor align="center">Nome/Docente</td>
	<td class=titulor align="center">Cargo</td>
	<td class=titulor align="center">Portaria</td>
	<td class=titulor align="center">Inicio em</td>
	<td class=titulor align="center">Término</td>
	<td class=titulor align="center">C.H.</td>
<!--	<td class=titulor align="center">Compl.</td>
	<td class=titulor align="center">Evento</td> -->
</tr>
<%
linhas=2
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 

if linhas=52 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<p style='margin-top:0; margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<p class=titulo>Nomeações " & portaria & ""
	linhas=1
	response.write "<table border='1' cellpadding='0' cellspacing='1' style='border-collapse: collapse' width='1000'>"
	response.write "<tr>"
	response.write "<td class=titulor align='center'>Nomeação</td>"
'	response.write "<td class=titulor align='center'>Chapa</td>"
	response.write "<td class=titulor align='center'>Nome/Docente</td>"
	response.write "<td class=titulor align='center'>Cargo</td>"
	response.write "<td class=titulor align='center'>Portaria</td>"
	response.write "<td class=titulor align='center'>Inicio em</td>"
	response.write "<td class=titulor align='center'>Término</td>"
	response.write "<td class=titulor align='center'>C.H.</td>"
'	response.write "<td class=titulor align='center'>Compl.</td>"
'	response.write "<td class=titulor align='center'>Evento</td>"
	response.write "</tr>"
	linhas=linhas+1
end if

%>
<tr>
	<td class="campor" height="10" nowrap ><%=rs("nomeacao") %></td>
<!--	<td class="campor" height="10" nowrap ><%=rs("Chapa") %></td> -->
	<td class="campor" height="10" nowrap ><%=left(rs("nome"),40) %> (<%=rs("chapa")%>)</td>
	<td class="campor" height="10" nowrap ><%=rs("cargo") %></td>
	<td class="campor" height="10" nowrap ><%=rs("portaria") %></td>
	<td class="campor" height="10" nowrap align="center"><%=rs("mand_ini") %></td>
	<td class="campor" height="10" nowrap align="center"><%=rs("mand_fim") %></td>
	<td class="campor" height="10" nowrap align="center"><%=rs("ch") %></td>
<!--	<td class="campor" height="10" nowrap align="center"><%=rs("complemento")%></td>
	<td class="campor" height="10" nowrap align="center"><%=rs("codeve")%></td> -->
<%
rs.movenext
linhas=linhas+1
loop
total=rs.recordcount
rs.close
else 'recordcount=0
	response.write "<tr><td class=""campor"" colspan='9'>"
	response.write "Não indicações para esta categoria de nomeação"
	response.write "</td></tr>"
end if 'recordcount=0
%>
</table>
<%
	pagina=pagina+1
	response.write "<p style='margin-top:0; margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"

end if 'temp=0

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>