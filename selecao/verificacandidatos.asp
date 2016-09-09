<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a68")="N" or session("a68")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Grade Horária</title>
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
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
%>
<p class=titulo>Verificação de Candidatos já entrevistados&nbsp;<%=titulo %>
<br>
<form method="POST" action="verificacandidatos.asp" name="form">
<p>
<select size="1" name="funcao" class=a  onChange="javascript:submit()">
	<option value="0">Selecione uma função</option>
<%
sql2="select f.codigo, f.nome from corporerm.dbo.pfuncao f, rs_requisicao r where r.funcao=f.codigo collate database_default group by f.codigo, f.nome order by nome "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
if request.form("funcao")=rs("codigo") then tempf="selected" else tempf=""
%>
	<option value="<%=rs("codigo")%>" <%=tempf%>><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
	</select>  
</form>
<!-- impressão do documento -->
<%
if request.form("funcao")<>"" then funcao=request.form("funcao") else funcao="-1"
sql1="SELECT c.id_candidato, c.nome_candidato, c.idade, c.telefone, " & _
"r.id_requisicao, r.funcao, r.descricao, r.dt_abertura, r.dt_encerramento " & _
"FROM rs_candidato c, rs_requisicao r WHERE c.id_requisicao = r.id_requisicao " & _
"AND r.funcao='" & funcao & "' ORDER BY c.nome_candidato "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650>
<tr>
	<td class=grupo>Candidato</td>
	<td class=grupo>Idade</td>
	<td class=grupo>Telefone</td>
	<td class=grupo>Vaga candidatada</td>
	<td class=grupo>Abertura</td>
	<td class=grupo>Encerr.</td>
</tr>
<%
linha=0
rs.movefirst:do while not rs.eof
if linha=0 then classe="campol" else classe="campo"
%>
<tr>
	<td class=<%=classe%>><%=rs("nome_candidato")%></td>
	<td class=<%=classe%>><%=rs("idade")%></td>
	<td class=<%=classe%>><%=rs("telefone")%></td>
	<td class=<%=classe%>>
		<a class=r href="requisicaover.asp?codigo=<%=rs("id_requisicao")%>" onclick="NewWindow(this.href,'Requisicao_ver','595','500','yes','center');return false" onfocus="this.blur()">
		<%if rs("descricao")="" or isnull(rs("descricao")) then response.write "Sem descrição" else response.write rs("descricao")%></a></td>
	</td>
	<td class=<%=classe%>><%=rs("dt_abertura")%></td>
	<td class=<%=classe%>><%=rs("dt_encerramento")%></td>
</tr>
<%
rs.movenext
if linha=0 then linha=1 else linha=0
loop
end if
rs.close

set rs=nothing
conexao.close
set conexao=nothing
%>
</table>
</body>
</html>