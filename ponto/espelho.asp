<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a4")="N" or session("a4")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Controle de devolução de espelho</title>
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
<script language="JavaScript" type="text/javascript"><!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() {	form.chapa.value=form.nome.value; }
function chapa1() {	form.nome.value=form.chapa.value; }
--></script>
</head>
<body>
<%
'if request.form("B3")<>"" then response.redirect "rhform_recibotcar2.asp"
response.write request.form("B3")
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
if request("acao")="excluir" then
	id=Request.QueryString("id")
	sql="delete from espelho where chapa='" & id & "'"
	conexao.execute sql
	manutencao=1
end if

'if request.form("B2")<>"" then
if request.form<>"" then
	iCount=request("Count")
	'for iLoop=0 to iCount
	'	aid=request("id" & iLoop)
	'	achapa=request("chapa" & iLoop)
	'	strSql="Update rh_ticket_recibo Set chapa = '" & achapa & "' Where chapa='" & aid
	'	conexao.execute strSql, , adCmdText
	'next
	if request.form("chapa")<>"" then
		chapa=numzero(request.form("chapa"),5)
		sSql="Insert Into espelho (chapa) "
		sSql=sSql & "Values ('" & chapa & "' "
		sSql=sSql & ")"
		conexao.Execute sSQL, , adCmdText
	end if
	manutencao=1
end if
'if manutencao=1 then response.redirect "rhcursos.asp?codigo=" & session("idcurso")
%>
<p class=titulo>Controle de devolução de espelho de ponto
<form method="POST" name="form" action="espelho.asp">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="center">Nome</td>
	<td class=titulo align="center"><font size="2">&nbsp;</font></td>
</tr>
<tr>
	<td class=campo><input type="text" class="form_box2" name="chapa" size="5" onchange="chapa1()" setfocus></td>
	<td class=campo><select size="1" name="nome" onchange="nome1()">
<%
sqltemp="select chapa, nome from pfunc where chapa<'10000' or chapa>'90000' order by nome"
sqltemp="select a.chapa, f.nome from corporerm.dbo.aparfun a, corporerm.dbo.pfunc f where a.chapa=f.chapa and f.codsituacao<>'D' and f.codsindicato<>'03' order by f.nome"
rs.Open sqltemp, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
%>
<option value="<%=rs("chapa")%>"><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
	</select>
	</td>
	<td class=campo>
	<input type="hidden" name="Count" value="<%=tcount-1%>">
	<input type="submit" value="Salvar" class="button" name="B2">
	</td>
</tr>
<%
sqlc="SELECT z.chapa, nome from espelho z, corporerm.dbo.pfunc f where z.chapa=f.chapa collate database_default order by nome "
rs.Open sqlc, ,adOpenStatic, adLockReadOnly
laststatus=""
if rs.recordcount>0 then
	tcount=0
	rs.movefirst
	do while not rs.eof 
%>
<tr>
	<td class=campo align="center"><%=rs("chapa")%></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo align="center">
		<a href="espelho.asp?acao=excluir&id=<%=rs("chapa")%>">
		<img border="0" src="../images/Trash.gif"></a>
	</td>
</tr>
<%
	rs.movenext
	tcount=tcount+1
	loop
else 'recordcount=0
 	response.write "<tr><td class=grupo colspan='3'>"
	response.write "<p>Não há funcionários selecionados.</td></tr>"
end if 'recordcount=0
rs.close
%>
</table>
</form>
<hr>
<DIV style=""page-break-after:always""></DIV>
<%
sql="select a.chapa, f.nome, f.codsituacao as sit, f.codsecao, s.descricao as setor from corporerm.dbo.aparfun a, corporerm.dbo.pfunc f, corporerm.dbo.psecao s " & _
"where a.chapa=f.chapa and f.codsecao=s.codigo and f.codsituacao<>'D' and f.codsindicato<>'03' " & _
"and a.chapa collate database_default not in (select e.chapa from espelho e) " & _
"order by f.nome"
rs.Open sql, ,adOpenStatic, adLockReadOnly

total=0
rs.movefirst
response.write "<table border='1' cellpadding='1' cellspacing='0' style='border-collapse:collapse' width='690'>"
response.write "<tr><td class=grupo colspan=" & rs.fields.count & ">Funcionários que ainda não devolveram o espelho</td></tr>"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulo>" & ucase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=campo>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
rs.close
response.write "<p>"
%>

</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>