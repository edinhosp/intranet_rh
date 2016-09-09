<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a6")="N" or session("a6")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Seleção temporária de funcionários</title>
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
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
--></script>
<%
response.write request.form("B3")
dim conexao, conexao2, chapach
dim rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
'set conexao2=server.createobject ("ADODB.Connection")
'conexao2.Open application("consql")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
if request("acao")="excluir" then
	id=Request.QueryString("id")
	sql="delete from zselecao where id_sel=" & id
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
		sqln="select nome from corporerm.dbo.pfunc where chapa='" & chapa & "' "
		rs2.Open sqln, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then nome=rs2("nome") else nome=""
		rs2.close
		sSql="Insert Into zselecao (chapa, sessao ) "
		sSql=sSql & "Values ('" & chapa & "','" & session("usuariomaster") & "' "
		sSql=sSql & ")"
		conexao.Execute sSQL, , adCmdText
	end if
	manutencao=1
end if
'if manutencao=1 then response.redirect "rhcursos.asp?codigo=" & session("idcurso")
%>
<p class="titulo">Seleção especial de funcionários
<form method="POST" name="form" action="especial.asp">
<table border="1" bordercolor="gray" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="500">
	<tr>
		<td class="titulo" align="center">Chapa</td>
		<td class="titulo" align="center">Nome</td>
		<td class="titulo" align="center"><font size="2">&nbsp;</font></td>
	</tr>

	<tr>
		<td class="campo"><input type="text" class="form_box2" name="chapa" size="5" onchange="chapa1()"></td>
		<td class="campo"><select size="1" name="nome" onchange="nome1()"><option value>Selecione um funcionário...</option>
<%
sqltemp="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' order by nome"
rs2.Open sqltemp, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof 
%>
<option value="<%=rs2("chapa")%>"><%=rs2("nome")%></option>
<%
rs2.movenext:loop
rs2.close
%>
		</select>
		</td>
		<td class="campo">
		<input type="hidden" name="Count" value="<%=tcount-1%>">
		<input type="submit" value="Salvar" class="button" name="B2">
		</td>
	</tr>


<%
sqlc="SELECT id_sel, z.chapa, f.nome, z.sessao from zselecao z, corporerm.dbo.pfunc f where f.chapa=z.chapa collate database_default and z.sessao='" & session("usuariomaster") & "' order by nome "
rs.Open sqlc, ,adOpenStatic, adLockReadOnly
laststatus=""
if rs.recordcount>0 then
	tcount=0
	rs.movefirst
	do while not rs.eof 
%>
	<tr>
		<td class="campo" align="center"><%=rs("chapa")%></td>
		<td class="campo"><%=rs("nome")%></td>
		<td class="campo" align="center">
		<a href="especial.asp?acao=excluir&amp;id=<%=rs("id_sel")%>">
		<img border="0" src="../images/Trash.gif" WIDTH="16" HEIGHT="16"></a>
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
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
'conexao2.close
'set conexao2=nothing
%>