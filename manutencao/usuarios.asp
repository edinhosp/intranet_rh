<!-- #config timefmt="%m/%d/%y" -->
<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
'	Response.buffer=true
'	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a99")="N" or session("a99")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cadastro de Usuários</title>
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
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=yes';
win=window.open(mypage,myname,settings);}
// -->
</script>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
sql1="select u.usuario, f.nome, 'apelido'=u.nome, u.grupo, f.codsituacao, u.ativo, u.master " & _
"from usuarios u left join corporerm.dbo.pfunc f on f.chapa=u.usuario collate database_default " & _
"where ativo in (0,1) " & _
"order by ativo desc, u.nome "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<form method="POST" name="form" action="usuarios.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Cadastro de Usuários</p>

<table border="0" cellspacing="1" cellpadding="1" style="border-collapse: collapse" width="690px">
<tr>
	<td class="campo" align="left">
		<input type="radio" name="status" value="A">Ativos  <input type="radio" name="status" value="D">Desativados
	</td>
	<td class="campo" align="right">
		<a href="usuarios_nova.asp?codigo=" onclick="NewWindow(this.href,'Usuario_Nova','550','300','yes','center');return false" onfocus="this.blur()">
		<img src="../imagesr/page_new.gif" border="0" width="10" alt="Inclusão de Usuário"></a>
	</td>
</tr>
</table>

<table border="1" cellspacing="1" cellpadding="1" style="border-collapse: collapse" width="690px">
<tr>
	<td class="titulo">Chapa</td>
	<td class="titulo">Nome</td>
	<td class="titulo">Nome Usuário</td>
	<td class="titulo">Grupo</td>
	<td class="titulo">Ativo</td>
	<td class="titulo">Mst</td>
	<td class="titulo">Validação</td>
	<td class="titulo"></td>
</tr>
<%
do while not rs.eof
valida=""
if rs("ativo")=true then ativo="status_ok.png" else ativo="status_nulo.png"
if rs("master")=true then master="status_ok.png" else master="status_nulo.png"
if rs("ativo")=true and rs("codsituacao")="D" then valida=valida & "Desativar situação D | "
if rs("grupo")="COORD.CURSO" and rs("ativo")=true then
	sql2="select top 1 chapa from n_indicacoes where chapa='" & rs("usuario") & "' and id_nomeacao=12 and GETDATE() between MAND_INI and MAND_FIM"
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then checkcoord=1 else checkcoord=0
	rs2.close
	if checkcoord=0 then valida=valida & "Desativar Coord.Curso | "
end if 
%>
<tr>
	<td class="campo">
		<a class=r href="usuarios_alteracao.asp?codigo=<%=rs("usuario")%>" onclick="NewWindow(this.href,'Usuario_Alterar','550','300','yes','center');return false" onfocus="this.blur()">
		<%=rs("usuario")%></a>
	</td>
	<td class="campo"><%=rs("nome")%></td>
	<td class="campo"><%=rs("apelido")%></td>
	<td class="campo"><%=rs("grupo")%></td>
	<td class="campo" align="center"><img src="../imagesr/<%=ativo%>"></td>
	<td class="campo" align="center"><img src="../imagesr/<%=master%>"></td>
	<td class="campo"><font color="red"><%=valida%></font></td>
	<td class="campo"><img src="../imagesr/tables.gif">
	</td>
</tr>
<%
rs.movenext
loop
%>
</table>


<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>

<!-- -->
<!-- -->
</form>
</body>
</html>