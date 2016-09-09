<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a36")="N" or session("a36")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Saldo de Dias de Férias - <%=request("nome")%></title>
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
conexao.Open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
set rsq=server.createobject ("ADODB.Recordset")
Set rsq.ActiveConnection = conexao
chapa=request("chapa")

sqla="select f.chapa, f.nome, f.codsecao, f.codsituacao, s.descricao " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.psecao s where s.codigo=f.codsecao "
sqlb="and f.chapa='" & chapa & "' "

sql1=sqla & sqlb & sqlc
rsq.Open sql1, ,adOpenStatic, adLockReadOnly
	
sql2="select c.id_fer, c.chapa, c.dtfimper, c.dtinigozo, c.dtfimgozo, c.dias, c.tipo, c.obs " & _
"from ferias c where c.chapa='" & request("chapa") & "' " & _
"order by c.dtinigozo, c.dtfimper, c.tipo  "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
'id_autonomo=request("Codigo")	
%>

<p style="margin-top: 0; margin-bottom: 0" class="titulo">
<% if session("a36")="T" then %>
<a href="ferias_ver.asp?chapa=<%=chapa%>" onMouseOver="window.status='Clique aqui para atualizar após as alterações'; return true" onMouseOut="window.status=''; return true" onmouseover>
<img border="0" src="../images/write.gif" alt="Clique para atualizar" WIDTH="16" HEIGHT="16">
<font size="1">!</font>
</a>
<% end if %>
Controle de Saldo de Férias</p>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr><td class="grupo">Dados Pessoais</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class="titulor">&nbsp;Chapa</td>
	<td class="titulor">&nbsp;Nome do funcionario</td>
	<td class="titulor">&nbsp;Situação</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rsq("chapa")%></td>
	<td class="campor"><b>&nbsp;<%=rsq("nome")%></b></td>
	<td class="campor">&nbsp;<%=rsq("codsituacao")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class="titulor">&nbsp;Código</td>
	<td class="titulor">&nbsp;Seção</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rsq("codsecao")%></td>
	<td class="campor">&nbsp;<%=rsq("descricao")%></td>
</tr>
</table>

<!-- quadro inicio mudanca-->
<table border="1" bordercolor="#CCCCCC" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><th class="titulo" colspan="12">Lançamentos de Férias</th></tr>
<tr>
	<td class="titulo" align="center">Per.Aquis.</td>
	<td class="titulo" align="center">Tipo</td>
	<td class="titulo" align="center">Início</td>
	<td class="titulo" align="center">Término</td>
	<td class="titulo" align="center">Dias</td>
	<td class="titulo" align="center">Obs.</td>
	<td class="titulo" align="center">Saldo</td>
	<td class="titulo" align="center">&nbsp;</td>
</tr>
<%
rsq.close
saldo=0
if rs2.recordcount>0 then
linhas=rs2.recordcount
rs2.movefirst
inicio=1
do while not rs2.eof
if rs2("tipo")="C" then saldo=saldo+cdbl(rs2("dias")) else saldo=saldo-cdbl(rs2("dias"))
if rs2("tipo")="C" then classe="campoar" else classe="camporr"
%>
<tr>
	<td class="campo" align="center"><%=rs2("dtfimper")%></td>
	<td class="campo" align="center"><%=rs2("tipo")%></td>
	<td class="campo" align="center"><%=rs2("dtinigozo")%></td>
	<td class="campo" align="center"><%=rs2("dtfimgozo")%></td>
	<td class="<%=classe%>" align="center"><b><%=rs2("dias")%></td>
	<td class="campor" align="left"><%=rs2("obs")%></td>
	<td class="campo" align="center"><%=saldo%></td>
	<td class="campor">&nbsp;
	<% if session("a36")="T" then %>
	<a href="fer_alteracao.asp?codigo=<%=rs2("id_fer")%>" onclick="NewWindow(this.href,'AlteracaoCSFer','420','200','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/folder95o.gif" width="14" alt="Alterar este lançamento"></a>
	<% end if %>
	</td>
</tr>
<%
if rs2.absoluteposition=rs2.recordcount then
end if
rs2.movenext
inicio=0
loop
else ' sem registros/planos
%>
  <tr><td class="campor" colspan="12">&nbsp;</td></tr>
<%
end if
%>
</table>
<!-- quadro fim mudanca -->
<table><tr>
<td valign="top">
<% if session("a36")="T" then %>
<a href="fer_nova.asp?codigo=<%=chapa%>" onclick="NewWindow(this.href,'InclusaoFer','420','200','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo lançamento" WIDTH="16" HEIGHT="16">
<font size="1">inserir novo lançamento</font></a>
<% end if %>
</td>
</tr></table>

</body>
</html>
<%
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>