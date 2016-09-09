<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")="N" or session("a72")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cadastro de Horários - Estagiários</title>
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

id_horario=request("Codigo")	

sqla="SELECT codigo, descricao, datacriacao, jsem, jmes, ativo " & _
"FROM est_cadhorario " & _
"WHERE codigo='" & request("codigo") & "' "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>
CADASTRO DE HORÁRIO</p>
<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr><td class=grupo>Dados </td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;Código</td>
	<td class=titulor>&nbsp;Descrição</td>
</tr>
<tr>
	<td class=campo><b>&nbsp;<%=rs("codigo")%></b></td>
	<td class=campo>&nbsp;<%=rs("descricao")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;Data Criação</td>
	<td class=titulor>&nbsp;Jorn. Semanal</td>
	<td class=titulor>&nbsp;Jorn. Mensal</td>
	<td class=titulor>&nbsp;Ativo</td>
</tr>
<tr>
	<td class=campo>&nbsp;<%=rs("datacriacao")%></td>
	<td class=campo>&nbsp;<%=horaload(rs("jsem"),2)%></td>
	<td class=campo>&nbsp;<%=horaload(rs("jmes"),2)%></td>
	<td class=campo>&nbsp;
	<%if rs("ativo")=0 then response.write "<img src='../images/bullet.gif'>" else response.write "<img src='../images/bullet_hl.gif'>"%>
	</td>
</tr>
</table>

<!-- quadro inicio mudanca-->
<table border="1" bordercolor="#CCCCCC" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><th class=titulo colspan=13>Batidas</th></tr>
<tr>
	<td class=titulor align="center" rowspan=2>Dia</td>
	<td class=titulor align="center" rowspan=2>Entr.</td>
	<td class=titulor align="center" rowspan=2>Saida</td>
	<td class=titulor align="center" rowspan=2>Entr.</td>
	<td class=titulor align="center" rowspan=2>Saida</td>
	<td class=titulor align="center" rowspan=2>Jornada</td>
	<td class=titulor align="center" rowspan=2>Comp.</td>
	<td class=titulor align="center" rowspan=2>Desc.</td>
	<td class=titulor align="center" colspan=2>Intervalo</td>
	<td class=titulor align="center" rowspan=2>&nbsp;</td>
</tr>
<tr>
	<td class=titulor align="center">Flex.?</td>
	<td class=titulor align="center">Embutido</td>
</tr>
<%

sqlm="select * from est_cadhorario_marcacoes where codigo='" & rs("codigo") & "' order by dia"
rs2.Open sqlm, ,adOpenStatic, adLockReadOnly

if rs2.recordcount>0 then
linhas=rs2.recordcount
rs2.movefirst
do while not rs2.eof
%>
<tr>
	<td class=campo align="center"><%=rs2("dia")%></td>
	<td class=campo align="center"><%=horaload(rs2("ent1"),2)%></td>
	<td class=campo align="center"><%=horaload(rs2("sai1"),2)%></td>
	<td class=campo align="center"><%=horaload(rs2("ent2"),2)%></td>
	<td class=campo align="center"><%=horaload(rs2("sai2"),2)%></td>
	<td class=campo align="center"><%=horaload(rs2("jorn"),2)%></td>
	<td class=campo align="center"><%if rs2("comp")=0 then response.write "<img src='../images/bullet.gif'>" else response.write "<img src='../images/bullet_hl.gif'>"%></td>
	<td class=campo align="center"><%if rs2("desc")=0 then response.write "<img src='../images/bullet.gif'>" else response.write "<img src='../images/bullet_hl.gif'>"%></td>
	<td class=campo align="center"><%if rs2("intflex")=0 then response.write "<img src='../images/bullet.gif'>" else response.write "<img src='../images/bullet_hl.gif'>"%></td>
	<td class=campo align="center"><%=horaload(rs2("intdentro"),2)%></td>

	<td class=campo valign=top align="center">
	<% if session("a72")="T" then %>
		<a href="hor_marc_alteracao.asp?codigo=<%=rs("codigo")%>&dia=<%=rs2("dia")%>" onclick="NewWindow(this.href,'AlteracaoMarcação','510','250','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/folder95o.gif" alt="Alterar este Dia"></a>
	<% end if %>
	</td>
</tr>
<%
rs2.movenext
loop
	sqlt="select sum(jorn) as tjorn from est_cadhorario_marcacoes where codigo='" & id_horario & "' "
	rs3.Open sqlt, ,adOpenStatic, adLockReadOnly
	tjorn=rs3("tjorn")
	rs3.close
%>
<tr>
	<td class=fundo colspan=5></td>
	<td class=campo align="center"><%=horaload(tjorn,2)%></td>
	<td class=fundo colspan=5></td>
</tr>
<%
else ' sem registros/planos
%>
<tr><td class="campor" colspan=9>&nbsp;</td></tr>
<%
end if
%>

</table>
<!-- quadro fim mudanca -->
<table><tr>
<td valign="top">
<% if session("a72")="T" then %>
<a href="hor_marc_nova.asp?codigo=<%=id_horario%>" onclick="NewWindow(this.href,'InclusaoMarcacao','510','200','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir nova marcação">
<font size="1">inserir nova marcação</font></a>
<% end if %>
</td>
</tr></table>

</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
set rs3=nothing
conexao.close
set conexao=nothing
%>