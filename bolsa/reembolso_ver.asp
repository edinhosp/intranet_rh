<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a65")="N" or session("a65")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Reembolsos - <%=request("nome")%></title>
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
chapa=request("chapa")

sqla="select f.chapa, f.nome, f.codsecao, f.codsituacao, s.descricao " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.psecao s where s.codigo=f.codsecao "
sqlb="and f.chapa='" & chapa & "' "
sql1=sqla & sqlb & sqlc
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr><td class=realce>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>
<% if session("a65")="T" then %>
<a href="reembolso_ver.asp?chapa=<%=chapa%>" onMouseOver="window.status='Clique aqui para atualizar após as alterações'; return true" onMouseOut="window.status=''; return true" onmouseover >
<img border="0" src="../images/write.gif" alt="Clique para atualizar">
<font size="1">!</font>
</a>
<% end if %>
Controle de Reembolsos</p>
</td>
<td>
<% if session("a65")="T" then %>
<a href="reembolso_nova.asp?codigo=<%=chapa%>&id=<%=request("id_bolsa")%>&porc=<%=porc%>" onclick="NewWindow(this.href,'InclusaoReembolso','420','200','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo reembolso">
<font size="1">inserir novo reembolso</font></a>
<% end if %>
</td>
</tr></table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr><td class=grupo>Dados Pessoais</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;Chapa</td>
	<td class=titulor>&nbsp;Nome do funcionario</td>
	<td class=titulor>&nbsp;Situação</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rs("chapa")%></td>
	<td class="campor"><b>&nbsp;<%=rs("nome")%></b></td>
	<td class="campor">&nbsp;<%=rs("codsituacao")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;Código</td>
	<td class=titulor>&nbsp;Seção</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rs("codsecao")%></td>
	<td class="campor">&nbsp;<%=rs("descricao")%></td>
</tr>
</table>

<!-- quadro inicio mudanca-->
<table border="1" bordercolor="#CCCCCC" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><th class=titulo colspan=6>Lançamentos de Reembolsos</th></tr>
<tr>
	<td class=titulo align="center">Mês Base</td>
	<td class=titulo align="center">Mensalidade</td>
	<td class=titulo align="center">Perc.%</td>
	<td class=titulo align="center">Reembolso</td>
	<td class=titulo align="center">Dt.Pagamento</td>
	<td class=titulo align="center">&nbsp;</td>
</tr>
<%
rs.close
totalpag=0
sql2="select r.id_reembolso, r.id_bolsa, r.chapa, r.mes_base, r.mensalidade, r.porcentagem, r.reembolso, r.data_pagamento " & _
"from bolsistas_reembolso r where r.chapa='" & request("chapa") & "' and r.id_bolsa=" & request("id_bolsa") & " " & _
"order by r.mes_base "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
'id_autonomo=request("Codigo")	
if rs2.recordcount>0 then
linhas=rs2.recordcount
rs2.movefirst
inicio=1
do while not rs2.eof
mesbase=numzero(month(rs2("mes_base")),2) & "/" & year(rs2("mes_base"))%>
<tr>
	<td class=campo align="center"><%=mesbase%></td>
	<td class=campo align="center"><%=formatnumber(rs2("mensalidade"),2)%></td>
	<td class=campo align="center"><%=rs2("porcentagem")%></td>
	<td class=campo align="center"><%=formatnumber(rs2("reembolso"),2)%></td>
	<td class=campo align="center"><%=rs2("data_pagamento")%></td>
	<td class="campor">&nbsp;
	<% if session("a65")="T" then %>
	<a href="reembolso_alteracao.asp?codigo=<%=rs2("id_reembolso")%>" onclick="NewWindow(this.href,'AlteracaoReembolso','420','200','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/folder95o.gif" width=14 alt="Alterar este lançamento"></a>
	<% end if %>
	</td>
</tr>
<%
if rs2.absoluteposition=rs2.recordcount then
end if
porc=rs2("porcentagem")
totalpag=totalpag+cdbl(rs2("reembolso"))
rs2.movenext
inicio=0
loop
%>
<tr>
	<td class=titulo colspan=3>Total recebido</td>
	<td class=campo align="center"><%=formatnumber(totalpag,2)%></td>
	<td class=titulo colspan=2></td>
</tr>

<%
else ' sem registros/planos
%>
  <tr><td class="campor" colspan=6>&nbsp;</td></tr>
<%
end if
%>
</table>
<!-- quadro fim mudanca -->
<table><tr>
<td valign="top">
<% if session("a65")="T" then %>
<a href="reembolso_nova.asp?codigo=<%=chapa%>&id=<%=request("id_bolsa")%>&porc=<%=porc%>" onclick="NewWindow(this.href,'InclusaoReembolso','420','200','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo reembolso">
<font size="1">inserir novo reembolso</font></a>
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