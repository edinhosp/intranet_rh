<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a47")="N" or session("a47")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Lançamentos em Folha - <%=request("nome")%></title>
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
"from corporerm.dbo.pfunc f, corporerm.dbo.psecao s where f.codsecao=s.codigo "
sqlb="and f.chapa='" & chapa & "' "

sql1=sqla & sqlb & sqlc
rs.Open sql1, ,adOpenStatic, adLockReadOnly
	
sql2="select a.id_adm, a.ano, a.mes, a.chapa, a.codevento, a.valor, a.nrovezes, b.descricao, b.valhordiaref " & _
"from apont_adm a, apont_adm_eventos b where a.chapa='" & request("chapa") & "' and b.codigo=a.codevento " & _
"order by a.ano, a.mes, a.codevento "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
'id_autonomo=request("Codigo")	
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>
<% if session("a47")="T" then %>
<a href="apontadm_ver.asp?chapa=<%=chapa%>" onMouseOver="window.status='Clique aqui para atualizar após as alterações'; return true" onMouseOut="window.status=''; return true" onmouseover >
<img border="0" src="../images/write.gif" alt="Clique para atualizar">
<font size="1">!</font>
</a>
<% end if %>
Lançamentos em Folha</p>

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
<tr><th class=titulo colspan=12>Lançamentos para Códigos Fixos</th></tr>
<tr>
	<td class=titulo align="center">Ano</td>
	<td class=titulo align="center">Mês</td>
	<td class=titulo align="center">Cod.</td>
	<td class=titulo align="center">Evento</td>
	<td class=titulo align="center">Tp</td>
	<td class=titulo align="center">Valor</td>
	<td class=titulo align="center">Vezes</td>
	<td class=titulo align="center">&nbsp;</td>
</tr>
<%
rs.close
if rs2.recordcount>0 then
linhas=rs2.recordcount
rs2.movefirst
inicio=1
do while not rs2.eof
if rs2("valhordiaref")="D" then
	valor=rs2("valor")
elseif rs2("valhordiaref")="H" then
	valor1=cdbl(rs2("valor"))
	hora=int(valor1/60)
	minuto=valor1-(hora*60)
	valor=hora & ":" & numzero(minuto,2)
	'valor=formatdatetime((valor1/60)/24,4)
elseif rs2("valhordiaref")="V" then
	valor=formatnumber(rs2("valor"),2)
else
	valor=rs2("valor")
end if
%>
<tr>
	<td class=campo align="center"><%=rs2("ano")%></td>
	<td class=campo align="center"><%=rs2("mes")%></td>
	<td class=campo align="center"><%=rs2("codevento")%></td>
	<td class=campo><%=rs2("descricao")%></td>
	<td class=campo align="center"><%=rs2("valhordiaref")%></td>
	<td class=campo align="right"><%=valor%>&nbsp;</td>
	<td class=campo align="center"><%=rs2("nrovezes")%></td>
	<td class="campor">&nbsp;
	<% if session("a47")="T" then %>
	<a href="apont_alteracao.asp?codigo=<%=rs2("id_adm")%>" onclick="NewWindow(this.href,'AlteracaoFolha','420','200','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/folder95o.gif" width=14 alt="Alterar este lançamento"></a>
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
<tr><td class="campor" colspan=12>&nbsp;</td></tr>
<%
end if
%>
</table>
<!-- quadro fim mudanca -->
<table><tr>
<td valign="top">
<% if session("a47")="T" then %>
<a href="apont_nova.asp?codigo=<%=chapa%>" onclick="NewWindow(this.href,'InclusaoFolha','420','200','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo teto">
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