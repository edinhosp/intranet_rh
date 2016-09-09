<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a61")="N" or session("a61")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Teto Máximo - <%=request("nome")%></title>
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
	
sql2="select * from rhcontroleteto where chapa='" & request("chapa") & "' order by ano desc, mes desc"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	'id_autonomo=request("Codigo")	
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>
<% if session("a61")="T" or session("a61")="C" then %>
<a href="controleteto_ver.asp?chapa=<%=chapa%>" onMouseOver="window.status='Clique aqui para atualizar após as alterações'; return true" onMouseOut="window.status=''; return true" onmouseover >
<img border="0" src="../images/write.gif" alt="Clique para atualizar">
<font size="1">!</font>
</a>
<% end if %>
CONTROLE DE TETO MÁXIMO</p>
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
<tr>
<th class=titulo colspan=12>
<a href="teto_nova.asp?codigo=<%=chapa%>" onclick="NewWindow(this.href,'InclusaoTeto','450','250','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo teto"></a>
Controle de Entrega</th></tr>
<tr>
	<td class=titulor align="center">Ano</td>
	<td class=titulor align="center">Mês</td>
	<td class=titulor align="center">Entrega</td>
	<td class=titulor align="center">Empresa</td>
	<td class=titulor align="center">Proporc</td>
	<td class=titulor align="center">Cópia?</td>
	<td class=titulor align="center">&nbsp;</td>
</tr>
<%
rs.close
if rs2.recordcount>0 then
linhas=rs2.recordcount
rs2.movefirst
inicio=1
do while not rs2.eof
%>
<tr>
	<td class="campor" align="center"><%=rs2("ano") %></td>
	<td class="campor" align="center"><%=rs2("mes") %>    </td>
	<td class="campor" align="center"><%=rs2("data") %>   </td>
	<td class="campor"><%=rs2("empresa")%></td>
	<td class="campor" align="right"><%=formatnumber(rs2("proporcional"),2)%></td>
	<td class="campor" align="center">&nbsp;<%if rs2("copia")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor">&nbsp; 
	<% if session("a61")="T"  or session("a61")="C" then %>
	<a href="teto_alteracao.asp?codigo=<%=rs2("id_teto")%>" onclick="NewWindow(this.href,'AlteracaoTeto','450','250','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/folder95o.gif" alt="Alterar este teto"></a>
	<% end if %>
	</td>
</tr>
<%
if rs2.absoluteposition=rs2.recordcount then
	session("tetoultimaempresa")=rs2("empresa")
	session("tetoultimocnpj")=rs2("cnpj")
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
<% if session("a61")="T" or session("a61")="C" then %>
<a href="teto_nova.asp?codigo=<%=chapa%>" onclick="NewWindow(this.href,'InclusaoTeto','450','250','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo teto">
<font size="1">inserir novo teto</font></a>
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