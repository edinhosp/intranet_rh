<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a88")="N" or session("a88")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Empréstimos Consignados - <%=request("nome")%></title>
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
chapa=request("chapa")

sqla="select f.chapa, f.nome, f.codsecao, f.codsituacao, s.descricao " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.psecao s where s.codigo=f.codsecao "
sqlb="and f.chapa='" & chapa & "' "

sql1=sqla & sqlb & sqlc
rs.Open sql1, ,adOpenStatic, adLockReadOnly
	
sql2="select idemp, chapa, data, valor, nprestacoes, vprestacao, venc1, vencu, obs, contrato, dt_conv, dt_assfieo, dt_banco " & _
",Status=case when vencu<getdate() then 'Quitado' else case when obs like '%quitado%' then 'Quitado' else 'Em aberto' end end " & _
"from emprestimos where chapa='" & request("chapa") & "' order by data "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly

%>

<p style="margin-top: 0; margin-bottom: 0" class="titulo">
<% if session("a88")="T" then %>
<a href="emprestimo_ver.asp?chapa=<%=chapa%>" onMouseOver="window.status='Clique aqui para atualizar após as alterações'; return true" onMouseOut="window.status=''; return true" onmouseover>
<img border="0" src="../images/write.gif" alt="Clique para atualizar" WIDTH="16" HEIGHT="16">
<font size="1">!</font>
</a>
<% end if %>
Controle de Emprestimos Consignados</p>

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
	<td class="campor">&nbsp;<%=rs("chapa")%></td>
	<td class="campor"><b>&nbsp;<%=rs("nome")%></b></td>
	<td class="campor">&nbsp;<%=rs("codsituacao")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class="titulor">&nbsp;Código</td>
	<td class="titulor">&nbsp;Seção</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rs("codsecao")%></td>
	<td class="campor">&nbsp;<%=rs("descricao")%></td>
</tr>
</table>

<!-- quadro inicio mudanca-->
<table border="1" bordercolor="#CCCCCC" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><th class="titulo" colspan="11">Lançamentos dos Emprestimos Consignados</th></tr>
<tr>
	<td class="titulo" align="center" rowspan=2>Data</td>
	<td class="titulo" align="center" rowspan=2 style="border-right:2px solid #000000">Valor</td>
	<td class="titulo" align="center" rowspan=2># Parc.</td>
	<td class="titulo" align="center" rowspan=2>$ Parc.</td>
	<td class="titulo" align="center" rowspan=2 style="border-left:2px solid #000000">1º Venc.</td>
	<td class="titulo" align="center" rowspan=2 style="border-right:2px solid #000000">Ult.Venc.</td>
	<td class="titulo" align="center" colspan=2 style="border-right:2px solid #000000">Datas Envio</td>
	<td class="titulo" align="center" rowspan=2>Contrato</td>
	<td class="titulo" align="center" rowspan=2 style="border-right:2px solid #000000">Status</td>
	<td class="titulo" align="center" rowspan=2>&nbsp;</td>
</tr>
<tr>
	<td class="titulo" align="center">P/Pro-Reit.</td>
	<td class="titulo" align="center" style="border-right:2px solid #000000">P/Banco</td>
</tr>
<%
rs.close
saldo=0
if rs2.recordcount>0 then
linhas=rs2.recordcount
rs2.movefirst
inicio=1
do while not rs2.eof
if rs2("contrato")="" or isnull(rs2("contrato")) then contrato="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" else contrato=rs2("contrato")
if rs2("status")="Em aberto" then corletra="Red" else corletra="Green"
%>
<tr>
	<td class="campo" align="center"><%=rs2("data")%></td>
	<td class="campo" align="right" style="border-right:2px solid #000000"><%=formatnumber(rs2("valor"),2)%></td>
	<td class="campo" align="center"><%=rs2("nprestacoes")%></td>
	<td class="campo" align="right"><%=formatnumber(rs2("vprestacao"),2)%></td>
	<td class="campo" align="center" style="border-left:2px solid #000000"><%=rs2("venc1")%></td>
	<td class="campo" align="center" style="border-right:2px solid #000000"><%=rs2("vencu")%></td>
	<td class="campo" align="center" style="border-left:2px solid #000000"><%=rs2("dt_assfieo")%></td>
	<td class="campo" align="center" style="border-right:2px solid #000000"><%=rs2("dt_banco")%></td>
	<td class="campor" align="left"><%=contrato%></td>
	<td class="campor" align="left" style="border-right:2px solid #000000"><b><font color=<%=corletra%>><%=rs2("status")%></font></td>
	<td class="campor">&nbsp;
	<% if session("a88")="T" then %>
	<a href="emprestimo_alteracao.asp?codigo=<%=rs2("idemp")%>" onclick="NewWindow(this.href,'AlteracaoEmprestimo','420','250','no','center');return false" onfocus="this.blur()">
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
  <tr><td class="campor" colspan="9">&nbsp;</td></tr>
<%
end if
%>
</table>
<!-- quadro fim mudanca -->
<table><tr>
<td valign="top">
<% if session("a88")="T" then %>
<a href="emprestimo_nova.asp?codigo=<%=chapa%>" onclick="NewWindow(this.href,'InclusaoEmprestimo','420','250','no','center');return false" onfocus="this.blur()">
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