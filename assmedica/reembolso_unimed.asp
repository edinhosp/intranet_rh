<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a59")="N" or session("a59")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Reembolso Unimed</title>
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
function nome1() { form.chapa.value=form.nome.value;form.submit(); }
function chapa1() { form.nome.value=form.chapa.value;form.submit(); }
--></script>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request.form("B1")="" or request.form("id")="" then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Sele��o de funcion�rio para reembolso
<form method="POST" action="reembolso_unimed.asp" name="form">
<%
sqla="SELECT m.chapa, f.nome from assmed_mudanca m, corporerm.dbo.pfunc f, assmed_planos p " & _
"where m.chapa=f.chapa collate database_default and (m.empresa=p.codigo and m.plano=p.plano) " & _
"and m.empresa='U' and ((getdate() between m.ivigencia and m.fvigencia) or ('20141031' between m.ivigencia and m.fvigencia)) order by f.nome"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=request.form("chapa")%>">
<select name="nome" class=a onchange="nome1()">
	<option value="0">Selecione o funcion�rio</option>
<%
rs.movefirst:do while not rs.eof
if request.form("chapa")=rs("chapa") then temps="selected" else temps=""
%>
	<option value="<%=rs("chapa")%>" <%=temps%> ><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<br><br>
<table border="1" cellpadding="2" cellspacing="2" style="border-collapse: collapse">
<tr>
	<td class=titulo></td>
	<td class=titulo>Tipo</td>
	<td class=titulo>Nome</td>
</tr>
<%
sqlt="SELECT 'Titular' as tipo, m.chapa, f.nome collate database_default as nome from assmed_mudanca m, corporerm.dbo.pfunc f, assmed_planos p " & _
"where m.chapa=f.chapa collate database_default and (m.empresa=p.codigo and m.plano=p.plano) " & _
"and m.chapa='" & request.form("chapa") & "' and m.empresa='U' and '20141031' between m.ivigencia and m.fvigencia "
sqld="SELECT 'Dependente' as tipo, m.id_mud, d.dependente " & _
"from assmed_dep d, assmed_dep_mudanca m, corporerm.dbo.pfunc f, assmed_planos p " & _
"where d.chapa=m.chapa and d.nrodepend=m.nrodepend and d.chapa=f.chapa collate database_default and (m.empresa=p.codigo and m.plano=p.plano) " & _
"and d.chapa='" & request.form("chapa") & "' and m.empresa='U' and '20141031' between m.ivigencia and m.fvigencia "
sqlfinal=sqlt & " union all " & sqld
rs.Open sqlfinal, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
do while not rs.eof
if rs("tipo")="Titular" then tipo="T" else tipo="D"
%>
<tr>
	<td class=campo><input type="radio" name="id" value="<%=tipo%><%=rs("chapa")%>"></td>
	<td class=campo><%=rs("tipo")%></td>
	<td class=campo><%=rs("nome")%></td>
</tr>
<%
rs.movenext
loop
end if
rs.close
%>
</table>
<br>
<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>
<%
end if

if request.form("B1")<>"" and request.form("id")<>"" then
temp=request.form("id")
tipo=left(temp,1)
codigo=right(temp,len(temp)-1)
chapa=request.form("chapa")
sqla="SELECT f.NOME, f.CODSITUACAO, f.CHAPA, f.DATAADMISSAO, f.CODSECAO, f.codsindicato, s.DESCRICAO AS Secao, " & _
"p.dtnascimento, p.telefone1, p.telefone2, p.telefone3, p.email, p.cpf " & _
", f.codbancopagto, f.codagenciapagto, f.contapagamento, m.codigo, p.cidade " & _
"FROM corporerm.dbo.PFUNC AS f inner join corporerm.dbo.PSECAO s on f.codsecao=s.codigo " & _
"inner join corporerm.dbo.PPESSOA p on p.codigo=f.codpessoa " & _
"inner join assmed_mudanca m on m.chapa=f.chapa collate database_default " & _
"and '20141031' between ivigencia and fvigencia and m.empresa='U' "
sqlb="AND f.CHAPA='" & chapa & "' "
sql1=sqla & sqlb
rs.Open sql1, ,adOpenStatic, adLockReadOnly

session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
idade=int((now()-rs("dtnascimento"))/365.25)

%>
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td valign="center" align="left" valign=middle class="campop"><font style="font-size:16px"><b>Solicita��o de Reembolso - Seguro Sa�de</b></font></td>
	<td valign="center" align="right" valign=middle><img src="../images/logo_unimed.jpg" border="0"></td>
</tr>
<tr><td colspan=2 height=5></td></tr>
</table>
<%
if tipo="T" then
	sqlplano="SELECT codigo, plano FROM assmed_mudanca " & _
	"WHERE chapa='" & rs("chapa") & "' and '20141031' between ivigencia and fvigencia and empresa in ('U') "
elseif tipo="D" then
	sqlplano="select m.plano, m.codigo, d.dependente from assmed_dep_mudanca m, assmed_dep d " & _
	"where m.chapa=d.chapa and m.nrodepend=d.nrodepend and m.id_mud=" & codigo
end if
rs3.Open sqlplano, ,adOpenStatic, adLockReadOnly
plano=rs3("plano")
carteirinha=rs3("codigo")
if tipo="T" then usuario=rs("nome") else usuario=rs3("dependente")
rs3.close
set rs3=nothing
%>


<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" colspan=2 style=""><b>Identifica��o</td>
</tr>
<tr>
	<td width="80%" class="campor" bordercolor="#000000" style="border-top:2px solid;border-right:1px solid">
	&nbsp;Nome do Estipulante / Empresa (Somente para Planos Coletivos)</td>
	<td width="20%" class="campor" bordercolor="#000000" style="border-top:2px solid;">
	&nbsp;Data de Emiss�o</td>
</tr>
<tr>
	<td class="campop" style="border-right:1px solid" height=25>&nbsp;FUNDACAO INSTITUTO DE ENSINO PARA OSASCO</td>
	<td class="campop" style="">&nbsp;<%=formatdatetime(now(),2)%></td>
</tr></table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td width="65%" class="campor" bordercolor="#000000" style="border-top:1px solid;border-right:1px solid">
	&nbsp;Nome do Segurado Titular</td>
	<td width="35%" class="campor" bordercolor="#000000" style="border-top:1px solid;">
	&nbsp;C�digo do Segurado Titular (obrigat�rio)</td>
</tr>
<tr>
	<td class="campop" style="border-right:1px solid" height=25>&nbsp;<%=rs("nome")%></td>
	<td class="campop" style="">&nbsp;<%=rs("codigo")%></td>
</tr></table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td width="25%" class="campor" bordercolor="#000000" style="border-top:1px solid;border-right:1px solid">
	&nbsp;Munic�pio Residencial</td>
	<td width="55%" class="campor" bordercolor="#000000" style="border-top:1px solid;border-right:1px solid">
	&nbsp;E-mail (Obrigat�rio e em Letra de Forma)</td>
	<td width="20%" class="campor" bordercolor="#000000" style="border-top:1px solid;">
	&nbsp;Telefone</td>
</tr>
<tr>
	<td class="campop" style="border-right:1px solid;border-bottom:2px solid" height=25>&nbsp;<%=rs("cidade")%></td>
	<td class="campop" style="border-right:1px solid;border-bottom:2px solid" >&nbsp;<input type=text class="form_input10" value="<%=rs("email")%>" size="45"></td>
	<td class="campop" style=";border-bottom:2px solid">&nbsp;<%=rs("telefone1")%></td>
</tr></table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td height=13 style="font-family:Arial Narrow;font-size:10px"><b>*O VALOR DO REEMBOLSO PODER� SER CREDITADO PARA O TITULAR OU DEPENDENTE, DESDE QUE O CPF/MF INFORMADO SEJA DO TITULAR DA CONTA CORRENTE</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690" height=30>
<tr>
	<td class=campo style="border-bottom:1px solid" valign="middle">
	&nbsp;<img src="../images/box_0.gif" width="22" height="22" border="0" alt=""></td>
	<td class=campo style="border-bottom:1px solid" valign="middle">&nbsp;Titular</td>
	<td class=campo style="border-bottom:1px solid" valign="middle">
	&nbsp;<img src="../images/box_0.gif" width="22" height="22" border="0" alt=""></td>
	<td class=campo style="border-bottom:1px solid" valign="middle">&nbsp;Dependente</td>
	<td class=campo width=75% style="border-bottom:1px solid" valign="middle"></td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690" height=30>
<tr>
	<td class=campo style="border-bottom:1px solid" valign="middle">
	&nbsp;<img src="../images/box_0.gif" width="22" height="22" border="0" alt=""></td>
	<td class=campo style="border-bottom:1px solid" valign="middle">&nbsp;Conta Corrente</td>
	<td class=campo style="border-bottom:1px solid" valign="middle">
	&nbsp;<img src="../images/box_0.gif" width="22" height="22" border="0" alt=""></td>
	<td class=campo style="border-bottom:1px solid" valign="middle">&nbsp;Conta Poupan�a</td>
	<td class=campo style="border-bottom:1px solid" valign="middle">
	&nbsp;<img src="../images/box_0.gif" width="22" height="22" border="0" alt=""></td>
	<td class=campo style="border-bottom:1px solid" valign="middle">&nbsp;Ordem de Pagamento</td>
	<td class=campo width=45% style="border-bottom:1px solid" valign="middle"></td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" bordercolor="#000000" style="border-right:1px solid">
	&nbsp;Nome do Banco</td>
	<td class="campor" bordercolor="#000000" style="border-right:1px solid">
	&nbsp;N� Banco</td>
	<td class="campor" bordercolor="#000000" style="border-right:1px solid">
	&nbsp;Ag�ncia N�mero</td>
	<td class="campor" bordercolor="#000000" style="border-right:1px solid">
	&nbsp;D�gito</td>
	<td class="campor" bordercolor="#000000" style="border-right:1px solid">
	&nbsp;Conta Corrente N�mero</td>
	<td class="campor" bordercolor="#000000" style="border-right:1px solid">
	&nbsp;D�gito</td>
	<td class="campor" bordercolor="#000000" style="">
	&nbsp;CPF/MF*</td>
</tr>
<tr>
	<td class="campop" style="border-right:1px solid;border-bottom:2px solid" height=25>&nbsp;<input type=text class="form_input10" value="<%if rs("codbancopagto")="237" then response.write "Banco Bradesco"%>" size="15"></td>
	<td class="campop" style="border-right:1px solid;border-bottom:2px solid">&nbsp;<input type=text class="form_input10" value="<%=rs("codbancopagto")%>" size="5"></td>
	<td class="campop" style="border-right:1px solid;border-bottom:2px solid">&nbsp;<input type=text class="form_input10" value="<%=left(rs("codagenciapagto"),len(rs("codagenciapagto"))-1)%>" size="10"></td>
	<td class="campop" style="border-right:1px solid;border-bottom:2px solid">&nbsp;<input type=text class="form_input10" value="<%=right(rs("codagenciapagto"),1)%>" size="2"></td>
	<td class="campop" style="border-right:1px solid;border-bottom:2px solid">&nbsp;<input type=text class="form_input10" value="<%=left(rs("contapagamento"), len(rs("contapagamento"))-1)%>" size="12"></td>
	<td class="campop" style="border-right:1px solid;border-bottom:2px solid">&nbsp;<input type=text class="form_input10" value="<%=right(rs("contapagamento"),1)%>" size="2"></td>
	<td class="campop" style="border-bottom:2px solid">&nbsp;<input type=text class="form_input10" value="<%=left(rs("cpf"),3) & "." & mid(rs("cpf"),4,3) & "." & mid(rs("cpf"),7,3) & "-" & right(rs("cpf"),2)  %>" size="20"></td>
</tr></table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" colspan=4 style="" height=35 valign="bottom"><b>Dados do Comprovante</td>
</tr>
<tr>
	<td class="campor" colspan=1 bordercolor="#000000" style="border-top:2px solid;border-right:1px solid" align="right">
	&nbsp;O preenchimento deste campo � obrigat�rio<img src="../images/arrow.gif" width="13" height="10" border="0" alt=""></td>
	<td class="campor" bordercolor="#000000" style="border-top:2px solid;border-right:1px solid">
	Informar os �ltimos 4 d�gitos da<br>carteira do Segurado Dependente</td>
	<td class="campor" colspan=2 bordercolor="#000000" style="border-top:2px solid;">&nbsp;</td>
</tr>
<tr>
	<td class="campor" bordercolor="#000000" style="border-top:1px solid;border-right:1px solid">
	&nbsp;Nome do Segurado</td>
	<td class="campor" bordercolor="#000000" style="border-top:1px solid;border-right:1px solid">
	&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-top:1px solid;border-right:1px solid">
	&nbsp;Quantidade de Recibos</td>
	<td class="campor" bordercolor="#000000" style="border-top:1px solid;">
	&nbsp;Valor Total por Paciente</td>
</tr>
<tr>
	<td class="campop" width=47% style="border-right:1px solid;border-bottom:1px solid" height=30>&nbsp;<input type=text class="form_input10" value="<%=usuario%>" size="45"></td>
	<td class="campop" width=23% style="border-right:1px solid;border-bottom:1px solid">&nbsp;<input type=text class="form_input10" value="<%=right(carteirinha,4)%>" size="6"></td>
	<td class="campop" width=15% style="border-right:1px solid;border-bottom:1px solid">&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td>
	<td class="campop" width=15% style="border-bottom:1px solid">&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td>
</tr>
<%
for a=1 to 4
%>
<tr>
	<td class="campop" width=47% style="border-right:1px solid;border-bottom:1px solid" height=30>&nbsp;<input type=text class="form_input10" value="<%%>" size="45"></td>
	<td class="campop" width=23% style="border-right:1px solid;border-bottom:1px solid">&nbsp;<input type=text class="form_input10" value="<%%>" size="6"></td>
	<td class="campop" width=15% style="border-right:1px solid;border-bottom:1px solid">&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td>
	<td class="campop" width=15% style="border-bottom:1px solid">&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td>
</tr>
<%
next
%>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" colspan=3 style="border-top:1px solid;border-bottom:2px solid" height=30 valign="bottom"><b>Uso Exclusivo da UNIMED Seguros Sa�de S.A.</td>
</tr>
<tr>
	<td class="campor" bordercolor="#000000" style="border-top:1px solid;border-right:1px solid">
	&nbsp;NR</td>
	<td class="campor" bordercolor="#000000" style="border-top:1px solid;border-right:1px solid">
	&nbsp;Valor Avisado</td>
	<td class="campor" bordercolor="#000000" style="border-top:1px solid;">
	&nbsp;Valor Reembolsado</td>
</tr>
<tr>
	<td class="campop" width=33% style="border-right:1px solid;border-bottom:1px solid" height=30>&nbsp;<input type=text class="form_input10" value="<%%>" size="5"></td>
	<td class="campop" width=33% style="border-right:1px solid;border-bottom:1px solid">&nbsp;<input type=text class="form_input10" value="<%%>" size="6"></td>
	<td class="campop" width=33% style="border-bottom:1px solid">&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td>
</tr>
<tr>
	<td class="campor" colspan=3 bordercolor="#000000" style="border-top:1px solid;">
	&nbsp;Observa��es</td>
</tr>
<tr><td class="campop" colspan=3 style="border-bottom:1px solid"  >&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td></tr>
<tr><td class="campop" colspan=3 style="border-bottom:1px solid" height=25 >&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td></tr>
<tr><td class="campop" colspan=3 style="border-bottom:1px solid" height=25 >&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td></tr>
<tr><td class="campop" colspan=3 style="border-bottom:1px solid" height=25 >&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td></tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" bordercolor="#000000" style="border-top:0px solid;border-right:1px solid">
	&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-top:0px solid;">
	&nbsp;Carimbo e Visto do Analista</td>
</tr>
<tr>
	<td class="campop" width=55% height=35 style="border-right:1px solid;border-bottom:2px solid">&nbsp;<input type=text class="form_input10" value="<%%>" size="6"></td>
	<td class="campop" width=45% style="border-bottom:2px solid">&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td>
</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo style="background:#000000" width=5></td>
	<td class=campo>
	<b>Unimed Seguros Sa�de S.A.</b> - CNPJ/MF 04.487.255/0001-81<br>
	Alameda Ministro Rocha Azevedo, 366  CEP 01410-901 S�o Paulo SP<br>
	Atendimento Nacional: 0800 016 6633 - Atendimento ao Deficiente Auditivo: 0800 770 3611<br>
	<b>www.segurosunimed.com.br</b>
	</td>
	<td class=campo valign=top align="right"><img src="../images/ans_unimed.gif" border="0" alt=""></td>
</tr>
</table>
<br>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo><img src="../images/tesoura1.gif" width="56" height="38" border="0" alt=""></td>
	<td class=campo width=100%><hr style="border:2px #000000 dotted"></td>
</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" colspan=2 style="border-top:0px solid;border-bottom:2px solid" height=30 valign="bottom"><b>Solicita��o de Reembolso - Protocolo</td>
</tr>
<tr>
	<td class="campor" bordercolor="#000000" style="border-top:1px solid;border-right:1px solid">
	&nbsp;Nome do Segurado Titular</td>
	<td class="campor" bordercolor="#000000" style="border-top:1px solid;">
	&nbsp;Protocolo da Unimed</td>
</tr>
<tr>
	<td class="campop" width=55% style="border-right:1px solid;border-bottom:1px solid" height=30>&nbsp;<input type=text class="form_input10" value="<%=rs("nome")%>" size="50"></td>
	<td class="campop" width=45% style="border-bottom:1px solid">&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" bordercolor="#000000" style="border-top:0px solid;border-right:1px solid">
	&nbsp;Data da entrega da Solicita��o</td>
	<td class="campor" bordercolor="#000000" style="border-top:0px solid;border-right:1px solid">
	&nbsp;Qtde. Recibos</td>
	<td class="campor" bordercolor="#000000" style="border-top:0px solid;border-right:1px solid">
	&nbsp;Valor Total dos Recibos</td>
	<td class="campor" bordercolor="#000000" style="border-top:0px solid;">
	&nbsp;</td>
</tr>
<tr>
	<td class="campop" width=20% style="border-right:1px solid;border-bottom:2px solid" height=30>&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td>
	<td class="campop" width=15% style="border-right:1px solid;border-bottom:2px solid" height=30>&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td>
	<td class="campop" width=20% style="border-right:1px solid;border-bottom:2px solid" height=30>&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td>
	<td class="campop" width=45% style="border-bottom:2px solid">&nbsp;<input type=text class="form_input10" value="<%%>" size="10"></td>
</tr>
</table>

<DIV style="page-break-after:always"></DIV> 

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" colspan=3 style="border-top:0px solid;border-bottom:2px solid" height=30 valign="bottom"><b>Requisitos para Solicita��o de Reembolso do Produto Seguro Sa�de</td>
</tr>
<tr>
	<td class=campo width=48% valign=top>
	<b>CONSIDERA��ES GERAIS:</b><br>
	1) Todos os documentos devem ser originais.<br>
	2) Para seu acompanhamento, recomendamos guardar uma c�pia dos documentos apresentados.<br>
	3) A Seguradora disp�e de at� 15 dias �teis, contados da data do recebimento da <b>documenta��o completa</b>, para efetuar o pagamento das despesas cobertas contratualmente.<br>
	4) O reembolso poder� ser solicitado, informando os dados banc�rios do Titular ou se preferir, de seu Dependente. 
	Obs: Para isto, mencionar sempre os dados banc�rios e CPF/MF do Titular da Conta Corrente.<br>
	5) Os reembolsos realizados s�o creditados somente na Conta Corrente do segurado Titular e Dependente (se for paciente), 
	atrav�s de Ordem de Pagamento ou Conta Poupan�a.<br>
	6) Reembolsos entregues na Unimed Seguros at� 10h00 ser�o reembolsados atrav�s de sistema imediato.<br>
	<br>
	<b>1 � CONSULTA M�DICA</b><br>
	Recibo do m�dico ou Nota Fiscal quitada da institui��o que efetuou o atendimento.<br>
	<br>
	Conte�do do documento:<br>
	-Nome do paciente<br>
	-Data da consulta<br>
	-Valor cobrado (num�rico e por extenso)<br>
	-Descri��o do tipo de atendimento/especialidade<br>
	-No recibo, dever�o constar os dados do M�dico (nome, CPF/MF, CRM, especialidade, assinatura sobre carimbo e endere�o completo)<br>
	-Na nota fiscal, dever�o constar os dados da institui��o (nome, CNPJ/MF e endere�o completo) com carimbo do profissional respons�vel pelo atendimento e carimbo de recebido com data e/ou autentica��o mec�nica.<br>
	<br>
	<b>2 � EXAMES LABORATORIAIS E RADIOL�GICOS</b><br>
	Apresentar:<br>
	a) Nota Fiscal quitada da institui��o que efetuou o atendimento;<br>
	b) Pedido do m�dico solicitante.<br>
	<br>
	Conte�do da Nota Fiscal:<br>
	-Nome do paciente<br>
	-Data do atendimento<br>
	-Valor cobrado (num�rico e por extenso)<br>
	-Nome de cada exame realizado com o respectivo valor unit�rio<br>
	-Regi�o corp�rea (exame por imagem)<br>
	-Dados da institui��o (nome, CNPJ/MF e endere�o completo)<br>
	-Cobran�a de taxas diversas, materiais e medicamentos devem vir discriminados (nomes e valores)<br>
	<br>
	<b>3 � TERAPIAS (Fisioterapia, Radioterapia, Escleroterapia, outras).</b><br>
	Apresentar:<br>
	a) Recibo ou Nota Fiscal quitada do prestador que realizou o atendimento.<br>
	b) Relat�rio do m�dico solicitante atualizado a cada tr�s meses informando o diagn�stico, tempo de exist�ncia da doen�a e tratamento proposto.<br>
	c) Informar as datas das sess�es realizadas (Fisio, Fono, Psico, Etc).<br>
	<br>
	Conte�do do recibo ou da Nota Fiscal:<br>
	-Nome do paciente<br>
	-Data da consulta<br>
	-Valor cobrado (num�rico e por extenso)<br>
	-Descri��o do tipo de atendimento<br>
	-No recibo, dever�o constar os dados do Prestador (nome, CPF/MF, n�mero de inscri��o no Conselho Regional (CRM/CREDITO/CRP/CRF), especialidade, assinatura sobre carimbo e endere�o completo).<Br>
	-Na nota fiscal, dever�o constar os dados da institui��o (nome, CNPJ/MF e endere�o completo).<br>
	<br>
	</td>
	<td class=campo width=4%></td>
	<td class=campo width=48% valign=top>
	<b>4 � DESPESAS HOSPITALARES</b><br>
	Apresentar:<br>
	a) Relat�rio emitido pelo m�dico assistente informando diagn�stico, tempo de exist�ncia da doen�a, tratamento realizado, per�odo de interna��o e quantidade de visitas hospitalares.<br>
	b) C�pia do(s) laudo(s) se for(em) realizado(s) exame(s) anatomo(s) patol�gico(s) ou polissonografia(s).<br>
	c) Recibos ou Nota Fiscal dos profissionais (cirurgi�o, auxiliar, anestesista, instrumentador, assist�ncia ao rec�m-nascido e visitas hospitalares).<br>
	d) Nota Fiscal quitada da entidade hospitalar.<br>
	<br>
	Conte�do do(s) recibo(s) ou da Nota Fiscal:<br>
	-Nome do paciente<br>
	-Data do evento<br>
	-Valor cobrado (num�rico e por extenso)<br>
	-Recibos de honor�rios m�dicos dever�o ser individualizados e constar os dados do profissional (nome, CPF/MF, CRM, fun��o exercida no evento e assinatura sobre carimbo).<br>
	-Para Honor�rios apresentados em Nota Fiscal, dever�o constar os dados da institui��o (nome, CNPJ/MF e endere�o completo) e descri��o da equipe m�dica (nome, CRM, posi��o e valor cobrado para cada profissional).<br>
	-Na Nota Fiscal hospitalar dever� constar data do per�odo da interna��o, descritivo com valores e quantidades individuais das despesas, taxas, servi�os complementares, materiais e medicamentos e vir acompanhada do recibo de quita��o ou carimbo de recebido com data e/ou autentica��o mec�nica.<Br>
	<br>
	<b>5 � PR�TESES E �RTESES LIGADAS AO ATO CIRURGICO</b><br>
	Verifique em seu contrato a abrang�ncia desta cobertura. Havendo cobertura, apresentar:<br>
	<br>
	a) Nota Fiscal quitada do prestador.<br>
	b) C�pia da Nota Fiscal do Fornecedor<Br>
	c) Relat�rio m�dico justificando a implanta��o do aparelho.<br>
	<br>
	Conte�do da Nota Fiscal:<br>
	-Nome do paciente<Br>
	-Data do atendimento<br>
	-Valor cobrado (num�rico e por extenso)<br>
	-Descri��o do tipo do aparelho<br>
	-Dados da institui��o (nome, CNPJ/MF e endere�o completo)<br>
	<br>
	<b>6 � REMO��O EM AMBUL�NCIA</b><br>
	Apresentar:<br>
	a) Nota Fiscal quitada do prestador.<br>
	b) Relat�rio m�dico informando o diagn�stico do paciente e necessidade da remo��o.<br>
	<br>
	Conte�do da Nota Fiscal:<br>
	-Nome do paciente<Br>
	-Data do atendimento<br>
	-Valor cobrado (num�rico e por extenso)<br>
	-Dados do prestador (nome, CNPJ/MF e endere�o completo)<br>
	-Descri��o do total de quil�metros rodados, valor unit�rio da quilometragem, local de partida e destino, tipo de ambul�ncia (UTI ou simples)<br>
	-Se houver cobran�a de Taxas/Honor�rios, discriminar.	
	</td>
</tr>	
<tr>
	<td class=campo colspan=3 style="border-top:1px solid;border-bottom:2px solid;font-family:Arial Narrow" height=40 valign="middle">
	<b>IMPORTANTE: Todas solicita��es de reembolso passam por an�lise t�cnica e m�dica. Havendo necessidade, a Seguradora se reserva o direito de solicitar documentos
	ou informa��es complementares para melhor classifica��o do procedimento de acordo com o plano contratato.</b>
	</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo style="background:#000000" width=5></td>
	<td class=campo>
	<b>Unimed Seguros Sa�de S.A.</b> - CNPJ/MF 04.487.255/0001-81<br>
	Alameda Ministro Rocha Azevedo, 366  CEP 01410-901 S�o Paulo SP<br>
	Atendimento Nacional: 0800 016 6633 - Atendimento ao Deficiente Auditivo: 0800 770 3611<br>
	<b>www.segurosunimed.com.br</b>
	</td>
	<td class=campo valign=top align="right"><img src="../images/ans_unimed.gif" border="0" alt=""></td>
</tr>
</table>



<%
rs.close
set rs=nothing
%>
</table>
<%
set rs=nothing
set rs2=nothing
end if ' 

conexao.close
set conexao=nothing
%>
</body>
</html>