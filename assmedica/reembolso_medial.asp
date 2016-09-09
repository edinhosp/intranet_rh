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
<title>Reembolso Medial</title>
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
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para reembolso
<form method="POST" action="reembolso_medial.asp" name="form">
<%
sqla="SELECT m.chapa, f.nome from assmed_mudanca m, corporerm.dbo.pfunc f, assmed_planos p " & _
"where m.chapa=f.chapa collate database_default and (m.empresa=p.codigo and m.plano=p.plano) " & _
"and m.empresa='M' and p.reemb<>0 and '07/31/2010' between m.ivigencia and m.fvigencia order by f.nome"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=request.form("chapa")%>">
<select name="nome" class=a onchange="nome1()">
	<option value="0">Selecione o funcionário</option>
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
"and m.chapa='" & request.form("chapa") & "' and m.empresa='M' and p.reemb<>0 and '07/31/2010' between m.ivigencia and m.fvigencia "
sqld="SELECT 'Dependente' as tipo, m.id_mud, d.dependente " & _
"from assmed_dep d, assmed_dep_mudanca m, corporerm.dbo.pfunc f, assmed_planos p " & _
"where d.id_dep=m.id_dep and d.chapa=f.chapa collate database_default and (m.empresa=p.codigo and m.plano=p.plano) " & _
"and d.chapa='" & request.form("chapa") & "' and m.empresa='M' and p.reemb<>0 and '07/31/2010' between m.ivigencia and m.fvigencia "
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
", f.codbancopagto, f.codagenciapagto, f.contapagamento " & _
"FROM corporerm.dbo.PFUNC AS f, corporerm.dbo.PSECAO s, corporerm.dbo.PPESSOA p " & _
"WHERE f.CODSECAO = s.CODIGO and p.codigo=f.codpessoa "
sqlb="AND f.CHAPA='" & chapa & "' "
sql1=sqla & sqlb
rs.Open sql1, ,adOpenStatic, adLockReadOnly

session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
idade=int((now()-rs("dtnascimento"))/365.25)

%>
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=100>
<tr><td valign="center" align="left"  valign=middle><img src="../images/medial2.gif" border="0"></td>
	<td valign="center" align="right" valign=middle><font size=2><b>SOLICITAÇÃO DE REEMBOLSO</b></font></td>
</tr>
<tr><td align="left" class="campor" colspan=2><b>Os recibos ou notas fiscais originais devem ser apresentados à Medial Saúde, 
no máximo, até 90 (noventa) dias, contados a partir da data do evento.
</tr>
<tr><td colspan=2 class="campop" align="center"><b>USUÁRIO</td></tr>
</table>
<%
if tipo="T" then
	sqlplano="SELECT codigo, plano FROM assmed_mudanca " & _
	"WHERE chapa='" & rs("chapa") & "' and '07/31/2010' between ivigencia and fvigencia and empresa in ('M') "
elseif tipo="D" then
	sqlplano="select m.plano, m.codigo, d.dependente from assmed_dep_mudanca m, assmed_dep d " & _
	"where m.id_dep=d.id_dep and m.id_mud=" & codigo
end if
rs3.Open sqlplano, ,adOpenStatic, adLockReadOnly
plano=rs3("plano")
carteirinha=rs3("codigo")
if tipo="T" then usuario=rs("nome") else usuario=rs3("dependente")
rs3.close
set rs3=nothing
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=40>
<tr>
	<td width="60%" class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Empresa</td>
	<td width=1% class="campor">&nbsp;</td>
	<td width="40%" class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Nome do Plano</td>
</tr>
<tr>
	<td width="60%" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;FUNDACAO INSTITUTO DE ENSINO PARA OSASCO</td>
	<td width=1% class="campor">&nbsp;</td>
	<td width="40%" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<%=plano%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=3><tr><td></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=40>
<tr>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Nome do Titular do Plano</td>
</tr>
<tr>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<%=rs("nome")%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=3><tr><td></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=40>
<tr>
	<td colspan=16 class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Número do Cartão de Identificação Medial Saúde</td>
	<td width=1% class="campor">&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Nome do Usuário Solicitante do Reembolso</td>
</tr>
<tr>
<%for a=1 to 16%>
	<td align="center" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	<%=mid(carteirinha,a,1)%></td>
<%next%>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	<input type="text" value="<%=usuario%>" size=50 class=form_input10></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=3><tr><td></td></tr></table>
  
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=40>
<tr>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;DDD - Tel. Residencial</td>
	<td width=1% class="campor">&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;DDD - Tel. Comercial</td>
</tr>
<tr>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;(11) <%=rs("telefone1")%></td>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;(11) <%if rs("telefone3")="" or isnull(rs("telefone3")) then response.write "3651-9999" else response.write rs("telefone3")%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=5><tr><td></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=40>
<tr>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;DDD - Celular</td>
	<td width=1% class="campor">&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;E-mail</td>
</tr>
<tr>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;(11) <%=rs("telefone2")%></td>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<%=rs("email")%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=3><tr><td></td></tr></table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td colspan=1 class="campop" align="center"><b>DEPÓSITO</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=40>
<tr>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Nome do Correntista</td>
	<td width=1% class="campor">&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;CPF/CNPJ</td>
</tr>
<tr>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<input type=text class="form_input10" value="<%=rs("nome")%>" size="45"></td>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<input type=text class="form_input10" value="<%=left(rs("cpf"),3) & "." & mid(rs("cpf"),4,3) & "." & mid(rs("cpf"),7,3) & "-" & right(rs("cpf"),2)  %>" size="15"></td>
</tr></table>
  
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=3><tr><td></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=40>
<tr>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Banco</td>
	<td width=1% class="campor">&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Nº do Banco</td>
	<td width=1% class="campor">&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Nº Agência</td>
	<td width=1% class="campor">&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Dígito</td>
	<td width=1% class="campor">&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Nº Conta Corrente</td>
	<td width=1% class="campor">&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Dígito</td>
</tr>
<tr>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<input type=text class="form_input10" value="<%if rs("codbancopagto")="237" then response.write "Banco Bradesco"%>" size="15"></td>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<input type=text class="form_input10" value="<%=rs("codbancopagto")%>" size="5"></td>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<input type=text class="form_input10" value="<%=left(rs("codagenciapagto"),len(rs("codagenciapagto"))-1)%>" size="4"></td>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<input type=text class="form_input10" value="<%=right(rs("codagenciapagto"),1)%>" size="2"></td>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<input type=text class="form_input10" value="<%=left(rs("contapagamento"), len(rs("contapagamento"))-1)%>" size="8"></td>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<input type=text class="form_input10" value="<%=right(rs("contapagamento"),1)%>" size="2"></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=3><tr><td></td></tr></table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td colspan=1 class="campop" align="center"><b>DADOS PARA REEMBOLSO</td></tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=3><tr><td></td></tr></table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td colspan=6 height=10></td></tr>
<tr>
	<td width=1% class="campor">&nbsp;</td>
	<td height=25 class=campo align="center" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:0;border-bottom-width:1">
	&nbsp;Nome do Prestador</td>
	<td class=campo align="center" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;CNPJ/CPF</td>
	<td class=campo align="center" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;Telefone</td>
	<td class=campo align="center" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:1;border-bottom-width:1">
	&nbsp;C.R.M.</td>
	<td width=1% class="campor">&nbsp;</td>
</tr>
<% for a=1 to 4 %>
<tr>
	<td width=1% class="campor">&nbsp;</td>
	<td height=25 bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:0;border-bottom-width:1">
	<input type=text class="form_input10" value="" size="54"></td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:0;border-bottom-width:1">
	<input type=text class="form_input10" value="" size="20"></td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:0;border-bottom-width:1">
	<input type=text class="form_input10" value="" size="9"></td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:0;border-bottom-width:1">
	<input type=text class="form_input10" value="" size="8"></td>
	<td width=1% class="campor">&nbsp;</td>
</tr>
<% next %>
<tr><td colspan=6 height=10></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td width=1% class="campor">&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:0;border-bottom-width:0">
	&nbsp;Valor Total dos Recibos Anexos</td>
	<td width=1% class="campor">&nbsp;</td>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:0;border-bottom-width:0">
	&nbsp;Quantidade dos Recibos Anexo</td>
	<td width=1% class="campor">&nbsp;</td>
	<td colspan=6 class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:0;border-bottom-width:0">
	&nbsp;Data</td>
	<td width=1% class="campor">&nbsp;</td>
</tr>
<tr>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:0;border-bottom-width:1">
	<input type=text class="form_input10" value="" size="15">&nbsp;</td>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:0;border-bottom-width:1">
	<input type=text class="form_input10" value="" size="8">&nbsp;</td>
	<td width=1% class="campor">&nbsp;</td>
	&nbsp;</td>
<%for a=1 to 6%>
	<td align="center" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:1;border-bottom-width:1">
	<%=mid(numzero(day(now),2)&numzero(month(now),2)&numzero(right(year(now),2),2),a,1)%></td>
<%next%>
	<td width=1% class="campor">&nbsp;</td>
</tr></table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td width=1% class="campor">&nbsp;</td>
	<td width=40% class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:0;border-bottom-width:0">
	&nbsp;Assinatura do Usuário Titular do Plano</td>
	<td width=1% class="campor">&nbsp;</td>
	<td width=30% class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:0;border-bottom-width:0">
	&nbsp;Local</td>
	<td width=28% class="campor">&nbsp;</td>
</tr>
<tr>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:0;border-bottom-width:1">
	&nbsp;</td>
	<td width=1% class="campor">&nbsp;</td>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:0;border-left-width:0;border-bottom-width:1">
	&nbsp;</td>
	<td width=1% class="campor">&nbsp;</td>
	&nbsp;</td>
</tr>
<tr><td colspan=5 height=10></td></tr>
</table>
</td></tr></table>

<br> 
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=5>
<tr><td style="border-style:dotted;border-top-width:1;border-right-width:0;border-left-width:0;border-bottom-width:0">
</td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td colspan=1 class="campop" align="center"><b>SOLICITAÇÃO DE REEMBOLSO - Protocolo de Recebimento</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=40>
<tr>
	<td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Nome do Usuário Solicitante do Reembolso</td>
</tr>
<tr>
	<td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<%=rs("nome")%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=5><tr><td></td></tr></table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td valign=top class=campo>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="200" valign=top>
<tr><td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Quantidade Total dos Recibos</td>
</tr>
<tr><td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<input type=text class="form_input10" value="" size="8"></td>
</tr>
<tr><td width=1% height=5 class="campor"></td>
</tr>
<tr><td class="campor" bordercolor="#000000" style="border-style:solid;border-top-width:1;border-right-width:1;border-left-width:1;border-bottom-width:0">
	&nbsp;Valor Total</td>
</tr>
<tr><td bordercolor="#000000" style="border-style:solid;border-top-width:0;border-right-width:1;border-left-width:1;border-bottom-width:1">
	&nbsp;<input type=text class="form_input10" value="" size="15"></td>
</tr>
<tr><td>&nbsp;</td></tr>
</table>

<br>
<br>
<br><b>
Enviar para:<br>
Medial Saúde<br>
A/C Setor de Reembolso<br>
Av. Marquês de São Vicente, 600<br>
Barra Funda<br>
São Paulo/SP<br>
CEP 01139-002
</td>
<td width=1% class="campor">&nbsp;</td>
<td valign=top class="campor">
<p style="margin-top:0;margin-bottom:0;text-align="center";font-size:8pt"><b>OBSERVAÇÕES IMPORTANTES:
<p style="margin-top:0;margin-bottom:0;text-align="left";font-size:8pt">
<b>- Os recibos ou notas fiscais devem conter:</b><br>
&nbsp;&nbsp;CPF/CNPJ do prestador; CRM; assinatura sob carimbo; nome do Usuário; CID 10; descrição do serviço prestado (ex.: consulta,
hemograma, RX do tórax, etc); data de prestação do serviço;<br>
- Nos tratamentos de Fisioterapia, Acunpuntura, Fonoaudiologia e Psicoterapia, anexar Relatório Médico com datas das sessões 
e Ficha de Frequência;<br>
- Nos exames de Ressonância, Tomografia, Cintilografia, Hemodinâmica, Mapeamento Cerebral ou outros procedimentos de alta
complexidade, anexar Relatório Médico contendo diagnóstico, CID 10 e código AMB para análise clínica.<br>
- Nas Internações Clínicas ou Cirúrgicas, anexar Relatório Médico Pós Cirúrgico contendo diagnóstico, CID 10 e código AMB para análise médica;<br>
- Despesas hospitalares devem conter discriminação das taxas, materiais e medicamentos;<br>
<b>- A falta de informação na documentação apresentada poderá acarretar a devolução da documentação para regularização;<br>
- O prazo para reembolso será de até 30 (trinta) dias, contados a partir da data de entrega da documentação completa na Medial Saúde.<br>
- Para evitar extravios, recomendadmos o envio de Solicitações de Reembolso e Recibos / Notas Fiscais via portador ou por carta registrada (AR).
</td>
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