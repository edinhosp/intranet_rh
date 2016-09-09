<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=false
	Server.ScriptTimeout = 1200
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a42")="N" or session("a42")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Alteração Cadastral</title>
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
<body style="margin-left:40px">
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
rs.CursorLocation = adUseClient
rs2.CursorLocation = adUseClient
rs3.CursorLocation = adUseClient

if request.form<>"" then
	idchapa=request.form("D1")
	if idchapa="Todos" then
		sqld=""
	elseif idchapa="Narciso" then
		sqld=" and left(f.codsecao,2)='01' and f.codtipo='N' "
	elseif idchapa="Narciso1" then
		sqld=" and left(f.codsecao,4)='01.1' and f.codtipo='N' "
	elseif idchapa="Narciso2" then
		sqld=" and left(f.codsecao,4)='01.2' and f.codtipo='N' "
	elseif idchapa="Narciso3" then
		sqld=" and left(f.codsecao,4)='01.3' and f.codtipo='N' "
	elseif idchapa="Yara" then
		sqld=" and left(f.codsecao,2)='03' and f.codtipo='N' "
	elseif idchapa="Yara1" then
		sqld=" and left(f.codsecao,4)='03.1' and f.codtipo='N' "
	elseif idchapa="Yara2" then
		sqld=" and left(f.codsecao,4)='03.2' and f.codtipo='N' "
	elseif idchapa="Yara3" then
		sqld=" and left(f.codsecao,4)='03.3' and f.codtipo='N' "
	elseif idchapa="Wilson" then
		sqld=" and left(f.codsecao,2)='04' and f.codtipo='N' "
	elseif idchapa="EstagioY" then
		sqld=" and f.codtipo='T' and f.codsecao<>'03.1.999' and left(f.codsecao,2)='03' "
	elseif idchapa="EstagioN" then
		sqld=" and f.codtipo='T' and f.codsecao<>'03.1.999' and left(f.codsecao,2)='01' "
	elseif idchapa="EstagioW" then
		sqld=" and f.codtipo='T' and f.codsecao<>'03.1.999' and left(f.codsecao,2)='04' "
	elseif idchapa="Falta" then
		sqld=" and f.chapa not in (select chapa from zselecao where sessao='0') and f.admissao<'12/31/2005' "
	else
		sqld=" and f.chapa='" & idchapa & "'"
	end if
	
	sqlc="SELECT top 500 * from qry_funcionarios f " & _
	"where f.codsituacao in ('A','F','E') and f.chapa>'00000' "
	sqle="order by f.codsecao, f.chapa "
	sqlb=sqlc & sqld & sqle
	'response.write sqlb
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
	temp=0
	titulo=rs("chapa") & " - " & rs("nome")
else
	temp=1
end if

if temp=1 then
	sqla="SELECT chapa, nome from corporerm.dbo.pfunc f where f.codsituacao<>'D' and f.codsecao<>'03.1.999' " & _
	"order by nome"
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<p class=titulo>Ficha de Alteração Cadastral&nbsp;<%=titulo %>
<br>
<form method="POST" action="rh_cadastral.asp" name="form">
<p><select size="1" name="D1">
	<option value="Todos">Todos</option>
	<option value="Falta">Reemissão dos que não entregaram</option>
	<option value="Narciso">Todos-Campus Narciso</option>
	<option value="Narciso1">Todos-Campus Narciso-Administrativo</option>
	<option value="Narciso2">Todos-Campus Narciso-Adm.Acadêmico</option>
	<option value="Narciso3">Todos-Campus Narciso-Acadêmico</option>
	<option value="Yara">Todos-Campus V.Yara</option>
	<option value="Yara1">Todos-Campus V.Yara-Administrativo</option>
	<option value="Yara2">Todos-Campus V.Yara-Adm.Acadêmico</option>
	<option value="Yara3">Todos-Campus V.Yara-Acadêmico</option>
	<option value="Wilson">Todos-Campus Jd.Wilson</option>
	<option value="EstagioY">Estagiários-Vila Yara</option>
	<option value="EstagioN">Estagiários-Narciso</option>
	<option value="EstagioW">Estagiários-Jd.Wilson</option>
<%
rs.movefirst:do while not rs.eof
%>
	<option value="<%=rs("chapa")%>"><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<br>
<input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>

<!-- impressão do documento -->
<%
end if

if temp<>1 then ' temp=0

'******************************
'response.write "<table border='1' cellpadding='0' cellspacing='1' style='border-collapse:collapse' width='650'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
'next
'response.write "</tr>"
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor">" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext
'loop
'response.write "</table>"
'rs.close
'response.write "<p>"
'*****************************

rs.movefirst
do while not rs.eof 
%>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=650 height="1000">
<tr><td class="campop" align="left" valign=top>
<!-- fim borda -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left"><b>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO<br></td><td class="campop" align="right"><%=now()%></tr></tr>
<tr><td class="campop" align="center" colspan=2><b><font size=3>FICHA DE ATUALIZAÇÃO CADASTRAL<br></td></tr>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>DADOS PESSOAIS</td>
	<td class="campop" align="right"><b><font size=2><%=rs("codsecao")%> - <%=rs("secao")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="center" height="20px" width=70><i><b>Código</td>
	<td class=campo align="center"><i><b>Nome Completo</td>
	<td class=campo align="center"><i><b>Apelido</td>
</tr>
<tr><td class=campo align="center" valign="middle" rowspan=2 style="border:1px solid #000000"><b>&nbsp;<%=rs("chapa")%></td>
	<td class="campop" height="20px" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;<b><%=rs("nome")%></td>
	<td class=campo align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("nome")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("apelido")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="center" height="20px" width="180px" valign="bottom"><i><b>Estado Civil</td>
	<td class=campo align="center" width="180px" valign="bottom"><i><b>Cidade Nascimento</td>
	<td class=campo align="center" valign="bottom"><i><b>Estado Nascimento</td>
	<td class=campo align="center" valign="bottom"><i><b>Data Nascimento</td>
</tr>
<%diab=day(rs("dtnascimento")):mesb=monthname(month(rs("dtnascimento"))):anob=year(rs("dtnascimento"))%>
<tr><td class=<%if isnull(rs("estadocivil")) then response.write "fundor" else response.write "campor"%> height="20px" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class=<%if isnull(rs("naturalidade")) then response.write "fundor" else response.write "campor"%> align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class=<%if isnull(rs("estadonatal")) then response.write "fundor" else response.write "campor"%> align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("estcivil")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("naturalidade")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("estadonatal")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=diab & "/" & mesb & "/" & anob%></td>
</tr>
</table>

<%
ehfieo=instr(1,rs("email"),"unifieo",1)
if ehfieo>0 then aviso1="<b><--------Informe um <br>email pessoal</b>" else aviso1=""
%>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class=campo align="center" width=200 height="20px" valign="bottom"><i><b>Escolaridade</td>
	<td class=campo align="center" valign="bottom"><i><b>E-mail de contato</td>
</tr>
<tr><td class="campor" height="20px" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campor" align="right" valign="top" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<%=aviso1%>&nbsp;</td>
</tr>
<tr><td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("instrucao")%></td>
	<td class=fundor align="left" width="300px" style="border:1px dotted #000000"><b>&nbsp;<%=rs("email")%></td>
</tr>
</table>

<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>ENDEREÇO / TELEFONES</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="center" valign="bottom"><i><b>Tipo</td>
	<td class=campo align="center" valign="bottom"><i><b>Endereço</td>
	<td class=campo align="center" valign="bottom" width="100px"><i><b>Número</td>
	<td class=campo align="center" valign="bottom" width="150px"><i><b>Complemento</td>
</tr>
<tr><td class="campor" height="20px" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campor" height="20px" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class=<%if isnull(rs("numero")) then response.write "fundor" else response.write "campor"%> align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=primeironome(rs("rua"))%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("rua")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("numero")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("complemento")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="center" height="20px" valign="bottom"><i><b>Bairro</td>
	<td class=campo align="center" valign="bottom"><i><b>Cidade</td>
	<td class=campo align="center" valign="bottom" width="70px"><i><b>Estado</td>
	<td class=campo align="center" valign="bottom" width="100px"><i><b>CEP</td>
</tr>
<tr><td class=<%if isnull(rs("bairro")) then response.write "fundor" else response.write "campor"%> height="20px" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class=<%if isnull(rs("cep")) then response.write "fundor" else response.write "campor"%> align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("bairro")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("cidade")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("estado")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("cep")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="center" height="20px" valign="bottom"><i><b>Residência</td>
	<td class=campo align="center" valign="bottom"><i><b>Celular</td>
	<td class=campo align="center" valign="bottom"><i><b>Rec./Coml.</td>
	<td class=campo align="center" valign="bottom"><i><b>Telefone FAX</td>
</tr>
<tr><td class="campor" height="20px" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("telefone1")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("telefone2")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("telefone3")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("fax")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="center" height="20px" valign="bottom"><i><b>Informação sobre a Residência</td>
	<td class=campo align="center" valign="bottom"><i><b>Informação sobre Financimento</td>
</tr>
<tr><td class="campop" height="25px" align="center" valign="middle" style="border-bottom:1px solid #000000;border-left:1px solid #000000">
	A residência é própria?&nbsp;<img src="..\images\bullet2.gif" border="0" alt="">&nbsp;Sim
	&nbsp;<img src="..\images\bullet2.gif" border="0" alt="">&nbsp;Não
	&nbsp;</td>
	<td class="campop" align="center" valign="middle" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	Foi financiada pelo FGTS?&nbsp;<img src="..\images\bullet2.gif" border="0" alt="">&nbsp;Sim
	&nbsp;<img src="..\images\bullet2.gif" border="0" alt="">&nbsp;Não
	&nbsp;</td>
</tr>
<tr><td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%%></td>
</tr>
</table>


<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>DOCUMENTOS</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="left" colspan=4><b>Carteira Profissional</td>
	<td class=campo align="left" colspan=3><b>Título de Eleitor</td></tr>
<tr><td class=campo align="left" valign="bottom"><i><b>Número</td>
	<td class=campo align="left" valign="bottom"><i><b>Série</td>
	<td class=campo align="left" valign="bottom" width="100px"><i><b>Data Emissão</td>
	<td class=campo align="left" valign="bottom" width="70px"><i><b>UF Emissão</td>
	<td class=campo align="left" valign="bottom"><i><b>Número</td>
	<td class=campo align="left" valign="bottom"><i><b>Zona</td>
	<td class=campo align="left" valign="bottom"><i><b>Seção</td>
</tr>
<tr><td class="<%if isnull(rs("carteiratrab")) then response.write "fundor" else response.write "campor"%>" height="20px" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("seriecarttrab")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("dtcarttrab")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("ufcarttrab")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("tituloeleitor")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("zonatiteleitor")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("secaotiteleitor")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("carteiratrab")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("seriecarttrab")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("dtcarttrab")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("ufcarttrab")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("tituloeleitor")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("zonatiteleitor")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("secaotiteleitor")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="left" colspan=4><b>Identidade (R.G.)</td>
	<td class=campo align="left" colspan=3><b>Habilitação (CNH)</td></tr>
<tr><td class=campo align="left" valign="bottom"><i><b>Número</td>
	<td class=campo align="left" valign="bottom" width="100px"><i><b>Data Emissão</td>
	<td class=campo align="left" valign="bottom" width="100px"><i><b>Orgão Emissor</td>
	<td class=campo align="left" valign="bottom" width="70px"><i><b>UF Emissão</td>
	<td class=campo align="left" valign="bottom" width="130px"><i><b>Número</td>
	<td class=campo align="left" valign="bottom" width="70px"><i><b>Categoria</td>
	<td class=campo align="left" valign="bottom"><i><b>Validade</td>
</tr>
<tr><td class="<%if isnull(rs("cartidentidade")) then response.write "fundor" else response.write "campor"%>" height="20px" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("dtemissaoident")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("orgemissorident")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("ufcartident")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("cartmotorista")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("tipocarthabilit")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("dtvenchabilit")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("cartidentidade")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("dtemissaoident")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("orgemissorident")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("ufcartident")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("cartmotorista")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("tipocarthabilit")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("dtvenchabilit")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="left" colspan=2><b>Reservista</td>
	<td class=campo align="left" colspan=1><b>C.P.F.</td>
	<td class=campo align="left" colspan=4><b>Orgão de classe.</td>
	</tr>
<tr><td class=campo align="left" valign="bottom"><i><b>Número</td>
	<td class=campo align="left" valign="bottom" width="40px"><i><b>Categoria</td>
	<td class=campo align="left" valign="bottom" width="90px"><i><b>Número</td>
	<td class=campo align="left" valign="bottom" width="110px"><i><b>Número</td>
	<td class=campo align="left" valign="bottom" width="130px"><i><b>Orgão</td>
	<td class=campo align="left" valign="bottom" width="70px"><i><b>Expedição</td>
	<td class=campo align="left" valign="bottom"><i><b>Validade</td>
</tr>
<tr><td class="<%if isnull(rs("certifreserv")) or rs("sexo")="F" then response.write "fundor" else response.write "campor"%>" height="20px" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("categmilitar")) or rs("sexo")="F" then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("cpf")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>

	<td class="<%if isnull(rs("regprofissional")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("regprofissional")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("regprofissional")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="<%if isnull(rs("regprofissional")) then response.write "fundor" else response.write "campor"%>" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("certifreserv")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("categmilitar")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("cpf")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=rs("regprofissional")%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%%></td>
</tr>
</table>

<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>FILIAÇÃO</td></tr>
</table>
<%
sqla="SELECT D.nome FROM corporerm.dbo.PFDEPEND D " & _
"WHERE D.CHAPA='" & rs("chapa") & "' and grauparentesco='6' ORDER BY D.nrodepend"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then pai=rs3("nome") else pai="&nbsp;"
rs3.close
sqla="SELECT D.nome FROM corporerm.dbo.PFDEPEND D " & _
"WHERE D.CHAPA='" & rs("chapa") & "' and grauparentesco='7' ORDER BY D.nrodepend"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then mae=rs3("nome") else mae="&nbsp;"
rs3.close
%>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="center" valign="bottom" width="50%"><i><b>Nome da Mãe</td>
	<td class=campo align="center" valign="bottom" width="50%"><i><b>Nome do Pai</td>
</tr>
<tr><td class="campor" height="20px" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">&nbsp;</td>
</tr>
<tr><td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=mae%></td>
	<td class=fundor align="left" style="border:1px dotted #000000"><b>&nbsp;<%=pai%></td>
</tr>
</table>



<DIV align="center"><B>--------------> NÃO SE ESQUEÇA DE PREENCHER O VERSO <--------------</B></DIV>
<!-- fim borda -->
</td></tr>
</table>
<%
response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
%>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=650 height="1000">
<tr><td class="campop" align="left" valign=top height=100%>
<!-- fim borda -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left"><b>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO<br></td><td class="campop" align="right">Página 2</tr></tr>
<tr><td class="campop" align="center" colspan=2><b><font size=3>FICHA DE ATUALIZAÇÃO CADASTRAL<br></td></tr>
<tr><td class="campop" align="right" colspan=2><b><font size=2><%=rs("chapa")%> - <%=rs("nome")%></td></tr>
</table>


<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>FAMÍLIA E DEPENDENTES LEGAIS</td></tr>
</table>
<%
FaltaCPF=""
sqla="SELECT D.*, MAE, p.plano FROM corporerm.dbo.PFDEPEND D LEFT JOIN corporerm.dbo.PFDEPENDCOMPL C ON D.CHAPA=C.CHAPA AND D.NRODEPEND=C.NRODEPEND " & _
"left join (select distinct chapa, nrodepend, plano='S' from assmed_dep_mudanca where GETDATE() between ivigencia and fvigencia and empresa not in ('UC','IP')) p on p.chapa=d.CHAPA collate database_default and p.nrodepend=d.NRODEPEND " & _
"WHERE D.CHAPA='" & rs("chapa") & "' and grauparentesco not in ('6','7','P','G') ORDER BY D.nrodepend "
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr>
	<td class=titulo width=200>Nome</td>
	<td class=titulo>Parentesco</td>
	<td class=titulo>Data Nasc.</td>
	<td class=titulo>Sexo</td>
	<td class=titulo>IRRF</td>
	<td class=titulo width=200>Nome da Mãe do dependente</td>
</tr>
<%
if rs3.recordcount>0 then
rs3.movefirst:do while not rs3.eof
sql="select descricao from corporerm.dbo.pcodparent where codcliente='" & rs3("grauparentesco") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then parentesco=trim(rs2("descricao")) else parentesco=""
if parentesco="Filho(a) Valido" and rs3("sexo")="M" then parentesco="Filho"
if parentesco="Filho(a) Valido" and rs3("sexo")="F" then parentesco="Filha"
rs2.close
if rs3("plano")="S" and isnull(rs3("CPF")) then FaltaCPF=FaltaCPF & "<b>---->Informar CPF de " & rs3("nome") & ": _____________________<br></b>"
%>
<tr>
	<td class=campo align="left"><%=rs3("nome")%></td>
	<td class=campo align="left"><%=parentesco%></td>
	<td class=campo align="center"><%=rs3("dtnascimento")%></td>
	<td class=campo align="center"><%=rs3("sexo")%></td>
	<td class=campo align="center"><%if rs3("incirrf")=1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=campo align="left"><%=rs3("mae")%></td>
</tr>
<%
rs3.movenext:loop
end if 'recordcount
rs3.close
for a=1 to 2
%>
<tr>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
<%
next
%>
<tr>
	<td class="campop" align="left" colspan=6><%=FaltaCPF%></td>
</tr>
</table>
<p style="margin-top:0;margin-bottom:0;font-size:9pt;text-align:right"><b>Só é considerado dependente para Imposto de Renda, aqueles com <font face='Wingdings'>ü</font> no IRRF.
<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>VEÍCULOS</td></tr>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b>A atualização do cadastro de veículos é essencial para
a emissão do novo crachá de estacionamento.</td></tr>
</table>

<%
sqla="select id_veiculo, marca,modelo, ano, cor, placa, dtcadastro, dttermino " & _
"from veiculos where chapa='" & rs("chapa") & "' and dttermino is null " & _
"order by dtcadastro, placa "
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=titulop>Marca</td>
	<td class=titulop>Modelo</td>
	<td class=titulop>Ano</td>
	<td class=titulop width=80>Cor</td>
	<td class=titulop>Placa</td>
	<td class=titulop>Cadastro</td>
</tr>
<%
if rs3.recordcount>0 then
rs3.movefirst:do while not rs3.eof
%>
<tr><td class="campop"><%=rs3("marca")%></td>
	<td class="campop"><%=rs3("modelo")%></td>
	<td class="campop"><%=rs3("ano")%></td>
	<td class="campop"><%=rs3("cor")%></td>
	<td class="campop" nowrap><%=rs3("placa")%></td>
	<td class="campop"><%=rs3("dtcadastro")%></td>
</tr>
<%
rs3.movenext:loop
end if ' recordcount
rs3.close
for a=1 to 2
%>
<tr>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
<% 
next

if rs("aposentado")=1 then 
	aposs="../images/bolax.gif"
	aposn="../images/bola.gif"
else 
	aposs="../images/bola.gif"
	aposn="../images/bolax.gif"
end if
%>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>VIDA PROFISSIONAL</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left">• Possui outro(s) emprego(s) em empresas ou instituições de ensino?</td>
<td><img src="../images/bola.gif" width="18" height="18" border="0"></td>
<td class="campop">SIM</td>
<td><img src="../images/bola.gif" width="18" height="18" border="0"></td>
<td class="campop">NÃO</td>
</tr></table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr>
	<td class="campop" rowspan=4 valign=top width=50>Quais?</td>
	<td class=titulop width=290>Empresa</td>
	<td class=titulop width=200>Cargo ou Função</td>
	<td class=titulop width=100>Desde de:</td>
</tr>
<%
for a=1 to 3
%>
<tr>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
<%
next
%>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left">• <b>É Aposentado</b> pela Previdência Social (iniciativa privada ou serviço público)?</td>
<td><img src="<%=aposs%>" width="18" height="18" border="0"></td>
<td class="campop">SIM</td>
<td><img src="<%=aposn%>" width="18" height="18" border="0"></td>
<td class="campop">NÃO</td>
</tr></table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr>
	<td class="campop" rowspan=2 valign=middle width=50>Se não</td>
	<td class="campop" width=500 height=25 valign=middle style="border-bottom: 1px solid">
	Qual o seu tempo total de trabalho para efeito de aposentadoria?</td>
	<td class="campop" width=90 align="right" style="border-bottom: 1px solid">anos</td>
</tr>
<tr>
	<td class="campop" width=500 height=25 valign=middle style="border-bottom: 1px solid">
	Tempo restante projetado para V.Sa. se aposentar?</td>
	<td class="campop" width=90 align="right" style="border-bottom: 1px solid">anos</td>
</tr>
</table>

<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr>
	<td class="campop" colspan=2 align="left" style="border-bottom: 1px solid #000000;border-top: 1px solid #000000">
	Declaro que as informações prestadas neste documento são a expressão da verdade, responsabilizando-me por elas
	na <u>forma da lei</u>, e comprometendo-me ainda a sempre
	que quaisquer destas informações forem alteradas comunicar ao Departamento de Recursos Humanos.
	</td>
</tr>
<tr>
	<td class="campop" align="left" width=50% rowspan=2>Osasco, _______de _________________de <%=year(now)%></td>
	<td class="campop" align="left" style="border-bottom: 1px solid #000000"><br>&nbsp;</td>
</tr>
<tr>
	<td class="campop" align="left">assinatura do empregado</td>
</tr>
</table>

<br>
<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr>
	<td class="campop" align="left" style="border: 1px solid #000000;border-top: 1px solid #000000">
<b>Observações:</b>
<br> 1. Devolver esta ficha preenchida até o <b>dia <%=formatdatetime(now()+15,2)%></b>.
<br> 2. Se houve alteração em seu estado civil, favor anexar cópia da certidão (casamento, divórcio, etc.).
	
	</td>
</tr>
<tr><td class="campor" valign=top height=15 align="right">
<%
pagina=pagina+1:response.write pagina
'paginai=paginai+1:resto=paginai mod 2: if resto=0 then response.write int(paginai/2)
%>
</td></tr>
</table>
<p style="margin-top:0;margin-bottom:0;font-size:14pt"><b></p><br>

<!-- fim borda -->

</td></tr>
</table>

<%
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página -->
rs.movenext
loop

response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr>
	<td class=fundo>Setor</td><td class=fundo>Chapa</td><td class=fundo>Nome</td><td class=fundo>Controle</td>
</tr>
<%
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class="campor"><%=rs("secao")%></td><td class=campo><%=rs("chapa")%></td><td class=campo><%=rs("nome")%></td><td class=campo>&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<%
rs.movenext
loop

rs.close

end if

set rs=nothing
set rs2=nothing
set rs3=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>