<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a42")="N" or session("a42")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Altera��o Cadastral</title>
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
<p class=titulo>Ficha de Altera��o Cadastral&nbsp;<%=titulo %>
<br>
<form method="POST" action="rh_cadastral_1.asp" name="form">
<p><select size="1" name="D1">
	<option value="Todos">Todos</option>
	<option value="Falta">Reemiss�o dos que n�o entregaram</option>
	<option value="Narciso">Todos-Campus Narciso</option>
	<option value="Narciso1">Todos-Campus Narciso-Administrativo</option>
	<option value="Narciso2">Todos-Campus Narciso-Adm.Acad�mico</option>
	<option value="Narciso3">Todos-Campus Narciso-Acad�mico</option>
	<option value="Yara">Todos-Campus V.Yara</option>
	<option value="Yara1">Todos-Campus V.Yara-Administrativo</option>
	<option value="Yara2">Todos-Campus V.Yara-Adm.Acad�mico</option>
	<option value="Yara3">Todos-Campus V.Yara-Acad�mico</option>
	<option value="Wilson">Todos-Campus Jd.Wilson</option>
	<option value="EstagioY">Estagi�rios-Vila Yara</option>
	<option value="EstagioN">Estagi�rios-Narciso</option>
	<option value="EstagioW">Estagi�rios-Jd.Wilson</option>
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

<!-- impress�o do documento -->
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
<tr><td class="campop" align="left"><b>FUNDA��O INSTITUTO DE ENSINO PARA OSASCO<br></td><td class="campop" align="right"><%=now()%></tr></tr>
<tr><td class="campop" align="center" colspan=2><b><font size=3>FICHA DE ATUALIZA��O CADASTRAL<br></td></tr>
<tr><td class="campop" align="right" colspan=2><b><font size=2><%=rs("codsecao")%> - <%=rs("secao")%></td></tr>
<tr><td class="campop" align="left" colspan=2 style="border-bottom: 2 solid #000000"><b><i>DADOS PESSOAIS</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" width=70>Seu C�digo</td>
	<td class="campop" align="left">Nome Completo</td>
	<td class="campop" align="left">Apelido</td>
</tr>
<tr><td class="campop" align="left"><b>&nbsp;<%=rs("chapa")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("nome")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("apelido")%></td>
</tr>
<tr><td class="campop" align="left">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left">Estado Civil</td>
	<td class="campop" align="left">Cidade Nascimento</td>
	<td class="campop" align="left">Estado Nascimento</td>
	<td class="campop" align="left">Data Nascimento</td>
</tr><%diab=day(rs("dtnascimento")):mesb=monthname(month(rs("dtnascimento"))):anob=year(rs("dtnascimento"))%>
<tr><td class="campop" align="left"><b>&nbsp;<%=rs("estcivil")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("naturalidade")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("estadonatal")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=diab & "/" & mesb & "/" & anob%></td>
</tr>
<tr><td class=<%if isnull(rs("estadocivil")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class=<%if isnull(rs("naturalidade")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class=<%if isnull(rs("estadonatal")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
</table>

<%
sqlf="select top 5 f.codinstrucao, t.tipo, f.curso, f.instituicao, f.dataconclusao, f.localinst from uprofformacao_ f, uprof_tipo t  where t.codinstrucao=f.codinstrucao and codprof='" & rs("chapa") & "' order by t.codinstrucao desc, f.dataconclusao desc "
rs2.Open sqlf, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then linhasg=rs2.recordcount+1 else linhasg=1
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" width=130>Escolaridade</td>
	<td class="campop" align="left" rowspan=3 valign=top>
<%
if rs("grauinstrucao")>"7" then
%>
	<table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" width=100%>
	<tr>
		<td class=campo align="left">Cursos Gradua��o/Especializa��o</td>
		<td class=campo align="left">Local/Institui��o</td>
		<td class=campo align="left">Ano</td>
	</tr>
<%
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
	<tr><td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000"><%=rs2("tipo") & "-" & rs2("curso")%></td>
		<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000"><%=rs2("localinst")%></td>
		<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000"><%=year(rs2("dataconclusao"))%></td>
	</tr>
<%
rs2.movenext
loop
end if 'rs2.recordcount
%>
	<tr><td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
		<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
		<td class="campor" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	</tr>
	</table>
<%
end if 'instrucao>7
rs2.close
%>
	</td>
</tr>
<tr><td class="campop" align="left"><b>&nbsp;<%=rs("instrucao")%></td>
</tr>
<tr><td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
</table>

<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>ENDERE�O / TELEFONES</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left">Endere�o</td>
	<td class="campop" align="left">N�mero</td>
	<td class="campop" align="left">Complemento</td>
</tr>
<tr><td class="campop" align="left"><b>&nbsp;<%=rs("rua")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("numero")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("complemento")%></td>
</tr>
<tr><td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class=<%if isnull(rs("numero")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left">Bairro</td>
	<td class="campop" align="left">Cidade</td>
	<td class="campop" align="left">Estado</td>
	<td class="campop" align="left">CEP</td>
</tr>
<tr><td class="campop" align="left"><b>&nbsp;<%=rs("bairro")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("cidade")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("estado")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("cep")%></td>
</tr>
<tr><td class=<%if isnull(rs("bairro")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class=<%if isnull(rs("cep")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" width=25%>Resid�ncia</td>
	<td class="campop" align="left" width=25%>Celular</td>
	<td class="campop" align="left" width=25%>Recados/Comercial</td>
	<td class="campop" align="left" width=25%>Telefone FAX</td>
</tr>
<tr><td class="campop" align="left"><b>&nbsp;<%=rs("telefone1")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("telefone2")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("telefone3")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("fax")%></td>
</tr>
<tr><td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" width=25%>Seu E-mail de contato (n�o informe o seu email do Unifieo)</td></tr>
<%
ehfieo=instr(1,rs("email"),"unifieo",1)
 %>
<tr><td class="campop" align="left"><b>&nbsp;<%=rs("email")%></td></tr>
<tr><td class=<%if isnull(rs("email")) then response.write "fundoc" else response.write "campop"%> align="right" style="border-bottom:1px solid #000000;border-left:1px solid #000000">
<%if ehfieo>0 then response.write "<b><--------Informe um email pessoal</b>"%>
&nbsp;</td></tr>
</table>

<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>DOCUMENTOS</td></tr>
</table>
<%
colunas=0
ctps   =rs("carteiratrab") :if isnull(ctps)    then colunas=colunas+1
serie  =rs("seriecarttrab"):if isnull(serie)   then colunas=colunas+1
emissao=rs("dtcarttrab")   :if isnull(emissao) then colunas=colunas+1
estado =rs("ufcarttrab")   :if isnull(estado)  then colunas=colunas+1
%>
<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" rowspan=2>Carteira<br>Profissional</td>
	<td class="campop" align="left">N�mero</td>
	<td class="campop" align="left">S�rie</td>
	<td class="campop" align="left">Data Emiss�o</td>
	<td class="campop" align="left">Estado de Emiss�o</td>
</tr>
<tr>
<%if isnull(ctps) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("carteiratrab")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=ctps%></td>
<%if isnull(serie) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("seriecarttrab")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=serie%></td>
<%if isnull(emissao) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("dtcarttrab")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=emissao%></td>
<%if isnull(estado) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("ufcarttrab")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=estado%></td>
</tr>
</table>

<%
colunas=0
titulo =rs("tituloeleitor")  :if isnull(titulo) then colunas=colunas+1
zona   =rs("zonatiteleitor") :if isnull(zona)   then colunas=colunas+1
secao  =rs("secaotiteleitor"):if isnull(secao)  then colunas=colunas+1
%>
<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" rowspan=2>T�tulo<br>Eleitoral</td>
	<td class="campop" align="left">N�mero</td>
	<td class="campop" align="left">Zona</td>
	<td class="campop" align="left">Se��o</td>
</tr>
<tr>
<%if isnull(titulo) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("tituloeleitor")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=titulo%></td>
<%if isnull(zona) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("zonatiteleitor")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=zona%></td>
<%if isnull(secao) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("secaotiteleitor")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=secao%></td>
</tr>
</table>

<%
colunas=0
identidade=rs("cartidentidade") :if isnull(identidade)    then colunas=colunas+1
emissao   =rs("dtemissaoident") :if isnull(emissao) then colunas=colunas+1
orgao     =rs("orgemissorident"):if isnull(orgao) then colunas=colunas+1
estado    =rs("ufcartident")    :if isnull(estado)  then colunas=colunas+1
%>
<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" rowspan=2>Identidade<br>(R.G.)</td>
	<td class="campop" align="left">N�mero</td>
	<td class="campop" align="left">Data Emiss�o</td>
	<td class="campop" align="left">Org�o Emissor</td>
	<td class="campop" align="left">Estado de Emiss�o</td>
</tr>
<tr>
<%if isnull(identidade) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("cartidentidade")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=identidade%></td>
<%if isnull(emissao) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("dtemissaoident")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=emissao%></td>
<%if isnull(orgao) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("orgemissorident")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=orgao%></td>
<%if isnull(estado) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("ufcartident")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=estado%></td>
</tr>
</table>

<%
colunas=0
reservista=rs("certifreserv"):if isnull(reservista) then colunas=colunas+1
categoria =rs("categmilitar"):if isnull(categoria)  then colunas=colunas+1
cpf       =rs("cpf")         :if isnull(cpf)        then colunas=colunas+1
%>
<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr>
<%if rs("sexo")="M" then %>
	<td class="campop" align="left" rowspan=2>Reservista<br></td>
	<td class="campop" align="left">N�mero</td>
	<td class="campop" align="left">Categoria</td>
<%end if%>
	<td class="campop" align="left" rowspan=2>C.P.F.<br></td>
	<td class="campop" align="left">N�mero</td>
</tr>
<tr>
<%if rs("sexo")="M" then %>
<%if isnull(reservista) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("certifreserv")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=reservista%></td>
<%if isnull(categoria) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("categmilitar")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=categoria%></td>
<%end if%>
<%if isnull(cpf) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("cpf")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=cpf%></td>
</tr>
</table>

<%
colunas=0
motorista =rs("cartmotorista")  :if isnull(motorista)  then colunas=colunas+1
tipo      =rs("tipocarthabilit"):if isnull(tipo)       then colunas=colunas+1
vencimento=rs("dtvenchabilit")  :if isnull(vencimento) then colunas=colunas+1
%>
<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" rowspan=2 width=25%>Carteira<br>Habilita��o</td>
	<td class="campop" align="left" width=25%>N�mero</td>
	<td class="campop" align="left" width=25%>Categoria</td>
	<td class="campop" align="left" width=25%>Data Validade</td>
</tr>
<tr>
<%if isnull(motorista) then tam=2 else tam=1%>
	<td class=<%if isnull(rs("cartmotorista")) then response.write "fundoc" else response.write "campop"%> align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=motorista%></td>
<%if isnull(tipo) then tam=2 else tam=1%>
	<td class="campop" align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=tipo%></td>
<%if isnull(vencimento) then tam=2 else tam=1%>
	<td class="campop" align="left" style="border-bottom:<%=tam%> solid #000000;border-left:<%=tam%> solid #000000"><b>&nbsp;<%=vencimento%></td>
</tr>
</table>
<DIV align="center"><B>--------------> N�O SE ESQUE�A DE PREENCHER O VERSO <--------------</B></DIV>
<!-- fim borda -->
</td></tr>
</table>
<%
response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a p�gina --> 
%>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=650 height="1000">
<tr><td class="campop" align="left" valign=top height=100%>
<!-- fim borda -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left"><b>FUNDA��O INSTITUTO DE ENSINO PARA OSASCO<br></td><td class="campop" align="right">P�gina 2</tr></tr>
<tr><td class="campop" align="center" colspan=2><b><font size=3>FICHA DE ATUALIZA��O CADASTRAL<br></td></tr>
<tr><td class="campop" align="right" colspan=2><b><font size=2><%=rs("chapa")%> - <%=rs("nome")%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>FILIA��O</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr>
	<td class=titulo width=50%>Nome da M�e</td>
	<td class=titulo width=50%>Nome do Pai</td>
</tr>
<%
sqla="SELECT D.* FROM corporerm.dbo.PFDEPEND D " & _
"WHERE D.CHAPA='" & rs("chapa") & "' and grauparentesco='6' ORDER BY D.nrodepend"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then pai=rs3("nome") else pai="&nbsp;"
rs3.close
sqla="SELECT D.* FROM corporerm.dbo.PFDEPEND D " & _
"WHERE D.CHAPA='" & rs("chapa") & "' and grauparentesco='7' ORDER BY D.nrodepend"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then mae=rs3("nome") else mae="&nbsp;"
rs3.close
%>
<tr>
	<td class=campo><%=mae%></td>
	<td class=campo><%=pai%></td>
</tr>
<tr>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>DEPENDENTES LEGAIS</td></tr>
</table>
<%
sqla="SELECT D.*, MAE FROM corporerm.dbo.PFDEPEND D, corporerm.dbo.PFDEPENDCOMPL C " & _
"WHERE D.CHAPA=C.CHAPA AND D.NRODEPEND=C.NRODEPEND AND D.CHAPA='" & rs("chapa") & "' and grauparentesco not in ('6','7') ORDER BY D.nrodepend"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr>
	<td class=titulo width=200>Nome</td>
	<td class=titulo>Parentesco</td>
	<td class=titulo>Data Nasc.</td>
	<td class=titulo>Sexo</td>
	<td class=titulo>IRRF</td>
	<td class=titulo width=200>Nome da M�e do dependente</td>
</tr>
<%
if rs3.recordcount>0 then
rs3.movefirst:do while not rs3.eof
sql="select descricao from corporerm.dbo.pcodparent where codcliente='" & rs3("grauparentesco") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then parentesco=trim(rs2("descricao")) else parentesco=""
if parentesco="Filho(a) Valido" and rs3("sexo")="M" then parentesco="Filho"
if parentesco="Filho(a) Valido" and rs3("sexo")="F" then parentesco="Filha"
rs2.close
%>
<tr>
	<td class=campo align="left"><%=rs3("nome")%></td>
	<td class=campo align="left"><%=parentesco%></td>
	<td class=campo align="center"><%=rs3("dtnascimento")%></td>
	<td class=campo align="center"><%=rs3("sexo")%></td>
	<td class=campo align="center"><%if rs3("incirrf")=1 then response.write "<font face='Wingdings'>�</font>" %></td>
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
</table>
<p style="margin-top:0;margin-bottom:0;font-size:9pt;text-align:right"><b>S� � considerado dependente para Imposto de Renda, aqueles com <font face='Wingdings'>�</font> no IRRF.
<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>VE�CULOS</td></tr>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b>A atualiza��o do cadastro de ve�culos � essencial para
a emiss�o do novo crach� de estacionamento.</td></tr>
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
	<td class=titulop>T�rmino</td>
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
	<td class="campop"><%=rs3("dttermino")%></td>
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
<tr><td class="campop" align="left">� Possui outro(s) emprego(s) em empresas ou institui��es de ensino?</td>
<td><img src="../images/bola.gif" width="18" height="18" border="0"></td>
<td class="campop">SIM</td>
<td><img src="../images/bola.gif" width="18" height="18" border="0"></td>
<td class="campop">N�O</td>
</tr></table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr>
	<td class="campop" rowspan=4 valign=top width=50>Quais?</td>
	<td class=titulop width=290>Empresa</td>
	<td class=titulop width=200>Cargo ou Fun��o</td>
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
<tr><td class="campop" align="left">� <b>� Aposentado</b> pela Previd�ncia Social (iniciativa privada ou servi�o p�blico)?</td>
<td><img src="<%=aposs%>" width="18" height="18" border="0"></td>
<td class="campop">SIM</td>
<td><img src="<%=aposn%>" width="18" height="18" border="0"></td>
<td class="campop">N�O</td>
</tr></table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr>
	<td class="campop" rowspan=2 valign=middle width=50>Se n�o</td>
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
	Declaro que as informa��es prestadas neste documento s�o a express�o da verdade, responsabilizando-me por elas
	na <u>forma da lei</u>, e comprometendo-me ainda a sempre
	que quaisquer destas informa��es forem alteradas comunicar ao Departamento de Recursos Humanos.
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
<b>Observa��es:</b>
<br> 1. Devolver esta ficha preenchida at� o <b>dia <%=formatdatetime(now()+15,2)%></b>.
<br> 2. Se houve altera��o em seu estado civil, favor anexar c�pia da certid�o (casamento, div�rcio, etc.).
	
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
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a p�gina -->
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