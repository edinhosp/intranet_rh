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
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
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
		sqld=" and f.chapa not in (select chapa from zselecao where sessao='0') and f.admissao<#12/31/2005# "
	else
		sqld=" and f.chapa='" & idchapa & "'"
	end if
	
	sqlc="SELECT top 500 * from qry_funcionarios f " & _
	"where f.codsituacao in ('A','F','E','P','Z','L') and f.chapa>'00000' "
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
<p class=titulo>Ficha de Verificação Cadastral&nbsp;<%=titulo %>
<br>
<form method="POST" action="rh_cadastral2.asp" name="form">
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
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=650 height="500">
<tr><td class="campop" align="left" valign=top>
<!-- fim borda -->

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border: 1px solid #000000;border-top: 1px solid #000000">
<b>Observações:</b>
<br> 1. Devolver esta ficha preenchida até o <b>dia <%=formatdatetime(now()+15,2)%></b>.
<br> 2. Se houve alteração em seu estado civil, favor anexar cópia da certidão (casamento, divórcio, etc.).
	</td></tr>
</table>
<br>


<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td><td class="campop" align="right"><%=rs("codsecao")%> - <%=rs("secao")%></tr></tr>
<tr><td class="campop" align="center" colspan=2><b><font size=3>FICHA DE ATUALIZAÇÃO CADASTRAL</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse:collapse;border-bottom:2 solid #000000;border-top:2 solid #000000" width=640>
<tr><td class="campop" align="left" width=70><b>&nbsp;<%=rs("chapa")%></td>
	<td class="campop" align="left"><b>&nbsp;<%=rs("nome")%></td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class=campo align="left" style="border-bottom: 2 dotted #000000"><b><i>ENDEREÇO</td>
	<td class=campo align="left" style="border-bottom: 2 dotted #000000"><b>Escolaridade</td>
</tr>
<%
if rs("complemento")<>"" then complemento=" - " & rs("complemento") else complemento=""
if rs("bairro")<>"" then bairro=" - " & rs("bairro") else bairro=""
if rs("email")<>"" then email=rs("email") else email="----"
%>
<tr><td class="campop" align="left">&nbsp;<%=rs("rua") & ", " & rs("numero") & complemento & bairro %></td>
	<td class="campop" align="left"><%=rs("instrucao")%></td></tr>
<tr><td class="campop" align="left">&nbsp;<%=rs("cidade") & " - CEP " & rs("cep") %></td>
	<td class="campop" align="left" style="border-bottom:1 dashed #000000">&nbsp;</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class=campo align="left" style="border-bottom: 2 dotted #000000"><b><i>TELEFONES</td>
	<td class=campo align="left" style="border-bottom: 2 dotted #000000"><b><i>EMAIL</td></tr>
<tr><td class="campop" align="left">&nbsp;<font style="font-family:Courier New">Res.      :</font> <%=rs("telefone1")%></td>
	<td class="campop" align="left" valign=top>&nbsp;<%=email%></td></tr>
<tr><td class="campop" align="left">&nbsp;<font style="font-family:Courier New">Cel.      :</font> <%=rs("telefone2")%></td>
	<td class=campo align="left" style="border-bottom: 2 dotted #000000"><b><i>ESTADO CIVIL</td></tr>
<tr><td class="campop" align="left">&nbsp;<font style="font-family:Courier New"><b>Recados   :</b></font> <%if rs("telefone3")="" or isnull(rs("telefone3")) then response.write string(12,"_") else response.write rs("telefone3")%> 
	<%if rs("fax")="" or isnull(rs("fax")) then contato=string(10,"_") else contato=rs("fax")%> (contato c/:<%=contato%>)
	<!--<br>&nbsp;<font style="font-family:Courier New">Fax       :</font> <%=rs("fax")%>--></td>
	<td class="campop" align="left">&nbsp;<%=rs("estcivil")%></td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="center"><b>Caso alguma das informações esteja incorreta ou incompleta, queira atualizá-la abaixo.</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>ATUALIZAÇÃO DE ENDEREÇO</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="left" width=60%>Endereço</td>
	<td class=campo align="left" width=10%>Número</td>
	<td class=campo align="left" width=30%>Complemento</td>
</tr>
<tr><td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="left" width=40%>Bairro</td>
	<td class=campo align="left" width=30%>Cidade</td>
	<td class=campo align="left" width=10%>Estado</td>
	<td class=campo align="left" width=20%>CEP</td>
</tr>
<tr><td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>ATUALIZAÇÃO DE TELEFONE/EMAIL</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="left" width=15%>Residência</td>
	<td class=campo align="left" width=15%>Celular</td>
	<td class=campo align="left" width=15%>Recados(Contato)</td>
	<td class=campo align="left" width=15%>Telefone FAX</td>
	<td class=campo align="left" width=40%>Email</td>
</tr>
<tr><td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border:1px solid #000000;">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
</table>

<!-- iconta -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>ATUALIZAÇÃO DE DADOS BANCÁRIOS</td></tr>
</table>
<%
if rs("banco")="237" and len(rs("agencia"))=6 then
	agencia=clng(rs("agencia"))
	agencia=left(agencia,len(agencia)-1)&"-"&right(agencia,1)
else
	agencia=rs("agencia")
end if
if rs("banco")="237" and len(rs("conta"))=8 then
	if right(rs("conta"),1)="P" then
		conta=rs("conta")
	else
		conta=clng(rs("conta"))
	end if
	conta=left(conta,len(conta)-1)&"-"&right(conta,1)
else
	conta=rs("conta")
end if
if rs("banco")="237" and rs("razao")<>"" then
	select case rs("razao")
		case "07.05"
			razao=rs("razao")&"-C.Corrente"
		case "07.38"
			razao=rs("razao")&"-C.Salário"
		case "10.51"
			razao=rs("razao")&"-C.Poupança"
		case else
			razao=rs("razao")&"-???"
	end select
else
	razao=rs("razao")
end if

%>
<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class=campo align="left" width=10%>Banco</td>
	<td class=campo align="left" width=12%>Agência</td>
	<td class=campo align="left" width=18%>C.Corrente</td>
	<td class=campo align="left" width=22%>Tipo</td>
	<td class="campor" align="left" width=38%>(Confirme o nº escrevendo "Correta" no espaço abaixo)</td>
</tr>
<tr><td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">
	&nbsp;<%=rs("banco")%></td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">
	&nbsp;<%=agencia%></td>
	<td class="campop" align="left" style="border:1px solid #000000;">&nbsp;<%=conta%></td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;<%=razao%></td>
	<td class="campop" align="left" style="border-bottom:1px solid #000000;border-left:1px solid #000000">&nbsp;</td>
</tr>
</table>
<!-- fconta -->

<%
sql="SELECT F.CHAPA, F.NOME, 'APELIDO' AS CAMPO, P.APELIDO AS CONTEUDO FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.APELIDO Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'ESTADO CIVIL' AS CAMPO, P.ESTADOCIVIL FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.ESTADOCIVIL Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'CIDADE DE NASCIMENTO' AS CAMPO, P.NATURALIDADE FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.NATURALIDADE Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'ESTADO DE NASCIMENTO' AS CAMPO, P.ESTADONATAL FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.ESTADONATAL Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'DATA DE NASCIMENTO' AS CAMPO, P.DTNASCIMENTO FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.DTNASCIMENTO Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'CTPS: NUMERO' AS CAMPO, P.CARTEIRATRAB FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.CARTEIRATRAB Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'CTPS: SERIE' AS CAMPO, P.SERIECARTTRAB FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.SERIECARTTRAB Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'CTPS: UF' AS CAMPO, P.UFCARTTRAB FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.UFCARTTRAB Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'CTPS: DATA DE EMISSÃO' AS CAMPO, P.DTCARTTRAB FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.DTCARTTRAB Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'TITULO ELEITORAL' AS CAMPO, P.TITULOELEITOR FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.TITULOELEITOR Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'TIT.ELEITORAL: ZONA' AS CAMPO, P.ZONATITELEITOR FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.ZONATITELEITOR Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'TIT.ELEITORAL: SEÇÃO' AS CAMPO, P.SECAOTITELEITOR FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.SECAOTITELEITOR Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'IDENTIDADE: NUMERO' AS CAMPO, P.CARTIDENTIDADE FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.CARTIDENTIDADE Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'IDENTIDADE: UF' AS CAMPO, P.UFCARTIDENT FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.UFCARTIDENT Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'IDENTIDADE: ORGÃO EMISSOR' AS CAMPO, P.ORGEMISSORIDENT FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.ORGEMISSORIDENT Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'IDENTIDADE: DATA EMISSÃO' AS CAMPO, P.DTEMISSAOIDENT FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.DTEMISSAOIDENT Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'RESERVISTA/DISPENSA' AS CAMPO, P.CERTIFRESERV FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.CERTIFRESERV Is Null AND F.CODSITUACAO<>'D' AND P.SEXO='M' AND P.NACIONALIDADE='10' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'CATEGORIA MILITAR' AS CAMPO, P.CATEGMILITAR FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.CATEGMILITAR Is Null AND F.CODSITUACAO<>'D' AND P.SEXO='M' AND P.NACIONALIDADE='10' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'C.P.F.' AS CAMPO, P.CPF FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.CPF Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'CART.HABILITAÇÃO' AS CAMPO, P.CARTMOTORISTA FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.CARTMOTORISTA Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'CART.HABILITAÇÃO: CATEGORIA' AS CAMPO, P.TIPOCARTHABILIT FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.TIPOCARTHABILIT Is Null AND F.CODSITUACAO<>'D' AND P.CARTMOTORISTA Is Not Null and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'CART.HABILITAÇÃO: VENCIMENTO' AS CAMPO, P.DTVENCHABILIT FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE (P.DTVENCHABILIT Is Null Or P.DTVENCHABILIT<getdate()) AND F.CODSITUACAO<>'D' AND P.CARTMOTORISTA Is Not Null and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'CURSO DE GRADUAÇÃO' AS CAMPO, null AS GRADUACAO FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE F.CODSITUACAO<>'D' AND P.GRAUINSTRUCAO>'8' " & _
"AND F.CHAPA collate database_default NOT IN (SELECT CODPROF FROM UPROFFORMACAO_ WHERE CODINSTRUCAO='9' GROUP BY CODPROF) and f.chapa collate database_default='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'EMAIL' AS CAMPO, P.EMAIL FROM corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO " & _
"WHERE P.EMAIL Is Null AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'NOME DA MÃE' AS CAMPO, D.NOME FROM corporerm.dbo.PFDEPEND D RIGHT JOIN (corporerm.dbo.PFUNC F INNER JOIN corporerm.dbo.PPESSOA P ON F.CODPESSOA=P.CODIGO) ON D.CHAPA=F.CHAPA " & _
"WHERE (D.NOME Is Null Or D.NOME='N/D') AND D.GRAUPARENTESCO='7' AND F.CODSITUACAO<>'D' and f.chapa='" & rs("chapa") & "' " & _
"UNION ALL " & _
"SELECT F.CHAPA, F.NOME, 'REG.PROFISSIONAL' AS CAMPO, P.REGPROFISSIONAL FROM corporerm.dbo.PFUNC AS F INNER JOIN corporerm.dbo.PPESSOA AS P ON F.CODPESSOA = P.CODIGO " & _
"WHERE P.REGPROFISSIONAL Is Null AND F.CODSITUACAO<>'D' AND P.GRAUINSTRUCAO>='9' and f.chapa='" & rs("chapa") & "' " 
'response.write sql
rs2.Open sql, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
%>
<table border="0" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class=titulop align="center" style="border: 1px solid #000000"><b>VERIFICAÇÃO DE CADASTRO</td></tr>
<tr><td class="campop" align="left">As informações abaixo constam sem preenchimento ou desatualizadas. Por gentileza,
atualize-as.</td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<%
rs2.movefirst:do while not rs2.eof
'******************************
'response.write "<table border='1' cellpadding='0' cellspacing='1' style='border-collapse:collapse' width='650'>"
'response.write "<tr>"
'for a= 0 to rs2.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs2.fields(a).name) & "</td>"
'next
'response.write "</tr>"
'do while not rs2.eof 
'response.write "<tr>"
'for a= 0 to rs2.fields.count-1
'	response.write "<td class="campor">" & rs2.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs2.movenext
'loop
'response.write "</table>"
'rs2.close
'response.write "<p>"
'*****************************
%>
<tr>
	<td width=50% class="campop" align="left" style="border-bottom:1 dashed #000000"><b><%=rs2("campo")%></td>
	<td width=25% class="campop" align="left" style="border-bottom: 1px solid #000000"><%=rs2("conteudo")%></td>
	<td width=25% class="campop" align="left" style="border-bottom:1 dotted #000000">&nbsp;</td>
</tr>
<%
rs2.movenext
loop
%>
</table>
<%
end if ' campos vazios
rs2.close
%>
<br>
<%
sql="select aposentado, dtaposentadoria from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "' "
rs2.Open sql, ,adOpenStatic, adLockReadOnly
if rs2("aposentado")=1 then 
	aposentado=1:imagem1="../images/bolax.gif":imagem2="../images/bola.gif"
else
	aposentado=0:imagem1="../images/bola.gif":imagem2="../images/bolax.gif"
end if
dtaposentadoria=rs2("dtaposentadoria")
rs2.close
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left" style="border-bottom: 2 solid #000000"><b><i>VIDA PROFISSIONAL</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left">• <b>É Aposentado</b> pela Previdência Social (iniciativa privada ou serviço público)?</td>
<td><img src="<%=imagem1%>" width="18" height="18" border="0"></td><td class="campop">SIM</td>
<td><img src="<%=imagem2%>" width="18" height="18" border="0"></td><td class="campop">NÃO</td>
</tr></table>

<%
if aposentado=0 then
sql="select data_tempo, tempo_trabalho, tempo_restante from pfunc_compl where chapa='" & rs("chapa") & "' "
rs2.Open sql, ,adOpenStatic, adLockReadOnly
anos=0:anostrab="":anosrest=""
if rs2.recordcount>0 then
	if rs2("data_tempo")<>"" then anos=int((now()-rs2("data_tempo"))/365.25)
	if rs2("tempo_trabalho")>0 then anostrab=rs2("tempo_trabalho")+anos
	if rs2("tempo_restante")>0 then anosrest=rs2("tempo_restante")-anos
end if
rs2.close
%>
<table border="0" cellpadding="1" cellspacing="5" style="border-collapse: collapse" width=640>
<tr><td class="campop" rowspan=2 valign=middle width=50>Se não</td>
	<td class="campop" width=500 height=25 valign=middle style="border-bottom: 1px solid">
	Qual o seu tempo total de trabalho para efeito de aposentadoria?</td>
	<td class="campop" width=90 align="right" style="border-bottom: 1px solid"><%=anostrab%> anos</td>
</tr>
<tr><td class="campop" width=500 height=25 valign=middle style="border-bottom: 1px solid">
	Tempo restante projetado para V.Sa. se aposentar?</td>
	<td class="campop" width=90 align="right" style="border-bottom: 1px solid"><%=anosrest%> anos</td>
</tr>
</table>
<%end if%>

<%
sql="select outra_empresa from pfunc_compl where chapa='" & rs("chapa") & "' "
rs2.Open sql, ,adOpenStatic, adLockReadOnly
outra_empresa=0:imagem1="../images/bola.gif":imagem2="../images/bolax.gif"
if rs2.recordcount>0 then
	if rs2("outra_empresa")=-1 then outra_empresa=1:imagem1="../images/bolax.gif":imagem2="../images/bola.gif"
end if 'rs2.recordcount
rs2.close
%>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campop" align="left">• Possui outro(s) emprego(s) em empresas ou instituições de ensino?</td>
<td><img src="<%=imagem1%>" width="18" height="18" border="0"></td><td class="campop">SIM</td>
<td><img src="<%=imagem2%>" width="18" height="18" border="0"></td><td class="campop">NÃO</td>
</tr></table>

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

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr><td class="campor" valign=top height=15 align="right">
<%pagina=pagina+1:'response.write pagina
response.write rs.absoluteposition
'paginai=paginai+1:resto=paginai mod 2: if resto=0 then response.write int(paginai/2)%>
</td></tr>
</table>
<br>

<!-- fim borda -->
</td></tr>
</table>

<%
'**************** fim - alterações para termo de responsabilidade internet 
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página -->
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