<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a27")="N" or session("a27")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Rescisão de Nomeações</title>
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
	
sqlc="SELECT top 2 i.id_nomeacao, n.NOMEACAO, i.id_indicado, i.CHAPA, i.NOME, i.PORTARIA, " & _
"i.codeve, i.CARGO, i.MAND_INI, i.MAND_FIM, i.CH, i.obs, s.codigo, s.descricao " & _
"FROM n_indicacoes as i, n_nomeacoes as n, qry_nomeacoes_setor s " & _
"WHERE i.id_nomeacao = n.id_nomeacao and (i.mand_fim>#12/1/2003# or isnull(i.mand_fim)) " & _
"and i.chapa=s.chapa " 
sqle="order by nome "
sqle=""
sqld=""
sqlc="SELECT top 5 i.CHAPA, i.NOME, f.RUA, f.NUMERO, f.COMPLEMENTO, f.BAIRRO, f.ESTADO, f.CIDADE, f.CEP " & _
"FROM (n_indicacoes AS i INNER JOIN n_nomeacoes AS n ON i.id_nomeacao = n.id_nomeacao) INNER JOIN qry_funcionarios f ON i.CHAPA=f.CHAPA collate database_default " & _
"WHERE ((i.id_nomeacao Not In (20,19,4,1,21)) AND (i.MAND_FIM>'" & dtaccess(now) & "' Or i.MAND_FIM Is Null) AND (f.CODSITUACAO<>'D')) " & _
"GROUP BY i.CHAPA, i.NOME, f.RUA, f.NUMERO, f.COMPLEMENTO, f.BAIRRO, f.ESTADO, f.CIDADE, f.CEP "  & _
"order by i.nome, cep "

sqlb=sqlc & sqld & sqle
rs.Open sqlb, ,adOpenStatic, adLockReadOnly
temp=0
titulo=rs("chapa") & " - " & rs("nome")
'	session("nomeacao_chapa")=rs("chapa")
'	session("nomeacao_id")=""
'	session("nomeacao_descr")=""

rs.movefirst
do while not rs.eof 
%>
	<p style="font-size: 12pt; text-align:justify">
	<br><b>Of. Reitoria <input type="text" class="form_input10" value="001/04"></b>
	<br>
	<br>
	<p style="font-size: 12pt; text-align:right">
	<br>Osasco, <input type="text" class="form_input10" value="05 de janeiro de 2004."><br>
	<br>&nbsp;
	<br>
	<br>
	<br>
	<p style="margin-top:0; margin-bottom:0;font-size:10pt"><b>
<%=rs("nome")%><br>
<%=rs("rua") & " " & rs("numero") & " " & rs("complemento")%><br>
<%=trim(rs("bairro"))%><br>
<%=rs("cidade") & " - " & rs("estado")%><br>
<%=rs("cep")%></b>
	<br>
	<br>
	<p style="font-size: 12pt; text-align:justify">
	<br><b>Prezado Professor</b>
	<br>&nbsp;
	<br>&nbsp;
	<br>
<p style="margin-left:-30px">_</p>
 	<br>
	<br>
<p style="font-size: 12pt; text-align:justify">
	<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Encaminhamos a V.Sa. a Portaria nº 05/2004, desta Reitoria, que suspende o exercício dos cargos em comissão.<br>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Esperamos continuar contando com a sua colaboração em designação futura.
	<br>
	<br>
	<br>
	<p style="font-size: 12pt; text-align:center">
	<br>Atenciosamente,
	<br>
	<br>
	<br>
	<br><b>A Reitoria</b>
	<p style="font-size: 12pt; text-align:justify">
	
<%
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
rs.movenext
loop
rs.close

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>