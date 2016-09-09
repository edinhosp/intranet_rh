<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a63")="N" or session("a63")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cadastro de Autônomos</title>
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
dim conexao, conexao2
dim rs, rs2
set conexao=server.createobject ("ADODB.Connection")
'conexao.Open Application("consql")

registros=50 'Session("RegistrosPorPagina")

sqla="select p.id_lanc, p.id_autonomo, a.nome_autonomo, p.data_emissao, p.data_pagamento, p.descricao_servico, p.valor_liquido " & _
"from autonomo_rpa p, autonomo a " & _
"where a.id_autonomo=p.id_autonomo and p.data_pagamento is null "
sqlb=""
sqlc="ORDER BY a.nome_autonomo "

sql1=sqla & sqlb &  sqlc

'if Session("PrimeiraVez")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	conexao.open Application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	Session("Pagina")=1
	MostraDados
	Session("PrimeiraVez")="Nao"
'else
'	if request("folha")="" then pagina=1
'	if request.form("pagina")<>"" then pagina=request.form("pagina")
'	if request("folha")<>"" then pagina=request("folha")
'	Session("Pagina")=pagina
'	conexao.cursorlocation = 3 'aduseclient
'	conexao.open Application("conexao")
'	set rs=server.createobject ("ADODB.Recordset")
'	rs.CacheSize = registros
'	rs.PageSize = registros
'	set rs.ActiveConnection = conexao
'	rs.Open sql1, ,adOpenStatic, adLockReadOnly
'	if rs.recordcount>0 then 	MostraDados
'end if	

Sub MostraDados()
	if rs.recordcount>0 then 	rs.AbsolutePage=Session("Pagina") 'vai para o número da pagina armazenado na sessão
End Sub
%>
<form method="POST" name="form" action="autonomo_pendente.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Cadastro de Autônomos</p>
<table border="0" width="650" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="70%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
	response.write "<img src='../images/setafirst0.gif' border='0'>"
	response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
	response.write "<a href=""autonomo_pendente.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
	response.write "<a href=""autonomo_pendente.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
end if

response.write "&nbsp;<b>"
response.write "<select size='1' name='pagina' onChange='javascript:submit()'>"
for selpag=1 to rs.pagecount
	if selpag=atual then selpag1="selected" else selpag1=""
	response.write "<option value=" & selpag & " " & selpag1 & ">" & selpag & "</option>"
next
response.write "</select>"
response.write "/" & rs.pagecount & "</b>&nbsp;"

if atual=rs.pagecount or rs.pagecount=0 then
	response.write "<img src='../images/setanext0.gif' border='0'>"
	response.write "<img src='../images/setalast0.gif' border='0'>"
else
	response.write "<a href=""autonomo_pendente.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
	response.write "<a href=""autonomo_pendente.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
	<td class=campo width="30%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="690" cellpadding="1" style="border-collapse: collapse">
<tr>
	<td class=fundor align="center">Contr.         </td>
	<td class=titulor align="center">Nome autônomo </td>
	<td class=titulor align="center">Tipo prestação</td>
	<td class=titulor align="center">Data Emissão  </td>
	<td class=titulor align="center">V.Líquido     </td>
</tr>
<%
linha=1
'rs.movefirst:do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
%>
<tr>
	<td class="campor" align="center">
	<% if session("a63")="T" then %>
		<a href="rpa_alteracao.asp?codigo=<%=rs("id_lanc")%>" onclick="NewWindow(this.href,'AlteracaoRPA','510','350','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/folder95o.gif" alt="Alterar este RPA"></a>
	<% end if %>
	</td>
	<td class="campor"><%=rs("nome_autonomo")%></td>
	<td class="campor"><%=rs("descricao_servico")%></td>
	<td class="campor" align="center"><%=rs("data_emissao")%></td>
	<td class="campor" align="right"><%=formatnumber(rs("valor_liquido"),2) %></td>
</tr>
<%
if linha=1 then linha=0 else linha=1
rs.movenext
if rs.eof then exit for
'loop
Next

else 'sem registros
%>
<td class=grupo colspan=11>Esta seleção não mostra nenhum registro.</td>
<%
end if
%>
</table>
<!--
<input name="B2" type="submit" class="button" value="Clique para Filtrar">
-->
</form>
</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>