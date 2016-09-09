<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a94")="N" or session("a94")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Manutenção Cadastro de Uniformes</title>
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
'set conexao2=server.createobject ("ADODB.Connection")
'conexao2.Open Application("consql")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
sessao=session.sessionid
%>
<p class=titulo>Checagem do Controle de Uniformes
<%
'***** atualizacao de novos funcionários  ******
sql2="SELECT chapa, nome FROM corporerm.dbo.pfunc WHERE codsituacao<>'D' and codsindicato<>'03' and chapa<'10000' and codtipo='N' " & _
"union all " & _
"select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' and chapa>'90000' and codtipo='T' "
'response.write sql2
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
response.write "<p>Funcionários ativos: " & rs2.recordcount
total=0
do while not rs2.eof
	sql1="select chapa from uniforme_func_cat where chapa='" & rs2("chapa") & "'"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
		'nada a fazer
	else
		sql3="insert into uniforme_func_cat (chapa, inicio, id_cat ) " & _
		"select '" & rs2("chapa") & "', getdate(), 0 "
		conexao.execute sql3
		total=total+1
		response.write "<br>Inseriu " & rs2("chapa") & " - " & rs2("nome")
	end if
	rs.close
rs2.movenext
loop
rs2.close
response.write "<br>Novos funcionários: " & total

%>
<p class=titulo>Uniformes sem Categoria
<table border="1" cellpadding="1" width="600" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Cod.</td>
	<td class=titulo>Descrição</td>
	<td class=titulo>Tamanho</td>
	<td class=titulo>&nbsp;</td>
</tr>
<%
sql3="SELECT i.id_item, i.descricao, i.tamanho, l.id_item AS checar " & _
"FROM uniforme_item i LEFT JOIN uniforme_link l ON i.id_item=l.id_item " & _
"WHERE l.id_item is null"
rs.Open sql3, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo><%=rs("id_item")%></td>
	<td class=campo><%=rs("descricao")%></td>
	<td class=campo><%=rs("tamanho")%></td>
	<td class=campo align="center">
    <% if session("a94")="T" then %>
		<a href="itens_alteracao.asp?codigo=<%=rs("id_item")%>" onclick="NewWindow(this.href,'AlteracaoUniforme','530','330','no','center');return false" onfocus="this.blur()">
		<img border='0' src='../images/folder95o.gif'></a>
	<% end if %>
	</td>
</tr>
<%
rs.movenext
loop
else
	response.write "<td class=grupo colspan='4'>&nbsp;Sem uniformes sem categorização pendentes.</td>"
end if
rs.close
%>
</table>
<DIV style="page-break-after:always"></DIV> <!-- Aqui quebra a página --> 

<p class=titulo>Funcionários sem Categoria de Uniforme
<table border="1" cellpadding="1" width="600" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Setor</td>
	<td class=titulo>Categoria</td>
</tr>
<%
sql3="select u.id_fcat, u.chapa, u.id_cat, f.nome, s.descricao as setor, f.codsituacao " & _
"from uniforme_func_cat u, corporerm.dbo.pfunc f, corporerm.dbo.psecao s where u.chapa=f.chapa collate database_default and " & _
"f.codsecao=s.codigo and (u.id_cat=0) and codtipo in ('N','T') order by u.chapa "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
if rs("codsituacao")="D" then cor="red" else cor="black"
%>
<tr>
	<td class=campo>
	<% if session("a94")="T" then %>
		<a href="func_cat_alteracao.asp?codigo=<%=rs("id_fcat")%>" onclick="NewWindow(this.href,'Alteracao','440','170','yes','center');return false" onfocus="this.blur()">
		<font size="1"><%=rs("chapa")%></font></a>
	<% else %>
		<%=rs("chapa")%>
	<% end if %>
	</td>
	<td class=campo><font color=<%=cor%>><%=rs("nome")%></td>
	<td class=campo><%=rs("setor")%></td>
	<td class=campo></td>
</tr>
<%
rs.movenext
loop
else
	response.write "<td class=grupo colspan='4'>&nbsp;Sem cadastros de categorias pendentes.</td>"
end if
rs.close
%>
</table>
<DIV style="page-break-after:always"></DIV> <!-- Aqui quebra a página --> 
<p class=titulo>Funcionários sem numeração de uniforme
<table border="1" cellpadding="1" width="650" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulor>Categoria</td>
</tr>
<%
sql3="SELECT fc.chapa, f.NOME, c.id_cat, c.descricao, min(f.codsituacao) as codsituacao " & _
"FROM (uniforme_func_item AS fi RIGHT JOIN (uniforme_func_cat AS fc INNER JOIN corporerm.dbo.PFUNC AS f ON fc.chapa=f.CHAPA collate database_default) ON fi.id_fcat = fc.id_fcat) LEFT JOIN uniforme_categoria AS c ON fc.id_cat = c.id_cat " & _
"where codsituacao<>'D' " & _
"GROUP BY fc.chapa, f.NOME, c.id_cat, c.descricao " & _
"HAVING c.id_cat<>8 AND Count(fi.id_fitem)=0 "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
if rs("codsituacao")="D" then cor="red" else cor="black"
%>
<tr>
	<td class=campo>
    <% if session("a94")="T" then %>
	<a href="funcionario_ver.asp?codigo=<%=rs("chapa")%>" onclick="NewWindow(this.href,'Alteracao','690','500','yes','center');return false" onfocus="this.blur()">
	<font size="1"><%=rs("chapa")%></font></a>
	<% else %>
		<%=rs("chapa")%>
	<% end if %>
	</td>
	<td class=campo><font color=<%=cor%>><%=rs("nome")%></td>
	<td class=campo><%=rs("descricao")%></td>
</tr>
<%
rs.movenext
loop
else
	response.write "<td class=grupo colspan='13'>&nbsp;Sem cadastros de numeração pendentes.</td>"
end if 'rs.recordcount
rs.close
%>
</table>
<%

conexao.close
set conexao=nothing
set rs=nothing
set rs2=nothing
%>
</body>
</html>