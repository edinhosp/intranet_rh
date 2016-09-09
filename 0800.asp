<%@ Language=VBScript %>
<!-- #Include file="ADOVBS.INC" -->
<!-- #Include file="funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a37")="N" or session("a37")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Funcionários - UNIFIEO</title>
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
<link rel="stylesheet" type="text/css" href="diversos.css">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
registros=Session("RegistrosPorPagina")
dim conexao, conexao2, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rs2=server.createobject ("ADODB.Recordset")
set rs2.ActiveConnection = conexao

if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"
	
	if request.form("localizar")="" then
		session("loc37")="Digite o nome ou parte dele"
	else
		session("loc37")=request.form("localizar")
	end if

	if isnumeric(session("loc37")) then
		session("sql37d")="where (f.chapa like '%" & session("loc37") & "%') "
	else
		session("sql37d")="where (f.nome like '%" & session("loc37") & "%') "
	end if

	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
if session("sql37d")="" then session("sql37d")="where chapa='00000' "

sqla="SELECT chapa, f.nome, codsituacao, situacao=t.descricao, secao=s.descricao, bloco, funcao=c.nome, f.codhorario, horario=h.descricao, campus=case left(f.codsecao,2) when '01' then 'Narciso' when '03' then 'V.Yara' when '04' then 'Jd.Wilson' end, p.email, p.telefone1, p.telefone2, fax, codsindicato " & _
"FROM corporerm.dbo.pfunc f inner join corporerm.dbo.ppessoa p on p.codigo=f.codpessoa " & _
"inner join corporerm.dbo.psecao s on s.codigo=f.codsecao inner join corporerm.dbo.pfuncao c on c.codigo=f.codfuncao " & _
"inner join corporerm.dbo.ahorario h on h.codigo=f.codhorario inner join corporerm.dbo.pcodsituacao t on t.codcliente=f.codsituacao " & _
"left join blocos b on b.codsecao=f.codsecao collate database_default "
sqlb=" and (chapa<'10000' or chapa>'90000') "
sqlc="order by case codsituacao when 'D' then 3 when 'A' then 1 when 'E' then 1 when 'F' then 1 when 'Z' then 1 else 2 end, f.nome "

sql1=sqla & session("sql37d") & sqlb & sqlc

if Session("PrimeiraVez")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	'conexao.open Application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	Session("Pagina")=1
	MostraDados
	Session("PrimeiraVez")="Nao"
else
	if request("folha")="" then pagina=1
	if request.form("pagina")<>"" then pagina=request.form("pagina")
	if request("folha")<>"" then pagina=request("folha")
	Session("Pagina")=pagina
	conexao.cursorlocation = 3 'aduseclient
	'conexao.open Application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 	MostraDados
end if	

Sub MostraDados()
	if rs.recordcount>0 then 	rs.AbsolutePage=Session("Pagina") 'vai para o número da pagina armazenado na sessão
End Sub
%>
<form method="POST" action="0800.asp" name="form">
<input type="hidden" name="vez1" value="<%=session("PrimeiraVez")%>">
<p class=titulo style="margin-top: 2; margin-bottom: 2;border-bottom:1px solid">Localizar Funcionários - UNIFIEO</p>
<p style="margin-top: 0; margin-bottom: 5;"><font color="blue">
<b>Localizar por nome: <input type="text" name="localizar" size=35 value="<%=session("loc37")%>">
Nomes por Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
</font>
<input name="B2" type="submit" class="button" value="Clique para Filtrar"></p>

<table border="0" width="690" cellspacing="0" style="border-collapse: collapse" cellpadding="0">
<tr>
    <td class=campo width="60%" valign="center" align="left">Página: 
<%
Session("Load1")="1"
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""0800.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""0800.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
end if

response.write "&nbsp;<b>"
response.write "<select size='1' name='pagina' onchange='javascript:submit()'>"
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
response.write "<a href=""0800.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""0800.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
    <td class=campo width="20%" valign="center" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
	<td class=campo width="20%" valign="top" align="right">
	</td>
  </tr>
<tr><td class="campop" height=5></td></tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="690" cellpadding="1" style="border-collapse: collapse">
<tr>
	<td class="campop" colspan=6>Cores de situação do funcionário:&nbsp;&nbsp;&nbsp;
	<font color=black>Ativo</font> &nbsp;&nbsp;&nbsp;
	<font color=red>Demitido</font> &nbsp;&nbsp;&nbsp;
	<font color=green>Afastado</font> 
	</td>
</tr>
<tr>
    <td class=titulop align="center">Nome</td>
    <td class=titulop align="center">Setor / Campus</td>
    <td class=titulop align="center">Função / Horario</td>
    <td class=titulop align="center">Contatos</td>
    <td class=titulop align="center">Ramal</td>
    <td class=titulop align="center">Obs.</td>
</tr>
<%
linha=1
if rs.recordcount>0 then
For Contador=1 to registros

obs=""
if rs("codsituacao")="D" then
	corfonte="red"
elseif rs("codsituacao")="P" or rs("codsituacao")="E" or rs("codsituacao")="L" or rs("codsituacao")="I" then
	corfonte="green"
else
	corfonte="black"
end if
if rs("codsindicato")="03" then
	horario="Ver Secr.Curso " & rs("bloco")
else
	horario=rs("horario")
end if
if rs("codsituacao")="D" then horario=""
if rs("codsituacao")<>"D" then
sqlf="select tipo='Férias', venc=dtvencferias, ini=inicprogferias1, fim=fimprogferias1 " & _
"from corporerm.dbo.PFUNC where CHAPA='" & rs("chapa") & "' and GETDATE() between INICPROGFERIAS1 and fimprogferias1+1 " & _
"union " & _
"select tipo='Descanso', DTFIMPER, DTINIGOZO, dtfimgozo " & _
"from ferias where CHAPA='" & rs("chapa") & "' and GETDATE() between DTINIGOZO and DTFIMGOZO+1 "
rs2.Open sqlf, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then obs=rs2("tipo")
rs2.close
end if
if rs("codsindicato")="03" and (now>=dateserial(year(now)-1,12,24) and now<=dateserial(year(now),1,23)) then obs="Recesso"
%>
<tr>
    <td class="campop"><font color=<%=corfonte%> > <%=rs("nome")%></td>
    <td class=campo><font color=<%=corfonte%> > <%=rs("secao")%><br> -><%=rs("campus")%></td>
    <td class=campo><font color=<%=corfonte%> > <%=rs("funcao")%><br> -><%=horario%></td>
    <td class=campo><font color=<%=corfonte%> > <%=rs("telefone1")%> - <%=rs("telefone2")%><br> -><%=rs("email")%></td>
    <td class=campo nowrap><font color=<%=corfonte%> > <%=rs("fax")%></td>
    <td class="campor"><font color=<%=corfonte%> > <%=obs%></td>
</tr>
<%
if linha=1 then linha=0 else linha=1
rs.movenext
if rs.eof then exit for
Next

else 'sem registros
%>
<td class=grupo colspan=10>Esta seleção não mostra nenhum registro.</td>
<%
end if

rs.close
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</table>

</form>
</body>
</html>