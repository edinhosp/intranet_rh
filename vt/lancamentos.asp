<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a83")="N" or session("a83")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Lançamentos Vale-Transporte</title>
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
registros=Session("RegistrosPorPagina")
registros=500
dim conexao, conexao2
dim rs, rs2
set conexao=server.createobject ("ADODB.Connection")
'conexao.Open Application("consql")

if request.form("B2")<>"" then
	Session("PrimeiraVez83")="Sim"

	if request.form("secao")="" then session("sel83")="Todas" else session("sel83")=request.form("secao")
	if request.form("localizar")="" then session("loc83")="" else session("loc83")=request.form("localizar")
	'if isnumeric(session("loc83"))=true then session("loc83")=numzero(session("loc83"),5)

	if session("sel83")<>"Todas" then
		session("sql83b")="AND (a.tipo_prestacao='" & session("sel83") & "') "
	else
		session("sql83b")=""
	end if

	if session("loc83")<>"" then
		if isnumeric(session("loc83")) then
			session("sql83d")="AND ((a.cpf like '%" & session("loc83") & "%') "
			session("sql83d")=session("sql83d") & "or (a.nit like '%" & session("loc83") & "%')) "
		else
			session("sql83d")="AND (a.nome_autonomo like '%" & session("loc83") & "%') "
		end if
	else
		session("sql83d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")
lasttipo=request.form("tipo")
if lasttipo="" then lasttipo="Todos"

sqla="SELECT vt_saldo.id_saldo, vt_saldo.data, vt_saldo.id_tipo, vt_saldo_tipo.tipo, vt_saldo_tipo.fator, " & _
"vt_saldo.codigo, PTARIFA.DESCRICAO, PTARIFA.VALOR, vt_saldo.tarifa, vt_saldo.quantidade, vt_saldo.total, " & _
"vt_saldo.chapa, PFUNC.NOME " & _
"FROM ((vt_saldo INNER JOIN vt_saldo_tipo ON vt_saldo.id_tipo = vt_saldo_tipo.id_tipo) LEFT JOIN corporerm.dbo.PFUNC PFUNC ON vt_saldo.chapa = PFUNC.CHAPA COLLATE DATABASE_DEFAULT) INNER JOIN corporerm.dbo.PTARIFA PTARIFA ON vt_saldo.codigo = PTARIFA.CODIGO COLLATE DATABASE_DEFAULT "
sqlb="WHERE vt_saldo.deletada=0 and getdate() between iniciovigencia and finalvigencia "
sqlc="ORDER BY vt_saldo.data desc, vt_saldo.id_tipo, vt_saldo.codigo; "

sql1=sqla & sqlb & session("sql83b") & session("sql83d") & sqlc

if Session("PrimeiraVez83")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	conexao.open Application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	Session("Pagina")=1
	MostraDados
	Session("PrimeiraVez83")="Nao"
else
	if request("folha")="" then pagina=1
	if request.form("pagina")<>"" then pagina=request.form("pagina")
	if request("folha")<>"" then pagina=request("folha")
	Session("Pagina")=pagina
	conexao.cursorlocation = 3 'aduseclient
	conexao.open Application("conexao")
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
<form method="POST" name="form" action="lancamentos.asp">
<p class="titulo" style="margin-top: 0; margin-bottom: 0">Lançamentos - Vale-Transporte</p>
<table border="0" width="650" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class="campo" width="70%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""lancamentos.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""lancamentos.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
response.write "<a href=""lancamentos.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""lancamentos.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
	<td class="campo" width="30%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount & _
   "<br> Página atual: " & Session("Pagina") & "/" & rs.pagecount & ""
%>
	</td>
</tr>
</table>

<table border="1" cellspacing="0" cellpadding="1" style="border-collapse: collapse" width="690">
<tr>
	<td class="titulo" align="center">Data </td>
	<td class="titulo" align="center">Movimento</td>
	<td class="titulo" align="center">VT</td>
	<td class="titulo" align="center">Funcionário</td>
	<td class="titulo" align="center">Tarifa</td>
	<td class="titulo" align="center">Quantidade</td>
	<td class="titulo" align="center">Total</td>
	<td class="titulo" align="center">&nbsp;    </td>
</tr>
<%
linha=1
'rs.movefirst:do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
if rs("fator")=-1 then estilo="<font color=red>" else estilo=""
%>
<tr>
	<td class="campo"><%=rs("data") %></td>
	<td class="campo"><%=rs("tipo")%></td>
	<td class="campo"><%=rs("descricao")%></td>
	<td class="campo"><%=rs("nome")%>&nbsp;(<%=rs("chapa")%>)</td>
	<td class="campo" align="right"><%=formatnumber(rs("tarifa"),2)%>&nbsp;</td>
	<td class="campo" align="right"><%=estilo%><%=formatnumber(rs("quantidade"),0)%>&nbsp;</td>
	<td class="campo" align="right"><%=estilo%><%=formatnumber(rs("total"),2)%>&nbsp;</td>
	<td class="campo" align="center">
	<% if session("a83")="T" then %>
	<% if rs("id_tipo")<>0 then %>
	<a href="mov_alteracao.asp?codigo=<%=rs("id_saldo")%>" onclick="NewWindow(this.href,'AlteracaoVT','420','200','no','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width="13"></a>
	<% end if %>
	<% end if %>
	</td>
</tr>
<%
if linha=1 then linha=0 else linha=1
rs.movenext
if rs.eof then exit for
'loop
Next

else 'sem registros
%>
<td class="grupo" colspan="10">Esta seleção não mostra nenhum registro.</td>
<%
end if
%>
</table>
<% if session("a83")="T" then %>
<a href="mov_nova.asp" onclick="NewWindow(this.href,'InclusaoVT','420','200','no','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif" WIDTH="16" HEIGHT="16">
<font size="1">inserir novo lançamento</font></a>
<% end if %>
<br>
<font size="1">
Registros/Página: <input type="text" name="regpag" size="3" value="<%=Session("RegistrosPorPagina")%>">
</font>
<input name="B2" type="submit" class="button" value="Clique para Filtrar">
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