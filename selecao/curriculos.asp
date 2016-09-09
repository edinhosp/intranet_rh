<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a68")="N" or session("a68")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cadastro de Curriculos - Trabalhe Conosco</title>
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
	Session("PrimeiraVez")="Sim"

	if request.form("secao")="" then session("sel68")="Todas" else session("sel68")=request.form("secao")
	if request.form("funcao")="" then session("emp68")="Todas" else session("emp68")=request.form("funcao")
	if request.form("requisitante")="" then session("req68")="Todas" else session("req68")=request.form("requisitante")
	if request.form("tipo")="" then session("t68")="Todos" else session("t68")=request.form("tipo")

	if request.form("localizar")="" then session("loc68")="" else session("loc68")=request.form("localizar")
	
	if request.form("tipo")="Todos" then
		session("tipo68")=""
	elseif request.form("tipo")="02" then
		session("tipo68")="and rq.motivo='02' "
	elseif request.form("tipo")="03" then
		session("tipo68")="and rq.motivo='03' "
	else
		session("tipo68")="and rq.motivo='04' "
	end if

	if isnumeric(session("loc68"))=true then session("loc68")=numzero(session("loc68"),5)

	if session("sel68")<>"Todas" then
		session("sql68b")="AND (rq.secao='" & session("sel68") & "') "
	else
		session("sql68b")=""
	end if

	if session("emp68")<>"Todas" then
		session("sql68c")="AND (rq.funcao='" & session("emp68") & "') "
	else
		session("sql68c")=""
	end if

	if session("req68")<>"Todas" then
		session("sql68e")="AND (rq.requisitante='" & session("req68") & "') "
	else
		session("sql68e")=""
	end if

	if session("loc68")<>"" then
		if isnumeric(session("loc68")) then
			session("sql68d")="AND (rq.descricao like '%" & session("loc68") & "%') "
		else
			session("sql68d")="AND (rq.descricao like '%" & session("loc68") & "%') "
		end if
	else
		session("sql68d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if

registros=Session("RegistrosPorPagina")
lasttipo=session("tp68")
if lasttipo="" then lasttipo="Todos"

sqla="SELECT rq.id_requisicao, rq.descricao, rq.funcao, rq.secao, rq.requisitante, rq.motivo, " & _
"rq.dt_abertura, rq.dt_encerramento " & _
"FROM rs_requisicao rq " & _
"WHERE rq.id_requisicao>0 "
'sqla="select urc_id, urc_nome, 
sqlb=""
sqlc="ORDER BY rq.descricao, rq.dt_abertura "

sql1=sqla & sqlb & session("sql68b") & session("sql68d") & session("sql68c") & session("sql68e") & session("tipo68") & sqlc

if Session("PrimeiraVez")<>"Nao" then
	conexao.cursorlocation = 3 'aduseclient
	conexao.open Application("mysqlfieo")
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
	conexao.open Application("mysqlfieo")
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
<form method="POST" name="form" action="requisicao.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Curriculos - Trabalhe Conosco</p>
<table border="0" width="650" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="55%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""requisicao.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""requisicao.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
response.write "<a href=""requisicao.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""requisicao.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
end if
%>
	</td>
	<td class=campo width="30%" valign="top" align="center">
<% if session("a68")="T" then %>
<a href="requisicao_nova.asp" onclick="NewWindow(this.href,'Requisicao_nova','635','510','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
<font size="1">inserir nova requisição</font></a>
<% end if %>
	</td>
	<td class=campo width="15%" valign="top" align="right">
<%
Response.write "Registros: " & rs.recordcount
%>
	</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellspacing="0" width="650" cellpadding="1" style="border-collapse: collapse">
<tr>
    <td class=titulo align="center">Descrição   </td>
    <td class=titulo align="center">Função      </td>
    <td class=titulo align="center">Seção       </td>
    <td class=titulo align="center">Requisitante</td>
    <td class=titulo align="center">Motivo      </td>
    <td class=titulo align="center">Abertura    </td>
    <td class=titulo align="center">Encerr.     </td>
    <td class=titulo align="center">V           </td>
    <td class=titulo align="center">A           </td>
</tr>
<%
linha=1
'rs.movefirst
'do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
%>
<tr>
    <td class=campo>
		<a class=r href="requisicaover.asp?codigo=<%=rs("id_requisicao")%>" onclick="NewWindow(this.href,'Requisicao_ver','595','500','yes','center');return false" onfocus="this.blur()">
		<%if rs("descricao")="" or isnull(rs("descricao")) then response.write "Sem descrição" else response.write rs("descricao")%></a></td>
    <td class=campo><%=rs("funcao")%></td>
    <td class=campo><%=rs("secao") %></td>
    <td class=campo><%=rs("requisitante") %></td>
    <td class=campo><%=rs("motivo") %></td>
    <td class=campo align="center">&nbsp;<%=rs("dt_abertura")%></td>
    <td class=campo align="center">&nbsp;<%=rs("dt_encerramento")%></td>
	<td class=campo align="center">
    <% if session("a68")="T" then %>
      <a href="frm_requisicao.asp?codigo=<%=rs("id_requisicao")%>" onclick="NewWindow(this.href,'frm_requisicao','695','450','yes','center');return false" onfocus="this.blur()">
	<img src="../images/LeafSearch.gif" border="0" height=14 alt="Imprimir requisição"></a>
	<% end if %>
	</td>
	<td class=campo align="center">
    <% if session("a68")="T" then %>
      <a href="requisicao_alteracao.asp?codigo=<%=rs("id_requisicao")%>" onclick="NewWindow(this.href,'Requisicao_alterar','635','500','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Write.gif" border="0" height=14></a>
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
<td class=grupo colspan=9>Esta seleção não mostra nenhum registro.</td>
<%
end if
%>
</table>
<gr>
<font size="1">

Filtrar Seção: <select size="1" name="secao">
<option value="Todas" <%if session("sel68")="Todas" then response.write "selected"%>>Todas Seções</option>
<%
sql2="SELECT rq.secao, S.DESCRICAO FROM rs_requisicao rq, corporerm.dbo.PSECAO S " & _
"WHERE rq.secao = S.CODIGO collate database_default GROUP BY rq.secao, S.DESCRICAO " & _
"order by S.descricao"
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("secao")%>" <%if session("sel68")=rs2("secao") then response.write "selected"%>><%=rs2("secao") & " - " & rs2("descricao")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>
&nbsp;&nbsp;&nbsp;

Filtrar Função: <select size="1" name="funcao">
<option value="Todas" <%if session("emp68")="Todas" then response.write "selected"%>>Todas Funções</option>
<%
sql2="SELECT rq.funcao, f.nome FROM rs_requisicao rq, corporerm.dbo.PFUNCAO F " & _
"WHERE rq.funcao = f.CODIGO collate database_default GROUP BY rq.funcao, f.nome " & _
"order by f.nome"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("funcao")%>" <%if session("emp68")=rs2("funcao") then response.write "selected"%>><%=rs2("funcao") & " - " & rs2("nome")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>
<br>
Filtrar Requisitante: <select size="1" name="requisitante">
<option value="Todas" <%if session("req68")="Todos" then response.write "selected"%>>Todos Requisitantes</option>
<%
sql2="SELECT requisitante, nome FROM rs_requisicao, corporerm.dbo.PFUNC f WHERE requisitante=f.CHAPA collate database_default GROUP BY requisitante, NOME ORDER BY NOME "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
%>
	<option value="<%=rs2("requisitante")%>" <%if session("req402")=rs2("requisitante") then response.write "selected"%>><%=rs2("requisitante") & " - " & rs2("nome")%></option>
<%
rs2.movenext:loop
rs2.close
%>
</select>
Filtrar Motivo: 
<select name="tipo">
	<option value="Todos" <%if lasttipo="Todos" then response.write "selected"%>> Todos</option>
	<option value="02" <%if lasttipo="02" then response.write "selected"%>> Substituição</option>
	<option value="03" <%if lasttipo="03" then response.write "selected"%>> Vaga nova</option>
	<option value="04" <%if lasttipo="04" then response.write "selected"%>> Aumento de quadro</option>
</select>
<br>
Localizar por descrição: <input type="text" name="localizar" size=35 value="<%=session("loc68")%>">
Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
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