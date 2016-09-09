<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a81")="N" or session("a81")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Assistência Médica</title>
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

	if request.form("secao")=""     then session("sel81")="Todas" else session("sel81")=request.form("secao")
	if request.form("empresa")=""   then session("emp81")="Todas" else session("emp81")=request.form("empresa")
	if request.form("localizar")="" then session("loc81")=""      else session("loc81")=request.form("localizar")
		
	if request.form("tipo")="Todos" then
		session("tipo81")=""
	elseif request.form("tipo")="Prof" then
		session("tipo81")="and f.codsindicato='03' "
	else
		session("tipo81")="and f.codsindicato<>'03' "
	end if

	if isnumeric(session("loc81"))=true then session("loc81")=numzero(session("loc81"),5)

	if session("sel81")<>"Todas" then
		session("sql81b")="AND (f.codsecao='" & session("sel81") & "') "
	else
		session("sql81b")=""
	end if

	if session("emp81")<>"Todas" then
		session("sql81c")="AND (a.empresa='" & session("emp81") & "') "
	else
		session("sql81c")=""
	end if

	if session("loc81")<>"" then
		if isnumeric(session("loc81")) then
			session("sql81d")="AND (a.chapa like '%" & session("loc81") & "%') "
		else
			session("sql81d")="AND (f.nome like '%" & session("loc81") & "%') "
		end if
	else
		session("sql81d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if
registros=Session("RegistrosPorPagina")
lasttipo=request.form("tipo")
if lasttipo="" then lasttipo="Todos"

sqla="SELECT a.empresa, a.CHAPA, f.NOME, f.CODSITUACAO, f.DATAADMISSAO, f.DATADEMISSAO, " & _
"f.CODSECAO, f.CODSINDICATO " & _
"FROM assmed_beneficiario a, corporerm.dbo.PFUNC f " & _
"WHERE a.CHAPA = f.CHAPA collate database_default "
sqlb=""
sqlc="ORDER BY f.NOME, a.empresa "

sql1=sqla & sqlb & session("sql81b") & session("sql81d") & session("sql81c") & session("tipo81") & sqlc

if Session("PrimeiraVez")<>"Nao" then
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
<form method="POST" name="form" action="controle.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Controle de Assistência Médica</p>
<table border="0" width="600" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
  <tr>
    <td class=campo width="70%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""controle.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""controle.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
response.write "<a href=""controle.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""controle.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
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

<table border="1" cellspacing="0" width="600" cellpadding="0" style="border-collapse: collapse">
<tr>
    <td class=titulo align="center">Empresa </td>
    <td class=titulo align="center">Chapa   </td>
    <td class=titulo align="center">Nome    </td>
    <td class=titulo align="center">Sit.    </td>
    <td class=titulo align="center">Seção   </td>
    <td class=titulo align="center">Sind.   </td>
    <td class=titulo align="center">Admissão</td>
    <td class=titulo align="center">Saída   </td>
    <td class=titulo align="center">X       </td>
</tr>
<%
linha=1
'rs.movefirst
'do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
%>
<tr>
    <td class=campo align="center"><%=rs("empresa") %></td>
    <td class=campo align="center"><%=rs("chapa")%></td>
    <td class=campo><%=rs("nome") %></td>
    <td class=campo align="center"><%=rs("codsituacao") %></td>
    <td class=campo align="center"><%=rs("codsecao") %></td>
    <td class=campo align="center"><%=rs("codsindicato")%></td>
    <td class=campo align="center">&nbsp;<%=rs("dataadmissao")%></td>
    <td class=campo align="center">&nbsp;<%=rs("datademissao")%></td>
	<td class=campo align="center">
    <% if session("a81")="T" or session("a81")="C" then %>
      <a href="controle_ver.asp?codigo=<%=rs("chapa")%>" onclick="NewWindow(this.href,'Alteracao','690','500','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
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
<font size="1">
<%
sql2="SELECT S.CODIGO, S.DESCRICAO FROM assmed_beneficiario B, corporerm.dbo.PFUNC F, corporerm.dbo.PSECAO S " & _
"WHERE B.CHAPA=F.CHAPA collate database_default AND F.CODSECAO = S.CODIGO GROUP BY S.CODIGO, S.DESCRICAO " & _
"order by S.descricao"
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Seção: <select size="1" name="secao">
<option value="Todas" <%if session("sel81")="Todas" then response.write "selected"%>>Todas Seções</option>
<%
rs2.movefirst
do while not rs2.eof
%>
    <option value="<%=rs2("codigo")%>" <%if session("sel81")=rs2("codigo") then response.write "selected"%>><%=rs2("codigo") & " - " & rs2("descricao")%></option>
<%
rs2.movenext
loop
rs2.close
%>
</select>
<br>
Empresa de Saúde: <select size="1" name="empresa">
  <option value="Todas">Todas Empresas</option>
<%
sql2="select codigo, operadora from assmed_empresa order by operadora"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
%>
    <option value="<%=rs2("codigo")%>" <%if session("emp81")=rs2("codigo") then response.write "selected"%>><%=rs2("codigo") & " - " & rs2("operadora")%></option>
<%
rs2.movenext
loop
rs2.close
%>
</select>
	<input type="radio" name="tipo" value="Todos" <%if lasttipo="Todos" then response.write "checked"%>> Todos 
	<input type="radio" name="tipo" value="Prof"  <%if lasttipo="Prof"  then response.write "checked"%>> Professores 
	<input type="radio" name="tipo" value="Adm"   <%if lasttipo="Adm"   then response.write "checked"%>> Administrativos
<br>
Localizar por nome/chapa: <input type="text" name="localizar" size=35 value="<%=session("loc81")%>">
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