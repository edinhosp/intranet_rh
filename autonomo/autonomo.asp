<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a52")="N" or session("a52")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
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

if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"

	if request.form("secao")="" then
		session("sel52")="Todas"
	else
		session("sel52")=request.form("secao")
	end if

	if request.form("localizar")="" then
		session("loc52")=""
	else
		session("loc52")=request.form("localizar")
	end if
		
	'if isnumeric(session("loc52"))=true then session("loc52")=numzero(session("loc52"),5)

	if session("sel52")<>"Todas" then
		session("sql52b")="AND (a.tipo_prestacao='" & session("sel52") & "') "
	else
		session("sql52b")=""
	end if

	if session("loc52")<>"" then
		if isnumeric(session("loc52")) then
			session("sql52d")="AND ((a.cpf like '%" & session("loc52") & "%') "
			session("sql52d")=session("sql52d") & "or (a.nit like '%" & session("loc52") & "%')) "
		else
			session("sql52d")="AND (a.nome_autonomo like '%" & session("loc52") & "%') "
		end if
	else
		session("sql52d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if

registros=Session("RegistrosPorPagina")
lasttipo=request.form("tipo")
if lasttipo="" then lasttipo="Todos"

sqla="SELECT a.id_autonomo, a.nome_autonomo, a.tipo_prestacao, a.cpf, a.nit, a.telefone, a.cbo  " & _
", rua, numero, complemento, bairro, cidade, estado, cep, celular, conta,agencia " & _
"FROM autonomo a " & _
"WHERE a.id_autonomo>0 "
sqlb=""
sqlc="ORDER BY a.nome_autonomo "

sql1=sqla & sqlb & session("sql52b") & session("sql52d") & sqlc

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
<form method="POST" name="form" action="autonomo.asp">
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
	response.write "<a href=""autonomo.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
	response.write "<a href=""autonomo.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
	response.write "<a href=""autonomo.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
	response.write "<a href=""autonomo.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
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
	<td class=titulor align="center">Nome autônomo </td>
	<td class=titulor align="center">Tipo prestação</td>
	<td class=titulor align="center">C.P.F.        </td>
	<td class=titulor align="center">PIS/NIT       </td>
	<td class=titulor align="center">Telefone      </td>
	<td class=titulor align="center">CBO           </td>
	<td class=titulor align="center">End.          </td>
	<td class=titulor align="center">Cta.          </td>
	<td class=titulor align="center">&nbsp;        </td>
	<td class=titulor align="center">RPA           </td>
	<td class=fundor align="center">Contr.         </td>
</tr>
<%
linha=1
'rs.movefirst:do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
cpf=rs("cpf")
if cpf<>"" then cpf=left(cpf,3) & "." & mid(cpf,4,3) & "." & mid(cpf,7,3) & "-" & right(cpf,2)
if rs("rua")<>"" and rs("cidade")<>"" and rs("cep")<>"" then endereco=1 else endereco=0
if rs("agencia")<>"" and rs("conta")<>"" then conta=1 else conta=0
%>
<tr>
	<td class="campor"><%=rs("nome_autonomo") %></td>
	<td class="campor"><%=rs("tipo_prestacao")%></td>
	<td class="campor" nowrap>&nbsp;<%=cpf %>&nbsp;</td>
	<td class="campor">&nbsp;<%=rs("nit") %>&nbsp;</td>
	<td class="campor" nowrap><%=rs("telefone") %></td>
	<td class="campor"><%=rs("cbo")%></td>
	<td class="campor" align="center"><%if endereco=0 then response.write "<img src='../images/bullet.gif'>" else response.write "<img src='../images/bullet_hl.gif'>"%></td>
	<td class="campor" align="center"><%if conta=0 then response.write "<img src='../images/bullet.gif'>" else response.write "<img src='../images/bullet_hl.gif'>"%></td>
	<td class="campor" align="center">
	<% if session("a52")="T" or session("a52")="C" then %>
		<a href="autonomo_alteracao.asp?codigo=<%=rs("id_autonomo")%>" onclick="NewWindow(this.href,'AlteracaoAutonomo','510','330','no','center');return false" onfocus="this.blur()">
		<img src="../images/Folder95O.gif" border="0" width=13 alt="Alterar os dados cadastrais"></a>
	<% end if %>
	</td>
	<td class="campor" align="center">
	<% if session("a52")="T" or session("a52")="C" then %>
		<a href="rpa.asp?codigo=<%=rs("id_autonomo")%>" onclick="NewWindow(this.href,'RPA','600','400','yes','center');return false" onfocus="this.blur()">
		<img src="../images/Money.gif" border="0" width=13 alt="Ver controle de pagamentos"></a>
	<% end if %>
	</td>
	<td class="campor" align="center">
		<a href="autonomo_contrato.asp?codigo=<%=rs("id_autonomo")%>" onclick="NewWindow(this.href,'ContratoAutonomo','690','400','yes','center');return false" onfocus="this.blur()">
		<img src="../images/Leaf.gif" width="16" border="0" alt="Emitir Contrato de Prestação de Serviços"></a>
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
<td class=grupo colspan=11>Esta seleção não mostra nenhum registro.</td>
<%
end if
%>
</table>
<%
if session("a52")="T" then
%>
<a href="autonomo_nova.asp" onclick="NewWindow(this.href,'InclusaoAutonomo','510','330','no','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif" alt="Cadastrar novo autonomo">
<font size="1">inserir novo autônomo</font></a><br>
<%
end if
%>
<font size="1">
<%
sql2="select tipo_prestacao as servico from autonomo group by tipo_prestacao"
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Serviço: <select size="1" name="secao">
<option value="Todas" <%if session("sel52")="Todas" then response.write "selected"%>>Todos serviços</option>
<%
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
%>
    <option value="<%=rs2("servico")%>" <%if session("sel52")=rs2("servico") then response.write "selected"%>><%=rs2("servico")%></option>
<%
rs2.movenext
loop
end if 'rs2.recordcount
rs2.close
%>
</select>
<br>
Localizar por nome/CPF: <input type="text" name="localizar" size=35 value="<%=session("loc52")%>">
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