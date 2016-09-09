<%@ Language=VBScript %>
<!-- #Include file="../ADOVBS.INC" -->
<!-- #Include file="../funcoes.INC" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a35")="N" or session("a35")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cadastro de Funcionários</title>
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
dim conexao, conexao2, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.cursorlocation = 3 'aduseclient
conexao.Open Application("conexao")
set conexao2=server.createobject ("ADODB.Connection")
conexao2.Open Application("conexao")
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao2
	
select case request("ordem")
	case "chapa"
		session("ordemfunc")="ORDER BY chapa "
	case "nome"
		session("ordemfunc")="ORDER BY p.nome, chapa "
	case "admissao"
		session("ordemfunc")="ORDER BY dataadmissao, p.nome "
	case "secao"
		session("ordemfunc")="ORDER BY codsecao, p.nome "
	case "funcaoc"
		session("ordemfunc")="ORDER BY codfuncao, p.nome "
	case "situacao"
		session("ordemfunc")="ORDER BY codsituacao, p.nome "
	case "recebimento"
		session("ordemfunc")="ORDER BY codrecebimento, p.nome "
	case "tipo"
		session("ordemfunc")="ORDER BY codtipo, p.nome "
	case "demissao"
		session("ordemfunc")="ORDER BY datademissao, p.nome "
	case "pis"
		session("ordemfunc")="ORDER BY pispasep "
	case "funcaon"
		session("ordemfunc")="ORDER BY fu.nome, p.nome "
	case else
		session("ordemfunc")=session("ordemfunc")
end select
if session("ordemfunc")="" then session("ordemfunc")="ORDER BY chapa "
	
if request.form("B2")<>"" then
	Session("PrimeiraVez")="Sim"

	if request.form("secao")="" then session("sel35")="Todas" else session("sel35")=request.form("secao")
	if request.form("funcao")="" then session("emp35")="Todas" else session("emp35")=request.form("funcao")
	if request.form("situacao")="" then session("req35")="Todas" else session("req35")=request.form("situacao")
	if request.form("localizar")="" then session("loc35")="" else session("loc35")=request.form("localizar")
	session("ltipo35")=request.form("tipo"):if session("ltipo35")="" then session("ltipo35")="Todos"

	select case session("ltipo35")
		case "A"
			session("sql35f")="and (p.codtipo='N' and p.codsindicato<>'03') "				
		case "E"
			session("sql35f")="and (p.codtipo='T') "
		case "P"
			session("sql35f")="and (p.codtipo='N' and p.codsindicato='03') "
		case "C"
			session("sql35f")="and (p.codtipo='A') "
		case else
			session("sql35f")=""
	end select

	if session("sel35")<>"Todas" then
		session("sql35b")="AND (p.codsecao='" & session("sel35") & "') "
	else
		session("sql35b")=""
	end if

	if session("emp35")<>"Todas" then
		session("sql35c")="AND (p.codfuncao='" & session("emp35") & "') "
	else
		session("sql35c")=""
	end if

	if session("req35")<>"Todas" then
		session("sql35e")="AND (p.codsituacao='" & session("req35") & "') "
	else
		session("sql35e")=""
	end if

	'if isnumeric(session("loc35"))=true then session("loc35")=numzero(session("loc35"),5)
	if session("loc35")<>"" then
		if isnumeric(session("loc35")) then
			if len(session("loc35"))=11 then
				session("sql35d")="AND (p.pispasep='" & session("loc35") & "') "
			else
				if isnumeric(session("loc35"))=true then session("loc35")=numzero(session("loc35"),5)
				session("sql35d")="AND (p.chapa like '%" & session("loc35") & "%') "
			end if
		else
			session("sql35d")="AND (p.nome like '%" & session("loc35") & "%') "
		end if
	else
		session("sql35d")=""
	end if
	if isnumeric(request.form("regpag")) then session("RegistrosporPagina")=request.form("regpag")
end if

registros=Session("RegistrosPorPagina")

sqla="select chapa, p.nome, dataadmissao, codsecao, codfuncao, codsituacao, codrecebimento, " & _
"codtipo, datademissao, pispasep, fu.nome as funcao, pe.dtnascimento, codsindicato " & _
"from corporerm.dbo.pfunc p, corporerm.dbo.pfuncao fu, corporerm.dbo.ppessoa pe " & _
"where p.codfuncao=fu.codigo and p.codpessoa=pe.codigo "
sqlb=""
sqlc=session("ordemfunc")

sql1=sqla & sqlb & session("sql35b") & session("sql35d") & session("sql35c") & session("sql35e") & session("sql35f") & sqlc

if Session("PrimeiraVez")<>"Nao" then
	'conexao.cursorlocation = 3 'aduseclient
	'conexao.open Application("consql")
	set rs=server.createobject ("ADODB.Recordset")
	rs.CacheSize = registros
	rs.PageSize = registros
	set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	Session("Pagina")=1
	MostraDados
	Session("PrimeiraVez")="Nao"
else
'	if request("folha")="" then
'      		pagina=1
'		else
'			pagina=request("folha")
'		end if
	if request("folha")="" then pagina=1
	if request.form("pagina")<>"" then pagina=request.form("pagina")
	if request("folha")<>"" then pagina=request("folha")
	Session("Pagina")=pagina
	'conexao.cursorlocation = 3 'aduseclient
	'conexao.open Application("consql")
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
<form method="POST" name="form" action="funcionarios.asp">
<p class=titulo style="margin-top: 0; margin-bottom: 0">Funcionários</p>
<table border="0" width="690" cellspacing="1" style="border-collapse: collapse" cellpadding="0">
<tr>
	<td class=campo width="70%" valign="top" align="left">Página: 
<%
atual=session("Pagina"):atual=cint(atual)
if atual=1 then
response.write "<img src='../images/setafirst0.gif' border='0'>"
response.write "<img src='../images/setaprevious0.gif' border='0'>"
else
response.write "<a href=""funcionarios.asp?folha=" & 1 & chr(34) & "><img src='../images/setafirst1.gif' border='0'></a>"
response.write "<a href=""funcionarios.asp?folha=" & atual-1 & chr(34) & "><img src='../images/setaprevious1.gif' border='0'></a>"
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
response.write "<a href=""funcionarios.asp?folha=" & atual+1 & chr(34) & "><img src='../images/setanext1.gif' border='0'></a>"
response.write "<a href=""funcionarios.asp?folha=" & rs.pagecount & chr(34) & "><img src='../images/setalast1.gif' border='0'></a>"
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
	<td class=fundor align="center"><a class=r href="funcionarios.asp?ordem=chapa"      >Chapa</a></td>
	<td class=fundor align="center"><a class=r href="funcionarios.asp?ordem=nome"       >Nome</a></td>
	<td class=fundor align="center"><a class=r href="funcionarios.asp?ordem=admissao"   >Admissão</a>    </td>
	<td class=fundor align="center"><a class=r href="funcionarios.asp?ordem=secao"      >Seção</a>       </td>
	<td class=fundor align="center"><a class=r href="funcionarios.asp?ordem=funcaoc"    >Cod.Função</a>  </td>
	<td class=fundor align="center"><a class=r href="funcionarios.asp?ordem=situacao"   >Situação</a>    </td>
	<td class=fundor align="center"><a class=r href="funcionarios.asp?ordem=recebimento">Cod.Rec.</a>    </td>
	<td class=fundor align="center"><a class=r href="funcionarios.asp?ordem=tipo"       >Tipo</a>        </td>
	<td class=fundor align="center"><a class=r href="funcionarios.asp?ordem=demissao"   >Demissão</a>    </td>
	<td class=fundor align="center"><a class=r href="funcionarios.asp?ordem=pis"        >PIS/PASEP</a>   </td>
	<td class=fundor align="center"><a class=r href="funcionarios.asp?ordem=funcaon"    >Descrição Função</a></td>
	<td class=fundor align="center">Aniv</td>
	<td class=fundor align="center">Fer</td>
	<td class=fundor align="center">FF</td>
</tr>
<%
linha=1
'rs.movefirst
'do while not rs.eof 
if rs.recordcount>0 then
For Contador=1 to registros
if month(rs("dtnascimento"))=month(now) then fmtaniv="<b>" else fmtaniv=""
aniv=numzero(day(rs("dtnascimento")),2) & "/" & fmtaniv & numzero(month(rs("dtnascimento")),2)
if rs("codsituacao")="D" then
	corfonte="red"
elseif rs("codsituacao")="P" or rs("codsituacao")="E" or rs("codsituacao")="L" or rs("codsituacao")="I" then
	corfonte="green"
else
	corfonte="black"
end if
if rs("codsindicato")="03" and rs("codsituacao")<>"D" then
sqlbloco="select bloco from blocos where codsecao='" & rs("codsecao") & "' "
rs2.Open sqlbloco, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
bloco="<font color='blue'>" & rs2("bloco") & "</font>"
end if
rs2.close
else
bloco=""
end if
%>
<tr>
	<td class="campor" align="center">
    <% if session("a35")="T" or session("a35")="C" then %>
      <a class=r href="funcionarios_ver.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'Funcionario_ver','690','500','yes','center');return false" onfocus="this.blur()">
	<%=rs("chapa")%></a>
	<%else%>
	<%=rs("chapa")%>
	<% end if %>
	</td>
	<td class="campor" nowrap><font color=<%=corfonte%> > <%=rs("nome")%></td>
	<td class="campor"><font color=<%=corfonte%> > <%=rs("dataadmissao") %></td>
	<td class="campor"><font color=<%=corfonte%> > <%=rs("codsecao") %></td>
	<td class="campor"><font color=<%=corfonte%> > <%=rs("codfuncao") %></td>
	<td class="campor"><font color=<%=corfonte%> > <%=rs("codsituacao")%></td>
	<td class="campor"><font color=<%=corfonte%> > <%=rs("codrecebimento")%></td>
	<td class="campor"><font color=<%=corfonte%> > <%=rs("codtipo") %></td>
	<td class="campor"><font color=<%=corfonte%> > <%=rs("datademissao")%><%=bloco%></td>
	<td class="campor"><font color=<%=corfonte%> > <%=rs("pispasep")%></td>
	<td class="campor" nowrap><font color=<%=corfonte%> > <%=rs("funcao")%></td>
	<td class="campor">&nbsp;<%=aniv%></td>
	<td class="campor">
		<a href="hstferias.asp?chapa=<%=rs("chapa")%>&nomefunc=<%=rs("nome")%>" onclick="NewWindow(this.href,'FichaFerias','800','300','yes','center');return false" onfocus="this.blur()">
	<img src="../images/ferias.gif" border="0" width=14 alt="Histórico de Férias"></a>
	</td>
	<td class="campor" nowrap>
		<a href="hstfichaf.asp?chapa=<%=rs("chapa")%>&nomefunc=<%=rs("nome")%>" onclick="NewWindow(this.href,'FichaFinanceira','800','500','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Moneybag.gif" border="0" width=13 alt="Ficha Financeira"></a>
<%if session("usuariomaster")="02379" or session("usuariomaster")="00259" or session("usuariomaster")="02552" then%>
		<a href="hstfichafc.asp?chapa=<%=rs("chapa")%>&nomefunc=<%=rs("nome")%>" onclick="NewWindow(this.href,'FichaFinanceira','800','500','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Moneybag.gif" border="0" width=13 alt="Ficha Financeira Caixa"></a>
<%end if%>
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
<tr><td class=grupo colspan=9>Esta seleção não mostra nenhum registro.</td></tr>
<%
end if
%>
</table>

<p><font size="1">
<%
sql2="select f.codsecao secao, s.descricao from corporerm.dbo.pfunc f, corporerm.dbo.psecao s where f.codsecao=s.codigo group by f.codsecao, s.descricao order by s.descricao"
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Seção: <select size="1" name="secao">
	<option value="Todas" <%if session("sel35")="Todas" then response.write "selected"%>>Todas Seções</option>
<%
rs2.movefirst
do while not rs2.eof
%>
	<option value="<%=rs2("secao")%>" <%if session("sel35")=rs2("secao") then response.write "selected"%>><%=rs2("secao") & " - " & rs2("descricao")%></option>
<%
rs2.movenext
loop
rs2.close
%>
</select>
&nbsp;&nbsp;

<%
sql2="select f.codfuncao as funcao, cf.nome from corporerm.dbo.pfunc f, corporerm.dbo.pfuncao cf where f.codfuncao=cf.codigo group by f.codfuncao, cf.nome order by cf.nome"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Função: <select size="1" name="funcao">
	<option value="Todas" <%if session("emp35")="Todas" then response.write "selected"%>>Todas Funções</option>
<%
rs2.movefirst
do while not rs2.eof
%>
	<option value="<%=rs2("funcao")%>" <%if session("emp35")=rs2("funcao") then response.write "selected"%>><%=rs2("funcao") & " - " & rs2("nome")%></option>
<%
rs2.movenext
loop
rs2.close
%>
</select>
<br>
<%
sql2="select codcliente, descricao from corporerm.dbo.pcodsituacao"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
Filtrar Situação: <select size="1" name="situacao">
	<option value="Todas" <%if session("req35")="Todos" then response.write "selected"%>>Todas Situações</option>
<%
rs2.movefirst
do while not rs2.eof
%>
	<option value="<%=rs2("codcliente")%>" <%if session("req35")=rs2("codcliente") then response.write "selected"%>><%=rs2("codcliente") & " - " & rs2("descricao")%></option>
<%
rs2.movenext
loop
rs2.close
%>
</select>
Filtrar Tipo: 
<select name="tipo">
	<option value="Todos" <%if session("ltipo35")="Todos" then response.write "selected"%>> Todos</option>
	<option value="A" <%if session("ltipo35")="A" then response.write "selected"%>> Administrativos</option>
	<option value="E" <%if session("ltipo35")="E" then response.write "selected"%>> Estagiários</option>
	<option value="P" <%if session("ltipo35")="P" then response.write "selected"%>> Professores</option>
	<option value="C" <%if session("ltipo35")="C" then response.write "selected"%>> Autonomos</option>
</select>
<br>
Localizar por nome/chapa: <input type="text" name="localizar" size=35 value="<%=session("loc35")%>">
Registros/Página: <input type="text" name="regpag" size=3 value="<%=Session("RegistrosPorPagina")%>">
</font>
<input name="B2" type="submit" class="button" value="Clique para Filtrar">
</form>
<img src="inseto.gif" width="75" height="56" border="0" alt="">
</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>