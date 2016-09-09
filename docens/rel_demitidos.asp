<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a17")="N" or session("a17")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Docentes Demitidos</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if (month(request.form("dtfim"))=6 or month(request.form("dtfim"))=12) and year(request.form("dtfim"))=year(now) then redutor=12 else redutor=0

sqla="SELECT f.CODSITUACAO, f.CHAPA, f.NOME AS NomeF, f.DATAADMISSAO, f.DATADEMISSAO, s.descricao as secao, f.CODSECAO, f.FUNCAO, f.TITULACAOPAGA, f.INSTRUCAOMEC " & _
"FROM dc_professor f, corporerm.dbo.psecao s " & _
"WHERE f.codsecao=s.codigo and f.CODSITUACAO='D' and datademissao between '" & dtaccess(request.form("dtini")) & "' and (cast('" & dtaccess(request.form("dtfim")) & "' as datetime)-cast(" & redutor & " as datetime)) "
select case request.form("R1")
	case "nome"
		sqlb="ORDER BY f.nome"
	case "demissao"
		sqlb="ORDER BY f.datademissao"
	case "setor"
		sqlb="ORDER BY f.secao, f.nome"
end select
sql1=sqla & sqlb

if request.form="" then
%>
<form method="POST" action="rel_demitidos.asp">
<p style="margin-bottom: 0"><font color="#0000FF"><b>Ordem para emissão do relatório de &quot;Docentes
  Demitidos&quot;</b></font></p>
  <table border="0" cellpadding="0" cellspacing="3">
    <tr>
      <td class=campo>entre <input type="text" name="dtini" size=8 class=a value="<%=dateserial(year(now),1,1)%>"> 
	  e <input type="text" name="dtfim" size=8 class=a value="<%=dateserial(year(now),12,31)%>"></td>
    </tr>
    <tr>
      <td class=campo><input type="radio" name="R1" value="nome" checked>por ordem de nome</td>
    </tr>
    <tr>
      <td class=campo><input type="radio" name="R1" value="demissao">por ordem de data de demissão</td>
    </tr>
    <tr>
      <td class=campo><input type="radio" name="R1" value="setor">por ordem de setor / nome</td>
    </tr>
  </table>
  <p><input type="submit" class=button value="Visualizar relatório" name="B1"></p>
</form>
<p style="margin-top: 0; margin-bottom: 0"><font color="#FF0000">Configure a página do seu navegador (Internet
Explorer, Netscape, Mozilla, etc) no sentido RETRATO.</font></p>
<%
end if

if request.form<>"" then
%>
<p class=titulo>Docentes Demitidos
<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
    <td class="titulor" align="center">Chapa</td>
    <td class="titulor" align="center">Nome</td>
    <td class="titulor" align="center">Admissão</td>
    <td class="titulor" align="center">Saída</td>
    <td class="titulor" align="center">Curso/Seção</td>
  </tr>
<%
linhas=2
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 
chapach=rs("chapa")
session("chapa")=chapach
if linhas>62 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<p style='margin-top:0; margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<p style='margin-top:0; margin-bottom:0' class=""titulo"">Docentes Demitidos"
	linhas=1
	response.write "<table border='1' bordercolor=#000000 cellpadding='1' cellspacing='0' style='border-collapse: collapse' width='650'>"
	response.write "<tr>"
	response.write "<td class=""titulor"" align=""center"">Chapa</td>"
	response.write "<td class=""titulor"" align=""center"">Nome</td>"
	response.write "<td class=""titulor"" align=""center"">Admissão</td>"
	response.write "<td class=""titulor"" align=""center"">Saida</td>"
	response.write "<td class=""titulor"" align=""center"">Curso/Seção</td>"
	response.write "</tr>"
	linhas=linhas+1
end if
%>
  <tr>
    <td class="campor" align="center">&nbsp;<%=rs("chapa")%></td>
    <td class="campor">&nbsp;<%=rs("nomef")%></td>
    <td class="campor" align="center">&nbsp;<%=rs("dataadmissao")%></td>
    <td class="campor" align="center">&nbsp;<%=rs("datademissao")%></td>
    <td class="campor">&nbsp;<%=rs("secao")%></td>
  </tr>
<%
linhas=linhas+1
rs.movenext
loop
else
%>
<tr>
    <td class="campor" colspan=5>-----------------------------------</td>
</tr>
<%
end if 'rs.recordcount
rs.close
%>
</table>
<%
	pagina=pagina+1
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
%>
<%
end if 'request.form
%>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>