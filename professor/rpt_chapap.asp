<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a89")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Relatório de Apontamento por Chapa</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, rt(10), rd(10)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	mesbase=request.form("mesbase")
	if request.form("resumo")="ON" then resumo=1 else resumo=0
	sessao=session.sessionid
	sql="SELECT qa.mes_base, gc.CODCCUSTO AS sec, gc.CURSO, qa.chapa, qa.NOME, gc.sal, Sum(qa.aula_prev) AS taulap, Sum(qa.aula_dada) AS taulad, gc.orient, Sum(qa.orientacao) AS torient, gc.superv, Sum(qa.supervisao) AS tsuperv " & _
"FROM qry_apontamentop AS qa INNER JOIN g2cursoeve AS gc ON qa.doc = gc.coddoc " & _
"GROUP BY qa.mes_base, gc.CODCCUSTO, gc.CURSO, qa.chapa, qa.NOME, gc.sal, gc.orient, gc.superv " & _
"HAVING qa.Mes_base='" & dtaccess(mesbase) & "' AND Sum(qa.Selec)<>0 " & _
"ORDER BY qa.NOME, qa.chapa;"
	'response.write "<br>" & sql
end if
%>

<% if request.form="" then %>
<p class=titulo>Geração de relatório de apontamento por chapa
<form method="POST" action="rpt_chapap.asp">
<p>Mês base para emissão: <select size="1" name="mesbase">
<%
sqla="SELECT mes_base FROM clc_cargap group by mes_base order by mes_base desc" 
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst
mesbase=rsc("mes_base")
do while not rsc.eof
%>
          <option value="<%=rsc("mes_base")%>" <%=tempt%>><%=rsc("mes_base")%></option>
<%
rsc.movenext
loop
rsc.close
%>
        </select></p>		
  <p><input type="submit" value="Visualizar relatório" name="Gerar" class="button"></p>
</form>
<%
else
%>
<table border="0" cellpadding="2" width="650" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td align="left"  >Ocorrências do Apontamento dos Docentes</td>
    <td align="center">Recursos Humanos</td>
    <td align="right" >Mês-Base: <%=mesbase%></td>
  </tr>
</table>
<table border="0" cellpadding="1" width="650" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td class=titulor colspan=2>Curso   </td>
    <td class=titulor colspan=2 align="center">Aulas Dadas</td>
    <td class=titulor colspan=2 align="center">Orientação</td>
    <td class=titulor colspan=2 align="center">Supervisão</td>
  </tr>
<%
linha=2
rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof
if linha>69 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<br>"
	response.write "<p style='margin-top: 0; margin-bottom: 0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<table border='0' cellpadding='2' width='650' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  >Ocorrências do Apontamento dos Docentes</td>"
	response.write "<td align='center'>Recursos Humanos</td>"
	response.write "<td align='right' >Mês-Base: " & mesbase & "</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table border='0' cellpadding='1' width='650' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=titulor colspan=2>Curso   </td>"
	response.write "<td class=titulor colspan=2 align=""center"">Aulas Dadas</td>"
	response.write "<td class=titulor colspan=2 align=""center"">Orientação</td>"
	response.write "<td class=titulor colspan=2 align=""center"">Supervisão</td>"
	response.write "</tr>"
	linha=2
end if

if lastchapa<>rs("chapa") then
%>
  <tr>
    <td class="campor" style="border-top: 1px solid #000000"><%=rs("chapa")%></td>
    <td class="campor" style="border-top: 1px solid #000000" colspan=2><b><%=rs("nome")%></b></td>
    <td class="campor" style="border-top: 1px solid #000000" colspan=8>&nbsp;</td>
  </tr>
<%
linha=linha+1
end if 'lastchapa

if rs("taulad")="" or isnull(rs("taulad")) then ev_aula="&nbsp;" else ev_aula=rs("sal")
if rs("torient")="" or isnull(rs("torient")) then ev_orient="&nbsp;" else ev_orient=rs("orient")
if rs("tsuperv")="" or isnull(rs("tsuperv")) then ev_superv="&nbsp;" else ev_superv=rs("superv")
%>
  <tr>
    <td class="campor" style="border-top: 1 dotted #000000" ><%=rs("sec")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" ><%=rs("curso")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=ev_aula%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=rs("taulad")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=ev_orient%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=rs("torient")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=ev_superv%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=rs("tsuperv")%></td>
  </tr>
<%
lastchapa=rs("chapa")
linha=linha+1

rs.movenext
loop
rs.close

%>
  <tr>
    <td class="campor" style="border-top: 1px solid #000000" colspan=11>&nbsp;</td>
  </tr>
</table>
<%
linha=linha+1

if linha>69 then
	pagina=pagina+1
	response.write "<br>"
	response.write "<p style='margin-top: 0; margin-bottom: 0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<table border='0' cellpadding='2' width='650' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  >Ocorrências do Apontamento dos Docentes</td>"
	response.write "<td align='center'>Recursos Humanos</td>"
	response.write "<td align='right' >Mês-Base: " & mesbase & "</td>"
	response.write "</tr>"
	response.write "</table>"
	linha=1
end if

linha=linha+1
pagina=pagina+1
response.write "<br>"
response.write "<p style='margin-top: 0; margin-bottom: 0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"

end if
%>
</body>
</html>
<%
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%>