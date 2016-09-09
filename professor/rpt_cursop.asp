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
<title>Relatório de Apontamento por Curso</title>
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
"ORDER BY gc.curso, qa.CHAPA "
	'response.write "<br>" & sql
end if
%>

<% if request.form="" then %>
<p class=titulo>Geração de relatório de apontamento por curso
<form method="POST" action="rpt_cursop.asp">
<p>Mês base para emissão: <select size="1" name="mesbase">
<%
sqla="SELECT mes_base FROM clc_cargap group by mes_base order by mes_base desc " 
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

<%
linha=2:inicio=0
tinj=0:tjd =0:trep=0:tjab=0:tdp =0:tae =0
rs.Open sql, ,adOpenStatic, adLockReadOnly

	response.write "<table border='0' cellpadding='2' width='650' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  >Ocorrências do Apontamento dos Docentes</td>"
	response.write "<td align='center'>Recursos Humanos</td>"
	response.write "<td align='right' >Mês-Base: " & mesbase & "</td>"
	response.write "</tr>":linha=1
	response.write "</table>"
	response.write "<table border='0' cellpadding='1' width='650' cellspacing='0' style='border-collapse: collapse'>"

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
	response.write "<tr><td class=titulor colspan=8>" & rs("sec") & " - " & rs("curso") & "</td></tr>"
	response.write "<tr>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' colspan=2>Docente   </td>"
	response.write "<td class=titulor colspan=2 align=""center"">Aulas Dadas</td>"
	response.write "<td class=titulor colspan=2 align=""center"">Orientação</td>"
	response.write "<td class=titulor colspan=2 align=""center"">Supervisão</td>"
	response.write "</tr>"
	linha=3
end if

if lastsecao<>rs("sec") then
	response.write "<tr>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' colspan=2>&nbsp;</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">&nbsp;</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">" & tau & "</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">&nbsp;</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">" & tor & "</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">&nbsp;</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">" & tsu & "</td>"
	response.write "</tr>":linha=linha+1
	tau=0:tor =0:tsu=0

	if inicio=1 then response.write "<tr><td class=""campor"" colspan=8>&nbsp;</td></tr>":linha=linha+1
	response.write "<tr><td class=titulor colspan=8>" & rs("sec") & " - " & rs("curso") & "</td></tr>":linha=linha+1
	response.write "<tr>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' colspan=2>Docente   </td>"
	response.write "<td class=titulor colspan=2 align=""center"">Aulas Dadas</td>"
	response.write "<td class=titulor colspan=2 align=""center"">Orientação</td>"
	response.write "<td class=titulor colspan=2 align=""center"">Supervisão</td>"
	response.write "</tr>":linha=linha+1

end if 'lastsecao

if rs("taulad")="" or isnull(rs("taulad")) then ev_aula="&nbsp;" else ev_aula=rs("sal")
if rs("torient")="" or isnull(rs("torient")) then ev_orient="&nbsp;" else ev_orient=rs("orient")
if rs("tsuperv")="" or isnull(rs("tsuperv")) then ev_superv="&nbsp;" else ev_superv=rs("superv")
%>
  <tr>
    <td class="campor" style="border-top: 1 dotted #000000" ><%=rs("chapa")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" ><%=rs("nome")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=ev_aula%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=rs("taulad")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=ev_orient%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=rs("torient")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=ev_superv%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=rs("tsuperv")%></td>
  </tr>
<%
lastsecao=rs("sec")
linha=linha+1
inicio=1
if rs("taulad") ="" or isnull(rs("taulad") ) then tau=tau else tau=tau+rs("taulad") 
if rs("torient")="" or isnull(rs("torient")) then tor=tor else tor=tor+rs("torient")   
if rs("tsuperv")="" or isnull(rs("tsuperv")) then tsu=tsu else tsu=tsu+rs("tsuperv")
rs.movenext
loop
rs.close

%>
  <tr>
    <td class="campor" style="border-top: 1px solid #000000" colspan=2>&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000" align="center">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000" align="center">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000" align="center"><%=tau%></td>
    <td class="campor" style="border-top: 1px solid #000000" align="center">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000" align="center"><%=tor%></td>
    <td class="campor" style="border-top: 1px solid #000000" align="center">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000" align="center"><%=tsu%></td>
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