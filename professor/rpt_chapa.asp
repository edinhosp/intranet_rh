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
	sql="SELECT qa.mes_base, gc.CODCCUSTO AS sec, gc.CURSO, qa.chapa, qa.NOME, gc.falta, Sum(qa.I) AS F_Inj, Sum(qa.Repos) AS Repos, Sum(qa.JD) AS F_JD, Sum(qa.JA) AS F_JAb, " & _
	"gc.depen AS Dep, Sum(qa.DP) AS DP, gc.aextra AS AulaExtra, Sum(qa.Extra) AS Extras, sum(qa.atraso) as atrasos " & _
"FROM qry_apontamento qa INNER JOIN g2cursoeve gc ON qa.doc=gc.coddoc " & _
"GROUP BY qa.mes_base, gc.CODCCUSTO, gc.CURSO, qa.chapa, qa.NOME, gc.falta, gc.depen, gc.aextra " & _
"HAVING qa.Mes_base='" & dtaccess(mesbase) & "' AND Sum(qa.Selec)<>0 " & _
"ORDER BY qa.NOME, qa.CHAPA "
	'response.write "<br>" & sql
end if
%>

<% if request.form="" then %>
<p class=titulo>Geração de relatório de apontamento por chapa
<form method="POST" action="rpt_chapa.asp">
<p>Mês base para emissão: <select size="1" name="mesbase">
<%
sqla="SELECT mes_base FROM clc_carga group by mes_base order by mes_base desc " 
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
<table border="0" cellpadding="2" width="690" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td align="left"  >Ocorrências do Apontamento dos Docentes</td>
    <td align="center">Recursos Humanos</td>
    <td align="right" >Mês-Base: <%=mesbase%></td>
  </tr>
</table>
<table border="0" cellpadding="1" width="690" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td class=titulor colspan=2>Curso   </td>
    <td class=titulor align="center">Falta    </td>
    <td class=titulor align="center">Injust</td>
    <td class=titulor align="center">Justif</td>
    <td class=titulor align="center">Repos.</td>
    <td class=titulor align="center">Abon  </td>
    <td class=titulor colspan=1 align="center">Atrasos</td>
    <td class=titulor colspan=2 align="center">Depend.</td>
    <td class=titulor colspan=2 align="center">A.Extras</td>
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
	response.write "<table border='0' cellpadding='2' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  >Ocorrências do Apontamento dos Docentes</td>"
	response.write "<td align='center'>Recursos Humanos</td>"
	response.write "<td align='right' >Mês-Base: " & mesbase & "</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table border='0' cellpadding='1' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=titulor colspan=2>Curso   </td>"
	response.write "<td class=titulor align=""center"">Falta    </td>"
	response.write "<td class=titulor align=""center"">Injust</td>"
	response.write "<td class=titulor align=""center"">Justif</td>"
	response.write "<td class=titulor align=""center"">Repos.</td>"
	response.write "<td class=titulor align=""center"">Abon  </td>"
	response.write "<td class=titulor colspan=1 align=""center"">Atrasos</td>"
	response.write "<td class=titulor colspan=2 align=""center"">Depend.</td>"
	response.write "<td class=titulor colspan=2 align=""center"">A.Extras</td>"
	response.write "</tr>"
	linha=2
end if

if lastchapa<>rs("chapa") then
%>
  <tr>
    <td class="campor" style="border-top: 1px solid #000000"><%=rs("chapa")%></td>
    <td class="campor" style="border-top: 1px solid #000000" colspan=2><b><%=rs("nome")%></b></td>
    <td class="campor" style="border-top: 1px solid #000000" colspan=9>&nbsp;</td>
  </tr>
<%
linha=linha+1
end if 'lastchapa

if rs("extras")="" or isnull(rs("extras")) then ev_ae="&nbsp;" else ev_ae=rs("aulaextra")
if rs("dp")="" or isnull(rs("dp")) then ev_dp="&nbsp;" else ev_dp=rs("dep")
if (rs("f_jd")="" or isnull(rs("f_jd"))) and _
	(rs("f_inj")="" or isnull(rs("f_inj"))) then
		ev_fal="&nbsp;" 
	else 
		ev_fal=rs("falta")
	end if
%>
  <tr>
    <td class="campor" style="border-top: 1 dotted #000000" ><%=rs("sec")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" ><%=rs("curso")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=ev_fal%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=rs("f_inj")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=rs("f_jd")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=rs("repos")%></td>
    <td class="campor" style="border-top: 1 dotted #000000;border-right:1px dotted" align="center"><%=rs("f_jab")%></td>
    <td class="campor" style="border-top: 1 dotted #000000;border-right:1px dotted" align="center"><%=rs("atrasos")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=ev_dp%></td>
    <td class="campor" style="border-top: 1 dotted #000000;border-right:1px dotted" align="center"><%=rs("dp")%></td>
    <td class="campor" style="border-top: 1 dotted #000000" align="center"><%=ev_ae%></td>
    <td class="campor" style="border-top: 1 dotted #000000;border-right:1px dotted" align="center"><%=rs("extras")%></td>
  </tr>
<%
lastchapa=rs("chapa")
linha=linha+1

rs.movenext
loop
rs.close

%>
  <tr>
    <td class="campor" style="border-top: 1px solid #000000" colspan=12>&nbsp;</td>
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