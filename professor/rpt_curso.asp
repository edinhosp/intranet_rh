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
	sql="SELECT qa.mes_base, gc.CODCCUSTO AS sec, gc.CURSO, qa.chapa, qa.NOME, Sum(qa.I) AS F_Inj, Sum(qa.Repos) AS Repos, Sum(qa.JD) AS F_JD, Sum(qa.JA) AS F_JAb, " & _
	"sum(case when i is null then 0 else i end + case when jd is null then 0 else jd end + case when ja is null then 0 else ja end) as totalfaltas, sum(aulas) as aulas, " & _
	"Sum(qa.DP) AS DP, Sum(qa.Extra) AS Extras, sum(qa.atraso) as atrasos, MIN(coordenador) as coord " & _
"FROM qry_apontamento qa INNER JOIN g2cursoeve gc ON qa.doc = gc.coddoc " & _
"GROUP BY qa.mes_base, gc.CODCCUSTO, gc.CURSO, qa.chapa, qa.NOME " & _
"HAVING qa.Mes_base='" & dtaccess(mesbase) & "' AND Sum(qa.Selec)<>0 " & _
"ORDER BY gc.curso, qa.CHAPA "
	'response.write "<br>" & sql
end if
%>

<% if request.form="" then %>
<p class=titulo>Geração de relatório de apontamento por curso
<form method="POST" action="rpt_curso.asp">
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

<%
linha=2:inicio=0
tinj=0:tjd =0:trep=0:tjab=0:tf=tinj+tjd+tjab:tatr=0
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
	response.write "<b>" & rs("sec") & " - " & rs("curso") & "</b>"
	response.write "<br>Coordenador: <b>" & rs("coord") & "</b>"

	response.write "<table border='0' cellpadding='1' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' colspan=2 rowspan=2>Docente   </td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' colspan=4 align=""center"">Faltas</td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' rowspan=2 align=""center"">Repos.</td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' rowspan=2 align=""center"">Atrasos<br>(min.)</td>"
	response.write "</tr><tr>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' align=""center"">Injust</td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' align=""center"">Justif</td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' align=""center"">Abon   </td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' align=""center"">Total  </td>"
	response.write "</tr>"
	linha=3
end if

if lastsecao<>rs("sec") then
	if inicio=1 then
	response.write "<tr>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' colspan=2>&nbsp;</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">" & tinj & "</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">" & tjd & "</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">" & tjab & "</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">" & tf & "</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">" & trep & "</td>"
	response.write "<td class=""campor"" style='border-top: 1px solid #000000' align=""center"">" & tatr & "</td>"
	response.write "</tr>"
	tinj=0:tjd =0:trep=0:tjab=0:tf=tinj+tjd+tjab:tatr=0
	
	pagina=pagina+1
	response.write "</table>"
	response.write "<br>"
	response.write "<p style='margin-top: 0; margin-bottom: 0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	end if
	response.write "<table border='0' cellpadding='2' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  >Ocorrências do Apontamento dos Docentes</td>"
	response.write "<td align='center'>Recursos Humanos</td>"
	response.write "<td align='right' >Mês-Base: " & mesbase & "</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<b>" & rs("sec") & " - " & rs("curso") & "</b>"
	response.write "<br>Coordenador: <b>" & rs("coord") & "</b>"

	response.write "<table border='0' cellpadding='1' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' colspan=2 rowspan=2>Docente   </td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' colspan=4 align=""center"">Faltas</td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' rowspan=2 align=""center"">Repos.</td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' rowspan=2 align=""center"">Atrasos<br>(min.)</td>"
	response.write "</tr><tr>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' align=""center"">Injust</td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' align=""center"">Justif</td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' align=""center"">Abon   </td>"
	response.write "<td class=titulor style='border-bottom: 1px solid #000000' align=""center"">Total  </td>"
	response.write "</tr>"
	linha=3
end if 'lastsecao

%>
  <tr>
    <td class="campor" style="border-top:1px dotted #000000" ><%=rs("chapa")%></td>
    <td class="campor" style="border-top:1px dotted #000000" ><%=rs("nome")%></td>
    <td class="campor" style="border-top:1px dotted #000000" align="center"><%=rs("f_inj")%></td>
    <td class="campor" style="border-top:1px dotted #000000" align="center"><%=rs("f_jd")%></td>
    <td class="campor" style="border-top:1px dotted #000000;border-right:1px dotted #000000" align="center"><%=rs("f_jab")%></td>
    <td class="campor" style="border-top:1px dotted #000000;border-right:1px dotted #000000" align="center"><%=rs("totalfaltas")%></td>
    <td class="campor" style="border-top:1px dotted #000000;border-right:1px dotted #000000" align="center"><%=rs("repos")%></td>
    <td class="campor" style="border-top:1px dotted #000000;border-right:1px dotted #000000" align="center"><%=rs("atrasos")%></td>
  </tr>
<%
lastsecao=rs("sec")
linha=linha+1
inicio=1
if rs("f_inj") ="" or isnull(rs("f_inj") ) then tinj=tinj else tinj=tinj+rs("f_inj")
if rs("f_jd")  ="" or isnull(rs("f_jd")  ) then tjd =tjd  else tjd =tjd +rs("f_jd") 
if rs("repos") ="" or isnull(rs("repos") ) then trep=trep else trep=trep+rs("repos")
if rs("f_jab") ="" or isnull(rs("f_jab") ) then tjab=tjab else tjab=tjab+rs("f_jab")
if rs("dp")    ="" or isnull(rs("dp")    ) then tdp =tdp  else tdp =tdp +rs("dp")   
if rs("extras")="" or isnull(rs("extras")) then tae =tae  else tae =tae +rs("extras")
if rs("atrasos")="" or isnull(rs("atrasos")) then tatr=tatr  else tatr=tatr +rs("atrasos")
tf=tinj+tjd+tjab
rs.movenext
loop
rs.close

%>
  <tr>
    <td class="campor" style="border-top: 1px solid #000000" colspan=2>&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000" align="center"><%=tinj%></td>
    <td class="campor" style="border-top: 1px solid #000000" align="center"><%=tjd%></td>
    <td class="campor" style="border-top: 1px solid #000000" align="center"><%=tjab%></td>
    <td class="campor" style="border-top: 1px solid #000000" align="center"><%=tf%></td>
    <td class="campor" style="border-top: 1px solid #000000" align="center"><%=trep%></td>
    <td class="campor" style="border-top: 1px solid #000000" align="center"><%=tatr%></td>
  </tr>
</table>
<%
linha=linha+1

if linha>69 then
	pagina=pagina+1
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