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
<title>Capas para Apontamento de Ponto Docente</title>
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
sql="TRANSFORM Sum(IIf([Selec]=0,Null,[Selec])) AS Expr3 " & _
"SELECT qa.mes_base, gc.CODCCUSTO AS sec, gc.CURSO, qa.chapa, qa.NOME, IIf(([I])>0,'Faltas ',IIf(([Repos])>0,'Reposição ',IIf(([atraso])>0,'Atrasos ',IIf(([JD])>0,'Faltas Just ',IIf(([JA])>0,'Faltas Abon ',IIf(([DP])>0,'Dependentes ',IIf(([extra])>0,'A.Extras ',Null))))))) AS Grupam, Sum(IIf([Selec]=0,Null,[Selec])) AS Total1 " & _
"FROM qry_apontamento qa INNER JOIN g2cursoeve gc ON qa.doc=gc.coddoc " & _
"WHERE (((IIf(([I])>0,'Faltas ',IIf(([Repos])>0,'Reposição ',IIf(([atraso])>0,'Atrasos ',IIf(([JD])>0,'Faltas Just ',IIf(([JA])>0,'Faltas Abon ',IIf(([DP])>0,'Dependentes ',IIf(([extra])>0,'A.Extras ',Null)))))))) Is Not Null) " & _
"AND ((qa.Mes_base)=#" & dtaccess(mesbase) & "#)) " & _
"GROUP BY qa.mes_base, gc.CODCCUSTO, gc.CURSO, qa.chapa, qa.NOME, IIf(([I])>0,'Faltas ',IIf(([Repos])>0,'Reposição ',IIf(([atraso])>0,'Atrasos ',IIf(([JD])>0,'Faltas Just ',IIf(([JA])>0,'Faltas Abon ',IIf(([DP])>0,'Dependentes ',IIf(([extra])>0,'A.Extras ',Null))))))) " & _
"ORDER BY gc.CURSO, qa.NOME, IIf(([I])>0,'Faltas ',IIf(([Repos])>0,'Reposição ',IIf(([atraso])>0,'Atrasos ',IIf(([JD])>0,'Faltas Just ',IIf(([JA])>0,'Faltas Abon ',IIf(([DP])>0,'Dependentes ',IIf(([extra])>0,'A.Extras ',Null))))))) " & _
"PIVOT IIf([dia_mes] Is Null,0,Day([dia_mes])) In (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,0);" 
sql="SELECT qa.mes_base, gc.CODCCUSTO AS sec, gc.CURSO, qa.chapa, qa.NOME, " & _
"grupam=case when i>0 then 'Faltas' else case when repos>0 then 'Reposição' else case when qa.atraso>0 then 'Atrasos' else case when JD>0 then 'Faltas Just ' else case when JD>0 then 'Faltas Abon ' else case when DP>0 then 'Dependentes ' else case when extra>0 then 'A.Extras ' else null end end end end end end end, " & _
"'01'=sum(case when day(dia_mes)=01 then case when selec=0 then null else selec end end), " & _
"'02'=sum(case when day(dia_mes)=02 then case when selec=0 then null else selec end end), " & _
"'03'=sum(case when day(dia_mes)=03 then case when selec=0 then null else selec end end), " & _
"'04'=sum(case when day(dia_mes)=04 then case when selec=0 then null else selec end end), " & _
"'05'=sum(case when day(dia_mes)=05 then case when selec=0 then null else selec end end), " & _
"'06'=sum(case when day(dia_mes)=06 then case when selec=0 then null else selec end end), " & _
"'07'=sum(case when day(dia_mes)=07 then case when selec=0 then null else selec end end), " & _
"'08'=sum(case when day(dia_mes)=08 then case when selec=0 then null else selec end end), " & _
"'09'=sum(case when day(dia_mes)=09 then case when selec=0 then null else selec end end), " & _
"'10'=sum(case when day(dia_mes)=10 then case when selec=0 then null else selec end end), " & _
"'11'=sum(case when day(dia_mes)=11 then case when selec=0 then null else selec end end), " & _
"'12'=sum(case when day(dia_mes)=12 then case when selec=0 then null else selec end end), " & _
"'13'=sum(case when day(dia_mes)=13 then case when selec=0 then null else selec end end), " & _
"'14'=sum(case when day(dia_mes)=14 then case when selec=0 then null else selec end end), " & _
"'15'=sum(case when day(dia_mes)=15 then case when selec=0 then null else selec end end), " & _
"'16'=sum(case when day(dia_mes)=16 then case when selec=0 then null else selec end end), " & _
"'17'=sum(case when day(dia_mes)=17 then case when selec=0 then null else selec end end), " & _
"'18'=sum(case when day(dia_mes)=18 then case when selec=0 then null else selec end end), " & _
"'19'=sum(case when day(dia_mes)=19 then case when selec=0 then null else selec end end), " & _
"'20'=sum(case when day(dia_mes)=20 then case when selec=0 then null else selec end end), " & _
"'21'=sum(case when day(dia_mes)=21 then case when selec=0 then null else selec end end), " & _
"'22'=sum(case when day(dia_mes)=22 then case when selec=0 then null else selec end end), " & _
"'23'=sum(case when day(dia_mes)=23 then case when selec=0 then null else selec end end), " & _
"'24'=sum(case when day(dia_mes)=24 then case when selec=0 then null else selec end end), " & _
"'25'=sum(case when day(dia_mes)=25 then case when selec=0 then null else selec end end), " & _
"'26'=sum(case when day(dia_mes)=26 then case when selec=0 then null else selec end end), " & _
"'27'=sum(case when day(dia_mes)=27 then case when selec=0 then null else selec end end), " & _
"'28'=sum(case when day(dia_mes)=28 then case when selec=0 then null else selec end end), " & _
"'29'=sum(case when day(dia_mes)=29 then case when selec=0 then null else selec end end), " & _
"'30'=sum(case when day(dia_mes)=30 then case when selec=0 then null else selec end end), " & _
"'31'=sum(case when day(dia_mes)=31 then case when selec=0 then null else selec end end), " & _
"'00'=sum(case when dia_mes is null then case when selec=0 then null else selec end end), " & _
"Sum(case when selec=0 then null else selec end) AS Total1  " & _
"FROM qry_apontamento qa INNER JOIN g2cursoeve gc ON qa.doc=gc.coddoc  " & _
"WHERE case when i>0 then 'Faltas' else case when repos>0 then 'Reposição' else case when qa.atraso>0 then 'Atrasos' else case when JD>0 then 'Faltas Just ' else case when JD>0 then 'Faltas Abon ' else case when DP>0 then 'Dependentes ' else case when extra>0 then 'A.Extras ' else null end end end end end end end Is Not Null  " & _
"AND qa.Mes_base='" & dtaccess(mesbase) & "' " & _ 
"GROUP BY qa.mes_base, gc.CODCCUSTO, gc.CURSO, qa.chapa, qa.NOME, " & _
"case when i>0 then 'Faltas' else case when repos>0 then 'Reposição' else case when qa.atraso>0 then 'Atrasos' else case when JD>0 then 'Faltas Just ' else case when JD>0 then 'Faltas Abon ' else case when DP>0 then 'Dependentes ' else case when extra>0 then 'A.Extras ' else null end end end end end end end  " & _
"ORDER BY gc.CURSO, qa.NOME,  " & _
"case when i>0 then 'Faltas' else case when repos>0 then 'Reposição' else case when qa.atraso>0 then 'Atrasos' else case when JD>0 then 'Faltas Just ' else case when JD>0 then 'Faltas Abon ' else case when DP>0 then 'Dependentes ' else case when extra>0 then 'A.Extras ' else null end end end end end end end  "
	'response.write "<br>" & sql
end if
%>

<% if request.form="" then %>
<p class=titulo>Geração de Capas para Apontamento de Ponto Docente
<form method="POST" action="rpt_capa.asp">
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
linha=1:inicio=0
rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof
if linha>46 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<br>"
	response.write "<p style='margin-top: 0; margin-bottom: 0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<table border='0' cellpadding='2' width='950' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  >Apontamento dos Docentes</td>"
	response.write "<td align='center'>Mês-Base: " & mesbase & "</td>"
	response.write "<td align='right' >Recursos Humanos</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<b>" & rs("sec") & " - <font size=2>" & rs("curso") & "</font></b>"

	response.write "<table border='0' cellpadding='1' width='950' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=titulor style='border: 1px solid #000000' >Chapa   </td>"
	response.write "<td class=titulor style='border: 1px solid #000000' >Professor   </td>"
	response.write "<td class=titulor style='border: 1px solid #000000' align=""center"">Evento </td>"
	response.write "<td class=titulor style='border: 1px solid #000000' align=""center"">Tot</td>"
	for a=1 to 31
	response.write "<td class=titulor style='border: 1px solid #000000' align=""center"">" & numzero(a,2) & "</td>"
	next
	response.write "<td class=titulor style='border: 1px solid #000000' align=""center"">0</td>"
	response.write "</tr>"
	linha=3
end if

if lastsecao<>rs("sec") then
	if inicio=1 then
		pagina=pagina+1
		response.write "</table>"
		response.write "<br>"
		response.write "<p style='margin-top: 0; margin-bottom: 0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
		response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	end if
	response.write "<table border='0' cellpadding='2' width='950' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  >Apontamento dos Docentes</td>"
	response.write "<td align='center'>Mês-Base: " & mesbase & "</td>"
	response.write "<td align='right' >Recursos Humanos</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<b>" & rs("sec") & " - <font size=2>" & rs("curso") & "</font></b>"

	response.write "<table border='0' cellpadding='1' width='950' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=titulor style='border: 1px solid #000000' >Chapa   </td>"
	response.write "<td class=titulor style='border: 1px solid #000000' >Professor   </td>"
	response.write "<td class=titulor style='border: 1px solid #000000' align=""center"">Evento </td>"
	response.write "<td class=titulor style='border: 1px solid #000000' align=""center"">Tot</td>"
	for a=1 to 31
	response.write "<td class=titulor style='border: 1px solid #000000' align=""center"">" & numzero(a,2) & "</td>"
	next
	response.write "<td class=titulor style='border: 1px solid #000000' align=""center"">0</td>"
	response.write "</tr>"
	linha=3
end if 'lastsecao

%>
  <tr>
    <td class="campor" style="border: 1px solid #000000"><%=rs("chapa")%>&nbsp;</td>
    <td class="campor" style="border: 1px solid #000000"><%=rs("nome")%></td>
    <td class="campor" style="border: 1px solid #000000"><%=rs("grupam")%></td>
    <td class="campor" style="border: 1px solid #000000" align="right"><%=rs("total1")%>&nbsp;</td>
<%
for a=6 to 6+31
%>
    <td class="campor" style="border: 1px solid #000000" align="center"><%=rs.fields(a)%></td>
<%
next
%>

  </tr>
<%
lastsecao=rs("sec")
linha=linha+1
inicio=1
rs.movenext
loop
rs.close

%>
</table>
<%
linha=linha+1
if linha>46 then
	pagina=pagina+1
	response.write "<br>"
	response.write "<p style='margin-top: 0; margin-bottom: 0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<table border='0' cellpadding='2' width='950' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  >Apontamento dos Docentes</td>"
	response.write "<td align='center'>Mês-Base: " & mesbase & "</td>"
	response.write "<td align='right' >Recursos Humanos</td>"
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