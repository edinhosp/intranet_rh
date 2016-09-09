<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a77")="N" or session("a77")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Checagem - Cargos e Salários</title>
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
	set rs2=server.createobject ("ADODB.Recordset")
	Set rs2.ActiveConnection = conexao
	set rs3=server.createobject ("ADODB.Recordset")
	Set rs3.ActiveConnection = conexao
	data=dateserial(2016,6,1)
	ttf=0:ttd=0:ttb=0:tt9=0:tqf=0:tqd=0:tqb=0:tq9=0
%>
<p style="margin-top:0;margin-bottom:0;color:Blue;font-size:9pt;text-align:left">
<b>Análise da Tabela Docentes - <%=monthname(month(data)) & "/" & year(data)%><br>&nbsp;</font></p>
<!-- inicio funcionarios -->
<%
sql="select chapa, nome, dataadmissao, codfuncao, funcao, codsecao, secao, titulacao=tab_instr, codnivelsal, ge, grauinstrucao, instrucao, tab_ref, tab_grade, titulacaopaga " & _
"from dc_professor f " & _
"where F.CODSITUACAO in ('A','F','Z','E') AND F.CODSINDICATO='03' and codtipo='N' " & _
"/*and f.chapa in (select chapa collate database_default from corporerm.dbo.pfsalcmp where codevento<>'027') */" & _
"ORDER BY F.CHAPA "

rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof
titulacao=rs("titulacao")
nivel=rs("codnivelsal")
reformulacao=rs("tab_ref")
%>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=650>
<tr><td class=titulor align="center">Chapa</td>
	<td class=titulor align="center">Nome</td>
	<td class=titulor align="center">Função</td>
	<td class=titulor align="center">Seção</td>
	<td class=titulor align="center">Titulação</td>
	<td class=titulor align="center">Nivel</td>
	<td class=titulor align="center">Admissao</td>
</tr>
<tr><td class="campoa"r><%=rs("chapa")%></td>
	<td class="campoa"r><%=rs("nome")%></td>
	<td class="campoa"r><%=rs("funcao")%>&nbsp;</td>
	<td class="campoa"r><%=rs("secao")%>&nbsp;</td>
	<td class="campoa"r><%=rs("grauinstrucao") & "-" & rs("instrucao")%></td>
	<td class="campoa"r><%=rs("codnivelsal")%>&nbsp;<%=titulacao%></td>
	<td class="campoa"r><%=rs("dataadmissao")%></td>
</tr>
</table>
<!-- inicio sal.composto -->
<%
sql3="SELECT s.CHAPA, s.CODEVENTO, e.DESCRICAO, s.NROSALARIO, s.JORNADA, s.VALOR, [JORNADA]/60.00 AS HSMES, [JORNADA]/60.00/4.50 AS HSSEM, round(VALOR/(JORNADA/60.00),2) AS HORA, s.INICIOVIGENCIA, s.FIMVIGENCIA " & _
"FROM corporerm.dbo.PFSALCMP s INNER JOIN corporerm.dbo.PEVENTO e ON s.CODEVENTO=e.CODIGO " & _
"WHERE s.CHAPA='" & rs("chapa") & "' /*and s.codevento<>'027'*/ " & _
"ORDER BY s.CHAPA, s.NROSALARIO "
rs3.Open sql3, ,adOpenStatic, adLockReadOnly

if rs3.recordcount>0 then
%>
<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=650 >
<tr>
	<td class=titulor align="center" width=15>#</td>
	<td class=titulor align="center" width=35>Cod.</td>
	<td class=titulor align="center" width=243>Descrição</td>
	<td class=titulor align="center" width=50>Valor</td>
	<td class=titulor align="center" width=40>Hs.Mês</td>
	<td class=titulor align="center" width=40>Vr.Hora</td>
	<td class=grupor align="center" width=2></td>
	<td class=titulor align="center" width=120 nowrap>Titulação</td>
	<td class=titulor align="center" width=30>Faixa</td>
	<td class=titulor align="center" width=40>Vr.Hora</td>
	<td class=titulor align="center" width=35>Status</td>
</tr>
<%
mensagem=""
tsalario=0:thoras=0
rs3.movefirst
do while not rs3.eof
tsalario=tsalario+cdbl(rs3("valor"))
thoras=thoras+cdbl(rs3("hsmes"))
ttsalario=ttsalario+cdbl(rs3("valor"))
tthoras=tthoras+cdbl(rs3("hsmes"))
%>
<tr>
	<td class="campor" align="center"><%=rs3("nrosalario")%></td>
	<td class="campor" align="center"><%=rs3("codevento")%></td>
	<td class="campor"><%=rs3("descricao")%></td>
	<td class="campor" align="right"><%=formatnumber(rs3("valor"),2)%>&nbsp;</td>
	<td class="campor" align="right"><%=rs3("hsmes")%>&nbsp;</td>
	<td class="campor" align="right"><%=formatnumber(rs3("hora"),2)%>&nbsp;</td>
	<td class=grupor></td>
<%
sql2="SELECT csd_cursos.evento, csd_cursos.tabela, csd_faixas.dt_faixa, csd_titulos.titulacao, csd_titulos.nivel, csd_titulos.titulo, csd_titulos.faixasalarial, csd_faixas.valoraula " & _
"FROM (csd_cursos INNER JOIN csd_titulos ON csd_cursos.tabela = csd_titulos.tabela) INNER JOIN csd_faixas ON csd_titulos.faixasalarial = csd_faixas.faixasalarial " & _
"WHERE csd_cursos.evento='" & rs3("codevento") & "' " & _
"AND '" & dtaccess(data) & "' Between [ivigencia] And [fvigencia] " & _
"AND csd_faixas.dt_faixa='" & dtaccess(data) & "' " & _
"AND csd_titulos.titulacao='" & titulacao & "' AND csd_titulos.nivel='" & nivel & "' and reformulacao='" & reformulacao & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
titulo=rs2("titulo")
faixa=rs2("faixasalarial")
valor=cdbl(rs2("valoraula"))
else
titulo=""
faixa=""
valor=0
end if 'recordcount rs2
rs2.close
if valor>0 and cdbl(rs3("hora"))>0 then 
	diferenca=valor-cdbl(rs3("hora")) 
else 
	diferenca=0
end if

if cdbl(diferenca)<>0 then totaldif=cdbl(totaldif) + cdbl(diferenca)
if cdbl(diferenca)<>0 then diferenca=">" & formatnumber(cdbl(diferenca),2) else diferenca=""
if diferenca<>"" then mensagem=mensagem & "<br>" & rs("chapa") & " - " & rs3("codevento") & " : " & cstr(rs3("hora")) & " -> " & cstr(valor)
if diferenca<>"" then fundo="grupor" else fundo=campor
%>
	<td class="campor"><%=titulo%></td>
	<td class="campor"><%=faixa%></td>
	<td class="campor" align="right"><%=formatnumber(valor,2)%>&nbsp;</td>
	<td class=<%=fundo%> nowrap><%=diferenca%></td>
</tr>
<%
rs3.movenext
loop
end if 'recordcount rs3
rs3.close
tvhora=tsalario/thoras
ttvhora=ttsalario/tthoras
%>
<tr>
	<td class=titulor colspan=3></td>
	<td class="campor" align="right"><%=formatnumber(tsalario,2)%>&nbsp;</td>
	<td class="campor" align="right"><%=thoras%>&nbsp;</td>
	<td class="campor" align="right"><%=formatnumber(tvhora,2)%>&nbsp;</td>
	<td class=grupor></td>
	<td class=titulor colspan=4></td>
</tr>
</table>
<!-- fim sal.composto -->
&nbsp;
<%
if rs("titulacaopaga")="F" then ttf=ttf+tsalario:tqf=tqf+thoras
if rs("titulacaopaga")="D" then ttd=ttd+tsalario:tqd=tqd+thoras
if rs("titulacaopaga")="B" then ttb=ttb+tsalario:tqb=tqb+thoras
if rs("titulacaopaga")="9" then tt9=tt9+tsalario:tq9=tq9+thoras

rs.movenext
loop
rs.close
%>
<!-- fim funcionarios -->
<br>Total Salários: <%=formatnumber(ttsalario,2)%>
<br>Total Horas: <%=tthoras%>
<br>Hora Média: <%=formatnumber(ttvhora,2)%>
<br>><%=formatnumber(cdbl(totaldif),2)%>
<%
'a1=ttf/tqf: response.write "Doutor: " & a1 & "<br>"
'a2=ttd/tqd: response.write "Mestre: " & a2 & "<br>"
'a3=ttb/tqb: response.write "Especialista: " & a3 & "<br>"
'a4=tt9/tq9: response.write "Graduado: " & a4 & "<br>"
response.write mensagem
%>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>