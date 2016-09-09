<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 1200
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a38")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cálculo de Enquadramento Salarial - Professores</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set conexao2=server.createobject ("ADODB.Connection")
conexao2.Open application("consql")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao2
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao2
set rsn=server.createobject ("ADODB.Recordset")
Set rsn.ActiveConnection = conexao

mes=request.form("mes")
mes=4
ano=request.form("ano")
ano=2005
udia=day(dateserial(ano,mes+1,1)-1)
sqln="select top 20 chapa from pfunc where codsindicato='03' and codsituacao in ('A','F') " & _
" and chapa in (select chapa from quem_nomeacoes) " & _
"order by chapa "
sqln="select top 10 chapa from pfunc where codsindicato='03' and codsituacao in ('A','F') " & _
" and chapa>'00000' " & _
"order by chapa "
sqln="select top 100 chapa from pfunc where codsindicato='03' and codsituacao in ('A','F','Z') " & _
" and chapa in (select chapa from grades_rt where codevento in ('128','255','256','257','258','034') ) " & _
"order by chapa "
sqln="select top 1000 chapa from cpcd group by chapa "

rsn.Open sqln, ,adOpenStatic, adLockReadOnly
rsn.movefirst
do while not rsn.eof
chapa=numzero(request.form("chapa"),5)
chapa=rsn("chapa")
taes=0:taem=0:tsaldo=0:saldoa=0
	
sqld="select f.nome, c.nome as funcao, s.descricao as setor from pfunc f, psecao s, pfuncao c where f.codsecao=s.codigo and f.codfuncao=c.codigo and f.chapa='" & chapa & "'"
rsd.Open sqld, ,adOpenStatic, adLockReadOnly
nome=rsd("nome"):setor=rsd("setor"):funcao=rsd("funcao")
rsd.close
linha=0
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=campo colspan=3>RECLASSIFICAÇÃO CPCD</td>
<td class=campo align="right"><%=rsn.absoluteposition%></td>
</tr>
<tr><td class="campor">Chapa</td>
	<td class="campor">Nome</td>
	<td class="campor">Setor</td>
	<td class="campor">Função</td></tr>
<tr><td class="campor"><%=chapa%></td>
	<td class="campor"><b><%=nome%></b></td>
	<td class="campor"><%=Setor%></td>
	<td class="campor"><%=funcao%></td></tr>
</table>
<table border="1" bordercolor="gray" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulor align="center" rowspan=2>Mês Ref.</td>
	<td class=titulor align="center" rowspan=2>Descrição</td>
	<td class=titulor align="center" rowspan=2>Data Pagto</td>
	<td class=titulor align="center" rowspan=2>Valor Pago</td>
	<td class=titulor align="center" colspan=2>Valor Aula</td>
	<td class=titulor align="center" rowspan=2>Diferença</td>
	<td class=titulor align="center" rowspan=2>Índice</td>
	<td class=titulor align="center" rowspan=2>Juros</td>
	<td class=titulor align="center" rowspan=2>Total</td>
</tr>
<tr>
	<td class=titulor align="center">Pago</td>
	<td class=titulor align="center">Devido</td>
</tr>

<%
linha=5
total=0
sqlcr="select [ano/mes] as competencia, eve, descricao, dtpagto, ref, valor, vh_epoca, vh_tit, total1, indice, meses, total4 " & _
"from cpcd where chapa='" & chapa & "' order by [ano/mes], vd desc, eve, dtpagto "
'marcações do chronus
rs.Open sqlcr, ,adOpenStatic, adLockReadOnly
inicio=0
rs.movefirst
do while not rs.eof
if lastcomp<>rs("competencia") and inicio=1 then
	response.write "<tr><td class=titulor colspan=3>Totais do mês</td>"
	response.write "<td class=titulor align="right">" & formatnumber(totalpago,2) & "&nbsp;</td>"
	response.write "<td class=titulor colspan=2>&nbsp;</td>"
	response.write "<td class=titulor align="right">" & formatnumber(totaldif,2) & "&nbsp;</td>"
	response.write "<td class=titulor colspan=2>&nbsp;</td>"
	response.write "<td class=titulor align="right">" & formatnumber(totalpag,2) & "&nbsp;</td>"
	linha=linha+1:totalpago=0:totaldif=0:totalpag=0
end if
if linha>70 then
	response.write "</table>"
	response.write "<DIV style=""page-break-after:always""></DIV>"
	response.write "<table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='650'>"
	response.write "<tr><td class=campo colspan=3>RECLASSIFICAÇÃO CPCD</td>"
	response.write "<td class=campo align="right">" & rsn.absoluteposition & "</td>"
	response.write "</tr>"
	response.write "<tr><td class="campor">Chapa</td>"
	response.write "	<td class="campor">Nome</td>"
	response.write "	<td class="campor">Setor</td>"
	response.write "	<td class="campor">Função</td></tr>"
	response.write "<tr><td class="campor">" & chapa & "</td>"
	response.write "	<td class="campor"><b>" & nome & "</b></td>"
	response.write "	<td class="campor">" & setor & "</td>"
	response.write "	<td class="campor">" & funcao & "</td></tr>"
	response.write "</table>"
	response.write "<table border='1' bordercolor='gray' cellpadding='1' cellspacing='0' style='border-collapse: collapse' width='650'>"
	response.write "<tr>"
	response.write "	<td class=titulor align="center" rowspan=2>Mês Ref.</td>"
	response.write "	<td class=titulor align="center" rowspan=2>Descrição</td>"
	response.write "	<td class=titulor align="center" rowspan=2>Data Pagto</td>"
	response.write "	<td class=titulor align="center" rowspan=2>Valor Pago</td>"
	response.write "	<td class=titulor align="center" colspan=2>Valor Aula</td>"
	response.write "	<td class=titulor align="center" rowspan=2>Diferença</td>"
	response.write "	<td class=titulor align="center" rowspan=2>Índice</td>"
	response.write "	<td class=titulor align="center" rowspan=2>Juros</td>"
	response.write "	<td class=titulor align="center" rowspan=2>Total</td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "	<td class=titulor align="center">Pago</td>"
	response.write "	<td class=titulor align="center">Devido</td>"
	response.write "</tr>"
	linha=5
end if
%>
<tr>
	<td class="campor" align="center"><%=rs("competencia")%></td>
	<td class="campor" align="left"><%=rs("descricao")%></td>
	<td class="campor" align="center"><%=rs("dtpagto")%></td>
	<td class="campor" align="right"><%=formatnumber(rs("valor"),2)%>&nbsp;</td>
	<td class="campor" align="right"><%=formatnumber(rs("vh_epoca"),2)%>&nbsp;</td>
	<td class="campor" align="right"><%=formatnumber(rs("vh_tit"),2)%>&nbsp;</td>
	<td class="campor" align="right"><%=formatnumber(rs("total1"),2)%>&nbsp;</td>
	<td class="campor" align="right"><%=formatnumber(rs("indice"),6)%>&nbsp;</td>
	<td class="campor" align="right"><%=formatpercent(rs("meses")/100,0)%>&nbsp;</td>
	<td class="campor" align="right"><%=formatnumber(rs("total4"),2)%>&nbsp;</td>
</tr>
<%
inicio=1
linha=linha+1
totalpago=totalpago+rs("valor")
totaldif=totaldif+rs("total1")
totalpag=totalpag+rs("total4")
ttotalpago=ttotalpago+rs("valor")
ttotaldif=ttotaldif+rs("total1")
ttotalpag=ttotalpag+rs("total4")
lastcomp=rs("competencia")
rs.movenext
loop
response.write "<tr><td class=titulor colspan=3>Totais do mês</td>"
response.write "<td class=titulor align="right">" & formatnumber(totalpago,2) & "&nbsp;</td>"
response.write "<td class=titulor colspan=2>&nbsp;</td>"
response.write "<td class=titulor align="right">" & formatnumber(totaldif,2) & "&nbsp;</td>"
response.write "<td class=titulor colspan=2>&nbsp;</td>"
response.write "<td class=titulor align="right">" & formatnumber(totalpag,2) & "&nbsp;</td>"
linha=linha+1:totalpago=0:totaldif=0:totalpag=0

response.write "<tr><td class=titulor colspan=6>TOTAL GERAL</td>"
response.write "<td class=titulor align="right">" & formatnumber(ttotaldif,2) & "&nbsp;</td>"
response.write "<td class=titulor colspan=2>&nbsp;</td>"
response.write "<td class=titulor align="right">" & formatnumber(ttotalpag,2) & "&nbsp;</td>"

rs.close
%>
</table>
<%
if ttotalpag<100 then 
	parcelas=1 
	vparc=ttotalpag
	texto=" 1 parcela de " & formatnumber(vparc,2)
else 
	parcelas=3
	vparc=int((ttotalpag/parcelas)*100)/100
	vparc1=ttotalpag-(vparc*2)
	texto=" 3 parcelas de " & formatnumber(vparc,2) & " / " & formatnumber(vparc,2) & " / " & formatnumber(vparc1,2)
end if
%>
<p style="margin-top: 0; margin-bottom: 0"><font size=1><b>Atualizado de acordo com a Tabela para Atualização de Débitos Trabalhistas - Junho/2005 - TRT 2ª Região/SP.
</font></p>
<br>
<p style="margin-top: 0; margin-bottom: 0"><font size=2><b>Valor dividido em <%=texto%></b></font></p>

<DIV style="page-break-after:always"></DIV>
<%
rsn.movenext
ttotalpago=0:ttotaldif=0:ttotalpag=0
loop
rsn.close

set rs=nothing
set rs2=nothing
set rsd=nothing
set rsn=nothing
conexao.close
set conexao=nothing
conexao2.close
set conexao2=nothing
%>
</body>
</html>