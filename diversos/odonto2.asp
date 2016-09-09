<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Pesquisa de interesse para adesão de Plano Odontológico</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

sql0="select perg1 as total from tab_odonto"
rs.Open sql0, ,adOpenStatic, adLockReadOnly
devolvido=rs.recordcount
rs.close
%>
<div align="left">

<!-- inicio formulario -->
<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td height=30 align="left" class="campop" style="border-bottom:solid 2 #000000">
	<b>Pesquisa de interesse para adesão de Plano Odontológico</td>
	<td align="right" valign=middle style="border-bottom:solid 2 #000000">
	<img src="../images/logo_centro_universitario_unifieo_big.jpg" width="150" border="0" alt="">
	</td>
</tr>
<tr><td height=15></td></tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop"><p style="margin-top:2;margin-bottom:2;line-height:15px;text-align:justify"><b>
	<font color=blue>Formulários distribuídos: aproximadamente 330<br>
	<font color=black>Formulários devolvidos: <font color=red><%=devolvido%> <font color=green>(<%=formatpercent(devolvido/330,2)%>)
	</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop"><p style="margin-top:6;margin-bottom:5;line-height:20px;text-align:justify">
	<B>1 - Alguma vez você recebeu orientação sobre a importância da odontologia para a sua saúde geral e sobre
	Prevenção em Odontologia?
	</td>
</tr>
<tr>
	<td class="campop">
<!-- tabulação -->
<table style="border-collapse: collapse" border="1" cellpadding="3" cellspacing="0">
<tr><td class=titulo align="center">Resposta</td><td class=titulo align="center">Total</t><td class=titulo align="center">%</td></tr>
<%
	sql1="SELECT perg1, Count(perg1) AS freq FROM tab_odonto GROUP BY perg1 ORDER BY Count(perg1) DESC;"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof
	if rs("perg1")="-" then texto="Não responderam"
	if rs("perg1")="N" then texto="NÃO"
	if rs("perg1")="S" then texto="SIM"
%>
<tr>
	<td class=campo><%=texto%></td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/devolvido,2)%></td>
</tr>
<%
	rs.movenext
	loop
	rs.close
%>
</table>
<!-- tabulação -->
	</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop"><p style="margin-top:6;margin-bottom:5;line-height:20px;text-align:justify">
	<B>2 - Você já tem plano odontológico?
	</td>
</tr>
<tr>
	<td class="campop">
<!-- tabulação -->
<table style="border-collapse: collapse" border="1" cellpadding="3" cellspacing="0">
<tr><td class=titulo align="center">Resposta</td><td class=titulo align="center">Total</t><td class=titulo align="center">%</td></tr>
<%
	sql1="SELECT perg2, Count(perg2) AS freq FROM tab_odonto GROUP BY perg2 ORDER BY Count(perg2) DESC;"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof
	if rs("perg2")="-" then texto="Não responderam"
	if rs("perg2")="N" then texto="NÃO"
	if rs("perg2")="S" then texto="SIM"
%>
<tr>
	<td class=campo><%=texto%></td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/devolvido,2)%></td>
</tr>
<%
	rs.movenext
	loop
	rs.close
%>
</table>
<!-- tabulação -->
	</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop"><p style="margin-top:6;margin-bottom:5;line-height:20px;text-align:justify">
	<B>3 - Caso o UNIFIEO oferecesse um Plano Coletivo Odontológico para você e seus dependentes, com ampla cobertura para
	restaurações, tratamentos de canal, tratamentos gengivais, emergência 24 horas, odontopediatria, prevenção, entre
	outros em consultórios e clínicas particulares, você teria interesse em participar?
	</td>
</tr>
<tr>
	<td class="campop">
<!-- tabulação -->
<table style="border-collapse: collapse" border="1" cellpadding="3" cellspacing="0">
<tr><td class=titulo align="center">Resposta</td><td class=titulo align="center">Total</t><td class=titulo align="center">%</td></tr>
<%
	sql1="SELECT perg3, Count(perg3) AS freq FROM tab_odonto GROUP BY perg3 ORDER BY Count(perg3) DESC;"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof
	if rs("perg3")="-" then texto="Não responderam"
	if rs("perg3")="N" then texto="NÃO"
	if rs("perg3")="S" then texto="SIM"
%>
<tr>
	<td class=campo><%=texto%></td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/devolvido,2)%></td>
</tr>
<%
	rs.movenext
	loop
	rs.close
%>
</table>
<!-- tabulação -->
	</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop"><p style="margin-top:6;margin-bottom:5;line-height:20px;text-align:justify">
	<B>4 - Pelo serviço oferecido você aceitaria pagar entre R$ 10,00 e R$ 12,00 mensais por pessoa a ser descontado
	no seu holerite?
	</td>
</tr>
<tr>
	<td class="campop">
	
<table><tr><td>
<!-- tabulação -->
<table style="border-collapse: collapse" border="1" cellpadding="3" cellspacing="0">
<tr><td class=titulo align="center" colspan=3>Total de respostas</td></tr>
<tr><td class=titulo align="center">Resposta</td><td class=titulo align="center">Total</t><td class=titulo align="center">%</td></tr>
<%
	sql1="SELECT perg4, Count(perg4) AS freq FROM tab_odonto GROUP BY perg4 ORDER BY Count(perg4) DESC;"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof
	if rs("perg4")="-" then texto="Não responderam"
	if rs("perg4")="N" then texto="NÃO"
	if rs("perg4")="S" then texto="SIM"
%>
<tr>
	<td class=campo><%=texto%></td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/devolvido,2)%></td>
</tr>
<%
	rs.movenext
	loop
	rs.close
%>
</table>
<!-- tabulação -->
</td><td>
<!-- tabulação -->
<table style="border-collapse: collapse" border="1" cellpadding="3" cellspacing="0">
<tr><td class=titulo align="center" colspan=3>Respostas Sim no item 3</td></tr>
<tr><td class=titulo align="center">Resposta</td><td class=titulo align="center">Total</t><td class=titulo align="center">%</td></tr>
<%
	sql0="select count(perg3) as total3 from tab_odonto where perg3='S' "
	rs.Open sql0, ,adOpenStatic, adLockReadOnly:total3=rs("total3"):rs.close
	sql0="select count(perg3) as total3 from tab_odonto where perg3='N' "
	rs.Open sql0, ,adOpenStatic, adLockReadOnly:total3n=rs("total3"):rs.close
	sql1="SELECT perg4, Count(perg4) AS freq FROM tab_odonto where perg3='S' GROUP BY perg4 ORDER BY Count(perg4) DESC;"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof
	if rs("perg4")="-" then texto="Não responderam"
	if rs("perg4")="N" then texto="NÃO"
	if rs("perg4")="S" then texto="SIM"
%>
<tr>
	<td class=campo><%=texto%></td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/total3,2)%></td>
</tr>
<%
	rs.movenext
	loop
	rs.close
%>
</table>
<!-- tabulação -->
</td></tr></table>

	</td>
</tr>
</table>

<table style="border-collapse: collapse" border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td class="campop" colspan=3><p style="margin-top:6;margin-bottom:5;line-height:20px;text-align:justify">
	<B>5 - O que motivaria a sua participação em um plano odontológico?
	</td>
</tr>
<tr>
	<td class="campop">
	
<table><tr><td>
<!-- tabulação -->
<table style="border-collapse: collapse" border="1" cellpadding="3" cellspacing="0">
<tr><td class=titulo align="center" colspan=3>Total dos "SIM"</td></tr>
<tr><td class=titulo align="center">Resposta</td><td class=titulo align="center">Total</t><td class=titulo align="center">%</td></tr>
<%
	sql1="SELECT Mid([perg5],1,1) AS opcao, Count(Mid([perg5],1,1)) AS freq FROM tab_odonto " & _
	"WHERE perg3='S' GROUP BY Mid([perg5],1,1) HAVING Mid([perg5],1,1)='X';"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst:do while not rs.eof
%>
<tr>
	<td class=campo>Preço</td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/total3,2)%></td>
</tr>
<%
	rs.movenext:loop:rs.close
%>
<%
	sql1="SELECT Mid([perg5],2,1) AS opcao, Count(Mid([perg5],2,1)) AS freq FROM tab_odonto " & _
	"WHERE perg3='S' GROUP BY Mid([perg5],2,1) HAVING Mid([perg5],2,1)='X';"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst:do while not rs.eof
%>
<tr>
	<td class=campo>Rede Credenciada Ampla</td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/total3,2)%></td>
</tr>
<%
	rs.movenext:loop:rs.close
%>
<%
	sql1="SELECT Mid([perg5],3,1) AS opcao, Count(Mid([perg5],3,1)) AS freq FROM tab_odonto " & _
	"WHERE perg3='S' GROUP BY Mid([perg5],3,1) HAVING Mid([perg5],3,1)='X';"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst:do while not rs.eof
%>
<tr>
	<td class=campo>Cobertura Ampla</td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/total3,2)%></td>
</tr>
<%
	rs.movenext:loop:rs.close
%>
<%
	sql1="SELECT Mid([perg5],4,1) AS opcao, Count(Mid([perg5],4,1)) AS freq FROM tab_odonto " & _
	"WHERE perg3='S' GROUP BY Mid([perg5],4,1) HAVING Mid([perg5],4,1)='X';"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst:do while not rs.eof
%>
<tr>
	<td class=campo>Outros</td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/total3,2)%></td>
</tr>
<%
	rs.movenext:loop:rs.close
%>

</table>
<!-- tabulação -->
</td><td>
<!-- tabulação -->
<table style="border-collapse: collapse" border="1" cellpadding="3" cellspacing="0">
<tr><td class=titulo align="center" colspan=3>Respostas dos "NÃO"</td></tr>
<tr><td class=titulo align="center">Resposta</td><td class=titulo align="center">Total</t><td class=titulo align="center">%</td></tr>
<%
	sql1="SELECT Mid([perg5],1,1) AS opcao, Count(Mid([perg5],1,1)) AS freq FROM tab_odonto " & _
	"WHERE perg3='N' GROUP BY Mid([perg5],1,1) HAVING Mid([perg5],1,1)='X';"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst:do while not rs.eof
%>
<tr>
	<td class=campo>Preço</td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/total3n,2)%></td>
</tr>
<%
	rs.movenext:loop:rs.close
%>
<%
	sql1="SELECT Mid([perg5],2,1) AS opcao, Count(Mid([perg5],2,1)) AS freq FROM tab_odonto " & _
	"WHERE perg3='N' GROUP BY Mid([perg5],2,1) HAVING Mid([perg5],2,1)='X';"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst:do while not rs.eof
%>
<tr>
	<td class=campo>Rede Credenciada Ampla</td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/total3n,2)%></td>
</tr>
<%
	rs.movenext:loop:rs.close
%>
<%
	sql1="SELECT Mid([perg5],3,1) AS opcao, Count(Mid([perg5],3,1)) AS freq FROM tab_odonto " & _
	"WHERE perg3='N' GROUP BY Mid([perg5],3,1) HAVING Mid([perg5],3,1)='X';"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst:do while not rs.eof
%>
<tr>
	<td class=campo>Cobertura Ampla</td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/total3N,2)%></td>
</tr>
<%
	rs.movenext:loop:rs.close
%>
<%
	sql1="SELECT Mid([perg5],4,1) AS opcao, Count(Mid([perg5],4,1)) AS freq FROM tab_odonto " & _
	"WHERE perg3='N' GROUP BY Mid([perg5],4,1) HAVING Mid([perg5],4,1)='X';"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst:do while not rs.eof
%>
<tr>
	<td class=campo>Outros</td>
	<td class=campo align="center"><%=rs("freq")%></td>
	<td class=campo align="center"><%=formatpercent(rs("freq")/total3N,2)%></td>
</tr>
<%
	rs.movenext:loop:rs.close
%>

</table>
<!-- tabulação -->
</td></tr></table>
	
	</td>
</tr>
</table>
<!-- celula fim para definir tamanho -->	
	
</div>
</body>
</html>