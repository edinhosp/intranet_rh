<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a70")="N" or session("a70")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Protocolo de Cesta Básica</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
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
sessao=session.sessionid

if request("sessao")<>"" then
sql="select opcao, t.sessao, substring(registro,47,4)+'/'+substring(registro,51,2)+'/'+substring(registro,53,2) as datapedido, " & _
"substring(registro,55,4)+'/'+substring(registro,59,2)+'/'+substring(registro,61,2) as dataliberacao, t2.horapedido, t3.valorpedido " & _
"from ttcbasica t, (select sessao, substring(registro,33,8) as horapedido from ttcbasica where campo1 in ('01')) as t2,  " & _
"(select sessao, substring(registro,14,14) as valorpedido from ttcbasica where campo1 in ('05')) as t3 " & _
"where t.campo1 in ('02') and t.sessao=t2.sessao and t.sessao=t3.sessao and t.sessao='" & request("sessao") & "' " & _
"order by substring(registro,47,4)+'/'+substring(registro,51,2)+'/'+substring(registro,53,2) "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	valorpedido=formatnumber(rs("valorpedido")/100,2)
	datapedido=cdate(rs("datapedido"))
end if
rs.close
end if

if request("acao")="excluir" then
	s=Request.QueryString("s")
	t=Request.QueryString("t")
	sql="delete from ttcbasica where sessao='" & s & "' and opcao='" & t & "'"
	conexao.execute sql
	manutencaocb=1
end if

%>

<% if request("sessao")="" then %>
<p class=titulo style="margin-top:0;margin-bottom:0">Geração de Protocolo para de Cesta Básica<p>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=450>
<tr>
	<td class=titulop align="center">Controle</td>
	<td class=titulop align="center">Tp</td>
	<td class=titulop align="center">Data Pedido</td>
	<td class=titulop align="center">Hora Pedido</td>
	<td class=titulop align="center">Data Liberação</td>
	<td class=titulop align="center">Valor Pedido</td>
    <td class=titulop align="center">&nbsp;</td>
</tr>
<%
sql="select top 100 opcao, t.sessao, convert(datetime,substring(registro,47,4)+'/'+substring(registro,51,2)+'/'+substring(registro,53,2)) as datapedido, " & _
"substring(registro,47,4) as ano, substring(registro,51,2) as mes, " & _
"convert(datetime,substring(registro,55,4)+'/'+substring(registro,59,2)+'/'+substring(registro,61,2)) as dataliberacao, t2.horapedido, t3.valorpedido " & _
"from ttcbasica t, (select sessao, substring(registro,33,8) as horapedido from ttcbasica where campo1 in ('01')) as t2,  " & _
"(select sessao, substring(registro,14,14) as valorpedido from ttcbasica where campo1 in ('05')) as t3 " & _
"where opcao in ('C','V') and t.campo1 in ('02') and t.sessao=t2.sessao and t.sessao=t3.sessao order by convert(datetime,substring(registro,47,4)+'/'+substring(registro,51,2)+'/'+substring(registro,53,2)) desc " 
rsc.Open sql, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
rsc.movefirst
do while not rsc.eof
sql2="select ano, mes, count(sessao) as freq from (" & sql & ") as t where mes='" & numzero(month(rsc("datapedido")),2) & "' and ano='" & year(rsc("datapedido")) & "' group by mes, ano, opcao "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
if rs("freq")>1 and cint(rs("ano"))=year(now) and cint(rs("mes"))>=month(now)-1 then existe=1 else existe=0
rs.close
%>
<tr>
	<td class="campop"><a href="cb_protocolo.asp?sessao=<%=rsc("sessao")%>&opcao=<%=rsc("opcao")%>"><%=rsc("sessao")%></a></td>
	<td class="campop" align="center"><%=rsc("opcao")%></td>
	<td class="campop" align="center"><%=rsc("datapedido")%></td>
	<td class="campop" align="center"><%=rsc("horapedido")%></td>
	<td class="campop" align="center"><%=rsc("dataliberacao")%></td>
	<td class="campop" align="right"><%=formatnumber(rsc("valorpedido")/100,2)%></td>
	<td class=campo align="center">&nbsp;
	<%if existe=1 then%>
		<a href="cb_protocolo.asp?acao=excluir&s=<%=rsc("sessao")%>&t=<%=rsc("opcao")%>">
		<img border="0" src="../images/Trash.gif"></a>
	<%end if%>
	</td>
</tr>
<%
rsc.movenext
loop
else
	response.write "<tr><td class=campo colspan=5>Não existem movimentos</td></tr>"
end if
rsc.close
%>
</table>
<%
else

sql="select t.sessao, convert(datetime,substring(registro,47,4)+'/'+substring(registro,51,2)+'/'+substring(registro,53,2)) as datapedido, " & _
"substring(registro,47,4) as ano, substring(registro,51,2) as mes, " & _
"convert(datetime,substring(registro,55,4)+'/'+substring(registro,59,2)+'/'+substring(registro,61,2)) as dataliberacao " & _
"from ttcbasica t " & _
"where t.campo1 in ('02') and t.sessao='" & request("sessao") & "' and opcao='" & request("opcao") & "' " & _
"order by convert(datetime,substring(registro,47,4)+'/'+substring(registro,51,2)+'/'+substring(registro,53,2)) " 
rs.Open sql, ,adOpenStatic, adLockReadOnly
dataliberacao=rs("dataliberacao")
rs.close
if request("opcao")="V" then textocab="TICKET-ALIMENTAÇÃO" else textocab="CESTA BÁSICA"

linha=5:inicio=0
sql="SELECT Left(campo2,2) as codcampus, codsecao=case when substring(registro,70,26)='JD.WILSON' then '04.0.000' else Right(campo2,8) end, campo3 as chapa, substring(registro,70,26) as Campus, " & _
"substring(registro,6,26) as Setor, substring(registro,112,30) as Nome FROM ttcbasica " & _
"WHERE sessao='" & request("sessao") & "' and opcao='" & request("opcao") & "' AND campo1='04' " & _
"ORDER BY substring(registro,112,30), substring(registro,70,26), substring(registro,6,26)"
'response.write sql
rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof
if linha>37 then
	inicio=0
	pagina=pagina+1
	if inicio=0 then
		response.write "</table>"
		response.write "<DIV style=""page-break-after:always""></DIV>"
	end if
	linha=5
end if
if inicio=0 then
%>
<table border="0" cellpadding="2" width="650" cellspacing="0" style="border-collapse: collapse">
<tr><td colspan=3 class=titulop align="center">RELAÇÃO DE FUNCIONÁRIOS QUE RECEBEM <%=textocab%></td></tr>
<tr>
    <td class=titulo align="left"  >Data Liberação: <%=dataliberacao%></td>
    <td class=titulo align="center">&nbsp;</td>
    <td class=titulo align="right" >Recursos Humanos</td>
</tr>
<tr>
    <td class=campo align="left" colspan=2 style="border:1px solid #000000">Para quem recebe remuneração de até 4 pisos salariais.</td>
    <td class=campo align="right" style="border:1px solid #000000">MÊS: <%=ucase(monthname(month(dataliberacao)))&"/"&year(dataliberacao)%></td>
</tr>
<tr>
	<td class=campo colspan=3 style="border:1px solid #000000"><b>-</b></td>
</tr>
</table>
<br>
<table border="0" cellpadding="0" width="650" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo align="center">Ref.</td>
    <td class=titulo align="center">Nome</td>
	<td class=titulo align="center">Assinatura</td>
	<td class=titulo align="center">Data</td>
</tr>
<%
inicio=1
end if

%>
<tr>
	<td class=campo width=40 style="border:1px solid #000000" height=25 valign=middle>&nbsp;<%=rs("chapa")%></td>
    <td class=campo width=215 style="border: 1px solid;font-family:'Courier New'" nowrap valign=middle><%=rs("nome")%> </td>
    <td class=campo style="border:1px solid #000000" valign=middle>&nbsp;</td>
    <td class=campo width=100 style="border:1px solid #000000" valign=middle>&nbsp;</td>
</tr>
<%
linha=linha+1
'lastsecao=rs("codsecao"):inicio=0
rs.movenext:loop


'****************** especial pesquisa
	pesquisa=0
	if pesquisa=1 then
	response.write "</table>"
	response.write "<DIV style=""page-break-after:always""></DIV>"
	linha=5:inicio=1
	rs.movefirst:do while not rs.eof
	if lastsecao<>rs("codsecao") then
		pagina=pagina+1
		if inicio=0 then
			response.write "</table>" : response.write "<DIV style=""page-break-after:always""></DIV>"
		end if
	%>
	<table border="0" cellpadding="2" width="650" cellspacing="0" style="border-collapse: collapse">
	<tr><td colspan=3 class=titulop align="center">PESQUISA DE PREFERÊNCIA - CESTA BÁSICA</td></tr>
	<tr>
		<td class=titulo align="left"  >MÊS: <%=ucase(monthname(month(dataliberacao)))&"/"&year(dataliberacao)%></td>
		<td class=titulo align="center">&nbsp;</td>
		<td class=titulo align="right" >Recursos Humanos</td>
	</tr>
	<tr>
		<td class=campo align="left" colspan=3 style="border:1px solid #000000">Estamos realizando uma pesquisa para determinar a preferência dos funcionários quanto ao modo
		de recebimento do benefício da cesta básica.
		<br>Se você prefere receber o crédito no cartão alimentação assinale a opção (Cartão Alimentação)
		<br>Se você prefere receber a caixa da cesta básica assinale (Cesta Básica)
	</tr>
	<tr>
		<td class=campo colspan=3 style="border:1px solid #000000"><b>Campus: <%=rs("campus")%> - Deptº: <%=rs("setor")%></b></td>
	</tr>
	</table>
	<br>
	<table border="0" cellpadding="0" width="650" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td class=titulo align="center">Ref.</td>
		<td class=titulo align="center">Nome</td>
		<td class=titulo align="center">Cartão Alimentação</td>
		<td class=titulo align="center">Cesta Básica</td>
	</tr>
	<% end if %>
	<tr>
		<td class=campo width=40 style="border:1px solid #000000" height=25 valign=middle>&nbsp;<%=rs("chapa")%></td>
		<td class=campo width=350 style="border: 1px solid;font-family:'Courier New'" nowrap valign=middle><%=rs("nome")%> </td>
		<td class=campo width=130 style="border:1px solid #000000" valign=middle align="center">[&nbsp;&nbsp;&nbsp;]</td>
		<td class=campo width=130 style="border:1px solid #000000" valign=middle align="center">[&nbsp;&nbsp;&nbsp;]</td>
	</tr>
	<%
	linha=linha+1
	lastsecao=rs("codsecao"):inicio=0
	rs.movenext:loop
	end if 'pesquisa
'****************** fim especial pesquisa

rs.close
%>
</table>
<%
end if
%>
</body>
</html>
<%
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing

'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a=0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
'next
'response.write "</tr>"
'if rs.recordcount>0 then rs.movefirst
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'*************** fim teste **********************
%>