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
<title>Rateio de Cesta Básica</title>
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
"where t.campo1 in ('02') and t.sessao=t2.sessao and t.sessao=t3.sessao and t.sessao='" & request("sessao") & "' and opcao in ('C','V') " & _
"order by convert(datetime,substring(registro,51,2)+'/'+substring(registro,53,2)+'/'+substring(registro,47,4)) desc "
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
<p class=titulo style="margin-top:0;margin-bottom:0">Geração de rateio de Cesta Básica<p>
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
sql="select top 100 opcao, t.sessao, convert(datetime,substring(registro,51,2)+'/'+substring(registro,53,2)+'/'+substring(registro,47,4)) as datapedido, " & _
"substring(registro,47,4) as ano, substring(registro,51,2) as mes, " & _
"convert(datetime,substring(registro,55,4)+'/'+substring(registro,59,2)+'/'+substring(registro,61,2)) as dataliberacao, t2.horapedido, t3.valorpedido " & _
"from ttcbasica t, (select sessao, substring(registro,33,8) as horapedido from ttcbasica where campo1 in ('01')) as t2,  " & _
"(select sessao, substring(registro,14,14) as valorpedido from ttcbasica where campo1 in ('05')) as t3 " & _
"where t.campo1 in ('02') and t.sessao=t2.sessao and t.sessao=t3.sessao and opcao in ('C','V') " & _
"order by convert(datetime,substring(registro,51,2)+'/'+substring(registro,53,2)+'/'+substring(registro,47,4)) desc "
rsc.Open sql, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
rsc.movefirst
do while not rsc.eof
sql2="select ano, mes, count(sessao) as freq from (" & sql & ") as t where mes='" & numzero(month(rsc("datapedido")),2) & "' and ano='" & year(rsc("datapedido")) & "' group by mes, ano, opcao "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
if rs("freq")=>1 and cint(rs("ano"))=cint(year(now)) and cint(rs("mes"))>=cint(month(now)) then existe=1 else existe=0
rs.close
%>
<tr>
	<td class="campop"><a href="cb_rateio.asp?sessao=<%=rsc("sessao")%>&opcao=<%=rsc("opcao")%>&total=<%=rsc("valorpedido")/100%>"><%=rsc("sessao")%></a></td>
	<td class="campop" align="center"><%=rsc("opcao")%></td>
	<td class="campop" align="center"><%=rsc("datapedido")%></td>
	<td class="campop" align="center"><%=rsc("horapedido")%></td>
	<td class="campop" align="center"><%=rsc("dataliberacao")%></td>
	<td class="campop" align="right"><%=formatnumber(rsc("valorpedido")/100,2)%></td>
	<td class=campo align="center">&nbsp;
	<%if existe=1 then%>
		<a href="cb_rateio.asp?acao=excluir&s=<%=rsc("sessao")%>&t=<%=rsc("opcao")%>">
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
%>
<table border="0" cellpadding="2" width="650" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td class=grupo align="left"  >Controle de Cesta-Básica</td>
    <td class=grupo align="center">Rateio de Nota Fiscal - Data-Pedido: <%=datapedido%></td>
    <td class=grupo align="right" >Empresa: TICKET</td>
  </tr>
</table>
<%
if request("opcao")="V" then
	titulo1="Valor"
	titulo2="Taxa"
	titulo3="Total"
else
	titulo1="Quant"
	titulo2="Taxa"
	titulo3="Perc.%"
end if

%>
<table border="0" cellpadding="0" width="650" cellspacing="0" style="border-collapse: collapse">
  <tr>
	<td class=titulo>Conta</td>
    <td class=titulo>Código   </td>
    <td class=titulo>Descrição</td>
    <td class=titulo align="center"><%=titulo1%></td>
	<td class=titulo align="center"><%=titulo2%></td>
	<td class=titulo align="center"><%=titulo3%></td>
  </tr>
<%
sql="select sessao, right(campo2,8) as codsecao, s.descricao, total=sum(convert(float,substring(registro,101,9))/100), ttaxa=sum(taxa), " & _
"conta=case when substring(campo2,6,1)='1' then '636' else case when substring(campo2,6,1)='2' then '521' else case when substring(campo2,6,1)='3' then '408' else '' end end end " & _
"from ttcbasica t, corporerm.dbo.psecao s " & _
"where campo1 in ('04') and sessao='" & request("sessao") & "' and opcao='" & request("opcao") & "' and right(campo2,8)=s.codigo collate database_default " & _
"group by sessao, right(campo2,8), s.descricao, case when substring(campo2,6,1)='1' then '636' else case when substring(campo2,6,1)='2' then '521' else case when substring(campo2,6,1)='3' then '408' else '' end end end " & _
"order by case when substring(campo2,6,1)='1' then '636' else case when substring(campo2,6,1)='2' then '521' else case when substring(campo2,6,1)='3' then '408' else '' end end end "
'response.write sql
linha=2
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalcb=0:totaltx=0:totalen=0
rs.movefirst
do while not rs.eof
totalcb=totalcb+cdbl(rs("total"))
totaltx=totaltx+cdbl(rs("ttaxa"))
if request("opcao")="V" then
	valor3=rs("ttaxa")+rs("total"):valor3a=formatnumber(valor3,2)
	valor4=totalcb+totaltx:valor4a=formatnumber(valor4,2)
else
	valor3=rs("total")/request("total"):valor3a=formatpercent(valor3,2)
	valor4=1:valor4a=formatpercent(valor4,2)
end if
%>
  <tr>
	<td class=campo style="border-bottom: 1px solid">&nbsp;<%=rs("conta")%></td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width:1">&nbsp;<%=rs("codsecao")%> </td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width:1">&nbsp;<%=rs("descricao")%></td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width:1" align="right">&nbsp;<%=formatnumber(rs("total"),2)%>&nbsp;</td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width:1" align="right">&nbsp;<%=formatnumber(rs("ttaxa"),2)%>&nbsp;</td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width:1" align="right">&nbsp;<%=valor3a%>&nbsp;</td>
  </tr>
<%
linha=linha+1
rs.movenext
loop
rs.close
%>
  <tr>
  	<td class=titulo style="border-top: 1px solid #000000">&nbsp;</td>
    <td class=titulo style="border-top: 1px solid #000000">&nbsp;</td>
    <td class=titulo style="border-top: 1px solid #000000">&nbsp;</td>
    <td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totalcb,2)%>&nbsp;</td>
    <td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totaltx,2)%>&nbsp;</td>
    <td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=valor4a%>&nbsp;</td>
  </tr>
</table>
<%
linha=linha+1
pagina=pagina+1
response.write "<br>"
response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"

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