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
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Pessoas no Pedido de VT</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, rs, rs2, t(4)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

sessao=request("sessao")
anterior=request("anterior")
primeiro=request("primeiro")
compara=request("compara")

if compara<>"yes" then

sql1="select registro from ttvtransporte where campo1='09' and sessao='" & sessao & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
	datapedido=mid(rs("registro"),25,10)
	horapedido=mid(rs("registro"),35,8)
	taxa      =mid(rs("registro"),54,16)*1
	pedido    =mid(rs("registro"),70,16)*1
	pessoas   =mid(rs("registro"),43,5)*1
rs.close

sqla="SELECT sessao, campus=Left(CODSECAO,2), codtipo, f.chapa, nome, total, dias=f.DIASUTPROXMES " & _
"from ( SELECT vt.sessao, chapa=vt.campo3, " & _
"sum( substring(registro,106,14)*(convert(float,substring(registro,120,14))/100) ) as total " & _
"FROM ttvtransporte vt " & _
"GROUP BY vt.sessao, vt.campo3, vt.campo1 " & _
"HAVING vt.sessao='" & sessao & "' AND vt.campo1='04' " & _
") v inner join corporerm.dbo.pfunc f on f.chapa collate database_default=v.chapa " & _
"order by campus, CODTIPO, nome "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
if primeiro=1 then medio=rs("total")/rs("dias") else medio=0
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=550>
<th class=titulo colspan=7>Relação de Pessoas incluidas no pedido feito em <%=datapedido%> - <%=horapedido%></th>
<tr>
	<td class=titulor >Campus</td>
	<td class=titulor >Tipo</td>
	<td class=titulor >Chapa</td>
	<td class=titulor align="center">Nome</td>
	<td class=titulor align="center">Valor</td>
	<%if primeiro=1 then%>
	<td class=titulor align="center">Dias</td>
	<td class=titulor >Média</td>
	<%else%>
	<td class=titulor colspan=2></td>
	<%end if%>
</tr>
<%
rs.movefirst
do while not rs.eof
if ultimoc<>rs("campus") and ultimoc<>"" then
	response.write "<tr><td class=titulor colspan=4>Sub-total do <i>campus</i> " & ultimoc & "</td><td class=titulor align=""right"">" & formatnumber(subtotal,2) & "</td><td class=""campor"" colspan=2></td></tr>"
	total=total+subtotal:subtotal=0
end if
%>
<tr>
	<td class="campor" align="center"><%=rs("campus")%></td>
	<td class="campor" align="center"><%=rs("codtipo")%></td>
	<td class="campor" align="center"><%=rs("chapa")%></td>
	<td class="campor" align="left"><%=rs("nome")%></td>
	<td class="campor" align="right"><%=formatnumber(rs("total"),2)%></td>
	<%if primeiro=1 then%>
	<td class="campor" align="center"><%=rs("dias")%></td>
	<td class="campor" align="right"><%=formatnumber(medio,2)%></td>
	<%else%>
	<td class="campor" colspan=2></td>
	<%end if%>
</tr>
<%
subtotal=subtotal+rs("total")
ultimoc=rs("campus")
rs.movenext
loop
response.write "<tr><td class=titulor colspan=4>Sub-total do <i>campus</i> " & ultimoc & "</td><td class=titulor align=""right"">" & formatnumber(subtotal,2) & "</td><td class=""campor"" colspan=2></td></tr>"
total=total+subtotal:subtotal=0
response.write "<tr><td class=titulor colspan=4>Total dos <i>campis</i> </td><td class=titulor align=""right"">" & formatnumber(total,2) & "</td><td class=""campor"" colspan=2></td></tr>"
%>
<tr>
	<td class=grupo colspan=7><%=rs.recordcount%> funcionários</td>
</tr>
<%
else
	response.write "<tr><td class=campo colspan=6>Sem compra efetuada</td></tr>"
end if
%>
</table>
<p>Valor do Pedido: <%=pedido%>
<br>Taxa: <%=taxa%>
<br>

<%
rs.close
if primeiro=1 then
	response.write "<br>Comparar com"
	sql="SET DATEFORMAT dmy select top 3 t.sessao, substring(registro,25,10) as datapedido " & _
	"from ttvtransporte t where t.campo1 in ('09') and t.empresa='VB' and t.sessao<>'" & sessao & "' " & _
	"order by convert(datetime,substring(registro,25,10)) desc"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	do while not rs.eof
	%>
	<a href="vt_analiticovb.asp?compara=yes&sessao=<%=sessao%>&anterior=<%=rs("sessao")%>"><%=rs("datapedido")%></a> |
	<%
	rs.movenext
	loop
	rs.close
end if

else 'compara=yes

sql1="select registro from ttvtransporte where campo1='09' and sessao='" & sessao & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
	datapedido = mid(rs("registro"),25,10) : horapedido = mid(rs("registro"),35,8)
	taxa = mid(rs("registro"),54,16)*1 : pedido = mid(rs("registro"),70,16)*1
	pessoas = mid(rs("registro"),43,5)*1
rs.close

sql1="select registro from ttvtransporte where campo1='09' and sessao='" & anterior & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
	datapedido2 = mid(rs("registro"),25,10) : horapedido2 = mid(rs("registro"),35,8)
	taxa2 = mid(rs("registro"),54,16)*1 : pedido2 = mid(rs("registro"),70,16)*1
	pessoas2 = mid(rs("registro"),43,5)*1
rs.close

sql2="If Exists(Select * from Tempdb..SysObjects Where Name Like '##compara%')  drop table ##compara"
conexao.execute sql2

sql2="select campus=left(codsecao,2), z.chapa, f.nome, f.codtipo " & _
", atual= sum(case when sessao='" & sessao & "' then total else 0 end), dias1=min(f.DIASUTPROXMES) " & _
", anterior= sum(case when sessao='" & anterior & "' then total else 0 end), dias2=min(f.DIASUTEISMES) " & _
"from ( " & _
"SELECT vt.sessao, chapa=vt.campo3, total=sum( substring(registro,106,14)*(convert(float,substring(registro,120,14))/100) ) " & _
"FROM ttvtransporte vt GROUP BY vt.sessao, vt.campo3, vt.campo1 HAVING vt.sessao='" & sessao & "' AND vt.campo1='04' " & _
"union " & _
"SELECT vt.sessao, chapa=vt.campo3, total=sum( substring(registro,106,14)*(convert(float,substring(registro,120,14))/100) ) " & _
"FROM ttvtransporte vt GROUP BY vt.sessao, vt.campo3, vt.campo1 HAVING vt.sessao='" & anterior & "' AND vt.campo1='04' " & _
") z inner join corporerm.dbo.PFUNC f on f.CHAPA collate database_default=z.chapa " & _
"group by left(codsecao,2), z.chapa, f.nome, f.codtipo " & _
"order by left(codsecao,2), f.nome "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
total=0
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=600>
<th class=titulo colspan=11>Comparação dos Pedidos efetuados em <%=datapedido%> e <%=datapedido2%></th>
<tr>
	<td class=titulor rowspan=2>Campus</td>
	<td class=titulor rowspan=2 >Tipo</td>
	<td class=titulor rowspan=2 >Chapa</td>
	<td class=titulor rowspan=2 align="center">Nome</td>
	<td class=titulor colspan=3>Atual</td>
	<td class=titulor colspan=3>Anterior</td>
	<td class=titulor rowspan=2 align="center">Var.</td>
</tr>
<tr>	
	<td class=titulor align="center">Valor</td>
	<td class=titulor align="center">Dias</td>
	<td class=titulor align="center">Média</td>
	<td class=titulor align="center">Valor</td>
	<td class=titulor align="center">Dias</td>
	<td class=titulor align="center">Média</td>
</tr>
<%
rs.movefirst
do while not rs.eof
if ultimoc<>rs("campus") and ultimoc<>"" then
	response.write "<tr><td class=titulor colspan=4>Sub-total do <i>campus</i> " & ultimoc & "</td><td class=titulor align=""right"">" & formatnumber(subtotal1,2) & "</td><td class=""campor"" colspan=2></td><td class=titulor align=""right"">" & formatnumber(subtotal2,2) & "</td><td class=""campor"" colspan=3></td></tr>"
	total1=total1+subtotal1:subtotal1=0
	total2=total2+subtotal2:subtotal2=0
end if
if rs("dias1")>0 then media1=rs("atual")/rs("dias1") else media1=0
if rs("dias2")>0 then media2=rs("anterior")/rs("dias2") else media2=0
variacao=media1-media2
if variacao<-0.1 or variacao>0.1 then variacao=formatnumber(variacao,2) else variacao="-----"
%>
<tr>
	<td class="campor" align="center"><%=rs("campus")%></td>
	<td class="campor" align="center"><%=rs("codtipo")%></td>
	<td class="campor" align="center"><%=rs("chapa")%></td>
	<td class="campor" align="left"><%=rs("nome")%></td>
	<td class="campor" align="right"><%=formatnumber(rs("atual"),2)%></td>
	<td class="campor" align="center"><%=rs("dias1")%></td>
	<td class="campor" align="right"><%=formatnumber(media1,2)%></td>
	<td class="campor" align="right"><%=formatnumber(rs("anterior"),2)%></td>
	<td class="campor" align="center"><%=rs("dias2")%></td>
	<td class="campor" align="right"><%=formatnumber(media2,2)%></td>
	<td class="campor" align="right"><%=variacao%></td>
</tr>
<%
subtotal1=subtotal1+rs("atual")
subtotal2=subtotal2+rs("anterior")
ultimoc=rs("campus")
rs.movenext
loop
response.write "<tr><td class=titulor colspan=4>Sub-total do <i>campus</i> " & ultimoc & "</td><td class=titulor align=""right"">" & formatnumber(subtotal1,2) & "</td><td class=""campor"" colspan=2></td><td class=titulor align=""right"">" & formatnumber(subtotal2,2) & "</td><td class=""campor"" colspan=3></td></tr>"
total=total+subtotal:subtotal=0
response.write "<tr><td class=titulor colspan=4>Total dos <i>campis</i> </td><td class=titulor align=""right"">" & formatnumber(total1,2) & "</td><td class=""campor"" colspan=2></td><td class=titulor align=""right"">" & formatnumber(total2,2) & "</td><td class=""campor"" colspan=3></td></tr>"
%>
</table>

<%
rs.close
end if 'compara

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>
