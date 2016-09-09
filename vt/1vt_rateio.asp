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
<title>Rateio de Vale Transporte</title>
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

if request("sessao")<>"" then
	sql="select t.sessao, substring(registro,25,8) as datapedido, substring(registro,68,16) as valorpedido, substring(registro,52,16) as taxa, t2.horapedido " & _
"from ttvtransporte t, (select sessao, substring(registro,33,8) as horapedido from ttvtransporte where campo1 in ('01')) as t2  " & _
"where t.campo1 in ('03') and t.sessao=t2.sessao and t.sessao='" & request("sessao") & "' /*order by convert(datetime,substring(registro,25,8))*/ "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	valorpedido=formatnumber(replace(rs("valorpedido"),".",","),2)
	valortaxa=formatnumber(replace(rs("taxa"),".",","),2)
	datapedido=rs("datapedido")
	rs.close
end if
%>

<% if request("sessao")="" then %>
<p class=titulo>Geração de rateio de Vale Transporte
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=450>
<tr>
	<td class=titulop align="center">Controle</td>
	<td class=titulop align="center">Data Pedido</td>
	<td class=titulop align="center">Hora Pedido</td>
	<td class=titulop align="center">Valor Pedido</td>
	<td class=titulop align="center">Taxa Pedido</td>
</tr>
<%
sql="SET DATEFORMAT dmy select t.sessao, substring(registro,25,8) as datapedido, substring(registro,68,16) as valorpedido, substring(registro,52,16) as taxa, t2.horapedido " & _
"from ttvtransporte t, (select sessao, substring(registro,33,8) as horapedido from ttvtransporte where campo1 in ('01')) as t2  " & _
"where t.campo1 in ('03') and t.sessao=t2.sessao order by convert(datetime,substring(registro,25,8)) desc, substring(registro,68,16) "
rsc.Open sql, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
%>
<tr>
	<td class="campop"><a href="vt_rateio.asp?sessao=<%=rsc("sessao")%>"><%=rsc("sessao")%></a></td>
	<td class="campop" align="center"><%=rsc("datapedido")%></td>
	<td class="campop" align="center"><%=rsc("horapedido")%></td>
	<td class="campop" align="right"><%=formatnumber(replace(rsc("valorpedido"),".",","),2)%></td>
	<td class="campop" align="right"><%=formatnumber(replace(rsc("taxa"),".",","),2)%></td>
</tr>
<%
rsc.movenext
loop
rsc.close
%>
</table>
<%
else
%>
<table border="0" cellpadding="2" width="690" cellspacing="" style="border-collapse: collapse">
  <tr>
    <td class=grupo align="left"  >Controle de Vale-Transporte</td>
    <td class=grupo align="center">Rateio de Nota Fiscal - Data-Pedido: <%=datapedido%></td>
    <td class=grupo align="right" >Empresa: TICKET</td>
  </tr>
</table>
<table border="0" cellpadding="0" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>###</td>
    <td class=titulo>Cód.Custo</td>
    <td class=titulo>Descrição</td>
    <td class=titulo align="center">Valor</td>
    <td class=titulo align="center">Taxa</td>
    <td class=titulo align="center">Entrega</td>
    <td class=titulo align="center">Repasses</td>
    <td class=titulo align="center">Acerto</td>
    <td class=titulo align="center">Total</td>
  </tr>
<%
sql="select valor from sptarifa where codigo='BA1'"
rs.Open sql, ,adOpenStatic, adLockReadOnly
taxa1=cdbl(rs("valor"))'$ barueri 
rs.close
taxa2=1.75 '% sptrans
taxa3=2.00 '% cmt - BOM
taxa4=1.75 '% carapicuiba - PEC
taxa4=0
if cdate(datapedido)<dateserial(2009,4,1) then taxa4=0
sql="select sessao, total=sum(case when campo1='07' then substring(registro,60,8)*(convert(float,substring(registro,68,9))/1) else 0 end), " & _
"sptrans=sum(case when substring(registro,83,7)='SPTRANS' or substring(registro,77,6) in ('SPTRAN','METRO','FEPASA') then substring(registro,60,8)*(convert(float,substring(registro,68,9))/1) else 0 end), " & _
"cmt=sum(case when substring(registro,77,6)='CMT' then substring(registro,60,8)*(convert(float,substring(registro,68,9))/1) else 0 end), " & _
"pec=sum(case when substring(registro,83,4)='PEC ' then  substring(registro,60,8)*(convert(float,substring(registro,68,9))/1) else 0 end) " & _
"from ttvtransporte t where campo1 in ('07') and sessao='" & request("sessao") & "' group by sessao "
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalsptrans=rs("sptrans")
totalcmt=rs("cmt")
totalpec=rs("pec")
rs.close
taxasptrans= int(totalsptrans*taxa2+0.5)/100
taxacmt    = int(totalcmt    *taxa3+0.5)/100
taxapec    = int(totalpec    *taxa4+0.5)/100

sql="select sessao, right(campo2,8) as codsecao, s.descricao, " & _
"total=sum(case when campo1='07' then substring(registro,60,8)*(convert(float,substring(registro,68,9))/1.00) else 0 end), " & _
"conta=case when substring(campo2,6,1)='1' then '633' else case when substring(campo2,6,1)='2' then '518' else case when substring(campo2,6,1)='3' then '405' else '' end end end, " & _
"repasse=sum(case when campo1='A' then substring(registro,60,8)*(convert(float,substring(registro,68,9))/1.00) else 0 end), " & _
"repasse1=sum(case when substring(registro,77,6) in ('B.B BA','B.B TR') then " & nraccess(taxa1) & " else 0 end), " & _
"repasse2=sum(case when substring(registro,83,7)='SPTRANS' or substring(registro,77,6) in ('SPTRAN','FEPASA') or substring(registro,77,5)='METRO' " & _
 "then round(substring(registro,60,8)*convert(float,substring(registro,68,9))*" & nraccess(taxa2) & ",0)/100 else 0 end), " & _
"repasse3=sum(case when substring(registro,77,6) in ('CMT') " & _
 "then round(substring(registro,60,8)*convert(float,substring(registro,68,9))*" & nraccess(taxa3) & ",0)/100 else 0 end), " & _
"repasse4=sum(case when substring(registro,83,4) in ('PEC ') " & _
 "then round(substring(registro,60,8)*convert(float,substring(registro,68,9))*" & nraccess(taxa4) & ",0)/100 else 0 end), " & _
"estorno=sum(case when campo1='A' then round(substring(registro,60,8)*convert(float,substring(registro,68,9)),2)/1.00 else 0 end) " & _
"from ttvtransporte t, corporerm.dbo.psecao s " & _
"where campo1 in ('07','A') and sessao='" & request("sessao") & "' and right(campo2,8)=s.codigo collate database_default " & _
"group by sessao, right(campo2,8), s.descricao, " & _
"case when substring(campo2,6,1)='1' then '633' else case when substring(campo2,6,1)='2' then '518' else case when substring(campo2,6,1)='3' then '405' else '' end end end " & _
"order by case when substring(campo2,6,1)='1' then '633' else case when substring(campo2,6,1)='2' then '518' else case when substring(campo2,6,1)='3' then '405' else '' end end end "
linha=2
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalvr=0:totaltx=0:totalen=0
rs.movefirst:do while not rs.eof
	totalvr=totalvr+cdbl(rs("total"))
	taxa=int((cdbl(rs("total")*valortaxa)/valorpedido)*100)/100
	entrega=int((cdbl(rs("total")*115.11)/valorpedido)*100)/100
	repasse1=cdbl(rs("repasse1")):repasse2=rs("repasse2"):repasse3=rs("repasse3"):repasse4=rs("repasse4")
	if cdbl(rs("total"))=0 then repasse1=0
	if rs.recordcount=rs.absoluteposition then taxa=valortaxa-totaltx
	if rs.recordcount=rs.absoluteposition then entrega=115.11-totalen
	if rs.recordcount=rs.absoluteposition then repasse2=taxasptrans-totalre2
	if rs.recordcount=rs.absoluteposition then repasse3=taxacmt-totalre3
	if rs.recordcount=rs.absoluteposition then repasse4=taxapec-totalre4
	totaltx=totaltx+taxa
	totalen=totalen+entrega
	totalacerto=totalacerto+cdbl(rs("estorno"))
	totalre1=totalre1+repasse1:totalre2=totalre2+repasse2:totalre3=totalre3+repasse3:totalre4=totalre4+repasse4
	rateio=cdbl(rs("total"))+taxa+entrega+repasse1+repasse2+repasse3+repasse4+cdbl(rs("estorno"))
	totalra=totalra+rateio
%>
<tr>
	<td class=campo style="border-bottom: 1px solid"><%=rs("conta")%></td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("codsecao")%> </td>
    <td class="campor" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("descricao")%></td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1" align="right"><%=formatnumber(rs("total"),2)%></td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1" align="right"><%=formatnumber(taxa,2)%>       </td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1" align="right"><%=formatnumber(entrega,2)%>    </td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1" align="right"><%=formatnumber(repasse1+repasse2+repasse3+repasse4,2)%>   </td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1" align="right"><%=formatnumber(rs("estorno"),2)%>   </td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1" align="right"><%=formatnumber(rateio,2)%>     </td>
</tr>
<%
linha=linha+1
rs.movenext:loop
rs.close
%>
<tr>
	<td class=titulo style="border-top: 1px solid #000000">&nbsp;</td>
	<td class=titulo style="border-top: 1px solid #000000">&nbsp;</td>
	<td class=titulo style="border-top: 1px solid #000000">&nbsp;</td>
	<td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totalvr,2)%></td>
	<td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totaltx,2)%></td>
	<td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totalen,2)%></td>
	<td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totalre1+totalre2+totalre3+totalre4,2)%></td>
	<td class=fundo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totalacerto,2)%></td>
	<td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totalra,2)%></td>
</tr>
</table>
<p style="margin-top:0;margin-bottom:0;font-family:Courier New;font-size:8pt">Total de Repasses:
<br>B.B. Barueri (R$ 2,50 por cartão....: <%=formatnumber(totalre1,2)%>
<br>SPTRANS (1,75% sobre total).........: <%=formatnumber(totalre2,2)%>
<br>CMT-BOM (2% sobre total)............: <%=formatnumber(totalre3,2)%>
<br>PEC-Carapicuiba (1,75% sobre total).: <%=formatnumber(totalre4,2)%>
<%

linha=linha+1
pagina=pagina+1
response.write "<br>"
response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"

if session("usuariomaster")="02379" then
response.write "<DIV style=""page-break-after:always""></DIV>"
sql="SELECT sessao, campo1, Mid(registro,77,6) AS Empresa, Mid(registro,83,12) AS Bilhete, " & _
"Sum(CDbl(Mid(registro,60,8))) AS Quant, Mid(registro,68,9)/100 AS Unitario, " & _
"Sum(CDbl(Mid(registro,60,8))*Mid(registro,68,9)/100) AS Total " & _
"FROM ttvtransporte " & _
"GROUP BY sessao, campo1, Mid(registro,77,6), Mid(registro,83,12), Mid(registro,68,9)/100 " & _
"HAVING sessao='" & request("sessao") & "' AND campo1='07' "
sql="SELECT sessao, campo1, substring(registro,77,6) AS Empresa, substring(registro,83,12) AS Bilhete, " & _
"Sum(convert(integer,substring(registro,60,8))) AS Quant, min(convert(float,substring(registro,68,9))) AS Unitario, " & _
"Sum(convert(float,substring(registro,60,8))*convert(float,substring(registro,68,9))/1) AS Total " & _
"FROM ttvtransporte GROUP BY sessao, campo1, substring(registro,77,6), substring(registro,83,12) " & _
"HAVING sessao='" & request("sessao") & "' AND campo1='07' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
'*************** inicio teste **********************
%>
<table border='1' cellpadding='2' cellspacing='0' style='border-collapse:collapse'>
<tr>
	<td class=titulo>Empresa</td>
	<td class=titulo>Bilhete</td>
	<td class=titulo>Quant.</td>
	<td class=titulo>Unitário</td>
	<td class=titulo>Total</td>
</tr>
<%
chkqt=0:chktt=0
do while not rs.eof
%>
<tr>
	<td class=campo><%=rs("empresa")%></td>
	<td class=campo><%=rs("bilhete")%></td>
	<td class=campo align="right"><%=rs("quant")%></td>
	<td class=campo align="right"><%=formatnumber(rs("unitario"),2)%></td>
	<td class=campo align="right"><%=formatnumber(rs("total"),2)%></td>
</tr>
<%
chkqt=chkqt+rs("quant"):chktt=chktt+rs("total")
rs.movenext
loop
%>
<tr>
	<td class=titulo colspan=2>Total</td>
	<td class=titulo align="right"><%=chkqt%></td>
	<td class=titulo></td>
	<td class=titulo align="right"><%=formatnumber(chktt,2)%></td>
</tr>
</table>
<%
'*************** fim teste **********************
rs.close
end if 'session

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