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
<script language="javascript" type="text/javascript">
<!--
/****************************************************
     Author: Eric King
     Url: http://redrival.com/eak/index.shtml
     This script is free to use as long as this info is left in
     Featured on Dynamic Drive script library (http://www.dynamicdrive.com)
****************************************************/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>

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
rs.CursorLocation=3
rsc.CursorLocation=3

if request("sessao")<>"" then
sql="SET DATEFORMAT dmy select t.sessao, substring(registro,25,10) as datapedido, substring(registro,95,8) as iniciopedido, substring(registro,70,16) as valorpedido, substring(registro,54,16) as taxa, substring(registro,35,8) as horapedido " & _
"from ttvtransporte t where t.campo1 in ('09') and t.empresa='VB' and sessao='" & request("sessao") & "' " & _
"order by convert(datetime,substring(registro,25,10)) desc, substring(registro,70,16) "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	valorpedido=formatnumber(replace(rs("valorpedido"),".",","),2)
	valorpedido=rs("valorpedido")+0
	valortaxa=formatnumber(replace(rs("taxa"),".",","),2)
	valortaxa=rs("taxa")+0
	datapedido=rs("datapedido")
	iniciopedido=rs("iniciopedido")
	rs.close
end if
if request("apagar")<>"" then
	sql="delete from ttvtransporte where sessao='" & request("apagar") & "'"
	conexao.execute sql
end if
%>

<% if request("sessao")="" or request("apagar")<>"" then %>
<p class=titulo>Geração de rateio de Vale Transporte
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=550>
<tr>
	<td class=titulop align="center">Controle</td>
	<td class=titulop align="center">Data Pedido</td>
	<td class=titulop align="center">Hora Pedido</td>
	<td class=titulop align="center">Valor Pedido</td>
	<td class=titulop align="center">Taxa Pedido</td>
	<td class=titulop align="center"></td>
</tr>
<%
sql="SET DATEFORMAT dmy select t.sessao, substring(registro,25,10) as datapedido, substring(registro,70,16) as valorpedido, substring(registro,54,16) as taxa, substring(registro,35,8) as horapedido " & _
"from ttvtransporte t where t.campo1 in ('09') and t.empresa='VB' " & _
"order by convert(datetime,substring(registro,25,10)) desc, substring(registro,35,8) desc "
rsc.Open sql, ,adOpenStatic, adLockReadOnly
if rsc.recordcount=0 or rsc.recordcount=-1 then
%>
<tr><td class="campop" colspan=5>Não existem rateios</td></tr>
<%
else 'rsc.recordcount=0
'-----------tem rateios
rsc.movefirst
do while not rsc.eof
if rsc.absoluteposition=1 then primeiro=1 else primeiro=0
%>
<tr>
	<td class="campop"><a href="vt_rateiovb.asp?sessao=<%=rsc("sessao")%>"><%=rsc("sessao")%></a>
	&nbsp;<a href="vt_rateiovb.asp?sessao=<%=rsc("sessao")%>&entrega=30"><img src="../images/truck.gif" border=0 alt="Gerar o rateio com taxa de entrega"></a>
	<%if chkdata=rsc("datapedido") or session("usuariomaster")="02379" then%>
	&nbsp;<a href="vt_rateiovb.asp?apagar=<%=rsc("sessao")%>"><img src="../images/trash.gif" border=0 alt="Apagar rateio duplicado"></a>
	<%end if%>
	</td>
	<td class="campop" align="center"><%=rsc("datapedido")%></td>
	<td class="campop" align="center"><%=rsc("horapedido")%></td>
	<td class="campop" align="right"><%=formatnumber(replace(rsc("valorpedido"),".",","),2)%></td>
	<td class="campop" align="right"><%=formatnumber(replace(rsc("taxa"),".",","),2)%></td>
	<td class="campop" align="center">
	<a class=r href="vt_analiticovb.asp?sessao=<%=rsc("sessao")%>&primeiro=<%=primeiro%>" onclick="NewWindow(this.href,'AnaliticoCompraVT','645','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Leaf.gif" width="16" height="16" border="0" alt="Ver funcionários incluidos na compra"></a>	
	</td>
</tr>
<%
chkdata=rsc("datapedido")
rsc.movenext
loop
rsc.close
end if 'rsc.recordcount=0
%>
</table>
<%
else
%>
<table border="0" cellpadding="2" width="690" cellspacing="" style="border-collapse: collapse">
  <tr>
    <td class=grupo align="left"  >Controle de Vale-Transporte</td>
    <td class=grupo align="center">Rateio de Nota Fiscal - Data-Pedido: <%=datapedido%></td>
    <td class=grupo align="right" >Empresa: VB</td>
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
'$ tx entrega
sql="select valor from iParametros where parametro='txentrega'"
rs.open sql, ,adOpenStatic, adLockReadOnly:txentrega=cdbl(replace(rs("valor"),".",",")):rs.close
txentrega=request("entrega")
'$ barueri 
sql="select top 1 valor from corporerm.dbo.ptarifa where codigo='BA1' and '" & iniciopedido & "' between iniciovigencia and finalvigencia "
rs.Open sql, ,adOpenStatic, adLockReadOnly:taxa1=cdbl(rs("valor")):rs.close
'$ cmto
sql="select top 1 valor from corporerm.dbo.ptarifa where codigo='03' and '" & iniciopedido & "' between iniciovigencia and finalvigencia "
rs.Open sql, ,adOpenStatic, adLockReadOnly:taxa5=cdbl(rs("valor")):rs.close
'% sptrans
sql="select valor from iParametros where parametro='taxa2'"
rs.open sql, ,adOpenStatic, adLockReadOnly:taxa2=cdbl(replace(rs("valor"),".",",")):rs.close
'% cmt - BOM
sql="select valor from iParametros where parametro='taxa3'"
rs.open sql, ,adOpenStatic, adLockReadOnly:taxa3=cdbl(replace(rs("valor"),".",",")):rs.close
'% carapicuiba - PEC
sql="select valor from iParametros where parametro='taxa4'"
rs.open sql, ,adOpenStatic, adLockReadOnly:taxa4=cdbl(replace(rs("valor"),".",",")):rs.close
'$ indexador de cartao
sql="select top 1 valor from corporerm.dbo.ptarifa where codigo='SPI' and '" & iniciopedido & "' between iniciovigencia and finalvigencia "
rs.Open sql, ,adOpenStatic, adLockReadOnly:vbmanut=cdbl(rs("valor"))*0.2:rs.close
sql="select valor from iParametros where parametro='vbmanut'"
rs.open sql, ,adOpenStatic, adLockReadOnly:vbmanut=cdbl(replace(rs("valor"),".",",")):rs.close
'$ taxa portabilidade
sql="select valor from iParametros where parametro='vbport'"
rs.open sql, ,adOpenStatic, adLockReadOnly:vbport=cdbl(replace(rs("valor"),".",",")):rs.close
'$ taxa administração
sql="select valor from iParametros where parametro='vbadmin'"
rs.open sql, ,adOpenStatic, adLockReadOnly:vbadmin=cdbl(replace(rs("valor"),".",",")):rs.close

sql=" select sessao, total=sum(case when campo1='04' then substring(registro,106,14)*(convert(float,substring(registro,120,14))/100) else 0 end), " & _
"sptrans=sum(case when substring(registro,31,6) in ('695','696','697','701') then substring(registro,106,14)*(convert(float,substring(registro,120,14))/100) else 0 end), " & _
"cmt=sum(case when substring(registro,31,6) in ('512','513','514','515','517','518'/*,'6312'*/) then substring(registro,106,14)*(convert(float,substring(registro,120,14))/100) else 0 end), " & _
"pec=sum(case when substring(registro,31,6) in ('6411') then substring(registro,106,14)*(convert(float,substring(registro,120,14))/100) else 0 end), " & _
"bem=sum(case when substring(registro,31,6) in ('5998') then substring(registro,106,14)*(convert(float,substring(registro,120,14))/100) else 0 end), " & _
"bbtt=sum(case when substring(registro,31,6) in ('6407','6409') then substring(registro,106,14)*(convert(float,substring(registro,120,14))/100) else 0 end), " & _
"nrsptrans=sum(case when substring(registro,31,6) in ('695','696','697','701') then 1 else 0 end), " & _
"nrcmt=sum(case when substring(registro,31,6) in ('512','513','514','515','517','518') then 1 else 0 end), " & _
"nrpec=sum(case when substring(registro,31,6) in ('6411') then 1 else 0 end), " & _
"nrbem=sum(case when substring(registro,31,6) in ('5998') then 1 else 0 end), " & _
"nrbbtt=sum(case when substring(registro,31,6) in ('6407') then 1 else 0 end) " & _ 
"from ttvtransporte t where campo1 in ('04') and sessao='" & request("sessao") & "' group by sessao "
'('6407','6409','6312') 
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalsptrans=rs("sptrans")
totalcmt=rs("cmt")
totalpec=rs("pec")
totalcartoes=rs("nrsptrans")+rs("nrcmt")+rs("nrpec")+rs("nrbem") +rs("nrbbtt")
stringcartoes=" | " & rs("nrsptrans") & " Sptrans | " & (rs("nrcmt")+rs("nrpec")+rs("nrbbtt")) & " BOM/BB | " & rs("nrbem") & " BEM"
rs.close
taxasptrans= int(totalsptrans*taxa2+0.5)/100
taxacmt    = int(totalcmt    *taxa3+0.5)/100
taxapec    = int(totalpec    *taxa4+0.5)/100
taxaport   = vbport

sql="select sessao, right(campo2,8) as codsecao, s.descricao, " & _
"total=sum(case when campo1='04' then substring(registro,106,14)*(convert(float,substring(registro,120,14))/100) else 0 end), " & _
"conta=case when substring(campo2,6,1)='1' then '633' else case when substring(campo2,6,1)='2' then '518' else case when substring(campo2,6,1)='3' then '405' else '' end end end, " & _
"repasse=sum(case when campo1='A' then substring(registro,106,14)*(convert(float,substring(registro,120,14))/100) else 0 end), " & _
"repasse1=sum(case when substring(registro,31,6) in ('6407','6409','6402') then " & nraccess(taxa1) & " else 0 end), " & _
"repasse2=sum(case when substring(registro,31,6) in ('695','696','697','701') then round(substring(registro,106,14)*(convert(float,substring(registro,120,14))/100)*" & nraccess(taxa2) & ",0)/100 else 0 end), " & _
"repasse3=sum(case when substring(registro,31,6) in ('512','513','514','515','517','518'/*,'6312'*/) then round(substring(registro,106,14)*(convert(float,substring(registro,120,14))/100)*" & nraccess(taxa3) & ",0)/100 else 0 end), " & _
"repasse4=sum(case when substring(registro,31,6) in ('6411') then round(substring(registro,106,14)*(convert(float,substring(registro,120,14))/100)*" & nraccess(taxa4) & ",0)/100 else 0 end), " & _
"repasse5=sum(case when substring(registro,31,6) in ('5998') then " & nraccess(taxa5) & " else 0 end), " & _
"index1=sum(case when substring(registro,31,6) in ('6407','6402') then 1 else 0 end)*" & nraccess(vbmanut) & ", " & _
"index2=sum(case when substring(registro,31,6) in ('695','696','697','701') then 1 else 0 end)*" & nraccess(vbmanut) & ", " & _
"index3=sum(case when substring(registro,31,6) in ('512','513','514','515','517','518') then 1 else 0 end)*" & nraccess(vbmanut) & ", " & _
"index4=sum(case when substring(registro,31,6) in ('6411') then 1 else 0 end)*" & nraccess(vbmanut) & ", " & _
"index5=sum(case when substring(registro,31,6) in ('5998') then 1 else 0 end)*" & nraccess(vbmanut) & ", " & _
"estorno=sum(case when campo1='A' then round(substring(registro,106,14)*(convert(float,substring(registro,120,14))/100),2)/1.00 else 0 end) " & _
"from ttvtransporte t, corporerm.dbo.psecao s " & _
"where campo1 in ('04','A') and sessao='" & request("sessao") & "' and right(campo2,8)=s.codigo collate database_default " & _
"group by sessao, right(campo2,8), s.descricao, " & _
"case when substring(campo2,6,1)='1' then '633' else case when substring(campo2,6,1)='2' then '518' else case when substring(campo2,6,1)='3' then '405' else '' end end end " & _
"order by case when substring(campo2,6,1)='1' then '633' else case when substring(campo2,6,1)='2' then '518' else case when substring(campo2,6,1)='3' then '405' else '' end end end "
linha=2
rs.Open sql, ,adOpenStatic, adLockReadOnly

totalvr=0:totaltx=0:totalen=0:totalpt=0:totalmanut=0
rs.movefirst:do while not rs.eof
	repasse1=cdbl(rs("repasse1")):repasse2=rs("repasse2"):repasse3=rs("repasse3"):repasse4=rs("repasse4"):repasse5=cdbl(rs("repasse5"))
	trepasse=repasse1+repasse2+repasse3+repasse4+repasse5
	totalvr=totalvr+cdbl(rs("total"))
	taxa=int(( (cdbl(rs("total") + trepasse) * valortaxa)/valorpedido)*100)/100
	entrega=int((cdbl(rs("total")*txentrega)/valorpedido)*100)/100
	portab=int((cdbl(rs("total")*taxaport)/valorpedido)*100)/100
	manutencao=cdbl(rs("index1"))+cdbl(rs("index2"))+cdbl(rs("index3"))+cdbl(rs("index4"))+cdbl(rs("index5"))
	if cdbl(rs("total"))=0 then repasse1=0
	
	
	if rs.recordcount=rs.absoluteposition then
		repasse2=taxasptrans-totalre2
		repasse3=taxacmt-totalre3
		repasse4=taxapec-totalre4

		novataxa=int(((totalre1+totalre2+totalre3+totalre4+totalre5)+(repasse1+repasse2+repasse3+repasse4+repasse5))*vbadmin)/100
'response.write novataxa
'response.write "<br>vt "&valortaxa
'response.write "<br>tre"&(totalre1+totalre2+totalre3+totalre4+totalre5)
'response.write "<br>re "&(repasse1+repasse2+repasse3+repasse4+repasse5)
'response.write "<br>tx "&vbadmin
		taxa=valortaxa+novataxa-totaltx
		entrega=txentrega-totalen
		portab=taxaport-totalpt
	end if
	totaltx=totaltx+taxa
	totalen=totalen+entrega
	totalpt=totalpt+portab
	totalacerto=totalacerto+cdbl(rs("estorno"))
	totalmanut=totalmanut+manutencao
	totalre1=totalre1+repasse1:totalre2=totalre2+repasse2:totalre3=totalre3+repasse3:totalre4=totalre4+repasse4:totalre5=totalre5+repasse5
	rateio=cdbl(rs("total"))+taxa+entrega+portab+repasse1+repasse2+repasse3+repasse4+repasse5+cdbl(rs("estorno"))+manutencao
	totalra=totalra+rateio
%>
<tr>
	<td class=campo style="border-bottom: 1px solid"><%=rs("conta")%></td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("codsecao")%> </td>
    <td class="campor" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("descricao")%></td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1" align="right"><%=formatnumber(rs("total"),2)%></td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1" align="right"><%=formatnumber(taxa,2)%></td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1" align="right"><%=formatnumber(entrega+portab,2)%>    </td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1" align="right"><%=formatnumber(repasse1+repasse2+repasse3+repasse4+repasse5+manutencao,2)%>   </td>
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
	<td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totalen+totalpt,2)%></td>
	<td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totalre1+totalre2+totalre3+totalre4+totalre5+totalmanut,2)%></td>
	<td class=fundo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totalacerto,2)%></td>
	<td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totalra,2)%></td>
</tr>
</table>
<p style="margin-top:0;margin-bottom:0;font-family:Courier New;font-size:8pt">Total de Repasses:
<br>B.B. Barueri (R$ <%=formatnumber(taxa1,2)%> por cartão)...: <%=formatnumber(totalre1,2)%>
<br>SPTRANS (<%=formatnumber(taxa2,2)%>% sobre total).........: <%=formatnumber(totalre2,2)%>
<br>CMT-BOM (<%=formatnumber(taxa3,2)%>% sobre total).........: <%=formatnumber(totalre3,2)%>
<br>PEC-Carapicuiba (<%=formatnumber(taxa4,2)%>% sobre total).: <%=formatnumber(totalre4,2)%>
<br>CMTO Bem Osasco (R$ <%=formatnumber(taxa5,2)%> p/cartão)..: <%=formatnumber(totalre5,2)%>
<br>Crédito em <%=totalcartoes%> cartões x R$ <%=vbmanut%>....: <%=formatnumber(totalmanut,2)%> <%=stringcartoes%>
<br>Taxa de Portabilidade...............: <%=formatnumber(vbport,2)%>
<%

linha=linha+1
pagina=pagina+1
response.write "<br>"
response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"

if session("usuariomaster")="02379" or session("usuariomaster")="02675" then
response.write "<DIV style=""page-break-after:always""></DIV>"
sql="SELECT sessao, campo1, substring(registro,46,60) AS Empresa, convert(float,substring(registro,120,14)) AS Bilhete, " & _
"Sum(convert(integer,substring(registro,106,14))) AS Quant, min((convert(float,substring(registro,120,14))/100)) AS Unitario, " & _
"Sum(convert(float,substring(registro,106,14))*(convert(float,substring(registro,120,14))/100)/1) AS Total " & _
"FROM ttvtransporte GROUP BY sessao, campo1, substring(registro,46,60), substring(registro,120,14) " & _
"HAVING sessao='" & request("sessao") & "' AND campo1='04' order by substring(registro,46,60)"
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