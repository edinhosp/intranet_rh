<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a88")="N" or session("a88")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Relação de Ocorrências</title>
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
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
'rs.Open sql, ,adOpenStatic, adLockReadOnly
	
if request.form="" then	
data1=now
data2=dateserial(year(data1),month(data1),1)
data3=dateserial(year(data1),month(data1),day(data1)-1)
sql="SELECT d.CODSECAO, s.DESCRICAO " & _
"FROM corporerm.dbo.APARFUN AS c, corporerm.dbo.PFUNC AS d, corporerm.dbo.PSECAO s " & _
"WHERE c.CHAPA = d.CHAPA and d.CODSECAO = s.CODIGO and d.CODSITUACAO<>'D' AND d.CODSINDICATO<>'03' " & _
"GROUP BY d.CODSECAO, s.DESCRICAO " & _
"ORDER BY s.descricao, d.CODSECAO "
%>
<form method="POST" action="rpt_ocorrencias.asp" name="form">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=420>
<tr>
	<td class=titulo colspan=3>Relação de Ocorrências - Ponto</td>
</tr>
<tr>
	<td class=grupo>Data Inicial</td>
	<td class=grupo>Data Final</td>
	<td class=grupo>Opções</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="dti" size="8" value="<%=data2%>"></td>
	<td class=titulo><input type="text" name="dtf" size="8" value="<%=data3%>"></td>
	<td class=titulo><input type="checkbox" name="quebra" value="on"> Quebra de página por setor?</td>
</tr>
<tr>
	<td class=grupo colspan=3>Setor</td>
</tr>
<tr>
	<td class=titulo colspan=3>
	<select size="1" name="setor">
	<option value="0">Todos setores</option>
<%
rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
%>
	<option value="<%=rs("codsecao")%>"><%=rs("codsecao") & "-" & rs("descricao")%></option>
<%
rs.movenext:loop
rs.close
%>	
	</select>
	</td>
</tr>
<tr>
	<td class=titulo colspan=3>
	<input type="text" value="" size="5" maxlength="5" name="ch1">
	<input type="text" value="" size="5" maxlength="5" name="ch2">
	<input type="text" value="" size="5" maxlength="5" name="ch3">
	<input type="text" value="" size="5" maxlength="5" name="ch4">
	<input type="text" value="" size="5" maxlength="5" name="ch5">
	</td>
</tr>
<tr>
	<td class=titulo colspan=3><input type="submit" value="Gerar relatório" name="Gerar" class="button">
	</td>
</tr>
</table>
</form>
<%

else 'request.form
inicio=1
data1=now
data2=dateserial(year(data1),month(data1)-1,1)
data0=request.form("dti")
data0f=request.form("dtf")
dataant=dateserial(year(data0),month(data0),day(data0)-1)
datarel=dateserial(year(data0),month(data0),day(data0))
datarelf=dateserial(year(data0f),month(data0f),day(data0f))

if request.form("setor")="0" then criterio="" else criterio=" AND d.codsecao='" & request.form("setor") & "' "
ch1=request.form("ch1"):ch2=request.form("ch2"):ch3=request.form("ch3"):ch4=request.form("ch4"):ch5=request.form("ch5")
if ch1<>"" or ch2<>"" or ch3<>"" or ch4<>"" or ch5<>"" then
	chapas=" and d.chapa in ("
	if ch1<>"" then chapas=chapas & "'" & ch1 & "'"
		if ch1<>"" and ch2<>"" then chapas=chapas  & ","
	if ch2<>"" then chapas=chapas & "'" & ch2 & "'"
		if ch2<>"" and ch3<>"" then chapas=chapas  & ","
	if ch3<>"" then chapas=chapas & "'" & ch3 & "'"
		if ch3<>"" and ch4<>"" then chapas=chapas  & ","
	if ch4<>"" then chapas=chapas & "'" & ch4 & "'"
		if ch4<>"" and ch5<>"" then chapas=chapas  & ","
	if ch5<>"" then chapas=chapas & "'" & ch5 & "'"
	chapas=chapas & ") "
end if

sqlant="select chapa, screditos=sum(screditos), sdebitos=sum(sdebitos) from (" & _
"SELECT CHAPA, Sum(EXTRAFAIXA1+EXTRAFAIXA2+EXTRAFAIXA3+EXTRAFAIXA4+EXTRAFAIXA5+EXTRADESC1+EXTRADESC2+EXTRAFER1+EXTRACOMP1) AS screditos, Sum(ATRASO+FALTA) AS sdebitos FROM corporerm.dbo.ABANCOHORFUN WHERE DATA<='" & dtaccess(dataant) & "' and data>='20120701' GROUP BY CHAPA " & _
"union select chapa, screditos=sum(extraant), sdebitos=sum(atrasoant+faltaant) from corporerm.dbo.asaldobancohor where inicioper='20120701' group by chapa " & _
") z group by chapa "

sqlmov="SELECT CHAPA, Sum(EXTRAFAIXA1+EXTRAFAIXA2+EXTRAFAIXA3+EXTRAFAIXA4+EXTRAFAIXA5+EXTRADESC1+EXTRADESC2+EXTRAFER1+EXTRACOMP1) AS creditos, Sum(ATRASO+FALTA) AS debitos FROM corporerm.dbo.ABANCOHORFUN WHERE DATA Between '" & dtaccess(datarel) & "' And '" & dtaccess(datarelf) & "' and data>='20120701' GROUP BY CHAPA"
sqlocor="select chapa, tbase=sum(base), thtrab=sum(htrab), tatraso=sum(atraso), tfalta=sum(falta), tabono=sum(abono), tadicional=sum(adicional), textra=sum(extraexecutado) from corporerm.dbo.AAFHTFUN where DATA Between '" & dtaccess(datarel) & "' And '" & dtaccess(datarelf) & "' GROUP BY CHAPA "

sql="SELECT d.CODSECAO, s.DESCRICAO, d.CHAPA, d.NOME, d.CODSITUACAO, d.CODSINDICATO, sa.screditos, sa.sdebitos, sm.creditos, sm.debitos, " & _
"tbase, thtrab, tatraso, tfalta, tabono, tadicional, textra " & _
"FROM ((((corporerm.dbo.PFUNC d INNER JOIN corporerm.dbo.APARFUN c ON d.CHAPA=c.CHAPA) INNER JOIN corporerm.dbo.PSECAO s ON d.CODSECAO=s.CODIGO) " & _
"LEFT JOIN (" & sqlocor & ") as h on h.chapa=d.chapa) " & _
"LEFT JOIN (" & sqlant & ") AS sa ON d.CHAPA=sa.CHAPA) LEFT JOIN (" & sqlmov & ") AS sm ON d.CHAPA=sm.CHAPA " & _
"WHERE d.CODSITUACAO in ('A','F','Z') AND d.CODSINDICATO<>'03' " & criterio & chapas & _
" and d.dataadmissao<'" & dtaccess(datarel) & "' " & _
"ORDER BY d.CODSECAO, d.NOME "

'response.write "<br>1 " & sqlant
'response.write "<br>2 " & sqlmov
'response.write "<br>3 " & sql
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border="0" cellpadding="2" width="999" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=campo align="left"  >Relatório Ponto</td>
	<td class=campo align="center">Ocorrências <%=monthname(month(datarel),0) & "/" & year(datarel)%> a <%=monthname(month(datarelf),0) & "/" & year(datarelf)%></td>
	<td class=campo align="right" ><%=now%> - Pág. <%pagina=pagina+1:response.write pagina%></td>
</tr>
</table>
<table border="0" cellpadding="1" width="999" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo rowspan=2 style="border: 1px solid #000000">Chapa</td>
	<td class=titulo rowspan=2 style="border: 1px solid #000000">Nome do Funcionário</td>
	<td class=titulo rowspan=2 style="border: 1px solid #000000">Sit.</td>
	<td class=titulo align="center" colspan=2 style="border: 1px solid #000000">Saldo anterior</td>
	<td class=titulo align="center" colspan=5 style="border: 1px solid #000000">Ocorrências</td>
	<td class=titulo align="center" colspan=3 style='border: 1px solid #000000'>Absenteismo</td>
	<td class=titulo align="center" colspan=2 style="border: 1px solid #000000">Saldo atual</td>
</tr>
<tr>
	<td class=titulo align="center" style="border: 1px solid #000000">Credor</td>
	<td class=titulo align="center" style="border: 1px solid #000000">Devedor</td>

	<td class=titulo align="center" style="border: 1px solid #000000">Atraso</td>
	<td class=titulo align="center" style="border: 1px solid #000000">Falta</td>
	<td class=titulo align="center" style="border: 1px solid #000000">Extra</td>
	<td class=titulo align="center" style="border: 1px solid #000000">Abono</td>
	<td class=titulo align="center" style="border: 1px solid #000000">Adic.N.</td>

	<td class=titulo align="center" style='border: 1px solid #000000'>Base</td>
	<td class=titulo align="center" style='border: 1px solid #000000'>H.Trab</td>
	<td class=titulo align="center" style='border: 1px solid #000000'>%</td>

	<td class=titulo align="center" style="border: 1px solid #000000">Credor</td>
	<td class=titulo align="center" style="border: 1px solid #000000">Devedor</td>
</tr>
<%
linha=3
rs.movefirst:do while not rs.eof
estilo="style='border-top: 1px solid #000000'"
estilo2="style='border-top: 1px solid #000000;border-right: 1px solid #000000'"
if linha>69 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<DIV style=""page-break-after:always""></DIV>"
	response.write "<table border='0' cellpadding='2' width='999' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=campo align=""left""  >Relatório de Ponto</td>"
	response.write "<td class=campo align=""center"">Ocorrências " & monthname(month(datarel),0) & "/" & year(datarel) & " a " & monthname(month(datarelf),0) & "/" & year(datarelf) & "</td>"
	response.write "<td class=campo align=""right"">" & now & " - Pág. " & pagina & "</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table border='0' cellpadding='1' width='999' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=titulo rowspan=2 style='border: 1px solid #000000'>Chapa</td>"
	response.write "<td class=titulo rowspan=2 style='border: 1px solid #000000'>Nome do Funcionário</td>"
	response.write "<td class=titulo rowspan=2 style='border: 1px solid #000000'>Sit.</td>"
	response.write "<td class=titulo align=""center"" colspan=2 style='border: 1px solid #000000'>Saldo anterior</td>"
	response.write "<td class=titulo align=""center"" colspan=5 style='border: 1px solid #000000'>Ocorrências</td>"
	response.write "<td class=titulo align=""center"" colspan=3 style='border: 1px solid #000000'>Absenteismo</td>"
	response.write "<td class=titulo align=""center"" colspan=2 style='border: 1px solid #000000'>Saldo atual</td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Credor</td>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Devedor</td>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Atraso</td>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Falta</td>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Extra</td>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Abono</td>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Adic.N.</td>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Base</td>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>H.Trab</td>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>%</td>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Credor</td>"
	response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Devedor</td>"
	response.write "</tr>"
	linha=3
end if
if lastsecao<>rs("codsecao") then
	if request.form("quebra")="on" and inicio=0 then
		pagina=pagina+1
		
	response.write "<tr>"
	response.write "<td height=20 class=titulo colspan=3 " & estilo2 & ">T O T A I S&nbsp;</td>"
	response.write "<td class=titulo " & estilo2 & " colspan=2 align=""center"">" & iif(santerior<0,"<font color=red>","") & horaload(abs(santerior),2) & "</td>"
	response.write "<td class=titulo " & estilo & " align=""center"">" & horaload(satraso,2) & "</td>"
	response.write "<td class=titulo " & estilo & " align=""center"">" & horaload(sfalta,2) & "</td>"
	response.write "<td class=titulo " & estilo2 & " align=""center"">" & horaload(sextra,2) & "</td>"
	response.write "<td class=titulo " & estilo & " align=""center"">" & horaload(sabono,2) & "</td>"
	response.write "<td class=titulo " & estilo2 & " align=""center"">" & horaload(sadicional,2) & "</td>"
	response.write "<td class=titulo " & estilo & " align=""center"">" & horaload(sbase,2) & "</td>"
	response.write "<td class=titulo " & estilo & " align=""center"">" & horaload(shtrab,2) & "</td>"
	if shtrab>sbase or (shtrab+sbase)=0 then sabsent=0 else sabsent=1-shtrab/sbase
	if sabsent>0 then sfabsent=formatpercent(sabsent,2) else sfabsent="&nbsp;"
	response.write "<td class=titulo " & estilo2 & " align=""center"">" & sfabsent & "</td>"
	response.write "<td class=titulo " & estilo2 & " colspan=2 align=""center"">" & iif(satual<0,"<font color=red>","") & horaload(abs(satual),2) & "</td>"
	response.write "</tr>"
	santerior=0:satraso=0:sfalta=0:sextra=0:sabono=0:sadicional=0:satual=0:sbase=0:shtrab=0
				tatraso=0:tfalta=0:textra=0:tabono=0:tadicional=0:tbase=0:thtrab=0

		response.write "</table>"
		response.write "<DIV style=""page-break-after:always""></DIV>"
		response.write "<table border='0' cellpadding='2' width='999' cellspacing='0' style='border-collapse: collapse'>"
		response.write "<tr>"
		response.write "<td class=campo align=""left""  >Relatório de Ponto</td>"
		response.write "<td class=campo align=""center"">Ocorrências " & monthname(month(datarel),0) & "/" & year(datarel) & " a " & monthname(month(datarelf),0) & "/" & year(datarelf) & "</td>"
		response.write "<td class=campo align=""right"">" & now & " - Pág. " & pagina & "</td>"
		response.write "</tr>"
		response.write "</table>"
		response.write "<table border='0' cellpadding='1' width='999' cellspacing='0' style='border-collapse: collapse'>"
		response.write "<tr>"
		response.write "<td class=titulo rowspan=2 style='border: 1px solid #000000'>Chapa</td>"
		response.write "<td class=titulo rowspan=2 style='border: 1px solid #000000'>Nome do Funcionário</td>"
		response.write "<td class=titulo rowspan=2 style='border: 1px solid #000000'>Sit.</td>"
		response.write "<td class=titulo align=""center"" colspan=2 style='border: 1px solid #000000'>Saldo anterior</td>"
		response.write "<td class=titulo align=""center"" colspan=5 style='border: 1px solid #000000'>Ocorrências</td>"
		response.write "<td class=titulo align=""center"" colspan=3 style='border: 1px solid #000000'>Absenteismo</td>"
		response.write "<td class=titulo align=""center"" colspan=2 style='border: 1px solid #000000'>Saldo atual</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Credor</td>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Devedor</td>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Atraso</td>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Falta</td>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Extra</td>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Abono</td>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Adic.N.</td>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Base</td>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>H.Trab</td>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>%</td>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Credor</td>"
		response.write "<td class=titulo align=""center"" style='border: 1px solid #000000'>Devedor</td>"
		response.write "</tr>"
		linha=3
	end if
	response.write "<tr>"
	response.write "<td class=grupo>" & rs("codsecao") & "</td>"
	response.write "<td class=grupo colspan=14>" & rs("descricao") & "</td>"
	response.write "</tr>"
	inicio=0
	linha=linha+1
end if
if isnull(rs("screditos")) then screditos=0 else screditos=rs("screditos")
if isnull(rs("sdebitos")) then sdebitos=0 else sdebitos=rs("sdebitos")
if isnull(rs("creditos")) then creditos=0 else creditos=rs("creditos")
if isnull(rs("debitos")) then debitos=0 else debitos=rs("debitos")
saldo_ant=screditos-sdebitos
if saldo_ant<0 then fator=-1 else fator=1
if saldo_ant<0 then saldo_ant=abs(saldo_ant)
saldoant=horaload(saldo_ant,2)
mcreditos=horaload(creditos,2)
mdebitos=horaload(debitos,2)
saldo_atu=screditos-sdebitos+creditos-debitos
if saldo_atu<0 then fatorf=-1 else fatorf=1
if saldo_atu<0 then saldo_atu=abs(saldo_atu)
saldoatu=horaload(saldo_atu,2)
linha=linha+1
if creditos=0 then mcreditos="&nbsp;"
if debitos=0 then mdebitos="&nbsp;"
tatraso=0:tfalta=0:tabono=0:textra=0:tadicional=0:tbase=0:thtrab=0
if isnull(rs("tatraso"))    then tatraso=0    else tatraso=rs("tatraso")       :tatraso1=horaload(tatraso,2)
if isnull(rs("tfalta"))     then tfalta=0     else tfalta=rs("tfalta")         :tfalta1 =horaload(tfalta,2)
if isnull(rs("tabono"))     then tabono=0     else tabono=rs("tabono")         :tabono1 =horaload(tabono,2)
if isnull(rs("textra"))     then textra=0     else textra=rs("textra")         :textra1 =horaload(textra,2)
if isnull(rs("tadicional")) then tadicional=0 else tadicional=rs("tadicional") :tadicional1=horaload(tadicional,2)
if isnull(rs("tbase"))      then tbase=0      else tbase=rs("tbase")           :tbase2=tbase :tbase1=horaload(tbase,2)
if isnull(rs("thtrab"))     then thtrab=0     else thtrab=rs("thtrab")         :thtrab2=thtrab :thtrab1 =horaload(thtrab,2)
santerior=santerior+screditos-sdebitos
satraso=satraso+tatraso
sfalta=sfalta+tfalta
sabono=sabono+tabono
sextra=sextra+textra
sadicional=sadicional+tadicional
sbase=sbase+tbase
shtrab=shtrab+thtrab
satual=satual+screditos-sdebitos+creditos-debitos
%>
<tr>
	<td height=20 class=campo <%=estilo%>><%=rs("chapa")%>&nbsp;</td>
	<td class=campo <%=estilo%>><%=rs("nome")%></td>
	<td class=campo <%=estilo2%>><%=rs("codsituacao")%></td>
	<td class=campo <%=estilo%> align="center"><%if fator=1 then response.write saldoant%></td>
	<td class=campo <%=estilo2%> align="center"><font color=red><%if fator=-1 then response.write saldoant%></td>
	
	<td class=campo <%=estilo%> align="center"><%=tatraso1%></td>
	<td class=campo <%=estilo%> align="center"><%=tfalta1%></td>
	<td class=campo <%=estilo2%> align="center"><%=textra1%></td>
	<td class=campo <%=estilo%> align="center"><%=tabono1%></td>
	<td class=campo <%=estilo2%> align="center"><%=tadicional1%></td>
	
	<td class=campo <%=estilo%> align="center"><%=tbase1%></td>
	<td class=campo <%=estilo%> align="center"><%=thtrab1%></td>
<%if thtrab2>tbase2 or (thtrab2+tbase2)=0 then absent=0 else absent=1-thtrab2/tbase2
if absent>0 then fabsent=formatpercent(absent,2) else fabsent="&nbsp;"%>
	<td class=campo <%=estilo2%> align="center"><%=fabsent%></td>

	<td class=campo <%=estilo%> align="center"><%if fatorf=1 then response.write saldoatu%></td>
	<td class=campo <%=estilo2%> align="center"><font color=red><%if fatorf=-1 then response.write saldoatu%></td>

</tr>
<%
lastsecao=rs("codsecao")
tatraso1="":tfalta1="":tabono1="":textra1="":tadicional1="":tbase1="":thtrab1="":
rs.movenext:loop
rs.close
%>
<tr>
	<td height=20 class=titulo colspan=3 <%=estilo2%>>T O T A I S&nbsp;</td>
	<td class=titulo <%=estilo2%> colspan=2 align="center"><%if santerior<0 then response.write "<font color=red>"%><%=horaload(abs(santerior),2)%></td>
	
	<td class=titulo <%=estilo%> align="center"><%=horaload(satraso,2)%></td>
	<td class=titulo <%=estilo%> align="center"><%=horaload(sfalta,2)%></td>
	<td class=titulo <%=estilo2%> align="center"><%=horaload(sextra,2)%></td>
	<td class=titulo <%=estilo%> align="center"><%=horaload(sabono,2)%></td>
	<td class=titulo <%=estilo2%> align="center"><%=horaload(sadicional,2)%></td>

	<td class=titulo <%=estilo%> align="center"><%=horaload(sbase,2)%></td>
	<td class=titulo <%=estilo%> align="center"><%=horaload(shtrab,2)%></td>
<%if shtrab>sbase or (shtrab+sbase)=0 then sabsent=0 else sabsent=1-shtrab/sbase
if sabsent>0 then sfabsent=formatpercent(sabsent,2) else sfabsent="&nbsp;"%>
	<td class=titulo <%=estilo2%> align="center"><%=sfabsent%></td>
	
	<td class=titulo <%=estilo2%> colspan=2 align="center"><%if satual<0 then response.write "<font color=red>"%><%=horaload(abs(satual),2)%></td>
</tr>


</table>

<%
end if 'request.form
'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>