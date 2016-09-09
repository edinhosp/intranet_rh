<%@ Language=VBScript %>
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
<title>Relação de Saldo de Banco de Horas</title>
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
<form method="POST" action="bancohoras.asp" name="form">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=420>
<tr>
	<td class=titulo colspan=3>Relação de Saldo de Banco de Horas</td>
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

sqlt="INSERT INTO ABANCOHORFUN (CODCOLIGADA,CHAPA,DATA,EXTRAFAIXA1,EXTRAFAIXA2,EXTRAFAIXA3,EXTRAFAIXA4,EXTRAFAIXA5,EXTRADESC1,EXTRADESC2,EXTRAFER1,EXTRAFER2,EXTRACOMP1,EXTRACOMP2,FALTA,ATRASO) " & _
"SELECT DISTINCTROW aparfun.CODCOLIGADA, aparfun.CHAPA, '" & dtaccess(datarel) & "', 0,0,0,0,0,0,0,0,0,0,0,0,0 " & _
"FROM APARFUN INNER JOIN PFUNC ON APARFUN.CHAPA = PFUNC.CHAPA WHERE CODSITUACAO<>'D' AND CODSINDICATO<>'03' "
'response.write sqlt

if request.form("setor")="0" then criterio="" else criterio=" AND d.codsecao='" & request.form("setor") & "' "

sqlant="select chapa, screditos=sum(screditos), sdebitos=sum(sdebitos) from (" & _
"SELECT CHAPA, Sum(EXTRAFAIXA1+EXTRAFAIXA2+EXTRAFAIXA3+EXTRAFAIXA4+EXTRAFAIXA5+EXTRADESC1+EXTRADESC2+EXTRAFER1+EXTRAFER2+EXTRACOMP1+EXTRACOMP2) AS screditos, Sum(ATRASO+FALTA) AS sdebitos FROM corporerm.dbo.ABANCOHORFUN WHERE DATA<='" & dtaccess(dataant) & "' and data>='20120701' GROUP BY CHAPA " & _
"union select chapa, screditos=sum(extraant), sdebitos=sum(atrasoant+faltaant) from corporerm.dbo.asaldobancohor where inicioper='20120701' group by chapa " & _
") z group by chapa "

sqlmov="SELECT CHAPA, Sum(EXTRAFAIXA1+EXTRAFAIXA2+EXTRAFAIXA3+EXTRAFAIXA4+EXTRAFAIXA5+EXTRADESC1+EXTRADESC2+EXTRAFER1+EXTRAFER2+EXTRACOMP1+EXTRACOMP2) AS creditos, Sum(ATRASO+FALTA) AS debitos FROM corporerm.dbo.ABANCOHORFUN WHERE DATA Between '" & dtaccess(datarel) & "' And '" & dtaccess(datarelf) & "' and data>='20120701' GROUP BY CHAPA"

sql="SELECT d.CODSECAO, s.DESCRICAO, d.CHAPA, d.NOME, d.CODSITUACAO, d.CODSINDICATO, sa.screditos, sa.sdebitos, sm.creditos, sm.debitos " & _
"FROM corporerm.dbo.APARFUN AS c, corporerm.dbo.PFUNC AS d, corporerm.dbo.PSECAO s, (" & sqlant & ") as sa, (" & sqlmov & ") as sm  " & _
"WHERE sa.chapa=d.chapa and sm.chapa=d.chapa and c.CHAPA = d.CHAPA and d.CODSECAO = s.CODIGO and d.CODSITUACAO<>'D' AND d.CODSINDICATO<>'03' " & criterio & _
"ORDER BY d.CODSECAO, d.NOME "
sql="SELECT d.CODSECAO, s.DESCRICAO, d.CHAPA, d.NOME, d.CODSITUACAO, d.CODSINDICATO, sa.screditos, sa.sdebitos, sm.creditos, sm.debitos " & _
"FROM (((corporerm.dbo.PFUNC AS d INNER JOIN corporerm.dbo.APARFUN AS c ON d.CHAPA = c.CHAPA) INNER JOIN corporerm.dbo.PSECAO AS s ON d.CODSECAO = s.CODIGO) LEFT JOIN (" & sqlant & ") AS sa ON d.CHAPA = sa.CHAPA) LEFT JOIN (" & sqlmov & ") AS sm ON d.CHAPA = sm.CHAPA " & _
"WHERE d.CODSITUACAO in ('A','F','Z') AND d.CODSINDICATO<>'03' " & criterio & _
" and d.chapa not in ('00099','02538','00554','02653','00822','02297','00093') " & _
"ORDER BY d.CODSECAO, d.NOME "

'response.write "<br>1 " & sqlant
'response.write "<br>2 " & sqlmov
'response.write "<br>3 " & sql
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border="0" cellpadding="2" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campor" align="left"  >Controle de Saldo</td>
	<td class="campor" align="center">Banco de Horas <%=monthname(month(datarel),0) & "/" & year(datarel)%> </td>
	<td class="campor" align="right" ><%=now%> - Pág. <%pagina=pagina+1:response.write pagina%></td>
</tr>
</table>
<table border="0" cellpadding="1" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulor rowspan=2 style="border: 1px solid #000000">Chapa</td>
	<td class=titulor rowspan=2 style="border: 1px solid #000000">Nome do Funcionário</td>
	<td class=titulor rowspan=2 style="border: 1px solid #000000">Sit.</td>
	<td class=titulor align="center" colspan=2 style="border: 1px solid #000000">Saldo anterior</td>
	<td class=titulor align="center" colspan=2 style="border: 1px solid #000000">Movimento do mês</td>
	<td class=titulor align="center" colspan=2 style="border: 1px solid #000000">Saldo atual</td>
</tr>
<tr>
	<td class=titulor align="center" style="border: 1px solid #000000">Credor</td>
	<td class=titulor align="center" style="border: 1px solid #000000">Devedor</td>
	<td class=titulor align="center" style="border: 1px solid #000000">Créditos</td>
	<td class=titulor align="center" style="border: 1px solid #000000">Débitos</td>
	<td class=titulor align="center" style="border: 1px solid #000000">Credor</td>
	<td class=titulor align="center" style="border: 1px solid #000000">Devedor</td>
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
	response.write "<table border='0' cellpadding='2' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=""campor"" align=""left""  >Controle de Saldo</td>"
	response.write "<td class=""campor"" align=""center"">Banco de Horas</td>"
	response.write "<td class=""campor"" align=""right"">" & now & " - Pág. " & pagina & "</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table border='0' cellpadding='1' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=titulor rowspan=2 style='border: 1px solid #000000'>Chapa</td>"
	response.write "<td class=titulor rowspan=2 style='border: 1px solid #000000'>Nome do Funcionário</td>"
	response.write "<td class=titulor rowspan=2 style='border: 1px solid #000000'>Sit.</td>"
	response.write "<td class=titulor align=""center"" colspan=2 style='border: 1px solid #000000'>Saldo anterior</td>"
	response.write "<td class=titulor align=""center"" colspan=2 style='border: 1px solid #000000'>Movimento do mês</td>"
	response.write "<td class=titulor align=""center"" colspan=2 style='border: 1px solid #000000'>Saldo atual</td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Credor</td>"
	response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Devedor</td>"
	response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Créditos</td>"
	response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Débitos</td>"
	response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Credor</td>"
	response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Devedor</td>"
	response.write "</tr>"
	linha=3
end if
if lastsecao<>rs("codsecao") then
	if request.form("quebra")="on" and inicio=0 then
		pagina=pagina+1
		response.write "</table>"
		response.write "<DIV style=""page-break-after:always""></DIV>"
		response.write "<table border='0' cellpadding='2' width='690' cellspacing='0' style='border-collapse: collapse'>"
		response.write "<tr>"
		response.write "<td class=""campor"" align=""left""  >Controle de Saldo</td>"
		response.write "<td class=""campor"" align=""center"">Banco de Horas " & monthname(month(datarel),0) & "/" & year(datarel) &  "</td>"
		response.write "<td class=""campor"" align=""right"">" & now & " - Pág. " & pagina & "</td>"
		response.write "</tr>"
		response.write "</table>"
		response.write "<table border='0' cellpadding='1' width='690' cellspacing='0' style='border-collapse: collapse'>"
		response.write "<tr>"
		response.write "<td class=titulor rowspan=2 style='border: 1px solid #000000'>Chapa</td>"
		response.write "<td class=titulor rowspan=2 style='border: 1px solid #000000'>Nome do Funcionário</td>"
		response.write "<td class=titulor rowspan=2 style='border: 1px solid #000000'>Sit.</td>"
		response.write "<td class=titulor align=""center"" colspan=2 style='border: 1px solid #000000'>Saldo anterior</td>"
		response.write "<td class=titulor align=""center"" colspan=2 style='border: 1px solid #000000'>Movimento do mês</td>"
		response.write "<td class=titulor align=""center"" colspan=2 style='border: 1px solid #000000'>Saldo atual</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Credor</td>"
		response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Devedor</td>"
		response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Créditos</td>"
		response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Débitos</td>"
		response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Credor</td>"
		response.write "<td class=titulor align=""center"" style='border: 1px solid #000000'>Devedor</td>"
		response.write "</tr>"
		linha=3
	end if
	response.write "<tr>"
	response.write "<td class=grupo>" & rs("codsecao") & "</td>"
	response.write "<td class=grupo colspan=8>" & rs("descricao") & "</td>"
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
sa_hora=int(saldo_ant/60)
sa_minuto=saldo_ant-(sa_hora*60)
saldoant=sa_hora & ":" & numzero(sa_minuto,2)

saldo_mes=creditos-debitos
mcreditos=int(creditos/60) & ":" & numzero((creditos-(int(creditos/60)*60)),2)
mdebitos=int(debitos/60) & ":" & numzero((debitos-(int(debitos/60)*60)),2)

saldo_atu=screditos-sdebitos+creditos-debitos
if saldo_atu<0 then fatorf=-1 else fatorf=1
if saldo_atu<0 then saldo_atu=abs(saldo_atu)
st_hora=int(saldo_atu/60)
st_minuto=saldo_atu-(st_hora*60)
saldoatu=st_hora & ":" & numzero(st_minuto,2)
linha=linha+1
if creditos=0 then mcreditos="&nbsp;"
if debitos=0 then mdebitos="&nbsp;"
%>
<tr>
	<td class="campor" <%=estilo%>><%=rs("chapa")%>&nbsp;</td>
	<td class="campor" <%=estilo%>><%=rs("nome")%></td>
	<td class="campor" <%=estilo%>><%=rs("codsituacao")%></td>
	<td class="campor" <%=estilo%> align="right"><%if fator=1 then response.write saldoant%>&nbsp;</td>
	<td class="campor" <%=estilo2%> align="right"><font color=red><%if fator=-1 then response.write saldoant%>&nbsp;</td>
	<td class="campor" <%=estilo%> align="right"><%=mcreditos%>&nbsp;</td>
	<td class="campor" <%=estilo2%> align="right"><%=mdebitos%>&nbsp;</td>
	<td class="campor" <%=estilo%> align="right"><%if fatorf=1 then response.write saldoatu%>&nbsp;</td>
	<td class="campor" <%=estilo2%> align="right"><font color=red><%if fatorf=-1 then response.write saldoatu%>&nbsp;</td>
</tr>
<%
lastsecao=rs("codsecao")
rs.movenext:loop
rs.close
%>
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