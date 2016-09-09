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
<title>Relação de Atestados Médicos</title>
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
data1=now()
datai=dateserial(year(data1),month(data1)-1,1)
dataf=dateserial(year(data1),month(data1),1)-1
sql="SELECT d.CODSECAO, secao as descricao " & _
"FROM atestados AS c, qry_funcionarios d " & _
"WHERE c.CHAPA = d.CHAPA collate database_default and d.CODSITUACAO<>'D' AND d.CODSINDICATO<>'03' " & _
"GROUP BY d.CODSECAO, secao " & _
"ORDER BY secao, d.CODSECAO "
%>
<form method="POST" action="atestado_rel1.asp" name="form">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=420>
<tr>
	<td class=titulo colspan=3>Relação de Atestados Médicos</td>
</tr>
<tr>
	<td class=grupo>Data Inicial</td>
	<td class=grupo>Data Final</td>
	<td class=grupo>Opções</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="dti" size="8" value="<%=datai%>"></td>
	<td class=titulo><input type="text" name="dtf" size="8" value="<%=dataf%>"></td>
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
	<td class=grupo colspan=3>Tipo</td>
</tr>
<tr>
	<td class=titulo colspan=3>
	<select size="1" name="tipo">
	<option value="0">Todos Funcionários</option>
	<option value="1">Administrativos</option>
	<option value="2">Professores</option>
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

if request.form("setor")="0" then 
	criterio1="" 
	criterio2="" 
	criterio3="" 
else 
	criterio1=" AND f.codsecao='" & request.form("setor") & "' "
	criterio2=" AND f.codsecao='" & request.form("setor") & "' "
	criterio3=" AND f.codsecao='" & request.form("setor") & "' "
end if
if request.form("tipo")="0" then tipo=""
if request.form("tipo")="1" then tipo=" and codsindicato<>'03' "
if request.form("tipo")="2" then tipo=" and codsindicato='03' "

sql="SELECT a.chapa, a.data1, a.data2, a.dias, a.cid, a.crm, a.medico, a.clinica, a.parcial " & _
"FROM atestados a, pfunc f " & _
"WHERE (a.data1 Between '" & dtaccess(datarel) & "' And '" & dtaccess(datarelf) & "' AND a.data2 Between '" & dtaccess(datarel) & "' And '" & dtaccess(datarelf) & "') " & _
"OR (a.data2>='" & dtaccess(datarel) & "') OR (a.data1<='" & dtaccess(datarelf) & "' AND a.data2>'" & dtaccess(datarelf) & "') " & _
"AND a.chapa=f.chapa " & _
"ORDER BY a.chapa "
sql="SELECT a.chapa, f.NOME, f.codsecao, f.codsituacao, secao, funcao, a.data1, a.data2, a.dias, a.cid, a.crm, a.medico, a.clinica, a.parcial " & _
"FROM atestados a INNER JOIN (select chapa, nome, codsecao, secao, codsituacao, codfuncao, funcao, codsindicato from qry_funcionarios where codsituacao<>'D' " & tipo & " ) f " & _
"ON a.chapa=f.CHAPA collate database_default " & _
"WHERE (a.data1 Between '" & dtaccess(datarel) & "' And '" & dtaccess(datarelf) & "' AND a.data2 Between '" & dtaccess(datarel) & "' And '" & dtaccess(datarelf) & "' " & criterio1 & ") " & _
"OR (a.data1<='" & dtaccess(datarelf) & "' AND a.data2>='" & dtaccess(datarel) & "' " & criterio2 & ") " & _
"OR (a.data1<='" & dtaccess(datarelf) & "' AND a.data2>'" & dtaccess(datarelf) & "' " & criterio3 & ") " & _
"ORDER BY f.codsecao, a.chapa "
'response.write sql
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border="0" cellpadding="2" width="990" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campor" align="left"  >Controle de Atestados</td>
	<td class="campor" align="center">Justificativas/Abonos <%=monthname(month(datarel),0) & "/" & year(datarel)%> </td>
	<td class="campor" align="right" ><%=now%> - Pág. <%pagina=pagina+1:response.write pagina%></td>
</tr>
</table>
<table border="0" cellpadding="1" width="990" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulor rowspan=2 style="border: 1px solid #000000">Chapa</td>
	<td class=titulor rowspan=2 style="border: 1px solid #000000">Nome do Funcionário</td>
	<td class=titulor rowspan=2 style="border: 1px solid #000000">Sit.</td>
	<td class=titulor align="center" colspan=2 style="border: 1px solid #000000">Período Afastamento</td>
	<td class=titulor align="center" colspan=2 style="border: 1px solid #000000">Número de Dias</td>
	<td class=titulor align="center" colspan=3 style="border: 1px solid #000000">Detalhes</td>
</tr>
<tr>
	<td class=titulor align="center" style="border: 1px solid #000000">Inicio</td>
	<td class=titulor align="center" style="border: 1px solid #000000">Término</td>
	<td class=titulor align="center" style="border: 1px solid #000000">Dias</td>
	<td class=titulor align="center" style="border: 1px solid #000000">Horas</td>
	<td class=titulor align="center" style="border: 1px solid #000000">CID</td>
	<td class=titulor align="center" style="border: 1px solid #000000">Clinica</td>
	<td class=titulor align="center" style="border: 1px solid #000000">Médico</td>
</tr>
<%
linha=3
rs.movefirst:do while not rs.eof
estilo="style='border-top: 1px solid #000000'"
estilo2="style='border-top: 1px solid #000000;border-right: 1px solid #000000'"
if linha>46 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<DIV style=""page-break-after:always""></DIV>"
	response.write "<table border='0' cellpadding='2' width='990' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=""campor"" align=""left""  >Controle de Atestados</td>"
	response.write "<td class=""campor"" align=""center"">Justificativas/Abonos</td>"
	response.write "<td class=""campor"" align=""right"">" & now & " - Pág. " & pagina & "</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table border='0' cellpadding='1' width='990' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=""titulor"" rowspan=2 style='border: 1px solid #000000'>Chapa</td>"
	response.write "<td class=""titulor"" rowspan=2 style='border: 1px solid #000000'>Nome do Funcionário</td>"
	response.write "<td class=""titulor"" rowspan=2 style='border: 1px solid #000000'>Sit.</td>"
	response.write "<td class=""titulor"" align=""center"" colspan=2 style='border: 1px solid #000000'>Período Afastamento</td>"
	response.write "<td class=""titulor"" align=""center"" colspan=2 style='border: 1px solid #000000'>Número de Dias</td>"
	response.write "<td class=""titulor"" align=""center"" colspan=3 style='border: 1px solid #000000'>Detalhes</td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Início</td>"
	response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Término</td>"
	response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Dias</td>"
	response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Horas</td>"
	response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>CID</td>"
	response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Clínica</td>"
	response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Médico</td>"
	response.write "</tr>"
	linha=3
end if
if lastsecao<>rs("codsecao") then
	if request.form("quebra")="on" and inicio=0 then
		pagina=pagina+1
		response.write "</table>"
		response.write "<DIV style=""page-break-after:always""></DIV>"
		response.write "<table border='0' cellpadding='2' width='990' cellspacing='0' style='border-collapse: collapse'>"
		response.write "<tr>"
		response.write "<td class=""campor"" align=""left""  >Controle de Atestados</td>"
		response.write "<td class=""campor"" align=""center"">Justificativas/Abonos " & monthname(month(datarel),0) & "/" & year(datarel) &  "</td>"
		response.write "<td class=""campor"" align=""right"">" & now & " - Pág. " & pagina & "</td>"
		response.write "</tr>"
		response.write "</table>"
		response.write "<table border='0' cellpadding='1' width='990' cellspacing='0' style='border-collapse: collapse'>"
		response.write "<tr>"
		response.write "<td class=""titulor"" rowspan=2 style='border: 1px solid #000000'>Chapa</td>"
		response.write "<td class=""titulor"" rowspan=2 style='border: 1px solid #000000'>Nome do Funcionário</td>"
		response.write "<td class=""titulor"" rowspan=2 style='border: 1px solid #000000'>Sit.</td>"
		response.write "<td class=""titulor"" align=""center"" colspan=2 style='border: 1px solid #000000'>Período Afastamento</td>"
		response.write "<td class=""titulor"" align=""center"" colspan=2 style='border: 1px solid #000000'>Número de Dias</td>"
		response.write "<td class=""titulor"" align=""center"" colspan=3 style='border: 1px solid #000000'>Detalhes</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Início</td>"
		response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Término</td>"
		response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Dias</td>"
		response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Horas</td>"
		response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>CID</td>"
		response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Clínica</td>"
		response.write "<td class=""titulor"" align=""center"" style='border: 1px solid #000000'>Médico</td>"
		response.write "</tr>"
		linha=3
	end if
	response.write "<tr>"
	response.write "<td class=grupo>" & rs("codsecao") & "</td>"
	response.write "<td class=grupo colspan=9>" & rs("secao") & "</td>"
	response.write "</tr>"
	inicio=0
	linha=linha+1
end if

linha=linha+1
valor1=rs("parcial")
hora=int(valor1)
minuto=int((valor1-hora)*100)
vhoras=hora&":"&numzero(minuto,2)
if cdbl(rs("parcial"))>0 then dias="&nbsp;" else dias=rs("dias")
if cdbl(rs("parcial"))=0 then vhoras="&nbsp;"
%>
<tr>
	<td class="campor" <%=estilo%>  align="left"><%=rs("chapa")%>&nbsp;</td>
	<td class="campor" <%=estilo%>  align="left"><%=rs("nome")%></td>
	<td class="campor" <%=estilo%>  align="left"><%=rs("codsituacao")%></td>
	<td class="campor" <%=estilo%>  align="center"><%=rs("data1")%></td>
	<td class="campor" <%=estilo2%> align="center"><%=rs("data2")%></td>
	<td class="campor" <%=estilo%>  align="center"><%=dias%></td>
	<td class="campor" <%=estilo2%> align="center"><%=vhoras%></td>
	<td class="campor" <%=estilo%>  align="left"><%=rs("cid")%></td>
	<td class="campor" <%=estilo2%> align="left"><%=rs("clinica")%></td>
	<td class="campor" <%=estilo2%> align="left"><%=rs("medico")%></td>
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