<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a46")="N" or session("a46")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Controle de Falta de Marcações e Justificativa</title>
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
	set rs2=server.createobject ("ADODB.Recordset")
	Set rs2.ActiveConnection = conexao
	set rs3=server.createobject ("ADODB.Recordset")
	Set rs3.ActiveConnection = conexao
	
if request.form="" then
%>
<p class=titulo>Verificação da Quantidade de Falta de Marcações
<form method="POST" action="n4_faltamarcacao.asp">
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo>Verificar Total de faltas de marcações entre</td></tr>
<%
hoje=int(now())
diasem=weekday(hoje)
d2=hoje - (diasem-1)
d1=d2-6
%>
<tr>
	<td class=titulo>de <input type="text" name="d1" value="<%=d1%>" size="9"> até <input type="text" name="d2" value="<%=d2%>" size="9"></td>
</tr>
<tr><td class=titulo>
	<input type="text" value="" size="5" maxlength="5" name="ch1">
	<input type="text" value="" size="5" maxlength="5" name="ch2">
	<input type="text" value="" size="5" maxlength="5" name="ch3">
	<input type="text" value="" size="5" maxlength="5" name="ch4">
	<input type="text" value="" size="5" maxlength="5" name="ch5">
</td></tr>

<tr><td colspan=3 class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3">
</td></tr>
</table>
</form>
<hr>
<%
else 'request.form <>''

datai=cdate(request.form("d1"))
dataf=cdate(request.form("d2"))
linha=0:pagina=0

teste=1

ch1=request.form("ch1"):ch2=request.form("ch2"):ch3=request.form("ch3"):ch4=request.form("ch4"):ch5=request.form("ch5")
if ch1<>"" or ch2<>"" or ch3<>"" or ch4<>"" or ch5<>"" then
	chapas=" and f.chapa in ("
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

sqld="select distinct a.chapa, f.nome, vezes=count(tipo), f.codsecao, f.secao, f.codhorario, h.DESCRICAO, f.sexo, f.email " & _
"from corporerm.dbo.aafdt a inner join qry_funcionarios f on f.chapa=a.chapa " & _
"inner join corporerm.dbo.AHORARIO h on h.CODIGO=f.codhorario " & _
"where convert(date,datahora) between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "' and tipo='I' " & _
"and f.codsituacao<>'D'  " & chapas & " and a.justificativa not in ('Folga NR-17','Marcação Teste','Acerto de Marcação') " & _
"group by a.chapa, f.nome, f.codsecao, f.secao, f.codhorario, h.DESCRICAO, f.sexo, f.email " & _
"order by f.codsecao, f.nome "
rs.Open sqld, ,adOpenStatic, adLockReadOnly
totalpag=int(rs.recordcount/65)+1
do while not rs.eof
if linha=0 then 'or linha>64 then
	'if linha<>0 then
	'	pagina=pagina+1
	'	response.write "<tr><td class="campor" colspan=7 style='border-top:1px solid #000000'>Página " & pagina & "/" & totalpag & " - " & now() & "</td></tr>"
	'	response.write "</table>"
	'	response.write "<DIV style=""page-break-after:always""></DIV>"
	'end if
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td class=titulo colspan=7 align="center">Relatório do Total de Falta de Marcações - De <%=datai%> a <%=dataf%></td></td>
<tr>
	<td class=titulo align="center">Funcionário</td>
	<td class=titulo align="center">#</td>
	<td class=titulo align="center">Justificativa</td>
	<td class=titulo align="center"></td>
</tr>
<%
	if linha<>0 then linha=0
end if 'linha
if rs("chapa")<>ultchapa then cab=1 else cab=0
'obs=rs.absoluteposition & "-" & obs 
%>
<tr>
<%
if cab=1 then estilo="border-top:1px solid #000000" else estilo="border-top:0px solid #000000"
%>
	<td class=campo style="<%=estilo%>" valign="top" >
	<%=rs("chapa")%> - <b><%=rs("nome")%></b><br>
	<%=rs("codsecao")%> - <%=rs("secao")%><br>
	<%=rs("descricao")%>
	</td>
	<td class=campo style="<%=estilo%><%=";border-left:1px dotted #000000"%>" valign="top" align="center" ><%=rs("vezes")%></td>
	
	<td class=campo style="<%=estilo%>" valign="top" >

	<!-- quadro dos dias com marcações incompletas -->
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<!--
<tr>
	<td class=titulo>Data</td>
	<td class=titulo>Hora</td>
	<td class=titulo>Justificativa</td>
	<td class=titulo></td>
</tr>
-->
<%
sql2="select distinct a.chapa, data=convert(date,datahora), diasem=datepart(dw,convert(date,datahora)), hora=convert(time,datahora), tipo, justificativa, datahora " & _
"from corporerm.dbo.aafdt a " & _
"where convert(date,datahora) between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "' and tipo='I' and a.chapa='" & rs("chapa") & "' " & _
"and a.justificativa not in ('Folga NR-17','Marcação Teste','Acerto de Marcação') " & _
"order by datahora "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof 
%>
<tr>
	<td class=campo width="70px" align="center"><%=cdate(rs2("data"))%></td>
	<td class=campo width="30px" align="center"><%=(left(rs2("hora"),5))%></td>
	<td class="campor" width="120px" align="left" nowrap><%=rs2("justificativa")%></td>
	<td class=campo>
<%
sql3="select batida from corporerm.dbo.abatfun where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs2("data")) & "' order by batida"
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
rs3.movefirst
do while not rs3.eof
	batida=rs3("batida")
	hora=int(batida/60)
	minuto=batida-(hora*60)
	temp=numzero(hora,2) & ":" & numzero(minuto,2)
	'response.write temp
	'if rs3.absoluteposition<rs3.recordcount then response.write " - "
rs3.movenext
loop
else
	'response.write "-"
end if
rs3.close
%>
	</td>
</tr>	
<%
%>

<%
rs2.movenext
loop
end if 'rs2.recordcount>0
rs2.close
%>
	</table>
<!-- final do quadro dos dias com marcações incompletas -->	
	</td>
	<td class=campo style="<%=estilo%>" valign="top" >
	<a href="n4_email.asp?chapa=<%=rs("chapa")%>&datai=<%=datai%>&dataf=<%=dataf%>" onclick="NewWindow(this.href,'Selecao_email','690','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/email_go.png" border="0" width="15" alt="Enviar•Email"></a>
	<br>
	<a href="n4_print.asp?chapa=<%=rs("chapa")%>&datai=<%=datai%>&dataf=<%=dataf%>" onclick="NewWindow(this.href,'Selecao_print','690','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/printer.gif" border="0" width="15" alt="Enviar•Impressora"></a>
	</td>
</tr>

<%
linha=linha+1
ultchapa=rs("chapa")
rs.movenext
loop
rs.close
pagina=pagina+1
%>
<tr><td class="campor" colspan=6 style='border-top:1px solid #000000'>Página <%=pagina & "/" & totalpag%> - <%=now()%></td></tr>
</table>

<%
end if ' request.form	
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>