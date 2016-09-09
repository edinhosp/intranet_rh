<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 60000
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a46")="N" or session("a46")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Falta de Marcações e Justificativa</title>
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
<p class=titulo>Verificação de Falta de Marcações
<form method="POST" action="n3_justificativa.asp">
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo>Verificar faltas de marcações entre</td></tr>
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
sql1="delete from _marcacoes_checagem "
if teste=1 then conexao.Execute sql1, , adCmdText

sql1="insert into _marcacoes_checagem (chapa, data) " & _
"select a.chapa, a.data from (  " & _
"	select chapa, data from corporerm.dbo.ABATFUN group by CHAPA, data having count(batida) not in (2,4,6,8)  " & _
") a  " & _
"inner join corporerm.dbo.pfunc f on f.chapa=a.chapa  " & _
"where f.codsindicato<>'03' and f.codhorario not in ('183','186','184','02398','02540','251','00818','02466','00357','00869','02193','00445','02329','02449','02196','02345','02194','00865','02121','280','02468')  " & _
"and a.data between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "'  " & _
"group by a.chapa, a.data  " 
'if teste=1 then conexao.Execute sql1, , adCmdText

sql1="insert into _marcacoes_checagem (chapa, data)  " & _
"select a.chapa, a.data  " & _
"from corporerm.dbo.abatfun a inner join corporerm.dbo.pfunc f on f.chapa=a.chapa " & _
"where f.codsindicato<>'03' and f.codhorario not in ('Diretores') and f.jornadamensal/60>=180 and datepart(dw,a.data)<>7  " & _
"and a.data between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "'  " & _
"group by a.chapa, a.data  " & _
"having count(a.batida)=2 and sum( (case when natureza=0 or natureza=4 then -1 else 1 end) *batida)>420  "
nome=session("usuariomaster")

sql1="insert into _marcacoes_checagem (chapa, data)  " & _
"select a.chapa, a.data  " & _
"from n3_justificativa_s1 a " 
sql2="delete from _marcacoes_checagem where data not between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "' "
if teste=1 then conexao.Execute sql1, , adCmdText
if teste=1 then conexao.Execute sql2, , adCmdText

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

sqld="select distinct z.chapa, f.nome, f.codsecao, f.secao, f.codhorario, h.DESCRICAO, f.sexo, f.email " & _
"from ( " & _
"	select a.chapa, a.data from _marcacoes_checagem a  " & _
"	where a.data between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "'  " & _
"	group by a.chapa, a.data " & _
") as z " & _
"inner join qry_funcionarios f on f.chapa collate database_default=z.chapa " & _
"inner join corporerm.dbo.AHORARIO h on h.CODIGO=f.codhorario " & _
"where f.codsituacao<>'D'  " & chapas & _
"order by f.codsecao, f.nome  "
rs.Open sqld, ,adOpenStatic, adLockReadOnly
totalpag=int(rs.recordcount/65)+1
do while not rs.eof
if linha=0 or linha>64 then
	if linha<>0 then
		pagina=pagina+1
		response.write "<tr><td class=""campor"" colspan=7 style='border-top:1px solid #000000'>Página " & pagina & "/" & totalpag & " - " & now() & "</td></tr>"
		response.write "</table>"
		response.write "<DIV style=""page-break-after:always""></DIV>"
	end if
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td class=titulo colspan=7 align="center">Relatório de Falta de Marcações - De <%=datai%> a <%=dataf%></td></td>
<tr>
	<td class=titulo align="center">Funcionário</td>
	<td class=titulo align="center">Marcações</td>
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
	<%=rs("codhorario")%> - <%=rs("descricao")%>
	</td>
	<td class=campo style="<%=estilo%>" valign="top" >

	<!-- quadro dos dias com marcações incompletas -->
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo width=140>Data</td>
	<td class=titulo width=260>Marcações efetuadas</td>
</tr>
<%
sql2="select a.data, datepart(dw,a.data) as diasem, envio=max(c.dtenvio), tipo=max(c.tipo), vezes=count(c.dtenvio) " & _
"from _marcacoes_checagem a left join n3controle c on c.chapa=a.chapa and c.data=a.data " & _
"where a.chapa='" & rs("chapa") & "' and a.data between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "' " & _
"group by a.chapa, a.data order by a.data" 
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof 
%>
<tr>
	<td class=campo align="center"><%=rs2("data")%> (<%=weekdayname(weekday(rs2("data")),1)%>)</td>
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
	response.write temp
	if rs3.absoluteposition<rs3.recordcount then response.write " - "
rs3.movenext
loop
else
	response.write "-"
end if
rs3.close
%>
	</td>
</tr>	
<%
if rs2("vezes")>0 then
	if rs2("tipo")="E" then tipo="Email" else tipo="Formulário"
	if rs2("vezes")>1 then texto1="vezes" else texto1="vez"
%>
<tr>
	<td class=fundor colspan=2><font color=red><b>
	<%="Ultimo envio em " & rs2("envio") & " por " & tipo & " (" & rs2("vezes") & texto1 & ")"%>
	</b></font>
	</td>
</tr>
<%
end if 'rs2("vezes")>0
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
	<a href="n3_email.asp?chapa=<%=rs("chapa")%>&datai=<%=datai%>&dataf=<%=dataf%>" onclick="NewWindow(this.href,'Selecao_email','690','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/email_go.png" border="0" width="15" alt="Enviar•Email"></a>
	<br>
	<a href="n3_print.asp?chapa=<%=rs("chapa")%>&datai=<%=datai%>&dataf=<%=dataf%>" onclick="NewWindow(this.href,'Selecao_print','690','400','yes','center');return false" onfocus="this.blur()">
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