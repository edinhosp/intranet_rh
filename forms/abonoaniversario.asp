<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a76")="N" or session("a76")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Formulário para Abono Aniversário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function nome1() { form.chapa.value=form.nome.value; }
function chapa1() { form.nome.value=form.chapa.value; }
function secao1() { form.nsecao.value=form.secao.value;form.submit() }
function secao2() { form.secao.value=form.nsecao.value; }
--></script>
</head>
<body style="margin-left:20px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
sessao=session.sessionid

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("B1")="" then
frmchapa=request.form("chapa")
%>
<form name="form" action="abonoaniversario.asp" method="post">
<table border="1" bordercorlor="#CCCCCC" cellpadding="4" cellspacing="0" style="border-collapse: collapse" width=500>
<tr>
	<td class=titulo colspan=3>Abono Aniversário de funcionário</td>
</tr>
<tr>
	<td class=campo>Funcionário</td>
	<td class=campo><input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=request.form("chapa")%>"></td>
	<td class=campo><select name="nome" class=a onchange="nome1()">
		<option value="0">Selecione o funcionário ou o mês</option>
		<option value="m01a">Aniversários de JANEIRO deste ano</option>
		<option value="m01p">Aniversários de JANEIRO do ano seguinte</option>
		<option value="m02">Aniversários de FEVEREIRO</option>
		<option value="m03">Aniversários de MARÇO</option>
		<option value="m04">Aniversários de ABRIL</option>
		<option value="m05">Aniversários de MAIO</option>
		<option value="m06">Aniversários de JUNHO</option>
		<option value="m07">Aniversários de JULHO</option>
		<option value="m08">Aniversários de AGOSTO</option>
		<option value="m09">Aniversários de SETEMBRO</option>
		<option value="m10">Aniversários de OUTUBRO</option>
		<option value="m11">Aniversários de NOVEMBRO</option>
		<option value="m12">Aniversários de DEZEMBRO</option>
		<option value="0"><hr></option>
<%
sql="select p.chapa, p.nome from corporerm.dbo.pfunc p where  p.codsituacao<>'D' and p.codsindicato<>'03' and p.codtipo='N' " & _
"order by p.nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
if frmchapa=rs("chapa") then tempc="selected" else tempc=""
%>
		<option value="<%=rs("chapa")%>" <%=tempc%>> <%=rs("nome")%></option>
<%
rs.movenext:loop
%>
	</select>
	</td>
</tr>
<tr>
	<td class=campo colspan=5>&nbsp;
	<input type="submit" value="Visualizar" class=button name="B1">
	<input type="checkbox" value="ON" name="semdireito"><b>Reimprimir formulário com mensagem de perda de direito
	</td>
</tr>
</table>

</form>
<%
else

frmchapa=request.form("chapa")

if left(frmchapa,1)="m" then
	mes=mid(frmchapa,2,2)
	qualano=mid(frmchapa,4,1)
	sql="select chapa, nome, codsituacao, funcao, codsecao, secao, dtnascimento " & _
	"from qry_funcionarios " & _
	"where codsindicato<>'03' and codtipo='N' and codsituacao in ('A','F','Z') and month(dtnascimento)=" & mes & " " & _
	"order by codsecao, nome "
	lista=1
else
	sql="select chapa, nome, codsituacao, funcao, codsecao, secao, dtnascimento " & _
	"from qry_funcionarios " & _
	"where codsindicato<>'03' and codtipo='N' and codsituacao in ('A','F','Z') and chapa='" & frmchapa & "' "
	lista=0
end if
if qualano="p" then qano=1 else qano=0
rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof
%>
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr>
		<td><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
		<td align="right"><p style="font-size:18pt"><b>Abono Aniversário</b><td>
	</tr>
</table>
<br><br>

<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border-left: 1px solid;border-right: 1px solid;border-top: 1px solid;font-size:10pt">
	<i>Nome do Empregado</i></td></tr>
	<tr><td class="campop" style="border-left: 1px solid;border-right: 1px solid;border-bottom: 1px solid;font-size:12pt">
	<b><%=rs("nome")%></b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border-left: 1px solid;border-right: 1px solid;font-size:10pt">
	<i>Chapa</i></td>
	<td class="campop" style="border-right: 1px solid;font-size:10pt">
	<i>Departamento</i></td></tr>
	<tr><td class="campop" style="border-right: 1px solid;border-left: 1px solid;border-bottom: 1px solid;font-size:12pt">
	<%=rs("chapa")%></td>
	<td class="campop" style="border-right: 1px solid;border-bottom: 1px solid;font-size:12pt">
	<%=rs("codsecao")%> - <%=rs("secao")%></td></tr>
</table>
<br>
<%
dtnasc=rs("dtnascimento")
aniv=dateserial(year(now)+qano,month(dtnasc),day(dtnasc))

sqld="select diaferiado from corporerm.dbo.gferiado where diaferiado='" & dtaccess(aniv) & "' "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then feriado=1 else feriado=0
rs2.close
diasem=weekday(aniv)
diasem2=weekdayname(diasem)

'if diasem=1 or diasem=7 then semana=1 else semana=0
if diasem=1 then semana=1 else semana=0

if feriado=0 and semana=0 then
	diadescanso=formatdatetime(aniv,2) & " (" & diasem2 & ") "
else
	diadescanso="____/____/_____ (_____________) "
	diadescanso=""
end if
%>

<%if request.form("semdireito")="" then%>
<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border:1px solid;font-size:12pt;text-align:justify">

De acordo com a Portaria nº 13/2015 da Reitoria, e tendo em vista que meu aniversário
será em <%=formatdatetime(aniv,2)%> (<%=diasem2%>):
<!--
<br>
<br>[&nbsp;&nbsp;&nbsp;] gratificação no valor correspondente a um dia de salário;
-->
<br>
<%if feriado=1 or semana=1 then%>
<br>[&nbsp;&nbsp;&nbsp;] <input type="text" class="form_input10" style="font-size:12pt;" value="" size=70>
<%else%>
<!-- <br>[&nbsp;&nbsp;&nbsp;] um dia de descanso na data de &nbsp;<input type="text" class="form_input10" style="font-size:12pt;" value="<%=diadescanso%>" size=40>&nbsp;. 
-->
<br>[&nbsp;&nbsp;&nbsp;] <input type="text" class="form_input10" style="font-size:12pt;" value="" size=70>
<%end if%>
<br>
<br>
<br>
<!--
E de acordo com o art. 4º da Portaria nº 13/2015 de 10/02/2015, estou ciente de que perderei o abono caso tenha faltado nos 12 (doze)
meses anteriores ao meu aniversário.-->
	</td></tr>
</table>
<%else%>
<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border:1px solid;font-size:12pt;text-align:justify">

De acordo com a Portaria nº 13/2015 da Reitoria, e mais especificamente em relação
ao art. 4º informamos que nos últimos 12 (doze) meses o critério estabelecido
para o direito não foi atingido.
<br>
<br>

Queremos aproveitar a ocasião e desejar que em <%=day(aniv)& "/" & month(aniv)%>, seu dia seja repleto de surpresas agradáveis.


	</td></tr>
</table>
<%end if%>

<br>

<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border:1px solid;font-size:12pt" width=300>
	<br>Osasco, ______/__________/_______
	<br><br><br><br>_______________________________
	<br>    <%=rs("nome")%>
	</td>
	<td class="campop" style="border:1px solid;font-size:12pt">
	<br>Ciente em: ______/______/______
	<br><br><br><br>_______________________________
	<br>    Encarregado/Supervisor
	</td></tr>
</table>
<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
<tr><td>
<p style="margin-bottom:0;margin-top:0;text-align:right"><%=rs.absoluteposition%>/<%=rs.recordcount%>
</td></tr></table>
<%
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>"
rs.movenext
loop

if lista=1 then
	response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<table border="1" cellpadding="2" cellspacing="0" width="650" bordercolor="#000000">
<tr><td class=grupo colspan=6>Relação de Abonos do mês <%=mes%></td></tr>
<tr>
	<td class=titulo>#</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Setor</td>
	<td class=titulo>Aniverário</td>
	<td class=titulo>Controle</td>
</tr>
<%
	rs.movefirst
	do while not rs.eof
%>
<tr>
	<td class=campo><%=rs.absoluteposition%></td>
	<td class=campo><%=rs("chapa")%></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo><%=rs("secao")%></td>
	<td class=campo align="center"><%=numzero(day(rs("dtnascimento")),2) & "/" & numzero(month(rs("dtnascimento")),2)%></td>
<%
fano=year(now)+qano & mes & "01"
sqlf="SELECT ANOCOMP, MESCOMP, eve=CASE WHEN CODEVENTO='087' THEN 'FALTAS (D)' ELSE 'FALTAS (H)' END, REF " & _
"FROM CORPORERM.DBO.PFFINANC WHERE CHAPA='" & rs("chapa") & "' AND DTPAGTO between DATEADD(""YYYY"",-1,'" & fano & "') and '" & fano & "' AND CODEVENTO IN ('087','088') "
'response.write sqlf
rs2.Open sqlf, ,adOpenStatic, adLockReadOnly
%>	
	<td class=fundor nowrap>
<%
do while not rs2.eof
	response.write rs2("mescomp") & "/" & rs2("anocomp") & ":" & rs2("eve") &"=" & rs2("ref")
	if rs2.recordcount>0 and rs2.absoluteposition<rs2.recordcount then response.write "<br>"
rs2.movenext:loop
rs2.close
%>
	</td>
</tr>
<%
	rs.movenext
	loop
%>
</table>
<%
end if

rs.close

end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>