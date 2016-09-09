<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a43")="N" or session("a43")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1251">
<title>SD - Seguro Desemprego</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"><!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() {	form.chapa.value=form.nome.value; }
function chapa1() {	form.nome.value=form.chapa.value; }
--></script>
</head>
<body style="margin-left:0px;margin-top:0px;">
<%
'<body style="margin-left:42px;margin-top: 80px;">
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("consql")
sessao=session.sessionid
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

espacamento=5

if request.form="" then
sql="select p.chapa, p.nome from pfunc p where p.chapa<'10000' and p.codtipo='N' order by p.nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form name="form" action="sd_formulario.asp" method="post">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Sele&ccedil;&atilde;o de Funcion&aacute;rio para emiss&atilde;o de SD</td>
</tr>
<tr>
	<td class=campo>Funcion&aacute;rio</td>
	<td class=campo><input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()"></td>
	<td class=campo>
		<select name="nome" class=a onchange="nome1()">
		<option value="0"> Selecione o funcion&aacute;rio</option>
<%
rs.movefirst
do while not rs.eof
%>
		<option value="<%=rs("chapa")%>"> <%=rs("nome")%></option>
<%
rs.movenext
loop
rs.close
%>
		</select>
	</td>
</tr>
<tr>
	<td class=campo colspan=3>&nbsp;
		<input type="submit" value="Visualizar" class=button name="B1">
	</td>
</tr>
</table>
</form>

<%
else 'request.form

sql="select f.chapa, f.nome, f.pispasep, f.dataadmissao, f.datademissao, p.dtnascimento, c.nome funcao, f.codsecao, d.mae, " & _
"p.carteiratrab, p.seriecarttrab, p.ufcarttrab, p.rua, p.numero, p.complemento, p.bairro, p.cidade, p.estado, p.cep, p.telefone1, " & _
"f.pispasep, p.carteiratrab, p.seriecarttrab, p.ufcarttrab, p.cpf, c.cbo2002, p.sexo, p.grauinstrucao, f.jornadamensal, " & _
"f.codsindicato, f.temavisoprevio " & _
"from pfunc f, ppessoa p, pfuncao c, (select chapa, nome as mae from pfdepend where grauparentesco='7') d " & _
"where f.codpessoa=p.codigo and c.codigo=f.codfuncao and d.chapa=f.chapa and f.chapa='" & request.form("chapa") & "' "
'response.write sql
rs.Open sql, ,adOpenStatic, adLockReadOnly
campus=left(rs("codsecao"),2)
if campus="01" then cnpj="73.066.166/0001-20"
if campus="03" then cnpj="73.066.166/0003-92" 
if campus="04" then cnpj="73.066.166/0004-73" 
%>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
ti=16
for a=1 to 40
letra=mid(rs("nome"),a,1)
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';" width=<%=ti%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
</tr>
</table>

<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
ti=16
for a=1 to 40
letra=mid(rs("mae"),a,1)
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';" width=<%=ti%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
</tr>
</table>

<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
ti=16
for a=1 to 40
letra=mid(rs("rua") & " " & rs("numero"),a,1)
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';" width=<%=ti%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
</tr>
</table>

<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
ti=16
for a=1 to 40
if rs("complemento")="" or isnull(rs("complemento")) then complemento=space(16) else complemento=rs("complemento")
telefone=textopuro(rs("telefone1"),2)
if left(telefone,4)="0011" then telefone=right(telefone,len(telefone)-4)
if left(telefone,3)="011" then telefone=right(telefone,len(telefone)-3)
letra=mid(espaco2(complemento,17) & replace(rs("cep"),"-"," ") & " " & rs("estado") & " " & espaco2(telefone,10)   ,a,1)
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';" width=<%=ti%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
</tr>
</table>

<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
ti=16
for a=1 to 40
if isnull(rs("ufcarttrab")) or rs("ufcarttrab")="" then ufctps="  " else ufctps=rs("ufcarttrab")
letra=mid(espaco2(rs("pispasep"),14) & espaco2(rs("carteiratrab"),7) & espaco2(rs("seriecarttrab"),3) & espaco2(ufctps,2) & "   " & espaco2(rs("cpf"),11)   ,a,1)
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';" width=<%=ti%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
</tr>
</table>

<br>
<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
ti=16
for a=1 to 40
letra=mid(space(5) & "1" & " " & espaco2(textopuro(cnpj,2),14) & "  " & "80314"   ,a,1)
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';" width=<%=ti%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
</tr>
</table>

<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
ti=16
for a=1 to 7
cbo=left(rs("cbo2002"),5) & " " & right(rs("cbo2002"),1)
letra=mid( espaco2(cbo,7) & rs("funcao")   ,a,1)
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';" width=<%=ti%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';">&nbsp;<%=rs("funcao")%></td>
</tr>
</table>

<br>
<br>
<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
ti=16
admissao=numzero(day(rs("dataadmissao")),2) & numzero(month(rs("dataadmissao")),2) & right(year(rs("dataadmissao")),2)
demissao=numzero(day(rs("datademissao")),2) & numzero(month(rs("datademissao")),2) & right(year(rs("datademissao")),2)
if rs("sexo")="F" then sexo="2" else sexo="1"
if rs("grauinstrucao")>"9" then instrucao="9" else instrucao=rs("grauinstrucao")
nasc=numzero(day(rs("dtnascimento")),2) & numzero(month(rs("dtnascimento")),2) & right(year(rs("dtnascimento")),2)
jornada=rs("jornadamensal")/60
if rs("codsindicato")="03" then jmes=jornada/4.5 else jmes=jornada/5
jmes=numzero(int(jmes),2)
for a=1 to 40
letra=mid( espaco2(admissao,8) & espaco2(demissao,8) & space(4) & espaco2(sexo,5) & espaco2(instrucao,3) & espaco2(nasc,8) & espaco2(jmes,2)   ,a,1)
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';" width=<%=ti%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
</tr>
</table>

<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
sqlsal="select top 4 mescomp, baseinss from pfperff where chapa='" & rs("chapa") & "' and nroperiodo<>4 order by anocomp desc, mescomp desc"
rs2.Open sqlsal, ,adOpenStatic, adLockReadOnly
tsal=0:ns=3
rs2.movefirst
'if rs2.recordcount>3 then rs2.movenext
rs2.move 1
mes3=numzero(rs2("mescomp"),2)
sal3=cdbl(rs2("baseinss"))*100:tsal=tsal+cdbl(rs2("baseinss"))
if rs2.eof=false then rs2.move 1
if rs2.eof=true then 
	sal2="":tsal=tsal:mes2=""
else 
	mes2=numzero(rs2("mescomp"),2)
	sal2=cdbl(rs2("baseinss"))*100:tsal=tsal+cdbl(rs2("baseinss"))
end if
if rs2.eof=false then rs2.move 1
if rs2.eof=true then 
	sal1="":tsal=tsal:mes1=""
else
	mes1=numzero(rs2("mescomp"),2)
	sal1=cdbl(rs2("baseinss"))*100:tsal=tsal+cdbl(rs2("baseinss"))
end if
ti=16
for a=1 to 41
letra=mid( espaco2(mes1,3) & espaco1(sal1,10) & space(1) & espaco2(mes2,3) & espaco1(sal2,10) & space(1) & espaco2(mes3,3) & espaco1(sal3,10)    ,a,1)
if a=14 then ti=8
if a=17 then ti=8
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';" width=<%=ti%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
</tr>
</table>

<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
tsal=tsal*100
meses=datediff("m", rs("dataadmissao") , rs("datademissao") )
if meses>36 then meses=36
ti=16
for a=1 to 40
if a<10 then ti=19
letra=mid( espaco1(tsal,10) & "   " & "104" & space(21) & meses    ,a,1)
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';" width=<%=ti%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
</tr>
</table>

<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
if meses>6 then recebeu="1" else recebeu="2"
if rs("temavisoprevio")="1" then aviso="1" else aviso="2"
ti=16
for a=1 to 40
if a<10 then ti=19
letra=mid( space(9) & recebeu & space(10) & aviso    ,a,1)
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Courier New';" width=<%=ti%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
</tr>
</table>
<%
rs.close
end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>