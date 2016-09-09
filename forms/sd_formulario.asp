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
<body style="margin-left:3px;margin-top:80px;">
<%
'<body style="margin-left:42px;margin-top: 80px;">
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
sessao=session.sessionid
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

espacamento=5

if request.form="" then
sql="select p.chapa, p.nome from corporerm.dbo.pfunc p where p.chapa<'10000' and p.codtipo='N' order by p.nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form name="form" action="sd_formulario.asp" method="post">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Sele&#231;&#227;o de Funcion&#225;rio para emiss&#227;o de SD</td>
</tr>
<tr>
	<td class=campo>Funcion&#225;rio</td>
	<td class=campo><input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()"></td>
	<td class=campo>
		<select name="nome" class=a onchange="nome1()">
		<option value="0"> Selecione o funcion&#225;rio</option>
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
dim ta(41)
sql="select f.chapa, f.nome, f.pispasep, f.dataadmissao, f.datademissao, p.dtnascimento, c.nome as funcao, f.codsecao, d.mae, " & _
"p.carteiratrab, p.seriecarttrab, p.ufcarttrab, p.rua, p.numero, p.complemento, p.bairro, p.cidade, p.estado, p.cep, p.telefone1, " & _
"f.pispasep, p.carteiratrab, p.seriecarttrab, p.ufcarttrab, p.cpf, c.cbo2002, p.sexo, p.grauinstrucao, f.jornadamensal, " & _
"f.codsindicato, f.temavisoprevio " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.pfuncao c, (select chapa, nome as mae from corporerm.dbo.pfdepend where grauparentesco='7') d " & _
"where f.codpessoa=p.codigo and c.codigo=f.codfuncao and d.chapa=f.chapa and f.chapa='" & request.form("chapa") & "' "
'response.write sql
rs.Open sql, ,adOpenStatic, adLockReadOnly
campus=left(rs("codsecao"),2)
if campus="01" then cnpj="73.063.166/0001-20"
if campus="03" then cnpj="73.063.166/0003-92" 
if campus="04" then cnpj="73.063.166/0004-73" 
%>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
ti=17
ta(1) =30:ta(2) =30:ta(3) =30:ta(4) =30:ta(5) =30:ta(6) =30:ta(7) =30:ta(8) =30:
ta(9) =30:ta(10)=30:ta(11)=30:ta(12)=30:ta(13)=30:ta(14)=30:ta(15)=30:ta(16)=30:
ta(17)=30:ta(18)=30:ta(19)=30:ta(20)=30:ta(21)=30:ta(22)=30:ta(23)=30:ta(24)=30:
ta(25)=30:ta(26)=30:ta(27)=30:ta(28)=30:ta(29)=30:ta(30)=30:ta(31)=30:ta(32)=30:
ta(33)=30:ta(34)=30:ta(35)=30:ta(36)=30:ta(37)=30:ta(38)=30:ta(39)=30:ta(40)=30:
for a=1 to 40
letra=mid(rs("nome"),a,1)
if letra=" " then letra=" "
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=<%=ta(a)%>  align="center">
	<%=letra%></td>
<%
if ti=17 then ti=17 else ti=17
next
%>
</tr>
</table>

<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
ti=17
for a=1 to 40
letra=mid(rs("mae"),a,1)
if letra=" " then letra=" "
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=<%=ta(a)%> align="center">
	<%=letra%></td>
<%
if ti=17 then ti=17 else ti=17
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
if letra=" " then letra=" "
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=<%=ta(a)%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=19 else ti=16
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
if rs("complemento")="" or isnull(rs("complemento")) then complemento=" " else complemento=rs("complemento")
telefone=textopuro(rs("telefone1"),2)
if left(telefone,4)="0011" then telefone=right(telefone,len(telefone)-4)
if left(telefone,3)="011" then telefone=right(telefone,len(telefone)-3)
letra=mid(espaco2(complemento,17) & " " & replace(rs("cep"),"-"," ") & " " & rs("estado") & " " & espaco2(telefone,10)   ,a,1)
if letra=" " then letra=" "
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=<%=ta(a)%> align="center">
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
for a=7 to 11:ta(a)=40:next:ta(20)=40:ta(22)=55:ta(29)=55:ta(30)=55:for a=31 to 40:ta(a)=35:next
for a=1 to 40
if isnull(rs("ufcarttrab")) or rs("ufcarttrab")="" then ufctps="  " else ufctps=rs("ufcarttrab")
letra=mid(espaco2(rs("pispasep"),14) & espaco2(rs("carteiratrab"),7) & espaco2(rs("seriecarttrab"),3) & espaco2(ufctps,2) & space(3) & espaco2(rs("cpf"),11)   ,a,1)
if letra=" " then letra=" "
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=<%=ta(a)%> align="center">
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
for a=7 to 11:ta(a)=30:next:ta(20)=30:ta(22)=30:ta(29)=30:ta(30)=30:for a=31 to 40:ta(a)=30:next
ta(9)=35
ti=16
for a=1 to 40
letra=mid(space(5) & "1" & " " & espaco2(textopuro(cnpj,2),14) & "  " & "80314"   ,a,1)
if letra=" " then letra=" "
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=<%=ta(a)%> align="center">
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
ta(9)=30
ti=16
cbo=left(rs("cbo2002"),5) & " " & right(rs("cbo2002"),1)
for a=1 to 7
letra=mid( cbo  ,a,1)
if letra=" " then letra=" "
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=<%=20%> align="center">
	<%=letra%></td>
<%
if ti=16 then ti=18 else ti=16
next
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=400 align="left" nowrap>&nbsp;<%=rs("funcao")%></td>
</tr>
</table>

<br>
<br>
<br>
<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<%
ti=16
ta(7)=25:ta(38)=20
admissao=numzero(day(rs("dataadmissao")),2) & numzero(month(rs("dataadmissao")),2) & right(year(rs("dataadmissao")),2)
demissao=numzero(day(rs("datademissao")),2) & numzero(month(rs("datademissao")),2) & right(year(rs("datademissao")),2)
if rs("sexo")="F" then sexo="2" else sexo="1"
if rs("grauinstrucao")>"9" then instrucao="9" else instrucao=rs("grauinstrucao")
nasc=numzero(day(rs("dtnascimento")),2) & numzero(month(rs("dtnascimento")),2) & right(year(rs("dtnascimento")),2)
jornada=rs("jornadamensal")/60
if rs("codsindicato")="03" then jmes=jornada/4.5 else jmes=jornada/5
jmes=numzero(int(jmes),2)
for a=1 to 40
letra=mid( espaco2(admissao,8) & espaco2(demissao,8) & space(4) & sexo & space(5) & instrucao & space(2) & espaco2(nasc,8) & espaco2(jmes,2)   ,a,1)
if letra=" " then letra=" "
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=<%=ta(a)%> align="center">
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
ta(7)=30:ta(38)=30
ta(4)=60:ta(14)=30:ta(17)=30:ta(23)=30:ta(27)=30:ta(30)=30
sqlsal="select top 4 mescomp, baseinss=sum(baseinss) from corporerm.dbo.pfperff where chapa='" & rs("chapa") & "' and nroperiodo<>4 " & _
"group by anocomp, mescomp order by anocomp desc, mescomp desc"
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
letra=mid( espaco2(mes1,3) & espaco1(sal1,10) & space(0) & espaco2(mes2,3) & espaco1(sal2,11) & space(0) & espaco2(mes3,3) & espaco1(sal3,10)    ,a,1)
if letra=" " then letra=" "
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=<%=ta(a)%> align="center">
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
ta(4)=30:ta(14)=30:ta(17)=30:ta(23)=30:ta(27)=30:ta(30)=30
for a=1 to 40
letra=mid( espaco1(tsal,10) & "  " & "104" & space(23) & meses    ,a,1)
if letra=" " then letra=" "
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=<%=ta(a)%> align="center">
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
letra=mid( space(11) & recebeu & space(8) & aviso    ,a,1)
if letra=" " then letra=" "
%>
<td style="font-size:11pt;letter-spacing:0px;font-family:'Arial';" width=<%=ta(a)%> align="center">
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