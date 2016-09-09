<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a42")="N" or session("a42")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Reembolso Medial</title>
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
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function nome1() { form.chapa.value=form.nome.value;form.submit(); }
function chapa1() { form.nome.value=form.chapa.value;form.submit(); }
--></script>
<%
dim conexao, rs, rs2
dim mes(12)
mes(1)="Janeiro":mes(2)="Fevereiro":mes(3)="Março":mes(4)="Abril":mes(5)="Maio":mes(6)="Junho"
mes(7)="Julho":mes(8)="Agosto":mes(9)="Setembro":mes(10)="Outubro":mes(11)="Novembro":mes(12)="Dezembro"
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request.form("B1")="" or request.form("id")="" then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para Recibo de Férias
<form method="POST" action="recibo_ferias.asp" name="form">
<%
sqla="select f.chapa, f.nome from corporerm.dbo.pfunc f where f.chapa in (select distinct chapa from corporerm.dbo.pfuferiasrecibo) order by f.nome "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=request.form("chapa")%>">
<select name="nome" class=a onchange="nome1()">
	<option value="0">Selecione o funcionário</option>
	<option value="00000" <%if request.form("chapa")="00000" then response.write "selected"%>>Por data de pagamento</option>
<%
rs.movefirst:do while not rs.eof
if request.form("chapa")=rs("chapa") then temps="selected" else temps=""
%>
	<option value="<%=rs("chapa")%>" <%=temps%> ><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<br><br>
<table border="1" cellpadding="2" cellspacing="2" style="border-collapse: collapse">
<tr>
	<td class=titulo></td>
	<td class=titulo>Vencimento</td>
	<td class=titulo>Período</td>
	<td class=titulo>Data Pagto</td>
</tr>
<%
if request.form("chapa")<>"00000" then
sqlp="select chapa, dtvencimento, nroperiodo, dtpagto, dtaviso from corporerm.dbo.pfperfer_old where chapa='" & request.form("chapa") & "' order by dtvencimento, nroperiodo "
sqlp="select CHAPA, 'dtvencimento'=FIMPERAQUIS, 'dtinicio'=DATAINICIO, 'dtpagto'=DATAPAGTO, 'dtaviso'=DATAAVISO from corporerm.dbo.PFUFERIASPER where CHAPA='" & request.form("chapa") & "' order by FIMPERAQUIS, DATAPAGTO "
rs.Open sqlp, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	do while not rs.eof
%>
<tr>
	<td class=campo><input type="radio" name="id" value="<%=rs("dtinicio") & rs("dtvencimento")%>"></td>
	<td class=campo><%=rs("dtvencimento")%></td>
	<td class=campo><%=rs("dtinicio")%></td>
	<td class=campo><%=rs("dtpagto")%></td>
</tr>
<%
	rs.movenext:loop
end if
rs.close
else
sqlp="SELECT 'DTPAGTO'=datapagto, Count(CHAPA) AS Total FROM corporerm.dbo.pfuferiasper GROUP BY DaTaPAGTO HAVING DaTaPAGTO>getdate()-30 ORDER BY DaTaPAGTO DESC"
sqlp="SELECT 'DTPAGTO'=p.datapagto, Count(p.CHAPA) AS Total FROM corporerm.dbo.pfuferiasper p inner join corporerm.dbo.PFUFERIASRECIBO r on r.CHAPA=p.CHAPA and r.FIMPERAQUIS=p.FIMPERAQUIS and r.DATAPAGTO=p.DATAPAGTO GROUP BY p.DaTaPAGTO HAVING p.DaTaPAGTO>getdate()-30 ORDER BY p.DaTaPAGTO DESC "
rs.Open sqlp, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	do while not rs.eof
%>
<tr>
	<td class=campo><input type="radio" name="id" value="00/00/0000<%=rs("dtpagto")%>"></td>
	<td class=campo>(<%=rs("total")%> recibos)</td>
	<td class=campo>&nbsp;</td>
	<td class=campo><%=rs("dtpagto")%></td>
</tr>
<%
	rs.movenext:loop
end if
rs.close

end if
%>
</table>
<br>
<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>
<%
end if

if request.form("B1")<>"" and request.form("id")<>"" then
temp=request.form("id")
dtinicio=left(temp,10)
dtvencimento=right(temp,len(temp)-10)
chapa=request.form("chapa")
contador=0

if chapa="00000" then
	sqlsel="select distinct chapa, fimperaquis, datainicio from corporerm.dbo.pfuferiasrecibo where datapagto='" & dtaccess(dtvencimento) & "' "
	rs.Open sqlsel, ,adOpenStatic, adLockReadOnly
	do while not rs.eof
		redim preserve ch(contador):ch(contador)=rs("chapa")
		redim preserve dv(contador):dv(contador)=rs("fimperaquis")
		redim preserve np(contador):np(contador)=rs("datainicio")
	rs.movenext
	contador=contador+1
	loop
	rs.close
else
	redim ch(0),dv(0),np(0)
	ch(0)=chapa
	dv(0)=dtvencimento
	np(0)=dtinicio
end if

for b=0 to ubound(ch)
sql1="select f.chapa, f.nome, f.codsecao, f.funcao, f.carteiratrab, f.seriecarttrab, s.cgc, s.rua, s.numero, s.bairro, s.cidade, f.salario, f.codrecebimento " & _
"from qry_funcionarios f, corporerm.dbo.psecao s where s.codigo=f.codsecao and f.chapa='" & ch(b) & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")

sql2="select dtiniperaquis, dtfimperaquis, dtinigozo, dtfimgozo, diasabono, nrofaltas " & _
"from corporerm.dbo.pfhstfer_old where chapa='" & ch(b) & "' and nroperiodo=" & np(b) & " and dtfimperaquis='" & dtaccess(dv(b)) & "' "
sql2="select dtiniperaquis=f.INICIOPERAQUIS, dtfimperaquis=p.FIMPERAQUIS, dtinigozo=p.DATAINICIO, dtfimgozo=p.DATAFIM, diasabono=NRODIASABONO, nrofaltas=p.FALTAS " & _
"from corporerm.dbo.PFUFERIASPER p inner join corporerm.dbo.PFUFERIAS f on f.CHAPA=p.CHAPA and f.FIMPERAQUIS=p.FIMPERAQUIS " & _
"where p.CHAPA='" & ch(b) & "' and DATAINICIO='" & dtaccess(np(b)) & "' and p.FIMPERAQUIS='" & dtaccess(dv(b)) & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
d1=day(rs2("dtiniperaquis"))
m1=mes(month(rs2("dtiniperaquis")))
a1=year(rs2("dtiniperaquis"))
d2=day(rs2("dtfimperaquis"))
m2=mes(month(rs2("dtfimperaquis")))
a2=year(rs2("dtfimperaquis"))
if isnull(rs2("dtinigozo")) or rs2("dtinigozo")="" then
	sql3="select inicprogferias1, fimprogferias1 from corporerm.dbo.pfunc where chapa='" & ch(b) & "'"
	rs3.Open sql3, ,adOpenStatic, adLockReadOnly
	inigozo=rs3("inicprogferias1"):fimgozo=rs3("fimprogferias1")
	rs3.close
else
	inigozo=rs2("dtinigozo"):fimgozo=rs2("dtfimgozo")
end if
d3=day(inigozo)
m3=mes(month(inigozo))
a3=year(inigozo)
d4=day(fimgozo)
m4=mes(month(fimgozo))
a4=year(fimgozo)
dabono=rs2("diasabono")
faltas=rs2("nrofaltas")
rs2.close

sql2="select dtpagto, dtaviso from corporerm.dbo.pfperfer_old where chapa='" & ch(b) & "' and nroperiodo=" & np(b) & " and dtvencimento='" & dtaccess(dv(b)) & "' "
sql2="select dtpagto=datapagto, dtaviso=dataaviso  " & _
"from corporerm.dbo.PFUFERIASPER p inner join corporerm.dbo.PFUFERIAS f on f.CHAPA=p.CHAPA and f.FIMPERAQUIS=p.FIMPERAQUIS " & _
"where p.CHAPA='" & ch(b) & "' and DATAINICIO='" & dtaccess(np(b)) & "' and p.FIMPERAQUIS='" & dtaccess(dv(b)) & "' "

rs2.Open sql2, ,adOpenStatic, adLockReadOnly
d5=day(rs2("dtaviso")):calculo=rs2("dtaviso")
m5=mes(month(rs2("dtaviso")))
a5=year(rs2("dtaviso"))
d6=day(rs2("dtpagto"))
m6=mes(month(rs2("dtpagto")))
a6=year(rs2("dtpagto"))
rs2.close

sql21="select top 1 dtmudanca from corporerm.dbo.pfhstsal where chapa='" & ch(b) & "' and dtmudanca<='" & dtaccess(calculo) & "' order by dtmudanca desc"
sql2="select salario from corporerm.dbo.pfhstsal where chapa='" & ch(b) & "' and dtmudanca in (" & sql21 & ") "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
salario=rs2("salario")
rs2.close
%>
<div align="center">
<center>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campop" align="center" valign="center"><font size="+1"><b>AVISO E RECIBO DE FÉRIAS</font></td></tr>
<tr><td class=campo align="center">(Para atender ao Decreto-Lei nº 5.452 de 01/05/1943, com as alterações do Decreto-Lei nº 1.535 de 13/04/1977)</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=10><tr><td></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=fundop align="center"><b>AVISO PRÉVIO DE FÉRIAS</td></tr>
<tr><td class=fundo align="center">(de acordo com o art. 135 da C.L.T., participando no mínimo com 30 dias de antecedência)</td></tr>
</table>
	
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=10><tr><td></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campor" style="border:1px solid #000000;">&nbsp;Nome do empregado</td>
	<td class="campor" style="border:1px solid #000000;">&nbsp;Nº da CTPS</td>
	<td class="campor" style="border:1px solid #000000;">&nbsp;Série</td>
	<td class="campor" style="border:1px solid #000000;">&nbsp;Registro</td>
</tr>
<tr>
	<td class="campop" height=25 align="center" style="border-left:1px solid #000000;"><%=rs("nome")%></td>
	<td class="campop" align="center" style="border-left:1px solid #000000;"><%=rs("carteiratrab")%></td>
	<td class="campop" align="center" style="border-left:1px solid #000000;"><%=rs("seriecarttrab")%></td>
	<td class="campop" align="center" style="border-left:1px solid #000000;border-right:1px solid #000000;"><%=rs("chapa")%></td>
</tr>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campor" align="center" style="border:1px solid #000000;">&nbsp;Período de Aquisição</td><tr>
<tr>
	<td class="campop" height=25 align="center" style="border-left:1px solid #000000;border-right:1px solid #000000;">
	De <%=d1%> de <%=m1%> de <%=a1%> A <%=d2%> de <%=m2%> de <%=a2%></td></tr>
<tr><td class="campor" align="center" style="border:1px solid #000000;">&nbsp;Período de Gozo das Férias</td><tr>
<tr>
	<td class="campop" height=25 align="center" style="border-left:1px solid #000000;border-right:1px solid #000000;">
	De <%=d3%> de <%=m3%> de <%=a3%> A <%=d4%> de <%=m4%> de <%=a4%></td></tr>
<tr><td class="campop" style="border-bottom:1px solid #000000;"></td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=10><tr><td></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=campo colspan=3 align="center" style="border:1px solid #000000;">&nbsp;BASE PARA CÁLCULO DA REMUNERAÇÃO DE FÉRIAS</td><tr>
<tr>
	<td class="campor" style="border:1px solid #000000;" width=33%>&nbsp;Faltas não justif.</td>
	<td class="campor" style="border:1px solid #000000;" width=33%>&nbsp;Salário base</td>
	<td class="campor" style="border:1px solid #000000;" width=33%>&nbsp;Base de cálculo</td>
</tr>
<tr>
	<td class="campop" height=25 align="center" style="border-left:1px solid #000000;"><%=faltas%></td>
	<td class="campop" align="center" style="border-left:1px solid #000000;"><%=formatnumber(salario,2)%></td>
	<td class="campop" align="center" style="border-left:1px solid #000000;border-right:1px solid #000000;"><%=rs("codrecebimento")%></td>
</tr>
<tr><td class="campop" colspan=3 style="border-bottom:1px solid #000000;"></td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=10><tr><td></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo align="center" style="border:1px solid #000000;">Cod.</td>
	<td class=campo align="center" style="border:1px solid #000000;">Ref. #</td>
	<td class=campo align="center" style="border:1px solid #000000;">Descrição</td>
	<td class=campo align="center" style="border:1px solid #000000;">Proventos</td>
	<td class=campo align="center" style="border:1px solid #000000;">Descontos</td>
</tr>
<%
sql2="select r.codevento, r.ref, r.valor, e.descricao, e.provdescbase " & _
"from corporerm.dbo.pfferias_old r, corporerm.dbo.pevento e where e.codigo=r.codevento and r.chapa='" & ch(b) & "' and nroperiodo=" & np(b) & " and dtvencimento='" & dtaccess(dv(b)) & "' " & _
" and provdescbase<>'B' and valor>0 order by provdescbase desc, codevento"
sql2="select r.CODEVENTO, r.REF, r.VALOR, e.DESCRICAO, e.PROVDESCBASE " & _
"from corporerm.dbo.PFUFERIASVERBAS r inner join corporerm.dbo.PEVENTO e on r.CODEVENTO=e.CODIGO " & _
"inner join corporerm.dbo.PFUFERIASPER p on p.CHAPA=r.CHAPA and p.FIMPERAQUIS=r.FIMPERAQUIS and p.DATAPAGTO=r.DATAPAGTO " & _
"where r.CHAPA='" & ch(b) & "' and r.FIMPERAQUIS='" & dtaccess(dv(b)) & "' and p.DATAINICIO='" & dtaccess(np(b)) & "' and PROVDESCBASE<>'B' and r.VALOR>0 " & _
"order by PROVDESCBASE desc, codevento "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
totprov=0:totdesc=0
do while not rs2.eof
valor=cdbl(rs2("valor"))
if rs2("provdescbase")="P" then impprov=formatnumber(valor,2) else impprov="&nbsp;"
if rs2("provdescbase")="D" then impdesc=formatnumber(valor,2) else impdesc="&nbsp;"
if rs2("provdescbase")="P" then totprov=totprov+valor
if rs2("provdescbase")="D" then totdesc=totdesc+valor
if cdbl(rs2("ref"))=0 then ref="&nbsp;" else ref=rs2("ref")
%>
<tr>
	<td class="campop" style="border-left:1px solid #000000;" align="center" height=20><%=rs2("codevento")%></td>
	<td class="campop" style="border-left:1px solid #000000;" align="right"><%=ref%>&nbsp;&nbsp;</td>
	<td class="campop" style="border-left:1px solid #000000;">&nbsp;<%=rs2("descricao")%></td>
	<td class="campop" style="border-left:1px solid #000000;" align="right"><%=impprov%>&nbsp;&nbsp;</td>
	<td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000" align="right"><%=impdesc%>&nbsp;&nbsp;</td>
</tr>
<%
rs2.movenext:loop
rs2.close
liquido=totprov-totdesc
%>
<tr>
	<td class="campop" colspan=3 style="border-left:1px solid #000000;border-top:1px solid #000000">&nbsp;Totais</td>
	<td class="campop" style="border-left:1px solid #000000;border-top:1px solid #000000" align="right"><%=formatnumber(totprov,2)%>&nbsp;&nbsp;</td>
	<td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000;border-top:1px solid #000000" align="right"><%=formatnumber(totdesc,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campop" colspan=3 align="right" style="border-left:1px solid #000000;border-top:1px solid #000000">&nbsp;Líquido&nbsp;&nbsp;</td>
	<td class="campop" colspan=2 align="center" style="border-left:1px solid #000000;border-right:1px solid #000000;border-top:1px solid #000000" align="right"><%=formatnumber(liquido,2)%>&nbsp;&nbsp;</td>
</tr>
<tr><td class="campop" colspan=5 style="border-bottom:3px double #000000;"></td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=10><tr><td></td></tr></table>


<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campop"><p style="margin-top:5;text-align:justify; line-height:20px"><%for a=1 to 10:response.write "&nbsp;":next%>
	Pela presente comunicamos-lhe que, de acordo com a legislação vigente, ser-lhe-ão concedidas férias 
	relativas ao período acima descrito e a sua disposição fica a importância líquida de R$ <%=formatnumber(liquido,2)%> (<u><%=extenso2(liquido)%> 
	<%for a=1 to (150-len(extenso2(liquido))):response.write " x":next%></u>), a ser paga adiantadamente.
	</td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campop" height=30 align="left" style="">Osasco, <%=d5%> de <%=m5%> de <%=a5%></td><td class="campop" width=10%></td><td class="campop" width=45%></td>
</tr>
<tr><td class="campor" align="center" style="border-top:1px solid #000000;">&nbsp;Local e data</td><td class="campop" width=10%></td><td class="campop" width=45%></td>
</tr>
<tr><td class="campor" height=35 align="left" valign=top style="">&nbsp;CIENTE</td><td class="campop" width=10%></td>
	<td class="campor" height=35 align="left" valign=top style="">&nbsp;</td>
</tr>
<tr><td class="campor" align="center" style="border-top:1px solid #000000;">&nbsp;Assinatura do empregado</td><td class="campop" width=10%></td>
	<td class="campor" align="center" style="border-top:1px solid #000000;">&nbsp;Assinatura do empregador</td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=10>
<tr><td height=15></td></tr>
<tr><td style="border-bottom:2px dotted #000000"></td></tr>
<tr><td height=15></td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=fundop align="center"><b>RECIBO DE FÉRIAS</td></tr>
<tr><td class=fundo align="center">(de acordo com o parágrafo único do art. 145 da C.L.T)</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campop"><p style="margin-top:5;text-align:justify; line-height:20px">
	 Recebi da firma <u>&nbsp;FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO<%for a=1 to 10:response.write "&nbsp;":next%></u>
	 estabelecida à <u>&nbsp;<%=rs("rua") & ", " & rs("numero") & " - " & rs("bairro")%><%for a=1 to 10:response.write "&nbsp;":next%></u>
	 em <u>&nbsp;<%=rs("cidade")%><%for a=1 to 10:response.write "&nbsp;":next%></u> a importância de R$ <u>&nbsp;<%=formatnumber(liquido,2)%>&nbsp;</u>
	 <u>&nbsp;(<%=extenso2(liquido)%>)&nbsp;</u> que me é paga antecipadamente por motivo das minhas férias regulamentares, ora concedidas e que vou gozar de
	 acordo com a descrição acima, tudo conforme o aviso que recebi em tempo, ao qual dei meu "ciente".<br>
	 Para clareza e documento, firmo o presente recibo, dando a firma plena e geral quitação.
	</td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campop" height=40 align="left" style="">Osasco, <%=d6%> de <%=m6%> de <%=a6%></td><td class="campop" width=10%></td><td class="campop" width=45%></td>
</tr>
<tr><td class="campor" align="center" style="border-top:1px solid #000000;">&nbsp;Local e data</td><td class="campop" width=10%></td>
	<td class="campor" align="center" style="border-top:1px solid #000000;">&nbsp;Assinatura do empregado</td>
</tr>
</table>

<%
if b<ubound(ch) then response.write "<DIV style=""page-break-after:always""></DIV>"
rs.close

next

set rs=nothing
%>
<%
set rs=nothing
set rs2=nothing
end if ' 

conexao.close
set conexao=nothing
%>
</body>
</html>