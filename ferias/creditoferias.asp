<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")="N" or session("a48")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Carta de Crédito de Férias</title>
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
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
--></script>

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

if request.form("b1")="" then
%>
<p class=titulo>Emissão de Autorização de Crédito - Férias&nbsp;<%=titulo %>
<form method="POST" action="creditoferias.asp" name="form">
<table border="0" width="250" cellspacing="0"cellpadding="3">
<tr>
	<td class=titulo>Data de Pagamento de Férias</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="dtpagto" onChange="javascript:submit()">
	<option>Selecione uma data</option>
&nbsp;
<%
if isdate(request.form("dtpagto"))=true then dtpagto=cdate(request.form("dtpagto"))
sql2="SELECT DTPAGTO, Count(Chapa) as Recibos FROM corporerm.dbo.pfperfer_old GROUP BY DTPAGTO HAVING DTPAGTO>=getdate()-30 ORDER BY DTPAGTO;"
rs.Open sql2, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
if dtpagto=rs("dtpagto") then temp1="selected" else temp1=""
%>
	<option value="<%=rs("dtpagto")%>" <%=temp1%>><%=rs("dtpagto")%>&nbsp;&nbsp;&nbsp; (<%=rs("recibos")%> recibos)</option>
<%
rs.movenext:loop
rs.close
%>
	</select>
	&nbsp;Pg.<input type="text" name="diaspag" value="2" size=2>dias
</td>
</tr>
<tr>
	<td class=fundo>
		<input type="radio" name="sind" value="01" onClick="javascript:submit()" <%if request.form("sind")="01" then response.write "checked" %> > Administrativos 
	 	<input type="radio" name="sind" value="03" onClick="javascript:submit()" <%if request.form("sind")="03" then response.write "checked" %> > Professores
	</td>
</tr>
</table>

<table border="0" width="250" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo><input type="submit" value="Visualizar" class="button" name="B1"></td>
</tr>
</table>

<%
vezes=0
if request.form<>"" then
if request.form("sind")="03" then selecao=" and f.codsindicato='03' " else selecao=" and f.codsindicato<>'03' "
if request.form("sind")="" then selecao=" and f.codsindicato<>'03' "
sql1="select r.chapa, f.nome, f.codbancopagto, f.codagenciapagto, f.contapagamento, r.dtvencimento, r.dtpagto, opbancaria razao " & _
"from corporerm.dbo.pfperfer_old r, corporerm.dbo.pfunc f " & _
"where r.chapa=f.chapa and r.dtpagto='" & dtaccess(request.form("dtpagto")) & "' " & selecao
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<br>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo></td>
</tr>
<%
rs.movefirst
do while not rs.eof
banco=rs("codbancopagto")
if banco<>"237" then classe="campol" else classe="campo"
if rs("razao")<>"07.05" then classe="campov" else classe="campo"
%>
<tr>
	<td class=<%=classe%>><%=rs("chapa")%></td>
	<td class=<%=classe%>><%=rs("nome")%></td>
	<td class=<%=classe%>>
		<input type="checkbox" name="emitir<%=vezes%>" value="ON" <%="checked"%> >
		<input type="hidden" name="id<%=vezes%>" value="<%=rs("chapa")%>">
		<input type="hidden" name="dt<%=vezes%>" value="<%=rs("dtpagto")%>">
	</td>
</tr>
<%
vezes=vezes+1
rs.movenext
loop
session("credferimp")=vezes-1
end if
rs.close
%>
</table>
<%
end if
%>
</form>
<%
else ' request.form<>""
	vez=session("credferimp")
	sql="delete from creditoferias where sessao='" & session.sessionid & "' "
	conexao.execute sql
	for a=0 to vez
		id=request.form("id" & a)
		dtpg=request.form("dt" & a)
		emitir=request.form("emitir" & a)
		dtpg=request.form("dt" & a)
		'response.write id & " " & tabela & " " & emitir & "<br>"
		if emitir="ON" then
			sql="INSERT INTO creditoferias ( sessao, data, chapa ) SELECT '" & session.sessionid & "', '" & dtaccess(dtpg) & "', '" & id & "'"
			conexao.execute sql
		end if
	next

valor=4599.99
dtpagto=cdate(request.form("dtpagto"))
dtpagto1=dtpagto
if dtpagto=dateserial(2005,12,30) then dtpagto1=dateserial(2005,12,29) 'else dtpagto1=dtpagto
if dtpagto=dateserial(2006,12,29) then dtpagto1=dateserial(2006,12,28) 'else dtpagto1=dtpagto
if dtpagto=dateserial(2007,12,31) then dtpagto1=dateserial(2007,12,28) 'else dtpagto1=dtpagto
if dtpagto=dateserial(2008,12,31) then dtpagto1=dateserial(2008,12,30) 'else dtpagto1=dtpagto
if dtpagto=dateserial(2009,12,31) then dtpagto1=dateserial(2009,12,30) 'else dtpagto1=dtpagto
if dtpagto=dateserial(2010,12,31) then dtpagto1=dateserial(2010,12,30) 'else dtpagto1=dtpagto
'parei
'	rs.Open sqlb, ,adOpenStatic, adLockReadOnly

sql1="select r.chapa, f.nome, f.codagenciapagto, f.contapagamento, r.dtvencimento, r.dtpagto " & _
"from corporerm.dbo.pfperfer_old r, corporerm.dbo.pfunc f " & _
"where r.chapa=f.chapa and r.dtpagto='" & dtaccess(dtpagto) & "' and r.chapa collate database_default in (select chapa from creditoferias where sessao='" & session.sessionid & "') order by r.chapa "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
vezes=int(total/13)
if total=13 then vezes=vezes else vezes=vezes+1
rs.close
ultima="00000"
'***************** <=13 cabe numa folha
for giro=1 to vezes

sql1="select top 13 r.chapa, f.nome, f.codagenciapagto, f.contapagamento, r.dtvencimento, r.dtpagto " & _
"from corporerm.dbo.pfperfer_old r, corporerm.dbo.pfunc f " & _
"where r.chapa=f.chapa and r.dtpagto='" & dtaccess(dtpagto) & "' and r.chapa>'" & ultima & "' and r.chapa collate database_default in (select chapa from creditoferias where sessao='" & session.sessionid & "') "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
rs.movefirst:primeira=rs("chapa")
rs.movelast:ultima=rs("chapa")
total=rs.recordcount
rs.movefirst

sql2="SELECT r.DTPAGTO, " & _
"sum(case when provdescbase='D' then -1 else 1 end * valor) as liquido " & _
"FROM corporerm.dbo.pfperfer_old r, corporerm.dbo.pfferias_old AS l, corporerm.dbo.PEVENTO e " & _
"WHERE r.dtvencimento=l.dtvencimento and r.chapa=l.chapa and l.codevento=e.codigo " & _
"and r.nroperiodo=l.nroperiodo " & _
"and r.chapa between '" & primeira & "' and '" & ultima & "' " & _
"and r.chapa collate database_default in (select chapa from creditoferias where sessao='" & session.sessionid & "') " & _
"and e.PROVDESCBASE In ('D','P') GROUP BY r.DTPAGTO HAVING r.DTPAGTO='" & dtaccess(dtpagto) & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
valor=cdbl(rs2("liquido"))
rs2.close

%>
<!-- <div align="right"> -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="620" height="990">
<tr><td height=111><img border="0" src="../images/aguia.jpg" width="236"></td> </tr>
<!-- corpo da carta -->
<tr><td height=800 class="campop" align="left" valign=top>
<%
data1=dtpagto-2
sqld="select diaferiado from corporerm.dbo.gferiado " & _
"where diaferiado='" & dtaccess(data1) & "' "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then data1=data1-1
rs2.close
if weekday(data1)=7 then data1=data1-1
if weekday(data1)=1 then data1=data1-2
dia=day(data1)
mes=monthname(month(data1))
ano=year(data1)
%>
<p align="right" style="margin-top:0;margin-bottom:0;font-size:12pt">
<input type="text" size=40 class="form_input10" style="text-align:right;font-size:12pt" value="<%="Osasco, " & dia & " de " & mes & " de " & ano%>">
<p align="left" style="margin-top:0;margin-bottom:0;font-size:12pt">
<br>
<br>
<br>
Ao<br>
Banco Bradesco S/A<br>
Agência 3390 - Ag.Empresa Alphaville<br>
Barueri - SP<br>
<br>
<br>
Prezados Senhores:<br>
<br>
<br>
<br>
<%
if total=1 then frase="" else frase="s"
if total=1 then frase2="" else frase2="es"
%>
<p align="justify" style="margin-top:0;margin-bottom:0;font-size:11pt;text-indent:200pt;line-height:150%">
Pela presente autorizamos levar a débito de nossa conta corrente nº 564.600-6, o valor de <b>R$ <%=formatnumber(valor,2)%> (<%=extenso2(valor)%>)</b>, para
pagamento das <b>férias do mês de <%=monthname(month(dtpagto+request.form("diaspag")))%>/<%=year(dtpagto+request.form("diaspag"))%></b>, creditando na<%=frase%> conta<%=frase%> corrente<%=frase%> do<%=frase%> funcionário<%=frase%> o<%=frase%> respectivo<%=frase%> valor<%=frase2%>, 
para crédito em <b><%=dtpagto1%></b> conforme abaixo:
<div align="center">
	<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse" width=90%>
	<tr><td class=titulop align="center">Nome</td>
		<td class=titulop align="center">Agência</td>
		<td class=titulop align="center">Conta Corrente</td>
		<td class=titulop align="center">Valor</td>
	</tr>
<%
do while not rs.eof
sql3="SELECT r.CHAPA, r.DTVENCIMENTO, r.DTPAGTO, sum(case when provdescbase='D' then -1 else 1 end * valor) AS Liquido " & _
"FROM corporerm.dbo.pfperfer_old r, corporerm.dbo.pfferias_old l, corporerm.dbo.PEVENTO e " & _
"WHERE r.dtvencimento=l.dtvencimento and r.chapa=l.chapa and l.codevento=e.codigo " & _
"and r.nroperiodo=l.nroperiodo " & _
"AND e.PROVDESCBASE in ('D','P') GROUP BY r.CHAPA, r.DTVENCIMENTO, r.DTPAGTO " & _
"HAVING r.DTPAGTO='" & dtaccess(dtpagto) & "' and r.chapa='" & rs("chapa") & "' " & _
"and r.chapa in (select chapa collate database_default from creditoferias where sessao='" & session.sessionid & "') "
rs2.Open sql3, ,adOpenStatic, adLockReadOnly
recibo=rs2("liquido")
rs2.close
%>
	<tr><td class="campop"><%=rs("nome")%></td>
		<td class="campop" align="center"><%=rs("codagenciapagto")%></td>
		<td class="campop" align="center"><%=rs("contapagamento")%></td>
		<td class="campop" align="right"><%=formatnumber(recibo,2)%>&nbsp;&nbsp;</td>
	</tr>
<%
rs.movenext
loop
%>	
	</table>
</div>
<br>
<br>
<p align="justify" style="margin-top:0;margin-bottom:0;font-size:11pt;text-indent:300pt">
Atenciosamente
	
	


</td></tr>
<!-- final do corpo da carta -->

<!-- rodapé da carta -->
<tr><td height=80>
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
		<tr><td><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b></td> </tr>
		<tr><td><font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000</font></td></tr>
		<tr><td><font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP 06020-190 - Fone: (011) 3651-9999</font></td></tr>
		<tr><td><font face="Arial Narrow">Caixa Postal - ACF - Jd. Ipê - nº 2530 - Osasco - SP - CEP 06053-990</font></td></tr>
	</table>
</td></tr>
<!-- final do rodapé da carta -->
</table>
<!-- </div> -->
<%
rs.close
if giro<vezes then response.write "<DIV style=""page-break-after:always""></DIV>"
next ' giro de paginas
%>

<%
end if 
%>
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>