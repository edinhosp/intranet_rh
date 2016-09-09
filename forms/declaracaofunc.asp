<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a40")="N" or session("a40")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Declaração de Vínculo</title>
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
set rsi=server.createobject ("ADODB.Recordset")
Set rsi.ActiveConnection = conexao
set rsl=server.createobject ("ADODB.Recordset")
Set rsl.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
teste=0

if request("codigo")<>"" or request.form<>"" then
	if request.form="" then temp=request("codigo")
	if request("codigo")="" then temp=request.form("chapa")
	if isnumeric(temp) then
		info=1
		temp=numzero(temp,5)
		sqlb="AND f.CHAPA='" & temp & "' "
	else
		info=2
		sqlb="AND f.nome like '%" & temp & "%' order by f.nome"
	end if
	sqla="SELECT F.CHAPA, F.NOME, C.NOME AS FUNCAO, P.CARTEIRATRAB, P.SERIECARTTRAB, f.codtipo, P.cartidentidade, " & _
	"F.DATAADMISSAO, f.datademissao, P.SEXO, f.codsituacao, s.descricao as secao, f.codsecao, f.salario " & _
	"FROM corporerm.dbo.PFUNC F, corporerm.dbo.PPESSOA P, corporerm.dbo.PFUNCAO C, corporerm.dbo.psecao s " & _
	"WHERE F.CODPESSOA = P.CODIGO AND F.CODFUNCAO = C.CODIGO and f.codsecao=s.codigo "
	'tipod=request("tipo")
	'refd=request("ref")
	'topsal=request.form("topsal")
	
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	session("chapa")=rs("chapa")
	chapa=rs("chapa")
	nome=rs("nome"):admissao=rs("dataadmissao"):funcao=rs("funcao")
	temp=0
	if rs.recordcount>0 and session("cartateto")<>"L" then temp=2
	'tipo=request.form("R1")
	'if request.form("ref")="ON" then ref=1 else ref=0
	'if request.form("end")="ON" then ende=1 else ende=0
	'topsald=request("topsal")
	session("40tipo")=session("40tipo")
	if request.form("R1")="" then session("40tipo")=request("tipo")
	if request.form("R1")<>"" then session("40tipo")=request.form("R1")
	session("40topsal")=session("40topsal")
	if request.form("topsal")="" then session("40topsal")=request("topsal")
	if request.form("topsal")<>"" then session("40topsal")=request.form("topsal")
	session("40ende")=session("40ende")
	if request.form("ende")="" then session("40ende")=request("ende")
	if request.form("ende")<>"" then session("40ende")=request.form("ende")
	session("40ref")=session("40ref")
	if request.form("ref")="" then session("40ref")=request("ref")
	if request.form("ref")<>"" then session("40ref")=request.form("ref")
	session("40sal")=session("40sal")
	if request.form("salario")="" then session("40salario")=request("salario")
	if request.form("salario")<>"" then session("40salario")=request.form("salario")
	session("40tudo")=session("40tudo")
	if request.form("tudo")="" then session("40tudo")=request("tudo")
	if request.form("tudo")<>"" then session("40tudo")=request.form("tudo")
	session("40parag")=session("40parag")
	if request.form("parag")="" then session("40parag")=request("parag")
	if request.form("parag")<>"" then session("40parag")=request.form("parag")
	session("40printobs")=session("40printobs")
	if request.form("printobs")="" then session("40printobs")=request("printobs")
	if request.form("printobs")<>"" then session("40printobs")=request.form("printobs")
	session("40assinatura")=session("40assinatura")
	if request.form("assinatura")="" then session("40assinatura")=request("assinatura")
	if request.form("assinatura")<>"" then session("40assinatura")=request.form("assinatura")
	'session("decl_tipo")=tipo
	'session("decl_ref")=ref
	'session("decl_ende")=ende
	'session("decl_topsald")=topsal
else
	temp=1
end if

if temp=1 then
	session("cartateto")="F"
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção do funcionário para emissão de declaração
<form method="POST" action="declaracaofunc.asp" name="form">
<p style="margin-top: 0; margin-bottom: 0">
Chapa/Nome <input type="text" name="chapa" size="20" class="form_box" value="<%=session("chapa")%>">
</p>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Opções para a declaração</td>
</tr>
<tr>
	<td class=campo>
	<p style="margin-top: 0; margin-bottom: 1">
	<input type="radio" name="R1" value="1" checked>normal&nbsp;&nbsp;
	<input type="checkbox" name="ref" value="ON">com referência 
	<input type="checkbox" name="ende" value="ON">imprimir endereço
	<input type="checkbox" name="salario" value="ON">imprimir salário<br>
	<input type="radio" name="R1" value="2">não está de aviso prévio ou contrato de experiência<br>
	<input type="radio" name="R1" value="3">não está de aviso prévio/contrato de experiência mais os 
	<input type="text" name="topsal" size="3" class="form_box" value="3">últimos salários<br>
	<input type="radio" name="R1" value="4">com disciplinas que leciona (apenas para professor)<br>
	<input type="radio" name="R1" value="5">com listas das disciplinas que lecionou (apenas para professor)
<%if session("usuariomaster")="02379" or session("usuariomaster")="00259" then%>
	<input type="checkbox" name="tudo" value="ON">geral<br>
	<input type="checkbox" name="printobs" value="ON" checked>imprime rodapé de obs.
	<input type="checkbox" name="assinatura" value="ON">Imprime Assinatura
<%else%>
	<input type="checkbox" name="tudo" value="ON" disabled>geral<br>
	<input type="checkbox" name="printobs" value="ON" checked disabled>imprime rodapé de obs.
	<input type="checkbox" name="assinatura" value="ON">Imprime Assinatura
<%end if%>
	
	</td>
</tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0">
<input type="submit" value="Pesquisar" name="B1" class="button"></p>
<textarea name="parag" cols="50" rows="5" class=form_input>

</textarea>

</form>
<%
elseif temp=0 then
session("cartateto")="C"
'if request.form<>"" then
if rs("sexo")="F" then v1="a" else v1="o"
if rs("sexo")="F" then v2="a" else v2=""
if rs("sexo")="F" then v3="à" else v3="ao"
if rs("codtipo")="N" then tipof="funcionári" else tipof="estagiári"
%>
<div align="center"><center>
<table border="0" cellpadding="5" width="620" cellspacing="0" height="1000">
<!-- linha da aguia -->
<tr><td height=112><img src="../images/aguia.jpg" border="0" width="236" height="111" alt=""></td></tr>
<!-- linha intermediaria -->
<tr><td height="20">&nbsp;</td></tr>
<!-- linha declaracao -->
<tr><td height=50 valign="center" align="center"><b><font size="4">D E C L A R A Ç Ã O</font></b></td></tr>
<!-- linha intermediaria -->
<tr><td height="20">&nbsp;</td></tr>
<!-- corpo da declaracao -->
<tr><td height=100% valign=top>
<%
if rs("codsituacao")="D" or (rs("codsituacao")="A" and rs("datademissao")<>"") then
	texto1=" foi ":texto2=" de ":texto3=" a " & rs("datademissao")
else
	texto1=" é ":texto2=" desde ":texto3=""
end if
if session("40ende")="ON" then
	campus=left(rs("codsecao"),2)
	texto_end="."
	if campus="01" then texto_end=" e seu local de trabalho é R. Narciso Sturlini, 883 - Osasco"
	if campus="03" then texto_end=" e seu local de trabalho é Av. Franz Voegeli, 300 - Osasco"
	if campus="04" then texto_end=" e seu local de trabalho é Av. Franz Voegeli, 1743 - Osasco"
else
	texto_end=""
end if
if session("40salario")="ON" then
	remuneracao=" e percebe a remuneração bruta de R$ " & formatnumber(cdbl(rs("salario"))*1,2) & " (" & extenso2(cdbl(rs("salario"))*1) & ")"
else
	remuneracao=""
end if
%>	
	<p>&nbsp;</p>
	<p align="justify"><font size="3">Declaramos aos orgãos interessados, que <%=v1%> Sr<%=v2%>. <%=rs("nome")%>,
	portador<%=v2%> <%if rs("carteiratrab")<>"" or not isnull(rs("carteiratrab")) then%> da CTPS nº <%=rs("carteiratrab")%> / <%=rs("seriecarttrab")%>
	<%else%> do R.G. nº <%=rs("cartidentidade")%><%end if%>
	,<%=texto1%>
	<%=tipof%><%=v1%> desta Instituição de Ensino Superior <%=texto2%>
	<%=rs("dataadmissao")%><%=texto3%>, exercendo a função de <%=rs("funcao")%>, no
	departamento de <%=rs("secao")%><%=texto_end%><%=remuneracao%>.</font></p>

<%
if session("40parag")<>"" then
%>
	<p align="justify"><font size="3"><%=session("40parag")%></font></p>
<%
end if

if session("40tipo")="2" then
%>
	<p align="justify"><font size="3">Declaramos, também, que <%=v1%> referid<%=v1%>
	funcionári<%=v1%> não está na presente data em cumprimento de aviso prévio
	 ou em período de experiência.</font></p>
<%
end if
if session("40tipo")="3" then
	sql="select top " & session("40topsal") & " chapa, anocomp, mescomp, sum(baseinss) as baseinss from corporerm.dbo.pfperff where chapa='" & rs("chapa") & "' group by chapa, anocomp, mescomp order by anocomp desc, mescomp desc "
	rsi.open sql
%>
	<p align="justify"><font size="3">Declaramos, também, que <%=v1%> referid<%=v1%>
	funcionári<%=v1%> não está na presente data em cumprimento de aviso prévio
	ou em período de experiência, que sua data-base para correção salarial é o mês de Março e que nos últimos meses seus rendimentos foram os seguintes:</font></p>
	<div align="center"><center>
<!-- tabela dos ultimos salarios -->
	<table border="1" cellpadding="2" cellspacing="0"><tr>
		<td class=titulo align="center">Ano</td>
		<td class=titulo align="center">Mês</td>
		<td class=titulo align="center">Valor base/bruto</td>
		<td class=titulo align="center">Valor Líquido</td>
	</tr>
<%
	rsi.movefirst
	do while not rsi.eof
%>
	<tr>
		<td class=campo align="center">&nbsp;<%=rsi("anocomp")%></td>
		<td class=campo align="center">&nbsp;<%=rsi("mescomp")%></td>
<%
	sqlliquido="select f.chapa, sum(case when provdescbase='P' then valor else 0 end) as bruto, sum(case e.provdescbase when 'P' then f.valor*1 else f.valor*-1 end) liquido " & _
	"from corporerm.dbo.pffinanc f, corporerm.dbo.pevento e " & _
	"where f.chapa='" & rs("chapa") & "' and f.anocomp=" & rsi("anocomp") & " and f.mescomp=" & rsi("mescomp") & " " & _
	"and f.codevento=e.codigo and e.provdescbase in ('P','D') and f.codevento not in ('060','301','315','330','020') " & _
	"and f.nroperiodo=2 group by f.chapa"
	rsl.open sqlliquido, ,adOpenStatic, adLockReadOnly
	if rsl.recordcount>0 then 
		liquido=formatnumber(rsl("liquido"),2) 
		bruto=formatnumber(rsl("bruto"),2) 
	else 
		liquido="&nbsp;"
		bruto="&nbsp;"
	end if
%>
		<td class=campo align="center">&nbsp;<%=bruto%></td>
		<td class=campo align="center">&nbsp;<%=liquido%></td>
<%
	rsl.close
%>
	</tr>
<%
	rsi.movenext
	loop
	rsi.close
	set rsi=nothing
%>
	</table></center></div>
<!-- fim tabela ultimos salarios -->
<%
	set rsl=nothing
end if

'---------------
if session("40tipo")="4" then
'---------------

if rs("codsituacao")="D" then
	dataaula=rs("datademissao")
	texto=" que no período letivo anterior a data de " & dataaula & ", "
	ctexto="va"
	situacao="D"
else
	texto=" que atualmente, "
	dataaula=now()
	ctexto=""
	situacao="-"
end if
sqldisciplina="SELECT materia, curso, sum(ta) as ch from g2ch g inner join g2cursoeve c on c.coddoc=g.coddoc WHERE CHAPA1='" & rs("chapa") & "' " & _
"AND '" & dtaccess(dataaula) & "' between inicio and termino GROUP BY materia, curso "
rsd.open sqldisciplina, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then testedem=0 else testedem=1
rsd.close
if testedem=1 then
sqldisciplina="SELECT top 1 termino from g2ch WHERE CHAPA1='" & rs("chapa") & "' GROUP BY termino "
rsd.open sqldisciplina, ,adOpenStatic, adLockReadOnly
if rsd.recordcount=0 then dataaula=now() else dataaula=rsd("termino")
rsd.close
end if
%>
	<p align="justify"><font size="3">Declaramos, também, <%=texto%> <%=v1%> referid<%=v1%>
	professor<%=v2%>  ministra<%=ctexto%> as seguintes disciplinas nesta Instituição de
	Ensino Superior:</font></p>
<%
	sqldisciplina="SELECT materia, curso, count(ta) as ch from g2ch g inner join g2cursoeve c on c.coddoc=g.coddoc WHERE CHAPA1='" & rs("chapa") & "' " & _
	"AND '" & dtaccess(dataaula) & "' between inicio and termino GROUP BY materia, curso "
	rsd.open sqldisciplina, ,adOpenStatic, adLockReadOnly
	if rsd.recordcount>0 then
%>
<div align="center"><center>
<!-- tabela dos ultimos salarios -->
	<table border="1" bordercolor="#000000" cellpadding="1" cellspacing="0">
	<tr>
		<td class=titulo align="center">Disciplina</td>
		<td class=titulo align="center">Curso</td>
		<td class=titulo align="center">Carga horária semanal (horas)</td>
	</tr>
<%	
		rsd.movefirst
		do while not rsd.eof
%>
	<tr>
		<td class=campo>&nbsp;<%=rsd("materia")%></td>
		<td class=campo>&nbsp;<%=rsd("curso")%></td>
		<td class=campo align="center">&nbsp;<%=formatnumber(rsd("ch"),0)%></td>
	</tr>
<%
		rsd.movenext
		loop
	end if
	rsd.close
%>
	</table></center></div>
<%
'---------------
end if
'---------------


'---------------
if session("40tipo")="5" then
'---------------
if rs("codsituacao")="D" then
	dataaula=rs("datademissao")
	texto=" que no período letivo anterior a data de " & dataaula & ", "
	ctexto="va"
	situacao="D"
else
	texto=" que atualmente, "
	dataaula=now()
	ctexto=""
	situacao="-"
end if
%>
	<p align="justify"><font size="3">Declaramos, também, que durante a vigência do contrato de trabalho,
	<%=v1%> referid<%=v1%> professor<%=v2%>  ministrou as seguintes disciplinas nesta Instituição de
	Ensino Superior:</font></p>
<%
ultcur=""
sqldisciplina="SELECT DISTINCT g.coddoc, c.CURSO, materia FROM g2ch g " & _
"inner join g2cursos c on c.coddoc=g.coddoc and c.codcur=g.codcur and c.codper=g.codper " & _
"WHERE chapa1='" & rs("chapa") & "' order by g.coddoc, materia"
rsd.open sqldisciplina, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
	rsd.movefirst
	do while not rsd.eof
	if ultcur<>rsd("coddoc") and ultcur<>"" then 
		response.write " <i><u>no curso de " & ultcur2 & "</u></i>."
		response.write "<br>"
	elseif ultcur<>"" and (rsd.absoluteposition<>rsd.recordcount or rsd.absoluteposition=rsd.recordcount) then
		response.write ", "
	end if
	response.write rsd("materia")

	ultcur=rsd("coddoc"):ultcur2=rsd("curso")
	rsd.movenext
	loop
	response.write " <i><u>no curso de " & ultcur2 & "</u></i>."
end if
rsd.close
%>
<%
'---------------
end if
'---------------


if session("40ref")="ON" then
%>
	<p align="justify"><font size="3">Igualmente, declaramos que até o momento nada consta em nossos
	arquivos que possa desabonar a conduta d<%=v1%> referid<%=v1%> funcionári<%=v1%>.</font></p>
<%
end if
%>

	<p align="justify"><font size="3">Recebam nossas considerações.</font></p>
	<br>
	<p><font size="3">Atenciosamente</font></p>
	<br>

<!-- tabela data e assinatura -->
	<table border="0" cellpadding="0" width="100%" cellspacing="0">
	<tr>
<%if day(now())=1 then dia="1º" else dia=day(now())%>
		<td width="50%" valign="top">
		<p><font size="3">Osasco,&nbsp;<%=dia & " de " & monthname(month(now())) & " de " & year(now()) %></font></p>
<%
	if session("40assinatura")="ON" then
		teste=cint(left(rs("codsecao"),2))
%>
		<img src="../images/assinaturarmsa2.gif" height="96" border="0" alt="">
<%
else
%>
		<br><br>
		<p><font size="3">_____________________________________<br>
<%
end if
%>		
		</font></p>
		</td>
		<!-- carimbo cgc -->
<%if teste=1 then %>
		<td width="50%" valign="top">&nbsp;
		<div align="center"><center>
		<table border="0" cellpadding="0" width="240" cellspacing="0">
		<tr><td width="1" style="border-left-style: solid; border-left-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240" rowspan="2">
				<p align="center"><b><font size="4" color="#808080">73.063.166/0001-20</font></b></td>
			<td width="1" style="border-right-style: solid; border-right-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
		<tr><td width="1"></td><td width="1"></td></tr>
		<tr><td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td><td width="1"></td></tr>
		<tr><td width="1"></td><td width="240" align="center">
				<b><font color="#808080">FUNDAÇÃO INSTITUTO DE<br>ENSINO PARA OSASCO</font></b></td>
			<td width="1"></td></tr>
		<tr><td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td><td width="1"></td></tr>
		<tr><td width="1">&nbsp;</td><td width="240" rowspan="2" align="center">
				<font color="#808080">Rua Narciso Sturlini, 883<br>
				Jd. Umuarama - CEP 06018-903<br>OSASCO - SP</font></td><td width="1"></td></tr>
		<tr><td width="1" style="border-left-style: solid; border-left-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="1" style="border-right-style: solid; border-right-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
		</table></center></div>
		<p>&nbsp;
		</td>
<%end if%>
<%if teste=3 then %>
		<td width="50%" valign="top">&nbsp;
		<div align="center"><center>
		<table border="0" cellpadding="0" width="240" cellspacing="0">
		<tr><td width="1" style="border-left-style: solid; border-left-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240" rowspan="2">
				<p align="center"><b><font size="4" color="#808080">73.063.166/0003-92</font></b></td>
			<td width="1" style="border-right-style: solid; border-right-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
		<tr><td width="1"></td><td width="1"></td></tr>
		<tr><td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td><td width="1"></td></tr>
		<tr><td width="1"></td><td width="240" align="center">
				<b><font color="#808080">FUNDAÇÃO INSTITUTO DE<br>ENSINO PARA OSASCO</font></b></td>
			<td width="1"></td></tr>
		<tr><td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td><td width="1"></td></tr>
		<tr><td width="1">&nbsp;</td><td width="240" rowspan="2" align="center">
				<font color="#808080">Av. Franz Voegelli, 300<br>
				Vila Yara - CEP 06020-090<br>OSASCO - SP</font></td><td width="1"></td></tr>
		<tr><td width="1" style="border-left-style: solid; border-left-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="1" style="border-right-style: solid; border-right-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
		</table></center></div>
		<p>&nbsp;
		</td>
<%end if%>
		</tr>
	</table>
<!-- fim tabela assinatura/data -->

	</td>
</tr>
<!-- linha intermediaria -->
<tr><td height="20">&nbsp;</td></tr>
<tr><td height="1"><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
<tr><td height="1"><font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000 - C.N.P.J. 73.063.166/0001-20</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP 06020-190 - Fone: (011) 3651-9999 - C.N.P.J. 73.063.166/0003-92</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 1743 - Osasco - SP - CEP 06020-190 - Fone: (011) 3654-0655 - C.N.P.J. 73.063.166/0004-73</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Caixa Postal - ACF - Jd. Ipê - nº 2530 - Osasco - SP - CEP 06053-990</font></td></tr>
</table>
</center></div>

<%
rs.close

'****************************
if session("40tipo")="4" and session("40tudo")="ON" then
%>

<!-- inicio quadro graduação -->

<%
sqld="SELECT chapa1, curso, materia, inicio, termino, Sum(ta) AS aulas, periodo=case when inicio is null then perlet else convert(nvarchar,day(inicio))+'/'+convert(nvarchar,month(inicio))+'/'+convert(char(4),year(inicio)) + ' a ' + convert(nvarchar,day(termino))+'/'+convert(nvarchar,month(termino))+'/'+convert(char(4),year(termino)) end FROM g2ch g inner join g2cursoeve c on c.coddoc=g.coddoc " & _
"GROUP BY chapa1, curso, materia, inicio, termino, case when inicio is null then perlet else convert(nvarchar,day(inicio))+'/'+convert(nvarchar,month(inicio))+'/'+convert(char(4),year(inicio)) + ' a ' + convert(nvarchar,day(termino))+'/'+convert(nvarchar,month(termino))+'/'+convert(char(4),year(termino)) end " & _
"HAVING chapa1='" & chapa & "' ORDER BY chapa1, curso, materia, inicio;"
rsd.Open sqld, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
	response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="620" height="1000">
<tr><td class=campo valign="top" height="99%">
<!-------------- -->
<%
rsd.movefirst:do while not rsd.eof
if rsd.absoluteposition=1 or linha>48 then
	if linha>48 then
		response.write "</table>"
		response.write "<DIV style=""page-break-after:always""></DIV>"
	end if
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="620">
<tr><td class="campop" align="center" style="font-size:14pt">A N E X O&nbsp;&nbsp;I<br>DISCIPLINAS MINISTRADAS - GRADUAÇÃO</td></tr></table>
<br>

<table border="0" cellpadding="2" cellspacing="0" width="620" style="border-collapse: collapse">
<tr><td class=campo width=15% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">CHAPA</td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">NOME</td>
</tr>
<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000"><b><%=chapa%></td>
	<td class="campop" style="border-right:1px solid #000000"><b><%=nome%></td>
</tr></table>
<table border="0" cellpadding="2" cellspacing="0" width="620" style="border-collapse: collapse">
<tr><td class=campo width=20% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">ADMISSÃO</td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">CARGO ATUAL</td>
</tr>
<tr><td class="campop" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b><%=admissao%></td>
	<td class="campop" style="border-bottom:1px solid #000000;border-right:1px solid #000000"><b><%=funcao%></td>
</tr>
</table>
<br>

<table border="0" cellpadding="2" cellspacing="0" width="620" style="border-collapse: collapse">
<tr><td class=titulo align="center">Curso</td>
	<td class=titulo align="center">Disciplina</td>
	<td class=titulo align="center">Aulas<br>p/sem</td>
	<td class=titulo align="center">Período</td>
</tr>
<%
linha=0
end if
%>
<tr><td class=campo align="left"   style="border:1px solid #000000"><%=rsd("curso")%></td>
	<td class=campo align="left"   style="border:1px solid #000000"><%=rsd("materia")%></td>
	<td class=campo align="center" style="border:1px solid #000000"><%=rsd("aulas")%></td>
	<td class=campo align="center" style="border:1px solid #000000" nowrap><%=rsd("periodo")%></td>
</tr>
<%
if len(rsd("materia"))>40 or len(rsd("curso"))>30 then linha=linha+2 else linha=linha+1
rsd.movenext:loop
%>
</table>
</td></tr>
<% if session("40printobs")="ON" then %>
<tr><td class=titulo height="1%"><b>A finalidade deste documento é meramente informativa. A efetiva comprovação das aulas ministradas deve ser feita pelo professor através de outros documentos.</td></tr>
<% end if %>

</table></div> <!-- tabela borda -->
<%
response.write "</table>"
end if 'rsd.recordcount
rsd.close
%>

<!-- inicio quadro da pos-graduação -->
<%
sqld="SELECT chapa1, curso, materia, inicio, termino, Sum(ta) AS aulas, periodo=case when inicio is null then perlet else convert(nvarchar,day(inicio))+'/'+convert(nvarchar,month(inicio))+'/'+convert(char(4),year(inicio)) + ' a ' + convert(nvarchar,day(termino))+'/'+convert(nvarchar,month(termino))+'/'+convert(char(4),year(termino)) end FROM g5ch g inner join g2cursoeve c on c.coddoc=g.coddoc " & _
"GROUP BY chapa1, curso, materia, inicio, termino, case when inicio is null then perlet else convert(nvarchar,day(inicio))+'/'+convert(nvarchar,month(inicio))+'/'+convert(char(4),year(inicio)) + ' a ' + convert(nvarchar,day(termino))+'/'+convert(nvarchar,month(termino))+'/'+convert(char(4),year(termino)) end " & _
"HAVING chapa1='" & chapa & "' ORDER BY chapa1, curso, materia, inicio;"
rsd.Open sqld, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
	response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="620" height="1000">
<tr><td class=campo valign="top" height="99%">
<!-------------- -->
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="620">
<tr><td class="campop" align="center" style="font-size:14pt">A N E X O&nbsp;&nbsp;II<br>DISCIPLINAS MINISTRADAS - PÓS-GRADUAÇÃO</td></tr></table>
<br>

<table border="0" cellpadding="2" cellspacing="0" width="620" style="border-collapse: collapse">
<tr><td class=campo width=15% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">CHAPA</td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">NOME</td>
</tr>
<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000"><b><%=chapa%></td>
	<td class="campop" style="border-right:1px solid #000000"><b><%=nome%></td>
</tr></table>
<table border="0" cellpadding="2" cellspacing="0" width="620" style="border-collapse: collapse">
<tr><td class=campo width=20% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">ADMISSÃO</td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">CARGO ATUAL</td>
</tr>
<tr><td class="campop" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b><%=admissao%></td>
	<td class="campop" style="border-bottom:1px solid #000000;border-right:1px solid #000000"><b><%=funcao%></td>
</tr>
</table>
<br>

<table border="0" cellpadding="2" cellspacing="0" width="620" style="border-collapse: collapse">
<tr><td class=titulo align="center">Curso</td>
	<td class=titulo align="center">Disciplina</td>
	<td class=titulo align="center">Aulas<br>p/sem</td>
	<td class=titulo align="center">Período</td>
</tr>
<%
rsd.movefirst:do while not rsd.eof
%>
<tr><td class=campo align="left"   style="border:1px solid #000000"><%=rsd("curso")%></td>
	<td class=campo align="left"   style="border:1px solid #000000"><%=rsd("materia")%></td>
	<td class=campo align="center" style="border:1px solid #000000"><%=rsd("aulas")%></td>
	<td class=campo align="center" style="border:1px solid #000000" nowrap><%=rsd("periodo")%></td>
</tr>
<%
rsd.movenext:loop
%>
</table>
</td></tr>
<% if session("40printobs")="ON" then %>
<tr><td class=titulo height="1%"><b>A finalidade deste documento é meramente informativa. A efetiva comprovação das aulas ministradas deve ser feita pelo professor através de outros documentos.</td></tr>
<% end if %>

</table></div> <!-- tabela borda -->
<%
response.write "</table>"
end if 'rsd.recordcount
rsd.close
%>

<!-- inicio quadro das nomeações -->

<%
sqld="SELECT i.CHAPA, n.NOMEACAO, i.CARGO, i.PORTARIA, i.entrega, i.MAND_INI, i.MAND_FIM, i.CH AS aulas, " & _
"periodo=case when mand_ini is null then 'Não disponível' else convert(nvarchar,day(mand_ini))+'/'+convert(nvarchar,month(mand_ini))+'/'+convert(char(4),year(mand_ini)) + ' a ' + convert(nvarchar,day(mand_fim))+'/'+convert(nvarchar,month(mand_fim))+'/'+convert(char(4),year(mand_fim)) end " & _
"FROM n_indicacoes AS i INNER JOIN n_nomeacoes AS n ON i.id_nomeacao = n.id_nomeacao " & _
"WHERE i.CHAPA='" & chapa & "' ORDER BY i.CHAPA, i.MAND_INI;"

rsd.Open sqld, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
	response.write "<DIV style=""page-break-after:always""></DIV>"
%>
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="620" height="1000">
<tr><td class=campo valign="top" height="99%">
<!-------------- -->
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="620">
<tr><td class="campop" align="center" style="font-size:14pt">A N E X O&nbsp;&nbsp;III<br>OUTRAS ATIVIDADES</td></tr></table>
<br>

<table border="0" cellpadding="2" cellspacing="0" width="620" style="border-collapse: collapse">
<tr><td class=campo width=15% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">CHAPA</td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">NOME</td>
</tr>
<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000"><b><%=chapa%></td>
	<td class="campop" style="border-right:1px solid #000000"><b><%=nome%></td>
</tr></table>
<table border="0" cellpadding="2" cellspacing="0" width="620" style="border-collapse: collapse">
<tr><td class=campo width=20% style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">ADMISSÃO</td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">CARGO ATUAL</td>
</tr>
<tr><td class="campop" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000"><b><%=admissao%></td>
	<td class="campop" style="border-bottom:1px solid #000000;border-right:1px solid #000000"><b><%=funcao%></td>
</tr>
</table>
<br>

<table border="0" cellpadding="2" cellspacing="0" width="620" style="border-collapse: collapse">
<tr><td class=titulo align="center">Atividade designada</td>
	<td class=titulo align="center">Documento de Autorização</td>
	<td class=titulo align="center">Horas<br>p/sem</td>
	<td class=titulo align="center">Período</td>
</tr>
<%
rsd.movefirst:do while not rsd.eof
if (rsd("entrega")="" or isnull(rsd("entrega"))) and situacao<>"D" then estilo="text-decoration:line-through;color:gray" else estilo=""
if session("40printobs")<>"ON" then estilo=""
%>
<tr><td class=campo align="left"   style="border:1px solid #000000;<%=estilo%>"><%=rsd("nomeacao")%></td>
	<td class=campo align="left"   style="border:1px solid #000000;<%=estilo%>"><%=rsd("portaria")%></td>
	<td class=campo align="center" style="border:1px solid #000000;<%=estilo%>">&nbsp;<%=rsd("aulas")%></td>
	<td class=campo align="center" style="border:1px solid #000000;<%=estilo%>" nowrap><%=rsd("periodo")%></td>
</tr>
<%
rsd.movenext:loop
%>
</table>
</td></tr>
<% if session("40printobs")="ON" then %>
<tr><td class=titulo height="1%"><b>A finalidade deste documento é meramente informativa. A efetiva comprovação das aulas ministradas deve ser feita pelo professor através de outros documentos.</td></tr>
<% end if %>
</table></div> <!-- tabela borda -->
<%
response.write "</table>"
end if 'rsd.recordcount
rsd.close
%>

<!-- final dos quadros -->

<%
end if 'tipo 4 = anexo disciplinas
'****************************

elseif temp=2 then
session("cartateto")="L"
%>
<!-- mostrar funcionarios e as contribuições -->
<table border="1" cellpadding="0" width="550" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>&nbsp;Chapa</td>
	<td class=titulo>&nbsp;Nome</td>
    <td class=titulo>&nbsp;Situacao</td>
</tr>
<%
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo>&nbsp;<%=rs("chapa")%></td>
	<td class=campo>&nbsp;<a href="declaracaofunc.asp?codigo=<%=rs("chapa")%>&tipo=<%=session("40tipo")%>&ref=<%=session("40ref")%>&ende=<%=session("40ende")%>&topsal=<%=session("40topsal")%>&salario=<%=session("40salario")%>&tudo=<%=session("40tudo")%>&printobs=<%=session("40printobs")%>&assinatura=<%=session("40assinatura")%>&parag=<%=session("40parag")%>"><%=rs("nome")%></a></td>
	<td class=campo>&nbsp;<%=rs("codsituacao")%></td>
</tr>
<%
rs.movenext
loop
%>
</table>
<%
rs.close
end if ' temps

set rsd=nothing
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>