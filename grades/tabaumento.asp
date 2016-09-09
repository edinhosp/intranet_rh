<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a80")="N" or session("a80")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Checagem de Atribuição de Aulas</title>
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
</head>
<body>
<%
	dim conexao, rs, rs2
	set conexao=server.createobject ("ADODB.Connection")
	conexao.Open application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	Set rs.ActiveConnection = conexao
	set rs2=server.createobject ("ADODB.Recordset")
	Set rs2.ActiveConnection = conexao
	if request.form<>"" then
		if request.form("B3")<>"" then
			finaliza=1
		else
			finaliza=0
		end if
	end if
	
if finaliza=0 then
%>
<p class=titulo>Seleção para impressão de Checagem de atribuição de carga horária - aumento</p>
<form method="POST" action="tabaumento.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="400">
<tr><td class=titulo colspan=2>Curso Graduação</td></tr>
<tr><td class=titulo colspan=2><select size="1" name="coddoc">
	<option value="0" selected>Selecione um curso</option>
<%
sqla="SELECT c.coddoc, c.curso from g2cursoeve c, grades_per p where p.coddoc=c.coddoc and p.coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') GROUP BY c.coddoc, c.curso order by c.curso "
sqla="SELECT c.coddoc, c.curso from g2cursoeve c, g2turmas t where t.coddoc=c.coddoc and c.coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') GROUP BY c.coddoc, c.curso order by c.curso "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof 
%>
<option <%if request.form("codcur")=rs("coddoc") then response.write "selected "%> value="<%=rs("coddoc")%>"><%=rs("curso")%> (<%=rs("coddoc")%>)</option>
<%
rs.movenext
loop
rs.close
%>  
	</select></td></tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="400">
<tr><td align="center" class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3"></td></tr>
</table>
</form>
<hr>
<%
end if 'finaliza 0

'******************************** inicio impressao
if finaliza=1 then
	inicio=now()
	coddoc=request.form("coddoc")
	periodos=0
	
periodo1="20092":periodo2="20101"
	sql1="SELECT t.chapa1, f.NOME, f.DATAADMISSAO, t.coddoc, t.curso, t.[" & periodo1 & "] as ant, t.[" & periodo2 & "] as atual, f.CODSITUACAO, " & _
	"variacao=(case when [" & periodo2 & "]=null then 0 else [" & periodo2 & "] end) - (case when [" & periodo1 & "]=null then 0 else [" & periodo1 & "] end) " & _
	"FROM totalizador_chor_cur AS t, corporerm.dbo.pfunc AS f " & _
	"WHERE t.chapa1=f.CHAPA collate database_default AND t.coddoc='" & coddoc & "' AND f.CODSITUACAO In ('A','Z','F') " & _
	"and ((case when [" & periodo2 & "]=null then 0 else [" & periodo2 & "] end) - (case when [" & periodo1 & "]=null then 0 else [" & periodo1 & "] end))>0 "
	'response.write "<br>" & sql1
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	'response.write "<br>" & rs.recordcount
dtfinal="01/31/2010"
dtinicio="02/01/2010"
'dim volta(3), fimper(3)
'volta(0)="20031":fimper(0)=dateserial(2003,7,31)
'volta(1)="20032":fimper(1)=dateserial(2004,1,31)
'volta(2)="20041":fimper(2)=dateserial(2004,7,31)
'volta(3)="20042":fimper(3)=dateserial(2005,1,31)
%>
<table border="1" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=980>
<tr><td class=grupo colspan=9>Atribuição de Aulas - Checagem de aumento de carga horária</td></tr>
<tr><td class="campor" colspan=9>
Segundo a cláusula 20 da Convenção Coletiva dos Professores, o professor só pode 
ter sua carga horária reduzida em razão de supressão de disciplina, classe ou turma, em virtude de alteração na estrutura 
curricular prevista, quando a carga horária de algumas disciplinas são diminuídas ou extintas. Ocorrendo isto, o professor 
da disciplina terá prioridade para preenchimento de vaga existente em outra disciplina na qual possua habilitação legal.
<br>A redução de carga horária fora destas situações, <b>constitui redução salarial ilegal</b>, 
segundo o artigo 478 da CLT e o Precedente Normativo do TST nº 78, sujeitando a Fundação e os responsáveis à sanções legais.
</td></tr>
<%
if rs.recordcount>0 then
%>
<tr>
	<td class=titulo colspan=9>CURSO DE <%=rs("curso")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<font size=1><%=rs("coddoc")%></font></td>
</tr>
<tr>
	<td class=titulo align="center" rowspan=2>Chapa</td>
	<td class=titulo align="center" rowspan=2>Nome</td>
	<td class=titulo align="center" colspan=2><%=right(periodo1,1)%>º sem/<%=left(periodo1,4)%></td>
	<td class=titulo align="center" colspan=2><%=right(periodo2,1)%>º sem/<%=left(periodo2,4)%></td>
</tr>
<tr>
	<td class=titulo align="center">#</td>
	<td class=titulo align="center">Aulas</td>
	<td class=titulo align="center">#</td>
	<td class=titulo align="center">Aulas</td>
</tr>
<%
rs.movefirst
do while not rs.eof
'	if rs2("turno")="1" then turno="Mat"
'	if rs2("turno")="2" then turno="Vesp"
'	if rs2("turno")="3" then turno="Not"
'	if rs2("turno")="5" then turno="Vesp-EF"
%>	
<tr>
	<td class="campol" align="center" style="font-size:7pt"><%=rs("chapa1")%></td>
	<td class="campol" align="left" style="font-size:7pt"><%=rs("nome")%></td>
	<td class="campol" align="center" style="font-size:7pt" valign="top"><%=rs("ant")%></td>
	<td class="campol" align="left" style="font-size:7pt" valign="top">
<%
sqlsem1="select codtur, turno, serie, turma, codmat, materia, sum(ta) as aulas from g2ch g " & _
"where chapa1='" & rs("chapa1") & "' and coddoc='" & coddoc & "' and '" & dtfinal & "' between inicio and termino and deletada=0 and ativo=1 " & _
"group by codtur, turno, serie, turma, codmat, materia "
rs2.Open sqlsem1, ,adOpenStatic, adLockReadOnly
numero=0
if rs2.recordcount>0 then
	rs2.movefirst
	do while not rs2.eof
	response.write rs2("codtur")
	response.write " - "
	'if lastnome=rs2("nome") then response.write "  o mesmo  " else response.write rs2("nome")
	response.write rs2("materia")
	response.write " ("
	response.write rs2("aulas") & ")"
	'if fimper(a)<>cdate(rs2("termino")) then response.write " até " & rs2("termino")
	if rs2.recordcount>1 and rs2.absoluteposition<rs2.recordcount then response.write "<br>"
	'lastnome=rs2("nome")
		redim preserve turma(numero)
		redim preserve materia(numero)
		redim preserve nmateria(numero)
		redim preserve checagem(numero)
		turma(numero)=rs2("codtur")
		materia(numero)=rs2("codmat"):nmateria(numero)=rs2("materia")
		checagem(numero)=trim(rs2("codtur") & rs2("codmat"))
	rs2.movenext
		numero=numero+1
	loop
else
	response.write "-"
	redim preserve turma(numero)
	redim preserve materia(numero)
	redim preserve nmateria(numero)
	redim preserve checagem(numero)
	turma(numero)=""
	materia(numero)="":nmateria(numero)=""
	checagem(numero)=""
end if
rs2.close
'next
%>	
	</td>
	<td class="campol" align="center" style="font-size:7pt" valign="top"><%=rs("atual")%></td>
	<td class="campol" align="left" style="font-size:7pt" valign="top">
<%
sqlsem1="select codtur, turno, serie, turma, codmat, materia, sum(ta) as aulas from g2ch g " & _
"where chapa1='" & rs("chapa1") & "' and coddoc='" & coddoc & "' and '" & dtinicio & "' between inicio and termino and deletada=0 and ativo=1 " & _
"group by codtur, turno, serie, turma, codmat, materia "
rs2.Open sqlsem1, ,adOpenStatic, adLockReadOnly
numero2=0
if rs2.recordcount>0 then
	rs2.movefirst
	do while not rs2.eof
	response.write rs2("codtur")
	response.write " - "
	'if lastnome=rs2("nome") then response.write "  o mesmo  " else response.write rs2("nome")
	response.write rs2("materia")
	response.write " ("
	response.write rs2("aulas") & ")"
	'if fimper(a)<>cdate(rs2("termino")) then response.write " até " & rs2("termino")
	if rs2.recordcount>1 and rs2.absoluteposition<rs2.recordcount then response.write "<br>"
	'lastnome=rs2("nome")
	checagem2=trim(rs2("codtur") & rs2("codmat"))
	redim preserve turma2(numero2)
	redim preserve materia2(numero2)
	redim preserve nmateria2(numero2)
	redim preserve checagemf(numero2)
	turma2(numero2)=rs2("codtur")
	materia2(numero2)=rs2("codmat"):nmateria2(numero2)=rs2("materia")
	checagemf(numero2)=rs2("codtur") & rs2("codmat")
	numero2=numero2+1
	rs2.movenext
	loop
else
	redim preserve turma2(numero2)
	redim preserve materia2(numero2)
	redim preserve nmateria2(numero2)
	redim preserve checagemf(numero2)
	turma2(numero2)=""
	materia2(numero2)="":nmateria2(numero2)=""
	checagemf(numero2)=""
	'response.write "-"
end if
rs2.close
'next
%>	
<!--
<hr style="margin-top:0;margin-bottom:0">
-->
<br><b>Aulas suprimidas:</b>
<%
maximo=0
for a=0 to ubound(checagem)
	redim preserve diminuida(maximo)
	diminuida(a)=checagem(a)
	maximo=maximo+1
next

for check=0 to ubound(checagem)
	for recheck=0 to ubound(checagemf)
		if checagem(check)=checagemf(recheck) then diminuida(check)=""
	next
next

for a=0 to ubound(diminuida)
	if diminuida(a)<>"" then
		sqlsupri="select chapa1, f.nome, materia, codtur, codmat, f.codsituacao, sum(ta) as aulas from g2ch g, corporerm.dbo.pfunc f " & _
		"where coddoc='" & coddoc & "' and left(codtur,7)='" & left(turma(a),7) & "' and codmat='" & materia(a) & "' and '" & dtinicio & "' between inicio and termino and f.chapa collate database_default=g.chapa1  and deletada=0 and ativo=1 " & _
		"group by chapa1, f.nome, materia, codtur, codmat, f.codsituacao "
		rs2.Open sqlsupri, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then
			rs2.movefirst
			do while not rs2.eof
				response.write "<br>" & rs2("codtur") & " - " & rs2("materia") & " -foi para-> <b>" & rs2("nome") & "</b> (" & rs2("aulas") & ") (" & rs2("codsituacao") & ")"
			rs2.movenext
			loop
		else		
			response.write "<br>" & turma(a) & " - " & nmateria(a) & " -> não formou turma"
		end if
		rs2.close
	end if
next
%>
<!--
<hr style="margin-top:0;margin-bottom:0">
-->
<Br><b>Aulas incluídas:</b>
<%
maximo=0
for a=0 to ubound(checagemf)
	redim preserve aumentada(maximo)
	aumentada(a)=checagemf(a)
	maximo=maximo+1
next

for check=0 to ubound(checagemf)
	for recheck=0 to ubound(checagem)
		if checagemf(check)=checagem(recheck) then aumentada(check)=""
	next
next

for a=0 to ubound(aumentada)
	if aumentada(a)<>"" then
		sqlsupri="select chapa1, f.nome, materia, codtur, codmat, f.codsituacao, sum(ta) as aulas from g2ch g, corporerm.dbo.pfunc f " & _
		"where coddoc='" & codcur & "' and left(codtur,7)='" & left(turma2(a),7) & "' and codmat='" & materia2(a) & "' and '" & dtfinal & "' between inicio and termino and f.chapa collate database_default=g.chapa1  and deletada=0 and ativo=1 " & _
		"group by chapa1, f.nome, materia, codtur, codmat, f.codsituacao "
		rs2.Open sqlsupri, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then
			rs2.movefirst
			do while not rs2.eof
				response.write "<br>" & rs2("codtur") & " - " & rs2("materia") & " -era de-> <b>" & rs2("nome") & "</b> (" & rs2("aulas") & ") (" & rs2("codsituacao") & ")"
			rs2.movenext
			loop
		else		
			response.write "<br>" & turma2(a) & " - " & nmateria2(a) & " -> turma nova"
		end if
		rs2.close
	end if
next
%>
	</td>
</tr>

<%
rs.movenext
loop
else 'recordcount
%>
<tr>
	<td class=titulo colspan=9>Não houve aumento de carga horária dos professores</td>
</tr>

<%
end if ' recordcount
rs.close
%>
</table>

<%	
termino=now()
duracao=termino-inicio
response.write right(formatdatetime(duracao,3),5)
end if 'finaliza 1

set rs=nothing
set rs2=nothing
'set rs3=nothing
'set rs4=nothing
conexao.close
'conexao2.close
set conexao=nothing
'set conexao2=nothing
%>
</body>
</html>