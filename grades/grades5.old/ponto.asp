<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a91")="N" or session("a91")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Grade Horária</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:40px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set conexao2=server.createobject ("ADODB.Connection")
conexao2.Open application("consql")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
set rsq=server.createobject ("ADODB.Recordset")
Set rsq.ActiveConnection = conexao2
dim emes(12),edia(31)
mesagora=month(now)
anoagora=year(now)
diaagora=day(now)

	if request("d")<>"" and request.form("diaagora")="" then diaagora=request("d") 
	if request("m")<>"" and request.form("mesagora")="" then mesagora=request("m")
	if request("a")<>"" and request.form("anoagora")="" then anoagora=request("a")
	if request("c")<>"" and request.form("codcur")  ="" then codcur  =request("c")
	if request("o")<>"" and request.form("opcao")  =""  then opcao   =request("o")
	if request("or")<>"" and request.form("ordem")  =""  then ordem  =request("or")
	if request.form("diaagora")<>"" and request("d")="" then diaagora=request.form("diaform")
	if request.form("mesagora")<>"" and request("m")="" then mesagora=request.form("mesform")
	if request.form("anoagora")<>"" and request("a")="" then anoagora=request.form("anoform")
	if request.form("codcur")  <>"" and request("c")="" then codcur  =request.form("codcur")
	if request.form("opcao")   <>"" and request("o")="" then opcao   =request.form("opcao")
	if request.form("ordem")   <>"" and request("or")="" then ordem  =request.form("ordem")
		
	if request.form<>"" then
		if request.form("B3")<>"" then
			finaliza=1
		else
			finaliza=0
			mesagora=request.form("mesform")
			anoagora=request.form("anoform")
			diaagora=request.form("diaform")
		end if
		if request.form("avanca")<>"" then
			mesagora=mesagora+1
			if mesagora>12 then
				mesagora=1
				anoagora=anoagora+1
			end if
		end if
		if request.form("volta")<>"" then
			mesagora=mesagora-1
			if mesagora<1 then
				mesagora=12
				anoagora=anoagora-1
			end if
		end if
		if request.form("avancay")<>"" then anoagora=anoagora+1
		if request.form("voltay")<>"" then anoagora=anoagora-1
	end if
	if opcao<>"" then
		turno=left(opcao,1)
		grade=right(opcao,1)
	else
		turno=""
		grade=""
	end if

	sqld="select day(diaferiado) as dia1 from corporerm.dbo.gferiado " & _
	"where month(diaferiado)=" & mesagora & " and year(diaferiado)=" & anoagora & " " & _
	"group by day(diaferiado) "
	rsq.Open sqld, ,adOpenStatic, adLockReadOnly
	if rsq.recordcount>0 then
	rsq.movefirst:do while not rsq.eof 
		edia(rsq("dia1"))=1
	rsq.movenext:loop
	end if
	rsq.close

if finaliza=0 then
%>
<p class=titulo>Seleção para impressão do livro de ponto</p>
<form method="POST" action="ponto.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=450>
<tr><td class=titulo>Curso:</td>
	<td class=titulo>Tipo:</td>
</tr>
<tr><td class=titulo><select size="1" name="codcur" onChange="javascript:submit()">
	<option value="0" selected>Selecione um curso</option>
<%
sqla="SELECT coddoc, curso from grades_5 GROUP BY coddoc, curso order by curso "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof 
%>
<option <%if codcur=rs("coddoc") then response.write "selected "%> value="<%=rs("coddoc")%>"><%=rs("curso")%> (<%=rs("coddoc")%>)</option>
<%
rs.movenext
loop
rs.close
if codcur="" then codcur=0
%>  
	</select></td>
<td class=titulo><select size="1" name="opcao" onChange="javascript:submit()">
<%
sqla="SELECT tipo from grades_5chi where coddoc='" & codcur & "' GROUP BY tipo order by tipo "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 
if rs.absoluteposition=1 and opcao="0" then opcao=rs("tipo")
%>
<option <%if opcao=rs("tipo") then response.write "selected "%> value="<%=rs("tipo")%>"><%=rs("tipo")%></option>
<%
rs.movenext
loop
else
%>
	<option value="0" selected>Selecione um tipo</option>
<%
end if 'recordcount >0
rs.close
%>  
	</select></td>
</tr>
</table>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=450>
<tr><td class=titulo width=200>
<!-- calencario -->
<%
	diasemana=weekday(dateserial(anoagora,mesagora,1))
	ultimodia=day(dateserial(anoagora,mesagora+1,1)-1)
	ultimo=0
	emes(1)="Janeiro":emes(2)="Fevereiro":emes(3)="Março":emes(4)="Abril":emes(5)="Maio":emes(6)="Junho"
	emes(7)="Julho":emes(8)="Agosto":emes(9)="Setembro":emes(10)="Outubro":emes(11)="Novembro":emes(12)="Dezembro"
%>
<input type="hidden" name="mesform" value="<%=mesagora%>">
<input type="hidden" name="anoform" value="<%=anoagora%>">
<input type="hidden" name="diaform" value="<%=diaagora%>">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=175>
<tr>
	<td class=campo><input type="submit" value="<<" name="voltay" class=button></td>
	<td class=campo><input type="submit" value="<" name="volta" class=button></td>
	<td class="campor" width="100%" align="center">
		<font color="#000080"><b><%=emes(mesagora)& "/" & anoagora%></font></td>
	<td class=campo><input type="submit" value=">" name="avanca" class=button></td>
	<td class=campo><input type="submit" value=">>" name="avancay" class=button></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=175>
<tr>
	<td class=campo align="center">Dom</td>
	<td class=campo align="center">Seg</td>
	<td class=campo align="center">Ter</td>
	<td class=campo align="center">Qua</td>
	<td class=campo align="center">Qui</td>
	<td class=campo align="center">Sex</td>
	<td class=campo align="center">Sab</td>
</tr>
<tr>
<%
for linha=1 to 7
	response.write "<td class=campo align='center'>"
	if linha=diasemana then
		ultimo=1
		if edia(ultimo)=1 or linha=1 then 'é feriado
			response.write "<font color='#FF0000'>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</font>"
		else
			if dateserial(anoagora,mesagora,ultimo)<=int(now+10) then
				response.write "<a href='ponto.asp?d=" & ultimo & "&m=" & mesagora & "&a=" & anoagora & "&c=" & codcur & "&o=" & opcao  & "&or=" & ordem & "' class=r>"
			end if
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		end if		
	elseif ultimo>=1 then
		ultimo=ultimo+1
		if edia(ultimo)=1 then
			response.write "<font color='#FF0000'>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</font>"
		else
			if dateserial(anoagora,mesagora,ultimo)<=int(now+10) then
				response.write "<a href='ponto.asp?d=" & ultimo & "&m=" & mesagora & "&a=" & anoagora & "&c=" & codcur & "&o=" & opcao & "&or=" & ordem & "' class=r>"
			end if
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		end if		
	end if
	response.write "</td>"
next
response.write "</tr>"

vartemp1=ultimodia-ultimo
vartemp2=int(vartemp1/7)
if (vartemp1/7)-vartemp2>0 then vartemp2=vartemp2+1
for sem=1 to vartemp2
	response.write "<tr>"
	for l2=1 to 7
		response.write "<td class=campo align='center'>"
		ultimo=ultimo+1
		if ultimo<=ultimodia then 
			if edia(ultimo)=1 or l2=1 then
				response.write "<font color='#FF0000'>"
				if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				response.write "</font>"
			else
				if dateserial(anoagora,mesagora,ultimo)<=int(now+10) then
					response.write "<a href='ponto.asp?d=" & ultimo & "&m=" & mesagora & "&a=" & anoagora & "&c=" & codcur & "&o=" & opcao  & "&or=" & ordem & "' class=r>"
				end if
				if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				response.write "</a>"
			end if
		end if
		response.write "</td>"
	next
	response.write "</tr>"
next
dataponto=dateserial(anoagora,mesagora,diaagora)

grade=0
if opcao="Convidados" then grade=2
if opcao="CLT" then grade=1
'sql2="select top 1 data, dialetivo, pagina, limite from grades_dias " & _
'"where codcur=" & codcur & " and turno=" & turno & " and grade=" & grade & " " & _
'"ORDER BY codcur, turno, grade, limite, data DESC "
sql2="select top 1 data, dialetivo, pagina, limite from grades_dias " & _
"where coddoc='" & codcur & "' and grade=" & grade & " " & _
"ORDER BY coddoc, turno, grade, limite, data DESC "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
	idata=rsc("data")
	idialetivo=rsc("dialetivo")
	pdialetivo=0 'idialetivo+1
	ipagina=rsc("pagina")
	ppagina=ipagina+1
	ilimite=rsc("limite")
else
	idata=""
	idialetivo=""
	ipagina=""
	ilimite=""
end if
rsc.close
'sql2="select top 1 data, dialetivo, pagina, limite from grades_dias " & _
'"where codcur=" & codcur & " and turno=" & turno & " and grade=" & grade & " " & _
'"and data=#" & dtaccess(dataponto) & "# " & _
'"ORDER BY codcur, turno, grade, limite, data DESC "
sql2="select top 1 data, dialetivo, pagina, limite from grades_dias " & _
"where coddoc='" & codcur & "' and grade=" & grade & " " & _
"and data='" & dtaccess(dataponto) & "' " & _
"ORDER BY coddoc, turno, grade, limite, data DESC "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
	ppagina=cdbl(rsc("pagina"))
	pdialetivo=cdbl(rsc("dialetivo"))
end if
rsc.close
if request.form("pg")<>"" then ppagina=request.form("pg")
if request.form("dl")<>"" then pdialetivo=0 'request.form("dl")
if ordem="" then ordem="nome"
%>
</table>
<!-- fim calencario -->
</td>
<td class=titulo valign=top>

	<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
	<tr><td align="center" class=grupo colspan=3>Ultima Impressão</td></tr>
	<tr><td align="center" class=fundo>&nbsp;Data&nbsp;</td>
	<td align="center" class=fundo>&nbsp;Dia Letivo&nbsp;</td>
	<td align="center" class=fundo>&nbsp;Página&nbsp;</td></tr>
	<tr><td align="center" class=campo>&nbsp;<%=idata%></td>
	<td align="center" class=campo>&nbsp;<%=idialetivo%></td>
	<td align="center" class=campo>&nbsp;<%=ipagina%></td></tr>
	</table>

Dia Selecionado:
<br><b><font color='#0000FF'><%=diaagora%>/<%=monthname(mesagora)%>/<%=anoagora%> (<%=weekdayname(weekday(dataponto))%>)</font></b>

	<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
	<tr><td class=titulo>&nbsp;Dia Letivo&nbsp;</td>
	<td	class=titulo>&nbsp;Página&nbsp;</td></tr>
	<tr><td class=titulo><input type="text" name="dl" size="5" maxlength="5" class=a value="<%=pdialetivo%>"></td>
	<td class=titulo><input type="text" name="pg" size="5" maxlength="5" class=a value="<%=ppagina%>"></td></tr>
	</table>
</td>
</tr></table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=450>
<tr><td align="center" class=titulo>
Imprimir a folha em ordem: 
<input type="radio" name="ordem" value="nome" <%if ordem="nome" then response.write "checked"%>> de nome
<input type="radio" name="ordem" value="turma" <%if ordem="turma" then response.write "checked"%>> de turma<br>
	</td></tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=450>
<tr><td align="center" class=titulo>
	<input type="submit" value="Clique para Visualizar" class="button" name="B3"></td></tr>
</table>
</form>
<hr>
Observação: Temporariamente, não está sendo feito controle de dia letivo e página impressa.<br>
Por gentileza, preencher manualmente.
<%
end if 'finaliza=0

'******************************** inicio impressao
if finaliza=1 then
	codcur=request.form("codcur")
	'turno=left(request.form("opcao"),1)
	'grade=right(request.form("opcao"),1)
	mes=request.form("mesform"):ano=request.form("anoform"):dia=request.form("diaform")
	dialetivo=request.form("dl")
	pagina=request.form("pg")
	opcao=request.form("opcao")
	periodoletivo=ano & "/"
	dataponto=dateserial(ano,mes,dia)	
	diasemana=weekday(dataponto)
	if pagina="" then pagina=2
	if dialetivo="" then dialetivo=1
	if pagina=1 then pagina=2

if turno="6" then turno="61,62" else turno=turno	
if opcao="CLT" then tipo=1
if opcao="Convidados" then tipo=2
sqla="SELECT gch.perlet, gch.perlet2, gch.perlet3, gch.coddoc, gch.curso, gch.turno, gch.serie, gch.turma, " & _
"gch.codtur, gch.diasem, gch.codmat, gch.materia, gch.chapa1, f.NOME, gch.a1, gch.a2, gch.a3, gch.a4, gch.a5, gch.a6, " & _
"gp.diretor, gp.coordenador, gp.chefedepto " & _
"FROM ((grades_5chi AS gch INNER JOIN grades_per AS gp ON (gch.coddoc=gp.coddoc) AND (gch.perlet=gp.perlet) AND (gch.perlet2=gp.perlet2)) " & _
"INNER JOIN grades_aux_prof AS f ON gch.chapa1=f.CHAPA) " & _
"INNER JOIN grades_gc ON (gp.perlet=grades_gc.perlet) AND (gch.serie=grades_gc.serie) AND (gp.coddoc=grades_gc.coddoc) " & _
"WHERE gch.coddoc='" & codcur & "' AND gch.diasem=DATEPART (dw,'" & dtaccess(dataponto) & "') and gch.tipo='" & opcao & "' " & _
"AND '" & dtaccess(dataponto) & "' Between [inicio] And [termino] AND '" & dtaccess(dataponto) & "' Between [pini] And [pfim] "
if request.form("ordem")="turma" then sqla=sqla & "ORDER BY gch.turno, gch.serie, gch.turma, a5,a3,a1 "
if request.form("ordem")="nome" then sqla=sqla & "ORDER BY gch.turno, f.nome "
'response.write sqla

rs.Open sqla, ,adOpenStatic, adLockReadOnly
tamanho=640
if rs.recordcount>0 then
rs.movefirst

'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'rs.movefirst
'*************** fim teste **********************
diretor=rs("diretor"):coordenador=rs("coordenador"):chefedepto=rs("chefedepto")
'if codcur<>10 and (turno<>"3") and left(rs("perlet"),4)>="2004" and isnull(diretor) then diretor="Maria Celia Soares Hungria de Luca"
'if codcur<>10 and (turno="3") and left(rs("perlet"),4)>="2004" and isnull(diretor) then diretor="Luiz Carlos de Azevedo Filho"
diretor=""

rs.movefirst
if rs("turno")="71" then turnod="Matutino"
if rs("turno")="73" then turnod="Noturno"
if rs("turno")="75" then turnod="Noturno"
classes=rs.fields.count-14:classes=6
linhas=rs.recordcount
icurso=rs("curso")
if tipo=1 then titulo="PONTO DO PESSOAL DOCENTE " else titulo="FOLHA DE PRESENÇA "
	
'response.write "Turno: " & turno & "<br>"
%>
<!-- borda -->
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho+10%>" height=1020>
<tr><td class=campo valign=top height=100%>
<!-- ponto -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho%>">
<tr>
	<td class="campop" align="left"><b><%=titulo%>DO CURSO DE <%=rs("curso")%></b></td>
	<td class="campop" align="right"><b><font size=5><%=pagina%></b></td>
</tr>
<tr>
	<td class="campop" align="left"><%=formatdatetime(dataponto,1)%></td>
	<td class="campop" align="right" nowrap>&nbsp; <!--<%=dialetivo%>º dia letivo--></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho%>">
<tr>
	<td class="campop" align="left"><b>Período</td>
	<td class="campop" align="center"><b>&nbsp;</td>
	<td class="campop" align="right"><b>&nbsp;</td>
</tr>

<table border="1" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho%>">
<tr>
	<td class=titulo align="center" rowspan=1>Nº</td>
	<td class=titulo align="center" rowspan=1>Nome / Disciplina</td>
	<td class=titulo align="center" rowspan=1>Assinatura</td>
	<td class=titulo align="center" colspan=<%=classes%>>Classes</td>
	<td class=fundo align="center" rowspan=1>Faltas</td>
	<td class=fundo align="center" rowspan=1>Observações</td>
</tr>
<%
rs.movefirst
do while not rs.eof 
if linhas<20 then
	altura=40
elseif linhas<23 then
	altura=35
else
	altura=30
end if

if lastperlet<>rs("perlet")&rs("turno") then
if rs("turno")="1" then turnod="Matutino"
if rs("turno")="2" then turnod="Vespertino"
if rs("turno")="3" then turnod="Noturno"
if rs("turno")="31" then turnod="Noturno"
sqla="select horini,horfim from grd_defhor where codds=" & diasemana & " and codtn in (" & rs("turno") & ") order by horini "
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
cols=rsc.recordcount
%>
<tr>
	<td class=campo colspan=3 height=<%=altura%>>Período Letivo: <%=rs("perlet")%> - <%=turnod%></td>
<%
rsc.movefirst
do while not rsc.eof
'for a= 1 to 6'14 to rs.fields.count-1
%>
	<td align="center" style="font-size:6pt; font-family:tahoma;font-weight:normal;background-color:Silver;color:Black;"><%=rsc("horini")%><br><%=rsc("horfim")%></td>
<%
'next
rsc.movenext
loop
	if cols<6 then
		for a=cols+1 to 6
		response.write "<td align="center" style='font-size:6pt; font-family:tahoma;font-weight:normal;background-color:Silver;color:Black;'></td>"
		next
	end if
rsc.close
%>
	<td class="campor" colspan=2 height=<%=altura%>>&nbsp;</td>
</tr>
<%
end if

linha=4+1
	
linha=linha+1
%>
<tr>
	<td class="campor" align="center" height=<%=altura%>><%=rs("chapa1")%></td>
	<td class="campor" ><b><%=rs("nome")%></b><br><%=rs("materia")%></td>
	<td class="campop" width=150>&nbsp;</td>
	<td class=campo align="center"><%if rs("a1")="1" then response.write rs("serie")&rs("turma") else response.write "&nbsp;"%></td>
	<td class=campo align="center"><%if rs("a2")="1" then response.write rs("serie")&rs("turma") else response.write "&nbsp;"%></td>
	<td class=campo align="center"><%if rs("a3")="1" then response.write rs("serie")&rs("turma") else response.write "&nbsp;"%></td>
	<td class=campo align="center"><%if rs("a4")="1" then response.write rs("serie")&rs("turma") else response.write "&nbsp;"%></td>
	<td class=campo align="center"><%if rs("a5")="1" then response.write rs("serie")&rs("turma") else response.write "&nbsp;"%></td>
	<td class=campo align="center"><%if rs("a6")="1" then response.write rs("serie")&rs("turma") else response.write "&nbsp;"%></td>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
</tr>
<%
linha=linha+1
lastperlet=rs("perlet")&rs("turno")
rs.movenext
loop

lbranco=""
if linhas<35-3 then
for b=linhas+1 to linhas+3 '35
%>
<tr>
	<td class="campor" height=30>&nbsp;</td>
	<td class="campor" nowrap><b>&nbsp;</b><br>&nbsp;</td>
	<td class="campop" width=150>&nbsp;</td>
<%for a=1 to 6 '14 to rs.fields.count-1 %>	
	<td class=campo align="center">&nbsp;</td>
<%next%>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
</tr>
<%
next
end if
%>
</table>
<!-- ponto -->

<!-- borda -->
</td></tr>
<tr><td class="campor" valign=top height=30>Diretor do curso: <%=diretor%><br>Coordenador do curso: <%=coordenador%>
</td></tr>
</table>
<%
icodcur=codcur
iturno=0 'turno
igrade=grade
idata=dataponto
idl=0 'dialetivo
ipg=pagina
if opcao="CLT" then igrade=1
if opcao="Convidados" then igrade=2
if session("usuariomaster")="023791" then 
	response.write turno & " turno <br>"
	response.write dataponto & " dataponto <br>"
	response.write dialetivo & " dialetivo <br>"
	response.write pagina & " pagina <br>"
	response.write codcur & " codcur <br>"
	response.write request.form
end if
sql="select coddoc from grades_dias where coddoc='" & icodcur & "' and turno=" & iturno & " and grade=" & igrade & " and data='" & dtaccess(idata) & "' "
rs2.Open sql, ,adOpenStatic, adLockReadOnly
existe=rs2.recordcount:rs2.close
if existe=0 then
	sql="INSERT INTO grades_dias (coddoc, temp, grade,turno, data, dialetivo, pagina) SELECT '" & icodcur & "','" & icurso & "'," & igrade & "," & iturno & ",'" & dtaccess(idata) & "'," & idl & "," & ipg & " "
	conexao.execute sql
else
	sql="UPDATE grades_dias SET pagina=" & ipg & ",dialetivo=" & idl & " ,grade=" & igrade & " WHERE data='" & dtaccess(idata) & "' AND turno=" & iturno & " AND coddoc='" & icodcur & "'"
	conexao.execute sql
end if ' existe

else 'sem registros
%>
<p>
<b><font color="#FF0000">
Esta seleção não mostra nenhum registro.</font></b></p>
<%
end if 'recordcount
%>
<%
rs.close
set rs2=nothing
end if ' finaliza=1

'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>