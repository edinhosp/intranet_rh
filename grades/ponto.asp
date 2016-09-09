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
<title>Grade Horária</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:40px">
<%
dim conexao, chapach, rs, rs1
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
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
	rs1.Open sqld, ,adOpenStatic, adLockReadOnly
	if rs1.recordcount>0 then
	rs1.movefirst:do while not rs1.eof 
		edia(rs1("dia1"))=1
	rs1.movenext:loop
	end if
	rs1.close

if finaliza=0 then
%>
<p class=titulo>Seleção para impressão do livro de ponto</p>
<form method="POST" action="ponto.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=450>
<tr><td class=titulo>Curso:</td>
<td class=titulo>Período:</td>
</tr>
<tr><td class=titulo><select size="1" name="codcur" onChange="javascript:submit()">
<%
sqla="SELECT c.coddoc, e.curso, tpcurso, descricao=case tpcurso when 'G' then 'Graduação' when 'L' then 'Cursos Livres' when 'M' then 'Mestrado' when 'P' then 'Pós-Graduação' else '' end " & _
"from g2cursos c, g2cursoeve e where tpcurso not in ('Z') and c.coddoc=e.coddoc and c.coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') group by tpcurso, e.curso, c.coddoc " & _
"order by tpcurso, e.curso "
rs1.Open sqla, ,adOpenStatic, adLockReadOnly
rs1.movefirst
do while not rs1.eof 
if rs1("tpcurso")<>grupoanterior then response.write "<option style='background:CCFFCC' value='" & rs1("tpcurso") & "'>------- " & ucase(rs1("descricao")) & " --------</option>"
%>
<option <%if codcur=rs1("coddoc") then response.write "selected "%> value="<%=rs1("coddoc")%>"><%=rs1("curso")%></option>
<%
grupoanterior=rs1("tpcurso")
rs1.movenext
loop
rs1.close
%>  
	<option value="POSN" <%if codcur="POSN" then response.write "selected"%>>Cursos de Pós - Narciso</option>
	<option value="POSYP" <%if codcur="POSYP" then response.write "selected"%>>Cursos de Pós - V.Yara - B.Prata</option>
	<option value="POSYV" <%if codcur="POSYV" then response.write "selected"%>>Cursos de Pós - V.Yara - B.Verde</option>

	</select></td>
<%
if codcur="" then codcur1="='0'":codcur="0"
if codcur="POSYP" then codcur1=    "in ('ART','PIN','PCL','MPS','PHO','PCI','DOP') "
if codcur="POSYV" then codcur1=    "in ('QTE','FDF') "
if codcur="POSN"  then codcur1="not in ('ART','PIN','PCL','MPS','FDF','MDI','DOP') "
%>
<td class=titulo><select size="1" name="opcao" onChange="javascript:submit()">
	<option value="0" selected>Selecione um período</option>
<%
if codcur="" then codcur=0
if turno="" then turno=9
sqla="select t.turno, h.descturno from g2aulas a, g2turmas t, corporerm.dbo.eturnos h " & _
"where t.coddoc='" & codcur & "' and a.id_grdturma=t.id_grdturma and h.codturno=t.turno " & _
"group by t.turno, h.descturno "
rs1.Open sqla, ,adOpenStatic, adLockReadOnly
if rs1.recordcount>0 then
rs1.movefirst
do while not rs1.eof 
if rs1.absoluteposition=1 and opcao="0" then opcao=rs1("turno")
%>
<option <%if cstr(opcao)=cstr(rs1("turno")) then response.write "selected"%> value="<%=rs1("turno")%>"><%=rs1("descturno")%></option>
<%
rs1.movenext
loop
end if 'recordcount >0
rs1.close
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
diasimpr=5
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
			if dateserial(anoagora,mesagora,ultimo)<=int(now+diasimpr) then
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
			if dateserial(anoagora,mesagora,ultimo)<=int(now+diasimpr) then
				response.write "<a href='ponto.asp?d=" & ultimo & "&m=" & mesagora & "&a=" & anoagora & "&c=" & codcur & "&o=" & opcao & "&or=" & ordem & "' class=r>"
			end if
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			if dateserial(anoagora,mesagora,ultimo)<=int(now+diasimpr) then
				response.write "</a>"
			end if
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
				if dateserial(anoagora,mesagora,ultimo)<=int(now+diasimpr) then
					response.write "<a href='ponto.asp?d=" & ultimo & "&m=" & mesagora & "&a=" & anoagora & "&c=" & codcur & "&o=" & opcao  & "&or=" & ordem & "' class=r>"
				end if
				if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				if dateserial(anoagora,mesagora,ultimo)<=int(now+diasimpr) then
					response.write "</a>"
				end if
			end if
		end if
		response.write "</td>"
	next
	response.write "</tr>"
next
dataponto=dateserial(anoagora,mesagora,diaagora)
sql2="select top 1 data, dialetivo, pagina from g2diaslivro " & _
"where coddoc='" & codcur & "' and turno=" & turno & " and impresso=1 " & _
"ORDER BY coddoc, turno, data DESC "
rs1.Open sql2, ,adOpenStatic, adLockReadOnly
if rs1.recordcount>0 then
	idata=rs1("data")
	idialetivo=rs1("dialetivo")
	ipagina=rs1("pagina")
else
	idata=""
	idialetivo=""
	ipagina=""
end if
rs1.close
sql2="select top 1 data, dialetivo, pagina from g2diaslivro " & _
"where coddoc='" & codcur & "' and turno=" & turno & " " & _
"and data='" & dtaccess(dataponto) & "' " & _
"ORDER BY coddoc, turno, data DESC "
rs1.Open sql2, ,adOpenStatic, adLockReadOnly
if rs1.recordcount>0 then
	ppagina=cdbl(rs1("pagina"))
	pdialetivo=cdbl(rs1("dialetivo"))
end if
rs1.close
if request.form("pg")<>"" then ppagina   =request.form("pg")
if request.form("dl")<>"" then pdialetivo=request.form("dl")
if request.form("pg")<>"" then ppagina   =ppagina
if request.form("dl")<>"" then pdialetivo=pdialetivo
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
	codcur=request.form("codcur"):coddoc=codcur
	turno=request.form("opcao")
	mes=request.form("mesform"):ano=request.form("anoform"):dia=request.form("diaform")
	dialetivo=request.form("dl")
	pagina=request.form("pg")
	periodoletivo=ano & "/"
	dataponto=dateserial(ano,mes,dia)	
	diasemana=weekday(dataponto)
	if pagina="" then pagina=2
	if dialetivo="" then dialetivo=1
	if pagina=1 then pagina=2

sql0="select tpcurso from g2cursos where coddoc='" & coddoc & "'"
rs1.Open sql0, ,adOpenStatic, adLockReadOnly
if rs1.recordcount>0 then tpcurso=rs1("tpcurso") else tpcurso=""
rs1.close
codcur1=""
if codcur="POSYP" then codcur1=    " in ('ART','PIN','PCL','PHO','PCI','DOP') ":tpcurso="P "
if codcur="POSYV" then codcur1=    " in ('QTE','FDF')":tpcurso="P "
if codcur="POSN"  then codcur1=" not in ('ART','PIN','PCL','MPS','FDF','MDI','DOP') ":tpcurso="P "

if tpcurso="G " OR tpcurso="G" then
	teste="g"
	sql1="select min(pos) minimo, max(pos) maximo from g2ch where coddoc='" & codcur & "' and '" & dtaccess(dataponto) & "' between inicio and termino and diasem=datepart (dw,'" & dtaccess(dataponto) & "') and turno=" & turno
	rs1.Open sql1, ,adOpenStatic, adLockReadOnly
	maximo=rs1("maximo"):minimo=rs1("minimo")
	rs1.close
	sql2="select a.chapa1, f.nome, a.coddoc, a.codtur, a.perlet, a.codmat, m. materia, a.diasem, a.turno, a.prof "
		sql3="select pos, g.codhor, horini, horfim, descricao from g2ch g where coddoc='" & codcur & "' and '" & dtaccess(dataponto) & "' between inicio and termino and diasem=datepart (dw,'" & dtaccess(dataponto) & "') and turno=" & turno & " group by pos, g.codhor, horini, horfim, descricao "
		rs1.Open sql3, ,adOpenStatic, adLockReadOnly
		do while not rs1.eof
			sql2=sql2 & ",'" & rs1("horini") & "'=max(case when a.pos=" & rs1("pos") & " then 'X' else '' end) "
		rs1.movenext:loop
		rs1.close
	sql2=sql2 & ", '' tipoprof "
	sql2=sql2 & "from g2ch a, g2defhor h, corporerm.dbo.umaterias m, grades_aux_prof f " & _
	"where h.codhor=a.codhor and m.codmat collate database_default=a.codmat " & _
	"and a.turno=" & turno & " and a.deletada=0 and a.coddoc='" & codcur & "' and '" & dtaccess(dataponto) & "' between a.inicio and a.termino " & _
	"and a.diasem=datepart (dw,'" & dtaccess(dataponto) & "') and f.chapa=a.chapa1 " & _
	"group by a.chapa1, f.nome, a.coddoc, a.codtur, a.perlet, a.codmat, m. materia, a.diasem, a.turno, a.prof " 
	if request.form("ordem")="turma" then sql2=sql2 & "ORDER BY a.perlet, a.codtur, f.nome "
	if request.form("ordem")="nome" then sql2=sql2 & "ORDER BY a.perlet, f.nome "
elseif codcur="POSN" or codcur="POSYV" or codcur="POSYP" then
	teste="3 ou 1"
	sql1="select a.coddoc, c.curso, a.chapa1, f.nome, /*d.id_data, d.id_grdaula,*/ d.data, d.aulas, codtur=min(a.codtur)+'<br>'+max(a.codtur), a.codmat, a.materia, a.perlet, " & _
	"tipoprof=case when chapa1<'10000' then 'Casa' else 'Convidado' end, datepart(w,data) diasem, turno, prof " & _
	"from ((g5datas d inner join g5ch a on a.id_grdaula=d.id_grdaula) " & _
	"left join grades_aux_prof f on f.chapa=a.chapa1) " & _
	"inner join g2cursoeve c on a.coddoc=c.coddoc " & _
	"where a.deletada=0 and d.deletada=0 and d.data='" & dtaccess(dataponto) & "' and c.coddoc " & codcur1 & "" & _
	"group by a.coddoc, c.curso, a.chapa1, f.nome, d.data, d.aulas, a.codmat, a.materia, a.perlet, " & _
	"case when chapa1<'10000' then 'Casa' else 'Convidado' end, datepart(w,data), turno, prof " & _
	"order by case when chapa1<'10000' then 'Casa' else 'Convidado' end, a.coddoc "
	sql2=sql1
else
	teste="resto"
	sql1="select a.coddoc, c.curso, a.chapa1, f.nome,  /*d.id_data, d.id_grdaula,*/ d.data, d.aulas, codtur=min(a.codtur)+'<br>'+max(a.codtur), a.codmat, a.materia, a.perlet, " & _
	"tipoprof=case when chapa1<'10000' then 'Casa' else 'Convidado' end, datepart(w,data) diasem, turno, prof " & _
	"from ((g5datas d inner join g5ch a on a.id_grdaula=d.id_grdaula) " & _
	"left join grades_aux_prof f on f.chapa=a.chapa1) " & _
	"inner join g2cursoeve c on a.coddoc=c.coddoc " & _
	"where a.deletada=0 and d.deletada=0 and d.data='" & dtaccess(dataponto) & "' and a.coddoc='" & coddoc & "' " & _
	"group by a.coddoc, c.curso, a.chapa1, f.nome, d.data, d.aulas, a.codmat, a.materia, a.perlet, " & _
	"case when chapa1<'10000' then 'Casa' else 'Convidado' end, datepart(w,data), turno, prof " & _
	"order by case when chapa1<'10000' then 'Casa' else 'Convidado' end, a.coddoc "
	sql2=sql1
end if
'response.write "<br>" & teste & "<br>" & sql2
rs.Open sql2, ,adOpenStatic, adLockReadOnly
tamanho=670
if rs.recordcount>0 then
rs.movefirst

if tpcurso="P " or tpcurso="P" and session("usuariomaster")="02379" then
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
'*************** fim teste **********************
end if

rs.movefirst
	sql4="select descturno from corporerm.dbo.eturnos where codturno=" & rs("turno")
	rs1.Open sql4, ,adOpenStatic, adLockReadOnly
	if rs1.recordcount>0 then turno=rs1("descturno") else turno="---"
	rs1.close
	sql5="select curso from g2cursoeve where coddoc='" & rs("coddoc") & "'"
	rs1.Open sql5, ,adOpenStatic, adLockReadOnly
	if rs1.recordcount>0 then icurso=rs1("curso") else icurso="---"
	rs1.close
classes=rs.fields.count-14:classes=(rs.fields.count-1)-10
if left(tpcurso,1)<>"G" then classes=3:icurso="CURSOS DE PÓS-GRADUAÇAO "
iturno=rs("turno")
linhas=rs.recordcount
%>
<!-- borda -->
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho+10%>" height=1020>
<tr><td class=campo valign=top height=100%>
<!-- ponto -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho%>">
<tr>
	<td class="campop" align="left"><b>PONTO DO PESSOAL DOCENTE DO CURSO DE <%=icurso%></b></td>
	<td class="campop" align="right"><b><font size=5><%=pagina%></b></td>
</tr>
<tr>
	<td class="campop" align="left"><%=formatdatetime(dataponto,1)%></td>
	<td class="campop" align="right" nowrap><%=dialetivo%>º dia letivo</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho%>">
<tr>
	<td class="campop" align="left"><b>Período <%=turno%></td>
	<td class="campop" align="center"><b>&nbsp;</td>
	<td class="campop" align="right"><b>&nbsp;</td>
</tr>

<table border="1" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho%>">
<tr>
	<td class=titulo align="center" rowspan=1>Nº</td>
	<td class=titulo align="center" rowspan=1>Nome / Disciplina</td>
	<td class=titulo align="center" rowspan=1>Turma</td>
	<td class=titulo align="center" rowspan=1>Assinatura</td>
	<td class=titulo align="center" colspan=<%=classes%>>Classes</td>
	<td class=fundo align="center" rowspan=1>Observações</td>
</tr>
<%
linha=4+1
rs.movefirst
do while not rs.eof 
altura=35 'para padronizar

'-----------------------
'Cabeçalho pagina
if linha>29 or (lasttipo<>rs("tipoprof") and rs.absoluteposition>1) then
if left(tpcurso,1)<>"G" then
for b=1 to 5 '35
%>
<tr>
	<td class="campor" height=35>&nbsp;</td>
	<td class="campor" nowrap><b>&nbsp;</b><br>&nbsp;</td>
	<td class="campop" width=20>&nbsp;</td>
	<td class="campop" width=150>&nbsp;</td>
<%
for a=1 to 3
%>	
	<td class=campo align="center">&nbsp;</td>
<%
next
%>
	<td class=campo>&nbsp;</td>
</tr>
<%
linha=linha+1
next
end if
'end if

response.write "</table>"
'-- ponto --
'-- borda --
response.write "</td></tr>"
response.write "</table>"

response.write "<DIV style=""page-break-after:always""></DIV>"
'-- borda --
response.write "<table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='" & tamanho+10 & "' height=1020>"
response.write "<tr><td class=campo valign=top height='100%'>"
'-- ponto --
response.write "<table border='0' cellpadding='3' cellspacing='0' style='border-collapse: collapse' width='" & tamanho & "'>"
response.write "<tr>"
response.write "	<td class=""campop"" align=""left""><b>PONTO DO PESSOAL DOCENTE DO CURSO DE " & icurso & "</b></td>"
response.write "	<td class=""campop"" align=""right""><b><font size=5>" & pagina & "V</b></td>"
response.write "</tr>"
response.write "<tr>"
response.write "	<td class=""campop"" align=""left"">" & formatdatetime(dataponto,1) & "</td>"
response.write "	<td class=""campop"" align=""right"" nowrap>" & dialetivo & "º dia letivo</td>"
response.write "</tr>"
response.write "</table>"

response.write "<table border='0' bordercolor='#000000' cellpadding='3' cellspacing='0' style='border-collapse: collapse' width='" & tamanho & "'>"
response.write "<tr>"
response.write "	<td class=""campop"" align=""left""><b>Período " & turno & "</td>"
response.write "	<td class=""campop"" align=""center""><b>&nbsp;</td>"
response.write "	<td class=""campop"" align=""right""><b>&nbsp;</td>"
response.write "</tr>"

response.write "<table border='1' bordercolor='#000000' cellpadding='1' cellspacing='0' style='border-collapse: collapse' width='" & tamanho & "'>"
response.write "<tr>"
response.write "	<td class=""titulo"" align=""center"" rowspan=1>Nº</td>"
response.write "	<td class=""titulo"" align=""center"" rowspan=1>Nome / Disciplina</td>"
response.write "	<td class=""titulo"" align=""center"" rowspan=1>Turma</td>"
response.write "	<td class=""titulo"" align=""center"" rowspan=1>Assinatura</td>"
response.write "	<td class=""titulo"" align=""center"" colspan=" & classes & ">Classes</td>"
response.write "	<td class=""fundo"" align=""center"" rowspan=1>Observações</td>"
response.write "</tr>"
linha=4+1

end if
'-----------------------
if lastperlet<>rs("perlet")&rs("turno") or linha=5 then
%>
<tr>
	<td class=campo colspan=4 height=<%=altura%>>Período Letivo: <%=rs("perlet")%> - <%=turno%></td>
<%
if left(tpcurso,1)="G" then
	for a=10 to rs.fields.count-2
%>
	<td align="center" style="font-size:6pt; font-family:tahoma;font-weight:normal;background-color:Silver;color:Black;"><%=rs.fields(a).name%></td>
<%
	next
else
%>
	<td align="center" style="font-size:6pt; font-family:tahoma;font-weight:normal;background-color:Silver;color:Black;">#Aulas</td>
	<td align="center" width=70 style="font-size:6pt; font-family:tahoma;font-weight:normal;background-color:Silver;color:Black;">Inicio</td>
	<td align="center" width=70 style="font-size:6pt; font-family:tahoma;font-weight:normal;background-color:Silver;color:Black;">Termino</td>
<%
end if
%>
	<td class="campor" colspan=1 height=<%=altura%>>&nbsp;<%%></td>
</tr>
<%
linha=linha+1
end if
%>
<tr>
	<td class="campor" align="center" height=<%=altura%>><%=rs("chapa1")%></td>
	<td class="campor"><b><%=rs("nome")%></b>
	<p style="font-size:7pt;margin-top:0;margin-bottom:0"><%=rs("materia")%>
<%if left(tpcurso,1)<>"G" then%>
	<p style="font-size:7pt;margin-top:0;margin-bottom:0"><%=rs("curso")%>
<%end if%>
	</td>
	<td class="campor" width=20><%=rs("codtur")%></td>
	<td class="campop" width=150>&nbsp;</td>
<%
if left(tpcurso,1)="G" then
	for b=10 to rs.fields.count-2
%>	
	<td class="campop" align="center" colspan=1><%=rs.fields(b)%></td>
<%
	next
else
%>	
	<td class="campop" align="center" colspan=1><%=rs("aulas")%></td>
	<td class="campop" align="center" colspan=1>&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td class="campop" align="center" colspan=1>&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;</td>
<%
end if
%>
	<td class=campo>&nbsp;</td>
</tr>
<%
'***********************
linha=linha+1
'***********************
lastperlet=rs("perlet")&rs("turno")
lasttipo=rs("tipoprof")
rs.movenext
loop

lbranco=""
'**** IMPRESSAO PROFESSORES MESTRADO *****
if coddoc="MPS" or coddoc="MDI" then 
sqlm="select p.chapa, f.nome from g2pontopos p " & _
"inner join grades_aux_prof f on f.chapa=p.chapa collate database_default " & _
"where p.CodDoc='" & coddoc & "' " & _
"and p.chapa collate database_default not in ( " & _
"select distinct a.chapa1 from g5datas d inner join g5ch a on a.id_grdaula=d.id_grdaula where a.deletada=0 and d.deletada=0 and d.data='" & dtaccess(dataponto) & "' and a.coddoc='" & coddoc & "' " & _
") order by f.nome "
rs1.Open sqlm, ,adOpenStatic, adLockReadOnly
if rs1.recordcount>0 then
do while not rs1.eof
%>
<tr>
	<td class="campor" align="center" height=<%=altura%>><%=rs1("chapa")%></td>
	<td class="campor"><b><%=rs1("nome")%></b></td>
<%for a=1 to 6%>
	<td class=campo align="center">&nbsp;</td>
<%next%>
</tr>
<%
rs1.movenext
loop
end if
rs1.close
end if '**** fim impressao mestrado *****
'if linhas<35-3 then
for b=1 to 5 '35
%>
<tr>
	<td class="campor" height=35>&nbsp;</td>
	<td class="campor" nowrap><b>&nbsp;</b><br>&nbsp;</td>
	<td class="campop" width=20>&nbsp;</td>
	<td class="campop" width=150>&nbsp;</td>
<%
if left(tpcurso,1)="G" then
for a=10 to rs.fields.count-2
%>	
	<td class=campo align="center">&nbsp;</td>
<%
next
else
for a=1 to 3
%>	
	<td class=campo align="center">&nbsp;</td>
<%
next
end if
%>
	<td class=campo>&nbsp;</td>
</tr>
<%
linha=linha+1
next
'end if
if coddoc="MPS" or coddoc="MDI" then 
	response.write "<tr><td class=""campop"" colspan=8><b>Além da assinatura, anotar o horário trabalhado (aulas, orientações etc)</td></tr>"
end if
%>
</table>
<br>
<br>___________________________________
<br>Coordenador do curso

<!-- ponto -->

<!-- borda -->
</td></tr>
</table>
<%
icodcur=codcur
iturno=iturno
igrade=grade
idata=dataponto
idl=dialetivo
ipg=pagina
sql="select coddoc from g2diaslivro where coddoc='" & icodcur & "' and turno=" & iturno & " and data='" & dtaccess(idata) & "' "
rs1.Open sql, ,adOpenStatic, adLockReadOnly
existe=rs1.recordcount:rs1.close
if existe=0 then
	sql="INSERT INTO g2diaslivro (coddoc, turno, data, dialetivo, pagina, obs) SELECT '" & icodcur & "'," & iturno & ",'" & dtaccess(idata) & "'," & idl & "," & ipg & ",'' "
	conexao.execute sql
else
	sql="UPDATE g2diaslivro SET impresso=1, pagina=" & ipg & ",dialetivo=" & idl & " WHERE data='" & dtaccess(idata) & "' AND turno=" & iturno & " AND coddoc='" & icodcur & "'"
	sql="UPDATE g2diaslivro SET impresso=1 WHERE data='" & dtaccess(idata) & "' AND turno=" & iturno & " AND coddoc='" & icodcur & "'"
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
set rs1=nothing
end if ' finaliza=1

'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>