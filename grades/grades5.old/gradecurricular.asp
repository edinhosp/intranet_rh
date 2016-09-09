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
<title>Grade Curricular</title>
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
dim conexao, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
set rs4=server.createobject ("ADODB.Recordset")
Set rs4.ActiveConnection = conexao
'rs.Open sql1, ,adOpenStatic, adLockReadOnly
if request.form<>"" then
	if request.form("B3")<>"" then
		finaliza=1
	else
		finaliza=0
	end if
end if
	
if finaliza=0 then
'if request.form("dataemmissao")<>"" then dataemissao=request.form("dataemissao") else dataemissao=dateserial(year(now),month(now)+1,1)
%>
<p class=titulo>Seleção para impressão de Grade curricular</p>
<form method="POST" action="gradecurricular.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="400">
<tr><td class=titulo colspan=2>Curso Graduação</td></tr>
<tr><td class=titulo colspan=2><select size="1" name="codcur">
	<option value="0" selected>Selecione um curso</option>
<%
sqla="SELECT coddoc, curso from grades_5 where coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') GROUP BY coddoc, curso order by curso "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
%>
	<option <%if request.form("coddoc")=rs("coddoc") then response.write "selected "%> value="<%=rs("coddoc")%>"><%=rs("curso")%> (<%=rs("coddoc")%>)</option>
<%
rs.movenext:loop
rs.close
%>  
	</select></td></tr>
<tr>
	<td class=titulo>Grade Curricular vigente em: <select size="1" name="perlet">
<%
sqla="select perlet2 from grades_5 group by perlet2 order by perlet2 desc"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
%>
	<option <%if request.form("perlet")=rs("perlet2") then response.write "selected"%> value="<%=rs("perlet2")%>"><%=rs("perlet2")%></option>
<%
rs.movenext:loop
rs.close
%>  
	</select>
	
	</td>
</tr>
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
	codcur=right(request.form("codcur"),len(request.form("codcur"))-3)
	codcur=request.form("codcur")
	'enfase=left(request.form("codcur"),3)
	datae=request.form("perlet")
	periodos=0
	
	sql1="SELECT per.coddoc, per.codcur, per.curso, per.perlet, per.perlet2, per.perletsg, per.pini, per.pfim, gc.GC, gc.serie, mat.CODMAT, mat.MATERIA, " &_
	"mat.NAULASSEM, mat.CARGAHORARIA, per.enfase " & _
	"FROM grades_per as per, grades_gc as gc, grades_materias as mat " & _
	"WHERE (per.perlet=gc.perlet) AND (per.coddoc=gc.coddoc) AND (per.gc=gc.gc) " & _
	"AND (gc.serie=mat.serie) AND (gc.GC=mat.GC) AND (gc.coddoc=mat.coddoc) " & _
	"AND perlet2='" & datae & "' AND per.coddoc='" & codcur & "' " & _
	"group by per.coddoc, per.curso, per.codcur, per.perlet, per.perlet2, per.perletsg, per.pini, per.pfim, gc.GC, gc.serie, mat.CODMAT, mat.MATERIA, mat.NAULASSEM, mat.CARGAHORARIA, per.enfase " & _	
	"ORDER BY per.coddoc, per.curso, per.perlet, per.pini, gc.GC, gc.serie, mat.CODMAT, mat.MATERIA "
	'response.write "<br>" & sql1
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	'response.write "<br>" & rs.recordcount
	if rs.recordcount>0 then
tamanho=900
%>
<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=<%=tamanho%>>
<tr>
	<td class=titulo colspan=3>CURSO DE <%=rs("curso")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<font size=1><%=rs("codcur")%></font></td>
</tr>
<tr>
	<td class=titulo align="center">Período Letivo</td>
	<td class=titulo align="center">Duração</td>
	<td class=titulo align="center">Grade Curricular</td>
</tr>
<tr>
	<td class="campol" align="center"><%=rs("perlet")%></td>
	<td class="campol" align="center"><%=rs("pini") & " a " & rs("pfim")%></td>
	<td class="campol" align="center"><%=rs("GC")%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=<%=tamanho%>>
<tr>
	<td class=titulor align="center" width=25>Per.</td>
	<td class=titulor align="center" width=<%=tamanho-25-35-30-530%> >Disciplina</td>
	<td class=titulor align="center" width=35>Aulas</td>
	<td class=titulor align="center" width=30>C.H.</td>
	<td class=titulor align="center" width=530>&nbsp;</td>
</tr>
<%
lastperlet=rs("perlet")
lastperlet2=rs("perlet2")
lastgc=rs("gc")
rs.movefirst
do while not rs.eof

if lastperlet<>rs("perlet") or lastgc<>rs("gc") then
periodos=periodos+1
%>
<!-- -->
<!--- checagem de disciplinas fora do padrão -->
<%
	perlet=lastperlet
	perle2=lastperlet2
	sql4="SELECT g.perlet, g.coddoc, g.codtur, g.turno, g.serie, g.turma, g.diasem, g.descricao, g.codmat, g.materia, g.chapa1, p.nome, g.usuarioa, g.usuarioc, g.id_grade " & _
	"FROM grades_5ch AS g, grades_aux_prof AS p  " & _
	"WHERE g.chapa1=p.chapa and g.perlet='" & perlet & "' and g.perlet2='" & perlet2 & "' AND g.coddoc='" & codcur & "' " & _
	"and g.codmat not in (SELECT gm.CODMAT " & _
	"FROM grades_gc AS gc INNER JOIN grades_materias AS gm ON (gc.serie = gm.periodo) AND (gc.GC = gm.GC) AND (gc.codcur = gm.CODCUR) " & _
	"WHERE gc.coddoc='" & codcur & "' AND gc.perlet='" & perlet & "' and gc.serie=g.serie)"
	rs4.Open sql4, ,adOpenStatic, adLockReadOnly
'	RESPONSE.WRITE SQL4
	if rs4.recordcount>0 then
%>
<tr>
	<td class=campo colspan=5>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=100%>
<tr>
	<td class=titulo colspan=7>Checagem de Disciplinas fora da Grade Curricular</td>
</tr>
<tr>
	<td class=campo>Turno</td>
	<td class=campo>Turma</td>
	<td class=campo>Dia</td>
	<td class=campo>Inicio</td>
	<td class=campo>Disciplina</td>
	<td class=campo>Professor</td>
	<td class=campo></td>
</tr>
<%
rs4.movefirst
do while not rs4.eof
	if isnull(rs4("usuarioa")) then usuario=rs4("usuarioc") else usuario=rs4("usuarioa")
	sql="select nome from usuarios where usuario='" & usuario & "'"
	rs2.Open sql, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then usuarion=rs2("nome") else usuarion=""
	rs2.close:turno1=rs4("turno")
	if rs4("turno")="1" then turno="Mat"
	if rs4("turno")="2" then turno="Vesp"
	if rs4("turno")="3" then turno="Not"
	if rs4("turno")="5" then turno="Vesp-EF"
	if rs4("turno")="61" then turno="Integral"
	if rs4("turno")="62" then turno="Integral"
%>
<tr>
	<td class=campo><%=turno%></td>
	<td class=campo><%=rs4("serie")&rs4("turma")%></td>
	<td class=campo><%=weekdayname(rs4("diasem"),1)%></td>
	<td class=campo><%=rs4("descricao")%></td>
	<td class=campo><%=rs4("codmat") & " - " & rs4("materia")%></td>
	<td class=campo><%=rs4("chapa1") & " - " & rs4("nome")%></td>
	<td class=campo>
    <% if session("a81")="T" then %>
	<a href="grade_alteracao.asp?codigo=<%=rs4("id_grade")%>" onclick="NewWindow(this.href,'AlteracaoGrade','550','330','no','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a><%=usuarion%>
	<% end if %>
	</td>
</tr>
<%
rs4.movenext
loop
%>
</table>
	</td>
<tr>
<%
end if	' rs.recordcount
rs4.close
%>	
<!--- checagem de disciplinas fora do padrão -->


</table>
<br>
<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=<%=tamanho%>>
<tr>
	<td class=titulo align="center">Período Letivo</td>
	<td class=titulo align="center">Duração</td>
	<td class=titulo align="center">Grade Curricular</td>
</tr>
<tr>
	<td class="campol" align="center"><%=rs("perlet")%></td>
	<td class="campol" align="center"><%=rs("pini") & " a " & rs("pfim")%></td>
	<td class="campol" align="center"><%=rs("GC")%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=<%=tamanho%>>
<tr>
	<td class=titulor align="center" width=25>Per.</td>
	<td class=titulor align="center" width=<%=tamanho-25-35-30-530%> >Disciplina</td>
	<td class=titulor align="center" width=35>Aulas</td>
	<td class=titulor align="center" width=30>C.H.</td>
	<td class=titulor align="center" width=530>&nbsp;</td>
</tr>
<%
end if ' mudanca de periodo letivo ou grade
%>

<tr>
	<td class="campor" align="center"><%=rs("serie")%></td>
	<td class="campor" ><%=rs("materia")%><br><%=rs("codmat")%></td>
	<td class="campor" align="center"><%=rs("naulassem")%></td>
	<td class="campor" align="center"><%=rs("cargahoraria")%></td>
	<td class="campor" valign=top align="left" style="border: 0 "> 
<%
	sql2="SELECT serie, turma, codtur, turno=round((case when turno<10 then turno*10 else turno end)/10,0), sum(g.ta) AS aulas FROM grades_5chi AS g " & _
	"WHERE g.perlet='" & rs("perlet") & "' AND g.coddoc='" & rs("coddoc") & "' " & _
	"AND g.serie=" & rs("serie") & " AND g.codmat='" & rs("codmat") & "' and deletada=0 and ativo=1 " & _
	"AND perlet2='" & datae & "' and g.enfase='" & rs("enfase") & "' " & _
	"GROUP BY codtur, coddoc, round((case when turno<10 then turno*10 else turno end)/10,0), turma, serie "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then
%>
	<table border="1" bordercolor="#CCCCCC" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=530>
	<tr>
		<td class="campoa"r width=35 style="border-top: 2px solid #000000">Turma</td>
		<td class="campoa"r width=15 style="border-top: 2px solid #000000">T.</td>
		<td class="campoa"r width=350>Professores</td>
		<td class="campoa"r width=100>Tp.Alunos</td>
		<td class="campoa"r width=30>Qtde</td>
	</tr>
<%
	rs2.movefirst
	do while not rs2.eof
	if rs2("aulas")<>rs("naulassem") then dif="<font color=red>!</font>" else dif="&nbsp;"
	if rs2("aulas")<>rs("naulassem") then cldif="campolr" else cldif="campoar"
	
	if rs2("turno")="1" then turno="Mat"
	if rs2("turno")="2" then turno="Vesp"
	if rs2("turno")="3" then turno="Not"
	if rs2("turno")="5" then turno="Vesp-EF"
	if rs2("turno")="61" then turno="Integral"
	if rs2("turno")="62" then turno="Integral"
	if rs2("turno")="6" then turno="Integral"
	'turma=right(rs2("sturma"),1)
	turma=rs2("turma")
	perlet=rs("perletsg")
'	"HAVING UMATALUN.STATUS In ('01','15','20') "

	sql3="SELECT uma.STATUS, usm1.DESCRICAO, Count(umc.MATALUNO) AS alunos " & _
	"FROM corporerm.dbo.USITMAT AS usm1 INNER JOIN ((((corporerm.dbo.UMATRICPL umc INNER JOIN corporerm.dbo.UMATALUN uma ON (umc.GRADE=uma.GRADE) AND (umc.CODPER=uma.CODPER) " & _
	"AND (umc.CODCUR=uma.CODCUR) AND (umc.PERLETIVO=uma.PERLETIVO) AND (umc.MATALUNO=uma.MATALUNO) AND (umc.CODCOLIGADA=uma.CODCOLIGADA) AND (umc.CODFILIAL=uma.CODFILIAL)) " & _
	"INNER JOIN corporerm.dbo.UMATERIAS um ON (uma.CODCOLIGADA=um.CODCOLIGADA) AND (uma.CODMAT=um.CODMAT)) " & _
	"INNER JOIN corporerm.dbo.UALUCURSO uac ON (umc.GRADE=uac.GRADE) AND (umc.CODPER=uac.CODPER) AND (umc.CODCUR=uac.CODCUR) AND (umc.MATALUNO=uac.MATALUNO) AND (umc.CODCOLIGADA=uac.CODCOLIGADA)) " & _
	"INNER JOIN corporerm.dbo.USITMAT usm ON uac.STATUS=usm.CODSITMAT) ON usm1.CODSITMAT=uma.STATUS " & _
	"WHERE umc.PERLETIVO='" & perlet & "' " & _
	"AND umc.CODCUR=" & rs("codcur") & " " & _
	"AND CODTUR='" & rs2("codtur") & "' " & _
	"AND uma.CODMAT='" & rs("codmat") & "' " & _
	"GROUP BY uma.STATUS, usm1.DESCRICAO "
	rs3.Open sql3, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then linhas=rs3.recordcount else linhas=1
%>
	<tr>
<!--		<td width=35 class="campoa"r rowspan="<%=linhas%>" align="center"><%=rs2("serie") & rs2("turma")%><br>(<%=turno%>)</td> -->
		<td width=35 class="campoa"r rowspan="<%=linhas%>" align="center"><%=rs2("codtur")%></td>
		<td width=15 class=<%=cldif%> rowspan="<%=linhas%>" align="center"><%=rs2("aulas")%></td>
		<td width=350 class="campoa"r rowspan="<%=linhas%>" valign=top>
<%
	sql4="SELECT g.diasem, g.descricao, g.chapa1, p.NOME, g.prof, juntar, jturma, dividir " & _
	"FROM grades_5ch AS g, grades_aux_prof AS p " & _
	"WHERE g.chapa1 = p.CHAPA and g.perlet='" & rs("perlet") & "' AND g.coddoc='" & codcur & "' " & _
	"AND g.serie=" & rs("serie") & " AND g.turma='" & turma & "' AND g.codmat='" & rs("codmat") & "' " & _
	"and g.codtur='" & rs2("codtur") & "' " & _
	"AND perlet2='" & datae & "' and g.deletada=0 and g.ativo=1 " & _
	"ORDER BY g.prof, g.diasem, g.descricao "
	rs4.Open sql4, ,adOpenStatic, adLockReadOnly
	if rs4.recordcount>0 then
%>		
	<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=350>
<%
	rs4.movefirst
	do while not rs4.eof
	if rs4("juntar")=-1 then juntar=" " & rs4("chapa1") &" Junta turma" else juntar="&nbsp;"
	if rs4("dividir")=-1 then dividir=" Divide turma" else dividir="&nbsp;"
	if rs4("prof")="2" then prof="(2) " else prof=""
%>
	<tr>
		<td class="campoa"r width=270 ><%if rs4("chapa1")="99999" then response.write "<font color=red>"%><%=rs4("nome") & "</font>" & prof%><%=juntar & " " & dividir%></td>
		<td class="campoa"r width=20 style="border-left: 1px solid #000000"><%=weekdayname(rs4("diasem"),1)%></td>
		<td class="campoa"r width=60 style="border-left: 1px solid #000000" nowrap><%=rs4("descricao")%></td>
	</tr>
<%
	rs4.movenext
	loop
%>
	</table>
<%
	end if
	rs4.close
%>		
		</td>
<%
	if rs3.recordcount>0 then
	rs3.movefirst
	do while not rs3.eof
	if rs3.absoluteposition>1 then response.write "<tr>"
%>
		<td class="campoa"r width=100><%=left(rs3("descricao"),20)%></td>
		<td class="campoa"r width=30 align="center"><%=rs3("alunos")%></td>
	</tr>
<%
	rs3.movenext
	loop
	else
%>
		<td class="campoa"r width=100></td>
		<td class="campoa"r width=30></td>
	</tr>
<%
	end if 'rs3.recordcount>0
	rs3.close
%>
<%
	rs2.movenext
	loop
%>
	</table>
<%
	end if ' rs2.recordcount>0
	rs2.close
%>
	
	</td>
</tr>
<%
lastperlet=rs("perlet")
lastgc=rs("gc")
rs.movenext
loop
rs.close
%>

<!--- checagem de disciplinas fora do padrão -->
<% if periodos>0 then %>
<%
	perlet=lastperlet
	perle2=lastperlet2
	sql4="SELECT g.perlet, g.codcur, g.turno, g.serie, g.turma, g.diasem, g.descricao, g.codmat, g.materia, g.chapa1, p.nome, g.usuarioa, g.usuarioc, g.id_grade " & _
	"FROM g2ch AS g, (select chapa, nome from pfunc union all select chapa, nome from grades_novos) AS p  " & _
	"WHERE g.chapa1=p.chapa and g.perlet='" & perlet & "' and g.perlet2='" & perlet2 & "' AND g.codcur=" & codcur & " " & _
	"and g.codmat not in (SELECT gm.CODMAT " & _
	"FROM grades_gc AS gc INNER JOIN grades_materias AS gm ON (gc.serie = gm.periodo) AND (gc.GC = gm.GC) AND (gc.codcur = gm.CODCUR) " & _
	"WHERE gc.codcur=" & codcur & " AND gc.perlet='" & perlet & "' and gc.serie=g.serie)"
	rs4.Open sql4, ,adOpenStatic, adLockReadOnly
'	RESPONSE.WRITE SQL4
	if rs4.recordcount>0 then
%>
<tr>
	<td class=campo colspan=5>
<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=100%>
<tr>
	<td class=titulo colspan=7>Checagem de Disciplinas fora da Grade Curricular</td>
</tr>
<tr>
	<td class=campo>Turno</td>
	<td class=campo>Turma</td>
	<td class=campo>Dia</td>
	<td class=campo>Inicio</td>
	<td class=campo>Disciplina</td>
	<td class=campo>Professor</td>
	<td class=campo></td>
</tr>
<%
rs4.movefirst
do while not rs4.eof
	if isnull(rs4("usuarioa")) then usuario=rs4("usuarioc") else usuario=rs4("usuarioa")
	sql="select nome from usuarios where usuario='" & usuario & "'"
	rs2.Open sql, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then usuarion=rs2("nome") else usuarion=""
	rs2.close

	if rs4("turno")="1" then turno="Mat"
	if rs4("turno")="2" then turno="Vesp"
	if rs4("turno")="3" then turno="Not"
	if rs4("turno")="5" then turno="Vesp-EF"
	if rs4("turno")="61" then turno="Integral"
	if rs4("turno")="62" then turno="Integral"

%>
<tr>
	<td class=campo><%=turno%></td>
	<td class=campo><%=rs4("serie")&rs4("turma")%></td>
	<td class=campo><%=weekdayname(rs4("diasem"),1)%></td>
	<td class=campo><%=rs4("descricao")%></td>
	<td class=campo><%=rs4("codmat") & " - " & rs4("materia")%></td>
	<td class=campo><%=rs4("chapa1") & " - " & rs4("nome")%></td>
	<td class=campo>
    <% if session("a81")="T" then %>
      <a href="grade_alteracao.asp?codigo=<%=rs4("id_grade")%>" onclick="NewWindow(this.href,'AlteracaoGrade','550','330','no','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a><%=usuarion%>
	<% end if %>
	</td>

</tr>
<%
rs4.movenext
loop
%>
</table>
<%
end if	
rs4.close
%>	
	</td>
<tr>
<!-- -->
<% end if 'periodos %>

</table>	
<%	
	else 'rs.recordcount>0
		response.write "<p class=realce>Esta seleção não mostra nenhum registro."
	end if 'rs.recordcount>0	

termino=now()
duracao=(termino-inicio)
Response.write "<p class=realce><font size=1> Inicio: " & inicio & " Termino: " & termino & " Duracao: " & formatdatetime(duracao,3) & "</font></p>"
	
end if 'finaliza 1

set rs=nothing
set rs2=nothing
set rs3=nothing
set rs4=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>