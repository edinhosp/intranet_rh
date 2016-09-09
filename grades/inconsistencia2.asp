<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a80")="N" or session("a80")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<?xml version="1.0" encoding="iso-8859-1"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Checagem de Inconsistências na Grade Horária</title>
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
conexao.Open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
sessao=session.sessionid
inicio=now()
fast=0 : fast2=0 : fast3=0 : fast4=0 : fast5=0
periodo="2013%2"
periodo2="('2013/2','2013/0')"
fimper=dateserial(2013,07,31)
inicioper=dateserial(2013,08,1)

sql="update g2aulas set jpai=0, juntar=0, juntar_id=null, jturma=null " & _
"where inicio>='" & dtaccess(inicioper) & "' and chapa1 like '99%'"
conexao.execute sql

if int(now())<inicioper then datacheck=inicioper else datacheck=int(now())

'if request("status")="" then status="A" else status=request("status")
'if session("ultimainconsistencia")<>"" then status=session("ultimainconsistencia") else status=request("status")
if request("status")="" and session("ultimainconsistencia")="" then session("ultimainconsistencia")="A"
if request("status")="" then status=session("ultimainconsistencia") else status=request("status")
session("ultimainconsistencia")=status

aba0="border-top:2pt double #000000;border-left:2pt double #000000;border-right:3pt double #000000;font-weight:normal;"
aba1="border-top:2pt solid #000000;border-left:2pt solid #000000;border-right:3pt solid #000000;font-weight:bold;"
if status="A"  then abaa=aba1  else abaa=aba0
if status="A2" then abaa2=aba1 else abaa2=aba0
if status="D"  then abad=aba1  else abad=aba0
if status="C"  then abac=aba1  else abac=aba0
if status="J"  then abaj=aba1  else abaj=aba0
if status="J2" then abaj2=aba1 else abaj2=aba0
if session("inconsis_hoje")="" then
	sql="delete from g2aulas where deletada=1":conexao.execute sql
	session("inconsis_hoje")="ok"
end if
%>
<form name="form">
<table border="0" bordercolor="#CCCCCC" cellpadding="2" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class="campop"><b>Checagem de Inconsistências na Grade Horária</b></td>
	<td class=campo width="5">&nbsp;</td>

	<td class=campo width="70" align="center" style="<%=abaa%>">
	<a href="inconsistencia2.asp?status=A">	Ativos</a></td>
	<td class=campo width="5">&nbsp;</td>

	<td class=campo width="70" align="center" style="<%=abaa2%>">
	<a href="inconsistencia2.asp?status=A2">	Acima/Abaixo<br>Limite</a></td>
	<td class=campo width="5">&nbsp;</td>

	<td class=campo width="70" align="center" style="<%=abad%>">
	<a href="inconsistencia2.asp?status=D">	Afastados<br>Demitidos</a></td>
	<td class=campo width="5">&nbsp;</td>
	<td class=campo width="70" align="center" style="<%=abac%>">
	<a href="inconsistencia2.asp?status=C">	Sem aula /<br>Contratação</a></td>
	<td class=campo width="5">&nbsp;</td>

	<td class=campo width="70" align="center" style="<%=abaj%>">
	<a href="inconsistencia2.asp?status=J">	Problemas Junções</a></td>
	<td class=campo width="5">&nbsp;</td>
	<td class=campo width="70" align="center" style="<%=abaj2%>">
	<a href="inconsistencia2.asp?status=J2">	Junções </a></td>
	<td class=campo width="5">&nbsp;</td>
</tr>
</table>

<% 
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
if status="A" then 
%>
<!--
períodos de inicio e termino
-->
<%
if fast=0 then
sql3="select g.id_grdaula, coddoc, codtur, codmat, materia, diasem, descricao, g.inicio, g.termino, p.inicio inicioc, p.termino terminoc, codhor, id_grdturma, usuarioa, usuarioc " & _
"from g2ch g, g2periodoaula p where g.perlet=p.perlet and deletada=0 and g.perlet in " & periodo2 & " " & _
"and (g.inicio<p.inicio or g.termino>p.termino) and '" & dtaccess(datacheck) & "' between g.inicio and g.termino "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Datas de inicio ou término das aulas incorretos (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center">Curso</td>
	<td class=titulor align="center">Turma</td>
	<td class=titulor align="center">Dia</td>
	<td class=titulor align="center">Aula</td>
	<td class=titulor align="center">Disciplina/Professor</td>
	<td class=titulor align="center">Inicio/Termino</td>
	<td class=titulor align="center">&nbsp;    </td>
</tr>
<%
rs.movefirst:do while not rs.eof
msg=""
if isnull(rs("usuarioa")) then usuario=rs("usuarioc") else usuario=rs("usuarioa")
sql="select nome from usuarios where usuario='" & usuario & "'"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then usuarion=rs2("nome") else usuarion=""
rs2.close
%>
<tr>
	<td class="campor"><%=rs("coddoc")%></td>
	<td class="campor"><%=rs("codtur")%></td>
    <td class="campor"><%=weekdayname(rs("diasem"),1)%></td>
	<td class="campor"><%=rs("descricao")%></td>
    <td class="campor"><%=rs("materia")%></td>
    <td class="campor"><%=rs("inicio")&" - "&rs("termino")%></td>
	<td class="campor">
	<% if session("a80")="T" then %>
	<a href="gradenovaaula.asp?idaula=<%=rs("id_grdaula")%>&idturma=<%=rs("id_grdturma")%>&codhor=<%=rs("codhor")%>" onclick="NewWindow(this.href,'AlteracaoGrade','650','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
	<% end if %> <span style="font-size:8px">(<%=usuarion%>)</span>
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
end if 'fast=0
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>

<!--
Disciplinas do EAD atribuidas à professores
-->
<%
if fast=0 then

sql3="select g.id_grdaula, coddoc, codtur, chapa1, f.nome, codmat, materia, diasem, descricao, codhor, id_grdturma, codsituacao, usuarioc, usuarioa "  & _
"from g2ch g left join dc_professor f on g.chapa1=f.chapa where deletada=0 and g.perlet in " & periodo2 & " and '" & dtaccess(datacheck) & "' between inicio and termino " & _
"and ( " & _
"(g.codmat in ('G1436','G0006','G0245','G0508','G1573') ) " & _
"or (g.codmat in ('G0002','G0050','G0191','G0197') and coddoc in ('AEM','CCO','CCT','CNI','QUI','CBI','FIS') ) /*filosofia e etica*/ " & _
"or (g.codmat in ('G0096','G0073','G0437') and coddoc='CCO') /*comunicacao e expressao*/ " & _
"or (g.codmat in ('G0001','G0052','G0471') and coddoc='CCT') /*portugues*/ " & _
"or (g.codmat in ('G0001','G0052','G0471') and coddoc='DDI') /*portugues*/ " & _
"or (g.codmat in ('G0001','G0052','G0471') and coddoc='EDF') /*portugues*/ " & _
"or (g.codmat in ('G0001','G0052','G0471') and coddoc='FAR') /*portugues*/ " & _
"or (g.codmat in ('G0001','G0052','G0471') and coddoc='FIS') /*portugues*/ " & _
"or (g.codmat in ('G0001','G0052','G0471') and coddoc='QUI') /*portugues*/ " & _
"or (g.codmat in ('G0001','G0052','G0471') and coddoc='SIN') /*portugues*/ " & _
"or (g.codmat in ('G0001','G0052','G0471') and coddoc='TIN') /*portugues*/ " & _
"or (g.codmat in ('G0096','G0073','G0437') and coddoc='TOE') /*comunicacao e expressao*/ " & _
"or (g.codmat in ('G0001','G0052','G0471') and coddoc='TGC') /*portugues*/ " & _
"or (g.codmat in ('G0001','G0052','G0471') and coddoc='TGF') /*portugues*/ " & _
"or (g.codmat in ('G0001','G0052','G0471') and coddoc='TGE') /*portugues*/ " & _
"or (g.codmat in ('G0096','G0073','G0437') and coddoc='TGE') /*comunicacao e expressao*/ " & _
"or (g.codmat in ('G0096','G0073','G0437') and coddoc='TRC') /*comunicacao e expressao*/ " & _
") and chapa1<>'99998' " & _
"order by coddoc, codtur "

rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Disciplinas do EAD atribuidas à professores - Codigo correto: 99998 (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center">Curso</td>
	<td class=titulor align="center">Turma</td>
	<td class=titulor align="center">Dia</td>
	<td class=titulor align="center">Aula</td>
	<td class=titulor align="center">Disciplina</td>
	<td class=titulor align="center">Professor</td>
	<td class=titulor align="center">Sit</td>
	<td class=titulor align="center">&nbsp;    </td>
</tr>
<%
rs.movefirst:do while not rs.eof
msg=""
if isnull(rs("usuarioa")) then usuario=rs("usuarioc") else usuario=rs("usuarioa")
sql="select nome from usuarios where usuario='" & usuario & "'"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then usuarion=rs2("nome") else usuarion=""
rs2.close
%>
<tr>
	<td class="campor"><%=rs("coddoc")%></td>
	<td class="campor"><%=rs("codtur")%></td>
    <td class="campor"><%=weekdayname(rs("diasem"),1)%></td>
	<td class="campor"><%=rs("descricao")%></td>
    <td class="campor"><%=rs("materia")%></td>
    <td class="campor"><%=rs("nome")%></td>
    <td class="campor"><%=rs("codsituacao")%></td>
	<td class="campor">
	<% if session("a80")="T" then %>
	<a href="gradenovaaula.asp?idaula=<%=rs("id_grdaula")%>&idturma=<%=rs("id_grdturma")%>&codhor=<%=rs("codhor")%>" onclick="NewWindow(this.href,'AlteracaoGrade','650','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
	<% end if %> <span style="font-size:8px">(<%=usuarion%>)</span>
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
end if 'fast=0
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>

<!--
Diferenças entre disciplinas lançadas em relação as grades
-->
<%
if fast=0 then
sql3a="select t.coddoc, t.codcur, t.codper, t.grade, t.codtur, t.serie, u.codmat, m.materia, aulasgc=case when u.naulassem is null then 0 else u.naulassem end " & _
", z.codtur, aulasgh=case when aulas is null then 0 else aulas end, " & _
"dif=(case when aulas is null then 0 else aulas end)-(case when u.naulassem is null then 0 else u.naulassem end), usuarioa, usuarioc " & _
"from ((g2turmas t inner join corporerm.dbo.ugrade u on t.codcur=u.codcur and t.codper=u.codper and t.grade=u.grade and t.serie=u.periodo) " & _
"inner join corporerm.dbo.umaterias m on m.codmat=u.codmat) left join ( " & _
"select t.coddoc, t.codcur, t.codper, t.grade, t.codtur, t.serie, g.codmat, aulas=count(codmat), max(usuarioa) usuarioa, max(usuarioc) usuarioc " & _
"from g2aulas g right join g2turmas t on g.id_grdturma=t.id_grdturma " & _
"where t.perlet in " & periodo2 & " and deletada=0 and ativo=1 " & _
"and '" & dtaccess(datacheck) & "' between inicio and termino " & _
"group by t.coddoc, t.codcur, t.codper, t.grade, t.codtur, t.serie, g.codmat " & _
") z on z.codcur=u.codcur and z.codper=u.codper and z.grade=u.grade and z.codtur=t.codtur and z.codmat=u.codmat collate database_default " & _
"where t.perlet in " & periodo2 & " " & _
"and  (case when aulas is null then 0 else aulas end)-(case when u.naulassem is null then 0 else u.naulassem end)<>0 " & _
"order by t.coddoc, t.codtur "

rs.Open sql3a, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Divergências entre quantidades de aulas lançadas em relação à grade curricular (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center" rowspan=2>Curso</td>
	<td class=titulor align="center" rowspan=2>Turma</td>
	<td class=titulor align="center" rowspan=2>Disciplina</td>
	<td class=titulor align="center" colspan=2>Aulas</td>
	<td class=titulor align="center" rowspan=2>&nbsp;    </td>
</tr>
<tr><td class=titulor align="center">lançadas</td>
	<td class=titulor align="center">na GC</td>

<%
rs.movefirst:do while not rs.eof
msg=""
if isnull(rs("usuarioa")) then usuario=rs("usuarioc") else usuario=rs("usuarioa")
sql="select nome from usuarios where usuario='" & usuario & "'"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then usuarion=rs2("nome") else usuarion=""
rs2.close
%>
<tr>
	<td class="campor" height=13><%=rs("coddoc")%></td>
	<td class="campor"><%=rs.fields(4).value%></td>
    <td class="campor"><%=rs("materia")%></td>
	<td class="campor" align="center"><%=rs("aulasgh")%></td>
    <td class="campor" align="center"><%=rs("aulasgc")%></td>
	<td class="campor">
	<% if session("a80")="T" then %>
	<% end if %> <span style="font-size:8px">(<%=usuarion%>)</span>
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
end if 'fast=0
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>


<!--
turmas sem lançamentos
-->
<%
if fast=0 then
sql3="select t.coddoc, c.curso, t.perlet, t.codtur, count(a.id_grdturma) lançamentos " & _
"from (g2turmas t left join (select * from g2aulas where deletada=0) a on a.id_grdturma=t.id_grdturma) " & _
"inner join g2cursoeve c on c.coddoc=t.coddoc " & _
"where perlet in " & periodo2 & "  " & _
"group by t.coddoc, c.curso, t.perlet, t.codtur " & _
"having count(a.id_grdturma)<16 "

rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Turmas sem lançamentos ou com lançamentos incompletos (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center">Curso</td>
	<td class=titulor align="center">Per.Letivo</td>
	<td class=titulor align="center">Turma</td>
	<td class=titulor align="center"># Lanç.</td>
	<td class=titulor align="center">&nbsp;    </td>
</tr>
<%
rs.movefirst:do while not rs.eof
msg=""
%>
<tr>
	<td class="campor" height=16><%=rs("curso")%></td>
	<td class="campor"><%=rs("perlet")%></td>
	<td class="campor" align="center"><%=rs("codtur")%></td>
    <td class="campor" align="center"><%=rs("lançamentos")%></td>
	<td class="campor">
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
end if 'fast=0
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>


<% 
end if 

'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
if status="A2" then 
%>

<!--
Professores acima do limite
-->
<%
antigo=1
if antigo=0 and fast=0 then
sql3="select z.chapa, f.nome, z.aulas, limite=case when limite is null then 20 else limite end from (( " & _
"select chapa1 chapa, aulas=sum(ta) " & _
"from g2ch g where deletada=0 and juntar=0 and chapa1<'10000'  " & _
"and g.perlet in " & periodo2 & " and '" & dtaccess(datacheck) & "' between inicio and termino  " & _
"group by chapa1 " & _
") z left join g2limite l on z.chapa=l.chapa) left join grades_aux_prof f on f.chapa=z.chapa " & _
"where z.aulas>case when limite is null then 20 else limite end "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Professores com aulas atribuidas acima do limite individual (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center" rowspan=2>Chapa</td>
	<td class=titulor align="center" rowspan=2>Nome do professor</td>
	<td class=titulor align="center" colspan=2>Aulas</td>
	<td class=titulor align="center" rowspan=2>&nbsp;    </td>
</tr>
<tr><td class=titulor align="center">atribuidas</td>
	<td class=titulor align="center">Limite</td>

<%
rs.movefirst:do while not rs.eof
msg=""
%>
<tr>
	<td class="campor" height=16><%=rs("chapa")%></td>
	<td class="campor"><%=rs("nome")%></td>
	<td class="campor" align="center"><%=rs("aulas")%></td>
    <td class="campor" align="center"><%=rs("limite")%></td>
	<td class="campor">
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>
<%
end if 'fast=0
%>

<!--
Professores abaixo do limite
-->
<%
if fast=0 then
sql3="select z.chapa, f.nome, z.aulas, limite=case when limite is null then 20 else limite end from (( " & _
"select chapa1 chapa, aulas=sum(ta) " & _
"from g2ch g where deletada=0 and juntar=0 and chapa1<'10000'  " & _
"and g.perlet in " & periodo2 & " and '" & dtaccess(datacheck) & "' between inicio and termino  " & _
"group by chapa1 " & _
") z left join g2limite l on z.chapa=l.chapa) left join grades_aux_prof f on f.chapa=z.chapa " & _
"where z.aulas<case when limite is null then 20 else limite end "
sql3="select * from ( " & _
"select f.chapa, nome, codsituacao, situacao=case when codsituacao='A' then 'Ativo' else 'Afastado' end, " & _
"Limite=limite, Anterior=s0.aulas, Atual=s1.aulas " & _
",status=case when s1.aulas<s0.aulas then '1.Diminuição da carga horária' " & _
"	when s1.aulas>s0.aulas and s1.aulas<=limite then '2.Aumento da carga horária' " & _
"	when s1.aulas>s0.aulas and s1.aulas>limite then '3.Aumento e excesso limite' " & _
"	when s1.aulas>s0.aulas then '2.Aumento da carga horária' " & _
"	when s1.aulas is null and s0.aulas>0 then '4.Sem aulas atribuidas' " & _
"	else '' end " & _
"from dc_professor f " & _
"left join g2limite l on l.chapa=f.chapa " & _
"left join g2semestrei s1 on s1.chapa1=f.chapa and s1.inicio='" & dtaccess(inicioper) & "' " & _
"left join g2semestret s0 on s0.chapa1=f.chapa and s0.termino='" & dtaccess(fimper) & "' " & _
"where codsituacao<>'D' and codtipo='N' " & _
") z order by status desc, codsituacao, nome " 

rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Professores com aulas atribuidas abaixo/acima do limite individual (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center" rowspan=2>Chapa</td>
	<td class=titulor align="center" rowspan=2>Nome do professor</td>
	<td class=titulor align="center" rowspan=2>Situação</td>
	<td class=titulor align="center" rowspan=2>Limite (*)</td>

	<td class=titulor align="center" colspan=2>Aulas</td>
	<td class=titulor align="center" rowspan=2>Status</td>
	<td class=titulor align="center" rowspan=2>&nbsp;    </td>
</tr>
<tr><td class=titulor align="center">Anterior</td>
	<td class=titulor align="center">Atual</td>

<%
rs.movefirst:do while not rs.eof
msg=""
if ultimostatus<>rs("status") then
	response.write "<tr><td class=grupo colspan=8>" & rs("status") & "</td></tr>"
end if
%>
<tr>
	<td class="campor" height=16><%=rs("chapa")%></td>
	<td class="campor"><%=rs("nome")%></td>
	<td class="campor"><%=rs("situacao")%></td>
    <td class="campor" align="center"><font color=gray><%=rs("limite")%></td>

	<td class="campor" align="center"><%=rs("anterior")%></td>
    <td class="campor" align="center"><%=rs("atual")%></td>

    <td class="campor" align="left"><%=rs("status")%></td>
	<td class="campor">
	</td>
</tr>
<%
ultimostatus=rs("status"):rs.movenext:loop
end if
rs.close
end if 'fast=0
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>

<% 
end if 


'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
if status="D" then 
%>

<!--
professores demitidos ou afastados
-->
<%
if fast=0 then
sql3="select g.id_grdaula, coddoc, codtur, chapa1, f.nome, codmat, materia, diasem, descricao, codhor, id_grdturma, codsituacao, usuarioc, usuarioa " & _
"from g2ch g left join dc_professor f on g.chapa1=f.chapa where deletada=0 and g.perlet in " & periodo2 & " and '" & dtaccess(datacheck) & "' between inicio and termino " & _
"and f.codsituacao not in ('A','F','Z') " & _
"order by coddoc, codtur, diasem "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Professores demitidos ou afastados com aulas atribuidas (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center">Curso</td>
	<td class=titulor align="center">Turma</td>
	<td class=titulor align="center">Dia</td>
	<td class=titulor align="center">Aula</td>
	<td class=titulor align="center">Disciplina</td>
	<td class=titulor align="center">Professor</td>
	<td class=titulor align="center">Sit</td>
	<td class=titulor align="center">&nbsp;</td>
</tr>
<%
rs.movefirst:do while not rs.eof
msg=""
if isnull(rs("usuarioa")) then usuario=rs("usuarioc") else usuario=rs("usuarioa")
sql="select nome from usuarios where usuario='" & usuario & "'"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then usuarion=rs2("nome") else usuarion=""
rs2.close
%>
<tr>
	<td class="campor"><%=rs("coddoc")%></td>
	<td class="campor" nowrap><%=rs("codtur")%></td>
    <td class="campor"><%=weekdayname(rs("diasem"),1)%></td>
	<td class="campor" nowrap><%=rs("descricao")%></td>
    <td class="campor"><%=rs("materia")%></td>
    <td class="campor"><%=left(rs("nome"),30)%></td>
    <td class="campor"><%=rs("codsituacao")%></td>
	<td class="campor" nowrap>
	<% if session("a80")="T" then %>
	<a href="gradenovaaula.asp?idaula=<%=rs("id_grdaula")%>&idturma=<%=rs("id_grdturma")%>&codhor=<%=rs("codhor")%>" onclick="NewWindow(this.href,'AlteracaoGrade','650','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
	<% end if %> <span style="font-size:8px">(<%=usuarion%>)</span>
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
end if 'fast=0
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>

<!--
Aulas de professores demitidos NÃO atribuidas a professores novos
-->
<%
if fast=0 then
sql3a="select g.id_grdaula, g.coddoc, g.codtur, g.chapa1, f.nome, f.tab_grade, f.instrucaomec, s.valoraula salnovo, " & _
" g.codmat, g.materia, diasem, descricao, codhor, id_grdturma, codsituacao " & _
", z.chapa1 achapa, z.nome anome, z.instrucaomec ainstrucaomec, z.tab_grade atab_grade, s2.valoraula saldem, usuarioc, usuarioa " & _
"from (((g2ch g left join dc_professor f on g.chapa1=f.chapa) " & _
"inner join ( " & _
"	select distinct g.coddoc, codtur, chapa1, f.nome, codmat, materia, codnivelsal, titulacaopaga, instrucaomec, tab_ref, tab_grade " & _
"	from (g2ch g left join dc_professor f on g.chapa1=f.chapa) " & _
"	where deletada=0 and g.perlet in ('2013/0','2013/1') and f.codsituacao in ('D') and '" & dtaccess(fimper) & "' between inicio and termino " & _
") z on z.coddoc=g.coddoc and left(z.codtur,7)=left(g.codtur,7) and z.codmat=g.codmat " & _
") inner join salarios_curso_faixa s on s.coddoc=g.coddoc and s.titulacao=f.titulacaopaga collate database_default and s.nivel=f.codnivelsal collate database_default and s.reformulacao=f.tab_ref " & _
") inner join salarios_curso_faixa s2 on s2.coddoc=z.coddoc and s2.titulacao=z.titulacaopaga collate database_default and s2.nivel=z.codnivelsal collate database_default and s2.reformulacao=z.tab_ref " & _
"where excecao=0 and deletada=0 and g.perlet in " & periodo2 & "  and '" & dtaccess(inicioper) & "' between inicio and termino " & _
"and (f.tab_grade>=3) order by g.coddoc, g.codtur, g.diasem "
'response.write sql3a
rs.Open sql3a, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Aulas de professores demitidos NÃO atribuidas a professores novos (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center">Curso</td>
	<td class=titulor align="center">Turma</td>
	<td class=titulor align="center">Dia</td>
	<td class=titulor align="center">Aula</td>
	<td class=titulor align="center">Disciplina</td>
	<td class=titulor align="center">Professor antigo<br>Professor novo</td>
	<td class=titulor align="center">&nbsp;    </td>
</tr>
<%
rs.movefirst:do while not rs.eof
msg=""
if isnull(rs("usuarioa")) then usuario=rs("usuarioc") else usuario=rs("usuarioa")
sql="select nome from usuarios where usuario='" & usuario & "'"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then usuarion=rs2("nome") else usuarion=""
rs2.close
%>
<tr>
	<td class="campor"><%=rs("coddoc")%></td>
	<td class="campor"><%=rs("codtur")%></td>
    <td class="campor"><%=weekdayname(rs("diasem"),1)%></td>
	<td class="campor" nowrap><%=rs("descricao")%></td>
    <td class="campor"><%=rs("materia")%></td>
    <td class="campor" nowrap><font color=red><%=left(rs("anome"),30)%> (<%=left(rs("ainstrucaomec"),1)&"-"&rs("atab_grade")%>) : <%if session("usuariogrupo")="RH" then response.write rs("saldem")%>
		<br><font color=blue><%=left(rs("nome"),30)%> (<%=left(rs("instrucaomec"),1)&"-"&rs("tab_grade")%>) : <%if session("usuariogrupo")="RH" then response.write rs("salnovo")%></td>
	<td class="campor">
	<% if session("a80")="T" then %>
	<a href="gradenovaaula.asp?idaula=<%=rs("id_grdaula")%>&idturma=<%=rs("id_grdturma")%>&codhor=<%=rs("codhor")%>" onclick="NewWindow(this.href,'AlteracaoGrade','650','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
	<% end if %> <span style="font-size:8px">(<%=usuarion%>)</span>
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
end if 'fast=0
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>

<% 
end if 

'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
if status="C" then 
%>
<!--
Professores sem aulas atribuidas no semestre
-->
<%
if fast=0 then
sql3="select f.chapa, f.nome, codsecao, aula_ant, aula_atual, codsituacao from (dc_professor f " & _
"left join (select chapa1, aula_atual=sum(ta) from g2ch g where deletada=0 and g.perlet in ('2013/0','2013/2') and '08/01/2013' between inicio and termino and chapa1<'10000' group by chapa1) a on a.chapa1=f.chapa " & _
") left join (select chapa1, aula_ant=sum(ta) from g2ch g where deletada=0 and g.perlet in ('2013/0','2013/1') and '07/31/2013' between inicio and termino and chapa1<'10000' group by chapa1) b on b.chapa1=f.chapa " & _
"where codsituacao<>'D' and aula_atual is null and (aula_ant is not null or aula_ant is null)  " & _
"and f.chapa not in (select chapa collate database_default from quem_nomeacoes) and codtipo='N' " & _
"and f.chapa not in ('00057','00056','00066','00063','01514','01164','01165','02768','02820','02823') and codsituacao in ('A','F','E','Z') "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Professores sem aulas atribuidas no semestre (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center" rowspan=2>Chapa</td>
	<td class=titulor align="center" rowspan=2>Professor</td>
	<td class=titulor align="center" rowspan=2>Cod.Seção</td>
	<td class=titulor align="center" colspan=2>Aulas no semestre</td>
	<td class=titulor align="center" rowspan=2>&nbsp;    </td>
</tr>
<tr><td class=titulor align="center">anterior</td>
	<td class=titulor align="center">atual</td>
<%
rs.movefirst:do while not rs.eof
msg=""
'if isnull(rs("usuarioa")) then usuario=rs("usuarioc") else usuario=rs("usuarioa")
'sql="select nome from usuarios where usuario='" & usuario & "'"
'rs2.Open sql, ,adOpenStatic, adLockReadOnly
'if rs2.recordcount>0 then usuarion=rs2("nome") else usuarion=""
'rs2.close
%>
<tr>
	<td class="campor"><%=rs("chapa")%></td>
	<td class="campor"><%=rs("nome")%></td>
	<td class="campor" nowrap><%=rs("codsecao")%></td>
    <td class="campor" align="center"><%=rs("aula_ant")%></td>
    <td class="campor" align="center"><%=rs("aula_atual")%></td>
    <td class="campor" align="center"><%=rs("codsituacao")%></td>
</tr>
<%
rs.movenext:loop
end if
rs.close
end if 'fast=0
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>

<!--
Aulas sem professor atribuido ou a ser contratado
-->
<%
if fast=0 then
sql3="select g.id_grdaula, coddoc, codtur, chapa1, f.nome, codmat, materia, diasem, descricao, codhor, id_grdturma, codsituacao, usuarioa, usuarioc " & _
"from g2ch g left join grades_aux_prof f on g.chapa1=f.chapa " & _
"where deletada=0 and g.perlet in " & periodo2 & " " & _
"and (g.chapa1 like '99%' or chapa1='0') and g.chapa1 not in ('99998') and '" & dtaccess(datacheck) & "' between inicio and termino " & _
"order by coddoc, diasem, codtur, descricao, f.nome, materia "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Aulas sem professor atribuido ou a ser contratado (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center">Curso</td>
	<td class=titulor align="center">Turma</td>
	<td class=titulor align="center">Dia</td>
	<td class=titulor align="center">Aula</td>
	<td class=titulor align="center">Disciplina</td>
	<td class=titulor align="center">Professor</td>
	<td class=titulor align="center">&nbsp;    </td>
</tr>
<%
rs.movefirst:do while not rs.eof
msg=""
if isnull(rs("usuarioa")) then usuario=rs("usuarioc") else usuario=rs("usuarioa")
sql="select nome from usuarios where usuario='" & usuario & "'"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then usuarion=rs2("nome") else usuarion=""
rs2.close
if len(rs("materia"))>40 then materia="<span style=""font-size:8px"">" & rs("materia") & "</span>" else materia=rs("materia")
%>
<tr>
	<td class="campor"><%=rs("coddoc")%></td>
	<td class="campor" nowrap><%=rs("codtur")%></td>
    <td class="campor"><%=weekdayname(rs("diasem"),1)%></td>
	<td class="campor" nowrap><%=rs("descricao")%></td>
    <td class="campor"><%=materia%></td>
    <td class="campor" nowrap><%=left(rs("nome"),30)%></td>
	<td class="campor" nowrap>
	<% if session("a80")="T" then %>
	<a href="gradenovaaula.asp?idaula=<%=rs("id_grdaula")%>&idturma=<%=rs("id_grdturma")%>&codhor=<%=rs("codhor")%>" onclick="NewWindow(this.href,'AlteracaoGrade','650','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
	<% end if %> <span style="font-size:8px">(<%=usuarion%>)</span>
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
end if 'fast=0
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>

<%
end if 'status=C

'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
if status="J" then 
%>

<!--
Inconsistencias na junção de turmas
-->
<%
if fast=0 then
sql="update g2aulas set jpai=0 " & _
"where id_grdaula in ( " & _
"select id_grdaula from ( " & _
"select a.id_grdaula, a.jpai, (select count(id_grdaula) j from g2aulas where juntar_id=a.id_grdaula and deletada=0 /*and termino<getdate()*/) juntadas " & _
"from g2aulas a " & _
"where jpai=1 and deletada=0 " & _
") z where juntadas=0 " & _
") "
conexao.execute sql
if session("usuariomaster")="02379" then complsql="where chapa1>='0' " else complsql=""
sql3="select * from ( " & _
"select id_grdaula, coddoc, codtur, chapa1, f.nome, codmat, materia, g.diasem, g.descricao, codhor, id_grdturma, inicio, termino, status, usuarioa, usuarioc, juntar,jpai " & _
"from (g2ch g inner join ( " & _
"select chapa, diasem, descricao, taula, tjuntar, tjpai, status " & _
"from (select chapa, diasem, descricao, taula, tjuntar, tjpai, status=case when (tjuntar+tjpai)>0 and taula=1 then 'Checar - Não junta' else case when (tjuntar+tjpai)<>taula and taula>1 then 'Checar - Junta' else null end end " & _
"from (select chapa1 chapa, diasem, descricao, taula=count(id_grdaula), tjuntar=sum(case when juntar=1 then 1 else 0 end), tjpai=sum(case when jpai=1 then 1 else 0 end) from g2ch g where deletada=0 and chapa1<'10000' and perlet in " & periodo2 & " and '" & dtaccess(datacheck) & "' between inicio and termino group by chapa1, diasem, descricao) z ) y " & _
"where status is not null ) x on x.chapa=g.chapa1 and x.diasem=g.diasem and x.descricao=g.descricao) " & _
"inner join corporerm.dbo.pfunc /*grades_aux_prof*/ f on g.chapa1=f.chapa collate database_default " & _
"where deletada=0 and ativo=1 and perlet in " & periodo2 & " and '" & dtaccess(datacheck) & "' between inicio and termino " & _
"union " & _
"select id_grdaula, coddoc, codtur, chapa1, f.nome, codmat, materia, g.diasem, g.descricao, codhor, id_grdturma, inicio, termino, status='Não junta', usuarioa, usuarioc, juntar,jpai " & _
"from g2ch g inner join corporerm.dbo.pfunc /*grades_aux_prof*/ f on g.chapa1=f.chapa collate database_default where juntar=1 and chapa1 not like '99%' and juntar_id not in (select id_grdaula from g2aulas) and deletada=0 and perlet in " & periodo2 & " and '" & dtaccess(datacheck) & "' between inicio and termino " & _
") j " & complsql & _
"order by chapa1, diasem, descricao, juntar " 
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Inconsistências na junção de turmas (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center">Curso</td>
	<td class=titulor align="center">Turma</td>
	<td class=titulor align="center">Dia</td>
	<td class=titulor align="center">Aula</td>
	<td class=titulor align="center">Disciplina</td>
	<td class=titulor align="center">Professor</td>
	<td class=titulor align="center">Status</td>
	<td class=titulor align="center">&nbsp;    </td>
</tr>
<%
rs.movefirst:do while not rs.eof
msg=""
if isnull(rs("usuarioa")) then usuario=rs("usuarioc") else usuario=rs("usuarioa")
sql="select nome from usuarios where usuario='" & usuario & "'"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then usuarion=rs2("nome") else usuarion=""
rs2.close
if len(rs("materia"))>40 then materia="<span style=""font-size:7px"">" & rs("materia") & "</span>" else materia="<span style=""font-size:8px"">" & rs("materia") & "</span>" 
%>
<tr>
	<td class="campor"><span style="font-size:8px"><%=rs("coddoc")%></span></td>
	<td class="campor" nowrap><span style="font-size:8px"><%=rs("codtur")%></span></td>
    <td class="campor"><span style="font-size:8px"><%=weekdayname(rs("diasem"),1)%></span></td>
	<td class="campor" nowrap><span style="font-size:8px"><%=rs("descricao")%></span></td>
    <td class="campor"><%=materia%></td>
    <td class="campor" nowrap><span style="font-size:8px"><%=left(rs("nome"),30)%></span></td>
    <td class="campor"><span style="font-size:8px"><%=rs("status")%></span></td>
	<td class="campor">
	<% if session("a80")="T" then %>
	<a href="gradenovaaula.asp?idaula=<%=rs("id_grdaula")%>&idturma=<%=rs("id_grdturma")%>&codhor=<%=rs("codhor")%>" onclick="NewWindow(this.href,'AlteracaoGrade','650','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
	<% end if %> <span style="font-size:8px">(<%=usuarion%>)</span>
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
end if 'fast=0
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>



<%
end if 'status=J


'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
'***********************************************************************************************************************************************************
if status="J2" then 
%>

<!--
junção de turmas
-->
<%
if fast=0 then
sql3="select id_grdaula, coddoc, codtur, chapa1, nome, codmat, materia, diasem, descricao, codhor, id_grdturma, inicio, termino, usuarioa, usuarioc, juntar, juntar_id, jturma, jpai, dividir " & _
"from g2ch g inner join grades_aux_prof f on g.chapa1=f.chapa where id_grdaula in (" & _
"select distinct juntar_id from g2ch g where chapa1<'10000' and juntar=1 and deletada=0 and g.perlet in " & periodo2 & " and '" & dtaccess(datacheck) & "' between inicio and termino " & _
") and chapa1<'10000' order by chapa1, diasem, descricao"
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="700" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8 height=20 valign=middle align="center">Disciplinas e professores com junção de aulas (<%=total%>)</td></tr>
<%
if rs.recordcount>0 then
%>
<tr><td class=titulor align="center">Curso</td>
	<td class=titulor align="center">Turma</td>
	<td class=titulor align="center">Dia</td>
	<td class=titulor align="center">Aula</td>
	<td class=titulor align="center">Disciplina<br>Professor</td>
	<td class=titulor align="center">&nbsp;    </td>
	<td class=titulor align="center">Junta com </td>
</tr>
<%
rs.movefirst:do while not rs.eof
msg=""
if isnull(rs("usuarioa")) then usuario=rs("usuarioc") else usuario=rs("usuarioa")
sql="select nome from usuarios where usuario='" & usuario & "'"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then usuarion=rs2("nome") else usuarion=""
rs2.close
%>
<tr>
	<td class="campor" style="border-bottom:2px solid #000000" valign=top><%=rs("coddoc")%></td>
	<td class="campor" style="border-bottom:2px solid #000000" valign=top><%=rs("codtur")%></td>
    <td class="campor" style="border-bottom:2px solid #000000" valign=top><%=weekdayname(rs("diasem"),1)%></td>
	<td class="campor" style="border-bottom:2px solid #000000" valign=top nowrap><%=rs("descricao")%></td>
    <td class="campor" style="border-bottom:2px solid #000000" valign=top><i><%=rs("materia")%></i><br><b><%=left(rs("nome"),30)%></td>
	<td class="campor" style="border-bottom:2px solid #000000" valign=top>
	<% if session("a80")="T" then %>
	<a href="gradenovaaula.asp?idaula=<%=rs("id_grdaula")%>&idturma=<%=rs("id_grdturma")%>&codhor=<%=rs("codhor")%>" onclick="NewWindow(this.href,'AlteracaoGrade','650','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
	<% end if %>
	</td>
    <td class="campor" style="border-bottom:2px solid #000000" valign=top nowrap>
	<table border="1" bordercolor="#CCCCCC" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" >
<%	
teste=99
IF teste=99 THEN
sqlf="select id_grdaula, coddoc, codtur, chapa1, codmat, materia, diasem, descricao, codhor, id_grdturma, inicio, termino, juntar, juntar_id, jturma, jpai, dividir " & _
"from g2ch g where juntar_id=" & rs("id_grdaula") & " order by chapa1, diasem, descricao"
rs2.Open sqlf, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>	
	<tr>
		<td	class="campor" style="border-bottom:2px dotted #000000" width=20><%=rs2("coddoc")%></td>
		<td	class="campor" style="border-bottom:2px dotted #000000" width=40><%=rs2("codtur")%></td>
	    <td class="campor" style="border-bottom:2px dotted #000000" width=100%><%=rs2("materia")%></td>
		<td class="campor" style="border-bottom:2px dotted #000000" width=18valign=top>
		<% if session("a80")="T" then %>
		<a href="gradenovaaula.asp?idaula=<%=rs2("id_grdaula")%>&idturma=<%=rs2("id_grdturma")%>&codhor=<%=rs2("codhor")%>" onclick="NewWindow(this.href,'AlteracaoGrade','650','400','yes','center');return false" onfocus="this.blur()">
		<img src="../images/Folder95O.gif" border="0" width=13></a>
		<% end if %>
		</td>
	</tr>
<%
rs2.movenext:loop
rs2.close
END IF
%>	
	</table>
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
end if 'fast=0
%>
	<tr><td class="campoa" colspan=8><%=total%> registros</td></tr>
</table>
<br>


<%
end if 'status=J2

termino=now()
duracao=termino-inicio
response.write "<br>" & formatdatetime(duracao,3)
conexao.close
set conexao=nothing
set rs=nothing
%>
</form>
</body>
</html>