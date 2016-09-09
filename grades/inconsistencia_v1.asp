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
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
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
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
sessao=session.sessionid
inicio=now()
%>
<p class=titulo style="margin-top: 0; margin-bottom: 0">Checagem de Inconsistências na Grade Horária</p>
<%
'***** checagem de horarios duplicados *****
%>
<a href="#hordupl" title="Horários de Aulas Duplicados"></a>
<a href="#profhordupl" title="Professores com Horários de Aulas Duplicados"></a>
<a href="#afastados" title="Professores Afastados ou Licenciados"></a>
<br>
<a name="hordupl"></a><table border="1" bordercolor="#CCCCCC" cellpadding="0" width="690" cellspacing="0" style="border-collapse: collapse">
  <tr><td class=grupo colspan=14>HORÁRIOS DE AULAS DUPLICADOS</td></tr>
  <tr>
    <td class=titulor align="center">P.Let.</td>
    <td class=titulor align="center">Curso     </td>
    <td class=titulor align="center">Per.      </td>
    <td class=titulor align="center">Turma     </td>
    <td class=titulor align="center">Dia       </td>
    <td class=titulor align="center" colspan=6>Aula      </td>
    <td class=titulor align="center">Disciplina/Professor</td>
    <td class=titulor align="center">j/d       </td>
    <td class=titulor align="center">&nbsp;    </td>
  </tr>
<%
fast=0
fast2=0
fast3=0
fast4=0
fast5=0
periodo="2009%2"

if fast=0 then
sql3="select g2.*, f.nome as professor from grades_2 g2, " & _
"(select g.id_grade from g2ch g, " & _
"(SELECT perlet, perlet2, coddoc, curso, turno, codtur, diasem, posicao, Count(id_grade) AS vezes FROM g2ch WHERE perlet2 like '" & periodo & "' and deletada=0 and ativo=1 and dividir=0 and prof=1 GROUP BY perlet, perlet2, coddoc, curso, turno, codtur, diasem, posicao HAVING Count(id_grade)>1) t " & _
"where t.perlet=g.perlet and t.perlet2=g.perlet2 and t.coddoc=g.coddoc and t.turno=g.turno and t.codtur=g.codtur and t.diasem=g.diasem and t.posicao=g.posicao group by g.id_grade) t2, " & _
"grades_aux_prof AS f " & _
"where f.chapa=g2.chapa1 and g2.id_grade=t2.id_grade " & _
"order by g2.coddoc, g2.serie, g2.turma, g2.turno, g2.diasem, g2.a5,g2.a3,g2.a1  "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
'if rs("dividir")=-1 then rs.movenext
if rs("turno")="1" then turno="Mat"
if rs("turno")="2" then turno="Vesp"
if rs("turno")="3" then turno="Not"
if rs("turno")="5" then turno="Vesp-EF"
if rs("turno")="61" then turno="Integral"
if rs("turno")="62" then turno="Integral"

professor1=rs("professor")
msg=""
if isnull(rs("usuarioa")) then usuario=rs("usuarioc") else usuario=rs("usuarioa")
sql="select nome from usuarios where usuario='" & usuario & "'"
rsc.Open sql, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then usuarion=rsc("nome") else usuarion=""
rsc.close
'if rs("dividir")=-1 then rs.movenext
%>
  <tr>
    <td class="campor"><%=rs("perlet")%></td>
    <td class="campor"><%=rs("curso") %></td>
    <td class="campor"><%=turno%></td>
    <td class="campor" align="center"><%=rs("codtur")%></td>
    <td class="campor"><%=weekdayname(rs("diasem"),1) %></td>
	<td class="campor" nowrap>
	<%if rs("a1")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class="campor" nowrap>
	<%if rs("a2")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class="campor" nowrap>
	<%if rs("a3")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class="campor" nowrap>
	<%if rs("a4")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class="campor" nowrap>
	<%if rs("a5")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class="campor" nowrap>
	<%if rs("a6")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>

    <td class="campor"><%=rs("materia") %>&nbsp;<font color="#FF0000"><%=msg%></font><br><font color="#0000FF"><%=professor1%></font> (<%=rs("chapa1")%>)
	</td>
    <td class="campor"><%=formatnumber(rs("juntar"),0) & "/" & formatnumber(rs("dividir"),0)%></td>

	<td class="campor">
    <% if session("a80")="T" then %>
      <a href="grade_alteracao.asp?codigo=<%=rs("id_grade")%>" onclick="NewWindow(this.href,'AlteracaoGrade','550','330','no','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a><%=usuarion%>
	<% end if %>
	</td>
  </tr>
<%
rs.movenext
loop
end if
rs.close
end if 'fast=0
%>
  <tr><td class="campoa" colspan=14><%=total%> registros</td></tr>
</table>
<%
'********** checagem de professores com horários duplicados
%>
<br>
<a name="profhordupl"></a><table border="1" cellpadding="0" width="690" cellspacing="0" style="border-collapse: collapse">
  <tr><td class=grupo colspan=14>PROFESSORES COM HORÁRIOS DE AULAS DUPLICADOS</td></tr>
  <tr>
    <td class=titulor align="center">P.Let.</td>
    <td class=titulor align="center">Curso     </td>
    <td class=titulor align="center">Per.      </td>
    <td class=titulor align="center">Turma     </td>
    <td class=titulor align="center">Dia       </td>
    <td class=titulor align="center" colspan=6>Aula   </td>
    <td class=titulor align="center">Disciplina/Professor</td>
    <td class=titulor align="center">j/d     </td>
    <td class=titulor align="center">&nbsp;    </td>
  </tr>
<%
if fast2=0 then
sql3="select g2.*, f.nome as professor from grades_2 g2, " & _
"(select g.id_grade from g2ch g, " & _
"(SELECT chapa1, perlet3, diasem, turno, posicao, Count(id_grade) AS vezes FROM g2ch WHERE deletada=0 and ativo=1 AND juntar=0 and chapa1<>'99999' AND perlet2 like '" & periodo & "' GROUP BY chapa1, perlet3, diasem, turno, posicao HAVING Count(id_grade)>1) t " & _
"where t.chapa1=g.chapa1 and t.perlet3=g.perlet3 and t.diasem=g.diasem and t.turno=g.turno and t.posicao=g.posicao group by g.id_grade) t2, " & _
"grades_aux_prof AS f " & _
"where f.chapa=g2.chapa1 and g2.id_grade=t2.id_grade " & _
"order by g2.chapa1, g2.coddoc, g2.serie, g2.turma, g2.turno, g2.diasem, g2.a5,g2.a3,g2.a1  "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
if rs("turno")="1" then turno="Mat"
if rs("turno")="2" then turno="Vesp"
if rs("turno")="3" then turno="Not"
if rs("turno")="61" then turno="Integral"
if rs("turno")="62" then turno="Integral"

'if rs("horini")="" then horini="&nbsp;" else horini=formatdatetime(rs("horini"),4)
'if rs("horfim")="" then horfim="&nbsp;" else horfim=formatdatetime(rs("horfim"),4)
professor1=rs("professor")
msg=""
'if rs("pini")<>rs("inicio") then msg="Iniciou: " & rs("inicio")
'if rs("pfim")<>rs("termino") then msg="Encerrou: " & rs("termino")
if isnull(rs("usuarioa")) then usuario=rs("usuarioc") else usuario=rs("usuarioa")
sql="select nome from usuarios where usuario='" & usuario & "'"
rsc.Open sql, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then usuarion=rsc("nome") else usuarion=""
rsc.close
%>
  <tr>
    <td class="campor"><%=rs("perlet")%></td>
    <td class="campor"><%=rs("curso") %></td>
    <td class="campor"><%=turno%></td>
    <td class="campor" align="center" nowrap><%=rs("codtur")%></td>
    <td class="campor"><%=weekdayname(rs("diasem"),1) %></td>
	<td class="campor" nowrap>
	<%if rs("a1")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class="campor" nowrap>
	<%if rs("a2")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class="campor" nowrap>
	<%if rs("a3")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class="campor" nowrap>
	<%if rs("a4")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class="campor" nowrap>
	<%if rs("a5")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
	<td class="campor" nowrap>
	<%if rs("a6")=1 then response.write "<font face='Wingdings'>ü</font>" else response.write "&nbsp;&nbsp;" %></td>
    <td class="campor"><%=rs("materia") %>&nbsp;<font color="#FF0000"><%=msg%></font><br><font color="#0000FF"><%=professor1%></font> (<%=rs("chapa1")%>)
	</td>
    <td class="campor"><%=formatnumber(rs("juntar"),0) & "/" & formatnumber(rs("dividir"),0)%></td>

	<td class="campor">
    <% if session("a80")="T" then %>
      <a href="grade_alteracao.asp?codigo=<%=rs("id_grade")%>" onclick="NewWindow(this.href,'AlteracaoGrade','550','425','no','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a><%=usuarion%>
	<% end if %>
	</td>
	<td class="campor">

	</td>
  </tr>

<%
rs.movenext
loop
end if
rs.close
end if 'fast2
%>
  <tr><td class="campoa" colspan=14><%=total%> registros</td></tr>
</table>
<%
'********** checagem de professores afastados
%>
<br>
<a name="afastados"></a><table border="1" cellpadding="0" width="690" cellspacing="0" style="border-collapse: collapse">
  <tr><td class=grupo colspan=10>Professores Afastados/Licenciados com grade horária</td></tr>
  <tr>
    <td class=titulor align="center">P.Let.</td>
    <td class=titulor align="center">Curso     </td>
    <td class=titulor align="center">Professor</td>
    <td class=titulor align="center">Inicio    </td>
    <td class=titulor align="center">Term.     </td>
    <td class=titulor align="center">CHS     </td>
    <td class=titulor align="center">&nbsp;    </td>
  </tr>
<%
if fast3=0 then
sql3="SELECT g.perlet, g.curso, g.chapa1, g.inicio, g.termino, f.NOME AS PROFESSOR, min(t.aulas) as aulas " & _
"FROM g2ch AS g, grades_aux_prof AS f, " & _
"(select chapa1, count(aula) as aulas from g2ch where perlet2 like '" & periodo & "' group by chapa1) as t " & _
"WHERE g.chapa1 = f.CHAPA and g.chapa1=t.chapa1 " & _
"and g.id_grade>0 AND g.deletada=0  " & _
"AND f.codsituacao not in ('A','F','Z','D') and g.perlet2 like '" & periodo & "' " & _
"group by g.perlet, g.chapa1, g.curso, g.inicio, g.termino, f.NOME " & _
"ORDER BY g.chapa1, g.perlet "

rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
professor1=rs("professor")
msg=""
%>
  <tr>
    <td class="campor"><%=rs("perlet")%></td>
    <td class="campor"><%=rs("curso") %></td>
    <td class="campor"><%=professor1%> (<%=rs("chapa1")%>)</td>
    <td class="campor"><%=rs("inicio")%></td>
    <td class="campor"><%=rs("termino") %></td>
    <td class="campor" align="center"><%=rs("aulas") %></td>
	<td class="campor">
	</td>
  </tr>
<%
rs.movenext
loop
end if
rs.close
end if 'fast3
%>
  <tr><td class="campoa" colspan=10><%=total%> registros</td></tr>
</table>

<%
'********** checagem de professores sem aulas
%>
<br>
<a name="afastados"></a><table border="1" cellpadding="0" width="690" cellspacing="0" style="border-collapse: collapse">
  <tr><td class=grupo colspan=10>Professores Ativos sem atribuições na grade horária</td></tr>
  <tr>
    <td class=titulor align="center">Curso Anterior</td>
    <td class=titulor align="center">Professor</td>
    <td class=titulor align="center">Aulas Sem.Anterior</td>
    <td class=titulor align="center">Aulas Sem.Atual</td>
    <td class=titulor align="center">&nbsp;  </td>
  </tr>
<%
if fast4=0 then
sql3="SELECT g.chapa1, f.NOME, f.CODSITUACAO, f.secao as DESCRICAO, " & _
"'20082'=sum(case when g.perlet3='20082' then g.ta else 0 end), " & _
"'20091'=sum(case when g.perlet3='20091' then g.ta else 0 end) " & _
"FROM g2ch g LEFT JOIN qry_funcionarios f ON g.chapa1=f.CHAPA collate database_default " & _
"WHERE f.CODSITUACAO In ('A','F','Z') AND (g.perlet3='20082' Or g.perlet3='20091') " & _
"GROUP BY g.chapa1, f.NOME, f.CODSITUACAO, f.secao "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount:total=0
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
professor1=rs("nome")
msg=""
a1=rs("20082")
a2=rs("20091")
'if (isnull(rs("20082")) and rs("20091")<>0) or (isnull(rs("20091")) and rs("20082")<>0) then
if (a1=0 and a2<>0) or (a2=0 and a1<>0) then
if a1=0 then a1="-"
if a2=0 then a2="-"
%>
  <tr>
    <td class="campor"><%=rs("descricao") %></td>
    <td class="campor"><%=professor1%> (<%=rs("chapa1")%>)</td>
    <td class="campor" align="center"><%=a1 %></td>
    <td class="campor" align="center"><%=a2 %></td>
	<td class="campor">
	</td>
  </tr>
<%
total=total+1
end if
rs.movenext
loop
end if
rs.close
end if 'fast4
%>
  <tr><td class="campoa" colspan=10><%=total%> registros</td></tr>
</table>


<%
'********** checagem de 20 aulas
%>
<br>
<a name="acima20"><table border="1" cellpadding="0" width="690" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=10>Professores acima das 20 aulas semanais (só serão pagos após aprovação da Reitoria)</td></tr>
<tr>
	<td class=titulor align="center">P.Let.    </td>
	<td class=titulor align="center">Chapa     </td>
	<td class=titulor align="center">Professor</td>
	<td class=titulor align="center">Aulas     </td>
	<td class=titulor align="center">&nbsp;    </td>
</tr>
<%
if fast5=0 then
sql3="SELECT perlet3, g.chapa1, f.nome AS PROFESSOR, Count(g.codmat) AS aulas " & _
"FROM g2ch AS g, grades_aux_prof AS f " & _
"WHERE g.chapa1 = f.chapa collate database_default " & _
"AND g.id_grade>0 AND g.deletada=0 AND g.ativo=1 AND g.juntar=0  " & _
"AND g.perlet2 like '" & periodo & "' " & _
"GROUP BY perlet3, g.chapa1, f.nome " & _
"HAVING Count(g.codmat)>20 " & _
"ORDER BY Count(g.codmat) DESC, F.NOME "
'and getdate() between g.inicio and g.termino
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
professor1=rs("professor")
msg=""
%>
<tr>
	<td class="campor"><%=rs("perlet3")%></td>
	<td class="campor"><%=rs("chapa1") %></td>
	<td class="campor"><%=rs("professor") %></td>
	<td class="campor" align="center"><%=rs("aulas")%></td>
	<td class="campor"></td>
</tr>
<%
rs.movenext
loop
end if
rs.close
end if 'fast5
%>
<tr><td class="campoa" colspan=10><%=total%> registros</td></tr>
</table>

<%
termino=now()
duracao=termino-inicio
response.write "<br>" & formatdatetime(duracao,3)
conexao.close
set conexao=nothing
set rs=nothing
%>
</body>
</html>