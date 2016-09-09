<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("acesso")>2 then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
'accesso func 1 prof 2
if session("a100")="N" or session("a100")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Informações do Professor</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
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
<link rel="stylesheet" type="text/css" href="tabcontent.css" />
<script type="text/javascript" src="tabcontent.js">
/***********************************************
* Tab Content script v2.2- © Dynamic Drive DHTML code library (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit Dynamic Drive at http://www.dynamicdrive.com/ for full source code
***********************************************/
</script>
</head>
<body>
<%
dim conexao, rs, rs2, formato(2)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
if request.form("chapa")<>"" then 
	chapa=request.form("chapa") 
else 
	chapa=session("usuariomaster")
end if
corcheck="black"
pixel=96/2.54
point=72/2.54
pointp=72.27/2.54

if request.form="" then
end if ' request.form=""

if request.form<>"" then
	vez=request.form("vezes")
	for a=1 to vez
		info    =request.form("m" & a):anterior=request.form("a" & a)
		coddoc=request.form("coddoc")
		codcur=request.form("hcodcur")
		codper=request.form("hcodper")
		grade=request.form("hgrade")
		if info<>"" and anterior="" then
			sql="insert into grades_aderencia (chapa, coddoc, codcur, codper, grade, codmat) " & _
			"select '" & chapa & "', '" & coddoc & "'," & codcur & "," & codper & "," & grade & ",'" & info & "';"
			conexao.execute sql
		end if
		if info="" and anterior<>"" then
			sql="delete from grades_aderencia " & _
			"where chapa='" & chapa & "' and coddoc='" & coddoc & "' and codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and codmat='" & anterior & "' "
			conexao.execute sql
		end if
	next
end if 'request.form<>""

%>
<!-- -->
<!-- -->
<form method="POST" action="aderencia.asp" name="form">
<%
if session("acesso")=2 or session("usuariogrupo")="COORD.CURSO" then
%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" >
<tr><td valign=top style="border-right:3px double silver;border-bottom:3px double silver" width=150 height=600>
<!-- -->
<p style="margin-top:0;margin-bottom:0" class=titulo><%=session("usuarioname")%></p>
<hr>
<p style="margin-top:0;margin-bottom:5"><a href="../academic/disponibilidade.asp">
<img src="../images/Clock.gif" width="16" height="16" border="0" alt="">Disponibilidade</a></p>

<p style="margin-top:0;margin-bottom:5">
<img src="../images/BookO.gif" width="16" height="16" border="0" alt="">Aderência</p>

<br><br>
<p style="margin-top:0;margin-bottom:5"><a href="../academic/meusplanos.asp">
<img src="../images/BookO.gif" width="16" height="16" border="0" alt="">Plano de Ensino</a>

<br><br>
<p style="margin-top:0;margin-bottom:5"><a href="../academic/espelho.asp">
<img src="../images/espelho.jpg" width="16" height="16" border="0" alt="">Marcação de Ponto</a></p>

<br><br><br><br><br><br><br>
<p style="margin-top:0;margin-bottom:0"><a href="../indexp.asp">
<img src="../images/setafirst0.gif" width="12" height="12" border="0" alt="">Início</a>
<!-- -->
</td><td valign=top style="border-bottom:3px double silver">
<p style="margin-top:0;margin-bottom:10" class=titulo>Aderência à Grade Curricular</p>
<!-- -->
<%
else ' para acesso=1
%>
<p style="margin-top:0;margin-bottom:10" class=titulo>Aderência à Grade Curricular</p>
<select size=1 name="chapa" onchange="javascript:submit();">
	<option value="0">Selecione....</option>
<%
sqlc="select chapa, nome from grades_aux_prof where codsituacao<>'D' and chapa<'10000' order by nome"
sqlc="select p.chapa, p.nome, z.aderencia from grades_aux_prof p left join (select chapa, count(codmat) aderencia from grades_aderencia group by chapa) z on z.chapa=p.chapa where codsituacao<>'D' and p.chapa<'10000' order by p.nome"
rs.Open sqlc, ,adOpenStatic, adLockReadOnly
do while not rs.eof
if rs("chapa")=chapa then txt="selected" else txt=""
if rs("aderencia")>0 then estilo="style='background:CCFFCC;'" else estilo="" 'estilo="style='background:FFFFFF;'"
%>
	<option <%=estilo%>  value="<%=rs("chapa")%>" <%=txt%>><%=rs("nome")%> (<%=rs("aderencia")%>) </option>
<%
rs.movenext
loop
rs.close
%>
</select>
<%
end if
%>

<table border="0" cellpadding="3" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" >
<tr>
	<td class=titulo height=30 valign='middle'>Curso:</td>
	<td class=titulo><select size="1" name="coddoc" onChange="javascript:submit()">
	<option value="" selected>Selecione um curso</option>
<%
sqla="select distinct tpcurso, coddoc, curso, descricao=case tpcurso when 'G' then 'Graduação' when 'L' then 'Cursos Livres' when 'M' then 'Mestrado' when 'P' then 'Pós-Graduação' else '' end " & _
"from grades_vigentes order by tpcurso, curso"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof
if rs("tpcurso")<>grupoanterior then response.write "<option style='background:CCFFCC' value='" & rs("tpcurso") & "'>------- " & ucase(rs("descricao")) & " --------</option>"
%>
	<option <%if request.form("coddoc")=rs("coddoc") then response.write "selected "%> value="<%=rs("coddoc")%>"><%=rs("curso")%></option>
<%
grupoanterior=rs("tpcurso")
rs.movenext
loop
rs.close
%>  
	</select>
	</td>

	<td class=titulo>Grade</td>
	<td class=titulo><select size="1" name="grade" onChange="javascript:submit()" class=a>
	<option value="" selected>Selecione uma grade</option>
<%
sqla="select codpergrade, descricao, codcur, codper, grade from grades_vigentes where coddoc='" & request.form("coddoc") & "' order by codcur, codper, grade"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst : do while not rs.eof 
%>
	<option <%if request.form("grade")=rs("codpergrade") then response.write "selected "%> value="<%=rs("codpergrade")%>"><%=rs("descricao")%></option>
<%
rs.movenext:loop
end if
rs.close
%>  
	</select>
	</td>
</table>
<br>
<input type="hidden" name="acoddoc" value="<%=request.form("coddoc")%>">
<input type="hidden" name="agc" value="<%=request.form("grade")%>">
<%
vezes=1
coddoc=request.form("coddoc")
gc=request.form("grade")
acoddoc=request.form("acoddoc")
if acoddoc<>coddoc and acoddoc<>"" then gc=""
sql="select codcur, codper, grade from grades_vigentes where codpergrade='" & request.form("grade") & "'"
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then codcur=rs("codcur"):codper=rs("codper"):grade=rs("grade")
rs.close
%>
<input type="hidden" name="hcodcur" value="<%=codcur%>">
<input type="hidden" name="hcodper" value="<%=codper%>">
<input type="hidden" name="hgrade" value="<%=grade%>">

<%
'if request.form("coddoc")<>"" and request.form("gc")<>"" then
if coddoc<>"" and gc<>"" then
%>
<ul id="countrytabs" class="shadetabs">
<%
sql="select menor, maior, ultserie from grades_vigentes where coddoc='" & request.form("coddoc") & "' and codpergrade='" & request.form("grade") & "'"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
menor=rs2("menor")
maior=rs2("maior")
ultima=rs2("ultserie")
rs2.close
for a=menor to ultima
%>
<li><a href="#" rel="country<%=a%>"><font style="font-size:7pt">Período</font> <%=a%></a></li>
<%
next
%>
</ul>

<div style="border:1px solid gray; width:650px; height:400px; margin-bottom: 1em; padding: 10px">
<%
for a=menor to ultima
%>
<div id="country<%=a%>" class="tabcontent">
<table border="0" cellpadding="2" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='<%=13.0*pixel%>px' >
<tr>
	<td class=titulop width='<%=8.0*pixel%>px' align='center'>Disciplina</td>
	<td class=titulop width='<%=1.0*pixel%>px' align='center'>C.H.</td>
	<td class=titulor width='<%=1.5*pixel%>px' align='center'>Ultimo<br>semestre</td>
	<td class=titulor width='<%=1.5*pixel%>px' align='center'>Pode<br>lecionar</td>
	<td class=titulop width='<%=1.0*pixel%>px' align='center'>Ver</td>
</tr>
<%
linha=1 :iniserie=0 '*******
formato(0)="style='background-color:gainsboro'"
formato(1)="style='background-color:#FFFFFF'"
sql="select v.coddoc, v.codcur, v.codper, v.grade, v.menor, v.maior, v.codpergrade, u.periodo, u.codmat, m.materia, u.naulassem, u.cargahoraria, ultperlet " & _
", (select chapa from grades_aderencia where codcur=v.codcur and codmat=u.codmat collate database_default and grade=v.grade and codper=v.codper and chapa='" & chapa & "') escolha " & _
"from grades_vigentes v inner join corporerm.dbo.ugrade u on u.codcur=v.codcur and u.codper=v.codper and u.grade=v.grade " & _
"inner join corporerm.dbo.umaterias m on m.codmat=u.codmat " & _
"where coddoc='" & request.form("coddoc") & "' and codpergrade='" & request.form("grade") & "' and periodo=" & a & " "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then 
	do while not rs.eof
%>
<tr>
	<td class=campo <%=style1%> <%=formato(linha)%> ><%=rs("materia")%>
<%
'if teste=2379 then
if (session("usuariomaster")="02379" or session("usuariomaster")="00259" or session("usuariomaster")="00099") and request.form("chapa")="0" then
sqlt="select a.chapa, f.nome, f.codsituacao from grades_aderencia a inner join dc_professor f on f.chapa=a.chapa " & _
"where codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and codmat='" & rs("codmat") & "' and f.codsituacao<>'D'"
rs2.Open sqlt, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then 
	do while not rs2.eof
		if rs2.absoluteposition=1 then response.write "<br>("
		response.write "<font style=""font-size:6pt"">" & rs2("nome")
		if rs2.absoluteposition<rs2.recordcount and rs2.recordcount>1 then response.write ", " else response.write ")"
	rs2.movenext
	loop
end if
rs2.close
end if
%>	
	</td>
	<td class=campo <%=style1%> <%=formato(linha)%> align="center"><%=rs("cargahoraria")%></td>
	<td class=campo <%=style1%> <%=formato(linha)%> align="center"><%=rs("ultperlet")%></td>
	<td class=campo <%=style1%> <%=formato(linha)%> align="center">
	<input onclick="javascript:submit();" type="checkbox" name="m<%=vezes%>" value="<%=rs("codmat")%>" 
 	<%if rs("escolha")=chapa then response.write "checked style='background:" & corcheck &";'"%>
 	title="Marque para informar que está habilitado">
	<input type="hidden" value="<%if rs("escolha")=chapa then response.write rs("codmat")%>" name="a<%=vezes%>">

	</td>
	<td class=campo <%=style1%> <%=formato(linha)%> align="center">
		<a class=r href="aderencia_pe.asp?codcur=<%=rs("codcur")%>&codper=<%=rs("codper")%>&grade=<%=rs("grade")%>&serie=<%=rs("periodo")%>&codmat=<%=rs("codmat")%>&perlet=<%=rs("ultperlet")%>" onclick="NewWindow(this.href,'form_pe','695','550','yes','center');return false" onfocus="this.blur()">
		ver</a>
	</td>
</tr>
<%
	lastper=rs("periodo")
	vezes=vezes+1
	if linha=0 then linha=1 else linha=0
	iniserie=0
	rs.movenext
	loop
end if 'rs.recordcount
rs.close
'response.write "<td class="campop" rowspan=1 colspan=6 align="center" style='border-top:2 solid #000000'></td>"
%>
</table>

</div>
<%
next
%>
</div>

<script type="text/javascript">
var countries=new ddtabcontent("countrytabs")
countries.setpersist(true)
countries.setselectedClassTarget("link") //"link" or "linkparent"
countries.init()
</script>
<!--
<p><a href="javascript:countries.cycleit('prev')" style="margin-right: 25px; text-decoration:none">volta</a> <a href="javascript: countries.expandit(3)">Click here to select last tab</a> <a href="javascript:countries.cycleit('next')" style="margin-left: 25px; text-decoration:none">avança</a></p>
-->

<input type="hidden" name="vezes" value="<%=vezes-1%>">
<!-- -->
<!-- -->
<%
end if 'request.forms <>""

if request.form("chapa")<>"" then
sqllista="select c.nome, p.habilitacao, d.descricao, a.codmat, m.materia " & _
"from grades_aderencia a " & _
"inner join corporerm.dbo.uperiodos p on p.codcur=a.codcur and p.codper=a.codper /*and p.gradeatual=a.grade*/ " & _
"inner join corporerm.dbo.ucursos c on c.codcur=a.codcur " & _
"inner join corporerm.dbo.umaterias m on m.codmat collate database_default=a.codmat " & _
"inner join corporerm.dbo.udefgrade d on d.codcur=a.codcur and d.codper=a.codper and d.grade=a.grade " & _
"where chapa='" & request.form("chapa") & "' order by a.codcur, a.codper, a.grade, materia"
rs.Open sqllista, ,adOpenStatic, adLockReadOnly

'*************** inicio teste **********************
totaldisp=0
response.write "<table border='1' bordercolor='#000000' cellpadding='0' cellspacing='0' style='border-collapse:collapse' width='600'>"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=""titulor"">" & ucase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	if a>2 then alinhamento="center" else alinhamento="left"
	response.write "<td class=""campor"" nowrap align='" & alinhamento & "'>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext:loop
response.write "</table>"
'*************** fim teste **********************
rs.close
end if

'rs.close
'set rs=nothing
'conexao.close
'set conexao=nothing
%>
<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="../images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>