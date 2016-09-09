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
chapa=session("usuariomaster")
corcheck="black"
'sqla="SELECT dc_carga.CURSO FROM dc_carga GROUP BY dc_carga.CURSO;"
'rs.Open sql1, ,adOpenStatic, adLockReadOnly

if request.form="" then

end if ' request.form=""

if request.form<>"" then
	'response.write request.form
	vez=request.form("vezes")
	for a=1 to vez
		info    =request.form("m" & a)
		anterior=request.form("a" & a)
		'response.write info & " " & anterior & " " & "<br>"
		if info<>"" and anterior="" then
			sql="insert into grades_aderencia (chapa, coddoc, codmat, gc) " & _
			"select '" & chapa & "', '" & request.form("coddoc") & "','" & info & "','" & request.form("gc") & "';"
			conexao.execute sql
		end if
		if info="" and anterior<>"" then
			sql="delete from grades_aderencia where chapa='" & chapa & "' and coddoc='" & request.form("coddoc") & "' and gc='" & request.form("gc") & "' " & _
			"and codmat='" & anterior & "' "
			response.write sql
			conexao.execute sql
		end if
	next
end if 'request.form<>""

sql6=""
'rs.Open sql6, ,adOpenStatic, adLockReadOnly
'do while not rs.eof
'rs.movenext
'loop
'rs.close

%>
<!-- -->
<!-- -->
<table><tr><td valign=top style="border-right:3px double #000000" width=150>
<!-- -->
<p style="margin-top:0;margin-bottom:0" class=titulo><%=session("usuarioname")%></p>
<hr>
<p style="margin-top:0;margin-bottom:5"><a href="../academic/disponibilidade.asp">
<img src="../images/Clock.gif" width="16" height="16" border="0" alt="">Disponibilidade</a></p>
<p style="margin-top:0;margin-bottom:5">
<img src="../images/BookO.gif" width="16" height="16" border="0" alt="">Aderência</p>
<!-- -->
</td><td valign=top>
<!-- -->

<p style="margin-top:0;margin-bottom:10" class=titulo>Aderência à Grade Curricular</p>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" >
<form method="POST" action="aderencia.asp" name="form">
<tr>
	<td class=titulo height=30 valign='middle'>Curso:</td>
	<td class=titulo><select size="1" name="coddoc" onChange="javascript:submit()">
	<option value="" selected>Selecione um curso</option>
<%sqla="select p.coddoc, c.curso from grades_plano p, g2cursoeve c where left(p.perlet,4) in (year(getdate()),year(getdate())-1) and p.coddoc=c.coddoc group by p.coddoc, c.curso order by p.coddoc":rs.Open sqla, ,adOpenStatic, adLockReadOnly:rs.movefirst:do while not rs.eof %>
	<option <%if request.form("coddoc")=rs("coddoc") then response.write "selected "%> value="<%=rs("coddoc")%>"><%=rs("curso")%></option>
<%rs.movenext:loop:rs.close%>  
	</select>
	</td>

	<td class=titulo>Grade</td>
	<td class=titulo><select size="1" name="gc" onChange="javascript:submit()">
	<option value="" selected>Selecione uma grade</option>
<%
sqla="select p.coddoc, p.gc from grades_plano p where left(p.perlet,4) in (year(getdate()),year(getdate())-1) " & _
"and p.coddoc='" & request.form("coddoc") & "' group by p.coddoc, p.gc order by p.coddoc"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst : do while not rs.eof 
%>
	<option <%if request.form("gc")=rs("gc") then response.write "selected "%> value="<%=rs("gc")%>"><%=rs("gc")%></option>
<%
rs.movenext:loop
end if
rs.close
%>  
	</select>
	</td>
</table>
<hr>
<%

%>
<table border="0" cellpadding="2" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" >
<tr>
	<td class=titulop>Período</td>
	<td class=titulop>Disciplina</td>
	<td class=titulop title="Carga horária semestral ou anual da disciplina">C.H.</td>
	<td class=titulor title="Último semestre que a disciplina foi ministrada">Ultimo<br>semestre</td>
	<td class=titulor title="Indicar se pode ou não ministrar esta disciplina">Pode<br>lecionar</td>
	<td class=titulop title="Visualizar o plano de ensino quando disponível">Ver</td>
</tr>
<%
if request.form("coddoc")<>"" and request.form("gc")<>"" then
linha=1 :iniserie=0:vezes=1 '*******
formato(0)="style='background-color:gainsboro'"
formato(1)="style='background-color:#FFFFFF'"
sql="select distinct top 100 percent m.coddoc, m.gc, m.serie, m.codmat, m.materia, m.naulassem, m.cargahoraria, max(p.perlet) ultperlet " & _
"from grades_materias m left join grades_plano p " & _
"on p.coddoc=m.coddoc and p.gc=m.gc and p.serie=m.serie and p.codmat=m.codmat " & _
"where m.coddoc='" & request.form("coddoc") & "' and m.gc='" & request.form("gc") & "'" & _
"group by m.coddoc, m.gc, m.serie, m.codmat, m.materia, m.naulassem, m.cargahoraria " & _
"order by m.serie, m.codmat"
sql="select distinct top 100 percent p.coddoc, p.gc, p.serie, p.codmat, p.materia, p.naulassem, p.cargahoraria, p.ultperlet " & _
", (select chapa from grades_aderencia where coddoc=p.coddoc and codmat=p.codmat and gc=p.gc and chapa='" & chapa & "') escolha " & _
"from grades_pe p " & _
"where p.coddoc='" & request.form("coddoc") & "' and p.gc='" & request.form("gc") & "' " & _
"order by p.serie, p.codmat "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then 
do while not rs.eof
%>
<tr>
<%
if lastper<>rs("serie") then
	sqlg="select count(serie) as linhas from (" & sql & ") t where serie=" & rs("serie")
	rs2.Open sqlg, ,adOpenStatic, adLockReadOnly
	lin_per=rs2("linhas")
	rs2.close
	barra1=" style='border-top:2 solid #000000'"
	response.write "<td class="campop" rowspan=" & lin_per & " align="center" style='border-top:2 solid #000000'><b>" & rs("serie") & "</td>"
	iniserie=1
end if
if iniserie=1 then style1="style='border-top:2 solid #000000;'" else style1=""
%>
	<td class=campo <%=style1%> <%=formato(linha)%> ><%=rs("materia")%></td>
	<td class=campo <%=style1%> <%=formato(linha)%> align="center"><%=rs("cargahoraria")%></td>
	<td class=campo <%=style1%> <%=formato(linha)%> align="center"><%=rs("ultperlet")%></td>
	<td class=campo <%=style1%> <%=formato(linha)%> align="center">
	<input onclick="javascript:submit();" type="checkbox" name="m<%=vezes%>" value="<%=rs("codmat")%>" 
 	<%if rs("escolha")=chapa then response.write "checked style='background:" & corcheck &";'"%>
 	title="Marque para informar que está habilitado">
	<input type="hidden" value="<%if rs("escolha")=chapa then response.write rs("codmat")%>" name="a<%=vezes%>">

	</td>
	<td class=campo <%=style1%> <%=formato(linha)%> align="center">
		<a class=r href="aderencia_pe.asp?coddoc=<%=rs("coddoc")%>&codmat=<%=rs("codmat")%>&gc=<%=rs("gc")%>&serie=<%=rs("serie")%>&perlet=<%=rs("ultperlet")%>" onclick="NewWindow(this.href,'form_pe','695','550','yes','center');return false" onfocus="this.blur()">
		ver</a>
	</td>
</tr>
<%
lastper=rs("serie")
vezes=vezes+1
if linha=0 then linha=1 else linha=0
iniserie=0
rs.movenext
loop
end if 'rs.recordcount
rs.close
%>








<%
	response.write "<td class="campop" rowspan=1 colspan=6 align="center" style='border-top:2 solid #000000'></td>"
%>
</table>
<input type="hidden" name="vezes" value="<%=vezes-1%>">


<!-- -->
<!-- -->
<%
end if 'request.forms <>""
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