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
'response.write "<br>" & session("acesso")
'response.write "<br>" & session("usuariomaster")
'response.write "<br>" & session("usuariogrupo")

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
largura=600
pixel=96/2.54
point=72/2.54
pointp=72.27/2.54

if request.form="" then
end if ' request.form=""

if request.form<>"" then
end if 'request.form<>""

%>
<!-- -->
<!-- -->
<form method="POST" action="meusplanos.asp" name="form">
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

<p style="margin-top:0;margin-bottom:5"><a href="../academic/aderencia.asp">
<img src="../images/BookO.gif" width="16" height="16" border="0" alt="">Aderência</a></p>

<br><br>
<p style="margin-top:0;margin-bottom:5">
<img src="../images/BookO.gif" width="16" height="16" border="0" alt="">Plano de Ensino

<br><br>
<p style="margin-top:0;margin-bottom:5"><a href="../academic/espelho.asp">
<img src="../images/espelho.jpg" width="16" height="16" border="0" alt="">Marcação de Ponto</a></p>

<br><br><br><br><br><br><br>
<p style="margin-top:0;margin-bottom:0"><a href="../indexp.asp">
<img src="../images/setafirst0.gif" width="12" height="12" border="0" alt="">Início</a>
<!-- -->
</td><td valign=top style="border-bottom:3px double silver">
<p style="margin-top:0;margin-bottom:10" class=titulo>Planos de Ensino</p>
<!-- -->
<%
else ' para acesso=1
%>
<p style="margin-top:0;margin-bottom:10" class=titulo>Planos de Ensino</p>
<select size=1 name="chapa" onchange="javascript:submit();">
	<option value="0">Selecione....</option>
<%
sqlc="select chapa, nome from grades_aux_prof where codsituacao<>'D' and chapa<'10000' order by nome"
rs.Open sqlc, ,adOpenStatic, adLockReadOnly
do while not rs.eof
if rs("chapa")=chapa then txt="selected" else txt=""
%>
	<option value="<%=rs("chapa")%>" <%=txt%>><%=rs("nome")%></option>
<%
rs.movenext
loop
rs.close
%>
</select>
<%
end if
%>

<%
sql0="insert into grades_plano (coddoc, codcur, codper, grade, perlet, serie, codmat, pa, novo, prof, coord) " & _
"select distinct g.coddoc, g.codcur, g.codper, g.grade, g.perlet, g.serie, g.codmat, pa=0, novo=1,0,0 " & _
"from g2ch g left join grades_plano p on p.coddoc=g.coddoc and p.codcur=g.codcur and p.codper=g.codper and p.grade=g.grade " & _
"	and p.codmat=g.codmat and p.perlet=g.perlet and p.serie=g.serie " & _
"where g.chapa1='" & chapa & "' and getdate() between g.inicio and g.termino and p.id_plano is null "
conexao.execute sql0

'*************** inicio teste **********************
sql1="select distinct g.coddoc, g.codcur, g.codper, g.grade, u.habilitacao, g.perlet, g.serie, g.codmat, g.materia, p.id_plano, pa, novo, p.prof, coord " & _
"from g2ch g " & _
"left join corporerm.dbo.uperiodos u on u.codcur=g.codcur and u.codper=g.codper " & _
"left join grades_plano p on p.coddoc=g.coddoc and p.codcur=g.codcur and p.codper=g.codper and p.grade=g.grade " & _
"	and p.codmat=g.codmat and p.perlet=g.perlet and p.serie=g.serie " & _
"where g.chapa1='" & chapa & "' and convert(datetime,convert(integer,getdate())) between g.inicio and g.termino "
rs.Open sql1, ,adOpenStatic, adLockReadOnly

'if session("usuariomaster")="02379" then
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
'next
'response.write "</tr>"
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext:loop
'response.write "</table>"
'response.write "# " & rs.recordcount & "<br>"
'end if
'*************** fim teste **********************
%>

<table border="1" cellpadding="4" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width="<%=largura%>">
<tr>
	<td colspan=5 class=titulop align="center" style="border-bottom:2px dotted #000000">Disciplinas ministradas em <%=formatdatetime(now(),2)%></td>
</tr>
<tr>
	<td class=titulo width="240">Curso</td>
	<td class=titulo width="30">Per.Let.</td>
	<td class=titulo width="30">Sem.</td>
	<td class=titulo width="280">Disciplina</td>
	<td class=titulo width="40"></td>
</tr>	
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class="campor"><%=rs("habilitacao")%></td>
	<td class=campo align="center"><%=rs("perlet")%></td>
	<td class=campo align="center"><%=rs("serie")%></td>
	<td class=campo><%=rs("materia")%></td>
	<td class=campo>
<%if rs("pa")=false or rs("coord")=false or isnull(rs("pa")) then%>
		<a href="plano_alteracao.asp?codigo=<%=rs("id_plano")%>" onclick="NewWindow(this.href,'planoensino_altera','635','580','yes','center');return false" onfocus="this.blur()">
		<img src="../images/novo.gif" width="17" height="17" border="0" alt="Alterar este plano de ensino"></a>
<%else%>
	<img src="../images/Stop.gif" width="16" height="16" border="0" alt="Alterações não permitidas.">
<%end if%>
<%if rs("novo")=false or rs("pa")=true or rs("coord")=true then%>
	<a class=r href="plano_ensino.asp?codigo=<%=rs("id_plano")%>" onclick="NewWindow(this.href,'form_pe','695','450','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Leaf.gif" width="16" height="16" border="0" alt="Visualizar este plano de ensino"></a>
<%end if%>
	</td>
</tr>	
<%
rs.movenext
loop
end if 'rs.recordcount>0
rs.close
%>

</table>


<!-- -->


<!-- -->


<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="../images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>