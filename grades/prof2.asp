<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a32")="N" or session("a32")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Indicação de 2º Professor</title>
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
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request("codigo")<>"" or request.form<>"" then
	if request.form="" then indicacao=request("codigo")
	if request("codigo")="" then indicacao=request.form("D1")
	
	sql="SELECT g.id_grade, g.perlet, g.perlet2, g.coddoc, g2cursoeve.CURSO, g.turno, g.serie, g.turma, g.codtur, g.diasem, " & _
	"g.a1, g.a2, g.a3, g.a4, g.a5, g.a6, g.codmat, m.MATERIA, g.chapa1, p1.nome AS prof1, g.chapa2, p2.nome AS prof2, g.codsala, " & _
	"g.alunos, g.obs, g.justificativa " & _
	"FROM (((grades_2 AS g INNER JOIN corporerm.dbo.UMATERIAS AS m ON g.codmat collate database_default = m.CODMAT) INNER JOIN grades_aux_prof AS p1 ON g.chapa1 collate database_default= p1.chapa collate database_default) " & _
	"LEFT JOIN grades_aux_prof AS p2 ON g.chapa2 collate database_default = p2.chapa collate database_default) INNER JOIN g2cursoeve ON g.coddoc = g2cursoeve.coddoc " & _
	"WHERE g.perlet2 Like '2008%2' AND g2cursoeve.coddoc='" & indicacao & "' AND g.prof2=1 " & _
	"ORDER BY g.perlet2, g.coddoc, g.codtur, g.diasem, g.codmat; "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
	titulo=rs("curso")
	end if
	temp=0
	'session("nomeacao_chapa")=rs("chapa")
	'session("nomeacao_id")=""
	'session("nomeacao_descr")=""
else
	temp=1
	'session("nomeacao_chapa")=""
	'session("nomeacao_id")=""
	'session("nomeacao_descr")=""
end if
%>
<%
if temp=1 then
	sqla="SELECT gc.coddoc, gc.CURSO FROM grades_2 AS g INNER JOIN g2cursoeve AS gc ON gc.coddoc = g.coddoc " & _
	"WHERE g.prof2=1 and perlet2 like '2008%2'  GROUP BY gc.coddoc, gc.CURSO ORDER BY gc.CURSO;"
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<p class=titulo>Indicação de 2º Professor &nbsp;<%=titulo %><br>
<form method="POST" action="prof2.asp">
	<p><select size="1" name="D1">
<%
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
%>
	<option value="<%=rs("coddoc")%>"><%=rs("curso")%></option>
<%
rs.movenext:loop
rs.close
else
	response.write "<option value='0'>Sem lançamentos</option>"
end if
%>
	</select>
	<br>
	<input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>
<%
else ' temp=0
%>
<p class=titulo>
Indicação de 2º Professor &nbsp;<%=titulo %><br>

<table border="1" bordercolor="gray" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=titulo align="center">Per.Let.</td>
	<td class=titulo align="center">Curso   </td>
	<td class=titulo align="center">Turno   </td>
	<td class=titulo align="center">Turma   </td>
	<td class=titulo align="center">Dia     </td>
	<td class=titulo align="center">Matéria</td>
	<td class=titulo align="center">Prof.Titular/Indicado</td>
	<td class=titulo align="center"><img border="0" src="../images/Magnify.gif"></td>
</tr>
<%
laststatus=""
inicio=1
Set arquivo=CreateObject("Scripting.FileSystemObject")

if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 
if lastperlet<>rs("perlet2") then
	perlet=rs("perlet"):perlet2=rs("perlet2")
	sql0="SELECT p.curso, p.diretor, p.coordenador, p.adjunto, p.chefedepto, p.secretaria " & _
	"FROM grades_per AS p INNER JOIN g2cursoeve AS c ON p.coddoc = c.coddoc " & _
	"WHERE c.coddoc='" & request.form("D1") & "' AND p.perlet='" & perlet & "' AND p.perlet2='" & perlet2 & "' " & _
	"GROUP BY p.curso, p.diretor, p.coordenador, p.adjunto, p.chefedepto, p.secretaria;"
	rs2.Open sql0, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then coordenador=rs2("coordenador") else coordenador="?????"
	rs2.close

	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="prof2_" & session.sessionid & rs("perlet2") & ".vbs"
	lote=caminho & nomefile
	Set leitura=arquivo.CreateTextFile(lote, true)
	leitura.writeline "Dim objeto, doc"
	leitura.writeline "Set objeto = WScript.CreateObject(""Word.Application"")"
	leitura.writeline "'Set doc=objeto.Documents.Add()"
	leitura.writeline "objeto.Documents.Add()"
	leitura.writeline "Set Doc = objeto.ActiveDocument"
	leitura.writeline "objeto.Visible=True"
	leitura.writeline "texto=""Osasco, "" & day(now) & "" de "" & monthname(month(now)) & "" de "" & year(now) & vbcrlf"
	leitura.writeline "texto=texto & vbcrlf & vbcrlf"
	leitura.writeline "texto=texto & ""Magnífico Pró-Reitor Acadêmico"" & vbcrlf"
	leitura.writeline "texto=texto & ""Dr. Luiz Fernando da Costa e Silva"" & vbcrlf"
	leitura.writeline "texto=texto & vbcrlf & vbcrlf & vbcrlf & vbcrlf"
	leitura.writeline "texto=texto & ""Ref. Inclusão de Professor Auxiliar"" & vbcrlf"
	leitura.writeline "texto=texto & vbcrlf & vbcrlf"
	leitura.writeline "texto=texto & ""Conforme as justificativas abaixo discriminadas, solicito autorização para a inclusão """
	leitura.writeline "texto=texto & ""de professor auxiliar nas disciplinas abaixo para o curso de " & rs("curso") & ", durante o período """
	leitura.writeline "texto=texto & ""letivo de " & rs("perlet") & "."""
	leitura.writeline "texto=texto & vbcrlf & vbcrlf"

	response.write "<tr><td class=""campol"" colspan='10'>"
	response.write "&nbsp;<a href='../temp/" & nomefile & "'><img src='../images/printer.gif' border=0>" & rs("perlet") & "</a></td></tr>"
end if
%>
<tr>
	<td class=campo><%=rs("perlet") %></td>
	<td class=campo><%=rs("coddoc")%>-<%=rs("curso")%></td>
	<td class=campo align="center"><%=rs("turno") %></td>
	<td class=campo align="center"><%=rs("codtur") %></td>
	<td class=campo align="center"><%=weekdayname(rs("diasem"),1) %></td>
	<td class=campo><%=rs("codmat")%>-<%=rs("materia")%></td>
	<td class=campo><%=rs("chapa1")%>-<%=rs("prof1")%></td>
	<td class=campo align="center" rowspan=2>
	<% if session("a32")="T" then %>
		<a href="grade_alteracao.asp?codigo=<%=rs("id_grade")%>" onclick="NewWindow(this.href,'AlteracaoGrade','550','425','no','center');return false" onfocus="this.blur()">
		<img src="../images/Folder95O.gif" border="0"></a>
	<% end if %>
	</td>
</tr>
<tr>
	<td class="campor" colspan=6>Justificativa: <%=rs("justificativa")%></td>
	<td class=campo><%=rs("chapa2")%>-<%=rs("prof2")%></td>
</tr>
<tr>
	<td class="campoa"r colspan=8 height=5></td>
</tr>
<%
	leitura.writeline "texto=texto & ""Matéria: " & rs("materia") & " - Turma: " & rs("codtur") & " - Dia: " & weekdayname(rs("diasem"),1) & " - nº alunos: " & rs("alunos") & """ & vbcrlf"
	leitura.writeline "texto=texto & ""Professor titular: " & rs("prof1") & """ & vbcrlf"
	leitura.writeline "texto=texto & ""Professor auxiliar: " & rs("prof2") & """ & vbcrlf"
	leitura.writeline "texto=texto & ""Justificativa: " & rs("justificativa") & """ & vbcrlf & vbcrlf"
lastperlet=rs("perlet2")
inicio=0
rs.movenext

if not rs.eof then
	if lastperlet<>rs("perlet2") then
	leitura.writeline "texto=texto & ""Atenciosamente."" & vbcrlf"
	leitura.writeline "texto=texto & vbcrlf & vbcrlf"
	leitura.writeline "texto=texto & vbcrlf & vbcrlf"
	leitura.writeline "texto=texto & vbcrlf & vbcrlf"
	leitura.writeline "texto=texto & """ & coordenador & """ & vbcrlf"
	leitura.writeline "texto=texto & ""Coordenador do curso de " & titulo & """ & vbcrlf"
	leitura.writeline "doc.Content.Text = texto"
	leitura.close
	end if
end if

loop
rs.close

	leitura.writeline "texto=texto & ""Atenciosamente."" & vbcrlf"
	leitura.writeline "texto=texto & vbcrlf & vbcrlf"
	leitura.writeline "texto=texto & vbcrlf & vbcrlf"
	leitura.writeline "texto=texto & vbcrlf & vbcrlf"
	leitura.writeline "texto=texto & """ & coordenador & """ & vbcrlf"
	leitura.writeline "texto=texto & ""Coordenador do curso de " & titulo & """ & vbcrlf"
	leitura.writeline "doc.Content.Text = texto"
leitura.close
end if
%>
</table>

<%
end if 'rs.rec

set leitura=nothing
set arquivo=nothing
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>