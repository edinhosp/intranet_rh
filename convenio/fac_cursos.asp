<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a55")="N" or session("a55")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cursos oferecidos pelo Convênio</title>
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

if request("acao")="excluir" then
	curso=Request.QueryString("curso")
	sql="delete from rhconveniobec where id_curso=" & curso
	conexao.execute sql
	manutencao=1
end if

if request.form("B2")<>"" then
	iCount=request("Count")
	for iLoop=0 to iCount
		aid=request("id" & iLoop)
		acurso=request("curso" & iLoop)
		strSql="Update rhconveniobec Set cursos = '" & acurso & "' Where id_curso=" & aid
		conexao.execute strSql, , adCmdText
	next
	if request.form("novocurso")<>"" then
		sSql="Insert Into rhconveniobec (id_faculdade, cursos ) "
		sSql=sSql & "Values (" & session("idcurso") & ", '" & request.form("novocurso") & "'"
		sSql=sSql & ")"
		conexao.Execute sSQL, , adCmdText
	end if
	manutencao=1
end if

'if manutencao=1 then response.redirect "rhcursos.asp?codigo=" & session("idcurso")
if manutencao=1 then response.write "<script>javascript:parent.frmMain.location.href='fac_cursos.asp?codigo=" & session("idcurso") & "';</script>"
if request("codigo")<>"" or request.form("B1")<>"" then
	if request.form="" then idfaculdade=request("codigo")
	if request("codigo")="" then idfaculdade=request.form("D1")
	sqla="SELECT id_faculdade, faculdade FROM rhconveniobe where id_faculdade=" & idfaculdade

	set rs2=server.createobject ("ADODB.Recordset")
	Set rs2 = conexao.Execute (sqla, , adCmdText)
	faculdade=rs2("faculdade")

	sqlc="SELECT id_faculdade, id_curso, cursos " & _
	"FROM rhconveniobec "

	sqld=" where id_faculdade=" & idfaculdade
	sqle=" ORDER BY cursos "
	sqlb=sqlc & sqld & sqle
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
	temp=0
	manutencao=0
	session("idcurso")=idfaculdade
else
	temp=1
end if
%>

<p class=titulo>Cursos oferecidos -&nbsp;<%=faculdade%>
<%
if temp=1 then
	sqla="SELECT id_faculdade, faculdade FROM rhconveniobe order by faculdade "
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<form method="POST" name="form1" action="fac_cursos.asp">
  <p><select size="1" name="D1" style="font-size: 8 pt">
<%
rs.movefirst
do while not rs.eof
%>
  <option value="<%=rs("id_faculdade")%>"><%=rs("faculdade")%></option>
<%
rs.movenext
loop
rs.close
%>
  </select>
  <br>
  <input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>
<%
else ' temp=0
%>
<form method="POST" name="form2" action="fac_cursos.asp">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo align="center">Curso</td>
	<td class=titulo align="center">&nbsp;</td>
</tr>
<%
laststatus=""
if rs.recordcount>0 then
tcount=0
rs.movefirst
do while not rs.eof 
%>
<tr>
	<td>
	<input type="hidden" name="id<%=tcount%>" value="<%=rs("id_curso") %>">
	<input type="text" class="form_input" name="curso<%=tcount%>" size="80" value="<%=rs("cursos") %>">
	</td>
	<td align="center">
	<% if session("a55")="T" then %>
		<a href="fac_cursos.asp?acao=excluir&curso=<%=rs("id_curso")%>">
		<img border="0" src="../images/Trash.gif"></a>
	<% end if %>
	</td>
</tr>
<%
rs.movenext
tcount=tcount+1
loop
rs.close
else 'recordcount=0
	response.write "<tr><td colspan='9' class=grupo>"
	response.write "<p>Não há cursos para esta instituição</td></tr>"
end if 'recordcount=0

if session("a55")="T" and temp=0 then
%>
  <tr><td colspan=2><input type="text" class="form_input" name="novocurso" size="80"></td></tr>
<% end if 'grant %>
</table>
  <input type="hidden" name="Count" value="<%=tcount-1%>">
<%if session("a55")="T" and temp=0 then %>
  <input type="submit" value="Salvar" class="button" name="B2">
<% end if 'grant %>
</form>
<%
end if 'temp=0
%>
</body>

</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>