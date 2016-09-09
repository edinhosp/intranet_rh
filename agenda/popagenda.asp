<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a96")="N" or session("a96")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Agenda de Compromissos e Lembretes</title>
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
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
usuario=session("usuariomaster")
if usuario="02379" or usuario="00892" or usuario="00259" or usuario="02159" or usuario="00595" or usuario="02463" or usuario="90261" or usuario="90344" then fotos=1 else fotos=0

dim emes(12),edia(31),fdia(40),edia2(31),edia3(31),fdia2(40),fdia3(40)
mesagora=month(now)
anoagora=year(now)
diaagora=day(now)

datasql=dateserial(anoagora,mesagora,diaagora)
incremento=0
datasql2=dateserial(anoagora,mesagora,diaagora+incremento)

sqla="Select * from ( " 
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u where u.usuario=a.usuarioc and a.tipo=0 and a.usuarioc='" & session("usuariomaster") & "' and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u where u.usuario=a.usuarioc and a.tipo=2 and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u where u.usuario=a.usuarioc and a.tipo=3 and a.usuarioc='" & session("usuariomaster") & "' and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u, agenda_3 a3 where a3.id_agenda=a.id_agenda and u.usuario=a.usuarioc and a.tipo=3 and a3.codigo='" & session("usuariomaster") & "' and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' "
sqla=sqla & "union all "
sqla=sqla & "SELECT a.*, u.nome from agenda a, usuarios u, agenda_1 a1 where a1.id_agenda=a.id_agenda and u.usuario=a.usuarioc and a.tipo=1 and a.data between '" & dtaccess(datasql) & "' and '" & dtaccess(datasql2) & "' and a1.codigo in (select grupo from usuarios where usuario='" & session ("usuariomaster") & "') "
sqla=sqla & ") as s order by data, hora "
'response.write sqla
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="400">
<tr><td colspan=4 class="campol" align="center"><font size="2">&nbsp;Lembretes/Compromissos do dia <b><%=datasql%>&nbsp;</font></td></tr>
<tr><td class=fundo align="center"></td>
	<td class=fundo align="center" valign=middle width=35>Hora</td>
	<td class=fundo align="center" valign=middle width=295>Compromisso</td>
	<td class=fundo align="center" valign=middle width=70>Anotação<br>criada por:</td>
</tr>
<%
if rs.recordcount>0 then
	rs.movefirst:do while not rs.eof 
%>
<tr><td class=campo align="center" height=16>
	<img src="../images/LeafSearch.gif" width="16" height="16" border="0" alt="">
	</td>
	<td class=campo align="center"><%if rs("hora")<>"" then response.write formatdatetime(rs("hora"),4) else response.write "-"%></td>
	<td class=campo>&nbsp;<%=rs("compromisso")%></td>
	<td class=campo>&nbsp;<%=rs("nome")%></td></tr>
<%
rs.movenext:loop
else
	response.write "<tr><td colspan='4'><font size='2'>&nbsp;Não há compromissos no dia de hoje.&nbsp;&nbsp;</font></td></tr>"
end if 'if recordcount
%>
</table>
<%
rs.close

set rs=nothing
conexao.close
set conexao=nothing
%>

</body>
</html>