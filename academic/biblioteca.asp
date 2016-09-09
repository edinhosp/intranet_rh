<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a93")="N" or session("a93")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Biblioteca - Controle de Aquisições</title>
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

'	sqla="SELECT chapa, nome from n_indicacoes group by chapa, nome order by nome"
'	rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<form method="POST" action="biblioteca.asp">
<p class=titulo>Biblioteca - Controle de Aquisições:&nbsp;
	<select size="1" name="tipo" onchange="javascript:submit()">
		<option value="P" <%if request.form("tipo")="P" then response.write "selected"%> >Aquisições solicitadas (através do plano de ensino)</option>
		<option value="A" <%if request.form("tipo")="A" then response.write "selected"%> >Aquisições autorizadas (através da biblioteca)</option>
		<option value="C" <%if request.form("tipo")="C" then response.write "selected"%> >Aquisições efetuadas (através da biblioteca)</option>
		<option value="N" <%if request.form("tipo")="N" then response.write "selected"%> >Aquisições negadas</option>
	</select>
	<input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>
<%

if request.form("tipo")<>"" then 'novo

if request.form="" then status=request("codigo")
if request("codigo")="" then status=request.form("D1")
status=request.form("tipo")
	
sql0="select top 100 referencia=convert(nvarchar(255),referencia), status=min(status), complementar, quant=count(complementar) " & _
"from grades_plano_biblio " & _
"where status='" & status & "' " & _
"group by convert(nvarchar(255),referencia), complementar "
rs.Open sql0, ,adOpenStatic, adLockReadOnly
temp=0
select case status
	case "P"
		titulo="Aquisições solicitadas"
	case "A"
		titulo="Aquisições autorizadas"
	case "C"
		titulo="Aquisições efetuadas"
	case "N"
		titulo="Aquisições negadas"
end select

%>
<p class=titulo style="margin-bottom:0px;margin-top:0px"><%=titulo %><br>

<table border="0" bordercolor="gray" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=titulo align="center" style="border:1px solid">Referência bibliográfica</td>
	<td class=titulo align="center" style="border:1px solid">Status</td>
	<td class=titulo align="center" style="border:1px solid">Bibliografia</td>
	<td class=titulo align="center" style="border:1px solid">Solicitações</td>
	<td class=titulo align="center" style="border:1px solid"><img border="0" src="../images/Magnify.gif"></td>
</tr>
<%
laststatus=""
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 
if laststatus<>rs("status") then
	'response.write "<tr><td class=grupo colspan='10'>"
	'response.write "&nbsp;" & rs("status") & "</td></tr>"
end if
%>
<tr>
	<td class=campo style="border:1px solid;border-top:2px solid #000000"><b><%=rs("referencia") %></td>
	<td class=campo style="border:1px solid;border-top:2px solid #000000" align="center"><%=rs("status") %></td>
	<td class=campo style="border:1px solid;border-top:2px solid #000000"><%if rs("complementar")=1 then response.write "complementar"%></td>
	<td class=campo style="border:1px solid;border-top:2px solid #000000" align="center"><%=rs("quant") %></td>
	<td class=campo style="border:1px solid;border-top:2px solid #000000" align="center"></td>
</tr>	
<%
sql1="select id_biblio, b.id_plano, b.usuarioc, f.nome nomec, b.usuarioa, f2.nome nomea, p.coddoc, p.codcur, u.habilitacao, p.codper, p.grade, p.codmat, m.materia, p.perlet " & _
"from grades_plano_biblio b " & _
"inner join grades_plano p on p.id_plano=b.id_plano " & _
"inner join corporerm.dbo.umaterias m on m.codmat collate database_default=p.codmat " & _
"inner join corporerm.dbo.uperiodos u on u.codcur=p.codcur and u.codper=p.codper " & _
"left join corporerm.dbo.pfunc f on f.chapa collate database_default=b.usuarioc " & _
"left join corporerm.dbo.pfunc f2 on f2.chapa collate database_default=b.usuarioa " & _
"where referencia like '" & rs("referencia") & "' and status='" & status & "' "
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
do while not rs2.eof
%>
<tr>
	<td class=campo colspan=4 style="border-left:1px solid">
<%
	response.write "<p style='font-size:8pt;margin-left:15px;text-indent:-8px'>•&nbsp;para a disciplina <b><i>" & rs2("materia") & "</i></b> do curso <font size=1><i>" & rs2("habilitacao") & "</i></font> pelo professor <u><i>" & rs2("nomec") & "</i></u>"
%>
	</td>
	<td class=campo align="center" style="border-left:1px solid;border-right:1px solid">
	<% if session("a93")="T" then %>
		<a href="biblioteca_status.asp?codigo=<%=rs2("id_biblio")%>" onclick="NewWindow(this.href,'AlteracaoAquisicao','520','430','yes','center');return false" onfocus="this.blur()">
		<img border='0' src='../images/folder95o.gif'></a>
	<% end if %>
	</td>
</tr>
<%
rs2.movenext
loop
end if 'rs2.recordcount
rs2.close

laststatus=rs("status")
rs.movenext
loop
response.write "<tr><td class=campo colspan=5 style='border-top:1px solid' height=5></td></tr>"
end if 'rs.recordcount>0
rs.close
end if
%>
</table>

<%
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>