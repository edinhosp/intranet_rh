<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a21")="N" or session("a21")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Nomeados</title>
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
set rsRHT=server.createobject ("ADODB.Recordset")
Set rsRHT.ActiveConnection = conexao

if request("codigo")<>"" or request.form<>"" then
	if request.form="" then idnomeacao=request("codigo")
	if request("codigo")="" then idnomeacao=request.form("D1")
	sqla="SELECT id_nomeacao, nomeacao, criacao FROM n_nomeacoes where id_nomeacao=" & idnomeacao

	set rs2=server.createobject ("ADODB.Recordset")
	Set rs2 = conexao.Execute (sqla, , adCmdText)
	Nomeacao=rs2("nomeacao")
	'session("nomeacao_id")   =rs2("id_nomeacao")
	'session("nomeacao_descr")=rs2("nomeacao")
	'session("nomeacao_chapa")=""

	if request("status")="" then status="A" else status=request("status")
	sqla="SELECT Status=case when mand_fim<getdate() then 'Vencidas' else 'Ativas' end, " & _
	"id_nomeacao, PORTARIA, id_indicado, CHAPA, NOME, CARGO, coddoc, " & _
	"codeve, MAND_INI, MAND_FIM, alunos, CH, OBS, CONTRATO, entrega " & _
	"FROM n_indicacoes " & _
	"where id_nomeacao=" & idnomeacao & " " & _
	"and (case when mand_fim<getdate() then 'V' else 'A' end)='" & status & "' " & _
	"ORDER BY (case when mand_fim<getdate() then 'Vencidas' else 'Ativas' end), n_indicacoes.NOME, mand_ini desc, mand_fim "
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
	temp=0 'request by nomeacoes_tipo
else
	temp=1 'request by self
	'session("nomeacao_chapa")=""
	'session("nomeacao_id")=""
	'session("nomeacao_descr")=""
end if
%>

<%
if temp=1 then
%>
<p><b>Nomeações para&nbsp;<%=nomeacao%></b></p>
<form method="POST" action="nomeados.asp">
	<select size="1" name="D1">
<%
	sqla="SELECT id_nomeacao, nomeacao, criacao FROM n_nomeacoes order by nomeacao"
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
	rs.movefirst:do while not rs.eof
%>
	<option value="<%=rs("id_nomeacao")%>"><%=rs("nomeacao")%></option>
<%
	rs.movenext:loop
	rs.close
%>
	</select><br>
	<input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>
<%
else ' temp=0
aba0="border-top:2pt double #000000;border-left:2pt double #000000;border-right:3pt double #000000"
aba1="border-top:2pt solid #000000;border-left:2pt solid #000000;border-right:3pt solid #000000"
if status="A" then abaa=aba1 else abaa=aba0
if status="V" then abav=aba1 else abav=aba0
%>
<br>
<table cellpadding="2" cellspacing="0" width="690">
<tr>
	<td class=campo><b>
	<% if session("a21")="T"then %>
	<a href="nomeados_nova.asp?codigo=<%=idnomeacao%>" onclick="NewWindow(this.href,'Inclusao','520','330','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/Appointment.gif"></a>
	<% end if %>
	Nomeações para&nbsp;<%=nomeacao%></b></td>
	<td class=campo width="5">&nbsp;</td>
	<td class=campo width="70" align="center" style="<%=abaa%>">
	<a href="nomeados.asp?codigo=<%=idnomeacao%>&status=A">
	Ativas</a></td>
	<td class=campo width="5">&nbsp;</td>
	<td class=campo width="70" align="center" style="<%=abav%>">
	<a href="nomeados.asp?codigo=<%=idnomeacao%>&status=V">
	Vencidas</a></td>
	<td class=campo width="5">&nbsp;</td>
</tr>
</table>

<table border="1" bordercolor="gray" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=titulor align="center">Portaria</td>
	<td class=titulor align="center">Nome/Docente</td>
	<td class=titulor align="center">Cargo</td>
	<td class=titulor align="center">Folha</td>
	<td class=titulor align="center">Inicio em</td>
	<td class=titulor align="center">Término</td>
	<td class=titulor align="center">C.H.</td>
	<td class=titulor align="center"><img border="0" src="../images/Magnify.gif"></td>
	<td class=titulor align="center">Contrato</td>
	<td class=titulor align="center">Dev.Contr</td>
</tr>
<%
laststatus=""
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof 
if laststatus<>rs("status") then
	response.write "<tr><td class=grupo colspan='10'>"
	response.write "&nbsp;" & rs("status")
	response.write "</td></tr>"
end if
%>
<tr>
  	<%if rs("obs")<>"" then response.write "<td class=""campor"" rowspan=""2"">" else response.write "<td class=""campor"">" %>
    <%=rs("portaria") %></td>
  	<%if rs("obs")<>"" then response.write "<td class=""campor"" rowspan=""2"">" else response.write "<td class=""campor"">" %>
    <%=rs("nome") & " (" & rs("chapa") & ")" %></td>
	<td class="campor"><%=rs("cargo") %></td>
	<td class="campor" align="center"><%=rs("codeve")%><br><%=rs("coddoc")%></td>
	<td class="campor" align="center"><%=rs("mand_ini") %></td>
	<td class="campor" align="center"><%=rs("mand_fim") %></td>
	<td class="campor" align="center"><%=rs("ch") %></td>
	<td class="campor" align="center">
    <% if session("a21")="T" then %>
		<a href="nomeados_alteracao.asp?codigo=<%=rs("id_indicado")%>" onclick="NewWindow(this.href,'AlteracaoNomeados','520','330','no','center');return false" onfocus="this.blur()">
		<img border='0' src='../images/folder95o.gif'></a>
	<% end if %>
	</td>

	<td class="campor" align="center">
<%
if session("a21")="T" then
	'*****************************************************************************************
	sqlrht="select chapa from quem_nomeacoes where tipo in ('RHT','RT') and chapa='" & rs("chapa") & "'"
	rsrht.Open sqlrht, ,adOpenStatic, adLockReadOnly
	if rsrht.recordcount>0 then contrht=1 else contrht=0
	rsrht.close
	contrht=0
	'*****************************************************************************************
	if idnomeacao=84 then
%>
		<a href="contrato_rht.asp?codigo=<%=rs("id_indicado")%>" onclick="NewWindow(this.href,'ContratoNomeados','690','400','yes','center');return false" onfocus="this.blur()">
	<%else
		if idnomeacao=12 or contrht=0 then%>
		<a href="nomeados_contrato.asp?codigo=<%=rs("id_indicado")%>" onclick="NewWindow(this.href,'ContratoNomeados','690','400','yes','center');return false" onfocus="this.blur()">
	<%	
	end if
	end if%>		
<%
end if
if rs("contrato")<>"" then 
	response.write "<font size='1'>" &rs("contrato")
else
	if (session("a21")="T" and contrht=0) or idnomeacao=84 or idnomeacao=12 then
		response.write "<img border='0' src='../images/novo.gif'>"
	else 
		response.write "&nbsp;"
	end if
end if
if session("a21")="T" then response.write "</a>" 
%>
	</td>
	<td class="campor" align="center"><%=rs("entrega") %></td>
</tr>
<%
if rs("obs")<>"" then
  response.write "<tr>"
  response.write "<td class=""campor"" colspan='8'>"
  response.write "&nbsp;" & rs("obs")
  response.write "</td></tr>"
end if

laststatus=rs("status")
rs.movenext:loop
rs.close
else 'recordcount=0
  response.write "<tr><td colspan='9' class=campo>"
  response.write "<p>Não indicações para esta categoria de nomeação"
  response.write "</td></tr>"
end if 'recordcount=0

end if 'temp=0
%>
</table>

<% 
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>