<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a23")="N" or session("a23")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Nomeações por docente</title>
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

'if temp=1 then
	sqla="SELECT chapa, nome from n_indicacoes group by chapa, nome order by nome"
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<form method="POST" action="nomeados_nome.asp">
<p class=titulo>Nomeações para&nbsp;<%=titulo %>
	<select size="1" name="chapa" onchange="javascript:submit()">
<%
rs.movefirst:do while not rs.eof
if request.form("chapa")=rs("chapa") then tempc="selected" else tempc=""
%>
	<option value="<%=rs("chapa")%>" <%=tempc%> ><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
	</select>
	<input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>
<%
'else ' temp=0
if request.form("chapa")<>"" then 'novo

if request.form="" then idnomeacao=request("codigo")
if request("codigo")="" then idnomeacao=request.form("D1")
chapa=request.form("chapa")
	
sqlc="SELECT Status=case when mand_fim<getdate() then 'Vencidas' else 'Ativas' end, " & _
"i.id_nomeacao, n.NOMEACAO, i.id_indicado, i.CHAPA, i.NOME, i.PORTARIA, " & _
"i.codeve, i.CARGO, i.MAND_INI, i.MAND_FIM, i.CH, i.obs, i.contrato, i.entrega " & _
"FROM n_indicacoes as i, n_nomeacoes as n " & _
"WHERE i.id_nomeacao = n.id_nomeacao "
sqld=" and i.chapa='" & chapa & "'"
sqle=" ORDER BY (case when mand_fim<getdate() then 'Vencidas' else 'Ativas' end), n.nomeacao, i.mand_ini "
sqlb=sqlc & sqld & sqle
rs.Open sqlb, ,adOpenStatic, adLockReadOnly
temp=0
titulo=rs("chapa") & " - " & rs("nome")
%>
<p class=titulo>
<% if session("a23")="T" and temp=0 then %>
<a href="nomeados_nova.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'InclusaoNomeacao','520','330','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif">
</font></a> <% end if %>
Nomeações para&nbsp;<%=titulo %><br>

<table border="1" bordercolor="gray" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=titulo align="center">Nomeação</td>
	<td class=titulo align="center">Portaria  </td>
	<td class=titulo align="center">Cargo     </td>
	<td class=titulo align="center">Folha     </td>
	<td class=titulo align="center">Inicio em </td>
	<td class=titulo align="center">Término   </td>
	<td class=titulo align="center">C.H.      </td>
	<td class=titulo align="center"><img border="0" src="../images/Magnify.gif"></td>
	<td class=titulo align="center">Contrato  </td>
	<td class=titulo align="center">Dev.<br>Contr.</td>
</tr>
<%
laststatus=""
rs.movefirst
do while not rs.eof 
if laststatus<>rs("status") then
  response.write "<tr><td class=grupo colspan='10'>"
  response.write "&nbsp;" & rs("status") & "</td></tr>"
end if
%>
	<tr>
  	<%if rs("obs")<>"" then response.write "<td class=campo rowspan='2'>" else response.write "<td class=campo>" %>
    	<%=rs("nomeacao") %></td>
  	<%if rs("obs")<>"" then response.write "<td class=campo rowspan='2'>" else response.write "<td class=campo>" %>
    	<%=rs("portaria") %></td>
	<td class=campo><%=rs("cargo") %></td>
	<td class=campo><%=rs("codeve") %></td>
	<td class=campo align="center"><%=rs("mand_ini") %></td>
	<td class=campo align="center"><%=rs("mand_fim") %></td>
	<td class=campo align="center"><%=rs("ch") %></td>
	<td class=campo align="center">
	<% if session("a23")="T" then %>
		<a href="nomeados_alteracao.asp?codigo=<%=rs("id_indicado")%>" onclick="NewWindow(this.href,'AlteracaoNomeacao','520','330','no','center');return false" onfocus="this.blur()">
		<img border='0' src='../images/folder95o.gif'></a>
	<% end if %>
	</td>
	<td class=campo align="center">
	<% if session("a23")="T" then %>
		<a href="nomeados_contrato.asp?codigo=<%=rs("id_indicado")%>" onclick="NewWindow(this.href,'ImpressaoContrato','690','400','yes','center');return false" onfocus="this.blur()">
	<% end if %>
	<font size="1">
<%
	if rs("contrato")<>"" then 
		response.write rs("contrato")
	else
		if session("a23")="T" then
			response.write "<img border='0' src='../images/novo.gif'>"
		else 
			response.write "&nbsp;"
		end if
	end if
%>
	</font>
<% if session("a23")="T" then response.write "</a>" %>
	</td>
	<td class="campor" align="center"><%=rs("entrega") %></td>
</tr>
<%
laststatus=rs("status")
if rs("obs")<>"" then
	response.write "<tr>"
	response.write "<td class=campo colspan='8'>"
	response.write "&nbsp;" & rs("obs") & "</td></tr>"
end if

rs.movenext
loop
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