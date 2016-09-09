<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a25")="N" or session("a25")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Vencimento das Nomeações</title>
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
'	session("nomeacao_chapa")=""
'	session("nomeacao_id")=""
'	session("nomeacao_descr")=""

if request.form<>"" then
	data1=dtaccess(request.form("t1"))
	data2=dtaccess(request.form("t2"))
	
	sqlb="SELECT Status=case when mand_fim<getdate() then 'Vencidas' else 'Ativas' end, " & _
	"i.id_nomeacao, n.NOMEACAO, i.id_indicado, i.PORTARIA, i.CHAPA, i.NOME, " & _
	"i.CARGO, i.MAND_INI, i.MAND_FIM, i.CH " & _
	"FROM n_indicacoes AS i INNER JOIN n_nomeacoes AS n ON i.id_nomeacao = n.id_nomeacao " & _
	"WHERE i.MAND_FIM Between '" & data1 & "' And '" & data2 & "' " & _
	" ORDER BY case when mand_fim<getdate() then 'Vencidas' else 'Ativas' end, n.nomeacao, i.nome "
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
	temp=0
	titulo=""
else
	temp=1
end if
%>
<p class=titulo>Nomeações&nbsp;<%=titulo %>
<%
if temp=1 then
data_1=dateserial(year(now),month(now),1)
data_2=dateserial(year(now),month(now)+1,1)-1
%>
<form method="POST" action="nomeacoes_termino.asp">
  <p>vencendo entre <input type="text" name="T1" size="12" value="<%=data_1%>" style="text-align:center">
  e <input type="text" name="T2" size="12" value="<%=data_2%>" style="text-align:center">
  <br>
  <input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>
<%
else ' temp=0
%>
<table border="1" bordercolor="gray" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=titulo align="center">Nomeação</td>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="center">Nome    </td>
	<td class=titulo align="center">Portaria</td>
	<td class=titulo align="center">Cargo   </td>
	<td class=titulo align="center">Inicio em</td>
	<td class=titulo align="center">Término  </td>
	<td class=titulo align="center">C.H.     </td>
	<td class=titulo align="center"><img border="0" src="../images/Magnify.gif"></td>
</tr>
<%
laststatus=""
if rs.recordcount>0  then
rs.movefirst
do while not rs.eof 
if laststatus<>rs("status") then
	response.write "<tr><td class=grupo colspan='9'>"
	response.write "&nbsp;" & rs("status") & "</td></tr>"
end if
%>
<tr>
	<td class=campo><a href="nomeados.asp?codigo=<%=rs("id_nomeacao")%>">
		<font size="1"><%=rs("nomeacao") %></font></a></td>
	<td class=campo><a href="nomeados_nome.asp?codigo=<%=rs("chapa")%>"><font size="1"><%=rs("chapa") %></font></a></td>
	<td class=campo><%=rs("nome") %></td>
	<td class=campo><%=rs("portaria") %></td>
	<td class=campo><%=rs("cargo") %></td>
	<td class=campo align="center"><%=rs("mand_ini") %></td>
	<td class=campo align="center"><%=rs("mand_fim") %></td>
	<td class=campo align="center"><%=rs("ch") %></td>
	<td class=campo align="center">
<% if session("a25")="T" then %>
	<a href="nomeados_alteracao.asp?codigo=<%=rs("id_indicado")%>" onclick="NewWindow(this.href,'Alteracao','520','330','no','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width="16" height="16" alt="Clique para alterar a nomeação"></a>
<% end if %>
	</td>
</tr>
<%
laststatus=rs("status")
rs.movenext
loop
else
%>
<tr><td class-campo colspan=9>Não existem portarias vencendo neste período.</td></tr>
<%
rs.close
end if'rs.recordcount
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