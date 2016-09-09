<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Confirmação de convites</title>
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
<link rel="stylesheet" type="text/css" href="<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>
Confirmação de Convites</p>
<form method="POST" action="convites.asp" name="form">
<P>Digite parte do nome para buscar: <input type="text" name="buscanome" size="40" maxlength="256" value="<%=request.form("buscanome")%>" />

<%
if request.form<>"" then
if request.form("buscanome")<>"" then
	buscanome=ucase(request.form("buscanome"))
	filtro="where nome like '%" & buscanome & "%' "
	quant=" top 100 "
else
	filtro=""
	quant=" top 10 "
end if

sql1="select " & quant & " codigo, local, tratamento, nome, descricao, dtconfirmacao, confirmado, naovai from _conv " & filtro & " order by nome "
rs.Open sql1, ,adOpenStatic, adLockReadOnly

%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width="650">
<tr>
	<td class=titulo>Tratamento</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Local</td>
	<td class=titulo>Data</td>
	<td class=titulo>confirmado por</td>
	<td class=titulo>-</td>
</tr>
<%
do while not rs.eof
nome=ucase(rs("nome")):nnn=""
if rs("naovai")=true then nnn="Ausência conf.-<b>Não</b> vem"
if rs("confirmado")<>"" and rs("naovai")=false then nnn="Presença conf."
%>
<tr>
	<td class=campo><%=rs("tratamento")%></td>
	<td class=campo nowrap><%=replace(nome,buscanome,"<font color=blue><b>" & buscanome & "</b></font>")%>
	<br><%=rs("descricao")%>
	</td>
	<td class="campor"><%=nnn%></td>
	<td class=campo><%=rs("dtconfirmacao")%></td>
	<td class=campo><%=rs("confirmado")%></td>
	<td class=campo>
	<a href="convites2.asp?codigo=<%=rs("codigo")%>" onclick="NewWindow(this.href,'confirmacao','400','150','no','center');return false" onfocus="this.blur()">
	<img src="images/bookO.gif" border="0" width=13 alt="confirmar presença"></a>
	</td>
</tr>
<%
rs.movenext:loop
rs.close
%>
</table>

<%
%>	




<%
end if 'request.form("B1")
%>
</form>
</body>
</html>
<%
conexao.close
set conexao=nothing
%>