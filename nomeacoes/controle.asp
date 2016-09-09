<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=false
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a23")="N" or session("a23")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Assinaturas</title>
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

%>
<form method="POST" action="controle.asp">
<p class=titulo>Controle de Assinaturas e Arquivo
</form>
<%
sql="SELECT i.id_indicado, 'Status'=case when mand_fim<getdate() then 'Vencidas' else 'Ativas' end, " & _
"n.NOMEACAO, i.CHAPA, i.NOME, i.PORTARIA, i.CARGO, i.MAND_INI, i.MAND_FIM, i.CH, i.obs, i.contrato, i.entrega " & _
"FROM n_indicacoes as i inner join n_nomeacoes as n on i.id_nomeacao=n.id_nomeacao " & _
"WHERE CONTRATO is not null and (entrega is null or entrega='') " & _
"ORDER BY (case when mand_fim<getdate() then 'Vencidas' else 'Ativas' end), i.NOME "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border="1" bordercolor="gray" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=titulo align="center">Status</td>
	<td class=titulo align="center">Pessoa  </td>
	<td class=titulo align="center">Nomeação/Portaria</td>
	<td class=titulo align="center">Período</td>
	<td class=titulo align="center">C.H.</td>
	<td class=titulo align="center"><img border="0" src="../images/Magnify.gif"></td>
	<td class=titulo align="center">Contrato  </td>
	<td class=titulo align="center">Dev.<br>Contr.</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 
%>
<tr>
	<td class=campo><%=rs("status") %></td>
	<td class=campo><%=rs("chapa") & " - " & rs("nome") %></td>
	<td class=campo><%=rs("nomeacao") & "<br>" & rs("cargo") & "<br>" & rs("portaria")  %></td>
	<td class=campo><%=rs("mand_ini") & " a " & rs("mand_fim") %></td>
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
rs.movenext
loop
rs.close
else
%>
<tr><td class=campo colspan=8>Não existe contratos pendentes para assinatura e arquivo.</td></tr>
<%
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