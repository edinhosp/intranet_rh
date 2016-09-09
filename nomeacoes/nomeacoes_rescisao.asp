<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a26")="N" or session("a26")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Rescisão de Nomeações</title>
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
<body style="margin-left:40px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
if request.form<>"" then
	idnomeacao=request.form("D1")
	if idnomeacao="Todos" then
		sqld=""
	else
		sqld=" and i.chapa='" & idnomeacao & "'"
	end if
	
	sqlc="SELECT i.id_nomeacao, n.NOMEACAO, i.id_indicado, i.CHAPA, i.NOME, i.PORTARIA, " & _
	"i.codeve, i.CARGO, i.MAND_INI, i.MAND_FIM, i.CH, i.obs, s.codigo, s.descricao " & _
	"FROM n_indicacoes as i, n_nomeacoes as n, qry_nomeacoes_setor s " & _
	"WHERE i.id_nomeacao = n.id_nomeacao and (i.mand_fim>'" & dtaccess(now) & "' or i.mand_fim is null) " & _
	"and i.chapa=s.chapa " 
	sqle="order by s.codigo, nome "
	sqlb=sqlc & sqld & sqle
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
	temp=0
	titulo=rs("chapa") & " - " & rs("nome")
		session("nomeacao_chapa")=rs("chapa")
'		session("nomeacao_id")=""
'		session("nomeacao_descr")=""
else
		temp=1
'		session("nomeacao_chapa")=""
'		session("nomeacao_id")=""
'		session("nomeacao_descr")=""
end if

if temp=1 then
	sqla="SELECT chapa, nome from n_indicacoes " & _
	"where mand_fim>'" & dtaccess(now()) & "' or mand_fim is null " & _
	"group by chapa, nome order by nome"
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<p class=titulo>Nomeações para&nbsp;<%=titulo %>
<br>
<form method="POST" action="nomeacoes_rescisao.asp" name="form">
<p><select size="1" name="D1">
	<option value="Todos">Todos</option>
<%
rs.movefirst
do while not rs.eof
%>
	<option value="<%=rs("chapa")%>"><%=rs("nome")%></option>
<%
rs.movenext
loop
rs.close
%>
	</select>
	<br>
	Data para imprimir:
	<input type="text" name="datarescisao" value="&nbsp;&nbsp;&nbsp;&nbsp;de <%=monthname(month(now))%> de <%=year(now)%>" size=30><br>
	<input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>

<!-- impressão do documento -->
<%
else ' temp=0
%>
<%
rs.movefirst
do while not rs.eof 
if rs("cargo")<>"" then complemento=" (" & rs("cargo") & ")" else complemento=""
if rs("portaria")<>"" then
	portaria=" através "
	if left(rs("portaria"),8)="Portaria" or left(rs("portaria"),6)="Instru" then
		portaria=portaria & "da " & rs("portaria")
	else
		portaria=portaria & "do " & rs("portaria")
	end if
else
	portaria=""
end if
if rs("mand_fim")<>"" then
	vencimento=", com vencimento em " & rs("mand_fim")
else
	vencimento=""
end if
%>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=650 height="1000">
  <tr>
    <td height=100% class="campop" align="justify" valign=top colspan=2>
	<p style="font-size: 12pt; text-align:justify">
	<br>
	<br>
	<br>
	<br>Osasco, <%=request.form("datarescisao")%>.<br>
	<br>&nbsp;
	<br>
	<br>
	<br>
	<br>À<br>Fundação Instituto de Ensino para Osasco<br>
	<br>&nbsp;
	<br>&nbsp;
	<br>
	<br>
	<br>
	<br>A partir desta data coloco à disposição da Diretoria desta Instituição, o cargo 
	de confiança e comissão objeto da nomeação para a atividade de: <%=rs("nomeacao")%> 
	<%=complemento%>, que me foi designado <%=portaria%> <%=vencimento%>.<br>
	<br>
	<br>
	<br>
	<br>
	<br>Atenciosamente,
	<br>
	<br>
	<br>
	<br>
	<br><%=string(50,"_")%>	
	<br><%=rs("nome")%>
	
	
	</td>
  </tr>
  <tr>
    <td height=50 class="campor" align="justify">
<p style='margin-top:0; margin-bottom:0'><font size='1'><%=rs("codigo")%>  - <%=rs("descricao")%></font></p>
	</td>
	<td class="campop" align="right"><%=rs("chapa")%>&nbsp;&nbsp;</td>
  </tr>
</table>
<%

if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
rs.movenext
loop
rs.close
end if
%>

<%
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>