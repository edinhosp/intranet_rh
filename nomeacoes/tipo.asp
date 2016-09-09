<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a20")="N" or session("a20")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Grupo de Nomeações</title>
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

sqla="SELECT id_nomeacao, nomeacao, criacao, extinta FROM n_nomeacoes ORDER by nomeacao"
sqla="SELECT n.id_nomeacao, n.NOMEACAO, n.CRIACAO, n.extinta, Count(i.id_indicado) AS Quant " & _
", 'Ativas'=sum(case when getdate() between i.mand_ini and i.mand_fim then 1 else 0 end) " & _
", 'Contrato'=sum(case when getdate() between i.mand_ini and i.mand_fim and i.CONTRATO IS null then 1 else 0 end) " & _
"FROM n_nomeacoes AS n LEFT JOIN n_indicacoes AS i ON n.id_nomeacao = i.id_nomeacao " & _
"GROUP BY n.id_nomeacao, n.NOMEACAO, n.CRIACAO, n.extinta " & _
"ORDER BY n.NOMEACAO "
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>Grupo de Nomeações

<table border="1" cellspacing="0" cellpadding="2" style="border-collapse: collapse" width="600">
<tr>
	<td class=titulo align="center" rowspan="2">Nomeações para</td>
	<td class=titulo align="center" rowspan="2">Criada pela</td>
	<td class=titulo align="center" colspan="3">Indicações</td>
	<td class=titulo align="center" rowspan="2">Extinta</td>
	<td class=titulo align="center" rowspan="2"><img border="0" src="../images/Magnify.gif"></td>
</tr>
<tr>
	<td class=titulo align="center">Ativas</td>
	<td class=titulo align="center">Contr.</td>
	<td class=titulo align="center">Total</td>
</tr>
  
<%
rs.movefirst
do while not rs.eof 
if rs("criacao")<>"" then temp=" (" & rs("criacao") & ")" else temp=""
if rs("ativas")=0 then ativas="" else ativas=rs("ativas")
if rs("contrato")=0 then contrato="" else contrato=rs("contrato")
%>
  <tr>
    <td class=campo><a href="nomeados.asp?codigo=<%=rs("id_nomeacao")%>" class=r>
    <%=rs("nomeacao")%></a></td>
    <td class=campo><%=rs("criacao") %></td>
    <td class=campo align="center"><%=ativas%></td>
    <td class=campo align="center"><%=contrato%></td>
    <td class=campo align="center"><%=rs("quant") %></td>
    <td class=campo align="center">
    <% if rs("extinta")=0 then response.write "<img border='0' src='../images/bullet.gif'>" else response.write "<img border='0' src='../images/bullet_hl.gif'>" %> 
    </td>
    <td class=campo align="center">&nbsp;
    <% if session("a20")="T" then %>
      <a href="tipo_alteracao.asp?codigo=<%=rs("id_nomeacao")%>" onclick="NewWindow(this.href,'AlteracaoGrupoNomeacao','520','150','no','center');return false" onfocus="this.blur()">
	  <img border="0" src="../images/folder95o.gif"></a>
	<% end if %>
    </td>
  </tr>
<%
rs.movenext
loop
%>
<tr><td colspan=7 class=titulo valign="center" align="right">
<% if session("a20")="T" then %>
<a href="tipo_nova.asp" onclick="NewWindow(this.href,'InclusaoGrupoNomeacao','520','150','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif">inserir novo grupo de nomeações</a>
<% end if %>
</td></tr>
</table>

<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>