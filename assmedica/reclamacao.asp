<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a85")="N" or session("a85")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Formulário Reclamação/Sugestão</title>
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
set conexao2=server.createobject ("ADODB.Connection")
conexao2.Open application("conexao")
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao2
%>

<table border="1" bordercolor="#CCCCCC" cellpadding="0" cellspacing="0" width="690" height=950 style="border-collapse: collapse">
<tr>
	<td valign=top>

<div align="center">	
<table border="0" bordercolor="#000000" cellpadding="0" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulop align="center" height=30 style="border:0 solid #000000">FICHA DE OCORRÊNCIAS / SERVIÇOS MÉDICOS</td>
</tr>
</table>
<%
teste=2
if teste=1 then imagem="../images/round_square.jpg":tam=15
if teste=2 then imagem="../images/bola.gif":tam=22

%>
<table border="0" bordercolor="#000000" cellpadding="0" width="650" cellspacing="0" style="border-collapse: collapse">
<tr><td height=5></td></tr>
<tr>
	<td class="campop" valign=middle><img src="<%=imagem%>" width="<%=tam%>" height="<%=tam%>" border="0" alt=""></td>
	<td class="campop" align="left">Sugestão</td>
	<td class="campop" valign=middle><img src="<%=imagem%>" width="<%=tam%>" height="<%=tam%>" border="0" alt=""></td>
	<td class="campop" align="left">Reclamação</td>
	<td class="campop" valign=middle><img src="<%=imagem%>" width="<%=tam%>" height="<%=tam%>" border="0" alt=""></td>
	<td class="campop" align="left">Elogio</td>
	<td class="campop" valign=middle><img src="<%=imagem%>" width="<%=tam%>" height="<%=tam%>" border="0" alt=""></td>
	<td class="campop" align="left">Outra (especifique:________________________)</td>
</tr>
<tr><td height=5></td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" width="650" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td height=45 colspan=2 class="campor" valign=top>&nbsp;Nome do Titular<br></td>
</tr>
<tr>
	<td height=45 colspan=2 class="campor" valign=top>&nbsp;Nome do Dependente (preencher se a ocorrência no atendimento for com o dependente)<br></td>
</tr>
<tr>
	<td height=45 class="campor" valign=top>&nbsp;Operadora<br>
	<font style="font-size:13px">&nbsp;[&nbsp;&nbsp;&nbsp;] Unimed Seguros [&nbsp;&nbsp;&nbsp;] Metlife Odonto [&nbsp;&nbsp;&nbsp;] Intermédica
	</td>
	<td height=45 width=250 class="campor" valign=top>&nbsp;Plano<br>
	</td>
</tr>
<tr>
	<td height=45 colspan=2 class="campor" valign=top>&nbsp;Local de Atendimento (nome da clínica/laboratório)<br></td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="0" width="650" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td height=45 class="campor" valign=top style="border-bottom: 1px solid;border-left: 1px solid;border-right: 1px solid">&nbsp;Endereço<br>
	</td>
	<td height=45 width=200 class="campor" valign=top style="border-bottom: 1px solid;border-left: 1px solid;border-right: 1px solid">&nbsp;Cidade<br>
	</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="0" width="650" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td height=45 class="campor" valign=top style="border-bottom: 1px solid;border-left: 1px solid;border-right: 1px solid">&nbsp;Tipo do serviço<br>
	<font style="font-size:13px">&nbsp;[&nbsp;&nbsp;&nbsp;] Consulta [&nbsp;&nbsp;&nbsp;] Exame [&nbsp;&nbsp;&nbsp;] Outros _________________
	</td>
	<td height=45 width=270 class="campor" valign=top style="border-bottom: 1px solid;border-left: 1px solid;border-right: 1px solid">&nbsp;Especialidade<br>
	</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="0" width="650" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td height=45 class="campor" valign=top style="border-left: 1px solid;border-right: 1px solid">&nbsp;Nome do funcionário/médico/atendente<br>
	</td>
	<td height=45 width=180 class="campor" valign=top style="border-left: 1px solid;border-right: 1px solid">&nbsp;CRM/identificação<br>
	</td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="0" width="650" cellspacing="0" style="border-collapse: collapse">
	<tr><td height=30 valign=middle class=titulo align="center">Ocorrência (descreva nas linhas abaixo a sua reclamação/sugestão/elogio)</td></tr>
<%for a=1 to 15%>
	<tr><td height=30 class="campop" align="center">&nbsp;</td></tr>
<%next%>
</table>
<table border="0" bordercolor="#000000" cellpadding="0" width="650" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td height=45 class="campor" valign=top style="border-bottom: 1px solid;border-left: 1px solid;border-right: 1px solid">&nbsp;Local e Data<br>
	</td>
	<td height=45 width=380 class="campor" valign=top style="border-bottom: 1px solid;border-left: 1px solid;border-right: 1px solid">&nbsp;Assinatura<br>
	</td>
</tr>
</table>








</div>
</td></tr></table>	
<%
%>
</body>
</html>
<%
set rs3=nothing
conexao2.close
set conexao2=nothing
%>