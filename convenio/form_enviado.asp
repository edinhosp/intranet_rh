<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a79")="N" or session("a79")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Formulário de Encaminhamento</title>
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
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

id_env=request("codigo")

sql="select b.chapa, f.nome, c.nome as funcao, f.dataadmissao, p.cartidentidade as rg, p.cpf, " & _
"m.faculdade, m.mantenedora, m.contato, m.email, m.telefone, " & _
"b.curso, b.periodo, b.encaminhamento, b.obs, b.anoletivo, b.data " & _
"from corporerm.dbo.pfunc f, rhconveniados b, rhconveniobe m, corporerm.dbo.ppessoa p, corporerm.dbo.pfuncao c " & _
"where f.chapa collate database_default=b.chapa and b.id_faculdade=m.id_faculdade and f.codpessoa=p.codigo and f.codfuncao=c.codigo " & _
"and b.id_env=" & id_env
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<div align="right">

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650" height="970">
<tr>
	<td><img border="0" src="../images/aguia.jpg" width="236"></td>
</tr>
<tr>
	<td><p align="center"><b><font size="3">CONVÊNIO BOLSA DE ESTUDO - GRADUAÇÃO</font></b></p>
      <p align="center">&nbsp;</td>
</tr>
<tr>
	<td style="border: 1px solid #000000;">
	<table border="0" cellpadding="4" width="100%" cellspacing="0">
	<tr>
		<td width="100">Emitente:</td>
		<td colspan="3">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td>
	</tr>
	<tr>
		<td width="100"></td>
		<td colspan="3"><b>CENTRO UNIVERSITÁRIO FIEO</b></td>
	</tr>
	<tr>
		<td width="100">Responsável:</td>
		<td>ROGERIO MATEUS DOS SANTOS ARAUJO</td>
		<td>Telefone:</td>
		<td>3651-9972</td>
	</tr>
	<tr>
		<td width="100">E-mail:</td>
		<td>rogerio@unifieo.br</td>
		<td></td>
		<td></td>
	</tr>
	</table>
	</td>
</tr>
<tr>
	<td style="border: 1px solid #000000">
	<table border="0" cellpadding="4" width="100%" cellspacing="0">
	<tr>
		<td width="150">Instituição Conveniada:</td>
		<td colspan="3"><%=ucase(rs("mantenedora")) %></td>
	</tr>
	<tr>
		<td width="150"></td>
		<td colspan="3"><b><%=ucase(rs("faculdade")) %></b></td>
	</tr>
	<tr>
		<td width="150">Responsável:</td>
		<td><%=rs("contato") %></td>
		<td>Telefone:</td>
		<td><%=rs("telefone") %></td>
	</tr>
	<tr>
		<td width="150">E-mail:</td>
		<td><%=rs("email") %></td>
		<td></td>
		<td></td>
	</tr>
	</table>
	</td>
</tr>
<tr>
	<td style="border-left: 1px solid #000000;border-right: 1px solid #000000;border-top: 1px solid #000000;">
<%if rs("encaminhamento")="1" then %>
    <p>Estamos encaminhando nosso funcionário para inscrição no concurso
	Vestibular <%=rs("anoletivo")%>, conforme convênio de Bolsa de Estudos. <%end if%>
<%if rs("encaminhamento")="2" then %>
	<p>Estamos encaminhando nosso funcionário aprovado no concurso Vestibular <%=rs("anoletivo")%>,
	a fim de efetivar a sua matrícula nesta Instituição de Ensino. <%end if%>
<%if rs("encaminhamento")="3" then %>
	<p>Solicitamos a renovação de matrícula para o ano letivo de <%=rs("anoletivo")%> do
	funcionário abaixo relacionada junto a esta Instituição de Ensino. <%end if%>
    </td>
</tr>
<tr><td style="border-left: 1px solid #000000;border-right: 1px solid #000000"></td></tr>
<tr><td style="border-left: 1px solid #000000;border-right: 1px solid #000000;border-bottom: 1px solid #000000">    
	<table border="0" cellpadding="5" width="100%" cellspacing="0">
	<tr>
		<td width="150">Nome:</td>
		<td><b> <%=rs("nome")%>    </b></td>
	</tr>
	<tr>
		<td width="150">Cargo:</td>
		<td> <%=rs("funcao")%>    </td>
	</tr>
	<tr>
		<td width="150">Admissão:</td>
		<td> <%=rs("dataadmissao")%></td>
	</tr>
	<tr>
		<td width="150">R.G.:</td>
		<td> <%=rs("rg")%>    </td>
	</tr>
	<tr>
		<td width="150">C.P.F.:</td>
		<td> <%=rs("cpf")%>    </td>
	</tr>
	<tr>
		<td width="150">Curso:</td>
		<td> <%=rs("curso")%>    </td>
	</tr>
	<tr>
		<td width="150">Semestre:</td>
		<td><input type="text" size="15" class="form_input10"></td>
	</tr>
	<tr>
		<td width="150">Período:</td>
		<td> <%=rs("periodo")%></td>
	</tr>
	<tr>
		<td width="150"><i>Campus</i>:</td>
		<td><input type="text" size="25" class="form_input10"></td>
	</tr>
	</table>
	</td>
</tr>
<tr><td><p>&nbsp;</td></tr>

<tr><td>OBS.: <%=rs("obs")%></td></tr>

<tr><td>&nbsp;</td></tr>
<tr>
	<td>Atenciosamente,
		<p><font size="2">Osasco,&nbsp;<%=day(rs("data")) & " de " & monthname(month(rs("data"))) & " de " & year(rs("data")) %></font></p>
		<p>&nbsp;</p>
    	<p>______________________________________________<br>
		<input type="text" size="50" class="form_input10" value="ROGERIO MATEUS DOS SANTOS ARAUJO"></p>
		<p>&nbsp;</p>
	</td>
</tr>
<tr><td><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b></td></tr>
<tr><td><font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000</font></td></tr>
<tr><td><font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP06020-190 - Fone: (011) 3651-9999</font></td></tr>
<tr><td><font face="Arial Narrow">Caixa Postal - ACF - Jd. Ipê - nº 2530 - Osasco - SP - CEP 06053-990</font></td></tr>
</table>
</div>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

</body>
</html>
<%
rs.close
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%>