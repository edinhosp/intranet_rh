<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a54")="N" or session("a54")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Vencimento de Contrato de Experi�ncia</title>
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

if request("codigo")="" then
	temp=1
	sqla="select chapa, nome, dataadmissao, dataadmissao+89 as venc from corporerm.dbo.pfunc " & _
	"where codsituacao<>'D' and codtipo='N' and dataadmissao+89>=getdate()-1 order by dataadmissao"
	'response.write sqla
	sql1=sqla
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	titulo=""
else
	temp=0
	sqla="SELECT f.CHAPA, f.NOME, f.DATAADMISSAO, f.DATAADMISSAO+89 AS Venc, f.CODFUNCAO, c.NOME AS FUNCAO, f.CODSECAO, s.DESCRICAO AS SECAO, f1.NOME as chefe, p.SEXO " & _
"FROM ((((corporerm.dbo.PFUNC f INNER JOIN corporerm.dbo.PSECAO s ON f.CODSECAO=s.CODIGO) INNER JOIN corporerm.dbo.PFUNCAO c ON f.CODFUNCAO=c.CODIGO) " & _
"LEFT JOIN corporerm.dbo.PSUBSTCHEFE ch ON f.CODSECAO=ch.CODSECAO) LEFT JOIN corporerm.dbo.PFUNC AS f1 ON ch.CHAPASUBST=f1.CHAPA) INNER JOIN corporerm.dbo.PPESSOA p ON f.CODPESSOA=p.CODIGO " & _
"WHERE f.CHAPA='" & request("codigo") & "' "
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
end if

if temp=1 then
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=grupo>Emiss�o de avalia��o de experi�ncia</td></tr>
</table>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="center">Nome</td>
	<td class=titulo align="center">Admiss�o</td>
	<td class=titulo align="center">Vencimento</td>
</tr>
<%
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo><%=rs("chapa")%></td>
	<td class=campo><a href="experiencia_avaliacao.asp?codigo=<%=rs("chapa")%>"><%=rs("nome")%></a></td>
	<td class=campo align="center"><%=rs("dataadmissao") %></td>
	<td class=campo align="center"><%=rs("venc") %></td>
</tr>
<%
rs.movenext
loop
rs.close
%>
</table>
<%
else ' temp=0
if rs("sexo")="M" then v1="o" else v1="a"
if rs("sexo")="M" then v2="" else v2="a"
%>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr><td class="campop"><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width=225></td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><b>Of. <%=rs("chapa")%> - RH</b></td></tr>
	<tr><td class="campop" align="right">
	<input type="text" name="txt1" class="form_input" size="29" value="Osasco, <%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %>" style="font-size:10pt">
	</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><input type="text" name="txt0" class="form_input" size="5" value="�"><br>
	<input type="text" name="txt1" class="form_input" size="60" value="Sr(a). <%=rs("chefe")%>" style="font-size:10pt"><br>
	<input type="text" name="txt2" class="form_input" size="60" value="<%=rs("secao")%>" style="font-size:10pt"><br>
	</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">Em atendimento a sua solicita��o, foi contratad<%=v1%> no dia <b><%=rs("dataadmissao")%></b> para prestar servi�os 
	a esse departamento, <%=v1%> Sr<%=v2%>. <b><%=rs("nome")%></b>. Considerando que seu contrato de experi�ncia de tr�s meses terminar� 
	em <b><%=rs("venc")%></b>, solicitamos que V.Sa. se manifeste por escrito, na avalia��o anexa sobre o desempenho d<%=v1%> referid<%=v1%> 
	funcion�ri<%=v1%>, informando-nos, com a maior brevidade poss�vel, se <%=v1%> mesm<%=v1%> atende as exig�ncias do cargo e preenche 
	os requisitos indispens�veis a sua admiss�o definitiva. Em resumo, se amolda aos padr�es da FIEO.</td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tal informa��o, sigilosa e important�ssima, evitar� que o servi�o sofra solu��o de continuidade e tenhamos 
	despesas significativas, totalmente desnecess�rias, com a dispensa d<%=v1%> funcion�ri<%=v1%> logo ap�s a expira��o do contrato 
	experimental.</td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Contando com a costumeira 
	colabora��o de V.Sa. apresentamos nossas cordiais sauda��es.</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop">
	<p align="center" style="line-height: 25px">
	<input type="text" name="txt1" class="form_input" size="60" value="LUIZ FERNANDO DA COSTA E SILVA" style="text-align:center;font-size:10pt;font-weight:bold"><br>
	<input type="text" name="txt2" class="form_input" size="60" value="Reitor" style="text-align:center;font-size:10pt;font-weight:bold"><br>
	</td></tr>
	<tr><td class="campop"></td></tr>
</table>
<DIV style="page-break-after:always"></DIV>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campop" align="center"><b>Relat�rio Sigiloso de Acompanhamento Funcional</b></td>
</tr>
</table>
<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campop" colspan=2>Nome: <b><%=rs("nome")%></td></tr>
<tr>
	<td class="campop">Dept�: <b><%=rs("secao")%></td>
	<td class="campop">Admiss�o: <b><%=rs("dataadmissao")%></td>
</tr>
</table>
<br>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulop align="center">Opini�o sobre o desempenho do funcion�rio</td>
</tr>
</table>
<br>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campop" align="center" rowspan=2>Caracter�sticas</td>
	<td class="campop" align="center" colspan=3>Classifica��o</td>
</tr>
<tr>
	<td class="campop" align="center">Insatisfat�rio</td>
	<td class="campop" align="center">Satisfat�rio</td>
	<td class="campop" align="center">+ q/Satisfat�rio</td>
</tr>
<tr>
	<td class="campop">Relacionamento Pessoal/Integra��o</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
</tr>
<tr>
	<td class="campop">Apresenta��o Pessoal</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
</tr>
<tr>
	<td class="campop">Assiduidade</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
</tr>
<tr>
	<td class="campop">Qualidade dos servi�os</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
</tr>
<tr>
	<td class="campop">Pontualidade</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
</tr>
<tr>
	<td class="campop">Interesse pelos servi�os</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
</tr>
</table>
<br>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr><td class=titulop align="center">Outras considera��es sobre o desempenho do funcion�rio</td></tr>
	<tr><td class="campop" align="center">&nbsp;</td></tr>
	<tr><td class="campop" align="center">&nbsp;</td></tr>
	<tr><td class="campop" align="center">&nbsp;</td></tr>
	<tr><td class="campop" align="center">&nbsp;</td></tr>
	<tr><td class="campop" align="center">&nbsp;</td></tr>
	<tr><td class="campop" align="center">&nbsp;</td></tr>
</table>
<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td rowspan=2 width=40><table border="1" bordercolor="#000000" cellpadding=0 cellspacing=0 width="100%" style="border-collapse: collapse">
		<tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table></td>
	<td class="campop">&nbsp;EFETIVAR</td>
</tr>
<tr>
	<td class="campop">&nbsp;N�O EFETIVAR</td>
</tr>
</table>
<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campop" align="center"><br>&nbsp;<br>&nbsp;<br>&nbsp;Assinatura do chefe</td>
</tr>
</table>
<br>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulor align="center">Devolver em envelope lacrado ao Recursos Humanos</td>
</tr>
</table>
<%
rs.close
end if 'temp=0

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>