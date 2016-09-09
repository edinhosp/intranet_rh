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
<title>Vencimento de Contrato de Experiência</title>
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
<tr><td class=grupo>Emissão de avaliação de experiência</td></tr>
</table>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="center">Nome</td>
	<td class=titulo align="center">Admissão</td>
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
	<tr><td class="campop"><input type="text" name="txt0" class="form_input" size="5" value="À"><br>
	<input type="text" name="txt1" class="form_input" size="60" value="Sr(a). <%=rs("chefe")%>" style="font-size:10pt"><br>
	<input type="text" name="txt2" class="form_input" size="60" value="<%=rs("secao")%>" style="font-size:10pt"><br>
	</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">Em atendimento a sua solicitação, foi contratad<%=v1%> no dia <b><%=rs("dataadmissao")%></b> para prestar serviços 
	a esse departamento, <%=v1%> Sr<%=v2%>. <b><%=rs("nome")%></b>. Considerando que seu contrato de experiência de três meses terminará 
	em <b><%=rs("venc")%></b>, solicitamos que V.Sa. se manifeste por escrito, na avaliação anexa sobre o desempenho d<%=v1%> referid<%=v1%> 
	funcionári<%=v1%>, informando-nos, com a maior brevidade possível, se <%=v1%> mesm<%=v1%> atende as exigências do cargo e preenche 
	os requisitos indispensáveis a sua admissão definitiva. Em resumo, se amolda aos padrões da FIEO.</td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tal informação, sigilosa e importantíssima, evitará que o serviço sofra solução de continuidade e tenhamos 
	despesas significativas, totalmente desnecessárias, com a dispensa d<%=v1%> funcionári<%=v1%> logo após a expiração do contrato 
	experimental.</td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Contando com a costumeira 
	colaboração de V.Sa. apresentamos nossas cordiais saudações.</td></tr>
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
	<td class="campop" align="center"><b>Relatório Sigiloso de Acompanhamento Funcional</b></td>
</tr>
</table>
<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campop" colspan=2>Nome: <b><%=rs("nome")%></td></tr>
<tr>
	<td class="campop">Deptº: <b><%=rs("secao")%></td>
	<td class="campop">Admissão: <b><%=rs("dataadmissao")%></td>
</tr>
</table>
<br>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulop align="center">Opinião sobre o desempenho do funcionário</td>
</tr>
</table>
<br>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campop" align="center" rowspan=2>Características</td>
	<td class="campop" align="center" colspan=3>Classificação</td>
</tr>
<tr>
	<td class="campop" align="center">Insatisfatório</td>
	<td class="campop" align="center">Satisfatório</td>
	<td class="campop" align="center">+ q/Satisfatório</td>
</tr>
<tr>
	<td class="campop">Relacionamento Pessoal/Integração</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
</tr>
<tr>
	<td class="campop">Apresentação Pessoal</td>
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
	<td class="campop">Qualidade dos serviços</td>
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
	<td class="campop">Interesse pelos serviços</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center">&nbsp;</td>
</tr>
</table>
<br>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr><td class=titulop align="center">Outras considerações sobre o desempenho do funcionário</td></tr>
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
	<td class="campop">&nbsp;NÃO EFETIVAR</td>
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