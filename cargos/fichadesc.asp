<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a77")="N" or session("a77")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Descrição de Cargo</title>
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
dim conexao, conexao2, chapach
dim rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("B1")="" then
%>
<!-- modelo do relatorio inicio -->
<!-- modelo do relatorio final -->
<form method="POST" action="fichadesc.asp" name="form">
<table border=0 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=500>
<tr><td valign=top colspan=2>
<p style="margin-bottom: 0" class=realce><b>Seleções para a impressão de &quot;Descrição de Cargo&quot;</b></p>
</td></tr>
<tr>
	<td class=titulor nowrap>Tipo da Seleção</td>
	<td class=titulor>Conteúdo da Seleção</td>
</tr>
<tr>
	<td class="campor" nowrap><select size="1" name="selecao" onChange="javascript:submit()">
		<option value="1" <%if request.form("selecao")="1" then response.write "selected"%> >Todos</option>
		<option value="2" <%if request.form("selecao")="2" then response.write "selected"%> >Setor</option>
		<option value="3" <%if request.form("selecao")="3" then response.write "selected"%> >Funcionário</option>
		<option value="4" <%if request.form("selecao")="4" then response.write "selected"%> >Cargo</option>
		<option value="5" <%if request.form("selecao")="5" then response.write "selected"%> >Em branco</option>
	</select>
	</td>
	<td class="campor">
<%
combo=0
select case request.form("selecao")
	case "2" 'setor
		combo=1:sqltemp="select codsecao as codigo, secao as descricao from qry_funcionarios f where f.codsindicato<>'03' and f.codsituacao<>'D' group by codsecao, secao order by secao "
	case "3" 'funcionario
		combo=1:sqltemp="select chapa as codigo, nome as descricao from qry_funcionarios f where f.codsindicato<>'03' and f.codsituacao<>'D' group by chapa, nome order by nome "
	case "4" 'cargo
		combo=1:sqltemp="select funcao as codigo, funcao as descricao from qry_funcionarios f where f.codsindicato<>'03' and f.codsituacao<>'D' group by funcao, funcao order by funcao "
end select
if combo=1 then
%>
<select size="1" name="cselecao">
<%
rs.Open sqltemp, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
%>
<option value="<%=rs("codigo")%>"><%=rs("descricao")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<%
end if 'selecao combo 1
%>
		</td>
	</tr>
<tr><td valign=top colspan=2>
<p><input type="submit" class=button value="Visualizar Relatório" name="B1"></p>
</td></tr>

<tr><td valign=top class=campoe colspan=2>
<p style="margin-top: 0; margin-bottom: 0"><font color="#FF0000">Configure a página do seu navegador (Internet
Explorer, Netscape, Mozilla, etc) no sentido RETRATO.</font></p>
</td></tr></table>

</form>
<%
end if  'if do request.form

'***************************************************************

if request.form("B1")<>"" then

filtro="":filtro2="":selecao=""
escolha=request.form("selecao")
select case request.form("selecao")
	case "1" 'todos
		filtrow="WHERE F.CODSINDICATO<>'03' AND F.CODSITUACAO<>'D' AND F.CODTIPO='N' "
		filtroh=""
		selecao="Seleção: todos registros"
	case "2" 'setor
		filtrow="WHERE F.CODSECAO='" & request.form("cselecao") & "' AND F.CODSINDICATO<>'03' AND F.CODSITUACAO<>'D' AND F.CODTIPO='N' "
		filtroh=""
		selecao="Seleção: " & request.form("cselecao")
	case "3" 'funcionario
		filtrow="WHERE F.CHAPA='" & request.form("cselecao") & "' AND F.CODSINDICATO<>'03' AND F.CODSITUACAO<>'D' AND F.CODTIPO='N' "
		filtroh=""
		selecao="Seleção: " & request.form("cselecao")
	case "4" 'cargo
		filtrow="WHERE C.NOME='" & request.form("cselecao") & "' AND F.CODSINDICATO<>'03' AND F.CODSITUACAO<>'D' AND F.CODTIPO='N' "
		filtroh=""
		selecao="Seleção: " & request.form("cselecao")
	case "5" 'em branco
		filtrow=""
		filtroh=""
		selecao="Seleção: branco"
end select

sqla="SELECT F.CHAPA, F.NOME, F.CODSINDICATO, F.CODSITUACAO, F.CODSECAO, S.DESCRICAO AS SECAO, F.CODFUNCAO, C.NOME AS FUNCAO, F.DATAADMISSAO, C.CBO2002 " & _
"FROM (corporerm.dbo.PFUNC AS F INNER JOIN corporerm.dbo.PSECAO AS S ON F.CODSECAO = S.CODIGO) INNER JOIN corporerm.dbo.PFUNCAO AS C ON F.CODFUNCAO = C.CODIGO "
sqlb=Filtrow
sqlc="" 'GROUP
sqld=Filtroh
sqle="ORDER BY F.CODSECAO, F.NOME "
if escolha<>"5" then 
	sqlz=sqla & sqlb & sqlc & sqld & sqle
else
	sqlz="select top 1 chapa from corporerm.dbo.pfunc"
end if

rs.Open sqlz, ,adOpenStatic, adLockReadOnly
do while not rs.eof

if escolha="5" then 
	nome="":funcao="":dataadmissao="":secao=""
else
	nome=rs("nome"):funcao=rs("funcao"):dataadmissao=rs("dataadmissao"):secao=rs("secao")
end if
%>
<div align="right">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690" height="1000">
<tr><td class=campo valign="top">

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr><td class=campo align="left" valign="top" width=150>
	<img src="../images/logo_centro_universitario_unifieo_big.jpg" border=0 width=150>
	</td>
<td class=campo align="center" valign="top" height=40 colspan=1>
	<b><font size="3">DESCRIÇÃO DE CARGO</font></b><br>
	(Antes de preencher leia as instruções em anexo)
	</td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr><td class=campo valign="top" width=70%>
	<!-- -->
	<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=100%>
		<tr><td class="campor" height=10 valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid"><b>NOME DO OCUPANTE:</td></tr>
		<tr><td class="campop" height=30 valign="middle" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"><b><%=nome%></td></tr>
		<tr><td class=campo height=5 valign="top"></td></tr>
		<tr><td class="campor" height=10 valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid"><b>TÍTULO DO CARGO:</td></tr>
		<tr><td class="campop" height=30 valign="middle" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"><b><%=funcao%></td></tr>
		<tr><td class=campo height=5 valign="top"></td></tr>
		<tr><td class="campor" height=10 valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid"><b>ADMISSÃO:</td></tr>
		<tr><td class="campop" height=30 valign="middle" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"><b><%=dataadmissao%></td></tr>
		<tr><td class=campo height=5 valign="top"></td></tr>
		<tr><td class="campor" height=10 valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid"><b>ÁREA:</td></tr>
		<tr><td class="campop" height=30 valign="middle" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"><b><%=secao%></td></tr>
	</table>
	<!-- -->
	</td>
	
	<td class=campo valign="top" width=30% align="center">
	<!-- -->
	<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=90%>
		<tr><td class=campo colspan=2 align="center" height=20>Organograma</td></tr>
		<tr><td class="campor" width=70% height=40 style="border:1px solid" valign=top align="center">cargo</td>
			<td class=campo width=30% valign=middle rowspan=2>Superior<br>Imediato</td></tr>
		<tr><td class="campor" width=70% height=40 style="border:1px solid" valign=top align="center">nome</td></tr>

		<tr><td class=campo height=10 valign="top" align="center" style="font-size:12pt"><b><!--&#8597;--></td><td class=campo></td></tr>
	
		<tr><td class="campor" width=70% height=40 style="border:0px solid" valign=top align="center"><!--cargo--></td>
			<td class=campo width=30% valign=middle rowspan=2><!-- Cargo<br>Descrito --></td></tr>
		<tr><td class="campor" width=70% height=40 style="border:0px solid" valign=top align="center"><!--nome--></td></tr>
	</table>
	<!-- -->
	</td>
</tr>
<tr><td class=campo align="center" valign="top" height=5 colspan=2></td></tr>

<tr><td class=campo align="center" valign="top" colspan=2>
	<!-- -->
	<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=100%>
		<tr><td class="campop" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid"><b>DESCRIÇÃO RESUMIDA DO CARGO:</td></tr>
		<tr><td class="campop" valign="top" height=30 style="border-bottom:1px dashed;border-left:1px solid;border-right:1px solid">&nbsp;</td></tr>
		<tr><td class="campop" valign="top" height=30 style="border-bottom:1px dashed;border-left:1px solid;border-right:1px solid">&nbsp;</td></tr>
		<tr><td class="campop" valign="top" height=5 style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"></td></tr>
	</table>
	<!-- -->
	</td></tr>
<tr><td class=campo align="center" valign="top" height=5 colspan=2></td></tr>

<tr><td class=campo align="left" valign="top" colspan=2>
	<b><font size="2">PRINCIPAIS RESPONSABILIDADES:</font></b>
	(caso necessário utilize folha em branco)
	</td></tr>

<tr><td class=campo align="center" valign="top" colspan=2 style="border:1px solid">
	<!-- --> <%linha1=35%>
	<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="4" style="border-collapse: collapse" width=100%>
		<tr><td class="campop" valign="middle" colspan=2 height=20 style="border-bottom:1px solid">No espaço abaixo descreva as tarefas diárias (aquelas que se repetem com regularidade).</td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class="campop" valign="middle" colspan=2 height=20 style="border-bottom:1px solid">No espaço abaixo descreva as suas tarefas periódicas e a sua frequência (mensal, trimestral, semestral ou anual), da mesma forma que no item anterior.</td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
		<tr><td class=campo width="5%" style="border-bottom:1px solid;border-right:1px solid" height=<%=linha1%>></td>
			<td class=campo width="95%" style="border-bottom:1px solid"></td></tr>
	</table>
	<!-- -->
	</td></tr>
</table>
<%
%>
</td></tr> <!-- tabela pagina -->
</table>
</div>

<DIV style="page-break-after:always"></DIV> <!-- quebra de pagina -->

<div align="left">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690" height="1000">
<tr><td class=campo valign="top">

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="100%" height=100%>
	<tr><td class=campo align="center" valign="middle" height=25 colspan=3><b><font size="2">ESPECIFICAÇÃO DO CARGO</font></b></td></tr>
	<tr><td class=fundo align="center" valign="middle" height=25 colspan=3><b><font size="2">REQUISITOS DO CARGO</font></b>   </td></tr>
	<tr><td class="campop" colspan=3><b>01. NIVEL DE INSTRUÇÃO</b><br>
		Qual o nível de instrução que você considera como mínimo necessário para ocupar o seu cargo? <u>Não mencione o seu nível de instrução</u>, mas aquele
		necessário para o exercício satisfatório do cargo.</td></tr>
	<tr><td class=campo colspan=3>
	<!-- -->
	<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=100%>
		<tr><td class="campop" width=5% height=30 valign=middle> (&nbsp;&nbsp;&nbsp;) </td>
			<td class="campop" width=30%>Ensino Fundamental incompleto </td>
			<td class="campop" width=5%> (&nbsp;&nbsp;&nbsp;)  </td>
			<td class="campop" width=25%>Ensino Médio incompleto</td>
			<td class="campop" width=5%> (&nbsp;&nbsp;&nbsp;)  </td>
			<td class="campop" width=30%>Ensino Superior incompleto </td></tr>
		<tr><td class="campop" width=5% height=30 valign=middle> (&nbsp;&nbsp;&nbsp;) </td>
			<td class="campop" width=30%>Ensino Fundamental completo </td>
			<td class="campop" width=5%> (&nbsp;&nbsp;&nbsp;)  </td>
			<td class="campop" width=25%>Ensino Médio completo</td>
			<td class="campop" width=5%> (&nbsp;&nbsp;&nbsp;)  </td>
			<td class="campop" width=30%>Ensino Superior completo </td></tr>
	</table>
	<!-- -->
	</td></tr>
	<tr><td class="campop" align="left" valign="top" height=45 colspan=3 style="border-bottom:1px dashed">
		No caso de necessidade de Curso Técnico ou Superior, mencionar o nome do curso:</td></tr>
	<tr><td class="campop" align="left" valign="top" height=60 colspan=3 style="border-bottom:1px dashed">
		Além do nível de instrução assinalado, há necessidade de algum tipo de especialização acadêmica complementar 
		(Pós-graduação, MBA, Mestrado ...)? Especifique:</td></tr>

	<tr><td class=campo align="center" valign="top" height=10 colspan=3></td></tr>
	<tr><td class="campop" colspan=3><b>02. CONHECIMENTOS TÉCNICOS OU ESPECÍFICOS PARA REALIZAÇÃO DAS ATIVIDADES</b></td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>

	<tr><td class=campo align="center" valign="top" height=10 colspan=3></td></tr>
	<tr><td class="campop" colspan=3><b>03. TEMPO DE EXPERIÊNCIA</b><br>
		Qual o tempo mínimo de experiência que o ocupante do cargo deve possuir, considerando o grau de instrução apontado no ítem 1, para
		desempenhar satisfatoriamente as tarefas do cargo?</td></tr>
	<tr><td class="campop" colspan=3 height=30>Anos ___________ <%for a=1 to 10%>&nbsp;<%next%> Meses ___________
	</td></tr>

	<tr><td class=campo align="center" valign="top" height=10 colspan=3></td></tr>
	<tr><td class="campop" colspan=3><b>04. IDIOMAS</b><br>
		Indique, além do português, outros idiomas necessários para o desempenho de suas atividades?</td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>
	
	<tr><td class=campo align="center" valign="top" height=10 colspan=3></td></tr>
	<tr><td class="campop" colspan=3><b>05. OBSERVAÇÕES / ESCLARECIMENTOS</b><br>
		Espaço reservado para complemento, observações, sugestões e esclarecimentos dos itens tratados:</td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>
	<tr><td class="campop" colspan=3 height=30 style="border-bottom:1px dashed"></td></tr>

	<tr><td class=campo align="center" valign="top" height=100% colspan=3></td></tr>

	<tr><td class="campop" height=10 width=40% valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Assinatura do funcionário:</td>
		<td class="campop" height=10 width=20% valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Data:</td>
		<td class="campop" height=10 width=40% valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Assinatura do Superior Imeditato:</td>		
	</tr>
	<tr><td class="campop" height=40 valign="middle" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"></td>
		<td class="campop" height=40 valign="middle" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"></td>
		<td class="campop" height=40 valign="middle" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"></td>
	</tr>

	<tr><td class="campor" align="right" valign="bottom" height=10 colspan=3>FDC mod.1 - 2008</td></tr>
		
</table>

</td></tr> <!-- tabela pagina -->
</table>
</div>
<%	
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>"
rs.movenext
loop

rs.close
set rs=nothing
pagina=pagina+1

end if 'if do request.form

conexao.close
set conexao=nothing
%>
</body>
</html>