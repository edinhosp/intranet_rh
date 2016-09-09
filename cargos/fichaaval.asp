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
<title>Avaliação de Desempenho Funcional</title>
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
<form method="POST" action="fichaaval.asp" name="form">
<table border=0 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=500>
<tr><td valign=top colspan=2>
<p style="margin-bottom: 0" class=realce><b>Seleções para a impressão de &quot;Avaliação de Desempenho Funcional&quot;</b></p>
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
		<option value="5" <%if request.form("selecao")="5" then response.write "selected"%> >Grupo</option>
		<option value="6" <%if request.form("selecao")="6" then response.write "selected"%> >Em branco</option>
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
	case "5" 'grupo
		combo=1:sqltemp="select codgrupoocup as codigo, nomegrupoocup as descricao from corporerm.dbo.vgrupoocupacional f "
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
	case "5" 'grupo
		filtrow="WHERE grupo='" & request.form("cselecao") & "' AND F.CODSINDICATO<>'03' AND F.CODSITUACAO<>'D' AND F.CODTIPO='N' "
		filtroh=""
		selecao="Seleção: " & request.form("cselecao")
	case "5" 'em branco
		filtrow=""
		filtroh=""
		selecao="Seleção: branco"
end select

sqla="SELECT top 5 F.CHAPA, F.NOME, F.CODSINDICATO, F.CODSITUACAO, F.CODSECAO, S.DESCRICAO AS SECAO, F.CODFUNCAO, C.NOME AS FUNCAO, " & _
"F.DATAADMISSAO, C.CBO2002, g.Grupo, o.NOMEGRUPOOCUP nomegrupo " & _
", campus=case left(codsecao,2) when '01' then 'Narciso' when '03' then 'V.Yara' when '04' then 'Jd.Wilson' end " & _
"FROM (corporerm.dbo.PFUNC AS F INNER JOIN corporerm.dbo.PSECAO AS S ON F.CODSECAO = S.CODIGO) " & _
"INNER JOIN corporerm.dbo.PFUNCAO AS C ON F.CODFUNCAO = C.CODIGO " & _
"inner join iAvaliacaoGC g on g.codfuncao=f.codfuncao collate database_default " & _
"left join corporerm.dbo.VGRUPOOCUPACIONAL o on o.CODGRUPOOCUP collate database_default=g.grupo "
sqlb=Filtrow
sqlc="" 'GROUP
sqld=Filtroh
sqle="ORDER BY F.CODSECAO, F.NOME "
if escolha<>"6" then 
	sqlz=sqla & sqlb & sqlc & sqld & sqle
else
	sqlz="select top 1 chapa from corporerm.dbo.pfunc"
end if

rs.Open sqlz, ,adOpenStatic, adLockReadOnly
do while not rs.eof

if escolha="6" then 
	nome="":funcao="":dataadmissao="":secao=""
else
	nome=rs("nome"):funcao=rs("funcao"):dataadmissao=rs("dataadmissao"):secao=rs("secao")
end if
%>
<div align="right">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690" height="1000">
<tr><td class=campo valign="top">

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campop" align="left" valign="top" height=40>
	<b>AVALIAÇÃO DE DESEMPENHO FUNCIONAL<br><%=ucase(rs("nomegrupo"))%></b>
	</td>
	<td class=campo align="right" valign="top" width=150>
	<img src="../images/logo_centro_universitario_unifieo_big.jpg" border=0 width=150>
	</td>
</tr>
<tr><td class=campo colspan=2 valign="top" height=5 style="border-top:1px solid"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campor" valign="top"><b>AVALIADO:</td>
	<td class="campor" valign="top"><b>Unidade:</td>
	<td class="campor" valign="top"><b>Período considerado:</td>
</tr>
<tr>
	<td class="campo" valign="top">&nbsp;&nbsp;<%=rs("nome")%></td>
	<td class="campo" valign="top">&nbsp;&nbsp;<%=rs("campus")%></td>
	<td class="campo" valign="top">&nbsp;&nbsp;____/____ a ____/____</td>
</tr>
<tr><td class=campo colspan=3 valign="top" height=5 style="border-top:0px solid"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campor" valign="top"><b>Data Admissão:</td>
	<td class="campor" valign="top"><b>Cargo Atual:</td>
	<td class="campor" valign="top"><b>Área/Depto:</td>
</tr>
<tr>
	<td class="campo" valign="top">&nbsp;&nbsp;<%=rs("dataadmissao")%></td>
	<td class="campo" valign="top">&nbsp;&nbsp;<%=rs("funcao")%></td>
	<td class="campo" valign="top">&nbsp;&nbsp;<%=rs("secao")%></td>
</tr>
<tr><td class=campo colspan=3 valign="top" height=5 style="border-top:0px solid"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campor" valign="top"><b>AVALIADOR (Nome):</td>
	<td class="campor" valign="top"><b>Cargo do Avaliador:</td>
</tr>
<tr>
	<td class="campo" valign="top">&nbsp;&nbsp;</td>
	<td class="campo" valign="top">&nbsp;&nbsp;</td>
</tr>
<tr><td class=campo colspan=2 valign="top" height=5 style="border-top:1px solid"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campor" width=100 valign="middle" align="center" style="border:1px solid"><b>Escala de Conceitos da Avaliação</td>
	<td class="campor" valign="top" style="border:1px solid">
	<b>(1) Crítico (Muito abaixo do Esperado)</b> - necessitando urgentemente desenvolver.
	<b>(2) Regular (Abaixo do Esperado)</b> - necessitando ainda desenvolver.
	<b>(3) Bom (Dentro do Esperado)</b> - atende satisfatoriamente, podendo melhorar, desenvolver.
	<b>(4) Muito Bom (Acima do Esperado)</b> - atende plenamente.
	<b>(5) Ótimo (Muito acima do Esperado)</b> - atende com excelência, supera o esperado.
	</td>
	<td class="campor" width="" align="center" style="border:1px solid"><img src="..\images\aval01.png" border="0"></td>
	<td class="campor" width="" align="center" style="border:1px solid"><img src="..\images\aval02.png" border="0"></td>
	<td class="campor" width="" align="center" style="border:1px solid"><img src="..\images\aval03.png" border="0"></td>
</tr>
<tr><td class=campo colspan=5 valign="top" height=1 style="border:0px solid"></td></tr>
</table>

<p style="font-size:7pt;margin-bottom:0;margin-top:3"><b>BLOCO I - CONHECIMENTOS DO TRABALHO</p>
<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<%
sqlb1="select bloco, fator, p1, p2, af from iAvaliacaoTopicos where grupo='" & rs("grupo") & "' and bloco=1 order by bloco, fator"
rs2.Open sqlb1, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>
<tr>
	<td class="campor" align="center" width="25" height="20"><%=rs2("bloco")&"."&rs2("fator")%></td>
	<td class="campor"><b><%=rs2("p1")%></b> - <%=rs2("p2")%></td>
	<td class="campor" width="26"></td>
	<td class="campor" width="26"></td>
	<td class="campor" width="26" align="center"><%=rs2("af")%></td>
</tr>	
<%
rs2.movenext
loop
rs2.close
%>
</table>

<p style="font-size:7pt;margin-bottom:0;margin-top:3"><b>BLOCO II - GESTÃO DO TRABALHO</p>
<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<%
sqlb1="select bloco, fator, p1, p2, af from iAvaliacaoTopicos where grupo='" & rs("grupo") & "' and bloco=2 order by bloco, fator"
rs2.Open sqlb1, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>
<tr>
	<td class="campor" align="center" width="25" height="20"><%=rs2("bloco")&"."&rs2("fator")%></td>
	<td class="campor"><b><%=rs2("p1")%></b> - <%=rs2("p2")%></td>
	<td class="campor" width="26"></td>
	<td class="campor" width="26"></td>
	<td class="campor" width="26" align="center"><%=rs2("af")%></td>
</tr>	
<%
rs2.movenext
loop
rs2.close
%>
</table>

<p style="font-size:7pt;margin-bottom:0;margin-top:3"><b>BLOCO III - CARACTERÍSTICAS PESSOAIS</p>
<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<%
sqlb1="select bloco, fator, p1, p2, af from iAvaliacaoTopicos where grupo='" & rs("grupo") & "' and bloco=3 order by bloco, fator"
rs2.Open sqlb1, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>
<tr>
	<td class="campor" align="center" width="25" height="20"><%=rs2("bloco")&"."&rs2("fator")%></td>
	<td class="campor"><b><%=rs2("p1")%></b> - <%=rs2("p2")%></td>
	<td class="campor" width="26"></td>
	<td class="campor" width="26"></td>
	<td class="campor" width="26" align="center"><%=rs2("af")%></td>
</tr>	
<%
rs2.movenext
loop
rs2.close
%>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campop" align="center" style="border-bottom:1px solid">
	RESULTADO DA AVALIAÇÃO - CONCEITOS NOS BLOCOS E GERAL
	</td></tr>
<tr><td class=campo valign="top" height=5 style="border:0px solid"></td></tr>
</table>	

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campor" width="25"></td>
	<td class="campor"></td>
	<td style="font-size:7pt" class="campor">Nº de Fatores</td>
	<td style="font-size:7pt" class="campor">Pontos Total</td>
	<td style="font-size:7pt" class="campor">Conceito</td>
	<td class="campor"></td>
	<td class="campor"></td>
	<td style="font-size:7pt" class="campor">Nº de Fatores</td>
	<td style="font-size:7pt" class="campor">Pontos Total</td>
	<td style="font-size:7pt" class="campor">Conceito</td>
	<td class="campor"></td>
	<td class="campor"></td>
	<td style="font-size:7pt" class="campor">Nº de Fatores</td>
	<td style="font-size:7pt" class="campor">Pontos Total</td>
	<td style="font-size:7pt" class="campor">Conceito</td>
	<td class="campor"></td>
</tr>
<tr>
	<td class="campor" width="25" height="20"></td>
	<td class="campor" style="border:1px solid" align="center"><b>Bloco I</td>
	<td class="campor" style="border:1px solid" title="N.Fatores"></td>
	<td class="campor" style="border:1px solid" title="Pontos T"></td>
	<td class="campor" style="border:1px solid" title="Conceito"></td>
	<td class="campor" width="15" title="-"></td>
	<td class="campor" style="border:1px solid" align="center"><b>Bloco II</td>
	<td class="campor" style="border:1px solid" title="N.Fatores"></td>
	<td class="campor" style="border:1px solid" title="Pontos T"></td>
	<td class="campor" style="border:1px solid" title="Conceito"></td>
	<td class="campor" width="15" title="-"></td>
	<td class="campor" style="border:1px solid" align="center"><b>Bloco III</td>
	<td class="campor" style="border:1px solid" title="N.Fatores"></td>
	<td class="campor" style="border:1px solid" title="Pontos T"></td>
	<td class="campor" style="border:1px solid" title="Conceito"></td>
	<td class="campor" width="25" title="-"></td>
</tr>
<tr><td class="campor" height="5" colspan="16"></td></tr>
<tr>
	<td class="campor" colspan="6" height="18"></td>
	<td class="campor" style="border:1px solid" align="center" colspan=3><b>CONCEITO GERAL = </td>
	<td class="campor" style="border:1px solid" title="Conceito"></td>
	<td class="campor" colspan="6"></td>
</tr>
<tr><td class="campor" height="7" colspan="16" style="border-bottom:1px solid"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr><td class="campo" height="13" colspan="2" style="border-bottom:1px solid" align="center">Considerações do Avaliador</td></tr>
<tr><td class="campor" height="" colspan="2" style="border-bottom:1px solid" align="center">
	<b>Feedback</b> - Destaque os aspectos que julge importantes para o desenvolvimento do colaborador AVALIADO, mencionando
	os pontos que precisam ser melhorados (desenvolvidos ou aprimorados) e também os pontos positivos (pontos fortes), como
	reforço as boas práticas para o exercício da sua função/cargo.
	</td></tr>
<tr><td class="campor" align="center" height="20" style="border-bottom:1px solid"><b>Pontos que devem ser melhorados</td>
	<td class="campor" align="center" style="border-bottom:1px solid;border-left:1px solid"><b>Pontos Fortes que se destacam</td>
	</tr>
<%for a=1 to 4%>	
<tr><td class="campor" align="center" height="20" style="border-bottom:1px solid"></td>
	<td class="campor" align="center" style="border-bottom:1px solid;border-left:1px solid"></td>
	</tr>
<%next%>

<tr><td class="campor" align="center" valign="bottom" height="40" style="">_______________________|__________<br>
                                                                           assinatura - avaliado  &nbsp;&nbsp;&nbsp; data</td>
	<td class="campor" align="center" valign="bottom" style="">_______________________|__________<br>
	                                                           assinatura - avaliador &nbsp;&nbsp;&nbsp; data</td>
	</tr>
</table>

<!-- -->
</td></tr>
</table>
</div>

<DIV style="page-break-after:always"></DIV> <!-- quebra de pagina -->

<div align="left">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690" height="1000">
<tr><td class=campo valign="top">

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campop" align="left" valign="top" height=40>
	<b>AVALIAÇÃO DE DESEMPENHO FUNCIONAL<br><%=ucase(rs("nomegrupo"))%></b>
	</td>
	<td class=campo align="right" valign="top" width=150>
	<img src="../images/logo_centro_universitario_unifieo_big.jpg" border=0 width=150>
	</td>
</tr>
<tr><td class=campo colspan=2 valign="top" height=5 style="border-top:1px solid"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campor" valign="top"><b>AVALIADO:</td>
	<td class="campor" valign="top"><b>Unidade:</td>
	<td class="campor" valign="top"><b>Período considerado:</td>
</tr>
<tr>
	<td class="campo" valign="top">&nbsp;&nbsp;<%=rs("nome")%></td>
	<td class="campo" valign="top">&nbsp;&nbsp;<%=rs("campus")%></td>
	<td class="campo" valign="top">&nbsp;&nbsp;____/____ a ____/____</td>
</tr>
<tr><td class=campo colspan=3 valign="top" height=5 style="border-top:0px solid"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campor" valign="top"><b>Data Admissão:</td>
	<td class="campor" valign="top"><b>Cargo Atual:</td>
	<td class="campor" valign="top"><b>Área/Depto:</td>
</tr>
<tr>
	<td class="campo" valign="top">&nbsp;&nbsp;<%=rs("dataadmissao")%></td>
	<td class="campo" valign="top">&nbsp;&nbsp;<%=rs("funcao")%></td>
	<td class="campo" valign="top">&nbsp;&nbsp;<%=rs("secao")%></td>
</tr>
<tr><td class=campo colspan=3 valign="top" height=5 style="border-top:0px solid"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campor" valign="top"><b>AVALIADOR (Nome):</td>
	<td class="campor" valign="top"><b>Cargo do Avaliador:</td>
</tr>
<tr>
	<td class="campo" valign="top">&nbsp;&nbsp;</td>
	<td class="campo" valign="top">&nbsp;&nbsp;</td>
</tr>
<tr><td class=campo colspan=2 valign="top" height="15" style="border-top:1px solid"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr><td class="campo" height="23" colspan="7" style="border-bottom:1px solid;border-top:1px solid" align="center"><b>PLANO DE AÇÃO DE DESENVOLVIMENTO INDIVIDUAL - PADI</td></tr>
<tr><td class="campor" height="" colspan="7" style="border-bottom:1px solid;border-top:1px solid" align="center">
	<b>GESTOR / AVALIADOR</b> - Estabeleça juntamente com o avaliado, um plano de ação visando melhorar os pontos apontados
	como deficientes e que precisam ser trabalhados.
	</td></tr>
<tr><td class=campo colspan="7" valign="top" height="4" style=""></td></tr>
<tr>
	<td class="campor" valign="middle" align="center">Tipo de ação sugerida:</td>
	<td class="campo" width="25" valign="middle" align="center" style="border:1px solid">1</td>
	<td class="campor" valign="middle">Simples orientações com acompanhamento</td>
	<td class="campo" width="25" valign="middle" align="center" style="border:1px solid">2</td>
	<td class="campor" valign="middle">Aprendizagem prática com pessoa indicada para <i>coaching</i> com monitoração direta</td>
	<td class="campo" width="25" valign="middle" align="center" style="border:1px solid">3</td>
	<td class="campor" valign="middle">Indicação de curso(s) para desenvolvimento ou aprimoramento (avaliar com a área de RH)</td>
</tr>
<tr><td class=campo colspan="7" valign="top" height="4" style="border-bottom:1px solid"></td></tr>
<tr><td class=campo colspan="7" valign="bottom" align="center" height="20" style="border-bottom:1px solid">
	<b>Indique o fator avaliado e que foi identificado como oportunidade de melhoria, necessitando desenvolvê-lo.</td></tr>
<tr><td class=campo colspan="7" valign="top" height="5" style="border-bottom:0px solid"></td></tr>
</table>

<%
for a=1 to 4
%>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campor" valign="middle" width="60" height="50"><b>FATOR:</b><br>(numeração)</td>
	<td class="campor" valign="top" width="35" height="" style="border-bottom:1px dotted"></td>
	<td class="campor" valign="top" width="5" height=""></td>
	<td class="campor" valign="top" style="border-bottom:1px dotted">Mencione a deficiência identificada:</td>
</tr>
<tr>
	<td class="campor" valign="midlle" width="60" height="50">Tipo de ação a ser desenvolvida:</td>
	<td class="campor" valign="top" width="35" height="" style="border-bottom:1px dotted"></td>
	<td class="campor" valign="top" width="5" height=""></td>
	<td class="campor" valign="top" style="border-bottom:1px dotted">Comente a sua realização:</td>
</tr>
<tr><td class=campo colspan="3" valign="top" height="3" style="border-bottom:0px solid"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campor" valign="middle" height="40"><b>Tempo necessário para apresentar melhorias:</b> (negociado com o avaliado)</td>
	<td class="campor" valign="top" style="border-bottom:1px dotted" width="100">&nbsp;</td>
	<td class="campor" valign="middle" ><b>Nova avaliação no fator, após a ação de desenvolvimento</b> (indique o mês da sua realização)</td>
	<td class="campor" style="border-bottom:1px dotted" width="100">&nbsp;</td>
</tr>
<tr><td class=campo colspan="3" valign="top" height="3" style="border-bottom:0px solid"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class="campor" valign="middle" height="40" nowrap><b>Resultado apresentado na nova avaliação:</b><br>(comente o resultado apresentado pelo colaborador)</td>
	<td class="campor" valign="top" style="border-bottom:1px dotted" width="490">&nbsp;</td>
</tr>
<tr><td class=campo colspan="3" valign="top" height="3" style="border-bottom:1px solid"></td></tr>
</table>
<%
next
%>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr><td class="campor" align="center" valign="bottom" height="40" style="">_______________________|__________<br>
                                                                           assinatura - avaliado  &nbsp;&nbsp;&nbsp; data</td>
	<td class="campor" align="center" valign="bottom" style="">_______________________|__________<br>
	                                                           assinatura - avaliador &nbsp;&nbsp;&nbsp; data</td>
</tr>
</table>

<!-- tabela pagina -->
</td></tr>
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