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
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request("codigo")="" then
	temp=1
	sqla="select chapa, nome, dataadmissao, dataadmissao+89 as venc from corporerm.dbo.pfunc " & _
	"where codsituacao<>'D' and codtipo='N' and dataadmissao+89>=getdate()-1 order by dataadmissao"
	sqla="select chapa, nome, dataadmissao, dataadmissao+89 as venc, DATAADMISSAO+44 as venc1 " & _
	", periodo=case when convert(integer,GETDATE()-dataadmissao)<=45 then '1' else '2' end " & _
	"from corporerm.dbo.pfunc where codsituacao<>'D' and codtipo='N' and dataadmissao+89>=getdate()-1 order by dataadmissao "
	sqla="select chapa, nome, dataadmissao, venc1, venc, periodo, dtbase=case periodo when '1' then venc1 else venc end " & _
	"from ( select chapa, nome, dataadmissao, dataadmissao+89 as venc, DATAADMISSAO+44 as venc1, periodo=case when convert(integer,GETDATE()-dataadmissao)<=45 then '1' else '2' end " & _
	"from corporerm.dbo.pfunc where codsituacao<>'D' and codtipo='N' and dataadmissao+89>=getdate()-1 " & _
	") z order by case periodo when '1' then venc1 else venc end, nome "
	'response.write sqla
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
	titulo=""
else
	temp=0
	sqla="SELECT f.CHAPA, f.NOME, f.DATAADMISSAO, f.DATAADMISSAO+89 AS Venc, f.DATAADMISSAO+44 AS Venc1, f.CODFUNCAO, c.NOME AS FUNCAO, f.CODSECAO, s.DESCRICAO AS SECAO, f1.NOME as chefe, p.SEXO " & _
"FROM ((((corporerm.dbo.PFUNC f INNER JOIN corporerm.dbo.PSECAO s ON f.CODSECAO=s.CODIGO) INNER JOIN corporerm.dbo.PFUNCAO c ON f.CODFUNCAO=c.CODIGO) " & _
"LEFT JOIN corporerm.dbo.PSUBSTCHEFE ch ON f.CODSECAO=ch.CODSECAO) LEFT JOIN corporerm.dbo.PFUNC AS f1 ON ch.CHAPASUBST=f1.CHAPA) LEFT JOIN corporerm.dbo.PPESSOA p ON f.CODPESSOA=p.CODIGO " & _
"WHERE f.CHAPA='" & request("codigo") & "' --and (getdate()>ch.datainicio and getdate()<ch.datafim or ch.datafim is null )"
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
	<td class=titulo align="center">Situação</td>
	<td class=titulo align="center">Venc.1ºper</td>
	<td class=titulo align="center">Venc.2ºper</td>
</tr>
<%
rs.movefirst
do while not rs.eof
if rs("periodo")="1" then situacao="<font color=green>Dentro do 1º período</font>"
if rs("periodo")="2" then situacao="<font color=blue>Dentro do 2º período</font>"
if now()>rs("venc") then situacao="<font color=red>Vencido</font>"
%>
<tr>
	<td class=campo style="border-bottom:1px dotted silver"><%=rs("chapa")%></td>
	<td class=campo style="border-bottom:1px dotted silver"><a class=r href="experiencia_avaliacao.asp?codigo=<%=rs("chapa")%>&periodo=<%=rs("periodo")%>"><%=rs("nome")%></a></td>
	<td class=campo style="border-bottom:1px dotted silver" align="center"><%=rs("dataadmissao") %></td>
	<td class=campo style="border-bottom:1px dotted silver" align="center" nowrap><%=situacao%></td>
	<td class=campo style="border-bottom:1px dotted silver" align="center"><%=rs("venc1") %></td>
	<td class=campo style="border-bottom:1px dotted silver" align="center"><%=rs("venc") %></td>
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
periodo=request("periodo")
if periodo=1 then vencper=rs("venc1") else vencper=rs("venc")

%>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr><td class="campop"><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width=225></td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><b>Of. <%=rs("chapa")%> - RH</b></td></tr>
	<tr><td class="campop" align="right">
	<input type="text" name="txt1" class="form_input" size="29" value="Osasco, <%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %>" style="font-size:10pt">
	</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><input type="text" name="txt0" class="form_input" size="5" value="À" style="font-size:10pt"><br>
	<input type="text" name="txt1" class="form_input" size="60" value="Sr(a). <%=rs("chefe")%>" style="font-size:10pt"><br>
	<input type="text" name="txt2" class="form_input" size="60" value="<%=rs("secao")%>" style="font-size:10pt"><br>
	</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">Em atendimento a sua solicitação, foi contratad<%=v1%> no dia <b><%=rs("dataadmissao")%></b> para prestar serviços 
	a esse departamento, <%=v1%> Sr<%=v2%>. <b><%=rs("nome")%></b>. Considerando que 
	<%if periodo=1 then response.write "o primeiro período do contrato de experiência terminará"%>
	<%if periodo=2 then response.write "seu contrato de experiência de três meses terminará"%>
	em <b><%=vencper%></b>, 
	<%if periodo=1 then%>
	solicitamos que V.Sa. se manifeste por escrito, na avaliação anexa sobre o desempenho d<%=v1%> referid<%=v1%> 
	funcionári<%=v1%> e o oriente caso não esteja apresentando ou tenha as exigências ideais para o cargo 
	e na hipótese de prorrogação a avaliação definitiva ocorrerá em 45 dias.
	<%end if%>
	<%if periodo=2 then%>
	solicitamos que V.Sa. se manifeste por escrito, na avaliação anexa sobre o desempenho d<%=v1%> referid<%=v1%> 
	funcionári<%=v1%>, informando-nos, com a maior brevidade possível, se <%=v1%> mesm<%=v1%> atende as exigências do cargo e preenche 
	os requisitos indispensáveis a sua admissão definitiva. Em resumo, se amolda aos padrões da FIEO.
	<%end if%>
	</td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tal informação, sigilosa e importantíssima, evitará que o serviço sofra solução de continuidade e tenhamos 
	despesas significativas, totalmente desnecessárias, com a dispensa d<%=v1%> funcionári<%=v1%> logo após a expiração do contrato 
	experimental.
	</td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Contando com a costumeira 
	colaboração de V.Sa. apresentamos nossas cordiais saudações.</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop">
	<p align="center" style="line-height: 25px">
	<input type="text" name="txt1" class="form_input" size="60" value="ROGERIO MATEUS DOS SANTOS ARAUJO" style="text-align:center;font-size:10pt;font-weight:bold"><br>
	<input type="text" name="txt2" class="form_input" size="60" value="Recursos Humanos" style="text-align:center;font-size:10pt;font-weight:bold"><br>
	</td></tr>
	<tr><td class="campop"></td></tr>
</table>
<DIV style="page-break-after:always"></DIV>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" style="border:1px solid #00000" align="center"><b>AVALIAÇÃO FUNCIONAL DURANTE O PERÍODO DE EXPERIÊNCIA</td>
	<td class="campop"><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width=175></td>
</tr>
<tr><td class=campo colspan=2 height=10  style="border-bottom:1px solid #000000"></td></tr>
<tr>
	<td class=campo colspan=2 align="center">
	<i>Com o objetivo de avaliar se o processo de adaptação e integração do novo colaborador está sendo adequadamente acompanhado <br>
	e se sua capacidade técnica e profissional estão correspondendo às expectativas	desejadas, solicitamos o preenchimento <br>
	deste formulário devolvendo-o impreterivelmente, até a data	estipulada.
	</td>
</tr>
<tr><td class=campo colspan=2 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campop" rowspan=2 align="center" valign="middle" style="border:1px solid"><b>1º PERÍODO</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Vencimento:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Devolver até:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Visto da área de Recursos Humanos</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-left:1px solid" align="center"><b><%=rs("venc1")%></td>
	<td class="campop" valign="top" style="border-left:1px solid" align="center"><b><%=rs("venc1")-10%></td>
	<td class="campop" valign="top" style="border-left:1px solid;border-right:1px solid"></td>
</tr>
<tr>
	<td class="campop" rowspan=2 align="center" valign="middle" style="border:1px solid"><b>2º PERÍODO</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Vencimento:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Devolver até:</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Visto da área de Recursos Humanos</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-left:1px solid;border-bottom:1px solid" align="center">
	<%if periodo=2 then%><b><%=rs("venc")%><%end if%>&nbsp;</td>
	<td class="campop" valign="top" style="border-left:1px solid;border-bottom:1px solid" align="center">
	<%if periodo=2 then%><b><%=rs("venc")-10%><%end if%>&nbsp;</td>
	<td class="campop" valign="top" style="border-left:1px solid;border-right:1px solid;border-bottom:1px solid"></td>
</tr>
<tr><td class=campo colspan=4 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">Nome do Funcionário</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid">RE</td>
	<td class="campor" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">Data de Admissão</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid"><b><%=rs("nome")%></td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid"><b><%=rs("chapa")%></td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"><b><%=rs("dataadmissao")%></td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Cargo</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Área / Depto</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid;border-right:1px solid">Superior Hierárquico-Nome</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid"><b><%=rs("funcao")%></td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid"><b><%=rs("secao")%></td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"><b><%=rs("chefe")%></td>
</tr>
<tr><td class=campo colspan=3 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td align="center">

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo align="center" valign="middle" rowspan=2>ITENS PARA AVALIAÇÃO</td>
	<td class=grupo align="center" valign="middle" colspan=4 style="border-right:2px solid">1º Período</td>
	<td class=grupo align="center" valign="middle" colspan=4>2º Período</td>
</tr>
<tr>
	<td class="campor" align="center" valign="middle">ÓTIMO</td>
	<td class="campor" align="center" valign="middle">BOM</td>
	<td class="campor" align="center" valign="middle">REGULAR</td>
	<td class="camposs" align="center" valign="middle" style="border-right:2px solid">ABAIXO DO<br>ESPERADO</td>
	<td class="campor" align="center" valign="middle">ÓTIMO</td>
	<td class="campor" align="center" valign="middle">BOM</td>
	<td class="campor" align="center" valign="middle">REGULAR</td>
	<td class="camposs" align="center" valign="middle">ABAIXO DO<br>ESPERADO</td>
</tr>
<%
chapa=rs("chapa")
sql="select distinct chapa from iAvExp where chapa='" & chapa & "'"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
	existe=rs2.recordcount
rs2.close
if existe=0 then
	sqli="insert into iAvExp (chapa, idItem, create_user, create_data) select '" & chapa & "', idItem, '" & session("usuariomaster") & "', GETDATE() from iAvExpItens"
	conexao.execute sqli
	existe=1
end if

sqla="select a.chapa, i.Tipo, i.Ordem, a.idItem, i.Descricao, a.P1Aval, a.P2Aval, a.Anotacao " & _
"from iAvExp a inner join iAvExpItens i on i.iditem=a.iditem " & _
"where a.chapa='" & chapa & "' and i.Tipo='IA' order by i.Tipo, i.Ordem"
rs2.Open sqla, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
p1=rs2("P1Aval"):p2=rs2("P2Aval")
%>
<tr>
	<td class=campo><%=rs2("Descricao")%></td>
	<td class=Campo align="center"><%if p1="O" then response.write "X"%></td>
	<td class=Campo align="center"><%if p1="B" then response.write "X"%></td>
	<td class=Campo align="center"><%if p1="R" then response.write "X"%></td>
	<td class=Campo align="center" style="border-right:2px solid"><%if p1="A" then response.write "X"%></td>
	<td class=Campo align="center"><%if p2="O" then response.write "X"%></td>
	<td class=Campo align="center"><%if p2="B" then response.write "X"%></td>
	<td class=Campo align="center"><%if p2="R" then response.write "X"%></td>
	<td class=Campo align="center"><%if p2="A" then response.write "X"%></td>
</tr>
<%
rs2.movenext
loop
rs2.close
%>
</table>
</td></tr></table>
<%
sqlp1="select a.chapa, i.Tipo, i.Ordem, a.idItem, i.Descricao, a.P1Aval, a.P2Aval, a.Anotacao from iAvExp a inner join iAvExpItens i on i.iditem=a.iditem where a.chapa='" & chapa & "' and i.tipo='P1'"
rs2.Open sqlp1, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
	if rs2("descricao")="Decisao" 					then p1decisao		=rs2("p1aval")
	if rs2("descricao")="Justificar"				then p1justificar	=rs2("p1aval")
	if rs2("descricao")="Pontos a serem melhorados"	then p1pontos		=rs2("p1aval")
	if rs2("descricao")="Por meio de"				then p1pormeio		=rs2("p1aval")
	if rs2("descricao")="Treinamento em"			then p1treinamento	=rs2("p1aval")
	if rs2("descricao")="Data"						then p1data			=rs2("p1aval")
	if rs2("descricao")="Avaliador"					then p1avaliador	=rs2("p1aval")
rs2.movenext:loop
rs2.close
%>

<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=grupo align="center" valign="middle" rowspan=6>
	1<br>º<br> <br>P<br>E<br>R<br>Í<br>O<br>D<br>O</td>
	<td class=campo height="25" valign="middle" style="border-top:1px solid;border-right:1px solid;border-bottom:1px solid">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Decisão: </b>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;<b><%if p1decisao="P" then response.write "X"%></b>&nbsp;&nbsp;] Prorrogar 
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;<b><%if p1decisao="D" then response.write "X"%></b>&nbsp;&nbsp;] Dispensar
	</td>
</tr>
<tr><td class="campor" height="25" valign="top" style="border-bottom:1px dotted;border-right:1px solid">Justificar</td></tr>
<tr><td class=campo style="border-bottom:1px solid;border-right:1px solid"><%=p1justificar%>&nbsp;</td></tr>
<tr><td class="campor" height="25" valign="top" style="border-bottom:1px dotted;border-right:1px solid">Pontos a serem melhorados ou considerados</td></tr>
<tr><td class=campo style="border-bottom:1px solid;border-right:1px solid"><%=p1pontos%>&nbsp;</td></tr>
<tr><td class=campo height="25" valign="middle" style="border-bottom:1px solid;border-right:1px solid">
	&nbsp;&nbsp;&nbsp;<b>Por meio de </b>
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;<b><%if p1pormeio="A" then response.write "X"%></b>&nbsp;&nbsp;] Acompanhamento/Orientação 
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;<b><%if p1pormeio="T" then response.write "X"%></b>&nbsp;&nbsp;] Treinamento em 
	<%if p1treinamento<>"" then response.write p1treinamento else response.write "______________________________________________"%>
	</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Data da Devolução</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid;border-right:1px solid">Visto do Superior Hierárquico</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid"><%=p1data%>&nbsp;</td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"><%=p1avaliador%>&nbsp;</td>
</tr>
<tr><td class=campo colspan=2 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>

<%
sqlp2="select a.chapa, i.Tipo, i.Ordem, a.idItem, i.Descricao, a.P1Aval, a.P2Aval, a.Anotacao from iAvExp a inner join iAvExpItens i on i.iditem=a.iditem where a.chapa='" & chapa & "' and i.tipo='P1'"
rs2.Open sqlp2, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
	if rs2("descricao")="Decisao" 					then p2decisao		=rs2("p2aval")
	if rs2("descricao")="Justificar"				then p2justificar	=rs2("p2aval")
	if rs2("descricao")="Pontos a serem melhorados"	then p2pontos		=rs2("p2aval")
	if rs2("descricao")="Por meio de"				then p2pormeio		=rs2("p2aval")
	if rs2("descricao")="Treinamento em"			then p2treinamento	=rs2("p2aval")
	if rs2("descricao")="Data"						then p2data			=rs2("p2aval")
	if rs2("descricao")="Avaliador"					then p2avaliador	=rs2("p2aval")
rs2.movenext:loop
rs2.close
%>


<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=grupo align="center" valign="middle" rowspan=6>
	2<br>º<br> <br>P<br>E<br>R<br>Í<br>O<br>D<br>O</td>
	<td class=campo height="25" valign="middle" style="border-top:1px solid;border-right:1px solid;border-bottom:1px solid">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Decisão: </b>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;<b><%if p2decisao="P" then response.write "X"%></b>&nbsp;&nbsp;] Efetivar 
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;<b><%if p2decisao="D" then response.write "X"%></b>&nbsp;&nbsp;] Dispensar
	</td>
</tr>
<tr><td class="campor" height="25" valign="top" style="border-bottom:1px dotted;border-right:1px solid">Justificar</td></tr>
<tr><td class=campo style="border-bottom:1px solid;border-right:1px solid"><%=p2justificar%>&nbsp;</td></tr>
<tr><td class="campor" height="25" valign="top" style="border-bottom:1px dotted;border-right:1px solid">Pontos a serem melhorados ou considerados</td></tr>
<tr><td class=campo style="border-bottom:1px solid;border-right:1px solid"><%=p2pontos%>&nbsp;</td></tr>
<tr><td class=campo height="25" valign="middle" style="border-bottom:1px solid;border-right:1px solid">
	&nbsp;&nbsp;&nbsp;<b>Por meio de </b>
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;<b><%if p2pormeio="A" then response.write "X"%></b>&nbsp;&nbsp;] Acompanhamento/Orientação 
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;<b><%if p2pormeio="T" then response.write "X"%></b>&nbsp;&nbsp;] Treinamento em 
	<%if p2treinamento<>"" then response.write p2treinamento else response.write "______________________________________________"%>
	</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Data da Devolução</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid;border-right:1px solid">Visto do Superior Hierárquico</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid"><%=p2data%>&nbsp;</td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid"><%=p2avaliador%>&nbsp;</td>
</tr>
<tr><td class=campo colspan=2 height=10 style="border-bottom:0px solid #000000"></td></tr>
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