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
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function nome1() { form.chapa.value=form.nome.value;form.submit(); }
function chapa1() { form.nome.value=form.chapa.value;form.submit(); }
--></script>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form("Salvar")<>"" then
	for a=1 to 23
		sql="update iAvExp set P1Aval='" & request.form("1P"&a) & "', P2Aval='" & request.form("2P"&a) & "' " & _
		"where idItem=" & a & " and chapa='" & request.form("chapa") & "' "
		conexao.execute sql
	next
end if
'if request.form("nome")<>"" and request.form("chapa")="" then valor_chapa=request.form("nome")
'if request.form("chapa")<>"" then valor_chapa=request.form("chapa")

%>
<form name="form" action="exp_AvExp.asp" method="post">

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=grupo>Cadastro da Avaliação dos Períodos de Experiência</td></tr>
</table>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="left">Nome do Funcionário</td>
	<td class=titulo align="center">Admissão</td>
	<td class=titulo align="left">Função</td>
</tr>
<tr>
	<td class=campo><input type="text" name="chapa" size="5" maxlength="5" value="<%=request.form("chapa")%>" onchange="chapa1()"></td>
	<td class=campo><select name="nome" class=a onchange="nome1()">
		<option value="0"> Selecione o funcionário</option>
<%
admissao="":funcao=""
sql="select p.chapa, p.nome, p.admissao, p.funcao from qry_funcionarios p where  p.codsituacao<>'D' and p.codsindicato<>'03' and p.codtipo='N' " & _
"order by p.nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
if request.form("chapa")=rs("chapa") or request.form("nome")=rs("chapa") then
	tempc="selected" 
	admissao=rs("admissao")
	funcao=rs("funcao")
else 
	tempc=""
end if
%>
		<option value="<%=rs("chapa")%>" <%=tempc%>> <%=rs("nome")%></option>
<%
rs.movenext:loop
%>
	</select>
	</td>
	<td class=campo align="center"><%=admissao%></td>
	<td class=campo><%=funcao%></td>
</tr>
<tr><td class=campo colspan=4 height=15></td></tr>
<%
rs.close
%>
</table>
<%
if request.form("chapa")<>"" then
chapa=request.form("chapa")
sql="select distinct chapa from iAvExp where chapa='" & chapa & "'"
rs.Open sql, ,adOpenStatic, adLockReadOnly
existe=rs.recordcount
rs.close
if existe=0 then
	sqli="insert into iAvExp (chapa, idItem, create_user, create_data) select '" & chapa & "', idItem, '" & session("usuariomaster") & "', GETDATE() from iAvExpItens"
	conexao.execute sqli
	existe=1
end if
if existe=1 then
	sqla="select a.chapa, i.Tipo, i.Ordem, a.idItem, i.Descricao, a.P1Aval, a.P2Aval, a.Anotacao " & _
	"from iAvExp a inner join iAvExpItens i on i.iditem=a.iditem " & _
	"where a.chapa='" & chapa & "' and i.Tipo='IA' order by i.Tipo, i.Ordem"
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
	rs.movefirst
%>
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
do while not rs.eof
p1=rs("P1Aval"):p2=rs("P2Aval")
%>
<tr>
	<td class=campo><%=rs("Descricao")%></td>
	<td class=Campo align="center"><input type="radio" name="1P<%=rs("idItem")%>" value="O" <%if p1="O" then response.write "checked"%> ></td>
	<td class=Campo align="center"><input type="radio" name="1P<%=rs("idItem")%>" value="B" <%if p1="B" then response.write "checked"%>></td>
	<td class=Campo align="center"><input type="radio" name="1P<%=rs("idItem")%>" value="R" <%if p1="R" then response.write "checked"%>></td>
	<td class=Campo align="center" style="border-right:2px solid"><input type="radio" name="1P<%=rs("idItem")%>" value="A" <%if p1="A" then response.write "checked"%>></td>
	<td class=Campo align="center"><input type="radio" name="2P<%=rs("idItem")%>" value="O" <%if p2="O" then response.write "checked"%>></td>
	<td class=Campo align="center"><input type="radio" name="2P<%=rs("idItem")%>" value="B" <%if p2="B" then response.write "checked"%>></td>
	<td class=Campo align="center"><input type="radio" name="2P<%=rs("idItem")%>" value="R" <%if p2="R" then response.write "checked"%>></td>
	<td class=Campo align="center"><input type="radio" name="2P<%=rs("idItem")%>" value="A" <%if p2="A" then response.write "checked"%>></td>
</tr>
<%
rs.movenext
loop
rs.close
%>
</table>

<%
sqlp1="select a.chapa, i.Tipo, i.Ordem, a.idItem, i.Descricao, a.P1Aval, a.P2Aval, a.Anotacao from iAvExp a inner join iAvExpItens i on i.iditem=a.iditem where a.chapa='" & chapa & "' and i.tipo='P1'"
rs.Open sqlp1, ,adOpenStatic, adLockReadOnly
do while not rs.eof
	if rs("descricao")="Decisao" 					then p1decisao		=rs("p1aval")
	if rs("descricao")="Justificar"					then p1justificar	=rs("p1aval")
	if rs("descricao")="Pontos a serem melhorados"	then p1pontos		=rs("p1aval")
	if rs("descricao")="Por meio de"				then p1pormeio		=rs("p1aval")
	if rs("descricao")="Treinamento em"				then p1treinamento	=rs("p1aval")
	if rs("descricao")="Data"						then p1data			=rs("p1aval")
	if rs("descricao")="Avaliador"					then p1avaliador	=rs("p1aval")
rs.movenext:loop
rs.close
%>

<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=grupo align="center" valign="middle" rowspan=6>
	1<br>º<br> <br>P<br>E<br>R<br>Í<br>O<br>D<br>O</td>
	<td class=campo height="25" valign="middle" style="border-top:1px solid;border-right:1px solid;border-bottom:1px solid">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Decisão: </b>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;
	<input type="radio" name="1P17" value="P" <%if p1decisao="P" then response.write "checked"%>>&nbsp;&nbsp;] Prorrogar 
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;
	<input type="radio" name="1P17" value="D" <%if p1decisao="D" then response.write "checked"%>>&nbsp;&nbsp;] Dispensar
	</td>
</tr>
<tr><td class="campor" height="20" valign="middle" style="border-bottom:1px solid;border-right:1px solid">
	Justificar <input type="text" name="1P18" size="90" value="<%=p1justificar%>"></td></tr>
<tr><td class="campor" height="20" valign="middle" style="border-bottom:1px solid;border-right:1px solid">
	Pontos a serem melhorados ou considerados <input type="text" name="1P19" size="70" value="<%=p1pontos%>"></td></tr>
<tr><td class=campo height="25" valign="middle" style="border-bottom:1px solid;border-right:1px solid">
	&nbsp;&nbsp;&nbsp;<b>Por meio de </b>
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;
	<input type="radio" name="1P20" value="A" <%if p1pormeio="A" then response.write "checked"%>>&nbsp;&nbsp;] Acompanhamento/Orientação 
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;
	<input type="radio" name="1P20" value="T" <%if p1pormeio="T" then response.write "checked"%>>&nbsp;&nbsp;] Treinamento em
	<input type="text" name="1P21" size="30" value="<%=p1treinamento%>">
	</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Data da Devolução</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid;border-right:1px solid">Visto do Superior Hierárquico</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid">
	<input type="text" name="1P22" size="20" value="<%=p1data%>"></td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid">
	<input type="text" name="1P23" size="40" value="<%=p1avaliador%>"></td>
</tr>
<tr><td class=campo colspan=2 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>

<%
sqlp2="select a.chapa, i.Tipo, i.Ordem, a.idItem, i.Descricao, a.P1Aval, a.P2Aval, a.Anotacao from iAvExp a inner join iAvExpItens i on i.iditem=a.iditem where a.chapa='" & chapa & "' and i.tipo='P1'"
rs.Open sqlp2, ,adOpenStatic, adLockReadOnly
do while not rs.eof
	if rs("descricao")="Decisao" 					then p2decisao		=rs("p2aval")
	if rs("descricao")="Justificar"					then p2justificar	=rs("p2aval")
	if rs("descricao")="Pontos a serem melhorados"	then p2pontos		=rs("p2aval")
	if rs("descricao")="Por meio de"				then p2pormeio		=rs("p2aval")
	if rs("descricao")="Treinamento em"				then p2treinamento	=rs("p2aval")
	if rs("descricao")="Data"						then p2data			=rs("p2aval")
	if rs("descricao")="Avaliador"					then p2avaliador	=rs("p2aval")
rs.movenext:loop
rs.close
%>

<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=grupo align="center" valign="middle" rowspan=6>
	2<br>º<br> <br>P<br>E<br>R<br>Í<br>O<br>D<br>O</td>
	<td class=campo height="25" valign="middle" style="border-top:1px solid;border-right:1px solid;border-bottom:1px solid">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Decisão: </b>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;
	<input type="radio" name="2P17" value="P" <%if p2decisao="P" then response.write "checked"%>>&nbsp;&nbsp;] Prorrogar 
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;
	<input type="radio" name="2P17" value="D" <%if p2decisao="D" then response.write "checked"%>>&nbsp;&nbsp;] Dispensar
	</td>
</tr>
<tr><td class="campor" height="20" valign="middle" style="border-bottom:1px solid;border-right:1px solid">
	Justificar <input type="text" name="2P18" size="90" value="<%=p2justificar%>"></td></tr>
<tr><td class="campor" height="20" valign="middle" style="border-bottom:1px solid;border-right:1px solid">
	Pontos a serem melhorados ou considerados <input type="text" name="2P19" size="70" value="<%=p2pontos%>"></td></tr>
<tr><td class=campo height="25" valign="middle" style="border-bottom:1px solid;border-right:1px solid">
	&nbsp;&nbsp;&nbsp;<b>Por meio de </b>
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;
	<input type="radio" name="2P20" value="A" <%if p2pormeio="A" then response.write "checked"%>>&nbsp;&nbsp;] Acompanhamento/Orientação 
	&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;
	<input type="radio" name="2P20" value="T" <%if p2pormeio="T" then response.write "checked"%>>&nbsp;&nbsp;] Treinamento em
	<input type="text" name="2P21" size="30" value="<%=p2treinamento%>">
	</td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid">Data da Devolução</td>
	<td class="campor" valign="top" style="border-top:0px solid;border-left:1px solid;border-right:1px solid">Visto do Superior Hierárquico</td>
</tr>
<tr>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid">
	<input type="text" name="2P22" size="20" value="<%=p2data%>"></td>
	<td class="campop" valign="top" style="border-bottom:1px solid;border-left:1px solid;border-right:1px solid">
	<input type="text" name="2P23" size="40" value="<%=p2avaliador%>"></td>
</tr>
<tr><td class=campo colspan=2 height=10 style="border-bottom:0px solid #000000"></td></tr>
</table>

<%
end if 'existe=1
end if 'request.form("chapa")<>""
%>

<%
set rs=nothing
conexao.close
set conexao=nothing
%>
</table>
<input type="submit" name="Salvar" Value="Salvar">
</form>
</body>
</html>