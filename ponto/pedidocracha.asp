<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a46")="N" or session("a46")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Solicitação de 2a.via Crachá</title>
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
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("B1")="" then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para emissão
<form method="POST" action="pedidocracha.asp" name="form">
<%
sqla="SELECT f.chapa, f.nome from corporerm.dbo.pfunc f " & _
"where f.codsituacao<>'D' order by f.nome"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=request.form("chapa")%>">
<select name="nome" class=a onchange="nome1()">
	<option value="0">Selecione o funcionário</option>
<%
rs.movefirst:do while not rs.eof
if request.form("chapa")=rs("chapa") then temps="selected" else temps=""
%>
	<option value="<%=rs("chapa")%>" <%=temps%> ><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<br><br>
<table border="1" cellpadding="2" cellspacing="2" style="border-collapse: collapse">
<tr>
	<td class=titulo>Data</td>
	<td class=titulo>Via</td>
	<td class=titulo>Motivo</td>
</tr>
<%
sql="select chapa, via, data, motivo from pedido_cracha where chapa='" & request.form("chapa") & "' order by data"
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
do while not rs.eof
%>
<tr>
	<td class=campo><%=rs("data")%></td>
	<td class=campo align="center"><%=rs("via")%></td>
	<td class=campo><%=rs("motivo")%></td>
</tr>
<%
via=rs("via")
rs.movenext:loop
end if
rs.close
%>
<tr>
	<td class=campo><input type="text" class="form_apt" name="data" size="8" value="<%=request.form("data")%>"> </td>
	<td class=campo><input type="text" class="form_box" name="via"  size="3" value="<%=via%>" style="text-align:center"> </td>
	<td class=campo><input type="text" class="form_apt" name="motivo" size="40" value="<%=request.form("motivo")%>" style="text-align:left"> </td>
</tr>

</table>
<br>
<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>
<%
end if

if request.form("B1")<>"" then

dtpedido=request.form("data")
via=request.form("via")
motivo=request.form("motivo")
chapa=request.form("chapa")
digito1=digito(chapa)
if request.form("data")<>"" and request.form("motivo")<>"" then
	sqlc="select chapa from pedido_cracha where chapa='" & chapa & "' and data='" & dtaccess(dtpedido) & "' "
	rs.Open sqlc, ,adOpenStatic, adLockReadOnly
	if rs.recordcount=0 then
		sqli="insert into pedido_cracha (chapa,via,data,motivo) select '" & chapa &"','" & via & "','" & dtaccess(dtpedido) & "','" & motivo & "' "
		conexao.execute sqli
	end if
	rs.close
end if

sql1="select f.chapa, f.nome, f.codsecao, s.descricao, p.sexo, p.apelido, f.codsindicato " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.psecao s, corporerm.dbo.ppessoa p " & _
"where f.codsecao=s.codigo and f.codpessoa=p.codigo and f.chapa='" & chapa & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
if rs("codsindicato")="03" then
	setorcargo="PROFESSOR"
	if rs("sexo")="F" then setorcargo=setorcargo & "A"
	tipo="p"
else
	setorcargo=rs("descricao")
	tipo="a"
end if
%>
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo height=70><img src="../images/logo_centro_universitario_unifieo_big.jpg" width="186" border="0" alt=""></td>
	<td class="campop" align="center"><b><font size=3>Ficha para emissão de crachá de identificação</td>
</tr>
<tr>
	<td class=campo valign="center" align="center" height=200 style="border: 1px solid #000000">
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="160" height="190">
		<tr><td style="border:4 double #000000" align="center">
		<img border="0" src="../func_foto.asp?chapa=<%=chapa%>" width="120">
		</td></tr></table>
	</td>
	<td class=campo valign="top">
		<table border="1" bordercolor="#000000" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=100%>
		<tr>
			<td class="campop" style="font-size:11pt;font-weight:bold" align="left" valign=middle height=30 width=100>Código</td>
			<td class="campop" style="font-size:12pt;font-weight:normal" align="left" valign=middle><%=rs("chapa")%> - <%=digito1%></td>
		</tr><tr>
			<td class="campop" style="font-size:11pt;font-weight:bold" align="left" valign=middle height=30>Nome</td>
			<td class="campop" style="font-size:12pt;font-weight:normal" align="left" valign=middle><%=rs("nome")%></td>
		</tr><tr>
			<td class=titulop colspan=2 style="font-size:11pt;font-weight:bold" align="center" valign=middle height=30> Dados para o crachá</td>
		</tr><tr>
			<td class="campop" style="font-size:11pt;font-weight:bold" align="left" valign=middle height=30>Nome</td>
			<td class="campop" style="font-size:12pt;font-weight:normal" align="left" valign=middle><%=rs("apelido")%></td>
		</tr><tr>
			<td class="campop" style="font-size:11pt;font-weight:bold" align="left" valign=middle height=30>Setor ou cargo</td>
			<td class="campop" style="font-size:12pt;font-weight:normal" align="left" valign=middle><%=setorcargo%></td>
		</tr><tr>
			<td class="campop" style="font-size:11pt;font-weight:bold" align="left" valign=middle height=50>Tipo</td>
			<td class="campop" style="font-size:12pt;font-weight:normal" align="left" valign=middle>
			<input type="radio" name="tipo" value="a" <%if tipo="a" then response.write "checked"%> >Vermelho (Administrativo)<br>
			<input type="radio" name="tipo" value="p" <%if tipo="p" then response.write "checked"%> >Azul (Acadêmico)
			</td>
		</tr>
		</table
	</td>
</tr>
<tr>
	<td class="campop" heigh=100 valign="middle" colspan=2><br><br><hr><br><br></td>
</tr>
<tr>
	<td class="campop" colspan=2 valign="top" height=300 style="border: 1px solid #000000">
		<table border="1" bordercolor="#000000" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=100%>
		<tr>
			<td class=titulo width=80>Data Pedido</td>	
			<td class=titulo width=40>Via</td>
			<td class=titulo>Motivo</td>
			<td class=titulo>Autorização de Desconto</td>
		</tr>
<%
sql="select chapa, via, data, motivo from pedido_cracha where chapa='" & request.form("chapa") & "' order by data"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
do while not rs2.eof
%>
<tr>
	<td class=campo align="center" height=30><%=rs2("data")%></td>
	<td class=campo align="center"><%=rs2("via")%></td>
	<td class=campo><%=rs2("motivo")%></td>
	<td class=campo>&nbsp;</td>
</tr>
<%
rs2.movenext:loop
end if
rs2.close
%>
	
	</td>
</tr>

<%
rs.close
set rs=nothing
%>
</table>
<%
set rs=nothing
set rs2=nothing
end if ' 

conexao.close
set conexao=nothing
%>
</body>
</html>