<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a40")="N" or session("a40")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Declaração de Vínculo para o INSS</title>
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
set rsi=server.createobject ("ADODB.Recordset")
Set rsi.ActiveConnection = conexao
set rsl=server.createobject ("ADODB.Recordset")
Set rsl.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
teste=0

if request.form<>"" then session("40quem")=request.form("quem"):session("40qfuncao")=request.form("funcao")
if request("codigo")<>"" or request.form<>"" then
	if request.form="" then temp=request("codigo")
	if request("codigo")="" then temp=request.form("chapa")
	if request.form("anome")="" then
		if isnumeric(temp) then
			info=1
			temp=numzero(temp,5)
			sqlb="AND f.CHAPA='" & temp & "' "
		else
			info=2
			sqlb="AND f.nome like '%" & temp & "%' order by f.nome"
		end if
		sqla="SELECT F.CHAPA, F.NOME, C.NOME AS FUNCAO, P.CARTEIRATRAB, P.SERIECARTTRAB, " & _
		"F.DATAADMISSAO, f.datademissao, P.SEXO, f.codsituacao, s.descricao as secao " & _
		"FROM corporerm.dbo.PFUNC F, corporerm.dbo.PPESSOA P, corporerm.dbo.PFUNCAO C, corporerm.dbo.psecao s " & _
		"WHERE F.CODPESSOA = P.CODIGO AND F.CODFUNCAO = C.CODIGO and f.codsecao=s.codigo "
		sql1=sqla & sqlb
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		session("40chapa")=rs("chapa")
		session("40nome")=rs("nome")
		session("40admissao")=rs("dataadmissao")
		session("40demissao")=rs("datademissao")
		session("40ctps")=rs("carteiratrab")
		session("40serie")=rs("seriecarttrab")
		session("40funcao")=rs("funcao")
		session("40sexo")=rs("sexo")
		session("40situacao")=rs("codsituacao")
		if rs.recordcount>1 and session("cartateto")<>"L" then temp=2
		else 'request.form("anome")<>""
		session("40nome")=ucase(request.form("anome"))
		session("40admissao")=request.form("aadmissao")
		session("40demissao")=request.form("ademissao")
		session("40ctps")=request.form("actps")
		session("40serie")=request.form("aserie")
		session("40funcao")=ucase(request.form("afuncao"))
		session("40sexo")=request.form("asexo")
		session("40situacao")="D"
	end if 'request.form("anome")
	temp=0
	'if request.form("anome")<>"" and session("cartateto")<>"L" then temp=2
else
	temp=1
end if

if temp=1 then
	session("cartateto")="F"
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção do funcionário para emissão de declaração
<form method="POST" action="declaracaoinss.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo colspan=2>Opções para a declaração</td>
</tr>
<tr><td class=fundo>Chapa/Nome</td><td class=fundo><input type="text" name="chapa" size="20" class="form_box" value="<%=session("chapa")%>"></td>
</tr>
<tr><td class=titulo>Quem assina:</td><td class=fundo><input type="text" name="quem" size="50" value="ROGERIO MATEUS DOS SANTOS ARAUJO"></td>
</tr>
<tr><td class=titulo>Função:</td><td class=fundo><input type="text" name="funcao" size="30" value="Supervisor de Recursos Humanos"></td>
</tr>
<tr><td height=5 colspan=2></td></tr>
<tr>
	<td class="campol" colspan=2><b>Para ex-funcionários antigos não cadastrados<br>no banco de dados</td>
</tr>
<tr><td class=fundo>Nome</td>
	<td class=fundo><input type="text" name="anome" value="" size="40"></td></tr>
<tr><td class=fundo>Data Admissão</td>
	<td class=fundo><input type="text" name="aadmissao" value="" size="8"></td></tr>
<tr><td class=fundo>Data Demissão</td>
	<td class=fundo><input type="text" name="ademissao" value="" size="8"></td></tr>
<tr><td class=fundo>CTPS</td>
	<td class=fundo><input type="text" name="actps" value="" size="6"></td></tr>
<tr><td class=fundo>Série</td>
	<td class=fundo><input type="text" name="aserie" value="" size="4"></td></tr>
<tr><td class=fundo>Função</td>
	<td class=fundo><input type="text" name="afuncao" value="" size="20"></td></tr>
<tr><td class=fundo>Sexo</td>
	<td class=fundo><select name="asexo"><option value="M">Masculino</option>
	<option value="F">Feminino</option></select></td></tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0">
<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>
<%
elseif temp=0 then
session("cartateto")="C"
'if request.form<>"" then
if session("40sexo")="F" then v1="a" else v1="o"
if session("40sexo")="F" then v2="a" else v2=""
if session("40sexo")="F" then v3="à" else v3="ao"
%>
<div align="center"><center>
<table border="0" cellpadding="5" width="620" cellspacing="0" height="1000">
<!-- linha da aguia -->
<tr><td height=112><img src="../images/aguia.jpg" border="0" width="236" height="111" alt=""></td></tr>
<!-- linha intermediaria -->
<tr><td height="20">&nbsp;</td></tr>
<!-- linha declaracao -->
<tr><td height=50 valign="center" align="center"><b><font size="4">D E C L A R A Ç Ã O</font></b></td></tr>
<!-- linha intermediaria -->
<tr><td height="20">&nbsp;</td></tr>
<!-- corpo da declaracao -->
<tr><td height=100% valign=top>
<%
if session("40situacao")="D" then
	demissao=day(session("40demissao")) & " de " & monthname(month(session("40demissao"))) & " de " & year(session("40demissao"))
	texto1=" esteve ":texto2=" de ":texto3=" a " & demissao
else
	texto1=" está ":texto2=" de ":texto3=" até hoje"
end if
admissao=day(session("40admissao")) & " de " & monthname(month(session("40admissao"))) & " de " & year(session("40admissao"))
%>	
	<p>&nbsp;</p>
	<p align="justify"><font size="3"><%=session("40quem")%>, na qualidade de <%=session("40qfuncao")%>, da empresa FUNDAÇÃO 
	INSTITUTO DE ENSINO PARA OSASCO, declara que <%=v1%> Sr<%=v2%>. <%=session("40nome")%>,
	portador<%=v2%> da CTPS nº <%=session("40ctps")%> série <%=session("40serie")%>,<%=texto1%>
	a serviço da supra mencionada empresa, no período <%=texto2%>
	<%=admissao%><%=texto3%>, exercendo a função de <%=session("40funcao")%>, tendo sido estes elementos extraídos dos LIVROS, 
	FICHAS, FOLHAS DE PAGAMENTO, CARTÕES DE PONTO, etc., existentes em nossos arquivos e que desde já ficam à disposição 
	do INSS no seguinte endereço: Av. Franz Voegelli, 300 - Osasco - São Paulo.</font></p>

	<p align="justify"><font size="3">Declara, também, estar ciente de que se em qualquer época, ficar provada a 
	inexatidão destas declarações, estará incursa nos artigos 171 e 299 do Código Penal.</font></p>
	<p align="justify">&nbsp;</p>

<!-- tabela data e assinatura -->
	<table border="0" cellpadding="0" width="100%" cellspacing="0">
	<tr>
		<td width="50%" valign="top">
<%if day(now())=1 then dia="1º" else dia=day(now())%>
		<p><font size="3">Osasco,&nbsp;<%=dia & " de " & monthname(month(now())) & " de " & year(now()) %></font></p>
		<p><font size="3">Atenciosamente</font></p>
		<p>&nbsp;</p>
		<p><font size="3">_____________________________________<br>
		<%=session("40quem")%><br><%=session("40qfuncao")%></font></p>
		</td>
		<!-- carimbo cgc -->
<%if teste=1 then %>
		<td width="50%" valign="top">&nbsp;
		<div align="center"><center>
		<table border="0" cellpadding="0" width="240" cellspacing="0">
		<tr><td width="1" style="border-left-style: solid; border-left-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240" rowspan="2">
				<p align="center"><b><font size="4" color="#808080">73.063.166/0001-20</font></b></td>
			<td width="1" style="border-right-style: solid; border-right-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
		<tr><td width="1"></td><td width="1"></td></tr>
		<tr><td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td><td width="1"></td></tr>
		<tr><td width="1"></td><td width="240" align="center">
				<b><font color="#808080">FUNDAÇÃO INSTITUTO DE<br>ENSINO PARA OSASCO</font></b></td>
			<td width="1"></td></tr>
		<tr><td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td><td width="1"></td></tr>
		<tr><td width="1">&nbsp;</td><td width="240" rowspan="2" align="center">
				<font color="#808080">Rua Narciso Sturlini, 883<br>
				Jd. Umuarama - CEP 06018-903<br>OSASCO - SP</font></td><td width="1"></td></tr>
		<tr><td width="1" style="border-left-style: solid; border-left-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="1" style="border-right-style: solid; border-right-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
		</table></center></div>
		<p>&nbsp;
		</td>
<%end if%>
		</tr>
	</table>
<!-- fim tabela assinatura/data -->

	</td>
</tr>
<!-- linha intermediaria -->
<tr><td height="20">&nbsp;</td></tr>
<tr><td height="1"><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
<tr><td height="1"><font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000 - C.N.P.J. 73.063.166/0001-20</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP 06020-190 - Fone: (011) 3651-9999 - C.N.P.J. 73.063.166/0003-92</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 1743 - Osasco - SP - CEP 06020-190 - Fone: (011) 3654-0655 - C.N.P.J. 73.063.166/0004-73</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Caixa Postal - ACF - Jd. Ipê - nº 2530 - Osasco - SP - CEP 06053-990</font></td></tr>
</table>
</center></div>

<%
'rs.close
'set rs=nothing

elseif temp=2 then
session("cartateto")="L"
%>
<!-- mostrar funcionarios e as contribuições -->
<table border="1" cellpadding="0" width="550" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>&nbsp;Chapa</td>
	<td class=titulo>&nbsp;Nome</td>
    <td class=titulo>&nbsp;Situacao</td>
</tr>
<%
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo>&nbsp;<%=rs("chapa")%></td>
	<td class=campo>&nbsp;<a href="declaracaoinss.asp?codigo=<%=rs("chapa")%>"><%=rs("nome")%></a></td>
	<td class=campo>&nbsp;<%=rs("codsituacao")%></td>
</tr>
<%
rs.movenext
loop
%>
</table>
<%
rs.close
set rs=nothing
end if ' temps
conexao.close
set conexao=nothing
%>
</body>
</html>