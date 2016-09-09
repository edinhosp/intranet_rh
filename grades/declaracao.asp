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
<title>Declaração de Disciplinas</title>
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
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
teste=0

if request("codigo")<>"" or request.form<>"" then
	if request.form="" then temp=request("codigo")
	if request("codigo")="" then temp=request.form("chapa")
	if isnumeric(temp) then
		info=1
		temp=numzero(temp,5)
		sqlb="AND f.CHAPA='" & temp & "' "
	else
		info=2
		sqlb="AND f.nome like '%" & temp & "%' order by f.nome"
	end if
	sqla="SELECT F.CHAPA, F.NOME, C.NOME AS FUNCAO, P.CARTEIRATRAB, P.SERIECARTTRAB, f.codtipo, P.cartidentidade, " & _
	"F.DATAADMISSAO, f.datademissao, P.SEXO, f.codsituacao, s.descricao as secao, f.codsecao, f.salario " & _
	"FROM corporerm.dbo.PFUNC F, corporerm.dbo.PPESSOA P, corporerm.dbo.PFUNCAO C, corporerm.dbo.psecao s " & _
	"WHERE F.CODPESSOA = P.CODIGO AND F.CODFUNCAO = C.CODIGO and f.codsecao=s.codigo and codtipo='A' "
	
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	session("chapa")=rs("chapa")
	chapa=rs("chapa")
	nome=rs("nome"):admissao=rs("dataadmissao"):funcao=rs("funcao")
	temp=0
	if rs.recordcount>0 and session("cartateto")<>"L" then temp=2
	session("40parag")=session("40parag")
	if request.form("parag")="" then session("40parag")=request("parag")
	if request.form("parag")<>"" then session("40parag")=request.form("parag")
else
	temp=1
end if

if temp=1 then
	session("cartateto")="F"
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de pessoa para emissão de declaração
<form method="POST" action="declaracao.asp" name="form">
<p style="margin-top: 0; margin-bottom: 0">
Chapa/Nome <input type="text" name="chapa" size="20" class="form_box" value="<%=session("chapa")%>">
</p>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Opções para a declaração</td>
</tr>
<tr>
	<td class=campo>
	<p style="margin-top: 0; margin-bottom: 1">
	</td>
</tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0">
<input type="submit" value="Pesquisar" name="B1" class="button"></p>
<textarea name="parag" cols="50" rows="5" class=form_input>

</textarea>

</form>
<%
elseif temp=0 then
session("cartateto")="C"
'if request.form<>"" then
if rs("sexo")="F" then v1="a" else v1="o"
if rs("sexo")="F" then v2="a" else v2=""
if rs("sexo")="F" then v3="à" else v3="ao"
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
if rs("codsituacao")="D" or (rs("codsituacao")="A" and rs("datademissao")<>"") then
	texto1=" prestou serviços ":texto2=" de ":texto3=" a " & rs("datademissao")
else
	texto1=" presta serviços ":texto2=" desde ":texto3=""
end if
%>	
	<p>&nbsp;</p>
	<p align="justify"><font size="3">Declaramos aos orgãos interessados, que <%=v1%> Sr<%=v2%>. <%=rs("nome")%>,
	portador<%=v2%> <%if rs("carteiratrab")<>"" or not isnull(rs("carteiratrab")) then%> da CTPS nº <%=rs("carteiratrab")%> / <%=rs("seriecarttrab")%>
	<%else%> do R.G. nº <%=rs("cartidentidade")%><%end if%>
	, prestou serviços nesta Instituição de Ensino Superior, ministrando módulos de duração programada em cursos
	de pós-graduação.</font></p>
<%
if session("40parag")<>"" then
%>
	<p align="justify"><font size="3"><%=session("40parag")%></font></p>
<%
end if
%>

<!-- inicio quadro da pos-graduação -->
<%
sqld="SELECT chapa1, curso, materia, inicio, termino, Sum(ta) AS aulas, periodo=case when inicio is null then perlet else convert(nvarchar,day(inicio))+'/'+convert(nvarchar,month(inicio))+'/'+convert(char(4),year(inicio)) + ' a ' + convert(nvarchar,day(termino))+'/'+convert(nvarchar,month(termino))+'/'+convert(char(4),year(termino)) end FROM g5ch g inner join g2cursoeve c on c.coddoc=g.coddoc " & _
"GROUP BY chapa1, curso, materia, inicio, termino, case when inicio is null then perlet else convert(nvarchar,day(inicio))+'/'+convert(nvarchar,month(inicio))+'/'+convert(char(4),year(inicio)) + ' a ' + convert(nvarchar,day(termino))+'/'+convert(nvarchar,month(termino))+'/'+convert(char(4),year(termino)) end " & _
"HAVING chapa1='" & chapa & "' ORDER BY inicio, chapa1, curso, materia;"
rs1.Open sqld, ,adOpenStatic, adLockReadOnly
if rs1.recordcount>0 then

%>
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="620" >
<tr><td class=campo valign="top" height="99%">
<!-------------- -->

<table border="0" cellpadding="2" cellspacing="0" width="620" style="border-collapse: collapse">
<tr><td class=titulo align="center">Curso</td>
	<td class=titulo align="center">Disciplina</td>
	<td class=titulo align="center">Horas</td>
	<td class=titulo align="center">Período</td>
</tr>
<%
rs1.movefirst:do while not rs1.eof
%>
<tr><td class=campo align="left"   style="border:1px solid #000000"><%=rs1("curso")%></td>
	<td class=campo align="left"   style="border:1px solid #000000"><%=rs1("materia")%></td>
	<td class=campo align="center" style="border:1px solid #000000"><%=rs1("aulas")%></td>
	<td class=campo align="center" style="border:1px solid #000000" nowrap><%=rs1("periodo")%></td>
</tr>
<%
rs1.movenext:loop
%>
</table>

</td></tr>
</table></div> <!-- tabela borda -->
<%

end if 'rsd.recordcount
rs1.close
%>
<!-- inicio quadro das nomeações -->
	
	<p align="justify"><font size="3">Recebam nossas considerações.</font></p>
	<p align="justify">&nbsp;</p>
	<p><font size="3">Atenciosamente</font></p>
	<p align="justify">&nbsp;</p>

<!-- tabela data e assinatura -->
	<table border="0" cellpadding="0" width="100%" cellspacing="0">
	<tr>
<%if day(now())=1 then dia="1º" else dia=day(now())%>
		<td width="50%" valign="top">
		<p><font size="3">Osasco,&nbsp;<%=dia & " de " & monthname(month(now())) & " de " & year(now()) %></font></p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p><font size="3">_____________________________________<br>
		</font></p>
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
rs.close
'****************************
%>

<%
'****************************
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
	<td class=campo>&nbsp;<a href="declaracao.asp?codigo=<%=rs("chapa")%>&parag=<%=session("40parag")%>"><%=rs("nome")%></a></td>
	<td class=campo>&nbsp;<%=rs("codsituacao")%></td>
</tr>
<%
rs.movenext
loop
%>
</table>
<%
rs.close
end if ' temps

set rsd=nothing
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>