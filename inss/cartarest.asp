<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a39")="N" or session("a39")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Carta de Restituição de INSS</title>
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
	if request.form("meses")="" then session("mesrest")=session("mesrest") else session("mesrest")=request.form("meses")
	if request.form("anoi")="" then session("anoirest")=session("anoirest") else session("anoirest")=request.form("anoi")
	if request.form("anof")="" then session("anofrest")=session("anofrest") else session("anofrest")=request.form("anof")
	sqla="SELECT f.CHAPA, f.NOME, c.NOME AS FUNCAO, P.CARTEIRATRAB, P.SERIECARTTRAB, P.CPF, " & _
	"F.PISPASEP, F.DATAADMISSAO, F.DATADEMISSAO, P.SEXO, f.codsituacao " & _
	"FROM corporerm.dbo.PFUNC f, corporerm.dbo.PPESSOA P, corporerm.dbo.PFUNCAO c " & _
	"WHERE f.codpessoa=p.codigo and f.codfuncao=c.codigo " 
	
	sql1=sqla & sqlb
	set rs=server.createobject ("ADODB.Recordset")
	Set rs.ActiveConnection = conexao
	set rsi=server.createobject ("ADODB.Recordset")
	Set rsi.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	session("chapa")=rs("chapa")
	chapa=rs("chapa")
	nome=rs("nome")
	temp=0
	if rs.recordcount>0 and session("cartateto")<>"L" then temp=2
else
	temp=1
end if
%>

<%
if temp=1 then
session("cartateto")="F"
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>
Seleção do funcionário para emissão de carta de restituição
<form method="POST" action="cartarest.asp" name="form">
  <p style="margin-top: 0; margin-bottom: 0">Chapa/Nome <input type="text" name="chapa" size="20" class="form_box" value="<%=session("chapa")%>"></p>
  <p style="margin-top: 0; margin-bottom: 0">Quantidade de meses para listar <input type="text" name="meses" size="3" value="60"></p>
  <p style="margin-top: 0; margin-bottom: 0">Entre os anos de <input type="text" name="anoi" size="6" value="<%=year(now)-4%>">
  e  <input type="text" name="anof" size="6" value="<%=year(now)%>">
  </p>
  
  <p style="margin-top: 0; margin-bottom: 0">Não imprimir zerados 	<input type="checkbox" name="zerado" value="ON" checked></p>
  <p style="margin-top: 0; margin-bottom: 0">
  <input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>

<%
elseif temp=0 then
session("cartateto")="C"
'if request.form<>"" then
if rs("sexo")="F" then v1="a" else v1="o"
if rs("sexo")="F" then v2="a" else v2=""
if rs("sexo")="F" then v3="à" else v3="ao"
if rs("codsituacao")="D" then
	texto1=" foi ":texto2=" de ":texto3=" a " & rs("datademissao")
else
	texto1=" é ":texto2=" desde ":texto3=""
end if
%>
<div align="center">
<table border="0" cellpadding="5" width="620" cellspacing="0" height="1000">
<tr>
	<td width="100%"><img border="0" src="../images/aguia.jpg"></td></tr>
<tr>
	<td width="100%">&nbsp;</td></tr>
<tr>
	<td width="100%" align="center"><b><font size="4">DECLARAÇÃO</font></b></td></tr>
<tr>
	<td width="100%">
	<p>&nbsp;</p>
	<p align="justify">Declaramos aos orgãos interessados, que <%=v1%> Sr<%=v2%>. <%=rs("nome")%>,
	portador<%=v2%> da CTPS nº <%=rs("carteiratrab")%> / <%=rs("seriecarttrab")%>, 
	do C.P.F. nº <%=rs("cpf")%> e do PIS/PASEP nº <%=rs("pispasep")%>,
	<%=texto1%>	funcionári<%=v1%> desta Instituição de Ensino Superior <%=texto2%>
	<%=rs("dataadmissao")%><%=texto3%>, exercendo a função de <%=rs("funcao")%>.</p>
	<p align="justify">Esclarecemos ainda que descontamos, recolhemos e não
	devolvemos as contribuições abaixo mencionadas para <%=v1%> referid<%=v1%> funcionári<%=v1%>,
	e não compensamos a importância em GRPS nem pleiteamos a restituição
	junto ao INSS.</p>
<%
sqlv="select top " & session("mesrest") & " anocomp, mescomp, nroperiodo, baseinss, inss, inssferias, insscalcusuario, baseinss13, inss13 " & _
"from pfperff " & _
"where chapa='" & rs("chapa") & "' and (baseinss<>0 or baseinss13<>0) and anocomp between " & session("anoirest") & " and " & session("anofrest") & "" & _
"ORDER BY ANOCOMP desc, MESCOMP desc " 

sqla="select anocomp, mescomp, nroperiodo, baseinss, inss, inssferias, insscalcusuario, baseinss13, inss13 " & _
"from corporerm.dbo.pfperff " & _
"where chapa='" & rs("chapa") & "' and (baseinss<>0 or baseinss13<>0) and anocomp between " & session("anoirest") & " and " & session("anofrest") & "" 
'"ORDER BY ANOCOMP desc, MESCOMP desc " 
'response.write "<br>" & sqla

sqlm="select anocomp, mescomp, nroperiodo, baseinss, inss, inssferias, insscalcusuario, baseinss13, inss13 " & _
"from corporerm.dbo.pfperffcompl " & _
"where chapa='" & rs("chapa") & "' and (baseinss<>0 or baseinss13<>0) and anocomp between " & session("anoirest") & " and " & session("anofrest") & "" 
'"ORDER BY ANOCOMP desc, MESCOMP desc " 
'response.write "<br>" & sqlm

if request("zeros")="y" then restricao="WHERE INSS>0 or INSSFERIAS>0 or inss13>0 or insscalcusuario>0 " else restricao=""
sqlv="select top " & session("mesrest") & " * from (" & sqla & " union all " & sqlm & ") t " & restricao & " ORDER BY ANOCOMP desc, MESCOMP desc "
'response.write "<br>" & sqlv
rsi.Open sqlv, ,adOpenStatic, adLockReadOnly
quant=rsi.recordcount
resto=quant mod 15
if resto=0 then resto=0 else resto=1
colunas=int(quant/15) + resto
'if colunas>3 then colunas=3
espacamento=3
if quant>15 then espacamento=2
if quant>30 then espacamento=1
if quant>45 then espacamento=0
%>
	<div align="center">
	<center>
	<table border="0" cellspacing="0">
	<tr>
<%
rsi.movefirst
'do while not rsi.eof
%>
<% for a=1 to colunas %>
	<td class=campo colspan=3 valign=top>
	<table border="1" cellpadding="<%=espacamento%>" cellspacing="1" style="border-collapse: collapse">
	<tr><td class=titulo align="center">Compe-<br>tência</td>
		<td class=titulo align="center">Salário<br>Contrib.</td>
		<td class=titulo align="center">Valor<br>INSS</td></tr>

<!-- tabela -->			
<% 
if a=colunas then final=quant else final=a*15
for b=a*15-14 to final
rsi.absoluteposition=b
if rsi("inssferias")="" or isnull(rsi("inssferias")) then inssferias=0 else inssferias=rsi("inssferias")
if rsi("inss")="" or isnull(rsi("inss")) then inss=0 else inss=rsi("inss")
if rsi("insscalcusuario")="" or isnull(rsi("insscalcusuario")) then inssc=0 else inssc=rsi("insscalcusuario")
if rsi("inss13")="" or isnull(rsi("inss13")) or cdbl(rsi("baseinss13"))=0 then inss13=0 else inss13=rsi("inss13")
total=cdbl(inss)+cdbl(inssferias)+cdbl(inssc)+cdbl(inss13)
base=cdbl(rsi("baseinss"))+cdbl(rsi("baseinss13"))
if cdbl(rsi("baseinss13"))>0 then b13=" 13º" else b13=""
%>
	<tr>
		<td class=campo align="center">&nbsp;<%=numzero(rsi("mescomp"),2) & "/" & rsi("anocomp") %> <%=b13%></td>
		<td class=campo align="center">&nbsp;<%=formatnumber(base,2) %></td>
		<td class=campo align="center">&nbsp;<%=formatnumber(total,2) %></td></tr>
<%next %>			
	</table>
	</td>
<%next %>
<%
'rsi.movenext
'loop
%>
<!-- tabela -->			
	</tr>
<%
rsi.close
%>
	</table>
	</center>
	</div>
	</td></tr>

<tr><td width="100%">
	<table border="0" cellpadding="0" width="100%" cellspacing="0">
<%if day(now())=1 then dia="1º" else dia=day(now())%>
	<tr><td width="50%" valign="top">
		<p><font size="2">Osasco,&nbsp;<%=dia & " de " & monthname(month(now())) & " de " & year(now()) %></font></p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p><font size="2">_____________________________________<br>
		</font><input type="text" name="nome" size="70" maxlength="256" class=form_input></p>
		</td>

<%if teste=1 then %>
		<td width="50%" valign="top">&nbsp;
		<div align="center">
		<center>
		<table border="0" cellpadding="0" width="240" cellspacing="0">
		<tr>
			<td width="1" style="border-left: 3px solid; border-top-style: 3px solid;"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240" rowspan="2">
				<p align="center"><b><font size="4" color="#808080">73.063.166/0001-20</font></b></td>
			<td width="1" style="border-right: 3px solid; border-top: 3px solid;"><img border="0" src="../images/branco.gif" width=10 height=10></td>
		</tr>
		<tr>
			<td width="1"></td>
			<td width="1"></td>
		</tr>
		<tr>
			<td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td>
			<td width="1"></td>
		</tr>
		<tr>
			<td width="1"></td>
			<td width="240">
			<p align="center"><b><font color="#808080">FUNDAÇÃO INSTITUTO DE<br>ENSINO PARA OSASCO</font></b></td>
			<td width="1"></td>
		</tr>
		<tr>
			<td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td>
			<td width="1"></td>
		</tr>
		<tr>
			<td width="1">&nbsp;</td>
			<td width="240" rowspan="2">
			<p align="center"><font color="#808080">Rua Narciso Sturlini, 883<br>
			Jd. Umuarama - CEP 06018-903<br>
			OSASCO - SP</font></td>
			<td width="1"></td>
		</tr>
		<tr>
			<td width="1" style="border-left: 3px solid; border-bottom: 3px solid;"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="1" style="border-right: 3px solid; border-bottom: 3px solid;"><img border="0" src="../images/branco.gif" width=10 height=10></td>
		</tr>
		</table>
		</center>
		</div>
		<p>&nbsp;
<%end if%>
	</td>
	</tr>
	</table>
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td></tr>
<tr>
	<td height="15"><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b></td>
</tr>
<tr>
	<td height="15">
		<font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP
		06018-903 - Fone: (011) 3681-6000<%if teste=0 then response.write " - C.N.P.J. 73.063.166/0001-20" %></font></td>
</tr>
<tr>
	<td height="15">
		<font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP
		06020-190 - Fone: (011) 3651-9999<%if teste=0 then response.write " - C.N.P.J. 73.063.166/0003-92" %></font></td>
</tr>
<%if teste=0 then%>
<tr>
	<td height="15">
		<font face="Arial Narrow">Av. Franz Voegeli, 1743 - Osasco - SP - CEP
		06020-190 - Fone: (011) 3654-0655 - C.N.P.J. 73.063.166/0004-73</font></td>
</tr>
<%end if%>
<tr>
	<td height="15">
	<font face="Arial Narrow">Caixa Postal - ACF - Jd. Ipê - nº 2530 -
	Osasco - SP - CEP 06053-990</font>
	</td></tr>
</table>
</div>
<%
rs.close
set rs=nothing
elseif temp=2 then
session("cartateto")="L"
%>
<!-- mostrar funcionarios e as contribuições -->
<table border="1" cellpadding="0" width="550" cellspacing="0">
<tr>
	<td class=titulo>&nbsp;Chapa</td>
	<td class=titulo>&nbsp;Nome</td>
	<td class=titulo>&nbsp;Situacao</td>
</tr>
<%
rs.movefirst
do while not rs.eof
if request.form("zerado")="ON" then zeros="y" else zeros="n"
%>
<tr>
	<td class=campo>&nbsp;<%=rs("chapa")%></td>
	<td class=campo>&nbsp;<a href="cartarest.asp?codigo=<%=rs("chapa")%>&zeros=<%=zeros%>"><%=rs("nome")%></a></td>
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
%>
</body>
</html>
<%
conexao.close
set conexao=nothing
%>