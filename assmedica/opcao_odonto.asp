<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a85")="N" or session("a85")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Opção Intermédica</title>
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
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request("codigo")<>"" or request.form<>"" then
	if request.form="" then temp=request("codigo")
	if request("codigo")="" then temp=request.form("chapa")
	if isnumeric(temp) then
		info=1
		temp=numzero(temp,5)
		sqlb="AND f.CHAPA='" & temp & "' "
	else
		info=2
		sqlb="AND f.nome like '%" & temp & "%' "
	end if

	sqla="SELECT f.NOME, f.CODSITUACAO, f.CHAPA, f.DATAADMISSAO, f.CODSECAO, f.codsindicato, s.DESCRICAO AS Secao, p.dtnascimento, f.salario " & _
	"FROM corporerm.dbo.PFUNC AS f, corporerm.dbo.PSECAO s, corporerm.dbo.PPESSOA p " & _
	"WHERE f.CODSECAO = s.CODIGO and p.codigo=f.codpessoa "

	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	session("chapa")=rs("chapa")
	temp=0
	if rs.recordcount>1 then temp=2
else
	temp=1
end if

if temp=1 then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário (administrativo ou professor)
<form method="POST" action="opcao.asp">
  <p style="margin-top: 0; margin-bottom: 0">Chapa/Nome <input type="text" name="chapa" size="20" class="form_box" value="<%=session("chapa")%>">
  <input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>

<%
elseif temp=0 then
'if request.form<>"" then
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
idade=int((now()-rs("dtnascimento"))/365.25)
if rs("codsindicato")="03" then empresa="I" else empresa="U"
select case empresa
	case "I"
		dt_inicio="01/10/2003"
		operadora="Intermédica Sistema de Saúde"
		planogratis="EXTRA"
		anterior="SAMCIL"
		valor=formatnumber(40.95,2)
		tipo="PROFESSOR"
		clausula="cláusula 49 item 5"
		copar=cdbl(4.095)
	case "M"
		dt_inicio="19/05/2003"
		operadora="Medial Saúde"
		planogratis="CLASSICO I"
		anterior="AMESP"
		valor=formatnumber(86.10,2)
		tipo="ADMINISTRATIVO"
		clausula="cláusula 40 item 5"
		copar=cdbl(8.61)
	case "U"
		dt_inicio="01/08/2010"
		operadora="Unimed Seguros"
		planogratis="BÁSICO"
		anterior="MEDIAL"
		valor=formatnumber(104.90,2)
		tipo="ADMINISTRATIVO"
		clausula="cláusula 40 item 5"
		copar=cdbl(10.49)
end select
inicial=0
%>
<table border="0" bordercolor="#000000" cellpadding="4" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td colspan=2 class=titulop style="border: 1px solid #000000"><b>OPÇÕES AO PLANO DE ASSISTÊNCIA MÉDICO-HOSPITALAR</b></td>
</tr>
<tr>
	<td class=campo><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<td class="campop" style="border-bottom: 1px solid #000000"><p style="text-align:justify"><b>1.</b> Não desejo me filiar ao plano de saúde
	proposto, por já estar filiado a plano de saúde em outra instituição ou particular, e para tanto estou renunciando conforme 
	documento escrito à parte.</td>
</tr>
<tr>
	<td class=campo><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<td class="campop" style="border-bottom: 1px solid #000000"><p style="text-align:justify"><b>2.</b> Desejo optar pelo plano de saúde abaixo 
	assinalado e contribuir na modalidade de co-participação, conforme artigo 30 da Lei nº 9656/98 e <%=clausula%> da Convenção Coletiva 
	de Trabalho, que permite continuar a usufruir do plano de saúde após rescisão do contrato de trabalho sem justa causa, por um 
	período mínimo de 6 meses e máximo de 24 meses, conforme artigo 30 § 1º da referida lei.</td>
</tr>
<tr>
	<td class=campo><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<td class="campop" style="border-bottom: 1px solid #000000"><p style="text-align:justify"><b>3.</b> Não desejo contribuir para o plano de 
	saúde na modalidade de co-participação, porém desejo optar pelo plano de saúde abaixo assinalado.</td>
</tr>
<tr>
	<td colspan=2 class="campop">	
	&nbsp;<br>_____________________________________<br>
	<%=rs("chapa")%>  - <%=rs("nome") %></p>
	</td>
</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="4" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulop style="border: 1px solid #000000"><b>AUTORIZAÇÃO PARA DESCONTO E/OU INCLUSÃO</b></td>
</tr>
<tr>
	<td class="campop">
	Eu,&nbsp;<%=rs("nome") %> (<%=idade%>), desejo por livre e espontânea vontade, optar por
	um plano de assistência médica diferenciado, identificado abaixo:</td>
</tr>
<tr>
	<td class="campop" valign=top align="center">
<%
sqla="SELECT empresa, plano FROM assmed_mudanca " & _
"WHERE chapa='" & rs("chapa") & "' AND getdate() Between [ivigencia] And [fvigencia] and plano not in ('Não Participa','IP') "
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then plano=rs3("plano") else plano=""
rs3.close

if cdbl(rs("salario"))<10000 then stringplano="'%diamante%'" else stringplano="'diamante%'"
if cdbl(rs("salario"))<3000 then limitep=3 else limitep=4
sqlplano="SELECT codigo, seq, plano, valor, reembolso FROM (select * from assmed_planos where (codigo='I' and seq<=3) or (codigo='U' and seq<=" & limitep & ")) a " & _
"WHERE codigo='" & empresa & "' AND plano Not Like 'agr%' ORDER BY seq "
rs3.Open sqlplano, ,adOpenStatic, adLockReadOnly
%>	
	<table border="1" bordercolor="#000000" cellpadding="1" width="500" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td align="center">Opção</font></td>
		<td align="center">Planos</font></td>
		<td align="center">Custo</font></td>
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular<br>(opção 2)</font></td>
<%if empresa="U" then%>		
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular ou Dependente<br>(opção 3)</font></td>
<%else%>
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular (opção 3)</font></td>
		<td align="center" class=campo>Desconto por&nbsp;<br>Dependente (opção 3)</font></td>
<%end if%>		
	</tr>
<%
rs3.movefirst
do while not rs3.eof
if plano=rs3("plano") then campof="fundo" else campof="campop"
if empresa="I" then desconto3=rs3("valor") else desconto3=rs3("reembolso")
%>
	<tr>
		<td class=<%=campof%> align="center"><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
		<td class=<%=campof%>>&nbsp;<%=rs3("plano")%></font></td>
		<td class=<%=campof%> align="center"><%=formatnumber(rs3("valor"),2)%></td>
		<td class=<%=campof%> align="center"><%=formatnumber(rs3("reembolso")+copar,2)%></td>
<%if empresa="U" then%>
		<td class=<%=campof%> align="center"><%=formatnumber(desconto3,2)%></td>
<%else%>
		<td class=<%=campof%> align="center"><%=formatnumber(rs3("reembolso"),2)%></td>
		<td class=<%=campof%> align="center"><%=formatnumber(desconto3,2)%></td>
<%end if%>
	</tr>
<%
rs3.movenext
loop
%>
	</table>
<%
rs3.close
set rs3=nothing
%>
	</td>
</tr>
<tr>
	<td class="campop">Desejo também&nbsp;incluir os meus dependentes legais (esposa, filhos
	até 21 anos) abaixo relacionados:</td>
</tr>
<tr>
	<td class="campop" valign=top align="center">
<%
sql2="select d.chapa, d.nome, d.dtnascimento, d.grauparentesco, p.descricao, " & _
"datediff(yy,dtnascimento,getdate()) AS idade, c.mae " & _
"from corporerm.dbo.pfdepend d, corporerm.dbo.pcodparent p, corporerm.dbo.pfdependcompl c " & _
"where d.grauparentesco=p.codcliente and c.nrodepend=d.nrodepend and d.chapa=c.chapa " & _
"and d.chapa='" & session("chapa") & "' and d.grauparentesco not in ('6','7') "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
	<table border="1" bordercolor="#CCCCCC" cellpadding="1" width="500" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td align="center"><font size="1">Opção</font></td>
		<td align="center"><font size="1">Nome do Dependente</font></td>
		<td align="center"><font size="1">Grau de&nbsp;<br> Parentesco</font></td>
		<td align="center"><font size="1">Data de&nbsp;<br> Nascimento</font></td>
		<td align="center"><font size="1">Idade</font></td>
	</tr>
<%
totaldep=0
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
idade=int((now()-rs2("dtnascimento"))/365.25)
if rs2("grauparentesco")="1" and idade>20 then
else
%>
	<tr>
		<td align="center" rowspan="2"><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
		<td><font size="2">&nbsp;<%=rs2("nome") %></font></td>
		<td><font size="2">&nbsp;<%=rs2("descricao") %></font></td>
		<td><font size="2">&nbsp;<%=rs2("dtnascimento") %></font></td>
		<td><font size="2">&nbsp;<%=rs2("idade")%></font></td>
	</tr>
	<tr>
		<td class="campor" colspan="4">Nome da mãe do dependente: <font size=2><%=rs2("mae")%>&nbsp;</td>
	</tr>
<%
totaldep=totaldep+1
end if
rs2.movenext
loop
rs2.close
set rs2=nothing
end if

if totaldep<=3 then linhasdep=1 '3
if totaldep<=2 then linhasdep=2 '2
if totaldep<=1 then linhasdep=3 '1
if totaldep>3 then linhasdep=0 '4

for a=0 to linhasdep-1 '3
%>
	<tr>
		<td align="center" rowspan="2"><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td class="campor" colspan="4">Nome da mãe do dependente:<font size=2>&nbsp;</td>
	</tr>
<%next%>
	</table>
	</td>
</tr>
<tr>
	<td class="campop"><p align="justify">Autorizo o desconto mensal em meu salário, através da folha de pagamento, da diferença 
	de valores entre o plano de saúde "<%=planogratis%>" a que tenho direito atualmente como <%=tipo%> e o plano acima por mim 
	escolhido. Estou ciente de que a inclusão do(s) meu(s) dependente(s) será paga integralmente por mim, autorizando desde já, o
desconto em meu salário. Nesta data a aludida diferença entre os planos mencionados é de R$ __________, devendo sofrer reajuste 
quando forem corrigidos os valores cobrados da contratante (FIEO) e que segundo critérios estabelecidos pela <%=operadora%>, 
qualquer alteração no plano só poderei fazer no aniversário do contrato, ou seja, todo mês de <%=monthname(month(dt_inicio))%> de
cada ano.</font></p>
<p><font size="2">Osasco,&nbsp;<%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %><br>
Autorizo o desconto.</font></p>
<p><font size="2">_____________________________________<br>
<%=rs("chapa")%>  - <%=rs("nome") %></font></p>
	</td>
</tr>
</table>
<%
rs.close
set rs=nothing
elseif temp=2 then
%>

<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="550" cellspacing="0" style="border-collapse: collapse">
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
	<td class=campo>&nbsp;<a href="opcao.asp?codigo=<%=rs("chapa")%>"><%=rs("nome")%></a></td>
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