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
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
	idchapa=request.form("D1")
	if idchapa="Todos" then
		sqld=" and a.empresa in ('I','U') "
	elseif idchapa="I" then
		sqld=" and a.empresa in ('I') "
	elseif idchapa="U" then
		sqld=" and a.empresa in ('U') "
	elseif idchapa="E" then
		sqld=" and a.chapa in (select d.chapa from assmed_dep d, assmed_dep_mudanca m where d.parentesco='Esposo' and m.id_dep=d.id_dep and m.empresa='U' and getdate() between m.ivigencia and m.fvigencia) "
	else
		sqld=" and a.chapa='" & idchapa & "' and a.empresa in ('I','U') "
	end if

	sqlc="SELECT a.chapa, a.empresa, a.plano, a.codigo, a.inclusao, f.CODSITUACAO, f.NOME, f.CODSECAO, s.DESCRICAO " & _
	"FROM assmed_mudanca a, corporerm.dbo.pfunc f, corporerm.dbo.PSECAO s WHERE a.chapa=f.chapa collate database_default and f.codsecao=s.codigo " & _
	"AND getdate() Between [ivigencia] And [fvigencia] AND f.CODSITUACAO<>'D' "
	sqle="order by f.codsecao, a.chapa "
	sqlb=sqlc & sqld & sqle
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
	temp=0
else
	temp=1
end if

if temp=1 then
	sqla="SELECT chapa, nome from corporerm.dbo.pfunc f where f.codsituacao<>'D' and f.codsecao<>'03.1.999' order by nome"
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>
Impressão de Opção por Co-Participação n assistência médica
<form method="POST" action="coparticipacao.asp">
<p style="margin-top: 0; margin-bottom: 0">
<select size="1" name="D1">
	<option value="Todos">Todos</option>
	<option value="I">Intermédica</option>
	<option value="U">Unimed Seguros</option>
	<option value="E">Esposos</option>
<%
rs.movefirst:do while not rs.eof
%>
	<option value="<%=rs("chapa")%>"><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<input type="submit" value="Imprimir" name="B1" class="button"></p>
</form>
<%
elseif temp=0 then
meiapagina=1
rs.movefirst
do while not rs.eof

select case rs("empresa")
	case "I"
		operadora="Intermédica Sistema de Saúde"
		planogratis="EXTRA"
		clausula="cláusula 49 item 5"
		copar=cdbl(4.24)
	case "U"
		operadora="Unimed Seguros"
		planogratis="BÁSICO"
		clausula="cláusula 40 item 5"
		copar=cdbl(10.49)
end select
inicial=0:total1=0:total2=0
titulo=rs("chapa") & " - " & rs("nome")

if meiapagina=0 then meiapagina=1 else meiapagina=0
%>
<!-- quadro dividor -->
<table border="0" bordercolor="#000000" cellpadding="0" width="691" cellspacing="0" style="border-collapse: collapse">
<tr><td valign=top>

<!-- inicio do recibo -->
<table border="0" bordercolor="#000000" cellpadding="4" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td colspan=2 class=titulop style="border: 1px solid #000000"><b>OPÇÕES AO PLANO DE ASSISTÊNCIA MÉDICO-HOSPITALAR</b></td>
</tr>
<tr>
	<td class=campo><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<td class=campo style="border-bottom: 1px solid #000000"><b>1.</b> Desejo contribuir na modalidade 
	de co-participação, conforme artigo 30 da Lei nº 9656/98 e <%=clausula%> da Convenção Coletiva 
	de Trabalho, que permite continuar a usufruir do plano de saúde após rescisão do contrato de trabalho sem justa causa, por um 
	período mínimo de 6 meses e máximo de 24 meses, conforme artigo 30 § 1º da referida lei.</td>
</tr>
<tr>
	<td class=campo><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<td class=campo style="border-bottom: 1px solid #000000"><b>2.</b> Não desejo contribuir para o plano de 
	saúde na modalidade de co-participação, permanecendo as condições anteriores.</td>
</tr>
<tr>
	<td colspan=2 class=campo>	
	&nbsp;<br>_____________________________________<br>
	<%=rs("chapa") & " - " & rs("nome")%>
	</td>
</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="4" width="690" cellspacing="0" style="border-collapse: collapse">
<tr><td width=30></td><td width=250></td><td width=80></td><td width=330></td></tr>
<tr>
	<td colspan=2 class=fundo style="border-top: 1px solid #000000;border-bottom: 1px solid;border-left: 1px solid"><b>RENOVAÇÃO DA AUTORIZAÇÃO PARA DESCONTO</b></td>
	<td colspan=2 class=campo valign=top align="center" style="">
<%
sqlplano="SELECT codigo, seq, plano, valor, reembolso FROM assmed_planos " & _
"WHERE codigo='" & rs("empresa") & "' AND plano='" & rs("plano") & "' "
rs2.Open sqlplano, ,adOpenStatic, adLockReadOnly
%>	
	<table border="1" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td align="center" class=campo>&nbsp; Planos &nbsp;</font></td>
		<td align="center" class=campo>&nbsp; Custo &nbsp;</font></td>
		<td align="center" class=campo>&nbsp; Desconto Titular (opção 1) &nbsp;</font></td>
		<td align="center" class=campo>&nbsp; Desconto Titular (opção 2) &nbsp;</font></td>
	</tr>
	<tr>
		<td class=campo>&nbsp;<%=rs2("plano")%>&nbsp;</font></td>
		<td class=campo align="center"><%=formatnumber(rs2("valor"),2)%></td>
		<td class=campo align="center"><%=formatnumber(rs2("reembolso")+copar,2)%></td>
		<td class=campo align="center"><%=formatnumber(rs2("reembolso"),2)%></td>
	</tr>
	</table>
<%
total1=total1+rs2("reembolso")+copar
total2=total2+rs2("reembolso")
rs2.close
%>
	</td>
</tr>
<tr>
	<td colspan=4 class=campo valign=top align="center">
	</td>
</tr>
<tr>
	<td colspan=4 class="campop"><b></td>
</tr>
<tr>
	<td colspan=4 class="campop" valign=top align="center">
<%
sql2="SELECT d.chapa, d.dependente, d.sexo, d.nascimento, d.parentesco, d.mae, m.empresa, m.plano, p.valor, p.reembolso, d.cpf " & _
"FROM assmed_dep d, assmed_dep_mudanca m, assmed_planos p " & _
"WHERE d.chapa=m.chapa and d.nrodepend=m.nrodepend and m.plano=p.plano AND m.empresa=p.codigo " & _
"AND d.chapa='" & rs("chapa") & "' AND getdate() Between [ivigencia] And [fvigencia] "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
	<table border="0" bordercolor="#000000" cellpadding="1" width="650" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td colspan=4 class=campo align="center" style="border: 1px solid">Dependentes inscritos na vigência do contrato de trabalho:</td>
		<td align="center" class=campo style="border: 1px solid">Desconto<br>Dep.(opção 1)</td>
		<td align="center" class=campo style="border: 1px solid">Desconto<br>Dep.(opção 2)</td>
	</tr>
<%
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
idade=int((now()-rs2("nascimento"))/365.25)
if rs("empresa")="M" then
	desconto=rs2("reembolso")
else
	desconto=rs2("valor")
end if
if rs2("parentesco")="Esposo" then desconto=rs2("valor")
if rs("plano")="Diamante I" or rs("plano")="Diamante III" then desconto=rs2("reembolso")
desconto1=desconto+copar:if desconto1>rs2("valor") then desconto1=desconto
desconto2=desconto
totaldep=rs2.recordcount
if totaldep>=5 then meiapagina=1
%>
	<tr>
		<td class=fundo rowspan=2 align="center" style="border-bottom: 1px solid"><%=rs2.absoluteposition%></td>
		<td class=campo style="border-left: 1px solid"><u><%=rs2("dependente")%></td>
		<td class=campo>Data Nasc: <u><%=rs2("nascimento")%></td>
		<td class=campo style="border-right: 1px solid">Parent.:<u><%=rs2("parentesco")%></td>
		<td class=campo rowspan=2 style="border: 1px solid" align="center"><%=formatnumber(desconto1,2)%></td>
		<td class=campo rowspan=2 style="border: 1px solid" align="center"><%=formatnumber(desconto2,2)%></td>
	</tr>
	<tr>
		<td class=campo colspan=2 style="border-bottom: 1px solid;border-left: 1px solid">Nome da mãe do dependente: <%=rs2("mae")%></td>
		<td class=campo colspan=1 style="border-bottom: 1px solid;border-right: 1px solid">CPF: <%=rs2("cpf")%></td>
	</tr>
<%
total1=total1+desconto1
total2=total2+desconto2
rs2.movenext
loop
else
%>
	<tr>
		<td class=campo colspan=6 style="border: 1px solid" align="center"><b>Nenhum dependente inscrito</td>
	</tr>
<%
end if 'rs2.recordcount
rs2.close
%>
	</table>
	</td>
</tr>
<tr>
	<td colspan=4 class=campo>Renovo a autorização do desconto mensal em meu salário, através da folha de pagamento, da diferença 
	de valores entre o plano de saúde "<%=planogratis%>" a que tenho direito atualmente como <%=tipo%> e o plano acima por mim 
	escolhido:</td>
</tr>
<tr>
	<td class=campo valign="center" width=30><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<td class=campo colspan=3 valign="center" width=660>Co-participação: <%=formatnumber(total1,2)%> (<%=extenso2(total1)%>)</td>
</tr>
<tr>
	<td class=campo valign=middle width=30><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
	<td class=campo colspan=3 valign="center">Condições atuais: <%=formatnumber(total2,2)%> (<%=extenso2(total2)%>)</td>
</tr>
<tr>
	<td colspan=3 class=campo width=600>Osasco,&nbsp;<%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %><br>
	Autorizo o desconto.<br><br>
	_____________________________________<br>
	<%=rs("chapa")%>  - <%=rs("nome") %>
	</td>
	<td class="campor" align="right" valign=bottom><%=rs("descricao")%> (<%=rs("codsituacao")%>)</td>
</tr>
</table>
<!-- final do recibo -->

<!-- quadro dividor -->
</td></tr>
</table>
<!-- quadro dividor -->

<%
if meiapagina=0 then response.write "<hr>"
if meiapagina=1 and rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>"
'response.write "<DIV style=""page-break-after:always""></DIV>"
rs.movenext
loop
rs.close

end if ' temps
%>
</body>
</html>
<%
set rs2=nothing
conexao.close
set conexao=nothing
%>