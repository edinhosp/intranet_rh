<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Crachá de Estacionamento</title>
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
	chapa=request("chapa")
	ano=request("ano")
	sql1="SELECT * from veiculos_alunosfunc where chapa='" & chapa & "' and validade='" & ano & "' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	codigo=right(year(now),1)
	soma=0
	for a=2 to 3
		numero=mid(chapa,a,1)
		soma=soma+numero
	next
	codigo=codigo & numzero(soma,2):soma=0
	for a=4 to 5
		numero=mid(chapa,a,1)
		soma=soma+numero
	next
	codigo=codigo & numzero(soma,2):soma=0
soma=0:pos=0
for a=1 to len(rs("placa"))
	numero=mid(rs("placa"),a,1)
	if isnumeric(numero)=true then
		soma=soma+numero
		pos=pos+1
	end if
	if pos=2 then
		pos=0		
		codigo=codigo & numzero(soma,2):soma=0
	end if
next
codigo=codigo & "1"

sql2="select chapa, nome, descricao from (select chapa, nome, codsecao as descricao, codsindicato from grades_novos union all select f.chapa collate database_default, f.nome collate database_default, f.secao collate database_default, f.codsindicato collate database_default from qry_funcionarios f) f " & _
"where chapa='" & rs("chapa") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
nome=rs2("nome"):secao=rs2("descricao")
rs2.close

texto1="&nbsp;"
if rs("campus_estudo")="VY" then texto1="<b>V. Yara</b>":str1="y"
if rs("campus_estudo")="NS" then texto1="<b>Narciso</b>":str1="n"
if rs("periodo")="M" then str2="7":texto2="Matutino"
if rs("periodo")="N" then str2="19":texto2="Noturno"
str2=""
imagem="_" & str1 & str2 & "7.bmp"

'width anterior=1028
%>
<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' height=320 width=539>
<!-- quadro com os dados -->
<tr><td width=537 height=318 valign=top align="left" style="background-color:transparent;border:1px dotted #000000;background:transparent url('../images/<%=imagem%>') no-repeat center;">

	<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' height=318 width=537>
	<tr><td height=50 width=225 valign=top align="left" style="background-color:transparent">
			<img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
		<td width=<%=537-225%> valign=top style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:30pt;text-align:right"><b><%=rs("chapa")%>&nbsp;</td>
	</tr>
	<tr><td height=100% valign=middle align="center" colspan=2 style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:38pt;text-align:center"><b>
<%
do while not rs.eof
response.write rs("placa")
if rs.recordcount>1 then
	if rs.absoluteposition<rs.recordcount then response.write "<br>"
end if
rs.movenext
loop
rs.movefirst
%>
		</td>
	</tr>
	<tr><td height=25 valign=top style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:24pt;text-align:left"><i>Campus</td>
		<td style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:30pt;text-align:right"><b><%=rs("matricula")%>&nbsp;</td>
	</tr>
	<tr><td height=40 valign=top style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:40pt;text-align:left"><i><%=texto1%></td>
		<td style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:20pt;text-align:right"><i>Validade: <%=rs("validade")%></i>&nbsp;
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:10pt;text-align:right"><b># <%=codigo%>&nbsp;</td>
	</tr>
	</table>	
	
<!-- quadro com o texto -->
</td>
<!--
<td width=487 height=318 valign=top align="left" style="background-color:transparent;border:1px dotted #000000;">

	<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' height=318 width=487>
	<tr><td valign=top align="left" class="campop">
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:16pt;text-align:justify">
	<b>Observações:</b>
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:10pt;text-align:justify">
	1. Esta plaqueta dá direito ao funcionário ingressar seu veículo no estacionamento do UNIFIEO - Centro Universitário FIEO.
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:10pt;text-align:justify">
	2. Durante a permanência do veículo no pátio do estacionamento, a plaqueta deverá estar junto ao parabrisa em lugar visível.
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:10pt;text-align:justify">
	3. O funcionário que ceder, emprestar ou fizer mau uso da presente plaqueta, terá sua vaga sumariamente cancelada.
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:10pt;text-align:justify">
	4. No caso de perda ou extravio da plaqueta o funcionário deverá comunicar, de imediato, a Tesouraria do UNIFIEO para solicitar a segunda via.
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:10pt;text-align:justify">
	5. Adquirindo a vaga, o funcionário-usuário concorda com seus termos e fica ciente de que o UNIFIEO não se responsabiliza, em nenhuma hipótese, por furto 
	de acessórios e/ou objetos deixados no interior do veículo, nem por danos causados ao mesmo.
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:10pt;text-align:justify">
	6. Os casos omissos serão resolvidos pela Diretoria do UNIFIEO.
	</td></tr>
	<tr><td valign=top align="left" class="campor">
	<%=rs.absoluteposition%>&nbsp;&nbsp;Nome: <b><%=nome%></b> &nbsp;&nbsp;Local: <%=secao%>
	</td></tr>
	</table>

</td>
-->
</tr>
</table>

<hr>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr>
		<td><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
		<td align="right"><p style="font-size:18pt"><b>Termo de Compromisso</b><td>
	</tr>
</table>
<br><br>

<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border-left: 1px solid;border-right: 1px solid;border-top: 1px solid;font-size:10pt">
	<i>Nome do Empregado</i></td></tr>
	<tr><td class="campop" style="border-left: 1px solid;border-right: 1px solid;border-bottom: 1px solid;font-size:12pt">
	<b><%=nome%></b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border-left: 1px solid;border-right: 1px solid;font-size:10pt">
	<i>Chapa</i></td>
	<td class="campop" style="border-right: 1px solid;font-size:10pt">
	<i>Departamento</i></td></tr>
	<tr><td class="campop" style="border-right: 1px solid;border-left: 1px solid;border-bottom: 1px solid;font-size:12pt">
	<%=rs("chapa")%></td>
	<td class="campop" style="border-right: 1px solid;border-bottom: 1px solid;font-size:12pt">
	<%=secao%></td></tr>
</table>
<br>
<%
%>
<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border:1px solid;font-size:12pt;text-align:justify">
Como funcionário da FIEO e aluno do Unifieo, estou recebendo nesta data um crachá para estacionamento durante
o período letivo de aulas, compromentendo-me:<br>
a) que somente eu farei uso deste crachá;<br>
b) a apenas utilizá-lo no campus <%=texto1%>, durante o período <%=texto2%>;<br>
c) a comunicar o RH quando houver alteração do veículo <%=rs("placa")%>;<br>
<br>
Estou ciente de que a não observância destes requisitos, sujeitar-me-á à perda da concessão ao estacionamento.
<br>
	</td></tr>
</table>

<br>

<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border:1px solid;font-size:12pt" width=300>
	<br>Osasco, ______/__________/_______
	<br><br><br><br>_______________________________
	<br>    <%=nome%>
	</td>
</tr>
</table>
<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
<tr><td>
<p style="margin-bottom:0;margin-top:0;text-align:right"><%=rs.absoluteposition%>/<%=rs.recordcount%>
</td></tr></table>




</body>
</html>
<%
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>