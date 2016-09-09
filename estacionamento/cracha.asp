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
	tipo=request("t"):if tipo="" then tipo="H"
	sql1="SELECT v.chapa, v.dttermino, v.marca, v.modelo, v.ano, v.cor, v.placa, va.vy, va.ns, va.bp, va.jw, " & _
	"va.inicio, va.termino, va.cartao, va.obs, va.pavy, va.pans, va.pabp, va.pajw " & _
	"FROM veiculos AS v INNER JOIN veiculos_a AS va ON v.chapa = va.chapa " & _
	"WHERE v.chapa='" & chapa & "' AND v.dttermino is null and termino>=getdate() "
if request("c")="vy" then sql1=sql1 & " and vy=1 "
if request("c")="ns" then sql1=sql1 & " and ns=1 "
if request("c")="bp" then sql1=sql1 & " and bp=1 "
if request("c")="jw" then sql1=sql1 & " and jw=1 "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	codigo=right(year(now),1)
	soma=0
	for a=2 to 3
		numero=mid(chapa,a,1)
		IF isnumeric(numero) then numero=numero else numero=0
		soma=soma+numero
	next
	codigo=codigo & numzero(soma,2):soma=0
	for a=4 to 5
		numero=mid(chapa,a,1)
		IF isnumeric(numero) then numero=numero else numero=0
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
sql2="select chapa, nome, descricao from (select chapa, nome, codsecao as descricao, codsindicato from grades_novos union all select f.chapa collate database_default, f.nome collate database_default, f.secao collate database_default, f.codsindicato from qry_funcionarios f) f " & _
"where chapa='" & rs("chapa") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
nome=rs2("nome"):secao=rs2("descricao")
rs2.close
texto1="&nbsp;"
if request("c")="vy" then texto1="<b>V. Yara</b>":imagem="1sem_y.GIF":imagemv="1sem_y_v.gif"
if request("c")="ns" then texto1="<b>Narciso</b>":imagem="1sem_n.GIF":imagemv="1sem_n_v.gif"
if request("c")="bp" then texto1="<b>B.Park</b>":imagem="1sem_y.GIF":imagemv="1sem_y_v.gif"
if request("c")="jw" then texto1="<b>J.Wilson</b>":imagem="1sem_w.GIF":imagemv="1sem_w_v.gif"

if tipo="H" then
	t1w=537:t1h=318
	t2w=487:t2h=318
%>
<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' height=<%=t1h+2%> width=<%=t1w+t2w+4%> >
<!-- quadro com os dados -->
<tr><td width=<%=t1w%> height=<%=t1h%> valign=top align="left" style="background-color:transparent;border:1px dotted #000000;background:transparent url(../images/<%=imagem%>)  no-repeat center;">


	<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=<%=t1w%> height=<%=t1h%>>
	<tr><td height=50 width=225 valign=top align="left" style="background-color:transparent">
			<img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
		<td width=<%=537-225%> valign=top style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:30pt;text-align:right"><b><%=rs("chapa")%>&nbsp;</td>
	</tr>
	<tr><td height=100% valign=middle align="center" colspan=2 style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:28pt;text-align:center"><b>
<%
do while not rs.eof
response.write rs("placa")
if rs.recordcount>1 then
	if rs.absoluteposition<rs.recordcount then response.write "<br>"
end if
rs.movenext:loop
rs.movefirst
%>
		</td>
	</tr>
	<tr><td height=25 valign=top style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:24pt;text-align:left"><i>Campus</td>
		<td style="background-color:transparent"></td>
	</tr>
	<tr><td height=40 valign=top style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:40pt;text-align:left"><i><%=texto1%></td>
		<td style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:18pt;text-align:right"><i>Validade: <%=ucase(monthname(month(rs("termino")),2))&"/"&year(rs("termino"))%></i>&nbsp;
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:10pt;text-align:right"><b># <%=codigo%>&nbsp;</td>
	</tr>
	</table>	
	
<!-- quadro com o texto -->
</td><td width=<%=t2w%> height=<%=t2h%> valign=top align="left" style="background-color:transparent;border:1px dotted #000000;">
	
	<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=<%=t2w%> height=<%=t2h%>>
	<tr><td valign=top align="left" class="campop">
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:14pt;text-align:justify">
	<b>Observações:</b>
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:9pt;text-align:justify">
	<!--1. Esta plaqueta dá direito ao funcionário ingressar seu veículo no estacionamento do UNIFIEO - Centro Universitário FIEO. -->
	1. Este cartão, permite ao funcionário/usuário, a título de cortesia gratuita, a comodidade de acesso de seu automóvel devidamente cadastrado ao pátio de estacionamento do campus apropriado, do Centro Universitário UNIFIEO.
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:9pt;text-align:justify">
	<!--2. Durante a permanência do veículo no pátio do estacionamento, a plaqueta deverá estar junto ao parabrisa em lugar visível.-->
	2. O próprio funcionário/usuário deverá estacionar seu veículo em local adequado, colocar o cartão junto ao vidro frontal do automóvel em local visível e levar consigo as chaves.
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:9pt;text-align:justify">
	<!--3. O funcionário que ceder, emprestar ou fizer mau uso da presente plaqueta, terá sua vaga sumariamente cancelada.-->
	3. O funcionário/usuário que ceder, emprestar ou não fizer bom uso do presente cartão e da comodidade que lhe é concedida terá sua vaga sumariamente cancelada.
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:9pt;text-align:justify">
	<!--4. No caso de perda ou extravio da plaqueta o funcionário deverá comunicar, de imediato, a Tesouraria do UNIFIEO para solicitar a segunda via.-->
	4. O motorista que causar eventuais prejuízos aos demais usuários do pátio de estacionamento deverá se responsabilizar pelos danos ocasionados.
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:9pt;text-align:justify">
	<!--5. Adquirindo a vaga, o funcionário-usuário concorda com seus termos e fica ciente de que o UNIFIEO não se responsabiliza, em nenhuma hipótese, por furto 
	de acessórios e/ou objetos deixados no interior do veículo, nem por danos causados ao mesmo.-->
	5. No caso de perda ou extravio do cartão de identificação, o funcionário/usuário deverá, de imediato, comunicar o ocorrido à Inspetoria do UNIFIEO.
	<p style="margin-top:2;margin-bottom:0;color:Black;font-size:9pt;text-align:justify">
	6. O funcionário/usuário tem plena ciência de que o UNIFIEO não se responsabiliza, em nenhuma hipótese, por eventuais danos ou furtos de objetos deixados no interior do veículo, bem como no próprio automóvel, perpetrados por terceiros ou provenientes de caso furtuito ou força maior.
	<br>7. Os casos omissos serão resolvidos pela Diretoria do UNIFIEO.
	</td></tr>
	<tr><td valign=top align="left" class="campor">
	<%=rs.absoluteposition%>&nbsp;&nbsp;Nome: <b><%=nome%></b> &nbsp;&nbsp;Local: <%=secao%>
	</td></tr>
	</table>

</td></tr>
</table>
<%
else 'tipo V
%>

<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=297>
<tr><td valign=top class="campor" colspan=2 style="border-top: 1px dotted #000000;border-left: 1px dotted #000000;border-right: 1px dotted #000000">
&nbsp;&nbsp;Nome: <%=nome%>
	</td></tr>
<tr><td valign=top class="campor" style="border-left: 1px dotted #000000" nowrap>
&nbsp;&nbsp;Setor: <%=secao%>
	</td><td class="campor" align="right" style="border-right: 1px dotted #000000"><%=rs.absoluteposition%></td></tr>
</table><br>

<table border='0' cellpadding='2' cellspacing='0' width=297 height=399 style="border-collapse: collapse;background: transparent url('../images/<%=imagemv%>') no-repeat center;" >
<tr><td style="background-color: transparent;" valign=top align="center" height=75 style="border-top: 1px dotted #000000;border-left: 1px dotted #000000;border-right: 1px dotted #000000">
<img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt="">
<b><%=codigo%>
</td></tr>

<tr><td style="background-color: transparent;" valign=top align="center" height=25>
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:30pt;text-align:center">
<b><%=rs("chapa")%></b></p>
</td></tr>

<tr><td style="background-color: transparent;" valign="center" align="center" height=100%>
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:24pt;text-align:center">
<b>
<%
do while not rs.eof
response.write rs("placa")
if rs.recordcount>1 then
	if rs.absoluteposition<rs.recordcount then response.write "<br>"
end if
rs.movenext:loop
rs.movefirst
%>
</p>
</td></tr>

<tr><td style="background-color: transparent;" valign=top align="center" height=80>
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:24pt;text-align:left">
<i>&nbsp;Campus
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:40pt;text-align:left">
&nbsp;<%=texto1%></p>
</td></tr>

<tr><td style="background-color: transparent;" valign=top height=28 style="border-bottom: 1px dotted #000000;border-left: 1px dotted #000000;border-right: 1px dotted #000000">
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:18pt;text-align:center">
<i>Validade: <%=ucase(monthname(month(rs("termino")),2))&"/"&year(rs("termino"))%></i></p>
</td></tr>
</table>

<table border='0' cellpadding='1' cellspacing='0' style='border-collapse: collapse' width=337 height=399>
<tr><td width=297 valign="center" align="center" style="border:1px dotted #000000;border-bottom:1px dotted #000000"><img src="../images/fundo_cracha2.gif" border="0">
</td>
<td valign="center" align="center"><img src="../images/tesoura3.gif" border="0" width="38" height="56" alt=""></td>
</tr>
</table>

<%
end if 'tipo
%>


</body>
</html>
<%
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>