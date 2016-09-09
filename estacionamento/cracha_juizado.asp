<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a87")="N" or session("a87")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
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

if request.form<>"" then
	if request.form("B3")<>"" then
		finaliza=1
	else
		finaliza=0
	end if
end if
if request.form("cracha")="" then tc="V" else tc=request.form("cracha")
if tc="H" then w1=565:h1=350
if tc="V" then w1=330:h1=400

if finaliza=0 then
sql="select cracha from veiculos_juizado"
rs.Open sql, ,adOpenStatic, adLockReadOnly
ultimo=rs("cracha")
rs.close
if month(now)<=6 then
	validade="JUN/" & year(now)
elseif month(now)<=11 then
	validade="DEZ/" & year(now)
else
	validade="JUN/" & year(now)+1
end if
%>
<p class=titulo>Seleção para impressão geral de crachás do Juizado Especial</p>
<form method="POST" action="cracha_juizado.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=250>
<tr><td class=titulo>Campus/Seção</td>
</tr>
<tr><td class=titulo><select size="1" name="D1">
	<option value="Narciso">Para Campus Narciso</option>
	<option value="Yara">Para Campus V.Yara</option>
	</select></td>
	</tr>
<tr><td class=titulo>A partir do nº <input type="text" name="numero" value="<%=ultimo+1%>" size=3>
	Quantidade <input type="text" name="quantidade" value="10" size=3>
	</td></tr>
<tr><td class=titulo>Validade <input type="text" name="validade" value="<%=validade%>" size=10></td></tr>
<tr>
	<td class=titulo colspan=2>
	<input type="radio" name="cracha" value="V" onclick="javascript:form.submit();" <%if tc="V" then response.write "checked"%> > Vertical
	<input type="radio" name="cracha" value="H" onclick="javascript:form.submit();" <%if tc="H" then response.write "checked"%> > Horizontal
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=250>
<tr><td align="center" class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3"></td></tr>
</table>
</form>
<hr>
<%
end if 'finaliza=0

if finaliza=1 then
	tipo=request.form("cracha"):if tipo="" then tipo="H"

	idc=request.form("D1")
	if idc="Narciso" then
		texto1="<b>Narciso</b>":imagem="1sem_n.GIF":imagemv="1sem_n_v.gif"
	elseif idc="Yara" then
		texto1="<b>V. Yara</b>":imagem="1sem_y.GIF":imagemv="1sem_y_v.gif"
	else
		texto1=""
	end if
quantidade=cint(request.form("quantidade"))
numeroinicial=cint(request.form("numero"))

for a=0 to quantidade-1

if tipo="H" then
	t1w=537:t1h=318
	t2w=487:t2h=318
%>
<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' height=320 width=1028>
<!-- quadro com os dados -->
<tr><td width=537 height=318 valign=top align="left" style="background-color:transparent;border:1px dotted #000000;background:transparent url('../images/<%=imagem%>') no-repeat center;">

	<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' height=318 width=537>
	<tr><td height=50 width=225 valign=top align="left" style="background-color:transparent">
			<img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
		<td width=<%=537-225%> valign=top style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:30pt;text-align:right"><b><%=numeroinicial+a%>&nbsp;</td>
	</tr>
	<tr><td height=100% valign=middle align="center" colspan=2 style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:36pt;text-align:center"><b><%="JUIZADO ESPECIAL"%></td>
	</tr>
	<tr><td height=25 valign=top style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:24pt;text-align:left"><i>Campus</td>
		<td style="background-color:transparent"></td>
	</tr>
	<tr><td height=40 valign=top style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:40pt;text-align:left"><i><%=texto1%></td>
		<td style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:18pt;text-align:right"><i>Validade: <%=request.form("validade")%>&nbsp;</td>
	</tr>
	</table>	
	
<!-- quadro com o texto -->
</td><td width=487 height=318 valign=top align="left" style="background-color:transparent;border:2px dotted #000000;">
	
	<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' height=318 width=487>
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
	Para uso exclusivo dos funcionários e estagiários do Juizado Especial Civel - Comarca de Osasco
	</td></tr>
	</table>

</td></tr>
</table>
<%

else 'vertical

%>
<table border='0' cellpadding='2' cellspacing='0' width=297 height=399 style="border-collapse: collapse;background: transparent url('../images/<%=imagemv%>') no-repeat center;">
<tr><td style="background-color: transparent;" valign=top align="center" height=75 style="border-top: 1 dotted #000000;border-left: 1 dotted #000000;border-right: 1 dotted #000000">
<img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt="">
<b><%=codigo%>
</td></tr>

<tr><td style="background-color: transparent;" valign=top align="center" height=25>
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:36pt;text-align:center">
<b><%=numeroinicial+a%></b></p>
</td></tr>
<%
%>
<tr><td style="background-color: transparent;" valign="center" align="center" height=100%>
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:24pt;text-align:center">
<b>
<%
response.write "JUIZADO ESPECIAL"
%>
</p>
</td></tr>

<tr><td style="background-color: transparent;" valign=top align="center" height=80>
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:24pt;text-align:left">
<i>&nbsp;Campus
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:40pt;text-align:left">
&nbsp;<%=texto1%></p>
</td></tr>

<tr><td style="background-color: transparent;" valign=top height=28 style="border-bottom: 1 dotted #000000;border-left: 1 dotted #000000;border-right: 1 dotted #000000">
<p style="margin-top:0;margin-bottom:0;color:Black;font-size:18pt;text-align:center">
<i>Validade: <%=request.form("validade")%></i></p>
</td></tr>

</table>

<table border='0' cellpadding='1' cellspacing='0' style='border-collapse: collapse' width=297 height=399>
<tr><td valign="center" align="center" style="border:1 dotted #000000;border-bottom:1 dotted #000000"><img src="../images/fundo_cracha2.gif" border="0">
</td></tr>
</table>

<%
end if 'tipo

if a<quantidade-1 then response.write "<DIV style=""page-break-after:always""></DIV>"
impresso=numeroinicial+a
next
sql="update veiculos_juizado set cracha=" & impresso
conexao.Execute sql, , adCmdText

end if 'finaliza=1
%>
</body>
</html>
<%

set rs=nothing
conexao.close
set conexao=nothing
%>