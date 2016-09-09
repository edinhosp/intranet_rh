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
	set rs2=server.createobject ("ADODB.Recordset")
	Set rs2.ActiveConnection = conexao
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
%>
<p class=titulo>Seleção para impressão geral de crachás</p>
<form method="POST" action="cracha_geral.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=250>
<tr><td class=titulo>Campus/Seção</td>
</tr>
<tr><td class=titulo><select size="1" name="D1">
	<option <%if request.form("d1")="Narciso"  then response.write "selected"%> value="Narciso">Todos-Campus Narciso</option>
	<option <%if request.form("d1")="Yara"     then response.write "selected"%> value="Yara">Todos-Campus V.Yara</option>
	<option <%if request.form("d1")="Wilson"   then response.write "selected"%> value="Wilson">Todos-Campus J.Wilson</option>
	<option <%if request.form("d1")="ANarciso" then response.write "selected"%> value="ANarciso">Administrativos-Campus Narciso</option>
	<option <%if request.form("d1")="AYara"    then response.write "selected"%> value="AYara">Administrativos-Campus V.Yara</option>
	<option <%if request.form("d1")="PNarciso" then response.write "selected"%> value="PNarciso">Professores-Campus Narciso</option>
	<option <%if request.form("d1")="PYara"    then response.write "selected"%> value="PYara">Professores-Campus V.Yara</option>
	</select></td>
</tr>
<tr>
	<td class=titulo colspan=2>
	<input type="radio" name="cracha" value="V" onclick="javascript:form.submit();" <%if tc="V" then response.write "checked"%> > Vertical
	<input type="radio" name="cracha" value="H" onclick="javascript:form.submit();" <%if tc="H" then response.write "checked"%> > Horizontal
	</td>
</tr>
<tr><td align="center" class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3"></td>
</tr>
</table>
</form>
<hr>
<%
end if 'finaliza=0

if finaliza=1 then
	tipo=request.form("cracha"):if tipo="" then tipo="H"

	idchapa=request.form("D1")
	if idchapa="Narciso" then
		sqld=" and v.ns=1 ":texto1="<b>Narciso</b>":imagem="1sem_n.GIF":imagemv="1sem_n_v.gif"
	elseif idchapa="Yara" then
		sqld=" and v.vy=1 ":texto1="<b>V. Yara</b>":imagem="1sem_y.GIF":imagemv="1sem_y_v.gif"
	elseif idchapa="Wilson" then
		sqld=" and v.jw=1 ":texto1="<b>J.Wilson</b>":imagem="1sem_w.GIF":imagemv="1sem_w_v.gif"
	elseif idchapa="ANarciso" then
		sqld=" and v.ns=1 and f.codsindicato<>'03' ":texto1="<b>Narciso</b>":imagem="1sem_n.GIF":imagemv="1sem_n_v.gif"
	elseif idchapa="AYara" then
		sqld=" and v.vy=1 and f.codsindicato<>'03' ":texto1="<b>V. Yara</b>":imagem="1sem_y.GIF":imagemv="1sem_y_v.gif"
	elseif idchapa="PNarciso" then
		sqld=" and v.ns=1 and f.codsindicato='03' ":texto1="<b>Narciso</b>":imagem="1sem_n.GIF":imagemv="1sem_n_v.gif"
	elseif idchapa="PYara" then
		sqld=" and v.vy=1 and f.codsindicato='03' ":texto1="<b>V. Yara</b>":imagem="1sem_y.GIF":imagemv="1sem_y_v.gif"
	else
		sqld=""
	end if

	dataemissao=dateserial(2013,8,1)
sql1="SELECT v.chapa, f.nome, f.descricao, v.termino FROM veiculos_a v , " & _
"(select chapa, nome, codsecao as descricao, codsindicato, codsituacao from grades_novos union all select f.chapa collate database_default, f.nome collate database_default, f.codsecao collate database_default+'-'+f.secao as secao1, f.codsindicato collate database_default, f.codsituacao collate database_default from qry_funcionarios f ) f " & _
"WHERE f.chapa collate database_default=v.chapa and f.codsituacao in ('A','F','Z') " & sqld & " and v.termino>getdate() and status='A' " & _
"GROUP BY v.chapa, f.nome, f.descricao, v.termino order by f.descricao, f.nome "
'response.write sql1
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	rs.movefirst:	do while not rs.eof
	chapa=rs("chapa")
	codigo=right(year(now),1)
	soma=0
	for a=2 to 3
		numero=mid(chapa,a,1):if isnumeric(numero)=true then soma=soma+numero else soma=soma
	next
	codigo=codigo & numzero(soma,2):soma=0
	for a=4 to 5
		numero=mid(chapa,a,1):soma=soma+numero
	next
	codigo=codigo & numzero(soma,2):soma=0
'response.write chapa
	sql2="select top 3 placa from veiculos where chapa='" & chapa & "' and dttermino is null "
	'response.write sql2
	'response.write ">>>> " & chapa
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	rs2.movefirst

	soma=0:pos=0
	for a=1 to len(rs2("placa"))
		numero=mid(rs2("placa"),a,1)
		if isnumeric(numero)=true then
			soma=soma+numero
			pos=pos+1
		end if
		if pos=2 then
			pos=0		
			codigo=codigo & numzero(soma,2):soma=0
		end if
	next
	codigo=codigo & "2"
	
if tipo="H" then
	t1w=537:t1h=318
	t2w=487:t2h=318
%>
<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' height=320 width=1028>
<!-- quadro com os dados -->
<tr><td width=<%=t1w%> height=<%=t1h%> valign=top align="left" style="background-color:transparent;border:1px dotted #000000;background:transparent url('../images/<%=imagem%>') no-repeat center;">

	<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=<%=t1w%> height=<%=t1h%>>
	<tr><td height=50 width=225 valign=top align="left" style="background-color:transparent">
			<img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
		<td width=<%=537-225%> valign=top style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:30pt;text-align:right"><b><%=rs("chapa")%>&nbsp;</td>
	</tr>
	<tr><td height=100% valign=middle align="center" colspan=2 style="background-color:transparent">
			<p style="margin-top:0;margin-bottom:0;color:Black;font-size:28pt;text-align:center"><b>
<%
do while not rs2.eof
response.write rs2("placa")
if rs2.recordcount>1 then
	if rs2.absoluteposition<rs2.recordcount then response.write "<br>"
end if
rs2.movenext:loop
%>
		</td>
	</tr>
	<tr><td height=25 valign=top style="background-color:transparent"><p style="margin-top:0;margin-bottom:0;color:Black;font-size:24pt;text-align:left"><i>Campus</td>
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
	<%=rs.absoluteposition%>&nbsp;&nbsp;Nome: <b><%=rs("nome")%></b> &nbsp;&nbsp;Local: <%=rs("descricao")%>
	</td></tr>
	</table>

</td></tr>
</table>
<%
else 'vertical
%>

<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=297>
<tr><td valign=top class="campor" colspan=2 style="border-top: 1px dotted #000000;border-left: 1px dotted #000000;border-right: 1px dotted #000000">
&nbsp;&nbsp;Nome: <%=rs("nome")%>
	</td></tr>
<tr><td valign=top class="campor" style="border-left: 1 dotted #000000" nowrap>
&nbsp;&nbsp;Setor: <%=rs("descricao")%>
	</td><td class="campor" align="right" style="border-right: 1px dotted #000000"><%=rs.absoluteposition%></td></tr>
</table><br>

<table border='0' cellpadding='2' cellspacing='0' width=297 height=399 style="border-collapse: collapse;background: transparent url('../images/<%=imagemv%>') no-repeat center;">
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
do while not rs2.eof
response.write rs2("placa")
if rs2.recordcount>1 then
	if rs2.absoluteposition<rs2.recordcount then response.write "<br>"
end if
rs2.movenext:loop
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
<tr><td style="background-color: transparent;" width=297 valign="center" align="center" style="border:1px dotted #000000;border-bottom:1px dotted #000000"><img src="../images/fundo_cracha2.gif" border="0">
</td>
<td valign="center" align="center"><img src="../images/tesoura3.gif" border="0" width="38" height="56" alt=""></td>
</tr>
</table>

<%
end if 'tipo 

rs2.close
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página -->
rs.movenext
loop

rs.close
end if 'finaliza=1
%>
</body>
</html>
<%

set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>