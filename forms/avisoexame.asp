<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a76")="N" or session("a76")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Opções de Comunicação</title>
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
function chapa1() { form.nome.value=form.chapa.value; form.submit(); }
function nome1() { form.chapa.value=form.nome.value; form.submit(); }
--></script>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsi=server.createobject ("ADODB.Recordset")
Set rsi.ActiveConnection = conexao

teste=0
if request.form("B1")="" then
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>
Opções e parametros para emissão do Aviso da Homologação/Exame</p>
<form method="POST" action="avisoexame.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width=500>
<tr>
	<td class=fundo>Chapa <input type="text" name="chapa" size="5" class="form_box" onchange="chapa1()" value="<%=request.form("chapa")%>"></td>
	<td class=fundo><select name="nome" onchange="nome1()">
	<option value="00000">Selecione...</option>
<%
sql="select chapa, nome from corporerm.dbo.pfunc where (codsituacao in ('A','F','Z','L') and codtipo='N') or (codsituacao in ('D') and datademissao>getdate()-15) order by nome"
rsi.Open sql, ,adOpenStatic, adLockReadOnly
rsi.movefirst
do while not rsi.eof
if rsi("chapa")=request.form("chapa") then tmpproc="selected" else tmpproc=""
%>
	<option value="<%=rsi("chapa")%>" <%=tmpproc%>><%=rsi("nome")%></option>
<%
rsi.movenext
loop
rsi.close
%>
	</select></td>
</tr>
</table>

<%
if request.form("chapa")<>"" then
sql1="select dataadmissao, datademissao from corporerm.dbo.pfunc where chapa='" & request.form("chapa") & "' "
rsi.Open sql1, ,adOpenStatic, adLockReadOnly
dataadmissao=rsi("dataadmissao")
if isnull(rsi("datademissao")) then dtdemissao="" else dtdemissao=rsi("datademissao")
rsi.close
end if

if request.form("datasaida")<>"" then datasaida=request.form("datasaida")
if dtdemissao<>"" then datasaida=dtdemissao
if request.form("localpag")<>"" then localpag=request.form("localpag")
if request.form("dthomologacao")<>"" then dthomologacao=request.form("dthomologacao")
if request.form("hrhomologacao")<>"" then hrhomologacao=request.form("hrhomologacao")
if request.form("localexame")<>"" then localexame=request.form("localexame")
if request.form("dtexame")<>"" then dtexame=request.form("dtexame")
if request.form("hrexame")<>"" then hrexame=request.form("hrexame")
if request.form("hrexame2")<>"" then hrexame2=request.form("hrexame2")
if request.form("preposto")<>"" then preposto=request.form("preposto") else preposto="02918"
if request.form("motivodrt")<>"" then motivodrt=request.form("motivodrt")
%>
<table border=0 cellpadding=5 width=500 bordercolor=black style="border-collapse:collapse">
<tr><td valign=top class=fundo style="border-bottom:2 solid;border-right:2 solid">

<table border="0" bordercolor=#000000 cellpadding="2" cellspacing="0" style="border-collapse:collapse">
<tr>
	<td class=fundo><b>Data de Admissão:</td><td class=fundo align="center"><%=dataadmissao%></td>
</tr>
<tr>
	<td class=fundo><b>Data de Saída:</td><td class=fundo><input style="text-align:center" type="text" name="datasaida" size=10 value="<%=datasaida%>" onchange="javascript:submit();"></td>
</tr>
<tr>
	<td class=fundo>Tempo de Serviço:</td><td class=fundo align="center"><%if datasaida<>"" then ts=(cdate(datasaida)-cdate(dataadmissao))/365.25 else ts=0%><%="+" & int(ts) & " anos"%></td>
</tr>
</table>

</td>
<td valign=top class=fundo style="border-bottom:2 solid">

<table border="0" bordercolor=#000000 cellpadding="2" cellspacing="0" style="border-collapse:collapse" height=100%>
<tr>
	<td class=fundo valign=middle><b>Local Homologação</td>
	<td class=fundo>
<%if ts>0 and ts<1 then%>	
	<input type="radio" name="localpag" value="fieo" <%if localpag="fieo" then response.write "checked"%> >Recursos Humanos<br>
<%elseif ts>=1 then%>
	<input type="radio" name="localpag" value="sindicato" <%if localpag="sindicato" then response.write "checked"%> onclick="javascript:submit();">Sindicato/Federação<br>
	<input type="radio" name="localpag" value="drt" <%if localpag="drt" then response.write "checked"%> onclick="javascript:submit();">DRT
<%else%>
	<font color=red><b>Informe a data de saida.
<%end if%>
	</td>
</tr>
</table>                 

</td></tr>
<tr><td valign=top class=fundo style="border-right:2 solid">

<table border="0" bordercolor=#000000 cellpadding="2" cellspacing="0" style="border-collapse:collapse">
<tr><td class=fundo><b>Data da Homologação/Pagamento</td></tr>
<tr><td class=fundo><input style="text-align:center" type="text" name="dthomologacao" size=10 value="<%=dthomologacao%>"> às
<input style="text-align:center" type="text" name="hrhomologacao" size=6 value="<%=hrhomologacao%>">
	</td></tr>
</table>

</td><td valign=top class=fundo>

<table border="0" bordercolor=#000000 cellpadding="2" cellspacing="0" style="border-collapse:collapse">
<tr><td class=fundo><b>Data do Exame Médico</td><td class=fundo><input type="radio" name="localexame" value="osasco" <%if localexame="osasco" then response.write "checked"%>>Osasco</td></tr>
<tr><td class=fundo><input style="text-align:center" type="text" name="dtexame" size=10 value="<%=dtexame%>"> às
<input style="text-align:center" type="text" name="hrexame" size=6 value="<%=hrexame%>">-
<input style="text-align:center" type="text" name="hrexame2" size=6 value="<%=hrexame2%>">
	</td><td class=fundo><input type="radio" name="localexame" value="brigadeiro" <%if localexame="brigadeiro" then response.write "checked"%>>B.Funda</td></tr>
<tr><td class=fundo></td><td class=fundo><input type="radio" name="localexame" value="paltino" <%if localexame="paltino" then response.write "checked"%>>Outros</td></tr>
</table>

</td></tr></table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width=500>
<tr>
	<td class=fundo>Chapa Preposto: <input style="text-align:center" type="text" name="preposto" size=6 value="<%=preposto%>">
	<td class=fundo valign=top>
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width=500>
<tr>
	<td class=fundo align="center"><input type="submit" value="Imprimir Aviso" name="B1" class="button"></td>
</tr>
</table>

</form>

<%
end if 'request.form("B1")

if request.form("B1")<>"" then
	'response.write request.form 
	chapa=request.form("chapa")
	datasaida=request.form("datasaida")
	localpag=request.form("localpag") 'sindicato/drt/fieo
	dthomologacao=request.form("dthomologacao")
	hrhomologacao=request.form("hrhomologacao")
	localexame=request.form("localexame") 'osasco/brigadeiro
	dtexame=request.form("dtexame")
	hrexame=request.form("hrexame")
	hrexame2=request.form("hrexame2")
	preposto=request.form("preposto")
	motivodrt=request.form("motivodrt")
	if preposto="" then preposto="02813"

sql1="select sexo, nome from qry_funcionarios where chapa='" & preposto & "' "
rsi.Open sql1, ,adOpenStatic, adLockReadOnly
nomepreposto=rsi("nome")
sexopreposto=rsi("sexo")
rsi.close
if sexopreposto="F" then b1="a " else b1="o "
if sexopreposto="F" then b2="a. " else b2=". "

	dataextenso=day(datasaida) & " de " & monthname(month(datasaida)) & " de " & year(datasaida)
	sql1="select f.chapa, f.nome, f.admissao, f.sexo, f.carteiratrab, f.seriecarttrab, f.ufcarttrab, " & _
	"f.codsecao, f.secao, f.codsindicato from qry_funcionarios f where f.chapa='" & chapa & "' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	
	if rs("sexo")="F" then str1="Sra. " else str1="Sr. "

local03="Sindicato dos Professores de Osasco (SINPRO)"
local01="Federação dos Trabalhadores em Estabelecimento de Ensino de São Paulo (FETEE)"
local01="Sindicato dos Auxiliares de Administração Escolar de São Paulo (SAAESP)"
endereco03="Av. Deputado Emilio Carlos, 937 - Osasco - SP"
endereco01="Rua das Cassuarinas, 109 - Jd. Oriental - SP"
endereco01="Rua Tenente Avelar Pires de Azevedo, 289 Sala 13 - Centro - Osasco"

select case localpag
	case "fieo"
		local="Departamento de Recursos Humanos"
		endereco="Av. Franz Voegelli, 300 - Osasco - SP"
	case "sindicato"
		if rs("codsindicato")="03" then
			local=local03
			endereco=endereco03
		else
			local=local01
			endereco=endereco01
		end if
	case "drt"
		local="Ministério do Trabalho"
		endereco="Rua Narciso Sturlini, 124 - Osasco - SP"
		endereco="Rua Santa Teresinha, 59 - Osasco - SP"
end select
if dthomologacao="" then
	dthomo2=" no dia ____/_____/______ às _____:_____ horas"
else
	dthomo2=" no dia " & dthomologacao & " às " & hrhomologacao & " horas"
end if

dtexame21=" na Rua Itabuna, 93 - Centro de Osasco"
dtexame22=" na Av. Thomas Edison, 305 - Barra Funda - SP"
dtexame23=" em um dos outros locais disponíveis no Recursos Humanos"
if dtexame="" then
	dtexame2=" poderá ser realizado e/ou agendado nos seguintes endereços:<br>"
	dtexame2=dtexame2 & "• " & dtexame21 & " através do telefone 3184-0099<br>"
	dtexame2=dtexame2 & "• " & dtexame22 & " através do telefone 3392-1305<br>"
	dtexame2=dtexame2 & "• " & dtexame23 & ""
else
	if hrexame2="" then 
		txt0=" às "
		txt1=hrexame
	else
		txt0=" das "
		txt1=hrexame & " às " & hrexame2 & " horas"
	end if
	dtexame2=" será realizado no dia " & dtexame & txt0 & txt1
	select case localexame
		case "osasco"
			end_exame=dtexame21
		case "brigadeiro"
			end_exame=dtexame22
		case "paltino"
			end_exame=dtexame23
			end_exame=dtexame21
	end select
end if
datalimite=dateserial(year(datasaida),month(datasaida),day(datasaida)+10)
if weekday(datalimite)=1 then datalimite=datalimite-2
if weekday(datalimite)=7 then datalimite=datalimite-1

if rs("sexo")="M" then a1="o" else a1="a"
if rs("codsindicato")="03" then a2="o" else a2="a"
if rs("codsindicato")="03" then categoria=local03 else categoria=local01
if rs("codsindicato")="03" then enderecosind=endereco03 else enderecosind=endereco01
%>	
<%
'**************** carta
%>
<div align="center"><center>
<table border="0" cellpadding="5" width="620" cellspacing="0" height="1000">
<!-- linha da aguia -->
<tr><td height=112><img src="../images/aguia.jpg" border="0" width="236" height="111" alt=""></td></tr>
<!-- linha intermediaria -->
<tr><td height="20">&nbsp;</td></tr>
<!-- linha declaracao -->
<tr><td height=50 valign="center" align="left"><font size="3">Osasco, <%=dataextenso%></td></tr>
<!-- linha intermediaria -->
<tr><td height="20">&nbsp;</td></tr>
<!-- linha declaracao -->
<tr><td height=50 valign="center" align="left"><font size="3">Sr<%=a1%>. <%=rs("nome")%></td></tr>
<!-- corpo da declaracao -->
<tr><td height=100% valign=top>
<%
%>	
	<p>&nbsp;</p>
<p style="margin-top:0;margin-bottom:0;text-align:justify;font-size:12pt">
<br>
<br>Dando continuidade ao seu processo de desligamento, solicitamos o seu comparecimento ao endereço abaixo para a
realização do exame médico demissional, o qual é obrigatório por lei e que, na falta, a homologação não poderá ser 
realizada. O agendamento poderá ser feito por telefone, de acordo com a sua preferência.
<br>
<br>Informamos que a homologação será realizada no <%=local%>, sito à <%=endereco%><%=dthomo2%>.
<br>
<br>O exame médico demissional <%=dtexame2%> <%=end_exame%>.
<br>
<br>


<!-- tabela data e assinatura -->
<table border="0" cellpadding="0" width="100%" cellspacing="0">
<tr>
	<td valign=top width=50%><p style="margin-top:0;margin-bottom:0;text-align:right;font-size:12pt">
	Atenciosamente&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><Br><br>
	</td>
</tr>
<tr>
	<td valign="top"><p style="margin-top:0;margin-bottom:0;text-align:justify;font-size:12pt">
Ciente: ______/_______/________
<br><br>Assinatura:______________________________________
	</td>

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
%>

<%
end if 'request.form("B1")
%>
</body>
</html>
<%
conexao.close
set conexao=nothing
%>