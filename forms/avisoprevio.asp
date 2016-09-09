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
<title>Opções de Aviso Prévio</title>
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
if request.form("geraCRM")="ON" then geraCRM=1 else geraCRM=0
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>
Opções e parametros para emissão de Aviso Prévio</p>
<form method="POST" action="avisoprevio.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width=500>
<tr>
	<td class=fundo>Chapa <input type="text" name="chapa" size="5" class="form_box" onchange="chapa1()" value="<%=request.form("chapa")%>"></td>
	<td class=fundo><select name="nome" onchange="nome1()">
	<option value="00000">Selecione...</option>
<%
sql="select chapa, nome from corporerm.dbo.pfunc where (codsituacao in ('A','F','Z') and codtipo='N') or (codsituacao in ('D') and datademissao>getdate()-15) order by nome"
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
	</select>
	<input type="checkbox" name="geraCRM" value="ON" <%if geraCRM=1 then response.write "checked"%>  >Gera CRM?
	</td>
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

if request.form("dataaviso")<>"" then dataaviso=cdate(request.form("dataaviso"))
if dtdemissao<>"" then dataaviso=dtdemissao
if request.form("localpag")<>"" then localpag=request.form("localpag")
if request.form("dthomologacao")<>"" then dthomologacao=request.form("dthomologacao")
if request.form("hrhomologacao")<>"" then hrhomologacao=request.form("hrhomologacao")
if request.form("localexame")<>"" then localexame=request.form("localexame")
if request.form("dtexame")<>"" then dtexame=request.form("dtexame")
if request.form("hrexame")<>"" then hrexame=request.form("hrexame")
if request.form("hrexame2")<>"" then hrexame2=request.form("hrexame2")
if request.form("preposto")<>"" then preposto=request.form("preposto") else preposto="03062"
if request.form("motivodrt")<>"" then motivodrt=request.form("motivodrt")
if request.form("tipoaviso")<>"" then tipoaviso=request.form("tipoaviso")

anos=datediff("yyyy", dataadmissao, dataaviso)
teste=dateadd("yyyy",anos,dataadmissao)
if teste>dataaviso then dajuste=-1 else dajuste=0
anosp=anos-0+dajuste
diasap=anosp*3
if diasap<0 then diasap=0
if diasap>60 then diasap=60
dataproj=dataaviso+(diasap+30)
datasaida=dataaviso+30
database=dateserial(year(dataproj),3,1)
databasei=database-30
if dataproj=>databasei and dataproj<=database then antecede="SIM" else antecede="NÃO"
session("apdiasap")=diasap+30
session("apantecede")=left(antecede,1)

%>

<table border=0 cellpadding=5 width=500 bordercolor=black style="border-collapse:collapse">
<tr><td valign=top class=fundo style="border-bottom:2 solid;border-right:2 solid">

<table border="0" bordercolor=#000000 cellpadding="2" cellspacing="0" style="border-collapse:collapse">
<tr>
	<td class=fundo><b>Data de Admissão:</td><td class=fundo align="center"><%=dataadmissao%></td>
</tr>
<tr>
	<td class=fundo>Tipo Aviso</td><td class=fundo align="left">
	<input type="radio" name="tipoaviso" value="I" <%if tipoaviso="I" then response.write "checked"%> onclick="javascript:submit();">Indenizado<br>
	<input type="radio" name="tipoaviso" value="T" <%if tipoaviso="T" then response.write "checked"%> onclick="javascript:submit();">Trabalhado<br>
	</td>
</tr>
<tr>
	<td class=fundo><b>Data do Aviso:</td><td class=fundo><input style="text-align:center" type="text" name="dataaviso" size=10 value="<%=dataaviso%>" onchange="javascript:submit();"></td>
</tr>
<tr>
	<td class=fundo><b>Data da Saída:</td>
	<td class=fundo>
<%
if tipoaviso="I" then datasaida=dataaviso
if tipoaviso="T" then datasaida=dataaviso-1+30
%><%=datasaida%>
	<input type="hidden" name="datasaida" value="<%=datasaida%>">
	</td>
</tr>
<tr>
	<td class=fundo><b>Data de Baixa:</td><td class=fundo><%=dataproj%></td>
</tr>
<tr>
	<td class=fundo>Tempo de Serviço:</td><td class=fundo align="center"><%if dataaviso<>"" then ts=(cdate(dataaviso)-cdate(dataadmissao))/365.25 else ts=0%><%="+" & int(ts) & " anos"%></td>
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
	</td>
	<td class=fundo valign=top>
<%if localpag="drt" then%>
	<table border="0" bordercolor=#000000 cellpadding="2" cellspacing="0" style="border-collapse:collapse" height=100%>
	<tr>
	<td class=fundo valign=middle><b>Motivos para DRT</td>
	<td class=fundo>
	<input type="radio" name="motivodrt" value="mot1" <%if motivodrt="mot1" then response.write "checked"%>>O sindicato cobra<br>
	<input type="radio" name="motivodrt" value="mot2" <%if motivodrt="mot2" then response.write "checked"%>>Não há representante na localidade<br>
	<input type="radio" name="motivodrt" value="mot3" <%if motivodrt="mot3" then response.write "checked"%>>Débitos da empresa c/o sindicato
	</td>
	</tr>
	</table>                 
<%end if%>
	
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width=500>
<tr>
	<td class=fundo align="center"><input type="submit" value="Imprimir Aviso" name="B1" class="button"></td>
</tr>
</table>

</form>

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse:collapse" >
<tr>
	<td class=fundo>Anos completos de trabalho:</td><td class=campo><%=anosp%></td>
</tr>
<tr>
	<td class=fundo>Dias a acrescentar:</td><td class=campo><%=diasap%></td>
</tr>
<tr>
	<td class=fundo>Total dias do Aviso Prévio:</td><td class=campo><%=30+diasap%></td>
</tr>
<tr>
	<td class=fundo>Final do aviso projetado:</td><td class=campo><%=dataproj%></td>
</tr>
<tr>
	<td class=fundo>Antecede 30 dias da data-base:
	<br><font style="font-size:8px">(entre <%=databasei%> e <%=database%>)</font></td><td class=campo><%=antecede%></td>
</tr>

</table>

<%

end if 'request.form("B1")

if request.form("B1")<>"" then
	'response.write request.form 
	chapa=request.form("chapa")
	dataaviso=request.form("dataaviso")
	localpag=request.form("localpag") 'sindicato/drt/fieo
	dthomologacao=request.form("dthomologacao")
	hrhomologacao=request.form("hrhomologacao")
	localexame=request.form("localexame") 'osasco/brigadeiro
	dtexame=request.form("dtexame")
	hrexame=request.form("hrexame")
	hrexame2=request.form("hrexame2")
	preposto=request.form("preposto")
	motivodrt=request.form("motivodrt")
	datasaida=request.form("datasaida")
	tipoaviso=request.form("tipoaviso")
	if tipoaviso="I" then limite=10 else limite=1
	datalimite=dateserial(year(datasaida),month(datasaida),day(datasaida)+limite)
	if weekday(datalimite)=1 then datalimite=datalimite-2
	if weekday(datalimite)=7 then datalimite=datalimite-1

	if preposto="" then preposto="02918"

sql1="select sexo, nome from qry_funcionarios where chapa='" & preposto & "' "
rsi.Open sql1, ,adOpenStatic, adLockReadOnly
nomepreposto=rsi("nome")
sexopreposto=rsi("sexo")
rsi.close
if sexopreposto="F" then b1="a " else b1="o "
if sexopreposto="F" then b2="a. " else b2=". "

	dataextenso=day(dataaviso) & " de " & monthname(month(dataaviso)) & " de " & year(dataaviso)
	sql1="select f.chapa, f.nome, f.admissao, f.sexo, f.carteiratrab, f.seriecarttrab, f.ufcarttrab, " & _
	"f.codsecao, f.secao, f.codsindicato from qry_funcionarios f where f.chapa='" & chapa & "' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	
	if rs("sexo")="F" then str1="Sra. " else str1="Sr. "
	if rs("sexo")="F" then str2="a" else str2="o"
	if tipoaviso="I" then str_ap="INDENIZADO" else str_ap="TRABALHADO"
%>
<div align="right"><right>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width="620">
<tr>
	<td class=campo align="center"><p style="font-size:12pt;margin-top: 0; margin-bottom: 0">
	<b>AVISO PRÉVIO DO EMPREGADOR</p>(<%=str_ap%>)</td>
</tr>
</table>
<br>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width="620" height=890>
<tr>
	<td class="campop" valign=top><p style="margin-top:0;margin-bottom:0;text-align:justify;font-size:12pt">
<br>Osasco, <%=dataextenso%><br>
<br>
<br><%=str1%> <b><%=rs("nome")%></b>
<br>CTPS nº <%=rs("carteiratrab")%> / <%=rs("seriecarttrab")%> / <%=rs("ufcarttrab")%>
	<%for a=1 to 10:response.write "&nbsp;":next%><br>SEÇÃO/DEPTO: <%=rs("secao")%>
<br>
<br>
<%if rs("codsindicato")="03" then%>
<!--
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Como é de conhecimento público, as instituições de ensino superior enfrentam situação difícil diante da
	concorrência predatória causada pela autorização indiscrimanada da abertura de novas faculdades e cursos a
	preços irrisórios, sem o compromisso com a qualidade de ensino que sempre norteou os rumos do UNIFIEO.
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tal quadro lamentavelmente reflete nesta casa de ensino, que, verificando a diminuição da procura em seus
	cursos, vê-se obrigada a redimensionar seu quadro funcional para evitar mal maior.
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Por tais motivos e sem qualquer demérito ao trabalho que V.Sa. desenvolveu por todo o tempo em que aqui 
	laborou, não resta outra alternativa a não ser comunicar sua dispensa das funções exercidas a partir de hoje, 
	ficando, portanto, convocado a comparecer no Departamento de Recursos Humanos, 
	de posse da Carteira de Trabalho e Previdência Social, crachá de identificação, cartão de estacionamento, 
	cartões da Assistência Médica e outros pertences da empresa que porventura estejam em seu poder.
-->
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Como é de conhecimento geral, as instituições particulares de ensino superior 
	vêm passando por momento de difícil solução, decorrente de vários fatores, entre os quais sobrelevam a 
	competição nem sempre comedida de entidades congêneres, os índices de inadimplência e evasão do alunado, 
	a exigência de se conceder um número elevado de bolsas a alunos carentes.
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Por outro lado, a preservação da qualidade do ensino aqui ministrado norteou 
	os rumos do UNIFIEO, no sentido de garantir essa continuidade, a qual não pode ser atingida ou alterada; 
	mas, para enfrentar estas dificuldades, há que adotar providências que venham melhor equacionar a oferta de 
	seus cursos, redimensionando, por sua vez, os seus quadros docentes.
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Por tais motivos, vimo-nos na contigência de dispensá-l<%=str2%> das funções 
	exercidas a partir de hoje, solicitando o seu comparecimento no Departamento de Recursos Humanos, de posse 
	da Carteira de Trabalho e Previdência Social, crachá de identificação, cartão de estacionamento, 
	cartões da Assistência Médica e outros pertences da empresa que porventura estejam em seu poder.
<%else%>
	<%if tipoaviso="I" then%>
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Como é de conhecimento geral, as instituições particulares de ensino superior 
	vêm passando por momento de difícil solução, decorrente de vários fatores, entre os quais sobrelevem os 
	índices de inadimplência e evasão do alunado. Por tais motivos, vimo-nos na contigência de rever a atual 
	estrutura do quadro administrativo, resultando na dispensa das funções exercidas por V.Sa., a partir de hoje, 
	solicitando o seu comparecimento no Departamento de Recursos Humanos, de posse da Carteira de Trabalho 
	e Previdência Social, crachá de identificação, cartão de estacionamento, cartões da Assistência Médica 
	e outros pertences da empresa que porventura estejam em seu poder.
	<%else%>
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Como é de conhecimento geral, as instituições particulares de ensino superior 
	vêm passando por momento de difícil solução, decorrente de vários fatores, entre os quais sobrelevem os 
	índices de inadimplência e evasão do alunado. Por tais motivos, vimo-nos na contigência de rever a atual 
	estrutura do quadro administrativo, resultando na dispensa das funções exercidas por V.Sa., 
	em <%=datasaida%>, ou seja, 30 dias a contar desta data.
	<%end if%>
<%end if 'sindicato%>
<%if tipoaviso="T" then%>
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Durante o período de cumprimento do aviso prévio, sua jornada de trabalho poderá
	ser reduzida, sem prejuízo da remuneração, da seguinte maneira:
	<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;] redução de 2 (duas) horas diárias em seu horário normal de trabalho;
	<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[&nbsp;&nbsp;&nbsp;] redução de 7 (sete) dias corridos no período de <%=cdate(datasaida)-6%> a <%=datasaida%>.
<br>
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;No dia seguinte ao término do seu aviso prévio, em <%=datalimite%>, deverá
	V.Sa. comparecer de posse da Carteira de Trabalho e Previdência Social, crachá de identificação, 
	cartão de estacionamento, cartões da Assistência Médica e outros pertences da empresa que porventura 
	estejam em seu poder para o cumprimento das formalidades legais exigidas para a rescisão contratual.
<%end if%>
<%
local03="Sindicato dos Professores de Osasco (SINPRO)"
local01="Federação dos Trabalhadores em Estabelecimento de Ensino de São Paulo (FETEE)"
local01="Sindicato dos Auxiliares de Administração Escolar de São Paulo (SAAESP)"
local01="Sindicato dos Auxiliares de Administração Escolar de Osasco e Região (SAAEO)"
endereco03="Av. Deputado Emilio Carlos, 937 - Osasco - SP"
endereco01="Rua das Cassuarinas, 109 - Jd. Oriental - SP"
endereco01="Rua Tenente Avelar Pires de Azevedo, 289 Sala 13 - Centro - Osasco"
endereco01="Rua Mariano J. M. Ferraz, 125 Sala 12 - Centro - Osasco"

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
	dthomo2=". O dia e horário será comunicado posteriormente"
else
	dthomo2=", no dia " & dthomologacao & " às " & hrhomologacao & " horas"
end if

dtexame21=" na Rua Itabuna, 93 - Centro de Osasco"
dtexame22=" na Av. Thomas Edison, 305 - Barra Funda - SP"
dtexame23=" em um dos outros locais disponíveis no Recursos Humanos"
if dtexame="" then
	dtexame2=" poderá ser agendado e realizado nos seguintes endereços:<br>"
	dtexame2=dtexame2 & "• " & dtexame21 & " através do telefone 3184-0099<br>"
	dtexame2=dtexame2 & "• " & dtexame22 & " através do telefone 3392-1305"
	dtexame2=dtexame2 & "• " & dtexame23 & ""
	dtexame2=" será agendado pelo Recursos Humanos e a guia lhe será entregue quando do seu comparecimento"
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
%>	
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;A homologação e quitação será realizada no <%=local%>, sito à <%=endereco%><%=dthomo2%>. 
	O pagamento das verbas rescisórias será creditado em sua conta-corrente até o dia <%=datalimite%>.
<br>
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;O exame médico demissional <%=dtexame2%> <%=end_exame%>.
<br>
<%if rs("codsindicato")="03" then%>
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Esperando contar, no futuro, com a colaboração prestada por V.Sa. no período em que trabalhou nesta instituição,
	subscrevemo-nos,
<%else%>
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Agradecendo sua colaboração no período em que trabalhou nesta instituição, subscrevemo-nos,
<%end if 'sindicato%>

<br>
<br>Atenciosamente,
<br>
<br>________________________________________
<br>Empregador
<br>
<br>
<br>________________________________________
<br>Ciente do Empregado
<br>
<br>Data:_____/_____/______
	
	</td>
</tr>
<tr><td class="campor" align="right"><%=session("apdiasap")&"-"&session("apantecede")%></td></tr>
</table>
</right></div>

<%
'**************** carta para DRT
if request.form("localpag")="drt" then
response.write "<DIV style=""page-break-after:always""></DIV>"
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
<tr><td height=50 valign="center" align="left"><font size="3">À<br>Subdelegacia do Trabalho em Osasco</td></tr>
<!-- corpo da declaracao -->
<tr><td height=100% valign=top>
<%
if rs("sexo")="M" then a1="o" else a1="a"
if rs("codsindicato")="03" then a2="o" else a2="a"
if rs("codsindicato")="03" then categoria=local03 else categoria=local01
if rs("codsindicato")="03" then enderecosind=endereco03 else enderecosind=endereco01
select case motivodrt
	case "mot1"
		motivo="o representante sindical, " & categoria & " situado a " & enderecosind & " cobra pela homologação o valor de R$ ____,00"
	case "mot2"
		motivo="o representante sindical, " & categoria & " situado a " & enderecosind & " não possui unidade nesta localidade"
	case "mot3"
		motivo="o representante sindical, " & categoria & " não homologa alegando que a empresa tem débitos"
end select
%>	
	<p>&nbsp;</p>
	<p align="justify"><font size="3">A Fundação Instituto de Enisno para Osasco - FIEO, telefone 3651-9972, estabelecida
	à Rua Narciso Sturlini, 883 - Jd. Umuarama - Osasco - SP, CEP 06018-903, com CNPJ nº 73.063.166/0001-20, vem pelo presente,
	requerer que seja feita a homologação d<%=a1%> seguinte ex-funcionári<%=a1%>:
	<br>
	<br>01 - <%=rs("nome")%>
	<br>
	<br>representad<%=a1%> pel<%=a2%>&nbsp;<%=categoria%>.
	<br>
	<br>Para tanto nomeia como preposto, <%=b1%> Sr<%=b2%> <%=nomepreposto%>.
	<br>
	<br>Tal solicitação prende-se ao fato de que <%=motivo%>.

	<p align="justify">&nbsp;</p>

<!-- tabela data e assinatura -->
	<table border="0" cellpadding="0" width="100%" cellspacing="0">
	<tr>
		<td width="50%" valign="top">
		<p>&nbsp;</p>
		<p><font size="3">_____________________________________<br>
		DEPTO DE RECURSOS HUMANOS</font></p>
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
end if
response.write "<DIV style=""page-break-after:always""></DIV>"

sqlu="select top 1 c.id_cat, u.descricao from uniforme_func_cat c inner join uniforme_categoria u on u.id_cat=c.id_cat " & _
"where chapa='" & rs("chapa") & "' and inicio<GETDATE() order by inicio desc"
rsi.Open sqlu, ,adOpenStatic, adLockReadOnly
if rsi.recordcount>0 then categoria=rsi("descricao") else categoria=""
rsi.close

if categoria<>"" then
for a=1 to 2
%>
<div align="center">
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width="620">
<tr><td class=titulop style="border: 1px solid;" align="center" height=40 valign=middle>RECIBO - DEVOLUÇÃO DE UNIFORME</td></tr>
<tr><td height=5 style="border:0 solid;"></td></tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width="620">
<tr>
	<td class="campop" height=45 valign=top><span style="font-size:9px;"><b>Chapa</b></span><br><br>&nbsp;<%=rs("chapa")%></td>
	<td class="campop" valign=top><span style="font-size:9px;"><b>Nome do funcionário</b></span><br><br>&nbsp;<%=rs("nome")%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width="620">
<tr>
	<td class="campop" height=45 valign=top><span style="font-size:9px;"><b>Data da saída</b></span><br><br>&nbsp;<%=dataaviso%></td>
	<td class="campop" valign=top><span style="font-size:9px;"><b>Categoria do uniforme</b></span><br><br>&nbsp;<%=categoria%></td>
</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width="620">
<tr><td height=5 colspan=2 style="border:0 solid;"></td></tr>
<tr><td class=fundo colspan=2 height=25 align="left" style="border: 1px solid"><b>PARA USO DO SUPRIMENTOS</b></td></tr>
<tr><td class=campo colspan=2 valign=top align="center" style="border: 1px solid">&nbsp;

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width="90%">
<tr><td height=35>[&nbsp;&nbsp;] Devolveu todas as peças.</td>
	<td width=20></td>
	<td style="border: 1px solid" colspan=3>Peças não devolvidas</td>
</tr>
<tr><td height=35>[&nbsp;&nbsp;] Devolveu ______ peças.</td>
	<td width=20>--></td>
	<td style="border: 1px solid" width=30> </td>
	<td style="border: 1px solid" width=75> x R$ 10,00 </td>
	<td style="border: 1px solid" width=95> = R$ </td>
</tr>
<tr><td height=35>[&nbsp;&nbsp;] Não devolveu nenhuma peça.</td>
	<td width=20>--></td>
	<td style="border: 1px solid"> </td>
	<td style="border: 1px solid"> x R$ 10,00 </td>
	<td style="border: 1px solid"> = R$ </td>

</tr>
</table>
&nbsp;

</td></tr>
<tr>
	<td class="campop" valign=top style="border: 1px solid"><span style="font-size:9px;"><b>Data</b></span><br>&nbsp;</td>
	<td class="campop" valign=top style="border: 1px solid"><span style="font-size:9px;"><b>Confirmo as informações acima</b><br><br>&nbsp;Assinatura - Suprimentos</span></td>
</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width="620">
<tr><td class="campor">
<%
if a=1 then
	response.write "<p style=""font-size:9px;margin-top:0px""><b>Via Funcionário:</b> Entregar esta via junto com as peças de uniforme devolvidas.</p>"
else 
	response.write "<p style=""font-size:9px;margin-top:0px""><b>Via Suprimentos:</b> Um dia após a data da saída enviar esta via ao RH, preenchida e assinada e com o número de uniformes faltantes, se for o caso.</p>"
end if
%>
</td></tr></table>
</div>
<%
if a=1 then
	response.write "<br><br><br><hr><br>"
end if
next
end if 'categoria de uniforme

response.write "<DIV style=""page-break-after:always""></DIV>"
%>

<div align="center"><center>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse:collapse" width="620">
<tr>
	<td class="campop" align="center"><font size=4>DOCUMENTOS NECESSÁRIOS PARA HOMOLAGAÇÃO NA DRT</td>
</tr>
<tr>
	<td class="campop"><font size=3>
<ol>
	<li>Carta de Preposto em 2 (duas) Vias sem rasuras.</li>
	<li>Termo de Rescisão de Contrato em 4 (quatro) Vias. Justificar no verso das 4 vias, o porque está homologando na DRT e não no sindicato da categoria.</li>
	<li>Acordo, Convenção ou Dissídio Coletivo completo.</li>
	<li>Extrato atualizado do FGTS, e Guias de Recolhimento dos meses que não constam no extrato.</li>
	<li>Ficha ou Livro de Registro e CTPS do empregado atualizada.</li>
	<li>Pagamento em:</li>
	<blockquote>
		<ul type="disc">
  		<li>cheque administrativo nominal ao empregado</li>
		<li>dinheiro</li>
		<li>prova bancária de quitação</li>
		</ul>
	</blockquote>
	<li>Comunicação da Dispensa (CD) e Requerimento do Seguro Desemprego para fins de habilitação, quando devido.</li>
	<li>Exame Médico Demissional.</li>
	<li>Guia de Recolhimento Rescisório do FGTS e da Contribuição Social (Multa do FGTS).</li>
	<li>Demonstrativo de parcelas variáveis para fins de cálculo das verbas rescisórias.</li>
	<li>As empresas cadastradas no SIMPLES deverão apresentar o comprovante no ato da homologação.</li>
</ol>	
	</td>
</tr>
<tr>
	<td class="campop"><font size=4>ATENÇÃO:<br><font size=3>Na falta de algum item acima, será impossível realizar a homologação.
	</td>
</tr>
</table>

<%
if request.form("geraCRM")="ON" or geraCRM=1 then
	if tipoaviso="T" then prazo=1 else prazo=10
	dataaviso=dataaviso
	datasaida=datasaida
	datapagto=datalimite
	sqlcrm="insert into iCRM_Fluxo (idCRM, chapaC, Chapa, DtFluxo, Anotacao, DtVencimento, Status, create_user, create_data) "
	sqlr001=sqlcrm & "select 'R001', '" & session("usuariomaster") & "', '" & chapa & "', '" & dtaccess(dataaviso) & "',null,'" & dtaccess(dataaviso) & "','A', '" & session("usuariomaster") & "', getdate() "
	conexao.execute sqlr001
	
sqlins="select idCRM from iCRM_Atividades where Atividade='Rescisão' and idCRM not in ('R001')"
rsi.Open sqlins, ,adOpenStatic, adLockReadOnly
do while not rsi.eof
	sqlr=sqlcrm & "select '" & rsi("idCRM") & "', '" & session("usuariomaster") & "', '" & chapa & "', '" & dtaccess(dataaviso) & "',null,'" & dtaccess(datalimite) & "','A', '" & session("usuariomaster") & "', getdate() "
	conexao.execute sqlr
rsi.movenext
loop
rsi.close

end if 'geraCRM

end if 'request.form("B1")
%>
</body>
</html>
<%
conexao.close
set conexao=nothing
%>