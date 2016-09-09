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
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Declaração Opcional de Plano de Saúde</title>
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
set conexao2=server.createobject ("ADODB.Connection")
conexao2.Open application("conexao")

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao2

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

	sqla="SELECT f.NOME, f.CODSITUACAO, f.CHAPA, f.DATAADMISSAO, f.CODSECAO, f.codsindicato, s.DESCRICAO AS Secao, " & _
	"p.dtnascimento, p.telefone1, p.telefone2, p.telefone3, p.email, p.cpf, p.estadocivil, c.nome as funcao, " & _
	"p.cartidentidade, p.cpf, p.dtnascimento, p.sexo, p.rua, p.numero, p.complemento, p.bairro, p.cidade, p.cep, p.estado, " & _
	"p.telefone1, f.datademissao, f.dtaposentadoria, f.aposentado, f.tipodemissao, p.grauinstrucao " & _
	"FROM corporerm.dbo.PFUNC f, corporerm.dbo.PSECAO s, corporerm.dbo.PPESSOA p, corporerm.dbo.PFUNCAO c " & _
	"WHERE f.CODSECAO=s.CODIGO and p.codigo=f.codpessoa and c.codigo=f.codfuncao "

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
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para Declaração opcional de Plano de Saúde - INTERMéDICA
<form method="POST" action="decl_copart_intermedica.asp">
	<p style="margin-top: 0; margin-bottom: 0">Chapa/Nome <input type="text" name="chapa" size="20" class="form_box" value="<%=session("chapa")%>">
	<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>

<%
elseif temp=0 then
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
idade=int((now()-rs("dtnascimento"))/365.25)
if rs("datademissao")="" or isnull(rs("datademissao")) then rsdatademissao=now() else rsdatademissao=rs("datademissao")

sqlplano="SELECT codigo, plano FROM assmed_mudanca " & _
"WHERE chapa='" & rs("chapa") & "' and '" & dtaccess(rsdatademissao) & "' between ivigencia and fvigencia "
rs3.Open sqlplano, ,adOpenStatic, adLockReadOnly
plano=rs3("plano")
carteirinha=rs3("codigo")
rs3.close
sqlmae="select nome from corporerm.dbo.pfdepend where chapa='"& rs("chapa") & "' and grauparentesco='7'"
rs3.Open sqlmae, ,adOpenStatic, adLockReadOnly
mae=rs3("nome")
rs3.close

if rs("aposentado")=1 then 
	Sapos="&nbsp;X&nbsp;":Napos="&nbsp;&nbsp;&nbsp;&nbsp;"
else 
	Sapos="&nbsp;&nbsp;&nbsp;&nbsp;":Napos="&nbsp;X&nbsp;"
end if
if rs("tipodemissao")="2" or rs("tipodemissao")="A" then 
	Sdem="&nbsp;X&nbsp;":Ndem="&nbsp;&nbsp;&nbsp;&nbsp;"
else 
	Sdem="&nbsp;&nbsp;&nbsp;&nbsp;":Ndem="&nbsp;X&nbsp;"
end if

idade=int((now()-rs("dtnascimento"))/365.25)
dia4=day(rs("dtaposentadoria")):if dia4="" or isnull(dia4) then dia4="  " else dia4=numzero(dia4,2)
mes4=month(rs("dtaposentadoria")):if mes4="" or isnull(mes4) then mes4="  " else mes4=numzero(mes4,2)
ano4=year(rs("dtaposentadoria")):if ano4="" or isnull(ano4) then ano4="  " else ano4=right(ano4,2)
dtaposent=dia4&mes4&ano4

'052 desconto co-participação 076 desconto assistencia médica
sqlp="select count(chapa) as vezes from corporerm.dbo.pffinanc where codevento IN ('076','076I','076U') and chapa='" & rs("chapa") & "' "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
meses2=rs3("vezes")
if meses2="" or isnull(meses2) then meses2=0
rs3.close
sqlp="select count(chapa) as vezes from corporerm.dbo.pffinanccompl where codevento IN ('076','076I','076U') and chapa='" & rs("chapa") & "' "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
mes=rs3("vezes")
if mes="" or isnull(mes) then mes=0
meses2=meses2+mes
rs3.close
sqlp="select count(chapa) as vezes from corporerm.dbo.pffinanc where codevento IN ('052','052I') and chapa='" & rs("chapa") & "' "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
meses=rs3("vezes")
if meses="" or isnull(meses) then meses=0
rs3.close
sqlp="select count(chapa) as vezes from corporerm.dbo.pffinanccompl where codevento IN ('052','052I') and chapa='" & rs("chapa") & "' "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
if mes="" or isnull(mes) then mes=0
meses=meses+mes
rs3.close
if meses2>meses then maior=meses2 else maior=meses
cano=int((meses+meses2)/12)
cmes=(meses+meses2)-(cano*12)
dini=dtdemissao
sqlp="select max(valor) ultima from corporerm.dbo.pffinanc where codevento in ('052','052U','052I','076','076I','076U') and chapa='" & rs("chapa") & "' " & _
"--and mescomp=" & month(rs("datademissao")) & " and anocomp=" & year(rs("datademissao")) & " "
rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
ultimo=rs3("ultima")
if ultimo="" or isnull(ultimo) then ultimo=0
rs3.close

if meses2>0 or meses>0 then 
	Scont="&nbsp;X&nbsp;":Ncont="&nbsp;&nbsp;&nbsp;&nbsp;"
else 
	Scont="&nbsp;&nbsp;&nbsp;&nbsp;":Ncont="&nbsp;X&nbsp;"
end if

%>

<div align="center">
<center>
<!-- ----------------------------- -->
<table border="0" cellpadding="5" width="690" cellspacing="0" height="1000">
<!-- linha da aguia -->
<tr><td height=112><img src="../images/aguia.jpg" border="0" width="236" height="111" alt=""></td></tr>
<!-- corpo da declaracao -->
<tr><td height=100% valign=top style="text-align:justify">
<!-- ----------------------------- -->
<p><b>DECLARAÇÃO DE CIÊNCIA SOBRE OS DIREITOS DOS ARTs. 30 e 31 DA LEI DE PLANOS DE SAÚDE - LEI Nº 9656/98</b>
<p>
<p style="margin-bottom:2px;margin-top:5px">Eu, <u><b><%=rs("nome")%><%for a=1 to (60-len(rs("nome"))):response.write "&nbsp;":next%></b></u>, 
CPF nº <%=rs("cpf")%>, ex-funcionário contributário do Plano de Assistência à Saúde da
<b>Intermédica Sistema de Saúde S/A</b>, contratado pela FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO 
com a qual mantive vínculo empregatício, declaro que em <u><%=rs("datademissao")%></u>,
data de formalização da comunicação do Aviso Prévio ou da comunicação da Aposentadoria, fui comunicado pela
minha ex-empregadora sobre o direito que a mim são conferidos pelos arts. 30 e 31 da Lei nº 9656/98, regulamentada
pela Resolução Normativa nº 279, da ANS, de 24 de novembro de 2011, alterada pelas Resoluções Normativas
nº 287 e 297/12 da ANS, no seguinte sentido:

<p style="margin-bottom:2px;margin-top:5px">É garantido aos ex-empregados, demitidos ou exonerados sem justa causa ou aposentados que contribuíram 
mensalmente para o pagamento da contraprestação pecuniária do plano privado de assistência à saúde em 
decorrência de vínculo empregatício, o direito de manterem a condição de beneficiários deste plano, 
nas mesmas condições de cobertura assistencial de que gozavam quando da vigência do vínculo de emprego, 
desde que assumam o pagamento integral da respectiva contraprestação pecuniária. 

<p style="margin-bottom:2px;margin-top:5px">Não são consideradas Contribuições, valores pagos pelo Titular:
<blockquote style="margin-top:0;margin-bottom:0">
a) relacionados à contribuição de dependentes e/ou agregados; e
<br>b) correspondentes à co-participação ou franquia paga única e exclusivamente em procedimentos como  fator de moderação, na utilização dos serviços de assistência médica.
</blockquote>

<p style="margin-bottom:2px;margin-top:5px">Tal benefício é extensivo aos dependentes inscritos quando da vigência do emprego, sendo certo que estes 
serão excluídos do contrato no término dos prazos estabelecidos em lei para manutenção do benefício ou na 
hipótese de perderem a condição de dependência prevista no contrato.

<p style="margin-bottom:2px;margin-top:5px">O período de duração do benefício varia de acordo com cada uma das situações abaixo descritas:

<p style="margin-bottom:2px;margin-top:5px">- É garantido ao ex-empregado demitido ou exonerado sem justa causa a manutenção do plano por um período igual
 a 1/3 (um terço) do tempo durante o qual contribuiu para o pagamento da contraprestação pecuniária do plano 
 privado de assistência à saúde, sendo-lhe garantido um período mínimo de 06 (seis) meses e um período máximo 
 de 24 (vinte e quatro) meses;

<p style="margin-bottom:2px;margin-top:5px">- É garantido ao ex-empregado que venha a se aposentar e que tenha contribuído por <b>10 (dez) anos ou mais</b>
para o pagamento da contraprestação pecuniária do plano privado de assistência à saúde, permanecer no plano por 
prazo indeterminado;

<p style="margin-bottom:2px;margin-top:5px">- É garantido ao ex-empregado que venha a se aposentar e que tenha contribuído por <b>menos de 10 (dez) anos</b> 
para o pagamento da contraprestação pecuniária do plano privado de assistência à saúde, permanecer no plano pelo 
 período igual ao número de anos em que participou do plano como contributário do plano.

<p style="margin-bottom:2px;margin-top:5px">Assim, recebido este comunicado, bem como explicações sobre tudo o que nele contido, declaro estar ciente
 e não ter dúvida de que devo exercer meu direito de opção pela manutenção ou não no plano privado de 
 assistência à saúde em até 30 (trinta) dias, contados da comunicação do Aviso Prévio ou da comunicação da 
 Aposentadoria, estando ciente também de que minha opção deverá ser manifestada por meio de Declaração de Opção 
 de Continuidade. 

<p>___________________, _____/______/________&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ciente, _______________________________________
<p style="margin-bottom:0px;margin-top:0px">
Local<%for a=1 to 30:response.write "&nbsp;":next%>data
<%for a=1 to 60:response.write "&nbsp;":next%>Assinatura do Beneficiário Titular

<p>____________________________________________
<p style="margin-bottom:0px;margin-top:0px">Assinatura da Contratante sob carimbo

<!-- ----------------------------- -->
</td></tr>
<!-- linha intermediaria -->
<tr><td height="5">&nbsp;</td></tr>
<tr><td height="1"><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
<tr><td height="1"><font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000 - C.N.P.J. 73.063.166/0001-20</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP 06020-190 - Fone: (011) 3651-9999 - C.N.P.J. 73.063.166/0003-92</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 1743 - Osasco - SP - CEP 06020-190 - Fone: (011) 3654-0655 - C.N.P.J. 73.063.166/0004-73</font></td></tr>
</table>
</center></div>
<!-- ----------------------------- -->
<DIV style="page-break-after:always"></DIV>

<div align="center">
<center>
<!-- ----------------------------- -->
<table border="0" cellpadding="5" width="690" cellspacing="0" height="1000">
<!-- linha da aguia -->
<tr><td height=112><img src="../images/aguia.jpg" border="0" width="236" height="111" alt=""></td></tr>
<!-- corpo da declaracao -->
<tr><td height=100% valign=top style="text-align:justify">
<!-- ----------------------------- -->

<p style="text-align:center"><b>DECLARAÇÃO DE OPÇÃO DE CONTINUIDADE</b></p>

<p style="margin-bottom:2px;margin-top:5px">Tendo firmado <b>DECLARAÇÃO DE CIÊNCIA SOBRE OS DIREITOS DOS ARTs. 30  e  31 DA LEI DE PLANOS DE SAÚDE – LEI Nº 9656/98</b>,  declaro expressamente:  

<p style="margin-top:0px;">(&nbsp;&nbsp;&nbsp;) <b>opção pela não</b> continuidade da condição de beneficiário no Plano de Assistência à Saúde;
<p style="margin-top:0px;">(&nbsp;&nbsp;&nbsp;) <b>opção pela continuidade</b> da condição de beneficiário no Plano de Assistência à 
	Saúde juntamente com  meus beneficiários nele inscritos, a qual formalizarei através da assinatura do 
	“Termo de Adesão”, diretamente na Intermédica em um dos endereços abaixo (Núcleo de Atendimento ao Cliente – NAC)
	, em horário comercial.
	
<p style="margin-top:5px;margin-bottom:0px">
<table border="0" cellpadding="1" width="690" cellspacing="0" height="">
<tr><td><b>São Paulo</b>: Praça Dom José Gaspar, 134 - 12º  andar - República - SP – SP – CEP: 01047-010    </td><td>(11) 3155-2125
<tr><td><b>Baixada Santista</b>: Rua Heitor Moraes, 19 – 1º andar – Boqueirão – Santos – SP - CEP: 11045-570</td><td>(13) 3229-1523
<tr><td><b>Sorocaba</b>: Av. Armando Salles de Oliveira, 504 – Trujillo – Sorocaba - SP - CEP: 18060-370    </td><td>(15) 3212-9318
<tr><td><b>Jundiaí</b>: Rua Antonio Segre, 295 - Jardim Brasil – Jundiaí - SP - CEP: 13201-843              </td><td>(11) 4583-0400
<tr><td><b>Rio de Janeiro</b>: Rua Sorocaba, 654 – Botafogo - Rio de Janeiro – RJ - CEP: 22271-110          </td><td>(21) 3984-2945
<tr><td><b>Minas Gerais</b>: Av. Olegário Maciel, 1195 – B. Lourdes – Belo Horizonte – MG – CEP: 30180-111  </td><td>(31) 2121-9018
<tr><td><b>Brasília</b>: SCS Quadra 05 Bloco B loja 80 – Asa Sul - Brasília – DF - CEP: 70305-904           </td><td>(61) 3704-7320
<tr><td><b>Recife</b>: Rua Barão de Itamaracá,142 – Espinheiro – Recife – PE - CEP: 52020-070               </td><td>(81) 2121-1030
<tr><td><b>Campinas</b>: Rua Carolina Florense, 201 – Guanabara - Campinas – SP – CEP: 13073-225            </td><td>(19) 3741-5620
<tr><td><b>Salvador</b>: Rua Lucaia, 156 – Rio Vermelho – BA - CEP: 41940-660                               </td><td>(71) 2104-3666
</table>

<p style="margin-bottom:2px;margin-top:5px">Estou ciente de que caso não exista NAC na cidade onde resido ou nas proximidades, deverei entrar em contato com a Central de Atendimento 24h, de minha localidade, cujo telefone está expresso na minha carteira do convênio, a fim de receber orientações dos procedimentos necessários.

<p style="margin-bottom:2px;margin-top:5px"><u><b>Relação de Documentos que deverão ser apresentados</b></u>:
<br>- Via original desta Declaração de Opção de Continuidade assinada pelo responsável pela Área de Recursos Humanos da empresa e pelo ex-empregado ou aposentado;
<br>- 03 cópias dos últimos holerites, acompanhadas dos seus originais, ou documentos emitidos pela empresa que demonstrem os descontos referentes à contribuição ao plano de assistência à saúde;
<br>- Comprovante de residência em nome do titular;
<br>- Termo de Rescisão do Contrato de Trabalho (original e cópia);
<br>- Carteira de Trabalho (original) e cópias: frente e verso da página com a foto; da página do registro, e da         página seguinte;
<br>- No caso de Aposentado: apresentar a Carta de Concessão da Aposentadoria no INSS (original e cópia);
<br>- Declaração da ex-empregadora assinada pelo responsável pela Área de Recursos Humanos e pelo ex-empregado ou aposentado informando o tempo de contribuição na operadora atual de planos de assistência à saúde e o de cada operadora porventura anteriormente contratada sucessivamente, na(s) qual(is) o ex-empregado ou aposentado tenha pago, referente à taxa do Titular, parcial ou integralmente, as mensalidades referentes a plano com padrão de acomodação e rede referenciada,  superior àquele oferecido e pago integralmente pela empresa (upgrade). 

<p style="margin-bottom:2px;margin-top:5px">Estou ciente, declaro que me foi explicado e concordo que:
<br> a) os prazos de direito, vigência, manutenção do benefício e demais condições estão estabelecidas no <b>Termo de Adesão</b>;
<br> b) a não formalização de adesão junto a Intermédica no supracitado prazo de 30 (trinta) dias, tornará automaticamente nula minha opção pela manutenção da condição de beneficiário no Plano de Assistência à Saúde.

<p style="margin-bottom:2px;margin-top:5px">São Paulo, ______ de_____________ de <%=year(now)%>
<br><br><br>

<p style="margin-bottom:0px;margin-top:5px">___________________________________                               
<br>Assinatura do Beneficiário Titular

<!-- ----------------------------- -->
</td></tr>
<!-- linha intermediaria -->
<tr><td height="1"></td></tr>
<tr><td height="1"><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
<tr><td height="1"><font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000 - C.N.P.J. 73.063.166/0001-20</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP 06020-190 - Fone: (011) 3651-9999 - C.N.P.J. 73.063.166/0003-92</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 1743 - Osasco - SP - CEP 06020-190 - Fone: (011) 3654-0655 - C.N.P.J. 73.063.166/0004-73</font></td></tr>
</table>
</center></div>
<!-- ----------------------------- -->
<DIV style="page-break-after:always"></DIV>

<div align="center">
<center>
<!-- ----------------------------- -->
<table border="0" cellpadding="5" width="690" cellspacing="0" height="1000">
<!-- linha da aguia -->
<tr><td height=112><img src="../images/aguia.jpg" border="0" width="236" height="111" alt=""></td></tr>
<!-- corpo da declaracao -->
<tr><td height=100% valign=top style="text-align:justify">
<!-- ----------------------------- -->

<p style="text-align:center"><b>INFORMAÇÕES REFERENTES AO DESLIGAMENTO DO FUNCIONÁRIO</b></p>

<p>Este questionário, de acordo com a Resolução Normativa Nº 279, da ANS, deverá ser preenchido pela 
empresa na data de formalização da comunicação do Aviso Prévio ou da comunicação da Aposentadoria ao funcionário.

<p>Nome do Funcionário: <b><%=rs("nome")%></b>
<br>CPF/MF: <b><%=rs("cpf")%></b>

<p> I) O Beneficiário foi excluído por demissão ou exoneração sem justa causa ou aposentadoria?
<br> ( <%=Sdem%> ) Sim     
<br> ( <%=Ndem%> ) Não

<p> II) O Beneficiário demitido ou exonerado sem justa causa é um Beneficiário aposentado que continuava trabalhando na Contratante?
<br> ( <%=Sapos%> ) Sim     
<br> ( <%=Napos%> ) Não

<p> III) O Beneficiário contribuía para o pagamento do plano privado de assistência à saúde?
<br> ( <%=Scont%> ) Sim     
<br> ( <%=Ncont%> ) Não

<p> IV) Por quanto tempo o Beneficiário contribuiu para o pagamento do plano privado de assistência à saúde?
<br> <u>&nbsp;&nbsp;<%=cano%>&nbsp;&nbsp;</u> anos <u>&nbsp;&nbsp;<%=cmes%>&nbsp;&nbsp;</u> meses

<p> V) O ex-empregado optou pela sua manutenção como Beneficiário?
<br> ( <%=Ncont%> ) Sim     
<br> ( <%=Ncont%> ) Não

<br>
<p><b>Importante</b>:
<p>Estas informações referentes ao desligamento do funcionário, assim como a 2ª via da DECLARAÇÃO DE OPÇÃO DE 
CONTINUIDADE, deverão ficar sob a guarda e responsabilidade da empresa, que se compromete expressamente a 
enviá-la à Intermédica, no prazo de até 5 (cinco) dias úteis, contados da data da solicitação.
<br>Na hipótese de apresentação intempestiva do documento requerido ou no caso de sua não apresentação, 
a Empresa assumirá, de plano todos os prejuízos eventualmente suportados em decorrência desta ação ou omissão.

<p style="margin-bottom:2px;margin-top:5px">_____________________, ______ de __________________ de <%=year(now)%>
<br><br><br>

<p style="margin-bottom:0px;margin-top:5px">___________________________________                               
<br>Assinatura da Empresa sob carimbo

<!-- ----------------------------- -->
</td></tr>
<!-- linha intermediaria -->
<tr><td height="1"></td></tr>
<tr><td height="1"><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
<tr><td height="1"><font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000 - C.N.P.J. 73.063.166/0001-20</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP 06020-190 - Fone: (011) 3651-9999 - C.N.P.J. 73.063.166/0003-92</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 1743 - Osasco - SP - CEP 06020-190 - Fone: (011) 3654-0655 - C.N.P.J. 73.063.166/0004-73</font></td></tr>
</table>
</center></div>
<!-- ----------------------------- -->

<%
rs.close
set rs=nothing

elseif temp=2 then
%>
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
	<td class=campo>&nbsp;<a href="decl_copart.asp?codigo=<%=rs("chapa")%>"><%=rs("nome")%></a></td>
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
set rs3=nothing
conexao2.close
set conexao2=nothing
%>
</body>
</html>