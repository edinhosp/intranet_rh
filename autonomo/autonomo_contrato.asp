<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a52")="N" or session("a52")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Contrato de Prestação de Serviços</title>
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
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
function mand_ini1(muda) {
	temp=form.mand_ini.value;
	inicio=new Date(temp.substr(6),temp.substr(3,2)-1,temp.substr(0,2));
	hoje=new Date();
	hoje.setDate(1);hoje.toLocaleString();
	fpgini="0" + hoje.getDate() + "/" + ((hoje.getMonth()+1)<10?"0":"") + (hoje.getMonth()+1) + "/" + hoje.getFullYear();
	//form.fpg_ini.value=fpgini;
	if (muda==1) { temp2=form.fpg_ini.value; hoje=new Date(temp2.substr(6),temp2.substr(3,2)-1,1); }
	dinicio=montharray[inicio.getMonth()]+" "+inicio.getDate()+", "+inicio.getFullYear()
	dmesfp=montharray[hoje.getMonth()]+" "+hoje.getDate()+", "+hoje.getFullYear()
	dias=(Math.round((Date.parse(dmesfp)-Date.parse(dinicio))/(24*60*60*1000))*1)
	semanas=Math.round(dias/7)
	dmesini=montharray[inicio.getMonth()]+" 1, "+inicio.getFullYear()
	if (dmesfp!=dmesini) {
		if (muda==0) { document.form.fpg_ini.value=fpgini }
		horas=document.form.ch.value
		document.form.complemento.value=horas*semanas
	} else {
		document.form.complemento.value=0
		if (muda==0) { document.form.fpg_ini.value=temp }
	}		
}
--></script>
<script language="VBScript">
	Sub termino_onChange
		inicio=document.form.inicio.value
		termino=document.form.termino.value
		dias=cdate(termino)-cdate(inicio)+1
		document.form.dias.value=dias
	End Sub
	Sub inicio_onChange
		inicio=document.form.inicio.value
		dias=document.form.dias.value
		document.form.termino.value=dateadd("d",cint(dias)-1,formatdatetime(document.form.inicio.value,2))
		'dias=cdate(termino)-cdate(inicio)+1
		'document.form.dias.value=dias
	End Sub
	Sub dias_onChange
		dias=document.form.dias.value
		document.form.termino.value=dateadd("d",cint(dias)-1,formatdatetime(document.form.inicio.value,2))
	End Sub
</script>

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

if request("codigo")<>"" then
	tipoform=2:	idautonomo=request("codigo")
	sqlb="SELECT id_autonomo, nome_autonomo, dtnascimento, sexo, tipo_prestacao, cpf, nit, rg, nacionalidade, estado_civil " & _
	"FROM autonomo where id_autonomo=" & idautonomo & " ORDER BY nome_autonomo "
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
	ct_codigo     =rs("id_autonomo")	: ct_nome       =rs("nome_autonomo")
	ct_nacional   =rs("nacionalidade")	: ct_estadocivil=rs("estado_civil")
	ct_sexo       =rs("sexo")			: ct_servico    =rs("tipo_prestacao")
	rs.close
elseif request.form<>"" then
	tipoform=0:	id_autonomo=request.form("id_autonomo")
	if request.form("nacionalidade")<>request.form("nacionalidade1") then
		sql="update autonomo set nacionalidade='" & request.form("nacionalidade") & "' where id_autonomo=" & id_autonomo
		conexao.execute sql
	end if
	if request.form("estado_civil")<>request.form("estado_civil1") then
		sql="update autonomo set estado_civil='" & request.form("estado_civil") & "' where id_autonomo=" & id_autonomo
		conexao.execute sql
	end if
	sqlc="SELECT a.id_autonomo, a.nome_autonomo, a.dtnascimento, a.tipo_prestacao, a.sexo, " & _
	"a.cpf, a.rg, a.rua, a.numero, a.complemento, a.bairro, a.cidade, a.cep, a.nacionalidade, a.estado_civil " & _
	"FROM autonomo a " & _
	"WHERE a.id_autonomo=" & id_autonomo & " "
	sqld=""
	sqle=" ORDER BY a.nome_autonomo "
	sqlb=sqlc & sqld & sqle
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
	ct_codigo     =rs("id_autonomo")	: ct_nome       =rs("nome_autonomo")
	ct_nacional   =rs("nacionalidade")	: ct_estadocivil=rs("estado_civil")
	ct_sexo       =rs("sexo")			: ct_servico    =rs("tipo_prestacao")
	'rs.close
else
	tipoform=1
	ct_codigo     =""	: ct_nome       =""
	ct_nacional   =""	: ct_estadocivil=""
	ct_sexo       =""	: ct_servico    =""
end if

if tipoform<>0 then
%>
<p class=titulo>Contrato de Prestação de Serviços para&nbsp;<%=titulo %>
<form method="POST" action="autonomo_contrato.asp" name="form">
<input type="hidden" name="id_autonomo1" value="<%=id_autonomo%>">
<table border="0" width="500" cellspacing="0"cellpadding="3">
<tr>
	<td class=titulo>Código</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo><input type="hidden" value="<%=ct_codigo%>" name="id_autonomo" size="8"><%=ct_codigo%></td>
	<td class=fundo><select size="1" name="nome_autonomo" onchange="nome1()">
	<option>Selecione um autônomo</option>
&nbsp;
<%
sql2="select chapa, nome from dc_professor where codsituacao<>'D' order by nome"
sql2="select id_autonomo as chapa, nome_autonomo as nome from autonomo where id_autonomo>0 "
if tipoform=2 then sql2=sql2 & " and id_autonomo=" & ct_codigo & "" else sql2=sql2 & " order by nome_autonomo"
rs.Open sql2, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
if ct_codigo=rs("chapa") then temp1="selected" else temp1=""
%>
	<option value="<%=rs("chapa")%>" <%=temp1%>><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
	</select></td>
</tr>
</table>

<table border="0" width="500" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo>Nacionalidade</td>
	<td class=titulo>Estado Civil</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=ct_nacional%>" name="nacionalidade" size="25"></td>
	<td class=fundo><input type="text" value="<%=ct_estadocivil%>" name="estado_civil" size="15"></td>
</tr>
</table>
<input type="hidden" value="<%=ct_nacional%>" name="nacionalidade1" size="25"></td>
<input type="hidden" value="<%=ct_estadocivil%>" name="estado_civil1" size="15"></td>

<table border="0" width="500" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo>Tipo Serviço</td>
	<td class=titulo>Valor hora R$</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=ct_servico%>" name="tipo_prestacao" size="50"></td>
	<td class=fundo><input type="text" value="<%=ct_valor%>" name="valor_hora" size="8"></td>
</tr>
</table>

<table border="0" width="500" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo>Prazo em Dias</td>
	<td class=titulo>Inicio</td>
	<td class=titulo>Término</td>
	<td class=titulo>Assinatura do Contrato</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=request.form("dias")%>" name="dias" size="4"></td>
	<td class=fundo><input type="text" value="<%=formatdatetime(now(),2)%>" name="inicio" size="9"></td>
	<td class=fundo><input type="text" value="" name="termino" size="9"></td>
	<td class=fundo><input type="text" value="<%=formatdatetime(now(),2)%>" name="assinatura" size="9"></td>
</tr>
</table>

<table border="0" width="500" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo><input type="submit" value="Visualizar" class="button" name="B1"></td>
</tr>
</table>
</form>
<%
else ' tipoform=0
if ct_sexo="F" then v1="a" else v1="o"
if ct_sexo="F" then v2="a" else v2=""
if ct_sexo="F" then v3="" else v3="o"
%>
<div align="right">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650" height="970" valign=top>
<!--<tr><td><img border="0" src="../images/aguia.jpg" width="236"></td> </tr>-->

<tr><td align="center" style="border: 1px solid #000000"><b><font size="3">CONTRATO DE PRESTAÇÃO DE SERVIÇOS</font></b></td></tr>

<tr><td align="left">Instrumento particular de Contrato de Prestação de Serviços, que entre si celebram,
	<br><br><b>DE UM LADO:<br></td></tr>
<%
if rs("sexo")="M" then a1="" else a1="a"
if rs("sexo")="M" then a2="o" else a2="a"
if rs("sexo")="M" then a3="ao" else a3="a"
cpf=rs("cpf"):if cpf<>"" then cpf=left(cpf,3) & "." & mid(cpf,4,3) & "." & mid(cpf,7,3) & "-" & right(cpf,2)
%>
<tr><td><p align="justify"><b>FIEO - FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</b>, 
	Instituição declarada de Utilidade Pública pelo Decreto Federal nº 90.564 de 27/11/1984; pela Lei Estadual nº 1763 de 20/09/78 
	e pelo Decreto Municipal nº 2605 de 23/08/72, com endereços na Rua Narciso Sturlini nº 883 e Av. Franz Voegelli nº 300 e 1743, 
	Osasco - SP, CNPJ nº 73.063.166/0001-20, mantenedora do <b>CENTRO UNIVERSITÁRIO FIEO - UNIFIEO</b>, celebra o presente CONTRATO
	DE PRESTAÇÃO DE SERVIÇOS E OUTRAS AVENÇAS, que terá vigência a partir da data de início da prestação de serviços, de acordo com
	as condições a seguir especificadas.</td></tr>

<tr><td align="left"><b>DE OUTRO LADO:<br></td></tr>

<tr><td><p align="justify"><b><%=rs("nome_autonomo")%></b>, residente na <%=rs("rua")%>&nbsp;<%=rs("numero")%>&nbsp;
	<%=rs("complemento")%>&nbsp;<%=rs("bairro")%> - <%=rs("cidade")%>, <%=rs("nacionalidade")%>, <%=rs("estado_civil")%>, 
	portador<%=a1%> da Cédula de Identidade RG nº <%=rs("rg")%> e CPF nº <%=CPF%>, doravante denominad<%=a2%> CONTRATAD<%=ucase(a2)%>.
</tr>

<tr><td><p align="justify">As PARTES resolvem firmar o presente CONTRATO, compromentendo-se a respeitar fielmente as Cláusulas 
	seguintes:</td>
</tr>

<tr><td align="center"><b>Cláusula 1 - DO OBJETO</td></tr>
<tr><td><p align="justify">1.1 Este CONTRATO tem por objeto exclusivo a prestação de serviços como <%=request.form("tipo_prestacao")%>,
	em local a ser definido pela CONTRATANTE.<br>
	1.2 As atividades da prestação de serviços ora contratado será realizado de acordo com as necessidades dos trabalhos.</td></tr>
<%
inicio=request.form("inicio")
termino=request.form("termino")
inicio=day(inicio) & " de " & monthname(month(inicio)) & " de " & year(inicio)
termino=day(termino) & " de " & monthname(month(termino)) & " de " & year(termino)
valor=request.form("valor_hora")
%>
<tr><td align="center"><b>Cláusula 2 - DO PRAZO</td></tr>
<tr><td><p align="justify">2.1 O presente CONTRATO tem prazo de <%=request.form("dias")%> dias, com início em <%=inicio%> e 
	término em <%=termino%>.</td></tr>

<tr><td align="center"><b>Cláusula 3 - DO PREÇO E DO PAGAMENTO</td></tr>
<tr><td><p align="justify">3.1 Pela prestação dos serviços, objeto do presente CONTRATO, a CONTRATANTE pagará <%=a3%> 
	CONTRATAD<%=ucase(a2)%>, a quantia de R$ <%=formatnumber(valor,2)%> ( <%=extenso2(valor)%>) a hora trabalhada, ficando por 
	conta d<%=a2%> CONTRATAD<%=ucase(a2)%>, as despesas com encargos sociais e as despesas com alimentação e transporte.<br>
	3.2 Ao término do trabalho, a CONTRATANTE emitirá RPA - Recibo de Pagamento de Autônomo, pagando <%=a3%> CONTRATAD<%=ucase(a2)%>
	até o décimo dia subseqüente do término da prestação dos serviços.</td></tr>

<tr><td align="center"><b>Cláusula 4 - DAS OBRIGAÇÕES DAS PARTES</td></tr>
<tr><td><p align="justify">4.1 A CONTRATANTE obriga-se a:<br>
	4.1.1 Pagar as contra prestações <%=a3%> CONTRATAD<%=ucase(a2)%> pontualmente;<br>
	4.1.2 Permitir o acesso as dependências d<%=a2%> CONTRATAD<%=ucase(a2)%> devidamente credenciad<%=a2%> pela CONTRATANTE, 
	e também colocar a disposição os materiais e equipamentos necessários visando o atendimento e a perfeita execução dos serviços
	objeto deste contrato.<br>
	<br>
	4.2 <%=ucase(a2)%> CONTRATAD<%=ucase(a2)%> obriga-se a:<br>
	4.2.1 Cumprir fielmente o objeto do presente contrato, bem como, manter o mais completo sigilo sobre quaisquer dados, materiais, 
	pormenores, informações, documentos, especificações técnicas ou comerciais, inovações ou aperfeiçoamentos da CONTRATANTE, de que 
	venha a ter conhecimento, ou acesso, ou que venha a lhe ser confiado, em razão deste contrato, sejam eles de interesse da 
	CONTRATANTE, ou de terceiros, não podendo, sob qualquer pretexto, divulgar, revelar, reproduzir, utilizar ou deles dar conhecimento
	a terceiros e estranhos a esta contratação, sob as penas da lei.<br>
	
	4.2.2 Para o desenvolvimento dessas atividades <%=ucase(a2)%> CONTRATAD<%=ucase(a2)%> não estará sujeito a controle de horário e
	habitualidade no fornecimento e execução dos serviços.<br>
	
	4.2.3 O inadimplemento desta cláusula implicará na retenção do pagamento da remuneração vincenda, por parte da CONTRATANTE.<br>
	4.2.4 Nenhuma obrigação fiscal, previdenciária ou trabalhista tocante aos serviços d<%=a2%> CONTRATAD<%=ucase(a2)%> serão de 
	responsabilidade da CONTRATANTE, ficando evidente que esta atividade é eventual, e portanto, não gerará em hipótese alguma 
	vínculo de trabalho.
	</td></tr>

<tr><td height=10 align="right">..........</td></tr>
	
</table>
<DIV style="page-break-after:always"></DIV>
<!-- pagina 2 -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650" valign="top" width="970">

<tr><td align="center"><b>Cláusula 5 - DA RESCISÃO</td></tr>
<tr><td><p align="justify">5.1 O inadimplemento, por qualquer das partes, de qualquer obrigação oriunda do presente CONTRATO, 
	<i>inclusive nos casos de imperícia, ofensa à ética profissional, apropriação indébita, quebra de sigilo, etc.</i> dara o direito
	à parte inocente de rescindir o presente CONTRATO, por justa causa, podendo ainda requerer indenização por danos e lucros 
	cessantes incorridos. Nesta hipótese, a parte inocente dará conhecimento, no prazo de 48 (quarenta e oito) horas, da sua 
	intenção.<br>
	5.2 Por fim, caso não haja interesse, a qualquer um das partes, em dar continuidade neste CONTRATO, a sua intenção de rescisão, 
	nenhuma indenização será devida.</td></tr>

<tr><td align="center"><b>Cláusula 6 - DISPOSIÇÕES GERAIS</td></tr>
<tr><td><p align="justify">6.1 <%=ucase(a2)%> CONTRATAD<%=ucase(a2)%> se responsabiliza integralmente pela boa execução dos serviços 
	deste contrato, ficando obrigad<%=a2%> a corrigir os eventuais erros ou omissões verificadas, desde que tenha sido responsável por
	estas, ficando sujeit<%=(a2)%> ao ressarcimento a título de danos morais e materiais em relação ao CONTRATANTE e a terceiros.<br>
	6.2 O presente contrato não estabelece entre as partes qualquer vínculo trabalhista ou societário, nem convenciona qualquer 
	associação com personalidade jurídica entre as partes contratantes, as quais continuam mantendo independência, sujeitando-se 
	exclusivamente ao pactuado neste contrato, cabendo, em razão disso, <%=ucase(a3)%> CONTRATAD<%=ucase(a2)%>, a responsabilidade 
	pela execução dos serviços, principalmente quanto aos encargos advindos da legislação fiscal, tributária, previdenciária, comercial e civil.<br>
	6.3 As cláusulas do contrato constante deste instrumento sempre prevalecerão sobre quaisquer acordos verbais ou escritos ajustatados 
	anteriormente à data de sua assinatura, sendo que a fixação de outras regras, que sirvam de norteamento ao desenvolvimento de seu objetivo,
	serão sempre feitas por escrito através de renovações escritas e assinaturas pelas partes.<br>
	6.4 Nos casos de omissão, dúvidas ou lides oriundas deste contrato, a ele aplicar-se-a as regras do Código Civil Brasileiro, relativas 
	à prestação de serviços, e do Código de Defesa do Consumidor, elegendo as partes o Foro da Comarca de Osasco, com a expressa renúncia 
	de qualquer outro, por mais privilegiado que seja, para dirimi-las.<br>

<!--	6.5 Fica eleito o Foro da Comarca de Osasco para dirimir eventuais dúvidas e controvérsias do presente 
	CONTRATO, renunciando as partes a qualquer outro, por mais privilegiado que seja.</td></tr>
-->
<tr><td height=10></td></tr>
	
<tr><td><p align="justify">Estando as PARTES assim justas e pactuadas, assina o presente em 2 (duas) vias de igual teor, na presença
	de 2 (duas) testemunhas.</td></tr>

<tr><td height=15></td></tr>

<%
ct_contrato=request.form("assinatura")
dia=day(ct_contrato)
mes=monthname(month(ct_contrato))
ano=year(ct_contrato)
%>
<tr><td align="right">Osasco,&nbsp;<%=dia & " de " & mes & " de " & ano %></td></tr>

<tr><td height=10></td></tr>

<tr><td>
	<table border="0" width="100%" cellspacing="0">
	<tr>
		<td width="50%">&nbsp;
			<p>_______________________________________<br>
			CONTRATANTE<br>
			FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td>
		<td width="50%">&nbsp;
			<p>_______________________________________<br>
			CONTRATAD<%=ucase(a2)%><br>
			<b><%=ct_nome%></b></td>
	</tr>
	</table>
</td></tr>

<tr><td>Testemunhas:</td></tr>

<tr>
	<td>
		<table border="0" width="100%" cellspacing="0">
		<tr>
			<td width="50%">&nbsp;<p>_______________________________________<br>
			Nome:<br>
			R.G.:</td>
			<td width="50%">&nbsp;<p>_______________________________________<br>
			Nome:<br>
			R.G.:</td>
		</tr>
		</table>
	</td>
</tr>

</table>
</div>
<%
end if 
%>
</body>
</html>
<%
if tipoform=0 and id_indicado<>"" then
	sqlz="UPDATE n_indicacoes SET CONTRATO = #" & dtaccess(ct_contrato) & "# "
	sqlz=sqlz & " WHERE id_indicado=" & id_indicado
	'response.write sqlz	
	conexao.execute sqlz
end if

set rs=nothing
conexao.close
set conexao=nothing
%>