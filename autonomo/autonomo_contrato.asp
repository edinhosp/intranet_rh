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
<title>Contrato de Presta��o de Servi�os</title>
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
<p class=titulo>Contrato de Presta��o de Servi�os para&nbsp;<%=titulo %>
<form method="POST" action="autonomo_contrato.asp" name="form">
<input type="hidden" name="id_autonomo1" value="<%=id_autonomo%>">
<table border="0" width="500" cellspacing="0"cellpadding="3">
<tr>
	<td class=titulo>C�digo</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo><input type="hidden" value="<%=ct_codigo%>" name="id_autonomo" size="8"><%=ct_codigo%></td>
	<td class=fundo><select size="1" name="nome_autonomo" onchange="nome1()">
	<option>Selecione um aut�nomo</option>
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
	<td class=titulo>Tipo Servi�o</td>
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
	<td class=titulo>T�rmino</td>
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

<tr><td align="center" style="border: 1px solid #000000"><b><font size="3">CONTRATO DE PRESTA��O DE SERVI�OS</font></b></td></tr>

<tr><td align="left">Instrumento particular de Contrato de Presta��o de Servi�os, que entre si celebram,
	<br><br><b>DE UM LADO:<br></td></tr>
<%
if rs("sexo")="M" then a1="" else a1="a"
if rs("sexo")="M" then a2="o" else a2="a"
if rs("sexo")="M" then a3="ao" else a3="a"
cpf=rs("cpf"):if cpf<>"" then cpf=left(cpf,3) & "." & mid(cpf,4,3) & "." & mid(cpf,7,3) & "-" & right(cpf,2)
%>
<tr><td><p align="justify"><b>FIEO - FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</b>, 
	Institui��o declarada de Utilidade P�blica pelo Decreto Federal n� 90.564 de 27/11/1984; pela Lei Estadual n� 1763 de 20/09/78 
	e pelo Decreto Municipal n� 2605 de 23/08/72, com endere�os na Rua Narciso Sturlini n� 883 e Av. Franz Voegelli n� 300 e 1743, 
	Osasco - SP, CNPJ n� 73.063.166/0001-20, mantenedora do <b>CENTRO UNIVERSIT�RIO FIEO - UNIFIEO</b>, celebra o presente CONTRATO
	DE PRESTA��O DE SERVI�OS E OUTRAS AVEN�AS, que ter� vig�ncia a partir da data de in�cio da presta��o de servi�os, de acordo com
	as condi��es a seguir especificadas.</td></tr>

<tr><td align="left"><b>DE OUTRO LADO:<br></td></tr>

<tr><td><p align="justify"><b><%=rs("nome_autonomo")%></b>, residente na <%=rs("rua")%>&nbsp;<%=rs("numero")%>&nbsp;
	<%=rs("complemento")%>&nbsp;<%=rs("bairro")%> - <%=rs("cidade")%>, <%=rs("nacionalidade")%>, <%=rs("estado_civil")%>, 
	portador<%=a1%> da C�dula de Identidade RG n� <%=rs("rg")%> e CPF n� <%=CPF%>, doravante denominad<%=a2%> CONTRATAD<%=ucase(a2)%>.
</tr>

<tr><td><p align="justify">As PARTES resolvem firmar o presente CONTRATO, compromentendo-se a respeitar fielmente as Cl�usulas 
	seguintes:</td>
</tr>

<tr><td align="center"><b>Cl�usula 1 - DO OBJETO</td></tr>
<tr><td><p align="justify">1.1 Este CONTRATO tem por objeto exclusivo a presta��o de servi�os como <%=request.form("tipo_prestacao")%>,
	em local a ser definido pela CONTRATANTE.<br>
	1.2 As atividades da presta��o de servi�os ora contratado ser� realizado de acordo com as necessidades dos trabalhos.</td></tr>
<%
inicio=request.form("inicio")
termino=request.form("termino")
inicio=day(inicio) & " de " & monthname(month(inicio)) & " de " & year(inicio)
termino=day(termino) & " de " & monthname(month(termino)) & " de " & year(termino)
valor=request.form("valor_hora")
%>
<tr><td align="center"><b>Cl�usula 2 - DO PRAZO</td></tr>
<tr><td><p align="justify">2.1 O presente CONTRATO tem prazo de <%=request.form("dias")%> dias, com in�cio em <%=inicio%> e 
	t�rmino em <%=termino%>.</td></tr>

<tr><td align="center"><b>Cl�usula 3 - DO PRE�O E DO PAGAMENTO</td></tr>
<tr><td><p align="justify">3.1 Pela presta��o dos servi�os, objeto do presente CONTRATO, a CONTRATANTE pagar� <%=a3%> 
	CONTRATAD<%=ucase(a2)%>, a quantia de R$ <%=formatnumber(valor,2)%> ( <%=extenso2(valor)%>) a hora trabalhada, ficando por 
	conta d<%=a2%> CONTRATAD<%=ucase(a2)%>, as despesas com encargos sociais e as despesas com alimenta��o e transporte.<br>
	3.2 Ao t�rmino do trabalho, a CONTRATANTE emitir� RPA - Recibo de Pagamento de Aut�nomo, pagando <%=a3%> CONTRATAD<%=ucase(a2)%>
	at� o d�cimo dia subseq�ente do t�rmino da presta��o dos servi�os.</td></tr>

<tr><td align="center"><b>Cl�usula 4 - DAS OBRIGA��ES DAS PARTES</td></tr>
<tr><td><p align="justify">4.1 A CONTRATANTE obriga-se a:<br>
	4.1.1 Pagar as contra presta��es <%=a3%> CONTRATAD<%=ucase(a2)%> pontualmente;<br>
	4.1.2 Permitir o acesso as depend�ncias d<%=a2%> CONTRATAD<%=ucase(a2)%> devidamente credenciad<%=a2%> pela CONTRATANTE, 
	e tamb�m colocar a disposi��o os materiais e equipamentos necess�rios visando o atendimento e a perfeita execu��o dos servi�os
	objeto deste contrato.<br>
	<br>
	4.2 <%=ucase(a2)%> CONTRATAD<%=ucase(a2)%> obriga-se a:<br>
	4.2.1 Cumprir fielmente o objeto do presente contrato, bem como, manter o mais completo sigilo sobre quaisquer dados, materiais, 
	pormenores, informa��es, documentos, especifica��es t�cnicas ou comerciais, inova��es ou aperfei�oamentos da CONTRATANTE, de que 
	venha a ter conhecimento, ou acesso, ou que venha a lhe ser confiado, em raz�o deste contrato, sejam eles de interesse da 
	CONTRATANTE, ou de terceiros, n�o podendo, sob qualquer pretexto, divulgar, revelar, reproduzir, utilizar ou deles dar conhecimento
	a terceiros e estranhos a esta contrata��o, sob as penas da lei.<br>
	
	4.2.2 Para o desenvolvimento dessas atividades <%=ucase(a2)%> CONTRATAD<%=ucase(a2)%> n�o estar� sujeito a controle de hor�rio e
	habitualidade no fornecimento e execu��o dos servi�os.<br>
	
	4.2.3 O inadimplemento desta cl�usula implicar� na reten��o do pagamento da remunera��o vincenda, por parte da CONTRATANTE.<br>
	4.2.4 Nenhuma obriga��o fiscal, previdenci�ria ou trabalhista tocante aos servi�os d<%=a2%> CONTRATAD<%=ucase(a2)%> ser�o de 
	responsabilidade da CONTRATANTE, ficando evidente que esta atividade � eventual, e portanto, n�o gerar� em hip�tese alguma 
	v�nculo de trabalho.
	</td></tr>

<tr><td height=10 align="right">..........</td></tr>
	
</table>
<DIV style="page-break-after:always"></DIV>
<!-- pagina 2 -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650" valign="top" width="970">

<tr><td align="center"><b>Cl�usula 5 - DA RESCIS�O</td></tr>
<tr><td><p align="justify">5.1 O inadimplemento, por qualquer das partes, de qualquer obriga��o oriunda do presente CONTRATO, 
	<i>inclusive nos casos de imper�cia, ofensa � �tica profissional, apropria��o ind�bita, quebra de sigilo, etc.</i> dara o direito
	� parte inocente de rescindir o presente CONTRATO, por justa causa, podendo ainda requerer indeniza��o por danos e lucros 
	cessantes incorridos. Nesta hip�tese, a parte inocente dar� conhecimento, no prazo de 48 (quarenta e oito) horas, da sua 
	inten��o.<br>
	5.2 Por fim, caso n�o haja interesse, a qualquer um das partes, em dar continuidade neste CONTRATO, a sua inten��o de rescis�o, 
	nenhuma indeniza��o ser� devida.</td></tr>

<tr><td align="center"><b>Cl�usula 6 - DISPOSI��ES GERAIS</td></tr>
<tr><td><p align="justify">6.1 <%=ucase(a2)%> CONTRATAD<%=ucase(a2)%> se responsabiliza integralmente pela boa execu��o dos servi�os 
	deste contrato, ficando obrigad<%=a2%> a corrigir os eventuais erros ou omiss�es verificadas, desde que tenha sido respons�vel por
	estas, ficando sujeit<%=(a2)%> ao ressarcimento a t�tulo de danos morais e materiais em rela��o ao CONTRATANTE e a terceiros.<br>
	6.2 O presente contrato n�o estabelece entre as partes qualquer v�nculo trabalhista ou societ�rio, nem convenciona qualquer 
	associa��o com personalidade jur�dica entre as partes contratantes, as quais continuam mantendo independ�ncia, sujeitando-se 
	exclusivamente ao pactuado neste contrato, cabendo, em raz�o disso, <%=ucase(a3)%> CONTRATAD<%=ucase(a2)%>, a responsabilidade 
	pela execu��o dos servi�os, principalmente quanto aos encargos advindos da legisla��o fiscal, tribut�ria, previdenci�ria, comercial e civil.<br>
	6.3 As cl�usulas do contrato constante deste instrumento sempre prevalecer�o sobre quaisquer acordos verbais ou escritos ajustatados 
	anteriormente � data de sua assinatura, sendo que a fixa��o de outras regras, que sirvam de norteamento ao desenvolvimento de seu objetivo,
	ser�o sempre feitas por escrito atrav�s de renova��es escritas e assinaturas pelas partes.<br>
	6.4 Nos casos de omiss�o, d�vidas ou lides oriundas deste contrato, a ele aplicar-se-a as regras do C�digo Civil Brasileiro, relativas 
	� presta��o de servi�os, e do C�digo de Defesa do Consumidor, elegendo as partes o Foro da Comarca de Osasco, com a expressa ren�ncia 
	de qualquer outro, por mais privilegiado que seja, para dirimi-las.<br>

<!--	6.5 Fica eleito o Foro da Comarca de Osasco para dirimir eventuais d�vidas e controv�rsias do presente 
	CONTRATO, renunciando as partes a qualquer outro, por mais privilegiado que seja.</td></tr>
-->
<tr><td height=10></td></tr>
	
<tr><td><p align="justify">Estando as PARTES assim justas e pactuadas, assina o presente em 2 (duas) vias de igual teor, na presen�a
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
			FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</td>
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