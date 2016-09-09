<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a42")="N" or session("a42")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Adendo ao Contrato de Trabalho</title>
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
	tipoform=0:	idnomeacao=request.form("chapa")
	sqlc="SELECT p.CHAPA, p.NOME, " & _
	"p.SEXO, p.RUA, p.NUMERO, p.COMPLEMENTO, p.BAIRRO, p.CIDADE, p.CEP, p.FUNCAO, " & _
	"p.CARTEIRATRAB, p.SERIECARTTRAB, p.codsecao, s.descricao " & _
	"FROM dc_professor AS p, corporerm.dbo.psecao as s " & _
	"WHERE p.CHAPA='" & idnomeacao & "' and p.codsecao=s.codigo "
	sqld=""
	sqle=" ORDER BY p.nome "
	sqlb=sqlc & sqld & sqle
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
		sqlh="select valor from corporerm.dbo.pvalfix where codigo='N8' and '" & dtaccess(request.form("mand_ini")) & "' between iniciovigencia and finalvigencia"
		rs2.Open sqlh, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then hora=rs2("valor") else hora=0
		rs2.close
	chapa      =request.form("chapa")      : nome      =rs("nome")
	atividades =request.form("atividades") : regime    =request.form("regime")
	inicio     =request.form("inicio")     : portaria  =request.form("portaria")

	rua        =rs("rua")                  : numero     =rs("numero")
	complemento=rs("complemento")       :	bairro     =rs("bairro")
	cidade     =rs("cidade")            :	cep        =rs("cep")
	ctps       =rs("carteiratrab")      :	serie      =rs("seriecarttrab")
	funcao     =rs("funcao")            :	tipo_adendo=request.form("tipo_adendo")
	sexo       =rs("sexo")
	secao      =rs("descricao")
	rs.close
else
	tipoform=1
		sqlh="select valor from corporerm.dbo.pvalfix where codigo='N8' and getdate() between iniciovigencia and finalvigencia"
		rs2.Open sqlh, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then hora=rs2("valor") else hora=0
		rs2.close
		chapa     ="" : nome=""
		atividades="Programa de Mestrado em Direitos Fundamentais" : regime="40"
		inicio    =formatdatetime(now,2) : portaria="Portaria nº "
end if


if tipoform<>0 then
%>
<p class=titulo>Adendo ao Contrato de Trabalho para&nbsp;<%=titulo %>
<form method="POST" action="contratort.asp" name="form">
<input type="hidden" name="id_indicado" value="<%=id_indicado%>">
<table border="0" width="400" cellspacing="0"cellpadding="3">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=chapa%>" name="chapa" onchange="chapa1()" size="8"></td>
	<td class=fundo><select size="1" name="nome" onchange="nome1()">
	<option>Selecione um professor</option>
&nbsp;
<%
sql2="select chapa, nome from dc_professor where codsituacao<>'D' order by nome"
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' and codsindicato='03' "
if tipoform=2 then sql2=sql2 & " and chapa='" & chapa & "'" else sql2=sql2 & " order by nome"
rs.Open sql2, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
if chapa=rs("chapa") then temp1="selected" else temp1=""
%>
	<option value="<%=rs("chapa")%>" <%=temp1%>><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
	</select></td>
</tr>
</table>

<table border="0" width="400" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo>&nbsp;</td>
	<td class=titulo>Atividades para</td>
	<td class=titulo>Regime</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="tipo_adendo" class=a>
		<option value="no">no</option>
		<option value="na">na</option>
		<option value="em">em</option>
		<option value="para">para</option>
		<option value="como">como</option>
		</select>
	</td>
	<td class=fundo><input type="text" value="<%=atividades%>" name="atividades" size="50"></td>
	<td class=fundo><b>RT </b><input type="text" value="<%=regime%>" name="regime" size="3"></td>
</tr>
</table>

<table border="0" width="400" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo>Inicio</td>
	<td class=titulo>Portaria de Nomeação</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=inicio%>" name="inicio" size="8"></td>
	<td class=fundo><input type="text" value="<%=portaria%>" name="portaria" size="50"></td>
</tr>
</table>


<table border="0" width="400" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo><input type="submit" value="Visualizar" class="button" name="B1"></td>
</tr>
</table>
</form>
<%
else ' tipoform=0
if sexo="F" then v1="a" else v1="o"
if sexo="F" then v2="a" else v2=""
if sexo="F" then v3="" else v3="o"
%>
<div align="right">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650" height="970">
<tr><td><img border="0" src="../images/aguia.jpg" width="236"></td> </tr>

<tr>
	<td><p align="center"><b><font size="3">ADENDO AO CONTRATO DE TRABALHO</font></b></p>
		<p align="center">&nbsp;</td>
</tr>

<tr>
	<td><p align="justify">Entre as partes, de um lado a <b>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</b>, 
	com sede a Av. Franz Voegeli, 300, Vila Yara, Osasco, CEP 06020-190, inscrita no CNPJ nº 73.063.166/0003-92, 
	denominada Contratante e de outro lado <%=v1%> Sr<%=v2%>. <b><%=nome%></b>, <%=funcao%>, residente e 
	domiciliad<%=v1%> à <%=rua%>&nbsp;<%=numero%>&nbsp;<%=complemento%> - <%=bairro%> - <%=cidade%> - 
	CEP <%=cep%>, portador<%=v2%> da CTPS nº<%=ctps%>/<%=serie%>, denominad<%=v1%> Professor<%=v2%>, 
	acordam o que se segue:</td>
</tr>

<tr>
	<td>
	<p align="justify">1. <%=ucase(v1)%> Professor<%=v2%> cumprirá atividades acadêmicas <%=tipo_adendo%>&nbsp;<%=atividades%>, em Regime
	de Trabalho - RT<%=regime%>, nelas estando incluídas as demais atividades, conforme cláusula 5.
	</td>
</tr>

<tr>
<%
if portaria="" or portaria="Portaria nº " then
	textoportaria="."
else
	textoportaria=", sendo " & v1 & " Professor" & v2 & " nomead" & v1 & " através da " & portaria & " da Reitoria."
end if
%>
	<td><p align="justify">2. As modalidades de Regime de Trabalho - RT<%=regime%>, tem início a partir de <%=inicio%>
	<%=textoportaria%></td>
</tr>

<tr>
	<td><p align="justify">3. No exercício de suas atividades está <%=(v1)%> Professor<%=v2%> sujeit<%=v1%> as normas 
	constantes do Regimento da Instituição de Ensino e do que prevê a legislação vigente.</td>
</tr>

<tr>
	<td><p align="justify">4. O presente termo aditivo poderá ser extinto a qualquer momento, não gerando ônus para as
	partes, quando ocorrer alterações na estrutura da Instituição ou de comum acordo entre as partes.</td>
</tr>

<tr>
	<td><p align="justify">5. Pelas atividades realizadas em RT<%=regime%>, <%=v1%> Professor<%=v2%> receberá pela
	contraprestação dos serviços, valores conforme tabela de remuneração específica para o curso. Incluem-se nos valores
	de Regime de Trabalho: horas-aulas em curso de graduação conforme disposto no § único; desenvolvimento de projeto de pesquisa aprovado no âmbito do
	curso de Mestrado/Especialização; orientação de até 10 (dez) mestrandos/pós-graduandos; atribuição de disciplina pela
	Coordenação do Curso de Mestrado/Especialização, devendo o professor assumir <%=regime%> horas semanais, as quais
	serão registradas através de ponto eletrônico, podendo cumprir até <%=regime/2%> horas semanais fora da Instituição,
	apresentando neste caso, relatório das atividades.
</tr>
<tr>
	<td>§ único. <%=ucase(v1)%> Professor<%=v2%> deverá atender ao curso de graduação quando convocado.</td>
</tr>

<!--
<tr>
	<td><p align="justify">6. Ficam sujeitas ao controle do ponto eletrônico todas as horas cumpridas, sejam no Campus
	Vila Yara ou no Campus Narciso.</td>
</tr>
-->
<tr>
	<td><p align="justify">6. O cumprimento da carga horária deverá ser comprovado mensalmente através de relatório
	das atividades desenvolvidas pelo PROFESSOR validado pelo coordenador do curso e entregue até o dia 10 do mês 
	subseqüente à Pró-Reitoria Acadêmica.</td>
</tr>

<tr>
	<td>E, por assim estarem de acordo, firmam o presente em 2 (duas) vias, uma das quais é entregue a<%=v3%> Professor<%=v2%>, 
	na presença de 2 (duas) testemunhas abaixo qualificadas.</td>
</tr>

<tr>
	<td>&nbsp;</td>
</tr>

<tr>
	<td>
<%
'if contrato="" then contrato=formatdatetime(now(),2)
if contrato="" then contrato=formatdatetime(inicio,2)
dia=day(contrato)
mes=monthname(month(contrato))
ano=year(contrato)
%>
		<p align="left">Osasco,&nbsp;<%=dia & " de " & mes & " de " & ano %></td>
</tr>

<tr><td>&nbsp;</td></tr>

<tr>
	<td>
		<table border="0" width="100%" cellspacing="0">
		<tr>
			<td width="50%">&nbsp;
				<p>_______________________________________<br>
				FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td>
			<td width="50%">&nbsp;
				<p>_______________________________________<br>
				<b><%=nome%></b></td>
		</tr>
		</table>
		<p>&nbsp;</td>
</tr>

<tr><td>Testemunhas:</td></tr>

<tr>
	<td>
		<table border="0" width="100%" cellspacing="0">
		<tr>
			<td width="50%">_______________________________________<br>
			Nome:<br>
			R.G.:</td>
			<td width="50%">_______________________________________<br>
			Nome:<br>
			R.G.:</td>
		</tr>
		</table>
	</td>
</tr>

<tr><td> &nbsp;<p align="right"><font size=1><%=secao%></font></p></td></tr>

<tr><td><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b></td> </tr>
<tr><td><font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000</font></td></tr>
<tr><td><font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP 06020-190 - Fone: (011) 3651-9999</font></td></tr>
<tr><td><font face="Arial Narrow">Caixa Postal - ACF - Jd. Ipê - nº 2530 - Osasco - SP - CEP 06053-990</font></td></tr>
</table>
</div>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<%
end if 
%>
</body>
</html>
<%
if tipoform=0 and id_indicado<>"" then
	sqlz="UPDATE n_indicacoes SET CONTRATO = #" & dtaccess(contrato) & "# "
	sqlz=sqlz & " WHERE id_indicado=" & id_indicado
	response.write sqlz	
	conexao.execute sqlz
end if

set rs=nothing
conexao.close
set conexao=nothing
%>