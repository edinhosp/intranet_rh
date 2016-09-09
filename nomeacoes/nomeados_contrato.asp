<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a29")="N" or session("a29")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
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
set conexao2=server.createobject ("ADODB.Connection")
conexao2.Open application("consql")
	
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao2

if request("codigo")<>"" then
	tipoform=2:	idnomeacao=request("codigo")
	sqlc="SELECT i.id_nomeacao, n.NOMEACAO, i.PORTARIA, i.id_indicado, i.CHAPA, " & _
	"i.NOME, i.CARGO, i.MAND_INI, i.MAND_FIM, i.alunos, i.CH, i.OBS, i.CONTRATO, " & _
	"p.SEXO, p.RUA, p.NUMERO, p.COMPLEMENTO, p.BAIRRO, p.CIDADE, p.CEP, p.FUNCAO, " & _
	"p.CARTEIRATRAB, p.SERIECARTTRAB, faixavalor=case when tab_grade=1 then 1 when tab_grade=2 then 2 else 3 end " & _
	"FROM n_indicacoes AS i, n_nomeacoes AS n, dc_professor AS p " & _
	"WHERE i.id_nomeacao = n.id_nomeacao AND i.CHAPA = p.CHAPA collate database_default "
	sqld=" and i.id_indicado=" & idnomeacao & ""
	sqle=" ORDER BY n.nomeacao "
	sqlb=sqlc & sqld & sqle
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
		session("id_nomeacao")=rs("id_nomeacao")
		select case rs("id_nomeacao")
			case 12 'coord grad
				tipo="CGR"
			case 16 'coord pos
				tipo="CPG"
			case 80 'paga
				tipo="E103"
			case else
				select case rs("faixavalor")
					case 5
						tipo="NOM"
					case 4
						tipo="NOM"
					case 3
						tipo="NOM"
					case 2
						tipo="A1"
					case 1
						tipo="E1"
				end select
		end select
		if rs("mand_ini")<now() then datac=rs("mand_ini") else datac=now()
		sqlh="select valor from pvalfix where codigo='" & tipo & "' and '" & dtaccess(datac) & "' between iniciovigencia and finalvigencia"
		rs2.Open sqlh, ,adOpenStatic, adLockReadOnly
RESPONSE.WRITE SQLH
		if rs2.recordcount>0 then hora=rs2("valor") else hora=0
		rs2.close
	'if hora=0 then hora=cdbl(request.form("ct_valor"))
	if rs("cargo")="" then cargocurso="" else cargocurso=" - " & rs("cargo")
	ct_chapa      =rs("chapa")         :	ct_nome       =rs("nome")
	ct_rua        =rs("rua")           :	ct_numero     =rs("numero")
	ct_complemento=rs("complemento")   :	ct_bairro     =rs("bairro")
	ct_cidade     =rs("cidade")        :	ct_cep        =rs("cep")
	ct_ctps       =rs("carteiratrab")  :	ct_serie      =rs("seriecarttrab")
	ct_funcao     =rs("funcao")        :	ct_nomeacao   =rs("nomeacao") & cargocurso
	ct_ch         =rs("ch")            :	ct_mand_ini   =rs("mand_ini")
	ct_mand_fim   =rs("mand_fim")      :	ct_valor      =hora
	ct_portaria   =rs("portaria")      :	id_indicado   =rs("id_indicado")
	ct_sexo       =rs("sexo")          :	ct_contrato   =rs("contrato")
	rs.close
elseif request.form<>"" then
	tipoform=0:	idnomeacao=request.form("chapa")
	sqlc="SELECT p.CHAPA, p.NOME, faixavalor=case when tab_grade=1 then 1 when tab_grade=2 then 2 else 3 end, " & _
	"p.SEXO, p.RUA, p.NUMERO, p.COMPLEMENTO, p.BAIRRO, p.CIDADE, p.CEP, p.FUNCAO, " & _
	"p.CARTEIRATRAB, p.SERIECARTTRAB, p.codsecao, s.descricao " & _
	"FROM dc_professor p, corporerm.dbo.psecao s " & _
	"WHERE s.codigo=p.codsecao and p.CHAPA='" & idnomeacao & "' "
	sqld=""
	sqle=" ORDER BY p.nome "
	sqlb=sqlc & sqld & sqle
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
		select case session("id_nomeacao")
			case 12 'coord grad
				tipo="CGR"
			case 16 'coord pos
				tipo="CPG"
			case 80 'paga
				tipo="E103"
			case else
				select case rs("faixavalor")
					case 5
						tipo="NOM"
					case 4
						tipo="NOM"
					case 3
						tipo="NOM"
					case 2
						tipo="A1"
					case 1
						tipo="E1"
				end select
		end select
		if request.form("mand_ini")<now() then datac=request.form("mand_ini") else datac=now()
		sqlh="select valor from corporerm.dbo.pvalfix where codigo='" & tipo & "' and '" & dtaccess(datac) & "' between iniciovigencia and finalvigencia"
		rs2.Open sqlh, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then hora=rs2("valor") else hora=0
		rs2.close
	'if hora=0 and session("id_nomeacao")=80 then hora=16.42
	ct_chapa      =request.form("chapa")   :	ct_nome       =rs("nome")
	ct_rua        =rs("rua")               :	ct_numero     =rs("numero")
	ct_complemento=rs("complemento")       :	ct_bairro     =rs("bairro")
	ct_cidade     =rs("cidade")            :	ct_cep        =rs("cep")
	ct_ctps       =rs("carteiratrab")      :	ct_serie      =rs("seriecarttrab")
	ct_funcao     =rs("funcao")            :	ct_nomeacao   =request.form("nomeacao")
	ct_ch         =request.form("ch")      :	ct_mand_ini   =request.form("mand_ini")
	ct_mand_fim   =request.form("mand_fim"):	ct_valor      =hora
	ct_portaria   =request.form("portaria"):	ct_sexo       =rs("sexo")
	ct_contrato   =request.form("data_assinatura"):	ct_secao      =rs("descricao")
	id_indicado   =request.form("id_indicado") : horas=request.form("horas")
	rs.close
else
	tipoform=1
		sqlh="select valor from corporerm.dbo.pvalfix where codigo='NOM' and getdate() between iniciovigencia and finalvigencia"
		rs2.Open sqlh, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then hora=rs2("valor") else hora=0
		rs2.close
		'if hora=0 then hora=cdbl(request.form("ct_valor")+0)
	ct_chapa      =""   :	ct_nome       =""
	ct_rua        =""   :	ct_numero     =""
	ct_complemento=""   :	ct_bairro     =""
	ct_cidade     =""   :	ct_cep        =""
	ct_ctps       =""   :	ct_serie      =""
	ct_funcao     =""   :	ct_nomeacao   =""
	ct_ch         =""   :	ct_mand_ini   =""
	ct_mand_fim   =""   :	ct_valor      =hora
	ct_portaria   =""   :	id_indicado   =""
end if
'ct_valor=request.form("valor_hora")
if request.form("valor_hora")<>"" then ct_valor=request.form("valor_hora") else ct_valor=ct_valor

if tipoform<>0 then
%>
<p class=titulo>Adendo ao Contrato de Trabalho para&nbsp;<%=titulo %>
<form method="POST" action="nomeados_contrato.asp" name="form">
<input type="hidden" name="id_indicado" value="<%=id_indicado%>">
<input type="hidden" name="id_nomeacao" value="<%=session("id_nomeacao")%>">
<table border="0" width="500" cellspacing="0"cellpadding="3">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=ct_chapa%>" name="chapa" onchange="chapa1()" size="8"></td>
	<td class=fundo><select size="1" name="nome" onchange="nome1()">
	<option>Selecione um professor</option>
&nbsp;
<%
sql2="select chapa, nome from dc_professor where codsituacao<>'D' order by nome"
sql2="select chapa, nome from dc_professor where codsituacao<>'D' "
sql2="select chapa, nome from dc_professor where codsituacao<>'' "
if tipoform=2 then sql2=sql2 & " and chapa='" & ct_chapa & "'" else sql2=sql2 & " order by nome"
rs.Open sql2, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
if ct_chapa=rs("chapa") then temp1="selected" else temp1=""
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
	<td class=titulo>Tipo Adendo</td>
	<td class=titulo>Nomeação</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="tipo_adendo" class=a>
		<option value="para">para</option>
		<option value="como">como</option>
		<option value="em">em</option>
		<option value="na">na</option>
		<option value="no">no</option>
		</select>
	</td>
	<td class=fundo><input type="text" value="<%=ct_nomeacao%>" name="nomeacao" size="50"></td>
</tr>
</table>

<table border="0" width="500" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo>Horas semanais</td>
	<td class=titulo>Valor hora R$</td>
	<td class=titulo>Portaria de Nomeação</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=ct_ch%>" name="ch" size="5"></td>
	<td class=fundo><input type="text" value="<%=ct_valor%>" name="valor_hora" size="8"></td>
	<td class=fundo><input type="text" value="<%=ct_portaria%>" name="portaria" size="50"></td>
</tr>
</table>

<table border="0" width="500" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo>Inicio</td>
	<td class=titulo>Término</td>
	<td class=titulo>Horas</td>
	<td class=titulo>Data de Assinatura</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=ct_mand_ini%>" name="mand_ini" size="14"></td>
	<td class=fundo><input type="text" value="<%=ct_mand_fim%>" name="mand_fim" size="14"></td>
	<td class=fundo><input type="text" value="<%=horas%>" name="horas" size="3"></td>
	<td class=fundo><input type="text" value="<%if ct_contrato="" or isnull(ct_contrato) then response.write ct_mand_ini else response.write ct_contrato%>" name="data_assinatura" size="14"></td>
</tr>
</table>

<table border="0" width="500" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo><input type="submit" value="Visualizar" class="button" name="B1"></td>
</tr>
</table>
</form>
<p style="font-family:'courier new'">

<%
sql3="select codigo, valor from corporerm.dbo.pvalfix where getdate() between iniciovigencia and finalvigencia and codigo in ('nom','cpg','cgr','e103')"
rs.Open sql3, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
texto=""
if rs("codigo")="NOM" then texto="Nomeações (Estágio, TCC etc)."
if rs("codigo")="CPG" then texto="Coord. Pós-Graduação........."
if rs("codigo")="CGR" then texto="Coord. Graduação............."
if rs("codigo")="E103" then texto="PAGA-Prog.Ativ.Gerais Acad..."
%>
Valor para (<%=texto%>) - R$ <%=formatnumber(rs("valor"),2)%><br>
<%
rs.movenext:loop
rs.close

else ' tipoform=0
if ct_sexo="F" then v1="a" else v1="o"
if ct_sexo="F" then v2="a" else v2=""
if ct_sexo="F" then v3="" else v3="o"

datatemp=ct_mand_ini
datatemp=now()

sqlparcial="select 'aula'=case when aulas is null then 0 else aulas end, 'atividade'=case when atividades is null then 0 else atividades end " & _
"from (SELECT aulas=Sum(ta) FROM g2ch WHERE '" & dtaccess(datatemp) & "' Between inicio And termino and chapa1='" & ct_chapa & "') a, " & _
"(SELECT atividades=sum(case when codeve is null or codeve='' then 0 else ch end) FROM n_indicacoes WHERE '" & dtaccess(datatemp) & "' Between mand_ini And mand_fim and CHAPA='" & ct_chapa & "') b "
rs.Open sqlparcial, ,adOpenStatic, adLockReadOnly
taulas=rs("aula")
tatividades=rs("atividade")
tpercentual=tatividades/(taulas+tatividades)
rs.close
if tpercentual>=0.25 and (taulas+tatividades)>=12 then parcial=1 else parcial=0
parcial=0
'response.write "<br>Au " & taulas
'response.write "<br>At " & tatividades
'response.write "<br>Per " & tpercentual
'response.write "<br>Par " & parcial

if right(ct_nomeacao,3)="RTP" then
	texto99="a fim de compor e integrar o <b>Regime de Tempo Parcial</b>"
else
	texto99="a fim de compor 40 horas semanais, integrando o <b>Regime de Tempo Integral</b>"
end if
%>
<div align="right">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690" height="970">
<tr><td><img border="0" src="../images/aguia.jpg" width="236"></td> </tr>

<tr>
	<td class=campo><p align="center"><b><font size="3">ADENDO AO CONTRATO DE TRABALHO</font></b></p>
	</td>
</tr>

<tr>
	<td class=campo><p align="justify">Entre as partes, de um lado a <b>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</b>, 
	com sede a Av. Franz Voegeli, 300, Vila Yara, Osasco, CEP 06020-190, inscrita no CNPJ nº 73.063.166/0003-92, 
	denominada Contratante e de outro lado <%=v1%> Sr<%=v2%>. <b><%=ct_nome%></b> (<%=ct_chapa%>), residente e 
	domiciliad<%=v1%> à <%=ct_rua%>&nbsp;<%=ct_numero%>&nbsp;<%=ct_complemento%> - <%=ct_bairro%> - <%=ct_cidade%> - 
	CEP <%=ct_cep%>, portador<%=v2%> da CTPS nº <%=ct_ctps%>/<%=ct_serie%>, denominad<%=v1%> Professor<%=v2%>, 
	acordam o que se segue:</td>
</tr>

<tr>
	<td class=campo>
<%
if ct_ch>1 then 
		plural="horas semanais":pl="as quais serão registradas através de relatório de atividades mensal"
	else 
		plural="hora semanal":pl="a qual será registrada através de relatório de atividade mensal"
end if 

if session("id_nomeacao")=80 or request.form("nomeacao")="Programa de Atividades Gerais Acadêmicas" then
	pl="sendo que no mínimo 50% (cinquenta por cento) destas horas deverão ser presenciais na Instituição"
	complitem3=", e deverá ser comprovada através de relatório mensal vistado pelo responsável das atividades e entregue na Reitoria"
	ate=" até "
%>
	<p align="justify">1. <%=ucase(v1)%> Professor<%=v2%>, a partir de <%=ct_mand_ini%>, <%=texto99%>,
	com no máximo 20 (vinte) aulas semanais, passa a dedicar além as atividades propostas e demais horas no <b>Programa de Atividades Gerais Acadêmicas</b>, <%=pl%>.
<%
else
	if parcial=0 then
%>
	<p align="justify">1. <%=ucase(v1)%> Professor<%=v2%>, a partir de <%=ct_mand_ini%>, <%=pl0%> passa a dedicar <%=ate%>
	<%=ct_ch%>&nbsp;<%=plural%>&nbsp;<%=request.form("tipo_adendo")%>&nbsp;<%=ct_nomeacao%>, <%=pl%>.
<%
	else
		if (taulas+tatividades)>=40 then txtRegime="Integral" else txtRegime="Parcial"
%>
<!--
	<p align="justify">1. <%=ucase(v1)%> Professor<%=v2%>, a partir de <%=ct_mand_ini%>, a fim de se enquadrar no regime de trabalho parcial passa a dedicar
	no mínimo 12 (doze) horas semanais, reservando 25% (vinte e cinco por cento) de sua carga horária para estudos, planejamento, avaliação e orientação do
	corpo discente, integrando o <b>Regime de Tempo <%=txtRegime%></b>.
-->
<%	
	end if
end if
%>
	 </td>
</tr>
<%if session("id_nomeacao")=80 or request.form("nomeacao")="Programa de Atividades Gerais Acadêmicas" then%> 
<tr>
	<td class="campor"><p align="justify" style="margin-bottom:0px">1.1 Além das horas já designadas, deverá incluir atividades do tipo abaixo (até o limite de <%=ct_ch%>h):<br>
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] proferir palestras no âmbito institucional, seja para funcionários administrativos do Unifieo, para professores, alunos ou para a comunidade;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] proferir palestras externas em organizações, seminários, fórums ou congressos, representando a instituição;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] realizar pesquisas com divulgação de conhecimentos (apresentação em congressos, publicação de artigos em revistas ou periódicos);
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] realizar tarefas ligadas à gestão acadêmica e demais atividades em consonância com a sua formação e/ou necessidades da Instituição;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] praticar supervisão e/ou orientação de estágios, trabalhos de conclusão de curso ou atividades didáticas junto aos alunos do curso para qual for designado;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] acompanhar, monitorar ou assistir projetos de iniciação científica e/ou pesquisa dos alunos participantes destes projetos;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] exercer atividades, como aulas de reforço, visando a preparação dos alunos para o ENADE;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] apresentar e/ou participar de projetos internos, visando a melhoria dos cursos, alunos, funcionários e da própria instituição;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] participar de projetos de avaliação e/ou comissões especiais.
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] ministrar aulas em cursos de pós-graduação.
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] dar orientação didática ou as previstas em grade curricular aos alunos.
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] ministrar aulas ou orientações aos alunos em regime de dependências.
	</td>
</tr>
<%end if%>

<tr><td class=campo>
<%if parcial=0 then%>
	<p align="justify">2. <%=ucase(v1)%> Professor<%=v2%> perceberá o valor da hora atividade vigente atualmente fixado 
	em R$ <%=formatnumber(ct_valor,2)%>, independente da docência, sendo este valor destacado em seu holerite.
<%else%>
	<p align="justify">2. <%=ucase(v1)%> Professor<%=v2%> pela atividade de 
	<%=ct_ch%>&nbsp;<%=plural%>&nbsp;<%=request.form("tipo_adendo")%>&nbsp;<%=ct_nomeacao%>,
	perceberá o valor da hora atividade vigente atualmente fixado 
	em R$ <%=formatnumber(ct_valor,2)%>, independente da docência, sendo este valor destacado em seu holerite.
<%end if%>
</td></tr>

<tr>
	<td class=campo><p align="justify">3. Esta atividade tem início em <%=ct_mand_ini%> e término em <%=ct_mand_fim%>, 
	sendo <%=v1%> professor<%=v2%> nomead<%=v1%> através da <%=ct_portaria%><%=complitem3%>.</td>
</tr>

<tr>
	<td class=campo><p align="justify">4. No exercício de suas atividades está <%=(v1)%> Professor<%=v2%> sujeit<%=v1%> as normas 
	constantes de Regimento da Instituição de Ensino e do que prevê a legislação de ensino superior vigente.</td>
</tr>

<tr>
	<td class=campo><p align="justify">5. Finda a atividade estipulada nas cláusulas anteriores, <%=(v1)%> Professor<%=v2%> continuará 
	a exercer a docência, conforme o Contrato de Trabalho inicial ou a carga horária que estiver ministrando na época.</td>
</tr>

<tr>
	<td class=campo><p align="justify">6. O presente termo aditivo poderá ser extinto a qualquer momento, não gerando ônus para as partes, através
	de atos normativos da Instituição, ou quando ocorrer alterações na estrutura da instituição de ensino ou na legislação do ensino.</td>
</tr>

<tr>
	<td class="campop">E, por assim estarem de acordo, firmam o presente em 2 (duas) vias, uma das quais é entregue a<%=v3%> Professor<%=v2%>, 
	na presença de 2 (duas) testemunhas abaixo qualificadas.</td>
</tr>

<tr><td class="campop">
<%
if ct_contrato="" then ct_contrato=formatdatetime(now(),2)
dia=day(ct_contrato)
mes=monthname(month(ct_contrato))
ano=year(ct_contrato)
if session("id_nomeacao")=80 then tam_assi=100 else tam_assi=180
%>
		<p align="left">Osasco,&nbsp;<%=dia & " de " & mes & " de " & ano %></td>
</tr>


<tr><td>
		<table border="0" width="100%" cellspacing="0">
		<tr>
			<td width="50%" valign="bottom"><img style="border-bottom:1px solid #000000" border="0" src="../images/assi_rmsa.jpg" width="<%=tam_assi%>"><br>
				
				FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td>
			<td width="50%">&nbsp;
				<p>_______________________________________<br>
				<b><%=ct_nome%></b></td>
		</tr>
		</table>
	</td>
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

<tr><td><p align="right"><font size=1><%=ct_secao%></font></p></td></tr>

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
	sqlz="UPDATE n_indicacoes SET CONTRATO = '" & dtaccess(ct_contrato) & "' "
	sqlz=sqlz & " WHERE id_indicado=" & id_indicado
	'response.write sqlz	
	conexao.execute sqlz
end if

set rs=nothing
conexao.close
set conexao=nothing
%>