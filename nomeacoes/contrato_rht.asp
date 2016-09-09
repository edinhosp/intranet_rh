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

if request("codigo")<>"" or request.form("idnomeacao")<>"" then
	tipoform=2
	if request.form("idnomeacao")<>"" then idnomeacao=request.form("idnomeacao") else idnomeacao=request("codigo")
	'idnomeacao=request("codigo")
	sqlc="SELECT i.id_nomeacao, n.NOMEACAO, i.PORTARIA, i.id_indicado, i.CHAPA, " & _
	"i.NOME, i.CARGO, i.MAND_INI, i.MAND_FIM, i.alunos, i.CH, i.OBS, i.CONTRATO, " & _
	"p.SEXO, p.RUA, p.NUMERO, p.COMPLEMENTO, p.BAIRRO, p.CIDADE, p.CEP, p.FUNCAO, " & _
	"p.CARTEIRATRAB, p.SERIECARTTRAB, p.secao, p.instrucao, p.grauinstrucao " & _
	"FROM n_indicacoes AS i, n_nomeacoes AS n, dc_professor AS p " & _
	"WHERE i.id_nomeacao = n.id_nomeacao AND i.CHAPA = p.CHAPA collate database_default " & _
	"and i.id_indicado=" & idnomeacao & " " & _
	" ORDER BY n.nomeacao "
	rs.Open sqlc, ,adOpenStatic, adLockReadOnly
	ct_chapa      =rs("chapa")         : ct_nome  =rs("nome")
	ct_rua        =rs("rua")           : ct_numero=rs("numero")
	ct_complemento=rs("complemento")   : ct_bairro=rs("bairro")
	ct_cidade     =rs("cidade")        : ct_cep   =rs("cep")
	ct_ctps       =rs("carteiratrab")  : ct_serie =rs("seriecarttrab")
	ct_funcao     =rs("funcao")
	ct_mand_ini   =rs("mand_ini")
	ct_mand_fim   =rs("mand_fim")
	ct_portaria   =rs("portaria") : ct_sexo=rs("sexo")
	ct_contrato   =rs("contrato") : ct_secao=rs("secao")
	ct_instrucao  =rs("instrucao") : ct_grau=rs("grauinstrucao")
	if ct_horas="" then ct_horas=40 else ct_horas=request.form("ct_horas")
	'if (rs("contrato")="" or isnull(rs("contrato"))) and request.form("data_assinatura")<>"" then ct_contrato=ct_mand_ini 'request.form("data_assinatura")
	'if request.form("data_assinatura")<>"" and request.form("data_assinatura")<>ct_contrato then ct_contrato=request.form("data_assinatura")
	if ct_contrato="" or isnull(ct_contrato) then ct_contrato=ct_mand_ini
	dia=day(ct_contrato)
	mes=monthname(month(ct_contrato))
	ano=year(ct_contrato)
	rs.close
end if

if request.form="" then
%>
<p class=titulo>Adendo ao Contrato de Trabalho para&nbsp;<%=titulo %>
<form method="POST" action="contrato_rht.asp" name="form">
<input type="hidden" name="idnomeacao" value="<%=idnomeacao%>">
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
sql2="select chapa, nome from dc_professor where codsituacao<>'D' and codsindicato='03' "
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
<tr><td class=titulo>Portaria de Nomeação</td></tr>
<tr>
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
	<td class=fundo><input type="text" value="<%=ct_horas%>" name="horas" size="3"></td>
	<td class=fundo><input type="text" value="<%if ct_contrato="" or isnull(ct_contrato) then response.write ct_mand_ini else response.write ct_contrato%>" name="data_assinatura" size="14"></td>
</tr>
</table>
<table border="0" width="500" cellspacing="0" cellpadding="3">
<tr><td class=titulo><font color=blue>Opcional: *</font> outras atividades a relacionar</td></tr>
<tr><td class=fundo><input type="text" value="<%%>" name="ativ1" size="70"></td></tr>
<tr><td class=fundo><input type="text" value="<%%>" name="ativ2" size="70"></td></tr>
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
tipoform=0
if ct_sexo="F" then v1="a" else v1="o"
if ct_sexo="F" then v2="a" else v2=""
if ct_sexo="F" then v3="" else v3="o"
if ct_chapa="02303" or ct_chapa="02757" then ct_grau="D"
'if ct_chapa="00892" then ct_grau="B"
'if ct_chapa="03112" then ct_grau="9"

ct_contrato=request.form("data_assinatura")
if ct_contrato="" or isnull(ct_contrato) then ct_contrato=ct_mand_ini
dia=day(ct_contrato):mes=monthname(month(ct_contrato)):ano=year(ct_contrato)

sqlv="select '9'=sum(case when CODIGO='RHT0' then VALOR else 0 end),'B'=sum(case when CODIGO='RHT1' then VALOR else 0 end),'D'=sum(case when CODIGO='RHT2' then VALOR else 0 end),'F'=sum(case when CODIGO='RHT3' then VALOR else 0 end) from corporerm.dbo.PVALFIX where '" & dtaccess(ct_contrato) & "' between INICIOVIGENCIA and FINALVIGENCIA and CODIGO in ('RHT0','RHT1','RHT2','RHT3')"
rs2.Open sqlv, ,adOpenStatic, adLockReadOnly
select case ct_grau
	case "B"
		valorht=cdbl(rs2("B"))
	case "D"
		valorht=cdbl(rs2("D"))
	case "F"
		valorht=cdbl(rs2("F"))
	case else
		valorht=cdbl(rs2("9"))
end select
rs2.close

%>
<div align="right">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690" height="970">
<tr><td><img border="0" src="../images/aguia.jpg" width="236"></td> </tr>

<tr>
	<td class=campo><p align="center"><b><font size="3">ADENDO AO CONTRATO DE TRABALHO</font></b></p></td>
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
    <% if ct_portaria<>"" then 
		txtportaria="e nomeado através da " & ct_portaria & ", "
	else 
		txtportaria=""
	end if %>
	<p align="justify">1. <%=ucase(v1)%> Professor<%=v2%>, a partir de <%=ct_mand_ini%>, a fim de compor <%=ct_horas%> (<%=extenson(ct_horas)%>) horas semanais, <%=txtportaria%>passa a 
	dedicar no máximo 20 (vinte) aulas semanais em sala de aula e demais horas no RHT <%=ct_horas%>, integrando o <b>Regime de Tempo Integral</b>.
	</td>
</tr>


<%
	final1=".":final2=".":final3=".":ativ1="":ativ2=""
	if request.form("ativ2")<>"" then final2=";":ativ2="<p align='justify' style='text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px'>[&nbsp;&nbsp;] " & request.form("ativ2") & final3 & ""
	if request.form("ativ1")<>"" then final1=";":ativ1="<p align='justify' style='text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px'>[&nbsp;&nbsp;] " & request.form("ativ1") & final2 & "</dd>"
%>
<tr>
	<td class="campor"><p align="justify" style="margin-bottom:0px">1.1 O Regime de Horas Trabalhadas, inclui as seguintes atividades:<br>
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] realizar tarefas ligadas à gestão acadêmica e demais atividades em consonância com a sua formação e/ou necessidades da Instituição, sob designação da Pró-Reitoria Acadêmica;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] proferir palestras no âmbito institucional, seja para funcionários administrativos do UNIFIEO, para professores, alunos ou para a comunidade;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] proferir palestras externas em organizações, seminários, fórums ou congressos, sempre <u>representando esta instituição</u>;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] realizar pesquisas com divulgação de conhecimentos (publicações anais de congressos, publicação de artigos em revistas ou periódicos), sempre em nome exclusivo desta Instituição;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] praticar supervisão e/ou orientação de estágios, trabalhos de conclusão de curso ou atividades didáticas junto aos alunos do curso para qual for designado;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] acompanhar, monitorar ou assistir projetos de iniciação científica e/ou pesquisa dos alunos participantes destes projetos;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] exercer atividades, como aulas de reforço, visando nivelamento ou à preparação dos alunos para o ENADE;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] apresentar e/ou participar de projetos internos, visando a melhoria dos cursos, alunos, funcionários e da própria instituição;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] participar de projetos de avaliação e/ou comissões especiais;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] ministrar aulas em cursos de pós-graduação;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] dar orientação didática ou as previstas em grade curricular aos alunos;
<p align="justify" style="text-indent:-25px;margin-left:25px;margin-top:0px;margin-bottom:0px;font-size:11px">[&nbsp;&nbsp;] dar aulas ou orientações aos alunos em regime de dependências<%=final1%>
	<%=ativ1%>
	<%=ativ2%>
	</td>
</tr>
<tr>
	<td class=campo><p align="justify">2. <%=ucase(v1)%> Professor<%=v2%> receberá um valor mensal atualmente fixado em R$ <%=formatnumber(valorht/40*ct_horas,2)%> (<%=extenso2(valorht/40*ct_horas)%>), já incluido os adicionais legais (DSR e Adicional Hora Atividade).</td>
</tr>

<tr>
	<td class=campo><p align="justify">3. As atividades deverão ser comprovadas através de relatório mensal vistado pelo responsável
	das atividades e entregue na Reitoria.</td>
</tr>

<tr>
	<td class=campo><p align="justify">4. No exercício de suas atividades está <%=(v1)%> Professor<%=v2%> sujeit<%=v1%> as normas 
	constantes de Regimento da Instituição de Ensino e do que prevê a legislação de ensino superior vigente.</td>
</tr>

<tr>
	<td class=campo><p align="justify">5. O presente contrato vigorará até <%=ct_mand_fim%>.</td>
</tr>

<tr>
	<td class="campop">E, por assim estarem de acordo, firmam o presente em 2 (duas) vias, uma das quais é entregue a<%=v3%> Professor<%=v2%>, 
	na presença de 2 (duas) testemunhas abaixo qualificadas.</td>
</tr>

<tr><td class=campo>&nbsp;</td></tr>

<tr><td class="campop">
<%
%>
		<p align="left">Osasco,&nbsp;<%=dia & " de " & mes & " de " & ano %></td>
</tr>


<tr><td>
		<table border="0" width="100%" cellspacing="0">
		<tr>
			<td width="50%">&nbsp;
				<p>_______________________________________<br>
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
if tipoform=0 and idnomeacao<>"" then
	sqlz="UPDATE n_indicacoes SET CONTRATO = '" & dtaccess(request.form("data_assinatura")) & "' "
	sqlz=sqlz & " WHERE id_indicado=" & idnomeacao
	'response.write sqlz
	conexao.execute sqlz
end if

set rs=nothing
conexao.close
set conexao=nothing
%>