<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a91")="N" or session("a91")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
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

if request("codigo")="" then
	compl=" order by nome, pini "
	if request("ordem")="curso" then compl=" order by curso, nome, pini "
	if request("ordem")="disciplina" then compl="order by materia, nome "
	if request("ordem")="nome" then compl=" order by nome, pini "
	if request("ordem")="chapa" then compl="order by chapa1, pini "
	if request("ordem")="inicio" then compl="order by pini, nome "
	
	temp=1
	sql1="select *, status=case when autorizado=1 then 'Reemissão' else 'Nova' end from ( " & _
	"select g.id_grdaula, g.perlet, g.codtur, c.curso, g.codmat, m.materia, g.chapa1, f.nome, pini=min(data), pfim=max(data), inicio, termino, autorizado, quando, ch=sum(aulas) " & _
	"from (((g5ch g inner join g2cursos c on c.coddoc=g.coddoc and c.codcur=g.codcur and c.codper=g.codper) " & _
	"inner join corporerm.dbo.umaterias m on m.codmat collate database_default=g.codmat) " & _
	"inner join grades_aux_prof f on f.chapa=g.chapa1) " & _
	"inner join g5datas d on d.id_grdaula=g.id_grdaula " & _
	"where g.deletada=0 and (g.chapa1<'10000') and d.c_emitido=0 " & _
	"group by g.id_grdaula, g.perlet, g.codtur, c.curso, g.codmat, m.materia, g.chapa1, f.nome, inicio, termino, autorizado, quando) z " & _
	"where ((autorizado=0 or autorizado is null) or pini<>inicio or pfim<>termino)  and pini>getdate()-120 " & _
	"and z.chapa1 not in (select chapa collate database_default from corporerm.dbo.pfsalcmp where codevento in ('255','256','257','258','128','138','RHT')) "  & _
	" and perlet IN ('2016/5','2016/6') " & compl

	sql2="union " & _
	"select *, status='RHT/RT' from ( " & _
	"select g.id_grdaula, g.perlet, g.codtur, c.curso, g.codmat, m.materia, g.chapa1, f.nome, pini=min(data), pfim=max(data), inicio, termino, autorizado, quando, ch=sum(aulas) " & _
	"from (((g5ch g inner join g2cursos c on c.coddoc=g.coddoc and c.codcur=g.codcur and c.codper=g.codper) " & _
	"inner join corporerm.dbo.umaterias m on m.codmat collate database_default=g.codmat) " & _
	"inner join grades_aux_prof f on f.chapa=g.chapa1) " & _
	"inner join g5datas d on d.id_grdaula=g.id_grdaula " & _
	"where g.deletada=0 and (g.chapa1<'10000') and d.c_emitido=0 " & _
	"group by g.id_grdaula, g.perlet, g.codtur, c.curso, g.codmat, m.materia, g.chapa1, f.nome, inicio, termino, autorizado, quando) z " & _
	"where ((autorizado=0 or autorizado is null) or pini<>inicio or pfim<>termino)  and pini>getdate()-120 " & _
	"and z.chapa1 in (select chapa collate database_default from corporerm.dbo.pfsalcmp where codevento in ('255','256','257','258','128','138','RHT')) " 
	sql=sql1 '& sql2
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	titulo=""
else
	temp=0
	sql="select *, status=case when autorizado=1 then 'Reemissão' else 'Nova' end from ( " & _
	"select g.id_grdaula, g.perlet, g.codtur, c.curso, g.codmat, m.materia, g.chapa1, f.nome, pini=min(data), pfim=max(data), inicio, termino, autorizado, quando, ch=sum(aulas) " & _
	"from (((g5ch g inner join g2cursos c on c.coddoc=g.coddoc and c.codcur=g.codcur and c.codper=g.codper) " & _
	"inner join corporerm.dbo.umaterias m on m.codmat collate database_default=g.codmat) " & _
	"inner join grades_aux_prof f on f.chapa=g.chapa1) " & _
	"inner join g5datas d on d.id_grdaula=g.id_grdaula " & _
	"where g.deletada=0 and (g.chapa1<'10000') and d.c_emitido=0 " & _
	"group by g.id_grdaula, g.perlet, g.codtur, c.curso, g.codmat, m.materia, g.chapa1, f.nome, inicio, termino, autorizado, quando) z " & _
	"where ((autorizado=0 or autorizado is null) or pini<>inicio or pfim<>termino) " & _
	"and id_grdaula=" & request("codigo")
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if temp=1 then
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td class=grupo>Emissão de Adendo de Contrato para professores (<%=rs.recordcount%>)</td></tr>
</table>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=titulo align="center"><a class="r" href="pos_contrato.asp?ordem=chapa"><b>#</a></td>
	<td class=titulo align="center"><a class="r" href="pos_contrato.asp?ordem=nome"><b>Nome</a></td>
	<td class=titulo align="center"><a class="r" href="pos_contrato.asp?ordem=curso"><b>Curso</a></td>
	<td class=titulo align="center"><a class="r" href="pos_contrato.asp?ordem=disciplina"><b>Disciplina</a></td>
	<td class=titulo align="center"><a class="r" href="pos_contrato.asp?ordem=inicio"><b>Período</a></td>
	<td class=titulo align="center">Aulas</td>
	<td class=titulo align="center"></td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo style="border-bottom:2px dotted"><%=rs("chapa1")%></td>
	<td class=campo style="border-bottom:2px dotted"><a href="pos_contrato.asp?codigo=<%=rs("id_grdaula")%>"><%=rs("nome")%></a></td>
	<td class=campo style="border-bottom:2px dotted"><%=rs("curso") %></td>
	<td class=campo style="border-bottom:2px dotted"><%=rs("materia") %> (<%=rs("codtur")%>)</td>
	<td class=campo style="border-bottom:2px dotted"><%=rs("pini") & " a " & rs("pfim") %></td>
	<td class=campo style="border-bottom:2px dotted"><%=rs("ch") %></td>
	<td class=campo style="border-bottom:2px dotted"><%=rs("status") %></td>
</tr>
<%
rs.movenext
loop
rs.close
else
	response.write "<tr><td class=""campop"" colspan=6 style=""border-bottom:2px dotted"">Sem contratos a emitir.</td></tr>"
end if
%>
</table>
<%
else ' temp=0

sql2="select nome, sexo, rua, numero, complemento, bairro, cidade, estado, cep, carteiratrab, seriecarttrab, secao, titulacaopaga, grauinstrucao from dc_professor where chapa='" & rs("chapa1") & "'"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
sexo=rs2("sexo"):t=rs2("titulacaopaga"):t=rs2("grauinstrucao")
if sexo="M" then v1="o" else v1="a"
if sexo="M" then v2="" else v2="a"
rs2.close
response.write t
if t<="B" then          sqlv="select valor from corporerm.dbo.pvalfix where codigo='POSD' and '" & dtaccess(rs("pini")) & "' between iniciovigencia and finalvigencia "
if t>"B" and t<"E" then sqlv="select valor from corporerm.dbo.pvalfix where codigo='POSE' and '" & dtaccess(rs("pini")) & "' between iniciovigencia and finalvigencia "
if t>="E" then          sqlv="select valor from corporerm.dbo.pvalfix where codigo='POSF' and '" & dtaccess(rs("pini")) & "' between iniciovigencia and finalvigencia "
rs2.Open sqlv, ,adOpenStatic, adLockReadOnly
valor=rs2("valor"):valor=cdbl(valor)
valor1=valor*0.05
valor2=(valor+valor1)*0.1667
valort=valor+valor1+valor2
rs2.close
rs2.Open sql2, ,adOpenStatic, adLockReadOnly

%>

<div align="right">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690" height="970">
<tr><td><img border="0" src="../images/aguia.jpg" width="236"></td> </tr>

<tr>
	<td class=campo><p align="center"><b><font size="3">ADENDO AO CONTRATO DE TRABALHO</font></b></p>
		<p align="center">&nbsp;</td>
</tr>

<tr>
	<td class=campo><p align="justify">Entre as partes, de um lado a <b>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</b>, 
	com sede a Av. Franz Voegeli, 300, Vila Yara, Osasco, CEP 06020-190, inscrita no CNPJ nº 73.063.166/0003-92, 
	denominada Contratante e de outro lado <%=v1%> Sr<%=v2%>. <b><%=rs2("nome")%></b> (<%=rs("chapa1")%>), residente e 
	domiciliad<%=v1%> à <%=rs2("rua")%>&nbsp;<%=rs2("numero")%>&nbsp;<%=rs2("complemento")%> - <%=rs2("bairro")%> - <%=rs2("cidade")%> - 
	CEP <%=rs2("cep")%>, portador<%=v2%> da CTPS nº<%=rs2("carteiratrab")%>/<%=rs2("seriecarttrab")%>, denominad<%=v1%> Professor<%=v2%>, 
	acordam o que se segue:</td>
</tr>

<tr>
	<td class=campo><p align="justify">1. <%=ucase(v1)%> Professor<%=v2%>, passa a ministrar o módulo de <b><%=rs("materia")%></b> no curso 
	<b><%=rs("curso")%></b> de pós-graduação, no período de <b><%=rs("pini")%></b> a <b><%=rs("pfim")%></b>, na turma <%=rs("codtur")%>.</td>
</tr>

<tr>
	<td class=campo><p align="justify">2. <%=ucase(v1)%> Professor<%=v2%> perceberá o valor fixo de R$ <%=formatnumber(valor,2)%>  por aula no módulo, incluindo-se
	o adicional de hora atividade e DSR, totalizando R$ <%=formatnumber(valort,2)%>, independente da docência na graduação, sendo este valor destacado em seu holerite.</td>
</tr>

<tr>
	<td class=campo><p align="justify">3. No exercício de suas atividades está <%=(v1)%> Professor<%=v2%> sujeit<%=v1%> as normas 
	constantes de Regimento da Instituição de Ensino e do que prevê a legislação de ensino superior vigente.</td>
</tr>

<tr>
	<td class=campo><p align="justify">4. Finda a atividade estipulada nas cláusulas anteriores, <%=(v1)%> Professor<%=v2%> continuará 
	a exercer a docência, conforme o Contrato de Trabalho inicial ou a carga horária que estiver ministrando na época.</td>
</tr>

<tr>
	<td class="campop">E, por assim estarem de acordo, firmam o presente em 2 (duas) vias, uma das quais é entregue a<%=v3%> Professor<%=v2%>, 
	na presença de 2 (duas) testemunhas abaixo qualificadas.</td>
</tr>

<tr><td class=campo>&nbsp;</td></tr>

<tr><td class="campop">
<%
if ct_contrato="" then ct_contrato=formatdatetime(now(),2)
dia=day(ct_contrato)
mes=monthname(month(ct_contrato))
ano=year(ct_contrato)
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
				<b><%=rs2("nome")%></b></td>
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

<tr><td><p align="right"><font size=1><%=rs2("secao")%></font></p></td></tr>

<tr><td><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b></td> </tr>
<tr><td><font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000</font></td></tr>
<tr><td><font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP 06020-190 - Fone: (011) 3651-9999</font></td></tr>
<tr><td><font face="Arial Narrow">Caixa Postal - ACF - Jd. Ipê - nº 2530 - Osasco - SP - CEP 06053-990</font></td></tr>
</table>
</div>


<%
sqlu="update g2aulas set inicio=ini, termino=fim " & _
"from g2aulas a inner join " & _
"(select id_grdaula, min(data) ini, max(data) fim from g2aulasdata where deletada=0 group by id_grdaula) d on a.id_grdaula=d.id_grdaula " & _
"where a.id_grdaula=" & rs("id_grdaula")
conexao.execute sqlu

rs2.close
rs.close
end if 'temp=0

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>