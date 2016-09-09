<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a89")="N" or session("a89")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Emissão de Avaliação de Professor Recem-Contratado</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"><!--
function nome1() {	form.chapa.value=form.nome.value; }
function chapa1() {	form.nome.value=form.chapa.value; }
--></script>
</head>
<body style="margin-left:20px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
sessao=session.sessionid
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

espacamento=5
if request.form="" then
sql="select p.chapa, p.nome, p.dataadmissao from corporerm.dbo.pfunc p where p.chapa<'10000' and p.codtipo='N' and codsituacao<>'D' and codsindicato='03' " & _
"order by p.dataadmissao desc, p.nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form name="form" action="avaliacao_prof.asp" method="post">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Seleção de Professor para emissão da Avaliação</td>
</tr>
<tr>
	<td class=campo>Professor</td>
	<td class=campo>
		<select name="chapa" class=a size=20 multiple>
		<option value="0"> Selecione o professor</option>
<%
rs.movefirst
do while not rs.eof
%>
		<option value="<%=rs("chapa")%>"> <%=rs("chapa") & " - " & rs("nome") & " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(" & rs("dataadmissao") & ")"%></option>
<%
rs.movenext
loop
rs.close
%>
		</select>
	</td>
</tr>
<tr>
	<td class=campo colspan=3>&nbsp;
		<input type="submit" value="Visualizar" class=button name="B1">
	</td>
</tr>
</table>
</form>

<%
else
'response.write request.form("chapa").count
chapas=request.form("chapa").count
%>
<table border="0" cellpadding="1" cellspacing="0" width="930" height=470 style="border-collapse: collapse">
<%
'****
sql="delete from ttavalprof where sessao='" & session("usuariomaster") & "'"
conexao.execute sql
dim perg(11), resp(11,4)
perg(1)="Apresentação Pessoal"
perg(2)="Adaptação e respeito às normas da Instituição"
perg(3)="Chegada e Saída da sala de aula no horário correto"
perg(4)="Com relação ao plano de ensino, cumpriu de forma"
perg(5)="Utilização de criatividade, conteúdos atualizados e recursos audio-visuais"
perg(6)="Transmite seus conhecimentos de modo"
perg(7)="Possui relacionamento com colegas e coordenador de forma"
perg(8)="Coopera e participa das atividades não docentes"
perg(9)="Houve reclamações dos alunos"
perg(10)="De modo geral, como avalia o desempenho desse docente"
perg(11)="V.Sa. indicaria este professor para outro curso ou disciplina?"

resp(1 ,1)="Excelente":resp(1 ,2)="Boa":resp(1 ,3)="Regular":resp(1 ,4)="Ruim"
resp(2 ,1)="Excelente":resp(2 ,2)="Boa":resp(2 ,3)="Regular":resp(2 ,4)="Ruim"
resp(3 ,1)="Excelente":resp(3 ,2)="Boa":resp(3 ,3)="Regular":resp(3 ,4)="Ruim"
resp(4 ,1)="Excelente":resp(4 ,2)="Boa":resp(4 ,3)="Regular":resp(4 ,4)="Ruim"
resp(5 ,1)="Excelente":resp(5 ,2)="Boa":resp(5 ,3)="Regular":resp(5 ,4)="Ruim"
resp(6 ,1)="Excelente":resp(6 ,2)="Boa":resp(6 ,3)="Regular":resp(6 ,4)="Ruim"
resp(7 ,1)="Excelente":resp(7 ,2)="Boa":resp(7 ,3)="Regular":resp(7 ,4)="Ruim"
resp(8 ,1)="Excelente":resp(8 ,2)="Boa":resp(8 ,3)="Regular":resp(8 ,4)="Ruim"
resp(9 ,1)="Nenhuma"  :resp(9 ,2)="Um/um motivo":resp(9 ,3)="Mais de uma":resp(9 ,4)="Muitas/Vários Motivos"
resp(10,1)="Excelente":resp(10,2)="Boa":resp(10,3)="Regular":resp(10,4)="Ruim"

for a=1 to chapas
'****
chapa=request.form("chapa").item(a)
sql="select f.chapa, f.nome, f.dataadmissao, f.sexo, s.codevento, coddoc, curso " & _
"from dc_professor f, corporerm.dbo.pfsalcmp s, g2cursoeve c " & _
"where f.chapa collate database_default=s.chapa and c.sal=s.codevento collate database_default and f.chapa='" & chapa & "' order by coddoc, nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
sufixo=" no "
do while not rs.eof
if rs("sexo")="F" then s1="a" else s1="o"
if rs("sexo")="F" then s2="a" else s2=""
sql2="insert into ttavalprof (sessao, chapa, coddoc) select '" & session("usuariomaster") & "','" & rs("chapa") & "','" & rs("coddoc") & "'"
conexao.execute sql2
%>
<table border="0" cellpadding="1" cellspacing="0" width="690" height=990 style="border-collapse: collapse">
<tr><td valign=top>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class=campo align="left"><img src="../images/logo_centro_universitario_unifieo_big.jpg" width="150"  border="0" alt=""></td>
	<td class="campop" align="center"><i><b>AVALIAÇÃO DE PROFESSOR RECEM-CONTRATADO</td></tr>
<tr><td colspan=2 class="campop" align="center">Relatório de Acompanhamento Funcional</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" height=10></td></tr>
<tr>
	<td class="campop" valign=top height=40><font style="font-size:7pt"><i><b>Docente Avaliado:</font></b></i><br>
	&nbsp;<%=rs("chapa")%>&nbsp;&nbsp;<%=rs("nome")%><td>
	<td class=campo rowspan=3 align="right">
	<img border="0" src="../func_foto.asp?chapa=<%=rs("chapa")%>"  height="120">
	</td>
</tr>
<tr>
	<td class="campop" valign=top height=40><font style="font-size:7pt"><i><b>Curso Principal:</font></b></i><br>
	&nbsp;<%=rs("curso")%><td>
</tr>
<tr>
	<td class="campop" valign=top height=40><font style="font-size:7pt"><i><b>Data de admissão:</font></b></i><br>
	&nbsp;<%=rs("dataadmissao")%><td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" height=5></td></tr>
<tr><td class="campop">Prezado(a) Coordenador(a)<br>
	Preencha a avaliação com base no desempenho d<%=s1%> professor<%=s2%> e ao término devolva ao Deptº de Recursos Humanos.
	</td></tr></table>

<%
for b=1 to 10
%>
<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" height=10></td></tr>
<tr><td class="campop" align="center" rowspan=2 width=25 style="border-right:1px solid #000000"><b> <%=b%> </td>
	<td class=campo width=5></td>
	<td class="campop" colspan=9 height=20><b> <%=perg(b)%> </td>
</tr>
<tr>
	<td class=campo width=5></td><%col=5%>
	<%for c=1 to 4%>
	<%if b=9 and c=4 then col2=150 else col2=100%>
	<td class=campo widht=30><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td><td class=campo width=<%=col2%> nowrap>&nbsp; <%=resp(b,c)%> </td>
	<%col=col+30+col2%>
	<%next%>
	<td class=campo width=<%=690-col%>></td>
</tr>
</table>
<%
next
%>
<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" height=10></td></tr>
<tr><td class="campop" align="center" rowspan=5 width=25 style="border-right:1px solid #000000"><b> 11 </td>
	<td class=campo width=5></td>
	<td class="campop" colspan=9 height=20><b> <%=perg(11)%> </td>
</tr>
<tr>
	<td class=campo width=5></td><%col=5%>
	<td class=campo widht=30><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td><td class=campo width=100 nowrap>&nbsp; Sim </td><%col=col+130%>
	<td class=campo widht=30><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td><td class=campo width=100 nowrap>&nbsp; Não </td><%col=col+130%>
	<td class=campo width=<%=690-col%>></td>
</tr>
<tr><td class=campo width=5></td>
	<td class="campop" colspan=9 height=20 style="border-bottom:1px dashed #000000">Justifique:</td>
</tr>
<tr><td class=campo width=5></td>
	<td class="campop" colspan=9 height=20 style="border-bottom:1px dashed #000000"></td>
</tr>
<tr><td class=campo width=5></td>
	<td class="campop" colspan=9 height=20 style="border-bottom:1px dashed #000000"></td>
</tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
<tr><td class="campor" height=10></td></tr>
<tr>
	<td width=345 height=50 class="campor" valign=top style="border-top:1px solid;border-left:1px solid;border-bottom:1px solid">&nbsp;Avaliador/Coordenador</td>
	<td width=230 class="campor" valign=top style="border-top:1px solid;border-left:1px solid;border-bottom:1px solid">&nbsp;Assinatura</td>
	<td width=115 class="campor" valign=top style="border:1px solid">&nbsp;Data</td>
</td>
</table>


</td></tr>
<tr><td class=campo height=100%></td></tr>
</table>
<%
rs.movenext
response.write "<DIV style=""page-break-after:always""></DIV>"
loop
rs.close

'****
next
'****
response.write "</table>"
%>
<%
'******************* cartas para acompanhar

sql3="select a.coddoc, curso, coordenador from ttavalprof a, g2cursoeve c where sessao='" & session("usuariomaster") & "' and c.coddoc=a.coddoc group by a.coddoc, curso, coordenador "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
<table border="0" cellpadding="1" cellspacing="0" width="690" height=990 style="border-collapse: collapse">
<tr><td valign=top>

<table border="0" cellpadding="0" cellspacing="0" width=100% style="border-collapse: collapse">
	<tr><td class="campop"><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width=225></td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><b>Of. RH-<%=rs("coddoc")%></b></td></tr>
	<tr><td class="campop" align="right">
	<input type="text" name="txt1" class="form_input" size="29" value="Osasco, <%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %>" style="font-size:10pt">
	</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><input type="text" name="txt0" class="form_input" size="5" value=""><br>
	<input type="text" name="txt1" class="form_input" size="60" value="Ilmo(a) Sr(a). <%=rs("coordenador")%>" style="font-size:10pt"><br>
	<input type="text" name="txt2" class="form_input" size="60" value="<%=rs("curso")%>" style="font-size:10pt"><br>
	</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Em atendimento a vossa solicitação, foram contratados para prestarem serviços a esse curso, os seguintes professores:

	<div align="center">
	<table border="1" cellpadding="3" cellspacing="0" width=80% style="border-collapse: collapse">
	<tr><td class=titulop>Nome</td>
		<td class=titulop align="center">Admissão</td>
	</tr>
<%
sql4="select a.chapa, f.nome, f.dataadmissao from ttavalprof a, dc_professor f where a.chapa=f.chapa and a.coddoc='" & rs("coddoc") & "' and a.sessao='" & session("usuariomaster") & "'"
rs2.Open sql4, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>
	<tr><td class="campop"><%=rs2("nome")%></td>
		<td class="campop" align="center"><%=rs2("dataadmissao")%></td>
	</tr>
<%
rs2.movenext
loop
rs2.close
%>	
	</table></div>
	<p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Solicitamos que V.Sa. se manifeste por escrito, na avaliação anexa sobre o desempenho dos referidos professores, 
	informando-nos, no prazo de 10 (dez) dias, se os mesmos atendem as exigências do cargo e preenche 
	os requisitos indispensáveis. Em resumo, se amolda aos padrões da FIEO.</td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Contando com a costumeira 
	colaboração de V.Sa. apresentamos nossas cordiais saudações.</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><p align="center" style="line-height: 25px"><b>LUIZ FERNANDO DA COSTA E SILVA<br>Pró-Reitor Acadêmico</b>
	</td></tr>
	<tr><td class="campop"></td></tr>
</table>
	
</td></tr>
<tr><td class=campo height=100%></td></tr>
</table>
<%
rs.movenext
response.write "<DIV style=""page-break-after:always""></DIV>"
loop
rs.close
%>

<%
end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>