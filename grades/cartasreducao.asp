<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a80")="N" or session("a8")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Carta aos professores</title>
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
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

sqlb="select top 5 t.chapa1 chapa, f.nome, sexo, bloco, secao,  f.codsituacao, [20121], [20122] " &_
"from totalizador_2ch t " & _
"inner join dc_professor f on f.chapa=t.chapa1 " & _
"inner join blocos b on b.codsecao=f.codsecao collate database_default " & _
"where [20122]<[20121] and f.codsituacao in ('A','F','E','Z') and chapa1>'0' " & _
"order by bloco, t.chapa1 "
rs.Open sqlb, ,adOpenStatic, adLockReadOnly

rs.movefirst
do while not rs.eof
%>
<div align="right">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td><img border="0" src="../images/logo_centro_universitario_unifieo_big.jpg" width=225></td><td width=35 rowspan=7></td></tr>
<tr><td><p align="center">&nbsp;</td></tr>
<tr>
	<td>
	Osasco, 13 de agosto de 2012<br>
    <br><b><%=rs("nome")%></b> (<%=rs("chapa")%>)<br>
    Bloco: <%=rs("bloco")%>&nbsp;-&nbsp;<%=rs("secao")%><br><br><br>
    Ref.: Redução da carga horária semestral de aulas no 2º semestre de 2012<br><br></p>
	<br>
	<%if rs("sexo")="F" then 
		response.write "Prezada Professora" 
		t1="a":t2="a"
	else 
		response.write "Prezado Professor"
		t1="o":t2=""
	end if%>
	<br>
    <p align="justify">Vimos informa-lhe que de acordo com a cláusula 22 da Convenção Coletiva do Sindicato
	dos Professores, que as suas aulas nos cursos de graduação, foram reduzidas de <%=rs("20121")%> para
	<%=rs("20122")%> aulas semanais, em decorrência de extinção e ou junção de disciplina, classe ou turma.</p>
	
	<div align="center">
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="450">
<tr>
	<td class=titulo rowspan=2>Curso</td>
	<td class=titulo colspan=2 align="center">Aulas no</td>
</tr>
<tr>
	<td class=titulo align="center">1º semestre</td>
	<td class=titulo align="center">2º semestre</td>
</tr>
<%
sqlc="select distinct g.chapa1, g.coddoc, curso, s1.tot1, s2.tot2 " & _
"from g2ch g " & _
"inner join g2cursoeve c on c.coddoc=g.coddoc " & _
"left join ( " & _
"select chapa1, coddoc, tot1=sum(ta) from g2ch where termino='07/31/2010' and deletada=0 group by chapa1, coddoc " & _
") s1 on s1.chapa1=g.chapa1 and s1.coddoc=g.coddoc " & _
"left join ( " & _
"select chapa1, coddoc, tot2=sum(ta) from g2ch where termino='01/31/2011' and deletada=0 group by chapa1, coddoc " & _
") s2 on s2.chapa1=g.chapa1 and s2.coddoc=g.coddoc " & _
"where termino in ('07/31/2010','01/31/2011') and g.chapa1='" & rs("chapa") & "' " 
rs2.Open sqlc, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
%>
<tr>
	<td class=campo><%=rs2("curso")%></td>
	<td class=campo align="center"><%=rs2("tot1")%></td>
	<td class=campo align="center"><%=rs2("tot2")%></td>
</tr>
<%
rs2.movenext
loop
rs2.close
%>
</table>
</div>
	<br>
    <p align="justify">	Lembramos ainda que V.Sa. deverá se manifestar por escrito, da aceitação ou não da redução da carga horária
	no prazo máximo de 3 (três) dias a partir do recebimento desta.
	<br>
	A ausência de manifestação d<%=t1%> professor<%=t2%> caracterizará a sua não aceitação e levará a instituição
	a tomar as providências que a legislação determinar.
	<br>
	<br>
	<br>
    </td></tr>


<tr><td>&nbsp;</td></tr>
<tr><td>Atenciosamente,    </td></tr>

<tr>
	<td>
	<p>&nbsp;
	<p>______________________________________________<br>
	Fundação Instituto de Ensino para Osasco
	<br>
	</td>
</tr>
<tr>
	<td class=campo>
	<br>
	<div align="center">
	<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="580">
	<tr><td class="campop">
	Este espaço é para sua aceitação ou não.
	<br><br><br><br><br>
	<%=rs("nome")%>
	</td></tr></table>
	</div

	</td>
</tr>
</table>
</div>
<%
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
rs.movenext
loop
%>
</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing

conexao.close
set conexao=nothing
%>