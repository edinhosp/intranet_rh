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
<title>Cartas de autorização para professores convidados</title>
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
	temp=1
	sql="select *, status=case when autorizado=1 then 'Reemissão' else 'Nova' end from ( " & _
	"select g.id_grdaula, g.perlet, g.codtur, c.curso, g.codmat, m.materia, g.chapa1, f.nome, pini=min(data), pfim=max(data), inicio, termino, autorizado, quando, ch=sum(aulas) " & _
	"from (((g5ch g inner join g2cursos c on c.coddoc=g.coddoc and c.codcur=g.codcur and c.codper=g.codper) " & _
	"inner join corporerm.dbo.umaterias m on m.codmat collate database_default=g.codmat) " & _
	"inner join grades_aux_prof f on f.chapa=g.chapa1) " & _
	"inner join g5datas d on d.id_grdaula=g.id_grdaula " & _
	"where g.deletada=0 and (g.chapa1>'10000' and g.chapa1<'99000') " & _
	"group by g.id_grdaula, g.perlet, g.codtur, c.curso, g.codmat, m.materia, g.chapa1, f.nome, inicio, termino, autorizado, quando ) z " & _
	"where ((autorizado=0 or autorizado is null) or (pini<>inicio or pfim<>termino)) " & _
	" and perlet  IN ('2016/5','2016/6') " & _
	"Order by curso, nome  "
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
	"where g.deletada=0 and (g.chapa1>'10000' and g.chapa1<'99000') " & _
	"group by g.id_grdaula, g.perlet, g.codtur, c.curso, g.codmat, m.materia, g.chapa1, f.nome, inicio, termino, autorizado, quando ) z " & _
	"where id_grdaula=" & request("codigo")
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if temp=1 then
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td class=grupo>Emissão de autorização para professores convidados (<%=rs.recordcount%>)</td></tr>
</table>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=titulo align="center">#</td>
	<td class=titulo align="center">Nome</td>
	<td class=titulo align="center">Curso</td>
	<td class=titulo align="center">Disciplina</td>
	<td class=titulo align="center">Período</td>
	<td class=titulo align="center">Aulas</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo style="border-bottom:2px dotted"><%=rs("chapa1")%></td>
	<td class=campo style="border-bottom:2px dotted"><a href="pos_cartaautorizacao.asp?codigo=<%=rs("id_grdaula")%>"><%=rs("nome")%></a></td>
	<td class=campo style="border-bottom:2px dotted"><%=rs("curso") %></td>
	<td class=campo style="border-bottom:2px dotted"><%=rs("materia") %></td>
	<td class=campo style="border-bottom:2px dotted"><%=rs("pini") & " a " & rs("pfim") %></td>
	<td class=campo style="border-bottom:2px dotted"><%=rs("ch") %></td>
</tr>
<%
rs.movenext
loop
rs.close
else
	response.write "<tr><td class=""campop"" colspan=6 style=""border-bottom:2px dotted"">Sem autorizações a emitir.</td></tr>"
end if
%>
</table>
<%
else ' temp=0

sql2="select sexo from grades_novos where chapa='" & rs("chapa1") & "' " & _
"union all " & _
"select sexo collate database_default from dc_professor where chapa='" & rs("chapa1") & "' "
'RESPONSE.WRITE SQL2
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
sexo=rs2("sexo")
rs2.close
if sexo="M" then v1="o" else v1="a"
if sexo="M" then v2="" else v2="a"

sqlt="select grauinstrucao from grades_novos where chapa='" & rs("chapa1") & "' union all " & _
"select grauinstrucao collate database_default from dc_professor where chapa='" & rs("chapa1") & "' "
rs2.Open sqlt, ,adOpenStatic, adLockReadOnly
grauinstrucao=rs2("grauinstrucao")
rs2.close
sqlt="select descricao from corporerm.dbo.pcodinstrucao where codcliente='" & grauinstrucao & "' "
rs2.Open sqlt, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then titulacao=rs2("descricao") else titulacao="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
rs2.close
%>
<div align="center">
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</td></tr>
	<tr><td class="campop"><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width=225></td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><b>Of. <%=rs("chapa1")%> - Secretaria de Cursos</b></td></tr>
	<tr><td class="campop" align="right">
	<input type="text" name="txt1" class="form_input" size="29" value="Osasco, <%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %>" style="font-size:10pt">
	</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><input type="text" name="txt0" class="form_input" size="5" value="Exmo Sr." style="font-size:10pt"><br>
	<input type="text" name="txt1" class="form_input" size="60" value="Dr. Luiz Fernando da Costa e Silva" style="font-size:10pt"><br>
	<input type="text" name="txt2" class="form_input" size="60" value="Reitoria do UNIFIEO" style="font-size:10pt"><br>
	</td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Venho por meio desta solicitar de Vossa Senhoria autorização para que <%=v1%> Sr<%=v2%>. <b><%=textopuro(rs("nome"),1)%></b> 
	(<%=titulacao%>) ministre como convidad<%=v1%>, a disciplina de <b><%=rs("materia")%></b> no curso de Pós-Graduação de <b><%=rs("curso")%></b>. 
	A carga horária total da disciplina ministrada será de <b><%=rs("ch")%></b> h/a, no período de <b><%=rs("pini")%></b> a <b><%=rs("pfim")%></b>.
	</td></tr>
	<tr><td class="campop"><p align="justify" style="line-height: 25px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Esperando contar com a sua compreensão,
	desde já agradeço e reitero meus protestos de estima e consideração.</td></tr>

	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop">&nbsp;<br>&nbsp;<br>&nbsp;<br></td></tr>
	<tr><td class="campop"><p align="center" style="line-height: 25px">

	<input type="text" name="txt1" class="form_input" size="60" value="Prof. " style="font-size:10pt;font-style:bold;text-align:center"><br>
	<input type="text" name="txt2" class="form_input" size="60" value="Coordenador do curso de <%=rs("curso")%>" style="font-size:10pt;text-align:center"><br>

	</td></tr>
	<tr><td class="campop"></td></tr>
</table>
</div>


<%
sqlu="update g2aulas set inicio=ini, termino=fim " & _
"from g2aulas a inner join " & _
"(select id_grdaula, min(data) ini, max(data) fim from g2aulasdata where deletada=0 group by id_grdaula) d on a.id_grdaula=d.id_grdaula " & _
"where a.id_grdaula=" & rs("id_grdaula")
'response.write sqlu
conexao.execute sqlu

rs.close

end if 'temp=0

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>