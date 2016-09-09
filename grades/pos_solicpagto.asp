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
<title>Solicitação de Pagamento para Convidados</title>
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
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
set rs4=server.createobject ("ADODB.Recordset")
Set rs4.ActiveConnection = conexao

if request.form("b1")="" then
%>
<p class=titulo>Solicitação de Pagamento - Convidados&nbsp;<%=titulo %>
<form method="POST" action="pos_solicpagto.asp" name="form">
<table border="0" width="350" cellspacing="0"cellpadding="3">
<tr>
	<td class=titulo>Nome do Convidado</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="chapa" onChange="javascript:submit()">
&nbsp;
<%
chapae=request.form("chapa")
sql2="select distinct p.chapa1, f.nome from g5listapagto p left join grades_aux_prof f on f.chapa=p.chapa1 where chapa1>'10000' and chapa1<'99999' order by f.nome"
rs.Open sql2, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	response.write "<option>Selecione...</option>"
	rs.movefirst
do while not rs.eof
	if chapae=rs("chapa1") then temp1="selected" else temp1=""
%>
	<option value="<%=rs("chapa1")%>" <%=temp1%>><%=rs("nome")%>&nbsp;&nbsp;&nbsp; </option>
<%
rs.movenext:loop
else
	response.write "<option>Não existe nenhum disponivel para pagamento</option>"
end if
rs.close
%>
	</select>
</td>
</tr>
</table>

<table border="0" width="350" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo><input type="submit" value="Visualizar" class="button" name="B1"></td>
</tr>
</table>

<%
vezes=0
if request.form<>"" then

sql1="select id_data, id_grdaula, data, aulas, id_grdturma, coddoc, codtur, codmat, materia, chapa1 " & _
"from g5listapagto where chapa1='" & request.form("chapa") & "' " & _
"order by data, coddoc"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<br>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Data</td>
	<td class=titulo>Disciplina</td>
	<td class=titulo>Turma</td>
	<td class=titulo># Aulas</td>
	<td class=titulo></td>
</tr>
<%
rs.movefirst
do while not rs.eof
temp=rs.absoluteposition mod 2
if temp=0 then classe="campol" else classe="campo"
%>
<tr>
	<td class=<%=classe%>><%=rs("data")%></td>
	<td class=<%=classe%>><%=rs("materia")%></td>
	<td class=<%=classe%>><%=rs("codtur")%></td>
	<td class=<%=classe%> align="center"><%=rs("aulas")%></td>
	<td class=<%=classe%>>
		<input type="checkbox" name="emitir<%=vezes%>" value="ON" <%=""%> >
		<input type="hidden" name="data<%=vezes%>" value="<%=rs("id_data")%>">
		<input type="hidden" name="aula<%=vezes%>" value="<%=rs("aulas")%>">
	</td>
</tr>
<%
vezes=vezes+1
rs.movenext
loop
session("credautonomo")=vezes-1
end if
rs.close
%>
</table>
<%
end if
%>
</form>
<%
else ' request.form<>""
	vez=session("credautonomo")
	chapa=request.form("chapa")
	sql="delete from ttpos_solpag where sessao='" & session("usuariomaster") & "'": conexao.Execute sql, , adCmdText
	for a=0 to vez
		id_data=request.form("data" & a)
		aulas=request.form("aula" & a)
		emitir=request.form("emitir" & a)
		'response.write "<br>" & chapa & " " & id_data & " " & emitir
		if emitir="ON" then
			sql="INSERT INTO ttpos_solpag ( sessao, id_data, aulas ) SELECT '" & session("usuariomaster") & "', " & id_data & ", " & aulas & ""
			conexao.Execute sql, , adCmdText 
		end if
	next
'parei

sql0="select nome"

sql1="select a.* from ttpos_solpag t inner join g5listapagto a on a.id_data=t.id_data where sessao='" & session("usuariomaster") & "'"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:primeira=rs("id_data"):primeira_data=rs("data")
rs.movelast:ultima=rs("id_data")
total=rs.recordcount
rs.movefirst

%>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="620" height="990">
<tr><td colspan=6 class=titulop height=35 valign="middle" align="center" style="border-bottom:2 solid"> S O L I C I T A Ç Ã O &nbsp; &nbsp; D E  &nbsp; &nbsp; P A G A M E N T O
</td></tr>
<!-- corpo da carta -->
<%
%>
<tr><td class=fundop colspan=3 align="center"> O R I G E M </td><td class=fundop colspan=3 align="center"> D E S T I N O </td></tr>
<tr>
	<td class=campo height=45><b>DEPTO/SEÇÃO:<br><input type="text" class="form_input" value="Secretaria Pós-Graduação" size=24></td>
	<td class=campo><b>DATA:<br><input type="text" class="form_input10" value="<%=int(now())%>" size=8></td>
	<td class=campo><b>NÚMERO:<br><input type="text" class="form_input10" value="" size=5></td>

	<td class=campo><b>DEPTO/SEÇÃO:<br><input type="text" class="form_input10" value="Contas a Pagar" size=12></td>
	<td class=campo><b>A ATENÇÃO DE:<br><input type="text" class="form_input10" value="Sr. Nascimento" size=13></td>
	<td class=campo><b>RECEBIDO EM:<br><input type="text" class="form_input10" value="" size=8></td>
</tr>
<tr>
	<td class="campop" colspan=6 height=50 style="border-bottom:2 solid">
	<b>ASSUNTO:</b><br>Solicitação de Pagamento para professor convidado</td>
</tr>
	
	
<tr><td colspan=6 height=100% class="campop" align="left" valign=top>

<p align="left" style="margin-top:0;margin-bottom:0;font-size:12pt">
<p align="justify" style="margin-top:0;margin-bottom:0;font-size:11pt;text-indent:10pt;line-height:150%">
<br>
<%
'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
'next
'response.write "</tr>"
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext:loop
'response.write "</table>"
'response.write "# " & rs.recordcount & "<br>"
'*************** fim teste **********************

sql2="select * from ( " & _
"select chapa, nome=nome collate database_default, sexo=sexo collate database_default, grauinstrucao=grauinstrucao collate database_default, cpf=cpf collate database_Default, " & _
"pispasep=pispasep collate database_default, cartidentidade=cartidentidade collate database_Default from dc_professor " & _
"union all " & _
"select chapa, nome, sexo, grauinstrucao, cpf, pispasep, cartidentidade from grades_novos " & _
") z where chapa='" & chapa & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
nome=rs2("nome")
grauinstrucao=rs2("grauinstrucao"):cpf=rs2("cpf"):pis=rs2("pispasep"):rg=rs2("cartidentidade")
rs2.close

sql3="select sum(aulas) total_aulas from (" & sql1 & ") z" 'ttpos_solpag where sessao='" & session("usuariomaster") & "'"
rs2.Open sql3, ,adOpenStatic, adLockReadOnly
total_aulas=rs2("total_aulas")
rs2.close
t=grauinstrucao

if chapa="98218" or chapa="98388" then
if t<="B" then          sqlv="select valor from corporerm.dbo.pvalfix where codigo='POSA' and '" & dtaccess(primeira_data) & "' between iniciovigencia and finalvigencia ":titulo="Especialista"
if t>"B" and t<"E" then sqlv="select valor from corporerm.dbo.pvalfix where codigo='POSB' and '" & dtaccess(primeira_Data) & "' between iniciovigencia and finalvigencia ":titulo="Mestre"
if t>="E" then          sqlv="select valor from corporerm.dbo.pvalfix where codigo='POSC' and '" & dtaccess(primeira_data) & "' between iniciovigencia and finalvigencia ":titulo="Doutor"
else
if t<="B" then          sqlv="select valor from corporerm.dbo.pvalfix where codigo='POSD' and '" & dtaccess(primeira_data) & "' between iniciovigencia and finalvigencia ":titulo="Especialista"
if t>"B" and t<"E" then sqlv="select valor from corporerm.dbo.pvalfix where codigo='POSE' and '" & dtaccess(primeira_Data) & "' between iniciovigencia and finalvigencia ":titulo="Mestre"
if t>="E" then          sqlv="select valor from corporerm.dbo.pvalfix where codigo='POSF' and '" & dtaccess(primeira_data) & "' between iniciovigencia and finalvigencia ":titulo="Doutor"
end if

rs2.Open sqlv, ,adOpenStatic, adLockReadOnly
valor=rs2("valor")
if t="9" then valor=40.36:titulo="Graduado"
valor=cdbl(valor)
valor1=valor*0.05
valor2=(valor+valor1)*0.1667
valort=round(valor+valor1+valor2,2)
rs2.close
'98218 98388
%>

<div align="center">
<table border='0' bordercolor="#000000" cellpadding='2' cellspacing='0' style='border-collapse:collapse' width=570>
<tr>
	<td class="campor" style="border-left:1px solid;border-top:1px solid;border-right:1px solid">Nome do prestador de serviço</td>
	<td class="campor" style="border-left:1px solid;border-top:1px solid;border-right:1px solid">Titulação</td>
</tr>
<tr>
	<td class="campop" style="border-left:1px solid;border-bottom:1px solid;border-right:1px solid"><%=nome%></td>
	<td class="campop" style="border-left:1px solid;border-bottom:1px solid;border-right:1px solid"><%=grauinstrucao & " - " & titulo%> </td>
</tr>
</table>
<table border='0' bordercolor="#000000" cellpadding='2' cellspacing='0' style='border-collapse:collapse' width=570>
<tr>
	<td class="campor" style="border-left:1px solid;border-top:1px solid;border-right:1px solid">CPF</td>
	<td class="campor" style="border-left:1px solid;border-top:1px solid;border-right:1px solid">PIS/PASEP/NIT</td>
	<td class="campor" style="border-left:1px solid;border-top:1px solid;border-right:1px solid">Outras informações</td>
</tr>
<tr>
	<td class="campop" style="border-left:1px solid;border-bottom:1px solid;border-right:1px solid"><%=cpf%></td>
	<td class="campop" style="border-left:1px solid;border-bottom:1px solid;border-right:1px solid"><%=pis%></td>
	<td class="campop" style="border-left:1px solid;border-bottom:1px solid;border-right:1px solid"><%=total_aulas%> (Valor Base: <%=valor%>)</td>
</tr>
</table>
</div>

<br>

<div align="center">
<table border='1' bordercolor="#000000" cellpadding='2' cellspacing='0' style='border-collapse:collapse' width=590>
<tr>
	<td class="campop" colspan=5>Cursos</td></tr>
<%
sqlc="select distinct y.coddoc, curso, codccusto from ( " & sql1 & ") y inner join g2cursoeve c on c.coddoc=y.coddoc "
rs2.Open sqlc, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>
	<tr>
		<td class=fundo colspan=5><b><%=rs2("curso")%></b> (<%=rs2("codccusto")%>)</td>
	</tr>
	<tr>
		<td class=campo colspan=2 ><b>Disciplina</td>
		<td class=campo align="center"><b>Datas</td>
		<td class=campo align="center"><b>Nº aulas</td>
		<td class=campo align="center"><b>Valor R$</td>
	</tr>
<%
	sqlm="select codmat, materia, sum(aulas) aulas from (" & sql1 & ") y  where coddoc='" & rs2("coddoc") & "' group by codmat, materia"
	rs3.Open sqlm, ,adOpenStatic, adLockReadOnly
	do while not rs3.eof
%>
	<tr>
		<td class=campo>&nbsp;</td>
		<td class=campo><%=rs3("materia")%></td>
		<td class=campo>
<%
	sqld="select data from (" & sql1 & ") y where coddoc='" & rs2("coddoc") & "' and codmat='" & rs3("codmat") & "' order by data"
	rs4.Open sqld, ,adOpenStatic, adLockReadOnly
	do while not rs4.eof
	response.write rs4("data")
	if rs4.absoluteposition<rs4.recordcount and rs4.recordcount>1 then response.write ", "
	rs4.movenext
	loop
	rs4.close
%>
		</td>
		<td class=campo align="center"><%=rs3("aulas")%></td>
		<td class=campo align="right"><%=formatnumber(rs3("aulas")*valort,2)%></td>
	</tr>
<%
	totalpag=totalpag+(rs3("aulas")*valort)
	rs3.movenext
	loop
	rs3.close	
rs2.movenext
loop
rs2.close
%>
	<tr>
		<td class=fundo align="left" colspan=4><b>Total Bruto a receber</td>
		<td class="campop" align="right"><b><%=formatnumber(totalpag,2)%></td>
	</tr>
</table>
</div>

	
</td></tr>
<!-- final do corpo da carta -->

<!-- rodapé da carta -->
<tr><td height=100 colspan=6 class="campor" style="border-top:2px dotted">

<Br>
<div align="center">
<table border='0' bordercolor="#000000" cellpadding='2' cellspacing='0' style='border-collapse:collapse' width=570>
<tr><td class="campop" colspan=3 style="border:1px solid">AUTORIZAÇÕES</td></tr>
<tr>
	<td class="campor" width=30% style="border-left:1px solid;border-top:1px solid;border-right:1px solid">Coordenador</td>
	<td class="campor" width=35% style="border-left:1px solid;border-top:1px solid;border-right:1px solid">Pró-Reitoria Acadêmica</td>
	<td class="campor" width=35% style="border-left:1px solid;border-top:1px solid;border-right:1px solid">Pró-Reitoria Administrativa</td>
</tr>
<tr>
	<td class="campop" style="border-left:1px solid;border-bottom:1px solid;border-right:1px solid">&nbsp;<br><br><br></td>
	<td class="campop" style="border-left:1px solid;border-bottom:1px solid;border-right:1px solid">&nbsp;</td>
	<td class="campop" style="border-left:1px solid;border-bottom:1px solid;border-right:1px solid">&nbsp;</td>
</tr>
</table>
</div>
<br>

</td></tr>

<tr><td height=30 colspan=6 class="campor"><%=session("usuariomaster")%> - Autenticador: <%=int(totalpag*day(now)*month(now)*year(now))%></td></tr>
</table>
<!-- </div> -->

<%
else
%>
	

<%
end if ' rs.recordcount>0
rs.close

%>

<%
end if 
%>
</body>
</html>
<%
set rs=nothing
set rs2=nothing
set rs3=nothing
set rs4=nothing
conexao.close
set conexao=nothing
%>