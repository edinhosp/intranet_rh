<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a91")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Grade Horária</title>
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
<script language="JavaScript" type="text/javascript"><!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome3() {	form.chapa1.value=form.nome1.value;
form.submit();	}
function chapa3() {	form.nome1.value=form.chapa1.value;
form.submit();	}
function nome4() {	form.chapa2.value=form.nome2.value;
form.submit();	}
function chapa4() {	form.nome2.value=form.chapa2.value;
form.submit();	}
function diasem1() {
	ok=0
	dia=form.diasem.value
	diaant=form.diasemant.value
	if (diaant==0) {ok=1}
	if (diaant==7 && (dia>0 && dia<7)) {ok=1}
	if ((diaant>0 && diaant<7) && dia==7) {ok=1}
	if (ok==1) {form.submit()}
}
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

<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
limitech=26
dim conexao, chapach, rs, rs2, a(6)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
'response.write request.form
	if request.form("bt_salvar")<>"" then
	'if request.form("salvar")="1" then
	tudook=1
		sqltemp="SELECT g.CURSO FROM g2cursoeve g WHERE g.coddoc='" & request.form("codcur") & "' "
		rsc.open sqltemp, ,adOpenStatic, adLockReadOnly
		if rsc.recordcount>0 then curso=rsc("curso") else curso=""
		rsc.close
		sqltemp="select materia from grd_materias where codmat='" & request.form("disciplina") & "'"
		'response.write sqltemp
		rsc.open sqltemp, ,adOpenStatic, adLockReadOnly
		if rsc.recordcount=1 then materia=rsc("materia") else materia=""
		rsc.close
		perlet=left(request.form("perlet"),6)
		perlet2=right(request.form("perlet"),6)
		sqltemp="select perletsg from grades_per where coddoc='" & request.form("codcur2") & "' and perlet='" & perlet & "' "
		rsc.open sqltemp, ,adOpenStatic, adLockReadOnly
		if rsc.recordcount=1 then perletsg=rsc("perletsg") else perletsg=perlet
		rsc.close

if request.form("a1")="" and request.form("a2")="" and request.form("a3")="" and request.form("a4")="" and request.form("a5")="" and request.form("a6")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione os horários de aula!');</script>"
end if
	
		sql="UPDATE grades_5 SET "
		sql=sql & "perlet    ='" & perlet & "', "
		sql=sql & "perlet2   ='" & perlet2 & "', "
		sql=sql & "perletsg  ='" & perletsg & "', "
		sql=sql & "coddoc    ='" & request.form("codcur") & "', "
		sql=sql & "curso     ='" & curso & "', "
		sql=sql & "turno     ='" & request.form("turno") & "', "
		sql=sql & "serie     ='" & request.form("serie") & "', "
		sql=sql & "turma     ='" & request.form("turma") & "', "
		codtur=request.form("grupo")
		if request.form("turno")="1" then ct1="M"
		if request.form("turno")="2" then ct1="V"
		if request.form("turno")="3" then ct1="N"
		if request.form("turno")="5" then ct1="V"
		if request.form("turno")="31" then ct1="N"
		if request.form("turno")="61" then ct1="I"
		if request.form("turno")="62" then ct1="I"
		if request.form("turno")="51" then ct1="N"
		if request.form("turno")="52" then ct1="N"
		if request.form("turno")="53" then ct1="N"
		codtur=codtur & ct1 & request.form("turma") & request.form("serie")
		sql=sql & "codtur    ='" & codtur & "', "
		sql=sql & "diasem    = " & request.form("diasem") & ", "
		sql=sql & "enfase    ='" & request.form("enfase") & "', "
		'if request.form("horini")="" then sql=sql & "horini=null," else sql=sql & " horini=#" & request.form("horini") & "#, "
		'horfim=cdate(request.form("horini")) + cdate("00:50")
		'if horfim="" then sql=sql & "horfim=null," else sql=sql & " horfim=#" & horfim & "#, "
 		'sql=sql & " thor=#00:50#, "
		if request.form("a1")="1" then a1=1 else a1="null"
		if request.form("a2")="1" then a2=1 else a2="null"
		if request.form("a3")="1" then a3=1 else a3="null"
		if request.form("a4")="1" then a4=1 else a4="null"
		if request.form("a5")="1" then a5=1 else a5="null"
		if request.form("a6")="1" then a6=1 else a6="null"
		sql=sql & "a1 = " & a1 & ", "
		sql=sql & "a2 = " & a2 & ", "
		sql=sql & "a3 = " & a3 & ", "
		sql=sql & "a4 = " & a4 & ", "
		sql=sql & "a5 = " & a5 & ", "
		sql=sql & "a6 = " & a6 & ", "
		adnot=0
		if request.form("turno")="3" then
			'if request.form("horini")="22:10" then
			if request.form("a6")="1" then adnot=1
			if request.form("diasem")="7" then adnot=0
		end if
		if request.form("inicio")="" then sql=sql & "inicio=null," else sql=sql & "inicio='" & dtaccess(request.form("inicio")) & "', "
		if request.form("termino")="" then sql=sql & "termino=null," else sql=sql & "termino='" & dtaccess(request.form("termino")) & "', "
		if request.form("juntar") ="ON" then sql=sql & "juntar =1, " else sql=sql & "juntar =0, "
		sql=sql & "jturma    = '" & request.form("jturma") & "', "
		if request.form("dividir")="ON" then sql=sql & "dividir=1, " else sql=sql & "dividir=0, "
		if request.form("extra")  ="ON" then sql=sql & "extra  =1, " else sql=sql & "extra  =0, "
		if request.form("demons") ="ON" then sql=sql & "demons =1, " else sql=sql & "demons =0, "
		sql=sql & "adnot   = " & adnot & ", "
		sql=sql & "codmat  ='" & request.form("disciplina") & "', "
		sql=sql & "materia ='" & materia & "', "
		if session("usuariomaster")<>"02379" then
 			sql=sql & "usuarioa='" & session("usuariomaster") & "', "
			sql=sql & "dataa   =now(), "
		end if
		sql=sql & "codsala ='" & request.form("sala") & "', "
		sql=sql & "chapa1  ='" & request.form("chapa1") & "' "
		if request.form("chapa2")="" or request.form("chapa2")="0" then sql=sql & ", chapa2= null "
		if request.form("chapa2")<>"" and request.form("chapa2")<>"0" then sql=sql & ", chapa2='" & request.form("chapa2") & "' "
		if request.form("necessita")="ON" then
			sql=sql & ", prof2=1"
			sql=sql & ", alunos=" & request.form("alunos") & " "
			sql=sql & ", justificativa='" & request.form("justificativa") & "' "
			if request.form("autorizado")="ON" then aut1=1 else aut1=0
			sql=sql & ", autorizado=" & aut1 & " "
			if request.form("quando")<>"" then aut2="'" & dtaccess(request.form("quando")) & "'" else aut2="null "
			sql=sql & ", quando=" & aut2
		else
			sql=sql & ", prof2=0 "
			sql=sql & ", alunos=0 "
			sql=sql & ", justificativa=null "
			sql=sql & ", autorizado=0 "
			sql=sql & ", quando=null "
		end if
		sql=sql & "WHERE id_grade=" & session("id_alt_grade")
		response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	'end if
	end if
	
else 'request.form=""
end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		'sql="DELETE id_grade FROM grades WHERE id_grade=" & session("id_alt_grade")
		sql="UPDATE grades_5 set deletada=1, usuarioa='" & session("usuariomaster") & "', dataa=now() WHERE id_grade=" & session("id_alt_grade")
		if session("usuariomaster")="02379" then sql="UPDATE grades_5 set deletada=-1 WHERE id_grade=" & session("id_alt_grade")
		conexao.Execute sql, , adCmdText
	end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and request.form("salvar")="0") then
	if request("codigo")="" then
		id_grade=session("id_alt_grade")
	else
		id_grade=request("codigo")
	end if
	sqla="select * from grades_5 "
	sqlb="where id_grade=" & id_grade
	sql1=sqla & sqlb 
	'response.write sql1
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if
%>

<%
if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("id_alt_grade")=rs("id_grade")

if request.form("codcur")<>"" then codcur=request.form("codcur") else codcur=rs("coddoc")
'if request.form("perlet")<>"" then perlet=request.form("perlet") else perlet=rs("perlet")

if request.form("perlet")<>"" then
	rperlet=request.form("perlet")
	perlet=left(request.form("perlet"),6)
	perlet2=right(request.form("perlet"),6)
	tipopl=mid(perlet2,5,1):tipople="..."
else
	perlet=rs("perlet")
	perlet2=rs("perlet2")
	tipopl=mid(perlet2,5,1)
	rperlet=perlet & perlet2
end if
	if tipopl="A" then tipople="Anual"
	if tipopl="S" then tipople="Semestral"
if request.form("turno")<>"" then turno=request.form("turno") else turno=rs("turno")
if request.form("grupo")<>"" then grupo=request.form("grupo") else grupo=left(rs("codtur"),3)
if request.form("serie")<>"" then serie=request.form("serie") else serie=rs("serie")
if request.form("turma")<>"" then turma=request.form("turma") else turma=rs("turma")
if request.form("sala")<>"" then sala=request.form("sala") else sala=rs("codsala")
if request.form("codcur")="10" then filial=1 else filial=3
'if request.form("horini")<>"" then horini=request.form("horini") else horini=rs("horini")
if request.form("disciplina")<>"" then codmat=request.form("disciplina") else codmat=rs("codmat")
if request.form("chapa1")<>"" then chapa1=request.form("chapa1") else chapa1=rs("chapa1")
if request.form("chapa2")<>"" then chapa2=request.form("chapa2") else chapa2=rs("chapa2")
if request.form("inicio")<>"" then inicio=request.form("inicio") else inicio=rs("inicio")
if request.form("termino")<>"" then termino=request.form("termino") else termino=rs("termino")
if request.form("codcur")="10" then filial=1 else filial=3
if rs("coddoc")="DIR" then filial=1 else filial=3

if request.form("alunos")=""        then alunos=rs("alunos") else alunos=request.form("alunos")
if request.form("justificativa")="" then justificativa=rs("justificativa") else justificativa=request.form("justificativa")
if request.form("quando")=""        then quando=rs("quando") else quando=request.form("quando")
necessita=0
if rs("prof2")=1 and request.form("necessita")="" then necessita=1
if request.form("necessita")="ON" then necessita=1
if request.form<>"" and request.form("necessita")="" then necessita=0
autorizado=""
if rs("autorizado")=1 and request.form("autorizado")="" then autorizado="checked"
if request.form("autorizado")="ON" then autorizado="checked"
if request.form("jturma")<>"" then jturma=request.form("jturma") else jturma=rs("jturma")
%>
<form method="POST" action="grade_alteracao.asp" name="form">
<input type="hidden" name="id_grade" size="4" value="<%=rs("id_grade")%>">  
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="100%">
	<tr><td class=grupo>Alteração de Grade Horária</td></tr>
</table>

<!-- Periodo Letivo / Curso -->
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo>Curso</td>
	<td class=titulo>Período Letivo</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="codcur" onfocus="javascript:window.status='Selecione o curso'" onChange="javascript:submit()">
<%
if session("usuariomaster")="02379" then lanc=" in (0,1) " else lanc=" in (1) "
sqla="SELECT coddoc, curso as nome FROM grades_per where tper='P' and lanc" & lanc & _
"GROUP BY coddoc, curso ORDER BY curso "
'"and coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') " & _
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
response.write "<option value='0'>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if codcur=rsd("coddoc") then tempc="selected" else tempc=""
%>
	<option value="<%=rsd("coddoc")%>" <%=tempc%>><%=rsd("nome")%> (<%=rsd("coddoc")%>)</option>
<%
rsd.movenext:loop
else
%>
	<option value="-1">Sem cursos cadastrados</option>
<%
end if
rsd.close
%>
	</select></td>
	<td class=fundo><select size="1" name="perlet" onfocus="javascript:window.status='Selecione o período'" onChange="javascript:submit()">
<%
sqla="SELECT perlet, perlet2 FROM grades_per where lanc " & lanc & " and coddoc='" & codcur & "' group by perlet, perlet2 "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
response.write "<option value='0'>....</option>"
rsd.movefirst:do while not rsd.eof
if rperlet=rsd("perlet") & rsd("perlet2") then temppl="selected" else temppl=""
%>
	<option value="<%=rsd("perlet") & rsd("perlet2")%>" <%=temppl%>><%=rsd("perlet")%></option>
<%
rsd.movenext:loop
else
%>
	<option value="-1">Sem períodos cadastrados</option>
<%
end if
rsd.close
%>
	</select>&nbsp; Grade <%=tipople%></td>
</tr>
</table>

<!-- Periodo / Serie / dia da semana -->
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo>Período</td>
	<td class=titulo>Série/Turma</td>
	<td class=titulo>Dia</td>
	<td class=titulo>Sala</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="turno" onfocus="javascript:window.status='Selecione o período'" onChange="javascript:submit()">
<%
sqla="SELECT codturno, descturno from eturnos where codturno in (1,2,3,5,31,51,52,53) "
if session("usuariomaster")="02379" then sqla="SELECT codturno, descturno from eturnos where codturno in (1,2,3,5,31,51,52,53) "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
response.write "<option value='0'>....</option>"
rsd.movefirst:do while not rsd.eof
if cint(turno)=cint(rsd("codturno")) then tempt="selected" else tempt=""
%>
	<option value="<%=rsd("codturno")%>" <%=tempt%>><%=rsd("descturno")%></option>
<%
rsd.movenext:loop
else
%>
	<option value="-1">Sem períodos cadastrados</option>
<%
end if
rsd.close
%>
	</select></td>
	
	<td class=fundo><select size="1" name="serie" onfocus="javascript:window.status='Selecione a serie/semestre'" onChange="javascript:submit()">
<%
sqla="SELECT serie from grades_gc where coddoc='" & codcur & "' and perlet='" & perlet & "' order by serie "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
response.write "<option value='0'>..</option>"
rsd.movefirst:do while not rsd.eof
if cint(serie)=cint(rsd("serie")) then temps="selected" else temps=""
%>
	<option value="<%=rsd("serie")%>" <%=temps%>><%=rsd("serie")%></option>
<%
rsd.movenext:loop
else
%>
	<option value="-1">Sem séries</option>
<%
end if
rsd.close
%>
	</select>
	<select size="1" name="turma" onfocus="javascript:window.status='Selecione a serie/semestre'">
<%
response.write "<option value='0'>....</option>"
for serie2=65 to 90
letra=chr(serie2)
'response.write letra
if turma=letra then temptm="selected" else temptm=""
%>
	<option value="<%=letra%>" <%=temptm%>><%=letra%></option>
<%
next
%>
	</select>
	<select size="1" name="grupo" onfocus="javascript:window.status='Selecione'" onChange="javascript:submit()">
<%
sqla="SELECT grupo from grades_per where coddoc='" & codcur & "' and perlet='" & perlet & "' order by grupo "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
'response.write "<option value='0'>..</option>"
rsd.movefirst:do while not rsd.eof
if grupo=rsd("grupo") then temps="selected" else temps=""
%>
		<option value="<%=rsd("grupo")%>" <%=temps%>><%=rsd("grupo")%></option>
<%
rsd.movenext:loop
else
%>
		<option value="-1">...</option>
<%
end if
rsd.close
%>
	</select>	
<%
sqla="SELECT enfase from grades_per where coddoc='" & codcur & "' and perlet='" & perlet & "' and grupo='" & grupo & "' "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then enfase=rsd("enfase") else enfase="NOR"
rsd.close
%>
<input type="hidden" name="enfase" value="<%=enfase%>">
	
	</td>
<input type="hidden" name="diasemant" value="<%=request.form("diasem")%>">
	<td class=fundo><select size="1" name="diasem" onfocus="javascript:window.status='Selecione o dia da semana'" onchange="diasem1()">
<%
if request.form("diasem")<>"" then diasem=request.form("diasem") else diasem=rs("diasem")
response.write "<option value='0'>..</option>"
for dia=2 to 7
diasem2=weekdayname(dia,-1)
if cint(diasem)=cint(dia) then tempd="selected" else tempd=""
%>
	<option value="<%=dia%>" <%=tempd%>><%=diasem2%></option>
<%
next
%>
	</select></td>

	<td class=fundo><select class=small size="1" name="sala" onfocus="javascript:window.status='Selecione a sala'">
<%
sqla="SELECT CODSALA as sala, SALADESC FROM ESALAS WHERE CODFILIAL=" & filial & " ORDER BY SALADESC "
sqla="select * from ( select codsala, saladesc, salacap, 'cadeiras' as tipo from esalas where codfilial=" & filial & "  UNION ALL select codsala, saladesc & ' (*)', salacap, tipo from grades_esalas) as salas order by saladesc "
sqla="select sala AS codsala, saladesc, salacap, tipo from grades_esalas where codfilial=" & filial & " order by saladesc "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
response.write "<option value=''>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if sala=rsd("codsala") then tempd="selected" else tempd=""
%>
	<option value="<%=rsd("codsala")%>" <%=tempd%>><%=rsd("saladesc")%></option>
<%
rsd.movenext:loop
else
%>
	<option value="-1">Sem cadastro</option>
<%
end if
rsd.close
%>
	</select></td>

</tr>
</table>
  
<!-- Hora Inicio/Termino / Disciplina -->
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo colspan=6>Horário da aula</td>
	<td class=titulo>Disciplina</td>
</tr>
<tr>
<%
if rs("a1")=1 then a(1)="checked" else if request.form("a1")="1" then a(1)="checked" else a(1)=""
if rs("a2")=1 then a(2)="checked" else if request.form("a2")="1" then a(2)="checked" else a(2)=""
if rs("a3")=1 then a(3)="checked" else if request.form("a3")="1" then a(3)="checked" else a(3)=""
if rs("a4")=1 then a(4)="checked" else if request.form("a4")="1" then a(4)="checked" else a(4)=""
if rs("a5")=1 then a(5)="checked" else if request.form("a5")="1" then a(5)="checked" else a(5)=""
if rs("a6")=1 then a(6)="checked" else if request.form("a6")="1" then a(6)="checked" else a(6)=""

sqla="select horini from grd_defhor where codds=" & diasem & " and codtn=" & turno & " order by horini "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
rsd.movefirst:do while not rsd.eof
%>
	<td class=fundo><%=rsd("horini")%><br><input type="checkbox" name="a<%=rsd.absoluteposition%>" value="1" <%=a(rsd.absoluteposition)%>></td>
<%
rsd.movenext:loop
else
%> 
	<td class=fundo colspan=6><font color="red">Selecione um período e/ou dia da semana</td>
<%
end if
rsd.close
%>

	<td class=fundo><select size="1" name="disciplina" onfocus="javascript:window.status='Selecione a disciplina'" onChange="javascript:submit()">
<%
sqla="SELECT c.coddoc, c.perlet, c.GC, c.serie, m.CODMAT, m.MATERIA, m.NAULASSEM " & _
"FROM grades_gc AS c INNER JOIN grades_materias AS m ON (c.serie=m.serie) AND (c.GC=m.GC) AND (c.coddoc=m.coddoc) " & _
"WHERE c.coddoc='" & codcur & "' AND c.perlet='" & perlet & "' AND c.serie=" & serie 
if session("usuariomaster")<>"02379" then sqla=sqla & " AND m.demons=0 "
sqla=sqla & " order by m.materia " 
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
response.write "<option value='0'>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if codmat=rsd("codmat") then tempdi="selected" else tempdi=""
%>
	<option value="<%=rsd("codmat")%>" <%=tempdi%>><%=rsd("materia")%></option>
<%
rsd.movenext:loop
else
%>
	<option value="-1">Sem disciplinas cadastradas</option>
<%
end if
rsd.close
%>
	</select>
	<a class=r href="hstdisciplina.asp?codmat=<%=codmat%>" onclick="NewWindow(this.href,'Pesquisa_disciplinas','545','200','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Magnify.gif" width="16" height="16" border="0" alt=""></a>
	</td>
</tr>
</table>

<!-- Chapa / Nome -->
<input type="hidden" name="chapa1ant" value="<%=request.form("chapa1")%>">
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>CH</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=chapa1%>" name="chapa1" size="8" onfocus="javascript:window.status='Informe o chapa do professor'" onchange="chapa3()"></td>
	<td class=fundo>&nbsp;
		<select size="1" name="nome1" onfocus="javascript:window.status='Selecione o Nome do Professor'" onchange="nome3()">
<%
sql2="select chapa, nome from grades_aux_prof "
if session("dp_chapa")<>"" then sql2=sql2 & "and chapa='" & session("dp_chapa") & "'" else sql2=sql2 & "order by nome"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione Professor 1....</option>"
rsc.movefirst:do while not rsc.eof
if chapa1=rsc("chapa") then temp="selected" else temp=""
%>
		<option value="<%=rsc("chapa")%>" <%=temp%>><%=rsc("nome")%></option>
<%rsc.movenext:loop%>
	</select></td>
	<td class=fundo align="left">&nbsp;
<%
sqlc="select count(chapa1) as taulas from grades_5ch where chapa1='" & chapa1 & "' and perlet2 like '" & left(perlet2,4) & "%" & right(perlet2,1) & "' and deletada=0 and juntar=0 "
rsd.Open sqlc, ,adOpenStatic, adLockReadOnly
taulas1=rsd("taulas")
rsd.close
%>	
<%if taulas1>0 then %>
<a class=r href="hstaulas.asp?chapa=<%=chapa1%>&ano=<%=left(perlet2,4)%>&semestre=<%=right(perlet2,1)%>" onclick="NewWindow(this.href,'Aulas_atribuidas','545','200','yes','center');return false" onfocus="this.blur()">
<%end if%>
<%=taulas1%> aulas
<%if taulas1>0 then %>
</a>
<%end if%>
	</td>	
</tr>
</table>
<%
if necessita=-1 then txtneed="checked" else txtneed=""
%>
<input type="hidden" name="chapa2ant" value="<%=request.form("chapa2")%>">
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr><td class=titulo colspan=3><input type="checkbox" name="necessita" value="ON" <%=txtneed%> onClick="javascript:submit()">
<font color="red">Necessita de 2º Professor?</td></tr>
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>CH</td>
</tr>
<%
'*************** 2 professor *************************
if necessita=-1 then
%>
<tr>
	<td class=fundo><input type="text" value="<%=chapa2%>" name="chapa2" size="8" onfocus="javascript:window.status='Informe o chapa do professor'" onchange="chapa4()"</td>
	<td class=fundo>&nbsp;
		<select size="1" name="nome2" onfocus="javascript:window.status='Selecione o Nome do Professor'" onchange="nome4()" >
<%
response.write "<option value='0'>Selecione Professor 2....</option>"
rsc.movefirst
do while not rsc.eof
if chapa2=rsc("chapa") then temp="selected" else temp=""
%>
			<option value="<%=rsc("chapa")%>" <%=temp%>><%=rsc("nome")%></option>
<%
rsc.movenext
loop
rsc.close
set rsc=nothing
%>
		</select></td>
	<td class=fundo align="left">
<%
sqlc="select count(chapa1) as taulas from grades_5ch where chapa1='" & chapa2 & "' and perlet2 like '" & left(perlet2,4) & "%" & right(perlet2,1) & "' and deletada=0 and juntar=0 "
rsd.Open sqlc, ,adOpenStatic, adLockReadOnly
taulas2=rsd("taulas")
rsd.close
%>	
<%if taulas2>0 then %>
<a class=r href="hstaulas.asp?chapa=<%=chapa2%>&ano=<%=left(perlet2,4)%>&semestre=<%=right(perlet2,1)%>" onclick="NewWindow(this.href,'Aulas_atribuidas','545','200','yes','center');return false" onfocus="this.blur()">
<%end if%>
<%=taulas2%> aulas
<%if taulas2>0 then %>
</a>
<%end if%>
	</td>	
</tr>
<tr>
	<td class=titulo>Alunos</td>
	<td class=titulo>Justificativa</td>
	<td class=titulo>Autorizado</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=alunos%>" name="alunos" size="3" onfocus="javascript:window.status='Informe o numero de alunos da turma'"></td>
	<td class=fundo><textarea name="justificativa" cols="42" rows="3"><%=justificativa%></textarea></td>
<%if session("usuariomaster")="02379" then%>
	<td class=fundo><input type="checkbox" name="autorizado" value="ON" <%=autorizado%>>
	<input type="text" name="quando" size="8" value="<%=quando%>"></td>
<%else%>
	<td class=fundo>&nbsp;</td>
<%end if%>
</tr>
<%
else '**2 prof**
	chapa2="":alunos=0:justificativa="":quando=""
end if
'******************************** 2 professor     **************
%>
</table>

<% if rs("juntar") =0 then juntar ="" else juntar ="checked" %>
<% if rs("dividir")=0 then dividir="" else dividir="checked" %>
<% if rs("extra")  =0 then extra  ="" else extra  ="checked" %>
<% if rs("demons") =0 then demons ="" else demons ="checked" %>
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulor>Início </td>
	<td class=titulor>Término</td>
	<td class=titulor>Junta turmas</td>
	<td class=titulor>Divide turmas</td>
<% if session("usuariomaster")="02379" then %>
	<td class=titulor>Aula Extra</td>
	<td class=titulor>Demonstr.</td>
<% end if %>
</tr>
<tr>
<% if session("usuariomaster")="02379" then %>
	<td class=fundo><input type="text" name="inicio" size="12" value="<%=inicio%>"></td>
	<td class=fundo><input type="text" name="termino" size="12" value="<%=termino%>"></td>
	<td class=fundo><input type="checkbox" name="juntar" value="ON" <%=juntar%>><input type="text" name="jturma" size="5" value="<%=jturma%>"></td>
	<td class=fundo><input type="checkbox" name="dividir" value="ON" <%=dividir%>></td>
	<td class=fundo><input type="checkbox" name="extra" value="ON" <%=extra%>></td>
	<td class=fundo><input type="checkbox" name="demons" value="ON" <%=demons%>></td>
<% else %>
	<td class=fundo><input type="text" name="inicio" size="12" value="<%=inicio%>"></td>
	<td class=fundo><input type="text" name="termino" size="12" value="<%=termino%>"></td>
    <td class=fundor><%if rs("juntar")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
    <td class=fundor><%if rs("dividir")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
<!--	<td class=fundo><input type="checkbox" name="juntar" value="ON"></td>
	<td class=fundo><input type="checkbox" name="dividir" value="ON"></td> -->
<% end if %>
</tr>
</table>
  
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo align="center">
  	<% if (taulas1>limitech and chapa1<>"99999") or (taulas2>limitech and chapa2<>"99999") then %>
	<font color=red>Professor excede o limite de 20 aulas!</font>
		<% if session("usuariomaster")="02379" then %>
			<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
		<% end if %>

	<% else %>
		<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<% end if%>
	<td class=titulo align="center"><input type="reset"  value="Desfazer" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser alterado!');</script>"
	end if
	'Response.write "<p>Registro atualizado.<br>"
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<!--
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>