<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a90")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Grade Horária</title>
<script language="javascript" type="text/javascript"><!--
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

dim conexao, conexao2, chapach, rs, rs2, ok
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		'if request.form("salvar")="1" then
		tudook=1
			sqla = "INSERT INTO grades_3 (perlet, perlet2, perletsg, coddoc, curso, turno, serie, turma, diasem, "
			sqla = sqla & "a1,a2,a3,a4,a5,a6, inicio, termino, juntar, dividir, extra, demons, adnot, codmat, materia, "
			sqla = sqla & "codsala, usuarioc, datac, codtur, chapa1 "
			sqla = sqla & " )"

if request.form("a1")="" and request.form("a2")="" and request.form("a3")="" and request.form("a4")="" and request.form("a5")="" and request.form("a6")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione os horários de aula!');</script>"
end if
			
			sqltemp="SELECT g.CURSO FROM g2cursoeve g WHERE g.coddoc='" & request.form("codcur") & "'"
			rsc.open sqltemp, ,adOpenStatic, adLockReadOnly
			if rsc.recordcount=1 then curso=rsc("curso") else curso=""
			rsc.close

			sqltemp="select materia from grd_materias where codmat='" & request.form("disciplina") & "'"
			rsc.open sqltemp, ,adOpenStatic, adLockReadOnly
			if rsc.recordcount=1 then materia=rsc("materia") else materia=""
			rsc.close

			perlet=left(request.form("perlet"),6)
			perlet2=right(request.form("perlet"),6)
		
			sqlb = " SELECT '" & perlet & "', "
			sqlb=sqlb & " '" & perlet2 & "', "
			sqlb=sqlb & " '" & perlet & "', "
			sqlb=sqlb & " '" & request.form("codcur") & "', "
			sqlb=sqlb & " '" & curso & "', "
			sqlb=sqlb & " '" & request.form("turno") & "', "
			sqlb=sqlb & " '" & request.form("serie") & "', "
			sqlb=sqlb & " '" & request.form("turma") & "', "
			sqlb=sqlb & " " & request.form("diasem") & ", "
			'if request.form("horini")="" then sqlb=sqlb & "null," else sqlb=sqlb & " #" & request.form("horini") & "#, "
			'horfim=cdate(request.form("horini")) + cdate("00:50")
			'sqlb=sqlb & " #" & horfim & "#, "
			'sqlb=sqlb & " #" & "00:50" & "#, "
			if request.form("a1")="1" then a1=1 else a1="null"
			if request.form("a2")="1" then a2=1 else a2="null"
			if request.form("a3")="1" then a3=1 else a3="null"
			if request.form("a4")="1" then a4=1 else a4="null"
			if request.form("a5")="1" then a5=1 else a5="null"
			if request.form("a6")="1" then a6=1 else a6="null"
			sqlb=sqlb & a1 & ", "
			sqlb=sqlb & a2 & ", "
			sqlb=sqlb & a3 & ", "
			sqlb=sqlb & a4 & ", "
			sqlb=sqlb & a5 & ", "
			sqlb=sqlb & a6 & ", "
			if request.form("inicio")="" then sqlb=sqlb & "null," else sqlb=sqlb & " '" & dtaccess(request.form("inicio")) & "', "
			if request.form("termino")="" then sqlb=sqlb & "null," else sqlb=sqlb & " '" & dtaccess(request.form("termino")) & "', "
			if request.form("juntar")="ON" then juntar = 1 else juntar = 0
			sqlb=sqlb & juntar & ", "
			if request.form("dividir")="ON" then dividir = 1 else dividir = 0
			sqlb=sqlb & dividir & ", "
			if request.form("extra")="ON" then extra = 1 else extra = 0
			sqlb=sqlb & extra & ", "
			if request.form("demons")="ON" then demons = 1 else demons = 0
			sqlb=sqlb & demons & ", "
			adnot=0
			'if request.form("turno")="3" then
			'	if request.form("a6")="1" then
			'		adnot=1			
			'	end if
			'end if
			sqlb=sqlb & " " & adnot & ", "
			sqlb=sqlb & " '" & request.form("disciplina") & "', "
			sqlb=sqlb & " '" & materia & "', "
			sqlb=sqlb & " '" & request.form("sala") & "', "
			sqlb=sqlb & " '" & session("usuariomaster") & "', "
			sqlb=sqlb & " getdate(), "
			codtur=request.form("grupo")
			if request.form("turno")="71" then ct1="M"
			if request.form("turno")="1" then ct1="M"
			if request.form("turno")="72" then ct1="V"
			if request.form("turno")="2" then ct1="V"
			if request.form("turno")="73" then ct1="N"
			if request.form("turno")="74" then ct1="N"
			if request.form("turno")="75" then ct1="N"
			codtur=codtur & "-" & ct1 & request.form("turma") & request.form("serie")
			sqlb=sqlb & " '" & codtur & "', "

			sqlb=sqlb & " '" & request.form("chapa1") & "' "
			sql = sqla & sqlb
			'response.write sql
			if tudook=1 then conexao.Execute sql, , adCmdText
		end if
	'end if 'request btsalvar
else 'request.form=""
end if

if request.form("bt_salvar")<>"" then
	if chapa2="0" then chapa2=""
	horini=cdate(request.form("horini"))+cdate("00:50")
	if request.form("horini")="22:10" then
		diasem=cint(request.form("diasem"))+1:if diasem=8 then diasem=2
		horini="19:30"
	else
		diasem=request.form("diasem")
		horini=request.form("horini")
	end if
else
	diasem=request.form("diasem")
	horini=request.form("horini")
end if	
	
'if request.form("bt_salvar")="" then
%>
<form method="POST" action="grade_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
	<tr><td class=grupo>Inclusão de Grade Horária</td></tr>
</table>
<%
'for each strItem in Request.form
'	Response.write stritem & " = " & request.form(stritem) & " "
'next
if request.form("codcur")="" then codcur="0" else codcur=request.form("codcur")
if request.form("codcur")="DIR" then filial=1 else filial=3
if request.form("turno")="" then turno=0 else turno=request.form("turno")
if request.form("diasem")="" then codds=0 else codds=request.form("diasem")
if request.form("serie")="" then serie=0 else serie=request.form("serie")
if request.form("perlet")="" then
	perlet="":perlet2="":tipopl="":tipople="..."
else
	perlet=left(request.form("perlet"),6)
	perlet2=right(request.form("perlet"),6)
	tipopl=mid(perlet2,5,1)
	if tipopl="A" then tipople="Anual"
	if tipopl="S" then tipople="Semestral"
end if
'horini=request.form("horini")
if request.form("disciplina")="" then disciplina=0 else disciplina=request.form("disciplina")
chapa1=request.form("chapa1")
chapa2=request.form("chapa2")

%>
<!-- Periodo Letivo / Curso -->
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Curso</td>
	<td class=titulo>Período Letivo</td>
</tr>
<tr>
<!--curso -->
	<td class=fundo><select size="1" name="codcur" onfocus="javascript:window.status='Selecione o curso'" onChange="javascript:submit()">
<%
if session("usuariomaster")="02379" then lanc=" in (0,1) " else lanc=" in (1) "
sqla="SELECT p.coddoc, p.curso as nome FROM grades_per p " & _
"where tper='L' and lanc" & lanc & _
"and p.coddoc in (select coddoc from grades_user where usuario='" & session("usuariomaster") & "') " & _
"GROUP BY p.coddoc, p.curso ORDER BY p.curso "

rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
response.write "<option value='0'>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if request.form("codcur")=rsd("coddoc") then tempc="selected" else tempc=""
%>
		<option value="<%=rsd("coddoc")%>" <%=tempc%>><%=rsd("nome")%></option>
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
sqla="SELECT perlet, perlet2 FROM grades_per where lanc=1 and coddoc='" & codcur & "' GROUP BY perlet, perlet2 "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
response.write "<option value='0'>....</option>"
rsd.movefirst:do while not rsd.eof
if request.form("perlet")=rsd("perlet") & rsd("perlet2") then temppl="selected" else temppl=""
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
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Período</td>
	<td class=titulo>Série/Turma</td>
	<td class=titulo>Dia da Semana</td>
	<td class=titulo>Sala</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="turno" onfocus="javascript:window.status='Selecione o período'" onChange="javascript:submit()">
<%
sqla="SELECT codturno, descturno from eturnos where codturno in (71,72,73,74,75,2,1) order by descturno "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
response.write "<option value='0'>....</option>"
rsd.movefirst:do while not rsd.eof
if cint(request.form("turno"))=cint(rsd("codturno")) then tempt="selected" else tempt=""
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
response.write "<option value='0'>..</option>"
for lserie=65 to 90
letra=chr(lserie)
'response.write letra
if request.form("turma")=letra then temptm="selected" else temptm=""
%>
		<option value="<%=letra%>" <%=temptm%>><%=letra%></option>
<%
next
%>
	</select></td>
<input type="hidden" name="diasemant" value="<%=request.form("diasem")%>">
	<td class=fundo><select size="1" name="diasem" onfocus="javascript:window.status='Selecione o dia da semana'" onchange="diasem1()">
<%
if diasem="" or isnull(diasem) then diasem=0
response.write "<option value='0'>....</option>"
for dia=2 to 7
diasemn=weekdayname(dia,-1)
if cint(diasem)=cint(dia) then tempd="selected" else tempd=""
%>
		<option value="<%=dia%>" <%=tempd%>><%=diasemn%></option>
<%
next
%>
	</select></td>

	<td class=fundo><select class=small size="1" name="sala" onfocus="javascript:window.status='Selecione a sala'">
<%
sqla="SELECT CODSALA as sala, SALADESC FROM ESALAS WHERE CODFILIAL=" & filial & " ORDER BY SALADESC "
sqla="select * from ( select codsala, saladesc, salacap, 'cadeiras' as tipo from esalas where codfilial=" & filial & "  UNION ALL select codsala, saladesc & ' (*)', salacap, tipo from grades_esalas) as salas order by saladesc "
sqla="select sala as codsala, saladesc, salacap, tipo from grades_esalas where codfilial=" & filial & " order by saladesc "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
response.write "<option value=''>Selecione....</option>"
rsd.movefirst:do while not rsd.eof
if request.form("sala")=rsd("codsala") then tempd="selected" else tempd=""
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
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo colspan=6>Horário da aula</td>
	<td class=titulo>Disciplina</td>
</tr>
<tr>
<%
cols=0
sqla="select horini, descricao from grd_defhor where codds=" & codds & " and codtn=" & turno & " order by horini "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
cols=rsd.recordcount
rsd.movefirst:do while not rsd.eof
%>
	<td class=fundo><%=rsd("descricao")%><br><input type="checkbox" name="a<%=rsd.absoluteposition%>" value="1" <%if request.form("A" & rsd.absoluteposition)="1" then response.write "checked"%>></td>
<%
rsd.movenext:loop
	if cols<6 then
		for a=cols+1 to 6
		response.write "<td class=fundo><input type=hidden name=a" & a & " value=0></td>"
		next
	end if
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
"FROM grades_gc AS c INNER JOIN grades_materias AS m ON (c.serie=m.serie) AND (c.GC=m.GC) AND (c.coddoc=m.CODdoc) " & _
"WHERE c.coddoc='" & codcur & "' AND c.perlet='" & perlet & "' AND c.serie=" & serie 
if session("usuariomaster")<>"02379" then sqla=sqla & " AND m.demons=0 "
sqla=sqla & " order by m.materia " 
response.write sqla
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
response.write "<option value=""0"">Selecione....</option>"
if rsd.recordcount>0 then
rsd.movefirst:do while not rsd.eof
if disciplina=rsd("codmat") then tempdi="selected" else tempdi=""
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
	<a class=r href="hstdisciplina.asp?codmat=<%=disciplina%>" onclick="NewWindow(this.href,'Pesquisa_disciplinas','545','200','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Magnify.gif" width="16" height="16" border="0" alt=""></a>
	</td>
	</tr>
</table>

<!-- Chapa / Nome -->
<input type="hidden" name="chapa1ant" value="<%=request.form("chapa1")%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>CH</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=chapa1%>" name="chapa1" size="8" onfocus="javascript:window.status='Informe o chapa do professor'" onchange="chapa3()"></td>
	<td class=fundo>
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
<%
rsc.movenext:loop
%>
		</select></td>
	<td class=fundo align="left">&nbsp;
<%
sqlc="select count(chapa1) as taulas from grades_3ch where chapa1='" & chapa1 & "' and perlet2 like '" & left(perlet2,4) & "%" & right(perlet2,1) & "' and juntar=0 "
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
sqla="select pini, pfim from grades_per where coddoc='" & codcur & "' and perlet='" & perlet & "' and perlet2='" & perlet2 & "' "
rsd.Open sqla, ,adOpenStatic, adLockReadOnly
if rsd.recordcount>0 then
	pini=rsd("pini"):pfim=rsd("pfim")
else
	pini="":pfim=""
end if
rsd.close
%>
 
<table border="0" cellpadding="3" cellspacing="0" width="500">
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
<% if session("usuariomaster")<>"02379" then %>
	<td class=fundo><%=pini%><input type="hidden" name="inicio" size="12" value="<%=pini%>"></td>
	<td class=fundo><%=pfim%><input type="hidden" name="termino" size="12" value="<%=pfim%>"></td>
	<td class=fundo><input type="checkbox" name="juntar" value="ON"></td>
	<td class=fundo><input type="checkbox" name="dividir" value="ON"></td>
<% else %>
	<td class=fundo><input type="text" name="inicio" size="12" value="<%=pini%>"></td>
	<td class=fundo><input type="text" name="termino" size="12" value="<%=pfim%>"></td>
	<td class=fundo><input type="checkbox" name="juntar" value="ON"></td>
	<td class=fundo><input type="checkbox" name="dividir" value="ON"></td>
	<td class=fundo><input type="checkbox" name="extra" value="ON"></td>
	<td class=fundo><input type="checkbox" name="demons" value="ON"></td>
<% end if %>
</tr>
</table>
  
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
	<% if (taulas1>=limitech or taulas2>=limitech) and chapa1<>"99999" then %>
	<font color=red>Professor excede o limite de 20 aulas!</font>
		<% if session("usuariomaster")="02379" then %>
			<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
		<% end if %>
	<% else %>
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
	<% end if%>
	</td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2" onfocus="javascript:window.status='Clique para desfazer e limpar a tela'"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()" onfocus="javascript:window.status='Clique aqui para fechar sem salvar'"></td>
</tr>
</table>
</form>
<%
'else
'rs.close
set rs=nothing
'end if   'request.form=""
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if

'	Response.write "<p>Registro salvo.<br>"
	'response.write '<script>javascript:top.window.close();</script>
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<!--
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if
%>
</body>
</html>