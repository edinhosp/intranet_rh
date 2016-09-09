<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a80")="N" or session("a80")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
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
--></script>
<script language="VBScript">
</script>

<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
Function IIf(condition,value1,value2)
	If condition Then IIf = value1 Else IIf = value2
End Function

ocorrencia="":tipomov=""

dim conexao, conexao2, chapach, rs, rs2, rs3, a(6)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request("idturma")<>"" then session("id_grdturma")=request("idturma")

if session("selecaoprofessor")="" then session("selecaoprofessor")="disciplina"
if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		if (request.form("chapa1")="" or request.form("chapa1")="0") then
			tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Preencha todos os campos do cadastro!');</script>"
		end if

'iif(request.form("extra")="ON",1,0)
		
		if request.form("quando")<>"" then dataaut="'" & dtaccess(request.form("quando")) & "'" else dataaut="null"
'		response.write "<br>" & request.form & "<br><br>"
		if request.form("gravado")="0" and request.form("substituir")<>"" then
'response.write "<br> Executou Rotina 1<br>"
			sql="update g2aulas set chapa1='" & request.form("chapa1") & "', ativo=1, usuarioa='" & session("usuariomaster") & "', dataa='" & dtaccess(now()) & "' " & _
			"where id_grdturma=" & request.form("id_grdturma") & " and id_grdaula=" & request.form("id_grdaula") & " and codmat='" & request.form("codmat") & "' "
			response.write sql
			if tudook=1 then conexao.Execute sql, , adCmdText
		end if		

		if request.form("gravado")="0" and request.form("substituir")="" then
'response.write "<br> Executou Rotina 2<br>"
			sql="insert into g2aulas (id_grdturma, codmat, chapa1, ativo, usuarioc, datac) " & _
			"select " & request.form("id_grdturma") & ",'" & request.form("codmat") & "', '" & request.form("chapa1") & "',1,'" & session("usuariomaster") & "','" & dtaccess(now()) & "' "
			if session("usuariogrupo")="RH" then
				sql="insert into g2aulas (id_grdturma, codmat, chapa1, ativo, usuarioc, datac, autorizado, quando, juntar) " & _
				"select " & request.form("id_grdturma") & ",'" & request.form("codmat") & "', '" & request.form("chapa1") & "',1,'" & session("usuariomaster") & "','" & dtaccess(now()) & "' " & _
				"," & iif(request.form("autorizado")="ON",1,0) & ", " & dataaut & _ 
				"," & iif(request.form("juntar")="ON",1,0)
			else
				sql="insert into g2aulas (id_grdturma, codmat, chapa1, ativo, usuarioc, datac) " & _
				"select " & request.form("id_grdturma") & ",'" & request.form("codmat") & "', '" & request.form("chapa1") & "',1,'" & session("usuariomaster") & "','" & dtaccess(now()) & "' "
			end if
			'response.write sql
			if tudook=1 then conexao.Execute sql, , adCmdText
		end if
		if request.form("gravado")="1" and request.form("chapa1")<>request.form("chapa1_o") and request.form("substituir")="" then
'response.write "<br> Executou Rotina 3<br>"
			sql="insert into g2aulas (id_grdturma, codmat, chapa1, ativo, usuarioc, datac) " & _
			"select " & request.form("id_grdturma") & ",'" & request.form("codmat") & "', '" & request.form("chapa1") & "',1,'" & session("usuariomaster") & "','" & dtaccess(now()) & "' "
			'response.write sql
			if tudook=1 then conexao.Execute sql, , adCmdText
		end if
		if request.form("gravado")="1" and request.form("chapa1")=request.form("chapa1_o") then
'response.write "<br> Executou Rotina 4<br>"
			sql="update g2aulas set autorizado=" & iif(request.form("autorizado")="ON",1,0) & ", quando=" & dataaut & " where id_grdaula=" & request.form("id_grdaula")
			if tudook=1 then conexao.Execute sql, , adCmdText
			sql="update g2aulas set juntar=" & iif(request.form("juntar")="ON",1,0) & " where id_grdaula=" & request.form("id_grdaula")
			if tudook=1 then conexao.Execute sql, , adCmdText
		end if
		
		iPago=request.form("nrdias"):'response.write "<br>" & ipago
		for iLoop=0 to iPago
			id_pago=request.form("pago" & iLoop)
			id_data=request.form("id" & iLoop)
			'response.write "<br>pago " & id_pago
			if id_pago<>"" then valor=1 else valor=0
			strSql="update g2aulasdata set pago=" & valor & " where id_data=" & id_data
			'response.write "<br>" & strsql
			conexao.execute strSql, , adCmdText
		next
		
	end if 'button=salvar
else 'request.form=""
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	'inicio=request.form("inicio")
	sql="UPDATE g2aulas set deletada=1, usuarioa='" & session("usuariomaster") & "', dataa=getdate() WHERE id_grdaula=" & request.form("id_grdaula")
	if session("usuariomaster")="02379" then sql="UPDATE g2aulas set deletada=1 WHERE id_grdaula=" & request.form("id_grdaula")
	'if now()>(cdate(inicio)+dias) then tudook=0:ocorrencia=ocorrencia & "<Br><font color=blue><b>As aulas já iniciaram. O lançamento não pode ser excluido!</b></font>"
	if session("usuariomaster")="02379" then tudook=1
	if tudook=1 then conexao.Execute sql, , adCmdText
	'
	'Ver quando excluir jpai=1 então zerar juntar=1 ou passar um quando count(juntar=1)>1 para jpai=1 
end if

if request.form("bt_selprof")<>"" then
	if session("selecaoprofessor")="disciplina" then session("selecaoprofessor")="todos" else session("selecaoprofessor")="disciplina"
end if
if session("selecaoprofessor")="disciplina" then txtbut="PROF." else txtbut="DISC."

'if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) or (request.form("bt_excluir")<>"" and tudook=0) then
if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and (tudook=0 or tudook=1)) or (request.form("bt_excluir")<>"" and (tudook=0 or tudook=1)) then
	'response.write request.form
	if request("codmat")="" then codmat=request.form("codmat") else codmat=request("codmat")
	sql1="select * from g2aulas where id_grdturma=" & session("id_grdturma") & " and codmat='" & codmat & "' and deletada=0 "
	if request.form("chapa1")<>"" then sql1=sql1 & " and chapa1='" & request.form("chapa1") & "'"
	if session("usuariomaster")="02379" then response.write "->" & sql1
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	sql2="select t.*, c.tipocurso, inicio, termino from g2turmas t, g2cursos c, g2periodoaula p " & _
	"where p.perlet=t.perlet and t.codcur=c.codcur and t.codper=c.codper " & _
	"and id_grdturma=" & session("id_grdturma") & " "
	
	rs1.Open sql2, ,adOpenStatic, adLockReadOnly

	if rs.recordcount=0 then
		gravado=0
		'codmat=request("codmat")
		chapa1=request.form("chapa1")
		id_grdaula=request.form("id_grdaula")
	else
		gravado=1
		if request.form("codmat") <>"" then codmat =request.form("codmat")  else codmat =rs("codmat")
		if request.form("chapa1") <>"" then chapa1 =request.form("chapa1")  else chapa1 =rs("chapa1")
		if request.form("id_grdaula") <>"" then id_grdaula =request.form("id_grdaula")  else id_grdaula =rs("id_grdaula")
		id_grdaula=rs("id_grdaula"):codmat =rs("codmat"):chapa1 =rs("chapa1")
		if (rs("autorizado")=true or rs("autorizado")=1) or request.form("autorizado")="ON" then autorizado="checked" else autorizado=""
		if (rs("juntar")=true or rs("juntar")=1) or request.form("juntar")="ON" then juntar="checked" else juntar=""
		if request.form("quando")<>"" then quando=request.form("quando") else quando=rs("quando")
	end if

 	sqlm="select materia from corporerm.dbo.umaterias where codmat='" & codmat & "' "
	rs2.Open sqlm, ,adOpenStatic, adLockReadOnly
	materia=rs2("materia")
	rs2.close
	
	if request("duplicar")=1 then
		did_grdturma=request("id_grdturma")
		dcodmat=request("codmat")
		dchapa1=request("chapa1")
		response.write "<br>Duplicar: " & did_grdturma & "/" & dcodmat & "/" & dchapa1 & "/" & id_grdaula
		sql1="insert into g2aulas (id_grdturma, codmat, chapa1, ativo, usuarioc, datac) " & _
		"select " & did_grdturma & ", '" & dcodmat & "', '" & dchapa1 & "', 1, '" & session("usuariomaster") & "', getdate() "
		conexao.execute sql1

		sql2="select id_grdaula from g2aulas where id_grdturma=" & did_grdturma & " and codmat='" & dcodmat & "' and chapa1='" & dchapa1 & "'"
		rs2.Open sql2, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then did_grdaula=rs2("id_grdaula") else did_grdaula=0
		rs2.close
		
		if did_grdaula>0 then
		sql3="insert into g2aulasdata (id_grdaula, data, codhor, qtaulas ) " & _
		"select " & did_grdaula & ", data, codhor, qtaulas from g2aulasdata where id_grdaula=" & id_grdaula
		conexao.execute sql3
		end if

	end if
%>
<form method="POST" action="gradenovapos.asp" name="form">
<input type="hidden" name="id_grdaula"  size="4" value="<%=id_grdaula%>"> 
<input type="hidden" name="id_grdturma" size="4" value="<%=session("id_grdturma")%>">
<input type="hidden" name="codmat" size="6" value="<%=codmat%>">
<input type="hidden" name="chapa1_o" size="6" value="<%=chapa1%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<input type="hidden" name="gravado" size="1" value="<%=gravado%>">

<!-- quadro -->
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr><td class=campo valign=top width="80%">
<!-- quadro -->

<table border="0" cellpadding="3" cellspacing="0" width="100%">
	<tr><td class=grupo>Alteração de Grade Horária (<%=id_grdaula%>)</td></tr>
</table>

  
<!-- Chapa / Nome -->
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulop colspan=2>Disciplina: <font color=blue><%=materia%></td>
</tr>
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=chapa1%>" name="chapa1" size="8" onfocus="javascript:window.status='Informe o chapa do professor'" onchange="chapa3()"></td>
	<td class=fundo>&nbsp;
		<select size="1" name="nome1" onfocus="javascript:window.status='Selecione o Nome do Professor'" onchange="nome3()">
<%
response.write "<option value='0'>Selecione Professor 1....</option>"
response.write "<option style='background:CCFFCC' value='0'>------- Professores atribuidos --------</option>"
sql2="select distinct g.chapa1 chapa, f.nome from g5ch g, grades_aux_prof f " & _
"where f.chapa=g.chapa1 and codmat='" & codmat & "' and coddoc='" & rs1("coddoc") & "' and id_grdturma=" & session("id_grdturma") & " " 
sql2a="select distinct g.chapa1 chapa from g5ch g, grades_aux_prof f " & _
"where f.chapa=g.chapa1 and codmat='" & codmat & "' and coddoc='" & rs1("coddoc") & "' and id_grdturma=" & session("id_grdturma") & " and deletada=0 " 
sql2=sql2 & "order by nome "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
temp=""
if chapa1=rs2("chapa") then temp="selected":selec1=1
%>
		<option value="<%=rs2("chapa")%>" <%=temp%>><%=rs2("nome")%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount>0
rs2.close

response.write "<option style='background:CCFFCC' value='0'>------- Professores disponíveis --------</option>"
sql2="select chapa, nome from grades_aux_prof "
if session("selecaoprofessor")="disciplina" then
	sql2=sql2 & " where chapa in (select chapa1 chapa from g5ch where codmat='" & codmat & "' and coddoc='" & rs1("coddoc") & "' group by chapa1) "
	sql2=sql2 & " order by nome "
else
	sql2=sql2 & " where chapa not in (" & sql2a & ") order by nome "
end if
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
response.write "<br>---->" & temp
if selec1<>1 then
	if chapa1=rs2("chapa") then temp2="selected" else temp2=""
end if
%>
		<option value="<%=rs2("chapa")%>" <%=temp2%>><%=rs2("nome")%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount>0
rs2.close
data1aula=""
%>
	</select>
	<%%>
	<input type="submit" name="bt_selprof" value="<%=txtbut%>" class="button" alt="Mostrar apenas professores da disciplina" onclick="javascript:submit()">
	</td>
</tr>

<!-- inicio consistencias professor 1 -->
<tr><td class=fundo colspan=3>
<%
if chapa1<>"" then
	ocorrencia=ocorrencia & "<font color=black><b>Professor 1: </b></font>"
	'------------- verifica se existe aula no mesmo horário para juntar --------------------
	'verificar depois no salvamento das datas

	'------------- verifica se é professor habitual da matéria / se não for, se vai ganhar mais que o do período anterior
	sql10="select coddoc, chapa1 chapa from g5ch where deletada=0 " & _
	"and codmat='" & codmat & "' and coddoc='" & rs1("coddoc") & "' and inicio<'" & dtaccess(rs1("inicio")) & "' and chapa1='" & chapa1 & "' group by coddoc, chapa1 "
	rs3.Open sql10, ,adOpenStatic, adLockReadOnly:existe=rs3.recordcount:rs3.close
	if existe=0 then
		ocorrencia=ocorrencia & "<br><br><font color=red>Professor não habitual."
	else
		ocorrencia=ocorrencia & "<br><br><font color=red>Professor habitual":inconsistencia="" 
	end if
	'----------------- mudando professor / verificar data de inicio das aulas ----------------------
	if chapa1_o<>chapa1 then 
		ocorrencia=ocorrencia & "<br><br><font color=green>Mudando professor: " & chapa1_o & " para " & chapa1
	end if
	'----------------- limite de aulas ----------------------------
	'nao se aplica a pois
end if
%>
<input type="hidden" name="limite1" value="<%=limite1%>">
<input type="hidden" name="taulas1" value="<%=taulas1%>">
<input type="hidden" name="obs" value="<%=inconsistencia%>">
</td></tr></table>
<!-- final consistencias professor 1-->

<%if gravado=1 then%>
<table border="0" bordercolor="#000000" cellpadding="4" cellspacing="0" width="100%">
<tr>
	<td class=titulop width=90>Datas</td>
	<td class=titulop width=25>&nbsp;</td>
	<td class=titulo width=50># Aulas</td>
	<td class=titulo>Horários</td>
	<td class=titulo>Pago?</td>
</tr>
<%
sqld="select d.id_data, d.id_grdaula, data, count(h.codhor) aulas, pago " & _
"from (g2aulas a inner join g2aulasdata d on a.id_grdaula=d.id_grdaula) left join g2aulashora h on h.id_data=d.id_data " & _
"where d.id_grdaula='" & id_grdaula & "' and (d.deletada=0 or d.deletada is null) and chapa1='" & chapa1 & "' " & _
"group by d.id_data, d.id_grdaula, data, pago order by data "
sqld="select d.id_data, d.id_grdaula, data, d.codhor, d.qtaulas aulas, pago, h.descricao " & _
"from g2aulas a inner join g2aulasdata d on a.id_grdaula=d.id_grdaula " & _
"left join g2defhor h on h.codhor=d.codhor " & _
"where d.id_grdaula='" & id_grdaula & "' and (d.deletada=0 or d.deletada is null) and chapa1='" & chapa1 & "' " & _
"order by data "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
taulas=0
if rs2.recordcount>0 then aulas=rs2.recordcount
do while not rs2.eof
if rs2.absoluteposition=1 then data1aula=rs2("data")
if rs2("pago")=true then pago="SIM" else pago="NÃO"
%>
<tr>
	<td class=fundo align="center" style="border-bottom:1px solid"><%=rs2("data")%> (<%=weekdayname(weekday(rs2("data")),1)%>)</td>
	<td class=fundo>
	<%if pago="NÃO" then%>
	<a class=r href="gradenovaposdata.asp?id_grdaula=<%=id_grdaula%>&id_data=<%=rs2("id_data")%>" onclick="NewWindow(this.href,'Data_hora','545','250','yes','center');return false" onfocus="this.blur()">
	<img src="../images/novo.gif" width="17" height="17" border="0" alt="Inserir/alterar nova data"></a>
	<%end if%>
	</td>
	<td class=fundo align="center" style="border-bottom:1px solid"><%=rs2("aulas")%></td>
	<td class=fundo style="border-bottom:1px solid">&nbsp;<%=rs2("descricao")%></td>
	<td class=fundo style="border-bottom:1px solid">&nbsp;
<%
if session("usuariogrupo")="RH" then
%>
	<input type="checkbox" name="pago<%=rs2.absoluteposition-1%>" value="<%=rs2("id_data")%>" <%if pago="SIM" then response.write "checked"%>>
	<input type="hidden" name="id<%=rs2.absoluteposition-1%>" value="<%=rs2("id_data")%>">
<%
else
	response.write pago
end if
%>
	</td>
</tr>
<%
taulas=taulas+rs2("aulas")
rs2.movenext
loop
rs2.close
%>	
<input type=hidden name=profsubs value="<%=data1aula%>">
	<tr>
		<td class=fundo>&nbsp;</td>
		<td class=fundo><a class=r href="gradenovaposdata.asp?id_grdaula=<%=id_grdaula%>&id_data=" onclick="NewWindow(this.href,'Data_hora','545','250','yes','center');return false" onfocus="this.blur()">
		<img src="../images/novo.gif" width="17" height="17" border="0" alt="Inserir/alterar nova data"></a>
		</td>
		<td class=fundo align="center" style="border-top:3px double #000000"><%=taulas%></td>
		<td class=fundo>&nbsp;</td>
		<td class=fundo>&nbsp;</td>
	</tr>
</table>
<%end if 'gravado=1%>  
<input type="hidden" name="nrdias" value="<%=aulas-1%>">


<% if session("usuariomaster")="02379" or Session("usuariogrupo")="RH" then %>
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr><td class=fundo>
	<input type="checkbox" name="autorizado" value="ON" <%=autorizado%>> Autorizado
	em <input type="text" name="quando" size="10" value="<%=quando%>">
</td>
<td class=fundo>
	<input type="checkbox" name="juntar" value="ON" <%=juntar%>> Junta Turma
</td>
</tr>
</table>
<% else %>
	Autorizado em <input type="hidden" name="quando" size="10" value="<%=quando%>"><%=quando%>
<% end if %>

<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer" class="button" name="B2"></td>
<%if gravado=1 and aulas=0 then%>
	<td class=titulo align="center"><input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
<%end if%>
</tr>
</table>


<!-- quadro -->
</td><td class=campo valign=top width="20%">
<!-- quadro -->

<table border="0" cellpadding="3" cellspacing="0" width="100%">
	<tr><td class=grupo>Ocorrências</td></tr>
	<tr><td class=campo>
	<%
	response.write ocorrencia

	if chapa1<>request.form("chapa1_o") then
	response.write "<br><br><input type=checkbox name=substituir value=""ON""> Apenas substituir o professor"
	end if
	%>
	</td></tr>
	<tr><td class=campo>
<%
sqldt="select id_grdturma, codtur from g2turmas t inner join ( " & _
"select codcur, codper, grade, perlet from g2turmas where id_grdturma=" & session("id_grdturma")  & "" & _
") s on s.codcur=t.codcur and s.codper=t.codper and s.grade=t.grade and s.perlet=t.perlet " & _
"where id_grdturma not in (" & session("id_grdturma") & ") "
rs2.Open sqldt, ,adOpenStatic, adLockReadOnly
do while not rs2.eof

sqled="select codmat, id_grdturma from g2aulas where id_grdturma=" & rs2("id_grdturma") & " and codmat=" & codmat & ""
rs3.Open sqled, ,adOpenStatic, adLockReadOnly
if rs3.recordcount=0 then
%>
	<a href="gradenovapos.asp?duplicar=1&id_grdturma=<%=rs2("id_grdturma")%>&codmat=<%=codmat%>&chapa1=<%=chapa1%>">
	Duplicar esta disciplina para a turma <%=rs2("codtur")%></a>
<%
end if
rs3.close

rs2.movenext
loop
rs2.close

%>
	</td></tr>
</table>

<!-- quadro para outros horarios -->
</td></tr>

</table>
<!-- quadro -->
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	if tudook=1 then
		'response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.location=window.opener.location;self.close();</script>"
        'response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.document.form.submit();self.close();</script>"
        response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.document.form.submit();</script>"
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