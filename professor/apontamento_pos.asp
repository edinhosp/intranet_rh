<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a73")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Apontamento dos Professores - Pós</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome1() {	form.chapa.value=form.nome.value;form.submit()	}
function chapa1() {	form.nome.value=form.chapa.value;form.submit()	}
function tipo_1() {
	ttemp=form.tipo1.checked
	form.tipo2.checked=false
	form.tipo1.checked=true
//	alert ("Tipo1 " + ttemp)
	form.submit()
}
function tipo_2() {
	ttemp=document.form.tipo2.checked
	form.tipo1.checked=false
	form.tipo2.checked=true
//	alert ("Tipo2 " + ttemp)
	form.submit()
}
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, marc(6), formato(6)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

tabela=690
	mesbasec=request.form("mesbase")
	cursoc  =request.form("curso")
	'if request.form("dtaula")="" then diac=mesbasec else diac=request.form("dtaula")
	diac=request.form("dtaula")
	chapac  =request.form("chapa")
	tipo1   =request.form("tipo1")
	tipo2   =request.form("tipo2")
	if tipo1="" and tipo2="" then tipo1="1"

if request.form<>"" then
	iCount=request.form("Count")
	for iLoop=0 to iCount
		id_carga=request.form("id_" & iLoop)
		if request.form("excluir" & iloop)<>"" then
			strSql="delete from clc_cargap where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
		if request.form("dad" & iloop)<>request.form("adad" & iloop) then
			if request.form("dad" & iloop)="" then dadas="null" else dadas=request.form("dad" & iloop)
			strSql="update clc_cargap set aula_dada=" & nraccess(dadas) & " where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
		if request.form("ori" & iloop)<>request.form("aori" & iloop) then
			if request.form("ori" & iloop)="" then orientacao="null" else orientacao=request.form("ori" & iloop)
			strSql="update clc_cargap set orient=" & nraccess(orientacao) & " where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
		if request.form("sup" & iloop)<>request.form("asup" & iloop) then
			if request.form("sup" & iloop)="" then supervisao="null" else supervisao=nraccess(request.form("sup" & iloop))
			strSql="update clc_cargap set superv=" & supervisao & " where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
		if request.form("adn" & iloop)<>request.form("aadn" & iloop) then
			if request.form("adn" & iloop)="" then adicnot="null" else adicnot=nraccess(request.form("adn" & iloop))
			strSql="update clc_cargap set adn=" & adicnot & " where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
		if request.form("obs" & iloop)<>request.form("aobs" & iloop) then
			strSql="update clc_cargap set observacao='" & request.form("obs" & iloop) & "' where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
	next
	if request.form("novachapa")<>"" then
		novachapa  =numzero(request.form("novachapa"),5)
		novadata   =request.form("novadata")
		novoperiodo=request.form("novoperiodo")
		diasem=weekday(novadata)
		sSql="Insert Into clc_cargap (mes_base,chapa,dia_mes,descr,doc,dia) "
		sSql=sSql & "Values ('" & dtaccess(mesbasec) & "', '" & novachapa & "','" & dtaccess(novadata) & "','" & novoperiodo & "','" & cursoc & "'," & diasem & ""
		sSql=sSql & ")"
		conexao.Execute sSQL, , adCmdText
	end if
	manutencao=1
end if
%>
<form name="form" action="apontamento_pos.asp" method="post">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Apontamento dos Professores - Pós-Graduação</td></tr>
</table>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr>
	<td class=titulor style="border-right:1px solid #000000">Mês Base</td>
	<td class=titulor><input type="checkbox" name="tipo1" value="1" <%if tipo1="1" then response.write "checked"%> onClick="tipo_1()">Curso</td>
	<td class=titulor style="border-right:1px solid #000000">Data</td>
	<td class=titulor><input type="checkbox" name="tipo2" value="1" <%if tipo2="1" then response.write "checked"%> onClick="tipo_2()">Chapa</td>
</tr>
<tr>
	<td class=titulor style="border-right:1px solid #000000;border-bottom:1px solid #000000">
	<select size="1" name="mesbase" onchange="javascript:submit()">
		<option value=""></option>
<%
if mesbasec="" then mesbasec=now()
sqla="SELECT mes_base FROM clc_cargap GROUP BY mes_base order by mes_base desc"
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if cdate(mesbasec)=cdate(rsc("mes_base")) then tempmb="selected" else tempmb=""
%>
          <option value="<%=rsc("mes_base")%>" <%=tempmb%>><%=rsc("mes_base")%></option>
<%
rsc.movenext:loop
rsc.close
%>
	</select>	
	</td>
	<td class=titulor style="border-bottom:1px solid #000000">
	<select size="1" name="curso" class=small onchange="javascript:submit()">
		<option value=""></option>
<%
sqla="SELECT c.doc, g.curso FROM clc_cargap c, g2cursoeve g WHERE c.doc=g.coddoc and " & _
"Mes_base='" & dtaccess(mesbasec) & "' GROUP BY c.doc, g.curso ORDER By g.curso "
'response.write sqla
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
rsc.movefirst:do while not rsc.eof
if cursoc=rsc("doc") then tempc="selected" else tempc=""
%>
          <option value="<%=rsc("doc")%>" <%=tempc%>><%=rsc("curso")%></option>
<%
rsc.movenext:loop
else
	response.write "<option value=''></option>"
end if
rsc.close
%>
	</select>	
	</td>
	<td class=titulor style="border-right:1px solid #000000;border-bottom:1px solid #000000">
	<select size="1" name="dtaula" onchange="javascript:submit()">
<%
sqla="SELECT dia_mes FROM clc_cargap WHERE Mes_base='" & dtaccess(mesbasec) & "' GROUP BY dia_mes "
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
if diac="" then diac=rsc("dia_mes")
rsc.movefirst:do while not rsc.eof
if cdate(diac)=cdate(rsc("dia_mes")) then tempd="selected" else tempd=""
%>
          <option value="<%=rsc("dia_mes")%>" <%=tempd%>><%=rsc("dia_mes")%></option>
<%
rsc.movenext:loop
else
	response.write "<option value=''></option>"
end if
rsc.close
%>
	</select>	
	</td>
	<td class=titulor style="border-bottom:1px solid #000000">
	<input type="text" name="chapa" value="<%=chapac%>" size="5" onchange="chapa1()">
	<select size="1" name="nome" class=small onchange="nome1()">
		<option value=""></option>
<%
if tipo2="1" then
sqla="SELECT c.CHAPA, P.NOME FROM clc_cargap c INNER JOIN corporerm.dbo.PFUNC p ON c.CHAPA=P.CHAPA collate database_default " & _
"where p.codsituacao in ('A','F','Z') and Mes_base='" & dtaccess(mesbasec) & "' " & _
"GROUP BY c.CHAPA, P.NOME ORDER BY P.NOME "
rsc.Open sqla, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if chapac=rsc("chapa") then tempch="selected" else tempch=""
%>
		<option value="<%=rsc("chapa")%>" <%=tempch%>><%=rsc("nome")%></option>
<%
rsc.movenext:loop
rsc.close
end if
%>
	</select>	
	</td>
</tr>
</table>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr>
<% if tipo1="1" then %>
	<td class=titulor>Chapa</td>
	<td class=titulor>Nome</td>
<%end if
if tipo2="1" then%>
	<td class=titulor>Curso</td>
	<td class=titulor>Data</td>
	<td class=titulor>Per.</td>
<%end if%>
	<td class=titulor>Dia</td>
	<td class=titulor>Tur</td>
	<td class=titulor>Prev.</td>
	<td class=titulor>Dada</td>
	<td class=titulor>Orient</td>
	<td class=titulor>Superv</td>
	<td class=titulor>Ad.Not.</td>
	<td class=titulor>Obs.</td>
	<td class=titulor><img src="../images/Trash.gif" border="0"></td>
</tr>
<%
if diac="" then diac=now()
	if tipo1="1" then
		sqla="select c.*, f.nome, g.curso from clc_cargap c, corporerm.dbo.pfunc f, g2cursoeve g " 
		sqla=sqla & "where f.chapa collate database_default=c.chapa and c.doc=g.coddoc and mes_base='" & dtaccess(mesbasec) & "' "
		sqla=sqla & "and doc='" & cursoc & "' "
		sqla=sqla & "and dia_mes='" & dtaccess(diac) & "' "
		sqla=sqla & "order by descr, nome, c.chapa, dia_mes "
	elseif tipo2="1" then
		sqla="select c.*, f.nome, g.curso from clc_cargap c, corporerm.dbo.pfunc f, g2cursoeve g " 
		sqla=sqla & "where f.chapa collate database_default=c.chapa and c.doc=g.coddoc and mes_base='" & dtaccess(mesbasec) & "' "
		sqla=sqla & "and c.chapa='" & chapac & "' "
		sqla=sqla & "order by g.curso, dia_mes, descr "
	else
		sqla="select * from clc_cargap " 
		sqla=sqla & "where mes_base='" & dtaccess(mesbasec) & "' "
	end if
	'response.write sqla
	rs.Open sqla, ,adOpenStatic, adLockReadOnly

if rs.recordcount>0 then
tcount=0

rs.movefirst
do while not rs.eof

if tipo1="1" then
if lastper=rs("descr") then a=a else response.write "<tr><td class=grupor colspan=13>" & rs("descr") & "</td></tr>"
end if
if tipo2="1" then
if lastcur=rs("doc") then a=a else response.write "<tr><td class=grupor colspan=14>" & rs("curso") & "</td></tr>"
end if
%>

<tr>
<input type="hidden" name="id_<%=tcount%>" value="<%=rs("id_carga")%>">
<% if tipo1="1" then %>
	<td class="campor" style="border-bottom:1px solid #000000"><%=rs("chapa")%></td>
	<td class="campor" style="border-bottom:1px solid #000000" nowrap><%=rs("nome")%></td>
<%end if
if tipo2="1" then%>
	<td class="campor" style="border-bottom:1px solid #000000" nowrap><%=lcase(rs("curso"))%></td>
	<td class="campor" style="border-bottom:1px solid #000000"><%=rs("dia_mes")%></td>
	<td class="campor" style="border-bottom:1px solid #000000"><%=rs("descr")%></td>
<%end if%>
	<td class="campor" style="border-bottom:1px solid #000000"><%=rs("dia")%></td>
	<td class="campor" style="border-bottom:1px solid #000000" nowrap><%=rs("turma")%></td>
	<td class="campor" style="border-bottom:1px solid #000000" align="center"><%=rs("aula_prev")%></td>
	<td class="campor" style="border-bottom:1px solid #000000">
		<input type="text" name="dad<%=tcount%>" value="<%=rs("aula_dada")%>" size="3" class=form_apt>
		<input type="hidden" name="adad<%=tcount%>" value="<%=rs("aula_dada")%>">
	</td>
	<td class="campor" style="border-bottom:1px solid #000000">
		<input type="text" name="ori<%=tcount%>" value="<%=rs("orient")%>" size="3" class=form_apt>
		<input type="hidden" name="aori<%=tcount%>" value="<%=rs("orient")%>">
	</td>
	<td class="campor" style="border-bottom:1px solid #000000">
		<input type="text" name="sup<%=tcount%>" value="<%=rs("superv")%>" size="3" class=form_apt>
		<input type="hidden" name="asup<%=tcount%>" value="<%=rs("superv")%>">
	</td>
	<td class="campor" style="border-bottom:1px solid #000000">
		<input type="text" name="adn<%=tcount%>" value="<%=rs("adn")%>" size="3" class=form_apt>
		<input type="hidden" name="aadn<%=tcount%>" value="<%=rs("adn")%>">
	</td>
	<td class="campor" style="border-bottom:1px solid #000000">
		<input type="text" name="obs<%=tcount%>" value="<%=rs("observacao")%>" size="15" class=form_apt>
		<input type="hidden" name="aobs<%=tcount%>" value="<%=rs("observacao")%>">
	</td>
    <td class="campor" style="border-bottom:1px solid #000000;border-left:1px solid #000000">
    <% if session("a73")="T" and isnull(rs("aula_prev")) then %>
	<input type="checkbox" name="excluir<%=tcount%>" value="1">
	<% end if %>
	<%
	if rs("aula_dada")>rs("aula_prev") or rs("adn")>1 then
		response.write "<b><font color=red>O lançamento está fora do esperado.</font></b>"
	
	end if
	%>
	</td>
<!-- marcacoes do dia -->
<%
	sqlcr="select chapa, day(data) as dia, data, batida, status from abatfun_m where " & _
	"chapa='" & rs("chapa") & "' and data='" & rs("dia_mes") & "' order by data, batida"
	sqlcr="select chapa, day(data) as dia, data, batida, status from abatfun_M where " & _
	"chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("dia_mes")) & "' order by batida"   'access
	'marcações do chronus
	sqlcr="select chapa, day(data) as dia, data, hora from _catraca where " & _
	"chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("dia_mes")) & "' order BY hora"  'sql
	sqlcr="select chapa, day(data) as dia, data, batida, status from corporerm.dbo.abatfun where " & _
	"chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("dia_mes")) & "' order BY batida"  'sql
	rs2.Open sqlcr, ,adOpenStatic, adLockReadOnly
	marcacao=0
	for b=1 to 6:marc(b)="":formato(b)="":next
	if rs2.recordcount>0 then
		rs2.movefirst
		do while not rs2.eof
		batida=formatdatetime((rs2("batida")/60)/24,4)   '--chronus
		'batida=formatdatetime(rs2("hora"),4)   '--catraca
		marcacao=marcacao+1 'else marcacao=1
		marc(marcacao)=batida
		'if rs2("status")="D" then formato(marcacao)="<font color='red'>" 'else formato(dia,marcacao)="<font color='black'>"
		rs2.movenext
		loop
	else 'recordcount rs2
		for b=1 to 6
			marc(b)=""
		next
	end if 'recordcount rs2
	if marcacao<6 then
		for a=marcacao+1 to 6
			marc(a)=""
		next
	end if
	rs2.close
%>
	<td class="campor" style="border-bottom:1px solid #000000">
<%
for a=1 to 6
	response.write "<font color=blue>|</font>" & formato(a) & marc(a) & "</font>"
next 
%>	
	</td>
</tr>
<%
lastper=rs("descr")
lastcur=rs("doc")
rs.movenext
tcount=tcount+1
loop
end if 'recordcount>0
%>
<% if tipo1="1" then %>
<tr><td class="campor" colspan=12 style="border-bottom:1px solid #000000">
<input type="text" name="novachapa" value="" size="7" class=form_apt>
	<select size="1" name="novoperiodo" class=small>
		<option value=""  <%if lastper="" then response.write "selected"%>></option>
		<option value="Manhã" <%if lastper="Manhã" then response.write "selected"%>>Manhã</option>
		<option value="Tarde" <%if lastper="Tarde" then response.write "selected"%>>Tarde</option>
		<option value="Noite" <%if lastper="Noite" then response.write "selected"%>>Noite</option>
	</select>	
<input type="text" name="novadata" value="<%=diac%>" size="10" class=form_apt>
</td></tr>
<%end if%>
</table>
<input type="hidden" name="Count" value="<%=tcount-1%>">
<input type="reset" value="Desfazer" class=button>
<input type="submit" value="Confirmar" class=button>
</form>

<%
'response.writ request.form
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>