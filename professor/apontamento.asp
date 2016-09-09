<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a38")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Apontamento dos Professores</title>
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
dim conexao, conexao2, chapach, rs, rs2, marc(20), formato(20)
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
			strSql="delete from clc_carga where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
		if request.form("fal" & iloop)<>request.form("afal" & iloop) then
			if request.form("fal" & iloop)="" then faltas="null" else faltas=request.form("fal" & iloop)
			strSql="update clc_carga set faltas=" & nraccess(faltas) & " where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
		if request.form("jus" & iloop)<>request.form("ajus" & iloop) then
			strSql="update clc_carga set justificativa='" & request.form("jus" & iloop) & "' where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
		if request.form("atr" & iloop)<>request.form("aatr" & iloop) then
			if request.form("atr" & iloop)="" then atrasos="null" else atrasos=request.form("atr" & iloop)
			strSql="update clc_carga set atraso=" & nraccess(atrasos) & " where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
		if request.form("ext" & iloop)<>request.form("aext" & iloop) then
			if request.form("ext" & iloop)="" then extra="null" else extra=request.form("ext" & iloop)
			strSql="update clc_carga set extra=" & extra & " where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
		if request.form("dep" & iloop)<>request.form("adep" & iloop) then
			if request.form("dep" & iloop)="" then dp="null" else dp=request.form("dep" & iloop)
			strSql="update clc_carga set dp=" & nraccess(dp) & " where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
		if request.form("rep" & iloop)<>request.form("arep" & iloop) then
			if request.form("rep" & iloop)="" then reposicao="null" else reposicao=request.form("rep" & iloop)
			strSql="update clc_carga set reposicao=" & reposicao & " where id_carga=" & id_carga
			'response.write strSql
			conexao.execute strSql, , adCmdText
		end if
		if request.form("obs" & iloop)<>request.form("aobs" & iloop) then
			strSql="update clc_carga set observacao='" & request.form("obs" & iloop) & "' where id_carga=" & id_carga
			conexao.execute strSql, , adCmdText
		end if
	next
	if request.form("novachapa")<>"" then
		novachapa  =numzero(request.form("novachapa"),5)
		novadata   =request.form("novadata")
		novoperiodo=request.form("novoperiodo")
		diasem=weekday(novadata)
		sSql="Insert Into clc_carga (mes_base,chapa,dia_mes,descr,doc,dia) "
		sSql=sSql & "Values ('" & dtaccess(mesbasec) & "', '" & novachapa & "','" & dtaccess(novadata) & "','" & novoperiodo & "','" & cursoc & "'," & diasem & ""
		sSql=sSql & ")"
		conexao.Execute sSQL, , adCmdText
	end if
	manutencao=1
end if
%>
<form name="form" action="apontamento.asp" method="post">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Apontamento dos Professores</td></tr>
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
sqla="SELECT mes_base FROM clc_carga GROUP BY mes_base order by mes_base desc"
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
		<option value="">Nenhum curso selecionado</option>
<%
sqla="SELECT CURSO FROM clc_carga WHERE Mes_base='" & dtaccess(mesbasec) & "' GROUP BY CURSO "
sqla="SELECT c.doc, g.curso FROM clc_carga c, g2cursoeve g WHERE c.doc=g.coddoc and " & _
"Mes_base='" & dtaccess(mesbasec) & "' GROUP BY c.doc, g.curso order by g.curso"
'response.write sqla
if request.form("mesbase")<>"" then
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
end if
if session("usuariomaster")="02379" then selectsize="1" else selectsize="1"
%>
	</select>	
	</td>
	<td class=titulor style="border-right:1px solid #000000;border-bottom:1px solid #000000">
	<select size="<%=selectsize%>" name="dtaula" onchange="javascript:submit()" >
<%
sqla="SELECT dia_mes FROM clc_carga WHERE Mes_base='" & dtaccess(mesbasec) & "' GROUP BY dia_mes order by dia_mes "
if request.form("mesbase")<>"" then
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
end if
%>
	</select>	
	</td>
	<td class=titulor style="border-bottom:1px solid #000000">
	<input type="text" name="chapa" value="<%=chapac%>" size="5" onchange="chapa1()">
	<select size="1" name="nome" class=small onchange="nome1()">
		<option value=""></option>
<%
if tipo2="1" then
sqla="SELECT c.CHAPA, P.NOME FROM clc_carga c INNER JOIN corporerm.dbo.PFUNC p ON c.CHAPA=P.CHAPA collate database_default " & _
"where p.codsituacao in ('A','F','Z') " & _
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
<!--	<td class=titulor>Curso</td> -->
	<td class=titulor>Data</td>
<!--	<td class=titulor>Per.</td> -->
<%end if%>
	<td class=titulor>Dia</td>
	<td class=titulor>Tur</td>
	<td class=titulor>Aula</td>
	<td class=titulor>Falta</td>
	<td class=titulor>Just.</td>
	<td class=titulor>Atraso</td>
	<td class=titulor>Extra</td>
	<td class=titulor>Dep.</td>
	<td class=titulor>Repos.</td>
	<td class=titulor>Obs.</td>
	<td class=titulor>Hor.</td>
	<td class=titulor><img src="../images/Trash.gif" border="0"></td>
</tr>
<%
	if mesbasec="" then mesbasec=now()
	if diac="" then diac=now()
	if tipo1="1" then
		sqla="select c.*, f.nome, g.curso from clc_carga c, corporerm.dbo.pfunc f, g2cursoeve g " 
		sqla=sqla & "where f.chapa collate database_default=c.chapa and c.doc=g.coddoc and mes_base='" & dtaccess(mesbasec) & "' "
		sqla=sqla & "and doc='" & cursoc & "' "
		sqla=sqla & "and dia_mes='" & dtaccess(diac) & "' "
		sqla=sqla & "order by descr, nome, c.chapa, dia_mes "
	elseif tipo2="1" then
		sqla="select c.*, f.nome, g.curso from clc_carga c, corporerm.dbo.pfunc f, g2cursoeve g " 
		sqla=sqla & "where f.chapa collate database_default=c.chapa and c.doc=g.coddoc and mes_base='" & dtaccess(mesbasec) & "' "
		sqla=sqla & "and c.chapa='" & chapac & "' "
		sqla=sqla & "order by g.curso, dia_mes, descr "
	else
		sqla="select * from clc_carga " 
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
if lastcur=rs("doc") then a=a else response.write "<tr><td class=grupor colspan=14><b>" & rs("curso") & "</td></tr>"
end if
if rs("horini")="" or isnull(rs("horini")) then horini="" else horini=formatdatetime(rs("horini"),4)
if rs("horfim")="" or isnull(rs("horfim")) then horfim="" else horfim=formatdatetime(rs("horfim"),4)

%>

<tr>
<input type="hidden" name="id_<%=tcount%>" value="<%=rs("id_carga")%>">
<% if tipo1="1" then %>
	<td class="campor" style="border-bottom:1px solid #000000"><%=rs("chapa")%></td>
	<td class="campor" style="border-bottom:1px solid #000000" nowrap><%=rs("nome")%></td>
<%end if
if tipo2="1" then
diames=numzero(day(rs("dia_mes")),2) & "/" &numzero(month(rs("dia_mes")),2)
%>
<!--	<td class="campor" style="border-bottom:1px solid #000000" nowrap><%=lcase(rs("curso"))%></td> -->
	<td class="campor" style="border-bottom:1px solid #000000" align="center"><b><font color="#006600"><%=diames%></font></td>
<!--	<td class="campor" style="border-bottom:1px solid #000000"><%=rs("descr")%></td> -->
<%end if%>
	<td class="campor" style="border-bottom:1px solid #000000" align="center"><%=weekdayname(rs("dia"),1)%></td>
	<td class="campor" style="border-bottom:1px solid #000000" nowrap><%=rs("turma")%></td>
	<td class="campor" style="border-bottom:1px solid #000000" align="center"><%=rs("aulas")%></td>
	<td class="campor" style="border-bottom:1px solid #000000">
		<input type="text" name="fal<%=tcount%>" value="<%=rs("faltas")%>" size="3" class=form_apt>
		<input type="hidden" name="afal<%=tcount%>" value="<%=rs("faltas")%>">
	</td>
	<td class="campor" style="border-bottom:1px solid #000000">
	<select size="1" name="jus<%=tcount%>" class=small>
		<option value=""  <%if rs("justificativa")="" then response.write "selected"%>></option>
		<option value="I" <%if rs("justificativa")="I" then response.write "selected"%>>Injustificada</option>
		<option value="J" <%if rs("justificativa")="J" then response.write "selected"%>>Just.-Abonar</option>
		<option value="D" <%if rs("justificativa")="D" then response.write "selected"%>>Just.-Descontar</option>
		<option value="A" <%if rs("justificativa")="A" then response.write "selected"%>>Atraso-Injust.-></option>
		<option value="B" <%if rs("justificativa")="B" then response.write "selected"%>>Atraso-Justif.-></option>
	</select>	
	<input type="hidden" name="ajus<%=tcount%>" value="<%=rs("justificativa")%>">
	</td>
	<td class="campor" style="border-bottom:1px solid #000000">
		<input type="text" name="atr<%=tcount%>" value="<%=rs("atraso")%>" size="3" class=form_apt>
		<input type="hidden" name="aatr<%=tcount%>" value="<%=rs("atraso")%>">
	</td>
	<td class="campor" style="border-bottom:1px solid #000000">
		<input type="text" name="ext<%=tcount%>" value="<%=rs("extra")%>" size="3" class=form_apt>
		<input type="hidden" name="aext<%=tcount%>" value="<%=rs("extra")%>">
	</td>
	<td class="campor" style="border-bottom:1px solid #000000">
		<input type="text" name="dep<%=tcount%>" value="<%=rs("dp")%>" size="3" class=form_apt>
		<input type="hidden" name="adep<%=tcount%>" value="<%=rs("dp")%>">
	</td>
	<td class="campor" style="border-bottom:1px solid #000000">
		<input type="text" name="rep<%=tcount%>" value="<%=rs("reposicao")%>" size="3" class=form_apt>
		<input type="hidden" name="arep<%=tcount%>" value="<%=rs("reposicao")%>">
	</td>
	<td class="campor" style="border-bottom:1px solid #000000">
		<input type="text" name="obs<%=tcount%>" value="<%=rs("observacao")%>" size="15" class=form_apt>
		<input type="hidden" name="aobs<%=tcount%>" value="<%=rs("observacao")%>">
	</td>
	<td class="campor" style="border-bottom:1px solid #000000" nowrap><font color="#0066cc"><%=horini%>-<%=horfim%></font></td>
    <td class="campor" style="border-bottom:1px solid #000000;border-left:1px solid #000000">
    <% if session("a38")="T" and isnull(rs("aulas")) then %>
	<input type="checkbox" name="excluir<%=tcount%>" value="1">
	<% end if %>
	</td>
<!-- marcacoes do dia -->
<%
	sqlcr="select chapa, day(data) as dia, data, batida, status from abatfun_m where " & _
	"chapa='" & rs("chapa") & "' and data='" & rs("dia_mes") & "' order by data, batida"
	sqlcr="select chapa, day(data) as dia, data, batida, status from abatfun_M where " & _
	"chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("dia_mes")) & "' order by batida"   'access
	sqlcr="select chapa, day(data) as dia, data, batida, status from corporerm.dbo.abatfun where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("dia_mes")) & "' order BY batida"  'sql
	if rs("descr")="Noite" then
		sqlcr="select top 8 chapa, day(data) as dia, data, hora from _catraca where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("dia_mes")) & "' order BY hora desc " 
	else 
		sqlcr="select top 8 chapa, day(data) as dia, data, hora from _catraca where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("dia_mes")) & "' order BY hora " 
	end if
	sqlcr="select chapa, day(data) as dia, data, batida, status from corporerm.dbo.abatfun where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("dia_mes")) & "' order BY batida"  'sql
	'marcações do chronus
	rs2.Open sqlcr, ,adOpenStatic, adLockReadOnly
	marcacao=0
	totalbatidas=rs2.recordcount
	for b=1 to 8:marc(b)="":formato(b)="":next
	if rs2.recordcount>0 then
		rs2.movefirst
		do while not rs2.eof
		'dia=rs2("dia")
		batida=formatdatetime((rs2("batida")/60)/24,4)
		'batida=formatdatetime(rs2("hora"),4)
		'if dia=diaant then 
		marcacao=marcacao+1 'else marcacao=1
		marc(marcacao)=batida
		'if rs2("status")="D" then formato(marcacao)="<font color='red'>" 'else formato(dia,marcacao)="<font color='black'>"
		'diaant=dia
		rs2.movenext
		loop
	else 'recordcount rs2
		for b=1 to 8
			marc(b)=""
		next
	end if 'recordcount rs2
	if marcacao<8 then
		for a=marcacao+1 to 8
			marc(a)=""
		next
	end if
	rs2.close
%>
	<td class="campor" style="border-bottom:1px solid #000000">
<%
for a=1 to 8
	response.write "<font color=blue>|</font>" & formato(a) & marc(a) & "</font>"
next 
%>	
<%
if totalbatidas=2 then
	if rs("horini")<>"" then
		i1=cdate(formatdatetime(rs("horini"),4))/24
		e1=cdate(marc(1))/24
		f1=cdate(formatdatetime(rs("horfim"),4))/24
		s1=cdate(marc(2))/24
		tolerancia=cdate("00:10:00")/24
		if e1>i1+tolerancia then response.write "<font color=blue>checar entrada</font>"
		if s1<f1-tolerancia then response.write "<font color=green>checar saida</font>"
	end if
end if
if totalbatidas=4 then

end if

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