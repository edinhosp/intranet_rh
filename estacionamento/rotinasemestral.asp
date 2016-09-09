<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a87")="N" or session("a87")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Rotinas Semestrais</title>
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
%>
<p class=titulo>Rotina Semestral para troca e renovação de crachás</p>
<form method="POST" action="rotinasemestral.asp" name="form">
<table border="0" bordercolor=black cellpadding="2" cellspacing="1" style="border-collapse: collapse" width=300>
<tr>
	<td class=titulo height=35>1. Cancelar demitidos</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R1">
</tr>
<tr>
	<td class=titulo height=35>2. Cancelar quem usa VT</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R2">
</tr>
<tr>
	<td class=titulo height=35>3. Criar período Administrativos</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R3">
</tr>
<tr>
	<td class=titulo height=35>4. Criar período Professores<Br>Mudança p/Grade Horária <br>&nbsp;&nbsp;&nbsp;(use cuidadosamente)</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R4">
</tr>
</table>

</form>
<hr>
<%
datager="'02/28/2016'"
if request.form("R1")<>"" then
	dtoper=formatdatetime(now,2)
	dtoper=dateserial(2016,2,29)
	sql1="select va.chapa, f.datademissao " & _
	"from veiculos_a va, corporerm.dbo.pfunc f where f.chapa collate database_default=va.chapa " & _
	"and (" & datager & " between inicio and termino) and f.codsituacao='D' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	response.write "Rotina 1: " & rs.recordcount & " ocorrências."
	if rs.recordcount>0 then
		do while not rs.eof
		sql2="select top 1 id_est, vy, ns, bp, cartao from veiculos_a where chapa='" & rs("chapa") & "' and " & datager & " between inicio and termino"
		rs2.Open sql2, ,adOpenStatic, adLockReadOnly
		id=rs2("id_est"):vy=cdbl(rs2("vy")):ns=cdbl(rs2("ns")):bp=cdbl(rs2("bp")):cartao=rs2("cartao")
		rs2.close
		sql3="update veiculos_a set pavy=" & vy & ", pans=" & ns & ", pabp=" & bp & ", cartao='" & cartao & "', vy=0,ns=0,bp=0"
		sql3="update veiculos_a set termino='" & dtaccess(rs("datademissao")) & "' where id_est=" & id
		conexao.execute sql3
		sql3="insert into veiculos_a (chapa,vy,ns,bp,inicio,termino,cartao,obs,pavy,pans,pabp,usuarioa,dataa) " & _
		"select '"&rs("chapa")&"',0,0,0,'"& dtaccess(dtoper)&"','"&dtaccess(dtoper)&"','"&cartao&"','demissão',"&vy&","&ns&","&bp&",'"&session("usuariomaster")&"','"&dtaccess(now)&"' "
		conexao.execute sql3
		rs.movenext
		loop
	end if
end if

if request.form("R2")<>"" then
	dtoper=formatdatetime(now,2)
	dtoper=dateserial(2016,2,29)
	sql1="select va.chapa, f.nome " & _
	"from veiculos_a va, corporerm.dbo.pfunc f where f.chapa collate database_default=va.chapa " & _
	"and (" & datager & " between inicio and termino) and f.codsituacao<>'D' and va.chapa in (select distinct chapa collate database_default from corporerm.dbo.pfvaletr where getdate() between dtinicio and dtfim)"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	response.write "Rotina 2: " & rs.recordcount & " ocorrências."
	if rs.recordcount>0 then
		do while not rs.eof
		sql2="select top 1 id_est, vy, ns, bp, cartao from veiculos_a where chapa='" & rs("chapa") & "' and " & datager & " between inicio and termino"
		rs2.Open sql2, ,adOpenStatic, adLockReadOnly
		id=rs2("id_est"):vy=cdbl(rs2("vy")):ns=cdbl(rs2("ns")):bp=cdbl(rs2("bp")):cartao=rs2("cartao")
		rs2.close
		sql3="update veiculos_a set pavy=" & vy & ", pans=" & ns & ", pabp=" & bp & ", cartao='" & cartao & "', vy=0,ns=0,bp=0"
		sql3="update veiculos_a set termino='" & dtaccess(cdate(dtoper)-1) & "' where id_est=" & id
		conexao.execute sql3
		sql3="insert into veiculos_a (chapa,vy,ns,bp,inicio,termino,cartao,obs,pavy,pans,pabp,usuarioa,dataa) " & _
		"select '"&rs("chapa")&"',0,0,0,'"&dtaccess(dtoper)&"','"&dtaccess(dtoper)&"','"&cartao&"','usa VT',"&vy&","&ns&","&bp&",'"&session("usuariomaster")&"','"&dtaccess(now)&"' "
		conexao.execute sql3
		rs.movenext
		loop
	end if
end if

if request.form("R3")<>"" then
	dtoper=formatdatetime(now,2)
	dtoper=dateserial(2016,2,28)
	termino=dateserial(2016,2,28)
	ninicio=dateserial(2016,3,1)
	ntermino=dateserial(2017,2,28)
	sql1="select va.chapa, f.nome, va.termino, vy, ns, bp, jw, id_est, cartao " & _
	"from veiculos_a va, corporerm.dbo.pfunc f where f.chapa collate database_default=va.chapa " & _
	"and cast(vy as integer)+ns+bp+jw>0 /*and va.status='A'*/ " & _
	"and termino='" & dtaccess(termino) & "' and f.codsituacao in ('A','F','Z') and f.codsindicato<>'03' " & _
	"order by va.chapa "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	response.write "Rotina 3: " & rs.recordcount & " ocorrências."
	
	if rs.recordcount>0 then
		do while not rs.eof
		id=rs("id_est"):vy=cdbl(rs("vy")):ns=cdbl(rs("ns")):bp=cdbl(rs("bp")):jw=cdbl(rs("jw")):cartao=rs("cartao")
		sql3="update veiculos_a set status='R' where id_est=" & id
		conexao.execute sql3
		sql3="insert into veiculos_a (chapa,vy,ns,bp,jw,inicio,termino,cartao,obs,pavy,pans,pabp,pajw,status,usuarioa,dataa) " & _
		"select '"&rs("chapa")&"',"&vy&","&ns&","&bp&","&jw&",'"&dtaccess(ninicio)&"','"&dtaccess(ntermino)&"','"&cartao&"','Lanc.Autom.',"&vy&","&ns&","&bp&","&jw&",'A','"&session("usuariomaster")&"','"&dtaccess(now)&"' "
		conexao.execute sql3
		rs.movenext
		loop
	end if
end if

if request.form("R4")<>"" then
	dtoper=formatdatetime(now,2)
	dtoper=dateserial(2016,2,28)
	termino=dateserial(2016,2,28)
	ninicio=dateserial(2016,3,1)
	ntermino=dateserial(2017,2,28)
	sql3="drop table grades_blocost "
	sql3="if exists (select 'True' from sysobjects where name='grades_blocost') drop table grades_blocost"
	conexao.execute sql3
	sql3="SELECT g.chapa1, g.NS, g.CO, g.AZ, g.AM, g.VE, g.LI, g.MA, g.BR, g.PR, " & _
	"(case when co is null then 0 else co end + case when az is null then 0 else az end + case when am is null then 0 else am end + " & _
	"case when ve is null then 0 else ve end + case when li is null then 0 else li end + case when ma is null then 0 else ma end) as t1 ," & _
	"(case when br is null then 0 else br end + case when pr is null then 0 else pr end) as t2 " & _
	"INTO grades_blocost FROM grades_blocos g"
	conexao.execute sql3

	sql1="select va.chapa, f.nome, va.termino, vy, ns, bp, jw, id_est, cartao " & _
	"from veiculos_a va, corporerm.dbo.pfunc f where f.chapa collate database_default=va.chapa " & _
	"and cast(vy as integer)+ns+bp+jw>0 /*and va.status='A'*/ " & _
	"and termino='" & dtaccess(termino) & "' and f.codsituacao in ('A','F','Z') and f.codsindicato='03' " & _
	"order by va.chapa "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	response.write "Rotina 4: " & rs.recordcount & " ocorrências."

	if rs.recordcount>0 then
		do while not rs.eof
		id=rs("id_est"):vy=cdbl(rs("vy")):ns=cdbl(rs("ns")):bp=cdbl(rs("bp")):jw=cdbl(rs("jw")):cartao=rs("cartao")
		sql3="update veiculos_a set status='R' where id_est=" & id
		conexao.execute sql3
		'checa narciso
		sql4="select ns from grades_blocost b where chapa1='" & rs("chapa") & "'"
		rs2.Open sql4, ,adOpenStatic, adLockReadOnly
			if rs2.recordcount>0 then
				narciso=rs2("ns")
				if narciso>0 then novons=-1 else novons=0
				if isnull(narciso) then novons=0
				if isnull(narciso) and ns=-1 then novons=-1
			else
				novons=ns
			end if
		rs2.close
		'checa brasil park e coral
		sql4="select t1, t2 from grades_blocost b where chapa1='" & rs("chapa") & "'"
		rs2.Open sql4, ,adOpenStatic, adLockReadOnly
			if rs2.recordcount>0 then
				bpark=rs2("t2"):coral=rs2("t1")
				if coral>bpark then novovy=-1 else novovy=0
				if bpark>coral then novobp=-1 else novobp=0
				if bpark>0 and coral>0 and bpark=coral then novovy=-1
			else
				novobp=bp:novovy=vy
			end if
		rs2.close
		'response.write "<br>" & rs("chapa")&"-" & rs("nome") & "-> ns: " & novons&"/"&ns & " bp: " & novobp&"/"&bp & " vy: " & novovy&"/"&vy
		'insere lançamento semestre
		sql3="insert into veiculos_a (chapa,vy,ns,bp,jw,inicio,termino,cartao,obs,pavy,pans,pabp,pajw,status,usuarioa,dataa) " & _
		"select '"&rs("chapa")&"',"&novovy&","&novons&","&novobp&","&jw&",'"&dtaccess(ninicio)&"','"&dtaccess(ntermino)&"','"&cartao&"','Lanc.Autom.',"&vy&","&ns&","&bp&","&jw&",'A','"&session("usuariomaster")&"','"&dtaccess(now)&"' "
		conexao.execute sql3

		rs.movenext
		loop
	end if

end if

'*************** inicio teste **********************
if request.form<>"" then
response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
response.write "<tr>"
for a=0 to rs.fields.count-1
	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
if rs.recordcount>0 then rs.movefirst
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=""campor"" nowrap>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
response.write "<p>"
end if
'*************** fim teste **********************
%>

</body>
</html>
<%

set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>