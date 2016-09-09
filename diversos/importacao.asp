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
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Rotinas Semestrais</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<script language="VBScript">
	Sub informacao(texto)
		document.form.campo.value=texto
	End Sub
</script>

<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set conexao2=server.createobject ("ADODB.Connection")
conexao2.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

%>
<p class=titulo>Importação de Horário</p>
<form name="form">
<input type="text" name="campo" size="80" class="form_input">
<!--
<br><input type="text" name="campo" size="10" class="form_apt" value="form_apt">
<br><input type="text" name="campo" size="10" class="form_box2" value="form_box2">
<br><input type="text" name="campo" size="10" class="form_box" value="form_box">
<br><input type="text" name="campo" size="10" class="form_input" value="form_input">
<br><input type="text" name="campo" size="10" class="form_input10" value="form_input10">
<br><input type="text" name="campo" size="10" class="form_ponto" value="form_ponto">
<br><input type="text" name="campo" size="10" class="help_input" value="help_input">
<br><input type="text" name="campo" size="10" class="pt8" value="pt8">
<br><input type="text" name="campo" size="10" class="pt9" value="pt9">
<br><input type="text" name="campo" size="10" class="pt10" value="pt10">
<br><input type="text" name="campo" size="10" class="proporcional" value="proporcional">
-->
</form>
<hr>
<%
inicio0=now()
%><script>Document.form.campo.value="1. Limpar tabelas temporárias"</script><%
sql1="delete from importacaohorario"
conexao.execute sql1
sql1="delete from importacaohorario2"
conexao.execute sql1
%><script>Document.form.campo.value="1. Limpar tabelas temporárias - Concluída"</script><%
for a=1 to 2000000:a=a:next

%><script>Document.form.campo.value="2. Gerar lançamentos temporários de horários"</script><%
sql2="INSERT INTO importacaohorario ( chapa, diasem, DIA, Ent1, Saida1, Ent2, Saida2, Ent3, Saida3, totalch ) " & _
"SELECT chapa, diasem, DIA, Ent1, Saida1, Ent2, Saida2, Ent3, Saida3, totalch FROM ttapontprof_2 " & _
"WHERE dia_mes Between '" & dtaccess(dateserial(year(now),month(now),1)) & "' And '" & dtaccess(dateserial(year(now),month(now)+1,1)-1) & "' " & _
"GROUP BY chapa, diasem, DIA, Ent1, Saida1, Ent2, Saida2, Ent3, Saida3, totalch "
conexao.execute sql2
%><script>Document.form.campo.value="2. Gerar lançamentos temporários de horários - Concluída"</script><%
for a=1 to 2000000:a=a:next

%><script>Document.form.campo.value="3. Completar os dias da semana"</script><%
sql3="select chapa from importacaohorario group by chapa"
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount:faltam=total-1
do while not rs.eof
	response.write "<script>Document.form.campo.value=""3. Completar os dias da semana: " & rs("chapa") & " - Faltam " & faltam &"""</script>"
	for a=1 to 7
		sql3a="select chapa, diasem from importacaohorario where chapa='" & rs("chapa") & "' and diasem=" & a
		rs2.Open sql3a, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then
		else
			sql3b="insert into importacaohorario (chapa, diasem, dia ) select '" & rs("chapa") & "', " & a & ", 0"
			conexao.execute sql3b
		end if
		rs2.close
	next
	faltam=faltam-1
rs.movenext
loop
rs.close
%><script>Document.form.campo.value="3. Completar os dias da semana - Concluída"</script><%
for a=1 to 2000000:a=a:next

%><script>Document.form.campo.value="4. Gerar lançamentos de horários"</script><%
sql4="INSERT INTO importacaohorario2 ( chapa, seq, tipo, codhorario ) " & _
"SELECT chapa, 0, '00', 'P' + chapa FROM importacaohorario " & _
"GROUP BY chapa "
conexao.execute sql4
sql4a="select chapa, codhorario from importacaohorario2 where tipo='00' "
rs.Open sql4a, ,adOpenStatic, adLockReadOnly
total=rs.recordcount:faltam=total-1
do while not rs.eof
	response.write "<script>Document.form.campo.value=""4. Criando Registro 00: " & rs("chapa") & " - Faltam " & faltam &"""</script>"
	turno=0 '0/1
	inativo=0 '0/1
	separador=""
	dthorario="28/01/2007"
	sql4b="update importacaohorario2 set registro='00" & espaco2(rs("codhorario"),10) & espaco2("Horário de Professor " & rs("chapa"),101) & dthorario & separador & turno & inativo & _
	"' where chapa='" & rs("chapa") & "' and tipo='00' "
	conexao.execute sql4b
	faltam=faltam-1
rs.movenext
loop
rs.close
%><script>Document.form.campo.value="4. Gerar lançamentos de horários - Concluída"</script><%
for a=1 to 2000000:a=a:next

%><script>Document.form.campo.value="5. Gerar lançamentos de horários"</script><%
sql5a="SELECT chapa, diasem, totalch FROM importacaohorario " & _
"WHERE diasem=1 AND totalch Is Null"
rs.Open sql5a, ,adOpenStatic, adLockReadOnly
total=rs.recordcount:faltam=total-1
do while not rs.eof
	response.write "<script>Document.form.campo.value=""5. Criando Registro 02: " & rs("chapa") & " - Faltam " & faltam &"""</script>"
	sql5b="insert into importacaohorario2 (tipo, seq, chapa, codhorario, registro ) " & _
	"select '02', 1, '" & rs("chapa") & "', 'P" & rs("chapa") & "', '02P" & rs("chapa") & "    1   00:00;24:00;'"
	conexao.execute sql5b
	sql5c="insert into importacaohorario2 (codhorario, seq, tipo, chapa, registro ) " & _
	"select 'P" & rs("chapa") & "', 8, 'LH', '" & rs("chapa") & "', " & _
	"'LHP" & rs("chapa") & "    " & espaco2(1,4) & espaco2("A",4) & "'" 
	conexao.execute sql5c
	faltam=faltam-1
rs.movenext
loop
rs.close
%><script>Document.form.campo.value="5. Gerar lançamentos de horários - Concluída"</script><%
for a=1 to 2000000:a=a:next

%><script>Document.form.campo.value="6. Gerar lançamentos de horários"</script><%
sql6a="SELECT chapa, diasem, totalch FROM importacaohorario " & _
"WHERE diasem<>1 AND totalch Is Null"
rs.Open sql6a, ,adOpenStatic, adLockReadOnly
total=rs.recordcount:faltam=total-1
do while not rs.eof
	response.write "<script>Document.form.campo.value=""6. Criando Registro 05: " & rs("chapa") & " - Faltam " & faltam &"""</script>"
	sql6b="insert into importacaohorario2 (tipo, seq, chapa, codhorario, registro ) " & _
	"select '05', " & rs("diasem") & ", '" & rs("chapa") & "', 'P" & rs("chapa") & "', '05P" & rs("chapa") & "    " & espaco2(rs("diasem"),4) & "'"
	conexao.execute sql6b
	faltam=faltam-1
rs.movenext
loop
rs.close
%><script>Document.form.campo.value="6. Gerar lançamentos de horários - Concluída"</script><%
for a=1 to 2000000:a=a:next

%><script>Document.form.campo.value="7. Gerar lançamentos de horários"</script><%
sql7a="SELECT chapa, diasem, ent1, saida1, ent2, saida2, ent3, saida3, totalch " & _
"FROM importacaohorario WHERE totalch Is Not Null"
rs.Open sql7a, ,adOpenStatic, adLockReadOnly
total=rs.recordcount:faltam=total-1:temp1=""
do while not rs.eof
	response.write "<script>Document.form.campo.value=""7. Criando Registro 01: " & rs("chapa") & " - Faltam " & faltam &"""</script>"
	if rs("ent1")<>"" then temp1=formatdatetime(rs("ent1"),4)&"E;"
	if rs("saida1")<>""  then temp1=temp1 & formatdatetime(rs("saida1"),4)&"S;"
	if rs("ent2")<>""  then temp1=temp1 & formatdatetime(rs("ent2"),4)&"E;"
	if rs("saida2")<>""  then temp1=temp1 & formatdatetime(rs("saida2"),4)&"S;"
	if rs("ent3")<>""  then temp1=temp1 & formatdatetime(rs("ent3"),4)&"E;"
	if rs("saida3")<>""  then temp1=temp1 & formatdatetime(rs("saida3"),4)&"S;"
	sql7b="insert into importacaohorario2 (codhorario, seq, tipo, chapa, registro ) " & _
	"select 'P" & rs("chapa") & "', " & rs("diasem") & ", '01', '" & rs("chapa") & "', " & _
	"'01P" & rs("chapa") & "    " & espaco2(rs("diasem"),4) & temp1 & "'" 
	'response.write sql7b
	'response.write "<br>" & rs.absoluteposition & " " & total &" - " & rs("chapa") & "<font size=1>" & sql7b & "</font>"
	conexao2.execute sql7b
	sql7c="insert into importacaohorario2 (codhorario, seq, tipo, chapa, registro ) " & _
	"select 'P" & rs("chapa") & "', " & 7+rs("diasem") &", 'LH', '" & rs("chapa") & "', " & _
	"'LHP" & rs("chapa") & "    " & espaco2(rs("diasem"),4) & espaco2(chr(64+rs("diasem")),4) & "'" 
	'conexao.execute sql7c
	faltam=faltam-1:temp1=""
rs.movenext
loop
rs.close
%><script>Document.form.campo.value="7. Gerar lançamentos de horários - Concluída"</script><%
for a=1 to 2000000:a=a:next




sql9="select registro from importacaohorario2 order by codhorario, seq, tipo"
'rs.Open sql9, ,adOpenStatic, adLockReadOnly
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
	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
response.write "<p>"
end if
'*************** fim teste **********************
'rs.close

	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="horarios.txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	sql="select registro from importacaohorario2 order by codhorario, seq, tipo"
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		registro=rs("registro")
		leitura.writeline registro
	rs.movenext
	loop
	rs.close
	termino=now()
	duracao=(termino-inicio0)
	Response.write "<p class=realce><font size=1> Inicio: " & inicio0 & " Termino: " & termino & " Duracao: " & formatdatetime(duracao,3) & "</font></p>"
	leitura.close
	set leitura=nothing
	set arquivo=nothing

%>
<a href="..\temp\<%=nomefile%>">Arquivo Horários</a>
</body>
</html>
<%

set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>