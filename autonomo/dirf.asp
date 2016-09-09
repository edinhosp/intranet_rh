<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a63")="N" or session("a63")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Geração de Arquivo DIRF para autônomos</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
inicio=now()
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	anobase=request.form("anobase")
    sql="DELETE FROM autonomo_dirf WHERE competencia='" & anobase & "' "
	conexao.execute sql
end if
%>
<p class=titulo>Geração de arquivo DIRF de autônomos</p>
<%
if request.form="" then
'mesfolha=month(dateserial(year(now),month(now)+1,1))
'anofolha=year(dateserial(year(now),month(now)+1,1))
%>
<form method="POST" action="dirf.asp" name="form">
<p>Ano-Base <select name="anobase">
<%
sqlmes="SELECT Year(data_pagamento) AS ano FROM autonomo_rpa GROUP BY Year(data_pagamento) order by Year(data_pagamento) desc "
rsc.Open sqlmes, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
if year(now)-1=rsc("ano") then tempc="selected" else tempc=""
%>
	<option value="<%=rsc("ano")%>" <%=tempc%>><%=rsc("ano")%></option>
<%
rsc.movenext
loop
rsc.close
%>
	</select></p>
<p><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></p>
</form>
<% 
else
anobase=request.form("anobase")
sequencia=1
' ************* linha 1 ***************
d01="Dirf"
d02=anobase+1
d03=anobase
d04="N" 'N-Original S-Retificadora
d05="" 'numzero(0,12) 'codigo recibo
d06="L35QJS2"

linha = d01 & chr(124) & d02 & chr(124) & d03 & chr(124) & d04 & chr(124) & d05 & chr(124) & d06 & chr(124)
string_sql = "INSERT INTO autonomo_dirf (competencia, linha, linhadirf) " & _
"SELECT '" & anobase & "', " & sequencia & ", '" & linha & "'"
conexao.execute string_sql

' ************* linha 2 ***************
sequencia=sequencia+1
d01="RESPO"
d02="18542005856"
d03="ROGERIO MATEUS DOS SANTOS ARAUJO"
d04="11"
d05="36519905"
d06="000000"
d07="36519987"
d08="rh@unifieo.br"

linha = d01 & chr(124) & d02 & chr(124) & d03 & chr(124) & d04 & chr(124) & d05 & chr(124) & d06 & chr(124) & d07 & chr(124) & d08 & chr(124)
string_sql = "INSERT INTO autonomo_dirf (competencia, linha, linhadirf) " & _
"SELECT '" & anobase & "', " & sequencia & ", '" & linha & "'"
conexao.execute string_sql

' ************* linha 3 ***************
sequencia=sequencia+1
d01="DECPJ"
d02="73063166000120"
d03="FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO"
d04="0" ' pj de direito privado
d05="00694967815"
d06="N" 'não é sócio ostensivo
d07="N" 'não é depositario de credito judicial
d08="N" 'não é instituição administradora de fundo
d09="N" 'não pagou rendimentos a residentes no exterior
d10="S" 'existe pagamento de valor por titular/dependente de plano de saude
d11="N" 'não existe pagamento relacionado a copa
d12="N" 'nao existe pagamento relacionado aos jogos olimpicos 2016
d13="N" 'não é declaração de extinção
d14="" 'space(8)

linha = d01 & chr(124) & d02 & chr(124) & d03 & chr(124) & d04 & chr(124) & d05 & chr(124) & d06 & chr(124) & d07 & chr(124) & d08 & chr(124) & d09 & chr(124) & d10 & chr(124) & d11 & chr(124) & d12 & chr(124) & d13 & chr(124) & d14 & chr(124)
string_sql = "INSERT INTO autonomo_dirf (competencia, linha, linhadirf) " & _
"SELECT '" & anobase & "', " & sequencia & ", '" & linha & "'"
conexao.execute string_sql

' ************* linha 4 ***************
sequencia=sequencia+1
d01="IDREC"
d02="0588"

linha = d01 & chr(124) & d02 & chr(124)
string_sql = "INSERT INTO autonomo_dirf (competencia, linha, linhadirf) " & _
"SELECT '" & anobase & "', " & sequencia & ", '" & linha & "'"
conexao.execute string_sql

' ************* linha 5 ***************
sql1="SELECT a.nome_autonomo, a.cpf " & _
"FROM autonomo AS a INNER JOIN autonomo_rpa AS r ON a.id_autonomo=r.id_autonomo  " & _
"WHERE Year(data_pagamento)=" & request.form("anobase") & " " & _
"GROUP BY a.nome_autonomo, a.cpf  " & _
"HAVING Sum(r.desconto_ir)>=0 " & _
"order by a.cpf " 
rsc.Open sql1, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
	'insere cpf
	sequencia=sequencia+1
	d01="BPFDEC"
	d02=espaco2(textopuro(rsc("cpf"),2),11)
	d03=rsc("nome_autonomo")
	d04=""
	linha = d01 & chr(124) & d02 & chr(124) & d03 & chr(124) & d04 & chr(124)
    string_sql = "INSERT INTO autonomo_dirf ( competencia, linha, linhadirf ) " & _
	"SELECT '" & anobase & "', " & sequencia & ", '" & linha & "'"
	conexao.execute string_sql

	'insere rendimentos
	sql="SELECT r01=sum(case when month(data_pagamento)=1 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"r02=sum(case when month(data_pagamento)=2  then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"r03=sum(case when month(data_pagamento)=3  then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"r04=sum(case when month(data_pagamento)=4  then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"r05=sum(case when month(data_pagamento)=5  then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"r06=sum(case when month(data_pagamento)=6  then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"r07=sum(case when month(data_pagamento)=7  then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"r08=sum(case when month(data_pagamento)=8  then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"r09=sum(case when month(data_pagamento)=9  then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"r10=sum(case when month(data_pagamento)=10 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"r11=sum(case when month(data_pagamento)=11 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"r12=sum(case when month(data_pagamento)=12 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
	"d01=sum(case when month(data_pagamento)=1  then [desconto_inss] else 0 end), " & _
	"d02=sum(case when month(data_pagamento)=2  then [desconto_inss] else 0 end), " & _
	"d03=sum(case when month(data_pagamento)=3  then [desconto_inss] else 0 end), " & _
	"d04=sum(case when month(data_pagamento)=4  then [desconto_inss] else 0 end), " & _
	"d05=sum(case when month(data_pagamento)=5  then [desconto_inss] else 0 end), " & _
	"d06=sum(case when month(data_pagamento)=6  then [desconto_inss] else 0 end), " & _
	"d07=sum(case when month(data_pagamento)=7  then [desconto_inss] else 0 end), " & _
	"d08=sum(case when month(data_pagamento)=8  then [desconto_inss] else 0 end), " & _
	"d09=sum(case when month(data_pagamento)=9  then [desconto_inss] else 0 end), " & _
	"d10=sum(case when month(data_pagamento)=10 then [desconto_inss] else 0 end), " & _
	"d11=sum(case when month(data_pagamento)=11 then [desconto_inss] else 0 end), " & _
	"d12=sum(case when month(data_pagamento)=12 then [desconto_inss] else 0 end), " & _
	"i01=sum(case when month(data_pagamento)=1  then [desconto_ir] else 0 end), " & _
	"i02=sum(case when month(data_pagamento)=2  then [desconto_ir] else 0 end), " & _
	"i03=sum(case when month(data_pagamento)=3  then [desconto_ir] else 0 end), " & _
	"i04=sum(case when month(data_pagamento)=4  then [desconto_ir] else 0 end), " & _
	"i05=sum(case when month(data_pagamento)=5  then [desconto_ir] else 0 end), " & _
	"i06=sum(case when month(data_pagamento)=6  then [desconto_ir] else 0 end), " & _
	"i07=sum(case when month(data_pagamento)=7  then [desconto_ir] else 0 end), " & _
	"i08=sum(case when month(data_pagamento)=8  then [desconto_ir] else 0 end), " & _
	"i09=sum(case when month(data_pagamento)=9  then [desconto_ir] else 0 end), " & _
	"i10=sum(case when month(data_pagamento)=10 then [desconto_ir] else 0 end), " & _
	"i11=sum(case when month(data_pagamento)=11 then [desconto_ir] else 0 end), " & _
	"i12=sum(case when month(data_pagamento)=12 then [desconto_ir] else 0 end) " & _
	"FROM autonomo AS a INNER JOIN autonomo_rpa AS r ON a.id_autonomo=r.id_autonomo  " & _
	"WHERE Year(data_pagamento)=" & anobase & " and a.cpf='" & rsc("cpf") & "' " & _
	"GROUP by a.cpf  " 
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	r01=nrfile(formatnumber(rs("r01"),2)): c01=nrfile(formatnumber(rs("d01"),2)): e01=nrfile(formatnumber(rs("i01"),2))
	r02=nrfile(formatnumber(rs("r02"),2)): c02=nrfile(formatnumber(rs("d02"),2)): e02=nrfile(formatnumber(rs("i02"),2))
	r03=nrfile(formatnumber(rs("r03"),2)): c03=nrfile(formatnumber(rs("d03"),2)): e03=nrfile(formatnumber(rs("i03"),2))
	r04=nrfile(formatnumber(rs("r04"),2)): c04=nrfile(formatnumber(rs("d04"),2)): e04=nrfile(formatnumber(rs("i04"),2))
	r05=nrfile(formatnumber(rs("r05"),2)): c05=nrfile(formatnumber(rs("d05"),2)): e05=nrfile(formatnumber(rs("i05"),2))
	r06=nrfile(formatnumber(rs("r06"),2)): c06=nrfile(formatnumber(rs("d06"),2)): e06=nrfile(formatnumber(rs("i06"),2))
	r07=nrfile(formatnumber(rs("r07"),2)): c07=nrfile(formatnumber(rs("d07"),2)): e07=nrfile(formatnumber(rs("i07"),2))
	r08=nrfile(formatnumber(rs("r08"),2)): c08=nrfile(formatnumber(rs("d08"),2)): e08=nrfile(formatnumber(rs("i08"),2))
	r09=nrfile(formatnumber(rs("r09"),2)): c09=nrfile(formatnumber(rs("d09"),2)): e09=nrfile(formatnumber(rs("i09"),2))
	r10=nrfile(formatnumber(rs("r10"),2)): c10=nrfile(formatnumber(rs("d10"),2)): e10=nrfile(formatnumber(rs("i10"),2))
	r11=nrfile(formatnumber(rs("r11"),2)): c11=nrfile(formatnumber(rs("d11"),2)): e11=nrfile(formatnumber(rs("i11"),2))
	r12=nrfile(formatnumber(rs("r12"),2)): c12=nrfile(formatnumber(rs("d12"),2)): e12=nrfile(formatnumber(rs("i12"),2))
	rs.close	
	v13=numzero(0,3)
	sequencia=sequencia+1
	d01="RTRT"
	linha = d01 & chr(124) & r01 & chr(124) & r02 & chr(124) & r03 & chr(124) & r04 & chr(124) & r05 & chr(124) & r06 & chr(124) & r07 & chr(124) & r08 & chr(124) & r09 & chr(124) & r10 & chr(124) & r11 & chr(124) & r12 & chr(124) & v13 & chr(124)
    string_sql = "INSERT INTO autonomo_dirf ( competencia, linha, linhadirf ) " & _
	"SELECT '" & anobase & "', " & sequencia & ", '" & linha & "'"
	conexao.execute string_sql
	sequencia=sequencia+1
	d01="RTPO"
	linha = d01 & chr(124) & c01 & chr(124) & c02 & chr(124) & c03 & chr(124) & c04 & chr(124) & c05 & chr(124) & c06 & chr(124) & c07 & chr(124) & c08 & chr(124) & c09 & chr(124) & c10 & chr(124) & c11 & chr(124) & c12 & chr(124) & v13 & chr(124)
    string_sql = "INSERT INTO autonomo_dirf ( competencia, linha, linhadirf ) " & _
	"SELECT '" & anobase & "', " & sequencia & ", '" & linha & "'"
	conexao.execute string_sql
	sequencia=sequencia+1
	d01="RTIRF"
	linha = d01 & chr(124) & e01 & chr(124) & e02 & chr(124) & e03 & chr(124) & e04 & chr(124) & e05 & chr(124) & e06 & chr(124) & e07 & chr(124) & e08 & chr(124) & e09 & chr(124) & e10 & chr(124) & e11 & chr(124) & e12 & chr(124) & v13 & chr(124)
    string_sql = "INSERT INTO autonomo_dirf ( competencia, linha, linhadirf ) " & _
	"SELECT '" & anobase & "', " & sequencia & ", '" & linha & "'"
	conexao.execute string_sql
	
rsc.movenext
loop
rsc.close

sequencia=sequencia+1
' ************* ultima linha  ***************
d01="FIMDirf"
linha = d01 & chr(124)
string_sql = "INSERT INTO autonomo_dirf (competencia, linha, linhadirf) " & _
"SELECT '" & anobase & "', " & sequencia & ", '" & linha & "'"
conexao.execute string_sql


sql=sql1
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalf=0:totalg=0
rs.movefirst
response.write "<table border='1' cellpadding='1' cellspacing='1' style='border-collapse: collapse' >"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>&nbsp;" & rs.fields(a).name & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	if rs.fields(a).type=5 then 
		conteudo=formatnumber(rs.fields(a),2) 
		response.write "<td align=""right"" class="""">&nbsp;" & conteudo & "</td>"
	else 
		conteudo=rs.fields(a)
		response.write "<td class=""campor"">&nbsp;" & conteudo & "</td>"
	end if
	'response.write "<td><font size='1'>&nbsp;" &rs.fields(a) & rs.fields(a).type & "</td>"
next
rs.movenext
loop
response.write "</table>"
rs.close
response.write "<p>"
%>

<%
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="dirf" & anobase &  ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	sql="SELECT linhadirf FROM autonomo_dirf " & _
	"WHERE competencia='" & request.form("anobase") & "' " & _
	"ORDER BY competencia, linha "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		leitura.writeline rs("linhadirf")
	rs.movenext
	loop
	rs.close
	termino=now()
	duracao=(termino-inicio)
	Response.write "<p class=realce><font size=1> Inicio: " & inicio & " Termino: " & termino & " Duracao: " & formatdatetime(duracao,3) & "</font></p>"
	leitura.close
	set leitura=nothing
	set arquivo=nothing

%>
<a href="..\temp\<%=nomefile%>">Arquivo DIRF <%=cmbmes%></a>
<%
end if 'request.form 
%> 

</body>

</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>