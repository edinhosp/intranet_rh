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
sequencia=1
' ************* linha 1 ***************
d01=numzero(sequencia,8)
d02="1"
d03="73063166000120"
d04="DIRF"
d05=request.form("anobase")
d06="O" 'O original
d07="1" '1 decl.normal
d08="2" '2 pj
d09="0" 'pj direito privado
d10="0" '0 nenhum c/imposto especial ou 1
d11=year(now)
d12="0"
d13=" "
d14=Espaco2("FUNDACAO INSTITUTO DE ENSINO PARA OSASCO", 60)
d15="00694967815"
d16=space(8)
d17=space(1) 
d18=space(42)
d19=space(12)
d20=space(229)
d21="18542005856"
d22=Espaco2("ROGERIO MATEUS DOS SANTOS ARAUJO", 60)
d23="0011"
d24="36519905"
d25="000000"
d26="00000000"
d27=Espaco2("rogerio@unifieo.br", 50)
d28=space(165)
d29=space(12) 'para uso do declarante
d30="9"
linha = d01 & d02 & d03 & d04 & d05 & d06 & d07 & d08 & d09 & d10
linha = linha & d11 & d12 & d13 & d14 & d15 & d16 & d17 & d18 & d19 & d20
linha = linha & d21 & d22 & d23 & d24 & d25 & d26 & d27 & d28 & d29 & d30
linha = linha '& d31
string_sql = "INSERT INTO autonomo_dirf (competencia, linha, linhadirf) " & _
"SELECT '" & d05 & "', " & sequencia & ", '" & linha & "'"
conexao.execute string_sql

' ************* linha 2 ***************
sql="SELECT a.nome_autonomo, a.cpf, Sum(r.desconto_ir) AS totir, " & _
"r01=sum(case when month(data_pagamento)=1 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d01=sum(case when month(data_pagamento)=1 then [desconto_inss] else 0 end), " & _
"i01=sum(case when month(data_pagamento)=1 then [desconto_ir] else 0 end), " & _
"r02=sum(case when month(data_pagamento)=2 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d02=sum(case when month(data_pagamento)=2 then [desconto_inss] else 0 end), " & _
"i02=sum(case when month(data_pagamento)=2 then [desconto_ir] else 0 end), " & _
"r03=sum(case when month(data_pagamento)=3 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d03=sum(case when month(data_pagamento)=3 then [desconto_inss] else 0 end), " & _
"i03=sum(case when month(data_pagamento)=3 then [desconto_ir] else 0 end), " & _
"r04=sum(case when month(data_pagamento)=4 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d04=sum(case when month(data_pagamento)=4 then [desconto_inss] else 0 end), " & _
"i04=sum(case when month(data_pagamento)=4 then [desconto_ir] else 0 end), " & _
"r05=sum(case when month(data_pagamento)=5 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d05=sum(case when month(data_pagamento)=5 then [desconto_inss] else 0 end), " & _
"i05=sum(case when month(data_pagamento)=5 then [desconto_ir] else 0 end), " & _
"r06=sum(case when month(data_pagamento)=6 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d06=sum(case when month(data_pagamento)=6 then [desconto_inss] else 0 end), " & _
"i06=sum(case when month(data_pagamento)=6 then [desconto_ir] else 0 end), " & _
"r07=sum(case when month(data_pagamento)=7 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d07=sum(case when month(data_pagamento)=7 then [desconto_inss] else 0 end), " & _
"i07=sum(case when month(data_pagamento)=7 then [desconto_ir] else 0 end), " & _
"r08=sum(case when month(data_pagamento)=8 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d08=sum(case when month(data_pagamento)=8 then [desconto_inss] else 0 end), " & _
"i08=sum(case when month(data_pagamento)=8 then [desconto_ir] else 0 end), " & _
"r09=sum(case when month(data_pagamento)=9 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d09=sum(case when month(data_pagamento)=9 then [desconto_inss] else 0 end), " & _
"i09=sum(case when month(data_pagamento)=9 then [desconto_ir] else 0 end), " & _
"r10=sum(case when month(data_pagamento)=10 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d10=sum(case when month(data_pagamento)=10 then [desconto_inss] else 0 end), " & _
"i10=sum(case when month(data_pagamento)=10 then [desconto_ir] else 0 end), " & _
"r11=sum(case when month(data_pagamento)=11 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d11=sum(case when month(data_pagamento)=11 then [desconto_inss] else 0 end), " & _
"i11=sum(case when month(data_pagamento)=11 then [desconto_ir] else 0 end), " & _
"r12=sum(case when month(data_pagamento)=12 then [servico_prestado]+[outros_rendimentos] else 0 end), " & _
"d12=sum(case when month(data_pagamento)=12 then [desconto_inss] else 0 end), " & _
"i12=sum(case when month(data_pagamento)=12 then [desconto_ir] else 0 end) " & _
"FROM autonomo AS a INNER JOIN autonomo_rpa AS r ON a.id_autonomo=r.id_autonomo  " & _
"WHERE Year(data_pagamento)=" & request.form("anobase") & " " & _
"GROUP BY a.nome_autonomo, a.cpf  " & _
"HAVING Sum(r.desconto_ir)>=0 "
rsc.Open sql, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
	sequencia=sequencia+1:e01=numzero(sequencia,8)
	sequencia=sequencia+1:f01=numzero(sequencia,8)
	e02="2"
	e03="73063166000120"
	e04="0588"
	e05="1" 'pessoa fisica
	e06="000" & espaco2(textopuro(rsc("cpf"),2),11)
	e07=espaco2(rsc("nome_autonomo"),60)
	e08=numzero(nrfile(formatnumber(rsc("r01"),2)),15): f08=numzero(nrfile(formatnumber(rsc("d01"),2)),15)
	e09=numzero(nrfile(formatnumber(0         ,2)),15): f09=numzero(nrfile(formatnumber(0         ,2)),15)
	e10=numzero(nrfile(formatnumber(rsc("i01"),2)),15): f10=numzero(nrfile(formatnumber(0         ,2)),15)

	e11=numzero(nrfile(formatnumber(rsc("r02"),2)),15): f11=numzero(nrfile(formatnumber(rsc("d02"),2)),15)
	e12=numzero(nrfile(formatnumber(0         ,2)),15): f12=numzero(nrfile(formatnumber(0         ,2)),15)
	e13=numzero(nrfile(formatnumber(rsc("i02"),2)),15): f13=numzero(nrfile(formatnumber(0         ,2)),15)

	e14=numzero(nrfile(formatnumber(rsc("r03"),2)),15): f14=numzero(nrfile(formatnumber(rsc("d03"),2)),15)
	e15=numzero(nrfile(formatnumber(0         ,2)),15): f15=numzero(nrfile(formatnumber(0         ,2)),15)
	e16=numzero(nrfile(formatnumber(rsc("i03"),2)),15): f16=numzero(nrfile(formatnumber(0         ,2)),15)

	e17=numzero(nrfile(formatnumber(rsc("r04"),2)),15): f17=numzero(nrfile(formatnumber(rsc("d04"),2)),15)
	e18=numzero(nrfile(formatnumber(0         ,2)),15): f18=numzero(nrfile(formatnumber(0         ,2)),15)
	e19=numzero(nrfile(formatnumber(rsc("i04"),2)),15): f19=numzero(nrfile(formatnumber(0         ,2)),15)

	e20=numzero(nrfile(formatnumber(rsc("r05"),2)),15): f20=numzero(nrfile(formatnumber(rsc("d05"),2)),15)
	e21=numzero(nrfile(formatnumber(0         ,2)),15): f21=numzero(nrfile(formatnumber(0         ,2)),15)
	e22=numzero(nrfile(formatnumber(rsc("i05"),2)),15): f22=numzero(nrfile(formatnumber(0         ,2)),15)

	e23=numzero(nrfile(formatnumber(rsc("r06"),2)),15): f23=numzero(nrfile(formatnumber(rsc("d06"),2)),15)
	e24=numzero(nrfile(formatnumber(0         ,2)),15): f24=numzero(nrfile(formatnumber(0         ,2)),15)
	e25=numzero(nrfile(formatnumber(rsc("i06"),2)),15): f25=numzero(nrfile(formatnumber(0         ,2)),15)

	e26=numzero(nrfile(formatnumber(rsc("r07"),2)),15): f26=numzero(nrfile(formatnumber(rsc("d07"),2)),15)
	e27=numzero(nrfile(formatnumber(0         ,2)),15): f27=numzero(nrfile(formatnumber(0         ,2)),15)
	e28=numzero(nrfile(formatnumber(rsc("i07"),2)),15): f28=numzero(nrfile(formatnumber(0         ,2)),15)

	e29=numzero(nrfile(formatnumber(rsc("r08"),2)),15): f29=numzero(nrfile(formatnumber(rsc("d08"),2)),15)
	e30=numzero(nrfile(formatnumber(0         ,2)),15): f30=numzero(nrfile(formatnumber(0         ,2)),15)
	e31=numzero(nrfile(formatnumber(rsc("i08"),2)),15): f31=numzero(nrfile(formatnumber(0         ,2)),15)

	e32=numzero(nrfile(formatnumber(rsc("r09"),2)),15): f32=numzero(nrfile(formatnumber(rsc("d09"),2)),15)
	e33=numzero(nrfile(formatnumber(0         ,2)),15): f33=numzero(nrfile(formatnumber(0         ,2)),15)
	e34=numzero(nrfile(formatnumber(rsc("i09"),2)),15): f34=numzero(nrfile(formatnumber(0         ,2)),15)

	e35=numzero(nrfile(formatnumber(rsc("r10"),2)),15): f35=numzero(nrfile(formatnumber(rsc("d10"),2)),15)
	e36=numzero(nrfile(formatnumber(0         ,2)),15): f36=numzero(nrfile(formatnumber(0         ,2)),15)
	e37=numzero(nrfile(formatnumber(rsc("i10"),2)),15): f37=numzero(nrfile(formatnumber(0         ,2)),15)

	e38=numzero(nrfile(formatnumber(rsc("r11"),2)),15): f38=numzero(nrfile(formatnumber(rsc("d11"),2)),15)
	e39=numzero(nrfile(formatnumber(0         ,2)),15): f39=numzero(nrfile(formatnumber(0         ,2)),15)
	e40=numzero(nrfile(formatnumber(rsc("i11"),2)),15): f40=numzero(nrfile(formatnumber(0         ,2)),15)

	e41=numzero(nrfile(formatnumber(rsc("r12"),2)),15): f41=numzero(nrfile(formatnumber(rsc("d12"),2)),15)
	e42=numzero(nrfile(formatnumber(0         ,2)),15): f42=numzero(nrfile(formatnumber(0         ,2)),15)
	e43=numzero(nrfile(formatnumber(rsc("i12"),2)),15): f43=numzero(nrfile(formatnumber(0         ,2)),15)

	e44=numzero(0,15)
	e45=numzero(0,15)
	e46=numzero(0,15)
	
	e47="0"
	e48="0":f48="1"
	e49=space(8)
	e50=space(32) 'para uso do declarante
	e51="9"
	partea1 = e01&e02&e03&e04&e05&e06&e07 
	partea2 = f01&e02&e03&e04&e05&e06&e07 
	parteb1= e08&e09&e10&e11&e12&e13&e14&e15&e16&e17&e18&e19&e20&e21&e22&e23&e24&e25&e26&e27&e28&e29&e30&e31&e32&e33&e34&e35&e36&e37&e38&e39&e40&e41&e42&e43
	parteb2= f08&f09&f10&f11&f12&f13&f14&f15&f16&f17&f18&f19&f20&f21&f22&f23&f24&f25&f26&f27&f28&f29&f30&f31&f32&f33&f34&f35&f36&f37&f38&f39&f40&f41&f42&f43
	partec1= e44&e45&e46&e47&e48&e49&e50&e51
	partec2= e44&e45&e46&e47&f48&e49&e50&e51
	linha1=partea1 & parteb1 & partec1
	linha2=partea1 & parteb2 & partec2
    string_sql = "INSERT INTO autonomo_dirf ( competencia, linha, linhadirf ) " & _
	"SELECT '" & d05 & "', " & (e01) & ", '" & linha1 & "'"
	conexao.execute string_sql
    string_sql = "INSERT INTO autonomo_dirf ( competencia, linha, linhadirf ) " & _
	"SELECT '" & d05 & "', " & (f01) & ", '" & linha2 & "'"
	conexao.execute string_sql
rsc.movenext
loop
rsc.close

sql=sql
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
		response.write "<td align="right" class="campor">&nbsp;" & conteudo & "</td>"
	else 
		conteudo=rs.fields(a)
		response.write "<td class="campor">&nbsp;" & conteudo & "</td>"
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
	nomefile="dirf" & ".099"
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