<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a70")="N" or session("a70")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Geração de Arquivo de Vale Transporte</title>
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
sessao=session.sessionid & "C"

if request.form<>"" then
	conexao.execute "delete from ttvtransporte where sessao='" & sessao & "' "

	datamin=request.form("datamin")
	datamin=year(datamin) & numzero(month(datamin),2) & numzero(day(datamin),2)
	datamax=request.form("datamax")
	datamax=year(datamax) & numzero(month(datamax),2) & numzero(day(datamax),2)
	data1=year(now) & numzero(month(now),2) & numzero(day(now),2)
	data2=formatdatetime(now,2)
	data2=numzero(day(now),2) & "/" & numzero(month(now),2) & "/" & numzero(right(year(now),2),2)
	hora1=numzero(hour(now),2) & "." &  numzero(minute(now),2) & "." &  numzero(second(now),2)
	dataini=dateserial(year(request.form("datamax")), month(request.form("datamax"))+1,1)
	datainicio=dataini
	datafim=dateserial(year(request.form("datamax")), month(request.form("datamax"))+2,1)-1
	dataini=year(dataini) & numzero(month(dataini),2) & numzero(day(dataini),2)
	datafim=year(datafim) & numzero(month(datafim),2) & numzero(day(datafim),2)
	
SQL="SELECT F.CHAPA, F.NOME, D.NOME AS MAE FROM corporerm.dbo.PFUNC AS F LEFT JOIN corporerm.dbo.PFDEPEND AS D ON F.CHAPA = D.CHAPA " & _
"WHERE F.DIASUTPROXMES>0 AND F.CODSITUACAO<>'D' AND D.NOME Is Null AND D.GRAUPARENTESCO='7' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
response.write "ATENÇÃO!!!! Funcionários sem mãe cadastrada!!!!"
do while not rs.eof
response.write "<br>" & rs("chapa") & "-" & rs("nome")
rs.movenext
loop
end if
rs.close
	
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
"select '" & sessao & "', '01', null, null, 'LSUP5" & espaco2(session("usuarioname"),8) & space(11) & data1 & hora1 & space(6) & space(261) & "' "
conexao.execute sql
'response.write "1"
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
"select '" & sessao & "', '02', null, null, 'T   A" & data1 & "V4.0" & space(60) & "' "
conexao.execute sql
'response.write "2"
sqlu="select left(codsecao,2) from ttvtcompl t, corporerm.dbo.pfunc f where f.chapa collate database_default=t.chapa and left(codsecao,2)='01' group by left(codsecao,2) "
rs.Open sqlu, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
"select '" & sessao & "', '04', '01', null, 'TTUN73.063.166/0001-20K." & data2 & "000001" & _
espaco2("NARCISO",26) & "AV  " & espaco2("FRANZ VOEGELLI",30) & "000300" & "5.ANDAR-RH" & espaco2("VILA YARA",15) & _ 
espaco2("OSASCO",25) & "06020-190SP" & datamin & datamax & espaco2(request.form("responsavel"),20) & _
"011365199050000S" & space(40) & "N" & "' "
conexao.execute sql
'response.write "3"
end if
rs.close

sqlu="select left(codsecao,2) from ttvtcompl t, corporerm.dbo.pfunc f where f.chapa collate database_default=t.chapa and left(codsecao,2)='03' group by left(codsecao,2) "
rs.Open sqlu, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
"select '" & sessao & "', '04', '03', null, 'TTUN73.063.166/0001-20K." & data2 & "000003" & _
espaco2("VILA YARA",26) & "AV  " & espaco2("FRANZ VOEGELLI",30) & "000300" & "5.ANDAR-RH" & espaco2("VILA YARA",15) & _ 
espaco2("OSASCO",25) & "06020-190SP" & datamin & datamax & espaco2(request.form("responsavel"),20) & _
"011365199050000S" & space(40) & "N" & "' "
conexao.execute sql
'response.write "4"
end if
rs.close

sqlu="select left(codsecao,2) from ttvtcompl t, corporerm.dbo.pfunc f where f.chapa collate database_default=t.chapa and left(codsecao,2)='04' group by left(codsecao,2) "
rs.Open sqlu, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
"select '" & sessao & "', '04', '04', null, 'TTUN73.063.166/0001-20K." & data2 & "000004" & _
espaco2("JD.WILSON",26) & "AV  " & espaco2("FRANZ VOEGELLI",30) & "000300" & "5.ANDAR-RH" & espaco2("VILA YARA",15) & _ 
espaco2("OSASCO",25) & "06020-190SP" & datamin & datamax & espaco2(request.form("responsavel"),20) & _
"011365199050000S" & space(40) & "N" & "' "
conexao.execute sql
'response.write "5"
end if
rs.close

sqlu="select codtipo from ttvtcompl t, corporerm.dbo.pfunc f where f.chapa collate database_default=t.chapa and codtipo='T' group by codtipo "
rs.Open sqlu, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
"select '" & sessao & "', '04', '90', null, 'TTUN73.063.166/0001-20K." & data2 & "000090" & _
espaco2("ESTAGIO",26) & "AV  " & espaco2("FRANZ VOEGELLI",30) & "000300" & "5.ANDAR-RH" & espaco2("VILA YARA",15) & _ 
espaco2("OSASCO",25) & "06020-190SP" & datamin & datamax & espaco2(request.form("responsavel"),20) & _
"011365199050000S" & space(40) & "N" & "' "
conexao.execute sql
'response.write "6"
end if
rs.close

	sql="SELECT f.CODTIPO, unidade=case when codtipo='T' then '90' else left(codsecao,2) end, f.CODSECAO, s.DESCRICAO " & _
"FROM corporerm.dbo.PFUNC f INNER JOIN corporerm.dbo.PSECAO s ON f.CODSECAO=s.CODIGO " & _
"WHERE f.CODSITUACAO<>'D' AND f.CHAPA collate database_default IN (SELECT CHAPA FROM TTVTcompl) " & _
"GROUP BY f.CODTIPO, case when codtipo='T' then '90' else left(codsecao,2) end, f.CODSECAO, s.DESCRICAO " & _
"ORDER BY f.CODTIPO, case when codtipo='T' then '90' else left(codsecao,2) end, f.CODSECAO "
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalde=rs.recordcount
rs.movefirst
do while not rs.eof
	unidade="0000" & left(rs("codsecao"),2)
	sql2="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
	"select '" & sessao & "', '05', '" & rs("unidade") & rs("codsecao") & "', null, 'TTDE73.063.166/0001-20K." & data2 & "0000" & rs("unidade") & _
	textopuro(rs("codsecao"),2) & espaco2(rs("descricao"),26) & space(20) & "73.063.166/0001-20" & "' "
	conexao.execute sql2
	'response.write "7"
rs.movenext
loop
rs.close
'response.write "<br>8"
	sql="SELECT F.CHAPA, unidade=case when codtipo='T' then '90' else left(codsecao,2) end, F.CODSECAO, F.NOME " & _
", P.CARTIDENTIDADE AS RG, P.CPF, P.DTNASCIMENTO, P.UFCARTIDENT, P.SEXO, M.MAE, P.RUA, P.NUMERO, P.COMPLEMENTO, " & _
"P.CIDADE, P.BAIRRO, P.CEP, P.ESTADO " & _
"FROM corporerm.dbo.PFUNC F, corporerm.dbo.PPESSOA P, QRY_MAE M " & _
"WHERE F.CODSITUACAO<>'D' and f.codpessoa=p.codigo and f.chapa=m.chapa " & _
"AND F.CHAPA collate database_default IN (SELECT CHAPA FROM TTVTcompl) " & _
"ORDER BY case when codtipo='T' then '90' else left(codsecao,2) end, F.NOME "
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalfu=rs.recordcount
rs.movefirst
do while not rs.eof
response.write "" & rs("chapa")
	unidade="0000" & left(rs("codsecao"),2)
	rg=textopuro(rs("rg"),2)
	if rg<>"" then rg=replace(rg,".","")
	if rg="" then rg=" "
	nasc=rs("dtnascimento")
	endereco=ucase(rs("rua"))
	if left(endereco,3)="AV." and left(endereco,2)<>"AV" and left(endereco,7)<>"AVENIDA" then
		endereco=replace(endereco,"AV.","AV   ")
	elseif left(endereco,2)="AV" and left(endereco,3)<>"AV." and left(endereco,7)<>"AVENIDA" then
		endereco=replace(endereco,"AV ","AV   ")
	elseif left(endereco,7)="AVENIDA" and left(endereco,3)<>"AV." and left(endereco,2)<>"AV" then
		endereco=replace(endereco,"AVENIDA ","AV   ")
	elseif left(endereco,3)="PCA" then
		endereco=replace(endereco,"PCA ","PÇ   ")
	elseif left(endereco,5)="PRACA" then
		endereco=replace(endereco,"PRACA ","PÇ   ")
	elseif left(endereco,7)="ESTRADA" then
		endereco=replace(endereco,"ESTRADA ","EST  ")
	elseif left(endereco,3)="RUA" then
		endereco=replace(endereco,"RUA ","RUA  ")
	else
		endereco="RUA" & space(2) & endereco
	end if
	bairro=rs("bairro")
	if not isnull(bairro) then bairro=replace(bairro,"'"," ") else bairro=""
	if isnull(rs("complemento")) then complemento="" else complemento=rs("complemento")
	if isnull(rs("cep")) then cep="" else cep=textopuro(rs("cep"),2)
	if isnull(rs("ufcartident")) then ufrg="" else ufrg=rs("ufcartident")
	nasc=year(nasc) & numzero(month(nasc),2) & numzero(day(nasc),2)
	sql2="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
	"select '" & sessao & "', '06', '" & rs("unidade") & rs("codsecao") & "', '" & rs("chapa") & "', " & _
	"'TTFU73.063.166/0001-20K." & data2 & "0000" & rs("unidade") & _
	textopuro(rs("codsecao"),2) & numzero(rs("chapa"),12) & espaco2(rs("nome"),30) & _
	espaco2(rg,10) & espaco2(textopuro(rs("cpf"),2),15) & nasc & espaco2(ufrg,2) & rs("sexo") & _
	espaco2(replace(rs("mae"),"'"," "),30) & espaco2(endereco,45) & numzero(rs("numero"),6) & espaco2(complemento,15) & _
	espaco2(rs("cidade"),40) & espaco2(bairro,30) & espaco2(cep,8) & espaco2(rs("estado"),2) & _
	space(14) & "' "
	conexao.execute sql2
rs.movenext
loop
rs.close

'response.write "<br>9"
	sql="SELECT unidade=case when codtipo='T' then '90' else left(codsecao,2) end, f.CODSECAO, f.CHAPA, t.DESCRICAO, s.vt_operadora, s.vt_bilhete, s.vt_tipo, s.vt_valor, t.VALOR, f.DIASUTPROXMES, fv.NROVIAGENS, DIASUTPROXMES*nroviagens AS Viagens, DIASUTPROXMES*nroviagens*t.valor AS Totalvr " & _
"FROM ttvtcompl a INNER JOIN (corporerm.dbo.PFUNC f INNER JOIN (((corporerm.dbo.PFVALETR fv INNER JOIN corporerm.dbo.PVALETR v ON fv.CODLINHA=v.CODIGO) INNER JOIN corporerm.dbo.PTARIFA t ON v.CODTARIFA=t.CODIGO) INNER JOIN sPTARIFA s ON t.CODIGO collate database_default=s.CODIGO) ON f.CHAPA collate database_default=fv.CHAPA) ON (a.CODLINHA=fv.CODLINHA collate database_default) AND (a.CHAPA=fv.CHAPA collate database_default) " & _
"WHERE f.DIASUTPROXMES>0 AND fv.NROVIAGENS>0 AND f.CODSITUACAO<>'D' AND fv.DTFIM>'" & dtaccess(datainicio) & "' " & _
"AND t.FINALVIGENCIA>'" & dtaccess(datainicio) & "' " & _
"ORDER BY case when codtipo='T' then '90' else left(codsecao,2) end, f.CODSECAO, f.CHAPA; "
	sql="SELECT unidade=case when codtipo='T' then '90' else left(codsecao,2) end, f.CODSECAO, f.CHAPA, t.DESCRICAO, s.vt_operadora, s.vt_bilhete, s.vt_tipo, s.vt_valor, t.VALOR, f.DIASUTPROXMES, a.QTVT AS Viagens, a.QTVT * t.VALOR AS Totalvr " & _
"FROM corporerm.dbo.PFUNC f INNER JOIN ttvtcompl a INNER JOIN corporerm.dbo.PVALETR v INNER JOIN corporerm.dbo.PTARIFA t ON v.CODTARIFA = t.CODIGO INNER JOIN SPTARIFA s ON t.CODIGO COLLATE database_default = s.CODIGO ON a.CODLINHA COLLATE database_default=v.CODIGO ON f.CHAPA = a.CHAPA COLLATE database_default " & _
"WHERE f.CODSITUACAO<>'D' AND t.FINALVIGENCIA>'" & dtaccess(datainicio) & "' " & _
"ORDER BY case when codtipo='T' then '90' else left(codsecao,2) end, f.CODSECAO, f.CHAPA; "
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalit=rs.recordcount:sequencia=1:totalvt=0
rs.movefirst
do while not rs.eof
	if (rs("vt_operadora")="CMTC" or rs("vt_operadora")="SPTRAN")  and (cdbl(rs("viagens"))*cdbl(rs("valor")))<10 then limitecmtc=limitecmtc & "<br>CMTC abaixo limite: " & rs("chapa") & "<br>"
	if (rs("vt_operadora")="CMTC" or rs("vt_operadora")="SPTRAN")  and (cdbl(rs("viagens"))*cdbl(rs("valor")))>300 then limitecmtc=limitecmtc & "<br>CMTC acima limite: " & rs("chapa") & "<br>"
	if lastchapa<>rs("chapa") then sequencia=1 else sequencia=sequencia
	if isnull(rs("vt_tipo")) then vt_tipo=space(4) else vt_tipo=espaco2(rs("vt_tipo"),4)
	unidade="0000" & left(rs("codsecao"),2)
	sql2="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
	"select '" & sessao & "', '07', '" & rs("unidade") & rs("codsecao") & "', '" & rs("chapa") & "', " & _
	"'TTIT73.063.166/0001-20K." & data2 & "0000" & rs("unidade") & _
	textopuro(rs("codsecao"),2) & numzero(rs("chapa"),12) & numzero(sequencia,3) & numzero(rs("viagens"),8) & _
	numzero(replace(replace(formatnumber(rs("valor"),2),".",""),",","."),9) & espaco2(rs("vt_operadora"),6) & _
	espaco2(rs("vt_bilhete"),12) & vt_tipo & "N" & "' "
	conexao.execute sql2
	sequencia=sequencia+1
	totalvt=totalvt+cdbl(rs("totalvr"))
	lastchapa=rs("chapa")
rs.movenext
loop
rs.close
perctaxa=2.78  'anterior 2.66
totaltaxa=int(totalvt * perctaxa + 0.5)/100
totalentrega=request.form("taxaentrega")

	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
"select '" & sessao & "', '03', null, null, 'TTPE73.063.166/0001-20K." & data2 & _
"0004" & numzero(totalde,4) & numzero(totalfu,5) & numzero(totalit,6) & _
numzero(replace(replace(formatnumber(totaltaxa,2),".",""),",","."),16) & _
numzero(replace(replace(formatnumber(totalvt,2),".",""),",","."),16) & _
data1 & space(1) & dataini & datafim & "00000000" & perctaxa & _
numzero(replace(replace(formatnumber(totalentrega,2),".",""),",","."),12) & _
"N" & "0001" & "AP%  R$ " & "' "
conexao.execute sql

	sql="select count(sessao) as totalreg from ttvtransporte where sessao='" & sessao & "' "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	totalreg=rs("totalreg")
	rs.close

	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
"select '" & sessao & "', '09', null, null, '9999" & numzero(totalreg,8) & space(152) & "' "
conexao.execute sql

	sql="select count(sessao) as totalarq from ttvtransporte where sessao='" & sessao & "' and campo1 not in ('01','02','09','10') "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	totalarq=rs("totalarq")
	rs.close

	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro) " & _
"select '" & sessao & "', '10', null, null, 'LSUP9" & "00000002" & "00000002" & numzero(totalarq,8) & space(277) & "' "
conexao.execute sql
	
	sql="select * from ttvtransporte where sessao='" & sessao & "' order by campo1, campo2, campo3"
end if
%>

<p class=titulo>Geração de arquivo de Vale Tranporte para pedido
<%
if request.form="" then
datamin=dateserial(year(now),month(now)+1,1)-1-9
datamax=dateserial(year(now),month(now)+1,1)-1-2
anofolha=year(dateserial(year(now),month(now)+1,1))
%>
<form method="POST" action="vt_arquivoc.asp">
	<p>Data mínima de Entrega: <input type="text" size="9" name="datamin" value="<%=datamin%>"><br>
	Data máxima de Entrega: <input type="text" size="9" name="datamax" value="<%=datamax%>"><br>
	Responsável Recebimento: <input type="text" size="20" name="responsavel" value="GRAZIELA"><br>
	Taxa de Entrega: R$ <input type="text" size="8" name="taxaentrega" value="115,11" class="vr">
	
	</p>
	<p><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></p>
</form>
<p class=titulo><b>Funcionários e bilhetes complementares</p>
<%
sqlc="select t.chapa, f.nome, t.codlinha, v.nomelinha as descricao, T.QTVT as quant from ttvtcompl t, corporerm.dbo.pfunc f, corporerm.dbo.pvaletr v " & _
"where t.chapa=f.chapa collate database_default and t.codlinha=v.codigo collate database_default ORDER BY t.chapa "
rs.Open sqlc, ,adOpenStatic, adLockReadOnly
rs.movefirst
response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse' >"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulo>" & ucase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=campo>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
rs.close
response.write "<p>"
%>
<% else %>
<%
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalf=0:totalg=0
rs.movefirst
response.write "<table border='1' cellpadding='1' cellspacing='1' style='border-collapse: collapse' >"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>&nbsp;" & rs.fields(a).name & "</td>"
next
response.write "<td class=titulor>Total</td>"
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	if rs.fields(a).type=5 then 
		conteudo=formatnumber(rs.fields(a),2) 
		response.write "<td align=""right"" class=""campor"">&nbsp;" & conteudo & "&nbsp;</td>"
	else 
		conteudo=rs.fields(a)
		response.write "<td class=""campor"">&nbsp;" & conteudo & "</td>"
	end if
next
response.write "<td class=""campor"" align=""right"">" & len(conteudo) & "</td>"
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
rs.close
response.write "<p>"
response.write limitecmtc
%>
Total de funcionarios: <%=totalfu%>
<br>Total do Pedido: <%=formatnumber(totalvt,2) %>
<br>Taxa de Serviço: <%=formatnumber(totaltaxa,2) %>
<br>Taxa de Entrega: <%=formatnumber(totalentrega,2) %>

<%
	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="vtk" & textopuro(data2,2) & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	sql="select * from ttvtransporte where sessao='" & sessao & "' order by campo1, campo2, campo3"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		leitura.writeline rs("registro")
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
<a href="..\temp\<%=nomefile%>">Arquivo Vale Transporte</a>
<%
end if 'request.form 
set rs=nothing
conexao.close
set conexao=nothing
%> 

</body>
</html>