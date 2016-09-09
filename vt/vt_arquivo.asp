<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a69")="N" or session("a69")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
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
sessao=session.sessionid

if request.form<>"" then
	conexao.execute "delete from ttvtransporte where sessao='" & sessao & "' "

	datamin=request.form("datamin")
	datamin=year(datamin) & numzero(month(datamin),2) & numzero(day(datamin),2)
	datamax=request.form("datamax")
	datamax=year(datamax) & numzero(month(datamax),2) & numzero(day(datamax),2)
	data1=year(now) & numzero(month(now),2) & numzero(day(now),2)
	data2=formatdatetime(now,2)
	data2=numzero(day(now),2) & "/" & numzero(month(now),2) & "/" & numzero(year(now),4)
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

'header
seq=1
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro,empresa) " & _
"select '" & sessao & "', '01', null, null, '" & "0" & "0100" & numzero(day(now),2) & numzero(month(now),2) & numzero(right(year(now),2),2) & _
"" & "73063166000120" & espaco2("FUNDACAO INSTITUTO DE ENSINO PARA OSASCO",60) & space(509) & numzero(seq,6) & "','VB' "
conexao.execute sql

'empresas e endereços
dim cnpj(2), ende(2), emp(2), tpe(2)
cnpj(0)="73063166000120":cnpj(1)="73063166000392":cnpj(2)="73063166000473"
emp(0)="01":emp(1)="03":emp(2)="04"
tpe(0)="1":tpe(1)="3":tpe(2)="3"
ende(0)="06020190" & espaco2("AV. FRANZ VOEGELLI",60) & espaco2("300",10) & espaco2("BLOCO BRANCO-5 ANDAR",40) & espaco2("IZABEL-RH",60)
ende(1)=ende(0):ende(2)=ende(0)
for a=0 to 0
	seq=seq+1
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro, empresa) " & _
"select '" & sessao & "', '02', '" & emp(a) & "', '01', '" & "1" & "010133519" & cnpj(a) & espaco2("FUNDACAO INSTITUTO DE ENSINO PARA OSASCO",60) & espaco2("UNIFIEO",40) & space(20) & space(20) & "  48" & _
"11" & "36519987  " & space(10) & espaco2(request.form("responsavel"),60) & "16  " & "02  " & "F" & "150688" & _
espaco2("izabel.rh@unifieo.br",50) & space(279) & numzero(seq,6) & "','VB' "
conexao.execute sql
next 
for a=0 to 2
	seq=seq+1
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro, empresa) " & _
"select '" & sessao & "', '02', '" & emp(a) & "', '02', '" & "2" & cnpj(a) & numzero(emp(a),4) & tpe(a) & ende(a) & _
space(396) & numzero(seq,6) & "','VB' "
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro, empresa) " & _
"select '" & sessao & "', '02', '" & emp(a) & "', '02', '" & "2" & cnpj(0) & numzero(emp(a),4) & tpe(a) & ende(a) & _
space(396) & numzero(seq,6) & "','VB' "
conexao.execute sql
next 

'funcionarios
	sql="SELECT F.CHAPA, unidade=left(codsecao,2), f.NOME, f.codsecao, depto=s.DESCRICAO, funcao=c.NOME, p.DTNASCIMENTO, p.CEP, " & _
"p.CPF, RG=p.CARTIDENTIDADE, p.DTEMISSAOIDENT, p.SEXO, m.MAE, m2.PAI, f.DATAADMISSAO, p.RUA, p.NUMERO, p.complemento, p.CIDADE, p.ESTADO, " & _
"ec=p.estadocivil, emissao=p.dtemissaoident " & _
"FROM corporerm.dbo.PFUNC F inner join corporerm.dbo.PPESSOA P on p.CODIGO=f.CODPESSOA " & _
"inner join corporerm.dbo.PSECAO s on s.CODIGO=f.CODSECAO inner join corporerm.dbo.PFUNCAO c on c.CODIGO=f.CODFUNCAO " & _
"left join QRY_MAE M on m.CHAPA=f.CHAPA left join qry_pai m2 on m2.CHAPA=f.chapa " & _
"WHERE F.DIASUTPROXMES>0 AND F.CODSITUACAO<>'D' and F.CHAPA collate database_default IN (SELECT CHAPA FROM TTVTCHAPA2 GROUP BY CHAPA) " & _
"ORDER BY left(codsecao,2), codsecao, F.chapa "
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalfu=rs.recordcount
rs.movefirst
do while not rs.eof
	seq=seq+1
	response.write " " & rs("chapa")
	if totalfu/30-int(totalfu/30)=0 then response.write "<br>"
	rg=replace(rs("rg"),".","")
	dig=InStr(rg,"-")
	if dig>0 then
		digrg=mid(rg,dig+1,1)
		rg=left(rg,dig-1)
	else
		rg=rg:digrg=""
	end if
	if isnull(rs("cep")) then cep=space(8) else cep=espaco2(textopuro(rs("cep"),2),8)
	if isnull(rs("pai")) then pai="" else pai=rs("pai")
	if isnull(rs("complemento")) then complemento="" else complemento=rs("complemento")
	nasc=numzero(day(rs("dtnascimento")),2) & numzero(month(rs("dtnascimento")),2) & numzero(right(year(rs("dtnascimento")),2),2)
	admissao=numzero(day(rs("dataadmissao")),2) & numzero(month(rs("dataadmissao")),2) & numzero(right(year(rs("dataadmissao")),4),4)
	if rs("emissao")<>"" then emissao=numzero(day(rs("emissao")),2) & numzero(month(rs("emissao")),2) & numzero(right(year(rs("emissao")),4),4) else emissao=""
	if rs("unidade")="01" then a=0 else if rs("unidade")="03" then a=1 else a=2
	if rs("ec")="C" then ec=2 else if rs("ec")="I" then ec=4 else if rs("ec")="V" then ec=5 else if rs("ec")="D" then ec=3 else ec=1
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro, empresa) " & _
"select '" & sessao & "', '03', '" & rs("unidade") & rs("codsecao") & "', '" & rs("chapa") & "', " & _
"'" & "3" & cnpj(0) & numzero(rs("unidade"),4) & numzero(rs("chapa"),15) & "001" & espaco2(rs("nome"),40) & _
espaco2(rs("depto"),40) & espaco2(rs("funcao"),30) & nasc & cep & ec & espaco2(textopuro(rs("cpf"),2),14) & _
espaco2(rg,14) & espaco2(emissao,8) & rs("sexo") & espaco2(rs("mae"),40) & espaco2(pai,40) & _
espaco2(admissao,8) & espaco2(rs("rua"),60) & espaco2(rs("numero"),10) & espaco2(complemento,40) & _
espaco2(ucase(rs("cidade")),60) & espaco2(rs("estado"),2) & espaco2(digrg,2) & space(133) & numzero(seq,6) & "','VB' "
	conexao.execute sql
rs.movenext
loop
rs.close

'itens de funcionarios
	sql="SELECT unidade=left(codsecao,2), f.CODSECAO, f.CHAPA, t.DESCRICAO, s.codvt, s2.descricao, t.VALOR, f.DIASUTPROXMES, " & _
"fv.NROVIAGENS, DIASUTPROXMES*nroviagens AS Viagens, DIASUTPROXMES*nroviagens*t.valor AS Totalvr " & _
"FROM ttvtchapa2 z INNER JOIN corporerm.dbo.PFUNC f ON f.CHAPA collate database_default=z.CHAPA " & _
"INNER JOIN corporerm.dbo.PFVALETR fv ON z.CODLINHA=fv.CODLINHA collate database_default AND z.CHAPA=fv.CHAPA collate database_default " & _
"INNER JOIN corporerm.dbo.PVALETR v ON fv.CODLINHA=v.CODIGO INNER JOIN corporerm.dbo.PTARIFA t ON v.CODTARIFA=t.CODIGO " & _
"INNER JOIN sPTARIFA s ON t.CODIGO collate database_default=s.CODIGO left join sptarifa_407 s2 on s.codvt=s2.codigo " & _
"WHERE f.DIASUTPROXMES>0 AND fv.NROVIAGENS>0 AND f.CODSITUACAO<>'D' AND fv.DTFIM>'" & dtaccess(datainicio) & "' " & _
"AND t.FINALVIGENCIA>'" & dtaccess(datainicio) & "' ORDER BY left(codsecao,2), f.CODSECAO, f.CHAPA; "
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalit=rs.recordcount:sequencia=1:totalvt=0
rs.movefirst
do while not rs.eof
	seq=seq+1
	if rs("codvt")="701" and (cdbl(rs("viagens"))*cdbl(rs("valor")))<10 then limitecmtc=limitecmtc & "<br>SPTRANS abaixo limite: " & rs("chapa") & "<br>"
	if rs("codvt")="701" and (cdbl(rs("viagens"))*cdbl(rs("valor")))>300 then limitecmtc=limitecmtc & "<br>SPTRANS acima limite: " & rs("chapa") & "<br>"
	unidade="0000" & left(rs("codsecao"),2)
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro, empresa) " & _
"select '" & sessao & "', '04', '" & rs("unidade") & rs("codsecao") & "', '" & rs("chapa") & "', " & _
"'" & "5" & cnpj(0) & numzero(rs("chapa"),15) & espaco2(rs("codvt"),15) & espaco2(rs("descricao"),60) & _
numzero(rs("viagens"),14) & numzero(replace(replace(formatnumber(rs("valor"),2),".",""),",",""),14) & "000000" & space(10) & "N" & numzero("01",6) & space(20) & _
"000" & "0" & "00" & "00000000" & space(404) & numzero(seq,6) & "','VB' "
	conexao.execute sql
	totalvt=totalvt+cdbl(rs("totalvr"))
	lastchapa=rs("chapa")
rs.movenext
loop
rs.close

sql="select valor from iParametros where parametro='vbadmin'"
rs.open sql, ,adOpenStatic, adLockReadOnly:perctaxa=cdbl(replace(rs("valor"),".",",")):rs.close
totaltaxa=int(totalvt * perctaxa + 0.5)/100

sql="select valor from iParametros where parametro='txentrega'"
rs.open sql, ,adOpenStatic, adLockReadOnly:totalentrega=cdbl(replace(rs("valor"),".",",")):rs.close
totalentrega=totalentrega

sql="select count(sessao) as totalreg from ttvtransporte where sessao='" & sessao & "' and campo1='02' and campo3='01' "
rs.Open sql, ,adOpenStatic, adLockReadOnly : treg1=rs("totalreg") : rs.close
sql="select count(sessao) as totalreg from ttvtransporte where sessao='" & sessao & "' and campo1='02' and campo3='02' "
rs.Open sql, ,adOpenStatic, adLockReadOnly : treg2=rs("totalreg") : rs.close
sql="select count(sessao) as totalreg from ttvtransporte where sessao='" & sessao & "' and campo1='03' "
rs.Open sql, ,adOpenStatic, adLockReadOnly : treg3=rs("totalreg") : rs.close
sql="select count(sessao) as totalreg from ttvtransporte where sessao='" & sessao & "' and campo1='04' "
rs.Open sql, ,adOpenStatic, adLockReadOnly : treg5=rs("totalreg") : rs.close
sql="select count(sessao) as totalreg from ttvtransporte where sessao='" & sessao & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly : treg=rs("totalreg")+1 : rs.close

	seq=seq+1
	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro,empresa) " & _
"select '" & sessao & "', '05', null, null, '9" & numzero(treg,6) & numzero(treg1,6) & numzero(treg2,6) & _
numzero(treg3,6) & "000000" & numzero(treg5,6) & space(557) & numzero(seq,6) & "','VB' "
conexao.execute sql

	sql="insert into ttvtransporte (sessao,campo1,campo2,campo3,registro,empresa) " & _
"select '" & sessao & "', '09', null, null, 'PEDIDO DE V.TRANSPORTE  " & data2 & _
hora1 & numzero(totalfu,5) & numzero(totalit,6) & _
numzero(replace(formatnumber(totaltaxa,2),".",""),16) & _
numzero(replace(formatnumber(totalvt,2),".",""),16) & _
data1 & space(1) & dataini & datafim & "00000000" & formatnumber(perctaxa,2) & "000000030,00" & "','VB' "
conexao.execute sql

	sql="select * from ttvtransporte where sessao='" & sessao & "' and empresa='VB' order by campo1, campo2, campo3"
end if
%>

<p class=titulo>Geração de arquivo de Vale Tranporte para pedido
<%
if request.form="" then
datamin=dateserial(year(now),month(now)+1,1)-1-9
datamax=dateserial(year(now),month(now)+1,1)-1-2
anofolha=year(dateserial(year(now),month(now)+1,1))
%>
<form method="POST" action="vt_arquivovb.asp">
	<p>Data mínima de Entrega: <input type="text" size="9" name="datamin" value="<%=datamin%>"><br>
	Data máxima de Entrega: <input type="text" size="9" name="datamax" value="<%=datamax%>"><br>
	Responsável Recebimento: <input type="text" size="20" name="responsavel" value="IZABEL"><br>
	</p>
	<p><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></p>
</form>
<% 
'datamax=year(datamax) & numzero(month(datamax),2) & numzero(day(datamax),2)
dataini=dateserial(year(datamax), month(datamax)+1,1)
datainicio=dataini

sql="DELETE FROM ttvtchapa":conexao.execute sql
sql="DELETE FROM ttvtchapa2":conexao.execute sql
sql="INSERT INTO ttvtchapa ( CHAPA ) SELECT CHAPA FROM corporerm.dbo.PFVALETR WHERE DTFIM>'" & dtaccess(datainicio) & "' GROUP BY CHAPA":response.write sql&"<br>":conexao.execute sql
sql="INSERT INTO ttvtchapa2 ( CHAPA, CODLINHA ) SELECT CHAPA, CODLINHA FROM corporerm.dbo.PFVALETR WHERE DTFIM>'" & dtaccess(datainicio) & "' GROUP BY CHAPA, CODLINHA":response.write sql:conexao.execute sql

else 

'Checagem de funcionários sem itens.
correcao1="UPDATE PFUNC SET DIASUTPROXMES=NULL WHERE CHAPA IN ("
sql1="SELECT it.campo3 AS chapa, fu.campo3 AS chapac " & _
"FROM (SELECT campo1, campo3 FROM ttvtransporte WHERE sessao='" & sessao & "' and campo1='03') fu " & _
"LEFT JOIN (SELECT campo1, campo3 FROM ttvtransporte WHERE sessao='" & sessao & "' and campo1='04') it ON fu.campo3 = it.campo3 " & _
"WHERE it.campo3 Is Null order by fu.campo3 "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
nrsql=1
if rs.recordcount>0 then
response.write "<br>ATENÇÃO!!!! Funcionários que não recebem vale-transporte no arquivo!!!!"
response.write "<br>SQL para correção:"
do while not rs.eof
	if nrsql>100 then sqlcorrecao=correcao1 & correcao & ")":correcao="":nrsql=0:response.write "<br>" & sqlcorrecao & "<br>"
	nrsql=nrsql+1
	correcao=correcao & "'" & rs("chapac") & "'"
	if nrsql<=100 then correcao=correcao & ","
rs.movenext:loop
end if
rs.close
sqlcorrecao=correcao1 & correcao & ")"
response.write "<br>" & sqlcorrecao & "<br>"
'checagem de itens sem funcionários
sql1="SELECT it.campo3 AS chapa " & _
"FROM (SELECT campo1, campo3 FROM ttvtransporte WHERE sessao='" & sessao & "' and campo1='03' and empresa='VB') fu " & _
"RIGHT JOIN (SELECT campo1, campo3 FROM ttvtransporte WHERE sessao='" & sessao & "' and campo1='04' and empresa='VB') it ON fu.campo3 = it.campo3 " & _
"WHERE fu.campo3 Is Null GROUP BY it.campo3"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
response.write "<br>ATENÇÃO!!!! Funcionários com problema no cadastro! Provavelmente sem nome de mãe."
do while not rs.eof
response.write "<br>" & rs("chapa")
rs.movenext
loop
end if
rs.close

'----------------------------------------
sql="select * from ttvtransporte where sessao='" & sessao & "' and empresa='VB' order by campo1, campo2, campo3"
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
		response.write "<td align="right" class="campor">&nbsp;" & conteudo & "&nbsp;</td>"
	else 
		conteudo=rs.fields(a)
		response.write "<td class="campor">&nbsp;" & conteudo & "</td>"
	end if
next
response.write "<td class="campor" align="right">" & len(conteudo) & "</td>"
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
	nomefile="vb" & textopuro(data2,2) & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	sql="select * from ttvtransporte where sessao='" & sessao & "' and campo1<>'09' and empresa='VB' order by campo1, campo2, campo3"
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