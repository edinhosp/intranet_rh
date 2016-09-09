<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")="N" or session("a48")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Custo da Folha</title>
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
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
--></script>

</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("Conexao2")

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form="" then
data=now()
data=dateserial(year(now),month(now)-1,1)
%>
<p class=titulo>Emissão de Custo da Folha Pag.&nbsp;<%=titulo %> - <font color="red"><b>BASE TESTE</b></font>
<form method="POST" action="custofolha_bteste.asp" name="form">
<table border="0" width="250" cellspacing="0"cellpadding="3">
<tr>
	<td class=titulo>Ano</td>
	<td class=titulo>Mês</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="anocomp" value="<%=year(data)%>" size="5"></td>
	<td class=fundo><input type="text" name="mescomp" value="<%=month(data)%>" size="3"></td>
</tr>
<tr>
	<td class=fundo colspan=2>Agrupa A/P? <input type="checkbox" name="agrupa" value="ON"></td>
</tr>
<tr>
	<td class=fundo colspan=2>Períodos: <input type="text" name="periodo" value="1,2,3,4,5,6,7,8,9,10,12,13" size="25"></td>
</tr>
</table>

<table border="0" width="250" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo><input type="submit" value="Visualizar" class="button" name="B1"></td>
</tr>
</table>
</form>
<%
else ' request.form<>""
anocomp=request.form("anocomp")
mescomp=request.form("mescomp")
sessao=session.sessionid
data=dateserial(anocomp,mescomp+1,1)
datam=dateserial(anocomp,mescomp-1,1)
mescompm=month(datam)
anocompm=year(datam)
periodo=request.form("periodo")
if request.form("agrupa")="ON" then agrupa=1 else agrupa=0
'	rs.Open sqlb, ,adOpenStatic, adLockReadOnly

sql1="delete from intranet_rh.dbo.grupo_custo_temp where sessao='" & sessao & "' "
conexao.execute sql1
sql1="delete from intranet_rh.dbo.grupo_custo_table where sessao='" & sessao & "' "
conexao.execute sql1

'**** grupos mes
sql2="INSERT INTO intranet_rh.dbo.grupo_custo_temp (Sessao, Ordem, Grupo, dia, codevento, codsin, ANOCOMP, MESCOMP, custo, Campus ) " & _
"SELECT '" & sessao & "', g.Ordem, g.Grupo, g.dia, g.CODIGO, f.CODSINDICATO, ff.ANOCOMP, ff.MESCOMP, Sum(VALOR*fator) AS custo, Left(CODSECAO,2) AS Campus " & _
"FROM grupo_custo g, corporerm_teste1.dbo.pffinanc ff, corporerm_teste1.dbo.PFUNC f WHERE ff.CHAPA=f.CHAPA and g.CODIGO=ff.CODEVENTO collate database_default " & _
"AND g.excecao='0' " & _
"and ff.nroperiodo in (" & periodo & ") " & _
"GROUP BY g.Ordem, g.Grupo, g.dia, g.CODIGO, f.CODSINDICATO, ff.ANOCOMP, ff.MESCOMP, Left(CODSECAO,2) " & _
"HAVING ff.ANOCOMP=" & anocomp & " AND ff.MESCOMP=" & mescomp & "; "
conexao.execute sql2

diasem=weekday(data)
if diasem<7 then dias=7-diasem
datair=data+dias+4
datair=dateserial(year(data),month(data)+1,10)
if weekday(datair)=1 then datair=datair-2
if weekday(datair)=7 then datair=datair-1

'*** IR RESCISAO
sql6="INSERT INTO grupo_custo_temp (Sessao, Ordem, Grupo, Campus, codsin, custo, ANOCOMP, MESCOMP, data, dia ) " & _
"SELECT '" & sessao & "', 8, 'IMPOSTO DE RENDA RESC', Left(CODSECAO,2) AS CAMPUS, pf.CODSINDICATO, Sum(valor) AS liquido, f.ANOCOMP, f.MESCOMP, f.DTPAGTO, 3 " & _
"FROM corporerm_teste1.dbo.pffinanc f, corporerm_teste1.dbo.pevento e, corporerm_teste1.dbo.pfunc pf " & _
"WHERE f.CODEVENTO=e.codigo AND pf.CHAPA=f.chapa AND e.PROVDESCBASE In ('P','D') " & _
"AND f.ANOCOMP=" & anocomp & " AND f.MESCOMP=" & mescomp & " AND f.codevento='312' and pf.codsituacao='D' " & _
"and f.nroperiodo in (" & periodo & ") " & _
"GROUP BY pf.CODTIPO, Left(CODSECAO,2), pf.CODSINDICATO, f.ANOCOMP, f.MESCOMP, f.DTPAGTO, day(f.dtpagto) " 
conexao.execute sql6

'**** grupos mes ant
sql3="INSERT INTO grupo_custo_temp (sessao, Ordem, Grupo, dia, codevento, codsin, ANOCOMP, MESCOMP, custo, Campus ) " & _
"SELECT '" & sessao & "', g.Ordem, g.Grupo, g.dia, g.CODIGO, f.CODSINDICATO, ff.ANOCOMP, ff.MESCOMP, Sum(VALOR*fator) AS custo, Left(CODSECAO,2) AS Campus " & _
"FROM grupo_custo g, corporerm_teste1.dbo.pffinanc ff, corporerm_teste1.dbo.PFUNC f WHERE ff.CHAPA=f.CHAPA AND g.CODIGO=ff.CODEVENTO collate database_default " & _
"AND g.excecao='1' " & _
"and ff.nroperiodo in (" & periodo & ") " & _
"GROUP BY g.Ordem, g.Grupo, g.dia, g.CODIGO, f.CODSINDICATO, ff.ANOCOMP, ff.MESCOMP, Left(CODSECAO,2) " & _
"HAVING ff.ANOCOMP=" & anocompm & " AND ff.MESCOMP=" & mescompm & "; "
conexao.execute sql3

if month(data)=2 then dia="28" else dia="dia"

sql3="update grupo_custo_temp set data='" & year(data) & "/" & month(data) & "/'+ cast(" & dia & " as char) where sessao='" & sessao & "' "
conexao.execute sql3
sql3="update grupo_custo_temp set data=(case when datepart(dw,data)=1 then data+1 else case when datepart(dw,data)=7 then data+2 else data end end) where ordem in (4,5,6,7) " & _
"and sessao='" & sessao & "' "
conexao.execute sql3
sql3="update grupo_custo_temp set data='" & dtaccess(datair) & "' where sessao='" & sessao & "' and grupo IN ('IMPOSTO DE RENDA','IMPOSTO DE RENDA RESC') "
conexao.execute sql3
sql3="update grupo_custo_temp set data=(case when datepart(dw,data)=1 then data-2 else case when datepart(dw,data)=7 then data-1 else data end end) where ordem in (1,2,3,20) " & _
"and sessao='" & sessao & "' "
conexao.execute sql3

'***** fgts
datafgts=dateserial(anocomp,mescomp+1,7)
diasem=weekday(datafgts):dias=0
if diasem=7 then dias=2
if diasem=1 then dias=1
datafgts=datafgts+dias
sql4="INSERT INTO grupo_custo_temp (sessao, Ordem, Grupo, dia, data, codsin, ANOCOMP, MESCOMP, custo, Campus ) " & _
"SELECT '" & sessao & "', 10, 'FGTS', 7, '" & dtaccess(datafgts) & "', f.CODSINDICATO, ff.ANOCOMP, ff.MESCOMP, " & _
"sum(round( ((case when basefgts<0 then 0 else basefgts end)+(case when basefgts13<0 then 0 else basefgts13 end)) *8.0+0,0)/100), Left(CODSECAO,2) AS CAMPUS " & _
"FROM corporerm_teste1.dbo.PFPERFF ff, corporerm_teste1.dbo.PFUNC f WHERE ff.CHAPA=f.CHAPA AND f.CODTIPO='N' " & _
"and ff.nroperiodo in (" & periodo & ") " & _
"GROUP BY f.CODSINDICATO, ff.ANOCOMP, ff.MESCOMP, Left(CODSECAO,2) " & _
"HAVING ff.ANOCOMP=" & anocomp & " AND ff.MESCOMP=" & mescomp & "; "
conexao.execute sql4

'***** pis
datapis=dateserial(anocomp,mescomp+1,15)
diasem=weekday(datapis):dias=0
if diasem=7 then dias=-1
if diasem=1 then dias=-2
datapis=datapis+dias
sql5="INSERT INTO grupo_custo_temp (Sessao, Ordem, Grupo, dia, data, codsin, ANOCOMP, MESCOMP, custo, Campus ) " & _
"SELECT '" & sessao & "', 11, 'PIS', 15, '" & dtaccess(datapis) & "', f.CODSINDICATO, ff.ANOCOMP, ff.MESCOMP, " & _
"sum(round((((case when basefgts<0 then 0 else basefgts end)+(case when basefgts13<0 then 0 else basefgts13 end))*1)+1/2,0)/100), Left(CODSECAO,2) AS CAMPUS " & _
"FROM corporerm_teste1.dbo.PFPERFF ff, corporerm_teste1.dbo.PFUNC f WHERE ff.CHAPA=f.CHAPA AND f.CODTIPO='N' " & _
"and ff.nroperiodo in (" & periodo & ") " & _
"GROUP BY f.CODSINDICATO, ff.ANOCOMP, ff.MESCOMP, Left(CODSECAO,2) " & _
"HAVING ff.ANOCOMP=" & anocomp & " AND ff.MESCOMP=" & mescomp & "; "
conexao.execute sql5

'*** LIQUIDO ESTAG
sql6="INSERT INTO grupo_custo_temp (Sessao, Ordem, Grupo, Campus, codsin, custo, ANOCOMP, MESCOMP, data, dia ) " & _
"SELECT '" & sessao & "', 12, 'LIQUIDO ESTAGIÁRIOS', Left(CODSECAO,2) AS CAMPUS, pf.CODSINDICATO, sum(case when provdescbase='D' then -1 else 1 end * case when valor<0 then 0 else valor end) as liquido, f.ANOCOMP, f.MESCOMP, f.DTPAGTO, 3 " & _
"FROM corporerm_teste1.dbo.pffinanc f, corporerm_teste1.dbo.pevento e, corporerm_teste1.dbo.pfunc pf " & _
"WHERE f.CODEVENTO=e.codigo AND pf.CHAPA=f.chapa AND e.PROVDESCBASE In ('P','D') " & _
"and f.nroperiodo in (" & periodo & ") " & _
"GROUP BY pf.CODTIPO, Left(CODSECAO,2), pf.CODSINDICATO, f.ANOCOMP, f.MESCOMP, f.DTPAGTO " & _
"HAVING f.ANOCOMP=" & anocomp & " AND f.MESCOMP=" & mescomp & " " & _
"AND pf.CODTIPO='T' AND sum(case when provdescbase='D' then -1 else 1 end * case when valor<0 then 0 else valor end)>0 "
conexao.execute sql6

'**** LIQUIDO FOLHA
sql7="INSERT INTO grupo_custo_temp (Sessao, Ordem, Grupo, Campus, codsin, custo, ANOCOMP, MESCOMP, data, dia ) " & _
"SELECT '" & sessao & "', 13, 'LIQUIDO FOLHA', Left(CODSECAO,2) AS CAMPUS, pf.CODSINDICATO, sum(case when provdescbase='D' then -1 else 1 end * case when valor<0 then 0 else valor end) as liquido, f.ANOCOMP, f.MESCOMP, f.DTPAGTO, 3 " & _
"FROM corporerm_teste1.dbo.pffinanc AS f, corporerm_teste1.dbo.pevento AS e, corporerm_teste1.dbo.pfunc AS pf " & _
"WHERE f.CODEVENTO=e.codigo AND pf.CHAPA=f.chapa AND e.PROVDESCBASE In ('P','D') " & _
"and f.nroperiodo in (" & periodo & ") " & _
"GROUP BY pf.CODTIPO, Left(CODSECAO,2), pf.CODSINDICATO, f.ANOCOMP, f.MESCOMP, f.dtpagto " & _
"HAVING f.ANOCOMP=" & anocomp & " AND f.MESCOMP=" & mescomp & " " & _
"AND pf.CODTIPO='N' AND sum(case when provdescbase='D' then -1 else 1 end * case when valor<0 then 0 else valor end)<>0 "
conexao.execute sql7
sql7a="select data, count(data) as freq from grupo_custo_temp g " & _
"where ordem=13 and sessao='" & sessao & "' group by data having count(data)>1"
rs.Open sql7a, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then datacerta=rs("data")
rs.close
sql7b="update grupo_custo_temp set data=#" & dtaccess(datacerta) & "# " & _
"where ordem=13 and sessao='" & sessao & "' "
'conexao.execute sql7b

'**** CREDITO FERIAS
sql8="INSERT INTO grupo_custo_temp (sessao, Ordem, Grupo, codsin, Campus, data, custo ) " & _
"SELECT '" & sessao & "', 14, 'CREDITO FÉRIAS', f.CODSINDICATO, Left(CODSECAO,2) AS Campus, pf.DTPAGTO, sum(case when provdescbase='D' then -1 else 1 end * case when valor<0 then 0 else valor end) as liquido " & _
"FROM corporerm_teste1.dbo.pfperfer_old pf, corporerm_teste1.dbo.pfferias_old fer, corporerm_teste1.dbo.PEVENTO e, corporerm_teste1.dbo.PFUNC f WHERE pf.chapa=f.chapa AND pf.nroperiodo=fer.nroperiodo AND pf.dtvencimento=fer.dtvencimento AND pf.chapa=fer.chapa AND fer.codevento=e.codigo " & _
"and (dtaviso+30)>=DATEADD(m,1,'" & anocomp & "-" & mescomp & "-01') and (dtaviso+30)<=(dateadd(m,2,'" & anocomp & "-" & mescomp & "-01')-1) AND e.PROVDESCBASE In ('P','D') " & _
"GROUP BY f.CODSINDICATO, Left(CODSECAO,2), pf.DTPAGTO ORDER BY pf.DTPAGTO "
sql8="INSERT INTO grupo_custo_temp (sessao, Ordem, Grupo, codsin, Campus, data, custo ) " & _
"SELECT '" & sessao & "', 14, 'CREDITO FÉRIAS', f.CODSINDICATO, Left(CODSECAO,2) AS Campus, 'DTPAGTO'=pf.DaTaPAGTO, sum(case when provdescbase='D' then -1 else 1 end * case when valor<0 then 0 else valor end) as liquido " & _
"FROM corporerm_teste1.dbo.PFUFERIASPER pf, corporerm_teste1.dbo.PFUFERIASVERBAS fer, corporerm_teste1.dbo.PEVENTO e, corporerm_teste1.dbo.PFUNC f " & _
"WHERE pf.chapa=f.chapa AND pf.DATAPAGTO=fer.datapagto AND pf.fimperaquis=fer.FIMPERAQUIS AND pf.chapa=fer.chapa AND fer.codevento=e.codigo " & _
"and (dataaviso+30)>=DATEADD(m,1,'" & anocomp & "-" & mescomp & "-01') and (dataaviso+30)<=(dateadd(m,2,'" & anocomp & "-" & mescomp & "-01')-1) AND e.PROVDESCBASE In ('P','D') " & _
"GROUP BY f.CODSINDICATO, Left(CODSECAO,2), pf.DaTaPAGTO ORDER BY pf.DaTaPAGTO "
conexao.execute sql8

'****** IR FERIAS
sql9="INSERT INTO grupo_custo_temp (sessao, Ordem, Grupo, codsin, Campus, data, custo, codevento ) " & _
"SELECT '" & sessao & "', 15, 'IR SOBRE FÉRIAS', f.CODSINDICATO, Left(CODSECAO,2) AS Campus, pf.DTPAGTO, Sum(fer.VALOR) as vr, fer.CODEVENTO " & _
"FROM corporerm_teste1.dbo.pfperfer_old pf, corporerm_teste1.dbo.pfferias_old fer, corporerm_teste1.dbo.PEVENTO e, corporerm_teste1.dbo.PFUNC f WHERE pf.CHAPA=f.CHAPA AND pf.nroperiodo=fer.nroperiodo AND pf.dtvencimento=fer.dtvencimento AND pf.chapa=fer.chapa AND fer.codevento=e.codigo " & _
"and (dtaviso+30)>=DATEADD(m,1,'" & anocomp & "-" & mescomp & "-01') and (dtaviso+30)<=(dateadd(m,2,'" & anocomp & "-" & mescomp & "-01')-1) AND e.PROVDESCBASE In ('P','D') " & _
"GROUP BY f.CODSINDICATO, Left(CODSECAO,2), pf.DTPAGTO, fer.CODEVENTO " & _
"HAVING fer.CODEVENTO='312' ORDER BY pf.DTPAGTO "
sql9="INSERT INTO grupo_custo_temp (sessao, Ordem, Grupo, codsin, Campus, data, custo, codevento ) " & _
"SELECT '" & sessao & "', 15, 'IR SOBRE FÉRIAS', f.CODSINDICATO, Left(CODSECAO,2) AS Campus, pf.DaTaPAGTO, Sum(fer.VALOR) as vr, fer.CODEVENTO " & _
"FROM corporerm_teste1.dbo.pfuferiasper pf, corporerm_teste1.dbo.pfuferiasverbas fer, corporerm_teste1.dbo.PEVENTO e, corporerm_teste1.dbo.PFUNC f " & _
"WHERE pf.CHAPA=f.CHAPA AND pf.datapagto=fer.datapagto AND pf.fimperaquis=fer.fimperaquis AND pf.chapa=fer.chapa AND fer.codevento=e.codigo " & _
"and (dataaviso+30)>=DATEADD(m,1,'" & anocomp & "-" & mescomp & "-01') and (dataaviso+30)<=(dateadd(m,2,'" & anocomp & "-" & mescomp & "-01')-1) AND e.PROVDESCBASE In ('P','D') " & _
"GROUP BY f.CODSINDICATO, Left(CODSECAO,2), pf.DaTaPAGTO, fer.CODEVENTO " & _
"HAVING fer.CODEVENTO='312' ORDER BY pf.DaTaPAGTO "
conexao.execute sql9

sql10="update grupo_custo_temp set data=data+(7-weekday(data))+4 where sessao='" & sessao & "' and ordem=15 "
sql10="update grupo_custo_temp set data=dateadd(m,1,cast(year(data) as char)+'/'+cast(month(data) as char)+'/10') where sessao='" & sessao & "' and ordem=15 "
conexao.execute sql10

sql11="SELECT sessao, codsin, data, Ordem, Grupo, " & _
"sum(case when campus='01' then custo else 0 end) as '01', sum(case when campus='02' then custo else 0 end) as '02', " & _
"sum(case when campus='03' then custo else 0 end) as '03', sum(case when campus='04' then custo else 0 end) as '04' " & _
"FROM grupo_custo_temp WHERE sessao='" & sessao & "' " & _
"GROUP BY sessao, codsin, data, Ordem, Grupo " & _
"ORDER BY data, Ordem "
rs.Open sql11, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof
	if isnull(rs("01")) then valor01=0 else valor01=nraccess(rs("01"))
	if isnull(rs("02")) then valor02=0 else valor02=nraccess(rs("02"))
	if isnull(rs("03")) then valor03=0 else valor03=nraccess(rs("03"))
	if isnull(rs("04")) then valor04=0 else valor04=nraccess(rs("04"))
	sql12="insert into grupo_custo_table (sessao,codsin,data,ordem,grupo,[01],[02],[03],[04]) " & _
	"values ('" & rs("sessao") & "','" & rs("codsin") & "','" & dtaccess(rs("data")) & "'," & rs("ordem") & ",'" & _
	rs("grupo") & "'," & valor01 & "," & valor02 & "," & valor03 & "," & valor04 & ") "
	conexao.execute sql12
rs.movenext
loop
rs.close

'********************************************

if agrupa=0 then

sqlc="select count(sessao) as linhas from grupo_custo_table where codsin='01' and sessao='" & sessao & "' group by sessao"
rs.Open sqlc, ,adOpenStatic, adLockReadOnly
lin_adm=rs("linhas")
rs.close
sqlc="select count(sessao) as linhas from grupo_custo_table where codsin='03' and sessao='" & sessao & "' group by sessao"
rs.Open sqlc, ,adOpenStatic, adLockReadOnly
lin_prof=rs("linhas")
rs.close

sql1="select sessao, codsin, data, grupo, sum([01]) as ns, sum([02]) as br, sum([03]) as vy, sum([04]) as jw " & _
"FROM grupo_custo_table " & _
"WHERE sessao='" & sessao & "' " & _
"group by sessao, codsin, data, grupo order by codsin, data "
rs.Open sql1, ,adOpenStatic, adLockReadOnly

total=rs.recordcount
rs.movefirst
%>
<!-- <div align="right"> -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650" height="990">
<tr><td valign=top>

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<%if bras=1 then colunas=8 else colunas=7%>
<tr><td class=titulop colspan=<%=colunas%> align="center">CUSTO DA FOLHA DE <%=ucase(monthname(mescomp))%>/<%=anocomp%></td></tr>
<tr>
	<td class="campop" colspan=2 align="center" style="border-bottom: 2 solid #000000"><b>DESCRIÇÃO</td>
	<td class="campop" align="center" style="border-bottom: 2 solid #000000"><b>NARCISO</td>
<%if bras=1 then%><td class="campop" style="border-bottom: 2 solid #000000"><b>BRÁS</td><%end if%>
	<td class="campop" align="center" style="border-bottom: 2 solid #000000"><b>VILA YARA</td>
	<td class="campop" align="center" style="border-bottom: 2 solid #000000"><b>JD. WILSON</td>
	<td class="campop" align="center" style="border-bottom: 2 solid #000000"><b>TOTAL</td>
	<td class="campop" align="center" style="border-bottom: 2 solid #000000"><b>DATA</td>
</tr>
<%
inicio=0
do while not rs.eof
%>
<tr>
<%
totall=0
if rs("codsin")<>lastsin then
	if rs("codsin")="01" then lin=lin_adm:texto="A<br>D<br>M<br>I<br>N<br>I<br>S<br>T<br>R<br>A<br>T<br>I<br>V<br>O<br>S":texto2="Administrativos"
	if rs("codsin")="03" then lin=lin_prof:texto="P<br>R<br>O<br>F<br>E<br>S<br>S<br>O<br>R<br>E<br>S":texto2="Professores"
	if inicio=1 then
		response.write "<tr><td class=""campop"" colspan=2 style='border-bottom:2 solid #000000'>Total Administrativos</td>"
		response.write "<td class=""campop"" style='border-bottom:2 solid #000000' align=""right"">" & formatnumber(totalns,2) & "</td>"
		if bras=1 then response.write "<td class=""campop"" style='border-bottom:2 solid #000000' align=""right"">" & formatnumber(totalbr,2) & "</td>"
		response.write "<td class=""campop"" style='border-bottom:2 solid #000000' align=""right"">" & formatnumber(totalvy,2) & "</td>"
		response.write "<td class=""campop"" style='border-bottom:2 solid #000000' align=""right"">" & formatnumber(totaljw,2) & "</td>"
		response.write "<td class=""campop"" style='border-bottom:2 solid #000000' align=""right""><b>" & formatnumber(total,2) & "</td>"
		response.write "<td class=fundo style='border-bottom:2 solid #000000'>&nbsp;</td></tr>"
		totalns=0:totalbr=0:totalvy=0:totaljw=0:total=0
	end if
	response.write "<td class=titulop align=""center"" style='border:2 solid #000000' rowspan=" & lin & ">" & texto & "</td>"
end if
totall=rs("ns")+rs("br")+rs("vy")+rs("jw")
totalns=totalns+rs("ns")
totalbr=totalbr+rs("br")
totalvy=totalvy+rs("vy")
totaljw=totaljw+rs("jw")
totalnsg=totalnsg+rs("ns")
totalbrg=totalbrg+rs("br")
totalvyg=totalvyg+rs("vy")
totaljwg=totaljwg+rs("jw")
total=total+totall
totalg=totalg+totall
%>
	<td class="campop"><%=rs("grupo")%></td>
	<td class="campop" align="right"><%=formatnumber(rs("ns"),2)%></td>
<%if bras=1 then%><td class="campop" align="right"><%=formatnumber(rs("br"),2)%></td><%end if%>
	<td class="campop" align="right"><%=formatnumber(rs("vy"),2)%></td>
	<td class="campop" align="right"><%=formatnumber(rs("jw"),2)%></td>
	<td class="campop" align="right"><b><%=formatnumber(totall,2)%></td>
	<td class="campop" align="center"><%=rs("data")%></td>
</tr>
<%
lastsin=rs("codsin")
inicio=1
rs.movenext
loop
rs.close
%>
<tr>
	<td class="campop" colspan=2 style='border-bottom:2 solid #000000'>Total Professores</td>
	<td class="campop" style='border-bottom:2 solid #000000' align="right"><%=formatnumber(totalns,2)%></td>
	<%if bras=1 then%><td class="campop" style='border-bottom:2 solid #000000' align="right"><%=formatnumber(totalbr,2)%></td><%end if%>
	<td class="campop" style='border-bottom:2 solid #000000' align="right"><%=formatnumber(totalvy,2)%></td>
	<td class="campop" style='border-bottom:2 solid #000000' align="right"><%=formatnumber(totaljw,2)%></td>
	<td class="campop" style='border-bottom:2 solid #000000' align="right"><b><%=formatnumber(total,2)%></td>
	<td class=fundo style='border-bottom:2 solid #000000'>&nbsp;</td>
</tr>
<tr>
	<td class="campop" colspan=2 style='border-bottom:2 solid #000000'>Total Geral</td>
	<td class="campop" style='border-bottom:2 solid #000000' align="right"><%=formatnumber(totalnsg,2)%></td>
	<%if bras=1 then%><td class="campop" style='border-bottom:2 solid #000000' align="right"><%=formatnumber(totalbrg,2)%></td><%end if%>
	<td class="campop" style='border-bottom:2 solid #000000' align="right"><%=formatnumber(totalvyg,2)%></td>
	<td class="campop" style='border-bottom:2 solid #000000' align="right"><%=formatnumber(totaljwg,2)%></td>
	<td class="campop" style='border-bottom:2 solid #000000' align="right"><b><%=formatnumber(totalg,2)%></td>
	<td class=fundo style='border-bottom:2 solid #000000'>&nbsp;</td>
</tr>
</table>

</td></tr></table>
<%

end if 'agrupa 0

if agrupa=1 then

sql1="select sessao, data, grupo, sum([01]) as ns, sum([02]) as br, sum([03]) as vy, sum([04]) as jw " & _
"FROM grupo_custo_table " & _
"WHERE sessao='" & sessao & "' " & _
"group by sessao, data, grupo order by data "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
rs.movefirst
%>
<!-- <div align="right"> -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650" height="990">
<tr><td valign=top>

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<%if bras=1 then colunas=7 else colunas=6%>
<tr><td class=titulop colspan=<%=colunas%> align="center">CUSTO DA FOLHA DE <%=ucase(monthname(mescomp))%>/<%=anocomp%></td></tr>
<tr>
	<td class="campop" align="center" style="border-bottom: 2 solid #000000"><b>DESCRIÇÃO</td>
	<td class="campop" align="center" style="border-bottom: 2 solid #000000"><b>NARCISO</td>
<%if bras=1 then%><td class="campop" style="border-bottom: 2 solid #000000"><b>BRÁS</td><%end if%>
	<td class="campop" align="center" style="border-bottom: 2 solid #000000"><b>VILA YARA</td>
	<td class="campop" align="center" style="border-bottom: 2 solid #000000"><b>JD. WILSON</td>
	<td class="campop" align="center" style="border-bottom: 2 solid #000000"><b>TOTAL</td>
	<td class="campop" align="center" style="border-bottom: 2 solid #000000"><b>DATA</td>
</tr>
<%
inicio=0
do while not rs.eof
%>
<tr>
<%
totall=0
totall=rs("ns")+rs("br")+rs("vy")+rs("jw")
totalns=totalns+rs("ns")
totalbr=totalbr+rs("br")
totalvy=totalvy+rs("vy")
totaljw=totaljw+rs("jw")
total=total+totall
totalg=totalg+totall
%>
	<td class="campop"><%=rs("grupo")%></td>
	<td class="campop" align="right"><%=formatnumber(rs("ns"),2)%></td>
<%if bras=1 then%><td class="campop" align="right"><%=formatnumber(rs("br"),2)%></td><%end if%>
	<td class="campop" align="right"><%=formatnumber(rs("vy"),2)%></td>
	<td class="campop" align="right"><%=formatnumber(rs("jw"),2)%></td>
	<td class="campop" align="right"><b><%=formatnumber(totall,2)%></td>
	<td class="campop" align="center"><%=rs("data")%></td>
</tr>
<%
inicio=1
rs.movenext
loop
rs.close
%>
<tr>
	<td class="campop" style='border-top:2 solid #000000'>Total Geral</td>
	<td class="campop" style='border-top:2 solid #000000' align="right"><%=formatnumber(totalns,2)%></td>
	<%if bras=1 then%><td class="campop" style='border-top:2 solid #000000' align="right"><%=formatnumber(totalbr,2)%></td><%end if%>
	<td class="campop" style='border-top:2 solid #000000' align="right"><%=formatnumber(totalvy,2)%></td>
	<td class="campop" style='border-top:2 solid #000000' align="right"><%=formatnumber(totaljw,2)%></td>
	<td class="campop" style='border-top:2 solid #000000' align="right"><b><%=formatnumber(totalg,2)%></td>
	<td class=fundo style='border-top:2 solid #000000'>&nbsp;</td>
</tr>
</table>

</td></tr></table>
<%
end if 'agrupa 1
%>

<%
end if  ' request.form
%>
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>