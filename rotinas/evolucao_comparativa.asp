<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a96")="N" or session("a96")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Evolução Comparativa</title>
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
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
executa=1

if request.form("B1")="" then
sql1="select top 1 anocomp as uano, mescomp as umes from corporerm.dbo.pffinanc group by anocomp, mescomp order by anocomp desc, mescomp desc"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
uano=rs("uano"):umes=rs("umes")
'response.write "<br>" & rs("uano")
'response.write "<br>" & rs("umes")
rs.close
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Parâmetros para emissão do quadro de Evolução
<form method="POST" action="evolucao_comparativa.asp" name="form">

<table border="1" cellpadding="2" cellspacing="2" style="border-collapse: collapse">
<tr>
	<td class=titulo>Meses</td>
	<td class=titulo>Eventos</td>
	<td class=titulo>Tipo</td>
	<td class=titulo colspan=2>Outros</td>
</tr>
<tr>
	<td class=campo valign=top>
<select size=16 name="meses" multiple>
<%
mes=umes:ano=uano
for c=1 to 60
datames=dateserial(ano,mes,1)
%>
<option value="<%=datames%>"><%=datames%></option>
<%
mes=mes-1
next
%>
</select>

</td>
<td class=campo valign=top>
	<input type="checkbox" name="eventos" value="P" checked>Proventos<br>
	<input type="checkbox" name="eventos" value="D">Descontos
	<hr>
	<input type="checkbox" name="totais" value="ON">Imprime<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;só totais
</td>
<td class=campo valign=top>
	<input type="radio" name="tipo" value="T" checked>Todos<br>
	<input type="radio" name="tipo" value="03">Prof. (-R-D)<br>
	<input type="radio" name="tipo" value="01">Admin.(-R-D)<br>
	<input type="radio" name="tipo" value="D">Diretoria<br>
	<input type="radio" name="tipo" value="R">Reitoria<br>
	<input type="radio" name="tipo" value="Z">Outros<br>
	<hr>
	<input type="radio" name="tipo" value="A">Administrativos<br>
	<input type="radio" name="tipo" value="P">Professores<br>
<hr>
	Agrupa A/P? <input type="checkbox" name="agrupa" value="ON">
</td>
<td class=campo valign=top>
	<b><u>Imprime?</u></b><br>
	•Ref.   <input type="checkbox" name="imp_ref" value="ON"><br>
	•Quant. <input type="checkbox" name="imp_quant" value="ON"><br>
	•% mês/mês <input type="checkbox" name="imp_perc" value="ON"><br>
	•Total <input type="checkbox" name="imp_tot" value="ON"><br>
	•Média <input type="checkbox" name="imp_med" value="ON"><br> 
	•Excl.Dem.do Mês <input type="checkbox" name="imp_tiradem" value="ON"><br> 
</td>
<td class=campo valign=top>
	<b>C.Custo</b><br>
	<input type="text" name="ccusto" size="8">
	<hr>
	<b>Titulação</b><br>
	<select size="10" name="titulacao" multiple>
		<option value="T" selected>Todas</option>
<%
sql1="select grauinstrucao as codcliente, instrucao as descricao from qry_funcionarios group by grauinstrucao, instrucao order by grauinstrucao desc"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
%>
	<option value="<%=rs("codcliente")%>"><%=rs("descricao")%></option>
<%
rs.movenext
loop
rs.close
%>
	</select>
</td>
</tr>
<tr>
	<td class=fundo colspan=5>
	<input type="submit" value="Pesquisar" name="B1" class="button">
	</td>
</tr>
</table>
</form>
<%
end if

if request.form("B1")<>"" then
executa=1:txtselecao=""
'response.write "<br>meses " & request.form("meses").count
'for a=1 to request.form("meses").count:response.write "<br>"&a&" " & request.form("meses").item(a):next
'response.write "eventos " & request.form("eventos")
'response.write "<br>tipo " & request.form("tipo")
'response.write "<br>agrupa " & request.form("agrupa")
'response.write "<br>imp ref " & request.form("imp_ref")
'response.write "<br>imp quant " & request.form("imp_quant")
'response.write "<br>imp mesames " & request.form("imp_perc")

texto1="("
for a=1 to request.form("meses").count
	'texto1=texto1 & "DateSerial(anocomp,mescomp,1)=#" &	dtaccess(request.form("meses").item(a)) & "#"
	texto1=texto1 & "convert(datetime,convert(char,anocomp)+'/'+convert(char,mescomp)+'/01')='" &	dtaccess(request.form("meses").item(a)) & "'"
	if request.form("meses").count>1 and a<request.form("meses").count then texto1=texto1 & " or "
next
texto1=texto1 & ") "
if request.form("eventos")="P" then texto2=" and e.provdescbase in ('P') "
if request.form("eventos")="D" then texto2=" and e.provdescbase in ('D') "
if request.form("eventos")="P, D" then texto2=" and e.provdescbase in ('P','D') "
'-----
if request.form("tipo")="T" then texto3="":texto3a=""
if request.form("tipo")="03" then texto3=" and f.codsindicato in ('03') ": texto3a=" and f.chapa collate database_default not in (select chapa from zselecao WHERE tipo in ('D','R')) "
if request.form("tipo")="01" then texto3=" and f.codsindicato in ('01') ": texto3a=" and f.chapa collate database_default not in (select chapa from zselecao WHERE tipo in ('D','R')) "
if request.form("tipo")="D"  then texto3a=" and f.chapa collate database_default in (select chapa from zselecao where tipo='D') "
if request.form("tipo")="R"  then texto3a=" and f.chapa collate database_default in (select chapa from zselecao where tipo='R') "
if request.form("tipo")="Z"  then texto3=" and f.codsindicato not in ('03','01') "
if request.form("tipo")="A"  then texto3=" and f.codsindicato in ('01') "
if request.form("tipo")="P"  then texto3=" and f.codsindicato in ('03') "
if request.form("agrupa")="ON" then 
	texto4="'00' as codsindicato, ":texto4a="codsindicato,":texto4b="'00', ":texto4b=""
else 
	texto4="f.codsindicato, ":texto4a="codsindicato, ":texto4b=texto4
end if
'-----
if request.form("imp_ref")="ON" then impref=1 else impref=0
if request.form("imp_quant")="ON" then impquant=1 else impquant=0
if request.form("imp_perc")="ON" then impporc=1 else impporc=0
if request.form("imp_tot")="ON" then imp_tot=1 else imp_tot=0
if request.form("imp_med")="ON" then imp_med=1 else imp_med=0
if request.form("imp_tiradem")="ON" then imp_tiradem=1 else imp_tiradem=0
'-----
if request.form("titulacao").count=1 and request.form("titulacao").item(1)="T" then
	texto6=""
else
	texto6=" and ("
	for a=1 to request.form("titulacao").count
		if request.form("titulacao").item(a)<>"T" then 
			texto6=texto6 & "f.grauinstrucao='" & request.form("titulacao").item(a) & "'"
			if request.form("titulacao").count>1 and a<request.form("titulacao").count then texto6=texto6 & " or "
		end if
	next
	texto6=texto6 & ") "
end if
'-----
if request.form("totais")="ON" then sototal=1 else sototal=0
'-----
if request.form("ccusto")<>"" then texto5=" and f.codsecao like '" & request.form("ccusto") & "%' " else texto5=""
'response.write "texto 5: " &texto5
sessao=session("usuariomaster")
sql1="delete from temp_evolucao where sessao='" & sessao & "'"
if executa=1 then conexao.execute sql1
sql1="delete from temp_evolucaocomp where sessao='" & sessao & "'"
if executa=1 then conexao.execute sql1

imp_dem="CASE When month(demissao)=mescomp and year(demissao)=anocomp then 'D' else 'A' end,"

sql1="INSERT INTO temp_evolucaocomp (sessao,datames,ANOCOMP,MESCOMP, " & texto4a & "CODSITUACAO,CODEVENTO,DESCRICAO,PROVDESCBASE,totqt,tothr,totref,total) " & _
"SELECT '" & sessao & "', convert(datetime,convert(char,anocomp)+'/'+convert(char,mescomp)+'/01') AS datames, ff.ANOCOMP, ff.MESCOMP, " & texto4 & imp_dem & "ff.CODEVENTO, e.DESCRICAO, e.PROVDESCBASE, " & _
"Count(ff.CODEVENTO) AS totqt, Sum(ff.HORA) AS tothr, Sum(ff.REF) AS totref, Sum(ff.VALOR) AS total " & _
"FROM (corporerm.dbo.PFFINANC ff INNER JOIN corporerm.dbo.PEVENTO e ON ff.CODEVENTO=e.CODIGO) INNER JOIN qry_funcionarios f ON ff.CHAPA collate database_default=f.CHAPA " & _
"where ff.valor>0 /*and (ff.chapa<'10000' or ff.chapa>='90000')*/ " & texto5 & texto3a & texto6 & _
"GROUP BY convert(datetime,convert(char,anocomp)+'/'+convert(char,mescomp)+'/01'), ff.ANOCOMP, ff.MESCOMP, " & texto4b & imp_dem & "ff.CODEVENTO, e.DESCRICAO, e.PROVDESCBASE " & _
"HAVING " & texto1 & texto2 & texto3 & ""
'if session("usuariomaster")="02379" then response.write "<br>" & sql1
if executa=1 then conexao.execute sql1

sql1="INSERT INTO temp_evolucaocomp (sessao,datames,ANOCOMP,MESCOMP, " & texto4a & "CODSITUACAO,CODEVENTO,DESCRICAO,PROVDESCBASE,totqt,tothr,totref,total) " & _
"SELECT '" & sessao & "', convert(datetime,convert(char,anocomp)+'/'+convert(char,mescomp)+'/01') AS datames, ff.ANOCOMP, ff.MESCOMP, " & texto4 & imp_dem & "ff.CODEVENTO, e.DESCRICAO, e.PROVDESCBASE, " & _
"Count(ff.CODEVENTO) AS totqt, Sum(ff.HORA) AS tothr, Sum(ff.REF) AS totref, Sum(ff.VALOR) AS total " & _
"FROM (corporerm.dbo.PFFINANC ff INNER JOIN corporerm.dbo.PEVENTO e ON ff.CODEVENTO=e.CODIGO) INNER JOIN qry_funcionarios f ON ff.CHAPA collate database_default=f.CHAPA " & _
"where ff.valor>0 /*and (ff.chapa<'10000' or ff.chapa>='90000')*/ " & texto5 & texto6 & _
"GROUP BY convert(datetime,convert(char,anocomp)+'/'+convert(char,mescomp)+'/01'), ff.ANOCOMP, ff.MESCOMP, " & texto4b & imp_dem & "ff.CODEVENTO, e.DESCRICAO, e.PROVDESCBASE " & _
"HAVING ff.codevento='LIQ' and " & texto1 & " and e.provdescbase in ('B') " & texto3 & ""
'if session("usuariomaster")="02379" then response.write "<br>" & sql1
if executa=1 then conexao.execute sql1

sql1="INSERT INTO temp_evolucaocomp (sessao,datames,ANOCOMP,MESCOMP, " & texto4a & "CODSITUACAO,CODEVENTO,DESCRICAO,PROVDESCBASE,totqt,tothr,totref,total) " & _
"SELECT '" & sessao & "', convert(datetime,convert(char,anocomp)+'/'+convert(char,mescomp)+'/01') AS datames, ff.ANOCOMP, ff.MESCOMP, " & texto4 & imp_dem & "g.ordem, g.grupo, 'B' as expr2, " & _
"Count(ff.CODEVENTO) AS totqt, Sum(ff.HORA) AS tothr, Sum(ff.REF) AS totref, Sum(ff.VALOR * g.fator) AS total " & _
"FROM grupo_custo g INNER JOIN ((corporerm.dbo.PFFINANC ff INNER JOIN corporerm.dbo.PEVENTO e ON ff.CODEVENTO=e.CODIGO) INNER JOIN qry_funcionarios f ON ff.CHAPA collate database_default=f.CHAPA) ON g.CODIGO=ff.CODEVENTO collate database_default " & _
"where ff.valor>0 and g.excecao='0' /*and (ff.chapa<'10000' or ff.chapa>='90000')*/ " & texto5 & texto6 & _
"GROUP BY convert(datetime,convert(char,anocomp)+'/'+convert(char,mescomp)+'/01'), ff.ANOCOMP, ff.MESCOMP, " & texto4b & imp_dem & "g.ordem, g.grupo " & _
"HAVING " & texto1 & texto3 & ""
'if session("usuariomaster")="02379" then response.write "<br>" & sql1
if executa=1 then conexao.execute sql1

if imp_tiradem=1 then
	sql1="delete from temp_evolucaocomp where sessao='" & sessao & "' and codsituacao='D' "
	if executa=1 then conexao.execute sql1
end if

sql2="INSERT INTO temp_evolucao ( sessao, codevento, descricao, provdescbase, codsindicato ) " & _
"SELECT sessao, CODEVENTO, DESCRICAO, PROVDESCBASE, codsindicato FROM temp_evolucaocomp " & _
"GROUP BY sessao, CODEVENTO, DESCRICAO, PROVDESCBASE, codsindicato HAVING sessao='" & sessao & "' "
'if session("usuariomaster")="02379" then response.write "<br>" & sql2
if executa=1 then conexao.execute sql2

numero=0
sql3="SELECT datames FROM temp_evolucaocomp where sessao='" & sessao & "' GROUP BY datames order by datames;"
rs.Open sql3, ,adOpenStatic, adLockReadOnly
do while not rs.eof
	redim preserve meses(numero)
	redim preserve total(numero)
	redim preserve liquido(numero)
	meses(numero)=rs("datames")
	total(numero)=0:liquido(numero)=0
	numero=numero+1
rs.movenext
loop
rs.close

sql4="SELECT sessao, CODEVENTO, DESCRICAO, PROVDESCBASE, CODSINDICATO "
for a=0 to ubound(meses)
	mest=dtaccess(meses(a))
	sql4=sql4 & ", '" & mest & "'=sum(case when datames='" & mest & "' then total else 0 end) "
next
sql4=sql4 & "FROM temp_evolucaocomp WHERE sessao='" & sessao & "' GROUP BY sessao, CODEVENTO, DESCRICAO, PROVDESCBASE, CODSINDICATO "
'response.write "<br>" & sql4
rs.Open sql4, ,adOpenStatic, adLockReadOnly
totalcampos=rs.fields.count-1
rs.movefirst:do while not rs.eof
	sql5="UPDATE temp_evolucao SET "
	if totalcampos>=5 then sql5=sql5 & "v1=" & nraccess(rs.fields(5))
	if totalcampos>=6 then sql5=sql5 & ", v2=" & nraccess(rs.fields(6))
	if totalcampos>=7 then sql5=sql5 & ", v3=" & nraccess(rs.fields(7))
	if totalcampos>=8 then sql5=sql5 & ", v4=" & nraccess(rs.fields(8))
	if totalcampos>=9 then sql5=sql5 & ", v5=" & nraccess(rs.fields(9))
	if totalcampos>=10 then sql5=sql5 & ", v6=" & nraccess(rs.fields(10))
	if totalcampos>=11 then sql5=sql5 & ", v7=" & nraccess(rs.fields(11))
	if totalcampos>=12 then sql5=sql5 & ", v8=" & nraccess(rs.fields(12))
	if totalcampos>=13 then sql5=sql5 & ", v9=" & nraccess(rs.fields(13))
	if totalcampos>=14 then sql5=sql5 & ", v10=" & nraccess(rs.fields(14))
	if totalcampos>=15 then sql5=sql5 & ", v11=" & nraccess(rs.fields(15))
	if totalcampos>=16 then sql5=sql5 & ", v12=" & nraccess(rs.fields(16))
	if totalcampos>=17 then sql5=sql5 & ", v13=" & nraccess(rs.fields(17))
	sql5=sql5 & " WHERE sessao='" & sessao & "' AND codevento='" & rs("codevento") & "' "
	if request.form("agrupa")="" then sql5=sql5 & " AND codsindicato='" & rs("codsindicato") & "' "
	'response.write "<br>" & sql5
	if executa=1 then conexao.execute sql5
rs.movenext:loop
rs.close

sql4="SELECT sessao, CODEVENTO, DESCRICAO, PROVDESCBASE, CODSINDICATO "
for a=0 to ubound(meses)
	mest=dtaccess(meses(a))
	sql4=sql4 & ", '" & mest & "'=sum(case when datames='" & mest & "' then totref else 0 end) "
next
sql4=sql4 & "FROM temp_evolucaocomp WHERE sessao='" & sessao & "' GROUP BY sessao, CODEVENTO, DESCRICAO, PROVDESCBASE, CODSINDICATO "
'response.write "<br>" & sql4
rs.Open sql4, ,adOpenStatic, adLockReadOnly
totalcampos=rs.fields.count-1
rs.movefirst:do while not rs.eof
	sql5="UPDATE temp_evolucao SET "
	if totalcampos>=5 then sql5=sql5 & "r1=" & nraccess(rs.fields(5))
	if totalcampos>=6 then sql5=sql5 & ", r2=" & nraccess(rs.fields(6))
	if totalcampos>=7 then sql5=sql5 & ", r3=" & nraccess(rs.fields(7))
	if totalcampos>=8 then sql5=sql5 & ", r4=" & nraccess(rs.fields(8))
	if totalcampos>=9 then sql5=sql5 & ", r5=" & nraccess(rs.fields(9))
	if totalcampos>=10 then sql5=sql5 & ", r6=" & nraccess(rs.fields(10))
	if totalcampos>=11 then sql5=sql5 & ", r7=" & nraccess(rs.fields(11))
	if totalcampos>=12 then sql5=sql5 & ", r8=" & nraccess(rs.fields(12))
	if totalcampos>=13 then sql5=sql5 & ", r9=" & nraccess(rs.fields(13))
	if totalcampos>=14 then sql5=sql5 & ", r10=" & nraccess(rs.fields(14))
	if totalcampos>=15 then sql5=sql5 & ", r11=" & nraccess(rs.fields(15))
	if totalcampos>=16 then sql5=sql5 & ", r12=" & nraccess(rs.fields(16))
	if totalcampos>=17 then sql5=sql5 & ", r13=" & nraccess(rs.fields(17))
	sql5=sql5 & " WHERE sessao='" & sessao & "' AND codevento='" & rs("codevento") & "' "
	if request.form("agrupa")="" then sql5=sql5 & " AND codsindicato='" & rs("codsindicato") & "' "
	'response.write "<br>" & sql5
	if executa=1 then conexao.execute sql5
rs.movenext:loop
rs.close

sql4="TRANSFORM Sum(totqt) AS SomaQt " & _
"SELECT sessao, CODEVENTO, DESCRICAO, PROVDESCBASE, CODSINDICATO FROM temp_evolucaocomp " & _
"WHERE sessao='" & sessao & "' GROUP BY sessao, CODEVENTO, DESCRICAO, PROVDESCBASE, CODSINDICATO " & _
"PIVOT datames "
sql4="SELECT sessao, CODEVENTO, DESCRICAO, PROVDESCBASE, CODSINDICATO "
for a=0 to ubound(meses)
	mest=dtaccess(meses(a))
	sql4=sql4 & ", '" & mest & "'=sum(case when datames='" & mest & "' then totqt else 0 end) "
next
sql4=sql4 & "FROM temp_evolucaocomp WHERE sessao='" & sessao & "' GROUP BY sessao, CODEVENTO, DESCRICAO, PROVDESCBASE, CODSINDICATO "
'response.write "<br>" & sql4
rs.Open sql4, ,adOpenStatic, adLockReadOnly
totalcampos=rs.fields.count-1
rs.movefirst:do while not rs.eof
	sql5="UPDATE temp_evolucao SET "
	if totalcampos>=5 then sql5=sql5 & "q1=" & nraccess(rs.fields(5))
	if totalcampos>=6 then sql5=sql5 & ", q2=" & nraccess(rs.fields(6))
	if totalcampos>=7 then sql5=sql5 & ", q3=" & nraccess(rs.fields(7))
	if totalcampos>=8 then sql5=sql5 & ", q4=" & nraccess(rs.fields(8))
	if totalcampos>=9 then sql5=sql5 & ", q5=" & nraccess(rs.fields(9))
	if totalcampos>=10 then sql5=sql5 & ", q6=" & nraccess(rs.fields(10))
	if totalcampos>=11 then sql5=sql5 & ", q7=" & nraccess(rs.fields(11))
	if totalcampos>=12 then sql5=sql5 & ", q8=" & nraccess(rs.fields(12))
	if totalcampos>=13 then sql5=sql5 & ", q9=" & nraccess(rs.fields(13))
	if totalcampos>=14 then sql5=sql5 & ", q10=" & nraccess(rs.fields(14))
	if totalcampos>=15 then sql5=sql5 & ", q11=" & nraccess(rs.fields(15))
	if totalcampos>=16 then sql5=sql5 & ", q12=" & nraccess(rs.fields(16))
	if totalcampos>=17 then sql5=sql5 & ", q13=" & nraccess(rs.fields(17))
	sql5=sql5 & " WHERE sessao='" & sessao & "' AND codevento='" & rs("codevento") & "' "
	if request.form("agrupa")="" then sql5=sql5 & " AND codsindicato='" & rs("codsindicato") & "' "
	'response.write "<br>" & sql5
	if executa=1 then conexao.execute sql5
rs.movenext:loop
rs.close

'response.write "<br>"
for b=0 to ubound(meses)
	'response.write b & "->" &  meses(b) & " "
next

sql6="select * from temp_evolucao where sessao='" & sessao & "' order by codsindicato, provdescbase desc, codevento "
rs.Open sql6, ,adOpenStatic, adLockReadOnly

'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a=0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'*************** fim teste **********************

sql7="SELECT Count(codsindicato) AS linhas FROM temp_evolucao WHERE sessao='" & sessao & "' AND codsindicato='01'"
rs2.Open sql7, ,adOpenStatic, adLockReadOnly
linhas01=rs2("linhas")
rs2.close
sql7="SELECT Count(codsindicato) AS linhas FROM temp_evolucao WHERE sessao='" & sessao & "' AND codsindicato='03'"
rs2.Open sql7, ,adOpenStatic, adLockReadOnly
linhas03=rs2("linhas")
rs2.close
sql7="SELECT Count(codsindicato) AS linhas FROM temp_evolucao WHERE sessao='" & sessao & "' AND codsindicato='00'"
rs2.Open sql7, ,adOpenStatic, adLockReadOnly
linhas00=rs2("linhas")
rs2.close

inicio=1:tgrupo=0
rs.movefirst
%>
<table border="1" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo rowspan=2>&nbsp;</td>
	<td class=titulo rowspan=2 align="center">Cod.</td>
	<td class=titulo rowspan=2 align="center">Evento</td>
<%for a=0 to ubound(meses)
	cspan=1+impquant+impref+impporc
	if a=ubound(meses) and impporc=0 then cspan=cspan+1
	response.write "<td class=titulo align=""center"" colspan=" & cspan & " style='border-right:2 solid #000000'>" & monthname(month(meses(a)))&"/"&year(meses(a)) & "</td>"
next 'a%>
<%if imp_tot=1 then%>
	<td class=titulo rowspan=2 align="center" style="border-right:2 solid #000000">Total</td><%end if%>
<%if imp_med=1 then%>
	<td class=titulo colspan=2 align="center" style="border-right:2 solid #000000">Média Ant.</td>
	<td class=titulo rowspan=2 align="center" style="border-right:2 solid #000000">Média Atual</td><%end if%>
</tr>
<tr>
<%for a=0 to ubound(meses)
	varborda=" style='border-right:2 solid #000000'"
	if impquant=1 then response.write "<td class=titulo align=""center"">Qt.</td>"
	if impref=1 then response.write "<td class=titulo align=""center"">Ref.</td>"
	if impporc=0 and a<ubound(meses) then txt1=varborda else txt1=""
	response.write "<td class=titulo align=""center""" & txt1 & ">Valor</td>"
	if impporc=1 or a=ubound(meses) then response.write "<td class=titulo align=""center"" style='border-right:2 solid #000000'>%</td>"
next 'a%>
<%if imp_med=1 then%>
	<td class=titulo align="center">$</td>
	<td class=titulo align="center" style='border-right:2 solid #000000'>% s/<%=monthname(month(meses(ubound(meses))),1)%></td><%end if%>
</tr>
<%
do while not rs.eof
if rs("codsindicato")="01" then textol="ADMINISTRATIVOS"
if rs("codsindicato")="03" then textol="PROFESSORES"
if rs("codsindicato")="00" then textol="TODOS"
'---------------------------------------------------------------------------------------
'***************************************************************************************
if lasttipo<>rs("provdescbase") or lastsind<>rs("codsindicato") then
	sqlc="SELECT Count(codevento) AS linhas FROM temp_evolucao " & _
	"WHERE sessao='" & sessao & "' AND provdescbase='" & rs("provdescbase") & "' AND codsindicato='" & rs("codsindicato") & "' "
	rs2.Open sqlc, ,adOpenStatic, adLockReadOnly
	linhas=rs2("linhas")
	rs2.close:primeira=1
	if sototal=1 then textocab=textol else textocab=""
	if rs("provdescbase")="P" then tipoeve="P - Proventos " & textocab
	if rs("provdescbase")="D" then tipoeve="D - Descontos " & textocab
	if rs("provdescbase")="B" then tipoeve="B - Eventos de Base " & textocab
	if inicio=0 then
%>
<tr>
	<td colspan=3 class=fundo style="border-top:2 solid;border-bottom:4 double">Total <%=lasttipo%></td>
<%for a=0 to ubound(meses)
	coluna=a+1%>
	<%if impquant=1 then%><td class=fundo align="right" style="border-top:2 solid;border-bottom:4 double">&nbsp;</td><%end if%>
	<%if impref=1 then%><td class=fundo align="right" style="border-top:2 solid;border-bottom:4 double">&nbsp;</td><%end if%>
<%if impporc=0 and a<ubound(meses) then txt1=varborda else txt1=""%>
	<td class=campo align="right" <%=txt1%> style="border-top:2 solid;border-bottom:4 double" ><%if isnumeric(total(a)) then response.write formatnumber(total(a),2)%></td>
<%if impporc=1 or a=ubound(meses) then%><td class="campolr" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double">
<%
if a>0 then
	valor2=total(a):if isnull(valor2) then valor2=0
	valor1=total(a-1):if isnull(valor1) then valor1=0
	if valor1=0 or valor2=0 then variacao="-" else variacao=(valor2 / valor1) - 1
	if isnumeric(variacao) then variacao=formatpercent(variacao,2)
	response.write variacao
end if
%>	
	</td><%end if%>
<%next 'a%>
<%if imp_tot=1 then
totalg=0:for a=0 to ubound(total):totalg=totalg+total(a):next
%>
	<td class="campoa" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%if isnull(totalg) then response.write "-" else response.write formatnumber(totalg,2)%></td>
<%end if
if imp_med=1 then
totalm1=0:for a=0 to ubound(total)-1:totalm1=totalm1+total(a):next:totalm1=totalm1/ubound(total)
totalm2=0:for a=0 to ubound(total):totalm2=totalm2+total(a):next:totalm2=totalm2/(ubound(total)+1)
%>
	<td class="campot" align="right" style="border-top:2 solid;border-bottom:4 double"><%if isnull(totalm1) then response.write "-" else response.write formatnumber(totalm1,2)%></td>
	<td class="campolr" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%
	v_ult_mes=total(ubound(total))
	if isnull(v_ult_mes) then v_ult_mes=0
	if v_ult_mes=0 or totalm1=0 then variacaom="-" else variacaom=(v_ult_mes/totalm1)-1
	if isnumeric(variacaom) then 
		if variacaom>1 then casas=0 else casas=2
		variacaom=formatpercent(variacaom,casas)
	end if
	response.write variacaom
	%></td>
	<td class="campot" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%if isnull(totalm2) then response.write "-" else response.write formatnumber(totalm2,2)%></td>
<%end if%>
</tr>
<%
	end if
	if lasttipo="D" then
	'------------- liquido ini ---------------------------
%>
<tr>
	<td colspan=3 class="campot" style="border-top:2 solid;border-bottom:4 double">Líquido</td>
<%for a=0 to ubound(meses)
	coluna=a+1%>
	<%if impquant=1 then%><td class="campot" align="right" style="border-top:2 solid;border-bottom:4 double">&nbsp;</td><%end if%>
	<%if impref=1 then%><td class="campot" align="right" style="border-top:2 solid;border-bottom:4 double">&nbsp;</td><%end if%>
<%if impporc=0 and a<ubound(meses) then txt1=varborda else txt1=""%>
	<td class="campot" align="right" <%=txt1%> style="border-top:2 solid;border-bottom:4 double" ><%if isnumeric(liquido(a)) then response.write formatnumber(liquido(a),2)%></td>
<%if impporc=1 or a=ubound(meses) then%><td class="campot"r align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double">
<%
if a>0 then
	valor2=liquido(a):if isnull(valor2) then valor2=0
	valor1=liquido(a-1):if isnull(valor1) then valor1=0
	if valor1=0 or valor2=0 then variacao="-" else variacao=(valor2 / valor1) - 1
	if isnumeric(variacao) then variacao=formatpercent(variacao,2)
	response.write variacao
end if
%>	
	</td><%end if%>
<%next 'a%>
<%if imp_tot=1 then
totalg=0:for a=0 to ubound(liquido):totalg=totalg+liquido(a):next
%>
	<td class="campoa" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%if isnull(totalg) then response.write "-" else response.write formatnumber(totalg,2)%></td>
<%end if
if imp_med=1 then
totalm1=0:for a=0 to ubound(liquido)-1:totalm1=totalm1+liquido(a):next:totalm1=totalm1/ubound(liquido)
totalm2=0:for a=0 to ubound(liquido):totalm2=totalm2+liquido(a):next:totalm2=totalm2/(ubound(liquido)+1)
%>
	<td class="campot" align="right" style="border-top:2 solid;border-bottom:4 double"><%if isnull(totalm1) then response.write "-" else response.write formatnumber(totalm1,2)%></td>
	<td class="campolr" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%
	v_ult_mes=liquido(ubound(liquido))
	if isnull(v_ult_mes) then v_ult_mes=0
	if v_ult_mes=0 or totalm1=0 then variacaom="-" else variacaom=(v_ult_mes/totalm1)-1
	if isnumeric(variacaom) then 
		if variacaom>1 then casas=0 else casas=2
		variacaom=formatpercent(variacaom,casas)
	end if
	response.write variacaom
	%></td>
	<td class="campot" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%if isnull(totalm2) then response.write "-" else response.write formatnumber(totalm2,2)%></td>
<%end if%>
</tr>
<%
	'------------- liquido fim ---------------------------
	for a=0 to ubound(liquido):liquido(a)=0:next
	end if 'lasttipo=D

	response.write "<tr><td class=""campoa"" colspan=28>" & tipoeve & "</td></tr>"
	for a=0 to ubound(total):total(a)=0:next
end if

'***************************************************************************************
'---------------------------------------------------------------------------------------

'->->->->->->->->->->->->->->-> Detalhes <-<-<-<-<-<-<-<-<-<-<-<-<-<-
tlinha=0:tmedia=0:tmediaa=0:tmedia1=0
for a=0 to ubound(meses)
	coluna=a+1
	vtemp=rs("v" & coluna)
	if isnull(vtemp) then vtemp=0
	tlinha=tlinha + vtemp
	if a<ubound(meses) then tmediaa=tmediaa+vtemp else tmediaa=tmediaa
	if rs("provdescbase")="D" then fator=-1 else fator=1
	total(a)=total(a)+vtemp
	liquido(a)=liquido(a)+(vtemp*fator)
next 'a
tmedia =tlinha / (ubound(meses)+1)
tmedia1=tmediaa / ((ubound(meses)+1)-1)
%>
<%
'
'
'
%>
<%
if sototal=0 then 'sototal
%>
<tr>
<%if primeira=1 then%>
	<td class=campo rowspan=<%=linhas%> align="center" valign=top>
<%
'if rs("codsindicato")="01" then textol="ADMINISTRATIVOS"
'if rs("codsindicato")="03" then textol="PROFESSORES"
'if rs("codsindicato")="00" then textol="TODOS"
response.write "<b>"
for b=1 to len(textol)
	response.write mid(textol,b,1) & "<br>"
next
%>	
	</td>
<%end if:primeira=0%>
	<td class=campo><%=rs("codevento")%></td>
	<td class=campo><%=rs("descricao")%></td>
<%for a=0 to ubound(meses)
	coluna=a+1%>
	<%if impquant=1 then%><td class=campo align="right"><%=rs("q" & coluna)%></td><%end if%>
	<%if impref=1 then%><td class=campo align="right"><%if rs("r" & coluna)<>0 then response.write rs("r" & coluna)%></td><%end if%>
<%if impporc=0 and a<ubound(meses) then txt1=varborda else txt1=""%>
	<td class=campo align="right" <%=txt1%> ><%if isnumeric(rs("v" & coluna)) then response.write formatnumber(rs("v" & coluna),2)%></td>
<%if impporc=1 or a=ubound(meses) then%><td class="campolr" align="right" style="border-right:2 solid #000000" nowrap>
<%
if a>0 then
	valor2=rs("v" & coluna):if isnull(valor2) then valor2=0
	valor1=rs("v" & coluna-1):if isnull(valor1) then valor1=0
	if valor1=0 or valor2=0 then variacao="-" else variacao=(valor2 / valor1) - 1
	if isnumeric(variacao) then
		if variacao>1 then casas=0 else casas=2
		variacao=formatpercent(variacao,casas)
	end if
	response.write variacao
end if
%>	
	</td><%end if%>
<%next 'a%>
<%if imp_tot=1 then%>
	<td class="campoa" align="right" style="border-right:2 solid #000000"><%if isnull(tlinha) then response.write "-" else response.write formatnumber(tlinha,2)%></td>
<%end if
if imp_med=1 then%>
	<td class="campot" align="right"><%if isnull(tmedia1) then response.write "-" else response.write formatnumber(tmedia1,2)%></td>
	<td class="campolr" align="right" style="border-right:2 solid #000000"><%
	v_ult_mes=rs("v" & ubound(meses)+1)
	if isnull(v_ult_mes) then v_ult_mes=0
	if v_ult_mes=0 or tmedia1=0 then variacaom="-" else variacaom=(v_ult_mes/tmedia1)-1
	if isnumeric(variacaom) then 
		if variacaom>1 then casas=0 else casas=2
		variacaom=formatpercent(variacaom,casas)
	end if
	response.write variacaom
	%></td>
	<td class="campot" align="right" style="border-right:2 solid #000000"><%if isnull(tmedia) then response.write "-" else response.write formatnumber(tmedia,2)%></td>
<%end if%>
<%%>
</tr>
<%
end if 'sototal
%>

<%
lasttipo=rs("provdescbase")
lastsind=rs("codsindicato")
inicio=0
rs.movenext
loop

'---------------------------------------------------------------------------------------
'***************************************************************************************
if sototal=1 then textocab=textol else textocab=""
if lasttipo="P" then tipoeve="P - Proventos " & textocab
if lasttipo="D" then tipoeve="D - Descontos " & textocab
if lasttipo="B" then tipoeve="B - Eventos de Base " & textocab
if inicio=0 then
%>
<tr>
	<td colspan=3 class=fundo style="border-top:2 solid;border-bottom:4 double">Total <%=lasttipo%></td>
<%for a=0 to ubound(meses)
	coluna=a+1%>
	<%if impquant=1 then%><td class=fundo align="right" style="border-top:2 solid;border-bottom:4 double">&nbsp;</td><%end if%>
	<%if impref=1 then%><td class=fundo align="right" style="border-top:2 solid;border-bottom:4 double">&nbsp;</td><%end if%>
<%if impporc=0 and a<ubound(meses) then txt1=varborda else txt1=""%>
	<td class=campo align="right" <%=txt1%> style="border-top:2 solid;border-bottom:4 double" ><%if isnumeric(total(a)) then response.write formatnumber(total(a),2)%></td>
<%if impporc=1 or a=ubound(meses) then%><td class="campolr" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double">
<%
if a>0 then
	valor2=total(a):if isnull(valor2) then valor2=0
	valor1=total(a-1):if isnull(valor1) then valor1=0
	if valor1=0 or valor2=0 then variacao="-" else variacao=(valor2 / valor1) - 1
	if isnumeric(variacao) then variacao=formatpercent(variacao,2)
	response.write variacao
end if
%>	
	</td><%end if%>
<%next 'a%>
<%if imp_tot=1 then
totalg=0:for a=0 to ubound(total):totalg=totalg+total(a):next
%>
	<td class="campoa" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%if isnull(totalg) then response.write "-" else response.write formatnumber(totalg,2)%></td>
<%end if
if imp_med=1 then
totalm1=0:for a=0 to ubound(total)-1:totalm1=totalm1+total(a):next:totalm1=totalm1/ubound(total)
totalm2=0:for a=0 to ubound(total):totalm2=totalm2+total(a):next:totalm2=totalm2/(ubound(total)+1)
%>
	<td class="campot" align="right" style="border-top:2 solid;border-bottom:4 double"><%if isnull(totalm1) then response.write "-" else response.write formatnumber(totalm1,2)%></td>
	<td class="campolr" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%
	v_ult_mes=total(ubound(total))
	if isnull(v_ult_mes) then v_ult_mes=0
	if v_ult_mes=0 or totalm1=0 then variacaom="-" else variacaom=(v_ult_mes/totalm1)-1
	if isnumeric(variacaom) then 
		if variacaom>1 then casas=0 else casas=2
		variacaom=formatpercent(variacaom,casas)
	end if
	response.write variacaom
	%></td>
	<td class="campot" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%if isnull(totalm2) then response.write "-" else response.write formatnumber(totalm2,2)%></td>
<%end if%>
</tr>
<%
	end if

	if lasttipo="D" then
	'------------- liquido ini ---------------------------
%>
<tr>
	<td colspan=3 class="campot" style="border-top:2 solid;border-bottom:4 double">Líquido</td>
<%for a=0 to ubound(meses)
	coluna=a+1%>
	<%if impquant=1 then%><td class="campot" align="right" style="border-top:2 solid;border-bottom:4 double">&nbsp;</td><%end if%>
	<%if impref=1 then%><td class="campot" align="right" style="border-top:2 solid;border-bottom:4 double">&nbsp;</td><%end if%>
<%if impporc=0 and a<ubound(meses) then txt1=varborda else txt1=""%>
	<td class="campot" align="right" <%=txt1%> style="border-top:2 solid;border-bottom:4 double" ><%if isnumeric(liquido(a)) then response.write formatnumber(liquido(a),2)%></td>
<%if impporc=1 or a=ubound(meses) then%><td class="campot"r align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double">
<%
if a>0 then
	valor2=liquido(a):if isnull(valor2) then valor2=0
	valor1=liquido(a-1):if isnull(valor1) then valor1=0
	if valor1=0 or valor2=0 then variacao="-" else variacao=(valor2 / valor1) - 1
	if isnumeric(variacao) then variacao=formatpercent(variacao,2)
	response.write variacao
end if
%>	
	</td><%end if%>
<%next 'a%>
<%if imp_tot=1 then
totalg=0:for a=0 to ubound(liquido):totalg=totalg+liquido(a):next
%>
	<td class="campoa" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%if isnull(totalg) then response.write "-" else response.write formatnumber(totalg,2)%></td>
<%end if
if imp_med=1 then
totalm1=0:for a=0 to ubound(liquido)-1:totalm1=totalm1+liquido(a):next:totalm1=totalm1/ubound(liquido)
totalm2=0:for a=0 to ubound(liquido):totalm2=totalm2+liquido(a):next:totalm2=totalm2/(ubound(liquido)+1)
%>
	<td class="campot" align="right" style="border-top:2 solid;border-bottom:4 double"><%if isnull(totalm1) then response.write "-" else response.write formatnumber(totalm1,2)%></td>
	<td class="campolr" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%
	v_ult_mes=liquido(ubound(liquido))
	if isnull(v_ult_mes) then v_ult_mes=0
	if v_ult_mes=0 or totalm1=0 then variacaom="-" else variacaom=(v_ult_mes/totalm1)-1
	if isnumeric(variacaom) then 
		if variacaom>1 then casas=0 else casas=2
		variacaom=formatpercent(variacaom,casas)
	end if
	response.write variacaom
	%></td>
	<td class="campot" align="right" style="border-right:2 solid #000000;border-top:2 solid;border-bottom:4 double"><%if isnull(totalm2) then response.write "-" else response.write formatnumber(totalm2,2)%></td>
<%end if%>
</tr>
<%
	'------------- liquido fim ---------------------------
	for a=0 to ubound(liquido):liquido(a)=0:next
	end if 'lasttipo=D
	
	for a=0 to ubound(total):total(a)=0:next
'***************************************************************************************
'---------------------------------------------------------------------------------------

rs.close
%>
</table>
<%
end if
%>
</body>
</html>
<%
'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>