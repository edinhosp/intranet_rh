<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a58")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Geração de Arquivo de Assistência Médica</title>
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

if request.form<>"" then
	prorata=request.form("prorata")
	ano=request.form("ano")
	mes=request.form("mes")
	mesq=numzero(mes,2)
	dataq=mesq & "/" & ano
	if request.form("estilo")="A" then
		datab=dateserial(ano,mes+1,1)
	else
		datab=dateserial(ano,mes+0,1)
	end if
	empresa=request.form("empresa")
	sessao=session("usuariomaster")
	
	conexao.execute "delete from assmed_folha where mesbase='" & dataq & "' and sessao='" & sessao & "' "
	'RESPONSE.WRITE   "delete from assmed_folha where mesbase='" & dataq & "' and sessao='" & sessao & "' "
	'conexao.execute "delete from assmed_folha where sessao='" & sessao & "' "
	
sql="INSERT INTO assmed_folha (sessao, mesbase, tp, CHAPA, principal, beneficiario, empresa, valor_plano, reemb) " & _
"SELECT '" & sessao & "' as sessao,'" & dataq & "' AS mesbase, 'Titular' AS tp, ab.chapa, f.nome AS principal, f.nome AS beneficiario, am.empresa, " & _
"valor_plano=case when " & prorata & "<>30 then round((valor/30)* " & prorata & ",2) else valor end, " & _
"reemb=case when " & prorata & "<>30 then round((reembolso/30)* " & prorata & ",2) else reembolso end " & _
"FROM assmed_beneficiario ab inner join assmed_mudanca am on ab.chapa=am.chapa " & _
"inner join assmed_planos ap on am.empresa=ap.codigo and am.plano=ap.plano " & _
"inner join corporerm.dbo.pfunc f on f.chapa collate database_default=ab.chapa " & _
"WHERE ('" & dtaccess(datab) & "' Between [ivigencia] And [fvigencia]) "
if empresa="T" then sql=sql & " and am.empresa in ('O','I','M','IP','MP','UC','U','V','BS','BP','C','CP') "
if empresa="O" then sql=sql & " and am.empresa in ('O','V') "
if empresa="I" then sql=sql & " and am.empresa in ('I','IP') "
if empresa="M" then sql=sql & " and (am.empresa in ('M','MP','UC','U','BS','BP','C','CP') or (am.empresa in ('I','IP') and codsindicato='01')) "
	conexao.execute sql

	sql="INSERT INTO assmed_folha (sessao, mesbase, tp, CHAPA, principal, beneficiario, empresa, valor_plano, reemb) " & _
"SELECT '" & sessao & "' as sessao, '" & dataq & "' AS Mesbase, 'Dependente' AS tp, ab.chapa, f.nome AS principal, ad.dependente AS beneficiario, adm.empresa, " & _
"valor_plano=case when " & prorata & "<>30 then round((valor/30)* " & prorata & ",2) else valor end, " & _
"reemb=case when " & prorata & "<>30 then round(((case when adm.empresa='I' and codsindicato='03' then [valor] else " & _
"case when ((adm.empresa='U' or adm.empresa='BS' or adm.empresa='C') and ((parentesco='Conjuge' or parentesco='Companheiro(a)') and sexo='M') and adm.plano not like 'SÊNIOR') or (adm.empresa='I' and (parentesco='Conjuge' or parentesco='Companheiro(a)') and sexo='M')  then [valor] " & _
"else reembolso end end)/30)*" & prorata & ",2) else " & _
"(case when adm.empresa='I' and codsindicato='03' then [valor] else " & _
"case when ((adm.empresa='U' or adm.empresa='BS' or adm.empresa='C') and ((parentesco='Conjuge' or parentesco='Companheiro(a)') and sexo='M') and adm.plano not like 'SÊNIOR') or (adm.empresa='I' and (parentesco='Conjuge' or parentesco='Companheiro(a)') and sexo='M') then [valor] " & _
"else [reembolso] end end) end " & _
"FROM assmed_beneficiario ab inner join assmed_dep ad on ab.chapa=ad.chapa " & _
"inner join assmed_dep_mudanca adm on ad.chapa=adm.chapa and ad.nrodepend=adm.nrodepend " & _
"inner join assmed_planos ap on adm.empresa=ap.codigo and adm.plano=ap.plano " & _
"inner join corporerm.dbo.pfunc f on f.chapa collate database_default=ab.chapa " & _
"WHERE ('" & dtaccess(datab) & "' Between [ivigencia] And [fvigencia]) "
if empresa="T" then sql=sql & " and adm.empresa in ('O','I','M','IP','MP','UC','U','V','BS','BP','C','CP') "
if empresa="O" then sql=sql & " and adm.empresa in ('O','V') "
if empresa="I" then sql=sql & " and adm.empresa in ('I','IP') "
if empresa="M" then sql=sql & " and (adm.empresa in ('M','MP','UC','U','BS','BP','C','CP') or (adm.empresa in ('I','IP') and codsindicato='01')) "
	conexao.execute sql

	sql="INSERT INTO assmed_folha (sessao, mesbase, tp, CHAPA, principal, beneficiario, empresa, valor_plano, reemb) " & _
"SELECT '" & sessao & "' as sessao, '" & dataq & "' AS Mesbase, " & _
"'Acerto' AS tp, " & _
"ab.chapa, f.nome AS principal, ac.descricao AS beneficiario, " & _
"ac.empresa, sum(ac.valor_acerto) AS valor_plano, sum(ac.reembolso) AS reemb " & _
"FROM assmed_beneficiario ab inner join assmed_acertos ac on ab.chapa=ac.chapa " & _
"inner join corporerm.dbo.pfunc f on f.chapa collate database_default=ab.chapa " & _
"WHERE ac.reembolso<>0 " & _
"and convert(smalldatetime,convert(char,month(data_acerto))+'/01/'+convert(char,year(data_acerto)))='" & dtaccess(datab) & "' "
if empresa="T" then sql=sql & " and ac.empresa in ('O','I','M','IP','MP','UC','U','V','BS','BP','C','CP') "
if empresa="O" then sql=sql & " and ac.empresa in ('O','V') "
if empresa="I" then sql=sql & " and ac.empresa in ('I','IP') "
if empresa="M" then sql=sql & " and (ac.empresa in ('M','MP','UC','U','BS','BP','C','CP') or (ac.empresa in ('I','IP') and codsindicato='01')) "
sql=sql & " group by ab.chapa, f.nome, ac.descricao, ac.empresa "
	conexao.execute sql
	
sql="update assmed_folha set reemb=0 where mesbase='" & dataq & "' and sessao='" & sessao & "' and chapa='00057' and tp<>'Acerto'"
conexao.execute sql
sql="update assmed_folha set reemb=0 where mesbase='" & dataq & "' and sessao='" & sessao & "' and chapa='00099' and tp<>'Acerto'"
conexao.execute sql
sql="update assmed_folha set reemb=0 where mesbase='" & dataq & "' and sessao='" & sessao & "' and chapa='02538' and tp<>'Acerto'"
conexao.execute sql
'sql="delete from assmed_folha where mesbase='" & dataq & "' and sessao='" & sessao & "' and chapa='02239'"
'conexao.execute sql
	
	sql="select * from assmed_folha where mesbase='" & dataq & "' and sessao='" & sessao & "' and reemb<>0 order by chapa, tp desc"
end if
%>

<p class=titulo>Geração de arquivo de Assistência Médica para o RM Labore
<%
if request.form="" then
mesfolha=month(dateserial(year(now),month(now),1))
anofolha=year(dateserial(year(now),month(now),1))
%>
<form method="POST" action="assmed_labore.asp">
<p style="margin-bottom:0;margin-top:0;background-color:#FFFFCC">
	Folha do Mês <input type="text" name="mes" size="2" value="<%=mesfolha%>"> / <input type="text" name="ano" size="4" value="<%=anofolha%>">
<p style="margin-bottom:0;margin-top:0;background-color:#CCFFCC">
	Cálculo Proporcional (pró-rata) <input type="text" name="prorata" size=2 value=30>
<p style="margin-bottom:0;margin-top:0;background-color:#FFCCCC">
	<input type="radio" name="empresa" value="M">Administrativos
   	<input type="radio" name="empresa" value="I">Intermédica 
   	<input type="radio" name="empresa" value="O">Metlife Odonto
   	<input type="radio" name="empresa" value="T" checked>Todas
<p style="margin-bottom:0;margin-top:0;background-color:silver">
	Desconto <input type="radio" name="estilo" value="A"> Antecipado
	<input type="radio" name="estilo" value="P" checked> Postergado</p>
  
<p><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></p>
</form>
<%
else

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
if lastfunc<>rs("chapa") then totalf=0
totalf=totalf+rs("reemb")
totalg=totalg+rs("reemb")
response.write "<tr>"
for a= 0 to rs.fields.count-1
	if rs.fields(a).type=5 then 
		conteudo=formatnumber(rs.fields(a),2) 
		response.write "<td align=""right"" class=""campor"">&nbsp;" & conteudo & "</td>"
	else 
		conteudo=rs.fields(a)
		response.write "<td class=""campor"">&nbsp;" & conteudo & "</td>"
	end if
	'response.write "<td><font size='1'>&nbsp;" &rs.fields(a) & rs.fields(a).type & "</td>"
next
response.write "<td class=""campor"" align=""right"">" & formatnumber(totalf,2) & "</td>"
response.write "</tr>"
lastfunc=rs("chapa")
rs.movenext
loop
response.write "<tr><td colspan=9 class=grupo align=""right"">" & formatnumber(totalg,2) & "</td></tr>"
response.write "</table>"
rs.close
response.write "<p>"

	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="assmedica" & request.form("ano") & mesq & empresa & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	sql="SELECT CHAPA, Sum(reemb) AS soma FROM assmed_folha " & _
	"WHERE mesbase='" & dataq & "' and sessao='" & sessao & "' GROUP BY CHAPA HAVING Sum(reemb)<>0"
	sql="select chapa, grupo=case when tp='Acerto' and beneficiario like '%via%' then '2via' else " & _
	"case when tp='Acerto' and reemb<0 then 'Devolucao' else " & _
	"case empresa when 'UC' then 'ParticipaçãoU' when 'IP' then 'ParticipaçãoI' when 'BP' then 'ParticipaçãoB' when 'CP' then 'ParticipaçãoC'  " & _
	"when 'V' then 'Odonto' when 'U' then 'AssistenciaU' when 'I' then 'AssistenciaI' when 'BS' then 'AssistenciaB' when 'C' then 'AssistenciaC' end end end, " & _
	"Sum(reemb) AS soma FROM assmed_folha " & _
	"WHERE mesbase='" & dataq & "' and sessao='" & sessao & "' GROUP BY CHAPA, " & _
	"case when tp='Acerto' and beneficiario like '%via%' then '2via' else case when tp='Acerto' and reemb<0 then 'Devolucao' else " & _
	"case empresa when 'UC' then 'ParticipaçãoU' when 'IP' then 'ParticipaçãoI' when 'BP' then 'ParticipaçãoB' when 'CP' then 'ParticipaçãoC' when 'V' then 'Odonto' when 'U' then 'AssistenciaU' when 'I' then 'AssistenciaI' when 'BS' then 'AssistenciaB' when 'C' then 'AssistenciaC' end end end "
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		chapa=espaco1(rs("chapa"),16)
		if rs("grupo")="AssistenciaU" then evento="076U":fator=1
		if rs("grupo")="AssistenciaB" then evento="076B":fator=1
		if rs("grupo")="AssistenciaC" then evento="076C":fator=1
		if rs("grupo")="AssistenciaI" then evento="076I":fator=1
		if rs("grupo")="Odonto" then evento="025L":fator=1
		if rs("grupo")="ParticipaçãoI" then evento="052I":fator=1
		if rs("grupo")="ParticipaçãoB" then evento="052B":fator=1
		if rs("grupo")="ParticipaçãoC" then evento="052C":fator=1
		if rs("grupo")="ParticipaçãoU" then evento="052U":fator=1
		if rs("grupo")="Devolucao" then evento="149":fator=-1
		if rs("grupo")="2via" then evento="079":fator=1
		evento=espaco1(evento,4)
		lancamento=rs("soma")*fator
		valor=espaco1(replace(formatnumber(lancamento,2),".",""),15)
		leitura.writeline chapa & ";" & evento & ";" & valor & ";001;01"
	rs.movenext
	loop
	rs.close
	sqlt="select CHAPA, 'soma'=SUM(valor_plano) from assmed_folha where mesbase='" & dataq & "' and sessao='" & sessao & "' and empresa not in ('BP','IP') group by chapa "
	rs.Open sqlt, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		chapa=espaco1(rs("chapa"),16)
		evento="CAM":fator=1
		evento=espaco1(evento,4)
		lancamento=rs("soma")*fator
		valor=espaco1(replace(formatnumber(lancamento,2),".",""),15)
		leitura.writeline chapa & ";" & evento & ";" & valor & ";001;01"
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
<a href="..\temp\<%=nomefile%>">Arquivo Desconto Assistência Médica</a>
<%
end if 'request.form 
set rs=nothing
conexao.close
set conexao=nothing
%> 
</body>
</html>