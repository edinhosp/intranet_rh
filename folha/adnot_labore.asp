<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Geração de Arquivo de Adicional Noturno</title>
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
	ano=request.form("ano")
	mes=request.form("mes")
	tipo=request.form("tipo")
	mesq=numzero(mes,2)
	dataq=mesq & "/" & ano
	datai=dateserial(ano,mes,1)
	dataf=dateserial(ano,mes+1,1)-1
	sessao=session("usuariomaster")
	
	conexao.execute "delete from adnot_folha where sessao='" & sessao & "'"

if tipo="A" or tipo="T" then
	sql="insert into adnot_folha (sessao, chapa, adn, nona, total, codevento, adncurso, adneve) " & _
	"select '" & sessao & "', a.chapa, sum(a.adicional) adn, round(sum(a.adicional*0.125),0) nona, round(sum(a.adicional*1.125),0) total, null, round(sum(a.adicional*1.125),0), '010' " & _
	"from corporerm.dbo.aafhtfun a, corporerm.dbo.pfunc f " & _
	"where f.chapa=a.chapa and a.data between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "' and a.chapa not in (select chapa from corporerm.dbo.pfsalcmp) and f.codsindicato<>'03' " & _
	"group by a.chapa having sum(a.adicional)>0 "
	conexao.execute sql
end if

if tipo="P" or tipo="T" then
	sql="insert into adnot_folha (sessao, chapa, adn, nona, total, codevento, adncurso, adneve) " & _
	"select '" & sessao & "', t.chapa, t.adn, t.nona, t.total, t.codevento, t.adncurso, " & _
	"adneve=case codevento when '138' then '387' when '255' then '387' when '256' then '387' when '257' then '387' when '258' then '387' when '192' then '373' when '688' then '373' when '108' then '359' when 'RHT' then '387' else adnot end " & _
	"from ( " & _
	"select a.chapa, sum(a.adicional) adn, round(sum(a.adicional*0.125),0) nona, round(sum(a.adicional*1.125),0) total, s.codevento, round(sum(a.adicional*1.125) * (convert(float,min(s.jornada))/max(f.jornadamensal)),0) adncurso " & _
	"from corporerm.dbo.aafhtfun a, corporerm.dbo.pfunc f, corporerm.dbo.pfsalcmp s " & _
	"where a.chapa=f.chapa and f.chapa=s.chapa and a.data between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "' and (f.codsindicato='03' or f.chapa='00374') " & _
	"group by a.chapa, s.codevento having sum(a.adicional)>0 " & _
	") t left join g2cursoeve c on t.codevento collate database_default=c.sal order by chapa, codevento "
	conexao.execute sql
	sql="insert into adnot_folha (sessao, chapa, adn, nona, total, codevento, adncurso, adneve) " & _
	"select '" & sessao & "', a.chapa, sum(a.adicional) adn, round(sum(a.adicional*0.125),0) nona, round(sum(a.adicional*1.125),0) total, null, round(sum(a.adicional*1.125),0), " & _
	"case a.chapa when '01165' then '951N' when '01164' then '952N' when '01057' then '953N' else '373' end " & _
	"from corporerm.dbo.aafhtfun a, corporerm.dbo.pfunc f " & _
	"where f.chapa=a.chapa and a.data between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "' and a.chapa not in (select chapa from corporerm.dbo.pfsalcmp) and f.codsindicato='03' " & _
	"group by a.chapa having sum(a.adicional)>0 " 
	conexao.execute sql
end if

conexao.execute "update adnot_folha set adncurso=adncurso/60 where sessao='" & sessao & "' and adneve not in ('010')"
conexao.execute "update adnot_folha set adneve='010' where sessao='" & sessao & "' and chapa='00374' and codevento='001'"
conexao.execute "update adnot_folha set adneve='384', adncurso=adncurso/60 where sessao='" & sessao & "' and chapa='02859'"
conexao.execute "update adnot_folha set adneve='349', adncurso=adncurso/60 where sessao='" & sessao & "' and chapa='02823'"

sql="select * from adnot_folha where sessao='" & sessao & "' and adncurso<>0 order by chapa, adneve "
end if
%>

<p class=titulo>Geração de arquivo de Adicional Noturno para o RM Labore
<%
if request.form="" then
mesfolha=month(dateserial(year(now),month(now)-1,1))
anofolha=year(dateserial(year(now),month(now)-1,1))
periodo=dateserial(anofolha,mesfolha,1) & " a " & dateserial(anofolha,mesfolha+1,1)-1
%>
<form method="POST" action="adnot_labore.asp">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Mês</td>
	<td class=titulo>Ano</td>
	<td class=titulo>Periodo</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="mes" size="4" value="<%=mesfolha%>" class=a></td>
	<td class=titulo><input type="text" name="ano" size="6" value="<%=anofolha%>"></td>
	<td class=fundo><input type=text name=periodo size=20 class=form_input value="<%=periodo%>"></td>
</tr>
<tr>
	<td class=titulo colspan=3>
	<input type="radio" name="tipo" value="P">Professores
   	<input type="radio" name="tipo" value="A">Administrativos 
   	<input type="radio" name="tipo" value="T" checked>Todas	
	</td>
</tr>
<tr>
	<td class=titulo colspan="3"><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></td>
</tr>
</table>
</form>
<%
else

rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst
response.write "<table border='1' cellpadding='1' cellspacing='1' style='border-collapse: collapse' >"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=""titulor"">&nbsp;" & rs.fields(a).name & "</td>"
next
response.write "</tr>"
do while not rs.eof 
if lastfunc<>rs("chapa") then totalf=0
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
response.write "</tr>"
lastfunc=rs("chapa")
rs.movenext
loop
response.write "</table>"
rs.close
response.write "<p>"

	caminho="c:\inetpub\wwwroot\rh\temp\"
	nomefile="adnot" & request.form("ano") & mesq & tipo & ".txt"
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	sql=sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	rs.movefirst
	do while not rs.eof 
		chapa=espaco1(rs("chapa"),16)
		evento=espaco1(rs("adneve"),4)
		lancamento=rs("adncurso")
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
<a href="..\temp\<%=nomefile%>">Arquivo Adicional Noturno</a>
<%
end if 'request.form 
set rs=nothing
conexao.close
set conexao=nothing
%> 
</body>
</html>