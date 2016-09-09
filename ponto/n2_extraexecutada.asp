<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 60000
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a46")="N" or session("a46")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Relatório de autorização de extras executadas</title>
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
	
if request.form="" then
%>
<p class=titulo>Relatório de Autorização de Horas
<form method="POST" action="n2_extraexecutada.asp">
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo>Verificar Extras Executadas e não autorizadas</td></tr>
<%
hoje=int(now())
diasem=weekday(hoje)
d2=hoje - (diasem-1)
d1=d2-6
%>
<tr>
	<td class=titulo>de <input type="text" name="d1" value="<%=d1%>" size="9"> até <input type="text" name="d2" value="<%=d2%>" size="9"></td>
</tr>
<tr><td class=titulo>
	Tipo de funcionário: <select name="tipofunc" size="1">
	<option value="F">Só funcionários</option>
	<option value="C">Só Chefes</option>
</select>
</td></tr>

<tr><td class=titulo>
	Apenas Seção: <select name="selsecao" size="1">
	<option value="T">Todas seções</option>
<%
sqls="SELECT distinct f.CODSECAO, f.Secao FROM corporerm.dbo.AAFHTFUN h inner join qry_funcionarios f on f.CHAPA=h.chapa where h.DATA>getdate()-60 and f.CODSINDICATO<>'03' and h.EXTRAEXECUTADO>h.EXTRAAUTORIZADO order by secao"
rs.Open sqls, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
	<option value="<%=rs("codsecao")%>"><%=rs("secao")%></option>
<%
rs.movenext
loop
rs.close
%>
</select>
</td></tr>
<tr><td class=titulo>
	<input type="text" value="" size="5" maxlength="5" name="ch1">
	<input type="text" value="" size="5" maxlength="5" name="ch2">
	<input type="text" value="" size="5" maxlength="5" name="ch3">
	<input type="text" value="" size="5" maxlength="5" name="ch4">
	<input type="text" value="" size="5" maxlength="5" name="ch5">
</td></tr>

<tr><td class=titulo>
	Imprimir apenas os faltantes da 1ª remessa: <input type="checkbox" name="imprimir_resto" value="ON">
	<br>Apagar a lista de faltantes:  <input type="checkbox" name="apagar_resto" value="ON">
</td></tr>

<tr><td class=titulo>
	Imprimir totais das colunas? <input type="checkbox" name="print_total" value="ON">
</td></tr>

<tr><td colspan=3 class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3">
</td></tr>
</table>
</form>
<hr>
<%
else 'request.form <>''
	'response.write request.form
	datai=request.form("d1")
	dataf=request.form("d2")
	tipof=request.form("tipofunc")
	linha=0:pagina=0:charchefe=64
	itotal=request.form("print_total")
	ttee=0
	ttea=0
	tten=0

'recalcula falta de justificativa para não gerar no relatorio de extras
	sql0="delete from _marcacoes_checagem "
	sql1="insert into _marcacoes_checagem (chapa, data)  select a.chapa, a.data  from n3_justificativa_s1 a " 
	sql2="delete from _marcacoes_checagem where data not between '" & dtaccess(datai) & "' and '" & dtaccess(dataf) & "' "
	conexao.Execute sql0, , adCmdText
	conexao.Execute sql1, , adCmdText
	conexao.Execute sql2, , adCmdText
'final do recalcula

	sqlcheck="select chapa from n2faltantes"
	rs.Open sqlcheck, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
	else
		sql1="insert into n2faltantes (chapa) select distinct chapa from _marcacoes_checagem"
		if teste=1 then conexao.Execute sql1, , adCmdText
	end if
	rs.close
	
sqld="declare @datai as datetime " & _
"set @datai='" & dtaccess(datai) & "' " & _
"SELECT h.chapa, f.nome, h.data, h.EXTRAEXECUTADO, h.EXTRAAUTORIZADO, f.CODSECAO, f.Secao, s.saldo " & _
", chefe=(select top 1 CHAPASUBST from corporerm.dbo.PSUBSTCHEFE where CODSECAO=f.codsecao and (DATAFIM is null or DATAFIM>GETDATE()) order by master desc, DATAINICIO desc) " & _
"FROM corporerm.dbo.AAFHTFUN h inner join qry_funcionarios f on f.CHAPA=h.chapa " & _
"left join ( " & _
"	select chapa, saldo=sum(saldo) from (" & _
"	select chapa, fimper, saldo=extraant+EXTRAATU-ATRASOANT-ATRASOATU-FALTAANT-FALTAATU " & _
"	from corporerm.dbo.ASALDOBANCOHOR s where FIMPER=(select top 1 fimper from corporerm.dbo.asaldobancohor where FIMPER<=@datai order by FIMPER desc) " & _
"	union " & _
"	select chapa, data, case when CODEVENTOPONTO='0001' or CODEVENTOPONTO='0002' then -1 else 1 end * valor " & _
"	from corporerm.dbo.ABANCOHORFUNDETALHE where DATA>(select top 1 fimper from corporerm.dbo.asaldobancohor where FIMPER<=@datai order by FIMPER desc) and DATA<@datai " & _
"	) z  " & _
"	group by CHAPA " & _
") s on s.chapa=h.chapa " & _
"where h.DATA between @datai and '" & dtaccess(dataf) & "'  " & _
"and f.CODSINDICATO<>'03' and h.EXTRAEXECUTADO>h.EXTRAAUTORIZADO " & _
"and f.chapa not in ('00099','00554','02297','02538','02653') "

sqld1=" and f.chapa collate database_default not in (select distinct chapa from _marcacoes_checagem) " 

if tipof="F" then
	sqle=" and h.chapa not in (select distinct CHAPASUBST from corporerm.dbo.PSUBSTCHEFE where (DATAFIM is null or DATAFIM>GETDATE()) and master=1 ) "
else
	sqle=" and h.chapa in (select distinct CHAPASUBST from corporerm.dbo.PSUBSTCHEFE where (DATAFIM is null or DATAFIM>GETDATE()) and master=1   ) "
end if

if request.form("selsecao")<>"T" then
	sqlf=" and f.codsecao='" & request.form("selsecao") & "' "
else
	sqlf=""
end if
ch1=request.form("ch1"):ch2=request.form("ch2"):ch3=request.form("ch3"):ch4=request.form("ch4"):ch5=request.form("ch5")
if ch1<>"" or ch2<>"" or ch3<>"" or ch4<>"" or ch5<>"" then
	chapas=" and h.chapa in ("
	if ch1<>"" then chapas=chapas & "'" & ch1 & "'"
		if ch1<>"" and ch2<>"" then chapas=chapas  & ","
	if ch2<>"" then chapas=chapas & "'" & ch2 & "'"
		if ch2<>"" and ch3<>"" then chapas=chapas  & ","
	if ch3<>"" then chapas=chapas & "'" & ch3 & "'"
		if ch3<>"" and ch4<>"" then chapas=chapas  & ","
	if ch4<>"" then chapas=chapas & "'" & ch4 & "'"
		if ch4<>"" and ch5<>"" then chapas=chapas  & ","
	if ch5<>"" then chapas=chapas & "'" & ch5 & "'"
	chapas=chapas & ") "
end if
sqlg=chapas

if request.form("imprimir_resto")="ON" then sqlh=" and h.chapa collate database_default in (select chapa from n2faltantes) " : sqld1=""

sqlz="order by chefe, codsecao, f.nome, h.CHAPA, h.DATA " 

sqlc=sqld & sqld1 & sqle & sqlf & sqlg & sqlh & sqlz

rs.CursorLocation = adUseClient
rs.Open sqlc, , adOpenStatic, adLockReadOnly

totalpag=int(rs.recordcount/65)+1
do while not rs.eof
if isnull(rs("chefe")) then chefe=rs("codsecao") else chefe=rs("chefe")

if request.form("apagar_resto")="ON" then
	sqlapaga="delete from n2faltantes"
	conexao.Execute sqlapaga, , adCmdText
end if

if linha=0 or linha>64 or ((ultchefe<>chefe) and tipof="F") then
	if linha<>0 then
		pagina=pagina+1
		response.write "<tr><td class=campo colspan=10 height=15></td></tr>"
		response.write "<tr><td class=campo colspan=3>Osasco, ______________________________</td>"
		response.write "<td class=campo colspan=7>___________________________________________<br>Assinatura do Chefe</td></tr>"
		response.write "<tr><td class=""campor"" colspan=7 style='border-top:1px solid #000000'>Página " & pagina & "/" & totalpag & " - " & now() & " - (1) Saldo sujeito à alteração.</td><td class=""campor"" colspan=4 style='border-top:1px solid #000000'><i><b>Favor devolver em até 48 horas.</i></td></tr>"
		response.write "</table>"
		response.write "<DIV style=""page-break-after:always""></DIV>"
	end if
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="990">
<tr><td class=titulo colspan=11 align="center" style="border:1px solid #000000">
	Relatório de Extras Executadas para Autorização de Inclusão - De <%=datai%> a <%=dataf%></td></td>
<tr>
	<td class=titulo valign="middle" align="center" style="border:1px solid black" rowspan=2 colspan=2 height=20>Funcionário</td>
	<td class=titulo valign="middle" align="center" style="border:1px solid black" rowspan=2>Saldo <font style="font-size:6px">(1)</font><br>em <%=formatdatetime(cdate(datai)-1,2)%></td>
	<td class=titulo valign="middle" align="center" style="border:1px solid black" rowspan=2>Data</td>
	<td class=titulo valign="middle" align="center" style="border:1px solid black" rowspan=2>Dia</td>
	<td class=titulo valign="middle" align="center" style="border:1px solid black" rowspan=2>Marcações</td>
	<td class=titulo valign="middle" align="center" style="border:1px solid black" colspan=3>Horas</td>
	<td class=titulo valign="middle" align="center" style="border:1px solid black" rowspan=2 colspan=2>Informação</td>
</tr>
<tr>
	<td class=titulor valign="middle" align="center" style="border:1px solid black">Executadas</td>
	<td class=titulor valign="middle" align="center" style="border:1px solid black">Autorizadas</td>
	<td class=titulor valign="middle" align="center" style="border:1px solid black">Não Autor.</td>
</tr>
<%
	if linha<>0 then linha=0
end if 'linha
if ultsecao<>rs("codsecao") then
	charchefe=charchefe+1
	response.write "<tr><td class=campo colspan=11 height=15></td></tr>"
	response.write "<tr><td class=""campol"" colspan=5 style='border-top:1px solid #000000'><b><i>Seção: " & rs("codsecao") & " - " & rs("secao") & "</td>"
	response.write "<td class=""campot"" colspan=4 style='border-top:1px solid #000000'>" & rs("chefe") & "</td>"
	response.write "<td class=""campot"" style='border-top:1px solid #000000'>Autorizo</td>"
	response.write "<td class=""campot"" style='border-top:1px solid #000000'>Não Autorizo</td>"
	response.write "</tr>"
	linha=linha+2
end if

Ext_NA=rs("extraexecutado")-rs("extraautorizado")
if rs("chapa")<>ultchapa then 
	cab=1
	response.write rs.absoluteposition
	response.write rs.recordcount
	if itotal="ON" and rs.absoluteposition>1 then
		response.write "<tr>"
		response.write "<td class=campo colspan=6></td>"
		response.write "<td class=campo align=""center"" style='border-top:2px double'>" & horaload(ttee,1) & "</td>"
		response.write "<td class=campo align=""center"" style='border-top:2px double'>" & horaload(ttea,1) & "</td>"
		response.write "<td class=campo align=""center"" style='border-top:2px double'>" & horaload(tten,1) & "</td>"
		response.write "<td class=campo colspan=2></td>"
		response.write "</tr>"
	end if
	ttee=0:ttea=0:tten=0
else 
	cab=0
end if
'obs=rs.absoluteposition & "-" & obs 
if rs("saldo")<0 then sinal="<font color=red>-" else sinal="<font color=blue>+"
%>
<tr>
<%if cab=1 then%>
	<td class=campo style="border-top:1px solid black;border-left:1px dotted black"><%=rs("chapa")%></td>
	<td class=campo style="border-top:1px solid black;border-left:1px dotted black"><%=rs("nome")%></td>
	<td class=campo style="border-top:1px solid black;border-left:1px dotted black" align="center"><%=sinal & horaload(abs(rs("saldo")),1)%></td>
<%else%>
	<td class=campo colspan=3 style="border-left:1px dotted black">&nbsp;</td>
<%
end if
if cab=1 then estilo="border-top:1px solid #000000;border-left:1px dotted black" else estilo="border-top:0px dotted #000000;border-left:1px dotted black"
%>
	<td class=campo style="<%=estilo%>" align="center" ><%=rs("data")%></td>
	<td class=campo style="<%=estilo%>" align="center" ><%=weekdayname(weekday(rs("data")),1)%></td>
	<td class=campo style="<%=estilo%>" align="left" >
<%
sqlb="select batida from corporerm.dbo.abatfun where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' order by batida "
rs2.Open sqlb, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
if rs2.absoluteposition<rs2.recordcount then response.write horaload(rs2("batida"),1) & " | " else response.write horaload(rs2("batida"),1)
rs2.movenext
loop
rs2.close
%>
	</td>
	<td class=campo style="<%=estilo%>" align="center">&nbsp;<%=horaload(rs("extraexecutado"),1)%></td>
	<td class=campo style="<%=estilo%>" align="center">&nbsp;<%=horaload(rs("extraautorizado"),1)%></td>
	<td class=campo style="<%=estilo%>" align="center">&nbsp;<%=horaload(Ext_NA,1)%></td>
	<td class=campo style="<%=estilo%>">[&nbsp;&nbsp;&nbsp;] ________________</td>
	<td class=campo style="<%=estilo%>">[&nbsp;&nbsp;&nbsp;] ________________</td>
</tr>
<%
linha=linha+1
ultchapa=rs("chapa")
ultsecao=rs("codsecao")
'if isnull(rs("chefe")) then ultchefe="" else 
ultchefe=chefe
ttee=ttee+rs("extraexecutado") : ttea=ttea+rs("extraautorizado") : tten=tten+Ext_NA

rs.movenext
loop
rs.close
pagina=pagina+1

response.write "<tr>"
response.write "<td class=campo colspan=6></td>"
response.write "<td class=campo align=""center"" style='border-top:2px double'>" & horaload(ttee,1) & "</td>"
response.write "<td class=campo align=""center"" style='border-top:2px double'>" & horaload(ttea,1) & "</td>"
response.write "<td class=campo align=""center"" style='border-top:2px double'>" & horaload(tten,1) & "</td>"
response.write "<td class=campo colspan=2></td>"
response.write "</tr>"

response.write "<tr><td class=campo colspan=11 height=15></td></tr>"
response.write "<tr><td class=campo colspan=3>Osasco, ______________________________</td>"
response.write "<td class=campo colspan=8>___________________________________________<br>Assinatura do Chefe</td></tr>"

%>
<tr><td class="campor" colspan=7 style='border-top:1px solid #000000'>Página <%=pagina & "/" & totalpag%> - <%=now()%> - (1) Saldo sujeito à alteração.</td>
	<td class="campor" colspan=4 style='border-top:1px solid #000000'><i><b>Favor devolver em até 48 horas.</i></td></tr>

</table>

<%
end if ' request.form	
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>