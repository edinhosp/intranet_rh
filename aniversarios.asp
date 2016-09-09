<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
if session("a2")="N" or session("a2")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Aniversariantes do dia</title>
<link rel="stylesheet" type="text/css" href="<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>
<!-- -->
<table><tr><td>
<!-- -->

<%
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("consql")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
usuario=session("usuariomaster")
if usuario="02379" or usuario="00892" or usuario="00259" or usuario="02159" or usuario="00595" or usuario="02463" or usuario="90261" or usuario="02552" or usuario="02592" or usuario="02675" then fotos=1 else fotos=0

dim emes(12),edia(31)
mesagora=month(now)
anoagora=year(now)
diaagora=day(now)

if request("d")<>"" and request.form("diaagora")="" then diaagora=request("d") 
if request("m")<>"" and request.form("mesagora")="" then mesagora=request("m")
if request.form("diaagora")<>"" and request("d")="" then diaagora=request.form("diaform")
if request.form("mesagora")<>"" and request("m")="" then mesagora=request.form("mesform")
		
if request.form<>"" then
	if request.form("B3")<>"" then
		finaliza=1
	else
		finaliza=0
		mesagora=request.form("mesform")
		diaagora=request.form("diaform")
	end if
	if request.form("avanca")<>"" then
		mesagora=mesagora+1
		if mesagora>12 then	mesagora=1
	end if
	if request.form("volta")<>"" then
		mesagora=mesagora-1
		if mesagora<1 then mesagora=12
	end if
end if

sqld="select day(diaferiado) as dia1 from gferiado " & _
"where month(diaferiado)=" & mesagora & " and year(diaferiado)=" & anoagora & " " & _
"group by day(diaferiado) "
rs.Open sqld, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof 
	edia(rs("dia1"))=1
rs.movenext:loop
end if
rs.close
	
dianiv=diaagora
mesniv=mesagora

sqla="SELECT PFUNC.CHAPA, Month([DTNASCIMENTO]) AS Mes, Day([DTNASCIMENTO]) AS Dia, " & _
"campus=case substring(codsecao,1,2) when '01' then 'Narciso' when '02' then 'Brás' when '03' then 'V.Yara' when '04' then 'Jd.Wilson' end, " & _
"PFUNC.NOME AS Nome_Aniversariante, (psecao.DESCRICAO) AS Setor,  " & _
"corpo=case when codtipo='T' then 'Estagiário' when codsindicato='03' then 'Professor' else 'Administrativo' end, " & _
"PPESSOA.DTNASCIMENTO, PPESSOA.EMAIL AS Expr3 " & _
"FROM (PFUNC INNER JOIN PPESSOA ON PFUNC.CODPESSOA = PPESSOA.CODIGO) INNER JOIN PSECAO ON PFUNC.CODSECAO = PSECAO.CODIGO " & _
"WHERE (PFUNC.CHAPA<'10000' Or PFUNC.CHAPA>='90000') " & _
"AND Month([DTNASCIMENTO])=" & mesniv & " " & _
"AND Day([DTNASCIMENTO])=" & dianiv & " " & _
"AND PFUNC.CODSITUACAO<>'D' " & _
"ORDER BY Month([DTNASCIMENTO]), Day([DTNASCIMENTO]), Campus, PFUNC.NOME "
'response.write sqla
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<!-- calencario -->
<%
diasemana=weekday(dateserial(anoagora,mesagora,1))
ultimodia=day(dateserial(anoagora,mesagora+1,1)-1)
ultimo=0
emes(1)="Janeiro":emes(2)="Fevereiro":emes(3)="Março":emes(4)="Abril":emes(5)="Maio":emes(6)="Junho"
emes(7)="Julho":emes(8)="Agosto":emes(9)="Setembro":emes(10)="Outubro":emes(11)="Novembro":emes(12)="Dezembro"
%>
<form method="POST" action="aniversarios.asp" name="form">

<input type="hidden" name="mesform" value="<%=mesagora%>">
<input type="hidden" name="diaform" value="<%=diaagora%>">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="175">
<tr>
	<td class="campo"><input type="submit" value="&lt;" name="volta" class="button"></td>
	<td class="campor" width="100%" align="center">
		<font color="#000080"><b><%=emes(mesagora)& "/" & anoagora%></font></td>
	<td class="campo"><input type="submit" value="&gt;" name="avanca" class="button"></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="175">
<tr>
	<td class="campo" align="center">Dom</td>
	<td class="campo" align="center">Seg</td>
	<td class="campo" align="center">Ter</td>
	<td class="campo" align="center">Qua</td>
	<td class="campo" align="center">Qui</td>
	<td class="campo" align="center">Sex</td>
	<td class="campo" align="center">Sab</td>
</tr>
<tr>
<%
for linha=1 to 7
	response.write "<td class=campo align='center'>"
	if linha=diasemana then
		ultimo=1
		if edia(ultimo)=1 or linha=1 then 'é feriado
			response.write "<a href='aniversarios.asp?d=" & ultimo & "&m=" & mesagora & "' class=r style='color:red'>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		else
			response.write "<a href='aniversarios.asp?d=" & ultimo & "&m=" & mesagora & "' class=r>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		end if		
	elseif ultimo>=1 then
		ultimo=ultimo+1
		if edia(ultimo)=1 then
			response.write "<a href='aniversarios.asp?d=" & ultimo & "&m=" & mesagora & "' class=r style='color:red'>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		else
			response.write "<a href='aniversarios.asp?d=" & ultimo & "&m=" & mesagora & "' class=r>"
			if cint(ultimo)=cint(diaagora) then response.write "<b>"
			response.write ultimo
			response.write "</a>"
		end if		
	end if
	response.write "</td>"
next
response.write "</tr>"

vartemp1=ultimodia-ultimo
vartemp2=int(vartemp1/7)
if (vartemp1/7)-vartemp2>0 then vartemp2=vartemp2+1
for sem=1 to vartemp2
	response.write "<tr>"
	for l2=1 to 7
		response.write "<td class=campo align='center'>"
		ultimo=ultimo+1
		if ultimo<=ultimodia then 
			if edia(ultimo)=1 or l2=1 then
				response.write "<a href='aniversarios.asp?d=" & ultimo & "&m=" & mesagora & "' class=r style='color:red'>"
				if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				response.write "</a>"
			else
				response.write "<a href='aniversarios.asp?d=" & ultimo & "&m=" & mesagora & "' class=r>"
				if cint(ultimo)=cint(diaagora) then response.write "<b>"
				response.write ultimo
				response.write "</a>"
			end if
		end if
		response.write "</td>"
	next
	response.write "</tr>"
next
%>
</table>
<!-- fim calencario -->
</form>
<p>

<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="400">
<tr>
<td class="titulo" align="center" width="300"><font size="2">&nbsp;Aniversariantes do dia <%=dianiv%>/<%=mesniv%>&nbsp;</font></td>
<% if fotos=1 then %>
<td class="titulo" align="center" width="100"><font size="2">&nbsp;Foto&nbsp;</font></td>
<% end if %>
</tr>
<%
if rs.recordcount>0 then
	rs.movefirst:do while not rs.eof 
%>
<tr>
<td valign="top"><b><%=rs("nome_aniversariante")%></b><font size="1"> (<%=rs("corpo")%>)</font>
<br><font size="1">Campus: <%=rs("campus") %></font>
<br><font size="1">Setor: <%=rs("setor") %></font>
<br><font size="1">Email: <a href="mailto:<%=rs("expr3")%>"><font size="1"><%=rs("expr3") %></font></a></font></td>

<% if fotos=1 then %>
	<td><font size="1">
	<img border="0" src="func_foto.asp?chapa=<%=rs("chapa")%>" width="100">
	</font></td>
<% end if %>
</tr>
<%
rs.movenext:loop
else
	response.write "<tr><td colspan='2'><font size='2'>&nbsp;Não há aniversariantes no dia de hoje.&nbsp;&nbsp;</font></td></tr>"
end if 'if recordcount
%>
</table>
<%
rs.close

set rs=nothing
conexao.close
set conexao=nothing
%>

<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>