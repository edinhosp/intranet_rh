<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("acesso")>2 then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
'accesso func 1 prof 2
if session("a100")="N" or session("a100")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Informações do Professor</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
dim formato(4)
if request.form("chapa")<>"" then 
	chapa=request.form("chapa") 
else 
	chapa=session("usuariomaster")
end if
sql1="select count(chapa) total from grades_disp where chapa='" & chapa & "'"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
total=rs("total")
rs.close

dim m(6,7), v(6,7), n(6,7)
corcheck="black"

if total=0 then
	for a=2 to 7
		sql2="insert into grades_disp (chapa, diasem) select '" & chapa & "'," & a & ""
		conexao.execute sql2
	next
end if
if total>0 and total<6 then
	for a=2 to 7
	sql4="select chapa from grades_disp where chapa='" & chapa & "' and diasem=" & a & ""
	rs.Open sql4, ,adOpenStatic, adLockReadOnly
	if rs.eof then
		sql5="insert into grades_disp (chapa, diasem) select '" & chapa & "'," & a & ""
		conexao.execute sql5
	end if
	rs.close
	next
end if

if (request.form<>"" and session("acesso")=2) or (request.form("Salvar")<>"" and session("acesso")=1) then
for m1=1 to 6
	for d=2 to 7
		m(m1,d)=0
		if request.form("m("&m1&","&d&")")="on" then m(m1,d)=1
		'response.write " / " & m1 & "-" & d & "-> ":response.write request.form("m("&m1&","&d&")")
	next
next

for v1=1 to 6
	for d=2 to 7
		v(v1,d)=0
		if request.form("v("&v1&","&d&")")="on" then v(v1,d)=1
		'response.write " / " & v1 & "-" & d & "-> ":response.write request.form("v("&v1&","&d&")")
	next
next

for n1=1 to 6
	for d=2 to 7
		n(n1,d)=0
		if request.form("n("&n1&","&d&")")="on" then n(n1,d)=1
		'response.write " / " & n1 & "-" & d & "-> ":response.write request.form("n("&n1&","&d&")")
	next
next
for u=2 to 7
	sql9="update grades_disp set "
	for t=1 to 6
		sql9=sql9 & "m0"&t&"="&m(t,u) & ","
		sql9=sql9 & "v0"&t&"="&v(t,u) & ","
		sql9=sql9 & "n0"&t&"="&n(t,u) & ","
	next
	sql9=left(sql9,len(sql9)-1)
	sql9=sql9 & " where chapa='" & chapa & "' and diasem=" & u & ""
	if session("master")=1 or session("usuariogrupo")="RH" then
		conexao.execute sql9
	else
		if u=7 then response.write "<script language='JavaScript' type='text/javascript'>alert('Alteração não permitida!');</script>"	
	end if
next
end if 'request.form<>""

sql6="select chapa, diasem, m01,m02,m03,m04,m05,m06, v01,v02,v03,v04,v05,v06, n01,n02,n03,n04,n05,n06 from grades_disp where chapa='" & chapa & "'"
rs.Open sql6, ,adOpenStatic, adLockReadOnly
do while not rs.eof
	m(1,rs("diasem"))=rs("m01"):m(2,rs("diasem"))=rs("m02"):m(3,rs("diasem"))=rs("m03"):m(4,rs("diasem"))=rs("m04"):m(5,rs("diasem"))=rs("m05"):m(6,rs("diasem"))=rs("m06")
	v(1,rs("diasem"))=rs("v01"):v(2,rs("diasem"))=rs("v02"):v(3,rs("diasem"))=rs("v03"):v(4,rs("diasem"))=rs("v04"):v(5,rs("diasem"))=rs("v05"):v(6,rs("diasem"))=rs("v06")
	n(1,rs("diasem"))=rs("n01"):n(2,rs("diasem"))=rs("n02"):n(3,rs("diasem"))=rs("n03"):n(4,rs("diasem"))=rs("n04"):n(5,rs("diasem"))=rs("n05"):n(6,rs("diasem"))=rs("n06")
rs.movenext
loop
rs.close

%>

<!-- -->
<form method="POST" action="disponibilidade.asp" name="form">

<%
if session("acesso")=2  then 'or session("usuariogrupo")="COORD.CURSO" then
%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" >
<tr><td valign=top style="border-right:3px double silver;border-bottom:3px double silver" width=150 height=600>
<!-- -->
<p style="margin-top:0;margin-bottom:0" class=titulo><%=session("usuarioname")%></p>
<hr>
<p style="margin-top:0;margin-bottom:5">
<img src="../images/Clock.gif" width="16" height="16" border="0" alt="">Disponibilidade</p>

<p style="margin-top:0;margin-bottom:5"><a href="../academic/aderencia.asp">
<img src="../images/BookO.gif" width="16" height="16" border="0" alt="">Aderência</a></p>

<br><br>
<p style="margin-top:0;margin-bottom:5"><a href="../academic/meusplanos.asp">
<img src="../images/BookO.gif" width="16" height="16" border="0" alt="">Plano de Ensino</a>

<br><br>
<p style="margin-top:0;margin-bottom:5"><a href="../academic/espelho.asp">
<img src="../images/espelho.jpg" width="16" height="16" border="0" alt="">Marcação de Ponto</a></p>

<br><br><br><br><br><br><br>
<p style="margin-top:0;margin-bottom:0"><a href="../indexp.asp">
<img src="../images/setafirst0.gif" width="12" height="12" border="0" alt="">Início</a>
<!-- -->
</td><td valign=top style="border-bottom:3px double silver">
<p style="margin-top:0;margin-bottom:10" class=titulo>Disponibilidade de Horários</p>
<!-- -->
<%
else ' para acesso=1
%>
<p style="margin-top:0;margin-bottom:10" class=titulo>Disponibilidade de Horários</p>
<select size=1 name="chapa" onchange="javascript:submit();">
	<option value="0">Selecione....</option>
<%
sqlc="select atualizado=case when dataa<getdate() then 1 else 0 end, f.chapa, nome from grades_aux_prof f left join (select chapa, dataa=max(dataa) from grades_disp group by chapa) d on d.chapa=f.chapa where codsituacao<>'D' and f.chapa<'10000' order by nome"
sqlc="select f.chapa, nome, disponivel from grades_aux_prof f left join (select chapa, disponivel=sum(convert(int,m01)+convert(int,m02)+convert(int,m03)+convert(int,m04)+convert(int,m05)+convert(int,m06)+convert(int,v01)+convert(int,v02)+convert(int,v03)+convert(int,v04)+convert(int,v05)+convert(int,v06)+convert(int,n01)+convert(int,n02)+convert(int,n03)+convert(int,n04)+convert(int,n05)+convert(int,n06)) from grades_disp group by chapa) d on d.chapa=f.chapa where codsituacao<>'D' and f.chapa<'10000' order by nome"
rs.Open sqlc, ,adOpenStatic, adLockReadOnly
do while not rs.eof
if rs("chapa")=chapa then txt="selected" else txt=""
if rs("disponivel")>0 then estilo="style='background:CCFFCC;'" else estilo="" 'estilo="style='background:FFFFFF;'"
%>
	<option <%=estilo%> value="<%=rs("chapa")%>" <%=txt%>><%=rs("nome")%></option>
<%
rs.movenext
loop
rs.close
%>
	<option value="02379">Edson teste</option>
</select>
<%
end if
%>

<table border="0" cellpadding="3" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" >
<tr>
	<td class=fundop colspan=2 style="border-bottom:1px solid"></td>
<%for a=2 to 7%>	
	<td class="campop" style="border-right:1px solid;border-left:1px solid;border-bottom:1px solid"><b><%=weekdayname(a,0)%></td>
<%next%>
	<td class=fundop colspan=2 style="border-bottom:1px solid"></td>
</tr>
<%
linha=0 '*******
formato(0)="style='background-color:#ccCCcc'"
formato(1)="style='background-color:#FFFFFF'"
corletra="white"
estilo="solid"
sql="select distinct descricao, periodo from g2defhor h, eturnos t where codtn=1 and h.codtn=t.codturno and tipocurso=2 "
rs.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs.eof
posicao=rs.absoluteposition
%>
<tr>
<%
if rs.absoluteposition=1 then
	response.write "<td class=""campoz"" style='color:" & corletra & ";border-left:1px solid;' rowspan=6 width=20 align=""center""><b>"
	for a=1 to len(rs("periodo"))
		response.write ucase(mid(rs("periodo"),a,1))
		if a<len(rs("periodo")) then response.write "<br>"
	next
	response.write "</td>"
end if
%>
	<td class="campop" height=30 <%=formato(linha)%> style="border-right:1px solid"><%=rs("descricao")%></td>
<%for a=2 to 7%>	
	<td class="campop" align="center" <%=formato(linha)%> style="border-right:1px <%=estilo%>">
<!-- escolhas  -->

	<input <%if session("acesso")=2 then response.write "onclick=""javascript:submit();"""%> type="checkbox" name="m(<%=posicao%>,<%=a%>)" value="on" <%if m(posicao,a)=true then response.write "checked style='background:" & corcheck &";'"%> >

<!-- escolhas  -->
	</td>
<%next%>
	<td class="campop" <%=formato(linha)%>><%=rs("descricao")%></td>
<%
if rs.absoluteposition=1 then
	response.write "<td class=""campoz"" style='color:" & corletra & ";border-left:1px solid' rowspan=6 width=20 align=""center""><b>"
	for a=1 to len(rs("periodo"))
		response.write ucase(mid(rs("periodo"),a,1))
		if a<len(rs("periodo")) then response.write "<br>"
	next
	response.write "</td>"
end if
%>
</tr>
<%
if linha=0 then linha=1 else linha=0
rs.movenext
loop
rs.close
%>

<tr><td class=grupo style="background-color:silver;border:1px solid" colspan=10 height=5></td></tr>

<%
sql="select distinct descricao, periodo from g2defhor h, eturnos t where codtn=5 and pos<=5 and h.codtn=t.codturno and tipocurso=2 "
rs.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs.eof
posicao=rs.absoluteposition
%>
<tr>
<%
if rs.absoluteposition=1 then
	response.write "<td class=""campoz"" style='color:" & corletra & ";border-left:1px solid' rowspan=5 width=20 align=""center""><b>"
	for a=1 to len(rs("periodo"))
		response.write ucase(mid(rs("periodo"),a,1))
		if a<len(rs("periodo")) then response.write "<br>"
	next
	response.write "</td>"
end if
%>
	<td class="campop" height=34 <%=formato(linha)%> style="border-right:1px solid"><%=rs("descricao")%></td>
<%for a=2 to 7:if a=7 then estatus="disabled" else estatus=""%>	
	<td class="campop" align="center" <%=formato(linha)%> style="border-right:1px <%=estilo%>">
<!-- escolhas  -->

	<input <%if session("acesso")=2 then response.write "onclick=""javascript:submit();"""%> type="checkbox" name="v(<%=posicao%>,<%=a%>)" <%=estatus%> value="on" <%if v(posicao,a)=true then response.write "checked style='background:" & corcheck &";'"%> >

<!-- escolhas  -->
	</td>
<%next%>
	<td class="campop" <%=formato(linha)%>><%=rs("descricao")%></td>
<%
if rs.absoluteposition=1 then
	response.write "<td class=""campoz"" style='color:" & corletra & ";border-left:1px solid' rowspan=5 width=20 align=""center""><b>"
	for a=1 to len(rs("periodo"))
		response.write ucase(mid(rs("periodo"),a,1))
		if a<len(rs("periodo")) then response.write "<br>"
	next
	response.write "</td>"
end if
%>
</tr>
<%
if linha=0 then linha=1 else linha=0
rs.movenext
loop
rs.close
%>

<tr><td class=grupo style="background-color:silver;border:1px solid" colspan=10 height=5></td></tr>

<%
sql="select distinct descricao, periodo from g2defhor h, eturnos t where codtn=3 and codds<>7 and h.codtn=t.codturno and tipocurso=2 "
rs.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs.eof
posicao=rs.absoluteposition
%>
<tr>
<%
if rs.absoluteposition=1 then
	response.write "<td class=""campoz"" style='color:" & corletra & ";border-left:1px solid' rowspan=6 width=20 align=""center""><b>"
	for a=1 to len(rs("periodo"))
		response.write ucase(mid(rs("periodo"),a,1))
		if a<len(rs("periodo")) then response.write "<br>"
	next
	response.write "</td>"
end if
%>
	<td class="campop" height=30 <%=formato(linha)%> style="border-right:1px solid"><%=rs("descricao")%></td>
<%for a=2 to 7:if a=7 then estatus="disabled" else estatus=""%>	
	<td class="campop" align="center" <%=formato(linha)%> style="border-right:1px <%=estilo%>">
<!-- escolhas  -->

	<input <%if session("acesso")=2 then response.write "onclick=""javascript:submit();"""%> type="checkbox" name="n(<%=posicao%>,<%=a%>)" <%=estatus%> value="on" <%if n(posicao,a)=true then response.write "checked style='background:" & corcheck &";'"%> >

<!-- escolhas  -->
	</td>
<%next%>
	<td class="campop" <%=formato(linha)%>><%=rs("descricao")%></td>
<%
if rs.absoluteposition=1 then
	response.write "<td class=""campoz"" style='color:" & corletra & ";border-left:1px solid' rowspan=6 width=20 align=""center""><b>"
	for a=1 to len(rs("periodo"))
		response.write ucase(mid(rs("periodo"),a,1))
		if a<len(rs("periodo")) then response.write "<br>"
	next
	response.write "</td>"
end if
%>
</tr>
<%
if linha=0 then linha=1 else linha=0
rs.movenext
loop
rs.close
%>
</table>
<%if session("acesso")=1 then%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" >
<tr><td class=campo align="center">
<input type="submit" name="Salvar" value="Clique aqui para salvar">
</td></tr></table>
<% end if%>
<!-- -->
<!-- -->
<%
'rs.close
'set rs=nothing
'conexao.close
'set conexao=nothing
%>
<%
if session("acesso")=2 then
%>
<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="../images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<%
end if
%>
<!-- -->
</body>
</html>