<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a35")="N" or session("a35")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Histórico Ficha Financeira</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, sal_anterior(10)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rsq=server.createobject ("ADODB.Recordset")
Set rsq.ActiveConnection = conexao

if request("chapa")<>"" then chapa=request("chapa") else chapa=request.form("chapa")
if request("nomefunc")<>"" then nomefunc=request("nomefunc") else nomefunc=request.form("nomefunc")

if request.form<>"" then
	anoform=request.form("anofrm")
else
	sql="select top 1 anocomp from corporerm.dbo.pffinanc where chapa='" & chapa & "' group by anocomp order by anocomp desc"
	rsq.Open sql, ,adOpenStatic, adLockReadOnly
	if rsq.recordcount>0 then anoform=rsq("anocomp") else anoform=year(now)
	rsq.close
end if
if request.form("bases")="ON" then sqlff="" else sqlff=" and e.provdescbase<>'B' "

	sessao=session("usuariomaster")
	sql2="delete from temp_ff where sessao='" & sessao & "'"
	conexao.execute sql2
	sqla="SELECT ff.NROPERIODO, ff.CHAPA, ff.ANOCOMP, e.PROVDESCBASE, ff.CODEVENTO, e.DESCRICAO, " & _
	"sum(case ff.mescomp when 1 then ff.valor else 0 end) '1', sum(case ff.mescomp when 2 then ff.valor else 0 end) '2', " & _
	"sum(case ff.mescomp when 3 then ff.valor else 0 end) '3', sum(case ff.mescomp when 4 then ff.valor else 0 end) '4', " & _
	"sum(case ff.mescomp when 5 then ff.valor else 0 end) '5', sum(case ff.mescomp when 6 then ff.valor else 0 end) '6', " & _
	"sum(case ff.mescomp when 7 then ff.valor else 0 end) '7', sum(case ff.mescomp when 8 then ff.valor else 0 end) '8', " & _
	"sum(case ff.mescomp when 9 then ff.valor else 0 end) '9', sum(case ff.mescomp when 10 then ff.valor else 0 end) '10', " & _
	"sum(case ff.mescomp when 11 then ff.valor else 0 end) '11', sum(case ff.mescomp when 12 then ff.valor else 0 end) '12' " & _
	"FROM (select * from corporerm.dbo.PFFINANC where chapa='" & chapa & "' AND ANOCOMP=" &  anoform & " union all " & _
	"select * from corporerm.dbo.PFFINANCCOMPL where chapa='" & chapa & "' AND ANOCOMP=" &  anoform & ") ff " & _
	"INNER JOIN corporerm.dbo.PEVENTO e ON ff.CODEVENTO = e.CODIGO " & _
	"WHERE ff.CHAPA='" &  chapa & "' AND ff.ANOCOMP=" &  anoform & " " & sqlff & _
	"GROUP BY ff.NROPERIODO, ff.CHAPA, ff.ANOCOMP, e.PROVDESCBASE, ff.CODEVENTO, e.DESCRICAO " & _
	"ORDER BY ff.NROPERIODO, e.PROVDESCBASE DESC , ff.CODEVENTO " 

	rsq.Open sqla, ,adOpenStatic, adLockReadOnly
	if rsq.recordcount>0 then
	rsq.movefirst:do while not rsq.eof
	sql2="INSERT INTO temp_ff ( sessao, nroperiodo, chapa, anocomp, provdescbase, codevento, descricao, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12 ) " & _
	"SELECT '" & sessao & "', " & rsq("nroperiodo") & ", '" & rsq("chapa") & "', " & rsq("anocomp") & ", " & _
	"'" & rsq("provdescbase") & "', '" & rsq("codevento") & "', '" & rsq("descricao") & "', " & _
	nraccess(rsq("1")) & ", " & _
	nraccess(rsq("2")) & ", " & _
	nraccess(rsq("3")) & ", " & _
	nraccess(rsq("4")) & ", " & _
	nraccess(rsq("5")) & ", " & _
	nraccess(rsq("6")) & ", " & _
	nraccess(rsq("7")) & ", " & _
	nraccess(rsq("8")) & ", " & _
	nraccess(rsq("9")) & ", " & _
	nraccess(rsq("10")) & ", " & _
	nraccess(rsq("11")) & ", " & _
	nraccess(rsq("12")) & " "
	'response.write sql2 & "<br>"
	conexao.execute sql2
	rsq.movenext:loop
	rsq.close
	
	sqla="SELECT ff.NROPERIODO, ff.CHAPA, ff.ANOCOMP, e.PROVDESCBASE, ff.CODEVENTO, e.DESCRICAO, " & _
	"sum(case ff.mescomp when 1 then ff.ref else 0 end) '1', sum(case ff.mescomp when 2 then ff.ref else 0 end) '2', " & _
	"sum(case ff.mescomp when 3 then ff.ref else 0 end) '3', sum(case ff.mescomp when 4 then ff.ref else 0 end) '4', " & _
	"sum(case ff.mescomp when 5 then ff.ref else 0 end) '5', sum(case ff.mescomp when 6 then ff.ref else 0 end) '6', " & _
	"sum(case ff.mescomp when 7 then ff.ref else 0 end) '7', sum(case ff.mescomp when 8 then ff.ref else 0 end) '8', " & _
	"sum(case ff.mescomp when 9 then ff.ref else 0 end) '9', sum(case ff.mescomp when 10 then ff.ref else 0 end) '10', " & _
	"sum(case ff.mescomp when 11 then ff.ref else 0 end) '11', sum(case ff.mescomp when 12 then ff.ref else 0 end) '12' " & _
	"FROM (select * from corporerm.dbo.PFFINANC where chapa='" & chapa & "' AND ANOCOMP=" &  anoform & " union all " & _
	"select * from corporerm.dbo.PFFINANCCOMPL where chapa='" & chapa & "' AND ANOCOMP=" &  anoform & ") ff " & _
	"INNER JOIN corporerm.dbo.PEVENTO e ON ff.CODEVENTO = e.CODIGO " & _
	"WHERE ff.CHAPA='" &  chapa & "' AND ff.ANOCOMP=" &  anoform & " " & sqlff & _
	"GROUP BY ff.NROPERIODO, ff.CHAPA, ff.ANOCOMP, e.PROVDESCBASE, ff.CODEVENTO, e.DESCRICAO " & _
	"ORDER BY ff.NROPERIODO, e.PROVDESCBASE DESC , ff.CODEVENTO "
	rsq.Open sqla, ,adOpenStatic, adLockReadOnly
	rsq.movefirst:do while not rsq.eof
	sql2="UPDATE temp_ff SET " & _
	"  r1 = " & nraccess(rsq("1")) & _
	", r2 = " & nraccess(rsq("2")) & _
	", r3 = " & nraccess(rsq("3")) & _
	", r4 = " & nraccess(rsq("4")) & _
	", r5 = " & nraccess(rsq("5")) & _
	", r6 = " & nraccess(rsq("6")) & _
	", r7 = " & nraccess(rsq("7")) & _
	", r8 = " & nraccess(rsq("8")) & _
	", r9 = " & nraccess(rsq("9")) & _
	", r10 = " & nraccess(rsq("10")) & _
	", r11 = " & nraccess(rsq("11")) & _
	", r12 = " & nraccess(rsq("12")) & _
	" WHERE sessao='" & sessao & "' AND nroperiodo=" & rsq("nroperiodo") & " AND " & _
	"chapa='" & rsq("chapa") & "' AND anocomp=" & rsq("anocomp") & " AND " & _
	"provdescbase='" & rsq("provdescbase") & "' AND codevento='" & rsq("codevento") & "' "
	'response.write sql2 & "<br>"
	conexao.execute sql2
	rsq.movenext:loop
	end if
	rsq.close	

sql="select * from temp_ff where sessao='" & sessao & "' and chapa='" & chapa & "' and anocomp=" & anoform & " " & _
"order by chapa, nroperiodo, provdescbase desc, codevento "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form method="POST" action="hstfichaf.asp" name="formff">
<input type="hidden" name="chapa" value="<%=chapa%>">
<input type="hidden" name="nomefunc" value="<%=nomefunc%>">
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=765>
<th class=titulo colspan=28>Histórico Ficha Financeira 
<select name="anofrm" onChange="javascript:submit()">
<%
sql="select anocomp from (select anocomp from corporerm.dbo.pffinanc where chapa='" & chapa & "' union all " & _
"select anocomp from corporerm.dbo.pffinanccompl where chapa='" & chapa & "') t group by anocomp order by anocomp desc"
rsq.Open sql, ,adOpenStatic, adLockReadOnly
'if rs2.recordcount>0 then anoform=rs2("anocomp") else anoform=year(now)
rsq.movefirst:do while not rsq.eof
anoform2=rsq("anocomp")
if cint(anoform2)=cint(anoform) then tempano="Selected" else tempano=""
%>
    	<option value="<%=anoform2%>" <%=tempano%>><%=anoform2%></option>
<%
rsq.movenext:loop
rsq.close
if request.form("bases")="ON" then cbases="checked" else cbases=""
%>
</select> - <%=chapa%> - <%=nomefunc%> 
<input type="checkbox" name="bases" value="ON" <%=cbases%> onClick="javascript:submit()">
</th>
<tr>
	<td class=fundor rowspan=2>Cod.</td>
	<td class=fundor rowspan=2>Descr.Evento</td>
	<td class=fundor rowspan=2>Tipo</td>
	<td class=titulor colspan=2 align="center">Jan.</td>
	<td class=titulor colspan=2 align="center">Fev.</td>
	<td class=titulor colspan=2 align="center">Mar.</td>
	<td class=titulor colspan=2 align="center">Abr.</td>
	<td class=titulor colspan=2 align="center">Mai.</td>
	<td class=titulor colspan=2 align="center">Jun.</td>
	<td class=titulor colspan=2 align="center">Jul.</td>
	<td class=titulor colspan=2 align="center">Ago.</td>
	<td class=titulor colspan=2 align="center">Set.</td>
	<td class=titulor colspan=2 align="center">Out.</td>
	<td class=titulor colspan=2 align="center">Nov.</td>
	<td class=titulor colspan=2 align="center">Dez.</td>
	<td class=fundor rowspan=2>Total Ano</td>
</tr>
<tr>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>
	<td class=fundor>Valor</td>
</tr>
<%
if rs.recordcount>0 then
'lastper=rs("nroperiodo")
'lasttipo=rs("provdescbase")
totalano=0:inicio=1
rs.movefirst
do while not rs.eof

if ((lasttipo<>rs("provdescbase")) or (lastper<>rs("nroperiodo"))) and lasttipo<>"B" then
	avar=" > " & lasttipo & " " & len(lasttipo) & " " & rs("provdescbase") & " " & len(rs("provdescbase"))
	if inicio=0 then
		response.write "<tr><td class=""campotr"" colspan=3 align=""right"">Total " & lasttipo & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv1,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv2,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv3,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv4,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv5,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv6,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv7,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv8,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv9,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv10,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv11,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv12,2) & "</td>"
		totalv=totalv1+totalv2+totalv3+totalv4+totalv5+totalv6+totalv7+totalv8+totalv9+totalv10+totalv11+totalv12
		response.write "<td class=""campotr"" align=""right"">" & formatnumber(totalv,2) & "</td></tr>"
	end if
	totalv1=0:totalv2=0:totalv3=0:totalv4=0:totalv5=0:totalv6=0:totalv7=0:totalv8=0:totalv9=0:totalv10=0:totalv11=0:totalv12=0:totalv=0
end if

if lastper<>rs("nroperiodo") then
	if inicio=0 then
		totall1=totalp1-totald1
		totall2=totalp2-totald2
		totall3=totalp3-totald3
		totall4=totalp4-totald4
		totall5=totalp5-totald5
		totall6=totalp6-totald6
		totall7=totalp7-totald7
		totall8=totalp8-totald8
		totall9=totalp9-totald9
		totall10=totalp10-totald10
		totall11=totalp11-totald11
		totall12=totalp12-totald12
		totall=totall1+totall2+totall3+totall4+totall5+totall6+totall7+totall8+totall9+totall10+totall11+totall12
		response.write "<tr><td class=""campolr"" align=""right"" colspan=3>" & "Valor Liquido Periodo" & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall1,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall2,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall3,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall4,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall5,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall6,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall7,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall8,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall9,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall10,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall11,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall12,2) & "</td>"
		response.write "<td class=""campolr"" align=""right"">" & formatnumber(totall,2) & "</td>"
		response.write "</tr>"
	totall1=0:totall2=0:totall3=0:totall4=0:totall5=0:totall6=0:totall7=0:totall8=0:totall9=0:totall10=0:totall11=0:totall12=0:totall=0
	end if
end if

if lastper<>rs("nroperiodo") then
	response.write "<tr><td class=grupor colspan=28>Período: " & rs("nroperiodo") & "</td></tr>"
	lasttipo=""
	totalv1=0:totalv2=0:totalv3=0:totalv4=0:totalv5=0:totalv6=0:totalv7=0:totalv8=0:totalv9=0:totalv10=0:totalv11=0:totalv12=0:totalv=0
	totall1=0:totall2=0:totall3=0:totall4=0:totall5=0:totall6=0:totall7=0:totall8=0:totall9=0:totall10=0:totall11=0:totall12=0:totall=0
	totalp1=0:totalp2=0:totalp3=0:totalp4=0:totalp5=0:totalp6=0:totalp7=0:totalp8=0:totalp9=0:totalp10=0:totalp11=0:totalp12=0:totalp=0
	totald1=0:totald2=0:totald3=0:totald4=0:totald5=0:totald6=0:totald7=0:totald8=0:totald9=0:totald10=0:totald11=0:totald12=0:totald=0
end if

if lasttipo<>rs("provdescbase") then
	select case rs("provdescbase")
		case "P"
			tipoeve="P - Proventos"
		case "D"
			tipoeve="D - Descontos"
		case "B"
			tipoeve="B - Eventos de Base"
	end select
	response.write "<tr><td class=""campoar"" colspan=28>" & tipoeve & "</td></tr>"
end if

if rs("v1")>0 then totalano=totalano+rs("v1"):totalv1=totalv1+cdbl(rs("v1"))
if rs("v2")>0 then totalano=totalano+rs("v2"):totalv2=totalv2+cdbl(rs("v2"))
if rs("v3")>0 then totalano=totalano+rs("v3"):totalv3=totalv3+cdbl(rs("v3"))
if rs("v4")>0 then totalano=totalano+rs("v4"):totalv4=totalv4+cdbl(rs("v4"))
if rs("v5")>0 then totalano=totalano+rs("v5"):totalv5=totalv5+cdbl(rs("v5"))
if rs("v6")>0 then totalano=totalano+rs("v6"):totalv6=totalv6+cdbl(rs("v6"))
if rs("v7")>0 then totalano=totalano+rs("v7"):totalv7=totalv7+cdbl(rs("v7"))
if rs("v8")>0 then totalano=totalano+rs("v8"):totalv8=totalv8+cdbl(rs("v8"))
if rs("v9")>0 then totalano=totalano+rs("v9"):totalv9=totalv9+cdbl(rs("v9"))
if rs("v10")>0 then totalano=totalano+rs("v10"):totalv10=totalv10+cdbl(rs("v10"))
if rs("v11")>0 then totalano=totalano+rs("v11"):totalv11=totalv11+cdbl(rs("v11"))
if rs("v12")>0 then totalano=totalano+rs("v12"):totalv12=totalv12+cdbl(rs("v12"))
if rs("provdescbase")="P" then
	if rs("v1")>0 then totalp1=totalp1+cdbl(rs("v1"))
	if rs("v2")>0 then totalp2=totalp2+cdbl(rs("v2"))
	if rs("v3")>0 then totalp3=totalp3+cdbl(rs("v3"))
	if rs("v4")>0 then totalp4=totalp4+cdbl(rs("v4"))
	if rs("v5")>0 then totalp5=totalp5+cdbl(rs("v5"))
	if rs("v6")>0 then totalp6=totalp6+cdbl(rs("v6"))
	if rs("v7")>0 then totalp7=totalp7+cdbl(rs("v7"))
	if rs("v8")>0 then totalp8=totalp8+cdbl(rs("v8"))
	if rs("v9")>0 then totalp9=totalp9+cdbl(rs("v9"))
	if rs("v10")>0 then totalp10=totalp10+cdbl(rs("v10"))
	if rs("v11")>0 then totalp11=totalp11+cdbl(rs("v11"))
	if rs("v12")>0 then totalp12=totalp12+cdbl(rs("v12"))
end if
if rs("provdescbase")="D" then
	if rs("v1")>0 then totald1=totald1+cdbl(rs("v1"))
	if rs("v2")>0 then totald2=totald2+cdbl(rs("v2"))
	if rs("v3")>0 then totald3=totald3+cdbl(rs("v3"))
	if rs("v4")>0 then totald4=totald4+cdbl(rs("v4"))
	if rs("v5")>0 then totald5=totald5+cdbl(rs("v5"))
	if rs("v6")>0 then totald6=totald6+cdbl(rs("v6"))
	if rs("v7")>0 then totald7=totald7+cdbl(rs("v7"))
	if rs("v8")>0 then totald8=totald8+cdbl(rs("v8"))
	if rs("v9")>0 then totald9=totald9+cdbl(rs("v9"))
	if rs("v10")>0 then totald10=totald10+cdbl(rs("v10"))
	if rs("v11")>0 then totald11=totald11+cdbl(rs("v11"))
	if rs("v12")>0 then totald12=totald12+cdbl(rs("v12"))
end if

%>
<tr>
	<td class="campor"><%=rs("codevento")%></td>
	<td class="campor"><%=lcase(rs("descricao"))%></td>
	<td class="campor"><%=rs("provdescbase")%></td>
	<td class="campor" align="right"><i><%if rs("r1")<>"" and rs("r1")>0 then response.write rs("r1")%></i></td>
	<td class="campor" align="right"><%if rs("v1")<>0 then response.write formatnumber(rs("v1"),2)%></td>
	<td class="campor" align="right"><i><%if rs("r2")<>"" and rs("r2")>0 then response.write rs("r2")%></i></td>
	<td class="campor" align="right"><%if rs("v2")<>0 then response.write formatnumber(rs("v2"),2)%></td>
	<td class="campor" align="right"><i><%if rs("r2")<>"" and rs("r3")>0 then response.write rs("r3")%></i></td>
	<td class="campor" align="right"><%if rs("v3")<>0 then response.write formatnumber(rs("v3"),2)%></td>
	<td class="campor" align="right"><i><%if rs("r4")<>"" and rs("r4")>0 then response.write rs("r4")%></i></td>
	<td class="campor" align="right"><%if rs("v4")<>0 then response.write formatnumber(rs("v4"),2)%></td>
	<td class="campor" align="right"><i><%if rs("r5")<>"" and rs("r5")>0 then response.write rs("r5")%></i></td>
	<td class="campor" align="right"><%if rs("v5")<>0 then response.write formatnumber(rs("v5"),2)%></td>
	<td class="campor" align="right"><i><%if rs("r6")<>"" and rs("r6")>0 then response.write rs("r6")%></i></td>
	<td class="campor" align="right"><%if rs("v6")<>0 then response.write formatnumber(rs("v6"),2)%></td>
	<td class="campor" align="right"><i><%if rs("r7")<>"" and rs("r7")>0 then response.write rs("r7")%></i></td>
	<td class="campor" align="right"><%if rs("v7")<>0 then response.write formatnumber(rs("v7"),2)%></td>
	<td class="campor" align="right"><i><%if rs("r8")<>"" and rs("r8")>0 then response.write rs("r8")%></i></td>
	<td class="campor" align="right"><%if rs("v8")<>0 then response.write formatnumber(rs("v8"),2)%></td>
	<td class="campor" align="right"><i><%if rs("r9")<>"" and rs("r9")>0 then response.write rs("r9")%></i></td>
	<td class="campor" align="right"><%if rs("v9")<>0 then response.write formatnumber(rs("v9"),2)%></td>
	<td class="campor" align="right"><i><%if rs("r10")<>"" and rs("r10")>0 then response.write rs("r10")%></i></td>
	<td class="campor" align="right"><%if rs("v10")<>0 then response.write formatnumber(rs("v10"),2)%></td>
	<td class="campor" align="right"><i><%if rs("r11")<>"" and rs("r11")>0 then response.write rs("r11")%></i></td>
	<td class="campor" align="right"><%if rs("v11")<>0 then response.write formatnumber(rs("v11"),2)%></td>
	<td class="campor" align="right"><i><%if rs("r12")<>"" and rs("r12")>0 then response.write rs("r12")%></i></td>
	<td class="campor" align="right"><%if rs("v12")<>0 then response.write formatnumber(rs("v12"),2)%></td>
	<td class="campor" align="right"><%=formatnumber(totalano,2)%></td>
</tr>
<%
lastper=rs("nroperiodo"):lasttipo=rs("provdescbase"):totalano=0
inicio=0
rs.movenext
loop
	if lasttipo<>"B" then
		response.write "<tr><td class=""campotr"" colspan=3 align=""right"">Total " & lasttipo & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv1,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv2,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv3,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv4,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv5,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv6,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv7,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv8,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv9,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv10,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv11,2) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv12,2) & "</td>"
		totalv=totalv1+totalv2+totalv3+totalv4+totalv5+totalv6+totalv7+totalv8+totalv9+totalv10+totalv11+totalv12
		response.write "<td class=""campotr"" align=""right"">" & formatnumber(totalv,2) & "</td></tr>"
		totalv1=0:totalv2=0:totalv3=0:totalv4=0:totalv5=0:totalv6=0:totalv7=0:totalv8=0:totalv9=0:totalv10=0:totalv11=0:totalv12=0:totalv=0
	end if

		totall1=totalp1-totald1
		totall2=totalp2-totald2
		totall3=totalp3-totald3
		totall4=totalp4-totald4
		totall5=totalp5-totald5
		totall6=totalp6-totald6
		totall7=totalp7-totald7
		totall8=totalp8-totald8
		totall9=totalp9-totald9
		totall10=totalp10-totald10
		totall11=totalp11-totald11
		totall12=totalp12-totald12
		totall=totall1+totall2+totall3+totall4+totall5+totall6+totall7+totall8+totall9+totall10+totall11+totall12
		response.write "<tr><td class=""campolr"" align=""right"" colspan=3>" & "Valor Liquido Periodo" & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall1,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall2,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall3,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall4,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall5,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall6,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall7,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall8,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall9,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall10,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall11,2) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(totall12,2) & "</td>"
		response.write "<td class=""campolr"" align=""right"">" & formatnumber(totall,2) & "</td>"
		response.write "</tr>"
	totall1=0:totall2=0:totall3=0:totall4=0:totall5=0:totall6=0:totall7=0:totall8=0:totall9=0:totall10=0:totall11=0:totall12=0:totall=0

else
	response.write "<tr><td class=campo colspan=28>Sem histórico de ficha financeira</td></tr>"
end if
%>
</table>
</form>

<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>