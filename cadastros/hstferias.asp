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
<title>Histórico de Férias</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
chapa=request("chapa")

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
'set rs2=server.createobject ("ADODB.Recordset")
'Set rs2.ActiveConnection = conexao
sqlb="select dtvencferias, saldoferias, inicprogferias1, fimprogferias1, nrodiasferias, " & _
"querabono, nrodiasabono, quer1aparc13o, feriascoletivas, dtpagtoferias, dtavisoferias " & _
"from corporerm.dbo.pfunc where CHAPA='" & chapa & "'"
rs.Open sqlb, ,adOpenStatic, adLockReadOnly	

vencfer=rs("dtvencferias")
limite=dateserial(year(vencfer)+1,month(vencfer),day(vencfer))-1
%>

<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=515>
<tr>
	<td class=fundor align="center">Venci-<br>mento</td>
	<td class=fundor align="center">Saldo</td>
	<td class=fundor align="center">Limite<br>Gozo</td>
	<td class=fundor align="center">Inicio</td>
	<td class=fundor align="center">Final</td>
	<td class=fundor align="center">Dias</td>
	<td class=fundor align="center">Quer<br>Abono</td>
	<td class=fundor align="center">Dias<br>Abono</td> 
	<td class=fundor align="center">Quer<br>1ª Parc.13</td>
	<td class=fundor align="center">Férias<br>Coletivas</td>
	<td class=fundor align="center">Data<br>Pagamento</td>
	<td class=fundor align="center">Data<br>Aviso</td>
</tr>
<tr>
	<td class="campor" align="center"><%=rs("dtvencferias")%></td>
	<td class="campor" align="center"><%=rs("saldoferias")%></td>
	<td class="campor" align="center"><%=limite%></td>
	<td class="campor" align="center"><%=rs("inicprogferias1")%></td>
	<td class="campor" align="center"><%=rs("fimprogferias1")%></td>
	<td class="campor" align="center"><%=rs("nrodiasferias")%></td>
	<td class="campor" align="center"><%if rs("querabono")=1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor" align="center"><%=rs("nrodiasabono")%></td>
	<td class="campor" align="center"><%if rs("quer1aparc13o")=1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor" align="center"><%if rs("feriascoletivas")=1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor" align="center"><%=rs("dtpagtoferias")%></td>
	<td class="campor" align="center"><%=rs("dtavisoferias")%></td>
</tr>
</table>

<br>
<%
rs.close

sqla="SELECT hf.*, p.DTPAGTO, p.dtvencimento " & _
"FROM corporerm.dbo.pfhstfer_old AS hf LEFT JOIN corporerm.dbo.pfperfer_old AS p ON (hf.DTFIMPERAQUIS = p.DTVENCIMENTO) AND (hf.NROPERIODO = p.NROPERIODO) AND (hf.CHAPA = p.CHAPA) " & _
"WHERE hf.CHAPA='" & chapa & "' ORDER BY hf.DTINIPERAQUIS, hf.NROPERIODO "
sqla="select p.CHAPA, 'dtfimperaquis'=p.FIMPERAQUIS, 'nroperiodo'=null, 'dtiniperaquis'=h.INICIOPERAQUIS, 'dtinigozo'=p.DATAINICIO, 'dtvencimento'=p.fimperaquis, " & _
"'dtfimgozo'=p.DATAFIM, 'diasabono'=p.NRODIASABONO, 'quer1aparc13o'=p.PAGA1APARC13O, p.FERIASCOLETIVAS, 'nrofaltas'=p.FALTAS, NRODIASFERIAS, PERIODOABERTO, 'dtpagto'=p.DATAPAGTO " & _
"from corporerm.dbo.PFUFERIASPER p inner join corporerm.dbo.PFUFERIAS h on h.CHAPA=p.CHAPA and h.FIMPERAQUIS=p.FIMPERAQUIS " & _
"where p.CHAPA='" & chapa & "' order by p.FIMPERAQUIS, p.DATAINICIO "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="0" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr><td valign=top align="left">

<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=490>
<th class=titulo colspan=11>Histórico de Fériass</th>
<tr>
	<td class=titulor align="center" rowspan=2>Per.</td>
	<td class=titulor align="center" colspan=2>Per.Aquisitivo</td>
	<td class=titulor align="center" colspan=2>Gozo</td>
	<td class=titulor align="center" colspan=2>Dias</td>
	<td class=titulor align="center" rowspan=2>1ª <br>Parc.13</td>
	<td class=titulor align="center" rowspan=2>Fér.<br>Coletivas</td>
	<td class=titulor align="center" rowspan=2>Faltas</td>
	<td class=titulor align="center" rowspan=2>Recibo</td>
</tr>
<tr>
	<td class=titulor align="center">Início</td>
	<td class=titulor align="center">Fim</td>
	<td class=titulor align="center">Inicio</td>
	<td class=titulor align="center">Fim</td>
	<td class=titulor align="center">Férias</td>
	<td class=titulor align="center">Abono</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
'sql="select descricao from pcodocortrab where codcliente=" & rs("codocorrencia") & ""
'response.write sql
'rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codocorrencia=rs2("descricao")
'rs2.close
if rs("diasabono")="" or isnull(rs("diasabono")) then diasabono=0 else diasabono=cdbl(rs("diasabono"))
if isnull(rs("dtinigozo")) then diasferias=0 else diasferias=cdate(rs("dtfimgozo"))-cdate(rs("dtinigozo"))+1
%>
<tr>
	<td class="campor" align="center"><%=rs("nroperiodo")%></td>
	<td class="campor" align="center"><%=rs("dtiniperaquis")%></td>
	<td class="campor" align="center"><%=rs("dtfimperaquis")%></td>
	<td class="campor" align="right"><%=rs("dtinigozo")%></td>
	<td class="campor" align="center"><%=rs("dtfimgozo")%></td>
	<td class="campor" align="center"><%=diasferias%></td>
	<td class="campor" align="center"><%=rs("diasabono")%></td>
	<td class="campor" align="center"><%if rs("quer1aparc13o")=1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor" align="center"><%if rs("feriascoletivas")=1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor" align="center"><%=rs("nrofaltas")%></td>
	<td class="campor" align="center">
	<a class="r9" href="hstferias.asp?chapa=<%=rs("chapa")%>&venc=<%=rs("dtvencimento")%>&pagto=<%=rs("dtpagto")%>">
	<%=rs("dtpagto")%></a>
	</td>
</tr>
<%
rs.movenext
loop
else
	response.write "<tr><td class=campo colspan=3>Sem lançamentos cadastrados</td></tr>"
end if
rs.close
%>
</table>

</td><td valign=top align="left">
<!-- quadro recibo -->
<%
if request("venc")<>"" then
chapa=request("chapa")
venc=request("venc")
pagto=request("pagto")
sqlr="SELECT F.CHAPA, F.DTVENCIMENTO, F.NROPERIODO, E.PROVDESCBASE, F.CODEVENTO, E.DESCRICAO, F.REF, F.VALOR, " & _
"base=case PROVDESCBASE when 'D' then -1 else 1 end * VALOR " & _
"FROM corporerm.dbo.pfferias_old F, corporerm.dbo.PEVENTO E WHERE F.CODEVENTO=E.CODIGO AND " & _
"F.CHAPA='" & chapa & "' AND F.DTVENCIMENTO='" & dtaccess(venc) & "' AND F.NROPERIODO=" & periodo & " " & _
"AND E.PROVDESCBASE<>'B' AND F.VALOR<>0 ORDER BY E.PROVDESCBASE DESC , F.CODEVENTO "
sqlr="select r.CHAPA, 'dtvencimento'=r.FIMPERAQUIS, 'nroperiodo'=r.DATAPAGTO, e.PROVDESCBASE, v.CODEVENTO, e.DESCRICAO, v.REF, v.VALOR, " & _
"base=case PROVDESCBASE when 'D' then -1 else 1 end * VALOR " & _
"from corporerm.dbo.PFUFERIASRECIBO r inner join corporerm.dbo.PFUFERIASVERBAS v on v.CHAPA=r.CHAPA and v.FIMPERAQUIS=r.FIMPERAQUIS and v.DATAPAGTO=r.DATAPAGTO inner join corporerm.dbo.PEVENTO e on e.CODIGO=v.CODEVENTO " & _
"where r.CHAPA='" & chapa & "' and r.FIMPERAQUIS='" & dtaccess(venc) & "' and r.DATAPAGTO='" & dtaccess(pagto) & "' " & _
"AND E.PROVDESCBASE<>'B' AND v.VALOR<>0 ORDER BY E.PROVDESCBASE DESC, v.CODEVENTO "
rs.Open sqlr, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=270>
<th class=titulo colspan=11>Recibo de Férias em <%=request("pagto")%></th>
<tr>
	<td class=titulor align="center">Cod.</td>
	<td class=titulor align="center">Descrição</td>
	<td class=titulor align="center">Ref.</td>
	<td class=titulor align="center">Rend.</td>
	<td class=titulor align="center">Desc.</td>
</tr>
<%
liquido=0
rs.movefirst
do while not rs.eof
if rs("provdescbase")="D" then vencimento="&nbsp;" else desconto="&nbsp;"
if rs("provdescbase")="P" then vencimento=formatnumber(rs("valor"),2) else desconto=formatnumber(rs("valor"),2)
liquido=liquido + cdbl(rs("base"))
%>
<tr>
	<td class="campor" align="center"><%=rs("codevento")%></td>
	<td class="campor" align="left"><%=rs("descricao")%></td>
	<td class="campor" align="center"><%=rs("ref")%></td>
	<td class="campor" align="right"><%=vencimento%></td>
	<td class="campor" align="right"><%=desconto%></td>
</tr>
<%
rs.movenext
loop
%>
<tr>
	<td class=titulor colspan=4>Líquido de Férias</td>
	<td class="campor" align="right"><%=formatnumber(liquido,2)%>
</tr>
</table>
<%
end if 'rs.recordcount
rs.close
%>


<%
end if 'request(pagto)
%>
<!-- quadro recibo -->
</td></tr></table>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>