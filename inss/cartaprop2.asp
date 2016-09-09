<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a39")="N" or session("a39")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Proporcionalidade</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<script language="VBScript">
	Sub Calcula()
	End Sub
</script></head>
<body>
<%
'response.write request.form
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
sqla="SELECT p.carteiratrab, p.seriecarttrab from qry_funcionarios p where chapa='" & session("chapa") & "' "
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
rs.Open sqla, ,adOpenStatic, adLockReadOnly
ctps=rs("carteiratrab")
serie=rs("seriecarttrab")
rs.close
'session("valorcarta")=request.form("valorcarta")

vsc=5189.82
vinss=570.88

if request.form<>"" then
	'if request.form("salario1")<>"" then salario1=cdbl(request.form("valorcarta")) else salario1=cdbl(request.form("salario1"))
	if request.form("salario2")<>"" then salario2=cdbl(request.form("salario2")) else salario2=0
	if request.form("salario3")<>"" then salario3=cdbl(request.form("salario3")) else salario3=0
	if request.form("salario4")<>"" then salario4=cdbl(request.form("salario4")) else salario4=0
	salario1=request.form("valorcarta")
	totalsalario=salario1+salario2+salario3+salario4
	sc=5189.82
	inss=570.88
	if totalsalario<sc then
		sc=totalsalario
		inss=int(sc * 11+0.5)/100
		'document.form.tsc.value=formatnumber(sc,2)
		'document.form.tinss.value=formatnumber(inss,2)
	else
		'document.form.tsc.value=formatnumber(sc,2)
		'document.form.tinss.value=formatnumber(inss,2)
	end if
	'response.write "<Br>1. " & salario1
	'response.write "<Br>2. " & sc
	'response.write "<Br>3. " & totalsalario
	sc1=salario1 * sc/totalsalario
	sc1=int(sc1*100+0.05)/100
	inss1=int(sc1 * 11+0.5)/100
	'document.form.sc1.value=formatnumber(sc1,2)
	'document.form.inss1.value=formatnumber(inss1,2)
	teste1=sc1:teste2=inss1
	if (salario2+salario3+salario4)>0 then
		prs1=formatnumber(sc1,2)
		pri1=formatnumber(inss1,2)
		pts=formatnumber(totalsalario,2)
		psc=formatnumber(sc,2)
		pinss=formatnumber(inss,2)
	else 
		prs1=""
		pri1=""
		pts=""
		psc=formatnumber(vsc,2)
		pinss=formatnumber(vinss,2)
	end if
	if salario2<>"" and salario2>0 then 
		sc2=(salario2 * sc) / totalsalario
		sc2=int(sc2*100+0.05)/100
		inss2=int(sc2 * 11+0.5)/100
		teste1=teste1+sc2:teste2=teste2+inss2
		'document.form.sc2.value=formatnumber(sc2,2)
		'document.form.inss2.value=formatnumber(inss2,2)
		if sc2>0 then prs2=formatnumber(sc2,2) else prs2=""
		if inss2>0 then pri2=formatnumber(inss2,2) else pri2=""
	else
		'document.form.sc2.value=""
		'document.form.inss2.value=""
	end if
	if salario3<>"" and salario3>0 then 
		sc3=(salario3 * sc) / totalsalario
		sc3=int(sc3*100+0.05)/100
		inss3=int(sc3 * 11+0.5)/100
		teste1=teste1+sc3:teste2=teste2+inss3
		'document.form.sc3.value=formatnumber(sc3,2)
		'document.form.inss3.value=formatnumber(inss3,2)
		if sc3>0 then prs3=formatnumber(sc3,2) else prs3=""
		if inss3>0 then pri3=formatnumber(inss3,2) else pri3=""
	else
		'document.form.sc3.value=""
		'document.form.inss3.value=""
	end if
	if salario4<>"" and salario4>0 then 
		sc4=(salario4 * sc) / totalsalario
		sc4=int(sc4*100+0.05)/100
		inss4=int(sc4 * 11+0.5)/100
		teste1=teste1+sc4:teste2=teste2+inss4
		'document.form.sc4.value=formatnumber(sc4,2)
		'document.form.inss4.value=formatnumber(inss4,2)
		if sc4>0 then prs4=formatnumber(sc4,2) else prs4=""
		if inss4>0 then pri4=formatnumber(inss4,2) else pri4=""
	else
		'document.form.sc4.value=""
		'document.form.inss4.value=""
	end if
	if teste1-sc<>0 then acerto1=teste1-sc:sc1=sc1-acerto1 'document.form.sc1.value=formatnumber(sc1-acerto1,2)
	if teste2-inss<>0 then acerto2=teste2-inss:inss1=inss1-acerto2 'document.form.inss1.value=formatnumber(inss1-acerto2,2)
	'document.form.empresa4.value=formatnumber(acerto2,2) & " - " & teste2
end if

%>
<form method="POST" name="form">
<input type="hidden" name="valorcarta" size="15" value="<%=request.form("valorcarta")%>">
<input type="hidden" name="mescarta" size="15" value="<%=request.form("mescarta")%>">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650>
<tr>
	<td class="campop" align="center">
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<font size=3><b>Declaração de Salários para Proporcionalidade da Contribuição do INSS</b></font>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	</td>
</tr>
<tr>
	<td class="campop"><p align=justify>Eu, <%=session("nomecarta")%>, portador da CTPS nº <%=ctps%> 
	série <%=serie%>, abaixo-assinado, declaro para fins de desconto proporcional 
da contribuição ao INSS, que meus salários referentes ao mês de <%=request.form("mescarta")%>, são 
os abaixo relacionados, comprometendo-me a comunicar em tempo hábil qualquer alteração nos salários 
e/ou vínculo de empregado regido pela C.L.T., cabendo única e exclusivamente a mim a responsabilidade 
pelas diferenças, multas e penalidades, etc., decorrente da omissão de informações ou informações 
incorretas.
	</td>
</tr>
<tr>
	<td class="campop" align="center">
	<p>&nbsp;</p>
	<font size=3><b>Demonstrativo de Salários e Contribuição</b></font>
	</td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650>
<tr>
	<td class="campop" align="center">#</td>
	<td class="campop" align="center">Instituição/Empresa</td>
	<td class="campop" align="center">Salário Base</td>
	<td class="campop" align="center">Salário Contribuição</td>
	<td class="campop" align="center">Contr. INSS</td>
</tr>
<tr>
	<td class="campop" height=30 align="center">1</td>
	<td class="campop" >Fundação Instituto de Ensino para Osasco</td>
	<input type="hidden" name="salario1" value="<%=request.form("salario1")%>">
	<td class="campop" align="right" width=110><%=formatnumber(request.form("valorcarta"),2)%>&nbsp;&nbsp;&nbsp;</td>
	<td class="campop" align="center" width=110><input type="text" name="sc1" value="<%=prs1%>" size=9 class=proporcional></td>
	<td class="campop" align="center" width=110><input type="text" name="inss1" value="<%=pri1%>" size=9 class=proporcional></td>
</tr>
<tr>
	<td class="campop" height=30 align="center">2</td>
	<td class="campop" ><input type="text" name="empresa2" value="<%=request.form("empresa2")%>" size="40" class="form_input10"></td>
	<td class="campop" align="center" width=110><input type="text" name="salario2" value="<%=request.form("salario2")%>" size=10 class=proporcional onchange="javascript:submit()"></td>
	<td class="campop" align="center"><input type="text" name="sc2" value="<%=prs2%>" size=9 class=proporcional></td>
	<td class="campop" align="center" width=110><input type="text" name="inss2" value="<%=pri2%>" size=9 class=proporcional></td>
</tr>
<tr>
	<td class="campop" height=30 align="center">3</td>
	<td class="campop" ><input type="text" name="empresa3" value="<%=request.form("empresa3")%>" size=40 class=form_input10></td>
	<td class="campop" align="center" width=110><input type="text" name="salario3" value="<%=request.form("salario3")%>" size=10 class=proporcional onchange="javascript:submit()"></td>
	<td class="campop" align="center"><input type="text" name="sc3" value="<%=prs3%>" size=9 class=proporcional></td>
	<td class="campop" align="center" width=110><input type="text" name="inss3" value="<%=pri3%>" size=9 class=proporcional></td>
</tr>
<tr>
	<td class="campop" height=30 align="center">4</td>
	<td class="campop" ><input type="text" name="empresa4" value="<%=request.form("empresa4")%>" size=40 class=form_input10></td>
	<td class="campop" align="center" width=110><input type="text" name="salario4" value="<%=request.form("salario4")%>" size=10 class=proporcional onchange="javascript:submit()"></td>
	<td class="campop" align="center"><input type="text" name="sc4" value="<%=prs4%>" size=9 class=proporcional></td>
	<td class="campop" align="center" width=110><input type="text" name="inss4" value="<%=pri4%>" size=9 class=proporcional></td>
</tr>
<tr>
	<td class="campop" height=30 align="center" colspan=2>Totais</td>
	<td class="campop" align="center" width=110><input type="text" name="salariototal" value="<%=pts%>" size=10 class=proporcional></td>
	<td class="campop" align="center" valign="middle">          <input type="text" name="tsc" value="<%=psc%>" size=9 class=proporcional>&nbsp;</td>
	<td class="campop" align="center" width=110><input type="text" name="tinss" value="<%=pinss%>" size=9 class=proporcional>&nbsp;</td>
</tr>
</table>
<hr>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650>
<tr>
	<td class="campop" align="center">#</td>
	<td class="campop" align="center">Carimbo e Assinatura</td>
	<td class="campop" align="center">Matrícula no INSS</td>
	<td class="campop" align="center">Data</td>
</tr>
<tr>
	<td class="campop" height=30 align="center">1</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center" ><input type="text" name="cnpj1" value="73.063.166/0003-92" size=18 class=form_input10></td>
	<td class="campop" align="center">&nbsp;<%=formatdatetime(now(),2)%></td>
</tr>
<tr>
	<td class="campop" height=30 align="center">2</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center"><input type="text" name="cnpj2" value="<%=request.form("cnpj2")%>" size=18 class=form_input10></td>
	<td class="campop" align="center"><input type="text" name="data2" value="<%=request.form("data2")%>" size=10 class=proporcional></td>
</tr>
<tr>
	<td class="campop" height=30 align="center">3</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center"><input type="text" name="cnpj3" value="<%=request.form("cnpj3")%>" size=18 class=form_input10></td>
	<td class="campop" align="center"><input type="text" name="data3" value="<%=request.form("data3")%>" size=10 class=proporcional></td>
</tr>
<tr>
	<td class="campop" height=30 align="center">4</td>
	<td class="campop" align="center">&nbsp;</td>
	<td class="campop" align="center"><input type="text" name="cnpj4" value="<%=request.form("cnpj4")%>" size=18 class=form_input10></td>
	<td class="campop" align="center"><input type="text" name="data4" value="<%=request.form("data4")%>" size=10 class=proporcional></td>
</tr>
</table>
<p>Data:______/_____/_______
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>_______________________________________________</p>
<p><%=request.form("nomecarta")%></p>
</form>
<%
'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>