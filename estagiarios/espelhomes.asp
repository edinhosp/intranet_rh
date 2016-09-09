<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")="N" or session("a72")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Espelho de Marcação Eletrônica</title>
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

chapa=request("chapa")
sql="select top 1 * from est_parametro "
rs.Open sql, ,adOpenStatic, adLockReadOnly
ano=rs("ano")
mes=rs("mes")
descricao=rs("descricao")
inicio1=rs("inicio")
fim1=rs("fim")
limite=rs("limite")
rs.close

data1=inicio1
data2=fim1
data1=cdate(data1):data2=cdate(data2)
sqld="select f.nome, f.codsindicato, c.nome as funcao, s.descricao as setor from corporerm.dbo.pfunc f, corporerm.dbo.psecao s, corporerm.dbo.pfuncao c where f.codsecao=s.codigo and f.codfuncao=c.codigo and f.chapa='" & chapa & "'"
'response.write sqld
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
sindicato=rs2("codsindicato")
if sindicato="03" then coluna=7 else coluna=5

%>
<%linha=linha+1%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class=campo style="border-top:5 double">Cartão de Marcações Eletrônicas</td></tr></table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class=campo>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td></tr></table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class=campo style="border-bottom:5 double">Período: de <b><%=data1%> até <%=data2%></td></tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=campo>Chapa</td>
	<td class=campo>Nome</td>
	<td class=campo>Setor</td>
	<td class=campo>Função</td></tr>
<tr><td class=campo><%=chapa%></td>
	<td class=campo><b><%=rs2("nome")%></b></td>
	<td class=campo><%=rs2("Setor")%></td>
	<td class=campo><%=rs2("funcao")%></td></tr>
</table>
<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" >
<tr>
	<td class=campo align="center">Data</td>
	<td class=campo align="center">Dia</td>
	<td class=campo align="center" style="border-right:2px solid #000000">Ind</td>
	<td class=campo align="center" style="border-right:2px solid #000000">Horas<br>Base</td>
	<td class=campo align="center" width=35>Ent1</td>
	<td class=campo align="center" width=35>Sai1</td>
	<td class=campo align="center" width=35>Ent2</td>
	<td class=campo align="center" width=35>Sai2</td>
	<td class=campo align="center" width=35>Ent3</td>
	<td class=campo align="center" width=35>Sai3</td>
	<td class=campo align="center" style="border-right:2px solid #000000">H.Trab.</td>
	<td class=campo align="center" >Extra<br>Aut.</td>
	<td class=campo align="center" >Atraso</td>
	<td class=campo align="center" style="border-right:2px solid #000000">Falta</td>
	<td class=campo align="center" style="border-right:2px solid #000000">Horas<br>Pagar</td>
</tr>
<tr><td class="campor" colspan=20 style="border-bottom:5 double"></td></td>
<%
nome=rs2("nome")
rs2.close
diasloop=datediff("d",data1,data2)+1:'response.write diasloop & "<br>"
diasloop=cint(diasloop)
totalchcumprir=0
totalchcumprida=0

Redim marc(diasloop,6), formato(diasloop,6)

for e=data1 to data2
	idmatriz=e-(data1)
	'marcações do chronus
	sqlcr="select chapa, day(data) as dia, data, batida, natureza, status from abatfunam where chapa='" & chapa & "' " & _
	"and data='" & dtaccess(e) & "' " & _
	"UNION ALL " & _
	"select chapa, day(data) as dia, data, batida, natureza, status from abatfun where chapa='" & chapa & "' " & _
	"and data='" & dtaccess(e) & "' "
	sqlcr="select chapa, day(data) as dian, data, dia, marc1, marc2, marc3, marc4, marc5, marc6, htrab, base, codigo, " & _
	"ajust1, ajust2, ajust3, ajust4, ajust5, ajust6, atraso, falta, extra, extraaut, descanso, feriado, dataaprov " & _
	"from est_batfun where chapa='" & chapa & "' and data='" & dtaccess(e) & "' "
	rs2.Open sqlcr, ,adOpenStatic, adLockReadOnly
	marcacao=1:extra=0:autor=0
	'------------------------------------------
	dia=rs2("dian"):data=rs2("data"):indice=rs2("dia")
	marc1=horaload(rs2("marc1"),2):if rs2("ajust1")>0 then marc1=horaload(rs2("ajust1"),2)
	marc2=horaload(rs2("marc2"),2):if rs2("ajust2")>0 then marc2=horaload(rs2("ajust2"),2)
	marc3=horaload(rs2("marc3"),2):if rs2("ajust3")>0 then marc3=horaload(rs2("ajust3"),2)
	marc4=horaload(rs2("marc4"),2):if rs2("ajust4")>0 then marc4=horaload(rs2("ajust4"),2)
	marc5=horaload(rs2("marc5"),2):if rs2("ajust5")>0 then marc5=horaload(rs2("ajust5"),2)
	marc6=horaload(rs2("marc6"),2):if rs2("ajust6")>0 then marc6=horaload(rs2("ajust6"),2)
	'------------------------------------------
	extra=rs2("extra")
	autor=rs2("extraaut")
	pontoextra=""
	if isnull(autor) and extra>0 then pontoextra="." 
	if autor>0 and extra>0 then pontoextra=horaload(autor,2)
	if autor>0 then extra1=autor else extra1=0
	htrab=rs2("htrab")
	base=rs2("base")
	checagem=htrab-base
	horas=base
	if abs(checagem)<=9 then horas=base else horas=htrab
	if autor>0 then horas=htrab
	thoras=thoras+horas
	tatraso=tatraso+rs2("atraso")
	textra=textra+extra1
	thtrab=thtrab+htrab
%>
<tr>
	<td class=campo align="center"><%=e%></td>
	<td class=campo align="center"><%=weekdayname(weekday(e),-1)%></td>
	<td class=campo align="center" style="border-right:2px solid #000000"><%=indice%></td>
	<td class=campo align="center" style="border-right:2px solid #000000"><%=horaload(rs2("base"),2)%></td>
<%if rs2("feriado")>0 or rs2("descanso")>0 then
	if rs2("feriado")>0 then texto="<b>FERIADO"
	if rs2("descanso")>0 then texto="<b>DESCANSO"
%>	
	<td class=campo align="left" colspan=6><%=texto%> </td>
<%else%>
	<td class=campo align="center" width=35><%=marc1%></td>
	<td class=campo align="center" width=35><%=marc2%></td>
	<td class=campo align="center" width=35><%=marc3%></td>
	<td class=campo align="center" width=35><%=marc4%></td>
	<td class=campo align="center" width=35><%=marc5%></td>
	<td class=campo align="center" width=35><%=marc6%></td>
<%end if%>
	<td class=campo align="center" style="border-right:2px solid #000000"><%=horaload(rs2("htrab"),2)%></td>
	<td class=campo align="center" style="border-right:1px dashed #000000"><%=pontoextra%></td>
	<td class=campo align="center" style="border-right:1px dashed #000000"><%=horaload(rs2("atraso"),2)%></td>
	<td class=campo align="center" style="border-right:2px solid #000000"><%=horaload(rs2("falta"),2)%></td>
	<td nowrap class=campo align="center" style="border-right:2px solid #000000"><%=horaload(horas,2)%></td>
</tr>
<%
	codigo=rs2("codigo")
	rs2.close
next 'llllllllll
sql="select descricao from est_Cadhorario where codigo='" & codigo & "'"
rs.Open sql, ,adOpenStatic, adLockReadOnly
descricao=rs("descricao")
rs.close

%>
<tr><td class="campor" colspan=20 style="border-bottom:5 double"></td></td>
<tr>
	<td class=campo colspan=9></td>
	<td class=campo colspan=6><br>
<pre>
Total de atrasos          : <%=horaload(tatraso,2)%>
Total de extras           : <%=horaload(textra,2)%>
Total de horas a pagar    : <%=horaload(thoras,2)%>
</pre>
	</td>
</tr>
<tr>
	<td class=campo colspan=15>
<pre>
Horário do Estágio:
<%=descricao%>
</pre>
	</td>
</tr>
<tr>
	<td class=campo colspan=15 height=100 valign="bottom">
<pre>
Reconheço a exatidão e confirmo a frequência constante deste cartão


                                        _____________________________________________
                                        <%=nome%>
</pre>
	</td>
</tr>
</table>
<%
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>