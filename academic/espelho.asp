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
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
if request.form("chapa")<>"" then 
	chapa=request.form("chapa") 
else 
	chapa=session("usuariomaster")
end if

%>
<!-- -->
<!-- -->
<form method="POST" action="espelho.asp" name="form">
<%
if session("acesso")=2 or session("usuariogrupo")="COORD.CURSO" or session("usuariogrupo")="JURIDICO"  then
%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" >
<tr><td valign=top style="border-right:3px double silver;border-bottom:3px double silver" width=150 height=600>
<!-- -->
<p style="margin-top:0;margin-bottom:0" class=titulo><%=session("usuarioname")%></p>
<hr>
<p style="margin-top:0;margin-bottom:5"><a href="../academic/disponibilidade.asp">
<img src="../images/Clock.gif" width="16" height="16" border="0" alt="">Disponibilidade</a></p>

<p style="margin-top:0;margin-bottom:5"><a href="../academic/aderencia.asp">
<img src="../images/BookO.gif" width="16" height="16" border="0" alt="">Aderência</a></p>

<br><br>
<p style="margin-top:0;margin-bottom:5"><a href="../academic/meusplanos.asp">
<img src="../images/BookO.gif" width="16" height="16" border="0" alt="">Plano de Ensino</a>

<br><br>
<p style="margin-top:0;margin-bottom:5">
<img src="../images/espelho.jpg" width="16" height="16" border="0" alt="">Marcação de Ponto</p>

<br><br><br><br><br><br><br>
<p style="margin-top:0;margin-bottom:0"><a href="../indexp.asp">
<img src="../images/setafirst0.gif" width="12" height="12" border="0" alt="">Início</a>
<!-- -->
</td><td valign=top style="border-bottom:3px double silver">
<p style="margin-top:0;margin-bottom:10" class=titulo>Espelho de Marcações no ponto eletrônico</p>
<!-- -->
<%
else ' para acesso=1
%>
<p style="margin-top:0;margin-bottom:10" class=titulo>Disponibilidade de Horários</p>
<select size=1 name="chapa" onchange="javascript:submit();">
	<option value="0">Selecione....</option>
<%
sqlc="select f.chapa, nome from grades_aux_prof f where codsituacao<>'D' and f.chapa<'10000' order by nome"
rs.Open sqlc, ,adOpenStatic, adLockReadOnly
do while not rs.eof
if rs("chapa")=chapa then txt="selected" else txt=""
'if rs("disponivel")>0 then estilo="style='background:CCFFCC;'" else estilo="" 'estilo="style='background:FFFFFF;'"
%>
	<option <%=estilo%> value="<%=rs("chapa")%>" <%=txt%>><%=rs("nome")%></option>
<%
rs.movenext
loop
rs.close
%>
</select>
<%
end if 'acesso
%>

<%
'chapa="01032"
sqlt="select distinct s.chapa, jornada=s.jornada, tipo=case when codevento='RHT' then 'RHT' when codevento between '255' and '258' then 'RT' " & _
"when codevento='128' then 'RT' when codevento='138' then 'RT' else 'N' end " & _
", ajuste=case when CODSECAO in ('03.3.600') then '1' else '0' end " & _
"from corporerm.dbo.pfsalcmp s inner join corporerm.dbo.pfunc f on f.chapa=s.chapa where s.chapa='" & chapa & "' "
rs.Open sqlt, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	tipo=rs("tipo")
	ajuste=rs("ajuste")
	jornada=rs("jornada")/60
else
	tipo="N"
	ajuste=rs("ajuste")
end if
rs.close

hoje=now():m2=month(hoje):a2=year(hoje)
'hoje=dateserial(2011,2,16)

if tipo="RT" then
	if day(hoje)>15 then mplus=0 else mplus=-1
	ini2=dateserial(year(hoje),month(hoje)+mplus,16):'response.write ini2
	fim2=dateserial(year(ini2),month(ini2)+1,15)  :'response.write fim2
	ini1=dateserial(year(ini2),month(ini2)-1,16)  :'response.write ini1
	fim1=dateserial(year(fim2),month(fim2)-1,15)  :'response.write fim1
elseif tipo="RHT" then
	ini2=dateserial(year(hoje),month(hoje),1)    :'response.write ini2 & " "
	fim2=dateserial(year(ini2),month(ini2)+1,1)-1:'response.write fim2 & " "
	ini1=dateserial(year(ini2),month(ini2)-1,1)  :'response.write ini1 & " "
	fim1=dateserial(year(fim2),month(fim2),1)-1  :'response.write fim1 & " "
	if fim2>now() then fim2=now()-1
else
	ini2=dateserial(year(hoje),month(hoje),1)    :'response.write ini2 & " "
	fim2=dateserial(year(ini2),month(ini2)+1,1)-1:'response.write fim2 & " "
	ini1=dateserial(year(ini2),month(ini2)-1,1)  :'response.write ini1 & " "
	fim1=dateserial(year(fim2),month(fim2),1)-1  :'response.write fim1 & " "
	if fim2>now() then fim2=now()-1
end if

if request.form<>"" or chapa<>"" or session("acesso")=2 or session("usuariogrupo")="COORD.CURSO" then
	chapa=chapa
	sqld="select f.nome, f.codsindicato, c.nome as funcao, s.descricao as setor from corporerm.dbo.pfunc f, corporerm.dbo.psecao s, corporerm.dbo.pfuncao c where f.codsecao=s.codigo and f.codfuncao=c.codigo and f.chapa='" & chapa & "'"
	rs.Open sqld, ,adOpenStatic, adLockReadOnly
	sindicato=rs("codsindicato")
	if sindicato="03" then coluna=7 else coluna=5
%>
<%linha=linha+1%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class=campo>Espelho de Marcação de Ponto</td></tr></table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class="campor">Chapa</td>
	<td class="campor">Nome</td>
	<td class="campor">Setor</td>
	<td class="campor">Função</td></tr>
<tr><td class="campor"><%=chapa%></td>
	<td class="campor"><b><%=rs("nome")%></b></td>
	<td class="campor"><%=rs("Setor")%></td>
	<td class="campor"><%=rs("funcao")%></td></tr>
</table>

<%
	rs.close
temporario=0
if temporario<>0 then
%>

<!-- quadro 2 meses -->
<table border="0" style="border-collapse: collapse">
<tr><td valign=top>
<%
for z=1 to 2
	if z=1 then data1=ini1:data2=fim1
	if z=2 then data1=ini2:data2=fim2
	if z=2 then response.write "</td><td valign=top>"
%>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr><td class=campo>Período: de <b><%=data1%> até <%=data2%></td></tr></table>

<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" >
<tr>
	<td class=grupo align="center" colspan=2 style="border-right:2px solid #000000">Datas</td>
	<td class=grupo align="center" colspan=7 style="border-right:2px solid #000000">Marcações efetuadas</td>
</tr>
<tr>
	<td class=titulo align="center">Data</td>
	<td class=titulo align="center" style="border-right:2px solid #000000">Dia</td>
	<td class=titulo align="center" width=35>1</td>
	<td class=titulo align="center" width=35>2</td>
	<td class=titulo align="center" width=35>3</td>
	<td class=titulo align="center" width=35>4</td>
	<td class=titulo align="center" width=35>5</td>
	<td class=titulo align="center" width=35>6</td>
	<td class=fundo align="center" style="border-right:2px solid #000000">H.Trab.</td>
</tr>
<%
diasloop=datediff("d",data1,data2)+1:'response.write diasloop & "<br>"
diasloop=cint(diasloop)
totalchcumprir=0
totalchcumprida=0
tcumprida1=0

Redim marc(diasloop,6), formato(diasloop,6)

for e=data1 to data2
	idmatriz=e-(data1)
	'marcações do chronus
	sqlcr="select chapa, day(data) as dia, data, batida, natureza, status from corporerm.dbo.abatfun where chapa='" & chapa & "' " & _
	"and data='" & dtaccess(e) & "' "
	rs2.Open sqlcr, ,adOpenStatic, adLockReadOnly
	marcacao=1
	if rs2.recordcount>0 then
		rs2.movefirst:do while not rs2.eof
		'------------------------------------------
		dia=rs2("dia"):data=rs2("data")
		natureza=rs2("natureza")
		batida=formatdatetime((rs2("batida")/60)/24,4)
		if dia=diaant then marcacao=marcacao+1 else marcacao=1
		'nat(dia,marcacao)=rs2("natureza")
		if natureza=0 or natureza=4 then natu=0
		if natureza=1 or natureza=5 then natu=1
		resto=marcacao mod 2
		if resto=0 and natu=0 then marcacao=marcacao+1 else marcacao=marcacao
		if resto<>0 and natu=1 then marcacao=marcacao+1 else marcacao=marcacao
		marc(idmatriz,marcacao)=batida:'response.write ">> " & idmatriz & " >> " & marc(idmatriz,marcacao) & "<br>"
		if rs2("status")="D" then formato(idmatriz,marcacao)="<font color='red'>" 'else formato(dia,marcacao)="<font color='black'>"
		diaant=dia
		'------------------------------------------
		rs2.movenext:loop
	else 'recordcount rs2
		for b=1 to 6:marc(idmatriz,b)="":next
	end if 'recordcount rs2
	rs2.close
next

dtponto=data1

for e=data1 to data2
	'dtponto=dateserial(ano,mes,a)
	dtponto=e
	idmatriz=e-(data1)
	if idmatriz=0 then indice=indice else indice=indice+1
	response.write "<tr>"

	response.write "<td class=campo align="center">" & dtponto & "</td>"
	response.write "<td class=campo align="center" style='border-right:2px solid #000000'>" & weekdayname(weekday(dtponto),-1) & "</td>"
	
	'*************ocorrencias
	sql1="select base, htrab from corporerm.dbo.aafhtfun where data='" & dtaccess(dtponto) & "' and chapa='" & chapa & "' "

	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 
		base=rs("base")
		htrab=rs("htrab")
	else
		base=0:htrab=0
	end if
	rs.close
	tbase=tbase+base

	for b=1 to 6
		batida=marc(idmatriz,b)
		if batida<>"" then ultima=b
		response.write "<td class=campo align="center">" & formato(idmatriz,b) & batida & "</font></td>"
	next
	If marc(idmatriz,1)="" and marc(idmatriz,2)="" and marc(idmatriz,3)="" and marc(idmatriz,4)="" and marc(idmatriz,5)="" and marc(idmatriz,6)="" then
		tot1=0:tot2=0:tot3=0
	else
		if marc(idmatriz,2)="" and marc(idmatriz,1)<>"" then tot1=0 else tot1=cdate(marc(idmatriz,2))-cdate(marc(idmatriz,1))
		if marc(idmatriz,4)="" and marc(idmatriz,3)<>"" then tot2=cdate(marc(idmatriz,3))-cdate(marc(idmatriz,2)) else tot2=cdate(marc(idmatriz,4))-cdate(marc(idmatriz,3))
		if marc(idmatriz,6)="" and marc(idmatriz,5)<>"" then tot3=cdate(marc(idmatriz,5))-cdate(marc(idmatriz,4)) else tot3=cdate(marc(idmatriz,6))-cdate(marc(idmatriz,5))
	end if

	thtrab=thtrab+htrab
	totc=tot1+tot2+tot3
	totch=formatdatetime(totc,4)
	if totc=0 then totch="-" else totch=formatdatetime(totc,4)
		
	response.write "<td class=campo align="center" style='border-right:2px solid #000000'>"
	if htrab>0 then response.write formatdatetime((htrab/60)/24,4) 
	if htrab=0 then response.write "<font color=gray>" & totch 
	response.write "</font></td>"

	if htrab>0 then tcumprida1=tcumprida1 + htrab:tcumprida2=tcumprida2 + htrab
	if htrab=0 then tcumprida2=tcumprida2 + (totc*24*60)
		
	response.write "</tr>"
next

if request.form("considerar")="ON" then totalgeral=tcumprida2 else totalgeral=tcumprida1
%>
<tr>
	<td class=titulo align="left" colspan=8>&nbsp;Totais</td>
	<td class=campo align="center" style="border-right:2px solid #000000"><%=formatnumber(totalgeral/60,2)%></td>
</tr>
</table>

<%
next
%>
<!-- fim quadro 2 meses -->
<%
end if 'temporario
%>

</td></tr></table>
<br>Instruções:
<%
if tipo="RHT" then 
	response.write "<br>Caso tenha que entregar algum relatório para a Diretoria, entregue-o até no máximo o dia 15 de cada mês."
	sqla="select aulas=sum(ta) from g2ch where chapa1='" & session("usuariomaster") & "' and '" & dtaccess(fim2) & "' between inicio and termino "
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
	aulas=rs("aulas")
	rs.close
	t1=(40-aulas + (aulas/60*50))*4
	jornada=int(t1+0.5)
end if
if tipo="RT" then 
	if ajuste=1 then
		jornada=int(jornada)
	else
		jornada=int(jornada/2)
	end if
	response.write ""
end if

sqlfer="select feriados=count(diaferiado) from corporerm.dbo.gferiado where diaferiado between '" & dtaccess(ini2) & "' and '" & dtaccess(fim2) & "'"
rs.Open sqlfer, ,adOpenStatic, adLockReadOnly
feriados=rs("feriados")
rs.close
%>
<script language="javascript">
function Calcula(){
var Parametro1=document.form.jornadabasica.value;
var Parametro2=document.form.suspensos.value;
var Soma= (form.jornadabasica.value*1) / 30 * (30-(form.suspensos.value*1)) ; 
var temp=Soma
document.form.jornada2.value=temp.toFixed(2);
}
</script>
<%if tipo<>"N" then%>
<br><input type="hidden" name="jornadabasica" value="<%=jornada%>">
<br>Jornada básica: <%=jornada%>
<br>Dias com suspensão de aulas ou feriados: <input type="text" size="1" name="suspensos" value="0" title="Dias com suspensão de aulas ou feriados" onKeyUp="Calcula()" onKeyPress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;" >
<br>Número de Horas a ser cumprida no ultimo período: <input type="text" size="4" name="jornada2" value="<%=jornada%>" title="Jornada a ser cumprida">
<br>
<%end if%>


<%
end if

set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>