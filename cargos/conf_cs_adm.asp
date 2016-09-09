<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a77")="N" or session("a77")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Checagem - Cargos e Salários</title>
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
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
set rss=server.createobject ("ADODB.Recordset")
Set rss.ActiveConnection = conexao
%>
<%
if request.form="" then
data_1=dateserial(year(now),month(now),1)
data_2=dateserial(year(now),month(now)+1,1)-1
%>
<p class=titulo>Conferência de Cargos e Salários</p>
<form method="POST" action="conf_cs_adm.asp">

<p style="margin-top:0;margin-bottom:0">Data da tabela: <select size="1" name="datatabela">
<%
sqla="SELECT data FROM cs_obs order by data desc"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
%>
<option value="<%=rs("data")%>"><%=rs("data")%></option>
<%
rs.movenext:loop
rs.close
%>  
</select><br>

<p style="margin-top:0;margin-bottom:0">Setor: <select size="1" name="setor">
<option value="0">Todos Setores</option>
<%
sqla="select distinct CODSECAO, secao from qry_funcionarios where CODSITUACAO<>'D' and CODTIPO='N' and CODSINDICATO<>'03' order by secao"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
%>
<option value="<%=rs("codsecao")%>"><%=rs("secao")%></option>
<%
rs.movenext:loop
rs.close
%>  
</select><br>

<p style="margin-top:0;margin-bottom:0"><input type="checkbox" name="histmes" value="ON">Somente alterações do mês 
<input type="text" name="meshst" size="2" maxlength="2" class="form_apt" value="<%=month(now())%>">
/<input type="text" name="anohst" size="4" maxlength="4" class="form_apt" value="<%=year(now())%>">
( Não pular página <input type="checkbox" name="quebrapag" value="ON">)
<br>
	
<input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>
<%
end if

if request.form<>"" then
data=request.form("datatabela")

if request.form("setor")<>"0" then sqlsetor=" and codsecao in ('" & request.form("setor") & "') " else sqlsetor=""

if request.form("histmes")="ON" then
	mes=request.form("meshst")
	ano=request.form("anohst")
	sqlchapa="and chapa in (select chapa from (SELECT CHAPA FROM corporerm.dbo.PFHSTSAL WHERE Month([DTMUDANCA])=" & mes & " AND Year([DTMUDANCA])=" & ano & " " & _
	"union all SELECT CHAPA FROM corporerm.dbo.PFHSTFCO WHERE Month([DTMUDANCA])=" & mes & " AND Year([DTMUDANCA])=" & ano & ") as t group by chapa) " & sqlsetor
else
	sqlchapa=""
end if

sql0="SELECT f.CHAPA, f.NOME, f.CODFUNCAO, lab.codfuncao AS codfuncaocs, lab.id_setor, lab.id_cargo, f.CODTIPO, c.cargo " & _
"FROM (cs_cargos_lab lab RIGHT JOIN corporerm.dbo.PFUNC f ON lab.codfuncao=f.CODFUNCAO collate database_default) LEFT JOIN cs_cargos c ON (lab.id_cargo=c.id_cargo) AND (lab.id_setor=c.id_setor) " & _
"WHERE lab.codfuncao Is Null AND f.CODTIPO='N' AND f.CODSITUACAO<>'D' AND f.CODSINDICATO<>'03'; "
rs2.Open sql0, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
	response.write "<font size='4'>Cargos não classificados</font><br>"
	do while not rs2.eof
	response.write rs2("chapa") & "-" & rs2("nome") & "-" & rs2("codfuncao") 
	response.write "<br>"
	rs2.movenext
	loop
	response.write "<hr>"
	response.write "<DIV style=""page-break-after:always""></DIV>"
end if
rs2.close
	
sql2="SELECT cs_salarios.id_setor, cs_areas.setor, cs_salarios.data " & _
"FROM cs_salarios INNER JOIN cs_areas ON cs_salarios.id_setor = cs_areas.id_setor " & _
"GROUP BY cs_salarios.id_setor, cs_areas.setor, cs_salarios.data " & _
"HAVING cs_salarios.data='" & dtaccess(data) & "' " & _
"ORDER BY cs_areas.setor "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
%>
<!--tirei titulo daqui -->
<%
titulo="<p style=""margin-top:0;margin-bottom:0;color:Blue;font-size:9pt;text-align:left""><b>Análise de Plano de Carreira Administrativo - " & rs2("setor") & "<br>" & monthname(month(data)) & "/" & year(data) & "<br>&nbsp;</font></p>"
response.write titulo
%>
<!-- inicio quadro do faixas do setor -->
<%
sql="SELECT a.id_setor, a.setor, c.ordem, c.cargo, s.data, s.n1, s.n2, s.n3, s.n4, s.n5, c.horas, s.id_cargo " & _
"FROM (cs_cargos AS c INNER JOIN cs_areas AS a ON c.id_setor = a.id_setor) " & _
"INNER JOIN cs_salarios AS s ON (c.id_cargo = s.id_cargo) AND (c.id_setor = s.id_setor) " & _
"WHERE s.data='" & dtaccess(data) & "' and a.id_setor='" & rs2("id_setor") & "' " & _
"ORDER BY a.setor, c.ordem "
rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<%
rs.movefirst
do while not rs.eof
sql3="SELECT f.CODSECAO, s.DESCRICAO, f.CHAPA, f.NOME, f.DATAADMISSAO, f.CODFUNCAO, c.NOME AS FUNCAO, f.GRUPOSALARIAL, f.SALARIO, [JORNADAMENSAL]/60 AS JORNADA, f.codnivelsal " & _
"FROM (cs_cargos_lab lab INNER JOIN (corporerm.dbo.PFUNC f INNER JOIN corporerm.dbo.PFUNCAO c ON f.CODFUNCAO=c.CODIGO) ON lab.codfuncao=f.CODFUNCAO collate database_default) INNER JOIN corporerm.dbo.PSECAO s ON f.CODSECAO=s.CODIGO " & _
"WHERE f.CODSINDICATO<>'03' AND f.CODTIPO='N' AND f.CODSITUACAO<>'D' " & _
"AND lab.id_setor='" & rs("id_setor") & "' AND lab.id_cargo='" & rs("id_cargo") & "' " & _
"" & sqlchapa & "" & sqlsetor & _
"ORDER BY f.NOME "
'response.write "<br>" & sql3
vazio=0
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
vazio=1
%>

<!--pus titulo aqui -->

<table border='1' bordercolor='#000000' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width=950>
<tr><td class=titulor colspan=5>&nbsp;Cargo</td>
	<td class=titulor align="center">Horas</td>
	<td class=titulor align="center" rowspan=2 colspan=2>&nbsp;</td>
	<td class=titulor align="center">N1</td>
	<td class=titulor align="center">N2</td>
	<td class=titulor align="center">N3</td>
	<td class=titulor align="center">N4</td>
	<td class=titulor align="center">N5</td>
	<td class=titulor align="center" rowspan=2 colspan=2>&nbsp;</td>
</tr>
<tr><td class="campoa" colspan=5>&nbsp;<%=rs("cargo")%></td>
	<td class="campoa" align="right"><%=rs("horas")%>&nbsp;</td>
	<td class="campoa"r align="right"><%=formatnumber(rs("n1"),2)%>&nbsp;</td>
	<td class="campoa"r align="right"><%=formatnumber(rs("n2"),2)%>&nbsp;</td>
	<td class="campoa"r align="right"><%=formatnumber(rs("n3"),2)%>&nbsp;</td>
	<td class="campoa"r align="right"><%=formatnumber(rs("n4"),2)%>&nbsp;</td>
	<td class="campoa"r align="right"><%=formatnumber(rs("n5"),2)%>&nbsp;</td>
</tr>
<!-- inicio funcionarios -->
<%
end if
rs3.close
sql3="SELECT f.CODSECAO, s.DESCRICAO, f.CHAPA, f.NOME, f.DATAADMISSAO, f.CODFUNCAO, c.NOME AS FUNCAO, f.GRUPOSALARIAL, f.SALARIO, [JORNADAMENSAL]/60 AS JORNADA, f.codnivelsal " & _
"FROM (cs_cargos_lab lab INNER JOIN (corporerm.dbo.PFUNC f INNER JOIN corporerm.dbo.PFUNCAO c ON f.CODFUNCAO=c.CODIGO) ON lab.codfuncao=f.CODFUNCAO collate database_default) INNER JOIN corporerm.dbo.PSECAO s ON f.CODSECAO=s.CODIGO " & _
"WHERE f.CODSINDICATO<>'03' AND f.CODTIPO='N' AND f.CODSITUACAO<>'D' " & _
"AND lab.id_setor='" & rs("id_setor") & "' AND lab.id_cargo='" & rs("id_cargo") & "' " & _
"" & sqlchapa & "" & sqlsetor & _
"ORDER BY f.codsecao, f.NOME "
'"AND PFUNC.DATAADMISSAO<=#12/31/2003# " & _
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
vazio=1
%>
<tr>
	<td class=titulor align="center">Função</td>
	<td class=titulor align="center">Seção</td>
	<td class=titulor align="center">Chapa</td>
	<td class=titulor align="center">Funcionário</td>
	<td class=titulor align="center">Adm.</td>
	<td class=titulor align="center">Jorn.</td>
	<td class=titulor align="center">Faixa</td>
	<td class=titulor align="center">abaixo</td>
	<td class=titulor align="center" colspan=5>Enquadramento</td>
	<td class=titulor align="center">acima</td>
	<td class=titulor align="center">Status</td>
</tr>
<%
rs3.movefirst
do while not rs3.eof
admissao=cdate(rs3("dataadmissao"))
jornadafu=cdbl(rs3("jornada"))
jornadacs=cdbl(rs("horas"))
'sqls="select top 1 salario from pfhstsal where chapa='" & rs3("chapa") & "' and dtmudanca<=#12/31/2003# order by dtmudanca desc"'
'rss.Open sqls, ,adOpenStatic, adLockReadOnly

salajustado=cdbl(rs3("salario")):ajustado=0
'salajustado=cdbl(rss("salario")):ajustado=0

formato="<font color='#000000'>"
if cdate(now)-admissao<90 then 
	if rs3("codnivelsal")=0 then
		formato="<font color='#0000ff'>"
		salajustado=(int(cdbl(rs3("salario"))/0.0095)/100)
		'salajustado=(int(cdbl(rss("salario"))/0.0095)/100)
		ajustado=1
	else
		salajustado=rs3("salario")
		'salajustado=rss("salario")
		ajustado=0
	end if
end if
if jornadafu<>jornadacs then
	formato="<font color='#ff0000'>"
	salajustado=cdbl(salajustado)
	salajustado=int(((salajustado/jornadafu)*jornadacs)*100)/100
	ajustado=1
end if
status=""
if rs3("gruposalarial")="N1" then salfaixa=rs("n1")
if rs3("gruposalarial")="N2" then salfaixa=rs("n2")
if rs3("gruposalarial")="N3" then salfaixa=rs("n3")
if rs3("gruposalarial")="N4" then salfaixa=rs("n4")
if rs3("gruposalarial")="N5" then salfaixa=rs("n5")
if rs3("gruposalarial")="ACN5" then salfaixa=0
if rs3("gruposalarial")="" or isnull(rs3("gruposalarial")) then status=" -Grupo Salarial não cadastrado"

arredondamento=cdbl(salajustado)-cdbl(salfaixa)
if arredondamento<=1 and arredondamento>=-1 then
	arredondamento=cdbl(salajustado)-cdbl(salfaixa)
	if arredondamento<-0.01 then status=status & " -Arrend.a menor: " & formatnumber(arredondamento,2)
	if arredondamento>0.01 then status=status & " -Arrend.a maior: " & formatnumber(arredondamento,2)
	if jornadafu<>jornadacs then
		salcorreto=int(((salfaixa/jornadacs)*jornadafu)*100)/100
		if salcorreto-salajustado>abs(0.01) then status=status & " - " & formatnumber(salcorreto,2)
	end if
end if
if arredondamento>1 or arredondamento<-1 then
	salajustado=cdbl(salajustado)
	tfaixa1=salajustado-rs("n1")
	tfaixa2=salajustado-rs("n2")
	tfaixa3=salajustado-rs("n3")
	tfaixa4=salajustado-rs("n4")
	tfaixa5=salajustado-rs("n5")
	if tfaixa1<=1 and tfaixa1>=-1 then 
		status=status & " -Grupo Correto: N1"
	elseif tfaixa2<=1 and tfaixa2>=-1 then
		status=status & " -Grupo Correto: N2"
	elseif tfaixa3<=1 and tfaixa3>=-1 then 
		status=status & " -Grupo Correto: N3"
	elseif tfaixa4<=1 and tfaixa4>=-1 then 
		status=status & " -Grupo Correto: N4"
	elseif tfaixa5<=1 and tfaixa5>=-1 then 
		status=status & " -Grupo Correto: N5"
	else
		status=status & " -Não checado"
	end if
end if
if rs3("gruposalarial")="ACN5" then status=status & " -Casos não classificados"
quadro1="&nbsp;"
quadro2="&nbsp;"
quadro3="&nbsp;"
quadro4="&nbsp;"
quadro5="&nbsp;"
quadro0="&nbsp;"
quadro9="&nbsp;"

if ajustado=1 then msgajuste="<br>(" & formato & formatnumber(salajustado,2) & ")</font>" else msgajuste=""
if cint(salajustado)>=cint(rs("n1")) and cint(salajustado)<cint(rs("n2")) then
	quadro1=formatnumber(rs3("salario"),2) & msgajuste
	'quadro1=formatnumber(rss("salario"),2) & msgajuste
elseif cint(salajustado)>=cint(rs("n2")) and cint(salajustado)<cint(rs("n3")) then
	quadro2=formatnumber(rs3("salario"),2) & msgajuste
	'quadro2=formatnumber(rss("salario"),2) & msgajuste
elseif cint(salajustado)>=cint(rs("n3")) and int(salajustado)<int(rs("n4")) then
	quadro3=formatnumber(rs3("salario"),2) & msgajuste
	'quadro3=formatnumber(rss("salario"),2) & msgajuste
elseif int(salajustado)>=int(rs("n4")) and int(salajustado)<int(rs("n5"))then
	quadro4=formatnumber(rs3("salario"),2) & msgajuste
	'quadro4=formatnumber(rss("salario"),2) & msgajuste
elseif int(salajustado)=int(rs("n5")) then
	quadro5=formatnumber(rs3("salario"),2) & msgajuste
	'quadro5=formatnumber(rss("salario"),2) & msgajuste
elseif int(salajustado)<int(rs("n1")) then
	quadro0=formatnumber(rs3("salario"),2) & msgajuste
	'quadro0=formatnumber(rss("salario"),2) & msgajuste
elseif int(salajustado)>int(rs("n5")) then
	quadro9=formatnumber(rs3("salario"),2) & msgajuste
	'quadro9=formatnumber(rss("salario"),2) & msgajuste
end if

'rss.close

%>
<tr>
	<td class="campor"><%=rs3("funcao")%>-<%=rs3("codfuncao")%></td>
	<td class="campor"><%=rs3("descricao")%></td>
	<td class="campor"><%=rs3("chapa")%></td>
	<td class="campor"><%=rs3("nome")%></td>
	<td class="campor"><%=rs3("dataadmissao")%></td>
	<td class="campor" align="center"><%=rs3("jornada")%></td>
	<td class="campor" align="center"><%=rs3("gruposalarial")%> (<%=rs3("codnivelsal")%>)</td>
	<td class="campor" align="center"><%=quadro0%></td>
	<td class="campor" align="center"><%=quadro1%></td>
	<td class="campor" align="center"><%=quadro2%></td>
	<td class="campor" align="center"><%=quadro3%></td>
	<td class="campor" align="center"><%=quadro4%></td>
	<td class="campor" align="center"><%=quadro5%></td>
	<td class="campor" align="center"><%=quadro9%></td>
	<td class="campor" ><%=status%></td>
</tr>
<%
rs3.movenext
tfaixa1="":tfaixa2="":tfaixa3="":tfaixa4="":tfaixa5=""
loop
end if 'recordcount rs3
rs3.close

if vazio=1 then
%>
</table>
&nbsp;
<!-- fim funcionarios -->
<%
end if

rs.movenext
loop
rs.close
%>

<!-- fim quadro do faixas do setor -->

<%
rs2.movenext
if request.form("quebrapag")="ON" then 

else
	if rs2.absoluteposition<rs2.recordcount and vazio=1 then response.write "<DIV style=""page-break-after:always""></DIV>"
end if
vazio=0
loop
rs2.close
end if 'request.form
%>

<%
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>