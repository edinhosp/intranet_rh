<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a53")="N" or session("a53")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Vencimento de Contrato de Experiência</title>
<script language="javascript" type="text/javascript">
<!--
/****************************************************
     Author: Eric King
     Url: http://redrival.com/eak/index.shtml
     This script is free to use as long as this info is left in
     Featured on Dynamic Drive script library (http://www.dynamicdrive.com)
****************************************************/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>
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
%>

<p class=titulo>Contrato de Experiência&nbsp;<%=titulo %>

<%
if request.form="" then
data_1=dateserial(year(now),month(now),1)
data_2=dateserial(year(now),month(now)+1,1)-1
%>
<form method="POST" action="experiencia_termino.asp">
	<p>vencendo entre <input type="text" name="T1" size="12" value="<%=data_1%>" style="text-align:center">
	e <input type="text" name="T2" size="12" value="<%=data_2%>" style="text-align:center">
	<br>
	Data da tabela: <select size="1" name="datatabela">
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

	<input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>
<%
end if

if request.form<>"" then
	data1=dtaccess(request.form("t1"))
	data2=dtaccess(request.form("t2"))
	datatabela=dtaccess(request.form("datatabela"))
sqla="SELECT f.CHAPA, f.NOME, f.CODSECAO, f.CODFUNCAO, c.NOME as funcao, f.JORNADA, f.SALARIO, f.GRUPOSALARIAL, f.CODNIVELSAL, f.DATAADMISSAO, " & _
"[DATAADMISSAO]+44 AS Fim1P, [DATAADMISSAO]+45 AS Inicio2P, [DATAADMISSAO]+89 AS Fim2P, f.jornadamensal/60 as jornadames, " & _
"faixa=case when gruposalarial='N1' then N1 else case when gruposalarial='N2' then N2 else case when gruposalarial='N3' then N3 else case when gruposalarial='N4' then N4 else case when gruposalarial='N5' then N5 else 0 end end end end end, " & _
"faixa_ajustada=(case when gruposalarial='N1' then N1 else case when gruposalarial='N2' then N2 else case when gruposalarial='N3' then N3 else case when gruposalarial='N4' then N4 else case when gruposalarial='N5' then N5 else 0 end end end end end/Horas)*(Jornadamensal/60), " & _
"proporc=case s.id_setor when 'NCLA' then Salario else ceiling((((salario/30) * case day(dataadmissao+89) when 31 then 30 else day(dataadmissao+89) end) +(((case when gruposalarial='N1' then N1 else case when gruposalarial='N2' then N2 else case when gruposalarial='N3' then N3 else case when gruposalarial='N4' then N4 else case when gruposalarial='N5' then N5 else 0 end end end end end/horas) *(jornadamensal/60))/30) *(30-case day(dataadmissao+89) when 31 then 30 else day(dataadmissao+89) end)) *100+0.5)/100 end, " & _
"s.horas, s.n1, s.n2, s.n3, s.n4, s.n5, s.id_cargo, s.id_setor " & _
"FROM corporerm.dbo.PFUNC f inner join corporerm.dbo.PFUNCAO c on c.CODIGO collate database_default=f.CODFUNCAO " & _
"left join ( " & _
"	select cl.codfuncao, cl.id_setor, cl.id_cargo, cc.cargo, cc.horas, s.n1, s.n2, s.n3, s.n4, s.n5 " & _
"	from cs_cargos_lab cl inner join cs_salarios s on s.id_cargo=cl.id_cargo and s.id_setor=cl.id_setor " & _
"	inner join cs_cargos cc on cc.id_setor=cl.id_setor and cc.id_cargo=cl.id_cargo where s.data='" & datatabela & "' " & _
") s on s.codfuncao collate database_default=f.CODFUNCAO " & _
"WHERE f.CODTIPO<>'T' AND f.CODSINDICATO<>'03' AND f.CODSITUACAO<>'D' "

sqla1="and (DATAADMISSAO+44) Between '" & data1 & "' And '" & data2 & "' "
sqla2="and (DATAADMISSAO+89) Between '" & data1 & "' And '" & data2 & "' "
sqlb=" ORDER BY DATAADMISSAO+89, f.nome "
sql1=sqla & sqla1 & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=titulop>Relatório de Prazos de Experiência e ajuste salarial Pós-Admissional</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=grupo>Vencimentos de 1º Período</td></tr>
</table>  
<br>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="center">Nome</td>
	<td class=titulo align="center">Cod.</td>
	<td class=titulo align="center">Nome Função</td>
	<td class=titulo align="center">Admissão</td>
	<td class=titulo align="center">Venc. 1º Período</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 
%>
<tr>
	<td class=campo><%=rs("chapa")%></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo><%=rs("codfuncao") %></td>
	<td class=campo><%=rs("funcao") %></td>
	<td class=campo align="center"><%=rs("dataadmissao") %></td>
	<td class=campo align="center"><b><%=rs("fim1p")%></b></td>
</tr>
<%
rs.movenext
loop
end if 'recordcount
rs.close
%>
</table>
<%
sql1=sqla & sqla2 & sqlb
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<br>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="950">
<tr><td class=grupo>Vencimentos de 2º Período</td></tr>
</table>  
<br>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="950">
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="center">Nome</td>
	<td class=titulo align="center">Cod.</td>
	<td class=titulo align="center">Nome Função</td>
	<td class=titulo align="center">Admissão</td>
	<td class=titulo align="center">Venc.1ºPer</td>
	<td class=titulo align="center">Venc.2ºPer</td>
	<td class=titulo align="center">Salário<br>Atual</td>
	<td class=grupo align="center">Grp.</td>
	<td class=grupo align="center">Faixa<br>Salarial</td>
	<td class=grupo align="center">Jornada<br>Func.</td>
	<td class=grupo align="center">Jornada<br>Faixa</td>
	<td class=grupo align="center">Faixa<br>Ajustada</td>
	<td class=titulo align="center">Salário<br>Propor</td>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 
%>
<tr>
	<td class=campo><%=rs("chapa")%></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo><%=rs("codfuncao") %></td>
	<td class=campo><%=rs("funcao") %></td>
	<td class=campo align="center"><%=rs("dataadmissao") %></td>
	<td class=campo align="center"><%=rs("fim1p") %></td>
	<td class=campo align="center"><b><%=rs("fim2p")%></b></td>
	<td class=campo align="right"><%=formatnumber(rs("salario"),2)%>&nbsp;</td>
	<td class=campo align="center"><%=rs("gruposalarial") %></td>
	<td class=campo align="right"><%if isnull(rs("faixa")) then response.write "---" else response.write formatnumber(rs("faixa"),2)%>&nbsp;</td>
	<td class=campo align="center"><%=rs("jornadames") %></td>
	<td class=campo align="center"><%=rs("horas") %></td>
	<td class=campo align="right"><%if isnull(rs("faixa_ajustada")) then response.write "---" else response.write formatnumber(rs("faixa_ajustada"),2)%>&nbsp;</td>
	<td class=campo align="right"><%if isnull(rs("proporc")) then response.write "---" else response.write formatnumber(rs("proporc"),2)%>&nbsp;</td>
</tr>
<%
rs.movenext
loop
end if 'rs.recordcount
rs.close
%>
</table>
<%
end if '

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>