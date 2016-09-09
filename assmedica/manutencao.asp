<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a84")="N" or session("a84")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Manutenção Assistência Médica</title>
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
conexao.Open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
sessao=session.sessionid
%>
<p class=titulo>Checagem da Assistência Médica
<%
if now()>dateserial(2006,6,1) then
response.write "<font color=blue>"
response.write "<p style='margin-top:0;margin-bottom:0'><b>---"
response.write "<font color=black>"
end if

'***** atualizacao de novos funcionários  ******
sql1="SELECT CHAPA, nome " & _
"FROM corporerm.dbo.pfunc WHERE CHAPA<'10000' AND CODSITUACAO<>'D'"
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
rs2.movefirst
response.write "<p>Funcionários ativos: " & rs2.recordcount
total=0
do while not rs2.eof
	sql2="select chapa from assmed_beneficiario where chapa='" & rs2("chapa") & "'"
	rs.Open sql2, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
	
	else
		sql3="insert into assmed_beneficiario (chapa, empresa) " & _
		"select '" & rs2("chapa") & "', 'N'"
		response.write "<br>" & sql3
		conexao.execute sql3
		total=total+1
		response.write "<br>Inseriu " & rs2("chapa") & " - " & rs2("nome")
	end if
	rs.close
rs2.movenext
loop
rs2.close
response.write "<p>Novos funcionários: " & total

sql="delete from ttassmedplatual where sessao='" & sessao & "' "
conexao.execute sql

sql="insert into ttassmedplatual (sessao, chapa, inicio) " & _
"SELECT '" & sessao & "', chapa, Max(ivigencia) AS inicio from assmed_mudanca where empresa not in ('T','IP','MP','O','BP','V') GROUP BY chapa "
conexao.execute sql

sql="drop table ttassmedplatualt"
sql="if exists (select 'True' from sysobjects where name='ttassmedplatualt') drop table ttassmedplatualt"
conexao.execute sql

sql="SELECT am.chapa, am.empresa as emp, am.plano as pl, am.codigo as cod INTO ttassmedplatualt " & _
"FROM ttassmedplatual a1 INNER JOIN assmed_mudanca am ON (a1.chapa = am.chapa) AND (a1.inicio = am.ivigencia) " & _
"WHERE a1.sessao='" & sessao & "' and am.empresa not in ('T','IP','MP','O','M','UC','BS','BP','V') "
conexao.execute sql

sql="UPDATE ttassmedplatualt INNER JOIN assmed_beneficiario ON ttassmedplatualt.chapa = assmed_beneficiario.CHAPA " & _
"SET assmed_beneficiario.empresa = emp, assmed_beneficiario.plano = pl, assmed_beneficiario.codigo = cod "
sql="update assmed_beneficiario SET assmed_beneficiario.empresa = emp, assmed_beneficiario.plano = pl, assmed_beneficiario.codigo = cod " & _
"FROM ttassmedplatualt INNER JOIN assmed_beneficiario ON ttassmedplatualt.chapa = assmed_beneficiario.CHAPA "
conexao.execute sql

if request("limpar")="True" then
	sql="UPDATE assmed_mudanca SET oper = '', uoper = [oper] WHERE oper<>'';"
	conexao.execute sql
	sql="UPDATE assmed_dep_mudanca SET oper = '', uoper = [oper] WHERE oper<>'';"
	conexao.execute sql
end if

%>
<p class=titulo>Funcionários ainda sem cadastro de assistência médica
<table border="1" cellpadding="0" width="600" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Chapa  </td>
	<td class=titulo>Nome   </td>
	<td class=titulo>Empresa</td>
	<td class=titulo>Plano  </td>
	<td class=titulo align="center">Data Limite</td>
	<td class=titulo align="center">Faltam</td>
	<td class=titulo>Ass.Médica</td>
</tr>
<%
sql3="SELECT CHAPA, NOME, empresa, plano, codigo FROM assmed_beneficiario WHERE empresa='N'"
sql3="select a.*, f.nome, (f.dataadmissao+28) as datalimite, round((convert(float,f.dataadmissao)+28)-convert(float,getdate())+1,0) as dias, " & _
"saude=case when f.codsindicato='03' then 'Intermedica' else 'Unimed' end " & _
"from assmed_beneficiario a, corporerm.dbo.pfunc f where a.empresa='N' and f.chapa collate database_default=a.chapa and f.codsituacao<>'D' order by f.dataadmissao "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo>&nbsp;
	<% if session("a84")="T" then %>
		<a href="controle_ver.asp?codigo=<%=rs("chapa")%>" onclick="NewWindow(this.href,'Alteracao','690','500','yes','center');return false" onfocus="this.blur()">
		<font size="1"><%=rs("chapa")%></font></a>
	<% else %>
		<%=rs("chapa")%>
	<% end if %>
	</td>
	<td class=campo>&nbsp;<%=rs("nome")%>   </td>
	<td class=campo>&nbsp;<%=rs("empresa")%></td>
	<td class=campo>&nbsp;<%=rs("plano")%>  </td>
	<td class=campo align="center">&nbsp;<%=rs("datalimite")%> </td>
	<td class=campo align="center">&nbsp;<%=rs("dias")%> </td>
	<td class=campo>&nbsp;<%=rs("saude")%> </td>
</tr>
<%
rs.movenext
loop
else
	response.write "<td class=grupo colspan='7'>&nbsp;Sem cadastros pendentes.</td>"
end if
rs.close
%>
</table>
<DIV style="page-break-after:always"></DIV> <!-- Aqui quebra a página --> 
<p class=titulo>Funcionários demitidos a serem excluídos
<table border="1" cellpadding="0" width="600" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo align="center">Permanecer até</td>
	<td class=titulo>Chapa  </td>
	<td class=titulo>Nome   </td>
	<td class=titulo>Empresa</td>
	<td class=titulo>Plano  </td>
	<td class=titulo>Código </td>
</tr>
<%
SendIp=request.servervariables("LOCAL_ADDR")
sql3="SELECT getdate() AS Hoje, f.CODSITUACAO, m.ivigencia, m.fvigencia, b.CHAPA, f.NOME, m.empresa, m.plano, m.codigo, dt_canc " & _
"FROM (assmed_beneficiario b INNER JOIN assmed_mudanca m ON b.CHAPA=m.chapa) INNER JOIN corporerm.dbo.PFUNC f ON b.CHAPA = f.CHAPA collate database_default " & _
"WHERE ( (getdate() Between ivigencia And fvigencia) ) AND (f.CODSITUACAO='D' OR (f.codsituacao<>'D' and f.datademissao is not null)) " & _
"ORDER BY F.CHAPA "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo align="center"><%=rs("dt_canc")%> </td>
	<td class=campo>&nbsp;
    <% if session("a84")="T" then %>
	<a href="controle_ver.asp?codigo=<%=rs("chapa")%>" onclick="NewWindow(this.href,'Alteracao','690','500','yes','center');return false" onfocus="this.blur()">
		<font size="1"><%=rs("chapa")%></font></a>
	<% else %>
		<%=rs("chapa")%>
	<% end if %>
	</td>
	<td class=campo>&nbsp;<%=rs("nome")%>   </td>
	<td class=campo>&nbsp;<%=rs("empresa")%></td>
	<td class=campo>&nbsp;<%=rs("plano")%>  </td>
	<td class=campo>&nbsp;<%=rs("codigo")%> </td>
</tr>
<%
rs.movenext
loop
end if 'rs.recordcount
rs.close
%>
</table>

<DIV style="page-break-after:always"></DIV> <!-- Aqui quebra a página --> 
<p class=titulo>Dependentes Maiores de 21 anos
<table border="1" cellpadding="0" width="600" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo align="center">Permanecer até</td>
	<td class=titulo>Chapa  </td>
	<td class=titulo>Nome   </td>
	<td class=titulo>Dependente</td>
	<td class=titulo>Plano</td>
	<td class=titulo>Tipo  </td>
</tr>
<%
SendIp=request.servervariables("LOCAL_ADDR")
sql3="SELECT ad.chapa, f.nome, ad.dependente, ad.nascimento, ad.parentesco, adm.empresa, adm.plano, adm.ivigencia, adm.fvigencia, " & _
"(dateadd(""yy"",-21,dateadd(""mm"",1,dateadd(""dd"",-day(getdate())+1,getdate())))) as expr1 " & _
"FROM assmed_dep ad, assmed_beneficiario ab, corporerm.dbo.pfunc f, corporerm.dbo.psecao s, assmed_dep_mudanca adm " & _
"WHERE ad.chapa=ab.chapa and ab.chapa=f.chapa collate database_default and f.codsecao=s.codigo and ad.chapa=adm.chapa and ad.nrodepend=adm.nrodepend " & _
"AND ad.nascimento<(dateadd(""yy"",-21,dateadd(""mm"",1,dateadd(""dd"",-day(getdate())+1,getdate())))) " & _
"AND ad.parentesco like 'filh%' AND adm.fvigencia>getdate() and adm.empresa in ('U','I','BS') " & _
"ORDER BY ad.nascimento desc "
rs.Open sql3, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo align="center"></td>
	<td class=campo>&nbsp;
    <% if session("a84")="T" then %>
	<a href="controle_ver.asp?codigo=<%=rs("chapa")%>" onclick="NewWindow(this.href,'Alteracao','690','500','yes','center');return false" onfocus="this.blur()">
		<font size="1"><%=rs("chapa")%></font></a>
	<% else %>
		<%=rs("chapa")%>
	<% end if %>
	</td>
	<td class=campo>&nbsp;<%=rs("nome")%>   </td>
	<td class=campo>&nbsp;<%=rs("dependente")%></td>
	<td class=campo>&nbsp;<%=rs("plano")%></td>
	<td class=campo>&nbsp;<%=rs("parentesco")%>  </td>
</tr>
<%
rs.movenext
loop
end if 'rs.recordcount
rs.close
%>
</table>









<%
if session("usuariomaster")="02379" or session("usuariomaster")="02675" then
%>
<a href="manutencao.asp?limpar=True">Passar Status atual para Status anterior</a>
<%
end if

conexao.close
set conexao=nothing
set rs=nothing
set rs2=nothing
%>
</body>
</html>