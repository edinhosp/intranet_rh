<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a86")="N" or session("a86")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Carta aos funcionários</title>
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
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao


sql3="SELECT ad.chapa, ab.nome, ad.dependente, ad.nascimento, ad.parentesco, adm.empresa, adm.plano, adm.ivigencia, adm.fvigencia, " & _
"(dateadd(""yy"",-21,dateadd(""mm"",1,dateadd(""dd"",-day(getdate())+1,getdate())))) as expr1 " & _
"FROM assmed_dep ad, assmed_beneficiario ab, corporerm.dbo.pfunc f, corporerm.dbo.psecao s, assmed_dep_mudanca adm " & _
"WHERE ad.chapa=ab.chapa and ab.chapa=f.chapa collate database_default and f.codsecao=s.codigo and ad.id_dep=adm.id_dep " & _
"AND ad.nascimento<(dateadd(""yy"",-21,dateadd(""mm"",1,dateadd(""dd"",-day(getdate())+1,getdate())))) " & _
"AND ad.parentesco like 'filh%' AND adm.fvigencia>getdate() and adm.empresa in ('U','V','I','BS') " & _
"ORDER BY ad.nascimento desc "

sqlb="SELECT ad.chapa, ab.nome, f.codsecao, s.descricao, ad.dependente, " & _
"(convert(char,year(getdate())-21)+'/'+convert(char,month(getdate())+1)+'/01') as expr1, ad.nascimento, ad.parentesco, adm.empresa, " & _
"adm.plano, adm.ivigencia, adm.fvigencia " & _
"FROM assmed_dep ad, assmed_beneficiario ab, corporerm.dbo.pfunc f, corporerm.dbo.psecao s, assmed_dep_mudanca adm " & _
"WHERE ad.chapa=ab.chapa and ab.chapa=f.chapa collate database_default and f.codsecao=s.codigo and ad.id_dep=adm.id_dep " & _
"AND ad.nascimento<(convert(char,year(getdate())-21)+'/'+convert(char,month(getdate())+1)+'/01') " & _
"AND ad.parentesco like 'filh%' AND adm.fvigencia>getdate() and adm.empresa in ('U','I','BS') " & _
"ORDER BY ad.nascimento desc "

sqlb="SELECT ad.chapa, f.nome, f.codsecao, ad.nascimento, ad.parentesco, s.descricao, ad.dependente, adm.plano, adm.ivigencia, adm.fvigencia, " & _
"(dateadd(""yy"",-21,dateadd(""mm"",1,dateadd(""dd"",-day(getdate())+1,getdate())))) as expr1 " & _
"FROM assmed_dep ad, assmed_beneficiario ab, corporerm.dbo.pfunc f, corporerm.dbo.psecao s, assmed_dep_mudanca adm " & _
"WHERE ad.chapa=ab.chapa and ab.chapa=f.chapa collate database_default and f.codsecao=s.codigo and ad.chapa=adm.chapa and ad.nrodepend=adm.nrodepend " & _
"AND ad.nascimento<(dateadd(""yy"",-21,dateadd(""mm"",1,dateadd(""dd"",-day(getdate())+15,getdate())))) " & _
"AND ad.parentesco like 'filh%' AND adm.fvigencia>getdate() and adm.empresa in ('U','I','BS') " & _
"ORDER BY ad.nascimento desc "


rs.Open sqlb, ,adOpenStatic, adLockReadOnly

if rs.recordcount>0 then
do while not rs.eof
%>
<div align="right">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td><img border="0" src="../images/logo_centro_universitario_unifieo_big.jpg" width=225></td></tr>
<tr><td><p align="center">&nbsp;</td></tr>
<tr>
	<td>Ao Sr(a).<br>
    <b><%=rs("nome")%></b> (<%=rs("chapa")%>)<br>
    <%=rs("codsecao")%>-&nbsp;<%=rs("descricao")%><br><br>
    Ref.: Dependente com mais de 21 anos<br><br></p>
    <p align="justify">Vimos comunicar-lhe que pelo motivo dos seus dependentes
	abaixo relacionados terem atigindio a idade limite de 21 anos, estarão sendo excluídos do
	plano de assistência médica <%=rs("plano")%> a partir desta data.<br>
	Solicitamos enviar as respectivas carteiras de saúde ao departamento de Recursos Humanos para 
	serem devolvidas à empresa de assistência médica.<br><br>
	Lembramos que o uso indevido da carteira de saúde poderá ser cobrado do titular pela empresa
	de saúde.</p>
    </td></tr>

<tr>
	<td>Dependentes a serem excluídos:	
	<table border="1" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width=500>
	<tr>
		<td>Nome do Dependente</td>
		<td>Nascimento</td>
		<td>Idade</td>
		<td>Parentesco</td>
	</tr>
	<tr>
		<td><%=rs("dependente")%></td>
		<td><%=rs("nascimento")%></td>
		<td><%=int((now()-rs("nascimento"))/365.25)%></td>
		<td><%=rs("parentesco")%></td>
	</tr>
	</table>		
	</td>
</tr>

<tr><td>&nbsp;</td></tr>
<tr><td>Atenciosamente,    </td></tr>

<tr>
	<td>
	<p align="left">Osasco,&nbsp;<%=day(now)%> de <%=monthname(month(now))%> de <%=year(now)%>
	<p>&nbsp;
	<p>______________________________________________<br>
	Fundação Instituto de Ensino para Osasco<br>
	Recursos Humanos&nbsp;</td>
</tr>
</table>
</div>
<%
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
rs.movenext
loop
%>
</body>
</html>
<%
else
%>
<p style="color:red">Não existem dependentes de assistência médica com idade maior de 21 anos.

<%
end if
rs.close
set rs=nothing

conexao.close
set conexao=nothing
%>