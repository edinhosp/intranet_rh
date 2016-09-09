<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a64")="N" or session("a64")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Titular de Bolsa de Estudos</title>
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

sqla="select f.nome, f.codsituacao, f.chapa, f.dataadmissao, c.nome as funcao, " & _
"f.codsecao, s.descricao AS Secao, p.grauinstrucao, f.datademissao " & _
"FROM corporerm.dbo.PFUNC AS f, corporerm.dbo.PPESSOA AS p, corporerm.dbo.PFUNCAO AS c, corporerm.dbo.PSECAO AS s " & _
"WHERE f.CODPESSOA=p.CODIGO and f.CODFUNCAO=c.CODIGO and f.CODSECAO=s.CODIGO "

if request.form("codigo")="" then codigo=request("codigo") else codigo=request.form("codigo")
sqlb="AND f.CHAPA='" & codigo & "' "
sqlc="ORDER BY f.CHAPA"
sql1=sqla & sqlb & sqlc
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>BOLSA DE ESTUDOS
<%
'rs.movefirst
'do while not rs.eof 
'session("chapabolsa")=rs("chapa")
'session("chapabolsanome")=rs("nome")
select case rs("codsituacao")
	case "A"
  		sit="Ativo"
	case "D"
		sit="Demitido"
	case "E"
		sit="Licenca Mater."
	case "F"
		sit="Ferias"
	case "I"
		sit="Apos. Invalidez"
	case "L"
		sit="Licenca s/venc"
	case "M"
		sit="Serv.Militar"
	case "O"
		sit="Doenca Ocupacional"
	case "P"
		sit="Af.Previdencia"
	case "R"
		sit="Licenca Remun."
	case "T"
		sit="Af.Ac.Trabalho"
	case "U"
		sit="Outros"
	case "V"
		sit="Aviso Previo"
	case "X"
		sit="C/Dem.no mes"
	case "Z"
		sit="Admissao prox.mes"
end select
select case rs("grauinstrucao")
	case "1"
		titulacao="Analfabeto"
	case "2"
		titulacao="Primario incompleto"
	case "3"
		titulacao="Primario completo"
	case "4"
		titulacao="Ginasial incompleto"
	case "5"
		titulacao="Ginasial completo"
	case "6"
		titulacao="Colegial incompleto"
	case "7"
		titulacao="Colegial completo"
	case "8"
		titulacao="Superior incompleto"
	case "9"
		titulacao="Graduado"
	case "A"
		titulacao="Especialista incompleto"
	case "B"
		titulacao="Especialista"
	case "C"
		titulacao="Mestrando"
	case "D"
		titulacao="Mestre"
	case "E"
		titulacao="Doutorando"
	case "F"
		titulacao="Doutor"
	case "G"
		titulacao="Livre Docente Incompleto"
	case "H"
		titulacao="Livre Docente"
end select
%>
<input type="hidden" name="codigo" value="<%=request("codigo")%>">
<table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" width="650">
<tr><td class=grupo>Dados Pessoais</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td width="100%" valign="top">
<!-- quadro -->
<table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>Chapa:</td>
	<td class=titulo>Nome:</td>
</tr>
<tr>
	<td class=campo>&nbsp;<%=rs("chapa")%>&nbsp;</td>
	<td class=campo>&nbsp;<%=rs("nome")%>&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>Situação:</td>
	<td class=titulo>Admissão:</td>
	<td class=titulo>Função:</td>
</tr>
<tr>
	<td class=campo>&nbsp;<%=sit%>&nbsp;<%if rs("codsituacao")="D" then response.write rs("datademissao")%></td>
	<td class=campo>&nbsp;<%=rs("dataadmissao")%>&nbsp;</td>
	<td class=campo>&nbsp;<%=rs("funcao")%>&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>Instrução/Titulação:</td>
	<td class=titulo>Seção:</td>
</tr>
<tr>
	<td class=campo>&nbsp;<%=titulacao%>&nbsp;</td>
	<td class=campo>&nbsp;<%=rs("codsecao")%>&nbsp;<%=rs("secao")%></td>
</tr>
</table>
<!-- quadro -->
	</td>
	<td width="170" valign="top">
	<font size="2">
	<img border="0" src="../func_foto.asp?chapa=<%=rs("chapa")%>"  width="140"></font>
	</td>
</tr>
</table>

<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=9>Outras Bolsas concecidadas a este Titular</td></tr>
<tr>
	<td class=titulor>Nome Bolsista</td>
	<td class=titulor>Parentesco/Tipo</td>
	<td class=titulor>Idade</td>
	<td class=titulor>Tipo Curso</td>
	<td class=titulor>Curso</td>
	<td class=titulor>Instituição</td>
	<td class=titulor>Tipo Bolsa</td>
	<td class=titulor>Situação</td>
	<td class=titulor>Compr.</td>
</tr>
<%
sql5="select b.id_bolsa, b.chapa, b.nome_bolsista, b.parentesco, b.dtnasc, b.tipocurso, b.curso, b.instituicao, b.matricula, " & _
"b.comprovante, b.observacao, b.tp_bolsa, t.descricao as dtp_bolsa, b.situacao, s.descricao as dsituacao " & _
"from bolsistas b, bolsistas_situacao s, bolsistas_tipo t " & _
"where s.id_sit=b.situacao and t.id_tp=b.tp_bolsa " & _
"and b.chapa='" & rs("chapa") & "' "
rs2.Open sql5, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
idade=int((now()-rs2("dtnasc"))/365.25)
%>		
<tr>
	<td class="campor">
	<a class=t href="bolsa_ver.asp?codigo=<%=rs2("id_bolsa")%>" onclick="NewWindow(this.href,'AlteracaoBolsa','690','600','no','center');return false" onfocus="this.blur()">	
	<%=rs2("nome_bolsista")%></a>
	</td>
	<td class="campor"><%=rs2("parentesco")%></td>
	<td class="campor"><%=idade%></td>
	<td class="campor"><%=rs2("tipocurso")%></td>
	<td class="campor"><%=rs2("curso")%></td>
	<td class="campor"><%=rs2("instituicao")%></td>
	<td class="campor"><%=rs2("dtp_bolsa")%></td>
	<td class="campor"><%=rs2("dsituacao")%></td>
	<td class="campor" valign="top">&nbsp;<%if rs2("comprovante")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
</tr>
<%
rs2.movenext
loop
end if
rs2.close
%>
</table>


</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>