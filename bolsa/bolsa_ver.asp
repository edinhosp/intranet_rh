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
<title>Bolsistas</title>
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

if request.form("codigo")="" then codigo=request("codigo") else codigo=request.form("codigo")

sql1="select b.id_bolsa, b.chapa, b.nome_bolsista, b.parentesco, b.dtnasc, b.tipocurso, b.curso, b.instituicao, b.matricula, " & _
"b.comprovante, b.observacao, b.tp_bolsa, t.descricao as dtp_bolsa, b.situacao, s.descricao as dsituacao " & _
"from bolsistas b, bolsistas_situacao s, bolsistas_tipo t " & _
"where s.id_sit=b.situacao and t.id_tp=b.tp_bolsa " & _
"and b.id_bolsa=" & codigo
rs.Open sql1, ,adOpenStatic, adLockReadOnly
sql2="select e.matricula, e.idimagem from corporerm.dbo.ealunos e where e.matricula='" & rs("matricula") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then imagem=rs2("idimagem") else imagem=0
rs2.close
sql2="select u.mataluno, u.status, s.descricao from corporerm.dbo.ualucurso u, corporerm.dbo.usitmat s where u.status=s.codsitmat and u.mataluno='" & rs("matricula") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then situacaoclassis=rs2("descricao") else situacaoclassis=""
rs2.close
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>BOLSA DE ESTUDOS</p>
<%
'session("chapabolsa")=rs("chapa")
'session("chapabolsanome")=rs("nome")
%>
<input type="hidden" name="codigo" value="<%=request("codigo")%>">
<table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" width="650">
<tr><td class=grupo colspan=5>Dados Pessoais</td></tr>
<tr>
	<td class=titulo>ID:</td>
	<td class=titulo colspan=3>Nome do Bolsista:</td>
	<td class=fundo width="140" valign="top" rowspan=8>
<%if rs("parentesco")="Titular" then %>
	<img border="0" src="../func_foto.asp?chapa=<%=rs("chapa")%>" width="140">
<%else%>
	<img border="0" src="../aluno_foto.asp?id=<%=imagem%>" width=140 >
<%end if%>
	</td>
</tr>
<tr>
	<td class=campo><%=rs("id_bolsa")%></td>
	<td class=campo colspan=3><b><%=rs("nome_bolsista")%></td>
</tr>
<tr>
	<td class=titulo>Parentesco</td>
	<td class=titulo>Data Nascimento</td>
	<td class=titulo>Situação Atual</td>
	<td class=titulo>Comprovante</td>
</tr>
<tr>
	<td class=campo><%=rs("parentesco")%></td>
	<td class=campo><%=rs("dtnasc")%> (<%=int((now-rs("dtnasc"))/365.25)%>)</td>
	<td class=campo><%=rs("dsituacao")%></td>
	<td class=campo><%if rs("comprovante")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
</tr>
<tr>
	<td class=titulo>Tipo do Curso:</td>
	<td class=titulo>Curso:</td>
	<td class=titulo>Instituição:</td>
	<td class=titulo>Matrícula:</td>
</tr>
<tr>
	<td class=campo><%=rs("tipocurso")%></td>
	<td class=campo><%=rs("curso")%></td>
	<td class=campo><%=rs("instituicao")%></td>
	<td class=campo>
		<a href="historico.asp?matricula=<%=rs("matricula")%>" onclick="NewWindow(this.href,'HistoricoEscolar','550','300','yes','center');return false" onfocus="this.blur()">
		<%=rs("matricula")%></a>
	</td>
</tr>
<tr>
	<td class=titulo>Tipo Bolsa</td>
	<td class=titulo colspan=2>Observação</td>
	<td class=titulo>Situação no Classis</td>
</tr>
<tr>
	<td class=campo><%=rs("dtp_bolsa")%></td>
	<td class=campo colspan=2><%=rs("observacao")%></td>
	<td class=campo><%=situacaoclassis%></td>
</tr>
</table>
<hr>
<table width=650><tr><td width=510>
<%
'rs.movenext
'loop
if rs("tp_bolsa")<>"6" then
sql2="select l.id_lanc, l.id_bolsa, l.renovacao, l.validade, l.ano_letivo, l.observacao, s.descricao, l.protocolo, " & _
"l.id_faculdade, l.curso, l.periodo " & _
"FROM bolsistas_lanc l, bolsistas_situacao s " & _
"WHERE s.id_sit=l.situacao AND l.id_bolsa=" & rs("id_bolsa") & " order by l.renovacao "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
%>
<table border="1" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Renovação</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Validade</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Período Letivo</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Situação</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Observação</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Protoc.</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">&nbsp;</td>
</tr>
<%
rs2.movefirst:do while not rs2.eof
%>
<tr>
	<td class="campor" align="center"><%=rs2("renovacao") %></td>
	<td class="campor" align="center"><%=rs2("validade") %></td>
	<td class="campor"><%=rs2("ano_letivo") %></td>
	<td class="campor"><%=rs2("descricao") %></td>
	<td class="campor"><%=rs2("observacao") %></td>
	<td class="campor" align="center"><%if rs2("protocolo")=1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor">&nbsp; 
	<% if session("a64")="T" and (rs2("protocolo")=0 or session("usuariomaster")="02379" or session("usuariomaster")="00259" or session("usuariomaster")="02977")  then %>
		<a href="lanc_alteracao.asp?codigo=<%=rs2("id_lanc")%>" onclick="NewWindow(this.href,'AlteracaoLanc','520','150','no','center');return false" onfocus="this.blur()">
		<img src="../images/Folder95O.gif" border="0" width=13 Alt="Alterar este lançamento"></a>
	<% end if %>
	</td>
</tr>
<%
rs2.movenext
loop
end if
rs2.close
%>
</table>
<%
else 'tipo bolsa
sql2="select l.id_lanc, l.id_bolsa, l.situacao, l.renovacao, l.validade, l.ano_letivo, l.observacao, s.descricao, l.protocolo, " & _
"l.id_faculdade, l.curso, l.periodo, f.faculdade " & _
"FROM bolsistas_lanc l, bolsistas_situacao s, rhconveniobe f " & _
"WHERE l.id_faculdade=f.id_faculdade and s.id_sit=l.situacao AND l.id_bolsa=" & rs("id_bolsa") & " order by l.renovacao "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
<table border="1" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Data</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Faculdade</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Curso</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Período</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">P.Letivo</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Tipo</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">Sit.</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">&nbsp;</td>
	<td style="border-bottom: 2px solid #000000" class=titulor align="center">&nbsp;</td>
</tr>
<%
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
tipoprot="?"
if rs2("protocolo")=1 then tipoprot="Inscr.Vestibular"
if rs2("protocolo")=2 then tipoprot="Matrícula"
if rs2("protocolo")=3 then tipoprot="Rematrícula"
session("64faculdade")=rs2("id_faculdade")
session("64curso")=rs2("curso")
session("64periodo")=rs2("periodo")
session("64protocolo")=rs2("protocolo")
%>
<tr>
	<td class="campor" align="center"><%=rs2("renovacao") %></td>
	<td class="campor"><%=rs2("faculdade") %></td>
	<td class="campor"><%=rs2("curso") %></td>
	<td class="campor"><%=rs2("periodo") %></td>
	<td class="campor"><%=rs2("ano_letivo") %></td>
	<td class="campor"><%=tipoprot%></td>
	<td class="campor" align="center"><%=rs2("situacao")%></td>
	<td class="campor">&nbsp; 
	<% if session("a64")="T" then %>
		<a href="convenio_alteracao.asp?codigo=<%=rs2("id_lanc")%>" onclick="NewWindow(this.href,'AlteracaoLanc','520','250','no','center');return false" onfocus="this.blur()">
		<img src="../images/Folder95O.gif" border="0" width=13 Alt="Alterar este lançamento"></a>
	<% end if %>
	</td>
	<td class="campor">
	<a href="form_enviado.asp?codigo=<%=rs2("id_lanc")%>" onclick="NewWindow(this.href,'FormularioConvenio','660','400','yes','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/leaf.gif" alt="Imprimir formulário"></a>
	</td>
</tr>
<%
rs2.movenext:loop
end if
rs2.close
%>
</table>
<%
end if 'tipo bolsa
%>

</td><td valign=top align="right">
<% if session("a64")="T" then %>
<%if rs("tp_bolsa")<>"6" then%>
<a href="lanc_nova.asp?codigo=<%=rs("id_bolsa")%>" onclick="NewWindow(this.href,'InclusaoLanc','520','150','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo lançamento">
<font size="1">Novo Lançamento</font></a>
<%else%>
<a href="convenio_nova.asp?codigo=<%=rs("id_bolsa")%>" onclick="NewWindow(this.href,'InclusaoLanc','520','250','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo lançamento">
<font size="1">Novo Lançamento</font></a>
<%end if%>
<% end if %>
</td></tr></table>
<hr>


<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8>Lançamentos de Bolsas no RM Classis</td></tr>
<tr>
	<td class=titulor>Per.Letivo</td>
	<td class=titulor>Status</td>
	<td class=titulor>Tipo Bolsa</td>
	<td class=titulor>Valor Desc.</td>
	<td class=titulor>Tipo Desc.</td>
	<td class=titulor>Inicio</td>
	<td class=titulor>Término</td>
	<td class=titulor>&nbsp;</td>
</tr>
<%
sql3="SELECT b.MATALUNO, b.PERLETIVO, b.CODBOL, t.DESCBOLSA, b.PERCDESC, b.TIPODESC, b.DTINICIO, b.DTFIM, b.TIPO, s.DESCRICAO " & _
"FROM (corporerm.dbo.ealubolsa AS b LEFT JOIN corporerm.dbo.etipobols AS t ON b.CODBOL = t.CODBOLSA) LEFT JOIN (corporerm.dbo.umatricpl AS u LEFT JOIN corporerm.dbo.usitmat AS s ON u.STATUS = s.CODSITMAT) ON (b.MATALUNO = u.MATALUNO) AND (b.PERLETIVO = u.PERLETIVO) " & _
"WHERE b.MATALUNO='" & rs("matricula") & "' ORDER BY b.PERLETIVO, b.DTINICIO "
rs2.Open sql3, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
tipob=""
if rs2("tipo")=8 then tipob="Somar bolsas"
if rs2("tipo")=9 then tipob="Aplicar bolsas em cascata"
if rs2("tipo")=10 then tipob="Utilizar o maior desconto"
%>		
<tr>
	<td class="campor"><%=rs2("perletivo")%></td>
	<td class="campor"><%=rs2("descricao")%></td>
	<td class="campor"><%=rs2("descbolsa")%></td>
	<td class="campor" align="right"><%=rs2("percdesc")%></td>
	<td class="campor"><%=rs2("tipodesc")%></td>
	<td class="campor"><%=rs2("dtinicio")%></td>
	<td class="campor"><%=rs2("dtfim")%></td>
	<td class="campor"><%=tipob%></td>
</tr>
<%
rs2.movenext
loop
end if
rs2.close
%>
</table>
<hr>

<!-- Titular e outras bolas -->
<%
sql4="select f.chapa, f.nome, f.codsituacao, f.situacao from qry_funcionarios f " & _
"where f.chapa='" & rs("chapa") & "' "
rs2.Open sql4, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=7>Funcionário Titular das Bolsas</td></tr>
<tr>
	<td class="campot"r colspan=7><b><%=rs2("chapa")%> - <%=rs2("nome")%></td>
</tr>
<tr><td class=grupo colspan=7>Outras Bolsas concecidadas a este Titular</td></tr>
<tr>
	<td class=titulor>Nome Bolsista</td>
	<td class=titulor>Parentesco/Tipo</td>
	<td class=titulor>Tipo Curso</td>
	<td class=titulor>Curso</td>
	<td class=titulor>Instituição</td>
	<td class=titulor>Tipo Bolsa</td>
	<td class=titulor>Situação</td>
</tr>
<%
rs2.close
sql5="select b.id_bolsa, b.chapa, b.nome_bolsista, b.parentesco, b.dtnasc, b.tipocurso, b.curso, b.instituicao, b.matricula, " & _
"b.comprovante, b.observacao, b.tp_bolsa, t.descricao as dtp_bolsa, b.situacao, s.descricao as dsituacao " & _
"from bolsistas b, bolsistas_situacao s, bolsistas_tipo t " & _
"where s.id_sit=b.situacao and t.id_tp=b.tp_bolsa " & _
"and b.chapa='" & rs("chapa") & "' " & _
"order by b.nome_bolsista "
rs2.Open sql5, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>		
<tr>
	<td class="campor">
	<a class=t href="bolsa_ver.asp?codigo=<%=rs2("id_bolsa")%>">
	<%=rs2("nome_bolsista")%></a>
	</td>
	<td class="campor"><%=rs2("parentesco")%></td>
	<td class="campor"><%=rs2("tipocurso")%></td>
	<td class="campor"><%=rs2("curso")%></td>
	<td class="campor"><%=rs2("instituicao")%></td>
	<td class="campor"><%=rs2("dtp_bolsa")%></td>
	<td class="campor"><%=rs2("dsituacao")%></td>
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