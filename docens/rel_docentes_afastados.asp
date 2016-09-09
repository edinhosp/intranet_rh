<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a16")="N" or session("a16")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Docentes Afastados</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("consql")

sqla="select f.codsituacao, si.descricao, f.chapa, f.nome, f.dataadmissao, f.codsecao, se.descricao as secao, " & _
"f.codfuncao, ca.nome as funcao, p.grauinstrucao, ci.descricao as instrucao, motivos.descmud, motivos.dtinicio, motivos.dtfinal  " & _
"from pfunc f, pcodsituacao si, psecao se, pfuncao ca, pcodinstrucao ci, ppessoa as p , " & _
"(SELECT PFHSTAFT.CHAPA, PFHSTAFT.DTINICIO, PFHSTAFT.DTFINAL, PFHSTAFT.TIPO, PCODSITUACAO.DESCRICAO AS desccod, PFHSTAFT.MOTIVO, PMUDSITUACAO.DESCRICAO AS descmud " & _
"FROM (PFHSTAFT INNER JOIN PCODSITUACAO ON PFHSTAFT.TIPO = PCODSITUACAO.CODCLIENTE) INNER JOIN PMUDSITUACAO ON PFHSTAFT.MOTIVO = PMUDSITUACAO.CODCLIENTE " & _
"WHERE (((PFHSTAFT.DTFINAL) Is Null Or (PFHSTAFT.DTFINAL)>getdate()))) as motivos " & _
"where f.codsituacao=si.codcliente and f.codfuncao=ca.codigo and f.codsecao=se.codigo and " & _
"f.codpessoa=p.codigo and p.grauinstrucao=ci.codcliente " & _
"and motivos.chapa=f.chapa " & _
"and f.codsituacao not in ('A','D') and f.codsindicato='03' " & _
"order by f.nome "
	set rs=server.createobject ("ADODB.Recordset")
	Set rs.ActiveConnection = conexao
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>

<p class=titulo>Docentes Afastados
<table border="1" bordercolor=#000000 cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
  <tr>
    <td class="titulor" align="center">Chapa</td>
    <td class="titulor" align="center">Nome</td>
    <td class="titulor" align="center">Admissão</td>
    <td class="titulor" align="center">Setor</td>
    <td class="titulor" align="center">Data Afast.</td>
    <td class="titulor" align="center">Situação</td>
    <td class="titulor" align="center">Motivo</td>
  </tr>
<%
linhas=2
rs.movefirst
do while not rs.eof 
chapach=rs("chapa")
session("chapa")=chapach
if rs("codsituacao")="E" then classe="campoar" else classe="campor"
%>
  <tr>
    <td height=28 class=<%=classe%> align="center"><%=rs("chapa")%></td>
    <td height=28 class=<%=classe%>><%=rs("nome")%></td>
    <td height=28 class=<%=classe%> align="center"><%=rs("dataadmissao")%></td>
    <td height=28 class=<%=classe%>><%=rs("secao")%></td>
    <td height=28 class=<%=classe%> align="center"><%=rs("dtinicio")%><%if rs("dtfinal")<>"" then response.write " a " & rs("dtfinal")%></td>
    <td height=28 class=<%=classe%>><%=rs("descricao")%></td>
    <td height=28 class=<%=classe%>><%=rs("descmud")%></td>
  </tr>
<%
linhas=linhas+1
rs.movenext
loop
%>
</table>
<%	pagina=pagina+1
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
%>
<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>