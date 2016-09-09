<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a37")="N" or session("a37")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Alunos</title>
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
dim conexao, conexao2, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set conexao2=server.createobject ("ADODB.Connection")
conexao2.open Application("Conexao")

sql1="select e.* from corporerm.dbo.ealunos e where e.matricula='" & request("matricula") & "' order by e.matricula "
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<p style="margin-top:0;margin-bottom:0" class=titulo>CADASTRO DE ALUNOS</p>
<%
'rs.movefirst:do while not rs.eof 'não há necessidade, é o unico registro
session("chapa")=rs("matricula")
session("chapanome")=rs("nome")
tabela=615
tbfoto=150
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Identificação  (<%=rs("codpessoa")%>)</td></tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td valign="top" class=fundo>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
  <tr>
    <td class=titulo>&nbsp;Matrícula</td>
    <td class=titulo>&nbsp;Nome</td>
  </tr>
  <tr>
    <td class=fundo>&nbsp;<input type="text" class=a size=15 value="<%=rs("matricula")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" class=a size=55 value="<%=rs("nome")%>" onfocus="this.blur()"></td>
  </tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
  <tr>
    <td class=titulo>&nbsp;Tipo</td>
    <td class=titulo>&nbsp;Data de Nascimento</td>
    <td class=titulo>&nbsp;Sexo</td>
  </tr>
<%
sql="select descricao from corporerm.dbo.pcodsexo where codcliente='" & rs("sexo") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then sexo=rs2("descricao")
rs2.close
if isnull(rs("tipoaluno")) then q01=0 else q01=rs("tipoaluno")
select case q01
	case 14	
		tipoaluno="BOLSA FAMILIA"
	case 15
		tipoaluno="GRAD. CONVÊNIO FAC."
	case 1
		tipoaluno="GRAD. NORMAL"
	case 4
		tipoaluno="GRAD. NORMAL VEST."
	case 2
		tipoaluno="GRAD. PREFEITURA"
	case 3
		tipoaluno="GRAD. PREFEITURA VEST."
	case 12
		tipoaluno="INGRESSANTE"
	case 13
		tipoaluno="NÃO MATRICULADO"
	case 6
		tipoaluno="PÓS EX-ALUNO"
	case 9
		tipoaluno="PÓS FUNCIONÁRIO"
	case 5
		tipoaluno="PÓS NORMAL"
	case 7
		tipoaluno="PÓS OAB"
	case 8
		tipoaluno="PÓS PARENTESCO"
	case 11
		tipoaluno="PÓS PREFEITURA"
	case 10
		tipoaluno="PÓS PROFESSOR"
end select
sql="select descricao from corporerm.dbo.utabtipo where codtipo=" & q01 & ""
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then tipoaluno=rs2("descricao")
rs2.close
%>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=1 value="<%=rs("tipoaluno")%>" onfocus="this.blur()">
	&nbsp;<input class=a type="text" size=20 value="<%=tipoaluno%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("dtnasc")%>" onfocus="this.blur()">&nbsp;(<%=int((now()-rs("dtnasc"))/365.25)%>)</td>
    <td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("sexo")%>" onfocus="this.blur()">
	&nbsp;<input type="text" size=10 value="<%=sexo%>" onfocus="this.blur()"></td>
  </tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
	<tr>
		<td class=titulo>&nbsp;Nacionalidade</td>
		<td class=titulo>&nbsp;Naturalidade</td>
	</tr>
<%
sql="select descricao from corporerm.dbo.pcodnacao where codcliente='" & rs("nacional") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then nacionalidade=rs2("descricao")
rs2.close
%>
	<tr>
		<td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("nacional")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=20 value="<%=nacionalidade%>" onfocus="this.blur()"></td>
		<td class=fundo>&nbsp;<input type="text" size=32 value="<%=rs("naturalidade")%>" onfocus="this.blur()"></td>
	</tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
	<tr>
		<td class=titulo>&nbsp;Estado Natal</td>
		<td class=titulo>&nbsp;Estado Civil</td>
	</tr>
<%
sql="select nome from corporerm.dbo.getd where codetd='" & rs("nates") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then estadonatal=trim(rs2("nome"))
rs2.close
sql="select descricao from corporerm.dbo.pcodestcivil where codcliente='" & rs("estcivil") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then estadocivil=trim(rs2("descricao"))
rs2.close
%>
	<tr>
		<td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("nates")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=30 value="<%=estadonatal%>" onfocus="this.blur()"></td>
		<td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("estcivil")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=20 value="<%=estadocivil%>" onfocus="this.blur()"></td>
	</tr>
</table>
    </td>
	<td width="<%=tbfoto%>" valign="top" class=fundo>
		<img border="0" src="../aluno_foto.asp?id=<%=rs("idimagem")%>"  width="<%=tbfoto%>">
	</td>
	</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
	<tr>
		<td class=titulo>&nbsp;E-mail</td>
		<td class=titulo valign=top width=300>&nbsp;Nome dos Pais</td>
	</tr>
	
	<tr>
		<td class=fundo valign=top>&nbsp;<input type="text" size=50 value="<%=rs("email")%>" onfocus="this.blur()">
		<br><%=rs("cedident")%>
		</td>
		<td class=fundo valign=top>
<%
if isnull(rs("pai")) then q02=0 else q02=rs("pai")
if isnull(rs("mae")) then q03=0 else q03=rs("mae")
if isnull(rs("respons")) then q04=0 else q04=rs("respons")
sql="select nome from corporerm.dbo.ppessoa where codigo=" & q02
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then pai=trim(rs2("nome"))
rs2.close
sql="select nome from corporerm.dbo.ppessoa where codigo=" & q03
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then mae=trim(rs2("nome"))
rs2.close
sql="select nome from corporerm.dbo.fcfo where codcfo='" & q04 & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then respons=trim(rs2("nome"))
rs2.close

%>
		<%=pai%><br><%=mae%><br>Respons.Financeiro: <%=respons%>
		</td>
	</tr>
</table>

<!-- cursos -->
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Cursos</td></tr>
</table>
<%
sqla="select u.codcur, c.nome, u.grade, u.codtun, t.descturno, u.status, s.descricao " & _
"from corporerm.dbo.ualucurso u, corporerm.dbo.ucursos c, corporerm.dbo.usitmat s, corporerm.dbo.eturnos t " & _
"where c.codcur=u.codcur and u.status=s.codsitmat and u.codtun=t.codturno " & _
"and u.mataluno='" & rs("matricula") & "' "
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
	<th class=titulo colspan=7>Cursos / Períodos
<%if session("a37")="T" then %>
  	<a class=t href="historico.asp?matricula=<%=rs("matricula")%>" onclick="NewWindow(this.href,'HistoricoEscolar','550','300','yes','center');return false" onfocus="this.blur()">
	<%=rs("matricula")%></a>
<% end if %>	
	
	</th>
	<tr>
		<td class=titulo rowspan=2>Curso</td>
		<td class=titulo rowspan=2>Grade</td>
		<td class=titulo rowspan=2>Turno</td>
		<td class=titulo rowspan=2>Status</td>
		<td class=titulo align="center" colspan=3>Períodos</td>
	</tr>
	<tr>
		<td class=titulo>Letivo</td>
		<td class=titulo>Per.#</td>
		<td class=titulo>Status</td>
	</tr>
<%
if rs3.recordcount>0 then
rs3.movefirst
do while not rs3.eof

sql2="select perletivo, periodo, status, descricao from corporerm.dbo.umatricpl u, corporerm.dbo.usitmat s " & _
"where s.codsitmat=u.status and mataluno='" & rs("matricula") & "' " & _
"and codcur=" & rs3("codcur") & " and grade='" & rs3("grade") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
linhas=rs2.recordcount:if linhas=0 then linhas=1
%>
	<tr>
		<td class=campo align="left" rowspan=<%=linhas%>><%=rs3("nome")%></td>
		<td class=campo align="left" rowspan=<%=linhas%>><%=rs3("grade")%></td>
		<td class=campo align="left" rowspan=<%=linhas%>><%=rs3("descturno")%></td>
		<td class=campo align="left" rowspan=<%=linhas%>><%=rs3("descricao")%></td>

<%
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
if rs2.absoluteposition>1 then response.write "<tr>"
%>
		<td class="campor"><%=rs2("perletivo")%></td>
		<td class="campor"><%=rs2("periodo")%></td>
		<td class="campor"><%=rs2("descricao")%></td>
		</tr>
<%
rs2.movenext
loop
else
%>
		<td class="campor" colspan=3></td></tr>
<%
end if
rs2.close
%>

<%
rs3.movenext
loop
end if
rs3.close
%>
</table>

<!-- fim cursos -->

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Endereço Principal</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo>&nbsp;Rua</td>
    <td class=titulo>&nbsp;Complemento</td>
  </tr>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=50 value="<%=rs("endaluno")%>" onfocus="this.blur()">
	<input type="text" size=10 value="<%=rs("numendalun")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=30 value="<%=rs("compendal")%>" onfocus="this.blur()"></td>
  </tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
	<tr>
		<td class=titulo>&nbsp;Bairro</td>
		<td class=titulo>&nbsp;Cidade</td>
		<td class=titulo>&nbsp;Estado</td>
	</tr>
<%
sql="select nome from corporerm.dbo.getd where codetd='" & rs("ufaluno") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then estado=trim(rs2("nome"))
rs2.close
%>
	<tr>
		<td class=fundo>&nbsp;<input type="text" size=30 value="<%=rs("bairroalun")%>" onfocus="this.blur()"></td>
		<td class=fundo>&nbsp;<input type="text" size=32 value="<%=rs("cidaluno")%>" onfocus="this.blur()"></td>
		<td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("ufaluno")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=30 value="<%=estado%>" onfocus="this.blur()"></td>
	</tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo>&nbsp;País</td>
    <td class=titulo>&nbsp;CEP</td>
    <td class=titulo>&nbsp;Telefone </td>
  </tr>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=16 value="<%=rs("paisaluno")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("cepaluno")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=35 value="<%=rs("telaluno")%>" onfocus="this.blur()"></td>
  </tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Bolsas</td></tr>
</table>

<%
sqla="SELECT B.*, T.DESCBOLSA from corporerm.dbo.EALUBOLSA B, corporerm.dbo.ETIPOBOLS T " & _
"WHERE B.CODBOL=T.CODBOLSA AND MATALUNO='" & rs("matricula") & "' ORDER BY B.perletivo"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
	<th class=titulo colspan=8>Cadastro de Bolsas</th>
	<tr>
		<td class=titulor>Período</td>
		<td class=titulor>Contrato</td>
		<td class=titulor>Tipo Bolsa</td>
		<td class=titulor>Tipo</td>
		<td class=titulor>Tipo Desc.</td>
		<td class=titulor>Desconto</td>
		<td class=titulor>Início</td>
		<td class=titulor>Término</td>
	</tr>
<%
if rs3.recordcount>0 then
rs3.movefirst
do while not rs3.eof
tipob=""
if rs3("tipo")=8 then tipob="Somar bolsas"
if rs3("tipo")=9 then tipob="Aplicar bolsas em cascata"
if rs3("tipo")=10 then tipob="Utilizar o maior desconto"
desconto=rs3("percdesc")
if rs3("tipodesc")="V" then tpdesc="Valor"
if rs3("tipodesc")="P" then tpdesc="Percentual"
if rs3("tipodesc")="V" and not isnull(rs3("percdesc")) then desconto=formatnumber(desconto,2)
if rs3("tipodesc")="P" and not isnull(rs3("percdesc")) then desconto=formatnumber(desconto,4)
%>
	<tr>
		<td class="campor" align="center"><%=rs3("perletivo")%></td>
		<td class="campor" align="left"><%=rs3("contrato")%></td>
		<td class="campor" align="left"><%=rs3("descbolsa")%></td>
		<td class="campor" align="left"><%=tipob%></td>
		<td class="campor" align="left"><%=rs3("tipodesc")%>-<%=tpdesc%></td>
		<td class="campor" align="right"><%=desconto%>&nbsp;</td>
		<td class="campor" align="center"><%=rs3("dtinicio")%></td>
		<td class="campor" align="center"><%=rs3("dtfim")%></td>
	</tr>
<%
rs3.movenext
loop
end if
rs3.close
%>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Vestibular</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo align="center" style="border: 1px solid #000000" colspan=2>&nbsp;Ensino Básico</td>
    <td class=titulo align="center" style="border: 1px solid #000000" colspan=2>&nbsp;Vestibular</td>
  </tr>
  <tr>
    <td class=titulo>&nbsp;Ano Conclusão</td>
    <td class=titulo>&nbsp;Instituição</td>
    <td class=titulo>&nbsp;Mês/Ano</td>
    <td class=titulo>&nbsp;Tipo</td>
  </tr>
<%
sql="SELECT u.ANOCONC2GRAU, u.CODCUR, c.NOME, u.COLEGIO2GRAU, u.DATAINGRESSO, u.TIPOINGRESSO, i.TIPOINGRESSO as tipo, u.ENTVESTIBULAR, u.CLASSIFICACAOVESTIB, u.PONTOSVESTIBULAR " & _
"FROM (corporerm.dbo.ualucurso AS u LEFT JOIN corporerm.dbo.utabingrs AS i ON u.TIPOINGRESSO = i.CODTIPING) INNER JOIN corporerm.dbo.UCURSOS AS c ON u.CODCUR = c.CODCUR " & _
"WHERE u.MATALUNO='" & rs("matricula") & "' "
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then 
	anoconclusao=rs2("anoconc2grau")
	instituicao=rs2("colegio2grau")
	ingresso=rs2("dataingresso")
	tipoingresso=rs2("tipo")
	entidade=rs2("entvestibular")
	classificacao=rs2("classificacaovestib")
	pontos=rs2("pontosvestibular")
end if
rs2.close
%>
  <tr>
    <td class=fundo>&nbsp;<input type="text" class=a size=8 value="<%=anoconclusao%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=30 value="<%=instituicao%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=8 value="<%=ingresso%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=30 value="<%=tipoingresso%>" onfocus="this.blur()"></td>
  </tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo>&nbsp;Entidade</td>
    <td class=titulo>&nbsp;Classificação</td>
    <td class=titulo>&nbsp;Pontos</td>
  </tr>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=50 value="<%=entidade%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=4 value="<%=classificacao%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=4 value="<%=pontos%>" onfocus="this.blur()"></td>
  </tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Saúde</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo colspan=3 align="center" style="border:1px solid #000000">&nbsp;Deficiências</td>
  </tr>
<%
sql="SELECT SANGUE, ALERGIAS, MEDICOS, TRATAMENTOS, REMEDIOS, SOCHOSP, OBSSAUDE, DEFICFISICO, DEFICVISUAL, DEFICAUDITIVO, INDDEFFIS, INDDEFVIS, INDDEFAUD " & _
"FROM corporerm.dbo.ealusaude WHERE MATALUNO='" & rs("matricula") & "' "
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then 
	SANGUE=rs2("SANGUE")
	ALERGIAS=rs2("ALERGIAS")
	MEDICOS=rs2("MEDICOS")
	TRATAMENTOS=rs2("TRATAMENTOS")
	REMEDIOS=rs2("REMEDIOS")
	SOCHOSP=rs2("SOCHOSP")
	OBSSAUDE=rs2("OBSSAUDE")
	DEFICFISICO=rs2("DEFICFISICO")
	DEFICVISUAL=rs2("DEFICVISUAL")
	DEFICAUDITIVO=rs2("DEFICAUDITIVO")
	INDDEFFIS=rs2("INDDEFFIS")
	INDDEFVIS=rs2("INDDEFVIS")
	INDDEFAUD =rs2("INDDEFAUD")
end if
rs2.close
%>
  <tr>
    <td class=titulo><input type="checkbox" value="1" <%if deficfisico="S" then response.write "checked"%> onfocus="this.blur()">&nbsp;Física</td>
    <td class=titulo><input type="checkbox" value="1" <%if deficvisual="S" then response.write "checked"%> onfocus="this.blur()">&nbsp;Visual</td>
    <td class=titulo><input type="checkbox" value="1" <%if deficauditivo="S" then response.write "checked"%> onfocus="this.blur()">&nbsp;Auditiva</td>
  </tr>
</table>

<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
	<tr><td class=titulo width=120>Alergias</td><td class=campo>&nbsp;<%=alergias%></td></tr>
	<tr><td class=titulo width=120>Remédios</td>          <td class=campo>&nbsp;<%=remedios%></td></tr>
	<tr><td class=titulo width=120>Médicos</td>           <td class=campo>&nbsp;<%=medicos%></td></tr>
	<tr><td class=titulo width=120>Socorro Hospitalar</td><td class=campo>&nbsp;<%=sochosp%></td></tr>
	<tr><td class=titulo width=120>Tratamento</td>        <td class=campo>&nbsp;<%=tratamentos%></td></tr>
	<tr><td class=titulo width=120>Outros</td>            <td class=campo>&nbsp;<%=obssaude%></td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Orientação</td></tr>
</table>
<%
sqla="SELECT e.nome, u.titulo, u.ano, u.cargahoraria from corporerm.dbo.uprofaluno u, corporerm.dbo.eprofes e " & _
"WHERE e.codprof=u.codprof and u.MATALUNO='" & rs("matricula") & "' "
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
	<th class=titulo colspan=4>Cadastro de Orientações</th>
	<tr>
		<td class=titulor>Professor</td>
		<td class=titulor>Título</td>
		<td class=titulor>Ano</td>
		<td class=titulor>C.Horária</td>
	</tr>
<%
'if ok=1234 then
if rs3.recordcount>0 then
rs3.movefirst
do while not rs3.eof
%>
	<tr>
		<td class="campor" align="left"><%=rs3("nome")%></td>
		<td class="campor" align="left"><%=rs3("titulo")%></td>
		<td class="campor" align="center"><%=rs3("ano")%></td>
		<td class="campor" align="center"><%=rs3("cargahoraria")%></td>
	</tr>
<%
rs3.movenext
loop
end if
rs3.close
%>
</table>

<%
rs.close
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
conexao2.close
set conexao2=nothing
%>
</body>
</html>