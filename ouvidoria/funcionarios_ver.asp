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
<title>Funcionários</title>
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

sqla="select f.*, p.* from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo "
sqlb="and f.chapa='" & request("chapa") & "' "
sqlc="order by f.chapa "
sql1=sqla & sqlb & sqlc
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<p style="margin-top:0;margin-bottom:0" class=titulo>CADASTRO DE FUNCIONÁRIOS</p>
<%
'rs.movefirst:do while not rs.eof 'não há necessidade, é o unico registro
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
tabela=615
tbfoto=150
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Identificação</td></tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td valign="top" class=fundo>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
  <tr>
    <td class=titulo>&nbsp;Código</td>
    <td class=titulo>&nbsp;Chapa</td>
    <td class=titulo>&nbsp;Nome</td>
  </tr>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=4 value="<%=rs("codigo")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" class=a size=5 value="<%=rs("chapa")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" class=a size=45 value="<%=rs("nome")%>" onfocus="this.blur()"></td>
  </tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
  <tr>
    <td class=titulo>&nbsp;Apelido</td>
    <td class=titulo>&nbsp;Data de Nascimento</td>
    <td class=titulo>&nbsp;Sexo</td>
  </tr>
<%
sql="select descricao from corporerm.dbo.pcodsexo where codcliente='" & rs("sexo") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then sexo=rs2("descricao")
rs2.close
%>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=15 value="<%=rs("apelido")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("dtnascimento")%>" onfocus="this.blur()">&nbsp;(<%=int((now()-rs("dtnascimento"))/365.25)%>)</td>
    <td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("sexo")%>" onfocus="this.blur()">
	&nbsp;<input type="text" size=20 value="<%=sexo%>" onfocus="this.blur()"></td>
  </tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela-tbfoto%>">
	<tr>
		<td class=titulo>&nbsp;Nacionalidade</td>
		<td class=titulo>&nbsp;Naturalidade</td>
	</tr>
<%
sql="select descricao from corporerm.dbo.pcodnacao where codcliente='" & rs("nacionalidade") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then nacionalidade=rs2("descricao")
rs2.close
%>
	<tr>
		<td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("nacionalidade")%>" onfocus="this.blur()">
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
sql="select nome from corporerm.dbo.getd where codetd='" & rs("estadonatal") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then estadonatal=trim(rs2("nome"))
rs2.close
sql="select descricao from corporerm.dbo.pcodestcivil where codcliente='" & rs("estadocivil") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then estadocivil=trim(rs2("descricao"))
rs2.close
%>
	<tr>
		<td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("estadonatal")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=30 value="<%=estadonatal%>" onfocus="this.blur()"></td>
		<td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("estadocivil")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=20 value="<%=estadocivil%>" onfocus="this.blur()"></td>
	</tr>
</table>
    </td>
	<td width="<%=tbfoto%>" valign="top" class=fundo>
		<img border="0" src="../func_foto.asp?chapa=<%=rs("chapa")%>"  width="<%=tbfoto%>">
	</td>
	</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
	<tr>
		<td class=titulo>&nbsp;Grau de Instrução</td>
		<td class=titulo valign=top width=300>&nbsp;Nome dos Pais</td>
	</tr>
	
<%
sql="select descricao from corporerm.dbo.pcodinstrucao where codcliente='" & rs("grauinstrucao") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then grauinstrucao=trim(rs2("descricao"))
rs2.close
%>
	<tr>
		<td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("grauinstrucao")%>" onfocus="this.blur()">
		&nbsp;<input type="text" class=a size=30 value="<%=grauinstrucao%>" onfocus="this.blur()"></td>
		<td class=fundo rowspan=3 valign=top>
<%
sql="select nome from corporerm.dbo.pfdepend where grauparentesco='6' and chapa='" & rs("chapa") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then pai=trim(rs2("nome"))
rs2.close
sql="select nome from corporerm.dbo.pfdepend where grauparentesco='7' and chapa='" & rs("chapa") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then mae=trim(rs2("nome"))
rs2.close
%>
		<%=pai%><br><%=mae%>	
		</td>
	</tr>
	<tr>
		<td class=titulo>&nbsp;Email</td>
	</tr>
	<tr>
		<td class=fundo>&nbsp;<input type="text" size=50 value="<%=rs("email")%>" onfocus="this.blur()"></td>
	</tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Endereço Principal</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo>&nbsp;Rua</td>
    <td class=titulo>&nbsp;Número</td>
    <td class=titulo>&nbsp;Complemento</td>
  </tr>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=40 value="<%=rs("rua")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("numero")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=30 value="<%=rs("complemento")%>" onfocus="this.blur()"></td>
  </tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
	<tr>
		<td class=titulo>&nbsp;Bairro</td>
		<td class=titulo>&nbsp;Cidade</td>
		<td class=titulo>&nbsp;Estado</td>
	</tr>
<%
sql="select nome from corporerm.dbo.getd where codetd='" & rs("estado") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then estado=trim(rs2("nome"))
rs2.close
%>
	<tr>
		<td class=fundo>&nbsp;<input type="text" size=30 value="<%=rs("bairro")%>" onfocus="this.blur()"></td>
		<td class=fundo>&nbsp;<input type="text" size=32 value="<%=rs("cidade")%>" onfocus="this.blur()"></td>
		<td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("estado")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=30 value="<%=estado%>" onfocus="this.blur()"></td>
	</tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo>&nbsp;País</td>
    <td class=titulo>&nbsp;CEP</td>
    <td class=titulo>&nbsp;Telefone I</td>
    <td class=titulo>&nbsp;Telefone II</td>
    <td class=titulo>&nbsp;Telefone III</td>
    <td class=titulo>&nbsp;Fax</td>
  </tr>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=16 value="<%=rs("pais")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("cep")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=15 value="<%=rs("telefone1")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=15 value="<%=rs("telefone2")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=15 value="<%=rs("telefone3")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=15 value="<%=rs("fax")%>" onfocus="this.blur()"></td>
  </tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Registro - Admissão</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo>&nbsp;Data de Admissão</td>
    <td class=titulo>&nbsp;Tipo de Admissão</td>
    <td class=titulo>&nbsp;</td>
    <td class=titulo>&nbsp;Motivo da Admissão</td>
  </tr>
<%
sql="select descricao from corporerm.dbo.ptpadmissao where codcliente='" & rs("tipoadmissao") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then tipoadmissao=trim(rs2("descricao"))
rs2.close
%>
  <tr>
    <td class=fundo>&nbsp;<input type="text" class=a size=8 value="<%=rs("dataadmissao")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=3 value="<%=rs("tipoadmissao")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=15 value="<%=tipoadmissao%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;Contrato com prazo <input type="checkbox" value="1" <%if rs("temprazocontr")=1 then response.write "checked"%> onfocus="this.blur()"></td>
<%
sql="select descricao from corporerm.dbo.pmotadmissao where codcliente='" & rs("motivoadmissao") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then motivoadmissao=trim(rs2("descricao"))
rs2.close
%>
    <td class=fundo>&nbsp;<input type="text" size=3 value="<%=rs("motivoadmissao")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=15 value="<%=motivoadmissao%>" onfocus="this.blur()"></td>
  </tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=fundo>&nbsp;Membro Sindical <input type="checkbox" value="1" <%if rs("membrosindical")=1 then response.write "checked"%> onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;Membro da Cipa <input type="checkbox" value="1" <%if rs("membrocipa")=1 then response.write "checked"%> onfocus="this.blur()"></td>
    <td class=fundo width=400>&nbsp;</td>
  </tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo>&nbsp;Sindicato</td>
  </tr>
<%
sql="select nome from corporerm.dbo.psindic where codigo='" & rs("codsindicato") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codsindicato=trim(rs2("nome"))
rs2.close
%>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=10 value="<%=rs("codsindicato")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=70 value="<%=codsindicato%>" onfocus="this.blur()"></td>
  </tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Registro - Admissão II</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=fundo valign=top>
	<!-- função / contribuição sindical -->
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<%
sql="select nome from corporerm.dbo.pfuncao where codigo='" & rs("codfuncao") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codfuncao=trim(rs2("nome"))
rs2.close
sql="select nomefaixa from corporerm.dbo.vfaixasalarial where codfaixa='" & rs("gruposalarial") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then gruposalarial=trim(rs2("nomefaixa"))
rs2.close
sql="select nomenivel from corporerm.dbo.vnivelfuncao where codnivel='" & rs("codnivelsal") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codnivelsal=trim(rs2("nomenivel"))
rs2.close
sql="select descricao from corporerm.dbo.pcodctsind where codcliente='" & rs("contribsindical") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then contribsindical=trim(rs2("descricao"))
rs2.close
%>
  <tr><td class=Titulo>&nbsp;Função</td></tr>
  <tr><td class=fundo><input type="text" size=9 value="<%=rs("codfuncao")%>" onfocus="this.blur()">
		&nbsp;<input type="text" class=a size=25 value="<%=codfuncao%>" onfocus="this.blur()">
		<a class=r href="hstfuncao.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'HistoricoFuncao','550','300','yes','center');return false" onfocus="this.blur()">
		<img src="../images/Form.gif" border="0" width="16" height="16" alt="Histórico de Função"></a>
		</td></tr>
  <tr><td class=Titulo>&nbsp;CAGED</td></tr>
  <tr>
    <td class=fundo>&nbsp;Deficiente:<br>
	&nbsp;Físico   <input type="checkbox" value="1" <%if rs("deficientefisico")=1 then response.write "checked"%> onfocus="this.blur()">
	&nbsp;Auditivo <input type="checkbox" value="1" <%if rs("deficienteauditivo")=1 then response.write "checked"%> onfocus="this.blur()">
	&nbsp;Fala     <input type="checkbox" value="1" <%if rs("deficientefala")=1 then response.write "checked"%> onfocus="this.blur()">
	&nbsp;Visual   <input type="checkbox" value="1" <%if rs("deficientevisual")=1 then response.write "checked"%> onfocus="this.blur()">
	&nbsp;Mental   <input type="checkbox" value="1" <%if rs("deficientemental")=1 then response.write "checked"%> onfocus="this.blur()">
	</td>
  </tr>
</table>
	
	</td>
    <td class=fundo valign=top>
	<!-- seção / rais / caged -->
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
<%
sql="select descricao from corporerm.dbo.psecao where codigo='" & rs("codsecao") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codsecao=trim(rs2("descricao"))
rs2.close
sql="select descricao from corporerm.dbo.pcodsitrais where codcliente='" & rs("situacaorais") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then situacaorais=trim(rs2("descricao"))
rs2.close
sql="select descricao from corporerm.dbo.pcodvinculo where codcliente='" & rs("vinculorais") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then vinculorais=trim(rs2("descricao"))
rs2.close
sql="select descricao from corporerm.dbo.pcorraca where codcliente='" & rs("corraca") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then corraca=trim(rs2("descricao"))
rs2.close
%>
  <tr><td class=Titulo>Seção</td></tr>
  <tr><td class=fundo><input type="text" size=9 value="<%=rs("codsecao")%>" onfocus="this.blur()">
		&nbsp;<input type="text" class=a size=30 value="<%=codsecao%>" onfocus="this.blur()">
		<a class=r href="hstsecao.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'HistoricoSecao','550','300','yes','center');return false" onfocus="this.blur()">
		<img src="../images/Form.gif" border="0" width="16" height="16" alt="Histórico de Seção"></a>
		</td></tr>
  <tr><td class=fundo>Cor/Raça</td></tr>
  <tr><td class=fundo><input type="text" size=3 value="<%=rs("corraca")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=20 value="<%=corraca%>" onfocus="this.blur()"></td></tr>
</table>
	</td>
  </tr>
</table>
<%if rs("codsituacao")="D" then %>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Demissão</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo>&nbsp;Data Demissão</td>
    <td class=titulo>&nbsp;Dt.Desligamento</td>
    <td class=titulo>&nbsp;Dt.Pagamento</td>
    <td class=titulo>&nbsp;Aviso Prévio</td>
    <td class=fundo>&nbsp;Data do aviso</td>
    <td class=fundo>&nbsp;Dias de aviso</td>
  </tr>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("datademissao")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("dtdesligamento")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("dtpagtorescisao")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;Tem aviso indenizado <input type="checkbox" value="1" <%if rs("temavisoprevio")=1 then response.write "checked"%> onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("dtavisoprevio")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=5 value="<%=rs("nrodiasaviso")%>" onfocus="this.blur()"></td>
  </tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo>&nbsp;Tipo de Demissão</td>
    <td class=titulo>&nbsp;Código de Saque</td>
    <td class=titulo>&nbsp;Motivo de demissão</td>
  </tr>
<%
sql="select descricao from corporerm.dbo.ptpdemissao where codcliente='" & rs("tipodemissao") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then tipodemissao=trim(rs2("descricao"))
rs2.close
sql="select descricao from corporerm.dbo.pcodsaque where codcliente='" & rs("codsaquefgts") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codsaquefgts=trim(rs2("descricao"))
rs2.close
sql="select descricao from corporerm.dbo.pmotdemissao where codcliente='" & rs("motivodemissao") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then motivodemissao=trim(rs2("descricao"))
rs2.close
%>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=3 value="<%=rs("tipodemissao")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=20 value="<%=tipodemissao%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=3 value="<%=rs("codsaquefgts")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=20 value="<%=codsaquefgts%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=3 value="<%=rs("motivodemissao")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=20 value="<%=motivodemissao%>" onfocus="this.blur()"></td>
  </tr>
</table>
<%end if%>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Base de Cálculo</td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=titulo>&nbsp;Forma de Recebimento</td>
    <td class=titulo>&nbsp;Tipo de Funcionário</td>
    <td class=titulo>&nbsp;Situação</td>
  </tr>
<%
sql="select descricao from corporerm.dbo.pcodreceb where codcliente='" & rs("codrecebimento") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codrecebimento=trim(rs2("descricao"))
rs2.close
sql="select descricao from corporerm.dbo.ptpfunc where codcliente='" & rs("codtipo") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codtipo=trim(rs2("descricao"))
rs2.close
sql="select descricao from corporerm.dbo.pcodsituacao where codcliente='" & rs("codsituacao") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codsituacao=trim(rs2("descricao"))
rs2.close
%>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=3 value="<%=rs("codrecebimento")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=15 value="<%=codrecebimento%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=3 value="<%=rs("codtipo")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=15 value="<%=codtipo%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=3 value="<%=rs("codsituacao")%>" onfocus="this.blur()">
		&nbsp;<input type="text" class=a size=15 value="<%=codsituacao%>" onfocus="this.blur()">
		<a class=r href="hstsituacao.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'HistoricoSituacao','550','300','yes','center');return false" onfocus="this.blur()">
		<img src="../images/Form.gif" border="0" width="16" height="16" alt="Histórico de Situação"></a>
	</td>
  </tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=titulo>&nbsp;Dependentes</td><td class=titulo>&nbsp;Aposentadoria</td></tr>
<tr><td class=titulo valign=top>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=300>
  <tr>
    <td class=fundo>&nbsp;Nro.Depend.IRRF</td>
    <td class=fundo>&nbsp;Nro.Depend.Sal.Família</td>
    <td class=fundo>&nbsp;</td>
  </tr>
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=5 value="<%=rs("nrodepirrf")%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=5 value="<%=rs("nrodepsalfam")%>" onfocus="this.blur()"></td>
	<td class=fundo><a class=r href="hstdepend.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'HistoricoNroDepend','550','300','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Form.gif" border="0" width="16" height="16" alt="Histórico de Nro.Dependentes"></a>
	</td>
  </tr>
 </table>
</td><td class=titulo valign=top>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td class=fundo>&nbsp;</td>
    <td class=fundo>&nbsp;Data Aposentadoria</td>
  </tr>
  <tr>
    <td class=fundo>&nbsp;Aposentado <input type="checkbox" value="1" <%if rs("aposentado")=1 then response.write "checked"%> onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("dtaposentadoria")%>" onfocus="this.blur()"></td>
  </tr>
 </table>
</td>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=titulo>&nbsp;Horário</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
  <tr>
    <td class=fundo>&nbsp;</td>
    <td class=fundo>&nbsp;Letra</td>
    <td class=fundo>&nbsp;</td>
  </tr>
<%
sql="select descricao from corporerm.dbo.ahorario where codigo='" & rs("codhorario") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codhorario=trim(rs2("descricao"))
rs2.close
sql="select descricao from corporerm.dbo.aindhor where codhorario='" & rs("codhorario") & "' and indiniciohor=" & rs("indiniciohor") & ""
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then letra=trim(rs2("descricao"))
rs2.close
%>  
  <tr>
    <td class=fundo>&nbsp;<input type="text" size=2 value="<%=rs("codhorario")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=75 value="<%=codhorario%>" onfocus="this.blur()"></td>
    <td class=fundo>&nbsp;<input type="text" size=2 value="<%=letra & " (" & rs("indiniciohor")%>)" onfocus="this.blur()"></td>
	<td class=fundo><a class=r href="hsthorario.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'HistoricoHorario','550','300','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Form.gif" border="0" width="16" height="16" alt="Histórico de Horários"></a>
	</td>
  </tr>
</table>


<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
  <tr><td class=grupo>Histórico de Afastamentos</td></tr>
</table>
<%
sqla="SELECT * FROM corporerm.dbo.PFHSTAFT WHERE CHAPA='" & rs("chapa") & "' ORDER BY dtinicio"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
	<th class=titulo colspan=7>Histórico de Afastamentos</th>
	<tr>
		<td class=titulor>Início</td>
		<td class=titulor>Final</td>
		<td class=titulor>Dias</td>
		<td class=titulor>Tipo</td>
		<td class=titulor>Desc.Tipo</td>
		<td class=titulor>Motivo</td>
		<td class=titulor>Desc.Motivo</td>
	</tr>
<%
if rs3.recordcount>0 then
rs3.movefirst
do while not rs3.eof
sql="select descricao from corporerm.dbo.pcodafast where codcliente='" & rs3("tipo") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then desctipo=trim(rs2("descricao")) else desctipo=""
rs2.close
sql="select descricao from corporerm.dbo.pmudsituacao where codcliente='" & rs3("motivo") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then descmotivo=trim(rs2("descricao")) else descmotivo=""
rs2.close
if rs3("dtfinal")="" or isnull(rs3("dtfinal")) then final=now() else final=rs3("dtfinal")
dias=int(final-rs3("dtinicio"))
ano=datediff("yyyy",rs3("dtinicio"),final)
mes=datediff("m",rs3("dtinicio"),final)
dia=datediff("d",rs3("dtinicio"),final)
%>
	<tr>
		<td class="campor" align="left"><%=rs3("dtinicio")%></td>
		<td class="campor" align="left"><%=rs3("dtfinal")%></td>
		<td class="campor" align="right"><%=dias%>&nbsp;</td>
		<td class="campor" align="left"><%=rs3("tipo")%></td>
		<td class="campor" align="left"><%=desctipo%></td>
		<td class="campor" align="left"><%=rs3("motivo")%></td>
		<td class="campor" align="left"><%=descmotivo%></td>
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
'rs2.close
set rs2=nothing
conexao.close
set conexao=nothing
conexao2.close
set conexao2=nothing
%>
</body>
</html>