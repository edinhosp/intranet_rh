<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a35")="N" or session("a35")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
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


if request("nomeacao")=1 then nomeacao=1 else nomeacao=0
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
set rs4=server.createobject ("ADODB.Recordset")
Set rs4.ActiveConnection = conexao
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
	<img border="0" src="func_foto.asp?chapa=<%=rs("chapa")%>"  width="<%=tbfoto%>">
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
	<td class=fundo>&nbsp;<input type="text" size=40 value="<%=rs("rua")%>" onfocus="this.blur()">
	<%googlemap=rs("rua") & ", " & rs("numero") & ", " & rs("cidade") & ", " & rs("cep")%>
		<a class=r href="http://maps.google.com.br/maps?f=q&source=s_q&hl=pt-BR&geocode=&q=<%=googlemap%>" target="_blank">
		<img src="../images/Form.gif" border="0" width="16" height="16" alt="Mapa do Endereço"></a>
	</td>
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
<tr><td class=grupo>Documentos</td></tr>
</table>

<table border="0" bordercolor="#808080" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr>
	<td class=titulo>&nbsp;Carteira de Identidade</td>
	<td class=titulo>&nbsp;Título de Eleitor</td>
	<td class=titulo>&nbsp;Carteira de Trabalho</td>
	<td class=titulo>&nbsp;Carteira de Motorista</td>
</tr>
<%
sql="select nome from corporerm.dbo.getd where codetd='" & rs("ufcartident") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then ufcartident=trim(rs2("nome"))
rs2.close
sql="select nome from corporerm.dbo.getd where codetd='" & rs("ufcarttrab") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then ufcarttrab=trim(rs2("nome"))
rs2.close
cpf=rs("cpf")
if cpf<>"" then cpf=left(cpf,3) & "." & mid(cpf,4,3) & "." & mid(cpf,7,3) & "-" & right(cpf,2)
sql="select nome from corporerm.dbo.gbanco where numbanco='" & rs("codbancopis") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then codbancopis=trim(rs2("nome"))
rs2.close
%>
<tr>
	<td class=fundo valign=top>&nbsp;Número<br>&nbsp;<input type="text" size=15 value="<%=rs("cartidentidade")%>" onfocus="this.blur()">
	<br>&nbsp;Data de Emissão<br>&nbsp;<input type="text" size=8 value="<%=rs("dtemissaoident")%>" onfocus="this.blur()">
	<br>&nbsp;Orgão Emissor<br>&nbsp;<input type="text" size=16 value="<%=rs("orgemissorident")%>" onfocus="this.blur()">
	<br>&nbsp;Estado Emissor<br>&nbsp;<input type="text" size=2 value="<%=rs("ufcartident")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=15 value="<%=ufcartident%>" onfocus="this.blur()">
	</td>
	<td class=fundo valign=top>&nbsp;Número<br>&nbsp;<input type="text" size=14 value="<%=rs("tituloeleitor")%>" onfocus="this.blur()">
	<br>&nbsp;Zona<br>&nbsp;<input type="text" size=6 value="<%=rs("zonatiteleitor")%>" onfocus="this.blur()">
	<br>&nbsp;Seção<br>&nbsp;<input type="text" size=6 value="<%=rs("secaotiteleitor")%>" onfocus="this.blur()">
	</td>
	<td class=fundo valign=top>&nbsp;Número<br>&nbsp;<input type="text" class=a size=10 value="<%=rs("carteiratrab")%>" onfocus="this.blur()">
	<br>&nbsp;Série<br>&nbsp;<input type="text" size=5 value="<%=rs("seriecarttrab")%>" onfocus="this.blur()">
	Carteira tipo NIT <input type="checkbox" value="1" <%if rs("nit")=1 then response.write "checked"%>>
	<br>&nbsp;Data de Emissão<br>&nbsp;<input type="text" size=8 value="<%=rs("dtcarttrab")%>" onfocus="this.blur()">
	<br>&nbsp;Estado Emissor<br>&nbsp;<input type="text" size=2 value="<%=rs("ufcarttrab")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=15 value="<%=ufcarttrab%>" onfocus="this.blur()">
	</td>
	<td class=fundo valign=top>&nbsp;Número<br>&nbsp;<input type="text" size=15 value="<%=rs("cartmotorista")%>" onfocus="this.blur()">
	<br>&nbsp;Tipo da Habilitação<br>&nbsp;<input type="text" size=5 value="<%=rs("tipocarthabilit")%>" onfocus="this.blur()">
	<br>&nbsp;Data de Vencimento<br>&nbsp;<input type="text" size=8 value="<%=rs("dtvenchabilit")%>" onfocus="this.blur()">
	</td>
</tr>
<tr>
	<td class=titulo>&nbsp;CPF</td>
	<td class=titulo>&nbsp;Registro Profissional</td>
	<td class=titulo>&nbsp;Carteira Reservista</td>
	<td class=titulo>&nbsp;PIS/PASEP</td>
</tr>
<tr>
	<td class=fundo valign=top>&nbsp;<input type="text" class=a size=14 value="<%=cpf%>" onfocus="this.blur()"></td>
	<td class=fundo valign=top>&nbsp;<input type="text" size=15 value="<%=rs("regprofissional")%>" onfocus="this.blur()"></td>
	<td class=fundo valign=top>&nbsp;Número<br>&nbsp;<input type="text" size=20 value="<%=rs("certifreserv")%>" onfocus="this.blur()">
	<br>&nbsp;Categoria<br>&nbsp;<input type="text" size=10 value="<%=rs("categmilitar")%>" onfocus="this.blur()">
	</td>
	<td class=fundo valign=top>&nbsp;<input type="text" size=14 value="<%=rs("pispasep")%>" onfocus="this.blur()">
	<br>&nbsp;Banco<br>&nbsp;<input type="text" size=2 value="<%=rs("codbancopis")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=15 value="<%=codbancopis%>" onfocus="this.blur()">
	<br>&nbsp;Data de Cadastramento<br>&nbsp;<input type="text" size=8 value="<%=rs("dtcadastropis")%>" onfocus="this.blur()">
	</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Campos Complementares</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr>
	<td class=titulo>&nbsp;Descrição</td>
	<td class=titulo>&nbsp;Valor do Campo</td>
</tr>
<%
sql="select * from corporerm.dbo.pfcompl where chapa='" & rs("chapa") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then
	ge=rs2("ge")
	remuneracao=rs2("remuneracao")
	admissao=rs2("admissao")
	brigada=rs2("brigada")
	assmedica=rs2("assmedica")
	assmedcod=rs2("assmedcod")
	setoraloc=rs2("setor")
	titulacaopaga=rs2("titulacaopaga")
end if
rs2.close
sql="select descricao from corporerm.dbo.gconsist where codcliente='" & assmedica & "' and codtabela='ASSMED'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then assmedicad=rs2("descricao")
rs2.close
sql="select descricao from corporerm.dbo.gconsist where codcliente='" & remuneracao & "' and codtabela='07'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then remuneracaod=rs2("descricao")
rs2.close
sql="select descricao from corporerm.dbo.psecao where codigo='" & setoraloc & "' "
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then setoralocnome=rs2("descricao")
rs2.close
sql="select descricao from corporerm.dbo.pcodinstrucao where codcliente='" & titulacaopaga & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then instrucaomec=trim(rs2("descricao"))
rs2.close

%>
<tr>
	<td class=fundo>&nbsp;Titulação para Pagamento/MEC</td>
	<td class=fundo>&nbsp;<%=titulacaopaga%> - <%=instrucaomec%></td>
</tr>
<tr>
	<td class=fundo>&nbsp;Setor / C.Custo de alocação</td>
	<td class=fundo>&nbsp;<%=setoraloc%> - <%=setoralocnome%></td>
</tr>
<tr>
	<td class=fundo>&nbsp;Admissão p/Prêmio de Permanência</td>
	<td class=fundo>&nbsp;<%=admissao%></td>
</tr>
<tr>
	<td class=fundo>&nbsp;Assistência Médica</td>
	<td class=fundo>&nbsp;<%=assmedica%> - <%=assmedicad%></td>
</tr>
<tr>
	<td class=fundo>&nbsp;Assistência Médica - Código</td>
	<td class=fundo>&nbsp;<%=assmedcod%></td>
</tr>
<tr>
	<td class=fundo>&nbsp;Grupo Especial</td>
	<td class=fundo>&nbsp;<%=ge%></td>
</tr>
<tr>
	<td class=fundo>&nbsp;Integrante da Brigada Incêndio</td>
	<td class=fundo>&nbsp;<%=brigada%></td>
</tr>
<tr>
	<td class=fundo>&nbsp;Remuneração</td>
	<td class=fundo>&nbsp;<%=remuneracao%> - <%=remuneracaod%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Registro - Admissão</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr>
	<td class=titulo>&nbsp;Data de Admissão</td>
	<td class=titulo>&nbsp;Tipo de Admissão</td>
	<td class=titulo>&nbsp;Data Base</td>
	<td class=titulo>&nbsp;</td>
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
	<td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("dtbase")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;Contrato com prazo <input type="checkbox" value="1" <%if rs("temprazocontr")=1 then response.write "checked"%> onfocus="this.blur()"></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr>
	<td class=titulo>&nbsp;Num.Ficha de Registro</td>
	<td class=titulo>&nbsp;Motivo da Admissão</td>
	<td class=titulo>&nbsp;Data da Transferência</td>
	<td class=titulo>&nbsp;Data do Final do Contrato</td>
</tr>
<%
sql="select descricao from corporerm.dbo.pmotadmissao where codcliente='" & rs("motivoadmissao") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then motivoadmissao=trim(rs2("descricao"))
rs2.close
%>
<tr>
	<td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("nrofichareg")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=3 value="<%=rs("motivoadmissao")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=15 value="<%=motivoadmissao%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("dttransferencia")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=8 value="<%=rs("fimprazocontr")%>" onfocus="this.blur()"></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr>
	<td class=titulo>&nbsp;Banco Pagamento</td>
	<td class=titulo>&nbsp;RMLabore.net</td>
</tr>
<%
sql="select nome from corporerm.dbo.gbanco where numbanco='" & rs("codbancopagto") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then codbancopagto=trim(rs2("nome"))
rs2.close
sql="select nome from corporerm.dbo.gagencia where numagencia='" & rs("codagenciapagto") & "' and numbanco='" & rs("codbancopagto") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then codagenciapagto=trim(rs2("nome"))
rs2.close
sql="select nome from corporerm.dbo.gusuario where codusuario='" & rs("codusuario") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then codusuario=trim(rs2("nome"))
rs2.close
sql="select descricao from corporerm.dbo.pgrpacessoquiosque where codgrupo='" & rs("codgrpquiosque") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then codgrpquiosque=trim(rs2("descricao"))
rs2.close
%>
<tr>
	<td class=fundo valign=top>&nbsp;<input type="text" size=3 value="<%=rs("codbancopagto")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=15 value="<%=codbancopagto%>" onfocus="this.blur()">	
		<br>&nbsp;Agência de Pagamento<br>&nbsp;<input type="text" size=6 value="<%=rs("codagenciapagto")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=25 value="<%=codagenciapagto%>" onfocus="this.blur()">
		<br>&nbsp;Número Conta Pagamento    Operação Bancária<br>&nbsp;<input type="text" size=15 value="<%=rs("contapagamento")%>" onfocus="this.blur()">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" size=10 value="<%=rs("opbancaria")%>" onfocus="this.blur()">
	</td>
	<td class=fundo valign=top>&nbsp;Grupo de acesso ao RMLabore.net<br>&nbsp;<input type="text" size=15 value="<%=rs("codgrpquiosque")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=15 value="<%=codgrpquiosque%>" onfocus="this.blur()">
		<br>&nbsp;Cód. Equipe<br>&nbsp;<input type="text" size=4 value="<%=rs("codequipe")%>" onfocus="this.blur()">
		<br>&nbsp;Usuário<br>&nbsp;<input type="text" size=15 value="<%=rs("codusuario")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=15 value="<%=codusuario%>" onfocus="this.blur()">	
	</td>
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
<tr><td class=titulo>&nbsp;Sindicato</td></tr>
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
	<tr><td class=fundo>&nbsp;Faixa Salarial</td></tr>
	<tr><td class=fundo><input type="text" size=10 value="<%=rs("gruposalarial")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=20 value="<%=gruposalarial%>" onfocus="this.blur()"></td></tr>
	<tr><td class=fundo>&nbsp;Nível Salarial</td></tr>
	<tr><td class=fundo><input type="text" size=10 value="<%=rs("codnivelsal")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=20 value="<%=codnivelsal%>" onfocus="this.blur()"></td></tr>
	<tr><td class=Titulo>&nbsp;Contribuição Sindical</td></tr>
	<tr><td class=fundo><input type="text" size=3 value="<%=rs("contribsindical")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=20 value="<%=contribsindical%>" onfocus="this.blur()">
		<a class=r href="hstcsind.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'HistoricoCSind','550','300','yes','center');return false" onfocus="this.blur()">
		<img src="../images/Form.gif" border="0" width="16" height="16" alt="Histórico de Contribuição Sindical"></a>
		</td></tr>
	<tr><td class=Titulo>&nbsp;CAGED</td></tr>
	<tr><td class=fundo>&nbsp;Deficiente:<br>
	&nbsp;Físico   <input type="checkbox" value="1" <%if rs("deficientefisico")=1 then response.write "checked"%> onfocus="this.blur()">
	&nbsp;Auditivo <input type="checkbox" value="1" <%if rs("deficienteauditivo")=1 then response.write "checked"%> onfocus="this.blur()">
	&nbsp;Fala     <input type="checkbox" value="1" <%if rs("deficientefala")=1 then response.write "checked"%> onfocus="this.blur()">
	&nbsp;Visual   <input type="checkbox" value="1" <%if rs("deficientevisual")=1 then response.write "checked"%> onfocus="this.blur()">
	&nbsp;Mental   <input type="checkbox" value="1" <%if rs("deficientemental")=1 then response.write "checked"%> onfocus="this.blur()">
	
	
	</td></tr>
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
	<tr><td class=fundo><b>Rais</b><br>Situação</td></tr>
	<tr><td class=fundo><input type="text" size=5 value="<%=rs("situacaorais")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=40 value="<%=situacaorais%>" onfocus="this.blur()"></td></tr>
	<tr><td class=fundo>Vínculo</td></tr>
	<tr><td class=fundo><input type="text" size=5 value="<%=rs("vinculorais")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=40 value="<%=vinculorais%>" onfocus="this.blur()"></td></tr>
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
<tr><td class=grupo>Vale Transporte</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr>
	<td class=fundo>&nbsp;</td>
	<td class=fundo>&nbsp;Mês Atual</td>
	<td class=fundo>&nbsp;Após Reajuste</td>
	<td class=fundo>&nbsp;Próximo Mês</td>
</tr>
<tr>
	<td class=fundo>&nbsp;Expediente Integral</td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("diasuteismes")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("diasutrestantes")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("diasutproxmes")%>" onfocus="this.blur()"></td>
</tr>
<tr>
	<td class=fundo>&nbsp;Meio Expediente</td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("diasutmeioexp")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("diasutrestmeio")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("diasutproxmeio")%>" onfocus="this.blur()"></td>
</tr>
</table>
<%
sqla="SELECT f.*, l.codtarifa FROM corporerm.dbo.PFVALETR f, corporerm.dbo.PVALETR L " & _
"WHERE CHAPA='" & rs("chapa") & "' AND F.CODlinha=L.CODIGO ORDER BY codlinha"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<th class=titulo colspan=7>Cadastro de Vale-Transporte</th>
<tr>
	<td class=titulor>Nro.Viagens</td>
	<td class=titulor>Nro.Viagens Meio Exp.</td>
	<td class=titulor>Cod.Linha</td>
	<td class=titulor>Nome da Linha</td>
	<td class=titulor>Valor</td>
	<td class=titulor>Início de uso</td>
	<td class=titulor>Término de uso</td>
</tr>
<%
rs3.movefirst
do while not rs3.eof
sql="select nomelinha from corporerm.dbo.pvaletr where codigo='" & rs3("codlinha") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then nomelinha=trim(rs2("nomelinha")) else nomelinha=""
rs2.close
sql="select valor from corporerm.dbo.ptarifa where codigo='" & rs3("codtarifa") & "' and getdate() between iniciovigencia and finalvigencia "
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then valortar=rs2("valor") else valortar=0
rs2.close
%>
<tr>
	<td class="campor" align="center"><%=rs3("nroviagens")%></td>
	<td class="campor" align="left"><%=rs3("nroviagmeioexp")%></td>
	<td class="campor" align="left"><%=rs3("codlinha")%></td>
	<td class="campor" align="left"><%=nomelinha%></td>
	<td class="campor" align="right"><%=formatnumber(valortar,2)%>&nbsp;</td>
	<td class="campor" align="center"><%=rs3("dtinicio")%></td>
	<td class="campor" align="center"><%=rs3("dtfim")%></td>
</tr>
<%
rs3.movenext
loop
%>
</table>
<%
end if
rs3.close
%>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
 <tr><td class=grupo>Registro - FGTS/SEFIP/INSS</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=titulo colspan=3>&nbsp;FGTS</td></tr>
<tr>
	<td class=fundo>&nbsp;Situação</td>
	<td class=fundo>&nbsp;Data Opção FGTS</td>
	<td class=fundo>&nbsp;Nro. Conta</td>
</tr>
<%
sql="select descricao from corporerm.dbo.pcodsitfgts where codcliente='" & rs("situacaofgts") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then situacaofgts=trim(rs2("descricao"))
rs2.close
sql="select nome from corporerm.dbo.gbanco where numbanco='" & rs("codbancofgts") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codbancofgts=trim(rs2("nome"))
rs2.close
saldofgts=rs("saldofgts")
if saldofgts<>"" then saldofgts=formatnumber(saldofgts,2) else saldofgts=""
%>
<tr>
	<td class=fundo>&nbsp;<input type="text" size=5 value="<%=rs("situacaofgts")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=30 value="<%=situacaofgts%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("dtopcaofgts")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=15 value="<%=rs("contafgts")%>" onfocus="this.blur()"></td>
</tr>
<tr>
	<td class=fundo>&nbsp;Banco</td>
	<td class=fundo>&nbsp;Data último saldo FGTS</td>
	<td class=fundo>&nbsp;Saldo</td>
</tr>
<tr>
	<td class=fundo>&nbsp;<input type="text" size=5 value="<%=rs("codbancofgts")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=30 value="<%=codbancofgts%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("dtsaldofgts")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=15 value="<%=saldofgts%>" onfocus="this.blur()"></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=titulo valign=top>&nbsp;SEFIP</td><td class=titulo valign=top>&nbsp;INSS</td></tr>
<%
sql="select descricao from corporerm.dbo.pcodocortrab where codcliente='" & rs("codocorrencia") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codocorrencia=trim(rs2("descricao"))
rs2.close
sql="select descricao from corporerm.dbo.pcodcattrab where codcliente='" & rs("codcategoria") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then codcategoria=trim(rs2("descricao"))
rs2.close
%>
<tr>
	<td class=fundo valign=top>&nbsp;Ocorrência<br>&nbsp;<input type="text" size=5 value="<%=rs("codocorrencia")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=50 value="<%=codocorrencia%>" onfocus="this.blur()">
		<a class=r href="hstsefip.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'HistoricoSefip','550','300','yes','center');return false" onfocus="this.blur()">
		<img src="../images/Form.gif" border="0" width="16" height="16" alt="Histórico de Ocorrências Sefip"></a>
		<br>&nbsp;Categoria<br>&nbsp;<input type="text" size=5 value="<%=rs("codcategoria")%>" onfocus="this.blur()">
		&nbsp;<input type="text" size=50 value="<%=codcategoria%>" onfocus="this.blur()">
	</td>
	<td class=fundo valign=top><input type="radio" name="s" value="optante" <%if rs("situacaoinss")=1 then response.write "checked"%>> Optante
	<br><br><input type="radio" name="s" value="não optante" <%if rs("situacaoinss")=0 then response.write "checked"%>> Não optante
	</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr>
	<td class=fundo>Histórico de Provisões
	<a class=r href="hstprovisao.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'HistoricoProvisao','550','300','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Form.gif" border="0" width="16" height="16" alt="Histórico de Provisões"></a>
	</td>
	<td class=fundo>Histórico de Exames
	<a class=r href="hstexames.asp?chapa=<%=rs("codpessoa")%>" onclick="NewWindow(this.href,'HistoricoExames','550','300','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Form.gif" border="0" width="16" height="16" alt="Histórico de Exames Médicos"></a>
	</td>
</tr>
</table>

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

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr>
	<td class=fundo>&nbsp;% Adiantam.</td>
	<td class=fundo>&nbsp;Ajuda de Custo</td>
	<td class=fundo>&nbsp;Arredondamento</td>
	<td class=fundo>&nbsp;Média Sal.Matern.</td>
	<td class=fundo>&nbsp;</td>
	<td class=fundo>&nbsp;</td>
	<td class=fundo>&nbsp;</td>
</tr>
<tr>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("percentadiant")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("ajudacusto")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("arredondamento")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rs("mediasalmatern")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;</td>
	<td class=fundo>&nbsp;</td>
	<td class=fundo>&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr>
	<td class=fundo>&nbsp;Deduz IRRF se maior de 65 anos <input type="checkbox" value="1" <%if rs("deduzirrf65")="1" then response.write "checked"%> onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;Tem dedução de CPMF <input type="checkbox" value="1" <%if rs("temdeducaocpmf")=1 then response.write "checked"%> onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;Utiliza Controle de Saldo de Verbas <input type="checkbox" value="1" <%if rs("usacontroledesaldo")=1 then response.write "checked"%> onfocus="this.blur()"></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=titulo>&nbsp;Salário</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr>
	<td class=fundo>&nbsp;</td>
	<td class=fundo>&nbsp;Mensal</td>
	<td class=fundo>&nbsp;Hora</td>
	<td class=fundo>&nbsp;Jornada</td>
	<td class=fundo>&nbsp;</td>
</tr>
<%
jornadamensal=formatdatetime((rs("jornadamensal")/60)/1,4)
jornadames=cdbl(rs("jornadamensal")/60)
horac=int(rs("jornadamensal")/60)
minutoc=int(((rs("jornadamensal")/60)-horac)*60)
jornadamensal=horac & ":" & numzero(minutoc,2)
if rs("usasalcomposto")=1 then hora="" else hora=formatnumber(cdbl(rs("salario"))/jornadames,2)
%>  
<tr>
	<td class=fundo>&nbsp;Usa Salário Composto <input type="checkbox" value="1" <%if rs("usasalcomposto")=1 then response.write "checked"%> onfocus="this.blur()"></td>
<%
salariomostra=formatnumber(rs("salario"),2)
if session("usuariomaster")="02589" then salariomostra="-"
if session("usuariomaster")="02589" then hora="-"
%>
	<td class=fundo>&nbsp;<input type="text" class=a size=15 value="<%=salariomostra%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=10 value="<%=hora%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" class=a size=10 value="<%=jornadamensal%>" onfocus="this.blur()"></td>
	<td class=fundo>
<%if session("usuariomaster")<>"02589" then%>
<a class=r href="hstsalario.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'HistoricoSalario','550','300','yes','center');return false" onfocus="this.blur()">
<img src="../images/Form.gif" border="0" width="16" height="16" alt="Histórico Salarial"></a><%end if%>
	</td>
</tr>
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

<!-- SALARIO COMPOSTO -->
<%if rs("usasalcomposto")=1 then %>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Salário Composto</td></tr>
</table>
<%
sqla="SELECT * FROM corporerm.dbo.PFSALCMP " & _
"WHERE CHAPA='" & rs("chapa") & "' ORDER BY nrosalario"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<th class=titulo colspan=10>Salário Composto</th>
<tr>
	<td class=titulor>Evento</td>
	<td class=titulor>Descricao do Evento</td>
	<td class=titulor>Valor</td>
	<td class=titulor>Jornada</td>
	<td class=titulor>Nro</td>
	<td class=titulor>Centro de Custo</td>
	<td class=titulor>Hora</td>
	<td class=titulor>Início</td>
	<td class=titulor>Término</td>
</tr>
<%
if rs3.recordcount>0 then 
rs3.movefirst
do while not rs3.eof
sql="select descricao from corporerm.dbo.pevento where codigo='" & rs3("codevento") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then evento=trim(rs2("descricao")) else evento=""
rs2.close
jornadames=cdbl(rs3("jornada")/60)
horac=int(rs3("jornada")/60)
minutoc=int(((rs3("jornada")/60)-horac)*60)
jornadamensal=horac & ":" & numzero(minutoc,2)
if cdbl(rs3("valor"))>0 then hora=formatnumber(cdbl(rs3("valor"))/jornadames,2) else hora=0
%>
<tr>
	<td class="campor" align="center"><%=rs3("codevento")%></td>
	<td class="campor" align="left"><%=evento%></td>
	<td class="campor" align="right"><%=formatnumber(rs3("valor"),2)%>&nbsp;&nbsp;</td>
	<td class="campor" align="center"><%=jornadamensal%></td>
	<td class="campor" align="center"><%=rs3("nrosalario")%></td>
	<td class="campor" align="right"><%=rs3("codccusto")%></td>
	<td class="campor" align="right"><%=hora%></td>
	<td class="campor" align="right"><%=rs3("iniciovigencia")%></td>
	<td class="campor" align="center"><%=rs3("fimvigencia")%></td>
</tr>
<%
rs3.movenext
loop
end if
rs3.close
%>
</table>
<%end if 'salario composto %>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Dados Contábeis</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tabela%>">
<tr>
	<td class=titulo colspan=2>&nbsp;Integração Contábil</td>
	<td class=titulo colspan=2>&nbsp;Integração Gerencial</td>
</tr>
<%
sql="select descricao from corporerm.dbo.pcodintcontfunc where codcliente='" & rs("integrcontabil") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then integrcontabil=trim(rs2("descricao")) else evento=""
rs2.close
sql="select descricao from corporerm.dbo.pcodintgerfunc where codcliente='" & rs("integrgerencial") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then integrgerencial=trim(rs2("descricao")) else evento=""
rs2.close
%>
<tr>
	<td class=fundo>&nbsp;<input type="text" size=10 value="<%=rs("integrcontabil")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=25 value="<%=integrcontabil%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=10 value="<%=rs("integrgerencial")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=25 value="<%=integrgerencial%>" onfocus="this.blur()"></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Dependentes</td></tr>
</table>
<%
sqla="SELECT * FROM corporerm.dbo.PFDEPEND " & _
"WHERE CHAPA='" & rs("chapa") & "' ORDER BY nrodepend"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<th class=titulo colspan=8>Cadastro de Dependentes</th>
<tr>
	<td class=titulor>Nro.Dep.</td>
	<td class=titulor>Nome</td>
	<td class=titulor>Parentesco</td>
	<td class=titulor>CPF</td>
	<td class=titulor>Nascimento</td>
	<td class=titulor>Sexo</td>
	<td class=titulor>Estado Civil</td>
	<td class=titulor>Inc.IRRF</td>
</tr>
<%
rs3.movefirst
do while not rs3.eof
sql="select descricao from corporerm.dbo.pcodparent where codcliente='" & rs3("grauparentesco") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then parentesco=trim(rs2("descricao")) else parentesco=""
rs2.close
%>
<tr>
	<td class="campor" align="center"><%=rs3("nrodepend")%></td>
	<td class="campor" align="left"><%=rs3("nome")%></td>
	<td class="campor" align="left"><%=rs3("grauparentesco") & " - " & parentesco%></td>
	<td class="campor" align="center"><%=rs3("cpf")%></td>
	<td class="campor" align="center"><%=rs3("dtnascimento")%></td>
	<td class="campor" align="center"><%=rs3("sexo")%></td>
	<td class="campor" align="center"><%=rs3("estadocivil")%></td>
	<td class="campor" align="center"><%if rs3("incirrf")=1 then response.write "<font face='Wingdings'>ü</font>" %></td>
</tr>
<%
rs3.movenext
loop
%>
</table>
<%
end if
rs3.close
%>

<%
sqla="SELECT * FROM corporerm.dbo.PFCODFIX " & _
"WHERE CHAPA='" & rs("chapa") & "' ORDER BY codevento"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Códigos Fixos</td></tr>
</table>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<th class=titulo colspan=6>Códigos Fixos</th>
<tr>
	<td class=titulor>Evento</td>
	<td class=titulor>Descrição</td>
	<td class=titulor>Valor</td>
	<td class=titulor>Nro.Vezes</td>
	<td class=titulor>Tipo</td>
	<td class=titulor>Descr.Tipo</td>
</tr>
<%
rs3.movefirst
do while not rs3.eof
sql="select descricao, valhordiaref from corporerm.dbo.pevento where codigo='" & rs3("codevento") & "'"
rs2.open sql, ,adOpenStatic
if rs2.recordcount>0 then
	codevento=trim(rs2("descricao"))
	tpe=rs2("valhordiaref")
end if
rs2.close
sql="select descricao from corporerm.dbo.ptpcodfixo where codcliente='" & rs3("tipo") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then desctipo=trim(rs2("descricao")) else desctipo=""
rs2.close
valor=rs3("valor"):if valor="" or isnull(valor) then valor=0
if tpe="V" then valor=formatnumber(valor,2)
if tpe="H" then 
	horac=int(cdbl(valor)/60)
	minutoc=int(((cdbl(valor)/60)-horac)*60)
	valor=horac & ":" & numzero(minutoc,2)
end if
%>
<tr>
	<td class="campor" align="center"><%=rs3("codevento")%></td>
	<td class="campor" align="left"><%=codevento%></td>
	<td class="campor" align="right"><%=valor%>&nbsp;</td>
	<td class="campor" align="center"><%=rs3("nrovezes")%></td>
	<td class="campor" align="center"><%=rs3("tipo")%></td>
	<td class="campor" align="left"><%=desctipo%></td>
</tr>
<%
rs3.movenext
loop
%>
</table>
<%
end if
rs3.close
%>

<%
sqla="SELECT * FROM corporerm.dbo.PFrateiofixo " & _
"WHERE CHAPA='" & rs("chapa") & "' ORDER BY codccusto"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Rateios Fixos</td></tr>
</table>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<th class=titulo colspan=3>Rateios Fixos</th>
<tr>
	<td class=titulor>Centro de Custo</td>
	<td class=titulor>Nome Centro Custo</td>
	<td class=titulor>Percentual/Valor</td>
</tr>
<%
rs3.movefirst
do while not rs3.eof
sql="select nome from corporerm.dbo.pccusto where codccusto='" & rs3("codccusto") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then ccusto=rs2("nome") else ccusto=""
rs2.close
%>
<tr>
	<td class="campor" align="left"><%=rs3("codccusto")%></td>
	<td class="campor" align="left"><%=ccusto%></td>
	<td class="campor" align="right"><%=rs3("valor")%>&nbsp;</td>
</tr>
<%
rs3.movenext
loop
%>
</table>
<%
end if
rs3.close
%>

<%
sqla="select e.CHAPA, e.CODIGO, e.DTEMPRESTIMO, e.VALORORIGINAL, e.NROPARCELAS, e.NROPARCPAGAS, e.SALDODEVEDOR " & _
"from corporerm.dbo.PFEMPRT e where e.CHAPA='" & rs("chapa") & "' order by e.dtemprestimo "
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Empréstimos</td></tr>
</table>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<th class=titulo colspan=7>Empréstimos</th>
<tr>
	<td class=titulor>Código</td>
	<td class=titulor align="center">Data</td>
	<td class=titulor align="center">Valor Total</td>
	<td class=titulor align="center">Nº de Parcelas</td>
	<td class=titulor align="center">Parcelas Pagas</td>
	<td class=titulor align="center">Saldo Devedor</td>
	<td class=titulor>Histórico</td>
</tr>
<%
rs3.movefirst
do while not rs3.eof
'sql="select nome from corporerm.dbo.pccusto where codccusto='" & rs3("codccusto") & "'"
'rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then ccusto=rs2("nome") else ccusto=""
'rs2.close
%>
<tr>
	<td class="campor" align="left"><%=rs3("codigo")%></td>
	<td class="campor" align="center"><%=rs3("dtemprestimo")%></td>
	<td class="campor" align="center"><%=formatnumber(rs3("valororiginal"),2)%></td>
	<td class="campor" align="center"><%=rs3("nroparcelas")%></td>
	<td class="campor" align="center"><%=rs3("nroparcpagas")%></td>
	<td class="campor" align="center"><%=formatnumber(rs3("saldodevedor"),2)%></td>
	<td class="campor" align="left">
	<a class=r href="hstemprestimos.asp?chapa=<%=rs("chapa")%>&codigo=<%=rs3("codigo")%>" onclick="NewWindow(this.href,'HistoricoEmprestimo','550','300','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Form.gif" border="0" width="16" height="16" alt="Histórico de Empréstimos"></a>
	</td>
</tr>
<%
rs3.movenext
loop
%>
</table>
<%
end if
rs3.close
%>



<%
sqla="SELECT * FROM corporerm.dbo.PFHSTAFT " & _
"WHERE CHAPA='" & rs("chapa") & "' ORDER BY dtinicio"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Histórico de Afastamentos</td></tr>
</table>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
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
%>
</table>
<%
end if
rs3.close
%>

<%
sqla="SELECT * FROM corporerm.dbo.PANOTAC " & _
"WHERE CODPESSOA=" & rs("codpessoa") & " ORDER BY tipo, nroanotacao"
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Anotações Pessoais</td></tr>
</table>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<th class=titulo colspan=7>Anotações Pessoais</th>
<tr>
	<td class=titulor>Tipo</td>
	<td class=titulor>Cod.</td>
	<td class=titulor>Data</td>
	<td class=titulor>Texto</td>
</tr>
<%
rs3.movefirst
do while not rs3.eof
sql="select descricao from corporerm.dbo.ptpanotacao where codcliente='" & rs3("tipo") & "'"
rs2.open sql, ,adOpenStatic:if rs2.recordcount>0 then tipo=trim(rs2("descricao")) else tipo=""
rs2.close
%>
<tr>
	<td class="campor" align="left"><%=rs3("tipo") & "-" & tipo%></td>
	<td class="campor" align="left"><%=rs3("nroanotacao")%></td>
	<td class="campor" align="left"><%=rs3("dtanotacao")%></td>
	<td class="campor" align="left"><%=rs3("texto")%></td>
</tr>
<%
rs3.movenext
loop
%>
</table>
<%
end if
rs3.close
%>

<%
sqla="SELECT * FROM emprestimos where chapa='" & rs("chapa") & "' ORDER BY data "
rs4.Open sqla, ,adOpenStatic, adLockReadOnly
if rs4.recordcount>0 then
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<tr><td class=grupo>Empréstimos Consignados</td></tr>
</table>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=<%=tabela%>>
<th class=titulo colspan=7>Empréstimos Consignados</th>
<tr>
	<td class=titulor>Data</td>
	<td class=titulor>Contrato</td>
	<td class=titulor>Valor</td>
	<td class=titulor>NºParc.</td>
	<td class=titulor>Vr Parc.</td>
	<td class=titulor>Início</td>
	<td class=titulor>Término</td>
</tr>
<%
rs4.movefirst
do while not rs4.eof
%>
<tr>
	<td class="campor" align="center"><%=rs4("data")%></td>
	<td class="campor" align="center"><%=rs4("contrato")%></td>
	<td class="campor" align="right"><%=formatnumber(rs4("valor"),2)%></td>
	<td class="campor" align="center"><%=rs4("nprestacoes")%></td>
	<td class="campor" align="right"><%=formatnumber(rs4("vprestacao"),2)%></td>
	<td class="campor" align="center"><%=rs4("venc1")%></td>
	<td class="campor" align="center"><%=rs4("vencu")%></td>
</tr>
<%
rs4.movenext
loop
%>
</table>
<%
end if
rs4.close
%>

<%
'rs.movenext:loop
rs.close
set rs=nothing
'rs2.close
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>