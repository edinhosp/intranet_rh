<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a59")="N" or session("a59")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Opção Medial Odonto</title>
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
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function nome1() { form.chapa.value=form.nome.value;form.submit(); }
function chapa1() { form.nome.value=form.chapa.value;form.submit(); }
--></script>
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

if request.form("B1")="" then 'or request.form("id")="" then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para opção de Odonto
<form method="POST" action="opcao_odonto.asp" name="form">
<%
sqla="SELECT f.chapa, f.NOME FROM corporerm.dbo.pfunc f  " & _
"WHERE f.CODSINDICATO<>'03' AND f.CODSITUACAO<>'D' GROUP BY f.chapa, f.NOME ORDER BY f.NOME;"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=request.form("chapa")%>">
<select name="nome" class=a onchange="nome1()">
	<option value="00000">===> Ficha em branco <===</option>
<%
rs.movefirst:do while not rs.eof
if request.form("chapa")=rs("chapa") then temps="selected" else temps=""
%>
	<option value="<%=rs("chapa")%>" <%=temps%> ><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<br><br>
<table border="1" cellpadding="2" cellspacing="2" style="border-collapse: collapse">
<tr>
	<td class=titulo></td>
	<td class=titulo>Tipo</td>
	<td class=titulo>Nome</td>
</tr>
<%
sqlt="SELECT 'Titular' as tipo, m.chapa, f.NOME FROM assmed_mudanca AS m INNER JOIN corporerm.dbo.pfunc AS f ON m.chapa = f.CHAPA collate database_default " & _
"WHERE f.CODSINDICATO<>'03' AND f.CODSITUACAO<>'D' and m.chapa='" & request.form("chapa") & "' GROUP BY m.chapa, f.NOME ORDER BY f.NOME;"
sqld=" "
sqlfinal=sqlt & " union all " & sqld
sqlfinal=sqlt
rs.Open sqlfinal, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
do while not rs.eof
if rs("tipo")="Titular" then tipo="T" else tipo="D"
%>
<tr>
	<td class=campo><input type="radio" name="id" value="<%=tipo%><%=rs("chapa")%>"></td>
	<td class=campo><%=rs("tipo")%></td>
	<td class=campo><%=rs("nome")%></td>
</tr>
<%
rs.movenext
loop
end if
rs.close
%>
</table>
<br>
<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>
<%
end if

if request.form("B1")<>"" then 'and request.form("id")<>"" then
'temp=request.form("id")
'tipo=left(temp,1)
'codigo=right(temp,len(temp)-1)
chapa=request.form("chapa")
sqla="select chapa, nome, admissao, dtnascimento, sexo, estcivil, mae, cpf, cartidentidade, orgemissorident, ufcartident, rua, numero, complemento, " & _
"bairro, cidade, cep, estado, email, telefone1, telefone2, telefone3, secao, funcao " & _
"from qry_funcionarios q where chapa<'10000' "
sqlb="AND q.CHAPA='" & chapa & "' "
sql1=sqla & sqlb
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
idade=int((now()-rs("dtnascimento"))/365.25)
else

end if
%>
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=100>
<tr><td valign="center" align="left"  valign=middle><img src="../images/medial_odonto.gif" border="0"></td>
	<td valign="center" align="center" valign=middle><font size=3><b>FICHA DE ADESÃO</b></font></td>
</tr>
</tr>
<tr><td colspan=2 height=25 class=fundop align="center" style="background='#72E4AB'"><b>DADOS DO BENEFICIÁRIO TITULAR</td></tr>
</table>
<%
if request.form("chapa")="00000" or request.form("chapa")="" then
	admissao="" : registro="" : secao="" : funcao="" : nome="" : nascimento=""
	mae="" : estcivil="" : sexo="" : cpf="" : rg="" : orgao="" : rua=""
	numero="" : complemento="" : bairro="" : cidade="" : cep="" : estado=""
	email="" : telefone1="" : telefone2="" : telefone3="" : sf="" : sm=""
else
	admissao=rs("admissao")
	registro=rs("chapa")
	secao=rs("secao")
	funcao=rs("funcao")
	nome=rs("nome")
	nascimento=rs("dtnascimento")
	mae=rs("mae")
	estcivil=rs("estcivil")
	sexo=rs("sexo")
	cpf=rs("cpf")
	rg=rs("cartidentidade")
	orgao=rs("orgemissorident") & " " & rs("ufcartident")
	rua=rs("rua")
	numero=rs("numero")
	complemento=rs("complemento")
	bairro=rs("bairro")
	cidade=rs("cidade")
	cep=rs("cep")
	estado=rs("estado")
	email=rs("email")
	telefone1=rs("telefone1")
	telefone2=""
	telefone3=rs("telefone2")
	if sexo="F" then sf="X" else sm="X"
end if
tamanho=32:tamanho2=30
corborda="#009999"
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td width="60%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Nome da Empresa</td>
	<td width="15%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Data de admissão</td>
	<td width="25%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:0 solid">	&nbsp;Número do Registro / Matrícula</td>
</tr>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;FUNDACAO INSTITUTO DE ENSINO PARA OSASCO</td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=admissao%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right:0 solid">	&nbsp;<%=registro%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td width="60%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Área/Setor</td>
	<td width="40%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:0 solid">	&nbsp;Área/Setor</td>
</tr>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=secao%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right:0 solid">	&nbsp;<%=funcao%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td width="80%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Nome Completo</td>
	<td width="20%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:0 solid">	&nbsp;Data nascimento</td>
</tr>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=nome%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right:0 solid">	&nbsp;<%=nascimento%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td width="20%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Sexo</td>
	<td width="25%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Estado civil</td>
	<td width="55%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:0 solid">	&nbsp;Nome da mãe</td>
</tr>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;[&nbsp;<%=sf%>&nbsp;] F &nbsp;&nbsp;[&nbsp;<%=sm%>&nbsp;] M</td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=estcivil%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right:0 solid">	&nbsp;<%=mae%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td width="35%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;CPF</td>
	<td width="25%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;RG</td>
	<td width="20%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Orgão emissor</td>
	<td width="20%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:0 solid">	&nbsp;País</td>
</tr>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=cpf%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=rg%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=orgao%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right:0 solid">	&nbsp;<%=""%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td width="65%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Endereço residencial</td>
	<td width="10%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Nº</td>
	<td width="15%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:0 solid">	&nbsp;Complemento</td>
</tr>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=rua%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=numero%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right:0 solid">	&nbsp;<%=complemento%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td width="40%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Bairro</td>
	<td width="30%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Cidade</td>
	<td width="20%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;CEP</td>
	<td width="10%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:0 solid">	&nbsp;UF</td>
</tr>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=bairro%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=cidade%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;<%=cep%></td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right:0 solid">	&nbsp;<%=estado%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td width="40%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;E-mail</td>
	<td width="20%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Tel. residencial</td>
	<td width="20%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Tel. comercial</td>
	<td width="20%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:0 solid">	&nbsp;Tel. celular</td>
</tr>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;<%=email%></td>
	<td bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;<%=telefone1%></td>
	<td bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;<%=""%></td>
	<td bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:0 solid">	&nbsp;<%=telefone3%></td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td colspan=2 height=25 class=fundop align="center" style="background='#72E4AB'"><b>DEPENDENTES</td></tr>
</table>

<%for a=1 to 5%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho2%>>
<tr>
	<td width="55%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;<b><%=a%>. Nome completo</b></td>
	<td width="15%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Parentesco *</td>
	<td width="20%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Data Nascimento</td>
	<td width="10%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:0 solid">	&nbsp;Sexo</td>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;</td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;</td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;</td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right:0 solid">	&nbsp;</td>
</tr></table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho2%>>
<tr>
	<td width="55%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;&nbsp;Nome da mãe do dependente</td>
	<td width="22%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right: 1px solid">	&nbsp;Celular</td>
	<td width="23%" class="campor" bordercolor=<%=corborda%> style="border-bottom:0 solid;border-right:0 solid">	&nbsp;CPF **</td>
<tr>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;</td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right: 1px solid">	&nbsp;</td>
	<td bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right:0 solid">	&nbsp;</td>
</tr></table>
<%next%>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=45>
<tr>
	<td class="campor" bordercolor=<%=corborda%> style="border-top:2 solid;border-bottom:3 solid" valign="middle">
	(*) PARENTESCO: Dependentes:  A = Cônjuge, B = Filho(a) até 24 anos, C = Tutelados ou adotivos, D = Companheiro(a). <br>
	<%for a=1 to 28%>&nbsp;<%next%>Os itens C e D são necessários documentos comprobatórios.<br>
	(**) <FONT COLOR="GRAY">Caso não preencha o nome da mãe.
	</td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho%>>
<tr>
	<td colspan=2 class="campop" bordercolor=<%=corborda%> style="border-bottom: 1px solid;border-right:0 solid" height=25 valign="middle">&nbsp;Plano: Medial Odonto</td>
</tr>
<tr>
	<td width="50%" bordercolor=<%=corborda%> style="border-bottom:3 solid;border-right: 1px solid">	&nbsp;Titular<br>&nbsp;R$</td>
	<td width="50%" bordercolor=<%=corborda%> style="border-bottom:3 solid;border-right:0 solid">	&nbsp;Dependentes<br>&nbsp;R$</td>
</tr>
<tr>
	<td colspan=2 class="campop" bordercolor=<%=corborda%> style="border-bottom:3 solid;border-right:0 solid" height=25 valign="middle">&nbsp;Valor Total R$</td>
</tr>
</table>

<br>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td colspan=2 height=25 class=fundop align="center" style="background='#72E4AB'"><b>CONDIÇÕES PARA VALIDADE DO BENEFÍCIO</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=<%=tamanho2%>>
<tr>
	<td class=campo bordercolor=<%=corborda%> style="border-top:0 solid;border-bottom:3 solid"><br>
	Declaro assumir integralmente o valor da taxa mensal correspondente a cada um dos inscritos, respeitando inclusive, majorações de valores por força de obrigações
	contratuais, autorizando assim, o desconto em folha de pagamento, comprometendo-me a permanecer no plano, ora referido, por um período mínimo de 12 meses. 
	O preenchimento desta ficha de adesão, garante minha inscrição ao plano odontológico e consequentemente a utilização dos serviços odontológicos MEDIAL ODONTO, 
	aceitando as condições contratuais gerais de atendimento e cobertura do plano.
	
	<p align="right">Adesão realizada em Osasco, _____ / _____ / _________
	<p align="right">Assinatura do Beneficiário _________________________________
	<br>&nbsp;
	</td>
</tr></table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td colspan=2 height=25 class=campo align="right"><img src="../images/medial_odonto_ans.gif" width="97" height="20" border="0" alt=""></td></tr>
</table>

<%
rs.close
set rs=nothing

set rs2=nothing
end if ' 

conexao.close
set conexao=nothing
%>
</body>
</html>