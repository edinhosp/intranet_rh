<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a52")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Autônomo</title>
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
<script language="javascript" type="text/javascript"><!--
function nomeu() {
	form.nome_autonomo.value=form.nome_autonomo.value.toUpperCase()
}
// --></script>

<script language="VBScript">
Function ChecaCPF(NUMERO, tipo)
'tipo checagem - 1 para verificar se está correto 2 - para informar qual o correto
   tamanho=len(numero)
   numero=string(11-tamanho,"0")&numero
   If Len(NUMERO) < 10 Then
      If tipo = 1 Then ChecaCPF = False
      If tipo = 2 Then ChecaCPF = String(11, "0")
      Exit Function
   End If
   digito1 = Mid(NUMERO, 10, 1)
   digito2 = Right(NUMERO, 1)
   soma = 0: nlm11 = 2
   For a = 1 To 9
      soma = soma + (a + 1) * cdbl((Mid(NUMERO, 10 - a, 1)))
   Next
   resto0 = 11 - (soma Mod 11)
   If resto0 > 9 Then resto0 = 0 Else resto0 = resto0
   soma = 0
   If resto0 = cint(digito1) Then
      For a = 1 To 10
         soma = soma + (a + 1) * (Mid(NUMERO, 11 - a, 1))
      Next
      resto2 = 11 - (soma Mod 11)
      If resto2 > 9 Then resto2 = 0 Else resto2 = resto2
   Else
      numero2 = Left(NUMERO, 9) & resto0 & Right(NUMERO, 1)
      For a = 1 To 10
         soma = soma + (a + 1) * (Mid(numero2, 11 - a, 1))
      Next
      resto2 = 11 - (soma Mod 11)
      If resto2 > 9 Then resto2 = 0 Else resto2 = resto2
      'resto2 = digito2
   End If
   If resto0 = cint(digito1) And resto2 = cint(digito2) Then
      If tipo = 1 Then ChecaCPF = True
      If tipo = 2 Then ChecaCPF = NUMERO
   Else
      If tipo = 1 Then ChecaCPF = False
      If tipo = 2 Then ChecaCPF = Left(NUMERO, 9) & resto0 & resto2
   End If
End Function

Function ChecaNIT(NUMERO, tipo)
'tipo checagem - 1 para verificar se está correto 2 - para informar qual o correto
   'NUMERO = Clng(NUMERO)
   digito = Right(NUMERO, 1)
   soma = 0: peso = 3298765432
   If NUMERO = 0 or Numero="" Then
      If tipo = 1 Then ChecaNIT = False
      If tipo = 2 Then ChecaNIT = String(11, "0")
      Exit Function
   End If
   If (Left(NUMERO, 10) < 1060000001 Or Left(NUMERO, 10) > 1069999999) And _
      (Left(NUMERO, 10) < 1090000000 Or Left(NUMERO, 10) > 1199027231) Then
      If tipo = 1 Then ChecaNIT = False
      If tipo = 2 Then ChecaNIT = "INVALIDO"
      Exit Function
   Else
      For a = 1 To 10
         soma = soma + cint(Mid(peso, a, 1)) * cint(Mid(NUMERO, a, 1))
      Next
      calculo = (soma Mod 11)
      If calculo = 0 Or calculo = 1 Then digreal = 0 Else digreal = 11 - calculo
      MsgBox "Calculo: " & calculo & " Digreal " & digreal & " Digito " & digito
      If cint(digito) = cint(digreal) Then
         If tipo = 1 Then ChecaNIT = True
         If tipo = 2 Then ChecaNIT = ChecaNIT
      Else
         If tipo = 1 Then ChecaNIT = False
         If tipo = 2 Then ChecaNIT = Left(NUMERO, 10) & digreal
      End If
   End If
End Function

Function ChecaPIS(NUMERO, tipo)
'tipo checagem - 1 para verificar se está correto 2 - para informar qual o correto
   'NUMERO = TextoPuro(NUMERO, 2)
   if numero="" then numero="00000000000"
   digito = cint(Right(NUMERO, 1))
   Saldo = 0: Mult = 2
   For a = 1 To 10
      Saldo = Saldo + Mult * cint(Mid(NUMERO, 11 - a, 1))
      If Mult < 9 Then Mult = Mult + 1 Else Mult = 2
   Next
   Resto = Saldo Mod 11
   digreal = 11 - Resto
   If digreal > 9 Then digreal = 0
   If digreal <> digito Then
      PisReal = Left(NUMERO, 10) & digreal
      If tipo = 1 Then ChecaPIS = False
      If tipo = 2 Then ChecaPIS = PisReal
      Exit Function
   End If
   If tipo = 1 Then ChecaPIS = True
   If tipo = 2 Then ChecaPIS = NUMERO
End Function

	Sub cpf_onChange
		tempcpf=document.form.cpf.value
		tempcpf=replace(tempcpf,".","")
		tempcpf=replace(tempcpf,"-","")
		if checacpf(tempcpf,1)=false then
			b=checacpf(tempcpf,2)
			a=msgbox("Este CPF está incorreto:" & vbcrlf & "O CPF provavelmente é " & b ,48,"Atenção")
		end if
	End sub
	
	Sub nit_onChange
		tempnit=document.form.nit.value
		tempnit=replace(tempnit,".","")
		tempnit=replace(tempnit,"-","")
		if checanit(tempnit,1)=false and checapis(tempnit,1)=false then
			d=checanit(tempnit,2)
			c=checapis(tempnit,2)
			a=msgbox("Este PIS/NIT está incorreto:" & vbcrlf & "Se NIT provavelmente é " & d & vbcrlf & "Se PIS provavelmente é " & c ,48,"Atenção")
		end if
	End sub

</script>
</head>
<body>
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function chapa1() { form.chapa.value=form.nome2.value;	}
function chapa2() {	form.nome2.value=form.chapa.value;	}
--></script>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	
if request.form("cbo")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('CBO: Selecione o código do CBO-2002!');</script>"
if request.form("tipo_prestacao")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Serviço prestado: Informe o serviços prestado!');</script>"
if request.form("nome_autonomo")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Nome do autônomo: Informe o nome!');</script>"

	sql = "INSERT INTO autonomo (nome_autonomo,telefone,celular,rua,numero,complemento,bairro,cidade,estado,cep,cpf,"
	if request.form("nit")<>"" then sql=sql & "nit, "
	sql = sql & "rg,orgao_rg,ccm,tipo_prestacao,cbo,bancocod,banconome,agencia,conta,sexo,dtnascimento "
	sql = sql & ") "

	sql2 = " SELECT '" & ucase(request.form("nome_autonomo")) & "' "
	sql2=sql2 & ", '" & request.form("telefone") & "' "
	sql2=sql2 & ", '" & request.form("celular") & "' "
	sql2=sql2 & ", '" & request.form("rua") & "' "
	sql2=sql2 & ", '" & request.form("numero") & "' "
	sql2=sql2 & ", '" & request.form("complemento") & "' "
	sql2=sql2 & ", '" & request.form("bairro") & "' "
	sql2=sql2 & ", '" & request.form("cidade") & "' "
	sql2=sql2 & ", '" & request.form("estado") & "' "
	sql2=sql2 & ", '" & request.form("cep") & "' "
	sql2=sql2 & ", '" & request.form("cpf") & "' "
	if request.form("nit")<>"" then sql2=sql2 & ", '" & request.form("nit") & "' "
	sql2=sql2 & ", '" & request.form("rg") & "' "
	sql2=sql2 & ", '" & request.form("orgao_rg") & "' "
	sql2=sql2 & ", '" & request.form("ccm") & "' "
	sql2=sql2 & ", '" & request.form("tipo_prestacao") & "' "
	sql2=sql2 & ", '" & request.form("cbo") & "' "
	sql2=sql2 & ", '" & request.form("bancocod") & "' "
	sql2=sql2 & ", '" & request.form("banconome") & "' "
	sql2=sql2 & ", '" & request.form("agencia") & "' "
	sql2=sql2 & ", '" & request.form("conta") & "' "
	sql2=sql2 & ", '" & request.form("sexo") & "' "
	if request.form("dtnascimento")="" then dtnascimento="null" else dtnascimento="'" & dtaccess(request.form("dtnascimento")) & "'"
	sql2=sql2 & ", " & dtnascimento & " "
	sql1 = sql & sql2 & ""
	'response.write sql1
	if tudook=1 then conexao.Execute sql1, , adCmdText
else 'request.form=""
end if

'if request.form="" then
%>
<form method="POST" action="autonomo_nova.asp" name="form">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr><td class=grupo>Inclusão de Autônomo</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>Nome do autônomo</td>
	<td class=titulo>Sexo</td>
	<td class=titulo>Data de Nascimento</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="nome_autonomo" size="50" value="<%=request.form("nome_autonomo")%>" onkeypress="nomeu()"></td>
	<td class=fundo><select name="sexo">
		<option value="F" <%if request.form("sexo")="F" then response.write "selected"%>>Feminino</option>
		<option value="M" <%if request.form("sexo")="M" then response.write "selected"%>>Masculino</option>
	</select>
	</td>
	<td class=fundo><input type="text" name="dtnascimento" size="12" value="<%=request.form("dtnascimento")%>"></td>
<!--onFocus="this.blur()"-->
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>Telefone</td>
	<td class=titulo>Celular</td>
	<td class=titulo>Rua   </td>
	<td class=titulo>Numero   </td>
</tr>
<tr>
	<td class=fundo><input type="text" name="telefone" size="12" value="<%=request.form("telefone")%>"></td>
	<td class=fundo><input type="text" name="celular" size="12" value="<%=request.form("celular")%>"></td>
	<td class=fundo><input type="text" name="rua" size="45" value="<%=request.form("rua")%>" onfocus="javascript:window.status='Informe o endereço do autônomo'"></td>
	<td class=fundo><input type="text" name="numero" size="5" value="<%=request.form("numero")%>" onfocus="javascript:window.status='Informe o endereço do autônomo'"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>Complemento   </td>
	<td class=titulo>Bairro  </td>
	<td class=titulo>Cidade   </td>
	<td class=titulo>UF   </td>
	<td class=titulo>CEP   </td>
</tr>
<tr>
	<td class=fundo><input type="text" name="complemento" size="15" value="<%=request.form("complemento")%>" onfocus="javascript:window.status='Informe o endereço do autônomo'"></td>
	<td class=fundo><input type="text" name="bairro" size="20" value="<%=request.form("bairro")%>" onfocus="javascript:window.status='Informe o endereço do autônomo'"></td>
	<td class=fundo><input type="text" name="cidade" size="15" value="<%=request.form("cidade")%>" onfocus="javascript:window.status='Informe o endereço do autônomo'"></td>
	<td class=fundo><input type="text" name="estado" size="2" value="<%=request.form("estado")%>" onfocus="javascript:window.status='Informe o endereço do autônomo'"></td>
	<td class=fundo><input type="text" name="cep" size="10" value="<%=request.form("cep")%>" onfocus="javascript:window.status='Informe o endereço do autônomo'"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>C.P.F.</td>
	<td class=titulo>PIS ou NIT</td>
	<td class=titulo>RG/Identidade</td>
	<td class=titulo>Orgão emissor RG</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="cpf"      size="14" onchange="validaCPF()" value="<%=request.form("cpf")%>"      onfocus="javascript:window.status='Informe o número do CPF'"></td>
	<td class=fundo><input type="text" name="nit"      size="14" value="<%=request.form("nit")%>"      onfocus="javascript:window.status='Informe o número do PIS, PASEP ou NIT'"></td>
	<td class=fundo><input type="text" name="rg"       size="14" value="<%=request.form("rg")%>"       onfocus="javascript:window.status='Informe o número da Identidade (RG)'"></td>
	<td class=fundo><input type="text" name="orgao_rg" size="14" value="<%=request.form("orgao_rg")%>" onfocus="javascript:window.status='Informe o orgão que emitiu o RG'"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo>C.C.M. (Inscr.Municipal)</td>
	<td class=titulo>Serviço habitualmente prestado</td>
	<td class=titulo>C.B.O.</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="ccm" size="14" value="<%=request.form("ccm")%>" onfocus="javascript:window.status='Informe o número da Inscrição municipal'"></td>
	<td class=fundo><input type="text" name="tipo_prestacao" size="44" value="<%=request.form("tipo_prestacao")%>"      onfocus="javascript:window.status='Informe uma descrição do serviço que este autônomo presta'"></td>
	<td class=fundo><input type="text" name="cbo" size="7" value="<%=request.form("cbo")%>" onfocus="javascript:window.status='Informe o número do CBO-2002'">
		<a href="pesquisa_cbo.asp" onclick="NewWindow(this.href,'PesquisaCBO','415','200','yes','center');return false" onfocus="this.blur()">
		<img src="../images/magnify.gif" border="0" width=13></a>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Cod.Banco</td>
	<td class=titulo>Banco</td>
	<td class=titulo>Agência</td>
	<td class=titulo>Conta Corrente</td>
	<td class=titulo>E-mail</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="bancocod"  size="3"  value="<%=request.form("bancocod")%>"      onfocus="javascript:window.status='Informe o código do Banco'"></td>
	<td class=fundo><input type="text" name="banconome" size="10" value="<%=request.form("banconome")%>"      onfocus="javascript:window.status='Informe o nome do Banco'"></td>
	<td class=fundo><input type="text" name="agencia"   size="6"  value="<%=request.form("agencia")%>"       onfocus="javascript:window.status='Informe o número da agência'"></td>
	<td class=fundo><input type="text" name="conta"     size="10" value="<%=request.form("conta")%>" onfocus="javascript:window.status='Informe o número da conta corrente'"></td>
	<td class=fundo><input type="text" name="email"     size="30" value="<%=request.form("email")%>" onfocus="javascript:window.status='Informe o email'"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2" onfocus="javascript:window.status='Clique para desfazer e limpar a tela'"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()" onfocus="javascript:window.status='Clique aqui para fechar sem salvar'"></td>
</tr>
</table>
</form>
<%
'else
'rs.close
set rs=nothing
'end if
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<!--
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if
%>
</body>
</html>