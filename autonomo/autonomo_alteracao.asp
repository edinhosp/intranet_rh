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
<title>Alteração de Autônomo</title>
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
	
</script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
	tudook=1	
if request.form("cbo")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('CBO: Selecione o código do CBO-2002!');</script>"
if request.form("tipo_prestacao")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Serviço prestado: Informe o serviços prestado!');</script>"
if request.form("nome_autonomo")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Nome do autônomo: Informe o nome!');</script>"

	sql="UPDATE autonomo SET "
	sql=sql & "nome_autonomo= '" & ucase(request.form("nome_autonomo"))      & "' "
	sql=sql & ",telefone     = '" & request.form("telefone") & "' "
	sql=sql & ",celular      = '" & request.form("celular") & "' "
	sql=sql & ",rua          = '" & request.form("rua") & "' "
	sql=sql & ",numero       = '" & request.form("numero") & "' "
	sql=sql & ",complemento  = '" & request.form("complemento") & "' "
	sql=sql & ",bairro       = '" & request.form("bairro") & "' "
	sql=sql & ",cidade       = '" & request.form("cidade") & "' "
	sql=sql & ",estado       = '" & request.form("estado") & "' "
	sql=sql & ",cep          = '" & request.form("cep") & "' "
	cpf=request.form("cpf"):cpf=replace(cpf,".",""):cpf=replace(cpf,"-","")
	sql=sql & ",cpf          = '" & cpf        & "' "
	sql=sql & ",nit          = '"  & request.form("nit")     & "' "
	sql=sql & ",rg           = '" & request.form("rg")       & "' "
	sql=sql & ",orgao_rg     = '" & request.form("orgao_rg") & "' "
	sql=sql & ",ccm          = '" & request.form("ccm")      & "' "
	sql=sql & ",tipo_prestacao='" & request.form("tipo_prestacao") & "' "
	sql=sql & ",cbo          = '"  & request.form("cbo")     & "' "
	sql=sql & ",bancocod     ='" & request.form("bancocod")  & "' "
	sql=sql & ",banconome    ='" & request.form("banconome") & "' "
	sql=sql & ",agencia      ='" & request.form("agencia")   & "' "
	sql=sql & ",conta        ='" & request.form("conta")     & "' "
	sql=sql & ",email        ='" & request.form("email")     & "' "
	sql=sql & ",sexo         ='" & request.form("sexo")      & "' "
	if request.form("dtnascimento")="" then dtnascimento="null" else dtnascimento="'" & dtaccess(request.form("dtnascimento")) & "'"
	sql=sql & ",dtnascimento =" & dtnascimento     & " "
	sql=sql & "WHERE id_autonomo=" & session("id_alt_autonomo")
	'response.write sql
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM autonomo WHERE id_autonomo=" & session("id_alt_autonomo")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")=null then
		id_autonomo=session("id_alt_autonomo")
	else
		id_autonomo=request("codigo")
	end if
	sqla="select * from autonomo "
	sqlb="where id_autonomo=" & id_autonomo
	sql1=sqla & sqlb 
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if


if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'rs.movefirst
'do while not rs.eof 
session("id_alt_autonomo")=rs("id_autonomo")
%>
<form method="POST" action="autonomo_alteracao.asp" name="form">
<input type="hidden" name="id_autonomo" size="4" value="<%=rs("id_autonomo")%>">  
<table border="0" cellpadding="3" cellspacing="0" width="500">
	<tr><td class=grupo>Alteração de Autônomo</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Nome do autônomo</td>
	<td class=titulo>Sexo</td>
	<td class=titulo>Data de Nascimento</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="nome_autonomo" size="50" value="<%=rs("nome_autonomo")%>"></td>
	<td class=fundo><select name="sexo">
		<option value="F" <%if rs("sexo")="F" then response.write "selected"%>>Feminino</option>
		<option value="M" <%if rs("sexo")="M" then response.write "selected"%>>Masculino</option>
	</select>
	</td>
	<td class=fundo><input type="text" name="dtnascimento" size="12" value="<%=rs("dtnascimento")%>"></td>
<!--onFocus="this.blur()"-->
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Telefone</td>
	<td class=titulo>Celular</td>
	<td class=titulo>Rua</td>
	<td class=titulo>Numero</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="telefone" size="12" value="<%=rs("telefone")%>"></td>
	<td class=fundo><input type="text" name="celular" size="12" value="<%=rs("celular")%>"></td>
	<td class=fundo><input type="text" name="rua" size="45" value="<%=rs("rua")%>" onfocus="javascript:window.status='Informe o endereço do autônomo'"></td>
	<td class=fundo><input type="text" name="numero" size="5" value="<%=rs("numero")%>" onfocus="javascript:window.status='Informe o numero'"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Complemento</td>
	<td class=titulo>Bairro</td>
	<td class=titulo>Cidade</td>
	<td class=titulo>UF</td>
	<td class=titulo>CEP</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="complemento" size="15" value="<%=rs("complemento")%>" onfocus="javascript:window.status='Informe o complemento do endereço (apto, casa, etc)'"></td>
	<td class=fundo><input type="text" name="bairro" size="20" value="<%=rs("bairro")%>" onfocus="javascript:window.status='Informe o bairro'"></td>
	<td class=fundo><input type="text" name="cidade" size="15" value="<%=rs("cidade")%>" onfocus="javascript:window.status='Informe a cidade'"></td>
	<td class=fundo><input type="text" name="estado" size="2" value="<%=rs("estado")%>" onfocus="javascript:window.status='Informe o Estado'"></td>
	<td class=fundo><input type="text" name="cep" size="10" value="<%=rs("cep")%>" onfocus="javascript:window.status='Informe o CEP'"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>C.P.F.</td>
	<td class=titulo>PIS ou NIT</td>
	<td class=titulo>RG/Identidade</td>
	<td class=titulo>Orgão emissor RG</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="cpf"      size="14" value="<%=rs("cpf")%>"      onfocus="javascript:window.status='Informe o número do CPF'"></td>
	<td class=fundo><input type="text" name="nit"      size="14" value="<%=rs("nit")%>"      onfocus="javascript:window.status='Informe o número do PIS, PASEP ou NIT'"></td>
	<td class=fundo><input type="text" name="rg"       size="14" value="<%=rs("rg")%>"       onfocus="javascript:window.status='Informe o número da Identidade (RG)'"></td>
	<td class=fundo><input type="text" name="orgao_rg" size="14" value="<%=rs("orgao_rg")%>" onfocus="javascript:window.status='Informe o orgão que emitiu o RG'"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>C.C.M. (Inscr.Municipal)</td>
	<td class=titulo>Serviço habitualmente prestado</td>
	<td class=titulo>C.B.O.</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="ccm" size="14" value="<%=rs("ccm")%>"      onfocus="javascript:window.status='Informe o número da Inscrição municipal'"></td>
	<td class=fundo><input type="text" name="tipo_prestacao" size="44" value="<%=rs("tipo_prestacao")%>"      onfocus="javascript:window.status='Informe uma descrição do serviço que este autônomo presta'"></td>
	<td class=fundo><input type="text" name="cbo" size="7" value="<%=rs("cbo")%>" onfocus="javascript:window.status='Informe o número do CBO-2002'">
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
	<td class=fundo><input type="text" name="bancocod"  size="3"  value="<%=rs("bancocod")%>"      onfocus="javascript:window.status='Informe o código do Banco'"></td>
	<td class=fundo><input type="text" name="banconome" size="10" value="<%=rs("banconome")%>"      onfocus="javascript:window.status='Informe o nome do Banco'"></td>
	<td class=fundo><input type="text" name="agencia"   size="6"  value="<%=rs("agencia")%>"       onfocus="javascript:window.status='Informe o número da agência'"></td>
	<td class=fundo><input type="text" name="conta"     size="10" value="<%=rs("conta")%>" onfocus="javascript:window.status='Informe o número da conta corrente'"></td>
	<td class=fundo><input type="text" name="email"     size="30" value="<%=rs("email")%>" onfocus="javascript:window.status='Informe o email'"></td>
</tr>
</table>
  
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser alterado!');</script>"
	end if
	'Response.write "<p>Registro atualizado.<br>"
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<!--
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>