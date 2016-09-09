<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Marcação - Estagiário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function renovacao1()	{ form.urenovacao.value=form.renovacao_anterior.value;	}
--></script>
<script language="VBScript">
	Sub pagina_onChange
		ok=true:dim frm:set frm=document.form
		if ok=true then frm.submit
	End Sub
	Sub ent1_onChange
		Jornada
	End Sub
	Sub sai1_onChange
		Jornada
	End Sub
	Sub ent2_onChange
		Jornada
	End Sub
	Sub sai2_onChange
		Jornada
	End Sub
	Sub Jornada()
		ent1=0:sai1=0:ent2=0:sai2=0
		msgbox document.form.ent1.value
		msgbox horasave(document.form.ent1.value)
		if document.form.ent1.value="" then ent1=0 else ent1=horasave(document.form.ent1.value)
		msgbox "ent1 " & ent1
		if document.form.sai1.value="" then sai1=0 else sai1=horasave(document.form.sai1.value)
		msgbox "sai1 " & sai1
		if document.form.ent2.value="" then ent2=0 else ent2=horasave(document.form.ent2.value)
		msgbox "ent2 " & ent2
		if document.form.sai2.value="" then sai2=0 else sai2=horasave(document.form.sai2.value)
		msgbox "sai2 " & sai2
		if sai1-ent1>0 then jorn1=(sai1-ent1) else jorn1=0
		if sai2-ent2>0 then jorn2=(sai2-ent2) else jorn2=0
		jorn=jorn1+jorn2
		document.form.jorn.value=horaload(jorn,2)
		'document.form.t.value=ent1&"-"&sai1&"-"&ent2&"-"&sai2 & chr(10) & jorn & "-" &jorn1&"-"&jorn2 & chr(10) & (sai1-ent1) & chr(10) & (sai2=ent2)
	End Sub
</script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(4), varcur(6)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	if request.form("comp")="ON" then    comp1=-1 else    comp1=0
	if request.form("desc")="ON" then    desc1=-1 else    desc1=0
	if request.form("intflex")="ON" then intflex1=-1 else intflex1=0
	ent1h=request.form("ent1_h")
	ent1m=request.form("ent1_m")
	sai1h=request.form("sai1_h")
	sai1m=request.form("sai1_m")
	ent2h=request.form("ent2_h")
	ent2m=request.form("ent2_m")
	sai2h=request.form("sai2_h")
	sai2m=request.form("sai2_m")
	jornh=request.form("jorn_h")
	jornm=request.form("jorn_m")
	
	if ent1h="" then ent1h=0:if ent1m="" then ent1m=0
	if sai1h="" then sai1h=0:if sai1m="" then sai1m=0
	if ent2h="" then ent2h=0:if ent2m="" then ent2m=0
	if sai2h="" then sai2h=0:if sai2m="" then sai2m=0
	if jornh="" then jornh=0:if jornm="" then jornm=0
	
	ent1=(ent1h*60)+ent1m
	sai1=(sai1h*60)+sai1m
	ent2=(ent2h*60)+ent2m
	sai2=(sai2h*60)+sai2m
	jorn=(jornh*60)+jornm

	sql="UPDATE est_cadhorario_marcacoes SET "
	sql=sql & "dia=" & request.form("dia") & ", "
	sql=sql & "ent1=" & ent1 & ", "
	sql=sql & "sai1=" & sai1 & ", "
	sql=sql & "ent2=" & ent2 & ", "
	sql=sql & "sai2=" & sai2 & ", "
	sql=sql & "jorn=" & jorn & ", "
	sql=sql & "[comp]=" & comp1 & ", "
	sql=sql & "[desc]=" & desc1 & ", "
	sql=sql & "intflex=" & intflex1 & " "
	'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
	'sql=sql & ",dataa   =getdate() "
	sql=sql & " WHERE codigo='" & session("idcadhor") & "' and dia=" & session("iddia")
	response.write sql
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM est_cadhorario_marcacoes WHERE codigo='" & session("idcadhor") & "' and dia=" & session("iddia")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")=null or request("codigo")="" then
		idcadhor=session("idcadhor")
		iddia=session("iddia")
		'if session("idcadhor")="" then id_cadhor=request.form("id_cadhor")
	else
		idcadhor=request("codigo")
		iddia=request("dia")
	end if
	sqla="select * from est_cadhorario_marcacoes " & _
	"where codigo='" & idcadhor & "' and dia=" & iddia
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
session("idcadhor")=rs("codigo")
session("iddia")=rs("dia")
if request.form("codigo")=""  then codigo=rs("codigo")          else codigo=request.form("codigo")
if request.form("dia")=""     then dia=rs("dia")                else dia=request.form("dia")
if request.form("ent1")=""    then ent1=horaload(rs("ent1"),2)  else ent1=request.form("ent1")
if request.form("sai1")=""    then sai1=horaload(rs("sai1"),2)  else sai1=request.form("sai1")
if request.form("ent2")=""    then ent2=horaload(rs("ent2"),2)  else ent2=request.form("ent2")
if request.form("sai2")=""    then sai2=horaload(rs("sai2"),2)  else sai2=request.form("sai2")
if request.form("jorn")=""    then jorn=horaload(rs("jorn"),2)  else jorn=request.form("jorn")
if request.form("comp")=""    then comp=rs("comp")              else comp=request.form("comp")
if request.form("desc")=""    then desc=rs("desc")               else desc=request.form("desc")
if request.form("intflex")="" then intflex=rs("intflex")        else intflex=request.form("intflex")
if comp<>0 or comp=true or comp="ON" then comp1="checked" else comp1=""
if desc<>0 or desc=true or desc="ON" then desc1="checked" else desc1=""
if intflex<>0 or intflex=true or intflex="ON" then intflex1="checked" else intflex1=""

%>
<form method="POST" action="hor_marc_alteracao.asp" name="form">
<input type="hidden" name="id_cadhor" size="4" value="<%=rs("codigo")%>" >  
<input type="hidden" name="id_dia" size="4" value="<%=rs("dia")%>" >  
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Cadastro de Marcações - Horário <%=idcadhor%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Dia</td>
	<td class=titulo>Entr.</td>
	<td class=titulo>Saida</td>
	<td class=titulo>Entr.</td>
	<td class=titulo>Saida</td>
</tr>
<tr>
	<td class=titulo><input type="text" name="dia" size="3" value="<%=dia%>" class="form_input" style="text-align:center"></td>
	<td class=fundo><input type="text" name="ent1" size="4" value="<%=ent1%>" class="form_input" style="text-align:center"></td>
	<td class=fundo><input type="text" name="sai1" size="4" value="<%=sai1%>" class="form_input" style="text-align:center"></td>
	<td class=fundo><input type="text" name="ent2" size="4" value="<%=ent2%>" class="form_input" style="text-align:center"></td>
	<td class=fundo><input type="text" name="sai2" size="4" value="<%=sai2%>" class="form_input" style="text-align:center"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Jornada</td>
	<td class=titulo>Comp.</td>
	<td class=titulo>Desc.</td>
	<td class=titulo>Int.Flex.?</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="jorn" size="4" value="<%=jorn%>" class="form_input" style="text-align:center"></td>
	<td class=fundo><input type="checkbox" name="comp" value="ON" <%=comp1%>>
	<td class=fundo><input type="checkbox" name="desc" value="ON" <%=desc1%>>
	<td class=fundo><input type="checkbox" name="intflex" value="ON" <%=intflex1%>>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=titulo colspan=3>&nbsp;</td></tr>
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Alterações  " class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
end if
set rs=nothing
set rsc=nothing
set rsnome=nothing
conexao.close
set conexao=nothing

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
%>
</body>
</html>