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
<title>Inclusão de Horário-Estagiário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function chapa1() {	form.chapa.value=form.nome.value;	}
function nome1() {	form.nome.value=form.chapa.value;	}
--></script>
<script language="VBScript">
	Sub pagina_onChange
		ok=true:dim frm:set frm=document.form
		if ok=true then frm.submit
	End Sub
	Sub ent1_h_onChange
		Jornada
	End Sub
	Sub ent1_m_onChange
		Jornada
	End Sub
	Sub sai1_h_onChange
		Jornada
	End Sub
	Sub sai1_m_onChange
		Jornada
	End Sub
	Sub ent2_h_onChange
		Jornada
	End Sub
	Sub ent2_m_onChange
		Jornada
	End Sub
	Sub sai2_h_onChange
		Jornada
	End Sub
	Sub sai2_m_onChange
		Jornada
	End Sub
	Sub Jornada()
		ent1=0:sai1=0:ent2=0:sai2=0
		if document.form.ent1_h.value="" then ent1h=0 else ent1h=document.form.ent1_h.value
		if document.form.ent1_m.value="" then ent1m=0 else ent1m=document.form.ent1_m.value
		if document.form.sai1_h.value="" then sai1h=0 else sai1h=document.form.sai1_h.value
		if document.form.sai1_m.value="" then sai1m=0 else sai1m=document.form.sai1_m.value
		if document.form.ent2_h.value="" then ent2h=0 else ent2h=document.form.ent2_h.value
		if document.form.ent2_m.value="" then ent2m=0 else ent2m=document.form.ent2_m.value
		if document.form.sai2_h.value="" then sai2h=0 else sai2h=document.form.sai2_h.value
		if document.form.sai2_m.value="" then sai2m=0 else sai2m=document.form.sai2_m.value
		ent1=ent1h*60+ent1m : sai1=sai1h*60+sai1m : ent2=ent2h*60+ent2m : sai2=sai2h*60+sai2m
		if sai1-ent1>0 then jorn1=(sai1-ent1) else jorn1=0
		if sai2-ent2>0 then jorn2=(sai2-ent2) else jorn2=0
		jorn=jorn1+jorn2
		jorn_h=int(jorn/60)
		jorn_m=jorn-(int(jorn/60)*60)
		document.form.jorn_h.value=jorn_h
		document.form.jorn_m.value=jorn_m
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

if request.form<>"" then
if request.form("bt_salvar")<>"" then
	tudook=1
	'response.write request.form
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

	sql = "INSERT INTO est_cadhorario_marcacoes (" 
	sql = sql & "codigo, dia, ent1, sai1, ent2, sai2, jorn, [comp], [desc], intflex "
	sql = sql & ") SELECT "
	sql = sql & "'" & request.form("codigo") & "', "
	sql = sql & request.form("dia") & ", "
	sql = sql & ent1 & ", "
	sql = sql & sai1 & ", "
	sql = sql & ent2 & ", "
	sql = sql & sai2 & ", "
	sql = sql & jorn & ", "
	sql = sql & comp1 & ", "
	sql = sql & desc1 & ", "
	sql = sql & intflex1 & " "
	'sql = sql  & ", '" & session("usuariomaster") & "', "
	'sql = sql  & ", getdate() "
	sql1 = sql
	response.write "<font size='1'>" & sql1
	if tudook=1 then conexao.Execute sql1, , adCmdText
end if
else 'request.form=""
end if

'if request.form="" then
if (request.form("bt_salvar")="") or (request.form("bt_salvar")<>"" and tudook=0) then
if request("codigo")="" then codigo=request.form("codigo") else codigo=request("codigo")
dia    =request.form("dia")
ent1_h =request.form("ent1_h")
ent1_m =request.form("ent1_m")
sai1_h =request.form("sai1_h")
sai1_m =request.form("sai1_m")
ent2_h =request.form("ent2_h")
ent2_m =request.form("ent2_m")
sai2_h =request.form("sai2_h")
sai2_m =request.form("sai2_m")
jorn_h =request.form("jorn_h")
jorn_m =request.form("jorn_m")
comp   =request.form("comp")
desc   =request.form("desc")
intflex=request.form("intflex")
if comp="ON" then comp1="checked" else comp1=""
if desc="ON" then desc1="checked" else desc1=""
if intflex="ON" then intflex1="checked" else intflex1=""
sqld="select max(dia) as udia from est_cadhorario_marcacoes where codigo='" & request("codigo") & "' "
rs.Open sqld, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then dia=rs("udia")+1 
rs.close
if dia="" or isnull(dia) then dia=1
%>
<form method="POST" action="hor_marc_nova.asp" name="form" >
<input type="hidden" name="codigo" size="3" value="<%=codigo%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Cadastro de Marcações - Horário <%=request("codigo")%></td></tr>
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
	<td class=fundo><input type="text" name="ent1_h" size="1" value="<%=ent1_h%>" class="form_input" style="text-align:center"><b>:<input type="text" name="ent1_m" size="1" value="<%=ent1_m%>" class="form_input" style="text-align:center"></td>
	<td class=fundo><input type="text" name="sai1_h" size="1" value="<%=sai1_h%>" class="form_input" style="text-align:center"><b>:<input type="text" name="sai1_m" size="1" value="<%=sai1_m%>" class="form_input" style="text-align:center"></td>
	<td class=fundo><input type="text" name="ent2_h" size="1" value="<%=ent2_h%>" class="form_input" style="text-align:center"><b>:<input type="text" name="ent2_m" size="1" value="<%=ent2_m%>" class="form_input" style="text-align:center"></td>
	<td class=fundo><input type="text" name="sai2_h" size="1" value="<%=sai2_h%>" class="form_input" style="text-align:center"><b>:<input type="text" name="sai2_m" size="1" value="<%=sai2_m%>" class="form_input" style="text-align:center"></td>
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
	<td class=fundo><input type="text" name="jorn_h" size="1" value="<%=jorn_h%>" class="form_input" style="text-align:center"><b>:<input type="text" name="jorn_m" size="1" value="<%=jorn_m%>" class="form_input" style="text-align:center"></td>
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
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>

</form>
<%
else
'rs.close
set rs=nothing
set rsc=nothing
end if
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if

'	Response.write "<p>Registro salvo.<br>"
	'response.write '<script>javascript:top.window.close();</script>
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