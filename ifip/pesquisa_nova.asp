<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a30")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Processo IFIP</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		if request.form("exp_cumpre")="ON" then exp_cumpre = -1 else exp_cumpre = 0
		sql = "INSERT INTO ifip_cadastro (" 
		sql = sql & "num_processo, status, titulo_pesquisa, linha_pesquisa, area_conhecimento, "
		sql = sql & "horas_semanais, dt_entrada, dt_termino, vigencia, aprov_valor, observacoes, "
		sql = sql & "aprov_depto, aprov_ifip, aprov_consepe, aprov_proadm, aprov_ciencia "
		sql = sql & ") "
		sql2 = " SELECT "
		sql2=sql2 & " '" & request.form("num_processo") & "', "
		sql2=sql2 & " '" & request.form("status") & "', "
		sql2=sql2 & " '" & request.form("titulo_pesquisa") & "', "
		sql2=sql2 & " '" & request.form("linha_pesquisa") & "', "
		sql2=sql2 & " '" & request.form("area_conhecimento") & "', "
		if request.form("horas_semanais")="" then sql2=sql2 & "null," else sql2=sql2 & nraccess(cdbl(request.form("horas_semanais"))) & ", "
		if request.form("dt_entrada")=""     then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("dt_entrada")) & "', "
		if request.form("dt_termino")=""     then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("dt_termino")) & "', "
		if request.form("vigencia")=""       then sql2=sql2 & "null," else sql2=sql2 & nraccess(cdbl(request.form("vigencia"))) & ", "
		if request.form("aprov_valor")=""    then sql2=sql2 & "null," else sql2=sql2 & nraccess(cdbl(request.form("aprov_valor"))) & ", "
		sql2=sql2 & " '" & request.form("observacoes") & "', "
		if request.form("aprov_depto")=""   then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("aprov_depto")) & "', "
		if request.form("aprov_ifip")=""    then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("aprov_ifip")) & "', "
		if request.form("aprov_consepe")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("aprov_consepe")) & "', "
		if request.form("aprov_proadm")=""  then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("aprov_proadm")) & "', "
		if request.form("aprov_ciencia")="" then sql2=sql2 & "null " else sql2=sql2 & " '" & dtaccess(request.form("aprov_ciencia")) & "' "
		sql1 = sql & sql2 & ""
		'response.write "<font size='2'>" & sql1
		conexao.Execute sql1, , adCmdText
	end if
else 'request.form=""
end if

if request.form="" then
%>
<form method="POST" action="pesquisa_nova.asp" name="form" >
<table border="0" cellpadding="3" cellspacing="0" width="600">
	<tr><td class=grupo>Inclusão de Processo IFIP</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
	<tr><td class=titulo>Nº Processo</td><td class=titulo>Título da Pesquisa</td></tr>
	<tr><td class=fundo><input type="text" name="num_processo" size="7" value=""></td>
		<td class=fundo rowspan=3>
			<textarea name="titulo_pesquisa" cols="60" rows="5"></textarea>
		</td>
	</tr>
	<tr><td class=titulo>Status</td></tr>
	<tr><td class=fundo>
	<select size="1" name="status" class=a>
		<option value="0">Selecione um status</option>
<%
sql2="select id_status, desc_status from ifip_wstatus"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
%>
          <option value="<%=rsc("id_status")%>"><%=rsc("desc_status")%></option>
<%
rsc.movenext:loop
rsc.close
%>
        </select>	
	</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
	<tr>
		<td class=titulo>Linha de Pesquisa</td>
		<td class=titulo>Área de Conhecimento</td>
		<td class=titulo>Horas</td>
	</tr>
	<tr>
    	<td class=fundo><input type="text" name="linha_pesquisa" size="25" value=""></td>
    	<td class=fundo><input type="text" name="area_conhecimento" size="50" value=""></td>
    	<td class=fundo><input type="text" name="horas_semanais" size="3" value="0"> p/sem.</td>
	</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" width="600">
	<tr><td class=fundo valign=top>

<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr>
		<td class=titulo>Início</td>
		<td class=titulo>Término</td>
		<td class=titulo>Vigência</td>
		<td class=titulo>Valor</td>
	</tr>
	<tr>
    	<td class=fundo><input type="text" name="dt_entrada" size="8" value=""></td>
    	<td class=fundo><input type="text" name="dt_termino" size="8" value=""></td>
    	<td class=fundo><input type="text" name="vigencia" size="3" value=""> meses</td>
    	<td class=fundo><input type="text" name="aprov_valor" size="8" value="" class=vr></td>
	</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr>
		<td class=titulo>Observações</td>
	</tr>
	<tr>
    	<td class=fundo><textarea name="observacoes" cols="50" rows="5"></textarea></td>
	</tr>
</table>

</td><td class=fundo>

<table border="0" cellpadding="3" cellspacing="0" width="190">
	<th class=titulo colspan=2>Aprovações</th>
	<tr><td class=titulo>Depto.</td>
    	<td class=fundo><input type="text" name="aprov_depto" size="8" value=""></td>
	</tr>
	<tr><td class=titulo>IFIP</td>
    	<td class=fundo><input type="text" name="aprov_ifip" size="8" value=""></td>
	</tr>
	<tr><td class=titulo>Pró-Adm.</td>
    	<td class=fundo><input type="text" name="aprov_proadm" size="8" value=""></td>
	</tr>
	<tr><td class=titulo>CONSEPE</td>
    	<td class=fundo><input type="text" name="aprov_consepe" size="8" value=""></td>
	</tr>
	<tr><td class=titulo>Ciência em</td>
    	<td class=fundo><input type="text" name="aprov_ciencia" size="8" value=""></td>
	</tr>
</table>

	</td></tr>  
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
	<tr>
		<td class=titulo align="center">
			<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
		<td class=titulo align="center">
			<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
		<td class=titulo align="center">
			<input type="button" value="Fechar" class="button" name="Bt_fechar" onClick="top.window.close()"></td>
	</tr>
</table>
</form>
<%
else
'rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" then
	Response.write "<p>Registro salvo.<br>"
	'response.write "<a href='javascript:window.close()'>Fechar Janela</a>"
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>