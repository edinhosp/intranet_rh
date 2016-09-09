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
<title>Alteração de Processo IFIP</title>
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
		sql="UPDATE ifip_cadastro SET "
		sql=sql & "num_processo   ='" & request.form("num_processo") & "', "
		sql=sql & "status         ='" & request.form("status") & "', "
		sql=sql & "titulo_pesquisa='" & request.form("titulo_pesquisa") & "', "
		sql=sql & "linha_pesquisa ='" & request.form("linha_pesquisa") & "', "
		sql=sql & "area_conhecimento='" & request.form("area_conhecimento") & "', "
		sql=sql & "horas_semanais = " & nraccess(cdbl(request.form("horas_semanais"))) & ", "
		if request.form("dt_entrada")=""     then
			sql=sql & "dt_entrada=null,"
		else
			sql=sql & "dt_entrada='" & dtaccess(request.form("dt_entrada")) & "', "
		end if
		if request.form("dt_termino")=""     then
			sql=sql & "dt_termino=null,"
		else
			sql=sql & "dt_termino='" & dtaccess(request.form("dt_termino")) & "', "
		end if
		if request.form("vigencia")=""     then
			sql=sql & "vigencia=null,"
		else
			sql=sql & "vigencia=" & nraccess(cdbl(request.form("vigencia"))) & ", "
		end if
		if request.form("aprov_valor")=""     then
			sql=sql & "aprov_valor=null,"
		else
			sql=sql & "aprov_valor=" & nraccess(cdbl(request.form("aprov_valor"))) & ", "
		end if
		sql=sql & "observacoes    ='" & request.form("observacoes") & "', "
		if request.form("aprov_depto")=""     then
			sql=sql & "aprov_depto=null,"
		else
			sql=sql & "aprov_depto='" & dtaccess(request.form("aprov_depto")) & "', "
		end if
		if request.form("aprov_ifip")=""     then
			sql=sql & "aprov_ifip=null,"
		else
			sql=sql & "aprov_ifip='" & dtaccess(request.form("aprov_ifip")) & "', "
		end if
		if request.form("aprov_consepe")=""     then
			sql=sql & "aprov_consepe=null,"
		else
			sql=sql & "aprov_consepe='" & dtaccess(request.form("aprov_consepe")) & "', "
		end if
		if request.form("aprov_proadm")=""     then
			sql=sql & "aprov_proadm=null,"
		else
			sql=sql & "aprov_proadm='" & dtaccess(request.form("aprov_proadm")) & "', "
		end if
		if request.form("aprov_ciencia")=""     then
			sql=sql & "aprov_ciencia=null "
		else
			sql=sql & "aprov_ciencia='" & dtaccess(request.form("aprov_ciencia")) & "' "
		end if

		sql=sql & " WHERE id_ifip=" & session("id_alt_ifip")
		'response.write sql
		conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		sql="DELETE FROM ifip_cadastro WHERE id_ifip=" & session("id_alt_ifip")
		conexao.Execute sql, , adCmdText
	end if

else 'request.form=""

	if request("codigo")=null then
		id_ifip=session("id_alt_ifip")
	else
		id_ifip=request("codigo")
	end if
	sqla="select * from ifip_cadastro "
	sqlb="where id_ifip=" & id_ifip
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if
%>
<%
if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_ifip")=rs("id_ifip")

%>
<form method="POST" action="pesquisa_alteracao.asp" name="form">
<input type="hidden" name="id_ifip" size="4" value="<%=rs("id_ifip")%>" style="font-size: 8 pt" >  
<table border="0" cellpadding="3" cellspacing="0" width="600">
	<tr><td class=grupo>Alteração de Processo IFIP - <%=rs("id_ifip")%></td></tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" width="600">
	<tr><td class=titulo>Nº Processo</td><td class=titulo>Título da Pesquisa</td></tr>
	<tr><td class=fundo><input type="text" name="num_processo" size="7" value="<%=rs("num_processo")%>"></td>
		<td class=fundo rowspan=3>
			<textarea name="titulo_pesquisa" cols="60" rows="5"><%=rs("titulo_pesquisa")%></textarea>
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
if rs("status")=rsc("id_status") then tmpfuncao="selected" else tmpfuncao=""
%>
          <option value="<%=rsc("id_status")%>" <%=tmpfuncao%>><%=rsc("desc_status")%></option>
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
    	<td class=fundo><input type="text" name="linha_pesquisa" size="25" value="<%=rs("linha_pesquisa")%>"></td>
    	<td class=fundo><input type="text" name="area_conhecimento" size="50" value="<%=rs("area_conhecimento")%>"></td>
    	<td class=fundo><input type="text" name="horas_semanais" size="3" value="<%=rs("horas_semanais")%>"> p/sem.</td>
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
    	<td class=fundo><input type="text" name="dt_entrada" size="8" value="<%=rs("dt_entrada")%>"></td>
    	<td class=fundo><input type="text" name="dt_termino" size="8" value="<%=rs("dt_termino")%>"></td>
    	<td class=fundo><input type="text" name="vigencia" size="3" value="<%=rs("vigencia")%>"> meses</td>
    	<td class=fundo><input type="text" name="aprov_valor" size="8" value="<%=rs("aprov_valor")%>" class=vr></td>
	</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr>
		<td class=titulo>Observações</td>
	</tr>
	<tr>
    	<td class=fundo><textarea name="observacoes" cols="50" rows="5"><%=rs("observacoes")%></textarea></td>
	</tr>
</table>

</td><td class=fundo>

<table border="0" cellpadding="3" cellspacing="0" width="190">
	<th class=titulo colspan=2>Aprovações</th>
	<tr><td class=titulo>Depto.</td>
    	<td class=fundo><input type="text" name="aprov_depto" size="8" value="<%=rs("aprov_depto")%>"></td>
	</tr>
	<tr><td class=titulo>IFIP</td>
    	<td class=fundo><input type="text" name="aprov_ifip" size="8" value="<%=rs("aprov_ifip")%>"></td>
	</tr>
	<tr><td class=titulo>Pró-Adm.</td>
    	<td class=fundo><input type="text" name="aprov_proadm" size="8" value="<%=rs("aprov_proadm")%>"></td>
	</tr>
	<tr><td class=titulo>CONSEPE</td>
    	<td class=fundo><input type="text" name="aprov_consepe" size="8" value="<%=rs("aprov_consepe")%>"></td>
	</tr>
	<tr><td class=titulo>Ciência em</td>
    	<td class=fundo><input type="text" name="aprov_ciencia" size="8" value="<%=rs("aprov_ciencia")%>"></td>
	</tr>
</table>

	</td></tr>  
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
	<tr>
		<td class=titulo align="center">
			<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
		<td class=titulo align="center">
			<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
		<td class=titulo align="center">
			<input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
	</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	Response.write "<p>Registro atualizado.<br>"
	'response.write "<a href='javascript:window.close()'>Fechar Janela</a>"
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
<%
end if

set rsc=nothing
set rsnome=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>