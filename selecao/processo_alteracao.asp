<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a68")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Alteração de Agendamento</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

	if request.form("bt_salvar")<>"" then
		tudook=1
		sql="UPDATE rs_agenda SET "
		sql=sql & "processo     ='" & request.form("processo") & "', "
		if request.form("processo_data")=""     then
			sql=sql & "processo_data=null,"
		else
			sql=sql & "processo_data='" & dtaccess(request.form("processo_data")) & "', "
		end if
		if request.form("processo_hora")=""     then
			sql=sql & "processo_hora=null,"
		else
			sql=sql & "processo_hora='" & request.form("processo_hora") & "', "
		end if
		sql=sql & "observacoes  ='" & request.form("observacoes")       & "' "
		sql=sql & "WHERE id_agenda=" & session("id_alt_agenda")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM rs_agenda WHERE id_agenda=" & session("id_alt_agenda")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")=null then
		id_agendao=session("id_alt_agenda")
	else
		id_agenda=request("codigo")
	end if
	sqla="select * from rs_agenda "
	sqlb="where id_agenda=" & id_agenda
	sql1=sqla & sqlb 
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_agenda")=rs("id_agenda")
%>
<form method="POST" action="processo_alteracao.asp">
<input type="hidden" name="id_agenda" size="4" value="<%=rs("id_agenda")%>">
  <table border="0" cellpadding="3" cellspacing="0" width="500">
    <tr><td class=grupo>Agenda para o Candidato: <%=request("candidato")%></td></tr>
  </table>
  <table border="0" cellpadding="3" cellspacing="0" width="500">
    <tr>
      <td class=titulo>Processo</td>
      <td class=titulo>Data</td>
      <td class=titulo>Hora</td>
    </tr>
    <tr>
      <td class=titulo>
	<select size="1" name="processo" class=a>
		<option value="0">Selecione um processo</option>
<%
sql2="select codigo, processo from rs_processo order by codigo "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
if rs("processo")=rs2("codigo") then tmpagenda="selected" else tmpagenda=""
%>
          <option value="<%=rs2("codigo")%>" <%=tmpagenda%>><%=rs2("processo")%></option>
<%
rs2.movenext:loop
rs2.close
set rs2=nothing
%>
        </select>	  
	  </td>
      <td class=titulo><input type="text" name="processo_data" size="8" value="<%=rs("processo_data")%>"></td>
      <td class=titulo><input type="text" name="processo_hora" size="8" value="<%=rs("processo_hora")%>"></td>
    </tr>
  </table>
  <table border="0" cellpadding="3" cellspacing="0" width="500">
    <tr>
      <td class=titulo>Observações</td>
    </tr>
    <tr>
      <td class=titulo><input type="text" name="observacoes" size="75" value="<%=rs("observacoes")%>"></td>
    </tr>
  </table>
  <table border="0" cellpadding="3" cellspacing="0" width="500">
    <tr>
      <td class=titulo align="center">
		<input type="submit" value="Salvar Alterações  " class="button" name="Bt_salvar">
      </td>
      <td class=titulo align="center">
       <input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
      <td class=titulo align="center">
       <input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>
    </tr>
  </table>
</form>
<%
rs.close
set rs=nothing
end if
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