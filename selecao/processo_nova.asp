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
<title>Inclusão de Agendamento</title>
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

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		sql = "INSERT INTO rs_agenda (id_candidato, processo, processo_data, processo_hora, observacoes "
		sql = sql & ") "
		sql2 = " SELECT " & request.form("id_candidato") & ", '" & _
		request.form("processo") & "', " 
		if request.form("processo_data")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("processo_data")) & "', "
		if request.form("processo_hora")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & request.form("processo_hora") & "', "
		sql2=sql2 & "'" & request.form("observacoes") & "' "
		sql1 = sql & sql2 & ""
		'response.write sql1
		if tudook=1 then conexao.Execute sql1, , adCmdText
	end if
else 'request.form=""
	
end if

if request.form="" or (request.form<>"" and tudook=0) then
%>
<form method="POST" action="processo_nova.asp">
<input type="hidden" name="id_candidato" size="4" value="<%=request("codigo")%>">
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
%>
          <option value="<%=rs2("codigo")%>"><%=rs2("processo")%></option>
<%
rs2.movenext:loop
rs2.close
set rs2=nothing
%>
        </select>	  
	  </td>
      <td class=titulo><input type="text" name="processo_data" size="8" value="<%=formatdatetime(now,2)%>"></td>
      <td class=titulo><input type="text" name="processo_hora" size="8" value="<%=formatdatetime(now,4)%>"></td>
    </tr>
  </table>
  <table border="0" cellpadding="3" cellspacing="0" width="500">
    <tr>
      <td class=titulo>Observações</td>
    </tr>
    <tr>
      <td class=titulo><input type="text" name="observacoes"  size="75" value=""></td>
    </tr>
  </table>
  <table border="0" cellpadding="3" cellspacing="0" width="500">
    <tr>
      <td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar">
      </td>
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