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
<title>Autorização de Horário fora do padrão - Estagiário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function renovacao1()	{ form.urenovacao.value=form.renovacao_anterior.value;	}
--></script>
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
	extra=request.form("extra"):extra=extra
	aut=request.form("extraaut"):aut=horasave(aut)
	ant=request.form("extraant"):if isnull(ant) or ant="" then ant=0
	if request.form("marc4")="" then marc4=0 else marc4=request.form("marc4")
	if request.form("marc3")="" then marc3=0 else marc3=request.form("marc3")
	if request.form("marc2")="" then marc2=0 else marc2=request.form("marc2")
	if request.form("marc1")="" then marc1=0 else marc1=request.form("marc1")
	alt=aut-ant
	response.write aut & "-" & ant & "-" & alt & "<br>"
	sql="UPDATE est_batfun SET "
	sql=sql & "extraaut =" & aut & " "
	'sql=sql & ",htrab   =htrab+" & alt & " "
	sql=sql & ",htrab   =" & marc4 & "-" & marc3 & "+" & marc2 & "-" & marc1 & "-" & extra & "+" & alt & " "
	if cdbl(extra)=cdbl(aut) then
		'sql=sql & ",ajust1=null ,ajust2=null ,ajust3=null ,ajust4=null ,ajust5=null ,ajust6=null "
		if request.form("ajust1")<>"" then sql=sql & ",ajust1=marc1 "
		if request.form("ajust2")<>"" then sql=sql & ",ajust2=marc2 "
		if request.form("ajust3")<>"" then sql=sql & ",ajust3=marc3 "
		if request.form("ajust4")<>"" then sql=sql & ",ajust4=marc4 "
	else
		if request.form("ajust1")<>"" then sql=sql & ",ajust1=hor1-" & aut & " "
		if request.form("ajust2")<>"" then sql=sql & ",ajust2=hor2+" & aut & " "
		if request.form("ajust3")<>"" then sql=sql & ",ajust3=hor3-" & aut & " "
		if request.form("ajust4")<>"" then sql=sql & ",ajust4=hor4+" & aut & " "
	end if
	sql=sql & ",travar   =1 "
	sql=sql & ",dataaprov=getdate() "
	sql=sql & ",usuario  ='" & session("usuariomaster") & "' "
	'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
	'sql=sql & ",dataa   =now() "
	sql=sql & " WHERE chapa='" & session("idchapa") & "' and data='" & dtaccess(session("iddia")) & "' "
	response.write sql
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	extra=request.form("extra"):extra=extra
	aut=request.form("extraaut"):aut=horasave(aut)
	ant=request.form("extraant"):if isnull(ant) or ant="" then ant=0
	sql="UPDATE est_batfun SET "
	sql=sql & "extraaut =null "
	sql=sql & ",htrab   =htrab-" & aut & " "
	if request.form("ajust1")<>"" then sql=sql & ",ajust1=hor1 "
	if request.form("ajust2")<>"" then sql=sql & ",ajust2=hor2 "
	if request.form("ajust3")<>"" then sql=sql & ",ajust3=hor3 "
	if request.form("ajust4")<>"" then sql=sql & ",ajust4=hor4 "
	sql=sql & ",travar   =0 "
	sql=sql & ",dataaprov='" & now() & "' "
	sql=sql & ",usuario  ='" & session("usuariomaster") & "' "
	'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
	'sql=sql & ",dataa   =getdate() "
	sql=sql & " WHERE chapa='" & session("idchapa") & "' and data='" & dtaccess(session("iddia")) & "' "
	response.write sql
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("chapa")=null or request("chapa")="" then
		idchapa=session("idchapa")
		iddia=session("iddia")
		'if session("idcadhor")="" then id_cadhor=request.form("id_cadhor")
	else
		idchapa=request("chapa")
		iddia=request("data")
	end if
	sqla="select * from est_batfun " & _
	"where chapa='" & idchapa & "' and data='" & dtaccess(iddia) & "' "
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
session("idchapa")=rs("chapa")
session("iddia")=rs("data")
ent1h=int(rs("extra")/60):ent1m=rs("extra")-ent1h*60
sai1h=int(rs("extraaut")/60):sai1m=rs("extraaut")-sai1h*60
if request.form("chapa")=""       then chapa=rs("chapa")             else chapa=request.form("chapa")
if request.form("data")=""        then data=rs("data")               else data=request.form("data")
if request.form("extra")=""       then extra=rs("extra")             else extra=request.form("extra")
if request.form("extraaut")=""    then extraaut=rs("extraaut")       else extraaut=request.form("extraaut")
if extraaut="" or isnull(extraaut) then extraaut=extra
%>
<form method="POST" action="espelhoponto_extraaut.asp" name="form">
<input type="hidden" name="id_chapa" size="5" value="<%=rs("chapa")%>" >  
<input type="hidden" name="id_dia" size="8" value="<%=rs("data")%>" >  
<input type="hidden" name="ajust1" size="4" value="<%=rs("ajust1")%>" >  
<input type="hidden" name="ajust2" size="4" value="<%=rs("ajust2")%>" >  
<input type="hidden" name="ajust3" size="4" value="<%=rs("ajust3")%>" >  
<input type="hidden" name="ajust4" size="4" value="<%=rs("ajust4")%>" >  
<input type="hidden" name="ajust5" size="4" value="<%=rs("ajust5")%>" >  
<input type="hidden" name="ajust6" size="4" value="<%=rs("ajust6")%>" >  
<input type="hidden" name="marc1" size="4" value="<%=rs("marc1")%>" >  
<input type="hidden" name="marc2" size="4" value="<%=rs("marc2")%>" >  
<input type="hidden" name="marc3" size="4" value="<%=rs("marc3")%>" >  
<input type="hidden" name="marc4" size="4" value="<%=rs("marc4")%>" >  
<input type="hidden" name="marc5" size="4" value="<%=rs("marc5")%>" >  
<input type="hidden" name="marc6" size="4" value="<%=rs("marc6")%>" >  
<input type="hidden" name="extra" size="4" value="<%=rs("extra")%>" >  
<input type="hidden" name="extraant" size="4" value="<%=rs("extraaut")%>" >  

<table border="0" cellpadding="3" cellspacing="0" width="250">
<tr><td class=grupo>Autorização - Data <%=session("iddia")%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="250">
<tr>
	<td class=titulo>Data</td>
	<td class=titulo>Extra</td>
	<td class=titulo>Autoriza</td>
</tr>
<tr>
	<td class=titulo align="center"><%=data%></td>
	<td class=fundo align="center"><%=horaload(extra,2)%></td>
	<td class=fundo>
		<input type="text" name="extraaut" size="5" value="<%=horaload(extraaut,2)%>" class="form_input" style="text-align:center"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="250">
<tr><td class=titulo colspan=3>&nbsp;</td></tr>
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="submit" value="Cancelar Autoriz." class="button" name="Bt_excluir"></td>
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
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.refresh;self.close();</script>"
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