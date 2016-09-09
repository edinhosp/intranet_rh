<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a88")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Atestado Médico</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript" src="../date.js"></script>
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
function mand_ini1(muda) {
	temp=form.dtinigozo.value;
	inicio=new Date(temp.substr(6),temp.substr(3,2)-1,temp.substr(0,2));
	temp2=form.dtfimgozo.value;
	termino=new Date(temp2.substr(6),temp2.substr(3,2)-1,temp2.substr(0,2));
	dinicio=montharray[inicio.getMonth()]+" "+inicio.getDate()+", "+inicio.getFullYear()
	dfinal=montharray[termino.getMonth()]+" "+termino.getDate()+", "+termino.getFullYear()
	dias=(Math.round((Date.parse(dfinal)-Date.parse(dinicio))/(24*60*60*1000))*1)+1
	document.form.dias.value=dias
}
--></script>
<script language="VBScript">
	Sub data2_onChange
		data2=document.form.data2.value
		data1=document.form.data1.value
		dias=cdate(data2)-cdate(data1)+1
		document.form.dias.value=dias
	End Sub
	Sub dias_onChange
		dias=document.form.dias.value
		document.form.data2.value=dateadd("d",cint(dias)-1,formatdatetime(document.form.data1.value,2))
		diasem=weekday(document.form.data2.value)
		if diasem=7 then diar=2 else diar=1
		document.form.data3.value=dateadd("d",diar,formatdatetime(document.form.data2.value,2))
	End Sub
</script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1

if request.form("data1")="" or request.form("data2")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe as datas do afastamento e término do atestado médico!');</script>"
end if
		dias=request.form("dias")
		'if request.form("tipo")="P" then dias=0
		sql="UPDATE atestados SET "
		if request.form("data1")<>"" then sql=sql & "data1='" & dtaccess(request.form("data1")) & "', "
		if request.form("data2")<>"" then sql=sql & "data2='" & dtaccess(request.form("data2")) & "', "
		if request.form("data3")<>"" then sql=sql & "data3='" & dtaccess(request.form("data3")) & "', "
		sql=sql & "dias      = " & dias & " "
		sql=sql & ",cid      ='" & request.form("cid") & "' "
		sql=sql & ",crm      ='" & request.form("crm") & "' "
		sql=sql & ",medico   ='" & request.form("medico") & "' "
		sql=sql & ",clinica  ='" & request.form("clinica") & "' "
		sql=sql & ",parcial  = " & nraccess(request.form("parcial")) & " "
		sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
		sql=sql & ",dataa   =getdate() "
		sql=sql & " WHERE id_atestado=" & session("id_alt_atestado")
		'response.write "<table width=100><tr><td class="campor">" & request.form & "</td></tr></table>"
		response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="UPDATE apont_adm set deletada=-1 WHERE id_adm=" & session("id_alt_adm")
		sql="DELETE FROM atestados WHERE id_atestado=" & session("id_alt_atestado")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
	if request("codigo")="" then
		id_atestado=session("id_alt_atestado")
		id_atestado=request.form("id_atestado")
	else
		id_atestado=request("codigo")
	end if
	sql="select * from atestados where id_atestado=" & id_atestado
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form("bt_salvar")="" and request.form("bt_excluir")="" then
session("id_alt_atestado")=rs("id_atestado")
sqlz="select nome from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "'"
set rsnome=server.createobject ("ADODB.Recordset")
set rsnome=conexao.Execute (sqlz, , adCmdText)
'response.write request.form
%>
<form method="POST" action="atestado_alteracao.asp" name="form">
<input type="hidden" name="id_atestado" size="4" value="<%=rs("id_atestado")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Alteração de Atestado Médico <%=rs("id_atestado")%></td></tr>
</table>

<!-- movimento / passe -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
<input type="hidden" name="chapa" size="5" value="<%=rs("chapa")%>">
	<td class=titulo><%=rs("id_atestado")%></td>
	<td class=titulo><%=rs("chapa")%></td>
	<td class=titulo><%=rsnome("nome")%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Afastamento</td>
	<td class=titulo>Dias</td>
	<td class=titulo>Término</td>
	<td class=titulo>Abono Parcial</td>
	<td class=titulo>Retorno Trabalho</td>
</tr>
<tr>
	<td class=fundo>
<!-- onFocus="javascript:vDateType='3'" onKeyUp="DateFormat(this,this.value,event,false,'3')" onBlur="DateFormat(this,this.value,event,true,'3')" -->
		<input type="text" name="data1" size="9" value="<%=rs("data1")%>"  >
	</td>
	<td class=fundo><input type="text" name="dias" size="3" value="<%=rs("dias")%>"></td>
	<td class=fundo>
		<input type="text" name="data2" size="9" value="<%=rs("data2")%>" >
	</td>
	<td class=fundo><input type="text" name="parcial" size="3" value="<%=rs("parcial")%>"></td>
	<td class=fundo>
		<input type="text" name="data3" onchange="" size="9" value="<%=rs("data3")%>" >
	</td>
</tr>
</table>

<!-- tarifa / quantidade / total -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>CID</td>
	<td class=titulo>CRM</td>
	<td class=titulo>Médico</td>
	<td class=titulo>Clínica</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="cid" size="6" value="<%=rs("cid")%>" ></td>
	<td class=fundo><input type="text" name="crm" size="6" value="<%=rs("crm")%>" ></td>
	<td class=fundo><input type="text" name="medico" size="20" value="<%=rs("medico")%>" ></td>
	<td class=fundo><input type="text" name="clinica" size="20" value="<%=rs("clinica")%>" ></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
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