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
<title>Alteração de Empréstimo Consignado</title>
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

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1

if request.form("data")="" or request.form("valor")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe a data e valor do empréstimo!');</script>"
end if
		sql="UPDATE emprestimos SET "
		sql=sql & "data='" & dtaccess(request.form("data")) & "', valor=" & nraccess(request.form("valor")) & " "
		sql=sql & ", nprestacoes=" & request.form("nprestacoes") & ", vprestacao=" & nraccess(request.form("vprestacao")) & " "
		sql=sql & ", venc1='" & dtaccess(request.form("venc1")) & "', vencu='" & dtaccess(request.form("vencu")) & "' "
		sql=sql & ", contrato='" & request.form("contrato") & "' "
		if request.form("dt_conv")<>"" then sql=sql & ", dt_conv='" & dtaccess(request.form("dt_conv")) & "' "
		if request.form("dt_assfieo")<>"" then sql=sql & ", dt_assfieo='" & dtaccess(request.form("dt_assfieo")) & "' "
		if request.form("dt_banco")<>"" then sql=sql & ", dt_banco='" & dtaccess(request.form("dt_banco")) & "' "
		sql=sql & ", obs='" & request.form("obs") & "' "
		'sql=sql & ",usuarioa='" & session("usuariomaster") & "' "
		'sql=sql & ",dataa   =getdate() "
		sql=sql & " WHERE idemp=" & session("id_alt_emp")
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM emprestimos WHERE idemp=" & session("id_alt_emp")
		conexao.Execute sql, , adCmdText
	end if
else 'request.form=""
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")="" then
		id_emp=session("id_alt_emp")
		id_emp=request.form("idemp")
	else
		id_emp=request("codigo")
	end if
	sql="select e.*, f.nome from emprestimos e, qry_funcionarios f where f.chapa collate database_default=e.chapa and idemp=" & id_emp
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
session("id_alt_emp")=rs("idemp")
%>
<form method="POST" action="emprestimo_alteracao.asp" name="form">
<input type="hidden" name="idemp" size="4" value="<%=rs("idemp")%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Alteração de Empréstimo Consignado <%=rs("idemp")%></td></tr>
</table>

<!--  -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
<input type="hidden" name="chapa" size="5" value="<%=rs("chapa")%>">
	<td class=titulo><%=rs("idemp")%></td>
	<td class=titulo><%=rs("chapa")%></td>
	<td class=titulo><%=rs("nome")%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Data</td>
	<td class=titulo>Valor</td>
	<td class=titulo># Parc.</td>
	<td class=titulo>$ Parc.</td>
	<td class=titulo>1º Venc.</td>
	<td class=titulo>Ult.Venc.</td>
</tr>
<tr>
<!-- onFocus="javascript:vDateType='3'" onKeyUp="DateFormat(this,this.value,event,false,'3')" onBlur="DateFormat(this,this.value,event,true,'3')" -->
	<td class=fundo><input type="text" name="data"        size="8" value="<%=rs("data")%>" ></td>
	<td class=fundo><input type="text" name="valor"       size="7" value="<%=rs("valor")%>"></td>
	<td class=fundo><input type="text" name="nprestacoes" size="3" value="<%=rs("nprestacoes")%>" ></td>
	<td class=fundo><input type="text" name="vprestacao"  size="6" value="<%=rs("vprestacao")%>"></td>
	<td class=fundo><input type="text" name="venc1"       size="8" value="<%=rs("venc1")%>" ></td>
	<td class=fundo><input type="text" name="vencu"       size="8" value="<%=rs("vencu")%>" ></td>
</tr>
</table>

<!--  -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Contrato</td>
	<td class=titulo>Data Conv.</td>
	<td class=titulo>Obs.</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="contrato" size="7" value="<%=rs("contrato")%>" ></td>
	<td class=fundo><input type="text" name="dt_conv" size="8" value="<%=rs("dt_conv")%>" ></td>
	<td class=fundo><input type="text" name="obs" size="40" value="<%=rs("obs")%>" ></td>
</tr>
</table>

<!--  -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo rowspan=2>Datas de Envio</td>
	<td class=titulo>P/Pro-Reitoria</td>
	<td class=titulo>P/Banco</td>
	<td class=titulo>Status</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="dt_assfieo" size="8" value="<%=rs("dt_assfieo")%>" ></td>
	<td class=fundo><input type="text" name="dt_banco" size="8" value="<%=rs("dt_banco")%>" ></td>
	<td class=fundo><%%></td>
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