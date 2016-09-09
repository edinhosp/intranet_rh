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
<title>Inclusão de Empréstimos Consignados</title>
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
	temp=form.data1.value;
	inicio=new Date(temp.substr(6),temp.substr(3,2)-1,temp.substr(0,2));
	temp2=form.data2.value;
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
	Sub parcial_onChange
		parcial=document.form.parcial.value
		if cdbl(parcial)>0 then document.form.data3.value=document.form.data2.value
	End Sub
</script>
</head>
<body>
<script src="../coolmenu/coolmenus_frame.js" type="text/javascript"></script>
<%
dim conexao, conexao2, chapach, rs, rs2, ok
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form<>"" then
		if request.form("bt_salvar")<>"" then
		tudook=1
		'if request.form("salvar")="1" then

if request.form("data")="" or request.form("valor")="" then
	tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe a data ou valor do empréstimo!');</script>"
end if
		sqla = "INSERT INTO emprestimos (chapa, data, "
		if request.form("dt_conv")<>"" then sqla=sqla & "dt_conv, "
		if request.form("dt_assfieo")<>"" then sqla=sqla & "dt_assfieo, "
		if request.form("dt_banco")<>"" then sqla=sqla & "dt_banco, "
		if request.form("obs")<>"" then sqla=sqla & "obs, "
		if request.form("contrato")<>"" then sqla=sqla & "contrato, "
		sqla = sqla & "valor, nprestacoes, vprestacao, venc1, vencu "
		sqla = sqla & " )"
		
		sqlb = " SELECT '" & request.form("chapa") & "'"
		sqlb=sqlb & ",'" & dtaccess(request.form("data")) & "' "
		if request.form("dt_conv")<>""    then sqlb=sqlb & ",'" & dtaccess(request.form("dt_conv")) & "' "
		if request.form("dt_assfieo")<>"" then sqlb=sqlb & ",'" & dtaccess(request.form("dt_assfieo")) & "' "
		if request.form("dt_banco")<>""   then sqlb=sqlb & ",'" & dtaccess(request.form("dt_banco")) & "' "
		if request.form("obs")<>""      then sqlb=sqlb & ",'" & request.form("obs") & "' "
		if request.form("contrato")<>"" then sqlb=sqlb & ",'" & request.form("contrato") & "' "
		sqlb=sqlb & ", " & nraccess(request.form("valor")) & ""
		sqlb=sqlb & ", " & request.form("nprestacoes") & " "
		sqlb=sqlb & ", " & nraccess(request.form("vprestacao")) & " "
		sqlb=sqlb & ",'" & dtaccess(request.form("venc1")) & "' "
		sqlb=sqlb & ",'" & dtaccess(request.form("vencu")) & "' "
		'sqlb=sqlb & ",'" & session("usuariomaster") & "'"
		'sqlb=sqlb & ",getdate()"
		sql = sqla & sqlb
		'response.write sql
		if tudook=1 then conexao.Execute sql, , adCmdText
		'end if
		end if 'request btsalvar
	else 'request.form=""
	end if

if (request.form("bt_salvar")="") or (request.form("bt_salvar")<>"" and tudook=0) then
%>
<form method="POST" action="emprestimo_nova.asp" name="form">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">
<table border="0" cellpadding="3" cellspacing="0" width="400">
	<tr><td class=grupo>Inclusão de Empréstimo Consignado</td></tr>
</table>
<%
'for each strItem in Request.form
'	Response.write stritem & " = " & request.form(stritem) & " "
'next
if request("codigo")<>"" then
	chapa=request("codigo")
elseif request.form("chapa")<>"" then
	chapa=request.form("chapa") 
else
	chapa=""
end if
%>
<!-- -->
<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo>0</td>
	<td class=fundo><input type="text" value="<%=chapa%>" name="chapa" size="5" onchange="chapa1()" onfocus="javascript:window.status='Informe o chapa do funcionário'"></td>
	<td class=fundo>
		<select size="1" name="nome" onchange="nome1()" onfocus="javascript:window.status='Selecione o Nome do Funcionário'" >
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where chapa<'10000' and codsituacao<>'D' order by nome "
'if session("dp_chapa")<>"" then sql2=sql2 & "and chapa='" & session("dp_chapa") & "'" else sql2=sql2 & "order by nome"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value=''>Selecione o funcionário....</option>"
rs2.movefirst:do while not rs2.eof
if chapa=rs2("chapa") then temp="selected" else temp=""
%>
		<option value="<%=rs2("chapa")%>" <%=temp%>><%=rs2("nome")%></option>
<%
rs2.movenext:loop
rs2.close
%>
		</select></td>
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
<!-- onFocus="javascript:vDateType='3'" onKeyUp="DateFormat(this,this.value,event,false,'3')" onBlur="DateFormat(this,this.value,event,true,'3')" -->
<tr>
	<td class=fundo><input type="text" name="data"        size="8" value="<%=request.form("data")%>" ></td>
	<td class=fundo><input type="text" name="valor"       size="7" value="<%=request.form("valor")%>"></td>
	<td class=fundo><input type="text" name="nprestacoes" size="3" value="<%=request.form("nprestacoes")%>" ></td>
	<td class=fundo><input type="text" name="vprestacao"  size="6" value="<%=request.form("vprestacao")%>"></td>
	<td class=fundo><input type="text" name="venc1"       size="8" value="<%=request.form("venc1")%>" ></td>
	<td class=fundo><input type="text" name="vencu"       size="8" value="<%=request.form("vencu")%>" ></td>
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
	<td class=fundo><input type="text" name="contrato" size="7" value="<%=request.form("contrato")%>" ></td>
	<td class=fundo><input type="text" name="dt_conv" size="8" value="<%=request.form("dt_conv")%>" ></td>
	<td class=fundo><input type="text" name="obs" size="40" value="<%=request.form("obs")%>" ></td>
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
	<td class=fundo><input type="text" name="dt_assfieo" size="8" value="<%=request.form("dt_assfieo")%>" ></td>
	<td class=fundo><input type="text" name="dt_banco" size="8" value="<%=request.form("dt_banco")%>" ></td>
	<td class=fundo><%%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="400">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
	</td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2" onfocus="javascript:window.status='Clique para desfazer e limpar a tela'"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()" onfocus="javascript:window.status='Clique aqui para fechar sem salvar'"></td>
</tr>
</table>
</form>
<%
'else
'rs.close
set rs=nothing
end if   'request.form=""
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if
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