<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 80000
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
if session("a45")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Query</title>
<link rel="stylesheet" type="text/css" href="<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="VBScript">
	Sub library_onChange
'		temp=document.form.library.value
'		tipo=left(temp,1)
'		temp=right(temp,len(temp)-1)
'		document.form.textosql.value=temp
'		if tipo="A" then
'			 'tipo="access" else tipo="sql"
'			document.form.bancoa.checked=true
'			document.form.bancos.checked=false
'		else
'			document.form.bancos.checked=true
'			document.form.bancoa.checked=false
'		end if
		'ok=true:dim frm:set frm=document.form
		'if ok=true then frm.submit
	End Sub
</script>
<script language="JavaScript" type="text/javascript"><!--
function nome1() {	form.chapa.value=form.nome.value; }
function chapa1() {	form.nome.value=form.chapa.value; }
function bancos1() {
	temp=form.bancos.checked
	form.bancoa.checked=false
	form.bancos.checked=true
}
function bancoa1() {
	temp=form.bancoa.checked
	form.bancoa.checked=true
	form.bancos.checked=false
}
function aumenta() {
	temp=form.bancoa.checked
	form.bancoa.checked=true
	form.bancos.checked=false
}
function library1() {
	temp=form.library.value
	tipo=temp.substring(0,1)
	temp=temp.substring(1,temp.length)
	form.textosql.value=temp
	if (tipo=='A') {
	form.bancoa.checked=true
	form.bancos.checked=false 
	} else {
	form.bancoa.checked=false
	form.bancos.checked=true
	}
}	
--></script>
</head>
<body>
<%
dim conexao, conexao1, rs, rsquery
set conexao=server.createobject ("ADODB.Connection")
conexao.open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
set rs.ActiveConnection = conexao

if request.form<>"" then 
	session("querytexto")=request.form("textosql")
	session("querybd")=request.form("banco")
end if

if session("querybd")="access" then bd1="checked" else bd1=""
if session("querybd")="sql"    then bd2="checked" else bd2=""
if request.form("bancoa")="ON" then bd1="checked" else bd1=""
if request.form("bancos")="ON"  then bd2="checked" else bd2=""
if request.form("nlinhas")="" then vlinha="10" else vlinha=request.form("nlinhas")
'if bd1="" and bd2="" then bd2="checked"
%>
<p style="margin-top: 0; margin-bottom: 0" class="realce">Query</p>
<p style="margin-top: 0; margin-bottom: 0" align="left">Para consultas do tipo ação (Create, Update, Alter, Delete, etc) desmarque a caixa de seleção abaixo.</p>
<form method="POST" action="query.asp" name="form">
<p style="margin-top: 0; margin-bottom: 0"><input type="checkbox" name="Select" value="ON" checked tabindex="1">Select	
<!--
					<input type="radio" name="banco" value="sql"    <%=bd2%>> sql 
					<input type="radio" name="banco" value="access" <%=bd1%>> access -->
					<input type="checkbox" name="bancos" value="ON" <%=bd2%> onchange="bancos1()"> sql 
					<input type="checkbox" name="bancoa" value="ON" <%=bd1%> onchange="bancoa1()"> access 
					<input type="texto" name="nlinhas" size=2 value="<%=vlinha%>" > linhas
<%
sqls="select descricao,tipo,sqltexto from bd_query order by descricao"
rs.Open sqls, ,adOpenStatic, adLockReadOnly
%>
<select name="library" onchange="library1()">
	<option value="">Selecione uma query</option>
<%
rs.movefirst
do while not rs.eof
if rs("tipo")&rs("sqltexto")=request.form("library") then temp1="selected" else temp1=""
%>	
	<option value="<%=rs("tipo")&rs("sqltexto")%>" <%=temp1%> ><%=rs("descricao")%></option>
<%
rs.movenext
loop
rs.close
if day(now)=22 and month(now)=8 and session("usuariomaster")="02379" then linhas=20 else linhas=7
%>
</select>
</p>
<p style="margin-top: 0; margin-bottom: 0"><textarea rows="<%=vlinha%>" class="p" name="textosql" cols="120"><%=session("querytexto")%></textarea></p>
<p style="margin-top: 0; margin-bottom: 0"><input type="submit" value="Executar" name="B1" class=button></p>
</form>

<%
if request.form<>"" then
	if request.form("bancos")="ON" then
		conexao.close
		'conexao.open Application("Consql")
		conexao.open Application("conexao")
	else
		conexao.close
		conexao.open Application("conexao")
	end if
	set rsquery=server.createobject ("ADODB.Recordset")
	set rsquery.ActiveConnection = conexao
	sql=request.form("textosql")
	session("querytexto")=sql
	if request.form("Select")="ON" or UCASE(left(sql,6))="SELECT" then
		rsquery.Open sql, ,adOpenStatic, adLockReadOnly
	else
		conexao.Execute Sql, , adCmdText
		SendIp=request.servervariables("REMOTE_ADDR")
		hoje = day(now) & "/" & month(now) & "/" & year(now)
		Sql=Replace(Sql,chr(39),chr(34))
	end if
	if request.form("Select")="ON" or ucase(left(sql,6))="SELECT" then 
	if not rsquery.eof then
		response.write "<table border='1' bordercolor='#CCCCCC' width='100%' cellpadding='0' cellspacing='0' style='border-collapse: collapse'><tr>"
		for x=0 to rsquery.fields.count-1
			response.write "<td class=titulor>" & rsquery.fields(x).name & "</td>"
		next
		response.write "</tr>"
		rsquery.movefirst
		registro=0
		do while not rsquery.eof
			response.write "<tr>"
			for x=0 to rsquery.fields.count-1
				response.write "<td class=""campor"">" 
				if rsquery.fields(x).value="" or isnull(rsquery.fields(x).value) then response.write "&nbsp;" else response.write rsquery.fields(x).value 
				'response.write rsquery.fields(x).value 
				response.write "</td>"
			next
			response.write "</tr>"
			rsquery.movenext
			registro=registro+1
		loop
		response.write "<tr><td class=grupo colspan="&rsquery.fields.count&">"&registro&" registros </td></tr>"
		response.write "</table>"
	else
		response.write "<p class=realce>Sem registros"
	end if
rsquery.close
'conexao.close
set rsquery=nothing
end if
end if
%>
</body>
</html>
<%
'set conexao=nothing
%>