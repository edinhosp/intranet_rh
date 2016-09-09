<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a75")="N" or session("a75")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Gerador de etiquetas</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"><!--
function library1() {
	temp=form.descricao.value
	tipo=temp.substring(0,1)
	temp=temp.substring(0,temp.length)
	form.textosql.value=temp
}	
--></script>
<script language="VBScript">
</script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsquery=server.createobject ("ADODB.Recordset")
set rsquery.ActiveConnection = conexao

pixel=96/2.54
point=72/2.54
pointp=72.27/2.54
vlinha=13
session("querytexto")=request.form("sqltexto")
session("querydescricao")=request.form("descricao")

if request.form<>"" then
	if request.form("S2")<>"" then
		if request.form("descricao")="" then 'atualizar
			sql="update etiqueta_query set " & _
			"sqltexto      ='" & replace(request.form("sqltexto"),"'","''") & "' where id_etiq=" & request.form("id_etiq")
			'response.write "<br>" & sql
			conexao.execute sql
		else 'salvar nova
			sql="insert into etiqueta_query (descricao, sqltexto) " & _
			"select '" & request.form("descricao") & "'" & _
			",'" & replace(request.form("sqltexto"),"'","''") & "'"
			'response.write "<br>" & sql
			conexao.execute sql
		end if
	end if 's1

	if request.form("E2")<>"" then
		response.write "<form method='POST' action='etiqueta3.asp' name='form'>"
		response.write "<input type=hidden name=msuperior    value=" & request.form("msuperior") & ">"
		response.write "<input type=hidden name=msuperiorp   value=" & request.form("msuperiorp") & ">"
		response.write "<input type=hidden name=altura       value=" & request.form("altura") & ">"
		response.write "<input type=hidden name=mdireita     value=" & request.form("mdireita") & ">"
		response.write "<input type=hidden name=largura      value=" & request.form("largura") & ">"
		response.write "<input type=hidden name=mesquerdap   value=" & request.form("mesquerdap") & ">"
		response.write "<input type=hidden name=mesquerda    value=" & request.form("mesquerda") & ">"
		response.write "<input type=hidden name=espacolinha  value=" & request.form("espacolinha") & ">"
		response.write "<input type=hidden name=espacocoluna value=" & request.form("espacocoluna") & ">"
		response.write "<input type=hidden name=linhas       value=" & request.form("linhas") & ">"
		response.write "<input type=hidden name=colunas      value=" & request.form("colunas") & ">"
		response.write "<input type=hidden name=sqltexto     value=""" & request.form("sqltexto") & """>"
		response.write "<input type=submit name=e1 class=button value='Clique para visualizar etiquetas'>"
		response.write "</form>"	
	end if	

end if

if request.form("e2")="" then
%>
<form method="POST" action="etiqueta2.asp" name="form">
<input type="hidden" name="colunas"      value="<%=request.form("colunas")%>">
<input type="hidden" name="linhas"       value="<%=request.form("linhas")%>">
<input type="hidden" name="msuperior"    value="<%=request.form("msuperior")%>">
<input type="hidden" name="mdireita"     value="<%=request.form("mdireita")%>">
<input type="hidden" name="mesquerda"    value="<%=request.form("mesquerda")%>">
<input type="hidden" name="altura"       value="<%=request.form("altura")%>">
<input type="hidden" name="largura"      value="<%=request.form("largura")%>">
<input type="hidden" name="espacolinha"  value="<%=request.form("espacolinha")%>">
<input type="hidden" name="espacocoluna" value="<%=request.form("espacocoluna")%>">
<input type="hidden" name="msuperiorp"   value="<%=request.form("msuperiorp")%>">
<input type="hidden" name="mesquerdap"   value="<%=request.form("mesquerdap")%>">

<p class=titulo>Gerador de Etiquetas - Etapa 2/3</p>
<%
sqls="select id_etiq, descricao, sqltexto from etiqueta_query order by descricao"
rs.Open sqls, ,adOpenStatic, adLockReadOnly
%>
<table border="0" bordercolor="#cccccc" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=500>
<tr><td class=campo valign=top>
<select name="id_etiq" size="10" onchange="javascript:submit()">
	<option value=0>Selecione uma query</option>
<%
rs.movefirst:do while not rs.eof
if rs("id_etiq")=cint(request.form("id_etiq")) then texto1="selected" else texto1=""
%>	
	<option value="<%=rs("id_etiq")%>" <%=texto1%> ><%=rs("descricao")%></option>
<%
rs.movenext:loop
rs.close

if request.form("id_etiq")="" then id_etiq=0 else id_etiq=request.form("id_etiq")
if request.form("id_etiqant")="" then id_etiqant=0 else id_etiqant=request.form("id_etiqant")
if cint(id_etiq)<>cint(id_etiqant) then
	sql="SELECT id_etiq, descricao, sqltexto " & _
	"FROM etiqueta_query where id_etiq=" & id_etiq
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
		sqltexto    =rs("sqltexto")
	end if
	rs.close
else 'request.form("idetiqueta")=request.form("idetiquetaant")
	sqltexto =session("querytexto")
end if
%>
</select>
<%
'response.write request.form("id_etiq") & "/" & request.form("id_etiqant") & "<br>"
'response.write id_etiq & "/" & id_etiqant & "<br>"
%>
<input type="hidden" name="id_etiqant" value="<%=request.form("id_etiq")%>">

</td><td class=campo valign=top>
<textarea rows="<%=vlinha%>" class="p" name="sqltexto" cols="100"><%=sqltexto%></textarea>
</td></tr></table>

<p style="margin-top: 0; margin-bottom: 0">
<input type="submit" value="Visualizar amostra dos dados" name="S1" class=button>
Nome do novo script SQL: <input type="text" name="descricao" value="<%=descricao%>" size=40 tabindex=10>
<input type="submit" class=button value="Salvar script" name="S2">
</p>

<table border="0" bordercolor="#cccccc" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=500>
<%
pag_etiqueta=cdbl(request.form("colunas")*request.form("linhas"))

if request.form("s1")<>"" then
	sql=request.form("sqltexto")
	session("querytexto")=sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	totalregistros=rs.recordcount
	response.write totalregistros
	paginas=int(totalregistros/pag_etiqueta)+1
	rs.close
	rsquery.Open sql, ,adOpenStatic, adLockReadOnly
	if totalregistros<10 then totalreg=totalregistros else totalreg=10
	if not rsquery.eof then
		response.write "<table border='1' bordercolor='#CCCCCC' cellpadding='0' cellspacing='0' style='border-collapse: collapse'><tr>"
		response.write "<tr><td class=grupor colspan=" & rsquery.fields.count & ">Amostra dos dados</td></tr>"
		for x=0 to rsquery.fields.count-1
			response.write "<td class=titulor>" & rsquery.fields(x).name & "</td>"
		next
		response.write "</tr>"
		rsquery.movefirst
		for l=1 to totalreg '10
			response.write "<tr>"
			for x=0 to rsquery.fields.count-1
				response.write "<td class="campor">" 
				if rsquery.fields(x).value="" or isnull(rsquery.fields(x).value) then response.write "&nbsp;" else response.write rsquery.fields(x).value 
				'response.write rsquery.fields(x).value 
				response.write "</td>"
			next
			response.write "</tr>"
			rsquery.movenext
		next
		response.write "<tr><td class=grupo colspan="&rsquery.fields.count&">"&totalregistros&" registros na seleção (" & paginas & " folhas de etiquetas) </td></tr>"
		response.write "</table>"
	else
		response.write "<p class=realce>Sem registros"
	end if
	rsquery.close
	set rsquery=nothing
else
		response.write "<table border='1' bordercolor='#CCCCCC' cellpadding='0' cellspacing='0' style='border-collapse: collapse'><tr>"
		response.write "<tr><td class=grupor colspan=5>Amostra dos dados</td></tr>"
		for l=1 to 10
			response.write "<tr>"
			for x=0 to 4
				response.write "<td class="campor">" 
				response.write "&nbsp;"
				response.write "</td>"
			next
			response.write "</tr>"
		next
		response.write "<tr><td class=grupo colspan=5> Sem registros na seleção (0 folha de etiquetas) </td></tr>"
		response.write "</table>"
end if 'request s1

%>

<br>
<input type="submit" class=button value="Continuar para próxima etapa" name="E2" >
</form>
<%

end if 'request.form e2

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>