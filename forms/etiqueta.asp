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
function descricao() { form.sqltexto.value=form.descricao.value; }
function sqltexto()  { form.descricao.value=form.sqltexto.value; }
function library1() {
	temp=form2.id_etiq.value
	tipo=temp.substring(0,1)
	temp=temp.substring(0,temp.length)
	form2.textosql.value=temp
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

pixel=96/2.54
point=72/2.54
pointp=72.27/2.54

if request.form<>"" then
	if request.form("S1")<>"" then
		if request.form("etiqueta")="" then 'atualizar
			sql="update etiqueta_cad set " & _
			"msuperior   =" & nraccess(request.form("msuperior")) & _
			",msuperiorp  =" & nraccess(request.form("msuperiorp")) & _
			",altura      =" & nraccess(request.form("altura")) & _
			",mdireita    =" & nraccess(request.form("mdireita")) & _
			",largura     =" & nraccess(request.form("largura")) & _
			",mesquerdap  =" & nraccess(request.form("mesquerdap")) & _
			",mesquerda   =" & nraccess(request.form("mesquerda")) & _
			",espacolinha =" & nraccess(request.form("espacolinha")) & _
			",espacocoluna=" & nraccess(request.form("espacocoluna")) & _
			",linhas      =" & nraccess(request.form("linhas")) & _
			",colunas     =" & nraccess(request.form("colunas")) & _
			" where id_etiq=" & request.form("idetiqueta")
			response.write "<br>" & sql
			conexao.execute sql
		else 'salvar nova
			sql="insert into etiqueta_cad (descricao, msuperior, msuperiorp, altura, mdireita, largura, mesquerdap, mesquerda, espacolinha, espacocoluna, linhas, colunas) " & _
			"select '" & request.form("etiqueta") & "'" & _
			"," & nraccess(request.form("msuperior")) & _
			"," & nraccess(request.form("msuperiorp")) & _
			"," & nraccess(request.form("altura")) & _
			"," & nraccess(request.form("mdireita")) & _
			"," & nraccess(request.form("largura")) & _
			"," & nraccess(request.form("mesquerdap")) & _
			"," & nraccess(request.form("mesquerda")) & _
			"," & nraccess(request.form("espacolinha")) & _
			"," & nraccess(request.form("espacocoluna")) & _
			"," & nraccess(request.form("linhas")) & _
			"," & nraccess(request.form("colunas"))
			response.write "<br>" & sql
			conexao.execute sql
		end if
	end if 's1

	if request.form("E1")<>"" then
		response.write "<form method='POST' action='etiqueta2.asp' name='form'>"
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
		response.write "<input type=submit name=e1 class=button value='Clique para selecionar dados'>"
		response.write "</form>"	
	end if	
end if

if request.form("E1")="" then 
%>
<form method="POST" action="etiqueta.asp" name="form">
<p class=titulo>Gerador de Etiquetas - Etapa 1/3</p>


<table border="0" bordercolor="#cccccc" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=500>
<tr><td class=grupo colspan=2>Configurar etiquetas 
	<select name="idetiqueta" size="1" onchange="javascript:submit()"><option value="0">Selecione um modelo</option>
<%
sql="select * from etiqueta_cad order by descricao"
rs.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs.eof
if rs("id_etiq")=cint(request.form("idetiqueta")) then texto1="selected" else texto1=""
%>
	<option value="<%=rs("id_etiq")%>" <%=texto1%>  ><%=rs("descricao")%></option>
<%
rs.movenext:loop
rs.close

if request.form("idetiqueta")="" then idetiqueta=0 else idetiqueta=request.form("idetiqueta")
if request.form("idetiquetaant")="" then idetiquetaant=0 else idetiquetaant=request.form("idetiquetaant")
if cint(idetiqueta)<>cint(idetiquetaant) then
	sql="SELECT id_etiq, descricao, msuperior, msuperiorp, altura, mdireita, largura, mesquerdap, mesquerda, espacolinha, espacocoluna, colunas, linhas " & _
	"FROM etiqueta_cad where id_etiq=" & idetiqueta
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
		colunas     =rs("colunas")
		linhas      =rs("linhas")
		msuperior   =rs("msuperior")
		mdireita    =rs("mdireita")
		mesquerda   =rs("mesquerda")
		altura      =rs("altura")
		largura     =rs("largura")
		espacolinha =rs("espacolinha")
		espacocoluna=rs("espacocoluna")
		msuperiorp  =rs("msuperiorp")
		mesquerdap  =rs("mesquerdap")
	end if
	rs.close
else 'request.form("idetiqueta")=request.form("idetiquetaant")
	colunas     =request.form("colunas")
	linhas      =request.form("linhas")
	msuperior   =request.form("msuperior")
	mdireita    =request.form("mdireita")
	mesquerda   =request.form("mesquerda")
	altura      =request.form("altura")
	largura     =request.form("largura")
	espacolinha =request.form("espacolinha")
	espacocoluna=request.form("espacocoluna")
	msuperiorp  =request.form("msuperiorp")
	mesquerdap  =request.form("mesquerdap")
end if

%>
	</select>
<input type="hidden" name="idetiquetaant" value="<%=request.form("idetiqueta")%>">
</td>
</tr>
<tr><td class=fundop>Nº Colunas</td>	
	<td class=fundop><input type="text" size="5" name="colunas" value="<%=colunas%>"></td>
</tr>
<tr><td class=fundop>Nº Linhas</td>	
	<td class=fundop><input type="text" size="5" name="linhas" value="<%=linhas%>"></td>
</tr>
</table>	
<%

larg=496:alt=191
c1=0.30:c2=2.00:c3=0.90:c4=0.85:c5=0.95:c6=0.75:c7=0.80:c8=0.95:c9=1.00:c10=0.80:c11=1.00:c12=0.65:c13=(larg-(c1+c2+c3+c4+c5+c6+c7+c8+c9+c10+c11+c12)*pixel)/pixel
l1=0.33:l2=0.60:l3=0.33:l4=0.57:l5=0.38:l6=0.25:l7=0.35:l8=0.18:l9=1.00:l10=(alt-(l1+l2+l3+l4+l5+l6+l7+l8+l9)*pixel)/pixel
'response.write (c1+c2+c3+c4+c5+c6+c7+c8+c9+c10+c11)*pixel
%>
<table border="0" bordercolor="red" cellpadding="1" width=500 height=200 cellspacing="0" style="background-color:transparent;border-collapse: collapse;background:transparent url(../images/fundo_etiqueta.jpg) ;">
<tr>
	<td height="<%=l1*pixel%>px" width="<%=c1*pixel%>px" style="background-color:transparent"></td>
	<td width="<%=c2*pixel%>px" style="background-color:transparent"></td>	<td width="<%=c3*pixel%>px" style="background-color:transparent"></td>
	<td width="<%=c4*pixel%>px" style="background-color:transparent"></td>	<td width="<%=c5*pixel%>px" style="background-color:transparent"></td>
	<td width="<%=c6*pixel%>px" style="background-color:transparent"></td>	<td width="<%=c7*pixel%>px" style="background-color:transparent"></td>
	<td width="<%=c8*pixel%>px" style="background-color:transparent"></td>	<td width="<%=c9*pixel%>px" style="background-color:transparent"></td>
	<td width="<%=c10*pixel%>px" style="background-color:transparent"></td>	<td width="<%=c11*pixel%>px" style="background-color:transparent"></td>
	<td width="<%=c12*pixel%>px" style="background-color:transparent"></td>	<td width="<%=c13*pixel%>px" style="background-color:transparent"></td>
</tr>
<tr>
	<td height="<%=l2*pixel%>px" width="<%=c1*pixel%>px" style="background-color:transparent"></td>
	<td colspan=2 width="<%=c2*pixel%>px" style="background-color:transparent"></td>
	<td colspan=2 width="<%=c4*pixel%>px" style="background-color:transparent" valign=top>
		<input type="text" size=6 class="help_input" value="<%=msuperior%>" tabindex=1 name="msuperior"></td>
	<td colspan=8 width="<%=c6*pixel%>px" style="background-color:transparent"></td>
</tr>
<tr>
	<td height="<%=l3*pixel%>px" width="<%=c1*pixel%>px" style="background-color:transparent"></td>
	<td colspan=8 width="<%=c2*pixel%>px" style="background-color:transparent"></td>
	<td colspan=2 rowspan=2 width="<%=c10*pixel%>px" style="background-color:transparent" valign=top>
		<input type="text" size=6 class="help_input" value="<%=msuperiorp%>" tabindex=6 name="msuperiorp"></td>
	<td colspan=2 width="<%=c12*pixel%>px" style="background-color:transparent"></td>
</tr>
<tr>
	<td height="<%=l4*pixel%>px" width="<%=c1*pixel%>px" style="background-color:transparent"></td>
	<td width="<%=c2*pixel%>px" style="background-color:transparent"></td>
	<td colspan=2 width="<%=c3*pixel%>px" style="background-color:transparent" valign=top>
		<input type="text" size=6 class="help_input" value="<%=altura%>" tabindex=3 name="altura"></td>
	<td colspan=5 width="<%=c5*pixel%>px" style="background-color:transparent"></td>
	<td colspan=2 width="<%=c12*pixel%>px" style="background-color:transparent"></td>
</tr>
<tr>
	<td height="<%=l5*pixel%>px" width="<%=c1*pixel%>px" style="background-color:transparent"></td>
	<td colspan=11 width="<%=c2*pixel%>px" style="background-color:transparent"></td>
	<td rowspan=2 width="<%=c13*pixel%>px" style="background-color:transparent" valign=top>
		<input type="text" size=6 class="help_input" value="<%=mdireita%>" tabindex=9 name="mdireita"></td>
</tr>
<tr>
	<td height="<%=l6*pixel%>px" width="<%=c1*pixel%>px" style="background-color:transparent"></td>
	<td colspan=3 width="<%=c2*pixel%>px" style="background-color:transparent"></td>
	<td colspan=2 rowspan=2 width="<%=c5*pixel%>px" style="background-color:transparent" valign=top>
		<input type="text" size=6 class="help_input" value="<%=largura%>" tabindex=4 name="largura"></td>
	<td colspan=6 width="<%=c7*pixel%>px" style="background-color:transparent"></td>
</tr>
<tr>
	<td height="<%=l7*pixel%>px" width="<%=c1*pixel%>px" style="background-color:transparent"></td>
	<td colspan=3 width="<%=c2*pixel%>px" style="background-color:transparent"></td>
	<td width="<%=c7*pixel%>px" style="background-color:transparent"></td>
	<td colspan=2 rowspan=3 width="<%=c8*pixel%>px" style="background-color:transparent" valign=top>
		<input type="text" size=6 class="help_input" value="<%=mesquerdap%>" tabindex=7 name="mesquerdap"></td>
	<td colspan=4 width="<%=c10*pixel%>px" style="background-color:transparent"></td>
</tr>
<tr>
	<td height="<%=l8*pixel%>px" width="<%=c1*pixel%>px" style="background-color:transparent"></td>
	<td rowspan=2 width="<%=c2*pixel%>px" style="background-color:transparent" valign=top>
		<input type="text" size=6 class="help_input" value="<%=mesquerda%>" tabindex=2 name="mesquerda"></td>
	<td colspan=5 width="<%=c3*pixel%>px" style="background-color:transparent"></td>
	<td colspan=4 width="<%=c10*pixel%>px" style="background-color:transparent"></td>
</tr>
<tr>
	<td height="<%=l9*pixel%>px" width="<%=c1*pixel%>px" style="background-color:transparent"></td>
	<td colspan=5 width="<%=c3*pixel%>px" style="background-color:transparent"></td>
	<td width="<%=c10*pixel%>px" style="background-color:transparent"></td>
	<td colspan=3 width="<%=c11*pixel%>px" style="background-color:transparent" valign=top>
		<input type="text" size=6 class="help_input" value="<%=espacolinha%>" tabindex=8 name="espacolinha"></td>
</tr>
<tr>
	<td height="<%=l10*pixel%>px" width="<%=c1*pixel%>px" style="background-color:transparent"></td>
	<td colspan=5 width="<%=c2*pixel%>px" style="background-color:transparent"></td>
	<td colspan=2 width="<%=c7*pixel%>px" style="background-color:transparent" valign=top>
		<input type="text" size=6 class="help_input" value="<%=espacocoluna%>" tabindex=5 name="espacocoluna"></td>
	<td colspan=5 width="<%=c9*pixel%>px" style="background-color:transparent"></td>
</tr>
</table>
<table border="0" bordercolor="#cccccc" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=500>
<tr><td class=grupo>Instruções:</td></tr>
<tr><td class=fundo>
1º - informe as medidas em centímetros<br>
2º - as margems superior e esquerda devem ser configuradas no seu navegador (Internet Explorer/Firefox/etc)<br>
3º - normalmente folhas de etiquetas são do tamanho Carta (8,5" x 11"). Configure o tamanho da página também.
</td</tr>
</tr>
<tr>
	<td class=titulo colspan=3>
	</td>
</tr>
</table>
Nome da etiqueta nova: <input type="text" name="etiqueta" value="" size=30 tabindex=10>
<input type="submit" class=button value="Salvar etiqueta" name="S1">
<br>
<input type="submit" class=button value="Continuar para próxima etapa" name="E1">
</form>
<%
end if 'request.form e1

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>