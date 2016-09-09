<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if (session("a93")="N" or session("a93")="") and (session("acesso")>2 or session("acesso")="") then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Inclusão de Bibliografia</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1()	{	form.chapa.value=form.nome.value;	}
function chapa1()	{	form.nome.value=form.chapa.value;	}
--></script>
</head>
<body>
<script src="../coolmenu/coolmenus_frame.js" type="text/javascript"></script>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request.form("bt_salvar")<>"" then
	tudook=1
	sql = "INSERT INTO grades_plano_biblio ("
	sql = sql & "id_plano, complementar, ordem, cod_acervo, referencia, usuarioc, datac, status"
	sql = sql & ") "
	if request.form("complementar")="" then complementar=0 else complementar=1
	if request.form("cod_acervo")="" then cod_acervo="null" else cod_acervo=request.form("cod_acervo")
if request.form("cod_acervo")="" and request.form("referencia")="" then tudook=0
	sql2 = " SELECT " & request.form("id_plano") & ", " & complementar & ", " & request.form("ordem") & ""
	sql2=sql2 & ", " & cod_acervo & ""
	sql2=sql2 & ", '" & request.form("referencia") & "'"
	sql2=sql2 & ", '" & session("usuariomaster") & "'"
	sql2=sql2 & ", getdate()"
	if session("perlet_atual_plano")>="2010" then
		status="null"
		if request.form("cod_acervo")="" and request.form("referencia")<>"" then status="'P'"
		sql2=sql2 & ", " & status
	else
		sql2=sql2 & ", null"
	end if 
	sql4 = sql & sql2 & ""
	'response.write "<font size='2'>" & sql4
	if tudook=1 then conexao.Execute sql4, , adCmdText
end if

if request("clique_acervo")<>"" then
	clique_acervo=request("clique_acervo")
else
	clique_acervo=request.form("cod_acervo")
end if

if request("clique_compra")<>"" then
	clique_compra=request("clique_compra")
else
	clique_compra=request.form("referencia")
end if

if request("codigo")<>"" then id_plano=request("codigo") else id_plano=session("insert_id_plano")
session("insert_id_plano")=id_plano
if session("insert_id_plano")="" and id_plano<>"" then session("insert_id_plano")=id_plano
if request("compl")<>"" then compl=request("compl")
'response.write "<br>:" & request("codigo")
'response.write "<br>:" & session("insert_id_plano")
'response.write "<br>:" & id_plano

largura=430
%>
<form method="POST" action="biblio_nova.asp" name="form" >
<input type="hidden" name="id_plano" value="<%=id_plano%>">
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr><td class=grupo>Inclusão de Bibliografia</td></tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo></td>
	<td class=titulo>Ordem</td>
	<td class=titulo width=100%></td>
</tr>
<tr>
	<td class=titulo><font color=gray>Novo</td>
<%
if request.form("complementar")="ON" or compl=1 then complementar="checked" else complementar=""
if complementar="checked" then comple=1 else comple=0
sqlo="select ultimo=max(ordem) from grades_plano_biblio where id_plano=" & id_plano & " and complementar=" & comple
rs.Open sqlo, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then ultimaordem=rs("ultimo") else ultimaordem=0
if isnull(ultimaordem) then ultimaordem=0
rs.close
%>
	<td class=titulo nowrap> Bibliografia Complementar <input type="checkbox" name="complementar" value="ON" <%=complementar%> onclick="javascript:submit();"></td>
	<td class=titulo><input type="text" name="ordem" size="3" value="<%=ultimaordem+1%>" ></td>
	<td class=titulo></td>
</tr>
</table>

<%
teste=0
'************
'if request.form("cod_acervo")="" or teste=1 then
	if request.form("referencia")<>"" then ref_digitada=1
	if clique_compra<>"" then referencia=clique_compra else referencia=request.form("referencia")

%>
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr>
	<td class=titulo>Referência <font size=1 color=blue><u>(Digite a referência caso não localize no acervo)</td>
</tr>
<tr>
	<td class=fundo>
<%
len2=int(len(request.form("referencia"))/55)+2
%>
	<textarea rows="<%=len2%>" name="referencia" cols="55" style="background-color: #FFFFCC"><%=clique_compra%></textarea>
<%
if session("perlet_atual_plano")>="2010" then
end if 
%>
	</td>
</tr>
</table>
<%
'else
	if request.form("cod_acervo")<>"" then ref_digitada=0
	if clique_acervo<>"" then cod_acervo=clique_acervo else cod_acervo=request.form("cod_acervo")
	sqlp="select referencia from pe_biblio where cod_acervo=" & cod_acervo
	if cod_acervo<>"" then 
		rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
		pesquisada=rs3("referencia")
		rs3.close
	end if
%>
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr>
	<td class=titulo>Cod.Acervo</td>
	<td class=titulo>Referência</td>
</tr>
<tr>
	<td class=fundo>
		<input type="text" name="cod_acervo" size="7" value="<%=clique_acervo%>" >
	</td>
	<td class=fundo>
		<%=pesquisada%>
	</td>
</tr>
</table>
<%
'end if 'request.form ("cod_acervo")
%>
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Registro    " class="button" name="Bt_salvar"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="button" value="Fechar             " class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>

<hr>
<!-- pesquisa -->
<%
if ref_digitada=0 then
%>
<br>
<table border=0 cellpadding=3 cellspacing=0 style="border-collapse: collapse" width=<%=largura%>>
<tr><td class=grupo colspan=2>Para pesquisar no acervo</td></tr>
<tr>
	<td class=titulor nowrap>Tipo</td>
	<td class=titulor>Conteúdo</td>
	</tr>
<tr>
	<td class="campot"r nowrap><select size="1" name="selecao" <!--onChange="javascript:submit()"--> >
		<option value="LIVRE" <%if request.form("selecao")="LIVRE" then response.write "selected"%> >Livre</option>
		<option value="TITULO" <%if request.form("selecao")="TITULO" then response.write "selected"%> >Titulo</option>
		<option value="AUTOR" <%if request.form("selecao")="AUTOR" then response.write "selected"%> >Autor</option>
		</select>
	</td>
	<td class="campot"r>
		<input type="text" name="conteudo" size="30" value="<%=request.form("conteudo")%>">
		<input type="submit" name="B1" value="Pesquisar no acervo">
	</td>
</tr>
<tr>
	<td class=titulor></td>
	<td class=titulor>Referência</td>
</tr>
<%
conteudo=request.form("conteudo")
if request.form("B1")<>"" then
	inicio=now()
	'procura por livros digitados (solicitados)
	sql0="select distinct cod_acervo=0, referencia=convert(nvarchar(255),referencia), status from grades_plano_biblio where (status is not null and status<>'N') and referencia like '%" & conteudo & "%'"
	rs3.Open sql0, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then
		procura1=rs3.recordcount
		do while not rs3.eof
		select case rs3("status")
			case "P"
				status="Aquisição solicitada"
			case "A"
				status="Aquisição autorizada"
			case "C"
				status="Adquirido para acervo"
		end select
%>
<tr>
	<td class="campor" style="border-bottom:2px solid #000000"><%=status%></a></td>
	<td class="campor" style="border-bottom:2px solid #000000" onclick="javascript:submit();"><a class=r href="biblio_nova.asp?clique_compra=<%=rs3("referencia")%>"><%=rs3("referencia")%></a></td>
</tr>
<%
		rs3.movenext:	loop
	end if 'rs3.recordcount
	rs3.close

	'procura por livros no acervo
	sql1="select cod_acervo, referencia, classificacao, obra, ano_publicacao, desc_tipo_obra from pe_biblio where "
	select case request.form("selecao")
		case "LIVRE"
			sql2=" livre like '%" & conteudo & "%' or assunto like '%" & conteudo & "%' "
		case "TITULO"
			sql2=" titulo like '%" & conteudo & "%' "
		case "AUTOR"
			sql2=" autor like '%" & conteudo & "%' or autor_principal like '%" & conteudo & "'"
	end select
	sql=sql1 & sql2 & " order by obra, ano_publicacao "
	inicio=now()
	rs3.Open sql, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then
		do while not rs3.eof
%>
<tr>
	<td class="campor" style="border-bottom:2px solid #000000"><%=rs3("cod_acervo")%></a></td>
	<td class="campor" style="border-bottom:2px solid #000000" onclick="javascript:submit();"><a class=r href="biblio_nova.asp?clique_acervo=<%=rs3("cod_acervo")%>"><%=rs3("referencia")%> (<font color=red><%=rs3("desc_tipo_obra")%></font>)</a></td>
</tr>
<%
		rs3.movenext:	loop
		termino=now():duracao=termino-inicio
		response.write "<tr><td class=grupo colspan=2>Pesquisou " & rs3.recordcount & " livros em " & cdbl(int(duracao*86400*100)/100) & " seg.</td></tr>"
	else
		response.write "<tr><td class=grupo colspan=2>Nenhum registro encontrado</td></tr>"
	end if 'rs3.recordcount
	rs3.close
end if 'request.form
%>
</table>
<%
end if 'ref_digitada=1
%>

</form>
<%
set rs=nothing
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
<%
end if
%>
</body>
</html>