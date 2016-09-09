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
<title>Alteração de Bibliografia</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1()	{	form.chapa.value=form.nome.value;	}
function chapa1()	{	form.nome.value=form.chapa.value;	}
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
set rs3.ActiveConnection = conexao

	if request.form("bt_salvar")<>"" then
		tudook=1
		sql="UPDATE grades_plano_biblio SET id_plano=id_plano "
		if request.form("cod_acervo")<>"" then sql=sql & ", cod_acervo=" & request.form("cod_acervo") else sql=sql & ", cod_acervo=null"
		if request.form("complementar")="ON" then complementar=1 else complementar=0
		if request.form("cod_acervo")="" and request.form("referencia")="" then tudook=0
		sql=sql & ", complementar=" & complementar 
		if request.form("ordem")<>"" then 
			sql=sql & ", ordem=" & request.form("ordem")
		else
			sqlo="select ultimo=max(ordem) from grades_plano_biblio where id_biblio=" & request.form("id_biblio") & " and complementar=" & complementar
			rs.Open sqlo, ,adOpenStatic, adLockReadOnly
			ultimaordem=rs("ultimo")
			rs.close
			sql=sql & ", ordem=" & ultimaordem+1
		end if
		if request.form("referencia")<>"" then sql=sql & ", referencia='" & request.form("referencia") & "'"
		sql=sql & ", usuarioa='" & session("usuariomaster") & "'"
		sql=sql & ", dataa=getdate() "
	if session("perlet_atual_plano")>="2010" then
		status="status=null "
		if request.form("cod_acervo")="" and request.form("referencia")<>"" then status="status='P' "
		sql=sql & ", " & status
	end if 
		sql=sql & "WHERE id_biblio=" & request.form("id_biblio")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM grades_plano_biblio WHERE id_biblio=" & request.form("id_biblio")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if isnull(request("codigo")) or request("codigo")="" then
		id_biblio=session("id_alt_biblio")
	else
		id_biblio=request("codigo")
	end if
	sql1="select b.* from grades_plano_biblio p where id_biblio=" & id_biblio
	sql1="select p.id_biblio, p.id_plano, p.complementar, p.cod_acervo, p.ordem, p.referencia digitada, p.status, b.referencia pesquisada " & _
	"from grades_plano_biblio p left join pe_biblio b on b.cod_acervo=p.cod_acervo " & _
	"where id_biblio=" & id_biblio & " "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if request("clique_acervo")<>"" then
	clique_acervo=request.form("clique_acervo")
end if

if request("clique_compra")<>"" then
	clique_compra=request.form("referencia")
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="" ) or request.form("b1")<>"" or (request.form("bt_salvar")<>"" and tudook=0) then

session("id_alt_biblio")=rs("id_biblio")
largura=430
%>

<form method="POST" action="biblio_alteracao.asp" name="form">
<input type="hidden" name="id_biblio" size="4" value="<%=rs("id_biblio")%>" style="font-size: 8 pt" >  
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr><td class=grupo>Alteração de Bibliografia</td></tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo></td>
	<td class=titulo>Ordem</td>
	<td class=titulo width=100%></td>
</tr>
<tr>
	<td class=titulo><font color=gray><%=rs("id_biblio")%></td>
<%if rs("complementar")=true then complementar="checked" else complementar=""%>
	<td class=titulo nowrap> Bibliografia Complementar <input type="checkbox" name="complementar" value="ON" <%=complementar%> ></td>
	<td class=titulo><input type="text" name="ordem" size="3" value="<%=rs("ordem")%>" ></td>
	<td class=titulo></td>
</tr>
</table>
<%
teste=0
if rs("cod_acervo")="" or rs("cod_acervo")=0 or isnull(rs("cod_acervo")) or request("clique_compra")<>"" or teste=1 then
	ref_digitada=1
	if request("clique_compra")="" then
		digitada=rs("digitada")
	else
		digitada=request("clique_compra")
	end if
%>
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr>
	<td class=titulo>Referência</td>
</tr>
<tr>
	<td class=fundo>
<%
len2=int(len(rs("digitada"))/55)+2
%>
	<textarea rows="<%=len2%>" name="referencia" cols="55" style="background-color: #FFFFCC"><%=digitada%></textarea>
	</td>
</tr>
</table>
<%
else
	ref_digitada=0
	pesquisada=rs("pesquisada")
	if request("clique_acervo")="" then 
		cod_acervo=rs("cod_acervo") 
	else 
		cod_acervo=request("clique_acervo")
		sqlp="select referencia from pe_biblio where cod_acervo=" & cod_acervo
		rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
		pesquisada=rs3("referencia")
	end if
%>
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr>
	<td class=titulo>Cod.Acervo</td>
	<td class=titulo>Referência</td>
</tr>
<tr>
	<td class=fundo>
		<input type="text" name="cod_acervo" size="7" value="<%=cod_acervo%>" >
	</td>
	<td class=fundo>
		<%=pesquisada%>
	</td>
</tr>
</table>
<%
end if
%>
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Alterações  " class="button" name="Bt_salvar" onclick="submit();"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>
</tr>
</table>

<hr>
<!-- pesquisa -->
<%
if ref_digitada=0 or request("clique_compra")<>"" then
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
		<input type="submit" name="b1" value="Pesquisar no acervo">
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
	<td class="campor" style="border-bottom:2px solid #000000"><a class=r href="biblio_alteracao.asp?clique_compra=<%=rs3("referencia")%>"><%=rs3("referencia")%></a></td>
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
	rs3.Open sql, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then
		do while not rs3.eof
%>
<tr>
	<td class="campor" style="border-bottom:2px solid #000000"><%=rs3("cod_acervo")%></a></td>
	<td class="campor" style="border-bottom:2px solid #000000"><a class=r href="biblio_alteracao.asp?clique_acervo=<%=rs3("cod_acervo")%>"><%=rs3("referencia")%> (<font color=red><%=rs3("desc_tipo_obra")%></font>)</a></td>
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
end if
%>
</body>
</html>