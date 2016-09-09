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
<title>UNIFIEO - Plano de Ensino</title>
<script language="javascript" type="text/javascript">
<!--
/****************************************************
     Author: Eric King
     Url: http://redrival.com/eak/index.shtml
     This script is free to use as long as this info is left in
     Featured on Dynamic Drive script library (http://www.dynamicdrive.com)
****************************************************/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>

<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

	if request.form("bt_salvar")<>"" then
		tudook=1

if request.form("justificativa")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe a Justificativa!');</script>"
if request.form("ementa")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe a Ementa!');</script>"
if request.form("objetivos_gerais")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe os Objetivos Gerais!');</script>"

		sql="UPDATE grades_plano SET "
		sql=sql & "justificativa     ='" & request.form("justificativa")      & "', "
		sql=sql & "ementa            ='" & request.form("ementa")             & "', "
		sql=sql & "objetivos_gerais  ='" & request.form("objetivos_gerais")   & "', "
		sql=sql & "unidades_tematicas='" & request.form("unidades_tematicas") & "', "
		sql=sql & "metodologia       ='" & request.form("metodologia")        & "', "
		sql=sql & "avaliacao         ='" & request.form("avaliacao")          & "', "
		sql=sql & "bibliografia      ='" & request.form("bibliografia")       & "', "
		sql=sql & "bibliografiac     ='" & request.form("bibliografiac")       & "', "
		if session("acesso")=2 then
			sql=sql & "novo=0,coord=0, "
		end if
		if request.form("pa")="ON" then 
			sql=sql & "pa=1, ":pa=1
		else 
			sql=sql & "pa=0, ":pa=0
		end if
		if request.form("prof")="ON" then 
			sql=sql & "prof=1, ":prof=1
		else 
			sql=sql & "prof=0, ":prof=0
		end if
		if request.form("paa")<>pa then
			sql=sql & "usuariop='" & session("usuariomaster") & "', datap=getdate(), "
		end if
		sql=sql & "usuarioa          ='" & session("usuariomaster")           & "', "
		sql=sql & "dataa=getdate() "
		sql=sql & " WHERE id_plano=" & session("id_alt_plano") & " "
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM grades_plano WHERE codmat='" & session("id_alt_plano") & "' "
		sql="update grades_plano set justificativa=null,ementa=null,objetivos_gerais=null,unidades_tematicas=null,metodologia=null,avaliacao=null,bibliografia=null,bibliografiac=null,novo=1,pa=0 where id_plano=" & session("id_alt_plano")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if isnull(request("codigo")) or request("codigo")="" then
		if session("id_alt_plano")="" then id_plano=request.form("id_plano") else id_plano=session("id_alt_plano")
		'id_plano=session("id_alt_plano")
	else
		id_plano=request("codigo")
	end if
	sqla="select p.*, m.materia from grades_plano p inner join corporerm.dbo.umaterias m on m.codmat collate database_default=p.codmat "
	sqlb="where p.id_plano=" & id_plano & " "
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then

session("id_alt_plano")=rs("id_plano")
if request.form("justificativa")     <>"" then justificativa     =request.form("justificativa")     else justificativa     =rs("justificativa")
if request.form("ementa")            <>"" then ementa            =request.form("ementa")            else ementa            =rs("ementa")
if request.form("objetivos_gerais")  <>"" then objetivos_gerais  =request.form("objetivos_gerais")  else objetivos_gerais  =rs("objetivos_gerais")
if request.form("unidades_tematicas")<>"" then unidades_tematicas=request.form("unidade_tematicas") else unidades_tematicas=rs("unidades_tematicas")
if request.form("metodologia")       <>"" then metodologia       =request.form("metodologia")       else metodologia       =rs("metodologia")
if request.form("avaliacao")         <>"" then avaliacao         =request.form("avaliacao")         else avaliacao         =rs("avaliacao")
if request.form("bibliografia")      <>"" then bibliografia      =request.form("bibliografia")      else bibliografia      =rs("bibliografia")
if request.form("bibliografiac")     <>"" then bibliografiac     =request.form("bibliografiac")     else bibliografiac     =rs("bibliografiac")
materia=rs("materia")
largura=600
session("perlet_atual_plano")=rs("perlet")
%>
<form method="POST" action="plano_alteracao.asp" name="form">
<input type="hidden" name="id_plano" size="4" value="<%=rs("id_plano")%>" style="font-size: 8 pt" >  
<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=grupo>Alteração de Plano de Ensino</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr><td class=fundo>Disciplina</td>
	<td class=fundop><%=rs("codmat")%> - <b><%=rs("materia")%></b> da Grade <%=rs("grade")%> no curso <%=rs("coddoc")%>, Período Letivo <%=rs("perlet")%>.
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">JUSTIFICATIVA</font></b></td></tr>
<%
len1=int(len(justificativa)/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len1%>" name="justificativa" cols="80" style="background-color: #FFFFCC"><%=justificativa%></textarea>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">EMENTA</font></b></td></tr>
<%
len2=int(len(ementa)/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len2%>" name="ementa" cols="80" style="background-color: #FFFFCC"><%=ementa%></textarea>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">OBJETIVOS GERAIS</font></b></td></tr>
<%
len3=int(len(objetivos_gerais)/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len3%>" name="objetivos_gerais" cols="80" style="background-color: #FFFFCC"><%=objetivos_gerais%></textarea>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">UNIDADES TEMÁTICAS</font></b></td></tr>
<%
len4=int(len(unidades_tematicas)/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len4%>" name="unidades_tematicas" cols="80" style="background-color: #FFFFCC"><%=unidades_tematicas%></textarea>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">METODOLOGIA</font></b></td></tr>
<%
len5=int(len(metodologia)/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len5%>" name="metodologia" cols="80" style="background-color: #FFFFCC"><%=metodologia%></textarea>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr><td class=fundo><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">AVALIAÇÃO</font></b></td></tr>
<%
len6=int(len(avaliacao)/70)+2
%>
<tr>
	<td class=titulo>
	<textarea rows="<%=len6%>" name="avaliacao" cols="80" style="background-color: #FFFFCC"><%=avaliacao%></textarea>
	</td>
</tr>
</table>

<!-- bibliografia -->
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr><td class=fundo align="left" colspan=2><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">BIBLIOGRAFIA BÁSICA</font></b></td>
	<td class=fundo align="right">
	<a href="biblio_nova.asp?codigo=<%=rs("id_plano")%>&compl=0" onclick="NewWindow(this.href,'bibliografia_nova','470','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/LeafSearch.gif" width="16" height="16" border="0" alt="Pesquisar e inserir"></a>
	</td>
</tr>
<%
len7=int(len(bibliografia)/70)+2
%>
<%
sql="select p.id_biblio, p.id_plano, p.complementar, p.cod_acervo, p.ordem, p.referencia digitada, p.status, b.referencia pesquisada " & _
"from grades_plano_biblio p left join pe_biblio b on b.cod_acervo=p.cod_acervo " & _
"where id_plano=" & id_plano & " and complementar=0 " & _
"order by ordem"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
if rs2("pesquisada")<>"" then referencia=rs2("pesquisada") else referencia=rs2("digitada")
%>
<tr>
	<td class="campoa"r style="border-bottom:1px solid #000000"><%=rs2("ordem")%></td>
	<td class="campoa"r style="border-bottom:1px solid #000000"><%=referencia%></td>
	<td class=fundo>
	<a href="biblio_alteracao.asp?codigo=<%=rs2("id_biblio")%>" onclick="NewWindow(this.href,'bibliografia_alterar','470','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/novo.gif" width="17" height="17" border="0" alt="Alterar"></a>
	</td>
</tr>
<%
rs2.movenext
loop
rs2.close
%>
</table>

<!-- bibliografia complementar-->
<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr><td class=fundo align="left" colspan=2><p style="margin-top: 0; margin-bottom: 0"><b><font color="#0000FF">BIBLIOGRAFIA COMPLEMENTAR</font></b></td>
	<td class=fundo align="right">
	<a href="biblio_nova.asp?codigo=<%=rs("id_plano")%>&compl=1" onclick="NewWindow(this.href,'bibliografia_nova','470','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/LeafSearch.gif" width="16" height="16" border="0" alt="Pesquisar e inserir"></a>
	</td>
</tr>
<%
len8=int(len(bibliografiac)/70)+2
%>
<%
sql="select p.id_biblio, p.id_plano, p.complementar, p.cod_acervo, p.ordem, p.referencia digitada, p.status, b.referencia pesquisada " & _
"from grades_plano_biblio p left join pe_biblio b on b.cod_acervo=p.cod_acervo " & _
"where id_plano=" & id_plano & " and complementar=1 " & _
"order by ordem"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
if rs2("pesquisada")<>"" then referencia=rs2("pesquisada") else referencia=rs2("digitada")
%>
<tr>
	<td class="campoa"r style="border-bottom:1px solid #000000"><%=rs2("ordem")%></td>
	<td class="campoa"r style="border-bottom:1px solid #000000"><%=referencia%></td>
	<td class=fundo>
	<a href="biblio_alteracao.asp?codigo=<%=rs2("id_biblio")%>" onclick="NewWindow(this.href,'bibliografiac_alterar','470','400','yes','center');return false" onfocus="this.blur()">
	<img src="../images/novo.gif" width="17" height="17" border="0" alt="Alterar"></a>
	</td>
</tr>
<%
rs2.movenext
loop
rs2.close
%>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<%
if session("a93")="T" or session("usuariogrupo")="RH" or session("usuariogrupo")="PLANEJAMENTO" then
	if (rs("pa")=false or isnull(rs("pa"))) and request.form("pa")<>"ON" then obs1="" else obs1="checked"
%>
<input type="hidden" name="paa" value="<%=rs("pa")%>">
<tr>
	<td class=titulo>Validado pelo Planejamento Acadêmico<input type="checkbox" name="pa" value="ON" <%=obs1%> ></td>
</tr>
<%
end if

if session("usuariogrupo")="PROFESSOR" or session("usuariogrupo")="RH" then
	if (rs("prof")=false or isnull(rs("prof"))) and request.form("prof")<>"ON" then obs2="" else obs2="checked"
%>
<input type="hidden" name="profa" value="<%=rs("prof")%>">
<tr>
	<td class=titulo style="border:1px dotted #000000"><font color=green>Professor:</font><font color=brown> Marque aqui se você já completou a digitação para ser validado pelo coordenador<input type="checkbox" name="prof" value="ON" <%=obs2%> ></td>
</tr>
<%
end if
%>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="<%=largura%>">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if
set rs2=nothing
set rsnome=nothing
conexao.close
set conexao=nothing

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.document.form.submit();self.close();</script>"
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