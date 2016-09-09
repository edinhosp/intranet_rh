<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a94")="N" or session("a94")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Uniforme</title>
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
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

sqlb="AND f.CHAPA='" & request("codigo") & "' "
sqlc="ORDER BY f.CHAPA"
sql1=sqla & sqlb & sqlc
sql1="select f.nome, f.codsituacao, f.chapa, f.dataadmissao, c.nome as funcao, " & _
"f.codsecao, s.descricao as secao, f.datademissao " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.pfuncao c, corporerm.dbo.psecao s " & _
"where f.codfuncao=c.codigo and f.codsecao=s.codigo " & _
"and f.chapa='" & request("codigo") & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>
<% if session("a94")="T" then %>
<a href="funcionario_ver.asp?codigo=<%=rs("chapa")%>" onMouseOver="window.status='Clique aqui para atualizar após as alterações'; return true" onMouseOut="window.status=''; return true" onmouseover >
<img border="0" src="../images/write.gif" alt="Clique para atualizar">
<font size="1">!</font>
</a>
<% end if %>
CONTROLE DE UNIFORME</p>
<%
sql2="select descricao from corporerm.dbo.pcodsituacao where codcliente='" & rs("codsituacao") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
sit=rs2("descricao")
rs2.close
%>
<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="650">
<tr><td class=grupo>Dados Pessoais</td></tr>
</table>

<%q1=320:q2=100:q3=225%>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td width="<%=q1%>" valign="top">
<!-- quadro -->
<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="<%=q1%>">
<tr>
	<td class=titulor>&nbsp;Chapa:</td>
	<td class=titulor>&nbsp;Nome:</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rs("chapa")%>&nbsp;</td>
	<td class="campor"><b>&nbsp;<%=rs("nome")%>&nbsp;</b></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="<%=q1%>">
<tr>
	<td class=titulor>&nbsp;Situação:</td>
	<td class=titulor>&nbsp;Admissão:</td>
	<td class=titulor>&nbsp;Função:</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=sit%>&nbsp;<%if rs("codsituacao")="D" then response.write rs("datademissao")%></td>
	<td class="campor">&nbsp;<%=rs("dataadmissao")%>&nbsp;</td>
	<td class="campor">&nbsp;<%=rs("funcao")%>&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="<%=q1%>">
<tr>
	<td class=titulor>&nbsp;Seção:</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rs("codsecao")%>&nbsp;<%=rs("secao")%></td>
</tr>
</table>

<!-- quadro inicio mudanca-->
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="<%=q1%>">
<tr><th class=titulo colspan=4> Histórico das Categorias</th></tr>
<tr>
	<td class=titulor align="center">Categoria</td>
	<td class=titulor align="center">Inicio</td>
	<td class=titulor align="center">&nbsp;</td>
	<td class=titulor align="center">&nbsp;</td>
</tr>
<%
sql2="SELECT id_fcat, fc.chapa, fc.id_cat, fc.inicio, c.descricao " & _
"FROM uniforme_func_cat fc, uniforme_categoria c " & _
"where fc.chapa='" & rs("chapa") & "' and fc.id_cat=c.id_cat order by fc.inicio, c.descricao "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
linhas=rs2.recordcount
rs2.movefirst
inicio=1
do while not rs2.eof
if cdate(rs2("inicio"))<int(now) then fundo="style='background-color:#ffcccc'" else fundo="style='background-color:#ccffcc'"
if rs2.absoluteposition<>rs2.recordcount then fundo="style='background-color:#ffcccc'" else fundo="style='background-color:#ccffcc'"
%>
<tr>
	<td <%=fundo%> class="campor"><%=rs2("descricao") %></td>
	<td <%=fundo%> class="campor"><%=rs2("inicio") %>    </td>
	<td <%=fundo%> class="campor">&nbsp; 
	<% if session("a94")="T" or session("a94")="C" then %>
		<a href="func_cat_alteracao.asp?codigo=<%=rs2("id_fcat")%>" onclick="NewWindow(this.href,'AlteracaoFuncCat','440','170','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/folder95o.gif" alt="alterar esta categoria" width=13></a>
	<% end if %>
	</td>
	<%if inicio=1 then %>
	<td class="campor" rowspan=<%=linhas%> valign="center" align="center">
	<% if session("a94")="T" or session("a94")="C" then %>
		<a href="func_cat_nova.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'InclusaoFuncCat','440','170','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/Appointment.gif" alt="inserir nova categoria"></a>
	<% end if %>
	</td>
	<% end if 'inicio=1%>
</tr>
<%
id_fcat=rs2("id_fcat"):id_cat=rs2("id_cat")
rs2.movenext
inicio=0
loop
else ' sem registros/planos
%>
<tr>
	<td class="campor" colspan=3>&nbsp;</td>
	<td class="campor" rowspan=1 valign="center" align="center">
	<% if session("a94")="T" or session("a94")="C" then %>
		<a href="func_cat_nova.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'InclusaoFuncCat','440','170','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/Appointment.gif" alt="inserir nova categoria"></a>
	<% end if %>
	</td>
</tr>
<%
end if
rs2.close
%>
</table>
<!-- fim quadro mudanca -->
</td>

<td width="<%=q2%>" valign="top" style="border-right:1 dotted #000000;border-left:1 dotted #000000">
	<img border="0" src="../func_foto.asp?chapa=<%=rs("chapa")%>"  width="<%=q2%>">
</td>

<td width="<%=q3%>" valign="top">
<!-- quadro uniformes -->
<%
'response.write id_fcat & "-" & id_cat

sqlc="select count(chapa) as total from uniforme_func_item where chapa='" & rs("chapa") & "' and id_fcat=" & id_fcat & " "
rs3.Open sqlc, ,adOpenStatic, adLockReadOnly
if rs3("total")<1 then
%>
<p style="margin-top:0;margin-bottom:0;text-align:center">
<a href="func_item_primeiro.asp?chapa=<%=rs("chapa")%>&id_cat=<%=id_cat%>&id_fcat=<%=id_fcat%>" onclick="NewWindow(this.href,'InclusaoFuncItem','420','350','yes','center');return false" onfocus="this.blur()">
Inserir uniformes para a categoria</a>
<%
else
%>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="<%=q3%>">
<tr><th class=titulo colspan=3>Uniformes cadastrados</th></tr>
<tr>
	<td class=titulor align="center">Uniforme</td>
	<td class=titulor align="center">Tamanho</td>
	<td class=titulor align="center">&nbsp;</td>
</tr>
<%
sql2="SELECT id_fitem, fi.chapa, fi.id_item, i.descricao, i.tamanho " & _
"FROM uniforme_func_item fi, uniforme_item i " & _
"where fi.chapa='" & rs("chapa") & "' and fi.id_item=i.id_item and id_fcat=" & id_fcat & " order by i.descricao "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
linhas=rs2.recordcount
rs2.movefirst
do while not rs2.eof
%>
<tr>
	<td <%=fundo%> class="campor"><%=rs2("descricao") %></td>
	<td <%=fundo%> class="campor" align="center"><%=rs2("tamanho") %>    </td>
	<td <%=fundo%> class="campor">&nbsp; 
	<% if session("a94")="T" or session("a94")="C" then %>
		<a href="func_item_alteracao.asp?codigo=<%=rs2("id_fitem")%>" onclick="NewWindow(this.href,'AlteracaoFuncItem','440','110','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/folder95o.gif" alt="alterar este uniforme" width=13></a>
	<% end if %>
	</td>
</tr>
<%
rs2.movenext
loop
end if 'rs2.recordcount
rs2.close
%>
</table>
	<% if session("a94")="T" or session("a94")="C" then %>
	<a href="func_item_nova.asp?chapa=<%=rs("chapa")%>&id_cat=<%=id_cat%>&id_fcat=<%=id_fcat%>" onclick="NewWindow(this.href,'InclusaoFuncItem','440','110','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/Appointment.gif" alt="inserir novo uniforme" width=16></a>
	<% end if %>

<%
end if
%>

<!-- fim quadro uniformes -->
</td>
</tr>
</table>
<%
'rs.movenext
'loop
%>
<hr>
<!-- movimentação -->
<table border="1" bordercolor="#CCCCCC" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><th class="titulo" colspan="12">Lançamentos de Movimentação</th></tr>
<tr>
	<td class="titulo" align="center">Data</td>
	<td class="titulo" align="center" Colspan=2>Tipo</td>
	<td class="titulo" align="center">Uniforme</td>
	<td class="titulo" align="center">Tam.</td>
	<td class="titulo" align="center">Novo</td>
	<td class="titulo" align="center">Usado</td>
	<td class="titulo" align="center">&nbsp;</td>
</tr>
<%
'a amarelo l verde t azul
sql2="select e.chapa, e.id_est, e.id_item, e.dt_movimento, e.id_mov, e.qt_novo, e.qt_usado, i.descricao as uniforme, " & _
"i.tamanho, i.codigorm, t.descricao as movimento, t.tipo " & _
"from uniforme_estoque e, uniforme_item i, uniforme_tpmov t " & _
"where e.id_item=i.id_item and e.id_mov=t.id_mov and e.chapa='" & rs("chapa") & "' " & _
"order by e.dt_movimento, i.descricao "

rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
linhas=rs2.recordcount
rs2.movefirst
inicio=1
do while not rs2.eof
tipo=""
if rs2("tipo")="1" then tipo="E" 
if rs2("tipo")="-1" then tipo="S" 
%>
<tr>
	<td class="campo" align="center"><%=rs2("dt_movimento")%></td>
	<td class="campo" align="left"><%=rs2("movimento")%></td>
	<td class="campo" align="center"><%=tipo%></td>
	<td class="campo" align="left"><%=rs2("uniforme")%></td>
	<td class="campo" align="center"><%=rs2("tamanho")%></td>
	<td class="campo" align="center"><%=rs2("qt_novo")*rs2("tipo")%></td>
	<td class="campo" align="center"><%=rs2("qt_usado")*rs2("tipo")%></td>
	<td class="campor">
	<% if session("a94")="T" then %>
	<a href="estoque_alteracao.asp?codigo=<%=rs2("id_est")%>" onclick="NewWindow(this.href,'AlteracaoEstoque','420','200','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/folder95o.gif" width="14" alt="Alterar este lançamento"></a>
	<% end if %>
	</td>
</tr>
<%
rs2.movenext
inicio=0
loop
else ' sem registros/planos
%>
  <tr><td class="campor" colspan="7">&nbsp;</td></tr>
<%
end if
%>
</table>
<!-- quadro fim mudanca -->

<table><tr>
<td valign="top">
<% if session("a94")="T" then %>
<a href="estoque_nova.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'InclusaoEstoque','420','200','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo lançamento" WIDTH="16" HEIGHT="16">
<font size="1">inserir novo lançamento</font></a>
<% end if %>
</td>
</tr></table>

</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
set rs3=nothing
conexao.close
set conexao=nothing
%>