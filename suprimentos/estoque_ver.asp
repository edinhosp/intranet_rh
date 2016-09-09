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
<title>Controle de Estoque Uniformes</title>
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
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
set rsq=server.createobject ("ADODB.Recordset")
Set rsq.ActiveConnection = conexao
item=request("item")
saldo=0:novo=0:usado=0

sql1="select i.id_item, i.descricao, i.codigorm, i.tamanho, i.sequencia, i.qt_novo, i.qt_usado, i.preco " & _
"from uniforme_item i where i.id_item=" & item
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>

<p style="margin-top: 0; margin-bottom: 0" class="titulo">
<% if session("a94")="T" then %>
<a href="estoque_ver.asp?item=<%=item%>" onMouseOver="window.status='Clique aqui para atualizar após as alterações'; return true" onMouseOut="window.status=''; return true" onmouseover>
<img border="0" src="../images/write.gif" alt="Clique para atualizar" width="16" height="16">
<font size="1">!</font>
</a>
<% end if %>
Controle de Estoque - Uniformes</p>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr><td class="grupo">Informação</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class="titulor"> Cod.</td>
	<td class="titulor"> Descrição</td>
	<td class="titulor"> Tamanho</td>
</tr>
<tr>
	<td class="campor"> <%=rs("id_item")%></td>
	<td class="campor"><b><%=rs("descricao")%></b></td>
	<td class="campor"> <%=rs("tamanho")%></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class="titulor"> Código RM</td>
	<td class="titulor"> Preço</td>
	<td class="titulor"> Est.Novo</td>
	<td class="titulor"> Est.Usado</td>
</tr>
<tr>
	<td class="campor"> <%=rs("codigorm")%></td>
	<td class="campor"> <%=formatnumber(rs("preco"),2)%></td>
	<td class="campor"> <%=rs("qt_novo")%></td>
	<td class="campor"> <%=rs("qt_usado")%></td>
</tr>
</table>

<!-- quadro inicio mudanca-->
<table border="1" bordercolor="#CCCCCC" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><th class="titulo" colspan="12">Lançamentos de Movimentação</th></tr>
<tr>
	<td class="titulo" align="center" rowspan=2>Data</td>
	<td class="titulo" align="center" rowspan=2 Colspan=2>Tipo</td>
	<td class="titulo" align="center" colspan=2>Estoque</td>
	<td class="titulo" align="center" rowspan=2>Funcionário</td>
	<td class="titulo" align="center" rowspan=2>&nbsp;</td>
</tr>
<tr>
	<td class="titulo" align="center">Novo</td>
	<td class="titulo" align="center">Usado</td>
</tr>
<tr> 
	<td class="campoa" colspan=3>Saldo Inicial</td>
	<td class="campoa" align="center"><%=rs("qt_novo")%></td>
	<td class="campoa" align="center"><%=rs("qt_usado")%></td>
	<td class="campoa"></td>	
	<td class="campoa"></td>	
</tr>
<%
'a amarelo l verde t azul
usado=usado+rs("qt_usado")
novo=novo+rs("qt_novo")
rs.close
saldo=0

sql2="select e.id_est, e.id_item, e.dt_movimento, e.id_mov, e.qt_novo, e.qt_usado, e.chapa, f.nome, t.descricao, t.tipo " & _
"from uniforme_estoque e, uniforme_tpmov t, pfunc f where t.id_mov=e.id_mov and e.chapa=f.chapa " & _
"and e.id_item=" & item & " order by e.dt_movimento "
sql2="SELECT e.id_est, e.id_item, e.dt_movimento, e.id_mov, e.qt_novo, e.qt_usado, e.chapa, f.NOME, t.descricao, t.tipo " & _
"FROM (uniforme_estoque AS e INNER JOIN uniforme_tpmov AS t ON e.id_mov=t.id_mov) LEFT JOIN corporerm.dbo.pfunc AS f ON e.chapa=f.CHAPA collate database_default " & _
"where e.id_item=" & item & " order by e.dt_movimento "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
linhas=rs2.recordcount
rs2.movefirst
inicio=1
do while not rs2.eof
if rs2("tipo")="1" then 
	tipo="E" 
elseif rs2("tipo")="-1" then 
	tipo="S" 
else tipo=""
end if
%>
<tr>
	<td class="campo" align="center"><%=rs2("dt_movimento")%></td>
	<td class="campo" align="left"><%=rs2("descricao")%></td>
	<td class="campo" align="left"><%=tipo%></td>
	<td class="campo" align="center"><%=rs2("qt_novo")*rs2("tipo")%></td>
	<td class="campo" align="center"><%=rs2("qt_usado")*rs2("tipo")%></td>
	<td class="campor" align="left">
    <% if session("a94")="T" or session("a94")="C" then %>
      <a class=r href="funcionario_ver.asp?codigo=<%=rs2("chapa")%>" onclick="NewWindow(this.href,'Alteracao','690','500','yes','center');return false" onfocus="this.blur()">
	<%=rs2("nome")%></a>
	<% end if %>
	</td>	
	<td class="campor">
	<% if session("a94")="T" then %>
	<a href="estoque_alteracao.asp?codigo=<%=rs2("id_est")%>" onclick="NewWindow(this.href,'AlteracaoEstoque','420','200','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/folder95o.gif" width="14" alt="Alterar este lançamento"></a>
	<% end if %>
	</td>
</tr>
<%
usado=usado+rs2("qt_usado")*rs2("tipo")
novo=novo+rs2("qt_novo")*rs2("tipo")
if rs2.absoluteposition=rs2.recordcount then
end if
rs2.movenext
inicio=0
loop
else ' sem registros/planos
%>
  <tr><td class="campor" colspan="7">&nbsp;</td></tr>
<%
end if
%>
<tr> 
	<td class="campot" colspan=3>Saldo Final</td>
	<td class="campot" align="center"><%=novo%></td>
	<td class="campot" align="center"><%=usado%></td>
	<td class="campot"></td>	
	<td class="campot"></td>	
</tr>
</table>
<!-- quadro fim mudanca -->
<table><tr>
<td valign="top">
<% if session("a94")="T" then %>
<a href="estoque_nova.asp?item=<%=item%>" onclick="NewWindow(this.href,'InclusaoEstoque','420','200','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo lançamento" WIDTH="16" HEIGHT="16">
<font size="1">inserir novo lançamento</font></a>
<% end if %>
</td>
</tr></table>

</body>
</html>
<%
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>