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
<title>Termo de Responsabilidade</title>
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
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
--></script>

</head>
<body>
<%
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("b1")="" then
%>
	<p class=titulo>Emissão de Termo de Responsabilidade
	<form method="POST" action="termoresp.asp" name="form">
	<table border="0" width="250" cellspacing="0"cellpadding="3">
	<tr>
		<td class=titulo>Data de Entrega</td>
	</tr>
	<tr>
		<td class=fundo><select size="1" name="dt_movimento" onChange="javascript:submit()">
		<option>Selecione uma data</option>
<%
		if isdate(request.form("dt_movimento"))=true then dt_movimento=cdate(request.form("dt_movimento"))
		sql2="select dt_movimento, count(chapa) as recibos from uniforme_estoque where id_mov in (1,3) group by dt_movimento order by dt_movimento;"
		rs.Open sql2, ,adOpenStatic, adLockReadOnly
		if rs.recordcount>0 then
		rs.movefirst:do while not rs.eof
		if dt_movimento=rs("dt_movimento") then temp1="selected" else temp1=""
%>
		<option value="<%=rs("dt_movimento")%>" <%=temp1%>><%=rs("dt_movimento")%>&nbsp;&nbsp;&nbsp; (<%=rs("recibos")%> itens)</option>
<%
		rs.movenext:loop
		else
		'response.write "<option value='0'>Sem lançamentos...</option>"
		end if
		rs.close
%>
		</select>
		</td>
	</tr>
	</table>

	<table border="0" width="250" cellspacing="0" cellpadding="3">
	<tr>
		<td class=titulo><input type="submit" value="Visualizar" class="button" name="B1"></td>
	</tr>
	</table>
<%
vezes=0
if request.form<>"" then
sql1="select e.chapa, f.nome, dt_movimento " & _
"from uniforme_estoque e, corporerm.dbo.pfunc f " & _
"where e.chapa=f.chapa collate database_default and e.dt_movimento='" & dtaccess(request.form("dt_movimento")) & "' " & _
"group by e.chapa, f.nome, dt_movimento "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<br>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo></td>
</tr>
<%
rs.movefirst:do while not rs.eof
classe="campo"
%>
<tr>
	<td class=<%=classe%>><%=rs("chapa")%></td>
	<td class=<%=classe%>><%=rs("nome")%></td>
	<td class=<%=classe%>>
		<input type="checkbox" name="emitir<%=vezes%>" value="ON" <%=""%> >
		<input type="hidden" name="id<%=vezes%>" value="<%=rs("chapa")%>">
		<input type="hidden" name="dt<%=vezes%>" value="<%=rs("dt_movimento")%>">
	</td>
</tr>
<%
vezes=vezes+1
rs.movenext:loop
session("unif_imp")=vezes-1
end if 'rs.recordcount
rs.close
%>
</table>
<%
end if 'request.form<>""
%>
</form>
<%
else ' request.form("b1")
	vez=session("unif_imp")
	dt_recibo=request.form("dt_movimento")
	sql="delete from uniforme_recibo where sessao='" & session.sessionid & "' "
	conexao.execute sql
	for a=0 to vez
		id=request.form("id" & a)
		dtmov=request.form("dt" & a)
		emitir=request.form("emitir" & a)
		'response.write id & " " & tabela & " " & emitir & "<br>"
		if emitir="ON" then
			sql="INSERT INTO uniforme_recibo ( sessao, data, chapa ) SELECT '" & session.sessionid & "', '" & dtaccess(dtmov) & "', '" & id & "'"
			conexao.execute sql
		end if
	next

sql1="select f.chapa, f.nome, f.codsecao, s.descricao as setor, p.sexo " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.psecao s, corporerm.dbo.ppessoa p " & _
"where f.codpessoa=p.codigo and f.codsecao=s.codigo and f.chapa collate database_default in (select chapa from uniforme_recibo where sessao='" & session.sessionid & "') order by f.nome "
rs.Open sql1, ,adOpenStatic, adLockReadOnly

if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 
if rs("sexo")="M" then s1="o" else s1="a"
if rs("sexo")="M" then s2="" else s2="a"
if rs("sexo")="M" then s3="o" else s3=""

%>
<!-- <div align="right"> -->
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650" height=990>
<tr>
	<td class="campop" align="left" valign=top height=62><img src="../images/logo_centro_universitario_unifieo_big.gif" width="200" border="0"></td>
</tr>
<tr>
	<td class="campop" align="center" valign=top height=50><b><u>TERMO DE RESPONSABILIDADE</u></b><br><%=rs("setor")%></td>
</tr>
<tr>
	<td class="campop" valign=top style="text-align:justify" height=100>Eu, <b><%=rs("nome")%></b>, 
	declaro ter recebido da FIEO-Fundação Instituto de Ensino para Osasco, os itens de uniforme abaixo relacionados
	para uso exclusivo nas instalações da empresa e durante o horário de trabalho.
	<br>Declaro, ainda, estar ciente de que:
	<br>a) me responsabilizo pelo ressarcimento em caso de extravio antes do período regular de substituição do uniforme;
	<br>b) devo devolver os itens de uniforme substituídos;
	<br>c) em caso de desligamento da empresa devo devolver os itens de uniformes em meu poder no prazo máximo de 24 horas sob pena de ter os seus valores descontados;
	<br>d) o valor de reembolso no caso de dano, extravio ou não devolução é de R$ 10,00 por peça de uniforme.
	</td>
</tr>

<tr>
	<td class="campop" valign=top align="center">
<!-- quadro dos uniformes entregue -->
	<table border="1" bordercolor="#000000" cellpadding="4" cellspacing="0" style="border-collapse: collapse" width=600>
	<tr><td class=titulop colspan=3>Discriminação dos Uniformes RECEBIDOS:</td></td>
	<tr>
		<td class=titulop width=400>Peça</td>
		<td class=titulop width=100>Tamanho</td>
		<td class=titulop width=100>Quantidade</td>
	</tr>
<%
sql2="select e.chapa, e.dt_movimento, e.id_item, e.id_mov, e.qt_novo, e.qt_usado, i.descricao, i.tamanho " & _
"from uniforme_estoque e, uniforme_item i, uniforme_tpmov t " & _
"where e.id_item=i.id_item and e.id_mov=t.id_mov " & _
"and t.tipo=-1 and e.chapa='" & rs("chapa") & "' and e.dt_movimento='" & dtaccess(request.form("dt_movimento")) & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof 
%>
	<tr>
		<td class="campop" align="left"><%=rs2("descricao")%></td>
		<td class="campop" align="center"><%=rs2("tamanho")%></td>
		<td class="campop" align="center"><%=rs2("qt_novo")+rs2("qt_usado")%></td>
	</tr>	
<%
rs2.movenext
loop
else
	for a=1 to 4
	response.write "<tr><td class=""campop"" align=""left"">&nbsp;</td>"
	response.write "<td class=""campop"" align=""left"">&nbsp;</td>"
	response.write "<td class=""campop"" align=""left"">&nbsp;</td></tr>"
	next 
end if 'rs2.recordcount>0
if rs2.recordcount>0 and rs2.recordcount<4 then
	for a=4 to (rs2.recordcount+1) step -1
	response.write "<tr><td class=""campop"" align=""left"">&nbsp;</td>"
	response.write "<td class=""campop"" align=""left"">&nbsp;</td>"
	response.write "<td class=""campop"" align=""left"">&nbsp;</td></tr>"
	next
end if
rs2.close
%>
	</table>
<!-- quadro dos uniformes entregue -->
	</td>
</tr>

<tr>
	<td class="campop" align="left">Osasco,&nbsp;<%=day(dt_recibo) & " de " & monthname(month(dt_recibo)) & " de " & year(dt_recibo) %></td>
</tr>
<tr>
	<td class="campop" align="left" height=50>Ass.:________________________________________<br>&nbsp;&nbsp;&nbsp;<%=rs("nome")%></td>
</tr>
<tr><td class="campop">&nbsp;</td></tr>
<tr><td class="campop" style="border-bottom:1 dotted #000000">&nbsp;</td></tr>
<tr><td class="campop">&nbsp;</td></tr>
<tr>
	<td class="campop" valign=top align="center">
<!-- quadro dos uniformes devolvidos -->
	<table border="1" bordercolor="#000000" cellpadding="4" cellspacing="0" style="border-collapse: collapse" width=600>
	<tr><td class=titulop colspan=3>Discriminação dos Uniformes DEVOLVIDOS:</td></td>
	<tr>
		<td class=titulop width=400>Peça</td>
		<td class=titulop width=100>Tamanho</td>
		<td class=titulop width=100>Quantidade</td>
	</tr>
<%
sql2="select e.chapa, e.dt_movimento, e.id_item, e.id_mov, e.qt_novo, e.qt_usado, i.descricao, i.tamanho " & _
"from uniforme_estoque e, uniforme_item i, uniforme_tpmov t " & _
"where e.id_item=i.id_item and e.id_mov=t.id_mov " & _
"and t.tipo=1 and e.chapa='" & rs("chapa") & "' and e.dt_movimento='" & dtaccess(request.form("dt_movimento")) & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof 
%>
	<tr>
		<td class="campop" align="left"><%=rs2("descricao")%></td>
		<td class="campop" align="center"><%=rs2("tamanho")%></td>
		<td class="campop" align="center"><%=rs2("qt_novo")+rs2("qt_usado")%></td>
	</tr>	
<%
rs2.movenext
loop
else
	for a=1 to 4
	response.write "<tr><td class=""campop"" align=""left"">&nbsp;</td>"
	response.write "<td class=""campop"" align=""left"">&nbsp;</td>"
	response.write "<td class=""campop"" align=""left"">&nbsp;</td></tr>"
	next 
end if 'rs2.recordcount>0
if rs2.recordcount>0 and rs2.recordcount<4 then
	for a=4 to (rs2.recordcount+1) step -1
	response.write "<tr><td class=""campop"" align=""left"">&nbsp;</td>"
	response.write "<td class=""campop"" align=""left"">&nbsp;</td>"
	response.write "<td class=""campop"" align=""left"">&nbsp;</td></tr>"
	next
end if
rs2.close
%>
	</table>
<!-- quadro dos uniformes devolvidos -->

	</td>
</tr>
<tr>
	<td class="campop" align="left">Osasco,&nbsp;<%=day(dt_recibo) & " de " & monthname(month(dt_recibo)) & " de " & year(dt_recibo) %></td>
</tr>
<tr>
	<td class="campop" align="left" height=50>Ass.:________________________________________<br>&nbsp;&nbsp;&nbsp;<%=rs("nome")%></td>
</tr>
<tr>
	<td class="campop" valign="center" height=50 style="border:2 solid #000000">
	Conferido:____________________________________  Ass:____________________________________
	</td>
</tr>
<tr><td class="campop">&nbsp;</td></tr>

</table>

<!-- </div> -->
<%
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>"
rs.movenext
loop
end if 'recordcount >formulario
rs.close
%>

<%
end if 
%>
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>