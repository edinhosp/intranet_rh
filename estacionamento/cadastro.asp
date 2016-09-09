<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a87")="N" or session("a87")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cadastramento de Veículos</title>
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
function nome1() { form.chapa.value=form.nome.value;form.submit();}
function chapa1() { form.nome.value=form.chapa.value;form.submit();}
--></script>
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rsq=server.createobject ("ADODB.Recordset")
Set rsq.ActiveConnection = conexao
set rsu=server.createobject ("ADODB.Recordset")
Set rsu.ActiveConnection = conexao
if request.form("chapa")<>"" then session("chapaveic")=request.form("chapa")
chapa=session("chapaveic")
if chapa="" then chapa="0"
sql1="SELECT dc_carga.CURSO FROM dc_carga GROUP BY dc_carga.CURSO;"
'rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<form method="POST" action="cadastro.asp" name="form">
<table border="0" cellpadding="3" cellspacing="0" width="420" style="border-collapse: collapse">
<tr><td class=titulo colspan=2><p style="margin-top:0;margin-bottom:0;color:Blue;font-size:10pt;text-align:left">
<b>Cadastramento de Veículos / Crachá de Estacionamento</font></p>
	</td></tr>
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=titulo><input type="text" value="<%=chapa%>" name="chapa" size="8" onchange="chapa1()" onfocus="javascript:window.status='Informe o chapa do funcionário'"></td>
	<td class=titulo>&nbsp;
	<select size="1" name="nome" onchange="nome1()" onfocus="javascript:window.status='Selecione o Nome do Funcionário'" >
<%
sql2="select t.chapa, t.nome from (select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' " & _
"union all " & _
"select chapa collate database_default, nome collate database_default from grades_novos) as t order by nome"
response.write sql2
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione Funcionário....</option>"
rs2.movefirst:do while not rs2.eof
if chapa=rs2("chapa") then tempc="selected" else tempc=""
%>
          <option value="<%=rs2("chapa")%>" <%=tempc%>><%=rs2("nome")%></option>
<%
rs2.movenext:loop
rs2.close
if request.form("cracha")="" then tc="H" else tc=request.form("cracha")
if tc="H" then w1=565:h1=350
if tc="V" then w1=330:h1=400
%>
	</select></td>
</tr>
<tr>
	<td class=titulo colspan=2>
	<input type="radio" name="cracha" value="V" onclick="javascript:form.submit();" <%if tc="V" then response.write "checked"%> > Vertical
	<input type="radio" name="cracha" value="H" onclick="javascript:form.submit();" <%if tc="H" then response.write "checked"%> > Horizontal
	</td>
</tr>
</table>
<!--
<input type="submit" value="Pesquisar" class="button" name="pesquisar" onfocus="javascript:window.status='Clique aqui para pesquisar'">
-->
</form>
<%
'*********** calculo *************
sql2="select * from grades_blocos where chapa1='" & chapa & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
if isnull(rs2("ns")) then ns=0 else ns=rs2("ns")
if isnull(rs2("co")) then co=0 else co=rs2("co")
if isnull(rs2("az")) then az=0 else az=rs2("az")
if isnull(rs2("am")) then am=0 else am=rs2("am")
if isnull(rs2("ve")) then ve=0 else ve=rs2("ve")
if isnull(rs2("li")) then li=0 else li=rs2("li")
if isnull(rs2("ma")) then ma=0 else ma=rs2("ma")
if isnull(rs2("br")) then br=0 else br=rs2("br")
if isnull(rs2("pr")) then pr=0 else pr=rs2("pr")
to1=co+az+am+ve+li+ma
to2=br+pr
if to2>to1 then classe2="campol" else classe2="campo"
if ns>0 then classe0="campoa" else classe0="campo"
if to1>to2 then classe1="campot" else classe1="campo"
%>
<table border="1" bordercolor="black" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="450">
<tr>
	<td class=titulo rowspan=2 align="center" width=50>Narciso</td>
	<td class=titulor colspan=6 align="center" width=300>Blocos Coral a Marron</td>
	<td class=titulor colspan=2 align="center" width=100>Blocos Branco e Prata</td>
</tr>
<tr>
	<td class=titulo align="center" width=50>CO</td>
	<td class=titulo align="center" width=50>AZ</td>
	<td class=titulo align="center" width=50>AM</td>
	<td class=titulo align="center" width=50>VE</td>
	<td class=titulo align="center" width=50>LI</td>
	<td class=titulo align="center" width=50>MA</td>
	<td class=titulo align="center" width=50>BR</td>
	<td class=titulo align="center" width=50>PR</td>
</tr>
<tr>
	<td class=campo align="center"><%=rs2("ns")%></td>
	<td class=campo align="center"><%=rs2("co")%></td>
	<td class=campo align="center"><%=rs2("az")%></td>
	<td class=campo align="center"><%=rs2("am")%></td>
	<td class=campo align="center"><%=rs2("ve")%></td>
	<td class=campo align="center"><%=rs2("li")%></td>
	<td class=campo align="center"><%=rs2("ma")%></td>
	<td class=campo align="center"><%=rs2("br")%></td>
	<td class=campo align="center"><%=rs2("pr")%></td>
</tr>
<tr>
	<td class=<%=classe0%> align="center"><%=ns%></td>
	<td class=<%=classe1%> colspan=6 align="center"><%=to1%></td>
	<td class=<%=classe2%> colspan=2 align="center"><%=to2%></td>
</tr>
</table>
<br>
<%
end if 'rs2.recordcount
rs2.close
%>

<!-- tabela -->
<table><tr><td valign=top>
<!-- tabela -->

<table border="1" bordercolor="Green" cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8>Cadastro de Veículos</td>
</tr>
<tr>
	<td class=titulor>Marca</td>
	<td class=titulor>Modelo</td>
	<td class=titulor>Ano</td>
	<td class=titulor>Cor</td>
	<td class=titulor>Placa</td>
	<td class=titulor>Cadastro</td>
	<td class=titulor>Cancelado</td>
	<td class=titulor>&nbsp;</td>
</tr>
<%
sql2="select diasuteismes, diasutrestantes, diasutproxmes, diasutmeioexp, diasutrestmeio, diasutproxmeio " & _
"from pfunc where chapa='" & chapa & "' "
sql2="SELECT Sum(NROVIAGENS) AS usa FROM corporerm.dbo.PFVALETR WHERE DTFIM>=getdate() AND CHAPA='" & chapa & "' "
rsq.Open sql2, ,adOpenStatic, adLockReadOnly
if rsq("usa")>0 then vtmes=rsq("usa") else vtmes=0
rsq.close
sql1="select id_veiculo, marca,modelo, ano, cor, placa, dtcadastro, dttermino " & _
"from veiculos where chapa='" & chapa & "' order by dtcadastro "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
if isnull(rs("dttermino")) then campo="campo" else campo="fundo"
vtmes=vtmes
%>
<tr>
	<td class=<%=campo%> ><%=rs("marca")%></td>
	<td class=<%=campo%> ><%=rs("modelo")%></td>
	<td class=<%=campo%> ><%=rs("ano")%></td>
	<td class=<%=campo%> ><%=rs("cor")%></td>
	<td class=<%=campo%>  nowrap><%=rs("placa")%></td>
	<td class=<%=campo%> ><%=rs("dtcadastro")%></td>
	<td class=<%=campo%> ><%=rs("dttermino")%></td>
	<td class=campo>
    <% if session("a87")<>"N" then %>
      <a href="cadastro_alteracao.asp?codigo=<%=rs("id_veiculo")%>" onclick="NewWindow(this.href,'AlteracaoVeiculo','520','240','no','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
	<% end if %>
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
%>
<tr><td class=campo colspan=13>
<% if session("a87")="T" then %>
<a class=r href="cadastro_nova.asp?chapa=<%=chapa%>" onclick="NewWindow(this.href,'InclusaoVeiculo','535','250','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
inserir novo veículo</a>
<% end if %>
</td>
</tr>
</table>

<!-- tabela -->
</td><td valign=top>
<!-- tabela -->
<table border="1" bordercolor="Green" cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=9>Controle de Estacionamento</td>
</tr>
<tr>
	<td class=titulor>Inicio</td>
	<td class=titulor>Termino</td>
	<td class=titulor>VY</td>
	<td class=titulor>BP</td>
	<td class=titulor>NS</td>
	<td class=titulor>JW</td>
	<td class=titulor>Cartão</td>
	<td class=titulor>Obs</td>
	<td class=titulor>&nbsp;</td>
</tr>
<%
sql2="select"
sql1="select id_est, vy, ns, bp, jw, inicio, termino, cartao, obs " & _
"from veiculos_a where chapa='" & chapa & "' order by termino, inicio "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
if rs("termino")>now then campo="campo" else campo="campot"
%>
<tr>
	<td class=<%=campo%> ><%=rs("inicio")%></td>
	<td class=<%=campo%> ><%=rs("termino")%></td>
	<td class=<%=campo%> valign="top" align="center">
	<%if vtmes=0 or vtmes="" or isnull(vtmes) then response.write "<a class=r href=""cracha.asp?chapa=" & chapa & "&c=vy&t=" & tc & """ onclick=""NewWindow(this.href,'ImpressaoCracha','" & w1 & "','" & h1 & "','yes','center');return false"" onfocus=""this.blur()"">" %>
	<%if rs("vy")=-1 then response.write "<img src='../images/truck.gif' width=13 border='0'>" %>
	<%if vtmes=0 or vtmes="" or isnull(vtmes) then response.write "</a>" %>
	</td>

	<td class=<%=campo%> valign="top" align="center">
	<%if vtmes=0 or vtmes="" or isnull(vtmes) then response.write "<a class=r href=""cracha.asp?chapa=" & chapa & "&c=bp&t=" & tc & """ onclick=""NewWindow(this.href,'ImpressaoCracha','" & w1 & "','" & h1 & "','yes','center');return false"" onfocus=""this.blur()"">" %>
	<%if rs("bp")=-1 then response.write "<img src='../images/truck.gif' width=13 border='0'>" %>
	<%if vtmes=0 or vtmes="" or isnull(vtmes) then response.write "</a>" %>
	</td>

	<td class=<%=campo%> valign="top" align="center">
	<%if vtmes=0 or vtmes="" or isnull(vtmes) then response.write "<a class=r href=""cracha.asp?chapa=" & chapa & "&c=ns&t=" & tc & """ onclick=""NewWindow(this.href,'ImpressaoCracha','" & w1 & "','" & h1 & "','yes','center');return false"" onfocus=""this.blur()"">" %>
	<%if rs("ns")=-1 then response.write "<img src='../images/truck.gif' width=13 border='0'>" %>
	<%if vtmes=0 or vtmes="" or isnull(vtmes) then response.write "</a>" %>
	</td>

	<td class=<%=campo%> valign="top" align="center">
	<%if vtmes=0 or vtmes="" or isnull(vtmes) then response.write "<a class=r href=""cracha.asp?chapa=" & chapa & "&c=jw&t=" & tc & """ onclick=""NewWindow(this.href,'ImpressaoCracha','" & w1 & "','" & h1 & "','yes','center');return false"" onfocus=""this.blur()"">" %>
	<%if rs("jw")=-1 then response.write "<img src='../images/truck.gif' width=13 border='0'>" %>
	<%if vtmes=0 or vtmes="" or isnull(vtmes) then response.write "</a>" %>
	</td>

	<td class=<%=campo%> ><%=rs("cartao")%></td>
	<td class=<%=campo%> ><%=rs("obs")%></td>
	<td class=campo>
    <% if session("a87")<>"N" and session("usuariomaster")="02379" then %>
	<a href="estac_alteracao.asp?codigo=<%=rs("id_est")%>" onclick="NewWindow(this.href,'AlteracaoEstacionamento','520','240','no','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
	<% end if %>
	</td>
</tr>
<%
pavy=cdbl(rs("vy"))
pans=cdbl(rs("ns"))
pabp=cdbl(rs("bp"))
pajw=cdbl(rs("jw"))
paid=rs("id_est")
acartao=rs("cartao")
fim=rs("termino")
rs.movenext:loop
end if
rs.close
%>
<tr><td class=campo colspan=13>
<% if session("a87")="T" then %>
<a class=r href="estac_nova.asp?chapa=<%=chapa%>&pavy=<%=pavy%>&pans=<%=pans%>&pabp=<%=pabp%> &pajw=<%=pajw%> &paid=<%=paid%>&acartao=<%=acartao%>&pafim=<%=fim%>" onclick="NewWindow(this.href,'InclusaoEstacionamento','535','250','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
inserir estacionamento</a>
<% end if %>
</td>
</tr>




</table>
<!-- tabela -->
</td></tr></table>
<!-- tabela -->


<br>

<p style="margin-top:0;margin-bottom:0;color:Blue;font-size:10pt;text-align:left">
<b>Uso de Vale-Transporte
<%
sql2="select diasuteismes, diasutrestantes, diasutproxmes, diasutmeioexp, diasutrestmeio, diasutproxmeio " & _
"from corporerm.dbo.pfunc where chapa='" & chapa & "' "
rsq.Open sql2, ,adOpenStatic, adLockReadOnly
if rsq.recordcount>0 then
vtmes=rsq("diasutproxmes")
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=420>
  <tr>
  	<td class=fundo>&nbsp;</td>
  	<td class=fundo>&nbsp;Mês Atual</td>
	<td class=fundo>&nbsp;Após Reajuste</td>
	<td class=fundo>&nbsp;Próximo Mês</td>
  </tr>
  <tr>
	<td class=fundo>&nbsp;Expediente Integral</td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rsq("diasuteismes")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rsq("diasutrestantes")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rsq("diasutproxmes")%>" onfocus="this.blur()"></td>
  </tr>
  <tr>
	<td class=fundo>&nbsp;Meio Expediente</td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rsq("diasutmeioexp")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rsq("diasutrestmeio")%>" onfocus="this.blur()"></td>
	<td class=fundo>&nbsp;<input type="text" size=9 value="<%=rsq("diasutproxmeio")%>" onfocus="this.blur()"></td>
  </tr>
</table>
<%
end if 'rs.recordcount
rsq.close
sqla="SELECT f.*, l.codtarifa FROM corporerm.dbo.PFVALETR f, corporerm.dbo.PVALETR L " & _
"WHERE CHAPA='" & chapa & "' AND F.CODlinha=L.CODIGO and dtfim>getdate() ORDER BY codlinha"
rsq.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<br>
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=420>
	<th class=titulo colspan=7>Cadastro de Vale-Transporte</th>
	<tr>
		<td class=titulor>Nr.de<br>Viagens</td>
		<td class=titulor>Cod.Linha</td>
		<td class=titulor>Nome da Linha</td>
		<td class=titulor>Valor</td>
		<td class=titulor>Início<br>de uso</td>
		<td class=titulor>Término<br>de uso</td>
	</tr>
<%
if rsq.recordcount>0 then
rsq.movefirst
do while not rsq.eof
sql="select nomelinha from corporerm.dbo.pvaletr where codigo='" & rsq("codlinha") & "'"
rsu.open sql, ,adOpenStatic:if rsu.recordcount>0 then nomelinha=trim(rsu("nomelinha")) else nomelinha=""
rsu.close
sql="select valor from ptarifa where codigo='" & rsq("codtarifa") & "' and now between iniciovigencia and finalvigencia "
sql="select valor from corporerm.dbo.ptarifa where codigo='" & rsq("codtarifa") & "' and getdate() between iniciovigencia and finalvigencia "
rsu.open sql, ,adOpenStatic:if rsu.recordcount>0 then valortar=rsu("valor") else valortar=0
rsu.close
%>
	<tr>
		<td class="campor" align="center"><%=rsq("nroviagens")%></td>
		<td class="campor" align="left"><%=rsq("codlinha")%></td>
		<td class="campor" align="left"><%=nomelinha%></td>
		<td class="campor" align="right"><%=formatnumber(valortar,2)%>&nbsp;</td>
		<td class="campor" align="center"><%=rsq("dtinicio")%></td>
		<td class="campor" align="center"><%=rsq("dtfim")%></td>
	</tr>
<%
rsq.movenext
loop
end if
rsq.close
%>
</table>

<%
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>