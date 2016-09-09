<%@ Language=VBScript %>
<!-- #Include file="..\adovbs.inc" -->
<!-- #Include file="..\funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a46")="N" or session("a46")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Horários</title>
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
	conexao.Open application("conexao")
	set rs=server.createobject ("ADODB.Recordset")
	Set rs.ActiveConnection = conexao
	set rsc=server.createobject ("ADODB.Recordset")
	Set rsc.ActiveConnection = conexao

	if request.form<>"" then
		codhorario=request.form("codhorario")
		sqla="SELECT codigo, descricao from corporerm.dbo.ahorario where codigo='" & codhorario & "' "
		set rs2=server.createobject ("ADODB.Recordset")
		Set rs2=conexao.Execute (sqla, , adCmdText)
		horario=rs2("codigo")
		descricao=rs2("descricao")
		session("horario")=codhorario
		temp=0
	else
		temp=1
	end if
%>
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function chapa1() {
	form.codigo.value=form.codhorario.value;
	}
function chapa2() {
	form.codhorario.value=form.codigo.value;
	}
--></script>
<p class=titulo>Conferência de Horários
<%
if temp=1 then
%>
<form method="POST" action="horarios.asp" name="form">
  <p>&nbsp;<input type="text" name="codigo" onchange="chapa2()" size="8" class=a>
  <select size="1" name="codhorario" onchange="chapa1()">
<%
sql2="select codigo, descricao from corporerm.dbo.ahorario order by descricao "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
if request.form("codhorario")=rsc("codigo") then tempz="selected" else tempz=""
%>
          <option value="<%=rsc("codigo")%>" <%=tempz%>><%=rsc("descricao")%></option>
<%
rsc.movenext
loop
rsc.close
%>
        </select>
  <br>
  <input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>

<%
else ' temp=0
sqlb="select max(indice) as loop from corporerm.dbo.abathor where codhorario='" & codhorario & "'"
Set rs2=conexao.Execute (sqlb, , adCmdText)
volta=rs2("loop")
%>
<p>Horário: <b><font color="#800000"> <%=horario%> </font></b> - <font color="#0000FF"> <%=descricao%></font></p>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
  <tr>
    <td class=titulo colspan=2 align="center">Indice</td>
    <td class=titulo colspan=4 align="center">Marcações</td>
    <td class=titulo colspan=2 align="center">Refeição</td>
    <td class=titulo colspan=2 align="center">Limite</td>
  </tr>
  <tr>
    <td class=campo align="center">Dia</td>
    <td class=campo align="center">Tipo</td>
    <td class=campo align="center">Entrada</td>
    <td class=campo align="center">Saida</td>
    <td class=campo align="center">Entrada</td>
    <td class=campo align="center">Saida</td>
    <td class=campo align="center">Inicio</td>
    <td class=campo align="center">Fim</td>
    <td class=campo align="center">Inicio</td>
    <td class=campo align="center">Fim</td>
  </tr>
<%
for voltas=1 to volta 
	sqlb="select batida, tipo from corporerm.dbo.abathor where codhorario='" & codhorario & "' "
	sqlb=sqlb & "and tipo=0 and indice=" & voltas
	sqlb=sqlb & " order by batida "
	rs.Open sqlb, ,adOpenStatic, adLockReadOnly
if rs.recordcount=0 then 'então é dia de descanso ou compensado
	sqld="select batida, tipo, inicio, fim from corporerm.dbo.abathor where codhorario='" & codhorario & "' "
	sqld=sqld & "and tipo in (1,2) and indice=" & voltas
	rsc.Open sqld, ,adOpenStatic, adLockReadOnly
	if rsc.recordcount>0 then
	tipoz=rsc("tipo")
	if tipoz=1 then tipo="Descanso"
	if tipoz=2 then tipo="Compensado"
	inicio=rsc("inicio")
	hora1=formatdatetime((inicio/60)/24,4)
	fim=rsc("fim")
	hora2=formatdatetime((fim/60)/24,4)
	end if
	rsc.close
%>
  <tr>
    <td class=campo align="center"><%=voltas%></td>
    <td class=campo align="center"><%=tipo%></td>
    <td class=campo colspan=6 align="center">-</td>
    <td class=campo align="center"><%=hora1%></td>
    <td class=campo align="center"><%=hora2%></td>
  </tr>
<%
else 'recordcount > 0
	quant=rs.recordcount
%>
  <tr>
    <td class=campo align="center"><%=voltas%></td>
    <td class=campo align="center">Normal</td>
<%
	batidas=0
	rs.movefirst
	do while not rs.eof
	batida=rs("batida")
	hora3=formatdatetime((batida/60)/24,4)
%>
	<td class=campo align="center"><%=hora3 %></td>
<%	
	batidas=batidas+1
	rs.movenext
	loop
	if batidas<4 then
	for a=batidas+1 to 4
%>
    <td class=campo align="center"></td>
<%
	next
	end if 'batidas=4

	sqld="select batida, tipo, inicio, fim from corporerm.dbo.abathor where codhorario='" & codhorario & "' "
	sqld=sqld & "and tipo in (4) and indice=" & voltas
	rsc.Open sqld, ,adOpenStatic, adLockReadOnly
	if rsc.recordcount=0 then
		response.write "<td colspan=4 align=""center""><font size=1>-</font></td>"
		rsc.close
	else
		inicio=rsc("inicio")
		hora1=formatdatetime((inicio/60)/24,4)
		fim=rsc("fim")
		hora2=formatdatetime((fim/60)/24,4)
		rsc.close
%>
    <td class=campo align="center"><%=hora1%></td>
    <td class=campo align="center"><%=hora2%></td>
<%
		sqld="select batida, tipo, inicio, fim from corporerm.dbo.abathor where codhorario='" & codhorario & "' "
		sqld=sqld & "and tipo in (5) and indice=" & voltas
		rsc.Open sqld, ,adOpenStatic, adLockReadOnly
		inicio=rsc("inicio")
		hora1=formatdatetime((inicio/60)/24,4)
		fim=rsc("fim")
		hora2=formatdatetime((fim/60)/24,4)
		rsc.close
%>
    <td class=campo align="center"><%=hora1%></td>
    <td class=campo align="center"><%=hora2%></td>
   </tr>
<%
	end if 'recordcount >0 para refeicao
end if 'recordcount >0
	rs.close
next

%>
</table>
<p><font color="#0000FF">Funcionários que estão neste horário:</font> 
<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse" width="650">
  <tr>
    <td class=titulo align="center">Chapa</td>
    <td class=titulo align="center">Nome</td>
    <td class=titulo align="center">Seção</td>
    <td class=titulo align="center">Setor</td>
  </tr>
<%
sqld="select f.chapa, f.nome, f.codsecao, s.descricao "
sqld=sqld & "from corporerm.dbo.pfunc f, corporerm.dbo.psecao s "
sqld=sqld & "where f.codsecao=s.codigo and f.codsituacao<>'D' and f.codtipo='N' "
sqld=sqld & "and f.codhorario='" & codhorario & "' "
rsc.Open sqld, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
rsc.movefirst
do while not rsc.eof
%>
  <tr>
    <td class=campo><%=rsc("chapa")%></td>
    <td class=campo><%=rsc("nome")%></td>
    <td class=campo><%=rsc("codsecao")%></td>
    <td class=campo><%=rsc("descricao")%></td>
  </tr>
<%
rsc.movenext
loop
else
%>
	<tr><td class=campo colspan=4>Não existem funcionários com este horário.</td></tr>
<%
end if
rsc.close
%>
</table>

<p><font color="#0000FF">Históricos deste horário:</font> 
<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650">
  <tr>
    <td class=titulo align="center">Chapa</td>
    <td class=titulo align="center">Nome</td>
    <td class=titulo align="center">Data</td>
    <td class=titulo align="center">Indice</td>
    <td class=titulo align="center">Situação</td>
  </tr>
<%
sqld="select f.chapa, f.nome, h.dtmudanca, h.codhorario, f.codsituacao, h.indiniciohor "
sqld=sqld & "from corporerm.dbo.pfunc f, corporerm.dbo.pfhsthor h "
sqld=sqld & "where f.chapa=h.chapa "
sqld=sqld & "and h.codhorario='" & codhorario & "' "
rsc.Open sqld, ,adOpenStatic, adLockReadOnly
if rsc.recordcount>0 then
rsc.movefirst
do while not rsc.eof
%>
  <tr>
    <td class=campo><%=rsc("chapa")%></td>
    <td class=campo><%=rsc("nome")%></td>
    <td class=campo align="center"><%=rsc("dtmudanca")%></td>
    <td class=campo align="center"><%=rsc("indiniciohor")%></td>
    <td class=campo align="center"><%=rsc("codsituacao")%></td>
  </tr>
<%
rsc.movenext
loop
else
%>
	<tr><td class=campo colspan=4>Não existe este horário nos históricos.</td></tr>
<%
end if
rsc.close
%>
</table>


<%
end if 'temp=0
%>
</body>
</html>
<%
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%>