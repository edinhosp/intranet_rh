<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a80")="N" or session("a80")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Data - Grade Horária Pós</title>
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
<script language="javascript" src="cal2.js">
/*
Xin's Popup calendar script-  Xin Yang (http://www.yxscripts.com/)
Script featured on/available at http://www.dynamicdrive.com/
This notice must stay intact for use
*/
</script>
<script language="javascript" src="cal_conf2.js"></script>

<script language="JavaScript" type="text/javascript"><!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome3() {	form.chapa1.value=form.nome1.value;
form.submit();	}
function chapa3() {	form.nome1.value=form.chapa1.value;
form.submit();	}
function nome4() {	form.chapa2.value=form.nome2.value;
form.submit();	}
function chapa4() {	form.nome2.value=form.chapa2.value;
form.submit();	}
--></script>
<script language="VBScript">
</script>

<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
Function IIf(condition,value1,value2)
	If condition Then IIf = value1 Else IIf = value2
End Function

ocorrencia="":tipomov=""

dim conexao, conexao2, chapach, rs, rs2, rs3, a(6)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request("id_grdaula")<>"" then session("id_grdaula")=request("id_grdaula")
session("id_data")=request("id_data"):if session("id_data")="" then session("id_data")="0"

if request.form<>"" then

	if request.form("bt_salvar")<>"" then
		tudook=1
		if (request.form("hora")="0" or request.form("data")="") then
			tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione uma data/horário de aula!');</script>"
		end if
		
		'verificar se a data já existe
		if tudook=1 then
			sql="select * from g2aulasdata where data='" & dtaccess(request.form("data")) & "' and id_grdaula=" & request.form("id_grdaula")
			rs.Open sql, ,adOpenStatic, adLockReadOnly
			if rs.recordcount=0 then 
				existedata=0: observacao="" : session("id_data")="0":codhor="":qtaulas=""
			else 
				existedata=1 : session("id_data")=rs("id_data"):observacao=rs("observacao"):apagada=rs("deletada")
				codhor=rs("codhor"):qtaulas=rs("qtaulas")
			end if
			rs.close
			if isnull(codhor) then codhor=""
			if isnull(qtaulas) then qtaulas=0
			if existedata=0 then
				sql="insert into g2aulasdata (data,id_grdaula,codhor,qtaulas,observacao) " & _
				"select '" & dtaccess(request.form("data")) & "'," & request.form("id_grdaula") & ", " & request.form("codhor") & ", " & request.form("qtaulas") & ", '" & request.form("observacao") & "'"
				conexao.execute sql
				mensagem="<font color=blue>Data inserida.</font>"
				sql="update g2aulas set autorizado=0 where id_grdaula=" & request.form("id_grdaula"):conexao.execute sql:response.write sql
			else
				mensagem="<font color=red>Data já existe.</font>"
				if apagada=true then conexao.execute "update g2aulasdata set deletada=1 where id_data=" & request.form("id_data")
				if request.form("observacao")<>observacao then conexao.execute "update g2aulasdata set observacao='" & request.form("observacao") & "' where id_data=" & request.form("id_data")
				if request.form("codhor")<>codhor then conexao.execute "update g2aulasdata set codhor=" & request.form("codhor") & " where id_data=" & request.form("id_data")
				if request.form("qtaulas")<>qtaulas then conexao.execute "update g2aulasdata set qtaulas=" & request.form("qtaulas") & " where id_data=" & request.form("id_data")
			end if
		end if

		sql="select * from g2aulasdata where data='" & dtaccess(request.form("data")) & "' and id_grdaula=" & request.form("id_grdaula")
		rs.Open sql, ,adOpenStatic, adLockReadOnly
		if rs.recordcount=0 then 
			existedata=0 : session("id_data")="0"
		else 
			existedata=1 : session("id_data")=rs("id_data")
		end if
		rs.close
		
	end if 'button=salvar
	
else 'request.form=""
end if

if request.form("datanova")<>"" then
	tudook=0
	if isdate(request.form("datanova"))=true then
		tudook=1
		sql="update g2aulasdata set data='" & dtaccess(request.form("datanova")) & "' where id_data=" & request.form("id_data") & ""
		if tudook=1 then conexao.execute sql, , adCmdText
		sql="update g2aulas set autorizado=0 where id_grdaula=" & request.form("id_grdaula"):conexao.execute sql,,adCmdText
		end if	
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="delete from g2aulasdata where data='" & dtaccess(request.form("data")) & "' and id_data=" & request.form("id_data")
	if tudook=1 then conexao.execute sql, , adCmdText
	sql="update g2aulas set autorizado=0 where id_grdaula=" & request.form("id_grdaula"):conexao.execute sql,,adCmdText
	mensagem="<font color=red>Data apagada.</font>"
end if

'if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) or (request.form("bt_excluir")<>"" and tudook=0) then
if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("datanova")<>"") or (request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" and (tudook=0 or tudook=1)) or (request.form("bt_excluir")<>"" and tudook=0) then
	'response.write request.form
	sql1="select * from g2aulasdata where id_grdaula=" & session("id_grdaula") & " and id_data=" & session("id_data")
	if session("usuariomaster")="02379" then response.write "->" & sql1	
	rs.Open sql1, ,adOpenStatic, adLockReadOnly

	if rs.recordcount=0 then
		gravado=0
		id_grdaula=request.form("id_grdaula")
		data=request.form("data")
		observacao=request.form("observacao")
		codhor=request.form("codhor")
		qtaulas=request.form("qtaulas")
	else
		gravado=1
		if request.form("id_grdaula")<>"" then id_grdaula =request.form("id_grdaula")  else id_grdaula =rs("id_grdaula")
		if request.form("datanova")<>"" then 
			data=rs("data")
		else
			if request.form("data")<>"" then data=request.form("data") else data=rs("data")
		end if
		if request.form("observacao")<>"" then observacao=request.form("observacao") else observacao=rs("observacao")
		if request.form("codhor")<>"" then codhor=request.form("codhor") else codhor=rs("codhor")
		if request.form("qtaulas")<>"" then qtaulas=request.form("qtaulas") else qtaulas=rs("qtaulas")
	end if
%>
<form method="POST" action="gradenovaposdata.asp" name="form">
<input type="hidden" name="id_grdaula" size="4" value="<%=session("id_grdaula")%>"> 
<input type="hidden" name="id_data"    size="4" value="<%=session("id_data")%>"> 
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">

<!-- quadro -->
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr><td class=campo valign=top width="80%">
<!-- quadro -->

<table border="0" cellpadding="3" cellspacing="0" width="100%">
	<tr><td class=grupo>Lançamento de Datas e Horários (<%=session("id_grdaula")%>)/(<%=session("id_data")%>)</td></tr>
</table>

  
<!--  -->
<table border="1" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo>Data</td>
	<td class=titulo>Horário da aula</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%if isdate(data)=true then response.write formatdatetime(data,2)%>" name="data" size="10" onfocus="javascript:window.status='Informe a data da aula'" onchange="javascript:submit()">
	<%if request.form("data")<>"" then response.write "<br>" & weekdayname(weekday(request.form("data")))%>
	<%if normal=1 then%>
	<small><a href="javascript:showCal('Calendar1')">Data</a></small>
	<%end if%>
	</td>

	<td class=fundo>&nbsp;
		<select size="1" name="codhor" onfocus="javascript:window.status='Selecione um dos horários de aula'" onchange="javascript:submit()">
<%
if data<>"" then diasemana=weekday(data) else diasemana=2
sql2="select codhor, descricao, horini, horfim, codtn from g2defhor where tipocurso in (5,6) and codds=" & diasemana & " order by descricao"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione o horario....</option>"
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
if isnull(codhor) then codhor=""
if cstr(codhor)=cstr(rs2("codhor")) and codhor<>"" then temp="selected" else temp=""
%>
		<option value="<%=rs2("codhor")%>" <%if rs2("codtn")=3 then response.write "style='background:CCFFCC'"%> <%=temp%>><%=rs2("descricao")%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount>0
rs2.close

%>
	</select>
	</td>
</tr>

<tr>
	<td class=fundo colspan=1>Quant.Aulas</td>
	<td class=fundo colspan=1>Obs. desta data</td>
</tr>
<%
if request.form("codhor")<>"" then
	sql3="select qtaulas from g2defhor where codhor=" & request.form("codhor")
	rs2.Open sql3, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then qtaulas=rs2("qtaulas") else qtaulas=0
	rs2.close
end if

%>
<tr>
	<td class=fundo>
		<input type="text" size="4" name="qtaulas" value="<%=qtaulas%>">
	</td>
	<td class=fundo colspan=1>
		<input type="text" size="50" name="observacao" value="<%=observacao%>">
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo align="center">
	<%if request.form("codhor")<>"0" then%>
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
	<%end if%>
	</td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
</tr>
</table>

<!-- quadro -->
</td><td class=campo valign=top width="20%">
<!-- quadro -->

<table border="0" cellpadding="3" cellspacing="0" width="100%">
	<tr><td class=grupo>Ocorrências</td></tr>
	<tr><td class=campo>
	<%
	response.write mensagem
	%>
	Transferir esta data para 
	<input type="text" value="" name="datanova" size="10" onfocus="javascript:window.status='Informe a nova data da aula'" onchange="javascript:submit()">
	</td></tr>
</table>

<!-- quadro para outros horarios -->
</td></tr>

</table>
<!-- quadro -->
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" or request.form("datanova")<>"" then
	if tudook=1 then
		'esponse.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.location.refresh;self.close();</script>"
        response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.document.form.submit();self.close();</script>"
        'response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.document.form.submit();</script>"
        'response.write "<script language='JavaScript' type='text/javascript'>window.opener.document.form.submit();</script>"
	end if
	if tudook=0 then
		'response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser alterado!');</script>"
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

conexao.close
set conexao=nothing
%>
</body>
</html>