<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a21")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inclusão de Nomeação</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">

<script language="JavaScript" type="text/javascript" src="../date.js"></script>
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
function mand_ini1(muda) {
	temp=form.mand_ini.value;
	inicio=new Date(temp.substr(6),temp.substr(3,2)-1,temp.substr(0,2));
	hoje=new Date();
	hoje.setDate(1);hoje.toLocaleString();
	fpgini="0" + hoje.getDate() + "/" + ((hoje.getMonth()+1)<10?"0":"") + (hoje.getMonth()+1) + "/" + hoje.getFullYear();
	//form.fpg_ini.value=fpgini;
	if (muda==1) { temp2=form.fpg_ini.value; hoje=new Date(temp2.substr(6),temp2.substr(3,2)-1,1); }
	dinicio=montharray[inicio.getMonth()]+" "+inicio.getDate()+", "+inicio.getFullYear()
	dmesfp=montharray[hoje.getMonth()]+" "+hoje.getDate()+", "+hoje.getFullYear()
	dias=(Math.round((Date.parse(dmesfp)-Date.parse(dinicio))/(24*60*60*1000))*1)
	semanas=Math.round(dias/7)
	dmesini=montharray[inicio.getMonth()]+" 1, "+inicio.getFullYear()
	if (dmesfp!=dmesini) {
		if (muda==0) { document.form.fpg_ini.value=fpgini }
		horas=document.form.ch.value
		document.form.complemento.value=horas*semanas
	} else {
		document.form.complemento.value=0
		if (muda==0) { document.form.fpg_ini.value=temp }
	}		
}

--></script>

</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		sql = "INSERT INTO n_indicacoes (ID_NOMEACAO, PORTARIA, CHAPA, NOME, CARGO, "
		sql = sql & "ALUNOS, CH, COMPLEMENTO, OBS, codeve, coddoc "
		sql = sql & ",MAND_INI "
		sql = sql & ",MAND_FIM, FPG_INI "
		sql = sql & ",CONTRATO, entrega, janeiro "
		sql = sql & ") "
		sqltemp="select nome from corporerm.dbo.pfunc where chapa='" & request.form("chapa") & "'"
		rsc.open sqltemp, ,adOpenStatic, adLockReadOnly
		if rsc.recordcount=1 then nome=rsc("nome") else nome=""
		rsc.close

		sql2 = " SELECT " & request.form("id_nomeacao") & ", "
		sql2=sql2 & " '" & request.form("portaria") & "', "
		sql2=sql2 & " '" & request.form("chapa") & "', "
		sql2=sql2 & " '" & nome & "', "
		sql2=sql2 & " '" & request.form("cargo") & "', "
		sql2=sql2 & " '" & request.form("alunos") & "', "
		sql2=sql2 & " " & nraccess(request.form("ch")) & ", "
		sql2=sql2 & " " & nraccess(request.form("complemento")) & ", "
		sql2=sql2 & " '" & request.form("obs") & "', "
		sql2=sql2 & " '" & request.form("codeve") & "', "
		sql2=sql2 & " '" & request.form("coddoc") & "', "
		if request.form("mand_ini")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("mand_ini")) & "', "
		if request.form("mand_fim")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("mand_fim")) & "', "
		if request.form("fpg_ini")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("fpg_ini")) & "', "
		if request.form("contrato")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("contrato")) & "', "
		if request.form("entrega")=""  then sql2=sql2 & "null, " else sql2=sql2 & " '" & dtaccess(request.form("entrega")) & "', "
		sql2=sql2 & " " & request.form("janeiro") & " "
		sql1 = sql & sql2 & ""
		'response.write "<font size='1'>" & sql1
		if tudook=1 then conexao.Execute sql1, , adCmdText
	end if
else 'request.form=""
end if
if request.form("id_nomeacao")="" then idnomeacao=request("codigo") else idnomeacao=request.form("id_nomeacao")
if idnomeacao<>"" then idnomeacao=cint(idnomeacao) else idnomeacao=0
if request.form("chapa")="" then idchapa=request("chapa") else idchapa=request.form("chapa")
if request.form="" or request.form("bt_salvar")="" then
%>
<form method="POST" action="nomeados_nova.asp" name="form" >
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=grupo>Nomeação para 
		<select size="1" name="id_nomeacao" onchange="submit()">
<%
	sqla="SELECT id_nomeacao, nomeacao, criacao, extinta FROM n_nomeacoes ORDER by nomeacao"
	rsd.Open sqla, ,adOpenStatic, adLockReadOnly
	rsd.movefirst:	do while not rsd.eof
	if cint(rsd("id_nomeacao"))=idnomeacao then temp="selected" else temp=""
%>
		<option value="<%=rsd("id_nomeacao")%>" <%=temp%>><%=rsd("nomeacao")%></option>
<%
	rsd.movenext:	loop
	rsd.close
%>
		</select>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Cód.</td>
	<td class=titulo>Portaria de Nomeação</td>
	<td class=titulo>Curso</td>
</tr>
<tr>
	<td class=titulo>0</td>
	<td class=fundo><input type="text" name="portaria" size="60"></td>
	<td class=fundo><select size="1" name="coddoc">
		<option value="">...</option>
<%
	sqla="SELECT e.coddoc from g2cursoeve e inner join g2cursos c on c.coddoc=e.coddoc where tipocurso in (2,6,5) and e.coddoc>'a' group by e.coddoc, tipocurso order by tipocurso, e.coddoc"
	rsd.Open sqla, ,adOpenStatic, adLockReadOnly
	rsd.movefirst:	do while not rsd.eof
	if rsd("coddoc")=request.form("coddoc") then temp="selected" else temp=""
%>
		<option value="<%=rsd("coddoc")%>" <%=temp%>><%=rsd("coddoc")%></option>
<%
	rsd.movenext:	loop
	rsd.close
%>
		</select>
	</td>

</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=idchapa%>" name="chapa" size="8" onchange="chapa1()"></td>
	<td class=fundo>
		<select size="1" name="nome" style="font-size: 8 pt" onchange="nome1()">
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' "
'if idchapa<>"" then sql2=sql2 & "and chapa='" & idchapa & "'"
sql2=sql2 & "order by nome"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
if idchapa=rsc("chapa") then temp="selected" else temp=""
%>
		<option value="<%=rsc("chapa")%>" <%=temp%>><%=rsc("nome")%></option>
<%
rsc.movenext
loop
rsc.close
%>
        </select>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Cargo/Curso nomeado</td>
	<td class=titulo>Carga<br>Horária</td>
	<td class=titulo>Período do mandato </td>
	<td class=titulo>Venc.Original</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="cargo"  size="30" value="<%=request.form("cargo")%>"></td>
	<td class=fundo><input type="text" name="ch"  size="5" value="<%=request.form("ch")%>" onchange="mand_ini1(1)"></td>
	<td class=fundo><input type="text" name="mand_ini" size="10" value="<%=request.form("mand_ini")%>" onchange="mand_ini1(0)" > <!-- onFocus="javascript:vDateType='3'" onKeyUp="DateFormat(this,this.value,event,false,'3')" onBlur="DateFormat(this,this.value,event,true,'3')" > -->
        a <input type="text" name="mand_fim"  size="10" value="<%=request.form("mand_fim")%>" ></td>
	<td class=fundo><input type="text" name="vorig" size="10" value="<%=request.form("vorig")%>" >
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Evento</td>
	<td class=titulo>Inicio Pagamento</td>
	<td class=titulo>Complemento</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="codeve">
		<option value="" <%=temp1%>>Sem ônus</option>
<%
	sqla="SELECT codevento, descricao FROM cnv_atividade " & _
	"WHERE id_nomeacao=" & idnomeacao & " ORDER BY descricao"
	rsc.Open sqla, ,adOpenStatic, adLockReadOnly
	if rsc.recordcount>0 then
	rsc.movefirst:	do while not rsc.eof
	if rsc("codevento")=request.form("codeve") then temp1="selected" else temp1=""
%>
		<option value="<%=rsc("codevento")%>" <%=temp1%>><%=rsc("descricao")%></option>
<%
	rsc.movenext:	loop
	end if
	rsc.close
%>
		</select>
	</td>

	<td class=fundo><input type="text" name="fpg_ini"  size="12" value="<%=request.form("fpg_ini")%>" onchange="mand_ini1(1)"  ></td>
	<td class=fundo><input type="text" name="complemento" size="5"  value="<%=request.form("complemento")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Mínimo<br>de alunos</td>
	<td class=titulo>Observação</td>
	<td class=titulo>Data emissão<br>do contrato</td>
	<td class=titulo>Devolução<br>do contrato</td>
	<td class=titulo>Pula<br>Janeiro</td>
</tr>
<tr>
	<td class=fundo>
		<table border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td class=titulo><input type="text" name="alunos" size="3"></td>
			<td class=titulo><font size="1">para TCC<br>ou Estágio</font></td>
		</tr>
		</table>
	</td>
	<td class=fundo><input type="text" name="obs"  size="30" value="<%=request.form("obs")%>"></td>
	<td class=fundo><input type="text" name="contrato" size="8" value="<%=request.form("contrato")%>" ></td>
	<td class=fundo><input type="text" name="entrega" size="8" value="<%=request.form("entrega")%>" ></td>
	<td class=fundo>
		<select size=1 name="janeiro">
			<option value="0" <%if request.form("janeiro")="0" then response.write "selected"%>>Não</option>
			<option value="-1" <%if request.form("janeiro")="-1" then response.write "selected"%>>Sim</option>
		</select>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar   " class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
else
'rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.location=window.opener.location;self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser gravado!');</script>"
	end if

'	Response.write "<p>Registro salvo.<br>"
	'response.write '<script>javascript:top.window.close();</script>
%>
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<!--
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if

set rsc=nothing
set rsd=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>