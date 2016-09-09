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
<title>Alteração de Nomeação <%=session("nomeacao_id")%></title>
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

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		sql="UPDATE n_indicacoes SET "
		sql=sql & "portaria = '"   & request.form("portaria") & "', "
		sql=sql & "id_nomeacao = "   & request.form("id_nomeacao") & ", "
		sql=sql & "chapa    = '"   & request.form("chapa")    & "', "
		sql=sql & "nome     = '"   & request.form("nome")     & "', "
		sql=sql & "cargo    = '"   & request.form("cargo")    & "', "
		sql=sql & "codeve   = '"   & request.form("codeve")    & "', "
		if request.form("mand_ini")<>"" then sql=sql & "mand_ini='" & dtaccess(request.form("mand_ini")) & "', " else sql=sql & "mand_ini=null, "
		if request.form("mand_fim")<>"" then sql=sql & "mand_fim='" & dtaccess(request.form("mand_fim")) & "', " else sql=sql & "mand_fim=null, "
		if request.form("vorig")<>"" then sql=sql & "vorig='" & dtaccess(request.form("vorig")) & "', " else sql=sql & "vorig=null, "
		sql=sql & "alunos   = '"   & request.form("alunos")   & "', "
		sql=sql & "ch       = '"   & nraccess(request.form("ch"))       & "', "
		sql=sql & "complemento= "  & nraccess(request.form("complemento")) & ", "
		sql=sql & "obs      = '"   & request.form("obs")      & "', "
		sql=sql & "coddoc   = '"   & request.form("coddoc")      & "', "
		if request.form("contrato")<>"" then sql=sql & "contrato='" & dtaccess(request.form("contrato")) & "', " else sql=sql & "contrato=null, "
		if request.form("fpg_ini")<>"" then sql=sql & "fpg_ini='" & dtaccess(request.form("fpg_ini")) & "', " else sql=sql & "fpg_ini=null, "
		if request.form("entrega")<>"" then sql=sql & "entrega = '"   & dtaccess(request.form("entrega")) & "', " else sql=sql & "entrega=null, "
		sql=sql & "janeiro  = " & request.form("janeiro") & " "
		sql=sql & "WHERE id_indicado=" & session("id_alt_indicado")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM n_indicacoes WHERE id_indicado=" & session("id_alt_indicado")
		conexao.Execute sql, , adCmdText
	end if

else 'request.form=""

	if request("codigo")=null then
		id_indicado=session("id_alt_indicado")
	else
		id_indicado=request("codigo")
	end if
	sqla="select * from n_indicacoes "
	sqlb="where id_indicado=" & id_indicado
	sql1=sqla & sqlb 
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if
%>

<%
if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_indicado")=rs("id_indicado")
%>
<form method="POST" action="nomeados_alteracao.asp" name="form">
<input type="hidden" name="id_indicado" size="4" value="<%=rs("id_indicado")%>">  
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=grupo>Nomeação para
		<select size="1" name="id_nomeacao" onFocus="this.blur()">
<%
	sqla="SELECT id_nomeacao, nomeacao, criacao, extinta FROM n_nomeacoes "
	sqla=sqla & "where id_nomeacao=" & rs("id_nomeacao") & " ORDER by nomeacao "
	response.write sqla
	rsc.Open sqla,,adOpenStatic, adLockReadOnly
	rsc.movefirst
	do while not rsc.eof
	if rsc("id_nomeacao")=rs("id_nomeacao") then temp="selected" else temp=""
%>
			<option value="<%=rsc("id_nomeacao")%>" <%=temp%>><%=rsc("nomeacao")%></option>
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
	<td class=titulo>Cód.</td>
	<td class=titulo>Portaria de Nomeação</td>
	<td class=titulo>Curso</td>
</tr>
<tr>
	<td class=titulo><%=rs("id_indicado")%></td>
	<td class=fundo><input class=a type="text" name="portaria" size="60" value="<%=rs("portaria")%>"></td>
	<td class=fundo><select size="1" name="coddoc">
		<option value="">...</option>
<%
	sqla="SELECT coddoc from g2cursoeve where tipocurso in (2,6,5) and coddoc>'a' order by tipocurso, coddoc"
	sqla="SELECT e.coddoc from g2cursoeve e inner join g2cursos c on c.coddoc=e.coddoc where tipocurso in (2,6,5) and e.coddoc>'a' group by e.coddoc, tipocurso order by tipocurso, e.coddoc"
	rsc.Open sqla, ,adOpenStatic, adLockReadOnly
	rsc.movefirst:	do while not rsc.eof
	if rsc("coddoc")=rs("coddoc") then temp="selected" else temp=""
%>
		<option value="<%=rsc("coddoc")%>" <%=temp%>><%=rsc("coddoc")%></option>
<%
	rsc.movenext:	loop
	rsc.close
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
	<td class=fundo><input type="text" name="chapa"  size="8" value="<%=rs("chapa")%>" onFocus="this.blur()"></td>
	<td class=fundo><input type="text" name="nome"  size="50" value="<%=rs("nome")%>" onFocus="this.blur()"></td>
<!--onFocus="this.blur()"-->
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Cargo/Curso nomeado</td>
	<td class=titulo>Período do mandato</td>
	<td class=titulo>Carga<br>Horária</td>
	<td class=titulo>Venc.Original</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="cargo"  size="30" value="<%=rs("cargo")%>"></td>
	<td class=fundo><input type="text" name="mand_ini" onchange="mand_ini1(0)"  size="10" value="<%=rs("mand_ini")%>" > <!-- onFocus="javascript:vDateType='3'" onKeyUp="DateFormat(this,this.value,event,false,'3')" onBlur="DateFormat(this,this.value,event,true,'3')" > -->
        a <input type="text" name="mand_fim" onchange=""  size="10" value="<%=rs("mand_fim")%>"  ></td>
	<td class=fundo><input type="text" name="ch"  size="4" value="<%=rs("ch")%>" onchange="mand_ini1(1)"></td>
	<td class=fundo><input type="text" name="vorig"  size="10" value="<%=rs("vorig")%>"  >
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
	"WHERE id_nomeacao=" & rs("id_nomeacao") & " ORDER BY descricao"
	rsc.Open sqla, ,adOpenStatic, adLockReadOnly
	if rsc.recordcount>0 then
	rsc.movefirst:	do while not rsc.eof
	if rsc("codevento")=rs("codeve") then temp1="selected" else temp1=""
%>
		<option value="<%=rsc("codevento")%>" <%=temp1%>><%=rsc("descricao")%></option>
<%
	rsc.movenext:	loop
	end if
	rsc.close
%>
		</select>
	</td>

	<td class=fundo><input type="text" name="fpg_ini" onchange="mand_ini1(1)" size="11" value="<%=rs("fpg_ini")%>"  ></td>
	<td class=fundo><input type="text" name="complemento" size="4" value="<%=rs("complemento")%>"></td>
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
			<td class=titulo><input type="text" name="alunos"  size="5" value="<%=rs("alunos")%>"></td>
			<td class=titulo><font size="1">para TCC<br>
              ou Estágio</font></td>
		</tr>
		</table>
	</td>
	<td class=fundo><input type="text" name="obs"  size="30" value="<%=rs("obs")%>"></td>
	<td class=fundo><input type="text" name="contrato" size="8" value="<%=rs("contrato")%>"  ></td>
	<td class=fundo><input type="text" name="entrega" size="8" value="<%=rs("entrega")%>"  ></td>
	<td class=fundo>
		<select size=1 name="janeiro">
			<option value="0" <%if rs("janeiro")=0 then response.write "selected"%>>Não</option>
			<option value="-1" <%if rs("janeiro")=-1 then response.write "selected"%>>Sim</option>
		</select>
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo align="center"><input type="submit" value="Salvar Alterações  " class="button" name="Bt_salvar"></td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro   " class="button" name="Bt_excluir"></td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end if

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	if tudook=1 then
		response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.location=window.opener.location;self.close();</script>"
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

set rsc=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>