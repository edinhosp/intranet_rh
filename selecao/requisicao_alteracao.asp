<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a68")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Requisição de Pessoal</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function funcao1(tipo) { 
	if (tipo=='c') { form.funcao.value=form.cfuncao.value }
	if (tipo=='n') { form.cfuncao.value=form.funcao.value }
	}
function secao1(tipo) { 
	if (tipo=='c') { form.secao.value=form.csecao.value }
	if (tipo=='n') { form.csecao.value=form.secao.value }
	}
function requisitante1(tipo) { 
	if (tipo=='c') { form.requisitante.value=form.crequisitante.value }
	if (tipo=='n') { form.crequisitante.value=form.requisitante.value }
	}
function chapas(tipo) { 
	if (tipo=='c') { form.chapasubst.value=form.cchapasubst.value }
	if (tipo=='n') { form.cchapasubst.value=form.chapasubst.value }
	}
function horario1(tipo) { 
	if (tipo=='c') { form.horario.value=form.chorario.value }
	if (tipo=='n') { form.chorario.value=form.horario.value }
	}
function escolaridade1(tipo) { 
	if (tipo=='c') { form.escolaridade.value=form.cescolaridade.value }
	if (tipo=='n') { form.cescolaridade.value=form.escolaridade.value }
	}
--></script>
<script language="VBScript">
	Sub motivo_onChange
		temp=document.form.motivo.value
		if temp="02" then 
			document.form.chapasubst.disabled=false 
			document.form.cchapasubst.disabled=false 
		else 
			document.form.chapasubst.disabled=true
			document.form.cchapasubst.disabled=true
		end if
	End Sub
	Sub salario_onChange
		temp=document.form.salario.value
		temp2=document.form.exp_cumpre.checked
		if temp2=true then fator=0.95 else fator=1
		document.form.salario.value=formatnumber(temp,2)
		document.form.salario_exp.value=formatnumber(temp*fator,2)
	End Sub
	Sub exp_cumpre_onClick
		temp=document.form.salario.value
		temp2=document.form.exp_cumpre.checked
		if temp2=true then fator=0.95 else fator=1
		document.form.salario_exp.value=formatnumber(temp*fator,2)
	End Sub
</script>

</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

	if request.form("bt_salvar")<>"" then
		tudook=1

if request.form("descricao")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe uma descrição para a vaga!');</script>"
if request.form("funcao")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione a função requisitada!');</script>"
if request.form("secao")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione a seção requisitante!');</script>"

		sql="UPDATE rs_requisicao SET "
		sql=sql & "descricao   ='" & request.form("descricao") & "', "
		sql=sql & "funcao      ='" & request.form("funcao") & "', "
		sql=sql & "secao       ='" & request.form("secao") & "', "
		sql=sql & "requisitante='" & request.form("requisitante") & "', "
		sql=sql & "motivo      ='" & request.form("motivo") & "', "
		sql=sql & "chapasubst  ='" & request.form("chapasubst") & "', "
		sql=sql & "id_area     =0," ' request.form("id_area") 
		sql=sql & "id_faixa    =0," ' request.form("id_faixa") 
		sql=sql & "faixa       ='" & request.form("faixa") & "', "
		if request.form("exp_cumpre")="ON" then 
			sql=sql & "exp_cumpre = 1, " 
		else
			sql=sql & "exp_cumpre = 0, "
		end if
		sql=sql & "salario     = " & nraccess(cdbl(request.form("salario"))) & ", "
		sql=sql & "tipo        ='" & request.form("tipo") & "', "
		sql=sql & "horario     ='" & request.form("horario") & "', "
		sql=sql & "escolaridade='" & request.form("escolaridade") & "', "
		sql=sql & "idademin    = " & request.form("idademin") & ", "
		sql=sql & "idademax    = " & request.form("idademax") & ", "
		sql=sql & "experiencia = " & request.form("experiencia") & ", "
		sql=sql & "sexo        ='" & request.form("sexo") & "', "
		sql=sql & "cursos      ='" & request.form("cursos") & "', "
		sql=sql & "deficiente  ='" & request.form("deficiente") & "', "
		sql=sql & "tp_def      ='" & request.form("tp_def") & "', "
		sql=sql & "outros      ='" & request.form("outros") & "', "
		if request.form("dt_abertura")=""     then
			sql=sql & "dt_abertura=null,"
		else
			sql=sql & "dt_abertura='" & dtaccess(request.form("dt_abertura")) & "', "
		end if
		if request.form("dt_encerramento")="" then
			sql=sql & "dt_encerramento=null,"
		else
			sql=sql & "dt_encerramento='" & dtaccess(request.form("dt_encerramento")) & "', "
		end if
		sql=sql & "qt_vagas    = " & request.form("qt_vagas") & " "
		sql=sql & " WHERE id_requisicao=" & session("id_alt_requisicao")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

	if request.form("bt_excluir")<>"" then
		tudook=1
		sql="DELETE FROM rs_requisicao WHERE id_requisicao=" & session("id_alt_requisicao")
		if tudook=1 then conexao.Execute sql, , adCmdText
	end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	if request("codigo")=null then
		id_requisicao=session("id_alt_requisicao")
	else
		id_requisicao=request("codigo")
	end if
	sqla="select * from rs_requisicao "
	sqlb="where id_requisicao=" & id_requisicao
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_requisicao")=rs("id_requisicao")
if rs("exp_cumpre")=0 then exp_cumpre="" else exp_cumpre="checked"
%>
<form method="POST" action="requisicao_alteracao.asp" name="form">
<input type="hidden" name="id_requisicao" size="4" value="<%=rs("id_requisicao")%>" style="font-size: 8 pt" >  
<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=grupo>Alteração de Requisição de Pessoal - <%=rs("id_requisicao")%></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=titulo>Descrição para a Requisição</td>
	<td class=fundo><input type="text" name="descricao" size="50" value="<%=rs("descricao")%>"></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=titulo>Função</td>
	<td class=fundo><input type="text" name="cfuncao" size="6" value="<%=rs("funcao")%>" onchange="funcao1('c')">
	<select size="1" name="funcao" class=a onchange="funcao1('n')">
		<option value="0">Selecione uma função</option>
<%
sql2="select codigo, nome from corporerm.dbo.pfuncao order by nome "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if rs("funcao")=rsc("codigo") then tmpfuncao="selected" else tmpfuncao=""
%>
          <option value="<%=rsc("codigo")%>" <%=tmpfuncao%>><%=rsc("nome")%></option>
<%
rsc.movenext:loop
rsc.close
%>
        </select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=titulo>Seção</td>
	<td class=fundo><input type="text" name="csecao" size="8" value="<%=rs("secao")%>" onchange="secao1('c')">
	<select size="1" name="secao" class=a onchange="secao1('n')">
		<option value="0">Selecione uma seção</option>
<%
sql2="select codigo, descricao from corporerm.dbo.psecao order by descricao "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if rs("secao")=rsc("codigo") then tmpsecao="selected" else tmpsecao=""
%>
          <option value="<%=rsc("codigo")%>" <%=tmpsecao%>><%=rsc("descricao")%></option>
<%
rsc.movenext:loop
rsc.close
%>
        </select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=titulo>Requisitante</td>
    <td class=fundo><input type="text" name="crequisitante" size="6" value="<%=rs("requisitante")%>" onchange="requisitante1('c')">
	<select size="1" name="requisitante" class=a onchange="requisitante1('n')">
		<option value="0">Selecione a pessao requisitante</option>
<%
sql2="SELECT C.CHAPASUBST, F.NOME FROM corporerm.dbo.PSUBSTCHEFE C INNER JOIN corporerm.dbo.PFUNC F ON C.CHAPASUBST=F.CHAPA GROUP BY C.CHAPASUBST, F.NOME ORDER BY F.NOME "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if rs("requisitante")=rsc("chapasubst") then tmpchapa="selected" else tmpchapa=""
%>
          <option value="<%=rsc("chapasubst")%>" <%=tmpchapa%>><%=rsc("nome")%></option>
<%
rsc.movenext:loop
rsc.close
%>
        </select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr>
    <td class=titulo>Motivo</td>
    <td class=titulo>Funcionário Substituido </td>
</tr>
<tr>
    <td class=fundo><select name="motivo">
 		<option value="02" <%if rs("motivo")="02" then response.write "selected"%>>Substituição</option>
	 	<option value="03" <%if rs("motivo")="03" then response.write "selected"%>>Vaga Nova</option>
 		<option value="04" <%if rs("motivo")="04" then response.write "selected"%>>Aumento de quadro</option>
	</select></td>
    <td class=fundo><input type="text" name="cchapasubst" size="6" value="<%=rs("chapasubst")%>" onchange="chapas('c')">
	<select size="1" name="chapasubst" class=a onchange="chapas('n')">
		<option value="0">Selecione a pessoa substituída</option>
<%
sql2="SELECT chapa, nome from corporerm.dbo.pfunc where (chapa<'10000' or chapa>'89999') and codtipo not in ('A') " & _
" order by nome"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if rs("chapasubst")=rsc("chapa") then tmpsubst="selected" else tmpsubst=""
%>
          <option value="<%=rsc("chapa")%>" <%=tmpsubst%>><%=rsc("nome")%></option>
<%
rsc.movenext:loop
rsc.close
%>
        </select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr>
    <td class=titulo>Tipo</td>
    <td class=titulo>&nbsp;</td>
    <td class=titulo>Salário</td>
    <td class=titulo>Faixa</td>
</tr>
<tr>
    <td class=fundo>
	<input type="radio" name="tipo" value="1" <%if rs("tipo")="1" then response.write "checked"%>>Normal<br>
	<input type="radio" name="tipo" value="2" <%if rs("tipo")="2" then response.write "checked"%>>Estagiário
	</td>
<%
if rs("salario")="" or isnull(rs("salario")) then salario=0 else salario=cdbl(rs("salario"))
if rs("exp_cumpre")=1 and rs("tipo")=1 then fator=0.95 else fator=1
salario_exp=cdbl(salario*fator)
%>
    <td class=fundo valign=top><input type="checkbox" name="exp_cumpre" value="ON" <%=exp_cumpre%>>Vai cumprir experiência?</td>
    <td class=fundo valign=top>Base&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;<input type="text" name="salario" size="8" value="<%=formatnumber(salario,2)%>" class=vr><br>
	Admissão:&nbsp;<input type="text" name="salario_exp" size="8" value="<%=formatnumber(salario_exp,2)%>" class=bloq>
	</td>
    <td class=fundo valign=top>
		<select name="faixa">
		<option value=""></option>
		<option value="N1" <%if rs("faixa")="N1" then response.write "selected"%>>N1</option>
		<option value="N2" <%if rs("faixa")="N2" then response.write "selected"%>>N2</option>
		<option value="N3" <%if rs("faixa")="N3" then response.write "selected"%>>N3</option>
		<option value="N4" <%if rs("faixa")="N4" then response.write "selected"%>>N4</option>
		<option value="N5" <%if rs("faixa")="N5" then response.write "selected"%>>N5</option>
		</select>
	</td>
  </tr>
</table>  
  
<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=titulo>Horário</td>
    <td class=fundo><input type="text" name="chorario" size="3" value="<%=rs("horario")%>" onchange="horario1('c')">
	<select size="1" name="horario" class=small onchange="horario1('n')">
		<option value="0">Selecione um horário</option>
<%
sql2="SELECT codigo, descricao from corporerm.dbo.ahorario order by descricao "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if rs("horario")=rsc("codigo") then tmphor="selected" else tmphor=""
%>
          <option value="<%=rsc("codigo")%>" <%=tmphor%>><%=rsc("descricao")%></option>
<%
rsc.movenext:loop
rsc.close
%>
        </select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=grupo>Requisitos</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr>
    <td class=titulo>Escolaridade Mínima</td>
    <td class=titulo>Idade</td>
    <td class=titulo>Sexo</td>
</tr>
<tr>
    <td class=fundo><input type="text" name="cescolaridade" size="3" value="<%=rs("escolaridade")%>" onchange="escolaridade1('c')">
	<select size="1" name="escolaridade" class=a onchange="escolaridade1('n')">
		<option value="0">Selecione a escolaridade</option>
<%
sql2="SELECT codcliente, descricao from corporerm.dbo.pcodinstrucao order by codcliente"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if rs("escolaridade")=rsc("codcliente") then tmpesc="selected" else tmpesc=""
%>
          <option value="<%=rsc("codcliente")%>" <%=tmpesc%>><%=rsc("descricao")%></option>
<%
rsc.movenext:loop
rsc.close
%>
        </select></td>
    <td class=fundo>Mínima <input type="text" name="idademin" size="3" value="<%=rs("idademin")%>" class=vr>
		Máxima <input type="text" name="idademax" size="3" value="<%=rs("idademax")%>" class=vr></td>
	<td class=fundo>
	<select name="sexo">
 	<option value="I" <%if rs("sexo")="I" then response.write "selected"%>> Indiferente</option>
 	<option value="F" <%if rs("sexo")="F" then response.write "selected"%>> Feminino</option>
 	<option value="M" <%if rs("sexo")="M" then response.write "selected"%>> Masculino</option>
	</select>
 
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr>
    <td class=titulo>Experiência Mínima</td>
    <td class=titulo>Cursos Exigidos</td>
    <td class=titulo>Deficiente</td>
</tr>
<tr>
    <td class=fundo valign=top><input type="text" name="experiencia" size="3" value="<%=rs("experiencia")%>"> anos</td>
    <td class=fundo><textarea name="cursos" cols="30" rows="3"><%=rs("cursos")%></textarea>
	</td>
	<td class=fundo>
	<select name="deficiente">
 	<option value="0" <%if rs("deficiente")="0" then response.write "selected"%>>Indiferente</option>
 	<option value="1" <%if rs("deficiente")="1" then response.write "selected"%>>Não Deficiente</option>
 	<option value="2" <%if rs("deficiente")="2" then response.write "selected"%>>Deficiente</option>
	</select><br>
	Tipo Deficiência: <input type="text" name="tp_def" size="15" value="<%=rs("tp_def")%>">
	</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr>
    <td class=titulo>Outros Requisitos</td>
    <td class=titulo>Dt.Abertura</td>
    <td class=titulo>Dt.Encerramento</td>
    <td class=titulo>Qt.Vagas</td>
</tr>
<tr>
    <td class=fundo valign=top><textarea name="outros" cols="40" rows="3"><%=rs("outros")%></textarea>
	</td>
	<td class=fundo valign=top><input type="text" name="dt_abertura" size="8" value="<%=rs("dt_abertura")%>"></td>
	<td class=fundo valign=top><input type="text" name="dt_encerramento" size="8" value="<%=rs("dt_encerramento")%>"></td>
	<td class=fundo valign=top><input type="text" name="qt_vagas" size="3" value="<%=rs("qt_vagas")%>"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
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
set rsc=nothing
set rsnome=nothing
conexao.close
set conexao=nothing

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
%>
</body>
</html>