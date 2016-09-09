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
<title>Inclusão de Requisição de Pessoal</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
function nome1()	{	form.chapa.value=form.nome.value;	}
function chapa1()	{	form.nome.value=form.chapa.value;	}
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

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1

if request.form("descricao")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Informe uma descrição para a vaga!');</script>"
if request.form("funcao")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione a função requisitada!');</script>"
if request.form("secao")="" then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione a seção requisitante!');</script>"
		
		if request.form("exp_cumpre")="ON" then exp_cumpre=1 else exp_cumpre=0
		sql = "INSERT INTO rs_requisicao (" 
		sql = sql & "descricao, funcao, secao, requisitante, motivo, chapasubst, id_area, id_faixa, "
		sql = sql & "exp_cumpre, salario, tipo, horario, escolaridade, idademin, idademax, "
		sql = sql & "experiencia, sexo, cursos, deficiente, tp_def, outros, dt_abertura, "
		sql = sql & "dt_encerramento, faixa, qt_vagas "
		sql = sql & ") "
		sql2 = " SELECT "
		sql2=sql2 & " '" & request.form("descricao") & "', "
		sql2=sql2 & " '" & request.form("funcao") & "', "
		sql2=sql2 & " '" & request.form("secao") & "', "
		sql2=sql2 & " '" & request.form("requisitante") & "', "
		sql2=sql2 & " '" & request.form("motivo") & "', "
		sql2=sql2 & " '" & request.form("chapasubst") & "', "
		sql2=sql2 & " 0," ' request.form("id_area") 
		sql2=sql2 & " 0," ' request.form("id_faixa") 
		sql2=sql2 & " " & exp_cumpre & ", "
		sql2=sql2 & " " & nraccess(cdbl(request.form("salario"))) & ", "
		sql2=sql2 & " '" & request.form("tipo") & "', "
		sql2=sql2 & " '" & request.form("horario") & "', "
		sql2=sql2 & " '" & request.form("escolaridade") & "', "
		sql2=sql2 & " " & request.form("idademin") & ", "
		sql2=sql2 & " " & request.form("idademax") & ", "
		sql2=sql2 & " " & request.form("experiencia") & ", "
		sql2=sql2 & " '" & request.form("sexo") & "', "
		sql2=sql2 & " '" & request.form("cursos") & "', "
		sql2=sql2 & " '" & request.form("deficiente") & "', "
		sql2=sql2 & " '" & request.form("tp_def") & "', "
		sql2=sql2 & " '" & request.form("outros") & "', "
		if request.form("dt_abertura")=""     then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("dt_abertura")) & "', "
		if request.form("dt_encerramento")="" then sql2=sql2 & "null," else sql2=sql2 & " '" & dtaccess(request.form("dt_encerramento")) & "', "
		sql2=sql2 & " '" & request.form("faixa") & "', "
		sql2=sql2 & " " & request.form("qt_vagas") & " "

		sql1 = sql & sql2 & ""
		'response.write "<font size='2'>" & sql1
		if tudook=1 then conexao.Execute sql1, , adCmdText
	end if
else 'request.form=""
end if

if request.form="" or (request.form<>"" and tudook=0) then
%>
<form method="POST" action="requisicao_nova.asp" name="form" >
<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=grupo>Inclusão de Requisição de Pessoal</td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=titulo>Descrição para a Requisição</td>
	<td class=fundo><input type="text" name="descricao" size="50" value="<%=request.form("descricao")%>" ></td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=titulo>Função</td>
	<td class=fundo><input type="text" name="cfuncao" size="6" onchange="funcao1('c')" value="<%=request.form("cfuncao")%>" >
	<select size="1" name="funcao" class=a onchange="funcao1('n')">
		<option value="0">Selecione uma função</option>
<%
sql2="select codigo, nome from corporerm.dbo.pfuncao order by nome "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if request.form("funcao")=rsc("codigo") then textof="selected" else textof=""
%>
          <option value="<%=rsc("codigo")%>" <%=textof%> ><%=rsc("nome")%></option>
<%
rsc.movenext:loop
rsc.close
%>
        </select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=titulo>Seção</td>
	<td class=fundo><input type="text" name="csecao" size="8" onchange="secao1('c')" value="<%=request.form("csecao")%>">
	<select size="1" name="secao" class=a onchange="secao1('n')">
		<option value="0">Selecione uma seção</option>
<%
sql2="select codigo, descricao from corporerm.dbo.psecao order by descricao "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if request.form("secao")=rsc("codigo") then textos="selected" else textos=""
%>
          <option value="<%=rsc("codigo")%>" <%=textos%> ><%=rsc("descricao")%></option>
<%
rsc.movenext:loop
rsc.close
%>
        </select></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=titulo>Requisitante</td>
	<td class=fundo><input type="text" name="crequisitante" size="6" onchange="requisitante1('c')" value="<%=request.form("crequisitante")%>">
	<select size="1" name="requisitante" class=a onchange="requisitante1('n')">
		<option value="0">Selecione a pessoa requisitante</option>
<%
sql2="SELECT C.CHAPASUBST, F.NOME FROM corporerm.dbo.PSUBSTCHEFE C INNER JOIN corporerm.dbo.PFUNC F ON C.CHAPASUBST=F.CHAPA GROUP BY C.CHAPASUBST, F.NOME ORDER BY F.NOME "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if request.form("requisitante")=rsc("chapasubst") then textos="selected" else textos=""
%>
          <option value="<%=rsc("chapasubst")%>" <%=textos%> ><%=rsc("nome")%></option>
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
 		<option value="02">Substituição</option>
	 	<option value="03">Vaga Nova</option>
 		<option value="04">Aumento de quadro</option>
	</select></td>
	<td class=fundo><input type="text" name="cchapasubst" size="6" onchange="chapas('c')" value="<%=request.form("cchapasubst")%>">
	<select size="1" name="chapasubst" class=a onchange="chapas('n')">
		<option value="0">Selecione a pessoa substituída</option>
<%
sql2="SELECT chapa, nome from corporerm.dbo.pfunc where (chapa<'10000' or chapa>'89999') and  " & _
" codsituacao<>'D' or (codsituacao='D' and datademissao>dateadd(m,-1,convert(nvarchar,year(getdate()))+'/'+convert(nvarchar,month(getdate())-0)+'/01'))  " & _
" order by nome"

rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if request.form("chapasubst")=rsc("chapa") then textoc="selected" else textoc=""
%>
          <option value="<%=rsc("chapa")%>" <%=textoc%> ><%=rsc("nome")%></option>
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
	<td class=titulo>Faixa&nbsp;</td>
</tr>
<tr>
	<td class=fundo>
	<input type="radio" name="tipo" value="1" checked>Normal<br>
	<input type="radio" name="tipo" value="2">Estagiário
	</td>
	<td class=fundo valign=top><input type="checkbox" name="exp_cumpre" value="ON">Vai cumprir experiência?</td>
	<td class=fundo valign=top>Base&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;<input type="text" name="salario" size="8" class=vr value="<%=request.form("salario")%>"><br>
	Admissão:&nbsp;<input type="text" name="salario_exp" size="8" class=bloq value="<%=request.form("salario_exp")%>">
	</td>
	<td class=fundo valign=top>
		<select name="faixa"><option value=""></option><option value="N1">N1</option>
		<option value="N2">N2</option><option value="N3">N3</option><option value="N4">N4</option>
		<option value="N5">N5</option></select>
	</td>
</tr>
</table>  
  
<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr><td class=titulo>Horário</td>
	<td class=fundo><input type="text" name="chorario" size="3" onchange="horario1('c')" value="<%=request.form("chorario")%>">
	<select size="1" name="horario" class=small onchange="horario1('n')">
		<option value="0">Selecione um horário</option>
<%
sql2="SELECT codigo, descricao from corporerm.dbo.ahorario order by descricao "
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if request.form("horario")=rsc("codigo") then textoh="selected" else textoh=""
%>
		<option value="<%=rsc("codigo")%>" <%=testeh%>><%=rsc("descricao")%></option>
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
	<td class=fundo><input type="text" name="cescolaridade" size="3" onchange="escolaridade1('c')" value="<%=request.form("cescolaridade")%>">
	<select size="1" name="escolaridade" class=a onchange="escolaridade1('n')">
		<option value="0">Selecione a escolaridade</option>
<%
sql2="SELECT codcliente, descricao from corporerm.dbo.pcodinstrucao order by codcliente"
rsc.Open sql2, ,adOpenStatic, adLockReadOnly
rsc.movefirst:do while not rsc.eof
if request.form("escolaridade")=rsc("codcliente") then tempe="selected" else tempe=""
%>
          <option value="<%=rsc("codcliente")%>" <%=tempe%>><%=rsc("descricao")%></option>
<%
rsc.movenext:loop
rsc.close
%>
        </select></td>
    <td class=fundo>Mínima <input type="text" name="idademin" size="3" value=16>
		Máxima <input type="text" name="idademax" size="3" value=99></td>
	<td class=fundo>
	<select name="sexo">
 	<option value="I"> Indiferente</option>
 	<option value="F"> Feminino</option>
 	<option value="M"> Masculino</option>
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
	<td class=fundo valign=top><input type="text" name="experiencia" size="3" value=1> anos</td>
	<td class=fundo><textarea name="cursos" cols="30" rows="3"></textarea></td>
	<td class=fundo>
	<select name="deficiente">
 	<option value="0">Indiferente</option>
 	<option value="1">Não Deficiente</option>
 	<option value="2">Deficiente</option>
	</select><br>
	Tipo Deficiência: <input type="text" name="tp_def" size="15" value="<%=request.form("tp_def")%>">
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
	<td class=fundo valign=top><textarea name="outros" cols="40" rows="3" value="<%=request.form("outros")%>" ></textarea></td>
	<td class=fundo valign=top><input type="text" name="dt_abertura" size="8" value=<%=formatdatetime(now,2)%> value="<%=request.form("dt_abertura")%>"></td>
	<td class=fundo valign=top><input type="text" name="dt_encerramento" size="8" value="<%=request.form("dt_encerramento")%>"></td>
	<td class=fundo valign=top><input type="text" name="qt_vagas" size="3" value=1></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="600">
<tr>
	<td class=titulo align="center">
		<input type="submit" value="Salvar Registro" class="button" name="Bt_salvar"></td>
	<td class=titulo align="center">
		<input type="reset"  value="Desfazer Alterações" class="button" name="B2"></td>
	<td class=titulo align="center">
		<input type="button" value="Fechar" class="button" name="Bt_fechar" onClick="top.window.close()"></td>
</tr>
</table>
</form>
<%
else
'rs.close
set rs=nothing
end if
conexao.close
set conexao=nothing

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
%>
</body>
</html>