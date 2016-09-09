<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a1")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Previa Curriculo</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:5px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("mysqlfieo")
'conexao.open "Driver={MySQL ODBC 3.51 Driver}; Server=localhost; Port=3306; Option=0; Socket=; Stmt=; Database=rhonline2; Uid=root; Pwd="
'conexao.open "Driver={MySQL ODBC 3.51 Driver}; Server=colossus2.fieo.br; Port=3306; Option=0; Socket=; Stmt=; Database=website; Uid=rh; Pwd=!@#qaz"
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

sessao=session.sessionid
cpf=request("codigo")
sql="SELECT * FROM tb_rh_candidato t where cpf='" & cpf & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
select case rs("tipo_candidato")
	case 1
		area="Docente"
	case 2
		area="Administrativa"
	case else
		area="Não definida"
end select
if rs("deficiente")="0" then deficiente="Não" else deficiente="Sim"
if rs("sexo")="1" then sexo="Masculino" else sexo="Feminino"

%>
<table border="0" cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td><b>CURRICULO</b><td>
</tr>
</table>

<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop" width=450><i>Nome</i><br><b><%=rs("nome")%></b>&nbsp;
	</td>
	<td class="campop"><i>Cadastro</i><br><%=rs("data_cadastro")%>&nbsp;
	</td>
	<td class="campop"><i>Area</i><br><%=area%>&nbsp;
	</td>
</tr>
</table>

<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop"><i>Deficiente?</i><br><%=deficiente%>&nbsp;
	</td>
	<td class="campop"><i>Sexo</i><br><%=sexo%>&nbsp;
	</td>
	<td class="campop"><i>Nascimento</i><br><%=rs("nascimento")%>&nbsp;
	</td>
</tr>
</table>

<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop"><i>Endereço</i><br><%=rs("endereco") & " " & rs("complemento")%>&nbsp;
	</td>
	<td class="campop"><i>Bairro</i><br><%=rs("bairro")%>&nbsp;
	</td>
	<td class="campop"><i>Cidade</i><br><%=rs("cidade")%>&nbsp;
	</td>
	<td class="campop"><i>UF</i><br><%=rs("uf")%>&nbsp;
	</td>
	<td class="campop"><i>CEP</i><br><%=rs("cep")%>&nbsp;
	</td>
</tr>
</table>

<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop"><i>Tel. Residencial</i><br><%=rs("tel_residencial")%>&nbsp;
	</td>
	<td class="campop"><i>Tel. Comercial</i><br><%=rs("tel_comercial")%>&nbsp;
	</td>
	<td class="campop"><i>Tel. Celular</i><br><%=rs("tel_celular")%>&nbsp;
	</td>
	<td class="campop"><i>Email</i><br><%=rs("email")%>&nbsp;
	</td>
</tr>
</table>

<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr>
	<td class="campop"><i>Observações</i><br><%=rs("observacoes")%>&nbsp;
	</td>
	<td class="campop"><i>Lattes</i><br><%if rs("lattes")<>"" then%><a href="<%=rs("lattes")%>" target=_blank><%=rs("lattes")%></a><%end if%>&nbsp;
	</td>
</tr>
</table>
<%
rs.close
%>

<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=grupo colspan=4>Formação</td></tr>
<tr>
	<td class=titulo>Tipo</td>
	<td class=titulo>Nome Curso</td>
	<td class=titulo>Local</td>
	<td class=titulo>Ano</td>
</tr>
<%
sql="SELECT f.cpf, f.nivel_curso, n.titulo, f.nome_curso, f.local_curso, f.ano_conclusao FROM tb_rh_formacao f " & _
"inner join tb_rh_ncurso n on n.id_ncurso=f.nivel_curso where cpf='" & cpf & "' order by nivel_curso"
rs.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
<tr>
	<td class=campo><%=rs("titulo")%></td>
	<td class=campo><%=rs("nome_curso")%></td>
	<td class=campo><%=rs("local_curso")%></td>
	<td class=campo><%=rs("ano_conclusao")%></td>
</tr>
<%
rs.movenext
loop
rs.close
%>
</table>

<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=grupo colspan=4>Experiência</td></tr>
<tr>
	<td class=titulo>Empresa</td>
	<td class=titulo>Período</td>
	<td class=titulo>Função/Salário</td>
	<td class=titulo>Atividades</td>
</tr>
<%
sql="SELECT h.cpf, h.nome_empresa, h.ramo_empresa, h.tel_empresa, admissao, demissao, funcao, salario, atividade_exercida " & _
"FROM tb_rh_hprofissional h where cpf='" & cpf & "' order by admissao"
rs.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
<tr>
	<td class=campo><%=rs("nome_empresa")%> (<%=rs("ramo_empresa")%>)</td>
	<td class=campo><%=rs("admissao")%> a <%=rs("demissao")%></td>
	<td class=campo><%=rs("funcao")%> / <%=rs("salario")%></td>
	<td class=campo><%=rs("atividade_exercida")%></td>
</tr>
<%
rs.movenext
loop
rs.close
%>
</table>

<table border="1" bordercolor=#000000 cellpadding="1" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td class=grupo colspan=4>Pretensão</td></tr>
<tr>
	<td class=titulo>Função</td>
	<td class=titulo>Salário</td>
	<td class=titulo>Habilidade</td>
</tr>
<%
sql="SELECT cpf, nome_funcao, salario, habilidade FROM tb_rh_pretensao p inner join tb_rh_funcao f on f.id_funcao=p.funcao " & _
"where cpf='" & cpf & "' order by id_pretensao " 
rs.Open sql, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
<tr>
	<td class=campo><%=rs("nome_funcao")%></td>
	<td class=campo><%=rs("salario")%></td>
	<td class=campo><%=rs("habilidade")%></td>
</tr>
<%
rs.movenext
loop
rs.close
%>
</table>





<%for a=1 to 4%>
<br>
<%next%>
<%

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>