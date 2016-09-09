<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a82")="N" or session("a82")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Relatório de Assistência Médica</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, rt(15), rd(15)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
sessao=session.sessionid:sessao=session("usuariomaster")

if request.form<>"" then
	mes=request.form("mes")
	ano=request.form("ano")

end if

if request.form="" then 
%>
<p class=titulo>Geração de relatório de conferência da Assistência Médica
<form method="POST" action="relacao.asp">
<p>Data base para o relatório: <input type="text" name="data" size="8" class=a value="<%=int(now())%>">
Dias a retroagir/avançar: <input type="text" name="fator" size="2" class=a value="<%=30%>">
<br>
Ordem: <select size="1" name="ordem">
	<option value="chapa">Chapa</option>
	<option value="principal">Nome</option>
</select><br>
</p>		
<p><input type="submit" value="Visualizar relatório" name="Gerar" class="button"></p>
</form>
<%
else
data=request.form("data"):data=dtaccess(data)
fator=request.form("fator")
ordem=request.form("ordem")

%>
<table border="0" cellpadding="2" width="990" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td align="left"  >Conferência de Assistência Médica</td>
	<td align="center">Relação de Beneficiários e Dependentes</td>
	<td align="right" ><%=cdate(request.form("data"))%></td>
</tr>
</table>
<table border="0" cellpadding="1" width="990" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulor>Chapa </td>
	<td class=titulor>Nome  </td>
	<td class=titulor>Nasc. </td>
	<td class=titulor>Sexo  </td>
	<td class=titulor colspan=6>Função/Seção</td>
	<td class=titulor>Datas </td>
	<td class=titulor>...</td>
</tr>
<%
sql="select chapa, f.nome, sit=CODSITUACAO, admissao=DATAADMISSAO, demissao=DATADEMISSAO, secao=s.DESCRICAO, funcao=c.NOME, " & _
"p.SEXO, p.DTNASCIMENTO, f.CODSECAO, f.CODFUNCAO " & _
"from corporerm.dbo.PFUNC f " & _
"inner join corporerm.dbo.PSECAO s on s.CODIGO=f.CODSECAO " & _
"inner join corporerm.dbo.PFUNCAO c on c.CODIGO=f.CODFUNCAO " & _
"inner join corporerm.dbo.PPESSOA p on p.CODIGO=f.CODPESSOA " & _
"where (DATAADMISSAO between DATEADD(D,-" & fator & ",'" & data & "') and DATEADD(D," & fator & ",'" & data & "') " & _
"or DATADEMISSAO between DATEADD(D,-" & fator & ",'" & data & "') and DATEADD(D," & fator & ",'" & data & "')) " & _
"and CODTIPO='N' " & _
"order by CODSITUACAO, DATAADMISSAO, DATADEMISSAO "
linha=2:limite=45 '72
rs.Open sql, ,adOpenStatic, adLockReadOnly
inicio=1
rs.movefirst
do while not rs.eof
if resumo=0 then
	if linha>limite then
		pagina=pagina+1
		response.write "</table>"
		response.write "<p style='margin-top:0;margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
		response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
		response.write "<table border='0' cellpadding='2' width='990' cellspacing='0' style='border-collapse: collapse'>"
		response.write "<tr>"
		response.write "<td align='left'  >Conferência de Assistência Médica</td>"
		response.write "<td align='center'>Relação de Beneficiários e Dependentes</td>"
		response.write "<td align='right' >Empresa: ...</td>"
		response.write "</tr>"
		response.write "</table>"
		response.write "<table border='0' cellpadding='0' width='990' cellspacing='0' style='border-collapse: collapse'>"
		response.write "<tr>"
		response.write "<td class=titulor>Chapa </td>"
		response.write "<td class=titulor>Nome  </td>"
		response.write "<td class=titulor>Nasc. </td>"
		response.write "<td class=titulor>Sexo  </td>"
		response.write "<td class=titulor colspan=6>Função/Seção</td>"
		response.write "<td class=titulor>Datas </td>"
		response.write "<td class=titulor>...</td>"
		response.write "</tr>"
		linha=2
	end if
end if 'resumo
if ultchapa<>rs("chapa") then estilo="style='border-top: 1px solid #000000'" else estilo=""
texto1="Entr: " & rs("admissao")
if rs("demissao")<>"" then texto1=texto1 & " - Saida: " & rs("demissao")

if resumo=0 then
%>
<tr>
	<td class="campor" <%=estilo%>><%=rs("chapa")%></td><%linha=linha+1%>
	<td class="campor" <%=estilo%>><%=rs("nome")%></td>
	<td class="campor" align="center" <%=estilo%>><%=rs("dtnascimento")%></td>
	<td class="campor" align="center" <%=estilo%>><%=rs("sexo")%></td>
	<td class="campor" colspan=6 <%=estilo%>><%=rs("funcao")%> / <%=replace(replace(rs("secao"),"CURSO DE ",""),"TECNOLOGIA","TEC.")%></td>
	<td class="campor" <%=estilo%>><%=texto1%></td>
	<td class="campor" <%=estilo%>> </td>
</tr>
<%
	sqlt="select empresa, operadora, m.plano, m.codigo, inclusao, fvigencia " & _
	"from assmed_mudanca m inner join assmed_planos p on p.codigo=m.empresa and m.plano=p.plano " & _
	"inner join assmed_empresa e on e.codigo=m.empresa " & _
	"where chapa='" & rs("chapa")& "' and p.tipo='M' and fvigencia>='" & data & "'"
	rs3.Open sqlt, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then
		if rs("sit")="D" then str4="Ex: (" & rs3("fvigencia") & ")" else str4=""
		str1="AM: " & rs3("operadora")
		str2=rs3("plano") & " (" & rs3("inclusao") & ")"
		str3=rs3("codigo")
	else
		str1="AM: -":str2="":str3="":str4=""
	end if
	rs3.close
	sqlt="select empresa, operadora, m.plano, m.codigo, inclusao, fvigencia " & _
	"from assmed_mudanca m inner join assmed_planos p on p.codigo=m.empresa and m.plano=p.plano " & _
	"inner join assmed_empresa e on e.codigo=m.empresa " & _
	"where chapa='" & rs("chapa")& "' and p.tipo='O' and fvigencia>='" & data & "'"
	rs3.Open sqlt, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then
		if rs("sit")="D" then str8="Ex: (" & rs3("fvigencia") & ")" else str8=""
		str5="AO: " & rs3("operadora")
		str6=rs3("plano") & " (" & rs3("inclusao") & ")"
		str7=rs3("codigo")
	else
		str5="AO: -":str6="":str7="":str8=""
	end if
	rs3.close
%>	
<tr>
	<td class="campor" colspan=7></td><%linha=linha+1%>
	<td class="campor"><i><%=str1%></i></td>
	<td class="campor"><i><%=str2%></i></td>
	<td class="campor"><i><%=str3%></i></td>
	<td class="campor"><i><%=str4%></i></td>
	<td class="campor"></td>
</tr>
<tr>
	<td class="campor" colspan=7></td><%linha=linha+1%>
	<td class="campor"><i><%=str5%></i></td>
	<td class="campor"><i><%=str6%></i></td>
	<td class="campor"><i><%=str7%></i></td>
	<td class="campor"><i><%=str8%></i></td>
	<td class="campor"></td>
</tr>
<%
'----------------- dependentes
sqld="select CHAPA, NRODEPEND, NOME, DTNASCIMENTO, SEXO, ESTADOCIVIL, GRAUPARENTESCO, parentesco=p.DESCRICAO, cpf " & _
"from corporerm.dbo.PFDEPEND d " & _
"inner join corporerm.dbo.PCODPARENT p on p.CODCLIENTE=d.GRAUPARENTESCO " & _
"where CHAPA='" & rs("chapa") & "' and GRAUPARENTESCO not in ('6','7') "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
	do while not rs2.eof
	idade=datediff("yyyy",rs2("dtnascimento"),now())
%>
<tr>
	<td class="campor"></td><%linha=linha+1%>
	<td class="campor"><i><%=rs2("nome")%></i></td>
		<%if rs2("dtnascimento")="" or isnull(rs2("dtnascimento")) then estilo="fundor" else estilo="campor"%>
	<td class=<%=estilo%> align="center"><i><%=rs2("dtnascimento")%></i></td>
		<%if rs2("sexo")="" or isnull(rs2("sexo")) then estilo="fundor" else estilo="campor"%>
	<td class=<%=estilo%> align="center"><i><%=rs2("sexo")%></i></td>
		<%if rs2("parentesco")="" or isnull(rs2("parentesco")) then estilo="fundor" else estilo="campor"%>
	<td class=<%=estilo%> ><i><%=rs2("parentesco")%></i></td>
		<%if rs2("estadocivil")="" or isnull(rs2("estadocivil")) then estilo="fundor" else estilo="campor"%>
	<td class=<%=estilo%> ><i><%=rs2("estadocivil")%></td>
		<%if idade="" or isnull(idade) then estilo="fundor" else estilo="campor"%>
	<td class=<%=estilo%> ><i><%=idade%></td>
<%
	estilo="campor"
	sqlp="select empresa, operadora, m.plano, m.codigo, inclusao, fvigencia " & _
	"from assmed_dep_mudanca m inner join assmed_planos p on p.codigo=m.empresa and m.plano=p.plano " & _
	"inner join assmed_empresa e on e.codigo=m.empresa " & _
	"where chapa='" & rs("chapa")& "' and nrodepend=" & rs2("nrodepend") & " and p.tipo='M' and fvigencia>='" & data & "'"
	rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then
		if rs("sit")="D" then str4="Ex: (" & rs3("fvigencia") & ")" else str4=""
		str1="AM: " & rs3("operadora")
		str2=rs3("plano") & " (" & rs3("inclusao") & ")"
		str3=rs3("codigo")
	else
		str1="AM: -":str2="":str3="":str4=""
	end if
	rs3.close
	sqlp="select empresa, operadora, m.plano, m.codigo, inclusao, fvigencia " & _
	"from assmed_dep_mudanca m inner join assmed_planos p on p.codigo=m.empresa and m.plano=p.plano " & _
	"inner join assmed_empresa e on e.codigo=m.empresa " & _
	"where chapa='" & rs("chapa")& "' and nrodepend=" & rs2("nrodepend") & " and p.tipo='O' and fvigencia>='" & data & "'"
	rs3.Open sqlp, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then
		if rs("sit")="D" then str8="Ex: (" & rs3("fvigencia") & ")" else str8=""
		str5="AO: " & rs3("operadora")
		str6=rs3("plano") & " (" & rs3("inclusao") & ")"
		str7=rs3("codigo")
	else
		str5="AO: -":str6="":str7="":str8=""
	end if
	rs3.close
	if (rs2("cpf")="" or isnull(rs2("cpf"))) and rs("sit")<>"D" and idade>=18 and (str1<>"AM: -" or str5<>"AO: -") then str4="CPF??":
%>	
	<td class="campor"><i><%=str1%></i></td>
	<td class="campor"><i><%=str2%></i></td>
	<td class="campor"><i><%=str3%></i></td>
	<td class="campor"><i><%=str4%></i></td>
	<td class="campor"></td>
</tr>
<tr>
	<td class="campor" colspan=7></td><%linha=linha+1%>
	<td class="campor"><i><%=str5%></i></td>
	<td class="campor"><i><%=str6%></i></td>
	<td class="campor"><i><%=str7%></i></td>
	<td class="campor"><i><%=str8%></i></td>
	<td class="campor"></td>
</tr>

<%
	rs2.movenext
	loop
end if
rs2.close
'-----------------------------
ultchapa=rs("chapa")
end if 'resumo
rs.movenext
loop
rs.close

%>
</table>
<%

if linha>limite-4 then
	pagina=pagina+1
	'response.write "<br>"
	response.write "<p style='margin-top:0;margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<table border='0' cellpadding='2' width='990' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  >Controle de Assistência Médica        </td>"
	response.write "<td align='center'>Relação de Beneficiários e Dependentes</td>"
	response.write "<td align='right' >Empresa: " & operadora & "</td>"
	response.write "</tr>"
	response.write "</table>"
	linha=1
end if
%>

<%
linha=linha+1
pagina=pagina+1
response.write "<br>"
response.write "<p style='margin-top:0;margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"

end if
%>
</body>
</html>
<%
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%>