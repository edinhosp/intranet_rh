<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Pesquisa Estacionamento</title>
<link rel="stylesheet" type="text/css" href="../diversos.css">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
%>
<body>
<%
if request.form("cmdLogin")<>"" then
	usuario1=ucase(request.form("txtUsuario"))
	senha1=request.form("txtpassword")
	usuario1=replace( request.form("txtUsuario") ,"--","")
	senha1=replace( request.form("txtpassword") ,"--","")
	sql="SELECT * From usuarios where usuario='" & usuario1 & "' and ativo=1"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 
		usuariof=rs("usuario"):senha=rs("password")
		Session("DescricaoErro")=""
		'testar se confere senha e usuario
		sqlrm="select chapa, codsituacao from corporerm.dbo.pfunc where chapa='" & usuario1 & "'"
		rs2.Open sqlrm, ,adOpenStatic, adLockReadOnly
			situ=rs2("codsituacao")
			if situ="A" or situ="E" or situ="F" or situ="Z" then permitido=1 else permitido=0
		rs2.close
		if usuariof=usuario1 and senha=senha1 and permitido=1 then
			Session("acesso")=1
			session("grant_estacionamento")=usuariof
			rp=request.cookies("vrh06")("registropagina")
			if rp<>"" then Session("RegistrosPorPagina")=rp else Session("registrosporpagina")=20
			cb=request.cookies("vrh06")("cabecalho")
			if cb="sim" then Session("cabecalho")="sim" else Session("cabecalho")="nao"
			Session("Usuarioname") =rs("nome")
			Session("usuariogrupo")=rs("grupo")
			Session("estilo")      =rs("estilo")
			if rs("master")=true then emaster=1 else emaster=0
			Session("master")      =emaster
			SendIp=request.servervariables("LOCAL_ADDR")
			SendIp=request.servervariables("REMOTE_ADDR")
			Session("UsuarioMaster")=ucase(usuario1)
			if Session("usuariomaster")<>"" then
			sqlz="INSERT INTO login ( usuario, entrada, sessao, ipcomp ) SELECT '" & usuario1 & "' AS Expr1, getdate() AS Expr2," & _
			" '" & Session.Sessionid & "' AS Expr3, '" & sendip & "';"
			conexao.Execute sqlz
			end if
			temp=rs("timeout"): if isnumeric(temp) then Session.timeout=temp else Session.timeout=20
			rs.close
			sqlu="select top 2 entrada from login where usuario='" & usuario1 & "' order by entrada desc "
			rs.Open sqlu, ,adOpenStatic, adLockReadOnly
			if rs.recordcount>1 then
				rs.movenext
				Session("lastacesso")=rs("entrada")
			else
				Session("lastacesso")="-.-"
			end if
			rs.close
			Session("DescricaoErro")=""
		else
			Session("DescricaoErro")="Usuario não cadastrado/Senha não confere"
		end if
	else 'rs.recordcount=0
		'teste professor
		rs.close
		sql="select chapa, nome, apelido, cartidentidade, cartmodelo19, codsituacao from qry_funcionarios where codsecao in ('03.2.008','01.2.008') and codsituacao<>'D' and chapa='" & usuario1 & "'"
		rs.Open sql, ,adOpenStatic, adLockReadOnly
		'response.write rs.recordcount
		if rs.recordcount>0 then
			if (rs("cartidentidade")="" or isnull(rs("cartidentidade"))) and rs("cartmodelo19")<>"" then ident=rs("cartmodelo19") else ident=rs("cartidentidade")
			senhaf=left(textopuro(ident,3),4)
			usuariof=rs("chapa"):senha=senhaf
			'response.write senhaf
			'response.write usuariof
			Session("DescricaoErro")=""
'-----------------------------------------------------
			if usuariof=usuario1 and senha=senha1 then
				Session("acesso")=2
				session("grant_estacionamento")=usuariof
				rp=request.cookies("vrh06")("registropagina")
				if rp<>"" then Session("RegistrosPorPagina")=rp else Session("registrosporpagina")=20
				cb=request.cookies("vrh06")("cabecalho")
				if cb="sim" then Session("cabecalho")="sim" else Session("cabecalho")="nao"
				Session("Usuarioname") =rs("apelido")
				Session("usuariogrupo")="INSPETORIA"
				'Session("grant_menu")  ="100"
				Session("a100")="T"
				Session("estilo")      ="diversos.css"
				SendIp=request.servervariables("LOCAL_ADDR")
				SendIp=request.servervariables("REMOTE_ADDR")
				Session("UsuarioMaster")=ucase(usuario1)
				if Session("usuariomaster")<>"" then
				sqlz="INSERT INTO login ( usuario, entrada, sessao, ipcomp ) SELECT '" & usuario1 & "' AS Expr1, getdate() AS Expr2," & _
				" '" & Session.Sessionid & "' AS Expr3, '" & sendip & "';"
				conexao.Execute sqlz
				end if
				temp=20: if isnumeric(temp) then Session.timeout=temp else Session.timeout=20
				rs.close
				sqlu="select top 2 entrada from login where usuario='" & usuario1 & "' order by entrada desc "
				rs.Open sqlu, ,adOpenStatic, adLockReadOnly
				if rs.recordcount>1 then
					rs.movenext
					Session("lastacesso")=rs("entrada")
				else
					Session("lastacesso")="-.-"
				end if
				rs.close
				Session("DescricaoErro")=""
			else
				Session("DescricaoErro")="Senha não confere"
			end if
'-----------------------------------------------------
		end if 'recordcount>0
	end if

	'conexao.close
end if

'if session("usuariomaster")="" or session("grant_estacionamento")="" then
if session("grant_estacionamento")="" then
tam=290
%>
<div align="center">
<form action="index.asp" method="post" name="formlogin">
<table border="0" cellpadding="0" cellspacing="0" style="background-color:transparent;border-collapse: collapse;background:transparent url(../images/acessorh.gif) no-repeat center;" width="620" height="350">
<tr><td colspan=2 height="220" style="background-color:transparent"></td></tr>
<tr><td width="<%=tam%>" height="39" style="background-color:transparent"></td>
	<td style="background-color:transparent" valign="top">
	<input type="text" name="txtUsuario" value="<%=Session("UsuarioMaster")%>" style="font-family:Tahoma; font-size:8pt; color:black; border:0px transparent; border-bottom:1px #000000 solid;background-color:white; " size="6" maxlength="6">
	</td>
</tr>
<tr><td width="<%=tam%>" height="40" style="background-color:transparent"></td>
	<td style="background-color:transparent" valign="top">
	<input type="password" name="txtPassword" value="" style="font-family:Tahoma; font-size:8pt; color:black; border:0px transparent; border-bottom:1px #000000 solid;background-color:white; " size="8" maxlength="8">
	</td>
</tr>
<tr><td width="<%=tam%>" style="background-color:transparent"></td>
	<td style="background-color:transparent" align="left">
	&nbsp;&nbsp;
	<input type="submit" class=button value=" Entrar " name="cmdLogin">
	</td></tr>
</table>	
</form>
</div>

<%
end if
'session("grant_estacionamento")=""
%>

<%
'if session("usuariomaster")<>"" or session("grant_estacionamento")<>"" then
if session("grant_estacionamento")<>"" then
%>
<form name="form" action="index.asp" method="post">
<table border="0" bordercolor="#000000" cellpadding="5" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=grupo colspan=5>Seleção para Pesquisa</td>
</tr>
<tr>
	<td>
		<input type="radio" name="tipo" value="p" <%if request.form("tipo")="p" then response.write "checked"%> onclick="tipo2('p')"> placa<br>
		<input type="radio" name="tipo" value="c" <%if request.form("tipo")="c" then response.write "checked"%> onclick="tipo2('c')"> cracha
	</td>
	<td>
		<input type="text" name="placa1" size="3" maxlength="3" value="<%=request.form("placa1")%>">-
		<input type="text" name="placa2" size="4" maxlength="4" value="<%=request.form("placa2")%>"><br>
		<input type="text" name="cracha" size="8" maxlength="8" value="<%=request.form("cracha")%>">
	</td>
	<td>
		<input type="submit" value="Pesquisar" name="pesquisar">
	</td>	
</tr>
</table>

</form>
<hr>
<%
tipo=request.form("tipo")
'**********************************************************************************************
if tipo="c" then
	valor=request.form("cracha")
	if isnumeric(valor)=true then
		valor=numzero(valor,5)
	end if
	sqlc="select f.chapa, f.nome, f.codsecao, s.descricao, f.codsindicato, f.codsituacao, " & _
	"situacao=case when f.codsituacao in ('A','F','Z') then 'ATIVO' else case when f.codsituacao='D' then 'DEMITIDO' else 'AFASTADO' end end " & _
	"from corporerm.dbo.pfunc f, corporerm.dbo.psecao s where s.codigo=f.codsecao and f.chapa='" & valor & "' "
	rs.Open sqlc, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse">	
<tr>
	<td class=titulo rowspan=2>Nome</td>
	<td class=titulo rowspan=2>Setor</td>
	<td class=titulo rowspan=2>Situação</td>
	<td class=titulo colspan=4 align="center">Estacionamentos</td>
</tr>
<tr>
	<td class=titulo>Narciso  </td><td class=titulo>Jd.Wilson</td><td class=titulo>Coral    </td><td class=titulo>B.Park   </td>
</tr>
<tr>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo><%=rs("descricao")%>
	<td class=campo><%=rs("situacao")%>
<%
	if rs("codsindicato")="03" and rs("codsituacao")<>"D" then
		sqlbloco="select bloco from blocos where codsecao='" & rs("codsecao") & "' "
		rs2.Open sqlbloco, ,adOpenStatic, adLockReadOnly
		bloco=" (<font color='blue'>" & rs2("bloco") & "</font>)"
		rs2.close
	else
		bloco=""
	end if
%>
	<%=bloco%></td>
<%
	sqle="select vy, ns, jw, bp from veiculos_a where chapa='" & rs("chapa")& "' and getdate() between inicio and termino "
	rs2.Open sqle, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then
	response.write rs2("vy")
%>
	<td class=campo align="center"><%if rs2("ns")=True then response.write "<img src='../images/truck.gif' width=13 border='0'>" %></td>
	<td class=campo align="center"><%if rs2("jw")=True then response.write "<img src='../images/truck.gif' width=13 border='0'>" %></td>
	<td class=campo align="center"><%if rs2("vy")=True then response.write "<img src='../images/truck.gif' width=13 border='0'>" %></td>
	<td class=campo align="center"><%if rs2("bp")=True then response.write "<img src='../images/truck.gif' width=13 border='0'>" %></td>
<%
	else
		response.write "<td class=grupo colspan=4 align=""center"">Não estaciona!!!</td>"
	end if
	rs2.close
%>
</tr>
</table>

<br>

<table border="1" bordercolor="Green" cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=8>Veículos Cadastrados</td>
</tr>
<tr>
	<td class=titulor>Marca</td>
	<td class=titulor>Modelo</td>
	<td class=titulor>Ano</td>
	<td class=titulor>Cor</td>
	<td class=titulor>Placa</td>
	<td class=titulor>Cancelado</td>
</tr>
<%
sqlv="select id_veiculo, marca,modelo, ano, cor, placa, dtcadastro, dttermino " & _
"from veiculos where chapa='" & rs("chapa") & "' and dttermino is null order by dttermino desc, dtcadastro "
rs2.Open sqlv, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
if isnull(rs2("dttermino")) then campo="campo" else campo="fundo"
%>
<tr>
	<td class=<%=campo%> ><%=rs2("marca")%></td>
	<td class=<%=campo%> ><%=rs2("modelo")%></td>
	<td class=<%=campo%> ><%=rs2("ano")%></td>
	<td class=<%=campo%> ><%=rs2("cor")%></td>
	<td class=<%=campo%>  nowrap><%=rs2("placa")%></td>
	<td class=<%=campo%> ><%=rs2("dttermino")%></td>
</tr>
<%
rs2.movenext:loop
end if
rs2.close
%>
</table>

<%
	else
		response.write "<p class=titulo>Não existe funcionário/professor com este código."
	end if 'rs.recordcount>0
	rs.close
end if 'tipo c


'**********************************************************************************************
if tipo="p" then
	valor1=request.form("placa1")
	valor2=request.form("placa2")
	if valor1<>"" and valor2="" then placa=valor1 & "%"
	if valor2<>"" and valor1="" then placa="%" & valor2
	if valor1<>"" and valor2<>"" then placa=valor1 & "%" & valor2
	sqlc="select f.chapa, f.nome, f.codsecao, s.descricao, f.codsindicato, f.codsituacao " & _
	"from pfunc f, psecao s where s.codigo=f.codsecao and f.chapa='" & valor & "' "
	sqlp="select v.placa, v.chapa, v.marca, v.modelo, v.cor, v.dtcadastro, v.dttermino, f.nome, f.codsituacao, s.descricao " & _
	", situacao=case when f.codsituacao in ('A','F','Z') then 'ATIVO' else case when f.codsituacao='D' then 'DEMITIDO' else 'AFASTADO' end end " & _
	"from veiculos v, corporerm.dbo.pfunc f, corporerm.dbo.psecao s " & _
	"where f.codsecao=s.codigo and f.chapa collate database_default=v.chapa and v.placa like '" & placa & "' order by v.placa "
	
	rs.Open sqlp, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse">	
<tr><td class=titulo rowspan=2>Placa</td>
	<td class=titulo rowspan=2>Modelo</td>
	<td class=titulo rowspan=2>Cor</td>
	<td class=titulo rowspan=2>Cancelado</td>
	<td class=titulo rowspan=2>Motorista</td>
	<td class=titulo rowspan=2>Situação</td>
	<td class=titulo colspan=4 align="center">Estacionamento</td>
</tr>
<tr>
	<td class=titulo>Narciso  </td><td class=titulo>Jd.Wilson</td><td class=titulo>Coral    </td><td class=titulo>B.Park   </td>
</tr>
<%
	do while not rs.eof
	if rs("codsituacao")="D" then classe="campov" else classe="campo"
%>
<tr>
	<td class=campo rowspan=2 align="center"><%=rs("placa")%></td>
	<td class=campo rowspan=2><%=rs("modelo")%></td>
	<td class=campo rowspan=2><%=rs("cor")%></td>
	<td class=campo rowspan=2 align="right"><%=rs("dttermino")%></td>
	<td class=<%=classe%>><%=rs("nome")%></td>
	<td class=<%=classe%>><%=rs("situacao")%></td>
<%
	sqle="select vy, ns, jw, bp from veiculos_a where chapa='" & rs("chapa")& "' and getdate() between inicio and termino "
	rs2.Open sqle, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then
%>	
	<td class=campo align="center"><%if rs2("ns")=True then response.write "<img src='../images/truck.gif' width=13 border='0'>" %></td>
	<td class=campo align="center"><%if rs2("jw")=True then response.write "<img src='../images/truck.gif' width=13 border='0'>" %></td>
	<td class=campo align="center"><%if rs2("vy")=True then response.write "<img src='../images/truck.gif' width=13 border='0'>" %></td>
	<td class=campo align="center"><%if rs2("bp")=True then response.write "<img src='../images/truck.gif' width=13 border='0'>" %></td>
<%
	else
		response.write "<td class=grupo colspan=4 align=""center"">Não estaciona!!!</td>"
	end if
	rs2.close
%>
</tr>
<tr>
	<td class=<%=classe%> colspan=2><%=rs("descricao")%></td>
</tr>
<%
	rs.movenext:loop
%>

<%
	else
		response.write "<p class=titulo>Não existem veículos com esta placa."
	end if 'rs.recordcount>0
	rs.close
end if 'tipo p
%>

<%

end if 'session com valor

set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>

</body>
</html>

<script language="VBScript" text="text/vbscript">
	Sub tipo2(tipo)
		temp=tipo
		if temp="p" then
			//document.form.tipo.value="p"
			document.form.cracha.value=""
			document.form.cracha.style.background="#cccccc"
			document.form.placa1.style.background="#ffffff"
			document.form.placa2.style.background="#ffffff"
			document.form.placa1.disabled=false
			document.form.placa2.disabled=false
			document.form.cracha.disabled=true
			document.form.pesquisar.disabled=false
			document.form.tipo(0).checked=true
		else 'cracha
			//document.form.tipo.value="c"
			document.form.placa1.value=""
			document.form.placa2.value=""
			document.form.cracha.style.background="#ffffff"
			document.form.placa1.style.background="#cccccc"
			document.form.placa2.style.background="#cccccc"
			document.form.placa1.disabled=true
			document.form.placa2.disabled=true
			document.form.cracha.disabled=false
			document.form.pesquisar.disabled=false
			document.form.tipo(1).checked=true
		end if
		//msgbox document.form.tipo(0).checked & document.form.tipo(1).checked 
	End Sub
	Sub cracha_onClick
		document.form.tipo(1).checked=true
		document.form.cracha.value=""
	End Sub
	Sub placa1_onClick
		document.form.tipo(0).checked=true
		document.form.placa1.value=""
	End Sub
	Sub placa2_onClick
		document.form.tipo(0).checked=true
		document.form.placa2.value=""
	End Sub
</script>