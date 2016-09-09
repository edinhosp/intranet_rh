<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")="N" or session("a48")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Conferência de Arquivo Retorno-Bradesco</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
inicio=now()
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
sessao=session.sessionid
%>
<p class=titulo>Conferência de arquivo Retorno-Bradesco</p>
<%
if request.form("Gerar")="" then
%>
<form method="POST" action="arqfopagretorno.asp" name="form">
<p style="margin-top: 0; margin-bottom: 0">
<textarea rows="15" class="p" name="texto" cols="140"><%=request.form("texto")%></textarea>
</p>
<p><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></p>
</form>
<%
end if '******************
%>
<%
if request.form("Gerar")<>"" then
	sql="delete from fopag_retorno where sessao='" & sessao & "'":conexao.execute sql
	posicao=1
	variavel=request.form("texto")
	tamanho=len(variavel)
	for a=1 to tamanho
		letra=mid(variavel, a, 1)
		if asc(letra)<>13 and asc(letra)<>10 then stringtxt=stringtxt & letra else stringtxt=stringtxt
		posicao=posicao+1
		if asc(letra)=13 then
			registro=stringtxt
			sql="insert into fopag_retorno (sessao, registro, ordem) values ('" & sessao & "', '" & registro & "'," & left(registro,1) & ")"
			conexao.execute sql
			stringtxt=""
			posicao=1
		end if
	next
	
	mens=""
	sql="select ordem, registro from fopag_retorno where sessao='" & sessao & "'"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
	do while not rs.eof
		if rs("ordem")=0 then
			err_agemp=mid(rs("registro"),121,1)
			err_ccemp=mid(rs("registro"),122,1)
			err_razao=mid(rs("registro"),123,1)
			err_dtdebito=mid(rs("registro"),124,1)
			if err_agemp="S" then mens=mens & "Erro na agência da Empresa<br>"
			if err_ccemp="S" then mens=mens & "Erro na conta da Empresa<br>"
			if err_razao="S" then mens=mens & "Erro na razão da Empresa<br>"
			if err_dtdebito="S" then mens=mens & "Erro na data de Débito<br>"
		end if
		if rs("ordem")=1 then
			funcionario=mid(rs("registro"),83,38) & " (" & mid(rs("registro"),127,13)/100 & ")"
			err_agfun=mid(rs("registro"),151,1)
			err_ccfun=mid(rs("registro"),152,1)
			if err_agfun="S" then mens=mens & "Erro na agência do funcionário " & funcionario & "<br>"
			if err_ccfun="S" then mens=mens & "Erro na conta corrente do funcionário " & funcionario & "<br>"
		end if
		if rs("ordem")=9 then
			err_total=mid(rs("registro"),15,1)
			if err_total="S" then mens=mens & "Erro no valor total do débito<br>"
		end if
	
	rs.movenext
	loop
	rs.close
	end if 'recordcount>0
	if mens="" then
		response.write "<p><font color=blue><b>O arquivo de Retorno não apresentou erros.</b></font>"
	else
		response.write "<p><font color=red><b>O arquivo de Retorno apresentou os seguintes erros:</b></font><br>"
		response.write mens	
	end if
%>

<%
end if 'request.form 
%> 

</body>
</html>
<%
set rs=nothing
set rs1=nothing
conexao.close
set conexao=nothing
%>
