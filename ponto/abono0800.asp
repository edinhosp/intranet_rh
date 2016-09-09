<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a46")="N" or session("a46")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Abono de intervalo - Atendimento</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<script language="VBScript">
	Sub informacao(texto)
		document.form.campo.value=texto
	End Sub
</script>

<%
dim conexao, rs
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form="" then
%>
<p class=titulo>Inclusão Abono dos 10 minutos - Atendimento
<form method="POST" action="abono0800.asp" name="form0">
<%
vezes=0
sql1="select f.chapa, f.nome, f.codsituacao, c.nome as funcao, a.descricao as horario " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.ahorario a, corporerm.dbo.pfuncao c " & _
"where f.codfuncao=c.codigo and f.codhorario=a.codigo and f.codsecao='03.1.021' and (f.codsituacao<>'D') "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=650>
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Sit</td>
	<td class=titulo>Função</td>
	<td class=titulo>Horário</td>
	<td class=titulo></td>
</tr>
<%
rs.movefirst:do while not rs.eof
classe="campo":classe2="campor"
%>
<tr>
	<td class=<%=classe%>><%=rs("chapa")%></td>
	<td class=<%=classe%>><%=rs("nome")%></td>
	<td class=<%=classe2%>><%=rs("codsituacao")%></td>
	<td class=<%=classe2%>><%=rs("funcao")%></td>
	<td class=<%=classe2%>><%=rs("horario")%></td>
	<td class=<%=classe%>>
		<input type="checkbox" name="emitir<%=vezes%>" value="ON" <%a="checked"%><%=b%> >
		<input type="hidden" name="id<%=vezes%>" value="<%=rs("chapa")%>">
	</td>
</tr>
<%
vezes=vezes+1
rs.movenext:loop
session("abono_inc")=vezes-1
end if
rs.close
%>
</table>
<table border="0" width="650" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo><input type="submit" value="Incluir intervalos" class="button" name="A1"></td>
	<td class=titulo><input type="submit" value="Incluir abonos" class="button" name="B1"></td>
	<td class=titulo><input type="submit" value="Apagar abonos" class="button" name="C1"></td>
</tr>
</table>

</form>
<%
end if 'request.form<>""
%>
<%

'*********************************************************************************************************************
if request.form("A1")<>"" then
	vez=session("abono_inc")

	sql="select iniciopermes, fimpermes from corporerm.dbo.aparam"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	inicio=rs("iniciopermes"):final=rs("fimpermes")
	rs.close
	response.write "<p><b>Efetuando inclusão de intervalos de 10 minutos no período de " & inicio & " a " & final & ".</b><br>"
	response.write "<form name='form'>"
	response.write "<input type='text' name='campo' size='80' class='form_ponto'>"
	response.write "</form>"

	for a=0 to vez
		chapa=request.form("id" & a)
		incluir=request.form("emitir" & a)
		if incluir="ON" then
			sql="select data, count(chapa) as total from corporerm.dbo.abatfun where chapa='" & chapa & "' and " & _
			"data between '" & dtaccess(inicio) & "' and '" & dtaccess(final) & "' group by data having count(chapa)=4 order by data"
			rs.Open sql, ,adOpenStatic, adLockReadOnly
			'response.write "<br>" & sql & "<br>" & rs.recordcount
			if rs.recordcount>0 then
				total=rs.recordcount:faltam=total-1
				rs.movefirst:do while not rs.eof
				if rs("total")=4 then
					sqlm="select j.batinicio, j.batfim from corporerm.dbo.ajorhor j, " & _
					"(select top 1 h.dtmudanca, h.codhorario, h.indiniciohor, a.descricao, a.databasehor, m.maxind, " & _
					"(round(convert(float,'" & dtaccess(rs("data")) & "'-a.databasehor+1),0)-(m.maxind-h.indiniciohor+1)) - (round((round(convert(float,'" & dtaccess(rs("data")) & "'-a.databasehor+1),0)-(m.maxind-h.indiniciohor))/m.maxind,0)*m.maxind) as indice_dia " & _
					"from corporerm.dbo.pfhsthor h, corporerm.dbo.ahorario a, (select codhorario, max(indice) as maxind from corporerm.dbo.abathor group by codhorario) m " & _
					"where h.chapa='" & chapa & "' and h.dtmudanca<='" & dtaccess(rs("data")) & "' and h.codhorario=a.codigo and m.codhorario=h.codhorario order by h.dtmudanca desc) I " & _
					"where j.codhorario=i.codhorario and (j.indinicio=i.indice_dia or j.indfim=i.indice_dia)"
					rs2.Open sqlm, ,adOpenStatic, adLockReadOnly
					if rs2.recordcount>0 then
						entrada=rs2("batinicio"):saida=rs2("batfim")
					end if 'rs2.recordcount
					rs2.close
					'response.write "<br>" & entrada & " " & saida
					randomize:int1ini=0:int1fim=0:int2ini=0:int2fim=0
					int1ini=entrada+110+int(rnd*20):'response.write "<br>" & int1ini
					sql2="select top 2 batida from corporerm.dbo.abatfun where chapa='" & chapa & "' and data='" & dtaccess(rs("data")) & "' order by batida asc"
					rs2.Open sql2, ,adOpenStatic, adLockReadOnly
					rs2.movelast
						m1=rs2("batida"):e1=(s1+10)
						if m1<int1ini then int1ini=m1-15
						int1fim=int1ini+10
						sql3="insert into corporerm.dbo.abatfun (codcoligada, chapa, data, batida, status, natureza) values (1,'" & chapa & "','" & dtaccess(rs("data")) & "'" & _
						",convert(bigint," & clng(int1ini+0) & "),'D',1)":response.write "<br>" & sql3:	
						conexao.execute sql3
						sql3="insert into corporerm.dbo.abatfun (codcoligada, chapa, data, batida, status, natureza) values (1,'" & chapa & "','" & dtaccess(rs("data")) & "'" & _
						",convert(bigint," & clng(int1fim+0) & "),'D',0)":'response.write "<br>" & sql3:	
						conexao.execute sql3
					rs2.close
					int2ini=entrada+280+int(rnd*20):'response.write "<br>" & int2ini
					sql2="select top 2 batida from corporerm.dbo.abatfun where chapa='" & chapa & "' and data='" & dtaccess(rs("data")) & "' order by batida desc"
					rs2.Open sql2, ,adOpenStatic, adLockReadOnly
					rs2.movelast
						m2=rs2("batida"):e2=s2+10
						if m2>int2ini then int2ini=m2+15
						int2fim=int2ini+10
						sql3="insert into corporerm.dbo.abatfun (codcoligada, chapa, data, batida, status, natureza) values (1,'" & chapa & "','" & dtaccess(rs("data")) & "'" & _
						",convert(bigint," & clng(int2ini+0) & "),'D',1)":'response.write "<br>" & sql3:	
						conexao.execute sql3
						sql3="insert into corporerm.dbo.abatfun (codcoligada, chapa, data, batida, status, natureza) values (1,'" & chapa & "','" & dtaccess(rs("data")) & "'" & _
						",convert(bigint," & clng(int2fim+0) & "),'D',0)":'response.write "<br>" & sql3:	
						conexao.execute sql3
					rs2.close
					response.write "<script>Document.form.campo.value=""Inserindo intervalos para: " & chapa & " no dia " & rs("data") & " - Faltam " & faltam & ".""</script>"
					'response.write "<br>" & "Inserindo intervalos para: " & chapa & " no dia " & rs("data") & " - Faltam " & faltam & "."
				end if 'total 8
				faltam=faltam-1
				rs.movenext:loop
			end if 'recordcount>0
			rs.close
		end if
	next
	response.write "<script>Document.form.campo.value=""Concluido!!!""</script>"
end if 'recordcount >formulario a1


'*********************************************************************************************************************
if request.form("B1")<>"" then
	vez=session("abono_inc")

	sql="select iniciopermes, fimpermes from corporerm.dbo.aparam"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	inicio=rs("iniciopermes"):final=rs("fimpermes")
	'inicio=dateserial(2012,4,1)
	'response.write inicio
	rs.close
	response.write "<p><b>Efetuando inclusão de abonos de 10 minutos no período de " & inicio & " a " & final & ".</b><br>"
	response.write "<form name='form'>"
	response.write "<input type='text' name='campo' size='80' class='form_ponto'>"
	response.write "</form>"

	for a=0 to vez
		chapa=request.form("id" & a)
		incluir=request.form("emitir" & a)
		if incluir="ON" then
			sql="select data, count(chapa) as total from corporerm.dbo.abatfun where chapa='" & chapa & "' and " & _
			"data between '" & dtaccess(inicio) & "' and '" & dtaccess(final) & "' group by data having count(chapa)=8 order by data"
			rs.Open sql, ,adOpenStatic, adLockReadOnly
			if rs.recordcount>0 then
				sql="delete from corporerm.dbo.aabonfun where chapa='" & chapa & "' and data between '" & dtaccess(inicio) & "' and '" & dtaccess(final) & "' and codabono='44' "
				conexao.execute sql
				response.write "<script>Document.form.campo.value=""Apagando abonos para: " & chapa & ".""</script>"
				'response.write "Apagando abonos para: " & chapa & "." & "<br>"
				total=rs.recordcount:faltam=total-1
				rs.movefirst:do while not rs.eof
				if rs("total")=8 then
					sql2="select top 2 batida from corporerm.dbo.abatfun where chapa='" & chapa & "' and data='" & dtaccess(rs("data")) & "' order by batida asc"
					rs2.Open sql2, ,adOpenStatic, adLockReadOnly
					rs2.movelast
						s1=rs2("batida"):e1=s1+10
						sql3="insert into corporerm.dbo.aabonfun (codcoligada, chapa, data, codabono, horainicio, horafim) " & _
						"values (1, '" & chapa & "', '" & dtaccess(rs("data")) & "', '44', " & s1 & ", " & e1 & " ) "
						response.write sql3
						conexao.execute sql3
					rs2.close
					sql2="select top 3 batida from corporerm.dbo.abatfun where chapa='" & chapa & "' and data='" & dtaccess(rs("data")) & "' order by batida desc"
					rs2.Open sql2, ,adOpenStatic, adLockReadOnly
					rs2.movelast
						s2=rs2("batida"):e2=s2+10
						sql3="insert into corporerm.dbo.aabonfun (codcoligada, chapa, data, codabono, horainicio, horafim, ) " & _
						"values (1,'" & chapa & "','" & dtaccess(rs("data")) & "','44'," & s2 & ", " & e2 & " ) "
						conexao.execute sql3
					rs2.close
					response.write "<script>Document.form.campo.value=""Inserindo abonos para: " & chapa & " no dia " & rs("data") & " - Faltam " & faltam &"""</script>"
					'response.write "Inserindo abonos para: " & chapa & " no dia " & rs("data") & " - Faltam " & faltam & "<br>"
				end if 'total 8
				faltam=faltam-1
				rs.movenext:loop
			end if 'recordcount>0
			rs.close
		end if
	next
	response.write "<script>Document.form.campo.value=""Concluido!!!""</script>"
end if 'recordcount >formulario B1


'*********************************************************************************************************************
if request.form("C1")<>"" then
	vez=session("abono_exc")

	sql="select iniciopermes, fimpermes from corporerm.dbo.aparam"
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	inicio=rs("iniciopermes"):final=rs("fimpermes")
	rs.close
	response.write "<p><b>Efetuando exclusão de abonos de 10 minutos no período de " & inicio & " a " & final & "</b>.<br>"
	response.write "<form name='form'>"
	response.write "<input type='text' name='campo' size='80' class='form_ponto'>"
	response.write "</form>"
	
	for a=0 to vez
		chapa=request.form("id" & a)
		incluir=request.form("emitir" & a)
		if incluir="ON" then
			sql="delete from corporerm.dbo.aabonfun where chapa='" & chapa & "' and data between '" & dtaccess(inicio) & "' and '" & dtaccess(final) & "' and codabono='44' "
			response.write sql
			conexao.execute sql
			response.write "<script>Document.form.campo.value=""Apagando abonos para: " & chapa & " no periodo de " & inicio & " a " & final & ".""</script>"
		end if
	next
	response.write "<script>Document.form.campo.value=""Concluido!!!""</script>"
end if 'recordcount >formulario C1

%>

<!-- <div align="right"> -->
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650" height=990>
</table>
<!-- </div> -->
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>