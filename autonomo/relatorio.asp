<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a63")="N" or session("a63")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Relatório de Pagamentos a Autônomos</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="VBScript">
	Sub dtemi1_onChange
		dataini=formatdatetime(document.form.dtemi1.value,2)
		datafim=dateserial(year(dataini),month(dataini)+1,1)-1
		document.form.dtemi1.value=dataini
		document.form.dtemi2.value=datafim
		document.form.dtpag1.value=""
		document.form.dtpag2.value=""
		document.form.tipo_emissao.option="emi"
	end Sub

	Sub dtemi2_onChange
		document.form.dtemi2.value=formatdatetime(document.form.dtemi2.value,2)
	end Sub

	Sub dtpag1_onChange
		dataini=formatdatetime(document.form.dtpag1.value,2)
		datafim=dateserial(year(dataini),month(dataini)+1,1)-1
		document.form.dtpag1.value=dataini
		document.form.dtpag2.value=datafim
		document.form.dtemi1.value=""
		document.form.dtemi2.value=""
		document.form.tipo_emissao.option.value="pag"
	end Sub
	Sub dtpag2_onChange
		document.form.dtpag2.value=formatdatetime(document.form.dtpag2.value,2)
	end Sub
	
</script>

</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, rt(10), rd(10)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	dtpag1=request.form("dtpag1")
	dtpag2=request.form("dtpag2")
	dtemi1=request.form("dtemi1")
	dtemi2=request.form("dtemi2")
	tipo_emissao=request.form("tipo_emissao")
	sessao=session.sessionid

sql1 = "SELECT a.nome_autonomo, a.cpf, a.nit, p.data_emissao, p.data_pagamento, " & _
"p.servico_prestado, p.outros_rendimentos, p.desconto_ir, p.desconto_inss, p.desconto_iss, " & _
"p.outros_descontos, p.valor_liquido, p.inss_outra_empresa " & _
"FROM autonomo a INNER JOIN autonomo_rpa p ON a.id_autonomo = p.id_autonomo "

select case tipo_emissao
	case "pag"
		data1=dtpag1:data2=dtpag2
		sql2="WHERE p.data_pagamento Between '" & dtaccess(data1) & "' And '" & dtaccess(data2) & "' "
		titulo="Data de Pagamento entre " & data1 & " e " & data2 & " "
	case "emi"
		data1=dtemi1:data2=dtemi2
		sql2="WHERE p.data_emissao Between '" & dtaccess(data1) & "' And '" & dtaccess(data2) & "' "
		titulo="Data de Emissão entre " & data1 & " e " & data2 & " "
	case "nul"
		sql2="WHERE p.data_pagamento is null "
		titulo="Sem Data de Pagamento"
end select
if sql2="" then sql3="WHERE a.id_autonomo>0 "
if request.form("descricao")<>"" then sql3=sql3 & " and descricao_servico like '%" & request.form("descricao") & "%' " else sql3=""

sql=sql1 & sql2	& sql3 &  "order by a.nome_autonomo, p.data_pagamento "
'response.write "<br>" & sql
end if

if request.form="" then
%>
<p class=titulo>Geração de relatório de pagamento a autônomos
<form method="POST" action="relatorio.asp" name="form">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td colspan="3" class=grupo>Selecionar tipo de emissão</td>
</tr>
<tr>
	<td class=titulo><input type="radio" name="tipo_emissao" value="pag" checked>Por data de pagamento&nbsp;</td>
	<td class=titulo>&nbsp;&nbsp;de <input type="text" name="dtpag1" size="8" value=""></td>
	<td class=titulo>a <input type="text" name="dtpag2" size="8" value="">&nbsp;</td>
</tr>
<tr>
	<td class=titulo><input type="radio" name="tipo_emissao" value="emi">Por data de emissão&nbsp;</td>
	<td class=titulo>&nbsp;&nbsp;de <input type="text" name="dtemi1" size="8" value=""></td>
	<td class=titulo>a <input type="text" name="dtemi2" size="8" value="">&nbsp;</td>
</tr>
<tr>
	<td colspan="3" class=titulo><input type="radio" name="tipo_emissao" value="nul">sem data de pagamento&nbsp;&nbsp;</td>
</tr>
<tr>
	<td colspan="3" class=titulo>Filtrar descrição: <input type="text" name="descricao" size=25 value=""></td>
</tr>
</table>
<input type="submit" value="Visualizar relatório" name="Gerar" class="button"></p>
</form>
<p><font color="#FF0000">Configure a página do seu navegador (Internet
Explorer, Netscape, Mozilla, etc) no sentido PAISAGEM.</font></p>
<%
else
%>
<table border="0" cellpadding="2" width="950" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td align="left"  ><b>Relação de Pagamentos de RPA</b></td>
	<td align="center">&nbsp;</td>
	<td align="right" ><%=titulo%></td>
</tr>
</table>

<table border="0" cellpadding="1" width="950" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulor>Nome Prestador</td>
	<td class=titulor>C.P.F.  </td>
	<td class=titulor>PIS/NIT </td>
	<td class=titulor>Emissão </td>
	<td class=titulor>Pagto.  </td>
	<td class=titulor>Serviços</td>
	<td class=titulor>Outros  </td>
	<td class=titulor>I.Renda </td>
	<td class=titulor>INSS    </td>
	<td class=titulor>ISS     </td>
	<td class=titulor>Outros  </td>
	<td class=titulor>Líquido </td>
	<td class=titulor>INSS outra</td>
</tr>
<%
linha=2
rs.Open sql, ,adOpenStatic, adLockReadOnly
sp=0:ord=0:ir=0:inss=0:iss=0:ode=0:liq=0:oe=0
rs.movefirst
do while not rs.eof
if linha>45 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<table border='0' cellpadding='2' width='950' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  ><b>Relação de Pagamentos de RPA</b></td>"
	response.write "<td align='center'>&nbsp;</td>"
	response.write "<td align='right' >" & titulo & "</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table border='0' cellpadding='1' width='950' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=titulor>Nome Prestador</td>"
	response.write "<td class=titulor>C.P.F.  </td>"
	response.write "<td class=titulor>PIS/NIT </td>"
	response.write "<td class=titulor>Emissão </td>"
	response.write "<td class=titulor>Pagto.  </td>"
	response.write "<td class=titulor>Serviços</td>"
	response.write "<td class=titulor>Outros  </td>"
	response.write "<td class=titulor>I.Renda </td>"
	response.write "<td class=titulor>INSS    </td>"
	response.write "<td class=titulor>ISS     </td>"
	response.write "<td class=titulor>Outros  </td>"
	response.write "<td class=titulor>Líquido </td>"
	response.write "<td class=titulor>INSS outra</td>"
	response.write "</tr>"
	linha=2
end if
sp  =sp  +rs("servico_prestado")
ord =ord +rs("outros_rendimentos")
ir  =ir  +rs("desconto_ir")
inss=inss+rs("desconto_inss")
iss =iss +rs("desconto_iss")
ode =ode +rs("outros_descontos")
liq =liq +rs("valor_liquido")
oe  =oe  +rs("inss_outra_empresa")
cpf=rs("cpf")
if cpf<>"" then cpf=left(cpf,3) & "." & mid(cpf,4,3) & "." & mid(cpf,7,3) & "-" & right(cpf,2)
%>
<tr>
	<td class="campor"><%=rs("nome_autonomo")%></td>
	<td class="campor"><%=cpf%></td>
	<td class="campor"><%=rs("nit")%></td>
	<td class="campor"><%=rs("data_emissao")%></td>
	<td class="campor"><%=rs("data_pagamento")%></td>
	<td class="campor" align="right"><%=formatnumber(rs("servico_prestado"),2)%></td>
	<td class="campor" align="right"><%=formatnumber(rs("outros_rendimentos"),2)%></td>
	<td class="campor" align="right"><%=formatnumber(rs("desconto_ir"),2)%></td>
	<td class="campor" align="right"><%=formatnumber(rs("desconto_inss"),2)%></td>
	<td class="campor" align="right"><%=formatnumber(rs("desconto_iss"),2)%></td>
	<td class="campor" align="right"><%=formatnumber(rs("outros_descontos"),2)%></td>
	<td class="campor" align="right"><%=formatnumber(rs("valor_liquido"),2)%></td>
	<td class="campor" align="right"><%=formatnumber(rs("inss_outra_empresa"),2)%></td>
</tr>
<%
linha=linha+1
rs.movenext
loop
rs.close
%>
<tr>
	<td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
	<td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
	<td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
	<td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
	<td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
	<td class="campor" style="border-top: 1px solid #000000" align="right"><%=formatnumber(sp,2)%></td>
	<td class="campor" style="border-top: 1px solid #000000" align="right"><%=formatnumber(ord,2)%></td>
	<td class="campor" style="border-top: 1px solid #000000" align="right"><%=formatnumber(ir,2)%></td>
	<td class="campor" style="border-top: 1px solid #000000" align="right"><%=formatnumber(inss,2)%></td>
	<td class="campor" style="border-top: 1px solid #000000" align="right"><%=formatnumber(iss,2)%></td>
	<td class="campor" style="border-top: 1px solid #000000" align="right"><%=formatnumber(ode,2)%></td>
	<td class="campor" style="border-top: 1px solid #000000" align="right"><%=formatnumber(liq,2)%></td>
	<td class="campor" style="border-top: 1px solid #000000" align="right"><%=formatnumber(oe,2)%></td>
</tr>
</table>
<%
linha=linha+1
pagina=pagina+1
'response.write "<br>"
response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
'response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 

end if 'request.form
%>
</body>
</html>
<%
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%>