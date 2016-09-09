<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a52")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Pagamento RPA</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">

</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, varpar(5)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rss=server.createobject ("ADODB.Recordset")
Set rss.ActiveConnection = conexao
set rst=server.createobject ("ADODB.Recordset")
Set rst.ActiveConnection = conexao

if request("codigo")<>"" then id_lanc=request("codigo") else id_lanc=request.form("id_autonomo")

if request.form("bt_excluir")<>"" then
	tudook=1
	sql="DELETE FROM autonomo_rpa WHERE id_lanc=" & session("id_alt_lanc")
	if tudook=1 then conexao.Execute sql, , adCmdText
	id_lanc=0
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
	'if request("codigo")=null then
	'	id_lanc=session("id_alt_lanc")
	'else
	'	id_lanc=request("codigo")
	'end if
	sql1="select * from autonomo_rpa where id_lanc=" & id_lanc
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if


'**************** Cálculo
'if request.form("bt_excluir")="" then

'************ inicialização *************
if request.form("data_emissao")="" then data_emissao=rs("data_emissao") else data_emissao=request.form("data_emissao")
if request.form("data_pagamento")="" then data_pagamento=rs("data_pagamento") else data_pagamento=request.form("data_pagamento")
if request.form("descricao_servico")="" then descricao_servico=rs("descricao_servico") else descricao_servico=request.form("descricao_servico")
if request.form("cod_serv")="" then cod_serv=rs("cod_serv") else cod_serv=request.form("cod_serv")
if request.form("aliquota")="" then aliquota=rs("aliquota") else aliquota=request.form("aliquota")
if request.form("servico_prestado")="" then servico_prestado=rs("servico_prestado") else servico_prestado=request.form("servico_prestado")
if request.form("desc_outros_rend")="" then desc_outros_rend=rs("desc_outros_rend") else desc_outros_rend=request.form("desc_outros_rend")
if request.form("outros_rendimentos")="" then outros_rendimentos=rs("outros_rendimentos") else outros_rendimentos=request.form("outros_rendimentos")
if request.form("salcont")="" then salcont=rs("salcont") else salcont=request.form("salcont")
if request.form("desconto_ir")="" then desconto_ir=rs("desconto_ir") else desconto_ir=request.form("desconto_ir")
if request.form("desconto_iss")="" then desconto_iss=rs("desconto_iss") else desconto_iss=request.form("desconto_iss")
if request.form("desconto_inss")="" then desconto_inss=rs("desconto_inss") else desconto_inss=request.form("desconto_inss")
if request.form("descricao_outros")="" then descricao_outros=rs("descricao_outros") else descricao_outros=request.form("descricao_outros")
if request.form("outros_descontos")="" then outros_descontos=rs("outros_descontos") else outros_descontos=request.form("outros_descontos")
if request.form("valor_liquido")="" then valor_liquido=rs("valor_liquido") else valor_liquido=request.form("valor_liquido")
if request.form("empresasc")="" then empresasc=rs("empresasc") else empresasc=request.form("empresasc")
if request.form("cnpjsc")="" then cnpjsc=rs("cnpjsc") else cnpjsc=request.form("cnpjsc")
if request.form("oe_valor")="" then oe_valor=rs("oe_valor") else oe_valor=request.form("oe_valor")
if request.form("oe_sc")="" then oe_sc=rs("oe_sc") else oe_sc=request.form("oe_sc")
if request.form("inss_outra_empresa")="" then inss_outra_empresa=rs("inss_outra_empresa") else inss_outra_empresa=request.form("inss_outra_empresa")
if request.form("proporcionaliza")="" then proporcionaliza=rs("proporcionaliza") else proporcionaliza=request.form("proporcionaliza")
if request.form("forcar_valores")="" then forcar_valores=rs("forcar_valores") else forcar_valores=request.form("forcar_valores")
if request("codigo")<>"" then id_autonomo=request("codigo") else id_autonomo=request.form("id_autonomo")
if request.form("apuracao_darf")="" then apuracao_darf=rs("apuracao_darf") else apuracao_darf=request.form("apuracao_darf")
if request.form("venc_darf")="" then venc_darf=rs("venc_darf") else venc_darf=request.form("venc_darf")
if request.form("ctl_darf")="" then ctl_darf=rs("ctl_darf") else ctl_darf=request.form("ctl_darf")

	'if request.form("data_pagamento")<>"" then data_pagamento=rs("data_pagamento") else data_pagamento=request.form("data_pagamento")
	if data_pagamento<>"" then
		diasem=weekday(data_pagamento)
		dias=7-diasem
		apuracao_darf=dateadd("d",cint(dias),formatdatetime(data_pagamento,2))
		venc_darf=dateadd("d",4,formatdatetime(apuracao_darf,2))
	End if
	'descricao_servico=request.form("descricao_servico"):if descricao_servico="" then descricao_servico="Servicos Prestados"
	'servico_prestado=request.form("servico_prestado"):if servico_prestado="" then servico_prestado=0
	'outros_rendimentos=request.form("outros_rendimentos"):if outros_rendimentos="" then outros_rendimentos=0
	'desconto_ir=request.form("desconto_ir"):if desconto_ir="" then desconto_ir=0
	'desconto_iss=request.form("desconto_iss"):if desconto_iss="" then desconto_iss=0
	'desconto_inss=request.form("desconto_inss"):if desconto_inss="" then desconto_inss=0
	'outros_descontos=request.form("outros_descontos"):if outros_descontos="" then outros_descontos=0
	'valor_liquido=request.form("valor_liquido"):if valor_liquido="" then valor_liquido=0
	'inss_outra_empresa=request.form("inss_outra_empresa"):if inss_outra_empresa="" then inss_outra_empresa=0
	'oe_valor=request.form("oe_valor"):if oe_valor="" then oe_valor=0
	'oe_sc=request.form("oe_sc"):if oe_sc="" then oe_sc=0
	'salcont=request.form("salcont"):if salcont="" then salcont=0
	'cod_serv=request.form("cod_serv"):if cod_serv="" then cod_serv="17.01"
	if cod_serv<>"" then
		sqlaliq="select aliquota from autonomo_iss where cod_serv='" & cod_serv & "' "
		rss.Open sqlaliq, ,adOpenStatic, adLockReadOnly
		aliquota=rss("aliquota")
		rss.close
	else
		aliquota=request.form("aliquota"):if aliquota="" then aliquota=2
	end if
	
	valor=cdbl(servico_prestado)+cdbl(outros_rendimentos) 'Valor base
	'data_emissao=request.form("data_emissao")
	if data_emissao="" then data_emissao=formatdatetime(now(),2)
	sqlinss="select faixa1=max(case when NROFAIXA=1 then LIMITESUPERIOR else 0 end), perc1=max(case when NROFAIXA=1 then PERCENTUAL else 0 end) " & _
	", faixa2=max(case when NROFAIXA=2 then LIMITESUPERIOR else 0 end), perc2=max(case when NROFAIXA=2 then PERCENTUAL else 0 end) " & _
	", faixa3=max(case when NROFAIXA=3 then LIMITESUPERIOR else 0 end), perc3=max(case when NROFAIXA=3 then PERCENTUAL else 0 end) " & _
	"from corporerm.dbo.PCALCVLR where CODTABCALC='01' and INICIOVIGENCIA=(select top 1 INICIOVIGENCIA from corporerm.dbo.PCALCVLR where INICIOVIGENCIA<='" & dtaccess(data_emissao) & "' and CODTABCALC='01' order by INICIOVIGENCIA desc)"
	rst.Open sqlinss, ,adOpenStatic, adLockReadOnly
	f1=cdbl(rst("faixa1")):f2=cdbl(rst("faixa2")):f3=cdbl(rst("faixa3")):p1=cdbl(rst("perc1")):p2=cdbl(rst("perc2")):p3=cdbl(rst("perc3"))
	if valor<=cdbl(rst("faixa1")) then 
		percinss=rst("perc1")
	elseif valor<=cdbl(rst("faixa2")) then 
		percinss=rst("perc2")
	elseif valor<=cdbl(rst("faixa3")) then 
		percinss=rst("perc3")
	end if
	limiteinss=int(cdbl(rst("faixa3"))*cint(rst("perc3")))/100
	tetoinss=cdbl(rst("faixa3"))
	if cint(percinss)=0 then
		calcinss=int(cdbl(rst("faixa3"))*cint(rst("perc3")))/100
	else
		calcinss=int(valor*cdbl(percinss))/100
	end if
	rst.close
	'response.write "<br>" & valor & "-" & percinss & "-" & calcinss

	if request.form("proporcionaliza")="ON" then proporcionaliza=1 else proporcionaliza=0
	if request.form("inscrito")<>"" then inscrito=1 else inscrito=0
	if request.form("forcar_valores")="ON" or rs("forcar_valores")=-1 then forcar_valores=-1 else forcar_valores=0

	if forcar_valores=0 then
		desconto_inss=calcinss
		descontado=inss_outra_empresa

		basef=0:baseo=0:base=0
		basef=cdbl(servico_prestado)+cdbl(outros_rendimentos)
		baseo=cdbl(oe_valor)
		base=basef+baseo
		if basef<tetoinss then f=0 else f=1
		if baseo<tetoinss then o=0 else o=1
		if base<tetoinss then t=0 else t=1
		'msgbox f & " " & o & " " & t
		'******** DESCONTO INSS ************
		if f=0 and o=0 and t=0 then 'ambos valores < teto e soma < texto
			'msgbox "Situacao 1"
			scf=basef : sco=baseo
					if sco<=f1 then 
						calcinss_sco=sco*p1
					elseif sco<=f2 then 
						calcinss_sco=sco*p2
					elseif valor<=f3 then 
						calcinss_sco=sco*p3
					else
						calcinss_sco=f3*p3
					end if
			inssf=int(scf*20)/100 : insso=int(calcinss_sco)/100
		end if
		if f=0 and o=0 and t=1 then 'ambos valores < teto e soma > texto
			if proporcionaliza=0 then
				'msgbox "Situacao 2-1"
				sco=baseo : scf=tetoinss-sco
					if sco<=f1 then 
						calcinss_sco=sco*p1
					elseif sco<=f2 then 
						calcinss_sco=sco*p2
					elseif valor<=f3 then 
						calcinss_sco=sco*p3
					else
						calcinss_sco=f3*p3
					end if
				insso=int(calcinss_sco)/100 : inssf=int(scf*20)/100
			else
				'msgbox "Situacao 2-2"
				sco=baseo*tetoinss/base:sco=int(sco*100)/100
				scf=basef*tetoinss/base:scf=tetoinss-sco
				sco=baseo : scf=tetoinss-sco
					if sco<=f1 then 
						calcinss_sco=sco*p1
					elseif sco<=f2 then 
						calcinss_sco=sco*p2
					elseif valor<=f3 then 
						calcinss_sco=sco*p3
					else
						calcinss_sco=f3*p3
					end if
				insso=int(calcinss_sco)/100 : inssf=int(scf*20)/100
			end if
		end if
		if (f=1 or o=1) and t=1 then
			if proporcionaliza=0 then
				'msgbox "Situacao 3-1"
				sco=baseo : if sco>tetoinss then sco=tetoinss
				scf=tetoinss-sco
				sco=baseo : scf=tetoinss-sco
					if sco<=f1 then 
						calcinss_sco=sco*p1
					elseif sco<=f2 then 
						calcinss_sco=sco*p2
					elseif valor<=f3 then 
						calcinss_sco=sco*p3
					else
						calcinss_sco=f3*p3
					end if
				insso=int(calcinss_sco)/100 : inssf=int(scf*20)/100
			else
				'msgbox "Situacao 3-2"
				sco=baseo*tetoinss/base:sco=int(sco*100)/100
				scf=basef*tetoinss/base:scf=tetoinss-sco
				sco=baseo : scf=tetoinss-sco
					if sco<=f1 then 
						calcinss_sco=sco*p1
					elseif sco<=f2 then 
						calcinss_sco=sco*p2
					elseif valor<=f3 then 
						calcinss_sco=sco*p3
					else
						calcinss_sco=f3*p3
					end if
				insso=int(calcinss_sco)/100 : inssf=int(scf*20)/100
			end if
		end if		

		oe_valor=baseo
		oe_sc=sco
		salcont=scf
		desconto_inss=inssf : if desconto_inss<0 then desconto_inss=0
		descontado=insso

		'******** DESCONTO IRRF ************
		baseir=basef-desconto_inss
		sqlir="select limite1=max(case when NROFAIXA=1 then LIMITESUPERIOR else 0 end), aliq1=max(case when NROFAIXA=1 then PERCENTUAL else 0 end), deducao1=max(case when NROFAIXA=1 then VALDEDUZIR else 0 end) " & _
		", limite2=max(case when NROFAIXA=2 then LIMITESUPERIOR else 0 end), aliq2=max(case when NROFAIXA=2 then PERCENTUAL else 0 end), deducao2=max(case when NROFAIXA=2 then VALDEDUZIR else 0 end) " & _
		", limite3=max(case when NROFAIXA=3 then LIMITESUPERIOR else 0 end), aliq3=max(case when NROFAIXA=3 then PERCENTUAL else 0 end), deducao3=max(case when NROFAIXA=3 then VALDEDUZIR else 0 end) " & _
		", limite4=max(case when NROFAIXA=4 then LIMITESUPERIOR else 0 end), aliq4=max(case when NROFAIXA=4 then PERCENTUAL else 0 end), deducao4=max(case when NROFAIXA=4 then VALDEDUZIR else 0 end) " & _
		", limite5=max(case when NROFAIXA=5 then LIMITESUPERIOR else 0 end), aliq5=max(case when NROFAIXA=5 then PERCENTUAL else 0 end), deducao5=max(case when NROFAIXA=5 then VALDEDUZIR else 0 end) " & _
		"from corporerm.dbo.PCALCVLR where CODTABCALC='02' and INICIOVIGENCIA=(select top 1 INICIOVIGENCIA from corporerm.dbo.PCALCVLR where INICIOVIGENCIA<='" & dtaccess(data_emissao) & "' and CODTABCALC='02' order by INICIOVIGENCIA desc)"
		rst.Open sqlir, ,adOpenStatic, adLockReadOnly
		if baseir<=cdbl(rst("limite1")) then
			vir=0
		elseif baseir<=cdbl(rst("limite2")) then
			vir=int(baseir*cdbl(rst("aliq2"))+0.5)/100:vir=vir-cdbl(rst("deducao2"))
		elseif baseir<=cdbl(rst("limite3")) then
			vir=int(baseir*cdbl(rst("aliq3"))+0.5)/100:vir=vir-cdbl(rst("deducao3"))
		elseif baseir<=cdbl(rst("limite4")) then
			vir=int(baseir*cdbl(rst("aliq4"))+0.5)/100:vir=vir-cdbl(rst("deducao4"))
		else
			vir=int(baseir*cdbl(rst("aliq5"))+0.5)/100:vir=vir-cdbl(rst("deducao5"))
		end if		
		if vir<0 then vir=0
		desconto_ir=vir
		rst.close

		'******** DESCONTO ISS ************
		baseiss=basef
		if inscrito=0 then desconto_iss=int(baseiss*aliquota+0.5)/100 else desconto_iss=0
	end if		

	servico_prestado  =formatnumber(servico_prestado,2)
	outros_rendimentos=formatnumber(outros_rendimentos,2)
	desconto_inss     =formatnumber(desconto_inss,2)
	desconto_ir       =formatnumber(desconto_ir,2)
	desconto_iss      =formatnumber(desconto_iss,2)
	outros_descontos  =formatnumber(outros_descontos,2)
	inss_outra_empresa=formatnumber(descontado,2)
	total_rendimentos=cdbl(servico_prestado)+outros_rendimentos
	total_descontos=cdbl(desconto_inss)+desconto_ir+desconto_iss+outros_descontos
	total_liquido=cdbl(total_rendimentos)-total_descontos
	total_rend        =formatnumber(total_rendimentos,2)
	total_desc        =formatnumber(total_descontos,2)
	valor_liquido     =formatnumber(total_liquido,2)
	oe_valor          =formatnumber(oe_valor,2)
	oe_sc             =formatnumber(sco,2)
	if scf=0 then scf=salcont
	salcont           =formatnumber(scf,2)
	
'end if

if request.form("bt_salvar")<>"" then
	tudook=1
	sql="UPDATE autonomo_rpa SET "
	if request.form("data_emissao")<>"" then sql=sql & "data_emissao = '"   & dtaccess(request.form("data_emissao")) & "', " else sql=sql & "data_emissao = null, "
	if request.form("data_pagamento")<>"" then sql=sql & "data_pagamento = '"   & dtaccess(request.form("data_pagamento")) & "', " else sql=sql & "data_pagamento = null, "
	if request.form("apuracao_darf")<>"" then sql=sql & "apuracao_darf = '"   & dtaccess(request.form("apuracao_darf")) & "', " else sql=sql & "apuracao_darf = null, "
	if request.form("venc_darf")<>"" then sql=sql & "venc_darf = '"   & dtaccess(request.form("venc_darf")) & "', " else sql=sql & "venc_darf = null, "
	sql=sql & "cod_serv          = '" & request.form("cod_serv")          & "', "
	sql=sql & "aliquota          = " & request.form("aliquota") & ", "
	sql=sql & "descricao_servico = '" & request.form("descricao_servico") & "', "
	sql=sql & "desc_outros_rend  = '" & request.form("desc_outros_rend")  & "', "
	sql=sql & "descricao_outros  = '" & request.form("descricao_outros")  & "', "
	sql=sql & "servico_prestado  = "  & nraccess(cdbl(request.form("servico_prestado")))  & ", "
	sql=sql & "outros_rendimentos= "  & nraccess(cdbl(request.form("outros_rendimentos")))  & ", "
	sql=sql & "desconto_inss     = "  & nraccess(cdbl(request.form("desconto_inss")))  & ", "
	sql=sql & "desconto_ir       = "  & nraccess(cdbl(request.form("desconto_ir")))  & ", "
	sql=sql & "desconto_iss      = "  & nraccess(cdbl(request.form("desconto_iss")))  & ", "
	sql=sql & "outros_descontos  = "  & nraccess(cdbl(request.form("outros_descontos")))  & ", "
	sql=sql & "valor_liquido     = "  & nraccess(cdbl(request.form("valor_liquido")))  & ", "
	sql=sql & "inss_outra_empresa= "  & nraccess(cdbl(request.form("inss_outra_empresa")))  & ", "
	sql=sql & "ctl_darf          = '" & request.form("ctl_darf")  & "', "

	sql=sql & "empresasc         = '" & request.form("empresasc")  & "', "
	sql=sql & "cnpjsc            = '" & request.form("cnpjsc")  & "', "
	sql=sql & "oe_valor          = "  & nraccess(cdbl(request.form("oe_valor")))  & ", "
	sql=sql & "oe_sc             = "  & nraccess(cdbl(request.form("oe_sc")))  & ", "
	sql=sql & "salcont           = "  & nraccess(cdbl(request.form("salcont")))  & ", "

	if request.form("proporcionaliza")="ON" then sql=sql & "proporcionaliza = 1, " else sql=sql & "proporcionaliza = 0, "
	if request.form("forcar_valores")="ON" then sql=sql & "forcar_valores = 1, " else sql=sql & "forcar_valores = 0, "
	if request.form("emitiu_darf")="0" then sql=sql & "emitiu_darf = 1, " else sql=sql & "emitiu_darf = 0, "
	if request.form("emitiu_rpa")="0" then sql=sql & "emitiu_rpa = 1 " else sql=sql & "emitiu_rpa = 0 "

	sql=sql & " WHERE id_lanc=" & session("id_alt_lanc")
	'response.write sql
	response.write request.form("cod_serv")
	if tudook=1 then conexao.Execute sql, , adCmdText
end if

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) then
'if request.form="" then
'rs.movefirst
'do while not rs.eof 
session("id_alt_lanc")=rs("id_lanc")

sql="select nome_autonomo, tipo_prestacao, ccm from autonomo where id_autonomo=" & rs("id_autonomo")
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
rsc.Open sql, ,adOpenStatic, adLockReadOnly
nome_autonomo=rsc("nome_autonomo")
tipo_prestacao=rsc("tipo_prestacao")
inscrito=rsc("ccm")
rsc.close
if rs("proporcionaliza")=-1 then proporcionaliza="checked" else proporcionaliza=""
if rs("forcar_valores")=-1 then forcar_valores=-1 else forcar_valores=0

%>
<form method="POST" action="rpa_alteracao.asp" name="form">
<input type="hidden" name="id_autonomo" value="<%=id_autonomo%>">
<input type="hidden" name="inscrito" value="<%=inscrito%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr><td class=grupo>Alteração de Pagamento RPA - <%=id_autonomo%> - <%=nome_autonomo%> </td></tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Emissão</td>
	<td class=titulo>Pagamento</td>
	<td class=titulo>Descrição dos Serviços</td>
</tr>
<tr>
	<td class=titulo valign=top><input type="text" name="data_emissao" size="8" value="<%=data_emissao%>" onchange="javascript:submit()"></td>
	<td class=titulo valign=top><input type="text" name="data_pagamento" size="8" value="<%=data_pagamento%>"></td>
	<td class=titulo><textarea name="descricao_servico" cols="45" rows="2"><%=descricao_servico%></textarea></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Classificação do Serviço</td>
	<td class=titulo>Alíquota ISS</td>
</tr>
<tr>
	<td class=titulo valign=top>
	<select name="cod_serv" onchange="javascript:submit()">
<%
	sqlserv="select cod_serv, desc_serv, usado from autonomo_iss order by cod_serv"
	rss.Open sqlserv, ,adOpenStatic, adLockReadOnly
	do while not rss.eof
	cs=rss("cod_serv")
	if rss("cod_serv")=cod_serv then txtcodserv="selected" else txtcodserv=""
	if rss("usado")="X" then opfun="#0C00FF" else opfun="#000"
%>	
	<option style="color:<%=opfun%>;" value="<%=rss("cod_serv")%>" <%=txtcodserv%>><%=left(rss("desc_serv"),70)%></option>
<%	
	rss.movenext
	loop
	rss.close
%>
	</select>
	</td>
	<td class=titulo valign=top><input type="text" name="aliquota" size="3" value="<%=aliquota%>" onchange="javascript:submit()"></td>
</tr>
</table>


<table border="0" cellpadding="2" cellspacing="0" width="500">
<tr>
	<td class=titulo width="172">Rendimentos</td>
	<td class=titulo width="74" align="center">Valor R$&nbsp;&nbsp;</td>
	<td class=titulo width="172">Descontos</td>
	<td class=titulo width="74" align="center">Valor R$&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class=titulo colspan="2" valign=top>
<!-- quadro rendimentos -->		
	<table border="0" cellpadding="0" cellspacing="0" width=246>
	<tr>
		<td class=fundo width="70%">Valor Serviços Prestados</td>
		<td class=fundo width="30%"><input type="text" name="servico_prestado" size="8" value="<%=formatnumber(servico_prestado,2)%>" class=vr onchange="javascript:submit()" ></td>
	</tr>
	<tr>
		<td class=fundor>Outros Rendimentos</td>
		<td class=fundo>&nbsp;</td>
	</tr>
	<tr>
		<td class=fundo><input type="text" name="desc_outros_rend" size="25" value="<%=desc_outros_rend%>"></td>
		<td class=fundo><input type="text" name="outros_rendimentos" size="8" value="<%=formatnumber(outros_rendimentos,2)%>" class=vr onchange="javascript:submit()"></td>
	</tr>
	<tr>
		<td class=fundo style="border-bottom: 1px solid #000000">&nbsp;</td>
		<td class=fundo style="border-bottom: 1px solid #000000">&nbsp;</td>
	</tr>
	<tr>
		<td class=fundo align="right">Total Rendimentos&nbsp;</td>
		<td class=fundo><input type="text" name="total_rend"  size="8" value="<%=formatnumber(cdbl(servico_prestado)+outros_rendimentos,2)%>" class=bloq onFocus="total_rend.blur()"></td>
	</tr>
	<tr>
		<td class=fundo>&nbsp;</td>
		<td class=fundo>&nbsp;</td>
	</tr>
	<tr>
		<td class=fundo colspan=2>S.C. Fieo <input type="text" name="salcont" size=7 value="<%if salcont>=0 then response.write formatnumber(salcont,2)%>" class="vr" onchange="javascript:submit()">
		Proporcionaliza? <input type="checkbox" name="proporcionaliza" value="ON" <%=proporcionaliza%>>
                   
	</tr>
</table>	
<!-- quadro rendimentos fim-->				
		</td>
		<td class=titulo colspan="2" valign=top>
<!-- quadro descontos -->		
	<table border="0" cellpadding="0" cellspacing="0" width=246>
	<tr>
		<td class=fundo width="70%">Imp.Renda Fonte</td>
		<td class=fundo width="30%"><input type="text" name="desconto_ir" size="8" value="<%=formatnumber(desconto_ir,2)%>" class="vr" onchange="javascript:submit()"></td>
	</tr>
	<tr>
		<td class=fundo>ISS - Fonte</td>
		<td class=fundo><input type="text" name="desconto_iss" size="8" value="<%=formatnumber(desconto_iss,2)%>" class="vr" onchange="javascript:submit()"></td>
	</tr>
	<tr>
		<td class=fundo>INSS</td>
		<td class=fundo><input type="text" name="desconto_inss" size="8" value="<%=formatnumber(desconto_inss,2)%>" class="vr" onchange="javascript:submit()"></td>
	</tr>
	<tr>
		<td class=fundor>Outros Descontos</td>
		<td class=fundo>&nbsp;</td>
	</tr>
	<tr>
		<td class=fundo><input type="text" name="descricao_outros" size="25" value="<%=descricao_outros%>"></td>
		<td class=fundo><input type="text" name="outros_descontos" size="8" value="<%=formatnumber(outros_descontos,2)%>" class="vr" onchange="javascript:submit()"></td>
	</tr>
	<tr>
		<td class=fundo style="border-bottom: 1px solid #000000">&nbsp;</td>
		<td class=fundo style="border-bottom: 1px solid #000000">&nbsp;</td>
	</tr>
	<tr>
		<td class=fundo align="right">Total Descontos&nbsp;</td>
		<td class=fundo><input type="text" name="total_desc"  size="8" value="<%=formatnumber(cdbl(desconto_ir)+desconto_iss+desconto_inss+outros_descontos,2)%>" class=bloq onFocus="total_desc.blur()"></td>
	</tr>
</table>
<!-- quadro descontos fim-->				
		</td>
	</tr>
	<tr>
		<td class=fundo width="50%" colspan="2"><input type="checkbox" name="forcar_valores" value="ON" <%if forcar_valores=-1 then response.write "checked"%> >&nbsp;Forçar Valores</td>
		<td class=titulo width="35%">Valor Líquido</td>
		<td class=fundo width="15%"><input type="text" name="valor_liquido" size="8" value="<%=formatnumber(valor_liquido,2)%>" class="bloq" onFocus="valor_liquido.blur()"></td>
	</tr>
</table>  

<hr style="background-color:Silver">
<table border="0" cellpadding="2" cellspacing="0" width="500">
<tr>
	<td colspan=5 class=fundo>Contribuições ao INSS outra empresa:</td>
</tr>
<tr>
	<td class=fundo>Empresa</td>
	<td class=fundo>CNPJ</td>
	<td class=fundo>Vr.Rend.</td>
	<td class=fundo>Vr SC</td>
	<td class=fundo>INSS desc</td>
</tr>
<tr>
	<td class=fundo><input type="text" name="empresasc" size=30 value="<%=empresasc%>"></td>
	<td class=fundo><input type="text" name="cnpjsc"    size=15 value="<%=cnpjsc%>"></td>
	<td class=fundo><input type="text" name="oe_valor"   size=7  value="<%=formatnumber(oe_valor,2)%>" class="vr" onchange="javascript:submit()">
	<td class=fundo><input type="text" name="oe_sc"   size=7  value="<%=formatnumber(oe_sc,2)%>" class="vr" onchange="javascript:submit()">
	<td class=fundo><input type="text" name="inss_outra_empresa" size="8" value="<%=formatnumber(inss_outra_empresa,2)%>" class="vr" onchange="javascript:submit()"></td>
</tr>
</table>

<!-- fim tabela -->  
  <table border="0" cellpadding="3" cellspacing="0" width="500">
    <tr>
      <td class=grupor colspan=3>DARF</td>
      <td class=grupor colspan=2>Emissão</td>
    </tr>
    <tr>
      <td class=titulo>Apuração</td>
      <td class=titulo>Vencimento</td>
      <td class=titulo>Controle</td>
      <td class=titulo>RPA</td>
      <td class=titulo>DARF</td>
    </tr>
<%
if rs("emitiu_darf")=0 then emitiu_darf="" else emitiu_darf="checked"
if rs("emitiu_rpa")=0 then emitiu_rpa="" else emitiu_rpa="checked"
%>
    <tr>
      <td class=titulo valign=top><input type="text" name="apuracao_darf" size="8" value="<%=apuracao_darf%>"></td>
      <td class=titulo valign=top><input type="text" name="venc_darf" size="8" value="<%=venc_darf%>"></td>
      <td class=titulo valign=top><input type="text" name="ctl_darf" size="15" value="<%=ctl_darf%>"></td>
	  <td class=fundo ><input type="checkbox" name="emitiu_rpa" value="0" <%=emitiu_rpa%>>&nbsp;</td>
	  <td class=fundo ><input type="checkbox" name="emitiu_darf" value="0" <%=emitiu_darf%>>&nbsp;</td>
    </tr>
  </table>

  <table border="0" cellpadding="3" cellspacing="0" width="500">
    <tr>
      <td class=titulo align="center">
		<input type="submit" value="Salvar Alterações" class="button" name="Bt_salvar">
      </td>
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