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
<title>Inclusão de Pagamento RPA</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="VBScript">
<!--
	Sub Calcula()
	End Sub
// -->
</script>

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
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

'**************** Cálculo
if request.form<>"" then
	data_pagamento=request.form("data_pagamento")
	response.write data_pagamento
	if data_pagamento<>"" then
		diasem=weekday(data_pagamento)
		dias=7-diasem
		apuracao_darf=dateadd("d",cint(dias),formatdatetime(data_pagamento,2))
		venc_darf=dateadd("d",4,formatdatetime(apuracao_darf,2))
	End if

	descricao_servico=request.form("descricao_servico"):if descricao_servico="" then descricao_servico="Servicos Prestados"
	cod_serv=request.form("cod_serv"):if cod_serv="" then cod_serv="17.01"
	if cod_serv<>"" then
		sqlaliq="select aliquota from autonomo_iss where cod_serv='" & cod_serv & "' "
		rss.Open sqlaliq, ,adOpenStatic, adLockReadOnly
		aliquota=rss("aliquota")
		rss.close
	else
		aliquota=request.form("aliquota"):if aliquota="" then aliquota=2
	end if
	servico_prestado=request.form("servico_prestado"):if servico_prestado="" then servico_prestado=0
	outros_rendimentos=request.form("outros_rendimentos"):if outros_rendimentos="" then outros_rendimentos=0
	desconto_ir=request.form("desconto_ir"):if desconto_ir="" then desconto_ir=0
	desconto_iss=request.form("desconto_iss"):if desconto_iss="" then desconto_iss=0
	desconto_inss=request.form("desconto_inss"):if desconto_inss="" then desconto_inss=0
	outros_descontos=request.form("outros_descontos"):if outros_descontos="" then outros_descontos=0
	valor_liquido=request.form("valor_liquido"):if valor_liquido="" then valor_liquido=0
	inss_outra_empresa=request.form("inss_outra_empresa"):if inss_outra_empresa="" then inss_outra_empresa=0
	oe_valor=request.form("oe_valor"):if oe_valor="" then oe_valor=0
	oe_sc=request.form("oe_sc"):if oe_sc="" then oe_sc=0
	salcont=request.form("salcont"):if salcont="" then salcont=0

	valor=cdbl(servico_prestado)+cdbl(outros_rendimentos) 'Valor base
	data_emissao=request.form("data_emissao")
	if data_emissao="" then data_emissao=formatdatetime(now(),2)
	sqlinss="select faixa1=max(case when NROFAIXA=1 then LIMITESUPERIOR else 0 end), perc1=max(case when NROFAIXA=1 then PERCENTUAL else 0 end) " & _
	", faixa2=max(case when NROFAIXA=2 then LIMITESUPERIOR else 0 end), perc2=max(case when NROFAIXA=2 then PERCENTUAL else 0 end) " & _
	", faixa3=max(case when NROFAIXA=3 then LIMITESUPERIOR else 0 end), perc3=max(case when NROFAIXA=3 then PERCENTUAL else 0 end) " & _
	"from corporerm.dbo.PCALCVLR where CODTABCALC='01' and INICIOVIGENCIA=(select top 1 INICIOVIGENCIA from corporerm.dbo.PCALCVLR where INICIOVIGENCIA<='" & dtaccess(data_emissao) & "' and CODTABCALC='01' order by INICIOVIGENCIA desc)"
	rs.Open sqlinss, ,adOpenStatic, adLockReadOnly
	f1=cdbl(rs("faixa1")):f2=cdbl(rs("faixa2")):f3=cdbl(rs("faixa3")):p1=cdbl(rs("perc1")):p2=cdbl(rs("perc2")):p3=cdbl(rs("perc3"))
	if valor<=cdbl(rs("faixa1")) then 
		percinss=rs("perc1")
	elseif valor<=cdbl(rs("faixa2")) then 
		percinss=rs("perc2")
	elseif valor<=cdbl(rs("faixa3")) then 
		percinss=rs("perc3")
	end if
	limiteinss=int(cdbl(rs("faixa3"))*cint(rs("perc3")))/100
	tetoinss=cdbl(rs("faixa3"))
	if cint(percinss)=0 then
		calcinss=int(cdbl(rs("faixa3"))*cint(rs("perc3")))/100
	else
		calcinss=int(valor*cdbl(percinss))/100
	end if
	rs.close
	'response.write "<br>" & valor & "-" & percinss & "-" & calcinss
	desconto_inss=calcinss

	descontado=inss_outra_empresa
	if request.form("proporcionaliza")="ON" then proporcionaliza=1 else proporcionaliza=0
	if request.form("inscrito")<>"" then inscrito=1 else inscrito=0

	if request.form("forcar_valores")<>"ON" then
	'if document.form.forcar_valores.checked=false then
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
		"from corporerm.dbo.PCALCVLR where CODTABCALC='02' and INICIOVIGENCIA=(select top 1 INICIOVIGENCIA from corporerm.dbo.PCALCVLR where INICIOVIGENCIA<='" & dtaccess(data_emissao) & "' and codtabcalc='02' order by INICIOVIGENCIA desc)"
		rs.Open sqlir, ,adOpenStatic, adLockReadOnly
		if baseir<=cdbl(rs("limite1")) then
			vir=0
		elseif baseir<=cdbl(rs("limite2")) then
			vir=int(baseir*cdbl(rs("aliq2"))+0.5)/100:vir=vir-cdbl(rs("deducao2"))
		elseif baseir<=cdbl(rs("limite3")) then
			vir=int(baseir*cdbl(rs("aliq3"))+0.5)/100:vir=vir-cdbl(rs("deducao3"))
		elseif baseir<=cdbl(rs("limite4")) then
			vir=int(baseir*cdbl(rs("aliq4"))+0.5)/100:vir=vir-cdbl(rs("deducao4"))
		else
			vir=int(baseir*cdbl(rs("aliq5"))+0.5)/100:vir=vir-cdbl(rs("deducao5"))
		end if		
		if vir<0 then vir=0
		desconto_ir=vir
		rs.close

		'******** DESCONTO ISS ************
		baseiss=basef
		if inscrito=0 then
			desconto_iss=int(baseiss*aliquota+0.5)/100 
		else 
			desconto_iss=0
		end if
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
	salcont           =formatnumber(scf,2)
	
end if

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=1
		sql = "INSERT INTO autonomo_rpa (" 
		sql = sql & "id_autonomo, data_emissao, descricao_servico, cod_serv, aliquota, servico_prestado, desconto_ir, "
		sql = sql & "desconto_iss, valor_liquido, desconto_inss, inss_outra_empresa, "
		sql = sql & "emitiu_rpa, emitiu_darf, proporcionaliza "
		if request.form("desc_outros_rend")<>""   then sql=sql & ", desc_outros_rend"
		if request.form("outros_rendimentos")<>"" then sql=sql & ", outros_rendimentos"
		if request.form("descricao_outros")<>""   then sql=sql & ", descricao_outros"
		if request.form("outros_descontos")<>""   then sql=sql & ", outros_descontos"
		if request.form("data_pagamento")<>""     then sql=sql & ", data_pagamento"
		if request.form("apuracao_darf")<>""      then sql=sql & ", apuracao_darf"
		if request.form("venc_darf")<>""          then sql=sql & ", venc_darf"
		if request.form("empresasc")<>""          then sql=sql & ", empresasc"
		if request.form("cnpjsc")<>""             then sql=sql & ", cnpjsc"
		if request.form("oe_valor")<>""           then sql=sql & ", oe_valor"
		if request.form("oe_sc")<>""              then sql=sql & ", oe_sc"
		if request.form("salcont")<>""            then sql=sql & ", salcont"
		sql = sql & ") "
		sql2 = " SELECT "
		sql2=sql2 & " " & request.form("id_autonomo") & ", "
		sql2=sql2 & " '" & dtaccess(request.form("data_emissao")) & "', "
		sql2=sql2 & " '" & request.form("descricao_servico") & "', "
		sql2=sql2 & " '" & request.form("cod_serv") & "', "
		sql2=sql2 & " '" & request.form("aliquota") & "', "
		sql2=sql2 & " " & nraccess(cdbl(request.form("servico_prestado"))) & ", "
		sql2=sql2 & " " & nraccess(cdbl(request.form("desconto_ir"))) & ", "
		sql2=sql2 & " " & nraccess(cdbl(request.form("desconto_iss"))) & ", "
		sql2=sql2 & " " & nraccess(cdbl(request.form("valor_liquido"))) & ", "
		sql2=sql2 & " " & nraccess(cdbl(request.form("desconto_inss"))) & ", "
		sql2=sql2 & " " & nraccess(cdbl(request.form("inss_outra_empresa"))) & ", "
		sql2=sql2 & " 0, "
		if cdbl(request.form("desconto_ir"))=0 then emitiu_darf=1 else emitiu_darf=0
		sql2=sql2 & " " & emitiu_darf & ", "
		if request.form("proporcionaliza")="ON" then proporcionaliza=1 else proporcionaliza=0
		sql2=sql2 & proporcionaliza
		if request.form("desc_outros_rend")<>""   then sql2=sql2 & ", '" & request.form("desc_outros_rend") & "' "
		if request.form("outros_rendimentos")<>"" then sql2=sql2 & ", " &  nraccess(cdbl(request.form("outros_rendimentos"))) & " "
		if request.form("descricao_outros")<>""   then sql2=sql2 & ", '" & request.form("descricao_outros") & "' "
		if request.form("outros_descontos")<>""   then sql2=sql2 & ", " &  nraccess(cdbl(request.form("outros_descontos"))) & " "
		if request.form("data_pagamento")<>""     then sql2=sql2 & ", '" & dtaccess(request.form("data_pagamento")) & "' "
		if request.form("apuracao_darf")<>""      then sql2=sql2 & ", '" & dtaccess(request.form("apuracao_darf")) & "' "
		if request.form("venc_darf")<>""          then sql2=sql2 & ", '" & dtaccess(request.form("venc_darf")) & "' "
		if request.form("empresasc")<>""          then sql2=sql2 & ", '" & request.form("empresasc") & "' "
		if request.form("cnpjsc")<>""             then sql2=sql2 & ", '" & request.form("cnpjsc") & "' "
		if request.form("oe_valor")<>""           then sql2=sql2 & ", " &  nraccess(cdbl(request.form("oe_valor"))) & " "
		if request.form("oe_sc")<>""              then sql2=sql2 & ", " &  nraccess(cdbl(request.form("oe_sc"))) & " "
		if request.form("salcont")<>""            then sql2=sql2 & ", " &  nraccess(cdbl(request.form("salcont"))) & " "
		sql1 = sql & sql2 & ""
		'response.write "<font size='2'>" & sql1
		if tudook=1 then conexao.Execute sql1, , adCmdText
		'response.write "<br>" & request.form
	end if

else 'request.form=""
end if
%>
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function chapa1() {	form.chapa.value=form.nome2.value;	}
function chapa2() {	form.nome2.value=form.chapa.value;	}
--></script>
<%
if request("codigo")="" then autonomo=request.form("id_autonomo") else autonomo=request("codigo")
sql="select nome_autonomo, tipo_prestacao, ccm from autonomo where id_autonomo=" & autonomo
rsc.Open sql, ,adOpenStatic, adLockReadOnly
nome_autonomo=rsc("nome_autonomo")
if descricao_servico="" then descricao_servico=rsc("tipo_prestacao")
inscrito=rsc("ccm")
rsc.close
if request.form("proporcionaliza")="ON" then proporcionaliza="checked" else proporcionaliza=""
%>
<form method="POST" action="rpa_nova.asp" name="form" >
<input type="hidden" name="id_autonomo" value="<%=autonomo%>">
<input type="hidden" name="apuracao_darf" value="<%=apuracao_darf%>">
<input type="hidden" name="venc_darf" value="<%=venc_darf%>">
<input type="hidden" name="inscrito" value="<%=inscrito%>">
<table border="0" cellpadding="3" cellspacing="0" width="500">
    <tr><td class=grupo>Inclusão de Pagamento RPA - <%=request("codigo")%> - <%=nome_autonomo%> </td></tr>
  </table>

<table border="0" cellpadding="3" cellspacing="0" width="500">
<tr>
	<td class=titulo>Emissão</td>
	<td class=titulo>Pagamento</td>
	<td class=titulo>Descrição dos Serviços</td>
</tr>
<tr>
	<td class=titulo valign=top><input type="text" name="data_emissao" size="8" value="<%=data_emissao%>" onchange="javascript:submit()" ></td>
	<td class=titulo valign=top><input type="text" name="data_pagamento" size="8" value="<%=data_pagamento%>" onchange="javascript:submit()"></td>
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
		<td class=fundo width="30%"><input type="text" name="servico_prestado" size="8" value="<%=servico_prestado%>" class=vr onchange="javascript:submit()"></td>
	</tr>
	<tr>
		<td class=fundor>Outros Rendimentos</td>
		<td class=fundo>&nbsp;</td>
	</tr>
	<tr>
		<td class=fundo><input type="text" name="desc_outros_rend" size="25" value=""></td>
		<td class=fundo><input type="text" name="outros_rendimentos" size="8" value="<%=outros_rendimentos%>" class=vr onchange="javascript:submit()"></td>
	</tr>
	<tr>
		<td class=fundo style="border-bottom: 1px solid #000000">&nbsp;</td>
		<td class=fundo style="border-bottom: 1px solid #000000">&nbsp;</td>
	</tr>
	<tr>
		<td class=fundo align="right">Total Rendimentos&nbsp;</td>
		<td class=fundo><input type="text" name="total_rend"  size="8" value="<%=total_rend%>" class=bloq onFocus="total_rend.blur()"></td>
	</tr>
	<tr>
		<td class=fundo>&nbsp;</td>
		<td class=fundo>&nbsp;</td>
	</tr>
	<tr>
		<td class=fundo colspan=2>S.C. Fieo <input type="text" name="salcont" size=7 value="<%=salcont%>" class="vr" onchange="javascript:submit()">
		Proporcionaliza? <input type="checkbox" name="proporcionaliza" value="ON" <%=proporcionaliza%>>
                   
		</td>
	</tr>
</table>	
<!-- quadro rendimentos fim-->				
		</td>
		<td class=titulo colspan="2" valign=top>
<!-- quadro descontos -->		
	<table border="0" cellpadding="0" cellspacing="0" width=246>
	<tr>
		<td class=fundo width="70%">Imp.Renda Fonte</td>
		<td class=fundo width="30%"><input type="text" name="desconto_ir" size="8" value="<%=desconto_ir%>" class="vr" onchange="javascript:submit()"></td>
	</tr>
	<tr>
		<td class=fundo>ISS - Fonte</td>
		<td class=fundo><input type="text" name="desconto_iss" size="8" value="<%=desconto_iss%>" class="vr" onchange="javascript:submit()"></td>
	</tr>
	<tr>
		<td class=fundo>INSS</td>
		<td class=fundo><input type="text" name="desconto_inss" size="8" value="<%=desconto_inss%>" class="vr" onchange="javascript:submit()"></td>
	</tr>
	<tr>
		<td class=fundor>Outros Descontos</td>
		<td class=fundo>&nbsp;</td>
	</tr>
	<tr>
		<td class=fundo><input type="text" name="descricao_outros" size="25" value="<%=request.form("descricao_outros")%>"></td>
		<td class=fundo><input type="text" name="outros_descontos" size="8" value="<%=outros_descontos%>" class="vr" onchange="javascript:submit()"></td>
	</tr>
	<tr>
		<td class=fundo style="border-bottom: 1px solid #000000">&nbsp;</td>
		<td class=fundo style="border-bottom: 1px solid #000000">&nbsp;</td>
	</tr>
	<tr>
		<td class=fundo align="right">Total Descontos&nbsp;</td>
		<td class=fundo><input type="text" name="total_desc"  size="8" value="<%=total_desc%>" class=bloq onFocus="total_desc.blur()"></td>
	</tr>
</table>
<!-- quadro descontos fim-->				
		</td>
	</tr>
	<tr>
		<td class=fundo width="50%" colspan="2"><input type="checkbox" name="forcar_valores" value="0">&nbsp;Forçar Valores</td>
		<td class=titulo width="35%">Valor Líquido</td>
		<td class=fundo width="15%"><input type="text" name="valor_liquido" size="8" value="<%=valor_liquido%>" class="bloq" onFocus="valor_liquido.blur()"></td>
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
	<td class=fundo><input type="text" name="empresasc" size=30></td>
	<td class=fundo><input type="text" name="cnpjsc"    size=15></td>
	<td class=fundo><input type="text" name="oe_valor"   size=7 value="<%=oe_valor%>" class="vr" onchange="javascript:submit()">
	<td class=fundo><input type="text" name="oe_sc"   size=7 value="<%=oe_sc%>" class="vr" onchange="javascript:submit()">
	<td class=fundo><input type="text" name="inss_outra_empresa" size="6" value="<%=inss_outra_empresa%>" class="vr" onchange="javascript:submit()"></td>
</tr>
</table>
  
<!-- fim tabela -->  
  <table border="0" cellpadding="3" cellspacing="0" width="500">
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
'else
'rs.close
set rs=nothing
'end if

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