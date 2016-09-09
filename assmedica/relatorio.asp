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
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form<>"" then
	pr1=request.form("prorata1")
	pr2=request.form("prorata2")
	empresa=request.form("empresa")
	datarel=request.form("data_base")
	if request.form("resumo")="ON" then resumo=1 else resumo=0
	if request.form("totalf")="ON" then totalf=1 else totalf=0
	if request.form("reajuste")="ON" then reajuste=1 else reajuste=0
	if request.form("anterior")="ON" then anterior=1 else anterior=0
	if request.form("contrato")<>"" then up=request.form("contrato") else up=""
	sessao=session.sessionid:sessao=session("usuariomaster")
	conexao.execute "delete from ttassmedrel where sessao='" & sessao & "' "
	sql="select operadora from assmed_empresa where codigo='" & empresa & "' "
	rs.open sql
	operadora=rs("operadora")
	rs.close
	ordem=request.form("ordem")
string1="[valor]":string2="[reembolso]"
if reajuste=1 then string1="[dif_vr]"
if reajuste=1 then string2="[dif_re]"
if anterior=1 then string1="[valor_ant]"
if anterior=1 then string2="[reembolso_ant]"
'atualizando up dependentes
sql="update assmed_dep_mudanca set up = m.up " & _
"FROM assmed_mudanca m INNER JOIN assmed_dep_mudanca dm ON m.chapa=dm.chapa AND m.empresa=dm.empresa AND m.plano=dm.plano;"
'response.write "<br>" & sql
conexao.execute sql

'inserindo titulares
sql="INSERT INTO ttassmedrel (sessao, data_base, tp, chapa, principal, beneficiario, codsecao, " & _
"empresa, plano, codigo, up, sexo, ingresso, validade, valor_plano, reemb, " & _
"rt1, rt2, rt3, rt4, rt5, rt6, rt7, rt8, rt9, rt10, rt11, rt12, " & _
"rd1, rd2, rd3, rd4, rd5, rd6, rd7, rd8, rd9, rd10, rd11, rd12, dtnascimento ) " & _
"SELECT '" & sessao & "', '" & dtaccess(datarel) & "', 'Titular', ab.chapa, f.nome, f.nome, " & _
"f.codsecao, am.empresa, am.plano, am.codigo, am.up, p.sexo, am.ivigencia, am.fvigencia, " & _
"case when " & pr1 & " <> " & pr2 & " then round(((" & string1 & " / " & pr2 & " ) * " & pr1 & ")*100+0.5,2)/100 else " & string1 & " end, " & _
"case when " & pr1 & " <> " & pr2 & " then round(((" & string2 & " / " & pr2 & " ) * " & pr1 & ")*100+0.5,2)/100 else " & string2 & " end, " & _
"case when seq=1 then 1 else 0 end, case when seq=2 then 1 else 0 end, case when seq=3 then 1 else 0 end, case when seq=4 then 1 else 0 end, " & _
"case when seq=5 then 1 else 0 end, case when seq=6 then 1 else 0 end, case when seq=7 then 1 else 0 end, case when seq=8 then 1 else 0 end, " & _
"case when seq=9 then 1 else 0 end, case when seq=10 then 1 else 0 end, case when seq=11 then 1 else 0 end, case when seq=12 then 1 else 0 end, " & _
"0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, p.dtnascimento " & _
"FROM assmed_beneficiario ab inner join corporerm.dbo.pfunc f on ab.chapa=f.chapa collate database_default " & _
"inner join assmed_mudanca am on ab.CHAPA=am.chapa " & _
"inner join assmed_planos ap on am.plano=ap.plano AND am.empresa=ap.codigo " & _
"inner join corporerm.dbo.ppessoa p on f.codpessoa=p.codigo " & _
"WHERE am.empresa='" & empresa & "' AND am.ivigencia<='" & dtaccess(datarel) & "' AND '" & dtaccess(datarel) & "' Between [ivigencia] And [fvigencia] "
'response.write "<br>" & sql
conexao.execute sql

'inserindo dependentes
sql="INSERT INTO ttassmedrel (sessao, data_base, tp, chapa, principal, beneficiario, codsecao, " & _
"empresa, plano, codigo, up, sexo, ingresso, validade, valor_plano, reemb, " & _
"rt1, rt2, rt3, rt4, rt5, rt6, rt7, rt8, rt9, rt10, rt11, rt12, " & _
"rd1, rd2, rd3, rd4, rd5, rd6, rd7, rd8, rd9, rd10, rd11, rd12, parentesco, dtnascimento ) " & _
"SELECT '" & sessao & "', '" & dtaccess(datarel) & "', 'Dependente', ab.chapa, f.NOME, ad.dependente, " & _
"f.codsecao, adm.empresa, adm.plano, adm.codigo, adm.up, ad.sexo, adm.ivigencia, adm.fvigencia, " & _
"case when " & pr1 & " <> " & pr2 & " then round(((" & string1 & " / " & pr2 & " ) * " & pr1 & ")*100+0.5,2)/100 else " & string1 & " end, " & _
"case when " & pr1 & " <> " & pr2 & " then /**/ round(((case when adm.empresa='S' then " & string1 & " else case when (adm.empresa='U' or adm.empresa='BS' or adm.empresa='C') and ((parentesco='Conjuge' or parentesco='Companheiro(a)') and sexo='M') and adm.plano not like 'SÊNIOR%' then " & string1 & " else " & string2 & " end end)/" & pr2 & ")*" & pr1 & ",2) " & _
"else /**/ case when adm.empresa='S' then " & string1 & " else case when (adm.empresa='U' or adm.empresa='BS' or adm.empresa='C') and ((parentesco='Conjuge' or parentesco='Companheiro(a)') and sexo='M') and adm.plano not like 'SÊNIOR%' then " & string1 & " else " & string2 & " end end end, " & _
"0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " & _
"case when seq=1 then 1 else 0 end, case when seq=2 then 1 else 0 end, case when seq=3 then 1 else 0 end, case when seq=4 then 1 else 0 end, " & _
"case when seq=5 then 1 else 0 end, case when seq=6 then 1 else 0 end, case when seq=7 then 1 else 0 end, case when seq=8 then 1 else 0 end, " & _
"case when seq=9 then 1 else 0 end, case when seq=10 then 1 else 0 end, case when seq=11 then 1 else 0 end, case when seq=12 then 1 else 0 end, " & _
"ad.parentesco, ad.nascimento " & _
"FROM assmed_beneficiario ab inner join corporerm.dbo.pfunc f on ab.chapa=f.chapa collate database_default " & _
"inner join assmed_dep ad on ab.CHAPA=ad.chapa " & _
"inner join assmed_dep_mudanca adm on ad.chapa=adm.chapa and ad.nrodepend=adm.nrodepend " & _
"inner join assmed_planos ap on adm.plano=ap.plano AND adm.empresa=ap.codigo " & _
"WHERE adm.empresa='" & empresa & "' AND adm.ivigencia<='" & dtaccess(datarel) & "' AND '" & dtaccess(datarel) & "' Between [ivigencia] And [fvigencia] "
'response.write "<br>" & sql
conexao.execute sql

'inserindo acertos
if reajuste=1 then string1="aa.dif_vr, aa.dif_re" else string1="aa.valor_acerto, aa.reembolso"
sql="INSERT INTO ttassmedrel (sessao, data_base, tp, chapa, principal, beneficiario, codsecao, " & _
"empresa, plano, codigo, ingresso, validade, valor_plano, reemb, up, " & _
"rt1, rt2, rt3, rt4, rt5, rt6, rt7, rt8, rt9, rt10, rt11, rt12, " & _
"rd1, rd2, rd3, rd4, rd5, rd6, rd7, rd8, rd9, rd10, rd11, rd12 ) " & _
"SELECT '" & sessao & "', '" & dtaccess(datarel) & "', 'Acerto', aa.chapa, f.nome, f.nome, " & _
"f.codsecao, aa.empresa, 'Ajuste Valor', codigo='', '" & dtaccess(datarel) & "','" & dtaccess(datarel) & "', " & string1 & ", " & _
" up='', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 " & _
"FROM corporerm.dbo.pfunc f, assmed_acertos aa " & _
"WHERE f.chapa collate database_default=aa.chapa " & _
"AND convert(char,year(data_acerto))+'/'+convert(char,month(data_acerto))+'/01'=convert(char,Year('" & dtaccess(datarel) & "'))+'/'+convert(char,Month('" & dtaccess(datarel) & "'))+'/01' " & _
"and aa.empresa='" & empresa & "' "
'response.write "<br>" & sql
conexao.execute sql

sql="SELECT sessao, data_base, chapa, tp, principal, beneficiario, codsecao, empresa, plano, codigo, up, sexo, " & _
"ingresso, validade, valor_plano, reemb, rt1, rt2, rt3, rt4, rt5, rt6, rt7, rt8, rt9, rt10, rt11, rt12, " & _
"rd1, rd2, rd3, rd4, rd5, rd6, rd7, rd8, rd9, rd10, rd11, rd12, parentesco, dtnascimento " & _
"FROM ttassmedrel " & _
"WHERE sessao='" & sessao & "' " & _
"ORDER BY " & ordem & ", tp DESC , beneficiario "
'response.write "<br>" & sql

sqlz="update ttassmedrel set valor_plano=0 where valor_plano is null": conexao.execute sqlz
sqlz="update ttassmedrel set reemb=0 where reemb is null": conexao.execute sqlz
if up<>"" then conexao.execute "delete from ttassmedrel where sessao='" & sessao & "' and up<>" & up & " "
end if

if request.form="" then 
%>
<p class=titulo>Geração de relatório da Assistência Médica
<form method="POST" action="relatorio.asp">
  <p>Data base para emissão: <input type="text" name="data_base" size="12" class=a value="<%=formatdatetime(now(),2)%>"><br>
  Divisão pro-rata: <input type="text" name="prorata1" size="4" value="30" class=a>/<input type="text" name="prorata2" size="4" value="30"><br>
  Empresa de Saúde: <select size="1" name="empresa">
<%
	sqla="SELECT * from assmed_empresa ORDER by operadora"
	rsc.Open sqla, ,adOpenStatic, adLockReadOnly
	rsc.movefirst
	cempresa=rsc("codigo")
	do while not rsc.eof
	if rsc("codigo")="C" then tempt="selected" else tempt=""
%>
          <option value="<%=rsc("codigo")%>" <%=tempt%>><%=rsc("operadora")%></option>
<%
	rsc.movenext
	loop
	rsc.close
%>
        </select><br>
Somente Resumo <input type="checkbox" name="resumo" value="ON"><br>
Imprimir total por funcionário <input type="checkbox" name="totalf" value="ON"><br>
Diferença de Reajuste? <input type="checkbox" name="reajuste" value="ON"><br>
Preço Anterior? <input type="checkbox" name="anterior" value="ON"><br>
SubContrato: <input type="text" size="3" name="contrato" value=""><br>
Ordem: <select size="1" name="ordem">
	<option value="chapa">Chapa</option>
	<option value="principal">Nome</option>
	<option value="up,chapa">UP, Chapa</option>
	<option value="up,principal">UP, Nome</option>
	<option value="up,beneficiario">UP/Beneficiario</option>
</select><br>
</p>		
<p><input type="submit" value="Visualizar relatório" name="Gerar" class="button"></p>
</form>
<%
else
%>
<table border="0" cellpadding="2" width="690" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td align="left"  >Controle de Assistência Médica</td>
    <td align="center">Relação de Beneficiários e Dependentes</td>
    <td align="right" >Empresa: <%=operadora%></td>
  </tr>
</table>
<table border="0" cellpadding="0" width="690" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td class=titulor>Chapa   </td>
    <td class=titulor>Nome    </td>
    <td class=titulor>Nasc.</td>
    <td class=titulor>Sexo</td>
    <td class=titulor>Inclusão</td>
    <td class=titulor>Plano   </td>
    <td class=titulor>Compos. </td>
    <td class=titulor align="right">Valor &nbsp;  </td>
    <td class=titulor align="right">Reemb.&nbsp;  </td>
    <td class=titulor>Validade</td>
    <td class=titulor>Seção   </td>
  </tr>
<%

linha=2:limite=82 '72
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalam=0:totalrb=0:inicio=1
rt1=0:rt2=0:rt3=0:rt4=0:rt5=0:rt6=0:rt7=0:rt8=0:rt9=0:rt10=0:rt11=0
rd1=0:rd2=0:rd3=0:rd4=0:rd5=0:rd6=0:rd7=0:rd8=0:rd9=0:rd10=0:rd11=0
rs.movefirst
do while not rs.eof
if resumo=0 then
if linha>limite then
	pagina=pagina+1
	response.write "</table>"
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<table border='0' cellpadding='2' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  >Controle de Assistência Médica</td>"
	response.write "<td align='center'>Relação de Beneficiários e Dependentes</td>"
	response.write "<td align='right' >Empresa: " & operadora & "</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table border='0' cellpadding='0' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=titulor>Chapa   </td>"
	response.write "<td class=titulor>Nome    </td>"
	response.write "<td class=titulor>Nasc.</td>"
	response.write "<td class=titulor>Sexo</td>"
	response.write "<td class=titulor>Inclusão</td>"
	response.write "<td class=titulor>Plano   </td>"
	response.write "<td class=titulor>Compos. </td>"
	response.write "<td class=titulor>Valor   </td>"
	response.write "<td class=titulor>Reemb.  </td>"
	response.write "<td class=titulor>Validade</td>"
	response.write "<td class=titulor>Seção   </td>"
	response.write "</tr>"
	linha=2
end if
end if 'resumo
totalam=totalam+cdbl(rs("valor_plano"))
totalrb=totalrb+cdbl(rs("reemb"))
rt1=rt1+rs("rt1"):rt2=rt2+rs("rt2"):rt3=rt3+rs("rt3"):rt4=rt4+rs("rt4"):rt5=rt5+rs("rt5")
rt6=rt6+rs("rt6"):rt7=rt7+rs("rt7"):rt8=rt8+rs("rt8"):rt9=rt9+rs("rt9"):rt10=rt10+rs("rt10")
rt11=rt11+rs("rt11"):rt12=rt12+rs("rt12"):
rd1=rd1+rs("rd1"):rd2=rd2+rs("rd2"):rd3=rd3+rs("rd3"):rd4=rd4+rs("rd4"):rd5=rd5+rs("rd5")
rd6=rd6+rs("rd6"):rd7=rd7+rs("rd7"):rd8=rd8+rs("rd8"):rd9=rd9+rs("rd9"):rd10=rd10+rs("rd10")
rd11=rd11+rs("rd11"):rd12=rd12+rs("rd12")
if rs("tp")<>"Dependente" and ordem<>"up,beneficiario" then estilo="style='border-top: 1px solid #000000'" else estilo=""
if datevalue(rs("validade"))="31/12/2020" or datevalue(rs("validade"))="30/09/03" then validade="&nbsp;" else validade=rs("validade")
if rs("reemb")=0 then reembolso="&nbsp;" else reembolso=formatnumber(rs("reemb"))
if rs("tp")<>"Dependente" then beneficiario="<b>" & rs("beneficiario") & "</b>" else beneficiario=rs("beneficiario")
if rs("tp")="Acerto" then totalacerto=totalacerto+rs("valor_plano")
	if ultchapa<>rs("chapa") and totalf=1 and inicio=0 then
%>
 <tr>
	<td class="campor" style="border-top: 1 dotted #000000" colspan=7>&nbsp;</td>
	<td class="campor" style="border-top: 1 dotted #000000" align="right"><i><%=formatnumber(totrel1,2)%></td>
	<td class="campor" style="border-top: 1 dotted #000000" align="right"><i><%=formatnumber(totrel2,2)%></td>
	<td class="campor" style="border-top: 1 dotted #000000" colspan=2>&nbsp;</td>
 </tr>
<%
	linha=linha+1
	end if
	if ultchapa<>rs("chapa") and totalf=1 then
		totrel1=0:totrel2=0
	end if
if resumo=0 then
%>
  <tr>
    <td class="campor" <%=estilo%>><%=rs("chapa")%></td>
    <td class="campor" <%=estilo%>><%=beneficiario %></td>
    <td class="campor" align="right" <%=estilo%>><%=rs("dtnascimento")%></td>
    <td class="campor" align="center" <%=estilo%>><%=rs("sexo")%></td>
    <td class="campor" align="center" <%=estilo%>><%=rs("ingresso")%></td>
    <td class="campor" <%=estilo%>><%=left(rs("plano"),14)%></td>
    <td class="campor" <%=estilo%>><%=rs("parentesco")%></td>
    <td class="campor" <%=estilo%> align="right"><%=formatnumber(rs("valor_plano"),2)%></td>
    <td class="campor" <%=estilo%> align="right"><%=reembolso %></td>
    <td class="campor" align="right" <%=estilo%>><%=validade %></td>
    <td class="campor" <%=estilo%>><%=rs("codsecao")%></td>
  </tr>
<%
linha=linha+1:inicio=0
totrel1=totrel1+rs("valor_plano")
totrel2=totrel2+rs("reemb")
ultchapa=rs("chapa")
end if 'resumo
rs.movenext
loop
rs.close

	if totalf=1 and inicio=0 then
%>
 <tr>
	<td class="campor" style="border-top: 1 dotted #000000" colspan=7>&nbsp;</td>
	<td class="campor" style="border-top: 1 dotted #000000" align="right"><i><%=formatnumber(totrel1,2)%></td>
	<td class="campor" style="border-top: 1 dotted #000000" align="right"><i><%=formatnumber(totrel2,2)%></td>
	<td class="campor" style="border-top: 1 dotted #000000" colspan=2>&nbsp;</td>
 </tr>
<%
	linha=linha+1
	end if
%>
  <tr>
    <td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000" align="right"><%=formatnumber(totalam,2)%></td>
    <td class="campor" style="border-top: 1px solid #000000" align="right"><%=formatnumber(totalrb,2)%></td>
    <td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
    <td class="campor" style="border-top: 1px solid #000000">&nbsp;</td>
  </tr>
</table>
<%
linha=linha+1

if linha>limite-9 then
	pagina=pagina+1
	'response.write "<br>"
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<table border='0' cellpadding='2' width='690' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td align='left'  >Controle de Assistência Médica        </td>"
	response.write "<td align='center'>Relação de Beneficiários e Dependentes</td>"
	response.write "<td align='right' >Empresa: " & operadora & "</td>"
	response.write "</tr>"
	response.write "</table>"
	linha=1
end if
rt(1)=rt1:rt(2)=rt2:rt(3)=rt3:rt(4)=rt4:rt(5)=rt5:rt(6)=rt6:rt(7)=rt7:rt(8)=rt8:rt(9)=rt9:rt(10)=rt10:rt(11)=rt11:rt(12)=rt12
rd(1)=rd1:rd(2)=rd2:rd(3)=rd3:rd(4)=rd4:rd(5)=rd5:rd(6)=rd6:rd(7)=rd7:rd(8)=rd8:rd(9)=rd9:rd(10)=rd10:rd(11)=rd11:rd(12)=rd12
sql="SELECT Max(seq) AS fim, Min(seq) AS inicio FROM assmed_planos " & _
"where codigo='" & empresa & "' "
rsc.open sql
inicio=rsc("inicio")
fim=rsc("fim")
rsc.close
%>
<p style="margin-top: 0; margin-bottom: 0"><font size="1"><b>Resumo de Vidas por Plano</b></font></p>
<%linha=linha+1%>
<table border="1" cellspacing="0" style="border-collapse: collapse" width="400">
  <tr>
    <td class=titulor bgcolor="#CCFFCC" style="border-bottom-style: solid; border-bottom-width: 1" align="center">Resumo</td>
    <td class=titulor bgcolor="#CCFFCC" style="border-bottom-style: solid; border-bottom-width: 1" align="center">Titulares</td>
    <td class=titulor bgcolor="#CCFFCC" style="border-bottom-style: solid; border-bottom-width: 1" align="center">Dependentes</td>
    <td class=titulor bgcolor="#CCFFCC" style="border-bottom-style: solid; border-bottom-width: 1" align="center">Total Vidas</td>
    <td class=titulor bgcolor="#CCFFCC" style="border-bottom-style: solid; border-bottom-width: 1" align="center">Valor Plano</td>
    <td class=titulor bgcolor="#CCFFCC" style="border-bottom-style: solid; border-bottom-width: 1" align="center">Totais</td>
  </tr>
<%
string1="[valor]":string2="[reembolso]"
if anterior=1 then string1="[valor_ant] as valor"
if anterior=1 then string2="[reembolso_ant]"
if reajuste=1 then string1="[dif_vr] as valor"
if reajuste=1 then string2="[dif_re]"
linha=linha+1
'fim=10
for a=inicio to fim
linha=linha+1
sql="select plano, " & string1 & " from assmed_planos where codigo='" & empresa & "' and seq=" & a
rsc.open sql, ,adOpenStatic, adLockReadOnly
if rsc.recordcount=1 then
'response.write "<Br>" & a & "-" & sql & "->" & rsc.recordcount
somatit=somatit+rt(a)
somadep=somadep+rd(a)
somag=somatit+somadep
if pr1 <> pr2 then
	valorpl=int(((rsc("valor")/pr2)*pr1)*100+0.5)/100
else
	valorpl=rsc("valor")
end if
tpp=valorpl *(rt(a)+rd(a))
tgp=tgp+tpp
%>
  <tr>
    <td class="campor">&nbsp;<%=rsc("plano")%></td>
    <td class="campor" align="right">&nbsp;<%=rt(a)%>&nbsp;</td>
    <td class="campor" align="right">&nbsp;<%=rd(a)%>&nbsp;</td>
    <td class="campor" align="right">&nbsp;<%=rt(a)+rd(a)%>&nbsp;</td>
    <td class="campor" align="right">&nbsp;<%=formatnumber(valorpl,2)%>&nbsp;</td>
    <td class="campor" align="right">&nbsp;<%=formatnumber(valorpl*(rt(a)+rd(a)),2)%>&nbsp;</td>
  </tr>
<%
end if 'rsc.recordcount
rsc.close
next
%>
  <tr>
    <td class=titulor style="border-top-style: solid; border-top-width: 1">&nbsp;Totais</td>
    <td class=titulor align="right" style="border-top-style: solid; border-top-width: 1">&nbsp;<%=somatit %></td>
    <td class=titulor align="right" style="border-top-style: solid; border-top-width: 1">&nbsp;<%=somadep %></td>
    <td class=titulor align="right" style="border-top-style: solid; border-top-width: 1">&nbsp;<%=somag %></td>
    <td class=titulor align="right" style="border-top-style: solid; border-top-width: 1">&nbsp;</td>
    <td class=titulor align="right" style="border-top-style: solid; border-top-width: 1">&nbsp;<%=formatnumber(tgp,2)%></td>
  </tr>
</table>
<%
linha=linha+1
%>
<p style="margin-top: 0; margin-bottom: 0"><font size="1">Total de Cobrança/Desconto Retroativo: R$ <%=formatnumber(totalacerto,2) %></font></p>
<%
linha=linha+1
pagina=pagina+1
response.write "<br>"
response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"

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