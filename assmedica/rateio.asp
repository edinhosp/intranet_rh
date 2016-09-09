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
<title>Rateio de Assistência Médica</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
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
	pr1=request.form("prorata1")
	pr2=request.form("prorata2")
	empresa=request.form("empresa")
	datarel=request.form("data_base")
	if request.form("reajuste")="ON" then reajuste=1 else reajuste=0
	if request.form("anterior")="ON" then anterior=1 else anterior=0
	sessao=session.sessionid
	conexao.execute "delete from ttassmedrel where sessao='" & sessao & "' "
	sql="select operadora from assmed_empresa where codigo='" & empresa & "' "
	rs.open sql
	operadora=rs("operadora")
	rs.close
string1="[valor]":string2="[reembolso]"
if reajuste=1 then string1="[dif_vr]"
if reajuste=1 then string2="[dif_re]"
if anterior=1 then string1="[valor_ant]"
if anterior=1 then string2="[reembolso_ant]"
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
"case when " & pr1 & " <> " & pr2 & " then /**/ round(((case when adm.empresa='S' then " & string1 & " else case when adm.empresa='M' and parentesco='Esposo' and adm.plano not like 'Diamante%' then " & string1 & " else " & string2 & " end end)/" & pr2 & ")*" & pr1 & ",2) " & _
"else /**/ case when adm.empresa='S' then " & string1 & " else case when adm.empresa='M' and parentesco='Esposo' and adm.plano not like 'Diamante%' then " & string1 & " else " & string2 & " end end end, " & _
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
"empresa, plano, codigo, ingresso, validade, valor_plano, reemb, " & _
"rt1, rt2, rt3, rt4, rt5, rt6, rt7, rt8, rt9, rt10, rt11, rt12, " & _
"rd1, rd2, rd3, rd4, rd5, rd6, rd7, rd8, rd9, rd10, rd11, rd12 ) " & _
"SELECT '" & sessao & "', '" & dtaccess(datarel) & "', 'Acerto', ab.chapa, f.nome, f.nome, " & _
"f.codsecao, ab.empresa, 'Ajuste Valor', ab.codigo, '" & dtaccess(datarel) & "','" & dtaccess(datarel) & "', " & string1 & ", " & _
"0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 " & _
"FROM assmed_beneficiario ab, corporerm.dbo.pfunc f, assmed_acertos aa " & _
"WHERE ab.chapa=f.chapa collate database_default and ab.chapa=aa.chapa " & _
"AND convert(char,year(data_acerto))+'/'+convert(char,month(data_acerto))+'/01'=convert(char,Year('" & dtaccess(datarel) & "'))+'/'+convert(char,Month('" & dtaccess(datarel) & "'))+'/01' " & _
"and aa.empresa='" & empresa & "' "
'response.write "<br>" & sql
conexao.execute sql

sql="SELECT sessao, data_base, chapa, tp, principal, beneficiario, codsecao, empresa, plano, codigo, " & _
"ingresso, validade, valor_plano, reemb, rt1, rt2, rt3, rt4, rt5, rt6, rt7, rt8, rt9, rt10, " & _
"rd1, rd2, rd3, rd4, rd5, rd6, rd7, rd8, rd9, rd10, parentesco " & _
"FROM ttassmedrel " & _
"WHERE sessao='" & sessao & "' " & _
"ORDER BY principal, chapa, tp DESC , beneficiario "
'response.write "<br>" & sql

sql="SELECT tt.codsecao, s.DESCRICAO, Sum(tt.valor_plano) AS Rateio, " & _
"conta=case when substring(codsecao,4,1)='1' then '634' else case when substring(codsecao,4,1)='2' then '519' else case when substring(codsecao,4,1)='3' then '406' else '' end end end " & _
"FROM ttassmedrel tt LEFT JOIN corporerm.dbo.PSECAO s ON tt.codsecao = s.CODIGO collate database_default " & _
"WHERE tt.sessao='" & sessao & "' " & _
"GROUP BY tt.codsecao, s.DESCRICAO, case when substring(codsecao,4,1)='1' then '634' else case when substring(codsecao,4,1)='2' then '519' else case when substring(codsecao,4,1)='3' then '406' else '' end end end " & _
"order by case when substring(codsecao,4,1)='1' then '634' else case when substring(codsecao,4,1)='2' then '519' else case when substring(codsecao,4,1)='3' then '406' else '' end end end, tt.codsecao "

end if

if request.form="" then
%>
<p class=titulo>Geração de rateio da Assistência Médica
<form method="POST" action="rateio.asp">
  <p>Data base para emissão: <input type="text" name="data_base" class=a size="12" value="<%=formatdatetime(now(),2)%>"><br>
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
Diferença de Reajuste? <input type="checkbox" name="reajuste" value="ON"><br>
Preço Anterior? <input type="checkbox" name="anterior" value="ON">
		</p>
  <p><input type="submit" value="Visualizar relatório" name="Gerar" class="button"></p>
</form>
<%
else
%>
<table border="0" cellpadding="2" width="650" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td class=grupo align="left"  >Controle de Assistência Médica</td>
    <td class=grupo align="center">Rateio de Nota Fiscal - Data-base: <%=datarel%></td>
    <td class=grupo align="right" >Empresa: <%=operadora%></td>
  </tr>
</table>
<table border="0" cellpadding="0" width="650" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td class=titulo>Conta    </td>
    <td class=titulo>Código   </td>
    <td class=titulo>Descrição</td>
    <td class=titulo align="center">Valor    </td>
  </tr>
<%
linha=2
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalam=0:totalrb=0
rs.movefirst
do while not rs.eof
if linha>70 then
	pagina=pagina+1
	response.write "</table>"
	'response.write "<br>"
	response.write "<p style='margin-top:0;margin-bottom:0'><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	response.write "<table border='0' cellpadding='2' width='650' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=grupo align='left'  >Controle de Assistência Médica</td>"
	response.write "<td class=grupo align='center'>Rateio de Nota Fiscal         </td>"
	response.write "<td class=grupo align='right' >Empresa: " & operadora &     "</td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "<table border='0' cellpadding='0' width='650' cellspacing='0' style='border-collapse: collapse'>"
	response.write "<tr>"
	response.write "<td class=titulo>Conta   </td>"
	response.write "<td class=titulo>Código   </td>"
	response.write "<td class=titulo>Descrição</td>"
	response.write "<td class=titulo align=""center"">Valor    </td>"
	response.write "</tr>"
	linha=2
end if
totalam=totalam+cdbl(rs("rateio"))
%>
  <tr>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1">&nbsp;<%=rs("conta")%> </td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1">&nbsp;<%=rs("codsecao")%> </td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1">&nbsp;<%=rs("descricao")%></td>
    <td class=campo style="border-bottom-style: solid; border-bottom-width: 1" align="right">&nbsp;<%=formatnumber(rs("rateio"),2)%>&nbsp;&nbsp;&nbsp;</td>
  </tr>
<%
linha=linha+1
rs.movenext
loop
rs.close
%>
  <tr>
    <td class=titulo style="border-top: 1px solid #000000">&nbsp;</td>
    <td class=titulo style="border-top: 1px solid #000000">&nbsp;</td>
    <td class=titulo style="border-top: 1px solid #000000">&nbsp;</td>
    <td class=titulo style="border-top: 1px solid #000000" align="right">&nbsp;<%=formatnumber(totalam,2)%></td>
  </tr>
</table>
<%
linha=linha+1
pagina=pagina+1
'response.write "<br>"
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