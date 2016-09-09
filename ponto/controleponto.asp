<%@ Language=VBScript %>
<!-- #Include file="..\adovbs.inc" -->
<!-- #Include file="..\funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a46")="N" or session("a46")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Ponto e Presença</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, rs, rsc
set conexao=server.createobject ("ADODB.Connection")
set rs=server.createobject ("ADODB.Recordset")
set rsc=server.createobject ("ADODB.Recordset")
conexao.Open application("conexao")

%>
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function nome1() {
	form.chapa.value=form.nome.value;
	}
function chapa1() {
	form.nome.value=form.chapa.value;
	}
--></script>
<%
if request.form="" then
%>
<p class=titulo>Emissão de Controle de Ponto e Presença</p>
<form method="POST" action="controleponto.asp" name="form">
  <input type="text" name="chapa" onchange="chapa1()" size="8" class=a>
  <select size="1" name="nome" onchange="nome1()">
  <option value="99">Todos</option>
  <option value="01">Campus Narciso</option>
  <option value="03">Campus Vila Yara</option>
  <option value="04">Campus Jardim Wilson</option>
<%
sql0="select chapa, nome from ifuncionarios where codsituacao<>'D' order by nome "
rsc.CursorLocation = adUseClient
rsc.Open sql0, conexao ,adOpenStatic, adLockReadOnly
Set rsc.ActiveConnection=nothing
rsc.movefirst:do while not rsc.eof
if request.form("nome")=rsc("chapa") then tempz="selected" else tempz=""
%>
          <option value="<%=rsc("chapa")%>" <%=tempz%>><%=rsc("nome")%></option>
<%
rsc.movenext
loop
rsc.close
mes=month(now())+1
ano=year(now())
if mes=13 then mes=1:ano=ano+1
%>
        </select>
  <br><font color="blue">Mês: <input type="text" name="mes" value="<%=mes%>" size="2"> Ano: <input type="text" name="ano" value="<%=ano%>" size="4">
  
  <br><input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>

<%
else ' visualizar
	codigo=request.form("chapa")
	mes=request.form("mes")
	ano=request.form("ano")
	select case codigo
		case "99"
			sqlfinal=" and codsindicato<>'03' and codtipo='N' ":controle="T:"
		case "01"
			sqlfinal=" and left(codsecao,2)='01' and codsindicato<>'03' and codtipo='N' ":controle="NS:"
		case "03"
			sqlfinal=" and left(codsecao,2)='03' and codsindicato<>'03' and codtipo='N' ":controle="VY:"
		case "04"
			sqlfinal=" and left(codsecao,2)='04' and codsindicato<>'03' and codtipo='N' ":controle="JW:"
		case else
			sqlfinal=" and (chapa='" & codigo & "' or codsecao='" & codigo & "') ":controle="S:"
	end select

	ultimo=day(dateserial(ano,mes+1,1)-1)
	dim feriado(31)
	for g=1 to ultimo
		feriado(g)=""
		sqlf="select nome, diaferiado from corporerm.dbo.gferiado where diaferiado='" & ano & numzero(mes,2) & numzero(g,2) & "' "
		rs.CursorLocation = adUseClient
		rs.Open sqlf, conexao ,adOpenStatic, adLockReadOnly
		if rs.recordcount>0 then feriado(g)=rs("nome")
		rs.close
	next 
		
	sql1="select f.chapa, f.nome, ctps=p.carteiratrab, serie=p.seriecarttrab, funcao=c.nome, f.codsecao, setor=s.descricao, s.cgc, s.rua, s.numero, horario=h.descricao " & _
	"from corporerm.dbo.pfunc f inner join corporerm.dbo.ppessoa p on p.codigo=f.codpessoa " & _
	"inner join corporerm.dbo.pfuncao c on c.codigo=f.codfuncao inner join corporerm.dbo.psecao s on s.codigo=f.codsecao " & _
	"inner join corporerm.dbo.ahorario h on h.codigo=f.codhorario " & _
	"where codsituacao<>'D' " & sqlfinal & " order by codsecao, f.nome "
	rs.CursorLocation = adUseClient
	rs.Open sql1, conexao ,adOpenStatic, adLockReadOnly
	Set rs.ActiveConnection=nothing
	do while not rs.eof
%>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr>
	<td class="campop"><b>CONTROLE DE PONTO E PRESENÇA</b></td>
	<td class=campo width=30></td>
	<td class="campop" align="right"><b><%=numzero(mes,2)%>/<%=ano%></b></td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr>
	<td class="campor" style="border-top:1px solid;border-right:1px solid">Empresa:</td>
	<td class="campor" style="border-top:1px solid;border-right:1px solid">CNPJ:</td>
	<td class="campor" style="border-top:1px solid;border-right:1px solid">Endereço:</td>	
	<td class="campor" style="border-top:1px solid;border-right:0px solid">Departamento</td>	
</tr>
<tr>
	<td class="campor" style="border-bottom:1px solid;border-right:1px solid">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td>
	<td class="campor" style="border-bottom:1px solid;border-right:1px solid"><%=rs("cgc")%></td>
	<td class="campor" style="border-bottom:1px solid;border-right:1px solid"><%=rs("rua") & " " & rs("numero")%></td>
	<td class="campor" style="border-bottom:1px solid;border-right:0px solid"><%=rs("setor")%></td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr>
	<td class="campor" style="border-top:0px solid;border-right:1px solid">Chapa</td>
	<td class="campor" style="border-top:0px solid;border-right:1px solid">Nome do Funcionário</td>
	<td class="campor" style="border-top:0px solid;border-right:1px solid">CTPS/Série</td>	
	<td class="campor" style="border-top:0px solid;border-right:0px solid">Função</td>
</tr>
<tr>
	<td class="campor" style="border-bottom:1px solid;border-right:1px solid"><b><%=rs("chapa")%></b></td>
	<td class="campor" style="border-bottom:1px solid;border-right:1px solid"><b><%=rs("nome")%></b></td>
	<td class="campor" style="border-bottom:1px solid;border-right:1px solid"><%=rs("ctps")&"/"&rs("serie")%></td>
	<td class="campor" style="border-bottom:1px solid;border-right:0px solid"><%=rs("funcao")%></td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr>
	<td class="campor" style="border-bottom:1px solid;border-right:0px solid">Horário: <%=rs("horario")%></td>	
</tr>
<tr><td class="campor" height=5></td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr>
	<td class=fundor rowspan=2 valign=middle align="center">Dia</td>
	<td class=fundor rowspan=2 valign=middle align="center">Entr.</td>
	<td class=fundor rowspan=2 valign=middle align="center">Saida</td>
	<td class=fundor rowspan=2 valign=middle align="center">Entr.</td>
	<td class=fundor rowspan=2 valign=middle align="center">Saida</td>
	<td class=fundor rowspan=2 valign=middle align="center">Entr.</td>
	<td class=fundor rowspan=2 valign=middle align="center">Saida</td>
	<td class=fundor colspan=4 align="center" height=5>Reservado (não preencha)</td>
</tr>
<tr>
	<td class="campor" height=10></td>
	<td class="campor"></td>
	<td class="campor"></td>
	<td class="campor"></td>
</tr>
<%
for d=1 to ultimo
data=dateserial(ano,mes,d)
diasem=weekday(data)
estilo="style='background:#ffffff'"
altura=25
if diasem=7 then '7 sabado
	estilo="style='background:#e1e1e1;font-size:7pt'":altura=20
end if
if diasem=1 then '1 domingo
	estilo="style='background:#9a9a9a;font-size:7pt'":altura=17
end if
if feriado(d)<>"" then 'feriado
	estilo="style='background:#aaaaaa;font-size:7pt'":altura=17
end if
'if d=6 or d=13 or d=20 or d=27 then estilo="fundo" else estilo="campo"
largura=65
%>
<tr>
	<td class=campo <%=estilo%> align="center" height=<%=altura%> width=45 nowrap><%=d & " (" & weekdayname(weekday(data),1) & ")"%></td>
<%if feriado(d)<>"" then %>
	<td class=campo colspan=4 <%=estilo%> align="center" ><%=feriado(d)%></td>
<%else%>
	<td class=campo <%=estilo%> width="<%=largura%>">&nbsp;</td>
	<td class=campo <%=estilo%> width="<%=largura%>">&nbsp;</td>
	<td class=campo <%=estilo%> width="<%=largura%>">&nbsp;</td>
	<td class=campo <%=estilo%> width="<%=largura%>">&nbsp;</td>
<%end if%>	
	<td class=campo <%=estilo%> width="<%=largura%>">&nbsp;</td>
	<td class=campo <%=estilo%> width="<%=largura%>">&nbsp;</td>
	<td class=campo <%=estilo%> >&nbsp;</td>
	<td class=campo <%=estilo%> >&nbsp;</td>
	<td class=campo <%=estilo%> >&nbsp;</td>
	<td class=campo <%=estilo%> >&nbsp;</td>
</tr>
<%
next
%>
</table>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td colspan=2 class="campor" height=5></td></tr>
<tr>
	<td colspan=1 class="campor" style="border-top:1px solid;border-bottom:1px solid;border-left:1px solid" width=500><u>Observações:</u>
	<br>-Embora a entrega seja mensal, o preenchimento deverá ser feito diariamente.
	<br>-Nas colunas destinadas ao horário, não preencha com outras informações além do seu horário de entrada ou saída.
	<br>-Nas colunas reservadas não faça nenhuma anotação.
	<br>-Caso tenha algum atestado médico ou outro comprovante de ausência, não anote, entregue no RH em até 48 horas.
	<br>-A não veracidade nas informações ensejará na aplicação das penalidades previstas na C.L.T.
	</td>
	<td class="campop" valign=middle align="center" style="border-top:1px solid;border-bottom:1px solid;border-right:1px solid"><b>Não Rasure</b></td>
</tr>
</table>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=690>
<tr>
	<td class=campo valign=top>Reconheço a exatidão e confirmo a frequência anotada.
	<br>
	<br>
	<br>______________________________________
	<br><%=rs("nome")%>
	</td>
	<td class=campo valign=top>Confirmo e autentico.
	<br>
	<br>
	<br>______________________________________
	<br>Assinatura da chefia
	<%for a=1 to 35:response.write "&nbsp;":next%><%=controle & rs.absoluteposition & "/" & rs.recordcount%>
	</td>
</tr>
</table>

<%
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>"
rs.movenext
loop
rs.close

%>


<%
end if ' request.form<>""
%>
</body>
</html>
<%
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%>