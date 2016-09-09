<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a42")="N" or session("a42")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Programação de Férias</title>
<script language="javascript" type="text/javascript">
<!--
/****************************************************
     Author: Eric King
     Url: http://redrival.com/eak/index.shtml
     This script is free to use as long as this info is left in
     Featured on Dynamic Drive script library (http://www.dynamicdrive.com)
****************************************************/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function secao1() { form.codsecao.value=form.secao.value;form.submit(); }
function codsecao1() { form.secao.value=form.codsecao.value;form.submit(); }
--></script>
<%
dim conexao, rs, rs2
dim mes(12)
mes(1)="Janeiro":mes(2)="Fevereiro":mes(3)="Março":mes(4)="Abril":mes(5)="Maio":mes(6)="Junho"
mes(7)="Julho":mes(8)="Agosto":mes(9)="Setembro":mes(10)="Outubro":mes(11)="Novembro":mes(12)="Dezembro"
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
sessao=session.sessionid:sessao=session("usuariomaster")

if request.form("B1")="" then
'---------- inicio e geração -----------------
sql1="delete from feriasprog where sessao='" & sessao & "'"
sql2="insert into feriasprog (sessao, chapa, codsecao, relsecao, dtvencferias, inicprogferias1, fimprogferias1, nrodiasferias, nrodiasabono) " & _
"select '" & sessao & "', chapa, codsecao, relsecao, dtvencferias, inicprogferias1, fimprogferias1, nrodiasferias, nrodiasabono " & _
"from ( " & _
"select f.chapa, f.codsecao, f.codsecao relsecao, 'dtvencferias'=p.FIMPERAQUIS, 'inicprogferias1'=p.DATAINICIO, 'fimprogferias1'=p.DATAFIM, p.nrodiasferias, p.nrodiasabono, 'faltas'=null " & _
"from corporerm.dbo.pfunc f inner join corporerm.dbo.pfuferiasper p on p.CHAPA=f.CHAPA " & _
"where f.codsindicato<>'03' and f.codsituacao<>'D' and p.SITUACAOFERIAS<>'F' and p.SITUACAOFERIAS='M' " & _
"union " & _
"select * from ( " & _
"select c.chapa, f.CODSECAO, relsecao=f.CODSECAO, c.FIMPERAQUIS, 'inicprogferias1'=null, 'fimprogferias1'=null,'nrodiasferias'=(case when c.CHAPA='00498' or c.CHAPA='01514' or c.CHAPA='00093' then 30 else case when faltas between 0 and 5 then 30 else case when faltas between 6 and 14 then 24 else case when faltas between 15 and 23 then 18 else case when faltas between 24 and 32 then 12 else 0 end end end end end) -dias, 'nrodiasabono'=null, faltas " & _
"from ( " & _
"select p.CHAPA, p.FIMPERAQUIS, p.PERIODOABERTO, dias=sum(case when m.NRODIASABONO IS null then 0 else m.NRODIASABONO end+case when m.NRODIASFERIAS IS null then 0 else m.NRODIASFERIAS end), " & _
"faltas=(select COUNT(datareferencia) from corporerm.dbo.AFALTAPARAFOLHA where CHAPA=p.CHAPA and DATAREFERENCIA between dateadd(d,1,dateadd(yy,-1,p.fimperaquis)) and p.fimperaquis) " & _
"from corporerm.dbo.PFUFERIAS p left join corporerm.dbo.PFUFERIASPER m on m.CHAPA=p.CHAPA and m.FIMPERAQUIS=p.FIMPERAQUIS where PERIODOABERTO=1 " & _
"group by  p.CHAPA, p.FIMPERAQUIS, p.PERIODOABERTO " & _
"having sum(case when m.NRODIASABONO IS null then 0 else m.NRODIASABONO end+case when m.NRODIASFERIAS IS null then 0 else m.NRODIASFERIAS end)<>30 " & _
") c inner join corporerm.dbo.PFUNC f on f.CHAPA=c.CHAPA where CODSITUACAO<>'D' and CODSINDICATO<>'03' and CODTIPO='N' " & _
") c where nrodiasferias>0 " & _
"union " & _
"select f.chapa, f.codsecao, f.codsecao relsecao, dateadd(yy,1,ultimo), null,null,null,null,null " & _
"from corporerm.dbo.pfunc f inner join ( " & _
"select f.chapa, ultimo=(select top 1 FIMPERAQUIS from corporerm.dbo.PFUFERIAS where CHAPA=f.chapa order by FIMPERAQUIS desc) " & _
"from corporerm.dbo.pfuferias f group by f.chapa " & _
") r on r.CHAPA=f.CHAPA " & _
"where f.codsindicato<>'03' and f.codtipo='N' and f.codsituacao not in ('P','I','D') /*and inicprogferias1 is null and ultimo<getdate()*/ " & _
") z "

'response.cookies("intranet_rh")("gerouhoje")="N"
if request.cookies("intranet_rh")("gerouhoje")<>"S" then
	'nao gerou
	response.cookies("intranet_rh")("gerouhoje")="S"
	response.cookies("intranet_rh").expires=dateadd("d",1,now)
	conexao.execute sql1
	conexao.execute sql2
else
	'ja gerou
	'response.cookies("intranet_rh")("gerouhoje")="N"
end if

%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Programação de Férias
<form method="POST" action="programacao.asp" name="form">
<input type="text" name="codsecao" size="8" maxlength="8" onchange="codsecao1()" value="<%=request.form("codsecao")%>">
<select name="secao" class=a onchange="secao1()">
	<option value="0">Selecione o departamento</option>
<!--	<option value="T">Todos departamentos</option> -->
<%
sqla="select p.relsecao codsecao, s.descricao nome from feriasprog p, corporerm.dbo.psecao s where s.codigo collate database_default=p.relsecao " & _
"group by p.relsecao, s.descricao order by s.descricao "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
if request.form("codsecao")=rs("codsecao") or session("ultimosetorprog")=rs("codsecao") then temps="selected" else temps=""
%>
	<option value="<%=rs("codsecao")%>" <%=temps%> ><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close

if request.form("print_abono")="ON" then print_abono="checked" else print_abono=""

%>
</select>
<br>
<input type="checkbox" name="print_abono" <%=print_abono%>>Imprimir colunas/obs. de Abono Pecuniário
<br>
<input type="submit" value="Gerar relatório" name="B1" class="button"></p>
</form>
<%
response.write request.cookies("intranet_rh")("gerouhoje")

if request.form("codsecao")="" then relsecao="99.9.999" else relsecao=request.form("codsecao")
session("ultimosetorprog")=relsecao
sql3="select p.chapa, f.nome, f.funcao, f.admissao, f.secao secao_labore, s.descricao secao_programacao " & _
"from feriasprog p, qry_funcionarios f, corporerm.dbo.psecao s " & _
"where f.chapa collate database_default=p.chapa and s.codigo collate database_default=p.relsecao and relsecao='" & relsecao & "' and sessao='" & sessao & "' " & _
"and codtipo='N' " & _
"group by p.chapa, f.nome, f.funcao, f.admissao, f.secao, s.descricao order by f.nome " 
rs.Open sql3, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" >
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome do funcionário</td>
	<td class=titulo>Admissão</td>
	<td class=titulo>Função</td>
	<td class=titulo>Setor no Labore</td>
	<td class=titulo>Setor p/Programação</td>
</tr>
<%
if rs.recordcount>0 then
do while not rs.eof
%>
<tr>
	<td class=campo><%=rs("chapa")%></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo><%=rs("admissao")%></td>
	<td class=campo><%=rs("funcao")%></td>
	<td class=campo><%=rs("secao_labore")%></td>
	<td class=campo>
	<a href="programacao_setor.asp?codigo=<%=rs("chapa")%>" onclick="NewWindow(this.href,'AlteracaoSetorProgramacao','450','150','yes','center');return false" onfocus="this.blur()">
	<%=rs("secao_programacao")%></a></td>
</tr>
<%
rs.movenext:loop
end if
rs.close

end if 'request.form ("B1")=""


if request.form("B1")<>"" then

if request.form("print_abono")="on" then print_abono="checked" else print_abono=""

sql4="select * from feriasprog p, qry_funcionarios f " & _
"where p.chapa=f.chapa collate database_default and p.relsecao='" & request.form("codsecao") & "' and sessao='" & session("usuariomaster") & "' " & _
"and codtipo='N' " & _
"order by f.nome, p.dtvencferias, inicprogferias1 "
rs.Open sql4, ,adOpenStatic, adLockReadOnly

sql5="select descricao from corporerm.dbo.psecao where codigo='" & request.form("codsecao") & "' "
rs2.Open sql5, ,adOpenStatic, adLockReadOnly
nomesetor=rs2("descricao")
rs2.close
linha=0:pagina=1
do while not rs.eof
if rs.absoluteposition<rs.recordcount then
	rs.movenext
	proxchapa=rs("chapa")
	rs.moveprevious
else
	proxchapa=""
end if
idade=datediff("yyyy",rs("dtnascimento"),rs("dtvencferias"))
if idade<18 or idade>50 then txt1="30 dias" else txt1="10 a 30 dias"
if proxchapa<>rs("chapa") then mudou=1 else mudou=0
if mudou=1 then estilo="style=""border-bottom:2px solid #000000""" else estilo=""
if mudou=1 then estilo2="style=""border-bottom:2px solid #000000""" else estilo2="style=""border-bottom:1px dotted #000000"""
if isnull(rs("inicprogferias1")) then inicio="____/____/______" else inicio=rs("inicprogferias1")
if isnull(rs("nrodiasferias")) or rs("nrodiasferias")=0 then diasferias="_____ <font size=1px>(" & txt1 & ")" else diasferias=rs("nrodiasferias")
if isnull(rs("nrodiasabono")) or rs("nrodiasabono")=0 then diasabono="(&nbsp;&nbsp;) Sim (&nbsp;&nbsp;) Não" else diasabono=rs("nrodiasabono")
if rs("nrodiasferias")=30 then diasabono="---"

if linha=0 or linha>24 then
	if linha>24 then
		response.write "</table>"
		pagina=pagina+1
		response.write "<DIV style=""page-break-after:always""></DIV>"
		linha=0
	end if
%>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="990">
<tr>
	<td class="campop" align="left" height=30 style="border-top:1px solid;border-bottom:1px solid;border-left:1px solid">&nbsp;UNIFIEO</td>
	<td class="campop" align="center" style="border-top:1px solid;border-bottom:1px solid;"><b>PLANEJAMENTO ANUAL DE FÉRIAS <%=YEAR(NOW)%>/<%=YEAR(NOW)+1%></TD>
	<td class="campop" align="right" style="border-top:1px solid;border-bottom:1px solid;border-right:1px solid"><%=now()%> - Pág.: <%=pagina%>&nbsp;</td>
</TR>
<tr><td colspan=3 height=5></td></tr>
<tr>
	<td class=titulop height=20>&nbsp;<%=request.form("codsecao")%></td>
	<td class=titulop colspan=2><%=nomesetor%></td>
</tr>
<tr><td colspan=3 height=5></td></tr>
</table>


<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="990">
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo>Nome do funcionário</td>
	<td class=titulo align="center">Data do<br>Vencimento</td>
	<td class=titulo>Data de<br>Saida</td>
	<td class=titulo>Dias de<br>descanso <font size=1px>(1)</td>
<%if print_abono="checked" then%>	
	<td class=titulo>Abono<br>pecuniário <font size=1px>(2)</td>
<%end if%>
	<td class=titulo width=20%>Observação</td>
</tr>
<%
end if
%>
<tr>
	<td class="campop" <%=estilo%> align="center" height=25><%if mudou=1 then response.write rs("chapa")%></td>
	<td class="campop" <%=estilo%>><b><%if mudou=1 then response.write rs("nome")%></td>
	<td class="campop" <%=estilo2%> align="center"><%=rs("dtvencferias")%></td>
	<td class="campop" <%=estilo2%>><%=inicio%></td>
	<td class="campop" <%=estilo2%>><%=diasferias%></td>
<%if print_abono="checked" then%>	
	<td class="campop" <%=estilo2%>><%=diasabono%></td>
<%end if%>
	<td class="campop" <%=estilo2%> style="border-left:1px solid"></td>
</tr>
<%
linha=linha+1
rs.movenext
loop
rs.close
%>

</table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="990">
<tr><td class="campor">
Observações:
<br>(1) Para maiores de 50 anos, o período de férias não pode ser dividido.
<br>&nbsp;&nbsp;&nbsp;&nbsp;Caso haja a divisão do período de férias, cada período não pode ser inferior a 10 dias e no máximo 2 períodos.
<br>(2) O intervalo entre um período e outro de férias deve ser de no mínimo 4 meses.
<br>(3) Na divisão de férias em 2 períodos, os 2 deverão ser agendados, o 1º no espaço indicado e o 2º em observações.
<br>(4) A data já agendada do planejamento anterior não poderá ser alterada.
<br>(5) Os setores de serviços essenciais e de atendimento devem procurar escalonar as férias dos seus funcionários a fim de não interromper a prestação de serviços.
<br>(6) O início das férias não poderá coincidir com sábado, domingo ou feriado e deverá ter início preferencialmente em uma segunda-feira.
<%if print_abono="checked" then%>	
<br>(7) O abono corresponde a 1/3 (um terço) dos dias adquiridos para descanso. Exemplo: para 30 dias (férias=20/abono=10, para 15 dias (férias=10/abono=5).
<%end if%>
</td><td class="campor">&nbsp;</td></tr>
<tr><td class="campor">&nbsp;</td><td class="campor"><br>_____________________________________<br>Assinatura do encarregado/coordenador</td></tr>

</table>

<%
%>
<%
set rs=nothing
set rs2=nothing
end if ' 

conexao.close
set conexao=nothing
%>
</body>
</html>