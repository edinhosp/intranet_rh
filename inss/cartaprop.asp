<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a39")="N" or session("a39")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Carta de Proporcionalidade</title>
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
function chapa1() { form.nome.value=form.chapa.value; }
function nome1() { form.chapa.value=form.nome.value; }
--></script>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsi=server.createobject ("ADODB.Recordset")
Set rsi.ActiveConnection = conexao
dataant=dateserial(year(now),month(now)-1,1)
session("mesf")=month(dataant)
session("anof")=year(dataant)

teste=0

if request("codigo")<>"" or request.form<>"" then
	if request.form="" then temp=request("codigo")
	if request("codigo")="" then temp=request.form("chapa")
 
	sqla="SELECT FF.CHAPA, ff.nroperiodo, FF.ANOCOMP, FF.MESCOMP, FF.CODEVENTO, E.DESCRICAO, e.provdescbase, [VALOR]*(case when provdescbase='P' then 1 else -1 end) AS ValorEvento, E.INCINSS, E.ESTINSS " & _
	"FROM corporerm.dbo.PFFINANC FF INNER JOIN corporerm.dbo.PEVENTO E ON FF.CODEVENTO = E.CODIGO " & _
	"WHERE (FF.CHAPA='" & request.form("chapa") & "' AND FF.ANOCOMP=" & request.form("T1") & " AND FF.MESCOMP=" & request.form("T2") & " AND E.INCINSS=1 AND (E.PROVDESCBASE='P' Or E.PROVDESCBASE='D')) OR " & _
	"(FF.CHAPA='" & request.form("chapa") & "' AND FF.ANOCOMP=" & request.form("T1") & " AND FF.MESCOMP=" & request.form("T2") & " AND E.ESTINSS=1 AND (E.PROVDESCBASE='P' Or E.PROVDESCBASE='D')) " & _
	"ORDER BY ff.nroperiodo, FF.ANOCOMP DESC , FF.MESCOMP DESC "
	
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
		session("chapa")=rs("chapa")
		chapa=rs("chapa")
	else
		session("chapa")=temp
		chapa=temp
	end if
	'nome=rs("nome")
	temp=0
	'if rs.recordcount>0 and session("cartateto")<>"L" then temp=2
	session("mesf")=request.form("T2")
	session("anof")=request.form("T1")
	mescarta=dateserial(session("anof"),session("mesf")+1,1)
else
	temp=1
end if
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>
Seleção do funcionário para emissão de comprovante
<form method="POST" action="cartaprop.asp" name="form">
<p style="margin-top: 0; margin-bottom: 0">&nbsp;Chapa <input type="text" name="chapa" size="5" class="form_box" onchange="chapa1()" value="<%=session("chapa")%>" >
<select name="nome" onchange="nome1()">
	<option value="00000">Selecione...</option>
<%
sql="select chapa, nome from corporerm.dbo.pfunc where codsituacao in ('A','F') order by nome"
rsi.Open sql, ,adOpenStatic, adLockReadOnly
rsi.movefirst
do while not rsi.eof
if rsi("chapa")=session("chapa") then tmpproc="selected" else tmpproc=""
if rsi("chapa")=session("chapa") then session("nomecarta")=rsi("nome")
%>
	<option value="<%=rsi("chapa")%>" <%=tmpproc%>><%=rsi("nome")%></option>
<%
rsi.movenext
loop
rsi.close
%>
</select>

  </p>
  <p style="margin-top: 0; margin-bottom: 0">
  Ano <input type="text" name="T1" size="5" value="<%=session("anof")%>"> 
  Mês <input type="text" name="T2" size="3" value="<%=session("mesf")%>"></p>
  <p style="margin-top: 0; margin-bottom: 0">
  <input type="submit" value="Pesquisar" name="B1" class="button">
  </p>
</form>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=400>
	<tr>
		<td class=titulo>Per.</td>
		<td class=titulo>Cod.</td>
		<td class=titulo>Descrição Evento</td>
		<td class=titulo>Tipo</td>
		<td class=titulo>Valor</td>
	</tr>
<%
if request.form<>"" then
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
	<tr>
		<td class=campo><%=rs("nroperiodo")%>&nbsp;</td>
		<td class=campo><%=rs("codevento")%>&nbsp;</td>
		<td class=campo><%=rs("descricao")%>&nbsp;</td>
		<td class=campo><%=rs("provdescbase")%>&nbsp;</td>
		<td class=campo align="right"><%=formatnumber(rs("valorevento"),2)%>&nbsp;</td>
	</tr>
<%
totalverbas=totalverbas+cdbl(rs("valorevento"))
rs.movenext
loop
end if 'rs.recordcount
rs.close
end if
%>
	<tr>
		<td class=campo colspan=4>Total</td>
		<td class=campo align="right"><%=formatnumber(totalverbas,2)%>&nbsp;</td>
	</tr>
</table>
<p>Opções para geração da carta</p>
<form method="POST" action="cartaprop2.asp" name="form2">
<input type="hidden" name="chapacarta" size="15" value="<%=session("chapa")%>">
<input type="hidden" name="nomecarta" size="15" value="<%=session("nomecarta")%>">

para o mês de <input type="text" name="mescarta" size="20" value="<%=ucase(monthname(month(mescarta+1))) & "/" & year(mescarta)%>">
<br>
valor do salário <input type="text" name="valorcarta" size="15" value="<%=formatnumber(totalverbas,2)%>">
<br>
  <input type="submit" value="Gerar" name="B2" class="button">
</form>
<p>É possivel digitar os valores das outras empresas e fechar a carta de proporção.

</body>

</html>
<%
conexao.close
set conexao=nothing
%>