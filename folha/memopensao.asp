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
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Memorando de Pagamento de Pensão</title>
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
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}
--></script>

</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

'response.write request.form

if request.form("b1")="" then
%>
<p class=titulo>Emissão de Memorando de Pagamento - Pensão&nbsp;<%=titulo %>
<form method="POST" action="memopensao.asp" name="form">
<table border="0" width="300" cellspacing="0"cellpadding="3">
<tr>
	<td class=titulo colspan=3>Período de Competência da Pensão</td>
</tr>
<%
if request.form("anocomp")="" then anocomp=year(now) else anocomp=request.form("anocomp")
if request.form("mescomp")="" then mescomp=month(now) else mescomp=request.form("mescomp")
if request.form("nroperiodo")="" then nroperiodo="2" else nroperiodo=request.form("nroperiodo")
if day(now)<25 and request.form("mescomp")="" and mescomp>1 then mescomp=mescomp-1
if day(now)<25 and request.form("mescomp")="" and mescomp=1 then mescomp=12
%>
<tr>
	<td class=fundo>Ano:<input type="text" name="anocomp" size="4" value="<%=anocomp%>"></td>
	<td class=fundo>Mês:<input type="text" name="mescomp" size="2" value="<%=mescomp%>"></td>
	<td class=fundo>NPer:<input type="text" name="nroperiodo" size="1" value="<%=nroperiodo%>" onChange="javascript:submit()"></td>
</tr>
<tr>
	<td class=fundo colspan=3 align="center"><input type="submit" value="Buscar pensionistas" class="button" name="B2">
	</td>
</tr>
</table>


<%
vezes=0
if request.form("B2")<>"" then

sql1="select chapa, nome, desconto, nrodepend, pensionista, percentual, valor=sum(valor) from (" & _
"select v.chapa, f.nome, v.desconto, d.NRODEPEND, d.NOME pensionista, percentual, m.VALOR " & _
"from corporerm.dbo.PFUNC f " & _
"inner join ( select d.CHAPA, d.NRODEPEND, d.NOME, d.RESPONSAVEL, percentual from corporerm.dbo.PFDEPEND d where d.INCPENSAO=1 and CHAPA<'10000' ) d on d.chapa=f.chapa " & _
"inner join ( " & _
"select v.chapa, anocomp, mescomp, nroperiodo, desconto=sum(v.valor) from corporerm.dbo.PFFINANC v " & _
"where v.CODEVENTO in (select codigo from corporerm.dbo.PEVENTO where PROVDESCBASE='D' and DESCRICAO like '%pensao%') group by v.chapa, anocomp, mescomp, nroperiodo " & _
") v on v.CHAPA=f.chapa " & _
"left join corporerm.dbo.PFDEPMOV m on m.ANOCOMP=v.ANOCOMP and m.MESCOMP=v.MESCOMP and m.NROPERIODO=v.NROPERIODO and m.CHAPA=v.CHAPA and m.NRODEPEND=d.NRODEPEND " & _
"where v.ANOCOMP=" & request.form("anocomp") & " and v.MESCOMP=" & request.form("mescomp") & " and v.NROPERIODO=" & request.form("nroperiodo") & " " & _
") z group by chapa, nome, desconto, nrodepend, pensionista, percentual "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulor align="center">Chapa</td>
	<td class=titulor align="center">Nome</td>
	<td class=titulor align="center">Desconto</td>
	<td class=titulor align="center">Pensionista</td>
	<td class=titulor align="center">Valor</td>
	<td class=titulor align="center">-</td>
</tr>
<%
rs.movefirst
do while not rs.eof
%>
<tr>
<%if chapaant<>rs("chapa") then%>
	<%totconf=0:desconto=rs("desconto")%>
	<td class="campor" style="border-top: 1px solid #000000;"><%=rs("chapa")%></td>
	<td class="campor" style="border-top: 1px solid #000000;"><%=rs("nome")%></td>
	<td class="campor" style="border-top: 1px solid #000000;" align="right"><%=formatnumber(rs("desconto"),2)%></td>
<%else%>
	<td class="campor" style="border-top:0 solid #000000;" colspan=3></td>
<%end if%>
	<td class="campor" style="border-top: 1px solid #000000;"><%=rs("pensionista")%></td>
<%
if rs("valor")="" or isnull(rs("valor")) then valorpensao=rs("desconto") else valorpensao=rs("valor")
if rs("valor")="" or isnull(rs("valor")) then corp="red" else corp="black"
totconf=cdbl(totconf)+cdbl(valorpensao)
%>
	<td class="campor" style="border-top: 1px solid #000000;">
		<input type="text" class="form_input7" style="text-align:right;color:<%=corp%>;" size="10" name="vr<%=vezes%>" value="<%=valorpensao%>">
	</td>

	<td class=campo>
		<input type="checkbox" style="size:7px" name="emitir<%=vezes%>" value="ON" <%="checked"%> >
		<input type="hidden" name="id<%=vezes%>" value="<%=rs("chapa")%>">
		<input type="hidden" name="nd<%=vezes%>" value="<%=rs("nrodepend")%>">
	</td>
</tr>
<%
vezes=vezes+1
chapaant=rs("chapa")
rs.movenext
%>
<%
loop
session("credpensao")=vezes-1
end if 'rs.recordcount
rs.close
	sql2="select descricao from corporerm.dbo.pfperff where ANOCOMP=" & request.form("anocomp") & " and MESCOMP=" & request.form("mescomp") & " and NROPERIODO=" & request.form("nroperiodo") & " and chapa='" & chapaant & "'"
	rs.Open sql2, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then descricao=rs("descricao") else descricao=""
	rs.close
	sql2="select distinct top 1 dtpagto from corporerm.dbo.pffinanc where ANOCOMP=" & request.form("anocomp") & " and MESCOMP=" & request.form("mescomp") & " and NROPERIODO=" & request.form("nroperiodo") & " and chapa='" & chapaant & "' and codevento in ('082','090','329','341','342') "
	rs.Open sql2, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then dtpagto=rs("dtpagto") else dtpagto=""
	rs.close
%>
</table>
<table border="0" width="300" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo>
	<input type="text" name="descricao" value="<%=descricao%>" size="40">
	<input type="text" name="dtpagto" value="<%=dtpagto%>" size="8">
	</td>
</tr>
<tr>
	<td class=titulo><input type="submit" value="Imprimir" class="button" name="B1"></td>
</tr>
</table>


<%
end if 'b2<>""

%>
</form>
<%
end if 'b1=""

if request.form("B1")<>"" then
	vez=session("credpensao")
	sql="delete from creditopensao where sessao='" & session.sessionid & "' "
	conexao.execute sql
	for a=0 to vez
		id=request.form("id" & a)
		nd=request.form("nd" & a)
		vr=request.form("vr" & a)
		emitir=request.form("emitir" & a)
		'response.write id & " " & nd & " " & vr & " " & emitir & "<br>"
		if emitir="ON" then
			sql="INSERT INTO creditopensao ( sessao, chapa, nrodepend, valor ) SELECT '" & session.sessionid & "', '" & id & "', " & nd & ", " & replace(vr,",",".") & ";"
			conexao.execute sql
		end if
	next

divisor=25
sql1="select c.chapa, f.nome, c.nrodepend, d.NOME pensionista, d.RESPONSAVEL, d.BANCO, b.NOMEREDUZIDO, d.AGENCIA, a.DIGAG, d.CONTACORRENTE, d.OPBANCARIA, c.valor " & _
"from creditopensao c inner join corporerm.dbo.PFDEPEND d on d.CHAPA=c.chapa collate database_default and d.NRODEPEND=c.nrodepend " & _
"inner join corporerm.dbo.PFUNC f on f.CHAPA collate database_default=c.chapa " & _
"left join corporerm.dbo.GAGENCIA a on a.NUMBANCO=d.BANCO and a.NUMAGENCIA=d.AGENCIA left join corporerm.dbo.GBANCO b on b.NUMBANCO=d.BANCO " & _
"where c.sessao='" & session.sessionid & "' order by d.banco, c.chapa"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
vezes=int(total/divisor)
if total=divisor then vezes=vezes else vezes=vezes+1
rs.close
ultimac="00000"
'***************** <=13 cabe numa folha
for giro=1 to vezes

sql1="select top 100 percent c.chapa, f.nome, c.nrodepend, d.NOME pensionista, d.cpf, d.RESPONSAVEL, d.BANCO, b.NOMEREDUZIDO, d.AGENCIA, a.DIGAG, d.CONTACORRENTE, d.OPBANCARIA, c.valor " & _
"from creditopensao c inner join corporerm.dbo.PFDEPEND d on d.CHAPA=c.chapa collate database_default and d.NRODEPEND=c.nrodepend " & _
"inner join corporerm.dbo.PFUNC f on f.CHAPA collate database_default=c.chapa " & _
"left join corporerm.dbo.GAGENCIA a on a.NUMBANCO=d.BANCO and a.NUMAGENCIA=d.AGENCIA left join corporerm.dbo.GBANCO b on b.NUMBANCO=d.BANCO " & _
"where c.sessao='" & session.sessionid & "' and c.chapa>'" & ultima & "' order by d.banco, c.chapa"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
rs.movefirst:primeirac=rs("chapa"):primeirab=rs("banco")
rs.movelast:ultimac=rs("chapa"):ultimab=rs("banco")
total=rs.recordcount
rs.movefirst

sql2="SELECT sum(valor) as liquido " & _
"FROM creditopensao where sessao='" & session.sessionid & "' " 
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
valor=cdbl(rs2("liquido"))
rs2.close

%>
<!-- <div align="right"> -->
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690" height="990">
<tr><td colspan=6 class=titulop height=35 valign="middle" align="center" style="border-bottom:2 solid"> M E M O R A N D O&nbsp; &nbsp;I N T E R N O
</td></tr>
<!-- corpo da carta -->
<%
data1=cdate(request.form("dtpagto"))
sqld="select diaferiado from corporerm.dbo.gferiado " & _
"where diaferiado='" & dtaccess(data1) & "' "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then data1=data1-1
rs2.close
if weekday(data1)=7 then data1=data1-1
if weekday(data1)=1 then data1=data1-2
dia=day(data1)
mes=monthname(month(data1))
ano=year(data1)
%>
<tr><td class=fundop colspan=3 align="center"> O R I G E M </td>
	<td class=fundop colspan=3 align="center"> D E S T I N O </td></tr>
<tr>
	<td class=campo height=45><b>DEPTO/SEÇÃO:<br><input type="text" class="form_input10" value="Recursos Humanos" size=15></td>
	<td class=campo><b>DATA:<br><input type="text" class="form_input10" value="<%=int(now())%>" size=10></td>
	<td class=campo><b>NÚMERO:<br><input type="text" class="form_input10" value="" size=6></td>

	<td class=campo><b>DEPTO/SEÇÃO:<br><input type="text" class="form_input10" value="Contas a Pagar" size=15></td>
	<td class=campo><b>A ATENÇÃO DE:<br><input type="text" class="form_input10" value="Sr. Nascimento" size=15></td>
	<td class=campo><b>RECEBIDO EM:<br><input type="text" class="form_input10" value="" size=10></td>
</tr>
<tr>
	<td class="campop" colspan=6 height=50 style="border-bottom:2 solid">
	<b>ASSUNTO:</b><br>Pagamento de Pensão Alimentícia (<%=request.form("descricao")%>)</td>
</tr>
	
	
<tr><td colspan=6 height=800 class="campop" align="left" valign=top>

<p align="left" style="margin-top:0;margin-bottom:0;font-size:12pt">
<%
if total=1 then frase="" else frase="s"
if total=1 then frase2="" else frase2="es"
%>
<p align="justify" style="margin-top:0;margin-bottom:0;font-size:10pt;text-indent:10pt;line-height:150%">
Solicitamos a emissão de cheques para que sejam feitos os respectivos depósitos conforme discriminação abaixo,
para pagamento de pensão alimentícia descontada na <%=request.form("descricao")%> para crédito em <%=request.form("dtpagto")%>,
no total de <b>R$ <%=formatnumber(valor,2)%></b>.
<br>

<div align="center">
	<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse" width=95%>
	<tr><td class=titulo align="center">Chapa</td>
		<td class=titulo align="center">Nome</td>
		<td class=titulo align="center">CPF</td>
		<td class=titulo align="center">Agência</td>
		<td class=titulo align="center">C.Corrente</td>
		<td class=titulo align="center">Oper.</td>
		<td class=titulo align="center">Valor</td>
	</tr>
<%
banco1=0:banco2=0
do while not rs.eof
	if rs("banco")="237" then banco1=banco1+cdbl(rs("valor")) else banco2=banco2+cdbl(rs("valor"))
	if rs("responsavel")<>"" then
		nomedeposito=rs("responsavel") & " (<span style='font-size:7px'>" & rs("pensionista") & "</span>)"
	else
		nomedeposito=rs("pensionista")
	end if
	if rs("banco")<>ubanco or isnull(ubanco) or ubanco="" then
		sql2="select banco, totalbanco=sum(valor) from (" & sql1 & ") b where banco='" & rs("banco") & "' group by banco "
		rs2.Open sql2, ,adOpenStatic, adLockReadOnly
		tbanco=rs2("totalbanco")
		rs2.close
%>
	<tr><td class=fundor colspan=7 style="border-top:2px solid #000000"><%=rs("banco")%>-<span style="font-size:10px"><%=rs("nomereduzido")%></span> - (<%=formatnumber(tbanco,2)%>)</td></tr>

<%
	end if
%>
	<tr>
		<td class=campo><%=rs("chapa")%></td>
		<td class=campo><%=nomedeposito%></td>
		<td class=campo><%=rs("cpf")%></td>
		<td class=campo>
			<input type="text" class="form_input" style="text-align:right" size="5" name=ag<%=rs.absoluteposition-1%> value="<%=rs("agencia")%><%if rs("digag")<>"" then response.write "-"&rs("digag")%>">
		</td>
		<td class=campo>
			<input type="text" class="form_input" style="text-align:right" size="10" name=ag<%=rs.absoluteposition-1%> value="<%=rs("contacorrente")%>">
		</td>
		<td class=campo><%=rs("opbancaria")%></td>
		<td class=campo align="right"><%=formatnumber(rs("valor"),2)%>&nbsp;&nbsp;</td>
	</tr>
<%
ubanco=rs("banco")
rs.movenext
loop
%>	
	</table>
</div>
<br>
<input type="text" class="form_input" style="text-align:left" size="50" name=obs value=".">

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr>
	<td class=campo>
Atenciosamente
<br>
<br>
__________________________________
	</td>
	<td class=campo>
<br>
<br>

<br>
<br>
<br>
<br>
	</td>
</tr>
</table>

</td></tr>
<!-- final do corpo da carta -->

<!-- rodapé da carta -->
<tr><td height=30 colspan=6 class="campor"><%=session("usuariomaster")%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="font-size:14px">Total Bradesco: <%=formatnumber(banco1,2)%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Outros Bancos: <%=formatnumber(banco2,2)%></span>
</td></tr>
<!-- final do rodapé da carta -->
</table>
<!-- </div> -->
<%
rs.close
if giro<vezes then response.write "<DIV style=""page-break-after:always""></DIV>"
next ' giro de paginas
%>

<%
end if 
%>
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>