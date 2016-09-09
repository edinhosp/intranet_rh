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
<title>Memorando de Pagamento de Rescisão</title>
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
<script language="VBScript">
	Sub Calcula()
		if document.form.salario2.value<>"" then salario2=cdbl(document.form.salario2.value) else salario2=0
		salario1=document.form.salario1.value
		totalsalario=salario1+salario2+salario3+salario4
		document.form.salariototal.value=formatnumber(totalsalario,2)
			document.form.inss2.value=formatnumber(inss2,2)
			document.form.inss2.value=""
		if teste1-sc<>0 then acerto1=teste1-sc:document.form.sc1.value=formatnumber(sc1-acerto1,2)
		if teste2-inss<>0 then acerto2=teste2-inss:document.form.inss1.value=formatnumber(inss1-acerto2,2)
		'document.form.empresa4.value=formatnumber(acerto2,2) & " - " & teste2
	End Sub
</script>
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

if request.form("b1")="" then
%>
<p class=titulo>Emissão de Memorando de Pagamento - Rescisão&nbsp;<%=titulo %>
<form method="POST" action="memorescisao.asp" name="form">
<table border="0" width="250" cellspacing="0"cellpadding="3">
<tr>
	<td class=titulo>Data de Pagamento da Rescisão</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="dtpagto" onChange="javascript:submit()">
	<option>Selecione uma data</option>
&nbsp;
<%
if isdate(request.form("dtpagto"))=true then dtpagto=cdate(request.form("dtpagto"))
sql2="select dtpagtorescisao, count(chapa) as recibos from corporerm.dbo.pfunc f where dtpagtorescisao>=getdate()-90 group by dtpagtorescisao order by dtpagtorescisao "
rs.Open sql2, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
if dtpagto=rs("dtpagtorescisao") then temp1="selected" else temp1=""
%>
	<option value="<%=rs("dtpagtorescisao")%>" <%=temp1%>><%=rs("dtpagtorescisao")%>&nbsp;&nbsp;&nbsp; (<%=rs("recibos")%> recibos)</option>
<%
rs.movenext:loop
rs.close
%>
	</select>
	&nbsp;Pg.<input type="text" name="diaspag" value="2" size=2>dias
</td>
</tr>
<tr>
	<td class=fundo>
	 	<input type="radio" name="sind" value="T" onClick="javascript:submit()" <%if request.form("sind")="" or request.form("sind")="T" then response.write "checked" %> > Todos
		<input type="radio" name="sind" value="01" onClick="javascript:submit()" <%if request.form("sind")="01" then response.write "checked" %> > Administrativos 
	 	<input type="radio" name="sind" value="03" onClick="javascript:submit()" <%if request.form("sind")="03" then response.write "checked" %> > Professores
	</td>
</tr>
</table>

<table border="0" width="250" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo><input type="submit" value="Visualizar" class="button" name="B1"></td>
</tr>
</table>

<%
vezes=0
if request.form<>"" then
if request.form("sind")="03" then selecao=" and f.codsindicato='03' "
if request.form("sind")="01" then selecao=" and f.codsindicato<>'03' "
if request.form("sind")="" then selecao=""
sql1="select f.chapa, f.nome, f.codbancopagto, f.codagenciapagto, f.contapagamento, f.datademissao, f.dtpagtorescisao, opbancaria razao " & _
"from corporerm.dbo.pfunc f " & _
"where f.dtpagtorescisao='" & dtaccess(request.form("dtpagto")) & "' " & selecao
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
%>
<br>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=400>
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Dt.Saida</td>
	<td class=titulo></td>
</tr>
<%
rs.movefirst
do while not rs.eof
banco=rs("codbancopagto")
if banco<>"237" then classe="campol" else classe="campo"
if rs("razao")<>"07.05" then classe="campov" else classe="campo"
%>
<tr>
	<td class=<%=classe%>><%=rs("chapa")%></td>
	<td class=<%=classe%>><%=rs("nome")%></td>
	<td class=<%=classe%>><%=rs("datademissao")%></td>
	<td class=<%=classe%>>
		<input type="checkbox" name="emitir<%=vezes%>" value="ON" <%="checked"%> >
		<input type="hidden" name="id<%=vezes%>" value="<%=rs("chapa")%>">
		<input type="hidden" name="dt<%=vezes%>" value="<%=rs("dtpagtorescisao")%>">
	</td>
</tr>
<%
vezes=vezes+1
rs.movenext
loop
session("credrescimp")=vezes-1
end if
rs.close
%>
</table>
<%
end if
%>
<p><input type="checkbox" name="cartas" value="ON">Imprimir cartas para o PAB

</form>
<%
else ' request.form<>""
	vez=session("credrescimp")
	sql="delete from creditorescisao where sessao='" & session.sessionid & "' "
	conexao.execute sql
	for a=0 to vez
		id=request.form("id" & a)
		dtpg=request.form("dt" & a)
		emitir=request.form("emitir" & a)
		dtpg=request.form("dt" & a)
		'response.write id & " " & tabela & " " & emitir & "<br>"
		if emitir="ON" then
			sql="INSERT INTO creditorescisao ( sessao, data, chapa ) SELECT '" & session.sessionid & "', '" & dtaccess(dtpg) & "', '" & id & "'"
			conexao.execute sql
		end if
	next

valor=4599.99
dtpagto=cdate(request.form("dtpagto"))
dtpagto1=dtpagto

sql1="select ff.chapa, f.nome, f.codagenciapagto, f.contapagamento, f.codsecao, f.dtpagtorescisao, tipodemissao, datademissao, e.chapa consig " & _
", Liquido=sum(case codevento when '308' then valor else null end) " & _
", BaseGFIP=sum(case when tipodemissao='4' or tipodemissao='8' then null else (case codevento when '308' then null else valor end) end) " & _
"from (corporerm.dbo.pffinanc ff inner join corporerm.dbo.pfunc f on ff.chapa=f.chapa) left join " & _
"(select chapa from emprestimos where (vencu>'" & dtaccess(dtpagto) & "' /* or obs not like '%quitado%'*/) group by chapa) e " & _
"on f.chapa collate database_default=e.chapa where ff.dtpagto='" & dtaccess(dtpagto) & "' and codevento in ('308','303','304','306','307','896') " & _
"and ff.chapa collate database_default in (select chapa from creditorescisao where sessao='" & session.sessionid & "') " & _
"group by ff.chapa, f.nome, f.codagenciapagto, f.contapagamento, f.codsecao, f.dtpagtorescisao, tipodemissao, datademissao, e.chapa order by ff.chapa "
sqlcarta=sql1
rs.Open sql1, ,adOpenStatic, adLockReadOnly
'response.write sql1
total=rs.recordcount
vezes=int(total/13)
if total=13 then vezes=vezes else vezes=vezes+1
rs.close
ultima="00000"
'***************** <=13 cabe numa folha
for giro=1 to vezes

sql1="select top 13 ff.chapa, f.nome, f.codagenciapagto, f.contapagamento, f.codsecao, f.dtpagtorescisao, tipodemissao " & _
", Liquido=sum(case codevento when '308' then valor else 0 end) " & _
", BaseGFIP=sum(case when tipodemissao='4' or tipodemissao='8' then 0 else (case codevento when '308' then 0 else valor end) end) " & _
"from corporerm.dbo.pffinanc ff, corporerm.dbo.pfunc f  where ff.chapa=f.chapa and ff.chapa>'" & ultima & "' " & _
"and ff.dtpagto='" & dtaccess(dtpagto) & "' and codevento in ('308','303','304','306','307','896') " & _
"and ff.chapa collate database_default in (select chapa from creditorescisao where sessao='" & session.sessionid & "') " & _
"group by ff.chapa, f.nome, f.codagenciapagto, f.contapagamento, f.codsecao, f.dtpagtorescisao, tipodemissao order by ff.chapa "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
rs.movefirst:primeira=rs("chapa")
rs.movelast:ultima=rs("chapa")
total=rs.recordcount
rs.movefirst
totalr=cdbl(0)
totalf=cdbl(0)
textoconsignado=""
%>
<div align="center">
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650" height="990">
<tr><td colspan=6 class=titulop height=35 valign="middle" align="center" style="border-bottom:2 solid"> M E M O R A N D O&nbsp; &nbsp;I N T E R N O
</td></tr>
<!-- corpo da carta -->
<%
data1=dtpagto-2
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
<tr><td class=fundop colspan=3 align="center"> O R I G E M </td><td class=fundop colspan=3 align="center"> D E S T I N O </td></tr>
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
	<b>ASSUNTO:</b><br>Pagamento de Verbas Rescisórias <%=monthname(month(dtpagto+request.form("diaspag")))%>/<%=year(dtpagto+request.form("diaspag"))%></td>
</tr>
	
	
<tr><td colspan=6 height=800 class="campop" align="left" valign=top>

<p align="left" style="margin-top:0;margin-bottom:0;font-size:12pt">
<br>
<br>
<br>
<%
if total=1 then frase="" else frase="s"
if total=1 then frase2="" else frase2="es"
%>
<p align="justify" style="margin-top:0;margin-bottom:0;font-size:11pt;text-indent:10pt;line-height:150%">
Solicitamos os valores, conforme abaixo discriminados, para pagamento em <b><%=rs("dtpagtorescisao")%></b> das Rescisões Contratuais do mês de <%=monthname(month(dtpagto+request.form("diaspag")))%>/<%=year(dtpagto+request.form("diaspag"))%>:
<br>

<div align="center">
	<table border="1" bordercolor="#000000" cellpadding="4" cellspacing="0" style="border-collapse: collapse" width=99%>
	<tr><td class=titulop align="center">Nome</td>
		<td class=titulop align="center">Agência</td>
		<td class=titulop align="center">C.C.</td>
		<td class=titulop align="center" style="border-left:2px solid #000000;border-right:2px solid #000000">Liquido<br>Rescisão</td>
		<td class=titulop align="center" style="border-left:2px solid #000000;border-right:2px solid #000000">F.G.T.S.</td>
	</tr>
<%
do while not rs.eof
rescisao=cdbl(rs("liquido"))
gfip=cdbl(rs("basegfip"))
sql2="select chapa from emprestimos where chapa='" & rs("chapa") & "' and (vencu>'" & dtaccess(dtpagto) & "' /*or obs not like '%quitado%'*/) "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then consignado=1 else consignado=0
rs2.close
if consignado=1 then
	c_emprestimo=int(rescisao*30)/100
	c_rescisao=rescisao-c_emprestimo
	textoconsignado="<b>(*)</b> Valores referentes a empréstimos consignados com o Bradesco."
end if
%>
	<tr><td class=campo><b><%=rs("nome")%></b><br>--><%=rs("codsecao")%></td>
		<td class=campo align="center">
		<input type="text" class="form_input" style="text-align:right" size="5" name=ag<%=rs.absoluteposition-1%> value="<%=rs("codagenciapagto")%>">
		</td>
		<td class=campo align="center">
		<input type="text" class="form_input" style="text-align:right" size="8" name=cc<%=rs.absoluteposition-1%> value="<%=rs("contapagamento")%>">
		</td>
<%if consignado=1 then%>
		<td class=campo align="right" style="border-left:2px solid #000000;border-right:2px solid #000000">
		<%=formatnumber(c_rescisao,2)%>&nbsp;&nbsp;<br>
		<%=formatnumber(c_emprestimo,2)%> <b>(*)</b>&nbsp;
		</td>
<%else%>
		<td class=campo align="right" style="border-left:2px solid #000000;border-right:2px solid #000000"><%=formatnumber(rescisao,2)%>&nbsp;&nbsp;</td>
<%
end if
if gfip>0 then
%>
		<td class=campo align="right" style="border-left:2px solid #000000;border-right:2px solid #000000">
		<input type="text" class="form_input" style="text-align:right" size="8" name=gfip<%=rs.absoluteposition-1%> value="<%=formatnumber(gfip,2)%>">&nbsp;&nbsp;</td>
<%else%>
		<td class=campo align="center" style="border-left:2px solid #000000;border-right:2px solid #000000">---</td>
<%end if%>
	</tr>
<%
totalr=totalr+rescisao
totalf=totalf+gfip
rs.movenext
loop
%>	
<tr>
	<td class=fundo colspan=3>Totais</td>
	<td class=fundo align="right" style="border-top:5px double #000000;"><%=formatnumber(totalr,2)%>&nbsp;&nbsp;</td>
	<td class=fundo align="right" style="border-top:5px double #000000;">
	<input type="text" class="form_input" style="text-align:right;background-color:Silver" size="10" name=totalf value="<%=formatnumber(totalf,2)%>">&nbsp;&nbsp;</td>

</tr>
<tr>
	<td class=fundo colspan=4>Total Geral</td>
	<td class=fundo align="right" style="border-top:5px double #000000;"><b>
	<input type="text" class="form_input" style="text-align:right;background-color:Silver;font-style:bold" size="10" name=totalf value="<%=formatnumber(totalf+totalr,2)%>">&nbsp;&nbsp;</td>

</tr>

	</table>
</div>
<br>
<%=textoconsignado%>
<input type="text" class="form_input" style="text-align:left" size="50" name=obs value=".">
<br>
<br>
<p align="justify" style="margin-top:0;margin-bottom:0;font-size:11pt;text-indent:12pt">
Atenciosamente
<br>
<br>
__________________________________
	
	
</td></tr>
<!-- final do corpo da carta -->

<!-- rodapé da carta -->
<tr><td height=30 colspan=6 class="campor"><%=session("usuariomaster")%>

</td></tr>
<!-- final do rodapé da carta -->
</table>
</div>
<%
rs.close
if giro<vezes then response.write "<DIV style=""page-break-after:always""></DIV>"
next ' giro de paginas

if request.form("cartas")="ON" then
' cartas para o banco
sql3=sqlcarta
rs.Open sql3, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
do while not rs.eof
%>
<br>
<div align="center"><center>
<table border="0" cellpadding="5" width="650" cellspacing="0" height="990">
<!-- linha da aguia -->
<tr><td height=112><img src="../images/aguia.jpg" border="0" width="236" height="111" alt=""></td></tr>
<!-- linha intermediaria -->
<tr><td height="20">&nbsp;</td></tr>
<!-- corpo da declaracao -->
<tr><td height=100% valign=top class="campop">
Osasco,&nbsp;<%=day(now) & " de " & monthname(month(now())) & " de " & year(now()) %><bR><br><br>
Ao<br>
Posto de Atendimento Banco Bradesco - FIEO<br>
Ag. 2856-8 - Cidade de Deus<Br>
Osasco - SP<br>
<bR><br><bR>
Prezados Senhores<br><br><br><br>
<p style="margin-bottom:5;text-indent: 40pt; text-align=justify">Pela presente informamos que o(a) Sr(a). <b><%=rs("nome")%></b>, correntista da agência <b><%=rs("codagenciapagto")%></b> 
com conta corrente nº <b><%=rs("contapagamento")%></b>, foi desligado(a) da Fundação Instituto de Ensino para Osasco em <b><%=rs("datademissao")%></b>, não fazendo
jus aos benefícios concedidos por este PAB.
<%
if rs("consig")=rs("chapa") then
	if cdbl(rs("liquido"))>0 then c_emprestimo=int(cdbl(rs("liquido"))*30)/100 else c_emprestimo=0
	'if cdbl(rs("liquido"))>0 then c_emprestimo=c_emprestimo else c_emprestimo=0
%>
<p style="margin-bottom:5;text-indent: 40pt; text-align=justify">Igualmente, informamos que para quitação parcial de empréstimos consignados, será encaminhado no dia
<%=rs("dtpagtorescisao")%> um cheque no valor de R$ <%=formatnumber(c_emprestimo,2)%>, correspondente a 30% do valor líquido da rescisão.
<%end if%>

<p style="margin-bottom:5;text-indent: 40pt; text-align=justify">Recebam nossas considerações.</b> 
<Br><br><br><br><br><br>
<p align="justify" style="margin-top:0;margin-bottom:0;font-size:11pt;text-indent:12pt">
Atenciosamente
<br>
<br>
__________________________________

	</td>
</tr>
<!-- linha intermediaria -->
<tr><td height="20">&nbsp;</td></tr>
<tr><td height="1"><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b><img border="0" src="../images/branco.gif" width=10 height=10></td></tr>
<tr><td height="1"><font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000 - C.N.P.J. 73.063.166/0001-20</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP 06020-190 - Fone: (011) 3651-9999 - C.N.P.J. 73.063.166/0003-92</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Av. Franz Voegeli, 1743 - Osasco - SP - CEP 06020-190 - Fone: (011) 3654-0655 - C.N.P.J. 73.063.166/0004-73</font></td></tr>
<tr><td height="1"><font face="Arial Narrow">Caixa Postal - ACF - Jd. Ipê - nº 2530 - Osasco - SP - CEP 06053-990</font></td></tr>
</table>
</center></div>
<%
rs.movenext
loop
rs.close

end if 'rs.recordcount cartas
end if 'request.form cartas

end if ' request.form
%>
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>