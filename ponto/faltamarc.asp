<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 20000
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a48")="N" or session("a48")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Checagem de Falta de Marca��es</title>
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
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request.form="" then
%>
<p class=titulo>Checagem e Emiss�o de Documento de Falta de Marca��es&nbsp;<%=titulo %>
<%
data_1=dateserial(year(now),month(now),1)
data_2=dateserial(year(now),month(now),day(now))-1
%>
<form method="POST" action="faltamarc.asp">
  <p>Marca��es incompletas entre <input type="text" name="T1" size="12" value="<%=data_1%>" style="text-align:center">
  e <input type="text" name="T2" size="12" value="<%=data_2%>" style="text-align:center">
  <br>
  <input type="checkbox" name="checagem" value="ON">Reemiss�o ap�s checar exce��es<br>
  <input type="checkbox" name="estagiario" value="ON">Somente estagi�rios<br>
  Apenas chapa: <input type="text" size="5" value="" name="chapa"><br>
  <input type="checkbox" name="formulario" value="ON">Imprimir verso com Formul�rio de Justificativa<br>
  <input type="checkbox" name="mesmapagina" value="ON">Imprimir Formul�rio de Justificativa na mesma p�gina<br>
  <input type="submit" value="Visualizar" class="button" name="B1"></p>
</form>
<%
else 'request.form
	data1=request.form("t1")
	data2=request.form("t2")

if request.form("checagem")<>"ON" then
sql1="delete from _marcacoes_checagem "
conexao.Execute sql1, , adCmdText

sql1="insert into _marcacoes_checagem (chapa, data) " & _
"select a.chapa, a.data from ( " & _
"select chapa, data from corporerm.dbo.ABATFUN group by CHAPA, data having count(batida) not in (2,4,6) " & _
") a inner join corporerm.dbo.pfunc f on f.chapa=a.chapa " & _
"where f.codsindicato<>'03' and f.codhorario not in ('183','186','184','02398','02540','251','00818','02466','00357','00869','02193','00445','02329','02449','02196','02345','02194','00865','02121','280','02468') " & _
"and a.data between '" & dtaccess(data1) & "' and '" & dtaccess(data2) & "' " & _
"group by a.chapa, a.data "
'response.write "<br><br>" & sql1
conexao.Execute sql1, , adCmdText

sql1="insert into _marcacoes_checagem (chapa, data) " & _
"select a.chapa, a.data from corporerm.dbo.abatfun a, corporerm.dbo.pfunc f " & _
"where f.chapa=a.chapa and f.codsindicato<>'03' and f.codhorario not in ('Diretores') and f.jornadamensal/60>=180 " & _
"and datepart(dw,a.data)<>7 and a.data between '" & dtaccess(data1) & "' and '" & dtaccess(data2) & "' " & _ 
"group by a.chapa, a.data having count(a.batida)=2 and sum( (case when natureza=0 or natureza=4 then -1 else 1 end) *batida)>420 "
'response.write "<br><br>" & sql1
conexao.Execute sql1, , adCmdText
end if

if request.form("estagiario")="ON" then textoe=" and f.codtipo='T' " else textoe=""	
if request.form("chapa")<>"" then textof=" and f.chapa='" & request.form("chapa") & "' " else textof=""
sql1="select z.chapa, f.nome, f.codsecao, f.codhorario, s.descricao, p.sexo " & _
"from " & _
"(select a.chapa, a.data from _marcacoes_checagem a where a.data between '" & dtaccess(data1) & "' and '" & dtaccess(data2) & "' group by a.chapa, a.data) as z, " & _
"corporerm.dbo.pfunc f, corporerm.dbo.psecao s, corporerm.dbo.ppessoa p  " & _
"where f.chapa collate database_default=z.chapa and f.codsecao=s.codigo and f.codpessoa=p.codigo and f.codsituacao<>'D' " & textoe & textof & _
"group by z.chapa, f.nome, f.codsecao, f.codhorario, s.descricao, p.sexo " & _
"order by f.codsecao, f.nome "
'response.write "<br><br>" & sql1
rs.Open sql1, ,adOpenStatic, adLockReadOnly

if rs.recordcount>0 then
rs.movefirst
do while not rs.eof 
if rs("sexo")="M" then s1="o" else s1="a"
if rs("sexo")="M" then s2="" else s2="a"
if rs("sexo")="M" then s3="o" else s3=""
%>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650"  >
<tr>
	<td class="campop" align="left" valign=top height=62><img src="../images/logo_centro_universitario_unifieo_big.gif" width="250" border="0"></td>
</tr>
<tr>
	<td class="campop" align="right">Osasco,&nbsp;<%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %></td>
</tr>
<tr>
	<td class="campop" valign=top>
	<br>A<%=s3%> Sr<%=s2%>.<br><%=rs("nome")%>&nbsp;(<%=rs("chapa")%>/<%=rs("codhorario")%>)<br>Setor: <%=rs("descricao")%><br>
	</td>
</tr>
<tr><td class="campop">&nbsp;</td></tr>
<tr>
	<td class="campop" valign=top>
		Ref.: Falta de marca��es em seu ponto eletr�nico.<br><br>
	<p style="margin-top:0;margin-bottom:0;text-align:justify">
	Ap�s verifica��o nas marca��es em seu ponto eletr�nico no m�s de <b><%=monthname(month(data1))%></b>, constatamos que
	em alguns dias, V.Sa. deixou de marcar uma ou mais vezes com o seu cart�o de identifica��o suas horas trabalhadas.
	</td>
</tr>
<tr>
	<td class="campop" valign=top><p style="margin-top:0;margin-bottom:0;text-align:justify">
	Relacionamos os dias em que suas marca��es est�o incompletas ou irregulares, e quais foram os hor�rios
	marcados, para auxiliar no preenchimento do quadro de justificativa de falta ou aus�ncia de marca��es.<br>
	Preencher o quadro "Justificativa para Aus�ncia de Marca��o de Ponto" abaixo, e devolver no prazo m�ximo de 48 horas
	ao Recursos Humanos, para regulariza��o, ficando ciente de que as informa��es ser�o
		inclu�das manualmente nas suas marca��es de ponto e conferidas com outros controles eletr�nicos dispon�veis, como catraca eletr�nica, 
		controle de estacionamento, entre outros.
	</td>
</tr>
<tr><td class="campop">&nbsp;</td></tr>
<%
sqla="select count(status) as esquecimento from corporerm.dbo.abatfun a " & _
"where a.chapa='" & rs("chapa") & "' and a.data between '" & dtaccess(data1) & "' and '" & dtaccess(data2) & "' " & _
"and status='D' "
rs2.Open sqla, ,adOpenStatic, adLockReadOnly
if rs2("esquecimento")>0 then esquecimento=rs2("esquecimento") else esquecimento=0
rs2.close
%>
<tr>
	<td class="campop" valign=top><p style="margin-top:0;margin-bottom:0;text-align:justify">
	<b>A Instru��o n� 19/2005-Reitoria de 09/11/2005 (alterada pela Portaria n� 308 de 10/11/2010), no seu item 5, regulamenta penalidades e limita o n�mero de esquecimentos
	 a 2 por m�s. At� o momento, nossos registros totalizam <%=esquecimento%>.
	</td>
</tr>
<tr>
	<td class="campop" valign=top>
<!-- quadro dos dias com marca��es incompletas -->
	<table border="1" bordercolor="#000000" cellpadding="4" cellspacing="0" style="border-collapse: collapse" width=600>
	<tr>
		<td class=titulop width=140>Data das marca��es</td>
		<td class=titulop width=260>Marca��es efetuadas no dia</td>
		<td class=titulop width=200>Observa��o
		<a href="excecao.asp?chapa=<%=rs("chapa")%>&data=0" onclick="NewWindow(this.href,'Inclusao','330','80','no','center');return false" onfocus="this.blur()">�</a>
		</td>
	</tr>
<%
'"and data not in (select data from _marcacoes_excecoes where chapa='" & rs("chapa") & "') " & _
sql2="select a.chapa, a.data, datepart(dw,a.data) as diasem " & _
"from _marcacoes_checagem a where a.chapa='" & rs("chapa") & "' and a.data between '" & dtaccess(data1) & "' and '" & dtaccess(data2) & "' " & _
"group by a.chapa, a.data order by a.data" 
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof 
%>
	<tr>
		<td class="campop" align="center"><%=rs2("data")%> (<%=weekdayname(weekday(rs2("data")),1)%>)</td>
		<td class="campop">
<%
sql3="select batida from corporerm.dbo.abatfun where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs2("data")) & "' order by batida"
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
rs3.movefirst
do while not rs3.eof
	batida=rs3("batida")
	hora=int(batida/60)
	minuto=batida-(hora*60)
	temp=numzero(hora,2) & ":" & numzero(minuto,2)
	response.write temp
	if rs3.absoluteposition<rs3.recordcount then response.write " - "
rs3.movenext
loop
else
	response.write "-"
end if
rs3.close
%>
		</td>
		<td class="campop">
<a href="excecao.asp?chapa=<%=rs("chapa")%>&data=<%=rs2("data")%>" onclick="NewWindow(this.href,'Inclusao','330','80','no','center');return false" onfocus="this.blur()">�</a>
		</td>
	</tr>	
<%
rs2.movenext
loop
end if 'rs2.recordcount>0
rs2.close
%>
	</table>
<!-- final do quadro dos dias com marca��es incompletas -->
	</td>
</tr>
<tr><td class="campop">&nbsp;</td></tr>
<tr><td class="campop">Atenciosamente,<br><br>Recursos Humanos<br><br></td></tr>
<tr><td class="campop">&nbsp;</td></tr>
<!--
<tr><td class="campop" height=100%>&nbsp;</td></tr>
-->
</table>

<%if request.form("formulario")="ON" then response.write "<DIV style=""page-break-after:always""></DIV>" %>
<%if request.form("formulario")="ON" or request.form("mesmapagina")="ON" then%>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650" height=450>
<tr>
	<td class="campop" valign=top>
	<!-- quadro formulario justificativa -->

	<table style="border-collapse: collapse" border="1" bordercolor="#CCCCCC" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td><img border="0" src="../images/logo_centro_universitario_unifieo_big.jpg" width="225" height="50"></td>
		<td align="center"><b><font size="2">Justificativa para Aus�ncia de Marca��o de Ponto</font></b></td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td valign="top"><font size="1">Departamento:</font><br>&nbsp;</td>
		<td width="150" valign="top"><font size="1">M�s:</font><br>&nbsp;</td>
		<td width="100" valign="top"><font size="1">Ano:</font><br><b><%=year(now())%></b></td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td width="80" valign="top"><font size="1">Chapa:</font><br>&nbsp;</td>
		<td valign="top"><font size="1">Nome do Funcion�rio:</font><br>&nbsp;</td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td width="100%" valign="top" colspan="12"><font size="1">Destina-se o presente controle a registrar informa��es do Empregado,
		relativas aos dias e hor�rio de trabalho face a justificativa assinalada. Fica ciente o empregado, e autoriza, que as informa��es ser�o
		inclu�das manualmente nas suas marca��es de ponto e conferidas com outros controles eletr�nicos dispon�veis, como catraca eletr�nica, controle de estacionamento, entre outros.</font></td></tr>
	<tr>
		<td class=fundor colspan=5 align="center" style="border:1px solid #000000"><i><b>Informe apenas as marca��es com problema</td>
		<td class=fundor colspan=7 align="center" style="border:1px solid #000000"><i><b>Assinale o motivo</td>
	</tr>
	<tr>
		<td width="30" valign="middle" rowspan="2" align="center"><font size="1">DIA</font></td>
		<td width="60" valign="middle" rowspan="2" align="center"><font size="1">Hor�rio de Entrada</font></td>
		<td            valign="top"    colspan="2" align="center"><font size="1">Intervalo para refei��o</font></td>
		<td width="60" valign="middle" rowspan="2" align="center"><font size="1">Hor�rio de Sa�da</font></td>
		<td            valign="middle" colspan="7" align="center"><font size="1">Justificativa p/ Aus�ncia</font></td>
	</tr>
	<tr>
		<td width="60" valign="top" align="center"><font size="1">Sa�da</font></td>
		<td width="60" valign="top" align="center"><font size="1">Retorno</font></td>
		<td width="20" valign="top" align="center"><font size="1">EM</font></td>
		<td width="20" valign="top" align="center"><font size="1">EC</font></td>
		<td width="20" valign="top" align="center"><font size="1">AM</font></td>
		<td width="20" valign="top" align="center"><font size="1">TE</font></td>
		<td width="20" valign="top" align="center"><font size="1">EL</font></td>
		<td width="20" valign="top" align="center"><font size="1">RD</font></td>
		<td width="210"valign="top" align="center"><font size="1">Outros</font></td>
	</tr>
<%for a=1 to 6%>
	<tr>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
		<td valign="top" height="25">&nbsp;</td>
	</tr>
<%next%>
	<tr>
		<td valign="top" colspan="12"><font size="1">C�d. Justificativas:<br>
		<b>EM</b> - Esquecimento de marca��o | 
		<b>EC</b> - Esquecimento do crach� |
		<b>AM</b> - Apagar marca��es em excesso/duplicidade<br>
		<b>TE</b> - Trabalho Externo <b><i>(anexar relat�rio identificando)</i></b> |
		<b>EL</b> - Erro Leitura |
		<b>RD</b> - Rel�gio desligado, atrasado etc 
		</font></td>
	</tr></table>

	<table style="border-collapse: collapse"  border="1" bordercolor="#CCCCCC" cellpadding="2" width="600" cellspacing="0">
	<tr>
		<td width="100%">
		<table style="border-collapse: collapse"  border="0" cellpadding="0" width="100%">
		<tr>
			<td width="30%" class="campor" valign="bottom">&nbsp;<br>_____________________<br>Data</td>
			<td width="30%" class="campor" valign="bottom">&nbsp;<br>__________________________<br>Assinatura do Funcion�rio</td>
			<td width="40%" class="campor" valign="bottom">De acordo:<br><br>&nbsp;&nbsp;&nbsp;___________________________________<br>&nbsp;&nbsp;&nbsp;Assinatura da Chefia</td>
		</tr></table>
		</td>
	</tr></table>

	<table border="0" cellpadding="2" width="600" cellspacing="0">
	<tr><td width="100%" align="right" class="campor">Form.RH 05/2012 - <%=rs.absoluteposition%></td>
	</tr></table>

<!-- final do quadro formulario justificativa -->	
	</td>
</tr>

<tr><td class=campo height=100%>&nbsp;</tr>
</table>

<%end if 'fim quadro if  %> 

<%
rs.movenext
'if rs.absoluteposition<rs.recordcount then 
response.write "<DIV style=""page-break-after:always""></DIV>"
loop

'********** rela��o
'response.write "<DIV style=""page-break-after:always""></DIV>
linharel=2
response.write "<table style=border-collapse:collapse' border=1 bordercolor=#CCCCCC cellpadding=0 width=600 cellspacing=0>"
	response.write "<tr><td colspan=4>Comunicados Falta de Marca��o - " & data1 & " a " & data2 & "</td></tr>"
	response.write "<tr>"
	response.write "<td class=titulo>#</td>"
	response.write "<td class=titulo>Chapa</td>"
	response.write "<td class=titulo>Nome</td>"
	response.write "<td class=titulo>Setor</td>"
	response.write "</tr>"
rs.movefirst
do while not rs.eof
	if linharel>68 then
		response.write "</table>"
		response.write "<DIV style=""page-break-after:always""></DIV>"
response.write "<table style=border-collapse:collapse' border=1 bordercolor=#CCCCCC cellpadding=0 width=600 cellspacing=0>"
	response.write "<tr><td colspan=4>Comunicados Falta de Marca��o - " & data1 & " a " & data2 & "</td></tr>"
	response.write "<tr>"
	response.write "<td class=titulo>#</td>"
	response.write "<td class=titulo>Chapa</td>"
	response.write "<td class=titulo>Nome</td>"
	response.write "<td class=titulo>Setor</td>"
	response.write "</tr>"
	linharel=2	
	end if
	response.write "<tr>"
	response.write "<td class=campo>" & rs.absoluteposition & "</td>"
	response.write "<td class=campo>" & rs("chapa") & "</td>"
	response.write "<td class=campo>" & rs("nome") & "</td>"
	response.write "<td class=campo>" & rs("descricao") & "</td>"
	response.write "</tr>"
	linharel=linharel+1
rs.movenext
loop
response.write "</table>"
'********** rela��o

end if 'recordcount
rs.close
%>
<%
end if 'request.form

set rs=nothing
set rs2=nothing
set rs3=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>