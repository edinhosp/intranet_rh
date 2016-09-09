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
<title>Carta de Crédito de Folha</title>
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

if request.form="" then
%>
<p class=titulo>Emissão de Autorização de Crédito - Folha Pag.&nbsp;<%=titulo %>
<form method="POST" action="creditofolha.asp" name="form">
<table border="0" width="250" cellspacing="0"cellpadding="3">
<tr>
	<td class=titulo>Data de Pagamento da Folha</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="dtpagto">
	<option>Selecione uma data</option>
&nbsp;
<%
sql2="select anocomp, mescomp, dtpagto, nroperiodo from corporerm.dbo.pffinanc group by anocomp, mescomp, dtpagto, nroperiodo having dtpagto>getdate()-60 order by dtpagto"
rs.Open sql2, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
if dtpagto=rs("dtpagto") then temp1="selected" else temp1=""
comp=numzero(rs("mescomp"),2)&"/"&rs("anocomp")
%>
	<option value="<%=numzero(rs("mescomp"),2)&rs("anocomp")&rs("nroperiodo")&rs("dtpagto")%>" <%=temp1%>><%=comp%>&nbsp;(Período: <%=rs("nroperiodo")%>) Pagto em <%=rs("dtpagto")%></option>
<%
rs.movenext:loop
rs.close
%>
	</select></td>
</tr>
</table>

<table border="0" width="250" cellspacing="0" cellpadding="3">
<tr>
	<td class=titulo><input type="submit" value="Visualizar" class="button" name="B1"></td>
</tr>
</table>
</form>
<%
else ' request.form<>""
temp=request.form("dtpagto")
mescomp=left(temp,2)
anocomp=mid(temp,3,4)
nroperiodo=mid(temp,7,1)
dtpagto=mid(temp,8,len(temp)-1)
'response.write anocomp & "<br>" & mescomp & "<br>" & nroperiodo & "<br>" & dtpagto
dtpagto=cdate(dtpagto)
'	rs.Open sqlb, ,adOpenStatic, adLockReadOnly

sql1="select codbancopagto, sum(liquido) as liquido, count(chapa) as quant from ( " & _
"select f.chapa, codbancopagto, SUM(case when provdescbase='D' then -1 else 1 end*valor) as liquido " & _
"from corporerm.dbo.pffinanc f, corporerm.dbo.pevento e, corporerm.dbo.pfunc pf " & _
"where f.codevento=e.codigo and pf.chapa=f.chapa and mescomp=" & mescomp & " and anocomp=" & anocomp & " " & _
"and nroperiodo=" & nroperiodo & " and provdescbase in ('P','D') group by f.chapa, codbancopagto " & _
"having SUM(case when provdescbase='D' then -1 else 1 end*valor)>0 ) as t where codbancopagto='237' group by codbancopagto "

rs.Open sql1, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
vezes=int(total/13)+1
rs.movefirst

valor=cdbl(rs("liquido"))
quant=rs("quant")
total=quant
%>
<!-- <div align="right"> -->
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="620" height="990">
<tr><td height=111><img border="0" src="../images/aguia.jpg" width="236"></td> </tr>
<!-- corpo da carta -->
<tr><td height=800 class="campop" align="left" valign=top>
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
<p align="right" style="margin-top:0;margin-bottom:0;font-size:12pt">Osasco,&nbsp;<%=dia & " de " & mes & " de " & ano %>
<p align="left" style="margin-top:0;margin-bottom:0;font-size:12pt">
<br>
<br>
<br>
Ao<br>
Banco Bradesco S/A<br>
Agência 3390 - Ag.Empresa Alphaville<br>
Barueri - SP<br>
<br>
<br>
Prezados Senhores:<br>
<br>
<br>
<br>
<%
if total=1 then frase="" else frase="s"
if total=1 then frase2="" else frase2="es"
%>
<p align="justify" style="margin-top:0;margin-bottom:0;font-size:12pt;text-indent:200pt;line-height:170%">
Pela presente autorizamos levar a débito de nossa conta corrente nº <b>564.600-6</b>, o valor de <b>R$ <%=formatnumber(valor,2)%></b> (<%=extenso2(valor)%>), 
referente ao pagamento de <b>Folha de Pagamento <%=mescomp&"/"&anocomp%></b>, creditando na<%=frase%> conta<%=frase%> 
corrente<%=frase%> do<%=frase%>&nbsp;<%=quant%> funcionário<%=frase%> conforme relação anexa e disquete, para crédito em <b><%=dtpagto%></b>.
<br>
<br>
<br>
<p align="justify" style="margin-top:0;margin-bottom:0;font-size:11pt;text-indent:300pt">
Atenciosamente
	
</td></tr>
<!-- final do corpo da carta -->

<!-- rodapé da carta -->
<tr><td height=80>
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
		<tr><td><b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b></td> </tr>
		<tr><td><font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000</font></td></tr>
		<tr><td><font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP 06020-190 - Fone: (011) 3651-9999</font></td></tr>
		<tr><td><font face="Arial Narrow">Caixa Postal - ACF - Jd. Ipê - nº 2530 - Osasco - SP - CEP 06053-990</font></td></tr>
	</table>
</td></tr>
<!-- final do rodapé da carta -->
</table>
<!-- </div> -->
<%
rs.close
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