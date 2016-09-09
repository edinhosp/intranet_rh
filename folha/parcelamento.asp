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
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Demonstrativo de Parcelamento de Crédito</title>
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
<script language="javascript" type="text/javascript">
function toggleAll(cb) 
{
        var val = cb.checked;
        var frm = document.forms[0];
        var len = frm.elements.length;
        var i=0;
        for( i=0 ; i<len ; i++) 
        {
                if (frm.elements[i].type=="checkbox" && frm.elements[i]!=cb) 
                {
                        frm.elements[i].checked=val;
                }
        }
}
</script>

</head>
<body>
<%
inicio=now()
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
sessao=session.sessionid
%>
<p class=titulo>Demonstrativo de Parcelamento de Crédito</p>
<%
'response.write "<br>" & session.sessionid
if request.form("Gerar")="" then
	mesfolha=month(dateserial(year(now),month(now)+1,1))
	anofolha=year(dateserial(year(now),month(now)+1,1))
%>
<form method="POST" action="parcelamento.asp" name="form">
<%
if request.form("tipopag")<>"" then tp=request.form("tipopag")
if request.form("ordvr")="ON" then ordvr="checked" else ordvr=""
if request.form("ultpg")<>"" then ultpg=request.form("ultpg") else ultpg=3
%>
<p>Selecione:
<input type="radio" name="tipopag" value="D" onclick="javascript:submit();" <%if tp="D" then response.write "checked"%> >Por Data
 | <i>Ordena liquido</i> <input type="checkbox" name="ordvr" value="ON" <%=ordvr%> onclick="javascript:submit();" >
<input type="radio" name="tipopag" value="C" onclick="javascript:submit();" <%if tp="C" then response.write "checked"%> >Por Chapa
 | <i> Listar os últimos <input type="text" name="ultpg" value="<%=ultpg%>" size="3" onchange="javascript:submit();" >
</p>
<%
if tp<>"" then   '******************
	tempd=request.form("dtpagto")
	divisor1=cint(instr(1,tempd,"!"))
	'divisor2=cint(instr(1,tempd,"@"))
	'divisor3=cint(instr(1,tempd,"#"))
	vezes=0
	if tp="D" then
		if divisor1="" or divisor1=0 then divisor1=2
		if len(tempd)>1 then dtpagto=left(tempd,divisor1-1) else dtpagto=int(now())
		if len(tempd)>1 then nroperiodo=mid(tempd,divisor1+1,len(tempd)-divisor1) else nroperiodo=0
		sql1="select distinct dtpagto, nroperiodo from creditofolhaparcelas order by dtpagto, nroperiodo"
		sql2=""
	else
		chapa=tempd
		sql1="select distinct p.chapa, f.nome from creditofolhaparcelas p inner join corporerm.dbo.PFUNC f on f.CHAPA=p.chapa collate database_default order by f.NOME"
		sql2=""
	end if

if tp="D" then textoselect="Selecione uma data" else textoselect="Selecione um nome"
%>
<select size="1" name="dtpagto" onChange="javascript:submit()">
<option value="0"><%=textoselect%></option>&nbsp;
<%
if isdate(dtpagto)=true then dtpagto=cdate(dtpagto)
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if tp="D" then 
	rs.movefirst:do while not rs.eof
	if tempd=rs("dtpagto")&"!"&rs("nroperiodo") then temp1="selected" else temp1=""
	descr1=" (" & rs("nroperiodo") & ")" 
%>
	<option value="<%=rs("dtpagto")&"!"&rs("nroperiodo")%>" <%=temp1%>> <%=rs("dtpagto")%>&nbsp;&nbsp;&nbsp; <%=descr1%></option>
<%
	rs.movenext:loop
end if
if tp="C" then
	rs.movefirst:do while not rs.eof
	if tempd=rs("chapa") then temp1="selected" else temp1=""
	descr1="" 
%>
	<option value="<%=rs("chapa")%>" <%=temp1%>> <%=rs("chapa")%> - <%=rs("nome")%></option>
<%
	rs.movenext:loop
end if
rs.close
%>
	</select>
<%
'*******************
nr_parcela=0
%>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr><td valign="top" class="campo">
<!--<input type="checkbox" name="checkall" onclick="toggleAll(this)" id="Checkbox1" /><font color=green>Selecionar todos</font>
-->
<br>
</td>
<td valign="top">

</td>
</tr>
</table>
<%
end if '******************

if request.form("dtpagto")<>"" then

if tp="D" then
	sqld="select distinct dataparc from creditofolhaparcelas where dtpagto='" & dtaccess(dtpagto) & "' and nroperiodo=" & nroperiodo & " and valorparc>0 and dataparc is not null order by dataparc"
	rs.Open sqld, ,adOpenStatic, adLockReadOnly
	matriz=0
	redim dataparc(rs.recordcount-1)
	redim tparc(rs.recordcount-1)
	tliquido=0:tsaldo=0:ttotal=0
	do while not rs.eof
		dataparc(matriz)=rs("dataparc")
		matriz=matriz+1
	rs.movenext:loop
	rs.close
	for a=0 to matriz-1
		'response.write "<br>" & a & " - " & dataparc(a)
	next
	if request.form("ordvr")="ON" then sqlo="order by p.liquido " else sqlo="order by p.chapa "
	sqlp="select p.chapa, f.NOME,  p.liquido "
	for a=0 to matriz-1
		sqlp=sqlp & ",'" & dataparc(a) & "'=sum(case when dataparc='" & dtaccess(dataparc(a)) & "' then valorparc else 0 end) "
	next
	sqlp=sqlp & ", 'Total'=SUM(valorparc), 'Saldo'=p.liquido-SUM(valorparc), " & _
	"'tipo'=max(case when codsindicato='03' then 'Prof.' else 'Adm.' end) " & _
	"from creditofolhaparcelas p inner join corporerm.dbo.PFUNC f on f.CHAPA=p.chapa collate database_default " & _
	"where dtpagto='" & dtaccess(dtpagto) & "' and nroperiodo=" & nroperiodo & " " & _
	"group by p.chapa, f.NOME, p.liquido " & sqlo
	
	rs.Open sqlp, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Liquido</td>
	<%for a=0 to matriz-1:response.write "<td class=titulor>" & dataparc(a) & "</td>":next%>
	<td class=titulo>Total</td>
	<td class=titulo>Saldo</td>
	<td class=titulo>Tipo</td>
</tr>
<%
do while not rs.eof
%>
<tr>
	<td class=campor><%=rs("chapa")%></td>
	<td class=campor><%=rs("nome")%></td>
	<td class=campor align=right><%=formatnumber(rs("liquido"),2)%></td>
	<%for a=0 to matriz-1:response.write "<td class=campor align=right>"%> <%if cdbl(rs.fields(a+3).value)>0 then response.write formatnumber(rs.fields(a+3).value,2) else response.write "" %><% response.write "</td>":next%>
	<td class=campor align=right><%if cdbl(rs("total"))>0 then response.write formatnumber(rs("total"),2) else response.write ""%></td>
	<td class=campor align=right><%if cdbl(rs("saldo"))>0 then response.write formatnumber(rs("saldo"),2) else response.write ""%></td>
	<td class=campor><%=rs("tipo")%></td>
</tr>
<%
tliquido=tliquido+cdbl(rs("liquido"))
ttotal=ttotal+cdbl(rs("total"))
tsaldo=tsaldo+cdbl(rs("saldo"))
for a=0 to matriz-1
	tparc(a)=tparc(a)+cdbl(rs.fields(a+3).value)
next
rs.movenext
loop
%>
<tr>
	<td class=campor colspan="2">Totais</td>
	<td class=campor align=right><%=formatnumber(tliquido,2)%></td>
	<%for a=0 to matriz-1:response.write "<td class=campor align=right>"%> <%if cdbl(tparc(a))>0 then response.write formatnumber(tparc(a),2) else response.write "" %><% response.write "</td>":next%>
	<td class=campor align=right><%if cdbl(ttotal)>0 then response.write formatnumber(ttotal,2) else response.write ""%></td>
	<td class=campor align=right><%if cdbl(tsaldo)>0 then response.write formatnumber(tsaldo,2) else response.write ""%></td>
	<td class=campor>&nbsp;</td>
</tr>
</table>

<%	
else
	chapa=request.form("dtpagto")
	sqld="select p.chapa, f.NOME, p.dtpagto, p.nroperiodo, p.liquido, p.dataparc, p.valorparc " & _
	"from creditofolhaparcelas p inner join corporerm.dbo.PFUNC f on f.CHAPA=p.chapa collate database_default " & _
	"where p.chapa='" & chapa & "' and dataparc is not null and valorparc>0 " & _
	"order by dtpagto, nroperiodo, nroparcela "
	sqld="select p.chapa, f.nome, p.dtpagto, p.nroperiodo, p.liquido, total=sum(valorparc), saldo=p.liquido-sum(valorparc) " & _
	"from creditofolhaparcelas p inner join corporerm.dbo.PFUNC f on f.CHAPA=p.chapa collate database_default " & _
	"inner join (select distinct top " & request.form("ultpg") & " dtpagto from creditofolhaparcelas where chapa='" & chapa & "' order by dtpagto desc ) d on d.dtpagto=p.dtpagto " & _
	"where p.chapa='" & chapa & "' and dataparc is not null and valorparc>0 " & _
	"group by p.chapa, f.nome, p.dtpagto, d.dtpagto, p.nroperiodo, p.liquido order by p.dtpagto, nroperiodo "
	rs.Open sqld, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr><td class="titulop" colspan="6"><%=rs("nome")%> </td></tr>
<tr>
	<td class=titulop>Data Venc.</td>
	<td class=titulop>Período</td>
	<td class=titulop>Liquido</td>
	<td class=titulop>Pago</td>
	<td class=titulop>Saldo</td>
	<td class=titulop>...</td>
</tr>
<%
do while not rs.eof
select case rs("nroperiodo")
	case 4
		periodo="13º Salário"
	case 1
		periodo="Rescisão"
	case 5
		periodo="Rescisão"
	case else
		periodo="Folha"
end select
	
%>
<tr>
	<td class=campop><%=rs("dtpagto")%></td>
	<td class=campop><%=periodo%></td>
	<td class=campop align=right><%=formatnumber(rs("liquido"),2)%></td>
	<td class=campop align=right><%=formatnumber(rs("total"),2)%></td>
	<td class=campop align=right><%=formatnumber(rs("saldo"),2)%></td>
	<td class=campop valign=top align=left>
<%
	sqlp="select nroparcela, dataparc, valorparc from creditofolhaparcelas where chapa='" & rs("chapa") & "' and dtpagto='" & dtaccess(rs("dtpagto")) & "' and nroperiodo=" & rs("nroperiodo") & " and dataparc is not null and valorparc>0 " 
	rs1.Open sqlp, ,adOpenStatic, adLockReadOnly
	do while not rs1.eof
		response.write rs1("dataparc") & " | " & formatnumber(rs1("valorparc"),2)
		if rs1.absoluteposition<rs1.recordcount then response.write "<br>"
	rs1.movenext:loop
	rs1.close
%>	
	</td>
</tr>
<%
rs.movenext
loop
%>
<tr><td class="fundo" colspan="6"><b><i>Valores sujeitos à conferência.</i></b></td></tr>
</table>
<%

end if
%>

<%
end if '****************** dtpagto<>""

end if '****************** request.form
%>
</form>


</body>

</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>
