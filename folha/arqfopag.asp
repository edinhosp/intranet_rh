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
<title>Geração de Arquivo Remessa-Bradesco</title>
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
<p class=titulo>Geração de arquivo Remessa-Bradesco</p>
<%
'response.write "<br>" & session.sessionid
if request.form("Gerar")="" then
	mesfolha=month(dateserial(year(now),month(now)+1,1))
	anofolha=year(dateserial(year(now),month(now)+1,1))
%>
<form method="POST" action="arqfopag.asp" name="form">
<%if request.form("tipopag")<>"" then tp=request.form("tipopag")%>
<p>Tipo de Pagamento:
<input type="radio" name="tipopag" value="F" onclick="javascript:submit();" <%if tp="F" then response.write "checked"%> >Férias
<input type="radio" name="tipopag" value="R" onclick="javascript:submit();" <%if tp="R" then response.write "checked"%> >Rescisão
<input type="radio" name="tipopag" value="P" onclick="javascript:submit();" <%if tp="P" then response.write "checked"%> >Folha Pagamento
</p>
<%
if tp<>"" then   '******************
	tempd=request.form("dtpagto")
	divisor1=cint(instr(1,tempd,"!"))
	divisor2=cint(instr(1,tempd,"@"))
	divisor3=cint(instr(1,tempd,"#"))
	vezes=0
	if len(tempd)>1 then anocomp=left(tempd,divisor1-1) else anocomp=year(now())
	if len(tempd)>1 then mescomp=mid(tempd,divisor1+1,divisor2-divisor1-1) else mescomp=month(now())
	if len(tempd)>1 then dtpagto=mid(tempd,divisor2+1,divisor3-divisor2-1) else dtpagto=int(now())
	if len(tempd)>1 then nroperiodo=mid(tempd,divisor3+1,len(tempd)-divisor3) else nroperiodo=0

	if tp="F" then sql1="SELECT anocomp=year(datapagto), mescomp=month(datapagto), 'DTPAGTO'=DATAPAGTO, nroperiodo=0, 'Recibos'=Count(Chapa) FROM corporerm.dbo.PFUFERIASRECIBO GROUP BY DATAPAGTO HAVING DATAPAGTO>=getdate()-120 ORDER BY DATAPAGTO;"
	if tp="F" then sql2="select codsindicato, r.chapa, nroperiodo=0, f.nome, f.codbancopagto, f.codagenciapagto, f.contapagamento, 'dtvencimento'=r.FIMPERAQUIS, 'DTPAGTO'=r.DATAPAGTO, opbancaria razao " & _
	", liquido=l.liquido, 'adiantado'=case when a.adiantado is null then 0 else a.adiantado end, 'saldo'=l.liquido-case when a.adiantado is null then 0 else a.adiantado end " & _
	"from corporerm.dbo.PFUFERIASRECIBO r inner join corporerm.dbo.pfunc f on f.chapa=r.chapa " & _
	"inner join ( " & _
	"	SELECT r.CHAPA, 'dtvencimento'=r.fimperaquis, r.DATAPAGTO, 'Liquido'=sum(case when provdescbase='D' then -1 else 1 end * valor) " & _
	"	FROM corporerm.dbo.pfuferiasrecibo r inner join corporerm.dbo.pfuferiasverbas l on r.fimperaquis=l.fimperaquis and r.chapa=l.chapa and r.datapagto=l.datapagto inner join corporerm.dbo.PEVENTO e on l.codevento=e.codigo " & _
	"	WHERE e.PROVDESCBASE in ('D','P') GROUP BY r.CHAPA, r.fimperaquis, r.DaTaPAGTO " & _
	") l on l.chapa=f.chapa and l.datapagto=r.datapagto " & _
	"left join ( " & _
	"	select chapa, dtpagto, nroperiodo, liquido, 'adiantado'=sum(case when valorparc is null then 0 else valorparc end) from creditofolhaparcelas where dtpagto='" & dtaccess(dtpagto) & "' and nroperiodo=" & nroperiodo & " group by chapa, dtpagto, nroperiodo, liquido " & _
	") a on a.chapa=r.CHAPA collate database_default and a.dtpagto='" & dtaccess(dtpagto) & "' and a.nroperiodo=0 and a.liquido=l.liquido " & _
	"where r.datapagto='" & dtaccess(dtpagto) & "' order by codsindicato, codbancopagto, opbancaria, nome "
	
	if tp="R" then sql1="select anocomp, mescomp, dtpagto, nroperiodo, 'Recibos'=COUNT(chapa) from ( " & _
	"select distinct s.anocomp, s.mescomp, s.DTPAGTO, s.chapa, ff.nroperiodo from corporerm.dbo.PFUNC f inner join corporerm.dbo.PFPERFF ff on ff.CHAPA=f.chapa " & _
	"inner join corporerm.dbo.PFFINANC s on s.CHAPA=ff.CHAPA and s.ANOCOMP=ff.ANOCOMP and s.MESCOMP=ff.MESCOMP and s.NROPERIODO=ff.nroperiodo and s.CHAPA=ff.CHAPA " & _
	"where (f.CODSITUACAO='D' or datademissao is not null) and ff.NROPERIODO not in (2,4) and s.DTPAGTO>GETDATE()-15 " & _
	") z group by anocomp, mescomp, dtpagto, nroperiodo "
	if tp="R" then sql2="select distinct f.codsindicato, s.DTPAGTO, s.chapa, s.nroperiodo, f.nome, f.codbancopagto, f.codagenciapagto, f.contapagamento, s.dtpagto, opbancaria razao " & _
	", liquido=s.valor, 'adiantado'=case when a.adiantado is null then 0 else a.adiantado end, 'saldo'=s.valor-case when a.adiantado is null then 0 else a.adiantado end " & _
	"from corporerm.dbo.PFUNC f inner join corporerm.dbo.PFPERFF ff on ff.CHAPA=f.chapa " & _
	"inner join corporerm.dbo.PFFINANC s on s.CHAPA=ff.CHAPA and s.ANOCOMP=ff.ANOCOMP and s.MESCOMP=ff.MESCOMP and s.NROPERIODO=ff.nroperiodo and s.CHAPA=ff.CHAPA " & _
	"left join ( " & _
	"	select chapa, dtpagto, nroperiodo, liquido, 'adiantado'=sum(case when valorparc is null then 0 else valorparc end) from creditofolhaparcelas where dtpagto='" & dtaccess(dtpagto) & "' and nroperiodo=" & nroperiodo & " group by chapa, dtpagto, nroperiodo, liquido " & _
	") a on a.chapa=s.CHAPA collate database_default and a.dtpagto='" & dtaccess(dtpagto) & "' and a.nroperiodo=s.nroperiodo and a.liquido=s.valor " & _
	"where (f.CODSITUACAO='D' or datademissao is not null) and ff.NROPERIODO not in (2,4) and s.DTPAGTO='" & dtaccess(dtpagto) & "' and CODEVENTO='308' order by f.codsindicato, codbancopagto, opbancaria, nome "

	if tp="P" then sql1="select anocomp, mescomp, dtpagto=max(dtpagto), nroperiodo, 'Recibos'=COUNT(chapa) from ( " & _
	"select distinct s.anocomp, s.mescomp, f.DTPAGTO, f.CHAPA, f.NROPERIODO from corporerm.dbo.PFPERFF s inner join corporerm.dbo.PFFINANC f on f.ANOCOMP=s.ANOCOMP and f.MESCOMP=s.MESCOMP and f.NROPERIODO=s.NROPERIODO where f.DTPAGTO>GETDATE()-120 " & _
	") z group by anocomp, mescomp, nroperiodo order by dtpagto "
	if tp="P" then sql2="select distinct p.codsindicato, s.CHAPA, s.NROPERIODO, p.nome, p.codbancopagto, p.codagenciapagto, p.contapagamento, opbancaria razao " & _
	", l.liquido, 'adiantado'=case when a.adiantado is null then 0 else a.adiantado end, 'saldo'=l.liquido-case when a.adiantado is null then 0 else a.adiantado end " & _
	"from corporerm.dbo.PFPERFF s /*inner join corporerm.dbo.PFFINANC f on f.ANOCOMP=s.ANOCOMP and f.MESCOMP=s.MESCOMP and f.NROPERIODO=s.NROPERIODO and f.CHAPA=s.CHAPA*/ " & _
	"inner join corporerm.dbo.PFUNC p on p.CHAPA=s.CHAPA " & _
	"inner join ( " & _
		"select chapa, anocomp, mescomp, nroperiodo, liquido=SUM(case when provdescbase='D' then -1 else 1 end*valor) from corporerm.dbo.pffinanc f inner join corporerm.dbo.pevento e on e.codigo=f.codevento " & _
		"where anocomp=f.anocomp and mescomp=f.mescomp and nroperiodo=f.nroperiodo and e.provdescbase in ('P','D') group by chapa, anocomp, mescomp, nroperiodo having SUM(case when provdescbase='D' then -1 else 1 end*valor)>0 " & _
	") l on l.chapa=s.chapa and l.ANOCOMP=s.ANOCOMP and l.MESCOMP=s.MESCOMP and l.NROPERIODO=s.NROPERIODO " & _
	"left join ( " & _
	"	select chapa, dtpagto, nroperiodo, liquido, 'adiantado'=sum(case when valorparc is null then 0 else valorparc end) from creditofolhaparcelas where dtpagto='" & dtaccess(dtpagto) & "' and nroperiodo=" & nroperiodo & " group by chapa, dtpagto, nroperiodo, liquido " & _
	") a on a.chapa=s.CHAPA collate database_default and a.dtpagto='" & dtaccess(dtpagto) & "' and a.nroperiodo=s.nroperiodo and a.liquido=l.liquido " & _
	"where s.anocomp=" & anocomp & " and s.mescomp=" & mescomp & " and s.NROPERIODO=" & nroperiodo & _
	" and p.codbancopagto='237' " & _
	" order by l.liquido, p.codsindicato, codbancopagto, opbancaria, /*l.liquido,*/ nome "
%>
<select size="1" name="dtpagto" onChange="javascript:submit()">
<option value="0">Selecione uma data</option>&nbsp;
<%
if isdate(dtpagto)=true then dtpagto=cdate(dtpagto)
rs.Open sql1, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
	if tempd=rs("anocomp") & "!" & rs("mescomp") & "@" & rs("dtpagto")&"#"&rs("nroperiodo") then temp1="selected" else temp1=""
	if tp="P" then descr1=" (" & rs("nroperiodo") & ")" else descr1=""
%>
	<option value="<%=rs("anocomp") & "!" & rs("mescomp") & "@" & rs("dtpagto")&"#"&rs("nroperiodo")%>" <%=temp1%>> <%=rs("anocomp") & "/" & rs("mescomp") & " = " & rs("dtpagto")%>&nbsp;&nbsp;&nbsp; (<%=rs("recibos")%> recibos) <%=descr1%></option>
<%
rs.movenext:loop
rs.close
%>
	</select>
<%
if request.form("numarquivo")="" then numarquivo=0 else numarquivo=request.form("numarquivo")
if request.form("arqreal")="ON" then arqreal="checked" else arqreal=""
if request.form("percentual")="" then percentual=100 else percentual=request.form("percentual")
if request.form("datacredito")="" then datacredito=formatdatetime(now(),2) else datacredito=request.form("datacredito")
if request.form("vminimo")="" then vminimo=0 else vminimo=request.form("vminimo")
if request.form("vmaximo")="" then vmaximo=99999 else vmaximo=request.form("vmaximo")
'*******************
nr_parcela=0
sqlp1="select 'ultima'=max(nroparcela), 'total'=SUM(valorparc) from creditofolhaparcelas where dtpagto='" & dtaccess(dtpagto) & "' and nroperiodo=" & nroperiodo & " group by dtpagto, nroperiodo, nroparcela "
sqlp1="select 'ultima'=max(nroparcela) from creditofolhaparcelas where dtpagto='" & dtaccess(dtpagto) & "' and nroperiodo=" & nroperiodo & " " 
rs1.Open sqlp1, ,adOpenStatic, adLockReadOnly
if isnull(rs1("ultima")) then pc_ultima=1 else pc_ultima=cint(rs1("ultima"))
rs1.close
sqlp2="select 'total'=SUM(valorparc) from creditofolhaparcelas where dtpagto='" & dtaccess(dtpagto) & "' and nroperiodo=" & nroperiodo & " and nroparcela=" & pc_ultima & " "
rs1.Open sqlp2, ,adOpenStatic, adLockReadOnly
if isnull(rs1("total")) then pc_total=0 else pc_total=cdbl(rs1("total"))
rs1.close
if pc_total=0 then nr_parcela=pc_ultima else nr_parcela=pc_ultima+1
'response.write "<br>" & pc_ultima
'response.write "<br>" & pc_total
'response.write "<br>" & nr_parcela
'response.write "<br>"
%>
	&nbsp;Pg.<input type="text" name="diaspag" value="2" size=2>dias
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr><td valign="top" class="campo">
<br>Numero de sequência do arquivo: <input type="text" name="numarquivo" value="<%=numarquivo%>" size=2>
<br>O arquivo é teste (desmarque para real) <input type="checkbox" name="arqreal" value="ON" <%=arqreal%>>
<input type="checkbox" name="checkall" onclick="toggleAll(this)" id="Checkbox1" /><font color=green>Selecionar todos</font>
<br>
Gerar <input type="text" name="percentual" value="<%=percentual%>" size="2" onchange="javascript:submit();" >% - 
Valor mínimo: <input type="text" name="vminimo" value="<%=vminimo%>" size="5" onchange="javascript:submit();" > - 
Valor máximo: <input type="text" name="vmaximo" value="<%=vmaximo%>" size="7" onchange="javascript:submit();" > -
Parcela sendo adiantada: <input type="text" name="nr_parcela" value="<%=nr_parcela%>" size="4" onchange="javascript:submit();" ><br>
Data de Crédito Real: <input type="text" name="datacredito" value="<%=datacredito%>"><br>
</td>
<td valign="top">

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr><td class="fundo">Parcela</td><td class="fundo">Data</td><td class="fundo">Valor</td><td class="fundo"></td></tr>
<%
rstotalparc=0.00
sqlparcelas="select dtpagto, nroperiodo, nroparcela, dataparc, Total=SUM(valorparc) from creditofolhaparcelas " & _
"where dtpagto='" & dtaccess(dtpagto) & "' and nroperiodo=" & nroperiodo & " group by dtpagto, nroperiodo, nroparcela, dataparc order by nroparcela "
rs1.Open sqlparcelas, ,adOpenStatic, adLockReadOnly
if rs1.recordcount>0 then
do while not rs1.eof
%>
<tr><td class="campo"><%=rs1("nroparcela")%></td>
	<td class="campo"><%=rs1("dataparc")%></td>
	<td class="campo"><%=formatnumber(rs1("total"),2)%></td>
	<td class="campo">
	<a href="arqfopagdel.asp?nroparcela=<%=rs1("nroparcela")%>&dataparc=<%=rs1("dataparc")%>&nroperiodo=<%=rs1("nroperiodo")%>&dtpagto=<%=rs1("dtpagto")%>" onclick="NewWindow(this.href,'ApagarParcela','490','300','yes','center');return false" onfocus="this.blur()">
	<img src="../images/trash.gif" border="0" alt="Apagar"></a>

	</td>
</tr>
<%
rstotalparc=rstotalparc+cdbl(rs1("total"))
rs1.movenext
loop
rs1.close
end if
%>
<tr><td class="fundo"></td><td class="fundo"></td><td class="fundo"><%=formatnumber(rstotalparc,2)%></td><td class="fundo"></td></tr>
</table>
</td>
</tr>
</table>
<%
end if '******************

if request.form("dtpagto")<>"" then
	prev_perc=cdbl(percentual) ': response.write "<br>" & prev_perc
	prev_vmin=cdbl(nraccess(vminimo)) ': response.write "<br>" & prev_vmin
	prev_vmax=cdbl(nraccess(vmaximo)) ': response.write "<br>" & prev_vmax

	rs.Open sql2, ,adOpenStatic, adLockReadOnly
	'response.write "<br>" & sql2

	if rs.recordcount>0 then
	subtotal=0
%>
<p><input type="submit" value="Gerar arquivo" name="Gerar" class="button"></p>

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>Sind.</td>
	<td class=titulo>Conta</td>
	<td class=titulo>Liquido</td>
	<td class=titulo></td>
	<td class=titulo>Já Pago</td>
	<td class=titulo>Saldo</td>
	<td class=titulo>Pagto Previsto</td>
	<td class=titulo>Sub-Total</td>
</tr>
<%
	tliquido=0
	rs.movefirst
	do while not rs.eof
	banco=rs("codbancopagto")
	classe="campo"
	if rs("razao")<>"07.05" then classe="campol"
	if banco<>"237" then classe="campov"
	'response.write "<br>chapa: " & rs("chapa")
	'response.write " | dtpagto: " & rs("dtpagto")
	'response.write " | nroperiodo: " & rs("nroperiodo")
	'response.write " | liquido: " & rs("liquido")
	saldo=cdbl(rs("saldo"))

	sqlparcela="if not exists (Select chapa from creditofolhaparcelas where chapa='" & rs("chapa") & "' and dtpagto='" & dtaccess(dtpagto) & "' and nroperiodo=" & rs("nroperiodo") & " and nroparcela=" & nr_parcela & ") " & _
"insert into creditofolhaparcelas (chapa, dtpagto, nroperiodo, liquido, nroparcela, valorparc) select '" & rs("chapa") & "', '" & dtaccess(dtpagto) & "', " & rs("nroperiodo") & ", " & nraccess(rs("liquido")) & ", " & nr_parcela & ", 0 " & _
"else " & _
"update creditofolhaparcelas set liquido=" & nraccess(rs("liquido")) & " where chapa='" & rs("chapa") & "' and dtpagto='" & dtaccess(dtpagto) & "' and nroperiodo=" & rs("nroperiodo") & " and nroparcela=" & nr_parcela & " "
conexao.execute sqlparcela
	
	if prev_perc<>100 then 
		previsto=int( (prev_perc)*cdbl(rs("saldo") ) )/100 ': response.write "<br>1 " & previsto
	else
		previsto=saldo ': response.write "<br>2 " & previsto
	end if
	if previsto <= prev_vmin then
		previsto=prev_vmin ': response.write "<br>3min "
		if previsto>saldo then previsto=saldo
	elseif previsto >= prev_vmax then 
		previsto=prev_vmax ': response.write "<br>4max "
	else
		'response.write "<br>5 "
	end if
	subtotal=subtotal+previsto
	if rs("codsindicato")="03" then corfonte="blue" else corfonte="red"
%>
<tr>
	<td class=<%=classe%>><font color="<%=corfonte%>"><%=rs("chapa")%></font></td>
	<td class=<%=classe%>><font color="<%=corfonte%>"><%=rs("nome")%></font></td>
	<td class=<%=classe%>><font color="<%=corfonte%>"><%=rs("codsindicato")%></font></td>
	<td class=<%=classe%>><font color="<%=corfonte%>"><%=rs("codbancopagto") & " / " & rs("razao") & " / " & rs("codagenciapagto") & "-" & rs("contapagamento")%></font></td>
	<td class=<%=classe%> align="right"><font color="<%=corfonte%>"><%=formatnumber(rs("liquido"),2)%></font></td>
	<td class=<%=classe%>><font color="<%=corfonte%>">
		<input type="checkbox" name="em<%=vezes%>" value="ON" <%if rs("codsindicato")="03" then response.write ""%> >
		<input type="hidden" name="id<%=vezes%>" value="<%=rs("chapa")%>">
		<input type="hidden" name="dt<%=vezes%>" value="<%=dtpagto%>">
		<input type="hidden" name="vr<%=vezes%>" value="<%=previsto%>"></font>
	</td>
	<td class=<%=classe%> align="right"><font color="<%=corfonte%>"><%=formatnumber(rs("adiantado"),2)%></font></td>
	<td class=<%=classe%> align="right"><font color="<%=corfonte%>"><%=formatnumber(rs("saldo"),2)%></font></td>
	<td class=<%=classe%> align="right"><font color="<%=corfonte%>"><%=formatnumber((previsto),2)%></font></td>
	<td class=<%=classe%> align="right"><font color="<%=corfonte%>"><%=formatnumber(subtotal,2)%></font></td>
</tr>
<%
	tliquido=tliquido+cdbl(rs("liquido"))
	vezes=vezes+1
	rs.movenext
	loop
	session("credferimp")=vezes-1
	'response.write "<br>" & session("credferimp")
	end if
rs.close
%>
<tr><td class="fundo" colspan="4" align="right">Total</td><td class="fundo" align="right"><%=formatnumber(tliquido,2)%></td><td class="fundo" colspan="5" align="right">&nbsp;</td></tr>
<tr><td class="fundo" colspan="4" align="right">Valor já adiantado</td><td class="fundo" align="right"><%=formatnumber(rstotalparc,2)%></td><td class="fundo" colspan="5" align="right">&nbsp;</td></tr>
<tr><td class="fundo" colspan="4" align="right">Saldo</td><td class="fundo" align="right"><%=formatnumber(tliquido-rstotalparc,2)%></td><td class="fundo" colspan="5" align="right">&nbsp;</td></tr>
</table>

<%
end if '****************** dtpagto<>""

end if '****************** request.form
%>
</form>

<%
if request.form("Gerar")<>"" then
	tempd=request.form("dtpagto")
	divisor1=cint(instr(1,tempd,"!"))
	divisor2=cint(instr(1,tempd,"@"))
	divisor3=cint(instr(1,tempd,"#"))
	if len(tempd)>1 then anocomp=left(tempd,divisor1-1) else anocomp=year(now())
	if len(tempd)>1 then mescomp=mid(tempd,divisor1+1,divisor2-divisor1-1) else mescomp=month(now())
	if len(tempd)>1 then dtpagto=mid(tempd,divisor2+1,divisor3-divisor2-1) else dtpagto=int(now())
	if len(tempd)>1 then nroperiodo=mid(tempd,divisor3+1,len(tempd)-divisor3) else nroperiodo=0
	sequencia=1
	sql="delete from creditofolha where sessao='" & session.sessionid & "' "
	conexao.execute sql
	sql="delete from fopag_remessa where sessao='" & session.sessionid & "' "
	conexao.execute sql
	'parcela=request.form("parcela"):parcela=cdbl(parcela)

' ************* linha header ***************
c01 = "01REMESSA03" & espaco2("CREDITO C/C",15)
c02 = "02856"
c03 = "07050"
c04 = "0564600"
c05 = "6"
c07 = space(2)
c08 = "71925"
c09 = Espaco2("FUNDACAO INSTITUTO DE ENSINO PARA OSASCO", 25)
c10 = "237" & espaco2("BRADESCO",15) & ddmm8(now()) & "01600BPI" 
c11 = ddmm8(dtpagto)
c11 = ddmm8(request.form("datacredito"))
c12 = space(1) & "N" & space(74)
c13 = numzero(sequencia,6)
linha = c01 & c02 & c03 & c04 & c05 & c06 & c07 & c08 & c09 & c10 & c11 & c12 & c13
string_sql = "INSERT INTO fopag_remessa (sessao, registro, ordem) " & _
"SELECT '" & sessao & "', '" & linha & "', '0' "
conexao.execute string_sql
sequencia=sequencia+1

' ************* linha transação ***************

vez=session("credferimp")
'response.write vez
for z=0 to vez step 1
	'response.write "<br>" & z
	em=request.form("em" & z)
	id=request.form("id" & z)
	dt=request.form("dt" & z)
	vr=request.form("vr" & z)
	'response.write "<Br>" & z & "->" & id & " " & dt & " " & em & " " & vr & "<br>"
	if em="ON" then
		sql="INSERT INTO creditofolha ( sessao, data, chapa, valorparc ) SELECT '" & session.sessionid & "', '" & dtaccess(dt) & "', '" & id & "', " & nraccess(vr) & ""
		conexao.execute sql
	end if
	if em<>"ON" then vr=0
	sqlp="update creditofolhaparcelas set dataparc='" & dtaccess(request.form("datacredito")) & "', valorparc=" & nraccess(vr) & _
	" where chapa='" & id & "' and dtpagto='" & dtaccess(dt) & "' and nroperiodo=" & nroperiodo & " and nroparcela=" & request.form("nr_parcela") & " " 
	'response.write "<br>" & sqlp
	conexao.execute sqlp
next
if request.form("tipopag")<>"" then tp=request.form("tipopag")
totalcredito=0
if tp="F" then sql2="SELECT r.CHAPA, 'dtvencimento'=r.fimperaquis, 'dtpagto'=r.DaTaPAGTO, sum(case when provdescbase='D' then -1 else 1 end * valor) AS Liquido " & _
	", p.nome, p.codbancopagto, p.codagenciapagto, p.contapagamento, opbancaria razao " & _
	"FROM corporerm.dbo.pfuferiasrecibo r inner join corporerm.dbo.pfuferiasverbas l on r.fimperaquis=l.fimperaquis and r.chapa=l.chapa and r.datapagto=l.datapagto " & _
	"inner join corporerm.dbo.PEVENTO e on l.codevento=e.codigo " & _
	"inner join corporerm.dbo.pfunc p on p.chapa=r.chapa " & _
	"WHERE e.PROVDESCBASE in ('D','P') GROUP BY r.CHAPA, r.fimperaquis, r.DaTaPAGTO, p.nome, p.codbancopagto, p.codagenciapagto, p.contapagamento, opbancaria " & _
	"HAVING r.DaTaPAGTO='" & dtaccess(dtpagto) & "' and r.chapa in (select chapa collate database_default from creditofolha where sessao='" & session.sessionid & "') "
if tp="F" then sql2="select distinct r.CHAPA, p.nome, p.codbancopagto, p.codagenciapagto, p.contapagamento, opbancaria razao, r.valorparc as liquido " & _
	"from creditofolha r inner join corporerm.dbo.PFUNC p on p.CHAPA=r.chapa collate database_default " & _
	"where sessao='" & session.sessionid & "' and valorparc>0 "
	
if tp="P" then sql2="select distinct f.DTPAGTO, s.CHAPA, f.NROPERIODO, p.nome, p.codbancopagto, p.codagenciapagto, p.contapagamento, opbancaria razao " & _
	", SUM(case when provdescbase='D' then -1 else 1 end*valor) as liquido " & _
	"from corporerm.dbo.PFPERFF s inner join corporerm.dbo.PFFINANC f on f.ANOCOMP=s.ANOCOMP and f.MESCOMP=s.MESCOMP and f.NROPERIODO=s.NROPERIODO and f.CHAPA=s.CHAPA " & _
	"inner join corporerm.dbo.PFUNC p on p.CHAPA=s.CHAPA inner join corporerm.dbo.pevento e on e.codigo=f.codevento " & _
	"where f.DTPAGTO='" & dtaccess(dtpagto) & "' and s.NROPERIODO=" & nroperiodo & " and codevento='LIQ' " & _
	"and codevento<>'308' and codbancopagto='237' and s.chapa in (select chapa collate database_default from creditofolha where sessao='" & session.sessionid & "') " & _
	"group by f.DTPAGTO, s.CHAPA, f.NROPERIODO, p.nome, p.codbancopagto, p.codagenciapagto, p.contapagamento, opbancaria " & _
	"having SUM(case when provdescbase='D' then -1 else 1 end*valor)>=0 " & _
	"order by opbancaria, nome "
if tp="P" then sql2="select distinct r.CHAPA, p.nome, p.codbancopagto, p.codagenciapagto, p.contapagamento, opbancaria razao, r.valorparc as liquido " & _
	"from creditofolha r inner join corporerm.dbo.PFUNC p on p.CHAPA=r.chapa collate database_default " & _
	"where sessao='" & session.sessionid & "' and valorparc>0 "
	
if tp="R" then sql2="select distinct s.DTPAGTO, s.chapa, f.nome, f.codbancopagto, f.codagenciapagto, f.contapagamento, s.dtpagto, opbancaria razao " & _
	", SUM(case when provdescbase='D' then -1 else 1 end*valor) as liquido " & _
	"from corporerm.dbo.PFUNC f inner join corporerm.dbo.PFPERFF ff on ff.CHAPA=f.chapa " & _
	"inner join corporerm.dbo.PFFINANC s on s.CHAPA=ff.CHAPA and s.ANOCOMP=ff.ANOCOMP and s.MESCOMP=ff.MESCOMP and s.NROPERIODO=ff.nroperiodo and s.CHAPA=ff.CHAPA " & _
	"inner join corporerm.dbo.pevento e on e.codigo=s.codevento " & _
	"where (f.CODSITUACAO='D' or f.datademissao is not null) and ff.NROPERIODO not in (2) and s.DTPAGTO='" & dtaccess(dtpagto) & "' " & _
	"and provdescbase in ('P','D') and codevento<>'308' and codbancopagto='237' and s.chapa in (select chapa collate database_default from creditofolha where sessao='" & session.sessionid & "') " & _
	"group by s.DTPAGTO, s.chapa, f.nome, f.codbancopagto, f.codagenciapagto, f.contapagamento, s.dtpagto, opbancaria " & _
	"having SUM(case when provdescbase='D' then -1 else 1 end*valor)>0 " & _
	"order by codbancopagto, opbancaria, nome "
if tp="R" then sql2="select distinct r.CHAPA, p.nome, p.codbancopagto, p.codagenciapagto, p.contapagamento, opbancaria razao, r.valorparc as liquido " & _
	"from creditofolha r inner join corporerm.dbo.PFUNC p on p.CHAPA=r.chapa collate database_default " & _
	"where sessao='" & session.sessionid & "' and valorparc>0 "

rs1.Open sql2, ,adOpenStatic, adLockReadOnly
rs1.movefirst
do while not rs1.eof
	c01 = "1" & space(61) 
	c02 = left(rs1("codagenciapagto"),5)
	c03 = textopuro(rs1("razao"),2) & 0
	c04 = left(rs1("contapagamento"),7)
	c05 = right(rs1("contapagamento"),1)
	c06 = space(2) & espaco2(rs1("nome"),38) & "0" & rs1("chapa")
	liquido1=cdbl(rs1("liquido"))
	vrem=formatnumber(liquido1,2)
    vrem=replace(vrem,".","")
    vrem=replace(vrem,",","")
    c07 = numzero(vrem,13)
	c08 = "298" & space(8) & space(44) & numzero(sequencia,6)
	linha = c01 & c02 & c03 & c04 & c05 & c06 & c07 & c08
	string_sql = "INSERT INTO fopag_remessa (sessao, registro, ordem) " & _
	"SELECT '" & sessao & "', '" & linha & "', '1' "
	conexao.execute string_sql
	totalcredito=cdbl(totalcredito)+cdbl(liquido1)
	sequencia=sequencia+1
	rs1.movenext
loop
rs1.close

' ************* linha trailler ***************
vrem=formatnumber(totalcredito,2)
vrem=replace(vrem,".","")
vrem=replace(vrem,",","")
linha = "9" & numzero(vrem,13) & Space(180) & numzero(sequencia,6)
string_sql = "INSERT INTO fopag_remessa (sessao, registro, ordem) " & _
"SELECT '" & sessao & "', '" & linha & "', '9' "
conexao.execute string_sql

sql="select * from fopag_remessa where sessao='" & sessao & "' order by ordem, substring(registro,195,6) "
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalf=0:totalg=0
rs.movefirst
response.write "<table border='1' cellpadding='1' cellspacing='1' style='border-collapse: collapse' >"
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=titulor>&nbsp;" & rs.fields(a).name & "</td>"
next
response.write "</tr>"
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	if rs.fields(a).type=5 then 
		conteudo=formatnumber(rs.fields(a),2) 
		response.write "<td align=""right"" class=""campor"">&nbsp;" & conteudo & "</td>"
	else 
		conteudo=rs.fields(a)
		if a=2 then conteudo=left(conteudo,126) & "<b><font color=blue>" & mid(conteudo,127,11) & "</font><font color=red>" & mid(conteudo,138,2) & "</font></b>" & mid(conteudo,140,61)
		response.write "<td class=""campor"">&nbsp;" & conteudo & "</td>"
	end if
	'response.write "<td><font size='1'>&nbsp;" &rs.fields(a) & rs.fields(a).type & "</td>"
next
rs.movenext
loop
response.write "</table>"

response.write "<p>"

	caminho="c:\inetpub\wwwroot\rh\temp\"
	if request.form("arqreal")="ON" then extensao=".TST" else extensao=".TXT"
	nomefile="FP" & numzero(day(now),2) & numzero(month(now),2) & request.form("numarquivo") & extensao
	lote=caminho & nomefile
	Set arquivo=CreateObject("Scripting.FileSystemObject")
	Set leitura=arquivo.CreateTextFile(lote, true)
	rs.movefirst
	do while not rs.eof 
		leitura.writeline rs("registro")
	rs.movenext
	loop
	rs.close
	termino=now()
	duracao=(termino-inicio)
	Response.write "<p class=realce><font size=1> Inicio: " & inicio & " Termino: " & termino & " Duracao: " & formatdatetime(duracao,3) & "</font></p>"
	leitura.close
	set leitura=nothing
	set arquivo=nothing

%>
<a href="..\temp\<%=nomefile%>">Arquivo Remessa <%=cmbmes%></a>
<%
end if 'request.form 
%> 

</body>

</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>
