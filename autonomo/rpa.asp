<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a52")="N" or session("a52")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Pagamentos - RPA</title>
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
conexao.Open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

sqla="SELECT id_lanc, id_autonomo, data_emissao, data_pagamento, descricao_servico, " & _
"servico_prestado, desc_outros_rend, outros_rendimentos, desconto_ir, desconto_inss, " & _
"desconto_iss, descricao_outros, outros_descontos, valor_liquido, apuracao_darf, venc_darf, " & _
"emitiu_rpa, emitiu_darf, ctl_darf, inss_outra_empresa " & _
"FROM autonomo_rpa " 
sqlb="WHERE id_autonomo=" & request("codigo") & " "
sqlc="ORDER BY data_emissao, id_lanc "

sql1=sqla & sqlb & sqlc
rs.Open sql1, ,adOpenStatic, adLockReadOnly
	
sql2="select * from autonomo where id_autonomo=" & request("codigo")
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
id_autonomo=request("Codigo")	
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>
<% if session("a52")="T" then %>
<a href="rpa.asp?codigo=<%=id_autonomo%>" onMouseOver="window.status='Clique aqui para atualizar após as alterações'; return true" onMouseOut="window.status=''; return true" onmouseover >
<img border="0" src="../images/write.gif" alt="Clique para atualizar">
<font size="1">!</font>
</a>
<% end if %>
CONTROLE DE PAGAMENTOS DE RPA</p>
<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr><td class=grupo>Dados Pessoais</td></tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;Nome do autônomo</td>
	<td class=titulor>&nbsp;Telefone</td>
</tr>
<tr>
	<td class="campor"><b>&nbsp;<%=rs2("nome_autonomo")%></b></td>
	<td class="campor">&nbsp;<%=rs2("telefone")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;Endereço</td>
	<td class=titulor>&nbsp;Tipo Serviço habitual</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rs2("rua") & " " & rs2("numero") & " " & rs2("complemento") & " - " & rs2("cidade") & " - " & rs2("cep") %></td>
	<td class="campor">&nbsp;<%=rs2("tipo_prestacao")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>&nbsp;CPF</td>
	<td class=titulor>&nbsp;PIS/NIT</td>
	<td class=titulor>&nbsp;Identidade</td>
	<td class=titulor>&nbsp;CCM</td>
	<td class=titulor>&nbsp;CBO</td>
	<td class=titulor>&nbsp;Banco</td>
	<td class=titulor>&nbsp;Agência</td>
	<td class=titulor>&nbsp;Conta</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rs2("cpf")%></td>
	<td class="campor">&nbsp;<%=rs2("nit")%></td>
	<td class="campor">&nbsp;<%=rs2("rg")%>&nbsp;<%=rs2("orgao_rg")%></td>
	<td class="campor">&nbsp;<%=rs2("ccm")%></td>
	<td class="campor">&nbsp;<%=rs2("cbo")%></td>
	<td class="campor">&nbsp;<%=rs2("bancocod") & " " & rs2("banconome")%></td>
	<td class="campor">&nbsp;<%=rs2("agencia")%></td>
	<td class="campor">&nbsp;<%=rs2("conta")%></td>
</tr>
</table>

<!-- quadro inicio mudanca-->
<table border="1" bordercolor="#CCCCCC" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><th class=titulo colspan=12>Ficha Financeira</th></tr>
<tr>
	<td class=titulor align="center">Emissao</td>
	<td class=titulor align="center">Pagto</td>
	<td class=titulor align="center">Descrição</td>
	<td class=titulor align="center">Total</td>
	<td class=titulor align="center">INSS</td>
	<td class=titulor align="center">IRRF</td>
	<td class=titulor align="center">ISS/O.</td>
	<td class=titulor align="center">Liquido</td>
	<td class=titulor align="center">RPA</td>
	<td class=titulor align="center">DARF</td>
	<td class=titulor align="center">&nbsp;</td>
	<td class=titulor align="center">&nbsp;</td>
</tr>
<%
rs2.close
if rs.recordcount>0 then
linhas=rs.recordcount
lastyear=year(rs("data_emissao"))
rs.movefirst
inicio=1
do while not rs.eof
total=rs("servico_prestado")+rs("outros_rendimentos")
odesc=rs("desconto_iss")+rs("outros_descontos")
if lastyear<>year(rs("data_emissao")) then
%>
<tr>
	<td class=totalr colspan=3>&nbsp;Total do ano</td>
	<td class=totalr align="right"><%=formatnumber(ttotal,2) %></td>
	<td class=totalr align="right"><%=formatnumber(tinss,2)%></td>
	<td class=totalr align="right"><%=formatnumber(tir,2)%></td>
	<td class=totalr align="right"><%=formatnumber(todesc,2)%></td>
	<td class=totalr align="right"><%=formatnumber(tliq,2)%></td>
	<td class=totalr colspan=4>&nbsp;</td>
</tr>
<%
	ttotal=0:tinss=0:tir=0:todesc=0:tliq=0
end if
%>
<tr>
	<td class="campor"><%=rs("data_emissao") %></td>
	<td class="campor"><%=rs("data_pagamento") %>    </td>
	<td class="campor"><%=rs("descricao_servico") %>   </td>
	<td class="campor" align="right"><%=formatnumber(total,2) %></td>
	<td class="campor" align="right"><%=formatnumber(rs("desconto_inss"),2)%></td>
	<td class="campor" align="right"><%=formatnumber(rs("desconto_ir"),2)%></td>
	<td class="campor" align="right"><%=formatnumber(odesc,2)%></td>
	<td class="campor" align="right"><%=formatnumber(rs("valor_liquido"),2)%></td>
	<td class="campor" align="center">&nbsp;<%if rs("emitiu_rpa")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor" align="center">&nbsp;<%if rs("emitiu_darf")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class="campor" valign=top align="center">
	<% if session("a52")="T" then %>
		<a href="rpa_alteracao.asp?codigo=<%=rs("id_lanc")%>" onclick="NewWindow(this.href,'AlteracaoRPA','510','450','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/folder95o.gif" alt="Alterar este RPA"></a>
	<% end if %>
	</td>
	<td class="campor" valign=top align="center">
	<% if session("a52")<>"N" then %>
		<a href="imprime_rpad.asp?codigo=<%=rs("id_lanc")%>" onclick="NewWindow(this.href,'ReciboRPA','685','350','yes','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/printer.gif" alt="Imprimir este RPA" ></a>
	<% end if %>
	</td>
</tr>
<%
ttotal=ttotal+total              :tgtotal=tgtotal+total
tinss=tinss+rs("desconto_inss")  :tginss=tginss+rs("desconto_inss")
tir=tir+rs("desconto_ir")        :tgir=tgir+rs("desconto_ir")
todesc=todesc+odesc              :tgodesc=tgodesc+odesc
tliq=tliq+rs("valor_liquido")    :tgliq=tgliq+rs("valor_liquido")
lastyear=year(rs("data_emissao"))
rs.movenext
inicio=0
loop
else ' sem registros/planos
%>
<tr><td class="campor" colspan=12>&nbsp;</td></tr>
<%
end if
%>
<tr>
	<td class=totalr colspan=3>&nbsp;Total do ano</td>
	<td class=totalr align="right"><%=formatnumber(ttotal,2) %></td>
	<td class=totalr align="right"><%=formatnumber(tinss,2)%></td>
	<td class=totalr align="right"><%=formatnumber(tir,2)%></td>
	<td class=totalr align="right"><%=formatnumber(todesc,2)%></td>
	<td class=totalr align="right"><%=formatnumber(tliq,2)%></td>
	<td class=totalr colspan=4>&nbsp;</td>
</tr>
<tr>
	<td class=totalr colspan=3>&nbsp;Total Geral</td>
	<td class=totalr align="right"><%=formatnumber(tgtotal,2) %></td>
	<td class=totalr align="right"><%=formatnumber(tginss,2)%></td>
	<td class=totalr align="right"><%=formatnumber(tgir,2)%></td>
	<td class=totalr align="right"><%=formatnumber(tgodesc,2)%></td>
	<td class=totalr align="right"><%=formatnumber(tgliq,2)%></td>
	<td class=totalr colspan=4>&nbsp;</td>
</tr>

</table>
<!-- quadro fim mudanca -->
<table><tr>
<td valign="top">
<% if session("a52")="T" then %>
<a href="rpa_nova.asp?codigo=<%=id_autonomo%>" onclick="NewWindow(this.href,'InclusaoRPA','510','400','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo RPa">
<font size="1">inserir novo RPA</font></a>
<% end if %>
</td>
</tr></table>

</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
set rs3=nothing
conexao.close
set conexao=nothing
%>