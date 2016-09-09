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
<title>Memorando de assistência médica</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<script language="VBScript" runat="server">
	Sub vr1_onChange
		ok=true:dim frm:set frm=document.form
		temp=document.form.vr1.value
		temp2=formatnumber(temp,2)
		Calcula
		document.form.vr1.value=temp2
	End Sub
	Sub vr2_onChange
		ok=true:dim frm:set frm=document.form
		temp=document.form.vr2.value
		temp2=formatnumber(temp,2)
		Calcula
		document.form.vr2.value=temp2
	End Sub
	Sub vr3_onChange
		ok=true:dim frm:set frm=document.form
		temp=document.form.vr3.value
		temp2=formatnumber(temp,2)
		Calcula
		document.form.vr3.value=temp2
	End Sub

	Sub Calcula()
		if document.form.vr1.value<>"" then valor1=cdbl(document.form.vr1.value) else valor1=0
		if document.form.vr2.value<>"" then valor2=cdbl(document.form.vr2.value) else valor2=0
		if document.form.vr3.value<>"" then valor3=cdbl(document.form.vr3.value) else valor3=0
		if document.form.vr4.value<>"" then valor4=cdbl(document.form.vr4.value) else valor4=0
		totalnf=valor1+valor2+valor3+valor4
		'etotal=extenso2(totalnf)
		document.form.totalnf.value=formatnumber(totalnf,2) 
	End Sub

</script></head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
'sqla="SELECT p.carteiratrab, p.seriecarttrab from qry_funcionarios p where chapa='" & request.form("chapacarta") & "' "
'rs.Open sqla, ,adOpenStatic, adLockReadOnly
'rs.close
sc=5189.82.75:inss=570.88

if request.form<>"" then
	if request.form("vr1")<>"" then vr1=cdbl(request.form("vr1")) else vr1=0
	if request.form("vr2")<>"" then vr2=cdbl(request.form("vr2")) else vr2=0
	if request.form("vr3")<>"" then vr3=cdbl(request.form("vr3")) else vr3=0
	if request.form("vr4")<>"" then vr4=cdbl(request.form("vr4")) else vr4=0
	totalmemo=vr1+vr2+vr3+vr4
end if


%>
<form method="POST" name="form">
<%
mes=month(now())
ano=year(now())
emes=monthname(mes)

%>
<!-- <div align="right"> -->
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="620" height="990">
<tr><td colspan=6 class=titulop height=35 valign="middle" align="center" style="border-bottom:2 solid"> M E M O R A N D O&nbsp; &nbsp;I N T E R N O
</td></tr>
<!-- corpo da carta -->
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
	<b>ASSUNTO:</b><br>
	<input type="text" class="form_input10" value="Pagamento de Assistência Médica <%response.write emes & "/" & ano%>" size=80>
	</td>
</tr>
	
	
<tr><td colspan=6 height=800 class="campop" align="left" valign=top>

<p align="left" style="margin-top:0;margin-bottom:0;font-size:12pt">
<br>
<br>
<br>
<p align="justify" style="margin-top:0;margin-bottom:0;font-size:11pt;text-indent:10pt;line-height:150%">
Solicitamos os valores para pagamento das notas fiscais de serviços de assistência médica,
conforme abaixo discriminado, <br> no total de <b>R$ <input name="totalnf" type="text" class="form_input10" value="<%=formatnumber(totalmemo,2)%>" size=80></b>
<br>

<div align="center">
	<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse" width=90%>
	<tr><td class=titulop align="center">Nota Fiscal</td>
		<td class=titulop align="center">Empresa</td>
		<td class=titulop align="center">Valor</td>
		<td class=titulop align="center">Vencimento</td>
	</tr>
	<tr><td class="campop"><input name="nf1"  type="text" class="form_input10" value="<%=request.form("nf1")%>" size=10></td>
		<td class="campop"><input name="emp1" type="text" class="form_input10" value="<%=request.form("emp1")%>" size=30></td>
		<td class="campop"><input name="vr1"  type="text" class="proporcional" value="<%=request.form("vr1")%>" onchange="javascript:submit();" size=15></td>
		<td class="campop"><input name="vc1"  type="text" class="form_input10" value="<%=request.form("vc1")%>" size=15></td>
	</tr>
	<tr><td class="campop"><input name="nf2"  type="text" class="form_input10" value="<%=request.form("nf2")%>" size=10></td>
		<td class="campop"><input name="emp2" type="text" class="form_input10" value="<%=request.form("emp2")%>" size=30></td>
		<td class="campop"><input name="vr2"  type="text" class="proporcional" value="<%=request.form("vr2")%>" onchange="javascript:submit();" size=15></td>
		<td class="campop"><input name="vc2"  type="text" class="form_input10" value="<%=request.form("vc2")%>" size=15></td>
	</tr>
	<tr><td class="campop"><input name="nf3"  type="text" class="form_input10" value="<%=request.form("nf3")%>" size=10></td>
		<td class="campop"><input name="emp3" type="text" class="form_input10" value="<%=request.form("emp3")%>" size=30></td>
		<td class="campop"><input name="vr3"  type="text" class="proporcional" value="<%=request.form("vr3")%>" onchange="javascript:submit();" size=15></td>
		<td class="campop"><input name="vc3"  type="text" class="form_input10" value="<%=request.form("vc3")%>" size=15></td>
	</tr>
	<tr><td class="campop"><input name="nf4"  type="text" class="form_input10" value="<%=request.form("nf4")%>" size=10></td>
		<td class="campop"><input name="emp4" type="text" class="form_input10" value="<%=request.form("emp4")%>" size=30></td>
		<td class="campop"><input name="vr4"  type="text" class="proporcional" value="<%=request.form("vr4")%>" onchange="javascript:submit();" size=15></td>
		<td class="campop"><input name="vc4"  type="text" class="form_input10" value="<%=request.form("vc4")%>" size=15></td>
	</tr>
	</table>
</div>
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
<!-- </div> -->








</form>
<%
'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>