<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a95")="N" or session("a95")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Op��o Formul�rios</title>
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
function nome1() { form.chapa.value=form.nome.value;form.submit(); }
function chapa1() { form.nome.value=form.chapa.value;form.submit(); }

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
--></script>

<%
Function IIf(condition,value1,value2)
	If condition Then IIf = value1 Else IIf = value2
End Function

dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request.form("B1")="" or request.form("form_id")="" then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Sele��o de funcion�rio e par�metros
<form method="POST" action="formularios.asp" name="form">
<%
sqla="SELECT f.chapa, f.NOME FROM corporerm.dbo.pfunc AS f " & _
"WHERE f.CODSITUACAO<>'D' and codtipo='N' ORDER BY f.NOME;"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=request.form("chapa")%>">
<select name="nome" class=a onchange="nome1()">
	<option value="00000">===> Ficha em branco <===</option>
<%
rs.movefirst:do while not rs.eof
if request.form("chapa")=rs("chapa") then temps="selected" else temps=""
%>
	<option value="<%=rs("chapa")%>" <%=temps%> ><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<br><br>
	<table border="1" bordercolor=#00000 cellpadding="0" cellspacing="0" style="border-collapse: collapse"><tr><td>
<table border="0" cellpadding="2" cellspacing="2" style="border-collapse: collapse" width=400>
<tr>
	<td class=titulop colspan=2>Data de emiss�o</td>
</tr>
<%
sql0="SELECT dataadmissao FROM corporerm.dbo.pfunc f where f.chapa='" & request.form("chapa") & "' ;"
rs.Open sql0, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	dataadmissao=rs("dataadmissao")
end if
rs.close
%>
<tr><td class=campo nowrap><input type="radio" name="emissao" value="A" checked> Data de Admissao</td><td class=campo width=95%><%=dataadmissao%><input type=hidden name="dataA" value="<%=dataadmissao%>"></td></tr>
<tr><td class=campo><input type="radio" name="emissao" value="H"> Data de Hoje            </td><td class=campo><%=formatdatetime(now(),2)%></td></tr>
<tr><td class=campo><input type="radio" name="emissao" value="D"> Data espec�fica         </td><td class=campo><input type="text" size=8 name="dataD"></td></tr>
<%
%>
</table>
	</td></tr></table>
<Br>
	<table border="1" bordercolor=#00000 cellpadding="0" cellspacing="0" style="border-collapse: collapse"><tr><td>
<table border="0" cellpadding="2" cellspacing="2" style="border-collapse: collapse" width=400>
<tr>
	<td class=fundo><input type="checkbox" name="checkall" onclick="toggleAll(this)" id="Checkbox1" />  </td>
	<td class=titulo>Formul�rio</td>
	<td class=titulo>Vias</td>
</tr>
<tr><td class="campop"><input type="checkbox" name="form_id" value="id0"></td><td class="campop">Termo de responsabilidade (Sal�rio-Fam�lia)</td><td class="campop"><input type="text" size="1" class="form_input10" name="via0" value="1"></td></tr>
<tr><td class="campop"><input type="checkbox" name="form_id" value="id1"></td><td class="campop">Recibo/Comprovante de entrega da CTPS      </td><td class="campop"><input type="text" size="1" class="form_input10" name="via1" value="1"></td></tr>
<tr><td class="campop"><input type="checkbox" name="form_id" value="id2"></td><td class="campop">Declara��o de Encargos p/Imposto de Renda  </td><td class="campop"><input type="text" size="1" class="form_input10" name="via2" value="1"></td></tr>
<tr><td class="campop"><input type="checkbox" name="form_id" value="id3"></td><td class="campop">Declara��o de Op��o de Vale-Transporte     </td><td class="campop"><input type="text" size="1" class="form_input10" name="via3" value="1"></td></tr>
<tr><td class="campop"><input type="checkbox" name="form_id" value="id4"></td><td class="campop">Contrato de Experi�ncia                    </td><td class="campop"><input type="text" size="1" class="form_input10" name="via4" value="2"></td></tr>
<tr><td class="campop"><input type="checkbox" name="form_id" value="id5"></td><td class="campop">Acordo de Compensa��o de Horas             </td><td class="campop"><input type="text" size="1" class="form_input10" name="via5" value="1"></td></tr>
<tr><td class="campop"><input type="checkbox" name="form_id" value="id6"></td><td class="campop">Ficha de Registro                          </td><td class="campop"><input type="text" size="1" class="form_input10" name="via6" value="1"></td></tr>
<tr><td class="campop"><input type="checkbox" name="form_id" value="id7"></td><td class="campop">Termo Internet                             </td><td class="campop"><input type="text" size="1" class="form_input10" name="via7" value="1"></td></tr>
<tr><td class="campop"><input type="checkbox" name="form_id" value="id8"></td><td class="campop">Op��o Assist�ncia M�dica                   </td><td class="campop"><input type="text" size="1" class="form_input10" name="via8" value="1"></td></tr>
<tr><td class="campop"><input type="checkbox" name="form_id" value="id9"></td><td class="campop">Op��o Cesta B�sica/V.Alimenta��o           </td><td class="campop"><input type="text" size="1" class="form_input10" name="via9" value="1"></td></tr>

</table>
	</td></tr></table>
<%
if session("usuariomaster")="02379" or session("usuariomaster")="02675" then
	response.write "<p>Data especial para relatorio <input type='text' size=8 name='dataInter'></p>"
end if
%>
<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>
<%
end if

if request.form("B1")<>"" and request.form("form_id")<>"" then
'temp=request.form("id")
'tipo=left(temp,1)
'codigo=right(temp,len(temp)-1)
chapa=request.form("chapa")
tamanho=32:tamanho2=30
largura1=650:largfoto=150
corborda="#009999"
for a=1 to 30:espacao=espacao & "&nbsp;":next

select case request.form("emissao")
	case "A"
		datarel=request.form("dataA")
	case "H"
		datarel=formatdatetime(now(),2)
	case "D"
		datarel=request.form("dataD")
end select

'***************************
'** INICIO DO FORMUL�RIO  **
'***************************
for a=1 to request.form("form_id").count
	if request.form("form_id").item(a)="id0" then formulario0="S"
next
if formulario0="S" then 'termo de responsabilidade
sqla="select chapa, nome, carteiratrab ctps, seriecarttrab serie from qry_funcionarios where chapa='" & chapa & "'"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
end if

for v=1 to request.form("via0")
%>
<center>
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690" height=900>
<tr>
	<td class=campo valign=top>
		<p align="center" style="margin-bottom:0px;margin-top:5px;font-size:16px">TERMO DE RESPONSABILIDADE</p>
		<p align="center" style="margin-bottom:5px;margin-top:0px;font-size:12px">(CONCESS�O DE SAL�RIO-FAM�LIA - PORTARIA N� MPAS - 3.040/82)</p>
<br>

<div align="center">
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class="campor" height=40><font style="font-size:9px">&nbsp;Empresa<br>
		<font style="font-size:14px">&nbsp;FUNDA��O INSTITUTO DE ENSINO PARA OSASCO
	</td>
	<td class="campor"><font style="font-size:9px">&nbsp;Matricula<br>
		<font style="font-size:14px">&nbsp;73.063.166
	</td>
</tr>
<tr>
	<td class="campor" colspan=2 height=40><font style="font-size:9px">&nbsp;Nome do Segurado<br>
		<font style="font-size:14px">&nbsp;<b><%=rs("nome")%></b> (<%=rs("chapa")%>)
	</td>
</tr>
<tr>
	<td class="campor" colspan=2 height=40><font style="font-size:9px">&nbsp;CTPS ou doc.identidade<br>
		<font style="font-size:14px">&nbsp;<%=rs("ctps") & "/" & rs("serie")%>
	</td>
</tr>
</table>
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class="campor" height=30 width=30 valign=middle align="center" rowspan=11><font style="font-size:11px">B<br>E<br>N<br>E<br>F<br>I<br>C<br>I<br>�<br>R<br>I<br>O<br>S</td>
	<td class="campor" height=30 width=450 align="center"><font style="font-size:12px">NOME DO FILHO</td>
	<td class="campor" height=30 align="center"><font style="font-size:12px">DATA DO NASCIMENTO</td>
</tr>
<%
sqld="select nome, dtnascimento from corporerm.dbo.pfdepend " & _
"where chapa='" & chapa & "' and grauparentesco in (1,3) and dateadd(yy,14,dtnascimento)>='" & dtaccess(datarel) & "' order by dtnascimento"
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
total=rs2.recordcount
do while not rs2.eof
%>
<tr>
	<td class="campor" height=30 ><font style="font-size:12px">&nbsp;<%=rs2("nome")%></td>
	<td class="campor" height=30 align="center"><font style="font-size:12px"><%=rs2("dtnascimento")%></td>
</tr>
<%
rs2.movenext
loop
rs2.close
if total<10 then
	for a=1 to (10-total)
%>
<tr>
	<td class="campor" height=30>&nbsp;</td>
	<td class="campor" height=30></td>
</tr>
<%
	next 
end if
%>
</table>

<!-- inicio texto -->
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td>
<p style="margin-bottom:0px;margin-top:5px;font-size:13px;text-align:justify;line-height:17px">
Pelo presente TERMO DE RESPONSABILIDADE declaro estar ciente de que deverei comunicar de imediato a ocorr�ncia dos seguintes fatos
ou ocorr�ncias que determinam a perda do direito ao sal�rio-fam�lia:
<br>- �BITO DE FILHO;
<br>- CESSA��O DA INVALIDEZ DE FILHO INV�LIDO;
<br>- SENTEN�A JUDICIAL QUE DETERMINE O PAGAMENTO A OUTREM (casos de desquite ou separa��o, abandono de filho ou perda do p�trio poder).
<br>
<br>Estou ciente, ainda de que a falta de cumprimento ora assumido, al�m de obrigar � devolu��o das import�ncias recebidas indevidamente, sujeitar-me-a
�s penalidades previstas no art. 171 do C�digo Penal e � rescis�o do contrato de trabalho, por justa causa, nos termos do art. 482 da Consolida��o das
Leis do Trabalho.
<br>
<br>Para filhos <u>menores do que 6 anos</u>, apresentarei a caderneta de vacina��o, e para filhos <u>de 7 a 14 anos</u> apresentarei comprovante de frequ�ncia escolar todo m�s de fevereiro e agosto.
</td></tr></table>
<!-- fim texto -->
<br>
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class="campor" height=60 width=520 valign=top><font style="font-size:9px">&nbsp;Local e data<br>
		<font style="font-size:14px"><br>&nbsp;Osasco, <%=datarel%>
	</td>
	<td class="campor" rowspan=2 height=60 valign=top><font style="font-size:9px">&nbsp;Impress�o Digital<br>
		<font style="font-size:14px">&nbsp;
	</td>
</tr>
<tr>
	<td colspan=2class="campor" colspan=2 height=50 valign=top><font style="font-size:9px">&nbsp;Assinatura<br>
		<font style="font-size:14px">&nbsp;
	</td>
</tr>
</table>
</div>

	</td>
</tr>
</table>

<div align="center">
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td>
<p style="margin-bottom:0px;margin-top:0px;font-size:9px;text-align:justify;">
1� via - empresa<br>2� via - segurado
</td></tr></table>
</div>

<%
'if request.form("via0")>1 and v<request.form("via0") then 
response.write "<DIV style=""page-break-after:always""></DIV>"
next 'via0
rs.close
end if 'id0

'***************************
'** INICIO DO FORMUL�RIO  **
'***************************
'if request.form("form_id").count>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
for a=1 to request.form("form_id").count
	if request.form("form_id").item(a)="id1" then formulario1="S"
next
'if formulario1="S" and (formulario0="S") then response.write "<DIV style=""page-break-after:always""></DIV>"
if formulario1="S" then 'recibo ctps

sqla="select chapa, nome, carteiratrab ctps, seriecarttrab serie, dtcarttrab emissao from qry_funcionarios where chapa='" & chapa & "'"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
end if

for v=1 to request.form("via1")
%>
<center>
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;border:1 dotted" width="690" height=900>
<tr>
	<td class=campo valign=top>
<br>

<div align="center">
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campop" height=50 colspan=3 style="font-size:15px;border:1px solid" align="center">RECIBO DE ENTREGA DA CARTEIRA DE TRABALHO<br>
		E PREVID�NCIA SOCIAL PARA ANOTA��ES
	</td>
</tr>
<tr><td class="campor" height=40 style="border:1px solid"><font style="font-size:9px">&nbsp;CTPS n�<br>
		<font style="font-size:14px">&nbsp;<%=rs("ctps")%>	</td>
	<td class="campor" style="border:1px solid"><font style="font-size:9px">&nbsp;S�RIE<br>
		<font style="font-size:14px">&nbsp;<%=rs("serie")%>	</td>
	<td class="campor" style="border:1px solid"><font style="font-size:9px">&nbsp;EMISS�O<br>
		<font style="font-size:14px">&nbsp;<%=rs("emissao")%>	</td>
</tr>
<tr><td class="campor" colspan=3 height=40 style="border:1px solid"><font style="font-size:9px">&nbsp;NOME DO EMPREGADO<br>
		<font style="font-size:14px">&nbsp;<%=rs("nome")%>
	</td>
</tr>
<tr><td class="campop" style="font-size:14px;border:1px solid" colspan=3 height=290 valign=top>
		<br>
		&nbsp;Recebemos a Carteira de Trabalho e Previd�ncia Social acima, anota��es necess�rias e que ser� devolvida dentro de 48 (quarenta e oito)
		horas, de acordo com a legisla��o em vigor.
		<br>
		<br>
		<br>
		<font style="font-size:14px"><br><%=espacao%>Osasco, <%=day(datarel)%> de <%=monthname(month(datarel))%> de <%=year(datarel)%>
		<br>
		<br>
		<br>
		<br><%=espacao%>___________________________________________________________
		<br><%=espacao%>FUNDA��O INSTITUTO DE ENSINO PARA OSASCO
	</td>
</tr>

<tr><td height=8 colspan=3 style="border-bottom:1px dotted #000000"></td></tr>
<tr><td height=8 colspan=3 style="border:0px solid"></td></tr>

<tr><td class="campop" height=50 colspan=3 style="font-size:15px;border:1px solid" align="center">COMPROVANTE DE DEVOLU��O DA CARTEIRA DE TRABALHO<br>
		E PREVID�NCIA SOCIAL
	</td>
</tr>
<tr><td class="campor" height=40 style="border:1px solid"><font style="font-size:9px">&nbsp;CTPS n�<br>
		<font style="font-size:14px">&nbsp;<%=rs("ctps")%>	</td>
	<td class="campor" style="border:1px solid"><font style="font-size:9px">&nbsp;S�RIE<br>
		<font style="font-size:14px">&nbsp;<%=rs("serie")%>	</td>
	<td class="campor" style="border:1px solid"><font style="font-size:9px">&nbsp;EMISS�O<br>
		<font style="font-size:14px">&nbsp;<%=rs("emissao")%>	</td>
</tr>
<tr><td class="campor" colspan=3 height=40 style="border:1px solid"><font style="font-size:9px">&nbsp;NOME DO EMPREGADO<br>
		<font style="font-size:14px">&nbsp;<%=rs("nome")%>
	</td>
</tr>
<tr><td class="campop" style="font-size:14px;border:1px solid" colspan=3 height=290 valign=top>
		<br>
		&nbsp;Recebi, em devolu��o, a Carteira de Trabalho e Previd�ncia Social com as respectivas anota��es.
		<br>
		<br>
		<br>
		<font style="font-size:14px"><br><%=espacao%>Osasco, <%=day(datarel)%> de <%=monthname(month(datarel))%> de <%=year(datarel)%>
		<br>
		<br>
		<br>
		<br><%=espacao%>___________________________________________________________
		<br><%=espacao%><%=rs("nome")%>
	</td>
</tr>
</table>

</div>
	</td>
</tr>
</table>

<%
'if request.form("via1")>1 and v<request.form("via1") then 
response.write "<DIV style=""page-break-after:always""></DIV>"
next 'via1
rs.close
end if 'id1

'***************************
'** INICIO DO FORMUL�RIO  **
'***************************
'if request.form("form_id").count>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
for a=1 to request.form("form_id").count
	if request.form("form_id").item(a)="id2" then formulario2="S"
next
'if formulario2="S" and (formulario1="S" or formulario0="S") then response.write "<DIV style=""page-break-after:always""></DIV>"
if formulario2="S" then 'declara��o de encargos
sqla="select chapa, nome, carteiratrab ctps, seriecarttrab serie, cpf, estcivil, rua, numero, complemento, bairro, cidade, estado, cep " & _
"from qry_funcionarios where chapa='" & chapa & "'"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
end if

for v=1 to request.form("via2")
%>
<center>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690" height=900>
<tr>
	<td class=campo valign=top>
<br>

<div align="center">
	
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class=campo height=40 align="center" style="font-size:14px;border:1px solid #000000"><b>DECLARA��O DE ENCARGOS DE FAM�LIA<br>PARA FINS DE IMPOSTO DE RENDA</td>
</tr>	
<tr><td height=10></td></tr>
<tr>
	<td class=campo height=30 align="left" style="font-size:12px;border:1px solid #000000"><b>&nbsp;� FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</td>
</tr>	
<tr><td height=10></td></tr>
</table>

<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td>
<p style="margin-bottom:0px;margin-top:5px;font-size:13px;text-align:justify;line-height:17px">
Nos termos da legisla��o do Imposto de Renda, venho pela presente informar-lhe que tenho como encargo(s) de fam�lia, a(s) pessoa(s) abaixo
relacionada(s):
</td></tr>
<tr><td height=5></td></tr>
</table>

<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class="campop" height=30 align="center" valign=middle>N� de<br>ordem</td>
	<td class="campop" height=30 align="center" valign=middle>Nome Completo</td>
	<td class="campop" height=30 align="center" valign=middle>Rela��o de<br>Depend�ncia</td>
	<td class="campop" height=30 align="center" valign=middle>Data de<br>Nascimento</td>
</tr>
<%
sqld="select chapa, nome, dtnascimento, grauparentesco, p.descricao " & _
"from corporerm.dbo.pfdepend d inner join corporerm.dbo.pcodparent p on p.codcliente=d.grauparentesco " & _
"where incirrf=1 and chapa='" & chapa & "' order by dtnascimento "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
total=rs2.recordcount
do while not rs2.eof
%>
<tr>
	<td class="campop" height=30 align="center"><font style="font-size:12px">&nbsp;<%=rs2.absoluteposition%></td>
	<td class="campop" height=30 ><font style="font-size:12px">&nbsp;<%=rs2("nome")%></td>
	<td class="campop" height=30 ><font style="font-size:12px">&nbsp;<%=rs2("descricao")%></td>
	<td class="campop" height=30 align="center"><font style="font-size:12px"><%=rs2("dtnascimento")%></td>
</tr>
<%
rs2.movenext
loop
rs2.close
if total<7 then
	for a=1 to (7-total)
%>
<tr>
	<td class=campo height=30>&nbsp;</td>
	<td class=campo height=30></td>
	<td class=campo height=30></td>
	<td class=campo height=30></td>
</tr>
<%
	next 
end if
%>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td colspan=2>
<p style="margin-bottom:0px;margin-top:5px;font-size:13px;text-align:justify;line-height:17px">
Declaro sob as penas da lei, que as informa��es aqui prestadas s�o verdadeiras e de minha inteira responsabilidade.
</td></tr>
<tr><td height=5></td></tr>
<tr><td class=campo>
	<font style="font-size:12px">Data
	<br><br>_____/_____/_____
</td><td class=campo>
	<font style="font-size:12px">Assinatura
	<br><br>__________________________________
</td></tr>
<tr><td height=15></td></tr>
</table>


<!-- inicio texto -->
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="<%=largura1%>">
<tr><td class="campor" height=40 style="border-top:1px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;Nome do declarante<br>
		<font style="font-size:14px">&nbsp;<%=rs("nome")%>	</td>
	<td class="campor" style="border-top:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;N� Registro<br>
		<font style="font-size:14px">&nbsp;<%=rs("chapa")%>	</td>
</tr>
</table>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" height=40 style="border-left:1px solid"><font style="font-size:9px">&nbsp;Endere�o<br>
		<font style="font-size:14px">&nbsp;<%=rs("rua")%>	</td>
	<td class="campor" style="border:0px solid"><font style="font-size:9px">&nbsp;N�<br>
		<font style="font-size:14px">&nbsp;<%=rs("numero")%>	</td>
	<td class="campor" style="border-right:1px solid"><font style="font-size:9px">&nbsp;Complemento<br>
		<font style="font-size:14px">&nbsp;<%=rs("complemento")%>	</td>
</tr>
</table>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" height=40 style="border-left:1px solid"><font style="font-size:9px">&nbsp;Bairro<br>
		<font style="font-size:14px">&nbsp;<%=rs("bairro")%>	</td>
	<td class="campor" style="border:0px solid"><font style="font-size:9px">&nbsp;Cidade<br>
		<font style="font-size:14px">&nbsp;<%=rs("cidade")%>	</td>
	<td class="campor" style="border:0px solid"><font style="font-size:9px">&nbsp;Estado<br>
		<font style="font-size:14px">&nbsp;<%=rs("estado")%>	</td>
	<td class="campor" style="border-right:1px solid"><font style="font-size:9px">&nbsp;CEP<br>
		<font style="font-size:14px">&nbsp;<%=rs("CEP")%>	</td>
</tr>
</table>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" height=40 style="border-left:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;Estado Civil<br>
		<font style="font-size:14px">&nbsp;<%=rs("estcivil")%>	</td>
</tr>
</table>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" height=40 style="border-bottom:1px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;CTPS n�<br>
		<font style="font-size:14px">&nbsp;<%=rs("ctps")%>	</td>
	<td class="campor" style="border-bottom:1px solid"><font style="font-size:9px">&nbsp;S�rie<br>
		<font style="font-size:14px">&nbsp;<%=rs("serie")%>	</td>
	<td class="campor" style="border-bottom:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;CPF n�<br>
		<font style="font-size:14px">&nbsp;<%=rs("cpf")%>	</td>
</tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td>
<p style="margin-bottom:0px;margin-top:5px;font-size:13px;text-align:justify;line-height:17px">
Anexar (c�pias simples):
<ul type="disc" style="margin-bottom:1px;margin-top:5px">
	<li>Certid�o de casamento;</li>
	<li>Certid�o de Nascimento dos filhos;</li>
	<li>Senten�a judicial que determine a guar de menor.</li>
</ul>
</td></tr>
<tr><td height=5></td></tr>
</table>

<!-- fim texto -->
</div>

	</td>
</tr>
</table>

<div align="center">
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td>
<p style="margin-bottom:0px;margin-top:0px;font-size:9px;text-align:justify;">
1� via - empresa
</td></tr></table>
</div>

<%
'if request.form("via2")>1 and v<request.form("via2") then 
response.write "<DIV style=""page-break-after:always""></DIV>"
next 'via2
rs.close
end if 'id2

'***************************
'** INICIO DO FORMUL�RIO  **
'***************************
'if request.form("form_id").count>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
for a=1 to request.form("form_id").count
	if request.form("form_id").item(a)="id3" then formulario3="S"
next
'if formulario3="S" and (formulario2="S" or formulario1="S" or formulario0="S") then response.write "<DIV style=""page-break-after:always""></DIV>"
if formulario3="S" then 'declara��o de op��o de vt
sqla="select chapa, nome, cpf, rua, numero, complemento, bairro, cidade, estado, cep " & _
"from qry_funcionarios where chapa='" & chapa & "'"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
end if

for v=1 to request.form("via3")
%>
<center>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690" height=990>
<tr>
	<td class=campo valign=top>

<div align="center">
	
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class=campo height=40 align="center" style="font-size:16px;border:1px solid #000000"><b>Declara��o de Op��o do Vale-Transporte</td>
</tr>
<tr>
	<td class="campop" style="font-size:12px;border:1px solid #000000"><p style="margin:5px">Pela presente venho declarar a minha op��o relativa � antecipa��o do benef�cio citado na Lei 7.418 de 16/12/1985, alterada pela 
		Lei 7.618 de 30/09/1987.
	</td>
</tr>
<tr>
	<td class="campop" style="font-size:12px;border:1px solid #000000"><p style="margin:5px">
	<b>[&nbsp;&nbsp;&nbsp;] N�o desejo receber o Vale-Transporte.</b>
	<br>
	<br>
	Osasco, ________ de _______________________ de ____________
	<br>
	<br>
	<br>
	___________________________________________________________<br>
	<%=rs("nome")%>
	</td>
</tr>

<tr><td height=5></td></tr>

<tr>
	<td class="campop" style="font-size:12px;border:1px solid #000000"><p style="margin:5px">
	<b>[&nbsp;&nbsp;&nbsp;] Desejo receber o Vale-Transporte.</b>
	<br>
	<p style="margin:5px;text-align:center"><b>Declara��o de Deslocamento</b>
	<p style="margin:5px;text-align:justify;line-height:19px">
	<%=rs("nome")%>, portador do C.P.F. n� <%=rs("cpf")%>, residente � <%=rs("rua")%>, <%=rs("numero")%> <%=rs("complemento")%> - <%=rs("bairro")%>, na cidade 
	de <%=rs("cidade")%>, necessito dos vales-transportes abaixo relacionados e declaro que &nbsp;<u>utilizarei exclusivamente para o deslocamento resid�ncia/trabalho</u> 
	e vice-versa, sujeitando-me as penalidades previstas em lei.<br>
	Autorizo, o desconto de 6% do meu sal�rio para participar como benefici�rio do Programa de Vale Transporte.
	</td>
</tr>
<tr><td height=5></td></tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="<%=largura1%>">
<tr>
	<td class=fundop height=239 width=15 align="center" valign=middle style="border:1px solid"><b>I<br>D<br>A
	
	</td>
	<td class="campop" rowspan=2 style="border-top:1px solid">
<%for a=1 to 6
	
%>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="635">
<tr><td class="campor" height=40 valign="top" style="border-right:1px solid;border-bottom:1px solid"><font style="font-size:9px">&nbsp;N�MERO/NOME DA LINHA<br>&nbsp;</td>
	<td class="campor" height=40 width=45 valign="top" style="border-right:1px solid;border-bottom:1px solid"><font style="font-size:9px">&nbsp;TIPO<br>&nbsp;</td>
	<td class="campor" height=40 width=150 valign="top" style="border-right:1px solid;border-bottom:1px solid"><font style="font-size:9px">&nbsp;EMPRESA<br>&nbsp;</td>
	<td class="campor" height=40 width=120 valign="top" style="border-right:1px solid;border-bottom:1px solid"><font style="font-size:9px">&nbsp;VALOR DA PASSAGEM<br>&nbsp;</td>
</tr>
</table>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="635">
<tr><td class="campor" height=40 valign="top" style="border-right:1px solid;border-bottom:2px solid"><font style="font-size:9px">&nbsp;LOCAL DE EMBARQUE<br>&nbsp;</td>
	<td class="campor" height=40 valign="top" style="border-right:1px solid;border-bottom:2px solid"><font style="font-size:9px">&nbsp;LOCAL DE DESEMBARQUE<br>&nbsp;</td>
</tr>
</table>
<%next%>
	</td>
</tr>
<tr>
	<td class=fundop height=240 width=15 align="center" valign=middle  style="border:1px solid"><b>V<br>O<br>L<br>T<br>A
	
	</td>
</tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="<%=largura1%>">
<tr>
	<td class=fundop>TIPO: T=CPTM | M=METRO | S=SPTRANS | I=INTEGRA��O | B=BOM | E=BEM | F=BENFICA | O=OUTROS
	
	</td>
</tr>
<tr>
	<td class="campop" style="font-size:12px;border-left:1px solid;border-right:1px solid;border-bottom:1px solid"><p style="margin:5px">
	Osasco, ________ de _______________________ de ____________
	<br>
	<br>
	<br>
	___________________________________________________________<br>
	<%=rs("nome")%>
	</td>
</tr>
</table>

</div>

	</td>
</tr>
</table>

<div align="center">
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td>
<p style="margin-bottom:0px;margin-top:0px;font-size:9px;text-align:justify;">
1� via - empresa
</td></tr></table>
</div>

<%
'if request.form("via3")>1 and v<request.form("via3") then 
response.write "<DIV style=""page-break-after:always""></DIV>"
next 'via3
rs.close
end if 'id3

'***************************
'** INICIO DO FORMUL�RIO  **
'***************************
'if request.form("form_id").count>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
for a=1 to request.form("form_id").count
	if request.form("form_id").item(a)="id4" then formulario4="S"
next
'if formulario4="S" and (formulario3="S" or formulario2="S" or formulario1="S" or formulario0="S") then response.write "<DIV style=""page-break-after:always""></DIV>"
if formulario4="S" then 'contrato de experiencia
sqla="select chapa, nome, carteiratrab ctps, seriecarttrab serie, cpf, f.funcao, f.rua, f.numero, f.complemento, f.bairro, f.cidade, f.estado, f.cep, cnpj=s.cgc, " & _
"frua=s.rua, fnumero=s.numero, fbairro=s.bairro, f.sexo, f.codsindicato " & _
"from qry_funcionarios f inner join corporerm.dbo.psecao s on s.codigo=f.codsecao where chapa='" & chapa & "'"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
end if
if rs("sexo")="F" then f1="a" else f1=""
if rs("sexo")="F" then f2="a" else f2="o"

if rs("codsindicato")="03" then
	pagto="por hora aula" 
	tipo="F"
	dias=90:edias="noventa"
	dias=45:edias="quarenta e cinco"
else 
	pagto="por m�s"
	tipo="A"
	dias=45:edias="quarenta e cinco"
end if

if request.form("emissao")="H" and rs("codsindicato")<>"03" then sqls="select salario, jornadamensal/60.00 jornada from corporerm.dbo.pfunc where chapa='" & chapa & "'"
if request.form("emissao")="A" and rs("codsindicato")<>"03" then sqls="select salario, jornada/60.00 jornada from corporerm.dbo.pfhstsal where chapa='" & chapa & "'  and dtmudanca between '" & dtaccess(datarel) & "' and dateadd(""hh"",23,'"& dtaccess(datarel) & "') "
if request.form("emissao")="D" and rs("codsindicato")<>"03" then sqls="select salario, jornada/60.00 jornada from corporerm.dbo.pfhstsal h where chapa='" &  chapa & "' and dtmudanca=(select top 1 dtmudanca from corporerm.dbo.pfhstsal where chapa='" & chapa & "' and dtmudanca<='" & dtaccess(datarel) & "' order by dtmudanca desc)"

if request.form("emissao")="H" and rs("codsindicato")="03" then sqls="select salario=salario/(jornada/60.00), jornadamensal/60.00 jornada from corporerm.dbo.pfunc where chapa='" & chapa & "'"
if request.form("emissao")="A" and rs("codsindicato")="03" then sqls="select salario=salario/(jornada/60.00), jornada/60.00 jornada from corporerm.dbo.pfhstsal where chapa='" & chapa & "'  and dtmudanca between '" & dtaccess(datarel) & "' and dateadd(""hh"",23,'"& dtaccess(datarel) & "') "
if request.form("emissao")="D" and rs("codsindicato")="03" then sqls="select salario=salario/(jornada/60.00), jornada/60.00 jornada from corporerm.dbo.pfhstsal h where chapa='" &  chapa & "' and dtmudanca=(select top 1 dtmudanca from corporerm.dbo.pfhstsal where chapa='" & chapa & "' and dtmudanca<='" & dtaccess(datarel) & "' order by dtmudanca desc)"
'response.write datarel & " - " & request.form("emissao") & "<br>" & sqls
rs2.Open sqls, ,adOpenStatic, adLockReadOnly
salario=rs2("salario")
jornada=rs2("jornada")
rs2.close

for v=1 to request.form("via4")
%>
<center>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse:collapse; border:0px dotted" width="690" height=990>
<tr>
	<td class=campo valign=top>

<div align="center">
	
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class=fundo height=25 align="center" style="font-size:14px;border:1px solid #000000"><b>CONTRATO DE TRABALHO A TITULO DE EXPERI�NCIA</td>
</tr>
<tr>
	<td class="campop" style="font-size:12px;border:0 solid">
	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	Entre a empresa <b>FIEO-FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</b>, com sede em Osasco-SP � <%=rs("frua")%>, <%=rs("fnumero")%> - <%=rs("fbairro")%>, doravante
	designada simplesmente EMPREGADORA e <b><%=rs("nome")%></b>, portador<%=f1%> da Carteira de Trabalho e Previd�ncia Social n� <%=rs("ctps")%> s�rie <%=rs("serie")%>, a seguir
	chamad<%=f2%> de apenas EMPREGADO, � celebrado o presente CONTRATO DE EXPERI�NCIA, que ter� vig�ncia a partir da data de in�cio de presta��o de servi�o, de acordo
	com as condi��es a seguir especificadas:
	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	1 - Fica o EMPREGADO admitido no quadro de funcion�rios da EMPREGADORA para exercer as fun��es de <b><%=rs("funcao")%></b>, mediante a remunera��o de 
	<b>R$ <%=formatnumber(salario,2)%></b> (<%=extenso2(cdbl(salario))%>) <%=pagto%>. A circunst�ncia, por�m, de ser a fun��o especificada n�o importa a intransferibilidade 
	do EMPREGADO para outros servi�os, no qual demonstre melhor capacidade de adapta��o desde que compat�vel com sua condi��o pessoal.
<%if tipo="A" then %>
	<dl style="margin-top:0px;margin-bottom:0px;font-size:12px;text-align:justify">
 	<dt></dt>
	<dd>1.1 - A Jornada de trabalho a ser cumprida � de <%=jornada%> horas por m�s. O hor�rio a ser cumprido ser� regulamentado em documento pr�prio a parte.</dd>
	</dl>
<%end if%>
	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	2 - O hor�rio de trabalho ser� anotado em sua ficha de registro e a eventual redu��o da jornada, por determina��o da EMPREGADORA, n�o inovar� este ajuste, permanecendo
	sempre �ntegra a obriga��o do EMPREGADO de cumprir o hor�rio que lhe for determinado, observando o limite legal.
	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	3 - Obriga-se tamb�m o EMPREGADO a prestar servi�os em horas extraordin�rias, sempre que lhe for determinado pela EMPREGADORA. O EMPREGADO receber� as horas
	extraordin�rias com o acr�scimo legal, salvo a ocorr�ncia de compensa��o, com a consequente redu��o da jornada de trabalho em outro dia.
<%if tipo="A" then %>
	<dl style="margin-top:0px;margin-bottom:0px;font-size:12px;text-align:justify">
 	<dt></dt>
	<dd>3.1 - A perman�ncia do EMPREGADO fora do hor�rio regular de trabalho, sem a devida autoriza��o, ensejar� a aplica��o de san��es disciplinares na forma da lei.</dd>
	</dl>
<%end if%>
	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	4 - Aceita o EMPREGADO, expressamente, a condi��o de prestar servi�os em qualquer dos turnos de trabalho, isto �, durante o dia como a noite, desde que sem
	simultaneidade, observado as prescri��es do assunto, quanto � remunera��o.

	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	5 - Fica ajustado nos termos do que disp�e o � 1� do artigo 469 da Consolida��o das Leis do Trabalho, que o EMPREGADO acatar� ordem emanada da EMPREGADORA para a
	presta��o de servi�os tanto na localidade de celebra��o do Contrato de Trabalho, como em qualquer outra Cidade, Capital ou Vila do Territ�rio Nacional, quer essa
	transfer�ncia seja transit�ria, quer seja definitiva.

	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	6 - Em caso de dano causado pelo EMPREGADO, fica a EMPREGADORA, autorizada a efetivar o desconto da import�ncia correspondente ao preju�zo, no qual far�, com 
	fundamento no � 1� do artigo 462 da Consolida��o das Leis do Trabalho, j� que essa possibilidade fica expressamente prevista em Contrato.

	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	7 - A justificativa de aus�ncia do EMPREGADO, deve observar a ordem preferencial dos atestados m�dicos estabelecida pelo Decreto n� 27.048 de 12/08/49 - art. 12, �� 
	1� e 2�, que regulamentou a Lei n� 605/49, conforme segue:
	<ol type="a" style="margin-top:0px;margin-bottom:0px">
 		<li>m�dico do Instituto Nacional de Seguro Social (INSS);</li>
	 	<li>m�dico da empresa ou por ela designado e pago;</li>
 		<li>m�dico do Servi�o Social da Ind�stria (SESI) ou do Servi�o Social do Com�rcio (SESC), conforme o caso;</li>
	 	<li>m�dico de reparti��o federal, estadual ou municipal, incumbida de assuntos de higiene ou sa�de;</li>
 		<li>m�dico do sindicato a que perten�a o EMPREGADO.</li>
	</ol>
 	A ordem preferencial estabelecida na Lei n� 605/49 para a justificativa de faltas ao trabalho d� ao EMPREGADOR o direito de aceitar ou n�o atestados fornecidos
	por m�dicos particulares.

<%if tipo="F" then%>
	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	8 - O EMPREGADO se obriga a cumprir fielmente o Contrato de Trabalho, bem como a preservar a integralidade, a confiabilidade e a confidencialidade sobre os documentos
	e as informa��es da EMPREGADORA a que tiver acesso, al�m de manter o mais completo sigilo sobre quaisquer dados materiais, pormenores, informa��es, documentos, especifica��es
	t�cnicas ou comerciais, inova��es ou aperfei�oamentos da EMPREGADORA de que venha a ter conhecimento, acesso ou lhe seja confiado em raz�o deste Contrato, n�o podendo, sob qualquer
	pretexto, divulgar, revelar, reproduzir, utilizar ou deles dar conhecimento a terceiros e estranhos a esta contrata��o, nem induzir alunos da EMPREGADORA a transferir-se para outra
	institui��o, sob as penas da lei.
<%end if%>	
<%if tipo="F" then c1="9" else c1="8"%>	
	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	<%=c1%> - O presente contrato, viger� durante <%=dias%> (<%=edias%> dias), sendo celebrado para as partes verificarem reciprocamente, a conveni�ncia ou n�o de se
	vincularem em car�ter definitivo a um Contrato de Trabalho. O EMPREGADOR passando a conhecer as aptid�es do EMPREGADO e suas qualidades pessoais e morais;
	o EMPREGADO verificando se o ambiente e os m�todos de trabalho atendem � sua conveni�ncia.
	</td>
</tr>
</table>
<DIV style="page-break-after:always"></DIV>	
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class="campop" style="font-size:12px;border:0 solid">
<%if tipo="F" then c2="10" else c2="9"%>	
	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	<%=c2%> - Opera-se a rescis�o do presente Contrato pela decorr�ncia do prazo supra ou por vontade de uma das partes; rescindindo-se por vontade do EMPREGADO ou pela
	EMPREGADORA com justa causa, nenhuma indeniza��o � devida; rescindindo-se, antes do prazo, pela EMPREGADORA, fica esta obrigada a pagar 50% dos sal�rios devidos 
	at� o final (metade do tempo combinado restante), nos termos do artigo 479 da CLT, sem preju�zo do disposto no Regulamento do FGTS. Nenhum aviso pr�vio � devido
	pela rescis�o do presente contrato.

	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
<%if tipo="F" then c3="11" else c3="10"%>	
	<%=c3%> - Na hip�tese deste ajuste transformar-se em Contrato de Prazo Indeterminado, pelo decurso do tempo, continuar�o em plena vig�ncia as cl�usulas 1 (um) a
	7 (sete), enquanto durarem as rela��es do EMPREGADO com a EMPREGADORA.

	<p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	E por estarem de pleno acordo, as partes contratantes, assinam o presente Contrato de Experi�ncia em duas vias, ficando a primeira em poder da EMPREGADORA, e a
	segunda via com o EMPREGADO, que dela dar� o competente recibo.
	</td>
</tr>

<tr><td height=5></td></tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="<%=largura1%>">
<tr>
	<td class="campop" style="font-size:12px;border:0px solid;"><p style="margin:2px;line-height:18px">
	Osasco,  <%=day(datarel)%> de <%=monthname(month(datarel))%> de <%=year(datarel)%>
	<br>
		<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="640">
		<tr>
			<td class=campo height=40 width="57%" style="border-bottom:1px solid"></td>
			<td class=campo height=40 width="3%" style="border-bottom:0px solid"></td>
			<td class=campo height=40 width="40%" style="border-bottom:1px solid"></td>
		</tr>
		<tr>
			<td class="campop" height=15 style="border-bottom:0px solid">FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</td>
			<td class="campop">&nbsp;</td>
			<td class="campop" height=15 style="border-bottom:0px solid">TESTEMUNHA</td>
		</tr>	
		<tr>
			<td class=campo height=40 style="border-bottom:1px solid"></td>
			<td class="campop">&nbsp;</td>
			<td class=campo height=40 style="border-bottom:1px solid"></td>
		</tr>
		<tr>
			<td class="campop" height=15 style="border-bottom:0px solid"><%=rs("nome")%></td>
			<td class="campop">&nbsp;</td>
			<td class="campop" height=15 style="border-bottom:0px solid">TESTEMUNHA</td>
		</tr>	
		</table>
	</td>
</tr>
</table>
<%
if tipo="A" or (tipo="F" and dias<90) then
data1=dateadd("d",44,datarel)
if weekday(data1)=7 then data1=dateadd("d",-1,data1)
if weekday(data1)=1 then data1=dateadd("d",-2,data1)
data2=dateadd("d",89,datarel)
%>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="<%=largura1%>">
<tr><td height=20></td></tr>
<tr>
	<td class=fundo height=25 align="center" style="font-size:14px;border:1px solid #000000"><b>TERMO DE PRORROGA��O</td>
</tr>
<tr><td height=5></td></tr>
<tr>
	<td class="campop" style="font-size:12px;border:0px solid;"><p style="margin:2px;line-height:18px;text-align:justify;margin-top:10px">
	Por m�tuo acordo entre as partes, fica o presente contrato de experi�ncia, que deveria vencer nesta data, prorrogado at� <%=data2%>.
	</td>
</tr>
<tr>
	<td class="campop" style="font-size:12px;border:0px solid;"><p style="margin:2px;line-height:18px">
	Osasco,  <%=day(data1)%> de <%=monthname(month(data1))%> de <%=year(data1)%>
	<br>
		<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="640">
		<tr>
			<td class=campo height=40 width="57%" style="border-bottom:1px solid"></td>
			<td class=campo height=40 width="3%" style="border-bottom:0px solid"></td>
			<td class=campo height=40 width="40%" style="border-bottom:1px solid"></td>
		</tr>
		<tr>
			<td class="campop" height=15 style="border-bottom:0px solid">FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</td>
			<td class="campop">&nbsp;</td>
			<td class="campop" height=15 style="border-bottom:0px solid">TESTEMUNHA</td>
		</tr>	
		<tr>
			<td class=campo height=40 style="border-bottom:1px solid"></td>
			<td class="campop">&nbsp;</td>
			<td class=campo height=40 style="border-bottom:1px solid"></td>
		</tr>
		<tr>
			<td class="campop" height=15 style="border-bottom:0px solid"><%=rs("nome")%></td>
			<td class="campop">&nbsp;</td>
			<td class="campop" height=15 style="border-bottom:0px solid">TESTEMUNHA</td>
		</tr>	
		</table>
	</td>
</tr>
</table>
<%
end if 'tipo a=prorroga��o
%>

</div>

	</td>
</tr>
</table>

<div align="center">
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td>
	<p style="margin-bottom:0px;margin-top:0px;font-size:9px;text-align:justify;">
</td></tr></table>
</div>

<%
'if request.form("via4")>1 and v<request.form("via4") then 
response.write "<DIV style=""page-break-after:always""></DIV>"
next 'via4
rs.close
end if 'id4

'***************************
'** INICIO DO FORMUL�RIO  **
'***************************
'if request.form("form_id").count>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
for a=1 to request.form("form_id").count
	if request.form("form_id").item(a)="id5" then formulario5="S"
next
'if formulario5="S" and (formulario4="S" or formulario3="S" or formulario2="S" or formulario1="S" or formulario0="S") then response.write "<DIV style=""page-break-after:always""></DIV>"
if formulario5="S" then 'acordo de compensa��o de horas
sqla="select chapa, nome, carteiratrab ctps, seriecarttrab serie, cpf, f.funcao, f.rua, f.numero, f.complemento, f.bairro, f.cidade, f.estado, f.cep, cnpj=s.cgc, " & _
"frua=s.rua, fnumero=s.numero, fbairro=s.bairro, fcidade=s.cidade, f.sexo, f.codsindicato, f.codhorario " & _
"from qry_funcionarios f inner join corporerm.dbo.psecao s on s.codigo=f.codsecao where chapa='" & chapa & "'"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
end if

if request.form("emissao")="H" and rs("codsindicato")<>"03" then sqls="select codhorario from corporerm.dbo.pfunc where chapa='" & chapa & "'"
if request.form("emissao")="A" and rs("codsindicato")<>"03" then sqls="select codhorario from corporerm.dbo.pfhsthor where chapa='" & chapa & "' and dtmudanca='" & dtaccess(datarel) & "'"
if request.form("emissao")="D" and rs("codsindicato")<>"03" then sqls="select codhorario from corporerm.dbo.pfhsthor h where chapa='" &  chapa & "' and dtmudanca=(select top 1 dtmudanca from corporerm.dbo.pfhsthor where chapa='" & chapa & "' and dtmudanca<='" & dtaccess(datarel) & "' order by dtmudanca desc)"

'response.write datarel & " - " & request.form("emissao") & "<br>" & sqls
rs2.Open sqls, ,adOpenStatic, adLockReadOnly
horario=rs2("codhorario")
rs2.close
'response.write " " & horario
for v=1 to request.form("via5")
%>
<center>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;border:1px dotted" width="690" height=990>
<tr>
	<td class=campo valign=top>
		<p align="center" style="margin-bottom:0px;margin-top:15px;font-size:16px">ACORDO PARA COMPENSA��O DE HORAS DE TRABALHO</p>
		<p align="center" style="margin-bottom:5px;margin-top:0px;font-size:12px">(INDIVIDUAL)</p>

<div align="center">
	
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class="campop" style="font-size:14px;border:0 solid"><p style="margin:5px;line-height:30px">
	Entre a empresa <u>FUNDA��O INSTITUTO DE ENSINO PARA OSASCO<%for t=1 to 10:response.write "&nbsp;":next%></u>, situada � <u><%=rs("frua")%>, <%=rs("fnumero")%> - 
	<%=rs("fbairro")%> em <%=rs("fcidade")%><%for t=1 to 10:response.write "&nbsp;":next%></u>, e o empregado abaixo assinado, portador da CTPS n� 
	<u><%=rs("ctps")%><%for t=1 to 10:response.write "&nbsp;":next%></u>, s�rie <u><%=rs("serie")%><%for t=1 to 10:response.write "&nbsp;":next%></u> fica
	convencionado, de acordo com o disposto no Art. 59 e seu � 2� do Decreto-Lei n� 5.452 de 01/05/1943 (Consolida��o das Leis do Trabalho), que o hor�rio do
	trabalho ser� o seguinte:
	</td>
</tr>
<%
sqlh="select codigo, descricao, databasehor, dias=(select max(indice) from corporerm.dbo.abathor where codhorario=a.codigo) " & _
"from corporerm.dbo.ahorario a where codigo='" & horario & "'"
rs2.Open sqlh, ,adOpenStatic, adLockReadOnly
descricao=rs2("descricao")
databasehor=rs2("databasehor")
dias=rs2("dias")
rs2.close

if dias>7 then vezes=2 else vezes=1
matriz=cint(dias-1)
redim qindice (matriz), qentrada(matriz), qsaida(matriz), qrefeicao(matriz), qtipo(matriz), textoh(matriz*2)

sqlq="select a.indice, entrada=min(batida), saida=max(batida), refeicao=min(i.intervalo), tipo='N' " & _
"from corporerm.dbo.abathor a " & _
"left join (select indice, inicio, fim, intervalo=fim-inicio from corporerm.dbo.abathor where codhorario='277' and tipo=4) i on i.indice=a.indice " & _
"where codhorario='" & horario & "' and tipo=0 group by a.indice " & _
"union " & _
"select a.indice, inicio, fim, 0, tipo='D' from corporerm.dbo.abathor a where codhorario='" & horario & "' and tipo in (1,2) "
'response.write sqlq
rs2.Open sqlq, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
	m2=rs2.absoluteposition-1
	qindice(m2)=rs2("indice")
		if qindice(m2)=7 or qindice(m2)=14 then temp=1 else temp=qindice(m2)+1
		if temp>7 then temp=temp-7
		qindice(m2)=weekdayname(temp)
	qentrada(m2)=rs2("entrada")
		qentrada(m2)=horaload(qentrada(m2),2)
	qsaida(m2)=rs2("saida")
		qsaida(m2)=horaload(qsaida(m2),2)
	qrefeicao(m2)=rs2("refeicao")
		qrefeicao(m2)=horaload(qrefeicao(m2),2)
	qtipo(m2)=rs2("tipo")
rs2.movenext
loop
rs2.close

'for a=0 to matriz
'	response.write "<br>" & a & " - " & qindice(a) & " - " & qentrada(a) & " - " & qsaida(a) & " - " & qrefeicao(a) & " - " & qtipo(a)
'next

initxt=0
textoh(initxt)=qindice(initxt) & " das " & qentrada(initxt) & " as " & qsaida(initxt) & iif(qrefeicao(b)<>""," com intervalo de " & qrefeicao(initxt) & ".","")
''response.write "<br>inicio: " & textoh(initxt)
ultimo=0
for a=0 to matriz
	'for b=a+1 to matriz
	b=a+1
if b<=matriz then
	if qtipo(a)="N" and b<=matriz then
	''response.write "<br>----------" & qtipo(a)
		if a>0 then ant=a-1 else ant=0
		if a=7 then 
			textoh(initxt+1)="E na semana alternada:"
			initxt=initxt+2:ultimo=a
		end if
		if qentrada(a)=qentrada(b) and qsaida(a)=qsaida(b) then
			textoh(initxt)=qindice(ultimo) & " a " & qindice(b) & " das " & qentrada(b) & " as " & qsaida(b) & iif(qrefeicao(b)<>""," com intervalo de " & qrefeicao(b) & ".","")
		elseif qentrada(a)=qentrada(ant) then
			'textoh(initxt)=qindice(ultimo) & " das " & qentrada(a) & " as " & qsaida(a) & " com intervalo de " & qrefeicao(a) & "."
			''response.write "<br>ultimo igual"
		elseif qentrada(a)<>qentrada(b) then
			initxt=initxt+1
			textoh(initxt)=qindice(a) & " das " & qentrada(a) & " as " & qsaida(a) & iif(qrefeicao(b)<>""," com intervalo de " & qrefeicao(a) & ".","")
			''response.write "<br>-->xxx a=" & a & " / b=" & b & " >> " & initxt
			ultimo=a+1
		else
			''response.write "<br>n�o � igual"
			initxt=initxt+1
			textoh(initxt)=qindice(a) & " das " & qentrada(b) & " as " & qsaida(b) & iif(qrefeicao(b)<>""," com intervalo de " & qrefeicao(b) & ".","")
			''response.write "<br>--> a=" & a & " / b=" & b & " >> " & initxt
			ultimo=a+1
		end if 'entrada igual
	end if 'nao � normal � descanso
end if
'	response.write "<br>=>" & a & ": " & initxt & "=>" & textoh(initxt)
	'next 'b
next 'a

sqlj="select jornada=sum(horasfalta) from corporerm.dbo.ajorhor where codhorario='" & horario & "'"
rs2.Open sqlj, ,adOpenStatic, adLockReadOnly
jornada=rs2("jornada")
if jornada>44 and matriz>6 then jornada=jornada/2
rs2.close
if jornada>0 then jornada=horaload(jornada,1)
%>
<tr>
	<td class=campo style="border:1px dotted">
	<input type="text" style="font-size:13px;height:25px;border:0px transparent" size="80" value="<%=textoh(0)%>">
	<input type="text" style="font-size:13px;height:25px;border:0px transparent" size="80" value="<%=textoh(1)%>">
	<input type="text" style="font-size:13px;height:25px;border:0px transparent" size="80" value="<%=textoh(2)%>">
	<input type="text" style="font-size:13px;height:25px;border:0px transparent" size="80" value="<%=textoh(3)%>">
	<input type="text" style="font-size:13px;height:25px;border:0px transparent" size="80" value="<%=textoh(4)%>">
	<input type="text" style="font-size:13px;height:25px;border:0px transparent" size="80" value="<%=textoh(5)%>">
	</td>
</tr>
<tr>
	<td class="campop" style="border:0 solid"><p style="margin:5px;line-height:30px">
	perfazendo o total de <input type="text" class=form_input10 style="border-bottom:1px dotted #000000;" value="<%=jornada%>" size=4> horas semanais.
	<br>
	E por estarem de pleno acordo, as partes contratantes assinam o presente em duas vias, o qual vigorar� 
	<input type="text" class=form_input10 style="border-bottom:1px dotted #000000;" value="por prazo indeterminado" size=30>.
	<br>
		<font style="font-size:12px"><br><%=espacao%>Osasco, 
		<input type="text" class=form_input10 value="<%=day(datarel)%>" size=3> de 
		<input type="text" class=form_input10 value="<%=monthname(month(datarel))%>" size=20> de 
		<input type="text" class=form_input10 value="<%=year(datarel)%>" size=5>
		<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="640">
		<tr>
			<td class=campo height=40 width="47%" style="border-bottom:1px solid"></td>
			<td class=campo height=40 width="3%" style="border-bottom:0px solid"></td>
			<td class=campo height=40 width="50%" style="border-bottom:1px solid"></td>
		</tr>
		<tr>
			<td class="campop" height=15 style="border-bottom:0px solid"><%=rs("nome")%></td>
			<td class="campop">&nbsp;</td>
			<td class="campop" height=15 style="border-bottom:0px solid">FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</td>
		</tr>	
		</table>
	</td>
</tr>
<tr><td height=5></td></tr>
</table>

</div>

	</td>
</tr>
</table>

<div align="center">
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td>
<p style="margin-bottom:0px;margin-top:0px;font-size:9px;text-align:justify;">
</td></tr></table>
</div>

<%
'if request.form("via5")>1 and v<request.form("via5") then 
response.write "<DIV style=""page-break-after:always""></DIV>"
next 'via5
rs.close
end if 'id5

'***************************
'** INICIO DO FORMUL�RIO  **
'***************************
'if request.form("form_id").count>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
for a=1 to request.form("form_id").count
	if request.form("form_id").item(a)="id6" then formulario6="S"
next
'if formulario6="S" and (formulario5="S" or formulario4="S" or formulario3="S" or formulario2="S" or formulario1="S" or formulario0="S") then response.write "<DIV style=""page-break-after:always""></DIV>"
if formulario6="S" then 'ficha de registro
sqla="select f.chapa, f.codpessoa, f.nome, sexo, admissao, demissao, instrucao, estcivil, dtnascimento, naturalidade, " & _
"estadonatal, carteiratrab, seriecarttrab, ufcarttrab, dtcarttrab, tituloeleitor, zonatiteleitor, secaotiteleitor, " & _
"cartidentidade, ufcartident, orgemissorident, dtemissaoident, cpf, certifreserv, categmilitar, mae, pai, " & _
"pispasep, dtcadastropis, nacionalidade, codrecebimento, regprofissional, jornadames, nrofichareg, " & _
"datachegada, cartmodelo19, nroreggeral, dtvencident, tipovisto, naturalizado, conjugebrasil, nrofilhosbrasil, nacionalidade2, " & _
"rua, numero, complemento, bairro, cidade, cep, admissao, codsindicato " & _
"from qry_funcionarios f where chapa='" & chapa & "'"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
end if

select case request.form("emissao")
	case "A"
		sqlsecao="select codsecao from corporerm.dbo.pfhstsec where chapa='" & chapa & "' and dtmudanca='" & dtaccess(datarel) & "'"
		sqlcargo="select codfuncao from corporerm.dbo.pfhstfco where chapa='" & chapa & "' and dtmudanca='" & dtaccess(datarel) & "'"
		sqlsalar="select salario, jornada from corporerm.dbo.pfhstsal where chapa='" & chapa & "' and dtmudanca between '" & dtaccess(datarel) & "' and dateadd(""hh"",23,'"& dtaccess(datarel) & "') "
		sqlhor  ="select codhorario from corporerm.dbo.pfhsthor where chapa='" & chapa & "' and dtmudanca='" & dtaccess(datarel) & "'"
	case "H"
		sqlsecao="select codsecao from corporerm.dbo.pfunc where chapa='" & chapa & "'"
		sqlcargo="select codfuncao from corporerm.dbo.pfunc where chapa='" & chapa & "'"
		sqlsalar="select salario, jornada=jornadamensal from corporerm.dbo.pfunc where chapa='" & chapa & "'"
		sqlhor  ="select codhorario from corporerm.dbo.pfunc where chapa='" & chapa & "'"
	case "D"
		sqlsecao="select codsecao from corporerm.dbo.pfhstsec where chapa='" & chapa & "' and dtmudanca=(select top 1 dtmudanca from corporerm.dbo.pfhstsec where chapa='" & chapa & "' and dtmudanca<='" & dtaccess(datarel) & "' order by dtmudanca desc) "
		sqlcargo="select codfuncao from corporerm.dbo.pfhstfco where chapa='" & chapa & "' and dtmudanca=(select top 1 dtmudanca from corporerm.dbo.pfhstfco where chapa='" & chapa & "' and dtmudanca<='" & dtaccess(datarel) & "' order by dtmudanca desc) "
		sqlsalar="select salario, jornada from corporerm.dbo.pfhstsal where chapa='" & chapa & "' and dtmudanca=(select top 1 dtmudanca from corporerm.dbo.pfhstsal where chapa='" & chapa & "' and dtmudanca<='" & dtaccess(datarel) & "' order by dtmudanca desc) "
		sqlhor  ="select codhorario from corporerm.dbo.pfhsthor where chapa='" & chapa & "' and dtmudanca=(select top 1 dtmudanca from corporerm.dbo.pfhsthor where chapa='" & chapa & "' and dtmudanca<='" & dtaccess(datarel) & "' order by dtmudanca desc) "
end select

sqlfuncao="select nome, cbo2002  from corporerm.dbo.pfuncao where codigo=(" & sqlcargo & ")"
rs2.Open sqlfuncao, ,adOpenStatic, adLockReadOnly
cargoimpressao=rs2("nome")
cboimpressao=rs2("cbo2002")
rs2.close
sqlhorario="select descricao from corporerm.dbo.ahorario where codigo=(" & sqlhor & ")"
rs2.Open sqlhorario, ,adOpenStatic, adLockReadOnly
horarioimpressao=rs2("descricao")
rs2.close
if rs("codsindicato")="03" then horarioimpressao="Hor�rio atribuido conforme grade curricular vigente."

sqlcnpj="select cgc, rua, numero, bairro, cidade, estado, cep, descricao from corporerm.dbo.psecao where codigo=(" & sqlsecao & ")"
rs2.Open sqlcnpj, ,adOpenStatic, adLockReadOnly
setorimpressao=rs2("descricao")
cnpj=rs2("cgc"):frua=rs2("rua"):fnumero=rs2("numero"):fbairro=rs2("bairro"):fcidade=rs2("cidade"):festado=rs2("estado"):fcep=rs2("cep")
rs2.close
for v=1 to request.form("via2")
%>
<center>
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;border:1px dotted" width="690" height=900>
<tr>
	<td class=campo valign=top>
<br>

<div align="center">

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1+10%>">
<tr><td class=campo height=30 align="center" style="font-size:14px;border:1px solid #000000"><b>FICHA DE REGISTRO DE EMPREGADO</td>
</tr>	
<tr><td height=10></td></tr>
</table>

<!-- --><%largfoto=115:ajuste=5:altura1=30%>
<table><tr><td width="<%=largfoto%>" height=150 valign="top" class=fundo style="border:1px solid #000000">
	<img border="0" src="..\func_foto.asp?chapa=<%=rs("chapa")%>" width="<%=largfoto%>" height=<%=largfoto/3*4%>>

<!-- -->

<!-- -->
</td><td class=campo valign=top>
<!-- -->

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1-largfoto-ajuste%>">
<tr><td class="campor" height=<%=altura1%> style="border-top:1px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>Empresa</b><br>
		<font style="font-size:12px">&nbsp;FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</td>
	<td class="campor" style="border-top:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>CNPJ</b><br>
		<font style="font-size:12px">&nbsp;<%=cnpj%></td>
	<td class="campor" width=100 style="border-top:1px solid;border-right:1px solid" align="center"><font style="font-size:9px">&nbsp;Matr�cula<br>
		<font style="font-size:12px">&nbsp;<b><%=chapa%></b></td>
</tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1-largfoto-ajuste%>">
<tr><td class="campor" height=<%=altura1%> style="border-top:0px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>Endere�o</b><br>
		<font style="font-size:12px">&nbsp;<%=frua%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Numero</b><br>
		<font style="font-size:12px">&nbsp;<%=fnumero%></td>
	<td class="campor" style="border-top:0px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>Bairro</b><br>
		<font style="font-size:12px">&nbsp;<%=fbairro%></td>
	<td class="campor" width=100 style="border-top:0px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;N� Ordem<br>
		<font style="font-size:12px">&nbsp;<%=rs("nrofichareg")%></td>
</tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1-largfoto-ajuste%>">
<tr><td class="campor" height=<%=altura1%> style="border-bottom:1px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>Cidade</b><br>
		<font style="font-size:12px">&nbsp;<%=fcidade%></td>
	<td class="campor" style="border-bottom:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Estado</b><br>
		<font style="font-size:12px">&nbsp;<%=festado%></td>
	<td class="campor" style="border-bottom:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>CEP</b><br>
		<font style="font-size:12px">&nbsp;<%=fcep%></td>
	<td class="campor" width=100 style="border-bottom:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;Emiss�o<br>
		<font style="font-size:12px">&nbsp;<%=datarel%></td>
</tr>
<tr><td height=10></td></tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="<%=largura1-largfoto-ajuste%>">
<tr><td class="campor" height=45 style="border-top:1px solid;border:1px solid"><font style="font-size:9px">&nbsp;<b>Nome do funcion�rio</b><br>
		<font style="font-size:15px">&nbsp;<b><%=rs("nome")%></b></td>
</tr>
</table>

<!-- -->
</td></tr>
</table>
<!-- -->

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" height=<%=altura1%> style="border-top:1px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>Data nascimento</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("dtnascimento")%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Local nascimento</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("naturalidade")%>&nbsp;<%=rs("estadonatal")%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Nacionalidade</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("nacionalidade2")%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Estado civil</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("estcivil")%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Sexo</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("sexo")%></td>
	<td class="campor" style="border-top:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>Instru��o</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("instrucao")%></td>
</tr>
</table>

<%
sqlconj="select nome from corporerm.dbo.pfdepend where chapa='" & chapa & "' and grauparentesco in ('5','C')" 
rs2.Open sqlconj, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then conjuge=rs2("nome") else conjuge=""
rs2.close
%>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" height=<%=altura1%> style="border-top:1px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>Filia��o</b><br>
		<font style="font-size:12px">&nbsp;M�e: <%=rs("mae")%><br>&nbsp;Pai: <%=rs("pai")%></td>
	<td class="campor" width=50% valign=top style="border-top:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b></b><br>
		<font style="font-size:12px">&nbsp;</td>
</tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" colspan=3 height=<%=altura1-18%> style="border-top:1px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>Carteira de Trabalho e Previd�ncia Social</b></td>
	<td class="campor" colspan=3 style="border-top:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>T�tulo de Eleitor</b></td>
	<td class="campor" colspan=3 style="border-top:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>Cart. de Identidade</b></td>
</tr>
<tr><td class="campor" height=<%=altura1%> style="border-top:0px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>N�</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("carteiratrab")%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>S�rie</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("seriecarttrab")%>&nbsp;<%=rs("ufcarttrab")%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Emiss�o</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("dtcarttrab")%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>N�</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("tituloeleitor")%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Zona</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("zonatiteleitor")%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Se��o</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("secaotiteleitor")%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>N�</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("cartidentidade")%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Emissor</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("orgemissorident")%>&nbsp;<%=rs("ufcartident")%></td>
	<td class="campor" style="border-top:0px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>Emiss�o</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("dtemissaoident")%></td>
</tr>
</table>	
	
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" height=<%=altura1%> style="border-top:1px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>C.P.F.</b><br>
		<font style="font-size:12px">&nbsp;<%=formatacpf(rs("cpf"))%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Reservista</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("certifreserv")%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Categoria</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("categmilitar")%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Reg. Profissional</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("regprofissional")%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>PIS/PASEP</b><br>
		<font style="font-size:12px">&nbsp;<%=formatapis(rs("pispasep"))%></td>
	<td class="campor" style="border-top:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>Data cadastro</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("dtcadastropis")%></td>
</tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" colspan=6 height=<%=altura1-18%> style="border-top:1px solid;border-left:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>Quando estrangeiro</b></td>
</tr>
<tr><td class="campor" height=<%=altura1%> style="border-top:0px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>Pa�s</b><br>
		<font style="font-size:12px">&nbsp;<%if rs("nacionalidade")<>"10" then response.write rs("nacionalidade2")%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Data chegada</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("datachegada")%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Mod. 19</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("cartmodelo19")%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>N� R.G.</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("nroreggeral")%></td>
	<td class="campor" style="border-top:0px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Validade</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("dtvencident")%></td>
	<td class="campor" style="border-top:0px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>Tipo Visto</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("tipovisto")%></td>
</tr>
</table>	

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" height=<%=altura1%> style="border-top:1px solid;border-left:1px solid;border-bottom:1px solid"><font style="font-size:9px">&nbsp;<b>Naturalizado?</b><br>
		<font style="font-size:12px">&nbsp;<%if rs("naturalizado")=1 then response.write "Sim"%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid;border-bottom:1px solid"><font style="font-size:9px">&nbsp;<b>Casado(a) com brasileira(o)</b><br>
		<font style="font-size:12px">&nbsp;<%if rs("conjugebrasil")=1 then response.write "Sim"%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid;border-bottom:1px solid"><font style="font-size:9px">&nbsp;<b>Tem filhos</b><br>
		<font style="font-size:12px">&nbsp;<%if rs("conjugebrasil")=1 then response.write rs("nrofilhosbrasil")%></td>
	<td class="campor" width=300 style="border-top:1px solid;border-right:1px solid;border-bottom:1px solid"><font style="font-size:9px">&nbsp;<b>Nome conjuge</b><br>
		<font style="font-size:12px">&nbsp;<%if rs("conjugebrasil")=1 and rs("conjugebrasil")<>"" and len(rs("conjugebrasil"))>1 then response.write rs("conjugebrasil")%></td>
</tr>
<tr><td height=10></td></tr>
</table>	

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" height=<%=altura1%> style="border-top:1px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>Endere�o</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("rua") & " " & rs("numero") & " " & rs("complemento")%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Bairro</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("bairro")%></td>
	<td class="campor" style="border-top:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>Cidade</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("cidade")%></td>
</tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" width=25 style="border:1px solid" rowspan=7 align="center"><font style="font-size:9px"><b>B<br>E<br>N<br>E<br>F<br>I<br>C<br>I<br>�<br>R<br>I<br>O<br>S</b></td>
	<td class="campor" height=<%=altura1-22%> style="border-top:1px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>Nome</b></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid"><font style="font-size:9px">&nbsp;<b>Parentesco</b></td>
	<td class="campor" style="border-top:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>Data Nasc.</b></td>
</tr>
<%
totaldep=0
sqld="select top 5 nrodepend, nome, dtnascimento, descricao, datediff(yy, dtnascimento, " & dtaccess(datarel) & ") " & _
"from corporerm.dbo.pfdepend d inner join corporerm.dbo.pcodparent c on c.codcliente=d.grauparentesco " & _
"where chapa='" & chapa & "' and dtnascimento<='" & dtaccess(datarel) & "' and (grauparentesco in ('5','C') " & _
"or (grauparentesco in ('1','3') and datediff(yy, dtnascimento, " & dtaccess(datarel) & ")<21)) " & _
"order by datediff(yy, dtnascimento, " & dtaccess(datarel) & ") "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
totaldep=rs2.recordcount
do while not rs2.eof
%>
<tr><td class="campor" height=<%=altura1-18%> style="border-bottom:1px dotted;border-left:1px solid"><font style="font-size:10px">&nbsp;<%=rs2("nome")%></td>
	<td class="campor" style="border-bottom:1px dotted;border-right:0px solid"><font style="font-size:10px">&nbsp;<%=rs2("descricao")%></td>
	<td class="campor" style="border-bottom:1px dotted;border-right:1px solid"><font style="font-size:10px">&nbsp;<%=rs2("dtnascimento")%></td>
</tr>
<%
rs2.movenext:loop
end if
rs2.close
for a=1 to (6-totaldep)
%>
<tr><td class="campor" height=<%=altura1-18%> style="border-bottom:1px dotted;border-left:1px solid"><font style="font-size:10px">&nbsp;</td>
	<td class="campor" style="border-bottom:1px dotted;border-right:0px solid"><font style="font-size:10px">&nbsp;</td>
	<td class="campor" style="border-bottom:1px dotted;border-right:1px solid"><font style="font-size:10px">&nbsp;</td>
</tr>
<%
next
%>
<tr><td height=10></td></tr>
</table>

<%
sqls="select sum(salario) tsalario, sum(jornada) tjornada from (" & sqlsalar & ") s "
rs2.Open sqls, ,adOpenStatic, adLockReadOnly
salario=cdbl(rs2("tsalario")): if salario="" or isnull(Salario) then salario=0
jornada=cint(rs2("tjornada")):jornada=jornada/60 
if salario>0 then hora= cdbl(salario)/ cdbl(jornada) else hora=0
rs2.close
if request.form("datainter")<>"" then admissao=formatdatetime(request.form("datainter"),2) else admissao=rs("admissao")
%>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" valign="top" height=<%=altura1+5%> style="border-top:2px solid;border-left:2px solid;border-bottom:0px solid"><font style="font-size:9px">&nbsp;<b>Data Admiss�o</b><br>
		<font style="font-size:12px">&nbsp;<%=admissao%></td>
	<td class="campor" valign="top" style="border-top:2px solid;border-right:0px solid;border-bottom:0px solid" nowrap><font style="font-size:9px">&nbsp;<b>Cargo</b><br>
		<font style="font-size:12px">&nbsp;<%=cargoimpressao%></td>
	<td class="campor" valign="top" style="border-top:2px solid;border-right:0px solid;border-bottom:0px solid"><font style="font-size:9px">&nbsp;<b>Se��o</b><br>
		<font style="font-size:12px">&nbsp;<%=setorimpressao%></td>
	<td class="campor" valign="top" style="border-top:2px solid;border-right:0px solid;border-bottom:0px solid"><font style="font-size:9px">&nbsp;<b>Sal�rio</b><br>
		<font style="font-size:12px">&nbsp;<input type="text" class="form_input10" size="24" value="<%=formatnumber(salario,2) & "/" & rs("codrecebimento")%> | <%=formatnumber(hora,2)%> (hora)"></td>
	<td class="campor" valign="top" style="border-top:2px solid;border-right:2px solid;border-bottom:0px solid"><font style="font-size:9px">&nbsp;<b>Jornada</b><br>
		<font style="font-size:12px">&nbsp;<%=jornada%> horas</td>
</tr>
</table>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" valign="top" width=100 height=<%=altura1+5%> style="border-top:1px solid;border-left:2px solid;border-bottom:2px solid"><font style="font-size:9px">&nbsp;<b>C�digo CBO</b><br>
		<font style="font-size:12px">&nbsp;<%=cboimpressao%></td>
	<td class="campor" valign="top" style="border-top:1px solid;border-right:2px solid;border-bottom:2px solid"><font style="font-size:9px">&nbsp;<b>Hor�rio de trabalho</b><br>
		<font style="font-size:12px">
		<input type="text" size="80" value="<%=horarioimpressao%>" style="font-family:Tahoma;font-size:10pt;color:black;border:0px transparent;background-color:white;">
		</td>
</tr>
<tr><td height=10></td></tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" rowspan=3 colspan=1 width=100 height=<%=altura1%> style="border:1px solid;" valign=top align="center"><font style="font-size:9px">&nbsp;<b>Polegar direito</b><br></td>
	<td class="campor" rowspan=1 colspan=5 valign=top><font style="font-size:10px">&nbsp;<b>Estou de pleno acordo com as declara��es acima que expressam a verdade</b><br></td>
	<td class="campor" rowspan=3 colspan=1 width=100 height=<%=altura1%> style="border:1px solid;" valign=top align="center"><font style="font-size:9px">&nbsp;<b>Data da Sa�da</b><br></td>
	</tr>
<tr><td class="campor" rowspan=1 colspan=1 width=10 height=80>&nbsp;</td>
	<td class="campor" rowspan=1 colspan=1 width=200 style="border-bottom:1px solid">&nbsp;</td>
	<td class="campor" rowspan=1 colspan=1 width=10>&nbsp;</td>
	<td class="campor" rowspan=1 colspan=1 width=200 style="border-bottom:1px solid">&nbsp;</td>
	<td class="campor" rowspan=1 colspan=1 width=10>&nbsp;</td>
	</tr>
<tr><td class="campor" rowspan=1 colspan=1 width=10 height=20>&nbsp;</td>
	<td class="campor" rowspan=1 colspan=1 width=200 style="border-bottom:0px solid">&nbsp;ASSINATURA DO EMPREGADO</td>
	<td class="campor" rowspan=1 colspan=1 width=10>&nbsp;</td>
	<td class="campor" rowspan=1 colspan=1 width=200 style="border-bottom:0px solid">&nbsp;ASSINATURA DO EMPREGADOR</td>
	<td class="campor" rowspan=1 colspan=1 width=10>&nbsp;</td>
	</tr>
</table>

<!-- inicio texto -->
<!-- fim texto -->
</div>

	</td>
</tr>
</table>

<%
verso=1
if verso=1 then
%>
<DIV style="page-break-after:always"></DIV>
<div align="center">
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class=campo height=30 align="center" style="font-size:14px;border:1px solid #000000"><b>FICHA DE ANOTA��ES / HISTORICO DE ALTERA��ES</td></tr>	
<tr><td height=10></td></tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" height=<%=altura1%> style="border-top:1px solid;border-left:1px solid"><font style="font-size:9px">&nbsp;<b>Empresa</b><br>
		<font style="font-size:12px">&nbsp;FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</td>
	<td class="campor" height=<%=altura1%> style="border-top:1px solid;border-left:0px solid"><font style="font-size:9px">&nbsp;<b>Endere�o</b><br>
		<font style="font-size:12px">&nbsp;<%=frua & " " & fnumero%></td>
	<td class="campor" style="border-top:1px solid;border-right:1px solid"><font style="font-size:9px">&nbsp;<b>Cidade</b><br>
		<font style="font-size:12px">&nbsp;<%=fcidade%></td>
</tr>
</table>
<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr><td class="campor" height=<%=altura1%> style="border-top:1px solid;border-left:1px solid;border-bottom:1px solid"><font style="font-size:9px">&nbsp;<b>Nome do funcion�rio</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("nome")%></td>
	<td class="campor" style="border-top:1px solid;border-right:0px solid;border-bottom:1px solid"><font style="font-size:9px">&nbsp;<b>Admiss�o</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("admissao")%></td>
	<td class="campor" style="border-top:1px solid;border-right:1px solid;border-bottom:1px solid"><font style="font-size:9px">&nbsp;<b>CTPS n�</b><br>
		<font style="font-size:12px">&nbsp;<%=rs("carteiratrab") & "/" & rs("seriecarttrab")%></td>
</tr>
<tr><td height=10></td></tr>
</table>



<!-- -->
<table width="<%=largura1%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse"><tr><td width=<%=largura1/2-2%> valign=top>
<!-- -->

<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=100%>
<tr><td class="campop" colspan=4 align="center">ALTERA��ES DE SAL�RIO</td></tr>
<tr><td class=titulo align="center">Data</td>
	<td class=titulo align="center">Valor</td>
	<td class=titulo align="center">Jornada</td>
	<td class=titulo align="center">Motivo</td>
</tr>
<%
sqlh1="select * from (select top 60 chapa, dtmudanca, motivo codmotivo, descricao motivo, nrosalario, salario, jornada, codevento " & _
"from corporerm.dbo.pfhstsal h inner join corporerm.dbo.pmotmudsal m on m.codcliente=h.motivo " & _
"where chapa='" & chapa & "' and h.motivo<>'11' order by dtmudanca desc) z order by dtmudanca"
rs2.Open sqlh1, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
jornada=rs2("jornada"):jornada=cdbl(jornada/60)
if rs2("codevento")="255" or rs2("codevento")="256" or rs2("codevento")="257" or rs2("codevento")="258" or rs2("codevento")="128" or rs2("codevento")="138" then RT=1 else RT=0
if rs2("codevento")<>"" and RT=0 then
	'response.write "<br>" & cdbl(rs2("salario")) & "-" & jornada
	if jornada=0 then salario=0 else salario=cdbl(rs2("salario"))/jornada
	sqleve="select coddoc from g2cursoeve where sal='" & rs2("codevento") & "'"
	rs3.Open sqleve, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then curso=rs3("coddoc") else curso=""
	rs3.close
	compl="/aula " & curso
else
	salario=rs2("salario")
	compl=""
end if
%>
<tr><td class="campor" align="center"><%=formatdatetime(rs2("dtmudanca"),2)%></td>
	<td class="campor" align="right"><%=formatnumber(salario,2) & compl%></td>
	<td class="campor" align="center"><%=jornada%></td>
	<td class="campor" align="left"><%=rs2("motivo")%></td>
</tr>
<%
rs2.movenext:loop
rs2.close
%>
</table>
<br>
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=100%>
<tr><td class="campop" colspan=5 align="center">AFASTAMENTOS/LICEN�AS</td></tr>
<tr><td class=titulo align="center">In�cio</td>
	<td class=titulo align="center">Termino</td>
	<td class=titulo align="center">Dias</td>
	<td class=titulo align="center">Motivo</td>
	<td class=titulo align="center">Situa��o</td>
</tr>
<%
sqlh4="select chapa, dtinicio, dtfinal, dias=datediff(d,dtinicio, dtfinal)+1, h.tipo, m.descricao tipo, h.motivo, s.descricao situacao, h.observacao " & _
"from corporerm.dbo.pfhstaft h inner join corporerm.dbo.pcodsituacao m on m.codcliente=h.tipo " & _
"inner join corporerm.dbo.pmudsituacao s on s.codcliente=h.motivo " & _
"where chapa='" & chapa & "' order by dtinicio "
rs2.Open sqlh4, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>
<tr><td class="campor" align="center"><%=formatdatetime(rs2("dtinicio"),2)%></td>
	<td class="campor" align="center"><%=rs2("dtfinal")%></td>
	<td class="campor" align="center"><%=rs2("dias")%></td>
	<td class="campor" align="left"><%=rs2("tipo")%></td>
	<td class="campor" align="left"><%=rs2("situacao")%></td>
</tr>
<tr><td class="campor" colspan=5 style="border-bottom:2px solid" align="left"><%=rs2("observacao")%></td>
</tr>
<%
rs2.movenext:loop
rs2.close
%>
</table>
<br>

<!-- -->
</td><td width=5></td><td width=<%=largura1/2-3%> valign=top>
<!-- -->

<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=100%>
<tr><td class="campop" colspan=3 align="center">ALTERA��ES DE CARGO</td></tr>
<tr><td class=titulo align="center">Data</td>
	<td class=titulo align="center">Cargo</td>
	<td class=titulo align="center">Motivo</td>
</tr>
<%
sqlh2="select chapa, dtmudanca, motivo, m.descricao motivo, codfuncao, nome funcao " & _
"from corporerm.dbo.pfhstfco h inner join corporerm.dbo.pmotmudfuncao m on m.codcliente=h.motivo " & _
"inner join corporerm.dbo.pfuncao f on f.codigo=h.codfuncao " & _
"where chapa='" & chapa & "' order by dtmudanca "
rs2.Open sqlh2, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>
<tr><td class="campor" align="center"><%=formatdatetime(rs2("dtmudanca"),2)%></td>
	<td class="campor" align="left"><%=rs2("funcao")%></td>
	<td class="campor" align="left"><%=rs2("motivo")%></td>
</tr>
<%
rs2.movenext:loop
rs2.close
%>
</table>
<br>
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=100%>
<tr><td class="campop" colspan=3 align="center">ALTERA��ES DE SE��O</td></tr>
<tr><td class=titulo align="center">Data</td>
	<td class=titulo align="center">Se��o</td>
	<td class=titulo align="center">Motivo</td>
</tr>
<%
sqlh3="select chapa, dtmudanca, motivo, m.descricao motivo, codsecao, f.descricao secao " & _
"from corporerm.dbo.pfhstsec h inner join corporerm.dbo.pmotmudsecao m on m.codcliente=h.motivo " & _
"inner join corporerm.dbo.psecao f on f.codigo=h.codsecao " & _
"where chapa='" & chapa & "' order by dtmudanca "
rs2.Open sqlh3, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>
<tr><td class="campor" align="center"><%=formatdatetime(rs2("dtmudanca"),2)%></td>
	<td class="campor" align="left"><%=rs2("secao")%></td>
	<td class="campor" align="left"><%=rs2("motivo")%></td>
</tr>
<%
rs2.movenext:loop
rs2.close
%>
</table>
<br>
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=100%>
<tr><td class="campop" colspan=3 align="center">CONTRIBUI��O SINDICAL</td></tr>
<tr><td class=titulo align="center">Ano</td>
	<td class=titulo align="center">Sindicato</td>
	<td class=titulo align="center">Valor</td>
</tr>
<%
sqlh5="select chapa, dtcontribuicao, ano=year(dtcontribuicao), codsindicato, nome sindicato, valor " & _
"from corporerm.dbo.pfhstcsd h inner join corporerm.dbo.psindic s on s.codigo=h.codsindicato " & _
"where chapa='" & chapa & "' order by dtcontribuicao "
rs2.Open sqlh5, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>
<tr><td class="campor" align="center"><%=rs2("ano")%></td>
	<td class="campor" align="left"><%=rs2("sindicato")%></td>
	<td class="campor" align="right"><%=formatnumber(rs2("valor"),2)%></td>
</tr>
<%
rs2.movenext:loop
rs2.close
%>
</table>
<br>
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=100%>
<tr><td class="campop" colspan=6 align="center">F�RIAS</td></tr>
<tr><td class=titulo align="center">Vencidas em</td>
	<td class=titulo align="center">Per�odo</td>
	<td class=titulo align="center">Dias</td>
	<td class=titulo align="center">Abono</td>
	<td class=titulo align="center">Faltas</td>
</tr>
<%
sqlh6="select chapa, inipa=dtiniperaquis, fimpa=dtfimperaquis, nroperiodo, inifer=dtinigozo, fimfer=dtfimgozo, dias=datediff(d,dtinigozo,dtfimgozo)+1, nrodiascorridos, diasabono, nrofaltas " & _
", abono=case when diasabono>=1 then 1 else 0 end " & _
"from corporerm.dbo.pfhstfer_old " & _
"where chapa='" & chapa & "' and datediff(d,dtinigozo,dtfimgozo)+1>0 order by dtfimperaquis, nroperiodo "
sqlh6="select p.CHAPA, inipa=f.INICIOPERAQUIS, fimpa=p.FIMPERAQUIS, p.DATAPAGTO, inifer=p.DATAINICIO, fimfer=p.DATAFIM, dias=p.NRODIASFERIAS, " & _
"nrodiascorridos=datediff(d,DATAINICIO,DATAFIM)+1, diasabono=p.NRODIASABONO, nrofaltas=p.FALTAS, abono=case when NRODIASABONO>0 then 1 else 0 end " & _
"from corporerm.dbo.PFUFERIASPER p inner join corporerm.dbo.PFUFERIAS f on f.CHAPA=p.CHAPA and f.FIMPERAQUIS=p.FIMPERAQUIS " & _
"where p.CHAPA='" & chapa & "' and SITUACAOFERIAS<>'M' order by p.FIMPERAQUIS, DATAPAGTO "
rs2.Open sqlh6, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
if rs2("abono")=1 then abono="X" else abono="-"
%>
<tr><td class="campor" align="center"><%=formatdatetime(rs2("fimpa"),2)%></td>
	<td class="campor" align="left"><%=rs2("inifer") & " a " & rs2("fimfer")%></td>
	<td class="campor" align="center"><%=rs2("dias")%></td>
	<td class="campor" align="center"><%=abono%></td>
	<td class="campor" align="center"><%=rs2("nrofaltas")%></td>
</tr>
<%
rs2.movenext:loop
rs2.close
%>
</table>
<br>

<!-- -->
</td></tr></table>
<!-- -->

<table border="0" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class="campor" height=80 valign=top>Documento de anota��es aprovado pela Portaria n� 41 de 28/03/2007 do MTE.<br>
	Osasco, <%=now()%>
	</td>
	<td class="campor" width=150 valign=top>
	<img src="../images/assinaturarmsa.gif" width="150" border="0" alt="">
	</td>
</tr>
</table>


<p style="margin-bottom:0px;margin-top:0px;font-size:9px;text-align:justify;">
</div>
<%
end if 'verso=1
%>

<%
'if request.form("via6")>1 and v<request.form("via6") then 
'response.write "<DIV style=""page-break-after:always""></DIV>" 'nao precisa quebrar pagina, � a ultima
next 'via6
rs.close
end if 'id6


'***************************
'** INICIO DO FORMUL�RIO  **
'***************************
'if request.form("form_id").count>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
for a=1 to request.form("form_id").count
	if request.form("form_id").item(a)="id7" then formulario7="S"
next
'if formulario7="S" and (formulario0="S") then response.write "<DIV style=""page-break-after:always""></DIV>"
if formulario7="S" then 'termo responsabilidade

sqla="select chapa, nome, cartidentidade from qry_funcionarios where chapa='" & chapa & "'"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
end if

for v=1 to request.form("via7")
%>

<center>
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;border:1 dotted" width="690" height=900>
<tr>
	<td class=campo valign=top>
<br>

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class=campo align="left"><img src="..\images\logo_centro_universitario_unifieo_big.jpg" width="150"> </td>
	<td class="campop" align="center"><b>TERMO DE RESPONSABILIDADE, CONFIABILIDADE E CONFIDENCIALIDADE</b></td>
	<td class=campo align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>	
</table>
<br>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" "<%=largura1%>">
<tr>
	<td class="campop" style="text-align:justify">Eu, <b><%=rs("nome")%></b>, portador do R.G. n� <%=rs("cartidentidade")%>, declaro haver 
	solicitado acesso � rede interna, sistemas, internet e e-mail, ficando plenamente esclarecido e ciente a respeito
	da pol�tica interna da institui��o, comprometendo-me a:
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	1. Acessar sistemas, rede de Internet/Intranet e a caixa postal (e-mail) somente por necessidade de servi�o
	ou por determina��o expressa de superior hier�rquico, realizando as tarefas e opera��es em estrita observ�ncia
	aos procedimentos, normas e disposi��es que regem os acessos � Internet/Intranet e respectiva utiliza��o da
	caixa postal e os e-mails;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	2. N�o revelar, fora do �mbito profissional, fato ou informa��o de qualquer natureza de que tenha conhecimento
	por for�a de minhas atribui��es, salvo em decorr�ncia de decis�o competente na esfera legal ou judicial, bem
	como de autoridade superior;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	3. Manter a necess�ria cautela quando da exibi��o de dados em tela, impressora ou na grava��o em meios
	eletr�nicos, a fim de evitar que deles venham a tomar ci�ncia pessoas n�o autorizadas;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	4. Para garantir a impossibilidade de acesso indevido por terceiros, n�o deverei me ausentar do terminal
	sem encerrar ou bloquear a sess�o do sistema;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	5. N�o revelar as minhas senhas de acesso (<i>login</i>) � rede e, sobretudo, de acesso aos sistemas, seja para qual
	pessoa for, devendo seguir as recomenda��es de seguran�a em rela��o � cria��o de uma senha forte, conforme
	pol�tica vigente, de forma a possibilitar que ela continue secreta;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	6. Responder, em todas as inst�ncias, pelas consequ�ncias das a��es ou omiss�es de minha parte que possam p�r
	em risco ou comprometer a exclusividade de conhecimento de minha senha ou das transa��es a que tenha acesso;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	7. Cuidar da integridade, confidencialidade e disponibilidade dos dados, informa��es e sistemas aos quais
	tenho acesso, devendo comunicar, por escrito, � chefia imediata, quaisquer ind�cios ou possibilidades de
	irregularidades, desvios ou falhas identificadas nos sistemas, sendo proibida a explora��o de falhas ou
	vulnerabilidades porventura existentes nos sistemas;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	8. N�o proceder � navega��o em <i>sites</i> fora do �mbito profissional, <i>sites</i> pornogr�ficos, 
	defensores do uso de drogas, de pedofilia ou de cunho racistas e similares. Tenho ci�ncia de que todos os
	acessos s�o monitorados, registrados e divulgados ao meu gestor;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	9. N�o efetuar <i>download</i> e <i>upload</i> de arquivos eletr�nicos fora do contexto profissional,
	sendo que � minha ci�ncia que todos e quaisquer tipo de arquivos baixados ou enviados � rede p�blica dever�o
	ser autorizados pela Divis�o de Tecnologia da Informa��o-DTI;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	10. Responder pelo uso de programas de mensagens instant�neas, sabendo que est�o liberados somente para
	fins profissionais. O acesso a este recurso somente ser� liberado ap�s autoriza��o do superior;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	11. Ter ci�ncia de que o acesso � informa��o de minha caixa postal (<i>e-mail</i>) n�o me garante direito
	sobre ela, nem me confere autoridade para liberar acesso a outras pessoas, pois constitui informa��es 
	pertinentes � FIEO e ao UNIFIEO, uma vez que devo fazer uso para melhor desempenhar minhas atividades;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	12. Em conjunto ao item 7 acima, fica expresso que constitui o descumprimento de normas regulamentares e 
	quebra de sigilo funcional divulgar dados obtidos por meio do uso de minha caixa postal (<i>e-mail</i>), a
	qual tenho acesso, seja para outros funcion�rios ou para terceiros n�o envolvidos nos trabalhos executados;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	13. Devo alterar minha senha, sempre que obrigat�rio ou que tenha suspei��o de descoberta por terceiros, n�o
	usando combina��es simples que possam ser facilmente descobertas;
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	14. O acesso � informa��o n�o me garante direito sobre ela, nem me confere autoridade para liberar acesso a
	outras pessoas.
	</td>
</tr>
<tr>
	<td class="campop" style="text-align:justify">
	Declaro, nesta data, ter ci�ncia e estar de acordo com os procedimentos acima descritos, comprometendo-me a
	respeit�-los e cumpr�-los plena e integralmente.
	</td>
</tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=640>
<tr>
	<td class="campop" colspan=2>Osasco, ________de _______________________de ________.
	</td>
</tr>
<tr>
	<td class="campop"><br>_____________________________________________<br><%=rs("nome")%></td>
	<td class="campop"><br>_____________________________________________<br>FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</td>
</tr>
</table>

</div>
	</td>
</tr>
</table>

<%
'if request.form("via7")>1 and v<request.form("via7") then 
response.write "<DIV style=""page-break-after:always""></DIV>"
next 'via7
rs.close
end if 'id7


'***************************
'** INICIO DO FORMUL�RIO  **
'***************************
'if request.form("form_id").count>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
for a=1 to request.form("form_id").count
	if request.form("form_id").item(a)="id8" then formulario8="S"
next
'if formulario8="S" and (formulario0="S") then response.write "<DIV style=""page-break-after:always""></DIV>"
if formulario8="S" then 'op��o assistencia m�dica
sqla="select chapa, nome from qry_funcionarios where chapa='" & chapa & "'"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
end if
for v=1 to request.form("via8")
%>
<center>
<!-- -->
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;border:1 dotted" width="690" height=900>
<tr><td class=campo valign=top>
<br>
<!-- -->
<!-- inicio formulario -->
<%
session("assmed_adm")=chapa
%>
<!-- #Include file="../assmedica/opcao_inc.asp"-->
<!-- final formulario -->
<!-- -->
</td></tr></table>
<!-- -->
<%
'if request.form("via8")>1 and v<request.form("via8") then 
response.write "<DIV style=""page-break-after:always""></DIV>"
next 'via8
rs.close
end if 'id8


'***************************
'** INICIO DO FORMUL�RIO  **
'***************************
'if request.form("form_id").count>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
for a=1 to request.form("form_id").count
	if request.form("form_id").item(a)="id9" then formulario9="S"
next
'if formulario9="S" and (formulario0="S") then response.write "<DIV style=""page-break-after:always""></DIV>"
if formulario9="S" then 'op��o assistencia m�dica
sqla="select chapa, nome, secao from qry_funcionarios where chapa='" & chapa & "'"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
end if
for v=1 to request.form("via9")
%>
<center>
<!-- -->
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;border:1 dotted" width="690" height=900>
<tr><td class=campo valign=top>
<br>
<!-- -->
<center>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
<tr>
	<td class="campo" width="150"><img src="../images/logo_centro_universitario_unifieo_big.jpg" width="150" border="0">
	</td>
	<td class="campop" width=<%=largura1-150%> align="center"><B>FORMUL�RIO DE OP��O</B><br>
	<span style="font-size:8pt">Cesta B�sica / Vale Alimenta��o</span>
	</td>
</tr>
<tr>
	<td height="35" class="campop" colspan="2"><span style="font-size:7pt"><b>Nome</b></span><br>
	<%=rs("nome")%></td>
</tr>
<tr>
	<td height="35" class="campop" colspan="2"><span style="font-size:7pt"><b>Setor</b></span><br>
	<%=rs("secao")%></td>
</tr>
<tr>
	<td class="campo" colspan="2"><br>
	<span style="font-size:10pt"><u>Fa�o a op��o</u>, conforme cl�usula "Cesta B�sica" da Conven��o Coletiva
	dos Auxiliares de Administra��o Escolar de Osasco, pelo seguinte:
	<br><br>
	[&nbsp;&nbsp;&nbsp;&nbsp;] Receber vale-alimenta��o atrav�s de valor creditado em cart�o magn�tico. (1)
	<br><br>
	[&nbsp;&nbsp;&nbsp;&nbsp;] Receber cesta b�sica com produtos, que devo retirar todo final de m�s. (2)
	<br><br>
	<u>Declaro estar ciente</u> do seguinte:<br>
	(1) o valor � creditado no �ltimo dia do m�s.<br>
	(2) a cesta deve ser retirada em at� 3 dias ap�s o dia 30 de cada m�s. Passado este prazo, � considerado
	que n�o desejo a cesta naquela m�s, sem que isto signifique a troca pelo vale-alimenta��o.<br>
	* Caso deseje, s� poderei <u>alterar</u> minha op��o at� o dia 5 de cada m�s.<br>
	<br>
	</td>
</tr>
<tr>
	<td class="campop" valign="top" height="70">
	<span style="font-size:7pt"><b>Data:</b></span><br>
	
	</td>
	<td class="campop" valign="top">
	<span style="font-size:7pt"><b>Assinatura:</b></span><br>

	</td>
</tr>
</table>


<!-- -->
</td></tr></table>
<!-- -->
<%
'if request.form("via9")>1 and v<request.form("via9") then 
response.write "<DIV style=""page-break-after:always""></DIV>"
next 'via9
rs.close
end if 'id9


'***************************
'** INICIO DO FORMUL�RIO  **
'***************************
'if request.form("form_id").count>1 then response.write "<DIV style=""page-break-after:always""></DIV>"
for a=1 to request.form("form_id").count
	if request.form("form_id").item(a)="id10" then formulario10="S"
next
'if formulario10="S" and (formulario0="S") then response.write "<DIV style=""page-break-after:always""></DIV>"
if formulario10="S" then 'op��o assistencia m�dica

sqla="select chapa, nome, cartidentidade from qry_funcionarios where chapa='" & chapa & "'"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
end if

for v=1 to request.form("via10")
%>

<center>
<!-- -->
<table border="1" bordercolor=#000000 cellpadding="0" cellspacing="0" style="border-collapse: collapse;border:1 dotted" width="690" height=900>
<tr><td class=campo valign=top>
<br>
<!-- -->

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=largura1%>">
</table>


<!-- -->
</td></tr></table>
<!-- -->

<%
'if request.form("via10")>1 and v<request.form("via10") then 
response.write "<DIV style=""page-break-after:always""></DIV>"
next 'via10
rs.close
end if 'id10


end if ' request.form
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>