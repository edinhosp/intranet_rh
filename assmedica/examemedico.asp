<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a85")="N" or session("a85")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Exame M�dico - Angular</title>
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
function nome1() { form.chapa.value=form.nome.value; }
function chapa1() { form.nome.value=form.chapa.value; }
--></script>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
rs.CursorLocation=3
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
dim endereco(10), tipo(10), extra(10), ende(10), hor(10), obs(10)
tipo(0)="P/A/D/MF/RT" : endereco(0)="Osasco - Rua Itabuna" : extra(0)="" : 
tipo(1)="P/A/D/MF/RT" : endereco(1)="S�o Paulo - Brooklin" : extra(1)="" : 
tipo(2)="P/A/D/MF/RT" : endereco(2)="S�o Paulo - Centro" : extra(2)="" : 
tipo(3)="P/A/D/MF/RT" : endereco(3)="Cotia" : extra(3)="" : 
tipo(4)="P/A/D/MF/RT" : endereco(4)="Jundiai" : extra(4)="" : 
tipo(5)="P/A/D/MF/RT" : endereco(5)="S�o Paulo - Santo Amaro" : extra(5)="" : 
tipo(6)="P/A/D/MF/RT" : endereco(6)="S�o Bernardo - Centro" : extra(6)="" : 
tipo(7)="P/A/D/MF/RT" : endereco(7)="S�o Paulo - Vila Mariana" : extra(7)="" : 
tipo(8)="P/A/D/MF/RT" : endereco(8)="S�o Paulo - Morumbi" : extra(8)="" : 
tipo(0)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o / Retorno"
tipo(1)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o / Retorno"
tipo(2)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o / Retorno"
tipo(3)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o / Retorno"
tipo(4)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o / Retorno"
tipo(5)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o / Retorno"
tipo(6)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o / Retorno"
tipo(7)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o / Retorno"
tipo(8)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o / Retorno"

if request.form("B1")="" or request.form("local")="" then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Sele��o de funcion�rio para exame m�dico - Angular
<form method="POST" action="examemedico.asp" name="form">
<%
sqla="select chapa, nome from corporerm.dbo.PFUNC f " & _
"where (codsituacao='D' and DATADEMISSAO>GETDATE()-45 and CHAPA<'10000') or CODSITUACAO<>'D' order by nome"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if request("chapa")<>"" then chapa=request("chapa") else chapa=request.form("chapa")
%>
<input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=chapa%>">
<select name="nome" class=a onchange="nome1()">
	<option value="">Selecione o funcion�rio</option>
	<option value="0">-=: FORMUL�RIO EM BRANCO :=-</option>
<%
rs.movefirst:do while not rs.eof
if request.form("chapa")=rs("chapa") or request("chapa")=rs("chapa") then temps="selected" else temps=""
%>
	<option value="<%=rs("chapa")%>" <%=temps%> ><%=rs("nome")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<br><br>
<table border="1" cellpadding="2" cellspacing="2" style="border-collapse: collapse">
<tr>
	<td class=titulo></td>
	<td class=titulo>Local de Exame</td>
	<td class=titulo>Tipo</td>
	<td class=titulo>Exames</td>
</tr>
<%
for a=0 to 8 'max 10
%>
<tr>
	<td class=campo><input type="radio" name="local" value="<%=a%>"></td>
	<td class=campo><%=endereco(a)%></td>
	<td class=campo><%=tipo(a)%></td>
	<td class=campo><%=extra(a)%></td>
</tr>
<%
next 'a
%>
</table>
<br>
<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>
<%
end if
'response.write request.form

if request.form("B1")<>"" and request.form("local")<>"" then
temp=request.form("local")
chapa=request.form("chapa")
if chapa<>"0" then
	sqla="select f.chapa, f.nome, f.DTNASCIMENTO, f.Secao, f.Funcao, f.CARTIDENTIDADE, s.cgc as cnpj " & _
	"from qry_funcionarios f inner join corporerm.dbo.PSECAO s on s.CODIGO=f.codsecao where chapa='" & chapa & "'"
	rs.Open sqla, ,adOpenStatic, adLockReadOnly
	nome=rs("nome")
	nascimento=rs("dtnascimento")
	idade=int((now()-rs("dtnascimento"))/365.25)
	setor=rs("secao")
	funcao=rs("funcao")
	rg=rs("cartidentidade")
	cnpj=rs("cnpj")
	rs.close
else
	nome="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	nascimento=space(20) & "&nbsp;"
	idade="&nbsp;"
	setor="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	funcao="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	rg="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	cnpj="&nbsp;73.063.166/_________-____"
end if
teste=0
if teste=1 then
	nome="MARCELO SANSINI TERRA"
	nascimento=space(20) & "&nbsp;"
	idade="&nbsp;"
	setor="CURSO TEC.MARKETING"
	funcao="PROFESSOR"
	rg="&nbsp;"
	cnpj="&nbsp;73.063.166/0003-92"
end if
ende(0)="Rua Itabuna, 93 - Centro de Osasco<br>Osasco - CEP 06010-120"
hor(0)="Seg. a Qui. 06h30 as 16h30 / Sex. 06h30 as 15h30"
ende(1)="Rua Joaquim Guarani, 105 - Brooklin"
hor(1)="Seg. a Qui. 07h30 as 11h30 e das 13h00 as 16h30<br>Sex. 07h00 as 11h00"
ende(2)="Rua Conselheiro Crispiniano, 40<br>Centro - CEP 01037-000"
hor(2)="Seg. a Sex. 08h30 as 11h30 e das 13h00 as 16h30"
ende(3)="Av. Professor Jos� Barreto, 111 - Cotia"
hor(3)="Seg. a Qui. 09h00 as 15h00"
ende(4)="Rua Francisco Morato, 226 - Vianeio - Jundiai"
hor(4)="Seg. a Sex. 08h30 as 17h00"
ende(5)="Rua Capit�o Tiago Luz, 113 - 1� andar - Salas 04/05/06<br>Santo Amaro"
hor(5)="Seg. a Sex. 08h00 as 12h00 e das 14h00 as 15h45"
ende(6)="Rua Marechal Deodoro, 1301 - Conjunto 2 - Centro<br>S�o Bernardo do Campo"
hor(6)="Seg. a Sex. 08h00 as 17h00"
ende(7)="Rua Vergueiro, 3215 - Vila Mariana<br>Pr�ximo ao Metro Vila Mariana e esquina com Lins de Vasconcelos"
hor(7)="Seg. a Sex. 08h00 as 11h00 e das 13h00 as 16h30"
ende(8)="Av. Padre Lebret, 766 - Morumbi"
hor(8)="Seg. a Sex. 08h00 as 11h30 e das 13h30 as 16h45<br>Os exames de Raio-X e Audiometria somente das 8h00 as 11h30"
%>
<div align="center">
<center>
<table border="0" bordercolor="#000000"cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td height="120px" align="left" valign="top" style="border-left:1px solid;border-top:1px solid;border-bottom:0px solid"><img src="../images/angular_novo.png" width="350px" border="0"></td>
	<td align="center" valign="middle" style="border-right:1px solid;border-top:1px solid;border-bottom:0px solid">
	<b>Guia de Encaminhamento<br>Comentada</b><br>
	<br><font size="1">Leia com aten��o as informa��es para<br>exame m�dico</font>.
	</td>
</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campo" height="40px" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	Empresa <b>(1)</b>: </td>
	<td class="campop" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</td>
	<td class="campo" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	Nome da Obra <b>(2)</b>: </td>
	<td class="campop" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	---------------</td>
</tr>
<tr>
	<td class="campo" height="40px" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid" nowrap>
	Nome do Funcion�rio <b>(3)</b>: </td>
	<td class="campop" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	<%=nome%></td>
	<td class="campo" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid" nowrap>
	Data de Nascimento <b>(4)</b>: </td>
	<td class="campop" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	<%=nascimento%></td>
</tr>
<tr>
	<td class="campo" height="40px" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	Fun��o <b>(5)</b>: </td>
	<td class="campop" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	<%=funcao%></td>
	<td class="campo" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	Setor <b>(6)</b>: </td>
	<td class="campop" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	<%=setor%></td>
</tr>
</table>	

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campo" height="40px" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	N� do RG <b>(7)</b>: </td>
	<td class="campop" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	<%=rg%></td>
	<td width="60%" class="campo" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	Assinatura e Carimbo da Empresa <b>(8)</b>: </td>
</tr>	
</table>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<%if temp=8 then%>
<tr>
	<td class="campo" height="35px" align="center" width="50%" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	<b>Tipo de Exame (9)</b>: </td>
	<td class="campo" colspan="2" align="center" width="50%" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	<b>Exames Complementares (10)</b>: </td>
</tr>	
<tr>
	<td class="campo" align="left" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	( ) Admissional<br>	( ) Demissional<br>	( ) Peri�dico<br>	( ) Mudan�a de Fun��o<br>	( ) Retorno ao Trabalho (agendado)
	</td>
	<td class="campo" align="left" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	( ) Exame Cl�nico<br>	( ) Audiometria<br>	( ) Hemograma<br>	( ) Urina<br>	( ) PPF (parasitol�gico)<br>	( ) Raio X T�rax PA<br>
	( ) EEG (Agendado)<br>	( ) ECG (Agendado)<br>	( ) Glicemia<br>	( ) Espirometria<br>	( ) Acuidade Visual<br>
	<br>
	</td>
	<td class="campo" align="left" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	( ) Acido Hip�rico<br>	( ) Ac. Metil-hip�rico<br>	( ) Avalia��o Psicol�gica<br>	( ) Avalia��o Psicossocial<br>
	( ) TGO<br> ( ) TGP<br> ( ) Gama GT<br> ( ) Cultura de Fezes (coprocultura)<br> ( ) Micol�gico de unha<br> ( ) VDRL
	<br>
	</td>
</tr>
<%else%>
<tr>
	<td class="campo" align="left" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	Tipo de Exame <b>(9)</b>: </td>
	<td class="campo" align="left" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	( ) Admissional<br>	( ) Demissional<br>	( ) Peri�dico<br>	( ) Mudan�a de Fun��o<br>	( ) Retorno ao Trabalho (agendado)
	</td>
	<td class="campo" align="left" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	Exames Complementares <b>(10)</b>: </td>
	<td class="campo" align="left" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	( ) Exame Cl�nico<br>	( ) Audiometria<br>	( ) Hemograma<br>	( ) Urina<br>	( ) PPF<br>	( ) Raio X T�rax PA<br>
	( ) EEG (Agendado)<br>	( ) ECG (Agendado)<br>	( ) Glicemia<br>	( ) Espirometria<br>	( ) Acuidade Visual<br>
	( ) Acido Hip�rico<br>	( ) Ac. Metil-hip�rico<br>	( ) Avalia��o Psicol�gica<br>	( ) Avalia��o Psicossocial<br>
	<br>
	</td>
</tr>
<%end if%>
</table>	

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td rowspan="3" width="100px" class="campo" align="left" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	<br>Informa��es para realiza��o do Exame M�dico na Cl�nica AngularMed:
	</td>
	<td class="campo" height="55px" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	Local:
	</td>
	<td class="campop" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	<%=ende(temp)%>
	</td>
</tr>	
<tr>
	<td class="campo" height="40px" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	Hor�rio de Atendimento:
	</td>
	<td class="campop" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	<%=hor(temp)%>
	</td>
</tr>	
<tr>
	<td class="campo" height="55px" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	Observa��es:
	</td>
	<td class="campop" align="left" valign="middle" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
	<%if temp=8 then%>
	<ul>
	<li>Portar documento RG ou CNH e esta guia de acompanhamento em m�os. A guia tamb�m pode ser enviada para o email: recepcao@angularmed.com.br</li>
	<li>No caso de retorno de Afastamento por Aux�lio Doen�a, deve-se trazer o laudo com alta do m�dico que acompanhou o afastamento.</li>
	<li>Temos estacionamento no local sem cobran�a, ou em frente ao Hospital Albert Einstein, n� 668, com tarifa.</li>
	<li>Os exames de EEG e ECG s�o realizados com agendamento pr�vio.</li>
	</ul>
	<%else%>
	Portar documento RG ou CNH e esta guia em m�os.<br>
	Atendimento por Ordem de Chegada.
	<%end if%>
	</td>
</tr>
<tr>
	<td class="campo" align="left" valign="middle" style="border-bottom:1px solid;" colspan="3">
</tr>
</table>	
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class="campor" align="center" bordercolor="#000000" style="border:0px solid">
	<img src="../images/angular_novo.png" width="120px" border="0"><br>
	Avenida Padre Lebret, 766 - Morumbi - S�o Paulo - SP - 05653-160<br>
	www.angularmed.com.br  -  (11) 3721-6268
	</td>
</tr>
</table>

<br>


<%
set rs=nothing
%>
</table>
<%
set rs=nothing
set rs2=nothing
end if ' 

conexao.close
set conexao=nothing
%>
</body>
</html>