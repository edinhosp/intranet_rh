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
dim endereco(10), tipo(10), extra(10), imagem(10), end1(10), end2(10), end3(10),textoimagem(10)
tipo(0)="P/A/D/MF" : endereco(0)="Osasco - Rua Itabuna" : extra(0)="" : imagem(0)="angular_semlogo.png"
tipo(1)="P/A/D/MF" : endereco(1)="Itapetininga" : extra(1)="Audiometria" : imagem(1)="angular_semlogo.png"
tipo(2)="P/A/D/MF" : endereco(2)="S�o Bernardo do Campo - Centro" : extra(2)="" : imagem(2)="angular_marechal.png":textoimagem(2)="<b>Instituto Marechal</b><br>Assessoria em Medicina e Seguran�a do Trabalho"
tipo(3)="P/A/D/MF" : endereco(3)="S�o Paulo - Ipiranga" : extra(3)="Audiometria" : imagem(3)="angular_semlogo.png"
tipo(4)="P/A/D/MF" : endereco(4)="S�o Paulo - Santo Amaro" : extra(4)="Audiometria/V�rios" : imagem(4)="angular_mta.png"
tipo(5)="P/A/D/MF" : endereco(5)="S�o Paulo - Vila Mariana" : extra(5)="Audiometria" : imagem(5)="angular_semlogo.png"
tipo(6)="N�o Realiza" : endereco(6)="S�o Paulo - Centro" : extra(6)="Audiometria/V�rios" : imagem(6)="angular_semlogo.png"
tipo(7)="P/A/D/MF" : endereco(7)="S�o Paulo - Barra Funda" : extra(7)="Audiometria" : imagem(7)="angular_semlogo.png"
tipo(8)="P/A/D/MF" : endereco(8)="S�o Paulo - Brooklin" : extra(8)="Audiometria" : imagem(8)="angular_semlogo.png"
tipo(9)="P/A/D/MF" : endereco(9)="Cotia - Centro" : extra(9)="Audiometria" : imagem(9)="angular_semlogo.png"
tipo(10)="RT" : endereco(10)="S�o Paulo - Morumbi" : extra(10)="" : imagem(10)="angular_semlogo.png"
tipo(0)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o"
tipo(1)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o"
tipo(2)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o"
tipo(3)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o"
tipo(4)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o"
tipo(5)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o"
tipo(6)="<font color=red>N�o Realiza</font>"
tipo(7)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o"
tipo(8)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o"
tipo(9)="Per�odico / Admissional / Demissional / Mudan�a de Fun��o"
tipo(10)="<font color=blue>Retorno ao Trabalho</font>"

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
for a=0 to 10
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
	nome=space(40) & "&nbsp;"
	nascimento=space(20) & "&nbsp;"
	idade="&nbsp;"
	setor="&nbsp;"
	funcao="&nbsp;"
	rg="&nbsp;"
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
%>
<div align="center">
<center>
<table border="0" bordercolor="#000000"cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr>
	<td height="80px" align="left" valign="top" style="border-left:1px solid;border-top:1px solid;border-bottom:0px solid"><img src="../images/angular.png" border="0"></td>
	<td align="right" valign="top" style="border-right:1px solid;border-top:1px solid;border-bottom:0px solid"><img src="../images/<%=imagem(temp)%>" border="0"><%=textoimagem(temp)%></td>
</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr>
	<td class="campop" align="center" valign="top" style="border-top:1px solid;border-left:1px solid;border-right:1px solid">
<!-- quadro com os dados -->
	<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="598">
	<tr>
		<td height="30px" width="10" class="campop">&nbsp;</td>
		<td class="campop" width="435">&nbsp;<%=nome%></td>
		<td class="campop" width="15">&nbsp;</td>
		<td class="campop" width="120" align="center">&nbsp;<%=nascimento%></td>
		<td class="campop" width="15">&nbsp;</td>
	</tr>
	<tr>
		<td height="5px" width="10" class="campor"></td>
		<td class="campor" style="border-top:1px solid"><b>NOME DO FUNCION�RIO</td>
		<td class="campor">&nbsp;</td>
		<td class="campor" align="center" style="border-top:1px solid"><b>DATA DE NASCIMENTO</td>
		<td class="campor">&nbsp;</td>
	</tr>
	<tr><td class="campo" height="10"></td></tr>
	</table>

	<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="598">
	<tr>
		<td height="30px" width="10" class="campop">&nbsp;</td>
		<td class="campop" width="200">&nbsp;<%=setor%></td>
		<td class="campop" width="15">&nbsp;</td>
		<td class="campop" width="200">&nbsp;<%=funcao%></td>
		<td class="campop" width="15">&nbsp;</td>
		<td class="campop" width="140" align="center">&nbsp;<%=rg%></td>
		<td class="campop" width="15">&nbsp;</td>
	</tr>
	<tr>
		<td height="5px" width="10" class="campor"></td>
		<td class="campor" style="border-top:1px solid"><b>SETOR</td>
		<td class="campor">&nbsp;</td>
		<td class="campor" align="center" style="border-top:1px solid"><b>FUN��O</td>
		<td class="campor">&nbsp;</td>
		<td class="campor" align="center" style="border-top:1px solid"><b>R.G.</td>
		<td class="campor">&nbsp;</td>
	</tr>
	<tr><td class="campo" height="10"></td></tr>
	</table>
	
	<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="598">
	<tr>
		<td height="50" width="10" class="campor"></td>
		<td class="campo" valign="top">
		Empresa: <b>FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</b>
		<br>CNPJ: <%=cnpj%>
		</td>
		<td class="campop" width="15">&nbsp;</td>
		<td class="campo" valign="bottom" style="border-bottom:0px solid">
		<%if teste=1 then%>
		<img src="../images/assi_edson.jpg" width=150>
		<%end if%>
		____________________________________<br>
		Carimbo ou Nome do Respons�vel
		</td>
		<td class="campop" width="15">&nbsp;</td>
	</tr>
	<tr><td class="campo" height="10"></td></tr>
	</table>
<%
if temp<>10 then fonteRT="gray" else fonteRT="black"
%>
	<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="598">
	<tr>
		<td class="campo" width="15">&nbsp;</td>
		<td class="campor" valign="top" align="left"><b><u>TIPO DE EXAME</u></td>
		<td class="campo" width="15">&nbsp;</td>
		<td class="campor" valign="top" align="left"><b><u>MOTIVO</u></td>
		<td class="campo" width="15">&nbsp;</td>
	</tr>
	<tr>
		<td class="campo" width="15">&nbsp;</td>
		<td class="campo" valign="top" align="left">
<%if temp<>6 then%>
		<img src="../images/bullet.gif"> ASO (Atestado de Sa�de Ocupacional)
<%
end if
if temp=1 or temp=3 or temp=4 or temp=5 or temp=6 or temp=7 or temp=8 or temp=9 then
		response.write "<br><img src=""../images/bullet.gif""> AUDIOMETRIA"
end if
if temp=4 then
		response.write "<br><img src=""../images/bullet.gif""> Tipagem Sangu�nea"
		response.write "<br><img src=""../images/bullet.gif""> Fator RH"
end if
if temp=4 or temp=6 then
		response.write "<br><img src=""../images/bullet.gif""> Hemograma"
		response.write "<br><img src=""../images/bullet.gif""> Acuidade Visual"
		response.write "<br><img src=""../images/bullet.gif""> Eletroencefalograma (EEG)"
		response.write "<br><img src=""../images/bullet.gif""> Eletrocardiograma (ECG)"
		response.write "<br><img src=""../images/bullet.gif""> Espirometria"
		response.write "<br><img src=""../images/bullet.gif""> Glicemia/Glicose"
end if
if temp=6 then
		response.write "<br><img src=""../images/bullet.gif""> RX Torax"
		response.write "<br><img src=""../images/bullet.gif""> Acido Hipurico"
		response.write "<br><img src=""../images/bullet.gif""> Methilhip�rico"
		response.write "<br><img src=""../images/bullet.gif""> Eritograma"
end if

%>		
		<br><img src="../images/bullet.gif"> ____________________________
		
		
		</td>
		<td class="campo" width="15">&nbsp;</td>
		<td class="campo" valign="top" align="left">
		<img src="../images/bullet.gif"> Peri�dico
		<br><img src="../images/bullet.gif"> Admiss�o
		<%for a=1 to 5:response.write "&nbsp;":next%> <img src="../images/bullet.gif"> <font color="<%=fonteRT%>">Retorno ao Trabalho</font>
		<br><img src="../images/bullet.gif"> Demiss�o
		<%for a=1 to 5:response.write "&nbsp;":next%>
		<img src="../images/bullet.gif"> Mudan�a de Fun��o
<%
if temp=6 then
	response.write "<br><br><br><br>"
	response.write "<br><img src=""../images/bullet.gif""> Entregar o exame para o funcion�rio"
	response.write "<br><img src=""../images/bullet.gif""> N�o entregar o exame para o funcion�rio"
end if
%>		
		</td>
		<td class="campo" width="15">&nbsp;</td>
	
	</tr>
	</table>

	<!-- quadro com os dados -->
	<bR><br>
	<Br><br><Br>
	</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid;border-right:1px solid">&nbsp;Obs.: O Funcion�rio dever� portar o R.G.
	<%for a=1 to 60:response.write "&nbsp;":next%>Data: ______/______/_______
</tr>
<tr>
	<td class="campo" align="center" bordercolor="#000000" style="border:1px solid">
<%
end1(0)="Telefone: (11) 3184-0099<br><b>Hor�rio de Atendimento:</b><br>Segunda a Quinta das 6:30 as 16:00<br>Sexta das 6:30 as 15:30<br>"
end2(0)="<b>Local do exame:</b> Rua Itabuna, 93 - Centro de Osasco - SP - CEP 06010-120<br>Pr�ximo � Prefeitura de Osasco, passar por baixo do pontilh�o.<br>"
end3(0)="<u>Atendimento por ordem de Chegada!</u><br><br>"
end1(1)="<b>Imedi</b> - Telefone: (15) 3271-7910<br><b>Hor�rio de Atendimento:</b><br>De Segunda a Sexta das 7:30 as 12:00 e das 14:00 as 17:00<br>"
end2(1)="<b>Local do exame:</b> Rua General Carneiro, 217 - Itapetininga<br>Pr�ximo � BR.<br>"
end3(1)="<u>Atendimento por ordem de Chegada!</u><br><br>"
end1(2)="Telefone: 4121-4145 ou 4123-7490<br><b>Hor�rio de Atendimento:</b><br>De Segunda a Sexta das 8:00 as 16:00<br>"
end2(2)="<b>Local do exame:</b> Rua Marechal Deodoro, 1301 - Conj. 2 - Centro<br>S�o Bernardo do Campo - SP<br>"
end3(2)="<u>Atendimento por ordem de Chegada e com esta guia para apresentar!</u><br><br>"
end1(3)="<b>Mesp Medicina</b> - Telefone: 2066-6166<br><b>Hor�rio de Atendimento:</b><br>De Segunda a Quinta das 8:00 as 11:00 e das 13:00 as 16:00<br>"
end2(3)="<b>Local do exame:</b> Rua das Juntas Provis�rias, 406 - Ipiranga - SP<br>Pr�ximo � Fabrica Viscont, ao lado da esta��o Nossa Senhora Aparecida do Expresso Tiradentes<br>"
end3(3)="<u>Ligar para agendar!</u><br><br>"
end1(4)="Telefone: 3805-3514 ou 5524-1370<br><b>Hor�rio de Atendimento:</b><br>De Segunda a Sexta<br>"
end2(4)="<b>Local do exame:</b> Rua Dr. Carlos Augusto Campos, 324 - Santo Amaro - SP<br>Pr�ximo ao McDonalds, altura do n� 540 da Av. Jo�o Dias<br>"
end3(4)="<u>Atendimento com hora marcada!</u><br><br>"
end1(5)="<b>NR9</b> - Telefone: 5574-6266<br><b>Hor�rio de Atendimento:</b><br>De Segunda a Sexta das 8:00 as 11:00 e das 13:00 as 15:00<br>"
end2(5)="<b>Local do exame:</b> Rua Vergueiro, 3215 - Vila Mariana - SP<br>Pr�ximo ao Metro Vila Mariana.<br>"
end3(5)="<u>Atendimento por ordem de Chegada!</u><br><br>"
end1(6)="<b></b> Telefone: 3159-0573<br><b>Hor�rio de Atendimento:</b><br>De Segunda a Sexta das 8:00 as 16:00<br>"
end2(6)="<b>Local do exame:</b> Rua 7 de Abril, 118 - 6� andar - conj. 601 - Centro - SP<br>Pr�ximo � esta��o de Metro Anhangaba�.<br>"
end3(6)="<u>N�O realiza ASO, apenas exames complementares!</u><br><br>"
end1(7)="<b>Prolabor</b> Telefone: 3392-1305<br><b>Hor�rio de Atendimento:</b><br>De Segunda a Quinta das 8:00 as 12:00 e das 14:00 as 17:00<br>"
end2(7)="<b>Local do exame:</b> Av. Thomas Edison, 305 - Barra Funda - SP<br>Pr�ximo � esta��o de Metro Barra Funda.<br>"
end3(7)="<u>Ligar para agendar!</u><br><br>"
end1(8)="<b>S.A.</b> Telefone: 5182-8221<br><b>Hor�rio de Atendimento:</b><br>De Segunda a Quinta das 8:00 as 11:30 e das 13:00 as 16:30<br>Sexta das 8:00 as 11:00<br>"
end2(8)="<b>Local do exame:</b> Rua Joaquim Guarani, 105 - Brooklin - SP<br>Pr�ximo ao Clube Banespa, travessa da Av.Santo Amaro, altura do n� 5200.<br>"
end3(8)="<u>Atendimento por agendamento!</u><br><br>"
end1(9)="<b>SeguraMed</b> Telefone: 4614-0153<br><b>Hor�rio de Atendimento:</b><br>De Segunda a Sexta das 11:00 as 12:00<br>"
end2(9)="<b>Local do exame:</b> Av. Prof. Jos� Barreto, 111 - Centro - Cotia - SP<br>Pr�ximo: em cima das Casas Pernambucanas.<br>"
end3(9)="<u>Ligar para marcar consulta!</u><br><br>"
end1(10)="<b>Cl�nica Angular</b> Telefone: 3721-6268<br><b>Hor�rio de Atendimento:</b><br>De Segunda a Quinta das 9:00 as 12:00 e das 13:30 as 16:30<br>Audiometria: segunda, quarta e quinta das 9:30 as 12:00<br>"
end2(10)="<b>Local do exame:</b> Av. Prof. Francisco Morato, 1956 - 1� andar - Butant� - SP (Proximo � Kalunga)<br>"
end3(10)="<u>Atendimento por ordem de chegada!</u><br><br>"
end1(10)="<b>Angular</b> Telefone: 2367-8192 / 3721-6268<br><b>Hor�rio de Atendimento:</b><br>De Segunda a Quinta das 8:30 as 12:00 e das 13:30 as 16:30<br>"
end2(10)="<b>Local do exame:</b> Av. Padre Lebret, 766 - Morumbi - SP<br>"
end3(10)="<u>Atendimento por ordem de chegada!</u><br><br>"
	response.write end1(temp)
	response.write end2(temp)
	response.write end3(temp)
%>	
	</td>
</tr>
</table>

<br>
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo><img src="../images/tesoura1.gif" width="56" height="38" border="0" alt=""></td>
	<td class=campo width=100%><hr style="border:2px #000000 dotted"></td>
</tr>
</table>

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