<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a80")="N" or session("a80")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Formulários Solicitados pela Previdência Social - BENEFÍCIO</title>
<!--
<link rel="stylesheet" type="text/css" href="../diversos.css">
-->
<link rel="SHORTCUT ICON" href="../images/rho.png">
<style type="text/css">
*{
padding: 0;
margin: 0;
top: 0px;
left: 0px;
}
.formularios-previdencia{
font-size: 11px;
width: 690px;	
margin: 0 auto 0 auto; 
font-family: Arial, Helvetica, sans-serif;
}
.formularios-previdencia h2{
color: #00008B;
text-align: center;
font-size: 16px;
}
.formularios-previdencia .observacao{	
text-align: center;
color: brown;
font-size: 12px;
}
.formularios-previdencia table{	
width: 100%;
}
.formularios-previdencia table tr td{	
/*border: #000 2px solid;*/
padding: 5px;
text-align: left;
vertical-align: text-top;
background: white;
}
.formularios-previdencia table td.titulo{	
color: #333;
background: #EEE;
text-align: center;
font-size: 20px;
font-weight: bold;
}
.formularios-previdencia table td input[type=text]{	
display: block;
position:relative;
width: 95%;
font-size: 16px;
font-style: normal;
}
#logo{
clear: both;
width: 285px;
height: 58px;
display: block;
margin: 0 auto 0 auto;
}
.quebra_de_pagina{
width: 690px;	
margin: 0 auto 0 auto;
text-align: center; 
}
.quebra_de_pagina p{
font-size: 10px;
color: #333;
margin: 0;
padding: 0; 
}
div.pagina{
z-index: -10;
/*background: #CCC;*/
width: 100%;
height: 1100px;
position: absolute;
}
.novapagina {
page-break-before: always;
}
.cl_divInstrucoes{
width: 95%;
text-align: justify;
}
</style>

</head>
<body>

<div class = "pagina"></div>
<div id="logo"><img src="topo_nome_mps.gif" /></div>

<div class="formularios-previdencia">
<br /><br />

<table>
<tr> 
	<td colspan="7" style="text-align: center;">
	<b>REQUERIMENTO DE BENEF&Iacute;CIO POR INCAPACIDADE</b>
	</td>
</tr>
<tr> 
	<td style="text-align: right"> Nome: </td>
	<td colspan="6"> <input type="text" name="textfield" size="80"> </td>
</tr>
<tr>
	<td nowrap style="text-align: right"> Data  de Nascimento: </td>
	<td colspan="2"> <input type="text" name="textfield2"> </td>
	<td style="text-align: right"> Nacionalidade: </td>
	<td colspan="3"> <input name="textfield3" type="text" size="20"> </td>				
</tr>
<tr> 
	<td style="text-align: right"> Rua/Av. </td>
	<td colspan="6"> <input name="textfield4" type="text" size="80"> </td>
</tr>
<tr>
	<td style="text-align: right"> Complemento </td>
	<td colspan="2"> <input name="textfield42" type="text" size="20"> </td>				
	<td style="text-align: right"> Cidade: </td>
	<td colspan="4"> <input name="textfield44" type="text" size="20"> </td>
</tr>
<tr> 
	<td style="text-align: right">  Bairro: </td>
	<td colspan="2"> <input name="textfield43" type="text" size="20"> </td>
	<td style="text-align: right"> Estado: </td>
	<td colspan="4"> <input name="textfield45" type="text" size="20"> </td>
</tr>
<tr> 
	<td style="text-align: right"> Sexo: </td>
	<td nowrap> M:  <input type="checkbox" name="checkbox" value="checkbox"> &nbsp; | &nbsp; F:  <input type="checkbox" name="checkbox2" value="checkbox"> </td>
	<td colspan="2" style="text-align: right"> CEP:</td>
	<td colspan="3"> <input name="textfield452" type="text" size="10"> </td>
</tr>
<tr> 
	<td colspan="2" style="text-align: right"> <b>DOC.  INSCRI&Ccedil;&Atilde;O - (N&ordm; e S&eacute;rie):</b> </td>
	<td colspan="5"> <input type="text" name="textfield4532"> </td>
</tr>
<tr> 
	<td height="21" style="text-align: right"> Estado Civil </td>
	<td > <input type="checkbox" name="checkbox3" value="checkbox"> Solteiro </td>
	<td colspan="2"> <input type="checkbox" name="checkbox5" value="checkbox"> Casado </td>
	<td colspan="3"> <b>TEM  OUTRA ATIVIDADE COM VINCULA&Ccedil;&Atilde;O &Agrave; PREVID&Ecirc;NCIA SOCIAL ? </b></td>
</tr>
<tr>
	<td height="25">&nbsp;  </td>
	<td > <input type="checkbox" name="checkbox4" value="checkbox"> Vi&uacute;vo </td>
	<td colspan="2"> <input type="checkbox" name="checkbox6" value="checkbox"> Desq/Divor </td>
	<td colspan="3"> <input type="checkbox" name="checkbox52" value="checkbox"> Sim &nbsp; | &nbsp; <input type="checkbox" name="checkbox53" value="checkbox"> N&atilde;o </td>
</tr>
<tr valign="middle"> 
	<td height="89" colspan="7" style="text-align: center;">
	<br /><br /><br /><br /><br />
	<b>ASSINATURA DO REQUERENTE ______________________________________________________</b>
	</td>
</tr>
<tr> 
	<td colspan="2" style="text-align: right"> NOME  DO PROCURADOR OU CURADOR: </td>
	<td colspan="5"> <input name="textfield46" type="text" size="60"> </td>
</tr>
<tr> 
	<td colspan="2" style="text-align: right"> ENDERE&Ccedil;O: </td>
	<td colspan="5"> <input name="textfield47" type="text" size="60"> </td>
</tr>
<tr> 
	<td colspan="7">&nbsp;  </td>
</tr>
<tr>
	<td colspan="7" style="text-align: center;">
	<b>ATESTADO DE AFASTAMENTO DO TRABALHO</b>
	</td>
</tr>
<tr> 
	<td style="text-align: right"> EMPRESA </td>
	<td colspan="2"> <input type="text" name="textfield4533"> </td>
	<td style="text-align: right"> N&ordm;  CNPJ: </td>
	<td colspan="3"> <input type="text" name="textfield45332"> </td>
</tr>
<tr> 
	<td style="text-align: right"> RUA/AV. </td>
	<td colspan="2"> <input type="text" name="textfield4534"> </td>
	<td style="text-align: right"> N&ordm;: </td>
	<td colspan="3"> <input name="textfield45333" type="text" size="5"> </td>
</tr>
<tr> 
	<td style="text-align: right"> COMPLEMENTO </td>
	<td colspan="2"> <input type="text" name="textfield4535"> </td>
	<td style="text-align: right"> BAIRRO: </td>
	<td colspan="3"> <input name="textfield45334" type="text" size="20"> &nbsp; </td>
</tr>
<tr> 
	<td style="text-align: right"> CIDADE </td>
	<td colspan="2"> <input type="text" name="textfield4536"> </td>
	<td style="text-align: right"> ESTADO: </td>
	<td colspan="3"> <input type="text" name="textfield45335">  &nbsp; </td>
</tr>
<tr> 
	<td style="text-align: right"> CEP: </td>
	<td colspan="2"> <input name="textfield453352" type="text" size="10"> </td>
	<td style="text-align: right"> <span style="font-weight: bold; color: red;"> CID: </span> </td>
	<td colspan="3"> <input type="text" name="textfield453353"> </td>
</tr>
<tr> 
	<td colspan="3" style="text-align: right"> <b> &Uacute;LTIMO DIA DE TRABALHO DO SEGURADO: </b> </td>
	<td colspan="4"> <input type="text" name="textfield4537"> </td>
</tr>
<tr bgcolor="#CCCCCC"> 
	<td colspan="7">&nbsp;  </td>
</tr>
<tr> 
	<td style="text-align: right"> <b>AFASTADO  POR:</b> </td>
	<td> <input type="checkbox" name="checkbox32" value="checkbox"> DOEN&Ccedil;A </td>
	<td nowrap> <input type="checkbox" name="checkbox34" value="checkbox"> ACIDENTE DO TRABALHO </td>
	<td> <input type="checkbox" name="checkbox33" value="checkbox"> F&Eacute;RIAS </td>
	<td colspan="3"> <input type="checkbox" name="checkbox35" value="checkbox"> ACIDENTE DE QUALQUER NATUREZA </td>
</tr>
<tr> 
	<td colspan="7">&nbsp;  </td>
</tr>
</table>
</div>

<div class="novapagina"></div>		
<div class="formularios-previdencia">

<table>
<tr> 
	<td colspan="4" style="text-align: center;">
		<b>DEPENDENTES PARA SAL&Aacute;RIO FAM&Iacute;LIA </b></td>
</tr>
<tr> 
	<td colspan="4">&nbsp;  </td>
</tr>
<tr> 
	<td height="21" style="text-align: center; width: 10%;"> PRENOME  DOS FILHOS </td>
	<td style="text-align: center; width: 10%;">DATA  NASC. </td>
	<td style="text-align: center; width: 10%;"> PRENOME DOS FILHOS </td>
	<td style="text-align: center; width: 10%;"> DATA NASC. </td>
</tr>
<tr > 
	<td style="text-align: center; width: 10%;"> <input name="textfield4537235" type="text" size="20"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield453722" type="text" size="10"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield45372355" type="text" size="20"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield4537225" type="text" size="10"> </td>
</tr>
<tr> 
	<td style="text-align: center; width: 10%;"> <input name="textfield45372352" type="text" size="20"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield4537222" type="text" size="10"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield45372356" type="text" size="20"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield4537226" type="text" size="10"> </td>
</tr>
<tr> 
	<td style="text-align: center; width: 10%;"> <input name="textfield45372353" type="text" size="20"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield4537223" type="text" size="10"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield45372357" type="text" size="20"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield4537227" type="text" size="10"> </td>
</tr>
<tr> 
	<td style="text-align: center; width: 10%;"> <input name="textfield45372354" type="text" size="20"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield4537224" type="text" size="10"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield45372358" type="text" size="20"> </td>
	<td style="text-align: center; width: 10%;"> <input name="textfield4537228" type="text" size="10"> </td>
</tr>
<tr> 
	<td colspan="7">&nbsp;  </td>
</tr>
<tr> 
	<td style="text-align: right;"> LOCALIDADE: </td>
	<td> <input name="textfield45372359" type="text" size="20"> </td>
	<td style="text-align: right;"> DATA: </td>
	<td> <input name="textfield45372282" type="text" size="10">  </td>
</tr>
<tr> 
	<td height="95" >&nbsp;  </td>
	<td  colspan="5" valign="bottom" style="text-align: center;">
	<br /><br /><br /><br /><br />
	_________________________________________________________________________<br />
	<b>ASSINATURA DO RESPONS&Aacute;VEL E CARIMBO DO CGC DA EMPRESA </b>
	</td>
</tr>
<tr> 
	<td colspan="7"><hr align="center"></td>
</tr>
<tr> 
	<td colspan="7">
		<div align="center">
		<b>I N S T R U &Ccedil; &Otilde; E S </b>
		</div>
		<div class="cl_divInstrucoes">
		1  - O requerimento deve ser sem rasuras e preenchido de prefer&ecirc;ncia &agrave; m&aacute;quina.<br />
		2 - No caso de segurado empregado, a empresa &eacute; respons&aacute;vel pelo preenchimento Atestado de Afastamento do Trabalho<br />
		3 - No m&ecirc;s do afastamento do trabalho a empresa efetuar&aacute; o pagamento integral do Sal&aacute;rio - Fam&iacute;lia, e o INSS far&aacute; o mesmo no m&ecirc;s da cessa&ccedil;&atilde;o do benef&iacute;cio, evitando-se assim, c&aacute;lculo de valores fracionados.
		</div>
	</td>
</tr>
</table>
</div>
<br />
</body>
</html>
