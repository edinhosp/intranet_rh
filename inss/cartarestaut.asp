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
<title>Carta de Restituição de INSS</title>
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
set rsi=server.createobject ("ADODB.Recordset")
Set rsi.ActiveConnection = conexao
teste=0

if request("codigo")<>"" or request.form<>"" then
	if request.form="" then temp=request("codigo")
	if request("codigo")="" then temp=request.form("id_autonomo")
	info=1
	sqlb="AND id_autonomo=" & temp & " "

	if request.form("meses")="" then session("mesrest")=session("mesrest") else session("mesrest")=request.form("meses")
	sqla="SELECT id_autonomo, nome_autonomo, cpf, nit, rg " & _
	"FROM autonomo " & _
	"WHERE id_autonomo>0 "
	
	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	session("id_autonomo")=rs("id_autonomo")
	id=rs("id_autonomo")
	nome=rs("nome_autonomo")
	temp=0
	if rs.recordcount>0 and session("cartateto")<>"L" then temp=2
else
	temp=1
end if
%>

<%
if temp=1 then
session("cartateto")="F"
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>
Seleção do autônomo para emissão de carta de restituição
<form method="POST" action="cartarestaut.asp" name="form">
  <p style="margin-top: 0; margin-bottom: 0">Chapa/Nome 
  <select name="id_autonomo">
<%
sql="SELECT id_autonomo, nome_autonomo FROM autonomo ORDER BY nome_autonomo "
rsi.Open sql, ,adOpenStatic, adLockReadOnly
rsi.movefirst
do while not rsi.eof
%>
  	<option value="<%=rsi("id_autonomo")%>"><%=rsi("nome_autonomo")%></option>
<%
rsi.movenext
loop
rsi.close
%>
  </select></p>
  <p style="margin-top: 0; margin-bottom: 0">Quantidade de meses para listar <input type="text" name="meses" size="5" value="12"></p>
  <p style="margin-top: 0; margin-bottom: 0">
  <input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>

<%
elseif temp=0 then
session("cartateto")="C"
'if request.form<>"" then

%>
<table border="0" cellpadding="5" width="620" cellspacing="0" height="1000">
  <tr><td width="100%"><img border="0" src="../images/aguia.jpg"></td></tr>

  <tr><td width="100%">&nbsp;</td></tr>

  <tr><td width="100%" align="center"><b><font size="4">DECLARAÇÃO</font></b></td></tr>

  <tr><td width="100%">
      <p>&nbsp;</p>
      <p align="justify">Declaramos para os devidos fins, que o(a) Sr(a). <%=rs("nome_autonomo")%>,
      inscrito no PIS/PASEP/NIT sob nº <%=rs("nit")%>, 
      do C.P.F. nº <%=rs("cpf")%> e do R.G. nº <%=rs("rg")%>,
      prestou serviços de natureza eventual a esta Instituição de Ensino Superior.</p>
      <p align="justify">Esclarecemos ainda que descontamos, recolhemos e não
      devolvemos as contribuições abaixo mencionadas para o(a) referido(a) contribuinte,
      e não compensamos a importância em GRPS nem pleiteamos a restituição
      junto ao INSS.</p>
<%
sqlv="SELECT top " & session("mesrest") & " id_autonomo, Year([data_pagamento]) AS ano, Month([data_pagamento]) AS mes, " & _
"Sum([servico_prestado]+[outros_rendimentos]) AS rend, Sum(desconto_inss) AS inss " & _
"FROM autonomo_rpa " & _
"WHERE ((data_pagamento Is Not Null)) " & _
"GROUP BY id_autonomo, Year([data_pagamento]), Month([data_pagamento]) " & _
"HAVING ((id_autonomo=" & rs("id_autonomo") & ")) " & _
"ORDER BY Year([data_pagamento]) DESC , Month([data_pagamento]) DESC "

rsi.Open sqlv, ,adOpenStatic, adLockReadOnly
quant=rsi.recordcount
resto=quant mod 12
if resto=0 then resto=0 else resto=1
colunas=int(quant/12) + resto
%>
      <div align="center">
        <center>
        <table border="0" cellspacing="0">
          <tr>
<%
rsi.movefirst
'do while not rsi.eof
%>
<% for a=1 to colunas %>
    <td class=campo colspan=3 valign=top>
	<table border="1" cellpadding="3" cellspacing="1" style="border-collapse: collapse">
	<tr><td class=titulo align="center">Competência</td>
        <td class=titulo align="center">Salário de Contribuição</td>
        <td class=titulo align="center">Valor de INSS</td></tr>

<!-- tabela -->			
<% 
if a=colunas then final=quant else final=a*12
for b=a*12-11 to final
rsi.absoluteposition=b
%><tr>
            <td class=campo align="center">&nbsp;<%=numzero(rsi("mes"),2) & "/" & rsi("ano") %></td>
            <td class=campo align="center">&nbsp;<%=formatnumber(rsi("rend"),2) %></td>
            <td class=campo align="center">&nbsp;<%=formatnumber(rsi("inss"),2) %></td></tr>
<%next %>			
</table>
</td>
<%next %>
<%
'rsi.movenext
'loop
%>
<!-- tabela -->			
          </tr>
<%
rsi.close
%>
        </table>
        </center>
      </div>
    </td></tr>

  <tr><td width="100%">
    <table border="0" cellpadding="0" width="100%" cellspacing="0">
      <tr><td width="50%" valign="top">
      <p><font size="2">Osasco,&nbsp;<%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %></font></p>
      <p>Atenciosamente</p>
      <p>&nbsp;</p>
      <p><font size="2">_____________________________________<br>
      </font><input type="text" name="nome" size="70" maxlength="256" class=form_input></p>
      </td>

<%if teste=1 then %>
        <td width="50%" valign="top">&nbsp;
          <div align="center">
            <center>
            <table border="0" cellpadding="0" width="240" cellspacing="0">
              <tr>
                <td width="1" style="border-left-style: solid; border-left-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
                <td width="240" rowspan="2">
                  <p align="center"><b><font size="4" color="#808080">73.063.166/0001-20</font></b></td>
                <td width="1" style="border-right-style: solid; border-right-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
              </tr>
              <tr>
                <td width="1"></td>
                <td width="1"></td>
              </tr>
              <tr>
                <td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
                <td width="240"></td>
                <td width="1"></td>
              </tr>
              <tr>
                <td width="1"></td>
                <td width="240">
                  <p align="center"><b><font color="#808080">FUNDAÇÃO INSTITUTO DE<br>
                  ENSINO PARA OSASCO</font></b></td>
                <td width="1"></td>
              </tr>
              <tr>
                <td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
                <td width="240"></td>
                <td width="1"></td>
              </tr>
              <tr>
                <td width="1">&nbsp;</td>
                <td width="240" rowspan="2">
                  <p align="center"><font color="#808080">Rua Narciso Sturlini, 883<br>
                  Jd. Umuarama - CEP 06018-903<br>
                  OSASCO - SP</font></td>
                <td width="1"></td>
              </tr>
              <tr>
                <td width="1" style="border-left-style: solid; border-left-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
                <td width="1" style="border-right-style: solid; border-right-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
              </tr>
            </table>
            </center>
          </div>
          <p>&nbsp;
<%end if%>
         </td>
      </tr>
    </table>
    </td>
  </tr>

  <tr><td>&nbsp;</td></tr>

  <tr><td>&nbsp;</td></tr>

  <tr>
    <td height="15">
      <b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b></td>
  </tr>
  <tr>
    <td height="15">
      <font face="Arial Narrow">R. Narciso Sturlini, 883 - Osasco - SP - CEP
      06018-903 - Fone: (011) 3681-6000<%if teste=0 then response.write " - C.N.P.J. 73.063.166/0001-20" %></font></td>
  </tr>
  <tr>
    <td height="15">
      <font face="Arial Narrow">Av. Franz Voegeli, 300 - Osasco - SP - CEP
      06020-190 - Fone: (011) 3651-9999<%if teste=0 then response.write " - C.N.P.J. 73.063.166/0003-92" %></font></td>
  </tr>
<%if teste=0 then%>
  <tr>
    <td height="15">
      <font face="Arial Narrow">Av. Franz Voegeli, 1743 - Osasco - SP - CEP
      06020-190 - Fone: (011) 3654-0655 - C.N.P.J. 73.063.166/0004-73</font></td>
  </tr>
<%end if%>
  <tr><td height="15">
      <font face="Arial Narrow">Caixa Postal - ACF - Jd. Ipê - nº 2530 -
      Osasco - SP - CEP 06053-990</font>
    </td></tr>

</table>

<%
rs.close
set rs=nothing
elseif temp=2 then
session("cartateto")="L"
%>
<!-- mostrar funcionarios e as contribuições -->
<table border="1" cellpadding="0" width="550" cellspacing="0">
  <tr>
    <td class=titulo>&nbsp;Código</td>
    <td class=titulo>&nbsp;Nome</td>
  </tr>
<%
rs.movefirst
do while not rs.eof
%>
  <tr>
    <td class=campo>&nbsp;<%=rs("id_autonomo")%></td>
    <td class=campo>&nbsp;<a href="cartarestaut.asp?codigo=<%=rs("id_autonomo")%>"><%=rs("nome_autonomo")%></a></td>
  </tr>
<%
rs.movenext
loop
%>

</table>
<%
rs.close
set rs=nothing
end if ' temps
%>
</body>
</html>
<%
conexao.close
set conexao=nothing
%>