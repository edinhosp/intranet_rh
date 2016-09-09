<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.redirect "intranet.asp"
if session("a1")="N" or session("a1")="" then response.redirect "intranet.asp"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Pesquisa Curriculos</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
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
</head>
<body>
<!-- -->
<table><tr><td>
<!-- -->
<form method="POST" action="banco_funcao.asp" name="form">

<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("MySQLfieo")
'conexao.open "Driver={MySQL ODBC 3.51 Driver}; Server=localhost; Port=3306; Option=0; Socket=; Stmt=; Database=rhonline2; Uid=root; Pwd="
'conexao.open "Driver={MySQL ODBC 3.51 Driver}; Server=colossus2.fieo.br; Port=3306; Option=0; Socket=; Stmt=; Database=website; Uid=rh; Pwd=!@#qaz"

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
rs.CursorLocation=3

%>

<table border="1" cellpadding="1" width="690" cellspacing="0" style="border-collapse: collapse">
<%
sqla="SELECT distinct p.funcao, f.nome_funcao, count(cpf) as curriculos FROM tb_rh_pretensao p " & _
"inner join tb_rh_funcao f on f.id_funcao=p.funcao group by p.funcao, f.nome_funcao order by f.nome_funcao"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<tr>
	<td class=titulo>Funções <%=rs.recordcount%></td>
	<td class=titulo>Curriculos</td>
</tr>
<tr>
	<td class="campor" valign=top width=150>
<%
rs.movefirst
do while not rs.eof
if cstr(rs("funcao"))=cstr(request("funcao")) then estilo="style='background:silver;'" else estilo=""
if cstr(rs("funcao"))=cstr(request("funcao")) and nomef="" then nomef=rs("nome_funcao")
%>
<a <%=estilo%> class=r href="banco_funcao.asp?funcao=<%=rs("funcao")%>"><%=rs("nome_funcao") & " (" & rs("curriculos") & ")"%></a><br>
<%
rs.movenext
loop
rs.close
%>
	</td>
	<td class="campor" valign=top><%=nomef%>
<%
if request("funcao")<>"" then
sqlb="SELECT c.nome, p.funcao, p.cpf, p.salario, p.habilidade, data_cadastro, nascimento, bairro, cidade, uf, tel_residencial, tel_celular, email, observacoes " & _
"FROM tb_rh_pretensao p inner join tb_rh_candidato c on c.cpf=p.cpf where p.funcao=" & request("funcao") & " order by data_cadastro desc "
rs.Open sqlb, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
<table border=1 width=530 style='border-collapse:collapse'>
<tr>
	<td class=campo valign=top><font color="green"><b>Nome:</b></font><br><b><%=ucase(rs("nome"))%></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<a class=r href="banco_curriculo.asp?codigo=<%=rs("cpf")%>" onclick="NewWindow(this.href,'form_curriculo','695','450','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Leaf.gif" width="16" height="16" border="0" alt="Visualizar o curriculo"></a></td>
	<td class=campo valign=top width=70><font color="green"><b>Cadastro:</b></font><br><%=rs("data_cadastro")%></td>
</tr>
</table>
<table border=1 width=530 style='border-collapse:collapse'>
<tr>
	<td class=campo valign=top width=50><font color="green"><b>Pretensão:</b></font><br><%=rs("salario")%></td>
	<td class=campo valign=top><font color="green"><b>Habilidades:</b></font><br><%=rs("habilidade")%></td>
</tr>
</table>
<table border=1 width=530 style='border-collapse:collapse'>
<tr>
	<td class=campo valign=top width=90><font color="green"><b>Nascimento:</b></font><br><%=rs("nascimento")%> (<%=int((now()-rs("nascimento"))/365.25)%>)</td>
	<td class=campo valign=top><font color="green"><b>Endereço:</b></font><br><%=rs("bairro") & " " & rs("cidade") & " " & rs("uf")%></td>
	<td class=campo valign=top><font color="green"><b>Telefone:</b></font><br><%=rs("tel_residencial") & " " & rs("tel_celular")%></td>
</tr>
</table>
<table border=1 width=530 style='border-collapse:collapse'>
<tr>
	<td class=campo valign=top width=50><font color="green"><b>Email:</b></font><br><%=rs("email")%></td>
	<td class=campo valign=top><font color="green"><b>Observações:</b></font><br><%=rs("observacoes")%></td>
</tr>
</table>


<hr style="color:blue"> 
<%
rs.movenext
loop
rs.close
end if
%>
	
	</td>
</tr>

</table>

<%
''*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'rs.movefirst
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'*************** fim teste **********************%>

<%

set rs=nothing
conexao.close
set conexao=nothing

%>

</form>
<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="../images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>

