<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>BO 60/05</TITLE>
<link rel="stylesheet" type="text/css" href="diversos.css">
</HEAD>
<BODY>
<%
dim conexao,rs,marc(6), formato(6)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
set rs4=server.createobject ("ADODB.Recordset")
Set rs4.ActiveConnection = conexao
%>
<table border='1' bordercolor='#000000' cellpadding='4' cellspacing='0' style='border-collapse:collapse'>
<%
sql="SELECT A.CHAPA FROM ABATFUN A, " & _
"(SELECT S.CHAPA, DTMUDANCA, CODSECAO FROM PFHSTSEC S, " & _
"(SELECT CHAPA, Max(DTMUDANCA) AS MUDANCA FROM PFHSTSEC GROUP BY CHAPA HAVING Max(DTMUDANCA)<=#1/17/2005# ORDER BY Max(DTMUDANCA) DESC) AS D " & _
"WHERE S.CHAPA=D.CHAPA AND DTMUDANCA=MUDANCA) S1 " & _
"WHERE S1.CHAPA=A.CHAPA AND LEFT(CODSECAO,2)='04' AND A.DATA=#01/17/05# " & _
"GROUP BY A.CHAPA "
rs3.Open sql, ,adOpenStatic, adLockReadOnly
rs3.movefirst
do while not rs3.eof
sqla="select f.nome, f.codsituacao, f.chapa, f.admissao, c.nome as funcao, f.codsecao, s.descricao as secao, f.estadocivil, " & _
"f.grauinstrucao, f.rua, f.numero, f.complemento, f.bairro, f.cidade, f.cep, f.telefone1, f.telefone2, f.telefone3, " & _
"f.fax, f.sexo, f.dtnascimento, f.email, f.cpf, f.cartidentidade, f.ufcartident, f.pispasep, f.codpessoa, f.demissao, " & _
"ss.descricao as sit, i.descricao as titulacao, f.dtemissaoident, f.mae " & _
"from qry_funcionarios f, pfuncao c, psecao s, pcodsituacao ss, pcodinstrucao i " & _
"where f.codfuncao = c.codigo and f.codsecao = s.codigo " & _
"and f.codsituacao = ss.codcliente and f.grauinstrucao=i.codcliente "
sqlb="and f.CHAPA='" & rs3("chapa") & "' "
sqlc="ORDER BY f.CHAPA"
sql1=sqla & sqlb & sqlc
rs.Open sql1, ,adOpenStatic, adLockReadOnly

'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
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
'rs.movefirst
'*************** fim teste **********************
%>
<tr>
	<td class=titulo style="border-style:solid;border-top-width:2" >Nome</td>
	<td class=campo colspan=3 style="border-style:solid;border-top-width:2"><b><%=rs("nome")%></td>
	<td class=titulo rowspan=8 style="border-style:solid;border-top-width:2"><img border="0" src="func_foto.asp?chapa=<%=rs("chapa")%>"  width="200"></td>
</tr>
<tr>
	<td class=titulo>Admissão</td>
	<td class=campo><%=rs("admissao")%></td>
	<td class=titulo>Demissão</td>
	<td class=campo><%=rs("demissao")%></td>
</tr>
<tr>
	<td class=titulo>Função</td>
	<td class=campo colspan=3><%=rs("funcao")%></td>
</tr>
<tr>
	<td class=titulo>Endereço</td>
	<td class=campo colspan=3><%=rs("rua") & " " & rs("numero") & " " & rs("complemento") & " - " & rs("bairro") & " - " & rs("cidade") & " - " & rs("CEP")%></td>
</tr>
<tr>
	<td class=titulo>Telefone</td>
	<td class=campo colspan=3><%=rs("telefone1") & " " & rs("telefone2")%></td>
</tr>
<tr>
	<td class=titulo>Dt.Nasc.</td>
	<td class=campo colspan=3><%=rs("dtnascimento")%></td>
</tr>
<tr>
	<td class=titulo>Cart.Identidade</td>
	<td class=campo colspan=3><%=rs("cartidentidade") & " / " & rs("ufcartident")%></td>
</tr>
<tr>
	<td class=titulo>Mãe</td>
	<td class=campo colspan=3><%=rs("mae")%></td>
</tr>
<tr>
	<td class=titulo>Marcações do ponto</t>
	<td class=campo colspan=4>
<%
sql="select batida from abatfun where chapa='" & rs("chapa") & "' and data=#1/17/05# order by batida"
rs2.Open sql, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
	valor=rs2("batida")
	hora=int(valor/60)
	minuto=valor-(hora*60)
	valord=hora & ":" & numzero(minuto,2)

response.write valord
response.write "-"
rs2.movenext
loop
rs2.close
%>
</td>
</tr>
<%
rs.close
if (rs3.absoluteposition/3)-int(rs3.absoluteposition/3)=0 then 
response.write "</table>"
response.write "<DIV style=""page-break-after:always""></DIV>"
response.write "<table border='1' bordercolor='#000000' cellpadding='4' cellspacing='0' style='border-collapse:collapse'>"
end if
rs3.movenext
loop
%>
<%
rs3.close
set rs3=nothing
conexao.close
set conexao=nothing
%>
</table>
</BODY>
</HTML>