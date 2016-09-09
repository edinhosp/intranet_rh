<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a100")="N" or session("a100")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Pesquisa CBO</title>
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
<link rel="stylesheet" type="text/css" href="tabcontent.css" />
<script type="text/javascript" src="tabcontent.js">
/***********************************************
* Tab Content script v2.2- © Dynamic Drive DHTML code library (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit Dynamic Drive at http://www.dynamicdrive.com/ for full source code
***********************************************/
</script>
</head>
<body>
<%
dim conexao, rs, rs2, formato(2)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao
corcheck="black"
pixel=96/2.54
point=72/2.54
pointp=72.27/2.54
id_familia=request("id_familia")

%>
<form method="POST" action="descricaocbo.asp" name="form">
<p style="margin-top:0px;margin-bottom:5px"><b>Classificação Brasileira de Ocupações<br>Relatório da Família
<%for a=1 to 120%>&nbsp;<%next%>
<a href="comparacbo.asp?id_familia=<%=id_familia%>">
<img src="../images/tjunta.jpg" width="16" height="14" border="0" alt="Comparar as ocupações"></a>
</p>

<br>
<table border="0" cellpadding="0" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='650px'>
<tr>
	<td class=titulop>Código</td>
	<td class=titulop>Títulos</td>
</tr>
<%
sql1="select codigo_familia_cbo, nome_familia, te_descricao_sumaria, te_cond_geral_exerc, te_formacao_exper, te_glossario, te_notas, id_conveniada from cbo_4familias_ocupacionais where id_familia=" & id_familia
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<tr>
	<td class="campop"><%=rs("codigo_familia_cbo")%></td>
	<td class="campop"><%=rs("nome_familia")%></td>
</tr>
</table>
<%
%>

<br>
<table border="0" cellpadding="0" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='650px'>
<tr>
	<td class=titulop colspan=2 >Ocupações / Sinônimos</td>
</tr>
<%
sql1="select id_ocupacao, nu_codigo_cbo, nm_ocupacao from cbo_5ocupacoes where id_familia=" & id_familia
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
cbo=left(rs2("nu_codigo_cbo"),4) & "-" & right(rs2("nu_codigo_cbo"),2)
%>
<tr>
	<td class="campop" valign=top colspan=2><%=cbo & " - <b>" & rs2("nm_ocupacao")%></td>
</tr>
<tr>
	<td class=campo valign=top style="border-bottom:1px solid #000000" width=60></td>
	<td class=campo valign=top style="border-bottom:1px solid #000000">
<%
sql2="select nm_titulo from cbo_5sinonimos where id_ocupacao=" & rs2("id_ocupacao")
rs3.Open sql2, ,adOpenStatic, adLockReadOnly
do while not rs3.eof
	response.write rs3("nm_titulo")
	if rs3.recordcount>1 and rs3.absoluteposition<rs3.recordcount then response.write ", "
rs3.movenext:loop
rs3.close
%>
	</td>
</tr>
<%
rs2.movenext:loop
rs2.close
%>
</table>


<table border="0" cellpadding="1" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='650px'>
<tr>
	<td class=titulop>Descrição Sumária</td>
</tr>
<tr>
	<td class="campop"><%=rs("te_descricao_sumaria")%></td>
</tr>
<tr>
	<td class=titulop>Formação e experiência</td>
</tr>
<tr>
	<td class="campop"><%=rs("te_formacao_exper")%></td>
</tr>
<tr>
	<td class=titulop>Condições gerais de exercício</td>
</tr>
<tr>
	<td class="campop"><%=rs("te_cond_geral_exerc")%></td>
</tr>
<tr>
	<td class=titulop>Esta família não compreende</td>
</tr>
<tr>
	<td class="campop">
<%
sql4="select id_familia_referenciada, codigo_familia_cbo, nome_familia from cbo_4referencias r inner join cbo_4familias_ocupacionais f on f.id_familia=r.id_familia_referenciada where r.id_familia=" & id_familia & " and tp_referencia in (2,1) "
rs2.Open sql4, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
do while not rs2.eof
	response.write rs2("codigo_familia_cbo") & " - " & rs2("nome_familia")
	if rs2.recordcount>1 and rs2.absoluteposition<rs2.recordcount then response.write "<br>"
rs2.movenext:loop
end if
rs2.close
%>&nbsp;
	</td>
</tr>
<tr>
	<td class=titulop>Notas</td>
</tr>
<tr>
	<td class="campop"><%=rs("te_notas")%>&nbsp;</td>
</tr>
<tr>
	<td class=titulop>GACS - Grande Área de Competência</td>
</tr>
<%
sql5="select id_gac, nome_ordem, nome_gac, id_tipo_gac from cbo_9gacs where id_familia=" & id_familia & " order by id_tipo_gac, nome_ordem "
rs2.Open sql5, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
do while not rs2.eof
%>
<tr><td class="campop" height=25 valign=middle><%=rs2("nome_ordem") & " - " & rs2("nome_gac")%></b></td></tr>
<%	
	'------------------
	sql6="select nu_ordem, nome_atividade from cbo_9atividades where id_gac=" & rs2("id_gac") & " order by nu_ordem"
	rs3.Open sql6, ,adOpenStatic, adLockReadOnly
	linha=0
	do while not rs3.eof
	if linha=1 then 
		fundo="style='background:#ffffff;'"
		linha=0
	else
		fundo="style='background:#ffffcc;'"
		linha=1
	end if
%>
<tr><td class="campop" <%=fundo%> valign=middle>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rs2("nome_ordem") & "." & rs3("nu_ordem") & " - " & rs3("nome_atividade")%></td></tr>
<%	
	rs3.movenext:loop
	rs3.close	
	'------------------
rs2.movenext:loop
end if
rs2.close
%>
<tr>
	<td class=titulop>Recursos de trabalho</td>
</tr>
<tr>
	<td class="campop">
<%
sql7="select nm_recurso_trabalho, in_publicacao from cbo_9recursos_trabalho where id_familia=" & id_familia
rs2.Open sql7, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
linha=0
do while not rs2.eof
	if linha=1 then 
		fundo="style='background:#ffffff;'"
		linha=0
	else
		fundo="style='background:#ffffcc;'"
		linha=1
	end if
if rs2("in_publicacao")=1 or rs2("in_publicacao")=true then texto1=" * " else texto1="&nbsp;&nbsp;&nbsp;"
%>
<tr><td class="campop" <%=fundo%> valign=middle><%=texto1%><%=rs2("nm_recurso_trabalho")%></td></tr>
<%	
rs2.movenext:loop
end if
rs2.close
%>&nbsp;
	</td>
</tr>
<tr>
	<td class="campor" align="right"><b>*</b> Recursos de trabalho mais importantes</td>
</tr>
<tr>
	<td class=titulop>Instituição Conveniada Responsável</td>
</tr>
<tr>
	<td class="campop">
<%
sql7="select nome_conveniada from cbo_tp4_conveniadas where id_conveniada=" & rs("id_conveniada")
rs2.Open sql7, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
%>
<tr><td class="campop" valign=middle>&nbsp;<%=rs2("nome_conveniada")%></td></tr>
<%	
end if
rs2.close
%>&nbsp;
	</td>
</tr>
<tr>
	<td class=titulop>Glossário</td>
</tr>
<tr>
	<td class="campop"><%=rs("te_glossario")%>&nbsp;</td>
</tr>




</table>


<%
'*************** inicio teste **********************
'totaldisp=0
'response.write "<table border='1' bordercolor='#000000' cellpadding='0' cellspacing='0' style='border-collapse:collapse' width='600'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
'next
'response.write "</tr>"
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	if a>2 then alinhamento="center" else alinhamento="left"
'	response.write "<td class="campor" nowrap align='" & alinhamento & "'>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext:loop
'response.write "</table>"
'*************** fim teste **********************

'rs.close
'set rs=nothing
'conexao.close
'set conexao=nothing
%>
<!-- -->
<!-- -->
</form>
</body>
</html>