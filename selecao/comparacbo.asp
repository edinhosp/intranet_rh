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

legenda=""
sql2="select id_ocupacao, id_familia, sg_ocupacao, nm_ocupacao, nu_codigo_cbo,id_ocupacao_gh from cbo_5ocupacoes where id_familia=" & id_familia
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
	legenda=legenda & "<a class=r href='comparacbo.asp?id_familia=" & id_familia & "&sigla=" & rs2("sg_ocupacao") & "'>" & rs2("sg_ocupacao") & "</a> - " & rs2("nm_ocupacao") & "<br>"
rs2.movenext:loop
rs2.close

%>
<form method="POST" action="comparacbo.asp" name="form">
<table border="0" cellpadding="0" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='690px'>
<tr>
	<td class="campop" align="left" valign=top>Classificação Brasileira de Ocupações<br>Atividades - Comparação</td>
	<td class=campo align="right" valign=top rowspan=2><%=legenda%></td>
</tr>
<%
sql1="select codigo_familia_cbo, nome_familia, te_descricao_sumaria, te_cond_geral_exerc, te_formacao_exper, te_glossario, te_notas, id_conveniada from cbo_4familias_ocupacionais where id_familia=" & id_familia
rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<tr>
	<td class="campop"><%=rs("codigo_familia_cbo")%> <b><%=rs("nome_familia")%></b></td>
</tr>
</table>

<%
rs.close


'---------------------------------------------------------
sql3="select g.id_gac, t.atividades, g.nome_ordem, g.nome_gac, a.id_atividade, a.nu_ordem, a.nome_atividade "
'---------------------------------------------------------
sql2="select id_ocupacao, id_familia, sg_ocupacao, nm_ocupacao, nu_codigo_cbo,id_ocupacao_gh from cbo_5ocupacoes where id_familia=" & id_familia
if request("sigla")<>"" then sql2=sql2 & " and sg_ocupacao='" & request("sigla") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
	if rs2("id_ocupacao_gh")>0 then gh=" or p.id_ocupacao=" & rs2("id_ocupacao_gh") & " " else gh=""
	sql3=sql3 & ", '" & rs2("sg_ocupacao") & "'=max(case when p.id_ocupacao=" & rs2("id_ocupacao") & gh &  " then 'X' else null end) "
rs2.movenext:loop
rs2.close
'---------------------------------------------------------
sql3=sql3 & " from cbo_9gacs g " & _
"inner join cbo_9atividades a on a.id_gac=g.id_gac " & _
"inner join cbo_5ocupacoes o on o.id_familia=g.id_familia " & _
"inner join ( " & _
"select id_gac, atividades=count(id_atividade) from cbo_9atividades group by id_gac " & _
") t on t.id_gac=g.id_gac " & _
"left join cbo_9perfis_ocupacionais p on p.id_atividade=a.id_atividade and p.id_ocupacao=o.id_ocupacao " & _
"where g.id_familia=" & id_familia & " " & _
"group by g.id_gac, t.atividades, g.nome_ordem, g.nome_gac, a.id_atividade, a.nu_ordem, a.nome_atividade " & _
"order by nome_ordem, nu_ordem "
'---------------------------------------------------------
rs.Open sql3, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="1" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='690px'>
<tr>
	<td class=titulop colspan=<%=rs.fields.count-7+1%>>Áreas / Atividades</td>
</tr>
<%
do while not rs.eof
	if rs("nome_ordem")<>ultimogac then mudou=1 else mudou=0

if mudou=1 then
%>
	<tr>
	<td class=grupo valign=top>&nbsp;<%=rs("nome_ordem") & " - " & rs("nome_gac")%></td>
	<%for a=7 to rs.fields.count-1%>
	<td class=grupo align="center"><%=rs.fields(a).name%></td>
	<%next%>
	</tr>
<%
end if
%>
<tr>
	<td class=campo><%=rs("nu_ordem") & ". " & rs("nome_atividade")%></td>
	<%for a=7 to rs.fields.count-1%>
	<td class=campo align="center"><%=rs.fields(a)%></td>
	<%next%>
</tr>	

<%
ultimogac=rs("nome_ordem")
rs.movenext
loop
%>
</table>

<%
'*************** inicio teste **********************
'rs.movefirst
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

rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
<!-- -->
<!-- -->
</form>
</body>
</html>