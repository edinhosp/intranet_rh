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
corcheck="black"
pixel=96/2.54
point=72/2.54
pointp=72.27/2.54

%>
<!-- -->
<!-- -->
<form method="POST" action="pesquisacbo.asp" name="form">
<p style="margin-top:0px;margin-bottom:2px"><b>Pesquisa</p>

<ul id="countrytabs" class="shadetabs">
<li><a href="#" rel="country1"><font style="font-size:7pt">Por Título</a></li>
<li><a href="#" rel="country2"><font style="font-size:7pt">Por Código</a></li>
<li><a href="#" rel="country3"><font style="font-size:7pt">Por Estrutura</a></li>
<li><a href="#" rel="country4"><font style="font-size:7pt">Por Título de A-Z</a></li>
</ul>

<div style="border:1px solid gray; width:690px; height:400px; margin-bottom: 1em; padding: 10px">
<!-- Por Título-->
<div id="country1" class="tabcontent">
<b><span style="font-size:10pt">Por Título</span></b>
<br><br><span style="font-size:8pt;font-weight:normal">Palavra chave:
<input type="text" name="palavra1" size="40" maxlength="40" value="<%=request.form("palavra1")%>" />
<input type="submit" value="Procurar" name="button1" />

<br><br>Procurar em: 
<br>
	<input type="checkbox" name="chkfamilia"  value="on" <%if request.form("chkfamilia")="on" or request.form="" then response.write "checked"%> />Famílias
	<input type="checkbox" name="chkocupacao" value="on" <%if request.form("chkocupacao")="on" or request.form=""  then response.write "checked"%>/>Ocupações
	<input type="checkbox" name="chksinonimo" value="on" <%if request.form("chksinonimo")="on" or request.form=""  then response.write "checked"%>/>Sinônimos
<br>
	<input type="checkbox" name="chkatividades" value="on" <%if request.form("chkatividades")="on" then response.write "checked"%> />Atividades
	<input type="checkbox" name="chkrecursos"   value="on" <%if request.form("chkrecursos")="on" then response.write "checked"%>/>Recursos de Trabalho
	<input type="checkbox" name="chkdescricao"  value="on" <%if request.form("chkdescricao")="on" then response.write "checked"%>/>Descrição
<br>
	<input type="checkbox" name="chkcondicoes" value="on" <%if request.form("chkcondicoes")="on" then response.write "checked"%>/>Condições do Trabalho
	<input type="checkbox" name="chkformacao"  value="on" <%if request.form("chkformacao")="on" then response.write "checked"%>/>Formação e Experiência
	<input type="checkbox" name="checkall" onclick="toggleAll(this)" id="Checkbox1" /><font color=green>Selecionar todos</font>
</span>

<br><br>
<%
if request.form("button1")<>"" and len(request.form("palavra1"))>0 then
	sql0=""
	if request.form("chkfamilia")="on" then sql1="select cbo=codigo_familia_cbo, nome=nome_familia, id_familia, Tipo='Família' from cbo_4familias_ocupacionais where nome_familia like '%" & request.form("palavra1") & "%' " else sql1=""
	if request.form("chkocupacao")="on" then sql2="select cbo=nu_codigo_cbo, nome=nm_ocupacao, id_familia, Tipo='Ocupação' from cbo_5ocupacoes where nm_ocupacao like '%" & request.form("palavra1") & "%' " else sql2=""
	if request.form("chksinonimo")="on" then sql3="select cbo=nu_codigo_cbo, nome=nm_titulo, id_familia, Tipo='Sinônimo' from cbo_5sinonimos s inner join cbo_5ocupacoes o on o.id_ocupacao=s.id_ocupacao where nm_titulo like '%" & request.form("palavra1") & "%' " else sql3=""
	if request.form("chkatividades")="on" then sql4="select cbo=codigo_familia_cbo, nome=nome_atividade, f.id_familia, tipo='Atividades' from cbo_9atividades a inner join cbo_9gacs g on g.id_gac=a.id_gac inner join cbo_4familias_ocupacionais f on f.id_familia=g.id_familia where nome_atividade like '%" & request.form("palavra1") & "%' union all select cbo=codigo_familia_cbo, nome=nome_gac, f.id_familia, tipo='Área de atividades' from cbo_9gacs g inner join cbo_4familias_ocupacionais f on f.id_familia=g.id_familia where nome_gac like '%" & request.form("palavra1") & "%' " else sql4=""
	if request.form("chkrecursos")="on" then sql5="select cbo=codigo_familia_cbo, nome=nm_recurso_trabalho, f.id_familia, tipo='Recurso de trabalho' from cbo_9recursos_trabalho r inner join cbo_4familias_ocupacionais f on f.id_familia=r.id_familia where nm_recurso_trabalho like '%" & request.form("palavra1") & "%' " else sql5=""
	if request.form("chkdescricao")="on" then sql6="select cbo=codigo_familia_cbo, nome=te_descricao_sumaria, id_familia, tipo='Descrição' from cbo_4familias_ocupacionais where te_descricao_sumaria like '%" & request.form("palavra1") & "%' " else sql6=""
	if request.form("chkcondicoes")="on" then sql7="select cbo=codigo_familia_cbo, nome=te_cond_geral_exerc, id_familia, tipo='Condições de trabalho' from cbo_4familias_ocupacionais where te_cond_geral_exerc like '%" & request.form("palavra1") & "%' " else sql7=""
	if request.form("chkformacao")="on" then sql8="select cbo=codigo_familia_cbo, nome=te_formacao_exper, id_familia, tipo='Formação e Experiência' from cbo_4familias_ocupacionais where te_formacao_exper like '%" & request.form("palavra1") & "%' " else sql8=""
	sql0="select cbo='0000', nome='', id_familia=0, tipo='' "
	if sql1<>"" and sql0<>"" then sql0=sql0 & " union " & sql1
	if sql2<>"" and sql0<>"" then sql0=sql0 & " union " & sql2
	if sql3<>"" and sql0<>"" then sql0=sql0 & " union " & sql3
	if sql4<>"" and sql0<>"" then sql0=sql0 & " union " & sql4
	if sql5<>"" and sql0<>"" then sql0=sql0 & " union " & sql5
	if sql6<>"" and sql0<>"" then sql0=sql0 & " union " & sql6
	if sql7<>"" and sql0<>"" then sql0=sql0 & " union " & sql7
	if sql8<>"" and sql0<>"" then sql0=sql0 & " union " & sql8
	sql1="select * from (" & sql0 & ") z where cbo<>'0000' order by nome "

	'sql0=sql1 & " union " & sql2 & " union " & sql3
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
    <table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='650px'>
	<tr>
		<td class="titulo" width='510px'>Título</td>
		<td class="titulo" width='70px'>Código</td>
		<td class="titulo" width='70px'>Tipo</td>
	</tr>
	</table>
    <div style="width:670px;overflow:auto;height:300px">
	<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='650px'>
<%	
	do while not rs.eof
	if rs("tipo")="Família" then familia=ucase(rs("nome")) else familia=rs("nome")
	if len(rs("cbo"))=6 then codigo_cbo=left(rs("cbo"),4)&"-"&right(rs("cbo"),2) else codigo_cbo=rs("cbo")
	if rs("tipo")="Sinônimo" then familia="<i>" & familia & "</i>"
	if rs("tipo")="Ocupação" then familia="<b>" & familia & "</b>"
	palavra=request.form("palavra1")
%>
	<div><a href="descricaocbo.asp?id_familia=<%=rs("id_familia")%>">
	<tr>
		<td class=campo width='510px' align="left"><%=left(replace(replace(familia,palavra,"<font color=blue><b>" & palavra & "</b></font>"),ucase(palavra),"<font color=blue><b>" & ucase(palavra) & "</b></font>"),100)%></td>
		<td class=campo width='70px' align="center"><%=codigo_cbo%></td>
		<td class=campo width='70px' align="center"><%=rs("Tipo")%></td>
	</tr>
	</a></div>
<%
	rs.movenext:loop
	rs.close
%>
        </table>
    </div>
<%
elseif request.form("button1")<>"" and len(request.form("palavra1"))=0 then
	response.write "<font color=red><b>Digite uma palavra para efetuar a pesquisa.</b></font>"
end if
%>

</div>
<!-- Por Código-->
<div id="country2" class="tabcontent">
<b><span style="font-size:10pt">Por Código</span></b>
<br><br><span style="font-size:8pt;font-weight:normal">Código: </span>
<input type="text" name="palavra2" size="6" maxlength="6" value="<%=request.form("palavra2")%>" />
<input type="submit" value="Procurar" name="button2" />
<br><br>
<%
if request.form("button2")<>"" and len(request.form("palavra2"))>=4 then
	sql1="select cbo=codigo_familia_cbo, nome=nome_familia, id_familia, Tipo='Família' from cbo_4familias_ocupacionais where codigo_familia_cbo like '" & request.form("palavra2") & "%' " & _
	"union all " & _
	"select cbo=nu_codigo_cbo, nome=nm_ocupacao, id_familia, Tipo='Ocupação' from cbo_5ocupacoes where nu_codigo_cbo like '" & request.form("palavra2") & "%' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
    <table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='650px'>
	<tr>
		<td class="titulo" width='510px'>Título</td>
		<td class="titulo" width='70px'>Código</td>
		<td class="titulo" width='70px'>Tipo</td>
	</tr>
	</table>
    <div style="width:670px;overflow:auto;height:300px">
	<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='650px'>
<%	
	do while not rs.eof
	if rs("tipo")="Família" then familia=ucase(rs("nome")) else familia=rs("nome")
	if len(rs("cbo"))=6 then codigo_cbo=left(rs("cbo"),4)&"-"&right(rs("cbo"),2) else codigo_cbo=rs("cbo")
	if rs("tipo")="Sinônimo" then familia="<i>" & familia & "</i>"
	if rs("tipo")="Ocupação" then familia="<b>" & familia & "</b>"
	palavra=request.form("palavra2")
%>
	<div><a href="descricaocbo.asp?id_familia=<%=rs("id_familia")%>">
	<tr>
		<td class=campo width='510px' align="left"><%=left(replace(replace(familia,palavra,"<font color=blue><b>" & palavra & "</b></font>"),ucase(palavra),"<font color=blue><b>" & ucase(palavra) & "</b></font>"),80)%></td>
		<td class=campo width='70px' align="center"><%=codigo_cbo%></td>
		<td class=campo width='70px' align="center"><%=rs("Tipo")%></td>
	</tr>
	</a></div>
<%
	rs.movenext:loop
	rs.close
%>
        </table>
    </div>
<%
elseif request.form("button2")<>"" and len(request.form("palavra2"))<4 then
	response.write "<font color=red><b>O código CBO tem no mínimo 4 dígitos. O código " & request.form("palavra2") & " não é válido.</b></font>"
end if
%>
</div>

<!-- Por Estrutura-->
<div id="country3" class="tabcontent">
<b><span style="font-size:10pt">Por Estrutura</span></b>
<br><br>
<span style="font-size:8pt">Grande Grupo:<br> 
	<select style="font-size:7pt"  name="grande_grupo" onchange="javascript:submit();">
	<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
	sql1="select codigo_gg, nome_gg, id_gg from cbo_1grandes_grupos order by codigo_gg, nome_gg"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	do while not rs.eof
%>
	<option <%if cstr(request.form("grande_grupo"))=cstr(rs("id_gg")) then response.write "selected"%> value="<%=rs("id_gg")%>"><%=rs("codigo_gg") & " - " & rs("nome_gg")%></option>
<%
	rs.movenext:loop
	rs.close
%>
	</select>
</span>
<br>
<span style="font-size:8pt">Subgrupo Principal: <br>
	<select style="font-size:7pt" name="subgrupo_principal" onchange="javascript:submit();">
	<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
if request.form("grande_grupo")<>"" then 
	sql1="select codigo_sgp, nome_sgp, id_sgp from cbo_2subgrupos_principais where id_gg=" & request.form("grande_grupo") & " order by codigo_sgp, nome_sgp"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	do while not rs.eof
%>
	<option <%if cstr(request.form("subgrupo_principal"))=cstr(rs("id_sgp")) then response.write "selected"%> value="<%=rs("id_sgp")%>"><%=rs("codigo_sgp") & " - " & rs("nome_sgp")%></option>
<%
	rs.movenext:loop
	rs.close
end if
%>
	</select>
</span>
<br>
<span style="font-size:8pt">Subgrupo: <br>
	<select style="font-size:7pt" name="subgrupo" onchange="javascript:submit();">
	<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
if request.form("subgrupo_principal")<>"" then 
	sql1="select codigo_sg, nome_sg, id_sg from cbo_3subgrupos where id_sgp=" & request.form("subgrupo_principal") & " order by codigo_sg, nome_sg"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	do while not rs.eof
%>
	<option <%if cstr(request.form("subgrupo"))=cstr(rs("id_sg")) then response.write "selected"%> value="<%=rs("id_sg")%>"><%=rs("codigo_sg") & " - " & rs("nome_sg")%></option>
<%
	rs.movenext:loop
	rs.close
end if
%>
	</select>
</span>
<br><br>
<%
if request.form("subgrupo")<>"" then
	sql1="select cbo=codigo_familia_cbo, nome=nome_familia, id_familia, Tipo='Família' from cbo_4familias_ocupacionais where id_sg=" & request.form("subgrupo") & " order by nome_familia"
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
    <table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='650px'>
	<tr>
		<td class="titulo" width='510px'>Título</td>
		<td class="titulo" width='70px'>Código</td>
		<td class="titulo" width='70px'>Tipo</td>
	</tr>
	</table>
    <div style="width:670px;overflow:auto;height:300px">
	<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='650px'>
<%	
	do while not rs.eof
	if rs("tipo")="Família" then familia=ucase(rs("nome")) else familia=rs("nome")
	if len(rs("cbo"))=6 then codigo_cbo=left(rs("cbo"),4)&"-"&right(rs("cbo"),2) else codigo_cbo=rs("cbo")
	if rs("tipo")="Sinônimo" then familia="<i>" & familia & "</i>"
	if rs("tipo")="Ocupação" then familia="<b>" & familia & "</b>"
%>
	<div><a href="descricaocbo.asp?id_familia=<%=rs("id_familia")%>">
	<tr>
		<td class=campo width='510px' align="left"><%=left(replace(replace(familia,palavra,"<font color=blue><b>" & palavra & "</b></font>"),ucase(palavra),"<font color=blue><b>" & ucase(palavra) & "</b></font>"),80)%></td>
		<td class=campo width='70px' align="center"><%=codigo_cbo%></td>
		<td class=campo width='70px' align="center"><%=rs("Tipo")%></td>
	</tr>
	</a></div>
<%
	rs.movenext:loop
	rs.close
%>
        </table>
    </div>
<%
end if
%>


</div>
<!-- Por Titulo de A-Z-->
<div id="country4" class="tabcontent">
<table border="0" cellpadding="2" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='<%=13.0*pixel%>px' >
</table>
<b><span style="font-size:10pt;">Por Títulos de A-Z</span></b>
<br><br><span style="font-size:9pt;background:silver;width:490;text-align:center">
<%
for a=65 to 90
	idbutton=a-64
	if request.form("button"&a)=chr(a) then letra=chr(a) else letra=letra
	if request.form("button"&a)=chr(a) then marca="bold" else marca="none"
%>
	<input type="submit" value="<%=chr(a)%>" name="button<%=a%>" style="border:1px solid gray;background:silver;color:black;font-weight:<%=marca%>" />
<%
next
%>
</span>
<br><br>
<%
if letra<>"" then
	sql1="select * from ( " & _
	"select cbo=codigo_familia_cbo, nome=nome_familia, id_familia, Tipo='Família' from cbo_4familias_ocupacionais where nome_familia like '" & letra & "%' " & _
	"union all " & _
	"select cbo=nu_codigo_cbo, nome=nm_ocupacao, id_familia, Tipo='Ocupação' from cbo_5ocupacoes where nm_ocupacao like '" & letra & "%' " & _
	"union all " & _
	"select cbo=nu_codigo_cbo, nome=nm_titulo, id_familia, Tipo='Sinônimo' from cbo_5sinonimos s inner join cbo_5ocupacoes o on o.id_ocupacao=s.id_ocupacao where nm_titulo like '" & letra & "%' " & _
	") z order by nome "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
    <table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='650px'>
	<tr>
		<td class="titulo" width='510px'>Título</td>
		<td class="titulo" width='70px'>Código</td>
		<td class="titulo" width='70px'>Tipo</td>
	</tr>
	</table>
    <div style="width:670px;overflow:auto;height:300px">
	<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" width='650px'>
<%	
	do while not rs.eof
	if rs("tipo")="Família" then familia=ucase(rs("nome")) else familia=rs("nome")
	if len(rs("cbo"))=6 then codigo_cbo=left(rs("cbo"),4)&"-"&right(rs("cbo"),2) else codigo_cbo=rs("cbo")
	if rs("tipo")="Sinônimo" then familia="<i>" & familia & "</i>"
	if rs("tipo")="Ocupação" then familia="<b>" & familia & "</b>"
%>
	<div><a href="descricaocbo.asp?id_familia=<%=rs("id_familia")%>">
	<tr>
		<td class=campo width='510px' align="left"><%=left(replace(replace(familia,palavra,"<font color=blue><b>" & palavra & "</b></font>"),ucase(palavra),"<font color=blue><b>" & ucase(palavra) & "</b></font>"),80)%></td>
		<td class=campo width='70px' align="center"><%=codigo_cbo%></td>
		<td class=campo width='70px' align="center"><%=rs("Tipo")%></td>
	</tr>
	</a></div>
<%
	rs.movenext:loop
	rs.close
%>
	</table>
    </div>

<%
end if
%>


</div>
<!-- -->

</div>

<script type="text/javascript">
var countries=new ddtabcontent("countrytabs")
countries.setpersist(true)
countries.setselectedClassTarget("link") //"link" or "linkparent"
countries.init()
</script>
<!--
<p><a href="javascript:countries.cycleit('prev')" style="margin-right: 25px; text-decoration:none">volta</a> <a href="javascript: countries.expandit(3)">Click here to select last tab</a> <a href="javascript:countries.cycleit('next')" style="margin-left: 25px; text-decoration:none">avança</a></p>
-->



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
</body>
</html>