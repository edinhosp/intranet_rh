<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a96")="N" or session("a96")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Checagem de Inconsistências Folha/Cadastro</title>
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
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() {	form.chapa.value=form.nome.value;	}
function chapa1() {	form.nome.value=form.chapa.value;	}

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

</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("b1")="" then
%>
<p class=titulo>Checagem de Inconsistências - Folha/Cadastro&nbsp;<%=titulo %>
<form method="POST" action="inconsistencia.asp" name="form">

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=400>
<tr>
	<td class=titulo>Categoria</td>
	<td class=titulo><input type="checkbox" name="checkall" onclick="toggleAll(this)" id="Checkbox1" /></td>
	<td class=titulo>Descrição</td>
	<td class=titulo>Parametros utilizados</td>
</tr>
<%
vezes=0
sql0="select id_check, categoria, id_script, nome_script, parametro from folha_check_script order by categoria, id_script "
rs.Open sql0, ,adOpenStatic, adLockReadOnly

do while not rs.eof
	if (rs.absoluteposition mod 2)=0 then classe="campol" else classe="campo"
	if ultimacategoria<>rs("categoria") then
		status=1
		sql0a="select itens=count(categoria) from folha_check_script where categoria='" & rs("categoria") & "'"
		rs2.Open sql0a, ,adOpenStatic, adLockReadOnly
			linhas=rs2("itens")
		rs2.close
	else
		status=0
	end if
%>
<tr>
<%if status=1 then %>
	<td class=campo rowspan=<%=linhas%>>
	<%
	tam=len(rs("categoria"))
	for a=1 to tam
		letra=mid(rs("categoria"),a,1)
		if letra=" " then response.write "<br>"
		response.write letra
	next
	%>
	</td>
<%end if%>
	<td class=<%=classe%>><%checado1="checked":checado0=""%>
		<input type="checkbox" name="checar<%=vezes%>" value="ON" <%=checado1%> >
		<input type="hidden" name="id_check<%=vezes%>" value="<%=rs("id_check")%>" >
	</td>
	<td class=<%=classe%>><%=rs("nome_script")%></td>
	<td class=<%=classe%>><%=rs("parametro")%></td>
</tr>
<%
	vezes=vezes+1
	session("check_inconsistencias")=vezes-1
	ultimacategoria=rs("categoria")
rs.movenext:loop
rs.close
%>
</table>

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=400>
<tr><td class=titulo>Parâmetros a utilizar:</td></tr>
<%
if day(now)<10 then
	mes=month(now)-1:mes_a=mes
elseif day(now)<20 then
	mes=month(now):mes_a=month(now)-1
else
	mes=month(now):mes_a=month(now)-1
end if
%>
<tr>
	<td class=titulo>Ano: <input type="text" name="ano" value="<%=year(now)%>" size=4>
	Mês Atual: <input type="text" name="mes" value="<%=mes%>" size=2>
	Mês Anterior: <input type="text" name="mesant" value="<%=mes_a%>" size=2>
	</td>
</tr>
<tr>
	<td class=titulo colspan=3><input type="submit" value="Visualizar" class="button" name="B1"></td>
</tr>
</table>

<%
%>
</form>
<%
else ' request.form<>""
	vez=session("check_inconsistencias")
	sql0="delete from folha_check where sessao='" & session("usuariomaster") & "'"
	conexao.execute sql0
	intro1="declare @sessao nvarchar(20) set @sessao='" & session("usuariomaster") & "' "
	intro2="declare @ano integer, @mes integer, @mesant integer " & _
	"set @ano=" & request.form("ano") & " " & _
	"set @mes=" & request.form("mes") & " " & _
	"set @mesant=" & request.form("mesant") & " " 
	intro3="insert into folha_check (chapa, nome, descricao, valor_atual, valor_previsto, sessao) "

	for a=0 to vez
		checar=request.form("checar" & a)
		id_check=request.form("id_check" & a)
		if checar="ON" then
			sql1="select texto_script from folha_check_script where id_check=" & id_check
			rs.Open sql1, ,adOpenStatic, adLockReadOnly
			texto=rs("texto_script")
			rs.close
			sql2=intro1 & intro2 & intro3 & texto
			'response.write "<br>" & sql2
			conexao.execute sql2
		end if
	next

sql3="select chapa, nome, descricao, valor_atual, valor_previsto " & _
"from folha_check where sessao='" & session("usuariomaster") & "'"
rs.Open sql3, ,adOpenStatic, adLockReadOnly
total=rs.recordcount
if rs.recordcount>0 then
rs.movefirst
inicio=1:repete=0:linha=0
%>
<%
do while not rs.eof
if linha>=50 then repete=1 else repete=0
if inicio=1 or repete=1 then 
	if repete=1 then response.write "</table><DIV style=""page-break-after:always""></DIV>"
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="660">
<tr>
	<td valign="middle" height=55><img border="0" src="../images/teste_logo2.png" height=55 ></td>
	<td valign="middle" align="center" class="campop"><b>Checagem de inconsistências no cadastro</b><br><%=now()%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="660" >
<tr>
	<td class=titulo align="center">Chapa</td>
	<td class=titulo align="center">Nome</td>
	<td class=titulo align="center">Inconsistência</td>
	<td class=titulo align="center">Valor atual</td>
	<td class=titulo align="center">Obs./correção</td>
</tr>
<%
	linha=0
end if
%>
<tr>
	<td class=campo align="center"><%=rs("chapa")%></td>
	<td class=campo><%=rs("nome")%></td>
	<td class=campo><%=rs("descricao")%></td>
	<td class=campo><%=rs("valor_atual")%></td>
	<td class=campo><%=rs("valor_previsto")%></td>
</tr>
<%
linha=linha+1
rs.movenext
inicio=0
loop
end if 'recordcount>0
rs.close
%>	
</table>
	
	


<%
if giro<vezes then response.write "<DIV style=""page-break-after:always""></DIV>"
%>

<%
end if 
%>
</body>
</html>
<%
set rs=nothing
conexao.close
set conexao=nothing
%>