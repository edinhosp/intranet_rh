<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a93")="N" or session("a93")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Plano de Ensino</title>
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
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if session("usuariogrupo")="RH" or session("usuariogrupo")="PLANEJAMENTO" or session("usuariogrupo")="SECR.GERAL" or session("usuariogrupo")="DTI" then podeduplicar=1 else podeduplicar=0

if request("origem")<>"" and request("destino")<>"" then
	sqlc="update grades_plano set justificativa=o.justificativa, ementa=o.ementa, objetivos_gerais=o.objetivos_gerais, unidades_tematicas=o.unidades_tematicas, " & _
	"metodologia=o.metodologia, avaliacao=o.avaliacao, bibliografia=o.bibliografia, bibliografiac=o.bibliografiac, novo=0, pa=0, " & _
	"usuarioc='" & session("usuariomaster") & "', datac=getdate() " & _
	"from grades_plano d, (Select * from grades_plano where id_plano=" & request("origem") & ") o where d.id_plano=" & request("destino") & " "
	conexao.execute sqlc	
	sqlb="insert into grades_plano_biblio (id_plano, complementar, cod_acervo, ordem, referencia, status) " & _
	"select " & request("destino") & ", complementar, cod_acervo, ordem, referencia, status " & _
	"from grades_plano_biblio where id_plano=" & request("origem") & ""
	conexao.execute sqlb
	response.write "<br>" & request("origem") & "->" & request("destino")
end if

if request("origemb")<>"" and request("destinob")<>"" then
	sqlb="insert into grades_plano_biblio (id_plano, complementar, cod_acervo, ordem, referencia, status) " & _
	"select " & request("destinob") & ", complementar, cod_acervo, ordem, referencia, status " & _
	"from grades_plano_biblio where id_plano=" & request("origemb") & ""
	conexao.execute sqlb
	response.write "<br>" & request("origemb") & "->" & request("destinob")
end if


coddoc="":gc="":perlet=""
'session("peperlet")="":session("pegc")="":session("pecoddoc")=""
if request.form("coddoc")<>"" then 
	coddoc=request.form("coddoc"):session("pecoddoc")=request.form("coddoc") 
else 
	coddoc=session("pecoddoc")
end if
if request.form("gc")<>"" then 
	gc=request.form("gc"):session("pegc")=request.form("gc")
else 
	gc=session("pegc")
end if
if request.form("perlet")<>"" then 
	perlet=request.form("perlet"):session("peperlet")=request.form("perlet") 
else 
	perlet=session("peperlet")
end if

'response.write "<br>" & session("pecoddoc") & "-" & request.form("coddoc")
'response.write "<br>" & session("pegc") & "-" & request.form("gc")
'response.write "<br>" & session("peperlet") & "-" & request.form("perlet")
%>
<p style="margin-top:0;margin-bottom:0" class=titulo>Plano de Ensino para&nbsp;<%=nomeacao%></p>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse:collapse;" >
<form method="POST" action="planoensino2.asp" name="form">
<tr>
	<td class=titulo height=30 valign='middle'>Curso:</td>
	<td class=titulo><select size="1" name="coddoc" onChange="javascript:submit()">
	<option value="" selected>Selecione um curso</option>
<%
sqla="select distinct tpcurso, coddoc, curso, descricao=case tpcurso when 'G' then 'Graduação' when 'L' then 'Cursos Livres' when 'M' then 'Mestrado' when 'P' then 'Pós-Graduação' else '' end " & _
"from grades_pe order by tpcurso, curso"

rs.Open sqla, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
if rs("tpcurso")<>grupoanterior then response.write "<option style='background:CCFFCC' value='" & rs("tpcurso") & "'>------- " & ucase(rs("descricao")) & " --------</option>"
%>
<option <%if coddoc=rs("coddoc") then response.write "selected "%> value="<%=rs("coddoc")%>"><%=rs("curso")%></option>
<%
grupoanterior=rs("tpcurso")
rs.movenext:loop
rs.close
%>  
	</select>
	</td>

	<td class=titulo>Grade</td>
	<td class=titulo><select size="1" name="gc" onChange="javascript:submit()">
	<option value="" selected>Selecione uma grade</option>
<%
if session("usuariogrupo")="PLANEJAMENTO" then filtrocurso=" and codcur>=500 " else filtrocurso=""
sqla="select codpergrade, descricao, codcur, codper, grade from grades_pe where coddoc='" & coddoc & "' " & filtrocurso & " order by codcur, codper, grade "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof 
%>
	<option <%if gc=rs("codpergrade") then response.write "selected "%> value="<%=rs("codpergrade")%>"><%=rs("descricao")%></option>
<%
rs.movenext:loop
end if
rs.close
%>  
	</select>
	</td>
<input type="hidden" name="acoddoc" value="<%=request.form("coddoc")%>">
<input type="hidden" name="agc" value="<%=request.form("gc")%>">
<%
vezes=1
'if request.form("coddoc")="" then coddoc=session("pecoddoc") else coddoc=request.form("coddoc")
'if request.form("gc")="" then gc=session("pegc") else gc=request.form("gc")
'if request.form("perlet")="" then perletsession("peperlet") else perlet=request.form("perlet")

acoddoc=request.form("acoddoc")
if acoddoc<>coddoc and acoddoc<>"" then gc=""
sql="select codcur, codper, grade from grades_pe where codpergrade='" & gc & "'"
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then codcur=rs("codcur"):codper=rs("codper"):grade=rs("grade")
rs.close
%>
<input type="hidden" name="hcodcur" value="<%=codcur%>">
<input type="hidden" name="hcodper" value="<%=codper%>">
<input type="hidden" name="hgrade" value="<%=grade%>">


	<td class=titulo>Per.Letivo</td>
	<td class=titulo><select size="1" name="perlet" onChange="javascript:submit()">
	<option value="" selected>Selecione um período</option>
<%
sqla="select distinct perlet from g2turmas where coddoc='" & coddoc & "' and codcur=" & codcur & " and codper=" & codper & " and grade=" & grade
rs.Open sqla, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof 
%>
<option <%if perlet=rs("perlet") then response.write "selected "%> value="<%=rs("perlet")%>"><%=rs("perlet")%></option>
<%
rs.movenext:loop
end if
rs.close
%>  
	</select>
	</td>
</tr>
</form>
</table>
<br>
<%
if (request.form("coddoc")<>"" or session("pecoddoc")<>"") and (request.form("gc")<>"" or session("pegc")<>"") and _
	(request.form("perlet")<>"" or session("peperlet")<>"")  then
%>
<table border="1" bordercolor="#000000" cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=grupo colspan=5> Curso: <%=coddoc%> - Grade Curricular: <%=grade%> - Período Letivo: <%=perlet%></td>
</tr>
<tr>
	<td class=titulo>Período</td>
	<td class=titulo>Disciplina</td>
	<td class=titulo>CH Sem.</td>
	<td class=titulo>CH Total</td>
	<td class=titulo>&nbsp;</td>
</tr>
<%
'*****************************************
' checa e insere series e disciplinas do plano
'*****************************************

sqli="insert into grades_plano (coddoc, codcur, codper, grade, perlet, serie, codmat, pa, novo, prof, coord) " & _
"select distinct g.coddoc, g.codcur, g.codper, g.grade, g.perlet, g.serie, g.codmat, 0, 1,0,0 " & _
"from g2ch g left join grades_plano u on u.codcur=g.codcur and u.codper=g.codper and u.grade=g.grade and u.serie=g.serie and u.perlet=g.perlet " & _
"where g.coddoc='" & coddoc & "' and g.codcur=" & codcur & " and g.codper=" & codper & " and g.grade=" & grade & " and g.perlet='" & perlet & "' and u.codmat is null " & _
"group by g.coddoc, g.codcur, g.codper, g.grade, g.perlet, g.serie, g.codmat "
conexao.Execute sqli, , adCmdText
sqli="insert into grades_plano (coddoc, codcur, codper, grade, perlet, serie, codmat, pa, novo, prof, coord) " & _
"select distinct g.coddoc, g.codcur, g.codper, g.grade, g.perlet, g.serie, g.codmat, 0, 1,0,0 " & _
"from g5ch g left join grades_plano u on u.codcur=g.codcur and u.codper=g.codper and u.grade=g.grade and u.serie=g.serie and u.perlet=g.perlet " & _
"where g.coddoc='" & coddoc & "' and g.codcur=" & codcur & " and g.codper=" & codper & " and g.grade=" & grade & " and g.perlet='" & perlet & "' and u.codmat is null " & _
"group by g.coddoc, g.codcur, g.codper, g.grade, g.perlet, g.serie, g.codmat "
conexao.Execute sqli, , adCmdText

sqlj="insert into grades_plano (coddoc, codcur, codper, grade, perlet, serie, codmat, pa, novo, prof, coord) " & _
"select '" & coddoc & "', u.codcur, u.codper, u.grade, '" & perlet & "', u.periodo, u.codmat, 0, 1,0,0 " & _
"from corporerm.dbo.ugrade u left join (select * from grades_plano where coddoc='" & coddoc & "' and perlet='" & perlet & "') p " & _
"	on p.codcur=u.codcur and p.codper=u.codper and p.grade=u.grade and p.serie=u.periodo and p.codmat=u.codmat collate database_default " & _
"where u.codcur=" & codcur & " and u.codper=" & codper & " and u.grade=" & grade & " and p.codmat is null " & _
"and periodo between (select min(serie) from grades_plano where coddoc='" & coddoc & "' and codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and perlet='" & perlet & "') " & _
"	and (select max(serie) from grades_plano where coddoc='" & coddoc & "' and codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and perlet='" & perlet & "') "
conexao.Execute sqlj, , adCmdText

'*****************************************
' busca series e disciplinas do plano
'*****************************************
sql="select p.id_plano, p.coddoc, p.codcur, p.codper, p.grade, p.perlet, p.serie, p.codmat, m.materia, g.naulassem, g.cargahoraria, p.pa, p.novo, p.justificativa " & _
", nb=(select count(id_plano) from grades_plano_biblio where id_plano=p.id_plano) " & _
"from grades_plano p inner join corporerm.dbo.umaterias m on m.codmat collate database_default=p.codmat " & _
"inner join corporerm.dbo.ugrade g on g.codmat=m.codmat and g.codcur=p.codcur and g.codper=p.codper and g.grade=p.grade " & _
"where p.coddoc='" & coddoc & "' and p.codcur=" & codcur & " and p.codper=" & codper & " and p.grade=" & grade & " and p.perlet='" & perlet & "' " & _
"order by p.serie, materia"
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof
barra1=""
pa=rs("pa"):novo=rs("novo")
%>
<tr>
<%
if lastper<>rs("serie") then
	sqlg="select count(serie) as linhas from grades_plano where coddoc='" & coddoc & "' and codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and perlet='" & perlet & "' and serie=" & rs("serie")
	rs2.Open sqlg, ,adOpenStatic, adLockReadOnly:lin_per=rs2("linhas")
	rs2.close
	barra1=" style='border-top:2 solid #000000'"
	response.write "<td class=""campop"" rowspan=" & lin_per & " align=""center"" style='border-bottom:2 solid #000000;border-top:2 solid #000000'><b>" & rs("serie") & "</td>"
end if
%>
	<td class=campo <%=barra1%> >
<%if novo=false or pa=true then%>
	<a class=r href="plano_ensino.asp?codigo=<%=rs("id_plano")%>" onclick="NewWindow(this.href,'form_pe','695','450','yes','center');return false" onfocus="this.blur()">
	<%=rs("codmat") & " - " & rs("materia")%></a>
<%else%>
	<%=rs("codmat") & " - " & rs("materia")%>
<%end if%>
	</td>
	<td class=campo align="center" <%=barra1%> ><%=rs("naulassem")%> </td>
	<td class=campo align="center" <%=barra1%> ><%=rs("cargahoraria")%> </td>
<%
if novo=true then
%>
	<td class=campo align="center" <%=barra1%> >&nbsp;
	<a href="plano_novo.asp?codigo=<%=rs("id_plano")%>" onclick="NewWindow(this.href,'planoensino_novo','635','580','yes','center');return false" onfocus="this.blur()">
	criar</a>
	</td>
<%
else
	if pa=true then classe="campov" else classe="campol"
%>
	<td class=<%=classe%> align="center" <%=barra1%> >&nbsp;
<%
	if (pa=false) or (pa=true and session("a93")="T") then
%>	
	<a href="plano_alteracao.asp?codigo=<%=rs("id_plano")%>" onclick="NewWindow(this.href,'planoensino_altera','635','580','yes','center');return false" onfocus="this.blur()">
	alterar</a>
<%
	end if
%>
	</td>
<%
end if
%>
<%
response.write "<td class=campo>"
	ultimo="":ultimob=""
if podeduplicar=1 then
	if ( isnull(rs("justificativa")) or rs("justificativa")="" ) then
		sqld="select ultimo=max(perlet) from grades_plano where coddoc='" & coddoc & "' and codcur=" & codcur & " and codper=" & codper & " and " & _
		"grade=" & grade & " and codmat='" & rs("codmat") & "' and perlet<'" & perlet & "' and justificativa is not null "
		rs2.Open sqld, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then ultimo=rs2("ultimo")
		rs2.close
		sqle="select id_plano from grades_plano where coddoc='" & coddoc & "' and codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and codmat='" & rs("codmat") & "' and perlet='" & ultimo & "' "
		rs2.Open sqle, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then id_copia=rs2("id_plano")
		rs2.close
	end if

	if rs("nb")=0 and rs("justificativa")<>"" then	
		sqld="select ultimo=max(perlet) from grades_plano where coddoc='" & coddoc & "' and codcur=" & codcur & " and codper=" & codper & " and " & _
		"grade=" & grade & " and codmat='" & rs("codmat") & "' and perlet<'" & perlet & "' and justificativa is not null "
		rs2.Open sqld, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then ultimob=rs2("ultimo")
		rs2.close
		sqle="select id_plano from grades_plano where coddoc='" & coddoc & "' and codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and codmat='" & rs("codmat") & "' and perlet='" & ultimob & "' "
		rs2.Open sqle, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then id_copiab=rs2("id_plano")
		rs2.close
	end if
end if
if ultimo<>"" then
%>
	<a class=r href="planoensino2.asp?origem=<%=id_copia%>&destino=<%=rs("id_plano")%>">Copiar <%=ultimo%></a>
<%
end if
if ultimob<>"" then
%>
	<a class=r href="planoensino2.asp?origemb=<%=id_copiab%>&destinob=<%=rs("id_plano")%>">Acertar Biblio</a>
<%
end if
response.write "</td>"
%>

</tr>
<%
lastper=rs("serie")
rs.movenext:loop
end if 'rs.recordcount>0
rs.close
%>
</table>
<%
end if 'request.form("codcur")<>""
%>
</body>
</html>
<%
set rs=nothing
set rs2=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%>