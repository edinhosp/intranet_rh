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
<title>UNIFIEO - Plano de Ensino</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:20px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

sessao=session.sessionid
codcur=request("codcur")
codper=request("codper")
grade=request("grade")
serie=request("serie")
codmat=request("codmat")
perlet=request("perlet")

sql="select max(perlet) ultimo from grades_plano where codcur=" & codcur & " and codper=" & codper & " and grade=" & grade & " and serie=" & serie & " and codmat='" & codmat & "' and (justificativa is not null) "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then ultimo=rs("ultimo") else ultimo=perlet
rs.close
	
sql="SELECT pe.id_plano, pe.CODMAT, u.MATERIA, pe.justificativa, pe.ementa, pe.objetivos_gerais, pe.unidades_tematicas, pe.metodologia, pe.avaliacao, pe.bibliografia, pe.bibliografiac, " & _
"pe.coddoc, c.CURSO, pe.codcur, pe.codper, pe.grade, pe.serie, g.NAULASSEM, g.CARGAHORARIA, c.DEPTO, pe.pa, pe.perlet " & _
"FROM grades_plano pe INNER JOIN corporerm.dbo.umaterias u ON pe.CODMAT = u.CODMAT collate database_default " & _
"INNER JOIN g2cursoeve c ON pe.coddoc = c.coddoc  " & _
"inner join corporerm.dbo.ugrade g on g.codcur=pe.codcur and g.codper=pe.codper and g.grade=pe.grade and g.codmat collate database_default=pe.codmat and g.periodo=serie " & _
"WHERE pe.codcur=" & codcur & " and pe.codper=" & codper & " and pe.grade=" & grade & " and pe.serie=" & serie & " and pe.codmat='" & codmat & "' and pe.perlet='" & ultimo & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly

tlr="style='border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000'"
tl="style='border-top:1px solid #000000;border-left:1px solid #000000'"
tr="style='border-top:1px solid #000000;border-right:1px solid #000000'"
l="style='border-left:1px solid #000000'"
r="style='border-right:1px solid #000000'"
lr="style='border-left:1px solid #000000;border-right:1px solid #000000'"
blr="style='border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000'"
bl="style='border-bottom:1px solid #000000;border-left:1px solid #000000'"
br="style='border-bottom:1px solid #000000;border-right:1px solid #000000'"
b="style='border-bottom:1px solid #000000'"
%>
<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td align="left" <%=tl%> width=110>
	<img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="110" alt=""></td>
	<td align="center" <%=tr%>>
	<b>PR�-REITORIA ACAD�MICA<br>PLANEJAMENTO ACAD�MICO</b></td>
</tr>
</table>
<%
if rs.recordcount>0 then
codigo=rs("id_plano")
%>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=campo <%=l%>>Curso:</td>
	<td class=campo><b><%=rs("curso")%><b></td>
	<td class=campo>Semestre/Per�odo:</td>
	<td class=campo <%=r%>><b><%=rs("serie")%></td>
</tr>
<tr><td class=campo <%=l%>>Disciplina:</td>
	<td class=campo><b><%=rs("materia")%></td>
	<td class=campo>C/H Total:</td>
	<td class=campo <%=r%>><b><%=rs("cargahoraria")%></td>
</tr>
<tr><td class=campo <%=l%>>Professor:</td>
	<td class=campo width=290><b>
<%
sqlp="select g.chapa1, f.nome from g2ch g, corporerm.dbo.pfunc f where g.chapa1=f.chapa collate database_default and g.codmat='" & rs("codmat") & "' " & _
"and perlet='" & rs("perlet") & "' and codcur=" & rs("codcur") & " and codper=" & rs("codper") & " and grade=" & rs("grade") & " group by g.chapa1, f.nome"
rs2.Open sqlp, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
	do while not rs2.eof
		professores=professores & rs2("nome")
		response.write rs2("nome")
		if rs2.recordcount>1 and rs2.absoluteposition<rs2.recordcount then response.write ", ":professores=professores & ", "
	rs2.movenext:loop
end if
rs2.close
%>	
	</td>
	<td class=campo>C/H Semanal:</td>
	<td class=campo <%=r%>><b><%=rs("naulassem")%></td>
</tr>
<tr><td class=campo <%=bl%>>Departamento:</td>
	<td class=campo <%=b%>><b><%=rs("depto")%></td>
	<td class=campo <%=b%>>Per�odo:</td>
	<td class=campo <%=br%>><b>
<%
sqlp="select g.turno, t.tipo from g2ch g, eturnos t where g.codmat='" & rs("codmat") & "' and g.turno=t.codturno " & _
"and perlet='" & rs("perlet") & "' and coddoc='" & rs("coddoc") & "' group by g.turno, t.tipo "
rs2.Open sqlp, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
	do while not rs2.eof
		periodos=periodos & rs2("tipo")
		response.write rs2("tipo")
		if rs2.recordcount>1 and rs2.absoluteposition<rs2.recordcount then response.write "/":periodos=periodos & "/"
	rs2.movenext:loop
end if
rs2.close
%>	
	</td>
</tr>
</table>
<%
dim texto(7),titulo(7):tamanho=0
titulo(0)="JUSTIFICATIVA"             : texto(0)=rs("justificativa")
titulo(1)="EMENTA"                    : texto(1)=rs("ementa")
titulo(2)="OBJETIVOS GERAIS"          : texto(2)=rs("objetivos_gerais")
titulo(3)="UNIDADES TEM�TICAS"        : texto(3)=rs("unidades_tematicas")
titulo(4)="METODOLOGIA"               : texto(4)=rs("metodologia")
titulo(5)="AVALIA��O"                 : texto(5)=rs("avaliacao")
titulo(6)="BIBLIOGRAFIA B�SICA"       : texto(6)=rs("bibliografia")
titulo(7)="BIBLIOGRAFIA COMPLEMENTAR" : texto(7)=rs("bibliografiac")
for a=0 to 7
	quadro=texto(a)
	if isnull(quadro)=false then quadro=replace(quadro,chr(13)&chr(10),"<br>")
	texto(a)=quadro
next

for a=0 to 7
tam=len(texto(a)):tamanho=tamanho+tam
if tamanho>53*60 then
end if
if a=6 then
	texto(6)=""
	sql="select p.id_biblio, p.id_plano, p.complementar, p.cod_acervo, p.ordem, p.referencia digitada, p.status, b.referencia pesquisada " & _
	"from grades_plano_biblio p left join pe_biblio b on b.cod_acervo=p.cod_acervo " & _
	"where id_plano=" & codigo & " and complementar=0 " & _
	"order by ordem"
	rs2.Open sql, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
	if rs2("pesquisada")<>"" then referencia=rs2("pesquisada") else referencia=rs2("digitada")
	texto(6)=texto(6) & referencia & "<br>"
	rs2.movenext:loop
	rs2.close
end if
if a=7 then
	texto(7)=""
	sql="select p.id_biblio, p.id_plano, p.complementar, p.cod_acervo, p.ordem, p.referencia digitada, p.status, b.referencia pesquisada " & _
	"from grades_plano_biblio p left join pe_biblio b on b.cod_acervo=p.cod_acervo " & _
	"where id_plano=" & codigo & " and complementar=1 " & _
	"order by ordem"
	rs2.Open sql, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
	if rs2("pesquisada")<>"" then referencia=rs2("pesquisada") else referencia=rs2("digitada")
	texto(7)=texto(7) & referencia & "<br>"
	rs2.movenext:loop
	rs2.close
end if

%>
<br><%'=tamanho & " / " & tamanho / 60%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=campo <%=tlr%> ><b><%=titulo(a)%></td></tr>
<tr><td class=campo <%=blr%> ><%=texto(a)%></td></tr>
</table>
<%
next
%>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class=campo align="left"><b><i>
<%
	if rs("pa")=1 or rs("pa")=true then valida="Validado pelo Planejamento Acad�mico." else valida="N�o validado pelo Planejamento Acad�mico"
	response.write valida
%>
	</td>
	<td class="campop" align="right"><b><i>Per�odo Letivo: <%=rs("perlet")%></td>
</tr>
</table>

<%
else
%>
<table border="0" cellpadding="5" cellspacing="0" width="650" style="border-collapse: collapse">
<tr><td align="left" <%=bl%> style="border-top:1px dashed #000000" width=110></td>
	<td align="center" <%=br%> style="border-top:1px dashed #000000"><b><font color=red>Sem plano de ensino cadastrado no momento.</td>
</tr>
</table>


<%
end if 'rs.recordcount
%>
<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>