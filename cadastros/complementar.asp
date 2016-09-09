<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a88")="N" or session("a88")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Formação Acadêmica</title>
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
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function nome1() { form.chapa.value=form.nome.value; form.submit(); }
function chapa1() { form.nome.value=form.chapa.value; form.submit(); }
--></script>

<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rsq=server.createobject ("ADODB.Recordset")
Set rsq.ActiveConnection = conexao

if request.form="" and session("compl_exec")="" then
	sql="INSERT INTO pfunc_compl (chapa) SELECT f.CHAPA FROM corporerm.dbo.pfunc f WHERE (f.chapa <'10000' OR f.chapa>='90000') and f.CHAPA collate database_default Not In (select c.chapa from pfunc_compl c);" 
	conexao.Execute sql, , adCmdText
	session("compl_exec")="S"
end if

if request.form("chapa")<>"" then session("chapaform")=request.form("chapa")
chapa=session("chapaform")
if chapa="" then chapa="0"

if request.form<>"" and request.form("nomea")<>request.form("nomeb") then
	sql="update pfunc_compl set nome='" & request.form("nomea") & "' where chapa='" & chapa & "'"
	conexao.execute sql
end if

%>
<form method="POST" action="complementar.asp" name="form">
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo colspan=2><p style="margin-top:0;margin-bottom:0;color:Blue;font-size:10pt;text-align:left">
<b>C A D A S T R O &nbsp;&nbsp;&nbsp;&nbsp; C O M P L E M E N T A R</font></p>
	</td></tr>
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=chapa%>" name="chapa" size="8" onchange="chapa1()" onfocus="javascript:window.status='Informe o chapa do professor'"></td>
	<td class=fundo>&nbsp;
	<select size="1" name="nome" onchange="nome1()" onfocus="javascript:window.status='Selecione o Nome do Professor'" >
<%
sql2="select chapa, nome from corporerm.dbo.pfunc where codsituacao<>'D' and codtipo='N' order by nome"
rsq.Open sql2, ,adOpenStatic, adLockReadOnly
response.write "<option value='0'>Selecione Funcionário....</option>"
rsq.movefirst:do while not rsq.eof
if chapa=rsq("chapa") then tempc="selected" else tempc=""
%>
          <option value="<%=rsq("chapa")%>" <%=tempc%>><%=rsq("nome")%></option>
<%
rsq.movenext:loop
rsq.close
%>
	</select></td>
</tr>
</table>
<!--
<input type="submit" value="Pesquisar" class="button" name="pesquisar" onfocus="javascript:window.status='Clique aqui para pesquisar'">
-->

<table border="1" bordercolor="Green" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td class=grupo colspan=8>Formação Acadêmica</td>
<td class=grupo align="center" colspan=1><font face="Wingdings" size=2>ê</font></td>
</tr>
<tr><td class=titulor>Tipo</td>
	<td class=titulor>Curso</td>
	<td class=titulor>Abrangência</td>
	<td class=titulor>Instituição</td>
	<td class=titulor>Local Inst.</td>
	<td class=titulor>Ano Conclusão</td>
	<td class=titulor>Data Conclusão</td>
	<td class=titulor>Comprovante</td>
	<td class=titulor>&nbsp;</td>
</tr>
<%
sql1="SELECT U.id_form, u.codprof, u.codinstrucao, t.tipo, u.curso, u.instituicao, u.localinst, u.dataconclusao, u.anoconclusao, u.comprovante, " & _
"u.abrangencia, a.descricao " & _
"FROM uprofformacao_ U, uprof_abrangencia a, uprof_tipo t " & _
"WHERE u.codprof='" & chapa & "' and u.abrangencia=a.abrangencia and t.codinstrucao=u.codinstrucao order by u.codinstrucao, u.anoconclusao "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr><td class=campo><%=rs("tipo")%></td>
	<td class=campo><%=rs("curso")%></td>
	<td class=campo><%=rs("descricao")%></td>
	<td class=campo><%=rs("instituicao")%></td>
	<td class=campo nowrap><%=rs("localinst")%></td>
	<td class=campo><%=rs("anoconclusao")%></td>
	<td class=campo><%=rs("dataconclusao")%></td>
	<td class=campo><%=rs("comprovante")%></td>
	<td class=campo>
	<a href="formacao_alteracao.asp?codigo=<%=rs("id_form")%>" onclick="NewWindow(this.href,'AlteracaoFormacao','520','240','no','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
%>
<tr><td class=campo colspan=8>
<a class=r href="formacao_nova.asp?chapa=<%=chapa%>" onclick="NewWindow(this.href,'InclusaoFormacao','535','250','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
inserir nova formação</a>
</td></tr>
</table>
<br>

<table border="1" bordercolor="Green" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=690>
<tr><td class=grupo widht=30%>Aposentadoria</td><td class=grupo width=70%>Outros empregos</td></tr>
<tr><td valign=top><!-- primeiro quadro -->

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<%
sql="select aposentado, dtaposentadoria from corporerm.dbo.pfunc where chapa='" & chapa & "' "
rsq.Open sql, ,adOpenStatic, adLockReadOnly
if rsq.recordcount>0 then aposentado=rsq("aposentado"):dtaposentadoria=rsq("dtaposentadoria")
rsq.close
sql="select tempo_trabalho, tempo_restante, outra_empresa, nome from pfunc_compl where chapa='" & chapa & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then tempo_trabalho=rs("tempo_trabalho"):tempo_restante=rs("tempo_restante"):outra_empresa=rs("outra_empresa"):nomea=rs("nome")
rs.close
%>
	<tr><td class=fundo>Aposentado? <%if aposentado=1 then response.write "<font face='Wingdings' size=5px>ü</font>" else response.write "<font face='Wingdings' size=5px>û</font>" %></td></tr>
	<tr><td class=fundo>Data Aposentadoria &nbsp;<input type="text" class="form_input10" size=8 value="<%=dtaposentadoria%>" onfocus="this.blur()"></td></tr>
	<tr><td class=fundo>Tempo Trabalho &nbsp;<input type="text" class="form_input10" size=8 value="<%=tempo_trabalho%>" onfocus="this.blur()"></td></tr>
	<tr><td class=fundo>Tempo Restante &nbsp;<input type="text" class="form_input10" size=8 value="<%=tempo_restante%>" onfocus="this.blur()"></td></tr>
</table>
<a href="aposentadoria_alteracao.asp?codigo=<%=chapa%>" onclick="NewWindow(this.href,'AlteracaoTempo','520','140','no','center');return false" onfocus="this.blur()">
<img src="../images/Folder95O.gif" border="0" width=13></a>

</td>
<td valign=top><!--  segundo quadro -->
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width=100%>
<%
%>
<tr><td class=campo colspan=5>Possui outro emprego? 
	<a href="outroemprego_alteracao.asp?codigo=<%=chapa%>" onclick="NewWindow(this.href,'AlteracaoTemEmprego','350','140','no','center');return false" onfocus="this.blur()">
	<%if outra_empresa=-1 then response.write "<font face='Wingdings' size=5px>ü</font>" else response.write "<font face='Wingdings' size=5px>û</font>" %>	</a>
	</td></tr>
<tr><td class=titulo align="center">Empresa</td>
	<td class=titulo align="center">Função</td>
	<td class=titulo align="center">Desde</td>
	<td class=titulo align="center">Até</td>
	<td class=titulo>&nbsp;</td>
</tr>
<% 
sql1="SELECT id_emp, chapa, empresa, cargo, desde, ate " & _
"FROM pfunc_empregos where chapa='" & chapa & "' order by desde "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr><td class=campo><%=rs("empresa")%></td>
	<td class=campo><%=rs("cargo")%></td>
	<td class=campo align="center"><%=rs("desde")%></td>
	<td class=campo align="center"><%=rs("ate")%></td>

	<td class=campo>
	<a href="empregos_alteracao.asp?codigo=<%=rs("id_emp")%>" onclick="NewWindow(this.href,'AlteracaoEmprego','520','200','no','center');return false" onfocus="this.blur()">
	<img src="../images/Folder95O.gif" border="0" width=13></a>
	</td>
</tr>
<%
rs.movenext:loop
end if
rs.close
%>
<tr><td class=campo colspan=5 style="border-top:2 solid #000000">
<a class=r href="empregos_nova.asp?chapa=<%=chapa%>" onclick="NewWindow(this.href,'InclusaoEmprego','535','250','yes','center');return false" onfocus="this.blur()"><img border="0" src="../images/Appointment.gif">
inserir novo emprego</a>
</td></tr>
</table>


<!-- segundo quadro -->
</td></tr>
</table>
<br>

<%
%>

<table border="1" bordercolor="Green" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td class=grupo>Outros</td></tr>
<tr><td class=titulor>Nome com acento</td></tr>
<tr><td class="campop">
	<input type="hidden" name="nomeb" value="<%=nomea%>" >
	<input type="text" name="nomea" size=80 value="<%=nomea%>" onchange="javascript:submit();">
</td></tr>
</table>

</form>


<p style="margin-top:0;margin-bottom:0;color:Blue;font-size:10pt;text-align:left">
<%
set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>