<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a30")="N" or session("a30")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Processo IFIP</title>
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
conexao.Open Application("conexao")

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

sql="SELECT i.*, ifip_wstatus.desc_status " & _
"FROM ifip_cadastro AS i LEFT JOIN ifip_wstatus ON i.status = ifip_wstatus.id_status " & _
"WHERE i.id_ifip=" & request("codigo")
rs.Open sql, ,adOpenStatic, adLockReadOnly
id_ifip=request("Codigo")	
%>

<p style="margin-top: 0; margin-bottom: 0" class=titulo>
<% if session("a30")="T" then %>
<a href="pesquisaver.asp?codigo=<%=id_ifip%>" onMouseOver="window.status='Clique aqui para atualizar após as alterações'; return true" onMouseOut="window.status=''; return true" onmouseover >
<img border="0" src="../images/write.gif" alt="Clique para atualizar"></a>
<% end if %>
PROCESSO IFIP</p>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><td class=grupo>Dados do processo</td></tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>Nº do Processo</td>
	<td class=titulor>Status</td>
	<td class=titulor>Título da Pesquisa</td>
</tr>
<tr>
	<td class="campor"><b><%=rs("num_processo")%></b></td>
	<td class="campor"><%=rs("desc_status")%></td>
	<td class="campor"><b><%=rs("titulo_pesquisa")%></b></td>
</tr>
</table>

<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="560">
<tr>
	<td class=titulor>Linha de Pesquisa</td>
	<td class=titulor>Área de conhecimento</td>
	<td class=titulor>Horas</td>
</tr>
<tr>
	<td class="campor"><%=rs("linha_pesquisa")%></td>
	<td class="campor"><%=rs("area_conhecimento")%></td>
	<td class="campor"><%=rs("horas_semanais")%> p/sem.</td>
</tr>
</table>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" width="560">
<tr><td class=campo valign=top>

	<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="390">
	<tr>
		<td class=titulor>Início</td>
		<td class=titulor>Término</td>
		<td class=titulor>Vigência</td>
		<td class=titulor>Valor</td>
	</tr>
	<tr>
    	<td class="campor"><%=rs("dt_entrada")%></td>
    	<td class="campor"><%=rs("dt_termino")%></td>
    	<td class="campor"><%=rs("vigencia")%> meses</td>
    	<td class="campor"><%=rs("aprov_valor")%></td>
	</tr>
	</table>

	<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="390">
	<tr><td class=titulor>Observações</td></tr>
	<tr><td class="campor"><%=rs("observacoes")%></td></tr>
	</table>

</td><td class=campo>

	<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width=165>
	<th class=titulo colspan=2>Aprovações</th>
	<tr><td class=titulor>Depto.</td>
    	<td class="campor" width=70><%=rs("aprov_depto")%></td>
	</tr>
	<tr><td class=titulor>IFIP</td>
    	<td class="campor"><%=rs("aprov_ifip")%></td>
	</tr>
	<tr><td class=titulor>Pró-Adm.</td>
    	<td class="campor"><%=rs("aprov_proadm")%></td>
	</tr>
	<tr><td class=titulor>CONSEPE</td>
    	<td class="campor"><%=rs("aprov_consepe")%></td>
	</tr>
	<tr><td class=titulor>Ciência em</td>
    	<td class="campor"><%=rs("aprov_ciencia")%></td>
	</tr>
	</table>

	</td></tr>  
</table>
<hr>

<!-- quadro inicio titulares-->
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><td class=grupo colspan=6 valign="center"><% if session("a30")="T" then %>
	<a href="titular_nova.asp?codigo=<%=id_ifip%>" onclick="NewWindow(this.href,'InclusaoTitular','510','150','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/Appointment.gif" alt="inserir novo titular"></a>
<% end if %>
Titulares</td></tr>
<tr>
	<td class=titulor align="center">Tipo</td>
	<td class=titulor align="center">Chapa</td>
	<td class=titulor align="center">Nome</td>
	<td class=titulor align="center">Titulação</td>
	<td class=titulor align="center">Seção</td>
	<td class=titulor align="center">&nbsp;</td>
</tr>
<%
sql2="select * from ifip_titulares WHERE id_ifip=" & request("codigo") & " ORDER BY tp_docente desc "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof

sql3="select id_titular, desc_titular from ifip_wtitular where id_titular='" & rs2("tp_docente") & "'"
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
tp_docente=rs3("desc_titular")
rs3.close

sql3="select funcionario, instrucao, secao from qry_funcionarios where chapa='" & rs2("chapa") & "' "
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
foto=80
%>
<tr>
	<td class="campor" valign=top><%=rs2("tp_docente")%> - <%=tp_docente%></td>
	<td class="campor" align="center" valign=top>
	<% if session("a30")="T" then %>
	<a class=r href="titular_alteracao.asp?codigo=<%=rs2("id_tit")%>" onclick="NewWindow(this.href,'AlteraTitular','510','150','no','center');return false" onfocus="this.blur()">
	<%=rs2("chapa") %></a><%else%><%=rs2("chapa") %><%end if%>
	</td>
	<td class="campor" valign=top><%=rs3("funcionario") %></td>
	<td class="campor" valign=top><%=rs3("instrucao")%></td>
	<td class="campor" valign=top><%=rs3("secao")%></td>
	<td class=fundo width="<%=foto%>" valign=top>
		<img border="0" src="../func_foto.asp?chapa=<%=rs2("chapa")%>"  width="<%=foto%>">
	</td>
<%
rs3.close
rs2.movenext
loop
else  'recordcount
%>
<tr><td class="campor" colspan=5>&nbsp;</td></tr>
<%
end if ' recordcount
rs2.close
%>
</table>
<!-- quadro fim titulares -->

<hr>
<!-- quadro inicio relatorios-->
<%
sql2="select max(sequencia) as topseq from ifip_relatorios where id_ifip=" & request("codigo")
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2("topseq")="" or isnull(rs2("topseq")) then sequencia=1 else sequencia=int(rs2("topseq"))+1
rs2.close
%>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><td class=grupo colspan=4><% if session("a30")="T" then %>
	<a href="relatorio_nova.asp?codigo=<%=id_ifip%>&sequencia=<%=sequencia%>" onclick="NewWindow(this.href,'InclusaoRelatório','510','150','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/Appointment.gif" alt="inserir novo relatório"></a>
<% end if %>
Relatórios</td></tr>
<tr>
	<td class=titulor align="center">Seq.</td>
	<td class=titulor align="center">Periodicidade</td>
	<td class=titulor align="center">Data Prevista</td>
	<td class=titulor align="center">Data Entrega</td>
</tr>
<%
sql2="select * from ifip_relatorios WHERE id_ifip=" & request("codigo") & " ORDER BY sequencia "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof

sql3="select id_periodicidade, desc_periodicidade from ifip_wperiodicidade where id_periodicidade='" & rs2("periodicidade") & "'"
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
periodicidade=rs3("desc_periodicidade")
rs3.close
%>
<tr>
	<td class="campor" valign=top align="center">
	<% if session("a30")="T" then %>
		<a class=r href="relatorio_alteracao.asp?codigo=<%=rs2("id_rel")%>" onclick="NewWindow(this.href,'AlteraRelatorio','510','150','no','center');return false" onfocus="this.blur()">
		<%=rs2("sequencia")%></a><%else%><%=rs2("sequencia")%><%end if%></td>
	<td class="campor" valign=top>&nbsp;<%=rs2("periodicidade") %> - <%=periodicidade%></td>
	<td class="campor" valign=top align="center"><%=rs2("dt_prevista") %></td>
	<td class="campor" valign=top align="center"><%=rs2("dt_apresentacao")%></td>
<%
rs2.movenext:loop
else  'recordcount
%>
	<tr><td class="campor" colspan=4>&nbsp;</td></tr>
<%
end if ' recordcount
rs2.close
%>
</table>
<!-- quadro fim relatorios -->

<hr>
<!-- quadro inicio publicacoes-->
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><td class=grupo colspan=5><% if session("a30")="T" then %>
	<a href="publicacoes_nova.asp?codigo=<%=id_ifip%>" onclick="NewWindow(this.href,'InclusaoPublicacao','510','180','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/Appointment.gif" alt="inserir nova publicação"></a>
<% end if %>
Publicações</td></tr>
<tr>
	<td class=titulor width=100 align="center">Dt.Publicação</td>
	<td class=titulor align="center">Título</td>
	<td class=titulor align="center">Local</td>
	<td class=titulor width=50 align="center">Número</td>
	<td class=titulor width=40 align="center">Página</td>
</tr>
<%
sql2="select * from ifip_publicacoes WHERE id_ifip=" & request("codigo") & " ORDER BY dt_publicacao "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
%>
<tr>
	<td class="campor" valign=top align="center">
	<% if session("a30")="T" then %>
	<a class=r href="publicacoes_alteracao.asp?codigo=<%=rs2("id_publ")%>" onclick="NewWindow(this.href,'AlteraPublicacao','510','180','no','center');return false" onfocus="this.blur()">
	<%=rs2("dt_publicacao")%></a><%else%><%=rs2("dt_publicacao")%><%end if%></td>
	<td class="campor" valign=top><%=rs2("titulo")%></td>
	<td class="campor" valign=top><%=rs2("local") %></td>
	<td class="campor" valign=top><%=rs2("numero")%></td>
	<td class="campor" valign=top><%=rs2("pagina")%></td>
</tr>
<%
rs2.movenext
loop
else  'recordcount
%>
	<tr><td class="campor" colspan=4>&nbsp;</td></tr>
<%
end if ' recordcount
rs2.close
%>
</table>
<!-- quadro fim publicacoes -->

<hr>
<!-- quadro inicio historicos-->
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="560">
<tr><td class=grupo colspan=4><% if session("a30")="T" then %>
<a href="historico_nova.asp?codigo=<%=id_ifip%>" onclick="NewWindow(this.href,'InclusaoHistorico','510','180','no','center');return false" onfocus="this.blur()">
<img border="0" src="../images/Appointment.gif" alt="inserir novo histórico"></a>
<% end if %>
Históricos</td></tr>
<tr>
	<td class=titulor align="center">Seq.</td>
	<td class=titulor align="center">Data</td>
	<td class=titulor align="center">Histórico</td>
	<td class=titulor align="center">Observação</td>
</tr>
<%
sql2="select * from ifip_historico WHERE id_ifip=" & request("codigo") & " ORDER BY sequencia, dt_historico "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
%>
<tr>
	<td class="campor" valign=top align="center">
	<% if session("a30")="T" then %>
	<a class=r href="historico_alteracao.asp?codigo=<%=rs2("id_hist")%>" onclick="NewWindow(this.href,'AlteraHistorico','510','180','no','center');return false" onfocus="this.blur()">
	<%=rs2("sequencia")%></a><%else%> <%=rs2("sequencia")%> <%end if%></td>
	<td class="campor" valign=top>&nbsp;<%=rs2("dt_historico")%></td>
	<td class="campor" valign=top><%=rs2("historico") %></td>
	<td class="campor" valign=top><%=rs2("observacao")%></td>
</tr>
<%
rs2.movenext:loop
else  'recordcount
%>
	<tr><td class="campor" colspan=4>&nbsp;</td></tr>
<%
end if ' recordcount
rs2.close
%>
</table>
<!-- quadro fim historicos -->

</body>
</html>
<%
rs.close
set rs=nothing
set rs2=nothing
set rs3=nothing
conexao.close
set conexao=nothing
%>