<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a81")="N" or session("a81")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Controle de Assistência Médica</title>
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

sqla="select f.nome, f.codsituacao, f.chapa, f.admissao, f.funcao, f.codsecao, f.Secao, f.grauinstrucao, " & _
"f.demissao, f.estadocivil, f.estcivil, f.dtnascimento, " & _
"a.empresa, a.plano, a.codigo, ae.operadora, s.cartaosus " & _
"FROM qry_funcionarios f inner join assmed_beneficiario a on a.chapa=f.chapa collate database_default " & _
"inner join assmed_empresa ae on a.empresa=ae.codigo " & _
"left join corporerm.dbo.vpcompl s on s.codpessoa=f.codpessoa "
if request.form("chapa")="" then chapa=request("codigo") else chapa=request.form("chapa")
sqlb="where f.CHAPA='" & chapa & "' "
sqlc="ORDER BY f.CHAPA"
sql1=sqla & sqlb & sqlc
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount=0 then
	sql3="insert into assmed_beneficiario (chapa, empresa) select '" & chapa & "', 'N'"
	conexao.execute sql3
	rs.close
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
end if
%>
<form name="form" method="POST" action="controle_ver.asp" >
<input type="hidden" name="chapa" value="<%=chapa%>">
<p style="margin-top: 0; margin-bottom: 0" class=titulo>
<% if session("a81")="T" then %>
<a href="controle_ver.asp?codigo=<%=chapa%>" onMouseOver="window.status='Clique aqui para atualizar após as alterações'; return true" onMouseOut="window.status=''; return true" onmouseover >
<img border="0" src="../images/write.gif" alt="Clique para atualizar">
<font size="1">!</font>
</a>
<% end if %>
CONTROLE DE ASSISTÊNCIA MÉDICA</p>
<%
'rs.movefirst
'do while not rs.eof 
'session("chapabolsa")=rs("chapa")
'session("chapabolsanome")=rs("nome")
sql2="select descricao from corporerm.dbo.pcodsituacao where codcliente='" & rs("codsituacao") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
sit=rs2("descricao")
rs2.close
sql2="select descricao from corporerm.dbo.pcodinstrucao where codcliente='" & rs("grauinstrucao") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
titulacao=rs2("descricao")
rs2.close
idadetit=int((now()-rs("dtnascimento"))/365.25)
%>
<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="650">
<tr><td class=grupo>Dados Pessoais</td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td width="100%" valign="top">
<!-- quadro -->

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulor>&nbsp;Chapa:</td>
	<td class=titulor>&nbsp;Nome:</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rs("chapa")%>&nbsp;</td>
	<td class="campor"><b>&nbsp;<%=rs("nome")%>&nbsp;</b></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulor>&nbsp;Situação:</td>
	<td class=titulor>&nbsp;Admissão:</td>
	<td class=titulor>&nbsp;Função:</td>
	<td class=titulor>&nbsp;Cartão SUS:</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=sit%>&nbsp;<%if rs("codsituacao")="D" then response.write rs("demissao")%></td>
	<td class="campor">&nbsp;<%=rs("admissao")%>&nbsp;</td>
	<td class="campor">&nbsp;<%=rs("funcao")%>&nbsp;</td>
	<td class="campor">&nbsp;<%=rs("cartaosus")%>&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulor>&nbsp;Instrução/Titulação:</td>
	<td class=titulor>&nbsp;Estado Civil</td>
	<td class=titulor>&nbsp;Seção:</td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=titulacao%>&nbsp;</td>
	<td class="campor">&nbsp;<%=rs("estcivil")%>&nbsp;</td>
	<td class="campor">&nbsp;<%=rs("codsecao")%>&nbsp;<%=rs("secao")%></td>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="500">
<tr>
	<td class=titulor>&nbsp;Empresa Atual</td>
	<td class=titulor>&nbsp;Plano Atual</td>
	<td class=titulor>&nbsp;Código </td>
	<td class=titulor>&nbsp;Nascimento </td>
	<td class=titulor>&nbsp;Idade </td>
</tr>
<tr>
	<td class="campor">&nbsp;<%=rs("operadora")%>&nbsp;</td>
	<td class="campor">&nbsp;<%=rs("plano")%>&nbsp;    </td>
	<td class="campor">&nbsp;<%=rs("codigo")%>&nbsp;   </td>
	<td class="campor">&nbsp;<%=rs("dtnascimento")%>&nbsp;</td>
	<td class="campor">&nbsp;<%=idadetit%>&nbsp;       </td>
</tr>
</table>

<!-- quadro inicio mudanca-->
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="490">
<tr><th class=titulo colspan=10>Histórico das mudanças de planos</th></tr>
<tr>
	<td class=titulor align="center">Empresa</td>
	<td class=titulor align="center">Plano</td>
	<td class=titulor align="center">Código</td>
	<td class=titulor align="center">Inclusão</td>
	<td class=titulor align="center">Cobrança</td>
	<td class=titulor align="center">Final</td>
	<td class=titulor align="center"><font face="Wingdings">ü</font></td>
	<td class=titulor align="center">Oper</td>
	<td class=titulor align="center">&nbsp;</td>
	<td class=titulor align="center">&nbsp;</td>
</tr>
<%
sql2="SELECT m.id_mudanca, m.empresa, e.operadora, m.plano, m.codigo, " & _
"m.inclusao, m.ivigencia, m.fvigencia, m.compr, m.oper, m.uoper " & _
"FROM assmed_mudanca m, assmed_empresa e " & _
"where m.chapa='" & rs("chapa") & "' and m.empresa=e.codigo order by m.ivigencia, m.fvigencia "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
linhas=rs2.recordcount
rs2.movefirst
inicio=1
do while not rs2.eof
final=rs2("fvigencia"):if isnull(final) then final=now-1
if cdate(final)<now then fundo="style='background-color:#ffcccc'" else fundo="style='background-color:#ccffcc'"
%>
<tr>
	<td <%=fundo%> class="campor"><%=rs2("operadora") %></td>
	<td <%=fundo%> class="campor"><%=rs2("plano") %>    </td>
	<td <%=fundo%> class="campor"><%=rs2("codigo") %>   </td>
	<td <%=fundo%> class="campor"><%=rs2("inclusao") %></td>
	<td <%=fundo%> class="campor"><%=rs2("ivigencia") %></td>
	<td <%=fundo%> class="campor"><%=rs2("fvigencia") %></td>
	<td <%=fundo%> class="campor">&nbsp;<%if rs2("compr")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td <%=fundo%> class="campor" align="center">&nbsp;<%=rs2("oper")%>/<%=rs2("uoper")%></td>
	<td <%=fundo%> class="campor">&nbsp; 
	<% if session("a81")="T" or session("a81")="C" then %>
		<a href="controle_alteracao.asp?codigo=<%=rs2("id_mudanca")%>" onclick="NewWindow(this.href,'AlteracaoPlano','520','270','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/folder95o.gif" alt="Alterar este plano" width=13></a>
	<% end if %>
	</td>
	<%if inicio=1 then %>
	<td class="campor" rowspan=<%=linhas%> valign="center" align="center">
	<% if session("a81")="T" or session("a81")="C" then %>
		<a href="controle_nova.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'InclusaoPlano','520','270','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/Appointment.gif" alt="inserir novo plano"></a>
	<% end if %>
	</td>
	<% end if 'inicio=1%>
</tr>
<%
rs2.movenext
inicio=0
loop
else ' sem registros/planos
%>
<tr>
	<td class="campor" colspan=8>&nbsp;</td>
	<td class="campor" rowspan=1 valign="center" align="center">
	<% if session("a81")="T" or session("a81")="C" then %>
		<a href="controle_nova.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'InclusaoPlano','520','270','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/Appointment.gif" alt="inserir novo plano"></a>
	<% end if %>
	</td>
</tr>
<%
end if
rs2.close
%>

</table>
<!-- quadro fim mudanca -->
	</td>
	<td width="170" valign="top">
		<img border="0" src="../func_foto.asp?chapa=<%=rs("chapa")%>"  width="150">
	</td>
</tr>
</table>
<%
'rs.movenext
'loop
if request.form("bases")="ON" then cbases="checked" else cbases=""
%>
<input type="checkbox" name="bases" value="ON" <%=cbases%> onClick="javascript:submit()">Mostrar todos dependentes e planos

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="650">
<tr><td class=grupo>Dependentes</td></tr>
</table>

<%
sql2="select d.NRODEPEND, NOME, CPF, DTNASCIMENTO, SEXO, ESTADOCIVIL, GRAUPARENTESCO, p.DESCRICAO as parentesco, DATAENTREGACERTIDAO dt_evento, CARTAOSUS, MAE " & _
", planosativos=(select COUNT(*) from assmed_dep_mudanca where GETDATE() between inclusao and fvigencia and chapa=d.CHAPA collate database_default and nrodepend=d.NRODEPEND) " & _
"from corporerm.dbo.PFDEPEND d inner join corporerm.dbo.PCODPARENT p on p.CODCLIENTE=d.GRAUPARENTESCO " & _
"left join corporerm.dbo.PFDEPENDCOMPL c on c.CHAPA=d.CHAPA and c.NRODEPEND=d.NRODEPEND " & _
"where d.CHAPA='" & rs("chapa") & "' order by grauparentesco, nome "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
%>
<!-- dependentes -->
<table border="1" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulor align="center">#</td>
	<td class=titulor align="center">Nome Dependente</td>
	<td class=titulor align="center">Sexo</td>
	<td class=titulor align="center">Nascimento</td>
	<td class=titulor align="center">Idade</td>
	<td class=titulor align="center">Parentesco</td>
	<td class=titulor align="center">Dt.Evento</td>
</tr>
<%
rs2.movefirst
do while not rs2.eof
idade=int((now()-rs2("dtnascimento"))/365.25)
if (request.form("bases")="ON") or (request.form("bases")="" and rs2("planosativos")>0) then
%>
<tr>
	<td class="campor" rowspan="3" align="center"><%=rs2("nrodepend")%>
	<br><% if session("a81")="T" or session("a81")="C" then %>
	<a href="ctr_dep_plano_nova.asp?chapa=<%=rs("chapa")%>&nrodepend=<%=rs2("nrodepend")%>" onclick="NewWindow(this.href,'InclusaoPlanoDep','520','270','no','center');return false" onfocus="this.blur()">
	<img border="0" src="../images/Appointment.gif" alt="inserir novo plano"></a>
	<% end if %>
	</td>
	<td class="campor" rowspan="1" width=240 valign=top><b><%=rs2("nome")%></b></td>
	<td class="campoa"r align="center"><%=rs2("sexo")%></td>
	<td class="campoa"r align="center"><%=rs2("dtnascimento")%></td>
	<td class="campoa"r align="center"><%=idade%></td>
	<td class="campoa"r><%=rs2("parentesco")%></td>
	<td class="campoa"r><%=rs2("dt_evento")%></td>
</tr>
<tr>
	<td class="campor" colspan=2 align="left"><font color="#808080">Nome da mãe: </font><%=rs2("mae")%></td>
	<td class="campor" colspan=2 valign=top><font color="#808080">
	<%if idade>=18 then response.write "CPF: </font>" & rs2("cpf")%></td>
	<td class="campor" colspan=2><font color="#808080">Cartão Sus:</font> <%=rs2("cartaosus")%></td>
</tr>
<tr>
	<td class=campo colspan=6 valign="top" align="center">
	
<!-- quadro inicio mudanca dependentes -->
<table border="1" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="">
<tr>
	<td class=titulor valign="top" align="center">Empresa</td>
	<td class=titulor valign="top" align="center">Plano  </td>
	<td class=titulor valign="top" align="center">Código </td>
	<td class=titulor valign="top" align="center">Inclusão</td>
	<td class=titulor valign="top" align="center">Cobr.</td>
	<td class=titulor valign="top" align="center">Término</td>
	<td class=titulor valign="top" align="center">&nbsp; </td>
	<td class=titulor valign="top" align="center">Oper   </td>
	<td class=titulor valign="top" align="center">&nbsp; </td>
</tr>
<%
sql3="SELECT m.id_mud, m.id_dep, m.empresa, e.operadora, m.plano, m.codigo, " & _
"m.inclusao, m.ivigencia, m.fvigencia, m.compr, m.oper, m.uoper " & _
"FROM assmed_dep_mudanca m, assmed_empresa e " & _
"where m.chapa='" & rs("chapa") & "' and nrodepend=" & rs2("nrodepend") & " and m.empresa=e.codigo order by m.ivigencia "
rs3.Open sql3, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then
%>
<%
rs3.movefirst
do while not rs3.eof
if cdate(rs3("fvigencia"))<now then fundo="style='background-color:#ffcccc'" else fundo="style='background-color:#ccffcc'"
%>
<tr>
	<td <%=fundo%> class="campor" valign="top" style="border-bottom: 2 solid #000080"><%=rs3("operadora") %></td>
	<td <%=fundo%> class="campor" valign="top" style="border-bottom: 2 solid #000080"><%=rs3("plano") %>    </td>
	<td <%=fundo%> class="campor" valign="top" style="border-bottom: 2 solid #000080"><%=rs3("codigo") %>   </td>
	<td <%=fundo%> class="campor" valign="top" style="border-bottom: 2 solid #000080"><%=rs3("inclusao") %></td>
	<td <%=fundo%> class="campor" valign="top" style="border-bottom: 2 solid #000080"><%=rs3("ivigencia") %></td>
	<td <%=fundo%> class="campor" valign="top" style="border-bottom: 2 solid #000080"><%=rs3("fvigencia") %></td>
	<td <%=fundo%> class="campor" valign="top" style="border-bottom: 2 solid #000080">&nbsp;<%if rs3("compr")=-1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td <%=fundo%> class="campor" valign="top" style="border-bottom: 2 solid #000080" align="center">&nbsp;<%=rs3("oper")%>/<%=rs3("uoper")%></td>
	<td <%=fundo%> class="campor" valign="top" style="border-bottom: 2 solid #000080">&nbsp; 
	<% if session("a81")="T" or session("a81")="C" then %>
		<a href="ctr_dep_plano_alteracao.asp?codigo=<%=rs3("id_mud")%>" onclick="NewWindow(this.href,'AlteracaoPlanoDep','520','270','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/folder95o.gif" alt="Alterar este plano" width=13></a>
	<% end if %>
	</td>

	</tr>
<%
rs3.movenext
loop
end if
rs3.close
%>
</table>
<!-- quadro fim mudanca dependentes -->
	</td>
<%
end if 'planos ativos>0 or request.form(cbases)=on
rs2.movenext
loop
else
%>
<tr>
	<td colspan="8" class=grupo>&nbsp;Sem dependentes cadastrados</td>
</tr>
<%
end if
rs2.close
%>
</table>
<!-- final dependentes -->
<!-- inicio acertos -->
<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="400">
<tr><td class=grupo>Acertos</td></tr>
</table>
<table border="1" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="400">
<%
sql4="SELECT id_acerto, data_acerto, empresa, descricao, valor_acerto, reembolso " & _
"FROM assmed_acertos " & _
"where chapa='" & rs("chapa") & "' and data_acerto>getdate()-90 order by data_acerto "
rs2.Open sql4, ,adOpenStatic, adLockReadOnly
if rs2.recordcount=0 then rowacerto=2 else rowacerto=rs2.recordcount+1
%>
<tr>
	<td class=titulor align="center">Data</td>
	<td class=titulor align="center">Empresa</td>
	<td class=titulor align="center">Descrição</td>
	<td class=titulor align="center">Vr.Acerto</td>
	<td class=titulor align="center">Reembolso</td>
	<td class="campor" rowspan="<%=rowacerto%>" align="center">
	<% if session("a81")="T" or session("a81")="C" then %>
		<a href="acerto_nova.asp?chapa=<%=rs("chapa")%>" onclick="NewWindow(this.href,'InclusaoAcerto','520','270','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/Appointment.gif" alt="inserir novo acerto de NF"></a>
	<% end if %>
	</td>
</tr>
<%
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
%>
<tr>
	<td class="campor" align="center"><a class=r href="acerto_alteracao.asp?codigo=<%=rs2("id_acerto")%>" onclick="NewWindow(this.href,'AlteracaoAcerto','520','230','no','center');return false" onfocus="this.blur()">
	<%=rs2("data_acerto")%></a></td>
	<td class="campor"><font size="1"><%=rs2("empresa")%></td>
	<td class="campor"><font size="1"><%=rs2("descricao")%></td>
	<td class="campor" align="center"><%=formatnumber(rs2("valor_acerto"),2)%></td>
	<td class="campor" align="center"><%=formatnumber(rs2("reembolso"),2)%></td>
</tr>
<%
rs2.movenext
loop
rs2.close
else
%>
<tr>
	<td class=grupo colspan="5">&nbsp;Sem Acertos cadastrados</td>
</tr>
<%
end if
%>
</table>
<!-- final acertos -->

<% if rs("codsituacao")="D" then %>
		<br>
		<a href="excecao.asp?codigo=<%=rs("chapa")%>" onclick="NewWindow(this.href,'InclusaoAcerto','500','150','no','center');return false" onfocus="this.blur()">
		<img border="0" src="../images/Appointment.gif" alt="inserir exceção de exclusão">Inserir exceção de Permanência</a>
<% end if %>

</form>
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