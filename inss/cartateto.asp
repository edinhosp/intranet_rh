<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a39")="N" or session("a39")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Carta de Teto Máximo</title>
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
conexao.Open application("conexao")
teste=0
valorteto=5189.82

if request("codigo")<>"" or request.form<>"" then
	if request.form="" then temp=request("codigo")
	if request("codigo")="" then temp=request.form("chapa")
	if isnumeric(temp) then
		info=1
		temp=numzero(temp,5)
		sqlb="AND f.CHAPA='" & temp & "' "
	else
		info=2
		sqlb="AND f.nome like '%" & temp & "%' order by f.nome"
	end if
 	sqla="SELECT F.CHAPA, F.NOME, C.NOME AS FUNCAO, P.CARTEIRATRAB, P.SERIECARTTRAB, " & _
 	"F.DATAADMISSAO, P.SEXO, f.codsituacao, f.pispasep " & _
	"FROM corporerm.dbo.PFUNC F, corporerm.dbo.PPESSOA P, corporerm.dbo.PFUNCAO C " & _
	"WHERE F.CODPESSOA = P.CODIGO AND F.CODFUNCAO = C.CODIGO "
	
	select case request.form("tipo")
		case "2"
			mes1=request.form("T1")
			mes2=request.form("T2")
			if left(mes1,2)="13" then mes1="12/" & mid(mes1,4,len(mes1)-3) & "-13º Salário"
			if left(mes2,2)="13" then mes2="12/" & mid(mes2,4,len(mes2)-3) & "-13º Salário"
			texto=", no período de "
			texto=texto & mes1
			if mes1=mes2 then
				texto=texto
			else
				texto=texto & " a " & mes2
			end if
			texto=texto & ""
		case "1"
			texto=""
	end select
	if request.form("tipo")="" then session("textoper")=session("textoper") else session("textoper")=texto
 	if request.form("empresa")="" then session("tetoempresa")=session("tetoempresa") else session("tetoempresa")=request.form("empresa")
 	if request.form("cnpj")="" then session("tetocnpj")=session("tetocnpj") else session("tetocnpj")=request.form("cnpj")
	if request.form("valorteto")="" then session("tetovalor")=session("tetovalor") else session("tetovalor")=request.form("valorteto")
	if request.form("assinatura")="" then session("tetoassinatura")=session("tetoassinatura") else session("tetoassinatura")=request.form("assinatura")
 	'session("40tipo")=session("40tipo")
	'if request.form("tipo")="" then session("40tipo")=request("tipo")
	'if request.form("tipo")<>"" then session("40tipo")=request.form("tipo")
	sql1=sqla & sqlb
	set rs=server.createobject ("ADODB.Recordset")
	Set rs.ActiveConnection = conexao
	set rs2=server.createobject ("ADODB.Recordset")
	Set rs2.ActiveConnection = conexao
	set rsi=server.createobject ("ADODB.Recordset")
	Set rsi.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	session("chapa")=rs("chapa")
	chapa=rs("chapa")
	nome=rs("nome")
	temp=0
	if rs.recordcount>0 and session("cartateto")<>"L" then temp=2
else
	temp=1
end if

if temp=1 then
session("cartateto")="F"
session("tetocnpj")=""
session("tetoempresa")=""
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>
Seleção do funcionário para emissão de comprovante
<form method="POST" action="cartateto.asp" name="form">
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;Chapa/Nome <input type="text" name="chapa" size="20" class="form_box" value="<%=session("chapa")%>">
  </p>
  <p style="margin-top: 0; margin-bottom: 0"><input type="radio" value="1" checked name="tipo">sem
  período</p>
  <p style="margin-top: 0; margin-bottom: 0"><input type="radio" value="2" name="tipo">no
  período de <input type="text" name="T1" size="8" value="<%=numzero(month(now),2) & "/" & year(now)%>"> a <input type="text" name="T2" size="8" value="<%=numzero(month(now)+0,2) & "/" & year(now)%>"></p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;Empresa a que se destina: <input type="text" name="empresa" size="30" class="form_box" value="<%=session("tetoempresa")%>">
  C.N.P.J. <input type="text" name="cnpj" size="18" class="form_box" value="<%=session("tetocnpj")%>">
  </p>
  <p style="margin-top: 0; margin-bottom: 0"><input type="checkbox" name="valorteto" value="ON">Imprime valor do teto?</p>
  <p style="margin-top: 0; margin-bottom: 0"><input type="checkbox" name="assinatura" value="ON">Imprime assinatura?</p>
  <p style="margin-top: 0; margin-bottom: 0">
  <input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>

<%
elseif temp=0 then
session("cartateto")="C"
'if request.form<>"" then
if rs("sexo")="F" then v1="a" else v1="o"
if rs("sexo")="F" then v2="a" else v2=""
if rs("sexo")="F" then v3="à" else v3="ao"
if rs("sexo")="F" then v4="a" else v4="e"
if session("tetovalor")="ON" then textovalor=", atualmente fixado em R$ " & formatnumber(valorteto,2) & "" else textovalor="."
%>
<div align="center">
  <center>
<table border="0" cellpadding="5" width="620" cellspacing="0" height="1000">
<tr>
	<td width="100%"><img border="0" src="../images/aguia.jpg"></td>
</tr>
<tr>
	<td width="100%">&nbsp;</td>
</tr>
<tr>
	<td width="100%">
		<p align="center"><b><font size="4">DECLARAÇÃO DE TETO MÁXIMO</font></b>
	</td>
</tr>
</center>
<tr>
	<td width="100%">
	<p>&nbsp;</p>
	<p align="justify">Declaramos aos orgãos interessados e comprovação de eventual
	fiscalização do I.N.S.S., que <%=v1%> Sr<%=v2%>. <%=rs("nome")%>,
	portador<%=v2%> da CTPS nº <%=rs("carteiratrab")%> / <%=rs("seriecarttrab")%>, PIS/PASEP nº <%=rs("pispasep")%>, é
	funcionári<%=v1%> desta Instituição de Ensino Superior desde
	<%=rs("dataadmissao")%>, exercendo a função de <%=rs("funcao")%>, contribuindo para a
	Previdência Social com o teto máximo vigente<%=session("textoper")%><%=textovalor%>
	</p>
	<p align="justify">Salientamos que, ocorrendo qualquer alteração na contribuição ou
	no desligamento d<%=v1%> funcionári<%=v1%> desta Instituição, caberá a el<%=v4%> a comunicação
	aos interessados em tempo hábil, ficando esta instituição isenta de qualquer
	responsabilidade quanto a possíveis consequências.</p>
	<p align="justify">Recebam nossas considerações.</p>
	<p align="justify">&nbsp;</p>
	<p><font size="2">Atenciosamente</font></p>
	<p align="justify">&nbsp;</p>
	</td>
</tr>
<tr>
	<td width="100%">
	<table border="0" cellpadding="0" width="100%" cellspacing="0">
	<tr>
<%if day(now())=1 then dia="1º" else dia=day(now())%>
		<td width="50%" valign="top">
		<p><font size="2">Osasco,&nbsp;<%=dia & " de " & monthname(month(now())) & " de " & year(now()) %></font></p>
<%
if session("tetoassinatura")<>"" then
%>
		<img src="../images/assinatura.jpg" height="96" border="0" alt="">
<%
else
%>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p><font size="2">_____________________________________<br>
<%
end if
%>		
		</font></p>
        </td>
<%if teste=1 then %>
		<td width="50%" valign="top">&nbsp;
		<div align="center">
		<center>
		<table border="0" cellpadding="0" width="240" cellspacing="0">
		<tr>
			<td width="1" style="border-left-style: solid; border-left-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240" rowspan="2">
				<p align="center"><b><font size="4" color="#808080">73.063.166/0001-20</font></b></td>
			<td width="1" style="border-right-style: solid; border-right-width: 3; border-top-style: solid; border-top-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
		</tr>
		<tr>
			<td width="1"></td>
			<td width="1"></td>
		</tr>
		<tr>
			<td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td>
			<td width="1"></td>
		</tr>
		<tr>
			<td width="1"></td>
			<td width="240">
				<p align="center"><b><font color="#808080">FUNDAÇÃO INSTITUTO DE<br>
				ENSINO PARA OSASCO</font></b></td>
			<td width="1"></td>
		</tr>
		<tr>
			<td width="1"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="240"></td>
			<td width="1"></td>
		</tr>
		<tr>
			<td width="1">&nbsp;</td>
			<td width="240" rowspan="2">
				<p align="center"><font color="#808080">Rua Narciso Sturlini, 883<br>
				Jd. Umuarama - CEP 06018-903<br>
				OSASCO - SP</font></td>
			<td width="1"></td>
		</tr>
		<tr>
			<td width="1" style="border-left-style: solid; border-left-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
			<td width="1" style="border-right-style: solid; border-right-width: 3; border-bottom-style: solid; border-bottom-width: 3"><img border="0" src="../images/branco.gif" width=10 height=10></td>
		</tr>
		</table>
		</center>
		</div>
		<p>&nbsp;
<%end if%>
		</td>
	</tr>
	</table>
	</td>
</tr>
<center>
<tr>
	<td width="100%">&nbsp;</td>
</tr>
<tr>
	<td width="100%">
	<p><b><u>Termo de Compromisso</u></b>
	<p align="justify">Comprometo-me a notificar em tempo hábil, sempre que ocorrer
	quaisquer alterações nas informações prestadas, quais sejam, meu
	desligamento, salário de contribuição, etc., ficando esta
	Instituição de Ensino Superior isenta de qualquer responsabilidade
	quanto à possíveis consequências.</p>
	<p>Data: _____/_____/_______</p>
	<p>&nbsp;</p>
	<p>_______________________________________________<br>
	<%=rs("nome")%> - <%=rs("chapa")%>
	</td>
</tr>
<tr>
	<td>Para IES: <%=session("tetoempresa")%></td>
</tr>
<tr>
	<td>&nbsp;</td>
</tr>
<tr><td class=campo>
	<b><font face="Arial Narrow">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</font></b>
	<br>R. Narciso Sturlini, 883 - Osasco - SP - CEP 06018-903 - Fone: (011) 3681-6000<%if teste=0 then response.write " - C.N.P.J. 73.063.166/0001-20" %>
	<br>Av. Franz Voegeli, 300 - Osasco - SP - CEP 06020-190 - Fone: (011) 3651-9999<%if teste=0 then response.write " - C.N.P.J. 73.063.166/0003-92" %>
<%if teste=0 then%>
	<br>Av. Franz Voegeli, 1743 - Osasco - SP - CEP 06020-190 - Fone: (011) 3654-0655 - C.N.P.J. 73.063.166/0004-73
<%end if%>
	<br>Caixa Postal - ACF - Jd. Ipê - nº 2530 - Osasco - SP - CEP 06053-990
	</td>
</tr>
</table>
</center>
</div>

<%
' inserir no controle de emissao
sql="INSERT INTO rhcontroletetofieo ( chapa, data_emissao, empresa, cnpj, usuario ) " & _
"SELECT '" & rs("chapa") & "', getdate(), '" & session("tetoempresa") & "', '" & session("tetocnpj") & "', " & _
"'" & session("usuariomaster") & "'"
conexao.execute sql

rs.close
set rs=nothing

elseif temp=2 then
session("cartateto")="L"
%>
<!-- mostrar funcionarios e as contribuições -->
<table border="1" cellpadding="0" width="550" cellspacing="0">
<tr>
	<td class=titulo>&nbsp;Chapa</td>
	<td class=titulo>&nbsp;Nome</td>
	<td class=titulo>&nbsp;Situacao</td>
</tr>
<%

rs.movefirst
do while not rs.eof

sqlinss="SELECT TOP 6 CHAPA, ANOCOMP, MESCOMP, NROPERIODO, BASEINSS, [INSS]+[INSSFERIAS] AS TINSS " & _
"FROM corporerm.dbo.PFPERFF " & _
"WHERE CHAPA='" & rs("chapa") & "' AND BASEINSS<>0 and nroperiodo<>10 " & _
"ORDER BY ANOCOMP DESC , MESCOMP DESC "
rsi.Open sqlinss, ,adOpenStatic, adLockReadOnly
if rsi.recordcount>0 then baseinss=cdbl(rsi("baseinss")) else baseinss=0
sql2="select salario, codsindicato from corporerm.dbo.pfunc where chapa='" & rs("chapa") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
salario=cdbl(rs2("salario"))
if rs2("codsindicato")="03" then salario=salario*1.225
rs2.close
sqlgrades="declare @chapa varchar(5) set @chapa='" & rs("chapa") & "' " & _
"select z.chapa, f.nome, f.codsecao, sal=sum( ceiling(convert(float,aulas*valoraula*1.225) + convert(float,valoraula*noturno*0.2))   ), tipo='Graduação' " & _
"from ( select f.chapa, f.instrucaomec, f.codsituacao, tab_instr, f.codnivelsal, " & _
"f.grauinstrucao, tab_ref, a.coddoc, a.sal, a.aulas, a.noturno, c.valoraula from dc_professor f, " & _
"(SELECT g.chapa1, gc.coddoc, gc.curso, gc.sal, " & _
"aulas=sum(case when juntar=1 then 0 else case when extra=1 then 0 else case when demons=1 then 0 else ta end end end)*4.5,  " & _
"gc.adnot, Sum(g.adnot)*4.5 AS noturno " & _
"FROM g2ch as g INNER JOIN g2cursoeve as gc ON gc.coddoc=g.coddoc " & _
"WHERE g.deletada=0 AND g.ativo in (1,0) and g.inicio<=getdate() AND g.demons>=0 " & _
"and getdate() between g.inicio and g.termino and chapa1=@chapa " & _
"GROUP BY g.chapa1, gc.coddoc, gc.curso, gc.CODCCUSTO, gc.sal, gc.adnot) a, " & _
"(SELECT c.evento, c.tabela, f.dt_faixa, t.titulacao, t.nivel, reformulacao, f.valoraula " & _
"FROM (csd_cursos c INNER JOIN csd_titulos t ON c.tabela=t.tabela) INNER JOIN csd_faixas f ON t.faixasalarial=f.faixasalarial " & _
"WHERE getdate() Between ivigencia And fvigencia AND f.dt_faixa='20120801') c " & _
"where f.codsituacao<>'D' and a.chapa1=f.chapa collate database_default " & _
"and a.sal=c.evento and c.nivel=f.codnivelsal collate database_default " & _
"and c.titulacao=f.tab_instr collate database_default " & _
"and c.reformulacao=f.tab_ref collate database_Default " & _
") z, corporerm.dbo.pfunc f " & _
"where z.chapa=f.chapa collate database_default " & _
"group by z.chapa, f.nome, f.codsecao "
'response.write sqlgrades
rs2.Open sqlgrades , ,adOpenStatic, adLockReadOnly
if rs2.bof=true then salgrade=0 else salgrade=cdbl(rs2("sal"))
rs2.close
%>
<tr>
	<td class=campo>&nbsp;<%=rs("chapa")%></td>
	<td class=campo>&nbsp;
<%if (baseinss>valorteto or (salario>valorteto or salgrade>valorteto)) or session("usuariomaster")<>"02379" then%>
<a href="cartateto.asp?codigo=<%=rs("chapa")%>">
<%end if%>
<%=rs("nome")%>
</a>
</td>
	<td class=campo>&nbsp;<%=rs("codsituacao")%></td>
</tr>
<tr><td class="campop" colspan=3 align="center">
	<font color="green"><b>Ultimo Salario: <%=formatnumber(salario,2)%></b>
	<br><font color="blue"><b>Salário pela grade atual: <%=formatnumber(salgrade*1.225,2)%></b>
</td></tr>
<tr><td class=campo colspan=3><table border=0 celpadding=2><tr>
	<td class=campo align="center" width=50>Ano</td>
	<td class=campo align="center" width=50>Mês</td>
	<td class=campo align="center" width=70>Base</td>
	<td class=campo align="center" width=70>INSS</td>
</tr>
<%
if rsi.recordcount>0 then
rsi.movefirst
do while not rsi.eof
%>
<tr>
	<td class=campo><%=rsi("anocomp")%></td>
	<td class=campo align="center"><%=rsi("mescomp")%></td>
	<td class=campo align="right"><%=formatnumber(rsi("baseinss"),2)%></td>
	<td class=campo align="right"><%=formatnumber(rsi("tinss"),2)%></td>
</tr>
<%
rsi.movenext
loop
else
response.write "<tr><td colspan='4'>Sem Bases de INSS no cadastro.</td></tr>"
end if
rsi.close
%>
</table></td></tr>
<%
rs.movenext
loop
%>

</table>
<p><font color='red' size='3'>Só emitir carta de teto para quem ganha acima de R$ <%=formatnumber(valorteto,2)%>
<%
rs.close
set rs=nothing
end if ' temps
conexao.close
set conexao=nothing
%>
</body>
</html>