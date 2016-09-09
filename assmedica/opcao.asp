<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a85")="N" or session("a85")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Opção Intermédica</title>
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
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request("codigo")<>"" or request.form<>"" then
	if request.form="" then temp=request("codigo")
	if request("codigo")="" then temp=request.form("chapa")
	if isnumeric(temp) then
		info=1
		temp=numzero(temp,5)
		sqlb="AND f.CHAPA='" & temp & "' "
	else
		info=2
		sqlb="AND f.nome like '%" & temp & "%' "
	end if

	sqla="SELECT f.NOME, f.CODSITUACAO, f.CHAPA, f.DATAADMISSAO, f.CODSECAO, f.codsindicato, s.DESCRICAO AS Secao, p.dtnascimento, f.salario " & _
	"FROM corporerm.dbo.PFUNC AS f, corporerm.dbo.PSECAO s, corporerm.dbo.PPESSOA p " & _
	"WHERE f.CODSECAO = s.CODIGO and p.codigo=f.codpessoa "

	sql1=sqla & sqlb
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	session("chapa")=rs("chapa")
	temp=0
	if rs.recordcount>1 then temp=2
else
	temp=1
end if

if temp=1 then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário (administrativo ou professor)
<form method="POST" action="opcao.asp">
  <p style="margin-top: 0; margin-bottom: 0">Chapa/Nome <input type="text" name="chapa" size="20" class="form_box" value="<%=session("chapa")%>">
  <input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>
<p><b><font color="#FF0000">Atenção para os períodos para mudança de plano:</font></b><p>
<table border="1" bordercolor="#CCCCCC" cellpadding="7" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td colspan="1" class=grupo>Prazos</td>
	<td class=grupo>Inclusão</td>
	<td class=grupo>Após o prazo</td>
	<td class=grupo>Mudança</td>
	<td class=grupo>Exclusão</td>
</tr>
<tr>
	<td class=campo>Unimed (Administrativos)</td>
	<td class=campo rowspan=2>•Na admissão<br>•até 30 dias nascimento<br>•até 30 dias casamento</td>
	<td class=campo>•Não poderá ser mais incluido</td>
	<td class=campo>•Agosto de cada ano/Promoção</td>
	<td class=campo>•Dependente só com 21 anos.<br>•Esposa só na separação.</td>
</tr>
<tr>
	<td class=campo>Intermédica(Professores)</td>
	<td class=campo>•Não poderá ser mais incluido</td>
	<td class=campo>•entre Setembro/Outubro</td>
	<td class=campo>•Dependente só com 21 anos.<br>•Esposa só na separação.</td>
</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellpadding="5" cellspacing="0" style="border-collapse: collapse">
<tr><td class=grupo colspan=3>Valores dos Planos</td></tr>
<tr><td valign=top>
<!-- Medial -->
<table border="1" bordercolor="#CCCCCC" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo colspan=2>Caixa Seguros</td></tr>
<tr><td class=fundo>Plano   </td><td class=fundo align="right">Valor </td></tr>
<tr><td class=campo>Fundamental Enfermaria</td><td class=campo align="right">197,55</td></tr>
<tr><td class=campo>Fundamental Apto.     </td><td class=campo align="right">246,30</td></tr>
<tr><td class=campo>Vital Enfermaria      </td><td class=campo align="right">265,08</td></tr>
<tr><td class=campo>Vital Apto.           </td><td class=campo align="right">329,52</td></tr>
<tr><td class=campo>Melhor Apto.          </td><td class=campo align="right">457,41</td></tr>
</table>
<!-- final Medial -->
</td><td valign=top>
<!-- Intermédica -->
<table border="1" bordercolor="#CCCCCC" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo colspan=2>Intermédica</td></tr>
<tr><td class=fundo rowspan=1>Plano</td>
	<td class=fundo rowspan=1>Valor</td>
	</tr>
<tr><td class=campo>Extra         </td><td class=campo align="center">89,38</td>
</tr>
<tr><td class=campo>Executivo     </td><td class=campo align="center">127,71</td>
</tr>
<tr><td class=campo>Executivo Plus</td><td class=campo align="center">217,14</td>
</tr>
<tr><td class=campo>Master</td><td class=campo align="center">-</td>
</tr>
</table>
<!-- final Intermédica -->
</td></tr>
</table>

<%
elseif temp=0 then
'if request.form<>"" then
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
idade=int((now()-rs("dtnascimento"))/365.25)
sqlempresa="select CHAPA, ASSMEDICA from corporerm.dbo.PFCOMPL where CHAPA='" & rs("chapa") & "' "
rs2.Open sqlempresa, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
	empresa=rs2("assmedica") ': response.write empresa
else
	if rs("codsindicato")="03" or rs("dataadmissao")>cdate(dateserial(2014,6,1))  then
		empresa="I" 
	elseif int(now())=>cdate(dateserial(2014,10,29)) then 
		empresa="BS" 
	else 
		empresa="C"
	end if
end if
rs2.close

select case empresa
	case "I"
		dt_inicio="01/10/2003"
		operadora="Intermédica Sistema de Saúde"
		planogratis="EXTRA"
		anterior="SAMCIL"
		valor=formatnumber(89.38,2)
		if rs("codsindicato")="03" then tipo="PROFESSOR" else tipo="ADMINISTRATIVO"
		if rs("codsindicato")="03" then clausula="cláusula 49 item 5" else clausula="cláusula 40 item 5"
		copar=cdbl(8.94)
	case "C"
		dt_inicio="01/02/2016"
		operadora="Caixa Seguros"
		planogratis="Fundamental Enfermaria"
		anterior="Bradesco Saúde"
		valor=formatnumber(197.55,2)
		tipo="ADMINISTRATIVO"
		clausula="cláusula 40 item 5"
		copar=cdbl(19.75)
	case "U"
		dt_inicio="01/08/2010"
		operadora="Unimed Seguros"
		planogratis="BÁSICO"
		anterior="MEDIAL"
		valor=formatnumber(166.67,2)
		tipo="ADMINISTRATIVO"
		clausula="cláusula 40 item 5"
		copar=cdbl(16.67)
	case "BS"
		dt_inicio="01/11/2014"
		operadora="Bradesco Saúde"
		planogratis="Perfil Enfermaria"
		anterior="UNIMED"
		valor=formatnumber(179.55,2)
		tipo="ADMINISTRATIVO"
		clausula="cláusula 40 item 5"
		copar=cdbl(17.96)
end select
inicial=0
%>
<table border="0" bordercolor="#000000" cellpadding="4" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td colspan=2 class=titulop align="center" style="border: 1px solid #000000"><b>OPÇÕES AO PLANO DE ASSISTÊNCIA MÉDICO-HOSPITALAR</b></td>
</tr>
<tr>
	<td class="campop">
	Eu,&nbsp;<%=rs("nome") %> (<%=idade%>), venho por livre e espontânea vontade, manifestar minhas opções em relação aos
	planos de assistência médica e/ou odontológica:</td>
</tr>
<tr>
	<td class=campo>
	<img src="../images/arrow.gif" width="13" height="10" border="0" alt="">________________________ ser incluido no plano de assistência médica conforme opção abaixo:
	<br><font style="font-size:8px">&nbsp;&nbsp;(Escreva "Desejo" ou "Não desejo")
	</td>
<tr>
	<td class=campo valign=top align="center">
<%
if now()<dateserial(2010,8,1) then datainicio=dateserial(2010,8,1) else datainicio=now()
sqla="SELECT empresa, plano FROM assmed_mudanca " & _
"WHERE chapa='" & rs("chapa") & "' AND '" & dtaccess(datainicio) & "' Between [ivigencia] And [fvigencia] and empresa in ('I','C','BS') "
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then plano=rs3("plano") else plano=""
rs3.close

if cdbl(rs("salario"))<4000 then limitep=4 else limitep=5
if rs("chapa")="00162" then limitep=5
if empresa="I" then codpar="IP"
if empresa="U" then codpar="UC"
if empresa="BS" then codpar="BP"
if empresa="C" then codpar="CP"
sqlpar="select valor from assmed_planos where seq=2 and codigo='" & codpar & "'"
rs3.Open sqlpar, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then copar=rs3("valor") else copar=0.00
rs3.close

sqlplano="SELECT codigo, seq, plano, valor, reembolso FROM (select * from assmed_planos where (codigo='I' and seq<=3) or (codigo='U' and seq<=" & limitep & ") or (codigo='BS' and seq<=" & limitep & ") or codigo='C'  ) a " & _
"WHERE codigo='" & empresa & "' AND plano Not Like 'agr%' ORDER BY seq "
rs3.Open sqlplano, ,adOpenStatic, adLockReadOnly
%>	
	<table border="1" bordercolor="#000000" cellpadding="1" width="600" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td align="center" width=22>Opção</font></td>
		<td align="center">Planos</font></td>
		<td align="center">Custo</font></td>
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular<br>(c/co-partic.)</td>
<%if empresa="C" or empresa="BS" or (empresa="I" and tipo="ADMINISTRATIVO") then%>		
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular/Dependente<br>(s/co-partic.)</td>
<%else%>
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular (s/co-partic.)</td>
		<td align="center" class=campo>Desconto por&nbsp;<br>Dependente (s/co-partic.)</td>
<%end if%>		
	</tr>
<%
if rs3.recordcount>0 then 
	rs3.movefirst
	do while not rs3.eof
	if plano=rs3("plano") then campof="fundo" else campof="campop"
	if empresa="I" then fator=1.00 else fator=1
	if empresa="I" and tipo="PROFESSOR" then desconto3=rs3("valor")*fator else desconto3=rs3("reembolso")*fator
	%>
		<tr>
			<td class=<%=campof%> align="center"><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
			<td class=<%=campof%>>&nbsp;<%=rs3("plano")%></font></td>
			<td class=<%=campof%> align="center"><%=formatnumber(rs3("valor")*fator,2)%></td>
			<td class=<%=campof%> align="center"><%=formatnumber(rs3("reembolso")*fator+copar,2)%></td>
	<%if empresa="C" or empresa="U" or (empresa="I" and tipo="ADMINISTRATIVO") then%>
			<td class=<%=campof%> align="center"><%=formatnumber(desconto3,2)%></td>
	<%else%>
			<td class=<%=campof%> align="center"><%=formatnumber(rs3("reembolso")*fator,2)%></td>
			<td class=<%=campof%> align="center"><%=formatnumber(desconto3,2)%></td>
	<%end if%>
		</tr>
	<%
	rs3.movenext
	loop
else
%>
		<tr>
			<td class=<%=campof%> align="center"><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
			<td class=<%=campof%></td>
			<td class=<%=campof%> align="center"></td>
			<td class=<%=campof%> align="center"></td>
			<td class=<%=campof%> align="center"></td>
			<td class=<%=campof%> align="center"></td>
		</tr>
<%
end if
%>
	</table>
<%
rs3.close
%>
	</td>
</tr>
<%
if empresa<>"I" or tipo="ADMINISTRATIVO" then
%>
<tr>
	<td class=campo>
	<img src="../images/arrow.gif" width="13" height="10" border="0" alt="">________________________ ser incluido no plano de assistência odontológica conforme opção abaixo:
	<br><font style="font-size:8px">&nbsp;&nbsp;(Escreva "Desejo" ou "Não desejo")
	</td>
<tr>
<tr>
	<td class=campo valign=top align="center">
<%
sqla="SELECT empresa, plano FROM assmed_mudanca " & _
"WHERE chapa='" & rs("chapa") & "' AND '" & dtaccess(datainicio) & "' Between [ivigencia] And [fvigencia] and empresa in ('O','V') "
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then plano=rs3("plano") else plano=""
if rs3.recordcount>0 then empresa=rs3("empresa") else empresa=""
rs3.close

if cdbl(rs("salario"))<3000 then limitep=3 else limitep=4
sqlplano="SELECT codigo, seq, plano, valor, reembolso FROM (select * from assmed_planos where (codigo='V' and seq<=" & limitep & ")) a " & _
"WHERE codigo='V' ORDER BY seq "
rs3.Open sqlplano, ,adOpenStatic, adLockReadOnly
%>	
	<table border="1" bordercolor="#000000" cellpadding="1" width="520" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td align="center" width=22>Opção</td>
		<td align="center">Planos</td>
		<td align="center">Custo</td>
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular ou dependente</td>
	</tr>
<%
rs3.movefirst
do while not rs3.eof
if plano=rs3("plano") then campof="fundo" else campof="campop"
%>
	<tr>
		<td class=<%=campof%> align="center"><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
		<td class=<%=campof%>>&nbsp;<%=rs3("plano")%></font></td>
		<td class=<%=campof%> align="center"><%=formatnumber(rs3("valor"),2)%></td>
		<td class=<%=campof%> align="center"><%=formatnumber(rs3("reembolso"),2)%></td>
	</tr>
<%
rs3.movenext
loop
%>
	</table>
<%
rs3.close
set rs3=nothing
%>
	</td>
</tr>

<%
end if ' final empresa <>I (sem odonto para professores)
%>

<tr>
	<td class=campo>Desejo também&nbsp;incluir os meus dependentes legais (esposa, filhos
	até 21 anos) abaixo relacionados:</td>
</tr>
<tr>
	<td class=campo valign=top align="center">
<%
sql2="select distinct d.chapa, nome=dependente, dtnascimento=nascimento, parentesco, " & _
"datediff(yy,nascimento,getdate()) AS idade, mae, cpf, saude=s.nrodepend, odonto=o.nrodepend " & _
"from assmed_dep d " & _
"left join (select chapa, nrodepend from assmed_dep_mudanca where empresa IN ('U','I') ) s on s.chapa=d.chapa and s.nrodepend=d.nrodepend " & _
"left join (select chapa, nrodepend from assmed_dep_mudanca where empresa='V') o on o.chapa=d.chapa and o.nrodepend=d.nrodepend " & _
"where d.chapa='" & session("chapa") & "' and (s.nrodepend is not null or o.nrodepend is not null) " & _
""
'response.write sql2
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
%>
	<table border="1" bordercolor="#CCCCCC" cellpadding="1" width="600" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td align="center"><font size="1">Opção<br>Saúde</td>
		<td align="center"><font size="1">Opção<br>Odonto</td>
		<td align="center"><font size="1">Nome do Dependente</td>
		<td align="center"><font size="1">Grau de&nbsp;<br> Parentesco</td>
		<td align="center"><font size="1">Data de&nbsp;<br> Nascimento</td>
		<td align="center"><font size="1">Idade</font></td>
	</tr>
<%
totaldep=0
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
if ( left(rs2("parentesco"),4)="Filh" or left(rs2("parentesco"),6)="Entead" ) and rs2("idade")>20 then
else
if rs2("saude")>0 then bolas="../images/bolax.gif" else bolas="../images/bola.gif"
if rs2("odonto")>0 then bolao="../images/bolax.gif" else bolao="../images/bola.gif"
%>
	<tr>
		<td align="center" rowspan="2"><img src=<%=bolas%> width="22" height="22" border="0" alt=""></td>
		<td align="center" rowspan="2"><img src=<%=bolao%> width="22" height="22" border="0" alt=""></td>
		<td><font size="2">&nbsp;<%=rs2("nome")%></td>
		<td><font size="2">&nbsp;<%=rs2("parentesco")%></td>
		<td><font size="2">&nbsp;<%=rs2("dtnascimento")%></td>
		<td><font size="2">&nbsp;<%=rs2("idade")%></td>
	</tr>
	<tr>
		<td class="campor" colspan="2">Nome da mãe do dependente:<br> <font size=2><%=rs2("mae")%>&nbsp;</td>
		<td class="campor" colspan="2">CPF do dependente:<br> <font size=2><%=rs2("cpf")%>&nbsp;</td>
	</tr>
<%
totaldep=totaldep+1
end if
rs2.movenext
loop
rs2.close
set rs2=nothing
end if

if totaldep<=3 then linhasdep=1 '3
if totaldep<=2 then linhasdep=2 '2
if totaldep<=1 then linhasdep=3 '1
if totaldep>3 then linhasdep=0 '4

for a=0 to linhasdep-1 '3
%>
	<tr>
		<td align="center" rowspan="2" height=25><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
		<td align="center" rowspan="2" height=25><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td class="campor" colspan="2" height=30>Nome da mãe do dependente:<br><font size=2>&nbsp;</td>
		<td class="campor" colspan="2">CPF do dependente:<br><font size=2>&nbsp;</td>
	</tr>
<%next%>
	</table>
	</td>
</tr>
<tr>
	<td class=campo>
	<img src="../images/arrow.gif" width="13" height="10" border="0" alt="">_________________ o desconto mensal em salário de R$ ____________, através da folha de pagamento:
	<br><font style="font-size:8px">&nbsp;&nbsp;(Escreva "Autorizo")
	<br><font style="font-size:11px"> - da diferença de valores entre o plano de saúde "<%=planogratis%>" a que tenho direito conforme convenção coletiva e
	os valores de plano diferenciado escolhido por mim e para meus dependentes.
	</td>
<tr>
<tr>
	<td class=campo>
	<img src="../images/arrow.gif" width="13" height="10" border="0" alt="">_________________ o desconto mensal em salário de R$ ____________, através da folha de pagamento:
	<br><font style="font-size:8px">&nbsp;&nbsp;(Escreva "Autorizo")
	<br><font style="font-size:11px"> - dos valores do plano odontológico escolhido por mim e por meus dependentes, se for o caso.
	</td>
<tr>
<tr>
	<td class=campo>
	<img src="../images/arrow.gif" width="13" height="10" border="0" alt="">_________________ contribuir no regime de co-participação;
	<br><font style="font-size:8px">&nbsp;&nbsp;(Escreva "Desejo" ou "Não desejo")
	<br><font style="font-size:11px"> - para permanecer no plano após demissão sem justa causa pelo período de 1/3 do tempo de contribuição, limitado a 2 anos.
	</td>
<tr>
<tr>
	<td class=campo>
	<img src="../images/arrow.gif" width="13" height="10" border="0" alt="">____________________ ciente de que:
	<br><font style="font-size:8px">&nbsp;&nbsp;(Escreva "Estou")
	<br><font style="font-size:11px"> - os valores acima serão reajustados de acordo com os critérios definidos em contrato entre a FIEO e a operadora de saúde;
	<br><font style="font-size:11px"> - a mudança para um plano superior NÃO poderá ser revertida;
	<br><font style="font-size:11px"> - a inclusão de novos dependentes após 30 dias do evento implicará em carência;
	<br><font style="font-size:11px"> - a exclusão dos dependentes atuais só poderá ser feita na perda de dependência (exemplo: maioridade para filhos, separação para conjuges etc).
	</td>
<tr>

<tr>
	<td class=campo><p align="justify">
	
<p><font size="2">Osasco,&nbsp;<%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %></font></p>
<p><font size="2">_____________________________________<br>
<%=rs("chapa")%>  - <%=rs("nome") %></font></p>
	</td>
</tr>
</table>
<%
rs.close
set rs=nothing
elseif temp=2 then
%>

<table border="1" bordercolor="#CCCCCC" cellpadding="0" width="550" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>&nbsp;Chapa</td>
	<td class=titulo>&nbsp;Nome</td>
	<td class=titulo>&nbsp;Situacao</td>
</tr>
<%
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class=campo>&nbsp;<%=rs("chapa")%></td>
	<td class=campo>&nbsp;<a href="opcao.asp?codigo=<%=rs("chapa")%>"><%=rs("nome")%></a></td>
	<td class=campo>&nbsp;<%=rs("codsituacao")%></td>
</tr>
<%
rs.movenext
loop
%>
</table>
<%
rs.close
set rs=nothing
end if ' temps

conexao.close
set conexao=nothing
%>
</body>
</html>