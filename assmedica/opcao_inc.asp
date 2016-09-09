<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Opção Intermédica</title>
</head>
<body>
<%
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

chapa=session("assmed_adm")
sql1="SELECT f.NOME, f.CODSITUACAO, f.CHAPA, f.DATAADMISSAO, f.CODSECAO, f.codsindicato, s.DESCRICAO AS Secao, p.dtnascimento, f.salario " & _
"FROM corporerm.dbo.PFUNC AS f, corporerm.dbo.PSECAO s, corporerm.dbo.PPESSOA p " & _
"WHERE f.CODSECAO = s.CODIGO and p.codigo=f.codpessoa and f.chapa='" & chapa & "'"
rs1.Open sql1, ,adOpenStatic, adLockReadOnly
session("chapa")=rs1("chapa")
session("chapanome")=rs1("nome")
idade=int((now()-rs1("dtnascimento"))/365.25)
if rs1("codsindicato")="03" then empresa="I" else empresa="U"
select case empresa
	case "I"
		dt_inicio="01/10/2003"
		operadora="Intermédica Sistema de Saúde"
		planogratis="EXTRA"
		anterior="SAMCIL"
		valor=formatnumber(49.47,2)
		tipo="PROFESSOR"
		clausula="cláusula 49 item 5"
		copar=cdbl(4.95)
	case "M"
		dt_inicio="19/05/2003"
		operadora="Medial Saúde"
		planogratis="CLASSICO I"
		anterior="AMESP"
		valor=formatnumber(86.10,2)
		tipo="ADMINISTRATIVO"
		clausula="cláusula 40 item 5"
		copar=cdbl(8.61)
	case "U"
		dt_inicio="01/08/2010"
		operadora="Unimed Seguros"
		planogratis="BÁSICO"
		anterior="MEDIAL"
		valor=formatnumber(142.45,2)
		tipo="ADMINISTRATIVO"
		clausula="cláusula 40 item 5"
		copar=cdbl(14.25)
end select
inicial=0
%>
<table border="0" bordercolor="#000000" cellpadding="4" width="690" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td colspan=2 class=titulop align="center" style="border: 1px solid #000000"><b>OPÇÕES AO PLANO DE ASSISTÊNCIA MÉDICO-HOSPITALAR</b></td>
</tr>
<tr>
	<td class="campop">
	Eu,&nbsp;<%=rs1("nome") %> (<%=idade%>), venho por livre e espontânea vontade, manifestar minhas opções em relação aos
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
"WHERE chapa='" & rs1("chapa") & "' AND '" & dtaccess(datainicio) & "' Between [ivigencia] And [fvigencia] and empresa in ('I','U') "
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then plano=rs3("plano") else plano=""
rs3.close

if cdbl(rs1("salario"))<3000 then limitep=3 else limitep=4
if rs1("chapa")="00162" then limitep=4
if empresa="I" then codpar="IP"
if empresa="U" then codpar="UC"
sqlpar="select valor from assmed_planos where seq=2 and codigo='" & codpar & "'"
rs3.Open sqlpar, ,adOpenStatic, adLockReadOnly
copar=rs3("valor")
rs3.close

sqlplano="SELECT codigo, seq, plano, valor, reembolso FROM (select * from assmed_planos where (codigo='I' and seq<=3) or (codigo='U' and seq<=" & limitep & ")) a " & _
"WHERE codigo='" & empresa & "' AND plano Not Like 'agr%' ORDER BY seq "
rs3.Open sqlplano, ,adOpenStatic, adLockReadOnly
%>	
	<table border="1" bordercolor="#000000" cellpadding="1" width="600" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td align="center" width=22>Opção</font></td>
		<td align="center">Planos</font></td>
		<td align="center">Custo</font></td>
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular<br>(c/co-partic.)</td>
<%if empresa="U" then%>		
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular/Dependente<br>(s/co-partic.)</td>
<%else%>
		<td align="center" class=campo>Desconto por&nbsp;<br>Titular (s/co-partic.)</td>
		<td align="center" class=campo>Desconto por&nbsp;<br>Dependente (s/co-partic.)</td>
<%end if%>		
	</tr>
<%
rs3.movefirst
do while not rs3.eof
if plano=rs3("plano") then campof="fundo" else campof="campop"
if empresa="I" then fator=1.00 else fator=1
if empresa="I" then desconto3=rs3("valor")*fator else desconto3=rs3("reembolso")*fator
%>
	<tr>
		<td class=<%=campof%> align="center"><img src="../images/bola.gif" width="22" height="22" border="0" alt=""></td>
		<td class=<%=campof%>>&nbsp;<%=rs3("plano")%></font></td>
		<td class=<%=campof%> align="center"><%=formatnumber(rs3("valor")*fator,2)%></td>
		<td class=<%=campof%> align="center"><%=formatnumber(rs3("reembolso")*fator+copar,2)%></td>
<%if empresa="U" then%>
		<td class=<%=campof%> align="center"><%=formatnumber(desconto3,2)%></td>
<%else%>
		<td class=<%=campof%> align="center"><%=formatnumber(rs3("reembolso")*fator,2)%></td>
		<td class=<%=campof%> align="center"><%=formatnumber(desconto3,2)%></td>
<%end if%>
	</tr>
<%
rs3.movenext
loop
%>
	</table>
<%
rs3.close
%>
	</td>
</tr>
<%
if empresa<>"I" then
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
"WHERE chapa='" & rs1("chapa") & "' AND '" & dtaccess(datainicio) & "' Between [ivigencia] And [fvigencia] and empresa in ('O','V') "
rs3.Open sqla, ,adOpenStatic, adLockReadOnly
if rs3.recordcount>0 then plano=rs3("plano") else plano=""
if rs3.recordcount>0 then empresa=rs3("empresa") else empresa=""
rs3.close

if cdbl(rs1("salario"))<3000 then limitep=3 else limitep=4
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
"left join (select chapa, nrodepend from assmed_dep_mudanca where empresa='U') s on s.chapa=d.chapa and s.nrodepend=d.nrodepend " & _
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
<%=rs1("chapa")%>  - <%=rs1("nome") %></font></p>
	</td>
</tr>
</table>
<%
rs1.close
set rs1=nothing

%>
</body>
</html>