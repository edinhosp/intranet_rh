<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")="N" or session("a72")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Rotinas Mensal</title>
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

sql="select top 1 * from est_parametro "
rs.Open sql, ,adOpenStatic, adLockReadOnly
ano=rs("ano")
mes=rs("mes")
descricao=rs("descricao")
inicio=rs("inicio")
fim=rs("fim")
limite=rs("limite")
rs.close
%>
<p class=titulo>Rotina Mensal para importação e cálculo de apontamento</p>
<form method="POST" action="rotinamensal.asp" name="form">
<table border="0" bordercolor=black cellpadding="2" cellspacing="1" style="border-collapse: collapse" width=300>
<tr>
	<td class=titulo height=35 colspan=2>
	Período atual: <%=inicio%> a <%=fim%>
	</td>
</tr>
<tr>
	<td class=titulo height=35>1. Criar datas do período</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R1">
</tr>
<tr>
	<td class=titulo height=35>2. Importar marcações Chronus</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R2">
</tr>
<tr>
	<td class=titulo height=35>3. Atribuir Horários</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R3">
</tr>
<tr>
	<td class=titulo height=35>4. Cálculo</td>
	<td class=fundo align="center"><input type="submit" value="Executar" class="button" name="R4">
</tr>
</table>

</form>
<hr>
<%
'*************************** R1 inicio ***************************
if request.form("R1")<>"" then
	rot1a=0:sql1="select chapa from est_histhor group by chapa":rs.Open sql1, ,adOpenStatic, adLockReadOnly
	do while not rs.eof
	for a=inicio to fim
		sql2="select chapa, data from est_batfun where chapa='" & rs("chapa") & "' and data='" & dtaccess(a) & "'"
		rs2.Open sql2, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount=0 then
			rot1a=rot1a+1:sql3="insert into est_batfun (chapa,data) select '" & rs("chapa") & "','" & dtaccess(a) & "' ":conexao.execute sql3
		end if
		rs2.close
	next
	rs.movenext:loop
	response.write "Rotina 1 - criando chapas e datas: " & rot1a & " ocorrências.<br>"
	'rs.close
	response.write "<p style=''><b>Rotina Finalizada</b>"
end if
'*************************** R1 final  ***************************

'*************************** R2 inicio ***************************
if request.form("R2")<>"" then
	sql1="select a.chapa, a.data from corporerm.dbo.abatfun a, est_batfun e where e.chapa=a.chapa collate database_default and e.data=a.data " & _
	"and a.data between '" & dtaccess(inicio) & "' and '" & dtaccess(fim) & "' " & _
	"group by a.chapa, a.data "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly	
	response.write "Rotina 2: " & rs.recordcount & " ocorrências."
	if rs.recordcount>0 then
		do while not rs.eof
			posicao=0
			sql3="update est_batfun set marc1=null, marc2=null, marc3=null, marc4=null, marc5=null, marc6=null where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "'"
			conexao.execute sql3
			sql2="select batida from corporerm.dbo.abatfun where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' order by batida"
			rs2.Open sql2, ,adOpenStatic, adLockReadOnly
			do while not rs2.eof
				posicao=posicao+1
				sql3="update est_batfun set marc" & posicao & "=" & rs2("batida") & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "'"
				conexao.execute sql3
			rs2.movenext
			loop
			rs2.close
		rs.movenext
		loop
	end if
	rs.close
	sql1="select chapa, data, marc1, marc2, marc3, marc4, marc5, marc6, htrab " & _
	"from est_batfun " & _
	"where data between '" & dtaccess(inicio) & "' and '" & dtaccess(fim) & "' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	do while not rs.eof
		htrab=0
		marc1=rs("marc1"):marc2=rs("marc2"):marc3=rs("marc3"):marc4=rs("marc4"):marc5=rs("marc5"):marc6=rs("marc6")
		if isnumeric(marc1)=true and isnumeric(marc2)=true then tot1=marc2-marc1 else tot1=0
		if isnumeric(marc3)=true and isnumeric(marc4)=true then tot2=marc4-marc3 else tot2=0
		if isnumeric(marc5)=true and isnumeric(marc6)=true then tot3=marc6-marc5 else tot3=0
		htrab=tot1+tot2+tot3
		sql2="update est_batfun set htrab=" & htrab & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
		conexao.execute sql2
	rs.movenext
	loop
	response.write "<p style=''><b>Rotina Finalizada</b>"
end if
'*************************** R2 final  ***************************

'*************************** R3 inicio ***************************
if request.form("R3")<>"" then

	sql2="select chapa from est_batfun where data between '" & dtaccess(inicio) & "' and '" & dtaccess(fim) & "' group by chapa "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly	
	do while not rs2.eof

	'verificando qual o ultimo codigo de hora e dia antes do periodo
		sql1="select top 1 chapa, codigo, dtmudanca, dia from est_histhor where chapa='" & rs2("chapa") & "' and dtmudanca<='" & dtaccess(inicio) & "' order by dtmudanca desc "
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		naotem=0
		if rs.recordcount>0 then
			indice=rs("dia"):adia=rs("dia")
			acodigo=rs("codigo")
			adtmudanca=rs("dtmudanca")
			naotem=1
		end if
		rs.close
		if naotem=0 then
			sql1="select top 1 chapa, codigo, dtmudanca, dia from est_histhor where chapa='" & rs2("chapa") & "' and dtmudanca>='" & dtaccess(inicio) & "' order by dtmudanca desc "
			rs.Open sql1, ,adOpenStatic, adLockReadOnly
			if rs.recordcount>0 then
				indice=rs("dia"):adia=rs("dia")
				acodigo=rs("codigo")
				adtmudanca=rs("dtmudanca")
				naotem=1
			end if
			rs.close
		end if
	response.write "X"
	'verificando o indice maior de dia
		sql2="select max(dia) as loop from est_cadhorario_marcacoes where codigo='" & acodigo & "'"
		rs.Open sql2, ,adOpenStatic, adLockReadOnly
		maxindice=rs("loop"):rs.close
	'preparando o loop
		dias=datediff("d",adtmudanca,inicio)
		novadata=adtmudanca
		for a=1 to dias
			novadata=novadata+1
			indice=indice+1
			if indice>maxindice then indice=1
		next
	'preparando o periodo de apontamento
		for a=novadata to fim

	sql1="update est_batfun set descanso=null,feriado=null where chapa='" & rs2("chapa") & "' and data='" & dtaccess(a) & "' "
	conexao.execute sql1

			if a=novadata then indice=indice else indice=indice+1
			sql1="select top 1 codigo, dia, dtmudanca from est_histhor where dtmudanca='" & dtaccess(a) & "' and chapa='" & rs2("chapa") & "' "
			rs.Open sql1, ,adOpenStatic, adLockReadOnly
			if rs.recordcount>0 then 
				indice=rs("dia")-0
				if acodigo<>rs("codigo") then
					sqlb="select max(dia) as loop from est_cadhorario_marcacoes where codigo='" & rs("codigo") & "'"
					Set rsd=conexao.Execute (sqlb, , adCmdText)
					maxindice=rsd("loop"):rsd.close
					acodigo=rs("codigo")
				end if
			else 
				indice=indice
			end if
			rs.close
			if indice>maxindice then indice=1
			'definindo codigo e dia a ser feito
			sql2="update est_batfun set codigo='" & acodigo & "', dia=" & indice & " where chapa='" & rs2("chapa") & "' and data='" & dtaccess(a) & "' "
			conexao.execute sql2
			'definindo os horarios
			sql1="select ent1, sai1, ent2, sai2, jorn, comp, [desc], intflex from est_cadhorario_marcacoes where codigo='" & acodigo & "' and dia=" & indice
			rs.Open sql1, ,adOpenStatic, adLockReadOnly
			ent1=rs("ent1"):sai1=rs("sai1"):ent2=rs("ent2"):sai2=rs("sai2"):jorn=rs("jorn"):comp=rs("comp"):desc=rs("desc"):intflex=rs("intflex")
			rs.close
			sql2="update est_batfun set hor1="&ent1&", hor2="&sai1&", hor3="&ent2&", hor4="&sai2&", base="&jorn & " where chapa='" & rs2("chapa") & "' and data='" & dtaccess(a) & "' "
			conexao.execute sql2
			if desc=-1 then 
				sql2="update est_batfun set descanso=1440 where chapa='" & rs2("chapa") & "' and data='" & dtaccess(a) & "' "
				conexao.execute sql2
			end if
			'feriado
			sql1="select nome from corporerm.dbo.gferiado where diaferiado='" & dtaccess(a) & "' "
			rs.Open sql1, ,adOpenStatic, adLockReadOnly
			if rs.recordcount>0 then 
				feriado=rs("nome")
				sql2="update est_batfun set hor1=0,hor2=0,hor3=0,hor4=0,base=0,feriado=1440 where chapa='" & rs2("chapa") & "' and data='" & dtaccess(a) & "' "
				conexao.execute sql2
			end if
			rs.close
			'
		next
	rs2.movenext
	loop
	rs2.close
	response.write "<br><b>Etapa Concluída!</b>"
end if
'*************************** R3 final  ***************************

'*************************** R4 inicio ***************************
if request.form("R4")<>"" then
	atraso=0
	sql2="select * from est_batfun where data between '" & dtaccess(inicio) & "' and '" & dtaccess(fim) & "' order by chapa, data "
	rs.Open sql2, ,adOpenStatic, adLockReadOnly	
	incr=int(rs.recordcount/10)
	do while not rs.eof
	if (rs.absoluteposition/incr)-int(rs.absoluteposition/incr)=0 then response.write "X"
	'limpar calculos anteriores
	sql1="update est_batfun set ajust3=null,ajust5=null,ajust6=null,falta=0,atraso=0,extra=0 where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
	conexao.execute sql1
	if rs("ajust1")>0 then t1="ajust1=ajust1" else t1="ajust1=null"
	if rs("ajust2")>0 then t2="ajust2=ajust2" else t2="ajust2=null"
	if rs("ajust4")>0 then t4="ajust4=ajust4" else t4="ajust4=null"
	t0=t1 & "," & t2 & "," & t4
	sql1="update est_batfun set " & t0 & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
	conexao.execute sql1
	'sem marcações ou marcações incompletas - vai gerar falta
	if rs("htrab")=0 and rs("base")>0 then 
		falta=rs("base"):trab=0
		sql1="update est_batfun set falta=" & falta & ", trab=" & trab & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
		conexao.execute sql1
	else
		falta=0
		sql1="update est_batfun set falta=0 where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
		conexao.execute sql1
	end if
	' atrasos
	atraso=0
	if rs("marc1") > rs("hor1")+limite then atraso=atraso + rs("marc1")-rs("hor1")
	if rs("hor4")>0 then
		if rs("marc4") < rs("hor4")-limite then atraso=atraso + rs("hor4")-rs("marc4")
	else
		if rs("marc2") < rs("hor2")-limite then atraso=atraso + rs("hor2")-rs("marc2")
	end if
	if falta>0 then atraso=0
	' intervalo
	if rs("hor3")>0 and rs("hor4")>0 then
		atrasoalmoco=0
		intervalo=rs("hor3")-rs("hor2")
		if rs("marc3")>0 and rs("marc4")>0 then intfeito=rs("marc3")-rs("marc2")
		if intfeito>(intervalo+limite) then 
			atraso=atraso + (intfeito-intervalo)
			atrasoalmoco=intfeito-intervalo
		end if
	end if
	if rs("base")=0 then atraso=0
	sql1="update est_batfun set atraso=" & atraso & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
	conexao.execute sql1

	' extras ou ajuste
	extra=0
	ajust1=0:ajust2=0:ajust4=0
	if rs("base")>0 then
		if rs("marc1") < rs("hor1")-limite then	
			extra=extra + (rs("hor1")-rs("marc1"))
			randomize timer
			ajust1=(rs("hor1")-limite) + int(rnd*9)+1
		end if
		if rs("hor4")>0 then
			if rs("marc4") > rs("hor4")+limite then 
				extra=extra + (rs("marc4")-rs("hor4"))
				randomize timer
				ajust4=(rs("hor4")-limite) + int(rnd*9)+1
			end if	
			'if rs("marc3") < rs("hor3")-limite then 
			'	extra=extra + (rs("hor3")-rs("marc3"))
			'	randomize timer
			'	ajust3=(rs("hor3")-limite) + int(rnd*9)+1
			'end if	
		else
			if rs("marc2") > rs("hor2")+limite then 
				extra=extra + (rs("marc2")-rs("hor2"))
				randomize timer
				ajust2=(rs("hor2")-limite) + int(rnd*9)+1
			end if
		end if
		if falta>0 then extra=0
		sql1="update est_batfun set extra=" & extra & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
		conexao.execute sql1
		if rs("travar")=0 then	
			sql1="update est_batfun set ajust1=" & ajust1 & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
			if ajust1<>0 then conexao.execute sql1
			sql1="update est_batfun set ajust2=" & ajust2 & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
			if ajust2<>0 then conexao.execute sql1
			sql1="update est_batfun set ajust4=" & ajust4 & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
			if ajust4<>0 then conexao.execute sql1
		end if 'travar

		htrab=0
		marc1=rs("marc1"):marc2=rs("marc2"):marc3=rs("marc3"):marc4=rs("marc4"):marc5=rs("marc5"):marc6=rs("marc6")
		if ajust1<>0 then marc1=ajust1
		if ajust2<>0 then marc2=ajust2
		if ajust4<>0 then marc4=ajust4
		if intervalo>0 then marc2=rs("hor2")
		if intervalo>0 then marc3=rs("hor3")
		if isnumeric(marc1)=true and isnumeric(marc2)=true then tot1=marc2-marc1 else tot1=0
		if isnumeric(marc3)=true and isnumeric(marc4)=true then tot2=marc4-marc3 else tot2=0
		if isnumeric(marc5)=true and isnumeric(marc6)=true then tot3=marc6-marc5 else tot3=0
		htrab=tot1+tot2+tot3-atrasoalmoco
		sql2="update est_batfun set htrab=" & htrab & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
		conexao.execute sql2
	else 'base=0
		htrab=0
		marc1=rs("marc1"):marc2=rs("marc2"):marc3=rs("marc3"):marc4=rs("marc4"):marc5=rs("marc5"):marc6=rs("marc6")
		if ajust1<>0 then marc1=ajust1
		if ajust2<>0 then marc2=ajust2
		if ajust4<>0 then marc4=ajust4
		if isnumeric(marc1)=true and isnumeric(marc2)=true then tot1=marc2-marc1 else tot1=0
		if isnumeric(marc3)=true and isnumeric(marc4)=true then tot2=marc4-marc3 else tot2=0
		if isnumeric(marc5)=true and isnumeric(marc6)=true then tot3=marc6-marc5 else tot3=0
		htrab=tot1+tot2+tot3
		sql2="update est_batfun set htrab=0, extra=" & htrab & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
		conexao.execute sql2
	end if	
	
	
	if rs("descanso")>0 or rs("feriado")>0 then
		if rs("marc1")>0 then 
			extra=rs("marc2")-rs("marc1")
			sql1="update est_batfun set extra=" & extra & " where chapa='" & rs("chapa") & "' and data='" & dtaccess(rs("data")) & "' "
			conexao.execute sql1
		end if
	end if

	rs.movenext
	loop
	response.write "<br><b>Etapa Concluída!</b>"
end if
'*************************** R4 final  ***************************

teste=0
if teste=1 then
'*************** inicio teste **********************
if request.form<>"" then
response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
response.write "<tr>"
for a=0 to rs.fields.count-1
	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "</td>"
next
response.write "</tr>"
if rs.recordcount>0 then rs.movefirst
do while not rs.eof 
response.write "<tr>"
for a= 0 to rs.fields.count-1
	response.write "<td class=""campor"" nowrap>" & rs.fields(a) & "</td>"
next
response.write "</tr>"
rs.movenext
loop
response.write "</table>"
response.write "<p>"
end if
'*************** fim teste **********************
end if 'teste=1
%>

</body>
</html>
<%

set rs=nothing
set rs2=nothing
conexao.close
set conexao=nothing
%>