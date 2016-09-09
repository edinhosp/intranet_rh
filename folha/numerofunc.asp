<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a45")="N" or session("a45")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Distribuição de Funcionários na Instituição</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, tgl(4,6), tl(4), tg(6), descricao(4)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

'*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'*************** fim teste **********************

dataatual=cdate(formatdatetime(now(),2))
datapesquisa=cdate(formatdatetime(request.form("dataquery"),2))
if request.form("dataquery")<>"" then
	datacampo=datapesquisa
	numero=2
else
	datacampo=dataatual
	numero=1
end if

if numero=1 then
	'administrativos
	sql1 ="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato<>'03' and codtipo='N' and left(codsecao,2)='01' and codsituacao in ('A','F','Z')"
	sql2 ="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato<>'03' and codtipo='N' and left(codsecao,2)='01' and codsituacao in ('L','E','I','M','O','P','T','U')"
	sql3 ="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato<>'03' and codtipo='N' and left(codsecao,2)='02' and codsituacao in ('A','F','Z')"
	sql4 ="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato<>'03' and codtipo='N' and left(codsecao,2)='02' and codsituacao in ('L','E','I','M','O','P','T','U')"
	sql5 ="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato<>'03' and codtipo='N' and left(codsecao,2)='03' and codsituacao in ('A','F','Z')"
	sql6 ="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato<>'03' and codtipo='N' and left(codsecao,2)='03' and codsituacao in ('L','E','I','M','O','P','T','U')"
	sql7 ="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato<>'03' and codtipo='N' and left(codsecao,2)='04' and codsituacao in ('A','F','Z')"
	sql8 ="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato<>'03' and codtipo='N' and left(codsecao,2)='04' and codsituacao in ('L','E','I','M','O','P','T','U')"
	'professores
	sql9 ="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato='03' and codtipo='N' and left(codsecao,2)='01' and codsituacao in ('A','F','Z')"
	sql10="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato='03' and codtipo='N' and left(codsecao,2)='01' and codsituacao in ('L','E','I','M','O','P','T','U')"
	sql11="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato='03' and codtipo='N' and left(codsecao,2)='02' and codsituacao in ('A','F','Z')"
	sql12="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato='03' and codtipo='N' and left(codsecao,2)='02' and codsituacao in ('L','E','I','M','O','P','T','U')"
	sql13="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato='03' and codtipo='N' and left(codsecao,2)='03' and codsituacao in ('A','F','Z')"
	sql14="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato='03' and codtipo='N' and left(codsecao,2)='03' and codsituacao in ('L','E','I','M','O','P','T','U')"
	sql15="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato='03' and codtipo='N' and left(codsecao,2)='04' and codsituacao in ('A','F','Z')"
	sql16="select count(chapa) as total from corporerm.dbo.pfunc where codsindicato='03' and codtipo='N' and left(codsecao,2)='04' and codsituacao in ('L','E','I','M','O','P','T','U')"
	'estagiarios
	sql17="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='01' and codsituacao in ('A','F','Z') and codsecao<>'01.1.999'"
	sql18="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='01' and codsituacao in ('L','E','I','M','O','P','T','U') and codsecao<>'01.1.999'"
	sql19="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='02' and codsituacao in ('A','F','Z') and codsecao<>'02.1.999'"
	sql20="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='02' and codsituacao in ('L','E','I','M','O','P','T','U') and codsecao<>'02.1.999'"
	sql21="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='03' and codsituacao in ('A','F','Z') and codsecao<>'03.1.999'"
	sql22="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='03' and codsituacao in ('L','E','I','M','O','P','T','U') and codsecao<>'03.1.999'"
	sql23="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='04' and codsituacao in ('A','F','Z') and codsecao<>'04.1.999'"
	sql24="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='04' and codsituacao in ('L','E','I','M','O','P','T','U') and codsecao<>'04.1.999'"
	'portador
	sql25="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='01' and codsituacao in ('A','F','Z')"
	sql26="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='01' and codsituacao in ('L','E','I','M','O','P','T','U')"
	sql27="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='02' and codsituacao in ('A','F','Z')"
	sql28="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='02' and codsituacao in ('L','E','I','M','O','P','T','U')"
	sql29="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='03' and codsituacao in ('A','F','Z')"
	sql30="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='03' and codsituacao in ('L','E','I','M','O','P','T','U')"
	sql31="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='04' and codsituacao in ('A','F','Z')"
	sql32="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='04' and codsituacao in ('L','E','I','M','O','P','T','U')"
else
	sql1 ="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codsindicato<>'03' and f.codtipo='N' and left(codsecao,2)='01' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql2 ="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codsindicato<>'03' and f.codtipo='N' and left(codsecao,2)='01' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa"
	sql3 ="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codsindicato<>'03' and f.codtipo='N' and left(codsecao,2)='02' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql4 ="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codsindicato<>'03' and f.codtipo='N' and left(codsecao,2)='02' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa"
	sql5 ="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codsindicato<>'03' and f.codtipo='N' and left(codsecao,2)='03' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql6 ="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codsindicato<>'03' and f.codtipo='N' and left(codsecao,2)='03' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa"
	sql7 ="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codsindicato<>'03' and f.codtipo='N' and left(codsecao,2)='04' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql8 ="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codsindicato<>'03' and f.codtipo='N' and left(codsecao,2)='04' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa"

	sql9 ="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codsindicato='03' and f.codtipo='N' and left(codsecao,2)='01' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql10="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codsindicato='03' and f.codtipo='N' and left(codsecao,2)='01' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa"
	sql11="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codsindicato='03' and f.codtipo='N' and left(codsecao,2)='02' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql12="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codsindicato='03' and f.codtipo='N' and left(codsecao,2)='02' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa"
	sql13="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codsindicato='03' and f.codtipo='N' and left(codsecao,2)='03' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql14="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codsindicato='03' and f.codtipo='N' and left(codsecao,2)='03' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa"
	sql15="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codsindicato='03' and f.codtipo='N' and left(codsecao,2)='04' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql16="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codsindicato='03' and f.codtipo='N' and left(codsecao,2)='04' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa"

	sql17="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codtipo='T' and left(codsecao,2)='01' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') and codsecao<>'01.1.999'"
	sql18="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codtipo='T' and left(codsecao,2)='01' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa and codsecao<>'01.1.999'"
	sql19="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codtipo='T' and left(codsecao,2)='02' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') and codsecao<>'02.1.999'"
	sql20="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codtipo='T' and left(codsecao,2)='02' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa and codsecao<>'02.1.999'"
	sql21="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codtipo='T' and left(codsecao,2)='03' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') and codsecao<>'03.1.999'"
	sql22="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codtipo='T' and left(codsecao,2)='03' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa and codsecao<>'03.1.999'"
	sql23="select count(f.chapa) as total from corporerm.dbo.pfunc f where f.codtipo='T' and left(codsecao,2)='04' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') and codsecao<>'04.1.999'"
	sql24="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.pfhstaft a where f.codtipo='T' and left(codsecao,2)='04' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa and codsecao<>'04.1.999'"
	
'	sql17="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='01' and codsituacao in ('A','F','Z') and codsecao<>'01.1.999'"
'	sql18="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='01' and codsituacao in ('L','E','I','M','O','P','T','U') and codsecao<>'01.1.999'"
'	sql19="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='02' and codsituacao in ('A','F','Z') and codsecao<>'02.1.999'"
'	sql20="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='02' and codsituacao in ('L','E','I','M','O','P','T','U') and codsecao<>'02.1.999'"
'	sql21="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='03' and codsituacao in ('A','F','Z') and codsecao<>'03.1.999'"
'	sql22="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='03' and codsituacao in ('L','E','I','M','O','P','T','U') and codsecao<>'03.1.999'"
'	sql23="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='04' and codsituacao in ('A','F','Z') and codsecao<>'04.1.999'"
'	sql24="select count(chapa) as total from corporerm.dbo.pfunc where codtipo='T' and left(codsecao,2)='04' and codsituacao in ('L','E','I','M','O','P','T','U') and codsecao<>'04.1.999'"

	sql25="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='01' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql26="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.pfhstaft a where a.chapa=f.chapa and f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='01' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa "
	sql27="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='02' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql28="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.pfhstaft a where a.chapa=f.chapa and f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='02' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa "
	sql29="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='03' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql30="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.pfhstaft a where a.chapa=f.chapa and f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='03' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa "
	sql31="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p where f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='04' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') --and f.chapa<'10000' "
	sql32="select count(f.chapa) as total from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.pfhstaft a where a.chapa=f.chapa and f.codpessoa=p.codigo and (deficientefisico=1 or deficienteauditivo=1 or deficientefala=1 or deficientevisual=1 or deficientemental=1) and codtipo='N' and left(codsecao,2)='04' and (dataadmissao<='" & dtaccess(datacampo) & "' and (dtdesligamento is null or dtdesligamento>'" & dtaccess(datacampo) & "')) and (DTTRANSFERENCIA is null or DTTRANSFERENCIA<='" & dtaccess(datacampo) & "') /*and f.chapa<'10000'*/ and (a.dtinicio<='" & dtaccess(datacampo) & "' and (a.dtfinal is null or a.dtfinal>'" & dtaccess(datacampo) & "')) and a.chapa=f.chapa "
end if

'*************administrativos
rs.Open sql1, ,adOpenStatic, adLockReadOnly : n_ad_a=rs("total") : rs.close
rs.Open sql2, ,adOpenStatic, adLockReadOnly : n_ad_l=rs("total") : rs.close:if numero=2 then n_ad_a=n_ad_a-n_ad_l
n_ad_t = n_ad_a + n_ad_l

rs.Open sql3, ,adOpenStatic, adLockReadOnly : b_ad_a=rs("total") : rs.close
rs.Open sql4, ,adOpenStatic, adLockReadOnly : b_ad_l=rs("total") : rs.close:if numero=2 then b_ad_a=b_ad_a-b_ad_l
b_ad_t = b_ad_a + b_ad_l

rs.Open sql5, ,adOpenStatic, adLockReadOnly : v_ad_a=rs("total") : rs.close
rs.Open sql6, ,adOpenStatic, adLockReadOnly : v_ad_l=rs("total") : rs.close:if numero=2 then v_ad_a=v_ad_a-v_ad_l
v_ad_t = v_ad_a + v_ad_l

rs.Open sql7, ,adOpenStatic, adLockReadOnly : j_ad_a=rs("total") : rs.close
rs.Open sql8, ,adOpenStatic, adLockReadOnly : j_ad_l=rs("total") : rs.close:if numero=2 then j_ad_a=j_ad_a-j_ad_l
j_ad_t = j_ad_a + j_ad_l

t_ad_a = n_ad_a + b_ad_a + v_ad_a + j_ad_a
t_ad_l = n_ad_l + b_ad_l + v_ad_l + j_ad_l
t_ad_t = n_ad_t + b_ad_t + v_ad_t + j_ad_t

'*************professores
rs.Open sql9, ,adOpenStatic, adLockReadOnly : n_pr_a=rs("total") : rs.close
rs.Open sql10, ,adOpenStatic, adLockReadOnly : n_pr_l=rs("total") : rs.close:if numero=2 then n_pr_a=n_pr_a-n_pr_l
n_pr_t = n_pr_a + n_pr_l

rs.Open sql11, ,adOpenStatic, adLockReadOnly : b_pr_a=rs("total") : rs.close
rs.Open sql12, ,adOpenStatic, adLockReadOnly : b_pr_l=rs("total") : rs.close:if numero=2 then b_pr_a=b_pr_a-b_pr_l
b_pr_t = b_pr_a + b_pr_l

rs.Open sql13, ,adOpenStatic, adLockReadOnly : v_pr_a=rs("total") : rs.close
rs.Open sql14, ,adOpenStatic, adLockReadOnly : v_pr_l=rs("total") : rs.close:if numero=2 then v_pr_a=v_pr_a-v_pr_l
v_pr_t = v_pr_a + v_pr_l

rs.Open sql15, ,adOpenStatic, adLockReadOnly : j_pr_a=rs("total") : rs.close
rs.Open sql16, ,adOpenStatic, adLockReadOnly : j_pr_l=rs("total") : rs.close:if numero=2 then j_pr_a=j_pr_a-j_pr_l
j_pr_t = j_pr_a + j_pr_l

t_pr_a = n_pr_a + b_pr_a + v_pr_a + j_pr_a
t_pr_l = n_pr_l + b_pr_l + v_pr_l + j_pr_l
t_pr_t = n_pr_t + b_pr_t + v_pr_t + j_pr_t

'********************Estagiarios
rs.Open sql17, ,adOpenStatic, adLockReadOnly : n_es_a=rs("total") : rs.close
rs.Open sql18, ,adOpenStatic, adLockReadOnly : n_es_l=rs("total") : rs.close
n_es_t = n_es_a + n_es_l

rs.Open sql19, ,adOpenStatic, adLockReadOnly : b_es_a=rs("total") : rs.close
rs.Open sql20, ,adOpenStatic, adLockReadOnly : b_es_l=rs("total") : rs.close
b_es_t = b_es_a + b_es_l

rs.Open sql21, ,adOpenStatic, adLockReadOnly : v_es_a=rs("total") : rs.close
rs.Open sql22, ,adOpenStatic, adLockReadOnly : v_es_l=rs("total") : rs.close
v_es_t = v_es_a + v_es_l

rs.Open sql23, ,adOpenStatic, adLockReadOnly : j_es_a=rs("total") : rs.close
rs.Open sql24, ,adOpenStatic, adLockReadOnly : j_es_l=rs("total") : rs.close
j_es_t = j_es_a + j_es_l

t_es_a = n_es_a + b_es_a + v_es_a + j_es_a
t_es_l = n_es_l + b_es_l + v_es_l + j_es_l
t_es_t = n_es_t + b_es_t + v_es_t + j_es_t

'********************Portadores de Deficiência
rs.Open sql25, ,adOpenStatic, adLockReadOnly : n_pd_a=rs("total") : rs.close
rs.Open sql26, ,adOpenStatic, adLockReadOnly : n_pd_l=rs("total") : rs.close:if numero=2 then n_pd_a=n_pd_a-n_pd_l
n_pd_t = n_pd_a + n_pd_l

rs.Open sql27, ,adOpenStatic, adLockReadOnly : b_pd_a=rs("total") : rs.close
rs.Open sql28, ,adOpenStatic, adLockReadOnly : b_pd_l=rs("total") : rs.close:if numero=2 then b_pd_a=b_pd_a-b_pd_l
b_pd_t = b_pd_a + b_pd_l

rs.Open sql29, ,adOpenStatic, adLockReadOnly : v_pd_a=rs("total") : rs.close
rs.Open sql30, ,adOpenStatic, adLockReadOnly : v_pd_l=rs("total") : rs.close:if numero=2 then v_pd_a=v_pd_a-v_pd_l
v_pd_t = v_pd_a + v_pd_l

rs.Open sql31, ,adOpenStatic, adLockReadOnly : j_pd_a=rs("total") : rs.close
rs.Open sql32, ,adOpenStatic, adLockReadOnly : j_pd_l=rs("total") : rs.close:if numero=2 then j_pd_a=j_pd_a-j_pd_l
j_pd_t = j_pd_a + j_pd_l

t_pd_a = n_pd_a + b_pd_a + v_pd_a + j_pd_a
t_pd_l = n_pd_l + b_pd_l + v_pd_l + j_pd_l
t_pd_t = n_pd_t + b_pd_t + v_pd_t + j_pd_t

%>
<form method="POST" name="form" action="numerofunc.asp">
<p class=titulo>Distribuição de Funcionários na Instituição em <input type=text name=dataquery size=10 class=subli2 value="<%=datacampo%>">
<table border="1" bordercolor=#000000 cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulo align="center" rowspan=2 style="border-right:2 solid #000000"></td>
	<td class="campoa" align="center" colspan=3 style="border-right:2 solid #000000">Campus Narciso</td>
	<td class="campot" align="center" colspan=3 style="border-right:2 solid #000000">Campus Brás</td>
	<td class="campol" align="center" colspan=3 style="border-right:2 solid #000000">Campus Vila Yara</td>
	<td class="campov" align="center" colspan=3 style="border-right:2 solid #000000">Campus Jd.Wilson</td>
	<td class=fundo align="center" colspan=3 style="border-right:2 solid #000000">Total Geral</td>
</tr>
<tr>
	<td class="campoa" align="center" >Ativos</td>
	<td class="campoa" align="center" >Lic.</td>
	<td class="campoa" align="center" style="border-right:2 solid #000000"><b>Total</td>
	<td class="campot" align="center" >Ativos</td>
	<td class="campot" align="center" >Lic.</td>
	<td class="campot" align="center" style="border-right:2 solid #000000"><b>Total</td>
	<td class="campol" align="center" >Ativos</td>
	<td class="campol" align="center" >Lic.</td>
	<td class="campol" align="center" style="border-right:2 solid #000000"><b>Total</td>
	<td class="campov" align="center" >Ativos</td>
	<td class="campov" align="center" >Lic.</td>
	<td class="campov" align="center" style="border-right:2 solid #000000"><b>Total</td>
	<td class=fundo align="center" >Ativos</td>
	<td class=fundo align="center" >Lic.</td>
	<td class=fundo align="center" style="border-right:2 solid #000000"><b>Total</td>
</tr>
<tr>
	<td class=campo style="border-right:2 solid #000000">Administrativos</td>
	<td class="campoa" align="center" ><%=n_ad_a%></td>
	<td class="campoa" align="center" ><%=n_ad_l%></td>
	<td class="campoa" align="center" style="border-right:2 solid #000000"><b><%=n_ad_t%></td>
	<td class="campot" align="center" ><%=b_ad_a%></td>
	<td class="campot" align="center" ><%=b_ad_l%></td>
	<td class="campot" align="center" style="border-right:2 solid #000000"><b><%=b_ad_t%></td>
	<td class="campol" align="center" ><%=v_ad_a%> (*)</td>
	<td class="campol" align="center" ><%=v_ad_l%></td>
	<td class="campol" align="center" style="border-right:2 solid #000000"><b><%=v_ad_t%></td>
	<td class="campov" align="center" ><%=j_ad_a%></td>
	<td class="campov" align="center" ><%=j_ad_l%></td>
	<td class="campov" align="center" style="border-right:2 solid #000000"><b><%=j_ad_t%></td>
	<td class=fundo align="center" ><%=t_ad_a%></td>
	<td class=fundo align="center" ><%=t_ad_l%></td>
	<td class=fundo align="center" style="border-right:2 solid #000000"><b><%=t_ad_t%></td>
</tr>
<tr>
	<td class=campo style="border-right:2 solid #000000">Professores</td>
	<td class="campoa" align="center" ><%=n_pr_a%></td>
	<td class="campoa" align="center" ><%=n_pr_l%></td>
	<td class="campoa" align="center" style="border-right:2 solid #000000"><b><%=n_pr_t%></td>
	<td class="campot" align="center" ><%=b_pr_a%></td>
	<td class="campot" align="center" ><%=b_pr_l%></td>
	<td class="campot" align="center" style="border-right:2 solid #000000"><b><%=b_pr_t%></td>
	<td class="campol" align="center" ><%=v_pr_a%></td>
	<td class="campol" align="center" ><%=v_pr_l%></td>
	<td class="campol" align="center" style="border-right:2 solid #000000"><b><%=v_pr_t%></td>
	<td class="campov" align="center" ><%=j_pr_a%></td>
	<td class="campov" align="center" ><%=j_pr_l%></td>
	<td class="campov" align="center" style="border-right:2 solid #000000"><b><%=j_pr_t%></td>
	<td class=fundo align="center" ><%=t_pr_a%></td>
	<td class=fundo align="center" ><%=t_pr_l%></td>
	<td class=fundo align="center" style="border-right:2 solid #000000"><b><%=t_pr_t%></td>
</tr>
<%
n_1_a = n_ad_a + n_pr_a
n_1_l = n_ad_l + n_pr_l
n_1_t = n_ad_t + n_pr_t
b_1_a = b_ad_a + b_pr_a
b_1_l = b_ad_l + b_pr_l
b_1_t = b_ad_t + b_pr_t
v_1_a = v_ad_a + v_pr_a
v_1_l = v_ad_l + v_pr_l
v_1_t = v_ad_t + v_pr_t
j_1_a = j_ad_a + j_pr_a
j_1_l = j_ad_l + j_pr_l
j_1_t = j_ad_t + j_pr_t
t_1_a = t_ad_a + t_pr_a
t_1_l = t_ad_l + t_pr_l
t_1_t = t_ad_t + t_pr_t

%>
<tr>
	<td class=campo style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000"><b>Total Funcionários</td>
	<td class="campoa" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000"><b><%=n_1_a%></td>
	<td class="campoa" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000"><b><%=n_1_l%></td>
	<td class="campoa" align="center" style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000"><b><%=n_1_t%></td>
	<td class="campot" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000"><b><%=b_1_a%></td>
	<td class="campot" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000"><b><%=b_1_l%></td>
	<td class="campot" align="center" style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000"><b><%=b_1_t%></td>
	<td class="campol" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000"><b><%=v_1_a%></td>
	<td class="campol" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000"><b><%=v_1_l%></td>
	<td class="campol" align="center" style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000"><b><%=v_1_t%></td>
	<td class="campov" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000"><b><%=j_1_a%></td>
	<td class="campov" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000"><b><%=j_1_l%></td>
	<td class="campov" align="center" style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000"><b><%=j_1_t%></td>
	<td class=fundo align="center" style="border-top:3 double #000000;border-bottom:3 double #000000"><b><%=t_1_a%></td>
	<td class=fundo align="center" style="border-top:3 double #000000;border-bottom:3 double #000000"><b><%=t_1_l%></td>
	<td class=fundo align="center" style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000"><b><%=t_1_t%></td>
</tr>
<tr>
	<td class=campo style="border-right:2 solid #000000">Estagiários</td>
	<td class="campoa" align="center" ><%=n_es_a%></td>
	<td class="campoa" align="center" ><%=n_es_l%></td>
	<td class="campoa" align="center" style="border-right:2 solid #000000"><b><%=n_es_t%></td>
	<td class="campot" align="center" ><%=b_es_a%></td>
	<td class="campot" align="center" ><%=b_es_l%></td>
	<td class="campot" align="center" style="border-right:2 solid #000000"><b><%=b_es_t%></td>
	<td class="campol" align="center" ><%=v_es_a%></td>
	<td class="campol" align="center" ><%=v_es_l%></td>
	<td class="campol" align="center" style="border-right:2 solid #000000"><b><%=v_es_t%></td>
	<td class="campov" align="center" ><%=j_es_a%></td>
	<td class="campov" align="center" ><%=j_es_l%></td>
	<td class="campov" align="center" style="border-right:2 solid #000000"><b><%=j_es_t%></td>
	<td class=fundo align="center" ><%=t_es_a%></td>
	<td class=fundo align="center" ><%=t_es_l%></td>
	<td class=fundo align="center" style="border-right:2 solid #000000"><b><%=t_es_t%></td>
</tr>
<%
n_2_a = n_1_a + n_es_a
n_2_l = n_1_l + n_es_l
n_2_t = n_1_t + n_es_t
b_2_a = b_1_a + b_es_a
b_2_l = b_1_l + b_es_l
b_2_t = b_1_t + b_es_t
v_2_a = v_1_a + v_es_a
v_2_l = v_1_l + v_es_l
v_2_t = v_1_t + v_es_t
j_2_a = j_1_a + j_es_a
j_2_l = j_1_l + j_es_l
j_2_t = j_1_t + j_es_t
t_2_a = t_1_a + t_es_a
t_2_l = t_1_l + t_es_l
t_2_t = t_1_t + t_es_t

%>
<tr>
	<td class=campo style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000">Total c/Estagiários</td>
	<td class="campoa" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000" ><%=n_2_a%></td>
	<td class="campoa" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000" ><%=n_2_l%></td>
	<td class="campoa" align="center" style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000"><%=n_2_t%></td>
	<td class="campot" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000" ><%=b_2_a%></td>
	<td class="campot" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000" ><%=b_2_l%></td>
	<td class="campot" align="center" style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000"><%=b_2_t%></td>
	<td class="campol" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000" ><%=v_2_a%></td>
	<td class="campol" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000" ><%=v_2_l%></td>
	<td class="campol" align="center" style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000"><%=v_2_t%></td>
	<td class="campov" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000" ><%=j_2_a%></td>
	<td class="campov" align="center" style="border-top:3 double #000000;border-bottom:3 double #000000" ><%=j_2_l%></td>
	<td class="campov" align="center" style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000"><%=j_2_t%></td>
	<td class=fundo align="center" style="border-top:3 double #000000;border-bottom:3 double #000000" ><%=t_2_a%></td>
	<td class=fundo align="center" style="border-top:3 double #000000;border-bottom:3 double #000000" ><%=t_2_l%></td>
	<td class=fundo align="center" style="border-right:2 solid #000000;border-top:3 double #000000;border-bottom:3 double #000000"><%=t_2_t%></td>
</tr>
<tr>
	<td class="campor" colspan=16>&nbsp;</td>
</tr>

<tr>
	<td class=campo style="border-right:2 solid #000000">Total P.D.</td>
	<td class="campoa" align="center" ><%=n_pd_a%></td>
	<td class="campoa" align="center" ><%=n_pd_l%></td>
	<td class="campoa" align="center" style="border-right:2 solid #000000"><b><%=n_pd_t%></td>
	<td class="campot" align="center" ><%=b_pd_a%></td>
	<td class="campot" align="center" ><%=b_pd_l%></td>
	<td class="campot" align="center" style="border-right:2 solid #000000"><b><%=b_pd_t%></td>
	<td class="campol" align="center" ><%=v_pd_a%></td>
	<td class="campol" align="center" ><%=v_pd_l%></td>
	<td class="campol" align="center" style="border-right:2 solid #000000"><b><%=v_pd_t%></td>
	<td class="campov" align="center" ><%=j_pd_a%></td>
	<td class="campov" align="center" ><%=j_pd_l%></td>
	<td class="campov" align="center" style="border-right:2 solid #000000"><b><%=j_pd_t%></td>
	<td class=fundo align="center" ><%=t_pd_a%></td>
	<td class=fundo align="center" ><%=t_pd_l%></td>
	<td class=fundo align="center" style="border-right:2 solid #000000"><b><%=t_pd_t%></td>
</tr>


</table>
</form>
<p><font size=1>(1) Cota de P.D. em relação ao Total: <%=int(t_1_t*0.04)%>
<br>(2) Defícit de P.D.: <%=int(t_1_t*0.04)-t_pd_t%>
<br>(*) Incluso no total:
<br>&nbsp;&nbsp;&nbsp;4 Pró-Reitores/Reitor
<br>&nbsp;&nbsp;&nbsp;6 Diretores
<br>&nbsp;&nbsp;&nbsp;1 Diretor/Professor
<%
	pagina=pagina+1
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
'else 'sem registros
%>
<%
'end if 'recordcount

'rs.close
'set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>