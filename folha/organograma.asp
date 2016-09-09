<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a92")="N" or session("a92")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Organograma</title>
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
datacampo=formatdatetime(now(),2)

sqla="select sum(case when codtipo='T' then 1 else 0 end) as est, sum(case when codtipo='N' and codsituacao in ('A','F','Z') then 1 else 0 end) as ativo, " & _
"sum(case when codtipo='N' and codsituacao not in ('A','F','Z') then 1 else 0 end) as afast, count(chapa) as numero, sum(salario+ats) as valor from pfunc_ats where codsecao='"
sqlb="' and codsituacao not in ('D') "

sqla="SELECT Sum(case when f.codtipo='T' then 1 else 0 end) AS est, Sum(case when f.codtipo='N' And f.codsituacao In ('A','F','Z') then 1 else 0 end) AS ativo, " & _
"Sum(case when f.codtipo='N' And f.codsituacao Not In ('A','F','Z') then 1 else 0 end) AS afast, Count(f.chapa) AS numero, " & _
"Sum(case when f.codsituacao In ('A','F','Z') Or f.codtipo='T' then f.salario+f.ats else 0 end) AS valor FROM pfunc_ats f, organograma_pessoas o " & _
"WHERE f.chapa collate database_default=o.chapa and f.codsituacao Not In ('D') AND o.organograma='"
sqlb="' "
sql1=sqla & "COPA-NS" & sqlb 'copa-ns
sql2=sqla & "FINANCEIRO-NS" & sqlb 'financeiro-ns
sql3=sqla & "RECURSOS HUMANOS-VY" & sqlb 'recursos humanos
sql4=sqla & "SUPRIMENTOS-VY" & sqlb 'suprimentos
sql5=sqla & "FINANCEIRO-VY" & sqlb 'financeiro-vy
sql6=sqla & "ARTE/MARKETING-VY" & sqlb 'arte/marketing-vy
sql7=sqla & "GRAFICA-VY" & sqlb 'grafica-vy
sql8=sqla & "SEVEN-SERVICOS DE EVENTOS-VY" & sqlb 'seven-vy
sql9=sqla & "TECNICO-VY" & sqlb 'tecnico-vy
sql10=sqla & "TECNICO-NS" & sqlb 'tecnico-ns
sql11=sqla & "CONTABILIDADE-VY" & sqlb 'contabilidade
sql12=sqla & "INSPETORIA DE ALUNOS-VY" & sqlb 'inspetoria-vy
sql13=sqla & "INSPETORIA DE ALUNOS-NS" & sqlb 'inspetoria-ns
sql14=sqla & "NEGOCIACAO-VY" & sqlb 'negociação
sql15=sqla & "COPA-VY" & sqlb 'copa-vy
sql16=sqla & "PORTARIA-VY" & sqlb 'portaria-v
sql17=sqla & "AUDIO E VIDEO-VY" & sqlb 'audio e video-vy
sql18=sqla & "PORTARIA-NS" & sqlb 'portaria-ns
sql19=sqla & "TELEFONIA-VY" & sqlb 'telefonia-vy
sql20=sqla & "PORTARIA-JW" & sqlb 'portaria-jw
sql21=sqla & "TELEFONIA-NS" & sqlb 'telefonia-ns
sql22=sqla & "MANUTENCAO-VY" & sqlb 'manutenção-vy
sql23=sqla & "MANUTENCAO-NS" & sqlb 'manutenção-ns
sql24=sqla & "SERVICOS GERAIS-VY" & sqlb 'serviços gerais-vy
sql25=sqla & "SERVICOS GERAIS-NS" & sqlb 'serviços gerais-ns
sql26=sqla & "SERVICOS GERAIS-JW" & sqlb 'serviços gerais-jw
sql27=sqla & "BIBLIOTECA-VY" & sqlb 'biblioteca-vy
sql28=sqla & "BIBLIOTECA-NS" & sqlb 'biblioteca-ns
sql29=sqla & "CAEF-VY" & sqlb 'caef-vy
sql30=sqla & "CAEF-NS" & sqlb 'caef-ns
sql31=sqla & "SECRETARIA GERAL-VY" & sqlb 'secretaria geral-vy
sql32=sqla & "SECRETARIA GERAL-NS" & sqlb 'secretaria geral-ns
sql33=sqla & "ARQUIVO GERAL-VY" & sqlb 'arquivo geral-vy
sql34=sqla & "SECRETARIA GERAL POS GRADUACAO-NS" & sqlb 'secretaria geral pós-graduação-ns
sql35=sqla & "OUVIDORIA-VY" & sqlb 'ouvidoria-vy
sql36=sqla & "CENTRAL DE ATENDIMENTO-VY" & sqlb 'central de atendimento-vy
sql37=sqla & "EDIFIEO-VY" & sqlb 'edifieo
sql38=sqla & "DEPTO. JURIDICO" & sqlb 'depto juridico
sql39=sqla & "JUIZADO DE PEQUENAS CAUSAS" & sqlb 'juizado de pequenas causas
sql40=sqla & "OBRA-VY" & sqlb 'obras
sql41=sqla & "PLANEJAMENTO-VY" & sqlb 'planejamento
sql42=sqla & "PROTOCOLO-VY" & sqlb 'recursos humanos-protocolo
sql43=sqla & "MOTORISTA RH-VY" & sqlb 'recursos humanos-motorista
sql44=sqla & "CPD-VY" & sqlb 'cpd-vy
sql45=sqla & "BOLSA DE ESTUDOS-VY" & sqlb 'bolsa-cristiane
sql46=sqla & "PESQUISA-VY" & sqlb 'bolsa-cristiane
sql47=sqla & "INSPETORIA DE ALUNOS-JW" & sqlb 'inspetoria-jw
sql48=sqla & "EFA-VY" & sqlb '
sql49=sqla & "CURSO EFA-VY" & sqlb '
sql50=sqla & "CURSOS ALFA-VY" & sqlb '
sql51=sqla & "SECRETARIA DO CURSO EFA-VY" & sqlb '
sql52=sqla & "CDHO-VY" & sqlb '
sql53=sqla & "CLINICA DE FISIOTERAPIA-JW" & sqlb '
sql54=sqla & "LABORATORIO DE FISIOTERAPIA-VY" & sqlb '
sql55=sqla & "LABORATORIO DE BIOLOGIA-VY" & sqlb '
sql56=sqla & "LABORATORIO DE FISICA-VY" & sqlb '
sql57=sqla & "LABORATORIO DE FOTOGRAFIA-VY" & sqlb '
sql58=sqla & "LABORATORIO DE RADIO-VY" & sqlb '
sql59=sqla & "LABORATORIO DE TV-VY" & sqlb '
sql60=sqla & "LABORATORIO DE QUIMICA-VY" & sqlb '
sql61=sqla & "LABORATORIO DE TURISMO-VY" & sqlb '
sql62=sqla & "RECEPCAO-JW" & sqlb '
sql63=sqla & "CLINICA DE PSICOPEDAGOGIA-JW" & sqlb '
sql64=sqla & "SECRETARIA DE CURSO - BL.MARRON" & sqlb '
sql65=sqla & "SECRETARIA DE CURSO - BL.VERDE" & sqlb '
sql66=sqla & "SECRETARIA DE CURSO - BL.PRATA" & sqlb '
sql67=sqla & "SECRETARIA POS GRADUACAO-VY" & sqlb '
sql68=sqla & "SECRETARIA POS GRADUACAO-NS" & sqlb '
sql69=sqla & "SECRETARIA DO CURSO DE DIREITO-NS" & sqlb '
sql70=sqla & "SECRETARIA DO SAJ-NS" & sqlb '
sql71=sqla & "PRAD-SECRETARIA" & sqlb '
sql72=sqla & "PRAD-ASSESSORIA" & sqlb '
sql73=sqla & "PRAC-SECRETARIA" & sqlb '
sql74=sqla & "PREC-SECRETARIA" & sqlb '
sql75=sqla & "PRDRC-SECRETARIA" & sqlb '
sql76=sqla & "PREC-ASSESSORIA" & sqlb '
sql77=sqla & "SECR. DIR.CURSO MATUTINO" & sqlb '
sql78=sqla & "SECR. DIR.CURSO TECNOLOGIA" & sqlb '
sql79=sqla & "SECR. DIR.CURSO NOTURNO" & sqlb '
sql80=sqla & "SECR. DIR.CURSO DIREITO" & sqlb '
sql81=sqla & "SEGURANCA E MEDICINA DO TRABALHO" & sqlb '
sql82=sqla & "PRAC-ASSESSORIA" & sqlb '

'*************administrativos
rs.Open sql1, ,adOpenStatic, adLockReadOnly : copa1_ativo=rs("ativo") : copa1_afast=rs("afast") : copa1_est=rs("est") : copa1_numero=rs("numero") : copa1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql2, ,adOpenStatic, adLockReadOnly : fin1_ativo=rs("ativo") : fin1_afast=rs("afast") : fin1_est=rs("est") : fin1_numero=rs("numero") : fin1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql3, ,adOpenStatic, adLockReadOnly : rh3_ativo=rs("ativo") : rh3_afast=rs("afast") : rh3_est=rs("est") : rh3_numero=rs("numero") : rh3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql4, ,adOpenStatic, adLockReadOnly : supr3_ativo=rs("ativo") : supr3_afast=rs("afast") : supr3_est=rs("est") : supr3_numero=rs("numero") : supr3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql5, ,adOpenStatic, adLockReadOnly : fin3_ativo=rs("ativo") : fin3_afast=rs("afast") : fin3_est=rs("est") : fin3_numero=rs("numero") : fin3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql6, ,adOpenStatic, adLockReadOnly : arte3_ativo=rs("ativo") : arte3_afast=rs("afast") : arte3_est=rs("est") : arte3_numero=rs("numero") : arte3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql7, ,adOpenStatic, adLockReadOnly : graf3_ativo=rs("ativo") : graf3_afast=rs("afast") : graf3_est=rs("est") : graf3_numero=rs("numero") : graf3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql8, ,adOpenStatic, adLockReadOnly : seven3_ativo=rs("ativo") : seven3_afast=rs("afast") : seven3_est=rs("est") : seven3_numero=rs("numero") : seven3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql9, ,adOpenStatic, adLockReadOnly : tec3_ativo=rs("ativo") : tec3_afast=rs("afast") : tec3_est=rs("est") : tec3_numero=rs("numero") : tec3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql10, ,adOpenStatic, adLockReadOnly : tec1_ativo=rs("ativo") : tec1_afast=rs("afast") : tec1_est=rs("est") : tec1_numero=rs("numero") : tec1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql11, ,adOpenStatic, adLockReadOnly : cont3_ativo=rs("ativo") : cont3_afast=rs("afast") : cont3_est=rs("est") : cont3_numero=rs("numero") : cont3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql12, ,adOpenStatic, adLockReadOnly : insp3_ativo=rs("ativo") : insp3_afast=rs("afast") : insp3_est=rs("est") : insp3_numero=rs("numero") : insp3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql13, ,adOpenStatic, adLockReadOnly : insp1_ativo=rs("ativo") : insp1_afast=rs("afast") : insp1_est=rs("est") : insp1_numero=rs("numero") : insp1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql14, ,adOpenStatic, adLockReadOnly : neg3_ativo=rs("ativo") : neg3_afast=rs("afast") : neg3_est=rs("est") : neg3_numero=rs("numero") : neg3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql15, ,adOpenStatic, adLockReadOnly : copa3_ativo=rs("ativo") : copa3_afast=rs("afast") : copa3_est=rs("est") : copa3_numero=rs("numero") : copa3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql16, ,adOpenStatic, adLockReadOnly : port3_ativo=rs("ativo") : port3_afast=rs("afast") : port3_est=rs("est") : port3_numero=rs("numero") : port3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql17, ,adOpenStatic, adLockReadOnly : audio3_ativo=rs("ativo") : audio3_afast=rs("afast") : audio3_est=rs("est") : audio3_numero=rs("numero") : audio3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql18, ,adOpenStatic, adLockReadOnly : port1_ativo=rs("ativo") : port1_afast=rs("afast") : port1_est=rs("est") : port1_numero=rs("numero") : port1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql19, ,adOpenStatic, adLockReadOnly : tele3_ativo=rs("ativo") : tele3_afast=rs("afast") : tele3_est=rs("est") : tele3_numero=rs("numero") : tele3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql20, ,adOpenStatic, adLockReadOnly : port4_ativo=rs("ativo") : port4_afast=rs("afast") : port4_est=rs("est") : port4_numero=rs("numero") : port4_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql21, ,adOpenStatic, adLockReadOnly : tele1_ativo=rs("ativo") : tele1_afast=rs("afast") : tele1_est=rs("est") : tele1_numero=rs("numero") : tele1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql22, ,adOpenStatic, adLockReadOnly : manu3_ativo=rs("ativo") : manu3_afast=rs("afast") : manu3_est=rs("est") : manu3_numero=rs("numero") : manu3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql23, ,adOpenStatic, adLockReadOnly : manu1_ativo=rs("ativo") : manu1_afast=rs("afast") : manu1_est=rs("est") : manu1_numero=rs("numero") : manu1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql24, ,adOpenStatic, adLockReadOnly : serv3_ativo=rs("ativo") : serv3_afast=rs("afast") : serv3_est=rs("est") : serv3_numero=rs("numero") : serv3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql25, ,adOpenStatic, adLockReadOnly : serv1_ativo=rs("ativo") : serv1_afast=rs("afast") : serv1_est=rs("est") : serv1_numero=rs("numero") : serv1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql26, ,adOpenStatic, adLockReadOnly : serv4_ativo=rs("ativo") : serv4_afast=rs("afast") : serv4_est=rs("est") : serv4_numero=rs("numero") : serv4_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql27, ,adOpenStatic, adLockReadOnly : bib3_ativo=rs("ativo") : bib3_afast=rs("afast") : bib3_est=rs("est") : bib3_numero=rs("numero") : bib3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql28, ,adOpenStatic, adLockReadOnly : bib1_ativo=rs("ativo") : bib1_afast=rs("afast") : bib1_est=rs("est") : bib1_numero=rs("numero") : bib1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql29, ,adOpenStatic, adLockReadOnly : caef3_ativo=rs("ativo") : caef3_afast=rs("afast") : caef3_est=rs("est") : caef3_numero=rs("numero") : caef3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql30, ,adOpenStatic, adLockReadOnly : caef1_ativo=rs("ativo") : caef1_afast=rs("afast") : caef1_est=rs("est") : caef1_numero=rs("numero") : caef1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql31, ,adOpenStatic, adLockReadOnly : secg3_ativo=rs("ativo") : secg3_afast=rs("afast") : secg3_est=rs("est") : secg3_numero=rs("numero") : secg3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql32, ,adOpenStatic, adLockReadOnly : secg1_ativo=rs("ativo") : secg1_afast=rs("afast") : secg1_est=rs("est") : secg1_numero=rs("numero") : secg1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql33, ,adOpenStatic, adLockReadOnly : arqg3_ativo=rs("ativo") : arqg3_afast=rs("afast") : arqg3_est=rs("est") : arqg3_numero=rs("numero") : arqg3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql34, ,adOpenStatic, adLockReadOnly : secgp1_ativo=rs("ativo") : secgp1_afast=rs("afast") : secgp1_est=rs("est") : secgp1_numero=rs("numero") : secgp1_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql35, ,adOpenStatic, adLockReadOnly : ouvi3_ativo=rs("ativo") : ouvi3_afast=rs("afast") : ouvi3_est=rs("est") : ouvi3_numero=rs("numero") : ouvi3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql36, ,adOpenStatic, adLockReadOnly : atend3_ativo=rs("ativo") : atend3_afast=rs("afast") : atend3_est=rs("est") : atend3_numero=rs("numero") : atend3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql37, ,adOpenStatic, adLockReadOnly : edifieo_ativo=rs("ativo") : edifieo_afast=rs("afast") : edifieo_est=rs("est") : edifieo_numero=rs("numero") : edifieo_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql38, ,adOpenStatic, adLockReadOnly : juri3_ativo=rs("ativo") : juri3_afast=rs("afast") : juri3_est=rs("est") : juri3_numero=rs("numero") : juri3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql39, ,adOpenStatic, adLockReadOnly : juri1_ativo=rs("ativo") : juri1_afast=rs("afast") : juri1_est=rs("est") : juri1_numero=rs("numero") : juri1_valor=formatnumber(rs("valor"),0) : rs.close

rs.Open sql40, ,adOpenStatic, adLockReadOnly : obra_ativo=rs("ativo") : obra_afast=rs("afast") : obra_est=rs("est") : obra_numero=rs("numero") : obra_valor=rs("valor") : rs.close
rs.Open sql41, ,adOpenStatic, adLockReadOnly : plan_ativo=rs("ativo") : plan_afast=rs("afast") : plan_est=rs("est") : plan_numero=rs("numero") : plan_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql42, ,adOpenStatic, adLockReadOnly : proto_ativo=rs("ativo") : proto_afast=rs("afast") : proto_est=rs("est") : proto_numero=rs("numero") : proto_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql43, ,adOpenStatic, adLockReadOnly : motor_ativo=rs("ativo") : motor_afast=rs("afast") : motor_est=rs("est") : motor_numero=rs("numero") : motor_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql44, ,adOpenStatic, adLockReadOnly : cpd3_ativo=rs("ativo") : cpd3_afast=rs("afast") : cpd3_est=rs("est") : cpd3_numero=rs("numero") : cpd3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql45, ,adOpenStatic, adLockReadOnly : bolsa3_ativo=rs("ativo") : bolsa3_afast=rs("afast") : bolsa3_est=rs("est") : bolsa3_numero=rs("numero") : bolsa3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql46, ,adOpenStatic, adLockReadOnly : pesq_ativo=rs("ativo") : pesq_afast=rs("afast") : pesq_est=rs("est") : pesq_numero=rs("numero") : pesq_valor=formatnumber(rs("valor"),0) : rs.close

rs.Open sql47, ,adOpenStatic, adLockReadOnly : insp4_ativo=rs("ativo") : insp4_afast=rs("afast") : insp4_est=rs("est") : insp4_numero=rs("numero") : insp4_valor=rs("valor") : rs.close
rs.Open sql48, ,adOpenStatic, adLockReadOnly : efa3_ativo=rs("ativo") : efa3_afast=rs("afast") : efa3_est=rs("est") : efa3_numero=rs("numero") : efa3_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql49, ,adOpenStatic, adLockReadOnly : cefa_ativo=rs("ativo") : cefa_afast=rs("afast") : cefa_est=rs("est") : cefa_numero=rs("numero") : cefa_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql50, ,adOpenStatic, adLockReadOnly : alfa_ativo=rs("ativo") : alfa_afast=rs("afast") : alfa_est=rs("est") : alfa_numero=rs("numero") : alfa_valor=formatnumber(rs("valor"),0) : rs.close

rs.Open sql51, ,adOpenStatic, adLockReadOnly : secefa_ativo=rs("ativo") : secefa_afast=rs("afast") : secefa_est=rs("est") : secefa_numero=rs("numero") : secefa_valor=rs("valor") : rs.close
rs.Open sql52, ,adOpenStatic, adLockReadOnly : cdho_ativo=rs("ativo") : cdho_afast=rs("afast") : cdho_est=rs("est") : cdho_numero=rs("numero") : cdho_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql53, ,adOpenStatic, adLockReadOnly : cfis_ativo=rs("ativo") : cfis_afast=rs("afast") : cfis_est=rs("est") : cfis_numero=rs("numero") : cfis_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql54, ,adOpenStatic, adLockReadOnly : lfis_ativo=rs("ativo") : lfis_afast=rs("afast") : lfis_est=rs("est") : lfis_numero=rs("numero") : lfis_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql55, ,adOpenStatic, adLockReadOnly : lbio_ativo=rs("ativo") : lbio_afast=rs("afast") : lbio_est=rs("est") : lbio_numero=rs("numero") : lbio_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql56, ,adOpenStatic, adLockReadOnly : lphi_ativo=rs("ativo") : lphi_afast=rs("afast") : lphi_est=rs("est") : lphi_numero=rs("numero") : lphi_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql57, ,adOpenStatic, adLockReadOnly : lpho_ativo=rs("ativo") : lpho_afast=rs("afast") : lpho_est=rs("est") : lpho_numero=rs("numero") : lpho_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql58, ,adOpenStatic, adLockReadOnly : lrad_ativo=rs("ativo") : lrad_afast=rs("afast") : lrad_est=rs("est") : lrad_numero=rs("numero") : lrad_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql59, ,adOpenStatic, adLockReadOnly : ltv_ativo=rs("ativo") : ltv_afast=rs("afast") : ltv_est=rs("est") : ltv_numero=rs("numero") : ltv_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql60, ,adOpenStatic, adLockReadOnly : lqui_ativo=rs("ativo") : lqui_afast=rs("afast") : lqui_est=rs("est") : lqui_numero=rs("numero") : lqui_valor=formatnumber(rs("valor"),0) : rs.close

rs.Open sql61, ,adOpenStatic, adLockReadOnly : ltur_ativo=rs("ativo") : ltur_afast=rs("afast") : ltur_est=rs("est") : ltur_numero=rs("numero") : ltur_valor=rs("valor") : rs.close
rs.Open sql62, ,adOpenStatic, adLockReadOnly : rec4_ativo=rs("ativo") : rec4_afast=rs("afast") : rec4_est=rs("est") : rec4_numero=rs("numero") : rec4_valor=rs("valor") : rs.close
rs.Open sql63, ,adOpenStatic, adLockReadOnly : cpsi_ativo=rs("ativo") : cpsi_afast=rs("afast") : cpsi_est=rs("est") : cpsi_numero=rs("numero") : cpsi_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql64, ,adOpenStatic, adLockReadOnly : scbm_ativo=rs("ativo") : scbm_afast=rs("afast") : scbm_est=rs("est") : scbm_numero=rs("numero") : scbm_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql65, ,adOpenStatic, adLockReadOnly : scbv_ativo=rs("ativo") : scbv_afast=rs("afast") : scbv_est=rs("est") : scbv_numero=rs("numero") : scbv_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql66, ,adOpenStatic, adLockReadOnly : scbp_ativo=rs("ativo") : scbp_afast=rs("afast") : scbp_est=rs("est") : scbp_numero=rs("numero") : scbp_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql67, ,adOpenStatic, adLockReadOnly : scpvy_ativo=rs("ativo") : scpvy_afast=rs("afast") : scpvy_est=rs("est") : scpvy_numero=rs("numero") : scpvy_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql68, ,adOpenStatic, adLockReadOnly : scpns_ativo=rs("ativo") : scpns_afast=rs("afast") : scpns_est=rs("est") : scpns_numero=rs("numero") : scpns_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql69, ,adOpenStatic, adLockReadOnly : scdir_ativo=rs("ativo") : scdir_afast=rs("afast") : scdir_est=rs("est") : scdir_numero=rs("numero") : scdir_valor=formatnumber(rs("valor"),0) : rs.close

rs.Open sql70, ,adOpenStatic, adLockReadOnly : ssaj_ativo=rs("ativo") : ssaj_afast=rs("afast") : ssaj_est=rs("est") : ssaj_numero=rs("numero") : ssaj_valor=rs("valor") : rs.close
rs.Open sql71, ,adOpenStatic, adLockReadOnly : pradsec_ativo=rs("ativo") : pradsec_afast=rs("afast") : pradsec_est=rs("est") : pradsec_numero=rs("numero") : pradsec_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql72, ,adOpenStatic, adLockReadOnly : pradass_ativo=rs("ativo") : pradass_afast=rs("afast") : pradass_est=rs("est") : pradass_numero=rs("numero") : if rs("valor")<>"" then pradass_valor=formatnumber(rs("valor"),0) else pradass_valor=0 : rs.close
rs.Open sql73, ,adOpenStatic, adLockReadOnly : pracsec_ativo=rs("ativo") : pracsec_afast=rs("afast") : pracsec_est=rs("est") : pracsec_numero=rs("numero") : pracsec_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql74, ,adOpenStatic, adLockReadOnly : precsec_ativo=rs("ativo") : precsec_afast=rs("afast") : precsec_est=rs("est") : precsec_numero=rs("numero") : precsec_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql75, ,adOpenStatic, adLockReadOnly : prdrcsec_ativo=rs("ativo") : prdrcsec_afast=rs("afast") : prdrcsec_est=rs("est") : prdrcsec_numero=rs("numero") : if rs("valor")<>"" then prdrcsec_valor=formatnumber(rs("valor"),0) else prdrcsec_valor=0 : rs.close
rs.Open sql76, ,adOpenStatic, adLockReadOnly : precass_ativo=rs("ativo") : precass_afast=rs("afast") : precass_est=rs("est") : precass_numero=rs("numero") : if rs("valor")<>"" then precass_valor=formatnumber(rs("valor"),0) else precass_valor=0 : rs.close
rs.Open sql77, ,adOpenStatic, adLockReadOnly : sdcm_ativo=rs("ativo") : sdcm_afast=rs("afast") : sdcm_est=rs("est") : sdcm_numero=rs("numero") : sdcm_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql78, ,adOpenStatic, adLockReadOnly : sdct_ativo=rs("ativo") : sdct_afast=rs("afast") : sdct_est=rs("est") : sdct_numero=rs("numero") : if rs("valor")<>"" then sdct_valor=formatnumber(rs("valor"),0) else sdct_valor=0 : rs.close
rs.Open sql79, ,adOpenStatic, adLockReadOnly : sdcn_ativo=rs("ativo") : sdcn_afast=rs("afast") : sdcn_est=rs("est") : sdcn_numero=rs("numero") : sdcn_valor=formatnumber(rs("valor"),0) : rs.close
rs.Open sql80, ,adOpenStatic, adLockReadOnly : sdcd_ativo=rs("ativo") : sdcd_afast=rs("afast") : sdcd_est=rs("est") : sdcd_numero=rs("numero") : if rs("valor")<>"" then sdcd_valor=formatnumber(rs("valor"),0) else sdcd_valor=0 : rs.close

rs.Open sql81, ,adOpenStatic, adLockReadOnly : smt_ativo=rs("ativo") : smt_afast=rs("afast") : smt_est=rs("est") : smt_numero=rs("numero") : smt_valor=rs("valor") : rs.close
rs.Open sql82, ,adOpenStatic, adLockReadOnly : pracass_ativo=rs("ativo") : pracass_afast=rs("afast") : pracass_est=rs("est") : pracass_numero=rs("numero") : pracass_valor=formatnumber(rs("valor"),0) : rs.close

if request.form("exibe")="ON" or session("a92")="C" then
	copa1_valor="-":fin1_valor="-":rh3_valor="-":supr3_valor="-":fin3_valor="-":arte3_valor="-":graf3_valor="-"
	seven3_valor="-":tec3_valor="-":tec1_valor="-":cont3_valor="-":insp3_valor="-":insp1_valor="-":neg3_valor="-"
	copa3_valor="-":port3_valor="-":audio3_valor="-":port1_valor="-":tele3_valor="-":port4_valor="-":tele1_valor="-"
	manu3_valor="-":manu1_valor="-":serv3_valor="-":serv1_valor="-":serv4_valor="-":bib3_valor="-":bib1_valor="-"
	caef3_valor="-":caef1_valor="-":secg3_valor="-":secg1_valor="-":arqg3_valor="-":secgp1_valor="-":ouvi3_valor="-"
	atend3_valor="-":edifieo_valor="-":juri3_valor="-":juri1_valor="-":obra_valor="-":plan_valor="-":proto_valor="-"
	motor_valor="-":cpd3_valor="-":bolsa3_valor="-":pesq_valor="-":insp4_valor="-":efa3_valor="-":cefa_valor="-"
	alfa_valor="-":secefa_valor="-":cdho_valor="-":cfis_valor="-":lfis_valor="-":lbio_valor="-":lphi_valor="-"
	lpho_valor="-":lrad_valor="-":ltv_valor="-":lqui_valor="-":ltur_valor="-":rec4_valor="-":cpsi_valor="-"
	scbm_valor="-":scbv_valor="-":scbp_valor="-":scpvy_valor="-":scpns_valor="-":scdir_valor="-":ssaj_valor="-"
	pradsec_valor="-":pradass_valor="-":pracsec_valor="-":precsec_valor="-":prdrcsec_valor="-":precass_valor="-"
	sdcm_valor="-":sdct_valor="-":sdcn_valor="-":sdcd_valor="-":smt_valor="-"
	cifrao=""
else
	cifrao="$"
end if
if request.form("exibe")="ON" then cexibe="checked" else cexibe=""
%>
<form method="POST" name="form" action="organograma.asp">
<p class=titulo>Organograma em <input type=text name=dataquery size=10 class=subli2 value="<%=datacampo%>">
<input type="checkbox" name="exibe" value="ON" <%=cexibe%> onClick="javascript:submit()">

<br>&nbsp;
<div align="center">
<!--

ZERO PÁGINA

-->

<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr><!-- 1 COLULNA -->
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 3 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 4 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 5 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
</tr>

<!-- 0 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=2  width=46></td>

	<td width=1 style="border-left: 1px solid #000000 #000000;border-bottom: 1px solid #000000 #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=97 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=5 align="center" class="campor">
	<b>REITORIA</b><br>
	Dr. José Cassio Soares Hungria</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td colspan=2  width=46></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 0 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5 style="border-left: 1px solid #000000"></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td style="" colspan=2 height=16 width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 3 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 4 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 5 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>

<!-- 1 LINHA -->
<tr><!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>PRO-REITORIA ADMINISTRATIVA</b><br>
	Dr. José Cassio Soares Hungria</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>PRO-REITORIA ACADÊMICA</b><br>
	Dr. Luiz Fernando da Costa e Silva</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>PRO-REITORIA EXTENSÃO E CULTURA</b><br>
	Dr. Luiz Carlos de Azevedo</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>PRO-REITORIA DESENV. RELAÇÕES COMUNITÁRIAS</b><br>
	Dr. Luiz Carlos de Azevedo</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
</tr>	

</table>

<hr>
<DIV style=""page-break-after:always""></DIV>

<!--

PRIMEIRA PAGINA

-->

<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr><!-- 1 COLULNA -->
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 2 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 3 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 4 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 5 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 6 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 7 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 8 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 9 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
</tr>

<!-- 0 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>PRO-REITORIA ADMINISTRATIVA</b><br>
	Dr. José Cassio Soares Hungria</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 9 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 0/a1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5 rowspan=2 width=93>
	<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td width=93 style="border: 1px solid #000000;" class="campor" align="center">
<a class=t href="org_view.asp?setor=PRAD-ASSESSORIA" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
		<b>Assessoria</b><br><%=pradass_ativo%> + (<%=pradass_est%>E <%=pradass_afast%>A)<br><%=cifrao%><%=pradass_valor%></td>
	</tr>
	</table>
	</td>

	<td width=5 style="border-bottom: 1px solid #000000"></td>
<!-- 4 COLULNA -->	
	<td style="border-bottom: 1px solid #000000" colspan=5 width=93></td>

	<td style="border-bottom: 1px solid #000000" width=5></td>
<!-- 5 COLULNA -->	
	<td style="border-bottom: 1px solid #000000" colspan=2 width=46></td>
	<td style="border-left: 1px solid #000000;border-bottom: 1px solid #000000" width=1></td>
	<td style="border-bottom: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-bottom: 1px solid #000000" width=5></td>
<!-- 6 COLULNA -->	
	<td style="border-bottom: 1px solid #000000" colspan=5  width=93></td>

	<td style="border-bottom: 1px solid #000000"  width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=5 rowspan=2 width=93>
	<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td width=93 style="border: 1px solid #000000;" class="campor" align="center">
<a class=t href="org_view.asp?setor=PRAD-SECRETARIA" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
		<b>Secretária/Motor.</b><br><%=pradsec_ativo%> + (<%=pradsec_est%>E <%=pradsec_afast%>A)<br><%=cifrao%><%=pradsec_valor%></td>
	</tr>
	</table>
	</td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 9 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 0/a2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
<!--	<td colspan=5  width=93></td> -->

	<td width=5></td>
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
<!--	<td colspan=5  width=93></td> -->

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 9 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>


<!-- 0/3 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 9 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>


<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=16 width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 2 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 3 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 4 COLULNA -->	
	<td style="border-top: 1px solid #000000"colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 5 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 6 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 7 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 8 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 9 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>

<!-- 1 LINHA -->
<tr><!-- 1 COLULNA -->
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=RECURSOS HUMANOS-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>REC. HUMANOS</b><br><%=rh3_ativo%> + (<%=rh3_est%>E <%=rh3_afast%>A)<br><%=cifrao%><%=rh3_valor%><br><font color=blue>Rogerio</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SUPRIMENTOS-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SUPRIMENTOS</b><br><%=supr3_ativo%> + (<%=supr3_est%>E <%=supr3_afast%>A)<br><%=cifrao%><%=supr3_valor%><br><font color=blue>Luisa</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=FINANCEIRO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>FINANCEIRO-VY</b><br><%=fin3_ativo%> + (<%=fin3_est%>E <%=fin3_afast%>A)<br><%=cifrao%><%=fin3_valor%><br><font color=blue>Marcos Ferrara</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=ARTE/MARKETING-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>ARTE/MKT</b><br><%=arte3_ativo%> + (<%=arte3_est%>E <%=arte3_afast%>A)<br><%=cifrao%><%=arte3_valor%><br><font color=green>Luciana</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=CPD-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>C.P.D.-VY</b><br><%=cpd3_ativo%> + (<%=cpd3_est%>E <%=cpd3_afast%>A)<br><%=cifrao%><%=cpd3_valor%><br><font color=blue>Eduardo Gross</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=CONTABILIDADE-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>CONTABILIDADE</b><br><%=cont3_ativo%> + (<%=cont3_est%>E <%=cont3_afast%>A)<br><%=cifrao%><%=cont3_valor%><br><font color=blue>Vicente</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=INSPETORIA DE ALUNOS-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>INSPETORIA-VY</b><br><%=insp3_ativo%> + (<%=insp3_est%>E <%=insp3_afast%>A)<br><%=cifrao%><%=insp3_valor%><br><font color=blue>Marcelo</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=NEGOCIACAO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>NEGOCIAÇÃO</b><br><%=neg3_ativo%> + (<%=neg3_est%>E <%=neg3_afast%>A)<br><%=cifrao%><%=neg3_valor%><br><font color=blue>Fábio</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 9 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SECRETARIA GERAL-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SEC.GERAL-VY</b><br><%=secg3_ativo%> + (<%=secg3_est%>E <%=secg3_afast%>A)<br><%=cifrao%><%=secg3_valor%><br><font color=blue>Celina</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
</tr>	

<!-- 2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=16 width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 9 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>
</tr>

<!-- 2 LINHA -->
<tr><!-- 1 COLULNA -->
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=PROTOCOLO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>PROTOCOLO</b><br><%=proto_ativo%> + (<%=proto_est%>E <%=proto_afast%>A)<br><%=cifrao%><%=proto_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=GRAFICA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>GRÁFICA</b><br><%=graf3_ativo%> + (<%=graf3_est%>E <%=graf3_afast%>A)<br><%=cifrao%><%=graf3_valor%><br><font color=green>Orlando</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=FINANCEIRO-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>FINANCEIRO-NS</b><br><%=fin1_ativo%> + (<%=fin1_est%>E <%=fin1_afast%>A)<br><%=cifrao%><%=fin1_valor%><br><font color=green>Oswaldo</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SEVEN-SERVICOS DE EVENTOS-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SEVEN-SERV.EVENTOS</b><br><%=seven3_ativo%> + (<%=seven3_est%>E <%=seven3_afast%>A)<br><%=cifrao%><%=seven3_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=TECNICO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>TÉCNICO-VY</b><br><%=tec3_ativo%> + (<%=tec3_est%>E <%=tec3_afast%>A)<br><%=cifrao%><%=tec3_valor%><br><font color=blue></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=INSPETORIA DE ALUNOS-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>INSPETORIA-NS</b><br><%=insp1_ativo%> + (<%=insp1_est%>E <%=insp1_afast%>A)<br><%=cifrao%><%=insp1_valor%><br><font color=green>Raul</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=BOLSA DE ESTUDOS-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>BOLSA DE ESTUDOS-VY</b><br><%=bolsa3_ativo%> + (<%=bolsa3_est%>E <%=bolsa3_afast%>A)<br><%=cifrao%><%=bolsa3_valor%><br><font color=blue>Mário Quinto</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 9 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SECRETARIA GERAL-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SEC.GERAL-NS</b><br><%=secg1_ativo%> + (<%=secg1_est%>E <%=secg1_afast%>A)<br><%=cifrao%><%=secg1_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
</tr>	

<!-- 3 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=16 width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 9 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>
</tr>

<!-- 3 LINHA -->
<tr><!-- 1 COLULNA -->
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=MOTORISTA RH-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>MOTORISTA</b><br><%=motor_ativo%> + (<%=motor_est%>E <%=motor_afast%>A)<br><%=cifrao%><%=motor_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=COPA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>COPA-VY</b><br><%=copa3_ativo%> + (<%=copa3_est%>E <%=copa3_afast%>A)<br><%=cifrao%><%=copa3_valor%><br><font color=green>Neusa</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=TECNICO-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>TÉCNIVO-NS</b><br><%=tec1_ativo%> + (<%=tec1_est%>E <%=tec1_afast%>A)<br><%=cifrao%><%=tec1_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=INSPETORIA DE ALUNOS-JW" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>INSPETORIA-JW</b><br><%=insp4_ativo%> + (<%=insp4_est%>E <%=insp4_afast%>A)<br><%=cifrao%><%=insp4_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 9 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=ARQUIVO GERAL-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>ARQ.GERAL-VY</b><br><%=arqg3_ativo%> + (<%=arqg3_est%>E <%=arqg3_afast%>A)<br><%=cifrao%><%=arqg3_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
</tr>	

<!-- 4 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>
	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 9 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>
</tr>

<!-- 4 LINHA -->
<tr><!-- 1 COLULNA -->
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SEGURANCA E MEDICINA DO TRABALHO" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SEG.MEDICINA TRABALHO</b><br><%=smt_ativo%> + (<%=smt_est%>E <%=smt_afast%>A)<br><%=cifrao%><%=smt_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=COPA-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>COPA-NS</b><br><%=copa1_ativo%> + (<%=copa1_est%>E <%=copa1_afast%>A)<br><%=cifrao%><%=copa1_valor%><br><font color=green>Anesia</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=AUDIO E VIDEO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>ÁUDIO E VIDEO</b><br><%=audio3_ativo%> + (<%=audio3_est%>E <%=audio3_afast%>A)<br><%=cifrao%><%=audio3_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=PORTARIA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>PORTARIA-VY</b><br><%=port3_ativo%> + (<%=port3_est%>E <%=port3_afast%>A)<br><%=cifrao%><%=port3_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 9 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SECRETARIA GERAL POS GRADUACAO-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SEG.GERAL PÓS-NS</b><br><%=secgp1_ativo%> + (<%=secgp1_est%>E <%=secgp1_afast%>A)<br><%=cifrao%><%=secgp1_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
</tr>

<!-- 5 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>
	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 5 LINHA -->
<tr><!-- 1 COLULNA -->
	<td width=93 colspan=5 class="campor"></td>
	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=TELEFONIA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>TELEFONIA-VY</b><br><%=tele3_ativo%> + (<%=tele3_est%>E <%=tele3_afast%>A)<br><%=cifrao%><%=tele3_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=PORTARIA-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>PORTARIA-NS</b><br><%=port1_ativo%> + (<%=port1_est%>E <%=port1_afast%>A)<br><%=cifrao%><%=port1_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>
</tr>	

<!-- 6 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>
	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 6 LINHA -->
<tr><!-- 1 COLULNA -->
	<td width=93 colspan=5 class="campor"></td>
	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=TELEFONIA-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>TELEFONIA-NS</b><br><%=tele1_ativo%> + (<%=tele1_est%>E <%=tele1_afast%>A)<br><%=cifrao%><%=tele1_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=PORTARIA-JW" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>PORTARIA-JW</b><br><%=port4_ativo%> + (<%=port4_est%>E <%=port4_afast%>A)<br><%=cifrao%><%=port4_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 9 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>
</tr>	

<!-- 7 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>
	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 7 LINHA -->
<tr><!-- 1 COLULNA -->
	<td width=93 colspan=5 class="campor"></td>
	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=MANUTENCAO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>MANUTENÇÃO-VY</b><br><%=manu3_ativo%> + (<%=manu3_est%>E <%=manu3_afast%>A)<br><%=cifrao%><%=manu3_valor%><br><font color=green>Nivalter</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>
</tr>	

<!-- 8 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>
	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 8 LINHA -->
<tr><!-- 1 COLULNA -->
	<td width=93 colspan=5 class="campor"></td>
	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=MANUTENCAO-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>MANUTENÇÃO-NS</b><br><%=manu1_ativo%> + (<%=manu1_est%>E <%=manu1_afast%>A)<br><%=cifrao%><%=manu1_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>
</tr>	

<!-- 9 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>
	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 9 LINHA -->
<tr><!-- 1 COLULNA -->
	<td width=93 colspan=5 class="campor"></td>
	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SERVICOS GERAIS-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SERV.GERAIS-VY</b><br><%=serv3_ativo%> + (<%=serv3_est%>E <%=serv3_afast%>A)<br><%=cifrao%><%=serv3_valor%><br><font color=green>Gregório</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>
</tr>	

<!-- 10 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>
	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 10 LINHA -->
<tr><!-- 1 COLULNA -->
	<td width=93 colspan=5 class="campor"></td>
	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SERVICOS GERAIS-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SERV.GERAIS-NS</b><br><%=serv1_ativo%> + (<%=serv1_est%>E <%=serv1_afast%>A)<br><%=cifrao%><%=serv1_valor%><br><font color=green>Manoel</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>
</tr>	

<!-- 11 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>
	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 11 LINHA -->
<tr><!-- 1 COLULNA -->
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SERVICOS GERAIS-JW" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SERV.GERAIS-JW</b><br><%=serv4_ativo%> + (<%=serv4_est%>E <%=serv4_afast%>A)<br><%=cifrao%><%=serv4_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>
</tr>	

</table>

<hr>
<DIV style="page-break-after:always"></DIV>

<!--

SEGUNDA PÁGINA

-->

<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr><!-- 1 COLULNA -->
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 2 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 3 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 4 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 5 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 6 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
</tr>

<!-- 0 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=2  width=46></td>

	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=97 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=5 align="center" class="campor">
	<b>PRO-REITORIA ACADÊMICA</b><br>
	Dr. Luiz Fernando da Costa e Silva</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td colspan=2  width=46></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 0/a1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5 rowspan=2 width=93>
	<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td width=93 style="border: 1px solid #000000;" class="campor" align="center">
<a class=t href="org_view.asp?setor=PRAC-ASSESSORIA" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
		<b>Assessoria</b><br><%=pracass_ativo%> + (<%=pracass_est%>E <%=pracass_afast%>A)<br><%=cifrao%><%=pracass_valor%></td>
	</tr>
	</table>
	</td>

	<td style="border-bottom: 1px solid #000000"  width=5></td>
<!-- 3 COLULNA -->	
	<td style="border-bottom: 1px solid #000000" colspan=5 width=93></td>

	<td style="border-left: 1px solid #000000;border-bottom: 1px solid #000000" width=5></td>
<!-- 4 COLULNA -->	
	<td style="border-bottom: 1px solid #000000" colspan=5 width=93></td>

	<td style="border-bottom: 1px solid #000000"  width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5 rowspan=2 width=93>
	<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td width=93 style="border: 1px solid #000000;" class="campor" align="center">
<a class=t href="org_view.asp?setor=PRAC-SECRETARIA" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
		<b>Secretária/Motor.</b><br><%=pracsec_ativo%> + (<%=pracsec_est%>E <%=pracsec_afast%>A)<br><%=cifrao%><%=pracsec_valor%></td>
	</tr>
	</table>
	</td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>

</tr>

<!-- 0/a2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
<!--	<td colspan=5  width=93></td> -->

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td style="border-left: 1px solid #000000;" width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=2  width=46></td>
	<td width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
<!--	<td colspan=5  width=93></td> -->

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 0 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5 style="border-left: 1px solid #000000"></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=16 width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 2 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 3 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 4 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 5 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 6 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>

<!-- 1 LINHA -->
<tr><!-- 1 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=BIBLIOTECA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>BIBLIOTECA-VY</b><br><%=bib3_ativo%> + (<%=bib3_est%>E <%=bib3_afast%>A)<br><%=cifrao%><%=bib3_valor%><br><font color=blue>Maria Helena</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=CAEF-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>CAEF-VY</b><br><%=caef3_ativo%> + (<%=caef3_est%>E <%=caef3_afast%>A)<br><%=cifrao%><%=caef3_valor%><br><font color=red>Profa.Ivanize</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=OUVIDORIA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>OUVIDORIA-VY</b><br><%=ouvi3_ativo%> + (<%=ouvi3_est%>E <%=ouvi3_afast%>A)<br><%=cifrao%><%=ouvi3_valor%><br><font color=blue>Mariana</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=EDIFIEO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>EDIFIEO-VY</b><br><%=edifieo_ativo%> + (<%=edifieo_est%>E <%=edifieo_afast%>A)<br><%=cifrao%><%=edifieo_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=DEPTO. JURIDICO" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>DEPTO.JURIDICO-VY</b><br><%=juri3_ativo%> + (<%=juri3_est%>E <%=juri3_afast%>A)<br><%=cifrao%><%=juri3_valor%><br><font color=blue>Dr.Carlos Alberto</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=PLANEJAMENTO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>PLAN.ACADÊMICO</b><br><%=plan_ativo%> + (<%=plan_est%>E <%=plan_afast%>A)<br><%=cifrao%><%=plan_valor%><br><font color=red>Dra.Gabriela</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
</tr>	

<!-- 2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>
</tr>

<!-- 2 LINHA -->
<tr><!-- 1 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=BIBLIOTECA-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>BIBLIOTECA-NS</b><br><%=bib1_ativo%> + (<%=bib1_est%>E <%=bib1_afast%>A)<br><%=cifrao%><%=bib1_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=CAEF-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>CAEF-NS</b><br><%=caef1_ativo%> + (<%=caef1_est%>E <%=caef1_afast%>A)<br><%=cifrao%><%=caef1_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=CENTRAL DE ATENDIMENTO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>CENTRAL ATEND.-VY</b><br><%=atend3_ativo%> + (<%=atend3_est%>E <%=atend3_afast%>A)<br><%=cifrao%><%=atend3_valor%><br><font color=green>Ana Paula</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=JUIZADO DE PEQUENAS CAUSAS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>JUIZ.PEQUENAS CAUSAS-NS</b><br><%=juri1_ativo%> + (<%=juri1_est%>E <%=juri1_afast%>A)<br><%=cifrao%><%=juri1_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=PESQUISA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>PESQUISA-VY</b><br><%=pesq_ativo%> + (<%=pesq_est%>E <%=pesq_afast%>A)<br><%=cifrao%><%=pesq_valor%><br><font color=red>Profa.Edna</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
</tr>	

</table>

<hr>
<DIV style="page-break-after:always"></DIV>

<!--

TERCEIRA PÁGINA

-->

<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr><!-- 1 COLULNA -->
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 2 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 3 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 4 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 5 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
</tr>

<!-- 0 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>
	
	<td width=5></td>
<!-- 3 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>PRO-REITORIA EXTENSÃO E CULTURA</b><br>
	Dr. Luiz Carlos de Azevedo</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 0/a1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 rowspan=2 width=93>
	<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td width=93 style="border: 1px solid #000000;" class="campor" align="center">
<a class=t href="org_view.asp?setor=PREC-ASSESSORIA" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
		<b>Assessoria</b><br><%=precass_ativo%> + (<%=precass_est%>E <%=precass_afast%>A)<br><%=cifrao%><%=precass_valor%></td>
	</tr>
	</table>
	</td>

	<td style="border-bottom: 1px solid #000000" width=5></td>
<!-- 2 COLULNA -->	
	<td style="border-bottom: 1px solid #000000" colspan=5  width=93></td>

	<td style="border-bottom: 1px solid #000000" width=5></td>
<!-- 3 COLULNA -->	
	<td style="border-bottom: 1px solid #000000" colspan=2 width=46></td>
	<td style="border-left: 1px solid #000000;border-bottom: 1px solid #000000" width=1></td>
	<td style="border-bottom: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-bottom: 1px solid #000000" width=5></td>
<!-- 4 COLULNA -->	
	<td style="border-bottom: 1px solid #000000" colspan=5  width=93></td>

	<td style="border-bottom: 1px solid #000000"  width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5 rowspan=2 width=93>
	<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td width=93 style="border: 1px solid #000000;" class="campor" align="center">
<a class=t href="org_view.asp?setor=PREC-SECRETARIA" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
		<b>Secretária/Motor.</b><br><%=precsec_ativo%> + (<%=precsec_est%>E <%=precsec_afast%>A)<br><%=cifrao%><%=precsec_valor%></td>
	</tr>
	</table>
	</td>
</tr>

<!-- 0/a2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
<!--	<td colspan=5 height=16 width=93></td>-->

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
<!--	<td colspan=5  width=93></td> -->

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 0 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2 height=16 width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 3 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 4 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 1 LINHA -->
<tr><!-- 1 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=EFA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>EFA</b><br><%=efa3_ativo%> + (<%=efa3_est%>E <%=efa3_afast%>A)<br><%=cifrao%><%=efa3_valor%><br><font color=blue>Maria Cristina</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=1 style="" class="campor"></td>
	<td width=91 style="" colspan=3 align="center" class="campor"></td>
	<td width=1 style="" class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=CURSOS ALFA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>ALFA</b><br><%=alfa_ativo%> + (<%=alfa_est%>E <%=alfa_afast%>A)<br><%=cifrao%><%=alfa_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>	

<!-- 2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2 height=9 width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 2/2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=2 height=9 width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td  style="border-top: 1px solid #000000"width=5></td>
<!-- 3 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 2 LINHA -->
<tr><!-- 1 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=CURSO EFA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>CURSOS EFA</b><br><%=cefa_ativo%> + (<%=cefa_est%>E <%=cefa_afast%>A)<br><%=cifrao%><%=cefa_valor%><br><font color=blue>Minighitti</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SECRETARIA DO CURSO EFA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SECR.CURSO EFA</b><br><%=secefa_ativo%> + (<%=secefa_est%>E <%=secefa_afast%>A)<br><%=cifrao%><%=secefa_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>	

</table>

<hr>
<DIV style="page-break-after:always"></DIV>

<!--

QUARTA PÁGINA

-->

<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr><!-- 1 COLULNA -->
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 2 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 3 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
</tr>

<!-- 0 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>PRO-REITORIA DESENV. REL. COMUNITÁRIAS</b><br>
	Dr. Luiz Carlos de Azevedo</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 0/a1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 width=46></td>
	<td style="border-left: 1px solid #000000;border-bottom: 1px solid #000000" width=1></td>
	<td style="border-bottom: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-bottom: 1px solid #000000" width=5></td>
<!-- 2 COLULNA -->	
	<td style="border-bottom: 1px solid #000000" colspan=5  width=93></td>

	<td style="border-bottom: 1px solid #000000"  width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5 rowspan=2 width=93>
	<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
	<tr>
		<td width=93 style="border: 1px solid #000000;" class="campor" align="center">
<a class=t href="org_view.asp?setor=PRDRC-SECRETARIA" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
		<b>Secretária/Motor.</b><br><%=prdrcsec_ativo%> + (<%=prdrcsec_est%>E <%=prdrcsec_afast%>A)<br><%=cifrao%><%=prdrcsec_valor%></td>
	</tr>
	</table>
	</td>
</tr>

<!-- 0/a2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2  width=46></td>
	<td style="" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
<!--	<td colspan=5  width=93></td> -->

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

</table>

<hr>
<DIV style="page-break-after:always"></DIV>

<!--

QUINTA PÁGINA

-->

<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr><!-- 1 COLULNA -->
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 2 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 3 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 4 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 5 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 6 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 7 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 8 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 9 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 10 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 11 COLULNA -->
	<td width=5></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
</tr>

<!-- 0/2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>DIRETORIAS DE CURSOS</b></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 9 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 10 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 11 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 0/3 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=5 height=16 width=93></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 3 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 4 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 5 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 6 COLULNA -->	
	<td colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td colspan=2 width=46></td>

	<td width=5></td>
<!-- 7 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 8 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 9 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 10 COLULNA -->	
	<td colspan=5  width=93></td>

	<td width=5></td>
<!-- 11 COLULNA -->	
	<td colspan=5  width=93></td>
</tr>

<!-- 0 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=16 width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 2 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 3 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-top: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 4 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 5 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 6 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-top: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 7 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 8 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-top: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 9 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 10 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-top: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 11 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>

<!-- 0 LINHA -->
<tr><!-- 1 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Curso de HISTÓRIA</b><br><font color=red>Profa. Maria Cecilia</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Curso de FISIOTERAPIA</b><br><font color=red>Prof. Reginaldo</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td colspan=5 width=93></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Curso de CIÊNCIAS BIOLÓGICAS</b><br><font color=red>Profa. Adriana</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Curso de TURISMO</b><br><font color=red>Prof. Mauro</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td colspan=5 width=93></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Curso de COMUNICAÇÃO SOCIAL</b><br><font color=red>Profª Helena</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td colspan=5 width=93></td>

	<td width=5></td>	
<!-- 9 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Curso de QUIMICA</b><br><font color=red>Profa. Márcia Biaggi</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 10 COLULNA -->	
	<td colspan=5 width=93></td>

	<td width=5></td>	
<!-- 11 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Curso de PSICO PEDAGOGIA</b><br><font color=red>Profa. Márcia Siqueira</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
</tr>	

<!-- 0/1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=9 width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 3 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 4 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 5 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 6 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 7 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 8 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 9 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 10 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 11 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>

<!-- 1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=9 width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td width=5></td>
<!-- 2 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 3 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 4 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 5 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 6 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 7 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 8 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 9 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 10 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 11 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>

<!-- 1 LINHA -->
<tr><!-- 1 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=CDHO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>CDHO</b><br><%=cdho_ativo%> + (<%=cdho_est%>E <%=cdho_afast%>A)<br><%=cifrao%><%=cdho_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=CLINICA DE FISIOTERAPIA-JW" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>CLIN. FISIOTERAPIA</b><br><%=cfis_ativo%> + (<%=cfis_est%>E <%=cfis_afast%>A)<br><%=cifrao%><%=cfis_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=LABORATORIO DE FISIOTERAPIA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>LAB. FISIOTERAPIA</b><br><%=lfis_ativo%> + (<%=lfis_est%>E <%=lfis_afast%>A)<br><%=cifrao%><%=lfis_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=LABORATORIO DE BIOLOGIA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>LAB. BIOLOGIA</b><br><%=lbio_ativo%> + (<%=lbio_est%>E <%=lbio_afast%>A)<br><%=cifrao%><%=lbio_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=LABORATORIO DE TURISMO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>LAB. TURISMO</b><br><%=ltur_ativo%> + (<%=ltur_est%>E <%=ltur_afast%>A)<br><%=cifrao%><%=ltur_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=LABORATORIO DE FOTOGRAFIA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>LAB. FOTOGRAFIA</b><br><%=lpho_ativo%> + (<%=lpho_est%>E <%=lpho_afast%>A)<br><%=cifrao%><%=lpho_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=LABORATORIO DE RADIO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>LAB. RÁDIO</b><br><%=lrad_ativo%> + (<%=lrad_est%>E <%=lrad_afast%>A)<br><%=cifrao%><%=lrad_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=LABORATORIO DE TV-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>LAB. TV</b><br><%=ltv_ativo%> + (<%=ltv_est%>E <%=ltv_afast%>A)<br><%=cifrao%><%=ltv_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 9 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=LABORATORIO DE QUIMICA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>LAB. QUIMICA</b><br><%=lqui_ativo%> + (<%=lqui_est%>E <%=lqui_afast%>A)<br><%=cifrao%><%=lqui_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 10 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=LABORATORIO DE FISICA-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>LAB. FÍSICA</b><br><%=lphi_ativo%> + (<%=lphi_est%>E <%=lphi_afast%>A)<br><%=cifrao%><%=lphi_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 11 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=CLINICA DE PSICOPEDAGOGIA-JW" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>CLIN. PSICO PEDAGOGIA</b><br><%=cpsi_ativo%> + (<%=cpsi_est%>E <%=cpsi_afast%>A)<br><%=cifrao%><%=cpsi_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
</tr>	

<!-- 2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=16 width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 2 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 3 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td style="" width=5></td>
<!-- 4 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td style="" width=5></td>
<!-- 5 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td style="" width=5></td>
<!-- 6 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td style="" width=5></td>
<!-- 7 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td style="" width=5></td>
<!-- 8 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td style="" width=5></td>
<!-- 9 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td style="" width=5></td>
<!-- 10 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td style="" width=5></td>
<!-- 11 COLULNA -->	
	<td style="" colspan=5  width=93></td>
</tr>

<!-- 1 LINHA -->
<tr><!-- 1 COLULNA -->	
	<td style="" colspan=5 height=16 width=93></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=RECEPCAO-JW" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>RECEPÇÃO-JW</b><br><%=rec4_ativo%> + (<%=rec4_est%>E <%=rec4_afast%>A)<br><%=cifrao%><%=rec4_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td width=5></td>	
<!-- 8 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td width=5></td>	
<!-- 9 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td width=5></td>	
<!-- 10 COLULNA -->	
	<td style="" colspan=5  width=93></td>

	<td width=5></td>	
<!-- 11 COLULNA -->	
	<td style="" colspan=5  width=93></td>
</tr>	
</table>

<hr>
<DIV style="page-break-after:always"></DIV>

<!--

SEXTA PÁGINA

-->

<table border="0" bordercolor=#CCCCCC cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr><!-- 1 COLULNA -->
	<td width=1></td><td width=45></td><td width=10></td><td width=45></td><td width=1></td>
<!-- 2 COLULNA -->
	<td width=15></td>
	<td width=1></td><td width=45></td><td width=10></td><td width=45></td><td width=1></td>
<!-- 3 COLULNA -->
	<td width=15></td>
	<td width=1></td><td width=45></td><td width=10></td><td width=45></td><td width=1></td>
<!-- 4 COLULNA -->
	<td width=15></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 5 COLULNA -->
	<td width=1></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
<!-- 6 COLULNA -->
	<td width=1></td>
	<td width=1></td><td width=45></td><td width=10></td><td width=45></td><td width=1></td>
<!-- 7 COLULNA -->
	<td width=15></td>
	<td width=1></td><td width=45></td><td width=1></td><td width=45></td><td width=1></td>
</tr>

<!-- 0 LINHA -->
<tr><!-- 1 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Diretoria Cursos Matutinos</b><br>Maria Célia Soares Hungria de Luca</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Diretoria Cursos Tecnologia</b><br>José Eduardo de Mello Freire</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Diretoria Cursos Noturnos</b><br>Luiz Carlos de Avezedo Filho</td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Diretoria Cursos Pós-Graduação</b></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>	

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
	<b>Diretoria Curso Direito</b><br> - </td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=93 colspan=5 class="campor"></td>	
</tr>	

<!-- 0/S1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=9 width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 2 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 3 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 4 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 5 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 6 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 7 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>

<!-- 0/S2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=9 width=46></td>
	<td style="border-left: 1px solid #000000;border-bottom: 1px solid #000000" width=1></td>
	<td style="border: 1px solid #000000" colspan=3 rowspan=2 class="campor" align="center">
<a class=t href="org_view.asp?setor=SECR. DIR.CURSO MATUTINO" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>Secretária</b><br><%=sdcm_ativo%> + (<%=sdcm_est%>E <%=sdcm_afast%>A)<br><%=cifrao%><%=sdcm_valor%></td>

<!-- 2 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-bottom: 1px solid #000000" width=1></td>
	<td style="border: 1px solid #000000" colspan=3 rowspan=2 class="campor" align="center"> 
<a class=t href="org_view.asp?setor=SECR. DIR.CURSO TECNOLOGIA" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>Secretária</b><br><%=sdct_ativo%> + (<%=sdct_est%>E <%=sdct_afast%>A)<br><%=cifrao%><%=sdct_valor%></td>

<!-- 3 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-bottom: 1px solid #000000" width=1></td>
	<td style="border: 1px solid #000000" colspan=3 rowspan=2 class="campor" align="center">
<a class=t href="org_view.asp?setor=SECR. DIR.CURSO NOTURNO" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>Secretária</b><br><%=sdcn_ativo%> + (<%=sdcn_est%>E <%=sdcn_afast%>A)<br><%=cifrao%><%=sdcn_valor%></td>

<!-- 4 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 5 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 6 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-bottom: 1px solid #000000" width=1></td>
	<td style="border: 1px solid #000000" colspan=3 rowspan=2 class="campor" align="center">
<a class=t href="org_view.asp?setor=SECR. DIR.CURSO DIREITO" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>Secretária</b><br><%=sdcd_ativo%> + (<%=sdcd_est%>E <%=sdcd_afast%>A)<br><%=cifrao%><%=sdcd_valor%></td>

<!-- 7 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>

<!-- 0/S3 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=9 width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
<!--	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td> -->
<!-- 2 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
<!--	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td> -->
<!-- 3 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
<!--	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td> -->
<!-- 4 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 5 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 6 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
<!--	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td> -->
<!-- 7 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>


<!-- 0/1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=9 width=46></td>
	<td style="border-left: 1px solid #000000;border-bottom: 1px solid #000000" width=1></td>
	<td style="border-bottom: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-bottom: 1px solid #000000" width=5></td>
<!-- 2 COLULNA -->	
	<td style="border-bottom: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-bottom: 1px solid #000000" width=1></td>
	<td style="border-bottom: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-bottom: 1px solid #000000" width=5></td>
<!-- 3 COLULNA -->	
	<td style="border-bottom: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 4 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 5 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 6 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 7 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>

<!-- 0/2 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=16 width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 2 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 3 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 4 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 5 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 6 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 7 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>


<!-- 1 LINHA -->
<tr><!-- 1 COLULNA / LINHA DOS TRACOS-->
	<td colspan=2 height=9 width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 2 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 3 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 4 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 5 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>

	<td style="" width=5></td>
<!-- 6 COLULNA -->	
	<td style="" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;border-top: 1px solid #000000" width=1></td>
	<td style="border-top: 1px solid #000000" colspan=2 width=46></td>

	<td style="border-top: 1px solid #000000" width=5></td>
<!-- 7 COLULNA -->	
	<td style="border-top: 1px solid #000000" colspan=2  width=46></td>
	<td style="border-left: 1px solid #000000;" width=1></td>
	<td style="" colspan=2 width=46></td>
</tr>

<!-- 1 LINHA -->
<tr><!-- 1 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SECRETARIA DE CURSO - BL.PRATA" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SECR.CURSO BL.PRATA</b><br><%=scbp_ativo%> + (<%=scbp_est%>E <%=scbp_afast%>A)<br><%=cifrao%><%=scbp_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 2 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SECRETARIA DE CURSO - BL.VERDE" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SECR.CURSO BL.VERDE</b><br><%=scbv_ativo%> + (<%=scbv_est%>E <%=scbv_afast%>A)<br><%=cifrao%><%=scbv_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 3 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SECRETARIA DE CURSO - BL.MARRON" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SECR.CURSO BL.MARRON</b><br><%=scbm_ativo%> + (<%=scbm_est%>E <%=scbm_afast%>A)<br><%=cifrao%><%=scbm_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 4 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SECRETARIA POS GRADUACAO-VY" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SECR.PÓS-VY</b><br><%=scpvy_ativo%> + (<%=scpvy_est%>E <%=scpvy_afast%>A)<br><%=cifrao%><%=scpvy_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 5 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SECRETARIA POS GRADUACAO-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SECR.PÓS-NS</b><br><%=scpns_ativo%> + (<%=scpns_est%>E <%=scpns_afast%>A)<br><%=cifrao%><%=scpns_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 6 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SECRETARIA DO CURSO DE DIREITO-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SECR.CURSO DIREITO-NS</b><br><%=scdir_ativo%> + (<%=scdir_est%>E <%=scdir_afast%>A)<br><%=cifrao%><%=scdir_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>

	<td width=5></td>	
<!-- 7 COLULNA -->	
	<td width=1 style="border-left: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
	<td width=91 style="border-bottom: 1px solid #000000;border-top: 1px solid #000000" colspan=3 align="center" class="campor">
<a class=t href="org_view.asp?setor=SECRETARIA DO SAJ-NS" onclick="NewWindow(this.href,'Ver_Organograma','550','300','yes','center');return false" onfocus="this.blur()"><font color=black>
	<b>SECR. SAJ-NS</b><br><%=ssaj_ativo%> + (<%=ssaj_est%>E <%=ssaj_afast%>A)<br><%=cifrao%><%=ssaj_valor%></td>
	<td width=1 style="border-right: 1px solid #000000;border-bottom: 1px solid #000000;border-top: 1px solid #000000" class="campor"></td>
</tr>	



</table>

</div>
</form>
<%
	pagina=pagina+1
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"

'rs.close
'set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>