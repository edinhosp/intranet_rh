<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a80")="N" or session("a80")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Alteração de Grade Horária</title>
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
<script language="JavaScript" type="text/javascript"><!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

function nome3() {	form.chapa1.value=form.nome1.value;
form.submit();	}
function chapa3() {	form.nome1.value=form.chapa1.value;
form.submit();	}
function nome4() {	form.chapa2.value=form.nome2.value;
form.submit();	}
function chapa4() {	form.nome2.value=form.chapa2.value;
form.submit();	}
--></script>
<script language="VBScript">
</script>

<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
'if session("usuariomaster")="02379" then response.write "<br>" & request.form
Function IIf(condition,value1,value2)
	If condition Then IIf = value1 Else IIf = value2
End Function

ocorrencia="":tipomov=""

dim conexao, conexao2, chapach, rs, rs2, rs3, a(6)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if session("selecaoprofessor")="" then session("selecaoprofessor")="disciplina"
if request("idaula")="0" then session("id_grdaula")=request("idaula") : session("tipomov")="N"
if request("idaula")>"0" then session("id_grdaula")=request("idaula") : session("tipomov")="A"
if request("idturma")<>"" then session("id_grdturma")=request("idturma")
if request("codhor")<>"" then session("codhor")=request("codhor")
'response.write "Debug: id_grdaula: " & session("id_grdaula") & " - tipo: " & session("tipomov") & " - id_grdturma: " & session("id_grdturma") & " - codhor: " & session("codhor") & " |"
	tolerancia_dias_para_excluir=25 : dias=tolerancia_dias_para_excluir

if request.form<>"" then
	if request.form("bt_salvar")<>"" then
		tudook=0
		limite1=request.form("limite1"):if limite1="" then limite1=0 else limite1=cdbl(limite1)
		taulas1=request.form("taulas1"):if taulas1>=0 then taulas1=cdbl(taulas1)
		limite2=request.form("limite2"):if limite2>=0 then limite2=cdbl(limite2)
		taulas2=request.form("taulas2"):if taulas2>=0 then taulas2=cdbl(taulas2)

		limite1c=request.form("limite1c"):if limite1c="" then limite1c=0 else limite1c=cdbl(limite1c)
		taulas1c=request.form("taulas1c"):if taulas1c>=0 then taulas1c=cdbl(taulas1c)
		limite2c=request.form("limite2c"):if limite2c="" then limite2c=0 else limite2c=cdbl(limite2c)
		taulas2c=request.form("taulas2c"):if taulas2c>=0 then taulas2c=cdbl(taulas2c)
		
		if (request.form("chapa1")<>request.form("chapa1_o") and request.form("chapa1_o")<>"") _
		or (request.form("codmat")<>request.form("codmat_o") and request.form("codmat_o")<>"") _
		or (request.form("chapa2")<>request.form("chapa2_o") ) then mudou1=1 else mudou1=0

	'response.write taulas1c & "-" & limite1c & "/" & taulas2c & "-" & limite2c
		if (request.form("codmat")="" or request.form("codmat")="0") or (request.form("chapa1")="" or request.form("chapa1")="0") then
			tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Preencha todos os campos do cadastro!');</script>"
		end if
		if request.form("limite111")="1" or request.form("limite112")="1" then tudook=0
		if request.form("necessita")="ON" and (request.form("chapa2")="0" or request.form("chapa2")="") then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('Selecione o 2º professor ou desmarque o campo ""Necessita""!');</script>"
		if (taulas1c>=limite1c) or (taulas2c>=limite2c and taulas2c>0) then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('(1) O professor excede o limite de aulas permitido!');</script>"

		if request.form("chapa2")="" then chapa2="null" else chapa2="'" & request.form("chapa2") & "'":chapa2a=mid(chapa2,2,5)
		if request.form("chapa1")="" or request.form("chapa1")="0" then chapa1="99999" else chapa1=request.form("chapa1")
		if request.form("alunos")="" then alunos=0 else alunos=request.form("alunos")
		if request.form("quando")="" then quando=int(now()) else quando=request.form("quando")
		inicio=request.form("inicio")
		termino=request.form("termino")
		deletada=0:ativo=1
		if session("idjpai")<>"" then 'o idjpai é definido durante a checagem de ocorrências
			'lançamento - este registro vai juntar com outro (o jpai) - se ele ainda não é jpai, neste execute vira jpai 
			sqlpai="update g2aulas set jpai=1, juntar_id=null, jturma=null where id_grdaula=" & session("idjpai") & " and (jpai=0 or jpai is null) "
			conexao.execute sqlpai, , adCmdText
			juntar=1:juntar_id=session("idjpai")
		else
			juntar=0:juntar_id="null"
		end if
		if (request.form("chapa1_o")<>request.form("chapa1") or request.form("codmat")<>request.form("codmat_o")) and request.form("jpai")="True" then
			MsgMaster=MsgMaster & "<br>Mudança disciplina/professor"
		end if
		if request.form("jpai")="True" then
			sqlson="select id_grdaula, id_grdturma from g2aulas where juntar_id=" & session("id_grdaula")
			rs.Open sqlson, ,adOpenStatic, adLockReadOnly
			total=rs.recordcount
			do while not rs.eof
				if rs.absoluteposition=1 then 'primeiro registro: vira pai/passa turma pra os filhos
					sqlj1="select codtur from g2turmas where id_grdturma=" & rs("id_grdturma")
					rs2.Open sqlj1, ,adOpenStatic, adLockReadOnly
					if rs2.recordcount>0 then 
						n_jturma=rs2("codtur"):n_juntar_id=rs("id_grdaula")
					else
						n_jturma="null":n_juntar_id="null"
					end if
					rs2.close
					sqlj2="update g2aulas set jpai=1, juntar=0, jturma=null, juntar_id=null where id_grdaula=" & rs("id_grdaula")
					conexao.execute sqlj2, , adCmdText
				else  'para os registros subsequentes
					sqlj1="update g2aulas set juntar_id=" & n_juntar_id & ", jturma='" & n_jturma & "' where id_grdaula=" & rs("id_grdaula")
					conexao.execute sqlj1, , adCmdText
				end if
			rs.movenext:loop
			rs.close
			sqlj3="update g2aulas set jpai=0 where id_grdaula=" & session("id_grdaula")
			conexao.execute sqlj3, , adCmdText
		end if

		if mudou1=1 and session("tipomov")="A" and now()>=(cdate(inicio)+dias) or request.form("forcainsercao")="ON" then
			if request.form("inicio")=request.form("inicio_o") then inicio=int(now()) else inicio=request.form("inicio")
			response.write "<br>" & request.form("inicio") & "- " & request.form("inicio_o")
			termino0=cdate(inicio)-1
			tipomov="S"
			sqla="update g2aulas set termino='" & dtaccess(termino0) & "' where id_grdaula=" & session("id_grdaula")
			response.write "<br>" & sqla
			conexao.execute sqla, , adCmdText
		end if	

		'checar novos limites e totais
		sql3="select count(id_grdaula) as taulas from g2ch where chapa1='" & chapa1 & "' and ('" & dtaccess(inicio) & "' between inicio and termino) and deletada=0 and juntar=0 and ativo=1"
		rs3.Open sql3, ,adOpenStatic, adLockReadOnly:taulas1=rs3("taulas")
		rs3.close
		sql3="select count(id_grdaula) as taulas from g2ch where chapa1='" & chapa2a & "' and ('" & dtaccess(inicio) & "' between inicio and termino) and deletada=0 and juntar=0 and ativo=1"
		rs3.Open sql3, ,adOpenStatic, adLockReadOnly:if rs3.recordcount>0 then taulas2=rs3("taulas") else taulas2=0
		rs3.close

		sql1="select codds, codtn, horini from g2defhor where codhor=" & request.form("codhor")
		rs.Open sql1, ,adOpenStatic, adLockReadOnly:diasem=rs("codds"):if hour(rs("horini"))=>22 then adnot=1 else adnot=0
		rs.close
		if left(request.form("chapa1"),2)="99" then limitet1=999 else limitet1=20
		if left(request.form("chapa2"),2)="99" then limitet2=999 else limitet2=20
		if session("tipomov")="N" or tipomov="S" then
			'response.write "<br>Salvar tipo N or S"
			'------------------------------------------------------------------------
			sql1="insert into g2aulas (id_grdturma, codmat, diasem, codhor, adnot, chapa1, chapa2, codsala, inicio, termino, " & _
			"juntar, jturma, juntar_id, dividir, dturma, extra, demons, obs, " & iif(request.form("necessita")="ON","prof2, alunos, justificativa, autorizado, quando,","") & " excecao, usuarioc, datac) "
			sql2="select " & request.form("id_grdturma")&", '" & request.form("codmat") & "', " & diasem & ", " & request.form("codhor") & ", " & adnot & _
			", '" & chapa1 & "', " & chapa2 & ", '" & request.form("codsala") & "', '" & dtaccess(inicio) & "', '" & dtaccess(termino) & "', " & _
			iif(request.form("juntar")="ON",1,0) & ", '" & request.form("jturma") & "', " & juntar_id & ", " & iif(request.form("dividir")="ON",1,0) & ", " & _
			"'" & request.form("dturma") & "', " & iif(request.form("extra")="ON",1,0) & ", " & iif(request.form("demons")="ON",1,0) & ", " & _
			"'" & request.form("obs") & "', " & _
			iif(request.form("necessita")="ON","1, " & alunos & ", '" & request.form("justificativa") & "', " & iif(request.form("autorizado")="ON",1,0) & ", '" & dtaccess(quando) & "', ","") & _
			"" & iif(request.form("excecao")="ON",1,0) & ", " & _
			"'" & session("usuariomaster") & "', '" & dtaccess(now()) & "' "
			sql=sql1 & sql2		
			'response.write "<br><b>" & sql & "</b><br>"
			tudook=1
			if juntar=0 then taulas1=taulas1+1 : if request.form("necessita")="ON" then taulas2=taulas2+1
			if (taulas1c>limite1c) or (taulas2c>limite2c and taulas2c>0) or (taulas1>limitet1) or (taulas2>limitet2) then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('(2) O professor excede o limite de aulas permitido!');</script>"
			if session("usuariomaster")="02379" or session("master")=1 then tudook=1
			if tudook=1 then conexao.Execute sql, , adCmdText
		end if
		if session("tipomov")="A" and tipomov<>"S" then
			'response.write "<br>A"
			sql="update g2aulas set " & _
			"id_grdturma=" & session("id_grdturma") & ", codmat='" & request.form("codmat") & "' ,diasem=" & diasem & ", codhor=" & request.form("codhor") & _
			", adnot=" & adnot & ", chapa1='" & chapa1 & "', chapa2=" & chapa2 & ", codsala='" & request.form("codsala") & "' ,inicio='" & dtaccess(inicio) & _
			"', termino='" & dtaccess(termino) & "', juntar=" & juntar & ", jturma='" & request.form("jturma") & "', juntar_id=" & juntar_id & _
			", dividir=" & iif(request.form("dividir")="ON",1,0) & ", dturma='" & request.form("dturma") & "' ,extra=" & iif(request.form("extra")="ON",1,0) & _
			", demons=" & iif(request.form("demons")="ON",1,0) & ", obs='" & request.form("obs") & "' " & _
			iif(request.form("necessita")="ON", ", prof2=1, alunos=" & alunos & ", justificativa='" & request.form("justificativa") & "', autorizado=" & iif(request.form("autorizado")="ON",1,0) & ", quando='" & dtaccess(quando) & "' ", ", prof2=0, alunos=null, justificativa=null, autorizado=null, quando=null") & _
			", excecao=" & iif(session("usuariomaster")="02379",iif(request.form("excecao")="ON",1,0),0) & ""
			if session("usuariomaster")="02379" then usua="" else usua=", usuarioa='" & session("usuariomaster") & "'"
			sql=sql & usua & ", dataa='" & dtaccess(now()) & "' where id_grdaula=" & session("id_grdaula")
			'response.write "<br><b>" & sql & "</b><br>"
			tudook=1
			if request.form("chapa1")<>request.form("chapa1_o") and juntar=0 then taulas1=taulas1+1
			if request.form("chapa2")<>request.form("chapa2_o") and juntar=0 then taulas2=taulas2+1
			'response.write "--->" & request.form("chapa1")
			'response.write "--->" & request.form("chapa1_o")
			response.write "--->" & taulas1
			'if juntar=0 then taulas1=taulas1+1 :
			'if request.form("necessita")="ON" then taulas2=taulas2+1
			if (taulas1c>limite1c) or (taulas2c>limite2c and taulas2c>0) or (taulas1>limitet1) or (taulas2>limitet2) then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('(3) O professor excede o limite de aulas permitido!');</script>"
			if session("usuariomaster")="02379" or session("master")=1 then tudook=1
			if tudook=1 then conexao.Execute sql, , adCmdText
		end if

		'vez=session("outros_codhor")
		'for b=1 to cdbl(vez)
		'	gravar=request.form("gravar" & b)
		'	ocodhor=request.form("codhor" & b)
		'	if gravar="ON" then
		'		'response.write "<br>Gravar horarios varios"
		'		sqlo4="select codds, codtn, horini from g2defhor where codhor=" & ocodhor
		'		rs.Open sqlo4, ,adOpenStatic, adLockReadOnly:diasem=rs("codds"):if hour(rs("horini"))=>22 then adnot=1 else adnot=0
		'		rs.close
		'		sqlo1="insert into g2aulas (id_grdturma, codmat, diasem, codhor, adnot, chapa1, chapa2, codsala, inicio, termino, " & _
		'		"juntar, jturma, juntar_id, dividir, dturma, extra, demons, obs, " & iif(request.form("necessita")="ON","prof2, alunos, justificativa, autorizado, quando,","") & " excecao, usuarioc, datac) "
		'		sqlo2="select " & request.form("id_grdturma")&", '" & request.form("codmat") & "', " & diasem & ", " & ocodhor & ", " & adnot & _
		'		", '" & chapa1 & "', " & chapa2 & ", '" & request.form("codsala") & "', '" & dtaccess(inicio) & "', '" & dtaccess(termino) & "', " & _
		'		iif(request.form("juntar")="ON",1,0) & ", '" & request.form("jturma") & "', " & juntar_id & ", " & iif(request.form("dividir")="ON",1,0) & ", " & _
		'		"'" & request.form("dturma") & "', " & iif(request.form("extra")="ON",1,0) & ", " & iif(request.form("demons")="ON",1,0) & ", " & _
		'		"'" & request.form("obs") & "', " & _
		'		iif(request.form("necessita")="ON","1, " & alunos & ", '" & request.form("justificativa") & "', " & iif(request.form("autorizado")="ON",1,0) & ", '" & dtaccess(quando) & "', ","") & _
		'		"" & iif(request.form("excecao")="ON",1,0) & ", " & _
		'		"'" & session("usuariomaster") & "', '" & dtaccess(now()) & "' "
		'		sqlo3=sqlo1 & sqlo2		
		'		'response.write "<br><b>" & sqlo3 & "</b><br>"
		'	if juntar=0 then taulas1=taulas1+1 : if request.form("necessita")="ON" then taulas2=taulas2+1
		'		if (taulas1c>=limite1c) or (taulas2c>=limite2c and taulas2c>0) then tudook=0:response.write "<script language='JavaScript' type='text/javascript'>alert('(4) O professor excede o limite de aulas permitido!');</script>"
		'		if tudook=1 then conexao.Execute sqlo3, , adCmdText
		'	end if
		'next

'tudook=0
	end if 'button=salvar
else 'request.form=""
end if

if request.form("bt_excluir")<>"" then
	tudook=1
	inicio=request.form("inicio")
	sql="UPDATE g2aulas set deletada=1, usuarioa='" & session("usuariomaster") & "', dataa=getdate() WHERE id_grdaula=" & session("id_grdaula")
	if session("usuariomaster")="02379" then sql="UPDATE g2aulas set deletada=1 WHERE id_grdaula=" & session("id_grdaula")
	if now()>(cdate(inicio)+dias) then tudook=0:ocorrencia=ocorrencia & "<Br><font color=blue><b>As aulas já iniciaram. O lançamento não pode ser excluido!</b></font>"
	if session("usuariomaster")="02379" or session("master")=1 then tudook=1
	if tudook=1 then conexao.Execute sql, , adCmdText
	'
	'Ver quando excluir jpai=1 então zerar juntar=1 ou passar um quando count(juntar=1)>1 para jpai=1 
end if

if request.form("bt_selprof")<>"" then
	if session("selecaoprofessor")="disciplina" then session("selecaoprofessor")="todos" else session("selecaoprofessor")="disciplina"
end if
if session("selecaoprofessor")="disciplina" then txtbut="PROF." else txtbut="DISC."

if (request.form("bt_salvar")="" and request.form("bt_excluir")="") or (request.form("bt_salvar")<>"" and tudook=0) or (request.form("bt_excluir")<>"" and tudook=0) then
	'response.write request.form
	sql1="select * from g2aulas where id_grdaula=" & session("id_grdaula")
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	sql2="select t.*, c.tipocurso, inicio, termino from g2turmas t, g2cursos c, g2periodoaula p " & _
	"where p.perlet=t.perlet and t.codcur=c.codcur and t.codper=c.codper " & _
	"and id_grdturma=" & session("id_grdturma")
	rs1.Open sql2, ,adOpenStatic, adLockReadOnly

	necessita=0 : autorizado="":excecao=""
	if session("tipomov")="N" then
		'codsala=request.form("codsala")
		if request.form("codsala")="" then codsala=rs1("codsala") else codsala=request.form("codsala")
		codmat=request.form("codmat")
		chapa1=request.form("chapa1")
		chapa2=request.form("chapa2")
		if request.form("inicio") ="" then inicio =rs1("inicio")  else inicio =request.form("inicio")
		if request.form("termino")="" then termino=rs1("termino") else termino=request.form("termino")
		alunos       =request.form("alunos")
		justificativa=request.form("justificativa")
		quando       =request.form("quando")
		jturma       =request.form("jturma")
		rsextra      =null
		rsdemons     =null
	else
		if request.form("codsala")<>"" then codsala=request.form("codsala") else codsala=rs("codsala")
		if request.form("codmat") <>"" then codmat =request.form("codmat")  else codmat =rs("codmat")
		if request.form("chapa1") <>"" then chapa1 =request.form("chapa1")  else chapa1 =rs("chapa1")
		if request.form("chapa2") <>"" then chapa2 =request.form("chapa2")  else chapa2 =rs("chapa2")
		if request.form("inicio") <>"" then inicio =request.form("inicio")  else inicio =rs("inicio")
		if request.form("termino")<>"" then termino=request.form("termino") else termino=rs("termino")
		if request.form("alunos")=""        then alunos=rs("alunos") else alunos=request.form("alunos")
		if request.form("justificativa")="" then justificativa=rs("justificativa") else justificativa=request.form("justificativa")
		if request.form("quando")=""        then quando=rs("quando") else quando=request.form("quando")
		if request.form("jturma")<>"" then jturma=request.form("jturma") else jturma=rs("jturma")
		if rs("prof2")=true and request.form("necessita")="" then necessita=-1
		if rs("autorizado")=true and request.form("autorizado")="" then autorizado="checked"
		if rs("excecao")=true and request.form("excecao")="" then excecao="checked"
		'-------------------------
		if rs("juntar") =0 then juntar ="" else juntar ="checked"
		if rs("dividir")=0 then dividir="" else dividir="checked"
		if rs("extra")  =0 then extra  ="" else extra  ="checked"
		if rs("demons") =0 then demons ="" else demons ="checked"
		rsextra=rs("extra")
		rsdemons=rs("demons")
	end if
	if rs1("coddoc")="DIR" then filial=1 else filial=3
	if request.form("necessita")="ON" then necessita=-1
	if request.form<>"" and request.form("necessita")="" then necessita=0
	if request.form("autorizado")="ON" then autorizado="checked"
	if request.form("excecao")="ON" then excecao="checked"
	if request.form("juntar") ="ON" then juntar ="checked" 'else juntar =""
	if request.form("dividir")="ON" then dividir="checked" 'else dividir=""

	if session("tipomov")="N" then
		codmat_o=request.form("codmat_o")
		chapa1_o=request.form("chapa1_o")
		chapa2_o=request.form("chapa2_o")
		jpai    =request.form("jpai")
		inicio_o=request.form("inicio_o")
	else
		'codmat_o=rs("codmat")
		'chapa1_o=rs("chapa1")
		'chapa2_o=rs("chapa2")
		'jpai    =rs("jpai")
		'inicio_o=rs("inicio")
	end if
	
%>
<form method="POST" action="gradenovaaula.asp" name="form">
<input type="hidden" name="id_grdaula"  size="4" value="<%=session("id_grdaula")%>"> 
<input type="hidden" name="id_grdturma" size="4" value="<%=session("id_grdturma")%>">
<input type="hidden" name="codhor"      size="4" value="<%=session("codhor")%>">
<input type="hidden" name="codmat_o"    size="4" value="<%=codmat%>">
<input type="hidden" name="chapa1_o"    size="4" value="<%=chapa1%>">
<input type="hidden" name="chapa2_o"    size="4" value="<%=chapa2%>">
<input type="hidden" name="jpai"        size="5" value="<%=jpai%>">
<input type="hidden" name="inicio_o"    size="4" value="<%=inicio_o%>">
<input type="hidden" name="salvar" value="<%=request.form("salvar")%>">

<!-- quadro -->
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr><td class=campo valign=top width="80%">
<!-- quadro -->

<table border="0" cellpadding="3" cellspacing="0" width="100%">
	<tr><td class=grupo>Alteração de Grade Horária (<%=session("id_grdaula")%>)</td></tr>
</table>

<!-- Periodo / Serie / dia da semana -->
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo>Dia</td>
	<td class=titulo>Hora</td>
	<td class=titulo>Sala</td>
</tr>
<tr>
	<td class=fundo>
<%
sql1="select codds, descricao, horfim from g2defhor where codhor=" & session("codhor")
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
	diasem=rs2("codds"):diasem_ex=weekdayname(diasem)
	horario=rs2("descricao"):horfim=rs2("horfim")
else
	diasem="???":diasem_ex=""
	horario="???":horfim=""
end if
rs2.close
%>	
	<%=diasem_ex%>	</td>
	<td class=fundo><font color=blue><b><%=horario%></td>
	<td class=fundo><select class=small size="1" name="codsala" onfocus="javascript:window.status='Selecione a sala'">
<%
sqla="select sala AS codsala, saladesc, salacap, tipo from grades_esalas where codfilial=" & filial & " order by saladesc "
rs2.Open sqla, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
response.write "<option value=''>Selecione....</option>"
rs2.movefirst:do while not rs2.eof
if codsala=rs2("codsala") then tempd="selected" else tempd=""
%>
	<option value="<%=rs2("codsala")%>" <%=tempd%>><%=rs2("saladesc")%></option>
<%
rs2.movenext:loop
else
%>
	<option value="-1">Sem cadastro</option>
<%
end if
rs2.close
%>
	</select></td>

</tr>
</table>
  
<!-- Hora Inicio/Termino / Disciplina -->
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo>Disciplina</td>
</tr>
<tr>
	<td class=fundo><select size="1" name="codmat" onfocus="javascript:window.status='Selecione a disciplina'" onChange="javascript:submit()">
<%
sqla="select g.codmat, m.materia, g.naulassem, g.cargahoraria, g.tipo " & _
"from (select codcur, codper, grade, periodo, codmat, naulassem, cargahoraria, tipo='O' from corporerm.dbo.ugrade " & _
	"UNION " & _
		"select codcur, codper, grade, periodo, codmat, naulassem, cargahoraria, tipo from g2_ugrade where tipo='G'  ) g " & _
"inner join corporerm.dbo.umaterias m on g.codmat=m.codmat " & _
"where codcur=" & rs1("codcur") & " and codper=" & rs1("codper") & " and grade=" & rs1("grade") & " and periodo=" & rs1("serie") & _
" and g.codmat not in ('0') " & _
" order by g.tipo desc, m.materia"
rs2.Open sqla, ,adOpenStatic, adLockReadOnly
if rs2.recordcount>0 then
response.write "<option value='0'>Selecione....</option>"
rs2.movefirst:do while not rs2.eof
if codmat=rs2("codmat") then tempdi="selected" else tempdi=""
if rs2("tipo")<>"O" then tipograde="(*) " else tipograde=""
%>
	<option value="<%=rs2("codmat")%>" <%=tempdi%>><%=tipograde%> <%=rs2("materia")%></option>
<%
rs2.movenext:loop
else
%>
	<option value="-1">Sem disciplinas cadastradas</option>
<%
end if
rs2.close
%>
	</select>
	<a class=r href="hstdisciplina.asp?codmat=<%=codmat%>" onclick="NewWindow(this.href,'Pesquisa_disciplinas','545','200','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Magnify.gif" width="16" height="16" border="0" alt=""></a>
	</td>
</tr>
</table>

<!-- Chapa / Nome -->
<!--
 *
**
 *
 *
 *
***-->
<%
'if session("usuariomaster")="02379" then
'if codmat<>codmat_o then chapa1=""
if request.form("codmat")<>"" and request.form("codmat_o")<>"" and request.form("codmat")<>request.form("codmat_o") then chapa1=""

existiu=0
if codmat<>"" then
	sqlt1="select codtur, coddoc from g2turmas where id_grdturma=" & session("id_Grdturma")
	rs2.Open sqlt1, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then 
		codturma=rs2("codtur"):icurso=rs2("coddoc")
	else 
		codturma="":icurso=""
	end if
	rs2.close
	if codturma<>"" then
		sqlt2="select distinct chapa1 from g2ch g inner join corporerm.dbo.pfunc f on f.chapa collate database_default=g.chapa1 where codmat='" & codmat & "' and termino='" & dtaccess(cdate(inicio)-1) & "' and codtur='" & codturma & "' and codsituacao in ('A','F') " & _
		" and chapa1 not in (select chapa from g2demissao where tipo='E') "
		rs2.Open sqlt2, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then 
			existiu=1 
			if rs2.recordcount=1 then response.write "<input type=hidden name=chapasemestreanterior value=""" & rs2("chapa1") & """>"
			if rs2.recordcount=1 then existiutotal=1
		else 
			existiu=0
		end if
		rs2.close
	else
		sqlt2=""
	end if
	if icurso="DIR" then existiu=0
end if
'end if
%>
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>CH</td>
	<td class=titulo>Disp.</td>
</tr>
<tr>
	<td class=fundo>
<%if session("usuariomaster")="02379" or session("master")=1 then %>	
	<input type="text" value="<%=chapa1%>" name="chapa1" size="8" onfocus="javascript:window.status='Informe o chapa do professor'" onchange="chapa3()">
<%else%>
	<input type="hidden" value="<%=chapa1%>" name="chapa1"><%=chapa1%>
<%end if%>
	</td>
	<td class=fundo>&nbsp;
		<select size="1" name="nome1" onfocus="javascript:window.status='Selecione o Nome do Professor'" onchange="nome3()">
<%
toprec=500
if existiu=1 then 
	response.write session("id_grdturma")
	sqlt9="select perlet from g2turmas where id_grdturma=" & session("id_grdturma")
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	codperlet=rs2("perlet")
	rs2.close

	sqlt3=sqlt2 & " union all " & _
	"select chapa from g2excecoes where codmat='" & codmat & "' and codtur='" & codturma & "' and perlet='" & codperlet & "' "

	sql2="select chapa, nome, instrucaomec, tipo, rt, tab_ref, tab_grade from grades_aux_prof "
	sql2=sql2 & " where chapa in (" & sqlt3 & ") "
	'IF session("usuariomaster")="02379" Then response.write sql2

elseif session("selecaoprofessor")="disciplina" then
	sql2="select chapa, nome, instrucaomec, tipo, rt, tab_ref, tab_grade from grades_aux_prof "
	sql2=sql2 & " where chapa in (select chapa1 chapa from g2ch where codmat='" & codmat & "' and coddoc='" & rs1("coddoc") & "' group by chapa1) "
	sql2=sql2 & " union all " & _ 
	"select distinct a.chapa, nome='Aderencia: '+f.nome, instrucaomec, tipo, rt, tab_ref, tab_grade from grades_aderencia a inner join grades_aux_prof f on f.chapa=a.chapa " & _
	"where codmat='" & codmat & "' and coddoc='" & rs1("coddoc") & "' " & _
	"and a.chapa not in (select chapa1 chapa from g2ch where codmat='" & codmat & "' and coddoc='" & rs1("coddoc") & "' group by chapa1) "
else
	sql2="select chapa, nome, instrucaomec, tipo, rt, tab_ref, tab_grade from grades_aux_prof "
	sql2=sql2 & " order by nome "
end if
'response.write sql2
if codmat="" or codmat="0" then sql2="select chapa='99999',nome='SEM PROFESSOR',instrucaomec='Graduado',tipo='CLT',rt=null,tab_ref='D',tab_grade=5"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if existiutotal=1 and rs2.recordcount>0 then existiuexcecao=1 

response.write "<option value='0'>Selecione Professor 1....</option>"
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
if chapa1=rs2("chapa") then temp="selected" else temp=""
select case rs2("tab_grade")
	case "5"
		fundo="style=""background:FF9999;"""
	case "4"
		fundo="style=""background:FFCCCC;"""
	case "3"
		fundo="style=""background:FFFFCC;"""
	case "2"
		fundo="style=""background:CCFFCC;"""
	case "1"
		fundo="style=""background:CCFFCC;"""
	case else
		fundo="style=""background:FFFFFF;"""
end select
if rs2("rt")=rs2("chapa") then fundo="style=""background:CCFFCC;"""
if rs2("tipo")<>"CLT" then tipo="(*) " else tipo=""
%>
		<option <%=fundo%> value="<%=rs2("chapa")%>" <%=temp%>><%=tipo & rs2("nome") & " - (" & rs2("tab_grade") & ")"%></option>
<%
rs2.movenext:loop
end if 'rs2.recordcount>0
response.write "<option value='99" & rs1("coddoc") & "'># A CONTRATAR PARA O CURSO#</option>"
if (codmat="G1436" or codmad="G0006" or codmat="G0245" or codmat="G0508" or codmat="G1573") _
	or ( (codmat="G0002" or codmat="G0050" or codmat="G0001" or codmat="G0052" or codmat="G0471") and ( rs1("coddoc")<>"LET" or rs1("coddoc")<>"CSJ") ) then
	response.write "<option value='99998'># E.A.D. #</option>"
end if

%>
	</select>
	
	<%if existiu<>1 then%>
	<input type="submit" name="bt_selprof" value="<%=txtbut%>" class="button" alt="Mostrar apenas professores da disciplina" onclick="javascript:submit()">
	<%end if%>
	</td>
	<td class=fundo align="left" valign="top">&nbsp;
<%
if existiuexcecao=1 then response.write "<input type=hidden name=chapasemestreanterior value=''>"
sql3="select count(id_grdaula) as taulas from g2ch where chapa1='" & chapa1 & "' and ('" & dtaccess(inicio) & "' between inicio and termino) and deletada=0 and juntar=0 and ativo=1"
rs3.Open sql3, ,adOpenStatic, adLockReadOnly:taulas1=rs3("taulas")
rs3.close
sql3="select count(id_grdaula) as taulas from g2ch where coddoc='" & rs1("coddoc") & "' and chapa1='" & chapa1 & "' and ('" & dtaccess(inicio) & "' between inicio and termino) and deletada=0 and juntar=0 and ativo=1"
rs3.Open sql3, ,adOpenStatic, adLockReadOnly:taulas1c=rs3("taulas")
rs3.close
%>	
<%if taulas1>0 or taulas1c>0 then %>
<a class=r href="hstaulas.asp?chapa=<%=chapa1%>&inicio=<%=inicio%>" onclick="NewWindow(this.href,'Aulas_atribuidas','545','200','yes','center');return false" onfocus="this.blur()">
<%end if%> 
<%=taulas1%> aulas (<% response.write rs1("coddoc")& ": " & taulas1c%>)
<%if taulas1>0 or taulas1c>0 then %> </a> <%end if%>
	</td>
	<td class=fundo align="center">
<a class=r href="hstdisp.asp?chapa=<%=chapa1%>&inicio=<%=inicio%>" onclick="NewWindow(this.href,'Aulas_disponiveis','545','200','yes','center');return false" onfocus="this.blur()">
<img src="../images/clock.gif" alt="" border="0" title="Clique para ver disponibilidade">
</a>
	</td>	
</tr>

<!-- inicio consistencias professor 1 -->
<tr><td class=fundo colspan=4>
<%
if chapa1<>"" then
	ocorrencia=ocorrencia & "<font color=black><b>Professor 1: </b></font>"
	'------------- verifica se existe aula no mesmo horário para juntar --------------------
	sql9="select t.codtur, m.materia, h.horini, a.inicio, a.termino, a.id_grdaula from g2aulas a, g2turmas t, corporerm.dbo.umaterias m, g2defhor h, " & _
	"(select horini, codds from g2defhor where codhor=" & session("codhor") & ") ch " & _
	"where chapa1='" & chapa1 & "' and '" & dtaccess(inicio) & "' between a.inicio and a.termino and a.id_grdturma<>" & session("id_grdturma") & " and deletada=0 and juntar=0 " & _
	"and h.horini=ch.horini and h.codds=ch.codds and a.id_grdturma=t.id_grdturma and m.codmat collate database_default=a.codmat and a.codhor=h.codhor order by t.codtur  "
	rs3.Open sql9, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then
		'do while not rs3.eof
		ocorrencia=ocorrencia & "<font color=red>Neste horário está em <b>" & rs3("materia") & "</b> na turma <b>" & rs3("codtur") & "." & "</b>"
		session("idjpai")=rs3("id_grdaula") : 		jturma=rs3("codtur")
		'rs3.movenext:loop
		ocorrencia=ocorrencia & "<br><font color=blue>A aula será juntada com a outra turma.</font>"
		juntar="checked"
	else
		session("idjpai")="" : juntar="" : jturma=""
	end if 
	rs3.close
	'------------- verifica se é professor habitual da matéria / se não for, se vai ganhar mais que o do período anterior
	sql10="select coddoc, chapa1 chapa from g2ch where deletada=0 " & _
	"and codmat='" & codmat & "' and coddoc='" & rs1("coddoc") & "' and inicio<'" & dtaccess(inicio) & "' and chapa1='" & chapa1 & "' group by coddoc, chapa1 "
	rs3.Open sql10, ,adOpenStatic, adLockReadOnly:existe=rs3.recordcount
	rs3.close
	if existe=0 then
		ocorrencia=ocorrencia & "<br><br><font color=red>Professor não habitual."
		sql11="select f.chapa chapa, v.valoraula valor from dc_professor f, csd_cursos c, csd_titulos t, csd_faixas v " & _
		"where c.tabela=t.tabela and t.titulacao=f.tab_instr collate database_default and t.nivel=f.codnivelsal collate database_default and t.reformulacao=f.tab_ref " & _
		"and t.faixasalarial=v.faixasalarial and v.dt_faixa in (select max(dt_faixa) from csd_faixas where dt_faixa<getdate()) " & _
		"and c.coddoc='" & rs1("coddoc") & "' and f.chapa='" & chapa1 & "' and getdate() between ivigencia and fvigencia "
		'response.write "<br>iniciou query " & now()
		if chapa1<>"99998" and chapa1<>"99999" then
			rs3.Open sql11, ,adOpenStatic, adLockReadOnly:
			if rs3.recordcount>0 then salprof=rs3("valor") else salprof=0
			rs3.close
		end if
		'response.write "<br>finalizou query " & now()
		'response.write "<br>iniciou query " & now()
		sql12="select min(valormin) valormin, max(valormax) valormax from g2comparasal where codmat='" & codmat & "' and coddoc='" & rs1("coddoc") & "'"
		rs3.Open sql12, ,adOpenStatic, adLockReadOnly
		if rs3.recordcount>0 then
			if salprof>rs3("valormax") then ocorrencia=ocorrencia & "<br><br>Este professor ganha <b>mais</b> que os habituais. Escolher outro.":inconsistencia="Inc.: Professor ganha + que os habituais."
			if salprof<rs3("valormin") then ocorrencia=ocorrencia & "<br><br>Este professor ganha <b><u>menos</u></b> que os habituais."
			if salprof=rs3("valormin") then ocorrencia=ocorrencia & "<br><br>Este professor ganha <b><u>igual</u></b> aos menores salários dos habituais."
		else
			ocorrencia=ocorrencia & "<Br><br>Não existe professores recebendo ainda.":inconsistencia=""
		end if
		rs3.close			
		'response.write "<br>finalizou query " & now()
	else
		ocorrencia=ocorrencia & "<br><br><font color=red>Professor habitual":inconsistencia="" 
	end if
	'----------------- mudando professor / verificar data de inicio das aulas ----------------------
	if chapa1_o<>chapa1 then 
		'ocorrencia=ocorrencia & "<br><br><font color=green>Mudando professor: " & chapa1_o & " para " & chapa1
	end if
	'----------------- limite de aulas ----------------------------
	tmpcur=rs1("coddoc")
	stringcurso=" in ('" & tmpcur & "','---') "
	if tmpcur="EGA" or tmpcur="EGC" or tmpcur="EGP" or tmpcur="EGT" or tmpcur="ECI" then stringcurso=" in ('CCO','EGA','EGC','EGP','EGT','ECI','---') "
	if tmpcur="AEM" or tmpcur="AMK" then stringcurso=" in ('AEM','AMK','---') "
	if tmpcur="TMK" or tmpcur="TGC" then stringcurso=" in ('TMK','TGC','---') "
	
	limite1=0:limite1c=0
	sql13="select limite=sum(limite) from g2limite where chapa='" & chapa1 & "' and coddoc " & stringcurso & " "
	rs3.Open sql13, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then limite1c=rs3("limite") 'else limite1=20
	rs3.close
	if left(chapa1,2)="99" then limite1c=999
	ocorrencia=ocorrencia & "<br><br><font color=gray>Limite de aulas " & rs1("coddoc") & ": " & limite1c & "</font>"
	'------------
	sql14="select chapa from g2disp where chapa='" & chapa1 & "' and codhor=" & session("codhor") & ""
	rs3.Open sql14, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then disponivel="OK" else disponivel="Horário indisponivel"
	rs3.close
	ocorrencia=ocorrencia & "<br><br><font color=blue>Disponibilidade: " & disponivel & "</font>"
	
	'------------------- checar limite de 11 horas --------------------------------
	if horfim<="09:10" then
		sqllim11="select id_grdaula, codtur, diasem, DESCRICAO from g2ch where chapa1='" & chapa1 & "' and diasem=" & diasem-1 & " and inicio='" & dtaccess(inicio) & "' and horfim>=convert(time,dateadd(hh,-11,'" & horfim & "')) and deletada=0 and juntar=0 "
		rs3.Open sqllim11, ,adOpenStatic, adLockReadOnly
		if rs3.recordcount>0 and left(chapa1,2)<>"99" then 
			ocorrencia=ocorrencia & "<br><br><font color=red>Existe aulas atribuidas dentro do limite de 11 horas na turma " & rs3("codtur") & " no horário das " & rs3("descricao") & "."
			limite111=1
		else
			limite111=0
		end if
		rs3.close
	elseif horfim>="20:50" then
		sqllim11="select id_grdaula, codtur, diasem, DESCRICAO from g2ch where chapa1='" & chapa1 & "' and diasem=" & diasem+1 & " and inicio='" & dtaccess(inicio) & "' and horini<=convert(time,dateadd(hh,+11,'" & horfim & "')) and deletada=0 and juntar=0 "
		rs3.Open sqllim11, ,adOpenStatic, adLockReadOnly
		if rs3.recordcount>0 and left(chapa1,2)<>"99" then 
			ocorrencia=ocorrencia & "<br><br><font color=red>Existe aulas atribuidas dentro do limite de 11 horas na turma " & rs3("codtur") & " no horário das " & rs3("descricao") & "."
			limite111=1
		else
			limite111=0
		end if
		rs3.close
	else 
		limite111=0
	end if

	end if
%>
<input type="hidden" name="limite111" value="<%=limite111%>">
<input type="hidden" name="limite1c" value="<%=limite1c%>">
<input type="hidden" name="taulas1c" value="<%=taulas1c%>">
<input type="hidden" name="obs" value="<%=inconsistencia%>">
</td></tr></table>
<!-- final consistencias professor 1-->

<%
if necessita=-1 then txtneed="checked" else txtneed=""
'*************** 2 professor *************************
%>
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr><td class=titulo colspan=3><input type="checkbox" name="necessita" value="ON" <%=txtneed%> onClick="javascript:submit()">
<font color="red">Necessita de 2º Professor?</td></tr>
<%
if necessita=-1 then
%>
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
	<td class=titulo>CH</td>
</tr>
<tr>
	<td class=fundo>
<%if session("usuariomaster")="02379" or session("master")=1 then %>	
	<input type="text" value="<%=chapa2%>" name="chapa2" size="8" onfocus="javascript:window.status='Informe o chapa do professor'" onchange="chapa4()">
<%else%>
	<input type="hidden" value="<%=chapa2%>" name="chapa2"><%=chapa2%>
<%end if%>
	</td>

	<td class=fundo>&nbsp;
		<select size="1" name="nome2" onfocus="javascript:window.status='Selecione o Nome do Professor'" onchange="nome4()" >
<%
response.write "<option value='0'>Selecione Professor 2....</option>"
if rs2.recordcount>0 then
rs2.movefirst:do while not rs2.eof
if chapa2=rs2("chapa") then temp="selected" else temp=""
%>
		<option value="<%=rs2("chapa")%>" <%=temp%>><%=rs2("nome") & " - (" & rs2("tab_grade") & ")"%></option>
<%
rs2.movenext:loop
end if
rs2.close
%>
		</select></td>
	<td class=fundo align="left">
<%
sql3="select count(id_grdaula) as taulas from g2ch where chapa1='" & chapa2 & "' and ('" & dtaccess(inicio) & "' between inicio and termino) and deletada=0 and juntar=0 and ativo=1"
rs3.Open sql3, ,adOpenStatic, adLockReadOnly:taulas2=rs3("taulas"):rs3.close
sql3="select count(id_grdaula) as taulas from g2ch where coddoc='" & rs1("coddoc") & "' and chapa1='" & chapa2 & "' and ('" & dtaccess(inicio) & "' between inicio and termino) and deletada=0 and juntar=0 and ativo=1"
rs3.Open sql3, ,adOpenStatic, adLockReadOnly:taulas2c=rs3("taulas"):rs3.close
%>	
<%if taulas2>0 or taulas2c>0 then %>
<a class=r href="hstaulas.asp?chapa=<%=chapa2%>&inicio=<%=inicio%>" onclick="NewWindow(this.href,'Aulas_atribuidas','545','200','yes','center');return false" onfocus="this.blur()">
<%end if%>
<%=taulas2%> aulas (<% response.write rs1("coddoc")& ": " & taulas2c%>)
<%if taulas2>0 or taulas2c>0 then %>
</a>
<%end if%>
	</td>	
</tr>

<!-- inicio consistencias professor 2 -->
<tr><td class=fundo colspan=3>
<%
if chapa2<>"" then
	ocorrencia=ocorrencia & "<hr><font color=black><b>Professor 2: </b></font>"
	'------------- verifica se existe aula no mesmo horário para juntar --------------------
	sql9="select t.codtur, m.materia, h.horini, a.inicio, a.termino, a.id_grdaula from g2aulas a, g2turmas t, corporerm.dbo.umaterias m, g2defhor h, " & _
	"(select horini, codds from g2defhor where codhor=" & session("codhor") & ") ch " & _
	"where chapa2='" & chapa2 & "' and '" & dtaccess(inicio) & "' between a.inicio and a.termino and a.id_grdturma<>" & session("id_grdturma") & " and deletada=0 and juntar=0 " & _
	"and h.horini=ch.horini and h.codds=ch.codds and a.id_grdturma=t.id_grdturma and m.codmat collate database_default=a.codmat and a.codhor=h.codhor "
	rs3.Open sql9, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then
		do while not rs3.eof
		ocorrencia=ocorrencia & "<font color=red>Neste horário está em <b>" & rs3("materia") & "</b> na turma <b>" & rs3("codtur") & "." & "</b>"
		session("idjpai")=rs3("id_grdaula") : 		jturma=rs3("codtur")
		rs3.movenext:loop
		ocorrencia=ocorrencia & "<br><br><font color=blue>A aula será juntada com a outra turma.</font>"
		juntar="checked"
	else
		session("idjpai")="" : juntar="" : jturma=""
	end if 
	rs3.close
	'------------- verifica se é professor habitual da matéria / se não for, se vai ganhar mais que o do período anterior
	sql10="select coddoc, chapa1 chapa from g2ch where deletada=0 " & _
	"and codmat='" & codmat & "' and coddoc='" & rs1("coddoc") & "' and inicio<'" & dtaccess(inicio) & "' and chapa1='" & chapa2 & "' group by coddoc, chapa1 "
	rs3.Open sql10, ,adOpenStatic, adLockReadOnly:existe=rs3.recordcount:rs3.close
	if existe=0 then
		ocorrencia=ocorrencia & "<br><br><font color=red>Professor não habitual."
		sql11="select f.chapa chapa, v.valoraula valor from dc_professor f, csd_cursos c, csd_titulos t, csd_faixas v " & _
		"where c.tabela=t.tabela and t.titulacao=f.tab_instr collate database_default and t.nivel=f.codnivelsal collate database_default and t.reformulacao=f.tab_ref " & _
		"and t.faixasalarial=v.faixasalarial and v.dt_faixa in (select max(dt_faixa) from csd_faixas where dt_faixa<getdate()) " & _
		"and c.coddoc='" & rs1("coddoc") & "' and f.chapa='" & chapa2 & "' and getdate() between ivigencia and fvigencia "
		'response.write "<br>iniciou query " & now()
		if chapa2<>"99998" and chapa2<>"99999" then
			rs3.Open sql11, ,adOpenStatic, adLockReadOnly:
			if rs3.recordcount>0 then salprof=rs3("valor") else salprof=0:
			rs3.close
		end if
		'response.write "<br>finalizou query " & now()
		'response.write "<br>iniciou query " & now()
		sql12="select min(valormin) valormin, max(valormax) valormax from g2comparasal where codmat='" & codmat & "' and coddoc='" & rs1("coddoc") & "'"
		rs3.Open sql12, ,adOpenStatic, adLockReadOnly
		if rs3.recordcount>0 then
			if salprof>rs3("valormax") then ocorrencia=ocorrencia & "<br><br>Este professor ganha <b>mais</b> que os habituais. Escolher outro.":inconsistencia="Inc.: Professor ganha + que os habituais."
			if salprof<rs3("valormin") then ocorrencia=ocorrencia & "<br><br>Este professor ganha <b><u>menos</u></b> que os habituais."
			if salprof=rs3("valormin") then ocorrencia=ocorrencia & "<br><br>Este professor ganha <b><u>igual</u></b> aos menores salários dos habituais."
		else
			ocorrencia=ocorrencia & "<Br><br>Não existe professores recebendo ainda.":inconsistencia=""
		end if
		rs3.close			
		'response.write "<br>finalizou query " & now()
	else
		ocorrencia=ocorrencia & "<br><br><font color=red>Professor habitual":inconsistencia="" 
	end if
	'----------------- mudando professor / verificar data de inicio das aulas ----------------------
	if chapa2_o<>chapa2 then 
		'ocorrencia=ocorrencia & "<br><br><font color=green>Mudando professor: " & chapa2_o & " para " & chapa2
	end if
	'----------------- limite de aulas ----------------------------
	tmpcur=rs1("coddoc")
	stringcurso=" in ('" & tmpcur & "','---') "
	if tmpcur="EGA" or tmpcur="EGC" or tmpcur="EGP" or tmpcur="EGT" or tmpcur="ECI" then stringcurso=" in ('CCO','EGA','EGC','EGP','EGT','ECI','---') "
	if tmpcur="AEM" or tmpcur="AMK" then stringcurso=" in ('AEM','AMK','---') "
	if tmpcur="TMK" or tmpcur="TGC" then stringcurso=" in ('TMK','TGC','---') "

	limite2=0:limite2c=0
	sql13="select limite=sum(limite) from g2limite where chapa='" & chapa2 & "' and coddoc " & stringcurso & " "
	rs3.Open sql13, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then limite2c=rs3("limite") 'else limite2=20
	rs3.close
	if left(chapa2,2)="99" then limite2c=999
	ocorrencia=ocorrencia & "<br><br><font color=gray>Limite de aulas " & rs1("coddoc") & ": " & limite2c & "</font>"
	'------------------------
	sql14="select chapa from g2disp where chapa='" & chapa2 & "' and codhor=" & session("codhor") & ""
	rs3.Open sql14, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then disponivel2="OK" else disponivel2="Horário indisponivel"
	rs3.close
	ocorrencia=ocorrencia & "<br><br><font color=blue>Disponibilidade: " & disponivel2 & "</font>"

	'------------------- checar limite de 11 horas --------------------------------
	if horfim<="09:10" then
		sqllim11="select id_grdaula, codtur, diasem, DESCRICAO from g2ch where chapa1='" & chapa2 & "' and diasem=" & diasem-1 & " and inicio='" & dtaccess(inicio) & "' and horfim>=convert(time,dateadd(hh,-11,'" & horfim & "')) and deletada=0 and juntar=0 "
		rs3.Open sqllim11, ,adOpenStatic, adLockReadOnly
		if rs3.recordcount>0 and left(chapa2,2)<>"99" then 
			ocorrencia=ocorrencia & "<br><br><font color=red>Existe aulas atribuidas dentro do limite de 11 horas na turma " & rs3("codtur") & " no horário das " & rs3("descricao") & "."
			limite112=1
		else
			limite112=0
		end if
		rs3.close
	else
		limite112=0
	end if
	
end if
%>
<input type="hidden" name="limite112" value="<%=limite112%>">
<input type="hidden" name="limite2c" value="<%=limite2c%>">
<input type="hidden" name="taulas2c" value="<%=taulas2c%>">
<input type="hidden" name="obs" value="<%=inconsistencia%>">
</td></tr>
<!-- final consistencia professor 2 -->

<tr>
	<td class=titulo>Alunos</td>
	<td class=titulo>Justificativa</td>
	<td class=titulo>Autorizado</td>
</tr>
<tr>
	<td class=fundo><input type="text" value="<%=alunos%>" name="alunos" size="3" onfocus="javascript:window.status='Informe o numero de alunos da turma'"></td>
	<td class=fundo><textarea name="justificativa" cols="42" rows="3"><%=justificativa%></textarea></td>
<%if session("usuariomaster")="02379" or session("master")=1  then%>
	<td class=fundo><input type="checkbox" name="autorizado" value="ON" <%=autorizado%>>
	<input type="text" name="quando" size="8" value="<%=quando%>">
	</td>
<%else%>
	<td class=fundo>&nbsp;</td>
<%end if%>
</tr>
<%
else '**2 prof**
	chapa2="":alunos=0:justificativa="":quando=""
end if
'******************************** 2 professor     **************
%>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulor>Início </td>
	<td class=titulor>Término</td>
	<td class=titulor>Junta turmas</td>
	<td class=titulor>Divide turmas</td>
<% if session("usuariomaster")="02379" or session("master")=1 then %>
	<td class=titulor>A.Extra</td>
	<td class=titulor>Demonstr.</td>
	<td class=titulor>Exc.</td>
<%else%>
	<td class=titulor colspan=3></td>
<% end if %>
</tr>
<tr>
	<td class=fundo><input type="text" name="inicio" size="12" value="<%=inicio%>" onchange="javascript:submit()"></td>
	<td class=fundo><input type="text" name="termino" size="12" value="<%=termino%>"></td>
	<td class=fundo><input type="checkbox" name="juntar" value="ON" <%=juntar%>  onfocus="this.blur()" > <input type="text" name="jturma" size="5" value="<%=jturma%>"></td>
	<td class=fundo><input type="checkbox" name="dividir" value="ON" <%=dividir%>></td>
<% if session("usuariomaster")="02379" or session("master")=1 then %>
	<td class=fundo><input type="checkbox" name="extra" value="ON" <%=extra%>></td>
	<td class=fundo><input type="checkbox" name="demons" value="ON" <%=demons%>></td>
	<td class=fundo><input type="checkbox" name="excecao" value="ON" <%=excecao%>></td>
<% else %>
    <td class=fundor><%if rsextra=1 then response.write "<font face='Wingdings'>ü</font>" %></td>
    <td class=fundor><%if rsdemons=1 then response.write "<font face='Wingdings'>ü</font>" %></td>
	<td class=fundor></td>
<% end if %>
</tr>
</table>

<table border="0" cellpadding="1" cellspacing="0" width="100%">
<%
'if session("tipomov")<>"A" or session("idjpai")="" then
if desabilitado then
%>
<tr><td class=grupo colspan=1>Outros horários</td></tr>
<tr><td class=fundo colspan=1>
	<b>Duplicar o lançamento para os seguintes horários:</b><br>
<%
sqlh="select h1.* from g2defhor h1, (select * from g2defhor where codhor=" & session("codhor") & ") h2 " & _
"where h2.codds=h1.codds and h2.codtn=h1.codtn and h2.tipocurso=h1.tipocurso and h1.pos>h2.pos " & _
"and h1.codhor not in (select codhor from g2aulas where id_grdturma=" & session("id_grdturma") & " and deletada=0 ) "
rs3.Open sqlh, ,adOpenStatic, adLockReadOnly
vezes=1
if rs3.recordcount>0 then
	do while not rs3.eof
		response.write "<input type='checkbox' name='gravar" & vezes &"' value='ON'>"
		response.write "<input type='hidden' name='codhor" & vezes & "' value='" & rs3("codhor") & "'>"
		response.write rs3("descricao")
		vezes=vezes+1
	rs3.movenext:loop
	session("outros_codhor")=vezes-1
end if
rs3.close
%>
	</td></tr>
<%
end if

if session("usuariomaster")="02379" or session("master")=1 then %>
<tr><td class=fundo height=5 style="margin-bottom:4px dashed #000000"><hr>
<input type="checkbox" name="forcainsercao" value="ON">Forçar novo lançamento
</td></tr>
<%
end if
%>
</table>
  
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr>
	<td class=titulo align="center">
	<%'if disponivel="Horário indisponivel" or disponivel2="Horário indisponivel" then%>
	<%if disponivel="" then%>
		<font color=red>Não pode salvar neste horário</font>
	<%elseif existiu=1 and existiutotal<>1 and request.form("chapasemestreanterior")<>"" and request.form("chapasemestreanterior")<>request.form("chapa1") and left(request.form("chapa1"),2)<>"99" then%>
		<font color=red>O professor desta disciplina não pode ser alterado.</font>
	<%elseif limite111=1 or limite112=1 then%>
		<font color=red>Não pode ser incluído em razão do horário.</font>
	<%else%>
		<input type="submit" value="Salvar Registro" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique aqui para salvar'">
	<%end if%>
	<%if ( (disponivel="Horário indisponivel" or disponivel2="Horário indisponivel") or (existiu=1 and request.form("chapasemestreanterior")<>request.form("chapa1")) or (limite111=1 or limite112=1) ) _
		and (session("usuariomaster")="02379" or session("master")=1) then%>
		<input type="submit" value="Salvar Exceção" class="button" name="bt_salvar" onfocus="javascript:window.status='Clique para salvar'">
	<%end if%>
		</td>
	<td class=titulo align="center"><input type="reset"  value="Desfazer" class="button" name="B2"></td>
	<td class=titulo align="center"><input type="submit" value="Excluir registro" class="button" name="Bt_excluir"></td>
</tr>
</table>

<!-- quadro -->
</td><td class=campo valign=top width="20%">
<!-- quadro -->

<table border="0" cellpadding="3" cellspacing="0" width="100%">
	<tr><td class=grupo>Ocorrências</td></tr>
	<tr><td class=campo>
	<%
	response.write ocorrencia
	%>
	</td></tr>
</table>

<!-- quadro para outros horarios -->
</td></tr>

</table>
<!-- quadro -->
</form>
<%
rs.close
set rs=nothing
end if

response.write MsgMaster

if request.form("bt_salvar")<>"" or request.form("bt_excluir")<>"" then
	if tudook=1 then
		'response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi alterado!');window.opener.location=window.opener.location;self.close();</script>"
        response.write "<script language='JavaScript' type='text/javascript'>alert('O Lançamento foi gravado!');window.opener.document.form.submit();self.close();</script>"
	else
		response.write "<script language='JavaScript' type='text/javascript'>alert('O lançamento Não pode ser alterado!');</script>"
	end if
	'Response.write "<p>Registro atualizado.<br>"
	'response.write "<a href='javascript:top.window.close()'>Fechar Janela</a>"
%>
<!--
<script language="Javascript">javascript:window.opener.location=window.opener.location</script>
<form>
<input type="button" value="Fechar" class="button" onClick="top.window.close()">
</form>
-->
<%
end if

conexao.close
set conexao=nothing
%>
</body>
</html>