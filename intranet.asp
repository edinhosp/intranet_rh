<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
response.cookies("vrh06")("firstlogin")="N"
response.cookies("vrh06").expires=dateadd("m",3,now)
%>
<%
firstlogin=request.cookies("vrh06")("firstlogin")
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("cmdLogin")<>"" then
	Session("DescricaoErro")="Usu�rio n�o cadastrado"
	loginmaster

Sub LoginMaster
	usuario1=ucase(request.form("txtUsuario"))
	senha1=request.form("txtpassword")
		usuario1=replace( request.form("txtUsuario") ,"--","")
		senha1=replace( request.form("txtpassword") ,"--","")
		usuario1=replace( usuario1 ,"=","")
		senha1=replace( senha1 ,"=","")
	sql="SELECT * From usuarios where ucase(usuario)='" & usuario1 & "'"
	sql="SELECT * From usuarios where usuario='" & usuario1 & "' and ativo=1"
	'response.write sql
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 
		usuariof=rs("usuario"):senha=rs("password")
		Session("DescricaoErro")=""
		'testar se confere senha e usuario
		sqlrm="select chapa, codsituacao from corporerm.dbo.pfunc where chapa='" & usuario1 & "'"
		rs2.Open sqlrm, ,adOpenStatic, adLockReadOnly
		if rs2.recordcount>0 then
			situ=rs2("codsituacao")
			if situ="A" or situ="E" or situ="F" or situ="Z" then permitido=1 else permitido=0
		end if
		if usuario1="02090" then permitido=1
		if usuario1="02653" then permitido=1
		IF USUARIO1="99999" THEN PERMITIDO=1
		rs2.close
		
		if usuariof=usuario1 and senha=senha1 and permitido=1 then
			Session("acesso")=1
			rp=request.cookies("vrh06")("registropagina")
			if rp<>"" then Session("RegistrosPorPagina")=rp else Session("registrosporpagina")=20
			cb=request.cookies("vrh06")("cabecalho")
			if cb="sim" then Session("cabecalho")="sim" else Session("cabecalho")="nao"
			Session("Usuarioname") =rs("nome")
			Session("usuariogrupo")=rs("grupo")
			'Session("grant_docens")=rs("docens")
			'Session("grant_ifip")  =rs("ifip")
			'Session("grant_rh")    =rs("rh")
			'Session("grant_curso") =rs("curso")
			'Session("grant_menu")  =rs("menu")
			Session("estilo")      =rs("estilo")
			if rs("master")=true then emaster=1 else emaster=0
			Session("master")      =emaster
			SendIp=request.servervariables("LOCAL_ADDR")
			SendIp=request.servervariables("REMOTE_ADDR")
			Session("UsuarioMaster")=ucase(usuario1)
			if Session("usuariomaster")<>"" then
			sqlz="INSERT INTO login ( usuario, entrada, sessao, ipcomp ) SELECT '" & usuario1 & "' AS Expr1, getdate() AS Expr2," & _
			" '" & Session.Sessionid & "' AS Expr3, '" & sendip & "';"
			conexao.Execute sqlz
			end if
			temp=rs("timeout"): if isnumeric(temp) then Session.timeout=temp else Session.timeout=20
			rs.close
			sqlu="select top 2 entrada from login where usuario='" & usuario1 & "' order by entrada desc "
			rs.Open sqlu, ,adOpenStatic, adLockReadOnly
			if rs.recordcount>1 then
				rs.movenext
				Session("lastacesso")=rs("entrada")
			else
				Session("lastacesso")="-.-"
			end if
			rs.close
			Session("DescricaoErro")=""
		else
			Session("DescricaoErro")="Usuario n�o cadastrado/Senha n�o confere"
		end if
	else 'rs.recordcount=0
		'teste professor
		rs.close
		sql="select chapa, nome, apelido, cartidentidade, cartmodelo19, codsituacao from dc_professor where codsituacao<>'D' and chapa='" & usuario1 & "'"
		rs.Open sql, ,adOpenStatic, adLockReadOnly
		if rs.recordcount>0 then
			if (rs("cartidentidade")="" or isnull(rs("cartidentidade"))) and rs("cartmodelo19")<>"" then ident=rs("cartmodelo19") else ident=rs("cartidentidade")
			senhaf=left(textopuro(ident,3),4)
			usuariof=rs("chapa"):senha=senhaf
			Session("DescricaoErro")=""
'-----------------------------------------------------
			if usuariof=usuario1 and senha=senha1 then
				Session("acesso")=2
				rp=request.cookies("vrh06")("registropagina")
				if rp<>"" then Session("RegistrosPorPagina")=rp else Session("registrosporpagina")=20
				cb=request.cookies("vrh06")("cabecalho")
				if cb="sim" then Session("cabecalho")="sim" else Session("cabecalho")="nao"
				Session("Usuarioname") =rs("apelido")
				Session("usuariogrupo")="PROFESSOR"
				'Session("grant_menu")  ="100"
				Session("a100")="T"
				Session("estilo")      ="diversos.css"
				SendIp=request.servervariables("LOCAL_ADDR")
				SendIp=request.servervariables("REMOTE_ADDR")
				Session("UsuarioMaster")=ucase(usuario1)
				if Session("usuariomaster")<>"" then
				sqlz="INSERT INTO login ( usuario, entrada, sessao, ipcomp ) SELECT '" & usuario1 & "' AS Expr1, getdate() AS Expr2," & _
				" '" & Session.Sessionid & "' AS Expr3, '" & sendip & "';"
				conexao.Execute sqlz
				end if
				temp=20: if isnumeric(temp) then Session.timeout=temp else Session.timeout=20
				rs.close
				sqlu="select top 2 entrada from login where usuario='" & usuario1 & "' order by entrada desc "
				rs.Open sqlu, ,adOpenStatic, adLockReadOnly
				if rs.recordcount>1 then
					rs.movenext
					Session("lastacesso")=rs("entrada")
				else
					Session("lastacesso")="-.-"
				end if
				rs.close
				Session("DescricaoErro")=""
			else
				Session("DescricaoErro")="Senha n�o confere"
			end if
'-----------------------------------------------------
		end if 'recordcount>0
	end if

	conexao.close
	Set rs=Nothing
	set conexao=nothing

End Sub

end if

if request.form("cmdLogout")<>"" then
	Session("DescricaoErro")=""
	Session("usuariomaster")="":Session("usuarioname")=""
	Session.Abandon
end if
if Session("estilo")="" then Session("estilo")="diversos.css"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>RH Online - Login</title>
<link rel="stylesheet" type="text/css" href="<%=Session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
<style type="text/css">
<!--
window1 {background: #000080; color: #FFFFFF; font: bold 8px; font-family: tahoma}
/* End of style section. Generated by AceHTML at 25/11/04 15:56:54 */
-->
</style>
</head>
<body>


<div align="center">

<%
'Session("usuariomaster")=""
if Session("usuariomaster")="" then
tam=290
%>
	<form action="intranet.asp" method="post" name="formlogin">
	
<table border="0" cellpadding="0" cellspacing="0" style="background-color:transparent;border-collapse: collapse;background:transparent url(images/acessorh.gif) no-repeat center;" width="620" height="350">
<tr><td colspan=2 height="220" style="background-color:transparent"></td></tr>
<tr><td width="<%=tam%>" height="39" style="background-color:transparent"></td>
	<td style="background-color:transparent" valign="top">
	<input type="text" name="txtUsuario" value="<%=Session("UsuarioMaster")%>" style="font-family:Tahoma; font-size:8pt; color:black; border:0px transparent; border-bottom:1px #000000 solid;background-color:white; " size="6" maxlength="6">
	</td>
</tr>
<tr><td width="<%=tam%>" height="40" style="background-color:transparent"></td>
	<td style="background-color:transparent" valign="top">
	<input type="password" name="txtPassword" value="" style="font-family:Tahoma; font-size:8pt; color:black; border:0px transparent; border-bottom:1px #000000 solid;background-color:white; " size="8" maxlength="8">
	</td>
</tr>
<tr><td width="<%=tam%>" style="background-color:transparent"></td>
	<td style="background-color:transparent" align="left">
	&nbsp;&nbsp;
	<input type="submit" class=button value=" Entrar " name="cmdLogin">
	</td></tr>
</table>	

	</form>
<%
else
%>
	<form action="intranet.asp" method="post" name="formlogout">

<table border="0" cellpadding="0" cellspacing="0" style="background-color:transparent;border-collapse: collapse;background:transparent url(images/acessorho.gif) no-repeat center;" width="620" height="350">
<tr><td class=campo style="background-color:transparent" align="center" valign=bottom>

	Usu�rio: <%=Session("usuarioname")%><br>
	Seu �ltimo acesso foi em <%=Session("lastacesso")%><br>
	IP: <%SendIp=request.servervariables("REMOTE_ADDR"):response.write Sendip%> 
	(<%SendIp=request.servervariables("LOCAL_ADDR"):response.write SendIp%>)
<br>
<br>
		<%if Session("acesso")=1 then%>
			<a href="index2.asp" onMouseOver="window.status='Clique para acesso ao menu principal';return true" onMouseOut="window.status='';return true">	<img src="images/setanext1.gif" width="12" height="12" border="0" alt="Clique para acesso ao menu principal">Clique para iniciar</a>
		<%end if%>
		<%if Session("acesso")=2 then%>
			<a href="indexp.asp" onMouseOver="window.status='Clique para acessar';return true" onMouseOut="window.status='';return true">	<img src="images/setanext1.gif" width="12" height="12" border="0" alt="Clique para acesso ao menu principal">Clique para iniciar</a>
		<%end if%>

</td></tr>
<tr><td style="background-color:transparent" align="center">
	&nbsp;&nbsp;
	<input type="submit" class=button value="ENCERRAR" name="cmdLogout">
	</td></tr>
</table>
	</form>
	
<%
end if
%>
</div>

<div align="center">
<table border="0" cellpadding="0" cellspacing="0" style="background-color:transparent;border-collapse: collapse;" width="620">
<tr>
	<td class=campo height="20">
<span id=tick1>
</span>&nbsp;
<script>
<!--
/*By JavaScript Kit
http://javascriptkit.com
Credit MUST stay intact for use
*/
function show1(){
   if (!document.all&&!document.getElementById)
      return
   thelement=document.getElementById? document.getElementById("tick1"): document.all.tick1
   var Digital=new Date()
   var hours=Digital.getHours()
   var minutes=Digital.getMinutes()
   var seconds=Digital.getSeconds()
   if (hours<=0)
      hours="0"+hours
   if (minutes<=9)
      minutes="0"+minutes
   if (seconds<=9)
      seconds="0"+seconds
   var ctime=hours+":"+minutes+":"+seconds+" "
   thelement.innerHTML="<b style='font-size:10;color:blue;'>"+ctime+"</b>"
   setTimeout("show1()",1000)
}
window.onload=show1
//-->
</script>
	</td>
	<td>
	<p style="font-size:10pt;font-family:tahoma;font-weight:normal;color:Red;"><%=Session("DescricaoErro")%>
	</td>
	<td class=campo align="right">
	&nbsp;<%=day(now()) & "/" & monthname(month(now()))%>&nbsp;
	</td>
</tr>
</table>

</div>


</body>
</html>