<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.redirect "intranet.asp"
'if session("a1")="N" or session("a1")="" then response.redirect "intranet.asp"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="refresh" content="180">
<title>Rodape</title>
<script language="javascript" type="text/javascript">
<!--
/*Author: Eric King     Url: http://redrival.com/eak/index.shtml     This script is free to use as long as this info is left in      Featured on Dynamic Drive script library (http://www.dynamicdrive.com)*/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>
<style type="text/css">
#dhtmltooltip{
	position: absolute;
	left: -300px;
	width: 150px;
	border: 1px solid black;
	padding: 2px;
	background-color: lightyellow;
	visibility: hidden;
	z-index: 100;
	/*Remove below line to remove shadow. Below line should always appear last within this CSS*/
	filter: progid:DXImageTransform.Microsoft.Shadow(color=gray,direction=135);
}
#dhtmlpointer{
	position:absolute;
	left: -300px;
	z-index: 101;
	visibility: hidden;
}
</style>
<link rel="stylesheet" type="text/css" href="<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<base target="frmMain">
<body style="background-color:#ffffcc;margin-top:0px;margin-left:0px;margin-right:0px;font-family:Tahoma;font-size:10px">
<script type="text/javascript">
/***********************************************
* Cool DHTML tooltip script II- © Dynamic Drive DHTML code library (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit Dynamic Drive at http://www.dynamicdrive.com/ for full source code
***********************************************/
var offsetfromcursorX=12 //Customize x offset of tooltip
var offsetfromcursorY=10 //Customize y offset of tooltip
var offsetdivfrompointerX=10 //Customize x offset of tooltip DIV relative to pointer image
var offsetdivfrompointerY=14 //Customize y offset of tooltip DIV relative to pointer image. Tip: Set it to (height_of_pointer_image-1).
document.write('<div id="dhtmltooltip"></div>') //write out tooltip DIV
document.write('<img id="dhtmlpointer" src="images/arrow2.gif">') //write out pointer image
var ie=document.all
var ns6=document.getElementById && !document.all
var enabletip=false
if (ie||ns6)
var tipobj=document.all? document.all["dhtmltooltip"] : document.getElementById? document.getElementById("dhtmltooltip") : ""
var pointerobj=document.all? document.all["dhtmlpointer"] : document.getElementById? document.getElementById("dhtmlpointer") : ""

function ietruebody(){
	return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}
function ddrivetip(thetext, thewidth, thecolor){
	if (ns6||ie){
		if (typeof thewidth!="undefined") tipobj.style.width=thewidth+"px"
		if (typeof thecolor!="undefined" && thecolor!="") tipobj.style.backgroundColor=thecolor
		tipobj.innerHTML=thetext
		enabletip=true
		return false
	}
}
function positiontip(e){
	if (enabletip){
	var nondefaultpos=false
	var curX=(ns6)?e.pageX : event.clientX+ietruebody().scrollLeft;
	var curY=(ns6)?e.pageY : event.clientY+ietruebody().scrollTop;
	//Find out how close the mouse is to the corner of the window
	var winwidth=ie&&!window.opera? ietruebody().clientWidth : window.innerWidth-20
	var winheight=ie&&!window.opera? ietruebody().clientHeight : window.innerHeight-20
	var rightedge=ie&&!window.opera? winwidth-event.clientX-offsetfromcursorX : winwidth-e.clientX-offsetfromcursorX
	var bottomedge=ie&&!window.opera? winheight-event.clientY-offsetfromcursorY : winheight-e.clientY-offsetfromcursorY
	var leftedge=(offsetfromcursorX<0)? offsetfromcursorX*(-1) : -1000
	//if the horizontal distance isn't enough to accomodate the width of the context menu
		if (rightedge<tipobj.offsetWidth){
		//move the horizontal position of the menu to the left by it's width
		tipobj.style.left=curX-tipobj.offsetWidth+"px"
		nondefaultpos=true
		}
	else if (curX<leftedge)
	tipobj.style.left="5px"
	else{
	//position the horizontal position of the menu where the mouse is positioned
	tipobj.style.left=curX+offsetfromcursorX-offsetdivfrompointerX+"px"
	pointerobj.style.left=curX+offsetfromcursorX+"px"
	}
//same concept with the vertical position
	if (bottomedge<tipobj.offsetHeight){
	tipobj.style.top=curY-tipobj.offsetHeight-offsetfromcursorY+"px"
	nondefaultpos=true
	}
	else{
	tipobj.style.top=curY+offsetfromcursorY+offsetdivfrompointerY+"px"
	pointerobj.style.top=curY+offsetfromcursorY+"px"
	}
	tipobj.style.visibility="visible"
	if (!nondefaultpos)
	pointerobj.style.visibility="visible"
	else
	pointerobj.style.visibility="hidden"
	}
}
function hideddrivetip(){
	if (ns6||ie){
	enabletip=false
	tipobj.style.visibility="hidden"
	pointerobj.style.visibility="hidden"
	tipobj.style.left="-1000px"
	tipobj.style.backgroundColor=''
	tipobj.style.width=''
	}
}
document.onmousemove=positiontip
</script>

<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
'rs.Open sql, ,adOpenStatic, adLockReadOnly
'rs.movefirst

if session("iCRMVer")="" then session("iCRMVer")=1
if session("iCRMTarefa")="" then session("iCRMTarefa")=0
if session("iCRMPessoa")="" then session("iCRMPessoa")=0
'if request("iCRMVer")="M" and (request("iCRMTarefa")="" or request("iCRMPessoa")="") then session("iCRMVer")=0 else session("iCRMVer")=1	'T/M
'if request("iCRMTarefa")="S" and (request("iCRMVer")="" or request("iCRMPessoa")="") then session("iCRMTarefa")=1 else session("iCRMTarefa")=0	'S/N
'if request("iCRMPessoa")="S" and (request("iCRMTarefa")="" or request("iCRMVer")="") then session("iCRMPessoa")=1 else session("iCRMPessoa")=0	'S/N
if request("iCRMVer")="M" then session("iCRMVer")=0 
if request("iCRMVer")="T" then session("iCRMVer")=1	'T/M
if request("iCRMTarefa")="S" then session("iCRMTarefa")=1 
if request("iCRMTarefa")="N" then session("iCRMTarefa")=0	'S/N
if request("iCRMPessoa")="S" then session("iCRMPessoa")=1 
if request("iCRMPessoa")="N" then session("iCRMPessoa")=0	'S/N

response.write "<br>" & session("iCRMVer") & "-" & session("iCRMTarefa") & "-" & session("iCRMPessoa")
response.write "<br>" & request("iCRMVer") & "-" & request("iCRMTarefa") & "-" & request("iCRMPessoa")
%>

<form method="POST" action="frmRodape.asp" target="_self" name="listaCRM" >
<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse:collapse" width="100%">
<tr>
	<td class="campor" valign="top" colspan=2>
		<span id=tick2></span>&nbsp;
		<a href="crm/lista_nova.asp" onclick="NewWindow(this.href,'CRM','400','230','yes','center');return false" onfocus="this.blur()">
		<img src="imagesr/page_new.gif" border="0" alt="Clique para inserir movimento"></a>

<%if session("iCRMVer")=1 then%>
	<a href="frmRodape.asp?iCRMVer=M" target="_self"><img src='imagesr/user.png' border='0' alt='Mostrar apenas os meus'></a>
<%else%>
	<a href="frmRodape.asp?iCRMVer=T" target="_self"><img src='imagesr/group.png' border='0' alt='Mostrar todos'></a>
<%end if%>

<%if session("iCRMTarefa")=1 then%>
	<a href="frmRodape.asp?iCRMTarefa=N" target="_self"><img src='imagesr/table_multiple.png' border='0' alt='Mostrar a lista das tarefas'></a>
<%else%>
	<a href="frmRodape.asp?iCRMTarefa=S" target="_self"><img src='imagesr/layout_content.png' border='0' alt='Mostrar agrupado por tipo'></a>
<%end if%>
	
<%if session("iCRMPessoa")=1 then%>
	<a href="frmRodape.asp?iCRMPessoa=N" target="_self"><img src='imagesr/table_multiple.png' border='0' alt='Mostrar a lista das pessoas'></a>
<%else%>
	<a href="frmRodape.asp?iCRMPessoa=S" target="_self"><img src='imagesr/page_user.gif' border='0' alt='Mostrar agrupado por pessoa'></a>
<%end if%>
		
	</td>
</tr>

<tr><td class="campos" valign="top" colspan=2><img src="imagesr/status_blue.png" border="0">
<img src="imagesr/status_green.png" border="0">
<img src="imagesr/status_yellow.png" border="0">
<img src="imagesr/status_red.png" border="0">
<img src="imagesr/status_ok.png" border="0">
<img src="imagesr/estouro.png" border="0">
<img src="imagesr/bomb.png" border="0">
</td></tr>
<%
if session("usuariomaster")="02379" or session("usuariomaster")="00259" or session("usuariogrupo")="RH" then

sqlt="select valorD from iCrm_Parametros where parametro='LastCheck'"
rs.Open sqlt, ,adOpenStatic, adLockReadOnly
if rs.recordcount=0 or isnull(rs("valorD")) or rs("valorD")="" or rs("valorD")+(1/24)<now() then
	sqls="select idCRM, script from iCRM_Script"
	rs2.Open sqls, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
	if rs2("script")<>"" then script=replace(rs2("script"),"#USER#",session("usuariomaster"))
	conexao.execute script
	rs2.movenext:loop
	rs2.close

	sqlu="update iCrm_Parametros set ValorD='" & dtaccess(now()) & "' where Parametro='LastCheck'"
	sqlu="update iCrm_Parametros set ValorD=getdate() where Parametro='LastCheck'"
	conexao.execute sqlu
end if '--
rs.close

'if request.form("todos")="ON" then
sqlmain="select * from iCRM_Lista where idFluxo>0 order by dia4, indicador "
sqlmain="select * from iCRM_Lista where idFluxo>0 "
sqlo1="order by dia4, indicador "
if session("iCRMVer")=1 then sqlw1="" else sqlw1="and (r='" & session("usuariomaster") & "' or S='" & session("usuariomaster") & "') "
if session("iCRMTarefa")=1 then sqlo1="order by atividade, dia4, indicador "
if session("iCRMPessoa")=1 then sqlo1="order by chapa, dia4, indicador "

sql=sqlmain & sqlw1 & sqlo1

rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount=0 then
	sql="select * from iCRM_Lista order by dia4, indicador "
	rs.close
	rs.Open sql, ,adOpenStatic, adLockReadOnly
end if
response.write "<tr><td class=""campos"">" & rs.recordcount & "</td></tr>"
rs.movefirst
do while not rs.eof
select case rs("indicador")
	case 1
		imagem="imagesr/status_green.png"
	case 2
		imagem="imagesr/status_yellow.png"
	case 3
		imagem="imagesr/status_red.png"
	case 4 and now()>rs("dtvencimento")
		imagem="imagesr/estouro.png"
	case 4 and now()<rs("dtvencimento")
		imagem="imagesr/bomb.png"
end select
select case rs("status")
	case "C"
		imagem="imagesr/status_blue.png"
	case "F"
		imagem="imagesr/status_ok.png"
end select
if rs("link_f")="" or isnull(rs("link_f")) then link=0 else link=1
linksub=rs("link_f")
if link=1 and rs("chapa")<>"" then linksub=replace(linksub,"#chapa#",rs("chapa"))
if link=1 and rs("dtvencimento")<>"" then linksub=replace(linksub,"#data#",rs("dtvencimento"))
txtTip="<p style=font-size:7pt>" & rs("atividade") & "<br>" & rs("tarefa") & "<br>" & _
rs("nome") & " (" & rs("chapa") & ")<br>" & _
"Prazo ate: " & rs("dtvencimento") & "<br>" & _
":" & rs("anotacao")
%>
<!--onMouseover="ddrivetip('JavaScriptKit.com JavaScript tutorials', 300)"; onMouseout="hideddrivetip()"-->
<tr onMouseover="ddrivetip('<%=txtTip%>', 110)"; onMouseout="hideddrivetip()">
	<td class="campos" valign=middle nowrap rowspan=2 style="border-bottom:1px dotted blue">
	<a class=r href="crm/lista_alteracao.asp?codigo=<%=rs("idFluxo")%>" onclick="NewWindow(this.href,'CRM','350','230','yes','center');return false" onfocus="this.blur()">
	<img src="<%=imagem%>" border="0"></a>
	</td>
	<td class="campos" valign=top nowrap>
	<%=rs("dtprazo")%> | <%=left(rs("nome"),10)%>
	</td>
</tr>
<tr onMouseover="ddrivetip('<%=txtTip%>',110)"; onMouseout="hideddrivetip()">
	<td class="campos" valign=top nowrap style="border-bottom:1px dotted blue">
	<%if link=1 then%>
		<a style="font-size:7pt" href="<%=linksub%>" target="frmMain"><%=rs("tarefa")%></a>
	<%else%>
		<%=rs("tarefa")%>
	<%end if%>
	</td>
</tr>
<%
rs.movenext
loop
rs.close

end if 'session usuariomaster
%>

</table>
</form>
<script>
<!--
/*By JavaScript Kit
http://javascriptkit.com
Credit MUST stay intact for use
*/
function show2(){
   if (!document.all&&!document.getElementById)
      return
   thelement=document.getElementById? document.getElementById("tick2"): document.all.tick2
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
   setTimeout("show2()",1000)
}
window.onload=show2
//-->
</script>
</body>
</html>