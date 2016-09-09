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
<title>Menu Principal</title>
<link rel="stylesheet" type="text/css" href="<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
<SCRIPT LANGUAGE="JavaScript">
function myprint() {
window.parent.frmMain.focus();
window.print();
}
</script>

<!-- MENU -->
<style type="text/css">
.menutitle{
cursor:pointer;
margin-bottom: 5px;
background-color:#ECECFF;
color:#000000;
width:110px;
padding:2px;
text-align:left;
font-weight:bold;
/*/*/border:1px solid #000000;/* */
}
.submenu{margin-bottom: 0.5em;}
</style>

<script type="text/javascript">
/***********************************************
* Switch Menu script- by Martial B of http://getElementById.com/
* Modified by Dynamic Drive for format & NS4/IE4 compatibility
* Visit http://www.dynamicdrive.com/ for full source code
***********************************************/
if (document.getElementById){ //DynamicDrive.com change
document.write('<style type="text/css">\n')
document.write('.submenu{display: none;}\n')
document.write('</style>\n')
}
function SwitchMenu(obj){
	if(document.getElementById){
	var el = document.getElementById(obj);
	var ar = document.getElementById("masterdiv").getElementsByTagName("span"); //DynamicDrive.com change
		if(el.style.display != "block"){ //DynamicDrive.com change
			for (var i=0; i<ar.length; i++){
				if (ar[i].className=="submenu") //DynamicDrive.com change
				ar[i].style.display = "none";
			}
			el.style.display = "block";
		}else{
			el.style.display = "none";
		}
	}
}
</script>
<!-- FIM MENU -->
</head>
<base target="frmMain">
<body style="background-color:#ffffcc;margin-top:0px;margin-left:2px;font-family:Tahoma;font-size:10px">
<a href="intranet.asp">
<img src="images/logo_centro_universitario_unifieo_big.gif" width="110" alt="" border=0></a>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
for a=1 to 100
	session("A" & a)=""
next
'**********************
session("A100")="T"
'**********************
sql="select m.* from intranet_menus m, (select idmenu as idmenu from intranet_grant where usuario='" & session("usuariomaster") & "' group by idmenu) g where m.idmenu=g.idmenu order by m.idmenu "
sql="select idmenu, menu from (SELECT gu.usuario, gr.idgrupo, gr.idmenu, m.menu FROM intranet_grupouser AS gu, intranet_grant AS gr, intranet_menus AS m " & _
"WHERE gu.idgrupo = gr.idgrupo AND gr.idmenu = m.idmenu GROUP BY gu.usuario, gr.idgrupo, gr.idmenu, m.menu " & _
"HAVING gu.usuario='" & session("usuariomaster") & "' UNION ALL " & _
"SELECT '',gr.idgrupo, gr.idmenu, m.menu FROM intranet_grant AS gr, intranet_menus AS m " & _
"WHERE gr.idmenu = m.idmenu GROUP BY gr.idgrupo, gr.idmenu, m.menu HAVING gr.idgrupo=1 ) as m1 " & _
"GROUP BY idmenu, menu ORDER BY idmenu "
rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst
%>
<!-- Keep all menus within masterdiv-->
<div id="masterdiv">
<%
do while not rs.eof
menu="menu" & rs("idmenu")
%>
<div class="menutitle" onClick="SwitchMenu('<%=menu%>')"><%=rs("menu")%></div>
<span class="submenu" id="<%=menu%>">
<%
sql2="select m.idsub, m.submenu, m.descricao, m.link from intranet_submenus m, intranet_grant g " & _
"where g.usuario='" & session("usuariomaster") & "' and (g.idsub=m.idsub and g.idmenu=m.idmenu) " & _
"and m.idmenu=" & rs("idmenu") & _
" order by m.idsub "
sql2="select usuario, idgrant, idgrupo, idmenu, idsub, descricao, link, acesso from ( " & _
"SELECT gu.usuario, gr.idgrant, gr.idgrupo, gr.idmenu, gr.idsub, sm.submenu, sm.descricao, sm.link, gr.acesso FROM (intranet_grupouser AS gu INNER JOIN intranet_grant AS gr ON gu.idgrupo = gr.idgrupo) INNER JOIN intranet_submenus AS sm ON (gr.idmenu = sm.idmenu) AND (gr.idsub = sm.idsub) " & _
"GROUP BY gu.usuario, gr.idgrant, gr.idgrupo, gr.idmenu, gr.idsub, sm.submenu, sm.descricao, sm.link, gr.acesso " & _
"HAVING gu.usuario='" & session("usuariomaster") & "' union all " & _
"SELECT '',gr.idgrant, gr.idgrupo, gr.idmenu, intranet_submenus.idsub, intranet_submenus.submenu, intranet_submenus.descricao, intranet_submenus.link, gr.acesso " & _
"FROM intranet_grant AS gr INNER JOIN intranet_submenus ON (gr.idmenu = intranet_submenus.idmenu) AND (gr.idsub = intranet_submenus.idsub) " & _
"GROUP BY gr.idgrant, gr.idgrupo, gr.idmenu, intranet_submenus.idsub, intranet_submenus.submenu, intranet_submenus.descricao, intranet_submenus.link, gr.acesso " & _
"HAVING gr.idgrupo=1 or gr.idgrupo=0 ) as m2 where idmenu=" & rs("idmenu") & " order by idmenu, idsub "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
session("A" & rs2("idgrant"))=rs2("acesso")
if rs2("descricao")="-" then
	response.write "<hr>"
else 
%>
•<a href="<%=rs2("link")%>" class="r"><%=rs2("descricao")%></a><font size="1">
 <%if session("usuariomaster")="02379" then response.write " (" & rs2("idgrant") & "-" & rs2("acesso") & ")"%><br>
<%
end if 
rs2.movenext
loop
rs2.close
%>
</span>
<%
rs.movenext
loop
rs.close
%>

<!--
	<img src="about.gif" onclick="SwitchMenu('sub6')"><br>
	<span class="submenu" id="sub6">
		- <a href="http://www.dynamicdrive.com/link.htm">Link to DD</a><br>
	</span>
-->
</div>

<span id=tick2>
</span>&nbsp;
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
<img src="images/Printer.gif" width="16" height="16" border="0" onclick="myprint()">
<%
response.write "<br>"
if session("usuariomaster")="02379" then response.write application("usuariosativos")

response.write "<br>"
if session("usuariogrupo")="COORD.CURSO" or session("usuariogrupo")="JURIDICO" then
%>
<a href="indexp.asp" target="_blank">Portal RH Professor</a>
<%
end if

%>
<%if session("usuariogrupo")="RH" then%>
<br>
<a href="http://10.0.1.91/docrh" target="_blank">DocRH</a>
<%end if%>
<%if session("usuariomaster")="023791" then%>
<a href="http://www.pokeplushies.com/feed/680004"><img src="http://www.pokeplushies.com/images/adoptables/680004.gif" border="0"><br>Click here to feed me a Rare Candy!</a><br><a href="http://www.pokeplushies.com">Get your own at PokePlushies!</a>
<a href="http://www.pokeplushies.com/feed/680034"><img src="http://www.pokeplushies.com/images/adoptables/680034.gif" border="0"><br>Click here to feed me a Star Fruit!</a><br><a href="http://www.flyffables.com">Get your own at Flyffables!</a>
<%end if%>
</body>
</html>