<html>
<!-- Generated by AceHTML Freeware http://freeware.acehtml.com -->
<!-- Creation date: 03/02/05 -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<title></title>
</head>
<body>
<!--
<script language="javascript">
document.write ("<font size=7>");
document.write ("<font size=2>Party? 181/213/233 = " +String.fromCharCode(181,213,233) + "<br>" );
document.write ("<font size=2>I come first 183/213/232/224/195/210 = " +String.fromCharCode(183,213,232,224,195,210) + "<br>" );
document.write ("<font size=2>Dont jam 205/194/232/210/225/168/193 = " +String.fromCharCode(205,194,232,210,225,168,193) + "<br>" );
document.write ("<font size=2>Play solo 162/205/224/180/213/232/194/199 = " +String.fromCharCode(165,205,224,180,213,232,194,199) + "<br>" );
document.write ("<font size=2>Ed 224/205/231/180 = " +String.fromCharCode(224,205,231,180) + "<br>" );
document.write ("<font size=7>");

for (i=161;i<=211;i++)
document.write (i + "-" + String.fromCharCode(i) + "<br>" );
for (i=224;i<=230;i++)
document.write (i + "-" + String.fromCharCode(i) + "<br>" );
for (i=239;i<=251;i++)
document.write (i + "-" + String.fromCharCode(i) + "<br>" );
for (i=212;i<=218;i++)
document.write (i + "-" + String.fromCharCode(i) + "<br>" );
for (i=231;i<=238;i++)
document.write (i + "-" + String.fromCharCode(i) + "<br>" );

</script>
-->

<%
response.write "<font size=2>Party? 181/213/233 = <font size=6>" & chr(181)&chr(213)&chr(233) :response.write "<br>"
response.write "<font size=2>I come first 183/213/232/224/195/210 = = <font size=6>" & chr(183)&chr(213)&chr(232)&chr(224)&chr(195)&chr(210) :response.write "<br>"
response.write "<font size=2>Dont jam 205/194/232/210/225/168/193 = <font size=6>" & chr(205)&chr(194)&chr(232)&chr(210)&chr(225)&chr(168)&chr(193) :response.write "<br>"
response.write "<font size=2>Play solo 162/205/224/180/213/232/194/199 = <font size=6>" & chr(162)&chr(205)&chr(224)&chr(180)&chr(213)&chr(232)&chr(194)&chr(199) :response.write "<br>"
response.write "<font size=2>Ed 224/205/231/180 = <font size=6>" & chr(224)&chr(205)&chr(231)&chr(180) :response.write "<br>"
response.write "<font size=6>"
response.write "<table border=1 cellpadding=9 cellspacing=0><tr><td valign=top><font size=2>"
for a= 161 to 180
	response.write a & " - <font size=6>" & chr(a) :response.write "<br><font size=2>"
next
response.write "</td><td valign=top><font size=2>"

for a= 181 to 200
	response.write a & " - <font size=6>" & chr(a) :response.write "<br><font size=2>"
next
response.write "</td><td valign=top><font size=2>"

for a= 201 to 211
	response.write a & " - <font size=6>" & chr(a) :response.write "<br><font size=2>"
next

for a= 224 to 230
	response.write a & " - <font size=6>" & chr(a) :response.write "<br><font size=2>"
next
response.write "</td><td valign=top><font size=2>"

for a= 239 to 251
	response.write a & " - <font size=6>" & chr(a) :response.write "<br><font size=2>"
next
response.write "</td><td valign=top><font size=2>"

for a= 212 to 218
	response.write a & " - <font size=6>" & chr(205)&chr(a) :response.write "<br><font size=2>"
next

for a= 231 to 238
	response.write a & " - <font size=6>" & chr(205)&chr(a) :response.write "<br><font size=2>"
next

response.write "</td></tr></table>"
%>
</body>
</html>