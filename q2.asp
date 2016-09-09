<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<html>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Teste Quina</title>
<body>
<p>Quina</p>

Jogos:
<%
dim n(40)
dim nv(7)
dim r(5)
r(1)=04 : r(2)=05 : r(3)=06 : r(4)=42 : r(5)=66
'r(1)=01 : r(2)=13 : r(3)=53 : r(4)=54 : r(5)=72
n(1)=26
n(2)=43
n(3)=67
n(4)=54
n(5)=16
n(6)=49
n(7)=80
n(8)=50
n(9)=22
n(10)=28
n(11)=29
n(12)=15
n(13)=39
n(14)=75
n(15)=10
n(16)=68
n(17)=36
n(18)=34
n(19)=40
n(20)=52
n(21)=7
n(22)=65
n(23)=72
n(24)=18
n(25)=77
n(26)=46
n(27)=37
n(28)=74
n(29)=79
n(30)=35
n(31)=1
n(32)=42
n(33)=27
n(34)=70
n(35)=76
n(36)=44
n(37)=59
n(38)=8
n(39)=12
n(40)=20

randomize timer
tvol3=0:tvol4=0:tvol5=0
apostas=7
jogos=3
comb=40
%>

<table>
<tr>
	<td valign=top>
<%
for a=1 to jogos
	tacerto=0
	response.write "<br>Volante " & a & ": "
	for b=1 to apostas
		numero=int(rnd()*comb)+1
		'nv(b)=numero
		nv(b)=n(numero)
		if b>1 then
			for c=1 to b-1
				if nv(c)=nv(b) then 
					randomize int(rnd*100)
					numero2=int(rnd()*comb)+1
					nv(b)=n(numero2)	
					c=c-1
				end if
			next
		end if
	'response.write " " & numzero(nv(b),2)
	next

	response.write " -> "
	for d=1 to apostas
		acertou=0
		for e=1 to 5
			if r(e)=nv(d) then
				acertou=1
				tacerto=tacerto+1
			end if
		next
		cor1="<font color=black>":cor2="</font>"
		if acertou=1 then cor1="<font color=blue><b>":cor2="</b></font>"
		response.write cor1 & numzero(nv(d),2) & cor2 & " "
	next
	if tacerto=3 then tvol3=tvol3+1
	if tacerto=4 then tvol4=tvol4+1
	if tacerto=5 then tvol5=tvol5+1
	response.write " -> " & tacerto

	'redim nv(5) preserve
next
%>
	</td>
	<td valign=top>
<%
response.write "<br> Volantes com 3 acertos: " & tvol3
response.write "<br> Volantes com 4 acertos: " & tvol4
response.write "<br> Volantes com 5 acertos: " & tvol5
%>
	</td>
</tr>
<table>

</body>
</html>