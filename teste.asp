<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Pesquisa sobre seu professor</TITLE>
<link rel="stylesheet" type="text/css" href="diversos.css">
</HEAD>
<BODY>
<%
if request("aluno")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://www.unifieo.br';</script>"
if request("professor")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://www.unifieo.br';</script>"
dim conexao,rs,marc(6), formato(6)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
'rs3.Open sql, ,adOpenStatic, adLockReadOnly

%>
<table border='1' bordercolor='#000000' cellpadding='4' cellspacing='0' style='border-collapse:collapse'>


</table>

<%
	response.write "Professor: " & request("professor")
	response.write "<br>"
	response.write "Aluno: "  & request("aluno")
%>
<br>O Unifieo deseja saber a sua opini�o sobre os seus professores neste semestre.
<br>Pode responder tranquilamente, que independentemente da sua opini�o voc� n�o ser� identificado perante o professor.
<br>
<Br>1. O professor apresentou no in�cio das aulas o conte�do a ser lecionado durante o semestre?
<br><input type="radio" name="p1" value="S">Sim
<br><input type="radio" name="p1" value="N">N�o
<br>
<br>2. O professor explica de maneira clara o conte�do da disciplina!
<br><input type="radio" name="p2" value="1">N�o explica
<br><input type="radio" name="p2" value="2">Explica, mas � dif�cil acompanhar
<br><input type="radio" name="p2" value="3">Sim
<br><input type="radio" name="p2" value="4">Sim e d� exemplos pr�ticos
<br>
<br>
</BODY>
</HTML>
