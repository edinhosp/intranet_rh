<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Crachá Provisório Aluno</title>

<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
--></script>
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

if request.form="" then
%>
<form method="POST" action="crachaaluno2.asp" name="form">
<table border="0" cellpadding="3" cellspacing="0" width="420" style="border-collapse: collapse">
<tr><td class=titulo colspan=3><p style="margin-top:0;margin-bottom:0;color:Blue;font-size:10pt;text-align:left">
<b>Emissão de Crachá Provisório</font></p>
	</td></tr>
<tr>
	<td class=titulo>Matrícula</td>
	<td class=titulo>Nome do Aluno</td>
	<td class=titulo>Via</td>
</tr>
<tr>
	<td class=titulo><input type="text" value="<%=chapa%>" name="chapa" size="8" onfocus="javascript:window.status='Informe a matrícula do aluno'"></td>
	<td class=titulo>&nbsp;
	</td>
	<td class=titulo><input type="text" value="01" name="via" size="2" onfocus="javascript:window.status='Informe a via'"></td>
</tr>
</table>
<input type="submit" value="Pesquisar" class="button" name="pesquisar" onfocus="javascript:window.status='Clique aqui para pesquisar'">
</form>

<%
else 'request.form
'-----------------------------------------------------
'Funcao: getCodigoBarras(ByVal Numeros)
'Sinopse: Rotina para gerar códigos de barra padrão 2of5 ou 25.
'Parametro:
'       Numeros: Números para a geração do código de barras
'Retorno: String (HTML com o código gerado)
'Autor: William Nazato (wil@merconet.com.br)
'-----------------------------------------------------
Function getCodigoBarras(ByVal Numeros)
    Dim F, F1, F2, i, Texto
    Dim arrCodigoBarra(99)
    Dim htmlCodigoBarra
    htmlCodigoBarra = ""
    Const Fino        = 2
    Const Largo       = 4
    Const Altura      = 55
    
    if isempty(arrCodigoBarra(0)) Then
        arrCodigoBarra(0) = "00110"
        arrCodigoBarra(1) = "10001"
        arrCodigoBarra(2) = "01001"
        arrCodigoBarra(3) = "11000"
        arrCodigoBarra(4) = "00101"
        arrCodigoBarra(5) = "10100"
        arrCodigoBarra(6) = "01100"
        arrCodigoBarra(7) = "00011"
        arrCodigoBarra(8) = "10010"
        arrCodigoBarra(9) = "01010"
        For F1 = 9 To 0 Step -1
            For F2 = 9 To 0 Step -1
                F = F1 * 10 + F2
                Texto = ""
                For i = 1 To 5
                    Texto = Texto & Mid(arrCodigoBarra(F1), i, 1) + Mid(arrCodigoBarra(F2), i, 1)
                Next
                arrCodigoBarra(f) = Texto
            Next
        Next
    End if

    'Construindo o código HTML do código de barras
    'Guarda inicial
    htmlCodigoBarra = htmlCodigoBarra & "<img src=p.jpg width=" & Fino & " height=" & Altura & " border=0>"
    htmlCodigoBarra = htmlCodigoBarra & "<img src=b.jpg width=" & Fino & " height=" & Altura & " border=0>"
    htmlCodigoBarra = htmlCodigoBarra & "<img src=p.jpg width=" & Fino & " height=" & Altura & " border=0>"
    htmlCodigoBarra = htmlCodigoBarra & "<img src=b.jpg width=" & Fino & " height=" & Altura & " border=0>"
    htmlCodigoBarra = htmlCodigoBarra & "<img"
    Texto = Numeros
    if Len(Texto) Mod 2 <> 0 Then Texto = "0" & Texto End if
    'HTML dos dados
    Do While Len(Texto) > 0
        i        = Cint(Left(Texto,2))
        Texto    = Right(Texto, Len(Texto)- 2)
        F        = arrCodigoBarra(i)
        For i = 1 To 10 Step 2
            If Mid(F, i, 1) = "0" Then
                F1 = Fino
            Else
                F1 = Largo
            End If
            
            htmlCodigoBarra = htmlCodigoBarra & " src=p.jpg width=" & F1 & " height=" & Altura & " border=0><img"
            
            If mid(F, i + 1, 1) = "0" Then
                F2 = Fino
            Else
                F2 = Largo
            End If

            htmlCodigoBarra = htmlCodigoBarra & " src=b.jpg width=" & F2 & " height=" & Altura & " border=0><img"
    
        Next
    Loop
    
    'Guarda final
    htmlCodigoBarra = htmlCodigoBarra & " src=p.jpg width=" & Largo & " height=" & Altura & " border=0>"
    htmlCodigoBarra = htmlCodigoBarra & "<img src=b.jpg width=" & Fino & " height=" & Altura & " border=0>"
    htmlCodigoBarra = htmlCodigoBarra & "<img src=p.jpg width=1 height=" & Altura & " border=0>"
    
    'Retornando a função
    getCodigoBarras    = htmlCodigoBarra
End Function

sql1="select distinct top 4 e.matricula, e.nome, e.dtnasc, e.idimagem, c.codcur, c1.nome as curso " & _
"from corporerm.dbo.ealunos e, corporerm.dbo.ualucurso c, corporerm.dbo.ucursos c1, corporerm.dbo.umatricpl pl, corporerm.dbo.ealuocor o " & _
"where e.matricula=c.mataluno and c.codcur=c1.codcur and pl.mataluno=c.mataluno and pl.codcur=c.codcur " & _
"and o.mataluno=e.matricula and o.codperlet=pl.perletivo and pl.status=c.status and c.status='01' " & _
"and o.codperlet like '2009%' and o.codgrpocor=1 and o.codocorrencia in (1,2) " & _
"order by e.nome "
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
%>
<div align="center">
<!-- *********************** -->
<table width="319" height="199" border=0 cellspacing=0 cellpadding=0 style='border-collapse: collapse' bordercolor="#dcdcdc">
<tr><td valign="top" align="left">
<img src="../images/logo_centro_universitario_unifieo_big.jpg" height="40" border="0" alt="">
</td>
<td valign="center" align="center">
<img border="0" src="../aluno_foto.asp?id=<%=rs2("idimagem")%>" height="90">
</td>
</tr>

<tr><td valign="top" align="center" colspan=2 class=fundo height=50>
	<table cellspacing=0 border=0 width=315><tr><td class=fundo align="left" valign="center" colspan=2>
	<p style="font-size:8px;margin-top:0;margin-bottom:0"><b>Nome</b></p>
	<p style="font-size:10px;margin-top:0;margin-bottom:0"><b><%=rs2("nome")%></b></p>
	</td></tr>
	<tr><td class=fundo align="left" valign="center">
	<p style="font-size:8px;margin-top:0;margin-bottom:0"><b>Prontuário</b></p>
	<p style="font-size:10px;margin-top:0;margin-bottom:0"><b><%=rs2("matricula")%></b></p>
	</td><td class=fundo align="left" valign="center">
	<p style="font-size:8px;margin-top:0;margin-bottom:0"><b>Via</b></p>
	<p style="font-size:10px;margin-top:0;margin-bottom:0"><b><%=request.form("via")%></b></p>
	</td></tr></table>
</td></tr>
</table>
<!-- *********************** -->

<DIV style="page-break-after:always"></DIV>

<!-- *********************** -->
<table width="319" height="199" border=0 cellspacing=0 cellpadding=0 style='border-collapse: collapse' bordercolor="#dcdcdc">
<tr><td valign="center" align="left">
<p style="font-size:10px;margin-top:0;margin-bottom:0"><b>Obrigatório para a entrada nos <i>campi</i> UNIFIEO.
<p style="font-size:9px;margin-top:0;margin-bottom:0">Este cartão é pessoal, intransferível e de inteira responsabilidade do aluno.
<br>Em caso de extravio ou dano, a 2ª via deve ser requerida na Secretaria Geral com o pagamento de taxa.
<br><br>
<p style="font-size:10px;margin-top:0;margin-bottom:0;text-align:center"><b>www.unifieo.br
</td></tr>

<tr><td valign="center" align="left">
	<p style="font-size:8px;margin-top:0;margin-bottom:0"><b>&nbsp;&nbsp;Curso</b></p>
	<p style="font-size:10px;margin-top:0;margin-bottom:0"><b>&nbsp;&nbsp;<%=rs2("curso")%></b></p>
	<p style="font-size:8px;margin-top:0;margin-bottom:0"><b>&nbsp;&nbsp;Dt.Nasc.</b></p>
	<p style="font-size:10px;margin-top:0;margin-bottom:0"><b>&nbsp;&nbsp;<%=rs2("dtnasc")%></b></p>
</td></tr>

<tr><td valign="bottom" align="center">
<%
Response.Write getCodigoBarras(rs2("matricula")&request.form("via")) & "<br>"
%>
<br>
</td></tr>
</table>
<!-- *********************** -->

</div>
<%
rs2.movenext
loop

rs2.close
end if 'request.form
%>
</html>