<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Crachá Provisório Funcionário Admitido</title>

<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript" type="text/javascript"> <!--
/***** script montado por Edson Benevides
Unifieo - 10/12/2004 *******************/
var montharray=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
function nome1() { form.chapa.value=form.nome.value;}
function chapa1() { form.nome.value=form.chapa.value;}
--></script>
</head>
<body>
<%

if request.form="" then
%>

<form method="POST" action="crachafuncp.asp" name="form">
<table border="0" cellpadding="3" cellspacing="0" width="420" style="border-collapse: collapse">
<tr><td class=titulo colspan=2><p style="margin-top:0;margin-bottom:0;color:Blue;font-size:10pt;text-align:left">
<b>Emissão de Crachá Provisório para novos funcionários</font></p>
	</td></tr>
<tr>
	<td class=titulo>Chapa</td>
	<td class=titulo>Nome</td>
</tr>
<tr>
	<td class=titulo><input type="text" value="<%=chapa%>" name="chapa" size="8" onfocus="javascript:window.status='Informe o chapa do funcionário'"></td>
	<td class=titulo><input type="text" value="<%=nome%>" name="nome" size="50" onfocus="javascript:window.status='Informe o nome do funcionário'"></td>
</tr>
<tr>
	<td class=titulo>Apelido</td>
	<td class=titulo>Setor</td>
</tr>
<tr>
	<td class=titulo><input type="text" value="<%=apelido%>" name="apelido" size="15" onfocus="javascript:window.status='Informe o primeiro nome do funcionário'"></td>
	<td class=titulo><input type="text" value="<%=setor%>" name="setor" size="50" onfocus="javascript:window.status='Informe o setor do funcionário'"></td>
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
    Const Largo        = 6
    Const Altura    = 105 '75
    
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

'Exemplo de geração do código de barras
'Substitua o valor do parâmetro abaixo pelo número do código de barras.
chapa=request.form("chapa")
tipo="F"
nome=request.form("nome")
descricao=request.form("setor")
apelido=request.form("apelido")

%>
<div align="center">
<table width="200" height="320" border=1 cellspacing=0 cellpadding=0 style='border-collapse: collapse' bordercolor="#dcdcdc">
<tr><td valign="top" align="center">

<!-- *********************** -->
<table width="199" height="319" border=0 cellspacing=0 cellpadding=0 style='border-collapse: collapse' bordercolor="#dcdcdc">
<tr><td valign="center" align="center">
<img border="0" src="func_foto.jpg"  width="90">
</td></tr>

<tr><td valign="top" align="center">
	<table cellspacing=7 border=0 width=199><tr><td class=fundo align="center" valign="center">
	<p style="font-size:16px"><b><%=apelido%></b></p>
	</td></tr></table>
</td></tr>

<tr><td valign="top" align="center">
<p style="font-size:12px"><b><%=descricao%></b></p>
</td></tr>

<tr><td valign="bottom" align="center">
<img src="../images/aguia.jpg" width="150" border="0" alt="">
</td></tr>
</table>
<!-- *********************** -->

<!-- page -->
</td></tr></table>
</div>
<DIV style="page-break-after:always"></DIV>
<div align="center">
<table width="220" height="320" border=0 cellspacing=0 cellpadding=0 style='border-collapse: collapse' bordercolor="#dcdcdc">
<tr><td width=20>&nbsp;</td><td valign="top" align="center">
<!-- page -->

<!-- *********************** -->
<table width="199" height="319" border=0 cellspacing=0 cellpadding=0 style='border-collapse: collapse' bordercolor="#dcdcdc">
<tr><td valign="center" align="center">
<br>
<p style="font-size:12px"><b><%=nome%></b></p>
</td></tr>

<tr><td valign="top" align="center">
<p style="font-size:11px"><b>Registro FIEO: <%=chapa%>-<%=digito(chapa)%></b></p>
</td></tr>

<tr><td valign="center" align="center">
<p style="font-size:10px"><b>
Uso Obrigatório, em local visível.<br>
Para obtenção de 2ª via, será<br>
cobrada a despesa com a emissão.<br>
Devolução obrigatória, em caso<br>
de desligamento.<br>
Em caso de perda, comunicar ao<br>
Recursos Humanos.<br>
Telefone: 3651-9905
</b></p>
</td></tr>

<tr><td valign="top" align="center">
<p style="font-size:14px;color:#808080"><b>PROVISÓRIO</b></p>
</td></tr>

<tr><td valign="bottom" align="center">
<%
Response.Write getCodigoBarras(chapa & digito(chapa))
'Response.Write (chapa & digito(chapa))
%>
<br>
</td></tr>
</table>
<!-- *********************** -->
</td></tr></table>

</div>
<%
end if 'request.form
%>
</html>