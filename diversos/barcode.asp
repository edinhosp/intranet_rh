<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<html><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<%
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
    Const Altura    = 75
    
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
chapa="02379"
tipo="F"
%>
<table width="200" height="320" border=1 cellspacing=0 cellpadding=0 style='border-collapse: collapse' bordercolor="#dcdcdc">
<tr>
<td valign="top" align="center">
&nbsp;
<%if tipo="F" then%>
<img border="0" src="../func_foto.asp?chapa=<%=chapa%>"  width="90">
<%else%>
<img border="0" src="../aluno_foto.asp?id=<%=rs("idimagem")%>" width="90">
<%end if%>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<%
Response.Write getCodigoBarras(chapa & digito(chapa))
response.write "<br>"
Response.Write (chapa & digito(chapa))
%>
</td>
</tr>
</table>
</html>

