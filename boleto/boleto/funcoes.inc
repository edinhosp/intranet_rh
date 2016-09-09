<%
Function I25Encode(StringNumber)
   Dim asPattern(), sSTART, sSTOP

   Redim asPattern(10)

   ' Padrões de Inicio e fim 
   sSTART = "NNNN"
   sSTOP  = "WNN"

   ' Padrões do código de barras 
   asPattern(0) = "NNWWN"
   asPattern(1) = "WNNNW"
   asPattern(2) = "NWNNW"
   asPattern(3) = "WWNNN"
   asPattern(4) = "NNWNW"
   asPattern(5) = "WNWNN"
   asPattern(6) = "NWWNN"
   asPattern(7) = "NNNWW"
   asPattern(8) = "WNNWN"
   asPattern(9) = "NWNWN"

   If (Len(StringNumber) Mod 2) <> 0 Then
      ' O número de caracteres no 
      ' argumento devem ser diferentes 
      I25Encode = ""
      Exit Function
   End If

      
   If Not IsNumeric(StringNumber) Then
      ' Argumento deve ser numérico 
      I25Encode = ""
      Exit Function
   Else
      If (InStr(StringNumber, ".") <> 0) Or _
         (InStr(StringNumber, ",") <> 0) Then
         ' O argumento é numero mais contem 
         ' caracteres invalidos para nós 
         I25Encode = ""
         Exit Function
      End If
   End If

   Dim sEncodedSTR, sUnit
   Dim iCharRead,   sDigit1, sDigit2, i

   sEncodedSTR = ""
   For iCharRead = 1 To Len(StringNumber) Step 2
       sDigit1 = asPattern( Asc( _
                 Mid( StringNumber, iCharRead, 1 ) ) - 48 )
       sDigit2 = asPattern( Asc( Mid( _
                 StringNumber, iCharRead + 1, 1 ) ) - 48 )

       sUnit = ""

       For i = 1 To 5
           sUnit = sUnit & Mid( sDigit1, i, 1 ) & _
                   Mid( sDigit2, i, 1 )
       Next

       sEncodedSTR = sEncodedSTR & sUnit
   Next
   I25Encode = sSTART & sEncodedSTR & sSTOP
End Function

Function Formata_Data(strData)
dia = day(strdata)
mes = month(strdata)
ano = year(strdata)
if dia < 9 then dia = "0" & dia
if mes < 9 then mes = "0" & mes
Formata_Data = dia & "/" & mes & "/" & ano
end function

Function Calculo_DV10(strNumero)
fator = 2
total = 0
For i = Len(strNumero) To 1 Step -1
	numero = Mid(strNumero, i, 1) * fator
	If numero > 9 Then
		numero = Cint(Left(numero, 1)) + Cint(Right(numero, 1))
	End If
	total = total + numero
	if fator = 2 then
		fator = 1
	else
		fator = 2
	end if
Next
resto = total Mod 10
resto = 10 - resto
if resto = 10 then
	Calculo_DV10 = 0
else
	Calculo_DV10 = resto
end if
End Function

Function calcula_DV_CodBarras(sequencia)
fator = 2
total = 0
For i = 43 To 1 step -1
	numero = Mid(sequencia, i, 1)
	If fator > 9 Then
		fator = 2
	End If
	numero = numero * fator
	total = total + numero
	fator = fator + 1
Next
resto = total mod 11
resultado = 11 - resto
If resultado = 10 Or resultado = 0 Then
    calcula_DV_CodBarras = 1
Else
    calcula_DV_CodBarras = resultado
End If
End Function

Function Monta_CodBarras()

database = CDate("7/10/1997")
fator = DateDiff("d", database, dt_venc)
valor = Int(valor_doc * 100)
while len(valor) < 10
    valor = "0" & valor
wend

codigo_sequencia = banco & moeda & fator & valor & carteira _
					& nossonumero & dv_nossonumero & agencia & conta & dv_conta & "000"

dvcb = calcula_DV_CodBarras(codigo_sequencia)
Monta_CodBarras = Left(codigo_sequencia, 4) & dvcb & Right(codigo_sequencia, 39)
End Function

Function Linha_Digitavel(sequencia_codigo_barra) 
seq1 = left(sequencia_codigo_barra,4) & mid(sequencia_codigo_barra,20,5)
seq2 = mid(sequencia_codigo_barra,25,10)
seq3 = mid(sequencia_codigo_barra,35,10)
seq4 = mid(sequencia_codigo_barra,6,14)
dvcb = mid(sequencia_codigo_barra,5,1)

dv1 = Calculo_DV10(seq1)
dv2 = Calculo_DV10(seq2)
dv3 = Calculo_DV10(seq3)

seq1 = Left(seq1 & dv1, 5) & "." & Mid(seq1 & dv1, 6, 5)
seq2 = Left(seq2 & dv2, 5) & "." & Mid(seq2 & dv2, 6, 6)
seq3 = Left(seq3 & dv3, 5) & "." & Mid(seq3 & dv3, 6, 6)


Linha_Digitavel = seq1 & " " & seq2 & " " & seq3 & " " & dvcb & " " & seq4
End Function%>