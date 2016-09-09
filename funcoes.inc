<%
Function SimNao (valor)
	if isnull(Valor) then Valor=0
	valor=cint(valor)
	if Valor=0 then temp="images/bullet_off.gif"
	if Valor=1 then temp="images/bullet_on.gif"
	SimNao=temp
End Function

Function Boleano(Valor)
	if isnull(Valor) then Valor=0
	valor=cint(valor)
	temp=""
	if Valor=0 then temp="Não"
	if Valor=1 then temp="Sim"
	Boleano=temp
End Function

Function IIf(condition,value1,value2)
	If condition Then IIf = value1 Else IIf = value2
End Function

Function JulDate(vdate)
	if isnull(vdate) then vdate=now()
	JulDate=int((Year(vdate) - 2000 + 100) * 1000 + vdate -datevalue("01/01/"&Year(vdate))+1)
End Function

Function HoraLoad(numero,tipo)
	'transforma hora no formato chronus para formato hh:mm
	if isnull(numero) or numero="" or numero="0" then numero=0
		hora=int(numero/60)
		minuto=numero-int(hora*60)
	if hora>=100 then hora=numzero(hora,len(hora))
	if hora<100 then hora=numzero(hora,2)
	if minuto>=0 then minuto=numzero(minuto,2)
	if tipo=2 and numero=0 then temp="" else temp=hora & ":" & minuto
	if tipo=1 then temp=hora & ":" & minuto
	HoraLoad=temp
End Function

Function HoraSave(numero)
	'transforma hora no formato hh:mm para o formato chronus
	if isnull(numero) then numero=0
	if numero="" then numero=0
	'msgbox "f1" & numero
	hora=hour(numero)
	'msgbox "f2" & hora
	minuto=minute(numero)
	'msgbox "f3" & minuto
	temp=hora*60+minuto
	'msgbox "f4" & temp
	HoraSave=temp
End Function

Function Digito(numero)
	if isnull(numero) then numero=0
	dim carac(99)
	for i=1 to len(numero)
		carac(i-1)=mid(numero,i,1)
	next
	for i=1 to len(numero)
		digito1=digito1 + carac(len(numero)-i)*(i+1)
	next
	Digito=11 - (digito1 mod 11)
	if Digito=11 then Digito=0
	if Digito=10 then Digito=0
End Function

Function numzero(numero,zeros)
	if zeros=0 or isnull(zeros) then zeros=1 'len(numero)
	tamanho=len(numero)
	numzero=string(zeros-tamanho,"0")&numero
End Function

Function Pass(pSerie)
    pSerie = UCase(pSerie)
    pNum = 0
    For x = 1 To Len(pSerie)
        pNum = pNum + Asc(Mid(pSerie, x, 1)) * x
    Next
    Pass = pNum
End Function

Function DtAccess(data)
    if isnull(data) then exit function
    temp=""&month(data)&"/"&day(data)&"/"&year(data)&""
    DtAccess=temp
End Function

Function Dataform(data)
    if isnull(data) then exit function
    temp=""&numzero(day(data),2)&"/"&numzero(month(data),2)&"/"&year(data)&""
    Dataform=temp
End Function

Function ddmm8(data)
    if isnull(data) then exit function
    temp=""&numzero(day(data),2)&""&numzero(month(data),2)&""&year(data)&""
    ddmm8=temp
End Function

Function DtSql(data)
    if isnull(data) then exit function
    temp=""&numzero(year(data),4)&numzero(month(data),2)&numzero(day(data),2)&""
    DtAccess=temp
End Function


Function DtMysql(data)
    if isnull(data) then exit function
    temp=""&year(data)&"/"&month(data)&"/"&day(data)&""
    DtMysql=temp
End Function

function horarm(valor)
    if isnull(valor) then exit function
    hora=int(valor)
    minuto=valor-hora
    horarm=hora & ":" & minuto
end function

Function NrFile(numero)
    newstring=""
    if numero="" then nraccess=0:exit function
    if isnull(numero) then nraccess="null":exit function
    for a=1 to len(numero)
	p=mid(numero,a,1)
	if p="," then p="" else p=p
	if p="." then p="" else p=p
	newstring=newstring & p
    next
    NrFile=newstring
End Function

Function NrAccess(numero)
    newstring=""
    if numero="" then nraccess=0:exit function
    if isnull(numero) then nraccess="null":exit function
    for a=1 to len(numero)
	p=mid(numero,a,1)
	if p="," then p="." else p=p
	newstring=newstring & p
    next
    NrAccess=newstring
End Function

Function Aspas(string)
    nova=""
    velha=string
    for x = 1 to len(velha)
        caracter = mid(velha, x, 1)
        if Asc(caracter)=39 or Asc(caracter)=44 then
	    letra=""
        else
            letra=caracter
        end if
	nova=nova+letra
    next
    Aspas=nova
End Function

Function TextoPuro(Texto, Tipo)
    'tipo 1 - converte letras
    'tipo 2 - converte letras e numeros
    'tipo 3 - converte numeros
    tamanho = Len(Texto)
    Texto2 = "":Tipo=cint(Tipo)
    If tamanho > 0 Then
    For L = 1 To tamanho
        Letra = Mid(Texto, L, 1)
        If Asc(Letra) >= 97 And Asc(Letra) <= 122 Then Letra = Chr(Asc(Letra) - 32)
        If Tipo = 1 Then
            If (Asc(Letra) >= 65 And Asc(Letra) <= 90) Or Asc(Letra) = 32 Then
                Texto2 = Texto2 & Letra
            Else
                Texto2 = Texto2 & ""
            End If
        End If
        If Tipo = 2 Then
            If (Asc(Letra) >= 65 And Asc(Letra) <= 90) Or Asc(Letra) = 32 _
             Or (Asc(Letra) >= 48 And Asc(Letra) <= 57) Then
                Texto2 = Texto2 & Letra
            Else
                Texto2 = Texto2 & ""
            End If
        End If
        If Tipo = 3 Then
            If (Asc(Letra) >= 48 And Asc(Letra) <= 57) Then
                Texto2 = Texto2 & Letra
            Else
                Texto2 = Texto2 & ""
            End If
        End If
    Next
    TextoPuro = Texto2

    End If
End Function

Function Espaco1(Texto, Quant)
    If IsNull(Texto) Then Texto = ""
    If IsNull(Texto) Then Exit Function
    Texto2 = Left(Texto, Quant)
    Texto3 = Space(Quant - Len(Texto2))
    Espaco1 = Texto3 + Texto2
End Function

Function Espaco2(Texto, Quant)
    If IsNull(Texto) Then Exit Function
    Texto2 = Left(Texto, Quant)
    Texto3 = Space(Quant - Len(Texto2))
    Espaco2 = Texto2 + Texto3
End Function

Function FormataCPF (Texto)
    If IsNull(Texto) Then Exit Function
	FormataCPF=left(texto,3) & "." & mid(texto,4,3) & "." & mid(texto,7,3) & "-" & right(texto,2)
End Function

Function FormataPIS (Texto)
    If IsNull(Texto) Then Exit Function
	FormataPIS=left(texto,3) & "." & mid(texto,4,5) & "." & mid(texto,9,2) & "." & right(texto,1)
End Function


Function ChecaCNPJ(NUMERO, tipo)
   If Len(NUMERO) < 12 Then
      If tipo = 1 Then ChecaCNPJ = False
      If tipo = 2 Then ChecaCNPJ = String(14, "0")
      Exit Function
   End If
   NUMERO = Format(TextoPuro(NUMERO, 2), "00000000000000")
   digito1 = Mid(NUMERO, 13, 1)
   digito2 = Right(NUMERO, 1)
   soma = 0: peso = 2
   For a = 1 To 12
      soma = soma + peso * Val(Mid(NUMERO, 13 - a, 1))
      If peso < 9 Then peso = peso + 1 Else peso = 2
   Next
   nl1 = 11 - (soma Mod 11)
   If nl1 > 9 Then nl1 = 0 Else nl1 = nl1
   peso = 2: soma = 0
   If nl1 = Val(digito1) Then
      For a = 1 To 13
         soma = soma + peso * Val(Mid(NUMERO, 14 - a, 1))
         If peso < 9 Then peso = peso + 1 Else peso = 2
      Next
      nl2 = 11 - (soma Mod 11)
      If nl2 > 9 Then nl2 = 0 Else nl2 = nl2
   Else
      numero2 = Left(NUMERO, 12) & nl1 & Right(NUMERO, 1)
      For a = 1 To 13
         soma = soma + peso * Val(Mid(numero2, 14 - a, 1))
         If peso < 9 Then peso = peso + 1 Else peso = 2
      Next
      nl2 = 11 - (soma Mod 11)
      If nl2 > 9 Then nl2 = 0 Else nl2 = nl2
      'nl2 = digito2
   End If
   If nl1 = Val(digito1) And nl2 = Val(digito2) Then
      If tipo = 1 Then ChecaCNPJ = True
      If tipo = 2 Then ChecaCNPJ = NUMERO
   Else
      If tipo = 1 Then ChecaCNPJ = False
      If tipo = 2 Then ChecaCNPJ = Left(NUMERO, 12) & nl1 & nl2
   End If
End Function

Function separadescricao(Texto, tipo)
    tam = Len(Texto)
    setor       = ""
    desccbo     = ""
    descsumaria = ""
    condicoes   = ""
    formacao    = ""
    fase = 1 : seq=0
    For a = 1 To tam
        letra = Mid(Texto, a, 1) : seq=seq+1
        If letra = "#" Then fase = fase + 1: a=seq
        if letra=chr(10) or letra=chr(13) then letra=""
        Select Case fase
            Case 1
                setor       = setor & letra
            Case 2
                desccbo     = desccbo & letra
            Case 3
                descsumaria = descsumaria & letra
            Case 4
                condicoes   = condicoes & letra
            Case 5
                formacao    = formacao & letra
        End Select
    Next
    Select Case tipo
        Case 1
            separadescricao = setor
        Case 2
            separadescricao = desccbo
        Case 3
            separadescricao = descsumaria
        Case 4
            separadescricao = condicoes
        Case 5
            separadescricao = formacao
    End Select
End Function

Function Veiculo(Texto, tipo)
    tam = Len(Texto)
    cracha=""
	marca =""
    modelo=""
    placa =""
	cor   =""
	ativo =""
    fase = 1:seq=0
    For a = 1 To tam
        letra = Mid(Texto, a, 1) : seq=seq+1
        If letra = "=" Then fase = fase + 1: a=seq
	if letra = "=" then letra=""
        Select Case fase
            Case 1
                cracha= cracha& letra
            Case 2
                marca = marca & letra
            Case 3
                modelo= modelo& letra
            Case 4
                placa = placa & letra
            Case 5
                cor   = cor   & letra
            Case 6
                ativo = ativp & letra
        End Select
    Next
    Select Case tipo
        Case 1
            Veiculo = trim(cracha)
        Case 2
            Veiculo = Trim(marca)
        Case 3
            Veiculo = trim(modelo)
        Case 4
            Veiculo = trim(placa)
        Case 5
            Veiculo = Trim(cor)
        Case 6
            Veiculo = trim(ativo)
    End Select
End Function

Function ChecaPIS(NUMERO, tipo)
'tipo checagem - 1 para verificar se está correto 2 - para informar qual o correto
   NUMERO = TextoPuro(NUMERO, 2)
   digito = Val(Right(NUMERO, 1))
   Saldo = 0: Mult = 2
   For a = 1 To 10
      Saldo = Saldo + Mult * Val(Mid(NUMERO, 11 - a, 1))
      If Mult < 9 Then Mult = Mult + 1 Else Mult = 2
   Next
   Resto = Saldo Mod 11
   digreal = 11 - Resto
   If digreal > 9 Then digreal = 0
   If digreal <> digito Then
      PisReal = Left(NUMERO, 10) & digreal
      If tipo = 1 Then ChecaPIS = False
      If tipo = 2 Then ChecaPIS = PisReal
      Exit Function
   End If
   If tipo = 1 Then ChecaPIS = True
   If tipo = 2 Then ChecaPIS = NUMERO
End Function

Function ChecaCPF(NUMERO, tipo)
'tipo checagem - 1 para verificar se está correto 2 - para informar qual o correto
   'NUMERO = TextoPuro(NUMERO, 2)
   'NUMERO = Format(NUMERO, "00000000000")
   tamanho=len(numero)
   numero=string(11-tamanho,"0")&numero
   'document.form.endereco.value=checacpf2(tempcpf,1)
   If Len(NUMERO) < 10 Then
      If tipo = 1 Then ChecaCPF = False
      If tipo = 2 Then ChecaCPF = String(11, "0")
      Exit Function
   End If
   digito1 = Mid(NUMERO, 10, 1)
   digito2 = Right(NUMERO, 1)
   soma = 0: nlm11 = 2
   For a = 1 To 9
      soma = soma + (a + 1) * cdbl((Mid(NUMERO, 10 - a, 1)))
   Next
   resto0 = 11 - (soma Mod 11)
   If resto0 > 9 Then resto0 = 0 Else resto0 = resto0
   soma = 0
   If resto0 = cint(digito1) Then
      For a = 1 To 10
         soma = soma + (a + 1) * (Mid(NUMERO, 11 - a, 1))
      Next
      resto2 = 11 - (soma Mod 11)
      If resto2 > 9 Then resto2 = 0 Else resto2 = resto2
   Else
      numero2 = Left(NUMERO, 9) & resto0 & Right(NUMERO, 1)
      For a = 1 To 10
         soma = soma + (a + 1) * (Mid(numero2, 11 - a, 1))
      Next
      resto2 = 11 - (soma Mod 11)
      If resto2 > 9 Then resto2 = 0 Else resto2 = resto2
      'resto2 = digito2
   End If
   If resto0 = cint(digito1) And resto2 = cint(digito2) Then
      If tipo = 1 Then ChecaCPF = True
      If tipo = 2 Then ChecaCPF = NUMERO
   Else
      If tipo = 1 Then ChecaCPF = False
      If tipo = 2 Then ChecaCPF = Left(NUMERO, 9) & resto0 & resto2
   End If
End Function

Function ChecaNIT(NUMERO, tipo)
'tipo checagem - 1 para verificar se está correto 2 - para informar qual o correto
   NUMERO = Val(NUMERO)
   digito = Right(NUMERO, 1)
   soma = 0: peso = 3298765432
   If NUMERO = 0 Then
      If tipo = 1 Then ChecaNIT = False
      If tipo = 2 Then ChecaNIT = String(11, "0")
      Exit Function
   End If
   If (Left(NUMERO, 10) < 1060000001 Or Left(NUMERO, 10) > 1069999999) And _
      (Left(NUMERO, 10) < 1090000000 Or Left(NUMERO, 10) > 1199027231) Then
      If tipo = 1 Then ChecaNIT = False
      If tipo = 2 Then ChecaNIT = "INVALIDO"
      Exit Function
   Else
      For a = 1 To 10
         soma = soma + Val(Mid(peso, a, 1)) * Val(Mid(NUMERO, a, 1))
      Next
      calculo = (soma Mod 11)
      If calculo = 0 Or calculo = 1 Then digreal = 0 Else digreal = 11 - calculo
      MsgBox "Calculo: " & calculo & " Digreal " & digreal & " Digito " & digito
      If Val(digito) = Val(digreal) Then
         If tipo = 1 Then ChecaNIT = True
         If tipo = 2 Then ChecaNIT = ChecaNIT
      Else
         If tipo = 1 Then ChecaNIT = False
         If tipo = 2 Then ChecaNIT = Left(NUMERO, 10) & digreal
      End If
   End If
End Function

Function ChecaCEI(NUMERO, tipo)
   NUMERO = Format(Val(NUMERO), "000000000000")
   If Len(NUMERO) <= 11 Then
      If tipo = 1 Then ChecaCEI = False
      If tipo = 2 Then ChecaCEI = String(12, "0")
      Exit Function
   End If
   digito = Right(NUMERO, 1): soma = 0
   peso = 74185216374
   If Val(Left(NUMERO, 2)) < 1 Or Val(Left(NUMERO, 2)) > 29 Then
      If tipo = 1 Then ChecaCEI = False
      If tipo = 2 Then ChecaCEI = String(12, "0")
      Exit Function
   Else
      Select Case Mid(NUMERO, 11, 1)
         Case 1,2,3,4,5
            If tipo = 1 Then ChecaCEI = False
            If tipo = 2 Then ChecaCEI = String(12, "0")
            Exit Function
         Case Else
            For a = 1 To 11
               soma = soma + Val(Mid(peso, a, 1)) * Val(Mid(NUMERO, a, 1))
            Next
            calculo = Right(soma, 2)
            digreal = Val(Left(calculo, 1)) + Val(Right(calculo, 1))
            digreal = 10 - Right(digreal, 1)
            If digreal > 9 Then digreal = 0
            If Val(digito) = Val(digreal) Then
               If tipo = 1 Then ChecaCEI = True
               If tipo = 2 Then ChecaCEI = NUMERO
            Else
               If tipo = 1 Then ChecaCEI = False
               If tipo = 2 Then ChecaCEI = Left(NUMERO, 11) & digreal
            End If
         End Select
   End If
End Function

Function Extenso2(nValor)
    If IsNull(nValor) Or nValor <= 0 Or nValor > 9999999.99 Then
        Exit Function
    End If
    'Declara as variaveis da funcao
    Dim nContador
    Dim nTamanho  'As Integer
    Dim cValor
    Dim cParte
    Dim cFinal    'As String
    ReDim aGrupo(4), aTexto(4) 'As String

    'Define natrizes com extensos parciais
    ReDim aUnid(19) 'As String
    aUnid(1) = "Um "
    aUnid(2) = "Dois "
    aUnid(3) = "Três "
    aUnid(4) = "Quatro "
    aUnid(5) = "Cinco "
    aUnid(6) = "Seis "
    aUnid(7) = "Sete "
    aUnid(8) = "Oito "
    aUnid(9) = "Nove "
    aUnid(10) = "Dez "
    aUnid(11) = "Onze "
    aUnid(12) = "Doze "
    aUnid(13) = "Treze "
    aUnid(14) = "Quatorze "
    aUnid(15) = "Quinze "
    aUnid(16) = "Dezesseis "
    aUnid(17) = "Dezessete "
    aUnid(18) = "Dezoito "
    aUnid(19) = "Dezenove "

    ReDim aDezena(9) 'As String
    aDezena(1) = "Dez "
    aDezena(2) = "Vinte "
    aDezena(3) = "Trinta "
    aDezena(4) = "Quarenta "
    aDezena(5) = "Cinquenta "
    aDezena(6) = "Sessenta "
    aDezena(7) = "Setenta "
    aDezena(8) = "Oitenta "
    aDezena(9) = "Noventa "

    ReDim aCentena(9) 'As String
    aCentena(1) = "Cento "
    aCentena(2) = "Duzentos "
    aCentena(3) = "Trezentos "
    aCentena(4) = "Quatrocentos "
    aCentena(5) = "Quinhentos "
    aCentena(6) = "Seiscentos "
    aCentena(7) = "Setecentos "
    aCentena(8) = "Oitocentos "
    aCentena(9) = "Novecentos "

    'Divide o valor em vários grupos
    cValor = FormatNumber(nValor, 2)
    cValor = replace(cValor,".","")
    cValor = replace(cValor,",",".")
    cValor = numzero(cValor,13)
    aGrupo(1) = Mid(cValor, 2, 3)
    aGrupo(2) = Mid(cValor, 5, 3)
    aGrupo(3) = Mid(cValor, 8, 3)
    aGrupo(4) = "0" + right(cValor,2)

    'Processa cada grupo
    For nContador = 1 To 4
        cParte = aGrupo(nContador)
        if cint(cParte) < 10 then
		nTamanho=1
	elseif cint(cParte) < 100 then
		nTamanho=2
	elseif cint(cParte) < 1000 then
		nTamanho=3
	end if
        'nTamanho = Switch(cint(cParte) < 10, 1, cint(cParte) < 100, 2, cint(cParte) < 1000, 3)
        If nTamanho = 3 Then
            If Right(cParte, 2) <> "00" Then
                aTexto(nContador) = aTexto(nContador) & aCentena(Left(cParte, 1)) & "e "
                nTamanho = 2
            Else
                'aTexto(nContador) = aTexto(nContador) & IIf(Left(cParte, 1) = "1", "Cem ", aCentena(Left(cParte, 1)))
		if Left(cParte, 1) = "1" then temp3="Cem " else	temp3=aCentena(Left(cParte, 1))
		aTexto(nContador) = aTexto(nContador) & temp3
            End If
        End If
        If nTamanho = 2 Then
            If cint(Right(cParte, 2)) < 20 Then
                aTexto(nContador) = aTexto(nContador) & aUnid(Right(cParte, 2))
            Else
                aTexto(nContador) = aTexto(nContador) & aDezena(Mid(cParte, 2, 1))
                If Right(cParte, 1) <> "0" Then
                    aTexto(nContador) = aTexto(nContador) & "e "
                    nTamanho = 1
                End If
            End If
        End If
        If nTamanho = 1 Then
            aTexto(nContador) = aTexto(nContador) & aUnid(Right(cParte, 1))
        End If
    Next

    'Gera o formato final do texto
    'If cint(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And cint(aGrupo(4)) <> 0 Then
    If (aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And cint(aGrupo(4)) <> 0 Then
        'cFinal = aTexto(4) & IIf(cint(aGrupo(4)) = 1, "Centavo", "Centavos")
	if cint(aGrupo(4))=1 then tempf="Centavo" else tempf="Centavos"
	cFinal=aTexto(4) & tempf
    Else
        cFinal = ""
        'cFinal = cFinal & IIf(cint(aGrupo(1)) <> 0, aTexto(1) & IIf(cint(aGrupo(1)) > 1, "Milhões ", "Milhão "), "")
	if cint(aGrupo(1)) > 1 then temp1="Milhões " else temp1="Milhão "
	if cint(aGrupo(1)) <>0 then tempg1=aTexto(1) & temp1
	cFinal = cFinal & tempg1

'        If cint(aGrupo(2) + aGrupo(3)) = 0 Then
        If (aGrupo(2) + aGrupo(3)) = 0 Then
            cFinal = cFinal & "de "
        Else
            'cFinal = cFinal & IIf(cint(aGrupo(2)) <> 0, aTexto(2) & "Mil ", "")
	    if cint(aGrupo(2)) <> 0 then tempt2=aTexto(2) & "Mil " else tempt2=""
	    cFinal = cFinal & tempt2 		
        End If
        cFinal = cFinal & aTexto(3) '& IIf(cint(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "Real ", "Reais ")
'	if cint(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1 then tempg3="Real " else tempg3="Reais "
	if (aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1 then tempg3="Real " else tempg3="Reais "
	cFinal = cFinal & tempg3
        'cFinal = cFinal & IIf(cint(aGrupo(4)) <> 0, "e " & aTexto(4) & IIf(cint(aGrupo(4)) = 1, "Centavo", "Centavos"), "")
	if cint(aGrupo(4)) = 1 then tempg4="Centavo" else tempg4="Centavos"
	if cint(aGrupo(4)) <> 0 then tempt4="e " & aTexto(4) & tempg4 else tempt4=""
	cFinal = cFinal & tempt4
    End If
    Extenso2 = cFinal
    'Extenso = cValor + " " + aGrupo(1) + " " + aGrupo(2) + " " + aGrupo(3) + " " + aGrupo(4) + " " + Str(nTamanho) + " " + aTexto(2) + " " + aTexto(3)
End Function

Function Extenson(nValor)
    If IsNull(nValor) Or nValor <= 0 Or nValor > 9999999 Then
        Exit Function
    End If
    'Declara as variaveis da funcao
    Dim nContador
    Dim nTamanho  'As Integer
    Dim cValor
    Dim cParte
    Dim cFinal    'As String
    ReDim aGrupo(4), aTexto(4) 'As String

    'Define natrizes com extensos parciais
    ReDim aUnid(19) 'As String
    aUnid(1) = "Um "
    aUnid(2) = "Dois "
    aUnid(3) = "Três "
    aUnid(4) = "Quatro "
    aUnid(5) = "Cinco "
    aUnid(6) = "Seis "
    aUnid(7) = "Sete "
    aUnid(8) = "Oito "
    aUnid(9) = "Nove "
    aUnid(10) = "Dez "
    aUnid(11) = "Onze "
    aUnid(12) = "Doze "
    aUnid(13) = "Treze "
    aUnid(14) = "Quatorze "
    aUnid(15) = "Quinze "
    aUnid(16) = "Dezesseis "
    aUnid(17) = "Dezessete "
    aUnid(18) = "Dezoito "
    aUnid(19) = "Dezenove "

    ReDim aDezena(9) 'As String
    aDezena(1) = "Dez "
    aDezena(2) = "Vinte "
    aDezena(3) = "Trinta "
    aDezena(4) = "Quarenta "
    aDezena(5) = "Cinquenta "
    aDezena(6) = "Sessenta "
    aDezena(7) = "Setenta "
    aDezena(8) = "Oitenta "
    aDezena(9) = "Noventa "

    ReDim aCentena(9) 'As String
    aCentena(1) = "Cento "
    aCentena(2) = "Duzentos "
    aCentena(3) = "Trezentos "
    aCentena(4) = "Quatrocentos "
    aCentena(5) = "Quinhentos "
    aCentena(6) = "Seiscentos "
    aCentena(7) = "Setecentos "
    aCentena(8) = "Oitocentos "
    aCentena(9) = "Novecentos "

    'Divide o valor em vários grupos
    cValor = FormatNumber(nValor, 2)
    cValor = replace(cValor,".","")
    cValor = replace(cValor,",",".")
    cValor = numzero(cValor,13)
    aGrupo(1) = Mid(cValor, 2, 3)
    aGrupo(2) = Mid(cValor, 5, 3)
    aGrupo(3) = Mid(cValor, 8, 3)
    aGrupo(4) = "0" + right(cValor,2)

    'Processa cada grupo
    For nContador = 1 To 4
        cParte = aGrupo(nContador)
        if cint(cParte) < 10 then
		nTamanho=1
	elseif cint(cParte) < 100 then
		nTamanho=2
	elseif cint(cParte) < 1000 then
		nTamanho=3
	end if
        'nTamanho = Switch(cint(cParte) < 10, 1, cint(cParte) < 100, 2, cint(cParte) < 1000, 3)
        If nTamanho = 3 Then
            If Right(cParte, 2) <> "00" Then
                aTexto(nContador) = aTexto(nContador) & aCentena(Left(cParte, 1)) & "e "
                nTamanho = 2
            Else
                'aTexto(nContador) = aTexto(nContador) & IIf(Left(cParte, 1) = "1", "Cem ", aCentena(Left(cParte, 1)))
		if Left(cParte, 1) = "1" then temp3="Cem " else	temp3=aCentena(Left(cParte, 1))
		aTexto(nContador) = aTexto(nContador) & temp3
            End If
        End If
        If nTamanho = 2 Then
            If cint(Right(cParte, 2)) < 20 Then
                aTexto(nContador) = aTexto(nContador) & aUnid(Right(cParte, 2))
            Else
                aTexto(nContador) = aTexto(nContador) & aDezena(Mid(cParte, 2, 1))
                If Right(cParte, 1) <> "0" Then
                    aTexto(nContador) = aTexto(nContador) & "e "
                    nTamanho = 1
                End If
            End If
        End If
        If nTamanho = 1 Then
            aTexto(nContador) = aTexto(nContador) & aUnid(Right(cParte, 1))
        End If
    Next

    'Gera o formato final do texto
    If cint(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And cint(aGrupo(4)) <> 0 Then
        'cFinal = aTexto(4) & IIf(cint(aGrupo(4)) = 1, "Centésimo", "Centésimos")
	if cint(aGrupo(4))=1 then tempf="Centésimo" else tempf="Centésimos"
	cFinal=aTexto(4) & tempf
    Else
        cFinal = ""
        'cFinal = cFinal & IIf(cint(aGrupo(1)) <> 0, aTexto(1) & IIf(cint(aGrupo(1)) > 1, "Milhões ", "Milhão "), "")
	if cint(aGrupo(1)) > 1 then temp1="Milhões " else temp1="Milhão "
	if cint(aGrupo(1)) <>0 then tempg1=aTexto(1) & temp1
	cFinal = cFinal & tempg1

        If cint(aGrupo(2) + aGrupo(3)) = 0 Then
            cFinal = cFinal & "de "
        Else
            'cFinal = cFinal & IIf(cint(aGrupo(2)) <> 0, aTexto(2) & "Mil ", "")
	    if cint(aGrupo(2)) <> 0 then tempt2=aTexto(2) & "Mil " else tempt2=""
	    cFinal = cFinal & tempt2 		
        End If
        cFinal = cFinal & aTexto(3) '& IIf(cint(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "Real ", "Reais ")
	if cint(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1 then tempg3=" " else tempg3=" "
	cFinal = cFinal & tempg3
        'cFinal = cFinal & IIf(cint(aGrupo(4)) <> 0, "e " & aTexto(4) & IIf(cint(aGrupo(4)) = 1, "Centésimo", "Centésimos"), "")
	if cint(aGrupo(4)) = 1 then tempg4="Centésimo" else tempg4="Centésimos"
	if cint(aGrupo(4)) <> 0 then tempt4="e " & aTexto(4) & tempg4 else tempt4=""
	cFinal = cFinal & tempt4
    End If
    Extenson = cFinal
    'Extenso = cValor + " " + aGrupo(1) + " " + aGrupo(2) + " " + aGrupo(3) + " " + aGrupo(4) + " " + Str(nTamanho) + " " + aTexto(2) + " " + aTexto(3)
End Function

Function Quadradinho(Texto, nQ)
    TR = "" & Chr(124) & ""
    TRi = Chr(124) & ""
    TRf = "" & Chr(124)
    tamanho = Len(Texto)
    If tamanho < nQ Then
        Texto2 = Espaco2(Texto, nQ)
    Else
        Texto2 = Left(Texto, nQ)
    End If
    Texto3 = TRi
    For L = 1 To (nQ - 1)
        letra = Mid(Texto2, L, 1)
        Texto3 = Texto3 + letra + TR
    Next
    Texto3 = Texto3 + Mid(Texto2, nQ, 1) + TRf
    Quadradinho = Texto3
End Function

Function Quadradinho2(Texto, nQ)
    TR = " " & Chr(124) & " "
    TRi = Chr(124) & " "
    TRf = " " & Chr(124)
    tamanho = Len(Texto)
    If tamanho < nQ Then
        Texto2 = Espaco2(Texto, nQ)
    Else
        Texto2 = Left(Texto, nQ)
    End If
    Texto3 = TRi
    For L = 1 To (nQ - 1)
        letra = Mid(Texto2, L, 1)
		if letra=" " then letra="&nbsp;"
        Texto3 = Texto3 + letra + TR
    Next 
    Texto3 = Texto3 + Mid(Texto2, nQ, 1) + TRf
    Quadradinho2 = Texto3
End Function

Function Espaco3(Texto, Quant)
    If IsNull(Texto) Then Exit Function
    Texto2 = Left(Texto, Quant)
    Texto3 = String(Quant - Len(Texto2),"_")
    Espaco3 = Texto2 + Texto3
End Function

Function fData3(data)
    if isnull(data) then exit function
    temp=""&numzero(day(data),2)&"/"&numzero(month(data),2)&"/"&numzero(year(data),4)&""
    fData3=temp
End Function

Function espacosd(Texto, nQ)
    TR = "&nbsp;"
    TRi = ""
    TRf = ""
    tamanho = Len(Texto)
    If tamanho < nQ Then
        Texto2 = Espaco2(Texto, nQ)
    Else
        Texto2 = Left(Texto, nQ)
    End If
    Texto3 = TRi
    For L = 1 To (nQ - 1)
        letra = Mid(Texto2, L, 1)
	if letra=" " then letra="&nbsp;"
        Texto3 = Texto3 + letra + TR
    Next 
    Texto3 = Texto3 + Mid(Texto2, nQ, 1) + TRf
    espacosd = Texto3
End Function

Function fQuadro3(Texto, nQ)
    TR = "<font color=silver>" & Chr(124) & "</font>"
    TRi = "<font color=silver>" & Chr(124) & "</font>"
    TRf = "" & "<font color=silver>" & Chr(124) & "</font>"
    tamanho = Len(Texto)
    If tamanho < nQ Then
        Texto2 = Espaco3(Texto, nQ)
    Else
        Texto2 = Left(Texto, nQ)
    End If
    Texto3 = TRi
    For L = 1 To (nQ - 1)
        letra = Mid(Texto2, L, 1)
		if letra=" " then letra="&nbsp;"
        Texto3 = Texto3 + letra + TR
    Next 
    Texto3 = Texto3 + Mid(Texto2, nQ, 1) + TRf
    fQuadro3 = Texto3
End Function

Function PrimeiroNome(numero)
    newstring=""
    if numero="" then PrimeiroNome=0:exit function
    for a=1 to len(numero)
		p=mid(numero,a,1)
		if p=" " then exit for
		newstring=newstring & p
    next
    PrimeiroNome=newstring
End Function

%>