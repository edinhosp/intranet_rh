<%
Function numzero(numero,zeros)
	if zeros=0 or isnull(zeros) then zeros=len(numero)
	tamanho=len(numero)
	numzero=string(zeros-tamanho,"0")&numero
End Function

Function DtAccess(data)
    if isnull(data) then exit function
    temp=""&month(data)&"/"&day(data)&"/"&year(data)&""
    DtAccess=temp
End Function

%>