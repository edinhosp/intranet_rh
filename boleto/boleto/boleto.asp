<!--#include file="funcoes.inc" -->
<%
connstr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("dados/dados.mdb") & ";"
Set oRec = Server.CreateObject("ADODB.RecordSet")
oRec.Open "select * from vw_boleto where 'cod =' " & request("id"), connstr, 1, 3, 1

valor_doc = oRec("valor")
dt_doc = oRec("dt_doc")
dt_venc = oRec("dt_venc")
banco = "341"
moeda= "9"
agencia = "0278"
conta = "55304"
dv_conta = "5"
carteira = "175"
num_doc = oRec("num_doc")
nossonumero = oRec("nosso_num")
dv_nossonumero = Calculo_DV10(agencia & conta & carteira & nossonumero)
cod_barra = Monta_CodBarras()
linha_dig = Linha_Digitavel(cod_barra)
nossonumero = carteira & "/" & nossonumero & "-" & dv_nossonumero
sacado = oRec("nome")
end_sacado = oRec("endereco") & "<br>CEP " & oRec("cep") & " - " & oRec("cidade") _
			& " / " & oRec("uf")
cod_cli = oRec("cod")
oRec.Close

%><HTML>
<HEAD>
<TITLE>SuperASP - Portal do programador ASP | BOLETO</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<META NAME="description" CONTENT="Hospedagem de sites com suporte para Access, MySQL, PHP e ASP. Todos os nossos planos de hospedagem contam com relatório de visitas WebTrends, totalmente em português. Oferecemos também apenas reserva de domínios, sem hospedagem.">
<META NAME="keywords" CONTENT="hospedagem de sites, web hosting, co-location, NT, ASP, PHP, Access, MySQL, website host, desconto, registro de domínios, transferência de domínios, domínios, webtrends, plano de parceria">
</HEAD>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="640">
  <tr> 
    <td width="320" height="50"><img src="images_boleto/logo_tw.gif" width="180" height="71"></td>
    <td align="right" valign="bottom" width="320"><font face="Arial" size="3"><b>Recibo do Sacado</b></font></td>
  </tr>
</table>
<table width="640" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF">
  <tr> 
    <td width="43%"><font face="Arial" size="1">&nbsp;Cedente<br>
      </font> <font face="Arial" size="2">SuperASP Ltda</font></td>
    <td width="21%"> 
      <font face="Arial" size="1">&nbsp;Agência/Cod. Cedente<br></font>
	  <div align="center"><font face="Arial" size="2"><%=agencia & "/" & conta & "-" & dv_conta%></font></div></td>
    <td width="18%"><font face="Arial" size="1">&nbsp;Data do Documento<br></font> 
      <div align="center"><font face="Arial" size="2"><%=dt_doc%></font></div></td>
    <td width="18%"> 
      <font face="Arial" size="1">&nbsp;Vencimento<br></font>
      <div align="right"><font face="Arial" size="2"><b><%=dt_venc%></b></font></div></td>
  </tr>
  <tr valign="top">
	<td width="43%"><font face="Arial" size="1">&nbsp;Sacado</font><br>
      <font face="Arial" size="2">&nbsp;<%=sacado%></font></td>
    <td width="21%"> 
      <font face="Arial" size="1">&nbsp;Número Documento<br></font>
	  <div align="center"><font face="Arial" size="2"><%=num_doc%></font></div>
  </td>
    <td width="18%"><font face="Arial" size="1">&nbsp;Nosso Número<br></font> 
      <div align="center"><font face="Arial" size="2"><%=nossonumero%></font></div>
  </td>
    <td width="18%"><font face="Arial" size="1">&nbsp;Valor do Documento<br></font>
    <div align="right"><font size="2" face="Arial"><b><%=FormatNumber(valor_doc,2)%></b></font></div></td>
  </tr>
  <tr> 
    <td colspan="4"><font face="Arial" size="1">&nbsp;Demonstrativo</font><br>
      <font face="Arial" size="2">&nbsp;Serviços de Internet solicitado no site 
      www.superasp.com.br<br>
      <br>
      </font></td>
  </tr>
</table>
<table border="0" cellpadding="2" cellspacing="0" width="640">
  <tr> 
    <td align="right" width="100%"><b><font face="Arial" size="2">Autenticação mecânica</font></b><br>
    </td>
  </tr>
  <tr> 
    <td align="middle" width="100%" height="17">&nbsp;</td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="640">
  <tr valign="top"> 
    <td height="30"><img src="images_boleto/tesoura.gif" width="640" height="8"></td>
  </tr>
</table>
<table border="0" cellPadding="0" cellSpacing="0" width="640">
  <tr>      
    <td width="112" valign="bottom"><img src="images_boleto/341.jpg" width="27" height="26" align="absmiddle"> 
      <font face="Arial" size="1"><b>Banco Ita&uacute; S.A.</b></font></td>
    <td width="66" align="center" valign="bottom"><img src="images_boleto/traco.gif" width="3" height="23" align="absbottom">&nbsp;<font face="Arial" size="4" color="#999999"><b>341-7</b></font>&nbsp;<img src="images_boleto/traco.gif" width="3" height="23" align="absbottom"> 
    </td>
    <td align="right" valign="bottom" width="462" nowrap><font face="Arial" size="3"><%=linha_dig%>&nbsp;</font></td>
  </tr>
</table>
<table border="1" cellPadding="1" cellSpacing="0" width="640" bordercolorlight="#FFFFFF">
  <tr>
    <td colspan="6"> 
      <table border="0" cellpadding="0" cellSpacing="0" width="100%">
        <tr> 
          <td valign="top" width="110"><font face="Arial" size="1">&nbsp;Local 
            de Pagamento</td>
          <td valign="middle" width="380"><font face="Arial" size="2"><b>Até o vencimento, preferencialmente no ITA&Uacute; ou BANERJ<br>
            Ap&oacute;s o vencimento, somente no ITA&Uacute; ou BANERJ</b></font></td>
        </tr>
      </table>
    </td>
    <td width="140"> 
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%" bordercolorlight="#FFFFFF">
        <tr> 
          <td><font face="Arial" size="1">&nbsp;Vencimento</font></td>
        </tr>
        <tr> 
          <td align="right"><font face="Arial" size="3"><b><%=dt_venc%>&nbsp;</b></font></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colSpan="6" vAlign="top"> 
      <table border="0" cellPadding="0" cellSpacing="0" width="100%">
        <tr> 
          <td valign="top"><font face="Arial" size="1">&nbsp;Cedente</font></td>
        </tr>
        <tr> 
          <td valign=top><font face="Arial" size="2">&nbsp;SuperASP Ltda.</font></td>
        </tr>
      </table>
    </td>
    <td width="140" valign="top"> 
      <table border="0" cellPadding="0" cellSpacing="0" width="100%">
        <tr>
          <td><font face="Arial" size="1">&nbsp;Agência/Código Cedente</font></td>
        </tr>
        <tr> 
          <td align="right"><font size="2" face="Arial"><%=agencia & "/" & conta & "-" & dv_conta%>&nbsp;</font></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td width="103" valign="top"> 
      <table border="0" cellPadding="0" cellSpacing="0" width="100%">
        <tr> 
          <td><font face="Arial" size="1">&nbsp;Data do documento</font></td>
        </tr>
        <tr> 
          <td align="center"><font face="Arial" size="2"><%=dt_doc%></font></td>
        </tr>
      </table>
    </td>
    <td colSpan="2" valign="top">
		<table border="0" cellPadding="0" cellSpacing="0" width="100%">
        <tr> 
          <td><font face=arial size=1>&nbsp;No. do documento</font></td>
        </tr>
        <tr> 
          <td align="center"><font face=arial size=2><%=num_doc%></font></td>
        </tr>
      </table>
    </td>
    <td width="62" valign="top"> 
      <table border="0" cellPadding="0" cellSpacing="0" width="98%">
        <tr> 
          <td><font face="Arial" size="1">&nbsp;Espécie doc</font></td>
        </tr>
        <tr> 
          <td align="center"><font face="Arial" size="2">DS</font></td>
        </tr>
      </table>
    </td>
    <td width="47" valign="top"> 
      <table border="0" cellPadding="0" cellSpacing="0" width="100%">
        <tr> 
          <td><font face="Arial" size="1">&nbsp;Aceite</font></td>
        </tr>
        <tr> 
          <td align="center"><font face="arial" size="2">N</font></td>
        </tr>
      </table>
    </td>
    <td width="126" valign="top"><font face="Arial" size="1">&nbsp;Data Processamento</font></td>
    <td width="140"> 
      <table border="0" cellPadding="0" cellSpacing="0" width="100%">
        <tr> 
          <td><font face="Arial" size="1">&nbsp;Nosso Número</font></td>
        </tr>
        <tr> 
          <td align="righ"t><font face="Arial" size="2"><%=nossonumero%>&nbsp;</font></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td valign="top" width="103"><font face="Arial" size="1">&nbsp;Uso do Banco</font></td>
    <td valign="top" width="64"> 
      <table border="0" cellPadding="0" cellSpacing="0" width="100%" align="center">
        <tr valign="top"> 
          <td><font face="arial" size="1">&nbsp;Carteira</font></td>
        </tr>
        <tr valign="top"> 
          <td align="center"><font face="arial" size="2"><%=carteira%></font></td>
        </tr>
      </table>
    </td>
    <td valign="top" width="68"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr valign="top"> 
          <td><font face="Arial" size="1">&nbsp;Moeda</font></td>
        </tr>
        <tr valign="top"> 
          <td align="center"><font face="Arial" size="2">R$</font></td>
        </tr>
      </table>
    </td>
    <td colSpan="2" valign="top"><font face="Arial" size="1">Quantidade</font></td>
    <td valign="top" width="126"><font face="Arial" size="1">&nbsp;Valor</font></td>
    <td width="140" valign="top"> 
      <table border="0" cellPadding="0" cellSpacing="0" height="100%" width="100%">
        <tr> 
          <td><font face="Arial" size="1">&nbsp;(=) Valor do Documento</font></td>
        </tr>
        <tr> 
          <td align="right" height="50%"><font face="Arial" size="2"><b><%=FormatNumber(valor_doc,2)%></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colSpan="6" vAlign="top" rowspan="5"> 
      <table border="0" cellPadding="0" cellSpacing="0" width="100%">
        <tr valign="center"> 
          <td colspan="2"><font face="Arial" size="1">&nbsp;Instruções</font></td>
        </tr>
        <tr vAlign="top"> 
          <td width="9"></td>
          <td width="481"><font face="Arial" size="2"><b>Não receber após o vencimento.</b></font></td>
        </tr>
      </table>
    </td>
    <td height="23" valign="top" width="140"><font face="Arial" size="1">&nbsp;(-) Descontos/Abatimento<br></font></td>
  </tr>
  <tr> 
    <td height="23" valign="top" width="140"><font face="Arial" size="1">&nbsp;(-) Outras 
      Deduções<br></font></td>
  </tr>
  <tr> 
    <td height="23" valign="top" width="140"><font face="Arial" size="1">&nbsp;(+) Mora/Multa 
      <br></font></td>
  </tr>
  <tr> 
    <td height="23" valign="top" width="140"><font face="Arial" size="1">&nbsp;(+) Outros 
      Acréscimos<br></font></td>
  </tr>
  <tr> 
    <td height="30" valign="top" width="140"><font face="Arial" size="1">&nbsp;(=) Valor 
      Cobrado<br></font></td>
  </tr>
  <tr> 
    <td colSpan="5"> 
      <table border="0" cellPadding="0" cellSpacing="0" width="100%">
        <tr> 
          <td vAlign=top width="44"><font face="Arial" size="1">&nbsp;Sacado 
            </font></td>
          <td vAlign=top width="316" height="40"><font face="Arial" size="1"><%=sacado%><BR>
            <%=end_sacado%></font></td>
        </tr>
      </table>
    </td>
    <td align="middle" colSpan="2" vAlign="bottom"><b><font face="Arial" size="3">&nbsp;Ficha de Compensação</font></b></td>
  </tr>
</table>
<table border="0" cellPadding="0" cellSpacing="0" width="637" height="86">
  <tr> 
    <td width="422" height="68" align="left" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="36" width="4%" valign="bottom" nowrap><%
sMyI25 = I25Encode(cod_barra)
bBAR = true
For iPos = 1 To Len(sMyI25)
	If (bBar) Then 
		sGIF = Mid(sMyI25, iPos, 1) & "b.gif"
	Else
		sGIF = Mid(sMyI25, iPos, 1) & "s.gif"
	End If
	Response.Write "<IMG SRC=""images_boleto/" & sGIF & """>"
	bBar = Not bBar
Next
%></td>
          <td height="53" width="96%" valign="bottom">&nbsp;</td>
        </tr>
      </table>
	</td>
    <td align="center" vAlign="top" width="253" height="68"><font face="Arial" size="1">Autenticação Mecânica</font></td>
  </tr>
  <tr valign="top"> 
    <td colSpan="2"><img src="images_boleto/tesoura.gif" width="640" height="8"></td>
  </tr>
  </table>
</body>
</html>
