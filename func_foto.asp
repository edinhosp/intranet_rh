<%
set conexao=server.createobject ("ADODB.Connection")
conexao.open Application("Consql")
set rs=server.createobject ("ADODB.Recordset")
'set rs.ActiveConnection = conexao

chapa=request("chapa")
Response.Expires = 0
Response.Buffer = TRUE
Response.Clear
Response.ContentType = "image/jpeg"

rs.open "SELECT F.CHAPA, P.CODIGO, P.IDIMAGEM, G.IMAGEM " & _
"FROM PFUNC f, PPESSOA p, GIMAGEM g " & _
"WHERE P.IDIMAGEM = G.ID AND F.CODPESSOA = P.CODIGO " & _
"AND F.CHAPA='" & chapa & "'",conexao

Response.BinaryWrite rs("imagem")
Response.End
rs.close
conexao.close
set rs=nothing
set conexao=nothing
%>