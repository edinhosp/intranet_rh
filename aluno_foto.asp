<%
set conexao=server.createobject ("ADODB.Connection")
conexao.open Application("Consql")
set rs=server.createobject ("ADODB.Recordset")
'set rs.ActiveConnection = conexao

idimagem=request("id")
Response.Expires = 0
Response.Buffer = TRUE
Response.Clear
Response.ContentType = "image/jpg"

rs.open "select id, codsistema, imagem FROM gimagem WHERE id=" & idimagem & "",conexao

Response.BinaryWrite rs("imagem")
Response.End
rs.close
conexao.close
set rs=nothing
set conexao=nothing
%>
