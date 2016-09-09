<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a93")="N" or session("a93")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"

%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Pesquisa Curriculos</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<script language="javascript" type="text/javascript">
<!--
/****************************************************
     Author: Eric King
     Url: http://redrival.com/eak/index.shtml
     This script is free to use as long as this info is left in
     Featured on Dynamic Drive script library (http://www.dynamicdrive.com)
****************************************************/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}

function toggleAll(cb) 
{
        var val = cb.checked;
        var frm = document.forms[0];
        var len = frm.elements.length;
        var i=0;
        for( i=0 ; i<len ; i++) 
        {
                if (frm.elements[i].type=="checkbox" && frm.elements[i]!=cb) 
                {
                        frm.elements[i].checked=val;
                }
        }
}
// -->
</script>
</head>
<body>
<!-- -->
<table><tr><td>
<!-- -->
<form method="POST" action="banco_funcao.asp" name="form">

<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("mysqlfieo")
'conexao.open "Driver={MySQL ODBC 3.51 Driver}; Server=localhost; Port=3306; Option=0; Socket=; Stmt=; Database=rhonline2; Uid=root; Pwd="
'conexao.open "Driver={MySQL ODBC 3.51 Driver}; Server=colossus2.fieo.br; Port=3306; Option=0; Socket=; Stmt=; Database=website; Uid=rh; Pwd=!@#qaz"

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
rs.CursorLocation=3
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs2.CursorLocation=3

if request.form("ordem")="" then ordem="c" else ordem=request.form("ordem")
button2=request.form("button2")
if session("usuariomaster")="02379" then response.write request.form & "<br><br>"

%>
Opções de pesquisa:<br>
<input type="radio" name="ordem" value="c" <%if ordem="c" then response.write "checked"%> >Cadastros mais recentes na frente<br>
<input type="radio" name="ordem" value="n" <%if ordem="n" then response.write "checked"%> >Por ordem de nome<br>
<Br>

Por nome do candidato
<span style="font-size:9pt;background:silver;width:490;text-align:center">
<%
for a=65 to 90
	idbutton=a-64
	if request.form("button"&a)=chr(a) then letra=chr(a) else letra=letra
	if request.form("button"&a)=chr(a) then marca="bold" else marca="none"
%>
	<input type="submit" value="<%=chr(a)%>" name="button<%=a%>" style="border:1px solid gray;background:silver;color:black;font-weight:<%=marca%>" />
<%
next
%>
</span>
<br>
<Br>

Por disciplina
<span style="font-size:9pt;background:silver;text-align:center">
<select name="disciplina" size="1" >
	<option value=0>Selecione uma disciplina</option>
<%
sqld="SELECT id_disciplina, disciplina, (select count(cpf) from tb_rh_rel_candidato_disciplina where disciplina=d.id_disciplina) as candidatos " & _
"FROM tb_rh_disciplina d order by disciplina "
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
if cstr(request.form("disciplina"))=cstr(rs2("id_disciplina")) then temp="selected" else temp=""
%>
	<option value="<%=rs2("id_disciplina")%>" <%=temp%>  ><%=rs2("disciplina") & " (" & rs2("candidatos") & ")"%></option>
<%	
rs2.movenext
loop
rs2.close
%>
</select>
<input type="submit" value="Procurar" name="button2" />
</span>
<br>
<Br>

<%
teste=0
if teste=1 then
%>
<span style="font-size:8pt;font-weight:normal">Palavra chave:
<input type="text" name="palavra1" size="40" maxlength="40" value="<%=request.form("palavra1")%>" />
<input type="submit" value="Procurar" name="button1" />

<br><br>Procurar em: 
<br>
	<input type="checkbox" name="chkfamilia"  value="on" <%if request.form("chkfamilia")="on" or request.form="" then response.write "checked"%> />Famílias
	<input type="checkbox" name="chkocupacao" value="on" <%if request.form("chkocupacao")="on" or request.form=""  then response.write "checked"%>/>Ocupações
	<input type="checkbox" name="chksinonimo" value="on" <%if request.form("chksinonimo")="on" or request.form=""  then response.write "checked"%>/>Sinônimos
<br>
	<input type="checkbox" name="chkatividades" value="on" <%if request.form("chkatividades")="on" then response.write "checked"%> />Atividades
	<input type="checkbox" name="chkrecursos"   value="on" <%if request.form("chkrecursos")="on" then response.write "checked"%>/>Recursos de Trabalho
	<input type="checkbox" name="chkdescricao"  value="on" <%if request.form("chkdescricao")="on" then response.write "checked"%>/>Descrição
<br>
	<input type="checkbox" name="chkcondicoes" value="on" <%if request.form("chkcondicoes")="on" then response.write "checked"%>/>Condições do Trabalho
	<input type="checkbox" name="chkformacao"  value="on" <%if request.form("chkformacao")="on" then response.write "checked"%>/>Formação e Experiência
	<input type="checkbox" name="checkall" onclick="toggleAll(this)" id="Checkbox1" /><font color=green>Selecionar todos</font>
</span>

<%
if request.form("button1")<>"" and len(request.form("palavra1"))>0 then
	sql0=""
	if request.form("chkfamilia")="on" then sql1="select cbo=codigo_familia_cbo, nome=nome_familia, id_familia, Tipo='Família' from cbo_4familias_ocupacionais where nome_familia like '%" & request.form("palavra1") & "%' " else sql1=""
	if request.form("chkocupacao")="on" then sql2="select cbo=nu_codigo_cbo, nome=nm_ocupacao, id_familia, Tipo='Ocupação' from cbo_5ocupacoes where nm_ocupacao like '%" & request.form("palavra1") & "%' " else sql2=""
	if request.form("chksinonimo")="on" then sql3="select cbo=nu_codigo_cbo, nome=nm_titulo, id_familia, Tipo='Sinônimo' from cbo_5sinonimos s inner join cbo_5ocupacoes o on o.id_ocupacao=s.id_ocupacao where nm_titulo like '%" & request.form("palavra1") & "%' " else sql3=""
	if request.form("chkatividades")="on" then sql4="select cbo=codigo_familia_cbo, nome=nome_atividade, f.id_familia, tipo='Atividades' from cbo_9atividades a inner join cbo_9gacs g on g.id_gac=a.id_gac inner join cbo_4familias_ocupacionais f on f.id_familia=g.id_familia where nome_atividade like '%" & request.form("palavra1") & "%' union all select cbo=codigo_familia_cbo, nome=nome_gac, f.id_familia, tipo='Área de atividades' from cbo_9gacs g inner join cbo_4familias_ocupacionais f on f.id_familia=g.id_familia where nome_gac like '%" & request.form("palavra1") & "%' " else sql4=""
	if request.form("chkrecursos")="on" then sql5="select cbo=codigo_familia_cbo, nome=nm_recurso_trabalho, f.id_familia, tipo='Recurso de trabalho' from cbo_9recursos_trabalho r inner join cbo_4familias_ocupacionais f on f.id_familia=r.id_familia where nm_recurso_trabalho like '%" & request.form("palavra1") & "%' " else sql5=""
	if request.form("chkdescricao")="on" then sql6="select cbo=codigo_familia_cbo, nome=te_descricao_sumaria, id_familia, tipo='Descrição' from cbo_4familias_ocupacionais where te_descricao_sumaria like '%" & request.form("palavra1") & "%' " else sql6=""
	if request.form("chkcondicoes")="on" then sql7="select cbo=codigo_familia_cbo, nome=te_cond_geral_exerc, id_familia, tipo='Condições de trabalho' from cbo_4familias_ocupacionais where te_cond_geral_exerc like '%" & request.form("palavra1") & "%' " else sql7=""
	if request.form("chkformacao")="on" then sql8="select cbo=codigo_familia_cbo, nome=te_formacao_exper, id_familia, tipo='Formação e Experiência' from cbo_4familias_ocupacionais where te_formacao_exper like '%" & request.form("palavra1") & "%' " else sql8=""
	sql0="select cbo='0000', nome='', id_familia=0, tipo='' "
	if sql1<>"" and sql0<>"" then sql0=sql0 & " union " & sql1
	if sql2<>"" and sql0<>"" then sql0=sql0 & " union " & sql2
	if sql3<>"" and sql0<>"" then sql0=sql0 & " union " & sql3
	if sql4<>"" and sql0<>"" then sql0=sql0 & " union " & sql4
	if sql5<>"" and sql0<>"" then sql0=sql0 & " union " & sql5
	if sql6<>"" and sql0<>"" then sql0=sql0 & " union " & sql6
	if sql7<>"" and sql0<>"" then sql0=sql0 & " union " & sql7
	if sql8<>"" and sql0<>"" then sql0=sql0 & " union " & sql8
	sql1="select * from (" & sql0 & ") z where cbo<>'0000' order by nome "
end if	
%>
<%
end if 'teste=1
%>

<table border="1" cellpadding="1" width="550" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo>Curriculos</td>
</tr>
<tr>
	<td class="campor" valign=top><%=nomef%>
<%
if letra<>"" or button2<>"" then
if ordem="c" then sqlo="order by data_cadastro desc " else sqlo="order by c.nome "

if letra<>"" then 
	sqlb="SELECT distinct c.nome, p.funcao, p.cpf, p.salario, data_cadastro, nascimento, bairro, cidade, uf, tel_residencial, tel_celular, email, observacoes " & _
	"FROM tb_rh_pretensao p inner join tb_rh_candidato c on c.cpf=p.cpf where p.funcao=1 " & _
	"and c.nome like '" & letra & "%' and nascimento>0 " & sqlo
end if
if button2<>"" then
	sqlb="SELECT distinct c.nome, p.funcao, p.cpf, p.salario, data_cadastro, nascimento, bairro, cidade, uf, tel_residencial, tel_celular, email, observacoes " & _
	"FROM tb_rh_pretensao p inner join tb_rh_candidato c on c.cpf=p.cpf where p.funcao=1 " & _
	"and nascimento>0 " & sqlo
	sqlb="SELECT distinct c.nome, p.funcao, c.cpf, p.salario, data_cadastro, nascimento, bairro, cidade, uf, tel_residencial, tel_celular, email, observacoes " & _
	"FROM tb_rh_candidato c inner join tb_rh_rel_candidato_disciplina d on d.cpf=c.cpf " & _
	"left join tb_rh_pretensao p on c.cpf=p.cpf " & _
	"where nascimento>0 and d.disciplina=" & request.form("disciplina") & " " & sqlo
end if	
'response.write sqlb
rs.Open sqlb, ,adOpenStatic, adLockReadOnly
do while not rs.eof
%>
<table border=1 width=530 style='border-collapse:collapse'>
<tr>
	<td class=campo valign=top><font color="green"><b>Nome:</b></font><br><b><%=ucase(rs("nome"))%></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<a class=r href="banco_curriculo.asp?codigo=<%=rs("cpf")%>" onclick="NewWindow(this.href,'form_curriculo','695','450','yes','center');return false" onfocus="this.blur()">
	<img src="../images/Leaf.gif" width="16" height="16" border="0" alt="Visualizar o curriculo"></a></td>
	<td class=campo valign=top width=70><font color="green"><b>Cadastro:</b></font><br><%=rs("data_cadastro")%></td>
</tr>
</table>
<table border=1 width=530 style='border-collapse:collapse'>
<tr>
	<td class=campo valign=top width=50><font color="green"><b>Pretensão:</b></font><br><%=rs("salario")%></td>
	<td class=campo valign=top><font color="green"><b>Habilidades:</b></font><br>
<%	
sqlh="SELECT habilidade FROM tb_rh_pretensao t where cpf='" & rs("cpf") & "' and funcao=1 " 
rs2.Open sqlh, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
	response.write rs2("habilidade")
	if rs.recordcount>1 and rs.absoluteposition<rs.recordcount then response.write "<Br>"
rs2.movenext
loop
rs2.close
	
%>	
	</td>
</tr>
</table>
<table border=1 width=530 style='border-collapse:collapse'>
<tr>
	<td class=campo valign=top width=90><font color="green"><b>Nascimento:</b></font><br><%=rs("nascimento")%> (<%=int((now()-rs("nascimento"))/365.25)%>)</td>
	<td class=campo valign=top><font color="green"><b>Endereço:</b></font><br><%=rs("bairro") & " " & rs("cidade") & " " & rs("uf")%></td>
	<td class=campo valign=top><font color="green"><b>Telefone:</b></font><br><%=rs("tel_residencial") & " " & rs("tel_celular")%></td>
</tr>
</table>
<table border=1 width=530 style='border-collapse:collapse'>
<tr>
	<td class=campo valign=top width=50><font color="green"><b>Email:</b></font><br><%=rs("email")%></td>
	<td class=campo valign=top><font color="green"><b>Observações:</b></font><br><%=rs("observacoes")%></td>
</tr>
</table>


<hr style="color:blue"> 
<%
rs.movenext
loop
rs.close
end if
%>
	
	</td>
</tr>

</table>

<%
''*************** inicio teste **********************
'response.write "<table border='1' cellpadding='1' cellspacing='2' style='border-collapse:collapse'>"
'response.write "<tr>"
'rs.movefirst
'for a= 0 to rs.fields.count-1
'	response.write "<td class=titulor>" & ucase(rs.fields(a).name) & "<br>" & a & "</td>"
'next
'response.write "</tr>"
'do while not rs.eof 
'response.write "<tr>"
'for a= 0 to rs.fields.count-1
'	response.write "<td class="campor" nowrap>" & rs.fields(a) & "</td>"
'next
'response.write "</tr>"
'rs.movenext
'loop
'response.write "</table>"
'response.write "<p>"
'*************** fim teste **********************%>

<%

set rs=nothing
conexao.close
set conexao=nothing

%>

</form>
<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="../images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>

