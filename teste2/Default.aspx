<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:GridView ID="GridView1" runat="server" 
            DataKeyNames="id_modulo" DataSourceID="SqlDataSource1" CellPadding="4" 
            ForeColor="#333333" GridLines="None" 
            AutoGenerateColumns="False">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:BoundField DataField="id_modulo" HeaderText="id_modulo" ReadOnly="True" 
                    SortExpression="id_modulo" />
                <asp:BoundField DataField="nome_modulo" HeaderText="nome_modulo" 
                    SortExpression="nome_modulo" />
                <asp:BoundField DataField="sigla_modulo" HeaderText="sigla_modulo" 
                    SortExpression="sigla_modulo" />
                <asp:BoundField DataField="inicio" HeaderText="inicio" 
                    SortExpression="inicio" />
                <asp:BoundField DataField="versao" HeaderText="versao" 
                    SortExpression="versao" />
            </Columns>
            <EditRowStyle BackColor="#2461BF" />
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#EFF3FB" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
            ConnectionString="<%$ ConnectionStrings:servico1ConnectionString1 %>" 
            ProviderName="<%$ ConnectionStrings:servico1ConnectionString1.ProviderName %>" 
            SelectCommand="SELECT * FROM soi_modulos"></asp:SqlDataSource>
    
    </div>
    </form>
</body>
</html>
