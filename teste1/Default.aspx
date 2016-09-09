<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:FormView ID="FormView1" runat="server" DataKeyNames="id_autonomo" 
            DataSourceID="SqlDataSource1" EnableModelValidation="True">
            <InsertItemTemplate>
                nome autonomo:
                <asp:TextBox ID="nome_autonomoTextBox" runat="server" 
                    Text='<%# Bind("nome_autonomo") %>' />
                <br />
                dtnascimento:
                <asp:TextBox ID="dtnascimentoTextBox" runat="server" 
                    Text='<%# Bind("dtnascimento") %>' />
                <br />
                sexo:
                <asp:TextBox ID="sexoTextBox" runat="server" Text='<%# Bind("sexo") %>' />
                <br />
                tipo_prestacao:
                <asp:TextBox ID="tipo_prestacaoTextBox" runat="server" 
                    Text='<%# Bind("tipo_prestacao") %>' />
                <br />
                nacionalidade:
                <asp:TextBox ID="nacionalidadeTextBox" runat="server" 
                    Text='<%# Bind("nacionalidade") %>' />
                <br />
                estado_civil:
                <asp:TextBox ID="estado_civilTextBox" runat="server" 
                    Text='<%# Bind("estado_civil") %>' />
                <br />
                cpf:
                <asp:TextBox ID="cpfTextBox" runat="server" Text='<%# Bind("cpf") %>' />
                <br />
                nit:
                <asp:TextBox ID="nitTextBox" runat="server" Text='<%# Bind("nit") %>' />
                <br />
                rg:
                <asp:TextBox ID="rgTextBox" runat="server" Text='<%# Bind("rg") %>' />
                <br />
                orgao_rg:
                <asp:TextBox ID="orgao_rgTextBox" runat="server" 
                    Text='<%# Bind("orgao_rg") %>' />
                <br />
                ccm:
                <asp:TextBox ID="ccmTextBox" runat="server" Text='<%# Bind("ccm") %>' />
                <br />
                telefone:
                <asp:TextBox ID="telefoneTextBox" runat="server" 
                    Text='<%# Bind("telefone") %>' />
                <br />
                celular:
                <asp:TextBox ID="celularTextBox" runat="server" Text='<%# Bind("celular") %>' />
                <br />
                cbo:
                <asp:TextBox ID="cboTextBox" runat="server" Text='<%# Bind("cbo") %>' />
                <br />
                rua:
                <asp:TextBox ID="ruaTextBox" runat="server" Text='<%# Bind("rua") %>' />
                <br />
                numero:
                <asp:TextBox ID="numeroTextBox" runat="server" Text='<%# Bind("numero") %>' />
                <br />
                complemento:
                <asp:TextBox ID="complementoTextBox" runat="server" 
                    Text='<%# Bind("complemento") %>' />
                <br />
                bairro:
                <asp:TextBox ID="bairroTextBox" runat="server" Text='<%# Bind("bairro") %>' />
                <br />
                cidade:
                <asp:TextBox ID="cidadeTextBox" runat="server" Text='<%# Bind("cidade") %>' />
                <br />
                estado:
                <asp:TextBox ID="estadoTextBox" runat="server" Text='<%# Bind("estado") %>' />
                <br />
                cep:
                <asp:TextBox ID="cepTextBox" runat="server" Text='<%# Bind("cep") %>' />
                <br />
                bancocod:
                <asp:TextBox ID="bancocodTextBox" runat="server" 
                    Text='<%# Bind("bancocod") %>' />
                <br />
                banconome:
                <asp:TextBox ID="banconomeTextBox" runat="server" 
                    Text='<%# Bind("banconome") %>' />
                <br />
                agencia:
                <asp:TextBox ID="agenciaTextBox" runat="server" Text='<%# Bind("agencia") %>' />
                <br />
                conta:
                <asp:TextBox ID="contaTextBox" runat="server" Text='<%# Bind("conta") %>' />
                <br />
                <asp:LinkButton ID="InsertButton" runat="server" CausesValidation="True" 
                    CommandName="Insert" Text="Insert" />
                &nbsp;<asp:LinkButton ID="InsertCancelButton" runat="server" 
                    CausesValidation="False" CommandName="Cancel" Text="Cancel" />
            </InsertItemTemplate>
            <ItemTemplate>
                &nbsp;<asp:LinkButton ID="NewButton" runat="server" CausesValidation="False" 
                    CommandName="New" Text="New" />
            </ItemTemplate>
        </asp:FormView>
    
        <asp:GridView ID="GridView1" runat="server" 
            AllowSorting="True" AutoGenerateColumns="False" CellPadding="4" 
            DataKeyNames="id_autonomo" DataSourceID="SqlDataSource1" 
            EmptyDataText="There are no data records to display." ForeColor="#333333" 
            GridLines="None" EnableModelValidation="True">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ButtonType="Image" EditImageUrl="~/images/page_edit.gif" EditText="Alterar" 
                    InsertText="Inserir" NewText="Novo" ShowEditButton="True" 
                    UpdateText="Salvar" CancelImageUrl="~/images/action_stop.gif" CancelText="Cancelar" 
                    DeleteText="Apagar" UpdateImageUrl="~/images/icon_accept.gif" />
                <asp:BoundField DataField="id_autonomo" HeaderText="#" ReadOnly="True" 
                    SortExpression="id_autonomo" />
                <asp:BoundField DataField="nome_autonomo" HeaderText="Nome Prestador" 
                    SortExpression="nome_autonomo" />
                <asp:BoundField DataField="dtnascimento" HeaderText="Nascimento" 
                    SortExpression="dtnascimento" />
                <asp:BoundField DataField="sexo" HeaderText="sexo" SortExpression="sexo" />
                <asp:BoundField DataField="tipo_prestacao" HeaderText="tipo_prestacao" 
                    SortExpression="tipo_prestacao" />
                <asp:BoundField DataField="nacionalidade" HeaderText="nacionalidade" 
                    SortExpression="nacionalidade" />
                <asp:BoundField DataField="estado_civil" HeaderText="estado_civil" 
                    SortExpression="estado_civil" />
                <asp:BoundField DataField="cpf" HeaderText="cpf" SortExpression="cpf" />
                <asp:BoundField DataField="nit" HeaderText="nit" SortExpression="nit" />
                <asp:BoundField DataField="rg" HeaderText="rg" SortExpression="rg" />
                <asp:BoundField DataField="orgao_rg" HeaderText="orgao_rg" 
                    SortExpression="orgao_rg" />
                <asp:BoundField DataField="ccm" HeaderText="ccm" SortExpression="ccm" />
                <asp:BoundField DataField="telefone" HeaderText="telefone" 
                    SortExpression="telefone" />
                <asp:BoundField DataField="celular" HeaderText="celular" 
                    SortExpression="celular" />
                <asp:BoundField DataField="cbo" HeaderText="cbo" SortExpression="cbo" />
                <asp:BoundField DataField="rua" HeaderText="rua" SortExpression="rua" />
                <asp:BoundField DataField="numero" HeaderText="numero" 
                    SortExpression="numero" />
                <asp:BoundField DataField="complemento" HeaderText="complemento" 
                    SortExpression="complemento" />
                <asp:BoundField DataField="bairro" HeaderText="bairro" 
                    SortExpression="bairro" />
                <asp:BoundField DataField="cidade" HeaderText="cidade" 
                    SortExpression="cidade" />
                <asp:BoundField DataField="estado" HeaderText="estado" 
                    SortExpression="estado" />
                <asp:BoundField DataField="cep" HeaderText="cep" SortExpression="cep" />
                <asp:BoundField DataField="bancocod" HeaderText="bancocod" 
                    SortExpression="bancocod" />
                <asp:BoundField DataField="banconome" HeaderText="banconome" 
                    SortExpression="banconome" />
                <asp:BoundField DataField="agencia" HeaderText="agencia" 
                    SortExpression="agencia" />
                <asp:BoundField DataField="conta" HeaderText="conta" SortExpression="conta" />
            </Columns>
            <EditRowStyle BackColor="#7C6F57" />
            <FooterStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#1C5E55" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#666666" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#E3EAEB" />
            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
            ConnectionString="<%$ ConnectionStrings:intranet_rhConnectionString1 %>" 
            DeleteCommand="DELETE FROM [autonomo] WHERE [id_autonomo] = @id_autonomo" 
            InsertCommand="INSERT INTO [autonomo] ([nome_autonomo], [dtnascimento], [sexo], [tipo_prestacao], [nacionalidade], [estado_civil], [cpf], [nit], [rg], [orgao_rg], [ccm], [telefone], [celular], [cbo], [rua], [numero], [complemento], [bairro], [cidade], [estado], [cep], [bancocod], [banconome], [agencia], [conta]) VALUES (@nome_autonomo, @dtnascimento, @sexo, @tipo_prestacao, @nacionalidade, @estado_civil, @cpf, @nit, @rg, @orgao_rg, @ccm, @telefone, @celular, @cbo, @rua, @numero, @complemento, @bairro, @cidade, @estado, @cep, @bancocod, @banconome, @agencia, @conta)" 
            ProviderName="<%$ ConnectionStrings:intranet_rhConnectionString1.ProviderName %>" 
            SelectCommand="select top 10 *
from autonomo" 
            
            UpdateCommand="UPDATE [autonomo] SET [nome_autonomo] = @nome_autonomo, [dtnascimento] = @dtnascimento, [sexo] = @sexo, [tipo_prestacao] = @tipo_prestacao, [nacionalidade] = @nacionalidade, [estado_civil] = @estado_civil, [cpf] = @cpf, [nit] = @nit, [rg] = @rg, [orgao_rg] = @orgao_rg, [ccm] = @ccm, [telefone] = @telefone, [celular] = @celular, [cbo] = @cbo, [rua] = @rua, [numero] = @numero, [complemento] = @complemento, [bairro] = @bairro, [cidade] = @cidade, [estado] = @estado, [cep] = @cep, [bancocod] = @bancocod, [banconome] = @banconome, [agencia] = @agencia, [conta] = @conta WHERE [id_autonomo] = @id_autonomo">
            <DeleteParameters>
                <asp:Parameter Name="id_autonomo" Type="Int32" />
            </DeleteParameters>
            <InsertParameters>
                <asp:Parameter Name="nome_autonomo" Type="String" />
                <asp:Parameter Name="dtnascimento" Type="DateTime" />
                <asp:Parameter Name="sexo" Type="String" />
                <asp:Parameter Name="tipo_prestacao" Type="String" />
                <asp:Parameter Name="nacionalidade" Type="String" />
                <asp:Parameter Name="estado_civil" Type="String" />
                <asp:Parameter Name="cpf" Type="String" />
                <asp:Parameter Name="nit" Type="String" />
                <asp:Parameter Name="rg" Type="String" />
                <asp:Parameter Name="orgao_rg" Type="String" />
                <asp:Parameter Name="ccm" Type="String" />
                <asp:Parameter Name="telefone" Type="String" />
                <asp:Parameter Name="celular" Type="String" />
                <asp:Parameter Name="cbo" Type="String" />
                <asp:Parameter Name="rua" Type="String" />
                <asp:Parameter Name="numero" Type="String" />
                <asp:Parameter Name="complemento" Type="String" />
                <asp:Parameter Name="bairro" Type="String" />
                <asp:Parameter Name="cidade" Type="String" />
                <asp:Parameter Name="estado" Type="String" />
                <asp:Parameter Name="cep" Type="String" />
                <asp:Parameter Name="bancocod" Type="String" />
                <asp:Parameter Name="banconome" Type="String" />
                <asp:Parameter Name="agencia" Type="String" />
                <asp:Parameter Name="conta" Type="String" />
            </InsertParameters>
            <UpdateParameters>
                <asp:Parameter Name="nome_autonomo" Type="String" />
                <asp:Parameter Name="dtnascimento" Type="DateTime" />
                <asp:Parameter Name="sexo" Type="String" />
                <asp:Parameter Name="tipo_prestacao" Type="String" />
                <asp:Parameter Name="nacionalidade" Type="String" />
                <asp:Parameter Name="estado_civil" Type="String" />
                <asp:Parameter Name="cpf" Type="String" />
                <asp:Parameter Name="nit" Type="String" />
                <asp:Parameter Name="rg" Type="String" />
                <asp:Parameter Name="orgao_rg" Type="String" />
                <asp:Parameter Name="ccm" Type="String" />
                <asp:Parameter Name="telefone" Type="String" />
                <asp:Parameter Name="celular" Type="String" />
                <asp:Parameter Name="cbo" Type="String" />
                <asp:Parameter Name="rua" Type="String" />
                <asp:Parameter Name="numero" Type="String" />
                <asp:Parameter Name="complemento" Type="String" />
                <asp:Parameter Name="bairro" Type="String" />
                <asp:Parameter Name="cidade" Type="String" />
                <asp:Parameter Name="estado" Type="String" />
                <asp:Parameter Name="cep" Type="String" />
                <asp:Parameter Name="bancocod" Type="String" />
                <asp:Parameter Name="banconome" Type="String" />
                <asp:Parameter Name="agencia" Type="String" />
                <asp:Parameter Name="conta" Type="String" />
                <asp:Parameter Name="id_autonomo" Type="Int32" />
            </UpdateParameters>
        </asp:SqlDataSource>
    
    </div>
    </form>
</body>
</html>
