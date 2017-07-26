<%@ Page Title="LVExercise" Language="VB" MasterPageFile="~/MasterPages/Frontend.master" 
   AutoEventWireup="false" CodeFile="LVExercise.aspx.vb" Inherits="Demos_LVExercise" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cpMainContent" Runat="Server">
    <p>
        <asp:TextBox ID="txtPreq" runat="server"></asp:TextBox>
        <asp:Label ID="Label1" runat="server" Text="PREQ"></asp:Label>
        <asp:TextBox ID="txtLoc" runat="server"></asp:TextBox>
        <asp:Label ID="Label10" runat="server" Text="Location"></asp:Label>
        <asp:Button ID="btnGO" runat="server" Text="GO" />
        <br />
        <br />

    </p>
    <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" 
            DataKeyNames="PREQ_ID,Loc_Id" DataSourceID="SqlDataSource2"  ShowFooter="True" HeaderStyle-ForeColor="Black" ShowHeaderWhenEmpty="True" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" CellPadding="1" Font-Size="Smaller" ForeColor="Black" GridLines="Vertical" HorizontalAlign="Justify" Width="1044px">
        <AlternatingRowStyle BackColor="#CCCCCC" Width="1044px" />
        <Columns>
            <asp:CommandField ShowDeleteButton="True" ShowEditButton="True" ShowInsertButton="true"/>
            <asp:BoundField DataField="PREQ_ID" HeaderText="PREQ_ID" ReadOnly="True" SortExpression="PREQ_ID" />
            <asp:BoundField DataField="Loc_Id" HeaderText="Loc_Id" ReadOnly="True" SortExpression="Loc_Id" />
            <asp:TemplateField HeaderText="Item" SortExpression="Item">
                <EditItemTemplate>
                    <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("Item") %>'></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="newItem" runat="server"></asp:TextBox>
                </FooterTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label1" runat="server" Text='<%# Bind("Item") %>'></asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Description" SortExpression="Description">
                <EditItemTemplate>
                    <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("Description") %>'></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="newDescription" runat="server" ></asp:TextBox>
                </FooterTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label2" runat="server" Text='<%# Bind("Description") %>'></asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Vendor_Item" SortExpression="Vendor_Item">
                <EditItemTemplate>
                    <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("Vendor_Item") %>'></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="newVendorItem" runat="server"></asp:TextBox>
                </FooterTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label3" runat="server" Text='<%# Bind("Vendor_Item") %>'></asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Cost" SortExpression="Cost">
                <EditItemTemplate>
                    <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("Cost") %>'></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="newCost" runat="server"></asp:TextBox>
                </FooterTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label4" runat="server" Text='<%# Bind("Cost") %>'></asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Retail" SortExpression="Retail">
                <EditItemTemplate>
                    <asp:TextBox ID="TextBox5" runat="server" Text='<%# Bind("Retail") %>'></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="newRetail" runat="server"></asp:TextBox>
                </FooterTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label5" runat="server" Text='<%# Bind("Retail") %>'></asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Category" SortExpression="Category">
                <EditItemTemplate>
                    <asp:TextBox ID="TextBox6" runat="server" Text='<%# Bind("Category") %>'></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="newCategory" runat="server"></asp:TextBox>
                </FooterTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label6" runat="server" Text='<%# Bind("Category") %>'></asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Sub_Category" SortExpression="Sub_Category">
                <EditItemTemplate>
                    <asp:TextBox ID="TextBox7" runat="server" Text='<%# Bind("Sub_Category") %>'></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="newSubCategory" runat="server"></asp:TextBox>
                </FooterTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label7" runat="server" Text='<%# Bind("Sub_Category") %>'></asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Stock_Units" SortExpression="Stock_Units">
                <EditItemTemplate>
                    <asp:TextBox ID="TextBox8" runat="server" Text='<%# Bind("Stock_Units") %>'></asp:TextBox>
                </EditItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="newStockUnits" runat="server">EACH</asp:TextBox>
                    <asp:Button ID="AddItem" runat="server" Text="Insert" />
                </FooterTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label8" runat="server" Text='<%# Bind("Stock_Units") %>'></asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
        </Columns>

        <EditRowStyle Font-Size="Smaller" />
        <EmptyDataRowStyle Font-Size="Smaller" />
        <FooterStyle BackColor="#CCCCCC" Font-Size="Smaller" Width="1044px" />
        <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" />
        <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
        <SortedAscendingCellStyle BackColor="#F1F1F1" />
        <SortedAscendingHeaderStyle BackColor="#808080" />
        <SortedDescendingCellStyle BackColor="#CAC9C9" />
        <SortedDescendingHeaderStyle BackColor="#383838" />
    </asp:GridView>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
        ConnectionString="<%$ ConnectionStrings:TCMConnectionString %>" 
        DeleteCommand="DELETE FROM [PREQ_Detail] WHERE [PREQ_ID] = @PREQ_ID AND [Loc_Id] = @Loc_Id" 
        InsertCommand="INSERT INTO [PREQ_Detail] ([PREQ_ID], [Loc_Id], [Item], [Description], [Vendor_Item], 
        [Cost], [Retail], [Category], [Sub_Category], [Stock_Units]) 
        VALUES (@PREQ_ID, @Loc_Id, @Item, @Description, @Vendor_Item, @Cost, @Retail, @Category, 
        @Sub_Category, @Stock_Units)" 
        SelectCommand="SELECT [PREQ_ID], [Loc_Id], [Item], [Description], [Vendor_Item], [Cost], [Retail], 
        [Category], [Sub_Category], [Stock_Units] FROM [PREQ_Detail] WHERE PREQ_ID = @preq AND Loc_Id = @loc" 
        UpdateCommand="UPDATE [PREQ_Detail] SET [Item] = @Item, [Description] = @Description, 
        [Vendor_Item] = @Vendor_Item, [Cost] = @Cost, [Retail] = @Retail, [Category] = @Category, 
        [Sub_Category] = @Sub_Category, [Stock_Units] = @Stock_Units WHERE [PREQ_ID] = @PREQ_ID 
        AND [Loc_Id] = @Loc_Id">
        <SelectParameters>
            <asp:ControlParameter ControlID="txtPREQ" Type="String" />
            <asp:ControlParameter ControlID="txtLoc" Type="String" />
        </SelectParameters>
        <DeleteParameters>
            <asp:Parameter Name="PREQ_ID" Type="String" />
            <asp:Parameter Name="Loc_Id" Type="String" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="PREQ_ID" Type="String" />
            <asp:Parameter Name="Loc_Id" Type="String" />
            <asp:Parameter Name="Item" Type="String" />
            <asp:Parameter Name="Description" Type="String" />
            <asp:Parameter Name="Vendor_Item" Type="String" />
            <asp:Parameter Name="Cost" Type="Decimal" />
            <asp:Parameter Name="Retail" Type="Decimal" />
            <asp:Parameter Name="Category" Type="String" />
            <asp:Parameter Name="Sub_Category" Type="String" />
            <asp:Parameter Name="Stock_Units" Type="String" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="Item" Type="String" />
            <asp:Parameter Name="Description" Type="String" />
            <asp:Parameter Name="Vendor_Item" Type="String" />
            <asp:Parameter Name="Cost" Type="Decimal" />
            <asp:Parameter Name="Retail" Type="Decimal" />
            <asp:Parameter Name="Category" Type="String" />
            <asp:Parameter Name="Sub_Category" Type="String" />
            <asp:Parameter Name="Stock_Units" Type="String" />
            <asp:Parameter Name="PREQ_ID" Type="String" />
            <asp:Parameter Name="Loc_Id" Type="String" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <p>

        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:TCMConnectionString %>" DeleteCommand="DELETE FROM [PREQ_Detail] WHERE [Loc_Id] = @Loc_Id AND [PREQ_ID] = @PREQ_ID" InsertCommand="INSERT INTO [PREQ_Detail] ([Loc_Id], [Item], [Description], [Vendor_Item], [Cost], [Retail], [Category], [Sub_Category], [Stock_Units], [PREQ_ID]) VALUES (@Loc_Id, @Item, @Description, @Vendor_Item, @Cost, @Retail, @Category, @Sub_Category, @Stock_Units, @PREQ_ID)" SelectCommand="SELECT [Loc_Id], [Item], [Description], [Vendor_Item], [Cost], [Retail], [Category], [Sub_Category], [Stock_Units], [PREQ_ID] FROM [PREQ_Detail]" UpdateCommand="UPDATE [PREQ_Detail] SET [Item] = @Item, [Description] = @Description, [Vendor_Item] = @Vendor_Item, [Cost] = @Cost, [Retail] = @Retail, [Category] = @Category, [Sub_Category] = @Sub_Category, [Stock_Units] = @Stock_Units WHERE [Loc_Id] = @Loc_Id AND [PREQ_ID] = @PREQ_ID">
            <DeleteParameters>
                <asp:Parameter Name="Loc_Id" Type="String" />
                <asp:Parameter Name="PREQ_ID" Type="String" />
            </DeleteParameters>
            <InsertParameters>
                <asp:Parameter Name="Loc_Id" Type="String" />
                <asp:Parameter Name="Item" Type="String" />
                <asp:Parameter Name="Description" Type="String" />
                <asp:Parameter Name="Vendor_Item" Type="String" />
                <asp:Parameter Name="Cost" Type="Decimal" />
                <asp:Parameter Name="Retail" Type="Decimal" />
                <asp:Parameter Name="Category" Type="String" />
                <asp:Parameter Name="Sub_Category" Type="String" />
                <asp:Parameter Name="Stock_Units" Type="String" />
                <asp:Parameter Name="PREQ_ID" Type="String" />
            </InsertParameters>
            <UpdateParameters>
                <asp:Parameter Name="Item" Type="String" />
                <asp:Parameter Name="Description" Type="String" />
                <asp:Parameter Name="Vendor_Item" Type="String" />
                <asp:Parameter Name="Cost" Type="Decimal" />
                <asp:Parameter Name="Retail" Type="Decimal" />
                <asp:Parameter Name="Category" Type="String" />
                <asp:Parameter Name="Sub_Category" Type="String" />
                <asp:Parameter Name="Stock_Units" Type="String" />
                <asp:Parameter Name="Loc_Id" Type="String" />
                <asp:Parameter Name="PREQ_ID" Type="String" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </p>
</asp:Content>

