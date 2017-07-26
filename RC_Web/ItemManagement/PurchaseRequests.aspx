<%@ Page Title="Purchase Requests" Language="VB" MasterPageFile="~/MasterPages/ItemManagement.master" AutoEventWireup="false" CodeFile="PurchaseRequests.aspx.vb" Inherits="ItemManagement_PurchaseRequests" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cpMainContent" Runat="Server">
    <asp:DropDownList ID="DropDownList1" runat="server" AppendDataBoundItems="True" AutoPostBack="True" DataSourceID="SqlDataSource1" DataTextField="PREQ_ID" DataValueField="PREQ_ID">
        <asp:ListItem Value="">Choose a PREQ</asp:ListItem>
    </asp:DropDownList>
    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataKeyNames="PREQ_ID,Loc_Id" 
        DataSourceID="SqlDataSource2">
        <Columns>
            <asp:HyperLinkField HeaderText="Purchase Request" />
            <asp:BoundField DataField="PREQ_ID" HeaderText="PREQ_ID" ReadOnly="True" SortExpression="PREQ_ID" />
            <asp:BoundField DataField="Batch_ID" HeaderText="Batch_ID" SortExpression="Batch_ID" />
            <asp:BoundField DataField="Loc_Id" HeaderText="Loc_Id" ReadOnly="True" SortExpression="Loc_Id" />
            <asp:BoundField DataField="Vendor_ID" HeaderText="Vendor_ID" SortExpression="Vendor_ID" />
            <asp:BoundField DataField="Buyer" HeaderText="Buyer" SortExpression="Buyer" />
            <asp:BoundField DataField="Order_Date" HeaderText="Order_Date" SortExpression="Order_Date" DataFormatString="{0:d}" />
            <asp:BoundField DataField="Deliver_Date" HeaderText="Deliver_Date" SortExpression="Deliver_Date" DataFormatString="{0:d}" />
            <asp:BoundField DataField="Cancel_Date" HeaderText="Cancel_Date" SortExpression="Cancel_Date" DataFormatString="{0:d}" />
            <asp:CommandField HeaderText="Delete" ShowDeleteButton="True" />
            <asp:CommandField HeaderText="Edit" ShowEditButton="True" />
        </Columns>
    </asp:GridView>
    <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
        ConnectionString="<%$ ConnectionStrings:TCMConnectionString %>" 
        SelectCommand="SELECT [PREQ_ID], [Batch_ID], [Loc_Id], [Vendor_ID], [Buyer], [Order_Date], [Deliver_Date], [Cancel_Date] FROM [PREQ_Header] WHERE ([PREQ_ID] = @PREQ_ID)" DeleteCommand="DELETE FROM [PREQ_Header] WHERE [PREQ_ID] = @PREQ_ID AND [Loc_Id] = @Loc_Id" InsertCommand="INSERT INTO [PREQ_Header] ([PREQ_ID], [Batch_ID], [Loc_Id], [Vendor_ID], [Buyer], [Order_Date], [Deliver_Date], [Cancel_Date]) VALUES (@PREQ_ID, @Batch_ID, @Loc_Id, @Vendor_ID, @Buyer, @Order_Date, @Deliver_Date, @Cancel_Date)" UpdateCommand="UPDATE [PREQ_Header] SET [Batch_ID] = @Batch_ID, [Vendor_ID] = @Vendor_ID, [Buyer] = @Buyer, [Order_Date] = @Order_Date, [Deliver_Date] = @Deliver_Date, [Cancel_Date] = @Cancel_Date WHERE [PREQ_ID] = @PREQ_ID AND [Loc_Id] = @Loc_Id">
        <DeleteParameters>
            <asp:Parameter Name="PREQ_ID" Type="String" />
            <asp:Parameter Name="Loc_Id" Type="String" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="PREQ_ID" Type="String" />
            <asp:Parameter Name="Batch_ID" Type="String" />
            <asp:Parameter Name="Loc_Id" Type="String" />
            <asp:Parameter Name="Vendor_ID" Type="String" />
            <asp:Parameter Name="Buyer" Type="String" />
            <asp:Parameter DbType="Date" Name="Order_Date" />
            <asp:Parameter DbType="Date" Name="Deliver_Date" />
            <asp:Parameter DbType="Date" Name="Cancel_Date" />
        </InsertParameters>
        <SelectParameters>
            <asp:ControlParameter ControlID="DropDownList1" Name="PREQ_ID" PropertyName="SelectedValue" 
                Type="String" />
        </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="Batch_ID" Type="String" />
            <asp:Parameter Name="Vendor_ID" Type="String" />
            <asp:Parameter Name="Buyer" Type="String" />
            <asp:Parameter DbType="Date" Name="Order_Date" />
            <asp:Parameter DbType="Date" Name="Deliver_Date" />
            <asp:Parameter DbType="Date" Name="Cancel_Date" />
            <asp:Parameter Name="PREQ_ID" Type="String" />
            <asp:Parameter Name="Loc_Id" Type="String" />
        </UpdateParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ 
        ConnectionStrings:TCMConnectionString %>" 
        DeleteCommand="DELETE FROM [PREQ_Header] WHERE [PREQ_ID] = @PREQ_ID AND [Loc_Id] = @Loc_Id" 
        InsertCommand="INSERT INTO [PREQ_Header] ([PREQ_ID], [Batch_ID], [Loc_Id], [Vendor_ID], [Buyer], 
        [Order_Date], [Deliver_Date], [Cancel_Date], [Last_Update]) 
        VALUES (@PREQ_ID, @Batch_ID, @Loc_Id, @Vendor_ID, @Buyer, @Order_Date, @Deliver_Date, 
        @Cancel_Date, @Last_Update)" 
        SelectCommand="SELECT * FROM [PREQ_Header]" UpdateCommand="UPDATE [PREQ_Header] SET [Batch_ID] = @Batch_ID, 
        [Vendor_ID] = @Vendor_ID, [Buyer] = @Buyer, [Order_Date] = @Order_Date, [Deliver_Date] = @Deliver_Date, 
        [Cancel_Date] = @Cancel_Date, [Last_Update] = @Last_Update 
        WHERE [PREQ_ID] = @PREQ_ID AND [Loc_Id] = @Loc_Id">
        <DeleteParameters>
            <asp:Parameter Name="PREQ_ID" Type="String" />
            <asp:Parameter Name="Loc_Id" Type="String" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="PREQ_ID" Type="String" />
            <asp:Parameter Name="Batch_ID" Type="String" />
            <asp:Parameter Name="Loc_Id" Type="String" />
            <asp:Parameter Name="Vendor_ID" Type="String" />
            <asp:Parameter Name="Buyer" Type="String" />
            <asp:Parameter DbType="Date" Name="Order_Date" />
            <asp:Parameter DbType="Date" Name="Deliver_Date" />
            <asp:Parameter DbType="Date" Name="Cancel_Date" />
            <asp:Parameter DbType="Date" Name="Last_Update" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="Batch_ID" Type="String" />
            <asp:Parameter Name="Vendor_ID" Type="String" />
            <asp:Parameter Name="Buyer" Type="String" />
            <asp:Parameter DbType="Date" Name="Order_Date" />
            <asp:Parameter DbType="Date" Name="Deliver_Date" />
            <asp:Parameter DbType="Date" Name="Cancel_Date" />
            <asp:Parameter DbType="Date" Name="Last_Update" />
            <asp:Parameter Name="PREQ_ID" Type="String" />
            <asp:Parameter Name="Loc_Id" Type="String" />
        </UpdateParameters>
    </asp:SqlDataSource>
</asp:Content>

