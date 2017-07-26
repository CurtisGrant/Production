<%@ Page Title="CreateNewPreq" Language="VB" MasterPageFile="~/MasterPages/Frontend.master"
    AutoEventWireup="false" CodeFile="CreateNewPreq.aspx.vb" Inherits="Tools_CreateNewPreq" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cpMainContent" runat="Server">
    <div id="Div1" style="width: 100%; overflow: auto;">
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>               
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                    ConnectionString="<%$ ConnectionStrings:TCMConnectionString %>" 
                    SelectCommand="SELECT [PREQ_NO], [Batch_Id], [Loc_Id], [Vendor_Id], [Buyer], [Order_Date], [Deliver_Date], [Cancel_Date] FROM [Purchase_Request_Header] WHERE ([PREQ_NO] = @PREQ_NO)">
                    <SelectParameters>
                        <asp:ControlParameter ControlID="ddlPreq" Name="PREQ_NO" PropertyName="SelectedValue" Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>
              
                
            </ContentTemplate>
        </asp:UpdatePanel>
        <asp:DropDownList ID="ddlPreq" runat="server" Visible="False">
        </asp:DropDownList>
        <br />
        <br />
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True" 
            AutoGenerateColumns="False" DataKeyNames="PREQ_NO,Loc_Id" DataSourceID="SqlDataSource1" 
            Font-Size="Small">
            <Columns>
                <asp:BoundField DataField="PREQ_NO" HeaderText="PREQ_NO" ReadOnly="True" SortExpression="PREQ_NO" />
                <asp:BoundField DataField="Batch_Id" HeaderText="Batch_Id" SortExpression="Batch_Id" />
                <asp:BoundField DataField="Loc_Id" HeaderText="Loc_Id" ReadOnly="True" SortExpression="Loc_Id" />
                <asp:BoundField DataField="Vendor_Id" HeaderText="Vendor_Id" SortExpression="Vendor_Id" />
                <asp:BoundField DataField="Buyer" HeaderText="Buyer" SortExpression="Buyer" />
                <asp:BoundField DataField="Order_Date" HeaderText="Order_Date" SortExpression="Order_Date" />
                <asp:BoundField DataField="Deliver_Date" HeaderText="Deliver_Date" SortExpression="Deliver_Date" />
                <asp:BoundField DataField="Cancel_Date" HeaderText="Cancel_Date" SortExpression="Cancel_Date" />
            </Columns>
        </asp:GridView>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" DataKeyNames="PREQ_ID,Loc_Id" 
            DataSourceID="SqlDataSource3" Font-Size="Small">
            <Columns>
                <asp:CommandField ShowSelectButton="True" />
                <asp:BoundField DataField="PREQ_ID" HeaderText="PREQ_ID" ReadOnly="True" SortExpression="PREQ_ID" />
                <asp:BoundField DataField="Batch_Id" HeaderText="Batch_Id" SortExpression="Batch_Id" />
                <asp:BoundField DataField="Loc_Id" HeaderText="Loc_Id" ReadOnly="True" SortExpression="Loc_Id" />
                <asp:BoundField DataField="Item" HeaderText="Item" SortExpression="Item" />
                <asp:BoundField DataField="Description" HeaderText="Description" SortExpression="Description" />
                <asp:BoundField DataField="Vendor_Item" HeaderText="Vendor_Item" SortExpression="Vendor_Item" />
                <asp:BoundField DataField="Cost" DataFormatString="{0:c}" HeaderText="Cost" SortExpression="Cost" />
                <asp:BoundField DataField="Retail" DataFormatString="{0:c}" HeaderText="Retail" SortExpression="Retail" />
                <asp:BoundField DataField="Category" HeaderText="Category" SortExpression="Category" />
                <asp:BoundField DataField="Sub_Category" HeaderText="Sub_Category" SortExpression="Sub_Category" />
                <asp:BoundField DataField="Stock_Units" HeaderText="Stock_Units" SortExpression="Stock_Units" />
            </Columns>
        </asp:GridView>
        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" 
            DataKeyNames="PREQ_ID,Loc_Id,SKU,Item" DataSourceID="SqlDataSource4" Font-Size="Small">
            <Columns>
                <asp:BoundField DataField="PREQ_ID" HeaderText="PREQ_ID" ReadOnly="True" SortExpression="PREQ_ID" />
                <asp:BoundField DataField="Loc_Id" HeaderText="Loc_Id" ReadOnly="True" SortExpression="Loc_Id" />
                <asp:BoundField DataField="SKU" HeaderText="SKU" ReadOnly="True" SortExpression="SKU" />
                <asp:BoundField DataField="UPC" HeaderText="UPC" SortExpression="UPC" />
                <asp:BoundField DataField="Cost" HeaderText="Cost" SortExpression="Cost" />
                <asp:BoundField DataField="Retail" HeaderText="Retail" SortExpression="Retail" />
                <asp:BoundField DataField="Qty" HeaderText="Qty" SortExpression="Qty" />
                <asp:BoundField DataField="Item" HeaderText="Item" ReadOnly="True" SortExpression="Item" />
                <asp:BoundField DataField="DIM1" HeaderText="DIM1" SortExpression="DIM1" />
                <asp:BoundField DataField="DIM2" HeaderText="DIM2" SortExpression="DIM2" />
                <asp:BoundField DataField="DIM3" HeaderText="DIM3" SortExpression="DIM3" />
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource4" runat="server" ConnectionString="<%$ ConnectionStrings:TCMConnectionString %>" 
            SelectCommand="SELECT * FROM [PREQ_Sku] WHERE ([PREQ_ID] = @PREQ_ID)">
            <SelectParameters>
                <asp:ControlParameter ControlID="GridView1" Name="PREQ_ID" PropertyName="SelectedValue" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        <br />

        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ 
            ConnectionStrings:TCMConnectionString %>" 
            SelectCommand="SELECT * FROM [PREQ_Detail] WHERE ([PREQ_ID] = @PREQ_ID)">
            <SelectParameters>
                <asp:ControlParameter ControlID="GridView1" Name="PREQ_ID" PropertyName="SelectedValue" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        <br />
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ 
            ConnectionStrings:TCMConnectionString %>" 
            SelectCommand="SELECT [PREQ_ID], [Batch_ID], [Loc_Id], [Vendor_ID], [Buyer], [Order_Date], 
            [Deliver_Date], [Cancel_Date] FROM [PREQ_Header] WHERE ([PREQ_ID] = @PREQ_ID)">
            <SelectParameters>
                <asp:ControlParameter ControlID="GridView1" Name="PREQ_ID" PropertyName="SelectedValue" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>





    </div>
<%--        <script type="text/javascript">
        function UploadFile(fileUpload) {
            if (fileUpload.value != '') {
                document.getElementById("<%=btnCreatePREQ.ClientID%>").click;
     }
 }
</script--%>>
</asp:Content>


