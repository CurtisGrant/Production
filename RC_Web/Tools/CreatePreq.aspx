<%@ Page Title="CreatePreq" Language="VB" MasterPageFile="~/MasterPages/Frontend.master" AutoEventWireup="false" 
    CodeFile="CreatePreq.aspx.vb" Inherits="_CreatePreq" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <link href="~/App_Themes/Monochrome/Monochrome.css" rel="stylesheet" type="text/css" />
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cpMainContent" Runat="Server">
    <p>
        &nbsp;</p>
    <asp:Panel ID="Panel1" runat="server" style="z-index: 1; left: 6px; top: -38px; position: relative; height: 34px; width: 840px">
        <asp:FileUpload ID="FileUpload1" runat="server" style="z-index: 1; left: 316px; top: 3px; position: absolute; width: 299px; margin-top: 0px" />
        <asp:DropDownList ID="ddPreq" runat="server" Font-Size="Small" Height="22px" style="z-index: 1; left: 6px; top: 3px; position: absolute; width: 300px">
        </asp:DropDownList>
        <asp:Button ID="MapFields" runat="server" Text="Map Fields" style="z-index: 1; left: 600px; top: 100px; position: absolute; 
                height: 30px" />
        <asp:Button ID="ResetColumnHeadings" runat="server" style="z-index: 1; left: 600px; top: 150px; position: absolute; 
                height: 30px" Text="Reset Column Headings" Height="20px" Visible="False" Width="200px" />        
        <asp:Button ID="CheckSkus" runat="server" style="z-index: 1; left: 600px; top: 250px; position: absolute;
                height: 30px" Text="Check SKUs" Visible="False" />
        <asp:Button ID="CreateSpreadsheet" runat="server" style="z-index: 1; left: 600px; top: 300px; position: absolute; 
                width: 198px; height: 30px;" Text="Create Spreadsheet(s)" Visible="False" />
          <asp:Label ID="mapLabel" runat="server" style="z-index: 1; left: 600px; top: 350px; position: absolute;
                height: 25px" Text="Not all required fields have been mapped" Visible="False" />  
        <br />
        <br />
        
    </asp:Panel>
    <div>
    <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
        <ContentTemplate>
            <asp:ListBox ID="RequiredFields" runat="server" Height="280px" Width="120px" Visible="False" style="z-index: 1; left: 42px; top: 0px; position: relative" >
                <asp:ListItem>*Item</asp:ListItem>
                <asp:ListItem>*Description</asp:ListItem>
                <asp:ListItem>Vendor ID</asp:ListItem>
                <asp:ListItem>Whse</asp:ListItem>
                <asp:ListItem>Cost</asp:ListItem>
                <asp:ListItem>Retail</asp:ListItem>
                <asp:ListItem>Category</asp:ListItem>
                <asp:ListItem>Sub Category</asp:ListItem>
                <asp:ListItem>Stock Unit</asp:ListItem>
                <asp:ListItem>Ecomm</asp:ListItem>
                <asp:ListItem>Order Qty</asp:ListItem>
                <asp:ListItem>Vendor Item</asp:ListItem>
            </asp:ListBox>
            <asp:ListBox ID="YourFields" runat="server" AutoPostBack="True" Height="280px" Width="120px" Visible="False" style="margin-top: 7px; z-index: 1; left: 135px; top: -1px; position: relative;"></asp:ListBox>
            
            <br />
        </ContentTemplate>
    </asp:UpdatePanel>
    </div>
    <p></p>
        <div>
            <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                <ContentTemplate>
                    <asp:GridView ID="GridView1" runat="server" CssClass="dynamic" Font-Size="Small" Width="840px">
            </asp:GridView>
                </ContentTemplate>            
            </asp:UpdatePanel>
             
        </div>
    <script type="text/javascript">
    function UploadFile(fileUpload) {
        if (fileUpload.value != '') {
            document.getElementById("<%=MapFields.ClientID%>").click();
        }
    }
</script>   
    
</asp:Content>

