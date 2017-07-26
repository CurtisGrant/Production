<%@ Page Title="Create Items" Language="VB" MasterPageFile="~/MasterPages/Frontend.master" 
   AutoEventWireup="false" CodeFile="CreateItems.aspx.vb" Inherits="_CreateItems" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <link href="App_Themes/Monochrome/Monochrome.css" rel="stylesheet" />
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cpMainContent" Runat="Server">
    <asp:Panel ID="Panel1" runat="server">
        <asp:FileUpload ID="FileUpload1" runat="server" Width="300px" />
        <asp:Button ID="LoadXL" onclick="LoadXL_Click" runat="server" Text="Load"/>       
        <asp:Button ID="CheckStatus" runat="server" Text="Check Status" Visible="False" />        
    </asp:Panel>    
    <br />
    <div>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
                <asp:GridView ID="gv1" runat="server" Font-Size="X-Small" Width="840px">
                    <AlternatingRowStyle BackColor="#CCCCCC" />
                </asp:GridView>
            </ContentTemplate>
        </asp:UpdatePanel>
    </div>
<script type="text/javascript">
   function UploadFile(fileUpload) {
     if (fileUpload.value != '') {
       document.getElementById("<%=LoadXL.ClientID%>").click();
        }
    }
</script>

</asp:Content>

