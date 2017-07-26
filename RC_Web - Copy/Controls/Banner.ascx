<%@ Control Language="VB" AutoEventWireup="false" CodeFile="Banner.ascx.vb" Inherits="Controls_Banner" %>
<style type="text/css">
    .auto-style1 {
        width: 3264px;
        height: 1836px;
    }
</style>

<asp:Panel ID="VerticalPanel" runat="server">
    <a href="http:google.com" target="_blank">
        <asp:Image ID="Image1" runat="server" AlternateText="This is a sample banner" 
            ImageUrl="~Images\Banner120x240.gif" />            
    </a>
</asp:Panel>

    

