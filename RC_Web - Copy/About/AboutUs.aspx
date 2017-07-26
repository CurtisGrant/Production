<%@ Page Title="About Us" Language="VB" MasterPageFile="~/MasterPages/Frontend.master" AutoEventWireup="false" 
    CodeFile="AboutUs.aspx.vb" Inherits="About_AboutUs" %>

<%@ Register src="~/Controls/Banner.ascx" tagname="Banner" tagprefix="uc1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cpMainContent" Runat="Server">
    <p>
        </p>
    <uc1:Banner ID="Banner1" runat="server" />
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        Page for About Us</p>
</asp:Content>

