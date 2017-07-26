<%@ Page Title="Create Reorders" Language="VB" MasterPageFile="~/MasterPages/Frontend.master" AutoEventWireup="false" 
    CodeFile="CreateReorders.aspx.vb" Inherits="Tools_CreateReorders" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cpMainContent" Runat="Server">
    <p>
    Create Reorders</p>
    <asp:FileUpload ID="FileUpload1" runat="server" />
    <asp:Label ID="Label1" runat="server"></asp:Label>
    <asp:Button ID="Button1" runat="server" Text="Load File" />
</asp:Content>

