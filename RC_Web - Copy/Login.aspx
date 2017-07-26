<%@ Page Title="Log in to Retail Clarity" Language="VB" MasterPageFile="~/MasterPages/Frontend.master" AutoEventWireup="false" 
    CodeFile="Login.aspx.vb" Inherits="Login" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cpMainContent" Runat="Server">
    <h1>Login here</h1>
<p>
    <asp:LoginView ID="LoginView1" runat="server">
        <AnonymousTemplate>
            <asp:Login ID="Login1" runat="server" CreateUserUrl="Signup.aspx" DestinationPageUrl="~/Default.aspx"
                CreateUserText="Create an account in Retail Clarity"> 
            </asp:Login>
        </AnonymousTemplate>
        <LoggedInTemplate>
            You are already logged in
        </LoggedInTemplate>
    </asp:LoginView>
</p>
    <asp:LoginStatus ID="LoginStatus1" runat="server" />
</asp:Content>

