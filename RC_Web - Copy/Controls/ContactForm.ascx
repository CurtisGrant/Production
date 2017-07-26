<%@ Control Language="VB" AutoEventWireup="false" CodeFile="ContactForm.ascx.vb" Inherits="Controls_ContactForm" %>
<style type="text/css">
    .auto-style1 {
        width: 100%;
    }
    .auto-style2 {
        height: 23px;
    }
    .auto-style3 {
        height: 64px;
    }
</style>
<script>
    function validatePhoneNumbers(source, args)
    {
        var phoneHome = document.getElementById('<%= PhoneHome.ClientID %>');
        var businessPhone = document.getElementById('<%= PhoneBusiness.ClientID %>');
        if (phoneHome.value !='' || phoneBusiness !='')
        {
            args.IsValid = true;
        }
        else {
            args.IsValid = false;
        }
    }
</script>
<table class="auto-style1">
    <tr>
        <td colspan="3">Some information here</td>
    </tr>
    <tr>
        <td>Name</td>
        <td>
            <asp:TextBox ID="Name" runat="server"></asp:TextBox>
        </td>
        <td>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="Name" 
                CssClass="ErrorMessage" ErrorMessage="Enter your name">ERROR!</asp:RequiredFieldValidator>
        </td>
    </tr>
    <tr>
        <td>E-Mail address</td>
        <td>
            <asp:TextBox ID="EmailAddress" runat="server"></asp:TextBox>
        </td>
        <td>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="EmailAddress" 
                ErrorMessage="Enter an e-mail address">ERROR!</asp:RequiredFieldValidator>
            <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ControlToValidate="EmailAddress" 
                ErrorMessage="Enter a valid e-mail address" 
                ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*">ERROR!</asp:RegularExpressionValidator>
        </td>
    </tr>
    <tr>
        <td>Repeat E-Mail address</td>
        <td>
            <asp:TextBox ID="ConfirmEmailAddress" runat="server"></asp:TextBox>
        </td>
        <td>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="ConfirmEmailAddress" 
                ErrorMessage="Retype the email address">ERROR!</asp:RequiredFieldValidator>
            <asp:CompareValidator ID="CompareValidator1" runat="server" ControlToCompare="EmailAddress" ControlToValidate="ConfirmEmailAddress" 
                ErrorMessage="The e-mail addresses do not match">ERROR!</asp:CompareValidator>
        </td>
    </tr>
    <tr>
        <td>Business Phone number</td>
        <td>
            <asp:TextBox ID="PhoneBusiness" runat="server"></asp:TextBox>
        </td>
        <td>
            <asp:CustomValidator ID="CustomValidator1" runat="server" ClientValidationFunction="validatePhoneNumbers" 
                CssClass="ErrorMessage" Display="Dynamic" ErrorMessage="Enter your home or business">ERROR!</asp:CustomValidator>
        </td>
    </tr>
    <tr>
        <td class="auto-style2">Home phone number</td>
        <td class="auto-style2">
            <asp:TextBox ID="PhoneHome" runat="server"></asp:TextBox>
        </td>
        <td class="auto-style2">&nbsp;</td>
    </tr>
    <tr>
        <td class="auto-style3">Comments</td>
        <td class="auto-style3">
            <asp:TextBox ID="Comments" runat="server" Height="54px" TextMode="MultiLine" Width="172px"></asp:TextBox>
        </td>
        <td class="auto-style3"></td>
    </tr>
    <tr>
        <td>&nbsp;</td>
        <td>
            <asp:Button ID="SendButton" runat="server" Text="Send" />
        </td>
        <td>&nbsp;</td>
    </tr>
    <tr>
        <td colspan="3">
            <asp:ValidationSummary ID="ValidationSummary1" runat="server" HeaderText="Please correct the following errors:" />
        </td>
    </tr>
</table>

