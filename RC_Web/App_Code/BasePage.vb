Imports Microsoft.VisualBasic

Public Class BasePage
    Inherits System.Web.UI.Page
    Private Sub Page_PreRender(sender As Object, e As EventArgs) Handles Me.PreRender
        If String.IsNullOrEmpty(Me.Title) OrElse Me.Title.Equals("Untitled Page",
                                                                 StringComparison.CurrentCultureIgnoreCase) Then
            Throw New Exception(
                "Page Title cannot be ""Untitled Page"" or an empty string,")
        End If
    End Sub
End Class
