
Partial Class Demos_Target:
    Inherits BasePage

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        Label1.Text = Request.QueryString.ToString()
    End Sub
End Class
