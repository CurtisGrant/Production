
Partial Class Demos_SalesChartMain
    Inherits System.Web.UI.Page
    Public theDate As Date
    Protected Sub txtDate_TextChanged(sender As Object, e As EventArgs) Handles txtDate.TextChanged
        theDate = CDate(txtDate.Text)
    End Sub
End Class
