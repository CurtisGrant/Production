Imports System.Data.SqlClient

Partial Class Tools_CreateReorders:
    Inherits BasePage
    Private Shared conString, servr, datbase, client, sql As String
    Private Shared con, con2, con3 As SqlConnection
    Private Shared cmd As SqlCommand
    Private Shared rdr As SqlDataReader
    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        conString = "Server=CURTIS-MOBILE;Initial Catalog=TCM;Integrated Security=True"
        con = New SqlConnection(conString)
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If IsPostBack Then
            Dim path As String = Server.MapPath("~/UploadedFiles/")
            Dim fileOK As Boolean = False
            If FileUpload1.HasFile Then
                Dim fileExtension As String
                fileExtension = System.IO.Path. _
                    GetExtension(FileUpload1.FileName).ToLower()
                Dim allowedExtensions As String() = _
                    {".jpg", ".jpeg", ".png", ".gif"}
                For i As Integer = 0 To allowedExtensions.Length - 1
                    If fileExtension = allowedExtensions(i) Then
                        fileOK = True
                    End If
                Next
                If fileOK Then
                    Try
                        FileUpload1.PostedFile.SaveAs(path & _
                             FileUpload1.FileName)
                        Label1.Text = "File uploaded!"
                    Catch ex As Exception
                        Label1.Text = "File could not be uploaded."
                    End Try
                Else
                    Label1.Text = "Cannot accept files of this type."
                End If
            End If
        End If

    End Sub
End Class
