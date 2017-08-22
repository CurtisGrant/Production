Imports System.IO
Imports System.Text

<CLSCompliant(True)>
Public Class ErrorLogger
    Public Sub New()

    End Sub

    Public Sub WriteToErrorLog(ByVal msg As String, ByVal stkTrace As String, ByVal title As String)

        'check and make the directory if necessary
        If Not System.IO.Directory.Exists(Module1.errorPath) Then
            System.IO.Directory.CreateDirectory(Module1.errorPath)
        End If

        'check the file
        Dim fs As FileStream = New FileStream(Module1.errorPath & "\errlog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim s As StreamWriter = New StreamWriter(fs)
        s.Close()
        fs.Close()

        'log it
        Dim fs1 As FileStream = New FileStream(Module1.errorPath & "\errlog.txt", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)
        s1.Write("Title: " & title & vbCrLf)
        s1.Write("Message: " & msg & vbCrLf)
        s1.Write("StackTrace: " & stkTrace & vbCrLf)
        s1.Write("Date/Time: " & DateTime.Now.ToString() & vbCrLf)
        s1.Write("===========================================================================================" & vbCrLf)
        s1.Close()
        fs1.Close()



    End Sub
End Class