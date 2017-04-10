Imports System.IO
Imports System.Text
Imports System.Xml

<CLSCompliant(True)> _
Public Class ErrorLogger
    Public Sub New()

    End Sub

    Public Sub WriteToErrorLog(ByVal msg As String, ByVal stkTrace As String, ByVal title As String)
        Dim errorPath As String = Module1.errorLog
        Dim client As String = Module1.client

        If Not System.IO.Directory.Exists(errorPath) Then
            System.IO.Directory.CreateDirectory(errorPath)
        End If

        Dim fs As FileStream = New FileStream(errorPath & "\errlog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim s As StreamWriter = New StreamWriter(fs)
        s.Close()
        fs.Close()

        Dim fs1 As FileStream = New FileStream(errorPath & "\errlog.txt", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)
        s1.Write("Title: " & client & " - " & title & vbCrLf)
        s1.Write("Message: " & msg & vbCrLf)
        s1.Write("StackTrace: " & stkTrace & vbCrLf)
        s1.Write("Date/Time: " & DateTime.Now.ToString() & vbCrLf)
        s1.Write("===========================================================================================" & vbCrLf)
        s1.Close()
        fs1.Close()

        If System.IO.File.Exists("\\LPS2-2\RetailClarity\RCSystem\SYSFAIL\failure.txt") Then
        Else
            File.Create("\\LPS2-2\RetailClarity\RCSystem\SYSFAIL\failure.txt")
        End If

    End Sub
End Class

