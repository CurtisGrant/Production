Imports System.IO
Imports System.Text

<CLSCompliant(True)> _
Public Class ErrorLogger
    Public Sub New()

    End Sub

    Public Sub WriteToErrorLog(ByVal msg As String, ByVal stkTrace As String, ByVal title As String)
        Dim errorPath As String = Module1.rcErrorPath
        If Not System.IO.Directory.Exists(rcErrorPath) Then
            System.IO.Directory.CreateDirectory(rcErrorPath)
        End If
        Dim fs As FileStream = New FileStream(rcErrorPath & "\errlog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim s As StreamWriter = New StreamWriter(fs)
        s.Close()
        fs.Close()
        Dim fs1 As FileStream = New FileStream(rcErrorPath & "\errlog.txt", FileMode.Append, FileAccess.Write)
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
