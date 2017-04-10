Imports System.IO
Imports System.Text

<CLSCompliant(True)> _
Public Class ErrorLogger
    Public Sub New()

    End Sub
    Public Sub WriteToErrorLog(ByVal msg As String, ByVal stkTrace As String, ByVal title As String)
        Dim errorPath As String = Module1.rcErrorPath
        Dim client As String = Module1.client

        If Not System.IO.Directory.Exists(rcErrorPath) Then
            System.IO.Directory.CreateDirectory(rcErrorPath)
        End If

        ''Dim fs As FileStream = New FileStream(rcErrorPath & "\errlog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        ''Dim s As StreamWriter = New StreamWriter(fs)
        ''s.Close()
        ''fs.Close()

        Dim fs1 As FileStream = New FileStream(rcErrorPath & "\errlog.txt", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)
        s1.Write("Title: " & client & " - " & title & vbCrLf)
        s1.Write("Message: " & msg & vbCrLf)
        s1.Write("StackTrace: " & stkTrace & vbCrLf)
        s1.Write("Date/Time: " & DateTime.Now.ToString() & vbCrLf)
        s1.Write("===========================================================================================" & vbCrLf)
        s1.Close()
        fs1.Close()

        If System.IO.File.Exists(rcErrorPath & "failure.txt") Then
            Dim fs As FileStream = New FileStream(rcErrorPath & "\failure.txt", FileMode.Append, FileAccess.Write)
            Dim s As StreamWriter = New StreamWriter(fs)
            s.WriteLine(client & " - " & title & " - " & msg)
            s.Close()
            fs.Close()
        Else
            Using sw As StreamWriter = File.CreateText(errorPath & "\failure.txt")
                sw.WriteLine(Err)
            End Using
        End If

    End Sub
End Class
