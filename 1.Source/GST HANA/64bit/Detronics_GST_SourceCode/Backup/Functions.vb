Imports System.IO
Public Class Functions
    Public Shared Sub WriteLog(ByVal Str As String)
        Dim oWrite As IO.StreamWriter
        Dim FilePath As String
        FilePath = Application.StartupPath + "\ErrorLog.txt"
        If IO.File.Exists(FilePath) Then
            oWrite = IO.File.AppendText(FilePath)
        Else
            oWrite = IO.File.CreateText(FilePath)
        End If
        oWrite.Write(Now.ToString() + ":" + Str + vbCrLf)
        oWrite.Close()
    End Sub
End Class
