Option Strict Off
Option Explicit On
Module SubMain
    Public Sub Main()
        Try
            Dim oConnection As Connection
            oConnection = New Connection
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
        End Try
    End Sub

End Module
