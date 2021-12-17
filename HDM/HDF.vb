Imports System
Imports HDM.Form1
Imports MySql.Data.MySqlClient

Module HDF
    Public colorrandomizer As Integer = 0

    Public Sub myquery()
        Dim conn As MySqlConnection

        conn = New MySqlConnection()
        '34.87.110.210
        conn.ConnectionString = cons
        Try
            conn.Open()
            MessageBox.Show("OK")
            conn.Close()
        Catch myerror As MySqlException
            MessageBox.Show("Error: " & myerror.Message)

        Finally
            conn.Dispose()
        End Try


    End Sub
    Public Sub radiovalchange()

    End Sub



    Public Sub notif(ByRef nmessage As String, Optional ByRef ncolor As Integer = 0)
        With Form1

            If ncolor = 3 Then
                .noti.SelectionColor = Color.Red
            Else
                If colorrandomizer = 0 Then
                    .noti.SelectionColor = Color.LimeGreen
                    colorrandomizer = 1
                Else
                    .noti.SelectionColor = Color.Cyan
                    colorrandomizer = 0
                End If
            End If

            .noti.AppendText(vbCrLf & " " & nmessage)
            .Refresh()
            .noti.ScrollToCaret()

        End With
    End Sub


End Module
