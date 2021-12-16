Imports System
Imports HDM.Form1

Module HDF
    Public colorrandomizer As Integer = 0

    Public Sub radiovalchange()

    End Sub



    Public Sub notif(ByRef nmessage As String, Optional ByRef ncolor As String = "")
        With Form1
            .noti.AppendText(vbCrLf & " " & nmessage)

            If colorrandomizer = 0 Then
                .noti.SelectionColor = Color.LimeGreen
                colorrandomizer = 1
            Else
                .noti.SelectionColor = Color.Cyan
                colorrandomizer = 0
            End If

            .Refresh()
            .noti.ScrollToCaret()

        End With
    End Sub


End Module
