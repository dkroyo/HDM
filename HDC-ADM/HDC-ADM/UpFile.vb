Imports System
Imports System.IO
Imports System.IO.File
Imports HDC_ADM.ADM
'Check if lates fire on sever is available
Public Class UpFile
    Public Sub New()
        Try
            err = 0

            Dim Q As String = """"
            Dim EXEname As String = "HDC-ADM.exe"
            'set Updated File Directorty
            'Dim myupdir As String = "C:\dBank\0002.DK.Reserves\Projects\0007.AttendanceMonitoringSheet\AddIns\Updates\"
            'must consider the file Directory
            '==================================================================================================
            Dim myupdir As String = "\\10.193.12.174\HDC_Users\Common\!Projects\HDC_ADM\Application\Updates\"
            '==================================================================================================
            'check if directory is True
            If Exists(myupdir) = True Then
                Beep()
                MsgBox("File Update Verification Error! Server connection not available.", vbOKOnly)
            Else
                'GET ACTUAL FILE UPDATE Time
                Dim myfilevername As String = Directory.GetCurrentDirectory
                myfilevername = Application.ExecutablePath
                Dim myfilever As DateTime = File.GetLastWriteTime(myfilevername)

                'get server file update time
                Dim myfileupname As String = myupdir & EXEname
                If Exists(myfileupname) = True Then
                    Dim myupfilever As DateTime = File.GetLastWriteTime(myfileupname)

                    LL("Server Application: " & myupfilever.ToString, 1)
                    LL("Current Application: " & myfilever.ToString, 2)

                    If myupfilever > myfilever Then

                        'Latest on server
                        If (MsgBox("The Latest File is Available on server!" & vbNewLine & "Do you wish to update now?", vbYesNo, "File Updates!") = vbYes) Then
                            'need to update.
                            Dim upbat As String = Directory.GetCurrentDirectory & "\update.bat"
                            If File.Exists(upbat) Then
                                File.Delete(upbat)
                            End If
                            '-----
                            Using sw As StreamWriter = CreateText(upbat)
                                sw.WriteLine("@echo on")
                                sw.WriteLine("timeout 5")
                                sw.WriteLine("copy " & Q & myfileupname & Q & " " & Q & myfilevername & Q, True)
                                sw.WriteLine("start " & myfilevername)
                                sw.WriteLine("del " & Q & upbat & Q)
                                sw.WriteLine("exit")

                            End Using
                            'START the Batch file
                            Process.Start(upbat)
                            'close the application!
                            ADM.Close()
                        Else
                            'no updates performed.
                        End If
                    Else
                        'OK version
                    End If
                Else
                    Beep()
                    MsgBox("Cannot Find EXE File on updates server!.", vbOKOnly)
                End If

            End If
        Catch ex As Exception
            Beep()
            LL(ex.Message.ToString, 3)
            err = 1
        End Try
    End Sub
End Class
