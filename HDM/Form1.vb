Imports HDM.HDF
Public Class Form1

    Public rhid As String, rBID As String, rHDNAME As String, rHDDATE As String, rHDTIME As String, rHDADDRESS As String, rHDPHONE As String, rHDMAIL As String, rHDTEMP As String, rBSUBMIT As String



    Public fever As String, pagod As String, tae As String, ulo As String, ubo As String, suka As String, sore As String, body As String, lost As String, dob As String, f2f As String, f2c As String, gtravel As String, ltravel As String


    'this section will handle the login
    Private Sub Blogin_Click(sender As Object, e As EventArgs) Handles Blogin.Click
        If loginpass.Text = "" Or loginuser.Text = "" Then
            Beep()
            notif("Both Username/Password is required!", 3)
            loginuser.Select()
        Else
            'create a hard coded login for administration and Debug
            Dim myuser As String = "dk", mypass As String = "dk"
            Dim youruser As String = loginuser.Text, yourpass As String = loginpass.Text


            If myuser = youruser And mypass = yourpass Then
            Else
                Beep()
                notif("Username/Password is incorrect!", 3)
                loginuser.Select()
                Exit Sub
            End If

            'if the hardcoded password is correct
            TC.TabPages.Remove(llogin)
            TC.TabPages.Insert(0, tmon)
            TC.TabPages.Insert(1, treg)
            TC.TabPages.Insert(2, tadmin)

            inithdfvalues() 'need to reset the values of the variables

        End If
    End Sub

    'Will be done during start up
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        noti.Text = "" 'Clean the notification area prior loading


        'Hide all Tab First
        TC.TabPages.Remove(llogin)
        TC.TabPages.Remove(treg)
        TC.TabPages.Remove(tmon)
        TC.TabPages.Remove(tadmin)

        'Show Tabs
        TC.TabPages.Insert(0, llogin)


    End Sub

    Private Sub yfever_CheckedChanged(sender As Object, e As EventArgs) Handles yfever.CheckedChanged

        If yfever.Checked = True Then
            fever = "YES"
        Else
            fever = "NO"
        End If
        notif("You selected " & fever & " for Fever (Lagnat).")
    End Sub




    Sub pb(ByRef Progrezz As Long)
        PBAR.Value = Progrezz
        Refresh()
    End Sub


    'this function will ensure that the program will start at fresh condition
    Sub inithdfvalues()
        pb(PBAR.Minimum)
        Dim c As Integer = 1
        notif("Initializing Health Declaration Form")
        fever = ""
        pagod = ""
        pb(10)
        tae = ""
        ulo = ""
        ubo = ""
        suka = ""
        sore = ""
        body = ""
        lost = ""
        pb(50)
        dob = ""
        f2f = ""
        f2c = ""
        gtravel = ""
        ltravel = ""
        notif("Finished Intializing the Health Declaration Form to preset values")
        pb(PBAR.Maximum)
    End Sub


End Class
