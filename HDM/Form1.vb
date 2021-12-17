Imports HDM.HDF
Imports MySql.Data.MySqlClient

Public Class Form1

    Public Shared cons As String = "server=34.87.110.210;user id=fj;password=unix;database=hdm"

    Public rhid As String, rBID As String, rHDNAME As String, rHDDATE As String, rHDTIME As String, rHDADDRESS As String, rHDPHONE As String, rHDMAIL As String, rHDTEMP As String, rBSUBMIT As String

    Private Sub loginpass_TextChanged(sender As Object, e As EventArgs) Handles loginpass.TextChanged

    End Sub

    Private Sub loginuser_TextChanged(sender As Object, e As EventArgs) Handles loginuser.TextChanged

    End Sub

    Private Sub TestConnectionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TestConnectionToolStripMenuItem.Click
        myquery()
    End Sub

    Private Sub BID_Click(sender As Object, e As EventArgs) Handles BID.Click
        If hdid.Text = "" Then Exit Sub
        pb(0)
        notif("Verifying User")
        'the code for now is the simplist
        Dim conn As MySqlConnection

        conn = New MySqlConnection()
        '34.87.110.210
        conn.ConnectionString = cons
        Try
            conn.Open()
            pb(50)
            'the connection is OK
            notif("Connected to the database")
            Dim str As String = "SELECT * FROM user WHERE emp_no like '" & hdid.Text & "'"

            Dim sqlcmd As New MySqlCommand
            With sqlcmd
                .CommandText = str
                .Connection = conn
            End With

            Application.DoEvents()

            Dim rd As MySqlDataReader = sqlcmd.ExecuteReader

            pb(75)
            If rd.Read Then
                HDNAME.Text = rd("fname").ToString & " " & rd("lname").ToString
                HDMAIL.Text = rd("email").ToString
                HDPHONE.Text = rd("mobile").ToString
                HDADDRESS.Text = rd("blk").ToString & " " & rd("street").ToString &
                    " " & rd("barangay").ToString & " " & rd("city").ToString & " " & rd("province").ToString
            End If

            HDDATE.Value = Today
            HDTIME.Value = Now

            rd.Close()
            conn.Close()

        Catch myerror As MySqlException
            notif("Error: " & myerror.Message, 3)

        Finally
            conn.Dispose()
        End Try
        notif("Finished Checking")
        pb(100)
    End Sub

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
                cuser.Text = "Current User: Administrator"
                clevel.Text = "Access Level: 3"
                notif("Signing in...")
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

        Me.llogin.Select()

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
        notif("")
        notif("")
        notif("")
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

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        pb(0)
        noti.Text = ""
        cuser.Text = "Sign In"
        clevel.Text = "No Access"
        loginuser.Select()
        loginuser.Focus()
        notif("Please Login to Continue")
        pb(100)
    End Sub

    Private Sub loginuser_KeyDown(sender As Object, e As KeyEventArgs) Handles loginuser.KeyDown
        If e.KeyCode = Keys.Enter Then
            loginpass.Select()
        End If
    End Sub

    Private Sub loginpass_KeyDown(sender As Object, e As KeyEventArgs) Handles loginpass.KeyDown
        If e.KeyCode = Keys.Enter Then
            Blogin.Select()
            Blogin.PerformClick()
        End If
    End Sub
End Class
