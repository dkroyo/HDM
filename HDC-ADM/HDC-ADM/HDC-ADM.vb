Imports System
Imports System.IO
Imports System.IO.File
Imports HDC_ADM.ADM_COMMON
Imports MySql.Data.MySqlClient


Public Class ADM
    Public Shared err As Integer = 0
    Public Shared maindb As String = "adm_main"
    Public Shared mirrordb As String = "adm_mirror"
    Public Shared userdb As String = "adm_user"
    Public Shared refdb As String = "adm_ref"
    Public Shared cons As String = "server=10.193.20.172;port=3306;uid=fj;pwd=unix;database=adm;Sslmode=None;"
    Public Shared smode1 As Long = 0
    Public Shared smode2 As Long = 0

    'declaring General Public Information
    Public Shared update_dir As String =
        "\\10.193.12.174\HDC_Users\Common\!Projects\HDC_ADM\Application\Updates\" 'this address must end with \ (back slash) 'it is server update data info
    Public Shared app_dir As String =
        Directory.GetCurrentDirectory() & "\" ' this is the application default directory and also must end with back slash (\)
    '# public values
    Public Shared hdcsn As String, fa As String, cal As String, sn As String, exp As String, eqmaker As String, eqmodel As String, eqtype As String,
        eqdisp As String, eqsite As String, eqbuild As String, eqroom As String, eqfloor As String, eqloc As String, eqpic As String, eqhost As String,
        equser As String, eqpass As String, eqip As String, eqos As String, eqantv As String, eqstatus As String, eqcals As String, eqcalD As String,
        eqcaldue As String, todate As String, totime As String, eqremarks As String, sql As String, search_str As String, aDate As String, atime As String,
        ID As String, user_login As String, user_passowd As String, user_fname As String, user_lname As String, user_pos As String, user_mail As String,
        user_boss As String, user_level As String, myusername As String, mypassword As String, memail As Boolean, bossmail As Boolean


    'this is the starting point! 
    Private Sub ADM_Load(sender As Object, e As EventArgs) Handles Me.Load
        'do not display the main Tab Yet
        '============================================================
        main_tab.DrawMode = TabDrawMode.OwnerDrawFixed
        With Me.main_tab.TabPages
            .Remove(adm_tab)
            .Remove(ref_tab)
            '.Remove(signup_tab)
            .Remove(query_tab)
        End With
        '============================================================

        'this for the sql file
        If File.Exists(app_dir & "MySql.Data.dll") Then
        Else
            File.Copy(update_dir & "MySql.Data.dll", app_dir & "MySql.Data.dll", True)
        End If
        Application.DoEvents()
        Qhistory_check.Checked = True 'for Query page as Read on Mirror server by default condition

        LogOutToolStripMenuItem.Enabled = False

        login_B.Enabled = False
        Application.DoEvents()
        Refresh()
        ender()

    End Sub


    'Form Ini Function
    Private Sub ADM_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Refresh()
        Threading.Thread.Sleep(1000)
        Refresh()
        pgmin()

        user_label.Text = ""
        'setting progress bar values!
        PGB.Maximum = PGB.Maximum
        PGB.Minimum = PGB.Minimum

        hdcsn_check.Enabled = False
        fa_check.Enabled = False
        cal_check.Enabled = False
        exp_check.Enabled = False
        sn_check.Enabled = False

        'setting the file name
        Me.Text = "HDC- ADM " & String.Format("Version {0}", My.Application.Info.Version.ToString)
        '
        LL("Initializing Managament System!", 1)
        err = 0
        'checking server connection
        server_connect(0)
        If err > 0 Then
            pgmax()
            Exit Sub
        End If
        '-----------------------------------------
        'check for file updates'
        Dim upmyfile As New UpFile
        If err > 0 Then
            pgmax()
            Exit Sub '<-- Always check every after process if ok
        End If

        'clean the form on start up
        clr_val()
        If err > 0 Then
            pgmax()
            Exit Sub '<-- Always check every after process if ok
        End If

        'drop down / combo box initialization
        If cbox_ini(0) > 1 Then
            pgmax()
            Exit Sub
        End If

        'get date and time val  
        If MyDataTime(0) > 0 Then
            pgmax()
            Exit Sub
        End If


        bossmail = False
        IncludeImmidiateSupperiorOnMaleToolStripMenuItem.Text = "Include Immidiate Supperior On Mail: NO"
        IncludeImmidiateSupperiorOnMaleToolStripMenuItem.BackColor = Color.Red

        memail = True
        SendMailUpdateToolStripMenuItem.Text = "Send Mail Update: YES"
        SendMailUpdateToolStripMenuItem.BackColor = Color.LightGreen

        rereadmode = "NO"
        ReReadOnUpdateToolStripMenuItem.BackColor = Color.Red
        ReReadOnUpdateToolStripMenuItem.ForeColor = Color.White

        'bago areh

        '-----------------------
        'dummy()
        '-----------------------
        login_B.Enabled = True
        ender()
        pgmax()
        'set focus on login
        user_box.Select()
        user_box.Focus()

        'restrict Access to Productivity menu
        EditToolStripMenuItem.Enabled = False
        ViewToolStripMenuItem.Enabled = False
        ToolsToolStripMenuItem.Enabled = False



        'this is temporary only
        'user_box.Text = "dk"
        'password_box.Text = "dk"
        'login_B.PerformClick()

    End Sub


    '## LOG OUT Button
    Private Sub LogOutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogOutToolStripMenuItem.Click
        logoutmode = 0
        err = 0
        Try
            pgmin()

            'If login_tab.Enabled = True Then
            '    LL("Unable to comply! You are not Logged In.", 3)
            '    Exit Sub
            'End If

            With Me.main_tab.TabPages
                .Insert(1, login_tab)
            End With
            Refresh()

            'this code is when login is successful
            With Me.main_tab.TabPages
                .Remove(adm_tab)
                .Remove(query_tab)
                ' If signup_tab.Enabled = True Then
                '.Remove(signup_tab)
                ' End If

                If user_level > 2 Then
                    .Remove(ref_tab)
                End If
            End With

            user_login = ""
            user_passowd = ""
            myusername = ""
            mypassword = ""
            user_fname = ""
            user_lname = ""
            user_pos = ""
            user_mail = ""
            user_boss = ""
            user_level = ""

            user_label.Text = ""

            clr_val()

            'restrict Access to Productivity menu
            EditToolStripMenuItem.Enabled = False
            ViewToolStripMenuItem.Enabled = False
            ToolsToolStripMenuItem.Enabled = False



        Catch ex As Exception
            Beep()
            LL(ex.Message.ToString, 3)
        End Try
        pgmax()
    End Sub




    '## LOGIn BUTTON
    Private Sub login_B_Click(sender As Object, e As EventArgs) Handles login_B.Click
        err = 0
        pgmin()
        Dim con As New MySqlConnection()
        Try
            user_label.Text = ""
            err = 0
            user_login = user_login
            user_passowd = user_passowd
            myusername = myusername
            mypassword = mypassword
            user_fname = user_fname
            user_lname = user_lname
            user_pos = user_pos
            user_mail = user_mail
            user_boss = user_boss
            user_level = user_level

            user_login = user_box.Text
            user_passowd = password_box.Text

            If user_login = "" Or user_passowd = "" Then
                Beep()
                LL("Incorrect User/Password!", 3)
                ender()
                pgmax()
                Exit Sub
            End If

            Dim ssrt As String = "SELECT * FROM adm_user where USERNAME = '" & user_login & "';"
            Application.DoEvents()

            con.ConnectionString = cons
            con.Open()

            Dim sqlcmd As New MySqlCommand
            With sqlcmd
                .CommandText = ssrt
                .Connection = con
            End With

            Dim rd As MySqlDataReader = sqlcmd.ExecuteReader

            If rd.Read Then

                user_fname = rd("FIRST_NAME").ToString
                user_lname = rd("LAST_NAME").ToString
                user_pos = rd("POSITIONx").ToString
                user_mail = rd("MyEmail").ToString
                user_boss = rd("SupEMail").ToString
                user_level = rd("CLEARANCE")
                myusername = rd("USERNAME").ToString
                mypassword = rd("PASSWORD").ToString

                If user_passowd <> mypassword Then
                    LL("Incorrect User/Password!!", 3)
                    password_box.Text = ""
                    user_passowd = ""
                    user_login = ""
                    Beep()
                    pgmax()
                    Exit Sub
                Else
                End If

                user_label.Text = "Current User: " & user_fname & " " & user_lname & " > Access Level: " & user_level

                LL("Login Succesful", 1)
                '============================================================================
                'still need to add the login argument but for the mean time no problem

                'restrict Access to Productivity menu
                EditToolStripMenuItem.Enabled = True
                ViewToolStripMenuItem.Enabled = True
                ToolsToolStripMenuItem.Enabled = True

                With Me.main_tab.TabPages
                    .Insert(1, query_tab)
                    .Insert(1, adm_tab)
                End With
                Refresh()

                'this code is when login is successful
                With Me.main_tab.TabPages
                    .Remove(login_tab)
                End With

                If user_level < 2 Then
                    update_B.Enabled = False
                    sql_B.Enabled = False
                Else
                    update_B.Enabled = True
                    sql_B.Enabled = True
                End If

                If user_level > 2 Then
                    With Me.main_tab.TabPages
                        .Insert(2, ref_tab)
                    End With
                End If

                If user_level < 5 Then
                    sm1.Enabled = False
                    sm2.Enabled = False
                    smode2 = 0
                    smode1 = 0
                Else
                    sm1.Enabled = True
                    sm2.Enabled = true
                End If

                '============================================================================
                logoutmode = 0
                Refresh()
                'user_passowd = ""
                'user_login = ""
                Application.DoEvents()
                QmyEQ_B.Text = user_fname & " Equip. List"

                Application.DoEvents()

                hdcsn_box.Select()
                hdcsn_box.Focus()
                LogOutToolStripMenuItem.Enabled = True
            Else
                Beep()
                LL("Incorrect User/Password!!", 3)
                password_box.Text = ""
                user_passowd = ""
                user_login = ""
                ender()
            End If
            rd.Close()

        Catch ex As Exception
            con.Close()
            Beep()
            LL(ex.Message.ToString, 3)
        End Try
        sql_box.Text = ""
        pgmax()
        ender()
    End Sub


    '## REad Button Handler
    Private Sub get_B_Click(sender As Object, e As EventArgs) Handles get_B.Click
        If user_level < 1 Then
            LL("Unauthorized Access!", 3)
            LL(user_fname & " your access level is only " & user_level, 1)
            ender()
            pgmax()
            Exit Sub
        End If

        err = 0
        pgmin()

        err = 0
        'This function is for read button
        'If txt_val(0) = 1 Then Exit Sub 'check weahter reference is OK
        searchCheck(0)
        If ID = "" Then
            LL("Please Input Search required Parameter first!", 3)
            Beep()
            ender()
            pgmax()
            Exit Sub
        End If
        pgval(10)
        'this is system generated date and Time.
        If MyDataTime(0) > 0 Then
            pgmax()
            Exit Sub 'must be before Getting values
        End If
        pgval(20)
        'collecting values on the boxes
        ''''If get_val(0) > 0 Then
        ''''    pgmax()
        ''''    Exit Sub 'not needed here~!
        ''''End If

        'create the read sql string
        If readsqlstr(0) > 0 Then
            pgmax()
            Exit Sub
        End If
        pgval(60)

        If manual_sql(0) > 0 Then
            pgmax()
            Exit Sub
        End If

        LL("Finished Query!", 1)
        pgmax()
        ender()
        'Back to index Position
        hdcsn_box.Focus()
        hdcsn_box.Select()

    End Sub

    '## Manual SQL
    Private Sub sql_B_Click(sender As Object, e As EventArgs) Handles sql_B.Click
        pgmin()
        err = 0
        If sql_box.Text = "" Then
            LL("Please provide SQL First!", 3)
            pgmax()
            Exit Sub
        End If
        pgval(50)
        sql = sql_box.Text
        Application.DoEvents()
        Refresh()
        clr_val()

        If manual_sql(0) > 0 Then
        End If
        pgmax()
    End Sub


    '## UPDATE BUTTON
    Private Sub update_B_Click(sender As Object, e As EventArgs) Handles update_B.Click

        pgmin()
        err = 0
        'this button will handle updating

        If user_level < 2 Then
            pgmax()
            Exit Sub
        End If
        pgval(10)
        'this is system generated date and Time.
        If MyDataTime(0) > 0 Then
            pgmax()
            Exit Sub 'must be before Getting values
        End If
        pgval(30)
        If txt_val(0) = 1 Then
            pgmax()
            Exit Sub
        End If

        'collecting values on the boxes
        If get_val(0) > 0 Then
            pgmax()
            Exit Sub 'not needed here~!
        End If
        pgval(60)
        'writing code4
        If write_sql(0) > 0 Then
            pgmax()
            Exit Sub
        End If

        '==================================== MAIL FUNCTION HANDLER ============================================
        LL("Initializing Mail Notification!", 0)
        'this not proper here:
        user_mail = user_mail
        user_boss = user_boss
        Application.DoEvents()
        'this function is to check to whom mail is needed to be sent!
        Dim new_user_mail As String, new_boss_mail As String

        If memail = True Then
            new_user_mail = user_mail
            LL("An email summary of " & hdcsn & " information will be sent to: " & new_user_mail, 1)
        Else
            new_user_mail = ""
            LL("Personal Mail Notification is dis-abled!", 3)
        End If
        Application.DoEvents()

        If bossmail = True Then
            new_boss_mail = user_boss
            LL("An email summary of " & hdcsn & " information will be sent to your Immediate Superior: " & new_boss_mail, 1)
        Else
            new_boss_mail = ""
            LL("Immediate Superior Mail Notification is dis-abled", 3)
        End If
        Application.DoEvents()
        'mailing Function
        If memail = False And bossmail = False Then
            mailup(new_user_mail, new_boss_mail, 0) 'this is special info gatherer
            LL("No Mail Notification!", 3)
            ender()
            Application.DoEvents()
            pgmax()
        Else
            If (mailup(new_user_mail, new_boss_mail, 0)) < 1 Then
                LL("Mail Sent", 1)
            Else
                ender()
                Beep()
                pgmax()
                Exit Sub
            End If
        End If
        Application.DoEvents()
        '==================================== MAIL FUNCTION HANDLER END ========================================
        pgval(80)
        clr_val()
        hdcsn_box.Text = hdcsn

        If rereadmode = "YES" Then
            get_B.PerformClick()
        Else
        End If

        LL("Finished Updating Records!", 2)

        ender()
        pgmax()
    End Sub

    Private Sub sn_check_CheckedChanged(sender As Object, e As EventArgs) Handles sn_check.CheckedChanged
        'sn_box.Select()
    End Sub


    'Clear Values Menu Button
    Private Sub clr_menu_Click(sender As Object, e As EventArgs) Handles clr_menu.Click
        pgmin()
        'this is to clear values on the boxes
        sql_box.Text = ""
        clr_val()
        ender()
        pgmax()
    End Sub


    '## the following functions is related to Selection of default Seach Condition!
    Private Sub hdcsn_box_Enter(sender As Object, e As EventArgs) Handles hdcsn_box.Enter
        hdcsn_check.Checked = True
    End Sub
    Private Sub fa_box_Enter(sender As Object, e As EventArgs) Handles fa_box.Enter
        fa_check.Checked = True
    End Sub


    Private Sub cal_box_Enter(sender As Object, e As EventArgs) Handles cal_box.Enter
        cal_check.Checked = True
    End Sub
    Public rereadmode As String = "NO"
    Private Sub ReReadOnUpdateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReReadOnUpdateToolStripMenuItem.Click
        If rereadmode = "YES" Then
            rereadmode = "NO"
            ReReadOnUpdateToolStripMenuItem.BackColor = Color.Red
            ReReadOnUpdateToolStripMenuItem.ForeColor = Color.White
        Else
            rereadmode = "YES"
            ReReadOnUpdateToolStripMenuItem.BackColor = Color.Green
            ReReadOnUpdateToolStripMenuItem.ForeColor = Color.White
        End If
        ReReadOnUpdateToolStripMenuItem.Text = "Re-Read on Update: " & rereadmode
    End Sub

    '# the following handle toggle
    Private Sub exp_box_Enter(sender As Object, e As EventArgs) Handles exp_box.Enter
        exp_check.Checked = True
    End Sub

    Private Sub exp_check_CheckedChanged(sender As Object, e As EventArgs) Handles exp_check.CheckedChanged
        'exp_box.Select()
    End Sub

    Private Sub cal_check_CheckedChanged(sender As Object, e As EventArgs) Handles cal_check.CheckedChanged
        'cal_box.Select()
    End Sub

    Private Sub fa_check_CheckedChanged(sender As Object, e As EventArgs) Handles fa_check.CheckedChanged
        'fa_box.Select()
    End Sub

    Private Sub hdcsn_check_CheckedChanged(sender As Object, e As EventArgs) Handles hdcsn_check.CheckedChanged
        'hdcsn_box.Select()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub sn_box_Enter(sender As Object, e As EventArgs) Handles sn_box.Enter
        sn_check.Checked = True
    End Sub


    'handles login entry
    Private Sub login_tab_Enter(sender As Object, e As EventArgs) Handles login_tab.Enter
        'login
        user_box.Text = ""
        password_box.Text = ""
    End Sub

    '-----------------------------------------------------------------------------------------
    '## Forgot Password
    Private Sub forgot_pass_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles forgot_pass.LinkClicked
        Beep()
        ''LL("Not Yet Supported", 1)
        If (F_Passowd(0) < 1) Then
        Else
        End If
        ender()
    End Sub

    Private Sub sm1_Click(sender As Object, e As EventArgs) Handles sm1.Click
        If smode1 = 0 Then
            smode1 = 1
            LL("Current Clearnace" & user_level, 1)
            sm1.BackColor = Color.Azure
            sm1.Text = "Special Mode 1: Enabled"
            LL("Special Mode 1: Enabled", 1)
            smode2 = 0
            sm2.BackColor = Color.Pink
            sm2.Text = "Special Mode 2: Dis-abled"
            LL("Special Mode 2: Dis-abled", 1)
        Else
            smode1 = 0
            LL("Current Clearnace" & user_level, 1)
            sm1.BackColor = Color.Pink
            sm1.Text = "Special Mode 1: Dis-abled"
            LL("Special Mode 1: Dis-abled", 3)
        End If



    End Sub

    Private Sub sm2_Click(sender As Object, e As EventArgs) Handles sm2.Click
        If smode2 = 0 Then
            smode2 = 1
            LL("Current Clearnace" & user_level, 1)
            sm2.BackColor = Color.Azure
            sm2.Text = "Special Mode 2: Enabled"
            LL("Special Mode 2: Enabled", 1)
            smode1 = 0
            sm1.BackColor = Color.Pink
            sm1.Text = "Special Mode 1: Dis-abled"
            LL("Special Mode 1: Dis-abled", 3)
        Else
            smode2 = 0
            LL("Current Clearnace" & user_level, 1)
            sm2.BackColor = Color.Pink
            sm2.Text = "Special Mode 2: Dis-abled"
            LL("Special Mode 2: Dis-abled", 1)
        End If
    End Sub





    Private Sub user_box_KeyDown(sender As Object, e As KeyEventArgs) Handles user_box.KeyDown
        If e.KeyCode = Keys.Enter Then
            password_box.Select()
        End If
    End Sub

    Private Sub password_box_KeyDown(sender As Object, e As KeyEventArgs) Handles password_box.KeyDown
        If e.KeyCode = Keys.Enter Then
            login_B.PerformClick()
        End If
    End Sub


    Private Sub hdcsn_box_KeyDown(sender As Object, e As KeyEventArgs) Handles hdcsn_box.KeyDown, fa_box.KeyDown, cal_box.KeyDown, exp_box.KeyDown, sn_box.KeyDown
        If e.KeyCode = Keys.Enter Then
            get_B.PerformClick()
        End If
    End Sub

    'close the App
    Private Sub exit_menu_Click(sender As Object, e As EventArgs) Handles exit_menu.Click
        Me.Close()
    End Sub

    'refresh the values 
    Private Sub RefreshDropDownListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshDropDownListToolStripMenuItem.Click
        pgmin()

        If cbox_ini(0) < 1 Then
        Else
            LL("Error occured during updating Drop Down list option!", 3)
        End If
        pgmax()
    End Sub

    'COMBO BOX Read Button
    Private Sub calgetlist_B_Click(sender As Object, e As EventArgs) Handles calgetlist_B.Click
        cbbox_read("CSTATUS")
        pgmax()
    End Sub

    Private Sub statusgetlist_B_Click(sender As Object, e As EventArgs) Handles statusgetlist_B.Click
        cbbox_read("ASTATUS")
        pgmax()
    End Sub

    Private Sub typegetlist_B_Click(sender As Object, e As EventArgs) Handles typegetlist_B.Click
        cbbox_read("TYPE")
        pgmax()
    End Sub

    Private Sub floorgetlist_B_Click(sender As Object, e As EventArgs) Handles floorgetlist_B.Click
        cbbox_read("DEPT")
        pgmax()
    End Sub

    Private Sub buildgetlist_B_Click(sender As Object, e As EventArgs) Handles buildgetlist_B.Click
        cbbox_read("BUILD")
        pgmax()
    End Sub

    Private Sub sitegetlist_B_Click(sender As Object, e As EventArgs) Handles sitegetlist_B.Click
        cbbox_read("SITE")
        pgmax()
    End Sub

    Private Sub refcal_b_Click(sender As Object, e As EventArgs) Handles refcal_b.Click
        pgmax()

        If refcal_box.Text = "" Then
            LL("Please Input Parameter before Adding", 3)
            ender()
            pgmax()
            Exit Sub
        End If
        cbbox_add("CSTATUS", refcal_box.Text)
        pgmax()
        ender()
    End Sub

    Private Sub refstatus_B_Click(sender As Object, e As EventArgs) Handles refstatus_B.Click
        pgmin()

        If refstatus_box.Text = "" Then
            LL("Please Input Parameter before Adding", 3)
            ender()
            pgmax()
            Exit Sub
        End If
        cbbox_add("ASTATUS", refstatus_box.Text)
        pgmax()
        ender()
    End Sub

    Private Sub reftype_B_Click(sender As Object, e As EventArgs) Handles reftype_B.Click
        pgmin()

        If reftype_box.Text = "" Then
            LL("Please Input Parameter before Adding", 3)
            ender()
            pgmax()
            Exit Sub
        End If
        cbbox_add("TYPE", reftype_box.Text)
        ender()
        pgmax()
    End Sub

    Private Sub reffloor_B_Click(sender As Object, e As EventArgs) Handles reffloor_B.Click
        pgmin()

        If reffloor_box.Text = "" Then
            LL("Please Input Parameter before Adding", 3)
            ender()
            pgmax()
            Exit Sub
        End If
        cbbox_add("DEPT", reffloor_box.Text)
        ender()
        pgmax()
    End Sub

    Private Sub refbuild_B_Click(sender As Object, e As EventArgs) Handles refbuild_B.Click
        pgmin()

        If refbuild_box.Text = "" Then
            LL("Please Input Parameter before Adding", 3)
            ender()
            pgmax()
            Exit Sub
        End If
        cbbox_add("BUILD", refbuild_box.Text)
        ender()
        pgmax()
    End Sub

    Private Sub refsiite_B_Click(sender As Object, e As EventArgs) Handles refsiite_B.Click
        pgmin()

        If refsite_box.Text = "" Then
            LL("Please Input Parameter before Adding", 3)
            ender()
            pgmax()
            Exit Sub
        End If
        cbbox_add("SITE", refsite_box.Text)
        ender()
        pgmax()
    End Sub
    Private Sub clx_Click(sender As Object, e As EventArgs) Handles clx.Click
        pgmin()
        CB_list.Text = ""
        refsite_box.Text = ""
        refbuild_box.Text = ""
        reffloor_box.Text = ""
        reftype_box.Text = ""
        refstatus_box.Text = ""
        refcal_box.Text = ""
        pgmax()
    End Sub

    Private Sub IncludeImmidiateSupperiorOnMaleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IncludeImmidiateSupperiorOnMaleToolStripMenuItem.Click
        If bossmail = False Then
            bossmail = True
            IncludeImmidiateSupperiorOnMaleToolStripMenuItem.Text = "Include Immidiate Supperior On Mail: YES"
            IncludeImmidiateSupperiorOnMaleToolStripMenuItem.BackColor = Color.LightGreen
        Else
            bossmail = False
            IncludeImmidiateSupperiorOnMaleToolStripMenuItem.Text = "Include Immidiate Supperior On Mail: NO"
            IncludeImmidiateSupperiorOnMaleToolStripMenuItem.BackColor = Color.Red
        End If
    End Sub

    Private Sub SendMailUpdateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SendMailUpdateToolStripMenuItem.Click
        If memail = False Then
            memail = True
            SendMailUpdateToolStripMenuItem.Text = "Send Mail Update: YES"
            SendMailUpdateToolStripMenuItem.BackColor = Color.LightGreen
        Else
            memail = False
            SendMailUpdateToolStripMenuItem.Text = "Send Mail Update: NO"
            SendMailUpdateToolStripMenuItem.BackColor = Color.Red
        End If
        pgmax()
    End Sub


    Private Sub auserclr_Click(sender As Object, e As EventArgs) Handles auserclr.Click
        auser_box.Text = ""
        apass_box.Text = ""
        afname_box.Text = ""
        alname_box.Text = ""
        aposition_box.Text = ""
        amymail_box.Text = ""
        abossmail_box.Text = ""
        alevel_box.Text = ""
        Application.DoEvents()
    End Sub


    Private Sub aread_B_Click(sender As Object, e As EventArgs) Handles aread_B.Click
        If auser_up(0) < 1 Then
        Else
        End If
        auser_box.Focus()
        auser_box.Select()
    End Sub

    Private Sub auser_box_KeyDown(sender As Object, e As KeyEventArgs) Handles auser_box.KeyDown
        If e.KeyCode = Keys.Enter Then
            aread_B.PerformClick()
        End If
    End Sub


    Private Sub aupdate_B_Click(sender As Object, e As EventArgs) Handles aupdate_B.Click
        aupdateup(0)
        auser_box.Focus()
        auser_box.Select()
    End Sub
    'this handle the sign up button
    Private Sub SignUpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SignUpToolStripMenuItem.Click
        '
        'LogOutToolStripMenuItem.Enabled = True
        'With Me.main_tab.TabPages
        '    .Insert(1, signup_tab)
        'End With
        main_tab.SelectedTab = signup_tab

    End Sub
    'THIS handles sign up button
    Private Sub bsignup_B_Click(sender As Object, e As EventArgs) Handles bsignup_B.Click

        If signup(0) < 1 Then
        Else
        End If
    End Sub
    'this is password Change Function
    Private Sub PasswordChangeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasswordChangeToolStripMenuItem.Click
        user_passowd = user_passowd
        user_login = user_login
        err = 0
        pgmin()
        Dim con As New MySqlConnection()
        Try
            con.ConnectionString = cons
            con.Open()

            Dim PCsql As New MySqlCommand
            PCsql.CommandText = "select * from adm_user where USERNAME = '" & user_login & "';"
            PCsql.Connection = con
            Dim rd As MySqlDataReader = PCsql.ExecuteReader
            Application.DoEvents()

            If rd.Read Then
                If user_passowd = rd("PASSWORD").ToString Then
                Else
                    rd.Close()
                    err = 1
                    LL("Login Password and Current Password is Invalid, Please try signing in again", 3)
                    con.Close()
                    pgmax()
                    ender()
                    Exit Sub
                End If
            Else
                rd.Close()
                err = 1
                LL("Database Reading Error!", 3)
                con.Close()
                pgmax()
                ender()
                Exit Sub
            End If
            rd.Close()
            Dim newpass As String = InputBox("Please Input New Password")
            Dim newpass2 As String = InputBox("Please Re-Input New Password")

            If newpass <> newpass2 Then
                err = 1
                LL("Password not match!", 3)
                con.Close()
                pgmax()
                ender()
                Exit Sub
            Else
            End If

            Dim sqlcmd As New MySqlCommand
            sqlcmd.CommandText = "UPDATE adm_user SET PASSWORD = '" & newpass & "' WHERE USERNAME = '" & user_login & "';"
            sqlcmd.Connection = con
            sqlcmd.ExecuteNonQuery()
            con.Close()

            LL("Password is Change, Please Login again!", 2)

            LogOutToolStripMenuItem.PerformClick()

            pgmax()
            ender()

        Catch ex As Exception
            err = 1
            LL(ex.Message.ToString, 3)
            con.Close()
            pgmax()
            ender()
        End Try
    End Sub

    Private Sub eqcals_box_SelectedIndexChanged(sender As Object, e As EventArgs) Handles eqcals_box.SelectedIndexChanged
        If eqcals_box.Text.Contains("ACTIVE") Then
            eqcals_box.BackColor = Color.Green
            eqcals_box.ForeColor = Color.White
        End If
        If eqcals_box.Text.Contains("IN ACTIVE") Then
            eqcals_box.BackColor = Color.Pink
            eqcals_box.ForeColor = Color.Black
        End If
        If Not eqcals_box.Text.Contains("ACTIVE") Then
            eqcals_box.BackColor = Color.White
            eqcals_box.ForeColor = Color.Black
        End If
    End Sub

    Private Sub eqstatus_box_Enter(sender As Object, e As EventArgs) Handles eqstatus_box.Enter
        eqstatus_box.BackColor = Color.White
        eqstatus_box.ForeColor = Color.Black
    End Sub
    Private Sub eqcals_box_Enter(sender As Object, e As EventArgs) Handles eqcals_box.Enter
        eqcals_box.BackColor = Color.White
        eqcals_box.ForeColor = Color.Black
    End Sub

    Private Sub eqstatus_box_SelectedIndexChanged(sender As Object, e As EventArgs) Handles eqstatus_box.SelectedIndexChanged
        If eqstatus_box.Text.Contains("ACTIVE") Then
            eqstatus_box.BackColor = Color.Green
            eqstatus_box.ForeColor = Color.White
        End If
        If eqstatus_box.Text.Contains("IN ACTIVE") Then
            eqstatus_box.BackColor = Color.Pink
            eqstatus_box.ForeColor = Color.Black
        End If
        If Not eqstatus_box.Text.Contains("ACTIVE") Then
            eqstatus_box.BackColor = Color.White
            eqstatus_box.ForeColor = Color.Black
        End If

    End Sub
    'history File Check
    Private Sub Qhistory_check_CheckedChanged(sender As Object, e As EventArgs) Handles Qhistory_check.CheckedChanged
        If Qhistory_check.Checked = True Then
            LL("I will display data from History Server", 1)
            ender()
        Else
            LL("I will display data from Summary Server", 2)
            ender()
        End If
    End Sub
    'HDC Control Number Mirroring
    Private Sub hdcsn_box_TextChanged(sender As Object, e As EventArgs) Handles hdcsn_box.TextChanged
        qID_box.Text = hdcsn_box.Text
    End Sub
    'Query Manual SQL Handler
    Private Sub qmanualsql_B_Click(sender As Object, e As EventArgs) Handles qmanualsql_B.Click
        pgmin()
        err = 0
        If qManualSQL_BOX.Text = "" Then
            LL("Please provide SQL First!", 3)
            pgmax()
            Exit Sub
        End If
        pgval(50)
        'sql = qManualSQL_BOX.Text
        Application.DoEvents()
        Refresh()
        clr_val()

        If QmanualSql(0) > 0 Then
        End If
        Me.qDG.CurrentCell = Me.qDG(0, 0)
        pgmax()
        ender()
        qID_box.Focus()
        qID_box.Select()
    End Sub


    'this will handle the query Procedure
    Private Sub qRead_B_Click(sender As Object, e As EventArgs) Handles qRead_B.Click
        err = 0

        If qID_box.Text = "" Then
            LL("Please Input HDC Control Number!", 3)
            pgmax()
            ender()
            Exit Sub
        End If

        pgmin()
        Dim sqlstr As String = ""
        If Not Qhistory_check.Checked = True Then
            'Form History File
            sqlstr = "Select * FROM ADM_MAIN WHERE HDC_SN = '" & qID_box.Text & "';"
        Else
            'from summary Server
            sqlstr = "Select * FROM ADM_MIRROR WHERE HDC_SN = '" & qID_box.Text & "';"
        End If
        pgval(50)
        qManualSQL_BOX.Text = sqlstr
        LL(sqlstr, 1)
        Application.DoEvents()

        If (QmanualSql(0) > 0) Then
        Else
        End If
        pgmax()
        ender()
        Me.qDG.CurrentCell = Me.qDG(0, 0)
        qID_box.Focus()
        qID_box.Select()

    End Sub

    Private Sub qID_box_KeyDown(sender As Object, e As KeyEventArgs) Handles qID_box.KeyDown
        If e.KeyCode = Keys.Enter Then
            qRead_B.PerformClick()
        End If
    End Sub

    Private Sub qmanualsql_Box_KeyDown(sender As Object, e As KeyEventArgs) Handles qManualSQL_BOX.KeyDown
        If e.KeyCode = Keys.Enter Then
            qmanualsql_B.PerformClick()
        End If
    End Sub

    'Multiple File Read Handler
    Private Sub Qreadmultiple_B_Click(sender As Object, e As EventArgs) Handles Qreadmultiple_B.Click
        pgmin()
        If Qreadmultiple_Box.Text = "" Then
            LL("Please Input Control Number First.", 3)
            pgmax()
            Beep()
            ender()
            Exit Sub
        End If
        If readMul(0) > 0 Then
        Else
        End If
        pgmax()
        ender()
    End Sub

    'this handles personal EQ List
    Private Sub QmyEQ_B_Click(sender As Object, e As EventArgs) Handles QmyEQ_B.Click
        pgmin()
        LL(user_fname.ToString, 1)
        LL(user_lname.ToString, 2)

        LL("Starting Equipment assignment query for " & user_fname, 1)

        qManualSQL_BOX.Text = ""
        Application.DoEvents()
        If Qhistory_check.Checked = True Then
            qManualSQL_BOX.Text = "SELECT * FROM ADM_MIRROR WHERE EQ_PIC like '" & user_fname & " " & user_lname & "' ORDER BY EQ_PIC;"
            LL("Reading from History Database", 0)
        Else
            qManualSQL_BOX.Text = "SELECT * FROM ADM_MAIN WHERE EQ_PIC like '" & user_fname & " " & user_lname & "' ORDER BY EQ_PIC;"
            LL("Reading from Main Database", 0)
        End If

        Application.DoEvents()

        If qManualSQL_BOX.Text = "" Then
            Beep()
            LL("No SQL String Found", 2)
            pgmax()
            ender()
        End If

        If QmanualSql(0) > 0 Then
        Else
            LL("Finished Equipment assignment query for " & user_fname, 2)
        End If
        Me.qDG.CurrentCell = Me.qDG(0, 0)
        pgmax()
        ender()
    End Sub

    'date and Time Picker

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim ndate As New Date
        ndate = DateTimePicker1.Value

        Dim DD As String, MM As String, YY As String
        DD = ndate.Day.ToString
        MM = ndate.Month.ToString
        YY = ndate.Year.ToString

        eqcalD_box.Text = MM & "/" & DD & "/" & YY
        Application.DoEvents()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        'eqcaldue_box
        Dim ndate As New Date
        ndate = DateTimePicker2.Value

        Dim DD As String, MM As String, YY As String
        DD = ndate.Day.ToString
        MM = ndate.Month.ToString
        YY = ndate.Year.ToString

        eqcaldue_box.Text = MM & "/" & DD & "/" & YY
        Application.DoEvents()
    End Sub

    'it will handle manual Query Tabulation
    Private Sub QTabulated_B_Click(sender As Object, e As EventArgs) Handles QTabulated_B.Click
        err = 0
        pgmin()
        LL("Counting Equipment", 1)
        If qManualSQL_BOX.Text = "" Then
            Beep()
            LL("NO SQL String Found!", 3)
            ender()
            pgmax()
        End If

        If QTabulated_box.Text = "" Then
            Beep()
            LL("Please select Tabulation and Grouping first!", 3)
            pgmax()
            ender()
        End If


        LL("Tabulating " & QTabulated_box.Text.ToString, 1)

        qManualSQL_BOX.Text = ""
        qManualSQL_BOX.Text = "SELECT " & QTabulated_box.Text & ", COUNT(HDC_SN) as 'Equipment Count' FROM ADM_MAIN GROUP BY " & QTabulated_box.Text & " ORDER BY " & QTabulated_box.Text

        If QmanualSql(0) > 0 Then
        Else
            LL("Finished Equipment Tabulation Query", 2)
        End If
        Me.qDG.CurrentCell = Me.qDG(0, 0)
        pgmax()
        ender()
    End Sub
    'this will hadle more complicated search!
    Private Sub qDG_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles qDG.CellDoubleClick

        err = 0
        pgmin()
        Dim rowid As Long = qDG.CurrentRow.Index
        Dim cellSTR As String = Trim(qDG.Rows(rowid).Cells(0).Value.ToString)
        Dim cellHDR As String = qDG.Columns(0).HeaderText.ToString

        qManualSQL_BOX.Text = ""
        qManualSQL_BOX.Text = "SELECT * FROM ADM_MAIN WHERE " & cellHDR & " like '" & cellSTR & "' ORDER BY " & cellHDR & ";"
        If QmanualSql(0) > 0 Then
        Else
        End If
        Me.qDG.CurrentCell = Me.qDG(0, 0)
        pgmax()
        ender()
    End Sub
    Public logoutmode As Long = 0
    Private Sub main_tab_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles main_tab.DrawItem

        If logoutmode > 0 Then Exit Sub

        Dim g As Graphics = e.Graphics
        Dim tp As TabPage = main_tab.TabPages(e.Index)
        Dim br As Brush
        Dim sf As New StringFormat
        Dim r As New RectangleF(e.Bounds.X, e.Bounds.Y + 2, e.Bounds.Width, e.Bounds.Height - 2)

        sf.Alignment = StringAlignment.Center

        Dim strTitle As String = tp.Text

        Select Case tp.Text
            Case "ADM"
                'Back Color
                br = New SolidBrush(Color.Blue)
                g.FillRectangle(br, e.Bounds)
                'Font Color
                br = New SolidBrush(Color.White)
                g.DrawString(strTitle, New Font(main_tab.Font, FontStyle.Bold), br, r, sf)
                'main_tab.TabPages(e.Index).Font = New Font(adm_tab.Font, FontStyle.Bold)
            Case "Login"
                'Back Color
                br = New SolidBrush(Color.Green)
                g.FillRectangle(br, e.Bounds)
                'Font Color
                br = New SolidBrush(Color.White)
                g.DrawString(strTitle, main_tab.Font, br, r, sf)
            Case "Ref"
                'Back Color
                br = New SolidBrush(Color.Pink)
                g.FillRectangle(br, e.Bounds)
                'Font Color
                br = New SolidBrush(Color.Black)
                g.DrawString(strTitle, main_tab.Font, br, r, sf)
            Case "Sign Up"
                'Back Color
                br = New SolidBrush(Color.Bisque)
                g.FillRectangle(br, e.Bounds)
                'Font Color
                br = New SolidBrush(Color.Black)
                g.DrawString(strTitle, main_tab.Font, br, r, sf)
            Case "Query"
                'Back Color
                br = New SolidBrush(Color.LightCyan)
                g.FillRectangle(br, e.Bounds)
                'Font Color
                br = New SolidBrush(Color.Black)
                g.DrawString(strTitle, main_tab.Font, br, r, sf)
            Case Else
                'Back Color
                br = New SolidBrush(Color.White)
                g.FillRectangle(br, e.Bounds)
                'Font Color
                br = New SolidBrush(Color.Black)
                g.DrawString(strTitle, main_tab.Font, br, r, sf)
        End Select

        g = Nothing
        tp = Nothing
        br = Nothing
        sf = Nothing
        r = Nothing

    End Sub



    Private Sub sql_box_KeyDown(sender As Object, e As KeyEventArgs) Handles sql_box.KeyDown
        If e.KeyCode = Keys.Enter Then
            sql_B.PerformClick()
        End If
    End Sub


End Class
