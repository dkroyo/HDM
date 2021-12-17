
Imports System
Imports System.IO
Imports System.IO.File
Imports HDC_ADM.ADM
Imports MySql.Data.MySqlClient
'Imports MySql.Data.MySql



Module ADM_COMMON
    '#server connection check function
    Public Function server_connect(ByVal res As Integer)
        ADM.PGB.Value = ADM.PGB.Minimum
        update_dir = update_dir
        err = 0
        Try
            If Directory.Exists(update_dir) Then
                LL("Connected to HDC Common Server!", 2)
            Else
                LL("Unable to connect to \\10.193.12.174, trying to established Connection!", 3)
                'trial 1 to connect
                Dim P As New Process
                P.StartInfo.FileName = "net.exe"
                P.StartInfo.Arguments = "use K: \\10.193.12.174\ /USER:toshiba\japsup 1048"
                P.StartInfo.CreateNoWindow = True
                P.Start()
                P.WaitForExit(2500)
                'Process.Start("CMD.exe", "/c net use \\10.193.12.174 /USER:toshiba\japsup 1048")
                'LL("CMD.exe /c net use \\10.193.12.174 /USER:toshiba\japsup 1048", 0)
                're check connection
                If Directory.Exists(update_dir) Then
                    LL("Connected to HDC Common Server after internal connection access.", 2)
                Else
                    LL("Still unable to connect to the server! Please try accessing manually.", 3)
                    LL("", 0)
                    LL("The App will continue boot but might not work properly!", 1)
                End If
            End If

        Catch ex As Exception
            err = 1
            Beep()
            LL(ex.Message.ToString, 2)
            ADM.PGB.Value = ADM.PGB.Minimum
        End Try
        ADM.PGB.Value = ADM.PGB.Maximum
        ender()
        Return res
    End Function

    '# Logger Function -> This function is basically for updating log on the Display
    Public Function LL(ByRef log_Str As String, ByRef iColor As Integer)
        With ADM
            ADM.PGB.Value = ADM.PGB.Minimum
            'initializing the declared variable
            log_Str = log_Str
            iColor = iColor

            'select Color Condition before Writing anything
            Select Case iColor
                Case 0 'white
                    .log_box.Select(.log_box.TextLength, 0)
                    .log_box.SelectionColor = Color.White
                    .log_box.SelectionFont = New Font("Century Gothic", 8.25, FontStyle.Regular)
                Case 1 'blue
                    .log_box.Select(.log_box.TextLength, 0)
                    .log_box.SelectionColor = Color.Cyan
                    .log_box.SelectionFont = New Font("Century Gothic", 8.25, FontStyle.Regular)
                Case 2 'Yellow
                    .log_box.Select(.log_box.TextLength, 0)
                    .log_box.SelectionColor = Color.Yellow
                    .log_box.SelectionFont = New Font("Century Gothic", 8.25, FontStyle.Regular)
                Case 3 'Red
                    .log_box.Select(.log_box.TextLength, 0)
                    .log_box.SelectionColor = Color.Red
                    .log_box.SelectionFont = New Font("Century Gothic", 10, FontStyle.Bold)
                Case Else
                    .log_box.SelectionFont = New Font("Century Gothic", 8.25, FontStyle.Regular)
            End Select
            ADM.PGB.Value = ADM.PGB.Maximum / 2
            'Initialize the components
            'Date, Time 
            Dim LL_datetime As String = Today.Month.ToString & "/" & Today.Day.ToString & "/" & Today.Year.ToString & " " & TimeOfDay.ToString("t")

            log_Str = LL_datetime & ">" & vbTab & log_Str
            ADM.log_box.AppendText(log_Str & vbNewLine)

            'selecting the last row of the logger data
            .log_box.SelectionStart = Len(.log_box.Text)

        End With
        Return log_Str
        ADM.PGB.Value = ADM.PGB.Maximum
        ender()
    End Function

    '#Clearing all values on the form
    Public Sub clr_val(Optional ByVal value As Integer = 0)
        ADM.PGB.Value = ADM.PGB.Minimum
        err = 0
        Try
            With ADM
                If value = 1 Then
                Else
                    .hdcsn_box.Text = ""
                    .fa_box.Text = ""
                    .cal_box.Text = ""
                    .sn_box.Text = ""
                    .exp_box.Text = ""
                End If
                .eqmaker_box.Text = ""
                .eqmodel_box.Text = ""
                .eqtype_box.Text = ""
                .eqdisp_box.Text = ""
                .eqsite_box.Text = ""
                .eqbuild_box.Text = ""
                .eqroom_box.Text = ""
                .eqfloor_box.Text = ""
                .eqloc_box.Text = ""
                .eqpic_box.Text = ""
                .eqhost_box.Text = ""
                .equser_box.Text = ""
                .eqpass_box.Text = ""
                .eqip_box.Text = ""
                .eqos_box.Text = ""
                .eqantv_box.Text = ""
                .eqstatus_box.Text = ""
                .eqcals_box.Text = ""
                .eqcalD_box.Text = ""
                .eqcaldue_box.Text = ""
                .todate_box.Text = ""
                .totime_box.Text = ""
                .eqremarks_box.Text = ""
                '.main_DG.Rows.Clear()
                '.main_DG.Columns.Clear()
                .main_DG.DataSource = ""
            End With
            LL("Form is cleared and initialized!", 0)
        Catch ex As Exception
            Beep()
            err = 1
            LL(ex.Message.ToString, 2)
            ADM.PGB.Value = ADM.PGB.Minimum
            pgmax()
            Exit Sub
        End Try
        ender()
        ADM.PGB.Value = ADM.PGB.Maximum
    End Sub

    'this function will initialized the search parameter!
    Public Function searchCheck(ByVal res As Integer)
        ADM.PGB.Value = ADM.PGB.Minimum
        'this time i will make the defaul sear conditio checked item.
        Try
            err = 0
            search_str = ""
            With ADM
                If .hdcsn_check.Checked = True Then
                    search_str = "HDC_SN"
                    ID = .hdcsn_box.Text
                End If
                If .fa_check.Checked = True Then
                    search_str = "FA_ID"
                    ID = .fa_box.Text
                End If
                If .cal_check.Checked = True Then
                    search_str = "CAL_ID"
                    ID = .cal_box.Text
                End If
                If .exp_check.Checked = True Then
                    search_str = "EXP_ID"
                    ID = .exp_box.Text
                End If
                If .sn_check.Checked = True Then
                    search_str = "SN_ID"
                    ID = .sn_box.Text
                End If
            End With
            ADM.PGB.Value = ADM.PGB.Maximum / 2
            If search_str = "" Then
                LL("Searh Paramter Invalid [Radio selection]", 3)
            End If
            LL("New Search reference: " & search_str, 2)
        Catch ex As Exception
            err = 1
            Beep()
            LL(ex.Message.ToString, 3)
            ADM.PGB.Value = ADM.PGB.Minimum
        End Try
        'ender()
        ADM.PGB.Value = ADM.PGB.Maximum
        Return res
    End Function
    'Text Box Verifier
    Public Function txt_val(ByVal res As Integer)
        ADM.PGB.Value = ADM.PGB.Minimum
        err = 0
        Try
            LL("Verifying required data availability?", 0)
            'boolean Expression

            With ADM
                If .hdcsn_box.Text = "" Or .fa_box.Text = "" Or .cal_box.Text = "" Or .sn_box.Text = "" Or .exp_box.Text = "" Or .eqmaker_box.Text = "" Or .eqmodel_box.Text = "" Or .eqtype_box.Text = "" Or .eqdisp_box.Text = "" Or .eqsite_box.Text = "" Or .eqbuild_box.Text = "" Or .eqroom_box.Text = "" Or .eqfloor_box.Text = "" Or .eqloc_box.Text = "" Or .eqpic_box.Text = "" Or .eqhost_box.Text = "" Or .equser_box.Text = "" Or .eqpass_box.Text = "" Or .eqip_box.Text = "" Or .eqos_box.Text = "" Or .eqantv_box.Text = "" Or .eqstatus_box.Text = "" Or .eqcals_box.Text = "" Or .eqcalD_box.Text = "" Or .eqcaldue_box.Text = "" Or .todate_box.Text = "" Or .totime_box.Text = "" Or .eqremarks_box.Text = "" Then
                    err = 1
                    res = 1
                    Beep()
                    LL("Please Fill all required values and do not leave blank!", 3)
                Else
                    res = 0
                End If
            End With
        Catch ex As Exception
            err = 1
            Beep()
            LL(ex.Message.ToString, 3)
            ADM.PGB.Value = ADM.PGB.Minimum
        End Try
        ender()
        ADM.PGB.Value = ADM.PGB.Maximum
        Return res

    End Function
    'selecting End of the Log
    Public Sub ender()
        Application.DoEvents()
        ADM.log_box.SelectionStart = Len(ADM.log_box.Text)
        ADM.log_box.Select()
        ADM.log_box.Focus()
        Application.DoEvents()
        ADM.hdcsn_box.Select()
    End Sub

    'date and time function Starndard Format!
    'Today.Month.ToString & "/" & Today.Day.ToString & "/" & Today.Year.ToString & " " & TimeOfDay.ToString("t")
    Public Function MyDataTime(ByVal res As Integer)

        ADM.PGB.Value = ADM.PGB.Minimum
        err = 0
        aDate = aDate
        atime = atime
        Try
            res = 0
            Dim DD As String, MM As String, YY As String, TT As String
            DD = Today.Day.ToString
            MM = Today.Month.ToString
            YY = Today.Year.ToString
            TT = TimeOfDay.ToString("t")
            aDate = MM & "/" & DD & "/" & YY
            atime = TT
            ADM.PGB.Value = ADM.PGB.Maximum / 2
            If aDate <> "" Or atime <> "" Then
                With ADM
                    .todate_box.Text = aDate
                    .totime_box.Text = atime
                End With
                Application.DoEvents()
                LL("Date and Time values are updated! [ " & aDate & " " & atime & "]", 2)
            Else
                err = 1
                res = 1
                LL("Failed to get current Date and Time!", 3)
            End If

            If ADM.hdcsn_box.Text = "" Then
            Else
                If ADM.eqcalD_box.Text = "" Then
                    ADM.eqcalD_box.Text = aDate
                Else
                End If

                If ADM.eqcaldue_box.Text = "" Then
                    ADM.eqcaldue_box.Text = aDate
                    'LL("Set Due Date is Today! Please Check again!", 3)
                Else
                End If

            End If


        Catch ex As Exception
            res = 1
            Beep()
            err = 1
            LL(ex.Message.ToString, 3)
            ADM.PGB.Value = ADM.PGB.Minimum
        End Try
        ender()
        ADM.PGB.Value = ADM.PGB.Maximum
        Return res
    End Function
    'SET VALUES
    Public Function set_val(ByVal res As Integer)
        ADM.PGB.Value = ADM.PGB.Minimum
        err = 0
        Try
            res = 0
            With ADM
                .hdcsn_box.Text = hdcsn
                .fa_box.Text = fa
                .cal_box.Text = cal
                .sn_box.Text = sn
                .exp_box.Text = exp
                .eqmaker_box.Text = eqmaker
                .eqmodel_box.Text = eqmodel
                .eqtype_box.Text = eqtype
                .eqdisp_box.Text = eqdisp
                .eqsite_box.Text = eqsite
                .eqbuild_box.Text = eqbuild
                .eqroom_box.Text = eqroom
                .eqfloor_box.Text = eqfloor
                .eqloc_box.Text = eqloc
                .eqpic_box.Text = eqpic
                .eqhost_box.Text = eqhost
                .equser_box.Text = equser
                .eqpass_box.Text = eqpass
                .eqip_box.Text = eqip
                .eqos_box.Text = eqos
                .eqantv_box.Text = eqantv
                .eqstatus_box.Text = eqstatus
                .eqcals_box.Text = eqcals
                .eqcalD_box.Text = eqcalD
                .eqcaldue_box.Text = eqcaldue
                .todate_box.Text = todate
                .totime_box.Text = totime
                .eqremarks_box.Text = eqremarks
            End With
            res = 0
            LL("required values are set!", 1)
        Catch ex As Exception
            res = 1
            Beep()
            err = 1
            LL(ex.Message.ToString, 3)
            ADM.PGB.Value = ADM.PGB.Minimum
        End Try
        ender()
        ADM.PGB.Value = ADM.PGB.Maximum
        Return res
    End Function
    'GET VALUES
    Public Function get_val(ByVal res As Integer)
        ADM.PGB.Value = ADM.PGB.Minimum
        err = 0
        Try
            res = 0
            With ADM
                hdcsn = .hdcsn_box.Text
                fa = .fa_box.Text
                cal = .cal_box.Text
                sn = .sn_box.Text
                exp = .exp_box.Text
                eqmaker = .eqmaker_box.Text
                eqmodel = .eqmodel_box.Text
                eqtype = .eqtype_box.Text
                eqdisp = .eqdisp_box.Text
                eqsite = .eqsite_box.Text
                eqbuild = .eqbuild_box.Text
                eqroom = .eqroom_box.Text
                eqfloor = .eqfloor_box.Text
                eqloc = .eqloc_box.Text
                'eqpic = .eqpic_box.Text
                'this option is added for Special mode 1
                If smode1 > 0 Or smode2 > 0 Then
                    eqpic = .eqpic_box.Text
                Else
                    eqpic = user_fname & " " & user_lname
                End If

                eqhost = .eqhost_box.Text
                equser = .equser_box.Text
                eqpass = .eqpass_box.Text
                eqip = .eqip_box.Text
                eqos = .eqos_box.Text
                eqantv = .eqantv_box.Text
                eqstatus = .eqstatus_box.Text
                eqcals = .eqcals_box.Text
                eqcalD = .eqcalD_box.Text
                eqcaldue = .eqcaldue_box.Text
                todate = .todate_box.Text
                totime = .totime_box.Text
                Dim rex As String() = Split(.eqremarks_box.Text, "|")
                If UBound(rex) < 1 Then
                    eqremarks = .eqremarks_box.Text & "|" & Environment.MachineName.ToString
                Else
                    eqremarks = rex(0) & "|" & rex(1)
                End If
                'this option is added for Special mode 1
                If smode1 > 0 Then
                    Dim a() As String = Split(eqremarks, "#")
                    Dim b As String = ""
                    For i = LBound(a) To UBound(a)
                        Select Case a(i)
                            Case " STATUS: UPDATED! - Auditor Check "
                            Case " STATUS: NOT UPDATED! - Auditor Check "
                            Case " Admin Update "
                            Case Else
                                If a(i) = " " Or a(i) = "" Then
                                Else
                                    b = b & a(i)
                                End If

                        End Select
                    Next

                    eqremarks = b
                    eqremarks = "# Admin Update # " & eqremarks
                Else
                End If

                If smode2 > 0 Then

                    Dim a() As String = Split(eqremarks, "#")
                    Dim b As String = ""
                    For i = LBound(a) To UBound(a)
                        Select Case a(i)
                            Case " STATUS: UPDATED! - Auditor Check "
                            Case " STATUS: NOT UPDATED! - Auditor Check "
                            Case " Admin Update "
                            Case Else
                                If a(i) = " " Or a(i) = "" Then
                                Else
                                    b = b & a(i)
                                End If

                        End Select
                    Next

                    eqremarks = b

                    Dim a_s As String
                    If (MsgBox("Equipment Informations are updated!", vbYesNo) = vbYes) Then
                        a_s = "STATUS: UPDATED!"
                    Else
                        a_s = "STATUS: NOT UPDATED!"
                    End If
                    eqremarks = "# " & a_s & " - Auditor Check # " & eqremarks
                Else
                End If

                If smode1 = 0 And smode2 = 0 Then
                    Dim a() As String = Split(eqremarks, "#")
                    Dim b As String = ""
                    For i = LBound(a) To UBound(a)
                        Select Case a(i)
                            Case " STATUS: UPDATED! - Auditor Check "
                            Case " Admin Update "
                            Case Else
                                If a(i) = " " Or a(i) = "" Then
                                Else
                                    b = b & a(i)
                                End If

                        End Select
                    Next
                    eqremarks = b
                End If
                eqremarks = eqremarks




            End With
            res = 0
            LL("required values are initialized!", 1)
        Catch ex As Exception
            res = 1
            Beep()
            err = 1
            LL(ex.Message.ToString, 3)
            ADM.PGB.Value = ADM.PGB.Minimum
        End Try
        ender()
        ADM.PGB.Value = ADM.PGB.Maximum
        Return res
    End Function
    'minimum PGBA
    Public Function pgmin()
        pgmin = True
        ADM.PGB.Value = ADM.PGB.Minimum
        Return pgmin
    End Function
    'Maximum PG.bar
    Public Function pgmax()
        pgmax = True
        ADM.PGB.Value = ADM.PGB.Maximum
        Return pgmax
    End Function

    Public Function pgval(ByVal Value As Long)
        ADM.PGB.Value = (ADM.PGB.Maximum / 100) * Value
        Return Value
    End Function

    'DUMMY OnLY
    Sub dummy()
        ADM.PGB.Value = ADM.PGB.Minimum
        With ADM
            .hdcsn_box.Text = "-"
            .fa_box.Text = "-"
            .cal_box.Text = "-"
            .sn_box.Text = "-"
            .exp_box.Text = "-"
            .eqmaker_box.Text = "-"
            .eqmodel_box.Text = "-"
            .eqtype_box.Text = "-"
            .eqdisp_box.Text = "-"
            .eqsite_box.Text = "-"
            .eqbuild_box.Text = "-"
            .eqroom_box.Text = "-"
            .eqfloor_box.Text = "-"
            .eqloc_box.Text = "-"
            .eqpic_box.Text = "-"
            .eqhost_box.Text = "-"
            .equser_box.Text = "-"
            .eqpass_box.Text = "-"
            .eqip_box.Text = "-"
            .eqos_box.Text = "-"
            .eqantv_box.Text = "-"
            .eqstatus_box.Text = "-"
            .eqcals_box.Text = "-"
            .eqcalD_box.Text = "-"
            .eqcaldue_box.Text = "-"
            .todate_box.Text = "-"
            .totime_box.Text = "-"
            .eqremarks_box.Text = "-"
        End With
        ADM.PGB.Value = ADM.PGB.Maximum
    End Sub
    '#Reading The database

    Private Sub last_srh()
        search_str = search_str
        ID = ID
        Select Case search_str
            Case "HDC_SN"
                ADM.hdcsn_box.Text = ID
            Case "FA_ID"
                ADM.fa_box.Text = ID
            Case "CAL_ID"
                ADM.cal_box.Text = ID
            Case "EXP_ID"
                ADM.exp_box.Text = ID
            Case "SN_ID"
                ADM.sn_box.Text = ID
        End Select

    End Sub


    Public Function readsqlstr(ByRef res As Integer)
        res = 0
        err = 0
        Dim con As New MySqlConnection()
        Try
            res = 0
            'Writing read sql command
            'If searchCheck(0) > 0 Then
            '    res = 1
            'Else
            res = 0
                'only good Search STR Values
                search_str = search_str
                sql = "SELECT * from " & maindb & " where " & search_str & " = '" & ID & "';"


                Application.DoEvents()
                'this should be the Query

                clr_val() 'need to reset the box for new writing condition
                ADM.sql_box.Text = sql
                LL("Opening Database connection!", 2)
                pgval(10)
                'Opening Server Connection. #MUST BE CLOSED
                con.ConnectionString = cons
                con.Open()

                Dim sqlcmd As New MySqlCommand
                With sqlcmd
                    .CommandText = sql
                    .Connection = con
                End With

                Application.DoEvents()

                Dim rd As MySqlDataReader = sqlcmd.ExecuteReader

                'check if has data
                If rd.Read Then
                    Application.DoEvents()
                    hdcsn = rd("HDC_SN").ToString
                    fa = rd("FA_ID").ToString
                    cal = rd("CAL_ID").ToString
                    sn = rd("SN_ID").ToString
                    exp = rd("EXP_ID").ToString
                    eqmaker = rd("EQ_MAKER").ToString
                    eqmodel = rd("EQ_MODEL").ToString
                    eqtype = rd("EQ_TYPE").ToString
                    eqdisp = rd("EQ_DISP").ToString
                    eqsite = rd("EQ_SITE").ToString
                    eqbuild = rd("EQ_BUILD").ToString
                    eqroom = rd("EQ_ROOM").ToString
                eqfloor = rd("EQ_DEPT").ToString
                eqloc = rd("EQ_LOC").ToString
                    eqpic = rd("EQ_PIC").ToString
                    eqhost = rd("EQ_HOST").ToString
                    equser = rd("EQ_USER").ToString
                    eqpass = rd("EQ_PASS").ToString
                    eqip = rd("EQ_IP").ToString
                    eqos = rd("EQ_OS").ToString
                    eqantv = rd("EQ_ANTV").ToString
                    eqstatus = rd("EQ_STATUS").ToString
                    eqcals = rd("EQ_CALS").ToString
                    eqcalD = rd("EQ_CALD").ToString
                    eqcaldue = rd("EQ_CALDUE").ToString
                    todate = rd("TO_DATE").ToString
                    totime = rd("TO_TIME").ToString
                    eqremarks = rd("EQ_REMARKS").ToString

                    Application.DoEvents()

                    If set_val(0) > 0 Then
                        LL("Error Occured during database data copy to text boxes", 3)
                    End If
                    pgval(55)
                Else
                    Beep()
                    res = 1
                    err = 0
                    LL("No Records Found!", 3)
                    clr_val()
                    last_srh()
                    'ADM.hdcsn_box.Text = ""
                End If
                rd.Close()
                LL("Closing database connection ", 1)
                con.Close()
            'End If
        Catch ex As Exception
            con.Close()
            res = 1
            Beep()
            err = 1
            LL(ex.Message.ToString, 3)
        End Try
        pgval(90)
        ender()
        pgmax()
        Return res
    End Function

    Public Sub cbbox_add(ByRef sqlmarker As String, ByRef sqlstr As String)
        err = 0
        Dim con As New MySqlConnection()
        Try
            LL("staring to initialize the drop down list!", 2)

            con.ConnectionString = cons
            con.Open()
            Dim sqlcmd As New MySqlCommand
            sqlcmd.Connection = con

            sql = "SELECT * from " & refdb & " where  avalue = '" & sqlstr & "' and MARKER = '" & sqlmarker & "';"
            sqlcmd.CommandText = sql
            Application.DoEvents()
            Dim rd As MySqlDataReader = sqlcmd.ExecuteReader
            If rd.Read Then
                Beep()
                err = 1
                LL("Parameter Already Exist!", 3)
                ender()
                pgmax()
                Exit Sub
            End If
            rd.Close()
            sql = "INSERT INTO " & refdb & "(MARKER, aVALUE) VALUES('" & sqlmarker & "', '" & sqlstr & "')"

            sqlcmd.CommandText = sql
            sqlcmd.ExecuteNonQuery()
            Application.DoEvents()
            LL("New Reference Added!", 2)
            ender()
            ADM.CB_list.Text = ""
            ADM.RefreshDropDownListToolStripMenuItem.PerformClick()
        Catch ex As Exception
            con.Close()
            LL(ex.Message.ToString, 3)
            Beep()
            ADM.log_box.Focus()
            pgmax()
            Exit Sub
        End Try
    End Sub

    Public Sub cbbox_read(ByRef sqlstr As String)
        ADM.CB_list.Text = ""
        err = 0
        Dim con As New MySqlConnection()
        Try
            LL("staring to initialize the drop down list!", 2)

            con.ConnectionString = cons
            con.Open()
            Dim sqlcmd As New MySqlCommand
            sqlcmd.Connection = con

            sql = "SELECT aValue from " & refdb & " where  MARKER = '" & sqlstr & "';"

            sqlcmd.CommandText = sql

            Application.DoEvents()
            Dim rd As MySqlDataReader = sqlcmd.ExecuteReader

            Do While rd.Read
                ADM.CB_list.AppendText(rd.GetString(0) & vbNewLine)
            Loop
            rd.Close()
            con.Close()
            ender()
            LL("Done Reading List", 1)
        Catch ex As Exception
            err = 1
            Beep()
            LL(ex.Message.ToString, 3)
            pgmax()
            Exit Sub
        End Try
    End Sub

    Function mailup(ByRef your_email As String, ByRef boss_mail As String, ByVal res As Long)
        your_email = your_email
        boss_mail = boss_mail
        res = res
        err = 0
        Try
            'just initialized sending information option
            memail = memail
            bossmail = bossmail

            ' now write the string update.
            Dim str_mail_body As String = "HDC Control Number: " & hdcsn & vbNewLine &
                                            "Fixed Asset ID: " & fa & vbNewLine &
                                            "Calibration ID: " & cal & vbNewLine &
                                            "Serial Number: " & sn & vbNewLine &
                                            "Expense Item ID: " & exp & vbNewLine &
                                            "Equipment Maker: " & eqmaker & vbNewLine &
                                            "Equipment Model: " & eqmodel & vbNewLine &
                                            "Equipment/Asset Type: " & eqtype & vbNewLine &
                                            "Description: " & eqdisp & vbNewLine &
                                            "Site: " & eqsite & vbNewLine &
                                            "Building: " & eqbuild & vbNewLine &
                                            "Room: " & eqroom & vbNewLine &
                                            "Floor: " & eqfloor & vbNewLine &
                                            "Actual Location: " & eqloc & vbNewLine &
                                            "PIC: " & eqpic & vbNewLine &
                                            "Hostname: " & eqhost & vbNewLine &
                                            "Username: " & equser & vbNewLine &
                                            "IP Address: " & eqip & vbNewLine &
                                            "Operating System: " & eqos & vbNewLine &
                                            "Antivirus Pattern: " & eqantv & vbNewLine &
                                            "Equipment/Asset Status: " & eqstatus & vbNewLine &
                                            "Calibration Status: " & eqcals & vbNewLine &
                                            "Calibration Date: " & eqcalD & vbNewLine &
                                            "Calibration Due Date: " & eqcaldue & vbNewLine &
                                            "Date: " & todate & vbNewLine &
                                            "Time: " & totime & vbNewLine &
                                            "Remarks: " & eqremarks & vbNewLine & vbNewLine & vbNewLine &
                                            "This is system generated Equipment information" & vbNewLine &
                                            "by; " & user_fname & " " & user_lname & vbNewLine &
                Environment.MachineName.ToString & " | " & Environment.UserName.ToString & " | " & Today.ToLongDateString

            Dim str_mail_subject As String = "HDC Equipment Information Update: " & hdcsn

            Application.DoEvents()
            Dim OutApp As Object
            OutApp = CreateObject("Outlook.Application")

            If Not OutApp Is Nothing Then
                LL("Outlook Version: " & OutApp.Version, 0)
            Else
                res = 0
                LL("No Outlook installed on the machine!", 3)
                LL("No Mail Notification, other process is run as usual!", 2)
                Return your_email
                Return bossmail
                Exit Function
            End If

            Dim OutMail As Object
            OutMail = OutApp.CreateItem(0)

            With OutMail
                .To = your_email
                .CC = boss_mail
                .BCC = "dankenneth.royo@toshiba.co.jp"
                .subject = str_mail_subject
                .body = str_mail_body
                .send
            End With

            OutMail = Nothing
            OutApp = Nothing
        Catch ex As Exception
            err = 1
            res = 1
            LL(ex.Message.ToString, 3)
            ender()
        End Try
        ender()
        Return res
        Return your_email
        Return bossmail
    End Function

    'this is for Forgot password Function
    Public Function F_Passowd(ByVal res As Long)
        res = 0
        err = 0
        Try
            'this function handles forgot Password!
            If ADM.user_box.Text = "" Then
                Beep()
                LL("Please Do not Leave username Blank!If you also don't remember your user name please input 'help'.", 3)
                Return res
            End If

            Dim F_myusername As String

            If ADM.user_box.Text = "help" Then
                F_myusername = "help"
            Else
                F_myusername = ADM.user_box.Text
            End If

            F_myusername = F_myusername
            LL("Initializing Forgot Password Function", 1)
            Application.DoEvents()
            Dim OutApp As Object
            OutApp = CreateObject("Outlook.Application")

            If Not OutApp Is Nothing Then
                LL("Outlook Version: " & OutApp.Version, 0)
            Else
                res = 0
                LL("No Outlook installed on the machine!", 3)
                Return res
            End If

            Dim OutMail As Object
            OutMail = OutApp.CreateItem(0)

            With OutMail
                .To = "dankenneth.royo@toshiba.co.jp"
                .CC = "dankenneth.royo@toshiba.co.jp"
                .BCC = ""
                .subject = "HDC-ADM Forgot Password"
                .body = "For Password reset! username: " & F_myusername
                .send
            End With

            OutMail = Nothing
            OutApp = Nothing
            LL("Forgot Password Notification was sent to the administrator!", 0)
        Catch ex As Exception
            LL(ex.Message.ToString, 3)
            err = 1
            res = 1
        End Try
        ender()
        Return res
    End Function

    'initializing the values of the comboboxes
    Public Function cbox_ini(ByVal res As Integer)
        pgmin()
        res = 0
        err = 0
        Dim con As New MySqlConnection()
        Try
            LL("staring to initialize the drop down list!", 2)
            pgval(10)

            con.ConnectionString = cons
            con.Open()
            Dim sqlcmd As New MySqlCommand
            sqlcmd.Connection = con

            'SITE
            '======================================================================
            sql = "SELECT aValue from " & refdb & " where  MARKER = 'SITE' ORDER BY aVALUE ASC;" '-----
            ADM.sql_box.Text = sql
            Application.DoEvents()
            sqlcmd.CommandText = sql

            Application.DoEvents()
            Dim rd As MySqlDataReader = sqlcmd.ExecuteReader
            LL("reading parameters for Site;", 0)                            '-----
            'clear the values before writing
            With ADM.eqsite_box.Items                                        '-----
                .Clear()
                Do While rd.Read
                    .Add(rd.GetString(0))
                Loop
                .Add("<Other>")
                .Add("")
            End With
            LL("initial SITE values are set.", 0)                        '-----
            '======================================================================
            rd.Close()
            'TYPE
            '======================================================================
            sql = "SELECT aValue from " & refdb & " where  MARKER = 'TYPE' ORDER BY aVALUE ASC;" '-----
            ADM.sql_box.Text = sql
            Application.DoEvents()
            sqlcmd.CommandText = sql

            Application.DoEvents()
            rd = sqlcmd.ExecuteReader
            LL("reading parameters for Asset Type;", 0)                      '-----
            'clear the values before writing
            With ADM.eqtype_box.Items                                        '-----
                .Clear()
                Do While rd.Read
                    .Add(rd.GetString(0))
                Loop
                .Add("<Other>")
                .Add("")
            End With
            LL("initial Asset Type values are set.", 0)                        '-----
            '======================================================================
            rd.Close()
            'BUILD
            '======================================================================
            sql = "SELECT aValue from " & refdb & " where  MARKER = 'BUILD' ORDER BY aVALUE ASC;" '-----
            ADM.sql_box.Text = sql
            Application.DoEvents()
            sqlcmd.CommandText = sql

            Application.DoEvents()
            rd = sqlcmd.ExecuteReader
            LL("reading parameters for Building;", 0)                      '-----
            'clear the values before writing
            With ADM.eqbuild_box.Items                                        '-----
                .Clear()
                Do While rd.Read
                    .Add(rd.GetString(0))
                Loop
                .Add("<Other>")
                .Add("")
            End With
            LL("initial Building values are set.", 0)                        '-----
            '======================================================================
            rd.Close()
            'DEPT
            '======================================================================
            sql = "SELECT aValue from " & refdb & " where  MARKER = 'DEPT' ORDER BY aVALUE ASC;" '-----
            ADM.sql_box.Text = sql
            Application.DoEvents()
            sqlcmd.CommandText = sql

            Application.DoEvents()
            rd = sqlcmd.ExecuteReader
            LL("reading parameters for Department/Section;", 0)                      '-----
            'clear the values before writing
            With ADM.eqfloor_box.Items                                        '-----
                .Clear()
                Do While rd.Read
                    .Add(rd.GetString(0))
                Loop
                .Add("<Other>")
                .Add("")
            End With
            LL("initial Department/Section values are set.", 0)                        '-----
            '======================================================================
            rd.Close()

            'Asset Status
            '======================================================================
            sql = "SELECT aValue from " & refdb & " where  MARKER = 'ASTATUS' ORDER BY aVALUE ASC;" '-----
            ADM.sql_box.Text = sql
            Application.DoEvents()
            sqlcmd.CommandText = sql

            Application.DoEvents()
            rd = sqlcmd.ExecuteReader
            LL("reading parameters for Asset STATUS;", 0)                      '-----
            'clear the values before writing
            With ADM.eqstatus_box.Items                                        '-----
                .Clear()
                Do While rd.Read
                    .Add(rd.GetString(0))
                Loop
                .Add("<Other>")
                .Add("")
            End With
            LL("initial Asset Status values are set.", 0)                        '-----
            '======================================================================
            rd.Close()
            ' Calibration
            '======================================================================
            sql = "SELECT aValue from " & refdb & " where  MARKER = 'CSTATUS' ORDER BY aVALUE ASC;" '-----
            ADM.sql_box.Text = sql
            Application.DoEvents()
            sqlcmd.CommandText = sql

            Application.DoEvents()
            rd = sqlcmd.ExecuteReader
            LL("reading parameters for Calibration/PM STATUS;", 0)                      '-----
            'clear the values before writing
            With ADM.eqcals_box.Items                                        '-----
                .Clear()
                Do While rd.Read
                    .Add(rd.GetString(0))
                Loop
                .Add("<Other>")
                .Add("")
            End With
            LL("initial Calibration/PM Status values are set.", 0)                        '-----
            '======================================================================
            rd.Close()
            'USER INFORMATION!
            With ADM.alevel_box.Items
                .Clear()
                .Add(1)
                .Add(2)
                .Add(3)
                .Add(4)
                .Add(5)
            End With

            With ADM.aposition_box.Items
                .Clear()
                .Add("Operator")
                .Add("Technician")
                .Add("Engineer")
                .Add("Staff")
                .Add("Supervisor")
                .Add("Manager")
                .Add("Others")
            End With


            With ADM.bpos_box.Items
                .Clear()
                .Add("Operator")
                .Add("Technician")
                .Add("Engineer")
                .Add("Staff")
                .Add("Supervisor")
                .Add("Manager")
                .Add("Others")
            End With
            With ADM.Qmarker_box.Items
                .Clear()
                .Add("HDC Control Number")
                .Add("Fixed Asset ID")
                .Add("Calibration ID")
                .Add("Expense ID")
                .Add("Serial Number")
            End With

            With ADM.QTabulated_box.Items
                .Clear()
                .Add("EQ_MAKER")
                .Add("EQ_MODEL")
                .Add("EQ_TYPE")
                .Add("EQ_SITE")
                .Add("EQ_BUILD")
                .Add("EQ_ROOM")
                .Add("EQ_DEPT")
                .Add("EQ_LOC")
                .Add("EQ_PIC")
                .Add("EQ_OS")
                .Add("EQ_ANTV")
                .Add("EQ_STATUS")
                .Add("EQ_CALS")
                .Add("EQ_TYPE, EQ_MODEL")
                .Add("EQ_TYPE, EQ_MAKER")
                .Add("EQ_TYPE, EQ_MAKER, EQ_MODEL")
                .Add("EQ_SITE, EQ_BUILD")
                .Add("EQ_SITE, EQ_BUILD, EQ_ROOM")
                .Add("EQ_SITE, EQ_DEPT")
                .Add("EQ_SITE, EQ_DEPT, EQ_PIC")
                .Add("EQ_DEPT, EQ_STATUS")
                .Add("EQ_DEPT, EQ_TYPE")
                .Add("EQ_DEPT, EQ_SITE, EQ_BUILD, EQ_TYPE")
                .Add("EQ_DEPT, EQ_SITE, EQ_BUILD, EQ_ROOM, EQ_TYPE")
                .Add("EQ_STATUS, EQ_TYPE")
                .Add("EQ_STATUS, EQ_DEPT")
                .Add("EQ_STATUS, EQ_SITE")
                .Add("EQ_STATUS, EQ_SITE, EQ_BUILD")
            End With
            ADM.QTabulated_box.Text = "EQ_TYPE"




            clr_val()
            '---------------------------------------
            pgval(90)
            con.Close()

        Catch ex As Exception
            con.Close()
            err = 1
            res = 1
            Beep()
            LL(ex.Message.ToString, 3)
        End Try
        pgmax()
        ender()
        Return res
    End Function

    'Manual SQL
    Function manual_sql(ByVal res As Integer)
        res = 0
        err = 0
        Dim con As New MySqlConnection()
        Try
            'sqlstr
            sql = sql
            sql = ADM.sql_box.Text
            user_level = user_level

            Application.DoEvents()

            Dim txt As String = sql.ToUpper

            'admin page!
            If txt.Contains("ADM_USER") Then
                If user_level < 5 Then
                    LL("Authentication error, security clearance level Is Not enough!", 3)
                    LL("It Requires at least Level 5", 1)
                    ender()
                    res = 1
                    Return res
                    Exit Function
                End If
            End If

            If txt.Contains("INSERT INTO") Then
                If user_level < 3 Then
                    LL("Authentication error, security clearance level Is Not enough!", 3)
                    LL("It Requires at least Level 3", 1)
                    ender()
                    res = 1
                    Return res
                    Exit Function
                End If
            End If

            If txt.Contains("UPDATE") Then
                If user_level < 2 Then
                    LL("Authentication error, security clearance level Is Not enough!", 3)
                    LL("It Requires at least Level 2", 1)
                    ender()
                    res = 1
                    Return res
                    Exit Function
                End If
            End If


            If txt.Contains("DELETE") Or txt.Contains("DROP") Then
                If user_level < 2 Then
                    LL("Authentication error, security clearance level Is Not enough!", 3)
                    LL("It Requires at least Level 9999", 1)
                    ender()
                    res = 1
                    Return res
                    Exit Function
                End If
            End If

            '===============================================================
            con.ConnectionString = cons
            con.Open()
            Application.DoEvents()

            Dim ADP As New MySqlDataAdapter(sql, con)
            Dim ds As New DataSet

            ADP.Fill(ds)

            ADM.main_DG.DataSource = ds.Tables(0)

            ADM.main_DG.AutoResizeColumns()

            con.Close()
            LL("Finished Manual SQL", 2)
            ender()
        Catch ex As Exception
            err = 1
            res = 1
            con.Close()
            Beep()
            LL(ex.Message.ToString, 3)
            ender()
        End Try
        Return res
    End Function

    Function QmanualSql(ByVal res As Long)
        err = 0
        res = 0
        pgmin()
        Application.DoEvents()

        Dim con As New MySqlConnection()
        Try
            'sqlstr
            Dim Qsql As String = ""
            Qsql = Qsql
            Qsql = ADM.qManualSQL_BOX.Text
            user_level = user_level
            ADM.qManualSQL_BOX.Text = Qsql.ToUpper()
            Application.DoEvents()

            Dim txt As String = Qsql.ToUpper

            'admin page!
            If txt.Contains("ADM_USER") Then
                If user_level < 5 Then
                    LL("Authentication error, security clearance level Is Not enough!", 3)
                    LL("It Requires at least Level 5", 1)
                    ender()
                    res = 1
                    Return res
                    Exit Function
                End If
            End If

            If txt.Contains("INSERT INTO") Then
                If user_level < 3 Then
                    LL("Authentication error, security clearance level Is Not enough!", 3)
                    LL("It Requires at least Level 3", 1)
                    ender()
                    res = 1
                    Return res
                    Exit Function
                End If
            End If

            If txt.Contains("UPDATE") Then
                If user_level < 2 Then
                    LL("Authentication error, security clearance level Is Not enough!", 3)
                    LL("It Requires at least Level 2", 1)
                    ender()
                    res = 1
                    Return res
                    Exit Function
                End If
            End If


            If txt.Contains("DELETE") Or txt.Contains("DROP") Then
                If user_level < 2 Then
                    LL("Authentication error, security clearance level Is Not enough!", 3)
                    LL("It Requires at least Level 9999", 1)
                    ender()
                    res = 1
                    Return res
                    Exit Function
                End If
            End If

            '===============================================================
            con.ConnectionString = cons
            con.Open()
            Application.DoEvents()

            ADM.qDG.DataSource = Nothing
            ADM.qDG.Rows.Clear()
            ADM.qDG.Columns.Clear()
            Application.DoEvents()

            Dim ADP As New MySqlDataAdapter(Qsql, con)
            Dim ds As New DataSet

            ADP.Fill(ds)

            ADM.qDG.DataSource = ds.Tables(0)

            ADM.qDG.AutoResizeColumns()

            con.Close()

            ADM.Refresh()

            For Each Col As DataColumn In ds.Tables(0).Columns
                If Col.ColumnName.ToString = "EQ_STATUS" Then
                    If ADM.qDG.Rows.Count > 0 Then
                        For i As Long = 0 To ADM.qDG.Rows.Count - 2
                            Dim toColor As String = ADM.qDG.Rows(i).Cells("EQ_STATUS").Value.ToString

                            If toColor.Contains("IN ACTIVE") Then
                                toColor = "IN ACTIVE"
                            ElseIf toColor.Contains("ACTIVE") Then
                                toColor = "ACTIVE"
                            ElseIf toColor.Contains("SCRAP") Then
                                toColor = "SCRAP"
                            ElseIf toColor.Contains("TRANSFERED TO OTHER DEPARTMENT") Then
                                toColor = "TTOD"

                            End If

                            Select Case toColor
                                Case "IN ACTIVE"
                                    ADM.qDG.Rows(i).DefaultCellStyle.BackColor = Color.Pink
                                Case "ACTIVE"
                                    ADM.qDG.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
                                Case "SCRAP"
                                    ADM.qDG.Rows(i).DefaultCellStyle.BackColor = Color.Black
                                    ADM.qDG.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                Case "TTOD"
                                    ADM.qDG.Rows(i).DefaultCellStyle.BackColor = Color.Blue
                                    ADM.qDG.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                Case Else

                            End Select

                        Next i
                    Else
                        'Do Nothing
                    End If
                Else
                    'do nothing
                End If
            Next
            'ADM.qDG.CurrentCell = ADM.qDG.Rows(0).Cells(0)

            ds = Nothing

            LL("Finished Manual SQL", 2)
            ender()
        Catch ex As Exception
            err = 1
            res = 1
            con.Close()
            Beep()
            LL(ex.Message.ToString, 3)
            ender()
        End Try

        Return res
    End Function



    'Write SQL
    Function write_sql(ByVal res As Integer)
        res = 0
        err = 0
        Dim con As New MySqlConnection()
        Try
            LL("Staring record update", 1)
            'writing
            sql = sql
            ID = ID

            ID = hdcsn
            sql = "SELECT * FROM " & maindb & " WHERE HDC_SN = '" & ID & "';"

            con.ConnectionString = cons
            con.Open()
            Dim sqlcmd As New MySqlCommand
            sqlcmd.CommandText = sql
            sqlcmd.Connection = con

            Dim rd As MySqlDataReader = sqlcmd.ExecuteReader
            Application.DoEvents()

            If rd.Read Then
                If user_level < 2 Then
                    Beep()
                    LL("Authentication error, security clearance level is not enough!", 3)
                    res = 1
                    Return res
                    Exit Function
                End If
                'security OK
                Application.DoEvents()
                Dim upsql As String = "UPDATE adm_main SET " &
                "FA_ID = '" & fa & "', " &
                "CAL_ID = '" & cal & "', " &
                "SN_ID = '" & sn & "', " &
                "EXP_ID = '" & exp & "', " &
                "EQ_MAKER = '" & eqmaker & "', " &
                "EQ_MODEL = '" & eqmodel & "', " &
                "EQ_TYPE = '" & eqtype & "', " &
                "EQ_DISP = '" & eqdisp & "', " &
                "EQ_SITE = '" & eqsite & "', " &
                "EQ_BUILD = '" & eqbuild & "', " &
                "EQ_ROOM = '" & eqroom & "', " &
                "EQ_DEPT = '" & eqfloor & "', " &
                "EQ_LOC = '" & eqloc & "', " &
                "EQ_PIC = '" & eqpic & "', " &
                "EQ_HOST = '" & eqhost & "', " &
                "EQ_USER = '" & equser & "', " &
                "EQ_PASS = '" & eqpass & "', " &
                "EQ_IP = '" & eqip & "', " &
                "EQ_OS = '" & eqos & "', " &
                "EQ_ANTV = '" & eqantv & "', " &
                "EQ_STATUS = '" & eqstatus & "', " &
                "EQ_CALS = '" & eqcals & "', " &
                "EQ_CALD = '" & eqcalD & "', " &
                "EQ_CALDUE = '" & eqcaldue & "', " &
                "TO_DATE = '" & todate & "', " &
                "TO_TIME = '" & totime & "', " &
                "EQ_REMARKS = '" & eqremarks & "' " &
                "WHERE HDC_SN = '" & hdcsn & "'"

                rd.Close()

                LL("Updating Main", 0)

                Dim sqlupdate As New MySqlCommand
                sqlupdate.Connection = con
                sqlupdate.CommandText = upsql
                sqlupdate.ExecuteNonQuery()
                LL(upsql, 0)
                LL("record is updated on main server!!", 1)

                'mirror server update
                Dim upsqlM As String = "INSERT INTO adm_mirror(" &
                    "HDC_SN, FA_ID, CAL_ID, SN_ID, EXP_ID, EQ_MAKER, EQ_MODEL, EQ_TYPE, EQ_DISP, EQ_SITE, " &
                    "EQ_BUILD, EQ_ROOM, EQ_DEPT, EQ_LOC, EQ_PIC, EQ_HOST, EQ_USER, EQ_PASS, EQ_IP, EQ_OS, " &
                    "EQ_ANTV, EQ_STATUS, EQ_CALS, EQ_CALD, EQ_CALDUE, TO_DATE, TO_TIME, EQ_REMARKS) VALUES('" &
                    hdcsn & "', '" & fa & "', '" & cal & "', '" & sn & "', '" & exp & "', '" & eqmaker & "', '" &
                    eqmodel & "', '" & eqtype & "', '" & eqdisp & "', '" & eqsite & "', '" & eqbuild & "', '" &
                    eqroom & "', '" & eqfloor & "', '" & eqloc & "', '" & eqpic & "', '" & eqhost & "', '" &
                    equser & "', '" & eqpass & "', '" & eqip & "', '" & eqos & "', '" & eqantv & "', '" &
                    eqstatus & "', '" & eqcals & "', '" & eqcalD & "', '" & eqcaldue & "', '" & todate & "', '" &
                    totime & "', '" & eqremarks & "')"
                LL("Updating Mirror", 0)

                Application.DoEvents()

                Dim sqlupdateM As New MySqlCommand
                sqlupdateM.Connection = con
                sqlupdateM.CommandText = upsqlM
                sqlupdateM.ExecuteNonQuery()

                Application.DoEvents()
                LL("record is updated on mirror server!!", 1)
                LL(upsqlM, 0)
            Else
                'No Record on database yet!
                rd.Close()
                'this code is added to you can add new records if not yet on the list
                If user_level < 3 Then
                    Beep()
                    LL("Authentication error, security clearance level is not enough!", 3)
                    res = 1
                    Return res
                    Exit Function
                End If

                If MsgBox("No Records Found! Add as new?", vbYesNo, "Add Records?") = vbYes Then
                Else
                    Beep()
                    LL("Transaction is cancelled!", 3)
                    res = 1
                    Return res
                    ender()
                    Exit Function
                End If

                Dim upsqlM As String = "INSERT INTO adm_mirror(" &
                    "HDC_SN, FA_ID, CAL_ID, SN_ID, EXP_ID, EQ_MAKER, EQ_MODEL, EQ_TYPE, EQ_DISP, EQ_SITE, " &
                    "EQ_BUILD, EQ_ROOM, EQ_DEPT, EQ_LOC, EQ_PIC, EQ_HOST, EQ_USER, EQ_PASS, EQ_IP, EQ_OS, " &
                    "EQ_ANTV, EQ_STATUS, EQ_CALS, EQ_CALD, EQ_CALDUE, TO_DATE, TO_TIME, EQ_REMARKS) VALUES('" &
                    hdcsn & "', '" & fa & "', '" & cal & "', '" & sn & "', '" & exp & "', '" & eqmaker & "', '" &
                    eqmodel & "', '" & eqtype & "', '" & eqdisp & "', '" & eqsite & "', '" & eqbuild & "', '" &
                    eqroom & "', '" & eqfloor & "', '" & eqloc & "', '" & eqpic & "', '" & eqhost & "', '" &
                    equser & "', '" & eqpass & "', '" & eqip & "', '" & eqos & "', '" & eqantv & "', '" &
                    eqstatus & "', '" & eqcals & "', '" & eqcalD & "', '" & eqcaldue & "', '" & todate & "', '" &
                    totime & "', '" & eqremarks & "')"
                LL("Updating Mirror", 0)

                Application.DoEvents()

                Dim sqlupdateM As New MySqlCommand
                sqlupdateM.Connection = con
                sqlupdateM.CommandText = upsqlM
                sqlupdateM.ExecuteNonQuery()

                Dim upsql As String = "INSERT INTO adm_main(" &
                    "HDC_SN, FA_ID, CAL_ID, SN_ID, EXP_ID, EQ_MAKER, EQ_MODEL, EQ_TYPE, EQ_DISP, EQ_SITE, " &
                    "EQ_BUILD, EQ_ROOM, EQ_DEPT, EQ_LOC, EQ_PIC, EQ_HOST, EQ_USER, EQ_PASS, EQ_IP, EQ_OS, " &
                    "EQ_ANTV, EQ_STATUS, EQ_CALS, EQ_CALD, EQ_CALDUE, TO_DATE, TO_TIME, EQ_REMARKS) VALUES('" &
                    hdcsn & "', '" & fa & "', '" & cal & "', '" & sn & "', '" & exp & "', '" & eqmaker & "', '" &
                    eqmodel & "', '" & eqtype & "', '" & eqdisp & "', '" & eqsite & "', '" & eqbuild & "', '" &
                    eqroom & "', '" & eqfloor & "', '" & eqloc & "', '" & eqpic & "', '" & eqhost & "', '" &
                    equser & "', '" & eqpass & "', '" & eqip & "', '" & eqos & "', '" & eqantv & "', '" &
                    eqstatus & "', '" & eqcals & "', '" & eqcalD & "', '" & eqcaldue & "', '" & todate & "', '" &
                    totime & "', '" & eqremarks & "')"
                LL("Updating Mirror", 0)

                Application.DoEvents()

                sqlupdateM.Connection = con
                sqlupdateM.CommandText = upsql
                sqlupdateM.ExecuteNonQuery()

                Application.DoEvents()

            End If
            ADM.hdcsn_box.Text = hdcsn
            con.Close()
            ender()
        Catch ex As Exception
            con.Close()
            LL(ex.Message.ToString(), 3)
            res = 1
            err = 1
            ender()
        End Try
        Return res
    End Function

    'this function is for  USER Management
    Function auser_up(ByVal res As Long)
        err = 0
        res = 0
        pgmin()
        Dim con As New MySqlConnection()
        Try
            LL("Initializing User Information Update!", 0)
            'check if the textbox reference is OK
            Dim aID As String = ADM.auser_box.Text

            If aID = "" Then
                res = 1
                ender()
                pgmax()
                LL("Please input user ID to proceed!", 3)
                con.Close()
                Return res
                Exit Function
            End If
            'With ADM
            '    Dim auser As String = .auser_box.Text, apass As String = .apass_box.Text, afname As String = .afname_box.Text,
            '    alname As String = .alname_box.Text, aposition As String = .aposition_box.Text, amymail As String = .amymail_box.Text,
            '    abossmail As String = .abossmail_box.Text, alevel As Integer = .alevel_box.Text

            '    If auser = "" Or apass = "" Or afname = "" Or alname = "" Or aposition = "" Or amymail = "" Or abossmail = "" Or alevel = 0 Then
            '        res = 1
            '        Return res
            '        ender()
            '        pgmax()
            '        LL("Please Provide all the needed information!", 3)
            '        Exit Function
            '    Else
            '    End If
            'End With
            Application.DoEvents()
            'Clear the Textbox Values

            ADM.apass_box.Text = ""
            ADM.afname_box.Text = ""
            ADM.alname_box.Text = ""
            ADM.aposition_box.Text = ""
            ADM.amymail_box.Text = ""
            ADM.abossmail_box.Text = ""
            ADM.alevel_box.Text = ""

            Dim asql As String = "SELECT * from " & userdb & " where USERNAME = '" & aID & "';"

            con.ConnectionString = cons
            con.Open()
            Dim sqlcmd As New MySqlCommand
            sqlcmd.CommandText = asql
            sqlcmd.Connection = con

            Application.DoEvents()
            Dim rd As MySqlDataReader = sqlcmd.ExecuteReader
            If rd.Read Then
                ADM.auser_box.Text = rd("USERNAME").ToString
                ADM.apass_box.Text = rd("PASSWORD").ToString
                ADM.afname_box.Text = rd("FIRST_NAME").ToString
                ADM.alname_box.Text = rd("LAST_NAME").ToString
                ADM.aposition_box.Text = rd("POSITIONx").ToString
                ADM.amymail_box.Text = rd("MyEMAIL").ToString
                ADM.abossmail_box.Text = rd("SupEMAIL").ToString
                ADM.alevel_box.Text = rd("CLEARANCE").ToString
                rd.Close()
            Else
                res = 1

                ender()
                pgmax()
                LL("No Records Found!", 3)
                con.Close()
                rd.Close()
                Return res
                Exit Function
            End If
            LL("Finished Reading User Information", 1)
            con.Close()
        Catch ex As Exception
            LL(ex.Message.ToString, 3)
            Beep()
            pgmax()
            ender()
            con.Close()
        End Try
        pgmax()
        ender()
        Return res
    End Function


    Function aupdateup(ByVal res As Long)
        res = 0
        err = 0
        pgmin()
        Dim con As New MySqlConnection()
        Try
            LL("Initializing User Information for Update", 2)

            LL("Initializing User Information Update!", 0)
            'check if the textbox reference is OK
            Dim aID As String = ADM.auser_box.Text

            If aID = "" Then
                res = 1
                ender()
                pgmax()
                LL("Please input user ID to proceed!", 3)
                con.Close()
                Return res
                Exit Function
            End If

            Dim auser As String = ADM.auser_box.Text
            Dim apass As String = ADM.apass_box.Text
            Dim afname As String = ADM.afname_box.Text
            Dim alname As String = ADM.alname_box.Text
            Dim aposition As String = ADM.aposition_box.Text
            Dim amymail As String = ADM.amymail_box.Text
            Dim abossmail As String = ADM.abossmail_box.Text
            Dim alevel As String = ADM.alevel_box.Text

            If auser = "" Or apass = "" Or afname = "" Or alname = "" Or aposition = "" Or amymail = "" Or abossmail = "" Or alevel = "" Then
                res = 1
                ender()
                pgmax()
                LL("Please Provide all the needed information!", 3)
                con.Close()
                Return res
                Exit Function
            Else
            End If
            Application.DoEvents()

            Dim asql As String = "SELECT * from " & userdb & " where USERNAME = '" & aID & "';"

            con.ConnectionString = cons
            con.Open()
            Dim sqlcmd As New MySqlCommand
            sqlcmd.CommandText = asql
            sqlcmd.Connection = con

            Dim aupsql As String = ""

            Application.DoEvents()
            Dim rd As MySqlDataReader = sqlcmd.ExecuteReader
            If rd.Read Then
                aupsql = "UPDATE " & userdb & " SET " &
                    "PASSWORD = '" & apass & "', " &
                "FIRST_NAME = '" & afname & "', " &
                "LAST_NAME = '" & alname & "', " &
                "POSITIONx = '" & aposition & "', " &
                "MyEMAIL = '" & amymail & "', " &
                "SupEMAIL = '" & abossmail & "', " &
                "CLEARANCE = '" & alevel & "' " &
                "WHERE USERNAME = '" & aID & "'; "
                rd.Close()

                Dim sqlupdate As New MySqlCommand
                sqlupdate.Connection = con
                sqlupdate.CommandText = aupsql
                LL(aupsql, 1)
                sqlupdate.ExecuteNonQuery()

                LL("Finished Updating User Information", 2)
            Else
                res = 1
                ender()
                pgmax()
                LL("Record dont Exist Yet!", 3)
                con.Close()
                Return res
                Exit Function
            End If
            rd.Close()

            ADM.apass_box.Text = ""
            ADM.afname_box.Text = ""
            ADM.alname_box.Text = ""
            ADM.aposition_box.Text = ""
            ADM.amymail_box.Text = ""
            ADM.abossmail_box.Text = ""
            ADM.alevel_box.Text = ""
            con.Close()
        Catch ex As Exception
            LL(ex.Message.ToString, 3)
        End Try
        ender()
        pgmax()
        Return res
    End Function

    'this function is for sign Up Condition
    Function signup(ByRef res As Long)
        res = 0
        err = 0
        pgmin()
        Dim con As New MySqlConnection()
        Try
            LL("Preparing Sign Up Function", 1)
            With ADM

                Application.DoEvents()
                .blevel_box.Text = "2"
                Application.DoEvents()
                .Refresh()
                Dim b_bol As Boolean
                b_bol = .buser_box.Text <> "" And .bpass_box.Text <> "" And .bfname_box.Text <> "" And .blname_box.Text <> "" And .bpos_box.Text <> "" And
                .bemail_box.Text <> "" And .bbossmail_box.Text <> "" And .blevel_box.Text <> ""

                If b_bol = False Then
                    con.Close()
                    Beep()
                    LL("Please Fill all required Information", 3)
                    res = 1
                    err = 1
                    pgmax()
                    ender()
                    Return res
                End If


                If .bemail_box.Text.Contains("@") Then
                Else
                    con.Close()
                    Beep()
                    LL("Invalid Email Address! Please try again.", 3)
                    res = 1
                    err = 1
                    pgmax()
                    ender()
                    Return res
                End If



                If (MsgBox("Are you sure on signing up?", vbYesNo) = vbYes) Then
                    ADM.Refresh()
                    LL("Sign up process was aborted!", 2)
                Else
                    res = 1
                    Return res
                End If
                con.ConnectionString = cons
                con.Open()

                Dim mailsql As New MySqlCommand
                mailsql.CommandText = "select * from adm_user where MyEMAIL = '" & .bemail_box.Text & "';"
                mailsql.Connection = con


                Dim rd As MySqlDataReader = mailsql.ExecuteReader
                Application.DoEvents()

                If rd.Read Then
                    rd.Close()
                    con.Close()
                    Beep()
                    LL("It seems that the email address you are trying to use already exist! Please contact Administrators for detials", 3)
                    res = 1
                    err = 1
                    pgmax()
                    ender()
                    Return res
                End If
                rd.Close()

                Application.DoEvents()
                Dim bsql As String = "INSERT INTO adm_user(USERNAME, PASSWORD, FIRST_NAME, LAST_NAME, POSITIONx, MyEMAIL, SupEMAIL, CLEARANCE) VALUES('" &
                    .buser_box.Text & "', '" & .bpass_box.Text & "', '" & .bfname_box.Text & "', '" & .blname_box.Text & "', '" & .bpos_box.Text & "', '" &
                    .bemail_box.Text & "', '" & .bbossmail_box.Text & "', '" & .blevel_box.Text & "')"

                Dim sqlcmd As New MySqlCommand
                sqlcmd.CommandText = bsql
                sqlcmd.Connection = con
                LL(bsql, 0)
                sqlcmd.ExecuteNonQuery()
                Application.DoEvents()
                con.Close()

                LL("Finished Data Upload!", 1)
                LL("If Error occured after this statement please inform administrator about your signing up manually, especially if you entended higher access level", 0)

                LL("Server Login Creation", 2)
                Application.DoEvents()
                LL("creating Mail Notification", 0)

                Application.DoEvents()
                Dim OutApp As Object
                OutApp = CreateObject("Outlook.Application")

                If Not OutApp Is Nothing Then
                    LL("Outlook Version: " & OutApp.Version, 0)
                Else
                    res = 0
                    LL("No Outlook installed on the machine!", 3)
                    LL("No Mail Notification, other process is run as usual!", 2)
                    Return res
                End If

                Dim OutMail As Object
                OutMail = OutApp.CreateItem(0)

                Dim str_mail_body As String =
                 "Username: " & .buser_box.Text & vbNewLine &
                "Password: " & .bpass_box.Text & vbNewLine &
                "First Name: " & .bfname_box.Text & vbNewLine &
                "Last Name: " & .blname_box.Text & vbNewLine &
                "Position: " & .bpos_box.Text & vbNewLine &
                "Email: " & .bemail_box.Text & vbNewLine &
                "Sup. Mail: " & .bbossmail_box.Text & vbNewLine &
                "Clearance: " & .blevel_box.Text

                With OutMail
                    .To = "dankenneth.royo@toshiba.co.jp"
                    .CC = ADM.bemail_box.Text
                    .BCC = "dankenneth.royo@toshiba.co.jp"
                    .subject = "HDC ADM Sign UP request"
                    .body = str_mail_body
                    .send
                End With

                OutMail = Nothing
                OutApp = Nothing


                LL("Finished Sign Up Process, Please Log In", 2)

                .buser_box.Text = ""
                .bpass_box.Text = ""
                .bfname_box.Text = ""
                .blname_box.Text = ""
                .bpos_box.Text = ""
                .bemail_box.Text = ""
                .bbossmail_box.Text = ""
                .blevel_box.Text = ""
                .Refresh()
                Application.DoEvents()

            End With
            pgmax()
        Catch ex As Exception
            Application.DoEvents()
            con.Close()
            pgmax()
            Beep()
            LL(ex.Message.ToString, 3)
            err = 1
        End Try
        ender()
        pgmax()
        Return res
    End Function

    'multiple File Read Function

    Function readMul(ByVal res As Long)
        err = 0
        res = pgmin()
        Dim con As New MySqlConnection()
        Try

            LL("Starting to read Multiple Control numbers", 0)
            Dim localdir As String = Directory.GetCurrentDirectory & "\"
            Dim localfile As String = localdir & "list.adm"

            Dim qMARKER As String = "HDC_SN"

            Select Case ADM.Qmarker_box.Text
                Case "HDC Control Number"
                    qMARKER = "HDC_SN"
                Case "Fixed Asset ID"
                    qMARKER = "FA_ID"
                Case "Calibration ID"
                    qMARKER = "CAL_ID"
                Case "Expense ID"
                    qMARKER = "EXP_ID"
                Case "Serial Number"
                    qMARKER = "SN_ID"
                Case Else
                    qMARKER = qMARKER
            End Select


            If File.Exists(localfile) Then
                File.Delete(localfile)
            End If
            'Save as File
            Using sw As StreamWriter = File.CreateText(localfile)
                sw.WriteLine(ADM.Qreadmultiple_Box.Text.ToString.ToUpper)
            End Using

            Dim QList As String() = ReadAllLines(localfile)

            Dim qSQL As String = "("

            For i = LBound(QList) To UBound(QList)
                If i = UBound(QList) Then
                    qSQL = qSQL & "'" & QList(i) & "')"
                Else
                    qSQL = qSQL & "'" & QList(i) & "', "
                End If
            Next i

            qSQL = qSQL
            LL(qSQL.ToString, 0)

            Application.DoEvents()

            If ADM.Qhistory_check.Checked = True Then
                qSQL = "Select * from ADM_MIRROR WHERE " & qMARKER & " in " & qSQL & " ORDER BY " & qMARKER
            Else
                qSQL = "Select * from ADM_MAIN WHERE " & qMARKER & " in " & qSQL & " ORDER BY " & qMARKER
            End If

            qSQL = qSQL

            ADM.qManualSQL_BOX.Text = qSQL.ToUpper
            Application.DoEvents()
            ADM.Refresh()

            ADM.qmanualsql_B.PerformClick()
            pgmax()
            ender()
            If File.Exists(localfile) Then
                File.Delete(localfile)
            End If
            Return res
        Catch ex As Exception
            err = 1
            res = 1
            Beep()
            LL(ex.Message.ToString, 3)
            pgmax()
            ender()
        End Try
        Return res
    End Function




End Module
