
Imports System.Data.OleDb

Imports Microsoft.Office.Interop

Imports System.Globalization

Imports System.IO

'Imports i00SpellCheck


Imports System.Net.Mail


Imports System.Windows

Imports DevExpress.XtraSpellChecker

Imports System.Data.SqlClient

Imports System.Net

Imports System.Net.Security

Imports System.Security.Cryptography.X509Certificates
Imports System.Threading

Public Class QAChatScorecard


    Dim SQL As String
    Dim con As New SqlConnection


    Dim One As Integer
    Dim two As Integer
    Dim three As Integer
    Dim Four As Integer
    Dim Five As Integer
    Dim Six As Integer
    Dim seven As Integer



    Dim totalQA As Integer

    ''Store Call Thread
    Dim StoreCallThread As System.Threading.Thread

    Dim dic_en_US As SpellCheckerOpenOfficeDictionary = New SpellCheckerOpenOfficeDictionary


    Dim Desk = My.Computer.FileSystem.SpecialDirectories.Desktop

    Dim ProgramDateForamt As String = "MM/dd/yyyy"

    Dim ProgramDate As DateTime

    Dim spellcheckDIR As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)

    Dim en_USaffPath = IO.Path.Combine(spellcheckDIR.DirectoryName, "en_US.aff")
    Dim en_USdicPath = IO.Path.Combine(spellcheckDIR.DirectoryName, "en_US.dic")


    Public Sub buttonEnables()


        btnSaveScoreCard.Enabled = True
        btnQaSetup.Enabled = True



        cboAgentName.Enabled = True

        cboSupervisor.Enabled = True


        cboSupervisorbox.Enabled = True

        btnSpellChecker.Enabled = True

        cboContactTypeChat.Enabled = True

        btnSave2.Enabled = True

    End Sub

    Public Sub buttondisables()

        btnSpellChecker.Enabled = False


        btnSaveScoreCard.Enabled = False
        btnQaSetup.Enabled = False


        cboAgentName.Enabled = False

        cboSupervisor.Enabled = False

        cboSupervisorbox.Enabled = False

        cboContactTypeChat.Enabled = False


        btnSave2.Enabled = False

    End Sub













    Private Sub ChatScorecard_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            '  PW.Hide()

            FillAutoFail()

            FillSRtype()

            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")


            cboContactTypeChat.SelectedIndex = -1





            If lblDecider.Text = "1" Then



                'txtRevCom.Visible = False
                'GroupBox10.Visible = False
                'Label75.Visible = False



                btnSaveScoreCard.Visible = True
                btnSave2.Visible = False


                BackgroundWorker3.RunWorkerAsync()



                '     QaSetupMod.connecttemp4()

            ElseIf lblDecider.Text = "2" Then



                btnSaveScoreCard.Visible = False
                btnSave2.Visible = True

                BackgroundWorker4.RunWorkerAsync()


                '    QaSetupMod.connecttemp10()

                cboContactTypeChat.Visible = False




            End If



            ''Spell Check

            SpellChecker2.SpellCheckMode = DevExpress.XtraSpellChecker.SpellCheckMode.AsYouType
            SpellChecker2.ParentContainer = Me
            SpellChecker2.CheckAsYouTypeOptions.CheckControlsInParentContainer = True
            SpellChecker2.SpellCheckMode = SpellCheckMode.AsYouType


            dic_en_US.DictionaryPath = en_USdicPath
            dic_en_US.GrammarPath = en_USaffPath
            dic_en_US.Culture = New CultureInfo("en-US")
            SpellChecker2.Dictionaries.Add(dic_en_US)


            '   SpellCheckLoadTimer.Enabled = True


            'dic_en_US.DictionaryPath = "\\NOAMIND01FIL05\Premier_Support\Qa Application\Dictionary\en_US.dic"
            'dic_en_US.GrammarPath = "\\NOAMIND01FIL05\Premier_Support\Qa Application\Dictionary\en_US.aff"
            'dic_en_US.Culture = New CultureInfo("en-US")
            'SpellChecker2.Dictionaries.Add(dic_en_US)








            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "MM/dd/yyyy"



            Me.WindowState = FormWindowState.Maximized
            Me.ActiveControl = cbo1_1



            Time.Enabled = True

            Control.CheckForIllegalCrossThreadCalls = False


            'Me.EnableControlExtensions()




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Public Sub FillSRtype()

        Try

            QaSetupMod.connecttemp18()

            sqltemp18 = "SELECT * FROM [SRType]"



            Dim cmdtemp As New SqlClient.SqlCommand




            cmdtemp.CommandText = sqltemp18

            cmdtemp.Connection = contemp18



            readertemp18 = cmdtemp.ExecuteReader


            While (readertemp18.Read())


                cboSRType.Items.Add(readertemp18("SRType"))


            End While



            cmdtemp.Dispose()

            contemp18.Close()

            Me.Cursor = Cursors.Hand




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try








    End Sub

    Public Sub QaTotalScore()


        '  Dim strQaScoreTotal As String
        Dim intQascoreTotal As Integer


        Dim int1_1 As Integer = cbo1_1.Text
        Dim int1_2 As Integer = cbo1_2.Text

        Dim int2_1 As Integer = cbo2_1.Text

        Dim int3_1 As Integer = cbo3_1.Text
        Dim int3_2 As Integer = cbo3_2.Text
        Dim int3_3 As Integer = cbo3_3.Text
        Dim int3_4 As Integer = cbo3_4.Text
        Dim int3_5 As Integer = cbo3_5.Text
        Dim int3_6 As Integer = cbo3_6.Text
        Dim int3_7 As Integer = cbo3_7.Text
        Dim int3_8 As Integer = cbo3_8.Text

        Dim int4_1 As Integer = Cbo4_1.Text
        Dim int4_2 As Integer = cbo4_2.Text
        Dim int4_3 As Integer = cbo4_3.Text


        Dim int5_1 As Integer = cbo5_1.Text
        Dim int5_2 As Integer = cbo5_2.Text

        Dim int6_1 As Integer = cbo6_1.Text
        Dim int6_2 As Integer = cbo6_2.Text
        Dim int6_3 As Integer = cbo6_3.Text


        Dim int7_1 As Integer = cbo7_1.Text
        Dim int7_2 As Integer = cbo7_2.Text
        Dim int7_3 As Integer = cbo7_3.Text
        Dim int7_4 As Integer = cbo7_4.Text
        Dim int7_5 As Integer = cbo7_5.Text
        Dim int7_6 As Integer = cbo7_6.Text


        One = int1_1 + int1_2

        two = int2_1
        three = int3_1 + int3_2 + int3_3 + int3_4 + int3_5 + int3_6 + int3_7 + int3_8

        Four = int4_1 + int4_2 + int4_3

        Five = int5_1 + int5_2

        Six = int6_1 + int6_2 + int6_3

        seven = int7_1 + int7_2 + int7_3 + int7_4 + int7_5 + int7_6





        intQascoreTotal = int1_1 + int1_2 + int2_1 + int3_1 + int3_2 + int3_3 + int3_4 + int3_5 + int3_6 + int3_7 + int3_8 + int4_1 + int4_2 + int4_3 + int5_1 + int5_2 + int6_1 + int6_2 + int6_3 + int7_1 + int7_2 + int7_3 + int7_4 + int7_5 + int7_6

        lblQAScore1.Text = intQascoreTotal



    End Sub

    Public Sub TCXscore()

        Dim intTCXscore As Integer
        Dim increase As Integer




        Dim int3_3 As Integer = cbo3_3.Text
        Dim int3_5 As Integer = cbo3_5.Text

        Dim int5_1 As Integer = cbo5_1.Text

        Dim int6_1 As Integer = cbo6_1.Text

        Dim int6_3 As Integer = cbo6_3.Text

        '' lblQaAvg.Text = Format(Val(result.ToString()), "0.00")

        increase = int3_3 + int3_5 + int5_1 + int6_1 + int6_3

        intTCXscore = increase / 30 * 100

        lblTCXscore.Text = Format(Val(intTCXscore.ToString()), "0")
        txtTCXScore.Text = Format(Val(intTCXscore.ToString()), "0")


    End Sub

    Public Sub CSatScore()

        Dim intCSATScore As Double
        Dim Cincrease As Integer


        Dim int1 As Integer = cboCSAT1.Text
        Dim int2 As Integer = cboCSAT2.Text
        Dim int3 As Integer = cboCSAT3.Text
        Dim int4 As Integer = cboCSAT4.Text
        Dim int5 As Integer = cboCSAT5.Text
        Dim int6 As Integer = cboCSAT6.Text

        Cincrease = int1 + int2 + int3 + int4 + int5 + int6

        intCSATScore = Cincrease / 6

        txtCSATScore.Text = intCSATScore.ToString("n")




    End Sub




    Public Sub Fillcombo33()


        Try



            QaSetupMod.connecttemp10()



            sqltemp10 = "SELECT * FROM [Agents] WHERE Supervisor='" & lblQAauditor1.Text & "' "





            Dim cmdtemp As New SqlCommand



            '  cmdtemp.CommandText = sqltemp

            cmdtemp.CommandText = sqltemp10

            cmdtemp.Connection = contemp10





            readertemp10 = cmdtemp.ExecuteReader



            While (readertemp10.Read())



                cboSupervisorbox.Items.Add(readertemp10("AgentName"))




            End While



            cmdtemp.Dispose()


            readertemp10.Close()



        Catch ex As Exception



            MsgBox(ex.Message)


        End Try




    End Sub


    Private Sub btnQaSetup_Click(sender As Object, e As EventArgs) Handles btnQaSetup.Click


        Try

            Me.Cursor = Cursors.WaitCursor


            reset()

            Form2.Clear()

            Form2.cboAgentName.Enabled = True

            Form2.cboContactType.Enabled = True

            Form2.cboSupervisor.Enabled = True



            'Form2.txtSRNumber.Text = txtSR.Text

            'Form2.txtContactID.Text = txtContactID.Text

            'Form2.txtContactName.Text = txtContactName.Text


            'Form2.txtContactEmail.Text = txtContactEmail.Text

            'Form2.txtContactPhone.Text = txtContactPhone.Text


            'Form2.txtAccountNum.Text = txtAccountNum.Text


            'Form2.txtCompany.Text = txtCompany.Text


            'Form2.txtOrderID.Text = txtOrderID.Text

            'Form2.DateTimePicker1.Text = DateTimePicker1.Text


            '' User
            ''Jira

            Form2.Show()

            Me.Hide()


            Me.Cursor = Cursors.Hand






        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub


    Public Sub resetcombo()

        cboAgentName.Items.Clear()

        ' cboAgentName.Text = "Agent Name"



    End Sub

    Public Sub Fillcombo()


        Try



            QaSetupMod.connecttemp4()




            '  sqltemp1 = "SELECT * FROM [Agents] WHERE Supervisor='" & lblQAauditor.Text & "' "


            sqltemp4 = "SELECT * FROM [Supervisor]"


            Dim cmdtemp As New SqlCommand



            '  cmdtemp.CommandText = sqltemp

            cmdtemp.CommandText = sqltemp4

            cmdtemp.Connection = contemp4





            readertemp4 = cmdtemp.ExecuteReader



            While (readertemp4.Read())



                cboSupervisor.Items.Add(readertemp4("FullName"))




            End While




            readertemp4.Close()


            cmdtemp.Dispose()
            MissedWeightsReset()



        Catch ex As Exception



            MsgBox(ex.Message)


        End Try



    End Sub




    Private Sub btnSaveScoreCard_Click(sender As Object, e As EventArgs) Handles btnSaveScoreCard.Click




        Try

            MissedWeightsCalc()

            If txtSR.Text <> "1-" And txtSR.MaskFull = False Then


                MsgBox("Please enter a valid SR#")


            Else




                If txtSR.Text = "1-" And txtContactID.Text = "" Then

                    MsgBox("A Service Request # or Contact ID is required before saving", MessageBoxButtons.RetryCancel)


                Else


                    If cboAgentName.Text = "Agent Name" Then


                        MsgBox("Please be advised you must select an 'agent name' before proceeding", MessageBoxButtons.RetryCancel)


                    Else


                        If cboSRType.Text = "" Then

                            MsgBox("A SR Type must be selected before saving", MessageBoxButtons.RetryCancel)

                            Me.ActiveControl = cboSRType

                        Else

                            'If dtpCondate.Value = Today Then




                            '    MsgBox("Are you sure the Contact date for this Audit is Today?", MessageBoxButtons.RetryCancel)


                            'Else
                            If cboSupervisor.Text = "Supervisor" Then


                                MsgBox("Please be advised you must select an 'Supervisor' before proceeding", MessageBoxButtons.RetryCancel)


                            Else

                                If txtTeamName.Text = "Please wait, Loading.." Then




                                    MsgBox("The agent’s team field is still loading, please wait until a team name appears before saving the scorecard", MessageBoxButtons.RetryCancel)


                                    Me.ActiveControl = txtTeamName


                                Else




                                    If cboAutoFail.Checked = True And cboAF.Text = "" Then


                                        MsgBox("Since this Audit was marked as 'Auto Fail', a reason must be selected before saving.", MessageBoxButtons.RetryCancel)



                                        Me.ActiveControl = cboAF


                                    Else

                                        If cboCSAT1.Text = "" Or cboCSAT2.Text = "" Or cboCSAT3.Text = "" Or cboCSAT4.Text = "" Or cboCSAT5.Text = "" Or cboCSAT6.Text = "" Then


                                            MsgBox("Please be advised you must fill out the CSAT Equivalency section below before you proceed", MessageBoxButtons.RetryCancel)

                                            Me.ActiveControl = cboCSAT1

                                        Else


                                            CSatScore()

                                            QaTotalScore()
                                            TCXscore()

                                            lblTCXscore.Visible = False

                                            If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then




                                            Else

                                                SplashScreenManager1.ShowWaitForm()
                                                Me.Cursor = Cursors.WaitCursor

                                                buttondisables()

                                                QAChatdisableControls()



                                                Me.ActiveControl = txtSR




                                                BackgroundWorker1.RunWorkerAsync()

                                                store()


                                                ' PleaseWait.ShowDialog()



                                            End If






                                        End If

                                    End If

                                    '




                                End If

                            End If

                        End If
                    End If
                End If
            End If





        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Public Sub MissedWeightsReset()

        txt1_1a.Text = 1
        txt1_2a.Text = 1


        txt2_1a.Text = 1

        txt3_1a.Text = 1
        txt3_2a.Text = 1
        txt3_3a.Text = 1
        txt3_4a.Text = 1
        txt3_5a.Text = 1
        txt3_6a.Text = 1
        txt3_7a.Text = 1
        txt3_8a.Text = 1

        txt4_1a.Text = 1
        txt4_2a.Text = 1
        txt4_3a.Text = 1

        txt5_1a.Text = 1
        txt5_2a.Text = 1

        txt6_1a.Text = 1
        txt6_2a.Text = 1
        txt6_3a.Text = 1

        txt7_1a.Text = 1
        txt7_2a.Text = 1
        txt7_3a.Text = 1
        txt7_4a.Text = 1
        txt7_5a.Text = 1
        txt7_6a.Text = 1



    End Sub




    Public Sub MissedWeightsCalc()






        If cbo1_1.Text = 0 Then

            txt1_1a.Text = "0"

        Else



        End If


        If cbo1_2.Text = 0 Then

            txt1_2a.Text = "0"

        Else



        End If





        If cbo2_1.Text = 0 Then

            txt2_1a.Text = "0"

        Else


        End If







        If cbo3_1.Text = 0 Then

            txt3_1a.Text = "0"

        Else


        End If


        If cbo3_2.Text = 0 Then

            txt3_2a.Text = "0"

        Else



        End If



        If cbo3_3.Text = 0 Then

            txt3_3a.Text = "0"

        Else



        End If



        If cbo3_4.Text = 0 Then

            txt3_4a.Text = "0"
        Else


        End If



        If cbo3_5.Text = 0 Then

            txt3_5a.Text = "0"

        Else



        End If



        If cbo3_6.Text = 0 Then

            txt3_6a.Text = "0"

        Else



        End If



        If cbo3_7.Text = 0 Then

            txt3_7a.Text = "0"

        Else


        End If



        If cbo3_8.Text = 0 Then

            txt3_8a.Text = "0"

        Else



        End If




        If Cbo4_1.Text = 0 Then

            txt4_1a.Text = "0"
        Else



        End If

        If cbo4_2.Text = 0 Then

            txt4_2a.Text = "0"

        Else



        End If



        If cbo4_3.Text = 0 Then

            txt4_3a.Text = "0"

        Else



        End If




        If cbo5_1.Text = 0 Then

            txt5_1a.Text = "0"

        Else



        End If


        If cbo5_2.Text = 0 Then


            txt5_2a.Text = "0"


        Else



        End If





        If cbo6_1.Text = 0 Then

            txt6_1a.Text = "0"

        Else



        End If



        If cbo6_2.Text = 0 Then

            txt6_2a.Text = "0"

        Else



        End If





        If cbo6_3.Text = 0 Then

            txt6_3a.Text = "0"

        Else

        End If






        If cbo7_1.Text = 0 Then

            txt7_1a.Text = "0"

        Else



        End If



        If cbo7_2.Text = 0 Then

            txt7_2a.Text = "0"

        Else



        End If



        If cbo7_3.Text = 0 Then

            txt7_3a.Text = "0"
        Else



        End If



        If cbo7_4.Text = 0 Then

            txt7_1a.Text = "0"

        Else


        End If



        If cbo7_5.Text = 0 Then

            txt7_5a.Text = "0"

        Else


        End If




        If cbo7_6.Text = 0 Then

            txt7_6a.Text = "0"

        Else



        End If








    End Sub


    Public Sub store()




        Try





            ' Test

            '    con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")


            '' Dyanic

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")



            'P Drive 
            con = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")


            'P new

            ' con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")


            con.Open()






            Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Supervisor],[TCX_Score],[Week_Number],[EditedQA],[1_1],[1_2],[2_1],[3_1],[3_2],[3_3],[3_4],[3_5],[3_6],[3_7],[3_8],[4_1],[4_2],[4_3],[5_1],[5_2],[6_1],[6_2],[6_3],[7_1],[7_2],[7_3],[7_4],[7_5],[7_6],[Month],[PendingDisputeID],[Dis_TCXScore],[SRType],[MainSupervisor],[CSATScore],[CSATQ1],[CSATQ2],[CSATQ3],[CSATQ4],[CSATQ5],[CSATQ6])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Supervisor, @TCX_Score, @Week_Number, @EditedQA,@1_1,@1_2,@2_1,@3_1,@3_2,@3_3,@3_4,@3_5,@3_6,@3_7,@3_8,@4_1,@4_2,@4_3,@5_1,@5_2,@6_1,@6_2,@6_3,@7_1,@7_2,@7_3,@7_4,@7_5,@7_6,@Month,@PendingDisputeID,@Dis_TCXScore,@SRType,@MainSupervisor,@CSATScore,@CSATQ1,@CSATQ2,@CSATQ3,@CSATQ4,@CSATQ5,@CSATQ6)"







            Using cmd As New SqlCommand(SQL, con)


                If txtSR.Text = "1-" Then

                    cmd.Parameters.AddWithValue("@SR", DBNull.Value)

                Else
                    cmd.Parameters.AddWithValue("@SR", txtSR.Text)

                End If



                ' cmd.Parameters.AddWithValue("@SR", txtSR.Text)
                cmd.Parameters.AddWithValue("@ContactID", txtContactID.Text)
                cmd.Parameters.AddWithValue("@CType", "Chat")
                cmd.Parameters.AddWithValue("@QA_Agent", cboAgentName.Text)
                cmd.Parameters.AddWithValue("@QA_Team", txtTeamName.Text)
                cmd.Parameters.AddWithValue("@QA_ContactDate", DateTimePicker1.Value)
                cmd.Parameters.AddWithValue("@QA_OrderID", txtOrderID.Text)
                cmd.Parameters.AddWithValue("@QA_Date", Now)
                cmd.Parameters.AddWithValue("@QA_Comments", txtQACom.Text)
                cmd.Parameters.AddWithValue("@QA_Opp", txtQAAOO.Text)
                cmd.Parameters.AddWithValue("@CI_Name", txtContactName.Text)
                cmd.Parameters.AddWithValue("@CI_Account", txtAccountNum.Text)
                cmd.Parameters.AddWithValue("@CI_Company", txtCompany.Text)
                cmd.Parameters.AddWithValue("@CI_Phone", txtContactPhone.Text)
                cmd.Parameters.AddWithValue("@CI_Email", txtContactEmail.Text)

                If txtRevCom.Text = "" Then


                    cmd.Parameters.AddWithValue("@Rev_Date", "9/9/2020")
                    cmd.Parameters.AddWithValue("@Rev_Manager", cboSupervisor.Text)
                    cmd.Parameters.AddWithValue("@Rev_Comments", "")
                    cmd.Parameters.AddWithValue("@PendingDisputeID", "Pending Review")

                ElseIf txtRevCom.Text <> "" Then

                    cmd.Parameters.AddWithValue("@Rev_Date", txtQADate.Text)
                    cmd.Parameters.AddWithValue("@Rev_Manager", cboSupervisor.Text)
                    cmd.Parameters.AddWithValue("@Rev_Comments", txtRevCom.Text)
                    cmd.Parameters.AddWithValue("@PendingDisputeID", "Reviewed")


                End If





                cmd.Parameters.AddWithValue("@One_1", cbo1_1.Text)
                cmd.Parameters.AddWithValue("@One_2", cbo1_2.Text)

                cmd.Parameters.AddWithValue("@One_1Note", txt1_1.Text)
                cmd.Parameters.AddWithValue("@One_2Note", txt1_2.Text)



                cmd.Parameters.AddWithValue("@Two_1", cbo2_1.Text)
                cmd.Parameters.AddWithValue("@Two_1Note", txt2_1.Text)


                cmd.Parameters.AddWithValue("@Three_1", cbo3_1.Text)
                cmd.Parameters.AddWithValue("@Three_2", cbo3_2.Text)
                cmd.Parameters.AddWithValue("@Three_3", cbo3_3.Text)
                cmd.Parameters.AddWithValue("@Three_4", cbo3_4.Text)
                cmd.Parameters.AddWithValue("@Three_5", cbo3_5.Text)
                cmd.Parameters.AddWithValue("@Three_6", cbo3_6.Text)
                cmd.Parameters.AddWithValue("@Three_7", cbo3_7.Text)
                cmd.Parameters.AddWithValue("@Three_8", cbo3_8.Text)


                cmd.Parameters.AddWithValue("@Three_1Note", txt3_1.Text)
                cmd.Parameters.AddWithValue("@Three_2Note", txt3_2.Text)
                cmd.Parameters.AddWithValue("@Three_3Note", txt3_3.Text)
                cmd.Parameters.AddWithValue("@Three_4Note", txt3_4.Text)
                cmd.Parameters.AddWithValue("@Three_5Note", txt3_5.Text)
                cmd.Parameters.AddWithValue("@Three_6Note", txt3_6.Text)
                cmd.Parameters.AddWithValue("@Three_7Note", txt3_7.Text)
                cmd.Parameters.AddWithValue("@Three_8Note", txt3_8.Text)

                cmd.Parameters.AddWithValue("@Four_1", Cbo4_1.Text)
                cmd.Parameters.AddWithValue("@Four_2", cbo4_2.Text)
                cmd.Parameters.AddWithValue("@Four_3", cbo4_3.Text)

                cmd.Parameters.AddWithValue("@Four_1Note", txt4_1.Text)
                cmd.Parameters.AddWithValue("@Four_2Note", txt4_2.Text)
                cmd.Parameters.AddWithValue("@Four_3Note", txt4_3.Text)

                cmd.Parameters.AddWithValue("@Five_1", cbo5_1.Text)
                cmd.Parameters.AddWithValue("@Five_2", cbo5_2.Text)


                cmd.Parameters.AddWithValue("Five_1Note", txt5_1.Text)
                cmd.Parameters.AddWithValue("Five_2Note", txt5_2.Text)




                cmd.Parameters.AddWithValue("@Six_1", cbo6_1.Text)
                cmd.Parameters.AddWithValue("@Six_2", cbo6_2.Text)
                cmd.Parameters.AddWithValue("@Six_3", cbo6_3.Text)



                cmd.Parameters.AddWithValue("@Six_1Note", txt6_1.Text)
                cmd.Parameters.AddWithValue("@Six_2Note", txt6_2.Text)
                cmd.Parameters.AddWithValue("@Six_3Note", txt6_3.Text)


                cmd.Parameters.AddWithValue("@Seven_1", cbo7_1.Text)
                cmd.Parameters.AddWithValue("@Seven_2", cbo7_2.Text)
                cmd.Parameters.AddWithValue("@Seven_3", cbo7_3.Text)
                cmd.Parameters.AddWithValue("@Seven_4", cbo7_4.Text)
                cmd.Parameters.AddWithValue("@Seven_5", cbo7_5.Text)
                cmd.Parameters.AddWithValue("@Seven_6", cbo7_6.Text)

                cmd.Parameters.AddWithValue("Seven_1Note", txt7_1.Text)
                cmd.Parameters.AddWithValue("Seven_2Note", txt7_2.Text)
                cmd.Parameters.AddWithValue("Seven_3Note", txt7_3.Text)
                cmd.Parameters.AddWithValue("Seven_4Note", txt7_4.Text)
                cmd.Parameters.AddWithValue("Seven_5Note", txt7_5.Text)
                cmd.Parameters.AddWithValue("Seven_6Note", txt7_6.Text)


                cmd.Parameters.AddWithValue("@1_1", txt1_1a.Text)
                cmd.Parameters.AddWithValue("@1_2", txt1_2a.Text)



                cmd.Parameters.AddWithValue("@2_1", txt2_1a.Text)

                cmd.Parameters.AddWithValue("@3_1", txt3_1a.Text)
                cmd.Parameters.AddWithValue("@3_2", txt3_2a.Text)
                cmd.Parameters.AddWithValue("@3_3", txt3_3a.Text)
                cmd.Parameters.AddWithValue("@3_4", txt3_4a.Text)
                cmd.Parameters.AddWithValue("@3_5", txt3_5a.Text)
                cmd.Parameters.AddWithValue("@3_6", txt3_6a.Text)
                cmd.Parameters.AddWithValue("@3_7", txt3_7a.Text)
                cmd.Parameters.AddWithValue("@3_8", txt3_8a.Text)

                cmd.Parameters.AddWithValue("@4_1", txt4_1a.Text)
                cmd.Parameters.AddWithValue("@4_2", txt4_2a.Text)
                cmd.Parameters.AddWithValue("@4_3", txt4_3a.Text)

                cmd.Parameters.AddWithValue("@5_1", txt5_1a.Text)
                cmd.Parameters.AddWithValue("@5_2", txt5_2a.Text)

                cmd.Parameters.AddWithValue("@6_1", txt6_1a.Text)
                cmd.Parameters.AddWithValue("@6_2", txt6_2a.Text)
                cmd.Parameters.AddWithValue("@6_3", txt6_3a.Text)


                cmd.Parameters.AddWithValue("@7_1", txt7_1a.Text)
                cmd.Parameters.AddWithValue("@7_2", txt7_2a.Text)
                cmd.Parameters.AddWithValue("@7_3", txt7_3a.Text)
                cmd.Parameters.AddWithValue("@7_4", txt7_4a.Text)
                cmd.Parameters.AddWithValue("@7_5", txt7_5a.Text)
                cmd.Parameters.AddWithValue("@7_6", txt7_6a.Text)



                If cboAutoFail.Checked = True Then

                    cmd.Parameters.AddWithValue("@QAScore", "0")
                    cmd.Parameters.AddWithValue("@Autofail", cboAF.Text)


                    cmd.Parameters.AddWithValue("@Dis_Score", "0")


                ElseIf cboAutoFail.Checked = False Then

                    cmd.Parameters.AddWithValue("@QAScore", txtQAScore.Text)
                    cmd.Parameters.AddWithValue("@Autofail", "N/a")
                    cmd.Parameters.AddWithValue("@Dis_Score", txtQAScore.Text)


                End If

                cmd.Parameters.AddWithValue("@Dis_TCXScore", txtTCXScore.Text)
                cmd.Parameters.AddWithValue("@Dis_Name", "")
                cmd.Parameters.AddWithValue("@Dis_Notes", "")
                cmd.Parameters.AddWithValue("@Dis_AppComments", "")

                cmd.Parameters.AddWithValue("@Auditor", lblQAauditor1.Text)
                cmd.Parameters.AddWithValue("@Supervisor", cboSupervisor.Text)
                cmd.Parameters.AddWithValue("@TCX_Score", txtTCXScore.Text)
                cmd.Parameters.AddWithValue("@Week_Number", Form2.lblYear.Text + " - " + "Week " + lblWeekNumber.Text)
                cmd.Parameters.AddWithValue("@EditedQA", "0")
                cmd.Parameters.AddWithValue("@Month", Form2.lblMonth.Text)
                cmd.Parameters.AddWithValue("@SRType", cboSRType.Text)
                cmd.Parameters.AddWithValue("@MainSupervisor", cboSupervisor.Text)


                cmd.Parameters.AddWithValue("@CSATScore", txtCSATScore.Text)

                cmd.Parameters.AddWithValue("@CSATQ1", cboCSAT1.Text)
                cmd.Parameters.AddWithValue("@CSATQ2", cboCSAT2.Text)
                cmd.Parameters.AddWithValue("@CSATQ3", cboCSAT3.Text)
                cmd.Parameters.AddWithValue("@CSATQ4", cboCSAT4.Text)
                cmd.Parameters.AddWithValue("@CSATQ5", cboCSAT5.Text)
                cmd.Parameters.AddWithValue("@CSATQ6", cboCSAT6.Text)



                cmd.ExecuteNonQuery()

                con.Close()



            End Using

            ' Saver.Enabled = True

            ExcelSaver.Enabled = True

        Catch ex As SqlException


            MsgBox(ex.Message)



            ProgressBar1.Value = 0
            lblprogr.Text = 0

            QAChatEnable()

            buttonEnables()

            Saver.Enabled = False


            Me.Cursor = Cursors.Hand




        Catch ex As Exception


            MsgBox(ex.Message)

            Me.Cursor = Cursors.Hand

            ProgressBar1.Value = 0
            lblprogr.Text = 0

            QAChatEnable()
            buttonEnables()


            Saver.Enabled = False


        End Try



    End Sub


    Public Sub store2()




        Try





            ' Test

            '   con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")


            '' Dyanic

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")



            'P Drive 
            con = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")


            'P new

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")


            con.Open()



            ' Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Supervisor],[TCX_Score],[EditedQA])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Supervisor, @TCX_Score, @EditedQA)"

            Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Supervisor],[TCX_Score],[Week_Number],[EditedQA],[1_1],[1_2],[2_1],[3_1],[3_2],[3_3],[3_4],[3_5],[3_6],[3_7],[3_8],[4_1],[4_2],[4_3],[5_1],[5_2],[6_1],[6_2],[6_3],[7_1],[7_2],[7_3],[7_4],[7_5],[7_6],[Month],[PendingDisputeID],[Dis_TCXScore],[SRType],[MainSupervisor],[CSATScore],[CSATQ1],[CSATQ2],[CSATQ3],[CSATQ4],[CSATQ5],[CSATQ6])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Supervisor, @TCX_Score, @Week_Number, @EditedQA,@1_1,@1_2,@2_1,@3_1,@3_2,@3_3,@3_4,@3_5,@3_6,@3_7,@3_8,@4_1,@4_2,@4_3,@5_1,@5_2,@6_1,@6_2,@6_3,@7_1,@7_2,@7_3,@7_4,@7_5,@7_6,@Month,@PendingDisputeID,@Dis_TCXScore,@SRType,@MainSupervisor,@CSATScore,@CSATQ1,@CSATQ2,@CSATQ3,@CSATQ4,@CSATQ5,@CSATQ6)"




            Using cmd As New SqlCommand(SQL, con)




                If txtSR.Text = "1-" Then

                    cmd.Parameters.AddWithValue("@SR", DBNull.Value)

                Else
                    cmd.Parameters.AddWithValue("@SR", txtSR.Text)

                End If



                '    cmd.Parameters.AddWithValue("@SR", txtSR.Text)
                cmd.Parameters.AddWithValue("@ContactID", txtContactID.Text)
                cmd.Parameters.AddWithValue("@CType", "Chat")
                cmd.Parameters.AddWithValue("@QA_Agent", cboSupervisorbox.Text)
                cmd.Parameters.AddWithValue("@QA_Team", txtTeamName.Text)
                cmd.Parameters.AddWithValue("@QA_ContactDate", DateTimePicker1.Value)
                cmd.Parameters.AddWithValue("@QA_OrderID", txtOrderID.Text)
                cmd.Parameters.AddWithValue("@QA_Date", Now)
                cmd.Parameters.AddWithValue("@QA_Comments", txtQACom.Text)
                cmd.Parameters.AddWithValue("@QA_Opp", txtQAAOO.Text)
                cmd.Parameters.AddWithValue("@CI_Name", txtContactName.Text)
                cmd.Parameters.AddWithValue("@CI_Account", txtAccountNum.Text)
                cmd.Parameters.AddWithValue("@CI_Company", txtCompany.Text)
                cmd.Parameters.AddWithValue("@CI_Phone", txtContactPhone.Text)
                cmd.Parameters.AddWithValue("@CI_Email", txtContactEmail.Text)



                If txtRevCom.Text = "" Then


                    cmd.Parameters.AddWithValue("@Rev_Date", "9/9/2020")
                    cmd.Parameters.AddWithValue("@Rev_Manager", lblQAauditor1.Text)
                    cmd.Parameters.AddWithValue("@Rev_Comments", "")
                    cmd.Parameters.AddWithValue("@PendingDisputeID", "Pending Review")

                ElseIf txtRevCom.Text <> "" Then

                    cmd.Parameters.AddWithValue("@Rev_Date", txtQADate.Text)
                    cmd.Parameters.AddWithValue("@Rev_Manager", lblQAauditor1.Text)
                    cmd.Parameters.AddWithValue("@Rev_Comments", txtRevCom.Text)
                    cmd.Parameters.AddWithValue("@PendingDisputeID", "Reviewed")


                End If




                '  cmd.Parameters.AddWithValue("@Dis_Score", lblQAScore1.Text)
                cmd.Parameters.AddWithValue("@Dis_TCXScore", txtTCXScore.Text)
                cmd.Parameters.AddWithValue("@Dis_Name", "")
                cmd.Parameters.AddWithValue("@Dis_Notes", "")
                cmd.Parameters.AddWithValue("@Dis_AppComments", "")


                cmd.Parameters.AddWithValue("@One_1", cbo1_1.Text)
                cmd.Parameters.AddWithValue("@One_2", cbo1_2.Text)

                cmd.Parameters.AddWithValue("@One_1Note", txt1_1.Text)
                cmd.Parameters.AddWithValue("@One_2Note", txt1_2.Text)



                cmd.Parameters.AddWithValue("@Two_1", cbo2_1.Text)
                cmd.Parameters.AddWithValue("@Two_1Note", txt2_1.Text)


                cmd.Parameters.AddWithValue("@Three_1", cbo3_1.Text)
                cmd.Parameters.AddWithValue("@Three_2", cbo3_2.Text)
                cmd.Parameters.AddWithValue("@Three_3", cbo3_3.Text)
                cmd.Parameters.AddWithValue("@Three_4", cbo3_4.Text)
                cmd.Parameters.AddWithValue("@Three_5", cbo3_5.Text)
                cmd.Parameters.AddWithValue("@Three_6", cbo3_6.Text)
                cmd.Parameters.AddWithValue("@Three_7", cbo3_7.Text)
                cmd.Parameters.AddWithValue("@Three_8", cbo3_8.Text)


                cmd.Parameters.AddWithValue("@Three_1Note", txt3_1.Text)
                cmd.Parameters.AddWithValue("@Three_2Note", txt3_2.Text)
                cmd.Parameters.AddWithValue("@Three_3Note", txt3_3.Text)
                cmd.Parameters.AddWithValue("@Three_4Note", txt3_4.Text)
                cmd.Parameters.AddWithValue("@Three_5Note", txt3_5.Text)
                cmd.Parameters.AddWithValue("@Three_6Note", txt3_6.Text)
                cmd.Parameters.AddWithValue("@Three_7Note", txt3_7.Text)
                cmd.Parameters.AddWithValue("@Three_8Note", txt3_8.Text)

                cmd.Parameters.AddWithValue("@Four_1", Cbo4_1.Text)
                cmd.Parameters.AddWithValue("@Four_2", cbo4_2.Text)
                cmd.Parameters.AddWithValue("@Four_3", cbo4_3.Text)

                cmd.Parameters.AddWithValue("@Four_1Note", txt4_1.Text)
                cmd.Parameters.AddWithValue("@Four_2Note", txt4_2.Text)
                cmd.Parameters.AddWithValue("@Four_3Note", txt4_3.Text)

                cmd.Parameters.AddWithValue("@Five_1", cbo5_1.Text)
                cmd.Parameters.AddWithValue("@Five_2", cbo5_2.Text)


                cmd.Parameters.AddWithValue("Five_1Note", txt5_1.Text)
                cmd.Parameters.AddWithValue("Five_2Note", txt5_2.Text)




                cmd.Parameters.AddWithValue("@Six_1", cbo6_1.Text)
                cmd.Parameters.AddWithValue("@Six_2", cbo6_2.Text)
                cmd.Parameters.AddWithValue("@Six_3", cbo6_3.Text)



                cmd.Parameters.AddWithValue("@Six_1Note", txt6_1.Text)
                cmd.Parameters.AddWithValue("@Six_2Note", txt6_2.Text)
                cmd.Parameters.AddWithValue("@Six_3Note", txt6_3.Text)


                cmd.Parameters.AddWithValue("@Seven_1", cbo7_1.Text)
                cmd.Parameters.AddWithValue("@Seven_2", cbo7_2.Text)
                cmd.Parameters.AddWithValue("@Seven_3", cbo7_3.Text)
                cmd.Parameters.AddWithValue("@Seven_4", cbo7_4.Text)
                cmd.Parameters.AddWithValue("@Seven_5", cbo7_5.Text)
                cmd.Parameters.AddWithValue("@Seven_6", cbo7_6.Text)

                cmd.Parameters.AddWithValue("Seven_1Note", txt7_1.Text)
                cmd.Parameters.AddWithValue("Seven_2Note", txt7_2.Text)
                cmd.Parameters.AddWithValue("Seven_3Note", txt7_3.Text)
                cmd.Parameters.AddWithValue("Seven_4Note", txt7_4.Text)
                cmd.Parameters.AddWithValue("Seven_5Note", txt7_5.Text)
                cmd.Parameters.AddWithValue("Seven_6Note", txt7_6.Text)

                cmd.Parameters.AddWithValue("@1_1", txt1_1a.Text)
                cmd.Parameters.AddWithValue("@1_2", txt1_2a.Text)



                cmd.Parameters.AddWithValue("@2_1", txt2_1a.Text)

                cmd.Parameters.AddWithValue("@3_1", txt3_1a.Text)
                cmd.Parameters.AddWithValue("@3_2", txt3_2a.Text)
                cmd.Parameters.AddWithValue("@3_3", txt3_3a.Text)
                cmd.Parameters.AddWithValue("@3_4", txt3_4a.Text)
                cmd.Parameters.AddWithValue("@3_5", txt3_5a.Text)
                cmd.Parameters.AddWithValue("@3_6", txt3_6a.Text)
                cmd.Parameters.AddWithValue("@3_7", txt3_7a.Text)
                cmd.Parameters.AddWithValue("@3_8", txt3_8a.Text)

                cmd.Parameters.AddWithValue("@4_1", txt4_1a.Text)
                cmd.Parameters.AddWithValue("@4_2", txt4_2a.Text)
                cmd.Parameters.AddWithValue("@4_3", txt4_3a.Text)

                cmd.Parameters.AddWithValue("@5_1", txt5_1a.Text)
                cmd.Parameters.AddWithValue("@5_2", txt5_2a.Text)

                cmd.Parameters.AddWithValue("@6_1", txt6_1a.Text)
                cmd.Parameters.AddWithValue("@6_2", txt6_2a.Text)
                cmd.Parameters.AddWithValue("@6_3", txt6_3a.Text)


                cmd.Parameters.AddWithValue("@7_1", txt7_1a.Text)
                cmd.Parameters.AddWithValue("@7_2", txt7_2a.Text)
                cmd.Parameters.AddWithValue("@7_3", txt7_3a.Text)
                cmd.Parameters.AddWithValue("@7_4", txt7_4a.Text)
                cmd.Parameters.AddWithValue("@7_5", txt7_5a.Text)
                cmd.Parameters.AddWithValue("@7_6", txt7_6a.Text)





                If cboAutoFail.Checked = True Then

                    cmd.Parameters.AddWithValue("@QAScore", "0")
                    cmd.Parameters.AddWithValue("@Autofail", cboAF.Text)
                    cmd.Parameters.AddWithValue("@Dis_Score", "0")

                ElseIf cboAutoFail.Checked = False Then

                    cmd.Parameters.AddWithValue("@QAScore", txtQAScore.Text)
                    cmd.Parameters.AddWithValue("@Autofail", "N/a")
                    cmd.Parameters.AddWithValue("@Dis_Score", txtQAScore.Text)

                End If

                cmd.Parameters.AddWithValue("@Auditor", lblQAauditor1.Text)
                cmd.Parameters.AddWithValue("@Supervisor", lblQAauditor1.Text)
                cmd.Parameters.AddWithValue("@TCX_Score", txtTCXScore.Text)
                cmd.Parameters.AddWithValue("@Week_Number", Form2.lblYear.Text + " - " + "Week " + lblWeekNumber.Text)
                cmd.Parameters.AddWithValue("@EditedQA", "0")
                cmd.Parameters.AddWithValue("@Month", Form2.lblMonth.Text)
                cmd.Parameters.AddWithValue("@SRType", cboSRType.Text)
                cmd.Parameters.AddWithValue("@MainSupervisor", lblQAauditor1.Text)


                cmd.Parameters.AddWithValue("@CSATScore", txtCSATScore.Text)

                cmd.Parameters.AddWithValue("@CSATQ1", cboCSAT1.Text)
                cmd.Parameters.AddWithValue("@CSATQ2", cboCSAT2.Text)
                cmd.Parameters.AddWithValue("@CSATQ3", cboCSAT3.Text)
                cmd.Parameters.AddWithValue("@CSATQ4", cboCSAT4.Text)
                cmd.Parameters.AddWithValue("@CSATQ5", cboCSAT5.Text)
                cmd.Parameters.AddWithValue("@CSATQ6", cboCSAT6.Text)



                cmd.ExecuteNonQuery()

                con.Close()



            End Using



            ' Saver2.Enabled = True


            ExcelSaver2.Enabled = True

        Catch ex As SqlException


            MsgBox(ex.Message)



            ProgressBar1.Value = 0
            lblprogr.Text = 0

            QAChatEnable()

            buttonEnables()

            Saver2.Enabled = False


            Me.Cursor = Cursors.Hand




        Catch ex As Exception


            MsgBox(ex.Message)

            Me.Cursor = Cursors.Hand

            ProgressBar1.Value = 0
            lblprogr.Text = 0

            QAChatEnable()

            Saver2.Enabled = False


        End Try



    End Sub

    Public Sub QAChatdisableControls()


        cbo1_1.Enabled = False
        cbo1_2.Enabled = False


        cbo2_1.Enabled = False

        cbo3_1.Enabled = False
        cbo3_2.Enabled = False
        cbo3_3.Enabled = False
        cbo3_4.Enabled = False
        cbo3_5.Enabled = False
        cbo3_6.Enabled = False
        cbo3_7.Enabled = False
        cbo3_8.Enabled = False

        Cbo4_1.Enabled = False
        cbo4_2.Enabled = False
        cbo4_3.Enabled = False


        cbo5_1.Enabled = False
        cbo5_2.Enabled = False

        cbo6_1.Enabled = False
        cbo6_2.Enabled = False
        cbo6_3.Enabled = False

        cbo7_1.Enabled = False
        cbo7_2.Enabled = False
        cbo7_3.Enabled = False
        cbo7_4.Enabled = False
        cbo7_5.Enabled = False
        cbo7_6.Enabled = False


        ''reset Textboxes

        'txt1_1.Enabled = False
        'txt1_2.Enabled = False




        'txt2_1.Enabled = False


        'txt3_1.Enabled = False
        'txt3_2.Enabled = False
        'txt3_3.Enabled = False
        'txt3_4.Enabled = False
        'txt3_5.Enabled = False
        'txt3_6.Enabled = False
        'txt3_7.Enabled = False
        'txt3_8.Enabled = False



        'txt4_1.Enabled = False
        'txt4_2.Enabled = False
        'txt4_3.Enabled = False

        'txt5_1.Enabled = False
        'txt5_2.Enabled = False




        'txt6_1.Enabled = False
        'txt6_2.Enabled = False
        'txt6_3.Enabled = False


        'txt7_1.Enabled = False
        'txt7_2.Enabled = False
        'txt7_3.Enabled = False
        'txt7_4.Enabled = False
        'txt7_5.Enabled = False
        'txt7_6.Enabled = False

    End Sub

    Public Sub resetatglance()

        ''Reset Scorecard at a glance info

        Me.ActiveControl = txtSR


        cboAgentName.Text = "Agent Name"
        '  cboTeamName.Text = "Team Name"

        txtSR.Clear()



        lblTCXscore.Visible = False


    End Sub


    Public Sub QAChatclear()


        ''Reset Comboboxes

        cbo1_1.Text = 3
        cbo1_2.Text = 3


        cbo2_1.Text = 4

        cbo3_1.Text = 2
        cbo3_2.Text = 4
        cbo3_3.Text = 7
        cbo3_4.Text = 5
        cbo3_5.Text = 4
        cbo3_6.Text = 3
        cbo3_7.Text = 3
        cbo3_8.Text = 2

        Cbo4_1.Text = 5
        cbo4_2.Text = 5
        cbo4_3.Text = 5


        cbo5_1.Text = 7
        cbo5_2.Text = 8

        cbo6_1.Text = 2
        cbo6_2.Text = 3
        cbo6_3.Text = 10

        cbo7_1.Text = 2
        cbo7_2.Text = 2
        cbo7_3.Text = 3
        cbo7_4.Text = 4
        cbo7_5.Text = 1
        cbo7_6.Text = 3


        ''reset Textboxes

        txt1_1.Clear()
        txt1_2.Clear()



        txt2_1.Clear()


        txt3_1.Clear()
        txt3_2.Clear()
        txt3_3.Clear()
        txt3_4.Clear()
        txt3_5.Clear()
        txt3_6.Clear()
        txt3_7.Clear()
        txt3_8.Clear()



        txt4_1.Clear()
        txt4_2.Clear()
        txt4_3.Clear()

        txt5_1.Clear()
        txt5_2.Clear()




        txt6_1.Clear()
        txt6_2.Clear()
        txt6_3.Clear()


        txt7_1.Clear()
        txt7_2.Clear()
        txt7_3.Clear()
        txt7_4.Clear()
        txt7_5.Clear()
        txt7_6.Clear()

        MissedWeightsReset()



        txt1_1.BackColor = Color.White

        txt1_2.BackColor = Color.White



        txt2_1.BackColor = Color.White


        txt3_1.BackColor = Color.White
        txt3_2.BackColor = Color.White
        txt3_3.BackColor = Color.White
        txt3_4.BackColor = Color.White
        txt3_5.BackColor = Color.White
        txt3_6.BackColor = Color.White
        txt3_7.BackColor = Color.White
        txt3_8.BackColor = Color.White



        txt4_1.BackColor = Color.White
        txt4_2.BackColor = Color.White
        txt4_3.BackColor = Color.White

        txt5_1.BackColor = Color.White
        txt5_2.BackColor = Color.White




        txt6_1.BackColor = Color.White
        txt6_2.BackColor = Color.White
        txt6_3.BackColor = Color.White


        txt7_1.BackColor = Color.White
        txt7_2.BackColor = Color.White
        txt7_3.BackColor = Color.White
        txt7_4.BackColor = Color.White
        txt7_5.BackColor = Color.White
        txt7_6.BackColor = Color.White


        txtQACom.Clear()
        txtQAAOO.Clear()


        txtAgentEmail.Clear()


        txtSR.Clear()
        txtContactID.Clear()
        txtContactName.Clear()
        txtContactEmail.Clear()
        txtContactPhone.Clear()
        txtAccountNum.Clear()
        txtCompany.Clear()
        txtOrderID.Clear()



        '  txtTeamName.Clear()



        cboCSAT1.SelectedIndex = -1
        cboCSAT2.SelectedIndex = -1
        cboCSAT3.SelectedIndex = -1
        cboCSAT4.SelectedIndex = -1
        cboCSAT5.SelectedIndex = -1
        cboCSAT6.SelectedIndex = -1

        txtCSATScore.Clear()
        txtTCXScore.Clear()

        cboSRType.SelectedIndex = -1

        lblQAScore1.Text = "100"
        txtQAScore.Text = "100"
        '    lblQAScore1.Visible = True

        lblTCXscore.Text = "100"

        txtRevCom.Clear()


    End Sub

    Public Sub QAChatEnable()




        ''Reset Comboboxes

        cbo1_1.Enabled = True
        cbo1_2.Enabled = True


        cbo2_1.Enabled = True

        cbo3_1.Enabled = True
        cbo3_2.Enabled = True
        cbo3_3.Enabled = True
        cbo3_4.Enabled = True
        cbo3_5.Enabled = True
        cbo3_6.Enabled = True
        cbo3_7.Enabled = True
        cbo3_8.Enabled = True

        Cbo4_1.Enabled = True
        cbo4_2.Enabled = True
        cbo4_3.Enabled = True


        cbo5_1.Enabled = True
        cbo5_2.Enabled = True

        cbo6_1.Enabled = True
        cbo6_2.Enabled = True
        cbo6_3.Enabled = True

        cbo7_1.Enabled = True
        cbo7_2.Enabled = True
        cbo7_3.Enabled = True
        cbo7_4.Enabled = True
        cbo7_5.Enabled = True
        cbo7_6.Enabled = True


        ''reset Textboxes

        txt1_1.Enabled = True
        txt1_2.Enabled = True




        txt2_1.Enabled = True


        txt3_1.Enabled = True
        txt3_2.Enabled = True
        txt3_3.Enabled = True
        txt3_4.Enabled = True
        txt3_5.Enabled = True
        txt3_6.Enabled = True
        txt3_7.Enabled = True
        txt3_8.Enabled = True



        txt4_1.Enabled = True
        txt4_2.Enabled = True
        txt4_3.Enabled = True

        txt5_1.Enabled = True
        txt5_2.Enabled = True




        txt6_1.Enabled = True
        txt6_2.Enabled = True
        txt6_3.Enabled = True


        txt7_1.Enabled = True
        txt7_2.Enabled = True
        txt7_3.Enabled = True
        txt7_4.Enabled = True
        txt7_5.Enabled = True
        txt7_6.Enabled = True

    End Sub



    Public Sub QAExcell()




        Try



            Dim oExcel As Object = CreateObject("Excel.Application")





            '' P Drive

            '   Dim oBook As Object = oExcel.Workbooks.Open("P:\QA Application\QA1\Chat.xlsx")




            '' Resouce
            Dim exeDir As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)
            Dim xlpath = IO.Path.Combine(exeDir.DirectoryName, "Chat.xlsx")


            Dim obook As Object = oExcel.Workbooks.Open(xlpath)




            Dim oSheet As Object = obook.Worksheets("Chat")  'or oBook.Worksheets("SheetName")




            oSheet.Range("D4").Value = "" & cbo1_1.Text
            oSheet.Range("D5").Value = "" & cbo1_2.Text



            oSheet.Range("H4").Value = "" & txt1_1.Text
            oSheet.Range("H5").Value = "" & txt1_2.Text

            ' oSheet.Range("C6").Value = "" & two

            oSheet.Range("D7").Value = "" & cbo2_1.Text
            oSheet.Range("H7").Value = "" & txt2_1.Text

            '   oSheet.Range("C8").Value = "" & three

            oSheet.Range("D9").Value = "" & cbo3_1.Text
            oSheet.Range("D10").Value = "" & cbo3_2.Text
            oSheet.Range("D11").Value = "" & cbo3_3.Text
            oSheet.Range("D12").Value = "" & cbo3_4.Text
            oSheet.Range("D13").Value = "" & cbo3_5.Text
            oSheet.Range("D14").Value = "" & cbo3_6.Text
            oSheet.Range("D15").Value = "" & cbo3_7.Text
            oSheet.Range("D16").Value = "" & cbo3_8.Text



            oSheet.Range("H9").Value = "" & txt3_1.Text
            oSheet.Range("H10").Value = "" & txt3_2.Text
            oSheet.Range("H11").Value = "" & txt3_3.Text
            oSheet.Range("H12").Value = "" & txt3_4.Text
            oSheet.Range("H13").Value = "" & txt3_5.Text
            oSheet.Range("H14").Value = "" & txt3_6.Text
            oSheet.Range("H15").Value = "" & txt3_7.Text
            oSheet.Range("H16").Value = "" & txt3_8.Text

            '   oSheet.Range("C17").Value = "" & Four

            oSheet.Range("D18").Value = "" & Cbo4_1.Text
            oSheet.Range("D19").Value = "" & cbo4_2.Text
            oSheet.Range("D20").Value = "" & cbo4_3.Text

            oSheet.Range("H18").Value = "" & txt4_1.Text
            oSheet.Range("H19").Value = "" & txt4_2.Text
            oSheet.Range("H20").Value = "" & txt4_3.Text


            ' oSheet.Range("C21").Value = "" & Five


            oSheet.Range("D22").Value = "" & cbo5_1.Text
            oSheet.Range("D23").Value = "" & cbo5_2.Text

            oSheet.Range("H22").Value = "" & txt5_1.Text
            oSheet.Range("H23").Value = "" & txt5_2.Text

            '   oSheet.Range("C24").Value = "" & Six


            oSheet.Range("D25").Value = "" & cbo6_1.Text
            oSheet.Range("D26").Value = "" & cbo6_2.Text
            oSheet.Range("D27").Value = "" & cbo6_3.Text


            oSheet.Range("H25").Value = "" & txt6_1.Text
            oSheet.Range("H26").Value = "" & txt6_2.Text
            oSheet.Range("H27").Value = "" & txt6_3.Text

            '   oSheet.Range("C28").Value = "" & seven


            oSheet.Range("D29").Value = "" & cbo7_1.Text
            oSheet.Range("D30").Value = "" & cbo7_2.Text
            oSheet.Range("D31").Value = "" & cbo7_3.Text
            oSheet.Range("D32").Value = "" & cbo7_4.Text
            oSheet.Range("D33").Value = "" & cbo7_5.Text
            oSheet.Range("D34").Value = "" & cbo7_6.Text


            oSheet.Range("H29").Value = "" & txt7_1.Text
            oSheet.Range("H30").Value = "" & txt7_2.Text
            oSheet.Range("H31").Value = "" & txt7_3.Text
            oSheet.Range("H32").Value = "" & txt7_4.Text
            oSheet.Range("H33").Value = "" & txt7_5.Text
            oSheet.Range("H34").Value = "" & txt7_6.Text




            If cboAutoFail.Checked = True And lblAutoZero.Visible = True Then

                oSheet.Range("C36").Value = "0"

            Else

                oSheet.Range("C36").Value = txtQAScore.Text


            End If


            oSheet.Range("C37").Value = txtTCXScore.Text


            oSheet.Range("C38").Value = txtSR.Text
            oSheet.Range("C39").Value = txtContactID.Text
            oSheet.Range("C40").Value = "Chat"
            oSheet.Range("C41").Value = "" & cboAgentName.Text
            oSheet.Range("C42").Value = "" & txtTeamName.Text

            oSheet.Range("C43").Value = DateTimePicker1.Text


            oSheet.Range("C44").Value = txtOrderID.Text

            oSheet.Range("C45").Value = "" & txtContactName.Text
            oSheet.Range("C46").Value = "" & txtContactPhone.Text
            oSheet.Range("C47").Value = "" & txtContactEmail.Text
            oSheet.Range("C48").Value = "" & txtCompany.Text
            oSheet.Range("C49").Value = "" & txtAccountNum.Text
            oSheet.Range("C50").Value = "" & cboAF.Text
            oSheet.Range("C51").Value = "" & lblQAauditor1.Text
            oSheet.Range("C52").Value = "" & txtQADate.Text



            oSheet.Range("B54").Value = txtQACom.Text
            oSheet.Range("B58").Value = txtQAAOO.Text





            ''Review

            oSheet.Range("C66").Value = "" & lblDate.Text
            oSheet.Range("C67").Value = "" & lblQAauditor1.Text
            oSheet.Range("B68").Value = "" & txtRevCom.Text


            oSheet.Range("C94").Value = "" & txtCSATScore.Text

            oSheet.Range("C95").Value = "" & cboCSAT1.Text
            oSheet.Range("C96").Value = "" & cboCSAT2.Text
            oSheet.Range("C97").Value = "" & cboCSAT3.Text
            oSheet.Range("C98").Value = "" & cboCSAT4.Text
            oSheet.Range("C99").Value = "" & cboCSAT5.Text
            oSheet.Range("C100").Value = "" & cboCSAT6.Text





            ''Dispute

            'oSheet.Range("C76").Value = "" & txtDisputeScore.Text
            'oSheet.Range("C77").Value = "" & lblrevMan.Text
            'oSheet.Range("B78").Value = "" & txtDisputeNotes.Text


            'oSheet.Range("C86").Value = "" & cboDisApproval.Text
            'oSheet.Range("B87").Value = "" & txtDiscomment.Text




            ' iF contactid is being used   

            If txtContactID.Text <> String.Empty And txtSR.Text = "1-" Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & "Chat QA Scorecard.xlsx")

                Saver2.Enabled = True

            End If

            ' If SR is being used
            If txtContactID.Text = String.Empty Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & "Chat QA Scorecard.xlsx")


                Saver.Enabled = True

            End If

            '' bothing being used
            If txtContactID.Text <> String.Empty And txtSR.Text <> "1-" Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & "Chat QA Scorecard.xlsx")


                Saver.Enabled = True

            End If


            oExcel.Quit()



        Catch ex As Exception



            MsgBox(ex.Message)



            ProgressBar1.Value = 0
            lblprogr.Text = 0

            QAChatEnable()

            buttonEnables()




        End Try

    End Sub



    Public Sub QAExcel2()




        Try



            Dim oExcel As Object = CreateObject("Excel.Application")





            '' P Drive

            '  Dim oBook As Object = oExcel.Workbooks.Open("P:\QA Application\QA1\Chat.xlsx")

            '' Resouce
            Dim exeDir As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)
            Dim xlpath = IO.Path.Combine(exeDir.DirectoryName, "Chat.xlsx")


            Dim obook As Object = oExcel.Workbooks.Open(xlpath)




            Dim oSheet As Object = obook.Worksheets("Chat")  'or oBook.Worksheets("SheetName")




            oSheet.Range("D4").Value = "" & cbo1_1.Text
            oSheet.Range("D5").Value = "" & cbo1_2.Text



            oSheet.Range("H4").Value = "" & txt1_1.Text
            oSheet.Range("H5").Value = "" & txt1_2.Text

            ' oSheet.Range("C6").Value = "" & two

            oSheet.Range("D7").Value = "" & cbo2_1.Text
            oSheet.Range("H7").Value = "" & txt2_1.Text

            '   oSheet.Range("C8").Value = "" & three

            oSheet.Range("D9").Value = "" & cbo3_1.Text
            oSheet.Range("D10").Value = "" & cbo3_2.Text
            oSheet.Range("D11").Value = "" & cbo3_3.Text
            oSheet.Range("D12").Value = "" & cbo3_4.Text
            oSheet.Range("D13").Value = "" & cbo3_5.Text
            oSheet.Range("D14").Value = "" & cbo3_6.Text
            oSheet.Range("D15").Value = "" & cbo3_7.Text
            oSheet.Range("D16").Value = "" & cbo3_8.Text



            oSheet.Range("H9").Value = "" & txt3_1.Text
            oSheet.Range("H10").Value = "" & txt3_2.Text
            oSheet.Range("H11").Value = "" & txt3_3.Text
            oSheet.Range("H12").Value = "" & txt3_4.Text
            oSheet.Range("H13").Value = "" & txt3_5.Text
            oSheet.Range("H14").Value = "" & txt3_6.Text
            oSheet.Range("H15").Value = "" & txt3_7.Text
            oSheet.Range("H16").Value = "" & txt3_8.Text

            '   oSheet.Range("C17").Value = "" & Four

            oSheet.Range("D18").Value = "" & Cbo4_1.Text
            oSheet.Range("D19").Value = "" & cbo4_2.Text
            oSheet.Range("D20").Value = "" & cbo4_3.Text

            oSheet.Range("H18").Value = "" & txt4_1.Text
            oSheet.Range("H19").Value = "" & txt4_2.Text
            oSheet.Range("H20").Value = "" & txt4_3.Text


            ' oSheet.Range("C21").Value = "" & Five


            oSheet.Range("D22").Value = "" & cbo5_1.Text
            oSheet.Range("D23").Value = "" & cbo5_2.Text

            oSheet.Range("H22").Value = "" & txt5_1.Text
            oSheet.Range("H23").Value = "" & txt5_2.Text

            '   oSheet.Range("C24").Value = "" & Six


            oSheet.Range("D25").Value = "" & cbo6_1.Text
            oSheet.Range("D26").Value = "" & cbo6_2.Text
            oSheet.Range("D27").Value = "" & cbo6_3.Text


            oSheet.Range("H25").Value = "" & txt6_1.Text
            oSheet.Range("H26").Value = "" & txt6_2.Text
            oSheet.Range("H27").Value = "" & txt6_3.Text

            '   oSheet.Range("C28").Value = "" & seven


            oSheet.Range("D29").Value = "" & cbo7_1.Text
            oSheet.Range("D30").Value = "" & cbo7_2.Text
            oSheet.Range("D31").Value = "" & cbo7_3.Text
            oSheet.Range("D32").Value = "" & cbo7_4.Text
            oSheet.Range("D33").Value = "" & cbo7_5.Text
            oSheet.Range("D34").Value = "" & cbo7_6.Text


            oSheet.Range("H29").Value = "" & txt7_1.Text
            oSheet.Range("H30").Value = "" & txt7_2.Text
            oSheet.Range("H31").Value = "" & txt7_3.Text
            oSheet.Range("H32").Value = "" & txt7_4.Text
            oSheet.Range("H33").Value = "" & txt7_5.Text
            oSheet.Range("H34").Value = "" & txt7_6.Text



            If cboAutoFail.Checked = True And lblAutoZero.Visible = True Then

                oSheet.Range("C36").Value = "0"

            Else

                oSheet.Range("C36").Value = txtQAScore.Text


            End If




            oSheet.Range("C37").Value = txtTCXScore.Text


            oSheet.Range("C38").Value = txtSR.Text
            oSheet.Range("C39").Value = txtContactID.Text
            oSheet.Range("C40").Value = "Chat"
            oSheet.Range("C41").Value = "" & cboSupervisorbox.Text
            oSheet.Range("C42").Value = "" & txtTeamName.Text

            oSheet.Range("C43").Value = DateTimePicker1.Text


            oSheet.Range("C44").Value = txtOrderID.Text

            oSheet.Range("C45").Value = "" & txtContactName.Text
            oSheet.Range("C46").Value = "" & txtContactPhone.Text
            oSheet.Range("C47").Value = "" & txtContactEmail.Text
            oSheet.Range("C48").Value = "" & txtCompany.Text
            oSheet.Range("C49").Value = "" & txtAccountNum.Text
            oSheet.Range("C50").Value = "" & cboAF.Text
            oSheet.Range("C51").Value = "" & lblQAauditor1.Text
            oSheet.Range("C52").Value = "" & txtQADate.Text



            oSheet.Range("B54").Value = txtQACom.Text
            oSheet.Range("B58").Value = txtQAAOO.Text





            ''Review

            oSheet.Range("C66").Value = "" & lblDate.Text
            oSheet.Range("C67").Value = "" & lblQAauditor1.Text
            oSheet.Range("B68").Value = "" & txtRevCom.Text


            oSheet.Range("C94").Value = "" & txtCSATScore.Text

            oSheet.Range("C95").Value = "" & cboCSAT1.Text
            oSheet.Range("C96").Value = "" & cboCSAT2.Text
            oSheet.Range("C97").Value = "" & cboCSAT3.Text
            oSheet.Range("C98").Value = "" & cboCSAT4.Text
            oSheet.Range("C99").Value = "" & cboCSAT5.Text
            oSheet.Range("C100").Value = "" & cboCSAT6.Text





            ''Dispute

            'oSheet.Range("C76").Value = "" & txtDisputeScore.Text
            'oSheet.Range("C77").Value = "" & lblrevMan.Text
            'oSheet.Range("B78").Value = "" & txtDisputeNotes.Text


            'oSheet.Range("C86").Value = "" & cboDisApproval.Text
            'oSheet.Range("B87").Value = "" & txtDiscomment.Text



            ' iF contactid is being used   

            If txtContactID.Text <> String.Empty And txtSR.Text = "1-" Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtContactID.Text & " " & cboSupervisorbox.Text & "-" & "Chat QA Scorecard.xlsx")

                Saver2.Enabled = True

            End If

            ' If SR is being used
            If txtContactID.Text = String.Empty Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "Chat QA Scorecard.xlsx")


                Saver.Enabled = True

            End If


            '' bothing being used
            If txtContactID.Text <> String.Empty And txtSR.Text <> "1-" Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "Chat QA Scorecard.xlsx")


                Saver.Enabled = True

            End If

            oExcel.Quit()




        Catch ex As Exception



            MsgBox(ex.Message)



            ProgressBar1.Value = 0
            lblprogr.Text = 0

            QAChatEnable()

            buttonEnables()



        End Try

    End Sub


    Private Sub Time_Tick(sender As Object, e As EventArgs) Handles Time.Tick


        txtQADate.Text = Date.Now.ToString("MM/dd/yyyy")


        lblDate.Text = Date.Now.ToString("MM/dd/yyyy")

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork


        Try

            For i = 0 To 100

                System.Threading.Thread.Sleep(30)
                Me.BackgroundWorker1.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i
            Next


            ''

            ' store()

            ' Send to Excell
            ' QAExcell()



            'StoreCallThread = New System.Threading.Thread(AddressOf store)
            ''
            'StoreCallThread.Start()


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub



    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

        ProgressBar1.Value = e.ProgressPercentage



    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        Me.Cursor = Cursors.Hand

        'PleaseWait.Hide()


        ''   lblQAScore1.Visible = True

        ''    If MsgBox(cboAgentName.Text & " " & "" & "scored a total of" & " " & lblQAScore1.Text & " " & "points on this QA audit, would you like to start a new ‘Chat’ audit?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then

        'If MsgBox("The audit for " & cboAgentName.Text & " " & "was successfully saved, would you like to start a New 'Chat’ audit?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then



        '    reset()


        '    buttonEnables()

        '    Form2.Clear()


        '    Form2.Show()

        '    Me.Hide()


        'Else

        '    buttonEnables()

        '    reset()

        '    '   cboTeamName.Text = "Team Name"
        '    cboAgentName.Text = "Agent Name"

        'End If




    End Sub

    Public Sub reset()


        ''Reset Scorecard at a glance info

        resetatglance()

        ''Reset scorecard

        QAChatclear()

        ''Transfer Qa Name to Wasetupform


        ' Form2.lblQAauditor.Text = lblQAauditor1.Text


        ''Reable buttons

        QAChatEnable()


        ' Me.Hide()


        cboAutoFail.Checked = False

        ProgressBar1.Value = 0
        lblprogr.Text = 0


        txtQACom.BackColor = Color.White

        cboAF.Visible = False








    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt1_1.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt1_1.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt1_1.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()


        Catch Excep As Exception
            MessageBox.Show(Excep.Message)

        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt1_2.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt1_2.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt1_2.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt2_1.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt2_1.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt2_1.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt3_1.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt3_1.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt3_1.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt3_2.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt3_2.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt3_2.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt3_3.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt3_3.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt3_3.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt3_4.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt3_4.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt3_4.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt3_5.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt3_5.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt3_5.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt3_6.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt3_6.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt3_6.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt3_7.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt3_7.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt3_7.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt3_8.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt3_8.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt3_8.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs)


        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt4_1.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt4_1.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt4_1.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try








    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt4_2.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt4_2.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt4_2.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt4_3.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt4_3.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt4_3.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt5_1.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt5_1.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt5_1.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt5_2.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt5_2.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt5_2.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt6_1.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt6_1.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt6_1.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt6_2.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt6_2.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt6_2.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt6_3.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt6_3.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt6_3.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt7_1.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt7_1.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt7_1.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub


    Private Sub Button22_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt7_2.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt7_2.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt7_2.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt7_3.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt7_3.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt7_3.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try




    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt7_4.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt7_4.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt7_4.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt7_5.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt7_5.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt7_5.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt7_6.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txt7_6.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txt7_6.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub


    'Private Sub cbo1_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo1_1.SelectedIndexChanged

    '    Dim int1_1 As Integer = cbo1_1.Text

    '    If cbo1_1.Text = 0 Then


    '        txt1_1.BackColor = Color.Yellow

    '    ElseIf cbo1_1.Text > 0 Then


    '        txt1_1.BackColor = Color.White


    '    End If






    'End Sub

    'Private Sub cbo1_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo1_2.SelectedIndexChanged

    '    Dim int1_2 As Integer = cbo1_2.Text

    '    If cbo1_2.Text = 0 Then


    '        txt1_2.BackColor = Color.Yellow

    '    ElseIf cbo1_2.Text > 0 Then


    '        txt1_2.BackColor = Color.White


    '    End If


    'End Sub

    'Private Sub cbo2_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo2_1.SelectedIndexChanged



    '    If cbo2_1.Text = 0 Then


    '        txt2_1.BackColor = Color.Yellow

    '    ElseIf cbo2_1.Text > 0 Then


    '        txt2_1.BackColor = Color.White


    '    End If

    'End Sub

    'Private Sub cbo3_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_1.SelectedIndexChanged

    '    If cbo3_1.Text = 0 Then


    '        txt3_1.BackColor = Color.Yellow

    '    ElseIf cbo3_1.Text > 0 Then


    '        txt3_1.BackColor = Color.White


    '    End If



    'End Sub



    'Private Sub cbo3_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_2.SelectedIndexChanged


    '    If cbo3_2.Text = 0 Then


    '        txt3_2.BackColor = Color.Yellow

    '    ElseIf cbo3_2.Text > 0 Then


    '        txt3_2.BackColor = Color.White


    '    End If



    'End Sub

    'Private Sub cbo3_3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_3.SelectedIndexChanged


    '    If cbo3_3.Text = 0 Then


    '        txt3_3.BackColor = Color.Yellow

    '    ElseIf cbo3_3.Text > 0 Then


    '        txt3_3.BackColor = Color.White


    '    End If





    'End Sub

    'Private Sub cbo3_4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_4.SelectedIndexChanged
    '    If cbo3_4.Text = 0 Then


    '        txt3_4.BackColor = Color.Yellow

    '    ElseIf cbo3_4.Text > 0 Then


    '        txt3_4.BackColor = Color.White


    '    End If


    'End Sub

    'Private Sub cbo3_5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_5.SelectedIndexChanged


    '    If cbo3_5.Text = 0 Then


    '        txt3_5.BackColor = Color.Yellow

    '    ElseIf cbo3_5.Text > 0 Then


    '        txt3_5.BackColor = Color.White


    '    End If


    'End Sub

    'Private Sub cbo3_6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_6.SelectedIndexChanged


    '    If cbo3_6.Text = 0 Then


    '        txt3_6.BackColor = Color.Yellow

    '    ElseIf cbo3_6.Text > 0 Then


    '        txt3_6.BackColor = Color.White


    '    End If






    'End Sub

    'Private Sub cbo3_7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_7.SelectedIndexChanged

    '    If cbo3_7.Text = 0 Then


    '        txt3_7.BackColor = Color.Yellow

    '    ElseIf cbo3_7.Text > 0 Then


    '        txt3_7.BackColor = Color.White


    '    End If


    'End Sub

    'Private Sub cbo3_8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_8.SelectedIndexChanged


    '    If cbo3_8.Text = 0 Then


    '        txt3_8.BackColor = Color.Yellow

    '    ElseIf cbo3_8.Text > 0 Then


    '        txt3_8.BackColor = Color.White


    '    End If


    'End Sub


    'Private Sub Cbo4_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cbo4_1.SelectedIndexChanged

    '    If Cbo4_1.Text = 0 Then


    '        txt4_1.BackColor = Color.Yellow

    '    ElseIf Cbo4_1.Text > 0 Then


    '        txt4_1.BackColor = Color.White


    '    End If
    'End Sub

    'Private Sub cbo4_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo4_2.SelectedIndexChanged

    '    If cbo4_2.Text = 0 Then


    '        txt4_2.BackColor = Color.Yellow

    '    ElseIf cbo4_2.Text > 0 Then


    '        txt4_2.BackColor = Color.White


    '    End If
    'End Sub

    'Private Sub cbo4_3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo4_3.SelectedIndexChanged

    '    If cbo4_3.Text = 0 Then


    '        txt4_3.BackColor = Color.Yellow

    '    ElseIf cbo4_3.Text > 0 Then


    '        txt4_3.BackColor = Color.White


    '    End If
    'End Sub

    'Private Sub cbo5_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo5_1.SelectedIndexChanged

    '    If cbo5_1.Text = 0 Then


    '        txt5_1.BackColor = Color.Yellow

    '    ElseIf cbo5_1.Text > 0 Then


    '        txt5_1.BackColor = Color.White


    '    End If
    'End Sub

    'Private Sub cbo5_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo5_2.SelectedIndexChanged

    '    If cbo5_2.Text = 0 Then


    '        txt5_2.BackColor = Color.Yellow

    '    ElseIf cbo5_2.Text > 0 Then


    '        txt5_2.BackColor = Color.White


    '    End If

    'End Sub

    'Private Sub cbo6_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo6_1.SelectedIndexChanged


    '    If cbo6_1.Text = 0 Then


    '        txt6_1.BackColor = Color.Yellow

    '    ElseIf cbo6_1.Text > 0 Then


    '        txt6_1.BackColor = Color.White


    '    End If




    'End Sub

    'Private Sub cbo6_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo6_2.SelectedIndexChanged


    '    If cbo6_2.Text = 0 Then


    '        txt6_2.BackColor = Color.Yellow

    '    ElseIf cbo6_2.Text > 0 Then


    '        txt6_2.BackColor = Color.White


    '    End If

    'End Sub

    'Private Sub cbo6_3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo6_3.SelectedIndexChanged
    '    If cbo6_3.Text = 0 Then


    '        txt6_3.BackColor = Color.Yellow

    '    ElseIf cbo6_3.Text > 0 Then


    '        txt6_3.BackColor = Color.White


    '    End If

    'End Sub

    'Private Sub cbo7_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo7_1.SelectedIndexChanged
    '    If cbo7_1.Text = 0 Then


    '        txt7_1.BackColor = Color.Yellow

    '    ElseIf cbo7_1.Text > 0 Then


    '        txt7_1.BackColor = Color.White


    '    End If
    'End Sub

    'Private Sub cbo7_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo7_2.SelectedIndexChanged
    '    If cbo7_2.Text = 0 Then


    '        txt7_2.BackColor = Color.Yellow

    '    ElseIf cbo7_2.Text > 0 Then


    '        txt7_2.BackColor = Color.White

    '    End If

    'End Sub

    'Private Sub cbo7_3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo7_3.SelectedIndexChanged
    '    If cbo7_3.Text = 0 Then


    '        txt7_3.BackColor = Color.Yellow

    '    ElseIf cbo7_3.Text > 0 Then


    '        txt7_3.BackColor = Color.White

    '    End If

    'End Sub

    'Private Sub cbo7_4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo7_4.SelectedIndexChanged
    '    If cbo7_4.Text = 0 Then


    '        txt7_4.BackColor = Color.Yellow

    '    ElseIf cbo7_4.Text > 0 Then


    '        txt7_4.BackColor = Color.White

    '    End If

    'End Sub

    'Private Sub cbo7_5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo7_5.SelectedIndexChanged
    '    If cbo7_5.Text = 0 Then


    '        txt7_5.BackColor = Color.Yellow

    '    ElseIf cbo7_5.Text > 0 Then


    '        txt7_5.BackColor = Color.White

    '    End If

    'End Sub

    'Private Sub cbo7_6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo7_6.SelectedIndexChanged
    '    If cbo7_6.Text = 0 Then


    '        txt7_6.BackColor = Color.Yellow

    '    ElseIf cbo7_6.Text > 0 Then


    '        txt7_6.BackColor = Color.White

    '    End If


    'End Sub


    Private Sub Button27_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txtQACom.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txtQACom.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txtQACom.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try









    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txtQAAOO.Text = "" Then

                Exit Sub
            End If

            objWord = New Word.Application()
            objTempDoc = objWord.Documents.Add
            objWord.Visible = False

            ' Position Word off the screen...this keeps Word invisible 
            ' throughout.
            objWord.WindowState = 0
            objWord.Top = -3000

            ' Copy the contents of the textbox to the clipboard
            Clipboard.SetDataObject(txtQAAOO.Text)

            ' With the temporary document, perform either a spell check or a 
            ' complete
            ' grammar check, based on user selection.
            With objTempDoc
                .Content.Paste()
                .Activate()


                .CheckSpelling()

                '  .CheckGrammar()


                ' After user has made changes, use the clipboard to
                ' transfer the contents back to the text box

                .Content.Copy()
                iData = Clipboard.GetDataObject
                If iData.GetDataPresent(DataFormats.Text) Then
                    txtQAAOO.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub QAChatScorecard_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        Try

            If MessageBox.Show("Are you sure to close this application?", "FADV Quality Assurance Application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

                End

            Else
                e.Cancel = True


            End If


        Catch ex As Exception

            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub cboAutoFail_CheckStateChanged(sender As Object, e As EventArgs)



        If cboAutoFail.CheckState = CheckState.Checked Then




            cboAF.Visible = True


        ElseIf cboAutoFail.CheckState = CheckState.Unchecked Then


            cboAF.Visible = False

            cboAF.Text = "N/a"

        End If

    End Sub





    Private Sub cboTeamName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSupervisor.SelectedIndexChanged

        Try

            Me.Cursor = Cursors.WaitCursor

            resetcombo()



            cboAgentName.Text = "Please wait, Loading.."

            BackgroundWorker5.RunWorkerAsync()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try
    End Sub

    Private Sub cboSupervisorbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSupervisorbox.SelectedIndexChanged



        Try


            txtTeamName.Text = "Please wait, Loading.."


            BackgroundWorker7.RunWorkerAsync()


        Catch ex As Exception



            MsgBox(ex.Message)


        End Try





    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork



        Try

            For i = 0 To 100

                System.Threading.Thread.Sleep(30)
                Me.BackgroundWorker1.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i
            Next




            ''

            ' store2()

            ' Send to Excell
            '  QAExcell()



            'StoreCallThread = New System.Threading.Thread(AddressOf store2)

            'StoreCallThread.Start()




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub BackgroundWorker2_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged


        ProgressBar1.Value = e.ProgressPercentage



    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted



        Me.Cursor = Cursors.Hand

        'PleaseWait.Hide()

        '' If MsgBox(cboSupervisorbox.Text & " " & "" & "scored a total of" & " " & lblQAScore1.Text & " " & "points on this QA audit, would you like to start a new ‘Chat’ audit?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then

        'If MsgBox("The audit for " & cboSupervisorbox.Text & " " & "was successfully saved, would you like to start a New 'Chat’ audit?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then

        '    reset()

        '    buttonEnables()


        '    Form2.Clear()


        '    Form2.Show()

        '    Me.Hide()


        'Else

        '    buttonEnables()

        '    reset()


        '    cboSupervisorbox.Text = "Agent Name"

        'End If








    End Sub

    Private Sub btnSpellChecker_Click(sender As Object, e As EventArgs) Handles btnSpellChecker.Click

        Try

            SpellChecker2.CheckContainer(Me)



        Catch ex As Exception



            MsgBox(ex.Message)


        End Try





    End Sub

    Private Sub cboContactType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboContactTypeChat.SelectedIndexChanged



        Dim msg = "Are you sure you want to change the Scorecard?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)




        If responce = MsgBoxResult.Yes Then




            If cboContactTypeChat.Text = "Call" Then

                QACallScorecard.txtSR.Text = txtSR.Text
                QACallScorecard.txtContactID.Text = txtContactID.Text


                QACallScorecard.txtContactName.Text = txtContactName.Text
                QACallScorecard.txtContactEmail.Text = txtContactEmail.Text
                QACallScorecard.txtContactPhone.Text = txtContactPhone.Text
                QACallScorecard.txtAccountNum.Text = txtAccountNum.Text
                QACallScorecard.txtCompany.Text = txtCompany.Text
                QACallScorecard.txtOrderID.Text = txtOrderID.Text


                ' QACallScorecard.dtpCondate.Text = ProgramDate.ToString(ProgramDateForamt)

                'QACallScorecard.dtpCondate.Text = DateTimePicker1.Text
                ' QACallScorecard.cboContactTypeCall.Text = "Call"
                QACallScorecard.lblQAauditor1.Text = lblQAauditor1.Text

                QACallScorecard.Show()



                Me.Hide()


            ElseIf cboContactTypeChat.Text = "Email" Then


                QAEmailScorecard.txtSR.Text = txtSR.Text
                QAEmailScorecard.txtContactID.Text = txtContactID.Text


                QAEmailScorecard.txtContactName.Text = txtContactName.Text
                QAEmailScorecard.txtContactEmail.Text = txtContactEmail.Text
                QAEmailScorecard.txtContactPhone.Text = txtContactPhone.Text
                QAEmailScorecard.txtAccountNum.Text = txtAccountNum.Text
                QAEmailScorecard.txtCompany.Text = txtCompany.Text
                QAEmailScorecard.txtOrderID.Text = txtOrderID.Text

                ' QAEmailScorecard.DateTimePicker1.Text = DateTimePicker1.Text

                '     QAEmailScorecard.DateTimePicker1.Text = ProgramDate.ToString(ProgramDateForamt)

                QAEmailScorecard.lblQAauditor1.Text = lblQAauditor1.Text

                ' QAEmailScorecard.cboContactTypeEmail.Text = "Email"


                QAEmailScorecard.Show()


                Me.Hide()


            ElseIf cboContactTypeChat.Text = "Chat" Then





            End If






        Else








        End If




    End Sub

    Private Sub BackgroundWorker3_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork


        Fillcombo()



    End Sub

    Private Sub BackgroundWorker4_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork


        Fillcombo33()


    End Sub

    Private Sub BackgroundWorker5_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork


        Try

            QaSetupMod.connecttemp4()


            sqltemp4 = "SELECT * FROM [Agents] WHERE Supervisor='" & cboSupervisor.Text & " ' "

            '    sqltemp2 = "SELECT * FROM [Teams]"



            Dim cmdtemp As New SqlCommand





            cmdtemp.CommandText = sqltemp4

            cmdtemp.Connection = contemp4



            readertemp4 = cmdtemp.ExecuteReader


            While (readertemp4.Read())


                cboAgentName.Items.Add(readertemp4("AgentName"))

                lblSupervisorEmail.Text = (readertemp4("SuperEmail"))


            End While



            cmdtemp.Dispose()

            readertemp4.Close()



            Me.Cursor = Cursors.Hand

        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted






        contemp4.Close()


        cboAgentName.Text = "Agent Name"

        Me.Cursor = Cursors.Hand



    End Sub

    Private Sub cboAgentName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAgentName.SelectedIndexChanged

        Try



            txtTeamName.Text = "Please wait, Loading.."

            BackgroundWorker6.RunWorkerAsync()



        Catch ex As Exception



            MsgBox(ex.Message)


        End Try
    End Sub

    Private Sub BackgroundWorker6_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker6.DoWork

        Try


            QaSetupMod.connecttemp8()


            '   Me.Cursor = Cursors.WaitCursor

            sqltemp8 = "SELECT * FROM [Agents] WHERE AgentName='" & cboAgentName.Text & " ' "



            Dim cmdtemp As New SqlCommand





            cmdtemp.CommandText = sqltemp8

            cmdtemp.Connection = contemp8



            readertemp8 = cmdtemp.ExecuteReader



            If (readertemp8.Read() = True) Then






                txtTeamName.Text = (readertemp8("Platform"))

                txtAgentEmail.Text = (readertemp8("AgentEmail"))




            End If


            cmdtemp.Dispose()

            readertemp8.Close()


            Me.Cursor = Cursors.Hand


        Catch ex As Exception



            MsgBox(ex.Message)


        End Try










    End Sub

    Private Sub BackgroundWorker7_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker7.DoWork

        Try
            '  Me.Cursor = Cursors.WaitCursor

            QaSetupMod.connecttemp8()


            sqltemp8 = "SELECT * FROM [Agents] WHERE AgentName='" & cboSupervisorbox.Text & " ' "



            Dim cmdtemp As New SqlCommand





            cmdtemp.CommandText = sqltemp8

            cmdtemp.Connection = contemp8



            readertemp8 = cmdtemp.ExecuteReader



            If (readertemp8.Read() = True) Then



                txtTeamName.Text = (readertemp8("Platform"))


                txtAgentEmail.Text = (readertemp8("AgentEmail"))

            End If



            cmdtemp.Dispose()


            readertemp8.Close()



        Catch ex As Exception



            MsgBox(ex.Message)


        End Try




    End Sub

    Private Sub BackgroundWorker7_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker7.RunWorkerCompleted





        contemp8.Close()






        Me.Cursor = Cursors.Hand



    End Sub

    Private Sub lblQAScore1_Click(sender As Object, e As EventArgs) Handles lblQAScore1.Click

        QaTotalScore()


    End Sub

    Private Sub cbo1_1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo1_1.SelectionChangeCommitted


        Try

            If cbo1_1.SelectedItem = 0 Then


                txt1_1.BackColor = Color.Yellow


                totalQA = Convert.ToInt32(lblQAScore1.Text) - 3



                lblQAScore1.Text = totalQA
                txtQAScore.Text = totalQA


            ElseIf cbo1_1.SelectedItem = 3 Then


                txt1_1.BackColor = Color.White


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else



                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 3



                    lblQAScore1.Text = totalQA
                    txtQAScore.Text = totalQA

                End If

            End If



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub cbo1_2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo1_2.SelectionChangeCommitted

        Try


            If cbo1_2.SelectedItem = 0 Then


                txt1_2.BackColor = Color.Yellow


                totalQA = Convert.ToInt32(lblQAScore1.Text) - 3



                lblQAScore1.Text = totalQA
                txtQAScore.Text = totalQA


            ElseIf cbo1_2.SelectedItem = 3 Then


                txt1_2.BackColor = Color.White


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else



                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 3



                    lblQAScore1.Text = totalQA
                    txtQAScore.Text = totalQA

                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub cbo2_1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo2_1.SelectionChangeCommitted

        Try

            If cbo2_1.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 4



                lblQAScore1.Text = totalQA
                txtQAScore.Text = totalQA

                txt2_1.BackColor = Color.Yellow


            ElseIf cbo2_1.SelectedItem = "4" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 4


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt2_1.BackColor = Color.White

                End If

            End If



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo3_1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo3_1.SelectionChangeCommitted

        Try
            If cbo3_1.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 2


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_1.BackColor = Color.Yellow

            ElseIf cbo3_1.SelectedItem = "2" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 2


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt3_1.BackColor = Color.White

                End If

            End If





        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub

    Private Sub cbo3_2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo3_2.SelectionChangeCommitted


        Try
            If cbo3_2.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 4


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_2.BackColor = Color.Yellow

            ElseIf cbo3_2.SelectedItem = "4" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 4


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt3_2.BackColor = Color.White

                End If

            End If




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo3_3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo3_3.SelectionChangeCommitted


        Try

            If cbo3_3.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 7


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_3.BackColor = Color.Yellow

            ElseIf cbo3_3.SelectedItem = "7" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 7


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt3_3.BackColor = Color.White

                End If

            End If



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub

    Private Sub cbo3_4_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo3_4.SelectionChangeCommitted


        Try

            If cbo3_4.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 5


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_4.BackColor = Color.Yellow

            ElseIf cbo3_4.SelectedItem = "5" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 5


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt3_4.BackColor = Color.White

                End If

            End If







        Catch ex As Exception



            MsgBox(ex.Message)

        End Try

















    End Sub

    Private Sub cbo3_5_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo3_5.SelectionChangeCommitted



        Try
            If cbo3_5.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 4


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_5.BackColor = Color.Yellow

            ElseIf cbo3_5.SelectedItem = "4" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 4


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt3_5.BackColor = Color.White

                End If

            End If






        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





    End Sub

    Private Sub cbo3_6_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo3_6.SelectionChangeCommitted


        Try
            If cbo3_6.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 3


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_6.BackColor = Color.Yellow

            ElseIf cbo3_6.SelectedItem = "3" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 3


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt3_6.BackColor = Color.White

                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub

    Private Sub cbo3_7_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo3_7.SelectionChangeCommitted


        Try
            If cbo3_7.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 3


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_7.BackColor = Color.Yellow

            ElseIf cbo3_7.SelectedItem = "3" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 3

                    txtQAScore.Text = totalQA

                    lblQAScore1.Text = totalQA

                    txt3_7.BackColor = Color.White

                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo3_8_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo3_8.SelectionChangeCommitted

        Try
            If cbo3_8.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 2


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_8.BackColor = Color.Yellow

            ElseIf cbo3_8.SelectedItem = "2" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 2


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt3_8.BackColor = Color.White

                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub Cbo4_1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Cbo4_1.SelectionChangeCommitted


        Try
            If Cbo4_1.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 5


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt4_1.BackColor = Color.Yellow

            ElseIf Cbo4_1.SelectedItem = "5" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else

                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 5


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt4_1.BackColor = Color.White

                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo4_2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo4_2.SelectionChangeCommitted


        Try
            If cbo4_2.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 5


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt4_2.BackColor = Color.Yellow

            ElseIf cbo4_2.SelectedItem = "5" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else



                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 5


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt4_2.BackColor = Color.White

                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo4_3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo4_3.SelectionChangeCommitted

        Try
            If cbo4_3.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 5


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt4_3.BackColor = Color.Yellow

            ElseIf cbo4_3.SelectedItem = "5" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else



                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 5


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt4_3.BackColor = Color.White

                End If

            End If

        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub cbo5_1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo5_1.SelectionChangeCommitted

        Try


            If cbo5_1.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 7


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt5_1.BackColor = Color.Yellow

            ElseIf cbo5_1.SelectedItem = "7" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 7


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt5_1.BackColor = Color.White

                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo5_2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo5_2.SelectionChangeCommitted


        Try



            If cbo5_2.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 8


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt5_2.BackColor = Color.Yellow

            ElseIf cbo5_2.SelectedItem = "8" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 8


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt5_2.BackColor = Color.White

                End If

            End If



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo6_1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo6_1.SelectionChangeCommitted

        Try

            If cbo6_1.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 2


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt6_1.BackColor = Color.Yellow

            ElseIf cbo6_1.SelectedItem = "2" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 2


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt6_1.BackColor = Color.White


                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub

    Private Sub cbo6_2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo6_2.SelectionChangeCommitted

        Try
            If cbo6_2.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 3


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt6_2.BackColor = Color.Yellow

            ElseIf cbo6_2.SelectedItem = "3" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 3


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt6_2.BackColor = Color.White


                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo6_3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo6_3.SelectionChangeCommitted


        Try
            If cbo6_3.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 10


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt6_3.BackColor = Color.Yellow

            ElseIf cbo6_3.SelectedItem = "10" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 10


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt6_3.BackColor = Color.White

                End If

            End If



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub cbo7_1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo7_1.SelectionChangeCommitted


        Try
            If cbo7_1.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 2


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt7_1.BackColor = Color.Yellow

            ElseIf cbo7_1.SelectedItem = "2" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else



                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 2


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt7_1.BackColor = Color.White

                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub

    Private Sub cbo7_2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo7_2.SelectionChangeCommitted


        Try
            If cbo7_2.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 2


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt7_2.BackColor = Color.Yellow

            ElseIf cbo7_2.SelectedItem = "2" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else



                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 2


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt7_2.BackColor = Color.White



                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo7_3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo7_3.SelectionChangeCommitted


        Try
            If cbo7_3.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 3


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt7_3.BackColor = Color.Yellow

            ElseIf cbo7_3.SelectedItem = "3" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else

                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 3


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt7_3.BackColor = Color.White

                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo7_4_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo7_4.SelectionChangeCommitted
        Try

            If cbo7_4.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 4


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt7_4.BackColor = Color.Yellow

            ElseIf cbo7_4.SelectedItem = "4" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else



                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 4


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt7_4.BackColor = Color.White

                End If


            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo7_5_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo7_5.SelectionChangeCommitted

        Try
            If cbo7_5.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 1


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt7_5.BackColor = Color.Yellow

            ElseIf cbo7_5.SelectedItem = "1" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 1


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt7_5.BackColor = Color.White

                End If


            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub cbo7_6_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo7_6.SelectionChangeCommitted

        Try
            If cbo7_6.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 3


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt7_6.BackColor = Color.Yellow

            ElseIf cbo7_6.SelectedItem = "3" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 3

                    txtQAScore.Text = totalQA

                    lblQAScore1.Text = totalQA

                    txt7_6.BackColor = Color.White

                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub btnSave2_Click(sender As Object, e As EventArgs) Handles btnSave2.Click

        Try

            MissedWeightsCalc()

            If txtSR.Text <> "1-" And txtSR.MaskFull = False Then


                MsgBox("Please enter a valid SR#")


            Else



                If txtSR.Text = "1-" And txtContactID.Text = "" Then

                    MsgBox("A Service Request # or Contact ID is required before saving", MessageBoxButtons.RetryCancel)


                Else


                    If cboSupervisorbox.Text = "Agent Name" Then


                        MsgBox("Please be advised you must select an 'agent name' before proceeding", MessageBoxButtons.RetryCancel)

                        Me.ActiveControl = cboSupervisorbox


                    Else

                        If cboSRType.Text = "" Then

                            MsgBox("A SR Type must be selected before saving", MessageBoxButtons.RetryCancel)

                            Me.ActiveControl = cboSRType

                        Else

                            'If dtpCondate.Value = Today Then




                            '    MsgBox("Are you sure the Contact date for this Audit is Today?", MessageBoxButtons.RetryCancel)


                            ' Else


                            If txtTeamName.Text = "Please wait, Loading.." Then




                                MsgBox("The agent’s team field is still loading, please wait until a team name appears before saving the scorecard", MessageBoxButtons.RetryCancel)


                                Me.ActiveControl = txtTeamName


                            Else




                                If cboAutoFail.Checked = True And cboAF.Text = "" Then


                                    MsgBox("Since this Audit was marked as 'Auto Fail', a reason must be selected before saving.", MessageBoxButtons.RetryCancel)



                                    Me.ActiveControl = cboAF


                                Else

                                    If cboCSAT1.Text = "" Or cboCSAT2.Text = "" Or cboCSAT3.Text = "" Or cboCSAT4.Text = "" Or cboCSAT5.Text = "" Or cboCSAT6.Text = "" Then


                                        MsgBox("Please be advised you must fill out the CSAT Equivalency section below before you proceed", MessageBoxButtons.RetryCancel)

                                        Me.ActiveControl = cboCSAT1

                                    Else



                                        CSatScore()

                                        QaTotalScore()
                                        TCXscore()
                                        lblTCXscore.Visible = False

                                        'If txtRevCom.Text = "" Then




                                        '    If MsgBox("Would you like to review the scorecard now before saving?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then



                                        '        Me.ActiveControl = txtRevCom

                                        '    Else



                                        '        Me.Cursor = Cursors.WaitCursor

                                        '        buttondisables()

                                        '        QAChatdisableControls()


                                        '        Me.ActiveControl = txtSR

                                        '        BackgroundWorker2.RunWorkerAsync()

                                        '        store2()


                                        '        'PleaseWait.ShowDialog()



                                        '    End If

                                        'End If

                                        ''


                                        'If txtRevCom.Text <> String.Empty Then


                                        If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then




                                        Else
                                            SplashScreenManager1.ShowWaitForm()

                                            Me.Cursor = Cursors.WaitCursor

                                            buttondisables()

                                            QAChatdisableControls()


                                            Me.ActiveControl = txtSR


                                            BackgroundWorker2.RunWorkerAsync()


                                            store2()

                                            ' PleaseWait.ShowDialog()




                                            '

                                        End If

                                    End If

                                End If
                                '




                            End If


                        End If

                    End If

                End If

            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub lblTCXscore_Click(sender As Object, e As EventArgs) Handles lblTCXscore.Click



        TCXscore()


    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

        Try


            Process.Start("P:\QA Application\QA1\ChatD.docx")


        Catch ex As Exception



            '  MsgBox(ex.Message)

            MsgBox("Make sure your are connected to the P drive.")

        End Try



    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted




        contemp10.Close()


    End Sub

    Private Sub BackgroundWorker6_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted









        contemp8.Close()



    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted







        contemp4.Close()








    End Sub





    Private Sub Saver_Tick(sender As Object, e As EventArgs) Handles Saver.Tick

        SplashScreenManager1.CloseWaitForm()
        Saver.Enabled = False



        Dim msg = "The excel scorecard was successfully saved to your QA2 folder; would you like to email the scorecard to the the agent?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)


        ''Master view
        If lblDecider.Text = "1" Then

            '' Send email based on supervisor or super user
            If responce = MsgBoxResult.Yes Then

                SplashScreenManager2.ShowWaitForm()
                ProgressBar1.Value = 0

                EmailBackground.RunWorkerAsync()


                SendEmail()


                '   SenderEmail1.Enabled = True





            Else

                reset()

                buttonEnables()




            End If

            '' Supervisor View
        ElseIf lblDecider.Text = "2" Then


            If responce = MsgBoxResult.Yes Then

                SplashScreenManager2.ShowWaitForm()
                ProgressBar1.Value = 0

                EmailBackground.RunWorkerAsync()


                SendEmail2()


                '   SenderEmail1.Enabled = True



            Else

                reset()

                buttonEnables()

                cboSupervisorbox.Text = "Agent Name"



            End If





        End If






        Saver.Enabled = False


        Me.Cursor = Cursors.Hand



    End Sub

    Private Sub Saver2_Tick(sender As Object, e As EventArgs) Handles Saver2.Tick


        SplashScreenManager1.CloseWaitForm()
        Saver2.Enabled = False



        Dim msg = "The excel scorecard was successfully saved to your QA2 folder; would you like to email the scorecard to the the agent?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)

        ''
        If lblDecider.Text = "2" Then


            If responce = MsgBoxResult.Yes Then

                SplashScreenManager2.ShowWaitForm()

                ProgressBar1.Value = 0

                EmailBackground.RunWorkerAsync()



                SendEmail2a()
            Else

                reset()

                buttonEnables()

                cboSupervisorbox.Text = "Agent Name"


            End If
            ''
        ElseIf lblDecider.Text = "1" Then




            If responce = MsgBoxResult.Yes Then

                SplashScreenManager2.ShowWaitForm()

                ProgressBar1.Value = 0

                EmailBackground.RunWorkerAsync()


                SendEmail1a()
            Else

                reset()

                buttonEnables()



            End If



        End If




        Saver2.Enabled = False
    End Sub

    Private Sub ExcelSaver_Tick(sender As Object, e As EventArgs) Handles ExcelSaver.Tick



        Try
            SplashScreenManager1.CloseWaitForm()

            ExcelSaver.Enabled = False


            Dim msg = "The audit was successfully saved; would you like to create the excel versioned scorecard now?"

            Dim title = "FADV QA Application"

            Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

            Dim responce = MsgBox(msg, style, title)





            If responce = MsgBoxResult.Yes Then

                SplashScreenManager1.ShowWaitForm()
                Me.Cursor = Cursors.WaitCursor

                ProgressBar1.Value = 0
                lblprogr.Text = 0


                BackgroundWorker1.RunWorkerAsync()




                QAExcell()




            Else


                QAChatEnable()

                buttonEnables()

                reset()




            End If

            ExcelSaver.Enabled = False

        Catch ex As Exception


            SplashScreenManager1.CloseWaitForm()
            ProgressBar1.Value = 0
            lblprogr.Text = 0

            QAChatEnable()

            buttonEnables()

            MsgBox(ex.Message)


            ExcelSaver.Enabled = False

        End Try







    End Sub

    Private Sub ExcelSaver2_Tick(sender As Object, e As EventArgs) Handles ExcelSaver2.Tick



        Try

            SplashScreenManager1.CloseWaitForm()

            ExcelSaver2.Enabled = False


            Dim msg = "The audit was successfully saved; would you like to create the excel versioned scorecard now?"

            Dim title = "FADV QA Application"

            Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

            Dim responce = MsgBox(msg, style, title)





            If responce = MsgBoxResult.Yes Then

                SplashScreenManager1.ShowWaitForm()
                Me.Cursor = Cursors.WaitCursor

                ProgressBar1.Value = 0
                lblprogr.Text = 0


                BackgroundWorker2.RunWorkerAsync()



                QAExcel2()




            Else

                QAChatEnable()

                buttonEnables()

                reset()



            End If

            ExcelSaver2.Enabled = False

        Catch ex As Exception
            SplashScreenManager1.CloseWaitForm()

            ProgressBar1.Value = 0
            lblprogr.Text = 0

            QAChatEnable()

            buttonEnables()

            MsgBox(ex.Message)


            ExcelSaver2.Enabled = False

        End Try












    End Sub

    Private Sub BackgroundWorker8_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker8.DoWork



        Try


            For i = 0 To 100

                System.Threading.Thread.Sleep(50)
                Me.BackgroundWorker1.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i
            Next




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try







    End Sub

    Private Sub BackgroundWorker8_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker8.ProgressChanged





        ProgressBar1.Value = e.ProgressPercentage


    End Sub

    Private Sub BackgroundWorker8_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker8.RunWorkerCompleted



        Me.Cursor = Cursors.Hand




    End Sub

    Private Sub QAChatScorecard_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown


        If e.Control And e.KeyCode.ToString = "S" Then

            MissedWeightsCalc()
            Me.ActiveControl = txtSR


            '' alt + s saver

            If lblDecider.Text = "1" Then



                If txtSR.Text <> "1-" And txtSR.MaskFull = False Then


                    MsgBox("Please enter a valid SR#")


                Else




                    If txtSR.Text = "1-" And txtContactID.Text = "" Then

                        MsgBox("A Service Request # or Contact ID is required before saving", MessageBoxButtons.RetryCancel)


                    Else





                        If cboAgentName.Text = "Agent Name" Then


                            MsgBox("Please be advised you must select an 'agent name' before proceeding", MessageBoxButtons.RetryCancel)


                        Else

                            If cboSRType.Text = "" Then

                                MsgBox("A SR Type must be selected before saving", MessageBoxButtons.RetryCancel)
                                Me.ActiveControl = cboSRType

                            Else

                                'If dtpCondate.Value = Today Then




                                '    MsgBox("Are you sure the Contact date for this Audit is Today?", MessageBoxButtons.RetryCancel)


                                'Else


                                If txtTeamName.Text = "Please wait, Loading.." Then




                                    MsgBox("The agent’s team field is still loading, please wait until a team name appears before saving the scorecard", MessageBoxButtons.RetryCancel)


                                    Me.ActiveControl = txtTeamName


                                Else




                                    If cboAutoFail.Checked = True And cboAF.Text = "" Then


                                        MsgBox("Since this Audit was marked as 'Auto Fail', a reason must be selected before saving.", MessageBoxButtons.RetryCancel)



                                        Me.ActiveControl = cboAF


                                    Else

                                        If cboCSAT1.Text = "" Or cboCSAT2.Text = "" Or cboCSAT3.Text = "" Or cboCSAT4.Text = "" Or cboCSAT5.Text = "" Or cboCSAT6.Text = "" Then


                                            MsgBox("Please be advised you must fill out the CSAT Equivalency section below before you proceed", MessageBoxButtons.RetryCancel)

                                            Me.ActiveControl = cboCSAT1

                                        Else

                                            CSatScore()

                                            QaTotalScore()


                                            TCXscore()

                                            lblTCXscore.Visible = False




                                            If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then


                                                SplashScreenManager1.ShowWaitForm()
                                                Me.Cursor = Cursors.WaitCursor

                                                buttondisables()

                                                QAChatdisableControls()



                                                Me.ActiveControl = txtSR




                                                BackgroundWorker1.RunWorkerAsync()

                                                'PleaseWait.ShowDialog()

                                                store()

                                            Else




                                                '' Do Nothing



                                            End If


                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If



            If lblDecider.Text = "2" Then




                If txtSR.Text <> "1-" And txtSR.MaskFull = False Then


                    MsgBox("Please enter a valid SR#")


                Else




                    If txtSR.Text = "1-" And txtContactID.Text = "" Then

                        MsgBox("A Service Request # or Contact ID is required before saving", MessageBoxButtons.RetryCancel)


                    Else



                        If cboSupervisorbox.Text = "Agent Name" Then


                            MsgBox("Please be advised you must select an 'agent name' before proceeding", MessageBoxButtons.RetryCancel)

                            Me.ActiveControl = cboSupervisorbox


                        Else

                            If cboSRType.Text = "" Then

                                MsgBox("A SR Type must be selected before saving", MessageBoxButtons.RetryCancel)
                                Me.ActiveControl = cboSRType

                            Else

                                'If dtpCondate.Value = Today Then




                                '    MsgBox("Are you sure the Contact date for this Audit is Today?", MessageBoxButtons.RetryCancel)


                                ' Else


                                If txtTeamName.Text = "Please wait, Loading.." Then




                                    MsgBox("The agent’s team field is still loading, please wait until a team name appears before saving the scorecard", MessageBoxButtons.RetryCancel)


                                    Me.ActiveControl = txtTeamName


                                Else




                                    If cboAutoFail.Checked = True And cboAF.Text = "" Then


                                        MsgBox("Since this Audit was marked as 'Auto Fail', a reason must be selected before saving.", MessageBoxButtons.RetryCancel)



                                        Me.ActiveControl = cboAF


                                    Else

                                        If cboCSAT1.Text = "" Or cboCSAT2.Text = "" Or cboCSAT3.Text = "" Or cboCSAT4.Text = "" Or cboCSAT5.Text = "" Or cboCSAT6.Text = "" Then


                                            MsgBox("Please be advised you must fill out the CSAT Equivalency section below before you proceed", MessageBoxButtons.RetryCancel)

                                            Me.ActiveControl = cboCSAT1

                                        Else

                                            CSatScore()

                                            QaTotalScore()

                                            TCXscore()

                                            lblTCXscore.Visible = False


                                            'If txtRevCom.Text = "" Then




                                            '    If MsgBox("Would you like to review the scorecard now before saving?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then



                                            '        Me.ActiveControl = txtRevCom

                                            '    Else



                                            '        Me.Cursor = Cursors.WaitCursor

                                            '        buttondisables()

                                            '        QAChatdisableControls()


                                            '        Me.ActiveControl = txtSR

                                            '        BackgroundWorker2.RunWorkerAsync()

                                            '        store2()

                                            '        ' PleaseWait.ShowDialog()



                                            '    End If

                                            'End If

                                            ''


                                            'If txtRevCom.Text <> String.Empty Then


                                            If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then




                                            Else

                                                SplashScreenManager1.ShowWaitForm()
                                                Me.Cursor = Cursors.WaitCursor

                                                buttondisables()

                                                QAChatdisableControls()




                                                Me.ActiveControl = txtSR


                                                BackgroundWorker2.RunWorkerAsync()

                                                store2()


                                                '    PleaseWait.ShowDialog()




                                                '

                                            End If

                                        End If

                                    End If
                                    '


                                End If


                            End If


                        End If

                    End If

                End If

            End If

        End If

        If e.Control And e.KeyCode.ToString = "X" Then

            SpellChecker2.CheckContainer(Me)



        End If

        If e.Control And e.KeyCode.ToString = "Z" Then



            Dim msg = "Are you sure you want to clear the Scorecard?"

            Dim title = "FADV QA Application"

            Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

            Dim responce = MsgBox(msg, style, title)





            If responce = MsgBoxResult.Yes Then

                reset()



            Else




            End If

        End If







    End Sub





    Private Shared Function Emailer(ByVal sender As Object, ByVal cert As X509Certificate, ByVal chain As X509Chain, ByVal errors As SslPolicyErrors) As Boolean

        Return True

    End Function





    Public Sub SendEmail()

        Try



            ' Dim attachment As Attachment = New Attachment("C:\Users\durraner\Documents\QASpreadSheet.xlsx")


            '   Dim attachment As Attachment = New Attachment(Desk & "\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & "Chat QA Scorecard.xlsx")


            Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & "Chat QA Scorecard.xlsx")



            Dim mail As New MailMessage




            Dim smtp As New SmtpClient("NOAMIND01MXC12.NOAM.FADV.NET")


            mail.Attachments.Add(attachment)


            mail.Subject = "QA Scorecard for SR#:" + txtSR.Text

            mail.To.Add(txtAgentEmail.Text)
            mail.CC.Add("CustomerCareQA@fadv.com")
            mail.CC.Add(lblUserEmail.Text)
            mail.CC.Add(lblSupervisorEmail.Text)





            mail.From = New MailAddress("CustomerCareQA@fadv.com")

            mail.Body = "Hello " + cboAgentName.Text + "," & vbCrLf &
           "" & vbCrLf &
            "I have attached your QA scorecard, if you have any questions or concerns please reach out to your supervisor." & vbCrLf &
            "" & vbCrLf &
            "Thank you," & vbCrLf &
            "QA Team"




           smtp.EnableSsl = False


            smtp.Credentials = New System.Net.NetworkCredential("durraner", Form2.lblEmailPassword.Text)



            smtp.Port = 587

            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf Emailer)



            smtp.Send(mail)

            SendEmailFin.Enabled = True

        Catch ex As Exception
            SplashScreenManager2.CloseWaitForm()

            EmailBackground.CancelAsync()

            SenderEmail1.Enabled = False

            buttonEnables()

            MsgBox(ex.Message)


            SenderEmail1.Enabled = False


            MsgBox(ex.Message)

        End Try




    End Sub


    Public Sub SendEmail2()

        Try



            ' Dim attachment As Attachment = New Attachment("C:\Users\durraner\Documents\QASpreadSheet.xlsx")


            '  Dim attachment As Attachment = New Attachment(Desk & "\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & "Chat QA Scorecard.xlsx")


            Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "Chat QA Scorecard.xlsx")



            Dim mail As New MailMessage




            Dim smtp As New SmtpClient("NOAMIND01MXC12.NOAM.FADV.NET")


            mail.Attachments.Add(attachment)


            mail.Subject = "QA Scorecard for SR#:" + txtSR.Text

            mail.To.Add(txtAgentEmail.Text)
            mail.CC.Add("CustomerCareQA@fadv.com")
            mail.CC.Add(lblUserEmail.Text)


            mail.From = New MailAddress("CustomerCareQA@fadv.com")


            mail.Body = "Hello " + cboSupervisorbox.Text + "," & vbCrLf &
           "" & vbCrLf &
            "I have attached your QA scorecard, if you have any questions or concerns please reach out to your supervisor." & vbCrLf &
            "" & vbCrLf &
            "Thank you," & vbCrLf &
            "QA Team"




           smtp.EnableSsl = False


            smtp.Credentials = New System.Net.NetworkCredential("durraner", Form2.lblEmailPassword.Text)



            smtp.Port = 587

            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf Emailer)



            smtp.Send(mail)

            SendEmailFin.Enabled = True

        Catch ex As Exception

            SplashScreenManager2.CloseWaitForm()

            MsgBox(ex.Message)



            EmailBackground.CancelAsync()

            SenderEmail2.Enabled = False

            buttonEnables()

            MsgBox(ex.Message)



        End Try




    End Sub


    Public Sub SendEmail1a()

        Try



            ' Dim attachment As Attachment = New Attachment("C:\Users\durraner\Documents\QASpreadSheet.xlsx")


            '  Dim attachment As Attachment = New Attachment(Desk & "\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & "Chat QA Scorecard.xlsx")


            Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & "Chat QA Scorecard.xlsx")

            '   Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & Desk & "\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "Chat QA Scorecard.xlsx")





            Dim mail As New MailMessage




            Dim smtp As New SmtpClient("NOAMIND01MXC12.NOAM.FADV.NET")


            mail.Attachments.Add(attachment)


            mail.Subject = "QA Scorecard for SR#:" + txtContactID.Text

            mail.To.Add(txtAgentEmail.Text)
            mail.CC.Add("CustomerCareQA@fadv.com")
            mail.CC.Add(lblUserEmail.Text)
            mail.CC.Add(lblSupervisorEmail.Text)





            mail.From = New MailAddress("CustomerCareQA@fadv.com")


            mail.Body = "Hello " + cboAgentName.Text + "," & vbCrLf &
           "" & vbCrLf &
            "I have attached your QA scorecard, if you have any questions or concerns please reach out to your supervisor." & vbCrLf &
            "" & vbCrLf &
            "Thank you," & vbCrLf &
            "QA Team"



           smtp.EnableSsl = False


            smtp.Credentials = New System.Net.NetworkCredential("durraner", Form2.lblEmailPassword.Text)



            smtp.Port = 587

            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf Emailer)



            smtp.Send(mail)



            SendEmailFin.Enabled = True



        Catch ex As Exception

            SplashScreenManager2.CloseWaitForm()

            EmailBackground.CancelAsync()

            SenderEmail1.Enabled = False

            buttonEnables()

            MsgBox(ex.Message)


            SenderEmail1.Enabled = False

        End Try




    End Sub


    Public Sub SendEmail2a()

        Try



            ' Dim attachment As Attachment = New Attachment("C:\Users\durraner\Documents\QASpreadSheet.xlsx")


            '  Dim attachment As Attachment = New Attachment(Desk & "\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & "Chat QA Scorecard.xlsx")

            Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtContactID.Text & " " & cboSupervisorbox.Text & "-" & "Chat QA Scorecard.xlsx")

            '    Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & Desk & "\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "Chat QA Scorecard.xlsx")


            Dim mail As New MailMessage




            Dim smtp As New SmtpClient("NOAMIND01MXC12.NOAM.FADV.NET")


            mail.Attachments.Add(attachment)


            mail.Subject = "QA Scorecard for SR#:" + txtContactID.Text

            mail.To.Add(txtAgentEmail.Text)
            mail.CC.Add("CustomerCareQA@fadv.com")
            mail.CC.Add(lblUserEmail.Text)


            mail.From = New MailAddress("CustomerCareQA@fadv.com")


            mail.Body = "Hello " + cboSupervisorbox.Text + "," & vbCrLf &
           "" & vbCrLf &
            "I have attached your QA scorecard, if you have any questions or concerns please reach out to your supervisor." & vbCrLf &
            "" & vbCrLf &
            "Thank you," & vbCrLf &
            "QA Team"





           smtp.EnableSsl = False


            smtp.Credentials = New System.Net.NetworkCredential("durraner", Form2.lblEmailPassword.Text)



            smtp.Port = 587

            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf Emailer)



            smtp.Send(mail)

            SendEmailFin.Enabled = True


        Catch ex As Exception

            SplashScreenManager2.CloseWaitForm()

            MsgBox(ex.Message)


            EmailBackground.CancelAsync()

            SenderEmail2.Enabled = False

            buttonEnables()

            MsgBox(ex.Message)





        End Try




    End Sub



    Private Sub SpellCheckLoadTimer_Tick(sender As Object, e As EventArgs) Handles SpellCheckLoadTimer.Tick



        SpellChecker2.ParentContainer = Me
        SpellChecker2.CheckAsYouTypeOptions.CheckControlsInParentContainer = True
        SpellChecker2.SpellCheckMode = SpellCheckMode.AsYouType


        SpellCheckLoadTimer.Enabled = False

    End Sub

    Private Sub EmailBackground_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles EmailBackground.DoWork

        Dim EmailBackground As System.ComponentModel.BackgroundWorker = CType(sender, System.ComponentModel.BackgroundWorker)

        For i = 0 To 100

            System.Threading.Thread.Sleep(40)

            If EmailBackground.CancellationPending Then

                e.Cancel = True

            Else

                Me.EmailBackground.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i


            End If


        Next



    End Sub

    Private Sub EmailBackground_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles EmailBackground.ProgressChanged

        ProgressBar1.Value = e.ProgressPercentage


    End Sub

    Private Sub EmailBackground_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles EmailBackground.RunWorkerCompleted






    End Sub

    Private Sub SenderEmail1_Tick(sender As Object, e As EventArgs) Handles SenderEmail1.Tick


        SendEmail()




        SenderEmail1.Enabled = False




    End Sub

    Private Sub SenderEmail2_Tick(sender As Object, e As EventArgs) Handles SenderEmail2.Tick





        SendEmail2()




        SenderEmail2.Enabled = False



    End Sub

    Private Sub cboAutoFail_CheckedChanged(sender As Object, e As EventArgs) Handles cboAutoFail.CheckedChanged


        Try


            If cboAutoFail.Checked = True Then





                cboAF.Visible = True

                lblQAScore1.Visible = False


                lblAutoZero.Visible = True



            ElseIf cboAutoFail.Checked = False Then


                cboAF.Visible = False

                cboAF.Text = "N/a"

                lblAutoZero.Visible = False

                txtQAScore.Text = lblQAScore1.Text

                '   lblQAScore1.Visible = True



            End If


        Catch ex As Exception

        End Try


    End Sub

    Private Sub SendEmailFin_Tick(sender As Object, e As EventArgs) Handles SendEmailFin.Tick

        SendEmailFin.Enabled = False

        Cursor = Cursors.Hand

        SplashScreenManager2.CloseWaitForm()


        Dim msg = "The scorecard was successfully emailed to the agent, would you like to audit a new chat?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)





        If responce = MsgBoxResult.Yes Then


            buttonEnables()

            reset()



        Else

            reset()

            buttonEnables()


            Form2.Clear()


            Form2.Show()

            Me.Hide()



        End If




        SendEmailFin.Enabled = False




    End Sub

    Public Sub FillAutoFail()

        Try

            QaSetupMod.connecttemp17()

            sqltemp17 = "SELECT * FROM [AutoFail]"



            Dim cmdtemp As New SqlClient.SqlCommand




            cmdtemp.CommandText = sqltemp17

            cmdtemp.Connection = contemp17



            readertemp17 = cmdtemp.ExecuteReader


            While (readertemp17.Read())


                cboAF.Items.Add(readertemp17("AutoFailReason"))


            End While



            cmdtemp.Dispose()

            contemp17.Close()

            Me.Cursor = Cursors.Hand




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try





    End Sub


End Class