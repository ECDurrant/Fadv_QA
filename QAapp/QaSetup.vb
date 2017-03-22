
Imports System.Data.OleDb
Imports System.Data.SqlClient

Imports System.Threading


Public Class QaSetup





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnHide.Click


        Me.Hide()



    End Sub

    Private Sub QaSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try



            QaSetupMod.connecttemp1()
            Fillcombo()

            QaSetupMod.connecttemp2()


            Me.ActiveControl = cboContactType

            Me.CenterToScreen()

            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "MM/dd/yyyy"






        Catch ex As Exception



            MsgBox(ex.Message)

        End Try






    End Sub


    Public Sub Fillcombo()


        Try





            sqltemp1 = "SELECT * FROM [Agents] WHERE Supervisor='" & lblQAauditor.Text & "' "





            Dim cmdtemp As New SqlCommand



            '  cmdtemp.CommandText = sqltemp

            cmdtemp.CommandText = sqltemp1

            cmdtemp.Connection = contemp1





            readertemp1 = cmdtemp.ExecuteReader



            While (readertemp1.Read())



                cboAgentName.Items.Add(readertemp1("AgentName"))




            End While






            cmdtemp.Dispose()

            readertemp1.Close()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try






    End Sub












    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Try

            If cboAgentName.Text = "Agent Name" Or cboContactType.Text = "Contact Type" Then






                MsgBox("Please be advised you must fill out all 'Agent Information' before proceeding", MessageBoxButtons.RetryCancel)


            Else



                '   If txtContactID.Text = "" And txtSRNumber.Text = "" Then

                'MsgBox("Please be advised a ContactID or SR# is required before saving", MessageBoxButtons.RetryCancel)


                '  Else

                If QACallScorecard.lblQAScore1.Visible = True Or QAEmailScorecard.lblQAScore1.Visible = True Or QAChatScorecard.lblQAScore1.Visible = True Or QALvl2CallScorecard.lblQAScore1.Visible = True Or QAlvl2EmailScorecard.lblQAScore.Visible = True Or QAResCallScorecard.lblQAScore1.Visible = True Or QAResidentEmailScorecard.lblQAScore.Visible = True Or QAConsuACallScorcard.lblQAScore1.Visible = True Then



                    MsgBox("You can not save a scorecard that has been scored already, press 'clear fields' button", MessageBoxButtons.OK)



                Else



                    If MsgBox("Are you sure you want to save and continue to the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then


                    Else



                        'Disable Controls



                        txtContactID.Enabled = False
                        txtContactEmail.Enabled = False
                        txtContactName.Enabled = False
                        txtContactPhone.Enabled = False
                        txtSRNumber.Enabled = False
                        txtOrderID.Enabled = False
                        txtAccountNum.Enabled = False
                        txtCompany.Enabled = False
                        txtJIRAbox.Enabled = False
                        txtUserID.Enabled = False
                        DateTimePicker1.Enabled = False


                        cboAgentName.Enabled = False

                        cboContactType.Enabled = False



                        '  QACallScorecard.lblDrive.Text = ComboBox1.Text




                        ' Select Scorcard based on Qa Audit Type

                        If cboContactType.Text = "Call" Then

                            Dim stringg As String = DateTimePicker1.Text


                            QACallScorecard.txtgnamebox.Text = txtContactName.Text
                            QACallScorecard.txtgphone.Text = txtContactPhone.Text
                            QACallScorecard.txtgemail.Text = txtContactEmail.Text
                            QACallScorecard.txtgacc.Text = txtAccountNum.Text
                            QACallScorecard.txtgcompany.Text = txtCompany.Text
                            QACallScorecard.txtgorderid.Text = txtOrderID.Text
                            QACallScorecard.txtgDatebox.Text = stringg
                            QACallScorecard.lbldrive2.Text = lblMDrive.Text
                            QACallScorecard.lblSCRN.Text = lblSCRN.Text




                            ''Call
                            '  QACallScorecard.lblAgentName1.Text = cboAgentName.Text
                            ' QACallScorecard.lblAgentTeam1.Text = txtAgentTeam.Text
                            '  QACallScorecard.lblContactType1.Text = cboContactType.Text
                            '  QACallScorecard.lblSRNumber1.Text = txtSRNumber.Text
                            '  QACallScorecard.lblContactID1.Text = txtContactID.Text


                            QACallScorecard.lblQAauditor1.Text = lblQAauditor.Text




                            QAEmailScorecard.Hide()
                            QAChatScorecard.Hide()
                            QALvl2CallScorecard.Hide()
                            QAlvl2EmailScorecard.Hide()
                            QAResCallScorecard.Hide()
                            QAResidentEmailScorecard.Hide()
                            QAConsuACallScorcard.Hide()



                            QACallScorecard.Show()

                            Me.Hide()



                        ElseIf cboContactType.Text = "Email" Then

                            Dim stringg As String = DateTimePicker1.Text


                            QAEmailScorecard.txtgnamebox.Text = txtContactName.Text
                            QAEmailScorecard.txtgphone.Text = txtContactPhone.Text
                            QAEmailScorecard.txtgemail.Text = txtContactEmail.Text
                            QAEmailScorecard.txtgacc.Text = txtAccountNum.Text
                            QAEmailScorecard.txtgcompany.Text = txtCompany.Text
                            QAEmailScorecard.txtgorderid.Text = txtOrderID.Text
                            QAEmailScorecard.txtgDatebox.Text = stringg
                            QAEmailScorecard.lbldrive2.Text = lblMDrive.Text
                            QAEmailScorecard.lblSCRN.Text = lblSCRN.Text



                            '''Email

                            'QAEmailScorecard.lblAgentName.Text = cboAgentName.Text
                            'QAEmailScorecard.lblAgentTeam.Text = txtAgentTeam.Text
                            'QAEmailScorecard.lblContactType.Text = cboContactType.Text
                            'QAEmailScorecard.lblSRNumber.Text = txtSRNumber.Text
                            'QAEmailScorecard.lblContactID1.Text = txtContactID.Text

                            'QAEmailScorecard.lblQAauditor.Text = lblQAauditor.Text


                            QAEmailScorecard.lbldrive2.Text = lblMDrive.Text
                            QAEmailScorecard.lblSCRN.Text = lblSCRN.Text


                            QACallScorecard.Hide()
                            QAChatScorecard.Hide()
                            QALvl2CallScorecard.Hide()
                            QAlvl2EmailScorecard.Hide()
                            QAResCallScorecard.Hide()
                            QAResidentEmailScorecard.Hide()
                            QAConsuACallScorcard.Hide()



                            QAEmailScorecard.Show()

                            Me.Hide()



                        ElseIf cboContactType.Text = "Chat" Then

                            Dim stringg As String = DateTimePicker1.Text


                            QAChatScorecard.txtgnamebox.Text = txtContactName.Text
                            QAChatScorecard.txtgphone.Text = txtContactPhone.Text
                            QAChatScorecard.txtgemail.Text = txtContactEmail.Text
                            QAChatScorecard.txtgacc.Text = txtAccountNum.Text
                            QAChatScorecard.txtgcompany.Text = txtCompany.Text
                            QAChatScorecard.txtgorderid.Text = txtOrderID.Text
                            QAChatScorecard.txtgDatebox.Text = stringg
                            QAChatScorecard.lbldrive2.Text = lblMDrive.Text
                            QAChatScorecard.lblSCRN1.Text = lblSCRN.Text






                            ''Chat
                            'QAChatScorecard.lblAgentName1.Text = cboAgentName.Text
                            'QAChatScorecard.lblAgentTeam1.Text = txtAgentTeam.Text
                            'QAChatScorecard.lblContactType1.Text = cboContactType.Text
                            'QAChatScorecard.lblSRNumber1.Text = txtSRNumber.Text
                            'QAChatScorecard.lblContactID1.Text = txtContactID.Text
                            'QAChatScorecard.lblQAauditor1.Text = lblQAauditor.Text




                            QAChatScorecard.lbldrive2.Text = lblMDrive.Text
                            QAChatScorecard.lblSCRN1.Text = lblSCRN.Text

                            QACallScorecard.Hide()
                            QAEmailScorecard.Hide()
                            QALvl2CallScorecard.Hide()
                            QAlvl2EmailScorecard.Hide()
                            QAResCallScorecard.Hide()
                            QAResidentEmailScorecard.Hide()
                            QAConsuACallScorcard.Hide()


                            QAChatScorecard.Show()
                            Me.Hide()






                        ElseIf cboContactType.Text = "Level 2 - Call" Then



                            Dim stringg As String = DateTimePicker1.Text


                            QALvl2CallScorecard.txtgnamebox.Text = txtContactName.Text
                            QALvl2CallScorecard.txtgphone.Text = txtContactPhone.Text
                            QALvl2CallScorecard.txtgemail.Text = txtContactEmail.Text
                            QALvl2CallScorecard.txtgacc.Text = txtAccountNum.Text
                            QALvl2CallScorecard.txtgcompany.Text = txtCompany.Text
                            QALvl2CallScorecard.txtgorderid.Text = txtOrderID.Text
                            QALvl2CallScorecard.txtgDatebox.Text = stringg
                            QALvl2CallScorecard.lbldrive2.Text = lblMDrive.Text
                            QALvl2CallScorecard.lblSCRN.Text = lblSCRN.Text

                            QALvl2CallScorecard.txtgjira.Text = txtJIRAbox.Text
                            QALvl2CallScorecard.txtguser.Text = txtUserID.Text




                            ''lvl 2 chat
                            QALvl2CallScorecard.lblAgentName1.Text = cboAgentName.Text
                            QALvl2CallScorecard.lblAgentTeam1.Text = txtAgentTeam.Text
                            QALvl2CallScorecard.lblContactType1.Text = cboContactType.Text
                            QALvl2CallScorecard.lblSRNumber1.Text = txtSRNumber.Text
                            QALvl2CallScorecard.lblContactID1.Text = txtContactID.Text
                            QALvl2CallScorecard.lblQAauditor1.Text = lblQAauditor.Text
                            QALvl2CallScorecard.lblJIRA.Text = txtJIRAbox.Text
                            QALvl2CallScorecard.lblUserID.Text = txtUserID.Text


                            QALvl2CallScorecard.lbldrive2.Text = lblMDrive.Text
                            QALvl2CallScorecard.lblSCRN.Text = lblSCRN.Text


                            QACallScorecard.Hide()
                            QAEmailScorecard.Hide()
                            QAChatScorecard.Hide()
                            QAlvl2EmailScorecard.Hide()
                            QAResCallScorecard.Hide()
                            QAResidentEmailScorecard.Hide()
                            QAConsuACallScorcard.Hide()




                            QALvl2CallScorecard.Show()

                            Me.Hide()



                        ElseIf cboContactType.Text = "Level 2 - Email" Then


                            Dim stringg As String = DateTimePicker1.Text


                            QAlvl2EmailScorecard.txtgnamebox.Text = txtContactName.Text
                            QAlvl2EmailScorecard.txtgphone.Text = txtContactPhone.Text
                            QAlvl2EmailScorecard.txtgemail.Text = txtContactEmail.Text
                            QAlvl2EmailScorecard.txtgacc.Text = txtAccountNum.Text
                            QAlvl2EmailScorecard.txtgcompany.Text = txtCompany.Text
                            QAlvl2EmailScorecard.txtgorderid.Text = txtOrderID.Text
                            QAlvl2EmailScorecard.txtgDatebox.Text = stringg
                            QAlvl2EmailScorecard.lbldrive2.Text = lblMDrive.Text
                            QAlvl2EmailScorecard.lblSCRN.Text = lblSCRN.Text

                            QAlvl2EmailScorecard.txtgjira.Text = txtJIRAbox.Text
                            QAlvl2EmailScorecard.txtguser.Text = txtUserID.Text




                            ''lvl 2 chat
                            QAlvl2EmailScorecard.lblAgentName.Text = cboAgentName.Text
                            QAlvl2EmailScorecard.lblAgentTeam.Text = txtAgentTeam.Text
                            QAlvl2EmailScorecard.lblContactType.Text = cboContactType.Text
                            QAlvl2EmailScorecard.lblSRNumber.Text = txtSRNumber.Text
                            QAlvl2EmailScorecard.lblContactID1.Text = txtContactID.Text
                            QAlvl2EmailScorecard.lblQAauditor.Text = lblQAauditor.Text
                            QAlvl2EmailScorecard.lblJIRA.Text = txtJIRAbox.Text
                            QAlvl2EmailScorecard.lblUserID.Text = txtUserID.Text


                            QAlvl2EmailScorecard.lbldrive2.Text = lblMDrive.Text
                            QAlvl2EmailScorecard.lblSCRN.Text = lblSCRN.Text



                            QACallScorecard.Hide()
                            QAEmailScorecard.Hide()
                            QAChatScorecard.Hide()
                            QALvl2CallScorecard.Hide()
                            QAResCallScorecard.Hide()
                            QAResidentEmailScorecard.Hide()
                            QAConsuACallScorcard.Hide()




                            QAlvl2EmailScorecard.Show()


                            Me.Hide()




                        ElseIf cboContactType.Text = "Resident - Call" Then

                            Dim stringg As String = DateTimePicker1.Text

                            QAResCallScorecard.txtgnamebox.Text = txtContactName.Text
                            QAResCallScorecard.txtgphone.Text = txtContactPhone.Text
                            QAResCallScorecard.txtgemail.Text = txtContactEmail.Text
                            QAResCallScorecard.txtgacc.Text = txtAccountNum.Text
                            QAResCallScorecard.txtgcompany.Text = txtCompany.Text
                            QAResCallScorecard.txtgorderid.Text = txtOrderID.Text
                            QAResCallScorecard.txtgDatebox.Text = stringg
                            QAResCallScorecard.lbldrive2.Text = lblMDrive.Text
                            QAResCallScorecard.lblSCRN.Text = lblSCRN.Text




                            QAResCallScorecard.lblAgentName1.Text = cboAgentName.Text
                            QAResCallScorecard.lblAgentTeam1.Text = txtAgentTeam.Text
                            QAResCallScorecard.lblContactType1.Text = cboContactType.Text
                            QAResCallScorecard.lblSRNumber1.Text = txtSRNumber.Text
                            QAResCallScorecard.lblContactID1.Text = txtContactID.Text
                            QAResCallScorecard.lblQAauditor1.Text = lblQAauditor.Text
                            '  QAResCallScorecard.lblJIRA.Text = txtJIRAbox.Text
                            ' QAResCallScorecard.lblUserID.Text = txtUserID.Text


                            QAResCallScorecard.lbldrive2.Text = lblMDrive.Text
                            QAResCallScorecard.lblSCRN.Text = lblSCRN.Text


                            QACallScorecard.Hide()
                            QAEmailScorecard.Hide()
                            QAChatScorecard.Hide()
                            QALvl2CallScorecard.Hide()
                            QAlvl2EmailScorecard.Hide()
                            QAResidentEmailScorecard.Hide()
                            QAConsuACallScorcard.Hide()



                            QAResCallScorecard.Show()


                            Me.Hide()




                        ElseIf cboContactType.Text = "Resident - Email" Then


                            Dim stringg As String = DateTimePicker1.Text

                            QAResidentEmailScorecard.txtgnamebox.Text = txtContactName.Text
                            QAResidentEmailScorecard.txtgphone.Text = txtContactPhone.Text
                            QAResidentEmailScorecard.txtgemail.Text = txtContactEmail.Text
                            QAResidentEmailScorecard.txtgacc.Text = txtAccountNum.Text
                            QAResidentEmailScorecard.txtgcompany.Text = txtCompany.Text
                            QAResidentEmailScorecard.txtgorderid.Text = txtOrderID.Text
                            QAResidentEmailScorecard.txtgDatebox.Text = stringg
                            QAResidentEmailScorecard.lbldrive2.Text = lblMDrive.Text
                            QAResidentEmailScorecard.lblSCRN.Text = lblSCRN.Text



                            QAResidentEmailScorecard.lblAgentName.Text = cboAgentName.Text
                            QAResidentEmailScorecard.lblAgentTeam.Text = txtAgentTeam.Text
                            QAResidentEmailScorecard.lblContactType.Text = cboContactType.Text
                            QAResidentEmailScorecard.lblSRNumber.Text = txtSRNumber.Text
                            QAResidentEmailScorecard.lblContactID1.Text = txtContactID.Text
                            QAResidentEmailScorecard.lblQAauditor.Text = lblQAauditor.Text
                            '  QAResCallScorecard.lblJIRA.Text = txtJIRAbox.Text
                            ' QAResCallScorecard.lblUserID.Text = txtUserID.Text



                            QAResCallScorecard.lbldrive2.Text = lblMDrive.Text
                            QAResCallScorecard.lblSCRN.Text = lblSCRN.Text


                            QACallScorecard.Hide()
                            QAEmailScorecard.Hide()
                            QAChatScorecard.Hide()
                            QALvl2CallScorecard.Hide()
                            QAlvl2EmailScorecard.Hide()
                            QAResCallScorecard.Hide()
                            QAConsuACallScorcard.Hide()





                            QAResidentEmailScorecard.Show()


                            Me.Hide()



                        ElseIf cboContactType.Text = "Consumer Advocacy - Call" Then



                            Dim stringg As String = DateTimePicker1.Text


                            QAConsuACallScorcard.txtgnamebox.Text = txtContactName.Text
                            QAConsuACallScorcard.txtgphone.Text = txtContactPhone.Text
                            QAConsuACallScorcard.txtgemail.Text = txtContactEmail.Text
                            QAConsuACallScorcard.txtgacc.Text = txtAccountNum.Text
                            QAConsuACallScorcard.txtgcompany.Text = txtCompany.Text
                            QAConsuACallScorcard.txtgorderid.Text = txtOrderID.Text
                            QAConsuACallScorcard.txtgDatebox.Text = stringg
                            QAConsuACallScorcard.lbldrive2.Text = lblMDrive.Text
                            QAConsuACallScorcard.lblSCRN.Text = lblSCRN.Text






                            QAConsuACallScorcard.lblAgentName1.Text = cboAgentName.Text
                            QAConsuACallScorcard.lblAgentTeam1.Text = txtAgentTeam.Text
                            QAConsuACallScorcard.lblContactType1.Text = cboContactType.Text
                            QAConsuACallScorcard.lblSRNumber1.Text = txtSRNumber.Text
                            QAConsuACallScorcard.lblContactID1.Text = txtContactID.Text
                            QAConsuACallScorcard.lblQAauditor1.Text = lblQAauditor.Text
                            '  QAResCallScorecard.lblJIRA.Text = txtJIRAbox.Text
                            ' QAResCallScorecard.lblUserID.Text = txtUserID.Text



                            QAConsuACallScorcard.lbldrive2.Text = lblMDrive.Text
                            QAConsuACallScorcard.lblSCRN.Text = lblSCRN.Text


                            QACallScorecard.Hide()
                            QAEmailScorecard.Hide()
                            QAChatScorecard.Hide()
                            QALvl2CallScorecard.Hide()
                            QAlvl2EmailScorecard.Hide()
                            QAResCallScorecard.Hide()
                            QAResidentEmailScorecard.Hide()







                            QAConsuACallScorcard.Show()



                            Me.Hide()










                        End If




                    End If



                End If

            End If


            '  End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub





    Private Sub Button1_Click_1(sender As Object, e As EventArgs)

        Try



            'Transfer label names to QAscorecard form

            QAScorecard.lblAgentName.Text = cboAgentName.Text
            QAScorecard.lblAgentTeam.Text = txtAgentTeam.Text
            QAScorecard.lblContactType.Text = cboContactType.Text
            QAScorecard.lblSRNumber.Text = txtSRNumber.Text



            txtContactID.Enabled = False
            txtContactEmail.Enabled = False
            txtContactName.Enabled = False
            txtContactPhone.Enabled = False
            txtSRNumber.Enabled = False
            txtOrderID.Enabled = False
            txtAccountNum.Enabled = False
            txtCompany.Enabled = False
            DateTimePicker1.Enabled = False


            cboAgentName.Enabled = False
            txtAgentTeam.Enabled = False
            cboContactType.Enabled = False

            btnSave.Enabled = False


            QAScorecard.Show()



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





    End Sub

    Public Sub Clear()

        Try


            txtContactID.Enabled = True
            txtContactEmail.Enabled = True
            txtContactName.Enabled = True
            txtContactPhone.Enabled = True
            txtSRNumber.Enabled = True
            txtOrderID.Enabled = True
            txtAccountNum.Enabled = True
            txtCompany.Enabled = True
            txtJIRAbox.Enabled = True
            txtUserID.Enabled = True

            DateTimePicker1.Enabled = True


            cboAgentName.Enabled = True
            txtAgentTeam.Enabled = True
            cboContactType.Enabled = True




            txtContactID.Clear()

            txtContactEmail.Clear()
            txtContactName.Clear()
            txtContactPhone.Clear()
            txtSRNumber.Clear()
            txtOrderID.Clear()
            txtAccountNum.Clear()
            txtCompany.Clear()
            txtJIRAbox.Clear()
            txtUserID.Clear()
            txtAgentTeam.Clear()


            btnSave.Enabled = True


            cboAgentName.Text = "Agent Name"

            cboContactType.Text = "Contact Type"


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub







    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        Try

            If MsgBox("Please be advised you about to clear and reset the scorecard, do you want to continue?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then



            Else



                Clear()



                QACallScorecard.reset()



                QACallScorecard.resetatglance()
                QACallScorecard.QACallclear()
                QACallScorecard.QACallEnable()



                QAEmailScorecard.reset()


                QAEmailScorecard.resetatglance()
                QAEmailScorecard.QAEmailclear()
                QAEmailScorecard.QAEmailEnable()




                QAChatScorecard.reset()


                QAChatScorecard.resetatglance()
                QAChatScorecard.QAChatclear()
                QAChatScorecard.QAChatEnable()




                QALvl2CallScorecard.reset()


                QALvl2CallScorecard.resetatglance()
                QALvl2CallScorecard.QAlvl2Callclear()
                QALvl2CallScorecard.QAlvl2CallEnable()


                QAlvl2EmailScorecard.reset()



                QAlvl2EmailScorecard.resetatglance()
                QAlvl2EmailScorecard.QAlvl2Emailclear()
                QAlvl2EmailScorecard.QAlvl2EmailEnable()


                QAResCallScorecard.reset()


                QAResCallScorecard.resetatglance()
                QAResCallScorecard.QACallclear()
                QAResCallScorecard.QACallEnable()


                QAResidentEmailScorecard.reset()


                QAResidentEmailScorecard.resetatglance()
                QAResidentEmailScorecard.QAEmailclear()
                QAResidentEmailScorecard.QAEmailEnable()


                QAConsuACallScorcard.reset()



                QAConsuACallScorcard.resetatglance()
                QAConsuACallScorcard.QACallclear()
                QAConsuACallScorcard.QACallEnable()







                QACallScorecard.Hide()
                QAEmailScorecard.Hide()
                QAChatScorecard.Hide()
                QALvl2CallScorecard.Hide()
                QAlvl2EmailScorecard.Hide()
                QAResCallScorecard.Hide()
                QAResidentEmailScorecard.Hide()
                QAConsuACallScorcard.Hide()


            End If




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub



    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click



        txtSRNumber.Text = "10-12323233"
        txtContactID.Text = "1212111"
        txtContactName.Text = "Crystal Smith"
        txtContactEmail.Text = "CrystalSmith@Gmail.com"
        txtContactPhone.Text = "5558889695"
        txtAccountNum.Text = "abc32323"
        txtCompany.Text = "Little Leauge"
        txtOrderID.Text = "95955555"
        txtJIRAbox.Text = "45488788"
        txtUserID.Text = "545454user"









    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        txtSRNumber.Text = "12-12323233"
        txtContactID.Text = "625984"
        txtContactName.Text = "Justin Fitzgerald"
        txtContactEmail.Text = "JTSDF@Yahooo.com"
        txtContactPhone.Text = "5559999695"
        txtAccountNum.Text = "abc101222"
        txtCompany.Text = "UPS"
        txtOrderID.Text = "1014477"
        txtJIRAbox.Text = "89844548"
        txtUserID.Text = "98985454user2"











    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click



        txtSRNumber.Text = "13-12323233"
        txtContactID.Text = "7825984"
        txtContactName.Text = "Heather Brown"
        txtContactEmail.Text = "HB@spcglobal.com"
        txtContactPhone.Text = "5559999695"
        txtAccountNum.Text = "bbc701222"
        txtCompany.Text = "Hannford"
        txtOrderID.Text = "11144778"
        txtJIRAbox.Text = "258488788"
        txtUserID.Text = "788454user3"













    End Sub


    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click

        Try

            If txtSRNumber.Enabled = True Then


                MsgBox("Edits can not be made at this time", MessageBoxButtons.OK)


            Else



                If QACallScorecard.lblQAScore1.Visible = True Or QAEmailScorecard.lblQAScore1.Visible = True Then



                    MsgBox("Edits can not be made to Scorecard Info after Audit has been saved", MessageBoxButtons.OK)



                Else




                    If MsgBox("Are you sure you want to make edits to the Scorecard Info?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then




                    Else

                        ''Enable Buttons

                        txtContactID.Enabled = True
                        txtContactEmail.Enabled = True
                        txtContactName.Enabled = True
                        txtContactPhone.Enabled = True
                        txtSRNumber.Enabled = True
                        txtOrderID.Enabled = True
                        txtAccountNum.Enabled = True
                        txtCompany.Enabled = True
                        txtUserID.Enabled = True
                        txtJIRAbox.Enabled = True

                        DateTimePicker1.Enabled = True


                        cboAgentName.Enabled = True
                        txtAgentTeam.Enabled = True
                        cboContactType.Enabled = True




                        'Transfer label names to QAscorecard form


                        'QAEmailScorecard.lblAgentName.Text = cboAgentName.Text
                        'QAEmailScorecard.lblAgentTeam.Text = txtAgentTeam.Text
                        'QAEmailScorecard.lblContactType.Text = cboContactType.Text
                        'QAEmailScorecard.lblSRNumber.Text = txtSRNumber.Text


                        'QAChatScorecard.cboAgentName.Text = cboAgentName.Text
                        'QAChatScorecard.cboTeamName.Text = txtAgentTeam.Text
                        'QAChatScorecard.lblContactType1.Text = cboContactType.Text
                        'QAChatScorecard.txtSR.Text = txtSRNumber.Text

                        'QAEmailScorecard.lblAgentName.Text = cboAgentName.Text
                        'QAEmailScorecard.lblAgentTeam.Text = txtAgentTeam.Text
                        'QAEmailScorecard.lblContactType.Text = cboContactType.Text
                        'QAEmailScorecard.lblSRNumber.Text = txtSRNumber.Text


                        QALvl2CallScorecard.lblAgentName1.Text = cboAgentName.Text
                        QALvl2CallScorecard.lblAgentTeam1.Text = txtAgentTeam.Text
                        QALvl2CallScorecard.lblContactType1.Text = cboContactType.Text
                        QALvl2CallScorecard.lblSRNumber1.Text = txtSRNumber.Text
                        QALvl2CallScorecard.lblUserID.Text = txtUserID.Text
                        QALvl2CallScorecard.lblJIRA.Text = txtJIRAbox.Text




                        QAlvl2EmailScorecard.lblAgentName.Text = cboAgentName.Text
                        QAlvl2EmailScorecard.lblAgentTeam.Text = txtAgentTeam.Text
                        QAlvl2EmailScorecard.lblContactType.Text = cboContactType.Text
                        QAlvl2EmailScorecard.lblSRNumber.Text = txtSRNumber.Text
                        QAlvl2EmailScorecard.lblUserID.Text = txtUserID.Text
                        QAlvl2EmailScorecard.lblJIRA.Text = txtJIRAbox.Text




                        QAResCallScorecard.lblAgentName1.Text = cboAgentName.Text
                        QAResCallScorecard.lblAgentTeam1.Text = txtAgentTeam.Text
                        QAResCallScorecard.lblContactType1.Text = cboContactType.Text
                        QAResCallScorecard.lblSRNumber1.Text = txtSRNumber.Text


                        QAResidentEmailScorecard.lblAgentName.Text = cboAgentName.Text
                        QAResidentEmailScorecard.lblAgentTeam.Text = txtAgentTeam.Text
                        QAResidentEmailScorecard.lblContactID1.Text = cboContactType.Text
                        QAResidentEmailScorecard.lblSRNumber.Text = txtSRNumber.Text


                        QAConsuACallScorcard.lblAgentName1.Text = cboAgentName.Text
                        QAConsuACallScorcard.lblAgentTeam1.Text = txtAgentTeam.Text
                        QAConsuACallScorcard.lblContactID1.Text = cboContactType.Text
                        QAConsuACallScorcard.lblSRNumber1.Text = txtSRNumber.Text





                    End If

                End If

            End If





        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub


    Private Sub cboAgentName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAgentName.SelectedIndexChanged



        Try





            sqltemp2 = "SELECT * FROM [Agents] WHERE AgentName='" & cboAgentName.Text & " ' "



            Dim cmdtemp As New SqlCommand





            cmdtemp.CommandText = sqltemp2

            cmdtemp.Connection = contemp2



            readertemp2 = cmdtemp.ExecuteReader



            If (readertemp2.Read() = True) Then



                txtAgentTeam.Text = (readertemp2("Platform"))



            End If



            cmdtemp.Dispose()

            readertemp2.Close()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try


    End Sub

    Private Sub QaSetup_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If MessageBox.Show("Are you sure to close this application?", "FADV Quality Assurance Application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            End

        Else
            e.Cancel = True


        End If




    End Sub


End Class