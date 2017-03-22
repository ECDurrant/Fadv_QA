
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports DevExpress.LookAndFeel

Public Class LogIn

    Dim con As New SqlConnection
    Dim con1 As New SqlConnection
    Dim con00 As New SqlConnection

    Dim SQL As String
    Dim SQL1 As String
    Dim SQL333 As String
    Dim goodiDA As New SqlClient.SqlDataAdapter
    Dim goodiDS As New DataSet
    Dim strCon As String = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
    Dim Ann As String
    Dim Ann2 As String
    Dim UpdateD As String
    Dim Changer As String
    Dim ConnectedSupervisor As String
    Dim UserRegion As String


    Dim Desk = My.Computer.FileSystem.SpecialDirectories.Desktop



    Dim Drive As String
    Dim SCRN As String
    Dim Decider As String


    ' Dim strCon As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

    '  Dim Desk As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)




    '  Dim strCon As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\QA.accdb"

    Public Sub LoginLoad()




        Try


            sqllogin = "SELECT * FROM [Login]"

            Dim cmdLogin As New SqlCommand

            cmdLogin.CommandText = sqllogin
            cmdLogin.Connection = conlogin


            readerlogin = cmdLogin.ExecuteReader

            While (readerlogin.Read())

                cboLogInNames.Items.Add(readerlogin("UserName"))


            End While


            cmdLogin.Dispose()
            readerlogin.Close()


            ''Error Checking
        Catch ex As SqlException

            If ConnectionState.Broken = True Then

                con.Close()
                con.Close()

                con.Open()
                con.Open()


                MsgBox("The connection to the P drive was interupted..@ login load procedure")


            End If

        Catch ex As SyntaxErrorException





        Catch ex As SystemException

            '  MsgBox("system error at fill combo procedure")


        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try



    End Sub






    Private Sub btnLogIn_Click(sender As Object, e As EventArgs) Handles btnLogIn.Click

        Try

            If cboLogInNames.Text = "Type Here..." And txtPassword.Text = "" Then




                MessageBox.Show("Please be advised that you must enter in a valid username and password in order to proceed", "Warning", MessageBoxButtons.RetryCancel)

                Me.ActiveControl = cboLogInNames



            Else

                If cboLogInNames.Text = "Type Here..." Then






                    MessageBox.Show("Please be advised that a valid username must be entered in order to proceed", "Warning", MessageBoxButtons.RetryCancel)


                    Me.ActiveControl = cboLogInNames


                Else

                    If cboLogInNames.Text = "" Then





                        MessageBox.Show("Please be advised that a valid username must be entered in order to proceed", "Warning", MessageBoxButtons.RetryCancel)


                        Me.ActiveControl = cboLogInNames




                    Else

                        If txtPassword.Text = "" Then






                            MessageBox.Show("Please be advised that a valid username must be entered in order to proceed", "Warning", MessageBoxButtons.RetryCancel)

                            Me.ActiveControl = txtPassword




                        Else




                            SplashScreenManager1.ShowWaitForm()

                            '  Me.Cursor = Cursors.WaitCursor

                            'If BackgroundWorker1.IsBusy = False Then

                            ' BackgroundWorker1.RunWorkerAsync()



                            'End If
                            '  PleaseWait.Show()


                            '  PW.Show()

                            '     Timer1.Enabled = True

                            QALogin()



                            '  QAlogin2()

                            '  QALogin33()







                        End If

                    End If

                End If

            End If




        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try




    End Sub

    Public Sub LogmeinQuick()





        cboLogInNames.Text = Form2.lblQAauditor.Text


        Form2.Show()

        Me.Hide()












    End Sub

    Public Sub QALogin()


        Dim con3 As New SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30")

        con3.Open()

        Dim cmd3 As New SqlClient.SqlCommand("Select * FROM [LogIn] WHERE UserName = '" & cboLogInNames.Text & " ' And Password = '" & txtPassword.Text & "'", con3)
        Dim dr3 As SqlDataReader = cmd3.ExecuteReader



        While dr3.Read


            lblDrive.Text = dr3(3).ToString

            lblSCRN.Text = dr3(4).ToString

            lblDeciderLogin.Text = dr3(5).ToString

            lblESDecider.Text = dr3(6).ToString

            lblUserEmail.Text = dr3(8).ToString

            lblEmailPassword.Text = dr3(9).ToString

            Ann = dr3(10).ToString

            UpdateD = dr3(11).ToString

            Ann2 = dr3(12).ToString

            Changer = dr3(13).ToString

            ConnectedSupervisor = dr3(14).ToString

            UserRegion = dr3(17).ToString



        End While


        If dr3.HasRows = True Then



            Form2.lblconnectedsupervisor1.Text = ConnectedSupervisor

            Form2.lblQAauditor.Text = cboLogInNames.Text
            Form2.lblQAauditor2.Text = cboLogInNames.Text
            Form2.lblQAAuditor3.Text = cboLogInNames.Text

            Form2.lblMDrive.Text = lblDrive.Text

            Form2.lblSCRN.Text = lblSCRN.Text

            Form2.lblDeciderDash.Text = lblDeciderLogin.Text

            Form2.lblESDecider.Text = lblESDecider.Text

            Form2.lblUserEmail.Text = lblUserEmail.Text

            Form2.lblEmailPassword.Text = lblEmailPassword.Text

            Form2.lblAppVersion.Text = UpdateD

            Form2.lblAnn.Text = Ann

            Form2.lblPleaseUpdateApp.Text = Ann2

            Form2.lblRegion.Text = UserRegion

            Form2.lblChanger.Text = Changer

            lblDrive.Text = Drive

            lblDeciderLogin.Text = Decider

            lblSCRN.Text = SCRN



            Form2.Show()

            Me.Hide()
            PW.Hide()
            SplashScreenManager1.CloseWaitForm()

        Else


            Me.Cursor = Cursors.Hand

            PW.Hide()

            SplashScreenManager1.CloseWaitForm()


            Dim QAIncorrectPass As Integer = MessageBox.Show("Username or Password Incorrect please try again", "FADV Quality Assurance Application", MessageBoxButtons.OK, MessageBoxIcon.Question)


            If QAIncorrectPass = DialogResult.OK Then




            End If



        End If


        con3.Close()




    End Sub




    Public Sub QAlogin2()


        Try
            '  con00 = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb")

            ' con00 = New SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")

            con00 = New SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30")

            con00.Open()


            'the query:
            Dim cmd As SqlCommand = New SqlCommand("Select * FROM [LogIn] WHERE UserName = '" & cboLogInNames.Text & " ' And Password = '" & txtPassword.Text & "'", con00)
            Dim dr As SqlDataReader = cmd.ExecuteReader


            ' the following variable is hold true if user is found, and false if user is not found
            Dim userFound As Boolean = False



            ' This is will hold drive information.



            'if found:

            While dr.Read
                userFound = True

                'Drive = dr("Drive").ToString
                'SCRN = dr("SCRN").ToString


                lblDrive.Text = dr(3).ToString

                lblSCRN.Text = dr(4).ToString

                lblDeciderLogin.Text = dr(5).ToString

                lblESDecider.Text = dr(6).ToString

                lblUserEmail.Text = dr(8).ToString

                lblEmailPassword.Text = dr(9).ToString

                Ann = dr(10).ToString

                UpdateD = dr(11).ToString

                Ann2 = dr(12).ToString

                Changer = dr(13).ToString

                ConnectedSupervisor = dr(14).ToString

                UserRegion = dr(17).ToString

            End While



            'checking the result
            If userFound = True Then


                Form2.lblconnectedsupervisor1.Text = ConnectedSupervisor

                Form2.lblQAauditor.Text = cboLogInNames.Text
                Form2.lblQAauditor2.Text = cboLogInNames.Text
                Form2.lblQAAuditor3.Text = cboLogInNames.Text

                Form2.lblMDrive.Text = lblDrive.Text

                Form2.lblSCRN.Text = lblSCRN.Text

                Form2.lblDeciderDash.Text = lblDeciderLogin.Text

                Form2.lblESDecider.Text = lblESDecider.Text

                Form2.lblUserEmail.Text = lblUserEmail.Text

                Form2.lblEmailPassword.Text = lblEmailPassword.Text

                Form2.lblAppVersion.Text = UpdateD

                Form2.lblAnn.Text = Ann

                Form2.lblPleaseUpdateApp.Text = Ann2

                Form2.lblRegion.Text = UserRegion

                Form2.lblChanger.Text = Changer

                lblDrive.Text = Drive

                lblDeciderLogin.Text = Decider

                lblSCRN.Text = SCRN



                Form2.Show()

                Me.Hide()
                PW.Hide()
                SplashScreenManager1.CloseWaitForm()

            Else


                '' If username and password does not match
                Timer1.Enabled = False

                Me.Cursor = Cursors.Hand

                PW.Hide()

                SplashScreenManager1.CloseWaitForm()


                Dim QAIncorrectPass As Integer = MessageBox.Show("Username or Password Incorrect please try again", "FADV Quality Assurance Application", MessageBoxButtons.OK, MessageBoxIcon.Question)


                If QAIncorrectPass = DialogResult.OK Then




                End If




            End If



            con00.Close()

            'dr.Close()


        Catch ex As Exception

            PW.Hide()

            MsgBox(ex.Message, 0 Or 48, "Alert")


            Timer1.Enabled = False

            PW.Hide()

        End Try



    End Sub




    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked


        QASignUp.Show()

        Me.Hide()




    End Sub




    Private Sub LogIn_Load(sender As Object, e As EventArgs) Handles MyBase.Load





        Me.CenterToScreen()


        Me.ActiveControl = cboLogInNames

        Log_In.connectlogin()
        LoginLoad()


        Control.CheckForIllegalCrossThreadCalls = False




    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click


        Me.ActiveControl = txtUserName

        txtPassword.Clear()
        ' txtUserName.Clear()

        cboLogInNames.Text = "Type Here..."




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnClose.Click



        Close()


    End Sub

    Private Sub lblForgotpass_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblForgotpass.LinkClicked



        'ForgotPass.Show()
        Me.Hide()






    End Sub

    Private Sub LogIn_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing


        If MessageBox.Show("Are you sure to close this application?", "FADV Quality Assurance Application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            'End

        Else
            e.Cancel = True


        End If









    End Sub



    Public Sub SetMainDrive()

















    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Try

            For i = 0 To 100

                System.Threading.Thread.Sleep(50)
                Me.BackgroundWorker1.ReportProgress(i)

                '   Label3.Text = i.ToString
                i = i
            Next

            QAlogin2()





        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try





    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted



    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged



        ProgressBar1.Value = e.ProgressPercentage


    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try
            QAlogin2()

            Timer1.Enabled = False


        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")


            Timer1.Enabled = False


            PW.Hide()

        End Try

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)




    End Sub
End Class