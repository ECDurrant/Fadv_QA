
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Threading


Public Class QASignUp

    Dim SQL As String
    Dim con As New OleDbConnection



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



        Me.Cursor = Cursors.WaitCursor


        If txtPassword.Text = "" Then




            MessageBox.Show("Please enter a password", "Warning", MessageBoxButtons.RetryCancel)

            Me.ActiveControl = txtPassword


            Me.Cursor = Cursors.Hand

        Else

            If txtSCRN.Text = "" Then




                MessageBox.Show("Please enter your SCRN ID", "Warning", MessageBoxButtons.RetryCancel)

                Me.ActiveControl = txtSCRN

                Me.Cursor = Cursors.Hand

            Else

                If cboDrive.Text = "Select" Then




                    MessageBox.Show("You must select your Main Drive from the list", "Warning", MessageBoxButtons.RetryCancel)

                    Me.ActiveControl = cboDrive


                    Me.Cursor = Cursors.Hand

                Else


                    If ComboBox1.Text = "Select your name from list" Then




                        MessageBox.Show("Please select your name from the list", "Warning", MessageBoxButtons.RetryCancel)

                        Me.ActiveControl = ComboBox1


                        Me.Cursor = Cursors.Hand
                    Else


                        If BackgroundWorker1.IsBusy = False Then

                            BackgroundWorker1.RunWorkerAsync()


                        End If


                    End If

                End If

            End If

        End If




    End Sub


    Public Sub store()

        Try


            '   con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")


            con = New System.Data.OleDb.OleDbConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")


            con.Open()



            Dim SQL As String = "INSERT INTO [Login] ([UserName], [Password], [Drive],[SCRN]) Values (?,?,?,?)"

            Using cmd As New OleDbCommand(SQL, con)



                cmd.Parameters.AddWithValue("@p1", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@p2", txtPassword.Text)
                cmd.Parameters.AddWithValue("@p3", cboDrive.Text)
                cmd.Parameters.AddWithValue("@p4", txtSCRN.Text)



                cmd.ExecuteNonQuery()

                con.Close()



            End Using


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Public Sub loginload()

        Try



            ' sqllogin = "SELECT * FROM [Login] WHERE UserName='" & cbologinnames.Text & "' "

            sqllogin = "SELECT * FROM [Login]"

            Dim cmdLogin As New SqlCommand

            cmdLogin.CommandText = sqllogin
            cmdLogin.Connection = conlogin


            readerlogin = cmdLogin.ExecuteReader

            While (readerlogin.Read())

                LogIn.cboLogInNames.Items.Add(readerlogin("UserName"))



            End While




            cmdLogin.Dispose()
            readerlogin.Close()


            ''Error Checking
        Catch ex As OleDbException

            If ConnectionState.Broken = True Then

                con.Close()
                con.Close()

                con.Open()
                con.Open()


                MsgBox("The connection to the P drive was interupted..@ fill combo procedure in loginload in QAsignup")


            End If

        Catch ex As SyntaxErrorException





        Catch ex As SystemException

            '  MsgBox("system error at fill combo procedure")


        Catch ex As Exception

            MsgBox(ex.Message, 0 Or 48, "Alert")

        End Try



    End Sub

    Public Sub DeleteN()





        Try




            Dim con As OleDbConnection

            Dim com As OleDbCommand

            '  con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")

            con = New System.Data.OleDb.OleDbConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")



            com = New OleDbCommand("delete from [Supervisor] where [FullName] =@ID", con)



            con.Open()



            com.Parameters.AddWithValue("@ID", ComboBox1.Text)





            com.ExecuteNonQuery()






            Me.Cursor = Cursors.Hand



            con.Close()



        Catch ex As Exception



            MsgBox(ex.Message, 0 Or 48, "Alert")



        End Try









    End Sub




    Private Sub QASignUp_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Me.CenterToScreen()

        Me.ActiveControl = ComboBox1


        Control.CheckForIllegalCrossThreadCalls = False


        QaSignUpmod.connecttemp()


        Fillcombo()





    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked





        '   Timer1.Enabled = True

        LogIn.Show()

        Me.Hide()



    End Sub





    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        LogIn.cboLogInNames.Items.Clear()

        Log_In.connectlogin()

        loginload()


        LogIn.Show()
        Me.Hide()


        Timer1.Enabled = False



    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork



        For i = 0 To 100

            System.Threading.Thread.Sleep(50)
            Me.BackgroundWorker1.ReportProgress(i)

            '  Label1.Text = i.ToString

            i = i
        Next

        store()

        DeleteN()




    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

        ProgressBar1.Value = e.ProgressPercentage



    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted



        '  MsgBox("Thank You for registering please look to your bottom right and select ‘login’ if your dont s")

        MsgBox("Thank You for registering, you will be taken back to the main login screen. if you dont see your name in dropdown just type name in")



        Me.Cursor = Cursors.Hand




        Me.Hide()

        LogIn.Show()




    End Sub


    Public Sub Fillcombo()



        Try





            sqltemp = "SELECT * FROM [Supervisor]"



            Dim cmdtemp As New OleDb.OleDbCommand





            cmdtemp.CommandText = sqltemp

            cmdtemp.Connection = contemp



            readertemp = cmdtemp.ExecuteReader

            While (readertemp.Read())


                ComboBox1.Items.Add(readertemp("FullName"))


            End While



            cmdtemp.Dispose()

            readertemp.Close()







        Catch ex As Exception



            MsgBox(ex.Message, 0 Or 48, "Alert")



        End Try





    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
End Class