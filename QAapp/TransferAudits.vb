

Imports System.Data.SqlClient

Public Class TransferAudits



    Private Sub TransferAudits_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Me.CenterToScreen()

        Loadcombo()



    End Sub


    Public Sub Loadcombo()

        Try

            QaSetupMod.connecttemp20()

            sqltemp20 = "SELECT * FROM [Supervisor]"


            Dim cmdtemp As New SqlClient.SqlCommand


            cmdtemp.CommandText = sqltemp20

            cmdtemp.Connection = contemp20



            readertemp20 = cmdtemp.ExecuteReader


            While (readertemp20.Read())


                cboFrom.Items.Add(readertemp20("FullName"))
                cboTo.Items.Add(readertemp20("FullName"))


            End While



            cmdtemp.Dispose()

            contemp20.Close()

            Me.Cursor = Cursors.Hand




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try





    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If cboFrom.Text = "" Then

            MsgBox("Please fill out the required fields", MessageBoxButtons.RetryCancel)


        Else


            If cboTo.Text = "" Then

                MsgBox("Please fill out the required fields", MessageBoxButtons.RetryCancel)


            Else




                Dim msg = "Are you sure you want to transfer the pending audits?"

                Dim title = "FADV QA Application"

                Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

                Dim responce = MsgBox(msg, style, title)

                If responce = MsgBoxResult.Yes Then

                    SplashScreenManager1.ShowWaitForm()


                    Dim con As SqlConnection = New SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")


                    Dim cmd As SqlCommand = New SqlCommand("Update QAMainDB Set Supervisor='" & cboTo.Text & "', AuditTransferManager='" & Form2.lblQAAuditor3.Text & "' WHERE Supervisor='" & cboFrom.Text & "' AND PendingDisputeID= 'Pending Review'", con)

                    '      Dim cmd As SqlCommand = New SqlCommand("Update QAMainDB Set Supervisor='" & cboTo.Text & "' WHERE Supervisor='" & cboFrom.Text & "' AND PendingDisputeID= 'Review Pending'", con)



                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()


                    Timer1.Enabled = True



                Else





                End If


            End If

        End If




    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick



        Timer1.Enabled = False


        SplashScreenManager1.CloseWaitForm()


        MsgBox("The Pending Audits from " & cboFrom.Text & " were successfully transferred to " & cboTo.Text & " please close this form and refresh your audit list")


        Me.Hide()






    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim msg = "Are you sure you want to close this form"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)

        If responce = MsgBoxResult.Yes Then

            Me.Hide()


        Else



        End If





    End Sub
End Class