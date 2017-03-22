


Imports System.Threading

Imports Microsoft.Office.Interop


Imports System.Data.OleDb


'Imports i00SpellCheck


Public Class QAConsuACallScorcard

    Dim SQL As String
    Dim con As New OleDbConnection


    Dim One As Integer
    Dim two As Integer
    Dim three As Integer
    Dim Four As Integer
    Dim Five As Integer
    Dim Six As Integer
    Dim seven As Integer


    ''Store Call Thread
    Dim StoreCallThread As System.Threading.Thread




    Private Sub QAConsuACallScorcard_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Try

            Me.WindowState = FormWindowState.Maximized

            Me.ActiveControl = cbo1_1


            ''Date
            Time.Enabled = True


            Control.CheckForIllegalCrossThreadCalls = False

            '  Me.EnableControlExtensions()





        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub




    Public Sub Store()




        Try



            ' Test

            con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")



            'P Drive 

            '   con = New System.Data.OleDb.OleDbConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")





            '' Dyanic


            ' con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")



            con.Open()



            Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA-Agent],[QA-Team],[QA-ContactDate],[QA-OrderID],[QA-Date],[QA-Comments],[QA-Opp],[CI-Name],[CI-Account],[CI-Company],[CI-Phone],[CI-Email],[Rev-Date],[Rev-Manager],[Rev-Comments],[Dis-Score],[Dis-Name],[Dis-Notes],[Dis-AppComments],[One-1],[One-2],[One-3],[One-1Note],[One-2Note],[One-3Note],[Two-1],[Two-1Note],[Three-1],[Three-2],[Three-3],[Three-4],[Three-5],[Three-6],[Three-7],[Three-8],[Three-9],[Three-1Note],[Three-2Note],[Three-3Note],[Three-4Note],[Three-5Note],[Three-6Note],[Three-7Note],[Three-8Note],[Three-9Note],[Four-1],[Four-2],[Four-3],[Four-1Note],[Four-2Note],[Four-3Note],[Five-1],[Five-2],[Five-1Note],[Five-2Note],[Six-1],[Six-2],[Six-3],[Six-4],[Six-1Note],[Six-2Note],[Six-3Note],[Six-4Note],[QAScore],[AutoFail],[Auditor]) Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

            Using cmd As New OleDbCommand(SQL, con)



                cmd.Parameters.AddWithValue("@p1", lblSRNumber1.Text)
                cmd.Parameters.AddWithValue("@p2", lblContactID1.Text)
                cmd.Parameters.AddWithValue("@p3", lblContactType1.Text)
                cmd.Parameters.AddWithValue("@p4", lblAgentName1.Text)
                cmd.Parameters.AddWithValue("@p5", lblAgentTeam1.Text)
                cmd.Parameters.AddWithValue("@p6", txtgDatebox.Text)
                cmd.Parameters.AddWithValue("@p7", txtgorderid.Text)
                cmd.Parameters.AddWithValue("@p8", Date.Now.ToString("MM/dd/yyyy"))
                cmd.Parameters.AddWithValue("@p9", txtQACom.Text)
                cmd.Parameters.AddWithValue("@p10", txtQAAOO.Text)
                cmd.Parameters.AddWithValue("@p11", txtgnamebox.Text)
                cmd.Parameters.AddWithValue("@p12", txtgacc.Text)
                cmd.Parameters.AddWithValue("@p13", txtgcompany.Text)
                cmd.Parameters.AddWithValue("@p14", txtgphone.Text)
                cmd.Parameters.AddWithValue("@p15", txtgemail.Text)
                cmd.Parameters.AddWithValue("@p16", "9/9/1988")
                cmd.Parameters.AddWithValue("@p17", "")
                cmd.Parameters.AddWithValue("@p18", "")
                cmd.Parameters.AddWithValue("@p19", "")
                cmd.Parameters.AddWithValue("@p20", "")
                cmd.Parameters.AddWithValue("@p21", "")
                cmd.Parameters.AddWithValue("@p22", "")


                cmd.Parameters.AddWithValue("@p23", cbo1_1.Text)
                cmd.Parameters.AddWithValue("@p24", cbo1_2.Text)
                cmd.Parameters.AddWithValue("@p25", cbo1_3.Text)

                cmd.Parameters.AddWithValue("@p26", txt1_1.Text)
                cmd.Parameters.AddWithValue("@p27", txt1_2.Text)
                cmd.Parameters.AddWithValue("@p28", txt1_3.Text)

                cmd.Parameters.AddWithValue("@p29", cbo2_1.Text)

                cmd.Parameters.AddWithValue("@p30", txt2_1.Text)

                cmd.Parameters.AddWithValue("@p31", cbo3_1.Text)
                cmd.Parameters.AddWithValue("@p32", cbo3_2.Text)
                cmd.Parameters.AddWithValue("@p33", cbo3_3.Text)
                cmd.Parameters.AddWithValue("@p34", cbo3_4.Text)
                cmd.Parameters.AddWithValue("@p35", cbo3_5.Text)

                cmd.Parameters.AddWithValue("@p36", cbo3_6.Text)

                cmd.Parameters.AddWithValue("@p37", cbo3_7.Text)
                cmd.Parameters.AddWithValue("@p38", cbo3_8.Text)
                cmd.Parameters.AddWithValue("@p39", cbo3_9.Text)


                cmd.Parameters.AddWithValue("@p40", txt3_1.Text)


                cmd.Parameters.AddWithValue("@p41", txt3_2.Text)
                cmd.Parameters.AddWithValue("@p42", txt3_3.Text)
                cmd.Parameters.AddWithValue("@p43", txt3_4.Text)
                cmd.Parameters.AddWithValue("@p44", txt3_5.Text)
                cmd.Parameters.AddWithValue("@p45", txt3_6.Text)
                cmd.Parameters.AddWithValue("@p46", txt3_7.Text)
                cmd.Parameters.AddWithValue("@p47", txt3_8.Text)
                cmd.Parameters.AddWithValue("@p48", txt3_9.Text)

                cmd.Parameters.AddWithValue("@p49", Cbo4_1.Text)
                cmd.Parameters.AddWithValue("@p50", cbo4_2.Text)
                cmd.Parameters.AddWithValue("@p51", cbo4_3.Text)

                cmd.Parameters.AddWithValue("@p52", txt4_1.Text)
                cmd.Parameters.AddWithValue("@p53", txt4_2.Text)
                cmd.Parameters.AddWithValue("@p54", txt4_3.Text)



                cmd.Parameters.AddWithValue("@p55", cbo5_1.Text)
                cmd.Parameters.AddWithValue("@p56", cbo5_2.Text)


                cmd.Parameters.AddWithValue("@p57", txt5_1.Text)
                cmd.Parameters.AddWithValue("@p58", txt5_2.Text)



                cmd.Parameters.AddWithValue("@p59", cbo6_1.Text)
                cmd.Parameters.AddWithValue("@p60", cbo6_2.Text)
                cmd.Parameters.AddWithValue("@p61", cbo6_3.Text)
                cmd.Parameters.AddWithValue("@p62", cbo6_4.Text)


                cmd.Parameters.AddWithValue("@p62", txt6_1.Text)
                cmd.Parameters.AddWithValue("@p63", txt6_2.Text)
                cmd.Parameters.AddWithValue("@p64", txt6_3.Text)
                cmd.Parameters.AddWithValue("@p65", txt6_4.Text)

                cmd.Parameters.AddWithValue("@p66", lblQAScore1.Text)

                If cboAutoFail.Checked Then

                    cmd.Parameters.AddWithValue("@p67", cboAF.Text)

                Else

                    cmd.Parameters.AddWithValue("@p67", "N/a")

                End If

                cmd.Parameters.AddWithValue("@p68", lblQAauditor1.Text)


                cmd.ExecuteNonQuery()

                con.Close()



            End Using


            ' MsgBox("Info saved")


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Public Sub QAExcell()




        Try



            Dim oExcel As Object = CreateObject("Excel.Application")




            ''Test

            '  Dim oBook As Object = oExcel.Workbooks.Open("C:\Users\playe\Desktop\QA\ScoreCard Excell\CaCallSc.xlsx")

            '' P Drive

            '   Dim oBook As Object = oExcel.Workbooks.Open("P:\SPC\QA\CaCallSc.xlsx")


            '' Dynamic

            Dim oBook As Object = oExcel.Workbooks.Open(lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\CaCallSc.xlsx")





            Dim oSheet As Object = oBook.Worksheets("CACallSc")  'or oBook.Worksheets("SheetName")

            'e.g. Read value from A2 cell to TextBox1 

            '   txt1_1.Text = oSheet.Range("F7").Value

            'e.g. Write value from TextBox1 to A2 cell


            oSheet.Range("C3").Value = "" & One


            oSheet.Range("C4").Value = "" & cbo1_1.Text
            oSheet.Range("C5").Value = "" & cbo1_2.Text
            oSheet.Range("C6").Value = "" & cbo1_3.Text



            oSheet.Range("D4").Value = "" & txt1_1.Text
            oSheet.Range("D5").Value = "" & txt1_2.Text
            oSheet.Range("D6").Value = "" & txt1_3.Text

            oSheet.Range("C7").Value = "" & two

            oSheet.Range("C8").Value = "" & cbo2_1.Text
            oSheet.Range("D8").Value = "" & txt2_1.Text

            oSheet.Range("C9").Value = "" & three

            oSheet.Range("C10").Value = "" & cbo3_1.Text
            oSheet.Range("C11").Value = "" & cbo3_2.Text
            oSheet.Range("C12").Value = "" & cbo3_3.Text
            oSheet.Range("C13").Value = "" & cbo3_4.Text
            oSheet.Range("C14").Value = "" & cbo3_5.Text
            oSheet.Range("C15").Value = "" & cbo3_6.Text
            oSheet.Range("C16").Value = "" & cbo3_7.Text
            oSheet.Range("C17").Value = "" & cbo3_8.Text
            oSheet.Range("C18").Value = "" & cbo3_9.Text


            oSheet.Range("D10").Value = "" & txt3_1.Text
            oSheet.Range("D11").Value = "" & txt3_2.Text
            oSheet.Range("D12").Value = "" & txt3_3.Text
            oSheet.Range("D13").Value = "" & txt3_4.Text
            oSheet.Range("D14").Value = "" & txt3_5.Text
            oSheet.Range("D15").Value = "" & txt3_6.Text
            oSheet.Range("D16").Value = "" & txt3_7.Text
            oSheet.Range("D17").Value = "" & txt3_8.Text
            oSheet.Range("D18").Value = "" & txt3_9.Text



            oSheet.Range("C19").Value = "" & Four

            oSheet.Range("C20").Value = "" & Cbo4_1.Text
            oSheet.Range("C21").Value = "" & cbo4_2.Text
            oSheet.Range("C22").Value = "" & cbo4_3.Text

            oSheet.Range("D20").Value = "" & txt4_1.Text
            oSheet.Range("D21").Value = "" & txt4_2.Text
            oSheet.Range("D22").Value = "" & txt4_3.Text


            oSheet.Range("C23").Value = "" & Five


            oSheet.Range("C23").Value = "" & cbo5_1.Text
            oSheet.Range("C24").Value = "" & cbo5_2.Text

            oSheet.Range("D24").Value = "" & txt5_1.Text
            oSheet.Range("D25").Value = "" & txt5_2.Text

            oSheet.Range("C26").Value = "" & Six


            oSheet.Range("C27").Value = "" & cbo6_1.Text
            oSheet.Range("C28").Value = "" & cbo6_2.Text
            oSheet.Range("C29").Value = "" & cbo6_3.Text
            oSheet.Range("C30").Value = "" & cbo6_4.Text

            oSheet.Range("D27").Value = "" & txt6_1.Text
            oSheet.Range("D28").Value = "" & txt6_2.Text
            oSheet.Range("D29").Value = "" & txt6_3.Text
            oSheet.Range("D30").Value = "" & txt6_4.Text



            oSheet.Range("A48").Value = "" & txtQACom.Text
            oSheet.Range("A62").Value = "" & txtQAAOO.Text




            oSheet.Range("C31").Value = lblQAScore1.Text




            oSheet.Range("B33").Value = lblSRNumber1.Text
            oSheet.Range("B34").Value = lblContactID1.Text
            oSheet.Range("B35").Value = lblContactType1.Text
            oSheet.Range("B36").Value = "" & lblAgentName1.Text
            oSheet.Range("B37").Value = "" & lblAgentTeam1.Text
            oSheet.Range("B38").Value = Form2.DateTimePicker1.Text
            oSheet.Range("B39").Value = Form2.txtOrderID.Text
            oSheet.Range("B40").Value = "" & Form2.txtContactName.Text
            oSheet.Range("B41").Value = "" & Form2.txtContactEmail.Text
            oSheet.Range("B42").Value = "" & Form2.txtContactPhone.Text
            oSheet.Range("B43").Value = "" & Form2.txtCompany.Text
            oSheet.Range("B44").Value = "" & Form2.txtAccountNum.Text
            oSheet.Range("B45").Value = "" & lblQAauditor1.Text
            oSheet.Range("B46").Value = "" & lblDate1.Text








            '' Test

            '  oBook.SaveAs("C:\Users\playe\Desktop\QA\" & "SR#" & lblSRNumber1.Text & "_" & lblAgentName1.Text & "_" & lblDate1.Text & " QA Scorecard.xlsx")



            '' P drive

            ' oBook.SaveAs("P:\SPC\QA\" & lblSRNumber1.Text & "_" & lblAgentName1.Text & "_" & lblDate1.Text & " QA Scorecard.xlsx")


            '' Dynamic

            oBook.SaveAs(lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA2\" & "SR#" & lblSRNumber1.Text & "_" & lblAgentName1.Text & "_" & lblDate1.Text & " QA Scorecard.xlsx")




            oExcel.Quit()


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try

    End Sub


    Private Sub btnSaveScoreCard_Click(sender As Object, e As EventArgs) Handles btnSaveScoreCard.Click


        Try




            '    If Form2.cboAgentName.Text = "Agent Name" Or Form2.cboContactType.Text = "Contact Type" Then

            If Form2.cboAgentName.Text = "Agent Name" Or Form2.cboContactType.Text = "Contact Type" Then


                MsgBox("Please be advised you must fill out all 'Agent Information' before proceeding", MessageBoxButtons.RetryCancel)


            Else




                If cboAutoFail.Checked = True And cboAF.Text = "Auto Fail Reason" Then


                    MsgBox("Since this Audit was marked as 'Auto Fail', a reason must be selected before saving.", MessageBoxButtons.RetryCancel)



                    Me.ActiveControl = cboAF

                Else




                    If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then



                    Else


                        MsgBox("Please wait while your audit is being saved")

                        If BackgroundWorker1.IsBusy = False Then

                            BackgroundWorker1.RunWorkerAsync()




                            If cboAutoFail.Checked = True Then


                                lblQAScore1.Text = "0"


                                lblQAScore1.Visible = True


                            Else



                                'Tally Qa Score

                                '
                                QaTotalScore()

                                ''Show Scorecard



                            End If

                        End If

                    End If

                End If

            End If





        Catch ex As Exception



            MsgBox(ex.Message)

        End Try






    End Sub

    Public Sub QaTotalScore()


        '  Dim strQaScoreTotal As String
        Dim intQascoreTotal As Integer


        Dim int1_1 As Integer = cbo1_1.Text
        Dim int1_2 As Integer = cbo1_2.Text
        Dim int1_3 As Integer = cbo1_3.Text

        Dim int2_1 As Integer = cbo2_1.Text

        Dim int3_1 As Integer = cbo3_1.Text
        Dim int3_2 As Integer = cbo3_2.Text
        Dim int3_3 As Integer = cbo3_3.Text
        Dim int3_4 As Integer = cbo3_4.Text
        Dim int3_5 As Integer = cbo3_5.Text
        Dim int3_6 As Integer = cbo3_6.Text
        Dim int3_7 As Integer = cbo3_7.Text
        Dim int3_8 As Integer = cbo3_8.Text
        Dim int3_9 As Integer = cbo3_9.Text

        Dim int4_1 As Integer = Cbo4_1.Text
        Dim int4_2 As Integer = cbo4_2.Text
        Dim int4_3 As Integer = cbo4_3.Text


        Dim int5_1 As Integer = cbo5_1.Text
        Dim int5_2 As Integer = cbo5_2.Text

        Dim int6_1 As Integer = cbo6_1.Text
        Dim int6_2 As Integer = cbo6_2.Text
        Dim int6_3 As Integer = cbo6_3.Text
        Dim int6_4 As Integer = cbo6_4.Text



        One = int1_1 + int1_2 + int1_3

        two = int2_1

        three = int3_1 + int3_2 + int3_3 + int3_4 + int3_5 + int3_6 + int3_7 + int3_8 + int3_9

        Four = int4_1 + int4_2 + int4_3

        Five = int5_1 + int5_2

        Six = int6_1 + int6_2 + int6_3 + int6_4





        intQascoreTotal = int1_1 + int1_2 + int1_3 + int2_1 + int3_1 + int3_2 + int3_3 + int3_4 + int3_5 + int3_6 + int3_7 + int3_8 + int3_9 + int4_1 + int4_2 + int4_3 + int5_1 + int5_2 + int6_1 + int6_2 + int6_3 + int6_4

        lblQAScore1.Text = intQascoreTotal

        lblQAScore1.Visible = True



    End Sub


    Public Sub QACalldisableControls()


        cbo1_1.Enabled = False
        cbo1_2.Enabled = False
        cbo1_3.Enabled = False

        cbo2_1.Enabled = False

        cbo3_1.Enabled = False
        cbo3_2.Enabled = False
        cbo3_3.Enabled = False
        cbo3_4.Enabled = False
        cbo3_5.Enabled = False
        cbo3_6.Enabled = False
        cbo3_7.Enabled = False
        cbo3_8.Enabled = False
        cbo3_9.Enabled = False

        Cbo4_1.Enabled = False
        cbo4_2.Enabled = False
        cbo4_3.Enabled = False


        cbo5_1.Enabled = False
        cbo5_2.Enabled = False

        cbo6_1.Enabled = False
        cbo6_2.Enabled = False
        cbo6_3.Enabled = False
        cbo6_4.Enabled = False


        ''reset Textboxes

        txt1_1.Enabled = False
        txt1_2.Enabled = False

        txt1_3.Enabled = False


        txt2_1.Enabled = False


        txt3_1.Enabled = False
        txt3_2.Enabled = False
        txt3_3.Enabled = False
        txt3_4.Enabled = False
        txt3_5.Enabled = False
        txt3_6.Enabled = False
        txt3_7.Enabled = False
        txt3_8.Enabled = False
        txt3_9.Enabled = False


        txt4_1.Enabled = False
        txt4_2.Enabled = False
        txt4_3.Enabled = False

        txt5_1.Enabled = False
        txt5_2.Enabled = False




        txt6_1.Enabled = False
        txt6_2.Enabled = False
        txt6_3.Enabled = False
        txt6_4.Enabled = False

    End Sub

    Public Sub resetatglance()

        ''Reset Scorecard at a glance info

        lblAgentName1.Text = "N/a"
        lblAgentTeam1.Text = "N/a"
        lblSRNumber1.Text = "N/a"
        lblContactType1.Text = "N/a"
        lblQAScore1.Text = "0"
    End Sub


    Public Sub QACallclear()


        ''Reset Comboboxes

        cbo1_1.Text = 2
        cbo1_2.Text = 1
        cbo1_3.Text = 2

        cbo2_1.Text = 25

        cbo3_1.Text = 5
        cbo3_2.Text = 3
        cbo3_3.Text = 5
        cbo3_4.Text = 4
        cbo3_5.Text = 5
        cbo3_6.Text = 3
        cbo3_7.Text = 5
        cbo3_8.Text = 1
        cbo3_9.Text = 4

        Cbo4_1.Text = 5
        cbo4_2.Text = 5
        cbo4_3.Text = 5


        cbo5_1.Text = 5
        cbo5_2.Text = 5

        cbo6_1.Text = 3
        cbo6_2.Text = 2
        cbo6_3.Text = 3
        cbo6_4.Text = 2



        ''reset Textboxes

        txt1_1.Clear()
        txt1_2.Clear()

        txt1_3.Clear()


        txt2_1.Clear()


        txt3_1.Clear()
        txt3_2.Clear()
        txt3_3.Clear()
        txt3_4.Clear()
        txt3_5.Clear()
        txt3_6.Clear()
        txt3_7.Clear()
        txt3_8.Clear()
        txt3_9.Clear()


        txt4_1.Clear()
        txt4_2.Clear()
        txt4_3.Clear()

        txt5_1.Clear()
        txt5_2.Clear()




        txt6_1.Clear()
        txt6_2.Clear()
        txt6_3.Clear()
        txt6_4.Clear()


        txtQAAOO.Clear()
        txtQACom.Clear()



        lblQAScore1.Visible = False


    End Sub

    Public Sub QACallEnable()




        ''Reset Comboboxes

        cbo1_1.Enabled = True
        cbo1_2.Enabled = True
        cbo1_3.Enabled = True

        cbo2_1.Enabled = True

        cbo3_1.Enabled = True
        cbo3_2.Enabled = True
        cbo3_3.Enabled = True
        cbo3_4.Enabled = True
        cbo3_5.Enabled = True
        cbo3_6.Enabled = True
        cbo3_7.Enabled = True
        cbo3_8.Enabled = True
        cbo3_9.Enabled = True



        Cbo4_1.Enabled = True
        cbo4_2.Enabled = True
        cbo4_3.Enabled = True


        cbo5_1.Enabled = True
        cbo5_2.Enabled = True

        cbo6_1.Enabled = True
        cbo6_2.Enabled = True
        cbo6_3.Enabled = True
        cbo6_4.Enabled = True



        ''reset Textboxes

        txt1_1.Enabled = True
        txt1_2.Enabled = True

        txt1_3.Enabled = True


        txt2_1.Enabled = True


        txt3_1.Enabled = True
        txt3_2.Enabled = True
        txt3_3.Enabled = True
        txt3_4.Enabled = True
        txt3_5.Enabled = True
        txt3_6.Enabled = True
        txt3_7.Enabled = True
        txt3_8.Enabled = True
        txt3_9.Enabled = True


        txt4_1.Enabled = True
        txt4_2.Enabled = True
        txt4_3.Enabled = True

        txt5_1.Enabled = True
        txt5_2.Enabled = True




        txt6_1.Enabled = True
        txt6_2.Enabled = True
        txt6_3.Enabled = True
        txt6_4.Enabled = True



    End Sub


    Private Sub btnQaSetup_Click(sender As Object, e As EventArgs) Handles btnQaSetup.Click


        Try


            If lblQAScore1.Visible = True Then

                reset()


            Else






                '  Form2.Show()
                Form2.Show()

                ''Transfer Qa auditor Name to form

                Form2.lblQAauditor.Text = lblQAauditor1.Text


                QAChatScorecard.lblQAauditor1.Text = lblQAauditor1.Text

                QAEmailScorecard.lblQAauditor1.Text = lblQAauditor1.Text

                QALvl2CallScorecard.lblQAauditor1.Text = lblQAauditor1.Text


            End If





        Catch ex As Exception

            MsgBox(ex.Message)


        End Try

    End Sub

    Private Sub Time_Tick(sender As Object, e As EventArgs) Handles Time.Tick


        lblDate1.Text = Date.Now.ToString("MMM dd yyyy")



    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork



        Try

            For i = 0 To 100

                System.Threading.Thread.Sleep(60)
                Me.BackgroundWorker1.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i
            Next


            ''


            '  Store()

            ' Send to Excell
            QAExcell()



            StoreCallThread = New System.Threading.Thread(AddressOf Store)
            '
            StoreCallThread.Start()



        Catch ex As Exception

            MsgBox(ex.Message)


        End Try

    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

        ProgressBar1.Value = e.ProgressPercentage



    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted



        lblQAScore1.Visible = True

        If MsgBox(lblAgentName1.Text & " " & "" & "scored a total of" & " " & lblQAScore1.Text & " " & "points on this QA audit,would you like to start a new one?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then


            MsgBox("You can now only review the saved scorecard, press 'QA Setup form' to clear and start a new audit")


        Else


            reset()




        End If


    End Sub



    Public Sub reset()

        ''Clear and reset the Form2 Tab

        Form2.Clear()


        Form2.Show()

        ''Reset Scorecard at a glance info

        resetatglance()

        ''Reset scorecard

        QACallclear()

        ''Transfer Qa Name to Wasetupform


        '  Form2.lblQAauditor.Text = lblQAauditor1.Text


        ''Reable buttons

        QACallEnable()


        Me.Hide()


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

    Private Sub Button3_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt1_3.Text = "" Then

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
            Clipboard.SetDataObject(txt1_3.Text)

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
                    txt1_3.Text = CType(iData.GetData(DataFormats.Text),
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

    Private Sub Button5_Click(sender As Object, e As EventArgs)
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

    Private Sub Button6_Click(sender As Object, e As EventArgs)
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
    Private Sub Button7_Click(sender As Object, e As EventArgs)

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

    Private Sub Button8_Click(sender As Object, e As EventArgs)

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

    Private Sub Button9_Click(sender As Object, e As EventArgs)

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

    Private Sub Button10_Click(sender As Object, e As EventArgs)

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

    Private Sub Button11_Click(sender As Object, e As EventArgs)

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

    Private Sub Button12_Click(sender As Object, e As EventArgs)

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

    Private Sub Button13_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt3_9.Text = "" Then

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
            Clipboard.SetDataObject(txt3_9.Text)

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
                    txt3_9.Text = CType(iData.GetData(DataFormats.Text),
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

    Private Sub Button21_Click(sender As Object, e As EventArgs)
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

    Private Sub Button22_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt6_4.Text = "" Then

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
            Clipboard.SetDataObject(txt6_4.Text)

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
                    txt6_4.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub


    Private Sub cbo1_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo1_1.SelectedIndexChanged

        Dim int1_1 As Integer = cbo1_1.Text

        If cbo1_1.Text = 0 Then


            txt1_1.BackColor = Color.Yellow

        ElseIf cbo1_1.Text > 0 Then


            txt1_1.BackColor = Color.White


        End If






    End Sub

    Private Sub cbo1_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo1_2.SelectedIndexChanged

        Dim int1_2 As Integer = cbo1_2.Text

        If cbo1_2.Text = 0 Then


            txt1_2.BackColor = Color.Yellow

        ElseIf cbo1_2.Text > 0 Then


            txt1_2.BackColor = Color.White


        End If





    End Sub

    Private Sub cbo1_3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo1_3.SelectedIndexChanged

        Dim int1_3 As Integer = cbo1_3.Text

        If cbo1_3.Text = 0 Then


            txt1_3.BackColor = Color.Yellow

        ElseIf cbo1_3.Text > 0 Then


            txt1_3.BackColor = Color.White


        End If




    End Sub


    Private Sub cbo2_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo2_1.SelectedIndexChanged



        If cbo2_1.Text = 0 Then


            txt2_1.BackColor = Color.Yellow

        ElseIf cbo2_1.Text > 0 Then


            txt2_1.BackColor = Color.White


        End If




    End Sub

    Private Sub cbo3_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_1.SelectedIndexChanged

        If cbo3_1.Text = 0 Then


            txt3_1.BackColor = Color.Yellow

        ElseIf cbo3_1.Text > 0 Then


            txt3_1.BackColor = Color.White


        End If



    End Sub



    Private Sub cbo3_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_2.SelectedIndexChanged


        If cbo3_2.Text = 0 Then


            txt3_2.BackColor = Color.Yellow

        ElseIf cbo3_2.Text > 0 Then


            txt3_2.BackColor = Color.White


        End If



    End Sub

    Private Sub cbo3_3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_3.SelectedIndexChanged


        If cbo3_3.Text = 0 Then


            txt3_3.BackColor = Color.Yellow

        ElseIf cbo3_3.Text > 0 Then


            txt3_3.BackColor = Color.White


        End If





    End Sub

    Private Sub cbo3_4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_4.SelectedIndexChanged
        If cbo3_4.Text = 0 Then


            txt3_4.BackColor = Color.Yellow

        ElseIf cbo3_4.Text > 0 Then


            txt3_4.BackColor = Color.White


        End If


    End Sub

    Private Sub cbo3_5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_5.SelectedIndexChanged


        If cbo3_5.Text = 0 Then


            txt3_5.BackColor = Color.Yellow

        ElseIf cbo3_5.Text > 0 Then


            txt3_5.BackColor = Color.White


        End If


    End Sub

    Private Sub cbo3_6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_6.SelectedIndexChanged


        If cbo3_6.Text = 0 Then


            txt3_6.BackColor = Color.Yellow

        ElseIf cbo3_6.Text > 0 Then


            txt3_6.BackColor = Color.White


        End If






    End Sub

    Private Sub cbo3_7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_7.SelectedIndexChanged

        If cbo3_7.Text = 0 Then


            txt3_7.BackColor = Color.Yellow

        ElseIf cbo3_7.Text > 0 Then


            txt3_7.BackColor = Color.White


        End If


    End Sub

    Private Sub cbo3_8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_8.SelectedIndexChanged


        If cbo3_8.Text = 0 Then


            txt3_8.BackColor = Color.Yellow

        ElseIf cbo3_8.Text > 0 Then


            txt3_8.BackColor = Color.White


        End If


    End Sub


    Private Sub cbo3_9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo3_9.SelectedIndexChanged


        If cbo3_9.Text = 0 Then


            txt3_9.BackColor = Color.Yellow

        ElseIf cbo3_9.Text > 0 Then


            txt3_9.BackColor = Color.White


        End If


    End Sub



    Private Sub Cbo4_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cbo4_1.SelectedIndexChanged

        If Cbo4_1.Text = 0 Then


            txt4_1.BackColor = Color.Yellow

        ElseIf cbo4_1.Text > 0 Then


            txt4_1.BackColor = Color.White


        End If
    End Sub

    Private Sub cbo4_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo4_2.SelectedIndexChanged

        If cbo4_2.Text = 0 Then


            txt4_2.BackColor = Color.Yellow

        ElseIf cbo4_2.Text > 0 Then


            txt4_2.BackColor = Color.White


        End If
    End Sub

    Private Sub cbo4_3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo4_3.SelectedIndexChanged

        If cbo4_3.Text = 0 Then


            txt4_3.BackColor = Color.Yellow

        ElseIf cbo4_3.Text > 0 Then


            txt4_3.BackColor = Color.White


        End If
    End Sub

    Private Sub cbo5_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo5_1.SelectedIndexChanged

        If cbo5_1.Text = 0 Then


            txt5_1.BackColor = Color.Yellow

        ElseIf cbo5_1.Text > 0 Then


            txt5_1.BackColor = Color.White


        End If
    End Sub

    Private Sub cbo5_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo5_2.SelectedIndexChanged

        If cbo5_2.Text = 0 Then


            txt5_2.BackColor = Color.Yellow

        ElseIf cbo5_2.Text > 0 Then


            txt5_2.BackColor = Color.White


        End If

    End Sub

    Private Sub cbo6_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo6_1.SelectedIndexChanged


        If cbo6_1.Text = 0 Then


            txt6_1.BackColor = Color.Yellow

        ElseIf cbo6_1.Text > 0 Then


            txt6_1.BackColor = Color.White


        End If




    End Sub

    Private Sub cbo6_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo6_2.SelectedIndexChanged


        If cbo6_2.Text = 0 Then


            txt6_2.BackColor = Color.Yellow

        ElseIf cbo6_2.Text > 0 Then


            txt6_2.BackColor = Color.White


        End If

    End Sub

    Private Sub cbo6_3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo6_3.SelectedIndexChanged
        If cbo6_3.Text = 0 Then


            txt6_3.BackColor = Color.Yellow

        ElseIf cbo6_3.Text > 0 Then


            txt6_3.BackColor = Color.White


        End If

    End Sub



    Private Sub cbo6_4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo6_4.SelectedIndexChanged
        If cbo6_4.Text = 0 Then


            txt6_4.BackColor = Color.Yellow

        ElseIf cbo6_4.Text > 0 Then


            txt6_4.BackColor = Color.White


        End If

    End Sub

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

    Private Sub QAConsuACallScorcard_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing


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

    Private Sub cboAutoFail_CheckStateChanged(sender As Object, e As EventArgs) Handles cboAutoFail.CheckStateChanged

        If cboAutoFail.CheckState = CheckState.Checked Then


            MsgBox("Are you sure you want to Auto Fail this agent? This will give a score of a 0, but the weights will still be recorded.")


            cboAF.Visible = True


        ElseIf cboAutoFail.CheckState = CheckState.Unchecked Then


            cboAF.Visible = False

            cboAF.Text = "N/a"

        End If






    End Sub
End Class