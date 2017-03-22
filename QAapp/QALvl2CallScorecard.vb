
Imports System.Threading

Imports Microsoft.Office.Interop

Imports System.Data.OleDb


'Imports i00SpellCheck


Public Class QALvl2CallScorecard


    Dim SQL As String
    Dim con As New OleDbConnection




    Dim One As Integer
    Dim two As Integer
    Dim three As Integer
    Dim Four As Integer
    Dim Five As Integer

    ''Store Call Thread
    Dim StoreCallThread As System.Threading.Thread


    Private Sub QALvl2CallScorecard_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

    Private Sub Time_Tick(sender As Object, e As EventArgs) Handles Time.Tick

        lblDate1.Text = Date.Now.ToString("MMM dd yyyy")


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


                QAEmailScorecard.lblQAauditor1.Text = lblQAauditor1.Text

                QAEmailScorecard.lblQAauditor1.Text = lblQAauditor1.Text


            End If





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





                            End If

                        End If

                    End If

                End If

            End If





        Catch ex As Exception



            MsgBox(ex.Message)

        End Try






    End Sub




    Public Sub store()




        Try

            ''Test 

            ' con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")


            'P Drive 

            con = New System.Data.OleDb.OleDbConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")




            '' Dyanic


            '  con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")




            con.Open()



            Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA-Agent],[QA-Team],[QA-ContactDate],[QA-OrderID],[QA-Date],[QA-Comments],[QA-Opp],[CI-Name],[CI-Account],[CI-Company],[CI-Phone],[CI-Email],[Rev-Date],[Rev-Manager],[Rev-Comments],[Dis-Score],[Dis-Name],[Dis-Notes],[Dis-AppComments],[One-1],[One-2],[One-1Note],[One-2Note],[Two-1],[Two-2],[Two-3],[Two-4],[Two-5],[Two-6],[Two-7],[Two-1Note],[Two-2Note],[Two-3Note],[Two-4Note],[Two-5Note],[Two-6Note],[Two-7Note], [Three-1],[Three-2],[Three-1Note],[Three-2Note],[Four-1],[Four-2],[Four-3],[Four-4],[Four-1Note],[Four-2Note],[Four-3Note],[Four-4Note],[Five-1],[Five-2],[Five-3],[Five-4],[Five-5],[Five-1Note],[Five-2Note],[Five-3Note],[Five-4Note],[Five-5Note],[QAScore],[JIRA],[UserID],[AutoFail],[Auditor]) Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

            Using cmd As New OleDbCommand(SQL, con)



                cmd.Parameters.AddWithValue("@p1", lblSRNumber1.Text)
                cmd.Parameters.AddWithValue("@p2", Form2.txtContactPhone.Text)
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


                cmd.Parameters.AddWithValue("@p25", txt1_1.Text)
                cmd.Parameters.AddWithValue("@p26", txt1_2.Text)


                cmd.Parameters.AddWithValue("@p27", cbo2_1.Text)

                cmd.Parameters.AddWithValue("@p28", cbo2_2.Text)
                cmd.Parameters.AddWithValue("@p29", cbo2_3.Text)
                cmd.Parameters.AddWithValue("@p30", cbo2_4.Text)
                cmd.Parameters.AddWithValue("@p31", cbo2_5.Text)
                cmd.Parameters.AddWithValue("@p32", cbo2_6.Text)
                cmd.Parameters.AddWithValue("@p33", cbo2_7.Text)


                cmd.Parameters.AddWithValue("@p34", txt2_1.Text)
                cmd.Parameters.AddWithValue("@p35", txt2_2.Text)
                cmd.Parameters.AddWithValue("@p36", txt2_3.Text)
                cmd.Parameters.AddWithValue("@p37", txt2_4.Text)
                cmd.Parameters.AddWithValue("@p38", txt2_5.Text)
                cmd.Parameters.AddWithValue("@p39", txt2_6.Text)
                cmd.Parameters.AddWithValue("@p40", txt2_7.Text)

                cmd.Parameters.AddWithValue("@p41", cbo3_1.Text)
                cmd.Parameters.AddWithValue("@p42", cbo3_2.Text)

                cmd.Parameters.AddWithValue("@p43", txt3_1.Text)
                cmd.Parameters.AddWithValue("@p44", txt3_2.Text)


                cmd.Parameters.AddWithValue("@p45", Cbo4_1.Text)
                cmd.Parameters.AddWithValue("@p46", cbo4_2.Text)
                cmd.Parameters.AddWithValue("@p47", cbo4_3.Text)
                cmd.Parameters.AddWithValue("@p48", cbo4_4.Text)


                cmd.Parameters.AddWithValue("@p49", txt4_1.Text)
                cmd.Parameters.AddWithValue("@p50", txt4_2.Text)
                cmd.Parameters.AddWithValue("@p51", txt4_3.Text)
                cmd.Parameters.AddWithValue("@p52", txt4_4.Text)

                cmd.Parameters.AddWithValue("@p53", cbo5_1.Text)
                cmd.Parameters.AddWithValue("@p54", cbo5_2.Text)
                cmd.Parameters.AddWithValue("@p55", cbo5_3.Text)
                cmd.Parameters.AddWithValue("@p56", cbo5_4.Text)
                cmd.Parameters.AddWithValue("@p57", cbo5_5.Text)




                cmd.Parameters.AddWithValue("@p58", txt5_1.Text)
                cmd.Parameters.AddWithValue("@p59", txt5_2.Text)
                cmd.Parameters.AddWithValue("@p60", txt5_3.Text)
                cmd.Parameters.AddWithValue("@p61", txt5_4.Text)
                cmd.Parameters.AddWithValue("@p62", txt5_5.Text)

                cmd.Parameters.AddWithValue("@p63", lblQAScore1.Text)
                cmd.Parameters.AddWithValue("@p64", txtgjira.Text)
                cmd.Parameters.AddWithValue("@p65", txtguser.Text)




                If cboAutoFail.Checked Then

                    cmd.Parameters.AddWithValue("@p66", cboAF.Text)

                Else

                    cmd.Parameters.AddWithValue("@p66", "N/a")

                End If



                cmd.Parameters.AddWithValue("@p67", lblQAauditor1.Text)


                cmd.ExecuteNonQuery()

                con.Close()



            End Using


            ' MsgBox("Info saved")


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try

    End Sub


    Public Sub QAlvl2CalldisableControls()


        cbo1_1.Enabled = False
        cbo1_2.Enabled = False


        cbo2_1.Enabled = False
        cbo2_2.Enabled = False
        cbo2_3.Enabled = False
        cbo2_4.Enabled = False
        cbo2_5.Enabled = False
        cbo2_6.Enabled = False
        cbo2_7.Enabled = False



        cbo3_1.Enabled = False
        cbo3_2.Enabled = False


        Cbo4_1.Enabled = False
        cbo4_2.Enabled = False
        cbo4_3.Enabled = False
        cbo4_4.Enabled = False

        cbo5_1.Enabled = False
        cbo5_2.Enabled = False
        cbo5_3.Enabled = False
        cbo5_4.Enabled = False
        cbo5_5.Enabled = False

        ''reset Textboxes

        txt1_1.Enabled = False
        txt1_2.Enabled = False




        txt2_1.Enabled = False
        txt2_2.Enabled = False
        txt2_3.Enabled = False
        txt2_4.Enabled = False
        txt2_5.Enabled = False
        txt2_6.Enabled = False
        txt2_7.Enabled = False




        txt3_1.Enabled = False
        txt3_2.Enabled = False




        txt4_1.Enabled = False
        txt4_2.Enabled = False
        txt4_3.Enabled = False
        txt4_4.Enabled = False


        txt5_1.Enabled = False
        txt5_2.Enabled = False
        txt5_3.Enabled = False
        txt5_4.Enabled = False
        txt5_5.Enabled = False



    End Sub

    Public Sub resetatglance()

        ''Reset Scorecard at a glance info




        lblAgentName1.Text = "N/a"
        lblAgentTeam1.Text = "N/a"
        lblSRNumber1.Text = "N/a"
        lblContactType1.Text = "N/a"
        lblJIRA.Text = "N/a"
        lblUserID.Text = "N/a"
        lblQAScore1.Text = "0"
    End Sub


    Public Sub QAlvl2Callclear()


        ''Reset Comboboxes

        cbo1_1.Text = 3
        cbo1_2.Text = 2


        cbo2_1.Text = 4
        cbo2_2.Text = 4
        cbo2_3.Text = 10
        cbo2_4.Text = 4
        cbo2_5.Text = 4
        cbo2_6.Text = 4
        cbo2_7.Text = 10

        cbo3_1.Text = 15
        cbo3_2.Text = 20


        Cbo4_1.Text = 7
        cbo4_2.Text = 8
        cbo4_3.Text = 2
        cbo4_4.Text = 3


        cbo5_1.Text = 2
        cbo5_2.Text = 2
        cbo5_3.Text = 2
        cbo5_4.Text = 2
        cbo5_5.Text = 2





        ''reset Textboxes

        txt1_1.Clear()
        txt1_2.Clear()



        txt2_1.Clear()
        txt2_2.Clear()
        txt2_3.Clear()
        txt2_4.Clear()
        txt2_5.Clear()
        txt2_6.Clear()
        txt2_7.Clear()




        txt3_1.Clear()
        txt3_2.Clear()



        txt4_1.Clear()
        txt4_2.Clear()
        txt4_3.Clear()
        txt4_4.Clear()

        txt5_1.Clear()
        txt5_2.Clear()
        txt5_3.Clear()
        txt5_4.Clear()
        txt5_5.Clear()

        txtQAAOO.Clear()
        txtQACom.Clear()




        lblQAScore1.Visible = False

    End Sub

    Public Sub QAlvl2CallEnable()




        ''Reset Comboboxes

        cbo1_1.Enabled = True
        cbo1_2.Enabled = True


        cbo2_1.Enabled = True
        cbo2_2.Enabled = True
        cbo2_3.Enabled = True
        cbo2_4.Enabled = True
        cbo2_5.Enabled = True
        cbo2_6.Enabled = True
        cbo2_7.Enabled = True


        cbo3_1.Enabled = True
        cbo3_2.Enabled = True


        Cbo4_1.Enabled = True
        cbo4_2.Enabled = True
        cbo4_3.Enabled = True
        cbo4_4.Enabled = True

        cbo5_1.Enabled = True
        cbo5_2.Enabled = True
        cbo5_3.Enabled = True
        cbo5_4.Enabled = True
        cbo5_5.Enabled = True



        ''reset Textboxes

        txt1_1.Enabled = True
        txt1_2.Enabled = True



        txt2_1.Enabled = True
        txt2_2.Enabled = True
        txt2_3.Enabled = True
        txt2_4.Enabled = True
        txt2_5.Enabled = True
        txt2_6.Enabled = True
        txt2_7.Enabled = True


        txt3_1.Enabled = True
        txt3_2.Enabled = True


        txt4_1.Enabled = True
        txt4_2.Enabled = True
        txt4_3.Enabled = True

        txt5_1.Enabled = True
        txt5_2.Enabled = True
        txt5_3.Enabled = True
        txt5_4.Enabled = True
        txt5_5.Enabled = True


    End Sub

    Public Sub QaTotalScore()


        '  Dim strQaScoreTotal As String
        Dim intQascoreTotal As Integer


        Dim int1_1 As Integer = cbo1_1.Text
        Dim int1_2 As Integer = cbo1_2.Text

        Dim int2_1 As Integer = cbo2_1.Text
        Dim int2_2 As Integer = cbo2_2.Text
        Dim int2_3 As Integer = cbo2_3.Text
        Dim int2_4 As Integer = cbo2_4.Text
        Dim int2_5 As Integer = cbo2_5.Text
        Dim int2_6 As Integer = cbo2_6.Text
        Dim int2_7 As Integer = cbo2_7.Text




        Dim int3_1 As Integer = cbo3_1.Text
        Dim int3_2 As Integer = cbo3_2.Text


        Dim int4_1 As Integer = Cbo4_1.Text
        Dim int4_2 As Integer = cbo4_2.Text
        Dim int4_3 As Integer = cbo4_3.Text
        Dim int4_4 As Integer = cbo4_4.Text






        Dim int5_1 As Integer = cbo5_1.Text
        Dim int5_2 As Integer = cbo5_2.Text
        Dim int5_3 As Integer = cbo5_3.Text
        Dim int5_4 As Integer = cbo5_4.Text
        Dim int5_5 As Integer = cbo5_5.Text


        One = int1_1 + int1_2

        two = int2_1 + int2_2 + int2_3 + int2_4 + int2_5 + int2_6 + int2_7

        three = int3_1 + int3_2

        Four = int4_1 + int4_2 + int4_3 + int4_4

        Five = int5_1 + int5_2 + int5_3 + int5_4 + int5_5














        intQascoreTotal = int1_1 + int1_2 + int2_1 + int2_2 + int2_3 + int2_4 + int2_5 + int2_6 + int2_7 + int3_1 + int3_2 + int4_1 + int4_2 + int4_3 + int4_4 + int5_1 + int5_2 + int5_3 + int5_4 + int5_5
        lblQAScore1.Text = intQascoreTotal

        lblQAScore1.Visible = True

    End Sub

    Public Sub QAExcell()




        Try



            Dim oExcel As Object = CreateObject("Excel.Application")



            ''Test

            '  Dim oBook As Object = oExcel.Workbooks.Open("C:\Users\playe\Desktop\QA\ScoreCard Excell\lvl2CallSc.xlsx")

            '' P Drive

            '   Dim oBook As Object = oExcel.Workbooks.Open("P:\SPC\QA\lvl2CallSc.xlsx")


            '' Dynamic

            Dim oBook As Object = oExcel.Workbooks.Open(lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\lvl2CallSc.xlsx")





            Dim oSheet As Object = oBook.Worksheets("lvl2CallSc")  'or oBook.Worksheets("SheetName")








            oSheet.Range("C3").Value = "" & One


            oSheet.Range("C4").Value = "" & cbo1_1.Text
            oSheet.Range("C5").Value = "" & cbo1_2.Text



            oSheet.Range("D4").Value = "" & txt1_1.Text
            oSheet.Range("D5").Value = "" & txt1_2.Text

            oSheet.Range("C6").Value = "" & two

            oSheet.Range("C7").Value = "" & cbo2_1.Text
            oSheet.Range("C8").Value = "" & cbo2_2.Text
            oSheet.Range("C9").Value = "" & cbo2_3.Text
            oSheet.Range("C10").Value = "" & cbo2_4.Text
            oSheet.Range("C11").Value = "" & cbo2_5.Text
            oSheet.Range("C12").Value = "" & cbo2_6.Text
            oSheet.Range("C13").Value = "" & cbo2_7.Text

            oSheet.Range("D7").Value = "" & txt2_1.Text
            oSheet.Range("D8").Value = "" & txt2_2.Text
            oSheet.Range("D9").Value = "" & txt2_3.Text
            oSheet.Range("D10").Value = "" & txt2_4.Text
            oSheet.Range("D11").Value = "" & txt2_5.Text
            oSheet.Range("D12").Value = "" & txt2_6.Text
            oSheet.Range("D13").Value = "" & txt2_7.Text




            oSheet.Range("C14").Value = "" & three

            oSheet.Range("C15").Value = "" & cbo3_1.Text
            oSheet.Range("C16").Value = "" & cbo3_2.Text



            oSheet.Range("D15").Value = "" & txt3_1.Text
            oSheet.Range("D16").Value = "" & txt3_2.Text


            oSheet.Range("C17").Value = "" & Four

            oSheet.Range("C18").Value = "" & Cbo4_1.Text
            oSheet.Range("C19").Value = "" & cbo4_2.Text
            oSheet.Range("C20").Value = "" & cbo4_3.Text
            oSheet.Range("C21").Value = "" & cbo4_4.Text

            oSheet.Range("D18").Value = "" & txt4_1.Text
            oSheet.Range("D19").Value = "" & txt4_2.Text
            oSheet.Range("D20").Value = "" & txt4_3.Text
            oSheet.Range("D21").Value = "" & txt4_4.Text

            oSheet.Range("C22").Value = "" & Five


            oSheet.Range("C23").Value = "" & cbo5_1.Text
            oSheet.Range("C24").Value = "" & cbo5_2.Text
            oSheet.Range("C25").Value = "" & cbo5_3.Text
            oSheet.Range("C26").Value = "" & cbo5_4.Text
            oSheet.Range("C27").Value = "" & cbo5_5.Text


            oSheet.Range("D23").Value = "" & txt5_1.Text
            oSheet.Range("D24").Value = "" & txt5_2.Text
            oSheet.Range("D25").Value = "" & txt5_3.Text
            oSheet.Range("D26").Value = "" & txt5_4.Text
            oSheet.Range("D27").Value = "" & txt5_5.Text



            oSheet.Range("C28").Value = lblQAScore1.Text

            oSheet.Range("A47").Value = lblQAScore1.Text
            oSheet.Range("A61").Value = lblQAScore1.Text






            oSheet.Range("B30").Value = lblSRNumber1.Text
            oSheet.Range("B31").Value = lblContactID1.Text
            oSheet.Range("B32").Value = lblContactType1.Text
            oSheet.Range("B33").Value = "" & lblAgentName1.Text
            oSheet.Range("B34").Value = "" & lblAgentTeam1.Text
            oSheet.Range("B35").Value = Form2.DateTimePicker1.Text
            oSheet.Range("B36").Value = Form2.txtOrderID.Text
            oSheet.Range("B37").Value = "" & txtgnamebox.Text
            oSheet.Range("B38").Value = "" & txtgemail.Text
            oSheet.Range("B39").Value = "" & txtgphone.Text
            oSheet.Range("B40").Value = "" & txtgcompany.Text
            oSheet.Range("B41").Value = "" & txtgacc.Text
            oSheet.Range("B42").Value = "" & txtgjira.Text
            oSheet.Range("B43").Value = "" & txtguser.Text
            oSheet.Range("B44").Value = "" & lblQAauditor1.Text
            oSheet.Range("B45").Value = "" & lblDate1.Text






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

    Private Sub D_Click(sender As Object, e As EventArgs)

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



            StoreCallThread = New System.Threading.Thread(AddressOf store)
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

        QAlvl2Callclear()

        ''Transfer Qa Name to Wasetupform


        '  Form2.lblQAauditor.Text = lblQAauditor1.Text


        ''Reable buttons

        QAlvl2CallEnable()


        Me.Hide()


        ProgressBar1.Value = 0
        lblprogr.Text = 0

        txtQACom.BackColor = Color.White


        cboAF.Visible = False

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

    Private Sub Button22_Click(sender As Object, e As EventArgs)
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

    Private Sub Button23_Click(sender As Object, e As EventArgs)
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

    Private Sub Button24_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt2_2.Text = "" Then

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
            Clipboard.SetDataObject(txt2_2.Text)

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
                    txt2_2.Text = CType(iData.GetData(DataFormats.Text),
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
            If txt2_3.Text = "" Then

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
            Clipboard.SetDataObject(txt2_3.Text)

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
                    txt2_3.Text = CType(iData.GetData(DataFormats.Text),
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
            If txt2_4.Text = "" Then

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
            Clipboard.SetDataObject(txt2_4.Text)

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
                    txt2_4.Text = CType(iData.GetData(DataFormats.Text),
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
            If txt2_5.Text = "" Then

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
            Clipboard.SetDataObject(txt2_5.Text)

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
                    txt2_5.Text = CType(iData.GetData(DataFormats.Text),
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
            If txt2_6.Text = "" Then

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
            Clipboard.SetDataObject(txt2_6.Text)

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
                    txt2_6.Text = CType(iData.GetData(DataFormats.Text),
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
            If txt2_7.Text = "" Then

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
            Clipboard.SetDataObject(txt2_7.Text)

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
                    txt2_7.Text = CType(iData.GetData(DataFormats.Text),
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

    Private Sub Button10_Click(sender As Object, e As EventArgs)
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

    Private Sub Button9_Click(sender As Object, e As EventArgs)
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

    Private Sub Button8_Click(sender As Object, e As EventArgs)
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

    Private Sub Button7_Click(sender As Object, e As EventArgs)
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

    Private Sub Button4_Click(sender As Object, e As EventArgs)

        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt4_4.Text = "" Then

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
            Clipboard.SetDataObject(txt4_4.Text)

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
                    txt4_4.Text = CType(iData.GetData(DataFormats.Text),
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

    Private Sub Button2_Click(sender As Object, e As EventArgs)
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

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Object
            Dim objTempDoc As Object

            ' Declare an IDataObject to hold the data returned from the 
            ' clipboard.
            Dim iData As IDataObject

            ' If there is no data to spell check, then exit sub here.
            If txt5_3.Text = "" Then

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
            Clipboard.SetDataObject(txt5_3.Text)

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
                    txt5_3.Text = CType(iData.GetData(DataFormats.Text),
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
            If txt5_4.Text = "" Then

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
            Clipboard.SetDataObject(txt5_4.Text)

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
                    txt5_4.Text = CType(iData.GetData(DataFormats.Text),
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
            If txt5_5.Text = "" Then

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
            Clipboard.SetDataObject(txt5_5.Text)

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
                    txt5_5.Text = CType(iData.GetData(DataFormats.Text),
                        String)
                End If
                .Saved = True

            End With

            objWord.Quit()




        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
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

    Private Sub QALvl2CallScorecard_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing


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