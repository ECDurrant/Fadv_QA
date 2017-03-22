
Imports System.Threading

Imports System.Data.OleDb


Imports Microsoft.Office.Interop

Imports i00SpellCheck

Public Class QAlvl2EmailScorecard


    Dim SQL As String
    Dim con As New OleDbConnection


    Dim One As Integer
    Dim two As Integer
    Dim three As Integer

    ''Store Call Thread
    Dim StoreCallThread As System.Threading.Thread





    Private Sub QAlvl2EmailScorecard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try


            Me.WindowState = FormWindowState.Maximized
            Me.ActiveControl = cbo1_1



            Time.Enabled = True

            Control.CheckForIllegalCrossThreadCalls = False

            ' Me.EnableControlExtensions()



        Catch ex As Exception



            MsgBox(ex.Message)


        End Try



    End Sub
    Public Sub QaTotalScore()


        '  Dim strQaScoreTotal As String
        Dim intQascoreTotal As Integer


        Dim int1_1 As Integer = cbo1_1.Text


        Dim int2_1 As Integer = cbo2_1.Text
        Dim int2_2 As Integer = cbo2_2.Text
        Dim int2_3 As Integer = cbo2_3.Text
        Dim int2_4 As Integer = cbo2_4.Text


        Dim int3_1 As Integer = cbo3_1.Text
        Dim int3_2 As Integer = cbo3_2.Text
        Dim int3_3 As Integer = cbo3_3.Text
        Dim int3_4 As Integer = cbo3_4.Text
        Dim int3_5 As Integer = cbo3_5.Text



        One = int1_1
        Two = int2_1
        three = int3_1 + int3_2 + int3_3 + int3_4 + int3_5


        intQascoreTotal = int1_1 + int2_1 + int2_2 + int2_3 + int2_4 + int3_1 + int3_2 + int3_3 + int3_4 + int3_5

        lblQAScore.Text = intQascoreTotal

        lblQAScore.Visible = True











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



            Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA-Agent],[QA-Team],[QA-ContactDate],[QA-OrderID],[QA-Date],[QA-Comments],[QA-Opp],[CI-Name],[CI-Account],[CI-Company],[CI-Phone],[CI-Email],[Rev-Date],[Rev-Manager],[Rev-Comments],[Dis-Score],[Dis-Name],[Dis-Notes],[Dis-AppComments],[One-1],[One-1Note],[Two-1],[Two-2],[Two-3],[Two-4],[Two-1Note],[Two-2Note],[Two-3Note],[Two-4Note],[Three-1],[Three-2],[Three-3],[Three-4],[Three-5],[Three-1Note],[Three-2Note],[Three-3Note],[Three-4Note],[Three-5Note],[QAScore],[JIRA],[UserID]) Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

            Using cmd As New OleDbCommand(SQL, con)



                cmd.Parameters.AddWithValue("@p1", lblSRNumber.Text)
                cmd.Parameters.AddWithValue("@p2", lblContactID1.Text)
                cmd.Parameters.AddWithValue("@p3", lblContactType.Text)
                cmd.Parameters.AddWithValue("@p4", lblAgentName.Text)
                cmd.Parameters.AddWithValue("@p5", lblAgentTeam.Text)
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
                cmd.Parameters.AddWithValue("@p24", txt1_1.Text)



                cmd.Parameters.AddWithValue("@p25", cbo2_1.Text)
                cmd.Parameters.AddWithValue("@p26", cbo2_2.Text)
                cmd.Parameters.AddWithValue("@p27", cbo2_3.Text)
                cmd.Parameters.AddWithValue("@p28", cbo2_4.Text)

                cmd.Parameters.AddWithValue("@p29", txt2_1.Text)
                cmd.Parameters.AddWithValue("@p30", txt2_2.Text)
                cmd.Parameters.AddWithValue("@p31", txt2_3.Text)
                cmd.Parameters.AddWithValue("@p32", txt2_4.Text)



                cmd.Parameters.AddWithValue("@p33", cbo3_1.Text)
                cmd.Parameters.AddWithValue("@p34", cbo3_2.Text)
                cmd.Parameters.AddWithValue("@p35", cbo3_3.Text)
                cmd.Parameters.AddWithValue("@p36", cbo3_4.Text)
                cmd.Parameters.AddWithValue("@p37", cbo3_5.Text)


                cmd.Parameters.AddWithValue("@p38", txt3_1.Text)
                cmd.Parameters.AddWithValue("@p39", txt3_2.Text)
                cmd.Parameters.AddWithValue("@p40", txt3_3.Text)
                cmd.Parameters.AddWithValue("@p41", txt3_4.Text)
                cmd.Parameters.AddWithValue("@p42", txt3_5.Text)


                cmd.Parameters.AddWithValue("@p43", lblQAScore.Text)
                cmd.Parameters.AddWithValue("@p44", lblJIRA.Text)
                cmd.Parameters.AddWithValue("@p45", lblUserID.Text)


                If cboAutoFail.Checked Then

                    cmd.Parameters.AddWithValue("@p46", cboAF.Text)

                Else

                    cmd.Parameters.AddWithValue("@p46", "N/a")

                End If


                cmd.Parameters.AddWithValue("@p47", lblQAauditor.Text)


                cmd.ExecuteNonQuery()

                con.Close()



            End Using


            ' MsgBox("Info saved")


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try

    End Sub


    Public Sub QAlvl2EmaildisableControls()


        cbo1_1.Enabled = False


        cbo2_1.Enabled = False
        cbo2_2.Enabled = False
        cbo2_3.Enabled = False
        cbo2_4.Enabled = False


        cbo3_1.Enabled = False
        cbo3_2.Enabled = False
        cbo3_3.Enabled = False
        cbo3_4.Enabled = False
        cbo3_5.Enabled = False



        ''reset Textboxes

        txt1_1.Enabled = False


        txt2_1.Enabled = False
        txt2_2.Enabled = False
        txt2_3.Enabled = False
        txt2_4.Enabled = False

        txt3_1.Enabled = False
        txt3_2.Enabled = False
        txt3_3.Enabled = False
        txt3_4.Enabled = False
        txt3_5.Enabled = False


    End Sub

    Public Sub resetatglance()

        ''Reset Scorecard at a glance info




        lblAgentName.Text = "N/a"
        lblAgentTeam.Text = "N/a"
        lblSRNumber.Text = "N/a"
        lblContactType.Text = "N/a"
        lblQAScore.Text = "0"
        lblJIRA.Text = "N/a"
        lblUserID.Text = "N/a"




    End Sub


    Public Sub QAlvl2Emailclear()


        ''Reset Comboboxes

        cbo1_1.Text = 25



        cbo2_1.Text = 2
        cbo2_2.Text = 10
        cbo2_3.Text = 3
        cbo2_4.Text = 10

        cbo3_1.Text = 3
        cbo3_2.Text = 4
        cbo3_3.Text = 3
        cbo3_4.Text = 20
        cbo3_5.Text = 20

        ''reset Textboxes

        txt1_1.Clear()



        txt2_1.Clear()
        txt2_2.Clear()
        txt2_3.Clear()
        txt2_4.Clear()






        txt3_1.Clear()
        txt3_2.Clear()
        txt3_3.Clear()
        txt3_4.Clear()
        txt3_5.Clear()


        txtQAAOO.Clear()
        txtQACom.Clear()




        lblQAScore.Visible = False

    End Sub

    Public Sub QAlvl2EmailEnable()




        ''Reset Comboboxes

        cbo1_1.Enabled = True



        cbo2_1.Enabled = True
        cbo2_2.Enabled = True
        cbo2_3.Enabled = True
        cbo2_4.Enabled = True



        cbo3_1.Enabled = True
        cbo3_2.Enabled = True
        cbo3_3.Enabled = True
        cbo3_4.Enabled = True
        cbo3_5.Enabled = True



        ''reset Textboxes

        txt1_1.Enabled = True





        txt2_1.Enabled = True
        txt2_2.Enabled = True
        txt2_3.Enabled = True
        txt2_4.Enabled = True


        txt3_1.Enabled = True
        txt3_2.Enabled = True
        txt3_3.Enabled = True
        txt3_4.Enabled = True
        txt3_5.Enabled = True


    End Sub



    Public Sub QAExcell()




        Try



            Dim oExcel As Object = CreateObject("Excel.Application")



            ''Test

            '  Dim oBook As Object = oExcel.Workbooks.Open("C:\Users\playe\Desktop\QA\ScoreCard Excell\lvl2EmailSc.xlsx")

            '' P Drive

            '   Dim oBook As Object = oExcel.Workbooks.Open("P:\SPC\QA\lvl2EmailSc.xlsx")


            '' Dynamic

            Dim oBook As Object = oExcel.Workbooks.Open(lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\lvl2EmailSc.xlsx")




            Dim oSheet As Object = oBook.Worksheets("lvl2EmailSc")  'or oBook.Worksheets("SheetName")








            oSheet.Range("C3").Value = "" & One


            oSheet.Range("C4").Value = "" & cbo1_1.Text


            oSheet.Range("D4").Value = "" & txt1_1.Text


            oSheet.Range("C5").Value = "" & two

            oSheet.Range("C6").Value = "" & cbo2_1.Text
            oSheet.Range("C7").Value = "" & cbo2_2.Text
            oSheet.Range("C8").Value = "" & cbo2_3.Text
            oSheet.Range("C9").Value = "" & cbo2_4.Text

            oSheet.Range("D6").Value = "" & txt2_1.Text
            oSheet.Range("D7").Value = "" & txt2_2.Text
            oSheet.Range("D8").Value = "" & txt2_3.Text
            oSheet.Range("D9").Value = "" & txt2_4.Text



            oSheet.Range("C10").Value = "" & three

            oSheet.Range("C11").Value = "" & cbo3_1.Text
            oSheet.Range("C12").Value = "" & cbo3_2.Text
            oSheet.Range("C13").Value = "" & cbo3_3.Text
            oSheet.Range("C14").Value = "" & cbo3_4.Text
            oSheet.Range("C15").Value = "" & cbo3_5.Text



            oSheet.Range("D11").Value = "" & txt3_1.Text
            oSheet.Range("D12").Value = "" & txt3_2.Text
            oSheet.Range("D13").Value = "" & txt3_3.Text
            oSheet.Range("D14").Value = "" & txt3_4.Text
            oSheet.Range("D15").Value = "" & txt3_5.Text



            oSheet.Range("C16").Value = lblQAScore.Text

            oSheet.Range("A35").Value = txtQACom.Text
            oSheet.Range("A49").Value = txtQAAOO.Text






            oSheet.Range("B18").Value = lblSRNumber.Text
            oSheet.Range("B19").Value = lblContactID1.Text
            oSheet.Range("B20").Value = lblContactType.Text
            oSheet.Range("B21").Value = "" & lblAgentName.Text
            oSheet.Range("B22").Value = "" & lblAgentTeam.Text
            oSheet.Range("B23").Value = Form2.DateTimePicker1.Text
            oSheet.Range("B24").Value = Form2.txtOrderID.Text
            oSheet.Range("B25").Value = "" & txtgnamebox.Text
            oSheet.Range("B26").Value = "" & txtgemail.Text
            oSheet.Range("B27").Value = "" & txtgphone.Text
            oSheet.Range("B28").Value = "" & txtgcompany.Text
            oSheet.Range("B29").Value = "" & txtgacc.Text
            oSheet.Range("B30").Value = "" & txtgjira.Text
            oSheet.Range("B31").Value = "" & txtguser.Text
            oSheet.Range("B32").Value = "" & lblQAauditor.Text
            oSheet.Range("B33").Value = "" & lblDate.Text



            '' Test

            '  oBook.SaveAs("C:\Users\playe\Desktop\QA\" & "SR#" & lblSRNumber1.Text & "_" & lblAgentName1.Text & "_" & lblDate1.Text & " QA Scorecard.xlsx")



            '' P drive

            ' oBook.SaveAs("P:\SPC\QA\" & lblSRNumber1.Text & "_" & lblAgentName1.Text & "_" & lblDate1.Text & " QA Scorecard.xlsx")


            '' Dynamic

            oBook.SaveAs(lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA2\" & "SR#" & lblSRNumber.Text & "_" & lblAgentName.Text & "_" & lblDate.Text & " QA Scorecard.xlsx")






            oExcel.Quit()





        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub Time_Tick(sender As Object, e As EventArgs) Handles Time.Tick


        lblDate.Text = Date.Now.ToString("MMM dd yyyy")



    End Sub

    Private Sub btnSaveScoreCard_Click(sender As Object, e As EventArgs) Handles btnSaveScoreCard.Click


        Try



            '   If Form2.cboAgentName.Text = "Agent Name" Or Form2.cboContactType.Text = "Contact Type" Then
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


                                lblQAScore.Text = "0"


                                lblQAScore.Visible = True


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

    Private Sub Time_Tick_1(sender As Object, e As EventArgs) Handles Time.Tick

        lblDate.Text = Date.Now.ToString("MMM dd yyyy")


    End Sub

    Private Sub btnQaSetup_Click(sender As Object, e As EventArgs) Handles btnQaSetup.Click


        Try



            If lblQAScore.Visible = True Then

                reset()


            Else




                Form2.Show()

            End If



        Catch ex As Exception


            MsgBox(ex.Message)


        End Try



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



        lblQAScore.Visible = True

        If MsgBox(lblAgentName.Text & " " & "" & "scored a total of" & " " & lblQAScore.Text & " " & "points on this QA audit,would you like to start a new one?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then


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

        QAlvl2Emailclear()

        ''Transfer Qa Name to Wasetupform


        '   Form2.lblQAauditor.Text = lblQAauditor.Text


        ''Reable buttons

        QAlvl2EmailEnable()


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

    Private Sub Button6_Click(sender As Object, e As EventArgs)
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

    Private Sub Button3_Click(sender As Object, e As EventArgs)
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

    Private Sub Button2_Click(sender As Object, e As EventArgs)
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

    Private Sub Button7_Click(sender As Object, e As EventArgs)
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


    Private Sub cbo1_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo1_1.SelectedIndexChanged

        Dim int1_1 As Integer = cbo1_1.Text

        If cbo1_1.Text = 0 Then


            txt1_1.BackColor = Color.Yellow

        ElseIf cbo1_1.Text > 0 Then


            txt1_1.BackColor = Color.White


        End If

    End Sub


    Private Sub cbo2_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo2_1.SelectedIndexChanged



        If cbo2_1.Text = 0 Then


            txt2_1.BackColor = Color.Yellow

        ElseIf cbo2_1.Text > 0 Then


            txt2_1.BackColor = Color.White


        End If


    End Sub

    Private Sub cbo2_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo2_2.SelectedIndexChanged



        If cbo2_2.Text = 0 Then


            txt2_2.BackColor = Color.Yellow

        ElseIf cbo2_2.Text > 0 Then


            txt2_2.BackColor = Color.White


        End If

    End Sub

    Private Sub cbo2_3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo2_3.SelectedIndexChanged



        If cbo2_3.Text = 0 Then


            txt2_3.BackColor = Color.Yellow

        ElseIf cbo2_3.Text > 0 Then


            txt2_3.BackColor = Color.White


        End If

    End Sub

    Private Sub cbo2_4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo2_4.SelectedIndexChanged



        If cbo2_4.Text = 0 Then


            txt2_4.BackColor = Color.Yellow

        ElseIf cbo2_4.Text > 0 Then


            txt2_4.BackColor = Color.White


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

    Private Sub QAlvl2EmailScorecard_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

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