
Imports Microsoft.Office.Interop.Excel


Imports System.Data.OleDb


Public Class QAScorecard


    Dim SQL As String
    Dim con As New OleDb.OleDbConnection



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

    

        Me.ActiveControl = cbo1_1




    End Sub

 

    Private Sub btnQaSetup_Click(sender As Object, e As EventArgs) Handles btnQaSetup.Click



        QaSetup.Show()


        ''Transfer Qa auditor Name to form

        QaSetup.lblQAauditor.Text = lblQAauditor.Text




    End Sub

    Public Sub Store()




        Try

            con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Documents\Visual Studio 2012\Projects\Qa\QA.accdb")



            con.Open()



            Dim SQL As String = "INSERT INTO [QAScorecardDB] ([SR], [ContactID], [CType],[QA-Agent],[QA-Team],[QA-ContactDate],[QA-OrderID],[QA-Date],[QA-Comments],[QA-Opp],[CI-Name],[CI-Account],[CI-Company],[CI-Phone],[CI-Email],[Rev-Date],[Rev-Manager],[Rev-Comments],[Dis-Score],[Dis-Name],[Dis-Notes],[Dis-AppComments],[One-1],[One-2],[One-3],[One-1Note],[One-2Note],[One-3Note],[Two-1],[Two-1Note],[Three-1],[Three-2],[Three-3],[Three-4],[Three-5],[Three-6],[Three-7],[Three-8],[Three-1Note],[Three-2Note],[Three-3Note],[Three-4Note],[Three-5Note],[Three-6Note],[Three-7Note],[Three-8Note],[Four-1],[Four-2],[Four-3],[Four-1Note],[Four-2Note],[Four-3Note],[Five-1],[Five-2],[Five-1Note],[Five-2Note],[Six-1],[Six-2],[Six-3],[Six-1Note],[Six-2Note],[Six-3Note],[Seven-1],[Seven-2],[Seven-3],[Seven-4],[Seven-5],[Seven-6],[Seven-1Note],[Seven-2Note],[Seven-3Note],[Seven-4Note],[Seven-5Note],[Seven-6Note],[QAScore]) Values ( ?, ?, ?, ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

            Using cmd As New OleDbCommand(SQL, con)



                cmd.Parameters.AddWithValue("@p1", lblSRNumber.Text)
                cmd.Parameters.AddWithValue("@p2", QaSetup.txtContactPhone.Text)
                cmd.Parameters.AddWithValue("@p3", lblContactType.Text)
                cmd.Parameters.AddWithValue("@p4", lblAgentName.Text)
                cmd.Parameters.AddWithValue("@p5", lblAgentTeam.Text)
                cmd.Parameters.AddWithValue("@p6", QaSetup.DateTimePicker1.Value)
                cmd.Parameters.AddWithValue("@p7", QaSetup.txtOrderID.Text)
                cmd.Parameters.AddWithValue("@p8", "Date")
                '   cmd.Parameters.AddWithValue("@p9", txtCommentes.Text)
                '  cmd.Parameters.AddWithValue("@p10", txtComments2.Text)
                cmd.Parameters.AddWithValue("@p11", QaSetup.txtContactID.Text)
                cmd.Parameters.AddWithValue("@p12", QaSetup.txtAccountNum.Text)
                cmd.Parameters.AddWithValue("@p13", QaSetup.txtCompany.Text)
                cmd.Parameters.AddWithValue("@p14", QaSetup.txtContactName.Text)
                cmd.Parameters.AddWithValue("@p15", QaSetup.txtContactEmail.Text)
                cmd.Parameters.AddWithValue("@p16", "Rev date")
                '  cmd.Parameters.AddWithValue("@p17", cboReviewManger.Text)
                ' cmd.Parameters.AddWithValue("@p18", txtReviewComments.Text)
                ' cmd.Parameters.AddWithValue("@p19", txtDisputeScore.Text)
                ' cmd.Parameters.AddWithValue("@p20", txtDisputerName.Text)
                '  cmd.Parameters.AddWithValue("@p21", txtDisputeComments.Text)
                ' cmd.Parameters.AddWithValue("@p22", txtDisputeAppComments.Text)


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



                cmd.Parameters.AddWithValue("@p39", txt3_1.Text)
                cmd.Parameters.AddWithValue("@p40", txt3_2.Text)
                cmd.Parameters.AddWithValue("@p41", txt3_3.Text)
                cmd.Parameters.AddWithValue("@p42", txt3_4.Text)
                cmd.Parameters.AddWithValue("@p43", txt3_5.Text)
                cmd.Parameters.AddWithValue("@p44", txt3_6.Text)
                cmd.Parameters.AddWithValue("@p45", txt3_7.Text)
                cmd.Parameters.AddWithValue("@p46", txt3_8.Text)


                cmd.Parameters.AddWithValue("@p47", Cbo4_1.Text)
                cmd.Parameters.AddWithValue("@p48", cbo4_2.Text)
                cmd.Parameters.AddWithValue("@p49", cbo4_3.Text)

                cmd.Parameters.AddWithValue("@p50", txt4_1.Text)
                cmd.Parameters.AddWithValue("@p51", txt4_2.Text)
                cmd.Parameters.AddWithValue("@p52", txt4_3.Text)



                cmd.Parameters.AddWithValue("@p53", cbo5_1.Text)
                cmd.Parameters.AddWithValue("@p54", cbo5_2.Text)


                cmd.Parameters.AddWithValue("@p55", txt5_1.Text)
                cmd.Parameters.AddWithValue("@p56", txt5_2.Text)



                cmd.Parameters.AddWithValue("@p57", cbo6_1.Text)
                cmd.Parameters.AddWithValue("@p58", cbo6_2.Text)
                cmd.Parameters.AddWithValue("@p59", cbo6_3.Text)



                cmd.Parameters.AddWithValue("@p60", txt6_1.Text)
                cmd.Parameters.AddWithValue("@p61", txt6_2.Text)
                cmd.Parameters.AddWithValue("@p62", txt6_3.Text)


                cmd.Parameters.AddWithValue("@p63", cbo7_1.Text)
                cmd.Parameters.AddWithValue("@p64", cbo7_2.Text)
                cmd.Parameters.AddWithValue("@p65", cbo7_3.Text)
                cmd.Parameters.AddWithValue("@p66", cbo7_4.Text)
                cmd.Parameters.AddWithValue("@p67", cbo7_5.Text)
                cmd.Parameters.AddWithValue("@p68", cbo7_6.Text)

                cmd.Parameters.AddWithValue("@p69", txt7_1.Text)
                cmd.Parameters.AddWithValue("@p70", txt7_2.Text)
                cmd.Parameters.AddWithValue("@p71", txt7_3.Text)
                cmd.Parameters.AddWithValue("@p72", txt7_4.Text)
                cmd.Parameters.AddWithValue("@p73", txt7_5.Text)
                cmd.Parameters.AddWithValue("@p74", txt7_6.Text)
                cmd.Parameters.AddWithValue("@p75", lblQAScore.Text)





                cmd.ExecuteNonQuery()

                con.Close()



            End Using


            ' MsgBox("Info saved")


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub


    Public Sub QAclear()


        ''Reset Comboboxes

        cbo1_1.Text = 2
        cbo1_2.Text = 1
        cbo1_3.Text = 2

        cbo2_1.Text = 15

        cbo3_1.Text = 2
        cbo3_2.Text = 1
        cbo3_3.Text = 3
        cbo3_4.Text = 4
        cbo3_5.Text = 3
        cbo3_6.Text = 3
        cbo3_7.Text = 3
        cbo3_8.Text = 1

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




        '     txtCommentes.Clear()



        '    txtComments2.Clear()


        '  cboReviewManger.Text = ""
        '  txtReviewComments.Clear()
        ' txtDisputeScore.Clear()
        '  txtDisputerName.Clear()
        ' txtDisputeComments.Clear()
        '  txtDisputeAppComments.Clear()
        '  txtDisputeScore.Clear()
        '  txtDisputeApprovalScore.Clear()



    End Sub




    Public Sub SavetoExcel()

        Try



            Dim oExcel As Object = CreateObject("Excel.Application")

            Dim oBook As Object = oExcel.Workbooks.Open("C:\Users\playe\Documents\Visual Studio 2012\Projects\Qa\MainQAExcell2.xlsx")

            Dim oSheet As Object = oBook.Worksheets("CallSheet")  'or oBook.Worksheets("SheetName")

            'e.g. Read value from A2 cell to TextBox1 

            '   txt1_1.Text = oSheet.Range("F7").Value

            'e.g. Write value from TextBox1 to A2 cell

            oSheet.Range("D3").Value = "" & txt1_1.Text



            'Save this Excel document / set fikder for all Qa scores

            oBook.SaveAs("C:\Users\playe\Documents\Visual Studio 2012\Projects\Qa\" & lblAgentName.Text & " " & "QA Scorecard.xlsx")



            oExcel.Quit()



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






        intQascoreTotal = int1_1 + int1_2 + int1_3 + int2_1 + int3_1 + int3_2 + int3_3 + int3_4 + int3_5 + int3_6 + int3_7 + int3_8 + int4_1 + int4_2 + int4_3 + int5_1 + int5_2 + int6_1 + int6_2 + int6_3 + int7_1 + int7_2 + int7_3 + int7_4 + int7_5 + int7_6

        lblQAScore.Text = intQascoreTotal











    End Sub


    Public Sub resetatglance()

        ''Reset Scorecard at a glance info
       
        lblAgentName.Text = "N/a"
        lblAgentTeam.Text = "N/a"
        lblSRNumber.Text = "N/a"
        lblContactType.Text = "N/a"
        lblQAScore.Text = "0"
    End Sub


    Private Sub btnSaveScoreCard_Click(sender As Object, e As EventArgs) Handles btnSaveScoreCard.Click

        Try

            '  If lblStored.Visible = True Then

            '  If MsgBox("Please be advised this Audit has already been saved and scored, proceeding will overwrite previous data", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.Cancel Then


            '  Else

            If QaSetup.cboAgentName.Text = "Agent Name" Or QaSetup.cboContactType.Text = "Contact Type" Then




                MsgBox("Please be advised you must fill out all 'Agent Information' before proceeding", MessageBoxButtons.RetryCancel)


            Else


                If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then



                Else


                    ''Store info to database

                    ' Store()


                    ''Tally Qa Score


                    '  QaTotalScore()

                    ''Show Scorecard


                    '   lblQAScore.Visible = True

                    SavetoExcel()





                    ''Disable Controls
                    ' disableControls()




                    If MsgBox("The QA Scorecard information successfully saved, Do you want to start a new one?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then




                    Else


                        ''Clear and reset the QaSetup Tab

                        QaSetup.Clear()


                        QaSetup.Show()

                        ''Reset Scorecard at a glance info

                        resetatglance()

                        ''Reset scorecard

                        QAclear()

                        ''Transfer Qa Name to Wasetupform


                        QaSetup.lblQAauditor.Text = lblQAauditor.Text


                        ''Reable buttons

                        enablecontrols()




                    End If


                End If


                'End If


            End If



        Catch ex As Exception

            MsgBox(ex.Message)

        End Try






    End Sub


    Public Sub disableControls()


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



        txt4_1.Enabled = False
        txt4_2.Enabled = False
        txt4_3.Enabled = False

        txt5_1.Enabled = False
        txt5_2.Enabled = False




        txt6_1.Enabled = False
        txt6_2.Enabled = False
        txt6_3.Enabled = False


        txt7_1.Enabled = False
        txt7_2.Enabled = False
        txt7_3.Enabled = False
        txt7_4.Enabled = False
        txt7_5.Enabled = False
        txt7_6.Enabled = False




        'txtCommentes.Enabled = False



        '   txtComments2.Enabled = False


        '   cboReviewManger.Enabled = False
        ' txtReviewComments.Enabled = False
        ' txtDisputeScore.Enabled = False
        ' txtDisputerName.Enabled = False
        ' txtDisputeComments.Enabled = False
        'txtDisputeAppComments.Enabled = False
        ' txtDisputeScore.Enabled = False
        ' txtDisputeApprovalScore.Enabled = False



    End Sub


    Public Sub enablecontrols()




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




        '  txtCommentes.Enabled = True




        'txtComments2.Enabled = True


        '   cboReviewManger.Enabled = True
        ' txtReviewComments.Enabled = True
        ' txtDisputeScore.Enabled = True
        ' txtDisputerName.Enabled = True
        ' txtDisputeComments.Enabled = True
        'txtDisputeAppComments.Enabled = True
        ' txtDisputeScore.Enabled = True
        ' txtDisputeApprovalScore.Enabled = True


    End Sub











    Private Sub btnClearScoreCard_Click(sender As Object, e As EventArgs) Handles btnClearScoreCard.Click

        Me.txt1_1.AutoSize = True






    End Sub
End Class
