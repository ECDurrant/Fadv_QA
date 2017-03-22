

Imports System.Threading

Imports Microsoft.Office.Interop

Imports DevExpress.XtraSpellChecker

Imports System.Globalization

Imports System.Data.SqlClient

Imports System.Data.OleDb

'Imports i00SpellCheck


Imports System.Net

Imports System.Net.Security

Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Mail

Public Class QAChatRev


    ''Store Call Thread
    Dim StoreCallThread As System.Threading.Thread

    'Store Call Thread
    Dim ToExcell As System.Threading.Thread

    Dim SQL As String
    Dim con As New SqlConnection


    Dim One As Integer
    Dim two As Integer
    Dim three As Integer
    Dim Four As Integer
    Dim Five As Integer
    Dim Six As Integer
    Dim seven As Integer


    Dim Desk = My.Computer.FileSystem.SpecialDirectories.Desktop


    Dim dic_en_US As SpellCheckerOpenOfficeDictionary = New SpellCheckerOpenOfficeDictionary


    Dim spellcheckDIR As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)

    Dim en_USaffPath = IO.Path.Combine(spellcheckDIR.DirectoryName, "en_US.aff")
    Dim en_USdicPath = IO.Path.Combine(spellcheckDIR.DirectoryName, "en_US.dic")




    Public Sub buttonEnables()

        btnApproval.Enabled = True

        btnSaveScoreCard.Enabled = True
        btnHide.Enabled = True



        btnSpellChecker.Enabled = True


        btnGenEx.Enabled = True

        btnSaveEdit.Enabled = True

        btnSaveDispute.Enabled = True

        btnEmail.Enabled = True

        btnDispute.Enabled = True



    End Sub

    Public Sub buttondisables()

        btnApproval.Enabled = False

        btnDispute.Enabled = False

        btnSaveScoreCard.Enabled = False
        btnHide.Enabled = False

        btnSpellChecker.Enabled = False

        btnGenEx.Enabled = False

        btnSaveEdit.Enabled = False


        btnSaveDispute.Enabled = False

        btnEmail.Enabled = False



    End Sub





    Private Sub QAChatRev_Load(sender As Object, e As EventArgs) Handles MyBase.Load




        Try

            GetEMAILInfo()


            ReaderEAdress()

            btnSaveDispute.SendToBack()

            If txtPendingReview.Text = "Pending Review" Then

                btnSaveScoreCard.Visible = True
                btnDispute.Visible = True
                grpDispute.Visible = False
                '   txtRevComments.BackColor = Color.LightGray
                btnDisabled.Visible = False


            ElseIf txtPendingReview.Text = "Reviewed" Then

                btnDispute.Visible = False
                btnDisabled.Visible = True
                btnSaveScoreCard.Visible = False
                btnApproval.Visible = False
                grpDispute.Enabled = False
                txtRevComments.ReadOnly = True

                lblDispute.Visible = True
                lblDis.Visible = True
                txtDisApp.Visible = True
                txtDisputeScore.Visible = True
                txtDisputedTCXScore.Visible = True


            End If




            If txtPendingReview.Text = "1" And lblDeciderX2.Text = "1" Then

                btnApproval.Visible = True
                btnDispute.Visible = False
                grpDispute.Visible = True
                grpDispute.Enabled = True
                txtDisApp.ReadOnly = True
                txtDisApp.Visible = True


                lblDis.Visible = True
                lblDispute.Visible = True
                radDisNo.Visible = True
                radDisYes.Visible = True
                txtDisputeScore.Visible = True
                txtDisputedTCXScore.Visible = True

            ElseIf txtPendingReview.Text = "1" And lblDeciderX2.Text = "2" Then


                btnDisabled.Visible = True
                btnDisabled.Text = "Awaiting approval"
                grpDispute.Visible = True
                grpDispute.Enabled = False

                lblDis.Visible = True
                lblDispute.Visible = True
                txtDisputeScore.Visible = True
                txtDisputedTCXScore.Visible = True
                txtDisApp.Visible = True
            End If




            If Form2.lblESDecider.Text = "1" Then

                '   editscorecardCheckBOX.Visible = True

                btnDelScorecard.Visible = True


            End If



            Me.WindowState = FormWindowState.Maximized


            SpellChecker1.SpellCheckMode = DevExpress.XtraSpellChecker.SpellCheckMode.AsYouType
            SpellChecker1.ParentContainer = Me
            SpellChecker1.CheckAsYouTypeOptions.CheckControlsInParentContainer = True
            SpellChecker1.SpellCheckMode = SpellCheckMode.AsYouType


            dic_en_US.DictionaryPath = en_USdicPath
            dic_en_US.GrammarPath = en_USaffPath
            dic_en_US.Culture = New CultureInfo("en-US")
            SpellChecker1.Dictionaries.Add(dic_en_US)



            'dic_en_US.DictionaryPath = "\\NOAMIND01FIL05\Premier_Support\Qa Application\Dictionary\en_US.dic"
            'dic_en_US.GrammarPath = "\\NOAMIND01FIL05\Premier_Support\Qa Application\Dictionary\en_US.aff"
            'dic_en_US.Culture = New CultureInfo("en-US")
            'SpellChecker1.Dictionaries.Add(dic_en_US)



            ' Me.ActiveControl = txtRevComments

            Me.ActiveControl = txtSR


            ''Date
            Time.Enabled = True


            Control.CheckForIllegalCrossThreadCalls = False


            'Me.EnableControlExtensions()


            dtpCondate.Format = DateTimePickerFormat.Custom
            dtpCondate.CustomFormat = "MM/dd/yyyy"




            dtpReviewdate.Format = DateTimePickerFormat.Custom
            dtpReviewdate.CustomFormat = "MM/dd/yyyy"








            If cboAF.Text <> "N/a" Then

                cboAutoFail.Checked = True


            End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try







    End Sub


    Public Sub ReaderEAdress()

        Try

            QaSetupMod.connecttemp19()

            sqltemp19 = "SELECT UserEmail FROM [Login] WHERE UserName='" & txtOrignalAuditor.Text & "'"

            '   sqltemp19 = "SELECT * FROM [Agents] WHERE Supervisor='" & txtOrignalAuditor.Text & "'"

            Dim cmdtemp As New SqlClient.SqlCommand




            cmdtemp.CommandText = sqltemp19

            cmdtemp.Connection = contemp19



            readertemp19 = cmdtemp.ExecuteReader


            While (readertemp19.Read())


                '    

                txtOrgAudEmail.Text = readertemp19("UserEmail")


            End While



            cmdtemp.Dispose()

            contemp19.Close()

            Me.Cursor = Cursors.Hand




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try





    End Sub
    Private Sub Time_Tick(sender As Object, e As EventArgs) Handles Time.Tick



        dtpReviewdate.Text = Date.Now.ToString("MM/dd/yyyy")



    End Sub

    Public Sub del()


        Try



            Dim con9 As SqlConnection
            Dim com9 As SqlCommand

            ' Test

            ' con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")

            '  con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb")





            'P Drive 

            con9 = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")

            'P nu Drive 

            '    con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")




            '' Dyanic



            '   con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")

            com9 = New SqlCommand("delete from [QAMainDB] where [ID] =@ID", con9)


            con9.Open()

            com9.Parameters.AddWithValue("@ID", lblGhostID.Text)

            com9.ExecuteNonQuery()

            con9.Close()







        Catch ex As Exception



            MsgBox(ex.Message)



        End Try

    End Sub



    Public Sub del1()


        Try



            Dim con9 As SqlConnection
            Dim com9 As SqlCommand

            ' Test

            '  con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")

            '  con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb")





            'P Drive 

            con9 = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")

            'P nu Drive 

            '   con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")




            '' Dyanic


            '   con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")



            com9 = New SqlCommand("delete from [QAMainDB] where [ID] =@ID", con9)


            con9.Open()

            com9.Parameters.AddWithValue("@ID", lblOLDID.Text)

            com9.ExecuteNonQuery()

            con9.Close()




        Catch ex As Exception

            DelSCTimer.Enabled = False

            MsgBox(ex.Message)

            DelSCTimer.Enabled = False

        End Try



    End Sub

    Public Sub Reset()


        editscorecardCheckBOX.Checked = False

        cboAutoFail.Checked = False

        cboAF.Text = "N/a"


        cboAutoFail.Checked = False



        ProgressBar1.Value = 0


        lblprogr.Text = 0



    End Sub

    Public Sub MissedWeightsCalc()






        If cbo1_1.Text = 0 Then

            txt1_1b.Text = "0"

        Else

            txt1_1b.Text = "1"

        End If


        If cbo1_2.Text = 0 Then

            txt1_2b.Text = "0"

        Else

            txt1_2b.Text = "1"

        End If





        If cbo2_1.Text = 0 Then

            txt2_1b.Text = "0"

        Else
            txt2_1b.Text = "1"

        End If







        If cbo3_1.Text = 0 Then

            txt3_1b.Text = "0"

        Else

            txt3_1b.Text = "1"
        End If


        If cbo3_2.Text = 0 Then

            txt3_2b.Text = "0"

        Else
            txt3_2b.Text = "1"


        End If



        If cbo3_3.Text = 0 Then

            txt3_3b.Text = "0"

        Else
            txt3_3b.Text = "1"


        End If



        If cbo3_4.Text = 0 Then

            txt3_4b.Text = "0"
        Else

            txt3_4b.Text = "1"
        End If



        If cbo3_5.Text = 0 Then

            txt3_5b.Text = "0"

        Else

            txt3_5b.Text = "1"

        End If



        If cbo3_6.Text = 0 Then

            txt3_6b.Text = "0"

        Else

            txt3_6b.Text = "1"

        End If



        If cbo3_7.Text = 0 Then

            txt3_7b.Text = "0"

        Else


            txt3_7b.Text = "1"

        End If



        If cbo3_8.Text = 0 Then

            txt3_8b.Text = "0"

        Else

            txt3_8b.Text = "1"

        End If




        If Cbo4_1.Text = 0 Then

            txt4_1b.Text = "0"
        Else

            txt4_1b.Text = "1"

        End If

        If cbo4_2.Text = 0 Then

            txt4_2b.Text = "0"

        Else


            txt4_2b.Text = "1"
        End If



        If cbo4_3.Text = 0 Then

            txt4_3b.Text = "0"

        Else

            txt4_3b.Text = "1"

        End If




        If cbo5_1.Text = 0 Then

            txt5_1b.Text = "0"

        Else


            txt5_1b.Text = "1"

        End If


        If cbo5_2.Text = 0 Then


            txt5_2b.Text = "0"


        Else


            txt5_2b.Text = "1"

        End If





        If cbo6_1.Text = 0 Then

            txt6_1b.Text = "0"

        Else


            txt6_1b.Text = "1"

        End If



        If cbo6_2.Text = 0 Then

            txt6_2b.Text = "0"

        Else

            txt6_2b.Text = "1"

        End If





        If cbo6_3.Text = 0 Then

            txt6_3b.Text = "0"

        Else

            txt6_3b.Text = "1"

        End If






        If cbo7_1.Text = 0 Then

            txt7_1b.Text = "0"

        Else

            txt7_1b.Text = "1"

        End If



        If cbo7_2.Text = 0 Then

            txt7_2b.Text = "0"

        Else

            txt7_2b.Text = "1"

        End If



        If cbo7_3.Text = 0 Then

            txt7_3b.Text = "0"
        Else

            txt7_3b.Text = "1"

        End If



        If cbo7_4.Text = 0 Then

            txt7_4b.Text = "0"

        Else

            txt7_4b.Text = "1"

        End If



        If cbo7_5.Text = 0 Then

            txt7_5b.Text = "0"

        Else

            txt7_5b.Text = "1"

        End If




        If cbo7_6.Text = 0 Then

            txt7_6b.Text = "0"

        Else

            txt7_6b.Text = "1"

        End If








    End Sub
    Public Sub Fillcombo()


        Try





            '  sqltemp1 = "Select * FROM [Agents] WHERE Supervisor='" & lblQAauditor.Text & "' "


            sqltemp13 = "SELECT * FROM [Teams]"


            Dim cmdtemp As New SqlCommand



            '  cmdtemp.CommandText = sqltemp

            cmdtemp.CommandText = sqltemp13

            cmdtemp.Connection = contemp13





            readertemp13 = cmdtemp.ExecuteReader



            While (readertemp13.Read())



                cboTeamName.Items.Add(readertemp13("Team"))




            End While






            cmdtemp.Dispose()

            readertemp13.Close()

            contemp13.Close()



        Catch ex As Exception



            MsgBox(ex.Message)


        End Try

    End Sub

    Public Sub save()



        Try


            ''saves review


            ' Test

            ' con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")


            '' Dyanic

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")



            'P Drive 
            con = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")



            'P nu Drive 

            'con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")



            con.Open()



            '   Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Dis_Approval],[Supervisor],[TCX_Score],[EditedQA])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Dis_Approval, @Supervisor, @TCX_Score, @EditedQA)"

            Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Supervisor],[TCX_Score],[Week_Number],[EditedQA],[1_1],[1_2],[2_1],[3_1],[3_2],[3_3],[3_4],[3_5],[3_6],[3_7],[3_8],[4_1],[4_2],[4_3],[5_1],[5_2],[6_1],[6_2],[6_3],[7_1],[7_2],[7_3],[7_4],[7_5],[7_6],[Month],[DisputedQA],[PendingDisputeID],[SRType],[MainSupervisor],[CSATScore],[CSATQ1],[CSATQ2],[CSATQ3],[CSATQ4],[CSATQ5],[CSATQ6],[Dis_TCXScore])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Supervisor, @TCX_Score, @Week_Number, @EditedQA,@1_1,@1_2,@2_1,@3_1,@3_2,@3_3,@3_4,@3_5,@3_6,@3_7,@3_8,@4_1,@4_2,@4_3,@5_1,@5_2,@6_1,@6_2,@6_3,@7_1,@7_2,@7_3,@7_4,@7_5,@7_6,@Month,@DisputedQA,@PendingDisputeID,@SRType,@MainSupervisor,@CSATScore,@CSATQ1,@CSATQ2,@CSATQ3,@CSATQ4,@CSATQ5,@CSATQ6,@Dis_TCXScore)"


            Using cmd As New SqlCommand(SQL, con)


                If txtSR.Text = "" Then

                    cmd.Parameters.AddWithValue("@SR", DBNull.Value)

                Else
                    cmd.Parameters.AddWithValue("@SR", txtSR.Text)

                End If





                cmd.Parameters.AddWithValue("@ContactID", txtContactID.Text)
                cmd.Parameters.AddWithValue("@CType", "Chat")
                cmd.Parameters.AddWithValue("@QA_Agent", txtAgentName.Text)
                cmd.Parameters.AddWithValue("@QA_Team", txtTeamName.Text)
                cmd.Parameters.AddWithValue("@QA_ContactDate", dtpCondate.Value)
                cmd.Parameters.AddWithValue("@QA_OrderID", txtOrderID.Text)
                cmd.Parameters.AddWithValue("@QA_Date", txtQADate.Text)
                cmd.Parameters.AddWithValue("@QA_Comments", txtQACom.Text)
                cmd.Parameters.AddWithValue("@QA_Opp", txtQAAOO.Text)
                cmd.Parameters.AddWithValue("@CI_Name", txtContactName.Text)
                cmd.Parameters.AddWithValue("@CI_Account", txtAccountNum.Text)
                cmd.Parameters.AddWithValue("@CI_Company", txtCompany.Text)
                cmd.Parameters.AddWithValue("@CI_Phone", txtContactPhone.Text)
                cmd.Parameters.AddWithValue("@CI_Email", txtContactEmail.Text)

                cmd.Parameters.AddWithValue("@Dis_Score", txtDisputeScore.Text)
                cmd.Parameters.AddWithValue("@Dis_Name", txtAgentName.Text)
                cmd.Parameters.AddWithValue("@Dis_Notes", txtDisputeNotes.Text)
                cmd.Parameters.AddWithValue("@Dis_AppComments", txtDiscomment.Text)
                cmd.Parameters.AddWithValue("@Rev_Date", dtpReviewdate.Value)
                cmd.Parameters.AddWithValue("@Rev_Manager", lblcurrentUser.Text)
                cmd.Parameters.AddWithValue("@Rev_Comments", txtRevComments.Text)



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
                cmd.Parameters.AddWithValue("@QAScore", txtQAScore.Text)
                cmd.Parameters.AddWithValue("@Autofail", cboAF.Text)
                cmd.Parameters.AddWithValue("@Dis_Approval", txtDisApp.Text)
                cmd.Parameters.AddWithValue("@Auditor", txtOrignalAuditor.Text)
                cmd.Parameters.AddWithValue("@Supervisor", txtSupervisor.Text)
                cmd.Parameters.AddWithValue("@TCX_Score", txtTCXScore.Text)
                cmd.Parameters.AddWithValue("@EditedQA", txtEditedQA.Text)
                cmd.Parameters.AddWithValue("@DisputedQA", txtDisputedQA.Text)

                cmd.Parameters.AddWithValue("@Week_Number", txtWeekNumber.Text)

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
                cmd.Parameters.AddWithValue("@Month", txtMonth.Text)
                cmd.Parameters.AddWithValue("@PendingDisputeID", "Reviewed")
                cmd.Parameters.AddWithValue("@SRType", txtSRType.Text)
                cmd.Parameters.AddWithValue("@MainSupervisor", txtSupervisor.Text)



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



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





    End Sub

    Public Sub SendEmaila()

        Try




            '  Dim attachment As Attachment = New Attachment("C:\Users\playe\Desktop\QA2\" & "" & txtSR.Text & " " & txtAgentName.Text & "-" & "Call QA Scorecard.xlsx")


            '   Dim attachment As Attachment = New Attachment(Desk & "\QA2\" & "" & txtSR.Text & " " & txtAgentName.Text & "-" & "Call QA Scorecard.xlsx")


            Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtContactID.Text & " " & txtAgentName.Text & "-" & "Chat QA Scorecard.xlsx")



            Dim mail As New MailMessage




            Dim smtp As New SmtpClient("NOAMIND01MXC12.NOAM.FADV.NET")


            mail.Attachments.Add(attachment)

            mail.Subject = "QA Scorecard for ContactID#:" + txtContactID.Text

            mail.To.Add(txtAgentEmail.Text)
            mail.CC.Add("CustomerCareQA@fadv.com")
            mail.CC.Add(lblUserEmail.Text)
            mail.CC.Add(lblSupervisorEmail.Text)





            mail.From = New MailAddress("CustomerCareQA@fadv.com")


            mail.Body = "Hello " + txtAgentName.Text + "," & vbCrLf &
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

            '     SplashScreenManager1.CloseWaitForm()

            EmailBackground.CancelAsync()


            MsgBox(ex.Message)

            SenderEmail1.Enabled = False


        End Try




    End Sub


    Public Sub SendEmail2Disputer()

        Try


            Dim mail As New MailMessage


            Dim smtp As New SmtpClient("NOAMIND01MXC12.NOAM.FADV.NET")

            If txtSR.Text = "" Then


                mail.Subject = "Contact ID: " + txtContactID.Text + " was disputed and requires approval in the QA App"

                mail.To.Add(txtAgentEmail.Text)
                mail.CC.Add("CustomerCareQA@fadv.com")
                mail.CC.Add(txtOrgAudEmail.Text)
                mail.CC.Add("Nick.DiVincenzo@Fadv.com")



                mail.From = New MailAddress("CustomerCareQA@fadv.com")


                mail.Body = "Hello," & vbCrLf &
               "" & vbCrLf &
             "Contact ID " + txtContactID.Text + " has been disputed and is pending approval by QA Team" & vbCrLf &
                "" & vbCrLf &
                "Thank you," & vbCrLf &
                "QA Team"

            Else

                mail.Subject = "SR#: " + txtSR.Text + " was disputed and requires approval in the QA App"

                mail.To.Add(txtAgentEmail.Text)
                mail.CC.Add("CustomerCareQA@fadv.com")
                mail.CC.Add(txtOrgAudEmail.Text)
                mail.CC.Add("Nick.DiVincenzo@Fadv.com")



                mail.From = New MailAddress("CustomerCareQA@fadv.com")


                mail.Body = "Hello," & vbCrLf &
               "" & vbCrLf &
                "SR " + txtSR.Text + " has been disputed and is pending approval by QA Team" & vbCrLf &
                "" & vbCrLf &
                "Thank you," & vbCrLf &
                "QA Team"





            End If





           smtp.EnableSsl = False


            smtp.Credentials = New System.Net.NetworkCredential("durraner", Form2.lblEmailPassword.Text)



            smtp.Port = 587

            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf Emailer)



            smtp.Send(mail)






        Catch ex As Exception


            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)



        End Try



    End Sub


    Public Sub SendEmail2DisputerGOC()

        Try


            Dim mail As New MailMessage


            Dim smtp As New SmtpClient("NOAMIND01MXC12.NOAM.FADV.NET")



            If txtSR.Text = "" Then


                mail.Subject = "Contact ID: " + txtContactID.Text + " was disputed and requires approval in the QA App"

                mail.To.Add(txtAgentEmail.Text)
                mail.CC.Add("CustomerCareQA@fadv.com")
                mail.CC.Add(txtOrgAudEmail.Text)
                mail.CC.Add("Anitha.thiagrajan@Fadv.com")



                mail.From = New MailAddress("CustomerCareQA@fadv.com")


                mail.Body = "Hello," & vbCrLf &
               "" & vbCrLf &
                "Contact ID " + txtContactID.Text + " has been disputed and is pending approval by QA Team" & vbCrLf &
                "" & vbCrLf &
                "Thank you," & vbCrLf &
                "QA Team"


            Else

                mail.Subject = "SR#: " + txtSR.Text + " was disputed and requires approval in the QA App"

                mail.To.Add(txtAgentEmail.Text)
                mail.CC.Add("CustomerCareQA@fadv.com")
                mail.CC.Add(txtOrgAudEmail.Text)
                mail.CC.Add("Anitha.thiagrajan@Fadv.com")



                mail.From = New MailAddress("CustomerCareQA@fadv.com")


                mail.Body = "Hello," & vbCrLf &
               "" & vbCrLf &
                "SR " + txtSR.Text + " has been disputed and is pending approval by QA Team" & vbCrLf &
                "" & vbCrLf &
                "Thank you," & vbCrLf &
                "QA Team"





            End If





           smtp.EnableSsl = False


            smtp.Credentials = New System.Net.NetworkCredential("durraner", Form2.lblEmailPassword.Text)



            smtp.Port = 587

            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf Emailer)



            smtp.Send(mail)






        Catch ex As Exception


            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)



        End Try



    End Sub





    Public Sub save2()



        Try


            ''saves dispute


            ' Test

            ' con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb")


            '' Dyanic

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")



            'P Drive 
            con = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")



            'P nu Drive 

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")



            con.Open()


            '  Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Dis_Approval],[Supervisor],[TCX_Score],[EditedQA])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Dis_Approval, @Supervisor, @TCX_Score, @EditedQA)"

            Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Supervisor],[TCX_Score],[Week_Number],[EditedQA],[1_1],[1_2],[2_1],[3_1],[3_2],[3_3],[3_4],[3_5],[3_6],[3_7],[3_8],[4_1],[4_2],[4_3],[5_1],[5_2],[6_1],[6_2],[6_3],[7_1],[7_2],[7_3],[7_4],[7_5],[7_6],[Month],[DisputedQA],[OLDID],[PendingDisputeID],[Dis_TCXScore],[Dis_AutoFail],[SRType],[MainSupervisor],[CSATScore],[CSATQ1],[CSATQ2],[CSATQ3],[CSATQ4],[CSATQ5],[CSATQ6])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Supervisor, @TCX_Score, @Week_Number, @EditedQA,@1_1,@1_2,@2_1,@3_1,@3_2,@3_3,@3_4,@3_5,@3_6,@3_7,@3_8,@4_1,@4_2,@4_3,@5_1,@5_2,@6_1,@6_2,@6_3,@7_1,@7_2,@7_3,@7_4,@7_5,@7_6,@Month,@DisputedQA,@OLDID,@PendingDisputeID,@Dis_TCXScore,@Dis_AutoFail,@SRType,@MainSupervisor,@CSATScore,@CSATQ1,@CSATQ2,@CSATQ3,@CSATQ4,@CSATQ5,@CSATQ6)"





            Using cmd As New SqlCommand(SQL, con)



                If txtSR.Text = "" Then

                    cmd.Parameters.AddWithValue("@SR", DBNull.Value)

                Else
                    cmd.Parameters.AddWithValue("@SR", txtSR.Text)

                End If


                cmd.Parameters.AddWithValue("@ContactID", txtContactID.Text)
                cmd.Parameters.AddWithValue("@CType", "Chat")
                cmd.Parameters.AddWithValue("@QA_Agent", txtAgentName.Text)
                cmd.Parameters.AddWithValue("@QA_Team", txtTeamName.Text)
                cmd.Parameters.AddWithValue("@QA_ContactDate", dtpCondate.Value)
                cmd.Parameters.AddWithValue("@QA_OrderID", txtOrderID.Text)
                cmd.Parameters.AddWithValue("@QA_Date", txtQADate.Text)
                cmd.Parameters.AddWithValue("@QA_Comments", txtQACom.Text)
                cmd.Parameters.AddWithValue("@QA_Opp", txtQAAOO.Text)
                cmd.Parameters.AddWithValue("@CI_Name", txtContactName.Text)
                cmd.Parameters.AddWithValue("@CI_Account", txtAccountNum.Text)
                cmd.Parameters.AddWithValue("@CI_Company", txtCompany.Text)
                cmd.Parameters.AddWithValue("@CI_Phone", txtContactPhone.Text)
                cmd.Parameters.AddWithValue("@CI_Email", txtContactEmail.Text)





                cmd.Parameters.AddWithValue("@Dis_Score", txtDisputeScore.Text)
                cmd.Parameters.AddWithValue("@Dis_TCXScore", txtDisputedTCXScore.Text)
                cmd.Parameters.AddWithValue("@Dis_Name", txtAgentName.Text)
                cmd.Parameters.AddWithValue("@Dis_Notes", txtDisputeNotes.Text)
                cmd.Parameters.AddWithValue("@Dis_AppComments", txtDiscomment.Text)
                cmd.Parameters.AddWithValue("@Rev_Date", "9/9/2021")
                cmd.Parameters.AddWithValue("@Rev_Manager", txtSupervisor.Text)
                cmd.Parameters.AddWithValue("@Rev_Comments", txtRevComments.Text)

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

                cmd.Parameters.AddWithValue("@QAScore", txtQAScore.Text)
                cmd.Parameters.AddWithValue("@Autofail", cboAF.Text)
                cmd.Parameters.AddWithValue("@Dis_AutoFail", txtghostAFreason.Text)

                cmd.Parameters.AddWithValue("@Auditor", txtOrignalAuditor.Text)
                cmd.Parameters.AddWithValue("@Dis_Approval", txtDisApp.Text)
                cmd.Parameters.AddWithValue("@Supervisor", txtSupervisor.Text)
                cmd.Parameters.AddWithValue("@TCX_Score", txtTCXScore.Text)
                cmd.Parameters.AddWithValue("@Week_Number", txtWeekNumber.Text)

                cmd.Parameters.AddWithValue("@EditedQA", "1")
                cmd.Parameters.AddWithValue("OLDID", lblGhostID.Text)



                '  cmd.Parameters.AddWithValue("SRType", )




                cmd.Parameters.AddWithValue("@DisputedQA", "1")
                    cmd.Parameters.AddWithValue("PendingDisputeID", "1")











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
                cmd.Parameters.AddWithValue("@Month", txtMonth.Text)
                cmd.Parameters.AddWithValue("@SRType", txtSRType.Text)
                cmd.Parameters.AddWithValue("@MainSupervisor", txtSupervisor.Text)


                cmd.Parameters.AddWithValue("@CSATScore", txtCSATScore.Text)

                cmd.Parameters.AddWithValue("@CSATQ1", cboCSAT1.Text)
                cmd.Parameters.AddWithValue("@CSATQ2", cboCSAT2.Text)
                cmd.Parameters.AddWithValue("@CSATQ3", cboCSAT3.Text)
                cmd.Parameters.AddWithValue("@CSATQ4", cboCSAT4.Text)
                cmd.Parameters.AddWithValue("@CSATQ5", cboCSAT5.Text)
                cmd.Parameters.AddWithValue("@CSATQ6", cboCSAT6.Text)
                cmd.Parameters.AddWithValue("@Dis_TCXScore", txtDisputedTCXScore.Text)




                cmd.ExecuteNonQuery()

                con.Close()



            End Using



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





    End Sub


    Public Sub save3()



        Try


            ''saves Yes dispute


            ' Test

            ' con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb")


            '' Dyanic

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")



            'P Drive 
            con = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")



            'P nu Drive 

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")



            con.Open()


            '  Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Dis_Approval],[Supervisor],[TCX_Score],[EditedQA])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Dis_Approval, @Supervisor, @TCX_Score, @EditedQA)"

            Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Dis_Approval],[Supervisor],[TCX_Score],[Week_Number],[EditedQA],[1_1],[1_2],[2_1],[3_1],[3_2],[3_3],[3_4],[3_5],[3_6],[3_7],[3_8],[4_1],[4_2],[4_3],[5_1],[5_2],[6_1],[6_2],[6_3],[7_1],[7_2],[7_3],[7_4],[7_5],[7_6],[Month],[DisputedQA],[OLDID],[PendingDisputeID],[Dis_TCXScore],[SRType],[MainSupervisor],[CSATScore],[CSATQ1],[CSATQ2],[CSATQ3],[CSATQ4],[CSATQ5],[CSATQ6])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Dis_Approval, @Supervisor, @TCX_Score, @Week_Number, @EditedQA,@1_1,@1_2,@2_1,@3_1,@3_2,@3_3,@3_4,@3_5,@3_6,@3_7,@3_8,@4_1,@4_2,@4_3,@5_1,@5_2,@6_1,@6_2,@6_3,@7_1,@7_2,@7_3,@7_4,@7_5,@7_6,@Month,@DisputedQA,@OLDID,@PendingDisputeID,@Dis_TCXScore,@SRType,@MainSupervisor,@CSATScore,@CSATQ1,@CSATQ2,@CSATQ3,@CSATQ4,@CSATQ5,@CSATQ6)"





            Using cmd As New SqlCommand(SQL, con)



                If txtSR.Text = "" Then

                    cmd.Parameters.AddWithValue("@SR", DBNull.Value)

                Else
                    cmd.Parameters.AddWithValue("@SR", txtSR.Text)

                End If


                cmd.Parameters.AddWithValue("@ContactID", txtContactID.Text)
                cmd.Parameters.AddWithValue("@CType", "Chat")
                cmd.Parameters.AddWithValue("@QA_Agent", txtAgentName.Text)
                cmd.Parameters.AddWithValue("@QA_Team", txtTeamName.Text)
                cmd.Parameters.AddWithValue("@QA_ContactDate", dtpCondate.Value)
                cmd.Parameters.AddWithValue("@QA_OrderID", txtOrderID.Text)
                cmd.Parameters.AddWithValue("@QA_Date", txtQADate.Text)
                cmd.Parameters.AddWithValue("@QA_Comments", txtQACom.Text)
                cmd.Parameters.AddWithValue("@QA_Opp", txtQAAOO.Text)
                cmd.Parameters.AddWithValue("@CI_Name", txtContactName.Text)
                cmd.Parameters.AddWithValue("@CI_Account", txtAccountNum.Text)
                cmd.Parameters.AddWithValue("@CI_Company", txtCompany.Text)
                cmd.Parameters.AddWithValue("@CI_Phone", txtContactPhone.Text)
                cmd.Parameters.AddWithValue("@CI_Email", txtContactEmail.Text)





                cmd.Parameters.AddWithValue("@Dis_Score", txtDisputeScore.Text)
                cmd.Parameters.AddWithValue("@Dis_TCXScore", txtDisputedTCXScore.Text)
                cmd.Parameters.AddWithValue("@Dis_Name", txtAgentName.Text)
                cmd.Parameters.AddWithValue("@Dis_Notes", txtDisputeNotes.Text)
                cmd.Parameters.AddWithValue("@Dis_AppComments", txtDiscomment.Text)
                cmd.Parameters.AddWithValue("@Rev_Date", dtpReviewdate.Value)
                cmd.Parameters.AddWithValue("@Rev_Manager", lblcurrentUser.Text)
                cmd.Parameters.AddWithValue("@Rev_Comments", txtRevComments.Text)

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

                cmd.Parameters.AddWithValue("@QAScore", txtQAScore.Text)
                cmd.Parameters.AddWithValue("@Autofail", cboAF.Text)
                cmd.Parameters.AddWithValue("@Auditor", txtOrignalAuditor.Text)
                cmd.Parameters.AddWithValue("@Dis_Approval", "Yes")
                cmd.Parameters.AddWithValue("@Supervisor", txtSupervisor.Text)
                cmd.Parameters.AddWithValue("@TCX_Score", txtTCXScore.Text)
                cmd.Parameters.AddWithValue("@Week_Number", txtWeekNumber.Text)

                cmd.Parameters.AddWithValue("@EditedQA", "1")
                cmd.Parameters.AddWithValue("OLDID", lblOLDID.Text)



                '  cmd.Parameters.AddWithValue("SRType", )



                cmd.Parameters.AddWithValue("@DisputedQA", "2")
                cmd.Parameters.AddWithValue("PendingDisputeID", "Reviewed")



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
                cmd.Parameters.AddWithValue("@Month", txtMonth.Text)
                cmd.Parameters.AddWithValue("@SRType", txtSRType.Text)
                cmd.Parameters.AddWithValue("@MainSupervisor", txtSupervisor.Text)

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



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





    End Sub


    Public Sub save4()



        Try


            ''Saves approve No


            ' Test

            ' con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb")


            '' Dyanic

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")



            'P Drive 
            con = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")



            'P nu Drive 

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")



            con.Open()


            '  Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Dis_Approval],[Supervisor],[TCX_Score],[EditedQA])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Dis_Approval, @Supervisor, @TCX_Score, @EditedQA)"

            Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Dis_Approval],[Supervisor],[TCX_Score],[Week_Number],[EditedQA],[1_1],[1_2],[2_1],[3_1],[3_2],[3_3],[3_4],[3_5],[3_6],[3_7],[3_8],[4_1],[4_2],[4_3],[5_1],[5_2],[6_1],[6_2],[6_3],[7_1],[7_2],[7_3],[7_4],[7_5],[7_6],[Month],[DisputedQA],[OLDID],[PendingDisputeID],[Dis_TCXScore],[SRType],[MainSupervisor],[CSATScore],[CSATQ1],[CSATQ2],[CSATQ3],[CSATQ4],[CSATQ5],[CSATQ6])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Dis_Approval, @Supervisor, @TCX_Score, @Week_Number, @EditedQA,@1_1,@1_2,@2_1,@3_1,@3_2,@3_3,@3_4,@3_5,@3_6,@3_7,@3_8,@4_1,@4_2,@4_3,@5_1,@5_2,@6_1,@6_2,@6_3,@7_1,@7_2,@7_3,@7_4,@7_5,@7_6,@Month,@DisputedQA,@OLDID,@PendingDisputeID,@Dis_TCXScore,@SRType,@MainSupervisor,@CSATScore,@CSATQ1,@CSATQ2,@CSATQ3,@CSATQ4,@CSATQ5,@CSATQ6)"





            Using cmd As New SqlCommand(SQL, con)



                If txtSR.Text = "" Then

                    cmd.Parameters.AddWithValue("@SR", DBNull.Value)

                Else
                    cmd.Parameters.AddWithValue("@SR", txtSR.Text)

                End If


                cmd.Parameters.AddWithValue("@ContactID", txtContactID.Text)
                cmd.Parameters.AddWithValue("@CType", "Chat")
                cmd.Parameters.AddWithValue("@QA_Agent", txtAgentName.Text)
                cmd.Parameters.AddWithValue("@QA_Team", txtTeamName.Text)
                cmd.Parameters.AddWithValue("@QA_ContactDate", dtpCondate.Value)
                cmd.Parameters.AddWithValue("@QA_OrderID", txtOrderID.Text)
                cmd.Parameters.AddWithValue("@QA_Date", txtQADate.Text)
                cmd.Parameters.AddWithValue("@QA_Comments", txtQACom.Text)
                cmd.Parameters.AddWithValue("@QA_Opp", txtQAAOO.Text)
                cmd.Parameters.AddWithValue("@CI_Name", txtContactName.Text)
                cmd.Parameters.AddWithValue("@CI_Account", txtAccountNum.Text)
                cmd.Parameters.AddWithValue("@CI_Company", txtCompany.Text)
                cmd.Parameters.AddWithValue("@CI_Phone", txtContactPhone.Text)
                cmd.Parameters.AddWithValue("@CI_Email", txtContactEmail.Text)





                cmd.Parameters.AddWithValue("@Dis_Score", txtDisputeScore.Text)
                cmd.Parameters.AddWithValue("@Dis_TCXScore", txtDisputedTCXScore.Text)
                cmd.Parameters.AddWithValue("@Dis_Name", txtAgentName.Text)
                cmd.Parameters.AddWithValue("@Dis_Notes", txtDisputeNotes.Text)
                cmd.Parameters.AddWithValue("@Dis_AppComments", txtDiscomment.Text)
                cmd.Parameters.AddWithValue("@Rev_Date", dtpReviewdate.Value)
                cmd.Parameters.AddWithValue("@Rev_Manager", lblcurrentUser.Text)
                cmd.Parameters.AddWithValue("@Rev_Comments", txtRevComments.Text)

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

                cmd.Parameters.AddWithValue("@QAScore", txtDisputeScore.Text)
                cmd.Parameters.AddWithValue("@Autofail", txtghostAFreason.Text)
                cmd.Parameters.AddWithValue("@Auditor", txtOrignalAuditor.Text)
                cmd.Parameters.AddWithValue("@Dis_Approval", "No")
                cmd.Parameters.AddWithValue("@Supervisor", txtSupervisor.Text)
                cmd.Parameters.AddWithValue("@TCX_Score", txtDisputedTCXScore.Text)
                cmd.Parameters.AddWithValue("@Week_Number", txtWeekNumber.Text)

                cmd.Parameters.AddWithValue("@EditedQA", "1")
                cmd.Parameters.AddWithValue("OLDID", lblOLDID.Text)



                '  cmd.Parameters.AddWithValue("SRType", )




                cmd.Parameters.AddWithValue("@DisputedQA", "3")
                cmd.Parameters.AddWithValue("PendingDisputeID", "Reviewed")





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
                cmd.Parameters.AddWithValue("@Month", txtMonth.Text)
                cmd.Parameters.AddWithValue("@SRType", txtSRType.Text)
                cmd.Parameters.AddWithValue("@MainSupervisor", txtSupervisor.Text)

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



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





    End Sub



    Public Sub saveEdit()



        Try





            ' Test

            ' con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb")


            '' Dyanic

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")



            'P Drive 
            con = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")



            'P nu Drive 

            '  con = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")



            con.Open()


            '  Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Dis_Approval],[Supervisor],[TCX_Score],[EditedQA])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Dis_Approval, @Supervisor, @TCX_Score, @EditedQA)"

            Dim SQL As String = "INSERT INTO [QAMainDB] ([SR], [ContactID], [CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_1Note],[One_2Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_8],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Three_8Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_2],[Six_3],[Six_1Note],[Six_2Note],[Six_3Note],[Seven_1],[Seven_2],[Seven_3],[Seven_4],[Seven_5],[Seven_6],[Seven_1Note],[Seven_2Note],[Seven_3Note],[Seven_4Note],[Seven_5Note],[Seven_6Note],[QAScore],[Autofail],[Auditor],[Supervisor],[TCX_Score],[Week_Number],[EditedQA],[1_1],[1_2],[2_1],[3_1],[3_2],[3_3],[3_4],[3_5],[3_6],[3_7],[3_8],[4_1],[4_2],[4_3],[5_1],[5_2],[6_1],[6_2],[6_3],[7_1],[7_2],[7_3],[7_4],[7_5],[7_6],[Month],[SRType])  Values (@SR, @ContactID, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_1Note, @One_2Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7, @Three_8, @Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Three_8Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_2, @Six_3, @Six_1Note, @Six_2Note, @Six_3Note, @Seven_1, @Seven_2, @Seven_3, @Seven_4, @Seven_5, @Seven_6, @Seven_1Note, @Seven_2Note, @Seven_3Note, @Seven_4Note, @Seven_5Note, @Seven_6Note, @QAScore, @Autofail, @Auditor, @Supervisor, @TCX_Score, @Week_Number, @EditedQA,@1_1,@1_2,@2_1,@3_1,@3_2,@3_3,@3_4,@3_5,@3_6,@3_7,@3_8,@4_1,@4_2,@4_3,@5_1,@5_2,@6_1,@6_2,@6_3,@7_1,@7_2,@7_3,@7_4,@7_5,@7_6,@Month)"





            Using cmd As New SqlCommand(SQL, con)



                If txtSR.Text = "" Then

                    cmd.Parameters.AddWithValue("@SR", DBNull.Value)

                Else
                    cmd.Parameters.AddWithValue("@SR", txtSR.Text)

                End If


                cmd.Parameters.AddWithValue("@ContactID", txtContactID.Text)
                cmd.Parameters.AddWithValue("@CType", "Chat")
                cmd.Parameters.AddWithValue("@QA_Agent", txtAgentName.Text)
                cmd.Parameters.AddWithValue("@QA_Team", txtTeamName.Text)
                cmd.Parameters.AddWithValue("@QA_ContactDate", dtpCondate.Value)
                cmd.Parameters.AddWithValue("@QA_OrderID", txtOrderID.Text)
                cmd.Parameters.AddWithValue("@QA_Date", txtQADate.Text)
                cmd.Parameters.AddWithValue("@QA_Comments", txtQACom.Text)
                cmd.Parameters.AddWithValue("@QA_Opp", txtQAAOO.Text)
                cmd.Parameters.AddWithValue("@CI_Name", txtContactName.Text)
                cmd.Parameters.AddWithValue("@CI_Account", txtAccountNum.Text)
                cmd.Parameters.AddWithValue("@CI_Company", txtCompany.Text)
                cmd.Parameters.AddWithValue("@CI_Phone", txtContactPhone.Text)
                cmd.Parameters.AddWithValue("@CI_Email", txtContactEmail.Text)





                cmd.Parameters.AddWithValue("@Dis_Score", txtDisputeScore.Text)
                cmd.Parameters.AddWithValue("@Dis_Name", txtAgentName.Text)
                cmd.Parameters.AddWithValue("@Dis_Notes", txtDisputeNotes.Text)
                cmd.Parameters.AddWithValue("@Dis_AppComments", txtDiscomment.Text)
                cmd.Parameters.AddWithValue("@Rev_Date", txtRevDate.Text)
                cmd.Parameters.AddWithValue("@Rev_Manager", txtSupervisor.Text)
                cmd.Parameters.AddWithValue("@Rev_Comments", txtRevComments.Text)

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

                cmd.Parameters.AddWithValue("@QAScore", txtQAScore.Text)
                cmd.Parameters.AddWithValue("@Autofail", cboAF.Text)
                cmd.Parameters.AddWithValue("@Auditor", txtOrignalAuditor.Text)
                cmd.Parameters.AddWithValue("@Dis_Approval", txtDisApp.Text)
                cmd.Parameters.AddWithValue("@Supervisor", txtSupervisor.Text)
                cmd.Parameters.AddWithValue("@TCX_Score", txtTCXScore.Text)
                cmd.Parameters.AddWithValue("@Week_Number", txtWeekNumber.Text)

                cmd.Parameters.AddWithValue("@EditedQA", "1")
                '    cmd.Parameters.AddWithValue("OLDID", lblGhostID.Text)



                '  cmd.Parameters.AddWithValue("SRType", )


                'If lblGhostDispute.Text = "Dispute Activated" Then

                '    cmd.Parameters.AddWithValue("@DisputedQA", "1")
                '    cmd.Parameters.AddWithValue("PendingDisputeID", "1")

                'Else



                'End If





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
                cmd.Parameters.AddWithValue("@Month", txtMonth.Text)
                cmd.Parameters.AddWithValue("@SRType", txtSRType.Text)


                cmd.ExecuteNonQuery()

                con.Close()



            End Using



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





    End Sub

    Public Sub QaTotalScore()


        If cboAutoFail.CheckState = CheckState.Checked Then


            txtQAScore.Text = "0"


        ElseIf cboAutoFail.CheckState = CheckState.Unchecked Then





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


            txtQAScore.Text = intQascoreTotal



        End If


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

        txtTCXScore.Text = Format(Val(intTCXscore.ToString()), "0")


    End Sub

    Public Sub QAExcell()




        Try



            Dim oExcel As Object = CreateObject("Excel.Application")





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






            '  oSheet.Range("C35").Value = txtQAScore.Text



            If cboAutoFail.Checked = True And cboAutoFail.Visible = True Then

                oSheet.Range("C36").Value = "0"

            Else

                oSheet.Range("C36").Value = txtQAScore.Text

            End If



            oSheet.Range("C37").Value = txtTCXScore.Text


            oSheet.Range("C38").Value = txtSR.Text
            oSheet.Range("C39").Value = txtContactID.Text
            oSheet.Range("C40").Value = "Chat"
            oSheet.Range("C41").Value = "" & txtAgentName.Text
            oSheet.Range("C42").Value = "" & txtTeamName.Text

            oSheet.Range("C43").Value = dtpCondate.Text


            oSheet.Range("C44").Value = txtOrderID.Text

            oSheet.Range("C45").Value = "" & txtContactName.Text
            oSheet.Range("C46").Value = "" & txtContactPhone.Text
            oSheet.Range("C47").Value = "" & txtContactEmail.Text
            oSheet.Range("C48").Value = "" & txtCompany.Text
            oSheet.Range("C49").Value = "" & txtAccountNum.Text
            oSheet.Range("C50").Value = "" & cboAF.Text
            oSheet.Range("C51").Value = "" & txtOrignalAuditor.Text
            oSheet.Range("C52").Value = "" & txtQADate.Text



            oSheet.Range("B54").Value = txtQACom.Text
            oSheet.Range("B58").Value = txtQAAOO.Text





            ''Review

            oSheet.Range("C66").Value = "" & dtpReviewdate.Text
            oSheet.Range("C67").Value = "" & txtSupervisor.Text
            oSheet.Range("B68").Value = "" & txtRevComments.Text


            ''Dispute

            oSheet.Range("C76").Value = "" & txtDisputeScore.Text
            oSheet.Range("C77").Value = "" & txtSupervisor.Text
            oSheet.Range("B78").Value = "" & txtDisputeNotes.Text


            oSheet.Range("C86").Value = "" & txtDisApp.Text
            oSheet.Range("B87").Value = "" & txtDiscomment.Text



            obook.SaveAs(Desk & "\QA2\" & "" & txtContactID.Text & " " & txtAgentName.Text & "-" & "Chat QA Scorecard.xlsx")


            oExcel.Quit()





            'If txtSR.Text = "" Then

            '    ''
            '    '  oBook.SaveAs("C:\Users\playe\Desktop\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Chat QA Scorecard.xlsx")

            '    '' Dynamic
            '    ' oBook.SaveAs(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Chat QA Scorecard.xlsx")

            '    ''Dyanmic
            '    oBook.SaveAs(Desk & "\QA2 \ " & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Call QA Scorecard.xlsx")



            '    oExcel.Quit()
            'Else

            '    '' Home 

            '    '  oBook.SaveAs("C:\Users\playe\Desktop\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Chat QA Scorecard.xlsx")

            '    '' Dynamic
            '    '  oBook.SaveAs(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Chat QA Scorecard.xlsx")



            '    '' Dynamic
            '    oBook.SaveAs(Desk & "\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Chat QA Scorecard.xlsx")




            '    oExcel.Quit()



            '    End If


        Catch ex As Exception

            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)

        End Try

    End Sub

    Public Sub QAExcell2()




        Try



            Dim oExcel As Object = CreateObject("Excel.Application")





            '' P Drive

            '  Dim oBook As Object = oExcel.Workbooks.Open("P:\QA Application\QA1\Chat.xlsx")


            '' Resouce
            Dim exeDir As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)
            Dim xlpath = IO.Path.Combine(exeDir.DirectoryName, "Chat.xlsx")
            Dim obook As Object = oExcel.Workbooks.Open(xlpath)



            Dim oSheet As Object = obook.Worksheets("Chat")  'or oBook.Worksheets("SheetName")








            '  oSheet.Range("C3").Value = "" & One


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






            '  oSheet.Range("C35").Value = txtQAScore.Text



            If cboAutoFail.Checked = True And cboAutoFail.Visible = True Then

                oSheet.Range("C36").Value = "0"

            Else

                oSheet.Range("C36").Value = txtQAScore.Text

            End If



            oSheet.Range("C37").Value = txtTCXScore.Text


            oSheet.Range("C38").Value = txtSR.Text
            oSheet.Range("C39").Value = txtContactID.Text
            oSheet.Range("C40").Value = "Chat"
            oSheet.Range("C41").Value = "" & txtAgentName.Text
            oSheet.Range("C42").Value = "" & txtTeamName.Text

            oSheet.Range("C43").Value = dtpCondate.Text


            oSheet.Range("C44").Value = txtOrderID.Text

            oSheet.Range("C45").Value = "" & txtContactName.Text
            oSheet.Range("C46").Value = "" & txtContactPhone.Text
            oSheet.Range("C47").Value = "" & txtContactEmail.Text
            oSheet.Range("C48").Value = "" & txtCompany.Text
            oSheet.Range("C49").Value = "" & txtAccountNum.Text
            oSheet.Range("C50").Value = "" & cboAF.Text
            oSheet.Range("C51").Value = "" & txtOrignalAuditor.Text
            oSheet.Range("C52").Value = "" & txtQADate.Text



            oSheet.Range("B54").Value = txtQACom.Text
            oSheet.Range("B58").Value = txtQAAOO.Text





            ''Review

            oSheet.Range("C66").Value = "" & dtpReviewdate.Text
            oSheet.Range("C67").Value = "" & txtSupervisor.Text
            oSheet.Range("B68").Value = "" & txtRevComments.Text


            ''Dispute

            oSheet.Range("C76").Value = "" & txtDisputeScore.Text
            oSheet.Range("C77").Value = "" & txtSupervisor.Text
            oSheet.Range("B78").Value = "" & txtDisputeNotes.Text


            oSheet.Range("C86").Value = "" & txtDisApp.Text
            oSheet.Range("B87").Value = "" & txtDiscomment.Text



            obook.SaveAs(Desk & "\QA2\" & "" & txtSR.Text & " " & txtAgentName.Text & "-" & "Chat QA Scorecard.xlsx")


            oExcel.Quit()


            'If txtSR.Text = "" Then

            '    ''
            '    '  oBook.SaveAs("C:\Users\playe\Desktop\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Chat QA Scorecard.xlsx")

            '    '' Dynamic
            '    ' oBook.SaveAs(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Chat QA Scorecard.xlsx")

            '    ''Dyanmic
            '    oBook.SaveAs(Desk & "\QA2 \ " & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Call QA Scorecard.xlsx")



            '    oExcel.Quit()
            'Else

            '    '' Home 

            '    '  oBook.SaveAs("C:\Users\playe\Desktop\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Chat QA Scorecard.xlsx")

            '    '' Dynamic
            '    '  oBook.SaveAs(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Chat QA Scorecard.xlsx")



            '    '' Dynamic
            '    oBook.SaveAs(Desk & "\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & cboTeamName.Text & "-" & "Chat QA Scorecard.xlsx")




            '    oExcel.Quit()



            '    End If


        Catch ex As Exception

            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)

        End Try

    End Sub



    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork



        Try

            ActiveControl = txtSR


            For i = 0 To 100

                System.Threading.Thread.Sleep(50)
                Me.BackgroundWorker1.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i
            Next

            ''
            save()


            ' Send to Excell
            '  QAExcell()



            'StoreCallThread = New System.Threading.Thread(AddressOf save)

            'StoreCallThread.Start()




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged


        ProgressBar1.Value = e.ProgressPercentage


    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted


        SplashScreenManager1.CloseWaitForm()


        buttonEnables()



        Me.Cursor = Cursors.Hand


        PleaseWait.Hide()



        MsgBox("This review was successfully saved, close scorecard and refresh your Audit List", MessageBoxButtons.OK)



        '  Form2.RevClear()


        Reset()




        Form2.Show()

        Me.Hide()


        dtpCondate.Value = Today


    End Sub



    Private Sub btnSaveScoreCard_Click(sender As Object, e As EventArgs) Handles btnSaveScoreCard.Click

        disable()


        Try





            If txtRevComments.Text = "" Then


                MsgBox("Please be advised you must fill out 'review comments section' before proceeding", MessageBoxButtons.RetryCancel)

                Me.ActiveControl = txtRevComments


                Me.Cursor = Cursors.Hand


            Else






                If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then



                Else



                    If BackgroundWorker1.IsBusy = False Then

                        SplashScreenManager1.ShowWaitForm()


                        Me.ActiveControl = txtSR

                        Me.Cursor = Cursors.WaitCursor

                        buttondisables()

                        BackgroundWorker1.RunWorkerAsync()

                        '   PleaseWait.ShowDialog()

                        QaTotalScore()




                        DelSCTimer.Enabled = True





                    End If

                End If


            End If



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub

    Private Sub btnHide_Click(sender As Object, e As EventArgs) Handles btnHide.Click





        '  Me.Hide()

        Me.Close()

        Form2.Show()



    End Sub

    Private Sub DelSCTimer_Tick(sender As Object, e As EventArgs) Handles DelSCTimer.Tick



        Try



            Dim con9 As SqlConnection
            Dim com9 As SqlCommand

            ' Test

            '   con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")

            '  con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")





            'P Drive 

            con9 = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")

            'P n Drive 

            '  con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")



            '' Dyanic


            '   con9 = New System.Data.sqlclient.sqlconnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lbldrive2.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA1\QA.accdb")




            com9 = New SqlCommand("delete from [QAMainDB] where [ID] =@ID", con9)


            con9.Open()

            com9.Parameters.AddWithValue("@ID", lblGhostID.Text)

            com9.ExecuteNonQuery()

            con9.Close()




            '   Form2.refreshDB()

            DelSCTimer.Enabled = False

        Catch ex As Exception

            DelSCTimer.Enabled = False

            MsgBox(ex.Message)

            DelSCTimer.Enabled = False

        End Try






    End Sub

    Private Sub btnSpellChecker_Click(sender As Object, e As EventArgs) Handles btnSpellChecker.Click
        Try

            SpellChecker1.CheckContainer(Me)



        Catch ex As Exception



            MsgBox(ex.Message)


        End Try


    End Sub


    Public Sub editall()


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


        ''Readonly

        txt1_1.ReadOnly = False
        txt1_2.ReadOnly = False




        txt2_1.ReadOnly = False


        txt3_1.ReadOnly = False
        txt3_2.ReadOnly = False
        txt3_3.ReadOnly = False
        txt3_4.ReadOnly = False
        txt3_5.ReadOnly = False
        txt3_6.ReadOnly = False
        txt3_7.ReadOnly = False
        txt3_8.ReadOnly = False



        txt4_1.ReadOnly = False
        txt4_2.ReadOnly = False
        txt4_3.ReadOnly = False

        txt5_1.ReadOnly = False
        txt5_2.ReadOnly = False




        txt6_1.ReadOnly = False
        txt6_2.ReadOnly = False
        txt6_3.ReadOnly = False


        txt7_1.ReadOnly = False
        txt7_2.ReadOnly = False
        txt7_3.ReadOnly = False
        txt7_4.ReadOnly = False
        txt7_5.ReadOnly = False
        txt7_6.ReadOnly = False



        txtSR.Enabled = True
        txtContactID.Enabled = True
        txtContactName.Enabled = True
        txtContactEmail.Enabled = True
        txtContactPhone.Enabled = True
        ' txtQADate.Enabled = True
        txtAccountNum.Enabled = True
        txtCompany.Enabled = True
        txtOrderID.Enabled = True
        dtpCondate.Enabled = True



        txtSR.ReadOnly = False
        txtContactID.ReadOnly = False
        txtContactName.ReadOnly = False
        txtContactEmail.ReadOnly = False
        txtContactPhone.ReadOnly = False
        ' txtQADate.Enabled = True
        txtAccountNum.ReadOnly = False
        txtCompany.ReadOnly = False
        txtOrderID.ReadOnly = False

        cboTeamName.Enabled = True
        cboAgentName.Enabled = True
        txtOrignalAuditor.Enabled = True
        txtSupervisor.Enabled = True

        cboAF.Enabled = True




    End Sub



    Public Sub DisputeEdit()




        If cbo1_1.Text = "0" Then

            cbo1_1.Enabled = True
            txt1_1.Enabled = True
            txt1_1.BackColor = Color.Yellow


        End If

        If cbo1_2.Text = "0" Then


            cbo1_2.Enabled = True
            txt1_2.Enabled = True
            txt1_2.BackColor = Color.Yellow


        End If



        If cbo2_1.Text = "0" Then

            cbo2_1.Enabled = True
            txt2_1.Enabled = True
            txt2_1.BackColor = Color.Yellow
        End If

        If cbo3_1.Text = "0" Then

            cbo3_1.Enabled = True
            txt3_1.Enabled = True
            txt3_1.BackColor = Color.Yellow
        End If


        If cbo3_2.Text = "0" Then

            cbo3_2.Enabled = True
            txt3_2.Enabled = True
            txt3_2.BackColor = Color.Yellow
        End If

        If cbo3_3.Text = "0" Then

            cbo3_3.Enabled = True
            txt3_3.Enabled = True
            txt3_3.BackColor = Color.Yellow
        End If

        If cbo3_4.Text = "0" Then

            cbo3_4.Enabled = True
            txt3_4.Enabled = True
            txt3_4.BackColor = Color.Yellow
        End If

        If cbo3_5.Text = "0" Then

            cbo3_5.Enabled = True
            txt3_5.Enabled = True
            txt3_5.BackColor = Color.Yellow
        End If

        If cbo3_6.Text = "0" Then

            cbo3_6.Enabled = True
            txt3_6.Enabled = True
            txt3_6.BackColor = Color.Yellow

        End If

        If cbo3_7.Text = "0" Then

            cbo3_7.Enabled = True
            txt3_7.Enabled = True
            txt3_7.BackColor = Color.Yellow

        End If

        If cbo3_8.Text = "0" Then

            cbo3_8.Enabled = True
            txt3_8.Enabled = True
            txt3_8.BackColor = Color.Yellow

        End If

        If Cbo4_1.Text = "0" Then

            Cbo4_1.Enabled = True
            txt4_1.Enabled = True
            txt4_1.BackColor = Color.Yellow


        End If

        If cbo4_2.Text = "0" Then

            cbo4_2.Enabled = True
            txt4_2.Enabled = True
            txt4_2.BackColor = Color.Yellow


        End If

        If cbo4_3.Text = "0" Then

            cbo4_3.Enabled = True
            txt4_3.Enabled = True
            txt4_3.BackColor = Color.Yellow

        End If

        If cbo5_1.Text = "0" Then

            cbo5_1.Enabled = True
            txt5_1.Enabled = True
            txt5_1.BackColor = Color.Yellow

        End If

        If cbo5_2.Text = "0" Then


            cbo5_2.Enabled = True
            txt5_2.Enabled = True
            txt5_2.BackColor = Color.Yellow
        End If


        If cbo6_1.Text = "0" Then

            cbo6_1.Enabled = True
            txt6_1.Enabled = True
            txt6_1.BackColor = Color.Yellow
        End If

        If cbo6_2.Text = "0" Then

            cbo6_2.Enabled = True
            txt6_2.Enabled = True
            txt6_2.BackColor = Color.Yellow
        End If

        If cbo6_3.Text = "0" Then

            cbo6_3.Enabled = True
            txt6_3.Enabled = True
            txt6_3.BackColor = Color.Yellow
        End If

        If cbo7_1.Text = "0" Then

            cbo7_1.Enabled = True
            txt7_1.Enabled = True
            txt7_1.BackColor = Color.Yellow
        End If


        If cbo7_2.Text = "0" Then

            cbo7_2.Enabled = True
            txt7_2.Enabled = True
            txt7_2.BackColor = Color.Yellow
        End If

        If cbo7_3.Text = "0" Then

            cbo7_3.Enabled = True
            txt7_3.Enabled = True
            txt7_3.BackColor = Color.Yellow
        End If

        If cbo7_4.Text = "0" Then

            cbo7_4.Enabled = True
            txt7_4.Enabled = True
            txt7_4.BackColor = Color.Yellow
        End If

        If cbo7_5.Text = "0" Then

            cbo7_5.Enabled = True
            txt7_5.Enabled = True
            txt7_5.BackColor = Color.Yellow
        End If

        If cbo7_6.Text = "0" Then

            cbo7_6.Enabled = True
            txt7_6.Enabled = True
            txt7_6.BackColor = Color.Yellow

        End If


        If cboAF.Text = Nothing Then


        Else

            cboAF.Enabled = True

        End If






    End Sub

    Public Sub RONLY()

        '' Read Only clear

        txt1_1.ReadOnly = False
        txt1_2.ReadOnly = False



        txt2_1.ReadOnly = False


        txt3_1.ReadOnly = False
        txt3_2.ReadOnly = False
        txt3_3.ReadOnly = False
        txt3_4.ReadOnly = False
        txt3_5.ReadOnly = False
        txt3_6.ReadOnly = False
        txt3_7.ReadOnly = False
        txt3_8.ReadOnly = False



        txt4_1.ReadOnly = False
        txt4_2.ReadOnly = False
        txt4_3.ReadOnly = False

        txt5_1.ReadOnly = False
        txt5_2.ReadOnly = False




        txt6_1.ReadOnly = False
        txt6_2.ReadOnly = False
        txt6_3.ReadOnly = False


        txt7_1.ReadOnly = False
        txt7_2.ReadOnly = False
        txt7_3.ReadOnly = False
        txt7_4.ReadOnly = False
        txt7_5.ReadOnly = False
        txt7_6.ReadOnly = False


    End Sub


    Public Sub disable()



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

        txt1_1.Enabled = False
        txt1_2.Enabled = False




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

        ''Read Only

        txt1_1.ReadOnly = True
        txt1_2.ReadOnly = True




        txt2_1.ReadOnly = True


        txt3_1.ReadOnly = True
        txt3_2.ReadOnly = True
        txt3_3.ReadOnly = True
        txt3_4.ReadOnly = True
        txt3_5.ReadOnly = True
        txt3_6.ReadOnly = True
        txt3_7.ReadOnly = True
        txt3_8.ReadOnly = True



        txt4_1.ReadOnly = True
        txt4_2.ReadOnly = True
        txt4_3.ReadOnly = True

        txt5_1.ReadOnly = True
        txt5_2.ReadOnly = True




        txt6_1.ReadOnly = True
        txt6_2.ReadOnly = True
        txt6_3.ReadOnly = True


        txt7_1.ReadOnly = True
        txt7_2.ReadOnly = True
        txt7_3.ReadOnly = True
        txt7_4.ReadOnly = True
        txt7_5.ReadOnly = True
        txt7_6.ReadOnly = True


        ' txtSR.Enabled = False
        txtContactID.Enabled = False
        txtContactName.Enabled = False
        txtContactEmail.Enabled = False
        txtContactPhone.Enabled = False
        ' txtQADate.Enabled = True
        txtAccountNum.Enabled = False
        txtCompany.Enabled = False
        txtOrderID.Enabled = False
        dtpCondate.Enabled = False


        cboTeamName.Enabled = False
        cboTeamName.Enabled = False

        txtOrignalAuditor.Enabled = False
        txtSupervisor.Enabled = False

        cboAF.Enabled = False
        cboAgentName.Enabled = False




    End Sub

    Private Sub btnEditAll_Click(sender As Object, e As EventArgs)




    End Sub



    Public Sub GetEMAILInfo()


        Try
            Using con01 = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")


                Dim SQL01 As String = "SELECT * FROM [Agents] WHERE AgentName= @AgentName"


                Using cmd01 As New SqlCommand(SQL01, con01)



                    cmd01.Parameters.AddWithValue("@AgentName", txtAgentName.Text)




                    con01.Open()



                    Dim reader01 As SqlDataReader

                    reader01 = cmd01.ExecuteReader()




                    While reader01.Read()

                        lblSupervisorEmail.Text = (reader01("SuperEmail"))
                        txtAgentEmail.Text = (reader01("AgentEmail"))


                    End While
                    reader01.Close()
                    con01.Close()

                End Using

            End Using

        Catch ex As Exception


            MsgBox(ex.Message)


        End Try


    End Sub

    Public Sub EmailID()




        Dim msg = "You are about to create an excel scorecard for this audit, do you want to proceed?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)





        If responce = MsgBoxResult.Yes Then

            SplashScreenManager1.ShowWaitForm()
            BackgroundWorker7.RunWorkerAsync()

            QAExcell()




        Else





        End If



    End Sub


    Public Sub SRID()





        Dim msg = "You are about to create an excel scorecard for this audit, do you want to proceed?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)





        If responce = MsgBoxResult.Yes Then

            SplashScreenManager1.ShowWaitForm()
            BackgroundWorker2.RunWorkerAsync()

            QAExcell2()




        Else






        End If


    End Sub




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnGenEx.Click


        Try

            GetEMAILInfo()



            If txtContactID.Text = "" And txtSR.Text <> "" Then

                SRID()


            ElseIf txtContactID.Text <> "" And txtSR.Text <> "" Then

                SRID()


            ElseIf txtContactID.Text <> "" And txtSR.Text = "" Then

                EmailID()

            End If



        Catch ex As Exception


            MsgBox(ex.Message)


        End Try


    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork



        For i = 0 To 100

            System.Threading.Thread.Sleep(55)
            Me.BackgroundWorker2.ReportProgress(i)

            lblprogr.Text = i.ToString

            i = i
        Next




    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted

        Try
            SplashScreenManager1.CloseWaitForm()

        Dim msg = "The excel scorecard was successfully saved to your QA2 folder; would you like to email the scorecard to the the agent?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)





        If responce = MsgBoxResult.Yes Then

            SplashScreenManager1.ShowWaitForm()

            ProgressBar1.Value = 0

            EmailBackground.RunWorkerAsync()


            SendEmail()


            ' SenderEmail1.Enabled = True





        Else

            Reset()

            buttonEnables()




        End If


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





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



    Private Sub cboTeamName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTeamName.SelectedIndexChanged



        Try

            Me.Cursor = Cursors.WaitCursor


            cboAgentName.Text = "Please wait, Loading.."

            txtSupervisor.Text = "Please wait, Loading.."

            '  resetcombo()
            cboAgentName.Items.Clear()


            BackgroundWorker4.RunWorkerAsync()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try










    End Sub

    Private Sub editscorecardCheckBOX_CheckedChanged(sender As Object, e As EventArgs) Handles editscorecardCheckBOX.CheckedChanged





        If editscorecardCheckBOX.CheckState = CheckState.Checked Then

            btnSaveEdit.Visible = True

            btnSaveScoreCard.Visible = False

            cboAutoFail.Visible = True





            cboAgentName.Visible = True
            cboTeamName.Visible = True


            txtAgentName.Visible = False
            txtTeamName.Visible = False





            editall()




        ElseIf editscorecardCheckBOX.CheckState = CheckState.Unchecked Then


            txtAgentName.Visible = True
            txtTeamName.Visible = True

            cboAgentName.Visible = False
            cboTeamName.Visible = False


            btnSaveEdit.Visible = False
            btnSaveScoreCard.Visible = True

            cboAutoFail.Visible = False

            disable()


        End If

    End Sub

    Private Sub cboAutoFail_CheckedChanged(sender As Object, e As EventArgs) Handles cboAutoFail.CheckedChanged



        If cboAutoFail.CheckState = CheckState.Checked Then


            txtQAScore.Text = "0"



        ElseIf cboAutoFail.CheckState = CheckState.Unchecked Then


            cboAF.Text = "N/a"



        End If





    End Sub

    Private Sub BackgroundWorker3_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork


        Fillcombo()


    End Sub

    Private Sub BackgroundWorker4_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork



        Try

            QaSetupMod.connecttemp1b()

            sqltemp1b = "Select * FROM [Agents] WHERE Platform='" & cboTeamName.Text & " ' "



            Dim cmdtemp1b As New SqlCommand




            cmdtemp1b.CommandText = sqltemp1b

            cmdtemp1b.Connection = contemp1b



            readertemp1b = cmdtemp1b.ExecuteReader


            While (readertemp1b.Read())




                cboAgentName.Items.Add(readertemp1b("AgentName"))



                txtAgentName.Text = readertemp1b(1).ToString

                txtSupervisor.Text = readertemp1b(2).ToString








            End While




            cmdtemp1b.Dispose()



            readertemp1b.Close()

            Me.Cursor = Cursors.Hand





        Catch ex As Exception



            MsgBox(ex.Message)


        End Try





    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted

        contemp1b.Close()


        cboAgentName.Text = "Agent Name"

        Me.Cursor = Cursors.Hand






    End Sub

    Private Sub cboAgentName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAgentName.SelectedIndexChanged



        Me.Cursor = Cursors.WaitCursor

        cboAgentName.Text = "Please wait, Loading.."

        BackgroundWorker5.RunWorkerAsync()






    End Sub

    Private Sub BackgroundWorker5_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork


        Try


            QaSetupMod.connecttemp12()

            sqltemp12 = "SELECT * FROM [Agents] WHERE AgentName='" & cboAgentName.Text & " ' "



            Dim cmdtemp As New SqlCommand





            cmdtemp.CommandText = sqltemp12

            cmdtemp.Connection = contemp12



            readertemp12 = cmdtemp.ExecuteReader



            If (readertemp12.Read() = True) Then






                txtSupervisor.Text = (readertemp12("Supervisor"))

                txtTeamName.Text = (readertemp12("Platform"))

                txtAgentName.Text = (readertemp12("AgentName"))

            End If




            readertemp12.Close()

            cmdtemp.Dispose()



        Catch ex As Exception



            MsgBox(ex.Message)


        End Try


    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted



        contemp12.Close()

        Me.Cursor = Cursors.Hand

    End Sub

    Private Sub btnSaveEdit_Click(sender As Object, e As EventArgs) Handles btnSaveEdit.Click



        MissedWeightsCalc()

        Me.ActiveControl = txtSR


        If MsgBox("Are you sure you want to save the edits to the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then



        Else

            del()

            If BackgroundWorker6.IsBusy = False Then

                '    DelTimer2.Enabled = True

                disable()


                Me.ActiveControl = txtSR

                Me.Cursor = Cursors.WaitCursor


                buttondisables()

                '  PleaseWait.ShowDialog()


                TCXscore()


                QaTotalScore()


                BackgroundWorker6.RunWorkerAsync()








            End If



        End If

    End Sub

    Private Sub BackgroundWorker6_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker6.DoWork


        Try
            ActiveControl = txtSR

            For i = 0 To 100

                System.Threading.Thread.Sleep(80)
                Me.BackgroundWorker6.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i
            Next


            saveEdit()



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub BackgroundWorker6_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted



        Reset()


        buttonEnables()

        MsgBox("the edit saved successfully", MessageBoxButtons.OK)

        Me.Cursor = Cursors.Hand

        editscorecardCheckBOX.Checked = False





    End Sub

    Private Sub BackgroundWorker6_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker6.ProgressChanged


        ProgressBar1.Value = e.ProgressPercentage


    End Sub

    Private Sub DelBackroundWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles DelBackroundWorker.DoWork


        del()

    End Sub

    Private Sub DelBackroundWorker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles DelBackroundWorker.RunWorkerCompleted


        Me.Cursor = Cursors.Hand


        MsgBox("This QA Audit has been deleted")


    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btnDelScorecard.Click


        Try

            If MsgBox("Are you sure you want to delete this QA audit?, it will permanently delete all instances of the audit.", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then



            Else

                Me.Cursor = Cursors.WaitCursor


                DelBackroundWorker.RunWorkerAsync()

            End If


        Catch ex As Exception



            MsgBox(ex.Message)



        End Try



    End Sub

    Private Sub SenderEmail1_Tick(sender As Object, e As EventArgs) Handles SenderEmail1.Tick





        SendEmail()



        SenderEmail1.Enabled = False


    End Sub

    Private Shared Function Emailer(ByVal sender As Object, ByVal cert As X509Certificate, ByVal chain As X509Chain, ByVal errors As SslPolicyErrors) As Boolean

        Return True

    End Function



    Public Sub SendEmail()

        Try



            ' Dim attachment As Attachment = New Attachment("C:\Users\durraner\Documents\QASpreadSheet.xlsx")


            '  Dim attachment As Attachment = New Attachment(Desk & "\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & "Call QA Scorecard.xlsx")


            Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtSR.Text & " " & txtAgentName.Text & "-" & "Chat QA Scorecard.xlsx")



            Dim mail As New MailMessage




            Dim smtp As New SmtpClient("NOAMIND01MXC12.NOAM.FADV.NET")


            mail.Attachments.Add(attachment)


            mail.Subject = "QA Scorecard for SR#:" + txtSR.Text

            mail.To.Add(txtAgentEmail.Text)
            mail.CC.Add("CustomerCareQA@fadv.com")
            mail.CC.Add(lblUserEmail.Text)
            mail.CC.Add(lblSupervisorEmail.Text)





            mail.From = New MailAddress("CustomerCareQA@fadv.com")

            mail.Body = "Hello " + txtAgentName.Text + "," & vbCrLf &
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
            '  SplashScreenManager1.CloseWaitForm()

            EmailBackground.CancelAsync()


            MsgBox(ex.Message)

            SenderEmail1.Enabled = False



        End Try




    End Sub

    Private Sub BackgroundWorker2_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged


        ProgressBar1.Value = e.ProgressPercentage


    End Sub

    Private Sub QAChatRev_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown




        '''Regular Save

        'If e.Control And e.KeyCode.ToString = "S" And btnSaveScoreCard.Visible = True Then

        '    MissedWeightsCalc()


        '    If txtRevComments.Text = "" Then


        '        MsgBox("Please be advised you must fill out 'review comments section' before proceeding", MessageBoxButtons.RetryCancel)

        '        Me.ActiveControl = txtRevComments


        '        Me.Cursor = Cursors.Hand


        '    Else




        '        If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then



        '        Else



        '            If BackgroundWorker1.IsBusy = False Then

        '                Me.ActiveControl = txtSR

        '                Me.Cursor = Cursors.WaitCursor

        '                buttondisables()

        '                BackgroundWorker1.RunWorkerAsync()

        '                '   PleaseWait.ShowDialog()

        '                QaTotalScore()

        '                TCXscore()


        '                DelSCTimer.Enabled = True





        '            End If

        '        End If


        '    End If



        'End If



        '''Dispute

        'If e.Control And e.KeyCode.ToString = "S" And btnSaveDispute.Visible = True Then

        '    MissedWeightsCalc()


        '    If MsgBox("Are you sure you want to save this dispute", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then



        '    Else

        '        If lblGhostDispute.Text = "Dispute Activated" And radDisYes.Checked = False And radDisNo.Checked = False Then


        '            MsgBox("You must approve this dispute before you save")

        '            Me.ActiveControl = radDisYes

        '        Else

        '            del()

        '            If BackgroundWorker5.IsBusy = False Then

        '                '    DelTimer2.Enabled = True

        '                disable()


        '                Me.ActiveControl = txtSR

        '                Me.Cursor = Cursors.WaitCursor


        '                buttondisables()

        '                '  PleaseWait.ShowDialog()


        '                TCXscore()

        '                QaTotalScore()


        '                BackgroundWorker6.RunWorkerAsync()








        '            End If



        '        End If


        '    End If






        'End If







        '''Edit

        'If e.Control And e.KeyCode.ToString = "S" And btnSaveEdit.Visible = True Then

        '    MissedWeightsCalc()


        '    If MsgBox("Are you sure you want to save the edits to the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then



        '    Else



        '        del()

        '        If BackgroundWorker5.IsBusy = False Then

        '            '    DelTimer2.Enabled = True

        '            disable()


        '            Me.ActiveControl = txtSR

        '            Me.Cursor = Cursors.WaitCursor


        '            buttondisables()

        '            '  PleaseWait.ShowDialog()


        '            TCXscore()

        '            QaTotalScore()


        '            BackgroundWorker6.RunWorkerAsync()








        '        End If



        '    End If


        'End If









































        If e.Control And e.KeyCode.ToString = "X" Then

            SpellChecker1.CheckContainer(Me)



        End If





    End Sub

    Private Sub SendEmailFin_Tick(sender As Object, e As EventArgs) Handles SendEmailFin.Tick


        SplashScreenManager1.CloseWaitForm()
        Me.Cursor = Cursors.Hand




        MsgBox("The scorecard was successfully emailed to the agent")

        SendEmailFin.Enabled = False






    End Sub




    Public Sub SendEmailSR()

        Dim msg = "Are you sure you want to email the scorecard to agent?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)





        If responce = MsgBoxResult.Yes Then

            SplashScreenManager1.ShowWaitForm()
            ProgressBar1.Value = 0

            EmailBackground.RunWorkerAsync()


            SendEmail()


        Else



        End If



    End Sub

    Public Sub SendEmailConID()



        Dim msg = "Are you sure you want to email the scorecard to agent?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)





        If responce = MsgBoxResult.Yes Then
            SplashScreenManager1.ShowWaitForm()

            ProgressBar1.Value = 0

            EmailBackground.RunWorkerAsync()


            SendEmaila()


        Else



        End If



    End Sub
    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles btnEmail.Click

        GetEMAILInfo()

        If txtContactID.Text = "" And txtSR.Text <> "" Then

            SendEmailSR()

        ElseIf txtContactID.Text <> "" And txtSR.Text <> "" Then

            SendEmailSR()

        ElseIf txtContactID.Text <> "" And txtSR.Text = "" Then

            SendEmailConID()

        End If




    End Sub


    Private Sub BackgroundWorker7_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker7.DoWork




        For i = 0 To 100

            System.Threading.Thread.Sleep(55)
            Me.BackgroundWorker2.ReportProgress(i)

            lblprogr.Text = i.ToString

            i = i
        Next


    End Sub

    Private Sub BackgroundWorker7_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker7.ProgressChanged

        ProgressBar1.Value = e.ProgressPercentage


    End Sub

    Private Sub BackgroundWorker7_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker7.RunWorkerCompleted

        Try


            SplashScreenManager1.CloseWaitForm()

        Dim msg = "The excel scorecard was successfully saved to your QA2 folder; would you like to email the scorecard to the the agent?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)





        If responce = MsgBoxResult.Yes Then

            SplashScreenManager1.ShowWaitForm()

            ProgressBar1.Value = 0

            EmailBackground.RunWorkerAsync()


            SendEmaila()




        Else

            Reset()

            buttonEnables()




        End If




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try






    End Sub

    Private Sub btnDispute_Click(sender As Object, e As EventArgs) Handles btnDispute.Click

        Dim msg1 = "Are you sure you want to dispute the QA score?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg1, style, title)





        If responce = MsgBoxResult.Yes Then

            '   btnSaveDispute.Location = New System.Drawing.Point(10, 679)

            btnSaveDispute.BringToFront()

            txtDisputeNotes.BackColor = Color.Yellow

            grpDispute.Enabled = True

            'Label82.Visible = False
            'txtDiscomment.Visible = False

            'lblDis.Visible = False
            'txtDisApp.Visible = False


            lblDispute.Visible = True
            txtDisputeScore.Visible = True
            txtDisputedTCXScore.Visible = True

            lblGhostDispute.Text = "Dispute Activated"

            MsgBox("Select the disputed weight leave a reason below")

            btnSaveEdit.Visible = False

            btnSaveScoreCard.Visible = False

            btnSaveDispute.Visible = True


            cboAutoFail.Visible = True



            grpDispute.Visible = True


            DisputeEdit()

            RONLY()



        Else





        End If






    End Sub

    Private Sub radDisYes_CheckedChanged(sender As Object, e As EventArgs) Handles radDisYes.CheckedChanged


        txtDisApp.Text = "Yes"

    End Sub

    Private Sub radDisNo_CheckedChanged(sender As Object, e As EventArgs) Handles radDisNo.CheckedChanged



        txtDisApp.Text = "No"


    End Sub

    Private Sub btnSaveDispute_Click(sender As Object, e As EventArgs) Handles btnSaveDispute.Click


        MissedWeightsCalc()


        If txtDisputeNotes.Text = "" Then


            MsgBox("You must comment on Dispute before proceeding")

            Me.ActiveControl = txtDisputeNotes
            txtDisputeNotes.BackColor = Color.Yellow

        Else


            If MsgBox("Are you sure you want to save and for approval?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then



            Else

                SplashScreenManager1.ShowWaitForm()

                del()

                If BackgroundWorker10.IsBusy = False Then



                    disable()


                    Me.ActiveControl = txtSR

                    Me.Cursor = Cursors.WaitCursor


                    buttondisables()



                    TCXscore()

                    QaTotalScore()



                    ActivateEmailTimer.Enabled = True

                    BackgroundWorker10.RunWorkerAsync()





                End If



            End If


        End If



    End Sub

    Private Sub cboAF_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAF.SelectedIndexChanged


        If cboAF.SelectedIndex = 4 Then

            cboAutoFail.Checked = False



        End If

        If cboAF.SelectedIndex <> 4 Then

            cboAutoFail.Checked = True

        End If






    End Sub

    Private Sub btnApproval_Click(sender As Object, e As EventArgs) Handles btnApproval.Click


        Try

            If MsgBox("Are you sure you want to approve this Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then



            Else

                If radDisYes.Checked = False And radDisNo.Checked = False Then


                    MsgBox("You must approve this dispute before you save")

                    Me.ActiveControl = radDisYes

                Else


                    If txtDiscomment.Text = "" Then


                        MsgBox("You must comment on Dispute approval before proceeding.")

                        Me.ActiveControl = txtDiscomment
                        txtDiscomment.BackColor = Color.Yellow

                    Else

                        If radDisYes.Checked = True Then

                            SplashScreenManager1.ShowWaitForm()
                            ActivateEmailTimer2.Enabled = True

                            '     del()
                            BackgroundWorker8.RunWorkerAsync()


                        ElseIf radDisNo.Checked = True Then

                            SplashScreenManager1.ShowWaitForm()
                            ActivateEmailTimer2.Enabled = True

                            ' del()
                            BackgroundWorker9.RunWorkerAsync()

                        End If






                    End If


                End If

            End If



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





    End Sub

    Private Sub BackgroundWorker8_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker8.DoWork


        Try
            ActiveControl = txtSR

            For i = 0 To 100

                System.Threading.Thread.Sleep(80)
                Me.BackgroundWorker6.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i
            Next


            save3()



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub BackgroundWorker8_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker8.ProgressChanged


        ProgressBar1.Value = e.ProgressPercentage

    End Sub

    Private Sub BackgroundWorker8_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker8.RunWorkerCompleted


        SplashScreenManager1.CloseWaitForm()
        del()

        editscorecardCheckBOX.Checked = False


        Reset()


        buttonEnables()

        MsgBox("the approval of YES was saved successfully - please close scorecard and refresh the Audit List", MessageBoxButtons.OK)

        Me.Cursor = Cursors.Hand






    End Sub

    Private Sub BackgroundWorker9_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker9.DoWork


        Try
            ActiveControl = txtSR

            For i = 0 To 100

                System.Threading.Thread.Sleep(80)
                Me.BackgroundWorker6.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i
            Next


            save4()



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub BackgroundWorker9_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker9.ProgressChanged

        ProgressBar1.Value = e.ProgressPercentage

    End Sub

    Private Sub BackgroundWorker9_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker9.RunWorkerCompleted


        SplashScreenManager1.CloseWaitForm()
        del()

        editscorecardCheckBOX.Checked = False


        Reset()


        buttonEnables()

        MsgBox("the approval of NO was saved successfully - please close scorecard and refresh the Audit List", MessageBoxButtons.OK)

        Me.Cursor = Cursors.Hand






    End Sub

    Private Sub BackgroundWorker10_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker10.DoWork

        Try
            ActiveControl = txtSR

            For i = 0 To 100

                System.Threading.Thread.Sleep(90)
                Me.BackgroundWorker6.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i
            Next


            save2()



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try

    End Sub

    Private Sub BackgroundWorker10_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker10.ProgressChanged


        ProgressBar1.Value = e.ProgressPercentage

    End Sub

    Private Sub BackgroundWorker10_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker10.RunWorkerCompleted

        SplashScreenManager1.CloseWaitForm()

        Reset()


        '  buttonEnables()

        btnHide.Enabled = True

        MsgBox("This dispute was successfully saved and sent for approval, close scorecard and refresh your Audit List", MessageBoxButtons.OK)

        Me.Cursor = Cursors.Hand

        editscorecardCheckBOX.Checked = False



    End Sub

    Public Sub ConfirmDisputeEmail()

        Try


            Dim mail As New MailMessage


            Dim smtp As New SmtpClient("NOAMIND01MXC12.NOAM.FADV.NET")

            If txtSR.Text = "" Then

                mail.Subject = "Contact ID: " + txtContactID.Text + " was approved by the QA Team"


                mail.CC.Add("CustomerCareQA@fadv.com")
                mail.CC.Add(txtOrgAudEmail.Text)
                mail.CC.Add("Nick.DiVincenzo@Fadv.com")



                mail.From = New MailAddress("CustomerCareQA@fadv.com")


                mail.Body = "Hello," & vbCrLf &
               "" & vbCrLf &
                "Contact ID " + txtContactID.Text + " has been approved, please check the approval status in the QA Application" & vbCrLf &
                "" & vbCrLf &
                "Thank you," & vbCrLf &
                "QA Team"


            Else

                mail.Subject = "SR#: " + txtSR.Text + " was approved by the QA Team"


                mail.CC.Add("CustomerCareQA@fadv.com")
                mail.CC.Add(txtOrgAudEmail.Text)
                mail.CC.Add("Nick.DiVincenzo@Fadv.com")



                mail.From = New MailAddress("CustomerCareQA@fadv.com")


                mail.Body = "Hello," & vbCrLf &
               "" & vbCrLf &
                "SR " + txtSR.Text + " has been approved, please check the approval status in the QA Application" & vbCrLf &
                "" & vbCrLf &
                "Thank you," & vbCrLf &
                "QA Team"



            End If

           smtp.EnableSsl = False


            smtp.Credentials = New System.Net.NetworkCredential("durraner", Form2.lblEmailPassword.Text)



            smtp.Port = 587

            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf Emailer)



            smtp.Send(mail)






        Catch ex As Exception


            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)



        End Try



    End Sub


    Public Sub ConfirmDisputeEmailGOC()

        Try


            Dim mail As New MailMessage


            Dim smtp As New SmtpClient("NOAMIND01MXC12.NOAM.FADV.NET")


            If txtSR.Text = "" Then

                mail.Subject = "Disputed QA Scorecard for ContactID: " + txtContactID.Text


                mail.CC.Add("CustomerCareQA@fadv.com")
                mail.CC.Add(txtOrgAudEmail.Text)
                mail.CC.Add("Anitha.thiagrajan@Fadv.com")



                mail.From = New MailAddress("CustomerCareQA@fadv.com")


                mail.Body = "Hello," & vbCrLf &
               "" & vbCrLf &
                "Contact ID " + txtContactID.Text + " has been approved, please check the approval status in the QA Application" & vbCrLf &
                "" & vbCrLf &
                "Thank you," & vbCrLf &
                "QA Team"


            Else

                mail.Subject = "Disputed QA Scorecard for SR#:" + txtSR.Text


                mail.CC.Add("CustomerCareQA@fadv.com")
                mail.CC.Add(txtOrgAudEmail.Text)
                mail.CC.Add("Anitha.thiagrajan@Fadv.com")



                mail.From = New MailAddress("CustomerCareQA@fadv.com")


                mail.Body = "Hello," & vbCrLf &
               "" & vbCrLf &
                "SR " + txtSR.Text + " has been approved, please check the approval status in the QA Application" & vbCrLf &
                "" & vbCrLf &
                "Thank you," & vbCrLf &
                "QA Team"



            End If




           smtp.EnableSsl = False


            smtp.Credentials = New System.Net.NetworkCredential("durraner", Form2.lblEmailPassword.Text)



            smtp.Port = 587

            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf Emailer)



            smtp.Send(mail)






        Catch ex As Exception


            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)



        End Try



    End Sub


    Private Sub ActivateEmailTimer_Tick(sender As Object, e As EventArgs) Handles ActivateEmailTimer.Tick


        If lblregion.Text = "GOC" Then

            SendEmail2DisputerGOC()

            ActivateEmailTimer.Enabled = False

        ElseIf lblregion.Text = "US" Then

            SendEmail2Disputer()

            ActivateEmailTimer.Enabled = False

        End If

        ActivateEmailTimer.Enabled = False


    End Sub

    Private Sub ActivateEmailTimer2_Tick(sender As Object, e As EventArgs) Handles ActivateEmailTimer2.Tick


        If lblregion.Text = "GOC" Then

            ConfirmDisputeEmailGOC()

            ActivateEmailTimer2.Enabled = False

        ElseIf lblregion.Text = "US" Then

            ConfirmDisputeEmail()

            ActivateEmailTimer2.Enabled = False
        End If


        ActivateEmailTimer2.Enabled = False


    End Sub


End Class