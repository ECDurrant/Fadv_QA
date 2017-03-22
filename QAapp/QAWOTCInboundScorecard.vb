

Imports System.Net.Mail

Imports Microsoft.Office.Interop


'Imports i00SpellCheck

Imports System.Windows

Imports DevExpress.XtraSpellChecker

Imports System.Globalization

Imports System.IO

Imports System.Data.OleDb

Imports System.Net

Imports System.Net.Security

Imports System.Security.Cryptography.X509Certificates


Imports System.Data.SqlClient



Imports System.Reflection
Imports System.Threading

Public Class QAWOTCInboundScorecard

    '' i changed the class name from QAScoreCard to QANuScorecard then changed the partial class name in the designer



    Dim SQL As String
    Dim con As New SqlClient.SqlConnection


    Dim One As Integer
    Dim two As Integer
    Dim three As Integer
    Dim Four As Integer
    Dim Five As Integer
    Dim Six As Integer
    Dim seven As Integer





    ''Store Call Thread
    Dim StoreCallThread As System.Threading.Thread

    'Store Call Thread
    Dim ToExcell As System.Threading.Thread

    Dim st As Date

    Dim intQascoreTotal As Integer

    Dim ComboName As String
    Dim ComboTeam As String


    Dim YesNoButton As MessageBoxResult



    Dim dic_en_US As SpellCheckerOpenOfficeDictionary = New SpellCheckerOpenOfficeDictionary



    Dim Desk = My.Computer.FileSystem.SpecialDirectories.Desktop

    Dim total As Integer = 100

    '  Dim totalQA As Integer = 0


    Dim AgentEmail As String


    Public Shared AgentEmailonCall As String

    Dim totalQA As Integer


    Dim combonum As Integer

    Dim int1_1 As Integer = 2

    Dim int1_2 As Integer = 1


    Dim int1_3 As Integer = 2


    Dim int2_1 As Integer = 15


    Dim int3_1 As Integer = 2


    Dim int3_2 As Integer = 1


    Dim int3_3 As Integer = 3



    Dim int3_4 As Integer = 4




    Dim int3_5 As Integer = 3


    Dim int3_6 As Integer = 3

    Dim int3_7 As Integer = 3

    Dim int3_8 As Integer = 1



    Dim int4_1 As Integer = 5



    Dim int4_2 As Integer = 5


    Dim int4_3 As Integer = 5


    Dim int5_1 As Integer = 7


    Dim int5_2 As Integer = 8



    Dim int6_1 As Integer = 5

    Dim int6_2 As Integer = 10


    Dim int7_1 As Integer = 5


    Dim int7_2 As Integer = 5

    Dim int7_3 As Integer = 5




    Dim QaTotal As Integer = 2 + 1 + 2 + 15 + 2 + 1 + 3 + 4 + 3 + 3 + 3 + 1 + 5 + 5 + 5 + 7 + 8 + 5 + 10 + 5 + 5 + 5


    Dim newTotal1 As Integer
    Dim newTotal2 As Integer
    Dim newTotal3 As Integer

    Dim newTotal4 As Integer
    Dim newTotal5 As Integer
    Dim newTotal6 As Integer
    Dim newTotal7 As Integer
    Dim newTotal8 As Integer
    Dim newTotal9 As Integer
    Dim newTotal10 As Integer
    Dim newTotal11 As Integer
    Dim newTotal12 As Integer
    Dim newTotal13 As Integer
    Dim newTotal14 As Integer
    Dim newTotal15 As Integer
    Dim newTotal16 As Integer
    Dim newTotal17 As Integer
    Dim newTotal18 As Integer
    Dim newTotal19 As Integer
    Dim newTotal20 As Integer
    Dim newTotal21 As Integer
    Dim newTotal22 As Integer
    Dim newTotal23 As Integer
    Dim newTotal24 As Integer
    Dim newTotal25 As Integer
    Dim newTotal26 As Integer


    Dim ProgramDateForamt As String = "MM/dd/yyyy"

    Dim ProgramDate As DateTime = Now

    Dim InterConDate As DateTime

    Dim spellcheckDIR As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)

    Dim en_USaffPath = IO.Path.Combine(spellcheckDIR.DirectoryName, "en_US.aff")
    Dim en_USdicPath = IO.Path.Combine(spellcheckDIR.DirectoryName, "en_US.dic")




    Public Sub buttonEnables()


        btnSaveScoreCard.Enabled = True
        btnQaSetup.Enabled = True

        btnSave2.Enabled = True

        cboAgentName.Enabled = True

        cboSupervisor.Enabled = True


        cboSupervisorbox.Enabled = True

        btnSpellChecker.Enabled = True

        cboContactTypeCall.Enabled = True

    End Sub

    Public Sub buttondisables()



        btnSave2.Enabled = False
        btnSaveScoreCard.Enabled = False
        btnQaSetup.Enabled = False
        btnSpellChecker.Enabled = False


        cboAgentName.Enabled = False

        cboSupervisor.Enabled = False

        cboSupervisorbox.Enabled = False


        cboContactTypeCall.Enabled = False

    End Sub

    Public Sub DictLoad()




        Dim dictionary As New SpellCheckerISpellDictionary()

        Dim affStream As Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("Dictionaries.en_US.aff")
        Dim dicStream As Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("Dictionaries.en_US.dic")
        Dim alphStream As Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("Dictionaries.EnglishAlphabet.txt")
        dictionary.Culture = New CultureInfo("en-US")
        dictionary.LoadFromStream(dicStream, affStream, alphStream)

        SpellChecker2.Dictionaries.Add(dictionary)


        SpellChecker2.ParentContainer = Me
        SpellChecker2.CheckAsYouTypeOptions.CheckControlsInParentContainer = True
        SpellChecker2.SpellCheckMode = SpellCheckMode.AsYouType





    End Sub


    Private Shared Function Emailer(ByVal sender As Object, ByVal cert As X509Certificate, ByVal chain As X509Chain, ByVal errors As SslPolicyErrors) As Boolean

        Return True

    End Function


    Public Sub SendEmail()

        Try



            ' Dim attachment As Attachment = New Attachment("C:\Users\durraner\Documents\QASpreadSheet.xlsx")


            '  Dim attachment As Attachment = New Attachment(Desk & "\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & "Call QA Scorecard.xlsx")


            Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & "Call QA Scorecard.xlsx")

            '   Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & Desk & "\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "Call QA Scorecard.xlsx")





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

        End Try




    End Sub


    Public Sub SendEmail2()

        Try



            ' Dim attachment As Attachment = New Attachment("C:\Users\durraner\Documents\QASpreadSheet.xlsx")


            '  Dim attachment As Attachment = New Attachment(Desk & "\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & "Call QA Scorecard.xlsx")


            Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "Call QA Scorecard.xlsx")

            '    Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & Desk & "\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "Call QA Scorecard.xlsx")


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


            '   Dim attachment As Attachment = New Attachment(Desk & "\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & "Call QA Scorecard.xlsx")


            Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & "Call QA Scorecard.xlsx")

            '   Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & Desk & "\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "Call QA Scorecard.xlsx")





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


            '  Dim attachment As Attachment = New Attachment(Desk & "\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & "Call QA Scorecard.xlsx")


            Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & "\Users\" & Form2.lblSCRN.Text & "\Desktop\QA2\" & "" & txtContactID.Text & " " & cboSupervisorbox.Text & "-" & "Call QA Scorecard.xlsx")

            '    Dim attachment As Attachment = New Attachment(Form2.lblMDrive.Text & Desk & "\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "Call QA Scorecard.xlsx")


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






    Private Sub CallScoreCard_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Try

            PW.Hide()


            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")

            FillAutoFail()

            FillCallerType()

            FillSRtype()

            If lblDecider.Text = "1" Then

                'txtRevcom.Visible = False
                'GroupBox11.Visible = False
                'Label40.Visible = False


                btnSaveScoreCard.Visible = True
                btnSave2.Visible = False




                BackgroundWorker3.RunWorkerAsync()

                '  Fillcombo()
                '   QaSetupMod.connecttemp3()


            ElseIf lblDecider.Text = "2" Then


                btnSaveScoreCard.Visible = False
                btnSave2.Visible = True
                cboContactTypeCall.Visible = False

                BackgroundWorker4.RunWorkerAsync()


                '    QaSetupMod.connecttemp9()
                '  Fillcombo33()

            End If



            '  DictLoad()


            ' SpellChecker2.SpellCheckMode = DevExpress.XtraSpellChecker.SpellCheckMode.AsYouType
            SpellChecker2.ParentContainer = Me
            SpellChecker2.CheckAsYouTypeOptions.CheckControlsInParentContainer = True
            SpellChecker2.SpellCheckMode = SpellCheckMode.AsYouType

            'SpellCheckLoadTimer.Enabled = True

            'dic_en_US.DictionaryPath = "\\NOAMIND01FIL05\Premier_Support\Qa Application\Dictionary\en_US.dic"
            'dic_en_US.GrammarPath = "\\NOAMIND01FIL05\Premier_Support\Qa Application\Dictionary\en_US.aff"
            'dic_en_US.Culture = New CultureInfo("en-US")
            'SpellChecker2.Dictionaries.Add(dic_en_US)


            dic_en_US.DictionaryPath = en_USdicPath
            dic_en_US.GrammarPath = en_USaffPath
            dic_en_US.Culture = New CultureInfo("en-US")
            SpellChecker2.Dictionaries.Add(dic_en_US)





            Me.WindowState = FormWindowState.Maximized

            Me.ActiveControl = cbo1_1




            ''Date
            Time.Enabled = True


            Control.CheckForIllegalCrossThreadCalls = False


            ''Me.EnableControlExtensions()


            dtpCondate.Format = DateTimePickerFormat.Custom
            dtpCondate.CustomFormat = "MM/dd/yyyy"




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub


    Public Sub combonumbers()







    End Sub


    Public Sub Curdate()

        Dim cd As String = txtQADate.Text


        Dim cdt = Date.ParseExact(cd, "dd/MM/yyyy", Nothing)





    End Sub


    Public Sub resetcombo()

        cboAgentName.Items.Clear()

        '   cboAgentName.Text = "Agent Name"



    End Sub



    Public Sub TCXscore()

        Dim intTCXscore As Integer
        Dim increase As Integer


        Dim int3_1 As Integer = cbo3_1.Text
        Dim int3_2 As Integer = cbo3_2.Text
        Dim int3_3 As Integer = cbo3_3.Text
        Dim int3_4 As Integer = cbo3_4.Text
        Dim int3_5 As Integer = cbo3_5.Text
        Dim int3_6 As Integer = cbo3_6.Text
        Dim int3_7 As Integer = cbo3_7.Text



        Dim int5_1 As Integer = cbo5_1.Text

        Dim int6_1 As Integer = cbo6_1.Text



        '' lblQaAvg.Text = Format(Val(result.ToString()), "0.00")

        increase = int3_1 + int3_2 + int3_3 + int3_4 + int3_5 + int3_6 + int3_7 + int5_1 + int6_1

        intTCXscore = increase / 42 * 100

        txtTCXScore.Text = Format(Val(intTCXscore.ToString()), "0")
        lblTCXscore.Text = Format(Val(intTCXscore.ToString()), "0")


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







    Public Sub Fillcombo()


        Try


            QaSetupMod.connecttemp3()



            '  sqltemp1 = "Select * FROM [Agents] WHERE Supervisor='" & lblQAauditor.Text & "' "


            sqltemp3 = "SELECT * FROM [Supervisor]"


            Dim cmdtemp As New SqlClient.SqlCommand



            '  cmdtemp.CommandText = sqltemp

            cmdtemp.CommandText = sqltemp3

            cmdtemp.Connection = contemp3





            readertemp3 = cmdtemp.ExecuteReader



            While (readertemp3.Read())



                cboSupervisor.Items.Add(readertemp3("FullName"))




            End While




            readertemp3.Close()

            cmdtemp.Dispose()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try




    End Sub

    Public Sub Fillcombo33()


        Try


            QaSetupMod.connecttemp9()


            sqltemp9 = "SELECT * FROM [Agents] WHERE Supervisor='" & lblQAauditor1.Text & "' "





            Dim cmdtemp As New SqlClient.SqlCommand



            '  cmdtemp.CommandText = sqltemp

            cmdtemp.CommandText = sqltemp9

            cmdtemp.Connection = contemp9





            readertemp9 = cmdtemp.ExecuteReader



            While (readertemp9.Read())



                cboSupervisorbox.Items.Add(readertemp9("AgentName"))




            End While




            readertemp9.Close()


            cmdtemp.Dispose()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try




    End Sub



    Public Sub MissedWeightsReset()

        txt1_1a.Text = 1
        txt1_2a.Text = 1
        txt1_3a.Text = 1

        txt2_1a.Text = 1

        txt3_1a.Text = 1
        txt3_2a.Text = 1
        txt3_3a.Text = 1
        txt3_4a.Text = 1
        txt3_5a.Text = 1
        txt3_6a.Text = 1
        txt3_7a.Text = 1


        txt4_1a.Text = 1
        txt4_2a.Text = 1
        txt4_3a.Text = 1

        txt5_1a.Text = 1
        txt5_2a.Text = 1

        txt6_1a.Text = 1


        txt7_1a.Text = 1
        txt7_2a.Text = 1
        txt7_3a.Text = 1



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



        If cbo1_3.Text = 0 Then

            txt1_3a.Text = "0"

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









    End Sub





    Public Sub Store()




        Try




            ' Test

            '  con = New System.Data.SqlClient.SqlConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb")



            'P Drive 

            con = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")



            con.Open()





            Dim SQL As String = "INSERT INTO [QAMainDB] ([ContactID],[EMID],[RegID],[AHT],[CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_3],[One_1Note],[One_2Note],[One_3Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_1Note],[Seven_1],[Seven_2],[Seven_3],[Seven_1Note],[Seven_2Note],[Seven_3Note],[QAScore],[Autofail],[Auditor],[Supervisor],[Week_Number],[EditedQA],[TCX_Score],[1_1],[1_2],[1_3],[2_1],[3_1],[3_2],[3_3],[3_4],[3_5],[3_6],[3_7],[4_1],[4_2],[4_3],[5_1],[5_2],[6_1],[7_1],[7_2],[7_3],[Month],[PendingDisputeID],[Dis_TCXScore],[SRType],[CSATScore],[CSATQ1],[CSATQ2],[CSATQ3],[CSATQ4],[CSATQ5],[CSATQ6],[CallerType])  Values (@ContactID,@EMID,@RegID,@AHT, @CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_3, @One_1Note, @One_2Note, @One_3Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7,@Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note, @Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_1Note, @Seven_1, @Seven_2, @Seven_3, @Seven_1Note, @Seven_2Note, @Seven_3Note, @QAScore, @Autofail, @Auditor, @Supervisor, @Week_Number, @EditedQA,@TCX_Score,@1_1,@1_2,@1_3,@2_1,@3_1,@3_2,@3_3,@3_4,@3_5,@3_6,@3_7,@4_1,@4_2,@4_3,@5_1,@5_2,@6_1,@7_1,@7_2,@7_3,@Month,@PendingDisputeID,@Dis_TCXScore,@SRType,@CSATScore,@CSATQ1,@CSATQ2,@CSATQ3,@CSATQ4,@CSATQ5,@CSATQ6,@CallerType)"



            Using cmd As New SqlCommand(SQL, con)




                cmd.Parameters.AddWithValue("@ContactID", txtContactID.Text)
                cmd.Parameters.AddWithValue("@EMID", txtEMID.Text)
                cmd.Parameters.AddWithValue("@RegID", txtRegID.Text)
                cmd.Parameters.AddWithValue("@AHT", txtAHT.Text)
                cmd.Parameters.AddWithValue("@CType", "WOTC Inbound")
                cmd.Parameters.AddWithValue("@QA_Agent", cboAgentName.Text)
                cmd.Parameters.AddWithValue("@QA_Team", txtTeamName.Text)
                cmd.Parameters.AddWithValue("@QA_ContactDate", dtpCondate.Value)
                cmd.Parameters.AddWithValue("@QA_Date", Now)
                cmd.Parameters.AddWithValue("@QA_OrderID", txtOrderID.Text)
                cmd.Parameters.AddWithValue("@QA_Comments", txtQACom.Text)
                cmd.Parameters.AddWithValue("@QA_Opp", txtQAAOO.Text)
                cmd.Parameters.AddWithValue("@CI_Name", txtContactName.Text)
                cmd.Parameters.AddWithValue("@CI_Account", txtAccountNum.Text)
                cmd.Parameters.AddWithValue("@CI_Company", txtCompany.Text)
                cmd.Parameters.AddWithValue("@CI_Phone", txtContactPhone.Text)
                cmd.Parameters.AddWithValue("@CI_Email", txtContactEmail.Text)

                If txtRevcom.Text = "" Then


                    cmd.Parameters.AddWithValue("@Rev_Date", "9/9/2020")
                    cmd.Parameters.AddWithValue("@Rev_Manager", cboSupervisor.Text)
                    cmd.Parameters.AddWithValue("@Rev_Comments", "")
                    cmd.Parameters.AddWithValue("@PendingDisputeID", "Pending Review")
                ElseIf txtRevcom.Text <> "" Then

                    cmd.Parameters.AddWithValue("@Rev_Date", txtQADate.Text)
                    cmd.Parameters.AddWithValue("@Rev_Manager", cboSupervisor.Text)
                    cmd.Parameters.AddWithValue("@Rev_Comments", txtRevcom.Text)
                    cmd.Parameters.AddWithValue("@PendingDisputeID", "Reviewed")


                End If


                '  cmd.Parameters.AddWithValue("@Dis_Score", lblQAScore1.Text)
                cmd.Parameters.AddWithValue("@Dis_TCXScore", txtTCXScore.Text)
                cmd.Parameters.AddWithValue("@Dis_Name", "")
                cmd.Parameters.AddWithValue("@Dis_Notes", "")
                cmd.Parameters.AddWithValue("@Dis_AppComments", "")

                cmd.Parameters.AddWithValue("@One_1", cbo1_1.Text)
                cmd.Parameters.AddWithValue("@One_2", cbo1_2.Text)
                cmd.Parameters.AddWithValue("@One_3", cbo1_3.Text)
                cmd.Parameters.AddWithValue("@One_1Note", txt1_1.Text)
                cmd.Parameters.AddWithValue("@One_2Note", txt1_2.Text)
                cmd.Parameters.AddWithValue("@One_3Note", txt1_3.Text)

                cmd.Parameters.AddWithValue("@Two_1", cbo2_1.Text)
                cmd.Parameters.AddWithValue("@Two_1Note", txt2_1.Text)

                cmd.Parameters.AddWithValue("@Three_1", cbo3_1.Text)
                cmd.Parameters.AddWithValue("@Three_2", cbo3_2.Text)
                cmd.Parameters.AddWithValue("@Three_3", cbo3_3.Text)
                cmd.Parameters.AddWithValue("@Three_4", cbo3_4.Text)
                cmd.Parameters.AddWithValue("@Three_5", cbo3_5.Text)
                cmd.Parameters.AddWithValue("@Three_6", cbo3_6.Text)
                cmd.Parameters.AddWithValue("@Three_7", cbo3_7.Text)



                cmd.Parameters.AddWithValue("@Three_1Note", txt3_1.Text)
                cmd.Parameters.AddWithValue("@Three_2Note", txt3_2.Text)
                cmd.Parameters.AddWithValue("@Three_3Note", txt3_3.Text)
                cmd.Parameters.AddWithValue("@Three_4Note", txt3_4.Text)
                cmd.Parameters.AddWithValue("@Three_5Note", txt3_5.Text)
                cmd.Parameters.AddWithValue("@Three_6Note", txt3_6.Text)
                cmd.Parameters.AddWithValue("@Three_7Note", txt3_7.Text)



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





                cmd.Parameters.AddWithValue("@Six_1Note", txt6_1.Text)




                cmd.Parameters.AddWithValue("@Seven_1", cbo7_1.Text)
                cmd.Parameters.AddWithValue("@Seven_2", cbo7_2.Text)
                cmd.Parameters.AddWithValue("@Seven_3", cbo7_3.Text)


                cmd.Parameters.AddWithValue("@Seven_1Note", txt7_1.Text)
                cmd.Parameters.AddWithValue("@Seven_2Note", txt7_2.Text)
                cmd.Parameters.AddWithValue("@Seven_3Note", txt7_3.Text)



                cmd.Parameters.AddWithValue("@1_1", txt1_1a.Text)
                cmd.Parameters.AddWithValue("@1_2", txt1_2a.Text)
                cmd.Parameters.AddWithValue("@1_3", txt1_3a.Text)


                cmd.Parameters.AddWithValue("@2_1", txt2_1a.Text)

                cmd.Parameters.AddWithValue("@3_1", txt3_1a.Text)
                cmd.Parameters.AddWithValue("@3_2", txt3_2a.Text)
                cmd.Parameters.AddWithValue("@3_3", txt3_3a.Text)
                cmd.Parameters.AddWithValue("@3_4", txt3_4a.Text)
                cmd.Parameters.AddWithValue("@3_5", txt3_5a.Text)
                cmd.Parameters.AddWithValue("@3_6", txt3_6a.Text)
                cmd.Parameters.AddWithValue("@3_7", txt3_7a.Text)


                cmd.Parameters.AddWithValue("@4_1", txt4_1a.Text)
                cmd.Parameters.AddWithValue("@4_2", txt4_2a.Text)
                cmd.Parameters.AddWithValue("@4_3", txt4_3a.Text)

                cmd.Parameters.AddWithValue("@5_1", txt5_1a.Text)
                cmd.Parameters.AddWithValue("@5_2", txt5_2a.Text)

                cmd.Parameters.AddWithValue("@6_1", txt6_1a.Text)




                cmd.Parameters.AddWithValue("@7_1", txt7_1a.Text)
                cmd.Parameters.AddWithValue("@7_2", txt7_2a.Text)
                cmd.Parameters.AddWithValue("@7_3", txt7_3a.Text)



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
                cmd.Parameters.AddWithValue("@Supervisor", cboSupervisor.Text)
                cmd.Parameters.AddWithValue("@Week_Number", Form2.lblYear.Text + " - " + "Week " + lblWeekNumber.Text)
                cmd.Parameters.AddWithValue("@TCX_Score", lblTCXscore.Text)
                cmd.Parameters.AddWithValue("@Month", Form2.lblMonth.Text)
                cmd.Parameters.AddWithValue("@EditedQA", "0")
                cmd.Parameters.AddWithValue("@SRType", cboSRType.Text)

                cmd.Parameters.AddWithValue("@CSATScore", txtCSATScore.Text)

                cmd.Parameters.AddWithValue("@CSATQ1", cboCSAT1.Text)
                cmd.Parameters.AddWithValue("@CSATQ2", cboCSAT2.Text)
                cmd.Parameters.AddWithValue("@CSATQ3", cboCSAT3.Text)
                cmd.Parameters.AddWithValue("@CSATQ4", cboCSAT4.Text)
                cmd.Parameters.AddWithValue("@CSATQ5", cboCSAT5.Text)
                cmd.Parameters.AddWithValue("@CSATQ6", cboCSAT6.Text)


                cmd.Parameters.AddWithValue("@CallerType", cboCallerType.Text)



                cmd.ExecuteNonQuery()

                con.Close()



            End Using


            ' MsgBox("Info saved")

            '   End If

            ' Saver.Enabled = True

            ExcelSaver.Enabled = True





        Catch ex As Exception

            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)

            Me.Cursor = Cursors.Hand

            ProgressBar1.Value = 0
            lblprogr.Text = 0


            QACallEnable()

            buttonEnables()

            Saver.Enabled = False


        End Try


    End Sub



    Public Sub Store2()




        Try




            con = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")




            con.Open()



            Dim SQL As String = "INSERT INTO [QAMainDB] ([ContactID],[EMID],[RegID],[AHT],[CType],[QA_Agent],[QA_Team],[QA_ContactDate],[QA_OrderID],[QA_Date],[QA_Comments],[QA_Opp],[CI_Name],[CI_Account],[CI_Company],[CI_Phone],[CI_Email],[Rev_Date],[Rev_Manager],[Rev_Comments],[Dis_Score],[Dis_Name],[Dis_Notes],[Dis_AppComments],[One_1],[One_2],[One_3],[One_1Note],[One_2Note],[One_3Note],[Two_1],[Two_1Note],[Three_1],[Three_2],[Three_3],[Three_4],[Three_5],[Three_6],[Three_7],[Three_1Note],[Three_2Note],[Three_3Note],[Three_4Note],[Three_5Note],[Three_6Note],[Three_7Note],[Four_1],[Four_2],[Four_3],[Four_1Note],[Four_2Note],[Four_3Note],[Five_1],[Five_2],[Five_1Note],[Five_2Note],[Six_1],[Six_1Note],[Seven_1],[Seven_2],[Seven_3],[Seven_1Note],[Seven_2Note],[Seven_3Note],[QAScore],[Autofail],[Auditor],[Supervisor],[Week_Number],[EditedQA],[TCX_Score],[1_1],[1_2],[1_3],[2_1],[3_1],[3_2],[3_3],[3_4],[3_5],[3_6],[3_7],[4_1],[4_2],[4_3],[5_1],[5_2],[6_1],[7_1],[7_2],[7_3],[Month],[PendingDisputeID],[Dis_TCXScore],[SRType],[CSATScore],[CSATQ1],[CSATQ2],[CSATQ3],[CSATQ4],[CSATQ5],[CSATQ6],[CallerType])  Values (@ContactID,@EMID,@RegID,@AHT,@CType, @QA_Agent, @QA_Team, @QA_ContactDate, @QA_OrderID, @QA_Date, @QA_Comments, @QA_Opp, @CI_Name, @CI_Account, @CI_Company, @CI_Phone, @CI_Email, @Rev_Date, @Rev_Manager, @Rev_Comments, @Dis_Score, @Dis_Name, @Dis_Notes, @Dis_AppComments, @One_1, @One_2, @One_3, @One_1Note, @One_2Note, @One_3Note, @Two_1, @Two_1Note, @Three_1, @Three_2, @Three_3, @Three_4, @Three_5, @Three_6, @Three_7,@Three_1Note, @Three_2Note, @Three_3Note, @Three_4Note, @Three_5Note, @Three_6Note, @Three_7Note,@Four_1, @Four_2, @Four_3, @Four_1Note, @Four_2Note, @Four_3Note, @Five_1, @Five_2, @Five_1Note, @Five_2Note, @Six_1, @Six_1Note, @Seven_1, @Seven_2, @Seven_3, @Seven_1Note, @Seven_2Note, @Seven_3Note, @QAScore, @Autofail, @Auditor, @Supervisor, @Week_Number, @EditedQA,@TCX_Score,@1_1,@1_2,@1_3,@2_1,@3_1,@3_2,@3_3,@3_4,@3_5,@3_6,@3_7,@4_1,@4_2,@4_3,@5_1,@5_2,@6_1,@7_1,@7_2,@7_3,@Month,@PendingDisputeID,@Dis_TCXScore,@SRType,@CSATScore,@CSATQ1,@CSATQ2,@CSATQ3,@CSATQ4,@CSATQ5,@CSATQ6,@CallerType)"





            Using cmd As New SqlCommand(SQL, con)








                '  cmd.Parameters.AddWithValue("@SR", txtSR.Text)
                cmd.Parameters.AddWithValue("@ContactID", txtContactID.Text)
                cmd.Parameters.AddWithValue("@EMID", txtEMID.Text)
                cmd.Parameters.AddWithValue("@RegID", txtRegID.Text)
                cmd.Parameters.AddWithValue("@AHT", txtAHT.Text)
                cmd.Parameters.AddWithValue("@CType", "WOTC Inbound")
                cmd.Parameters.AddWithValue("@QA_Agent", cboSupervisorbox.Text)
                cmd.Parameters.AddWithValue("@QA_Team", txtTeamName.Text)
                cmd.Parameters.AddWithValue("@QA_ContactDate", dtpCondate.Value)
                cmd.Parameters.AddWithValue("@QA_OrderID", txtOrderID.Text)
                cmd.Parameters.AddWithValue("@QA_Date", Now)
                cmd.Parameters.AddWithValue("@QA_Comments", txtQACom.Text)
                cmd.Parameters.AddWithValue("@QA_Opp", txtQAAOO.Text)
                cmd.Parameters.AddWithValue("@CI_Name", txtContactName.Text)
                cmd.Parameters.AddWithValue("@CI_Account", txtAccountNum.Text)
                cmd.Parameters.AddWithValue("@CI_Company", txtCompany.Text)
                cmd.Parameters.AddWithValue("@CI_Phone", txtContactPhone.Text)
                cmd.Parameters.AddWithValue("@CI_Email", txtContactEmail.Text)


                If txtRevcom.Text = "" Then


                    cmd.Parameters.AddWithValue("@Rev_Date", "9/9/2020")
                    cmd.Parameters.AddWithValue("@Rev_Manager", lblQAauditor1.Text)
                    cmd.Parameters.AddWithValue("@Rev_Comments", "")
                    cmd.Parameters.AddWithValue("@PendingDisputeID", "Pending Review")

                ElseIf txtRevcom.Text <> "" Then

                    cmd.Parameters.AddWithValue("@Rev_Date", txtQADate.Text)
                    cmd.Parameters.AddWithValue("@Rev_Manager", lblQAauditor1.Text)
                    cmd.Parameters.AddWithValue("@Rev_Comments", txtRevcom.Text)
                    cmd.Parameters.AddWithValue("@PendingDisputeID", "Reviewed")


                End If

                '  cmd.Parameters.AddWithValue("@Dis_Score", lblQAScore1.Text)



                cmd.Parameters.AddWithValue("@Dis_TCXScore", lblTCXscore.Text)
                cmd.Parameters.AddWithValue("@Dis_Name", "")
                cmd.Parameters.AddWithValue("@Dis_Notes", "")
                cmd.Parameters.AddWithValue("@Dis_AppComments", "")

                cmd.Parameters.AddWithValue("@One_1", cbo1_1.Text)
                cmd.Parameters.AddWithValue("@One_2", cbo1_2.Text)
                cmd.Parameters.AddWithValue("@One_3", cbo1_3.Text)
                cmd.Parameters.AddWithValue("@One_1Note", txt1_1.Text)
                cmd.Parameters.AddWithValue("@One_2Note", txt1_2.Text)
                cmd.Parameters.AddWithValue("@One_3Note", txt1_3.Text)

                cmd.Parameters.AddWithValue("@Two_1", cbo2_1.Text)
                cmd.Parameters.AddWithValue("@Two_1Note", txt2_1.Text)

                cmd.Parameters.AddWithValue("@Three_1", cbo3_1.Text)
                cmd.Parameters.AddWithValue("@Three_2", cbo3_2.Text)
                cmd.Parameters.AddWithValue("@Three_3", cbo3_3.Text)
                cmd.Parameters.AddWithValue("@Three_4", cbo3_4.Text)
                cmd.Parameters.AddWithValue("@Three_5", cbo3_5.Text)
                cmd.Parameters.AddWithValue("@Three_6", cbo3_6.Text)
                cmd.Parameters.AddWithValue("@Three_7", cbo3_7.Text)



                cmd.Parameters.AddWithValue("@Three_1Note", txt3_1.Text)
                cmd.Parameters.AddWithValue("@Three_2Note", txt3_2.Text)
                cmd.Parameters.AddWithValue("@Three_3Note", txt3_3.Text)
                cmd.Parameters.AddWithValue("@Three_4Note", txt3_4.Text)
                cmd.Parameters.AddWithValue("@Three_5Note", txt3_5.Text)
                cmd.Parameters.AddWithValue("@Three_6Note", txt3_6.Text)
                cmd.Parameters.AddWithValue("@Three_7Note", txt3_7.Text)



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





                cmd.Parameters.AddWithValue("@Six_1Note", txt6_1.Text)




                cmd.Parameters.AddWithValue("@Seven_1", cbo7_1.Text)
                cmd.Parameters.AddWithValue("@Seven_2", cbo7_2.Text)
                cmd.Parameters.AddWithValue("@Seven_3", cbo7_3.Text)

                cmd.Parameters.AddWithValue("@Seven_1Note", txt7_1.Text)
                cmd.Parameters.AddWithValue("@Seven_2Note", txt7_2.Text)
                cmd.Parameters.AddWithValue("@Seven_3Note", txt7_3.Text)


                cmd.Parameters.AddWithValue("@1_1", txt1_1a.Text)
                cmd.Parameters.AddWithValue("@1_2", txt1_2a.Text)
                cmd.Parameters.AddWithValue("@1_3", txt1_3a.Text)


                cmd.Parameters.AddWithValue("@2_1", txt2_1a.Text)

                cmd.Parameters.AddWithValue("@3_1", txt3_1a.Text)
                cmd.Parameters.AddWithValue("@3_2", txt3_2a.Text)
                cmd.Parameters.AddWithValue("@3_3", txt3_3a.Text)
                cmd.Parameters.AddWithValue("@3_4", txt3_4a.Text)
                cmd.Parameters.AddWithValue("@3_5", txt3_5a.Text)
                cmd.Parameters.AddWithValue("@3_6", txt3_6a.Text)
                cmd.Parameters.AddWithValue("@3_7", txt3_7a.Text)


                cmd.Parameters.AddWithValue("@4_1", txt4_1a.Text)
                cmd.Parameters.AddWithValue("@4_2", txt4_2a.Text)
                cmd.Parameters.AddWithValue("@4_3", txt4_3a.Text)

                cmd.Parameters.AddWithValue("@5_1", txt5_1a.Text)
                cmd.Parameters.AddWithValue("@5_2", txt5_2a.Text)

                cmd.Parameters.AddWithValue("@6_1", txt6_1a.Text)




                cmd.Parameters.AddWithValue("@7_1", txt7_1a.Text)
                cmd.Parameters.AddWithValue("@7_2", txt7_2a.Text)
                cmd.Parameters.AddWithValue("@7_3", txt7_3a.Text)




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
                cmd.Parameters.AddWithValue("@Week_Number", Form2.lblYear.Text + " - " + "Week " + lblWeekNumber.Text)
                cmd.Parameters.AddWithValue("@EditedQA", "0")
                cmd.Parameters.AddWithValue("@TCX_Score", txtTCXScore.Text)
                cmd.Parameters.AddWithValue("@Month", Form2.lblMonth.Text)
                cmd.Parameters.AddWithValue("@SRType", cboSRType.Text)


                cmd.Parameters.AddWithValue("@CSATScore", txtCSATScore.Text)

                cmd.Parameters.AddWithValue("@CSATQ1", cboCSAT1.Text)
                cmd.Parameters.AddWithValue("@CSATQ2", cboCSAT2.Text)
                cmd.Parameters.AddWithValue("@CSATQ3", cboCSAT3.Text)
                cmd.Parameters.AddWithValue("@CSATQ4", cboCSAT4.Text)
                cmd.Parameters.AddWithValue("@CSATQ5", cboCSAT5.Text)
                cmd.Parameters.AddWithValue("@CSATQ6", cboCSAT6.Text)

                cmd.Parameters.AddWithValue("@CallerType", cboCallerType.Text)

                cmd.ExecuteNonQuery()






                con.Close()



            End Using


            ' MsgBox("Info saved")

            '   End If


            ExcelSaver2.Enabled = True


            '  Saver2.Enabled = True



        Catch ex As Exception


            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)


            Me.Cursor = Cursors.Hand

            ProgressBar1.Value = 0
            lblprogr.Text = 0


            QACallEnable()

            buttonEnables()

            Saver2.Enabled = False

        End Try


    End Sub


    Public Sub QAExcell()



        Try



            Dim oExcel As Object = CreateObject("Excel.Application")


            '' P Drive

            '  Dim oBook As Object = oExcel.Workbooks.Open("P:\QA Application\QA1\Call.xlsx")



            '' Resouce
            Dim exeDir As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)
            Dim xlpath = IO.Path.Combine(exeDir.DirectoryName, "WOTCInbound.xlsx")


            Dim obook As Object = oExcel.Workbooks.Open(xlpath)


            Dim oSheet As Object = obook.Worksheets("WOTC Inbound")



            ' oSheet.Range("D3").Value = "" & One

            oSheet.Range("D4").Value = "" & cbo1_1.Text
            oSheet.Range("D5").Value = "" & cbo1_2.Text
            oSheet.Range("D6").Value = "" & cbo1_3.Text



            oSheet.Range("H4").Value = "" & txt1_1.Text
            oSheet.Range("H5").Value = "" & txt1_2.Text
            oSheet.Range("H6").Value = "" & txt1_3.Text

            '  oSheet.Range("D7").Value = "" & two

            oSheet.Range("D8").Value = "" & cbo2_1.Text
            oSheet.Range("H8").Value = "" & txt2_1.Text

            '   oSheet.Range("D9").Value = "" & three

            oSheet.Range("D10").Value = "" & cbo3_1.Text
            oSheet.Range("D11").Value = "" & cbo3_2.Text
            oSheet.Range("D12").Value = "" & cbo3_3.Text
            oSheet.Range("D13").Value = "" & cbo3_4.Text
            oSheet.Range("D14").Value = "" & cbo3_5.Text
            oSheet.Range("D15").Value = "" & cbo3_6.Text
            oSheet.Range("D16").Value = "" & cbo3_7.Text




            oSheet.Range("H10").Value = "" & txt3_1.Text
            oSheet.Range("H11").Value = "" & txt3_2.Text
            oSheet.Range("H12").Value = "" & txt3_3.Text
            oSheet.Range("H13").Value = "" & txt3_4.Text
            oSheet.Range("H14").Value = "" & txt3_5.Text
            oSheet.Range("H15").Value = "" & txt3_6.Text
            oSheet.Range("H16").Value = "" & txt3_7.Text


            '  oSheet.Range("D18").Value = "" & Four

            oSheet.Range("D18").Value = "" & Cbo4_1.Text
            oSheet.Range("D19").Value = "" & cbo4_2.Text
            oSheet.Range("D20").Value = "" & cbo4_3.Text

            oSheet.Range("H18").Value = "" & txt4_1.Text
            oSheet.Range("H19").Value = "" & txt4_2.Text
            oSheet.Range("H20").Value = "" & txt4_3.Text


            '  oSheet.Range("D22").Value = "" & Five


            oSheet.Range("D22").Value = "" & cbo5_1.Text
            oSheet.Range("D23").Value = "" & cbo5_2.Text

            oSheet.Range("H22").Value = "" & txt5_1.Text
            oSheet.Range("H23").Value = "" & txt5_2.Text

            '  oSheet.Range("D25").Value = "" & Six


            oSheet.Range("D25").Value = "" & cbo6_1.Text

            oSheet.Range("H25").Value = "" & txt6_1.Text



            '  oSheet.Range("D29").Value = "" & seven


            oSheet.Range("D27").Value = "" & cbo7_1.Text
            oSheet.Range("D28").Value = "" & cbo7_2.Text
            oSheet.Range("D29").Value = "" & cbo7_3.Text



            oSheet.Range("H27").Value = "" & txt7_1.Text
            oSheet.Range("H28").Value = "" & txt7_2.Text
            oSheet.Range("H29").Value = "" & txt7_3.Text



            If cboAutoFail.Checked = True Then

                oSheet.Range("D30").Value = "0"

            Else

                oSheet.Range("D30").Value = txtQAScore.Text

            End If


            oSheet.Range("C31").Value = txtCSATScore.Text


            oSheet.Range("C32").Value = txtContactID.Text
            oSheet.Range("C33").Value = "" & cboAgentName.Text
            oSheet.Range("C34").Value = "WOTC Inbound"
            oSheet.Range("C35").Value = "" & cboCallerType.Text
            oSheet.Range("C36").Value = "" & txtCompany.Text
            oSheet.Range("C37").Value = "" & txtEMID.Text
            oSheet.Range("C38").Value = "" & txtRegID.Text
            oSheet.Range("C39").Value = "" & txtAHT.Text
            oSheet.Range("C40").Value = dtpCondate.Text
            oSheet.Range("C41").Value = "" & cboAF.Text
            oSheet.Range("C70").Value = "" & lblQAauditor1.Text




            oSheet.Range("B43").Value = txtQACom.Text
            oSheet.Range("B47").Value = txtQAAOO.Text
            oSheet.Range("B63").Value = txtRevcom.Text


            oSheet.Range("C54").Value = "" & txtCSATScore.Text

            oSheet.Range("C55").Value = "" & cboCSAT1.Text
            oSheet.Range("C56").Value = "" & cboCSAT2.Text
            oSheet.Range("C57").Value = "" & cboCSAT3.Text
            oSheet.Range("C58").Value = "" & cboCSAT4.Text
            oSheet.Range("C59").Value = "" & cboCSAT5.Text
            oSheet.Range("C60").Value = "" & cboCSAT6.Text


            ' iF contactid is being used   

            If txtContactID.Text <> String.Empty And txtSR.Text = "1-" Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtContactID.Text & " " & cboAgentName.Text & "-" & "WOTC Inbound QA Scorecard.xlsx")

                Saver2.Enabled = True


            End If

            ' If SR is being used
            If txtContactID.Text = String.Empty Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & "WOTC Inbound QA Scorecard.xlsx")


                Saver.Enabled = True

            End If

            ''if both are filled out
            If txtContactID.Text <> String.Empty And txtSR.Text <> "1-" Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtSR.Text & " " & cboAgentName.Text & "-" & " WOTC Inbound QA Scorecard.xlsx")


                Saver.Enabled = True

            End If


            oExcel.Quit()

            ''    Saver.Enabled = True


        Catch ex As Exception

            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)


            Me.Cursor = Cursors.Hand

            ProgressBar1.Value = 0
            lblprogr.Text = 0


            QACallEnable()

            buttonEnables()

            Saver.Enabled = False


        End Try

    End Sub

    Public Sub QAExcel2()


        Try



            Dim oExcel As Object = CreateObject("Excel.Application")


            '' P Drive

            '   Dim oBook As Object = oExcel.Workbooks.Open("P:\QA Application\QA1\Call.xlsx")



            '' Resouce
            Dim exeDir As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)
            Dim xlpath = IO.Path.Combine(exeDir.DirectoryName, "WOTCInbound.xlsx")


            Dim obook As Object = oExcel.Workbooks.Open(xlpath)


            Dim oSheet As Object = obook.Worksheets("WOTC Inbound")



            ' oSheet.Range("D3").Value = "" & One

            oSheet.Range("D4").Value = "" & cbo1_1.Text
            oSheet.Range("D5").Value = "" & cbo1_2.Text
            oSheet.Range("D6").Value = "" & cbo1_3.Text



            oSheet.Range("H4").Value = "" & txt1_1.Text
            oSheet.Range("H5").Value = "" & txt1_2.Text
            oSheet.Range("H6").Value = "" & txt1_3.Text

            '  oSheet.Range("D7").Value = "" & two

            oSheet.Range("D8").Value = "" & cbo2_1.Text
            oSheet.Range("H8").Value = "" & txt2_1.Text

            '   oSheet.Range("D9").Value = "" & three

            oSheet.Range("D10").Value = "" & cbo3_1.Text
            oSheet.Range("D11").Value = "" & cbo3_2.Text
            oSheet.Range("D12").Value = "" & cbo3_3.Text
            oSheet.Range("D13").Value = "" & cbo3_4.Text
            oSheet.Range("D14").Value = "" & cbo3_5.Text
            oSheet.Range("D15").Value = "" & cbo3_6.Text
            oSheet.Range("D16").Value = "" & cbo3_7.Text




            oSheet.Range("H10").Value = "" & txt3_1.Text
            oSheet.Range("H11").Value = "" & txt3_2.Text
            oSheet.Range("H12").Value = "" & txt3_3.Text
            oSheet.Range("H13").Value = "" & txt3_4.Text
            oSheet.Range("H14").Value = "" & txt3_5.Text
            oSheet.Range("H15").Value = "" & txt3_6.Text
            oSheet.Range("H16").Value = "" & txt3_7.Text


            '  oSheet.Range("D18").Value = "" & Four

            oSheet.Range("D18").Value = "" & Cbo4_1.Text
            oSheet.Range("D19").Value = "" & cbo4_2.Text
            oSheet.Range("D20").Value = "" & cbo4_3.Text

            oSheet.Range("H18").Value = "" & txt4_1.Text
            oSheet.Range("H19").Value = "" & txt4_2.Text
            oSheet.Range("H20").Value = "" & txt4_3.Text


            '  oSheet.Range("D22").Value = "" & Five


            oSheet.Range("D22").Value = "" & cbo5_1.Text
            oSheet.Range("D23").Value = "" & cbo5_2.Text

            oSheet.Range("H22").Value = "" & txt5_1.Text
            oSheet.Range("H23").Value = "" & txt5_2.Text

            '  oSheet.Range("D25").Value = "" & Six


            oSheet.Range("D25").Value = "" & cbo6_1.Text

            oSheet.Range("H25").Value = "" & txt6_1.Text



            '  oSheet.Range("D29").Value = "" & seven


            oSheet.Range("D27").Value = "" & cbo7_1.Text
            oSheet.Range("D28").Value = "" & cbo7_2.Text
            oSheet.Range("D29").Value = "" & cbo7_3.Text



            oSheet.Range("H27").Value = "" & txt7_1.Text
            oSheet.Range("H28").Value = "" & txt7_2.Text
            oSheet.Range("H29").Value = "" & txt7_3.Text


            If cboAutoFail.Checked = True Then

                oSheet.Range("D30").Value = "0"

            Else

                oSheet.Range("D30").Value = txtQAScore.Text

            End If


            oSheet.Range("C31").Value = txtCSATScore.Text


            oSheet.Range("C32").Value = txtContactID.Text
            oSheet.Range("C33").Value = "" & cboAgentName.Text
            oSheet.Range("C34").Value = "WOTC Inbound"
            oSheet.Range("C35").Value = "" & cboCallerType.Text
            oSheet.Range("C36").Value = "" & txtCompany.Text
            oSheet.Range("C37").Value = "" & txtEMID.Text
            oSheet.Range("C38").Value = "" & txtRegID.Text
            oSheet.Range("C39").Value = "" & txtAHT.Text
            oSheet.Range("C40").Value = dtpCondate.Text
            oSheet.Range("C41").Value = "" & cboAF.Text
            oSheet.Range("C70").Value = "" & lblQAauditor1.Text




            oSheet.Range("B43").Value = txtQACom.Text
            oSheet.Range("B47").Value = txtQAAOO.Text
            oSheet.Range("B63").Value = txtRevcom.Text

            oSheet.Range("C54").Value = "" & txtCSATScore.Text

            oSheet.Range("C55").Value = "" & cboCSAT1.Text
            oSheet.Range("C56").Value = "" & cboCSAT2.Text
            oSheet.Range("C57").Value = "" & cboCSAT3.Text
            oSheet.Range("C58").Value = "" & cboCSAT4.Text
            oSheet.Range("C59").Value = "" & cboCSAT5.Text
            oSheet.Range("C60").Value = "" & cboCSAT6.Text


            ' iF contactid is being used   

            If txtContactID.Text <> String.Empty And txtSR.Text = "1-" Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtContactID.Text & " " & cboSupervisorbox.Text & "-" & "WOTC Inbound QA Scorecard.xlsx")

                Saver2.Enabled = True

            End If

            ' If SR is being used
            If txtContactID.Text = String.Empty Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "WOTC Inbound QA Scorecard.xlsx")


                Saver.Enabled = True

            End If

            ''If both are filled out

            If txtContactID.Text <> String.Empty And txtSR.Text <> "1-" Then

                obook.SaveAs(Desk & "\QA2\" & "" & txtSR.Text & " " & cboSupervisorbox.Text & "-" & "WOTC Inbound QA Scorecard.xlsx")


                Saver.Enabled = True

            End If

            oExcel.Quit()




        Catch ex As Exception

            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)



            Me.Cursor = Cursors.Hand

            ProgressBar1.Value = 0
            lblprogr.Text = 0


            QACallEnable()

            buttonEnables()

            Saver.Enabled = False



        End Try











    End Sub




    Private Sub btnSaveScoreCard_Click(sender As Object, e As EventArgs) Handles btnSaveScoreCard.Click



        '    Try
        MissedWeightsCalc()

        'If txtSR.Text <> "1-" And txtSR.MaskFull = False Then


        '    MsgBox("Please enter a valid SR#")


        'Else


        If txtContactID.Text = "" Then

            MsgBox("A Contact ID Is required before saving", MessageBoxButtons.RetryCancel)


        Else


            If cboSRType.Text = "" Then

                MsgBox("A SR Type must be selected before saving", MessageBoxButtons.RetryCancel)

                Me.ActiveControl = cboSRType

            Else

                If cboCallerType.Text = "" Then

                    MsgBox("A Caller Type must be selected before saving", MessageBoxButtons.RetryCancel)

                    Me.ActiveControl = cboCallerType

                Else

                    If cboSupervisor.Text = "Supervisor" Then


                    MsgBox("Please be advised you must select an 'Supervisor' before proceeding", MessageBoxButtons.RetryCancel)


                Else

                    If cboAgentName.Text = "Agent Name" Then


                        MsgBox("Please be advised you must select an 'agent name' before proceeding", MessageBoxButtons.RetryCancel)


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

                                        If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then




                                        Else

                                            SplashScreenManager1.ShowWaitForm()
                                            Me.Cursor = Cursors.WaitCursor

                                            buttondisables()

                                            QACalldisableControls()



                                            Me.ActiveControl = txtContactID




                                            BackgroundWorker1.RunWorkerAsync()

                                            'PleaseWait.ShowDialog()

                                            Store()




                                        End If





                                    End If

                                End If

                            End If

                            '




                        End If

                    End If

                End If
            End If

        End If



        'Catch ex As Exception



        '    MsgBox(ex.Message)



        'End Try






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


        Dim int4_1 As Integer = Cbo4_1.Text
        Dim int4_2 As Integer = cbo4_2.Text
        Dim int4_3 As Integer = cbo4_3.Text


        Dim int5_1 As Integer = cbo5_1.Text
        Dim int5_2 As Integer = cbo5_2.Text

        Dim int6_1 As Integer = cbo6_1.Text




        Dim int7_1 As Integer = cbo7_1.Text
        Dim int7_2 As Integer = cbo7_2.Text
        Dim int7_3 As Integer = cbo7_3.Text






        One = int1_1 + int1_2 + int1_3

        two = int2_1

        three = int3_1 + int3_2 + int3_3 + int3_4 + int3_5 + int3_6 + int3_7

        Four = int4_1 + int4_2 + int4_3

        Five = int5_1 + int5_2

        Six = int6_1

        seven = int7_1 + int7_2 + int7_3






        intQascoreTotal = int1_1 + int1_2 + int1_3 + int2_1 + int3_1 + int3_2 + int3_3 + int3_4 + int3_5 + int3_6 + int3_7 + int4_1 + int4_2 + int4_3 + int5_1 + int5_2 + int6_1 + int7_1 + int7_2 + int7_3


        lblQAScore1.Text = intQascoreTotal
        txtQAScore.Text = intQascoreTotal




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
        ' cbo3_8.Enabled = False

        Cbo4_1.Enabled = False
        cbo4_2.Enabled = False
        cbo4_3.Enabled = False


        cbo5_1.Enabled = False
        cbo5_2.Enabled = False

        cbo6_1.Enabled = False
        'cbo6_2.Enabled = False


        cbo7_1.Enabled = False
        cbo7_2.Enabled = False
        cbo7_3.Enabled = False



        ''reset Textboxes

        'txt1_1.Enabled = False
        'txt1_2.Enabled = False

        'txt1_3.Enabled = False


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
        lblTCXscore.Visible = False



        cboAgentName.Text = "Agent Name"
        '  cboTeamName.Text = "Team Name"

        txtSR.Clear()
        '  lblQAScore1.Text = "100"
    End Sub


    Public Sub QACallclear()


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
        cbo3_7.Text = 4


        Cbo4_1.Text = 5
        cbo4_2.Text = 5
        cbo4_3.Text = 5


        cbo5_1.Text = 7
        cbo5_2.Text = 8

        cbo6_1.Text = 15


        cbo7_1.Text = 5
        cbo7_2.Text = 5
        cbo7_3.Text = 5


        MissedWeightsReset()


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




        txt4_1.Clear()
        txt4_2.Clear()
        txt4_3.Clear()

        txt5_1.Clear()
        txt5_2.Clear()




        txt6_1.Clear()




        txt7_1.Clear()
        txt7_2.Clear()
        txt7_3.Clear()


        txtQAAOO.Clear()
        txtQACom.Clear()

        txtAgentEmail.Clear()



        ''
        txt1_1.BackColor = Color.White




        txt1_2.BackColor = Color.White

        txt1_3.BackColor = Color.White


        txt2_1.BackColor = Color.White


        txt3_1.BackColor = Color.White
        txt3_2.BackColor = Color.White
        txt3_3.BackColor = Color.White
        txt3_4.BackColor = Color.White
        txt3_5.BackColor = Color.White
        txt3_6.BackColor = Color.White
        txt3_7.BackColor = Color.White




        txt4_1.BackColor = Color.White
        txt4_2.BackColor = Color.White
        txt4_3.BackColor = Color.White

        txt5_1.BackColor = Color.White
        txt5_2.BackColor = Color.White




        txt6_1.BackColor = Color.White



        txt7_1.BackColor = Color.White
        txt7_2.BackColor = Color.White
        txt7_3.BackColor = Color.White




        txtSR.Clear()
        txtContactID.Clear()
        txtContactName.Clear()
        txtContactEmail.Clear()
        txtContactPhone.Clear()
        txtAccountNum.Clear()
        txtCompany.Clear()
        txtOrderID.Clear()

        txtRevcom.Clear()



        ' txtTeamName.Clear()


        cboCSAT1.SelectedIndex = -1
        cboCSAT2.SelectedIndex = -1
        cboCSAT3.SelectedIndex = -1
        cboCSAT4.SelectedIndex = -1
        cboCSAT5.SelectedIndex = -1
        cboCSAT6.SelectedIndex = -1

        txtCSATScore.Clear()
        txtTCXScore.Clear()

        txtQAScore.Text = "100"
        lblQAScore1.Text = "100"
        ' lblQAScore1.Visible = True
        cboSRType.SelectedIndex = -1



        txtEMID.Clear()
        txtAHT.Clear()
        txtRegID.Clear()
        cboCallerType.SelectedIndex = -1


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
        ' cbo3_8.Enabled = True

        Cbo4_1.Enabled = True
        cbo4_2.Enabled = True
        cbo4_3.Enabled = True


        cbo5_1.Enabled = True
        cbo5_2.Enabled = True

        cbo6_1.Enabled = True
        '  cbo6_2.Enabled = True


        cbo7_1.Enabled = True
        cbo7_2.Enabled = True
        cbo7_3.Enabled = True



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
        '  txt3_8.Enabled = True



        txt4_1.Enabled = True
        txt4_2.Enabled = True
        txt4_3.Enabled = True

        txt5_1.Enabled = True
        txt5_2.Enabled = True




        txt6_1.Enabled = True
        '    txt6_2.Enabled = True



        txt7_1.Enabled = True
        txt7_2.Enabled = True
        txt7_3.Enabled = True

    End Sub


    Private Sub btnQaSetup_Click(sender As Object, e As EventArgs) Handles btnQaSetup.Click

        Try

            Me.Cursor = Cursors.WaitCursor


            reset()

            Form2.Clear()


            Form2.cboAgentName.Enabled = True

            Form2.cboContactType.Enabled = True

            Form2.cboSupervisor.Enabled = True




            'Form2.cboAgentName.Text = cboAgentName.Text

            'Form2.cboSupervisor.Text = cboSupervisor.Text







            'Form2.txtSRNumber.Text = txtSR.Text

            'Form2.txtContactID.Text = txtContactID.Text

            'Form2.txtContactName.Text = txtContactName.Text


            'Form2.txtContactEmail.Text = txtContactEmail.Text

            'Form2.txtContactPhone.Text = txtContactPhone.Text


            'Form2.txtAccountNum.Text = txtAccountNum.Text


            'Form2.txtCompany.Text = txtCompany.Text


            'Form2.txtOrderID.Text = txtOrderID.Text

            'Form2.DateTimePicker1.Text = dtpCondate.Text


            '' User
            ''Jira

            Form2.Show()

            Me.Hide()


            Me.Cursor = Cursors.Hand




        Catch ex As Exception



            MsgBox(ex.Message)

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





            ' Store()


            '  Catch ex As SqlException




        Catch ex As Exception



            MsgBox(ex.Message)




        End Try





    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

        ProgressBar1.Value = e.ProgressPercentage



    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted






        Me.Cursor = Cursors.Hand


        ''PleaseWait.Hide()

        ''If MsgBox("The audit for " & cboAgentName.Text & " " & "was successfully saved, would you like to start a New 'Call’ audit?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then


        ''    reset()

        ''    buttonEnables()


        ''    Form2.Clear()


        ''    Form2.Show()

        ''    Me.Hide()


        ''Else

        ''    buttonEnables()

        ''    reset()

        ''    '  cboTeamName.Text = "Team Name"

        ''End If




    End Sub



    Public Sub reset()



        ''Reset Scorecard at a glance info

        resetatglance()

        ''Reset scorecard

        QACallclear()

        ''Transfer Qa Name to Qasetupform


        '  Form2.lblQAauditor.Text = lblQAauditor1.Text


        ''Reable buttons

        QACallEnable()


        Me.Cursor = Cursors.Hand

        '  Me.Hide()

        cboAutoFail.Checked = False

        cboAF.Visible = False


        ProgressBar1.Value = 0
        lblprogr.Text = 0


        ' txtQACom.BackColor = Color.White





    End Sub

    'Public Sub spellcheck()


    '    Try
    '        ' Create Word and temporary document objects.
    '        Dim objWord As Object
    '        Dim objTempDoc As Object

    '        ' Declare an IDataObject to hold the data returned from the 
    '        ' clipboard.
    '        Dim iData As IDataObject

    '        ' If there is no data to spell check, then exit sub here.
    '        If txt1_1.Text = "" Then

    '            Exit Sub
    '        End If

    '        objWord = New Word.Application()
    '        objTempDoc = objWord.Documents.Add
    '        objWord.Visible = False

    '        ' Position Word off the screen...this keeps Word invisible 
    '        ' throughout.
    '        objWord.WindowState = 0
    '        objWord.Top = -3000

    '        ' Copy the contents of the textbox to the clipboard
    '        Clipboard.SetDataObject(txt1_1.Text)

    '        ' With the temporary document, perform either a spell check or a 
    '        ' complete
    '        ' grammar check, based on user selection.
    '        With objTempDoc
    '            .Content.Paste()
    '            .Activate()

    '            .CheckSpelling()

    '            .CheckGrammar()

    '            ' After user has made changes, use the clipboard to
    '            ' transfer the contents back to the text box

    '            .Content.Copy()
    '            iData = Clipboard.GetDataObject
    '            If iData.GetDataPresent(DataFormats.Text) Then
    '                txt1_1.Text = CType(iData.GetData(DataFormats.Text),
    '                    String)
    '            End If
    '            .Saved = True



    '        End With

    '        objWord.Quit()


    '    Catch Excep As Exception
    '        MessageBox.Show(Excep.Message)

    '    End Try









    'End Sub






    Private Sub SpellOrGrammarCheck(ByVal blnSpellOnly As Boolean)


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
                .close()

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











    Private Sub Button13_Click(sender As Object, e As EventArgs)


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





    Private Sub BackgroundWorker2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork




        Try

            For i = 0 To 100

                System.Threading.Thread.Sleep(30)
                Me.BackgroundWorker1.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i
            Next




            ''


            ' Store2()

            ' Send to Excell
            '  QAExcell()



            'StoreCallThread = New System.Threading.Thread(AddressOf Store2)

            'StoreCallThread.Start()




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub



    Private Sub QACallScorecard_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        Try

            If MessageBox.Show("Are you sure to close this application?", "FADV Quality Assurance Application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then

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


            MsgBox("Are you sure you want to Auto Fail this agent? This will give a score of a 0, but the weights will still be recorded.")


            cboAF.Visible = True




        ElseIf cboAutoFail.CheckState = CheckState.Unchecked Then


            cboAF.Visible = False



        End If



    End Sub









    Private Sub cboTeamName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSupervisor.SelectedIndexChanged


        Try




            Me.Cursor = Cursors.WaitCursor


            cboAgentName.Text = "Please wait, Loading.."

            resetcombo()

            BackgroundWorker5.RunWorkerAsync()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try






    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        txtSR.Text = "10-12323233"
        txtContactID.Text = "1212111"
        txtContactName.Text = "Crystal Smith"
        txtContactEmail.Text = "CrystalSmith@Gmail.com"
        txtContactPhone.Text = "5558889695"
        txtAccountNum.Text = "abc32323"
        txtCompany.Text = "Little Leauge"
        txtOrderID.Text = "95955555"

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSupervisorbox.SelectedIndexChanged

        Try

            txtTeamName.Text = "Please wait, Loading.."


            BackgroundWorker7.RunWorkerAsync()


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


        '' If MsgBox(cboSupervisorbox.Text & " " & "" & "scored a total of" & " " & lblQAScore1.Text & " " & "points on this QA audit, would you like to start a new ‘Call’ audit?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then

        'If MsgBox("The audit for " & cboSupervisorbox.Text & " " & "was successfully saved, would you like to start a New 'Call’ audit?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then


        '    reset()

        '    buttonEnables()


        '    Form2.Clear()


        '    Form2.Show()

        '    Me.Hide()


        'Else


        '    buttonEnables()

        '    reset()

        '    'cboSupervisor.Text = "Supervisor"
        '    'cboAgentName.Text = "Agent Name"
        '    cboSupervisorbox.Text = "Agent Name"
        '    'cboSupervisorbox.Text = "Agent Name"

        'End If







    End Sub



    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles btnSpellChecker.Click

        Try

            SpellChecker2.CheckContainer(Me)



        Catch ex As Exception



            MsgBox(ex.Message)


        End Try



    End Sub


    Private Sub cboContactType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboContactTypeCall.SelectedIndexChanged




        Dim msg = "Are you sure you want to change the Scorecard?"

        Dim title = "FADV QA Application"

        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        Dim responce = MsgBox(msg, style, title)





        If responce = MsgBoxResult.Yes Then



            If cboContactTypeCall.Text = "Call" Then




            ElseIf cboContactTypeCall.Text = "Email" Then


                QAEmailScorecard.cboAgentName.Text = cboAgentName.Text
                QAEmailScorecard.txtTeamName.Text = txtTeamName.Text


                QAEmailScorecard.txtSR.Text = txtSR.Text
                QAEmailScorecard.txtContactID.Text = txtContactID.Text


                QAEmailScorecard.txtContactName.Text = txtContactName.Text
                QAEmailScorecard.txtContactEmail.Text = txtContactEmail.Text
                QAEmailScorecard.txtContactPhone.Text = txtContactPhone.Text
                QAEmailScorecard.txtAccountNum.Text = txtAccountNum.Text
                QAEmailScorecard.txtCompany.Text = txtCompany.Text
                QAEmailScorecard.txtOrderID.Text = txtOrderID.Text
                '  QAEmailScorecard.DateTimePicker1.Text = dtpCondate.Text

                '  QAEmailScorecard.DateTimePicker1.Text = ProgramDate.ToString(ProgramDateForamt)


                QAEmailScorecard.lblQAauditor1.Text = lblQAauditor1.Text

                ' QAEmailScorecard.cboContactTypeEmail.Text = "Email"

                QAEmailScorecard.Show()


                Me.Hide()


            ElseIf cboContactTypeCall.Text = "Chat" Then


                QAChatScorecard.txtSR.Text = txtSR.Text
                QAChatScorecard.txtContactID.Text = txtContactID.Text


                QAChatScorecard.txtContactName.Text = txtContactName.Text
                QAChatScorecard.txtContactEmail.Text = txtContactEmail.Text
                QAChatScorecard.txtContactPhone.Text = txtContactPhone.Text
                QAChatScorecard.txtAccountNum.Text = txtAccountNum.Text
                QAChatScorecard.txtCompany.Text = txtCompany.Text
                QAChatScorecard.txtOrderID.Text = txtOrderID.Text

                'QAChatScorecard.DateTimePicker1.Text = ProgramDate.ToString(ProgramDateForamt)

                ' QAChatScorecard.DateTimePicker1.Text = ProgramDate.ToString(dtpCondate.Text)

                QAChatScorecard.lblQAauditor1.Text = lblQAauditor1.Text

                '    QAChatScorecard.cboContactTypeChat.Text = "Chat"

                QAChatScorecard.Show()
                Me.Hide()



            End If







        Else




        End If



    End Sub

    Private Sub txtSR_TextChanged(sender As Object, e As EventArgs)




    End Sub

    Private Sub BackgroundWorker3_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork


        Fillcombo()


    End Sub

    Private Sub BackgroundWorker4_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork



        Fillcombo33()


    End Sub

    Private Sub BackgroundWorker5_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork

        Try

            QaSetupMod.connecttemp2()

            sqltemp2 = "SELECT * FROM [Agents] WHERE Supervisor='" & cboSupervisor.Text & " ' "



            Dim cmdtemp As New SqlClient.SqlCommand




            cmdtemp.CommandText = sqltemp2

            cmdtemp.Connection = contemp2



            readertemp2 = cmdtemp.ExecuteReader


            While (readertemp2.Read())


                cboAgentName.Items.Add(readertemp2("AgentName"))

                lblSupervisorEmail.Text = (readertemp2("SuperEmail"))


            End While



            cmdtemp.Dispose()



            Me.Cursor = Cursors.Hand










        Catch ex As Exception



            MsgBox(ex.Message)


        End Try



    End Sub

    Private Sub cboAgentName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAgentName.SelectedIndexChanged


        txtTeamName.Text = "Please wait, Loading.."

        BackgroundWorker6.RunWorkerAsync()


    End Sub

    Private Sub BackgroundWorker6_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker6.DoWork
        Try



            QaSetupMod.connecttemp8()

            '   Me.Cursor = Cursors.WaitCursor

            sqltemp8 = "SELECT * FROM [Agents] WHERE AgentName='" & cboAgentName.Text & " ' "



            Dim cmdtemp As New SqlClient.SqlCommand





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

            ' Me.Cursor = Cursors.WaitCursor

            QaSetupMod.connecttemp8()


            sqltemp8 = "SELECT * FROM [Agents] WHERE AgentName='" & cboSupervisorbox.Text & " ' "



            Dim cmdtemp As New SqlClient.SqlCommand





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

    Private Sub BackgroundWorker7_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker7.RunWorkerCompleted


        'cboAgentName.Text = "Agent Name"



        contemp8.Close()


    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted






        contemp2.Close()




        cboAgentName.Text = "Agent Name"

    End Sub

    Private Sub BackgroundWorker8_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker8.ProgressChanged

        ProgressBar1.Value = e.ProgressPercentage

    End Sub



    Private Sub BackgroundWorker8_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker8.RunWorkerCompleted



        Me.Cursor = Cursors.Hand


        PleaseWait.Hide()


        If MsgBox(cboAgentName.Text & " " & "" & "scored a total of" & " " & lblQAScore1.Text & " " & "points on this QA audit, would you like to start a new ‘Call’ audit?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then


            reset()

            buttonEnables()


            Form2.Clear()


            Form2.Show()

            Me.Hide()


        Else

            buttonEnables()

            reset()

            '  cboTeamName.Text = "Team Name"

        End If



    End Sub

    Private Sub BackgroundWorker9_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker9.DoWork

        Try

            For i = 0 To 100

                System.Threading.Thread.Sleep(60)
                Me.BackgroundWorker1.ReportProgress(i)

                lblprogr.Text = i.ToString

                i = i
            Next




            ''
            Store2()


            ' Send to Excell
            '  QAExcell()



            'StoreCallThread = New System.Threading.Thread(AddressOf Store2)

            'StoreCallThread.Start()




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub

    Private Sub BackgroundWorker9_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker9.ProgressChanged


        ProgressBar1.Value = e.ProgressPercentage

    End Sub

    Private Sub BackgroundWorker9_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker9.RunWorkerCompleted


        Me.Cursor = Cursors.Hand
        PleaseWait.Hide()


        If MsgBox(cboSupervisorbox.Text & " " & "" & "scored a total of" & " " & lblQAScore1.Text & " " & "points on this QA audit, would you like to start a new ‘Call’ audit?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then


            reset()

            buttonEnables()


            Form2.Clear()


            Form2.Show()

            Me.Hide()


        Else


            buttonEnables()

            reset()

            'cboSupervisor.Text = "Supervisor"
            'cboAgentName.Text = "Agent Name"
            cboSupervisorbox.Text = "Agent Name"
            'cboSupervisorbox.Text = "Agent Name"

        End If



    End Sub

    Private Sub cboAutoFail_CheckedChanged(sender As Object, e As EventArgs) Handles cboAutoFail.CheckedChanged


        If cboAutoFail.Checked = True Then




            cboAF.Visible = True

            lblQAScore1.Visible = False


            lblAutoZero.Visible = True



        ElseIf cboAutoFail.Checked = False Then


            cboAF.Visible = False

            cboAF.Text = "N/a"

            lblAutoZero.Visible = False

            '     lblQAScore1.Visible = True

            txtQAScore.Text = lblQAScore1.Text


            If lblQAScore1.Text = 0 Then


                lblQAScore1.Text = "100"


            End If

        End If







    End Sub



    Private Sub cbo1_1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo1_1.SelectionChangeCommitted


        Try


            If cbo1_1.SelectedItem = 0 Then


                txt1_1.BackColor = Color.Yellow


                totalQA = Convert.ToInt32(lblQAScore1.Text) - 2



                lblQAScore1.Text = totalQA
                txtQAScore.Text = totalQA


            ElseIf cbo1_1.SelectedItem = 2 Then


                txt1_1.BackColor = Color.White


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else



                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 2


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA


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

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 1


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

            ElseIf cbo1_2.SelectedItem = 1 Then



                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    txt1_2.BackColor = Color.White


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 1


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA


                End If

            End If



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub

    Private Sub cbo1_3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo1_3.SelectionChangeCommitted

        Try

            If cbo1_3.SelectedItem = 0 Then



                txt1_3.BackColor = Color.Yellow

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 2


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA


            ElseIf cbo1_3.SelectedItem = 2 Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    txt1_3.BackColor = Color.White

                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 2


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA


                End If

            End If



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub



    Private Sub cbo2_1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo2_1.SelectionChangeCommitted

        Try


            If cbo2_1.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 15


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA


                txt2_1.BackColor = Color.Yellow


            ElseIf cbo2_1.SelectedItem = "15" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 15


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

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 1


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_2.BackColor = Color.Yellow

            ElseIf cbo3_2.SelectedItem = "1" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 1


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

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 3


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_3.BackColor = Color.Yellow

            ElseIf cbo3_3.SelectedItem = "3" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 3


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

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 4


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_4.BackColor = Color.Yellow

            ElseIf cbo3_4.SelectedItem = "4" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else

                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 4


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

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 3


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_5.BackColor = Color.Yellow

            ElseIf cbo3_5.SelectedItem = "3" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else

                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 3


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

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 4


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt3_7.BackColor = Color.Yellow

            ElseIf cbo3_7.SelectedItem = "4" Then


                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 4


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt3_7.BackColor = Color.White

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

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 15


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt6_1.BackColor = Color.Yellow

            ElseIf cbo6_1.SelectedItem = "15" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else


                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 15


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt6_1.BackColor = Color.White


                End If

            End If



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub




    Private Sub cbo7_1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cbo7_1.SelectionChangeCommitted

        Try

            If cbo7_1.SelectedItem = "0" Then

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 5


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt7_1.BackColor = Color.Yellow

            ElseIf cbo7_1.SelectedItem = "5" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else



                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 5


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

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 5


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt7_2.BackColor = Color.Yellow

            ElseIf cbo7_2.SelectedItem = "5" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else



                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 5


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

                totalQA = Convert.ToInt32(lblQAScore1.Text) - 5


                txtQAScore.Text = totalQA
                lblQAScore1.Text = totalQA

                txt7_3.BackColor = Color.Yellow

            ElseIf cbo7_3.SelectedItem = "5" Then

                If lblQAScore1.Text = "100" Then

                    lblQAScore1.Text = "100"


                Else

                    totalQA = Convert.ToInt32(lblQAScore1.Text) + 5


                    txtQAScore.Text = totalQA
                    lblQAScore1.Text = totalQA

                    txt7_3.BackColor = Color.White

                End If

            End If



        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub




    Private Sub cbo1_1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo1_1.SelectedIndexChanged

        'If cbo1_1.SelectedItem = 0 Then


        '    txt1_1.BackColor = Color.Yellow


        '    totalQA = Convert.ToInt32(lblQAScore1.Text) - 2



        '    lblQAScore1.Text = totalQA



        'ElseIf cbo1_1.SelectedItem = 2 Then


        '    txt1_1.BackColor = Color.White


        '    If lblQAScore1.Text = "100" Then

        '        lblQAScore1.Text = "100"


        '    Else



        '        totalQA = Convert.ToInt32(lblQAScore1.Text) + 2



        '        lblQAScore1.Text = totalQA


        '    End If

        'End If


    End Sub

    Private Sub lblQAScore1_Click(sender As Object, e As EventArgs) Handles lblQAScore1.Click

        QaTotalScore()

    End Sub

    Private Sub cbo1_2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo1_2.SelectedIndexChanged

    End Sub

    Private Sub btnSave2_Click(sender As Object, e As EventArgs) Handles btnSave2.Click


        Try

            MissedWeightsCalc()

            'If txtSR.Text <> "1-" And txtSR.MaskFull = False Then


            '    MsgBox("Please enter a valid SR#")


            'Else


            If txtContactID.Text = "" Then

                MsgBox("A Contact ID is required before saving", MessageBoxButtons.RetryCancel)


            Else


                If cboSRType.Text = "" Then

                    MsgBox("A SR Type must be selected before saving", MessageBoxButtons.RetryCancel)
                    Me.ActiveControl = cboSRType

                Else

                    If cboCallerType.Text = "" Then

                        MsgBox("A Caller Type must be selected before saving", MessageBoxButtons.RetryCancel)

                        Me.ActiveControl = cboCallerType

                    Else


                        If cboSupervisorbox.Text = "Agent Name" Then


                            MsgBox("Please be advised you must select an 'agent name' before proceeding", MessageBoxButtons.RetryCancel)

                            Me.ActiveControl = cboSupervisorbox


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




                                        If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then




                                        Else
                                            SplashScreenManager1.ShowWaitForm()

                                            Me.Cursor = Cursors.WaitCursor

                                            buttondisables()

                                            QACalldisableControls()


                                            Me.ActiveControl = txtContactID


                                            BackgroundWorker2.RunWorkerAsync()

                                            Store2()


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







        Catch ex As Exception



            MsgBox(ex.Message)

        End Try










    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

        Try
            Process.Start("P:\QA Application\QA1\CallD.docx")



        Catch ex As Exception


            MsgBox("Make sure your are connected to the P drive.")
            '   MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub BackgroundWorker6_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted







        contemp8.Close()



    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted




        contemp3.Close()


    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted





        contemp9.Close()




    End Sub

    Private Sub Saver_Tick(sender As Object, e As EventArgs) Handles Saver.Tick


        Saver.Enabled = False
        SplashScreenManager1.CloseWaitForm()


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


                ProgressBar1.Value = 0

                SplashScreenManager2.ShowWaitForm()

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



        Saver2.Enabled = False
        SplashScreenManager1.CloseWaitForm()


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

                QACallEnable()

                buttonEnables()

                reset()

            End If



        Catch ex As Exception



            MsgBox(ex.Message)


            ExcelSaver.Enabled = False

            ProgressBar1.Value = 0
            lblprogr.Text = 0


            QACallEnable()

            buttonEnables()

            SplashScreenManager1.CloseWaitForm()

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

                QACallEnable()

                buttonEnables()


                reset()

            End If



        Catch ex As Exception



            MsgBox(ex.Message)


            ExcelSaver2.Enabled = False


            ProgressBar1.Value = 0
            lblprogr.Text = 0


            QACallEnable()

            buttonEnables()

            SplashScreenManager1.CloseWaitForm()


        End Try



    End Sub

    Private Sub QACallScorecard_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown


        'If e.Control And e.KeyCode.ToString = "S" Then

        '    MsgBox("GoODi")

        'End If







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

                                If cboCallerType.Text = "" Then

                                    MsgBox("A Caller Type must be selected before saving", MessageBoxButtons.RetryCancel)

                                    Me.ActiveControl = cboCallerType

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


                                                QaTotalScore()

                                                TCXscore()

                                                lblTCXscore.Visible = False


                                                If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then


                                                    SplashScreenManager1.ShowWaitForm()
                                                    Me.Cursor = Cursors.WaitCursor

                                                    buttondisables()

                                                    QACalldisableControls()



                                                    Me.ActiveControl = txtSR




                                                    BackgroundWorker1.RunWorkerAsync()

                                                    'PleaseWait.ShowDialog()

                                                    Store()

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

                            'If dtpCondate.Value = Today Then




                            '    MsgBox("Are you sure the Contact date for this Audit is Today?", MessageBoxButtons.RetryCancel)


                            ' Else

                            If cboSRType.Text = "" Then

                                MsgBox("A SR Type must be selected before saving", MessageBoxButtons.RetryCancel)
                                Me.ActiveControl = cboSRType

                            Else

                                If cboCallerType.Text = "" Then

                                    MsgBox("A Caller Type must be selected before saving", MessageBoxButtons.RetryCancel)

                                    Me.ActiveControl = cboCallerType

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


                                                QaTotalScore()

                                                TCXscore()

                                                lblTCXscore.Visible = False




                                                If MsgBox("Are you sure you want to save the Scorecard?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then




                                                Else

                                                    SplashScreenManager1.ShowWaitForm()
                                                    Me.Cursor = Cursors.WaitCursor

                                                    buttondisables()

                                                    QACalldisableControls()


                                                    Me.ActiveControl = txtSR


                                                    BackgroundWorker2.RunWorkerAsync()

                                                    Store2()


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

    Private Sub BackgroundWorker10_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker10.DoWork



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

    Private Sub BackgroundWorker10_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker10.ProgressChanged


        ProgressBar1.Value = e.ProgressPercentage





    End Sub

    Private Sub BackgroundWorker10_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker10.RunWorkerCompleted


        Me.Cursor = Cursors.Hand




    End Sub

    Private Sub SpellCheckLoadTimer_Tick(sender As Object, e As EventArgs) Handles SpellCheckLoadTimer.Tick



        SpellChecker2.ParentContainer = Me
        SpellChecker2.CheckAsYouTypeOptions.CheckControlsInParentContainer = True
        SpellChecker2.SpellCheckMode = SpellCheckMode.AsYouType




        SpellCheckLoadTimer.Enabled = False

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        ' DictLoad()

        ' System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceNames()

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




        'Me.Cursor = Cursors.Hand




        'Dim msg = "The scorecard was successfully emailed to the agent, would you like to audit a new call?"

        'Dim title = "FADV QA Application"

        'Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

        'Dim responce = MsgBox(msg, style, title)





        'If responce = MsgBoxResult.Yes Then


        '    buttonEnables()

        '    reset()



        'Else




        '    reset()

        '    buttonEnables()


        '    Form2.Clear()


        '    Form2.Show()

        '    Me.Hide()



        'End If






    End Sub

    Private Sub SenderEmail1_Tick(sender As Object, e As EventArgs) Handles SenderEmail1.Tick




        SendEmail()




        SenderEmail1.Enabled = False




    End Sub

    Private Sub SenderEmail2_Tick(sender As Object, e As EventArgs) Handles SenderEmail2.Tick





        SendEmail2()




        SenderEmail2.Enabled = False



    End Sub

    Private Sub SendEmailFin_Tick(sender As Object, e As EventArgs) Handles SendEmailFin.Tick


        SplashScreenManager2.CloseWaitForm()

        SendEmailFin.Enabled = False

        Me.Cursor = Cursors.Hand




        Dim msg = "The scorecard was successfully emailed to the agent, would you like to audit a new call?"

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

    Public Sub FillCallerType()

        Try

            QaSetupMod.connecttemp18()

            sqltemp18 = "SELECT * FROM [CallerTypeWOTC]"



            Dim cmdtemp As New SqlClient.SqlCommand




            cmdtemp.CommandText = sqltemp18

            cmdtemp.Connection = contemp18



            readertemp18 = cmdtemp.ExecuteReader


            While (readertemp18.Read())


                cboCallerType.Items.Add(readertemp18("CallerType"))


            End While



            cmdtemp.Dispose()

            contemp18.Close()

            Me.Cursor = Cursors.Hand




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try








    End Sub


    Public Sub FillAutoFail()

        Try

            QaSetupMod.connecttemp17()

            sqltemp17 = "SELECT * FROM [WOTCAutoFail]"



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

    Private Sub txtAccountNum_TextChanged(sender As Object, e As EventArgs) Handles txtAccountNum.TextChanged

    End Sub
End Class