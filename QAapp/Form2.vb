

Imports System.Data.OleDb
Imports System.Data.SqlClient

Imports System.Threading

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms.VisualStyles
Imports System.Collections
Imports System.Reflection

Imports DataGridViewAutoFilter
Imports System.IO
Imports DevExpress.Spreadsheet

Imports DevExpress

Imports Microsoft.Azure
Imports Microsoft.WindowsAzure.CloudStorageAccount
Imports Microsoft.WindowsAzure.StorageClient.BlobType


Imports Microsoft.Office.Interop

'Imports Microsoft.Office.Interop.Access

Imports Microsoft.Office.Interop.Excel
Imports DevExpress.Data
Imports DevExpress.XtraGrid
Imports System.Globalization
Imports System.Net
Imports DevExpress.DataAccess
Imports Microsoft.WindowsAzure
Imports Microsoft.WindowsAzure.StorageClient

Public Class Form2

    Dim rowCount As Integer

    Dim nametest As String


    Dim PendingTotal As Integer = 0
    Dim counter As Integer
    Dim counter2 As Integer
    Dim counter3 As Integer


    Dim PendingTotalsql As Integer = 0
    Dim countersql As Integer
    Dim counter2sql As Integer
    Dim counter3sql As Integer




    Dim Importname As String
    Dim Importteam As String
    Dim Importaudtype As String



    Public Shared AgentEmail As String
    Public Shared UserEmail As String




    Dim temp As Integer

    Dim Now = DateTime.Now


    Dim Desk = My.Computer.FileSystem.SpecialDirectories.Desktop


    Dim Overdue As Integer

    Dim ProgramDateForamt As String = "MM/dd/yyyy"

    Dim en As New CultureInfo("en-US")


    Dim exeDir1 As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)
    Dim QATrendspath1 = IO.Path.Combine(exeDir1.DirectoryName, "QATrendsM.xlsm")


    ' Dim QATrendspath1 = IO.Path.Combine(exeDir1.DirectoryName, "TestBook.xlsx")

    'Dim xlsApp As Excel.Application
    'Dim xlsWorkBooks As Excel.Workbooks = QATrendspath1


    Dim ChangerVariable As String






    Public Sub RowCounting()


        counter = 0
        counter2 = 0
        counter3 = 0





        rowCount = DataGridView2.RowCount



        lblTotalAudits.Text = DataGridView2.RowCount.ToString


        ''Loop For Showing the Pending review






        For i = 0 To (DataGridView2.Rows.Count - 1)



            Dim qad = CDate(DataGridView2.Rows(i).Cells("QA_Date").Value.ToString)

            Dim expiredate = qad.AddDays(7)

            Dim Threedaynoti = qad.AddDays(3)



            If Now > expiredate And DataGridView2.Rows(i).Cells("Rev_Date").Value = "9/9/2020" Then

                counter3 += 1


            End If



            '' Total pending Reivew


            If DataGridView2.Rows(i).Cells("Rev_Date").Value = "9/9/2020" Then



                Int32.TryParse(DataGridView2.Rows(i).Cells("Rev_Date").Value.ToString(), temp)

                PendingTotal += temp

                counter += 1





                '' Total Reviewed


            ElseIf qad < Now Then



                counter2 += 1


                '' To


                ' ElseIf Now > expiredate And DataGridView2.Rows(i).Cells("Rev_Date").Value = "9/9/2020" Then

            End If



        Next




        lblPenReview.Text = counter

        lblPassdue.Text = counter3
        lblTotalRev.Text = counter2
        lblTotalAudits.Text = rowCount

    End Sub



    Public Sub resetcounter()

        counter = 0
        counter2 = 0
        counter3 = 0



    End Sub


    Public Sub SuperVSetup()

        cboSupervisor.Visible = False

        cboAgentName.Visible = False

        cboSuperAgentBox.Visible = True



        Me.cboSuperAgentBox.Location = New System.Drawing.Point(14, 65)

        Me.cboContactType.Location = New System.Drawing.Point(14, 115)





    End Sub


    Public Sub WeekNumber()

        Dim datenow = DateTime.Now
        Dim Dfi = DateTimeFormatInfo.CurrentInfo
        Dim calander = Dfi.Calendar


        '   Dim weekofYear = calander.GetWeekOfYear(datenow, Dfi.CalendarWeekRule, Dfi.FirstDayOfWeek) - 1
        Dim weekofYear = calander.GetWeekOfYear(datenow, Dfi.CalendarWeekRule, Dfi.FirstDayOfWeek)
        Dim QAYear = calander.GetYear(datenow)
        Dim QAMonthNumber = calander.GetMonth(datenow)
        Dim QAMonth = MonthName(QAMonthNumber, False)


        lblMonth.Text = QAMonth

        lblWeekNumber.Text = weekofYear.ToString("D2")

        lblYear.Text = QAYear




    End Sub

    Public Sub LogInDecider()






        '' If lblDeciderDash.Text = "1" Then

        If lblDeciderDash.Text = "QaAuditor" Then



            ''Hides agent combo box
            cboSuperAgentBox.Visible = False
            ''Hides Error message
            Control.CheckForIllegalCrossThreadCalls = False

            ''Loads all Datafrom Dataset
            Me.QaMainDBTableAdapter7.Fill(Me.QADBDataSet6.QAMainDB)
            ''Fills the dropdownbox with list of supervisor names
            Fillcombo()

            Me.ActiveControl = cboContactType

            Me.CenterToScreen()


            ''Sets datte and time format for date time pickers
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "MM/dd/yyyy"

            DateTimePicker2.Format = DateTimePickerFormat.Custom
            DateTimePicker2.CustomFormat = "MM/dd/yyyy"

            DateTimePicker3.Format = DateTimePickerFormat.Custom
            DateTimePicker3.CustomFormat = "MM/dd/yyyy"




        ElseIf lblDeciderDash.Text = "Supervisor" Then

            Control.CheckForIllegalCrossThreadCalls = False


            Me.QaMainDBTableAdapter7.Fill(Me.QADBDataSet6.QAMainDB)


            ''Fill DataGrid based on user 
            QAMainDBBindingSource1.Filter = "[Supervisor] Like '%" & lblconnectedsupervisor1.Text & "%'"

            'QAMainDBBindingSource1.Filter = "[Supervisor] Like '%" & lblconnectedsupervisor2.Text & "%'"

            'QAMainDBBindingSource1.Filter = "[Supervisor] Like '%" & lblconnectedsupervisor3.Text & "%'"


            ' FillSpreadSheet()


            RowCounting()


            cboSupervisor.Visible = False

            cboAgentName.Visible = False

            cboSuperAgentBox.Visible = True


            Me.cboContactType.Location = New System.Drawing.Point(14, 115)


            Fillcombo33()

            Me.CenterToScreen()

            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "MM/dd/yyyy"

            DateTimePicker2.Format = DateTimePickerFormat.Custom
            DateTimePicker2.CustomFormat = "MM/dd/yyyy"

            DateTimePicker3.Format = DateTimePickerFormat.Custom
            DateTimePicker3.CustomFormat = "MM/dd/yyyy"



            RowCounting()


        ElseIf lblDeciderDash.Text = "TeamLead" Then

            Control.CheckForIllegalCrossThreadCalls = False


            Me.QaMainDBTableAdapter7.Fill(Me.QADBDataSet6.QAMainDB)


            ''Fill DataGrid based on user 
            QAMainDBBindingSource1.Filter = "[Supervisor] Like '%" & lblconnectedsupervisor1.Text & "%'"

            'QAMainDBBindingSource1.Filter = "[Supervisor] Like '%" & lblconnectedsupervisor2.Text & "%'"

            'QAMainDBBindingSource1.Filter = "[Supervisor] Like '%" & lblconnectedsupervisor3.Text & "%'"


            ' FillSpreadSheet()


            RowCounting()


            cboSupervisor.Visible = False

            cboAgentName.Visible = False

            cboSuperAgentBox.Visible = True


            Me.cboContactType.Location = New System.Drawing.Point(14, 115)


            Fillcombo33()

            Me.CenterToScreen()

            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "MM/dd/yyyy"

            DateTimePicker2.Format = DateTimePickerFormat.Custom
            DateTimePicker2.CustomFormat = "MM/dd/yyyy"

            DateTimePicker3.Format = DateTimePickerFormat.Custom
            DateTimePicker3.CustomFormat = "MM/dd/yyyy"



            RowCounting()

        ElseIf lblDeciderDash.Text = "GOCSupervisor" Then

            Control.CheckForIllegalCrossThreadCalls = False


            Me.QaMainDBTableAdapter7.Fill(Me.QADBDataSet6.QAMainDB)


            ''Fill DataGrid based on user 
            QAMainDBBindingSource1.Filter = "[Supervisor] Like '%" & lblconnectedsupervisor1.Text & "%'"


            ' FillSpreadSheet()


            RowCounting()


            cboSupervisor.Visible = False

            cboAgentName.Visible = False

            cboSuperAgentBox.Visible = True


            Me.cboContactType.Location = New System.Drawing.Point(14, 115)


            ''Fill drop down box by Location of agent
            Fillcombo34()

            Me.CenterToScreen()

            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "MM/dd/yyyy"

            DateTimePicker2.Format = DateTimePickerFormat.Custom
            DateTimePicker2.CustomFormat = "MM/dd/yyyy"

            DateTimePicker3.Format = DateTimePickerFormat.Custom
            DateTimePicker3.CustomFormat = "MM/dd/yyyy"



            RowCounting()




        ElseIf lblDeciderDash.Text = "GOCTemLead" Then

            Control.CheckForIllegalCrossThreadCalls = False


            Me.QaMainDBTableAdapter7.Fill(Me.QADBDataSet6.QAMainDB)


            ''Fill DataGrid based on user 
            QAMainDBBindingSource1.Filter = "[Supervisor] Like '%" & lblconnectedsupervisor1.Text & "%'"


            ' FillSpreadSheet()


            RowCounting()


            cboSupervisor.Visible = False

            cboAgentName.Visible = False

            cboSuperAgentBox.Visible = True


            Me.cboContactType.Location = New System.Drawing.Point(14, 115)

            ''Fill drop down box by Location of agent
            Fillcombo34()

            Me.CenterToScreen()

            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "MM/dd/yyyy"

            DateTimePicker2.Format = DateTimePickerFormat.Custom
            DateTimePicker2.CustomFormat = "MM/dd/yyyy"

            DateTimePicker3.Format = DateTimePickerFormat.Custom
            DateTimePicker3.CustomFormat = "MM/dd/yyyy"



            RowCounting()


        ElseIf lblDeciderDash.Text = "Admin" Then



            ''Hides agent combo box
            cboSuperAgentBox.Visible = False
            ''Hides Error message
            Control.CheckForIllegalCrossThreadCalls = False

            ''Loads all Datafrom Dataset
            Me.QaMainDBTableAdapter7.Fill(Me.QADBDataSet6.QAMainDB)
            ''Fills the dropdownbox with list of supervisor names
            Fillcombo()

            Me.ActiveControl = cboContactType

            Me.CenterToScreen()


            ''Sets datte and time format for date time pickers
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "MM/dd/yyyy"

            DateTimePicker2.Format = DateTimePickerFormat.Custom
            DateTimePicker2.CustomFormat = "MM/dd/yyyy"

            DateTimePicker3.Format = DateTimePickerFormat.Custom
            DateTimePicker3.CustomFormat = "MM/dd/yyyy"





        End If






    End Sub



    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            '  splashScreenManager1.CloseWaitForm()

            ''Annoucment tmer
            '   Timer3.Enabled = True

            ''Reads the number to change the annoument page
            ' Changer()

            Annoucer()

            GridView1.OptionsView.ShowIndicator = False

            ''Sets the culture to the app
            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")

            ' Dim ProgramdateString As String = DateTimePicker1.ToString(ProgramDateForamt)


            ''hide the tab pages
            TabControl1.TabPages.Remove(Tab3)
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.TabPages.Remove(TabPage4)

            ''Sets the week number for the app
            WeekNumber()

            ''
            If lblDeciderDash.Text = "Admin" Or lblDeciderDash.Text = "QaAuditor" Then

                lblDeciderX2.Text = "1"



            ElseIf lblDeciderDash.Text = "TeamLead" Or lblDeciderDash.Text = "Supervisor" Then


                lblDeciderX2.Text = "2"

            End If

            If lblDeciderDash.Text = "Admin" Then

                lnkTransferAudits.Visible = True


            End If

            Me.Cursor = Cursors.Hand


            LogInDecider()


        Catch ex As Exception

            ' RefreshWorkbook()

            MsgBox(ex.Message)



        End Try


    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnOLDSave.Click

        Try

            If cboAgentName.Text = "Agent Name" Or cboContactType.Text = "Contact Type" Then






                MsgBox("Please be advised you must fill out all 'Agent Information' before proceeding", MessageBoxButtons.RetryCancel)


            Else




                If QACallScorecard.lblQAScore1.Visible = True Or QAEmailScorecard.lblQAScore1.Visible = True Or QAChatScorecard.lblQAScore1.Visible = True Or QALvl2CallScorecard.lblQAScore1.Visible = True Or QAlvl2EmailScorecard.lblQAScore.Visible = True Or QAResCallScorecard.lblQAScore1.Visible = True Or QAResidentEmailScorecard.lblQAScore.Visible = True Or QAConsuACallScorcard.lblQAScore1.Visible = True Then



                    MsgBox("You can not save a scorecard that has been scored already, press 'clear fields' button", MessageBoxButtons.OK)



                Else



                    If MsgBox("Are you sure you want to save and continue to the Scorecard?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then


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

                        cboSupervisor.Enabled = False

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
                            ' QACallScorecard.lblAgentName1.Text = cboAgentName.Text
                            '  QACallScorecard.lblAgentTeam1.Text = cboTeamName.Text
                            ' QACallScorecard.lblContactType1.Text = cboContactType.Text
                            ' QACallScorecard.lblSRNumber1.Text = txtSRNumber.Text
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



                            ''Email

                            'QAEmailScorecard.cbo.Text = cboAgentName.Text
                            'QAEmailScorecard.lblAgentTeam.Text = cboTeamName.Text
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






                            '''Chat
                            'QAChatScorecard.lblAgentName1.Text = cboAgentName.Text
                            'QAChatScorecard.lblAgentTeam1.Text = cboTeamName.Text
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
                            QALvl2CallScorecard.lblAgentTeam1.Text = cboSupervisor.Text
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
                            QAlvl2EmailScorecard.lblAgentTeam.Text = cboSupervisor.Text
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
                            QAResCallScorecard.lblAgentTeam1.Text = cboSupervisor.Text
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
                            QAResidentEmailScorecard.lblAgentTeam.Text = cboSupervisor.Text
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
                            QAConsuACallScorcard.lblAgentTeam1.Text = cboSupervisor.Text
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

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click

        Try

            'If txtSRNumber.Enabled = True Then


            '    MsgBox("Edits can not be made at this time", MessageBoxButtons.OK)


            '   Else



            If QACallScorecard.lblQAScore1.Visible = True Or QAEmailScorecard.lblQAScore1.Visible = True Then



                MsgBox("Edits can not be made to Scorecard Info after Audit has been saved", MessageBoxButtons.OK)



            Else




                If MsgBox("Are you sure you want to make edits to the Scorecard Info?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then




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
                    'txtAgentTeam.Enabled = True
                    cboContactType.Enabled = True

                    cboSupervisor.Enabled = True


                    'Transfer label names to QAscorecard form


                    QACallScorecard.cboSupervisor.Text = cboAgentName.Text
                    QACallScorecard.cboSupervisor.Text = cboSupervisor.Text
                    '   QACallScorecard.lblContactType1.Text = cboContactType.Text
                    QACallScorecard.txtSR.Text = txtSRNumber.Text




                    QAEmailScorecard.cboAgentName.Text = cboAgentName.Text
                    QAEmailScorecard.cboSupervisor.Text = cboSupervisor.Text
                    '  QAEmailScorecard.lblContactType.Text = cboContactType.Text
                    QAEmailScorecard.txtSR.Text = txtSRNumber.Text


                    QAChatScorecard.cboAgentName.Text = cboAgentName.Text
                    QAChatScorecard.cboSupervisor.Text = cboSupervisor.Text
                    '  QAChatScorecard.lblContactType1.Text = cboContactType.Text
                    QAChatScorecard.txtSR.Text = txtSRNumber.Text







                    QALvl2CallScorecard.lblAgentName1.Text = cboAgentName.Text
                    QALvl2CallScorecard.lblAgentTeam1.Text = cboSupervisor.Text
                    QALvl2CallScorecard.lblContactType1.Text = cboContactType.Text
                    QALvl2CallScorecard.lblSRNumber1.Text = txtSRNumber.Text
                    QALvl2CallScorecard.lblUserID.Text = txtUserID.Text
                    QALvl2CallScorecard.lblJIRA.Text = txtJIRAbox.Text




                    QAlvl2EmailScorecard.lblAgentName.Text = cboAgentName.Text
                    QAlvl2EmailScorecard.lblAgentTeam.Text = cboSupervisor.Text
                    QAlvl2EmailScorecard.lblContactType.Text = cboContactType.Text
                    QAlvl2EmailScorecard.lblSRNumber.Text = txtSRNumber.Text
                    QAlvl2EmailScorecard.lblUserID.Text = txtUserID.Text
                    QAlvl2EmailScorecard.lblJIRA.Text = txtJIRAbox.Text




                    QAResCallScorecard.lblAgentName1.Text = cboAgentName.Text
                    QAResCallScorecard.lblAgentTeam1.Text = cboSupervisor.Text
                    QAResCallScorecard.lblContactType1.Text = cboContactType.Text
                    QAResCallScorecard.lblSRNumber1.Text = txtSRNumber.Text


                    QAResidentEmailScorecard.lblAgentName.Text = cboAgentName.Text
                    QAResidentEmailScorecard.lblAgentTeam.Text = cboSupervisor.Text
                    QAResidentEmailScorecard.lblContactID1.Text = cboContactType.Text
                    QAResidentEmailScorecard.lblSRNumber.Text = txtSRNumber.Text


                    QAConsuACallScorcard.lblAgentName1.Text = cboAgentName.Text
                    QAConsuACallScorcard.lblAgentTeam1.Text = cboSupervisor.Text
                    QAConsuACallScorcard.lblContactID1.Text = cboContactType.Text
                    QAConsuACallScorcard.lblSRNumber1.Text = txtSRNumber.Text





                End If

            End If

            ' End If





        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub


    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        Me.Cursor = Cursors.Hand

        Try

            If MsgBox("Please be advised you about to clear and reset the scorecard, do you want to continue?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then



            Else




                readertemp8.Close()

                contemp8.Close()


                readertemp1.Close()
                contemp1.Close()




                txtAgentTeam.Clear()


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

    Private Sub btnHide_Click(sender As Object, e As EventArgs) Handles btnHide.Click

        If cboSupervisor.Enabled = True Then

            MsgBox("You can not hide the dashboard when its the only form open or press 'save info' button first to proceed.")


        Else






            Me.Hide()

        End If






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
            'txtAgentTeam.Enabled = True
            cboContactType.Enabled = True
            cboSupervisor.Enabled = True



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
            ' txtAgentTeam.Clear()


            txtAgentEmail.Clear()




            btnOLDSave.Enabled = True


            cboAgentName.Text = "Agent Name"
            cboSupervisor.Text = "Supervisor"


            cboSuperAgentBox.Text = "Agent Name"

            cboContactType.Text = "Contact Type"


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub




    Public Sub Fillcombo()


        Try

            QaSetupMod.connecttemp16()



            '  sqltemp1 = "SELECT * FROM [Agents] WHERE Supervisor='" & lblQAauditor.Text & "' "


            sqltemp16 = "SELECT * FROM [Supervisor]"


            Dim cmdtemp As New SqlCommand



            '  cmdtemp.CommandText = sqltemp

            cmdtemp.CommandText = sqltemp16

            cmdtemp.Connection = contemp16





            readertemp16 = cmdtemp.ExecuteReader



            While (readertemp16.Read())



                cboSupervisor.Items.Add(readertemp16("FullName"))



            End While






            cmdtemp.Dispose()

            readertemp16.Close()

            contemp16.Close()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try






    End Sub




    Public Sub Fillcombo1()


        Try

            QaSetupMod.connecttemp1()





            sqltemp1 = "SELECT * FROM [Agents] WHERE Supervisor='" & cboSupervisor.Text & "'"





            Dim cmdtemp As New SqlCommand



            '  cmdtemp.CommandText = sqltemp

            cmdtemp.CommandText = sqltemp1

            cmdtemp.Connection = contemp1





            readertemp1 = cmdtemp.ExecuteReader



            While (readertemp1.Read())



                cboAgentName.Items.Add(readertemp1("AgentName"))


                lblSupervisorEmail.Text = (readertemp1("SuperEmail"))

            End While






            cmdtemp.Dispose()
            readertemp1.Close()
            contemp1.Close()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try


    End Sub


    Public Sub Fillcombo33()


        Try


            QaSetupMod.connecttemp6()


            sqltemp6 = "SELECT * FROM [Agents] WHERE Supervisor='" & lblconnectedsupervisor1.Text & "' "





            Dim cmdtemp As New SqlCommand



            '  cmdtemp.CommandText = sqltemp

            cmdtemp.CommandText = sqltemp6

            cmdtemp.Connection = contemp6





            readertemp6 = cmdtemp.ExecuteReader



            While (readertemp6.Read())



                cboSuperAgentBox.Items.Add(readertemp6("AgentName"))



            End While



            cmdtemp.Dispose()

            readertemp6.Close()

            contemp6.Close()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try




    End Sub



    Public Sub Fillcombo34()


        Try


            QaSetupMod.connecttemp6()


            '    sqltemp6 = "SELECT * FROM [Agents] WHERE Location='" & lblQAAuditor3.Text & "' "





            Dim cmdtemp As New SqlCommand



            '  cmdtemp.CommandText = sqltemp

            cmdtemp.CommandText = sqltemp6

            cmdtemp.Connection = contemp6





            readertemp6 = cmdtemp.ExecuteReader



            While (readertemp6.Read())



                cboSuperAgentBox.Items.Add(readertemp6("AgentName"))



            End While



            cmdtemp.Dispose()

            readertemp6.Close()

            contemp6.Close()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try




    End Sub






    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        Try

            If MessageBox.Show("Are you sure to close this application?", "FADV Quality Assurance Application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then

                'xlsApp.Workbooks("Testbook.xlsm").Close(SaveChanges:=False)

                End

            Else
                e.Cancel = True


            End If


        Catch ex As Exception

            MsgBox(ex.Message)

        End Try



    End Sub


    Public Sub resetcombo()

        cboAgentName.Items.Clear()

        cboAgentName.Text = "Agent Name"

        cboSuperAgentBox.Text = "Agent Name"


    End Sub


    Public Sub filling()



        sqltemp1 = "SELECT * FROM [Agents] WHERE Supervisor='" & cboSupervisor.Text & " ' "

        '    sqltemp2 = "SELECT * FROM [Teams]"



        Dim cmdtemp As New SqlClient.SqlCommand





        cmdtemp.CommandText = sqltemp1

        cmdtemp.Connection = contemp1



        readertemp1 = cmdtemp.ExecuteReader


        While (readertemp1.Read())




            cboAgentName.Items.Add(readertemp1("AgentName"))


            '    txtAgentTeam.Text = readertemp1(3).ToString

            'QACallScorecard.txtTeamName.Text = readertemp1(3).ToString
            'QAEmailScorecard.txtTeamName.Text = readertemp1(3).ToString
            'QAChatScorecard.txtTeamName.Text = readertemp1(3).ToString



        End While

        contemp1.Close()



        cmdtemp.Dispose()

        readertemp1.Close()

    End Sub



    Private Sub cboTeamName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSupervisor.SelectedIndexChanged



        Try
            Me.Cursor = Cursors.AppStarting

            ' resetcombo()

            cboAgentName.Items.Clear()

            cboAgentName.Text = "Please wait, Loading.."


            BackgroundWorker4.RunWorkerAsync()





        Catch ex As Exception



            MsgBox(ex.Message)


        End Try



    End Sub




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        QAMainDBBindingSource.Filter = "(Convert(ID, 'System.String') LIKE '" & TextBox1.Text & "')"







    End Sub

    Public Sub search1()



        If RadioButton1.Checked Then



            QAMainDBBindingSource1.Filter = "[Auditor] like '%" & txtSearchBox.Text & "%'"

            RowCounting()

        ElseIf RadioButton2.Checked Then


            QAMainDBBindingSource1.Filter = "[QA-Agent] like '%" & txtSearchBox.Text & "%'"

            RowCounting()

        ElseIf RadioButton3.Checked Then



            QAMainDBBindingSource1.Filter = "[SR] like '%" & txtSearchBox.Text & "%'"


            RowCounting()

        ElseIf RadioButton4.Checked Then

            QAMainDBBindingSource1.Filter = "[Ctype] like '%" & txtSearchBox.Text & "%'"

            RowCounting()

        ElseIf RadioButton5.Checked Then

            QAMainDBBindingSource1.Filter = "[Supervisor] like '%" & txtSearchBox.Text & "%'"

            RowCounting()


        ElseIf RadioButton6.Checked Then

            QAMainDBBindingSource1.Filter = "[Rev-Manager] like '%" & txtSearchBox.Text & "%'"

            RowCounting()



            '  Bindingsource7.Filter = "[Rev-Manager] like '%" & lblQAauditor.Text & "%'" & "[QA-Agent] Like '%" & txtSearchBox.Text & "'"

        End If



    End Sub




    Public Sub searcha1()


        If RadioButton1.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & "' AND [Auditor] like '%" & txtSearchBox.Text & "%'"



            RowCounting()

        ElseIf RadioButton2.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & "' AND [QA-Agent] like '%" & txtSearchBox.Text & "%'"



            RowCounting()

        ElseIf RadioButton3.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " ' AND [SR] like '%" & txtSearchBox.Text & "%'"




            RowCounting()

        ElseIf RadioButton4.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " 'AND [Ctype] like '%" & txtSearchBox.Text & "%'"



            RowCounting()

        ElseIf RadioButton5.Checked Then


            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " 'AND [QA-Team] like '%" & txtSearchBox.Text & "%'"



            RowCounting()


        ElseIf RadioButton6.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " 'AND [Rev-Manager] like '%" & txtSearchBox.Text & "%'"



            RowCounting()




        End If


    End Sub


    Public Sub search2()


        If RadioButton1.Checked Then



            QAMainDBBindingSource1.Filter = "[Auditor] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"



            RowCounting()

        ElseIf RadioButton2.Checked Then



            QAMainDBBindingSource1.Filter = "[QA-Agent] like '%" & txtSearchBox.Text & "%'  AND [Supervisor] ='" & lblQAauditor.Text & "'"

            RowCounting()

        ElseIf RadioButton3.Checked Then

            '  Bindingsource7.Filter = "[Supervisor] ='" & lblQAauditor.Text & "'"

            QAMainDBBindingSource1.Filter = "[Supervisor] ='" & lblQAauditor.Text & "' AND [SR] Like '%" & txtSearchBox.Text & "'"


            RowCounting()

        ElseIf RadioButton4.Checked Then

            QAMainDBBindingSource1.Filter = "[Ctype] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"

            RowCounting()

        ElseIf RadioButton5.Checked Then

            QAMainDBBindingSource1.Filter = "[Supervisor] Like '%" & txtSearchBox.Text & "%'  AND [Supervisor] ='" & lblQAauditor.Text & "'"

            RowCounting()


        ElseIf RadioButton6.Checked Then

            QAMainDBBindingSource1.Filter = "[Rev-Manager] like '%" & txtSearchBox.Text & "%'  AND [Supervisor] ='" & lblQAauditor.Text & "'"

            RowCounting()


        End If











    End Sub


    Public Sub search2a()



        If RadioButton1.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & "' AND [Auditor] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"



            RowCounting()

        ElseIf RadioButton2.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & "' AND [QA-Agent] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"



            RowCounting()

        ElseIf RadioButton3.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " ' AND [SR] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"




            RowCounting()

        ElseIf RadioButton4.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " 'AND [Ctype] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"



            RowCounting()

        ElseIf RadioButton5.Checked Then


            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " 'AND [QA-Team] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"



            RowCounting()


        ElseIf RadioButton6.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " 'AND [Rev-Manager] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"



            RowCounting()


        End If



    End Sub




    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Try


            If txtSearchBox.Text = "" Then

                MsgBox("Please Fill out the Search box", MessageBoxButtons.RetryCancel)

                Me.ActiveControl = txtSearchBox


                Me.Cursor = Cursors.Hand
            Else


                If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False And RadioButton6.Checked = False Then

                    MsgBox("Please select an option under ‘search criteria’", MessageBoxButtons.RetryCancel)


                    Me.Cursor = Cursors.Hand


                Else

                    '  If lblQAauditor.Text = "Carla Hardy" Or lblQAauditor.Text = "Daphne Nixon" Or lblQAauditor.Text = "Debolina Chatterjee" Or lblQAauditor.Text = "Gnanesh Jayaram" Or lblQAauditor.Text = "Nathan Beers" Or lblQAauditor.Text = "Neha Modak" Or lblQAauditor.Text = "Barb Hurley" Or lblQAauditor.Text = "Daniel Jones" Or lblQAauditor.Text = "Eric Durrant" Then

                    If lblDeciderDash.Text = "1" Then

                        If MsgBox("Is this search based on the Audit Date?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then

                            search1()





                        ElseIf lblDeciderDash.Text = "2" Then

                            searcha1()




                        End If


                    Else



                        If MsgBox("Is this search based on the Audit Date?", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then



                            search2()



                        Else



                            search2a()


                        End If





                    End If

                End If

            End If

        Catch ex As Exception



            MsgBox(ex.Message)


        End Try



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        RadioButton1.Checked = False

        RadioButton2.Checked = False


        RadioButton3.Checked = False


        RadioButton4.Checked = False

        RadioButton5.Checked = False

        RadioButton6.Checked = False



        RevClear()



    End Sub


    Public Sub RevClear()

        lblQaAvg.Text = 0


        txtSRPort.Clear()

        ProgressBar1.Value = 0

        txtSearchBox.Clear()

        DateTimePicker2.Value = Today
        DateTimePicker3.Value = Today

        lblQaAvg.Text = "0"



    End Sub

    Public Sub refreshDB()

        Try

            Me.QaMainDBTableAdapter7.Fill(Me.QADBDataSet6.QAMainDB)


            '   QAMainDBBindingSource1.Filter = "[Auditor] like '%" & "*" & "%'"
            '







            '  Me.QaMainDBTableAdapter7.Fill(Me.QADBDataSet6.QAMainDB)



            ' GridView1.PopulateColumns(GridControl1.DataSource)


            '   Me.QAMainDBTableAdapter.Fill(Me.QADataSet5.QAMainDB)


            ' Me.QAMainDBTableAdapter7.Fill(Me.QADataSet1.QAMainDB)

            '  QAMainDBBindingSource1.Filter = "[Auditor] like '%" & "*" & "%'"



            'resetcounter()

            'RowCounting()

            Me.Cursor = Cursors.Hand

            SplashScreenManager1.CloseWaitForm()


        Catch ex As Exception

            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)


            Me.Cursor = Cursors.Hand


        End Try


    End Sub

    Public Sub refreshDB2()

        Try


            Me.QaMainDBTableAdapter7.RefreshForSupervisor(Me.QADBDataSet6.QAMainDB)



            QAMainDBBindingSource1.Filter = "[Supervisor] like '%" & lblQAAuditor3.Text & "%'"
            '



            '   Dim SQLRefresh As String = "SELECT * From dbo.QAMainDB Where Supervisor=" + lblQAAuditor3.Text


            'Dim SQLRefresh As String = "SELECT * From dbo.QAMainDB Where Supervisor= [Supervisor]"



            'Dim command As New SqlCommand(SQLRefresh, contemp1b)



            'command.ExecuteNonQuery()




            'Me.QaMainDBTableAdapter7.Fill(Me.QADBDataSet6.QAMainDB)


            'QAMainDBBindingSource1.Filter = "[Supervisor] like '%" & lblQAAuditor3.Text & "%'"




            'resetcounter()

            'RowCounting()

            SplashScreenManager1.CloseWaitForm()

            Me.Cursor = Cursors.Hand

        Catch ex As Exception

            SplashScreenManager1.CloseWaitForm()

            MsgBox(ex.Message)



            Me.Cursor = Cursors.Hand

        End Try






    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


        Try



            PageLoadTimer.Start()

            ProgressBar1.Value = 0

            '  PleaseWait.ShowDialog()

            Me.Cursor = Cursors.AppStarting

            '    If lblQAauditor.Text = "Carla Hardy" Or lblQAauditor.Text = "Daphne Nixon" Or lblQAauditor.Text = "Debolina Chatterjee" Or lblQAauditor.Text = "Gnanesh Jayaram" Or lblQAauditor.Text = "Nathan Beers" Or lblQAauditor.Text = "Neha Modak" Or lblQAauditor.Text = "Barb Hurley" Or lblQAauditor.Text = "Daniel Jones" Or lblQAauditor.Text = "Eric Durrant" Or lblQAauditor.Text = "Nick DiVincenzo" Or lblQAauditor.Text = "Isiah Topa" Then

            If lblDeciderDash.Text = "1" Then

                refreshDB()

                lblQaAvg.Text = 0

                DateTimePicker2.Value = Today
                DateTimePicker3.Value = Today

            ElseIf lblDeciderDash.Text = "2" Then

                refreshDB2()

                lblQaAvg.Text = 0

                DateTimePicker2.Value = Today
                DateTimePicker3.Value = Today




            End If


            txtSRPort.Clear()
            txtSearchBox.Clear()

            RadioButton1.Checked = False

            RadioButton2.Checked = False


            RadioButton3.Checked = False


            RadioButton4.Checked = False

            RadioButton5.Checked = False

            RadioButton6.Checked = False






        Catch ex As Exception



            MsgBox(ex.Message)


        End Try





    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        'txtSRNumber.Text = "10-12323233"
        'txtContactID.Text = "1212111"
        'txtContactName.Text = "Crystal Smith"
        'txtContactEmail.Text = "CrystalSmith@Gmail.com"
        'txtContactPhone.Text = "5558889695"
        'txtAccountNum.Text = "abc32323"
        'txtCompany.Text = "Little Leauge"
        'txtOrderID.Text = "95955555"
        'txtJIRAbox.Text = "45488788"
        'txtUserID.Text = "545454user"

        TabControl1.TabPages.Insert(0, Tab3)





    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        'txtSRNumber.Text = "13-12323233"
        'txtContactID.Text = "7825984"
        'txtContactName.Text = "Heather Brown"
        'txtContactEmail.Text = "HB@spcglobal.com"
        'txtContactPhone.Text = "5559999695"
        'txtAccountNum.Text = "bbc701222"
        'txtCompany.Text = "Hannford"
        'txtOrderID.Text = "11144778"
        'txtJIRAbox.Text = "258488788"
        'txtUserID.Text = "788454user3"



        Timer3.Enabled = True





    End Sub

    Public Sub FiltDate()


        If RadioButton1.Checked = True Then


            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <='" & DateTimePicker3.Text & "' AND [Auditor] like '%" & txtSearchBox.Text & "%'"

            RowCounting()

        ElseIf RadioButton2.Checked = True Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & "' AND [QA-Agent] like '%" & txtSearchBox.Text & "%'"

            RowCounting()

        ElseIf RadioButton3.Checked = True Then


            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " ' AND [SR] like '%" & txtSearchBox.Text & "%'"


            RowCounting()

        ElseIf RadioButton4.Checked = True Then


            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " ' AND [Ctype] like '%" & txtSearchBox.Text & "%'"


            RowCounting()

        ElseIf RadioButton5.Checked = True Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " 'AND [QA-Team] like '%" & txtSearchBox.Text & "%'"

            RowCounting()


        ElseIf RadioButton6.Checked = True Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " 'AND [Rev-Manager] like '%" & txtSearchBox.Text & "%'"





        ElseIf RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False And RadioButton6.Checked = False Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " '"

            RowCounting()


        End If










    End Sub

    Public Sub filtdate2()





        If RadioButton1.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & "' AND [Auditor] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"



            RowCounting()

        ElseIf RadioButton2.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & "' AND [QA-Agent] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"



            RowCounting()

        ElseIf RadioButton3.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " ' AND [SR] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"




            RowCounting()

        ElseIf RadioButton4.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " 'AND [Ctype] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"



            RowCounting()

        ElseIf RadioButton5.Checked Then


            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " 'AND [QA-Team] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"



            RowCounting()


        ElseIf RadioButton6.Checked Then

            QAMainDBBindingSource1.Filter = "[QA-Date] >= '" & DateTimePicker2.Text & "' AND [QA-Date] <= '" & DateTimePicker3.Text & " 'AND [Rev-Manager] like '%" & txtSearchBox.Text & "%' AND [Supervisor] ='" & lblQAauditor.Text & "'"



            RowCounting()


        End If





    End Sub






    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles txtFilterDate.Click

        Try


            '  If lblQAauditor.Text = "Carla Hardy" Or lblQAauditor.Text = "Daphne Nixon" Or lblQAauditor.Text = "Debolina Chatterjee" Or lblQAauditor.Text = "Gnanesh Jayaram" Or lblQAauditor.Text = "Nathan Beers" Or lblQAauditor.Text = "Neha Modak" Or lblQAauditor.Text = "Barb Hurley" Or lblQAauditor.Text = "Daniel Jones" Or lblQAauditor.Text = "Eric Durrant" Then

            If lblDeciderDash.Text = "1" Then

                FiltDate()


            ElseIf lblDeciderDash.Text = "2" Then




                filtdate2()




            End If




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try



    End Sub

    Private Sub ExitProgramToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitProgramToolStripMenuItem.Click


        Try

            If MessageBox.Show("Are you sure to close this application?", "FADV Quality Assurance Application", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then

                End

            Else



            End If



        Catch ex As Exception

            MsgBox(ex.Message)

        End Try


    End Sub


    Public Sub ScorecardLaunch()





        Dim ProgramAsDate As DateTime = DateTimePicker1.Value

        Dim ProgramString As String = ProgramAsDate.ToString(ProgramDateForamt)







        ' Select Scorcard based on Qa Audit Type

        If cboContactType.Text = "Call" Then


            Dim stringg As String = DateTimePicker1.Text


            '' CALLGost fields
            QACallScorecard.txtgnamebox.Text = txtContactName.Text
            QACallScorecard.txtgphone.Text = txtContactPhone.Text
            QACallScorecard.txtgemail.Text = txtContactEmail.Text
            QACallScorecard.txtgacc.Text = txtAccountNum.Text
            QACallScorecard.txtgcompany.Text = txtCompany.Text
            QACallScorecard.txtgorderid.Text = txtOrderID.Text
            QACallScorecard.txtgDatebox.Text = stringg

            QACallScorecard.lbldrive2.Text = lblMDrive.Text
            QACallScorecard.lblSCRN.Text = lblSCRN.Text

            ''Callt visable fields

            QACallScorecard.cboAgentName.Text = cboAgentName.Text
            QACallScorecard.cboSupervisor.Text = cboSupervisor.Text



            QACallScorecard.lblContactType1.Text = cboContactType.Text
            QACallScorecard.txtSR.Text = txtSRNumber.Text
            QACallScorecard.txtContactID.Text = txtContactID.Text


            QACallScorecard.txtContactName.Text = txtContactName.Text
            QACallScorecard.txtContactEmail.Text = txtContactEmail.Text
            QACallScorecard.txtContactPhone.Text = txtContactPhone.Text
            QACallScorecard.txtAccountNum.Text = txtAccountNum.Text
            QACallScorecard.txtCompany.Text = txtCompany.Text
            QACallScorecard.txtOrderID.Text = txtOrderID.Text


            'QACallScorecard.dtpCondate.Text = stringg

            QACallScorecard.dtpCondate.Text = ProgramAsDate.ToString(ProgramDateForamt)



            QACallScorecard.lblQAauditor1.Text = lblQAauditor.Text


            QACallScorecard.lblWeekNumber.Text = lblWeekNumber.Text


            QAWOTCInboundScorecard.Hide()
            QAEmailScorecard.Hide()
            QAChatScorecard.Hide()
            QALvl2CallScorecard.Hide()
            QAlvl2EmailScorecard.Hide()
            QAResCallScorecard.Hide()
            QAResidentEmailScorecard.Hide()
            QAConsuACallScorcard.Hide()


            Me.Cursor = Cursors.Hand


            Me.Hide()
            QACallScorecard.Show()

        ElseIf cboContactType.Text = "Email" Then

            Dim stringg As String = DateTimePicker1.Text


            ''Email Gost fields
            QAEmailScorecard.txtgnamebox.Text = txtContactName.Text
            QAEmailScorecard.txtgphone.Text = txtContactPhone.Text
            QAEmailScorecard.txtgemail.Text = txtContactEmail.Text
            QAEmailScorecard.txtgacc.Text = txtAccountNum.Text
            QAEmailScorecard.txtgcompany.Text = txtCompany.Text
            QAEmailScorecard.txtgorderid.Text = txtOrderID.Text
            QAEmailScorecard.txtgDatebox.Text = stringg

            QAEmailScorecard.lbldrive2.Text = lblMDrive.Text
            QAEmailScorecard.lblSCRN.Text = lblSCRN.Text

            ''eMAIL visable fields

            QAEmailScorecard.cboAgentName.Text = cboAgentName.Text
            QAEmailScorecard.cboSupervisor.Text = cboSupervisor.Text




            QAEmailScorecard.lblContactType.Text = cboContactType.Text
            QAEmailScorecard.txtSR.Text = txtSRNumber.Text
            QAEmailScorecard.txtContactID.Text = txtContactID.Text


            QAEmailScorecard.txtContactName.Text = txtContactName.Text
            QAEmailScorecard.txtContactEmail.Text = txtContactEmail.Text
            QAEmailScorecard.txtContactPhone.Text = txtContactPhone.Text
            QAEmailScorecard.txtAccountNum.Text = txtAccountNum.Text
            QAEmailScorecard.txtCompany.Text = txtCompany.Text
            QAEmailScorecard.txtOrderID.Text = txtOrderID.Text

            QAEmailScorecard.DateTimePicker1.Text = ProgramAsDate.ToString(ProgramDateForamt)

            '  QAEmailScorecard.DateTimePicker1.Text = DateTimePicker1.Text

            QAEmailScorecard.lblQAauditor1.Text = lblQAauditor.Text


            QAEmailScorecard.lblWeekNumber.Text = lblWeekNumber.Text

            QAWOTCInboundScorecard.Hide()
            QACallScorecard.Hide()
            QAChatScorecard.Hide()
            QALvl2CallScorecard.Hide()
            QAlvl2EmailScorecard.Hide()
            QAResCallScorecard.Hide()
            QAResidentEmailScorecard.Hide()
            QAConsuACallScorcard.Hide()


            Me.Cursor = Cursors.Hand

            Me.Hide()


            QAEmailScorecard.Show()


        ElseIf cboContactType.Text = "Chat" Then

            Dim stringg As String = DateTimePicker1.Text



            ''chat Gost fields
            QAChatScorecard.txtgnamebox.Text = txtContactName.Text
            QAChatScorecard.txtgphone.Text = txtContactPhone.Text
            QAChatScorecard.txtgemail.Text = txtContactEmail.Text
            QAChatScorecard.txtgacc.Text = txtAccountNum.Text
            QAChatScorecard.txtgcompany.Text = txtCompany.Text
            QAChatScorecard.txtgorderid.Text = txtOrderID.Text
            QAChatScorecard.txtgDatebox.Text = stringg

            QAChatScorecard.lbldrive2.Text = lblMDrive.Text
            QAChatScorecard.lblSCRN1.Text = lblSCRN.Text





            ''Chat visable fields

            QAChatScorecard.cboAgentName.Text = cboAgentName.Text
            QAChatScorecard.cboSupervisor.Text = cboSupervisor.Text


            QAChatScorecard.lblContactType1.Text = cboContactType.Text
            QAChatScorecard.txtSR.Text = txtSRNumber.Text
            QAChatScorecard.txtContactID.Text = txtContactID.Text


            QAChatScorecard.txtContactName.Text = txtContactName.Text
            QAChatScorecard.txtContactEmail.Text = txtContactEmail.Text
            QAChatScorecard.txtContactPhone.Text = txtContactPhone.Text
            QAChatScorecard.txtAccountNum.Text = txtAccountNum.Text
            QAChatScorecard.txtCompany.Text = txtCompany.Text
            QAChatScorecard.txtOrderID.Text = txtOrderID.Text


            QAChatScorecard.DateTimePicker1.Text = ProgramAsDate.ToString(ProgramDateForamt)


            '  QAChatScorecard.DateTimePicker1.Text = DateTimePicker1.Text

            QAChatScorecard.lblQAauditor1.Text = lblQAauditor.Text


            QAChatScorecard.lblWeekNumber.Text = lblWeekNumber.Text

            QAWOTCInboundScorecard.Hide()
            QACallScorecard.Hide()
            QAEmailScorecard.Hide()
            QALvl2CallScorecard.Hide()
            QAlvl2EmailScorecard.Hide()
            QAResCallScorecard.Hide()
            QAResidentEmailScorecard.Hide()
            QAConsuACallScorcard.Hide()

            Me.Cursor = Cursors.Hand

            Me.Hide()

            QAChatScorecard.Show()



        ElseIf cboContactType.Text = "WOTC Inbound" Then


            Dim stringg As String = DateTimePicker1.Text


            '' CALLGost fields
            QAWOTCInboundScorecard.txtgnamebox.Text = txtContactName.Text
            QAWOTCInboundScorecard.txtgphone.Text = txtContactPhone.Text
            QAWOTCInboundScorecard.txtgemail.Text = txtContactEmail.Text
            QAWOTCInboundScorecard.txtgacc.Text = txtAccountNum.Text
            QAWOTCInboundScorecard.txtgcompany.Text = txtCompany.Text
            QAWOTCInboundScorecard.txtgorderid.Text = txtOrderID.Text
            QAWOTCInboundScorecard.txtgDatebox.Text = stringg

            QAWOTCInboundScorecard.lbldrive2.Text = lblMDrive.Text
            QAWOTCInboundScorecard.lblSCRN.Text = lblSCRN.Text

            ''Callt visable fields

            QAWOTCInboundScorecard.cboAgentName.Text = cboAgentName.Text
            QAWOTCInboundScorecard.cboSupervisor.Text = cboSupervisor.Text



            QAWOTCInboundScorecard.lblContactType1.Text = cboContactType.Text
            QAWOTCInboundScorecard.txtSR.Text = txtSRNumber.Text
            QAWOTCInboundScorecard.txtContactID.Text = txtContactID.Text


            QAWOTCInboundScorecard.txtContactName.Text = txtContactName.Text
            QAWOTCInboundScorecard.txtContactEmail.Text = txtContactEmail.Text
            QAWOTCInboundScorecard.txtContactPhone.Text = txtContactPhone.Text
            QAWOTCInboundScorecard.txtAccountNum.Text = txtAccountNum.Text
            QAWOTCInboundScorecard.txtCompany.Text = txtCompany.Text
            QAWOTCInboundScorecard.txtOrderID.Text = txtOrderID.Text


            'QANuCallScorecard.dtpCondate.Text = stringg

            QAWOTCInboundScorecard.dtpCondate.Text = ProgramAsDate.ToString(ProgramDateForamt)



            QAWOTCInboundScorecard.lblQAauditor1.Text = lblQAauditor.Text


            QAWOTCInboundScorecard.lblWeekNumber.Text = lblWeekNumber.Text



            QAEmailScorecard.Hide()
            QAChatScorecard.Hide()
            QALvl2CallScorecard.Hide()
            QAlvl2EmailScorecard.Hide()
            QAResCallScorecard.Hide()
            QAResidentEmailScorecard.Hide()
            QAConsuACallScorcard.Hide()
            QACallScorecard.Hide()

            Me.Cursor = Cursors.Hand


            Me.Hide()
            QAWOTCInboundScorecard.Show()




        End If












    End Sub

    Public Sub UpdateAppDecider()

        If lblAppVersion.Text <> lblAppVersion2.Text Then

            lblPleaseUpdateApp.Visible = True

            QACallScorecard.lblPleaseUpdateApp.Visible = True
            QAEmailScorecard.lblPleaseUpdateApp.Visible = True
            QAChatScorecard.lblPleaseUpdateApp.Visible = True

            Timer4.Enabled = True

        End If



    End Sub

    Public Sub Annoucer()

        Try

            Using con001 = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")



                Dim SQL01 As String = "Select * FROM [AppMessage]"


                Using cmd001 As New SqlCommand(SQL01, con001)



                    con001.Open()


                    Dim reader001 As SqlDataReader

                    reader001 = cmd001.ExecuteReader()





                    While reader001.Read()

                        ''
                        txtQAAppAnnoucment.Text = reader001(1).ToString


                    End While
                    reader001.Close()
                    con001.Close()


                End Using

            End Using


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub


    Public Sub RefreshDatatable()

        Using con003 = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")



            Dim SQL03 As String = "Select * FROM [QAMainDB]"


            Dim reader003 As SqlDataReader

            '   Dim NuDataTable As DataTable


            Using cmd003 As New SqlCommand(SQL03, con003)


                con003.Open()


                reader003 = cmd003.ExecuteReader()



                While reader003.Read()




                End While
                reader003.Close()
                con003.Close()


            End Using

        End Using







    End Sub



    Public Sub Changer()

        Try

            Using con001 = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")



                Dim SQL01 As String = "Select * FROM [LogIn]"


                Using cmd001 As New SqlCommand(SQL01, con001)



                    con001.Open()


                    Dim reader001 As SqlDataReader

                    reader001 = cmd001.ExecuteReader()





                    While reader001.Read()

                        '''
                        lblChanger.Text = reader001(13).ToString

                        txtChanger.Text = reader001(13).ToString
                        lblAnn.Text = reader001(10).ToString
                        lblAppVersion2.Text = reader001(11).ToString

                        If lblChanger.Text <> txtChanger.Text Then

                            Timer3.Enabled = True
                            ' MsgBox("1")

                        End If

                    End While
                    reader001.Close()
                    con001.Close()


                End Using

            End Using


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


    End Sub


    Private Sub btnSaveInfo_Click(sender As Object, e As EventArgs) Handles btnLaunchSC.Click

        Try

            Changer()

            ' Timer6.Enabled = True

            Me.Cursor = Cursors.AppStarting
            SplashScreenManager1.ShowWaitForm()
            Annoucer()

            QACallScorecard.lblWeekNumber.Text = lblWeekNumber.Text
            QAEmailScorecard.lblWeekNumber.Text = lblWeekNumber.Text
            QAChatScorecard.lblWeekNumber.Text = lblWeekNumber.Text




            QACallScorecard.lblUserEmail.Text = lblUserEmail.Text
            QAEmailScorecard.lblUserEmail.Text = lblUserEmail.Text
            QAChatScorecard.lblUserEmail.Text = lblUserEmail.Text



            QACallScorecard.lblMDrive.Text = lblMDrive.Text
            QAEmailScorecard.lblMdrive.Text = lblMDrive.Text
            QAChatScorecard.lblMDrive.Text = lblMDrive.Text


            QACallScorecard.lblSCRN1.Text = lblSCRN.Text
            QAChatScorecard.lblSCRN1.Text = lblSCRN.Text
            QAEmailScorecard.lblSCRN.Text = lblSCRN.Text



            QACallScorecard.lblSupervisorEmail.Text = lblSupervisorEmail.Text
            QAChatScorecard.lblSupervisorEmail.Text = lblSupervisorEmail.Text
            QAEmailScorecard.lblSupervisorEmail.Text = lblSupervisorEmail.Text


            QACallScorecard.txtAgentEmail.Text = txtAgentEmail.Text
            QAEmailScorecard.txtAgentEmail.Text = txtAgentEmail.Text
            QAChatScorecard.txtAgentEmail.Text = txtAgentEmail.Text

            QACallScorecard.lblEmailPassword.Text = lblEmailPassword.Text
            QAEmailScorecard.lblEmailPassword.Text = lblEmailPassword.Text
            QAChatScorecard.lblEmailPassword.Text = lblEmailPassword.Text




            UpdateAppDecider()





            '    QAChatScorecard.txtAgentEmail.Text = txtAgentEmail.Text

            '   QAEmailScorecard.txtAgentEmail.Text = txtAgentEmail.Text






            contemp1.Close()


            contemp6.Close()

            contemp16.Close()

            If cboContactType.Text = "Contact Type" Then

                SplashScreenManager1.CloseWaitForm()
                MsgBox("You must select a contact type before you can continue.", MessageBoxButtons.RetryCancel)

                Me.ActiveControl = cboContactType

                Me.Cursor = Cursors.Hand



            Else


                If txtAgentTeam.Text = "Please wait, Loading.." Then

                    SplashScreenManager1.CloseWaitForm()
                    MsgBox("The agent's team field is still loading, please wait until a team name appears before launching the scorecard.", MessageBoxButtons.RetryCancel)

                    Me.ActiveControl = txtAgentTeam

                    Me.Cursor = Cursors.Hand


                Else




                    If QACallScorecard.lblQAScore1.Visible = True Or QAEmailScorecard.lblQAScore1.Visible = True Or QAChatScorecard.lblQAScore1.Visible = True Or QALvl2CallScorecard.lblQAScore1.Visible = True Or QAlvl2EmailScorecard.lblQAScore.Visible = True Or QAResCallScorecard.lblQAScore1.Visible = True Or QAResidentEmailScorecard.lblQAScore.Visible = True Or QAConsuACallScorcard.lblQAScore1.Visible = True Then


                        SplashScreenManager1.CloseWaitForm()
                        MsgBox("You can not save a scorecard that has been scored already, press 'clear fields' button", MessageBoxButtons.OK)

                        Me.Cursor = Cursors.Hand



                    Else



                        QACallScorecard.AgentEmailonCall = AgentEmail


                        ''Launch scorecard

                        ScorecardLaunchMain()





                    End If


                End If
            End If





        Catch ex As Exception



            MsgBox(ex.Message)

        End Try

    End Sub

    Public Sub ScorecardLaunchMain()

        If lblDeciderDash.Text = "QaAuditor" Then

            If cboContactType.Text = "Call" Then
                ''Call
                SplashScreenManager1.CloseWaitForm()
                QACallScorecard.cboAgentName.Text = cboAgentName.Text
                QACallScorecard.cboSupervisor.Text = cboSupervisor.Text
                QACallScorecard.txtTeamName.Text = txtAgentTeam.Text

                'QACallScorecard.cboAgentName.Text = cboAgentName.Text
                'QACallScorecard.cboTeamName
                QACallScorecard.cboSupervisorbox.Visible = False

                QACallScorecard.lblDecider.Text = 1

                ScorecardLaunch()


            ElseIf cboContactType.Text = "Email" Then

                ''Email
                SplashScreenManager1.CloseWaitForm()
                QAEmailScorecard.txtTeamName.Text = txtAgentTeam.Text
                QAEmailScorecard.cboAgentName.Text = cboAgentName.Text
                QAEmailScorecard.cboSupervisor.Text = cboSupervisor.Text
                QAEmailScorecard.cboSupervisorbox.Visible = False


                ScorecardLaunch()


            ElseIf cboContactType.Text = "Chat" Then


                QAEmailScorecard.lblDecider.Text = 1


                SplashScreenManager1.CloseWaitForm()

                ''Chat 

                QAChatScorecard.txtTeamName.Text = txtAgentTeam.Text
                ' QAChatScorecard.cboContactTypeChat.Text = cboContactType.Text

                QAChatScorecard.cboSupervisorbox.Visible = False
                QAChatScorecard.cboAgentName.Text = cboAgentName.Text
                QAChatScorecard.cboSupervisor.Text = cboSupervisor.Text

                QAChatScorecard.lblDecider.Text = 1



                cboAgentName.Enabled = False

                cboContactType.Enabled = False

                cboSupervisor.Enabled = False




                ScorecardLaunch()


            ElseIf cboContactType.Text = "WOTC Inbound" Then
                ''Call
                SplashScreenManager1.CloseWaitForm()
                QAWOTCInboundScorecard.cboAgentName.Text = cboAgentName.Text
                QAWOTCInboundScorecard.cboSupervisor.Text = cboSupervisor.Text
                QAWOTCInboundScorecard.txtTeamName.Text = txtAgentTeam.Text
                QAWOTCInboundScorecard.txtAgentEmail.Text = txtAgentEmail.Text

                'QAWOTCInboundScorecard.cboAgentName.Text = cboAgentName.Text
                'QAWOTCInboundScorecard.cboTeamName
                QAWOTCInboundScorecard.cboSupervisorbox.Visible = False

                QAWOTCInboundScorecard.lblDecider.Text = 1

                ScorecardLaunch()




                End If

        End If



        ''Supervisor log in

        If lblDeciderDash.Text = "Supervisor" Then

            '    ElseIf lblQAauditor.Text IsNot "Carla Hardy" Then


            ''Call

            If cboContactType.Text = "Call" Then

                SplashScreenManager1.CloseWaitForm()
                QACallScorecard.txtTeamName.Text = txtAgentTeam.Text


                QACallScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text


                QACallScorecard.cboSupervisorbox.Location = New System.Drawing.Point(197, 266)
                QACallScorecard.txtTeamName.Location = New System.Drawing.Point(197, 315)
                QACallScorecard.Label16.Location = New System.Drawing.Point(194, 299)



                QACallScorecard.lblagentE.Location = New System.Drawing.Point(196, 348)
                QACallScorecard.txtAgentEmail.Location = New System.Drawing.Point(197, 366)




                QACallScorecard.cboSupervisor.Visible = False
                QACallScorecard.cboAgentName.Visible = False
                QACallScorecard.txtTeamName.Visible = True
                QACallScorecard.Label16.Visible = True

                QACallScorecard.lblContactTypeDropDown.Visible = False


                QACallScorecard.lblDecider.Text = 2


                ScorecardLaunch()

            ElseIf cboContactType.Text = "Email" Then

                SplashScreenManager1.CloseWaitForm()

                QAEmailScorecard.cboSupervisorbox.Location = New System.Drawing.Point(209, 266)
                QAEmailScorecard.txtTeamName.Location = New System.Drawing.Point(208, 318)
                QAEmailScorecard.Label13.Location = New System.Drawing.Point(205, 300)

                QAEmailScorecard.lblagentE.Location = New System.Drawing.Point(207, 353)
                QAEmailScorecard.txtAgentEmail.Location = New System.Drawing.Point(208, 370)




                QAEmailScorecard.txtTeamName.Text = txtAgentTeam.Text


                ''Email

                QAEmailScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text



                QAEmailScorecard.cboSupervisor.Visible = False
                QAEmailScorecard.cboAgentName.Visible = False
                QAEmailScorecard.txtTeamName.Visible = True
                QAEmailScorecard.Label13.Visible = True



                QAEmailScorecard.lblContactTypeDropDown.Visible = False
                QAEmailScorecard.lblDecider.Text = 2

                ScorecardLaunch()



            ElseIf cboContactType.Text = "Chat" Then
                ''Chat

                SplashScreenManager1.CloseWaitForm()

                QAChatScorecard.cboSupervisorbox.Location = New System.Drawing.Point(196, 265)
                QAChatScorecard.txtTeamName.Location = New System.Drawing.Point(196, 315)
                QAChatScorecard.Label16.Location = New System.Drawing.Point(193, 299)

                QAChatScorecard.lblagentE.Location = New System.Drawing.Point(195, 352)
                QAChatScorecard.txtAgentEmail.Location = New System.Drawing.Point(196, 370)



                QAChatScorecard.txtTeamName.Text = txtAgentTeam.Text


                QAChatScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text

                QAChatScorecard.lblContactTypeDropDown.Visible = False

                QAChatScorecard.cboSupervisor.Visible = False
                QAChatScorecard.cboAgentName.Visible = False
                QAChatScorecard.txtTeamName.Visible = True
                QAChatScorecard.Label16.Visible = True



                QAChatScorecard.lblDecider.Text = 2



                ScorecardLaunch()



            ElseIf cboContactType.Text = "WOTC Inbound" Then

                SplashScreenManager1.CloseWaitForm()
                QAWOTCInboundScorecard.txtTeamName.Text = txtAgentTeam.Text


                QAWOTCInboundScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text


                'QAWOTCInboundScorecard.cboSupervisorbox.Location = New System.Drawing.Point(197, 266)
                'QAWOTCInboundScorecard.txtTeamName.Location = New System.Drawing.Point(197, 315)
                'QAWOTCInboundScorecard.Label16.Location = New System.Drawing.Point(194, 299)



                'QAWOTCInboundScorecard.lblagentE.Location = New System.Drawing.Point(196, 348)
                'QAWOTCInboundScorecard.txtAgentEmail.Location = New System.Drawing.Point(197, 366)


                QAWOTCInboundScorecard.txtAgentEmail.Text = txtAgentEmail.Text

                QAWOTCInboundScorecard.cboSupervisor.Visible = False
                QAWOTCInboundScorecard.cboAgentName.Visible = False
                QAWOTCInboundScorecard.txtTeamName.Visible = True
                QAWOTCInboundScorecard.Label16.Visible = True

                QAWOTCInboundScorecard.lblContactTypeDropDown.Visible = False


                QAWOTCInboundScorecard.lblDecider.Text = 2


                ScorecardLaunch()





                End If
        End If


        If lblDeciderDash.Text = "Team Lead" Then

            '    ElseIf lblQAauditor.Text IsNot "Carla Hardy" Then


            ''Call

            If cboContactType.Text = "Call" Then

                SplashScreenManager1.CloseWaitForm()
                QACallScorecard.txtTeamName.Text = txtAgentTeam.Text


                QACallScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text


                QACallScorecard.cboSupervisorbox.Location = New System.Drawing.Point(197, 266)
                QACallScorecard.txtTeamName.Location = New System.Drawing.Point(197, 315)
                QACallScorecard.Label16.Location = New System.Drawing.Point(194, 299)



                QACallScorecard.lblagentE.Location = New System.Drawing.Point(196, 348)
                QACallScorecard.txtAgentEmail.Location = New System.Drawing.Point(197, 366)




                QACallScorecard.cboSupervisor.Visible = False
                QACallScorecard.cboAgentName.Visible = False
                QACallScorecard.txtTeamName.Visible = True
                QACallScorecard.Label16.Visible = True

                QACallScorecard.lblContactTypeDropDown.Visible = False


                QACallScorecard.lblDecider.Text = 2


                ScorecardLaunch()

            ElseIf cboContactType.Text = "Email" Then


                SplashScreenManager1.CloseWaitForm()
                QAEmailScorecard.cboSupervisorbox.Location = New System.Drawing.Point(209, 266)
                QAEmailScorecard.txtTeamName.Location = New System.Drawing.Point(208, 318)
                QAEmailScorecard.Label13.Location = New System.Drawing.Point(205, 300)

                QAEmailScorecard.lblagentE.Location = New System.Drawing.Point(207, 353)
                QAEmailScorecard.txtAgentEmail.Location = New System.Drawing.Point(208, 370)




                QAEmailScorecard.txtTeamName.Text = txtAgentTeam.Text


                ''Email

                QAEmailScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text



                QAEmailScorecard.cboSupervisor.Visible = False
                QAEmailScorecard.cboAgentName.Visible = False
                QAEmailScorecard.txtTeamName.Visible = True
                QAEmailScorecard.Label13.Visible = True



                QAEmailScorecard.lblContactTypeDropDown.Visible = False
                QAEmailScorecard.lblDecider.Text = 2

                ScorecardLaunch()



            ElseIf cboContactType.Text = "Chat" Then
                ''Chat

                SplashScreenManager1.CloseWaitForm()

                QAChatScorecard.cboSupervisorbox.Location = New System.Drawing.Point(196, 265)
                QAChatScorecard.txtTeamName.Location = New System.Drawing.Point(196, 315)
                QAChatScorecard.Label16.Location = New System.Drawing.Point(193, 299)

                QAChatScorecard.lblagentE.Location = New System.Drawing.Point(195, 352)
                QAChatScorecard.txtAgentEmail.Location = New System.Drawing.Point(196, 370)



                QAChatScorecard.txtTeamName.Text = txtAgentTeam.Text


                QAChatScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text

                QAChatScorecard.lblContactTypeDropDown.Visible = False

                QAChatScorecard.cboSupervisor.Visible = False
                QAChatScorecard.cboAgentName.Visible = False
                QAChatScorecard.txtTeamName.Visible = True
                QAChatScorecard.Label16.Visible = True



                QAChatScorecard.lblDecider.Text = 2



                ScorecardLaunch()


            ElseIf cboContactType.Text = "WOTC Inbound" Then

                SplashScreenManager1.CloseWaitForm()
                QAWOTCInboundScorecard.txtTeamName.Text = txtAgentTeam.Text


                QAWOTCInboundScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text


                'QAWOTCInboundScorecard.cboSupervisorbox.Location = New System.Drawing.Point(197, 266)
                'QAWOTCInboundScorecard.txtTeamName.Location = New System.Drawing.Point(197, 315)
                'QAWOTCInboundScorecard.Label16.Location = New System.Drawing.Point(194, 299)



                'QAWOTCInboundScorecard.lblagentE.Location = New System.Drawing.Point(196, 348)
                'QAWOTCInboundScorecard.txtAgentEmail.Location = New System.Drawing.Point(197, 366)


                QAWOTCInboundScorecard.txtAgentEmail.Text = txtAgentEmail.Text

                QAWOTCInboundScorecard.cboSupervisor.Visible = False
                QAWOTCInboundScorecard.cboAgentName.Visible = False
                QAWOTCInboundScorecard.txtTeamName.Visible = True
                QAWOTCInboundScorecard.Label16.Visible = True

                QAWOTCInboundScorecard.lblContactTypeDropDown.Visible = False


                QAWOTCInboundScorecard.lblDecider.Text = 2


                ScorecardLaunch()







                End If
            End If






        If lblDeciderDash.Text = "Admin" Then

            If cboContactType.Text = "Call" Then
                ''Call

                SplashScreenManager1.CloseWaitForm()

                QACallScorecard.cboAgentName.Text = cboAgentName.Text
                QACallScorecard.cboSupervisor.Text = cboSupervisor.Text
                QACallScorecard.txtTeamName.Text = txtAgentTeam.Text

                'QACallScorecard.cboAgentName.Text = cboAgentName.Text
                'QACallScorecard.cboTeamName
                QACallScorecard.cboSupervisorbox.Visible = False

                QACallScorecard.lblDecider.Text = 1

                ScorecardLaunch()


            ElseIf cboContactType.Text = "Email" Then

                ''Email

                SplashScreenManager1.CloseWaitForm()

                QAEmailScorecard.txtTeamName.Text = txtAgentTeam.Text
                QAEmailScorecard.cboAgentName.Text = cboAgentName.Text
                QAEmailScorecard.cboSupervisor.Text = cboSupervisor.Text
                QAEmailScorecard.cboSupervisorbox.Visible = False


                ScorecardLaunch()


            ElseIf cboContactType.Text = "Chat" Then


                QAEmailScorecard.lblDecider.Text = 1


                SplashScreenManager1.CloseWaitForm()

                ''Chat 

                QAChatScorecard.txtTeamName.Text = txtAgentTeam.Text
                ' QAChatScorecard.cboContactTypeChat.Text = cboContactType.Text

                QAChatScorecard.cboSupervisorbox.Visible = False
                QAChatScorecard.cboAgentName.Text = cboAgentName.Text
                QAChatScorecard.cboSupervisor.Text = cboSupervisor.Text

                QAChatScorecard.lblDecider.Text = 1



                cboAgentName.Enabled = False

                cboContactType.Enabled = False

                cboSupervisor.Enabled = False




                ScorecardLaunch()



            ElseIf cboContactType.Text = "WOTC Inbound" Then
                ''Call

                SplashScreenManager1.CloseWaitForm()

                QAWOTCInboundScorecard.cboAgentName.Text = cboAgentName.Text
                QAWOTCInboundScorecard.cboSupervisor.Text = cboSupervisor.Text
                QAWOTCInboundScorecard.txtTeamName.Text = txtAgentTeam.Text
                QAWOTCInboundScorecard.txtAgentEmail.Text = txtAgentEmail.Text

                'QAWOTCInboundScorecard.cboAgentName.Text = cboAgentName.Text
                'QAWOTCInboundScorecard.cboTeamName
                QAWOTCInboundScorecard.cboSupervisorbox.Visible = False

                QAWOTCInboundScorecard.lblDecider.Text = 1

                ScorecardLaunch()






                End If

        End If

        If lblDeciderDash.Text = "GOCSupervisor" Then

            '    ElseIf lblQAauditor.Text IsNot "Carla Hardy" Then


            ''Call

            If cboContactType.Text = "Call" Then


                QACallScorecard.txtTeamName.Text = txtAgentTeam.Text


                QACallScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text


                QACallScorecard.cboSupervisorbox.Location = New System.Drawing.Point(197, 266)
                QACallScorecard.txtTeamName.Location = New System.Drawing.Point(197, 315)
                QACallScorecard.Label16.Location = New System.Drawing.Point(194, 299)



                QACallScorecard.lblagentE.Location = New System.Drawing.Point(196, 348)
                QACallScorecard.txtAgentEmail.Location = New System.Drawing.Point(197, 366)




                QACallScorecard.cboSupervisor.Visible = False
                QACallScorecard.cboAgentName.Visible = False
                QACallScorecard.txtTeamName.Visible = True
                QACallScorecard.Label16.Visible = True

                QACallScorecard.lblContactTypeDropDown.Visible = False


                QACallScorecard.lblDecider.Text = 2


                ScorecardLaunch()

            ElseIf cboContactType.Text = "Email" Then



                QAEmailScorecard.cboSupervisorbox.Location = New System.Drawing.Point(209, 266)
                QAEmailScorecard.txtTeamName.Location = New System.Drawing.Point(208, 318)
                QAEmailScorecard.Label13.Location = New System.Drawing.Point(205, 300)

                QAEmailScorecard.lblagentE.Location = New System.Drawing.Point(207, 353)
                QAEmailScorecard.txtAgentEmail.Location = New System.Drawing.Point(208, 370)




                QAEmailScorecard.txtTeamName.Text = txtAgentTeam.Text


                ''Email

                QAEmailScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text



                QAEmailScorecard.cboSupervisor.Visible = False
                QAEmailScorecard.cboAgentName.Visible = False
                QAEmailScorecard.txtTeamName.Visible = True
                QAEmailScorecard.Label13.Visible = True



                QAEmailScorecard.lblContactTypeDropDown.Visible = False
                QAEmailScorecard.lblDecider.Text = 2

                ScorecardLaunch()



            ElseIf cboContactType.Text = "Chat" Then
                ''Chat



                QAChatScorecard.cboSupervisorbox.Location = New System.Drawing.Point(196, 265)
                QAChatScorecard.txtTeamName.Location = New System.Drawing.Point(196, 315)
                QAChatScorecard.Label16.Location = New System.Drawing.Point(193, 299)

                QAChatScorecard.lblagentE.Location = New System.Drawing.Point(195, 352)
                QAChatScorecard.txtAgentEmail.Location = New System.Drawing.Point(196, 370)



                QAChatScorecard.txtTeamName.Text = txtAgentTeam.Text


                QAChatScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text

                QAChatScorecard.lblContactTypeDropDown.Visible = False

                QAChatScorecard.cboSupervisor.Visible = False
                QAChatScorecard.cboAgentName.Visible = False
                QAChatScorecard.txtTeamName.Visible = True
                QAChatScorecard.Label16.Visible = True



                QAChatScorecard.lblDecider.Text = 2



                ScorecardLaunch()




            ElseIf cboContactType.Text = "WOTC Inbound" Then


                QAWOTCInboundScorecard.txtTeamName.Text = txtAgentTeam.Text


                QAWOTCInboundScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text


                'QAWOTCInboundScorecard.cboSupervisorbox.Location = New System.Drawing.Point(197, 266)
                'QAWOTCInboundScorecard.txtTeamName.Location = New System.Drawing.Point(197, 315)
                'QAWOTCInboundScorecard.Label16.Location = New System.Drawing.Point(194, 299)



                'QAWOTCInboundScorecard.lblagentE.Location = New System.Drawing.Point(196, 348)
                'QAWOTCInboundScorecard.txtAgentEmail.Location = New System.Drawing.Point(197, 366)

                QAWOTCInboundScorecard.txtAgentEmail.Text = txtAgentEmail.Text


                QAWOTCInboundScorecard.cboSupervisor.Visible = False
                QAWOTCInboundScorecard.cboAgentName.Visible = False
                QAWOTCInboundScorecard.txtTeamName.Visible = True
                QAWOTCInboundScorecard.Label16.Visible = True

                QAWOTCInboundScorecard.lblContactTypeDropDown.Visible = False


                QAWOTCInboundScorecard.lblDecider.Text = 2


                ScorecardLaunch()



                    End If

        End If


            If lblDeciderDash.Text = "GOCTeamLead" Then

                '    ElseIf lblQAauditor.Text IsNot "Carla Hardy" Then


                ''Call

                If cboContactType.Text = "Call" Then


                    QACallScorecard.txtTeamName.Text = txtAgentTeam.Text


                    QACallScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text


                    QACallScorecard.cboSupervisorbox.Location = New System.Drawing.Point(197, 266)
                    QACallScorecard.txtTeamName.Location = New System.Drawing.Point(197, 315)
                    QACallScorecard.Label16.Location = New System.Drawing.Point(194, 299)



                    QACallScorecard.lblagentE.Location = New System.Drawing.Point(196, 348)
                    QACallScorecard.txtAgentEmail.Location = New System.Drawing.Point(197, 366)




                    QACallScorecard.cboSupervisor.Visible = False
                    QACallScorecard.cboAgentName.Visible = False
                    QACallScorecard.txtTeamName.Visible = True
                    QACallScorecard.Label16.Visible = True

                    QACallScorecard.lblContactTypeDropDown.Visible = False


                    QACallScorecard.lblDecider.Text = 2


                    ScorecardLaunch()

                ElseIf cboContactType.Text = "Email" Then



                    QAEmailScorecard.cboSupervisorbox.Location = New System.Drawing.Point(209, 266)
                    QAEmailScorecard.txtTeamName.Location = New System.Drawing.Point(208, 318)
                    QAEmailScorecard.Label13.Location = New System.Drawing.Point(205, 300)

                    QAEmailScorecard.lblagentE.Location = New System.Drawing.Point(207, 353)
                    QAEmailScorecard.txtAgentEmail.Location = New System.Drawing.Point(208, 370)




                    QAEmailScorecard.txtTeamName.Text = txtAgentTeam.Text


                    ''Email

                    QAEmailScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text



                    QAEmailScorecard.cboSupervisor.Visible = False
                    QAEmailScorecard.cboAgentName.Visible = False
                    QAEmailScorecard.txtTeamName.Visible = True
                    QAEmailScorecard.Label13.Visible = True



                    QAEmailScorecard.lblContactTypeDropDown.Visible = False
                    QAEmailScorecard.lblDecider.Text = 2

                    ScorecardLaunch()



                ElseIf cboContactType.Text = "Chat" Then
                    ''Chat



                    QAChatScorecard.cboSupervisorbox.Location = New System.Drawing.Point(196, 265)
                    QAChatScorecard.txtTeamName.Location = New System.Drawing.Point(196, 315)
                    QAChatScorecard.Label16.Location = New System.Drawing.Point(193, 299)

                    QAChatScorecard.lblagentE.Location = New System.Drawing.Point(195, 352)
                    QAChatScorecard.txtAgentEmail.Location = New System.Drawing.Point(196, 370)



                    QAChatScorecard.txtTeamName.Text = txtAgentTeam.Text


                    QAChatScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text

                    QAChatScorecard.lblContactTypeDropDown.Visible = False

                    QAChatScorecard.cboSupervisor.Visible = False
                    QAChatScorecard.cboAgentName.Visible = False
                    QAChatScorecard.txtTeamName.Visible = True
                    QAChatScorecard.Label16.Visible = True



                    QAChatScorecard.lblDecider.Text = 2



                    ScorecardLaunch()



                If cboContactType.Text = "WOTC Inbound" Then


                    QAWOTCInboundScorecard.txtTeamName.Text = txtAgentTeam.Text


                    QAWOTCInboundScorecard.cboSupervisorbox.Text = cboSuperAgentBox.Text


                    'QAWOTCInboundScorecard.cboSupervisorbox.Location = New System.Drawing.Point(197, 266)
                    'QAWOTCInboundScorecard.txtTeamName.Location = New System.Drawing.Point(197, 315)
                    'QAWOTCInboundScorecard.Label16.Location = New System.Drawing.Point(194, 299)



                    'QAWOTCInboundScorecard.lblagentE.Location = New System.Drawing.Point(196, 348)
                    'QAWOTCInboundScorecard.txtAgentEmail.Location = New System.Drawing.Point(197, 366)


                    QAWOTCInboundScorecard.txtAgentEmail.Text = txtAgentEmail.Text

                    QAWOTCInboundScorecard.cboSupervisor.Visible = False
                    QAWOTCInboundScorecard.cboAgentName.Visible = False
                    QAWOTCInboundScorecard.txtTeamName.Visible = True
                    QAWOTCInboundScorecard.Label16.Visible = True

                    QAWOTCInboundScorecard.lblContactTypeDropDown.Visible = False


                    QAWOTCInboundScorecard.lblDecider.Text = 2


                    ScorecardLaunch()




                End If

            End If


            End If




    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Try

            Dim i As Integer

            With DataGridView2


                If e.RowIndex >= 0 Then

                    i = DataGridView2.CurrentRow.Index


                    Clipboard.SetDataObject(DataGridView2.GetClipboardContent)

                    ProgressBar1.Value = 0


                    If DataGridView2.Rows(i).Cells("DataGridViewTextBoxColumn2").Value.ToString = "" Then

                        txtSRPort.Text = DataGridView2.Rows(i).Cells("DataGridViewTextBoxColumn14").Value.ToString

                        txtAuditType.Text = DataGridView2.Rows(i).Cells("DataGridViewTextBoxColumn15").Value.ToString


                    Else

                        txtSRPort.Text = DataGridView2.Rows(i).Cells("DataGridViewTextBoxColumn2").Value.ToString

                        txtAuditType.Text = DataGridView2.Rows(i).Cells("DataGridViewTextBoxColumn15").Value.ToString



                        'TextBox4.Text = DataGridView2.Rows(i).Cells("DataGridViewTextBoxColumn16").Value.ToString

                        'TextBox5.Text = DataGridView2.Rows(i).Cells("DataGridViewTextBoxColumn17").Value.ToString

                        'TextBox6.Text = DataGridView2.Rows(i).Cells("DataGridViewTextBoxColumn15").Value.ToString


                        'txtImport.Text = DataGridView2.Rows(i).Cells("DataGridViewTextBoxColumn2").Value.ToString

                    End If


                End If



            End With




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click


        ''SpreadsheetControl1.LoadDocument("C:\Users\playe\Desktop\QA\QAExcel.xlsx", DocumentFormat.Xlsx)





        ' MsgBox("Updated")


        ExcelStuff.Macro()



        ' Revpen.Macro()


        RefreshExcel.Enabled = True







    End Sub


    Public Sub CallPort()








    End Sub

    Public Sub OpenScoreCard()



        ' Using con0 = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")

        Using con0 = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")



            ' Dim SQL0 As String = "SELECT * FROM [QAMainDB] WHERE SR= ? AND Ctype= ?"

            Dim SQL0 As String = "SELECT * FROM [QAMainDB] WHERE SR= ? or ContactID= ?"

            '    Dim SQL0 As String = "SELECT * FROM [QAMainDB] WHERE SR= ? "

            Using cmd0 As New SqlCommand(SQL0, con0)





                cmd0.Parameters.AddWithValue("@p1", txtSRPort.Text)
                cmd0.Parameters.AddWithValue("@p2", txtSRPort.Text)


                '   cmd0.Parameters.AddWithValue("@p2", DataGridViewTextBoxColumn18.ToString)



                con0.Open()



                Dim reader0 As SqlDataReader

                reader0 = cmd0.ExecuteReader()




                While reader0.Read()





                    If txtAuditType.Text = "Call" Then



                        QACallRev.txtSR.Text = reader0(1).ToString()
                        QACallRev.txtContactID.Text = reader0(2).ToString()
                        '    QACallScorecard.txtContactID.Text = reader0(3).ToString() contact type
                        QACallRev.cboAgentName.Text = reader0(4).ToString()
                        QACallRev.cboTeamName.Text = reader0(5).ToString()
                        QACallRev.dtpCondate.Text = reader0(6).ToString()
                        QACallRev.txtOrderID.Text = reader0(7).ToString()
                        QACallRev.txtQADate.Text = reader0(8).ToString()
                        QACallRev.txtQACom.Text = reader0(9).ToString()
                        QACallRev.txtQAAOO.Text = reader0(10).ToString()
                        QACallRev.txtContactName.Text = reader0(11).ToString()
                        QACallRev.txtAccountNum.Text = reader0(12).ToString()
                        QACallRev.txtCompany.Text = reader0(13).ToString
                        QACallRev.txtContactPhone.Text = reader0(14).ToString
                        QACallRev.txtContactEmail.Text = reader0(15).ToString

                        QACallRev.dtpReviewdate.Text = reader0(16).ToString
                        QACallRev.lblrevMan.Text = reader0(17).ToString
                        QACallRev.txtRevComments.Text = reader0(18).ToString
                        QACallRev.txtDisputeScore.Text = reader0(19).ToString
                        QACallRev.txtDisputerName.Text = reader0(20).ToString
                        QACallRev.txtDisputeNotes.Text = reader0(21).ToString
                        QACallRev.txtDisComment.Text = reader0(22).ToString


                        QACallRev.cbo1_1.Text = reader0(23).ToString
                        QACallRev.cbo1_2.Text = reader0(24).ToString
                        QACallRev.cbo1_3.Text = reader0(25).ToString

                        QACallRev.txt1_1.Text = reader0(32).ToString
                        QACallRev.txt1_2.Text = reader0(33).ToString
                        QACallRev.txt1_3.Text = reader0(34).ToString


                        QACallRev.cbo2_1.Text = reader0(41).ToString
                        QACallRev.txt2_1.Text = reader0(50).ToString



                        QACallRev.cbo3_1.Text = reader0(59).ToString
                        QACallRev.cbo3_2.Text = reader0(60).ToString
                        QACallRev.cbo3_3.Text = reader0(61).ToString
                        QACallRev.cbo3_4.Text = reader0(62).ToString
                        QACallRev.cbo3_5.Text = reader0(63).ToString
                        QACallRev.cbo3_6.Text = reader0(64).ToString
                        QACallRev.cbo3_7.Text = reader0(65).ToString
                        QACallRev.cbo3_8.Text = reader0(66).ToString


                        QACallRev.txt3_1.Text = reader0(68).ToString
                        QACallRev.txt3_2.Text = reader0(69).ToString
                        QACallRev.txt3_3.Text = reader0(70).ToString
                        QACallRev.txt3_4.Text = reader0(71).ToString
                        QACallRev.txt3_5.Text = reader0(72).ToString
                        QACallRev.txt3_6.Text = reader0(73).ToString
                        QACallRev.txt3_7.Text = reader0(74).ToString
                        QACallRev.txt3_8.Text = reader0(75).ToString


                        QACallRev.Cbo4_1.Text = reader0(77).ToString
                        QACallRev.cbo4_2.Text = reader0(78).ToString
                        QACallRev.cbo4_3.Text = reader0(79).ToString

                        QACallRev.txt4_1.Text = reader0(86).ToString
                        QACallRev.txt4_2.Text = reader0(87).ToString
                        QACallRev.txt4_3.Text = reader0(88).ToString


                        QACallRev.cbo5_1.Text = reader0(95).ToString
                        QACallRev.cbo5_2.Text = reader0(96).ToString


                        QACallRev.txt5_1.Text = reader0(104).ToString
                        QACallRev.txt5_2.Text = reader0(105).ToString


                        QACallRev.cbo6_1.Text = reader0(113).ToString
                        QACallRev.cbo6_2.Text = reader0(114).ToString
                        QACallRev.cbo6_3.Text = reader0(115).ToString

                        QACallRev.txt6_1.Text = reader0(122).ToString
                        QACallRev.txt6_2.Text = reader0(123).ToString
                        QACallRev.txt6_3.Text = reader0(124).ToString

                        QACallRev.cbo7_1.Text = reader0(131).ToString
                        QACallRev.cbo7_2.Text = reader0(132).ToString
                        QACallRev.cbo7_3.Text = reader0(133).ToString
                        QACallRev.cbo7_4.Text = reader0(134).ToString
                        QACallRev.cbo7_5.Text = reader0(135).ToString
                        QACallRev.cbo7_6.Text = reader0(136).ToString

                        QACallRev.txt7_1.Text = reader0(140).ToString
                        QACallRev.txt7_2.Text = reader0(141).ToString
                        QACallRev.txt7_3.Text = reader0(142).ToString
                        QACallRev.txt7_4.Text = reader0(143).ToString
                        QACallRev.txt7_5.Text = reader0(144).ToString
                        QACallRev.txt7_6.Text = reader0(145).ToString

                        QACallRev.txtQAScore.Text = reader0(149).ToString


                        QACallRev.cboAF.Text = reader0(152).ToString


                        QACallRev.txtOrignalAuditor.Text = reader0(153).ToString
                        QACallRev.txtSupervisor.Text = reader0(155).ToString

                        QACallRev.txtTCXScore.Text = reader0(156).ToString
                        QACallRev.txtWeekNumber.Text = reader0(157).ToString



                        QACallRev.lblrevMan.Text = lblQAauditor.Text


                        QACallRev.lblQAauditor1.Text = lblQAauditor.Text



                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                    ElseIf txtAuditType.Text = "Email" Then

                        '  CallTimer.Enabled = True



                        QAEmailRev.txtSR.Text = reader0(1).ToString()
                        QAEmailRev.txtContactID.Text = reader0(2).ToString()
                        'QACallScorecard.txtContactID.Text = reader0(3).ToString() contact type
                        QAEmailRev.cboAgentName.Text = reader0(4).ToString()
                        QAEmailRev.cboTeamName.Text = reader0(5).ToString()
                        QAEmailRev.dtpCondate.Text = reader0(6).ToString()
                        QAEmailRev.txtOrderID.Text = reader0(7).ToString()
                        QAEmailRev.txtQADate.Text = reader0(8).ToString()
                        QAEmailRev.txtQACom.Text = reader0(9).ToString()
                        QAEmailRev.txtQAAOO.Text = reader0(10).ToString()
                        QAEmailRev.txtContactName.Text = reader0(11).ToString

                        QAEmailRev.txtAccountNum.Text = reader0(12).ToString
                        QAEmailRev.txtCompany.Text = reader0(13).ToString
                        QAEmailRev.txtContactPhone.Text = reader0(14).ToString
                        QAEmailRev.txtContactEmail.Text = reader0(15).ToString

                        QAEmailRev.dtpReviewdate.Text = reader0(16).ToString
                        QAEmailRev.lblrevMan.Text = reader0(17).ToString
                        QAEmailRev.txtRevComments.Text = reader0(18).ToString
                        QAEmailRev.txtDisputeScore.Text = reader0(19).ToString
                        '   QAEmailRev.txtDisputerName.Text = reader0(20).ToString
                        QAEmailRev.txtDisputeNotes.Text = reader0(21).ToString
                        QAEmailRev.txtDisComment.Text = reader0(22).ToString




                        QAEmailRev.cbo1_1.Text = reader0(23).ToString
                        QAEmailRev.cbo1_2.Text = reader0(24).ToString
                        QAEmailRev.cbo1_3.Text = reader0(25).ToString

                        QAEmailRev.txt1_1.Text = reader0(32).ToString
                        QAEmailRev.txt1_2.Text = reader0(33).ToString
                        QAEmailRev.txt1_3.Text = reader0(34).ToString



                        QAEmailRev.cbo2_1.Text = reader0(41).ToString
                        QAEmailRev.cbo2_2.Text = reader0(42).ToString
                        QAEmailRev.cbo2_3.Text = reader0(43).ToString
                        QAEmailRev.cbo2_4.Text = reader0(44).ToString



                        QAEmailRev.txt2_1.Text = reader0(50).ToString
                        QAEmailRev.txt2_2.Text = reader0(51).ToString
                        QAEmailRev.txt2_3.Text = reader0(52).ToString
                        QAEmailRev.txt2_4.Text = reader0(53).ToString



                        QAEmailRev.cbo3_1.Text = reader0(59).ToString
                        QAEmailRev.cbo3_2.Text = reader0(60).ToString
                        QAEmailRev.cbo3_3.Text = reader0(61).ToString
                        QAEmailRev.cbo3_4.Text = reader0(62).ToString
                        QAEmailRev.cbo3_5.Text = reader0(63).ToString




                        QAEmailRev.txt3_1.Text = reader0(68).ToString
                        QAEmailRev.txt3_2.Text = reader0(69).ToString
                        QAEmailRev.txt3_3.Text = reader0(70).ToString
                        QAEmailRev.txt3_4.Text = reader0(71).ToString
                        QAEmailRev.txt3_5.Text = reader0(72).ToString






                        QAEmailRev.cbo4_1.Text = reader0(77).ToString
                        QAEmailRev.cbo4_2.Text = reader0(78).ToString
                        QAEmailRev.cbo4_3.Text = reader0(79).ToString
                        QAEmailRev.cbo4_4.Text = reader0(79).ToString



                        QAEmailRev.txt4_1.Text = reader0(86).ToString
                        QAEmailRev.txt4_2.Text = reader0(87).ToString
                        QAEmailRev.txt4_3.Text = reader0(88).ToString
                        QAEmailRev.txt4_4.Text = reader0(89).ToString


                        QAEmailRev.cbo5_1.Text = reader0(95).ToString
                        QAEmailRev.cbo5_2.Text = reader0(96).ToString
                        QAEmailRev.cbo5_3.Text = reader0(97).ToString
                        QAEmailRev.cbo5_4.Text = reader0(98).ToString
                        QAEmailRev.cbo5_5.Text = reader0(99).ToString
                        QAEmailRev.cbo5_6.Text = reader0(100).ToString

                        QAEmailRev.txt5_1.Text = reader0(104).ToString
                        QAEmailRev.txt5_2.Text = reader0(105).ToString
                        QAEmailRev.txt5_3.Text = reader0(106).ToString
                        QAEmailRev.txt5_4.Text = reader0(107).ToString
                        QAEmailRev.txt5_5.Text = reader0(108).ToString
                        QAEmailRev.txt5_6.Text = reader0(109).ToString

                        QAEmailRev.txtQAScore.Text = reader0(149).ToString


                        QAEmailRev.cboAF.Text = reader0(152).ToString


                        QAEmailRev.txtOrignalAuditor.Text = reader0(153).ToString
                        QAEmailRev.txtSupervisor.Text = reader0(155).ToString


                        QAEmailRev.txtTCXScore.Text = reader0(156).ToString

                        QAEmailRev.txtWeekNumber.Text = reader0(157).ToString

                        QAEmailRev.lblrevMan.Text = lblQAauditor.Text

                        QAEmailRev.lblQAauditor1.Text = lblQAauditor.Text




                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    ElseIf txtAuditType.Text = "Chat" Then



                        QAChatRev.txtSR.Text = reader0(1).ToString()
                        QAChatRev.txtContactID.Text = reader0(2).ToString()
                        '    QACallScorecard.txtContactID.Text = reader0(3).ToString() contact type
                        QAChatRev.cboAgentName.Text = reader0(4).ToString()
                        QAChatRev.cboTeamName.Text = reader0(5).ToString()
                        QAChatRev.dtpCondate.Text = reader0(6).ToString()
                        QAChatRev.txtOrderID.Text = reader0(7).ToString()
                        QAChatRev.txtQADate.Text = reader0(8).ToString()
                        QAChatRev.txtQACom.Text = reader0(9).ToString()
                        QAChatRev.txtQAAOO.Text = reader0(10).ToString()
                        QAChatRev.txtContactName.Text = reader0(11).ToString()
                        QAChatRev.txtAccountNum.Text = reader0(12).ToString()
                        QAChatRev.txtCompany.Text = reader0(13).ToString
                        QAChatRev.txtContactPhone.Text = reader0(14).ToString
                        QAChatRev.txtContactEmail.Text = reader0(15).ToString

                        QAChatRev.dtpReviewdate.Text = reader0(16).ToString
                        QAChatRev.lblrevMan.Text = reader0(17).ToString
                        QAChatRev.txtRevComments.Text = reader0(18).ToString
                        QAChatRev.txtDisputeScore.Text = reader0(19).ToString
                        ' QAChatRev.txtDisputerName.Text = reader0(20).ToString
                        QAChatRev.txtDisputeNotes.Text = reader0(21).ToString
                        QAChatRev.txtDiscomment.Text = reader0(22).ToString


                        QAChatRev.cbo1_1.Text = reader0(23).ToString
                        QAChatRev.cbo1_2.Text = reader0(24).ToString

                        QAChatRev.txt1_1.Text = reader0(32).ToString
                        QAChatRev.txt1_2.Text = reader0(33).ToString


                        QAChatRev.cbo2_1.Text = reader0(41).ToString
                        QAChatRev.txt2_1.Text = reader0(50).ToString




                        QAChatRev.cbo3_1.Text = reader0(59).ToString
                        QAChatRev.cbo3_2.Text = reader0(60).ToString
                        QAChatRev.cbo3_3.Text = reader0(61).ToString
                        QAChatRev.cbo3_4.Text = reader0(62).ToString
                        QAChatRev.cbo3_5.Text = reader0(63).ToString
                        QAChatRev.cbo3_6.Text = reader0(64).ToString
                        QAChatRev.cbo3_7.Text = reader0(65).ToString
                        QAChatRev.cbo3_8.Text = reader0(66).ToString



                        QAChatRev.txt3_1.Text = reader0(68).ToString
                        QAChatRev.txt3_2.Text = reader0(69).ToString
                        QAChatRev.txt3_3.Text = reader0(70).ToString
                        QAChatRev.txt3_4.Text = reader0(71).ToString
                        QAChatRev.txt3_5.Text = reader0(72).ToString
                        QAChatRev.txt3_6.Text = reader0(73).ToString
                        QAChatRev.txt3_7.Text = reader0(74).ToString
                        QAChatRev.txt3_8.Text = reader0(75).ToString


                        QAChatRev.Cbo4_1.Text = reader0(77).ToString
                        QAChatRev.cbo4_2.Text = reader0(78).ToString
                        QAChatRev.cbo4_3.Text = reader0(79).ToString


                        QAChatRev.txt4_1.Text = reader0(86).ToString
                        QAChatRev.txt4_2.Text = reader0(87).ToString
                        QAChatRev.txt4_3.Text = reader0(88).ToString


                        QAChatRev.cbo5_1.Text = reader0(95).ToString
                        QAChatRev.cbo5_2.Text = reader0(96).ToString


                        QAChatRev.txt5_1.Text = reader0(104).ToString
                        QAChatRev.txt5_2.Text = reader0(105).ToString



                        QAChatRev.cbo6_1.Text = reader0(113).ToString
                        QAChatRev.cbo6_2.Text = reader0(114).ToString
                        QAChatRev.cbo6_3.Text = reader0(115).ToString

                        QAChatRev.txt6_1.Text = reader0(122).ToString
                        QAChatRev.txt6_2.Text = reader0(123).ToString
                        QAChatRev.txt6_3.Text = reader0(124).ToString

                        QAChatRev.cbo7_1.Text = reader0(131).ToString
                        QAChatRev.cbo7_2.Text = reader0(132).ToString
                        QAChatRev.cbo7_3.Text = reader0(133).ToString
                        QAChatRev.cbo7_4.Text = reader0(134).ToString
                        QAChatRev.cbo7_5.Text = reader0(135).ToString
                        QAChatRev.cbo7_6.Text = reader0(136).ToString

                        QAChatRev.txt7_1.Text = reader0(140).ToString
                        QAChatRev.txt7_2.Text = reader0(141).ToString
                        QAChatRev.txt7_3.Text = reader0(142).ToString
                        QAChatRev.txt7_4.Text = reader0(143).ToString
                        QAChatRev.txt7_5.Text = reader0(144).ToString
                        QAChatRev.txt7_6.Text = reader0(145).ToString




                        QAChatRev.txtQAScore.Text = reader0(149).ToString


                        QAChatRev.cboAF.Text = reader0(152).ToString


                        QAChatRev.txtOrignalAuditor.Text = reader0(153).ToString

                        QAChatRev.txtSupervisor.Text = reader0(155).ToString

                        QAChatRev.txtTCXScore.Text = reader0(156).ToString

                        QAChatRev.txtWeekNumber.Text = reader0(157).ToString


                        QAChatRev.lblrevMan.Text = lblQAauditor.Text

                        QAChatRev.lblQAauditor1.Text = lblQAauditor.Text



                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    End If






                End While
                reader0.Close()


            End Using

        End Using





    End Sub





    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click



        Try

            Me.Cursor = Cursors.AppStarting

            If txtSRPort.Text = "" Then

                MsgBox("Please enter a SR# or Contact ID in field", MessageBoxButtons.RetryCancel)

                Me.ActiveControl = txtSRPort

                Me.Cursor = Cursors.Hand


            Else



                If BackgroundWorker2.IsBusy = False Then

                    BackgroundWorker2.RunWorkerAsync()

                    '   PleaseWait.ShowDialog()

                    ' Using con0 = New System.Data.OleDb.OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb")


                    Using con0 = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")

                        ' Using con0 = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")



                        ' Dim SQL0 As String = "SELECT * FROM [QAMainDB] WHERE SR= ? AND Ctype= ?"

                        Dim SQL0 As String = "SELECT * FROM [QAMainDB] WHERE SR= @SR or ContactID= @ContactID"

                        '    Dim SQL0 As String = "SELECT * FROM [QAMainDB] WHERE SR= ? "

                        Using cmd0 As New SqlCommand(SQL0, con0)





                            cmd0.Parameters.AddWithValue("@SR", txtSRPort.Text)
                            cmd0.Parameters.AddWithValue("@ContactID", txtSRPort.Text)


                            '   cmd0.Parameters.AddWithValue("@p2", DataGridViewTextBoxColumn18.ToString)



                            con0.Open()



                            Dim reader0 As SqlDataReader

                            reader0 = cmd0.ExecuteReader()




                            While reader0.Read()





                                If txtAuditType.Text = "Call" Then




                                    QACallRev.lblGhostSR.Text = reader0(1).ToString()
                                    QACallRev.txtSR.Text = reader0(1).ToString()
                                    QACallRev.txtContactID.Text = reader0(2).ToString()
                                    '    QACallScorecard.txtContactID.Text = reader0(3).ToString() contact type


                                    'QACallRev.cboAgentName.Text = reader0(4).ToString()
                                    'QACallRev.cboTeamName.Text = reader0(5).ToString()

                                    QACallRev.txtAgentName.Text = reader0(4).ToString()
                                    QACallRev.txtTeamName.Text = reader0(5).ToString()


                                    QACallRev.dtpCondate.Text = reader0(6).ToString()
                                    QACallRev.txtOrderID.Text = reader0(7).ToString()
                                    QACallRev.txtQADate.Text = reader0(8).ToString()
                                    QACallRev.txtQACom.Text = reader0(9).ToString()
                                    QACallRev.txtQAAOO.Text = reader0(10).ToString()
                                    QACallRev.txtContactName.Text = reader0(11).ToString()
                                    QACallRev.txtAccountNum.Text = reader0(12).ToString()
                                    QACallRev.txtCompany.Text = reader0(13).ToString
                                    QACallRev.txtContactPhone.Text = reader0(14).ToString
                                    QACallRev.txtContactEmail.Text = reader0(15).ToString

                                    QACallRev.dtpReviewdate.Text = reader0(16).ToString
                                    QACallRev.lblrevMan.Text = reader0(17).ToString
                                    QACallRev.txtRevComments.Text = reader0(18).ToString
                                    QACallRev.txtDisputeScore.Text = reader0(19).ToString
                                    QACallRev.txtDisputerName.Text = reader0(20).ToString
                                    QACallRev.txtDisputeNotes.Text = reader0(21).ToString
                                    QACallRev.txtDisComment.Text = reader0(22).ToString


                                    QACallRev.cbo1_1.Text = reader0(23).ToString
                                    QACallRev.cbo1_2.Text = reader0(24).ToString
                                    QACallRev.cbo1_3.Text = reader0(25).ToString

                                    QACallRev.txt1_1.Text = reader0(32).ToString
                                    QACallRev.txt1_2.Text = reader0(33).ToString
                                    QACallRev.txt1_3.Text = reader0(34).ToString


                                    QACallRev.cbo2_1.Text = reader0(41).ToString
                                    QACallRev.txt2_1.Text = reader0(50).ToString



                                    QACallRev.cbo3_1.Text = reader0(59).ToString
                                    QACallRev.cbo3_2.Text = reader0(60).ToString
                                    QACallRev.cbo3_3.Text = reader0(61).ToString
                                    QACallRev.cbo3_4.Text = reader0(62).ToString
                                    QACallRev.cbo3_5.Text = reader0(63).ToString
                                    QACallRev.cbo3_6.Text = reader0(64).ToString
                                    QACallRev.cbo3_7.Text = reader0(65).ToString
                                    QACallRev.cbo3_8.Text = reader0(66).ToString


                                    QACallRev.txt3_1.Text = reader0(68).ToString
                                    QACallRev.txt3_2.Text = reader0(69).ToString
                                    QACallRev.txt3_3.Text = reader0(70).ToString
                                    QACallRev.txt3_4.Text = reader0(71).ToString
                                    QACallRev.txt3_5.Text = reader0(72).ToString
                                    QACallRev.txt3_6.Text = reader0(73).ToString
                                    QACallRev.txt3_7.Text = reader0(74).ToString
                                    QACallRev.txt3_8.Text = reader0(75).ToString


                                    QACallRev.Cbo4_1.Text = reader0(77).ToString
                                    QACallRev.cbo4_2.Text = reader0(78).ToString
                                    QACallRev.cbo4_3.Text = reader0(79).ToString

                                    QACallRev.txt4_1.Text = reader0(86).ToString
                                    QACallRev.txt4_2.Text = reader0(87).ToString
                                    QACallRev.txt4_3.Text = reader0(88).ToString


                                    QACallRev.cbo5_1.Text = reader0(95).ToString
                                    QACallRev.cbo5_2.Text = reader0(96).ToString


                                    QACallRev.txt5_1.Text = reader0(104).ToString
                                    QACallRev.txt5_2.Text = reader0(105).ToString


                                    QACallRev.cbo6_1.Text = reader0(113).ToString
                                    QACallRev.cbo6_2.Text = reader0(114).ToString
                                    QACallRev.cbo6_3.Text = reader0(115).ToString

                                    QACallRev.txt6_1.Text = reader0(122).ToString
                                    QACallRev.txt6_2.Text = reader0(123).ToString
                                    QACallRev.txt6_3.Text = reader0(124).ToString

                                    QACallRev.cbo7_1.Text = reader0(131).ToString
                                    QACallRev.cbo7_2.Text = reader0(132).ToString
                                    QACallRev.cbo7_3.Text = reader0(133).ToString
                                    QACallRev.cbo7_4.Text = reader0(134).ToString
                                    QACallRev.cbo7_5.Text = reader0(135).ToString
                                    QACallRev.cbo7_6.Text = reader0(136).ToString

                                    QACallRev.txt7_1.Text = reader0(140).ToString
                                    QACallRev.txt7_2.Text = reader0(141).ToString
                                    QACallRev.txt7_3.Text = reader0(142).ToString
                                    QACallRev.txt7_4.Text = reader0(143).ToString
                                    QACallRev.txt7_5.Text = reader0(144).ToString
                                    QACallRev.txt7_6.Text = reader0(145).ToString

                                    QACallRev.txtQAScore.Text = reader0(149).ToString


                                    QACallRev.cboAF.Text = reader0(152).ToString


                                    QACallRev.txtOrignalAuditor.Text = reader0(153).ToString
                                    QACallRev.txtSupervisor.Text = reader0(155).ToString
                                    QACallRev.txtTCXScore.Text = reader0(156).ToString
                                    QACallRev.txtWeekNumber.Text = reader0(157).ToString



                                    QACallRev.txt1_1a.Text = reader0(164).ToString
                                    QACallRev.txt1_2a.Text = reader0(165).ToString
                                    QACallRev.txt1_3a.Text = reader0(166).ToString

                                    QACallRev.txt2_1a.Text = reader0(167).ToString

                                    QACallRev.txt3_1a.Text = reader0(171).ToString
                                    QACallRev.txt3_2a.Text = reader0(172).ToString
                                    QACallRev.txt3_3a.Text = reader0(173).ToString
                                    QACallRev.txt3_4a.Text = reader0(174).ToString
                                    QACallRev.txt3_5a.Text = reader0(175).ToString
                                    QACallRev.txt3_6a.Text = reader0(176).ToString
                                    QACallRev.txt3_7a.Text = reader0(177).ToString
                                    QACallRev.txt3_8a.Text = reader0(178).ToString

                                    QACallRev.txt4_1a.Text = reader0(179).ToString
                                    QACallRev.txt4_2a.Text = reader0(180).ToString
                                    QACallRev.txt4_3a.Text = reader0(181).ToString


                                    QACallRev.txt5_1a.Text = reader0(183).ToString
                                    QACallRev.txt5_2a.Text = reader0(184).ToString

                                    QACallRev.txt6_1a.Text = reader0(189).ToString
                                    QACallRev.txt6_2a.Text = reader0(190).ToString
                                    QACallRev.txt6_3a.Text = reader0(191).ToString


                                    QACallRev.txt7_1a.Text = reader0(192).ToString
                                    QACallRev.txt7_2a.Text = reader0(193).ToString
                                    QACallRev.txt7_3a.Text = reader0(194).ToString
                                    QACallRev.txt7_4a.Text = reader0(195).ToString
                                    QACallRev.txt7_5a.Text = reader0(196).ToString
                                    QACallRev.txt7_6a.Text = reader0(197).ToString






                                    QACallRev.lblrevMan.Text = lblQAauditor.Text


                                    QACallRev.lblQAauditor1.Text = lblQAauditor.Text



                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                                ElseIf txtAuditType.Text = "Email" Then

                                    '  CallTimer.Enabled = True


                                    '   QAEmailRev.lblGhostSR.Text = reader0(1).ToString()
                                    QAEmailRev.txtSR.Text = reader0(1).ToString()
                                    QAEmailRev.txtContactID.Text = reader0(2).ToString()
                                    'QACallScorecard.txtContactID.Text = reader0(3).ToString() contact type

                                    'QAEmailRev.cboAgentName.Text = reader0(4).ToString()
                                    'QAEmailRev.cboTeamName.Text = reader0(5).ToString()

                                    QAEmailRev.txtAgentName.Text = reader0(4).ToString()
                                    QAEmailRev.txtTeamName.Text = reader0(5).ToString()




                                    QAEmailRev.dtpCondate.Text = reader0(6).ToString()
                                    QAEmailRev.txtOrderID.Text = reader0(7).ToString()
                                    QAEmailRev.txtQADate.Text = reader0(8).ToString()
                                    QAEmailRev.txtQACom.Text = reader0(9).ToString()
                                    QAEmailRev.txtQAAOO.Text = reader0(10).ToString()
                                    QAEmailRev.txtContactName.Text = reader0(11).ToString

                                    QAEmailRev.txtAccountNum.Text = reader0(12).ToString
                                    QAEmailRev.txtCompany.Text = reader0(13).ToString
                                    QAEmailRev.txtContactPhone.Text = reader0(14).ToString
                                    QAEmailRev.txtContactEmail.Text = reader0(15).ToString

                                    QAEmailRev.dtpReviewdate.Text = reader0(16).ToString
                                    QAEmailRev.lblrevMan.Text = reader0(17).ToString
                                    QAEmailRev.txtRevComments.Text = reader0(18).ToString
                                    QAEmailRev.txtDisputeScore.Text = reader0(19).ToString
                                    '   QAEmailRev.txtDisputerName.Text = reader0(20).ToString
                                    QAEmailRev.txtDisputeNotes.Text = reader0(21).ToString
                                    QAEmailRev.txtDisComment.Text = reader0(22).ToString




                                    QAEmailRev.cbo1_1.Text = reader0(23).ToString
                                    QAEmailRev.cbo1_2.Text = reader0(24).ToString
                                    QAEmailRev.cbo1_3.Text = reader0(25).ToString

                                    QAEmailRev.txt1_1.Text = reader0(32).ToString
                                    QAEmailRev.txt1_2.Text = reader0(33).ToString
                                    QAEmailRev.txt1_3.Text = reader0(34).ToString



                                    QAEmailRev.cbo2_1.Text = reader0(41).ToString
                                    QAEmailRev.cbo2_2.Text = reader0(42).ToString
                                    QAEmailRev.cbo2_3.Text = reader0(43).ToString
                                    QAEmailRev.cbo2_4.Text = reader0(44).ToString



                                    QAEmailRev.txt2_1.Text = reader0(50).ToString
                                    QAEmailRev.txt2_2.Text = reader0(51).ToString
                                    QAEmailRev.txt2_3.Text = reader0(52).ToString
                                    QAEmailRev.txt2_4.Text = reader0(53).ToString



                                    QAEmailRev.cbo3_1.Text = reader0(59).ToString
                                    QAEmailRev.cbo3_2.Text = reader0(60).ToString
                                    QAEmailRev.cbo3_3.Text = reader0(61).ToString
                                    QAEmailRev.cbo3_4.Text = reader0(62).ToString
                                    QAEmailRev.cbo3_5.Text = reader0(63).ToString




                                    QAEmailRev.txt3_1.Text = reader0(68).ToString
                                    QAEmailRev.txt3_2.Text = reader0(69).ToString
                                    QAEmailRev.txt3_3.Text = reader0(70).ToString
                                    QAEmailRev.txt3_4.Text = reader0(71).ToString
                                    QAEmailRev.txt3_5.Text = reader0(72).ToString






                                    QAEmailRev.cbo4_1.Text = reader0(77).ToString
                                    QAEmailRev.cbo4_2.Text = reader0(78).ToString
                                    QAEmailRev.cbo4_3.Text = reader0(79).ToString
                                    QAEmailRev.cbo4_4.Text = reader0(79).ToString



                                    QAEmailRev.txt4_1.Text = reader0(86).ToString
                                    QAEmailRev.txt4_2.Text = reader0(87).ToString
                                    QAEmailRev.txt4_3.Text = reader0(88).ToString
                                    QAEmailRev.txt4_4.Text = reader0(89).ToString


                                    QAEmailRev.cbo5_1.Text = reader0(95).ToString
                                    QAEmailRev.cbo5_2.Text = reader0(96).ToString
                                    QAEmailRev.cbo5_3.Text = reader0(97).ToString
                                    QAEmailRev.cbo5_4.Text = reader0(98).ToString
                                    QAEmailRev.cbo5_5.Text = reader0(99).ToString
                                    QAEmailRev.cbo5_6.Text = reader0(100).ToString

                                    QAEmailRev.txt5_1.Text = reader0(104).ToString
                                    QAEmailRev.txt5_2.Text = reader0(105).ToString
                                    QAEmailRev.txt5_3.Text = reader0(106).ToString
                                    QAEmailRev.txt5_4.Text = reader0(107).ToString
                                    QAEmailRev.txt5_5.Text = reader0(108).ToString
                                    QAEmailRev.txt5_6.Text = reader0(109).ToString

                                    QAEmailRev.txtQAScore.Text = reader0(149).ToString


                                    QAEmailRev.cboAF.Text = reader0(152).ToString


                                    QAEmailRev.txtOrignalAuditor.Text = reader0(153).ToString
                                    QAEmailRev.txtSupervisor.Text = reader0(155).ToString
                                    QAEmailRev.txtTCXScore.Text = reader0(156).ToString
                                    QAEmailRev.txtWeekNumber.Text = reader0(157).ToString


                                    QAEmailRev.txt1_1a.Text = reader0(164).ToString
                                    QAEmailRev.txt1_2a.Text = reader0(165).ToString
                                    QAEmailRev.txt1_3a.Text = reader0(166).ToString

                                    QAEmailRev.txt2_1a.Text = reader0(167).ToString
                                    QAEmailRev.txt2_2a.Text = reader0(168).ToString
                                    QAEmailRev.txt2_3a.Text = reader0(169).ToString
                                    QAEmailRev.txt2_4a.Text = reader0(170).ToString

                                    QAEmailRev.txt3_1a.Text = reader0(171).ToString
                                    QAEmailRev.txt3_2a.Text = reader0(172).ToString
                                    QAEmailRev.txt3_3a.Text = reader0(173).ToString
                                    QAEmailRev.txt3_4a.Text = reader0(174).ToString
                                    QAEmailRev.txt3_5a.Text = reader0(175).ToString


                                    QAEmailRev.txt4_1a.Text = reader0(179).ToString
                                    QAEmailRev.txt4_2a.Text = reader0(180).ToString
                                    QAEmailRev.txt4_3a.Text = reader0(181).ToString
                                    QAEmailRev.txt4_4a.Text = reader0(182).ToString



                                    QAEmailRev.txt5_1a.Text = reader0(183).ToString
                                    QAEmailRev.txt5_2a.Text = reader0(184).ToString
                                    QAEmailRev.txt5_3a.Text = reader0(185).ToString
                                    QAEmailRev.txt5_4a.Text = reader0(186).ToString
                                    QAEmailRev.txt5_5a.Text = reader0(187).ToString
                                    QAEmailRev.txt5_6a.Text = reader0(188).ToString


                                    QAEmailRev.lblrevMan.Text = lblQAauditor.Text

                                    QAEmailRev.lblQAauditor1.Text = lblQAauditor.Text

                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                ElseIf txtAuditType.Text = "Chat" Then


                                    QAChatRev.lblGhostSR.Text = reader0(1).ToString()
                                    QAChatRev.txtSR.Text = reader0(1).ToString()
                                    QAChatRev.txtContactID.Text = reader0(2).ToString()
                                    '    QACallScorecard.txtContactID.Text = reader0(3).ToString() contact type

                                    'QAChatRev.cboAgentName.Text = reader0(4).ToString()
                                    'QAChatRev.cboTeamName.Text = reader0(5).ToString()


                                    QAChatRev.txtAgentName.Text = reader0(4).ToString()
                                    QAChatRev.txtTeamName.Text = reader0(5).ToString()


                                    QAChatRev.dtpCondate.Text = reader0(6).ToString()
                                    QAChatRev.txtOrderID.Text = reader0(7).ToString()
                                    QAChatRev.txtQADate.Text = reader0(8).ToString()
                                    QAChatRev.txtQACom.Text = reader0(9).ToString()
                                    QAChatRev.txtQAAOO.Text = reader0(10).ToString()
                                    QAChatRev.txtContactName.Text = reader0(11).ToString()
                                    QAChatRev.txtAccountNum.Text = reader0(12).ToString()
                                    QAChatRev.txtCompany.Text = reader0(13).ToString
                                    QAChatRev.txtContactPhone.Text = reader0(14).ToString
                                    QAChatRev.txtContactEmail.Text = reader0(15).ToString

                                    QAChatRev.dtpReviewdate.Text = reader0(16).ToString
                                    QAChatRev.lblrevMan.Text = reader0(17).ToString
                                    QAChatRev.txtRevComments.Text = reader0(18).ToString
                                    QAChatRev.txtDisputeScore.Text = reader0(19).ToString
                                    'QAChatRev.txtDisputerName.Text = reader0(20).ToString
                                    QAChatRev.txtDisputeNotes.Text = reader0(21).ToString
                                    QAChatRev.txtDiscomment.Text = reader0(22).ToString


                                    QAChatRev.cbo1_1.Text = reader0(23).ToString
                                    QAChatRev.cbo1_2.Text = reader0(24).ToString

                                    QAChatRev.txt1_1.Text = reader0(32).ToString
                                    QAChatRev.txt1_2.Text = reader0(33).ToString


                                    QAChatRev.cbo2_1.Text = reader0(41).ToString
                                    QAChatRev.txt2_1.Text = reader0(50).ToString




                                    QAChatRev.cbo3_1.Text = reader0(59).ToString
                                    QAChatRev.cbo3_2.Text = reader0(60).ToString
                                    QAChatRev.cbo3_3.Text = reader0(61).ToString
                                    QAChatRev.cbo3_4.Text = reader0(62).ToString
                                    QAChatRev.cbo3_5.Text = reader0(63).ToString
                                    QAChatRev.cbo3_6.Text = reader0(64).ToString
                                    QAChatRev.cbo3_7.Text = reader0(65).ToString
                                    QAChatRev.cbo3_8.Text = reader0(66).ToString



                                    QAChatRev.txt3_1.Text = reader0(68).ToString
                                    QAChatRev.txt3_2.Text = reader0(69).ToString
                                    QAChatRev.txt3_3.Text = reader0(70).ToString
                                    QAChatRev.txt3_4.Text = reader0(71).ToString
                                    QAChatRev.txt3_5.Text = reader0(72).ToString
                                    QAChatRev.txt3_6.Text = reader0(73).ToString
                                    QAChatRev.txt3_7.Text = reader0(74).ToString
                                    QAChatRev.txt3_8.Text = reader0(75).ToString


                                    QAChatRev.Cbo4_1.Text = reader0(77).ToString
                                    QAChatRev.cbo4_2.Text = reader0(78).ToString
                                    QAChatRev.cbo4_3.Text = reader0(79).ToString


                                    QAChatRev.txt4_1.Text = reader0(86).ToString
                                    QAChatRev.txt4_2.Text = reader0(87).ToString
                                    QAChatRev.txt4_3.Text = reader0(88).ToString


                                    QAChatRev.cbo5_1.Text = reader0(95).ToString
                                    QAChatRev.cbo5_2.Text = reader0(96).ToString


                                    QAChatRev.txt5_1.Text = reader0(104).ToString
                                    QAChatRev.txt5_2.Text = reader0(105).ToString



                                    QAChatRev.cbo6_1.Text = reader0(113).ToString
                                    QAChatRev.cbo6_2.Text = reader0(114).ToString
                                    QAChatRev.cbo6_3.Text = reader0(115).ToString

                                    QAChatRev.txt6_1.Text = reader0(122).ToString
                                    QAChatRev.txt6_2.Text = reader0(123).ToString
                                    QAChatRev.txt6_3.Text = reader0(124).ToString

                                    QAChatRev.cbo7_1.Text = reader0(131).ToString
                                    QAChatRev.cbo7_2.Text = reader0(132).ToString
                                    QAChatRev.cbo7_3.Text = reader0(133).ToString
                                    QAChatRev.cbo7_4.Text = reader0(134).ToString
                                    QAChatRev.cbo7_5.Text = reader0(135).ToString
                                    QAChatRev.cbo7_6.Text = reader0(136).ToString

                                    QAChatRev.txt7_1.Text = reader0(140).ToString
                                    QAChatRev.txt7_2.Text = reader0(141).ToString
                                    QAChatRev.txt7_3.Text = reader0(142).ToString
                                    QAChatRev.txt7_4.Text = reader0(143).ToString
                                    QAChatRev.txt7_5.Text = reader0(144).ToString
                                    QAChatRev.txt7_6.Text = reader0(145).ToString




                                    QAChatRev.txtQAScore.Text = reader0(149).ToString


                                    QAChatRev.cboAF.Text = reader0(152).ToString


                                    QAChatRev.txtOrignalAuditor.Text = reader0(153).ToString

                                    QAChatRev.txtSupervisor.Text = reader0(155).ToString
                                    QAChatRev.txtTCXScore.Text = reader0(156).ToString
                                    QAChatRev.txtWeekNumber.Text = reader0(157).ToString



                                    QAChatRev.lblrevMan.Text = lblQAauditor.Text

                                    QAChatRev.lblQAauditor1.Text = lblQAauditor.Text



                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                End If






                            End While
                            reader0.Close()


                        End Using

                    End Using




                    ' Me.Cursor = Cursors.Hand


                End If


            End If








        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





    End Sub







    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click


        'Dim PendingTotal As Integer = 0
        'Dim counter As Integer

        'Dim temp As Integer



        'For i = 0 To (DataGridView2.Rows.Count - 1)


        '    If DataGridView2.Rows(i).Cells("Rev_Date").Value = "9/9/2020" Then


        '        PendingTotal = Integer.Parse(DataGridView2.Rows(counter).Cells("Rev_Date").Value.ToString.Count(), temp)

        '        PendingTotal += temp


        '    End If


        'Next

        'For i = 0 To (DataGridView2.Rows.Count - 1)


        '    If DataGridView2.Rows(i).Cells("Rev_Date").Value = "9/9/2020" Then



        '        Int32.TryParse(DataGridView2.Rows(i).Cells("Rev_Date").Value.ToString(), temp)

        '        PendingTotal += temp

        '        counter += 1

        '        lblPenReview.Text = counter
        '    End If

        'Next



        'Dim Value As String = "9/9/2020"
        'Dim ColumnName As String = "Rev_Date"



        'lblPenReview.Text = DataGridView2 _
        '    .Rows.Cast(Of DataGridViewRow) _
        '    .Where(Function(row)
        '               Return (Not IsDBNull(row.Cells(ColumnName).Value)) AndAlso (row.Cells(ColumnName).Value = Value)
        '           End Function) _
        '    .Select(Function(row) row.Cells(ColumnName).Value).Count.ToString










        '  Dim







        'For i As Integer = 0 To (DataGridView2.Rows.Count - 1)


        '    Dim qad = CDate(DataGridView2.Rows(i).Cells("Rev_Date").Value.ToString)

        '    Dim expiredate = qad.AddDays(7)

        '    Dim Threedaynoti = qad.AddDays(3)


        '    If qad < Now Then



        '        counter2 += 1


        '    ElseIf Now > Threedaynoti And (DataGridView2.Rows(i).Cells("Rev_Date").Value = "9/9/2020") Then


        '        counter3 += 1



        '        End If


        '    Next


        '    lblTotalRev.Text = counter2




        For i = 0 To (DataGridView2.Rows.Count - 1)



            Dim qad = CDate(DataGridView2.Rows(i).Cells("Rev_Date").Value.ToString)

            Dim expiredate = qad.AddDays(7)

            Dim Threedaynoti = qad.AddDays(3)









            If Now > expiredate And qad = "9/9/2020" Then

                counter3 += 1


            End If


        Next



    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles CallTimer.Tick








    End Sub



    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting

        Try


            Dim Now = DateTime.Now

            ' Dim QaDate = CDate(DataGridView2.Rows(i).Cells("QA_Date").Value.ToString)


            Dim QaDate = CDate(DataGridView2.Rows(e.RowIndex).Cells("QA_Date").Value.ToString)



            Dim expiredate = QaDate.AddDays(7)

            Dim Threedaynoti = QaDate.AddDays(3)


            '  Dim




            If Now > expiredate And DataGridView2.Rows(e.RowIndex).Cells("Rev_Date").Value = "9/9/2020" Then

                e.CellStyle.BackColor = Color.Tomato


            ElseIf Now > Threedaynoti And DataGridView2.Rows(e.RowIndex).Cells("Rev_Date").Value = "9/9/2020" Then


                e.CellStyle.BackColor = Color.Yellow

            ElseIf DataGridView2.Rows(e.RowIndex).Cells("Rev_Date").Value = "9/9/2020" Then


                '  e.CellStyle.BackColor = Color.LawnGreen

                DataGridView2.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LawnGreen



            ElseIf DataGridView2.Rows(e.RowIndex).Cells("Rev_Date").Value IsNot "9/9/2020" Then


                e.CellStyle.BackColor = Color.White


            ElseIf DataGridView2.Rows(e.RowIndex).Cells("Rev_Date").Value = Nothing Then





            End If







        Catch ex As System.ArgumentOutOfRangeException


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try




    End Sub

    Private Sub RefreshExcel_Tick(sender As Object, e As EventArgs) Handles RefreshExcel.Tick

        RefreshExcel.Enabled = False


        'xlsApp.Workbooks("Testbook.xlsm").Close(SaveChanges:=False)


        ' SpreadsheetControl1.LoadDocument("P: \SPC\QA\QAExcell.xlsm", DocumentFormat.Xlsm)

        ' SpreadsheetControl1.LoadDocument("C:\Users\playe\Desktop\QA\QAExcellMaped.xlsm", DocumentFormat.Xlsm)


        ' MsgBox("Updated")


        RefreshExcel.Enabled = False


        RefreshExcel2.Enabled = True

    End Sub


    Private Sub Timer1_Tick_1(sender As Object, e As EventArgs) Handles Timer1.Tick


        Timer1.Enabled = False


        RowCounting()

        Timer1.Enabled = False



    End Sub

    Private Sub DataGridView2_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView2.SelectionChanged












    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles btnImportScorecard.Click


        Try


            Dim myStream As Stream = Nothing
            Dim openFileDialog1 As New OpenFileDialog()

            openFileDialog1.InitialDirectory = lblMDrive.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA2\"

            '  openFileDialog1.InitialDirectory = "C:\Users\playe\desktop\qa2"



            'openFileDialog1.Filter = "Excel |*.xlsx"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True


            If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                myStream = openFileDialog1.OpenFile()

                Process.Start(openFileDialog1.FileName)




            End If




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try






    End Sub



    Private Sub OpenFileDialog1_FileOk(sender As Object, e As CancelEventArgs) Handles OpenFileDialog1.FileOk


        'Dim fileDialog As OpenFileDialog = TryCast(sender, OpenFileDialog)

        'Dim selectedFile As String = fileDialog.FileName

        'If String.IsNullOrEmpty(selectedFile) OrElse selectedFile.Contains(".lnk") Then

        '    MessageBox.Show("Please select a valid Excel File")

        '    e.Cancel = True

        'End If
        'Return





    End Sub




    Private Sub BackgroundWorker2_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker2.DoWork

        Try

            For i = 0 To 100

                System.Threading.Thread.Sleep(69)


                Me.BackgroundWorker2.ReportProgress(i)

                '    lblprogr.Text = i.ToString

                i = i
            Next







        Catch ex As Exception



            MsgBox(ex.Message)

        End Try






    End Sub

    Private Sub BackgroundWorker2_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged


        ProgressBar1.Value = e.ProgressPercentage


    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted



        Try

            OpenCard.Enabled = True


            Me.Cursor = Cursors.Hand





            '  PleaseWait.Close()

            'If txtAuditType.Text = "Call" Then


            '    QACallRev.Show()


            'ElseIf txtAuditType.Text = "Email" Then


            '    QAEmailRev.Show()


            'ElseIf txtAuditType.Text = "Chat" Then


            '    QAChatRev.Show()




            'End If




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try






    End Sub

    Private Sub BackgroundWorker3_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker3.DoWork




    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click

        Try


            Dim AverageCalc As Double

            Dim average As Double

            Dim AVGcounter = 0

            Dim result As Double

            For i = 0 To (DataGridView2.Rows.Count - 1)



                If Double.TryParse(DataGridView2.Rows(i).Cells("DataGridViewTextBoxColumn23").Value.ToString(), average) Then




                    AverageCalc += average

                    AVGcounter += 1

                End If




                result = AverageCalc / rowCount

            Next


            ' lblQaAvg.Text = Format(Val(result.ToString()), "0.00")





        Catch ex As Exception



            MsgBox(ex.Message)

        End Try





    End Sub

    Private Sub cboSuperAgentBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSuperAgentBox.SelectedIndexChanged




        Try






            Me.Cursor = Cursors.AppStarting

            txtAgentTeam.Text = "Please wait, Loading.."

            BackgroundWorker6.RunWorkerAsync()


        Catch ex As Exception


            MsgBox(ex.Message)



        End Try

    End Sub

    Private Sub DataGridView2_Sorted(sender As Object, e As EventArgs) Handles DataGridView2.Sorted


        RowCounting()


    End Sub

    Private Sub PageLoadTimer_Tick(sender As Object, e As EventArgs) Handles PageLoadTimer.Tick


        PageLoadTimer.Enabled = False


        PleaseWait.Hide()

        PleaseWait.Hide()

        PageLoadTimer.Enabled = False

        Me.Cursor = Cursors.Hand

    End Sub

    Private Sub Button12_Click_1(sender As Object, e As EventArgs)




        If txtAuditType.Text = "Call" Then


            QACallRev.Show()

        ElseIf txtAuditType.Text = "Email" Then


            QAEmailRev.Show()


        ElseIf txtAuditType.Text = "Chat" Then


            QAChatRev.Show()



        End If





    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click



        TabControl1.TabPages.Insert(0, Tab3)

        '  QACallRev.Show()

        ' QAMainDBBindingSource1.Filter = "[Rev-Manager] like '%" & lblQAauditor.Text & "%'" And "[QA-Agent] Like '%" & txtSearchBox.Text & "'"

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick


        PleaseWait.Close()


        '  PleaseWait.Hide()

        QACallRev.ShowDialog()

        Timer2.Enabled = False

    End Sub

    Private Sub cboAgentName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAgentName.SelectedIndexChanged

        Try

            Me.Cursor = Cursors.AppStarting

            txtAgentTeam.Text = "Please wait, Loading.."

            BackgroundWorker5.RunWorkerAsync()



        Catch ex As Exception



            MsgBox(ex.Message)



        End Try




    End Sub

    Private Sub combofiller_Tick(sender As Object, e As EventArgs)





    End Sub

    Private Sub BackgroundWorker4_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker4.DoWork

        Fillcombo1()



    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted



        cboAgentName.Text = "Agent Name"

        Me.Cursor = Cursors.Hand


    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted



        'readertemp8.Close()
        'contemp8.Close()





    End Sub

    Private Sub BackgroundWorker5_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker5.DoWork


        Try


            QaSetupMod.connecttemp8()




            sqltemp8 = "SELECT * FROM [Agents] WHERE AgentName='" & cboAgentName.Text & " ' "




            Dim cmdtemp As New SqlCommand





            cmdtemp.CommandText = sqltemp8

            cmdtemp.Connection = contemp8



            readertemp8 = cmdtemp.ExecuteReader



            If (readertemp8.Read() = True) Then



                txtAgentTeam.Text = (readertemp8("Platform"))

                txtAgentEmail.Text = (readertemp8("AgentEmail"))



            End If



            cmdtemp.Dispose()


            readertemp8.Close()

            contemp8.Close()





            Me.Cursor = Cursors.Hand


        Catch ex As Exception



            MsgBox(ex.Message)

            MsgBox("Press 'Refresh Dashboard' button and try again")

        End Try





    End Sub

    Private Sub BackgroundWorker6_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker6.DoWork


        Try


            QaSetupMod.connecttemp8()



            Me.Cursor = Cursors.AppStarting

            sqltemp8 = "SELECT * FROM [Agents] WHERE AgentName='" & cboSuperAgentBox.Text & " ' "




            Dim cmdtemp As New SqlCommand





            cmdtemp.CommandText = sqltemp8

            cmdtemp.Connection = contemp8



            readertemp8 = cmdtemp.ExecuteReader



            If (readertemp8.Read() = True) Then



                txtAgentTeam.Text = (readertemp8("Platform"))

                txtAgentEmail.Text = (readertemp8("AgentEmail"))



            End If



            cmdtemp.Dispose()

            readertemp8.Close()

            contemp8.Close()



        Catch ex As Exception



            MsgBox(ex.Message)


        End Try




    End Sub

    Private Sub BackgroundWorker6_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted

        Me.Cursor = Cursors.Hand


        'readertemp8.Close()
        'contemp8.Close()



    End Sub

    Private Sub OpenCard_Tick(sender As Object, e As EventArgs) Handles OpenCard.Tick


        If txtAuditType.Text = "Call" Then


            QACallRev.Show()


        ElseIf txtAuditType.Text = "Email" Then


            QAEmailRev.Show()


        ElseIf txtAuditType.Text = "Chat" Then


            QAChatRev.Show()




        End If

        OpenCard.Enabled = False


    End Sub

    Private Sub MainLoader_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles MainLoader.ProgressChanged

    End Sub

    Private Sub MainLoader_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles MainLoader.RunWorkerCompleted

    End Sub

    Private Sub MainLoader_DoWork(sender As Object, e As DoWorkEventArgs) Handles MainLoader.DoWork

        Me.QaMainDBTableAdapter7.Fill(Me.QADBDataSet6.QAMainDB)


    End Sub

    Private Sub lblGreen_Click(sender As Object, e As EventArgs) Handles lblGreen.Click

        Try

            '   Bindingsource7.Filter = DefaultBackColor = Color.LawnGreen



            ''  Dim currencyManager1 As CurrencyManager = DirectCast(BindingContext(DataGridView2.DataSource), CurrencyManager)



            '''    For Each r As DataGridViewRow In DataGridView2.Rows

            '' If Not r.IsNewRow Then


            ' currencyManager1.SuspendBinding()


            ''If Not r.DefaultCellStyle.BackColor <> Color.LawnGreen Then


            'If r.DefaultCellStyle.BackColor <> r.DefaultCellStyle.BackColor = Color.LawnGreen Then




            ''DataGridView2.Rows.Remove(r)


            '  r.Visible = False



            '  currencyManager1.ResumeBinding()


            '''End If

            '''End If

            '''    Next




            Dim currencyManager1 As CurrencyManager = DirectCast(BindingContext(DataGridView2.DataSource), CurrencyManager)

            Dim i As Integer

            For i = 0 To (DataGridView2.Rows.Count - 1)

                '   DataGridView2.Rows(i).Visible = True

                If Not DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.LawnGreen Then


                    '  If DataGridView2.Rows(i).DefaultCellStyle.BackColor Not Color.LawnGreen Then



                    ''    currencyManager1.SuspendBinding()



                    ' DataGridView2.CurrentCell = Nothing

                    ' DataGridView2.Rows(i).Visible = False


                    'DataGridView2.Rows.hide(DataGridView2.Rows(i))


                    'DataGridView2.Rows.Remove(DataGridView2)

                    ''    currencyManager1.ResumeBinding()


                End If




            Next



            ' RowCounting()







        Catch ex As Exception



            MsgBox(ex.Message)


        End Try

    End Sub


    Private Sub Button12_Click_2(sender As Object, e As EventArgs) Handles Button12.Click




        Try

            SplashScreenManager1.ShowWaitForm()

            Annoucer()
            RefreshDatagrid()


        Catch ex As Exception


            MsgBox(ex.Message)



        End Try






    End Sub


    Public Sub RefreshDatagrid()





        DateTimePicker4.Value = Today
        DateTimePicker5.Value = Today



        ProgressBar1.Value = 0


        If lblDeciderDash.Text = "Admin" Then

            Me.Cursor = Cursors.AppStarting



            refreshDB()



        ElseIf lblDeciderDash.Text = "QaAuditor" Then


            Me.Cursor = Cursors.AppStarting



            refreshDB()


        ElseIf lblDeciderDash.Text = "Supervisor" Then


            Me.Cursor = Cursors.AppStarting



            refreshDB2()


        ElseIf lblDeciderDash.Text = "TeamLead" Then


            Me.Cursor = Cursors.AppStarting



            refreshDB2()



        ElseIf lblDeciderDash.Text = "GOCSupervisor" Then


            Me.Cursor = Cursors.AppStarting



            refreshDB2()

        ElseIf lblDeciderDash.Text = "GOCTeamLead" Then


            Me.Cursor = Cursors.AppStarting



            refreshDB2()


        End If





    End Sub





    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles txtSRPort3.Click


        Try

            Me.Cursor = Cursors.AppStarting

            SplashScreenManager1.ShowWaitForm()


            If txtSRPort2.Text = "" Then

                MsgBox("Please enter a SR# or Contact ID in field", MessageBoxButtons.RetryCancel)

                Me.ActiveControl = txtSRPort2

                Me.Cursor = Cursors.Hand


            Else



                If BackgroundWorker7.IsBusy = False Then

                    BackgroundWorker7.RunWorkerAsync()

                    '   PleaseWait.ShowDialog()

                    ' Using con0 = New System.Data.OleDb.OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb")


                    Using con00 = New System.Data.SqlClient.SqlConnection("Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")

                        ' Using con0 = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb")



                        ' Dim SQL0 As String = "SELECT * FROM [QAMainDB] WHERE SR= ? AND Ctype= ?"

                        Dim SQL00 As String = "SELECT * FROM [QAMainDB] WHERE SR= @SR or ContactID= @ContactID"

                        '    Dim SQL0 As String = "SELECT * FROM [QAMainDB] WHERE SR= ? "

                        Using cmd00 As New SqlCommand(SQL00, con00)





                            cmd00.Parameters.AddWithValue("@SR", txtSRPort2.Text)
                            cmd00.Parameters.AddWithValue("@ContactID", txtSRPort2.Text)


                            '   cmd0.Parameters.AddWithValue("@p2", DataGridViewTextBoxColumn18.ToString)



                            con00.Open()



                            Dim reader00 As SqlDataReader

                            reader00 = cmd00.ExecuteReader()




                            While reader00.Read()


                                txtAuditTypeTest.Text = reader00(3).ToString()


                                If txtAuditType2.Text = "Call" And txtAuditTypeTest.Text = "Call" Then


                                    QACallRev.lblregion.Text = lblRegion.Text

                                    QACallRev.lblDeciderX2.Text = lblDeciderX2.Text

                                    QACallRev.lblOLDID.Text = reader00(201).ToString()
                                    QACallRev.txtSRType.Text = reader00(202).ToString
                                    QACallRev.txtPendingReview.Text = reader00(203).ToString()



                                    QACallRev.lblcurrentUser.Text = lblQAAuditor3.Text
                                    QACallRev.lblUserType.Text = lblDeciderDash.Text




                                    QACallRev.lblUserEmail.Text = lblUserEmail.Text

                                    QACallRev.lblGhostID.Text = reader00(0).ToString()
                                    QACallRev.lblGhostSR.Text = reader00(1).ToString()
                                    QACallRev.txtSR.Text = reader00(1).ToString()
                                    QACallRev.txtContactID.Text = reader00(2).ToString()
                                    '    QACallScorecard.txtContactID.Text = reader00(3).ToString() contact type


                                    'QACallRev.cboAgentName.Text = reader00(4).ToString()
                                    'QACallRev.cboTeamName.Text = reader00(5).ToString()

                                    QACallRev.txtAgentName.Text = reader00(4).ToString()
                                    QACallRev.txtTeamName.Text = reader00(5).ToString()


                                    QACallRev.dtpCondate.Text = reader00(6).ToString()
                                    QACallRev.txtOrderID.Text = reader00(7).ToString()
                                    QACallRev.txtQADate.Text = reader00(8).ToString()
                                    QACallRev.txtQACom.Text = reader00(9).ToString()
                                    QACallRev.txtQAAOO.Text = reader00(10).ToString()
                                    QACallRev.txtContactName.Text = reader00(11).ToString()
                                    QACallRev.txtAccountNum.Text = reader00(12).ToString()
                                    QACallRev.txtCompany.Text = reader00(13).ToString
                                    QACallRev.txtContactPhone.Text = reader00(14).ToString
                                    QACallRev.txtContactEmail.Text = reader00(15).ToString

                                    QACallRev.dtpReviewdate.Text = reader00(16).ToString
                                    QACallRev.lblrevMan.Text = reader00(17).ToString
                                    QACallRev.txtRevComments.Text = reader00(18).ToString

                                    QACallRev.txtDisputerName.Text = reader00(20).ToString
                                    QACallRev.txtDisputeNotes.Text = reader00(21).ToString
                                    QACallRev.txtDisComment.Text = reader00(22).ToString


                                    QACallRev.cbo1_1.Text = reader00(23).ToString
                                    QACallRev.cbo1_2.Text = reader00(24).ToString
                                    QACallRev.cbo1_3.Text = reader00(25).ToString

                                    QACallRev.txt1_1.Text = reader00(32).ToString
                                    QACallRev.txt1_2.Text = reader00(33).ToString
                                    QACallRev.txt1_3.Text = reader00(34).ToString


                                    QACallRev.cbo2_1.Text = reader00(41).ToString
                                    QACallRev.txt2_1.Text = reader00(50).ToString



                                    QACallRev.cbo3_1.Text = reader00(59).ToString
                                    QACallRev.cbo3_2.Text = reader00(60).ToString
                                    QACallRev.cbo3_3.Text = reader00(61).ToString
                                    QACallRev.cbo3_4.Text = reader00(62).ToString
                                    QACallRev.cbo3_5.Text = reader00(63).ToString
                                    QACallRev.cbo3_6.Text = reader00(64).ToString
                                    QACallRev.cbo3_7.Text = reader00(65).ToString
                                    QACallRev.cbo3_8.Text = reader00(66).ToString


                                    QACallRev.txt3_1.Text = reader00(68).ToString
                                    QACallRev.txt3_2.Text = reader00(69).ToString
                                    QACallRev.txt3_3.Text = reader00(70).ToString
                                    QACallRev.txt3_4.Text = reader00(71).ToString
                                    QACallRev.txt3_5.Text = reader00(72).ToString
                                    QACallRev.txt3_6.Text = reader00(73).ToString
                                    QACallRev.txt3_7.Text = reader00(74).ToString
                                    QACallRev.txt3_8.Text = reader00(75).ToString


                                    QACallRev.Cbo4_1.Text = reader00(77).ToString
                                    QACallRev.cbo4_2.Text = reader00(78).ToString
                                    QACallRev.cbo4_3.Text = reader00(79).ToString

                                    QACallRev.txt4_1.Text = reader00(86).ToString
                                    QACallRev.txt4_2.Text = reader00(87).ToString
                                    QACallRev.txt4_3.Text = reader00(88).ToString


                                    QACallRev.cbo5_1.Text = reader00(95).ToString
                                    QACallRev.cbo5_2.Text = reader00(96).ToString


                                    QACallRev.txt5_1.Text = reader00(104).ToString
                                    QACallRev.txt5_2.Text = reader00(105).ToString


                                    QACallRev.cbo6_1.Text = reader00(113).ToString
                                    QACallRev.cbo6_2.Text = reader00(114).ToString
                                    QACallRev.cbo6_3.Text = reader00(115).ToString

                                    QACallRev.txt6_1.Text = reader00(122).ToString
                                    QACallRev.txt6_2.Text = reader00(123).ToString
                                    QACallRev.txt6_3.Text = reader00(124).ToString

                                    QACallRev.cbo7_1.Text = reader00(131).ToString
                                    QACallRev.cbo7_2.Text = reader00(132).ToString
                                    QACallRev.cbo7_3.Text = reader00(133).ToString
                                    QACallRev.cbo7_4.Text = reader00(134).ToString
                                    QACallRev.cbo7_5.Text = reader00(135).ToString
                                    QACallRev.cbo7_6.Text = reader00(136).ToString

                                    QACallRev.txt7_1.Text = reader00(140).ToString
                                    QACallRev.txt7_2.Text = reader00(141).ToString
                                    QACallRev.txt7_3.Text = reader00(142).ToString
                                    QACallRev.txt7_4.Text = reader00(143).ToString
                                    QACallRev.txt7_5.Text = reader00(144).ToString
                                    QACallRev.txt7_6.Text = reader00(145).ToString

                                    QACallRev.txtQAScore.Text = reader00(149).ToString


                                    QACallRev.cboAF.Text = reader00(152).ToString


                                    QACallRev.txtOrignalAuditor.Text = reader00(153).ToString
                                    QACallRev.txtSupervisor.Text = reader00(155).ToString

                                    QACallRev.txtTCXScore.Text = reader00(156).ToString
                                    QACallRev.txtWeekNumber.Text = reader00(157).ToString

                                    QACallRev.txtRevDate.Text = reader00(16).ToString


                                    QACallRev.txtMonth.Text = reader00(161).ToString



                                    QACallRev.txt1_1a.Text = reader00(162).ToString
                                    QACallRev.txt1_2a.Text = reader00(163).ToString
                                    QACallRev.txt1_3a.Text = reader00(164).ToString

                                    QACallRev.txt2_1a.Text = reader00(165).ToString

                                    QACallRev.txt3_1a.Text = reader00(169).ToString
                                    QACallRev.txt3_2a.Text = reader00(170).ToString
                                    QACallRev.txt3_3a.Text = reader00(171).ToString
                                    QACallRev.txt3_4a.Text = reader00(172).ToString
                                    QACallRev.txt3_5a.Text = reader00(173).ToString
                                    QACallRev.txt3_6a.Text = reader00(174).ToString
                                    QACallRev.txt3_7a.Text = reader00(175).ToString
                                    QACallRev.txt3_8a.Text = reader00(176).ToString

                                    QACallRev.txt4_1a.Text = reader00(177).ToString
                                    QACallRev.txt4_2a.Text = reader00(178).ToString
                                    QACallRev.txt4_3a.Text = reader00(179).ToString


                                    QACallRev.txt5_1a.Text = reader00(181).ToString
                                    QACallRev.txt5_2a.Text = reader00(182).ToString

                                    QACallRev.txt6_1a.Text = reader00(187).ToString
                                    QACallRev.txt6_2a.Text = reader00(188).ToString
                                    QACallRev.txt6_3a.Text = reader00(189).ToString


                                    QACallRev.txt7_1a.Text = reader00(190).ToString
                                    QACallRev.txt7_2a.Text = reader00(191).ToString
                                    QACallRev.txt7_3a.Text = reader00(192).ToString
                                    QACallRev.txt7_4a.Text = reader00(193).ToString
                                    QACallRev.txt7_5a.Text = reader00(194).ToString
                                    QACallRev.txt7_6a.Text = reader00(195).ToString


                                    QACallRev.txtEditedQA.Text = reader00(159).ToString
                                    QACallRev.txtDisputedQA.Text = reader00(160).ToString


                                    QACallRev.txtDisApp.Text = reader00(154).ToString
                                    QACallRev.txtDisputeScore.Text = reader00(19).ToString
                                    QACallRev.txtDisputedTCXScore.Text = reader00(204).ToString

                                    QACallRev.txtghostAFreason.Text = reader00(152).ToString


                                    '  QACallRev.lblrevMan.Text = lblQAauditor.Text


                                    QACallRev.lblQAauditor1.Text = lblQAauditor.Text

                                    QACallRev.lblEmailPassword.Text = lblEmailPassword.Text

                                    QACallRev.txtCSATScore.Text = reader00(208).ToString

                                    QACallRev.cboCSAT1.Text = reader00(209).ToString
                                    QACallRev.cboCSAT2.Text = reader00(210).ToString
                                    QACallRev.cboCSAT3.Text = reader00(211).ToString
                                    QACallRev.cboCSAT4.Text = reader00(212).ToString
                                    QACallRev.cboCSAT5.Text = reader00(213).ToString
                                    QACallRev.cboCSAT6.Text = reader00(214).ToString

                                    '   QACallRev.txtDisApp.Text = reader00(154).ToString

                                    QACallRev.txt1_1b.Text = reader00(219).ToString
                                    QACallRev.txt1_2b.Text = reader00(220).ToString
                                    QACallRev.txt1_3b.Text = reader00(221).ToString

                                    QACallRev.txt2_1b.Text = reader00(222).ToString

                                    QACallRev.txt3_1b.Text = reader00(226).ToString
                                    QACallRev.txt3_2b.Text = reader00(227).ToString
                                    QACallRev.txt3_3b.Text = reader00(228).ToString
                                    QACallRev.txt3_4b.Text = reader00(229).ToString
                                    QACallRev.txt3_5b.Text = reader00(230).ToString
                                    QACallRev.txt3_6b.Text = reader00(231).ToString
                                    QACallRev.txt3_7b.Text = reader00(232).ToString
                                    QACallRev.txt3_8b.Text = reader00(233).ToString

                                    QACallRev.txt4_1b.Text = reader00(234).ToString
                                    QACallRev.txt4_2b.Text = reader00(235).ToString
                                    QACallRev.txt4_3b.Text = reader00(236).ToString


                                    QACallRev.txt5_1b.Text = reader00(237).ToString
                                    QACallRev.txt5_2b.Text = reader00(238).ToString

                                    QACallRev.txt6_1b.Text = reader00(243).ToString
                                    QACallRev.txt6_2b.Text = reader00(244).ToString
                                    QACallRev.txt6_3b.Text = reader00(245).ToString


                                    QACallRev.txt7_1b.Text = reader00(246).ToString
                                    QACallRev.txt7_2b.Text = reader00(247).ToString
                                    QACallRev.txt7_3b.Text = reader00(248).ToString
                                    QACallRev.txt7_4b.Text = reader00(249).ToString
                                    QACallRev.txt7_5b.Text = reader00(250).ToString
                                    QACallRev.txt7_6b.Text = reader00(251).ToString








                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                                ElseIf txtAuditType2.Text = "Email" And txtAuditTypeTest.Text = "Email" Then

                                    '  CallTimer.Enabled = True

                                    QAEmailRev.lblregion.Text = lblRegion.Text

                                    QAEmailRev.lblDeciderX2.Text = lblDeciderX2.Text


                                    QAEmailRev.lblOLDID.Text = reader00(201).ToString
                                    QAEmailRev.txtSRType.Text = reader00(202).ToString
                                    QAEmailRev.txtPendingReview.Text = reader00(203).ToString


                                    QAEmailRev.lblcurrentUser.Text = lblQAAuditor3.Text
                                    QAEmailRev.lblUserType.Text = lblDeciderDash.Text



                                    QAEmailRev.lblUserEmail.Text = lblUserEmail.Text

                                    QAEmailRev.lblGhostID.Text = reader00(0).ToString()
                                    QAEmailRev.lblGhostSR.Text = reader00(1).ToString()
                                    QAEmailRev.txtSR.Text = reader00(1).ToString()
                                    QAEmailRev.txtContactID.Text = reader00(2).ToString()
                                    'QACallScorecard.txtContactID.Text = reader00(3).ToString() contact type

                                    'QAEmailRev.cboAgentName.Text = reader00(4).ToString()
                                    'QAEmailRev.cboTeamName.Text = reader00(5).ToString()

                                    QAEmailRev.txtAgentName.Text = reader00(4).ToString()
                                    QAEmailRev.txtTeamName.Text = reader00(5).ToString()




                                    QAEmailRev.dtpCondate.Text = reader00(6).ToString()
                                    QAEmailRev.txtOrderID.Text = reader00(7).ToString()
                                    QAEmailRev.txtQADate.Text = reader00(8).ToString()
                                    QAEmailRev.txtQACom.Text = reader00(9).ToString()
                                    QAEmailRev.txtQAAOO.Text = reader00(10).ToString()
                                    QAEmailRev.txtContactName.Text = reader00(11).ToString

                                    QAEmailRev.txtAccountNum.Text = reader00(12).ToString
                                    QAEmailRev.txtCompany.Text = reader00(13).ToString
                                    QAEmailRev.txtContactPhone.Text = reader00(14).ToString
                                    QAEmailRev.txtContactEmail.Text = reader00(15).ToString

                                    QAEmailRev.dtpReviewdate.Text = reader00(16).ToString
                                    QAEmailRev.lblrevMan.Text = reader00(17).ToString
                                    QAEmailRev.txtRevComments.Text = reader00(18).ToString
                                    ' QAEmailRev.txtDisApp.Text = reader00(19).ToString
                                    'QAEmailRev.txtDisputerName.Text = reader00(20).ToString
                                    QAEmailRev.txtDisputeNotes.Text = reader00(21).ToString
                                    QAEmailRev.txtDisComment.Text = reader00(22).ToString




                                    QAEmailRev.cbo1_1.Text = reader00(23).ToString
                                    QAEmailRev.cbo1_2.Text = reader00(24).ToString
                                    QAEmailRev.cbo1_3.Text = reader00(25).ToString

                                    QAEmailRev.txt1_1.Text = reader00(32).ToString
                                    QAEmailRev.txt1_2.Text = reader00(33).ToString
                                    QAEmailRev.txt1_3.Text = reader00(34).ToString



                                    QAEmailRev.cbo2_1.Text = reader00(41).ToString
                                    QAEmailRev.cbo2_2.Text = reader00(42).ToString
                                    QAEmailRev.cbo2_3.Text = reader00(43).ToString
                                    QAEmailRev.cbo2_4.Text = reader00(44).ToString



                                    QAEmailRev.txt2_1.Text = reader00(50).ToString
                                    QAEmailRev.txt2_2.Text = reader00(51).ToString
                                    QAEmailRev.txt2_3.Text = reader00(52).ToString
                                    QAEmailRev.txt2_4.Text = reader00(53).ToString



                                    QAEmailRev.cbo3_1.Text = reader00(59).ToString
                                    QAEmailRev.cbo3_2.Text = reader00(60).ToString
                                    QAEmailRev.cbo3_3.Text = reader00(61).ToString
                                    QAEmailRev.cbo3_4.Text = reader00(62).ToString
                                    QAEmailRev.cbo3_5.Text = reader00(63).ToString




                                    QAEmailRev.txt3_1.Text = reader00(68).ToString
                                    QAEmailRev.txt3_2.Text = reader00(69).ToString
                                    QAEmailRev.txt3_3.Text = reader00(70).ToString
                                    QAEmailRev.txt3_4.Text = reader00(71).ToString
                                    QAEmailRev.txt3_5.Text = reader00(72).ToString






                                    QAEmailRev.cbo4_1.Text = reader00(77).ToString
                                    QAEmailRev.cbo4_2.Text = reader00(78).ToString
                                    QAEmailRev.cbo4_3.Text = reader00(79).ToString
                                    QAEmailRev.cbo4_4.Text = reader00(80).ToString



                                    QAEmailRev.txt4_1.Text = reader00(86).ToString
                                    QAEmailRev.txt4_2.Text = reader00(87).ToString
                                    QAEmailRev.txt4_3.Text = reader00(88).ToString
                                    QAEmailRev.txt4_4.Text = reader00(89).ToString


                                    QAEmailRev.cbo5_1.Text = reader00(95).ToString
                                    QAEmailRev.cbo5_2.Text = reader00(96).ToString
                                    QAEmailRev.cbo5_3.Text = reader00(97).ToString
                                    QAEmailRev.cbo5_4.Text = reader00(98).ToString
                                    QAEmailRev.cbo5_5.Text = reader00(99).ToString
                                    QAEmailRev.cbo5_6.Text = reader00(100).ToString

                                    QAEmailRev.txt5_1.Text = reader00(104).ToString
                                    QAEmailRev.txt5_2.Text = reader00(105).ToString
                                    QAEmailRev.txt5_3.Text = reader00(106).ToString
                                    QAEmailRev.txt5_4.Text = reader00(107).ToString
                                    QAEmailRev.txt5_5.Text = reader00(108).ToString
                                    QAEmailRev.txt5_6.Text = reader00(109).ToString

                                    QAEmailRev.txtQAScore.Text = reader00(149).ToString


                                    QAEmailRev.cboAF.Text = reader00(152).ToString


                                    QAEmailRev.txtOrignalAuditor.Text = reader00(153).ToString
                                    QAEmailRev.txtSupervisor.Text = reader00(155).ToString
                                    QAEmailRev.txtTCXScore.Text = reader00(156).ToString
                                    QAEmailRev.txtWeekNumber.Text = reader00(157).ToString

                                    QAEmailRev.txtMonth.Text = reader00(161).ToString

                                    QAEmailRev.txt1_1a.Text = reader00(162).ToString
                                    QAEmailRev.txt1_2a.Text = reader00(163).ToString
                                    QAEmailRev.txt1_3a.Text = reader00(164).ToString

                                    QAEmailRev.txt2_1a.Text = reader00(165).ToString
                                    QAEmailRev.txt2_2a.Text = reader00(166).ToString
                                    QAEmailRev.txt2_3a.Text = reader00(167).ToString
                                    QAEmailRev.txt2_4a.Text = reader00(168).ToString

                                    QAEmailRev.txt3_1a.Text = reader00(169).ToString
                                    QAEmailRev.txt3_2a.Text = reader00(170).ToString
                                    QAEmailRev.txt3_3a.Text = reader00(171).ToString
                                    QAEmailRev.txt3_4a.Text = reader00(172).ToString
                                    QAEmailRev.txt3_5a.Text = reader00(173).ToString


                                    QAEmailRev.txt4_1a.Text = reader00(177).ToString
                                    QAEmailRev.txt4_2a.Text = reader00(178).ToString
                                    QAEmailRev.txt4_3a.Text = reader00(179).ToString
                                    QAEmailRev.txt4_4a.Text = reader00(180).ToString



                                    QAEmailRev.txt5_1a.Text = reader00(181).ToString
                                    QAEmailRev.txt5_2a.Text = reader00(182).ToString
                                    QAEmailRev.txt5_3a.Text = reader00(183).ToString
                                    QAEmailRev.txt5_4a.Text = reader00(184).ToString
                                    QAEmailRev.txt5_5a.Text = reader00(185).ToString
                                    QAEmailRev.txt5_6a.Text = reader00(186).ToString


                                    '     QAEmailRev.lblrevMan.Text = lblQAauditor.Text

                                    QAEmailRev.lblQAauditor1.Text = lblQAauditor.Text

                                    QAEmailRev.lblEmailPassword.Text = lblEmailPassword.Text


                                    QAEmailRev.txtEditedQA.Text = reader00(159).ToString
                                    QAEmailRev.txtDisputedQA.Text = reader00(160).ToString
                                    QAEmailRev.txtRevDate.Text = reader00(16).ToString


                                    ' QAEmailRev.txtDisApp.Text = reader00(154).ToString


                                    QAEmailRev.txtDisApp.Text = reader00(154).ToString
                                    QAEmailRev.txtDisputeScore.Text = reader00(19).ToString
                                    QAEmailRev.txtDisputedTCXScore.Text = reader00(204).ToString

                                    QAEmailRev.txtghostAFreason.Text = reader00(152).ToString


                                    QAEmailRev.txtCSATScore.Text = reader00(208).ToString

                                    QAEmailRev.cboCSAT1.Text = reader00(209).ToString
                                    QAEmailRev.cboCSAT2.Text = reader00(210).ToString
                                    QAEmailRev.cboCSAT3.Text = reader00(211).ToString
                                    QAEmailRev.cboCSAT4.Text = reader00(212).ToString
                                    QAEmailRev.cboCSAT5.Text = reader00(213).ToString
                                    QAEmailRev.cboCSAT6.Text = reader00(214).ToString




                                    QAEmailRev.txt1_1b.Text = reader00(219).ToString
                                    QAEmailRev.txt1_2b.Text = reader00(220).ToString
                                    QAEmailRev.txt1_3b.Text = reader00(221).ToString

                                    QAEmailRev.txt2_1b.Text = reader00(222).ToString
                                    QAEmailRev.txt2_2b.Text = reader00(223).ToString
                                    QAEmailRev.txt2_3b.Text = reader00(224).ToString
                                    QAEmailRev.txt2_4b.Text = reader00(225).ToString

                                    QAEmailRev.txt3_1b.Text = reader00(226).ToString
                                    QAEmailRev.txt3_2b.Text = reader00(227).ToString
                                    QAEmailRev.txt3_3b.Text = reader00(228).ToString
                                    QAEmailRev.txt3_4b.Text = reader00(229).ToString
                                    QAEmailRev.txt3_5b.Text = reader00(230).ToString


                                    QAEmailRev.txt4_1b.Text = reader00(234).ToString
                                    QAEmailRev.txt4_2b.Text = reader00(235).ToString
                                    QAEmailRev.txt4_3b.Text = reader00(236).ToString
                                    QAEmailRev.txt4_4b.Text = reader00(237).ToString

                                    QAEmailRev.txt5_1b.Text = reader00(237).ToString
                                    QAEmailRev.txt5_2b.Text = reader00(238).ToString
                                    QAEmailRev.txt5_3b.Text = reader00(239).ToString
                                    QAEmailRev.txt5_4b.Text = reader00(240).ToString
                                    QAEmailRev.txt5_5b.Text = reader00(241).ToString
                                    QAEmailRev.txt5_6b.Text = reader00(242).ToString




                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                ElseIf txtAuditType2.Text = "Chat" And txtAuditTypeTest.Text = "Chat" Then


                                    QAChatRev.lblregion.Text = lblRegion.Text
                                    QAChatRev.lblDeciderX2.Text = lblDeciderX2.Text



                                    QAChatRev.lblOLDID.Text = reader00(201).ToString
                                    QAChatRev.txtSRType.Text = reader00(202).ToString
                                    QAChatRev.txtPendingReview.Text = reader00(203).ToString


                                    QAChatRev.lblUserEmail.Text = lblUserEmail.Text


                                    QAChatRev.lblcurrentUser.Text = lblQAAuditor3.Text
                                    QAChatRev.lblUserType.Text = lblDeciderDash.Text


                                    QAChatRev.lblGhostID.Text = reader00(0).ToString()
                                    QAChatRev.lblGhostSR.Text = reader00(1).ToString()
                                    QAChatRev.txtSR.Text = reader00(1).ToString()
                                    QAChatRev.txtContactID.Text = reader00(2).ToString()
                                    '    QACallScorecard.txtContactID.Text = reader00(3).ToString() contact type

                                    'QAChatRev.cboAgentName.Text = reader00(4).ToString()
                                    'QAChatRev.cboTeamName.Text = reader00(5).ToString()


                                    QAChatRev.txtAgentName.Text = reader00(4).ToString()
                                    QAChatRev.txtTeamName.Text = reader00(5).ToString()


                                    QAChatRev.dtpCondate.Text = reader00(6).ToString()
                                    QAChatRev.txtOrderID.Text = reader00(7).ToString()
                                    QAChatRev.txtQADate.Text = reader00(8).ToString()
                                    QAChatRev.txtQACom.Text = reader00(9).ToString()
                                    QAChatRev.txtQAAOO.Text = reader00(10).ToString()
                                    QAChatRev.txtContactName.Text = reader00(11).ToString()
                                    QAChatRev.txtAccountNum.Text = reader00(12).ToString()
                                    QAChatRev.txtCompany.Text = reader00(13).ToString
                                    QAChatRev.txtContactPhone.Text = reader00(14).ToString
                                    QAChatRev.txtContactEmail.Text = reader00(15).ToString

                                    QAChatRev.dtpReviewdate.Text = reader00(16).ToString
                                    QAChatRev.lblrevMan.Text = reader00(17).ToString
                                    QAChatRev.txtRevComments.Text = reader00(18).ToString
                                    '     QAChatRev.txtDisApp.Text = reader00(19).ToString
                                    '   QAChatRev.txtDisputerName.Text = reader00(20).ToString
                                    QAChatRev.txtDisputeNotes.Text = reader00(21).ToString
                                    QAChatRev.txtDiscomment.Text = reader00(22).ToString


                                    QAChatRev.cbo1_1.Text = reader00(23).ToString
                                    QAChatRev.cbo1_2.Text = reader00(24).ToString

                                    QAChatRev.txt1_1.Text = reader00(32).ToString
                                    QAChatRev.txt1_2.Text = reader00(33).ToString


                                    QAChatRev.cbo2_1.Text = reader00(41).ToString
                                    QAChatRev.txt2_1.Text = reader00(50).ToString




                                    QAChatRev.cbo3_1.Text = reader00(59).ToString
                                    QAChatRev.cbo3_2.Text = reader00(60).ToString
                                    QAChatRev.cbo3_3.Text = reader00(61).ToString
                                    QAChatRev.cbo3_4.Text = reader00(62).ToString
                                    QAChatRev.cbo3_5.Text = reader00(63).ToString
                                    QAChatRev.cbo3_6.Text = reader00(64).ToString
                                    QAChatRev.cbo3_7.Text = reader00(65).ToString
                                    QAChatRev.cbo3_8.Text = reader00(66).ToString



                                    QAChatRev.txt3_1.Text = reader00(68).ToString
                                    QAChatRev.txt3_2.Text = reader00(69).ToString
                                    QAChatRev.txt3_3.Text = reader00(70).ToString
                                    QAChatRev.txt3_4.Text = reader00(71).ToString
                                    QAChatRev.txt3_5.Text = reader00(72).ToString
                                    QAChatRev.txt3_6.Text = reader00(73).ToString
                                    QAChatRev.txt3_7.Text = reader00(74).ToString
                                    QAChatRev.txt3_8.Text = reader00(75).ToString


                                    QAChatRev.Cbo4_1.Text = reader00(77).ToString
                                    QAChatRev.cbo4_2.Text = reader00(78).ToString
                                    QAChatRev.cbo4_3.Text = reader00(79).ToString


                                    QAChatRev.txt4_1.Text = reader00(86).ToString
                                    QAChatRev.txt4_2.Text = reader00(87).ToString
                                    QAChatRev.txt4_3.Text = reader00(88).ToString


                                    QAChatRev.cbo5_1.Text = reader00(95).ToString
                                    QAChatRev.cbo5_2.Text = reader00(96).ToString


                                    QAChatRev.txt5_1.Text = reader00(104).ToString
                                    QAChatRev.txt5_2.Text = reader00(105).ToString



                                    QAChatRev.cbo6_1.Text = reader00(113).ToString
                                    QAChatRev.cbo6_2.Text = reader00(114).ToString
                                    QAChatRev.cbo6_3.Text = reader00(115).ToString

                                    QAChatRev.txt6_1.Text = reader00(122).ToString
                                    QAChatRev.txt6_2.Text = reader00(123).ToString
                                    QAChatRev.txt6_3.Text = reader00(124).ToString

                                    QAChatRev.cbo7_1.Text = reader00(131).ToString
                                    QAChatRev.cbo7_2.Text = reader00(132).ToString
                                    QAChatRev.cbo7_3.Text = reader00(133).ToString
                                    QAChatRev.cbo7_4.Text = reader00(134).ToString
                                    QAChatRev.cbo7_5.Text = reader00(135).ToString
                                    QAChatRev.cbo7_6.Text = reader00(136).ToString

                                    QAChatRev.txt7_1.Text = reader00(140).ToString
                                    QAChatRev.txt7_2.Text = reader00(141).ToString
                                    QAChatRev.txt7_3.Text = reader00(142).ToString
                                    QAChatRev.txt7_4.Text = reader00(143).ToString
                                    QAChatRev.txt7_5.Text = reader00(144).ToString
                                    QAChatRev.txt7_6.Text = reader00(145).ToString




                                    QAChatRev.txtQAScore.Text = reader00(149).ToString


                                    QAChatRev.cboAF.Text = reader00(152).ToString


                                    QAChatRev.txtOrignalAuditor.Text = reader00(153).ToString

                                    QAChatRev.txtSupervisor.Text = reader00(155).ToString
                                    QAChatRev.txtTCXScore.Text = reader00(156).ToString
                                    QAChatRev.txtWeekNumber.Text = reader00(157).ToString

                                    QAChatRev.txtMonth.Text = reader00(161).ToString


                                    QAChatRev.txt1_1a.Text = reader00(162).ToString
                                    QAChatRev.txt1_2a.Text = reader00(163).ToString


                                    QAChatRev.txt2_1a.Text = reader00(165).ToString

                                    QAChatRev.txt3_1a.Text = reader00(169).ToString
                                    QAChatRev.txt3_2a.Text = reader00(170).ToString
                                    QAChatRev.txt3_3a.Text = reader00(171).ToString
                                    QAChatRev.txt3_4a.Text = reader00(172).ToString
                                    QAChatRev.txt3_5a.Text = reader00(173).ToString
                                    QAChatRev.txt3_6a.Text = reader00(174).ToString
                                    QAChatRev.txt3_7a.Text = reader00(175).ToString
                                    QAChatRev.txt3_8a.Text = reader00(176).ToString

                                    QAChatRev.txt4_1a.Text = reader00(177).ToString
                                    QAChatRev.txt4_2a.Text = reader00(178).ToString
                                    QAChatRev.txt4_3a.Text = reader00(179).ToString


                                    QAChatRev.txt5_1a.Text = reader00(181).ToString
                                    QAChatRev.txt5_2a.Text = reader00(182).ToString

                                    QAChatRev.txt6_1a.Text = reader00(187).ToString
                                    QAChatRev.txt6_2a.Text = reader00(188).ToString
                                    QAChatRev.txt6_3a.Text = reader00(189).ToString


                                    QAChatRev.txt7_1a.Text = reader00(190).ToString
                                    QAChatRev.txt7_2a.Text = reader00(191).ToString
                                    QAChatRev.txt7_3a.Text = reader00(192).ToString
                                    QAChatRev.txt7_4a.Text = reader00(193).ToString
                                    QAChatRev.txt7_5a.Text = reader00(194).ToString
                                    QAChatRev.txt7_6a.Text = reader00(195).ToString


                                    '   QAChatRev.lblrevMan.Text = lblQAauditor.Text

                                    QAChatRev.lblQAauditor1.Text = lblQAauditor.Text

                                    QAChatRev.lblEmailPassword.Text = lblEmailPassword.Text


                                    QAChatRev.txtEditedQA.Text = reader00(159).ToString
                                    QAChatRev.txtDisputedQA.Text = reader00(160).ToString
                                    QAChatRev.txtRevDate.Text = reader00(16).ToString


                                    QAChatRev.txtDisApp.Text = reader00(154).ToString
                                    QAChatRev.txtDisputeScore.Text = reader00(19).ToString
                                    QAChatRev.txtDisputedTCXScore.Text = reader00(204).ToString

                                    QAChatRev.txtghostAFreason.Text = reader00(152).ToString

                                    QAChatRev.txtCSATScore.Text = reader00(208).ToString

                                    QAChatRev.cboCSAT1.Text = reader00(209).ToString
                                    QAChatRev.cboCSAT2.Text = reader00(210).ToString
                                    QAChatRev.cboCSAT3.Text = reader00(211).ToString
                                    QAChatRev.cboCSAT4.Text = reader00(212).ToString
                                    QAChatRev.cboCSAT5.Text = reader00(213).ToString
                                    QAChatRev.cboCSAT6.Text = reader00(214).ToString

                                    'QAChatRev.txtDisApp.Text = reader00(154).ToString



                                    QAChatRev.txt1_1b.Text = reader00(219).ToString
                                    QAChatRev.txt1_2b.Text = reader00(220).ToString


                                    QAChatRev.txt2_1b.Text = reader00(222).ToString

                                    QAChatRev.txt3_1b.Text = reader00(226).ToString
                                    QAChatRev.txt3_2b.Text = reader00(227).ToString
                                    QAChatRev.txt3_3b.Text = reader00(228).ToString
                                    QAChatRev.txt3_4b.Text = reader00(229).ToString
                                    QAChatRev.txt3_5b.Text = reader00(230).ToString
                                    QAChatRev.txt3_6b.Text = reader00(231).ToString
                                    QAChatRev.txt3_7b.Text = reader00(232).ToString
                                    QAChatRev.txt3_8b.Text = reader00(233).ToString

                                    QAChatRev.txt4_1b.Text = reader00(234).ToString
                                    QAChatRev.txt4_2b.Text = reader00(235).ToString
                                    QAChatRev.txt4_3b.Text = reader00(236).ToString


                                    QAChatRev.txt5_1b.Text = reader00(237).ToString
                                    QAChatRev.txt5_2b.Text = reader00(238).ToString

                                    QAChatRev.txt6_1b.Text = reader00(243).ToString
                                    QAChatRev.txt6_2b.Text = reader00(244).ToString
                                    QAChatRev.txt6_3b.Text = reader00(245).ToString


                                    QAChatRev.txt7_1b.Text = reader00(246).ToString
                                    QAChatRev.txt7_2b.Text = reader00(247).ToString
                                    QAChatRev.txt7_3b.Text = reader00(248).ToString
                                    QAChatRev.txt7_4b.Text = reader00(249).ToString
                                    QAChatRev.txt7_5b.Text = reader00(250).ToString
                                    QAChatRev.txt7_6b.Text = reader00(251).ToString
                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                                ElseIf txtAuditType2.Text = "WOTC Inbound" And txtAuditTypeTest.Text = "WOTC Inbound" Then


                                    QAWOTCInboundRev.lblregion.Text = lblRegion.Text

                                        QAWOTCInboundRev.lblDeciderX2.Text = lblDeciderX2.Text

                                        QAWOTCInboundRev.lblOLDID.Text = reader00(201).ToString()
                                        QAWOTCInboundRev.txtSRType.Text = reader00(202).ToString
                                        QAWOTCInboundRev.txtPendingReview.Text = reader00(203).ToString()



                                        QAWOTCInboundRev.lblcurrentUser.Text = lblQAAuditor3.Text
                                        QAWOTCInboundRev.lblUserType.Text = lblDeciderDash.Text




                                        QAWOTCInboundRev.lblUserEmail.Text = lblUserEmail.Text


                                        '   QAWOTCInboundRev.lblGhostSR.Text = reader00(1).ToString()
                                        ' QAWOTCInboundRev.txtSR.Text = reader00(1).ToString()

                                        '    QACallScorecard.txtContactID.Text = reader00(3).ToString() contact type


                                        'QAWOTCInboundRev.cboAgentName.Text = reader00(4).ToString()
                                        'QAWOTCInboundRev.cboTeamName.Text = reader00(5).ToString()

                                        QAWOTCInboundRev.txtContactID.Text = reader00(2).ToString()
                                        QAWOTCInboundRev.lblGhostID.Text = reader00(0).ToString()


                                        QAWOTCInboundRev.txtAgentName.Text = reader00(4).ToString()
                                        QAWOTCInboundRev.txtTeamName.Text = reader00(5).ToString()


                                        QAWOTCInboundRev.dtpCondate.Text = reader00(6).ToString()
                                        QAWOTCInboundRev.txtOrderID.Text = reader00(7).ToString()
                                        QAWOTCInboundRev.txtQADate.Text = reader00(8).ToString()
                                        QAWOTCInboundRev.txtQACom.Text = reader00(9).ToString()
                                        QAWOTCInboundRev.txtQAAOO.Text = reader00(10).ToString()
                                        QAWOTCInboundRev.txtContactName.Text = reader00(11).ToString()
                                        QAWOTCInboundRev.txtAccountNum.Text = reader00(12).ToString()
                                        QAWOTCInboundRev.txtCompany.Text = reader00(13).ToString
                                        QAWOTCInboundRev.txtContactPhone.Text = reader00(14).ToString
                                        QAWOTCInboundRev.txtContactEmail.Text = reader00(15).ToString

                                        QAWOTCInboundRev.dtpReviewdate.Text = reader00(16).ToString
                                        QAWOTCInboundRev.lblrevMan.Text = reader00(17).ToString
                                        QAWOTCInboundRev.txtRevComments.Text = reader00(18).ToString

                                        QAWOTCInboundRev.txtDisputerName.Text = reader00(20).ToString
                                        QAWOTCInboundRev.txtDisputeNotes.Text = reader00(21).ToString
                                        QAWOTCInboundRev.txtDisComment.Text = reader00(22).ToString


                                        QAWOTCInboundRev.cbo1_1.Text = reader00(23).ToString
                                        QAWOTCInboundRev.cbo1_2.Text = reader00(24).ToString
                                        QAWOTCInboundRev.cbo1_3.Text = reader00(25).ToString

                                        QAWOTCInboundRev.txt1_1.Text = reader00(32).ToString
                                        QAWOTCInboundRev.txt1_2.Text = reader00(33).ToString
                                        QAWOTCInboundRev.txt1_3.Text = reader00(34).ToString


                                        QAWOTCInboundRev.cbo2_1.Text = reader00(41).ToString
                                        QAWOTCInboundRev.txt2_1.Text = reader00(50).ToString



                                        QAWOTCInboundRev.cbo3_1.Text = reader00(59).ToString
                                        QAWOTCInboundRev.cbo3_2.Text = reader00(60).ToString
                                        QAWOTCInboundRev.cbo3_3.Text = reader00(61).ToString
                                        QAWOTCInboundRev.cbo3_4.Text = reader00(62).ToString
                                        QAWOTCInboundRev.cbo3_5.Text = reader00(63).ToString
                                        QAWOTCInboundRev.cbo3_6.Text = reader00(64).ToString
                                        QAWOTCInboundRev.cbo3_7.Text = reader00(65).ToString



                                        QAWOTCInboundRev.txt3_1.Text = reader00(68).ToString
                                        QAWOTCInboundRev.txt3_2.Text = reader00(69).ToString
                                        QAWOTCInboundRev.txt3_3.Text = reader00(70).ToString
                                        QAWOTCInboundRev.txt3_4.Text = reader00(71).ToString
                                        QAWOTCInboundRev.txt3_5.Text = reader00(72).ToString
                                        QAWOTCInboundRev.txt3_6.Text = reader00(73).ToString
                                        QAWOTCInboundRev.txt3_7.Text = reader00(74).ToString



                                        QAWOTCInboundRev.Cbo4_1.Text = reader00(77).ToString
                                        QAWOTCInboundRev.cbo4_2.Text = reader00(78).ToString
                                        QAWOTCInboundRev.cbo4_3.Text = reader00(79).ToString

                                        QAWOTCInboundRev.txt4_1.Text = reader00(86).ToString
                                        QAWOTCInboundRev.txt4_2.Text = reader00(87).ToString
                                        QAWOTCInboundRev.txt4_3.Text = reader00(88).ToString


                                        QAWOTCInboundRev.cbo5_1.Text = reader00(95).ToString
                                        QAWOTCInboundRev.cbo5_2.Text = reader00(96).ToString


                                        QAWOTCInboundRev.txt5_1.Text = reader00(104).ToString
                                        QAWOTCInboundRev.txt5_2.Text = reader00(105).ToString


                                        QAWOTCInboundRev.cbo6_1.Text = reader00(113).ToString


                                        QAWOTCInboundRev.txt6_1.Text = reader00(122).ToString


                                        QAWOTCInboundRev.cbo7_1.Text = reader00(131).ToString
                                        QAWOTCInboundRev.cbo7_2.Text = reader00(132).ToString
                                        QAWOTCInboundRev.cbo7_3.Text = reader00(133).ToString


                                        QAWOTCInboundRev.txt7_1.Text = reader00(140).ToString
                                        QAWOTCInboundRev.txt7_2.Text = reader00(141).ToString
                                        QAWOTCInboundRev.txt7_3.Text = reader00(142).ToString


                                        QAWOTCInboundRev.txtQAScore.Text = reader00(149).ToString


                                        QAWOTCInboundRev.cboAF.Text = reader00(152).ToString


                                        QAWOTCInboundRev.txtOrignalAuditor.Text = reader00(153).ToString
                                        QAWOTCInboundRev.txtSupervisor.Text = reader00(155).ToString

                                        QAWOTCInboundRev.txtTCXScore.Text = reader00(156).ToString
                                        QAWOTCInboundRev.txtWeekNumber.Text = reader00(157).ToString

                                        QAWOTCInboundRev.txtRevDate.Text = reader00(16).ToString


                                        QAWOTCInboundRev.txtMonth.Text = reader00(161).ToString



                                        QAWOTCInboundRev.txt1_1a.Text = reader00(162).ToString
                                        QAWOTCInboundRev.txt1_2a.Text = reader00(163).ToString
                                        QAWOTCInboundRev.txt1_3a.Text = reader00(164).ToString

                                        QAWOTCInboundRev.txt2_1a.Text = reader00(165).ToString

                                        QAWOTCInboundRev.txt3_1a.Text = reader00(169).ToString
                                        QAWOTCInboundRev.txt3_2a.Text = reader00(170).ToString
                                        QAWOTCInboundRev.txt3_3a.Text = reader00(171).ToString
                                        QAWOTCInboundRev.txt3_4a.Text = reader00(172).ToString
                                        QAWOTCInboundRev.txt3_5a.Text = reader00(173).ToString
                                        QAWOTCInboundRev.txt3_6a.Text = reader00(174).ToString
                                        QAWOTCInboundRev.txt3_7a.Text = reader00(175).ToString


                                        QAWOTCInboundRev.txt4_1a.Text = reader00(177).ToString
                                        QAWOTCInboundRev.txt4_2a.Text = reader00(178).ToString
                                        QAWOTCInboundRev.txt4_3a.Text = reader00(179).ToString


                                        QAWOTCInboundRev.txt5_1a.Text = reader00(181).ToString
                                        QAWOTCInboundRev.txt5_2a.Text = reader00(182).ToString

                                        QAWOTCInboundRev.txt6_1a.Text = reader00(187).ToString



                                        QAWOTCInboundRev.txt7_1a.Text = reader00(190).ToString
                                        QAWOTCInboundRev.txt7_2a.Text = reader00(191).ToString
                                        QAWOTCInboundRev.txt7_3a.Text = reader00(192).ToString



                                        QAWOTCInboundRev.txtEditedQA.Text = reader00(159).ToString
                                        QAWOTCInboundRev.txtDisputedQA.Text = reader00(160).ToString


                                        QAWOTCInboundRev.txtDisApp.Text = reader00(154).ToString
                                        QAWOTCInboundRev.txtDisputeScore.Text = reader00(19).ToString
                                        QAWOTCInboundRev.txtDisputedTCXScore.Text = reader00(204).ToString

                                        QAWOTCInboundRev.txtghostAFreason.Text = reader00(152).ToString


                                        '  QAWOTCInboundRev.lblrevMan.Text = lblQAauditor.Text


                                        QAWOTCInboundRev.lblQAauditor1.Text = lblQAauditor.Text

                                        QAWOTCInboundRev.lblEmailPassword.Text = lblEmailPassword.Text

                                        QAWOTCInboundRev.txtCSATScore.Text = reader00(208).ToString

                                        QAWOTCInboundRev.cboCSAT1.Text = reader00(209).ToString
                                        QAWOTCInboundRev.cboCSAT2.Text = reader00(210).ToString
                                        QAWOTCInboundRev.cboCSAT3.Text = reader00(211).ToString
                                        QAWOTCInboundRev.cboCSAT4.Text = reader00(212).ToString
                                        QAWOTCInboundRev.cboCSAT5.Text = reader00(213).ToString
                                        QAWOTCInboundRev.cboCSAT6.Text = reader00(214).ToString


                                        QAWOTCInboundRev.txtEMID.Text = reader00(215).ToString
                                        QAWOTCInboundRev.txtRegID.Text = reader00(216).ToString
                                        QAWOTCInboundRev.txtAHT.Text = reader00(217).ToString
                                    QAWOTCInboundRev.txtCallerType.Text = reader00(218).ToString



                                    QAWOTCInboundRev.txt1_1b.Text = reader00(219).ToString
                                    QAWOTCInboundRev.txt1_2b.Text = reader00(220).ToString
                                    QAWOTCInboundRev.txt1_3b.Text = reader00(221).ToString

                                    QAWOTCInboundRev.txt2_1b.Text = reader00(222).ToString

                                    QAWOTCInboundRev.txt3_1b.Text = reader00(226).ToString
                                    QAWOTCInboundRev.txt3_2b.Text = reader00(227).ToString
                                    QAWOTCInboundRev.txt3_3b.Text = reader00(228).ToString
                                    QAWOTCInboundRev.txt3_4b.Text = reader00(229).ToString
                                    QAWOTCInboundRev.txt3_5b.Text = reader00(230).ToString
                                    QAWOTCInboundRev.txt3_6b.Text = reader00(231).ToString
                                    QAWOTCInboundRev.txt3_7b.Text = reader00(232).ToString

                                    QAWOTCInboundRev.txt4_1b.Text = reader00(234).ToString
                                    QAWOTCInboundRev.txt4_2b.Text = reader00(235).ToString
                                    QAWOTCInboundRev.txt4_3b.Text = reader00(236).ToString


                                    QAWOTCInboundRev.txt5_1b.Text = reader00(237).ToString
                                    QAWOTCInboundRev.txt5_2b.Text = reader00(238).ToString

                                    QAWOTCInboundRev.txt6_1b.Text = reader00(243).ToString



                                    QAWOTCInboundRev.txt7_1b.Text = reader00(246).ToString
                                    QAWOTCInboundRev.txt7_2b.Text = reader00(247).ToString
                                    QAWOTCInboundRev.txt7_3b.Text = reader00(248).ToString


                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                Else





                                    MsgBox("Please retry your request")

                                    SplashScreenManager1.CloseWaitForm()

                                End If






                            End While
                            reader00.Close()
                            con00.Close()


                        End Using

                    End Using




                    ' Me.Cursor = Cursors.Hand


                End If


            End If








        Catch ex As Exception



            MsgBox(ex.Message)

        End Try


























    End Sub

    Private Sub GridView1_RowCellClick(sender As Object, e As XtraGrid.Views.Grid.RowCellClickEventArgs) Handles GridView1.RowCellClick
        Try


            '    txtAuditType.Text = GridView1.GetRowCellDisplayText(i, "CType").ToString


            ProgressBar2.Value = 0

            '    ProgressBar1.Value = 0


            '    For i As Integer = 0 To GridView1.DataRowCount - 1





            '        If GridView1.GetRowCellValue(i, "SR").ToString() = "" Then





            '            txtAuditType2.Text = GridView1.GetRowCellDisplayText(i, "CType").ToString
            '            txtSRPort2.Text = GridView1.GetRowCellValue(i, "ContactID").ToString



            '        Else



            '            txtAuditType2.Text = GridView1.GetRowCellDisplayText(i, "CType").ToString
            '            txtSRPort2.Text = GridView1.GetRowCellValue(i, "ContactID").ToString




            '        End If




            '    Next





        Catch ex As Exception



            '    MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub GridView1_FocusedRowChanged(sender As Object, e As XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles GridView1.FocusedRowChanged

        Try

            If GridView1.GetRowCellValue(e.FocusedRowHandle, "SR").ToString = "" Then

                txtAuditType2.Text = GridView1.GetRowCellDisplayText(e.FocusedRowHandle, "Ctype").ToString
                txtSRPort2.Text = GridView1.GetRowCellValue(e.FocusedRowHandle, "ContactID").ToString



            Else



                txtAuditType2.Text = GridView1.GetRowCellDisplayText(e.FocusedRowHandle, "Ctype").ToString
                txtSRPort2.Text = GridView1.GetRowCellValue(e.FocusedRowHandle, "SR").ToString


            End If






        Catch ex As Exception



            '  MsgBox(ex.Message)

        End Try


    End Sub

    Private Sub BackgroundWorker7_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker7.DoWork

        Try

            For i = 0 To 100

                System.Threading.Thread.Sleep(32)


                Me.BackgroundWorker7.ReportProgress(i)

                '    lblprogr.Text = i.ToString

                i = i
            Next







        Catch ex As Exception



            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub BackgroundWorker7_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker7.ProgressChanged


        ProgressBar2.Value = e.ProgressPercentage


    End Sub

    Private Sub BackgroundWorker7_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker7.RunWorkerCompleted

        Try

            ' OpenCard.Enabled = True


            Me.Cursor = Cursors.Hand


            If txtAuditType2.Text = "Call" Then


                QACallRev.Show()
                SplashScreenManager1.CloseWaitForm()

            ElseIf txtAuditType2.Text = "Email" Then


                QAEmailRev.Show()
                SplashScreenManager1.CloseWaitForm()

            ElseIf txtAuditType2.Text = "Chat" Then


                QAChatRev.Show()

                SplashScreenManager1.CloseWaitForm()

            ElseIf txtAuditType2.Text = "WOTC Inbound" Then


                QAWOTCInboundRev.Show()

                SplashScreenManager1.CloseWaitForm()

            End If




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub

    Private Sub GridView1_CustomDrawRowIndicator(sender As Object, e As XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs) Handles GridView1.CustomDrawRowIndicator


        e.Info.ImageIndex = -1


    End Sub

    Private Sub GridView1_ColumnFilterChanged(sender As Object, e As EventArgs) Handles GridView1.ColumnFilterChanged


        Try



            RowCounting2()



        Catch ex As System.ArgumentOutOfRangeException


        Catch ex As Exception



            MsgBox(ex.Message)

        End Try



    End Sub





    Public Sub RowCounting2()


        'countersql = 0
        'counter2sql = 0
        'counter3sql = 0



        ''Loop For Showing the Pending review




        'For i = 0 To (GridView1.RowCount - 1)



        '    Dim qad = CDate(DataGridView2.Rows(i).Cells("QA_Date").Value.ToString)


        '    Dim qad00 = CDate(GridView1.GetRow("Rev_Date").ToString)


        '    Dim expiredate00 = qad00.AddDays(7)

        '    Dim Threedaynoti00 = qad00.AddDays(3)



        '    If Now > expiredate And DataGridView2.Rows(i).Cells("Rev_Date").Value = "9/9/2020" Then


        '        If GridView1.GetRow("Rev_Date").Value = "9/9/2020" Then

        '            counter3sql += 1


        '        End If







        '' Total pending Reivew


        'If GridView1.GetRow("Rev_Date").Value = "9/9/2020" Then



        '    Int32.TryParse(DataGridView2.Rows(i).Cells("Rev_Date").Value.ToString(), temp)

        '    PendingTotal += temp

        '    counter += 1





        '' Total Reviewed


        'ElseIf qad < Now Then



        '    counter2 += 1






        'End If



        'Next

        'lblTotalPassDueSQL1.Text = counter3sql


        'lblpendreviewsql.Text = counter

        'lblpastduesql.Text = counter3
        'lblTotalsql.Text = counter2
        'lblTotalsql.Text = rowCount

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click










    End Sub



    Private Sub GridView1_CustomSummaryCalculate(sender As Object, e As DevExpress.Data.CustomSummaryEventArgs) Handles GridView1.CustomSummaryCalculate






    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click



        Try


            Dim myStream As Stream = Nothing
            Dim openFileDialog1 As New OpenFileDialog()

            openFileDialog1.InitialDirectory = lblMDrive.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA2\"

            '  openFileDialog1.InitialDirectory = "C:\Users\playe\desktop\qa2"



            'openFileDialog1.Filter = "Excel |*.xlsx"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True


            If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                myStream = openFileDialog1.OpenFile()

                Process.Start(openFileDialog1.FileName)




            End If




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try




    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click




        Try

            SplashScreenManager1.ShowWaitForm()
            Me.Cursor = Cursors.AppStarting
            Me.QaMainDBTableAdapter7.ReadyForReview(Me.QADBDataSet6.QAMainDB)




            Me.Cursor = Cursors.Hand
            SplashScreenManager1.CloseWaitForm()

        Catch ex As Exception

            Me.Cursor = Cursors.Hand

            MsgBox(ex.Message)


        End Try






    End Sub



    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        Try

            SplashScreenManager1.ShowWaitForm()
            Me.Cursor = Cursors.AppStarting


            Me.QaMainDBTableAdapter7.ThreeDaysPastDue(Me.QADBDataSet6.QAMainDB)



            Me.Cursor = Cursors.Hand
            SplashScreenManager1.CloseWaitForm()
        Catch ex As Exception



            MsgBox(ex.Message)



            Me.Cursor = Cursors.Hand


        End Try





    End Sub



    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click


        Try
            SplashScreenManager1.ShowWaitForm()
            Me.Cursor = Cursors.AppStarting

            Me.QaMainDBTableAdapter7.OverDue(Me.QADBDataSet6.QAMainDB)


            Me.Cursor = Cursors.Hand
            SplashScreenManager1.CloseWaitForm()
        Catch ex As Exception



            MsgBox(ex.Message)

            Me.Cursor = Cursors.Hand

        End Try




    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click



        Try

            SplashScreenManager1.ShowWaitForm()
            Me.Cursor = Cursors.AppStarting

            Me.QaMainDBTableAdapter7.Reviewed(Me.QADBDataSet6.QAMainDB)


            Me.Cursor = Cursors.Hand
            SplashScreenManager1.CloseWaitForm()

        Catch ex As Exception



            MsgBox(ex.Message)

            Me.Cursor = Cursors.Hand


        End Try



    End Sub

    Private Sub BackgroundWorker8_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker8.DoWork

        Try


            refreshDB()

        Catch ex As Exception



            MsgBox(ex.Message)


        End Try



    End Sub

    Private Sub RefreshForSupervisorToolStripButton_Click(sender As Object, e As EventArgs)
        Try
            Me.QaMainDBTableAdapter7.RefreshForSupervisor(Me.QADBDataSet6.QAMainDB)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Button14_Click_1(sender As Object, e As EventArgs) Handles Button14.Click



        Try


            If lblDeciderDash.Text = "Admin" Then


                QAMainDBBindingSource1.Filter = "[QA_Date] >= '" & DateTimePicker5.Text & "' AND [QA_Date] <= '" & DateTimePicker4.Text & " '"





            ElseIf lblDeciderDash.Text = "QaAuditor" Then



                QAMainDBBindingSource1.Filter = "[QA_Date] >= '" & DateTimePicker5.Text & "' AND [QA_Date] <= '" & DateTimePicker4.Text & " '"


            ElseIf lblDeciderDash.Text = "Supervisor" Then




                QAMainDBBindingSource1.Filter = "[QA_Date] >= '" & DateTimePicker5.Text & "' AND [QA_Date] <= '" & DateTimePicker4.Text & "' AND [Supervisor] = '" & lblQAAuditor3.Text & "'"




            ElseIf lblDeciderDash.Text = "TeamLead" Then



                QAMainDBBindingSource1.Filter = "[QA_Date] >= '" & DateTimePicker5.Text & "' AND [QA_Date] <= '" & DateTimePicker4.Text & "' AND [Supervisor] = '" & lblQAAuditor3.Text & "'"




            ElseIf lblDeciderDash.Text = "GOCSupervisor" Then





                QAMainDBBindingSource1.Filter = "[QA_Date] >= '" & DateTimePicker5.Text & "' AND [QA_Date] <= '" & DateTimePicker4.Text & "' AND [Supervisor] = '" & lblQAAuditor3.Text & "'"





            ElseIf lblDeciderDash.Text = "GOCTeamLead" Then




                QAMainDBBindingSource1.Filter = "[QA_Date] >= '" & DateTimePicker5.Text & "' AND [QA_Date] <= '" & DateTimePicker4.Text & "' AND [Supervisor] = '" & lblQAAuditor3.Text & "'"




            End If







        Catch ex As Exception



            MsgBox(ex.Message)


        End Try





    End Sub

    Private Sub BackgroundWorker8_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker8.RunWorkerCompleted



        Me.Cursor = Cursors.Hand


    End Sub

    Private Sub BackgroundWorker9_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker9.DoWork

        Try

            refreshDB2()




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try




    End Sub

    Private Sub BackgroundWorker9_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker9.RunWorkerCompleted

        Me.Cursor = Cursors.Hand



    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click



        FillSpreadSheet()


    End Sub



    Public Sub RefreshWorkbook()

        'Dim xlsApp As Excel.Application = Nothing
        'Dim xlsWorkBooks As Excel.Workbooks = Nothing
        'Dim xlsWB As Excel.Workbook = QATrendspath1

        ''Try

        'xlsApp = New Excel.Application
        'xlsWorkBooks = xlsApp.Workbooks
        'xlsWB = xlsWorkBooks.Open(QATrendspath1)


        'xlsWB.Close()
        'xlsWB = Nothing
        'xlsApp.Quit()
        'xlsApp = Nothing




        'Catch ex As Exception

        'Finally

        '    'xlsWB.Close()
        '    'xlsWB = Nothing
        '    'xlsApp.Quit()
        '    'xlsApp = Nothing

        'End Try

    End Sub




    Public Sub FillSpreadSheet()

        Try



            ' SpreadsheetControl1.LoadDocument("C:\Users\playe\Desktop\QA\Book2.xlsm", DocumentFormat.Xlsm)

            ' SpreadsheetControl1.LoadDocument("C:\Users\playe\Desktop\QA\QAExcellMaped.xlsm", DocumentFormat.Xlsm)

            '  SpreadsheetControl1.LoadDocument("P:\SPC\QA\QAExcell.xlsm", DocumentFormat.Xlsm)


            '   SpreadsheetControl1.LoadDocument("C:\Users\durraner\Desktop\ReviewPending.xlsm", DocumentFormat.Xlsm)


            '   SpreadsheetControl1.LoadDocument(QATrendspath1, DocumentFormat.Xlsx)


            SpreadsheetControl1.LoadDocument(QATrendspath1, DocumentFormat.Xlsx)


            SpreadsheetControl1.Document.DocumentSettings.ShowPivotTableFieldList = False




        Catch ex As Exception

            SpreadsheetControl1.Refresh()

            MsgBox(ex.Message)

        End Try



    End Sub



    Private Sub SpreadsheetControl1_CellBeginEdit(sender As Object, e As XtraSpreadsheet.SpreadsheetCellCancelEventArgs) Handles SpreadsheetControl1.CellBeginEdit



        ' If e.ColumnIndex = 1 And e.RowIndex = 1 Then

        e.Cancel = True

        '   End If



    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboMWSelector.SelectedIndexChanged


        Try

            Dim msg = "Please wait while the scorecard loads"

            Dim title = "FADV QA Application"

            Dim style = MsgBoxStyle.OkOnly

            Dim responce = MsgBox(msg, style, title)



            If cboMWSelector.Text = "Call" Then

                MWBworker.RunWorkerAsync()
                '      FillSpreadSheet()

                MWTimer.Start()


            ElseIf cboMWSelector.Text = "Email" Then




            ElseIf cboMWSelector.Text = "Chat" Then



            End If




        Catch ex As Exception




            '  RefreshWorkbook()

            ' xlsApp.Workbooks(QATrendspath1).Close(SaveChanges:=False)

            MsgBox(ex.Message)

        End Try




    End Sub

    Private Sub MWBworker_DoWork(sender As Object, e As DoWorkEventArgs) Handles MWBworker.DoWork






        'Dim i As Integer

        'For i = 0 To 100

        '    System.Threading.Thread.Sleep(200)


        '    Me.MWBworker.ReportProgress(i)

        '    '  MWBworker.ReportProgress(CInt(100 * i / Max) & i.ToString)

        '    i = i

        'Next

        ' MWTimer.Start()


        FillSpreadSheet()









    End Sub

    Private Sub MWBworker_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles MWBworker.ProgressChanged



        '    ProgressBar3.Value = e.ProgressPercentage

    End Sub

    Private Sub MWBworker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles MWBworker.RunWorkerCompleted





    End Sub

    Private Sub MWTimer_Tick(sender As Object, e As EventArgs) Handles MWTimer.Tick

        ProgressBar3.Increment(5)



        If ProgressBar3.Value = 100 Then


            MWTimer.Stop()

        End If

        '   


    End Sub

    Private Sub MWBworker_Disposed(sender As Object, e As EventArgs) Handles MWBworker.Disposed

    End Sub

    Private Sub RefreshExcel2_Tick(sender As Object, e As EventArgs) Handles RefreshExcel2.Tick


        SpreadsheetControl1.LoadDocument(QATrendspath1, DocumentFormat.Xlsm)


        RefreshExcel2.Enabled = False



    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked



        Try


            Dim msg = "Are you sure you want to export to excel?"

            Dim title = "FADV QA Application"

            Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or MsgBoxStyle.Question

            Dim responce = MsgBox(msg, style, title)



            Dim myStream As Stream = Nothing
            Dim openFileDialog2 As New OpenFileDialog()

            openFileDialog2.InitialDirectory = lblMDrive.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA2\"


            If responce = MsgBoxResult.Yes Then



                '            String FileName = "C:\\MyFiles\\Grid.xls";
                'MyGridControl.ExportToXls(FileName);


                GridControl1.ExportToXlsx(lblMDrive.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA2\ExcelDump.xlsx")


                '   GridControl1.ExportToXlsx("C:\Users\playe\Desktop\QA2\ExcelDump.xlsx")


            Else





            End If









        Catch ex As Exception



            MsgBox(ex.Message)


        End Try















    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked



        Try


            Dim myStream As Stream = Nothing
            Dim openFileDialog1 As New OpenFileDialog()

            openFileDialog1.InitialDirectory = lblMDrive.Text & "\Users\" & lblSCRN.Text & "\Desktop\QA2\"

            '  openFileDialog1.InitialDirectory = "C:\Users\playe\desktop\qa2"



            'openFileDialog1.Filter = "Excel |*.xlsx"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True


            If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                myStream = openFileDialog1.OpenFile()

                Process.Start(openFileDialog1.FileName)




            End If




        Catch ex As Exception



            MsgBox(ex.Message)


        End Try


















    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick


        lblAnn.Visible = True

        If lblAnn.Right < 0 Then

            lblAnn.Left = Me.ClientSize.Width
        Else
            lblAnn.Left -= 10

        End If

        Timer5.Enabled = True


    End Sub

    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick

        lblPleaseUpdateApp.Visible = True

        If lblPleaseUpdateApp.Right < 0 Then

            lblPleaseUpdateApp.Left = Me.ClientSize.Width
        Else
            lblPleaseUpdateApp.Left -= 10

        End If


        Timer5.Enabled = True


    End Sub

    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick


        Timer3.Enabled = False

        Timer4.Enabled = False

        lblPleaseUpdateApp.Visible = False
        lblAnn.Visible = False


    End Sub

    Private Sub Timer6_Tick(sender As Object, e As EventArgs) Handles Timer6.Tick


        Changer()


    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click



        WebBrowser1.Navigate(QATrendspath1, False)



        'Dim WebC As New WebClient

        'WebC.Headers.Add(HttpRequestHeader.Cookie, WebBrowser1.Document.Cookie)

        'WebC.DownloadFile(QATrendspath1, "Copy of QA_Trends.xlsx")



    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click




        Shell(QATrendspath1)


    End Sub

    Private Sub WebBrowser1_FileDownload(sender As Object, e As EventArgs) Handles WebBrowser1.FileDownload



    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        Try

            Me.Cursor = Cursors.AppStarting
            SplashScreenManager1.ShowWaitForm()

            Me.QaMainDBTableAdapter7.DisputedQA(Me.QADBDataSet6.QAMainDB)


            Me.Cursor = Cursors.Hand
            SplashScreenManager1.CloseWaitForm()

        Catch ex As Exception



            MsgBox(ex.Message)

            Me.Cursor = Cursors.Hand


        End Try





    End Sub

    Private Sub LnkLableChangelog_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LnkLableChangelog.LinkClicked

        Try


            Process.Start("P:\QA Application\QA1\ChangeLog.xlsx")


        Catch ex As Exception

            MsgBox("Make sure your are connected to the P drive.")

        End Try



    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked


        '' Dim Desk = My.Computer.FileSystem.SpecialDirectories.Desktop

        '' Retrieve storage account from connection string.
        'Dim storageAccount As CloudStorageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("edurrantString"))

        '' Create the blob client.
        'Dim blobClient As CloudBlobClient = storageAccount.CreateCloudBlobClient()

        '' Retrieve reference to a previously created container.
        'Dim container As CloudBlobContainer = blobClient.GetContainerReference("edblob")

        '' Retrieve reference to a blob named "photo1.jpg".
        'Dim blockBlob As CloudBlockBlob = container.GetBlockBlobReference("QATrendsM.xlsm")

        '' Save blob contents to a file.
        ''    Using fileStream = System.IO.File.OpenWrite(Desk)

        'Using fileStream = System.IO.File.OpenRead("C: \Users\playe\Desktop\QA2")


        '    blockBlob.DownloadToStream(fileStream)




        'End Using


        Process.Start(QATrendspath1)



    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click

        Try

            Me.Cursor = Cursors.AppStarting
            SplashScreenManager1.ShowWaitForm()

            Me.QaMainDBTableAdapter7.ApprovedFill(Me.QADBDataSet6.QAMainDB)


            Me.Cursor = Cursors.Hand
            SplashScreenManager1.CloseWaitForm()

        Catch ex As Exception



            MsgBox(ex.Message)

            Me.Cursor = Cursors.Hand


        End Try







    End Sub

    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click

        Try

            Me.Cursor = Cursors.AppStarting
            SplashScreenManager1.ShowWaitForm()

            Me.QaMainDBTableAdapter7.DeniedFill(Me.QADBDataSet6.QAMainDB)


            Me.Cursor = Cursors.Hand
            SplashScreenManager1.CloseWaitForm()

        Catch ex As Exception



            MsgBox(ex.Message)

            Me.Cursor = Cursors.Hand


        End Try










    End Sub

    Private Sub RefillDGToolStripButton_Click(sender As Object, e As EventArgs)
        Try
            Me.QaMainDBTableAdapter7.RefillDG(Me.QADBDataSet6.QAMainDB)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub lnkTransferAudits_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnkTransferAudits.LinkClicked



        TransferAudits.Show()




    End Sub


End Class


