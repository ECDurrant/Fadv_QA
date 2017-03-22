
Imports DevExpress.DataAccess.Excel
Imports Microsoft.Office.Interop


Module ExcelStuff





    Sub Macro()

        Try









            'Dim xlApp = New Excel.Application

            'Dim xlApp1 = New Excel.Application

            'Dim xlWorkBook As Excel.Workbook

            Dim exeDir1 As New IO.FileInfo(Reflection.Assembly.GetExecutingAssembly.FullName)

            Dim QATrendspath1 = IO.Path.Combine(exeDir1.DirectoryName, "Copy of QA_Trends.xlsx")


            '  ~~> Start Excel And open the workbook.

            'xlWorkBook = xlApp.Workbooks.Open(QATrendspath1)



            '~~> Run the macros.
            '  xlApp.Run("Macro3")

            '   xlApp.Run("Refresh3")
            '   xlApp1.Run("Refresh4")

            '~~> Clean-up: Close the workbook and quit Excel.
            '    xlWorkBook.Close(False)

            '~~> Quit the Excel Application
            '    xlApp.Quit()




        Catch ex As Exception



            MsgBox(ex.Message)

        End Try






    End Sub


End Module
