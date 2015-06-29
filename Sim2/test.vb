Imports Excel = Microsoft.Office.Interop.Excel
Public Class test
    Private xlApp As Excel.Application
    Private xlWB As Excel.Workbook
    Public Sub OpenExcel()
        xlApp = New Excel.Application

        '  xlWB = xlApp.Workbooks.Open("C:\Users\User\Desktop\TheSimulator\WorksheetTemplate.xlsx")

        xlApp.Visible = True
    End Sub

End Class
