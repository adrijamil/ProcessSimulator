Imports Excel = Microsoft.Office.Interop.Excel
'Imports Microsoft.

Imports Math = System.Math
Imports System

Module Main
    Public myTest As test
    Sub Main()
        myTest = New test
        'Ancillary.makeCompDB()
        Dim str As String
        Dim MyCase As New SimCase
        Dim CaseName As String
        Dim myIO As InputOutput
        'myTest.OpenExcel()
        myIO = MyCase.IO
        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next

        myIO = MyCase.IO
        CaseName = "C:\Users\Fuad\Desktop\development\Sim2\FilesNeeded\test.txt"
        myIO.OpenExcelIFace()
        myIO.ReadInputFile(CaseName)
        Do
                str = Console.ReadLine
                'myIO.SendMessage(Str)
                If str = "Exit" Then Exit Do
        Loop
    End Sub

End Module
