Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office
Imports System.IO
Imports System.Xml.Serialization

Module Ancillary
'test'

    Sub makeCompDB()
        Dim WS As Worksheet
        ' Dim myLoc As String
        ' Dim myFileName As String
        ' Dim ifile As Integer
        Dim i As Integer
        Dim tempComp As CompForDb
        Dim tempComponent As Component
        Dim xlApp As Excel.Application
        Dim CompArray() As Component ' System.Collections.ArrayList

        Dim WB As Workbook
        ' Dim myObject As MySerializableClass = New MySerializableClass()
        ' Insert code to set properties and fields of the object.
        Dim mySerializer As XmlSerializer = New XmlSerializer(GetType(Component()))
        ' To write to a file, create a StreamWriter object.
        Dim myWriter As StreamWriter = New StreamWriter("C:\Users\Fuad\Desktop\development\Sim2\FilesNeeded\ComponentDatabase.xml")

        'this sub reads all component data from excel sheet and puts in .dat file

        xlApp = New Excel.Application
        WB = xlApp.Workbooks.Open("C:\Users\Fuad\Desktop\development\Sim2\FilesNeeded\TheSim_1.0.1_AJS07.xlsm")
        WS = WB.Worksheets("ComponentDatabase")
        xlApp.Visible = True


        i = 1
        Do While WS.Cells(2 + i, 2).value <> ""


            tempComp = New CompForDB
            With tempComp
                .Name = WS.Cells(2 + i, 2).value
                .ID_Num = i
                .Mw = WS.Cells(2 + i, 3).value
                .Cp = WS.Cells(2 + i, 4).value
                .Tc = WS.Cells(2 + i, 5).value
                .Pc = WS.Cells(2 + i, 6).value
                .Accentric = WS.Cells(2 + i, 7).value
                .Hvap = WS.Cells(2 + i, 8).value
                .RefState = WS.Cells(2 + i, 9).value
                .StdLiqDens = WS.Cells(2 + i, 10).value
            End With

            tempComponent = New Component

            tempComponent.MakeComponent(tempComp.Name, "", tempComp)

            ReDim Preserve CompArray(i - 1)
            CompArray(i - 1) = tempComponent





            i = i + 1
        Loop
        mySerializer.Serialize(myWriter, CompArray)
        Dim proc As System.Diagnostics.Process

        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next


        myWriter.Close()
    End Sub


    Sub ReadCompDB()
        ' Dim WS As Worksheet
        ' Dim myLoc As String
        ' Dim myFileName As String
        ' Dim ifile As Integer
        Dim i As Integer
        'Dim tempComp As CompForDb
        Dim tempComponent As Component
        '  Dim xlApp As Excel.Application
        '  Dim WB As Workbook
        ' Dim myObject As MySerializableClass = New MySerializableClass()
        ' Insert code to set properties and fields of the object.
        Dim mySerializer As XmlSerializer = New XmlSerializer(GetType(Component()))
        ' To write to a file, create a StreamWriter object.
        Dim CompArray() As Component
        Dim myFileStream As FileStream = New FileStream("C:\Users\User\Desktop\TheSimulator\ComponentDatabase.xml", FileMode.Open)
        'this sub reads all component data from excel sheet and puts in .dat file

        'xlApp = New Excel.Application
        '  WB = xlApp.Workbooks.Open("C:\Users\User\Desktop\TheSimulator\TheSim_1.0.1_AJS07.xlsm")
        ' WS = WB.Worksheets("ComponentDatabase")




        i = 1
        'Do Until myFileStream.Position = myFileStream.Length


        tempComponent = Nothing

        CompArray = CType(mySerializer.Deserialize(myFileStream), Component())
        For Each tempComponent In CompArray
            'Dim tempcomp As New CompForDB
            With tempComponent
                System.Console.WriteLine(.Name)
                System.Console.WriteLine(.MolWt)
                ' System.Console.WriteLine(.)
            End With
        Next

        i = i + 1
        ' Loop
        myFileStream.Close()

    End Sub

End Module
