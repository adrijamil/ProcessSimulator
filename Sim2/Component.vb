Imports System.IO
Imports System.Xml.Serialization
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office
Public Class Component

    ' ------------------------------------------------------------------------------------------------------
    ' ClassName   : Component
    ' Version     : 1.0
    ' Author      : AJS
    ' Purpose     : To hold information about a species. To be used as part of collection the Property Package objects
    ' Remarks     : This class should be versatile ie should work with any Property Package
    ' Improvements: 1) To be able to get data from a .dat file.
    '               2) To add properties as needed.
    ' ------------------------------------------------------------------------------------------------------
    '
    ' ------------------------------------------------------------------------------------------------------

    'Base variable types are used for components
    'This is because they should be constant. ie you do not need the functionality of the Variable object
    'All variables are public so the PropPack access them

    Public Name As String
    Public MolWt As Double 'g/mol
    Public CpMass As Double 'J/K/mol
    Public Tc As Double 'in Kelvin
    Public Pc As Double 'in Pascal
    Public AcFact As Double 'unitless
    Public Hvap As Double 'in J/mol
    Public RefState As String 'is it gas or liquid at ref T (25C)
    Public StdLiqDens As Double ' kg/m3

    ' this part is not used by anyone
    ' potentially be used by IdealGas proppack to get Pvap_i=f(T)
    'Private Structure AntoinesCoeffSet
    '    Public a As Double
    'b As Double
    'c As Double
    'tmax As Double
    'End Structure

    'Public Structure CompForDb
    '    Public Name As String
    '    Public ID_Num As Integer
    '    Public Mw As Double
    '    Public Cp As Double
    '    Public Tc As Double
    '    Public Pc As Double
    '    Public Accentric As Double
    '    Public Hvap As Double
    '    Public RefState As String
    '    Public StdLiqDens As Double
    'End Structure


    'ReadOnly Property Name As String
    '    Get
    '        Name = myName
    '    End Get
    'End Property

    Public Sub MakeComponent(daName As String, daDB As String, Optional tempComp As CompForDb = Nothing)

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

        'this sub reads all component data from excel sheet and puts in .dat file

        'xlApp = New Excel.Application
        '  WB = xlApp.Workbooks.Open("C:\Users\User\Desktop\TheSimulator\TheSim_1.0.1_AJS07.xlsm")

        'myFileName = daDB

        ' ifile = FreeFile()
        'Open (myFileName For Binary) As ifile
        'FileOpen(ifile, myFileName, OpenMode.Binary, OpenAccess.Default)

        i = 1
        ' tempComp = Nothing

        If Not tempComp Is Nothing Then
            With tempComp

                Name = .Name
                MolWt = .Mw
                CpMass = .Cp
                Tc = .Tc
                Pc = .Pc
                AcFact = .Accentric
                Hvap = .Hvap
                RefState = .RefState
                StdLiqDens = .StdLiqDens
            End With
            GoTo skip
        End If


        i = 1
        'Do Until myFileStream.Position = myFileStream.Length

        Dim myFileStream As FileStream = New FileStream("C:\Users\Fuad\Desktop\development\Sim2\FilesNeeded\ComponentDatabase.xml", FileMode.Open)
        tempComponent = Nothing

        CompArray = CType(mySerializer.Deserialize(myFileStream), Component())
        For Each tempComponent In CompArray
            'Dim tempcomp As New CompForDB
            With tempComponent
                ' MsgBox(.Name)
                If UCase(.Name) = UCase(daName) Then
                    Name = .Name
                    MolWt = .MolWt
                    CpMass = .CpMass
                    Tc = .Tc
                    Pc = .Pc
                    AcFact = .AcFact
                    Hvap = .Hvap
                    RefState = .RefState
                    StdLiqDens = .StdLiqDens
                End If

            End With
        Next


        myFileStream.Close()



skip:

    End Sub





    Public Sub New()

    End Sub

End Class
