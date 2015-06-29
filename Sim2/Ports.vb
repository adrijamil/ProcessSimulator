
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office

Public Class Port

    Private myVariable As Variable
    Private myMoniker As String
    Private myExcelCell As Range
    Private myExcelLocked As Boolean
    Private myIsConnection As Boolean
    Private myConnection As Connection

    Private Structure Connection
        Public ConnectionType As String 'inlet or outlet
        Public myStream As Stream
        Public myUnitOp As Object
    End Structure

    Property ExcelCell As Range
        Get
            ExcelCell = myExcelCell
        End Get
        Set(daRange As Range)
            myExcelCell = daRange
        End Set
    End Property

    ReadOnly Property Name As String
        Get
            Name = myVariable.Name
        End Get
    End Property

    Property Variable() As Variable
        Get
            Variable = myVariable
        End Get
        Set(daVar As Variable)
            myVariable = daVar
        End Set
    End Property

    Property ExcelLocked As Boolean
        Get
            ExcelLocked = myExcelLocked
        End Get
        Set(daBool As Boolean)
            myExcelLocked = daBool
        End Set
    End Property

    ReadOnly Property Moniker As String
        Get
            Moniker = myVariable.Moniker
        End Get
    End Property

    Public Sub SetVal(daVal As Object, Optional daUnit As String = "")
        'daval can be an array too
        Dim daInt As String
        '  Dim tempVar As Variable
        Dim tempVariant As Object
        Dim nComps As Integer

        'Dim tempstrm
        'Set tempstrm = myVariable.Parent


        If InStr(daVal, " ") <> 0 Then 'means im trying to pass a double in an array
            nComps = myVariable.Parent.Parent.PropPack.nComps
            daInt = Split(daVal, " ")(0)
            daVal = Split(daVal, " ")(1)
            tempVariant = myVariable.Val
            daVal = CDbl(daVal)
            If IsArray(tempVariant) = False Then
                ReDim tempVariant(0 To nComps - 1)
            End If
            tempVariant(daInt - 1) = daVal
            myVariable.Val = tempVariant
        Else
            myVariable.Val = daVal
        End If

        myVariable.IsCalculated = False

    End Sub

    Public Sub SetConnection(ByVal daType As String, daStream As Stream, daUO As Object)

        myIsConnection = True
        myConnection.ConnectionType = daType
        myConnection.myStream = daStream
        myConnection.myUnitOp = daUO

    End Sub
    Sub UpdateExcel()
        'can this work for arrays?
        'Application.EnableEvents = False

        If Not myExcelCell Is Nothing Then
            If IsArray(myVariable.Val) Then
                myExcelCell.Value = myExcelCell.Application.WorksheetFunction.Transpose(myVariable.Val)
            Else
                myExcelCell.Value = myVariable.Val
            End If
        End If
        'Application.EnableEvents = True

    End Sub

    Public Sub New()
        myIsConnection = False

    End Sub

End Class
