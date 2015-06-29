
Imports Microsoft.Office
Imports Excel = Microsoft.Office.Interop.Excel
Imports System
Imports System.IO
Imports System.Text

Public Class InputOutput


    Public WithEvents WS As Excel.Worksheet


    Private myCase As SimCase
    Private mySimPorts As New Collection
    Private myXLPorts As New Collection


    Private xlApp As Excel.Application
    Private xlWB As Excel.Workbook

    Property SimulationCase As SimCase
        Get
            SimulationCase = myCase
        End Get
        Set(daCase As SimCase)
            myCase = daCase
        End Set
    End Property

    Private Sub WS_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
        If Target.Address(False, False) = "B2" Then

        End If

    End Sub

    Private Sub XLBtnSendMessage_Click()
        MsgBox("fuck")

    End Sub
    Private Sub WS_Change1(Target As Range) Handles WS.Change
        '   Console.WriteLine(Target.Value)
        ' MsgBox(Target.Value)

        On Error GoTo errorhandler

        Dim tempPort As Port

        tempPort = myXLPorts(Target.Address)


        If Not tempPort.ExcelLocked = True Then
            If CStr(Target.Value) = "" Then
                Target.Value = "<empty>"
                tempPort.SetVal(-32767)
            Else
                tempPort.SetVal(Target.Value)
            End If
            tempPort.Variable.Parent.IsDirty = True
        Else
            'put the value back; cannot change
            tempPort.UpdateExcel()
        End If


        Dim tempFS As FlowSheet
        '"forgetting" loop
        For Each tempFS In myCase.FlowSheets
            tempFS.Forget()
        Next

        SolveAll()
        UpdatePorts()
errorhandler:
        If Err.Number <> 0 Then
            Console.WriteLine(Err.Description)
            Console.WriteLine(Err.Description)
            Err.Clear()
            Resume Next
        End If

    End Sub





    Public Function SendMessage(daString As String) As String ' send back a string like "0:no error" or "1:cannot compute"

        '    Dim daObj() As String
        ' Dim daAction As String
        Dim daName As String
        ' Dim i As Integer
        Dim tempFS As FlowSheet
        Dim tempstrm As Stream
        Dim daUnitStr As String
        Dim daObjStr As Object
        Dim ExitCode As Integer
        Dim tempUnitOp As Object
        Dim tempPort As Port
        Dim tempComp As Component
        Dim daValStr As String
        Dim daActionStr As String

        ' Debug.Print(daString)
        tempUnitOp = Nothing

        If myCase Is Nothing Then myCase = New SimCase


        ExitCode = 0 'by default no error
        daValStr = ""

        daActionStr = Split(UCase(daString), " ")(0)
        If UBound(Split(UCase(daString), " ")) > 1 Then
            daValStr = Split(UCase(daString), " ")(2)
        End If

        daString = Split(UCase(daString), " ")(1)

        If InStr(1, daString, ".") <> 0 Then
            daObjStr = Split(UCase(daString), ".")

            '    For i = 0 To UBound(daObjStr)
            '        Debug.Print daObjStr(i)
            '    Next
        Else
            daObjStr = daString

        End If

        Select Case daActionStr

            Case "ADD"
                ' Add stream, flowsheet, unitop
                ' Add component
                If IsArray(daObjStr) = False Then ' must be add flow sheet
                    'SYNTAX: "ADD FS(myflowsheetname)
                    daName = Mid(daObjStr, InStr(1, daObjStr, "(") + 1, InStr(1, daObjStr, ")") - InStr(1, daObjStr, "(") - 1)
                    tempFS = New FlowSheet
                    tempFS.Name = daName
                    myCase.FlowSheets.Add(tempFS, daName)
                Else ' either unit op or stream or porperty package
                    'SYNTAX: "ADD FS(myflowsheetname).STRM(mystreamname)
                    If Left(daObjStr(1), 4) = "STRM" Then 'add a stream
                        daName = Mid(daObjStr(1), InStr(1, daObjStr(1), "(") + 1, InStr(1, daObjStr(1), ")") - InStr(1, daObjStr(1), "(") - 1)
                        tempstrm = New Stream
                        tempFS = myCase.FlowSheets(Mid(daObjStr(0), InStr(1, daObjStr(0), "(") + 1, InStr(1, daObjStr(0), ")") - InStr(1, daObjStr(0), "(") - 1))
                        tempstrm.Name = daName

                        tempFS.AddStream(tempstrm)

                        'this needs to happen after I add to
                        'tempstrm.Map()

                        For Each tempPort In tempstrm.Ports
                            Console.Write(tempPort.Moniker & vbNewLine)

                            mySimPorts.Add(tempPort, tempPort.Moniker) ', tempPort.Moniker

                        Next

                    ElseIf Left(daObjStr(1), 6) = "UNITOP" Then ' must be Unit Op
                        'SYNTAX: "ADD FS(myflowsheetname).UNITOP.VALVE(myValveName)
                        daUnitStr = Left(daObjStr(2), InStr(1, daObjStr(2), "(") - 1)
                        tempFS = myCase.FlowSheets(Mid(daObjStr(0), InStr(1, daObjStr(0), "(") + 1, InStr(1, daObjStr(0), ")") - InStr(1, daObjStr(0), "(") - 1))
                        daName = Mid(daObjStr(2), InStr(1, daObjStr(2), "(") + 1, InStr(1, daObjStr(2), ")") - InStr(1, daObjStr(2), "(") - 1)
                        Select Case daUnitStr
                            Case "VALVE"
                                '   tempUnitOp = New Valve

                            Case "MIXER"
                                '  tempUnitOp = New Mixer
                            Case "SPLITTER"
                                '  tempUnitOp = New Splitter
                            Case Else

                        End Select

                        If Not tempUnitOp Is Nothing Then tempUnitOp.Name = daName
                        tempFS.AddUnitOp(tempUnitOp)

                        For Each tempPort In tempUnitOp.Ports
                            mySimPorts.Add(tempPort, tempPort.Moniker) ', tempPort.Moniker
                            '       myXLPorts.Add(tempPort.Moniker, tempPort.ExcelCell.Address)
                        Next

                    ElseIf Left(daObjStr(1), 8) = "PROPPACK" Then ' add property package
                        'SYNTAX: "ADD FS(myflowsheetname).PROPPACK.IDEAL(myIdeal)
                        daUnitStr = Left(daObjStr(2), InStr(1, daObjStr(2), "(") - 1)
                        tempFS = myCase.FlowSheets(Mid(daObjStr(0), InStr(1, daObjStr(0), "(") + 1, InStr(1, daObjStr(0), ")") - InStr(1, daObjStr(0), "(") - 1))
                        daName = Mid(daObjStr(2), InStr(1, daObjStr(2), "(") + 1, InStr(1, daObjStr(2), ")") - InStr(1, daObjStr(2), "(") - 1)

                        Select Case daUnitStr
                            Case "IDEAL"
                                tempFS.SetPackage("Ideal")
                            Case "REFPROP"
                                tempFS.SetPackage("RefProp")
                            Case Else

                        End Select

                        tempFS.PropPack.Name = daName

                    ElseIf Left(daObjStr(1), 9) = "COMPONENT" Then ' add property package
                        'SYNTAX: "ADD FS(myflowsheetname).COMPONENT(METHANE)
                        tempFS = myCase.FlowSheets(Mid(daObjStr(0), InStr(1, daObjStr(0), "(") + 1, InStr(1, daObjStr(0), ")") - InStr(1, daObjStr(0), "(") - 1))
                        'Set tempFS = myCase.FlowSheets("MYFLOWSHEETNAME")
                        daName = Mid(daObjStr(1), InStr(1, daObjStr(1), "(") + 1, InStr(1, daObjStr(1), ")") - InStr(1, daObjStr(1), "(") - 1)

                        tempComp = tempFS.PropPack.CreateComponent(daName)
                        tempFS.PropPack.AddComponent(tempComp)

                    End If

                End If
            Case "SET"
                'set stream specs
                'set operation specs
                'set solver auto on/off
                'handle connections here?
                tempFS = myCase.FlowSheets(Mid(daObjStr(0), InStr(1, daObjStr(0), "(") + 1, InStr(1, daObjStr(0), ")") - InStr(1, daObjStr(0), "(") - 1))
                If Left(daObjStr(1), 6) = "SOLVER" Then 'MUST BE solver
                    'SYNTAX: "SET FS(myflowsheetname).SOLVER TRUE
                    If daValStr = "TRUE" Then
                        tempFS.IsActive = True
                    Else
                        tempFS.IsActive = False
                    End If
                ElseIf Left(daObjStr(1), 4) = "STRM" Then
                    'SYNTAX: "SET FS(myflowsheetname).STRM(mystreamname).VarName Value Unit
                    tempstrm = tempFS.Streams(CStr(Mid(daObjStr(1), InStr(1, daObjStr(1), "(") + 1, InStr(1, daObjStr(1), ")") - InStr(1, daObjStr(1), "(") - 1)))
                    daName = daObjStr(2)
                    If InStr(1, daName, "(") <> 0 Then daName = Left(daName, Len(daName) - 3)
                    For Each tempPort In tempstrm.Ports

                        If UCase(tempPort.Name) = daName Then
                            If UBound(daObjStr) > 2 Then 'must be composition
                                daValStr = daObjStr(3) & " " & daValStr
                            End If
                            tempPort.SetVal(daValStr)
                            Exit For
                        End If
                    Next

                ElseIf Left(daObjStr(1), 6) = "UNITOP" Then
                    'SYNTAX: "SET FS(myflowsheetname).UNITOP.VALVE(myValve).VarName Value Unit
                    tempUnitOp = tempFS.UnitOps(Mid(daObjStr(2), InStr(1, daObjStr(2), "(") + 1, InStr(1, daObjStr(2), ")") - InStr(1, daObjStr(2), "(") - 1))
                    daName = daObjStr(3)
                    For Each tempPort In tempUnitOp.Ports
                        If tempPort.Name = daName Then
                            tempPort.SetVal(daValStr)
                            Exit For
                        End If
                    Next
                End If

            Case "DELETE"
                'delete stream/unit/fs
                'delete value
                'delete connection

            Case "CONNECT"
                'SYNTAX: "CONNECT FS(myflowsheetname).UNITOP.VALVE(myValveName).Inlet STRM(mystream)
                tempFS = myCase.FlowSheets(Mid(daObjStr(0), InStr(1, daObjStr(0), "(") + 1, InStr(1, daObjStr(0), ")") - InStr(1, daObjStr(0), "(") - 1))
                tempUnitOp = tempFS.UnitOps(Mid(daObjStr(2), InStr(1, daObjStr(2), "(") + 1, InStr(1, daObjStr(2), ")") - InStr(1, daObjStr(2), "(") - 1))
                'Debug.Print "fuck"
                tempstrm = tempFS.Streams(Mid(daValStr, InStr(1, daValStr, "(") + 1, InStr(1, daValStr, ")") - InStr(1, daValStr, "(") - 1))

                Select Case daObjStr(3)
                    Case "Inlet"
                        tempUnitOp.InletStreams.Add(tempstrm, tempstrm.Name)
                        tempstrm.ToUnit = tempUnitOp.Name
                    Case "Outlet"
                        tempUnitOp.OutletStreams.Add(tempstrm, tempstrm.Name)
                        tempstrm.FromUnit = tempUnitOp.Name
                End Select

                tempPort = New Port
                tempPort.SetConnection(daObjStr(3), tempstrm, tempUnitOp)

            Case "PRINT"
                'either by moniker or name


            Case Else
                ExitCode = 1


        End Select

        Select Case ExitCode
            Case 0

            Case 1

        End Select

        SendMessage = "A-OK"

        SolveAll()
        MapPorts()
        UpdatePorts()


    End Function
    Sub SolveAll()


        Dim tempFS As FlowSheet
        For Each tempFS In myCase.FlowSheets
            If tempFS.IsActive Then
                tempFS.Solve()
            End If
        Next


    End Sub


    Sub UpdatePorts()
        Dim j As Integer
        Dim CompList() As String
        If Not myCase.FlowSheets(1).PropPack Is Nothing Then


            For j = 0 To myCase.FlowSheets(1).PropPack.nComps - 1
                ReDim Preserve CompList(j)
                CompList(j) = myCase.FlowSheets(1).PropPack.Component(j).name
            Next

        End If
        Dim tempPort As Port
        xlApp.EnableEvents = False

        For Each tempPort In mySimPorts
            'MsgBox(tempPort.ExcelCell.Address)

            tempPort.UpdateExcel()
            If tempPort.Name = "Composition" Then
                tempPort.ExcelCell.Offset(0, -1).Value = tempPort.ExcelCell.Application.WorksheetFunction.Transpose(CompList)
            End If
        Next
        xlApp.EnableEvents = True

    End Sub
    Private Sub WS_BtnSendMessage_Click()
        MsgBox("fuckyou")

    End Sub
    Sub MapPorts()
        Dim tempFS As FlowSheet
        Dim tempstrm As Stream
        Dim tempPort As Port
        Dim darange As Range
        Dim daMoniker() As String
        Dim StrStart As Range
        Dim UOStart As Range
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim NC As Integer
        Dim CompList() As String

        On Error GoTo errorhandler
        darange = Nothing


        StrStart = WS.Range("D9")

        If myCase.FlowSheets.Count = 0 Then Exit Sub
        If myCase.FlowSheets(1).PropPack Is Nothing Then Exit Sub
        NC = myCase.FlowSheets(1).PropPack.nComps

        For Each tempstrm In myCase.FlowSheets(1).Streams

            i = i + 1
            For Each tempPort In tempstrm.Ports
                daMoniker = Split(tempPort.Moniker, ".")

                If Left(daMoniker(2), 5) = "PHASE" Then
                    Select Case tempPort.Variable.Parent.Name
                        Case "Overall"

                            Select Case daMoniker(3)
                                ' Case "MassEnthalpy":
                                '     Set daRange = StrStart.Offset(5, i)
                                Case "FMass"
                                    darange = StrStart.Offset(4, i)
                                    ' Case "MassCp":
                                    '    Set daRange = StrStart.Offset(5, i)
                                Case "MoleFrac"
                                    darange = StrStart.Offset(3, i)
                                Case "Mw"
                                    darange = StrStart.Offset(5, i)
                                Case "MassDensity"
                                    darange = StrStart.Offset(6, i)
                                Case "Composition"
                                    darange = WS.Range(StrStart.Offset(8, i).Address & ":" & StrStart.Offset(8 + NC - 1, i).Address)


                            End Select

                        Case "Vapour"

                            Select Case daMoniker(3)
                                ' Case "MassEnthalpy":
                                '     Set daRange = StrStart.Offset(5, i)
                                'Case "FMass":
                                '    Set daRange = StrStart.Offset(5, i)
                                ' Case "MassCp":
                                '    Set daRange = StrStart.Offset(5, i)
                                Case "MoleFrac"
                                    darange = StrStart.Offset(9 + NC, i)
                                Case "Mw"
                                    darange = StrStart.Offset(10 + NC, i)
                                Case "MassDensity"
                                    darange = StrStart.Offset(11 + NC, i)
                                Case "Composition"
                                    darange = WS.Range(StrStart.Offset(13 + NC, i).Address & ":" & StrStart.Offset(13 + NC + NC - 1, i).Address)

                            End Select

                        Case "Liquid"
                            Select Case daMoniker(3)
                                ' Case "MassEnthalpy":
                                '     Set daRange = StrStart.Offset(5, i)
                                'Case "FMass":
                                '    Set daRange = StrStart.Offset(5, i)
                                ' Case "MassCp":
                                '    Set daRange = StrStart.Offset(5, i)
                                Case "MoleFrac"
                                    darange = StrStart.Offset(14 + 2 * NC, i)
                                Case "Mw"
                                    darange = StrStart.Offset(15 + 2 * NC, i)
                                Case "MassDensity"
                                    darange = StrStart.Offset(16 + 2 * NC, i)
                                Case "Composition"
                                    darange = WS.Range(StrStart.Offset(18 + 2 * NC, i).Address & ":" & StrStart.Offset(18 + 2 * NC + NC - 1, i).Address)

                            End Select
                    End Select
                Else
                    Select Case daMoniker(2)
                        Case "Temperature"
                            darange = StrStart.Offset(2, i)
                        Case "Pressure"
                            darange = StrStart.Offset(1, i)
                        Case "VapourFraction"
                            darange = StrStart.Offset(3, i)
                    End Select

                End If

                If Not darange Is Nothing Then tempPort.ExcelCell = darange
                myXLPorts.Add(tempPort, tempPort.ExcelCell.Address)

                darange = Nothing

            Next
        Next
errorhandler:
        If Err.Number <> 0 Then
            Debug.Print(Err.Description)
            Err.Clear()
            Resume Next
        End If



    End Sub

    Public Sub ReadInputFile(daPath As String)

        Dim iFile As Integer
        Dim objReader As New StreamReader(daPath)
        Dim tempStr As String

        ' iFile = FreeFile()
        tempStr = ""

        '  FileOpen(iFile, daPath, OpenMode.Input, OpenAccess.Read)


        Do Until objReader.EndOfStream()
            tempStr = objReader.ReadLine
            ' Console.WriteLine(tempStr)

            SendMessage(tempStr)
        Loop
        FileClose(iFile)
    End Sub

    Public Sub OpenExcelIFace()
        xlApp = New Excel.Application

        xlWB = xlApp.Workbooks.Open("C:\Users\Fuad\Desktop\development\Sim2\FilesNeeded\Template.xlsm")
        'xlWB = xlApp.Workbooks.Add

        ' WS = xlWB.Worksheets("SimWorkSheet")
        WS = xlWB.Worksheets("Sheet1")
        xlApp.Visible = True
        '  XLBtn_SendMessage = WS.Shapes(0)




    End Sub
    Private Sub FormatSheet()

    End Sub

    'Private Sub xlApp_WorkbookBeforeClose(Wb As Workbook, ByRef Cancel As Boolean) Handles xlApp.WorkbookBeforeClose

    'End Sub
End Class


