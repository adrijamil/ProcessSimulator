Imports System.Diagnostics.Debug

Public Class FlowSheet

    ' ------------------------------------------------------------------------------------------------------
    ' ClassName   : FlowSheet
    ' Version     : 1.0
    ' Author      : AJS
    ' Purpose     : This class "owns" stream, unit operation and property package objects. The Solver is in here
    ' Remarks     : This class should be versatile ie should work with any Property Package
    ' Improvements: 1) To implement calculation levels.
    '               2) Think about putting the solver elsewhere. Does it help with multiple flowsheets? Maybe a SimCase object is needed
    '
    ' ------------------------------------------------------------------------------------------------------
    '
    ' ------------------------------------------------------------------------------------------------------
    Public IsActive As Boolean 'solver active or not
    Public Ports As New Collection

    Private myName As String
    Private myMoniker As String
    Private myStreams As New System.Collections.Generic.SortedList(Of String, Stream) ' should make this private and write property to access
    Private myUnitOps As New System.Collections.Generic.SortedList(Of String, Object) ' should make this private and write property to access
    Private myPropPack As Object ' dim as Object so you can use any Property Package object

    Private SolveStack As New System.Collections.Generic.Stack(Of Object) ' no one else needs access to the stack

    ReadOnly Property Streams(Optional ByVal daIndex As Object = -32767) As Object
        Get
            If IsNumeric(daIndex) Then
                If daIndex = -32767 Then
                    Streams = myStreams.Values
                Else
                    Streams = myStreams.Values(daIndex)
                End If
            Else
                Streams = myStreams(daIndex)
            End If
        End Get
    End Property



    ReadOnly Property UnitOps(Optional ByVal daIndex As Object = -32767) As Object 'can get either single object or whole collection
        Get
            If IsNumeric(daIndex) Then
                If daIndex = -32767 Then
                    UnitOps = myUnitOps.Values
                Else
                    UnitOps = myUnitOps.Values(daIndex)
                End If
            Else
                UnitOps = UnitOps(daIndex)
            End If
        End Get
    End Property

    ReadOnly Property Moniker As String
        Get
            Moniker = myMoniker
        End Get
    End Property

    Property Name As String
        Get
            Name = myName
        End Get
        Set(daName As String)
            myName = daName
            myMoniker = "FS:" & myName
        End Set
    End Property

    ReadOnly Property PropPack As Object
        Get
            PropPack = myPropPack
        End Get
    End Property
    ReadOnly Property IsSolved As Boolean
        Get
            'check if all streams are solved
            'generally, if all streams are solved then all unit ops have done their part [LOGIC CHECK]
            'if above comment is false then check all unit ops as well
            Dim tempstrm As Stream
            IsSolved = True
            For Each tempstrm In Streams
                If tempstrm.IsSolved = False Then
                    IsSolved = False
                    Exit For
                End If
            Next
        End Get
    End Property

    Public Sub SetPackage(daPackName As String)
        'Only Ideal and RefProp available at the moment
        'Should figure out how to show options when using this sub eg like how vbBlack, vbBlue keywords come out.
        Select Case daPackName
            Case "Ideal"
                myPropPack = New IdealGas
            Case "RefProp"
                '   myPropPack = New RefProp_PP
            Case Else
                'do something
        End Select
        myPropPack.Parent = Me
    End Sub

    Public Sub AddStream(daStream As Stream)
        daStream.SetParent(Me)
        myStreams.Add(daStream.Name, daStream)
    End Sub
    Public Sub AddUnitOp(daObject As Object)
        myUnitOps.Add(daObject.Name, daObject)
    End Sub

    Public Sub Forget()

        Dim tempVar As Variable
        Dim tempstrm As Stream
        Dim tempStrm2 As Stream
        Dim tempUO As Object
        Dim n As Integer
        Dim i As Integer
        Dim SomeItem As Object 'SomeItem might be a Unit Operation or a Stream

        On Error GoTo errhandler

        On Error GoTo errhandler
        '    SolveStack.Clear() ' can this go in the header?
        For Each tempstrm In myStreams.Values
            tempstrm.Forget()
        Next

        For Each tempUO In myUnitOps.Values
            tempUO.Forget()
        Next


errhandler:
        If Err.Number <> 0 Then
            Debug.Print(Err.Description)
            Err.Clear()
            Resume Next
        End If
    End Sub

    Public Sub Solve()

        Dim tempVar As Variable
        Dim tempstrm As Stream
        Dim tempStrm2 As Stream
        Dim tempUO As Object
        Dim n As Integer
        Dim i As Integer
        Dim j As Integer
        Dim SomeItem As Object 'SomeItem might be a Unit Operation or a Stream
        Dim StackCount As Integer

        On Error GoTo errhandler

        'add all to the stack
        'might be useful to write an algorithm to figure out which items should calculate first
        '  SolveStack = New System.Collections.Generic.Stack(Of Object) ' can this go in the header?
        SolveStack.Clear()

        For Each tempstrm In myStreams.Values
            SolveStack.Push(tempstrm)
        Next

        For Each tempUO In myUnitOps.Values
            SolveStack.Push(tempUO)
        Next
        StackCount = SolveStack.Count
        n = 0
        j = 0
        'start trying to solve each item in stack
        Do While IsSolved = False


            SomeItem = SolveStack.Peek
            SomeItem.Execute()
            If SomeItem.IsSolved = True Then
                SolveStack.Pop()
            End If

            If SolveStack.Count = 0 Then Exit Do ' then calcs are all done

            'n = 0
            'For Each SomeItem In Stack.Values
            '    n = n + 1
            '    SomeItem.Execute() 'For a streams, it will execute a flash if it can (DOF = 0). For unit ops, it will pass all the variables that it can.

            '    If SomeItem.IsSolved = True Then 'if it is solved (Stream flashed, UnitOp has passed all variables) remove it
            '        Stack.Remove(n)
            '        n = n - 1
            '    End If

            'If myStack.Count = 0 Then Exit Do ' then calcs are all done
            'Next

            ' for debugging purposes, it shouldnt take more than a few passes

            j = j + 1
            n = n + 1

            If j = 20 Then
                MsgBox("not solved, over 20 passes")
                Exit Do
            End If

            'make sure stack 
            If n = StackCount Then
                n = 0
                If SolveStack.Count = StackCount Then
                    Exit Do
                Else
                    StackCount = SolveStack.Count
                End If

            End If
        Loop

errhandler:
        If Err.Number <> 0 Then
            Debug.Print(Err.Description)
            Err.Clear()
            Resume Next
        End If
    End Sub

    Public Sub New()
        IsActive = False

        'give me a default name
    End Sub

End Class
