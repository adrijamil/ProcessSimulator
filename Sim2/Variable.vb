Public Class Variable
    Public IsKnown As Boolean
    Public IsCalculated As Boolean
    Public LinkedCell As String
    Public Unit As String
    Private myVal As Object
    Private myName As String 'can use this as a "moniker"
    Private myQuantityType As String
    Private myBaseVarType As String ' double, integer, string
    Private myParent As Object

    Property Name As String
        Get
            Name = myName
        End Get
        Set(daName As String)
            myName = daName
        End Set
    End Property

    ReadOnly Property Moniker() As String
        Get
            'what is my parent 'parent might be stream,phase, or unitop
            Moniker = myParent.Moniker & "." & myName
        End Get
    End Property

    Property Parent() As Object
        Get
            Parent = myParent
        End Get
        Set(daParent As Object)
            myParent = daParent
        End Set
    End Property

    WriteOnly Property QuantityType As String
        Set(daType As String)
            Select Case daType
                Case "Pressure", "Temperature"
                    myBaseVarType = "Double"
                Case "MassFlow", "MoleFlow", "MolecularWeight", "MassDensity", "MassEnthalpy", "MolarComposition", "MassHeatCapacity", "PhaseMoleFraction"
                    myBaseVarType = "Double"
                Case "Boolean" 'if its boolean then it is always known
                    IsKnown = True
                    myBaseVarType = "Boolean"
                Case "Integer"
                    myBaseVarType = "Integer"
            End Select

            myQuantityType = daType
        End Set
    End Property

    Property Val As Object


        Set(ByVal daVal As Object)
            'can accept array or lone double
            'can also be used for settings: booleans, integers

            myVal = daVal
            Dim i As Integer

            Dim isEmpty As Boolean

            isEmpty = False

            Select Case myBaseVarType
                Case "Double", "Integer", ""
                    'check IsArray or not
                    If IsArray(daVal) Then
                        For i = 1 To UBound(daVal)
                            If daVal(i) = -32767 Then
                                isEmpty = True
                            End If
                        Next

                        If isEmpty = False Then
                            IsKnown = True
                        End If

                    Else
                        If Not daVal = -32767 Then
                            IsKnown = True
                        Else
                            IsCalculated = True
                        End If
                    End If

                Case Else 'Booleans. They should never have to be calculated (unless to activate choke for example)

                    IsKnown = True
                    IsCalculated = False
            End Select
        End Set
        Get
            Val = myVal
        End Get
    End Property


    Public Sub New()
        IsCalculated = True
        myVal = -32767
        IsKnown = False
    End Sub

End Class
