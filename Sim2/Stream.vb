Public Class Stream
    Public Name As String
    Public Ports As New System.Collections.Generic.List(Of Port)
    Public Phases As New System.Collections.Generic.SortedList(Of String, Phase)
    Public ToUnit As String
    Public FromUnit As String
    Public myParent As FlowSheet 'make this private

    Private myPressure As New Variable
    Private myTemperature As New Variable
    Private myVapourFraction As Variable
    Private myPropPack As Object
    Private Components() As String
    Private CompMws() As Double
    Private myIsSolved As Boolean
    Private myIsDirty As Boolean

    ReadOnly Property Moniker As String

        Get  ' can only get moniker
            Moniker = myParent.Moniker & ".STRM:" & Name
        End Get

    End Property

    WriteOnly Property IsDirty As Boolean
        Set(daBool As Boolean)
            myIsDirty = daBool
        End Set
    End Property

    ReadOnly Property Temperature() As Variable
        Get
            Temperature = myTemperature
        End Get
    End Property

    ReadOnly Property Pressure() As Variable
        Get
            Pressure = myPressure
        End Get
    End Property
    ReadOnly Property VapourFraction() As Variable
        Get
            VapourFraction = myVapourFraction
        End Get
    End Property

    ReadOnly Property PropPack() As Object
        Get
            PropPack = myPropPack
        End Get
    End Property
    ReadOnly Property FMass() As Variable
        Get
            FMass = Phases(1).FMass
        End Get
    End Property
    ReadOnly Property Mw() As Variable
        Get
            Mw = Phases(1).Mw
        End Get
    End Property
    ReadOnly Property MassDensity() As Variable
        Get
            MassDensity = Phases(1).MassDensity
        End Get
    End Property
    ReadOnly Property Composition() As Variable
        Get
            Composition = Phases(1).Composition
        End Get
    End Property
    ReadOnly Property MassEnthalpy() As Variable
        Get
            MassEnthalpy = Phases(1).MassEnthalpy
        End Get
    End Property
    ReadOnly Property MassCp() As Variable
        Get
            MassCp = Phases(1).MassCp
        End Get
    End Property

    ReadOnly Property CanSolve() As Boolean
        Get
            If Pressure.IsKnown And (Temperature.IsKnown Or MassEnthalpy.IsKnown) And Composition.IsKnown Then
                CanSolve = True
            Else
                CanSolve = False
            End If
        End Get
    End Property

    Property IsSolved As Boolean
        Get
            IsSolved = False

            'check all variables
            If Pressure.IsKnown And Temperature.IsKnown And Composition.IsKnown And MassEnthalpy.IsKnown Then
                IsSolved = True
            End If

        End Get
        Set(daBool As Boolean)

            myIsSolved = daBool
            'Dim tempVar As Variable

            'forget my calculated vars
            If daBool = False Then
                'For Each tempVar In vars
                '    If tempVar.IsCalculated = True Then
                '        tempVar.IsKnown = False
                '        tempVar.Val = -32767
                '    End If
                'Next

                If Not myParent Is Nothing Then

                    'go upstream
                    If Not ToUnit = vbNullString Then
                        If myParent.UnitOps(ToUnit).IsSolved Then
                            myParent.UnitOps(ToUnit).IsSolved = False
                        End If
                    End If
                    'go downstream
                    If Not FromUnit = vbNullString Then
                        If myParent.UnitOps(FromUnit).IsSolved Then
                            myParent.UnitOps(FromUnit).IsSolved = False
                        End If
                    End If
                End If

            End If

        End Set
    End Property



    Public Sub AddPhase(daPhase As Phase)
        Dim tempPort As Port

        For Each tempPort In daPhase.Ports
            Ports.Add(tempPort) ', tempPort.Moniker
        Next

        Phases.Add(daPhase.Name, daPhase)

    End Sub

    Public Sub SetParent(daParent As FlowSheet)
        myParent = daParent
        myPropPack = myParent.PropPack

    End Sub

    Public Sub Forget()
        If myIsDirty = True Then
            If myPressure.IsCalculated Then myPressure.IsKnown = False
            If myTemperature.IsCalculated Then myTemperature.IsKnown = False
            If myVapourFraction.IsCalculated Then myVapourFraction.IsKnown = False
            If MassEnthalpy.IsCalculated Then MassEnthalpy.IsKnown = False
        End If

    End Sub

    Public Sub Execute()

        '  Dim DaUnknown As String
        '    Dim daKnown As String
        '    Dim daMw As Double
        '   Dim daDens As Double
        Dim P As Double
        Dim t As Double
        P = Pressure.Val
        t = Temperature.Val

        If CanSolve = False Then
            Exit Sub
        End If


        'using property package
        If MassEnthalpy.IsKnown = False Then
            myPropPack.PT_Flash(Me)
        ElseIf Temperature.IsKnown = False Then
            myPropPack.PH_Flash(Me)
        End If

        IsSolved = True


    End Sub


    Public Sub New()

        Dim tempPort As Port
        Dim tempVap As Phase
        Dim tempLiq As Phase
        Dim tempOverall As Phase
        myIsSolved = False
        ToUnit = ""
        FromUnit = ""

        Pressure.QuantityType = "Pressure"
        Pressure.Name = "Pressure"
        Pressure.Parent = Me

        tempPort = New Port
        tempPort.Variable = (Pressure)
        Ports.Add(tempPort)

        Temperature.QuantityType = "Temperature"
        Temperature.Name = "Temperature"
        Temperature.Parent = Me
        tempPort = New Port
        tempPort.Variable = (Temperature)
        Ports.Add(tempPort)


        tempOverall = New Phase
        tempOverall.SetParent(Me)
        tempOverall.Name = "Overall"
        AddPhase(tempOverall)


        tempVap = New Phase
        tempVap.SetParent(Me)
        tempVap.Name = "Vapour"
        AddPhase(tempVap)
        'link the variables so you can have 2 ports to same variable
        '-so can spec vfrac via stream but not via phase
        myVapourFraction = tempVap.MoleFrac

        tempLiq = New Phase
        tempLiq.SetParent(Me)
        tempLiq.Name = "Liquid"
        AddPhase(tempLiq)

        myIsDirty = True

    End Sub



    Public Sub Map()
        Dim tempPort As Port


        For Each tempPort In Ports
            Select Case tempPort.Name
            End Select
        Next

    End Sub


    Public Function PhasePresent(daString As String) As Boolean
        PhasePresent = False

        Dim tempPhase As Phase
        For Each tempPhase In Phases.Values

            If tempPhase.Name = daString Then PhasePresent = True
        Next

    End Function
End Class
