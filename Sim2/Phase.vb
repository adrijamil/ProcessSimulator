Public Class Phase
    Public Name As String

    'Public F As Double
    Private myParent As Stream

    'myvariables
    Private myFMass As New Variable
    Private myMw As New Variable
    Private myMassDensity As New Variable
    Private myComposition As New Variable
    Private myMassEnthalpy As New Variable
    Private myMassCp As New Variable
    Private myMoleFrac As New Variable
    Private myPorts As New Collection
    Private myComponents() As String 'this is just a list of components- not really important for algorithms
    Private myCompMws() As Double 'check if this is actually used


    ReadOnly Property Ports As Collection
        Get
            Ports = myPorts
        End Get
    End Property
    ReadOnly Property Moniker As String
        Get
            Moniker = myParent.Moniker & ".PHASE:" & Name
        End Get
    End Property
    ReadOnly Property MoleFrac As Variable
        Get
            MoleFrac = myMoleFrac
        End Get
    End Property
    ReadOnly Property FMass As Variable
        Get
            FMass = myFMass
        End Get
    End Property
    ReadOnly Property Mw As Variable
        Get
            Mw = myMw
        End Get
    End Property
    ReadOnly Property MassDensity As Variable
        Get
            MassDensity = myMassDensity
        End Get
    End Property
    ReadOnly Property Composition As Variable
        Get
            Composition = myComposition
        End Get
    End Property
    ReadOnly Property MassEnthalpy As Variable
        Get
            MassEnthalpy = myMassEnthalpy
        End Get
    End Property
    ReadOnly Property MassCp As Variable
        Get
            MassCp = myMassCp
        End Get
    End Property
    ReadOnly Property Parent As Stream
        Get
            Parent = myParent
        End Get
    End Property

    Sub SetParent(daParent As Stream)
        'Dim tempPort As Port
        myParent = daParent
        'For Each tempPort In myPorts
        '    paParent.Ports.Add(tempPort)
        'Next
    End Sub

    Public Sub CalcMw()
        Dim i As Integer
        Mw.Val = 0

        For i = 0 To UBound(Composition.Val) - 1
            Mw.Val = Mw.Val + Composition.Val(i) * myParent.PropPack.Component(i).MolWt
        Next

    End Sub
    Public Sub New()

        Dim tempPort As Port

        myFMass.QuantityType = "FMass"
        myFMass.Name = "FMass"
        myFMass.Parent = Me
        tempPort = New Port
        tempPort.Variable = myFMass
        myPorts.Add(tempPort) '

        myMw.QuantityType = "Mw"
        myMw.Name = "Mw"
        myMw.Parent = Me
        tempPort = New Port
        tempPort.Variable = myMw
        myPorts.Add(tempPort)

        myMassDensity.QuantityType = "MassDensity"
        myMassDensity.Name = "MassDensity"
        myMassDensity.Parent = Me
        tempPort = New Port
        tempPort.Variable = myMassDensity
        myPorts.Add(tempPort)

        myMassEnthalpy.QuantityType = "MassEnthalpy"
        myMassEnthalpy.Name = "MassEnthalpy"
        myMassEnthalpy.Parent = Me
        tempPort = New Port
        tempPort.Variable = myMassEnthalpy
        myPorts.Add(tempPort)


        myMoleFrac.QuantityType = "MoleFrac"
        myMoleFrac.Name = "MoleFrac"
        myMoleFrac.Parent = Me
        tempPort = New Port
        tempPort.Variable = myMoleFrac
        myPorts.Add(tempPort)

        myComposition.QuantityType = "Composition"
        myComposition.Name = "Composition"
        myComposition.Parent = Me
        tempPort = New Port
        tempPort.Variable = myComposition
        myPorts.Add(tempPort)

        'myVapFrac.QuantityType = "MoleFraction"
        'myVapFrac.Name = "VapourFraction"
        ' myVapFrac.Parent = Me
        ' Set tempPort = New Port
        'tempPort.SetVar myVapFrac
        'myPorts.Add tempPort

        myMassCp.QuantityType = "MassHeatCapacity"
        myMassCp.Name = "MassCp"
        myMassCp.Parent = Me
        tempPort = New Port
        tempPort.Variable = myMassCp
        myPorts.Add(tempPort)

    End Sub

    Function Clone() As Phase
        Dim tempPhase As New Phase

        With tempPhase
            .SetParent(myParent)
            .FMass.Val = myFMass.Val
            .Mw.Val = myMw.Val
            .MassDensity.Val = myMassDensity.Val
            .MassEnthalpy.Val = myMassEnthalpy.Val
            .Composition.Val = myComposition.Val
            .MassCp.Val = myMassCp.Val
            .MoleFrac.Val = myMoleFrac.Val
        End With

        Clone = tempPhase

    End Function
End Class
